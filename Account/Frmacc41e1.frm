VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc41e1 
   AutoRedraw      =   -1  'True
   Caption         =   "簽收作業查詢"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5040
   ScaleWidth      =   8865
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      MaxLength       =   1
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton CmdSearch 
      Caption         =   "查詢"
      Height          =   375
      Left            =   7800
      TabIndex        =   3
      Top             =   182
      Width           =   915
   End
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
      Height          =   330
      Left            =   5280
      TabIndex        =   2
      Top             =   204
      Width           =   2400
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
      Height          =   330
      Left            =   2535
      TabIndex        =   1
      Top             =   204
      Width           =   2400
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
      Height          =   330
      ItemData        =   "Frmacc41e1.frx":0000
      Left            =   135
      List            =   "Frmacc41e1.frx":0002
      TabIndex        =   0
      Top             =   204
      Width           =   2052
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc41e1.frx":0004
      Height          =   3850
      Left            =   0
      TabIndex        =   5
      Top             =   1060
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "A2301"
         Caption         =   "簽收單號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "A2302"
         Caption         =   "繳款日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0##/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "A2303"
         Caption         =   "智權人員"
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
      BeginProperty Column03 
         DataField       =   "A2308"
         Caption         =   "繳收據日"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0##/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "CU04"
         Caption         =   "客戶名稱"
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
         DataField       =   "A2322"
         Caption         =   "銀存科目"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Type"
         Caption         =   "類別"
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
      BeginProperty Column07 
         DataField       =   "Amount"
         Caption         =   "簽收金額"
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
         Size            =   275
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1065.26
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   6240
      Top             =   600
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
   Begin VB.Label Label3 
      Caption         =   "所別：       (1.北所 2.中所 3.南所 4.高所 空白.全所)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   645
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   12
      Top             =   4764
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
      Height          =   255
      Left            =   5025
      TabIndex        =   7
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
      Left            =   2295
      TabIndex        =   6
      Top             =   210
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc41e1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/07 Form2.0已修改 DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit
Dim adoadodc1 As New ADODB.Recordset
'Added by Lydia 2015/11/13 F12功能的按鈕
Private Sub cmdSearch_Click()
    KeyDefine vbKeyF12
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
   'Modified by Lydia 2015/11/17
   'Me.Width = 8850
   Me.Width = 8985
   Me.Height = 5500 'Modify by Amy 2024/08/21 原:5400
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   strCon1 = "智權人員(未繳收據)"
   strCon2 = "智權人員"
   strCon3 = "繳款日期"
   strCon4 = "繳收據日"
   strCon5 = "銀存科目" 'Add by Amy 2018/02/02 +銀存科目
   Combo1.AddItem MsgText(31)
   Combo1.AddItem strCon1
   Combo1.AddItem strCon2
   Combo1.AddItem strCon3
   Combo1.AddItem strCon4
   Combo1.AddItem strCon5 'Add by Amy 2018/02/02 +銀存科目
   Combo1 = MsgText(31)
   'Added by Lydia 2015/11/17 北所人員可以看它所
   If pub_strUserOffice <> "1" Then
      Text1.Enabled = False
   End If
   Text1.Text = pub_strUserOffice
   'end 2015/11/17
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Adodc1.Recordset.RecordCount <> 0 Then
      strCompanyNo = Adodc1.Recordset.Fields("a2301").Value
   Else
      strCompanyNo = MsgText(601)
      strItemNo = MsgText(601)
   End If
   StatusClear
   tool1_enabled
   With Frmacc41e0
      .Enabled = True
      .Show
      If strCompanyNo <> MsgText(601) Then
         .txtA2301 = strCompanyNo
         .ReadData .txtA2301
      End If
   End With
   'Add by Amy 2014/01/07 未清會帶入其他Form
   strCompanyNo = MsgText(601)
   strItemNo = MsgText(601)
   'end 2014/01/07
   Set Frmacc41e1 = Nothing
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Screen.MousePointer = vbHourglass
         Acc230Query
         Screen.MousePointer = vbDefault
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  傳票資料查詢
'
'*************************************************
Private Sub Acc230Query()
Dim strSort As String   'add by sonia 2015/11/25 加排序條件
Dim strQ As String 'Add by Amy 2018/02/02

On Error GoTo Checking
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   Select Case Combo1
      Case strCon1 '智權人員(未繳收據)
         strCondition = "a2303"
         strSort = " order by a2302 desc,A2301"  'add by sonia 2015/11/25 加排序條件
      Case strCon2 '智權人員
         strCondition = "a2303"
         strSort = " order by a2303,a2302 desc,A2301"  'add by sonia 2015/11/25 加排序條件
      Case strCon3 '繳款日期
         strCondition = "a2302"
         strSort = " order by a2302 desc,A2301"  'add by sonia 2015/11/25 加排序條件
      Case strCon4 '繳收據日
         strCondition = "a2308"
         strSort = " order by decode(A2308,NULL,0,1),A2308 desc,A2301"  'add by sonia 2015/11/25 加排序條件
      'Add by Amy 2018/02/02
      Case strCon5 '銀存科目
         If Text1.Text <> "" Then
            strQ = " And A2305='" & Text1.Text & "'"
         End If
         '抓取銀行科目最大簽收單號之資料
         If Combo2 = Combo3 Then
            strQ = " And a2301=(Select  Max(a2301)  From  Acc230  Where a2322='" & Combo2 & "'" & strQ & ") "
         Else
            strQ = " And a2301 IN (Select  Max(a2301)  From  Acc230  Where a2322>='" & Combo2 & "' and a2322<='" & Combo3 & "'" & strQ & " Group by a2322) "
         End If
         strSort = " Order by a2322"
      Case MsgText(31) '全部
         strSort = " order by a2302 desc,A2301"  'add by sonia 2015/11/25 加排序條件
      Case Else
         Exit Sub
   End Select
   
   'Modified by Lydia 2015/11/17 +收款類別
'   strSql = "select A2301,A2302,ST02 A2303,A2308,CU04,NVL(A2306,0)+NVL(A2317,0)+NVL(A2318,0)+NVL(A2319,0)+NVL(A2320,0) Amount" & _
'            " from acc230, STAFF, customer where ST01(+)=A2303" & _
'            " AND CU01(+)=SUBSTR(A2304,1,8) AND CU02(+)=SUBSTR(A2304,9,1)"
   'Modify by Amy 2018/02/02 +顯示銀存科目
   'Modified by Morgan 2023/1/7 客戶名稱中->英->日
   strSql = "select A2301,A2302,ST02 A2303,A2308,Nvl(CU04,nvl(RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90),CU06)) CU04,A2322,SUBSTR(DECODE(SIGN(A2306),1,'支票')||DECODE(SIGN(A2317),1,'現金')||DECODE(SIGN(A2318),1,'銀存')||DECODE(SIGN(A2319),1,'暫收')||DECODE(SIGN(A2320),1,'其他'),1,2) Type" & _
            ",NVL(A2306,0)+NVL(A2317,0)+NVL(A2318,0)+NVL(A2319,0)+NVL(A2320,0) Amount" & _
            " from acc230, STAFF, customer where ST01(+)=A2303" & strQ & _
            " AND CU01(+)=SUBSTR(A2304,1,8) AND CU02(+)=SUBSTR(A2304,9,1)"
   '不是選全部且不是選銀存科目
   If Combo1 <> MsgText(31) And Combo1 <> strCon5 Then
      If Combo2 <> MsgText(601) Then
         If Combo1 = strCon1 Or Combo1 = strCon2 Then
            strSql = strSql & " and " & strCondition & ">='" & Combo2 & "'"
         Else
            strSql = strSql & " and " & strCondition & ">=" & Val(Combo2)
         End If
      End If
      If Combo3 <> MsgText(601) Then
         If Combo1 = strCon1 Or Combo1 = strCon2 Then
            strSql = strSql & " and " & strCondition & "<='" & Combo3 & "'"
         Else
            strSql = strSql & " and " & strCondition & "<=" & Val(Combo3)
         End If
      End If
   End If
   'end 2018/02/02
   '智權人員(未繳收據)
   If Combo1 = strCon1 Then
      strSql = strSql & " AND A2308 IS NULL"
      'Added by Morgan 2014/10/9
      strSql = strSql & " and (a2321 is null or to_char(a2321,'hh24miss')<>'000000')"
      'end 2014/10/9
   End If
   'Modified by Lydia 2015/11/17 改為選擇所別
'   If pub_strUserOffice <> "1" Then 'Added by Morgan 2015/6/17 改北所人員可看全部
'      strSql = strSql & " and A2305='" & pub_strUserOffice & "'"
'   End If
   If Text1.Text <> "" Then
      'Modified by Morgan 2023/2/7 所別:輸入人員、智權人員、1911-1913都要
      'strSql = strSql & " and A2305='" & Text1.Text & "'"
      strSql = strSql & " and exists(select * from staff where st01=A2303 and ('" & Text1.Text & "' in (st06,A2305)"
      If Text1.Text = "2" Then
         strSql = strSql & " or A2322='1911')"
      ElseIf Text1.Text = "3" Then
         strSql = strSql & " or A2322='1912')"
      ElseIf Text1.Text = "4" Then
         strSql = strSql & " or A2322='1913')"
      Else
         strSql = strSql & ")"
      End If
      strSql = strSql & ")"
      'end 2023/2/7
   End If
   'modify by sonia 2015/11/25 淑芳要求以繳款日期由大至小排序
   'strSql = strSql & " order by decode(A2308,NULL,0,1),A2301"
   strSql = strSql & strSort
   'end 2015/11/25
   adoadodc1.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
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
   adoadodc1.Open "select * from acc230 where rownum<1 order by 1 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Added by Lydia 2015/11/17
Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub
Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1.Text <> "" Then
      If Val(Text1) < 0 And Val(Text1) > 4 Then
         MsgBox "請輸入1-4", , MsgText(5)
         Cancel = True
         Text1.SetFocus
      End If
   End If
End Sub
