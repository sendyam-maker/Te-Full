VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc1193 
   AutoRedraw      =   -1  'True
   Caption         =   "銷帳退費查詢"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      Height          =   312
      Left            =   5748
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
      Height          =   312
      Left            =   2628
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
      Height          =   312
      Left            =   228
      TabIndex        =   0
      Top             =   240
      Width           =   2052
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1193.frx":0000
      Height          =   4275
      Left            =   135
      TabIndex        =   3
      Top             =   660
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   7541
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "銷帳退費資料"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "a0s01"
         Caption         =   "銷帳退費單號"
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
      BeginProperty Column01 
         DataField       =   "a0s03"
         Caption         =   "銷退日期"
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
      BeginProperty Column02 
         DataField       =   "a0s02"
         Caption         =   "收據單號"
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
         DataField       =   "a0s10"
         Caption         =   "轉出單號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "NAME"
         Caption         =   "收據抬頭或客戶名稱"
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
         DataField       =   "AMOUNT"
         Caption         =   "銷退金額"
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
         DataField       =   "SALES"
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
      BeginProperty Column07 
         DataField       =   "a0s18"
         Caption         =   "備註"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
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
         Size            =   344
         BeginProperty Column00 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   7065.071
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   228
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
      Left            =   -12
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
      Left            =   5508
      TabIndex        =   5
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
      Left            =   2388
      TabIndex        =   4
      Top             =   240
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc1193"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoacc0s0 As New ADODB.Recordset
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
   strCon1 = "銷帳退費單號"
   strCon2 = "收據單號"
   strCon3 = "銷退日期"
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
   If Adodc1.Recordset.RecordCount <> 0 Then
      strItemNo = Adodc1.Recordset.Fields("a0s01").Value
      If IsNull(Adodc1.Recordset.Fields("a0s02").Value) Then
         strCon1 = MsgText(601)
      Else
         strCon1 = Adodc1.Recordset.Fields("a0s02").Value
      End If
   Else
      strItemNo = MsgText(601)
      strCon1 = MsgText(601)
   End If
   StatusClear
   tool1_enabled
   Frmacc1190.Enabled = True
   Frmacc1190.Show
   Set Frmacc1193 = Nothing
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
         strCondition = "a0s01"
      Case strCon2
         strCondition = "a0s02"
      Case strCon3
         strCondition = "a0s03"
   End Select
   If strCondition = MsgText(601) Then
      Exit Sub
   End If
   Combo2.Clear
   Combo3.Clear
   adoacc0s0.CursorLocation = adUseClient
   adoacc0s0.Open "select distinct " & strCondition & " from acc0s0 order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc0s0.EOF = False
      If IsNull(adoacc0s0.Fields(0).Value) = False Then
         Combo2.AddItem adoacc0s0.Fields(0).Value
         Combo3.AddItem adoacc0s0.Fields(0).Value
      End If
      adoacc0s0.MoveNext
   Loop
   adoacc0s0.Close
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Acc0s0Query
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  票據資料查詢
'
'*************************************************
Private Sub Acc0s0Query()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   Select Case Combo1
      Case strCon1
         strCondition = "a0s01"
      Case strCon2
         strCondition = "a0s02"
      Case strCon3
         strCondition = "a0s03"
      Case MsgText(31)
         '2005/10/25 MODIFY BY SONIA加欄位
         'adoadodc1.Open "select * from acc0s0 order by a0s01 asc", adoTaie, adOpenStatic, adLockReadOnly
         adoadodc1.Open "select A0S01,A0S03,A0S02,A0S10,A0K04 AS NAME,A0S05+A0S06+A0S07 AS AMOUNT,ST02 AS SALES,A0S18 from acc0s0,ACC0K0,STAFF          where A0S02=A0K01 AND A0K20=ST01(+) UNION " & _
                        "select A0S01,A0S03,A0S02,A0S10,CU04  AS NAME,A0S05+A0S06+A0S07 AS AMOUNT,ST02 AS SALES,A0S18 from acc0s0,ACC0T0,STAFF,CUSTOMER where AND A0S02=A0T01 AND A0T05=ST01(+) AND SUBSTR(A0T06,1,8)=CU01(+) AND SUBSTR(A0T06,9,1)=CU02(+) " & _
                        " order by a0s01 asc", adoTaie, adOpenStatic, adLockReadOnly
         Adodc1.Recordset.Requery
         Exit Sub
      Case Else
         Exit Sub
   End Select
   If Combo3 = MsgText(601) Then
      If Combo1 = strCon3 Then
         '2005/10/25 MODIFY BY SONIA加欄位
         'adoadodc1.Open "select * from acc0s0 where " & strCondition & " = " & Val(Combo2) & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         adoadodc1.Open "select A0S01,A0S03,A0S02,A0S10,A0K04 AS NAME,A0S05+A0S06+A0S07 AS AMOUNT,ST02 AS SALES,A0S18 from acc0s0,ACC0K0,STAFF          where " & strCondition & " = " & Val(Combo2) & " AND A0S02=A0K01 AND A0K20=ST01(+) UNION " & _
                        "select A0S01,A0S03,A0S02,A0S10,CU04  AS NAME,A0S05+A0S06+A0S07 AS AMOUNT,ST02 AS SALES,A0S18 from acc0s0,ACC0T0,STAFF,CUSTOMER where " & strCondition & " = " & Val(Combo2) & " AND A0S02=A0T01 AND A0T05=ST01(+) AND SUBSTR(A0T06,1,8)=CU01(+) AND SUBSTR(A0T06,9,1)=CU02(+) " & _
                        " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
      Else
         '2005/10/25 MODIFY BY SONIA加欄位
         'adoadodc1.Open "select * from acc0s0 where " & strCondition & " = '" & Combo2 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         adoadodc1.Open "select A0S01,A0S03,A0S02,A0S10,A0K04 AS NAME,A0S05+A0S06+A0S07 AS AMOUNT,ST02 AS SALES,A0S18 from acc0s0,ACC0K0,STAFF          where " & strCondition & " = '" & Combo2 & "' AND A0S02=A0K01 AND A0K20=ST01(+) UNION " & _
                        "select A0S01,A0S03,A0S02,A0S10,CU04  AS NAME,A0S05+A0S06+A0S07 AS AMOUNT,ST02 AS SALES,A0S18 from acc0s0,ACC0T0,STAFF,CUSTOMER where " & strCondition & " = '" & Combo2 & "' AND A0S02=A0T01 AND A0T05=ST01(+) AND SUBSTR(A0T06,1,8)=CU01(+) AND SUBSTR(A0T06,9,1)=CU02(+) " & _
                        " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
      End If
   Else
      If Combo2 = MsgText(601) Then
         If Combo1 = strCon3 Then
            '2005/10/25 MODIFY BY SONIA加欄位
            'adoadodc1.Open "select * from acc0s0 where " & strCondition & " <= " & Val(Combo3) & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
             adoadodc1.Open "select A0S01,A0S03,A0S02,A0S10,A0K04 AS NAME,A0S05+A0S06+A0S07 AS AMOUNT,ST02 AS SALES,A0S18 from acc0s0,ACC0K0,STAFF          where " & strCondition & " <= " & Val(Combo3) & " AND A0S02=A0K01 AND A0K20=ST01(+) UNION " & _
                            "select A0S01,A0S03,A0S02,A0S10,CU04  AS NAME,A0S05+A0S06+A0S07 AS AMOUNT,ST02 AS SALES,A0S18 from acc0s0,ACC0T0,STAFF,CUSTOMER where " & strCondition & " <= " & Val(Combo3) & " AND A0S02=A0T01 AND A0T05=ST01(+) AND SUBSTR(A0T06,1,8)=CU01(+) AND SUBSTR(A0T06,9,1)=CU02(+) " & _
                            " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
        Else
            '2005/10/25 MODIFY BY SONIA加欄位
            'adoadodc1.Open "select * from acc0s0 where " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            adoadodc1.Open "select A0S01,A0S03,A0S02,A0S10,A0K04 AS NAME,A0S05+A0S06+A0S07 AS AMOUNT,ST02 AS SALES,A0S18 from acc0s0,ACC0K0,STAFF          where " & strCondition & " <= '" & Combo3 & "' AND A0S02=A0K01 AND A0K20=ST01(+) UNION " & _
                           "select A0S01,A0S03,A0S02,A0S10,CU04  AS NAME,A0S05+A0S06+A0S07 AS AMOUNT,ST02 AS SALES,A0S18 from acc0s0,ACC0T0,STAFF,CUSTOMER where " & strCondition & " <= '" & Combo3 & "' AND A0S02=A0T01 AND A0T05=ST01(+) AND SUBSTR(A0T06,1,8)=CU01(+) AND SUBSTR(A0T06,9,1)=CU02(+) " & _
                           " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         End If
      Else
         If Combo1 = strCon3 Then
            '2005/10/25 MODIFY BY SONIA加欄位
            'adoadodc1.Open "select * from acc0s0 where " & strCondition & " >= " & Val(Combo2) & " and " & strCondition & " <= " & Val(Combo3) & " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            adoadodc1.Open "select A0S01,A0S03,A0S02,A0S10,A0K04 AS NAME,A0S05+A0S06+A0S07 AS AMOUNT,ST02 AS SALES,A0S18 from acc0s0,ACC0K0,STAFF          where " & strCondition & " >= " & Val(Combo2) & " and " & strCondition & " <= " & Val(Combo3) & " AND A0S02=A0K01 AND A0K20=ST01(+) UNION " & _
                           "select A0S01,A0S03,A0S02,A0S10,CU04  AS NAME,A0S05+A0S06+A0S07 AS AMOUNT,ST02 AS SALES,A0S18 from acc0s0,ACC0T0,STAFF,CUSTOMER where " & strCondition & " >= " & Val(Combo2) & " and " & strCondition & " <= " & Val(Combo3) & " AND A0S02=A0T01 AND A0T05=ST01(+) AND SUBSTR(A0T06,1,8)=CU01(+) AND SUBSTR(A0T06,9,1)=CU02(+) " & _
                           " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         Else
            '2005/10/25 MODIFY BY SONIA加欄位
            'adoadodc1.Open "select * from acc0s0 where " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
            adoadodc1.Open "select A0S01,A0S03,A0S02,A0S10,A0K04 AS NAME,A0S05+A0S06+A0S07 AS AMOUNT,ST02 AS SALES,A0S18 from acc0s0,ACC0K0,STAFF          where " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' AND A0S02=A0K01 AND A0K20=ST01(+) UNION " & _
                           "select A0S01,A0S03,A0S02,A0S10,CU04  AS NAME,A0S05+A0S06+A0S07 AS AMOUNT,ST02 AS SALES,A0S18 from acc0s0,ACC0T0,STAFF,CUSTOMER where " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' AND A0S02=A0T01 AND A0T05=ST01(+) AND SUBSTR(A0T06,1,8)=CU01(+) AND SUBSTR(A0T06,9,1)=CU02(+) " & _
                           " order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
         End If
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
   adoadodc1.Open "select * from acc0s0 where a0s01 = '" & Combo2 & "' order by a0s01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub



