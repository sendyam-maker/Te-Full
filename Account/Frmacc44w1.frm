VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc44w1 
   AutoRedraw      =   -1  'True
   Caption         =   "代填繳款書客戶查詢"
   ClientHeight    =   5172
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5172
   ScaleWidth      =   8760
   Begin VB.Frame Frame2 
      Height          =   350
      Left            =   6360
      TabIndex        =   8
      Top             =   -50
      Width           =   2100
      Begin VB.OptionButton Option3 
         Caption         =   "模糊比對"
         Height          =   180
         Index           =   1
         Left            =   1050
         TabIndex        =   10
         Top             =   144
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton Option3 
         Caption         =   "字首比對"
         Height          =   180
         Index           =   0
         Left            =   72
         TabIndex        =   9
         Top             =   144
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdA49 
      Caption         =   "基本資料維護"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5730
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   4680
      Width           =   1905
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      ItemData        =   "Frmacc44w1.frx":0000
      Left            =   5730
      List            =   "Frmacc44w1.frx":0002
      TabIndex        =   3
      Top             =   336
      Width           =   2772
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      ItemData        =   "Frmacc44w1.frx":0004
      Left            =   2610
      List            =   "Frmacc44w1.frx":0006
      TabIndex        =   2
      Top             =   336
      Width           =   2772
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      ItemData        =   "Frmacc44w1.frx":0008
      Left            =   210
      List            =   "Frmacc44w1.frx":000A
      TabIndex        =   1
      Top             =   336
      Width           =   2052
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   210
      Top             =   450
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGrid1 
      Height          =   3900
      Left            =   210
      TabIndex        =   11
      Top             =   720
      Width           =   8292
      _ExtentX        =   14626
      _ExtentY        =   6879
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "選「收據抬頭」且起迄條件相同"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   216
      Left            =   210
      TabIndex        =   7
      Top             =   50
      Width           =   3192
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5496
      TabIndex        =   5
      Top             =   336
      Width           =   132
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "="
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2376
      TabIndex        =   4
      Top             =   336
      Width           =   132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "註：按ESC鍵，即可離開查詢，回上一畫面！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   300
      TabIndex        =   0
      Top             =   4740
      Width           =   5265
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc44w1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2022/3/15 Form2.0已修改(DataGrid1改字型)
'Create by Sindy 2017/6/19
Option Explicit

Public adoacc420 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Dim m_ConSQL As String 'Add By Sindy 2019/12/10 控制分所人員操作時,只可查詢該所資料
Dim i As Integer, strFieldN(), intWidth() 'Add by Amy 2025/02/20
Public stPreCon As String 'Add by Amy 2025/03/04 前畫面條件

Private Sub cmdA49_Click()
Dim rsA As New ADODB.Recordset
Dim strQueryData As String
   
   'Modify by Amy 2025/02/20 原DataGrid,選抬頭名稱查詢可能多筆
   'If Adodc1.Recordset.RecordCount > 0 Then strQueryData = Adodc1.Recordset.Fields("cu04").Value
   Call SetColor(, , strQueryData)
   'end 2025/02/20
   If Trim(strQueryData) <> "" Then
      '客戶檔
      strSql = "select cu01,cu02 from customer where '" & ChgSQL(strQueryData) & "'=cu04" & _
              " or '" & ChgSQL(strQueryData) & "'=cu05||' '||cu88||' '||cu89||' '||cu90" & _
              " or '" & ChgSQL(strQueryData) & "'=cu06"
      If rsA.State = adStateOpen Then rsA.Close
      rsA.CursorLocation = adUseClient
      rsA.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         'tool1_enabled
         Me.MousePointer = vbHourglass
         'MenuDisabled
         strUserLevel = Me.Name
         Frmacc21r0.SetParent Me 'Add By Sindy 2016/11/29
         Frmacc21r0.txtKey = rsA.Fields("cu01") & rsA.Fields("cu02")
         Frmacc21r0.Command1_Click
         Frmacc21r0.Show
         Me.Hide
         Me.MousePointer = vbDefault
      Else
         '收據抬頭
         strSql = "select a4201 from acc420 where a4201='" & ChgSQL(strQueryData) & "'"
         If rsA.State = adStateOpen Then rsA.Close
         rsA.CursorLocation = adUseClient
         rsA.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            'tool1_enabled
            Me.MousePointer = vbHourglass
            'MenuDisabled
            strUserLevel = Me.Name
            strCompanyNo = ChgSQL(strQueryData)
            Frmacc11p0.SetParent Me 'Add By Sindy 2016/11/29
            strSaveConfirm = ""
            Frmacc11p0.textA4201 = Trim(strQueryData)
            Frmacc11p0.bolCallMe = True 'Add By Sindy 2015/9/14
            Frmacc11p0.Command3_Click
            Frmacc11p0.Show
            Me.Hide
            Me.MousePointer = vbDefault
         End If
      End If
   Else
      MsgBox "請點選一筆資料做查詢！", vbCritical
   End If
   
   Set rsA = Nothing
End Sub

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
Dim strST06 As String
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 9000
   Me.Height = 5595
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   'Add By Sindy 2019/12/10 控制分所人員操作時,只可查詢該所資料
   strST06 = PUB_GetST06(strUserNum)
   m_ConSQL = ""
   If strST06 <> "1" Then
      m_ConSQL = " and st06='" & strST06 & "'"
   End If
   '2019/12/10 END
   
   strCon1 = "收據抬頭" 'Modify by Amy 2025/02/20 原:客戶名稱-瑞婷
   strCon2 = "客戶編號"
   strCon3 = "智權人員" 'Add By Sindy 2019/12/10
   If strST06 = "1" Then strCon4 = "所別" 'Add By Sindy 2019/12/10
   Combo1.AddItem MsgText(31)
   Combo1.AddItem strCon1
   Combo1.AddItem strCon2
   Combo1.AddItem strCon3 'Add By Sindy 2019/12/10
   If strST06 = "1" Then Combo1.AddItem strCon4 'Add By Sindy 2019/12/10
   Combo1 = MsgText(31)
   SelectScope
   OpenTable
   Call SetGridWidth(True)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   'MenuEnabled
   StatusClear
   'tool1_enabled
   tool3_enabled
   stPreCon = "" 'Add by Amy 2025/03/04
   Frmacc44w0.Enabled = True
   Frmacc44w0.Show
   Set Frmacc44w1 = Nothing
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
         strCondition = "cu04"
      Case strCon2
         strCondition = "cu01"
      'Add By Sindy 2019/12/10
      Case strCon3
         strCondition = "st01"
      Case strCon4
         strCondition = "st06"
      '2019/12/10 END
   End Select
   If strCondition = MsgText(601) Then
      Exit Sub
   End If
   Combo2.Clear
   Combo3.Clear
   adoacc420.CursorLocation = adUseClient
   adoacc420.Open "select distinct " & strCondition & " from acc420 order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc420.EOF = False
      If IsNull(adoacc420.Fields(0).Value) = False Then
         Combo2.AddItem adoacc420.Fields(0).Value
         Combo3.AddItem adoacc420.Fields(0).Value
      End If
      adoacc420.MoveNext
   Loop
   adoacc420.Close
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Screen.MousePointer = vbHourglass
         Acc020Query
         Screen.MousePointer = vbDefault
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  資料查詢
'
'*************************************************
Private Sub Acc020Query()
   Dim strWhere As String, strQ As String 'Add by Amy 2025/02/20
   
On Error GoTo Checking
   
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   Select Case Combo1
      Case strCon1
         strCondition = "cu04"
      Case strCon2
         strCondition = "cu01"
      'Add By Sindy 2019/12/10
      Case strCon3
         strCondition = "st01"
      Case strCon4
         strCondition = "st06"
      '2019/12/10 END
      Case MsgText(31) '全部
      'Modify by Amy 2025/02/20 原使用DataGird,下拉選單[客戶名稱]改為[收據抬頭]並增加字首比對及模糊比對
      '                                                     且客戶已給代填同意書檔案顯示Ｖ(加cu168/a4220)改存公司別(可多筆[,]區分)
'         adoadodc1.Open "select rownum cnt,cu01,cu04,st02,st06nm,st01,st06 from(" & _
'                        " select cu01||cu02 cu01,nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)) cu04,st02,DECODE(ST06,'1','北','2','中','3','南','4','高','其他') AS st06nm,st01,st06" & _
'                        " From customer,staff" & _
'                        " where cu168='Y' and cu13=st01(+)" & m_ConSQL & _
'                        " Union" & _
'                        " select '',a4201,st02,DECODE(ST06,'1','北','2','中','3','南','4','高','其他') AS st06nm,st01,st06" & _
'                        " from acc420,staff" & _
'                        " where a4220='Y' and a4206=st01(+)" & m_ConSQL & _
'                        ") order by cu01,cu04", adoTaie, adOpenStatic, adLockReadOnly
'         Adodc1.Recordset.Requery
'         Exit Sub
      Case Else
         Exit Sub
   End Select
'   If Combo3 = MsgText(601) Then
'      adoadodc1.Open "select rownum cnt,cu01,cu04,st02,st06nm,st01,st06 from(" & _
'                     " select cu01||cu02 cu01,nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)) cu04,st02,DECODE(ST06,'1','北','2','中','3','南','4','高','其他') AS st06nm,st01,st06" & _
'                     " From customer,staff" & _
'                     " where cu168='Y' and cu13=st01(+)" & m_ConSQL & _
'                     " Union" & _
'                     " select '',a4201,st02,DECODE(ST06,'1','北','2','中','3','南','4','高','其他') AS st06nm,st01,st06" & _
'                     " from acc420,staff" & _
'                     " where a4220='Y' and a4206=st01(+)" & m_ConSQL & _
'                     ") where " & strCondition & " = '" & Combo2 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      If Combo2 = MsgText(601) Then
'         adoadodc1.Open "select rownum cnt,cu01,cu04,st02,st06nm,st01,st06 from(" & _
'                        " select cu01||cu02 cu01,nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)) cu04,st02,DECODE(ST06,'1','北','2','中','3','南','4','高','其他') AS st06nm,st01,st06" & _
'                        " From customer,staff" & _
'                        " where cu168='Y' and cu13=st01(+)" & m_ConSQL & _
'                        " Union" & _
'                        " select '',a4201,st02,DECODE(ST06,'1','北','2','中','3','南','4','高','其他') AS st06nm,st01,st06" & _
'                        " from acc420,staff" & _
'                        " where a4220='Y' and a4206=st01(+)" & m_ConSQL & _
'                        ") where " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
'      Else
'         adoadodc1.Open "select rownum cnt,cu01,cu04,st02,st06nm,st01,st06 from(" & _
'                        " select cu01||cu02 cu01,nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)) cu04,st02,DECODE(ST06,'1','北','2','中','3','南','4','高','其他') AS st06nm,st01,st06" & _
'                        " From customer,staff" & _
'                        " where cu168='Y' and cu13=st01(+)" & m_ConSQL & _
'                        " Union" & _
'                        " select '',a4201,st02,DECODE(ST06,'1','北','2','中','3','南','4','高','其他') AS st06nm,st01,st06" & _
'                        " from acc420,staff" & _
'                        " where a4220='Y' and a4206=st01(+)" & m_ConSQL & _
'                        ") where " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
'      End If
'   End If
'   Adodc1.Recordset.Requery
   
   '全部
   If Combo1 = MsgText(31) Then
      'Modify by Amy 2025/03/04 +前畫面條件
      strWhere = stPreCon & " Order by cu01,cu04"
   'end 2025/03/04
   '選 收據抬頭 且起/迄 條件輸相同
   ElseIf Combo1 = strCon1 And Trim(Combo2) <> MsgText(601) And Trim(Combo3) <> MsgText(601) And Trim(Combo2) = Trim(Combo3) Then
      '字首或模糊比對
      If Option3(0).Value = True Then
         strWhere = "=1"
      Else
          strWhere = ">0"
      End If
      strWhere = "And InStr(" & strCondition & "," & CNULL(ChgSQL(Combo2)) & ")" & strWhere & " "
   '迄條件 為空
   ElseIf Combo3 = MsgText(601) Then
      strWhere = "And " & strCondition & " = '" & ChgSQL(Combo2) & "' "
   Else
      If Combo2 = MsgText(601) Then
         strWhere = "And " & strCondition & " <= '" & ChgSQL(Combo3) & "' "
      Else
          strWhere = "And " & strCondition & " >= '" & ChgSQL(Combo2) & "' And " & strCondition & " <= '" & ChgSQL(Combo3) & "' "
      End If
   End If
   If InStr(UCase(strWhere), "ORDER BY ") = 0 Then
      'Modify by Amy 2025/03/04 +前畫面條件
      strWhere = strWhere & stPreCon & " Order by " & strCondition & " asc"
   End If
   'Modify by Amy 2025/03/04 +代填方式(cu181/a4228)及前畫面條件
   strQ = "select ' ' as V,' ' as WTC_1,' ' as WTC_L,rownum cnt,cu01,cu04,Decode(cu181,'1','每筆代繳','2','單筆超過2000元','') as cu181,st02,st06nm,st01,st06,WTC,stID from(" & _
                        " select cu01||cu02 cu01,nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)) cu04,cu181,st02,DECODE(ST06,'1','北','2','中','3','南','4','高','其他') AS st06nm" & _
                        ",st01,st06,cu168 as WTC,cu11 as stID" & _
                        " From customer,staff" & _
                        " Where cu168 is not null and cu13=st01(+)" & m_ConSQL & _
                        " Union" & _
                        " select '',a4201,a4228,st02,DECODE(ST06,'1','北','2','中','3','南','4','高','其他') AS st06nm" & _
                        ",st01,st06,a4220 as WTC,a4202 as STID" & _
                        " From acc420,staff" & _
                        " Where a4220 is not null and a4206=st01(+)" & m_ConSQL & _
                        ") where 1=1 " & strWhere
   adoadodc1.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
   Set MSHFGrid1.Recordset = adoadodc1
   SetGridWidth
   If Adodc1.Recordset.RecordCount = 0 Then
      MsgBox MsgText(33), , MsgText(5)
   Else
      SetWTCData
   End If
   'end 2025/02/20
   Exit Sub
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   Resume
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   'Modify by Amy 2025/02/20 +WTC_1/WTC_L/WTC/stID
   'Modify by Amy 2025/03/03 代填同意書改為每月代填同意書公司別(可能多筆) 原:cu168='Y' /a4220='Y'
   'Modify by Amy 2025/03/04 +代填方式(cu181/a4228)
   adoadodc1.Open "select ' ' as V,' ' as WTC_1,' ' as WTC_L,rownum cnt,cu01,cu04,Decode(cu181,'1','每筆代繳','2','單筆收據稅額超過2000元','') as cu181,st02,st06nm,st01,st06,WTC,stID from(" & _
                  " select cu01||cu02 cu01,nvl(cu04,nvl(cu05||' '||cu88||' '||cu89||' '||cu90,cu06)) cu04,cu181,st02,DECODE(ST06,'1','北','2','中','3','南','4','高','其他') AS st06nm" & _
                  ",st01,st06,cu168 as WTC,cu11 as stID" & _
                  " From customer,staff" & _
                  " where cu168 is not null and cu13=st01(+)" & m_ConSQL & _
                  " Union" & _
                  " select '',a4201,a4228 as cu181,st02,DECODE(ST06,'1','北','2','中','3','南','4','高','其他') AS st06nm" & _
                  ",st01,st06,a4220 as WTC,a4202 as STID" & _
                  " from acc420,staff" & _
                  " where a4220 is not null and a4206=st01(+)" & m_ConSQL & _
                  ") where cu04='' order by cu01,cu04", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Add by Amy 2025/02/20
Private Sub SetGridWidth(Optional ByVal IsFirst As Boolean = False)
   If IsFirst = True Then
      ReDim strFieldN(11)
      ReDim intWidth(11)
          
      strFieldN = Array(" ", "智慧-同意書", "法律-同意書", "序號", "客戶編號", "客戶/抬頭名稱", "代填方式", "智權人員", "所別", "st01", "st06", "WTC", "stID")
      intWidth = Array(200, 650, 650, 520, 1050, 2500, 1000, 1000, 600, 0, 0, 0, 0)
      MSHFGrid1.Cols = UBound(strFieldN) + 1
   End If

   MSHFGrid1.row = 0
   For i = LBound(strFieldN) To UBound(strFieldN)
      MSHFGrid1.col = i
      MSHFGrid1.ColWidth(i) = intWidth(i)
      MSHFGrid1.Text = strFieldN(i)
      
      MSHFGrid1.CellFontName = "新細明體-ExtB"
      If InStr(strFieldN(i), "同意書") > 0 Then
         MSHFGrid1.CellFontSize = 9
      Else
         MSHFGrid1.CellFontSize = 11
      End If
      MSHFGrid1.CellFontBold = True
      MSHFGrid1.CellAlignment = flexAlignLeftCenter
   Next i
End Sub

'客戶已給代填同意書檔案顯示O
Private Sub SetWTCData()
   Dim i As Integer, intU1 As Integer, intU2 As Integer, intC(2) As Integer
   Dim stWTCCmp As String, stCName As String, stID As String
   
   intU1 = GetValue("智慧-同意書")
   intU2 = GetValue("法律-同意書")
   intC(0) = GetValue("WTC")
   intC(1) = GetValue("客戶/抬頭名稱")
   intC(2) = GetValue("stID")
   With MSHFGrid1
      For i = 1 To .Rows - 1
         stWTCCmp = .TextMatrix(i, intC(0))
         stCName = .TextMatrix(i, intC(1))
         stID = .TextMatrix(i, intC(2))
         If stWTCCmp <> MsgText(601) Then
            '智慧所
            If InStr(stWTCCmp, "1") > 0 Then
               If ChkWithholdingTaxConsent(0, Me.Name, "1", stCName) = True Then
                  .TextMatrix(i, intU1) = "O"
                  .ColAlignment(intU1) = flexAlignCenterCenter
               End If
            End If
            '法律所
            If InStr(stWTCCmp, "1") > 0 Then
               If ChkWithholdingTaxConsent(0, Me.Name, "L", stCName) = True Then
                  .TextMatrix(i, intU2) = "O"
                  .col = intU2
                  .ColAlignment(intU2) = flexAlignCenterCenter
               End If
            End If
         End If
      Next i
   End With
End Sub

Private Function GetValue(pFieldN As String) As Integer
   Dim jj As Integer
 
    For jj = LBound(strFieldN) To UBound(strFieldN)
       If UCase(strFieldN(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

Private Sub MSHFGrid1_Click()
   Dim intC1 As Integer, iCurRow As Integer, m_Color As Long
   
   With MSHFGrid1
      intC1 = GetValue("序號")
      iCurRow = .MouseRow
      If iCurRow = 0 Then Exit Sub
      '序號 為空不執行 ex:一進入此畫面不小心點到
      If .TextMatrix(iCurRow, intC1) = MsgText(601) Then Exit Sub
      
      intC1 = GetValue("智慧-同意書") - 1
      If .TextMatrix(iCurRow, intC1) = "V" Then
         .TextMatrix(iCurRow, intC1) = ""
         m_Color = QBColor(15) '設回
      Else
         .TextMatrix(iCurRow, intC1) = "V"
         m_Color = &HFFC0C0 '整列底 藍色
      End If
      SetColor iCurRow, m_Color
   End With
End Sub

Private Sub SetColor(Optional ByVal pRow As Integer = 0, Optional ByVal pColor As Long = 0, Optional ByRef stCName As String)
   Dim i As Integer, intC1 As Integer
   
   stCName = ""
   With MSHFGrid1
      intC1 = GetValue("智慧-同意書") - 1
      .Visible = False
      '按[基本資料維護]鈕
      If pRow = 0 And pColor = 0 Then
         For i = 1 To .Rows - 1
            If .TextMatrix(i, intC1) = "V" Then
               stCName = .TextMatrix(i, GetValue("客戶/抬頭名稱"))
               .TextMatrix(i, intC1) = ""
               pColor = QBColor(15) '設回
               pRow = i
               Exit For
            End If
         Next i
      End If
      
      If pRow > 0 Then
         .row = pRow
         For i = 0 To .Cols - 1
            .col = i
            .CellBackColor = pColor
         Next i
      End If
      .Visible = True
   End With
End Sub
'end 2025/02/20
