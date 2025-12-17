VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc1211 
   AutoRedraw      =   -1  'True
   Caption         =   "收據/請款單,發票資料查詢"
   ClientHeight    =   5160
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8856
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   11.4
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   8856
   Begin VB.CommandButton Command2 
      Caption         =   "收據抬頭修改"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6747
      TabIndex        =   3
      Top             =   600
      Width           =   1785
   End
   Begin VB.TextBox txtInvoice 
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
      Left            =   4440
      MaxLength       =   10
      TabIndex        =   1
      Top             =   210
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   210
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      Caption         =   "收據內容"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7440
      TabIndex        =   2
      Top             =   240
      Width           =   1092
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Bindings        =   "Frmacc1211.frx":0000
      Height          =   3945
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   8475
      _ExtentX        =   14944
      _ExtentY        =   6964
      _Version        =   393216
      Cols            =   17
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   $"Frmacc1211.frx":0015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   17
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   240
      Top             =   840
      Visible         =   0   'False
      Width           =   960
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
   Begin VB.Label lblPS 
      Caption         =   "列：N 未列印收據，Z 不列印收據，          # 已開INVOICE"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   408
      Left            =   3168
      TabIndex        =   8
      Top             =   576
      Width           =   3036
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "紅色資料代表有銷帳或銷退"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   2700
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收據號碼"
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
      Left            =   240
      TabIndex        =   5
      Top             =   210
      Width           =   975
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
      Caption         =   "智權發票號碼"
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
      Left            =   3030
      TabIndex        =   4
      Top             =   240
      Width           =   1395
   End
End
Attribute VB_Name = "Frmacc1211"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/16 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit

Public adoadodc1 As New ADODB.Recordset
Dim m_blnFirstShow As Boolean
Dim i  As Integer, j As Integer 'Add by Amy 2014/03/10


Private Sub Command1_Click()
   'Add by Morgan 2005/3/28
   If Adodc1.Recordset.State = adStateClosed Then Exit Sub

   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   'Add by Amy 2014/03/10 DataGrid不使用
    With grdDataList
    .Enabled = False
    For i = 1 To .Rows - 1
        .col = 0
        .row = i
        If .Text = "V" Then
            .col = 0
            .Text = ""
            'Modified by Lydia 2025/05/23 Index +1
            If Val(.TextMatrix(i, 12)) > 0 Or Val(.TextMatrix(i, 13)) > 0 Or Val(.TextMatrix(i, 14)) > 0 Or Val(.TextMatrix(i, 15)) > 0 Then
            'end 2025/05/23
                For j = 0 To .Cols - 1
                    .col = j
                    .CellBackColor = &H8080FF
                Next j
            Else
                For j = 0 To .Cols - 1
                    .col = j
                    .CellBackColor = QBColor(15)
                Next j
            End If
            'Add by Amy 2014/03/10 +if Exit for
            If Not IsNull(grdDataList.Text) Then
                'modify by sonia 2013/12/12 因銷過帳加註'銷'故此處要取消
                'strItemNo = Adodc1.Recordset.Fields("a0k01").Value
                'Modified by Lydia 2016/09/05 抓Grid的收據號碼
                'strItemNo = Adodc1.Recordset.Fields("a0k01").Value
                strItemNo = "" & Trim(.TextMatrix(i, 4))
                'Mark by Amy 2014/03/10 拿掉顯示銷
                '   If InStr(strItemNo, "銷") > 0 Then
                '      strItemNo = Left(strItemNo, Len(strItemNo) - 1)
                '   End If
   
                strExitControl = MsgText(601)
                tool3_enabled
                Screen.MousePointer = vbHourglass
                Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
                Frmacc1210.Show
                Screen.MousePointer = vbDefault
                Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
                '   Unload Me
                'Exit Sub
                Exit For
            End If
        End If
    Next i
     .Enabled = True
    End With
    'end 2014/03/10
End Sub
Private Sub Command2_Click()
'Added by Lydia 2016/01/20 點選呼叫"收據抬頭修改"
   If grdDataList.Rows < 3 Then Exit Sub
   
   If Adodc1.Recordset.State = 1 Then
      If Adodc1.Recordset.RecordCount = 0 Then
         Exit Sub
      End If
      strItemNo = Adodc1.Recordset.Fields("a0k01").Value
      strTitle = Me.Name
      If Mid(strItemNo, 1, 1) = "E" Then
         tool14_enabled
         MenuDisabled
         Frmacc1140.Show
         Me.Enabled = False
      Else
         MsgBox "請點選收據/請款單資料..."
         strItemNo = ""
         strTitle = ""
      End If
   Else
      MsgBox "請先按 F12 查詢並點選單據資料..."
   End If
'end 2016/01/20
End Sub

'Add by Amy 2013/12/23
Private Sub Form_Activate()
   tool3_enabled
   strFormName = Name
End Sub
'end 2013/12/23

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyDefine KeyCode
End Sub

Private Sub Form_Load()

On Error GoTo flgErr

Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
    m_blnFirstShow = True
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   strExitControl = "Y" 'Added by Morgan 2014/1/2
   'Modify by Amy 2023/08/18 原:W8850 H5400
   Me.Width = 8950
   Me.Height = 5600
   'end 2023/08/18
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Text1 = "E"
   'Modify by Amy 2014/03/10
   'OpenTable
   SetGridWidth '
   'end 2014/03/10
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
   
   'Added by Lydia 2016/01/20 分所不可使用"收據抬頭修改"
   If pub_strUserOffice = "1" Then
      Command2.Visible = True
   Else
      Command2.Visible = False
   End If
   'end 2016/01/20
   
flgErr:
    If Err.Number <> 0 Then MsgBox Err.Source & "：" & Err.Description, vbCritical, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strExitControl = "Y" Then
      StatusClear
      strFormName = MsgText(601)
      KeyEnter vbKeyEscape
      MenuEnabled
      Set Frmacc1211 = Nothing
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
   '92.2.27 MODIFY BY SONIA 用A0K02時, FORM LOAD會很久
   'adoadodc1.Open "select * from acc0k0 where a0k02 >= " & Val(FCDate(MaskEdBox2.Text)) & " and a0k02 <= " & Val(FCDate(MaskEdBox3.Text)) & " and (a0k02 <> 0 and a0k02 is not null) order by a0k01 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.Open "select * from acc0k0 where a0k01 = '" & Text1 & "' order by a0k01 asc", adoTaie, adOpenStatic, adLockReadOnly
   '92.2.27 END
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number <> 0 Then MsgBox Err.Source & "：" & Err.Description, vbCritical, "OpenTable"
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
Dim strSql As String
'Add By Cheng 2003/06/16
Dim StrSQLa As String
Dim i, jj As Integer 'Add by Amy 2013/12/16

On Error GoTo Checking
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
    'Add By Cheng 2004/01/12
    '若非北所員工, 只能列印該所資料
    If pub_strUserOffice <> "1" Then
        strSql = strSql & " And ''||ST06='" & pub_strUserOffice & "' "
    End If
    'End
   
   If Text1 = "" Or Me.Text1.Text = "E" Then
      '2005/10/21 MODIFY BY SONIA
      'StrSQLa = "select a0k02, a0k11, st03, a0k01, a0k30, a0k03, a0k04, a0k06, a0k07, a0k30, Nvl(new.Service,0) As Service, Nvl(new.Tax,0) As Tax, a0k09 from acc0k0, staff, (select a1u02, sum(a1u07) as Service, sum(a1u09) as Tax from acc1u0 group by a1u02) new where a0k20 = st01 (+) and a0k01 = a1u02 (+) and (a0k02 <> 0 and a0k02 is not null)" & strSQL & " order by a0k01 asc"
      'Modified by Morgan 2011/11/24 是否合併改抓 a0j07
      'modify by sonia 2013/12/12 銷過帳加註a0k01
      'Modify by Amy 2014/01/13 改sql 修正智權發票號碼抓不到收據號碼造成錯誤並讓查無資料時能Refresh 資料
      'strSqla = "select a0k02, a0k11, st03, a0k01||decode(a0k10,null,null,'銷') a0k01,a0j07, a0k03, a0k04, a0J09, a0J10, Nvl(new.Service,0) As Service, Nvl(new.Tax,0) As Tax, a0k09 from acc0k0, ACC0J0, staff, (select a1u02, A1U03, sum(a1u07) as Service, sum(a1u09) as Tax from acc1u0 group by a1u02,A1U03) new where a0k20 = st01 (+) and (a0k02 <> 0 and a0k02 is not null) and a0k01 = a0J13 (+) and a0J13 = a1u02 (+) and a0J01 = a1u03 (+) " & strSql & " order by a0k01 asc"
      'Modify by Amy 2014/03/10
      'Modified by Lydia 2025/05/23 +列印收據 => decode(a0k40,null,a0k32,'#') a0k32
      StrSQLa = "select '',a0k02, a0k11, '' as st03, a0k01, decode(a0k40,null,a0k32,'#') a0k32,'' as axc01, a0k03, a0k04, 0 as a0J09, 0 as a0J10,'' as a0j07, 0 As Service, 0 As Tax, 0 as S_a1u08, 0 as S_a1u10, a0k09 from acc0k0 where a0k01 = '' order by a0k01 asc"
      '2005/10/21 END
      adoadodc1.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   Else
   'end 2014/01/13
      '2005/10/21 MODIFY BY SONIA
      'StrSQLa = "select a0k02, a0k11, st03, a0k01, a0k30, a0k03, a0k04, a0k06, a0k07, a0k30, Nvl(new.Service,0) As Service, Nvl(new.Tax,0) As Tax, a0k09 from acc0k0, staff, (select a1u02, sum(a1u07) as Service, sum(a1u09) as Tax from acc1u0 WHERE a1U02 = '" & Text1 & "' group by a1u02) new where a0k20 = st01 (+) and a0k01 = a1u02 (+) and a0k01 = '" & Text1 & "' " & strSQL
      'Modified by Morgan 2011/11/24 是否合併改抓 a0j07
      'modify by sonia 2013/12/12 銷過帳加註a0k01
      'Modify by Amy 2013/12/13  改sum(a1u07) 或sum(a1u09) 大於0 才顯示銷, + axc01,Sum(a1u08),Sum(a1u10),小計
      'StrSQLa = "select a0k02, a0k11, st03, a0k01||decode(a0k10,null,null,'銷') a0k01,a0j07, a0k03, a0k04, a0J09, a0J10, Nvl(new.Service,0) As Service, Nvl(new.Tax,0) As Tax, a0k09 from acc0k0, ACC0J0, staff, (select a1u02, A1U03, sum(a1u07) as Service, sum(a1u09) as Tax from acc1u0 WHERE a1U02 = '" & Text1 & "' group by a1u02,A1U03) new where a0k20 = st01 (+) and a0k01 = a0J13 (+) and a0J13 = a1u02 (+) and a0J01 = a1u03 (+) and a0k01 = '" & Text1 & "' " & strSql
      'Modify by Amy 2014/03/10 +空白欄並取消"銷"字 原:a0k01||Decode(Nvl(new.Service,0),0,Decode(Nvl(Tax,0),0,'','銷'),'銷') as a0k01
      'Modified by Lydia 2016/09/02 不論輸入母號或子號，都要同時帶出所有母子號資料
      'StrSQLa = "select '',a0k02, a0k11, st03, a0k01,axc01, a0k03, a0k04, a0J09, a0J10,a0j07, Nvl(new.Service,0) As Service, Nvl(new.Tax,0) As Tax, Nvl(S_a1u08,0) as S_a1u08, Nvl(S_a1u10,0) as S_a1u10, a0k09 from acc0k0, ACC0J0, staff, Acc431, (select a1u02, A1U03, sum(a1u07) as Service, sum(a1u09) as Tax, Sum(a1u08) as S_a1u08, Sum(a1u10) as S_a1u10 from acc1u0 WHERE a1U02 = '" & Text1 & "' group by a1u02,A1U03) new where a0k20 = st01 (+) and a0k01 = a0J13 (+) and a0J13 = a1u02 (+) and a0J01 = a1u03 (+) and a0k01 =axc02(+) and a0k01 = '" & Text1 & "' " & strSql
      'StrSQLa = StrSQLa & " Union all Select '',0,'小計',null,null,null,null,'合計：'||ltrim(to_char(sum(Nvl(a0J09,0))+sum(Nvl(a0J10,0))-sum(Nvl(Service,0))-sum(Nvl(Tax,0)),'999,999,999')),sum(Nvl(a0J09,0)),sum(Nvl(a0J10,0)),null,sum(Nvl(Service,0)) as Service, sum(Nvl(Tax,0)) as Tax, Sum(Nvl(S_a1u08,0)) as S_a1u08, Sum(Nvl(S_a1u10,0)) as S_a1u10,0 From ACC0J0 , (select a1u02, A1U03, sum(a1u07) as Service, sum(a1u09) as Tax, Sum(a1u08) as S_a1u08, Sum(a1u10) as S_a1u10 from acc1u0 WHERE a1U02 = '" & Text1 & "' group by a1u02,A1U03) new Where a0J13 = a1u02 (+) and a0J01 = a1u03 (+) and a0j13 = '" & Text1 & "' "
      '2005/10/21 END
      'Modified by Lydia 2025/05/23 +列印收據 => decode(a0k40,null,a0k32,'#') a0k32
      StrSQLa = "select '',a0k02, a0k11, st03, a0k01, decode(a0k40,null,a0k32,'#') a0k32 ,axc01, a0k03, a0k04, a0J09, a0J10,a0j07, Nvl(new.Service,0) As Service, Nvl(new.Tax,0) As Tax, Nvl(S_a1u08,0) as S_a1u08, Nvl(S_a1u10,0) as S_a1u10, a0k09 from acc0k0, ACC0J0, staff, Acc431, (select a1u02, A1U03, sum(a1u07) as Service, sum(a1u09) as Tax, Sum(a1u08) as S_a1u08, Sum(a1u10) as S_a1u10 from acc1u0 WHERE substr(a1U02,1,9) = '" & Mid(Text1, 1, 9) & "' group by a1u02,A1U03) new where a0k20 = st01 (+) and a0k01 = a0J13 (+) and a0J13 = a1u02 (+) and a0J01 = a1u03 (+) and a0k01 =axc02(+) and substr(a0k01,1,9) = '" & Mid(Text1, 1, 9) & "' " & strSql
      'Modified by Lydia 2025/05/23 '小計',null,null,null,null,'合計：'=> '小計',null,null,null,null,null,'合計：'
      StrSQLa = StrSQLa & " Union all Select '',0,'小計',null,null,null,null,null,'合計：'||ltrim(to_char(sum(Nvl(a0J09,0))+sum(Nvl(a0J10,0))-sum(Nvl(Service,0))-sum(Nvl(Tax,0)),'999,999,999')),sum(Nvl(a0J09,0)),sum(Nvl(a0J10,0)),null,sum(Nvl(Service,0)) as Service, sum(Nvl(Tax,0)) as Tax, Sum(Nvl(S_a1u08,0)) as S_a1u08, Sum(Nvl(S_a1u10,0)) as S_a1u10,0 From ACC0J0 , (select a1u02, A1U03, sum(a1u07) as Service, sum(a1u09) as Tax, Sum(a1u08) as S_a1u08, Sum(a1u10) as S_a1u10 from acc1u0 WHERE substr(a1U02,1,9) = '" & Mid(Text1, 1, 9) & "' group by a1u02,A1U03) new Where a0J13 = a1u02 (+) and a0J01 = a1u03 (+) and substr(a0j13,1,9) = '" & Mid(Text1, 1, 9) & "' "
      
      adoadodc1.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   End If
   'Modify by Amy 2014/03/10
   'Adodc1.Recordset.Requery
   Set Adodc1.Recordset = adoadodc1
   
   If Adodc1.Recordset.RecordCount = 0 Or Adodc1.Recordset.RecordCount = 1 Then 'Modify by Amy 2013/12/13 +判斷=1
      Adodc1.Recordset.Close
      grdDataList.Rows = 2
      grdDataList.Clear
      SetGridWidth
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
    'Add by Amy 2014/03/10
    Set grdDataList.Recordset = Adodc1.Recordset
    SetGridWidth
    RefreshGridData
   'end 2014/03/10
Checking:
   If Err.Number <> 0 Then MsgBox Err.Source & "：" & Err.Description, vbCritical, "AdodcRefresh"
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
            'Modify by Morgan 2005/3/28 加所別控制
            'AdodcRefresh
            'If Adodc1.Recordset.State = adStateOpen Then Adodc1.Recordset.Close
            'Modify by Amy 2014/03/10 不使用DataGrid
            'DataGrid1.Refresh
            'end 2014/03/10
            
            'Modify by Amy 2013/12/16 +智權發票號查詢
            If (Text1 <> "" And Text1 <> "E") Or txtInvoice <> "" Then
               If (Text1 = "" Or Text1 = "E") And txtInvoice <> "" Then
                 Text1 = GetAxc02(txtInvoice)
               End If
               'end 2013/12/16
               Erase strExc: strExc(1) = Text1
               If PUB_CheckCaseZone(strExc, pub_strUserOffice, "2") = True Then
                  AdodcRefresh
               End If
            Else
               AdodcRefresh
            End If
            '2005/3/28 end
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            MsgBox MsgText(181), , MsgText(5)
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   'Modify by Amy 2013/12/16 +智權發票號判斷
   '2005/10/21 ADD BY SONIA 因無A0K02之INDEX,故取消收據日期的條件
   If (Text1 = "" Or Me.Text1.Text = "E") And Trim(txtInvoice) = "" Then
      Exit Function
   End If
   '2005/10/21 END
   If Text1 <> MsgText(601) And Text1 <> MsgText(802) Then
      FormCheck = True
      Exit Function
   End If
   'Add by Amy 2013/12/16
   If txtInvoice <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   'end 2013/12/16
   FormCheck = False
End Function

'Add by Amy 2013/12/16 以發票號碼取得收據號碼
Private Function GetAxc02(ByVal strInvoice As String) As String
    Dim StrSqlB As String
    GetAxc02 = ""
    StrSqlB = "Select * From Acc431 Where axc01='" & strInvoice & "' "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, StrSqlB)
    If intI = 1 Then
        GetAxc02 = RsTemp.Fields("axc02")
    End If
End Function
'end 2013/12/16

'Add by Amy 2014/01/13
Private Sub txtInvoice_GotFocus()
    TextInverse txtInvoice
End Sub

Private Sub txtInvoice_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub
'end 2014/01/13

'Add by Amy 2014/03/10
Private Sub GrdDataList_Click()
With grdDataList
    .col = 0
    .row = .MouseRow
    If .Rows - 1 = .row Then Exit Sub
     .Visible = False
    If .row <> 0 Then
        If .Text = "V" Then
            .Text = ""
            'Modified by Lydia 2025/05/23 Index + 1
            If Val(.TextMatrix(.row, 12)) > 0 Or Val(.TextMatrix(.row, 13)) > 0 Or Val(.TextMatrix(.row, 14)) > 0 Or Val(.TextMatrix(.row, 15)) > 0 Then
            'end 2025/05/23
                For i = 0 To .Cols - 1
                    .col = i
                    .CellBackColor = &H8080FF
                Next i
            Else
                For i = 0 To .Cols - 1
                    .col = i
                    .CellBackColor = QBColor(15)
                Next i
            End If
        
        Else
            .Text = "V"
            For i = 0 To .Cols - 1
                .col = i
                .CellBackColor = &HFFC0C0
            Next i
        End If
    End If
    .Visible = True
End With
End Sub

'銷帳或銷退有數字以紅色標註(原DataGrid不使用)
Private Sub SetGridWidth()
    With grdDataList
        .FormatString = .FormatString
        .ColWidth(0) = 250
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(1) = 1140
        .ColAlignment(1) = flexAlignCenterCenter
        .ColWidth(2) = 600
        .ColAlignment(2) = flexAlignCenterCenter
        .ColWidth(3) = 810
        .ColAlignment(3) = flexAlignCenterCenter
        .ColWidth(4) = 1350
        .ColAlignment(4) = flexAlignCenterCenter
         'Addedby Lydia 2025/05/23 列印收據
        .ColWidth(5) = 300
        .ColAlignment(5) = flexAlignCenterCenter
        'Modified by Lydia 2025/05/23 後面Index+1
        .ColWidth(6) = 300
        .ColAlignment(6) = flexAlignLeftCenter
        .ColWidth(7) = 1335
        .ColAlignment(7) = flexAlignLeftCenter
        .ColWidth(8) = 3615
        .ColAlignment(8) = flexAlignLeftCenter
        .ColWidth(9) = 1515
        .ColAlignment(9) = flexAlignRightCenter
        .ColWidth(10) = 1515
        .ColAlignment(10) = flexAlignCenterCenter
        .ColWidth(11) = 1005
        .ColAlignment(11) = flexAlignRightCenter
        .ColWidth(12) = 1515
        .ColAlignment(12) = flexAlignRightCenter
        .ColWidth(13) = 1515
        .ColAlignment(13) = flexAlignRightCenter
        .ColWidth(14) = 1515
        .ColAlignment(14) = flexAlignRightCenter
        .ColWidth(15) = 1515
        .ColAlignment(15) = flexAlignRightCenter
        .ColWidth(16) = 1245
        .ColAlignment(16) = flexAlignLeftCenter
        'end 2025/05/23
    End With
End Sub

'資料的銷帳或銷退大於0則以紅色顯示
Private Sub RefreshGridData()
    With grdDataList
        If Adodc1.Recordset.RecordCount > 0 Then
            .Visible = False
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format(.TextMatrix(i, 1), DFormat)
                'Modified by Lydia 2025/05/23 index + 1
                .TextMatrix(i, 9) = Format(.TextMatrix(i, 9), DDollar2)
                .TextMatrix(i, 11) = Format(.TextMatrix(i, 11), DDollar2)
                .TextMatrix(i, 12) = Format(.TextMatrix(i, 12), DDollar2)
                .TextMatrix(i, 13) = Format(.TextMatrix(i, 13), DDollar2)
                .TextMatrix(i, 14) = Format(.TextMatrix(i, 14), DDollar2)
                .TextMatrix(i, 15) = Format(.TextMatrix(i, 15), DDollar2)
                .TextMatrix(i, 16) = Format(.TextMatrix(i, 16), DFormat)
                
                If .Rows - 1 <> i And (Val(.TextMatrix(i, 12)) > 0 Or Val(.TextMatrix(i, 13)) > 0 Or Val(.TextMatrix(i, 14)) > 0 Or Val(.TextMatrix(i, 15)) > 0) Then
                'end 2025/05/23
                    For j = 0 To .Cols - 1
                        grdDataList.row = i
                        grdDataList.col = j
                        grdDataList.CellBackColor = &H8080FF
                    Next j
                End If
            Next i
            .Visible = True
         End If
    End With
End Sub
'end 2014/03/10
