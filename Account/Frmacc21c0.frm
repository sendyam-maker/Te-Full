VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc21c0 
   AutoRedraw      =   -1  'True
   Caption         =   "結匯匯率輸入"
   ClientHeight    =   3252
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   7872
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3252
   ScaleWidth      =   7872
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   6000
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
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
      Left            =   4080
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc21c0.frx":0000
      Height          =   2295
      Left            =   1020
      TabIndex        =   3
      Top             =   720
      Width           =   5175
      _ExtentX        =   9123
      _ExtentY        =   4043
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   20
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
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "a1a03"
         Caption         =   "幣別"
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
         DataField       =   "a1a04"
         Caption         =   "匯率"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "a1a12"
         Caption         =   "匯率議價編號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00000000"
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
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   984.189
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1860.095
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "匯率儲存"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6240
      TabIndex        =   4
      Top             =   720
      Width           =   972
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2400
      Top             =   600
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
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "~"
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
      Left            =   5760
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "作業日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3120
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "結匯日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc21c0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/12/03 Form2.0已檢查 (無需修改的物件)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc190 As New ADODB.Recordset
Public adoacc1a0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset

Private Sub Command2_Click()
   Screen.MousePointer = vbHourglass
   Acc190Save
   KeyEnter vbKeyEscape
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   tool1_enabled
   MenuDisabled
   Frmacc21d0.Show
   Screen.MousePointer = vbDefault
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
   Adodc1.Recordset.UpdateBatch
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Select Case DataGrid1.col
            Case 1
               SendKeys "{DOWN}"
            'Added by Lydia 2017/09/29
            Case 2
               If Trim(DataGrid1.Text) = "" Or Len(DataGrid1.Text) = 12 Then
                  SendKeys "{DOWN}"
               ElseIf Trim(DataGrid1) <> "" Then
               End If
            'end 2017/09/29
         End Select
   End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(135)
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8000
   Me.Height = 3700 'Modify by Amy 2023/08/18
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   MaskEdBox1.Mask = MsgText(601)
   MaskEdBox1.Text = CFDate(ACDate(ServerDate))
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = MsgText(601)
   MaskEdBox2.Text = CFDate(ACDate(ServerDate))
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = MsgText(601)
   MaskEdBox3.Text = CFDate(ACDate(ServerDate))
   MaskEdBox3.Mask = DFormat
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(135)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc21c0 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1a0 where a1a01 = " & Val(FCDate(MaskEdBox1.Text)) & " order by a1a03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
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
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1a0 where a1a01 = " & Val(FCDate(MaskEdBox1.Text)) & " order by a1a03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
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
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Acc1a0Save
         MaskEdBox1.SetFocus
   End Select
   KeyEnter KeyCode
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label1 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label1 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
End Sub

'*************************************************
'  儲存資料表(國外結匯匯率資料)
'
'*************************************************
Private Sub Acc1a0Save()
On Error GoTo Checking
   adoacc190.CursorLocation = adUseClient
   'Modified by Lydia 2017/09/30 華銀整批媒體RMB改CNY
   'adoacc190.Open "select a1903 from acc190, acc180 where a1901 = a1801 and a1908 is null group by a1903", adoTaie, adOpenStatic, adLockReadOnly
   'Modified by Lydia 2024/09/03 排除匯款方式6-抵帳 =>  and a1811<>'6'
   adoacc190.Open "select decode(a1917||a1903,'JRMB','" & J_RMB & "',a1903) a1903 from acc190, acc180 where a1901 = a1801 and a1908 is null and a1811<>'6' group by decode(a1917||a1903,'JRMB','" & J_RMB & "',a1903) ", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc190.EOF = False
      adoacc1a0.CursorLocation = adUseClient
      'Modified by Lydia 2017/09/21 + a1a12
      adoacc1a0.Open "select a1a03,a1a12 from acc1a0 where a1a01 = " & Val(FCDate(MaskEdBox1.Text)) & " and a1a03 = '" & adoacc190.Fields("a1903").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc1a0.RecordCount = 0 Then
         'Modified by Lydia 2017/09/21 + A1A12 匯率議價編號(華銀結匯媒體使用)
         adoTaie.Execute "insert into acc1a0 (a1a01, a1a02, a1a03, a1a05, a1a06, a1a07, a1a11,a1a12) values (" & Val(FCDate(MaskEdBox1.Text)) & ", " & Val(strSrvDate(2)) & ", " & _
                         "'" & adoacc190.Fields("a1903").Value & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strUserNum & "', " & Val(strSrvDate(2)) & ",'')"
      'Added by Lydia 2017/09/21 +檢查
      ElseIf "" & adoacc1a0.Fields("a1a12") <> "" Then
        strExc(0) = PUB_StringFilter(PUB_GetSimpleName("" & adoacc1a0.Fields("a1a12")))
        If Len(strExc(0)) <> 12 Then
           MsgBox adoacc1a0.Fields("a1a03") & "匯率議價編號請輸入12碼英數字!", vbCritical, "資料檢核"
        End If
      'end 2017/09/21
      End If
      adoacc1a0.Close
      adoacc190.MoveNext
   Loop
   AdodcRefresh
   adoacc190.Close
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  儲存資料表(國外結匯匯率資料)
'
'*************************************************
Private Sub Acc190Save()
On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoacc1a0.CursorLocation = adUseClient
   adoacc1a0.Open "select * from acc1a0 where a1a01 = " & Val(FCDate(MaskEdBox1.Text)) & "", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc1a0.EOF = False
      If IsNull(adoacc1a0.Fields("a1a04").Value) = False Then
         'Added by Lydia 2017/09/30 華銀整批媒體RMB改CNY
         If "" & adoacc1a0.Fields("a1a03").Value = J_RMB And J_RMB <> "RMB" Then
            'Modified by Lydia 2024/09/03 排除匯款方式6-抵帳
            'adoTaie.Execute "update acc190 set a1905 = a1904 * " & Val(adoacc1a0.Fields("a1a04").Value) & ", a1906 = " & Val(adoacc1a0.Fields("a1a04").Value) & " where a1903 = 'RMB' and a1917='J' and a1908 is null"
            '暫收款退費之匯率改抓原暫收款之匯率
            'adoTaie.Execute "update acc190 set a1905=(select a1904*a1205 from acc130,acc120 where a1902=a1301 and a1303=a1201), " & _
                            "a1906=(select a1205 from acc130,acc120 where a1902=a1301 and a1303=a1201)" & _
                            " where a1903 = 'RMB' and a1917='J' and a1908 is null and substr(a1902,1,1)='O'"
            adoTaie.Execute "update acc190 set a1905 = a1904 * " & Val(adoacc1a0.Fields("a1a04").Value) & ", a1906 = " & Val(adoacc1a0.Fields("a1a04").Value) & " where a1903 = 'RMB' and a1917='J' and a1908 is null and a1901 in (select a1801 from acc180 where a1801=a1901 and a1811<>'6') "
            adoTaie.Execute "update acc190 set a1905=(select a1904*a1205 from acc130,acc120 where a1902=a1301 and a1303=a1201), " & _
                            "a1906=(select a1205 from acc130,acc120 where a1902=a1301 and a1303=a1201)" & _
                            " where a1903 = 'RMB' and a1917='J' and a1908 is null and substr(a1902,1,1)='O' and a1901 in (select a1801 from acc180 where a1801=a1901 and a1811<>'6')"
            'end 2024/09/03
         Else
         'end 2017/09/30
            'Modified by Lydia 2024/09/03 排除匯款方式6-抵帳
            'adoTaie.Execute "update acc190 set a1905 = a1904 * " & Val(adoacc1a0.Fields("a1a04").Value) & ", a1906 = " & Val(adoacc1a0.Fields("a1a04").Value) & " where a1903 = '" & adoacc1a0.Fields("a1a03").Value & "' and a1908 is null"
            '92.6.30 add by sonia 暫收款退費之匯率改抓原暫收款之匯率
            'adoTaie.Execute "update acc190 set a1905=(select a1904*a1205 from acc130,acc120 where a1902=a1301 and a1303=a1201), " & _
                            "a1906=(select a1205 from acc130,acc120 where a1902=a1301 and a1303=a1201)" & _
                            " where a1903 = '" & adoacc1a0.Fields("a1a03").Value & "' and a1908 is null and substr(a1902,1,1)='O'"
            '92.6.30 end
            adoTaie.Execute "update acc190 set a1905 = a1904 * " & Val(adoacc1a0.Fields("a1a04").Value) & ", a1906 = " & Val(adoacc1a0.Fields("a1a04").Value) & " where a1903 = '" & adoacc1a0.Fields("a1a03").Value & "' and a1908 is null and a1901 in (select a1801 from acc180 where a1801=a1901 and a1811<>'6')"
            adoTaie.Execute "update acc190 set a1905=(select a1904*a1205 from acc130,acc120 where a1902=a1301 and a1303=a1201), " & _
                            "a1906=(select a1205 from acc130,acc120 where a1902=a1301 and a1303=a1201)" & _
                            " where a1903 = '" & adoacc1a0.Fields("a1a03").Value & "' and a1908 is null and substr(a1902,1,1)='O' and a1901 in (select a1801 from acc180 where a1801=a1901 and a1811<>'6')"
            'end 2024/09/03
         End If
      End If
      adoacc1a0.MoveNext
   Loop
   adoacc1a0.Close
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
      MsgBox Label2 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox Label2 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
End Sub

Private Sub MaskEdBox3_Validate(Cancel As Boolean)
   If MaskEdBox3.Text = MsgText(601) Or MaskEdBox3.Text = MsgText(29) Then
      MsgBox Label2 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox3.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox3.Text) = MsgText(603) Then
      MsgBox Label2 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox3.SetFocus
      Exit Sub
   End If
End Sub

'Added by Lydia 2017/09/21
Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
Dim strTmp As String
   'Memo by Lydia 2017/09/21 因為是直接寫到DB,若DB欄位v2(15)長度不放大一些,在輸入時超過欄位限制會發生系統錯誤訊息
   If DataGrid1.col = 2 Then
      strTmp = DataGrid1.Text
      If Len(strTmp) - 1 >= 12 Then
         KeyAscii = 0
         MsgBox "匯率議價編號不可超過12碼英數字!", vbCritical, "資料檢核"
         DataGrid1.Text = Mid(strTmp, 1, 12)
      Else
         KeyAscii = UpperCase(KeyAscii)
      End If
   End If
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
   'Added by Lydia 2024/06/14 台銀近幾月日幣匯率皆給到小數位後8碼，為使結匯金額正確，匯率欄位增加容納至小數位後8碼
   If ColIndex = 1 And DataGrid1.Text <> "" Then
      'Modified by Lydia 2024/06/25 Trunc對於4.467會換成4.46666669
      'strExc(0) = Trunc(DataGrid1.Text, 8)
      strExc(0) = DataGrid1.Text
      If InStr(strExc(0), ",") > 0 Then
         strExc(0) = Mid(strExc(0), 1, InStr(strExc(0), ",") + 8)
      End If
      DataGrid1.Text = strExc(0)
   End If
   'end 2024/06/14
   If ColIndex = 2 And DataGrid1.Text <> "" Then
      strExc(0) = UCase(PUB_StringFilter(PUB_GetSimpleName(DataGrid1.Text)))
      If Len(strExc(0)) <> 12 Then
         MsgBox "匯率議價編號請輸入12碼英數字!", vbCritical, "資料檢核"
      End If
      DataGrid1.Text = Mid(strExc(0), 1, 12)
   End If
   
End Sub
'end 2017/09/21
