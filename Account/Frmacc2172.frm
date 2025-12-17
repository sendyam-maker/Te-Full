VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc2172 
   AutoRedraw      =   -1  'True
   Caption         =   "付款日期調整"
   ClientHeight    =   4728
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   4248
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4728
   ScaleWidth      =   4248
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2172.frx":0000
      Height          =   3375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   3735
      _ExtentX        =   6583
      _ExtentY        =   5948
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "A1F01"
         Caption         =   "原付款日期"
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
      BeginProperty Column01 
         DataField       =   "A1F02"
         Caption         =   "新付款日期"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   960
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1740
      TabIndex        =   0
      Top             =   210
      Width           =   1245
      _ExtentX        =   2201
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
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
      Left            =   1740
      TabIndex        =   1
      Top             =   570
      Width           =   1245
      _ExtentX        =   2201
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
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
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "新付款日期"
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
      TabIndex        =   4
      Top             =   600
      Width           =   1395
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "原付款日期"
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
      TabIndex        =   3
      Top             =   270
      Width           =   1395
   End
End
Attribute VB_Name = "Frmacc2172"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Morgan 2023/3/30
Option Explicit
Public adoadodc1 As New ADODB.Recordset

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   FormShow
   RecordShow
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 4365
   Me.Height = 5170 'Modify by Amy 2023/08/18 原5000
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   OpenTable
   AdodcRefresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc2172 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1F0 order by a1F01 desc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示欄位資料(美金匯率資料表)
'
'*************************************************
Public Sub FormShow()
   FormClear
   If Not Adodc1.Recordset.EOF Then
      MaskEdBox1.Mask = MsgText(601)
      MaskEdBox1.Text = CFDate(Adodc1.Recordset.Fields("a1f01").Value)
      MaskEdBox1.Tag = Adodc1.Recordset.Fields("a1f01").Value
      MaskEdBox1.Mask = DFormat
      MaskEdBox2.Mask = MsgText(601)
      MaskEdBox2.Text = CFDate(Adodc1.Recordset.Fields("a1f02").Value)
      MaskEdBox2.Mask = DFormat
   End If
   FormEnabled
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1F0 order by a1F01 desc", adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount > 0 Then
      RecordShow
      If MaskEdBox1.Tag <> "" Then
         Adodc1.Recordset.Find "a1f01=" & MaskEdBox1.Tag
         If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
      End If
   End If
   FormShow
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = Adodc1.Recordset.Bookmark & MsgText(35) & Adodc1.Recordset.RecordCount
End Sub

Public Sub FormClear()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
End Sub

Public Sub FormEnabled()
   '新增
   If strSaveConfirm = MsgText(3) Then
      MaskEdBox1.Enabled = True
      MaskEdBox2.Enabled = True
   '修改
   ElseIf strSaveConfirm = MsgText(4) Then
      MaskEdBox1.Enabled = False
      MaskEdBox2.Enabled = True
   Else
      MaskEdBox1.Enabled = False
      MaskEdBox2.Enabled = False
   End If
End Sub

Public Function EditCheck() As Boolean
   If MaskEdBox1.Text = "___/__/__" Then
      If strSaveConfirm = MsgText(4) Then
         MsgBox "請先點選欲修改的資料！", vbExclamation
      Else
         MsgBox "請先點選欲刪除的資料！", vbExclamation
      End If
      Exit Function
   End If
      
   If strSaveConfirm = MsgText(4) Then
      If MaskEdBox2.Text < CFDate(strSrvDate(2)) Then
         If MsgBox(Label2 & "小於系統日，是否確定要修改？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
         End If
      End If
   End If
   EditCheck = True
End Function

Public Function SaveCheck() As Boolean
   If MaskEdBox1.Text = "___/__/__" Then
      MsgBox "請輸入" & Label1 & "！", vbExclamation
      MaskEdBox1.SetFocus
      Exit Function
   End If
   
   If MaskEdBox2.Text = "___/__/__" Then
      MsgBox "請輸入" & Label2 & "！", vbExclamation
      MaskEdBox2.SetFocus
      Exit Function
   End If
   
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label1 & MsgText(63), , MsgText(5)
      MaskEdBox1.SetFocus
      Exit Function
   End If
   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox Label2 & MsgText(63), , MsgText(5)
      MaskEdBox2.SetFocus
      Exit Function
   End If
   
   If MaskEdBox1.Text < CFDate(strSrvDate(2)) Then
      If MsgBox(Label1 & "小於系統日，是否確定要繼續？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Function
      End If
   End If
      
   'Removed by Morgan 2023/6/2
   'If MaskEdBox1.Text >= MaskEdBox2.Text Then
   '   MsgBox Label2 & "必須大於" & Label1 & "！", vbExclamation
   '   MaskEdBox2.SetFocus
   '   Exit Function
   'End If
   'end 2023/6/2
   
   If strSaveConfirm = MsgText(4) Then
      If Replace(MaskEdBox2.Text, "/", "") = Adodc1.Recordset.Fields("a1f02").Value Then
         MsgBox Label2 & "並未修改！", vbExclamation
         MaskEdBox2.SetFocus
         Exit Function
      End If
   End If
   
   If ChkWorkDay(FCDate(MaskEdBox2.Text) + 19110000) = False Then
      MsgBox Label2 & "請輸入工作日！", vbExclamation, "日期錯誤！"
      MaskEdBox2.SetFocus
      Exit Function
   End If
   
   SaveCheck = True
End Function

Public Function FormDelete() As Boolean
   Dim stSQL As String, intR As Integer
   Dim stA1F01 As String, stA1F02 As String
   
   stA1F01 = Replace(MaskEdBox1.Text, "/", "")
   stA1F02 = Replace(MaskEdBox2.Text, "/", "")
   
   stSQL = "select a1501 from acc150  where a1527=" & stA1F02 & " and nvl(a1520,0)=0 and a1507 is null"
   intR = 1
   Set RsTemp = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      MsgBox "下列帳單的付款日期已設定為 " & MaskEdBox2.Text & "，本記錄不可刪除！" & vbCrLf & vbCrLf & RsTemp.GetString, vbCritical
      Exit Function
   End If
   
On Error GoTo ErrHnd
   adoTaie.BeginTrans
   
   stSQL = "delete acc1f0 where a1f01=" & stA1F01 & " and a1f02=" & stA1F02
   adoTaie.Execute stSQL, intR
   
   adoTaie.CommitTrans
   FormDelete = True
   MsgBox "刪除成功！", vbInformation
   FormClear
   AdodcRefresh
   Exit Function
   
ErrHnd:
   adoTaie.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Public Function FormSave() As Boolean
   Dim stSQL As String, intR As Integer
   Dim stA1F01 As String, stA1F02 As String
   Dim stChkDate As String, stRcvr As String, stRcvrlst As String
   Dim stSubject As String, stContent As String
   Dim stAgent As String, stBNo As String, stCNo As String, stDlr As String, stAmt As String

   If SaveCheck = False Then Exit Function
   stA1F01 = Replace(MaskEdBox1.Text, "/", "")
   stA1F02 = Replace(MaskEdBox2.Text, "/", "")
   
   adoTaie.BeginTrans
   
On Error GoTo ErrHnd
   '新增
   If strSaveConfirm = MsgText(3) Then
      stSQL = "insert into acc1f0(a1f01,a1f02,a1f03,a1f04,a1f05)" & _
         " values(" & stA1F01 & "," & stA1F02 & ",'" & strUserNum & "'" & _
         ",to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'))"
      stChkDate = stA1F01
   '修改
   Else
      stSQL = "update acc1f0 set a1f02=" & stA1F02 & ",a1f06='" & strUserNum & "'" & _
         ",a1f07=to_char(sysdate,'yyyymmdd'),a1f08=to_char(sysdate,'hh24miss')" & _
         " where a1f01=" & stA1F01
      stChkDate = Adodc1.Recordset.Fields("a1f02").Value
   End If
   adoTaie.Execute stSQL, intR
   
   stSQL = "select distinct a1516 R1,nvl(a1519,a1516) R2,a1528 R3,a1503,a1501,axf03,a1505,a1506" & _
      " from acc150,acc151 where a1527=" & stChkDate & " and nvl(a1520,0)=0 and a1507 is null and axf01(+)=a1501" & _
      " order by 1,2,3,4,5,6"
   intR = 1
   Set RsTemp = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      stSubject = "CF急件付款日期調整通知確認，若有疑問請洽財務人員"
      With RsTemp
      Do While Not .EOF
         stRcvr = .Fields("R1")
         If .Fields("R2") <> .Fields("R1") Then stRcvr = stRcvr & ";" & .Fields("R2")
         If Not IsNull(.Fields("R3")) Then stRcvr = stRcvr & ";" & .Fields("R3")
         If stRcvr <> stRcvrlst Then
            If stRcvrlst <> "" Then
               stContent = stContent & "<TR><TD>" & stAgent & "</TD><TD>" & stCNo & "</TD><TD>" & stBNo & "</TD><TD>" & stDlr & "</TD><TD align=right>" & stAmt & "</TD></TR>"
               stContent = stContent & "</TABLE>"
               'Modified by Morgan 2024/12/25 +mc11(內容有超過mc08大小)
               stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc11)" & _
                     " values( '" & strUserNum & "','" & stRcvrlst & "',to_char(sysdate,'yyyymmdd')" & _
                     ",to_char(sysdate,'hh24miss'),'" & ChgSQL(stSubject) & "','" & ChgSQL(Left(stContent, 4000)) & "','" & ChgSQL(Mid(stContent, 4001)) & "')"
               cnnConnection.Execute stSQL, intR
               stBNo = ""
            End If
            stRcvrlst = stRcvr
            stContent = "&nbsp;您好，付款日期" & CFDate(stChkDate) & "已調整為 " & MaskEdBox2.Text & "，更動清單如下：" & vbCrLf & _
               "<table border=1 cellspacing=0 width=700>" & _
               "<TR style=""background:#E1E1E1""><TD>代理人名稱</TD><TD>本所案號</TD><TD>單據編號</TD><TD>幣別</TD><TD>金額</TD></TR>"
         End If
         If .Fields("a1501") <> stBNo Then
            If stBNo <> "" Then
               stContent = stContent & "<TR><TD>" & stAgent & "</TD><TD>" & stCNo & "</TD><TD>" & stBNo & "</TD><TD>" & stDlr & "</TD><TD align=right>" & stAmt & "</TD></TR>"
            End If
            
            stSQL = "update acc150 set a1527=" & stA1F02 & " where a1501='" & .Fields("a1501") & "' and a1527=" & stChkDate
            cnnConnection.Execute stSQL, intR
            
            stAgent = ""
            If Not IsNull(.Fields("a1503")) Then
               ClsPDGetAgent .Fields("a1503"), stAgent
            End If
            stBNo = .Fields("a1501")
            stCNo = .Fields("axf03")
            stDlr = .Fields("a1505")
            stAmt = Format(.Fields("a1506"), DDollar)
         Else
            If stCNo <> .Fields("axf03") Then
               If Right(stCNo, 3) <> "..." Then
                  stCNo = stCNo & "..."
               End If
            End If
         End If
         .MoveNext
      Loop
      stContent = stContent & "<TR><TD>" & stAgent & "</TD><TD>" & stCNo & "</TD><TD>" & stBNo & "</TD><TD align=right>" & stDlr & "</TD><TD align=right>" & stAmt & "</TD></TR>"
      stContent = stContent & "</TABLE>"
      'Modified by Morgan 2024/12/25 +mc11(內容有超過mc08大小)
      stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc11)" & _
            " values( '" & strUserNum & "','" & stRcvrlst & "',to_char(sysdate,'yyyymmdd')" & _
            ",to_char(sysdate,'hh24miss'),'" & ChgSQL(stSubject) & "','" & ChgSQL(Left(stContent, 4000)) & "','" & ChgSQL(Mid(stContent, 4001)) & "')"
      cnnConnection.Execute stSQL, intR
      End With
   End If
   
   adoTaie.CommitTrans
   FormSave = True
   MaskEdBox1.Tag = stA1F01
   AdodcRefresh
   PUB_SendMailCache
   Exit Function
   
ErrHnd:
   adoTaie.RollbackTrans
   MsgBox Err.Description, vbCritical
End Function
