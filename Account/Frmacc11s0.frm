VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frmacc11s0 
   AutoRedraw      =   -1  'True
   Caption         =   "發票申報作業"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   5505
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3570
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
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
      Height          =   300
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1200
      Width           =   612
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   3
      Top             =   2040
      Width           =   5000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "執行(&E)"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   1560
      Width           =   5000
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1200
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
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
      Left            =   2760
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "是否重新產生申報資料            (Y:重新產生)"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1260
      Width           =   4425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "申報日期"
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
      Left            =   240
      TabIndex        =   8
      Top             =   780
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   -120
      Top             =   960
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "發票日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   300
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc11s0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/08 Form2.0已檢查 (無需修改的物件)
'2014/3/5 create By Sonia
Option Explicit

Public adoacc510 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim stra0807 As String               'J公司統一編號
Dim stra0808 As String               'J公司稅籍編號
Dim strCnt   As Integer              '寫入筆數
Dim m_i As Integer
Dim strTemp(1 To 16) As String
Dim TempFileName As String


Private Sub Command1_Click()
   adoacc510.CursorLocation = adUseClient
   adoacc510.Open "select nvl(a4111,0) a4111 from acc410 where a4101>=" & Val(Left(FCDate(MaskEdBox1.Text), 5)) & " and a4102<=" & Val(Left(FCDate(MaskEdBox2.Text), 5)) & "", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc510.RecordCount <> 0 Then
      If adoacc510.Fields("a4111").Value > 0 And Text2 <> "Y" Then
         adoacc510.Close
         Screen.MousePointer = vbDefault
         MsgBox "此期間發票已產生申報資料, 請輸入重新產生資料欄！", , MsgText(21)
         Exit Sub
      End If
   End If
   adoacc510.Close
      
   Transfer

End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5625
   Me.Height = 4545 'Modify by Amy 2023/10/06 原:4485
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath3)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   Text1 = "": Text2 = ""
   
   '發票日期預設上一期起日(前一個單月的1日)至上一期止日(前一個雙用的最後一日)
   If Mid(strSrvDate(1), 5, 2) Mod 2 > 0 Then
      MaskEdBox1.Text = TransDate(CompDate(1, -2, (Left(strSrvDate(1), 6) & "01")), 1)
      MaskEdBox1.Text = Mid(MaskEdBox1.Text, 1, 3) & "/" & Mid(MaskEdBox1.Text, 4, 2) & "/" & Mid(MaskEdBox1.Text, 6, 2)
      MaskEdBox2.Text = TransDate(CompDate(2, -1, (Left(strSrvDate(1), 6) & "01")), 1)
      MaskEdBox2.Text = Mid(MaskEdBox2.Text, 1, 3) & "/" & Mid(MaskEdBox2.Text, 4, 2) & "/" & Mid(MaskEdBox2.Text, 6, 2)
   Else
      MaskEdBox1.Text = TransDate(CompDate(1, -3, (Left(strSrvDate(1), 6) & "01")), 1)
      MaskEdBox1.Text = Mid(MaskEdBox1.Text, 1, 3) & "/" & Mid(MaskEdBox1.Text, 4, 2) & "/" & Mid(MaskEdBox1.Text, 6, 2)
      MaskEdBox2.Text = TransDate(CompDate(2, -1, CompDate(1, -1, (Left(strSrvDate(1), 6)) & "01")), 1)
      MaskEdBox2.Text = Mid(MaskEdBox2.Text, 1, 3) & "/" & Mid(MaskEdBox2.Text, 4, 2) & "/" & Mid(MaskEdBox2.Text, 6, 2)
   End If
   
   '預設系統日
   MaskEdBox3.Text = TransDate(strSrvDate(1), 1)
   MaskEdBox3.Text = Mid(MaskEdBox3.Text, 1, 3) & "/" & Mid(MaskEdBox3.Text, 4, 2) & "/" & Mid(MaskEdBox3.Text, 6, 2)
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = DFormat
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc11s0 = Nothing
End Sub

Private Sub Transfer()
Dim strText As String
Dim ff As Integer

On Error GoTo Checking
   
   Text1 = "": strCnt = 1  '流水號從1號起
   
   Screen.MousePointer = vbHourglass
   
   '抓J公司統一編號
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select a0807,a0808 from acc080 where a0801='J' ", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      stra0807 = "" & adoquery.Fields("a0807").Value
      stra0808 = "" & adoquery.Fields("a0808").Value
   Else
      stra0807 = ""
      stra0808 = ""
   End If
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   
   cnnConnection.BeginTrans
   
   '先刪除該期間的申報資料
   adoTaie.Execute "delete from acc510 where a5104 >= " & Val(Left(FCDate(MaskEdBox1.Text), 5)) & " and a5104 <= " & Val(Left(FCDate(MaskEdBox2.Text), 5)) & ""
   
   Text1 = "正在將銷項發票資料轉至申報資料檔......"
   ProgressBar1.Value = 0
   adoacc510.CursorLocation = adUseClient
   'add by sonia 2018/9/12 +acc0s0判斷銷退日期是否為當期銷退 (EK19738126)
   adoacc510.Open "select * from acc430,acc0s0 where a4302 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a4302 <= " & Val(FCDate(MaskEdBox2.Text)) & " and a0s26(+)=a4301 order by a4302, a4301", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc510.RecordCount <> 0 Then
      ProgressBar1.max = adoacc510.RecordCount
   End If
   Do While adoacc510.EOF = False
      DoEvents
      
      For m_i = 1 To 15
         strTemp(m_i) = ""
         Select Case m_i
            Case 12
               strTemp(m_i) = String(5, " ")
            Case 11, 13, 14, 15
               strTemp(m_i) = " "
         End Select
      Next m_i
      
     'modify by sonia 2021/5/17婉莘說當期艮要同時寫35及33二筆,故取消a4302< " & Val(FCDate(MaskEdBox1.Text)),MN31530044,以前資料都是手動加的
      ''非銷退
      'If IsNull(adoacc510.Fields("a4309")) Then            '格式代號
      '   'modify by sonia 2019/10/7 31->35
      '   strTemp(1) = "35"                                   '非銷退
      ''銷退
      'Else
      '   strTemp(1) = "33"                                   '銷退
      '   'add by sonia 2018/9/12 當期銷退才放33 (EK19738126)
      '   If Val(adoacc510.Fields("a0s03")) >= Val(FCDate(MaskEdBox2.Text)) Then
      '     'modify by sonia 2019/10/7 31->35
      '      strTemp(1) = "35"
      '   End If
      '   'end 2018/9/12
      'End If
       strTemp(1) = "35"
      'end 2021/5/17
      
      strTemp(2) = stra0808                                '申報營業人稅籍編號
      strTemp(3) = strCnt                                  '流水號
      strCnt = strCnt + 1
      strTemp(4) = Val(Left(adoacc510.Fields("a4302"), 5)) '申報年月
      strTemp(16) = Val(Left(adoacc510.Fields("a4302"), 5)) '發票所屬年月   2019/10/7 add by sonia
      '作廢
      If Val("" & adoacc510.Fields("a4308")) > 0 Then      '買受人統一編號
         strTemp(5) = ""                                      '作廢發票存空白
      '未作廢
      Else
         strTemp(5) = "" & adoacc510.Fields("a4303")       '未作廢存買受人統一編號
         If strTemp(5) = "00000000" Then strTemp(5) = ""      'add by sonia 2018/9/12 統一編號00000000者存空白
      End If
      strTemp(6) = stra0807                                '智權公司統一編號
      strTemp(7) = "" & adoacc510.Fields("a4301")          '發票號碼
      '作廢
      If Val("" & adoacc510.Fields("a4308")) > 0 Then      '銷售金額
         strTemp(8) = 0                                         '作廢發票存0
      '未作廢
      Else
         strTemp(8) = Val("" & adoacc510.Fields("a4304"))
      End If
      '作廢
      If Val("" & adoacc510.Fields("a4308")) > 0 Then      '課稅別
         strTemp(9) = "F"                                       '作廢發票存F
      '未作廢
      Else
         'modify by sonia 2019/6/28 +零稅率PU14798012
         'strTemp(9) = "1"                                       '未作廢存1
         If "" & adoacc510.Fields("a4323") = "Y" Then
            strTemp(9) = "2"                                    '零稅率存2
            strTemp(15) = "1"                                   '零稅率存1   'add by sonia 2019/7/11
         Else
            strTemp(9) = "1"                                    '未作廢存1
         End If
         'end 2019/6/28
      End If
      '作廢
      If Val("" & adoacc510.Fields("a4308")) > 0 Then      '營業稅額
         strTemp(10) = 0                                        '作廢發票存0
      '未作廢
      Else
         strTemp(10) = Val("" & adoacc510.Fields("a4305"))      '未作廢存營業稅額
      End If
      'modify by sonia 2019/10/7 +a5116(strTemp(16))
      adoTaie.Execute "insert into acc510 values (" & CNULL(strTemp(1)) & ", " & CNULL(strTemp(2)) & ", " & CNULL(strTemp(3)) & ", " & strTemp(4) & ", " & CNULL(strTemp(5)) & ", " & CNULL(strTemp(6)) & ", " & CNULL(strTemp(7)) & ", " & strTemp(8) & ", " & CNULL(strTemp(9)) & ", " & strTemp(10) & ", " & CNULL(strTemp(11)) & ", " & CNULL(strTemp(12)) & ", " & CNULL(strTemp(13)) & ", " & CNULL(strTemp(14)) & ", " & CNULL(strTemp(15)) & ", " & strTemp(16) & ")"
      
      ProgressBar1.Value = ProgressBar1.Value + 1
      adoacc510.MoveNext
   Loop
   adoacc510.Close
   
   Text1 = "正在將銷項發票跨月轉開資料轉至申報資料檔......"
   ProgressBar1.Value = 0
   adoacc510.CursorLocation = adUseClient
   adoacc510.Open "select * from acc430 where a4310 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a4310 <= " & Val(FCDate(MaskEdBox2.Text)) & " order by a4310, a4301", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc510.RecordCount <> 0 Then
      ProgressBar1.max = adoacc510.RecordCount
   End If
   Do While adoacc510.EOF = False
      DoEvents
      
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select * from acc460 where a4601 = " & CNULL(adoacc510.Fields("a4301")) & "", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount > 0 Then
         strSql = "update acc460 set a4606=" & Val(FCDate(MaskEdBox3.Text)) & " where a4601 = " & CNULL(adoacc510.Fields("a4301")) & ""
         adoTaie.Execute strSql
      End If
      
      For m_i = 1 To 15
         strTemp(m_i) = ""
         Select Case m_i
            Case 12
               strTemp(m_i) = String(5, " ")
            Case 11, 13, 14, 15
               strTemp(m_i) = " "
         End Select
      Next m_i
      
      strTemp(1) = "33"                                    '格式代號
      strTemp(2) = stra0808                                '申報營業人稅籍編號
      strTemp(3) = strCnt                                  '流水號
      strCnt = strCnt + 1
      strTemp(4) = Val(Left(adoacc510.Fields("a4310"), 5)) '申報年月(轉開日期)
      strTemp(16) = Val(Left(adoacc510.Fields("a4310"), 5)) '發票所屬年月   2019/10/7 add by sonia
      'end 2019/10/7
      strTemp(5) = "" & adoacc510.Fields("a4303")          '買受人統一編號
      If strTemp(5) = "00000000" Then strTemp(5) = ""         'add by sonia 2018/9/12 統一編號00000000者存空白
      strTemp(6) = stra0807                                '智權公司統一編號
      strTemp(7) = "" & adoacc510.Fields("a4301")          '發票號碼
      strTemp(8) = Val("" & adoquery.Fields("a4604"))      '銷售金額
      'modify by sonia 2019/6/28 +零稅率PU14798012
      'strTemp(9) = "1"                                     '課稅別
      If "" & adoacc510.Fields("a4323") = "Y" Then
         strTemp(9) = "2"                                  '零稅率存2
         strTemp(15) = "1"                                 '零稅率存1   'add by sonia 2019/7/11
      Else
         strTemp(9) = "1"                                  '其他存1
      End If
      'end 2019/6/28
      strTemp(10) = Val("" & adoquery.Fields("a4605"))     '營業稅額
      'modify by sonia 2019/10/7 +a5116(strTemp(16))
      adoTaie.Execute "insert into acc510 values (" & CNULL(strTemp(1)) & ", " & CNULL(strTemp(2)) & ", " & CNULL(strTemp(3)) & ", " & strTemp(4) & ", " & CNULL(strTemp(5)) & ", " & CNULL(strTemp(6)) & ", " & CNULL(strTemp(7)) & ", " & strTemp(8) & ", " & CNULL(strTemp(9)) & ", " & strTemp(10) & ", " & CNULL(strTemp(11)) & ", " & CNULL(strTemp(12)) & ", " & CNULL(strTemp(13)) & ", " & CNULL(strTemp(14)) & ", " & CNULL(strTemp(15)) & ", " & strTemp(16) & ")"
      
      adoquery.Close
      ProgressBar1.Value = ProgressBar1.Value + 1
      adoacc510.MoveNext
   Loop
   adoacc510.Close

   Text1 = "正在將銷退費資料轉至申報資料檔......"
   ProgressBar1.Value = 0
   adoacc510.CursorLocation = adUseClient
   'modify by sonia 2021/5/17婉莘說當期艮要同時寫35及33二筆,故取消a4302< " & Val(FCDate(MaskEdBox1.Text)),MN31530044,以前資料都是手動加的
   'adoacc510.Open "select * from acc0s0,acc430 where a0s03 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0s03 <= " & Val(FCDate(MaskEdBox2.Text)) & " and a0s26 is not null and a0s26=a4301(+) and a4302< " & Val(FCDate(MaskEdBox1.Text)) & " order by a0s03, a0s26", adoTaie, adOpenStatic, adLockReadOnly
   adoacc510.Open "select * from acc0s0,acc430 where a0s03 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0s03 <= " & Val(FCDate(MaskEdBox2.Text)) & " and a0s26 is not null and a0s26=a4301(+) order by a0s03, a0s26", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc510.RecordCount <> 0 Then
      ProgressBar1.max = adoacc510.RecordCount
   End If
   Do While adoacc510.EOF = False
      DoEvents
      
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select * from acc460,acc430 where a4601 = " & CNULL(adoacc510.Fields("a0s01")) & " and a4301 = " & CNULL(adoacc510.Fields("a0s26")) & "", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount > 0 Then
         strSql = "update acc460 set a4606=" & Val(FCDate(MaskEdBox3.Text)) & " where a4601 = " & CNULL(adoacc510.Fields("a4301")) & ""
         adoTaie.Execute strSql
      End If
      
      For m_i = 1 To 15
         strTemp(m_i) = ""
         Select Case m_i
            Case 12
               strTemp(m_i) = String(5, " ")
            Case 11, 13, 14, 15
               strTemp(m_i) = " "
         End Select
      Next m_i
      
      strTemp(1) = "33"                                    '格式代號
      strTemp(2) = stra0808                                '申報營業人稅籍編號
      strTemp(3) = strCnt                                  '流水號
      strCnt = strCnt + 1
      'modify by sonia 2018/5/14 婉莘說銷退資料的所屬年月改放發票號碼的所屬年月,但ACC510會重覆主KEY,請她再確認,YT19738091,2019/9/12因PU14798114抓銷退日期,婉莘申報時會錯誤,故10/7加欄位A5116
      strTemp(4) = Val(Left(adoacc510.Fields("a0S03"), 5)) '申報年月(銷退日期)
      strTemp(16) = Val(Left(adoquery.Fields("a4302"), 5)) '發票所屬年月   2019/10/7 add by sonia
      'end 2019/10/7
      strTemp(5) = "" & adoquery.Fields("a4303")           '買受人統一編號
      If strTemp(5) = "00000000" Then strTemp(5) = ""          'add by sonia 2018/9/12 統一編號00000000者存空白
      strTemp(6) = stra0807                                '智權公司統一編號
      strTemp(7) = "" & adoquery.Fields("a4301")           '發票號碼
      strTemp(8) = Val("" & adoquery.Fields("a4604"))      '銷售金額
      'modify by sonia 2019/6/28 +零稅率PU14798012
      'strTemp(9) = "1"                                     '課稅別
      If "" & adoquery.Fields("a4323") = "Y" Then
         strTemp(9) = "2"                                  '零稅率存2
         strTemp(15) = "1"                                 '零稅率存1   'add by sonia 2019/7/11
      Else
         strTemp(9) = "1"                                  '其他存1
      End If
      'end 2019/6/28
      strTemp(10) = Val("" & adoquery.Fields("a4605"))     '營業稅額
      'modify by sonia 2019/10/7 +a5116(strTemp(16))
      adoTaie.Execute "insert into acc510 values (" & CNULL(strTemp(1)) & ", " & CNULL(strTemp(2)) & ", " & CNULL(strTemp(3)) & ", " & strTemp(4) & ", " & CNULL(strTemp(5)) & ", " & CNULL(strTemp(6)) & ", " & CNULL(strTemp(7)) & ", " & strTemp(8) & ", " & CNULL(strTemp(9)) & ", " & strTemp(10) & ", " & CNULL(strTemp(11)) & ", " & CNULL(strTemp(12)) & ", " & CNULL(strTemp(13)) & ", " & CNULL(strTemp(14)) & ", " & CNULL(strTemp(15)) & ", " & strTemp(16) & ")"
      
      adoquery.Close
      ProgressBar1.Value = ProgressBar1.Value + 1
      adoacc510.MoveNext
   Loop
   adoacc510.Close

   Text1 = "正在將進項發票資料轉至申報資料檔......"
   ProgressBar1.Value = 0
   adoacc510.CursorLocation = adUseClient
   adoacc510.Open "select * from acc450 where a4503 >= '" & Val(FCDate(MaskEdBox1.Text)) & "' and a4503 <= '" & Val(FCDate(MaskEdBox2.Text)) & "' order by a4503, a4504", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc510.RecordCount <> 0 Then
      ProgressBar1.max = adoacc510.RecordCount
   End If
   Do While adoacc510.EOF = False
      DoEvents

      For m_i = 1 To 15
         strTemp(m_i) = ""
         Select Case m_i
            Case 12
               strTemp(m_i) = String(5, " ")
            Case 13, 14, 15
               strTemp(m_i) = " "
         End Select
      Next m_i
      
      strTemp(1) = "" & adoacc510.Fields("a4502")          '格式代號
      strTemp(2) = stra0808                                '申報營業人稅籍編號
      strTemp(3) = strCnt                                  '流水號
      strCnt = strCnt + 1
      strTemp(4) = Val(Left(adoacc510.Fields("a4503"), 5)) '所屬年月
      strTemp(16) = Val(Left(adoacc510.Fields("a4503"), 5)) '發票所屬年月   2019/10/7 add by sonia
      strTemp(5) = stra0807                                '智權公司統一編號
      strTemp(6) = "" & adoacc510.Fields("a4505")          '銷售人統一編號
      strTemp(7) = "" & adoacc510.Fields("a4504")          '發票號碼
      strTemp(8) = Val("" & adoacc510.Fields("a4507"))     '銷售金額
      strTemp(9) = "1"                                     '課稅別
      strTemp(10) = Val("" & adoacc510.Fields("a4508"))    '營業稅額
      strTemp(11) = Val("" & adoacc510.Fields("a4506"))    '扣抵代號
      'modify by sonia 2019/10/7 +a5116(strTemp(16))
      adoTaie.Execute "insert into acc510 values (" & CNULL(strTemp(1)) & ", " & CNULL(strTemp(2)) & ", " & CNULL(strTemp(3)) & ", " & strTemp(4) & ", " & CNULL(strTemp(5)) & ", " & CNULL(strTemp(6)) & ", " & CNULL(strTemp(7)) & ", " & strTemp(8) & ", " & CNULL(strTemp(9)) & ", " & strTemp(10) & ", " & CNULL(strTemp(11)) & ", " & CNULL(strTemp(12)) & ", " & CNULL(strTemp(13)) & ", " & CNULL(strTemp(14)) & ", " & CNULL(strTemp(15)) & ", " & strTemp(16) & ")"
      
      ProgressBar1.Value = ProgressBar1.Value + 1
      adoacc510.MoveNext
   Loop
   adoacc510.Close

   '更新發票號碼範圍檔的申報日期
   adoTaie.Execute "update acc410 set a4111= " & Val(FCDate(MaskEdBox3.Text)) & " where a4101>=" & Val(Left(FCDate(MaskEdBox1.Text), 5)) & " and a4102<=" & Val(Left(FCDate(MaskEdBox2.Text), 5)) & ""

   '抓當期申報資料寫入桌面文字檔ACC510.TXT
   Text1 = "正在將當期申報資料寫入桌面文字檔 " & stra0807 & ".TXT......"
   ProgressBar1.Value = 0
   adoacc510.CursorLocation = adUseClient
   adoacc510.Open "select * from acc510 where a5104 >= " & Val(Left(FCDate(MaskEdBox1.Text), 5)) & " and a5104 <= " & Val(Left(FCDate(MaskEdBox2.Text), 5)) & " order by to_number(a5103)", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc510.RecordCount <> 0 Then
      ProgressBar1.max = adoacc510.RecordCount
      
      If ff > 0 Then Close #ff
      ff = FreeFile
      TempFileName = PUB_Getdesktop & "\" & stra0807 & ".TXT"
      Open TempFileName For Output As ff
   End If
         
   Do While adoacc510.EOF = False
      DoEvents
      
      strTemp(1) = "" & adoacc510.Fields("a5101")
      strTemp(2) = "" & adoacc510.Fields("a5102")
      strTemp(3) = Right("0000000" & adoacc510.Fields("a5103"), 7)
      'modify by sonia 2019/10/7婉莘說銷退發票改抓發票所屬年月
      'strTemp(4) = Right("00000" & adoacc510.Fields("a5104"), 5)
      strTemp(4) = Right("00000" & adoacc510.Fields("a5116"), 5)
      'end 2019/10/7
      strTemp(5) = Left(CheckStr(adoacc510.Fields("a5105")) & "        ", 8)
      strTemp(6) = "" & adoacc510.Fields("a5106")
      strTemp(7) = "" & adoacc510.Fields("a5107")
      strTemp(8) = Right("000000000000" & adoacc510.Fields("a5108"), 12)
      strTemp(9) = "" & adoacc510.Fields("a5109")
      strTemp(10) = Right("0000000000" & adoacc510.Fields("a5110"), 10)
      strTemp(11) = "" & adoacc510.Fields("a5111")
      strTemp(12) = "" & adoacc510.Fields("a5112")
      strTemp(13) = "" & adoacc510.Fields("a5113")
      strTemp(14) = "" & adoacc510.Fields("a5114")
      strTemp(15) = "" & adoacc510.Fields("a5115")
      
      strText = ""
      For m_i = 1 To 15
          strText = strText & strTemp(m_i)
      Next m_i
      Print #ff, strText
      
      ProgressBar1.Value = ProgressBar1.Value + 1
      adoacc510.MoveNext
   Loop
   adoacc510.Close
   Close ff
   
   cnnConnection.CommitTrans
   
   Screen.MousePointer = vbDefault
   Text1 = "": Text2 = ""
   MsgBox "已產生發票申報資料, 請至桌面讀取 " & stra0807 & ".TXT 匯入資料！", , MsgText(21)
   
Checking:
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   If adoacc510.State = adStateOpen Then
      Text1 = "錯誤之發票號碼: " & strTemp(7)
      adoacc510.Close
   End If
   Screen.MousePointer = vbDefault
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Sub

Private Sub MaskEdBox3_Validate(Cancel As Boolean)
   If MaskEdBox3.Text = MsgText(601) Or MaskEdBox3.Text = MsgText(29) Then
      MsgBox Label3 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox3.SetFocus
      Exit Sub
   End If
   
   If DateCheck(MaskEdBox3.Text) = MsgText(603) Then
      MsgBox Label3 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox3.SetFocus
      Exit Sub
   End If
   
   If Val(FCDate(MaskEdBox3.Text)) < Val(strSrvDate(2)) Then
      MsgBox "申報日期不可小於系統日！"
      Cancel = True
      MaskEdBox3.SetFocus
      Exit Sub
   End If
   
   If ChkWorkDay(Val(DBDATE(FCDate(MaskEdBox3.Text)))) = False Then
      MsgBox "申報日期必須為工作日！"
      Cancel = True
   End If

End Sub

Private Sub Text2_GotFocus()
   TextInverse Me.Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> 89 Then
      KeyAscii = 0
   End If
End Sub


