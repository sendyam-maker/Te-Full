VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc2440 
   AutoRedraw      =   -1  'True
   Caption         =   "結匯明細表"
   ClientHeight    =   2100
   ClientLeft      =   4275
   ClientTop       =   1215
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2100
   ScaleWidth      =   5160
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
      Left            =   345
      TabIndex        =   4
      Top             =   1080
      Width           =   2565
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1320
      TabIndex        =   2
      Top             =   660
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
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
      Left            =   3240
      TabIndex        =   3
      Top             =   648
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   270
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   1575
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1572
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
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   1572
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
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "代理人"
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
      Left            =   360
      TabIndex        =   9
      Top             =   660
      Width           =   975
   End
   Begin VB.Label Label2 
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
      Left            =   3000
      TabIndex        =   8
      Top             =   660
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "結匯日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label7 
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
      Height          =   252
      Left            =   3000
      TabIndex        =   6
      Top             =   240
      Width           =   252
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc2440"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/2 日期欄已修改
Option Explicit
Public adoacc170 As New ADODB.Recordset
Public adocaseprogress As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoaccrpt204 As New ADODB.Recordset
Public adoaccrpt219 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim intCounter As Integer
Dim strAmount As String
Dim intLength As Integer
Dim strNo As String
Dim douAmount As Double
Dim douFAmount As Double
Dim strSql As String

Private Sub Command2_Click()
   If FormCheck = False Then
      MsgBox MsgText(195), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   strSql = ""
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = " and a1b03 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a1b03 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a1b02 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a1b02 <= '" & Text2 & "'"
   End If
   Select Case Combo2
      Case ComboItem(251)
         Accrpt204Delete
         ProduceData
         PrintData
      Case ComboItem(252)
         ProcessDetail
         PrintData1
   End Select
   FormClear
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5250
   Me.Height = 2500
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   Combo2.AddItem ComboItem(251)
'   Combo2.AddItem ComboItem(252)
   Combo2 = ComboItem(251)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc2440 = Nothing
End Sub
'Add by Morgan 2004/12/15 抓代理人抬頭
Private Function GetFAgentTitle(stFAgentNo As String) As String

   Dim stSQL As String
   
On Error GoTo ErrHnd

   '不管幣別抓任一筆
   stSQL = "Select * From ACC220 Where a2201='" & stFAgentNo & "' AND ROWNUM<2"
         
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         GetFAgentTitle = Trim("" & .Fields("a2203") & " " & .Fields("a2204") & " " & .Fields("a2205") & " " & .Fields("a2206"))
      Else
         stSQL = "Select * From acc1b0, acc190, ACC180 Where a1b02='" & stFAgentNo & "' and a1908 = a1b01 and a1801 = a1901 AND ROWNUM<2"
         CheckOC3
         .Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
         If .RecordCount > 0 Then
            GetFAgentTitle = "" & .Fields("a1810")
         End If
         If GetFAgentTitle = "" Then
            stSQL = "select * from fagent where fa01 = '" & Mid(stFAgentNo, 1, 8) & "' and fa02 = '" & Mid(stFAgentNo, 9, 1) & "'"
            CheckOC3
            .Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
               GetFAgentTitle = Trim("" & .Fields("fa05") & " " & .Fields("fa63") & " " & .Fields("fa64") & " " & .Fields("fa65"))
            End If
         End If
      End If
   End With
   
ErrHnd:

   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical

End Function

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim intPage As Integer
Dim intRecord As Integer
Dim strCurrency As String
Dim strDocNo As String
Dim strUNo As String
Dim StrSQLa As String
Dim m_strDomAmt2  As Double    '2011/8/19 add by sonia 國內收款金額

On Error GoTo Checking
   douAmount = 0
   douFAmount = 0
   strNo = ""
   strUNo = ""
   strCurrency = ""
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoaccrpt204.CursorLocation = adUseClient
   adoaccrpt204.Open "select * from accrpt204", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc170.CursorLocation = adUseClient
   'adoacc170.Open "select * from acc170, acc190, acc1c0, acc1b0 where a1709 = a1901 and a1702 = a1902 and a1702 = a1c03 and a1c01 = a1b01 and a1c02 = a1b02" & strSQL & " order by a1705 asc, a1901 asc, a1707 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Morgan 2006/3/16
   '避免資料重複先distinct acc1p0
   'StrSQLa = "select a1709, a1b03, a1705, a1907, a1707, axf03, axf02, a1903, a1907, a1b04, a1906, axf04, a1901, a1p07, a1701, a1p08, axf01, a1p03 from acc151, acc170, acc190, (select * from acc1p0, acc1b0 where a1p04 = a1b01||a1b02) new where axf01 = a1702 and axf01 = a1902 and axf03 = a1p17 and axf04 = a1p21 and a1908 = substr(a1p04, 1, length(a1p04) - 9)" & strSQL & " union " & _
   '              "select a1709, a1b03, a1705, a1907, a1707, axg03 as axf03, axg02 as axf02, a1903, a1907, a1b04, a1906, axg04 * (-1) as axf04, a1901, a1p07, a1701, a1p08, axg01 as axf01, a1p03 from acc161, acc170, acc190, (select * from acc1p0, acc1b0 where a1p04 = a1b01||a1b02) new where axg01 = a1702 and axg01 = a1902 and axg03 = a1p17 and axg04 = a1p21 and a1908 = substr(a1p04, 1, length(a1p04) - 9)" & strSQL & " order by a1705 asc, a1901 asc, a1707 asc, a1p03 asc"
   
   'Add by Morgan 2006/4/19
   '95.4.20以後匯票資料輸入時1p23改放"帳單編號+總收文號"
   If Val(FCDate(MaskEdBox1.Text)) > 950419 Then
      StrSQLa = "select a1709, a1b03, a1705, a1907, a1707, axf03, axf02, a1903, a1b04, a1906, axf04, a1901, a1p07, a1701, a1p08, axf01" & _
         " From acc1b0, ACC1P0, acc190, acc170, acc151" & _
         " where a1908(+)=a1b01 and a1702(+)=a1902 and axf01=a1902" & _
         " and a1p04(+) = a1b01||a1b02 and a1p17 is not null and a1p23=axf01||axf02 and axf04=a1p21" & strSql & _
         " Union" & _
         " select a1709, a1b03, a1705, a1907, a1707, axg03 as axf03, axg02 as axf02, a1903, a1b04, a1906, axg04 * (-1) as axf04, a1901, a1p07, a1701, a1p08, axg01 as axf01" & _
         " From acc1b0, ACC1P0, acc190, acc170, acc161" & _
         " where a1908(+)=a1b01 and a1702(+)=a1902 and axg01=a1902" & _
         " and a1p04(+) = a1b01||a1b02 and a1p17 is not null and a1p23=axg01||axg02 and axg04=a1p21" & strSql & _
         " order by a1705 asc, a1901 asc, a1707 asc"
   'Modify by Morgan 2006/3/30
   '95.4.1以後匯票資料輸入時1p23會放總收文號,改用新語法以避免資料重複
   ElseIf Val(FCDate(MaskEdBox1.Text)) > 950330 Then
      StrSQLa = "select a1709, a1b03, a1705, a1907, a1707, axf03, axf02, a1903, a1b04, a1906, axf04, a1901, a1p07, a1701, a1p08, axf01" & _
         " From acc1b0, ACC1P0, acc190, acc170, acc151" & _
         " where a1p04(+) = a1b01||a1b02 and a1p17 is not null" & _
         " and a1908(+)=a1b01 and a1702(+)=a1902 and axf01=a1902 and axf02=a1p23 and axf04=a1p21" & strSql & _
         " Union" & _
         " select a1709, a1b03, a1705, a1907, a1707, axg03 as axf03, axg02 as axf02, a1903, a1b04, a1906, axg04 * (-1) as axf04, a1901, a1p07, a1701, a1p08, axg01 as axf01" & _
         " From acc1b0, ACC1P0, acc190, acc170, acc161" & _
         " where a1p04(+) = a1b01||a1b02 and a1p17 is not null" & _
         " and a1908(+)=a1b01 and a1702(+)=a1902 and axg01=a1902 and axg02=a1p23 and axg04=a1p21" & strSql & _
         " order by a1705 asc, a1901 asc, a1707 asc"
   Else
      StrSQLa = "select a1709, a1b03, a1705, a1907, a1707, axf03, axf02, a1903, a1907, a1b04, a1906, axf04, a1901, a1p07, a1701, a1p08, axf01 from acc151, acc170, acc190, (select distinct  a1p04,a1p07, a1p08, a1p17, a1p21, a1b03, a1b04 from acc1p0, acc1b0 where a1p04 = a1b01||a1b02 " & strSql & ") new where axf03 = a1p17 and axf04 = a1p21 and axf01 = a1702 and axf01 = a1902 and a1908 = substr(a1p04, 1, length(a1p04) - 9) union " & _
                 "select a1709, a1b03, a1705, a1907, a1707, axg03 as axf03, axg02 as axf02, a1903, a1907, a1b04, a1906, axg04 * (-1) as axf04, a1901, a1p07, a1701, a1p08, axg01 as axf01 from acc161, acc170, acc190, (select distinct  a1p04,a1p07, a1p08, a1p17, a1p21, a1b03, a1b04 from acc1p0, acc1b0 where a1p04 = a1b01||a1b02 " & strSql & ") new where axg03 = a1p17 and axg04 = a1p21 and axg01 = a1702 and axg01 = a1902 and a1908 = substr(a1p04, 1, length(a1p04) - 9)" & " order by a1705 asc, a1901 asc, a1707 asc"
   
   End If
   adoacc170.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc170.RecordCount = 0 Then
      adoacc170.Close
      adoaccrpt204.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   intPage = 0
   intRecord = 1
   Do While adoacc170.EOF = False
      If strNo <> (adoacc170.Fields("a1709").Value) Then
         If strNo <> "" Then
            adoaccrpt204.AddNew
            adoaccrpt204.Fields("r20401").Value = strUserNum
            adoaccrpt204.Fields("r20407").Value = intPage
            adoaccrpt204.Fields("r20408").Value = Mid(strNo, 1, 9)
            adoaccrpt204.Fields("r20413").Value = "小計:"
            adoaccrpt204.Fields("r20416").Value = strCurrency
            adoaccrpt204.Fields("r20417").Value = douAmount
            '取小數兩位
            adoaccrpt204.Fields("r20420").Value = Format(douFAmount, "0.00")
            'End
            adoaccrpt204.Fields("r20421").Value = intRecord
            adoaccrpt204.UpdateBatch
         End If
         intPage = intPage + 1
         intRecord = 1
         douAmount = 0
         douFAmount = 0
         strNo = adoacc170.Fields("a1709").Value
      End If
'      If strDocNo = adoacc170.Fields("axf01").Value Then
      adoaccrpt204.AddNew
      adoaccrpt204.Fields("r20401").Value = strUserNum
      adoaccrpt204.Fields("r20402").Value = ReportTitle(204)
      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
         adoaccrpt204.Fields("r20403").Value = Val(FCDate(MaskEdBox1.Text))
      Else
         adoaccrpt204.Fields("r20403").Value = Null
      End If
      If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
         adoaccrpt204.Fields("r20404").Value = Val(FCDate(MaskEdBox2.Text))
      Else
         adoaccrpt204.Fields("r20404").Value = Null
      End If
      adoaccrpt204.Fields("r20405").Value = strUserNum
      adoaccrpt204.Fields("r20406").Value = Val(strSrvDate(2))
      adoaccrpt204.Fields("r20407").Value = intPage
      If IsNull(adoacc170.Fields("a1709").Value) Then
         adoaccrpt204.Fields("r20408").Value = Null
      Else
         adoaccrpt204.Fields("r20408").Value = adoacc170.Fields("a1709").Value
      End If
      If IsNull(adoacc170.Fields("a1b03").Value) Then
         adoaccrpt204.Fields("r20409").Value = Null
      Else
         adoaccrpt204.Fields("r20409").Value = adoacc170.Fields("a1b03").Value
      End If
      If IsNull(adoacc170.Fields("a1705").Value) Then
         adoaccrpt204.Fields("r20410").Value = Null
         adoaccrpt204.Fields("r20411").Value = Null
      Else
         adoaccrpt204.Fields("r20410").Value = adoacc170.Fields("a1705").Value
         'Modify by Morgan 2004/12/15 改同水單抓法
         'adoaccrpt204.Fields("r20411").Value = FagentQuery(adoacc170.Fields("a1705").Value, 1)
         adoaccrpt204.Fields("r20411").Value = GetFAgentTitle(adoacc170.Fields("a1705").Value)
         If adoaccrpt204.Fields("r20411").Value = "" Then
            adoaccrpt204.Fields("r20411").Value = FagentQuery(adoacc170.Fields("a1705").Value, 2)
         End If
      End If
      '2006/2/21 MODIFY BY SONIA 收據抬頭原抓a1907改抓收據或客戶名稱,因一帳單多案號時收據抬頭可能不同
      'If IsNull(adoacc170.Fields("a1907").Value) Then
      '   adoaccrpt204.Fields("r20413").Value = Null
      'Else
      '   adoaccrpt204.Fields("r20413").Value = adoacc170.Fields("a1907").Value
      'End If
      If IsNull(GetA0K04("" & adoacc170.Fields("axf03").Value, "" & adoacc170.Fields("axf02").Value)) Then
         adoaccrpt204.Fields("r20413").Value = Null
      Else
         adoaccrpt204.Fields("r20413").Value = GetA0K04("" & adoacc170.Fields("axf03").Value, "" & adoacc170.Fields("axf02").Value)
      End If
      If IsNull(adoacc170.Fields("a1707").Value) Then
         adoaccrpt204.Fields("r20412").Value = Null
         adoaccrpt204.Fields("r20414").Value = Null
         adoaccrpt204.Fields("r20415").Value = 0
      Else
         adoaccrpt204.Fields("r20412").Value = adoacc170.Fields("axf03").Value
         If adoacc170.Fields("a1701").Value = "1" Then
'2011/8/19 modify by sonia
'            adocaseprogress.CursorLocation = adUseClient
'            StrSQLa = "select a0k04, a0l02, a0m02, nvl(a0k17, 0)+nvl(a0k18, 0) as Amount from caseprogress, acc0k0, acc0m0, acc0l0, acc1u0 where cp60 = a0k01 and a0k01 = a0m02 and a0m01 = a0l01 and cp60 = a1u02 (+) and cp09 = a1u03 (+) and cp01 = '" & Mid(adoacc170.Fields("axf03").Value, 1, Len(adoacc170.Fields("axf03").Value) - 9) & "' and cp02 = '" & Mid(adoacc170.Fields("axf03").Value, Len(adoacc170.Fields("axf03").Value) - 8, 6) & "' and cp03 = '" & Mid(adoacc170.Fields("axf03").Value, Len(adoacc170.Fields("axf03").Value) - 2, 1) & "' and cp04 = '" & Mid(adoacc170.Fields("axf03").Value, Len(adoacc170.Fields("axf03").Value) - 1, 2) & "' and cp09 = '" & adoacc170.Fields("axf02").Value & "' union " & _
'                      "select null as a0k04, a0y02 as a0l02, a0z02 as a0m02, nvl(a1k30, 0) as Amount from caseprogress, acc1k0, acc0z0, acc0y0 where cp60 = a1k01 and a1k01 = a0z02 and a0z01 = a0y01 and cp01 = '" & Mid(adoacc170.Fields("axf03").Value, 1, Len(adoacc170.Fields("axf03").Value) - 9) & "' and cp02 = '" & Mid(adoacc170.Fields("axf03").Value, Len(adoacc170.Fields("axf03").Value) - 8, 6) & "' and cp03 = '" & Mid(adoacc170.Fields("axf03").Value, Len(adoacc170.Fields("axf03").Value) - 2, 1) & "' and cp04 = '" & Mid(adoacc170.Fields("axf03").Value, Len(adoacc170.Fields("axf03").Value) - 1, 2) & "' and cp09 = '" & adoacc170.Fields("axf02").Value & "'"
'            adocaseprogress.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'            If adocaseprogress.RecordCount <> 0 Then
'               If IsNull(adocaseprogress.Fields(1).Value) Then
'                  adoaccrpt204.Fields("r20414").Value = Null
'               Else
'                  adoaccrpt204.Fields("r20414").Value = adocaseprogress.Fields(1).Value
'               End If
'               If IsNull(adocaseprogress.Fields(3).Value) Then
'                  adoaccrpt204.Fields("r20415").Value = 0
'               Else
'                  adoaccrpt204.Fields("r20415").Value = Val(adocaseprogress.Fields(3).Value)
'               End If
'            Else
'               adoaccrpt204.Fields("r20414").Value = Null
'               adoaccrpt204.Fields("r20415").Value = 0
'            End If
'            adocaseprogress.Close
            If GetA1l02("" & adoacc170.Fields("axf03").Value, "" & adoacc170.Fields("axf02").Value, 0, m_strDomAmt2) = "" Then
               adoaccrpt204.Fields("r20414").Value = Null
               adoaccrpt204.Fields("r20415").Value = 0
            Else
               adoaccrpt204.Fields("r20414").Value = GetA1l02("" & adoacc170.Fields("axf03").Value, "" & adoacc170.Fields("axf02").Value, 0, m_strDomAmt2)
               adoaccrpt204.Fields("r20415").Value = Val(m_strDomAmt2)
            End If
'2011/8/19 end
         End If
      End If
      If IsNull(adoacc170.Fields("a1903").Value) Then
         adoaccrpt204.Fields("r20416").Value = Null
         strCurrency = ""
      Else
         adoaccrpt204.Fields("r20416").Value = adoacc170.Fields("a1903").Value
         strCurrency = adoacc170.Fields("a1903").Value
      End If
      If IsNull(adoacc170.Fields("a1p07").Value) Or adoacc170.Fields("a1p07").Value = 0 Then
         If IsNull(adoacc170.Fields("a1p08").Value) Then
            adoaccrpt204.Fields("r20417").Value = 0
         Else
            adoaccrpt204.Fields("r20417").Value = Val(adoacc170.Fields("a1p08").Value) * (-1)
            douAmount = douAmount - Val(adoacc170.Fields("a1p08").Value)
         End If
      Else
         adoaccrpt204.Fields("r20417").Value = adoacc170.Fields("a1p07").Value
         douAmount = douAmount + Val(adoacc170.Fields("a1p07").Value)
      End If
      If IsNull(adoacc170.Fields("a1b04").Value) Then
         adoaccrpt204.Fields("r20418").Value = 0
      Else
         adoaccrpt204.Fields("r20418").Value = adoacc170.Fields("a1b04").Value
      End If
      If IsNull(adoacc170.Fields("a1906").Value) Then
         adoaccrpt204.Fields("r20419").Value = 0
      Else
         adoaccrpt204.Fields("r20419").Value = adoacc170.Fields("a1906").Value
      End If
      If IsNull(adoacc170.Fields("axf04").Value) Then
         adoaccrpt204.Fields("r20420").Value = 0
      Else
         If adoacc170.Fields("a1701").Value = "1" Then
            adoaccrpt204.Fields("r20420").Value = Val("" & adoacc170.Fields("axf04").Value)
            douFAmount = douFAmount + CDbl(Val("" & adoacc170.Fields("axf04").Value))
         Else
            adoaccrpt204.Fields("r20420").Value = Val("" & adoacc170.Fields("axf04").Value)
            douFAmount = douFAmount + CDbl(Val("" & adoacc170.Fields("axf04").Value))
         End If
      End If
      adoaccrpt204.Fields("r20421").Value = intRecord
      adoaccrpt204.UpdateBatch
      intRecord = intRecord + 1
        '每頁18行明細
      If intRecord = 19 Then
         intRecord = 1
         intPage = intPage + 1
      End If
NextSkip:
      adoacc170.MoveNext
   Loop
   adoacc170.Close
   adoaccrpt204.AddNew
   adoaccrpt204.Fields("r20401").Value = strUserNum
   adoaccrpt204.Fields("r20407").Value = intPage
   adoaccrpt204.Fields("r20408").Value = Mid(strNo, 1, 9)
   adoaccrpt204.Fields("r20413").Value = "小計:"
   adoaccrpt204.Fields("r20416").Value = strCurrency
   adoaccrpt204.Fields("r20417").Value = douAmount
   adoaccrpt204.Fields("r20420").Value = douFAmount
   adoaccrpt204.Fields("r20421").Value = intRecord
   adoaccrpt204.UpdateBatch
   adoaccrpt204.AddNew
   adoaccrpt204.Fields("r20401").Value = strUserNum
   adoaccrpt204.Fields("r20407").Value = intPage + 1
   adoaccrpt204.Fields("r20408").Value = Mid(strNo, 1, 9)
   adoaccrpt204.Fields("r20413").Value = "合計:"
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(r20417) from accrpt204 where r20410 is null", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         adoaccrpt204.Fields("r20417").Value = 0
      Else
         adoaccrpt204.Fields("r20417").Value = adoaccsum.Fields(0).Value
      End If
   Else
      adoaccrpt204.Fields("r20417").Value = 0
   End If
   adoaccsum.Close
   adoaccrpt204.Fields("r20421").Value = intRecord + 1
   adoaccrpt204.UpdateBatch
   adoaccrpt204.Close
   StatusClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt204Delete()
   adoTaie.Execute "delete from accrpt204"
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text1 = ""
   Text2 = ""
   MaskEdBox1.SetFocus
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHead()
   intCounter = 0
   Printer.CurrentX = 4000
   Printer.CurrentY = 500 + intCounter * 300
   If IsNull(adoaccrpt204.Fields("r20402").Value) Then
      Printer.Print ""
   Else
      Printer.Print adoaccrpt204.Fields("r20402").Value
   End If
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = 500 + intCounter * 300
   Printer.Print "代理人: "
   Printer.CurrentX = 1300
   Printer.CurrentY = 500 + intCounter * 300
   If IsNull(adoaccrpt204.Fields("r20410").Value) Then
      Printer.Print ""
   Else
      Printer.Print adoaccrpt204.Fields("r20410").Value
   End If
   Printer.CurrentX = 3000
   Printer.CurrentY = 500 + intCounter * 300
   If IsNull(adoaccrpt204.Fields("r20411").Value) Then
      Printer.Print ""
   Else
      Printer.Print adoaccrpt204.Fields("r20411").Value
   End If
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = 500 + intCounter * 300
   Printer.Print "本所案號"
   Printer.CurrentX = 1500
   Printer.CurrentY = 500 + intCounter * 300
   Printer.Print "收據抬頭"
   Printer.CurrentX = 3500
   Printer.CurrentY = 500 + intCounter * 300
   Printer.Print "國內收款日"
   Printer.CurrentX = 4800
   Printer.CurrentY = 500 + intCounter * 300
   Printer.Print "國內收款"
   Printer.CurrentX = 6100
   Printer.CurrentY = 500 + intCounter * 300
   Printer.Print "幣別"
   Printer.CurrentX = 6900
   Printer.CurrentY = 500 + intCounter * 300
   Printer.Print "國外請款"
   Printer.CurrentX = 9200
   Printer.CurrentY = 500 + intCounter * 300
   Printer.Print "匯率"
   Printer.CurrentX = 10200
   Printer.CurrentY = 500 + intCounter * 300
   Printer.Print "結匯金額"
   Printer.Line (0, 500 + intCounter * 300 + 350)-(11300, 500 + intCounter * 300 + 350)
   intCounter = intCounter + 2
End Sub

'*************************************************
' 列印資料
'
'*************************************************
Private Sub PrintData()
Dim intPage As Integer

   intPage = 0
   Printer.ScaleMode = 1
   'Modify by Morgan 2008/3/25 XP自定紙張需手動設定並將印表機預設為該紙張
   '9x
   If pub_OS = "1" Then
      Printer.Height = 8000
      Printer.Width = 12992
   Else
      Printer.PaperSize = PUB_GetPaperSize(3)
   End If
   'end 2008/3/25
   'add by nick 2004/07/28 設定字型
   Printer.Font.Name = "細明體"
   Printer.FontSize = 12
   adoaccrpt204.CursorLocation = adUseClient
   adoaccrpt204.Open "select * from accrpt204 order by r20407 asc, r20421 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoaccrpt204.EOF = False
      If intPage <> adoaccrpt204.Fields("r20407").Value Then
         If intPage <> 0 Then
            Printer.NewPage
         End If
         PrintHead
         intPage = adoaccrpt204.Fields("r20407").Value
      End If
      If adoaccrpt204.Fields("r20413").Value = "小計:" Then
         Printer.Line (0, 500 + intCounter * 300 - 10)-(11300, 500 + intCounter * 300 - 10)
      End If
      Printer.CurrentX = 0
      Printer.CurrentY = 500 + intCounter * 300
      If IsNull(adoaccrpt204.Fields("r20412").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoaccrpt204.Fields("r20412").Value
      End If
      Printer.CurrentX = 1500
      Printer.CurrentY = 500 + intCounter * 300
      If IsNull(adoaccrpt204.Fields("r20413").Value) Then
         Printer.Print ""
      Else
         Printer.Print StrToStr(adoaccrpt204.Fields("r20413").Value, 8)
      End If
      Printer.CurrentX = 3500
      Printer.CurrentY = 500 + intCounter * 300
      If IsNull(adoaccrpt204.Fields("r20414").Value) Then
         Printer.Print ""
      Else
         Printer.Print CFDate(adoaccrpt204.Fields("r20414").Value)
      End If
      If IsNull(adoaccrpt204.Fields("r20415").Value) = True Or adoaccrpt204.Fields("r20415").Value = 0 Then
         strAmount = ""
      Else
         strAmount = Format(Val(adoaccrpt204.Fields("r20415").Value), DDollar)
      End If
      intLength = Printer.TextWidth(strAmount)
      Printer.CurrentX = 6000 - intLength
      Printer.CurrentY = 500 + intCounter * 300
      Printer.Print strAmount
      Printer.CurrentX = 6100
      Printer.CurrentY = 500 + intCounter * 300
      If IsNull(adoaccrpt204.Fields("r20416").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoaccrpt204.Fields("r20416").Value
      End If
      If IsNull(adoaccrpt204.Fields("r20420").Value) = True Or adoaccrpt204.Fields("r20417").Value = 0 Then
         strAmount = ""
      Else
         strAmount = Format(Val(adoaccrpt204.Fields("r20420").Value), FDollar)
      End If
      intLength = Printer.TextWidth(strAmount)
      Printer.CurrentX = 8100 - intLength
      Printer.CurrentY = 500 + intCounter * 300
      Printer.Print strAmount
      If IsNull(adoaccrpt204.Fields("r20419").Value) = True Or adoaccrpt204.Fields("r20419").Value = 0 Then
         strAmount = ""
      Else
         strAmount = Val(adoaccrpt204.Fields("r20419").Value)
      End If
      intLength = Printer.TextWidth(strAmount)
      Printer.CurrentX = 10000 - intLength
      Printer.CurrentY = 500 + intCounter * 300
      Printer.Print strAmount
      If IsNull(adoaccrpt204.Fields("r20417").Value) = True Or adoaccrpt204.Fields("r20420").Value = 0 Then
         strAmount = ""
      Else
         strAmount = Format(Val(adoaccrpt204.Fields("r20417").Value), DDollar)
      End If
      Printer.CurrentX = 10050
      Printer.CurrentY = 500 + intCounter * 300
      Printer.Print "NT"
      intLength = Printer.TextWidth(strAmount)
      Printer.CurrentX = 11300 - intLength
      Printer.CurrentY = 500 + intCounter * 300
      Printer.Print strAmount
      intCounter = intCounter + 1
      adoaccrpt204.MoveNext
   Loop
   Printer.EndDoc
   adoaccrpt204.Close
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   Select Case Len(Text1)
      Case 6
         Text1 = Text1 & "000"
      Case 8
         Text1 = Text1 & "0"
   End Select
   '2009/6/2 MODIFY BY SONIA 預設尾碼999
   'Text2 = Text1
   'Modify By Sindy 2014/8/11 999=>ZZZ
   'If Text1.Text <> "" Then Text2.Text = Left(Me.Text1.Text, 6) & "999"
   If Text1.Text <> "" Then Text2.Text = Left(Me.Text1.Text, 6) & "ZZZ"
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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

'*************************************************
'  產生匯票函件明細表
'
'*************************************************
Public Function ProcessDetail() As Boolean
Dim intRow As Integer
Dim strName As String
Dim strCase As String
Dim douAmount As Double
Dim strTotalCase As String

   intCounter = 1
   intRow = 1
   douAmount = 0
   adoTaie.Execute "delete from accrpt219 where r21901 = '" & strUserNum & "'"
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from acc190, acc180, acc1b0, fagent where a1901 = a1801 and a1908 = a1b01 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+)" & strSql & " order by a1803 asc, a1902 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoquery.EOF = False
      adoaccsum.CursorLocation = adUseClient
      '2007/3/2 modify by sonia 加入cp87,cp88
      'adoaccsum.Open "select cp01||cp02 from caseprogress where (cp61 = '" & adoquery.Fields("a1902").Value & "' or cp62 = '" & adoquery.Fields("a1902").Value & "' or cp63 = '" & adoquery.Fields("a1902").Value & "')", adoTaie, adOpenStatic, adLockReadOnly
      adoaccsum.Open "select cp01||cp02 from caseprogress where (cp61 = '" & adoquery.Fields("a1902").Value & "' or cp62 = '" & adoquery.Fields("a1902").Value & "' or cp63 = '" & adoquery.Fields("a1902").Value & "' or cp87 = '" & adoquery.Fields("a1902").Value & "' or cp88 = '" & adoquery.Fields("a1902").Value & "')", adoTaie, adOpenStatic, adLockReadOnly
      '2007/3/2 end
      Do While adoaccsum.EOF = False
         If intRow > 4 Then
            adoTaie.Execute "insert into accrpt219 values ('" & strUserNum & "', '" & adoquery.Fields("a1803").Value & "', '" & IIf(IsNull(adoquery.Fields("a1810").Value), adoquery.Fields("fa05").Value, adoquery.Fields("a1810").Value) & "', '" & strCase & "', null, 0, 0, " & intCounter & ")"
            strCase = ""
            intRow = 1
        End If
         If InStr(1, strTotalCase, adoaccsum.Fields(0).Value) > 0 Then
         Else
            strCase = strCase & " " & adoaccsum.Fields(0).Value
            strTotalCase = strTotalCase & " " & adoaccsum.Fields(0).Value
         End If
         intRow = intRow + 1
         adoaccsum.MoveNext
      Loop
      adoaccsum.Close
      douAmount = douAmount + Val(adoquery.Fields("a1904").Value)
      adoquery.MoveNext
      If adoquery.EOF = False Then
         If strName <> adoquery.Fields("a1803").Value Then
            strName = adoquery.Fields("a1803").Value
            adoquery.MovePrevious
            adoTaie.Execute "insert into accrpt219 values ('" & strUserNum & "', '" & adoquery.Fields("a1803").Value & "', '" & IIf(IsNull(adoquery.Fields("a1810").Value), adoquery.Fields("fa05").Value, adoquery.Fields("a1810").Value) & "', '" & strCase & "', '" & adoquery.Fields("a1903").Value & "', " & douAmount & ", " & douAmount & ", " & intCounter & ")"
            adoquery.MoveNext
            strCase = ""
            strTotalCase = ""
            douAmount = 0
            intRow = 1
         End If
      Else
         adoquery.MovePrevious
         adoTaie.Execute "insert into accrpt219 values ('" & strUserNum & "', '" & adoquery.Fields("a1803").Value & "', '" & IIf(IsNull(adoquery.Fields("a1810").Value), adoquery.Fields("fa05").Value, adoquery.Fields("a1810").Value) & "', '" & strCase & "', '" & adoquery.Fields("a1903").Value & "', " & douAmount & ", " & douAmount & ", " & intCounter & ")"
         adoquery.MoveNext
      End If
      intCounter = intCounter + 1
   Loop
   adoquery.Close
End Function

'*************************************************
' 列印資料(匯票函件明細表)
'
'*************************************************
Private Sub PrintData1()
Dim strName As String

   strName = ""
   intCounter = 1
   Printer.FontSize = 12
   adoaccrpt219.CursorLocation = adUseClient
   adoaccrpt219.Open "select * from accrpt219 where r21901 = '" & strUserNum & "' order by r21908 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoaccrpt219.EOF = False
      If intCounter > 40 Then
         intCounter = 1
         Printer.NewPage
         PrintHead1
      End If
      If strName <> adoaccrpt219.Fields("r21902").Value Then
         If strName = "" Then
            PrintHead1
         End If
         strName = adoaccrpt219.Fields("r21902").Value
         Printer.CurrentX = 0
         Printer.CurrentY = 500 + intCounter * 300
         If IsNull(adoaccrpt219.Fields("r21903").Value) Then
            Printer.Print ""
         Else
            Printer.Print Mid(adoaccrpt219.Fields("r21903").Value, 1, 17)
         End If
      End If
      Printer.CurrentX = 2500
      Printer.CurrentY = 500 + intCounter * 300
      If IsNull(adoaccrpt219.Fields("r21904").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoaccrpt219.Fields("r21904").Value
      End If
      Printer.CurrentX = 7200
      Printer.CurrentY = 500 + intCounter * 300
      If IsNull(adoaccrpt219.Fields("r21905").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoaccrpt219.Fields("r21905").Value
      End If
      If IsNull(adoaccrpt219.Fields("r21906").Value) = True Or adoaccrpt219.Fields("r21906").Value = 0 Then
         strAmount = ""
      Else
         strAmount = Format(Val(adoaccrpt219.Fields("r21906").Value), FDollar)
      End If
      intLength = Printer.TextWidth(strAmount)
      Printer.CurrentX = 9100 - intLength
      Printer.CurrentY = 500 + intCounter * 300
      Printer.Print strAmount
      If IsNull(adoaccrpt219.Fields("r21907").Value) = True Or adoaccrpt219.Fields("r21907").Value = 0 Then
         strAmount = ""
      Else
         strAmount = Format(Val(adoaccrpt219.Fields("r21907").Value), FDollar)
      End If
      intLength = Printer.TextWidth(strAmount)
      Printer.CurrentX = 10400 - intLength
      Printer.CurrentY = 500 + intCounter * 300
      Printer.Print strAmount
      intCounter = intCounter + 1
      adoaccrpt219.MoveNext
   Loop
   PrintSum1
   Printer.EndDoc
   adoaccrpt219.Close
End Sub

'*************************************************
'  抬頭列印(匯票函件明細表)
'
'*************************************************
Private Sub PrintHead1()
   intCounter = 0
   Printer.CurrentX = 4000
   Printer.CurrentY = 500 + intCounter * 300
   Printer.Print ReportTitle(219)
   intCounter = intCounter + 2
   Printer.CurrentX = 9000
   Printer.CurrentY = 500 + intCounter * 300
   Printer.Print Format(AFDate(CADate(ACDate(ServerDate))), "mmm. d, yyyy")
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = 500 + intCounter * 300
   Printer.Print "客戶名稱"
   Printer.CurrentX = 2500
   Printer.CurrentY = 500 + intCounter * 300
   Printer.Print "本所案號"
   Printer.CurrentX = 7200
   Printer.CurrentY = 500 + intCounter * 300
   Printer.Print "幣別"
   Printer.CurrentX = 8000
   Printer.CurrentY = 500 + intCounter * 300
   Printer.Print "小計"
   Printer.CurrentX = 9200
   Printer.CurrentY = 500 + intCounter * 300
   Printer.Print "合計"
   Printer.Line (0, 500 + intCounter * 300 + 350)-(10400, 500 + intCounter * 300 + 350)
   intCounter = intCounter + 2
End Sub

'*************************************************
'  列印小計
'
'*************************************************
Public Function PrintSum1() As Boolean
   If intCounter > 38 Then
      intCounter = 1
      Printer.NewPage
   End If
   intCounter = intCounter + 1
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select count(distinct a1b02), count(distinct a1b01) from acc1b0, acc190 where a1b01 = a1908" & strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) = False Then
         Printer.CurrentX = 0
         Printer.CurrentY = 500 + intCounter * 300
         Printer.Print "代理人合計: " & adoquery.Fields(0).Value & " 家"
      End If
      intCounter = intCounter + 1
      If IsNull(adoquery.Fields(1).Value) = False Then
         Printer.CurrentX = 0
         Printer.CurrentY = 500 + intCounter * 300
         Printer.Print "匯票合計: " & adoquery.Fields(1).Value & " 張"
      End If
      intCounter = intCounter + 1
   End If
   adoquery.Close
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select count(distinct axf03) from acc1b0, acc190, acc151 where a1b01 = a1908 and a1902 = axf01" & strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) = False Then
         Printer.CurrentX = 0
         Printer.CurrentY = 500 + intCounter * 300
         Printer.Print "件數合計: " & adoquery.Fields(0).Value & " 件"
         intCounter = intCounter + 1
      End If
   End If
   adoquery.Close
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select sum(a1p08) from acc1b0, acc1p0 where a1b01||a1b02 = a1p04" & strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) = False Then
         Printer.CurrentX = 0
         Printer.CurrentY = 500 + intCounter * 300
         Printer.Print "台幣結匯金額合計: " & Format(adoquery.Fields(0).Value, FDollar)
      End If
   End If
   adoquery.Close
End Function

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 = MsgText(601) Then
      Exit Sub
   End If
   Select Case Len(Text2)
      Case 6
         Text2 = Text2 & "000"
      Case 8
         Text2 = Text2 & "0"
   End Select
End Sub
