VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc34c0 
   AutoRedraw      =   -1  'True
   Caption         =   "銀行帳號別資金流動表"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   5160
   Begin VB.CheckBox Check1 
      Caption         =   "是否產生Excel檔案"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   1995
   End
   Begin VB.CommandButton Command1 
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
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   1440
      Width           =   4692
   End
   Begin VB.ComboBox Combo6 
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
      Left            =   3600
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo5 
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
      Left            =   960
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo4 
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
      Left            =   3600
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   1212
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
      Left            =   960
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   1812
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
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1572
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
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1800
      TabIndex        =   2
      Top             =   600
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1335
      Left            =   240
      Top             =   2040
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc34c0.frx":0000
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "2."
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
      Left            =   720
      TabIndex        =   14
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc34c0.frx":0442
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "1."
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
      Left            =   720
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "排序方式"
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
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "最後傳票日期"
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
      TabIndex        =   11
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label5 
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
      TabIndex        =   10
      Top             =   240
      Width           =   252
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "銀行代號"
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
      TabIndex        =   9
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc34c0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0e0 As New ADODB.Recordset
Public adoacc0h0 As New ADODB.Recordset
Public adoacc0g0 As New ADODB.Recordset
Public adoaccrpt313 As New ADODB.Recordset
Dim ado313Sum As New ADODB.Recordset 'Add by Amy 2013/08/12 合計用
Dim strSort1, strSort2 As String
'Modify by Amy 2013/08/12 改寫法
'Dim dllaccrpt313 As Object
Dim intCounter As Integer
Dim intRecord As Integer
Dim intPage As Integer
Dim SumAmt(2) As Double '合計用
'end 2013/08/12
Dim bolA4 As Boolean 'Add by amy 2013/08/22 改為A4列印-辜

Private Sub Combo3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo4.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo4_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo5.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo5_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo6.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo6_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Text1.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Command1_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   
   bolA4 = True 'Add by Amy 2013/08/22
      
   Screen.MousePointer = vbHourglass
   Accrpt313Delete
   ProduceData
   
   If adoaccrpt313.State = adStateOpen Then
      adoaccrpt313.Close
   End If
   adoaccrpt313.CursorLocation = adUseClient
   adoaccrpt313.Open "select * from accrpt313", adoTaie, adOpenStatic, adLockReadOnly
   
   If adoaccrpt313.RecordCount <> 0 Then
      'Modify by Amy 2013/08/12 改寫(不使用accReport) 且增加產生excel
      If ado313Sum.State = adStateOpen Then
            ado313Sum.Close
      End If
      ado313Sum.CursorLocation = adUseClient
      strExc(0) = "Select R31310,S31309,to_char(round(S31309/Total*100,2)) P From(" & _
                       "Select R31310,sum(R31309) S31309 From accrpt313 Where R31301='" & strUserNum & "' group by R31310),(select sum(R31309) Total from accrpt313 Where R31301='" & strUserNum & "') "
      ado313Sum.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
      
      If Check1.Value = 1 Then
        Call ExcelSaveNew
      ElseIf bolA4 Then
        Call PrintDataA4
        'dllaccrpt313.Acc34c0 ReportTitle(313), Text1, Text2, MaskEdBox1.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
        'PrintData '2013/08/22 改印A4 不印大報表
      End If
      'end
   End If
   If ado313Sum.State = adStateOpen Then 'Add by Amy 2013/11/12 +if修正查無資料error
        ado313Sum.Close
   End If
   adoaccrpt313.Close
   Screen.MousePointer = vbDefault
   FormClear
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)'Modify by Amy 2013/08/22
   Frmacc0000.StatusBar1.Panels(1).Text = "以 A4 列印"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
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
   'Modify by Amy 2013/08/12 +產生excel
   'Me.Height = 2100
   Me.Height = 2340
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
   Combo4.AddItem MsgText(1)
   Combo4.AddItem MsgText(2)
   Combo6.AddItem MsgText(1)
   Combo6.AddItem MsgText(2)
   Combo4 = MsgText(1)
   Combo6 = MsgText(1)
   ComboAdd
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   'Set dllaccrpt313 = CreateObject("AccReport.ReportSelect") 'Modify by Amy 2013/08/12 改寫法
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   'Set dllaccrpt313 = Nothing 'Modify by Amy 2013/08/12 改寫法
   Set Frmacc34c0 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

'*************************************************
'  Combo 項目新增
'
'*************************************************
Private Sub ComboAdd()
   strSort1 = "銀行代號"
   strSort2 = "銀行帳號"
   Combo3.AddItem strSort1
   Combo3.AddItem strSort2
   Combo5.AddItem strSort1
   Combo5.AddItem strSort2
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim strOrder1, strOrder2 As String
Dim strSql As String
Dim lngDueDate As Long
Dim intYear, intMonth As Integer
Dim adoacc0b0 As New ADODB.Recordset
Dim adoquery As New ADODB.Recordset
Dim strName As String
Dim strProDate As String
   
On Error GoTo Checking
   adoacc0b0.CursorLocation = adUseClient
   adoacc0b0.Open "select * from acc0b0", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0b0.RecordCount = 0 Then
      If Mid(ServerDate, 5, 2) = 1 Then
         intMonth = 12
         intYear = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) - 1
      Else
         intMonth = Val(Mid(ServerDate, 5, 2)) - 1
         intYear = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3))
      End If
      strProDate = intYear & IIf(intMonth > 9, intMonth, "0" & intMonth) & "00"
   Else
      If IsNull(adoacc0b0.Fields("a0b02").Value) Then
         If Mid(ServerDate, 5, 2) = 1 Then
            intMonth = 12
            intYear = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) - 1
         Else
            intMonth = Val(Mid(ServerDate, 5, 2)) - 1
            intYear = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3))
         End If
         strProDate = intYear & IIf(intMonth > 9, intMonth, "0" & intMonth) & "00"
      Else
         intMonth = Val(Mid(CFDate(adoacc0b0.Fields("a0b02").Value), 5, 2))
         intYear = Val(Mid(CFDate(adoacc0b0.Fields("a0b02").Value), 1, 3))
      End If
      strProDate = adoacc0b0.Fields("a0b02").Value
   End If
   adoacc0b0.Close
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Select Case Combo3
      Case strSort1
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0h01 asc"
         Else
            strOrder1 = " order by a0h01 desc"
         End If
      Case strSort2
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0h02 asc"
         Else
            strOrder1 = " order by a0h02 desc"
         End If
      Case Else
         strOrder1 = MsgText(601)
   End Select
   Select Case Combo5
      Case strSort1
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0h01 asc"
         Else
            strOrder2 = ", a0h01 desc"
         End If
      Case strSort2
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0h02 asc"
         Else
            strOrder2 = ", a0h02 desc"
         End If
      Case Else
         strOrder2 = MsgText(601)
   End Select
   If Text1 <> MsgText(601) Then
      strSql = " and a0h01 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0h01 <= '" & Text2 & "'"
   End If
   If strSql <> MsgText(601) Then
      strSql = " WHERE " & Mid(strSql, 5, Len(strSql) - 4)
   End If
   adoaccrpt313.CursorLocation = adUseClient
   adoaccrpt313.Open "select * from accrpt313 order by r31301,r31302", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0h0.CursorLocation = adUseClient
   adoacc0h0.Open "select * from acc0h0" & strSql & strOrder1 & strOrder2, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0h0.RecordCount = 0 Then
      adoacc0h0.Close
      adoaccrpt313.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc0h0.EOF = False
      adoacc0g0.CursorLocation = adUseClient
      adoacc0g0.Open "SELECT A0G09, A0G02 FROM ACC0G0 WHERE A0G01 = '" & adoacc0h0.Fields("A0H01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0g0.RecordCount <> 0 Then
         If adoacc0g0.Fields(0).Value = MsgText(602) Then
            lngDueDate = Val(ACDate(Format(CDate(Mid(CADate(FCDate(MaskEdBox1.Text)), 1, 4) & "/" & Mid(CADate(FCDate(MaskEdBox1.Text)), 5, 2) & "/" & Mid(CADate(FCDate(MaskEdBox1.Text)), 7, 2)) - 1, "YYYYMMDD")))
         Else
            lngDueDate = Val(ACDate(Format(CDate(Mid(CADate(FCDate(MaskEdBox1.Text)), 1, 4) & "/" & Mid(CADate(FCDate(MaskEdBox1.Text)), 5, 2) & "/" & Mid(CADate(FCDate(MaskEdBox1.Text)), 7, 2)) - 3, "YYYYMMDD")))
         End If
         If IsNull(adoacc0g0.Fields(1).Value) Then
            strName = ""
         Else
            strName = adoacc0g0.Fields(1).Value
         End If
      End If
      adoacc0g0.Close
      adoaccrpt313.AddNew
      adoaccrpt313.Fields("r31301").Value = strUserNum
      adoaccrpt313.Fields("r31302").Value = adoacc0h0.Fields("a0h01").Value
      adoaccrpt313.Fields("r31303").Value = adoacc0h0.Fields("a0h02").Value
      adoaccrpt313.Fields("r31304").Value = strName
      'Add by Amy 2013/08/12 +出名人及存款類別欄位
      adoaccrpt313.Fields("r31310").Value = adoacc0h0.Fields("a0h15").Value
      adoaccrpt313.Fields("r31311").Value = adoacc0h0.Fields("a0h16").Value
      'end 2013/08/12
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select a0408 from acc040 where a0403 = '1' and a0401 = " & intYear & " and a0402 = " & intMonth & " and a0405 = '" & adoacc0h0.Fields("a0h08").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields(0).Value) Then
            adoaccrpt313.Fields("r31306").Value = 0
         Else
            adoaccrpt313.Fields("r31306").Value = Val(adoquery.Fields(0).Value)
         End If
      Else
         adoaccrpt313.Fields("r31306").Value = 0
      End If
      adoquery.Close
      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
         adoaccrpt313.Fields("r31305").Value = Val(FCDate(MaskEdBox1.Text))
      Else
         adoaccrpt313.Fields("r31305").Value = Null
      End If
      '未收票據
      adoacc0e0.CursorLocation = adUseClient
      'Ken 91/03/26 -- Start
      'adoacc0e0.Open "select sum(a0e11) from acc0e0, acc0h0 where a0e19 = a0h01 and a0e20 = a0h02 and a0h08 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0e04 = '" & MsgText(18) & "' and (a0e10 <= " & Val(FCDate(MaskEdBox1.Text)) & " and a0e17 = 0 and a0e15 = 0 and a0e34 = 0 and a0e21 = 0)", adoTaie, adOpenStatic, adLockReadOnly
      adoacc0e0.Open "select sum(a0e11) from acc0e0, acc0h0 where a0e19 = a0h01 and a0e20 = a0h02 and a0h08 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0e04 = '" & MsgText(18) & "' and (a0e10 > " & Val(FCDate(MaskEdBox1.Text)) & " and a0e17 = 0 and a0e15 = 0 and a0e34 = 0 and a0e21 = 0)", adoTaie, adOpenStatic, adLockReadOnly
      'Ken 91/03/26 -- End
      If adoacc0e0.RecordCount <> 0 Then
         If IsNull(adoacc0e0.Fields(0).Value) Then
            adoaccrpt313.Fields("r31307").Value = 0
         Else
            adoaccrpt313.Fields("r31307").Value = Val(adoacc0e0.Fields(0).Value)
         End If
      Else
         adoaccrpt313.Fields("r31307").Value = 0
      End If
      adoacc0e0.Close
      '已收已入帳
      adoacc0e0.CursorLocation = adUseClient
      'Ken 91/03/26 -- Start
      'adoacc0e0.Open "select sum(ax206) from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and ax205 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0205 > " & Val(strProDate) & "", adoTaie, adOpenStatic, adLockReadOnly
      adoacc0e0.Open "select sum(ax206) from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and ax205 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0205 > " & Val(strProDate) & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & "", adoTaie, adOpenStatic, adLockReadOnly
      'Ken 91/03/26 -- End
      If adoacc0e0.RecordCount <> 0 Then
         If IsNull(adoacc0e0.Fields(0).Value) = False Then
            adoaccrpt313.Fields("r31307").Value = Val(adoaccrpt313.Fields("r31307").Value) + Val(adoacc0e0.Fields(0).Value)
         End If
      End If
      adoacc0e0.Close
      '未付票據
      adoacc0e0.CursorLocation = adUseClient
      'Ken 91/03/26 -- Start
      'adoacc0e0.Open "select sum(a0e11) from acc0e0, acc0h0 where a0e01 = a0h01 and a0e07 = a0h02 and a0h08 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0e04 = '" & MsgText(19) & "' and (a0e10 <= " & Val(FCDate(MaskEdBox1.Text)) & "  and a0e25 = 0 and (a0e37 = 0 or a0e37 is null))", adoTaie, adOpenStatic, adLockReadOnly
      adoacc0e0.Open "select sum(a0e11) from acc0e0, acc0h0 where a0e01 = a0h01 and a0e07 = a0h02 and a0h08 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0e04 = '" & MsgText(19) & "' and (a0e10 > " & Val(FCDate(MaskEdBox1.Text)) & "  and a0e25 = 0 and (a0e37 = 0 or a0e37 is null))", adoTaie, adOpenStatic, adLockReadOnly
      'Ken 91/03/26 -- End
      If adoacc0e0.RecordCount <> 0 Then
         If IsNull(adoacc0e0.Fields(0).Value) Then
            adoaccrpt313.Fields("r31308").Value = 0
         Else
            adoaccrpt313.Fields("r31308").Value = Val(adoacc0e0.Fields(0).Value)
         End If
      Else
         adoaccrpt313.Fields("r31308").Value = 0
      End If
      adoacc0e0.Close
      '已付已入帳
      adoacc0e0.CursorLocation = adUseClient
      'Ken 91/03/26 -- Start
      'adoacc0e0.Open "select sum(ax207) from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and ax205 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0205 > " & Val(strProDate) & "", adoTaie, adOpenStatic, adLockReadOnly
      adoacc0e0.Open "select sum(ax207) from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and ax205 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0205 > " & Val(strProDate) & " and a0205 <= " & Val(FCDate(MaskEdBox1.Text)) & "", adoTaie, adOpenStatic, adLockReadOnly
      'Ken 91/03/26 -- End
      If adoacc0e0.RecordCount <> 0 Then
         If IsNull(adoacc0e0.Fields(0).Value) = False Then
            adoaccrpt313.Fields("r31308").Value = Val(adoaccrpt313.Fields("r31308").Value) + Val(adoacc0e0.Fields(0).Value)
         End If
      End If
      adoacc0e0.Close
      If IsNull(adoaccrpt313.Fields("r31306").Value) Then
         adoaccrpt313.Fields("r31309").Value = Val(adoaccrpt313.Fields("r31307").Value) - Val(adoaccrpt313.Fields("r31308").Value)
      Else
         adoaccrpt313.Fields("r31309").Value = Val(adoaccrpt313.Fields("r31306").Value) + Val(adoaccrpt313.Fields("r31307").Value) - Val(adoaccrpt313.Fields("r31308").Value)
      End If
      adoaccrpt313.UpdateBatch
      adoacc0h0.MoveNext
   Loop
   adoacc0h0.Close
   adoaccrpt313.Close
   'adoTaie.Execute "delete from accrpt313 where r31302 is null" 'Modify by Amy 2013/08/12 排除數字都是0
   adoTaie.Execute "delete from accrpt313 where r31302 is null Or (r31306 = 0 And r31307 = 0 And r31308 = 0 And r31309 = 0)"
   'Add by Amy 2013/08/12 +新增合計(列印用)
   If Check1.Value = 0 Then
        SumAmt(0) = 0: SumAmt(1) = 0: SumAmt(2) = 0
        ado313Sum.CursorLocation = adUseClient
        strExc(0) = "Select  sum(R31307) S31307,sum(R31308) S31308,sum(R31309) S31309 From accrpt313 Where R31301='" & strUserNum & "' "
        ado313Sum.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
        If ado313Sum.RecordCount <> 0 Then
            SumAmt(0) = ado313Sum.Fields("S31307")
            SumAmt(1) = ado313Sum.Fields("S31308")
            SumAmt(2) = ado313Sum.Fields("S31309")
        End If
         ado313Sum.Close
   End If
   'end 2013/08/12

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
Private Sub Accrpt313Delete()
   adoTaie.Execute "delete from accrpt313"
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
   Text2 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   Check1.Value = 0 'Add by Amy 2013/08/12
   Combo3 = ""
   Combo5 = ""
   Text1.SetFocus
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
   FormCheck = False
End Function

'Add by Amy 2013/08/12 改不使用accReport(因需顯示2個表格)-大報表
Private Sub PrintData()
   Dim strAmount As String
   Dim intLength As Integer
   Dim isFirstG As Boolean '是否第一次印分類抬頭
   intCounter = 3: intRecord = 1: intPage = 1
   
   PrintHead
   PrintHeadD
   Do While adoaccrpt313.EOF = False
       If intRecord > 34 Then
            intPage = intPage + 1
            intRecord = 1
            Printer.NewPage
             intCounter = 3
            PrintHead
            PrintHeadD
       End If
       Printer.FontBold = False
        '存款類別
        Printer.CurrentX = 0
        Printer.CurrentY = 300 + intCounter * 300
        If IsNull(adoaccrpt313.Fields("R31311").Value) Then
            Printer.Print ""
        Else
            Printer.Print adoaccrpt313.Fields("R31311").Value
        End If
        '銀行名稱
        Printer.CurrentX = 1700
        Printer.CurrentY = 300 + intCounter * 300
        If IsNull(adoaccrpt313.Fields("R31304").Value) Then
            Printer.Print ""
        Else
            Printer.Print adoaccrpt313.Fields("R31304").Value
        End If
        '期初餘額
        If IsNull(adoaccrpt313.Fields("R31306").Value) = False Then
            strAmount = Format(Val(adoaccrpt313.Fields("R31306").Value), DDollar2)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 9400 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
        End If
        '收入
        If IsNull(adoaccrpt313.Fields("R31307").Value) = False Then
            strAmount = Format(Val(adoaccrpt313.Fields("R31307").Value), DDollar2)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 10700 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
        End If
        '支出
        If IsNull(adoaccrpt313.Fields("R31308").Value) = False Then
            strAmount = Format(Val(adoaccrpt313.Fields("R31308").Value), DDollar2)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 12000 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
        End If
        '帳號餘額
        If IsNull(adoaccrpt313.Fields("R31309").Value) = False Then
            strAmount = Format(Val(adoaccrpt313.Fields("R31309").Value), DDollar2)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 13300 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
        End If
        '出名人
        Printer.CurrentX = 13500
        Printer.CurrentY = 300 + intCounter * 300
        If IsNull(adoaccrpt313.Fields("R31310").Value) Then
            Printer.Print ""
        Else
            Printer.Print adoaccrpt313.Fields("R31310").Value
        End If

        intCounter = intCounter + 1
        intRecord = intRecord + 1
        adoaccrpt313.MoveNext
   Loop
   
   If intRecord > 34 Then
        intPage = intPage + 1
        intRecord = 1
        Printer.NewPage
        intCounter = 3
        PrintHead
        PrintHeadD
   End If
   '表尾合計
   Printer.FontBold = True
   intRecord = intRecord + 1
   Printer.CurrentX = 8600
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "合計:"
      
   Printer.FontBold = False
   Printer.Line (9500, 300 + intCounter * 300 - 50)-(12000, 300 + intCounter * 300 - 50)
   strAmount = Format(Val(SumAmt(0)), DDollar2)
   intLength = Printer.TextWidth(strAmount)
   Printer.CurrentX = 10700 - intLength
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print strAmount
   strAmount = Format(Val(SumAmt(1)), DDollar2)
   intLength = Printer.TextWidth(strAmount)
   Printer.CurrentX = 12000 - intLength
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print strAmount
   Printer.Line (9500, 300 + intCounter * 300 + 350)-(12000, 300 + intCounter * 300 + 350)
   Printer.Line (9500, 300 + intCounter * 300 + 400)-(12000, 300 + intCounter * 300 + 400)
   intCounter = intCounter + 2
   
      
   '分類明細表
   ado313Sum.MoveFirst
   If ado313Sum.RecordCount <> 0 Then
        isFirstG = True
        If intRecord > 34 Then
            intPage = intPage + 1
            intRecord = 1
            Printer.NewPage
             intCounter = 3
            PrintHead
            PrintHeadG
       ElseIf isFirstG Then
           isFirstG = False
           PrintHeadG
       End If
        Printer.FontBold = False
        Do While ado313Sum.EOF = False
           '出名人
            Printer.CurrentX = 0
            Printer.CurrentY = 300 + intCounter * 300
            If IsNull(ado313Sum.Fields("R31310").Value) Then
                Printer.Print ""
            Else
                Printer.Print ado313Sum.Fields("R31310").Value
            End If
            '金額
            If IsNull(ado313Sum.Fields("S31309").Value) = False Then
                strAmount = Format(Val(ado313Sum.Fields("S31309").Value), DDollar2)
                intLength = Printer.TextWidth(strAmount)
                Printer.CurrentX = 7400 - intLength
                Printer.CurrentY = 300 + intCounter * 300
                Printer.Print strAmount
            End If
            '所佔比率
            If IsNull(ado313Sum.Fields("P").Value) = False Then
                strAmount = Format(Val(ado313Sum.Fields("P").Value), FDollar)
                intLength = Printer.TextWidth(strAmount & "%")
                Printer.CurrentX = 9100 - intLength
                Printer.CurrentY = 300 + intCounter * 300
                Printer.Print strAmount & "%"
            End If
            
            intCounter = intCounter + 1
            intRecord = intRecord + 1
            ado313Sum.MoveNext
        Loop
        
        If intRecord > 34 Then
            intPage = intPage + 1
            intRecord = 1
            Printer.NewPage
            intCounter = 3
            PrintHead
            PrintHeadG
        End If
        '分類明細合計
        Printer.FontBold = True
        intRecord = intRecord + 1
        Printer.CurrentX = 2200
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print "合計:"
        
        Printer.Line (6000, 300 + intCounter * 300 - 50)-(7500, 300 + intCounter * 300 - 50)
        strAmount = Format(Val(SumAmt(2)), DDollar2)
        intLength = Printer.TextWidth(strAmount)
        Printer.CurrentX = 7400 - intLength
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print strAmount
        Printer.Line (6000, 300 + intCounter * 300 + 350)-(7500, 300 + intCounter * 300 + 350)
        Printer.Line (6000, 300 + intCounter * 300 + 400)-(7500, 300 + intCounter * 300 + 400)
        
        intCounter = intCounter + 2
        Printer.FontBold = True
        Printer.CurrentX = 8200
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print "*** 結束 ***"
   End If
   
   Printer.EndDoc
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHead()

   Printer.FontSize = 17
   Printer.FontBold = True
   Printer.CurrentX = 5000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print ReportTitle(313)
   
   Printer.FontSize = 12
   intCounter = intCounter + 2
   Printer.CurrentX = 5000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "銀行代號: " & Text1 & "~" & Text2
   intCounter = intCounter + 1
   Printer.CurrentX = 5000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "到期日期: " & MaskEdBox1
   
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印人員: " & StaffQuery(strUserNum)
   Printer.CurrentX = 12000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印日期: " & CFDate(ACDate(ServerDate))
   intCounter = intCounter + 1
   Printer.CurrentX = 12000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "頁次: " & intPage
  
End Sub

'*************************************************
'  明細表抬頭列印
'
'*************************************************
Private Sub PrintHeadD()

   intCounter = intCounter + 2
   Printer.CurrentX = 250
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "存款類別"
   Printer.CurrentX = 3000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "銀行名稱"
  
   Printer.CurrentX = 8100
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "期初餘額"
   Printer.CurrentX = 9700
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "收　入"
   Printer.CurrentX = 11100
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "支　出"
   Printer.CurrentX = 12200
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "帳號餘額"
   Printer.CurrentX = 15000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "出名人"
   Printer.CurrentX = 17500
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "備　註"
   

   
   Printer.Line (0, 300 + intCounter * 300 + 350)-(19700, 300 + intCounter * 300 + 350)
   intCounter = intCounter + 2
End Sub

'*************************************************
'  分類表抬頭列印
'
'*************************************************
Private Sub PrintHeadG()

   Printer.FontSize = 12
  
   intCounter = intCounter + 2
   Printer.CurrentX = 2200
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "以戶名分類明細"
   Printer.CurrentX = 6400
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "金　額"
  
   Printer.CurrentX = 7900
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "所佔比率"
   
   Printer.Line (0, 300 + intCounter * 300 + 350)-(9200, 300 + intCounter * 300 + 350)
   intCounter = intCounter + 2
End Sub

'*************************************************
'  轉成Excel檔案
'
'*************************************************
Private Sub ExcelSaveNew()
Dim xlsAgentPoint As New Excel.Application
Dim wksrpt As New Worksheet
Dim xlsFileName As String
Dim iRow As Integer, RsCount As Integer
Dim MaxColStr As String '最右邊欄位代碼
Dim GCount As Integer
Dim SumGRow As Integer    '戶名分類合計row

On Error GoTo ErrHnd

 If Dir(strExcelPath & "銀行帳號別資金流動表" & ServerDate & MsgText(43)) = MsgText(601) Then
    If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
    End If
 Else
      Kill strExcelPath & "銀行帳號別資金流動表" & ServerDate & MsgText(43)
 End If

 MaxColStr = Chr(Asc("a") + 7)
 RsCount = adoaccrpt313.RecordCount
 GCount = ado313Sum.RecordCount
 
 xlsAgentPoint.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
 xlsAgentPoint.Workbooks.add
 Set wksrpt = xlsAgentPoint.Worksheets(1)
 
 wksrpt.Columns("a:a").ColumnWidth = 13
 wksrpt.Columns("b:b").ColumnWidth = 30
 wksrpt.Columns("c:c").ColumnWidth = 13
 wksrpt.Columns("d:d").ColumnWidth = 13
 wksrpt.Columns("e:e").ColumnWidth = 13
 wksrpt.Columns("f:f").ColumnWidth = 13
 wksrpt.Columns("g:g").ColumnWidth = 22
 wksrpt.Columns("h:h").ColumnWidth = 13
 
 iRow = 1
 wksrpt.Range("a" & iRow).Value = ReportTitle(313)
 wksrpt.Range("a" & iRow).HorizontalAlignment = xlCenter
 With wksrpt.Range("a1:" & MaxColStr & "1")
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Size = 18
    .Font.Bold = True
    .MergeCells = True
 End With
 
 
  iRow = iRow + 2
 wksrpt.Range("c" & iRow).Value = "銀行代號："
 wksrpt.Range("c" & iRow).HorizontalAlignment = xlRight
 wksrpt.Range("d" & iRow).Value = Text1 & " ~ " & Text2
 wksrpt.Range("d" & iRow).HorizontalAlignment = xlLeft
 iRow = iRow + 1
 wksrpt.Range("c" & iRow).Value = "到期日期："
 wksrpt.Range("c" & iRow).HorizontalAlignment = xlRight
 wksrpt.Range("d" & iRow).Value = MaskEdBox1
 wksrpt.Range("d" & iRow).HorizontalAlignment = xlLeft
  iRow = iRow + 1
 wksrpt.Range("a" & iRow).Value = "列印人員："
 wksrpt.Range("a" & iRow).HorizontalAlignment = xlRight
 wksrpt.Range("b" & iRow).Value = StaffQuery(strUserNum)
 wksrpt.Range("a" & iRow).HorizontalAlignment = xlLeft
 wksrpt.Range("g" & iRow).Value = "列印日："
 wksrpt.Range("g" & iRow).HorizontalAlignment = xlRight
 wksrpt.Range("h" & iRow).Value = CFDate(ACDate(ServerDate))
 wksrpt.Range("h" & iRow).HorizontalAlignment = xlLeft
  iRow = iRow + 2
  
  '明細表抬頭
  wksrpt.Range("a" & iRow).Value = "存款類別"
  wksrpt.Range("b" & iRow).Value = "銀行名稱"
  wksrpt.Range("c" & iRow).Value = "期初餘額"
  wksrpt.Range("d" & iRow).Value = "收　入"
  wksrpt.Range("e" & iRow).Value = "支　出"
  wksrpt.Range("f" & iRow).Value = "帳號餘額"
  wksrpt.Range("g" & iRow).Value = "出　名　人"
  wksrpt.Range("h" & iRow).Value = "備　註"
  
  '明細表內容
  Do While adoaccrpt313.EOF = False
        iRow = iRow + 1

        wksrpt.Range("a" & iRow).Value = "" & adoaccrpt313.Fields("R31311")
        wksrpt.Range("b" & iRow).Value = "" & adoaccrpt313.Fields("R31304")
        wksrpt.Range("c" & iRow).Value = Val(adoaccrpt313.Fields("R31306"))
        wksrpt.Range("d" & iRow).Value = Val(adoaccrpt313.Fields("R31307"))
        wksrpt.Range("e" & iRow).Value = Val(adoaccrpt313.Fields("R31308"))
        wksrpt.Range("f" & iRow).Value = Val(adoaccrpt313.Fields("R31309"))
        wksrpt.Range("g" & iRow).Value = "" & adoaccrpt313.Fields("R31310")
        adoaccrpt313.MoveNext
  Loop
 
  '明細表合計
   iRow = iRow + 1
   wksrpt.Range("c" & iRow).Value = "合計"
   wksrpt.Range("d" & iRow).Formula = "=sum(d" & iRow - RsCount & ":d" & iRow - 1 & ")"
   wksrpt.Range("e" & iRow).Formula = "=sum(e" & iRow - RsCount & ":e" & iRow - 1 & ")"
   
   '明細表格式設定
   wksrpt.Range("a7:" & MaxColStr & "7").HorizontalAlignment = xlCenter
   wksrpt.Range("a7:" & MaxColStr & "7").Font.Bold = True
   wksrpt.Range("c" & iRow).HorizontalAlignment = xlCenter
   wksrpt.Range("c" & iRow).Font.Bold = True
   wksrpt.Range("c8:c" & iRow).NumberFormatLocal = "#,##0"
   wksrpt.Range("d8:d" & iRow).NumberFormatLocal = "#,##0"
   wksrpt.Range("e8:e" & iRow).NumberFormatLocal = "#,##0"
   wksrpt.Range("f8:f" & iRow).NumberFormatLocal = "#,##0"
 With wksrpt.Range("a7:" & MaxColStr & iRow)
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    .Borders(xlInsideVertical).LineStyle = xlContinuous
    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
 End With

  '分類表抬頭
   iRow = iRow + 2
  wksrpt.Range("b" & iRow).Value = "以戶名分類明細"
  wksrpt.Range("c" & iRow).Value = "金　額"
  wksrpt.Range("d" & iRow).Value = "所佔比率"

  '分類表內容
   SumGRow = iRow + 1 + GCount
   Do While ado313Sum.EOF = False
        iRow = iRow + 1

        wksrpt.Range("b" & iRow).Value = "" & ado313Sum.Fields("R31310")
        wksrpt.Range("c" & iRow).Value = Val(ado313Sum.Fields("S31309"))
        wksrpt.Range("d" & iRow).Formula = "=c" & iRow & "/$c$" & SumGRow
        ado313Sum.MoveNext
   Loop

   '分類表合計
    iRow = iRow + 1
   wksrpt.Range("b" & iRow).Value = "合計"
   wksrpt.Range("c" & iRow).Formula = "=sum(c" & iRow - GCount & ":c" & iRow - 1 & ")"
   
  '分類表格式設定
  wksrpt.Range("b" & iRow).HorizontalAlignment = xlCenter
  wksrpt.Range("b" & iRow).Font.Bold = True
  wksrpt.Range("b" & iRow - GCount - 1 & ":d" & iRow - GCount - 1).HorizontalAlignment = xlCenter
  wksrpt.Range("b" & iRow - GCount - 1 & ":d" & iRow - GCount - 1).Font.Bold = True
  wksrpt.Range("c" & iRow - GCount & ":c" & SumGRow).NumberFormatLocal = "#,##0"
  wksrpt.Range("d" & iRow - GCount & ":d" & SumGRow).NumberFormatLocal = "0.00%"
 With wksrpt.Range("b" & iRow - GCount - 1 & ":d" & SumGRow)
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    .Borders(xlInsideVertical).LineStyle = xlContinuous
    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
 End With
   
 'Modify by Amy 2015/05/21 原使用函數以為是抓A4紙張
 wksrpt.PageSetup.PaperSize = 9 '設定紙張 A4
 wksrpt.PageSetup.Orientation = xlLandscape '橫印
 wksrpt.PageSetup.PrintTitleRows = "$1:$6" '表頭保留7列
 wksrpt.PageSetup.PrintArea = "$A$1:$" & MaxColStr & "$" & iRow '設定列印範圍
      
 wksrpt.PageSetup.LeftMargin = xlsAgentPoint.InchesToPoints(0.5)
 wksrpt.PageSetup.RightMargin = xlsAgentPoint.InchesToPoints(0.5)
 wksrpt.PageSetup.TopMargin = xlsAgentPoint.InchesToPoints(0.3)
 wksrpt.PageSetup.BottomMargin = xlsAgentPoint.InchesToPoints(0.3)
 wksrpt.PageSetup.HeaderMargin = xlsAgentPoint.InchesToPoints(0.5)
 wksrpt.PageSetup.FooterMargin = xlsAgentPoint.InchesToPoints(0.5)
      
wksrpt.PageSetup.Zoom = 100 '縮放比例
   'Modify by Amy 2016/06/23 +判斷版本
   If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & "銀行帳號別資金流動表" & ServerDate & MsgText(43), FileFormat:=-4143
   Else
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & "銀行帳號別資金流動表" & ServerDate & MsgText(43), FileFormat:=56
   End If
   'end 2016/06/23
   xlsAgentPoint.Workbooks.Close
   xlsAgentPoint.Quit
   Set xlsAgentPoint = Nothing
   Set wksrpt = Nothing
   MsgBox "Excel檔案產生完成！（檔案位置：" & strExcelPath & "銀行帳號別資金流動表" & ServerDate & MsgText(43) & "）"
   Exit Sub
   
ErrHnd:
   If Not xlsAgentPoint Is Nothing Then
      xlsAgentPoint.Quit
      Set xlsAgentPoint = Nothing
      Set wksrpt = Nothing
   End If
   MsgBox Err.Description
End Sub
'end 2013/08/12

'Add by Amy 2013/08/22 改以A4印
Private Sub PrintDataA4()
   Dim strAmount As String
   Dim intLength As Integer
   Dim isFirstG As Boolean '是否第一次印分類抬頭
   intCounter = 3: intRecord = 1: intPage = 1
   
   Printer.PaperSize = PUB_GetPaperSize(9) '設定紙張 A4
   Printer.Orientation = xlLandscape '橫印
   PrintHeadA4
   PrintHeadDA4
   Do While adoaccrpt313.EOF = False
       If intRecord > 22 Then
            intPage = intPage + 1
            intRecord = 1
            Printer.NewPage
             intCounter = 3
            PrintHeadA4
            PrintHeadDA4
       End If
       Printer.FontBold = False
        '存款類別
        Printer.CurrentX = 0
        Printer.CurrentY = 300 + intCounter * 300
        If IsNull(adoaccrpt313.Fields("R31311").Value) Then
            Printer.Print ""
        Else
            Printer.Print adoaccrpt313.Fields("R31311").Value
        End If
        '銀行名稱
        Printer.CurrentX = 1700
        Printer.CurrentY = 300 + intCounter * 300
        If IsNull(adoaccrpt313.Fields("R31304").Value) Then
            Printer.Print ""
        Else
            Printer.Print adoaccrpt313.Fields("R31304").Value
        End If
        '期初餘額
        If IsNull(adoaccrpt313.Fields("R31306").Value) = False Then
            strAmount = Format(Val(adoaccrpt313.Fields("R31306").Value), DDollar2)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 7300 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
        End If
        '收入
        If IsNull(adoaccrpt313.Fields("R31307").Value) = False Then
            strAmount = Format(Val(adoaccrpt313.Fields("R31307").Value), DDollar2)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 8900 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
        End If
        '支出
        If IsNull(adoaccrpt313.Fields("R31308").Value) = False Then
            strAmount = Format(Val(adoaccrpt313.Fields("R31308").Value), DDollar2)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 10500 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
        End If
        '帳號餘額
        If IsNull(adoaccrpt313.Fields("R31309").Value) = False Then
            strAmount = Format(Val(adoaccrpt313.Fields("R31309").Value), DDollar2)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 12100 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
        End If
        '出名人
        Printer.CurrentX = 12300
        Printer.CurrentY = 300 + intCounter * 300
        If IsNull(adoaccrpt313.Fields("R31310").Value) Then
            Printer.Print ""
        Else
            Printer.Print adoaccrpt313.Fields("R31310").Value
        End If

        intCounter = intCounter + 1
        intRecord = intRecord + 1
        adoaccrpt313.MoveNext
   Loop
   
   If intRecord > 22 Then
        intPage = intPage + 1
        intRecord = 1
        Printer.NewPage
        intCounter = 3
        PrintHeadA4
        PrintHeadDA4
   End If
   '表尾合計
   Printer.FontBold = True
   intRecord = intRecord + 1
   Printer.CurrentX = 6800
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "合計:"
      
   Printer.FontBold = False
   Printer.Line (7700, 300 + intCounter * 300 - 50)-(10500, 300 + intCounter * 300 - 50)
   strAmount = Format(Val(SumAmt(0)), DDollar2)
   intLength = Printer.TextWidth(strAmount)
   Printer.CurrentX = 8900 - intLength
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print strAmount
   strAmount = Format(Val(SumAmt(1)), DDollar2)
   intLength = Printer.TextWidth(strAmount)
   Printer.CurrentX = 10500 - intLength
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print strAmount
   Printer.Line (7700, 300 + intCounter * 300 + 350)-(10500, 300 + intCounter * 300 + 350)
   Printer.Line (7700, 300 + intCounter * 300 + 400)-(10500, 300 + intCounter * 300 + 400)
   intCounter = intCounter + 2
    intRecord = intRecord + 2
      
   '分類明細表
   ado313Sum.MoveFirst
   If ado313Sum.RecordCount <> 0 Then
        isFirstG = True
                
        Do While ado313Sum.EOF = False
            If intRecord > 22 Then
                intPage = intPage + 1
                intRecord = 1
                Printer.NewPage
                intCounter = 3
                PrintHeadA4
                PrintHeadGA4
                If isFirstG = True Then isFirstG = False
            ElseIf isFirstG Then
                isFirstG = False
                PrintHeadGA4
            End If
            
           Printer.FontBold = False
           '出名人
            Printer.CurrentX = 0
            Printer.CurrentY = 300 + intCounter * 300
            If IsNull(ado313Sum.Fields("R31310").Value) Then
                Printer.Print ""
            Else
                Printer.Print ado313Sum.Fields("R31310").Value
            End If
            '金額
            If IsNull(ado313Sum.Fields("S31309").Value) = False Then
                strAmount = Format(Val(ado313Sum.Fields("S31309").Value), DDollar2)
                intLength = Printer.TextWidth(strAmount)
                Printer.CurrentX = 7400 - intLength
                Printer.CurrentY = 300 + intCounter * 300
                Printer.Print strAmount
            End If
            '所佔比率
            If IsNull(ado313Sum.Fields("P").Value) = False Then
                strAmount = Format(Val(ado313Sum.Fields("P").Value), FDollar)
                intLength = Printer.TextWidth(strAmount & "%")
                Printer.CurrentX = 9100 - intLength
                Printer.CurrentY = 300 + intCounter * 300
                Printer.Print strAmount & "%"
            End If
            
            intCounter = intCounter + 1
            intRecord = intRecord + 1
            ado313Sum.MoveNext
        Loop
        
        If intRecord > 22 Then
            intPage = intPage + 1
            intRecord = 1
            Printer.NewPage
            intCounter = 3
            PrintHeadA4
            PrintHeadGA4
        End If
        '分類明細合計
        Printer.FontBold = True
        intRecord = intRecord + 1
        Printer.CurrentX = 2200
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print "合計:"
        
        Printer.Line (6000, 300 + intCounter * 300 - 50)-(7500, 300 + intCounter * 300 - 50)
        strAmount = Format(Val(SumAmt(2)), DDollar2)
        intLength = Printer.TextWidth(strAmount)
        Printer.CurrentX = 7400 - intLength
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print strAmount
        Printer.Line (6000, 300 + intCounter * 300 + 350)-(7500, 300 + intCounter * 300 + 350)
        Printer.Line (6000, 300 + intCounter * 300 + 400)-(7500, 300 + intCounter * 300 + 400)
        
        intCounter = intCounter + 2
        Printer.FontBold = True
        Printer.CurrentX = 8200
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print "*** 結束 ***"
   End If
   
   Printer.EndDoc
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHeadA4()

   Printer.FontSize = 17
   Printer.FontBold = True
   Printer.CurrentX = 5000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print ReportTitle(313)
   
   Printer.FontSize = 12
   intCounter = intCounter + 2
   Printer.CurrentX = 5000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "銀行代號: " & Text1 & "~" & Text2
   intCounter = intCounter + 1
   Printer.CurrentX = 5000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "到期日期: " & MaskEdBox1
   
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印人員: " & StaffQuery(strUserNum)
   Printer.CurrentX = 12000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印日期: " & CFDate(ACDate(ServerDate))
   intCounter = intCounter + 1
   Printer.CurrentX = 12000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "頁次: " & intPage
  
End Sub

'*************************************************
'  明細表抬頭列印
'
'*************************************************
Private Sub PrintHeadDA4()

   intCounter = intCounter + 2
   Printer.CurrentX = 250
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "存款類別"
   Printer.CurrentX = 3000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "銀行名稱"
  
   Printer.CurrentX = 6200
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "期初餘額"
   Printer.CurrentX = 8000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "收　入"
   Printer.CurrentX = 9600
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "支　出"
   Printer.CurrentX = 11200
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "帳號餘額"
   Printer.CurrentX = 13000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "出名人"
   Printer.CurrentX = 14900
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "備　註"
   

   
   Printer.Line (0, 300 + intCounter * 300 + 350)-(16000, 300 + intCounter * 300 + 350)
   intCounter = intCounter + 2
End Sub

'*************************************************
'  分類表抬頭列印
'
'*************************************************
Private Sub PrintHeadGA4()

   Printer.FontSize = 12
  
   intCounter = intCounter + 2
   Printer.CurrentX = 2200
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "以戶名分類明細"
   Printer.CurrentX = 6400
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "金　額"
  
   Printer.CurrentX = 7900
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "所佔比率"
   
   Printer.Line (0, 300 + intCounter * 300 + 350)-(9200, 300 + intCounter * 300 + 350)
   intCounter = intCounter + 2
End Sub
