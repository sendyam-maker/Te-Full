VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc3440 
   AutoRedraw      =   -1  'True
   Caption         =   "銀行帳號別票據彙總表"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3570
   ScaleWidth      =   5160
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
      Top             =   1080
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
      Top             =   2640
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
      Top             =   2640
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
      Top             =   2280
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
      Top             =   2280
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
      Left            =   1320
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3240
      TabIndex        =   3
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
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "到期日期"
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
      TabIndex        =   15
      Top             =   600
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
      TabIndex        =   14
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1335
      Left            =   240
      Top             =   1800
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc3440.frx":0000
      Stretch         =   -1  'True
      Top             =   2640
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
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc3440.frx":0442
      Stretch         =   -1  'True
      Top             =   2280
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
      TabIndex        =   12
      Top             =   2280
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
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   135
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
Attribute VB_Name = "Frmacc3440"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0h0 As New ADODB.Recordset
Public adoacc0e0 As New ADODB.Recordset
Public adoaccrpt304 As New ADODB.Recordset
Public adoacc040 As New ADODB.Recordset
Public adoacc0b0 As New ADODB.Recordset
Dim strSort1, strSort2 As String
Dim dllaccrpt304 As Object

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
         'edit by nickc 2007/02/08
         'Combo7.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Command1_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   If MaskEdBox1.Text = MsgText(29) Or MaskEdBox2.Text = MsgText(29) Then
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt304Delete
   ProduceData
   If adoaccrpt304.State = adStateOpen Then
      adoaccrpt304.Close
   End If
   adoaccrpt304.CursorLocation = adUseClient
   adoaccrpt304.Open "select * from accrpt304", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt304.RecordCount <> 0 Then
      dllaccrpt304.Acc3440 ReportTitle(304), Text1, Text2, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   End If
   adoaccrpt304.Close
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
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
   Me.Height = 2200
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
   Combo4.AddItem MsgText(1)
   Combo4.AddItem MsgText(2)
   Combo6.AddItem MsgText(1)
   Combo6.AddItem MsgText(2)
   Combo4 = MsgText(1)
   Combo6 = MsgText(1)
   ComboAdd
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Set dllaccrpt304 = CreateObject("AccReport.ReportSelect")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt304 = Nothing
   Set Frmacc3440 = Nothing
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
   strSort1 = "銀行帳號"
   strSort2 = "銀行名稱"
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
Dim strOrder1 As String
Dim strOrder2 As String
Dim strSql As String
Dim intYear As Integer
Dim intMonth As Integer
Dim strStartDate As String
Dim strEdnDate As String
   
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
   Else
      If IsNull(adoacc0b0.Fields("a0b01").Value) Then
         If Mid(ServerDate, 5, 2) = 1 Then
            intMonth = 12
            intYear = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) - 1
         Else
            intMonth = Val(Mid(ServerDate, 5, 2)) - 1
            intYear = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3))
         End If
      Else
        intMonth = Val(Mid(CFDate(adoacc0b0.Fields("a0b01").Value), 5, 2))
        intYear = Val(Mid(CFDate(adoacc0b0.Fields("a0b01").Value), 1, 3))
      End If
   End If
   adoacc0b0.Close
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Select Case Combo3
      Case strSort1
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0h02 asc"
         Else
            strOrder1 = " order by a0h02 desc"
         End If
      Case strSort2
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0h03 asc"
         Else
            strOrder1 = " order by a0h03 desc"
         End If
      Case Else
         strOrder1 = MsgText(601)
   End Select
   Select Case Combo5
      Case strSort1
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0h02 asc"
         Else
            strOrder2 = ", a0h02 desc"
         End If
      Case strSort2
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0h03 asc"
         Else
            strOrder2 = ", a0h03 desc"
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
   adoaccrpt304.CursorLocation = adUseClient
   adoaccrpt304.Open "select * from accrpt304", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0h0.CursorLocation = adUseClient
   adoacc0h0.Open "select * from acc0h0, acc0g0 where a0h01 = a0g01" & strSql & strOrder1 & strOrder2, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0h0.RecordCount = 0 Then
      adoacc0h0.Close
      adoaccrpt304.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc0h0.EOF = False
      strSql = MsgText(601)
      'If adoacc0h0.Fields("a0g09").Value = MsgText(602) Then
      '   strStartDate = ACDate(Format(CDate(AFDate(CADate(FCDate(MaskEdBox1.Text)))) - 1, "YYYYMMDD"))
      '   strEndDate = ACDate(Format(CDate(AFDate(CADate(FCDate(MaskEdBox2.Text)))) - 1, "YYYYMMDD"))
      'Else
      '   strStartDate = ACDate(Format(CDate(AFDate(CADate(FCDate(MaskEdBox1.Text)))) - 3, "YYYYMMDD"))
      '   strEndDate = ACDate(Format(CDate(AFDate(CADate(FCDate(MaskEdBox2.Text)))) - 3, "YYYYMMDD"))
      'End If
      'If strStartDate <> MsgText(601) Then
      '   strSQL = strSQL & " and a0e10 >= " & Val(strStartDate) & ""
      'End If
      'If strEndDate <> MsgText(601) Then
      '   strSQL = strSQL & " and a0e10 <= " & Val(strEndDate) & ""
      'End If
      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
         strSql = strSql & " and a0e10 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      End If
      If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
         strSql = strSql & " and a0e10 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      End If
      adoaccrpt304.AddNew
      adoaccrpt304.Fields("r30401").Value = strUserNum
      adoaccrpt304.Fields("r30402").Value = adoacc0h0.Fields("a0h01").Value
      adoaccrpt304.Fields("r30403").Value = adoacc0h0.Fields("a0h02").Value
      adoaccrpt304.Fields("R30404").Value = A0g02Query(adoacc0h0.Fields("A0H01").Value)
      If IsNull(adoacc0h0.Fields("A0H08").Value) = False Then
         adoacc040.CursorLocation = adUseClient
         adoacc040.Open "select a0408 from acc040 where a0401 = " & intYear & " and a0403 = '1' and a0404 = '" & MsgText(55) & "' and a0405 = '" & adoacc0h0.Fields("a0h08").Value & "' and a0402 = " & intMonth & "", adoTaie, adOpenStatic, adLockReadOnly
         If adoacc040.RecordCount <> 0 Then
            adoacc040.MoveLast
            If IsNull(adoacc040.Fields(0).Value) Then
               adoaccrpt304.Fields("R30405").Value = 0
            Else
               adoaccrpt304.Fields("R30405").Value = adoacc040.Fields(0).Value
            End If
         Else
            adoaccrpt304.Fields("R30405").Value = 0
         End If
         adoacc040.Close
      Else
         adoaccrpt304.Fields("r30405").Value = 0
      End If
      adoacc0e0.CursorLocation = adUseClient
      adoacc0e0.Open "select sum(a0e11) from acc0e0 where a0e19 = '" & adoacc0h0.Fields("a0h01").Value & "' AND A0E20 = '" & adoacc0h0.Fields("A0H02").Value & "' and a0e04 = '" & MsgText(18) & "' and a0e15 = 0 and a0e17 = 0 and a0e21 = 0 and (a0e34 = 0 or a0e34 is null)" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0e0.RecordCount <> 0 Then
         If IsNull(adoacc0e0.Fields(0).Value) Then
            adoaccrpt304.Fields("r30406").Value = 0
         Else
            adoaccrpt304.Fields("r30406").Value = Val(adoacc0e0.Fields(0).Value)
         End If
      Else
         adoaccrpt304.Fields("r30406").Value = 0
      End If
      adoacc0e0.Close
      adoacc0e0.CursorLocation = adUseClient
      adoacc0e0.Open "select sum(a0e11) from acc0e0 where a0e01 = '" & adoacc0h0.Fields("a0h01").Value & "' and a0e07 = '" & adoacc0h0.Fields("a0h02").Value & "' and a0e04 = '" & MsgText(19) & "' and a0e37 = 0 and a0e25 = 0" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0e0.RecordCount <> 0 Then
         If IsNull(adoacc0e0.Fields(0).Value) Then
            adoaccrpt304.Fields("r30407").Value = 0
         Else
            adoaccrpt304.Fields("r30407").Value = Val(adoacc0e0.Fields(0).Value)
         End If
      Else
         adoaccrpt304.Fields("r30407").Value = 0
      End If
      adoaccrpt304.Fields("r30408").Value = Val(adoaccrpt304.Fields("r30405").Value) + Val(adoaccrpt304.Fields("r30406").Value) - Val(adoaccrpt304.Fields("r30407").Value)
      adoacc0e0.Close
      adoaccrpt304.UpdateBatch
      adoacc0h0.MoveNext
   Loop
   adoacc0h0.Close
   adoaccrpt304.Close
   adoTaie.Execute "delete from accrpt304 where r30402 is null"
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
Private Sub Accrpt304Delete()
   adoTaie.Execute "delete from accrpt304"
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
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
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
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function


