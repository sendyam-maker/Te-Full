VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc34d0 
   AutoRedraw      =   -1  'True
   Caption         =   "日期別資金流動預測表"
   ClientHeight    =   3408
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5220
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3408
   ScaleWidth      =   5220
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   255
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
      Left            =   960
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo4 
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
      Left            =   3600
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo5 
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
      Left            =   960
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo6 
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
      Left            =   3600
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo7 
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
      Left            =   960
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo8 
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
      Left            =   3600
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   990
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   555
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
      Left            =   3240
      TabIndex        =   2
      Top             =   555
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
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "Label4"
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "(1.台一 2.智權)"
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
      Left            =   1700
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "公 司 別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "3."
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
      Left            =   720
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "排序方式"
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
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "1."
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
      Left            =   720
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image Image2 
      Height          =   252
      Left            =   3000
      Picture         =   "Frmacc34d0.frx":0000
      Stretch         =   -1  'True
      Top             =   2160
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "2."
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
      Left            =   720
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image Image3 
      Height          =   252
      Left            =   3000
      Picture         =   "Frmacc34d0.frx":0442
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Image Image4 
      Height          =   252
      Left            =   3000
      Picture         =   "Frmacc34d0.frx":0884
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1572
      Left            =   240
      Top             =   1680
      Visible         =   0   'False
      Width           =   4692
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   555
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "到期日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   585
      Width           =   1095
   End
End
Attribute VB_Name = "Frmacc34d0"
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
Public adoacc0b0 As New ADODB.Recordset
Public adoaccrpt314 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Dim strSort1, strSort2, strSort3 As String
Dim dllaccrpt314 As Object

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
         Combo7.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo7_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo8.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo8_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         MaskEdBox2.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Command1_Click()
         
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt314Delete
   ProduceData
   adoaccrpt314.CursorLocation = adUseClient
   adoaccrpt314.Open "select * from accrpt314", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt314.RecordCount <> 0 Then
      '20140120START Modify By eric
      dllaccrpt314.Acc34d0 ReportTitle(314) & "-" & Label4, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
      'dllaccrpt314.Acc34d0 ReportTitle(314), MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
      '20140120END
   End If
   adoaccrpt314.Close
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
   Me.Height = 2050 'Modify by Amy 2023/08/18 原:1900
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
   Combo8.AddItem MsgText(1)
   Combo8.AddItem MsgText(2)
   Combo4 = MsgText(1)
   Combo6 = MsgText(1)
   Combo8 = MsgText(1)
   ComboAdd
      
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Set dllaccrpt314 = CreateObject("AccReport.ReportSelect")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt314 = Nothing
   Set Frmacc34d0 = Nothing
End Sub

'*************************************************
'  Combo 項目新增
'
'*************************************************
Private Sub ComboAdd()
   strSort1 = "到期日期"
   strSort2 = "銀行代號"
   strSort3 = "銀行帳號"
   Combo3.AddItem strSort1
   Combo5.AddItem strSort2
   Combo5.AddItem strSort3
   Combo7.AddItem strSort2
   Combo7.AddItem strSort3
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim strOrder1, strOrder2, strOrder3 As String
Dim strSql As String
   
On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Select Case Combo3
      Case strSort1
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e10 asc"
         Else
            strOrder1 = " order by a0e10 desc"
         End If
      Case Else
         strOrder1 = MsgText(601)
   End Select
   Select Case Combo5
      Case strSort2
         If Combo6 = MsgText(1) Then
            strOrder2 = " order by a0h01 asc"
         Else
            strOrder2 = " order by a0h01 desc"
         End If
      Case strSort3
         If Combo6 = MsgText(1) Then
            strOrder2 = " order by a0h02 asc"
         Else
            strOrder2 = " order by a0h02 desc"
         End If
      Case Else
         strOrder2 = MsgText(601)
   End Select
   Select Case Combo7
      Case strSort2
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0h01 asc"
         Else
            strOrder3 = ", a0h01 desc"
         End If
      Case strSort3
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0h02 asc"
         Else
            strOrder3 = ", a0h02 desc"
         End If
      Case Else
         strOrder3 = MsgText(601)
   End Select
   
   '20140120START Modify By eric
   If Text1 <> MsgText(601) Then
      strSql = " and a0e23 = '" & IIf(Text1 = "2", "J", "1") & "' "
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0e10 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   'If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
   '   strSql = " and a0e10 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   'End If
   '20140120END
   
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0e10 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   
   adoaccrpt314.CursorLocation = adUseClient
   adoaccrpt314.Open "select * from accrpt314", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   adoacc0e0.CursorLocation = adUseClient
'   adoacc0e0.Open "select a0e10, sum(a0e11) from acc0e0" & strSQL & " group by a0e10" & strOrder1, adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoacc0e0.EOF = False
      adoacc0h0.CursorLocation = adUseClient
      adoacc0h0.Open "select * from acc0h0" & strOrder2 & strOrder3, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0h0.RecordCount = 0 Then
         adoacc0h0.Close
         adoaccrpt314.Close
         MsgBox MsgText(28), , MsgText(5)
         Exit Sub
      End If
      Do While adoacc0h0.EOF = False
         adoaccrpt314.AddNew
         adoaccrpt314.Fields("r31401").Value = strUserNum
'         If IsNull(adoacc0e0.Fields(0).Value) Then
'            adoaccrpt314.Fields("r31402").Value = Null
'         Else
'            adoaccrpt314.Fields("r31402").Value = adoacc0e0.Fields(0).Value
'         End If
         If IsNull(adoacc0h0.Fields("a0h01").Value) Then
            adoaccrpt314.Fields("r31403").Value = Null
         Else
            adoaccrpt314.Fields("r31403").Value = adoacc0h0.Fields("a0h01").Value
         End If
         If IsNull(adoacc0h0.Fields("a0h02").Value) Then
            adoaccrpt314.Fields("r31404").Value = Null
         Else
            adoaccrpt314.Fields("r31404").Value = adoacc0h0.Fields("a0h02").Value
         End If
         adoaccrpt314.Fields("r31405").Value = A0g02Query(adoacc0h0.Fields("a0h01").Value)
         adoaccsum.CursorLocation = adUseClient
         adoaccsum.Open "select sum(a0e11) from acc0e0 where a0e19 = '" & adoacc0h0.Fields("a0h01").Value & "' AND A0E20 = '" & adoacc0h0.Fields("A0H02").Value & "' and a0e04 = '" & MsgText(18) & "' and a0e15 = 0 and a0e17 = 0 and a0e21 = 0 and (a0e34 = 0 or a0e34 is null)" & strSql, adoTaie, adOpenStatic, adLockReadOnly
         If adoaccsum.RecordCount <> 0 Then
            If IsNull(adoaccsum.Fields(0).Value) Then
               adoaccrpt314.Fields("r31406").Value = 0
            Else
               adoaccrpt314.Fields("r31406").Value = Val(adoaccsum.Fields(0).Value)
            End If
         Else
            adoaccrpt314.Fields("r31406").Value = 0
         End If
         adoaccsum.Close
         adoaccsum.CursorLocation = adUseClient
         adoaccsum.Open "select sum(a0e11) from acc0e0 where a0e01 = '" & adoacc0h0.Fields("a0h01").Value & "' and a0e07 = '" & adoacc0h0.Fields("a0h02").Value & "' and a0e04 = '" & MsgText(19) & "' and a0e37 = 0 and a0e25 = 0" & strSql, adoTaie, adOpenStatic, adLockReadOnly
         If adoaccsum.RecordCount <> 0 Then
            If IsNull(adoaccsum.Fields(0).Value) Then
               adoaccrpt314.Fields("r31408").Value = 0
            Else
               adoaccrpt314.Fields("r31408").Value = Val(adoaccsum.Fields(0).Value)
            End If
         Else
            adoaccrpt314.Fields("r31408").Value = 0
         End If
         adoaccsum.Close
         adoacc0b0.CursorLocation = adUseClient
         '20140120START Modify By eric
         adoacc0b0.Open "select a0b02 from acc0b0 where a0b04 = '" & IIf(Text1 = "2", "J", "1") & "' ", adoTaie, adOpenStatic, adLockReadOnly
         'adoacc0b0.Open "select a0b02 from acc0b0", adoTaie, adOpenStatic, adLockReadOnly
         '20140120END
         
         If adoacc0b0.RecordCount <> 0 Then
            If IsNull(adoacc0b0.Fields(0).Value) Then
               adoaccrpt314.Fields("r31410").Value = 0
            Else
               adoaccsum.CursorLocation = adUseClient
               '20140120START Modify By eric
               adoaccsum.Open "select sum(a0408) from acc040 where a0401 = " & Val(Mid(CFDate(adoacc0b0.Fields(0).Value), 1, 3)) & " and a0402 = " & Val(Mid(CFDate(adoacc0b0.Fields(0).Value), 5, 2)) & " AND A0403 ='" & IIf(Text1 = "2", "J", "1") & "' and a0405 = '" & adoacc0h0.Fields("a0h08").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
               'adoaccsum.Open "select sum(a0408) from acc040 where a0401 = " & Val(Mid(CFDate(adoacc0b0.Fields(0).Value), 1, 3)) & " and a0402 = " & Val(Mid(CFDate(adoacc0b0.Fields(0).Value), 5, 2)) & " and a0403 = '1' and a0405 = '" & adoacc0h0.Fields("a0h08").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
               '20140120END
               If adoaccsum.RecordCount <> 0 Then
                  If IsNull(adoaccsum.Fields(0).Value) Then
                     adoaccrpt314.Fields("r31410").Value = 0
                  Else
                     adoaccrpt314.Fields("r31410").Value = Val(adoaccsum.Fields(0).Value)
                  End If
               Else
                  adoaccrpt314.Fields("r31410").Value = 0
               End If
               adoaccsum.Close
            End If
         Else
            adoaccrpt314.Fields("r31410").Value = 0
         End If
         adoacc0b0.Close
         adoaccrpt314.Fields("r31411").Value = Val(adoaccrpt314.Fields("r31406").Value) - Val(adoaccrpt314.Fields("r31408").Value) + Val(adoaccrpt314.Fields("r31410").Value)
         adoaccrpt314.UpdateBatch
         adoacc0h0.MoveNext
      Loop
      adoacc0h0.Close
'      adoacc0e0.MoveNext
'   Loop
'   adoacc0e0.Close
   adoaccrpt314.Close
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
Private Sub Accrpt314Delete()
   adoTaie.Execute "delete from accrpt314"
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
   Combo3 = ""
   Combo5 = ""
   Combo7 = ""
'20140120START Modify By eric
   Label4 = ""
   Text1.Text = ""
   Text1.SetFocus
   'MaskEdBox1.SetFocus
'20140120END Modify By eric
   
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
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

'20140120START By eric
Private Sub Text1_LostFocus()
   If Text1.Text = "" Then
      MsgBox "公司別不可空白 !"
      Text1.SetFocus
      Exit Sub
   End If
   If Text1.Text <> "1" And Text1.Text <> "2" Then
      MsgBox "公司別僅能為 1 或 2 !"
      Text1.Text = ""
      Text1.SetFocus
      Exit Sub
   End If
End Sub

'20140120START By eric
Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

'20140120START By eric
Private Sub Text1_Change()
   Label4.Caption = A0802Query(IIf(Text1 = "2", "J", "1"))
End Sub
