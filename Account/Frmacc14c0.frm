VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc14c0 
   AutoRedraw      =   -1  'True
   Caption         =   "應付款統計表"
   ClientHeight    =   3024
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3024
   ScaleWidth      =   5160
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   3500
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1320
      Style           =   2  '單純下拉式
      TabIndex        =   16
      Top             =   2390
      Width           =   3450
   End
   Begin VB.TextBox Text4 
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
      Left            =   1320
      TabIndex        =   6
      Top             =   1560
      Width           =   615
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
      TabIndex        =   10
      Top             =   1920
      Width           =   4692
   End
   Begin VB.TextBox Text3 
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
      Left            =   3240
      TabIndex        =   3
      Top             =   840
      Width           =   1572
   End
   Begin VB.TextBox Text2 
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
      Left            =   1320
      TabIndex        =   2
      Top             =   840
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
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   612
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Top             =   1200
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
      TabIndex        =   5
      Top             =   1200
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
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
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
      TabIndex        =   18
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
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
      Left            =   300
      TabIndex        =   17
      Top             =   2426
      Width           =   972
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "(1: 全部 2: 未付)"
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
      Left            =   2040
      TabIndex        =   15
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "列印方式"
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
      TabIndex        =   14
      Top             =   1560
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   132
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
      TabIndex        =   13
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "入帳日期"
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
      TabIndex        =   12
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label7 
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
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "往來對象"
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
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "(1.廠商 2.客戶 3.員工)"
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
      Left            =   2040
      TabIndex        =   8
      Top             =   510
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "往來類別"
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
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc14c0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoacc0o0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoaccrpt112 As New ADODB.Recordset
Dim dllaccrpt112 As Object
Dim strPrinter As String 'Add By Sindy 2013/6/4

Private Sub Combo2_GotFocus()
    TextInverse Combo2
End Sub

'Add by Amy 2020/04/08
Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(Combo2) = MsgText(601) Then Exit Sub
    
    strCmp = Combo2
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp, strCmp) = 0 Then
        MsgBox Label9 & MsgText(63), , MsgText(5)
        Cancel = True
        Combo2.SetFocus
        Exit Sub
    ElseIf Len(Trim(Combo2)) = 1 Then
        Combo2 = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub

Private Sub Command1_Click()
   'Add by Amy 2014/02/06 +公司別
   Dim bCancel As Boolean
   Dim strCmp As String 'Add by Amy 2020/04/08
   
   'Modify by Amy 2020/04/08 改下拉 原:Text5
   'If Text5 = MsgText(601) Then
   If Trim(Combo2) = MsgText(601) Then
      MsgBox Label9 & MsgText(52), , MsgText(5)
      Exit Sub
   End If
   Call Combo2_Validate(bCancel)
   If bCancel = True Then
      Exit Sub
   End If
   'end 2020/04/08
   'end 2014/02/06
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt112Delete
   ProduceData
   PUB_SetOsDefaultPrinter Combo1 'Add By Sindy 2013/6/4
   'Modify by Amy 2020/04/08 公司別改下拉,名稱原:Text13
   If Trim(Combo2) <> MsgText(601) Then
      strCmp = Combo2
      If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
      End If
      strCmp = A0802Query(strCmp)
    End If
   dllaccrpt112.Acc14c0 strCmp & "," & ReportTitle(112), MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   'end 2020/04/08
   PUB_SetOsDefaultPrinter strPrinter 'Add By Sindy 2013/6/4
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5280 '5250
   Me.Height = 3435 '3000
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Add by Amy 2020/04/08
   Combo2.AddItem "", 0
   Call Pub_SetCboCmp(Combo2, False, False, False, , 1)
   'end 2020/04/08
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   Set dllaccrpt112 = CreateObject("AccReport.ReportSelect")
   
   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add By Sindy 2013/6/4
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   
   'Add By Sindy 2013/6/4
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '2013/6/4 END

   Set dllaccrpt112 = Nothing
   Set Frmacc14c0 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text1 = Mid(ComboItem(92), 1, 1) Then
      If Len(Text1) = 6 Then
         Text1 = AfterZero(Text1)
      End If
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim lngPreAmount As Long
Dim lngPayAmount As Long
Dim strSql As String
Dim strSum As String
Dim strCmp As String 'Add by Amy 2020/04/08

On Error GoTo Checking
   If Text1 <> MsgText(601) Then
      strSql = " and a0o02 = '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0o03 >= '" & Text2 & "'"
   End If
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and a0o03 <= '" & Text3 & "'"
   End If
   If MaskEdBox1.Text <> MsgText(601) Then
      strSql = strSql & " and a0o05 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0o05 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Text4 = "2" Then
      strSql = strSql & " and (a0o11 is null or a0o11 = 0)"
   End If
   If strSql <> MsgText(601) Then
      strSum = strSql
      strSql = " where " & Mid(strSql, 5, Len(strSql) - 4)
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoacc0o0.CursorLocation = adUseClient
   'Modify by Amy 2014/02/06 +公司別
   'Modify by Amy 2020/04/08 公司別改下拉 原:Text5
   strCmp = Combo2
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
   adoacc0o0.Open "select a0o02, a0o03 from acc0o0" & strSql & " And a0o07='" & strCmp & "' group by a0o02, a0o03", adoTaie, adOpenStatic, adLockReadOnly
   'end 2020/04/08
   If adoacc0o0.RecordCount = 0 Then
      adoacc0o0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc0o0.EOF = False
      adoaccrpt112.CursorLocation = adUseClient
      adoaccrpt112.Open "select * from accrpt112 where r11201 = '" & strUserNum & "' and r11202 = '" & adoacc0o0.Fields("a0o03").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adoaccrpt112.RecordCount = 0 Then
         adoaccrpt112.AddNew
         CustomerSave
      End If
      adoaccsum.CursorLocation = adUseClient
      'Modify by Amy 2014/02/06 改公司別 原:'1'
      'Modify by Amy 2020/04/08 公司別改下拉
      adoaccsum.Open "select sum(a1p08) from acc1p0, acc0o0 where a1p04 = a0o01 and a1p01 = '" & strCmp & "' and a1p02 = 'B' and a0o03 = '" & adoacc0o0.Fields("a0o03").Value & "' and a1p05 in ('2112', '2113')" & strSum & " union " & _
                     "select sum(a1p08) from acc1p0, acc0o0 where a1p04 = a0o09 and a1p01 = '" & strCmp & "' and a1p02 = 'E' and a0o03 = '" & adoacc0o0.Fields("a0o03").Value & "' and a1p05 in ('2112', '2113')" & strSum & " union " & _
                     "select sum(a1p08) from acc1p0, acc0o0 where a1p04 = a0o09 and a1p01 = '" & strCmp & "'  and a1p02 = 'Z' and a0o03 = '" & adoacc0o0.Fields("a0o03").Value & "' and a1p05 in ('2112', '2113')" & strSum, adoTaie, adOpenStatic, adLockReadOnly
      If adoaccsum.RecordCount <> 0 Then
         If IsNull(adoaccsum.Fields(0).Value) Then
            lngPreAmount = 0
         Else
            lngPreAmount = Val(adoaccsum.Fields(0).Value)
         End If
      Else
         lngPreAmount = 0
      End If
      adoaccsum.Close
      adoaccrpt112.Fields("r11204").Value = lngPreAmount
      adoaccsum.CursorLocation = adUseClient
      'Modify by Amy 2014/02/06 改公司別 原:'1'
      'Modify by Amy 2020/04/08 公司別改下拉
      adoaccsum.Open "select sum(a1p08) from acc1p0, acc0o0 where a1p04 = a0o01 and a1p01 = '" & strCmp & "' and a1p02 = 'B' and a0o03 = '" & adoacc0o0.Fields("a0o03").Value & "' and (a0o11 is not null and a0o11 <> 0) and a1p05 in ('2112', '2113')" & strSum & " union " & _
                     "select sum(a1p08) from acc1p0, acc0o0 where a1p04 = a0o09 and a1p01 = '" & strCmp & "' and a1p02 = 'E' and a0o03 = '" & adoacc0o0.Fields("a0o03").Value & "' and (a0o11 is not null and a0o11 <> 0) and a1p05 in ('2112', '2113')" & strSum & " union " & _
                     "select sum(a1p08) from acc1p0, acc0o0 where a1p04 = a0o09 and a1p01 = '" & strCmp & "' and a1p02 = 'Z' and a0o03 = '" & adoacc0o0.Fields("a0o03").Value & "' and (a0o11 is not null and a0o11 <> 0) and a1p05 in ('2112', '2113')" & strSum, adoTaie, adOpenStatic, adLockReadOnly
      If adoaccsum.RecordCount <> 0 Then
         If IsNull(adoaccsum.Fields(0).Value) Then
            adoaccrpt112.Fields("r11205").Value = 0
         Else
            adoaccrpt112.Fields("r11205").Value = adoaccsum.Fields(0).Value
         End If
      Else
         adoaccrpt112.Fields(0).Value = 0
      End If
      adoaccsum.Close
      adoaccrpt112.UpdateBatch
      adoaccrpt112.Close
      adoacc0o0.MoveNext
   Loop
   adoacc0o0.Close
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
Private Sub Accrpt112Delete()
   adoTaie.Execute "delete from accrpt112"
End Sub

'*************************************************
'  往來對象儲存
'
'*************************************************
Private Sub CustomerSave()
   adoaccrpt112.Fields("r11201").Value = strUserNum
   If IsNull(adoacc0o0.Fields("a0o03").Value) Then
      adoaccrpt112.Fields("r11202").Value = Null
   Else
      adoaccrpt112.Fields("r11202").Value = adoacc0o0.Fields("a0o03").Value
      Select Case adoacc0o0.Fields("a0o02").Value
         Case Mid(ComboItem(91), 1, 1)
            adoaccrpt112.Fields("r11203").Value = A0i02Query(adoacc0o0.Fields("a0o03").Value)
         Case Mid(ComboItem(92), 1, 1)
            adoaccrpt112.Fields("r11203").Value = CustomerQuery(adoacc0o0.Fields("a0o03").Value, 1)
         Case Mid(ComboItem(93), 1, 1)
            adoaccrpt112.Fields("r11203").Value = StaffQuery(adoacc0o0.Fields("a0o03").Value)
      End Select
   End If
   adoaccrpt112.Fields("r11204").Value = 0
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If Text1 = Mid(ComboItem(92), 1, 1) Then
      If Len(Text3) = 6 Then
         Text3 = AfterZero(Text3)
      End If
   End If
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   'Modify by Amy 2020/04/08 公司別改下拉
   'Add by Amy 2014/02/06
'   Text5 = ""
'   Text13 = ""
   'end 2014/02/06
   Combo2 = ""
   'end 2020/04/08
   Text1 = ""
   Text2 = ""
   Text3 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text4 = ""
   Combo2.SetFocus 'Modify by Amy 2020/04/08 原:Text5
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
   If Text3 <> MsgText(601) Then
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
   If Text4 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

'Mark by Amy 2020/04/08 公司別改下拉
'Add by Amy 2014/02/06
'Private Sub Text5_Change()
'    If Text5 = MsgText(601) Then
'        Text13 = ""
'        Exit Sub
'    End If
'    If Text5 = "1" Or Text5 = "J" Then
'        Text13 = A0802Query(Text5)
'    End If
'End Sub
'
'Private Sub Text5_GotFocus()
'    TextInverse Text5
'End Sub
'
'Private Sub Text5_KeyPress(KeyAscii As Integer)
'    KeyAscii = UpperCase(KeyAscii)
'End Sub
'
'Private Sub Text5_Validate(Cancel As Boolean)
'    If Text5 = "" Then Exit Sub
'    If Text5 <> "1" And Text5 <> "J" Then
'        Text13 = ""
'        MsgBox "公司別輸入錯誤請確認 ！"
'        Cancel = True
'        Exit Sub
'    End If
'End Sub
''end 2014/02/06
'end 2020/04/08
