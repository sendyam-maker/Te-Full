VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc14d0 
   AutoRedraw      =   -1  'True
   Caption         =   "國內廠商付款明細表(目前不使用)"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2355
   ScaleWidth      =   5160
   Begin VB.ComboBox CboCmp 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   60
      Width           =   3525
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      Style           =   2  '單純下拉式
      TabIndex        =   10
      Top             =   1920
      Width           =   3450
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
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   1572
   End
   Begin VB.TextBox Text3 
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
      TabIndex        =   2
      Top             =   480
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      Default         =   -1  'True
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
      TabIndex        =   5
      Top             =   1320
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   840
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
      TabIndex        =   4
      Top             =   840
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
      Caption         =   "公司別"
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
      Left            =   330
      TabIndex        =   12
      Top             =   150
      Width           =   855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
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
      Left            =   300
      TabIndex        =   11
      Top             =   1950
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "廠商編號"
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
      Left            =   330
      TabIndex        =   9
      Top             =   510
      Width           =   975
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
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "入帳日期"
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
      Left            =   330
      TabIndex        =   7
      Top             =   870
      Width           =   975
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
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   840
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc14d0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/05/11 目前不使用-瑞婷
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoacc0o0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoaccrpt113 As New ADODB.Recordset
Dim intPage, intCounter, intLength As Integer, lngTotal, lngAmount As Long
Dim bolSum As Boolean, strAmount As String
'Add By Cheng 2003/05/28
Dim m_blnPrintData As Boolean '是否有資料要列印
Dim strPrinter As String 'Add By Sindy 2013/6/4
Dim strCmp As String, strCmpN As String 'Add by Sindy 2020/04/17


'Add by Sindy 2020/04/17
Private Sub SetCompN()
    strCmpN = "": strCmp = ""
    If Trim(CboCmp) <> MsgText(601) Then
        strCmp = CboCmp
        If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
        End If
    End If
    strCmpN = GetAccReportCmpN(strCmp, False, True)
End Sub

Private Sub CboCmp_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboCmp_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(CboCmp) = MsgText(601) Then Exit Sub
    
    strCmp = CboCmp
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp, strCmp) = 0 Then
        MsgBox Label1 & MsgText(63), , MsgText(5)
        Cancel = True
        CboCmp.SetFocus
        Exit Sub
    ElseIf Len(Trim(CboCmp)) = 1 Then
        CboCmp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/04/17

Private Sub Command1_Click()
Dim strSupplyNo As String
Dim bCancel As Boolean  'Add by Amy 2014/01/28

   'Add by Amy 2014/01/28 +公司別不可為空
   'Modify By Sindy 2020/4/23
   'If Text1 = MsgText(601) Then
   If CboCmp.Text = MsgText(601) Then
   '2020/4/23 END
      MsgBox Label1 & MsgText(52), , MsgText(5)
      CboCmp.SetFocus
      Exit Sub
   End If
   Call CboCmp_Validate(bCancel)
   If bCancel = True Then
      CboCmp.SetFocus
      Exit Sub
   End If
   'end 2014/01/28
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   
   Call SetCompN 'Add by Sindy 2020/04/23
   
   Screen.MousePointer = vbHourglass
   intPage = 0
   intCounter = 0
   lngTotal = 0
   lngAmount = 0
   Accrpt113Delete
   ProduceData
   'Add By Cheng 2003/05/28
   '若無資料則退出函式
   If m_blnPrintData = False Then
      Screen.MousePointer = vbDefault
      FormClear
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
      Exit Sub
   End If
   PUB_RestorePrinter Combo1 'Add By Sindy 2013/6/4
   Printer.Font = "新細明體"
   adoaccrpt113.CursorLocation = adUseClient
   adoaccrpt113.Open "select * from accrpt113 where r11301 = '" & strUserNum & "' order by r11302 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Do While adoaccrpt113.EOF = False
      If strSupplyNo <> adoaccrpt113.Fields("r11302").Value Then
         If lngAmount <> 0 Then
            PrintSum
            lngTotal = 0
            lngAmount = 0
            Printer.NewPage
         End If
         intCounter = 0
         intPage = intPage + 1
         PrintHead
         strSupplyNo = adoaccrpt113.Fields("r11302").Value
      End If
      Printer.CurrentX = 2500
      Printer.CurrentY = 4000 + intCounter * 300
      If IsNull(adoaccrpt113.Fields("r11304").Value) = False Then
         Printer.Print adoaccrpt113.Fields("r11304").Value
      End If
      If IsNull(adoaccrpt113.Fields("r11305").Value) = False Then
         strAmount = Format(adoaccrpt113.Fields("r11305").Value, DDollar)
         intLength = Printer.TextWidth(strAmount)
         Printer.CurrentX = 6000 - intLength
         Printer.CurrentY = 4000 + intCounter * 300
         Printer.Print strAmount
         lngAmount = lngAmount + Val(adoaccrpt113.Fields("r11305").Value)
      End If
      intCounter = intCounter + 1
      lngTotal = lngTotal + 1
      If intCounter > 20 Then
         intCounter = 0
         intPage = intPage + 1
         Printer.NewPage
         PrintHead
      End If
      adoaccrpt113.MoveNext
   Loop
   PrintSum
   adoaccrpt113.Close
   Printer.EndDoc
   PUB_RestorePrinter strPrinter 'Add By Sindy 2013/6/4
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
   Me.Height = 2760 '2200
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
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add By Sindy 2013/6/4
   
   'Add by Sindy 2020/04/17 公司別改下拉
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, False, False, False, , 1)
   'end 2020/04/17
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

   Set Frmacc14d0 = Nothing
End Sub

'Mark by Sindy 2020/4/23 公司別改下拉式選單
''Add by Amy 2014/01/28 +公司別
'Private Sub Text1_Change()
'    If Text1 = MsgText(601) Then
'        Text13 = ""
'        Exit Sub
'    End If
'    If Text1 = "1" Or Text1 = "J" Then
'        Text13 = A0802Query(Text1)
'    End If
'End Sub
'
'Private Sub Text1_GotFocus()
'    TextInverse Text1
'End Sub
'
'Private Sub Text1_KeyPress(KeyAscii As Integer)
'    KeyAscii = UpperCase(KeyAscii)
'End Sub
'
'Private Sub Text1_Validate(Cancel As Boolean)
'    If Text1 = "" Then Exit Sub
'    If Text1 <> "1" And Text1 <> "J" Then
'        Text13 = ""
'        MsgBox "公司別輸入錯誤請確認 ！"
'        Cancel = True
'        Exit Sub
'    End If
'End Sub
''end 2014/01/28

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim strSameName As String
Dim strSql As String

On Error GoTo Checking
   If Text2 <> MsgText(601) Then
      strSql = " and a0o03 >= '" & Text2 & "'"
   End If
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and a0o03 <= '" & Text3 & "'"
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0o05 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0o05 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoaccrpt113.CursorLocation = adUseClient
   adoaccrpt113.Open "select * from accrpt113", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0o0.CursorLocation = adUseClient
   'Modify by Amy 2014/01/28 +公司別
   adoacc0o0.Open "select * from acc0o0 where a0o07='" & strCmp & "' And a0o02 = '1'" & strSql & " order by a0o04 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0o0.RecordCount = 0 Then
      adoacc0o0.Close
      adoaccrpt113.Close
      MsgBox MsgText(28), , MsgText(5)
        'Add By Cheng 2003/05/28
        '無資料可列印
        m_blnPrintData = False
      Exit Sub
   End If
    'Add By Cheng 2003/05/28
    '有資料可列印
    m_blnPrintData = True
   Do While adoacc0o0.EOF = False
      If IIf(IsNull(adoacc0o0.Fields("a0o03").Value), MsgText(601), adoacc0o0.Fields("a0o03").Value) <> strSameName Then
         strSameName = IIf(IsNull(adoacc0o0.Fields("a0o03").Value), MsgText(601), adoacc0o0.Fields("a0o03").Value)
      End If
      adoaccrpt113.AddNew
      adoaccrpt113.Fields("r11301").Value = strUserNum
      If IsNull(adoacc0o0.Fields("a0o03").Value) Then
         adoaccrpt113.Fields("r11302").Value = Null
      Else
         adoaccrpt113.Fields("r11302").Value = adoacc0o0.Fields("a0o03").Value
         adoaccrpt113.Fields("r11303").Value = A0i02Query(adoacc0o0.Fields("a0o03").Value)
      End If
      If IsNull(adoacc0o0.Fields("a0o04").Value) Then
         adoaccrpt113.Fields("r11304").Value = Null
      Else
         adoaccrpt113.Fields("r11304").Value = adoacc0o0.Fields("a0o04").Value
      End If
      adoaccsum.CursorLocation = adUseClient
      'Modify by Amy 2014/01/28 改公司別 原:'1'
      adoaccsum.Open "select sum(a1p08) from acc1p0 where a1p01 ='" & adoacc0o0.Fields("a0o07").Value & "'  and a1p02 = 'B' and a1p04 = '" & adoacc0o0.Fields("a0o01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoaccsum.RecordCount <> 0 Then
         If IsNull(adoaccsum.Fields(0).Value) Then
            adoaccrpt113.Fields("r11305").Value = 0
         Else
            adoaccrpt113.Fields("r11305").Value = adoaccsum.Fields(0).Value
         End If
      Else
         adoaccrpt113.Fields("r11305").Value = 0
      End If
      adoaccsum.Close
      adoaccrpt113.UpdateBatch
      adoacc0o0.MoveNext
   Loop
   adoacc0o0.Close
   adoaccrpt113.Close
   adoTaie.Execute "delete from accrpt113 where r11302 is null"
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
Private Sub Accrpt113Delete()
   adoTaie.Execute "delete from accrpt113"
End Sub

''*************************************************
''  合計計算
''
''*************************************************
'Private Sub Calculate()
'Dim strSql As String
'
'   If MaskEdBox1.Text <> MsgText(601) Then
'      strSql = " and a0o05 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) Then
'      strSql = strSql & " and a0o05 <= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   adoaccrpt113.AddNew
'   adoaccrpt113.Fields("r11301").Value = strUserNum
'   If IsNull(adoacc0o0.Fields("a0o03").Value) Then
'      adoaccrpt113.Fields("r11302").Value = Null
'   Else
'      adoaccrpt113.Fields("r11302").Value = adoacc0o0.Fields("a0o03").Value
'   End If
'   adoaccrpt113.Fields("r11306").Value = ReportSum(4)
'   adoaccrpt113.UpdateBatch
'   adoaccrpt113.AddNew
'   adoaccrpt113.Fields("r11301").Value = strUserNum
'   If IsNull(adoacc0o0.Fields("a0o03").Value) Then
'      adoaccrpt113.Fields("r11302").Value = Null
'   Else
'      adoaccrpt113.Fields("r11302").Value = adoacc0o0.Fields("a0o03").Value
'   End If
'   adoaccsum.CursorLocation = adUseClient
'   'Modify by Amy 2014/01/28 +公司別
'   adoaccsum.Open "select count(*) from acc0o0 where a0o07='" & strCmp & "' And a0o02 = '" & adoacc0o0.Fields("a0o02").Value & "' and a0o03 = '" & adoacc0o0.Fields("a0o03").Value & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt113.Fields("r11304").Value = 0
'      Else
'         adoaccrpt113.Fields("r11304").Value = adoaccsum.Fields(0).Value
'      End If
'   Else
'      adoaccrpt113.Fields("r11304").Value = 0
'   End If
'   adoaccsum.Close
'   adoaccsum.CursorLocation = adUseClient
'   'Modify by Amy 2014/01/28 改公司別 原:'1'
'   adoaccsum.Open "select sum(a1p08) from acc1p0, acc0o0 where a1p03 = a0o01 and a1p01 ='" & adoacc0o0.Fields("a0o07").Value & "' and a1p02 = 'B' and a1p04 = 'B' and a0o02 = '1' and a0o03 = '" & adoacc0o0.Fields("a0o03").Value & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt113.Fields("r11305").Value = 0
'      Else
'         adoaccrpt113.Fields("r11305").Value = adoaccsum.Fields(0).Value
'      End If
'   Else
'      adoaccrpt113.Fields("r11305").Value = 0
'   End If
'   adoaccsum.Close
'   adoaccrpt113.UpdateBatch
'End Sub

'*************************************************
' 列印抬頭
'
'*************************************************
Private Sub PrintHead()
   Printer.FontSize = 16
   'Add by Amy 2014/01/28 +公司抬頭
   Printer.CurrentX = 3000
   Printer.CurrentY = 500
   Printer.Print strCmpN 'Text13
   'end 2014/02/28
   Printer.CurrentX = 3000
   Printer.CurrentY = 1000
   Printer.Print ReportTitle(113)
   Printer.FontSize = 12
   Printer.CurrentX = 3100
   Printer.CurrentY = 2000
   Printer.Print "入帳日期:"
   Printer.CurrentX = 4100
   Printer.CurrentY = 2000
   Printer.Print MaskEdBox1.Text
   Printer.CurrentX = 5100
   Printer.CurrentY = 2000
   Printer.Print "~"
   Printer.CurrentX = 5200
   Printer.CurrentY = 2000
   Printer.Print MaskEdBox2.Text
   Printer.CurrentX = 500
   Printer.CurrentY = 2500
   Printer.Print "列印人員:"
   Printer.CurrentX = 1600
   Printer.CurrentY = 2500
   Printer.Print StaffQuery(strUserNum)
   Printer.CurrentX = 7000
   Printer.CurrentY = 2500
   Printer.Print "列印日期:"
   Printer.CurrentX = 8100
   Printer.CurrentY = 2500
   Printer.Print CFDate(ACDate(ServerDate))
   Printer.CurrentX = 500
   Printer.CurrentY = 2800
   Printer.Print "廠商代號:"
   Printer.CurrentX = 1600
   Printer.CurrentY = 2800
   If IsNull(adoaccrpt113.Fields("r11302").Value) = False Then
      Printer.Print adoaccrpt113.Fields("r11302").Value
   End If
   Printer.CurrentX = 7000
   Printer.CurrentY = 2800
   Printer.Print "頁次:"
   Printer.CurrentX = 8100
   Printer.CurrentY = 2800
   Printer.Print intPage
   Printer.CurrentX = 500
   Printer.CurrentY = 3100
   Printer.Print "廠商名稱:"
   Printer.CurrentX = 1600
   Printer.CurrentY = 3100
   If IsNull(adoaccrpt113.Fields("r11303").Value) = False Then
      Printer.Print adoaccrpt113.Fields("r11303").Value
   End If
   Printer.CurrentX = 3000
   Printer.CurrentY = 3500
   Printer.Print "發票號碼"
   Printer.CurrentX = 5000
   Printer.CurrentY = 3500
   Printer.Print "發票金額"
   Printer.Line (2500, 3900)-(6500, 3900)
End Sub

'*************************************************
' 列印合計
'
'*************************************************
Private Sub PrintSum()
   Printer.Line (2500, 13000)-(6500, 13000)
   Printer.CurrentX = 2500
   Printer.CurrentY = 13100
   Printer.Print "合計:"
   strAmount = Format(lngTotal, DDollar)
   intLength = Printer.TextWidth(strAmount)
   Printer.CurrentX = 4100 - intLength
   Printer.CurrentY = 13100
   Printer.Print strAmount
   Printer.CurrentX = 4200
   Printer.CurrentY = 13100
   Printer.Print "筆"
   strAmount = Format(lngAmount, DDollar)
   intLength = Printer.TextWidth(strAmount)
   Printer.CurrentX = 6000 - intLength
   Printer.CurrentY = 13100
   Printer.Print strAmount
   Printer.Line (2500, 13500)-(6500, 13500)
   Printer.Line (2500, 13600)-(6500, 13600)
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text2 = ""
   Text3 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   'Modify by Amy 2014/01/28
   'Text2.SetFocus
'   Text1 = ""
'   Text13 = ""
'   Text1.SetFocus
   'end 2014/01/28
   'Add By Sindy 2020/4/23
   CboCmp.ListIndex = -1
   CboCmp.SetFocus
   '2020/4/23 END
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
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
   FormCheck = False
End Function

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
