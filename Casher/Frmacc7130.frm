VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc7130 
   AutoRedraw      =   -1  'True
   Caption         =   "分所智權人員收款明細表"
   ClientHeight    =   3390
   ClientLeft      =   3660
   ClientTop       =   1875
   ClientWidth     =   5640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3390
   ScaleWidth      =   5640
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
      Left            =   2220
      MaxLength       =   1
      TabIndex        =   14
      Text            =   "Y"
      Top             =   1770
      Width           =   525
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1485
      TabIndex        =   12
      Top             =   2880
      Width           =   3495
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
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1350
      Width           =   945
   End
   Begin VB.OptionButton Option1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   780
      Width           =   195
   End
   Begin VB.OptionButton Option1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   330
      Value           =   -1  'True
      Width           =   195
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
      Left            =   450
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   2370
      Width           =   4692
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
      Left            =   1575
      MaxLength       =   5
      TabIndex        =   1
      Top             =   420
      Width           =   945
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1590
      TabIndex        =   3
      Top             =   870
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
      Left            =   3510
      TabIndex        =   4
      Top             =   870
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
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "是否列印明細資料：         (Y/N)"
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
      Left            =   210
      TabIndex        =   15
      Top             =   1800
      Width           =   3795
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   " 印表機"
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
      Left            =   450
      TabIndex        =   13
      Top             =   2895
      Width           =   975
   End
   Begin VB.Label lblSalesName 
      BackStyle       =   0  '透明
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
      Left            =   2250
      TabIndex        =   11
      Top             =   1380
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
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
      Left            =   180
      TabIndex        =   10
      Top             =   1380
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label3 
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
      Left            =   3270
      TabIndex        =   9
      Top             =   870
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "收款日期"
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
      Left            =   450
      TabIndex        =   8
      Top             =   870
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收款年月                  (Ex : 9301)"
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
      Left            =   450
      TabIndex        =   7
      Top             =   420
      Width           =   3795
   End
End
Attribute VB_Name = "Frmacc7130"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit
Public adoacc310 As New ADODB.Recordset
Dim strSql, strNo, strSalesNo, strSalesName As String
Dim intLength As Integer
Dim intCounter As Integer
Dim intPage As Integer
Dim PLeft(0 To 17) As String
Dim prnPrint As Printer
Dim strPrint As String

Private Sub Command1_Click()
    If FormCheck = False Then
        MsgBox MsgText(181), , MsgText(5)
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    For Each prnPrint In Printers
       If prnPrint.DeviceName = Combo1 Then
          Set Printer = prnPrint
       End If
    Next
    PrintDetail
    For Each prnPrint In Printers
       If prnPrint.DeviceName = strPrint Then
          Set Printer = prnPrint
       End If
    Next
    Screen.MousePointer = vbDefault
    FormClear
    Frmacc0000.StatusBar1.Panels(1).Text = "列印A4報表"
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
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim ii As Integer
   
    Me.Icon = LoadPicture(strIcoPath)
    strFormName = Name
    Me.Width = 5760
    Me.Height = 3795
    Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
    Image1 = LoadPicture(strBackPicPath4)
    sglWidth = Image1.Width
    sglHeight = Image1.Height
    For intX = 0 To Int(ScaleWidth / sglWidth)
        For intY = 0 To Int(ScaleHeight / sglHeight)
            PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
        Next
    Next
    Option1_Click 0
    Text1 = Left(strSrvDate(1), 6) - 191100
    MaskEdBox1.Mask = DFormat
    MaskEdBox2.Mask = DFormat
    Frmacc0000.StatusBar1.Panels(1).Text = "列印A4報表"
    SendKeys "{Tab}"
   strPrint = Printer.DeviceName
   For Each prnPrint In Printers
      'edit by nick 2004/11/11
      'If prnPrint.DeviceName <> Printer.DeviceName Then
         Combo1.AddItem prnPrint.DeviceName
      'End If
      If Combo1 = "" Then
         'edit by nick 2004/11/11
         'Combo1 = prnPrint.DeviceName
         Combo1 = Printer.DeviceName
      End If
   Next
    '設定列印印表機
    StrSQLa = "Select * From PrintStartPoint Where PSP01='" & strUserNum & "' And PSP02='" & Me.Name & "' And PSP03='" & Me.Combo1.Name & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    '若有資料
    If rsA.RecordCount > 0 Then
        If Me.Combo1.ListCount > 0 Then
            For ii = 0 To Me.Combo1.ListCount - 1
                If Me.Combo1.List(ii) = "" & rsA("PSP06").Value Then
                    Me.Combo1.ListIndex = ii
                    Exit For
                End If
            Next ii
        End If
        '記錄原設定值
        Me.Combo1.Tag = Me.Combo1.Text
    '若無資料
    Else
        '記錄原設定值
        Me.Combo1.Tag = ""
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '若印表機變動, 則更新列印設定
    If Me.Combo1.Text <> Me.Combo1.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
    End If
    strFormName = MsgText(601)
    KeyEnter vbKeyEscape
    MenuEnabled
    StatusClear
    Set Frmacc7130 = Nothing
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    If MaskEdBox1.Text = MsgText(29) Then
        Exit Sub
    End If
    MaskEdBox2.Mask = ""
    MaskEdBox2.Mask = DFormat
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
    Case 0
        Me.Text1.Enabled = True
        Me.Option1(1).Value = False
        Me.MaskEdBox1.Enabled = False
        Me.MaskEdBox2.Enabled = False
    Case 1
        Me.MaskEdBox1.Enabled = True
        Me.MaskEdBox2.Enabled = True
        Me.Option1(0).Value = False
        Me.Text1.Enabled = False
    End Select
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If Me.Text1.Text <> "" Then
        If IsDate(Format(Val(Me.Text1.Text & "01") + 19110000, "####/##/##")) = False Then
            MsgBox "收款年月輸入錯誤!!!", vbExclamation + vbOKOnly
            Cancel = True
        End If
    End If
    If Cancel = True Then Text1_GotFocus
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHead()
Dim strLocation As String

    GetPleft
    Printer.CurrentX = PLeft(3)
    Printer.CurrentY = 500
    Printer.Print IIf(pub_strUserOffice = "1", "北所", IIf(pub_strUserOffice = "2", "中所", IIf(pub_strUserOffice = "3", "南所", IIf(pub_strUserOffice = "4", "高所", "其他所")))) & _
                     IIf(Me.Option1(0).Value = True, " " & Left(Me.Text1.Text, Len(Me.Text1.Text) - 2) & " 年 " & Right(Me.Text1.Text, 2) & " 月 份", " " & Me.MaskEdBox1.Text & " ~ " & Me.MaskEdBox2.Text & " ") & "智權人員收款明細表"
    Printer.CurrentX = PLeft(0)
    Printer.CurrentY = 800
    'add by nick 2005/01/14  加入控制列不列印明細
    If Text3.Text = "Y" Then
        Printer.Print "智權人員：" & strSalesName
    Else
        Printer.Print "智權人員："
    End If
    Printer.CurrentX = PLeft(17)
    Printer.CurrentY = 800
    Printer.Print "列印日期：" & CFDate(strSrvDate(2))
           
    Printer.CurrentX = PLeft(0)
    Printer.CurrentY = 1100
    Printer.Print "收款日"
    Printer.CurrentX = PLeft(1)
    Printer.CurrentY = 1100
    Printer.Print "收款人"
    Printer.CurrentX = PLeft(2) + 6 * 90 - Printer.TextWidth("現金")
    Printer.CurrentY = 1100
    Printer.Print "客戶名稱"
    Printer.CurrentX = PLeft(3)
    Printer.CurrentY = 1100
    Printer.Print "案件性質"
    
    Printer.CurrentX = PLeft(17)
    Printer.CurrentY = 1100
    Printer.Print "頁　　數：" & intPage

    Printer.CurrentX = PLeft(4)
    Printer.CurrentY = 1400
    Printer.Print "人工號"
    Printer.CurrentX = PLeft(5)
    Printer.CurrentY = 1400
    Printer.Print "電腦號"
    Printer.CurrentX = PLeft(6) + 6 * 90 - Printer.TextWidth("現金")
    Printer.CurrentY = 1400
    Printer.Print "現金"
    Printer.CurrentX = PLeft(7) + 6 * 90 - Printer.TextWidth("支票")
    Printer.CurrentY = 1400
    Printer.Print "支票"
    Printer.CurrentX = PLeft(8)
    Printer.CurrentY = 1400
    Printer.Print "到期日"
    Printer.CurrentX = PLeft(9)
    Printer.CurrentY = 1400
    Printer.Print "帳號"
    Printer.CurrentX = PLeft(10)
    Printer.CurrentY = 1400
    Printer.Print "票號"
    Printer.CurrentX = PLeft(11)
    Printer.CurrentY = 1400
    Printer.Print "付款地"
    Printer.CurrentX = PLeft(12)
    Printer.CurrentY = 1400
    Printer.Print "扣繳日"
    Printer.CurrentX = PLeft(13) + 6 * 90 - Printer.TextWidth("扣繳額")
    Printer.CurrentY = 1400
    Printer.Print "扣繳額"
    Printer.CurrentX = PLeft(14) + 6 * 90 - Printer.TextWidth("留" & IIf(pub_strUserOffice = "1", "北所", IIf(pub_strUserOffice = "2", "中所", IIf(pub_strUserOffice = "3", "南所", IIf(pub_strUserOffice = "4", "高所", "其他所")))))
    Printer.CurrentY = 1400
    Printer.Print "留" & IIf(pub_strUserOffice = "1", "北所", IIf(pub_strUserOffice = "2", "中所", IIf(pub_strUserOffice = "3", "南所", IIf(pub_strUserOffice = "4", "高所", "其他所"))))
    Printer.CurrentX = PLeft(15) + 6 * 90 - Printer.TextWidth("點數")
    Printer.CurrentY = 1400
    Printer.Print "點數"
    Printer.CurrentX = PLeft(16)
    Printer.CurrentY = 1400
    Printer.Print "會計簽名"
    Printer.CurrentX = PLeft(17)
    Printer.CurrentY = 1400
    Printer.Print "備　　註"
    
    Printer.CurrentX = PLeft(0)
    Printer.CurrentY = 1700
    Printer.Print String(170, "=")
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
    If Me.Option1(0).Value = True Then
        Text1.SetFocus
    Else
        Me.MaskEdBox1.SetFocus
    End If
End Sub

'*************************************************
' 列印明細
'
'*************************************************
Private Sub PrintDetail()
Dim intCounter As Integer
Dim strName As String
Dim strOurCaseNo As String
Dim strCaseData As String
Dim strAmt As String
Dim strPoint As String
Dim dblCash, dblTCash As Double
Dim dblCheck, dblTCheck As Double
Dim dblTOT, dblTTOT As Double
'add by nick 2004/08/20
Dim dblPoint, dblTPoint As Double
'add by nick 2004/10/20
Dim dblMoney, dblTMoney As Double
    strSql = ""
    strNo = "": strSalesNo = "": strSalesName = ""
    strOurCaseNo = ""
    strCaseData = ""
    strAmt = ""
    strPoint = ""
    dblCash = 0: dblCheck = 0: dblTOT = 0
    dblTCash = 0: dblTCheck = 0: dblTTOT = 0
    'add by nick 2004/08/20
    dblPoint = 0: dblTPoint = 0
    'add by nick 2004/10/20
    dblMoney = 0: dblTMoney = 0
    If Me.Option1(0).Value = True Then
        If Text1 <> "" Then
            strSql = strSql & " And A3102>=" & Val(Me.Text1.Text & "01") & " And A3102<=" & Val(Me.Text1.Text & "31") & " "
        End If
    Else
        If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
            strSql = strSql & " And A3102>=" & Val(FCDate(MaskEdBox1.Text)) & " "
        End If
        If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
            strSql = strSql & " And A3102<=" & Val(FCDate(MaskEdBox2.Text)) & " "
        End If
    End If
    If Me.Text2.Text <> "" Then
        'edit by nick 2004/12/30
        'StrSql = StrSql & " And A0K20='" & Me.Text2.Text & "' "
        strSql = strSql & " And A3121='" & Me.Text2.Text & "' "
    End If
    Printer.Orientation = 2
    Printer.Font.Name = "新細明體"
    Printer.FontSize = 9
    'edit by nick 2004/08/20 讓分所可以查其他所 cancel
    'strSQL = "Select A3102 As 收款日, ST02 As 收款人, A0K04 As 收據抬頭, A0J02 As 本所案號, A0J20 As 案件性質名稱, Nvl(A0J09,0)+Nvl(A0J10,0) As 費用, A3104 As 人工號, A3103 As 電腦號, A3105 As 現金, A3106 As 支票, A3107 As 到期日, A3108 As 帳號, A3109 As 票號, A3110 As 付款地, A3111 As 扣繳日, A3112 As 扣繳額, A3113 As 留分所金額, Round(Nvl(A0J09,0)/1000,1) As 點數, A0J09, A0J10, A0K20 From ACC310, ACC0k0, ACC0J0, Staff Where A3103=A0K01(+) And A0K01=A0J13(+) And A0K20=ST01(+) And A3101='" & pub_strUserOffice & "' " & strSQL
    'strSQL = strSQL & " Order By ST03, A0K20, A3101, A3102, A3103, A3104 "
    'Modified by Morgan 2011/12/27 取消 a0j20
    strSql = "Select A3102 As 收款日, ST02 As 收款人, A3122 As 收據抬頭, A0J02 As 本所案號, getcp10desc(cp01,cp10,a0j04) As 案件性質名稱, Nvl(A0J09,0)+Nvl(A0J10,0) As 費用, A3104 As 人工號, A3103 As 電腦號, A3105 As 現金, A3106 As 支票, A3107 As 到期日, A3108 As 帳號, A3109 As 票號, A3110 As 付款地, A3111 As 扣繳日, A3112 As 扣繳額, A3113 As 留分所金額, A3123 As 點數, A0J09, A0J10, A3121,A3124 as 備註 From ACC310,  ACC0J0, Staff,caseprogress Where A3103=A0J13(+) And A3121=ST01(+)  And A3101='" & pub_strUserOffice & "' " & strSql & " and cp09(+)=a0j01 "
    '93.10.7 MODIFY BY SONIA
    'strSQL = strSQL & " Order By ST03, A3121, A3101, A3102, A3103, A3104 "
    '94.1.4 MODIFY BY SONIA
    'StrSql = StrSql & " Order By ST15, A3121, A3101, A3102, A3103, A3104 "
    strSql = strSql & " Order By ST15, A3121, A3101, A3115, A3116, A3103, A3104 "
    '94.1.4 END
    '93.10.7 END
    adoacc310.CursorLocation = adUseClient
    adoacc310.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
    If adoacc310.RecordCount = 0 Then
        adoacc310.Close
        MsgBox MsgText(28), , MsgText(5)
        Exit Sub
    Else
        adoacc310.MoveFirst
    End If
    'edit by nick 2004/08/20
    'intPage = 0
    intPage = 1
    intPage = intPage / 1
    'edit by nick 2004/08/20
    'strSalesNo = "" & adoacc310("A0K20").Value
    strSalesNo = "" & adoacc310("A3121").Value
    strSalesName = "" & adoacc310("收款人").Value
    PrintHead
    Do While adoacc310.EOF = False
        If intCounter >= 27 Then
            intCounter = 0
            Printer.NewPage
            intPage = intPage + 1
            PrintHead
        End If
        If strNo = "" Then
            strNo = "" & adoacc310("人工號").Value & adoacc310("電腦號").Value
            strCaseData = IIf(strOurCaseNo <> "" & adoacc310("本所案號").Value, ReConBNOurCaseNO("" & adoacc310("本所案號").Value), "") & adoacc310("案件性質名稱").Value
            If strOurCaseNo <> "" & adoacc310("本所案號").Value Then
                strOurCaseNo = "" & adoacc310("本所案號").Value
            End If
            strAmt = Val("" & adoacc310("A0J09").Value) + Val("" & adoacc310("A0J10").Value)
            'strPoint = Val("" & adoacc310("點數").Value)
            GoTo NextRec
        Else
            If strNo <> "" & adoacc310("人工號").Value & adoacc310("電腦號").Value Then
                adoacc310.MovePrevious
                GoTo PrintRec
            Else
                strCaseData = strCaseData & "及" & IIf(strOurCaseNo <> "" & adoacc310("本所案號").Value, ReConBNOurCaseNO("" & adoacc310("本所案號").Value), "") & adoacc310("案件性質名稱").Value
                If strOurCaseNo <> "" & adoacc310("本所案號").Value Then
                    strOurCaseNo = "" & adoacc310("本所案號").Value
                End If
                strAmt = Val(strAmt) + Val("" & adoacc310("A0J09").Value) + Val("" & adoacc310("A0J10").Value)
                'strPoint = Val(strPoint) + Val("" & adoacc310("點數").Value)
                GoTo NextRec
            End If
        End If
PrintRec:
        'edit by nick 2004/08/20
        'If strSalesNo <> "" & adoacc310.Fields("A0K20").Value Then
        If strSalesNo <> Trim("" & adoacc310.Fields("A3121").Value) Then
            Printer.CurrentX = PLeft(4)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print "現金：" & Format(dblCash, "#,##0")
            Printer.CurrentX = PLeft(6)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print "支票：" & Format(dblCheck, "#,##0")
            Printer.CurrentX = PLeft(9)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print "合計：" & Format(dblTOT, "#,##0")
            'add by nick 2004/08/20
            Printer.CurrentX = PLeft(12)
            Printer.CurrentY = 2000 + intCounter * 300
            'edit by nick 2004/10/20
            'Printer.Print "點數：" & Format(dblPoint, "#,##0")
            Printer.Print "點數：" & Format(dblPoint, "#,##0.000")
            'add by nick 2004/10/20
            Printer.CurrentX = PLeft(14)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print "留" & IIf(pub_strUserOffice = "1", "北所", IIf(pub_strUserOffice = "2", "中所", IIf(pub_strUserOffice = "3", "南所", IIf(pub_strUserOffice = "4", "高所", "其他所")))) & "：" & Format(dblMoney, "#,##0")
            'add by nick 2005/01/14  加入控制列不列印明細
            If Text3.Text = "N" Then
                Printer.CurrentX = PLeft(17)
                Printer.CurrentY = 2000 + intCounter * 300
                Printer.Print strSalesName
            End If
            intCounter = intCounter + 1
            dblCash = 0: dblCheck = 0: dblTOT = 0
            'add by nick 2004/08/20
            dblPoint = 0
            'add by nick 2004/10/20
            dblMoney = 0
            'edit by nick 2004/08/20
            'strSalesNo = "" & adoacc310.Fields("A0K20").Value
            strSalesNo = "" & adoacc310.Fields("A3121").Value
            strSalesName = "" & adoacc310.Fields("收款人").Value
            'add by nick 2005/01/14  加入控制列不列印明細
            If Text3.Text = "Y" Then
                intCounter = 0
                Printer.NewPage
                intPage = intPage + 1
                PrintHead
            'add by nick 2005/01/14  加入控制列不列印明細
            End If
        End If
        'add by nick 2005/01/14  加入控制列不列印明細
        If Text3.Text = "Y" Then
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print ChangeTStringToTDateString("" & adoacc310("收款日").Value)
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print "" & adoacc310("收款人")
            Printer.CurrentX = PLeft(2) + 6 * 90 - Printer.TextWidth("現金")
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print "" & adoacc310("收據抬頭").Value
            Printer.CurrentX = PLeft(3)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print strCaseData & IIf(strAmt = "0", "", strAmt)
            intCounter = intCounter + 1
                    
            Printer.CurrentX = PLeft(4)
            Printer.CurrentY = 2000 + intCounter * 300
    '        Printer.Print IIf(strCaseData = "", "" & adoacc310("人工號").Value, "")
            Printer.Print "" & adoacc310("人工號").Value
            Printer.CurrentX = PLeft(5)
            Printer.CurrentY = 2000 + intCounter * 300
    '        Printer.Print IIf(strCaseData <> "", "" & adoacc310("電腦號").Value, "")
            Printer.Print "" & adoacc310("電腦號").Value
            'edit by nick 2004/10/08
            'Printer.CurrentX = PLeft(6) + 6 * 90 - Printer.TextWidth("" & adoacc310("現金").Value)
            Printer.CurrentX = PLeft(6) + 6 * 90 - Printer.TextWidth(Format("" & adoacc310("現金").Value, "#,##0"))
            Printer.CurrentY = 2000 + intCounter * 300
            'edit by nick 2004/10/08
            'Printer.Print "" & adoacc310("現金").Value
            Printer.Print Format("" & adoacc310("現金").Value, "#,##0")
            'Printer.CurrentX = PLeft(7) + 6 * 90 - Printer.TextWidth("" & adoacc310("支票").Value)
            Printer.CurrentX = PLeft(7) + 6 * 90 - Printer.TextWidth(Format("" & adoacc310("支票").Value, "#,##0"))
            Printer.CurrentY = 2000 + intCounter * 300
            'edit by nick 2004/10/08
            'Printer.Print "" & adoacc310("支票").Value
            Printer.Print Format("" & adoacc310("支票").Value, "#,##0")
        'add by nick 2005/01/14  加入控制列不列印明細
        End If
        dblCash = dblCash + Val("" & adoacc310("現金").Value)
        dblCheck = dblCheck + Val("" & adoacc310("支票").Value)
        dblTOT = dblTOT + Val("" & adoacc310("現金").Value) + Val("" & adoacc310("支票").Value)
        'add by nick 2004/08/20
        dblPoint = dblPoint + Val("" & adoacc310("點數").Value)
        dblTCash = dblTCash + Val("" & adoacc310("現金").Value)
        dblTCheck = dblTCheck + Val("" & adoacc310("支票").Value)
        dblTTOT = dblTTOT + Val("" & adoacc310("現金").Value) + Val("" & adoacc310("支票").Value)
        'add by nick 2004/08/20
        dblTPoint = dblTPoint + Val("" & adoacc310("點數").Value)
        'add by nick 2004/10/20
        dblTMoney = dblTMoney + Val("" & adoacc310("留分所金額").Value)
        'add by nick 2005/01/14  加入控制列不列印明細
        If Text3.Text = "Y" Then
            Printer.CurrentX = PLeft(8)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print ChangeTStringToTDateString("" & adoacc310("到期日").Value)
            Printer.CurrentX = PLeft(9)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print "" & adoacc310("帳號").Value
            Printer.CurrentX = PLeft(10)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print "" & adoacc310("票號").Value
            Printer.CurrentX = PLeft(11)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print "" & adoacc310("付款地").Value
            Printer.CurrentX = PLeft(12)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print ChangeTStringToTDateString("" & adoacc310("扣繳日").Value)
            'edit by nick 2004/10/08
            'Printer.CurrentX = PLeft(13) + 6 * 90 - Printer.TextWidth("" & adoacc310("扣繳額").Value)
            Printer.CurrentX = PLeft(13) + 6 * 90 - Printer.TextWidth(Format("" & adoacc310("扣繳額").Value, "#,##0"))
            Printer.CurrentY = 2000 + intCounter * 300
            'edit by nick 2004/10/08
            'Printer.Print "" & adoacc310("扣繳額").Value
            Printer.Print Format("" & adoacc310("扣繳額").Value, "#,##0")
            'Printer.CurrentX = PLeft(14) + 6 * 90 - Printer.TextWidth("" & adoacc310("留分所金額").Value)
            Printer.CurrentX = PLeft(14) + 6 * 90 - Printer.TextWidth(Format("" & adoacc310("留分所金額").Value, "#,##0"))
            Printer.CurrentY = 2000 + intCounter * 300
            'edit by nick 2004/10/08
            'Printer.Print "" & adoacc310("留分所金額").Value
            Printer.Print Format("" & adoacc310("留分所金額").Value, "#,##0")
            Printer.CurrentX = PLeft(15) + 6 * 90 - Printer.TextWidth(Format(strPoint, "0.000"))
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print Format("" & adoacc310("點數").Value, "0.000")
            'add by nick 2004/08/26
            Printer.CurrentX = PLeft(17)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print StrToStr("" & adoacc310("備註").Value, 13)
            intCounter = intCounter + 1
                    
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print String(300, "-")
            intCounter = intCounter + 1
        'add by nick 2005/01/14  加入控制列不列印明細
        End If
        adoacc310.MoveNext
        If adoacc310.EOF = False Then
            strNo = "" & adoacc310("人工號").Value & adoacc310("電腦號").Value
            strCaseData = IIf(strOurCaseNo <> "" & adoacc310("本所案號").Value, ReConBNOurCaseNO("" & adoacc310("本所案號").Value), "") & adoacc310("案件性質名稱").Value
            If strOurCaseNo <> "" & adoacc310("本所案號").Value Then
                strOurCaseNo = "" & adoacc310("本所案號").Value
            End If
            strAmt = Val("" & adoacc310("A0J09").Value) + Val("" & adoacc310("A0J10").Value)
            'strPoint = Val("" & adoacc310("點數").Value)
            adoacc310.MoveNext
        End If
        GoTo NextRec1
NextRec:
        adoacc310.MoveNext
NextRec1:
    Loop
    If intCounter >= 27 Then
        intCounter = 0
        Printer.NewPage
        intPage = intPage + 1
        PrintHead
    End If
    adoacc310.MoveLast
        'edit by nick 2004/08/20
        'If strSalesNo <> "" & adoacc310.Fields("A0K20").Value Then
        If strSalesNo <> Trim("" & adoacc310.Fields("A3121").Value) Then
            Printer.CurrentX = PLeft(4)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print "現金：" & Format(dblCash, "#,##0")
            Printer.CurrentX = PLeft(6)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print "支票：" & Format(dblCheck, "#,##0")
            Printer.CurrentX = PLeft(9)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print "合計：" & Format(dblTOT, "#,##0")
            'add by nick 2004/08/20
            Printer.CurrentX = PLeft(12)
            Printer.CurrentY = 2000 + intCounter * 300
            'edit by nick 2004/10/20
            'Printer.Print "點數：" & Format(dblPoint, "#,##0")
            Printer.Print "點數：" & Format(dblPoint, "#,##0.000")
            'add by nick 2004/10/20
            Printer.CurrentX = PLeft(14)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print "留" & IIf(pub_strUserOffice = "1", "北所", IIf(pub_strUserOffice = "2", "中所", IIf(pub_strUserOffice = "3", "南所", IIf(pub_strUserOffice = "4", "高所", "其他所")))) & "：" & Format(dblMoney, "#,##0")
            'add by nick 2005/01/14  加入控制列不列印明細
            If Text3.Text = "N" Then
                Printer.CurrentX = PLeft(17)
                Printer.CurrentY = 2000 + intCounter * 300
                Printer.Print strSalesName
            End If
            intCounter = intCounter + 1
            dblCash = 0: dblCheck = 0: dblTOT = 0
            'add by nick 2004/08/20
            dblPoint = 0
            'add by nick 2004/10/20
            dblMoney = 0
            'edit by nick 2004/08/20
            'strSalesNo = "" & adoacc310.Fields("A0K20").Value
            strSalesNo = "" & adoacc310.Fields("A3121").Value
            strSalesName = "" & adoacc310.Fields("收款人").Value
            'add by nick 2005/01/14  加入控制列不列印明細
            If Text3.Text = "Y" Then
                intCounter = 0
                Printer.NewPage
                intPage = intPage + 1
                PrintHead
            'add by nick 2005/01/14  加入控制列不列印明細
            End If
        End If
        'add by nick 2005/01/14  加入控制列不列印明細
        If Text3.Text = "Y" Then
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print ChangeTStringToTDateString("" & adoacc310("收款日").Value)
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print "" & adoacc310("收款人")
            Printer.CurrentX = PLeft(2) + 6 * 90 - Printer.TextWidth("現金")
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print "" & adoacc310("收據抬頭").Value
            Printer.CurrentX = PLeft(3)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print strCaseData & IIf(strAmt = "0", "", strAmt)
            intCounter = intCounter + 1
                    
            Printer.CurrentX = PLeft(4)
            Printer.CurrentY = 2000 + intCounter * 300
    '        Printer.Print IIf(strCaseData = "", "" & adoacc310("人工號").Value, "")
            Printer.Print "" & adoacc310("人工號").Value
            Printer.CurrentX = PLeft(5)
            Printer.CurrentY = 2000 + intCounter * 300
    '        Printer.Print IIf(strCaseData <> "", "" & adoacc310("電腦號").Value, "")
            Printer.Print "" & adoacc310("電腦號").Value
            'edit by nick 2004/10/08
            'Printer.CurrentX = PLeft(6) + 6 * 90 - Printer.TextWidth("" & adoacc310("現金").Value)
            Printer.CurrentX = PLeft(6) + 6 * 90 - Printer.TextWidth(Format("" & adoacc310("現金").Value, "#,##0"))
            Printer.CurrentY = 2000 + intCounter * 300
            'edit by nick 2004/10/08
            'Printer.Print "" & adoacc310("現金").Value
            Printer.Print Format("" & adoacc310("現金").Value, "#,##0")
            'Printer.CurrentX = PLeft(7) + 6 * 90 - Printer.TextWidth("" & adoacc310("支票").Value)
            Printer.CurrentX = PLeft(7) + 6 * 90 - Printer.TextWidth(Format("" & adoacc310("支票").Value, "#,##0"))
            Printer.CurrentY = 2000 + intCounter * 300
            'edit by nick 2004/10/08
            'Printer.Print "" & adoacc310("支票").Value
            Printer.Print Format("" & adoacc310("支票").Value, "#,##0")
        'add by nick 2005/01/14  加入控制列不列印明細
        End If
        dblCash = dblCash + Val("" & adoacc310("現金").Value)
        dblCheck = dblCheck + Val("" & adoacc310("支票").Value)
        dblTOT = dblTOT + Val("" & adoacc310("現金").Value) + Val("" & adoacc310("支票").Value)
        'add by nick 2004/08/20
        dblPoint = dblPoint + Val("" & adoacc310("點數").Value)
        dblTCash = dblTCash + Val("" & adoacc310("現金").Value)
        dblTCheck = dblTCheck + Val("" & adoacc310("支票").Value)
        dblTTOT = dblTTOT + Val("" & adoacc310("現金").Value) + Val("" & adoacc310("支票").Value)
        'add by nick 2004/08/20
        dblTPoint = dblTPoint + Val("" & adoacc310("點數").Value)
        'add by nick 2004/10/20
        dblMoney = dblMoney + Val(adoacc310("留分所金額").Value)
        dblTMoney = dblTMoney + Val(adoacc310("留分所金額").Value)
        'add by nick 2005/01/14  加入控制列不列印明細
        If Text3.Text = "Y" Then
            Printer.CurrentX = PLeft(8)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print ChangeTStringToTDateString("" & adoacc310("到期日").Value)
            Printer.CurrentX = PLeft(9)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print "" & adoacc310("帳號").Value
            Printer.CurrentX = PLeft(10)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print "" & adoacc310("票號").Value
            Printer.CurrentX = PLeft(11)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print "" & adoacc310("付款地").Value
            Printer.CurrentX = PLeft(12)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print ChangeTStringToTDateString("" & adoacc310("扣繳日").Value)
            'edit by nick 2004/10/08
            'Printer.CurrentX = PLeft(13) + 6 * 90 - Printer.TextWidth("" & adoacc310("扣繳額").Value)
            Printer.CurrentX = PLeft(13) + 6 * 90 - Printer.TextWidth(Format("" & adoacc310("扣繳額").Value, "#,##0"))
            Printer.CurrentY = 2000 + intCounter * 300
            'edit by nick 2004/10/08
            'Printer.Print "" & adoacc310("扣繳額").Value
            Printer.Print Format("" & adoacc310("扣繳額").Value, "#,##0")
            'Printer.CurrentX = PLeft(14) + 6 * 90 - Printer.TextWidth("" & adoacc310("留分所金額").Value)
            Printer.CurrentX = PLeft(14) + 6 * 90 - Printer.TextWidth(Format("" & adoacc310("留分所金額").Value, "#,##0"))
            Printer.CurrentY = 2000 + intCounter * 300
            'edit by nick 2004/10/08
            'Printer.Print "" & adoacc310("留分所金額").Value
            Printer.Print Format("" & adoacc310("留分所金額").Value, "#,##0")
            Printer.CurrentX = PLeft(15) + 6 * 90 - Printer.TextWidth(Format(strPoint, "0.000"))
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print Format("" & adoacc310("點數").Value, "0.000")
            'add by nick 2004/08/26
            Printer.CurrentX = PLeft(17)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print StrToStr("" & adoacc310("備註").Value, 13)
            intCounter = intCounter + 1
                    
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print String(300, "-")
            intCounter = intCounter + 1
        'add by nick 2005/01/14  加入控制列不列印明細
        End If
        Printer.CurrentX = PLeft(4)
        Printer.CurrentY = 2000 + intCounter * 300
        Printer.Print "現金：" & Format(dblCash, "#,##0")
        Printer.CurrentX = PLeft(6)
        Printer.CurrentY = 2000 + intCounter * 300
        Printer.Print "支票：" & Format(dblCheck, "#,##0")
        Printer.CurrentX = PLeft(9)
        Printer.CurrentY = 2000 + intCounter * 300
        Printer.Print "合計：" & Format(dblTOT, "#,##0")
        'add by nick 2004/08/20
        Printer.CurrentX = PLeft(12)
        Printer.CurrentY = 2000 + intCounter * 300
        'edit by nick 2004/10/20
        'Printer.Print "點數：" & Format(dblPoint, "#,##0")
        Printer.Print "點數：" & Format(dblPoint, "#,##0.000")
        'add by nick 2004/10/20
        Printer.CurrentX = PLeft(14)
        Printer.CurrentY = 2000 + intCounter * 300
        Printer.Print "留" & IIf(pub_strUserOffice = "1", "北所", IIf(pub_strUserOffice = "2", "中所", IIf(pub_strUserOffice = "3", "南所", IIf(pub_strUserOffice = "4", "高所", "其他所")))) & "：" & Format(dblMoney, "#,##0")
            'add by nick 2005/01/14  加入控制列不列印明細
            If Text3.Text = "N" Then
                Printer.CurrentX = PLeft(17)
                Printer.CurrentY = 2000 + intCounter * 300
                Printer.Print strSalesName
            End If
        intCounter = intCounter + 1
        'add by nick 2005/01/14  加入控制列不列印明細
        If Text3.Text = "N" Then
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = 2000 + intCounter * 300
            Printer.Print String(300, "-")
            intCounter = intCounter + 1
        'add by nick 2005/01/14  加入控制列不列印明細
        End If
        dblCash = 0: dblCheck = 0: dblTOT = 0
        'add by nick 2004/08/20
        dblPoint = 0
        'add by nick 2004/10/20
        dblMoney = 0
        Printer.CurrentX = PLeft(4)
        Printer.CurrentY = 2000 + intCounter * 300
        Printer.Print "現金總計：" & Format(dblTCash, "#,##0")
        Printer.CurrentX = PLeft(6)
        Printer.CurrentY = 2000 + intCounter * 300
        Printer.Print "支票總計：" & Format(dblTCheck, "#,##0")
        Printer.CurrentX = PLeft(9)
        Printer.CurrentY = 2000 + intCounter * 300
        Printer.Print "總計：" & Format(dblTTOT, "#,##0")
        'add by nick 2004/08/20
        Printer.CurrentX = PLeft(12)
        Printer.CurrentY = 2000 + intCounter * 300
        'edit by nick 2004/10/20
        'Printer.Print "點數總計：" & Format(dblTPoint, "#,##0")
        Printer.Print "點數總計：" & Format(dblTPoint, "#,##0.000")
        'add by nick 2004/10/20
        Printer.CurrentX = PLeft(14)
        Printer.CurrentY = 2000 + intCounter * 300
        Printer.Print "留" & IIf(pub_strUserOffice = "1", "北所", IIf(pub_strUserOffice = "2", "中所", IIf(pub_strUserOffice = "3", "南所", IIf(pub_strUserOffice = "4", "高所", "其他所")))) & "：" & Format(dblTMoney, "#,##0")
        
        intCounter = intCounter + 1
        dblTCash = 0: dblTCheck = 0: dblTTOT = 0
        'add by nick 2004/08/20
        dblTPoint = 0
        'add by nick 2004/10/20
        dblTMoney = 0
    If adoacc310.State <> adStateClosed Then adoacc310.Close
    Set adoacc310 = Nothing
    Printer.EndDoc
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
    If Me.Option1(0).Value = True Then
        If Text1 <> MsgText(601) Then
           FormCheck = True
           Exit Function
        End If
    End If
    If Me.Option1(1).Value = True Then
        If MaskEdBox1.Text <> MsgText(29) Then
           FormCheck = True
           Exit Function
        End If
        If MaskEdBox2.Text <> MsgText(29) Then
           FormCheck = True
           Exit Function
        End If
    End If
    FormCheck = False
End Function

Private Sub Text2_GotFocus()
    TextInverse Me.Text2
End Sub

Private Function ReConBNOurCaseNO(strCaseNo As String) As String

If strCaseNo <> "" Then
    ReConBNOurCaseNO = Replace(Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "-" & Right(Left(strCaseNo, Len(strCaseNo) - 3), 6) & "-" & Right(Left(strCaseNo, Len(strCaseNo) - 2), 1) & "-" & Right(strCaseNo, 2), "-0-00", "")
Else
    ReConBNOurCaseNO = ""
End If

End Function

Private Sub GetPleft()

Erase PLeft
'PLeft(0) = 500
'PLeft(1) = 1940
'PLeft(2) = 3380
'PLeft(3) = 8180
'
'PLeft(4) = 500
'PLeft(5) = 1940
'PLeft(6) = 3380
'PLeft(7) = 4220
'PLeft(8) = 5060
'PLeft(9) = 6260
'PLeft(10) = 8180
'PLeft(11) = 10100
'PLeft(12) = 11660
'PLeft(13) = 12860
'PLeft(14) = 13700
'PLeft(15) = 14540
'PLeft(16) = 15380
'PLeft(17) = 17300


PLeft(0) = 500 * (3 / 4)
PLeft(1) = 1940 * (3 / 4)
PLeft(2) = 3380 * (3 / 4)
PLeft(3) = 8180 * (3 / 4) + (90 * 5)

PLeft(4) = 500 * (3 / 4)
PLeft(5) = 1940 * (3 / 4)
PLeft(6) = 3380 * (3 / 4)
PLeft(7) = 4220 * (3 / 4) + (90 * 2)
PLeft(8) = 5060 * (3 / 4) + (90 * 5)
PLeft(9) = 6260 * (3 / 4) + (90 * 5)
PLeft(10) = 8180 * (3 / 4) + (90 * 5)
PLeft(11) = 10100 * (3 / 4) + (90 * 5)
PLeft(12) = 11660 * (3 / 4) + (90 * 7)
PLeft(13) = 12860 * (3 / 4) + (90 * 7)
PLeft(14) = 13700 * (3 / 4) + (90 * 7)
PLeft(15) = 14540 * (3 / 4) + (90 * 9)
PLeft(16) = 15380 * (3 / 4) + (90 * 11)
PLeft(17) = 17300 * (3 / 4) + (90 * 13)

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    If Me.Text2.Text = "" Then Me.lblSalesName.Caption = "": Exit Sub
    Me.lblSalesName.Caption = GetStaffName(Me.Text2.Text)
    If Me.lblSalesName.Caption = "" Then
        MsgBox "智權人員輸入錯誤!!!", vbExclamation + vbOKOnly
        Cancel = True
    End If
    If Cancel = True Then Text2_GotFocus
End Sub

Private Sub Text3_GotFocus()
    TextInverse Me.Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
If KeyAscii <> 78 And KeyAscii <> 89 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
    If Text3.Text = "" Then
        MsgBox "是否列印明細，請輸入 Y 或 N！", , "錯誤！"
        Text3.SetFocus
        Cancel = True
    End If
End Sub
