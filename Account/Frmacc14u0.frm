VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc14u0 
   AutoRedraw      =   -1  'True
   Caption         =   "應收帳款財務處控管資料表"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3765
   ScaleWidth      =   5085
   Begin VB.TextBox Text4 
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
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   21
      Text            =   "Y"
      Top             =   2520
      Width           =   690
   End
   Begin VB.TextBox TxtClass 
      Height          =   264
      Left            =   120
      TabIndex        =   20
      Top             =   840
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.ComboBox cboClass 
      Height          =   300
      ItemData        =   "Frmacc14u0.frx":0000
      Left            =   1080
      List            =   "Frmacc14u0.frx":0002
      TabIndex        =   2
      Text            =   "cboClass"
      Top             =   1320
      Width           =   3000
   End
   Begin VB.CommandButton cmdRemClass 
      Caption         =   "移除↓"
      Height          =   285
      Left            =   4140
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1050
      Width           =   735
   End
   Begin VB.CommandButton cmdAddClass 
      Caption         =   "新增↑"
      Height          =   285
      Left            =   4140
      TabIndex        =   18
      Top             =   1350
      Width           =   735
   End
   Begin VB.ListBox lstClass 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      ItemData        =   "Frmacc14u0.frx":0004
      Left            =   1080
      List            =   "Frmacc14u0.frx":000B
      MultiSelect     =   1  '簡易多重選取
      Sorted          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   600
      Width           =   3000
   End
   Begin VB.ComboBox CboPrinter 
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
      Left            =   1140
      Style           =   2  '單純下拉式
      TabIndex        =   15
      Top             =   3360
      Width           =   3450
   End
   Begin VB.CommandButton Cmd_Excel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Excel"
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
      Left            =   2640
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   2900
      Width           =   2296
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
      Left            =   3000
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
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
      Height          =   315
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
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
      Height          =   315
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Cmd_Print 
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
      Left            =   120
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   2900
      Width           =   2296
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Top             =   120
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
      Left            =   3000
      TabIndex        =   1
      Top             =   120
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
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "是否含未列印收據            ( Y:含 )"
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
      Left            =   120
      TabIndex        =   22
      Top             =   2580
      Width           =   4605
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "印表機"
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
      Left            =   120
      TabIndex        =   16
      Top             =   3390
      Width           =   975
   End
   Begin VB.Label Lbl2 
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
      Left            =   2760
      TabIndex        =   14
      Top             =   1800
      Width           =   255
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
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblSalesName 
      BackStyle       =   0  '透明
      Caption         =   "智權人員名稱"
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
      Left            =   2760
      TabIndex        =   12
      Top             =   2160
      Width           =   1350
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "業務區"
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
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "控管類別"
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
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "控管日期"
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
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label LblS1 
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
      Left            =   2760
      TabIndex        =   8
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "Frmacc14u0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Amy 2015/04/07
Option Explicit

Public adoAcc14u0 As New ADODB.Recordset
Dim strPrinter As String, PLeft() As Integer, ColName() As String
Dim intCounter As Integer, intPage As Integer
Dim ii As Integer, strTemp As String, strCaseNo(1 To 4) As String
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
Dim strField, intWidth()
Dim intField As Integer
Dim strFileN As String 'Add by Amy 2015/05/14
Dim StrCP10Name As String 'Add by Amy 2015/07/03

Private Sub CboClass_Validate(Cancel As Boolean)
    Select Case cboClass
        'Modify by Amy 2015/08/24 +催款中/預計收款
        'Modify by Amy 2017/08/29 待收款 改為 請款中
        Case "", "全部", "待銷帳", "請款中", "未送件", "依流程請款", "其他", "會稿中", "尚未辦理", "催款中", "預計收款"
        Case Else
            ShowMsg Label2 & "錯誤, 請以下拉方式點選 !"
            Cancel = True
    End Select
End Sub

Private Sub Cmd_Excel_Click()
    If FormCheck = False Then
        MsgBox "請輸入查詢條件", , MsgText(5)
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    If ProduceData = True Then
        PrintExcel
    End If
    Screen.MousePointer = vbDefault
    FormClear
    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

Private Sub Cmd_Print_Click()
    If FormCheck = False Then
        MsgBox "請輸入查詢條件", , MsgText(5)
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    'Modify by Amy 2016/08/23 印表機設定修改
    PUB_RestorePrinter CboPrinter
    If ProduceData = True Then
        PrintReportA4
    End If
    PUB_RestorePrinter strPrinter
    'end 2016/08/23
    Screen.MousePointer = vbDefault
    FormClear
    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

Private Sub cmdAddClass_Click()
    If AddList(lstClass, cboClass) = True Then
        TxtClass = ComposeList(lstClass)
        cboClass = ""
    End If
    cboClass.SetFocus
End Sub

Private Function ComposeList(oList As ListBox, Optional p_iOpt As Integer = 0) As String
Dim iPos As Integer, stItem As String, strTemp As String
   
   strTemp = ""
   If oList.ListCount > 0 Then
      For intI = 0 To oList.ListCount - 1
         If p_iOpt = 0 Then
            iPos = InStr(oList.List(intI), Chr(1))
            If iPos > 0 Then
               stItem = Left(oList.List(intI), iPos - 1)
            Else
               stItem = oList.List(intI)
            End If
         Else
            stItem = Format(oList.ItemData(intI), "00")
         End If
         If intI = 0 Then
            strTemp = stItem
         Else
            strTemp = strTemp & "," & stItem
         End If
      Next
   End If
   ComposeList = strTemp
End Function

Private Sub cmdRemClass_Click()
    If RemoveList(lstClass) = True Then
        TxtClass = ComposeList(lstClass)
        cboClass.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim intX As Integer, intY As Integer
    Dim sglWidth As Single, sglHeight As Single

    Me.Icon = LoadPicture(strIcoPath)
    strFormName = Name
    Me.Width = 5205
    Me.Height = 4275
    Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
    Image1 = LoadPicture(strBackPicPath4)
    sglWidth = Image1.Width
    sglHeight = Image1.Height
    For intX = 0 To Int(ScaleWidth / sglWidth)
        For intY = 0 To Int(ScaleHeight / sglHeight)
            PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
        Next
    Next
    SetCombo cboClass
    lstClass.Clear
    MaskEdBox1.Mask = DFormat
    MaskEdBox2.Mask = DFormat
    lblSalesName.Caption = ""
    PUB_SetPrinter Me.Name, CboPrinter, strPrinter
    'Mark by Amy 2017/11/16 不預設,第一次進去選完會存
    'Add by Amy 2015/07/22 +印表機預設成1200-瑞婷
'    If Pub_StrUserSt03 = "M31" Then
'        CboPrinter = "HP LaserJet 1200 Series PCL"
'    End If
End Sub

'*************************************************
' 設定控管類別下拉選單
'
'*************************************************
Private Sub SetCombo(oCombo As ComboBox)
   With oCombo
      .Clear
      'Modify by Amy 2016/08/23 修改顯示順序-瑞婷
      .AddItem "全部"
      .AddItem "預計收款" 'Add by Amy 2015/08/24
      .AddItem "待銷帳"
      .AddItem "未送件"
      .AddItem "依流程請款"
      .AddItem "會稿中"
      .AddItem "尚未辦理"
      .AddItem "催款中" 'Add by Amy 2015/08/24
      .AddItem "請款中" 'Modify by Amy 2017/08/29 原:待收款
      .AddItem "其他"
      'end 2016/08/23
   End With
End Sub

Private Function AddList(oList As ListBox, oCombo As ComboBox, Optional p_iOpt As Integer = 0) As Boolean
    Dim idx As Integer, bFound As Boolean, stNewItem As String, iNewItemData As Integer
    Dim stSort As String, iPos As Integer
   
    If oCombo.Text = "" Then
        Exit Function
    End If
    'Add by Amy 2016/08/23 全部或重覆不需再加入
    If InStr(oList, "全部") > 0 Then
        Exit Function
    ElseIf InStr(oList, oCombo.Text) > 0 Then
        Exit Function
    End If
   
    '若有控制字元時後面為說明文字不抓
    iPos = InStr(oCombo, Chr(1))
    If iPos > 0 Then
        stNewItem = Left(oCombo, iPos - 1)
    Else
        stNewItem = oCombo
    End If
      
    If InStr(stNewItem, ",") > 0 Then
        MsgBox "逗號[,]為系統保留字，請改用其他符號！", vbExclamation
        oCombo.SetFocus
        Exit Function
    End If

    If stNewItem <> "" Then
        If bFound = False Then
            oList.AddItem stNewItem, 0
            If p_iOpt <> 0 Then
                oList.ItemData(0) = oCombo.ItemData(oCombo.ListIndex)
            End If
            AddList = True
        End If
    End If
End Function

Private Function RemoveList(oList As ListBox) As Boolean
Dim ii As Integer
   
   If oList.ListCount > 0 Then
      ii = 0
      Do While ii < oList.ListCount
         If oList.Selected(ii) = True Then
            RemoveList = True
            oList.RemoveItem ii
            ii = ii - 1
         End If
         ii = ii + 1
      Loop
   End If
End Function

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
    MaskEdBox1.SetFocus
    TxtClass = ""
    lstClass.Clear
    Text1 = ""
    Text2 = ""
    Text3 = ""
    lblSalesName.Caption = ""
End Sub
'
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

Private Sub Form_Unload(Cancel As Integer)
    strFormName = MsgText(601)
    KeyEnter vbKeyEscape
    MenuEnabled
    StatusClear
    If Me.CboPrinter.Text <> Me.CboPrinter.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.CboPrinter.Name, "0", "0", Me.CboPrinter.Text
    End If
    Set Frmacc14u0 = Nothing
End Sub

Private Sub Text1_GotFocus()
    TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
    TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    If Trim(Text2) <> "" Then
        If RunNick(Text1, Text2) = True Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub Text3_GotFocus()
    TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
    Cancel = False
    If Text3 <> "" And GetStaffName(Text3) = "" Then
        Cancel = True
        lblSalesName.Caption = ""
        MsgBox Label4.Caption & "不存在", vbCritical
        TextInverse Text3
        Exit Sub
    Else
        lblSalesName.Caption = GetStaffName(Text3)
    End If
End Sub

Private Function ProduceData() As Boolean
    Dim strQ As String, strClass As String, strTp() As String
    Dim ii As Integer
    Dim strWhere As String 'Add by Amy 2015/06/24
    
On Error GoTo Checking
    If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
        strQ = " And a0k38 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
    End If
    If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
        strQ = strQ & " And a0k38 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
    End If
    'Add by Amy 2016/03/04 +是否含未列印收據
    If Text4 = MsgText(601) Then strQ = strQ & " And A0K32 is null "
    If lstClass <> MsgText(601) Then
        If InStr(TxtClass, "全部") Then
            'Modify by Amy 2016/08/23 改顯示順序,寫成一致
            'Modify by Amy 2017/08/29 待收款 改為 請款中
            TxtClass = "預計收款,待銷帳,未送件,依流程請款,會稿中,尚未辦理,催款中,請款中,其他"
        End If
        
        strTp = Split(TxtClass, ",")
        For ii = 0 To UBound(strTp)
            strClass = strClass & " Or InStr(a0k39,'" & strTp(ii) & "')>0 "
        Next ii
        strQ = strQ & " And (" & Right(strClass, Len(strClass) - 3) & ") "
    End If
    If Text1 <> MsgText(601) Then
        strWhere = strWhere & " And st15 >= '" & Text1 & "' "
    End If
    If Text2 <> MsgText(601) Then
        strWhere = strWhere & " And st15 <= '" & Text2 & "' "
    End If
    If Text3 <> MsgText(601) Then
        strQ = strQ & " And a0k20='" & Text3 & "' "
    End If
    'Modify by Amy 2015/07/14 因acc1u0及acc1j0串出來的資料可能為多筆造成加總金額錯誤
'    strQ = "Select st02,sqldatet(a0k38) as a0k38,a0k39,a0k01,a0j02,Sum(Nvl(a0k06,0)+Nvl(a0k07,0))-Sum(Nvl(a1u07,0)+Nvl(a1u09,0)) as Amount,a0902,st15,st01 " & _
'                "From acc0k0,acc0j0,Staff,acc1u0,acc090 " & _
'                "Where (a0k09 is null or a0k09 = 0) And a0k01=a0j13(+) And a0k20=st01(+) And st15=a0901(+)" & _
'                "And a0j13=a1u02(+) And a1u03(+)=a0j01 " & strQ & _
'                " Group by a0902,st02,a0k38,a0k39,a0k01,a0j02,st15,st01" & _
'                " Order by st15,st01"
    'Modify by Amy 2015/07/21 +收據抬頭a0k04
    'Modify by Amy 2015/07/28 控管日期a0k38 欄改顯示收據日期a0k02
    'Modify by Amy 2016/06/22 排序+a0k39
    strQ = "Select st02,sqldatet(a0k02) as DocDate,a0k39,a0k01,cp01||cp02||cp03||cp04 CaseNo,Decode(nvl(Amt2,0)+nvl(Amt3,0),0, nvl(a0j09,0)+nvl(a0j10,0),nvl(a0j09,0)+nvl(a0j10,0)-nvl(Amt2,0)-nvl(Amt3,0)) Amount,a0j01,a0k08,a0902,st15,st01,cp10,a0k04 " & _
                "From acc0k0,Staff, acc090,CaseProgress," & _
                "(Select a1u03,a1u02, sum(nvl(a1u04,0)+nvl(a1u05,0)-nvl(a1u08,0)-nvl(a1u10,0)) Amt2,sum(nvl(a1u07,0)+nvl(a1u09,0)) Amt3 From acc1u0,acc0k0 Where (a0k09 is null or a0k09 = 0) And a0k01=a1u02 (+) And a1u02 is not null " & strQ & " Group by a1u02,a1u03) x, " & _
                "(Select a0j01,a0j13, sum(nvl(a0j09,0)) a0j09,sum(nvl(a0j10,0)) a0j10 From acc0j0,acc0k0 Where (a0k09 is null or a0k09 = 0) And a0k01=a0j13(+) And a0j13 is not null " & strQ & " Group by a0j01,a0j13) y " & _
                "Where (a0k09 is null or a0k09 = 0) And a0k01=a0j13(+) And a0j13=a1u02(+) And  a0j01=a1u03(+) And a0k20=st01(+) And st15=a0901(+) And a0j01=cp09(+) " & strQ & strWhere & _
                "And nvl(a0j09,0)+nvl(a0j10,0)- Nvl(Amt3,0)<> Nvl(Amt2,0) Order by a0k39,st15,st01,a0k01"
                
    If adoAcc14u0.State = adStateOpen Then adoAcc14u0.Close
    adoAcc14u0.CursorLocation = adUseClient
    adoAcc14u0.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
    If adoAcc14u0.RecordCount = 0 Then
        ProduceData = False
        adoAcc14u0.Close
        MsgBox MsgText(28), , MsgText(5)
        Exit Function
    Else
        ProduceData = True
    End If
    
Checking:
    If Err.Number = 0 Then
        Exit Function
    End If
    MsgBox Err.Description, , MsgText(5)
End Function

Private Sub PrintReportA4()
    Dim StaffNo As String
    Dim MaxRow As Integer 'Add by Amy 2015/07/03
    Dim strClass As String 'Add by Amy 2016/06/22
    
On Error GoTo Checking
    
    GetPleft
    MaxRow = 20 'Add by Amy 2015/07/03
    'Modify by Amy 2015/05/21 原:PUB_GetPaperSize(9)
    Printer.PaperSize = 9 '設定紙張 A4
    Printer.Orientation = 2 'Modify by Amy 2015/07/03 改橫印
    intPage = 1
    With adoAcc14u0
        Do While .EOF = False
            'Modify by Amy 2016/06/22 +strClass <> .Fields("a0k39")
            If StaffNo <> .Fields("st01") Or intCounter > MaxRow Or strClass <> .Fields("a0k39") Then
                intCounter = 1
                If intPage <> 1 Then Printer.NewPage
                PrintHeadA4 .Fields("a0k39")
                intPage = intPage + 1
                Printer.FontBold = False
            End If
            
            For ii = 1 To UBound(ColName)
                Select Case ii
                    Case 1 '日期
                        strTemp = "" & .Fields("DocDate") 'Modify by Amy 2015/07/28 原:控管日期
                        Printer.CurrentX = PLeft(ii - 1)
                    'Add by Amy 2015/07/21
                    Case 2 '客戶名稱
                        strTemp = StrToStr("" & .Fields("a0k04"), 6)
                        Printer.CurrentX = PLeft(ii - 1)
                    'Mark by Amy 2015/08/24 若選「預計收款」增加「預計//日收款」無法全部顯示
'                    Case 3 '控管類別
'                        strTemp = "" & .Fields("a0k39")
'                        Printer.CurrentX = PLeft(ii - 1)
                    Case 3 '收據號碼
                        strTemp = "" & .Fields("a0k01")
                        Printer.CurrentX = PLeft(ii - 1)
                    Case 4 '本所案號
                        strTemp = CaseNoAddSign("" & .Fields("CaseNo")) 'Modify by Amy 2015/07/14
                        Printer.CurrentX = PLeft(ii - 1)
                    Case 5 '案件名稱
                        strTemp = StrToStr(GetCaseName("" & .Fields("a0j01"), StrCP10Name), 6) 'Add by Amy 2015/07/03
                        Printer.CurrentX = PLeft(ii - 1)
                    'Add by Amy 2015/07/03
                    Case 6 '案件性質
                        strTemp = StrToStr(StrCP10Name, 6)
                        Printer.CurrentX = PLeft(ii - 1)
                    Case 7 '收據金額
                        strTemp = Format("" & .Fields("Amount"), DDollar)
                        Printer.CurrentX = PLeft(ii) - 100 - Printer.TextWidth(strTemp)
                    'Add by Amy 2015/07/03
                    Case 8 '收據備註
                        strTemp = StrToStr("" & .Fields("a0k08"), 23) 'Modify by Amy 2015/08/24 原:15
                        Printer.CurrentX = PLeft(ii - 1)
                    Case Else
                End Select
                
                Printer.CurrentY = 300 + intCounter * 300
                Printer.Print strTemp
            Next ii
            intCounter = intCounter + 1
            StaffNo = .Fields("st01")
            strClass = .Fields("a0k39") 'Add by Amy 2016/06/22
            .MoveNext
        Loop
    End With
    Printer.EndDoc

Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Modify by Amy +stClass
Private Sub PrintHeadA4(ByVal stClass As String)
    
    With adoAcc14u0
        strTemp = "應收帳款財務處控管資料表"
        Printer.FontSize = 16
        Printer.FontBold = True
        Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strTemp) / 2)
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print strTemp
        intCounter = intCounter + 2
    
        Printer.FontSize = 12
        'Add by Amy 2015/08/24 +若選「預計收款」增加「預計//日收款」造成備註欄位無法印出其資料
        '                                          瑞婷:每次只會選一個類別,故顯示於列印日期前
        'Modify by Amy 2016/06/22
        Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("控管類別：" & stClass) / 2)
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print "控管類別：" & stClass
        'end 2016/06/22
        intCounter = intCounter + 1
        'end 2015/08/24
        
        Printer.CurrentX = 12000 'Modify by Amy 2015/07/03 原:直印時8800
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print "列印日期：" & CFDate(ACDate(ServerDate))
        intCounter = intCounter + 1
       
        Printer.CurrentX = 200 'Modify by Amy 2015/07/03 原:直印時0
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print "智權人員："; .Fields("a0902") & " " & .Fields("st02")
        Printer.CurrentX = 12000 'Modify by Amy 2015/07/03 原:直印時8800
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print "列印人員：" & StaffQuery(strUserNum)
        intCounter = intCounter + 2
        
        For ii = 1 To UBound(ColName)
            Printer.CurrentX = PLeft(ii - 1) + (PLeft(ii) - PLeft(ii - 1) - Printer.TextWidth(ColName(ii)) - 100) / 2
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print ColName(ii)
            Printer.Line (PLeft(ii - 1), Printer.CurrentY)-(PLeft(ii) - 100, Printer.CurrentY)
        Next ii
        intCounter = intCounter + 1
    End With
End Sub

Private Sub GetPleft()
   'Modify by Amy 2015/07/03 增加欄位
   'Modify by Amy 2015/07/21 增加客戶名稱
   Dim ii As Integer
   ReDim PLeft(0 To 8)
   ReDim ColName(1 To 8)
   
   ii = 0
   PLeft(ii) = 200: ii = ii + 1
   PLeft(ii) = PLeft(ii - 1) + 1400: ColName(ii) = "收據日期": ii = ii + 1 'Modify by Amy 2015/07/28 原:控管日期
   PLeft(ii) = PLeft(ii - 1) + 1700: ColName(ii) = "客戶名稱": ii = ii + 1
   'PLeft(ii) = PLeft(ii - 1) + 1500: ColName(ii) = "控管類別": ii = ii + 1 'Mark by Amy 2015/08/24 若選「預計收款」增加「預計//日收款」後無法全部顯示
   PLeft(ii) = PLeft(ii - 1) + 1300: ColName(ii) = "收據號碼": ii = ii + 1
   PLeft(ii) = PLeft(ii - 1) + 2000: ColName(ii) = "本所案號": ii = ii + 1
   PLeft(ii) = PLeft(ii - 1) + 1700: ColName(ii) = "案件名稱": ii = ii + 1
   PLeft(ii) = PLeft(ii - 1) + 1700: ColName(ii) = "案件性質": ii = ii + 1 'Add by Amy 2015/07/03
   PLeft(ii) = PLeft(ii - 1) + 1400: ColName(ii) = "收據金額": ii = ii + 1
   PLeft(ii) = PLeft(ii - 1) + 5200: ColName(ii) = "收據備註": ii = ii + 1 'Modify by Amy 2015/08/24 原:3400
   
End Sub

Private Sub PrintExcel()
    Dim StaffNo As String, strFieldN As String, bolIsFirst As Boolean
    Dim strClass As String, intXlsSheet As Integer
    Dim strWkName As String 'Add by Amy 2017/09/25 for 2010 工作表名稱為中文
    
On Error GoTo Checking
    'Modfiy by Amy 2015/07/03 增加案件性質、收據金額 欄位
    'Modify by Amy 2015/07/21 增加客戶名稱(顯示收據抬頭a0k04)
    ReDim strField(8)
    ReDim intWidth(8)
    strField = Array("收據日期", "客戶名稱", "控管類別", "收據號碼", "本所案號", "案件名稱", _
                                 "案件性質", "收據金額", "收據備註")
    intWidth = Array(10, 13, 13, 13, 16, 13, 15, 13, 24)
    intField = 65: intCounter = 1: bolIsFirst = True: intXlsSheet = 1
    'end 2015/07/03
    
    'Add by Amy 2015/05/14 沒Save 執行第二次Excel 會當
    strFileN = "應收帳款財務控管資料表" & ACDate(ServerDate) & ServerTime & MsgText(43)
    If Dir(strExcelPath & strFileN) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & strFileN
    End If
    
    If bolIsFirst = True Then
        xlsAnnuity.SheetsInNewWorkbook = 3 'Added by Lydia 2019/03/13 預設工作表數量
        xlsAnnuity.Workbooks.add
        xlsAnnuity.Application.WindowState = xlMinimized
    End If
    
'Modify by Amy 2016/06/22 +依類別產生不同工作表
NextComp:
    If intXlsSheet > 3 Then
        xlsAnnuity.Worksheets.add
    End If
    'Modify by Amy 2017/09/25 for 工作表名稱改為中文
    If strWkName = MsgText(601) Then strWkName = Left(xlsAnnuity.Worksheets(1).Name, Len(xlsAnnuity.Worksheets(1).Name) - 1)
    Set wksAnnuity = xlsAnnuity.Worksheets(strWkName & intXlsSheet)
    'end 2017/09/25
    wksAnnuity.Activate
    With adoAcc14u0
        Call SetExcel(bolIsFirst)
        bolIsFirst = False
        Do While .EOF = False
            If strClass <> .Fields("a0k39") And strClass <> MsgText(601) Then
                '改工作表名稱
                wksAnnuity.Name = strClass
                 intXlsSheet = intXlsSheet + 1
                intCounter = 1: bolIsFirst = True: StaffNo = "": strClass = ""
                GoTo NextComp
            End If
            If StaffNo <> .Fields("st01") And StaffNo <> MsgText(601) Then
                '換頁
                wksAnnuity.Range("A" & intCounter).Select
                wksAnnuity.HPageBreaks.add Before:=wksAnnuity.Application.ActiveCell
                Call SetExcel
            End If
            For ii = 0 To UBound(strField)
                strFieldN = strField(ii)
                Select Case strField(ii)
                    'Modify by Amy 2015/07/28 改顯示收據日
                    Case "收據日期"
                        strTemp = "" & .Fields("DocDate")
                    'Add by Amy 2015/07/21
                    Case "客戶名稱"
                        strTemp = StrToStr("" & .Fields("a0k04"), 6)
                    Case "控管類別"
                        strTemp = "" & .Fields("a0k39")
                    Case "收據號碼"
                        strTemp = "" & .Fields("a0k01")
                    Case "本所案號"
                        strTemp = CaseNoAddSign("" & .Fields("CaseNo")) 'Modify by Amy 2015/07/16
                    Case "案件名稱"
                        strTemp = StrToStr(GetCaseName("" & .Fields("a0j01"), StrCP10Name), 6) 'Add by Amy 2015/07/03
                    Case "案件性質"
                        strTemp = StrCP10Name
                    Case "收據金額"
                        strTemp = Format("" & .Fields("Amount"), DDollar)
                        wksAnnuity.Range(Chr(GetValue(strFieldN) + intField) & intCounter).NumberFormatLocal = "#,##0"
                    Case "收據備註"
                        strTemp = "" & .Fields("a0k08")
                    Case Else
                End Select
                wksAnnuity.Range(Chr(GetValue(strFieldN) + 65) & intCounter).Value = strTemp
            Next ii
             intCounter = intCounter + 1
             StaffNo = .Fields("st01")
             strClass = .Fields("a0k39")
            .MoveNext
        Loop
    End With
    'Modify by Amy 2015/05/14
    'xlsAnnuity.Visible = True
    wksAnnuity.Name = strClass
    If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=-4143
   Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=56
   End If
    xlsAnnuity.Workbooks.Close
    xlsAnnuity.Quit
    Set xlsAnnuity = Nothing
    Set wksAnnuity = Nothing
    adoAcc14u0.Close
    MsgBox "Excel 檔已產生~"
    Exit Sub

Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=-4143
   Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=56
   End If
   'end 2016/06/22
   xlsAnnuity.Workbooks.Close
   xlsAnnuity.Quit
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
End Sub

Private Sub SetExcel(Optional IsFirst As Boolean = False)
    
    With wksAnnuity
        If IsFirst = True Then
            .Range(Chr(intField) & intCounter).Font.Size = 16
            .Range(Chr(intField) & intCounter).Font.Bold = True
            .Range(Chr(intField) & intCounter).Value = "應收帳財務處控管資料表"
            intCounter = intCounter + 1
            
            .PageSetup.PrintTitleRows = "$1:$1"
            .PageSetup.PaperSize = xlPaperA4    '設定紙張 A4
            .PageSetup.Orientation = xlLandscape  'Modify by Amy 2015/07/03 改橫印
            .PageSetup.LeftMargin = 28.34
            .PageSetup.RightMargin = 28.34
            .PageSetup.TopMargin = 42.51
            .PageSetup.BottomMargin = 42.51
            .PageSetup.HeaderMargin = 28.34
            .PageSetup.FooterMargin = 28.34
        End If
        .Range(Chr(intField) & intCounter).Value = "列印人員：" & StaffQuery(strUserNum)
        .Range(Chr(intField + UBound(strField) - 1) & intCounter).Value = "列印日期："
        .Range(Chr(intField + UBound(strField) - 1) & intCounter).HorizontalAlignment = xlHAlignRight
        .Range(Chr(intField + UBound(strField)) & intCounter).Value = CFDate(ACDate(ServerDate))
        .Range(Chr(intField + UBound(strField)) & intCounter).HorizontalAlignment = xlHAlignLeft
        intCounter = intCounter + 1
        
        .Range(Chr(intField) & intCounter).Value = "智權人員：" & adoAcc14u0.Fields("a0902") & " " & adoAcc14u0.Fields("st02")
        intCounter = intCounter + 1
        
        For ii = 0 To UBound(strField)
            .Columns(Chr(intField + ii) & ":" & Chr(intField + ii)).ColumnWidth = intWidth(ii)
            .Range(Chr(intField + ii) & intCounter).Value = strField(ii)
            .Range(Chr(intField + ii) & intCounter).HorizontalAlignment = xlCenter
        Next ii
        .Range(Chr(intField) & intCounter & ":" & Chr(UBound(strField) + intField) & intCounter).Select
        With .Application.Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        intCounter = intCounter + 1
    End With
End Sub

'重組本所案號(加-符號)
Private Function CaseNoAddSign(ByVal stra0j02 As String) As String
    strCaseNo(1) = "": strCaseNo(2) = "": strCaseNo(3) = "": strCaseNo(4) = ""
    CaseNoAddSign = ""
    If stra0j02 = MsgText(601) Then Exit Function
    strCaseNo(1) = Mid(stra0j02, 1, Len(stra0j02) - 9)
    strCaseNo(2) = Mid(stra0j02, (Len(stra0j02) - 9) + 1, 6)
    strCaseNo(3) = Mid(stra0j02, (Len(stra0j02) - 3) + 1, 1)
    strCaseNo(4) = Mid(stra0j02, (Len(stra0j02) - 2) + 1, 2)
    CaseNoAddSign = strCaseNo(1) & "-" & strCaseNo(2) & "-" & strCaseNo(3) & "-" & strCaseNo(4)
End Function

'取得案件名稱 中->英->日
'Modify by Amy 2015/07/03 +傳入總收文號,回傳案件性質名稱
Private Function GetCaseName(stra0j01 As String, ByRef StrCP10Name As String) As String
    Dim adoBase As New ADODB.Recordset
    Dim strBase As String
    
    GetCaseName = "": StrCP10Name = ""
    If strCaseNo(1) = MsgText(601) Then Exit Function
    strBase = "Select Decode(pa05,null,Decode(pa06,null,Nvl(pa07,''),pa06),pa05) as CaseName,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) as CP10 From Patent,CaseProgress,CasePropertyMap " & _
                    "Where pa01=cP01(+) And pa02=cP02(+) And pa03=cP03(+) And pa04=cP04(+) And cp01=cpm01(+) And cp10=cpm02(+) And pa01='" & strCaseNo(1) & "' And pa02='" & strCaseNo(2) & "' And pa03='" & strCaseNo(3) & "' And pa04='" & strCaseNo(4) & "' And cp09='" & stra0j01 & "'" & _
         " Union Select tm05 as CaseName,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) as CP10 From TradeMark,CaseProgress,CasePropertyMap " & _
         "Where tm01=cp01(+) And tm02=cp02(+) And tm03=cp03(+) And tm04=cp04(+) And cp01=cpm01(+) And cp10=cpm02(+) And tm01='" & strCaseNo(1) & "' And tm02='" & strCaseNo(2) & "' And tm03='" & strCaseNo(3) & "' And tm04='" & strCaseNo(4) & "' And cp09='" & stra0j01 & "'" & _
         " Union Select Decode(lc05,null,Decode(lc06,null,Nvl(lc07,''),lc06),lc05) as CaseName,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) as CP10 From LawCase,CaseProgress,CasePropertyMap " & _
         "Where lc01=cp01(+) And lc02=cp02(+) And lc03=cp03(+) And lc04=cp04(+) And cp01=cpm01(+) And cp10=cpm02(+) And lc01='" & strCaseNo(1) & "' And lc02='" & strCaseNo(2) & "' And lc03='" & strCaseNo(3) & "' And lc04='" & strCaseNo(4) & "' And cp09='" & stra0j01 & "'" & _
         " Union Select hc06 as CaseName,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) as CP10 From HireCase,CaseProgress,CasePropertyMap " & _
         "Where hc01=cp01(+) And hc02=cp02(+) And hc03=cp03(+) And hc04=cp04(+) And cp01=cpm01(+) And cp10=cpm02(+) And hc01='" & strCaseNo(1) & "' And hc02='" & strCaseNo(2) & "' And hc03='" & strCaseNo(3) & "' And hc04='" & strCaseNo(4) & "' And cp09='" & stra0j01 & "'" & _
         " Union Select Decode(sp05,null,Decode(sp06,null,Nvl(sp07,''),sp06),sp05) as CaseName,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) as CP10 From ServicePractice,CaseProgress,CasePropertyMap " & _
         "Where sp01=cp01(+) And sp02=cp02(+) And sp03=cp03(+) And sp04=cp04(+) And cp01=cpm01(+) And cp10=cpm02(+) And sp01='" & strCaseNo(1) & "' And sp02='" & strCaseNo(2) & "' And sp03='" & strCaseNo(3) & "' And sp04='" & strCaseNo(4) & "' And cp09='" & stra0j01 & "' "
    adoBase.CursorLocation = adUseClient
    adoBase.Open strBase, adoTaie, adOpenStatic, adLockReadOnly
    If adoBase.RecordCount <> 0 Then
        GetCaseName = "" & adoBase.Fields("CaseName")
        StrCP10Name = "" & adoBase.Fields("CP10")
    End If
End Function

Private Function GetValue(pFieldN As String) As Integer
   Dim jj As Integer
 
    For jj = 1 To UBound(strField)
       If UCase(strField(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

'Add by Amy 2016/03/04
Private Sub Text4_GotFocus()
    TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    Select Case KeyAscii
    Case 89, 8
        'Do Nothing
    Case Else
        KeyAscii = 0
    End Select
End Sub
'end 2016/03/04


