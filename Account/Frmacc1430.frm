VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc1430 
   AutoRedraw      =   -1  'True
   Caption         =   "付款工作底稿"
   ClientHeight    =   3384
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3384
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
      Left            =   1440
      TabIndex        =   0
      Top             =   150
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
      Height          =   345
      Left            =   1050
      Style           =   2  '單純下拉式
      TabIndex        =   16
      Top             =   2880
      Width           =   3945
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
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   12
      Top             =   1890
      Width           =   435
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
      Left            =   225
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   2310
      Width           =   4692
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
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   550
      Width           =   612
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1440
      TabIndex        =   5
      Top             =   960
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
      Left            =   3360
      TabIndex        =   7
      Top             =   960
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
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   1440
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
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   300
      Left            =   3360
      TabIndex        =   11
      Top             =   1440
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
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
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
      TabIndex        =   17
      Top             =   2940
      Width           =   885
   End
   Begin VB.Label Label8 
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
      Left            =   240
      TabIndex        =   15
      Top             =   150
      Width           =   975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "翻譯費是否含未收回回執單         ( N: 不含 )"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   14
      Top             =   1920
      Width           =   4680
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   135
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
      Left            =   3120
      TabIndex        =   10
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "欲處理日期"
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
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
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
      Left            =   3120
      TabIndex        =   6
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label3 
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
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "(1.廠商/員工 2.客戶)"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   550
      Width           =   2415
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
      Left            =   240
      TabIndex        =   1
      Top             =   550
      Width           =   1215
   End
End
Attribute VB_Name = "Frmacc1430"
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
Public adoaccrpt103 As New ADODB.Recordset
Public adoacc1p0 As New ADODB.Recordset
Dim lngCounter As Long
Dim dllaccrpt103 As Object
Dim strStartDate As String
Dim strEndDate As String
Dim strStartDate1 As String
Dim strEndDate1 As String
'Added by Lydia 2017/07/31 改成A4橫印
Dim strPrinter As String '系統預設印表機
Dim mPrtOrt As Integer  '原本預設印表機的列印方向
Private Const ciTitleFontSize = 14
Private Const ciFontSize = 10
Private Const ciStartX = 400, ciStartY = 400, ciColGap = 150
Dim PLeft(0 To 7) As Integer '欄位起始位置陣列
Dim PTitle(0 To 6) As String '欄位抬頭陣列
Dim strTemp(0 To 6) As String
Dim iPrint As Integer, iPage As Integer
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim iPageLine As Integer '頁面資料列
Dim strTmpSubTot As String
Dim strTmpTot As String
'end 2017/07/31

'Add by Amy 2020/04/10
Private Sub Combo2_GotFocus()
    TextInverse Combo2
End Sub

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
        MsgBox Label8 & MsgText(63), , MsgText(5)
        Cancel = True
        Combo2.SetFocus
        Exit Sub
    ElseIf Len(Trim(Combo2)) = 1 Then
        Combo2 = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/04/10

Private Sub Command1_Click()
'Added by Lydia 2017/07/31
Dim strGrp As String, intP As Integer
'Add by Amy 2014/01/16 +公司別
Dim bCancel As Boolean

   bCancel = False
   'Modify by Amy 2020/04/10 公司別改下拉 原:Text3
   If Trim(Combo2) = MsgText(601) Then
      MsgBox "請輸入" & Label8, , MsgText(5)
      Combo2.SetFocus
      Exit Sub
   End If
   Call Combo2_Validate(bCancel)
   If bCancel = True Then
      Exit Sub
   End If
   'end 2020/04/10
   'end 2014/01/16
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt103Delete
   ProduceData
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strStartDate = MaskEdBox1.Text
      strEndDate = MaskEdBox2.Text
   Else
      strStartDate = ""
      strEndDate = ""
   End If
   If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> MsgText(29) Then
      strStartDate1 = MaskEdBox3.Text
      strEndDate1 = MaskEdBox4.Text
   Else
      strStartDate1 = ""
      strEndDate1 = ""
   End If
   If adoaccrpt103.State = adStateOpen Then
      adoaccrpt103.Close
   End If
   adoaccrpt103.CursorLocation = adUseClient
   'Modified by Lydia 2017/09/11 排序(R10301+R10310+R10303)
   'adoaccrpt103.Open "select * from accrpt103", adoTaie, adOpenStatic, adLockReadOnly
   adoaccrpt103.Open "select * from accrpt103 order by r10301,r10310,r10303", adoTaie, adOpenStatic, adLockReadOnly
   
   If adoaccrpt103.RecordCount <> 0 Then
'      Select Case Text1
'         Case Mid(ComboItem(91), 1, 1)
'            dllaccrpt103.Acc1430 ReportTitle(103), MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'         Case Mid(ComboItem(92), 1, 1)
'            dllaccrpt103.Acc1430 ReportTitle(103), MaskEdBox3.Text, MaskEdBox4.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'         Case Mid(ComboItem(93), 1, 1), "4"
'            dllaccrpt103.Acc1430 ReportTitle(103), MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'      End Select
'      dllaccrpt103.Acc1430 ReportTitle(103), strStartDate, strEndDate, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
     'Modify by Amy 2014/01/16 +公司別名稱
      'Modified by Lydia 2017/07/31 改成A4橫印
      'dllaccrpt103.Acc1430 Label9 & "," & ReportTitle(103), strStartDate & "," & strStartDate1, strEndDate & "," & strEndDate1, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
      SettingPrtOS
      With adoaccrpt103
         .MoveFirst
         PrintHeader
         Do While Not .EOF
            '小計
            If strGrp <> "" And strGrp <> "" & .Fields("R10303") Then
               Call PrintTot(1)
            End If
            
            strTmpSubTot = Val(strTmpSubTot) + Val("" & .Fields("R10307"))
            strTmpTot = Val(strTmpTot) + Val("" & .Fields("R10307"))
            strGrp = "" & .Fields("R10303")
            '對象名稱
            strTemp(0) = PUB_StrToStr("" & .Fields("R10303"), 30)
            '金額
            strTemp(1) = Format("" & .Fields("R10307"), "##,##0")
            '開票情形
            strTemp(2) = ""
            '摘要
            strTemp(3) = PUB_StrToStr("" & .Fields("R10306"), 54)
            '入帳日期
            strTemp(4) = ChangeTStringToTDateString("" & .Fields("R10304"))
            '應付款單號
            strTemp(5) = PUB_StrToStr("" & .Fields("R10305"), 10)
            '傳票號碼
            strTemp(6) = PUB_StrToStr("" & .Fields("R10311"), 10)
            
            For intP = 0 To 6
                If intP = 1 Or intP > 3 Then  '靠右
                   Printer.CurrentX = PLeft(intP) + (PLeft(intP + 1) - PLeft(intP) - Printer.TextWidth(strTemp(intP)) - ciColGap)
                Else
                   Printer.CurrentX = PLeft(intP)
                End If
                Printer.CurrentY = iPrint
                Printer.Print strTemp(intP)
            Next intP
            PrintNewLine
            .MoveNext
         Loop
      End With
      'Add by Amy 2023/06/28 加最後一個 小計 ex:J公司 往來類別:1 入帳日:112/05/01~112/06/07 欲處理日:112/06/10~112/06/10 顯示V0046 1筆及V0411 2筆, V0411無小計
      Call PrintTot(1)
      '列印-合計
      Call PrintTot(2)
      'end 2017/07/31
   End If
   adoaccrpt103.Close
   Printer.EndDoc 'Added by Lydia 2020/04/10
   
   'Added by Lydia 2017/07/31 還原系統預設印表機
   PUB_RestorePrinter strPrinter
   Printer.Orientation = IIf(mPrtOrt = 0, 1, mPrtOrt)
   'end 2017/07/31
   Screen.MousePointer = vbDefault
   FormClear
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102) 'Remove by Lydia 2017/07/31
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   'Remove by Lydia 2017/07/31
   'If KeyCode <> vbKeyEscape Then
   '   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   'End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5250
   'Modified by Lydia 2017/07/31
   'Me.Height = 3185 'Modify by Amy 2014/01/16 原:3000
   Me.Height = 3795
   'end 2017/07/31
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
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = DFormat
   'Add by Amy 2020/04/10 公司別改下拉
   Combo2.AddItem "", 0
   Call Pub_SetCboCmp(Combo2, False, False, False, , 1)
   'end 2020/04/10
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102) 'Remove by Lydia 2017/07/31
   
   'Modified by Lydia 2017/07/31 預設印表機選單
   'Set dllaccrpt103 = CreateObject("AccReport.ReportSelect")
   PUB_SetPrinter Me.Name, Combo1, strPrinter
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   'Modified by Lydia 2017/07/31 若印表機變動, 則更新列印設定
   'Set dllaccrpt103 = Nothing
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   'end 2017/07/31
   
   Set Frmacc1430 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim strSql As String
Dim bolDel As Boolean 'Add by Amy 2014/01/23 '避免總金額=0,銷貨折讓單未收回 重覆Del的錯誤
Dim strCmp As String 'Add by Amy 2020/04/10

On Error GoTo Checking
   strSql = ""
   bolDel = False 'Add by Amy 2014/01/23
   If Text1 <> MsgText(601) Then
      Select Case Text1
         Case Mid(ComboItem(91), 1, 1), Mid(ComboItem(93), 1, 1)
            strSql = "and a0o02 <> '2'"
         Case Mid(ComboItem(92), 1, 1)
            strSql = "and a0o02 = '2'"
         Case Else
            strSql = "and a0o02 <> '2'"
      End Select
   Else
      Exit Sub
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0o05 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0o05 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> MsgText(29) Then
      strSql = strSql & " and a0o06 >= " & Val(FCDate(MaskEdBox3.Text)) & ""
   End If
   If MaskEdBox4.Text <> MsgText(601) And MaskEdBox4.Text <> MsgText(29) Then
      strSql = strSql & " and a0o06 <= " & Val(FCDate(MaskEdBox4.Text)) & ""
   End If
   'Add by Amy 2020/04/10 公司別改下拉
   If Trim(Combo2) <> MsgText(601) Then
      strCmp = Combo2
      If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
      End If
   End If
   'end 2020/04/10
   lngCounter = 0
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26) 'Remove by Lydia 2017/07/31
   adoaccrpt103.CursorLocation = adUseClient
   adoaccrpt103.Open "select * from accrpt103", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0o0.CursorLocation = adUseClient
   'Modify by Morgan 2007/11/5
   'strExc(0) = "select * from acc0o0 where (a0o11 is null or a0o11 = '')"
   'Modify by Amy 2014/01/16 +公司別
   'Modify by Amy 2020/04/10 公司別改下拉 原:Text3
   strExc(0) = "select * from acc0o0,acc250 where (a0o11 is null or a0o11 = '') and a2505(+)=a0o01 And a0o07='" & strCmp & "' "
   If Text2 = "N" Then
      strExc(0) = strExc(0) & " and (a2502 is null or a2502<>'5' or  a2510 is not null) "
   End If
   strExc(0) = strExc(0) & strSql & " order by a0o02 asc, a0o03 asc"
   adoacc0o0.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0o0.RecordCount = 0 Then
      adoacc0o0.Close
      adoaccrpt103.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc0o0.EOF = False
      adoaccrpt103.AddNew
      adoaccrpt103.Fields("r10301").Value = strUserNum
      adoaccrpt103.Fields("r10302").Value = Counter
      If IsNull(adoacc0o0.Fields("a0o02").Value) Then
         adoaccrpt103.Fields("r10303").Value = Null
      Else
         adoaccrpt103.Fields("r10303").Value = adoacc0o0.Fields("a0o03").Value
         Select Case adoacc0o0.Fields("a0o02").Value
            Case Mid(ComboItem(91), 1, 1)
               adoaccrpt103.Fields("r10303").Value = adoaccrpt103.Fields("r10303").Value & A0i02Query(adoacc0o0.Fields("a0o03").Value)
               'Add by Morgan 2007/11/5 若翻譯費回執未收回時上'*'號
               If "" & adoacc0o0("a2502") = "5" And IsNull(adoacc0o0("a2510")) Then
                  adoaccrpt103.Fields("r10303").Value = adoaccrpt103.Fields("r10303").Value & "*"
               End If
               'end 2007/11/5
            Case Mid(ComboItem(92), 1, 1)
               adoaccrpt103.Fields("r10303").Value = adoaccrpt103.Fields("r10303").Value & CustomerQuery(adoacc0o0.Fields("a0o03").Value, 1)
            Case Mid(ComboItem(93), 1, 1)
               adoaccrpt103.Fields("r10303").Value = adoaccrpt103.Fields("r10303").Value & StaffQuery(adoacc0o0.Fields("a0o03").Value)
         End Select
      End If
      If IsNull(adoacc0o0.Fields("a0o05").Value) Then
         adoaccrpt103.Fields("r10304").Value = Null
      Else
         adoaccrpt103.Fields("r10304").Value = adoacc0o0.Fields("a0o05").Value
      End If
      adoaccrpt103.Fields("r10305").Value = adoacc0o0.Fields("a0o01").Value
      adoaccsum.CursorLocation = adUseClient
      'Modify by Amy 2014/01/16 改公司別 原:'1'
      'Modify by Amy 2020/04/10 公司別改下拉 原:Text3
      strExc(1) = "select a1p14 from acc1p0 where a1p01 = '" & strCmp & "' and a1p02 = 'B' and a1p04 = '" & adoacc0o0.Fields("a0o01").Value & "' and a1p05 in ('2112', '2113') union " & _
                     "select a1p14 from acc1p0 where a1p01 = '" & strCmp & "' and a1p02 = 'E' and a1p23 = '" & adoacc0o0.Fields("a0o01").Value & "' and a1p05 in ('2112', '2113') union " & _
                     "select a1p14 from acc1p0 where a1p01 = '" & strCmp & "' and a1p02 = 'Z' and a1p04 = '" & adoacc0o0.Fields("a0o09").Value & "' and a1p05 in ('2112', '2113')"
      adoaccsum.Open strExc(1), adoTaie, adOpenStatic, adLockReadOnly
      If adoaccsum.RecordCount <> 0 Then
         If IsNull(adoaccsum.Fields("a1p14").Value) Then
            adoaccrpt103.Fields("r10306").Value = Null
         Else
            adoaccrpt103.Fields("r10306").Value = adoaccsum.Fields("a1p14").Value
         End If
      Else
         adoaccrpt103.Fields("r10306").Value = Null
      End If
      adoaccsum.Close
      adoaccsum.CursorLocation = adUseClient
      'Modify by Amy 2014/01/16 改公司別 原:'1'
      'Modify by Amy 2020/04/10 公司別改下拉 原:Text3
      strExc(1) = "select sum(a1p08) from acc1p0 where a1p01 = '" & strCmp & "' and a1p02 = 'B' and a1p04 = '" & adoacc0o0.Fields("a0o01").Value & "' and a1p05 in ('2112', '2113') union " & _
                     "select sum(a1p08) from acc1p0 where a1p01 = '" & strCmp & "' and a1p02 = 'E' and a1p23 = '" & adoacc0o0.Fields("a0o01").Value & "' and a1p05 in ('2112', '2113') union " & _
                     "select sum(a1p08) from acc1p0 where a1p01 = '" & strCmp & "' and a1p02 = 'Z' and a1p04 = '" & adoacc0o0.Fields("a0o09").Value & "' and a1p05 in ('2112', '2113')"
      adoaccsum.Open strExc(1), adoTaie, adOpenStatic, adLockReadOnly
      If adoaccsum.RecordCount <> 0 Then
         If IsNull(adoaccsum.Fields(0).Value) Then
            adoaccrpt103.Fields("r10307").Value = 0
         Else
            adoaccrpt103.Fields("r10307").Value = adoaccsum.Fields(0).Value
         End If
      Else
         adoaccrpt103.Fields("r10307").Value = 0
      End If
      adoaccsum.Close
      If IsNull(adoacc0o0.Fields("a0o06").Value) Then
         adoaccrpt103.Fields("r10309").Value = Null
      Else
         adoaccrpt103.Fields("r10309").Value = adoacc0o0.Fields("a0o06").Value
      End If
      If IsNull(adoacc0o0.Fields("a0o02").Value) Then
         adoaccrpt103.Fields("r10310").Value = Null
      Else
         adoaccrpt103.Fields("r10310").Value = adoacc0o0.Fields("a0o02").Value
      End If
      adoacc1p0.CursorLocation = adUseClient
      'Modify by Amy 2014/01/16 改公司別 原:'1'
      'Modify by Amy 2020/04/10 公司別改下拉 原:Text3
      strExc(1) = "select a1p22 from acc1p0 where a1p01 = '" & strCmp & "' and a1p02 = 'B' and a1p04 = '" & adoacc0o0.Fields("a0o01").Value & "' and a1p05 in ('2112', '2113') union " & _
                     "select a1p22 from acc1p0 where a1p01 = '" & strCmp & "' and a1p02 = 'E' and a1p04 = '" & adoacc0o0.Fields("a0o09").Value & "' and a1p05 in ('2112', '2113') union " & _
                     "select a1p22 from acc1p0 where a1p01 = '" & strCmp & "' and a1p02 = 'Z' and a1p04 = '" & adoacc0o0.Fields("a0o09").Value & "' and a1p05 in ('2112', '2113')"
      adoacc1p0.Open strExc(1), adoTaie, adOpenStatic, adLockReadOnly
      If adoacc1p0.RecordCount <> 0 Then
         If IsNull(adoacc1p0.Fields(0).Value) Then
            adoaccrpt103.Fields("r10311").Value = Null
         Else
            adoaccrpt103.Fields("r10311").Value = adoacc1p0.Fields(0).Value
         End If
      Else
         adoaccrpt103.Fields("r10311").Value = Null
      End If
      adoacc1p0.Close
      If adoaccrpt103.Fields("r10307").Value = 0 Then
         adoaccrpt103.Delete
         bolDel = True 'Add by Amy 2014/01/23
      End If
      
      'Add by Amy 2014/01/23 +if
      If bolDel = False Then
         'Add by Amy 2014/01/16 剔除銷貨折讓單未收回者(a2510及2519都為null)
         If adoacc0o0.Fields("a0o07").Value = "J" And adoacc0o0.Fields("a0o19").Value = "2" And IsNull(adoacc0o0.Fields("a2510").Value) And IsNull(adoacc0o0.Fields("a2519").Value) Then
             adoaccrpt103.Delete
         End If
         'end 2014/01/16
      End If
      'end 2014/01/23
      adoaccrpt103.UpdateBatch
      adoacc0o0.MoveNext
   Loop
   adoacc0o0.Close
   adoaccrpt103.Close
   adoTaie.Execute "delete from accrpt103 where r10302 is null"
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
Private Sub Accrpt103Delete()
   adoTaie.Execute "delete from accrpt103"
End Sub

'*************************************************
'  計數器
'
'*************************************************
Private Function Counter() As Long
   lngCounter = lngCounter + 1
   Counter = lngCounter
End Function

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   'Modify by Amy 2020/04/10 公司別改下拉
   'Text3 = "" 'Add by Amy 2014/01/16
   Combo2 = ""
   'end 2020/04/10
   Text1 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = ""
   MaskEdBox3.Text = ""
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = ""
   MaskEdBox4.Text = ""
   MaskEdBox4.Mask = DFormat
   'Modify by Amy 2020/04/10 公司別改下拉
   'Text3.SetFocus 'Modify by Amy 2014/01/16
   Combo2.SetFocus
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
'   Select Case Text1
'      Case Mid(ComboItem(91), 1, 1)
'         MaskEdBox1.Enabled = True
'         MaskEdBox2.Enabled = True
'         MaskEdBox3.Enabled = False
'         MaskEdBox4.Enabled = False
'      Case Mid(ComboItem(92), 1, 1)
'         MaskEdBox1.Enabled = False
'         MaskEdBox2.Enabled = False
'         MaskEdBox3.Enabled = True
'         MaskEdBox4.Enabled = True
'      Case Mid(ComboItem(93), 1, 1), "4"
'         MaskEdBox1.Enabled = True
'         MaskEdBox2.Enabled = True
'         MaskEdBox3.Enabled = False
'         MaskEdBox4.Enabled = False
'   End Select
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
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox3.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox4.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

Private Sub Text2_GotFocus()
   TextInverse Text2
   CloseIme
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Mark by Amy 2020/04/10 公司別改下拉
'Add by Amy 2014/01/15
'Private Sub Text3_Change()
'    Label9.Caption = A0802Query(Text3)
'End Sub
'
'Private Sub Text3_GotFocus()
'    TextInverse Text3
'End Sub
'
'Private Sub Text3_KeyPress(KeyAscii As Integer)
'    KeyAscii = UpperCase(KeyAscii)
'End Sub
'
'Private Sub Text3_Validate(Cancel As Boolean)
'    If Text3 = "" Then Exit Sub
'    If Text3 <> "1" And Text3 <> "J" Then
'        MsgBox "公司別輸入錯誤請確認 ！"
'        Cancel = True
'        Exit Sub
'    End If
'End Sub
'end 2014/01/15
'end 2020/04/10

'Added by Lydia 2017/07/31 列印設定
Private Sub SettingPrtOS()

    '設定印表機
     mPrtOrt = Printer.Orientation
     Printer.EndDoc
     PUB_RestorePrinter Combo1
     Printer.PaperSize = 9  'A4
     Printer.Orientation = 2 '2.橫印
     
     '設定欄位位置
     If Val(PLeft(0)) = 0 Then
        lngPageHeight = Printer.ScaleHeight
        lngPageWidth = Printer.ScaleWidth
        lngLineHeight = 300
        Printer.Font.Name = "新細明體"
        Printer.Font.Size = ciFontSize
        
        PTitle(0) = "對象名稱"
        PLeft(0) = ciStartX
        
        PTitle(1) = "金額"
        PLeft(1) = PLeft(0) + Printer.TextWidth(String(15, "　")) + ciColGap
        
        PTitle(2) = "開票情形"
        PLeft(2) = PLeft(1) + Printer.TextWidth(String(6, "　")) + ciColGap
        
        PTitle(3) = "摘要"
        PLeft(3) = PLeft(2) + Printer.TextWidth(String(10, "　")) + ciColGap
        
        PTitle(4) = "入帳日期"
        PLeft(4) = PLeft(3) + Printer.TextWidth(String(26, "　")) + ciColGap
        
        PTitle(5) = "應付款單號"
        PLeft(5) = PLeft(4) + Printer.TextWidth(String(5, "　")) + ciColGap
        
        PTitle(6) = "傳票號碼"
        PLeft(6) = PLeft(5) + Printer.TextWidth(String(5, "　")) + ciColGap
        
        '傳票號碼(止)
        PLeft(7) = PLeft(6) + Printer.TextWidth(String(5, "　")) + ciColGap
     End If
           
     iPrint = 0
     iPage = 0
     strTmpTot = "0"
     strTmpSubTot = "0"
End Sub
'Added by Lydia 2017/07/31 換行
Private Sub PrintNewLine(Optional ByVal mRate As Single = 1, Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)
   iPrint = iPrint + lngLineHeight * mRate
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint

      iPage = iPage + 1
      Printer.NewPage
      PrintHeader
   End If
End Sub

'Added by Lydia 2017/07/31 列印表頭
Private Sub PrintHeader()
Dim x1 As Integer
Dim strCmp As String 'Add by Amy 2020/04/10

iPrint = ciStartY
iPageLine = 0

Printer.Font.Size = ciTitleFontSize
Printer.Font.Bold = True

'公司別
'Modify by Amy 2020/04/10 公司別改下拉 原:Text3
If Trim(Combo2) <> "" Then
    strCmp = Combo2
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    strCmp = A0802Query(strCmp)
    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strCmp)) / 2
    Printer.CurrentY = iPrint
    Printer.Print strCmp
    PrintNewLine
End If
'end 2020/04/10
         
'報表名稱
Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(ReportTitle(103))) / 2
Printer.CurrentY = iPrint
Printer.Print ReportTitle(103)
PrintNewLine

Printer.Font.Size = ciFontSize
Printer.Font.Bold = True

'第一行
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "名稱含 * 號表回執未收回"

Printer.CurrentX = PLeft(2) + 685
Printer.CurrentY = iPrint
Printer.Print "入帳日期：" & IIf(Val(FCDate(MaskEdBox1.Text)) > 0, convForm(MaskEdBox1.Text, 10), String(4, "　")) & " ~ " & _
                           IIf(Val(FCDate(MaskEdBox2.Text)) > 0, convForm(MaskEdBox2.Text, 10), String(4, "　"))

Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

'第二行
PrintNewLine
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "列印人員：" & strUserName

Printer.CurrentX = PLeft(2) + 500
Printer.CurrentY = iPrint
Printer.Print "欲處理日期：" & IIf(Val(FCDate(MaskEdBox3.Text)) > 0, convForm(MaskEdBox3.Text, 10), String(4, "　")) & " ~ " & _
                           IIf(Val(FCDate(MaskEdBox4.Text)) > 0, convForm(MaskEdBox4.Text, 10), String(4, "　"))
                           
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & Printer.Page

'第三行 抬頭
PrintNewLine
For x1 = 0 To 6
    '水平置中
    Printer.CurrentX = PLeft(x1) + (PLeft(x1 + 1) - PLeft(x1) - Printer.TextWidth(PTitle(x1))) / 2
    Printer.CurrentY = iPrint
    Printer.Print PTitle(x1)
Next x1

PrintNewLine
Printer.Line (PLeft(0), iPrint)-(PLeft(7), iPrint)
iPrint = iPrint + 150

Printer.Font.Bold = False

End Sub

'Added by Lydia 2017/07/31 列印小計／合計
Private Sub PrintTot(ByVal aKind As Integer)
    Printer.Line (PLeft(1), iPrint)-(PLeft(3), iPrint)
    iPrint = iPrint + 150
    Printer.CurrentX = PLeft(1) - 800
    Printer.CurrentY = iPrint
    If aKind = 1 Then
        Printer.Print "小計："
        strExc(0) = Format(Val(strTmpSubTot), "##,##0")
        Printer.CurrentX = PLeft(1) + (PLeft(2) - PLeft(1) - Printer.TextWidth(strExc(0)) - ciColGap)
        Printer.CurrentY = iPrint
        Printer.Print strExc(0)
        strTmpSubTot = "0"
        PrintNewLine
    Else
        Printer.Print "合計："
        strExc(0) = Format(Val(strTmpTot), "##,##0")
        Printer.CurrentX = PLeft(1) + (PLeft(2) - PLeft(1) - Printer.TextWidth(strExc(0)) - ciColGap)
        Printer.CurrentY = iPrint
        Printer.Print strExc(0)
        PrintNewLine
        Printer.Line (PLeft(1), iPrint)-(PLeft(3), iPrint)
        iPrint = iPrint + 80
        Printer.Line (PLeft(1), iPrint)-(PLeft(3), iPrint)
        
        PrintNewLine
        
        Printer.Font.Bold = True
        strExc(1) = "***結束***"
        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strExc(1))) / 2
        Printer.CurrentY = iPrint
        Printer.Print strExc(1)
    End If
    
End Sub
