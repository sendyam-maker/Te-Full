VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frmacc11t0 
   AutoRedraw      =   -1  'True
   Caption         =   "發票上傳作業"
   ClientHeight    =   4350
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   6050
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4350
   ScaleWidth      =   6050
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   30
      TabIndex        =   15
      Top             =   3780
      Width           =   5790
      _ExtentX        =   10231
      _ExtentY        =   441
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CmdTran 
      BackColor       =   &H00C0FFC0&
      Caption         =   "上傳"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1470
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   900
      Width           =   1000
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   12
      Top             =   480
      Width           =   1150
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3450
      MaxLength       =   10
      TabIndex        =   11
      Text            =   "MX12345678"
      Top             =   480
      Width           =   1150
   End
   Begin VB.CommandButton CmdQuery 
      BackColor       =   &H00C0FFC0&
      Caption         =   "查詢上傳錯誤"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2580
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   900
      Width           =   3250
   End
   Begin VB.FileListBox File1 
      Height          =   420
      Left            =   30
      TabIndex        =   3
      Top             =   330
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox ChkBox 
      Caption         =   "要上傳折讓(銷退)資料"
      Height          =   225
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   1
      Top             =   1320
      Width           =   5790
   End
   Begin VB.CommandButton CmdXML 
      BackColor       =   &H00C0FFC0&
      Caption         =   "轉檔(&E)"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   60
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   900
      Width           =   1300
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   750
      Top             =   330
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   3450
      TabIndex        =   6
      Top             =   60
      Width           =   1155
      _ExtentX        =   2046
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
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
      Left            =   4680
      TabIndex        =   7
      Top             =   60
      Width           =   1155
      _ExtentX        =   2046
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
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
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   516
      Left            =   0
      TabIndex        =   16
      Top             =   4080
      Width           =   5900
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4590
      TabIndex        =   13
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "發票號碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2550
      TabIndex        =   10
      Top             =   570
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "上傳日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2550
      TabIndex        =   9
      Top             =   90
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4590
      TabIndex        =   8
      Top             =   60
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "按下「轉檔」產生XML，確認無誤，再按「上傳」"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   510
      Left            =   60
      TabIndex        =   4
      Top             =   435
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc11t0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/08 Form2.0已檢查 (無需修改的物件)
'2019/03/05 Create by Amy
Option Explicit

Const strFixInv_S = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & "<INVOICE>" & vbCrLf
Const strFixInv_E = "</INVOICE>"
Const strInvoiceType = "07" '發票類別-07.一般稅額計算之電子發票
Const strDonateMark = "0" '捐贈註記-0.非捐贈發票
Const strCustomsClearanceMark = "1" '通關方式註記-1.非經海關出口
Const strQty = "1" '數量
Const strTaxType = "1" '課稅別-1.應說
Const strTaxRate = "0.05" '稅率

Dim XmlPath As String, strPOSID As String, strPOSSN As String
Dim stToPath As String '檔案上傳路徑
Dim stQPath As String '上傳錯誤路徑
Dim strSellerID As String, strSellerAddr As String, strSellerTag As String '營業人統編/賣方地址/賣方資料(Taie)
Dim adoMain As New ADODB.Recordset
Dim adoQ As New ADODB.Recordset
Dim strQ As String, strFileName1 As String, strFileName2 As String
Dim i As Integer, intSeq As Integer
Dim arrA04MainF() As Variant, arrA04DetF() As Variant, arrA04AmtF() As Variant
Dim arrC04AF() As Variant, arrC04BF() As Variant, arrC04CF() As Variant
Dim arrA05() As Variant, arrC05() As Variant, arrB05() As Variant, arrD05() As Variant
Dim bolBtoB As Boolean '產生 B to B Tag
Dim stBuyerTag As String, stBuyerID As String, stBuyerName As String, stBuyerAddr As String, stCusNo As String
Dim stPCName As String, stRunMsg As String 'Add by Amy 2024/01/16
'Add by Amy 2024/12/20 避免同天正在上傳之發票,又產生xml檔 之條件語法-避免有未改到
Const stSqlFix As String = "And Not Exists(Select * From AccTmp11t0,BookRecord Where a4301=SubStr(R002,5,10) And R001 is null And r003=br04(+) And r004=br05(+) ) "

Private Function FormCheck() As Boolean
    Dim bCancel As Boolean
    Dim stTmp As String
    
    FormCheck = False
    If (MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29)) And (MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29)) _
        And Trim(Text2) = MsgText(601) And Trim(Text3) = MsgText(601) Then
        MsgBox "請擇一輸入查詢條件！", , MsgText(5)
        Text2.SetFocus
    End If
        
    If Val(FCDate(MaskEdBox1.Text)) > 0 Or Val(FCDate(MaskEdBox2.Text)) > 0 Then
        If Val(FCDate(MaskEdBox1.Text)) = 0 Then
            MsgBox "上傳日期起日不可為空！", , MsgText(5)
            Exit Function
        End If
        If Val(FCDate(MaskEdBox2.Text)) = 0 Then
            MsgBox "上傳日期迄日不可為空！", , MsgText(5)
            Exit Function
        End If
        Call MaskEdBox1_Validate(bCancel)
        If bCancel = True Then Exit Function
        '每月批次會刪除3個月前的資料
        If Val(Mid(DBDATE(DateAdd("m", -3, Format(strSrvDate(1), "####/##/##"))), 1, 6)) - 191100 >= Val(Replace(Mid(MaskEdBox1.Text, 1, 6), "/", "")) Then
            MsgBox "上傳日期迄日不可輸3個月前！", , MsgText(5)
            Exit Function
        End If
        Call MaskEdBox2_Validate(bCancel)
        If bCancel = True Then Exit Function
    End If
    
    If Trim(Text2) <> MsgText(601) Or Trim(Text3) <> MsgText(601) Then
        If Left(Text2, 2) <> Left(Text3, 2) Then
            MsgBox "發票號碼起迄前2碼英文不可以不一樣！", , MsgText(5)
            Exit Function
        End If
    End If
    
    FormCheck = True
    
End Function

Private Sub cmdQuery_Click()
    Dim j As Integer, k As Integer
    Dim bolCrossY As Boolean, bolCrossM As Boolean '是跨年/是跨月
    Dim stErrData As String, stSPath As String '搜尋路徑
    Dim stYear1 As String, stMonth1 As String, stDate1 As String
    Dim stYear2 As String, stMonth2 As String, stDate2 As String
    Dim intStartM As Integer, intEndM As Integer, intStartD As Integer, intEndD As String '起迄月/日
    Dim stInvNo As String
    Dim stTmp
        
    If FormCheck = False Then Exit Sub
    Text1 = ""
    
    '起
    stYear1 = Mid(MaskEdBox1, 1, 3)
    stMonth1 = Mid(MaskEdBox1, 5, 2)
    stDate1 = Mid(MaskEdBox1, 8, 2)
    '迄
    stYear2 = Mid(MaskEdBox2, 1, 3)
    stMonth2 = Mid(MaskEdBox2, 5, 2)
    stDate2 = Mid(MaskEdBox2, 8, 2)
    
    '若未輸入日期抓三個月前的日期(AutoBatch會刪三個月前的錯誤資料)
    If (MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29)) And (MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29)) Then
        stDate1 = DBDATE(DateAdd("m", -3, Format(strSrvDate(1), "####/##/##")))
        stYear1 = Val(Mid(stDate1, 1, 4)) - 1911
        stMonth1 = Mid(stDate1, 5, 2)
        stDate1 = "01"
        stYear2 = Mid(strSrvDate(1), 1, 4) - 1911
        stMonth2 = Mid(strSrvDate(1), 5, 2)
        stDate2 = "31"
    End If
    
    intStartM = Val(stMonth1): intEndM = Val(stMonth2)
    intStartD = Val(stDate1): intEndD = Val(stDate2)
    '跨年
    If stYear1 <> stYear2 Then
        bolCrossY = True
        intStartM = stYear1
        intEndM = 12
        intEndD = 31
    '只跨月
    ElseIf stMonth1 <> stMonth2 Then
        bolCrossM = True
        intEndD = 31
    End If
    
    '錯誤訊息依年/月/日分資料夾
    For i = Val(stYear1) To Val(stYear2)
        For j = intStartM To intEndM
            For k = Val(stDate1) To Val(stDate2)
                stSPath = stQPath & i + 1911 & "\" & Format(j, "00") & "\" & Format(k, "00")
                If Dir(stSPath, vbDirectory) <> "" Then
                    stErrData = stErrData & ";" & Replace(PUB_GetFileListOrderby(stSPath & "\", "*.xml", True), "||", ";")
                End If
            Next k
        Next j
        '下個年度改起始月
        If bolCrossY = True And intEndM = 12 Then
            intStartM = 1
            If Val(stYear2) = i Then
                bolCrossY = False
                intEndM = Val(stMonth2)
            End If
        End If
        If bolCrossM = True And intEndD = 31 Then
            intStartD = 1
            If Val(stMonth2) = j Then
                bolCrossM = False
                intEndM = Val(stDate2)
            End If
        End If
    Next i
    If stErrData <> MsgText(601) Then
        stTmp = Split(Mid(stErrData, 2), ";")
        For i = LBound(stTmp) To UBound(stTmp)
            stInvNo = stTmp(i)
            stInvNo = Mid(stInvNo, InStr(stInvNo, "_") + 1)
            stInvNo = Mid(stInvNo, 1, InStr(stInvNo, "_") - 1)
            If Text2 <> MsgText(601) Then
                If Left(stInvNo, 2) = Left(Text2, 2) _
                    And Val(Mid(stInvNo, 3)) >= Val(Mid(Text2, 3)) And Val(Mid(stInvNo, 3)) <= Val(Mid(Text3, 3)) Then
                    Text1 = Text1 & stInvNo & vbCrLf
                End If
            Else
                Text1 = Text1 & stInvNo & vbCrLf
            End If
        Next i
        Text1 = "上傳有誤之發票號碼：(請通知電腦中心處理)" & vbCrLf & Text1
    Else
        MsgBox "查無資料！"
    End If
         
End Sub

Private Sub CmdXML_Click()
     '設定Taie資料
    SetSellerTag
    '建立資料夾
    'Modify by Amy 2021/06/22 先判斷桌面有沒有xls資料夾
    If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
        MkDir strExcelPath
    End If
    If Dir(Mid(XmlPath, 1, Len(XmlPath) - 1), vbDirectory) = MsgText(601) Then
        MkDir XmlPath
    End If
   
    Text1 = "": strSellerTag = "": ProgressBar1.Value = 0
    Screen.MousePointer = vbHourglass
    Call SetSellerTag
    Text1 = "產生 發票訊息格式-開始" & vbCrLf & Text1
    stRunMsg = ""
    If TXT_InvoiceData = False Then
        Screen.MousePointer = vbDefault
        MsgBox "產生 發票有誤,請通知電腦中心！"
        Exit Sub
    End If
    Text1 = "產生 發票訊息格式-結束" & vbCrLf & Text1
    
    ProgressBar1.Value = 0
    Text1 = "產生 發票作廢資料-開始" & vbCrLf & Text1
    stRunMsg = ""
    If TXT_CancelInvoice = False Then
        Screen.MousePointer = vbDefault
        MsgBox "產生 發票作廢資料有誤,請通知電腦中心！"
        Exit Sub
    End If
    Text1 = "產生 發票作廢資料-結束" & vbCrLf & Text1
    Screen.MousePointer = vbDefault
    
    ProgressBar1.Value = 0
    If ChkBox.Value = 1 Then
        Text1 = "產生 折讓資料-開始" & vbCrLf & Text1
        stRunMsg = ""
        If TXT_DisCountData = False Then
            Screen.MousePointer = vbDefault
            MsgBox "產生 折讓資料有誤,請通知電腦中心！"
            Exit Sub
        End If
        Text1 = "產生 折讓資料-結束" & vbCrLf & Text1
        Screen.MousePointer = vbDefault
    End If
    '折讓作廢目前無(也不會修改)先不做
    '1131015 瑞婷產出折讓單 BH29497939 (為折讓單號及當時的發票號碼) 且已上傳盟立後,智權通知要作廢此折讓單之操作:
    '1.[財務]需至盟立操作折讓單作廢(婉莘操作)
    '2.[電腦中心]需更新Acc430 Tag 並於資料刪改記錄備註
    '   ACC430折讓作廢Tag拿掉 語法:Update Acc430 Set A4310=Null,A4324=Null,A4325=Null Where A4301='發票號' '(ex:發票號:BH29497939)
    
    Text1 = vbCrLf & Text1
    If Dir(XmlPath & "*.*") = MsgText(601) Then
        MsgBox "無檔案產生！"
    Else
        MsgBox "XML 檔案已產生！"
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Dim stChoose As String
    Dim intX As Integer, intY As Integer, sglWidth As Single, sglHeight As Single
    Dim strMsg As String 'Add by Amy 2019/12/27
    
    Label4.Caption = "": Label4.Visible = False 'Add by Amy 2024/10/24
    Me.Icon = LoadPicture(strIcoPath)
    strFormName = Name
    'Modify by Amy 2023/10/06 原 W:6000/H:4485
    Me.Width = 6170
    Me.Height = 4800
    Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
    Image1 = LoadPicture(strBackPicPath1)
    sglWidth = Image1.Width
    sglHeight = Image1.Height
    For intX = 0 To Int(ScaleWidth / sglWidth)
        For intY = 0 To Int(ScaleHeight / sglHeight)
            PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
        Next
    Next
    
    'Modify by Amy 2024/01/16 轉pdf程式移機,改抓系統特殊設定
    stPCName = Pub_GetSpecMan("分信主機名稱", False)
    '正式上傳設定
    XmlPath = strExcelPath & "EInvoice\"
    strPOSID = "1" '正式-POS機編號
    strPOSSN = "1B120D101718DDC40199" '正式-通道金鑰
    
'*** Memo 程式說明 ***
    '1.轉檔可選正式或測式Tag(以切的資料庫判斷)
    '2.電腦中心操作 [上傳] 鈕,只能上傳至 [測試] 資料庫(只能使用【 測試 】資料庫產生的資料)
    '3.若需產生之測式xml Tag問題為何,可使用 [PostMan] 軟體(將xml Tag貼至PostMan軟體中)
    '3.使用【 正式】資料庫產生的資料,需將產生之檔案[自行] copy User 產生之xml 資料夾後,請User按上傳
    '    當天產生發票未上傳,又馬上作廢後按上傳,導致盟立無對應之發票資料而錯誤,已作廢發票台一程式無法還原,拿類似[發票訊息格式],改資料及檔名上傳用
'*** End Memo 程式說明 ***
    
    'Modify by Amy 2024/09/25 A2004更換電腦,電腦名稱已修改 原:\\AA2004-\EInvoice\ (資料夾要設共用及開EveryOne權限)
    stToPath = "\\" & pub_HostName & "\EInvoice\"
    'Modify by Amy 2024/10/24 修改訊息,讓電腦中心知道如何上傳至盟立正式平台 (1131024 大樓突然斷電,無法用Amy電腦按上傳至M51-APP)
    If Pub_StrUserSt03 = "M51" Then
       Label4.Visible = True
       strMsg = "1.產生【 正式 】XML資料　2.產生《測試 》XML資料" & vbCrLf & _
                        "！！以上都只會上傳至盟立《測試平台》！！" & vbCrLf & vbCrLf & _
                        "需產生【 正式 】XML資料上傳至盟立【 正式平台】" & vbCrLf & _
                        "請使用VB修改變數"
        stChoose = Trim(InputBox(strMsg, "請選資料型態", "2"))
        '測試機用
        'Modify by Amy 2019/12/27 +strMsg提醒
       If strUserNum <> "A2004" Then
            stToPath = "\\A2004\EInvoice\"
        End If
        If stChoose = "2" Then
            strPOSID = "1" '測式-POS機編號
            strPOSSN = "f547c52b0be7c5e73125" '測式-通道金鑰
            strMsg = "會上傳至盟立《測試平台 》,請確認：" & vbCrLf & _
                        "1.至盟立《測試平台》網站,確認是否已新增字軌" & vbCrLf & _
                        "2.上傳至盟立《測試平台》,轉PDF程式需在Amy電腦執行" & vbCrLf & _
                        "    請確認Amy電腦是否開啟"
            If strUserNum <> "A2004" Then
               strMsg = strMsg & "！！！請確認程式與目前Amy電腦名稱是否相符！！！"
            End If
        Else
            strMsg = "請確認是否產生【 正式 】XML資料" & vbCrLf & _
                        "目前預設上傳為盟立《測試平台》,請注意" & vbCrLf & _
                        "1.【 正式 】XML資料上傳至盟立《測試平台》會錯誤" & vbCrLf & _
                        "2.若需產生【 正式 】XML資料上傳至盟立【 正式 】平台" & vbCrLf & _
                        "    請使用VB修改變數"
        End If
        MsgBox strMsg
        'end 2019/12/27
        
        'Mark by Amy 2019/12/27 改成 run \\AA2004-
'        If Dir(Mid(stToPath, 1, Len(stToPath) - 1), vbDirectory) = MsgText(601) Then
'            MkDir stToPath
'        End If
    Else
        stToPath = "\\" & stPCName & "\Einvoice\"
    End If
    
    If Label4.Visible = True Then
      Label4.Caption = "目前設定產生" & IIf(stChoose = "1", "【 正式】", "《測試 》") & "XML資料 -> " & _
                                       "上傳主機為 " & Replace(UCase(stToPath), "EINVOICE\", "")
    End If
    'end 2024/10/24
    Me.Caption = Me.Caption & "(" & stToPath & ")"
    
    '查詢錯誤使用
    stQPath = "\\" & stPCName & "\c$\551cron\Error\"
    'end 2024/01/16
    MaskEdBox1.Mask = ""
    MaskEdBox1.Text = CFDate(strSrvDate(2))
    MaskEdBox1.Mask = DFormat
    MaskEdBox2.Mask = ""
    MaskEdBox2.Text = CFDate(strSrvDate(2))
    MaskEdBox2.Mask = DFormat
    Text2 = "": Text3 = ""
    'Modify by Amy 2023/05/17 +允許電腦中心可下畫面條件產生(單筆)資料-測式用,故許電腦中心不預帶
    If Pub_StrUserSt03 <> "M51" Then Call SetTodayInvoiceNo '預帶當日最小及最大發票號碼
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strFormName = MsgText(601)
    KeyEnter vbKeyEscape
    MenuEnabled
    Set Frmacc11t0 = Nothing
End Sub

Private Sub CmdTran_Click()
    Dim oFilObj As FileSystemObject
    Dim RsQ As New ADODB.Recordset, strQ As String, stIns As String, intQ As Integer
    Dim stStep As String, stInvoice As String, stTranTime As String, stTmp As String, arrTmp 'Add by Amy 2024/09/24
    Dim bolTrans As Boolean, bolMoveFile As Boolean 'Add by Amy 2024/10/24
    
On Error GoTo ErrHand
    'Memo 電腦中心操作[上傳]鈕,請參閱 Form_Load 程式說明
   
    '讀取資料夾檔案
    File1.path = XmlPath
    File1.Refresh
    If File1.ListCount = 0 Then
        MsgBox XmlPath & "無檔案需上傳", vbExclamation
        Exit Sub
    End If
    '判斷是否有未關閉之檔案
    For i = 0 To File1.ListCount - 1
        If PUB_ChkFileOpening(XmlPath & File1.List(i)) = True Then
            MsgBox XmlPath & File1.List(i) & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        'Add by Amy 2024/09/24 記錄檔名前14碼
        'Modify by Amy 2025/02/04 +if ,銷退檔名=xmlTag+_+I單號9碼(a0s01)
        If Mid(File1.List(i), 5, 1) = "I" Then
            stInvoice = stInvoice & ";" & Left(File1.List(i), 13)
        Else
            stInvoice = stInvoice & ";" & Left(File1.List(i), 14)
        End If
    Next i
    '判斷Server是否有未上傳(於AutoPdf刪其Tag)
    strQ = "Select * From BookRecord Where br01=111111 "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        '避免Server正在傳 Server的File1未Refresh
        MsgBox "Server上仍有未上傳資料" & vbCrLf & "不可再上傳。", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
         
    '判斷是否有檔案未上傳(上傳主機 "C:\EInvoice\" ,其他資料夾user 可能有權限問題,故於轉pdf檢查)
    stStep = stToPath  'Add by Amy 2024/09/24 移機後利於判斷問題
    If Dir(stToPath) <> MsgText(601) Then
        MsgBox "Server (" & stStep & ")上仍有未上傳資料" & vbCrLf & "！！請通知電腦中心！！", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Add by Amy 2024/09/24 避免常發生上傳更新Tag重覆,將上傳之檔名前14碼寫入暫檔
    If stInvoice = MsgText(601) Then
        MsgBox "無stInvoce資料" & vbCrLf & "！！請通知電腦中心！！", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Text1 = String(20, "-") & vbCrLf & vbCrLf & Text1
    Text1 = "開始" & vbCrLf & Text1
    
    '檔案移至Server
    stStep = "搬移檔案至Server"
    
   'Modify by Amy 2024/10/24 改先移檔,確定成功再寫DB
'*** 搬移檔案 ***
    Set oFilObj = New FileSystemObject
    oFilObj.MoveFile XmlPath & "*.*", stToPath
    bolMoveFile = True
    Set oFilObj = Nothing
    If ChkXml(1, stToPath) = False Then
      stTmp = "已執行" & stStep & vbCrLf & _
                      "但 " & stToPath & " 仍無上傳之檔案"
      MsgBox stTmp & vbCrLf & "！！請通知電腦中心！！", vbExclamation
      Text1 = "　" & stStep & "-->失敗" & vbCrLf & Text1
      Text1 = "　！！請通知電腦中心！！" & vbCrLf & Text1
      Text1 = vbCrLf & "結束" & vbCrLf & Text1
      Screen.MousePointer = vbDefault
      Exit Sub
    End If
    Text1 = "　" & stStep & "-->成功" & vbCrLf & Text1
                        
'*** 更新上傳Tag ***
    stStep = "寫AccTmp11t0 Tag"
    'Modify by Amy 2024/09/24 +stStep/AccTmp11t0/BookRecord增加寫入日期時間
    stTranTime = ServerTime
    '** 上傳資料寫入暫存檔 **
    arrTmp = Split(Mid(stInvoice, 2), ";")
    
    cnnConnection.BeginTrans
    bolTrans = True
    For i = LBound(arrTmp) To UBound(arrTmp)
      stIns = "Insert Into AccTmp11t0 (ID,R002,R003,R004) " & _
                   "Values('" & strUserNum & "','" & arrTmp(i) & "'," & strSrvDate(1) & "," & stTranTime & ")"
      cnnConnection.Execute stIns
      Text1 = "　" & arrTmp(i) & vbCrLf & Text1 'Add by Amy 2024/11/04
    Next i
    Text1 = "　" & stStep & "-->成功" & vbCrLf & Text1
    
    '** 寫入觸發上傳盟立之Tag (轉PDF作業用)**
    stStep = "寫BookRecord Tag"
    stIns = "Insert into BookRecord (br01,br04,br05) Values(111111," & strSrvDate(1) & "," & stTranTime & ")"
    cnnConnection.Execute stIns
    'end 2024/09/24
    
    cnnConnection.CommitTrans
    Text1 = "　" & stStep & "-->成功" & vbCrLf & Text1
    Text1 = "結束" & vbCrLf & Text1
    
    MsgBox "上傳完成！"
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
   If bolMoveFile = True Then Set oFilObj = Nothing
   If bolTrans = True Then cnnConnection.RollbackTrans
   
   If stStep = stToPath And Err.Number = 52 Then
      'Err.Description=不正確的檔案名稱或數目
      stTmp = stStep & "->抓不到此路徑"
   ElseIf stStep = "搬移檔案至Server" And Err.Number = 76 Then
      stTmp = stToPath & "->請確認資料夾[權限][共用]頁籤權限已開"
   ElseIf InStr(stStep, "Tag") > 0 Then
      stTmp = "　" & stStep & "-->失敗"
      If bolMoveFile = True And ChkXml(1, stToPath) = True Then
         '搬回檔案
         Set oFilObj = New FileSystemObject
         oFilObj.MoveFile stToPath & "*.*", XmlPath
         Set oFilObj = Nothing
         If Dir(XmlPath) = MsgText(601) Then
            stTmp = "　！！檔案搬回" & XmlPath & "失敗！！" & vbCrLf & stTmp
         End If
      End If
   'end 2024/10/24
   Else
      'Err.Description=沒有權限(需確認[共用]EveryOne權限全開)
      stTmp = Err.Description
      If stStep <> MsgText(601) Then
         stTmp = stTmp & "-" & stStep
      End If
   End If
   Text1 = "　！！！上傳有誤！！！" & vbCrLf & _
                  stTmp & vbCrLf & Text1
   
   MsgBox "上傳有誤，請留畫面訊息" & vbCrLf & "請通知電腦中心！" & vbCrLf, vbExclamation
   Text1 = vbCrLf & "結束" & vbCrLf & Text1
   Screen.MousePointer = vbDefault
End Sub

'產生畫面條件發票資料(Tag不要轉大寫)
Private Function TXT_InvoiceData() As Boolean
    Dim rsA As New ADODB.Recordset
    Dim stQ As String, intQ As Integer
    Dim stOldInvNo As String, stOldAxc02 As String, stSys As String
    Dim stC0401 As String, stA0401 As String, stA0401_Fix As String, stC0401CTag As String, stC0401DTag As String 'C0401/A0401
    Dim stDetailTag As String, stTmp As String, intRow As Integer '列數
    Dim objStream1 As Object
    Dim stCU01 As String, stCU02 As String 'Add by Amy 2023/05/17
    Set objStream1 = CreateObject("ADODB.Stream")

On Error GoTo ErrHand
    
    TXT_InvoiceData = False
    'A0401 設定
    'Modify by Amy 2024/11/04 盟立加Tag:PrintMark/RandomNumber/ZeroTaxRateReason
    arrA04MainF = Array("InvoiceNumber", "InvoiceDate", "InvoiceTime", "Seller", "Buyer", _
                                "CheckNumber", "BuyerRemark", "MainRemark", "CustomsClearanceMark", "Category", _
                                "RelateNumber", "InvoiceType", "GroupMark", "DonateMark", "BondedAreaConfirm", _
                                "Attachment", "PrintMark", "RandomNumber", "ZeroTaxRateReason")
    'Modify by Amy 2024/11/04 盟立加Tag:TaxType
    arrA04DetF = Array("Description", "Quantity", "Unit", "UnitPrice", "Amount", _
                                    "SequenceNumber", "Remark", "RelateNumber", "TaxType")
    arrA04AmtF = Array("SalesAmount", "TaxType", "TaxRate", "TaxAmount", "TotalAmount", _
                                    "DiscountAmount", "OriginalCurrencyAmount", "ExchangeRate", "Currency")
    
    '=== A0401存證發票 Fix Tag ===
    stA0401_Fix = "<INVOICE_CODE>A0401</INVOICE_CODE>" & vbCrLf & _
                                "<POSSN>" & strPOSSN & "</POSSN>" & vbCrLf & _
                                "<POSID>" & strPOSID & "</POSID>" & vbCrLf & _
                                "<SYSTIME>" & Format(strSrvDate(1), "####-##-##") & " " & Format(Now, "HH:mm:ss") & "</SYSTIME>" & vbCrLf
    '=== C0401開立發票 D(Fix) Tag ===
    stC0401DTag = stC0401DTag & _
                        "<D1>" & strSellerID & "</D1>" & vbCrLf & _
                        "<D2>" & strPOSSN & "</D2>" & vbCrLf & _
                        "<D3>" & strPOSID & "</D3>" & vbCrLf & _
                        "<D4>" & Format(strSrvDate(1), "####-##-##") & " " & Format(Now, "HH:mm:ss") & "</D4>" & vbCrLf & _
                        "<D5></D5>" & vbCrLf
    'C0401設定
    Call SetC0401Field
     
    '抓取發票資料產生 C0401開立發票訊息格式/A0401平台存證開立發票訊息格式
    'Modify by Amy 2023/05/17 +a0k05,允許電腦中心下畫面條件產生(單筆)資料-測式用
    stQ = ""
    If Pub_StrUserSt03 = "M51" And (Text2 <> MsgText(601) Or Text3 <> MsgText(601)) Then
        If Text2 <> MsgText(601) Then
            stQ = stQ & "And a4301>='" & Text2 & "' "
        End If
        If Text3 <> MsgText(601) Then
            stQ = stQ & "And a4301<='" & Text3 & "' "
        End If
    End If
    'Modify by Amy 2024/10/29 +AccTmp11t0,避免同天正在上傳之發票,又產生xml檔
    'Modify by Amy 2024/12/20 將2024/10/29 條件寫成變數,避免有未改到
    stQ = "Select acc430.*,sqltime(a4313) as a4313T,axc02,a0k20,st01,st15,a0k03,a0k04,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as cuname,a0k33,a0k05 " & _
               "From Acc430,Acc431,Acc0k0,Customer,Staff " & _
               "Where a4301=axc01(+) And axc02=a0k01(+) And a0k20=st01(+) And substr(a0k03,1,8)=cu01(+) And substr(a0k03,9,1)=cu02(+) " & _
               "And a4302 >= " & Val(TranInvoiceDate) & " And Nvl(a4319,'0')='0' " & stQ & stSqlFix & _
               " Order by a4302, a4301"
    'end 2023/05/17
    intQ = 1
    Set adoMain = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        ProgressBar1.max = adoMain.RecordCount
        VarClear
        bolBtoB = False: intSeq = 0
        adoMain.MoveFirst
        Do While adoMain.EOF = False
            Call VarClear '清空變數值
            
            '一個發票號一個檔案
            If stOldInvNo <> "" & adoMain.Fields("a4301") And stOldInvNo <> MsgText(601) Then
                With objStream1
                    If bolBtoB = True Then
                       .WriteText strFixInv_S & stA0401_Fix & stA0401 & strFixInv_E
                    Else
                        .WriteText strFixInv_S & stC0401 & stC0401DTag & strFixInv_E
                    End If
                    bolBtoB = False
                    .SaveToFile strFileName1
                    .Close
                    Text1 = "   " & stOldInvNo & vbCrLf & Text1
                End With
                stC0401 = "": stA0401 = "": stC0401CTag = ""
            End If
            
            strFileName1 = "" & adoMain.Fields("a4301")
            '=== 買方資料設定 ====
            stBuyerID = "" & adoMain.Fields("a4303")
            '境外公司(統一編號 8個0)算個人
            If stBuyerID <> MsgText(601) And stBuyerID <> "00000000" Then
                bolBtoB = True
                '有買方統編產生A0401
                 strFileName1 = "A04_" & strFileName1
            Else
                '沒買方統編產生C0401
                stBuyerID = "0000000000" '統一編號若為空值(個人:傳10個0)
                strFileName1 = "C04_" & strFileName1
            End If
            stBuyerName = RepSpecWord(Trim("" & adoMain.Fields("a0k04")))
            '當天開立又取消尚未上傳,Acc0k0無資料,BuyerName會是空值但為必填Tag需填4個0
            If stBuyerName = MsgText(601) And "" & adoMain.Fields("a4308") <> MsgText(601) Then
                stBuyerName = "0000"
            End If
            'Modify by Amy 2023/05/17 改以a0k04 抓資料-秀玲
'            '收據抬頭為3個字以下(含3個字)以a0k03抓地址
'            'Modify by Amy 2019/07/25 因發票明細需show 營業地址(公司),故三摺(個人)show聯絡地址-瑞婷
'            If Len(Trim("" & adoMain.Fields("a0k04").Value)) <= 3 Then
'                If bolBtoB = True Then
'                    stBuyerAddr = GetCusAddr2(Trim("" & adoMain.Fields("a0k03")), True)
'                    'Memo by Amy 2023/05/17 不會有 有統編但抬頭<=3的資料
'                Else
'                    stBuyerAddr = GetCusAddr(Trim("" & adoMain.Fields("a0k03")), True)
'                End If
'            '以收據抬頭a0k04抓客戶地址,若不存在,再抓收據抬頭資料檔acc420的營業地址抬頭
'            Else
'                If bolBtoB = True Then
'                    stBuyerAddr = GetCusAddr2(Trim("" & adoMain.Fields("a0k04")), False)
'                Else
'                    'ex:境外公司 a4303(統編 bolBtoB = False)=8個0->產生「B to C個人」Tag
'                    'ex:E11117477/無統編/發票HK64861029/抬頭[江西大田精密科技有限公司]
'                    stBuyerAddr = GetCusAddr(Trim("" & adoMain.Fields("a0k04")), False)
'                End If
'            End If
'            'end 2019/07/25
            stCU01 = "": stCU02 = ""
            stRunMsg = "發票[" & adoMain.Fields("a4301") & "]抬頭+客戶編號抓資料" 'Add by Amy 2024/11/04
            '避免多筆資料抓錯,先抓抬頭+客戶編號,抓不到再抓抬頭資料的第一筆(同frmacc1440)-秀玲
            If GetTitleCustData("" & adoMain.Fields("a0k04").Value, "" & adoMain.Fields("a0k03"), "", stCU01, stCU02, , , , , , , , , , , , , , , , , , False, , , , , , , , , , , , Me.Name) = False Then
                Call GetTitleCustData("" & adoMain.Fields("a0k04").Value, "", "", stCU01, stCU02)
            End If
            '公司(有a4303-統編) or 可扣繳(公司 or 境外公司-->[境外公司]產生的是[個人]的Tag)
            'cu23中文地址 -> a4215 營業地址
            If bolBtoB = True Or "" & adoMain.Fields("a0k05") = "2" Then
                If stCU01 & stCU02 = MsgText(601) Then
                    stBuyerAddr = GetCusAddr2(Trim("" & adoMain.Fields("a0k04")), False)
                Else
                    stBuyerAddr = GetCusAddr2(stCU01 & stCU02, True)
                End If
            '個人 cu31聯絡地址 ->cu23 中文地址 ->a4203 郵寄地址 ->a4215 營業地址
            Else
                If stCU01 & stCU02 = MsgText(601) Then
                    stBuyerAddr = GetCusAddr(Trim("" & adoMain.Fields("a0k04")), False)
                Else
                    stBuyerAddr = GetCusAddr(stCU01 & stCU02, True)
                End If
            End If
            'end 2023/05/17
            stBuyerTag = GetBuyerTag(IIf(bolBtoB = True, "A04", "C04"))
            '=== End 買方資料設定 ====
          
            strFileName1 = XmlPath & strFileName1 & "_" & strSrvDate(2) & ".xml"
            '檔案存在先刪除
            If Dir(strFileName1) <> MsgText(601) Then
                Kill strFileName1
            End If
            '開啟檔案
            With objStream1
                .Type = adTypeText
                .Mode = 3
                .Open
                .Position = 0
                .Charset = "UTF-8"
            End With
            
            '*** Main Tag ***
            '有買方統編產生A0401(BtoB)
            If bolBtoB = True Then
               '=== A0401存證發票 ===
                stA0401 = stA0401 & "<Main>" & vbCrLf
                For i = LBound(arrA04MainF) To UBound(arrA04MainF)
                    stTmp = ""
                    Select Case arrA04MainF(i)
                        Case "InvoiceNumber" '發票號碼
                            stTmp = "" & adoMain.Fields("a4301")
                        Case "InvoiceDate" '發票日期
                            stTmp = Format(Val("" & adoMain.Fields("a4302")) + 19110000, "####-##-##")
                        Case "InvoiceTime" '發票時間
                            stTmp = Format("" & adoMain.Fields("a4313T"), "HH:mm:ss")
                        Case "Seller" '賣方
                            stTmp = strSellerTag
                        Case "Buyer" '買方
                            stTmp = stBuyerTag
                        Case "MainRemark" '總備註
                            stTmp = "" & adoMain.Fields("st15") & " " & adoMain.Fields("st01")
                        Case "InvoiceType" '發票類別
                            stTmp = strInvoiceType
                        Case "DonateMark" '捐贈註記
                            stTmp = strDonateMark
                        Case "CustomsClearanceMark" '通關方式註記
                            '「零稅率」時「通關方式註記」必填 1.非經海關出口
                            If "" & adoMain.Fields("a4323") = "Y" Then stTmp = strCustomsClearanceMark
                        Case "PrintMark" '紙本電子發票已列印註記 'Add by Amy 2024/11/04
                            stTmp = "Y"
                        '"CheckNumber","BuyerRemark", "Category","RelateNumber"
                        '發票檢查碼,買受人註記,沖帳別,相關號碼
                        '"GroupMark","BondedAreaConfirm","Attachment"
                        '彙開註記,買受人簽署適用零稅率註記,證明附件
                        '"RandomNumber","ZeroTaxRateReason"
                        '發票防偽隨機碼,零稅率原因
                    End Select
                    If arrA04MainF(i) <> "Seller" And arrA04MainF(i) <> "Buyer" Then
                        stTmp = "<" & arrA04MainF(i) & ">" & stTmp & "</" & arrA04MainF(i) & ">" & vbCrLf
                    Else
                        stTmp = stTmp
                    End If
                    stA0401 = stA0401 & stTmp
                Next i
                stA0401 = stA0401 & "</Main>" & vbCrLf
                
            '沒買方統編產生C0401(BtoC)
            Else
                '=== C0401開立發票 ===
                'Modify by Amy 2024/10/29 盟立加欄位 原:30
                For i = 1 To 33
                    stTmp = ""
                    Select Case i
                        '訊息類型-A1
                        Case GetFieldVal("C04A", "Invoice_Code")
                            stTmp = "C0401"
                        '發票號碼-A2
                        Case GetFieldVal("C04A", "InvoiceNumber")
                            stTmp = "" & adoMain.Fields("a4301")
                        '發票開立日期-A3
                        Case GetFieldVal("C04A", "InvoiceDate")
                            stTmp = Format(Val("" & adoMain.Fields("a4302")) + 19110000, "####-##-##")
                        '發票開立時間-A4
                        Case GetFieldVal("C04A", "InvoiceTime")
                            stTmp = Format("" & adoMain.Fields("a4313T"), "HH:mm:ss")
                        '買方統編-A5
                        Case GetFieldVal("C04A", "BuyerIdentifier")
                            stTmp = stBuyerID
                        '買方名稱-A6
                        Case GetFieldVal("C04A", "BuyerName")
                            stTmp = stBuyerName
                        '買方地址-A7
                        Case GetFieldVal("C04A", "BuyerAddress")
                            stTmp = stBuyerAddr
                        '買方負責人-A8(顯示於地址下方xxx收)
                        Case GetFieldVal("C04A", "BuyerPrincipal")
                            stTmp = stBuyerName
                        '客戶編號-A12
                        Case GetFieldVal("C04A", "BuyerCusNo")
                            stTmp = stCusNo
                        '總備註-A16(印於QCode下方)
                        Case GetFieldVal("C04A", "MainRemark")
                            stTmp = "" & adoMain.Fields("st15") & " " & adoMain.Fields("st01")
                        '通關方式註記-A17
                        Case GetFieldVal("C04A", "CustomsClearanceMark")
                            '「零稅率」時「通關方式註記」必填 1.非經海關出口
                            If "" & adoMain.Fields("a4323") = "Y" Then stTmp = strCustomsClearanceMark
                        '發票類別-A22
                        Case GetFieldVal("C04A", "InvoiceType")
                            stTmp = strInvoiceType
                        '捐贈註記-A24
                        Case GetFieldVal("C04A", "DonateMark")
                            stTmp = strDonateMark
                        '紙本電子發票已列印註記-A28
                        Case GetFieldVal("C04A", "PrintMark")
                            stTmp = "Y"
                        '發票防偽隨機碼-A30
                        Case GetFieldVal("C04A", "InvoceRandomNo")
                            '只有B to C才有「發票防偽隨機碼」-盟立文件C0401 p25
                            stTmp = "" & adoMain.Fields("a4301")
                            'Modify by Amy 2021/03/03 固定4碼 原:GetInvChkNumber(stTmp) ,傳入MN31530000回傳6,只產生3碼數字上傳會錯
                            stTmp = Format(GetInvChkNumber(stTmp), "00") & Format(Int((99 * Rnd) + 1), "00")
                        'Add by Amy 2024/10/29 '零稅率原因-A33
                        Case GetFieldVal("C04A", "ZeroTaxRateReason")
                            If "" & adoMain.Fields("a4323") = "Y" Then
                              stTmp = "72" '營業稅法第7條,第二款
                            End If
                        '"BuyerTEL", "BuyerFAX", "BuyerEmail", "BuyerRole"
                        '電話,傳真,Email,營業人角色註記
                        '"CheckNumber","BuyerRemark","GroupMark"
                        '發票檢查碼,買受人註記,彙開註記
                        '"TaxServiceName","AllowDate","AllowDoc","AllowNo"
                        '稅捐稽徵處,核准日,核准文,核准號
                        '"CarrierType", "CarrierId1", "CarrierId2", "NPOBAN"
                        '載具類別號碼,載具顯碼id,載具隱碼id,捐贈對象
                        '"RelateNumber","BondedAreaConfirm"
                        '相關號碼-A31,買受人簽署適用零稅率註記-A32
                    End Select
                    stC0401 = stC0401 & "<A" & i & ">" & stTmp & "</A" & i & ">" & vbCrLf
                Next i
            End If
            '*** End Main Tag ***

            '*** 明細 ***
            stRunMsg = "抓取發票明細資料"
            '依帳款類別變更欄位抓取發票明細資料
            If IsNull(adoMain.Fields("a0k33")) = False Then
                stDetailTag = GetInvoiceDetail1("" & adoMain.Fields("a4308"))
            Else
                stDetailTag = GetInvoiceDetail2("" & adoMain.Fields("a4308"))
            End If
            If bolBtoB = True Then
                stA0401 = stA0401 & "<Details>" & vbCrLf & stDetailTag & "</Details>" & vbCrLf
            Else
                'Modify by Amy 2019/08/21 於各筆明細加<B>…</B>
                'stC0401 = stC0401 & "<B>" & vbCrLf & stDetailTag & "</B>" & vbCrLf
                stC0401 = stC0401 & vbCrLf & stDetailTag & vbCrLf
            End If
                        
            '*** Amounts Tag ***
            stRunMsg = "Amounts Tag"
            '有買方統編 產生A0401(BtoB)
            If bolBtoB = True Then
                '=== A0401存證發票 ===
                stA0401 = stA0401 & "<Amount>" & vbCrLf
                For i = LBound(arrA04AmtF) To UBound(arrA04AmtF)
                    stTmp = ""
                    Select Case arrA04AmtF(i)
                        Case "SalesAmount" '應稅銷售額合計
                            stTmp = Val("" & adoMain.Fields("a4304"))
                        Case "TaxType" '課稅別
                            stTmp = strTaxType
                            If "" & adoMain.Fields("a4323") = "Y" Then stTmp = "2" '零稅率
                        Case "TaxRate" '稅率
                            stTmp = strTaxRate
                            If "" & adoMain.Fields("a4323") = "Y" Then stTmp = "0" '零稅率
                        Case "TaxAmount" '營業稅額
                            stTmp = Val("" & adoMain.Fields("a4305"))
                        Case "TotalAmount" '總計
                            stTmp = Val("" & adoMain.Fields("a4304")) + Val("" & adoMain.Fields("a4305"))
                        '"DiscountAmount","OriginalCurrencyAmount","ExchangeRate","Currency"
                        '扣抵金額,原幣金額,匯率,幣別
                    End Select
                    stA0401 = stA0401 & "<" & arrA04AmtF(i) & ">" & stTmp & "</" & arrA04AmtF(i) & ">" & vbCrLf
                Next i
                stA0401 = stA0401 & "</Amount>" & vbCrLf
            '沒買方統編 產生C0401(BtoC)
            Else
                '=== C0401開立發票 ===
                For i = 1 To 13
                    stTmp = ""
                    Select Case i
                        '應稅銷售額合計-C1
                        Case GetFieldVal("C04C", "SalesAmount")
                            stTmp = Val("" & adoMain.Fields("a4304"))
                            If "" & adoMain.Fields("a4323") = "Y" Then stTmp = "0" '零稅率時需為0
                        '免稅銷售額合計-C2
                        Case GetFieldVal("C04C", "FreeTaxSalesAmount")
                            stTmp = "0" '固定0,稅內含
                        '零稅率銷售額合計-C3
                        Case GetFieldVal("C04C", "ZeroTaxSalesAmount")
                            stTmp = "0" '固定0,稅內含
                            If "" & adoMain.Fields("a4323") = "Y" Then stTmp = Val("" & adoMain.Fields("a4304")) '零稅率時=應稅銷售額
                        '課稅別-C4
                        Case GetFieldVal("C04C", "TaxType")
                            stTmp = strTaxType
                            If "" & adoMain.Fields("a4323") = "Y" Then stTmp = "2" '零稅率(應稅銷售額合計/免稅銷售額合計 都需為0)
                        '稅率-C5
                        Case GetFieldVal("C04C", "TaxRate")
                            stTmp = strTaxRate
                            If "" & adoMain.Fields("a4323") = "Y" Then stTmp = "0" '零稅率
                        '營業稅額-C6
                        Case GetFieldVal("C04C", "TaxAmount")
                            stTmp = "0" '無買方統編固定為 0-盟立文件C0401 p25
                        '總計-C7
                        Case GetFieldVal("C04C", "TotalAmount") '總計=應稅銷售額,稅內含
                            stTmp = Val("" & adoMain.Fields("a4304"))
                        '備註二-C13
                        Case GetFieldVal("C04C", "Remark2") '會於明細聯顯示
                            stTmp = "" & adoMain.Fields("st15") & " " & adoMain.Fields("st01")
                        '"DiscountAmount","OriginalCurrencyAmount","ExchangeRate","Currency","Remark1
                        ''扣抵金額,原幣金額,匯率,幣別,備註一(C12印於QCode下方)
                    End Select
                     stC0401CTag = stC0401CTag & "<C" & i & ">" & stTmp & "</C" & i & ">" & vbCrLf
                Next i
                stC0401 = stC0401 & stC0401CTag
            End If
            '*** End Amounts Tag ***
            
            ProgressBar1.Value = ProgressBar1.Value + 1
            stOldInvNo = "" & adoMain.Fields("a4301")
            adoMain.MoveNext
        Loop
    End If
    adoMain.Close
    If stA0401 & stC0401 <> MsgText(601) Then
        stRunMsg = "存XML"
        With objStream1
            If bolBtoB = True Then
                .WriteText strFixInv_S & stA0401_Fix & stA0401 & strFixInv_E
            Else
                .WriteText strFixInv_S & stC0401 & stC0401DTag & strFixInv_E
            End If
            .SaveToFile strFileName1
            .Close
            Text1 = "   " & stOldInvNo & vbCrLf & Text1
        End With
    End If
    
    TXT_InvoiceData = True
    Exit Function
    
ErrHand:
    strExc(9) = "產生 有誤-" & Err.Description
    If InStr(stRunMsg, "抬頭+客戶編號") > 0 Then
      '測式時訊息,若使用[測式平台]測試[作廢]發票,要先產生一個正常發票上傳後,再上傳作廢發票
      strExc(9) = "產生 有誤-確認發票是否已作廢" & vbCrLf & _
                           "　(" & stRunMsg & "有誤)"
    ElseIf stRunMsg <> MsgText(601) Then
      strExc(9) = strExc(9) & "(" & stRunMsg & ")"
    End If
    Text1 = strExc(9) & vbCrLf & Text1
    If adoMain.State = adStateOpen Then adoMain.Close
    If objStream1.State = adStateOpen Then objStream1.Close
   
End Function

'產生折讓(銷退)發票資料
Private Function TXT_DisCountData() As Boolean
    Dim rsA As New ADODB.Recordset
    Dim stQ As String, intQ As Integer
    Dim stOldDisNo As String
    Dim arrMainF() As Variant, arrDetF() As Variant
    Dim objStream1 As Object
    Dim stDisMainTag As String, stDisDetTag As String, stDisAmtTag As String
    Dim stFix As String, stFix_S As String, stFix_E As String
    Dim stTmp As String
    Set objStream1 = CreateObject("ADODB.Stream")

On Error GoTo ErrHand
    
    TXT_DisCountData = False
    arrMainF = Array("AllowanceNumber", "AllowanceDate", "S_Address", "Identifier", "Name", "Address", _
                                "AllowanceType")
    arrDetF = Array("OriginalInvoiceDate", "OriginalInvoiceNumber", "OriginalDescription", "Quantity", "UnitPrice", _
                            "Amount", "Tax", "AllowanceSequenceNumber", "TaxType")
  
    '===  Fix Tag ===
    stFix_S = Replace(strFixInv_S, "<INVOICE>", "<Allowance>")
    stFix = "<SELLERID>" & strSellerID & "</SELLERID>" & vbCrLf & _
                "<POSSN>" & strPOSSN & "</POSSN>" & vbCrLf & _
                "<POSID>" & strPOSID & "</POSID>" & vbCrLf & _
                "<SYSTIME>" & Format(strSrvDate(1), "####-##-##") & " " & Format(Now, "HH:mm:ss") & "</SYSTIME>" & vbCrLf
    stFix_E = Replace(strFixInv_E, "</INVOICE>", "</Allowance>")
     
    '抓取發票資料產生 折讓=銷退(台一只有銷退)+轉開資料(發票有轉開日a4310 ex:發票當時開100,需作銷退30) 訊息格式
    'Modify by Amy 2019/12/16 原抓Acc0s0.*並拿掉 And a0s04='3' 且需加入有發票轉開日期且電子發票折讓上傳日為空(先給客戶發票但未付款,已申報後轉開)
    'Modify by Amy 2024/10/29 +AccTmp11t0,避免同天正在上傳之發票,又產生xml檔
    'Modify by Amy 2025/02/04 語法再調整
    stQ = "Select * From (" & _
               "Select a0s01,a0s02,a0s03,a0s04,a0s26,a4303,a0k03,a0k04 " & _
               "From Acc0s0,Acc430,Acc431,Acc0k0 " & _
               "Where a0s26=a4301(+) And a0s26 is not null And a4301=axc01(+) And axc02=a0k01(+) " & _
               " And a0s03 >= " & Val(TranInvoiceDate) & " And Nvl(A0s28,0)=0 "
    stQ = stQ & "Union All " & _
               "Select a4301 as a0s01,'' as a0s02,a4310 as a0s03,'' as a0s04,a4301 as a0s26,a4303,a0k03,a0k04 " & _
               "From Acc430,Acc431,Acc0k0 " & _
               "Where a4301=axc01(+) And SubStr(axc02,1,9)=a0k01(+) And Nvl(a4310,0)>=" & TranInvoiceDate & _
               " And Nvl(a4324,0)=0 " & _
               ") Where " & Replace(Mid(stSqlFix, 4), "a4301", "a0s01") & " Order by a0s03, a0s26"
    bolBtoB = False: intQ = 1
    Set adoMain = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        ProgressBar1.max = adoMain.RecordCount
        adoMain.MoveFirst
        Do While adoMain.EOF = False
            Call VarClear '清空變數值
            
            '一個折讓單號一個檔案
            If stOldDisNo <> "" & adoMain.Fields("a0s01") And stOldDisNo <> MsgText(601) Then
                With objStream1
                    If bolBtoB = True Then
                        .WriteText stFix_S & "<INVOICE_CODE>B0401</INVOICE_CODE>" & stFix & vbCrLf & stDisMainTag & stFix_E
                    Else
                        .WriteText stFix_S & "<INVOICE_CODE>D0401</INVOICE_CODE>" & stFix & vbCrLf & stDisMainTag & stFix_E
                    End If
                    bolBtoB = False
                    .SaveToFile strFileName1
                    .Close
                    Text1 = "   " & stOldDisNo & vbCrLf & Text1
                End With
                stDisMainTag = "": stDisDetTag = "": stDisAmtTag = ""
            End If
            strFileName1 = "" & adoMain.Fields("a0s01")
            
            '=== 買方資料設定 ====
            stBuyerID = "" & adoMain.Fields("a4303")
            '境外公司(統一編號 8個0)算個人
            If stBuyerID <> MsgText(601) And stBuyerID <> "00000000" Then
                bolBtoB = True
                '有買方統編產生B0401
                 strFileName1 = "B04_" & strFileName1
            Else
                '沒買方統編產生D0401
                stBuyerID = "0000000000" '統一編號若為空值(個人:傳10個0)
                strFileName1 = "D04_" & strFileName1
            End If
            stBuyerName = RepSpecWord(Trim("" & adoMain.Fields("a0k04")))
            '收據抬頭為3個字以下(含3個字)以a0k03抓地址
            'Modify by Amy 2019/07/25 因發票明細需show 營業地址(公司),故三摺(個人)show聯絡地址-瑞婷
            If Len(Trim("" & adoMain.Fields("a0k04"))) <= 3 Then
                If bolBtoB = True Then
                    stBuyerAddr = GetCusAddr2(Trim("" & adoMain.Fields("a0k03")), True)
                Else
                    stBuyerAddr = GetCusAddr(Trim("" & adoMain.Fields("a0k03")), True)
                End If
            '以收據抬頭a0k04抓客戶地址,若不存在,再抓收據抬頭資料檔acc420的營業地址抬頭
            Else
                If bolBtoB = True Then
                    stBuyerAddr = GetCusAddr2(Trim("" & adoMain.Fields("a0k04")), False)
                Else
                    stBuyerAddr = GetCusAddr(Trim("" & adoMain.Fields("a0k04")), False)
                End If
            End If
            'end 2019/07/25
            '=== End 買方資料設定 ====
           
            strFileName1 = XmlPath & strFileName1 & "_" & strSrvDate(2) & ".xml"
            '檔案存在先刪除
            If Dir(strFileName1) <> MsgText(601) Then
                Kill strFileName1
            End If
            '開啟檔案
            With objStream1
                .Type = adTypeText
                .Mode = 3
                .Open
                .Position = 0
                .Charset = "UTF-8"
            End With
       
            '=== Tag ===
            For i = LBound(arrMainF) To UBound(arrMainF)
                stTmp = ""
                Select Case arrMainF(i)
                    Case "AllowanceNumber"
                        stTmp = "" & adoMain.Fields("a0s01")
                    Case "AllowanceDate"
                        stTmp = Format(Val("" & adoMain.Fields("a0s03")) + 19110000, "####-##-##")
                    Case "S_Address" '賣方地址
                        stTmp = strSellerAddr
                    Case "Identifier" '買方統編
                        stTmp = stBuyerID
                    Case "Name" '買方名稱
                        stTmp = stBuyerName                        '
                    Case "Address" '買方地址
                        stTmp = stBuyerAddr
                    Case "AllowanceType" '折讓種類
                        stTmp = "2"
                        stDisDetTag = GetDisDetailTag("" & adoMain.Fields("a0s01"), "" & adoMain.Fields("a0s26"), arrDetF(), stDisAmtTag)
                End Select
                stDisMainTag = stDisMainTag & "<" & UCase(arrMainF(i)) & ">" & stTmp & "</" & UCase(arrMainF(i)) & ">" & vbCrLf
                '插入明細Tag
                If arrMainF(i) = "AllowanceType" Then
                    stDisMainTag = stDisMainTag & stDisDetTag
                End If
            Next i
            stDisMainTag = stDisMainTag & stDisAmtTag
           
            ProgressBar1.Value = ProgressBar1.Value + 1
            stOldDisNo = "" & adoMain.Fields("a0s01")
            adoMain.MoveNext
        Loop
    End If
    adoMain.Close
    
    If stDisMainTag <> MsgText(601) Then
        With objStream1
            If bolBtoB = True Then
                .WriteText stFix_S & "<INVOICE_CODE>B0401</INVOICE_CODE>" & stFix & vbCrLf & stDisMainTag & stFix_E
            Else
                .WriteText stFix_S & "<INVOICE_CODE>D0401</INVOICE_CODE>" & stFix & vbCrLf & stDisMainTag & stFix_E
            End If
            .SaveToFile strFileName1
            .Close
            Text1 = "   " & stOldDisNo & vbCrLf & Text1
        End With
    End If
    
    TXT_DisCountData = True
    Exit Function
    
ErrHand:
    Text1 = "產生有誤-" & Err.Description & vbCrLf & Text1
    If adoMain.State = adStateOpen Then adoMain.Close
    If objStream1.State = adStateOpen Then objStream1.Close
    
End Function

'產生畫面條件發票作廢資料
Private Function TXT_CancelInvoice() As Boolean
    Dim rsA As New ADODB.Recordset
    Dim stQ As String, intQ As Integer
    Dim stOldInvNo As String, stTag As String
    Dim stTmp As String
    Dim objStream1 As Object
    Set objStream1 = CreateObject("ADODB.Stream")

On Error GoTo ErrHand
    
    TXT_CancelInvoice = False
    'A0501 設定
    'Modify by Amy 2024/11/04 盟立加Tag:B_EMAIL_ADDRESS
    arrA05 = Array("INVOICE_CODE", "POSSN", "POSID", "SYSTIME", "CancelInvoiceNumber", _
                                "InvoiceDate", "BuyerId", "SellerId", "CancelDate", "CancelTime", _
                                "CancelReason", "B_EMAIL_ADDRESS", "ReturnTaxDocumentNumber", "Remark")
    'C0501 設定
    arrC05 = Array("Invoice_Code", "POSSN", "POSID", "Invoice_Number", "Invoice_Date", _
                                "BuyerId", "SellerId", "Cancel_Date", "Cancel_Time", "Cancel_Reason", _
                                "B_EMAIL_ADDRESS", "ReturnTaxDocuemtn_Number", "Remark", "SYSTIME")
    'end 2024/11/04
        
    '抓取發票作廢資料產生 A0501/C0501發票作廢訊息格式
    'Modify by Amy 2024/10/29 +AccTmp11t0,避免同天正在上傳之發票,又產生xml檔
    'Modify by Amy 2024/12/20 將2024/10/29 條件寫成變數,避免有未改到
    stQ = "Select * From Acc430 " & _
               "Where a4308 is not null And a4308 >= " & Val(TranInvoiceDate) & " And Nvl(a4321,'0')='0' " & stSqlFix & _
               " Order by a4302, a4301"
    intQ = 1
    Set adoMain = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        ProgressBar1.max = adoMain.RecordCount
        VarClear
        bolBtoB = False: intSeq = 0
        adoMain.MoveFirst
        Do While adoMain.EOF = False
            Call VarClear '清空變數值
            
            '一個發票號一個檔案
            If stOldInvNo <> "" & adoMain.Fields("a4301") And stOldInvNo <> MsgText(601) Then
                With objStream1
                    .WriteText strFixInv_S & stTag & strFixInv_E
                    bolBtoB = False
                    .SaveToFile strFileName1
                    .Close
                    Text1 = "   " & stOldInvNo & vbCrLf & Text1
                End With
                stTag = ""
            End If
            
            strFileName1 = "" & adoMain.Fields("a4301")
            '=== 買方資料設定 ====
            stBuyerID = "" & adoMain.Fields("a4303")
            '境外公司(統一編號 8個0)算個人
            If stBuyerID <> MsgText(601) And stBuyerID <> "00000000" Then
                bolBtoB = True
                '有買方統編產生A0501
                 strFileName1 = "A05_" & strFileName1
            Else
                '沒買方統編產生C0501
                stBuyerID = "0000000000" '統一編號若為空值(個人:傳10個0)
                strFileName1 = "C05_" & strFileName1
            End If
          
            strFileName1 = XmlPath & strFileName1 & "_" & strSrvDate(2) & ".xml"
            '檔案存在先刪除
            If Dir(strFileName1) <> MsgText(601) Then
                Kill strFileName1
            End If
            '開啟檔案
            With objStream1
                .Type = adTypeText
                .Mode = 3
                .Open
                .Position = 0
                .Charset = "UTF-8"
            End With
            
            '*** Main Tag ***
            '有買方統編產生A0501-Tag不轉大寫某些Tag固定大寫某些故定小寫否則會抓不到Tag(BtoB)
            If bolBtoB = True Then
               '=== A0501存證發票作廢 ===
                For i = LBound(arrA05) To UBound(arrA05)
                    stTmp = ""
                    Select Case arrA05(i)
                        Case "INVOICE_CODE" '訊息代碼
                            stTmp = "A0501"
                        Case "POSSN" 'POS機出廠序號(通道金鑰)
                            stTmp = strPOSSN
                        Case "POSID" 'POS機編號
                            stTmp = strPOSID
                        Case "SYSTIME" '系統時間
                            stTmp = Format(Now, "HH:mm:ss")
                        Case "CancelInvoiceNumber" '發票編號
                            stTmp = "" & adoMain.Fields("a4301")
                        Case "InvoiceDate" '發票開立日期
                            stTmp = Format(Val("" & adoMain.Fields("a4302")) + 19110000, "####-##-##")
                        Case "BuyerId" '買方統編
                            stTmp = stBuyerID
                        Case "SellerId" '賣方統編
                            stTmp = strSellerID
                        Case "CancelDate" '作廢日期
                            stTmp = Format(Val("" & adoMain.Fields("a4308")) + 19110000, "####-##-##")
                        Case "CancelTime" '作廢時間
                            stTmp = "00:00:00" 'Modify by Amy 2020/07/02 上傳至財政部會錯 原:24:00:00 -盟立
                        Case "CancelReason" '作廢原因
                            stTmp = "資料有誤"
                        '"B_EMAIL_ADDRESS","ReturnTaxDocumentNumber","Remark"
                        '買方EMAIL信箱,專案作廢核准文號,備註
                    End Select
                    stTag = stTag & "<" & arrA05(i) & ">" & stTmp & "</" & arrA05(i) & ">" & vbCrLf
                Next i
                
            '沒買方統編產生C0501(BtoC)
            Else
                '=== C0501發票作廢 ===
                For i = LBound(arrC05) To UBound(arrC05)
                    stTmp = ""
                    Select Case arrC05(i)
                        Case "Invoice_Code" '訊息類型
                            stTmp = "C0501"
                        Case "POSSN"  'POS機出廠序號(通道金鑰)
                            stTmp = strPOSSN
                        Case "POSID" 'POS機編號
                            stTmp = strPOSID
                        Case "Invoice_Number" '發票號碼
                            stTmp = "" & adoMain.Fields("a4301")
                        Case "Invoice_Date" '發票開立日期
                            stTmp = Format(Val("" & adoMain.Fields("a4302")) + 19110000, "####-##-##")
                        Case "BuyerId" '買方統編
                            stTmp = stBuyerID
                        Case "SellerId" '賣方統編
                            stTmp = strSellerID
                        Case "Cancel_Date" '作廢日期
                            stTmp = Format(Val("" & adoMain.Fields("a4308")) + 19110000, "####-##-##")
                        Case "Cancel_Time" '作廢時間
                            stTmp = "24:00:00"
                        Case "Cancel_Reason" '作廢原因
                            stTmp = "資料有誤"
                        Case "SYSTIME" '系統時間
                            stTmp = Format(Now, "HH:mm:ss")
                        '"ReturnTaxDocuemtn_Number","Remark"
                        '專案作廢核准文號,備註
                    End Select
                    stTag = stTag & "<" & UCase(arrC05(i)) & ">" & stTmp & "</" & UCase(arrC05(i)) & ">" & vbCrLf
                Next i
            End If
            '*** End Main Tag ***
                       
            ProgressBar1.Value = ProgressBar1.Value + 1
            stOldInvNo = "" & adoMain.Fields("a4301")
            adoMain.MoveNext
        Loop
    End If
    adoMain.Close
    If stTag <> MsgText(601) Then
        With objStream1
            .WriteText strFixInv_S & stTag & strFixInv_E
            .SaveToFile strFileName1
            .Close
            Text1 = "   " & stOldInvNo & vbCrLf & Text1
        End With
    End If
    
    TXT_CancelInvoice = True
    Exit Function
    
ErrHand:
    Text1 = "產生 有誤-" & Err.Description & vbCrLf & Text1
    If adoMain.State = adStateOpen Then adoMain.Close
    If objStream1.State = adStateOpen Then objStream1.Close
   
End Function

'抓取1月第一筆發票號
Private Function GetJanInvNo(ByVal stYM As String) As String
    Dim rsA As New ADODB.Recordset
    Dim stQ As String, intQ As Integer
    
    stQ = "Select * From Acc410 Where a4101=" & Val(stYM)
    intQ = 1
    Set rsA = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        GetJanInvNo = "" & rsA.Fields("a4103") & Format(rsA.Fields("a4104"), String(8, "0"))
    End If
    rsA.Close
End Function

'抓取買方資料(參考frmacc1610.PrintAddress)
'回傳地址
Private Function GetCusAddr(ByVal stCu As String, bolIsNo As Boolean) As String
    Dim rsA As New ADODB.Recordset
    Dim stQ As String, intQ As Integer
   
    GetCusAddr = ""
    'Modify by Amy 2023/05/17 原:會傳入 a0k03,改都抓a0k04(抬頭),傳入前會先以GetTitleCustData函數抓客戶編號再傳入
    If bolIsNo = True Then
        stQ = "CU01='" & Left(stCu, 8) & "' And CU02='" & Mid(stCu, 9, 1) & "' "
    '    If bolIsNo = False Then
    '        stQ = "CU04='" & stCu & "' "
    '    End If
        'Modify by Amy 2023/03/01 +國內同業、解除對造、不得代理專利、不得代理商標
        'Modify By Sindy 2025/6/27 +or cu80='其他' or cu80='業務自行處理' or cu80='國內同業' or cu80='解除對造' or cu80='不得代理專利' or cu80='不得代理商標'
        '                          改抓常變數
        stQ = stQ & "and (cu80 is null or instr('" & 客戶及代理人可讀取的狀態 & "',cu80)>0) and cu02=0 "
        
        stQ = "Select cu01,cu02,cu04,cu112,cu23,cu30,cu31 From Customer Where " & stQ
        intQ = 1
        Set rsA = ClsLawReadRstMsg(intQ, stQ)
        If intQ = 1 Then
            '先抓聯絡地址(與GetCusAddr2不同處)
            If Trim("" & rsA.Fields("cu31")) <> MsgText(601) Then
                GetCusAddr = Trim("" & rsA.Fields("cu30")) & " " & Trim("" & rsA.Fields("cu31"))
            '再抓中文地址
            ElseIf Trim("" & rsA.Fields("cu23")) <> MsgText(601) Then
                GetCusAddr = Trim("" & rsA.Fields("cu112")) & " " & Trim("" & rsA.Fields("cu23"))
            End If
        End If
    Else
    '以收據抬頭抓客戶檔之CU04,若為客戶資料則抓中文地址,若不存在,則再抓收據抬頭資料檔acc420的營業地址
    'ElseIf bolIsNo = False Then
    'end 2023/05/17
        stQ = "Select * From Acc420 Where a4201='" & stCu & "'"
        If rsA.State = adStateOpen Then rsA.Close
         intQ = 1
        Set rsA = ClsLawReadRstMsg(intQ, stQ)
        If intQ = 1 Then
            '先抓郵寄絡地址(與GetCusAddr2不同處)
            If Trim("" & rsA.Fields("a4203")) <> MsgText(601) Then
                GetCusAddr = Trim("" & rsA.Fields("a4203"))
            '再抓營業地址
            ElseIf Trim("" & rsA.Fields("a4215")) <> MsgText(601) Then
                GetCusAddr = Trim("" & rsA.Fields("a4215"))
            End If
        End If
    End If
    rsA.Close
End Function

'抓取買方資料(參考PUB_PrintCaseReceipt_Inv)
'回傳地址
Private Function GetCusAddr2(ByVal stCu As String, bolIsNo As Boolean) As String
    Dim rsA As New ADODB.Recordset
    Dim stQ As String, intQ As Integer
    
    GetCusAddr2 = ""
    'Modify by Amy 2023/05/17 原:會傳入 a0k03,改都抓a0k04(抬頭),傳入前會先以GetTitleCustData函數抓客戶編號再傳入
    If bolIsNo = True Then
        stQ = "CU01='" & Left(stCu, 8) & "' And CU02='" & Mid(stCu, 9, 1) & "' "
    '    If bolIsNo = False Then
    '        stQ = "CU04='" & stCu & "' "
    '    End If
        'Modify by Amy 2023/03/01 +國內同業、解除對造、不得代理專利、不得代理商標
        'Modify By Sindy 2025/6/27 +or cu80='其他' or cu80='業務自行處理' or cu80='國內同業' or cu80='解除對造' or cu80='不得代理專利' or cu80='不得代理商標'
        '                          改抓常變數
        stQ = stQ & "and (cu80 is null or instr('" & 客戶及代理人可讀取的狀態 & "',cu80)>0) and cu02=0 "
        
        stQ = "Select cu01,cu02,cu04,cu112,cu23,cu30,cu31 From Customer Where " & stQ
        intQ = 1
        Set rsA = ClsLawReadRstMsg(intQ, stQ)
        '以[收據抬頭]抓客戶檔之CU04,若為客戶資料則抓中文地址cu23
        If intQ = 1 Then
            If Trim("" & rsA.Fields("cu23")) <> MsgText(601) Then
                GetCusAddr2 = Trim("" & rsA.Fields("cu112")) & " " & Trim("" & rsA.Fields("cu23"))
            End If
        End If
    Else
    '客戶檔無資料,則再抓[收據抬頭資料檔.營業地址]acc420.a4215
    'ElseIf bolIsNo = False Then
    'end 2023/05/17
        stQ = "Select * From Acc420 Where a4201='" & stCu & "'"
        If rsA.State = adStateOpen Then rsA.Close
         intQ = 1
        Set rsA = ClsLawReadRstMsg(intQ, stQ)
        If intQ = 1 Then
            If Trim("" & rsA.Fields("a4215")) <> MsgText(601) Then
                GetCusAddr2 = Trim("" & rsA.Fields("a4215"))
            End If
        End If
    End If
    rsA.Close
End Function

'設定賣方資料(Taie Information)
Private Sub SetSellerTag()
    Dim RsQ As New ADODB.Recordset
    Dim stQ As String, intQ As Integer
   
    stQ = "Select a0807,a0802,a0804 From Acc080 Where a0801='J' "
    If RsQ.State = adStateOpen Then RsQ.Close
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        strSellerID = "" & RsQ.Fields("a0807") '統編
        For i = 0 To RsQ.Fields.Count - 1
            If i = 2 Then
                strSellerAddr = "" & RsQ.Fields("a0804")
            Else
                strSellerTag = strSellerTag & _
                                        "<" & IIf(i = 0, "Identifier", "Name") & ">" & _
                                        "" & RsQ.Fields(i) & _
                                        "</" & IIf(i = 0, "Identifier", "Name") & ">" & vbCrLf
            End If
        Next i
        strSellerTag = "<Seller>" & vbCrLf & strSellerTag & "</Seller>" & vbCrLf
    End If
End Sub

'回傳買方資料Tag
Private Function GetBuyerTag(ByVal stCode As String) As String
    If stCode = "A04" Then
        GetBuyerTag = "<Buyer>" & vbCrLf & _
                            "<Identifier>" & stBuyerID & "</Identifier>" & vbCrLf & _
                            "<Name>" & stBuyerName & "</Name>" & vbCrLf & _
                            "<Address>" & stBuyerAddr & "</Address>" & vbCrLf & _
                            "<PersonInCharge>" & stBuyerName & "</PersonInCharge>" & vbCrLf & _
                            "</Buyer>" & vbCrLf
        Exit Function
    End If
    
End Function

'回傳銷退金額
Private Function GetWriteOffVal(ByVal stAxc02 As String, ByVal stA0j01 As String) As String
    Dim RsQ As New ADODB.Recordset
    Dim stQ As String, intQ As Integer
   
    stQ = "Select Nvl(sum(nvl(a1u07,0))+sum(nvl(a1u09,0)),0) A1uAmt From Acc1u0 " & _
             "Where a1u02='" & stAxc02 & "' and a1u03='" & stA0j01 & "' "
    If RsQ.State = adStateOpen Then RsQ.Close
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        GetWriteOffVal = Val("" & RsQ.Fields(0))
    End If
    RsQ.Close
End Function

Private Sub VarClear()
    stBuyerTag = ""
    stBuyerID = ""
    stBuyerName = ""
    stBuyerAddr = ""
    stCusNo = ""
End Sub

'
Private Function GetDisDetailTag(ByVal stA4601 As String, ByVal stA4301 As String, ByRef arrField() As Variant, ByRef stAmt As String) As String
    Dim RsQ As New ADODB.Recordset
    Dim stQ As String, intQ As Integer
    Dim intSeq As Integer, j As Integer
    Dim stTag As String, stTP As String
    Dim stTaxAmt As String, stTotAmt As String
                                
    GetDisDetailTag = ""
    
    stQ = "Select * From Acc460,Acc430 " & _
                "Where a4601='" & stA4601 & "' And  a4301='" & stA4301 & "' "
    intQ = 1: intSeq = 1
    Set RsQ = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        RsQ.MoveFirst
        Do While RsQ.EOF = False
            For j = LBound(arrField) To UBound(arrField)
                Select Case arrField(j)
                    Case "OriginalInvoiceDate" '原發票日期
                        stTP = Format(Val("" & RsQ.Fields("a4302")) + 19110000, "####-##-##")
                    Case "OriginalInvoiceNumber" '原發票編號
                        stTP = "" & RsQ.Fields("a4301")
                    Case "OriginalDescription" '品名
                        stTP = "" & RsQ.Fields("a4602")
                    Case "Quantity"
                        stTP = strQty
                    Case "UnitPrice"
                        stTP = "" & RsQ.Fields("a4604")
                    Case "Amount"
                        stTP = "" & RsQ.Fields("a4604")
                    Case "Tax"
                        stTP = "" & RsQ.Fields("a4605")
                    Case "AllowanceSequenceNumber"
                        stTP = intSeq
                    Case "TaxType"
                        stTP = strTaxType
                End Select
                stTag = stTag & "<" & UCase(arrField(j)) & ">" & stTP & "</" & UCase(arrField(j)) & ">" & vbCrLf
            Next j
            stTaxAmt = Val(stTaxAmt) + Val("" & RsQ.Fields("a4605"))
            stTotAmt = Val(stTotAmt) + Val("" & RsQ.Fields("a4604"))
            RsQ.MoveNext
            intSeq = intSeq + 1
            GetDisDetailTag = GetDisDetailTag & _
                                        "<PRODUCTITEM>" & vbCrLf & stTag & "</PRODUCTITEM>" & vbCrLf
            stTag = ""
        Loop
        '執行一次
        If stAmt = MsgText(601) Then
            stAmt = "<TAXAMOUNT>" & stTaxAmt & "</TAXAMOUNT>" & vbCrLf & _
                          "<TOTALAMOUNT>" & stTotAmt & "</TOTALAMOUNT>" & vbCrLf
        End If
    End If
    RsQ.Close
End Function

'特殊字取代為全型字
Private Function RepSpecWord(ByVal stData) As String
    RepSpecWord = Replace(Replace(Replace(Replace(stData, "<", "＜"), ">", "＞"), "&", "＆"), "'", "’")
End Function

'設定 C0401對應欄位
Private Sub SetC0401Field()
    Dim j As Integer
    'Modify by Amy 2024/10/29 盟立加欄,原:arrC04AF(1 To 30)/arrC04BF(1 To 12)
    ReDim arrC04AF(1 To 33): ReDim arrC04BF(1 To 13): ReDim arrC04CF(1 To 13)
    
    '=== Main Tag ===
    For j = LBound(arrC04AF) To UBound(arrC04AF)
        Select Case j
            Case 1
                arrC04AF(j) = "Invoice_Code"
            Case 2
                arrC04AF(j) = "InvoiceNumber"
            Case 3
                arrC04AF(j) = "InvoiceDate"
            Case 4
                arrC04AF(j) = "InvoiceTime"
            Case 5 '買方統一編號
                arrC04AF(j) = "BuyerIdentifier"
            Case 6
                arrC04AF(j) = "BuyerName"
            Case 7
                arrC04AF(j) = "BuyerAddress"
            Case 8 '買方負責人
                arrC04AF(j) = "BuyerPrincipal"
            Case 9
                arrC04AF(j) = "BuyerTEL"
            Case 10
                arrC04AF(j) = "BuyerFAX"
            Case 11
                arrC04AF(j) = "BuyerEmail"
            Case 12
                arrC04AF(j) = "BuyerCusNo"
            Case 13 '買方營業人角色註記
                arrC04AF(j) = "BuyerRole"
            Case 14 '發票檢查碼
                arrC04AF(j) = "CheckNumber"
            Case 15 '買受人註記
                arrC04AF(j) = "BuyerRemark"
            Case 16 '總備註
                arrC04AF(j) = "MainRemark"
            Case 17 '通關方式註記
                arrC04AF(j) = "CustomsClearanceMark"
            Case 18 '稅捐稽徵處名稱
                arrC04AF(j) = "TaxServiceName"
            Case 19 '核准日
                arrC04AF(j) = "AllowDate"
            Case 20 '核准文
                arrC04AF(j) = "AllowDoc"
            Case 21 '核准號
                arrC04AF(j) = "AllowNo"
            Case 22 '發票類別
                arrC04AF(j) = "InvoiceType"
            Case 23 '彙開註記
                arrC04AF(j) = "GroupMark"
            Case 24 '捐贈註記
                arrC04AF(j) = "DonateMark"
            Case 25 '載具類別號碼
                arrC04AF(j) = "CarrierType"
            Case 26 '載具顯碼ID
                arrC04AF(j) = "CarrierId1"
            Case 27 '載具隱碼ID
                arrC04AF(j) = "CarrierId2"
            Case 28 '紙本電子發票已列印註記
                arrC04AF(j) = "PrintMark"
            Case 29 '發票捐贈對象
                arrC04AF(j) = "NPOBAN"
            Case 30 '發票防偽隨機碼
                arrC04AF(j) = "InvoceRandomNo"
            'Add by Amy 2024/10/29
            Case 31 '相關號碼
                arrC04AF(j) = "RelateNumber"
            Case 32 '買受人簽署適用零稅率註記
                arrC04AF(j) = "BondedAreaConfirm"
            Case 33 '零稅率原因
                arrC04AF(j) = "ZeroTaxRateReason"
            'end 2024/10/29
        End Select
    Next j
    '=== End Main Tag ===
    
    '=== Detail Tag ====
    For j = LBound(arrC04BF) To UBound(arrC04BF)
        Select Case j
            Case 1 '商品項目
                arrC04BF(j) = "ProductItem"
            Case 2 '品名
                arrC04BF(j) = "Description"
            Case 3 '數量
                arrC04BF(j) = "Quantity"
            Case 4 '單位
                arrC04BF(j) = "Unit"
            Case 5 '單價
                arrC04BF(j) = "UnitPrice"
            Case 6 '金額
                arrC04BF(j) = "Amount"
            Case 7 '明細排列序號
                arrC04BF(j) = "SequenceNumber"
            Case 8 '單一欄位備註
                arrC04BF(j) = "Remark"
            Case 9 '相關號碼
                arrC04BF(j) = "RelateNumber"
            Case 10 '未稅金額
                arrC04BF(j) = "UnTax"
            Case 11 '品號
                arrC04BF(j) = "DescriptionNo"
            Case 12 '品項條碼
                arrC04BF(j) = "Item_Number"
            Case 13 '課稅別 'Add by Amy 2024/10/29
                arrC04BF(j) = "TaxType"
        End Select
    Next j
    '=== End Detail Tag ====
    
    '=== Amount Tag ===
    For j = LBound(arrC04CF) To UBound(arrC04CF)
        Select Case j
            Case 1 '銷售額合計
                arrC04CF(j) = "SalesAmount"
            Case 2 '免稅銷售額合計
                arrC04CF(j) = "FreeTaxSalesAmount"
            Case 3 '零稅率銷售額合計
                arrC04CF(j) = "ZeroTaxSalesAmount"
            Case 4 '課稅別
                arrC04CF(j) = "TaxType"
            Case 5 '稅率
                arrC04CF(j) = "TaxRate"
            Case 6 '營業稅額
                arrC04CF(j) = "TaxAmount"
            Case 7 '總計
                arrC04CF(j) = "TotalAmount"
            Case 8 '扣抵金額
                arrC04CF(j) = "DiscountAmount"
            Case 9 '原幣金額
                arrC04CF(j) = "OriginalCurrencyAmount"
            Case 10 '匯率
                arrC04CF(j) = "ExchangeRate"
            Case 11 '幣別
                arrC04CF(j) = "Currency"
            Case 12 '備註一(會於QCode下方顯示)
                arrC04CF(j) = "Remark1"
            Case 13 '備註二(會於明細聯顯示)
                arrC04CF(j) = "Remark2"
        End Select
    Next j
    '=== End Amount Tag ===
   
End Sub

Private Function GetFieldVal(stCodeTag As String, pFieldN As String) As Integer
    Dim jj As Integer
    
    'Main Tag
    If stCodeTag = "C04A" Then
        For jj = LBound(arrC04AF) To UBound(arrC04AF)
           If UCase(arrC04AF(jj)) = UCase(pFieldN) Then
              GetFieldVal = jj
              Exit For
           End If
        Next jj
        Exit Function
    End If
    
    'Detail Tag
    If stCodeTag = "C04B" Then
        For jj = LBound(arrC04BF) To UBound(arrC04BF)
           If UCase(arrC04BF(jj)) = UCase(pFieldN) Then
              GetFieldVal = jj
              Exit For
           End If
        Next jj
        Exit Function
    End If
    
    'Amount Tag
    If stCodeTag = "C04C" Then
        For jj = LBound(arrC04CF) To UBound(arrC04CF)
           If UCase(arrC04CF(jj)) = UCase(pFieldN) Then
              GetFieldVal = jj
              Exit For
           End If
        Next jj
        Exit Function
    End If
End Function

'參考PUB_PrintCaseReceipt_Inv修改
Private Function GetInvoiceDetail1(ByVal stA4308 As String) As String
    Dim stQ As String, intQ As Integer
    Dim stItemDesc As String, stCaseNo As String '品名/案號
    Dim stTmp As String
    Dim stTotAmt As String, stAmount1 As String, stAmount2 As String, stAmt As String
    Dim stSys As String
    Dim stText As String, stTag As String
    Dim intRow As Integer
    
    GetInvoiceDetail1 = ""
    stQ = "Select * From Acc0j0,CaseProgress,Nation,CasePropertyMap " & _
            "Where a0j13='" & adoMain.Fields("axc02") & "' And a0j01=cp09(+) " & _
            "And cp01=cpm01(+) And cp10=cpm02(+) " & _
            "And (cp79 <> 0 or (cp79 = 0 And cp75 <> 0))  And na01(+)=a0j04 " & _
            "Order by a0j25 asc"
    intQ = 1
    Set adoQ = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        adoQ.MoveFirst
        stTotAmt = 0:  intRow = 0: stAmount1 = 0: stAmount2 = 0
        stItemDesc = adoQ.Fields("a0j22") '帳款類別
        Do While adoQ.EOF = False
            If stItemDesc <> adoQ.Fields("a0j22") Then '帳款類別
                '申請國家+專利/商標/案件性質
                stAmount2 = Val(stAmount2) + Val(stTotAmt) 'stAmount2=總金額
                If bolBtoB = False Then
                    stText = stTotAmt
                Else
                    stText = Round((stTotAmt / 1.05))
                    '計算後可能會有差額1,2元;最後一筆要攤平 ex.E10309340
                    If intRow = adoQ.RecordCount Then
                        stText = Round((Val(stAmount2) / 1.05)) - Val(stAmount1)
                    End If
                End If
                stAmount1 = Val(stAmount1) + Val(stText) 'stAmount1=總淨額(扣掉稅額)
                intRow = intRow + 1
                If bolBtoB = True Then
                    'Modify by Amy 2021/06/25 +本所案號
                    stTag = stTag & GetInvoiceDetailTag_A04(Mid(stTmp, 2), stText, intRow, "" & adoQ.Fields("a0j02"))
                Else
                    'Modify by Amy 2019/08/21 每筆明細都要有<B><B1></B1>...</B>
                    stTag = stTag & "<B>" & GetInvoiceDetailTag_C04(Mid(stTmp, 2), stText, intRow) & "</B>"
                End If
                stTotAmt = 0: stCaseNo = "": stTmp = ""
            End If
            'Modify by Amy 2020/07/09 拿掉if
            'If stCaseNo = "" Then
            stCaseNo = "" & adoQ.Fields("a0j02") '本所案號
            'end 2020/07/09
                  
            '金額
            stAmt = Val("" & adoQ.Fields("a0j09")) + Val("" & adoQ.Fields("a0j10"))
            '扣除銷帳金額
            stAmt = stAmt - Val(GetWriteOffVal("" & adoMain.Fields("axc02"), "" & adoQ.Fields("a0j01")))
            stTotAmt = Val(stTotAmt) + Val(stAmt)
      
            stSys = CheckSys(Left(adoQ.Fields("a0j02"), Len(adoQ.Fields("a0j02")) - 9))
            'Mark by Amy 2020/07/10 拿掉if 都顯示-瑞婷
            'If stItemDesc <> adoQ.Fields("a0j22") Or stTmp = MsgText(601) Then
                'Modify by Amy 2024/09/03 案號000不顯示(同 PUB_PrintCaseReceipt_J_Doc 判斷)
                'Modify by Amy 2024/09/23 同定稿要有-(同 PUB_PrintCaseReceipt_J_Doc 判斷)
                stTmp = stTmp & "/" & GetPrjNationName("" & adoQ.Fields("a0j04")) & IIf(stSys = "1" Or stSys = "5", "專利", IIf(stSys = "2" Or stSys = "6", "商標", "服務費")) & _
                             "/" & adoQ.Fields("a0j22") '原:& "/" & IIf(Right(stCaseNo, 3) = "000", Mid(stCaseNo, 1, Len(stCaseNo) - 3), stCaseNo)
                strExc(9) = Left(stCaseNo, Len(stCaseNo) - 3)
                strExc(10) = Right(stCaseNo, 3)
                strExc(8) = Left(strExc(9), Len(strExc(9)) - 6) & "-" & Right(strExc(9), 6)
                If Right(strExc(10), 3) <> "000" Then
                  strExc(8) = strExc(8) & "-" & Left(strExc(10), 1) & "-" & Right(strExc(10), 2)
                End If
                stTmp = stTmp & "/" & strExc(8)
                'end 2024/09/23
            'End If
            
            stItemDesc = adoQ.Fields("a0j22") '帳款類別
            adoQ.MoveNext
        Loop
        '最後一筆
        stAmount2 = Val(stAmount2) + Val(stTotAmt) 'stAmount2=總金額
        If bolBtoB = False Then
            stText = stTotAmt
        Else
            stText = Round((stTotAmt / 1.05))
            '計算後可能會有差額1,2元;最後一筆要攤平 ex.E10309340
            If intRow = adoQ.RecordCount Then
                stText = Round((Val(stAmount2) / 1.05)) - Val(stAmount1)
            End If
        End If
        intRow = intRow + 1
        If bolBtoB = True Then
            'Modify by Amy 2021/06/25 +本所案號
            stTag = stTag & GetInvoiceDetailTag_A04(Mid(stTmp, 2), stText, intRow, stCaseNo)
        Else
            'Modify by Amy 2019/08/21 每筆明細都要有<B><B1></B1>...</B>
            stTag = stTag & "<B>" & GetInvoiceDetailTag_C04(Mid(stTmp, 2), stText, intRow) & "</B>"
        End If
        
    Else
        stText = ""
        If Val(stA4308) <> 0 Then
            stItemDesc = "作廢": intRow = 1
        End If
        If bolBtoB = True Then
            'Modify by Amy 2021/06/25 +本所案號
            stTag = GetInvoiceDetailTag_A04(stItemDesc, stText, intRow, stCaseNo)
        Else
            'Modify by Amy 2019/08/21 每筆明細都要有<B><B1></B1>...</B>
            stTag = "<B>" & GetInvoiceDetailTag_C04(stItemDesc, stText, intRow) & "</B>"
        End If
    End If
    GetInvoiceDetail1 = stTag
    adoQ.Close
End Function

'參考PUB_PrintCaseReceipt_Inv修改
Private Function GetInvoiceDetail2(ByVal stA4308 As String) As String
    Dim stQ As String, intQ As Integer
    Dim stAccClass As String, stItemDesc As String, stCaseNo As String '帳款類別/品名/案號
    Dim stTotAmt As String, stAmount1 As String, stAmount2 As String, stAmt As String
    Dim stSys As String
    Dim stText As String, stTag As String
    Dim intRow As Integer
    
    GetInvoiceDetail2 = ""
    'Modify by Amy 2022/06/21 +cp79及cp75條件,避免銷帳還出現一筆金額0 的資料  ex:E11112336
    'Sindy加PUB_PrintCaseReceipt_Inv函數時,程式沒加到cp79及cp75條件,目前PUB_PrintCaseReceipt_Inv函數無使用,故不改
    stQ = "Select * From Acc0j0,CaseProgress,CasePropertyMap " & _
                "Where a0j13='" & adoMain.Fields("axc02") & "' And a0j01=cp09(+) " & _
                "And cp01=cpm01(+) and cp10=cpm02(+) And (cp79 <> 0 or (cp79 = 0 And cp75 <> 0)) " & _
                "Order by a0j01 asc"
    intQ = 1
    Set adoQ = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        adoQ.MoveFirst
        Do While adoQ.EOF = False
            intRow = intRow + 1
            stCaseNo = "" & adoQ.Fields("a0j02") 'Add by Amy 2021/06/25
            stSys = CheckSys(Left(adoQ.Fields("a0j02"), Len(adoQ.Fields("a0j02")) - 9))
            '申請國家+專利/商標/案件性質/本所案號
            'Modify by Amy 2024/09/03 案號000不顯示(同 PUB_PrintCaseReceipt_J_Doc 判斷)
            'Modify by Amy 2024/09/23 同定稿要有-(同 PUB_PrintCaseReceipt_J_Doc 判斷)
            stItemDesc = GetPrjNationName("" & adoQ.Fields("a0j04")) & IIf(stSys = "1" Or stSys = "5", "專利", IIf(stSys = "2" Or stSys = "6", "商標", "服務費")) & "/" & IIf("" & adoQ.Fields("a0j04") = "020", adoQ.Fields("cpm04"), adoQ.Fields("cpm03"))
                                '原:& _"/" & IIf(Right(stCaseNo, 3) = "000", Mid(stCaseNo, 1, Len(stCaseNo) - 3), stCaseNo)
            strExc(9) = Left(stCaseNo, Len(stCaseNo) - 3)
            strExc(10) = Right(stCaseNo, 3)
            strExc(8) = Left(strExc(9), Len(strExc(9)) - 6) & "-" & Right(strExc(9), 6)
            If Right(strExc(10), 3) <> "000" Then
               strExc(8) = strExc(8) & "-" & Left(strExc(10), 1) & "-" & Right(strExc(10), 2)
            End If
            stItemDesc = stItemDesc & "/" & strExc(8)
            'end 2024/09/23
            
            '金額
            stAmt = Val("" & adoQ.Fields("a0j09")) + Val("" & adoQ.Fields("a0j10"))
            '扣除銷帳金額
            stAmt = Val(stAmt) - Val(GetWriteOffVal("" & adoMain.Fields("axc02"), "" & adoQ.Fields("a0j01")))
            stAmount2 = Val(stAmount2) + Val(stAmt) 'stAmount2=總金額
            If bolBtoB = False Then
                stText = stAmt
            Else
                stText = Round((Val(stAmt) / 1.05))
                '計算後可能會有差額1,2元;最後一筆要攤平 ex.E10309340
                If intRow = adoQ.RecordCount Then
                    stText = Round((Val(stAmount2) / 1.05)) - Val(stAmount1)
                End If
            End If
            stAmount1 = Val(stAmount1) + Val(Format(stText)) 'stAmount1=總淨額(扣掉稅額)
            If bolBtoB = True Then
                'Modify by Amy 2021/06/25 +本所案號
                stTag = stTag & GetInvoiceDetailTag_A04(stItemDesc, stText, intRow, stCaseNo)
            Else
                'Modify by Amy 2019/08/21 每筆明細都要有<B><B1></B1>...</B>
                stTag = stTag & "<B>" & GetInvoiceDetailTag_C04(stItemDesc, stText, intRow) & "</B>"
            End If
            
            adoQ.MoveNext
        Loop
    '當日未上傳前取消Acc0k0會沒資料,但明細Tag仍要產生
    Else
        stText = ""
        If Val(stA4308) <> 0 Then
            stItemDesc = "作廢": intRow = 1
        End If
        If bolBtoB = True Then
            'Modify by Amy 2021/06/25 +本所案號
            stTag = GetInvoiceDetailTag_A04(stItemDesc, stText, intRow, stCaseNo)
        Else
            'Modify by Amy 2019/08/21 每筆明細都要有<B><B1></B1>...</B>
            stTag = "<B>" & GetInvoiceDetailTag_C04(stItemDesc, stText, intRow) & "</B>"
        End If
    End If
    GetInvoiceDetail2 = stTag
    adoQ.Close
End Function

'A0401存證發票 Detail Tag
'Modify by Amy 2021/06/25 +stCaseNo
Private Function GetInvoiceDetailTag_A04(stItemDesc As String, stUnitPrice As String, intSeq As Integer, stCaseNo As String) As String
    Dim stTP As String
    GetInvoiceDetailTag_A04 = ""
    For i = LBound(arrA04DetF) To UBound(arrA04DetF)
        stTP = ""
        Select Case arrA04DetF(i)
            Case "Description"
                stTP = stItemDesc
            Case "Quantity"
                stTP = strQty
            Case "UnitPrice"
                stTP = Val(stUnitPrice)
            Case "Amount"
                stTP = Val(stUnitPrice)
            Case "SequenceNumber"
                stTP = intSeq
            'Add by Amy 2021/06/25 單一欄位備註(明細備註)顯示案件名稱
            Case "Remark"
                If stCaseNo <> MsgText(601) Then
                    'Modify by Amy 2022/08/29 由於ACS可能不顯示案件名稱,故改抓共用function
                    'stTP = CaseNameShow(Mid(stCaseNo, 1, Len(stCaseNo) - 9), Mid(stCaseNo, Len(stCaseNo) - 8, 6), Mid(stCaseNo, Len(stCaseNo) - 2, 1), Mid(stCaseNo, Len(stCaseNo) - 1, Len(stCaseNo)), 1)
                    stTP = Pub_GetInvRemark(Me.Name, stCaseNo, , False, "" & adoMain.Fields("a4326"))
                End If
            Case "TaxType" '課稅別 'Add by Amy 2024/11/04
                stTP = "1" '應稅
            '"RelateNumber"
            '相關號碼
        End Select
        GetInvoiceDetailTag_A04 = GetInvoiceDetailTag_A04 & _
                                                "<" & arrA04DetF(i) & ">" & stTP & "</" & arrA04DetF(i) & ">" & vbCrLf
    Next i
    GetInvoiceDetailTag_A04 = "<ProductItem>" & vbCrLf & _
                                                 GetInvoiceDetailTag_A04 & _
                                                 "</ProductItem>" & vbCrLf
End Function

'C0401開立發票 Detail Tag
Private Function GetInvoiceDetailTag_C04(stItemDesc As String, stAmt As String, intSeq As Integer) As String
    Dim stTP As String
    GetInvoiceDetailTag_C04 = ""
    'Modify by Amy 2024/10/29 盟立加欄位 原:12
    For i = 1 To 13
        stTP = ""
        Select Case i
            Case GetFieldVal("C04B", "ProductItem")
                stTP = intSeq
            Case GetFieldVal("C04B", "Description")
                stTP = stItemDesc
            Case GetFieldVal("C04B", "Quantity")
                stTP = strQty
            Case GetFieldVal("C04B", "UnitPrice")
                stTP = Val(stAmt)
            Case GetFieldVal("C04B", "Amount")
                stTP = Val(stAmt)
            Case GetFieldVal("C04B", "SequenceNumber")
                stTP = intSeq
            'Add by Amy 2024/10/29 '課稅別
            Case GetFieldVal("C04B", "TaxType")
               stTP = strTaxType
               If "" & adoMain.Fields("a4323") = "Y" Then stTP = "2" '零稅率代碼
            ''"Unit","Remark","RelateNumber"
            '單位,單一欄位備註,相關號碼
        End Select
        stTP = "<B" & i & ">" & stTP & "</B" & i & ">" & vbCrLf
        GetInvoiceDetailTag_C04 = GetInvoiceDetailTag_C04 & stTP
    Next i
End Function

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    If MaskEdBox1.Text = MsgText(29) Or MaskEdBox1.Text = MsgText(601) Then Exit Sub
    
    If IsDate(ChangeTStringToWDateString(FCDate(MaskEdBox1.Text))) = False Then
        MsgBox "上傳日期輸入錯誤！", , MsgText(5)
        Cancel = True
        MaskEdBox1.SetFocus
    End If
    
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
    If MaskEdBox2.Text = MsgText(29) Or MaskEdBox2.Text = MsgText(601) Then Exit Sub
    
    If IsDate(ChangeTStringToWDateString(FCDate(MaskEdBox2.Text))) = False Then
        MsgBox "上傳日期輸入錯誤！", , MsgText(5)
        Cancel = True
        MaskEdBox2.SetFocus
    End If
    If MaskEdBox1.Text = MsgText(29) Or MaskEdBox1.Text = MsgText(601) Then Exit Sub
    
    If Val(FCDate(MaskEdBox1.Text)) > Val(FCDate(MaskEdBox2.Text)) Then
        MsgBox "上傳日期迄日不可大於起日！", , MsgText(5)
        Cancel = True
        MaskEdBox2.SetFocus
    End If
End Sub

'預設今日目前第一筆和最後一筆發票號碼
Private Sub SetTodayInvoiceNo()
    Dim stQ As String, intQ As Integer
    
    stQ = "Select Min(a4301) as InvNo,1 as Sort From Acc430 Where a4302=" & Val(strSrvDate(2)) & _
    " Union Select Max(a4301) as InvNo,2 as Sort From Acc430 Where a4302=" & Val(strSrvDate(2)) & _
    " Order by Sort"
    intQ = 1
    Set adoQ = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        adoQ.MoveFirst
        Do While adoQ.EOF = False
            If Val("" & adoQ.Fields("Sort")) = 1 Then
                Text2 = "" & adoQ.Fields("InvNo")
            Else
                Text3 = "" & adoQ.Fields("InvNo")
            End If
            adoQ.MoveNext
        Loop
    End If
End Sub

Private Sub Text2_GotFocus()
    TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_GotFocus()
    TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub
