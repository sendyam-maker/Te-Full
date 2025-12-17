VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc14i0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "廠商付款明細表"
   ClientHeight    =   2460
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   5412
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5412
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
      Top             =   45
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
      Left            =   750
      Style           =   2  '單純下拉式
      TabIndex        =   12
      Top             =   2040
      Width           =   4600
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
      Left            =   1440
      TabIndex        =   1
      Top             =   420
      Width           =   1572
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
      Left            =   3360
      TabIndex        =   2
      Top             =   420
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生Excel檔(&P)"
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
      Left            =   390
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   1560
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1440
      TabIndex        =   3
      Top             =   780
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
      TabIndex        =   4
      Top             =   780
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
      Left            =   1425
      TabIndex        =   5
      Top             =   1140
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
      Left            =   480
      TabIndex        =   14
      Top             =   45
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "印表機"
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
      Left            =   50
      TabIndex        =   13
      Top             =   2070
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "匯款日期"
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
      Left            =   480
      TabIndex        =   11
      Top             =   1170
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "廠商編號"
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
      Left            =   480
      TabIndex        =   10
      Top             =   420
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
      Left            =   3120
      TabIndex        =   9
      Top             =   420
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "付款日期"
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
      Left            =   480
      TabIndex        =   8
      Top             =   780
      Width           =   975
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
      TabIndex        =   7
      Top             =   780
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc14i0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Const cTwipsPerCentiMeter As Integer = 567
Dim strPrinter As String
'Add by Amy 2014/02/18
Dim strCompName As String '公司名稱
Dim strA1p05 As String '會計科目
Dim StrA0H02 As String '銀行帳號
Dim strA0H03 As String '帳戶名稱
Dim stra0807 As String '統一編號
'end 2014/02/18
Dim strCmp As String 'Add by Amy 2020/04/13 公司編號

'Add by Amy 2020/04/13
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
        MsgBox Label2 & MsgText(63), , MsgText(5)
        Cancel = True
        Combo2.SetFocus
        Exit Sub
    ElseIf Len(Trim(Combo2)) = 1 Then
        Combo2 = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/04/13

Private Sub Command1_Click()
   'Add by Amy 2014/02/18 +公司別
   Dim bCancel As Boolean
   
   bCancel = False
   'Modify by Amy 2020/04/13 改下拉 原:Text5
   strCmp = ""
   If Trim(Combo2) = MsgText(601) Then
        MsgBox Label9 & "不可為空值！", , MsgText(5)
        Combo2.SetFocus
        Exit Sub
   End If
   Call Combo2_Validate(bCancel)
   If bCancel = True Then
        Combo2.SetFocus
        Exit Sub
   End If
   'end 2020/04/13
   'end 2014/02/18
   Screen.MousePointer = vbHourglass
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      
   Else
      If TxtValidate = True Then
         MsgBox "請更換新紙！", vbInformation
         'Moddify by Amy 2020/04/13 公司別改下拉 原:Text5
         If Trim(Combo2) <> MsgText(601) Then
            strCmp = Combo2
            If InStr(strCmp, "　") > 0 Then
                  strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
            End If
         End If
         'Add by Amy 2014/02/18
'         If Text5 = "1" Then
'            strCompName = "台一"
'         Else
'            strCompName = "智權"
'         End If
         strCompName = A0802Query(strCmp, True)
         'end 2014/02/18
         ProduceData
         FormClear
      End If
   End If
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub
'Add by Morgan 2011/6/22
Private Function TxtValidate() As Boolean
   If ChkDate(DBDATE(MaskEdBox3)) = False Then
      MaskEdBox3.SetFocus
      Exit Function
   Else
      If ChkWorkDay(DBDATE(MaskEdBox3)) = False Then
         MsgBox "匯款日期必須為工作日!!"
         MaskEdBox3.SetFocus
         Exit Function
      End If
   End If
   TxtValidate = True
End Function

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
   Dim stCon As String
   Dim BolOk As Boolean

   If Text2 <> "" Then
      stCon = stCon & " and a0q03>='" & Text2 & "'"
   End If
   If Text3 <> "" Then
      stCon = stCon & " and a0q03<='" & Text3 & "'"
   End If
   If MaskEdBox1.Text <> MsgText(29) Then
      stCon = stCon & " and a0q01>=" & Val(FCDate(Me.MaskEdBox1.Text))
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      stCon = stCon & " and a0q01<=" & Val(FCDate(Me.MaskEdBox2.Text))
   End If
   
   '瑞興電匯
   'Modify by Morgan 2007/2/5 廠商名稱不抓a0q05 改抓a0i17 or a0i02
   'Modify by Morgan 2007/2/6 所有匯費都由廠商付
   'strSQL = "select a0i15 c1,nvl(a0i17,a0i02) c2" & _
      ",a0q06-decode(substr(a0q03,1,1),'F',0,30) c3" & _
      ",decode(substr(a0q03,1,1),'F',30,0) c4" & _
      ",decode(substr(a0q03,1,1),'F',0,30) c5" & _
      ",a0q06+decode(substr(a0q03,1,1),'F',30,0) c6" & _
      " From acc0q0, acc0i0" & _
      " where a0i01=a0q03 and a0i12='1'" & stCon & _
      " and exists(select * from acc1p0 where a1p04=a0q17 and a1p05='110202')" & _
      " order by 1"
   'Modify by Amy 2014/02/18 語法+公司別及依公司別帶不同a1p05
   'strSql = "select a0i15 c1,nvl(a0i17,a0i02) c2" & _
      ",a0q06-30 c3,0 c4,30 c5,a0q06 c6,A0I14 c7,a0i20 c8" & _
      " From acc0q0, acc0i0" & _
      " where a0i01=a0q03 and a0i12='1'" & stCon & _
      " and exists(select * from acc1p0 where a1p04=a0q17 and a1p05='110202')" & _
      " order by 1"
   'Modify by Amy 2020/04/13 公司別改為下拉 原:Text5
   If strCmp = "1" Then
      'modify by sonia 2020/5/12 110202->110602
      strA1p05 = "110602"
   'Add by Amy 2020/04/14 +L
   ElseIf strCmp = "L" Then
      strA1p05 = "110502"
   Else
      strA1p05 = "110303"
   End If
   StrA0H02 = ""
   strA0H03 = GetAcc0H0Data(strA1p05, StrA0H02)
   
   strSql = "select a0i15 c1,nvl(a0i17,a0i02) c2" & _
      ",a0q06-30 c3,0 c4,30 c5,a0q06 c6,A0I14 c7,a0i20 c8" & _
      " From acc0q0, acc0i0" & _
      " where a0i01=a0q03 and a0i12='1' And a0q19='" & strCmp & "' " & stCon & _
      " and exists(select * from acc1p0 where a1p04=a0q17 And a1p01=a0q19 and a1p05='" & strA1p05 & "')" & _
      " order by 1"
   'end 2020/04/13
   'end 2014/02/18
   'end 2007/2/6
   
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, adoTaie, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         ExcelSave
         'Added by Morgan 2011/6/20 加轉媒體檔
         Save2File
         BolOk = True
      End If
   End With

   '瑞興直存
   'Modify by Morgan 2007/2/5 廠商名稱不抓a0q05 改抓a0i17 or a0i02
   'Modify by Amy 2014/02/18 語法+公司別及依公司別帶不同a1p05
   'Modify by Amy 2020/04/13 公司別改為下拉 原:Text5
   strSql = "select nvl(a0i17,a0i02) c1,a0i14 c2 ,a0q06 c3" & _
      " From acc0q0, acc0i0" & _
      " where a0i01=a0q03 and a0i12='2' And a0q19='" & strCmp & "' " & stCon & _
      " and exists(select * from acc1p0 where a1p04=a0q17 And a1p01=a0q19 and a1p05='" & strA1p05 & "')" & _
      " order by 2"
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, adoTaie, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         ExcelSave1
         BolOk = True
      End If
   End With
   
  'Added by Morgan 2011/11/15
  '華銀直存
  'Modify by Amy 2014/02/18 語法+公司別及依公司別帶不同a1p05
  'Modify by Amy 2020/04/13 公司別改為下拉 原:Text5
  If strCmp = "1" Then
        strA1p05 = "110207"
   'Add by Amy 2020/04/14 +L
   ElseIf strCmp = "L" Then
        strA1p05 = ""
   Else
        strA1p05 = "110304"
   End If
   StrA0H02 = ""
   strA0H03 = GetAcc0H0Data(strA1p05, StrA0H02)
   stra0807 = GetA0807(IIf(strCmp = "1", "2", strCmp)) '台一統一編號
   
    strSql = "select a0i15 c1,nvl(a0i17,a0i02) c2" & _
      ",a0q06 c3,0 c4,0 c5,a0q06 c6,A0I14 c7,a0i20 c8,a0i18 c9" & _
      " From acc0q0, acc0i0" & _
      " where a0i01=a0q03 and a0i12='1'  And a0q19='" & strCmp & "' and substr(a0i20,1,3)='008'" & stCon & _
      " and exists(select * from acc1p0 where a1p04=a0q17 And a1p01=a0q19 and a1p05='" & strA1p05 & "')" & _
      " order by 1"
    'end 2014/02/18
  'end 2020/04/13
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, adoTaie, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         Call Save2File2(1)
         Call ExcelSave2(1)
         BolOk = True
      End If
   End With
   
   '華銀電匯
   'Modify by Amy 2014/02/18 語法+公司別及依公司別帶不同a1p05
   'Modify by Amy 2020/04/13 公司別改為下拉 原:Text5
   strSql = "select a0i15 c1,nvl(a0i17,a0i02) c2" & _
      ",a0q06-25 c3,0 c4,25 c5,a0q06 c6,A0I14 c7,a0i20 c8" & _
      " From acc0q0, acc0i0" & _
      " where a0i01=a0q03 and a0i12='1' And a0q19='" & strCmp & "' and substr(a0i20,1,3)<>'008'" & stCon & _
      " and exists(select * from acc1p0 where a1p04=a0q17 And a1p01=a0q19 and a1p05='" & strA1p05 & "')" & _
      " order by 1"
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, adoTaie, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         Call Save2File2(2)
         Call ExcelSave2(2)
         BolOk = True
      End If
   End With
   'end 2011/11/15
      
   If BolOk Then
      MsgBox "資料已建立！"
   Else
      MsgBox "無符合資料！"
   End If
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
   Me.Width = 5505 'Modify by Amy 2015/06/11 原:5250
   Me.Height = 2945 'Modify by Amy 2014/02/18 原:2640
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Add by Amy 2020/04/13
   Combo2.AddItem "", 0
   Call Pub_SetCboCmp(Combo2, False, False, False, , 1)
   'end 2020/04/13
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   'Add by Morgan 2011/6/22
   MaskEdBox3 = Format(TransDate(CompWorkDay(1, CompDate(2, 1, strSrvDate(1))), 1), "0##/##/##")
   MaskEdBox3.Mask = DFormat
   'end 2011/6/22
   'Add by Amy 2015/06/11 +預設紙張在4100的第一匣處-瑞婷
   If Pub_StrUserSt03 = "M31" Then
        Combo1.AddItem "HP LaserJet 4100 Series PCL(財務處.第1匣)"
   End If
   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add by Morgan 2011/7/8
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   Set Frmacc14i0 = Nothing
End Sub

Private Sub MaskEdBox3_GotFocus()
   MaskEdBoxInverse MaskEdBox3
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
   Text2.SetFocus
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Private Function FormCheck() As Boolean
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
End Function

'*************************************************
'  轉成Excel檔案
'
'*************************************************
Private Sub ExcelSave()
   Dim strFilePath As String, strTitle As String
   Dim xlsSalesPoint As New Excel.Application
   Dim wksTmp As New Worksheet
   Dim lngCounter As Long
   
   'Excel檔案路徑
   'Modify by Amy 2014/02/18 檔案名稱+公司名稱
   strTitle = Me.Caption & strCompName & "(瑞興電匯)"
   strFilePath = strExcelPath & strTitle & ACDate(ServerDate) & ServerTime & MsgText(43)
   If Dir(strFilePath) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strFilePath
   End If
   
   xlsSalesPoint.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
   xlsSalesPoint.Workbooks.add
   Set wksTmp = xlsSalesPoint.Worksheets(1)
   With wksTmp
      '欄寬
      .PageSetup.Orientation = xlPortrait  '直印
      .PageSetup.PrintTitleRows = "$1:$5"
      .PageSetup.CenterFooter = "第 &P 頁，共 &N 頁"
      .Columns("a:a").ColumnWidth = 10
      .Columns("b:b").ColumnWidth = 30
      .Columns("c:c").ColumnWidth = 8
      .Columns("d:d").ColumnWidth = 10
      .Columns("e:e").ColumnWidth = 10
      .Columns("f:f").ColumnWidth = 10
      '表頭
      .Range("a1").Value = strTitle
      .Range("a1:f1").Select
       With .Range("a1:f1")
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
       End With
      '統計日期
      .Range("a3").Value = "付款日期："
      .Range("b3").Value = MaskEdBox1.Text & " － " & MaskEdBox2.Text
      
      .Range("a5").Value = "帳號"
      .Range("a5").HorizontalAlignment = xlCenter
      .Range("b5").Value = "客戶名稱"
      .Range("b5").HorizontalAlignment = xlCenter
      .Range("c5").Value = "金額"
      .Range("c5").HorizontalAlignment = xlCenter
      .Range("d5").Value = "匯費(所付)"
      .Range("d5").HorizontalAlignment = xlCenter
      .Range("e5").Value = "匯費(客付)"
      .Range("e5").HorizontalAlignment = xlCenter
      .Range("f5").Value = "合計"
      .Range("f5").HorizontalAlignment = xlCenter
      
      lngCounter = 5
      Do While Not adoRecordset.EOF
         lngCounter = lngCounter + 1
         .Range("a" & lngCounter).NumberFormatLocal = "@"
         .Range("a" & lngCounter).Value = "" & adoRecordset(0)
         .Range("b" & lngCounter).Value = "" & adoRecordset(1)
         .Range("c" & lngCounter).Value = "" & adoRecordset(2)
         .Range("d" & lngCounter).Value = "" & adoRecordset(3)
         .Range("e" & lngCounter).Value = "" & adoRecordset(4)
         .Range("f" & lngCounter).Value = "" & adoRecordset(5)
         adoRecordset.MoveNext
      Loop
      .Range("c" & lngCounter + 1).Formula = "=sum(c6:c" & lngCounter & ")"
      .Range("d" & lngCounter + 1).Formula = "=sum(d6:d" & lngCounter & ")"
      .Range("e" & lngCounter + 1).Formula = "=sum(e6:e" & lngCounter & ")"
      .Range("f" & lngCounter + 1).Formula = "=sum(f6:f" & lngCounter & ")"
      lngCounter = lngCounter + 1
      
      .Range("c6:f" & lngCounter).NumberFormatLocal = "#,##0"
      
      '加框線
      .Range("a5:f" & lngCounter).Select
      With xlsSalesPoint.Selection
         .Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Borders(xlEdgeTop).LineStyle = xlContinuous
         .Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Borders(xlEdgeRight).LineStyle = xlContinuous
         .Borders(xlInsideVertical).LineStyle = xlContinuous
         .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
   End With
   'Modify by Amy 2016/06/23 +判斷版本
   If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFilePath, FileFormat:=-4143
   Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFilePath, FileFormat:=56
   End If
   'end 2016/06/23
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   StatusClear
End Sub

'*************************************************
'  轉成Excel檔案
'
'*************************************************
Private Sub ExcelSave1()
   Dim strFilePath As String, strTitle As String
   Dim xlsSalesPoint As New Excel.Application
   Dim wksTmp As New Worksheet
   Dim lngCounter As Long
   
   'Excel檔案路徑
   'Modify by Amy 2014/02/18 檔案名稱+公司名稱
   strTitle = Me.Caption & strCompName & "(瑞興直存)"
   strFilePath = strExcelPath & strTitle & ACDate(ServerDate) & ServerTime & MsgText(43)
   If Dir(strFilePath) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strFilePath
   End If
   
   xlsSalesPoint.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
   xlsSalesPoint.Workbooks.add
   Set wksTmp = xlsSalesPoint.Worksheets(1)
   With wksTmp
      '欄寬
      .PageSetup.Orientation = xlPortrait  '直印
      .PageSetup.PrintTitleRows = "$1:$5"
      .PageSetup.CenterFooter = "第 &P 頁，共 &N 頁"
      .Columns("a:a").ColumnWidth = 30
      .Columns("b:b").ColumnWidth = 20
      .Columns("c:c").ColumnWidth = 15
      
      '表頭
      .Range("a1").Value = strTitle
      .Range("a1:c1").Select
       With .Range("a1:c1")
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
       End With
      '統計日期
      .Range("a3").Value = "付款日期：" & MaskEdBox1.Text & " － " & MaskEdBox2.Text
      
      .Range("a5").Value = "客戶名稱"
      .Range("a5").HorizontalAlignment = xlCenter
      .Range("b5").Value = "帳號"
      .Range("b5").HorizontalAlignment = xlCenter
      .Range("c5").Value = "金額"
      .Range("c5").HorizontalAlignment = xlCenter
      
      lngCounter = 5
      Do While Not adoRecordset.EOF
         lngCounter = lngCounter + 1
         .Range("a" & lngCounter).Value = "" & adoRecordset(0)
         .Range("b" & lngCounter).NumberFormatLocal = "@"
         .Range("b" & lngCounter).Value = "" & adoRecordset(1)
         .Range("c" & lngCounter).Value = "" & adoRecordset(2)
         adoRecordset.MoveNext
      Loop
      .Range("c" & lngCounter + 1).Formula = "=sum(c6:c" & lngCounter & ")"
      
      lngCounter = lngCounter + 1
      
      .Range("c6:c" & lngCounter).NumberFormatLocal = "#,##0"
      
      '加框線
      .Range("a5:c" & lngCounter).Select
      With xlsSalesPoint.Selection
         .Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Borders(xlEdgeTop).LineStyle = xlContinuous
         .Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Borders(xlEdgeRight).LineStyle = xlContinuous
         .Borders(xlInsideVertical).LineStyle = xlContinuous
         .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
   End With
   'Modify by Amy 2016/06/23 +判斷版本
   If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFilePath, FileFormat:=-4143
   Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFilePath, FileFormat:=56
   End If
   'end 2016/06/23
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   StatusClear
End Sub

'Added by Morgan 2011/12/9 華銀
'*************************************************
'  轉成Excel檔案
'
'*************************************************
Private Sub ExcelSave2(pChoice As Integer)
   Dim strFilePath As String, strTitle As String
   Dim xlsSalesPoint As New Excel.Application
   Dim wksTmp As New Worksheet
   Dim lngCounter As Long
   
   'Excel檔案路徑
   'Modify by Amy 2014/02/18 檔案路徑及檔案名稱+公司名稱
   If pChoice = 1 Then
      strTitle = Me.Caption & strCompName & "(華銀直存)"
   Else
      strTitle = Me.Caption & strCompName & "(華銀電匯)"
   End If
   strFilePath = strExcelPath & strTitle & ACDate(ServerDate) & ServerTime & MsgText(43)
   If Dir(strFilePath) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strFilePath
   End If
   
   xlsSalesPoint.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
   xlsSalesPoint.Workbooks.add
   Set wksTmp = xlsSalesPoint.Worksheets(1)
   With wksTmp
      '欄寬
      .PageSetup.Orientation = xlPortrait  '直印
      .PageSetup.PrintTitleRows = "$1:$5"
      .PageSetup.CenterFooter = "第 &P 頁，共 &N 頁"
      .Columns("a:a").ColumnWidth = 30
      .Columns("b:b").ColumnWidth = 20
      .Columns("c:c").ColumnWidth = 15
      
      '表頭
      .Range("a1").Value = strTitle
      .Range("a1:c1").Select
       With .Range("a1:c1")
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
       End With
      '統計日期
      .Range("a3").Value = "付款日期：" & MaskEdBox1.Text & " － " & MaskEdBox2.Text
      
      .Range("a5").Value = "客戶名稱"
      .Range("a5").HorizontalAlignment = xlCenter
      .Range("b5").Value = "帳號"
      .Range("b5").HorizontalAlignment = xlCenter
      .Range("c5").Value = "金額"
      .Range("c5").HorizontalAlignment = xlCenter
      
      lngCounter = 5
      adoRecordset.MoveFirst
      Do While Not adoRecordset.EOF
         lngCounter = lngCounter + 1
         .Range("a" & lngCounter).Value = "" & adoRecordset("c2")
         .Range("b" & lngCounter).NumberFormatLocal = "@"
         .Range("b" & lngCounter).Value = "" & adoRecordset("c7")
         .Range("c" & lngCounter).Value = "" & adoRecordset("c3")
         adoRecordset.MoveNext
      Loop
      .Range("c" & lngCounter + 1).Formula = "=sum(c6:c" & lngCounter & ")"
      
      lngCounter = lngCounter + 1
      
      .Range("c6:c" & lngCounter).NumberFormatLocal = "#,##0"
      
      '加框線
      .Range("a5:c" & lngCounter).Select
      With xlsSalesPoint.Selection
         .Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Borders(xlEdgeTop).LineStyle = xlContinuous
         .Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Borders(xlEdgeRight).LineStyle = xlContinuous
         .Borders(xlInsideVertical).LineStyle = xlContinuous
         .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
   End With
   'Modify by Amy 2016/06/23 +判斷版本
   If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFilePath, FileFormat:=-4143
   Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFilePath, FileFormat:=56
   End If
   'end 2016/06/23
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   StatusClear
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

'Add by Morgan 2011/6/20
'產生媒體檔案並列印遞送單及收妥回條
Private Sub Save2File()
   Dim strFileName As String
   Dim ff As Integer
   Dim strText As String
   Dim strTmp(9) As String
   Dim iCount As Integer, lAmount As Long, lNet As Long
   Dim strRDate As String '匯款日期
   
   'strFileName = PUB_Getdesktop & "\sedbh_pc"
   'Modify by Amy 2014/02/18 檔案名稱+公司名稱
   strFileName = strExcelPath & strCompName & "sedbh_pc"
   'end 2014/02/18
   strRDate = DBDATE(MaskEdBox3)
   
   If Dir(strFileName) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strFileName
   End If
   
   If ff > 0 Then Close #ff
   ff = FreeFile
   Open strFileName For Output As ff
   
   Erase strTmp
   strTmp(1) = "1" '區別碼 9(1) 首筆為 1
   strTmp(2) = strRDate '日期 9(8) 西元年月日
   strTmp(3) = String(279, Chr(32)) '空白 X(279)
   strText = ""
   For intI = 1 To 3
      strText = strText & strTmp(intI)
   Next
   Print #ff, strText
   
   With adoRecordset
   .MoveFirst
   iCount = 0
   lAmount = 0
   lNet = 0
   Do While Not .EOF
      iCount = iCount + 1
      lAmount = lAmount + Val("" & .Fields("c6"))
      lNet = lNet + Val("" & .Fields("c3"))
      Erase strTmp
      strTmp(1) = "2" '區別碼 9(1) 明細為 2
      strTmp(2) = strRDate '日期 9(8) 西元年月日
      strTmp(3) = Format(Val("" & .Fields("c8")), String(7, "0")) '解款行 9(7) 3位銀行代碼4位分行代碼
      strTmp(4) = Format(Val("" & .Fields("c7")), String(14, "0")) '收款人帳號 9(14) 右靠左補零
      strTmp(5) = Format(Val("" & .Fields("c3")), String(13, "0")) '金額 9(13) 右靠左補零
      'Modify by Amy 2014/02/18
      'strTmp(6) = Left("台一國際專利商標事務所" & String(40, "　"), 40) '匯款人名 X(80) 中英全形40字，左靠右補全形空白
      strTmp(6) = Left(strA0H03 & String(40, "　"), 40)
      'end 2014/02/18
      strTmp(7) = Left(toDblFont("" & .Fields("c2")) & String(40, "　"), 40) '收款人名 X(80) 中英全形40字，左靠右補全形空白
      strTmp(8) = String(40, "　") '附言 X(80) 中英數全形40字，左靠右補全形空白
      strTmp(9) = Format(Val("" & .Fields("c5")), String(5, "0")) '匯費 9(5) 右靠左補零
      strText = ""
      For intI = 1 To 9
         strText = strText & strTmp(intI)
      Next
      Print #ff, strText
      .MoveNext
   Loop
   
   Erase strTmp
   strTmp(1) = "3" '區別碼 9(1) 尾筆為 3
   strTmp(2) = strRDate '日期 9(8) 西元年月日
   strTmp(3) = Format(iCount, String(7, "0")) '筆數 9(7) 右靠左補零
   strTmp(4) = Format(lNet, String(13, "0")) '金額 9(13) 右靠左補零
   strTmp(5) = String(259, Chr(32)) '空白 X(259)
   strText = ""
   For intI = 1 To 5
      strText = strText & strTmp(intI)
   Next
   Print #ff, strText
   Close ff
   
   PUB_RestorePrinter Combo1
   PrintDeliveryForm lAmount, iCount, lNet
   PUB_RestorePrinter strPrinter
   
   End With
End Sub
'Add by Morgan 2011/6/20
'列印整批匯款遞送單及收妥回條
Private Sub PrintDeliveryForm(pAmount As Long, pCount As Integer, pNet As Long)
   Dim ii As Integer
   Dim stFontName As String, dblFontSize As Double
   Dim lngTopMargin As Long, lngLeftMargin As Long
   Dim Px As Long, Py As Long, pX2 As Long, pY2 As Long
   Dim stYear As String, stMonth As String, StDay As String '列印日期
   Dim stYear1 As String, stMonth1 As String, stDay1 As String '匯款日期
   Dim stTerm As String
   
   stYear = Format(PUB_DBYEAR(strSrvDate(1)) - 1911)
   stMonth = Format(PUB_DBMONTH(strSrvDate(1)))
   StDay = Format(PUB_DBDAY(strSrvDate(1)))
   
   stYear1 = Format(PUB_DBYEAR(MaskEdBox3) - 1911)
   stMonth1 = Format(PUB_DBMONTH(MaskEdBox3))
   stDay1 = Format(PUB_DBDAY(MaskEdBox3))
   
   stFontName = Printer.FontName
   dblFontSize = Printer.FontSize
   lngTopMargin = (Printer.Height - Printer.ScaleHeight) / 2
   lngLeftMargin = (Printer.Width - Printer.ScaleWidth) / 2 - 0.5 * cTwipsPerCentiMeter
   
   Printer.PaperSize = 9
   Printer.Orientation = 1
   Printer.FontName = "標楷體"
   Printer.Font.Size = 18
   stTerm = "瑞興銀行整批匯款遞送單"
   Px = 1.5 * cTwipsPerCentiMeter - lngLeftMargin + (17 * cTwipsPerCentiMeter - Printer.TextWidth(stTerm)) / 2
   Py = 1.5 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   Printer.Font.Size = 14
   
   Printer.DrawWidth = 5
   '框
   Px = 1.5 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 2.5 * cTwipsPerCentiMeter - lngTopMargin
   pX2 = Px + 17 * cTwipsPerCentiMeter
   pY2 = Py + 10.6 * cTwipsPerCentiMeter
   Printer.Line (Px, Py)-(pX2, pY2), , B
   
   '縱線
   Px = 6.2 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 2.5 * cTwipsPerCentiMeter - lngTopMargin
   pX2 = Px
   pY2 = Py + 4.8 * cTwipsPerCentiMeter
   Printer.Line (Px, Py)-(pX2, pY2)
   
   '橫線
   For ii = 1 To 4
      Px = 1.5 * cTwipsPerCentiMeter - lngLeftMargin
      Py = (2.5 + 1.2 * ii) * cTwipsPerCentiMeter - lngTopMargin
      pX2 = Px + 17 * cTwipsPerCentiMeter
      pY2 = Py
      Printer.Line (Px, Py)-(pX2, pY2)
   Next
   Printer.DrawWidth = 1
   
   Px = 1.8 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 2.8 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print "委 託 人 帳 號 ："
   
   Px = Px + 6.5 * cTwipsPerCentiMeter
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print "0075-21-" & StrA0H02 'Modify by Amy 2014/02/18 原: "0075-21-0149980"
   
   Px = 1.8 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 3.9 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print "匯  款  總  額 ："
   
   Px = Px + 6.5 * cTwipsPerCentiMeter
   Py = 3.9 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print "$" & Format(pAmount, "#,###,###,###") & "-"
   
   Px = 1.8 * cTwipsPerCentiMeter - lngLeftMargin
   Py = Py + 0.5 * cTwipsPerCentiMeter
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Font.Size = 9
   Printer.Print "(匯款金額加匯款手續費)"
   Printer.Font.Size = 14
   
   Px = 1.8 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 5.2 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print "匯  款  日  期 ："
   
   Px = Px + 6.5 * cTwipsPerCentiMeter
   Py = 5.2 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stYear1 & " 年 " & stMonth1 & " 月 " & stDay1 & " 日(不得為例假日)"
   
   Px = 1.8 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 6.4 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print "筆 數 / 金 額  ："
   
   Px = Px + 6.5 * cTwipsPerCentiMeter
   Py = 6.4 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print pCount & " 筆 " & Format(pNet, "#,###,###,###") & " 元"
   
   Px = 6.2 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 9 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print "委託人簽章：" & strA0H03 'Modify by Amy 2014/02/18 原:台一國際專利商標事務所
   
   Px = 16.2 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 9 * cTwipsPerCentiMeter - lngTopMargin + 60
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Font.Size = 9
   Printer.Print "(蓋原留印鑑)"
   Printer.Font.Size = 14

   
   Px = 1.8 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 12.5 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print "中    華    民    國   " & stYear & "   年   " & stMonth & "   月   " & StDay & "   日"
   
   Printer.DrawStyle = vbDot
   Px = 1 * cTwipsPerCentiMeter - lngLeftMargin
   Py = Printer.ScaleHeight / 2
   pX2 = Px + 18 * cTwipsPerCentiMeter
   pY2 = Py
   For intI = 1 To 10
      Printer.Line (Px, Py)-(pX2, pY2)
   Next
   Printer.DrawStyle = vbSolid
   
   lngTopMargin = -Printer.ScaleHeight / 2
   
   Printer.Font.Size = 18
   stTerm = "收  妥  回  條"
   Px = 1.5 * cTwipsPerCentiMeter - lngLeftMargin + (17 * cTwipsPerCentiMeter - Printer.TextWidth(stTerm)) / 2
   Py = 1.5 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   Printer.Font.Size = 14
   
   Printer.DrawWidth = 5
   '框
   Px = 1.5 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 2.5 * cTwipsPerCentiMeter - lngTopMargin
   pX2 = Px + 17 * cTwipsPerCentiMeter
   pY2 = Py + 4.8 * cTwipsPerCentiMeter
   Printer.Line (Px, Py)-(pX2, pY2), , B
   
   '縱線
   Px = 6.2 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 2.5 * cTwipsPerCentiMeter - lngTopMargin
   pX2 = Px
   pY2 = Py + 4.8 * cTwipsPerCentiMeter
   Printer.Line (Px, Py)-(pX2, pY2)
   
   '橫線
   For ii = 1 To 3
      Px = 1.5 * cTwipsPerCentiMeter - lngLeftMargin
      Py = (2.5 + 1.2 * ii) * cTwipsPerCentiMeter - lngTopMargin
      pX2 = Px + 17 * cTwipsPerCentiMeter
      pY2 = Py
      Printer.Line (Px, Py)-(pX2, pY2)
   Next
   Printer.DrawWidth = 1
   
   Px = 1.8 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 2.8 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print "委 託 人 帳 號 ："
   
   Px = Px + 6.5 * cTwipsPerCentiMeter
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print "0075-21-" & StrA0H02 'Modify by Amy 2014/02/18 原: "0075-21-0149980"
   
   Px = 1.8 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 3.9 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print "匯  款  總  額 ："
   
   Px = Px + 6.5 * cTwipsPerCentiMeter
   Py = 3.9 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print "$" & Format(pAmount, "#,###,###,###") & "-"
   
   Px = 1.8 * cTwipsPerCentiMeter - lngLeftMargin
   Py = Py + 0.5 * cTwipsPerCentiMeter
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Font.Size = 8
   Printer.Print "(匯款金額加匯款手續費)"
   Printer.Font.Size = 14
   
   Px = 1.8 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 5.2 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print "匯  款  日  期 ："
   
   Px = Px + 6.5 * cTwipsPerCentiMeter
   Py = 5.2 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stYear1 & " 年 " & stMonth1 & " 月 " & stDay1 & " 日(不得為例假日)"
   
   Px = 1.8 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 6.4 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print "筆 數 / 金 額  ："
   
   Px = Px + 6.5 * cTwipsPerCentiMeter
   Py = 6.4 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print pCount & " 筆 " & Format(pNet, "#,###,###,###") & " 元"
   
   Px = 1.8 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 9 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print "瑞興銀行         分行       作業主管：          經辦："
   
   Px = 1.8 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 11 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print "中    華    民    國   " & stYear & "   年   " & stMonth & "   月   " & StDay & "   日"
   
   Printer.EndDoc
   Printer.FontName = stFontName
   Printer.FontSize = dblFontSize
End Sub

'Add by Morgan 2011/11/15
'產生媒體檔案並列印轉帳登錄單
'華銀
'iChoice:1=直存,2=電匯
Private Sub Save2File2(iChoice As Integer)
   Dim strFileName As String
   Dim ff As Integer
   Dim strText As String
   Dim strTmp(11) As String
   Dim iCount As Integer, lAmount As Long, lNet As Long
   Dim strRDate As String '匯款日期
   
   'Modify by Amy 2014/02/18 檔案名稱+公司名稱
   '直存
   If iChoice = 1 Then
      strFileName = strExcelPath & strCompName & "HNCBIN.TXT"
   '電匯
   Else
      strFileName = strExcelPath & strCompName & "HNCBRMNEW.DAT"
   End If
   'end 2018/02/18
   
   If Dir(strFileName) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strFileName
   End If
   
   strRDate = DBDATE(MaskEdBox3)
   
   If ff > 0 Then Close #ff
   ff = FreeFile
   Open strFileName For Output As ff
   
   With adoRecordset
   .MoveFirst
   iCount = 0
   lAmount = 0
   lNet = 0
   Do While Not .EOF
      iCount = iCount + 1
      lAmount = lAmount + Val("" & .Fields("c6"))
      lNet = lNet + Val("" & .Fields("c3"))
      Erase strTmp
      
      '直存
      If iChoice = 1 Then
         strTmp(1) = strRDate '轉帳日期 X(08)
         strTmp(2) = "1450" '主辦行 X(04) 分行代碼
         'Modify by Amy 2014/02/18 原: "04146457"
         strTmp(3) = Left(stra0807 & String(10, " "), 10) '營利編號 X(10)
         strTmp(4) = Right(String(16, "0") & .Fields("c7"), 16)  '帳號 9(16) 右靠左補零
         strTmp(5) = String(10, " ") '身分證 X(10) 英文要大寫 '可不填
         strTmp(6) = "106" '轉帳別 X(03) 固定為 106 1=入帳,06=當日傳送當日轉帳
         strTmp(7) = Format(Val("" & .Fields("c3")), String(11, "0")) & "00" '金額 9(11)V99 右靠左補零,含兩位小數
         strTmp(8) = String(10, " ") '資料編號 X(10) 空白
         strTmp(9) = String(12, " ") '保留 X(12) 空白
         strTmp(10) = "Y" '是否須檢核戶名 X(01) 空白:不檢核,Y:須檢核
         strTmp(11) = Left(toDblFont(UCase("" & .Fields("c2"))) & String(20, "　"), 20) '帳號戶名 X(40) 全形字
      '電匯
      Else
         strTmp(1) = strRDate '匯出日期 X(08)
         strTmp(2) = "1450" '主辦行 X(04) 分行代碼
         'Modify by Amy 2014/02/18 原: "04146457"
         strTmp(3) = Left(stra0807 & String(10, " "), 10)  '公司代號 X(10) 左靠右補空白 '2014/02/18
         strTmp(4) = Format(Val("" & .Fields("c8")), String(7, "0")) '解款行 9(7)
         strTmp(5) = Left(.Fields("c7") & String(14, " "), 14)  '收款人帳號 X(14) 左靠右補空白
         strTmp(6) = Left(toDblFont(UCase("" & .Fields("c2"))) & String(39, "　"), 39) '收款人號戶名 X(78) 全形字
         strTmp(7) = Format(Val("" & .Fields("c3")), String(12, "0")) & "00" '金額 9(12)V99 右靠左補零,含兩位小數
         'Modify by Amy  2014/02/18 原:"台一國際專利法律事務所"
         strTmp(8) = Left(strA0H03 & String(39, "　"), 39)  '匯款人 X(79) '2014/02/18
         strTmp(9) = String(39, "　") '附言 X(79)
         strTmp(10) = String(20, " ") '資料編號 X(20)
      End If
      
      strText = ""
      For intI = 1 To 11
         strText = strText & strTmp(intI)
      Next
      Print #ff, strText
      .MoveNext
   Loop
   
   Close ff
   
'
   PUB_RestorePrinter Combo1
   PrintDeliveryForm2 iChoice, lAmount, iCount, lNet
   PUB_RestorePrinter strPrinter
   
   End With
End Sub
'Added by Morgan 2011/11/15
'華南銀行電腦集中轉帳登錄單
Private Sub PrintDeliveryForm2(iChoice As Integer, pAmount As Long, pCount As Integer, pNet As Long)
   Dim ii As Integer
   Dim stFontName As String, dblFontSize As Double
   Dim lngTopMargin As Long, lngLeftMargin As Long
   Dim Px As Long, Py As Long, pX2 As Long, pY2 As Long
   Dim stYear As String, stMonth As String, StDay As String '列印日期
   Dim stYear1 As String, stMonth1 As String, stDay1 As String '匯款日期
   Dim stTerm As String
   
   stYear = Format(PUB_DBYEAR(strSrvDate(1)) - 1911)
   stMonth = Format(PUB_DBMONTH(strSrvDate(1)))
   StDay = Format(PUB_DBDAY(strSrvDate(1)))
   
   stYear1 = Format(PUB_DBYEAR(MaskEdBox3) - 1911)
   stMonth1 = Format(PUB_DBMONTH(MaskEdBox3))
   stDay1 = Format(PUB_DBDAY(MaskEdBox3))
   
   stFontName = Printer.FontName
   dblFontSize = Printer.FontSize
   lngTopMargin = (Printer.Height - Printer.ScaleHeight) / 2
   lngLeftMargin = (Printer.Width - Printer.ScaleWidth) / 2
   
   Printer.PaperSize = 9
   Printer.Orientation = 1
   Printer.FontName = "新細明體"
   Printer.Font.Size = 18
   Printer.Font.Bold = True
   
   stTerm = "華南商業銀行 電腦集中轉帳登錄單"
   Px = 1.5 * cTwipsPerCentiMeter - lngLeftMargin + (18 * cTwipsPerCentiMeter - Printer.TextWidth(stTerm)) / 2
   Py = 1.2 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
      
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   stTerm = "中華民國  " & stYear1 & "  年  " & stMonth1 & "  月  " & stDay1 & "  日"
   Px = 1.5 * cTwipsPerCentiMeter - lngLeftMargin + (18 * cTwipsPerCentiMeter - Printer.TextWidth(stTerm)) / 2
   Py = 2.2 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   Printer.DrawWidth = 5
   '框1
   Px = 1.5 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 3 * cTwipsPerCentiMeter - lngTopMargin
   pX2 = Px + 18 * cTwipsPerCentiMeter
   pY2 = Py + 13 * cTwipsPerCentiMeter
   Printer.Line (Px, Py)-(pX2, pY2), , B
   
   '縱線1
   Px = 3 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 3 * cTwipsPerCentiMeter - lngTopMargin
   pX2 = Px
   pY2 = Py + 13 * cTwipsPerCentiMeter
   Printer.Line (Px, Py)-(pX2, pY2)
   
   '橫線1
   Px = 3 * cTwipsPerCentiMeter - lngLeftMargin
   Py = Py + 1.5 * cTwipsPerCentiMeter
   pX2 = Px + 16.5 * cTwipsPerCentiMeter
   pY2 = Py
   Printer.Line (Px, Py)-(pX2, pY2)
   
   '橫線2
   Px = 3 * cTwipsPerCentiMeter - lngLeftMargin
   Py = Py + 1.2 * cTwipsPerCentiMeter
   pX2 = Px + 16.5 * cTwipsPerCentiMeter
   pY2 = Py
   Printer.Line (Px, Py)-(pX2, pY2)
   
   '橫線3
   Px = 3 * cTwipsPerCentiMeter - lngLeftMargin
   Py = Py + 3.8 * cTwipsPerCentiMeter
   pX2 = Px + 16.5 * cTwipsPerCentiMeter
   pY2 = Py
   Printer.Line (Px, Py)-(pX2, pY2)
   
   '橫線4-7
   For ii = 4 To 7
      Px = 3 * cTwipsPerCentiMeter - lngLeftMargin
      Py = Py + 1.3 * cTwipsPerCentiMeter
      pX2 = Px + 16.5 * cTwipsPerCentiMeter
      pY2 = Py
      Printer.Line (Px, Py)-(pX2, pY2)
   Next
   
   '縱線2
   Px = 11 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 12.1 * cTwipsPerCentiMeter - lngTopMargin
   pX2 = Px
   pY2 = Py + 1.3 * cTwipsPerCentiMeter
   Printer.Line (Px, Py)-(pX2, pY2)
   
   '框2
   Px = 1.5 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 16.5 * cTwipsPerCentiMeter - lngTopMargin
   pX2 = Px + 18 * cTwipsPerCentiMeter
   pY2 = Py + 5 * cTwipsPerCentiMeter
   Printer.Line (Px, Py)-(pX2, pY2), , B
   
   '框3
   Px = 1.5 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 22 * cTwipsPerCentiMeter - lngTopMargin
   pX2 = Px + 18 * cTwipsPerCentiMeter
   pY2 = Py + 3 * cTwipsPerCentiMeter
   Printer.Line (Px, Py)-(pX2, pY2), , B
   
   '橫線1
   Px = 1.5 * cTwipsPerCentiMeter - lngLeftMargin
   Py = Py + 0.5 * cTwipsPerCentiMeter
   pX2 = Px + 18 * cTwipsPerCentiMeter
   pY2 = Py
   Printer.Line (Px, Py)-(pX2, pY2)
   
   '縱線1
   Px = Px + 1.5 * cTwipsPerCentiMeter
   Py = 22 * cTwipsPerCentiMeter - lngTopMargin
   pX2 = Px
   pY2 = Py + 0.5 * cTwipsPerCentiMeter
   Printer.Line (Px, Py)-(pX2, pY2)
   
   '縱線2
   Px = Px + 2.8 * cTwipsPerCentiMeter
   Py = 22 * cTwipsPerCentiMeter - lngTopMargin
   pX2 = Px
   pY2 = Py + 0.5 * cTwipsPerCentiMeter
   Printer.Line (Px, Py)-(pX2, pY2)
   
   '縱線3
   Px = Px + 1.3 * cTwipsPerCentiMeter
   Py = 22 * cTwipsPerCentiMeter - lngTopMargin
   pX2 = Px
   pY2 = Py + 0.5 * cTwipsPerCentiMeter
   Printer.Line (Px, Py)-(pX2, pY2)
   
   '縱線4
   Px = Px + 2.1 * cTwipsPerCentiMeter
   Py = 22 * cTwipsPerCentiMeter - lngTopMargin
   pX2 = Px
   pY2 = Py + 0.5 * cTwipsPerCentiMeter
   Printer.Line (Px, Py)-(pX2, pY2)
   
   '縱線5
   Px = Px + 2.2 * cTwipsPerCentiMeter
   Py = 22 * cTwipsPerCentiMeter - lngTopMargin
   pX2 = Px
   pY2 = Py + 0.5 * cTwipsPerCentiMeter
   Printer.Line (Px, Py)-(pX2, pY2)
   
   '縱線6
   Px = Px + 2.7 * cTwipsPerCentiMeter
   Py = 22 * cTwipsPerCentiMeter - lngTopMargin
   pX2 = Px
   pY2 = Py + 0.5 * cTwipsPerCentiMeter
   Printer.Line (Px, Py)-(pX2, pY2)
   
   Printer.DrawWidth = 1
   
   Printer.Font.Size = 17
   stTerm = "申請單位填寫"
   Px = 1.5 * cTwipsPerCentiMeter - lngLeftMargin + (1.5 * cTwipsPerCentiMeter - Printer.TextWidth("申")) / 2
   Py = 2 * cTwipsPerCentiMeter - lngTopMargin
   For ii = 1 To Len(stTerm)
      Py = Py + 1.4 * cTwipsPerCentiMeter + Printer.TextHeight(Mid(stTerm, ii, 1))
      Printer.CurrentX = Px: Printer.CurrentY = Py
      Printer.Print Mid(stTerm, ii, 1)
   Next

   Printer.Font.Size = 13
   stTerm = "轉帳公司代號  (營利事業登記證統一編號)"
   Px = 3.5 * cTwipsPerCentiMeter
   Py = 3.1 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   Printer.Font.Size = 15
   stTerm = stra0807 'Modify by Amy 2014/02/18 原:"04146457"
   Px = 3.5 * cTwipsPerCentiMeter
   Py = 3.8 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   Printer.Font.Size = 12
   If iChoice = 1 Then
      stTerm = "轉帳性質：  0.□扣帳 (借方)      1.▓入帳 (貸方)      □整批跨行通匯"
   Else
      stTerm = "轉帳性質：  0.□扣帳 (借方)      1.□入帳 (貸方)      ▓整批跨行通匯"
   End If
   Px = 3.5 * cTwipsPerCentiMeter
   Py = 4.9 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   If iChoice = 1 Then
      stTerm = "轉帳業務： 06.▓當日傳送當日轉帳"
   Else
      stTerm = "轉帳業務："
   End If
   Px = 3.5 * cTwipsPerCentiMeter
   Py = 6.1 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
      
   stTerm = "轉帳日期： 民國  " & stYear1 & "  年  " & stMonth1 & "  月  " & stDay1 & "  日"
   Px = 3.5 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 9.9 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   stTerm = "金　　額： 新台幣 " & Replace(ChangeNumber(str(pNet)), "整", "")
   Px = 3.5 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 11.2 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   stTerm = "筆　　數： " & pCount & " 筆"
   Px = 3.5 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 12.5 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   stTerm = "NT$： " & Format(pNet, "#,###,###,###")
   Px = 11.5 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 12.5 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   stTerm = "轉帳公司帳號： 14510" & StrA0H02  'Modify by Amy 原: 145100202330
   Px = 3.5 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 13.8 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   If iChoice = 2 Then
      stTerm = "備註： 手續費 " & Format(pCount * 25, "#,###,###,###") & " 元"
      Px = 3.5 * cTwipsPerCentiMeter - lngLeftMargin
      Py = 15.1 * cTwipsPerCentiMeter - lngTopMargin
      Printer.CurrentX = Px: Printer.CurrentY = Py
      Printer.Print stTerm
   End If
   
   stTerm = "磁片檢驗情形："
   Px = 3.5 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 16.9 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   stTerm = "(蓋      章)"
   Px = 15.5 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 20.5 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   Printer.Font.Size = 10
   stTerm = "主辦行"
   Px = 1.7 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 22.1 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   stTerm = "公司代號"
   Px = 3.7 * cTwipsPerCentiMeter - lngLeftMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   stTerm = "業種"
   Px = 6.1 * cTwipsPerCentiMeter - lngLeftMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   stTerm = "轉帳日"
   Px = 7.7 * cTwipsPerCentiMeter - lngLeftMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   stTerm = "片  數"
   Px = 9.9 * cTwipsPerCentiMeter - lngLeftMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   stTerm = "金　額"
   Px = 12.1 * cTwipsPerCentiMeter - lngLeftMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   stTerm = "筆數"
   Px = 14.4 * cTwipsPerCentiMeter - lngLeftMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   Printer.Font.Size = 12
   stTerm = "經副襄理"
   Px = 1.6 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 25.2 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   stTerm = "經辦員"
   Px = 8.2 * cTwipsPerCentiMeter - lngLeftMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   stTerm = "發訊員"
   Px = 14.2 * cTwipsPerCentiMeter - lngLeftMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   stTerm = "初碼值："
   Px = 1.6 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 25.8 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   stTerm = "譯碼人："
   Px = 8.2 * cTwipsPerCentiMeter - lngLeftMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   stTerm = "押碼值："
   Px = 1.6 * cTwipsPerCentiMeter - lngLeftMargin
   Py = 26.4 * cTwipsPerCentiMeter - lngTopMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   stTerm = "押碼人："
   Px = 8.2 * cTwipsPerCentiMeter - lngLeftMargin
   Printer.CurrentX = Px: Printer.CurrentY = Py
   Printer.Print stTerm
   
   Printer.EndDoc
   Printer.FontName = stFontName
   Printer.FontSize = dblFontSize
End Sub

'Mark by Amy 2020/04/13 公司別改下拉
'Add by Amy 2014/02/18
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
'        KeyAscii = UpperCase(KeyAscii)
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
'end 2020/04/13

Private Function GetAcc0H0Data(ByVal strA0H08 As String, ByRef stA0H02 As String) As String
    Dim adoacc0h0 As New ADODB.Recordset
    Dim strQuery As String
    
    GetAcc0H0Data = ""
    strQuery = "Select * From Acc0H0 Where a0H08='" & strA0H08 & "' "
    intI = 1
    Set adoacc0h0 = ClsLawReadRstMsg(intI, strQuery)
    If intI = 1 Then
         If Not IsNull(adoacc0h0.Fields("a0h03").Value) Then
            GetAcc0H0Data = adoacc0h0.Fields("a0h03").Value
            StrA0H02 = adoacc0h0.Fields("a0h02").Value
         End If
    End If
End Function

Private Function GetA0807(ByVal strA0801 As String) As String
    Dim adoacc080 As New ADODB.Recordset
    Dim strQuery As String
    
    GetA0807 = ""
    strQuery = "Select * From Acc080 Where a0801='" & strA0801 & "' "
    intI = 1
    Set adoacc080 = ClsLawReadRstMsg(intI, strQuery)
    If intI = 1 Then
         If Not IsNull(adoacc080.Fields("a0807").Value) Then
            GetA0807 = adoacc080.Fields("a0807").Value
         End If
    End If
    
End Function
'end 2014/02/18
