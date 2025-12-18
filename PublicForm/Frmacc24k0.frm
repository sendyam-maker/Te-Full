VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc24k0 
   AutoRedraw      =   -1  'True
   Caption         =   "國外請款點數分析表"
   ClientHeight    =   3312
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5124
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3312
   ScaleWidth      =   5124
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Frmacc24k0.frx":0000
      Left            =   1305
      List            =   "Frmacc24k0.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   3
      Top             =   1020
      Width           =   3525
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Frmacc24k0.frx":0004
      Left            =   1305
      List            =   "Frmacc24k0.frx":0006
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   1470
      Width           =   3525
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1305
      TabIndex        =   6
      Top             =   2370
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      Default         =   -1  'True
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
      TabIndex        =   5
      Top             =   1890
      Width           =   4692
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   0
      Text            =   "ALL"
      Top             =   240
      Width           =   3510
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   630
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
      Top             =   630
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "PS：僅統計分配點數歸該部門的資料！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   13
      Top             =   2880
      Width           =   3330
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "統計部門："
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
      Top             =   1050
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "統計方式："
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
      TabIndex        =   11
      Top             =   1470
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   " 印表機："
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
      Left            =   300
      TabIndex        =   10
      Top             =   2370
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   630
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "請款日期："
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
      TabIndex        =   8
      Top             =   630
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "系統類別："
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
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc24k0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Create by Morgan 2010/8/24
Option Explicit

Dim strOPrinter As String
Dim prnstrPos As Integer
Dim PLeft() As Integer, ColName() As String, iPrint As Integer, iPage As Integer
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long

Sub GetPleft()
   ReDim PLeft(0 To 7)
   ReDim ColName(1 To 7)
   PLeft(0) = 500
   '2011/12/22 modify by sonia
   'PLeft(1) = 500: ColName(1) = "組別"
   'PLeft(2) = PLeft(1) + 1000: ColName(2) = "系統別"
   If Combo2.ListIndex = 2 Then
      PLeft(1) = 500: ColName(1) = "國籍"
   Else
      PLeft(1) = 500: ColName(1) = "組別"
   End If
   PLeft(2) = PLeft(1) + 1500: ColName(2) = "系統別"
   '2011/12/22 end
   PLeft(3) = PLeft(2) + 1000: ColName(3) = "案件性質"
   PLeft(4) = PLeft(3) + 2000: ColName(4) = "件數"
   PLeft(5) = PLeft(4) + 1000: ColName(5) = "點數"
   PLeft(6) = PLeft(5) + 1400: ColName(6) = "百分比"
   PLeft(7) = PLeft(6) + 1300: ColName(7) = "備註"
End Sub

Private Sub Combo3_Click()
   SetCombo2
End Sub

Private Sub Command2_Click()
   If FormCheck = False Then
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   If ProduceData Then
      PUB_RestorePrinter Combo1
      PrintData
      'FormClear
      PUB_RestorePrinter strOPrinter
   End If
   Screen.MousePointer = vbDefault
   SetStatusBar
End Sub

Private Sub Form_Activate()
   If Forms(0).Name = "mdiMain" Then
      Forms(0).ToolShow
   End If
End Sub

Private Sub SetStatusBar()
   StatusView "請更換 A4 紙張"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      SetStatusBar
   End If
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, 5310, 3720
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   SetCombo3
   If Left(Pub_StrUserSt03, 2) = "F2" Then
      Combo3.ListIndex = 0
      Combo3.Enabled = False
   ElseIf Left(Pub_StrUserSt03, 2) = "F1" Then
      Combo3.ListIndex = 1
      Combo3.Enabled = False
   ElseIf Left(Pub_StrUserSt03, 2) = "F4" Then
      Combo3.ListIndex = 2
      Combo3.Enabled = False
   Else
      Combo3.ListIndex = 0
   End If
   PUB_SetPrinter Me.Name, Combo1, strOPrinter
   SetStatusBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc24k0 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Me.Text1.Text <> "ALL" Then
      If Not CheckSysKind1("" & Me.Text1.Text) Then
         Me.Text1.SetFocus
         Cancel = True
      End If
   End If
   If Cancel Then Text1_GotFocus
End Sub
'檢查輸入的系統類別是否正確
Private Function CheckSysKind1(strSysKind As String) As Boolean
Dim arr1
Dim arr2
Dim ii As Integer
Dim jj As Integer
   
   CheckSysKind1 = False
   arr2 = Split(Me.Text1.Text, ",")
   For ii = LBound(arr2) To UBound(arr2)
      If CheckSys(arr2(ii)) = "" Then
         MsgBox "系統類別輸入錯誤!!!", vbExclamation + vbOKOnly
         Exit Function
      End If
   Next ii
   CheckSysKind1 = True
End Function

'*************************************************
'  產生報表資料
'
'*************************************************
Private Function ProduceData() As Boolean
   Dim stSQL As String, stVTB As String, stCon As String, ii As Integer, strSystemKind As String
   Dim arr1
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/23 清除查詢印表記錄檔欄位
   stCon = ""
   If Text1 <> "ALL" Then
      arr1 = Split(Me.Text1.Text, ",")
      For ii = LBound(arr1) To UBound(arr1)
         strSystemKind = strSystemKind & "'" & arr1(ii) & "',"
      Next ii
      strSystemKind = Left(strSystemKind, Len(strSystemKind) - 1)
      stCon = stCon & " AND A1K13 IN ( " & strSystemKind & " ) "
   End If
   If Text1 <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1(0) & Text1 'Add By Sindy 2010/12/23
   End If
   
   If MaskEdBox1 <> "___/__/__" Then
      stCon = stCon & " and a1k02 >= " & ChangeTDateStringToTString(MaskEdBox1)
   End If
   If MaskEdBox2 <> "___/__/__" Then
      stCon = stCon & " and a1k02 <= " & ChangeTDateStringToTString(MaskEdBox2)
   End If
   If MaskEdBox1 <> "___/__/__" Or MaskEdBox2 <> "___/__/__" Then
      pub_QL05 = pub_QL05 & ";" & Label4 & MaskEdBox1 & "-" & MaskEdBox2 'Add By Sindy 2010/12/23
   End If
   
   Select Case Combo3.ListIndex
      Case 0 '外專
         stCon = stCon & " and st03 like 'F2%'"
      Case 1 '外商
         stCon = stCon & " and st03 like 'F1%'"
      Case 2 '外法
         stCon = stCon & " and st03>='F3' and st03<'F5'"
   End Select
   If Combo3.Text <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label3 & Combo3.Text 'Add By Sindy 2010/12/23
   End If
   
   '有對應收文號的請款項目
   stVTB = "select '1' id,nvl(st16,st03) st16,cp01,cp10,cp09,a1n05,a1k01 cp60" & _
      " From acc1k0, caseprogress, acc1n0, staff" & _
      " Where a1k12||a1k25 is null" & stCon & _
      " and cp60(+)=a1k01" & _
      " and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n06(+) is null" & _
      " and st01(+)=a1n04"
      
   '核稿點數
   stVTB = stVTB & " Union All " & _
      " select '2' id,nvl(st16,st03) st16,cp01,cp10,cp09,a1n05,a1k01 cp60" & _
      " From acc1k0, acc1n0, staff, caseprogress" & _
      " Where a1k12||a1k25 is null" & stCon & _
      " and a1n01(+)=a1k01 and a1n02(+)='2' and a1n06='Y' and cp09(+)=a1n03" & _
      " and st01(+)=a1n04"
      
   '沒有對應收文號的請款項目
   stVTB = stVTB & " Union All " & _
      " select '3' id,'' st16,a1k13 cp01,'X' cp10,a1k01 cp09,a1n05,a1k01 cp60" & _
      " From acc1k0, acc1n0, staff" & _
      " Where a1k12||a1k25 is null" & stCon & _
      " and a1n01(+)=a1k01 and a1n02(+)='2' and rtrim(a1n03) is null" & _
      " and st01(+)=a1n04"
   
   Select Case Combo2.ListIndex
      Case 0 '承辦人組別
         stSQL = "select '" & strUserNum & "' id,nvl(st16,'5') R01" & _
            ",cp01 R02,cp10 R03,count(distinct cp09) R04" & _
            ",sum(a1n05) R05" & _
            ",count(distinct decode(id,2,cp09)) R06" & _
            ",sum(decode(id,2,a1n05)) R07" & _
            " from (" & stVTB & ") group by st16,cp01,cp10"
      Case 1 '案件技術或語言
         stSQL = "select '" & strUserNum & "' id,nvl(st16,nvl(pa150,'5')) R01" & _
            ",cp01 R02,cp10 R03,count(distinct cp09) R04" & _
            ",sum(a1n05) R05" & _
            ",count(distinct decode(id,2,cp09)) R06" & _
            ",sum(decode(id,2,a1n05)) R07" & _
            " from (" & stVTB & "),acc1k0,patent" & _
            " where a1k01(+)=cp60 and pa01(+)=a1k13 and pa02(+)=a1k14 and pa03(+)=a1k15 and pa04(+)=a1k16" & _
            " group by nvl(st16,nvl(pa150,'5')),cp01,cp10"
      Case 2 '代理人國籍
         stSQL = "select '" & strUserNum & "' id,substr(nvl(fa10,cu10),1,3) R01" & _
            ",cp01 R02,cp10 R03,count(distinct cp09) R04" & _
            ",sum(a1n05) R05" & _
            ",count(distinct decode(id,2,cp09)) R06" & _
            ",sum(decode(id,2,a1n05)) R07" & _
            " from (" & stVTB & "),acc1k0,fagent,customer" & _
            " where a1k01(+)=cp60 and cu01(+)=substr(a1k28,1,8) and cu02(+)=substr(a1k28,9) " & _
            " and fa01(+)=substr(a1k28,1,8) and fa02(+)=substr(a1k28,9)" & _
            " group by substr(nvl(fa10,cu10),1,3),cp01,cp10"
   End Select
   If Combo2.Text <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label2 & Combo2.Text 'Add By Sindy 2010/12/23
   End If
   
   cnnConnection.Execute "delete from accRPT24k0 where id='" & strUserNum & "'", intI
   cnnConnection.Execute "insert into accRPT24k0(id,R01,R02,R03,R04,R05,R06,R07) " & stSQL, intI
   If intI > 0 Then
      ProduceData = True
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/23
      MsgBox "無可列印資料!!!", vbExclamation
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
   Text1.SetFocus
End Sub

'*************************************************
'  列印報表
'
'*************************************************
Private Sub PrintData()
   Dim strGrp As String, strGrpName As String, strSubTot As String, strTot As String
   
   If Combo2.ListIndex = 2 Then
      'Modify By Sindy 2015/9/2 substrb(decode(R03,'X','其他',decode(cpm03,'（無）',cpm04,cpm03)),1,16) 案件性質 ==> substrb(decode(R03,'X','其他',decode(cpm01,null,R03,decode(cpm03,'（無）',cpm04,cpm03))),1,16) 案件性質
      strExc(0) = "select SUBSTRB(na03,1,6) 國籍" & _
         ",R02 系統別" & _
         ",substrb(decode(R03,'X','其他',decode(cpm01,null,R03,decode(cpm03,'（無）',cpm04,cpm03))),1,16) 案件性質" & _
         ",R04 件數" & _
         ",R05 點數" & _
         ",R06 核稿件數" & _
         ",R07 核稿點數,R01" & _
         " from accRPT24k0,casepropertymap,nation" & _
         " where id='" & strUserNum & "' and cpm01(+)=R02 and cpm02(+)=R03" & _
         " and na01(+)=R01 order by R01,R02,R03"
   Else
      '2011/12/22 modify by sonia 無組別者改抓部門名稱
      'strExc(0) = "select " & IIf(Combo3.ListIndex = 0, "SUBSTRB(CST16(R01),1,6)", "decode(R01,'5','其他',R01)") & " 組別" & _
         ",R02 系統別" & _
         ",substrb(decode(R03,'X','其他',cpm03),1,16) 案件性質" & _
         ",R04 件數" & _
         ",R05 點數" & _
         ",R06 核稿件數" & _
         ",R07 核稿點數,R01" & _
         " from accRPT24k0,casepropertymap" & _
         " where id='" & strUserNum & "' and cpm01(+)=R02 and cpm02(+)=R03 order by R01,R02,R03"
      'Modify By Sindy 2015/9/2 substrb(decode(R03,'X','其他',decode(cpm03,'（無）',cpm04,cpm03)),1,16) 案件性質 ==> substrb(decode(R03,'X','其他',decode(cpm01,null,R03,decode(cpm03,'（無）',cpm04,cpm03))),1,16) 案件性質
      strExc(0) = "select " & IIf(Combo3.ListIndex = 0, "SUBSTRB(decode(a0902,null,CST16(R01),a0902),1,10)", "decode(a0902,null,R01,a0902)") & " 組別" & _
         ",R02 系統別" & _
         ",substrb(decode(R03,'X','其他',decode(cpm01,null,R03,decode(cpm03,'（無）',cpm04,cpm03))),1,16) 案件性質" & _
         ",R04 件數" & _
         ",R05 點數" & _
         ",R06 核稿件數" & _
         ",R07 核稿點數,R01" & _
         " from accRPT24k0,casepropertymap,acc090" & _
         " where id='" & strUserNum & "' and cpm01(+)=R02 and cpm02(+)=R03 and R01=a0901(+) order by R01,R02,R03"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   InsertQueryLog (RsTemp.RecordCount) 'Add By Sindy 2010/12/23
   If intI = 1 Then
      GetPleft
      Printer.PaperSize = 9 'A4
      Printer.Orientation = 1 '直印
      lngPageHeight = Printer.ScaleHeight
      lngPageWidth = Printer.ScaleWidth
      lngLineHeight = 300
      iPage = 0
      PrintHeader
      With RsTemp
      strTot = GetSubTot()
      strGrp = .Fields("R01")
      strGrpName = "" & .Fields(0)
      strSubTot = GetSubTot(strGrp)
      Do While Not .EOF
         If strGrp <> "" & .Fields("R01") Then
            PrintSum strSubTot, strGrpName, strTot
            strGrp = "" & .Fields("R01")
            strGrpName = "" & .Fields(0)
            strSubTot = GetSubTot(strGrp)
         End If
         NewLine
         For intI = 1 To 3
            Printer.CurrentX = PLeft(intI)
            Printer.CurrentY = iPrint
            Printer.Print .Fields(intI - 1)
         Next
         
         '件數
         strExc(0) = Format(.Fields(3), "#")
         Printer.CurrentX = PLeft(5) - 150 - Printer.TextWidth(strExc(0))
         Printer.CurrentY = iPrint
         Printer.Print strExc(0)
         
         '點數
         strExc(0) = Format(.Fields(4), "#.000")
         Printer.CurrentX = PLeft(6) - 150 - Printer.TextWidth(strExc(0))
         Printer.CurrentY = iPrint
         Printer.Print strExc(0)
         
         '百分比
         If Val(strSubTot) > 0 Then
            strExc(0) = Format(100 * .Fields(4) / Val(strSubTot), "#.00") & "%"
            Printer.CurrentX = PLeft(7) - 150 - Printer.TextWidth(strExc(0))
            Printer.CurrentY = iPrint
            Printer.Print strExc(0)
         End If
         
         '備註
         If .Fields(5) > 0 Then
            Printer.CurrentX = PLeft(7)
            Printer.CurrentY = iPrint
            Printer.Print "含核稿 " & .Fields(5) & " 件 " & .Fields(6) & " 點"
         End If
         .MoveNext
      Loop
      PrintSum strSubTot, strGrpName, strTot
      PrintSum strTot
      End With
      Printer.EndDoc
   End If
End Sub

Private Function GetSubTot(Optional p_strGrp As String) As String
   strExc(0) = "select sum(R05) from accRPT24k0 where id='" & strUserNum & "'"
   If p_strGrp <> "" Then strExc(0) = strExc(0) & " and R01='" & p_strGrp & "'"
   intI = 1
   Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetSubTot = "" & AdoRecordSet3(0)
   End If
End Function

'*************************************************
'  列印抬頭
'
'*************************************************
Private Sub PrintHeader()
   Dim strPTmp As String
   iPage = iPage + 1
   
   iPrint = 500
   Printer.FontName = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strPTmp = Me.Caption & "(" & Combo3 & ")"
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   NewLine
   strPTmp = "請款日期：" & MaskEdBox1 & " - " & MaskEdBox2
   Printer.CurrentX = 4000
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   
   NewLine
   strPTmp = "統計方式：" & Combo2
   Printer.CurrentX = 4000
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   
   NewLine
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   NewLine
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
    
   NewLine , True
   NewLine
   
   For intI = 1 To 7
      Select Case intI
         Case 4, 5, 6
            Printer.CurrentX = PLeft(intI + 1) - 150 - Printer.TextWidth(ColName(intI))
         Case Else
            Printer.CurrentX = PLeft(intI)
      End Select
      Printer.CurrentY = iPrint
      Printer.Print ColName(intI)
   Next
   NewLine , True
End Sub

'*************************************************
' 列印合計
'
'*************************************************
Private Sub PrintSum(p_SubTot As String, Optional p_Grp As String, Optional p_Tot As String)
   '組別
   If p_Grp <> "" Then
      NewLine , True
   End If
   
   NewLine
   
   Printer.FontBold = True
   
   '組別
   If p_Grp <> "" Then
      strExc(0) = p_Grp
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = iPrint
      Printer.Print strExc(0)
   End If
   
   '小計
   If p_Grp <> "" Then
      strExc(0) = "小計："
   Else
      strExc(0) = "合計："
   End If
   Printer.CurrentX = PLeft(5) - 150 - Printer.TextWidth(strExc(0))
   Printer.CurrentY = iPrint
   Printer.Print strExc(0)
   '點數
   strExc(0) = Format(Val(p_SubTot), "#.000")
   Printer.CurrentX = PLeft(6) - 150 - Printer.TextWidth(strExc(0))
   Printer.CurrentY = iPrint
   Printer.Print strExc(0)
   '百分比
   If Val(p_Tot) > 0 Then
      strExc(0) = Format(100 * Val(p_SubTot) / Val(p_Tot), "#.00") & "%"
      Printer.CurrentX = PLeft(7) - 150 - Printer.TextWidth(strExc(0))
      Printer.CurrentY = iPrint
      Printer.Print strExc(0)
   End If
   
   Printer.FontBold = False
   
   NewLine , True
   
   '2011/12/22 add by sonia 合計後加印備註
   If p_Grp = "" Then
      iPrint = iPrint + 500
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = iPrint
      Printer.Print "備註：１．組別欄的其他為歸屬於部門的點數(非個人點數)季分配至該部門的點數"
      iPrint = iPrint + 500
      Printer.CurrentY = iPrint
      Printer.Print "　　　　　２．案件性質欄的其他為請款項目與收文案件性質無法配對的點數"
      'Added by Morgan 2020/6/19
      iPrint = iPrint + 500
      Printer.CurrentY = iPrint
      Printer.Print "　　　　　３．新案翻譯僅含上班翻譯及核稿點數，其餘點數則歸屬其他"
      'end 2020/6/19
      
      'Added by Morgan 2024/9/5
      iPrint = iPrint + 500
      Printer.CurrentY = iPrint
      Printer.Print "　　　　　４．含律師分配外專點數，不含分配內專點數"
      'end 2024/9/5
      
   End If
   '2011/12/22 end
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Private Function FormCheck() As Boolean
   If Text1.Text = "" Then
      MsgBox "請輸入系統類別!!!", vbExclamation + vbOKOnly
      Text1.SetFocus
      Exit Function
   End If
   If MaskEdBox1 = "___/__/__" Then
      MsgBox "請輸入請款日期起日!!!", vbExclamation + vbOKOnly
      MaskEdBox1.SetFocus
      Exit Function
   End If
   If Not CheckIsTaiwanDate(ChangeTDateStringToTString(MaskEdBox1)) Then
      MsgBox "請款日期起日輸入錯誤!!!", vbExclamation + vbOKOnly
      MaskEdBox1.SetFocus
      Exit Function
   End If
   If MaskEdBox2 = "___/__/__" Then
      MsgBox "請輸入請款日期迄日!!!", vbExclamation + vbOKOnly
      MaskEdBox2.SetFocus
      Exit Function
   End If
   If Not CheckIsTaiwanDate(ChangeTDateStringToTString(MaskEdBox2)) Then
      MsgBox "請款日期迄日輸入錯誤!!!", vbExclamation + vbOKOnly
      MaskEdBox2.SetFocus
      Exit Function
   End If
   If Combo2.ListIndex < 0 Then
      MsgBox "請選擇統計方式!!!", vbExclamation + vbOKOnly
      Combo2.SetFocus
      Exit Function
   End If
   FormCheck = True
End Function

Private Sub SetCombo3()
   Combo3.AddItem "外專", 0
   Combo3.AddItem "外商", 1
   Combo3.AddItem "外法", 2
End Sub

Private Sub SetCombo2()
   Combo2.Clear
   Select Case Combo3.ListIndex
      Case 0  '外專
         Combo2.AddItem "承辦人組別", 0
         Combo2.AddItem "案件技術或語言", 1
         Combo2.AddItem "請款對象國籍別", 2
      Case 1 '外商
         Combo2.AddItem "承辦人組別", 0
      Case 2 '外法
         Combo2.AddItem "承辦人組別", 0
   End Select
End Sub
'Modified by Morgan 2024/9/5 備註增加,保留列數也要增加
Private Sub NewLine(Optional ByVal iExtraLines As Integer = 5, Optional bPulsLine As Boolean)
   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      PrintLine
      Printer.NewPage
      PrintHeader
      iPrint = iPrint + lngLineHeight
   ElseIf bPulsLine Then
      PrintLine
   End If
End Sub

Private Sub PrintLine(Optional strChar As String = "-")
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print String(98, strChar)
End Sub
