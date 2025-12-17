VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc14l0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "收據帳款明細列印"
   ClientHeight    =   2655
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   5145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5145
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   1680
      Style           =   2  '單純下拉式
      TabIndex        =   15
      Top             =   1680
      Width           =   2820
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
      Index           =   1
      Left            =   3150
      MaxLength       =   9
      TabIndex        =   3
      Top             =   510
      Width           =   1590
   End
   Begin VB.CheckBox chk 
      Caption         =   "智權人員格式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   1845
      TabIndex        =   7
      Top             =   1290
      Width           =   1700
   End
   Begin VB.CheckBox chk 
      Caption         =   "客戶格式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   330
      TabIndex        =   6
      Top             =   1290
      Value           =   1  '核取
      Width           =   1470
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
      Index           =   1
      Left            =   3150
      MaxLength       =   9
      TabIndex        =   5
      Top             =   960
      Width           =   1575
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
      Index           =   0
      Left            =   1230
      MaxLength       =   9
      TabIndex        =   4
      Top             =   960
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
      Index           =   0
      Left            =   1230
      MaxLength       =   9
      TabIndex        =   2
      Top             =   510
      Width           =   1590
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
      Left            =   270
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   2040
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   1230
      TabIndex        =   0
      Top             =   60
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
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
      Left            =   3150
      TabIndex        =   1
      Top             =   60
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
   Begin VB.Label Label15 
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
      Left            =   600
      TabIndex        =   16
      Top             =   1710
      Width           =   750
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
      Index           =   2
      Left            =   2925
      TabIndex        =   14
      Top             =   540
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收據號碼"
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
      Index           =   1
      Left            =   270
      TabIndex        =   13
      Top             =   990
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
      Index           =   1
      Left            =   2910
      TabIndex        =   12
      Top             =   990
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號"
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
      Left            =   270
      TabIndex        =   11
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
      Index           =   0
      Left            =   2910
      TabIndex        =   10
      Top             =   60
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收據日期"
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
      Index           =   0
      Left            =   270
      TabIndex        =   9
      Top             =   60
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc14l0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Created by Morgan 2011/10/3
Option Explicit

Dim PLeft() As Integer
Dim intY As Integer
Dim strPrinter As String 'Add by Amy 2023/10/11

Private Sub chk_Click(Index As Integer)
   chk(Abs(Index - 1)).Value = Abs(chk(Index).Value - 1)
End Sub

Private Sub Command1_Click()
   Dim strCon As String
   Dim rsReport As ADODB.Recordset
   Dim lstNo As String, lstCaseNo As String
   Dim iPageNo As Integer, iItemCount As Integer
   Dim lngTotal As Long, lngPoint As Long
   
   If FormCheck = False Then Exit Sub
      
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strCon = strCon & " and a0k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strCon = strCon & " and a0k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Text1(0) <> "" Then
      strCon = strCon & " and a0k03 >= '" & Text1(0) & "'"
   End If
   If Text1(1) <> "" Then
      strCon = strCon & " and a0k03 <= '" & Text1(1) & "'"
   End If
   If Text2(0) <> "" Then
      strCon = strCon & " and a0k01 >= '" & Text2(0) & "'"
   End If
   If Text2(1) <> "" Then
      strCon = strCon & " and a0k01 <= '" & Text2(1) & "'"
   End If
   
   '排序:業務區,智權人員,客戶號,收據號
   strExc(0) = "select a0k01,A0K02,a0k03,a0k20,a0k22,CU04,cp01,cp02,cp03,cp04,cp05,cp09,a0j04,a0j09,a0j10,cpm03,cpm04,na03" & _
      " from acc0k0,customer,acc0j0,caseprogress,casepropertymap,nation" & _
      " where cu01(+)=substr(a0k03,1,8) and cu02(+)=substr(a0k03,9)" & strCon & _
      " and a0j13(+)=a0k01 and cp09(+)=a0j01 and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=a0j04" & _
      " order by a0k22,a0k20,a0k03,a0k01,cp09"
   intI = 1
   Set rsReport = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      PUB_RestorePrinter cmbPrinter 'Add by Amy 2023/10/11
      With rsReport
      GetPleft
      Printer.PaperSize = PUB_GetPaperSize(3) '中一刀
      Printer.Font.Name = "細明體"
      Printer.FontSize = 12
      lstNo = ""
      Do While Not .EOF
         If .Fields("a0k01") <> lstNo Then
            If lstNo <> "" Then
               PrintFoot lngTotal, lngPoint
               Printer.NewPage
            End If
            lstNo = .Fields("a0k01")
            PrintHead "" & .Fields("CU04"), .Fields("A0K02"), .Fields("A0K01"), iPageNo
            iItemCount = 0
            lngTotal = 0
            lngPoint = 0
            lstCaseNo = ""
         Else
            iItemCount = iItemCount + 1
            'Modified by Morgan 2012/6/8
            'If iItemCount > 15 Then
            If intY > Printer.Height - 1000 Then
               PrintHead .Fields("CU04"), .Fields("A0K02"), .Fields("A0K01"), iPageNo
               iItemCount = 0
            Else
               intY = intY + 300
            End If
         End If
         
         '本所案號
         If .Fields("cp01") & .Fields("cp02") & .Fields("cp03") & .Fields("cp04") <> lstCaseNo Then
            strExc(1) = .Fields("cp01") & .Fields("cp02") & .Fields("cp03") & .Fields("cp04")
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = intY
            Printer.Print strExc(1)
               
            '申請國家
            strExc(1) = "" & .Fields("na03")
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = intY
            Printer.Print StrToStr(strExc(1), 4)
         End If
                  
         '收文日
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = intY
         Printer.Print CFDate(TransDate("" & .Fields("cp05"), 1))
         
         '案件性質
         If .Fields("a0j04") = "020" Then
            strExc(1) = "" & .Fields("cpm04")
         Else
            strExc(1) = "" & .Fields("cpm03")
         End If
         Printer.CurrentX = PLeft(3)
         Printer.CurrentY = intY
         Printer.Print StrToStr(strExc(1), 5)
         
         '應收金額
         strExc(1) = Format(Val("" & .Fields("a0j09")) + Val("" & .Fields("a0j10")), DDollar)
         Printer.CurrentX = PLeft(5) - Printer.TextWidth(strExc(1)) - 200
         Printer.CurrentY = intY
         Printer.Print strExc(1)
         
         lngTotal = lngTotal + Val("" & .Fields("a0j09")) + Val("" & .Fields("a0j10"))
         lngPoint = lngPoint + Val("" & .Fields("a0j09"))
         
         '案件名稱
         If .Fields("cp01") & .Fields("cp02") & .Fields("cp03") & .Fields("cp04") <> lstCaseNo Then
            strExc(1) = GetPrjName(.Fields("cp01") & "-" & .Fields("cp02") & "-" & .Fields("cp03") & "-" & .Fields("cp04"))
            Printer.CurrentX = PLeft(5)
            Printer.CurrentY = intY
            Printer.Print StrToStr(strExc(1), 16)
         End If
         lstCaseNo = "" & .Fields("cp01") & .Fields("cp02") & .Fields("cp03") & .Fields("cp04")
         .MoveNext
      Loop
      PrintFoot lngTotal, lngPoint
      Printer.EndDoc
      End With
      PUB_RestorePrinter strPrinter 'Add by Amy 202310/11
   Else
      MsgBox "無符合資料！"
   End If
   Set rsReport = Nothing
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height
   
   'Removed by Morgan 2012/6/8 不要預設--瑞婷
   'MaskEdBox1.Text = CFDate(strSrvDate(2))
   MaskEdBox1.Mask = DFormat
   'MaskEdBox2.Text = CFDate(strSrvDate(2))
   MaskEdBox2.Mask = DFormat
   PUB_SetPrinter Me.Name, cmbPrinter, strPrinter 'Add by Amy 2023/10/11
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(151)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Amy 2023/10/11 若印表機變動, 則更新列印設定
   If cmbPrinter.Text <> cmbPrinter.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
   End If
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc14l0 = Nothing
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   MaskEdBox2.Mask = MsgText(601)
   MaskEdBox2.Text = MaskEdBox1.Text
   MaskEdBox2.Mask = DFormat
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   CloseIme
   If Index = 1 Then
      If Text1(0).Text <> "" Then
         'Modify By Sindy 2014/8/11 999=>ZZZ
         'Text1(1).Text = Left(Text1(0).Text, 6) & "999"
         Text1(1).Text = Left(Text1(0).Text, 6) & "ZZZ"
      End If
   End If
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)

End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Index = 1 Then Exit Sub
   
   If Text1(Index) = MsgText(601) Then
      Exit Sub
   End If
   Select Case Len(Text1(Index))
      Case 6
         Text1(Index) = Text1(Index) & "000"
      Case 8
         Text1(Index) = Text1(Index) & "0"
   End Select
End Sub

Private Sub Text2_GotFocus(Index As Integer)
   CloseIme
   TextInverse Text2(Index)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Private Function FormCheck() As Boolean
   If MaskEdBox1.Text <> MsgText(29) And MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If Text2(0) <> "" And Text2(1) <> "" Then
      FormCheck = True
      Exit Function
   End If
   MsgBox "條件不足，至少需輸入收據日期起迄或收據號碼起迄!!", vbCritical
End Function

Private Sub GetPleft()
   ReDim PLeft(6) As Integer
   
   '本所案號
   PLeft(0) = 500
   '申請國家
   PLeft(1) = PLeft(0) + 2.8 * 567
   '收文日期
   PLeft(2) = PLeft(1) + 2 * 567
   '案件性質
   PLeft(3) = PLeft(2) + 2.3 * 567
   '應收金額
   PLeft(4) = PLeft(3) + 2.5 * 567
   '案件名稱
   PLeft(5) = PLeft(4) + 2.5 * 567
End Sub

Private Sub PrintHead(pCustName As String, pDate As String, pNo As String, ByRef pPageNo As Integer)
   pPageNo = pPageNo + 1
   
   If pPageNo > 1 Then Printer.NewPage
   
   intY = 600
   
   strExc(1) = "帳款明細表"
   Printer.FontSize = 18
   Printer.CurrentX = 10.5 * 567 - Printer.TextWidth(strExc(1)) / 2
   Printer.CurrentY = intY
   Printer.Print strExc(1)
               
   intY = intY + 500
   
   strExc(1) = "客戶名稱：" & pCustName
   Printer.FontSize = 12
   Printer.CurrentX = 10.5 * 567 - Printer.TextWidth(strExc(1)) / 2
   Printer.CurrentY = intY
   Printer.Print strExc(1)
   
   Printer.CurrentX = 9000
   Printer.CurrentY = intY
   Printer.Print "收據日期：" & CFDate(pDate)
   
   intY = intY + 300
   
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = intY
   Printer.Print "頁　次：" & pPageNo
   
   Printer.CurrentX = 9000
   Printer.CurrentY = intY
   Printer.Print "收據編號：" & pNo
   
   intY = intY + 350
   
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = intY
   Printer.Print "本所案號"
   
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = intY
   Printer.Print "申請國家"
   
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = intY
   Printer.Print "收文日期"
   
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = intY
   Printer.Print "案件性質"
   
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = intY
   Printer.Print "應收金額"
   
   Printer.CurrentX = PLeft(5) + 2 * 567
   Printer.CurrentY = intY
   Printer.Print "案件名稱"
   
   intY = intY + 300
   
   Printer.Line (PLeft(0) - 50, intY)-(11300, intY)
   intY = intY + 50
End Sub

Private Sub PrintFoot(pTotal As Long, pPoint As Long)
   intY = intY + 300
   Printer.Line (PLeft(0) - 50, intY)-(11300, intY)
   
   intY = intY + 50
   If chk(1).Value = 1 Then
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = intY
      Printer.Print "點數總計：" & Round(pPoint / 1000, 2)
   End If
   
   strExc(1) = "合計：" & Format(pTotal, DDollar)
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(strExc(1)) - 200
   Printer.CurrentY = intY
   Printer.Print strExc(1)
End Sub
