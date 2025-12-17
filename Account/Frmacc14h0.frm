VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc14h0 
   AutoRedraw      =   -1  'True
   Caption         =   "收據抬頭修改清單"
   ClientHeight    =   1900
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   4860
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1900
   ScaleWidth      =   4860
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   1470
      Style           =   2  '單純下拉式
      TabIndex        =   7
      Top             =   810
      Width           =   2820
   End
   Begin VB.TextBox txtDate 
      Height          =   300
      Left            =   1470
      MaxLength       =   3
      TabIndex        =   0
      Top             =   60
      Width           =   795
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
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   1200
      Width           =   4365
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1470
      TabIndex        =   1
      Top             =   450
      Width           =   1335
      _ExtentX        =   2364
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
      Left            =   3150
      TabIndex        =   2
      Top             =   450
      Width           =   1335
      _ExtentX        =   2364
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
      Left            =   390
      TabIndex        =   8
      Top             =   840
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "扣繳年度               (空白表全部)"
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
      Index           =   1
      Left            =   390
      TabIndex        =   6
      Top             =   90
      Width           =   3300
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   135
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
      Left            =   2910
      TabIndex        =   5
      Top             =   450
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "修改日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   390
      TabIndex        =   4
      Top             =   450
      Width           =   1455
   End
End
Attribute VB_Name = "Frmacc14h0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoquery As New ADODB.Recordset
Dim PLeft(0 To 7) As Integer
Dim m_intPage As Integer
Dim m_iPrint As Integer
'預設印表機
Dim strPrinter As String  'Add by Amy 2022/04/11

Private Sub Command1_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
    PUB_RestorePrinter cmbPrinter 'Add by Amy 2022/04/11
    PrintData
    PUB_RestorePrinter strPrinter 'Add by Amy 2022/04/11
    FormClear
   Screen.MousePointer = vbDefault
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
   Me.Width = 4980
   'Modify by Amy 2023/10/11 原H:2055
   Me.Height = 2370
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
   PUB_SetPrinter Me.Name, cmbPrinter, strPrinter 'Add by Amy 2022/04/11
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Amy 2022/04/11若印表機變動, 則更新列印設定
   If cmbPrinter.Text <> cmbPrinter.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
   End If
   
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc14h0 = Nothing
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
   txtDate = ""
   txtDate.SetFocus
   MaskEdBox1.SetFocus
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   'Add by Morgan 2005/3/18 加開立年度
   If txtDate <> "" Then
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

Private Sub PrintData()
Dim strSkipPage As String
   
   strSql = "": strSkipPage = ""
   'Modify by Morgan 2005/3/18 加開立年度
   'strSQL = strSQL & " And A0K31>=" & Val(FCDate(Me.MaskEdBox1.Text)) & " "
   'strSQL = strSQL & " And A0K31<=" & Val(FCDate(Me.MaskEdBox2.Text)) & " "
   If txtDate <> "" Then
      strSql = strSql & " and A0k16=" & Val(txtDate)
   End If
   If Me.MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " And A0K31>=" & Val(FCDate(Me.MaskEdBox1.Text)) & " "
   End If
   If Me.MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " And A0K31<=" & Val(FCDate(Me.MaskEdBox2.Text)) & " "
   End If
   
    'Modify by Morgan 2004/10/26 加 公司(a0k11)
    'strSQL = "Select A0K01, A0L02, A0J02, A0K03, A0K04, A1P22, Max(A0K11) as Comp, Max(A0K06) as Fee From ACC0K0, ACC0M0, ACC0L0, ACC0J0, ACC1P0 Where A0K01=A0M02 And A0M01=A0L01 And A0K01=A0J13 And A0L01=A1P04" &_
    'strSQL & " Group By A0K01, A0L02, A0J02, A0K03, A0K04, A1P22 Order By A0K03, A1P22 "
    'Modify By Sindy 2010/5/18 加申請國家(a0k23)
'    strSql = "Select A0K01, A0L02, A0J02, A0K03, A0K04, A1P22, Max(A0K11) as Comp, Max(decode(a0k30,'Y',nvl(A0K07,0)+nvl(A0K06,0),nvl(A0k06,0))) as Fee From ACC0K0, ACC0M0, ACC0L0, ACC0J0, ACC1P0 Where A0K01=A0M02 And A0M01=A0L01 And A0K01=A0J13 And A0L01=A1P04 AND A0K31>0" & _
'                    strSql & " Group By A0K01, A0L02, A0J02, A0K03, A0K04, A1P22 Order By Comp,A1P22,A0k01 "
    'Modified by Morgan 2011/11/25 考慮是否合併改判斷 a0j07
    'strSql = "Select A0K01, A0L02, A0J02, A0K03, A0K04, A1P22, Max(A0K11) as Comp, Max(decode(a0k30,'Y',nvl(A0K07,0)+nvl(A0K06,0),nvl(A0k06,0))) as Fee,na03,A0k16 From ACC0K0, ACC0M0, ACC0L0, ACC0J0, ACC1P0, Nation Where A0K01=A0M02 And A0M01=A0L01 And A0K01=A0J13 And A0L01=A1P04 AND A0K31>0 AND a0k23=na01(+) " & _
                    strSql & " Group By A0k16,A0K01, A0L02, A0J02, A0K03, A0K04, A1P22,na03 Order By A0k16,Comp,A1P22,A0k01 "
    '2012/1/13 modify by sonia 服務費欄改抓傳票,合併者抓收文+規費,不合併者只抓收入
    'strSql = "Select distinct A0K01, A0L02, A0J02, A0K03, A0K04, A1P22,Comp,Fee,na03,A0k16" & _
             " from (select A0K01, A0L02, A0J02, A0K03, A0K04, Max(A0K11) as Comp, sum(decode(a0j07,'Y',nvl(A0j09,0)+nvl(A0j10,0),nvl(A0j09,0))) as Fee" & _
             " ,na03,A0k16,a0l01 From ACC0K0, ACC0M0, ACC0L0, ACC0J0,Nation Where A0K01=A0M02 And A0M01=A0L01 And A0K01=A0J13" & _
             " AND A0K31>0 AND a0k23=na01(+) " & strSql & _
             " Group By A0k16,A0K01, A0L01, A0L02, A0J02, A0K03, A0K04,na03) x,acc1p0 where a1p04 (+)=a0l01" & _
             " Order By A0k16,Comp,A1P22,A0k01 "
    'modify by sonia 2021/3/17 +and a0k11=a1p01(+)條件,否則E10913159,E10910403會重覆(因為作帳公司不同)
    'mdofied by Morgan 2022/12/8 2公司收據要抓1公司分錄
    'modify by sonia 2024/4/1 1.收入要扣除借方扣支援點數E11204340,2.法律所案源案件抓2407XX科目為收入E11223369
    'strSql = "Select A0K01, A0L02, A0J02, A0K03, A0K04, A1P22, Comp,decode(a0j07,'Y',收入+規費,收入) as Fee,na03,A0k16 from" & _
             " (select distinct A0K01, A0L02, A0J02, A0K03, A0K04, A0K11 as Comp, a0j07, A0k16, a0l01, na03 From ACC0K0, ACC0M0, ACC0L0, ACC0J0,Nation" & _
             " Where A0K01=A0M02 And A0M01=A0L01 And A0K01=A0J13 AND A0K31>0 AND a0k23=na01(+) " & strSql & ") x," & _
             " (select a1p04,a1p17,a1p22,sum(收入) 收入,sum(規費) 規費 from (select distinct A1P04,a1p17,A1P22,a1p05,decode(substr(a1p05,1,1),'4',a1p08,0) 收入,decode(substr(a1p05,1,4),'2201',a1p08,0) 規費" & _
             " from ACC0K0, ACC0M0, ACC0L0, acc1p0 where A0K01=A0M02 And A0M01=A0L01 AND A0K31>0 and A0L01=A1P04(+) and decode(a0k11,'2','1',a0k11)=a1p01(+) " & strSql & _
             " ) GROUP BY A1P04,A1P17,A1P22) y where A0L01=a1p04(+) AND A0J02=A1P17(+) Order By A0k16,Comp,A1P22,A0k01"
    strSql = "Select A0K01, A0L02, A0J02, A0K03, A0K04, A1P22, Comp,decode(a0j07,'Y',收入+規費,收入) as Fee,na03,A0k16 from" & _
             " (select distinct A0K01, A0L02, A0J02, A0K03, A0K04, A0K11 as Comp, a0j07, A0k16, a0l01, na03 From ACC0K0, ACC0M0, ACC0L0, ACC0J0,Nation" & _
             " Where A0K01=A0M02 And A0M01=A0L01 And A0K01=A0J13 AND A0K31>0 AND a0k23=na01(+) " & strSql & ") x," & _
             " (select a1p04,a1p17,a1p22,sum(收入) 收入,sum(規費) 規費 from (select distinct A1P04,a1p17,A1P22,a1p05,decode(substr(a1p05,1,1),'4',a1p07*-1+a1p08,decode(substr(a1p05,1,4),'2407',a1p07*-1+a1p08,0)) 收入,decode(substr(a1p05,1,4),'2201',a1p08,0) 規費" & _
             " from ACC0K0, ACC0M0, ACC0L0, acc1p0 where A0K01=A0M02 And A0M01=A0L01 AND A0K31>0 and A0L01=A1P04(+) and decode(a0k11,'2','1',a0k11)=a1p01(+) " & strSql & _
             " ) GROUP BY A1P04,A1P17,A1P22) y where A0L01=a1p04(+) AND A0J02=A1P17(+) Order By A0k16,Comp,A1P22,A0k01"
    adoquery.CursorLocation = adUseClient
    adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
    If adoquery.RecordCount <> 0 Then
        GetPrintLeft
        m_intPage = 0
        'PrintHead
        Do While adoquery.EOF = False
            'Modify By Sindy 2010/5/18
            If m_iPrint > 15100 Or _
               (strSkipPage <> adoquery.Fields("A0k16").Value) Then
                If strSkipPage <> "" Then Printer.NewPage
                Call PrintHead("" & adoquery.Fields("A0k16").Value)
            End If
            strSkipPage = "" & adoquery.Fields("A0k16").Value
            '2010/5/18 End
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = m_iPrint
            'Modify by Morgan 2004/10/26 加 公司
            Printer.Print " " & adoquery.Fields("Comp").Value
            Printer.CurrentX = PLeft(6)
            Printer.CurrentY = m_iPrint
            '2004/10/26 end
            Printer.Print "" & adoquery.Fields(0).Value
            'Modify by Morgan 2004/10/26 改印服務費
            'Printer.CurrentX = PLeft(1)
            'Printer.CurrentY = m_iPrint
            'Printer.Print "" & adoquery.Fields(1).Value
            'Modify by Morgan 2005/3/18
            'Printer.CurrentX = PLeft(2) - Printer.TextWidth(Format(Val("" & adoquery.Fields("Fee").Value), DDollar)) - 100
            Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(Val("" & adoquery.Fields("Fee").Value), DDollar)) - 200
            Printer.CurrentY = m_iPrint
            Printer.Print Format(Val("" & adoquery.Fields("Fee").Value), DDollar)
            '2004/10/26 end
            'Add By Sindy 2010/5/18
            Printer.CurrentX = PLeft(7)
            Printer.CurrentY = m_iPrint
            Printer.Print Left(Trim("" & adoquery.Fields("na03").Value) & "    ", 4)
            '2010/5/18 End
            Printer.CurrentX = PLeft(2)
            Printer.CurrentY = m_iPrint
            Printer.Print "" & adoquery.Fields(2).Value
            Printer.CurrentX = PLeft(3)
            Printer.CurrentY = m_iPrint
            Printer.Print "" & adoquery.Fields(3).Value
            Printer.CurrentX = PLeft(4)
            Printer.CurrentY = m_iPrint
            Printer.Print StrToStr("" & adoquery.Fields(4).Value, 15) '15個中文字
            Printer.CurrentX = PLeft(5)
            Printer.CurrentY = m_iPrint
            Printer.Print "" & adoquery.Fields(5).Value
            m_iPrint = m_iPrint + 300
            adoquery.MoveNext
        Loop
        Printer.CurrentX = PLeft(0)
        Printer.CurrentY = m_iPrint
        Printer.Print String(110, "=")
        m_iPrint = m_iPrint + 300
        Printer.EndDoc
        ShowPrintOk
    Else
        MsgBox MsgText(28), , MsgText(5)
    End If
    adoquery.Close
End Sub

Private Sub GetPrintLeft()
   'Modify by Morgan 2004/10/16
   'PLeft(0) = 0 '收據編號
   'PLeft(1) = PLeft(0) + 1664 '收款日期
   PLeft(0) = 0 '公司
'   PLeft(6) = PLeft(0) + 500 '收據編號
'   'Modify by Morgna 2005/3/18
'   'PLeft(1) = PLeft(0) + 1664 '服務費
'   ''2004/10/16 end
'   'PLeft(2) = PLeft(1) + 936 '本所案號
'   'PLeft(3) = PLeft(2) + 1664 '客戶編號
'   'PLeft(4) = PLeft(3) + 1040 '收據抬頭
'   'PLeft(5) = PLeft(4) + 3224 '收款傳票號碼
'   PLeft(2) = PLeft(0) + 1664 '本所案號
'   PLeft(3) = PLeft(2) + 1664 '客戶編號
'   PLeft(4) = PLeft(3) + 1040 '收據抬頭
'   PLeft(1) = PLeft(4) + 3224 '服務費
'   PLeft(5) = PLeft(1) + 936 '收款傳票號碼
   '2005/3/18 end
   
   PLeft(6) = 500 '收據編號
   PLeft(2) = 1700 '本所案號
   PLeft(3) = 4900 '客戶編號
   PLeft(4) = 5900 '收據抬頭
   PLeft(1) = 9000 '服務費
   PLeft(5) = 10000 '收款傳票號碼
   PLeft(7) = 3300 '申請國家 Add By Sindy 2010/5/18
End Sub

Private Sub PrintHead(strYear As String)
   m_intPage = m_intPage + 1
   m_iPrint = 0: Printer.Orientation = 1
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 20
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 3328
   Printer.CurrentY = m_iPrint
   Printer.Print "*** 收據抬頭修改清單 ***"
   m_iPrint = m_iPrint + 500
   Printer.Font.Size = 10
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = m_iPrint
   Printer.Print "列印人：" & GetStaffName(strUserNum)
   'Add by Morgan 2005/3/18
   'If txtDate <> "" Then
      Printer.CurrentX = 3744
      Printer.CurrentY = m_iPrint
      'Printer.Print "扣繳年度：" & txtDate
      Printer.Print "扣繳年度：" & strYear
      m_iPrint = m_iPrint + 300
   'End If
   '2005/3/8 end
   Printer.CurrentX = 3744
   Printer.CurrentY = m_iPrint
   Printer.Print "修改日期：" & Me.MaskEdBox1.Text & "－" & Me.MaskEdBox2.Text
   Printer.CurrentX = 9000
   Printer.CurrentY = m_iPrint
   Printer.Print "列印日期：" & CFDate(ACDate(ServerDate))
   m_iPrint = m_iPrint + 300
   Printer.CurrentX = 9000
   Printer.CurrentY = m_iPrint
   Printer.Print "頁　　次：" & str(m_intPage)
   m_iPrint = m_iPrint + 300
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = m_iPrint
   Printer.Print String(110, "-")
   m_iPrint = m_iPrint + 300
        
    Printer.CurrentX = PLeft(0)
    Printer.CurrentY = m_iPrint
    'Modify by Morgan 2004/10/26 加 公司
    Printer.Print "公司"
    Printer.CurrentX = PLeft(6)
    Printer.CurrentY = m_iPrint
    '2004/10/26 end
    Printer.Print "收據編號"
    Printer.CurrentX = PLeft(1)
    Printer.CurrentY = m_iPrint
    'Modify by Morgan 2004/10/26
    'Printer.Print "收款日期"
    Printer.Print "服務費"
    Printer.CurrentX = PLeft(2)
    Printer.CurrentY = m_iPrint
    Printer.Print "本所案號"
    'Add By Sindy 2010/5/18
    Printer.CurrentX = PLeft(7)
    Printer.CurrentY = m_iPrint
    Printer.Print "申請國家"
    '2010/5/18 End
    Printer.CurrentX = PLeft(3)
    Printer.CurrentY = m_iPrint
    Printer.Print "客戶編號"
    Printer.CurrentX = PLeft(4)
    Printer.CurrentY = m_iPrint
    Printer.Print "收據抬頭"
    Printer.CurrentX = PLeft(5)
    Printer.CurrentY = m_iPrint
    Printer.Print "傳票號碼"
    m_iPrint = m_iPrint + 300
   
    Printer.CurrentX = PLeft(0)
    Printer.CurrentY = m_iPrint
    Printer.Print String(110, "-")
    m_iPrint = m_iPrint + 300
End Sub

Private Sub txtDate_GotFocus()
   TextInverse txtDate
   CloseIme
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii < Asc("0") And KeyAscii > Asc("9") Then
      KeyAscii = 0
      Beep
   End If
End Sub

