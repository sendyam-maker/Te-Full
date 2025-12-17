VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmacc14k0 
   AutoRedraw      =   -1  'True
   Caption         =   "翻譯費明細表"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5172
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   5172
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
      Height          =   315
      Left            =   1530
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1290
      Width           =   1575
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
      Left            =   1530
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1650
      Width           =   300
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
      Height          =   315
      Left            =   1530
      MaxLength       =   6
      TabIndex        =   3
      Top             =   900
      Width           =   1575
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   1530
      Style           =   2  '單純下拉式
      TabIndex        =   6
      Top             =   2040
      Width           =   2820
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
      TabIndex        =   7
      Top             =   2550
      Width           =   4770
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
      Height          =   315
      Left            =   1530
      MaxLength       =   1
      TabIndex        =   2
      Top             =   495
      Width           =   300
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1530
      TabIndex        =   0
      Top             =   120
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
      Left            =   3435
      TabIndex        =   1
      Top             =   120
      Width           =   1575
      _ExtentX        =   2794
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
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "(迄)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1125
      TabIndex        =   15
      Top             =   1290
      Width           =   450
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "lblStaffName"
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
      Left            =   3240
      TabIndex        =   14
      Top             =   1320
      Width           =   1740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "印証明單             (N:不印)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   135
      TabIndex        =   13
      Top             =   1680
      Width           =   2700
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "員工代號 (起)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   135
      TabIndex        =   12
      Top             =   930
      Width           =   1350
   End
   Begin VB.Label lblStaffName 
      BackStyle       =   0  '透明
      Caption         =   "lblStaffName"
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
      Left            =   3240
      TabIndex        =   11
      Top             =   930
      Width           =   1740
   End
   Begin VB.Label Label15 
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
      Left            =   135
      TabIndex        =   10
      Top             =   2070
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   210
      Left            =   135
      TabIndex        =   9
      Top             =   165
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   3150
      X2              =   3380
      Y1              =   270
      Y2              =   270
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "身分                   (1:內翻  2:外翻  空白:全部)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   135
      TabIndex        =   8
      Top             =   525
      Width           =   4590
   End
End
Attribute VB_Name = "frmacc14k0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
'Add by Morgan 2007/6/4
Option Explicit

Public m_bolCalled As Boolean 'Added by Morgan 2024/5/24

Dim PLeft(0 To 6) As Integer
Dim m_intPage As Integer
Dim m_iPrint As Integer
'預設印表機
Dim m_DefaultPrinter As String, m_Prn As Printer
Dim m_Grp As String, m_RptNo As String, m_NAME As String, m_Rate1 As String, m_Rate2 As String, m_Rate3 As String, m_IdKind As String
Dim lngXo As Long, lngYo As Long, lngX As Long, lngY As Long '列印位置
Dim m_lPage As Long '總頁次
'Added by Morgan 2013/1/31
Dim m_lTax As Long '代扣所得稅
Dim m_lHFee As Long '代扣補充保費
'end 2013/1/31
Dim m_bolNew As Boolean, m_bolOld As Boolean 'Added by Morgan 2019/8/15

Private Sub Command1_Click()
   FormPrint
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   'Removed by Morgan 2024/5/22 不需要，且薪資要用會有錯
   'If KeyCode <> vbKeyEscape Then
   '   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   'End If
End Sub

Private Sub Form_Load()
   '表單初始化
   PUB_InitForm Me, 5295, 3465
   PUB_SetPrinter Me.Name, cmbPrinter, m_DefaultPrinter
   '畫面初值設定
   FormClear
   'Removed by Morgan 2024/5/22 不需要，且薪資要用會有錯
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
   'end 2024/5/22
End Sub

Private Sub FormClear()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text1 = ""
   Text2 = ""
   lblStaffName = ""
   Text3 = ""
   Label4 = ""
   Text4 = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   '若印表機變動, 則更新列印設定
   If cmbPrinter.Text <> cmbPrinter.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
   End If
   PUB_RestorePrinter m_DefaultPrinter
   
   Set frmacc14k0 = Nothing
End Sub


Private Function FormCheck() As Boolean
   If MaskEdBox1.Text = MsgText(29) Then
      MsgBox "入帳日期不可空白！"
      MaskEdBox1.SetFocus
      Exit Function
   ElseIf MaskEdBox2.Text = MsgText(29) Then
      MsgBox "入帳日期不可空白！"
      MaskEdBox2.SetFocus
      Exit Function
   Else
      FormCheck = True
   End If
End Function

Private Sub MaskEdBox1_GotFocus()
   If MaskEdBox1.Text <> MsgText(29) Then
      MaskEdBox1.SelStart = 0
      MaskEdBox1.SelLength = MaskEdBox1.MaxLength
   End If
End Sub

Private Sub MaskEdBox2_GotFocus()
   If MaskEdBox2.Text = MsgText(29) And MaskEdBox1.Text <> MsgText(29) Then
      MaskEdBox2 = MaskEdBox1
      MaskEdBox2.SelStart = 0
      MaskEdBox2.SelLength = MaskEdBox2.MaxLength
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 = "" Then
      lblStaffName = ""
   Else
      lblStaffName = GetStaffName(Text2, True)
   End If
End Sub

Private Sub Text3_GotFocus()
   If Text2 <> "" And Text3 = "" Then
      Text3 = Text2
   End If
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If Text3 = "" Then
      Label4 = ""
   Else
      Label4 = GetStaffName(Text3, True)
   End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub FormPrint()
   Dim strCon As String
   strCon = ""
   '入帳日期
   If MaskEdBox1.Text <> MsgText(29) Then
      strCon = strCon & " and a1p18>=" & Val(FCDate(MaskEdBox1.Text))
   'Added by Morgan 2014/11/6
   Else
      MsgBox "請輸入入帳日期起日！", vbExclamation
      MaskEdBox1.SetFocus
      Exit Sub
   'end 2014/11/6
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      strCon = strCon & " and a1p18<=" & Val(FCDate(MaskEdBox2.Text))
   'Added by Morgan 2014/11/6
   Else
      MsgBox "請輸入入帳日期止日！", vbExclamation
      MaskEdBox2.SetFocus
      Exit Sub
   'end 2014/11/6
   End If
   
   'Modify by Morgan 2011/3/3 改判斷部門
   '內翻
   If Text4 = "1" Then
      'strCon = strCon & " and s2.st04='1'"
      strCon = strCon & " and s1.st03='F52'"
   '外翻
   ElseIf Text4 = "2" Then
      'strCon = strCon & " and nvl(s2.st04,'2')='2'"
      strCon = strCon & " and s1.st03='F51'"
   End If
   
   '員工代號
   If Text2 <> "" Then
      strCon = strCon & " and a1p15>='" & Text2 & "'"
   End If
   '員工代號
   If Text3 <> "" Then
      strCon = strCon & " and a1p15<='" & Text3 & "'"
   End If
   
   'Added by Morgan 2013/2/18 王雅萍 F5542 直接開發票不必明細也不印證明單--瑞婷
   'strCon = strCon & " and a1p15<>'F5542'" 'Removed by Morgan 2018/4/2 改判斷有 Transfee 資料者(王雅萍)--婧瑄
   'end 2013/2/18
   
   'strExc(0) = " select decode(s2.st04,'1','內翻','外翻') C01" & _
      ",a1p15,a1p17,s1.st02||'('||a1p15||')' C02,a1p04,TF02,TF03,TF04,nvl(TF05,100) TF05" & _
      ",nvl(TF06,100) TF06,TF15,TF16,TF17,a1p07,SPR02,SPR03,SPR04,s1.st02,s2.st04,A2501,TF18" & _
      " from acc1p0,staff s1,staff_idmap,staff s2,TRANSFEE,STAFF_PAYRATE,ACC250" & _
      " where a1p05='6130'" & strCon & _
      " AND s1.st01(+)=a1p15 and sim02(+)=a1p15 and s2.st01(+)=sim01 and s1.st01 is not null" & _
      " AND TF07(+)=A1P04 AND TF14(+)=A1P17 AND SPR01(+)=A1P15" & _
      " AND A2502(+)='5' AND A2503(+)=A1P15 AND A2505(+)=A1P04 order by 1,2,3"
   'Modified by Morgan 2013/1/31 +補充保費,所得稅
   'Modified by Morgan 2017/2/13 +TF21,TF22,SPR11
   'Modified by Morgan 2018/4/2 剔除沒有 Transfee 資料者--婧瑄
   'Modified by Morgan 2019/7/3 不需再抓staff_idmap,新編號復職且關聯同一F編號時資料會重複 Ex:F5656顧家盛 ->A1025,A8008
   'strExc(0) = " select decode(s1.st03,'F52','內翻','外翻') C01" & _
      ",a1p15,a1p17,s1.st02||'('||a1p15||')' C02,a1p04,TF02,TF03,TF04,nvl(TF05,100) TF05" & _
      ",nvl(TF06,100) TF06,TF15,TF16,TF17,TF21,TF22,a1p07,SPR02,SPR03,SPR04,SPR11,s1.st02,decode(s1.st03,'F52','1','2') C03,A2501,TF18,OD06,OD13" & _
      " from acc1p0,staff s1,staff_idmap,staff s2,TRANSFEE,STAFF_PAYRATE,ACC250,othersalarydata" & _
      " where a1p05='6130'" & strCon & _
      " AND s1.st01(+)=a1p15 and sim02(+)=a1p15 and s2.st01(+)=sim01 and s1.st01 is not null" & _
      " AND TF07(+)=A1P04 AND TF14(+)=A1P17 and tf01 is not null AND SPR01(+)=A1P15" & _
      " AND A2502(+)='5' AND A2503(+)=A1P15 AND A2505(+)=A1P04" & _
      " AND OD03(+)=a1p15 AND OD02(+)=A1P18+19110000 order by 1,2,3"
   'Modified by Morgan 2019/8/15 +,TF27,EP07,SPR12,SPR13
   strExc(0) = " select decode(s1.st03,'F52','內翻','外翻') C01" & _
      ",a1p15,a1p17,s1.st02||'('||a1p15||')' C02,a1p04,TF02,TF03,TF04,nvl(TF05,100) TF05" & _
      ",nvl(TF06,100) TF06,TF15,TF16,TF17,TF21,TF22,a1p07,SPR02,SPR03,SPR04,SPR11,s1.st02,decode(s1.st03,'F52','1','2') C03,A2501,TF18,OD06,OD13,TF27,EP09,SPR12,SPR13" & _
      " from acc1p0,staff s1,TRANSFEE,STAFF_PAYRATE,ACC250,othersalarydata,engineerprogress" & _
      " where a1p05='6130'" & strCon & _
      " AND s1.st01(+)=a1p15 and s1.st01 is not null" & _
      " AND TF07(+)=A1P04 AND TF14(+)=A1P17 and tf01 is not null AND SPR01(+)=A1P15" & _
      " AND A2502(+)='5' AND A2503(+)=A1P15 AND A2505(+)=A1P04" & _
      " AND OD03(+)=a1p15 AND OD02(+)=A1P18+19110000 and ep02(+)=TF01 order by 1,2,3"
   'end 2019/7/3
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      DoPrint RsTemp.Clone
   Else
      MsgBox "查無資料！"
   End If
   FormClear
   
End Sub

Private Function AddNewA250(p_A2503 As String, p_A2504 As Long, p_A2505 As String, p_A2513 As String, p_A2514 As String) As String
   Dim stRetNo As String
   stRetNo = AutoNo("H", 5)
   strSql = "INSERT INTO ACC250(A2501,A2502,A2503,A2504,A2505,A2506,A2513,A2514) VALUES('" & stRetNo & "','5','" & p_A2503 & "'," & p_A2504 & ",'" & p_A2505 & "','" & strUserNum & "','" & ChgSQL(p_A2513) & "','" & ChgSQL(p_A2514) & "')"
   adoTaie.Execute strSql
   AddNewA250 = stRetNo
End Function

Private Sub DoPrint(p_Rst As ADODB.Recordset)

   Dim strTemp As String, lngTotal As Long
   Dim dblTranFee As Double, dblTypFee As Double, dblDisRate1 As Double, dblDisRate2 As Double
   Dim strName As String, strAddr As String, ii As Integer
   Dim strA2501 As String, strA2501s As String, strDesc As String, lngSubTotal As Long
   Dim strID As String, strNo As String, strLstNo As String
   Dim dblPlusRate As Double 'Add by Morgan 2007/8/13
   Dim arrA250
   'Added by Morgan 2019/8/15
   Dim rstTmp As ADODB.Recordset
   'end 2019/8/15
   
On Error GoTo flgErr

   Set rstTmp = p_Rst.Clone
   
   If Not m_bolCalled Then 'Added by Morgan 2024/5/24
   
      '設定使用者所選擇的印表機成預設印表機
      For Each m_Prn In Printers
         If m_Prn.DeviceName = cmbPrinter.Text Then
            Set Printer = m_Prn
            Exit For
         End If
      Next
      
   End If
   
   GetPleft
   Printer.FontName = "標楷體"
   Printer.Orientation = 1 '直印
   m_NAME = ""
   lngYo = 0
   m_lPage = 0
   With p_Rst
      .MoveFirst
      Do While Not .EOF
         strNo = "" & .Fields("a1p04")
         'Add by Morgan 2007/12/5
         '一個翻譯可能會有多張應付單
         '外翻要印回執信封
         If m_IdKind <> "1" And Text1 <> "N" Then
            If strLstNo <> "" And strLstNo <> strNo Then
               If strA2501 = "" Then
                  strA2501 = AddNewA250(strID, lngSubTotal, strLstNo, strName, strDesc)
               End If
               strA2501s = strA2501s & "," & strA2501
               strDesc = ""
               lngSubTotal = 0
            End If
         End If
         strA2501 = "" & .Fields("a2501")
         If m_NAME <> .Fields("C02") Then
            If m_NAME <> "" Then
               PrintTotal lngTotal
               '外翻要印回執信封
               If m_IdKind <> "1" And Text1 <> "N" Then
                  'Modify by Morgan 2007/12/5 回執單會印多張
                  'If strA2501 = "" Then
                  '   strA2501 = AddNewA250(strID, lngTotal, strNo, strName, strDesc)
                  'End If
                  'NewPage
                  'PUB_PrintReceipt5 strA2501, lngYo
                  If strA2501s <> "" Then
                     arrA250 = Split(strA2501s, ",")
                     For ii = 1 To UBound(arrA250)
                        strTemp = arrA250(ii)
                        NewPage
                        PUB_PrintReceipt5 strTemp, lngYo
                     Next
                  End If
                  strA2501s = ""
                  NewPage
                  PrintReCover strAddr, strName
                  'end 2007/12/5
               End If
            End If
            lngTotal = 0
            lngSubTotal = 0
            m_intPage = 0
            m_Grp = "" & .Fields("C01")
            m_IdKind = "" & .Fields("C03")
            strName = "" & .Fields("st02")
            strID = "" & .Fields("a1p15")
            strDesc = ""
            m_Rate3 = Val("" & .Fields("SPR11"))  'Added by Morgan 2017/2/13
            
            'Added by Morgan 2019/8/15
            m_bolNew = False: m_bolOld = False
            rstTmp.MoveFirst
            rstTmp.Find "C02='" & .Fields("C02") & "'", , adSearchForward, 1
            Do While Not rstTmp.EOF
               If rstTmp.Fields("ep09") >= "20190815" Then
                  m_bolNew = True
               Else
                  m_bolOld = True
               End If
               rstTmp.Find "C02='" & .Fields("C02") & "'", 1, adSearchForward
            Loop
            
            '新公式
            If m_bolNew And Not m_bolOld Then
               m_RptNo = "4"
               '英文
               If .Fields("TF27") = "1" Then
                  m_Rate1 = Val("" & .Fields("SPR12"))
               '日文
               ElseIf .Fields("TF27") = "2" Then
                  m_Rate1 = Val("" & .Fields("SPR13"))
               End If
            Else
            'end 2019/8/15
            
               '日文
               If Val("" & .Fields("TF03")) > 0 Then
                  m_RptNo = "2"
                  m_Rate1 = Val("" & .Fields("SPR03"))
               '英文
               Else
                  'Modified by Morgan 2017/2/13
                  'm_RptNo = "1"
                  If Val("" & .Fields("TF21")) > 0 Then
                     m_RptNo = "3"
                  Else
                     m_RptNo = "1"
                  End If
                  'end 2017/2/13
                  m_Rate1 = Val("" & .Fields("SPR02"))
               End If
               m_Rate2 = Val("" & .Fields("SPR04"))
               
            End If '新公式
            
            '外翻要印信封
            If m_IdKind <> "1" And Text1 <> "N" Then
               strAddr = GetAddr(strID)
               PrintCover strAddr, strName
            End If
            
            m_NAME = "" & .Fields("C02")
            
            'Added by Morgan 2013/1/31
            m_lTax = Val("" & .Fields("OD06"))
            m_lHFee = Val("" & .Fields("OD13"))
            'end 2013/1/31
            
            PrintHead
         Else
            NewLine
         End If
         
         'Added by Morgan 2019/8/15
         '新公式
         If .Fields("ep09") >= "20190815" Then
            m_RptNo = "3"
         Else
         'end 2019/8/15
         
            If Val("" & .Fields("TF03")) > 0 Then
               m_RptNo = "2"
            '英文
            Else
               If Val("" & .Fields("TF21")) > 0 Then
                  m_RptNo = "3"
               Else
                  m_RptNo = "1"
               End If
            End If
               
            '翻譯費
            If m_RptNo = "1" Then
               dblTranFee = Round((Val("" & .Fields("TF02")) + Val("" & .Fields("TF04"))) * Val(m_Rate1) / 1000)
            'Added by Morgan 2017/2/13
            '以英文字數計費
            ElseIf m_RptNo = "3" Then
               dblTranFee = Round(Val("" & .Fields("TF21")) * Val(m_Rate3) / 1000 * 0.8)
            'end 2017/2/13
            Else
               dblTranFee = Round(Val("" & .Fields("TF03")) * Val(m_Rate1) / 1000)
            End If
            '打字費
            'Modified by Morgan 2017/2/13 英文字數計費沒有打字費
            If m_RptNo = "3" Then
               dblTypFee = 0
            Else
               dblTypFee = Round((Val("" & .Fields("TF02")) + Val("" & .Fields("TF04"))) * Val(m_Rate2) / 1000)
            End If
         End If
            
         strTemp = "" & .Fields("a1p17")
         strTemp = ChgCaseNo(strTemp)

         strDesc = strDesc & IIf(strDesc = "", strTemp, "," & Mid(strTemp, 5))
         lngTotal = lngTotal + Val("" & .Fields("a1p07"))
         lngSubTotal = lngSubTotal + Val("" & .Fields("a1p07"))
         
         'Added by Morgan 2019/8/15
         '新公式
         If m_bolNew And Not m_bolOld Then
            '案號
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = lngYo + m_iPrint
            Printer.Print strTemp
            '原文字數
            strTemp = Val("" & .Fields("TF21"))
            Printer.CurrentX = PLeft(2) - Printer.TextWidth(strTemp) - 200
            Printer.CurrentY = lngYo + m_iPrint
            Printer.Print strTemp
            '相似折扣
            strTemp = Val("" & .Fields("TF05"))
            If Val(strTemp) > 0 And Val(strTemp) <> 100 Then
               Printer.CurrentX = PLeft(3) - Printer.TextWidth(strTemp) - 200
               Printer.CurrentY = lngYo + m_iPrint
               Printer.Print strTemp
            End If
            '瑕疵折扣
            strTemp = Val("" & .Fields("TF06"))
            If Val(strTemp) > 0 And Val(strTemp) <> 100 Then
               Printer.CurrentX = PLeft(4) - Printer.TextWidth(strTemp) - 200
               Printer.CurrentY = lngYo + m_iPrint
               Printer.Print strTemp
            End If
            '加成比率
            strTemp = Val("" & .Fields("TF18"))
            If Val(strTemp) > 0 And Val(strTemp) <> 100 Then
               Printer.CurrentX = PLeft(5) - Printer.TextWidth(strTemp) - 200
               Printer.CurrentY = lngYo + m_iPrint
               Printer.Print strTemp
            End If
            
            '翻譯費
            strTemp = Format(Val("" & .Fields("a1p07")), "#,##0")
            Printer.CurrentX = PLeft(6) - Printer.TextWidth(strTemp)
            Printer.CurrentY = lngYo + m_iPrint
            Printer.Print strTemp
            
         Else
         'end 2019/8/15
            
            'Added by Morgan 2019/8/15
            If .Fields("ep09") >= "20190815" Then
               strTemp = strTemp & "原"
               dblTypFee = 0
               dblTranFee = Val("" & .Fields("a1p07"))
            Else
            'end 2019/8/15
            
               'Added by Morgan 2017/2/13
               If m_RptNo = "3" Then
                  strTemp = strTemp & "英"
               End If
               'end 2017/2/13
               
            End If 'Added by Morgan 2018/8/15
            
            '相似折扣
            dblDisRate1 = Val("" & .Fields("TF05"))
            If dblDisRate1 > 0 And dblDisRate1 <> 100 Then
               strTemp = "*" & strTemp
            End If
            
            '瑕疵折扣
            dblDisRate2 = Val("" & .Fields("TF06"))
            If dblDisRate2 > 0 And dblDisRate2 <> 100 Then
               strTemp = "**" & strTemp
            End If
            
            '加成比率
            dblPlusRate = Val("" & .Fields("TF18"))
            If dblPlusRate > 0 And dblPlusRate <> 100 Then
               strTemp = "@" & strTemp
            End If
            
            '案號
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = lngYo + m_iPrint
            Printer.Print strTemp
            If m_RptNo = "1" Then
               '中文字數
               strTemp = Val("" & .Fields("TF02")) + Val("" & .Fields("TF04"))
            'Added by Morgan 2017/2/13
            ElseIf m_RptNo = "3" Then
               '英文字數
               strTemp = Val("" & .Fields("TF21"))
            'end 2017/2/13
            Else
               '日文字數
               strTemp = Val("" & .Fields("TF03"))
            End If
            Printer.CurrentX = PLeft(2) - Printer.TextWidth(strTemp) - 200
            Printer.CurrentY = lngYo + m_iPrint
            Printer.Print strTemp
            
            '翻譯費
            strTemp = Format(dblTranFee, "#,##0")
            Printer.CurrentX = PLeft(3) - Printer.TextWidth(strTemp) - 200
            Printer.CurrentY = lngYo + m_iPrint
            Printer.Print strTemp
            
            If m_RptNo = "2" Then
               '中文字數
               strTemp = Val("" & .Fields("TF02")) + Val("" & .Fields("TF04"))
               Printer.CurrentX = PLeft(4) - Printer.TextWidth(strTemp) - 200
               Printer.CurrentY = lngYo + m_iPrint
               Printer.Print strTemp
            End If
            
            '中打費
            strTemp = Format(dblTypFee, "#,##0")
            Printer.CurrentX = PLeft(5) - Printer.TextWidth(strTemp) - 200
            Printer.CurrentY = lngYo + m_iPrint
            Printer.Print strTemp
            
            '小計
            strTemp = Format(Val("" & .Fields("a1p07")), "#,##0")
            Printer.CurrentX = PLeft(6) - Printer.TextWidth(strTemp)
            Printer.CurrentY = lngYo + m_iPrint
            Printer.Print strTemp
            
         End If
         
         strLstNo = strNo
         .MoveNext
      Loop
      PrintTotal lngTotal
      '外翻要印回執信封
      If m_IdKind <> "1" And Text1 <> "N" Then
         'Modify by Morgan 2007/12/5
         'If strA2501 = "" Then
         '   strA2501 = AddNewA250(strID, lngTotal, strNo, strName, strDesc)
         'End If
         'NewPage
         'PUB_PrintReceipt5 strA2501, lngYo
         If strA2501 = "" Then
            strA2501 = AddNewA250(strID, lngSubTotal, strLstNo, strName, strDesc)
         End If
         strA2501s = strA2501s & "," & strA2501
         If strA2501s <> "" Then
            arrA250 = Split(strA2501s, ",")
            For ii = 1 To UBound(arrA250)
               strTemp = arrA250(ii)
               NewPage
               PUB_PrintReceipt5 strTemp, lngYo
            Next
         End If
         'end 2007/12/5
         NewPage
         PrintReCover strAddr, strName
      End If
      Printer.EndDoc
      If Not m_bolCalled Then 'Added by Morgan 2024/5/24
         MsgBox "列印完成！"
      End If
   End With
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Sub

Private Sub NewPage()
   m_lPage = m_lPage + 1
   If m_lPage > 1 Then
      If lngYo > 0 Then
         Printer.NewPage
         lngYo = 0
      Else
         lngYo = Printer.Height / 2
      End If
   End If
End Sub

Private Sub NewLine(Optional bolNoLine As Boolean = False)
   m_iPrint = m_iPrint + 300
   If m_iPrint + 900 > Printer.ScaleHeight / 2 Then
      Printer.DrawStyle = vbSolid
      If bolNoLine = False Then
         Printer.Line (PLeft(0), lngYo + m_iPrint)-(PLeft(6), lngYo + m_iPrint)
      End If
      PrintHead
   End If
End Sub

Private Sub GetPleft()
   PLeft(0) = 1000 '案號
   If m_RptNo = "2" Or m_RptNo = "4" Then
      PLeft(1) = PLeft(0) + 2000 '日文字數
      PLeft(2) = PLeft(1) + 1500 '翻譯費
      PLeft(3) = PLeft(2) + 1500 '中文字數
      PLeft(4) = PLeft(3) + 1500 '中打費
      PLeft(5) = PLeft(4) + 1500 '小計
      PLeft(6) = PLeft(5) + 1500 '右邊
   Else
      PLeft(1) = PLeft(0) + 2000 '中文字數
      PLeft(2) = PLeft(1) + 1500 '翻譯費
      PLeft(3) = PLeft(2) + 1500 '
      PLeft(4) = PLeft(2) + 1500 '中打費
      PLeft(5) = PLeft(4) + 1500 '小計
      PLeft(6) = PLeft(5) + 1500 '右邊
   End If
   
End Sub

Private Sub PrintHead()

   Dim strTemp As String
   
   NewPage

   m_intPage = m_intPage + 1
   m_iPrint = 500:
   GetPleft
   With Printer
      
      '表頭
      .Font.Size = 18
      .Font.Bold = True
      .Font.Underline = True
      
      strTemp = m_NAME & "翻譯費明細表"
      .CurrentX = PLeft(0) + (PLeft(6) - PLeft(0) - .TextWidth(strTemp)) / 2
      .CurrentY = lngYo + m_iPrint
      Printer.Print strTemp
      
      .Font.Size = 10
      .Font.Bold = False
      .Font.Underline = False
      
      '跳列
      m_iPrint = m_iPrint + 600
      
      '條件
      strTemp = "入帳日期: " & MaskEdBox1 & " － " & MaskEdBox2
      .CurrentX = PLeft(0) + (PLeft(6) - PLeft(0) - .TextWidth(strTemp)) / 2
      .CurrentY = lngYo + m_iPrint
      Printer.Print strTemp
      
      '跳列
      m_iPrint = m_iPrint + 300
      
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = lngYo + m_iPrint
      Printer.Print "列印人：" & strUserName
      
      
      strTemp = "列印日期：" & CFDate(strSrvDate(2))
      .CurrentX = PLeft(6) - .TextWidth(strTemp)
      .CurrentY = lngYo + m_iPrint
      Printer.Print strTemp
      
      '跳列
      m_iPrint = m_iPrint + 300
      
      strTemp = "列印日期：" & CFDate(strSrvDate(2))
      .CurrentX = PLeft(6) - .TextWidth(strTemp)
      .CurrentY = lngYo + m_iPrint
      Printer.Print "頁　　次：　" & str(m_intPage)
      
      m_iPrint = m_iPrint + 300
      DrawLine
      
      .Font.Size = 12
      .Font.Bold = True
      
      'Added by Morgan 2019/8/15
      '新公式
      If m_bolNew And Not m_bolOld Then
            .CurrentX = PLeft(0)
            .CurrentY = lngYo + m_iPrint
            Printer.Print "案號"
            .CurrentX = PLeft(1)
            .CurrentY = lngYo + m_iPrint
            Printer.Print "原文字數"
            .CurrentX = PLeft(2)
            .CurrentY = lngYo + m_iPrint
            Printer.Print "相似折扣"
            .CurrentX = PLeft(3)
            .CurrentY = lngYo + m_iPrint
            Printer.Print "瑕疵折扣"
            .CurrentX = PLeft(4)
            .CurrentY = lngYo + m_iPrint
            Printer.Print "加成比率"
            .CurrentX = PLeft(5)
            .CurrentY = lngYo + m_iPrint
            Printer.Print "翻譯費"
      Else
      'end 2019/8/15
      
         If m_RptNo = "2" Then
            .CurrentX = PLeft(0)
            .CurrentY = lngYo + m_iPrint
            Printer.Print "案號"
            .CurrentX = PLeft(1)
            .CurrentY = lngYo + m_iPrint
            Printer.Print "日文字數"
            .CurrentX = PLeft(2)
            .CurrentY = lngYo + m_iPrint
            Printer.Print "翻譯費"
            .CurrentX = PLeft(3)
            .CurrentY = lngYo + m_iPrint
            Printer.Print "中文字數"
            .CurrentX = PLeft(4)
            .CurrentY = lngYo + m_iPrint
            Printer.Print "中打費"
            .CurrentX = PLeft(5)
            .CurrentY = lngYo + m_iPrint
            Printer.Print "小計"
         Else
            .CurrentX = PLeft(0)
            .CurrentY = lngYo + m_iPrint
            Printer.Print "案號"
            .CurrentX = PLeft(1)
            .CurrentY = lngYo + m_iPrint
            
            'Modified by Morgan 2017/2/13
            'Printer.Print "中文字數"
            'Modified by Morgan 2019/8/15
            'If m_Rate3 > 0 Then
            '   Printer.Print "中/英文字數"
            'Else
            '   Printer.Print "中文字數"
            'End If
            Printer.Print "字數"
            'end 2019/8/15
            'end 2017/2/13
            .CurrentX = PLeft(2)
            .CurrentY = lngYo + m_iPrint
            Printer.Print "翻譯費"
            .CurrentX = PLeft(4)
            .CurrentY = lngYo + m_iPrint
            Printer.Print "中打費"
            .CurrentX = PLeft(5)
            .CurrentY = lngYo + m_iPrint
            Printer.Print "小計"
         End If
         
      End If
      .Font.Bold = False
      
      m_iPrint = m_iPrint + 300
      DrawLine
   End With
End Sub
   
Private Sub PrintTotal(p_lngTot As Long)
   Dim strTemp As String
   Dim ii As Integer
   
   NewLine
   
   DrawLine
   Printer.Font.Bold = True
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = lngYo + m_iPrint
   Printer.Print "總計"
   strTemp = Format(p_lngTot, "$#,##0")
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(strTemp)
   Printer.CurrentY = lngYo + m_iPrint
   Printer.Print strTemp
      
   'Modified by Morgan 2014/6/3
   'm_iPrint = m_iPrint + 300
   NewLine True
   
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = lngYo + m_iPrint
   Printer.Print "代扣所得稅"
   strTemp = Format(m_lTax, "$#,##0")
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(strTemp)
   Printer.CurrentY = lngYo + m_iPrint
   Printer.Print strTemp
   
   'Modified by Morgan 2014/6/3
   'm_iPrint = m_iPrint + 300
   NewLine True
   
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = lngYo + m_iPrint
   Printer.Print "代扣補充保費"
   strTemp = Format(m_lHFee, "$#,##0")
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(strTemp)
   Printer.CurrentY = lngYo + m_iPrint
   Printer.Print strTemp
   
   'Modified by Morgan 2014/6/3
   'm_iPrint = m_iPrint + 300
   NewLine True
   
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = lngYo + m_iPrint
   Printer.Print "實際金額"
   strTemp = Format(p_lngTot - m_lTax - m_lHFee, "$#,##0")
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(strTemp)
   Printer.CurrentY = lngYo + m_iPrint
   Printer.Print strTemp
   
   Printer.Font.Bold = False
   If m_iPrint + 2000 > Printer.ScaleHeight / 2 Then
      PrintHead
   End If
   
   Printer.Font.Size = 10
   m_iPrint = m_iPrint + 600
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = lngYo + m_iPrint
   
   'Added by Morgan 2019/8/15
   '新公式
   If m_bolNew And Not m_bolOld Then
      Printer.Print "備註：翻譯費率：" & m_Rate1 & "/每千個原文字數"
   
   Else
      If m_RptNo = "2" Then
         Printer.Print "備註：1. 翻譯費 rate：" & m_Rate1 & "/每千個日文字"
      Else
         Printer.Print "備註：1. 翻譯費 rate：" & m_Rate1 & "/每千個中文字"
      End If
      
      m_iPrint = m_iPrint
      Printer.CurrentX = PLeft(0) + Printer.TextWidth("備註：") + 4000
      Printer.CurrentY = lngYo + m_iPrint
      Printer.Print "2. 中打費 rate：" & m_Rate2 & "/每千個中文字"
      
      m_iPrint = m_iPrint + 300
      Printer.CurrentX = PLeft(0) + Printer.TextWidth("備註：")
      Printer.CurrentY = lngYo + m_iPrint
      Printer.Print "3. ""@"" 表示該案已計算加成部分"
      
      m_iPrint = m_iPrint
      Printer.CurrentX = PLeft(0) + Printer.TextWidth("備註：") + 4000
      Printer.CurrentY = lngYo + m_iPrint
      Printer.Print "4. ""*"" 表示該案已扣除相似折扣部分"
      
      m_iPrint = m_iPrint + 300
      Printer.CurrentX = PLeft(0) + Printer.TextWidth("備註：")
      Printer.CurrentY = lngYo + m_iPrint
      Printer.Print "5. ""**"" 表示該案已扣除瑕疵折扣部分"
      
      m_iPrint = m_iPrint
      Printer.CurrentX = PLeft(0) + Printer.TextWidth("備註：") + 4000
      Printer.CurrentY = lngYo + m_iPrint
      Printer.Print "6. ""***"" 表示該案已扣除相似折扣部分及瑕疵折扣部分"
      ii = 7
      'Added by Morgan 2017/2/13
      If m_Rate3 > 0 Then
         m_iPrint = m_iPrint + 300
         Printer.CurrentX = PLeft(0) + Printer.TextWidth("備註：")
         Printer.CurrentY = lngYo + m_iPrint
         Printer.Print "7. ""英""表示該案以英文字數計費(翻譯費=英文字數x英文翻譯費率x80%)"
         ii = 8
      End If
      'end 2017/2/13
      
      'Added by Morgan 2019/8/15
      If m_bolNew Then
         m_iPrint = m_iPrint + 300
         Printer.CurrentX = PLeft(0) + Printer.TextWidth("備註：")
         Printer.CurrentY = lngYo + m_iPrint
         Printer.Print ii & ". ""原""表示該案以新方式計算"
      End If
      'end 2019/8/15
   End If
End Sub

Private Sub DrawLine()
   Printer.DrawStyle = vbSolid
   Printer.DrawWidth = 4
   Printer.Line (PLeft(0), lngYo + m_iPrint)-(PLeft(6), lngYo + m_iPrint)
   m_iPrint = m_iPrint + 100
End Sub

Private Function ChgCaseNo(p_CaseNo As String) As String
   If Len(p_CaseNo) < 10 Then
      ChgCaseNo = p_CaseNo
   ElseIf Right(p_CaseNo, 3) = "000" Then
      ChgCaseNo = Left(p_CaseNo, Len(p_CaseNo) - 9) & "-" & Left(Right(p_CaseNo, 9), 6)
   Else
      ChgCaseNo = Left(p_CaseNo, Len(p_CaseNo) - 9) & "-" & Right(p_CaseNo, 9)
   End If
End Function
'信封
Private Sub PrintCover(p_CustAddr As String, p_CustName As String)
   Dim lngX1 As Long, lngY1 As Long, lngX2 As Long, lngY2 As Long
   
   NewPage
   
   Printer.FontSize = 14
   Printer.CurrentX = 750
   Printer.CurrentY = lngYo + 600
   Printer.Print "寄件人：台一國際智慧財產事務所"
   Printer.CurrentX = 750
   Printer.CurrentY = lngYo + 950
   Printer.Print "地　址：１０４台北市長安東路二段１１２號九樓"
   Printer.CurrentX = 750
   Printer.CurrentY = lngYo + 1300
   Printer.Print "電　話：０２－２５０６１０２３"
   
   Printer.FontSize = 18
   Printer.CurrentX = 9800
   Printer.CurrentY = lngYo + 600
   strExc(1) = "平信"
   Printer.Print strExc(1)
   
   lngX1 = 9750
   lngX2 = lngX1 + Printer.TextWidth(strExc(1)) + 100
   lngY1 = lngYo + 550
   lngY2 = lngY1 + Printer.TextHeight(strExc(1)) + 100
   Printer.Line (lngX1, lngY1)-(lngX2, lngY1)
   Printer.Line (lngX1, lngY1)-(lngX1, lngY2)
   Printer.Line (lngX2, lngY1)-(lngX2, lngY2)
   Printer.Line (lngX1, lngY2)-(lngX2, lngY2)
   
   Printer.FontSize = 14
   Printer.CurrentX = 1850
   Printer.CurrentY = lngYo + 3700
   strExc(1) = "收件人："
   Printer.Print strExc(1)
   
   lngX1 = 1850 + Printer.TextWidth(strExc(1))
   lngY1 = lngYo + 3700
   Pub_SmartPrint p_CustAddr, lngX1, lngY1, 125, 350
   
   lngY1 = lngY1 + 400
   Printer.FontSize = 20
   Pub_SmartPrint p_CustName, lngX1, lngY1, 125, 400
   
End Sub

'回執信封
Private Sub PrintReCover(p_CustAddr As String, p_CustName As String)
   Dim lngX1 As Long, lngY1 As Long, lngX2 As Long, lngY2 As Long
   
   NewPage
   
   Printer.FontSize = 14
   lngX1 = 750
   lngY1 = lngYo + 600
   Pub_SmartPrint "寄件人：" & p_CustName, lngX1, lngY1, 125, 350
   
   lngX1 = 750
   lngY1 = lngY1 + 350
   Pub_SmartPrint "地　址：" & p_CustAddr, lngX1, lngY1, 125, 350
      
   Printer.FontSize = 18
   Printer.CurrentX = 9800
   Printer.CurrentY = lngYo + 600
   strExc(1) = "平信"
   Printer.Print strExc(1)
   
   lngX1 = 9750
   lngX2 = lngX1 + Printer.TextWidth(strExc(1)) + 100
   lngY1 = lngYo + 550
   lngY2 = lngY1 + Printer.TextHeight(strExc(1)) + 100
   Printer.Line (lngX1, lngY1)-(lngX2, lngY1)
   Printer.Line (lngX1, lngY1)-(lngX1, lngY2)
   Printer.Line (lngX2, lngY1)-(lngX2, lngY2)
   Printer.Line (lngX1, lngY2)-(lngX2, lngY2)
   
   Printer.FontSize = 14
   Printer.CurrentX = 1850
   Printer.CurrentY = lngYo + 3700
   strExc(1) = "收件人：１０４台北市長安東路二段１１２號九樓"
   Printer.Print strExc(1)
   
   
   Printer.CurrentX = 1850 + Printer.TextWidth("收件人：")
   Printer.CurrentY = lngYo + 4100
   Printer.FontSize = 20
   Printer.Print "台一國際智慧財產事務所"
   
   Printer.CurrentX = 5000
   Printer.CurrentY = lngYo + 5800
   Printer.Print "會計部"
   
End Sub

Private Function GetAddr(p_ID As String) As String
   'Modify by Morgan 2009/1/21 廠商郵遞區號欄位自地址欄拆開a0i04
   strExc(0) = "select a0i04||a0i03 from acc0i0 where a0i01='" & p_ID & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetAddr = Trim("" & RsTemp(0))
   End If
End Function

'Added by Morgan 2024/5/22 因薪資系統會呼叫此表單,新增此同名函數以便與財務系統相容
Private Sub KeyEnter(InputCode As Integer)
   If InputCode = vbKeyEscape Then '離開
      If LCase(App.EXEName) = "account" Then
         tool4_enabled
      End If
      Unload Me
   End If
End Sub
