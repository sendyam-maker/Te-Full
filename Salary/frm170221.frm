VERSION 5.00
Begin VB.Form frm170221 
   BorderStyle     =   1  '單線固定
   Caption         =   "扣繳憑單套印"
   ClientHeight    =   3660
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4764
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4764
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   5
      Left            =   2010
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "Y"
      Top             =   2220
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   4
      Left            =   2865
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "Y"
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   1470
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "Y"
      Top             =   960
      Width           =   255
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   975
      Style           =   2  '單純下拉式
      TabIndex        =   15
      Top             =   2550
      Width           =   2730
   End
   Begin VB.TextBox txtSet 
      Height          =   285
      Index           =   1
      Left            =   1695
      TabIndex        =   14
      Top             =   2910
      Width           =   705
   End
   Begin VB.TextBox txtSet 
      Height          =   285
      Index           =   2
      Left            =   1695
      TabIndex        =   13
      Top             =   3210
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   1470
      MaxLength       =   12
      TabIndex        =   3
      Text            =   "123456789012"
      Top             =   1605
      Width           =   1290
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1470
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "1"
      Top             =   1290
      Width           =   255
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2100
      TabIndex        =   6
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3195
      TabIndex        =   7
      Top             =   90
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Index           =   0
      Left            =   1470
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "96"
      Top             =   630
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "是否為補印地址條：          ( Y 是 )"
      Height          =   180
      Index           =   8
      Left            =   255
      TabIndex        =   21
      Top             =   2280
      Width           =   2625
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "非本所在職員工是否印地址條：        ( Y 是 )"
      Height          =   180
      Index           =   7
      Left            =   255
      TabIndex        =   20
      Top             =   1980
      Width           =   3435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "股利格式：          ( Y 是 )"
      Height          =   180
      Index           =   6
      Left            =   435
      TabIndex        =   19
      Top             =   990
      Width           =   1905
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      Height          =   180
      Index           =   3
      Left            =   255
      TabIndex        =   18
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "橫軸偏移值(X)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   4
      Left            =   255
      TabIndex        =   17
      Top             =   2970
      Width           =   3240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "縱軸偏移值(Y)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   5
      Left            =   255
      TabIndex        =   16
      Top             =   3270
      Width           =   3240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "台一"
      Height          =   180
      Index           =   1
      Left            =   2865
      TabIndex        =   12
      Top             =   1650
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "台一專利商標"
      Height          =   180
      Index           =   0
      Left            =   1830
      TabIndex        =   11
      Top             =   1320
      Width           =   1080
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "所得人代號："
      Height          =   180
      Index           =   2
      Left            =   270
      TabIndex        =   10
      Top             =   1650
      Width           =   1080
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "公  司  別："
      Height          =   180
      Index           =   1
      Left            =   450
      TabIndex        =   9
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "扣繳年度："
      Height          =   180
      Index           =   0
      Left            =   450
      TabIndex        =   8
      Top             =   660
      Width           =   900
   End
End
Attribute VB_Name = "frm170221"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Memo by Morgan 2024/1/31 改新部門
'Create by Morgan 2009/1/23
Option Explicit

Dim m_Actived As Boolean
Dim m_DefaultPrinter As String

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         Screen.MousePointer = vbHourglass
         If TxtValidate = True Then
            Me.Enabled = False
            If cmbPrinter <> Printer.DeviceName Then
               PUB_RestorePrinter cmbPrinter
            End If
            PrintSheet
            '若印表機變動, 則更新列印設定
            If cmbPrinter.Tag <> cmbPrinter Or txtSet(1) <> txtSet(1).Tag Or txtSet(2) <> txtSet(2).Tag Then
                PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, Val(txtSet(1)), Val(txtSet(2)), Me.cmbPrinter.Text
            End If
            If Printer.DeviceName <> m_DefaultPrinter Then
               PUB_RestorePrinter m_DefaultPrinter
            End If
            Me.Enabled = True
         End If
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cmbPrinter, m_DefaultPrinter, , , Me.txtSet(1), Me.txtSet(2)
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set frm170221 = Nothing
End Sub

Private Sub FormReset()
   Dim oText As TextBox
   For Each oText In Text1
      Text1(oText.Index) = Empty
   Next
   Text1(0).Text = strSrvDate(2) \ 10000 - 1
   Text1(4).Text = "Y"
   Label1(0) = ""
   Label1(1) = ""
End Sub

Private Sub Form_Activate()
   If m_Actived = False Then
      FormReset
      Text1(0).SetFocus
      Text1_GotFocus 0
      m_Actived = True
   End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
   Case 0
      If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
         KeyAscii = 0
         Beep
      End If
   Case 3, 4, 5
      If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
         KeyAscii = 0
         Beep
      End If
   Case Else
   
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 1
      If Text1(Index) <> "" Then
         Label1(0) = CompNameQuery(Text1(Index))
         If Label1(0) = "" Then
            ShowMsg "公司別錯誤 !"
            Cancel = True
         End If
      End If
      
   Case 2
      If Text1(Index) <> "" Then
         If ClsPDGetOtherIncomer(Text1(Index), strExc(1)) = True Then
            Label1(1) = strExc(1)
         Else
            'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
            'If ChkStaffID(Replace(Text1(Index), "A", "0")) = True Then
            If ChkStaffID(Left(Text1(Index), 1) & Replace(Mid(Text1(Index), 2), "A", "0")) = True Then
               Cancel = True
            End If
            If Cancel = False Then
               If ClsPDGetStaffN(Text1(Index), strExc(1), , True) = False Then
                  Cancel = True
                  Label1(1) = ""
               Else
                  Label1(1) = strExc(1)
               End If
            End If
         End If
      End If
   End Select
End Sub


Private Sub PrintSheet()
   Dim YM As String, stCon As String, adoRst As ADODB.Recordset
   Dim Xo As Integer, Yo As Integer, xi As Long, yi As Long
   Dim dblUnitWidth As Double, dblUnitHeight As Double
   Dim strDesc1 As String, strDesc2 As String
   Dim strFontSize As String, strFontName As String
   Dim bOnlyAddress As Boolean, strLstID As String
   Dim stVTB As String
   
   Xo = 0 + Val(txtSet(1)) * 567
   Yo = -240 + Val(txtSet(2)) * 567
   dblUnitWidth = 1300 '欄位寬
   dblUnitHeight = 480 '欄位高
   
   stCon = ""
   '年度
   If Text1(0) <> "" Then
      stCon = stCon & " and id14=" & Val(Text1(0)) + 1911
   End If
   '公司
   If Text1(1) <> "" Then
      strExc(0) = "select a0807 from acc080 where a0801='" & Text1(1) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         stCon = stCon & " and id03='" & RsTemp(0) & "'"
      End If
   End If
   '所得人代號
   If Text1(2) <> "" Then
      stCon = stCon & " and id25='" & Text1(2) & "'"
   End If
   
'   'Add by Morgan 2009/5/25 補印有錯的
'   If Check1.Value = 1 Then
'      stCon = stCon & " and id05 in ('50','93') and substr(id17,1,10)>'0000000000'"
'   End If
   
   '所得人代號
   If Text1(3) = "Y" Then
      stCon = stCon & " and id05='54'"
   Else
      stCon = stCon & " and id05<>'54'"
   End If
   
   '翻譯人員抓相同ID的最小所內編號的部門及在職狀態
   'Modified by Morgan 2024/1/31 st03->decode(sign(id14-" & Left(新部門啟用日, 4) & "),-1,st03,st93)
   stVTB = "select id25 X1,substrb(min(lpad(st01,6,'0')||st04),7) X2,substrb(min(lpad(st01,6,'0')||st06),7) X3,substrb(min(lpad(st01,6,'0')||decode(sign(id14-" & Left(新部門啟用日, 4) & "),-1,st03,st93)),7) X4" & _
      " from IncomeData,staff where substr(id25,1,1)='F'" & stCon & _
      " and st26(+)=id06 and substrb(st01,1,1)<>'F'" & _
      " group by id25"
      
   '排序:在離職,部門,員工編號
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modified by Morgan 2012/2/3 修正外譯人員不會印地址問題
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   strExc(0) = "select nvl(decode(substr(id25,1,1),'F',X2,s2.st04),'2') stat,nvl(X3,s2.st06) comp,nvl(X4,id24) dep,id06 sid,a0802,a0804" & _
      ",s1.st02,oi03,s2.st04,a.*" & _
      " from IncomeData a,(" & stVTB & ") X,acc080,staff s1,OtherIncomer,staff s2" & _
      " where X1(+)=id25 and a0807(+)=id03 and s1.st01(+)=a0806 and OI01(+)=ID25" & _
      " and s2.st01(+)=substr(id25,1,2)||replace(substr(id25,3,1),'A','0')||substr(id25,4) " & stCon
   
   'Added by Morgan 2012/2/3
   If Text1(5) = "Y" Then
      strExc(0) = "select * from (" & strExc(0) & ") where stat='2'"
   End If
   'end 2012/2/3
   
   strExc(0) = strExc(0) & " order by 1 asc,2 asc,3 asc,4 asc"
   
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strFontSize = Printer.FontSize
      strFontName = Printer.FontName
      Printer.EndDoc
      
      With adoRst
      '只列印地址條
      If Text1(5) = "Y" Then GoTo PrintAddress
      
      Printer.PaperSize = PUB_GetPaperSize(3)
      Printer.Font = "細明體"
      
      Do While Not .EOF
         If .AbsolutePosition > 1 Then Printer.NewPage
         
         Printer.FontSize = 14
         yi = Yo + 1850
         '扣繳單位統編
         strExc(1) = Right(Val("" & .Fields("id03")) + 100000000, 8)
         xi = Xo + 2320 - Printer.TextWidth(strExc(1)) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '稽徵機關
         strExc(1) = "" & .Fields("id01")
         xi = Xo + 4050 - Printer.TextWidth(strExc(1)) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '製單編號(流水號+扣繳單位統編)
         strExc(1) = Right(Val("" & .Fields("id02")) + 100000000, 8) & Right(Val("" & .Fields("id03")) + 100000000, 8)
         xi = Xo + 6200 - Printer.TextWidth(strExc(1)) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '格式代號及所得類別
         strExc(1) = "" & .Fields("id05")
         If strExc(1) <> "54" Then
            xi = Xo + 8630 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
         
            yi = Yo + 2700
            '所得人統編
            strExc(1) = "" & .Fields("id06")
            xi = Xo + 2700 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)

            '有無住所(判斷其他所得人有無住所欄位)
            strExc(1) = "" & .Fields("oi03")
            '無
            If strExc(1) = "N" Then
               xi = Xo + 5050
            '有
            Else
               xi = Xo + 4200
            End If
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print "V"
            
            '所得人、執業...(共用欄位一,沒有的印所得人代號)
            strExc(1) = "" & .Fields("id11")
            If strExc(1) = "" Then
               strExc(1) = "" & .Fields("id25")
            End If
            xi = Xo + 6560 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
                        
            yi = Yo + 3170
            '所得人姓名
            strExc(1) = Trim("" & .Fields("id15"))
            'Modify by Morgan 2010/1/22
            'xi = Xo + 5215 - Printer.TextWidth(strExc(1)) / 2
            Printer.FontSize = 12
            xi = Xo + 4660 - Printer.TextWidth(strExc(1)) / 2
            
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            yi = Yo + 3570
            '所得人地址
            strExc(1) = Trim("" & .Fields("id16"))
            'Modify by Morgan 2010/1/22
            'xi = Xo + 6570 - Printer.TextWidth(strExc(1)) / 2
            'Do While xi < Xo + 3800
            xi = Xo + 6010 - Printer.TextWidth(strExc(1)) / 2
            Do While xi < Xo + 2800
               Printer.FontSize = Printer.FontSize - 1
               If Printer.FontSize < 9 Then
                  Printer.FontSize = 9
                  'Modify by Morgan 2010/1/22
                  'xi = Xo + 3700
                  xi = Xo + 2600
                  Exit Do
               End If
               'Modify by Morgan 2010/1/22
               'xi = Xo + 6570 - Printer.TextWidth(strExc(1)) / 2
               xi = Xo + 6010 - Printer.TextWidth(strExc(1)) / 2
            Loop
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            Printer.FontSize = 12
            yi = Yo + 4260
            '所得所屬年月
            '年
            strExc(1) = Val("" & .Fields("id14")) - 1911
            xi = Xo + 1870 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            '月
            strExc(1) = "" & .Fields("id22")
            xi = Xo + 2640 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            '年
            strExc(1) = Val("" & .Fields("id14")) - 1911
            xi = Xo + 3680 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            '月
            strExc(1) = "" & .Fields("id23")
            xi = Xo + 4500 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            Printer.FontSize = 14
            yi = Yo + 4320
            '所得給付年度
            strExc(1) = Val("" & .Fields("id14")) - 1911
            xi = Xo + 6470 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            '退休自提金額(薪資50,退職93)
            If .Fields("id05") = "50" Or .Fields("id05") = "93" Then
               strExc(1) = Format(Left("" & .Fields("id17"), 10), "#,###")
               xi = Xo + 8620 - Printer.TextWidth(strExc(1)) / 2
               Printer.CurrentX = xi: Printer.CurrentY = yi
               Printer.Print strExc(1)
            End If
                     
            yi = Yo + 5020
            '扣繳稅額
            strExc(1) = Format("" & .Fields("id09"), "#,###")
            xi = Xo + 6480 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            yi = Yo + 5230
            '給付總額
            strExc(1) = Format("" & .Fields("id08"), "#,###")
            xi = Xo + 2620 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            '給付淨額
            strExc(1) = Format("" & .Fields("id10"), "#,###")
            xi = Xo + 8640 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            'Modify by Morgan 2010/1/22
            'Printer.FontSize = 14
            Printer.FontSize = 12
            yi = Yo + 6250
            '扣繳單位名稱
            strExc(1) = Trim("" & .Fields("a0802"))
            xi = Xo + 4420 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            yi = Yo + 6720
            '扣繳單位地址
            strExc(1) = Trim("" & .Fields("a0804"))
            xi = Xo + 4420 - Printer.TextWidth(strExc(1)) / 2
            Do While xi < Xo + 1925
               Printer.FontSize = Printer.FontSize - 1
               If Printer.FontSize < 9 Then
                  Printer.FontSize = 9
                  xi = Xo + 1825
                  Exit Do
               End If
               xi = Xo + 4420 - Printer.TextWidth(strExc(1)) / 2
            Loop
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            'Add by Morgan 2010/1/22
            '格式代號說明
            Printer.FontSize = 12
            strExc(1) = GetDesc("" & .Fields("id05"))
            xi = Xo + 8200 - Printer.TextWidth(strExc(1)) / 2
            Do While xi < Xo + 7000
               Printer.FontSize = Printer.FontSize - 1
               If Printer.FontSize < 9 Then
                  Printer.FontSize = 9
                  xi = Xo + 7000
                  Exit Do
               End If
               xi = Xo + 8200 - Printer.TextWidth(strExc(1)) / 2
            Loop
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            
            'Modify by Morgan 2010/1/22
            'Printer.FontSize = 14
            Printer.FontSize = 12
            yi = Yo + 7115
            '扣繳義務人
            strExc(1) = Trim("" & .Fields("st02"))
            xi = Xo + 4420 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
         '股利
         Else
         
            yi = Yo + 2700
            '所得人統編
            strExc(1) = "" & .Fields("id06")
            xi = Xo + 2445 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            '有無住所(判斷其他所得人有無住所欄位)
            strExc(1) = "" & .Fields("oi03")
            '無
            If strExc(1) = "N" Then
               xi = Xo + 5960
            '有
            Else
               xi = Xo + 4600
            End If
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print "V"
            
            '所得人代號
            strExc(1) = "" & .Fields("id25")
            xi = Xo + 8035 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
                        
            yi = Yo + 3170
            
            Printer.FontSize = 12 'Add by Morgan 2010/1/22
            
            '所得人姓名
            strExc(1) = Trim("" & .Fields("id15"))
            xi = Xo + 5895 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            yi = Yo + 3670
            '所得人地址
            strExc(1) = Trim("" & .Fields("id16"))
            xi = Xo + 5895 - Printer.TextWidth(strExc(1)) / 2
            Do While xi < Xo + 2500
               Printer.FontSize = Printer.FontSize - 1
               If Printer.FontSize < 9 Then
                  Printer.FontSize = 9
                  xi = Xo + 2300
                  Exit Do
               End If
               xi = Xo + 5895 - Printer.TextWidth(strExc(1)) / 2
            Loop
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
                        
            Printer.FontSize = 14
            yi = Yo + 4520
            '所得所屬年度
            strExc(1) = Val("" & .Fields("id14")) - 1911
            xi = Xo + 3520 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            '所得給付年度
            xi = Xo + 6730 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            '稅額扣抵比率
            strExc(1) = Format(Left("" & .Fields("id11"), 4) / 100)
            xi = Xo + 8620 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            yi = Yo + 5300
            '股利總額
            strExc(1) = Format("" & .Fields("id08"), "#,###")
            xi = Xo + 2600 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi + 100
            Printer.Print strExc(1)
            
            '可扣抵稅額
            strExc(1) = Format("" & .Fields("id09"), "#,###")
            xi = Xo + 5270 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi + 100
            Printer.Print strExc(1)

            'Modify by Morgan 2010/1/21 調整位置,+印現金股利
            Printer.FontSize = 12
            '股利淨額
            strExc(1) = Format("" & .Fields("id10"), "#,###")
            xi = Xo + 7950 - Printer.TextWidth(strExc(1)) / 2
            'Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.CurrentX = xi: Printer.CurrentY = yi - 100
            Printer.Print strExc(1)
            
            Printer.FontSize = 10
            
            '資本公積股利
            strExc(1) = 0
            xi = Xo + 7100 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi + 400
            Printer.Print strExc(1)
            
            '現金股利
            strExc(1) = Format("" & .Fields("id10"), "#,###")
            xi = Xo + 7950 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi + 400
            Printer.Print strExc(1)
            
            '股票股利
            strExc(1) = 0
            xi = Xo + 8850 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi + 400
            Printer.Print strExc(1)
            
            'Modify by Morgan 2010/1/22
            'Printer.FontSize = 14
            Printer.FontSize = 12
            
            yi = Yo + 6150
            '扣繳單位名稱
            strExc(1) = Trim("" & .Fields("a0802"))
            xi = Xo + 4170 - Printer.TextWidth(strExc(1)) / 2
            'Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.CurrentX = xi: Printer.CurrentY = yi + 200
            Printer.Print strExc(1)
            
            yi = Yo + 6610
            '扣繳單位地址
            strExc(1) = Trim("" & .Fields("a0804"))
            xi = Xo + 4170 - Printer.TextWidth(strExc(1)) / 2
            Do While xi < Xo + 1900
               Printer.FontSize = Printer.FontSize - 1
               If Printer.FontSize < 9 Then
                  Printer.FontSize = 9
                  xi = Xo + 1880
                  Exit Do
               End If
               xi = Xo + 4170 - Printer.TextWidth(strExc(1)) / 2
            Loop
            'Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.CurrentX = xi: Printer.CurrentY = yi + 120
            Printer.Print strExc(1)
            
            'Modify by Morgan 2010/1/22
            'Printer.FontSize = 14
            Printer.FontSize = 12
            
            yi = Yo + 7020
            '扣繳義務人
            strExc(1) = Trim("" & .Fields("st02"))
            xi = Xo + 4170 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi + 40
            Printer.Print strExc(1)
            'end 2010/1/21
         End If
         .MoveNext
      Loop
      Printer.EndDoc
      
PrintAddress:
      If Text1(4) = "Y" Then
         .MoveFirst
         .Find "stat='2'"
         If Not .EOF Then
            If MsgBox("準備列印地址條，請放入正確的印表紙！", vbOKCancel + vbDefaultButton2) = vbOK Then
               frm170205.Hide
               Do While Not .EOF
                  If .Fields("stat") = "2" Then
                     If strLstID <> "" & .Fields("id06") Then
                        With frm170205
                        .FormReset
                        .Text1(0) = "" & adoRst.Fields("id25")
                        .Text1(1) = "" & adoRst.Fields("id25")
                        .Text1(5) = "Y"
                        .cmbPrinter = Me.cmbPrinter.Text
                        .m_bolBeCalled = True
                        .m_bolRegAddr = True 'Add by Morgan 2011/2/17
                        .m_bolVPrint = True 'Add by Morgan 2011/2/17
                        .PrintSheet
                        End With
                        strLstID = "" & .Fields("id06")
                     End If
                  End If
                  .MoveNext
               Loop
               Unload frm170205
            End If
         End If
      End If
      
      End With
      Printer.FontSize = strFontSize
      Printer.FontName = strFontName
      
      MsgBox "列印結束 !"
   Else
      MsgBox "無資料可列印 !"
   End If
   Set adoRst = Nothing
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   Dim oText As TextBox
   
   If Text1(0) = "" Then
      MsgBox "年度不可空白 !"
      Text1(0).SetFocus
      Exit Function
   End If
   For Each oText In Text1
      Text1_Validate oText.Index, bCancel
      If bCancel = True Then
         Text1(oText.Index).SetFocus
         Text1_GotFocus oText.Index
         Exit Function
      End If
   Next
   TxtValidate = True
End Function
'Add by Morgan 2010/1/22
Private Function GetDesc(stCode As String) As String
   Select Case stCode
      Case "50"
         GetDesc = "薪資"
      Case "51"
         GetDesc = "租賃"
      Case "53"
         GetDesc = "權利金"
      Case "5A"
         GetDesc = "金融業利息"
      Case "5B"
         GetDesc = "其他利息"
      Case "54", "54F"
         GetDesc = "股利或盈餘"
      Case "54Y"
         GetDesc = "其他營利所得"
      Case "9A"
         GetDesc = "執行業務"
      Case "9B"
         GetDesc = "稿費及演講鐘點費等七項"
      Case "91"
         GetDesc = "競技競賽及機會中獎獎金"
      Case "92"
         GetDesc = "其他"
      Case "93"
         GetDesc = "退職所得"
      Case "94"
         GetDesc = "行使員工認股權證所得"
      Case "95"
         GetDesc = "政府補助款"
   End Select
   GetDesc = stCode & " " & GetDesc
End Function
