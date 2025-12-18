VERSION 5.00
Begin VB.Form frm170201 
   BorderStyle     =   1  '單線固定
   Caption         =   "敘薪/換敘通知單"
   ClientHeight    =   1752
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3804
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1752
   ScaleWidth      =   3804
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   1080
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   1320
      Width           =   2460
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Index           =   3
      Left            =   2190
      MaxLength       =   6
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Index           =   2
      Left            =   1290
      MaxLength       =   6
      TabIndex        =   2
      Top             =   960
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Index           =   1
      Left            =   2190
      MaxLength       =   7
      TabIndex        =   1
      Text            =   "960131"
      Top             =   630
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Index           =   0
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   0
      Text            =   "960101"
      Top             =   630
      Width           =   765
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   2595
      TabIndex        =   5
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Left            =   1515
      TabIndex        =   6
      Top             =   90
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "印表機："
      Height          =   180
      Left            =   330
      TabIndex        =   9
      Top             =   1350
      Width           =   720
   End
   Begin VB.Line Line2 
      X1              =   1800
      X2              =   2460
      Y1              =   1050
      Y2              =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   330
      TabIndex        =   8
      Top             =   990
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   1800
      X2              =   2460
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "異動日期："
      Height          =   180
      Left            =   330
      TabIndex        =   7
      Top             =   660
      Width           =   900
   End
End
Attribute VB_Name = "frm170201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/4/18 改成Form2.0 (以圖片方式列印Unicode文字)
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Modified by Morgan 2024/1/31 改新部門順便刪除很多沒在用的舊程式碼,有需要的話看備份
'Create by Morgan 2009/2/6
Option Explicit

Dim m_DefaultPrinter As String
Dim iLineHeight As Integer, TwPerCm As Integer

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
   Screen.MousePointer = vbHourglass
   If TxtValidate = True Then
      Me.Enabled = False
      If cmbPrinter <> Printer.DeviceName Then PUB_RestorePrinter cmbPrinter
      
      PrintSheet
      
      '若印表機變動, 則更新列印設定
      If cmbPrinter.Tag <> cmbPrinter Then
         PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, 0, 0, Me.cmbPrinter.Text
      End If
      If Printer.DeviceName <> m_DefaultPrinter Then PUB_RestorePrinter m_DefaultPrinter
      Me.Enabled = True
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Function TxtValidate() As Boolean
   If Text1(0) = "" Then
      MsgBox "異動日期起日不可空白 !"
      Text1(0).SetFocus
      Exit Function
   ElseIf ChkDate(Text1(0)) = False Then
      Text1_GotFocus 0
      Text1(0).SetFocus
      Exit Function
   End If
   
   If Text1(1) = "" Then
      MsgBox "異動日期迄日不可空白 !"
      Text1_GotFocus 1
      Text1(1).SetFocus
      Exit Function
   ElseIf ChkDate(Text1(1)) = False Then
      Text1(1).SetFocus
      Exit Function
   End If
   TxtValidate = True
End Function

Private Sub PrintSheet()
   Dim stCon As String
   Dim strFontSize As String, strFontName As String
   Dim Xo As Integer, Yo As Integer, xi As Long, yi As Long
   Dim iWidth As Integer, iMargin As Integer
   Dim Px() As Integer, Py() As Integer
   Dim Header1 As String, Header2 As String
   Dim adoRst As ADODB.Recordset
   Dim DepOld As String, PosOld As String, TitOld As String, Reason As String
   Dim Comp1Old As String, Comp2Old As String, SL(11 To 39) As Long
   
   
   stCon = ""
   If Text1(0) <> "" Then
      stCon = stCon & " and sl02>=" & (Val(Text1(0)) + 19110000)
   End If

   If Text1(1) <> "" Then
      stCon = stCon & " and sl02<=" & (Val(Text1(1)) + 19110000)
   End If

   If Text1(2) <> "" Then
      stCon = stCon & " and sl01>='" & Text1(2) & "'"
   End If
   
   If Text1(3) <> "" Then
      stCon = stCon & " and sl01<='" & Text1(3) & "'"
   End If

   'Rem by Morgan 2009/5/27 不管是否調薪都要印
   'modify by sonia 2016/3/21 +st28試用期起日
   'Modified by Morgan 2024/1/31 +新部門acc090new
   strExc(0) = "select decode(sign(sl02-" & 新部門啟用日 & "),-1,st03,st93) st03,st02,a1.a0802 comp1,a2.a0802 comp2,decode(sign(sl02-" & 新部門啟用日 & "),-1,a0902,a0922) a0902,sc02" & _
      ",c1.ac03 Tit,c2.ac03 Pos,c3.ac03 Reason,x.*,st28" & _
      " from salarylog x,staff,acc080 a1,acc080 a2,acc090,acc090new" & _
      ",staff_change,allcode c1,allcode c2,allcode c3" & _
      " where st01(+)=sl01" & stCon & _
      " and a1.a0801(+)=sl33 and a2.a0801(+)=sl34 and a0901(+)=st03 and a0921(+)=st93" & _
      " and sc01(+)=sl01 and sc02(+)=sl02" & _
      " and c1.ac02(+)=st20 and c1.ac01(+)='01'" & _
      " and c2.ac02(+)=st21 and c2.ac01(+)='02'" & _
      " and c3.ac02(+)=sc03 and c3.ac01(+)='05'" & _
      " order by 1,sl01"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strFontSize = Printer.FontSize
      strFontName = Printer.FontName
      Printer.EndDoc
      Printer.PaperSize = 9
      Printer.Orientation = 1
      Printer.Font = "細明體"
      Printer.FontSize = 12
                 
      TwPerCm = 567
      iMargin = 150
      iWidth = 10300
      iLineHeight = Printer.TextHeight("字") + 45 '30
      
      Xo = (Printer.ScaleWidth - iWidth) / 2
      Yo = 500
      With adoRst
      Do While Not .EOF
         If .AbsolutePosition > 1 Then Printer.NewPage
         
         Printer.FontSize = 20
         Printer.FontBold = True
         yi = Yo + 750
         
         yi = yi + 2 * iLineHeight
         strExc(1) = "" & .Fields("comp1")
         xi = Xo + iWidth / 2 - Printer.TextWidth(strExc(1)) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
'2010/4/7 移到下面印
'         strExc(1) = "" & .Fields("comp1")
'         xi = Xo + iWidth / 2 - Printer.TextWidth(strExc(1)) / 2
'         Printer.CurrentX = xi: Printer.CurrentY = yi
'         Printer.Print strExc(1)
'2010/4/7 END
         
         Printer.FontSize = 16 '14
         'yi = yi + 2 * iLineHeight
         yi = yi + 3 * iLineHeight
         If .Fields("sl03") & .Fields("sl35") = "TN" Then
'2010/4/7 CANCEL BY SONIA
'            '所長
'            If Header1 = "" Then
'               strExc(0) = "select st02 from staff where st04='1' and st20='11'"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  Header1 = "" & RsTemp.Fields(0)
'               End If
'            End If
'            '總經理
'            If Header2 = "" Then
'               strExc(0) = "select st02 from staff where st04='1' and st20='21'"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  Header2 = "" & RsTemp.Fields(0)
'               End If
'            End If
'2010/4/7 END
            
            ReDim Px(15) As Integer
            ReDim Py(15) As Integer '11
            
            strExc(1) = "敘　薪　通　知　單"
            xi = Xo + iWidth / 2 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            Printer.FontSize = 12
            Printer.FontBold = False
            '2010/4/7 從下面移上來
            'yi = Py(11) - 1.5 * iLineHeight
            'xi = Px(0) + iMargin
            'yi = yi + 1.5 * iLineHeight
            yi = yi + 2.5 * iLineHeight
            'xi = Xo + iMargin
            xi = 6500
            Printer.CurrentX = xi: Printer.CurrentY = yi
            'modify by sonia 2019/11/26 incCNV_CHINESE_MINKO改用incCNV_CHINESE_MINKO1
            strExc(1) = "中華民國" & TranslateKeyWord(incCNV_CHINESE_MINKO1, strSrvDate(1), "")
            HPrint strExc(1)
            '2010/4/7 END
            
            yi = yi + 1.5 * iLineHeight
            'xi = Xo
            xi = 2400
            '畫表格
            Px(0) = xi: Py(0) = yi
            PrintTable1 Px, Py
            
            yi = Py(0) + 0.5 * iLineHeight
            xi = Px(0) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            strExc(1) = "員工編號"
            HPrint strExc(1)
            
            xi = Px(3) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print "" & .Fields("sl01")
            
            xi = Px(5) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            strExc(1) = "姓　名"
            HPrint strExc(1)
            
            xi = Px(8) + iMargin
            strExc(1) = "" & .Fields("st02")
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint strExc(1)
            
            yi = Py(1) + 0.5 * iLineHeight
            'xi = Px(10) + iMargin
            xi = Px(0) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            '2011/9/13 modify by sonia
            'strExc(1) = "任用日期"
            '2016/3/21 modify by sonia 辜說無試用期改印日期即可 A5001葉子寧
            'strExc(1) = "試用日期"
            If "" & .Fields("st28") <> "" Then
               strExc(1) = "試用日期"
            Else
               strExc(1) = "日　　期"
            End If
            '2016/3/21 end
            HPrint strExc(1)
            
            xi = Px(3) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print ChangeWStringToTDateString("" & .Fields("sl02"))
            
            yi = Py(2) + 0.5 * iLineHeight
            xi = Px(0) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            strExc(1) = "服務單位"
            HPrint strExc(1)
            
            xi = Px(3) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "" & .Fields("a0902")
            
            yi = Py(3) + 0.5 * iLineHeight
            'xi = Px(9) + iMargin
            xi = Px(0) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "所任職務"
            
            xi = Px(4) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "職位"
            
            xi = Px(5) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            'modify by sonia 2014/9/19 加"實習"二字-辜
            'HPrint "" & .Fields("Pos")
            '2016/3/21 modify by sonia 辜說無試用期者不印"實習"二字 A5001葉子寧
            'HPrint "實習" & .Fields("Pos")
            If "" & .Fields("st28") <> "" Then
               HPrint "實習" & .Fields("Pos")
            Else
               HPrint "" & .Fields("Pos")
            End If
            '2016/3/21 end
            
'            yi = Py(2) + 0.5 * iLineHeight
'            xi = Px(9) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            HPrint "職務"
            
            yi = Py(4) + 0.5 * iLineHeight
            xi = Px(4) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "職稱"
            
            xi = Px(5) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "" & .Fields("Tit")
            
            yi = Py(5) + 0.5 * iLineHeight
'            xi = Px(0) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            VPrint "" & .Fields("comp1")
            
            xi = Px(0) + iMargin + 400
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "本　薪"
            
            strExc(2) = Format(Val("" & .Fields("sl11")) + Val("" & .Fields("sl14")), "#,###")
            xi = Px(6) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
'            xi = Px(6) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            VPrint "" & .Fields("comp2")
'
'            xi = Px(7) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            HPrint "本　薪"
'
'            strExc(2) = Format(Val("" & .Fields("sl19")) + Val("" & .Fields("sl22")), "#,###")
'            xi = Px(11) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            xi = Px(12) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            HPrint "本　薪"
'
'            strExc(2) = Format(Val("" & .Fields("sl11")) + Val("" & .Fields("sl14")) + Val("" & .Fields("sl19")) + Val("" & .Fields("sl22")), "#,###")
'            xi = Px(15) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
                       
            yi = Py(6) + 0.5 * iLineHeight
            xi = Px(1) + iMargin + 100
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "職　務"
            
            strExc(1) = Format(Val("" & .Fields("sl12")), "#,###")
            xi = Px(6) - iMargin - Printer.TextWidth(strExc(1))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
'            xi = Px(8) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            HPrint "職務"
'
'            strExc(1) = Format(Val("" & .Fields("sl20")), "#,###")
'            xi = Px(11) - iMargin - Printer.TextWidth(strExc(1))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(1)
            
'            xi = Px(11) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            VPrint "總　　　　　　　　計"
'
'            xi = Px(13) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            HPrint "職務"
'
'            strExc(2) = Format(Val("" & .Fields("sl12")) + Val("" & .Fields("sl20")), "#,###")
'            xi = Px(15) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
            
            yi = Py(7)
            xi = Px(0) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            VPrint "津　　　　　　貼"
            
'            xi = Px(7) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            VPrint "津　　　　　貼"
'
'            xi = Px(12) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            VPrint "津　　　　　貼"

            yi = Py(7) + 0.5 * iLineHeight
            xi = Px(1) + iMargin + 100
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "證　照" '"技術"  'cancel by sonia 2020/5/6 技術改技術/證照,分二行
            
            'strExc(2) = Format(Val("" & .Fields("sl13")), "#,###")
            strExc(2) = Format(Val("" & .Fields("sl39")), "#,###")
            xi = Px(6) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
'            xi = Px(8) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            strExc(1) = "證照" '"技術證照"
'            HPrint strExc(1)  'cancel by sonia 2020/5/6 技術改技術/證照,分二行
'
'            'strExc(2) = Format(Val("" & .Fields("sl21")), "#,###")
'            strExc(2) = ""
'            xi = Px(11) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            xi = Px(13) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            strExc(1) = "證照" '"技術證照"
'            HPrint strExc(1)  'cancel by sonia 2020/5/6 技術改技術/證照,分二行
'
'            'strExc(2) = Format(Val("" & .Fields("sl13")) + Val("" & .Fields("sl21")), "#,###")
'            strExc(2) = Format(Val("" & .Fields("sl39")), "#,###")
'            xi = Px(15) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)

''add by sonia 2020/5/7 技術改技術/證照,分二行
'            Printer.FontSize = 11
'            yi = Py(5) + 0.25 * iLineHeight
'            xi = Px(2) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            HPrint "技術/"
'
'            xi = Px(8) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            strExc(1) = "技術/"
'            HPrint strExc(1)
'
'            xi = Px(13) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            strExc(1) = "技術/"
'            HPrint strExc(1)
'
'            yi = Py(5) + 1 * iLineHeight
'            xi = Px(2) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            HPrint "證照"
'
'            xi = Px(8) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            strExc(1) = "證照"
'            HPrint strExc(1)
'
'            xi = Px(13) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            strExc(1) = "證照"
'            HPrint strExc(1)
'            Printer.FontSize = 12
''end 2020/5/7
            
            yi = Py(8) + 0.5 * iLineHeight
            xi = Px(1) + iMargin + 100
            Printer.CurrentX = xi: Printer.CurrentY = yi
            strExc(1) = "差　旅"
            HPrint strExc(1)
            
            strExc(2) = Format(Val("" & .Fields("sl15")), "#,###")
            '2013/10/9 ADD BY SONIA 智權部同仁若差旅費無金額差旅費欄註明 "實報實銷"
            If Left("" & .Fields("st03"), 1) = "S" And "" & .Fields("st03") > "S10" Then
               If Val("" & .Fields("sl15")) = 0 Then strExc(2) = "實報實銷"
            End If
            '2013/10/9 END
            xi = Px(6) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
'            xi = Px(8) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            strExc(1) = "差旅"
'            HPrint strExc(1)
'
'            strExc(2) = Format(Val("" & .Fields("sl23")), "#,###")
'            '2013/10/9 ADD BY SONIA 智權部同仁若差旅費無金額差旅費欄註明 "實報實銷"
'            'MODIFY BY SONIA 2015/4/22 有本薪資料才印實報實銷
'            If Left("" & .Fields("st03"), 1) = "S" And "" & .Fields("st03") > "S10" And Val("" & .Fields("sl19")) + Val("" & .Fields("sl22")) > 0 Then
'               If Val("" & .Fields("sl23")) = 0 Then strExc(2) = "實報實銷"
'            End If
'            '2013/10/9 END
'            xi = Px(11) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            xi = Px(13) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            strExc(1) = "差旅"
'            HPrint strExc(1)
'
'            strExc(2) = Format(Val("" & .Fields("sl15")) + Val("" & .Fields("sl23")), "#,###")
'            '2013/10/9 ADD BY SONIA 智權部同仁若差旅費無金額差旅費欄註明 "實報實銷"
'            If Left("" & .Fields("st03"), 1) = "S" And "" & .Fields("st03") > "S10" Then
'               If Val("" & .Fields("sl15")) + Val("" & .Fields("sl23")) = 0 Then strExc(2) = "實報實銷"
'            End If
'            '2013/10/9 END
'            xi = Px(15) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
            
            yi = Py(9) + 0.5 * iLineHeight
            xi = Px(1) + iMargin + 100
            Printer.CurrentX = xi: Printer.CurrentY = yi
            strExc(1) = "房　租"
            HPrint strExc(1)
            
            strExc(2) = Format(Val("" & .Fields("sl16")), "#,###")
            xi = Px(6) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
'            xi = Px(8) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            strExc(1) = "房租"
'            HPrint strExc(1)
'
'            strExc(1) = Format(Val("" & .Fields("sl24")), "#,###")
'            xi = Px(11) - iMargin - Printer.TextWidth(strExc(1))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'
'            xi = Px(13) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            strExc(1) = "房租"
'            HPrint strExc(1)
'
'            strExc(2) = Format(Val("" & .Fields("sl16")) + Val("" & .Fields("sl24")), "#,###")
'            xi = Px(15) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
            
            
            yi = Py(10) + 0.5 * iLineHeight
            xi = Px(1) + iMargin + 100
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "技　術"
            
            strExc(2) = Format(Val("" & .Fields("sl13")), "#,###")
            xi = Px(6) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
'            xi = Px(8) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            strExc(1) = "技術"
'            HPrint strExc(1)
'
'            strExc(2) = Format(Val("" & .Fields("sl21")), "#,###")
'            xi = Px(11) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            xi = Px(13) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            strExc(1) = "技術"
'            HPrint strExc(1)
'
'            strExc(2) = Format(Val("" & .Fields("sl13")) + Val("" & .Fields("sl21")), "#,###")
'            xi = Px(15) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
            
            
            'yi = Py(8) + 0.5 * iLineHeight
            yi = Py(11) + 0.5 * iLineHeight
            xi = Px(1) + iMargin + 100
            Printer.CurrentX = xi: Printer.CurrentY = yi
            strExc(1) = "特　支"
            HPrint strExc(1)
            
            strExc(2) = Format(Val("" & .Fields("sl17")), "#,###")
            xi = Px(6) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
'            xi = Px(8) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            strExc(1) = "特支"
'            HPrint strExc(1)
'
'            strExc(2) = Format(Val("" & .Fields("sl25")), "#,###")
'            xi = Px(11) - iMargin - Printer.TextWidth(strExc(2))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(2)
'
'            xi = Px(13) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            strExc(1) = "特支"
'            HPrint strExc(1)
'
'            strExc(1) = Format(Val("" & .Fields("sl17")) + Val("" & .Fields("sl25")), "#,###")
'            xi = Px(15) - iMargin - Printer.TextWidth(strExc(1))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(1)
            
            
            'yi = Py(9) + 0.5 * iLineHeight
            yi = Py(12) + 0.5 * iLineHeight
            xi = Px(0) + iMargin + 400
            Printer.CurrentX = xi: Printer.CurrentY = yi
            strExc(1) = "合　計"
            HPrint strExc(1)
            
            strExc(2) = Format(Val("" & .Fields("sl11")) + Val("" & .Fields("sl12")) + Val("" & .Fields("sl39")) + Val("" & .Fields("sl13")) + Val("" & .Fields("sl14")) + Val("" & .Fields("sl15")) + Val("" & .Fields("sl16")) + Val("" & .Fields("sl17")), "#,###")
            xi = Px(6) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
'            xi = Px(7) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            strExc(1) = "合　計"
'            HPrint strExc(1)
'
'            strExc(3) = Format(Val("" & .Fields("sl19")) + Val("" & .Fields("sl20")) + Val("" & .Fields("sl21")) + Val("" & .Fields("sl22")) + Val("" & .Fields("sl23")) + Val("" & .Fields("sl24")) + Val("" & .Fields("sl25")), "#,###")
'            xi = Px(11) - iMargin - Printer.TextWidth(strExc(3))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(3)
'
'            xi = Px(12) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            strExc(1) = "合　計"
'            HPrint strExc(1)
'
'            strExc(4) = Format(Val(Format(strExc(2))) + Val(Format(strExc(3))), "#,###")
'            xi = Px(15) - iMargin - Printer.TextWidth(strExc(4))
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(4)
            
'2010/4/7 MODIBY BY SONIA 把公司別移下來,取消負責人改以蓋章方式
'            yi = Py(10) + 1.5 * iLineHeight
'            xi = Px(7) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            If .Fields("st03") = "R04" Then
'               strExc(1) = "董事長　　何金柱"
'               HPrint strExc(1)
'            Else
'               strExc(1) = "所　長　　" & Header1
'               HPrint strExc(1)
'
'               yi = Py(10) + 3.5 * iLineHeight
'               xi = Px(7) + iMargin
'               Printer.CurrentX = xi: Printer.CurrentY = yi
'               strExc(1) = "總經理　　" & Header2
'               HPrint strExc(1)
'            End If
            
            
'            Printer.FontSize = 14
'            Printer.FontBold = True
'            strExc(1) = "" & .Fields("comp1")
'            xi = Px(1) + iMargin
'            yi = yi + 5 * iLineHeight
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(1)
'            Printer.FontSize = 12
'            Printer.FontBold = False
'2010/4/7 END
'2010/4/7 CANCEL BY SONIA 移到上面去
'            yi = Py(11) - 1.5 * iLineHeight
'            xi = Px(0) + iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            strExc(1) = "中華民國" & TranslateKeyWord(incCNV_CHINESE_MINKO, strSrvDate(1), "")
'            HPrint strExc(1)
'2010/4/7 END

         Else
            '原人事資料
            '人事有異動
            If IsNull(.Fields("sc02")) = False Then
               '事由
               Reason = "" & .Fields("reason")
               '抓前次人事異動內容(部門,職稱,職位)
               'Modified by Morgan 2024/1/31 +新部門acc090new
               strExc(0) = "select decode(sign(sc02-" & 新部門啟用日 & "),-1, a0902,a0922) a0902,c1.ac03 Tit,c2.ac03 Pos" & _
                  " from staff_change a,acc090,acc090new,allcode c1,allcode c2" & _
                  " where sc01='" & .Fields("sl01") & "' and sc02" & _
                  "=(select max(b.sc02) from staff_change b where b.sc01=a.sc01 and b.sc02<" & .Fields("sc02") & ")" & _
                  " and a0901(+)=sc04 and a0921(+)=sc04" & _
                  " and c1.ac02(+)=sc05 and c1.ac01(+)='01'" & _
                  " and c2.ac02(+)=sc06 and c2.ac01(+)='02'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  DepOld = "" & RsTemp.Fields("a0902")
                  PosOld = "" & RsTemp.Fields("Pos")
                  TitOld = "" & RsTemp.Fields("Tit")
               Else
                  DepOld = ""
                  PosOld = ""
                  TitOld = ""
               End If
            '人事無異動
            Else
               '2011/9/13 modify by sonia
               'Reason = "調薪"
               Reason = "年度調薪"
               '2011/9/13 end
               DepOld = "" & .Fields("a0902")
               PosOld = "" & .Fields("Pos")
               TitOld = "" & .Fields("Tit")
            End If
            
            '原薪資資料
            Erase SL
            '抓前次薪資異動內容
            strExc(0) = "select a1.a0802 comp1,a2.a0802 comp2,x.*" & _
               " from salarylog x,acc080 a1,acc080 a2" & _
               " where sl01='" & .Fields("sl01") & "' and sl02" & _
               "=(select max(b.sl02) from salarylog b where b.sl01=x.sl01 and b.sl02<" & .Fields("sl02") & ")" & _
               " and a1.a0801(+)=sl33 and a2.a0801(+)=sl34"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               With RsTemp
                  Comp1Old = "" & .Fields("comp1")
                  Comp2Old = "" & .Fields("comp2")
                  'Modified by Morgan 2020/11/23
                  'For intI = 11 To 26
                  For intI = 11 To 39
                  'end 2020/11/23
                     SL(intI) = Val("" & .Fields("sl" & Format(intI, "00")))
                  Next
                  '2011/9/13 add by sonia 前次若為sl03='T'則改Reason
                  If "" & .Fields("sl03") = "T" Then Reason = "試用期滿調薪"
                  '2011/9/13 END
               End With
            End If
            
            ReDim Px(12) As Integer
            ReDim Py(20) As Integer
            
            strExc(1) = "換　敘　通　知　單"
            xi = Xo + iWidth / 2 - Printer.TextWidth(strExc(1)) / 2
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            
            Printer.FontSize = 12
            Printer.FontBold = False
            
            'yi = yi + 1.5 * iLineHeight
            yi = yi + 2.5 * iLineHeight
            'xi = Xo + iMargin
            xi = 7700
            Printer.CurrentX = xi: Printer.CurrentY = yi
            'modify by sonia 2019/11/26 incCNV_CHINESE_MINKO改用incCNV_CHINESE_MINKO1
            strExc(1) = "中華民國" & TranslateKeyWord(incCNV_CHINESE_MINKO1, strSrvDate(1), "")
            HPrint strExc(1)
            
            yi = yi + 1.5 * iLineHeight
            xi = Xo
            Px(0) = xi: Py(0) = yi
            PrintTable2 Px, Py
            
            yi = Py(0) + 0.5 * iLineHeight
            xi = Px(0) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "員工編號"
                        
            xi = Px(2) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print "" & .Fields("sl01")
            
            xi = Px(4) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "姓　名"
            
            xi = Px(5) + iMargin
                        
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "" & .Fields("st02")
            
                        
            '2011/9/13 modify by sonia
            'xi = Px(8) + iMargin
            'Printer.CurrentX = xi: Printer.CurrentY = yi
            'HPrint "事　由"
            xi = Px(7) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "事由"
            '2011/9/13 end
            
            '2011/9/13 modify by sonia
            'xi = Px(9) + iMargin
            xi = Px(8) + 400 + iMargin
            '2011/9/13 end
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint Reason
            
            yi = Py(1) + 0.5 * iLineHeight
            xi = Px(0) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "原服務單位"
                        
            xi = Px(2) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint DepOld
            
            xi = Px(5) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "現派調單位"
                        
            xi = Px(7) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "" & .Fields("a0902")
            
            yi = Py(2) + 0.5 * iLineHeight
            xi = Px(0) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "原任"
            
            xi = Px(1) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "職位"
            
            xi = Px(2) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint PosOld
            
            xi = Px(5) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "現調"
            
            xi = Px(6) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "職位"
            
            xi = Px(7) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "" & .Fields("Pos")
            
            yi = Py(3) + 0.5 * iLineHeight
            xi = Px(0) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "職務"
            
            xi = Px(1) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "職稱"
            
            xi = Px(2) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint TitOld
            
            xi = Px(5) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "職務"
            
            xi = Px(6) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "職稱"
            
            xi = Px(7) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "" & .Fields("Tit")
            
            yi = Py(4) + 0.5 * iLineHeight
            xi = Px(0) + 2 * iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            VPrint "原　　　支　　　薪　　　資"
            
'            xi = Px(1) + 2 * iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            VPrint Comp1Old
            
            xi = Px(2) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "本　　　薪"
            
            strExc(2) = Format(SL(11) + SL(14), "#,###")
            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            xi = Px(5) + 2 * iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            VPrint "核　　　定　　　薪　　　資"
            
'            xi = Px(6) + 2 * iMargin
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            VPrint "" & .Fields("comp1")
            
            xi = Px(7) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "本　　　薪"
            
            strExc(2) = Format(Val("" & .Fields("sl11")) + Val("" & .Fields("sl14")), "#,###")
            xi = Px(10) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            yi = Py(5) + 0.5 * iLineHeight
            xi = Px(3) '+ iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "職　務"
            
            strExc(2) = Format(SL(12), "#,###")
            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            xi = Px(8) '+ iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "職　務"
            
            strExc(2) = Format(Val("" & .Fields("sl12")), "#,###")
            xi = Px(10) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            yi = Py(6)
            'xi = Px(2) + iMargin
            xi = Px(1) + (iMargin * 2)
            Printer.CurrentX = xi: Printer.CurrentY = yi
            VPrint "津　　　　　　貼"
            
            'xi = Px(7) + iMargin
            xi = Px(6) + (iMargin * 2)
            Printer.CurrentX = xi: Printer.CurrentY = yi
            VPrint "津　　　　　　貼"
            
            yi = Py(6) + 0.5 * iLineHeight
            xi = Px(3) '+ iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "證　照"
            
            'strExc(2) = Format(SL(13), "#,###")
            strExc(2) = Format(SL(39), "#,###")
            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            xi = Px(8) '+ iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "證　照"
            
            'strExc(2) = Format(Val("" & .Fields("sl13")), "#,###")
            strExc(2) = Format(Val("" & .Fields("sl39")), "#,###")
            xi = Px(10) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            yi = Py(7) + 0.5 * iLineHeight
            xi = Px(3) '+ iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "差　旅"
            
            strExc(2) = Format(SL(15), "#,###")
            '2013/10/9 ADD BY SONIA 智權部同仁若差旅費無金額差旅費欄註明 "實報實銷"
            If Left("" & .Fields("st03"), 1) = "S" And "" & .Fields("st03") > "S10" Then
               If Val(SL(15)) = 0 Then strExc(2) = "實報實銷"
            End If
            '2013/10/9 END
            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            xi = Px(8) '+ iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "差　旅"
            
            strExc(2) = Format(Val("" & .Fields("sl15")), "#,###")
            '2013/10/9 ADD BY SONIA 智權部同仁若差旅費無金額差旅費欄註明 "實報實銷"
            If Left("" & .Fields("st03"), 1) = "S" And "" & .Fields("st03") > "S10" Then
               If Val("" & .Fields("sl15")) = 0 Then strExc(2) = "實報實銷"
            End If
            '2013/10/9 END
            xi = Px(10) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            yi = Py(8) + 0.5 * iLineHeight
            xi = Px(3) '+ iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "房　租"
            
            strExc(2) = Format(SL(16), "#,###")
            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            xi = Px(8) '+ iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "房　租"
            
            strExc(2) = Format(Val("" & .Fields("sl16")), "#,###")
            xi = Px(10) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            
            yi = Py(9) + 0.5 * iLineHeight
            xi = Px(3) '+ iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "技　術"
            
            strExc(2) = Format(SL(13), "#,###")
            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            xi = Px(8) '+ iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "技　術"
            
            strExc(2) = Format(Val("" & .Fields("sl13")), "#,###")
            xi = Px(10) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            
            yi = Py(10) + 0.5 * iLineHeight
            xi = Px(3) ' + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "特　支"
            
            strExc(2) = Format(SL(17), "#,###")
            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            xi = Px(8) '+ iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "特　支"
            
            strExc(2) = Format(Val("" & .Fields("sl17")), "#,###")
            xi = Px(10) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            yi = Py(11) + 0.5 * iLineHeight
            xi = Px(2) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "合　　　計"
            
            strExc(2) = Format(SL(11) + SL(12) + SL(13) + SL(39) + SL(14) + SL(15) + SL(16) + SL(17), "#,###")
            xi = Px(5) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
            
            xi = Px(7) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "合　　　計"
            
            strExc(2) = Format(Val("" & .Fields("sl11")) + Val("" & .Fields("sl12")) + Val("" & .Fields("sl13")) + Val("" & .Fields("sl39")) + Val("" & .Fields("sl14")) + Val("" & .Fields("sl15")) + Val("" & .Fields("sl16")) + Val("" & .Fields("sl17")), "#,###")
            xi = Px(10) - iMargin - Printer.TextWidth(strExc(2))
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(2)
           
            'yi = Py(19) + 0.5 * iLineHeight
            yi = Py(12) + 0.5 * iLineHeight
            'xi = Px(0) + 2 * iMargin
            xi = Px(0) + iMargin + 50
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "停支日期", 100
            
            xi = Px(2) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            strExc(1) = ChangeWStringToTDateString(CompDate(2, -1, "" & .Fields("sl02")))
            Printer.Print strExc(1)
            
            xi = Px(5) + iMargin + 50
            Printer.CurrentX = xi: Printer.CurrentY = yi
            HPrint "換敘日期", 100
            '具領日=換敘日(異動日)的次月5號
            xi = Px(7) + iMargin
            Printer.CurrentX = xi: Printer.CurrentY = yi
            strExc(1) = ChangeWStringToTDateString("" & .Fields("sl02"))
            strExc(1) = strExc(1) & " (" & ChangeWStringToTDateString(CompDate(1, 1, (Val("" & .Fields("sl02")) \ 100) & "05")) & " 具領)"
            Printer.Print strExc(1)
            
            '2010/4/7 ADD BY SONIA 自表頭移下來
'            Printer.FontSize = 14
'            Printer.FontBold = True
'            strExc(1) = "" & .Fields("comp1")
            xi = Px(0) + 2 * iMargin
            yi = yi + 3 * iLineHeight
'            Printer.CurrentX = xi: Printer.CurrentY = yi
'            Printer.Print strExc(1)
            Printer.FontSize = 12
            Printer.FontBold = False
            '2010/4/7 END
            
            '2010/5/4 ADD BY SONIA 加備註
            If "" & .Fields("sl37") <> "" Then
               xi = Px(5) + iMargin
               Printer.CurrentX = xi: Printer.CurrentY = yi
               '2012/10/31 MODIFY BY SONIA 移至右邊並分成二行
               Printer.Print "備註：" & Left(.Fields("sl37"), 20)
               If Mid(.Fields("sl37"), 21) <> "" Then
                  yi = yi + 1 * iLineHeight
                  xi = Px(5) + iMargin
                  Printer.CurrentX = xi: Printer.CurrentY = yi
                  Printer.Print "　　　ff" & Mid(.Fields("sl37"), 21)
               End If
            End If
            '2010/5/4 END
         End If
                 
         .MoveNext
      Loop
      Printer.EndDoc
      Printer.FontSize = strFontSize
      Printer.FontName = strFontName
      End With
      
      MsgBox "列印結束 !"
   Else
      MsgBox "無資料可列印 !"
   End If
   Set adoRst = Nothing
End Sub

Private Sub FormReset()
   Dim oText  As TextBox
   For Each oText In Text1
      oText.Text = Empty
   Next
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cmbPrinter, m_DefaultPrinter
   FormReset
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170201 = Nothing
End Sub


Private Sub Text1_GotFocus(Index As Integer)
   If Index = 1 Then
      If Text1(0) <> "" Then Text1(1) = Text1(0)
   ElseIf Index = 3 Then
      If Text1(2) <> "" Then Text1(3) = Text1(2)
   End If
   TextInverse Text1(Index)
   CloseIme
End Sub

'敘薪通知單表格
Private Sub PrintTable1(Px() As Integer, Py() As Integer)
   Px(1) = Px(0) + TwPerCm
   Px(2) = Px(1) + TwPerCm
   Px(3) = Px(2) + 1.5 * TwPerCm
   Px(4) = Px(2) + 1.5 * TwPerCm
   Px(5) = Px(4) + 2.5 * TwPerCm
   Px(6) = Px(4) + 2.5 * TwPerCm
   Px(7) = Px(6) + TwPerCm
   Px(8) = Px(7) + TwPerCm
   Px(9) = Px(8) + 1.5 * TwPerCm
   Px(10) = Px(9) + 1.5 * TwPerCm
   Px(11) = Px(10) + TwPerCm
   Px(12) = Px(11) + TwPerCm
   Px(13) = Px(12) + TwPerCm
   Px(14) = Px(13) + 1.5 * TwPerCm
   Px(15) = Px(14) + 2.5 * TwPerCm
   
   For intI = 1 To 14 '10
      Py(intI) = Py(intI - 1) + 2 * iLineHeight
   Next
   'Py(11) = Py(10) + 8 * iLineHeight
   'Py(12) = Py(11) + 8 * iLineHeight
   
   Printer.DrawWidth = 7
   
   '框
   'Printer.Line (Px(0), Py(0))-(Px(15), Py(11)), , B
   Printer.Line (Px(0), Py(0))-(Px(12), Py(13)), , B
   '縱線
   Printer.Line (Px(1), Py(6))-(Px(1), Py(12))
   Printer.Line (Px(3), Py(0))-(Px(3), Py(13))
   Printer.Line (Px(5), Py(0))-(Px(5), Py(1))
   Printer.Line (Px(8), Py(0))-(Px(8), Py(1))
   Printer.Line (Px(5), Py(3))-(Px(5), Py(5))
   
   '橫線
   Printer.Line (Px(0), Py(1))-(Px(12), Py(1))
   Printer.Line (Px(0), Py(2))-(Px(12), Py(2))
   
   Printer.Line (Px(0), Py(3))-(Px(12), Py(3))
   Printer.Line (Px(3), Py(4))-(Px(12), Py(4))
   
   Printer.Line (Px(0), Py(5))-(Px(12), Py(5))
   Printer.Line (Px(0), Py(6))-(Px(12), Py(6))
   
   For intI = 7 To 11
      Printer.Line (Px(1), Py(intI))-(Px(12), Py(intI))
   Next
   Printer.Line (Px(0), Py(12))-(Px(12), Py(12))
   
   Printer.DrawWidth = 1
End Sub

'換敘通知單表格
Private Sub PrintTable2(Px() As Integer, Py() As Integer)
   Px(1) = Px(0) + 1.5 * TwPerCm
   Px(2) = Px(1) + 1.5 * TwPerCm
   Px(3) = Px(2) + TwPerCm
   Px(4) = Px(3) + 2.5 * TwPerCm
   Px(5) = Px(4) + 2.5 * TwPerCm
   Px(6) = Px(5) + 1.5 * TwPerCm
   Px(7) = Px(6) + 1.5 * TwPerCm
   Px(8) = Px(7) + TwPerCm
   Px(9) = Px(8) + 2.5 * TwPerCm
   Px(10) = Px(9) + 2.5 * TwPerCm
      
   For intI = 1 To 13 '20
      Py(intI) = Py(intI - 1) + 2 * iLineHeight
   Next
            
   Printer.DrawWidth = 7
   
   '框
   'Printer.Line (Px(0), Py(0))-(Px(10), Py(20)), , B
   Printer.Line (Px(0), Py(0))-(Px(10), Py(13)), , B
   '縱線
   Printer.Line (Px(1), Py(2))-(Px(1), Py(12)) '19
   
   'Printer.Line (Px(2), Py(0))-(Px(2), Py(18))
   Printer.Line (Px(2), Py(0))-(Px(2), Py(4))
   Printer.Line (Px(2), Py(12))-(Px(2), Py(13))
   
   'Printer.Line (Px(3), Py(5))-(Px(3), Py(11)) '10
   Printer.Line (Px(2), Py(5))-(Px(2), Py(11)) '10
   
   'Printer.Line (Px(3), Py(12))-(Px(3), Py(10)) '17
   'Printer.Line (Px(3), Py(19))-(Px(3), Py(20))
   'Printer.Line (Px(3), Py(12))-(Px(3), Py(13))
   Printer.Line (Px(4), Py(0))-(Px(4), Py(1))
   Printer.Line (Px(4), Py(4))-(Px(4), Py(12)) '19
   Printer.Line (Px(5), Py(0))-(Px(5), Py(13)) '20
   Printer.Line (Px(6), Py(2))-(Px(6), Py(12)) '19
   
   'Printer.Line (Px(7), Py(0))-(Px(7), Py(18))
   Printer.Line (Px(7), Py(0))-(Px(7), Py(4))
   Printer.Line (Px(7), Py(12))-(Px(7), Py(13))
   
   'Printer.Line (Px(7), Py(19))-(Px(7), Py(12)) '20
   Printer.Line (Px(7), Py(12))-(Px(7), Py(13)) '20
   
   'Printer.Line (Px(8), Py(5))-(Px(8), Py(11)) '10
   Printer.Line (Px(7), Py(5))-(Px(7), Py(11)) '10
   
'   Printer.Line (Px(8), Py(12))-(Px(8), Py(17))
   Printer.Line (Px(8) + 400, Py(0))-(Px(8) + 400, Py(1))
   'Printer.Line (Px(9), Py(4))-(Px(9), Py(19))
   Printer.Line (Px(9), Py(4))-(Px(9), Py(12))
   
   '橫線
   Printer.Line (Px(0), Py(1))-(Px(10), Py(1))
   Printer.Line (Px(0), Py(2))-(Px(10), Py(2))
   Printer.Line (Px(1), Py(3))-(Px(5), Py(3))
   Printer.Line (Px(6), Py(3))-(Px(10), Py(3))
   
   Printer.Line (Px(0), Py(4))-(Px(10), Py(4))
   
   'Printer.Line (Px(2), Py(5))-(Px(5), Py(5))
   Printer.Line (Px(1), Py(5))-(Px(5), Py(5))
   
   'Printer.Line (Px(7), Py(5))-(Px(10), Py(5))
   Printer.Line (Px(6), Py(5))-(Px(10), Py(5))
   
   For intI = 6 To 11 '9
      'Printer.Line (Px(3), Py(intI))-(Px(5), Py(intI))
      Printer.Line (Px(2), Py(intI))-(Px(5), Py(intI))
      'Printer.Line (Px(8), Py(intI))-(Px(10), Py(intI))
      Printer.Line (Px(7), Py(intI))-(Px(10), Py(intI))
   Next
   Printer.Line (Px(3), Py(10))-(Px(5), Py(10))
   Printer.Line (Px(8), Py(10))-(Px(10), Py(10))
   
   'Printer.Line (Px(2), Py(11))-(Px(5), Py(11))
   Printer.Line (Px(1), Py(11))-(Px(5), Py(11))
   
   'Printer.Line (Px(7), Py(11))-(Px(10), Py(11))
   Printer.Line (Px(6), Py(11))-(Px(10), Py(11))
   
   Printer.Line (Px(0), Py(12))-(Px(5), Py(12))
   Printer.Line (Px(5), Py(12))-(Px(10), Py(12))
  
   Printer.DrawWidth = 1
End Sub


'垂直列印
Private Sub VPrint(sWords As String, Optional iGap As Integer = 50)
   Dim Xo As Integer, Yo As Integer, ii As Integer
   Xo = Printer.CurrentX
   Yo = Printer.CurrentY
   For ii = 1 To Len(sWords)
      Printer.CurrentX = Xo
      Printer.CurrentY = Yo + (ii - 1) * (Printer.TextHeight(Space(1)) + iGap)
      Printer.Print Mid(sWords, ii, 1)
   Next
End Sub

'水平列印
Private Sub HPrint(sWords As String, Optional iGap As Integer = 50)
   Dim Xo As Long, Yo As Long, ii As Integer
   Xo = Printer.CurrentX
   Yo = Printer.CurrentY
   'Modified by Morgan 2022/4/19 逐字檢查Unicode文字改以圖片方式列印
   'For ii = 1 To Len(sWords)
   '   Printer.CurrentX = Xo
   '   Printer.CurrentY = Yo
   '   Printer.Print Mid(sWords, ii, 1)
   '   Xo = Xo + Printer.TextWidth(Mid(sWords, ii, 1)) + iGap
   'Next
   PUB_PrintUnicodeText sWords, Xo, Yo, iGap
   'end 2022/4/19
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub


