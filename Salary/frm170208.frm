VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm170208 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工薪資單"
   ClientHeight    =   3708
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4164
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3708
   ScaleWidth      =   4164
   Begin VB.CheckBox Check5 
      Caption         =   "約定薪資翻譯人員(以EMail發送翻譯明細)"
      Height          =   255
      Left            =   225
      TabIndex        =   21
      Top             =   3096
      Width           =   3564
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H008080FF&
      Caption         =   "？"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3792
      Style           =   1  '圖片外觀
      TabIndex        =   20
      Top             =   2784
      Width           =   285
   End
   Begin VB.CheckBox Check4 
      Caption         =   "只EMail給自己(不會寄原收件人)"
      Height          =   255
      Left            =   225
      TabIndex        =   19
      Top             =   3408
      Width           =   3492
   End
   Begin VB.CheckBox Check3 
      Caption         =   "台一投資及離職同仁(以EMail發送薪資單)"
      Height          =   255
      Left            =   225
      TabIndex        =   17
      Top             =   2784
      Width           =   3540
   End
   Begin VB.TextBox Text5 
      Height          =   255
      Left            =   1125
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1140
      Width           =   495
   End
   Begin VB.CheckBox Check2 
      Caption         =   "A4格式"
      Height          =   255
      Left            =   180
      TabIndex        =   15
      Top             =   120
      Value           =   1  '核取
      Width           =   1500
   End
   Begin VB.CheckBox Check1 
      Caption         =   "只印要印薪資單者"
      Height          =   255
      Left            =   225
      TabIndex        =   14
      Top             =   2472
      Width           =   2175
   End
   Begin VB.TextBox txtSet 
      Height          =   285
      Index           =   2
      Left            =   1620
      TabIndex        =   6
      Top             =   2130
      Width           =   705
   End
   Begin VB.TextBox txtSet 
      Height          =   285
      Index           =   1
      Left            =   1620
      TabIndex        =   5
      Top             =   1830
      Width           =   705
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   945
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   1440
      Width           =   2955
   End
   Begin VB.TextBox Text3 
      Height          =   255
      Left            =   1140
      MaxLength       =   6
      TabIndex        =   2
      Top             =   840
      Width           =   765
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1890
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "5"
      Top             =   510
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "96"
      Top             =   510
      Width           =   435
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   2835
      TabIndex        =   8
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "執行(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1755
      TabIndex        =   7
      Top             =   60
      Width           =   975
   End
   Begin MSForms.Label lblName 
      Height          =   180
      Left            =   2016
      TabIndex        =   18
      Top             =   888
      Width           =   1332
      VariousPropertyBits=   8388627
      Caption         =   "姓名"
      Size            =   "2350;317"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "所別："
      Height          =   180
      Index           =   3
      Left            =   210
      TabIndex        =   16
      Top             =   1170
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "縱軸偏移值(Y)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   2
      Left            =   210
      TabIndex        =   13
      Top             =   2190
      Width           =   3240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "橫軸偏移值(X)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   12
      Top             =   1890
      Width           =   3240
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      Height          =   255
      Left            =   210
      TabIndex        =   11
      Top             =   1530
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "員工編號："
      Height          =   180
      Index           =   1
      Left            =   210
      TabIndex        =   10
      Top             =   870
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "薪資月份：            年         月"
      Height          =   180
      Left            =   210
      TabIndex        =   9
      Top             =   570
      Width           =   2205
   End
End
Attribute VB_Name = "frm170208"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by Morgan 2009/1/5
Option Explicit

Dim m_Actived As Boolean
Dim m_DefaultPrinter As String
Dim adoRst As ADODB.Recordset
Dim Xo As Integer, Yo As Integer, xi As Long, yi As Long
Dim dblUnitWidth As Double, dblUnitHeight As Double
Dim YM As String
Dim m_AttachPath As String

Private Sub SetCheck(pName As String)
   If Check1.Name = pName And Check1.Value = vbChecked Then
      Check5.Value = vbUnchecked
      Check3.Value = vbUnchecked
   ElseIf Check3.Name = pName And Check3.Value = vbChecked Then
      Check1.Value = vbUnchecked
      Check5.Value = vbUnchecked
   ElseIf Check5.Name = pName And Check5.Value = vbChecked Then
      Check1.Value = vbUnchecked
      Check3.Value = vbUnchecked
   End If
End Sub
Private Sub Check1_Click()
   SetCheck Check1.Name
End Sub

Private Sub Check3_Click()
   SetCheck Check3.Name
End Sub

Private Sub Check5_Click()
   SetCheck Check5.Name
End Sub

Private Sub cmdHelp_Click()
   MsgBoxU "離職同人發信條件：" & vbCrLf & "付款當月5號前離職者發信，但限入帳類別為薪資轉帳(2,4,5,6)者。" & vbCrLf & "(ex:10/4離職,會收到9&10月的薪資通知信.)", vbInformation
End Sub

Private Sub Mail2Translator()
   Dim stCon As String
   Dim lngAmtFee As Long, lngAmtThisMon As Long, lngAmtNextMon As Long
   Dim YM2 As String
   Dim stSubject As String, strContent As String
   Dim strPdfName As String, strMsg As String, strMsgOK As String, strMsgErr As String
   
   If Text3 <> "" Then
      stCon = stCon & " and sm01='" & Text3 & "'"
   End If
   
   strExc(0) = "select sm01,sm37,max(s1.st02) st02,max(s1.st18) st18,od03,nvl(sum(od05),0) od05 from salarymonth,staff s1,staff s2,othersalarydata" & _
      " where sm02=" & YM & " and substr(sm01,-2)>='9A' " & stCon & " and s1.st01(+)=sm01 and s2.st26(+)=s1.st26" & _
      " and od03(+)=s2.st01 and od02>=" & YM & "01 and od02<=" & YM & "31 and od04='01' group by sm01,sm37,od03"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      YM2 = Left(CompDate(1, 1, YM & "01"), 6)
      stSubject = Val(Text1) & "年" & Val(Text2) & "月翻譯費明細"
      With adoRst
      Do While Not .EOF
         lngAmtThisMon = 0
         lngAmtNextMon = 0
         lngAmtFee = .Fields("od05")
         '上月待補足金額(列本月翻譯預支)
         strExc(0) = "select nvl(sum(od05),0) from othersalarydata" & _
            " where od03='" & .Fields("sm01") & "' and od02>=" & YM & "01 and od02<=" & YM & "31 and od04='39'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            lngAmtThisMon = RsTemp(0)
         End If
         '本月待補足金額(列下月翻譯預支)
         strExc(0) = "select nvl(sum(od05),0) from othersalarydata" & _
            " where od03='" & .Fields("sm01") & "' and od02>=" & YM2 & "01 and od02<=" & YM2 & "31 and od04='39'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            lngAmtNextMon = RsTemp(0)
         End If
         strContent = "<DIV><FONT color=""black"">" & .Fields("st02") & "　您好," & vbCrLf & vbCrLf & _
            "本月翻譯費共計 " & Format(lngAmtFee, "#,##0") & " 元" & IIf(lngAmtFee > 0, ", 明細如附件", "") & "." & vbCrLf & _
            "其中," & vbCrLf & _
            (Left(YM2, 4) - 1911) & "年" & Val(Mid(YM2, 5)) & "月待補足金額為 " & Format(lngAmtNextMon, "#,##0") & " 元" & vbCrLf & _
            "(上月待補足 " & Format(lngAmtThisMon, "#,##0") & " 元 + 本薪 - 本月翻譯費 = " & Format(lngAmtNextMon, "#,##0") & " 元)" & vbCrLf & vbCrLf & _
            "以上" & vbCrLf & vbCrLf & CompNameQuery(.Fields("sm37")) & "</DIV>&nbsp;"
         strPdfName = .Fields("od03") & "_" & YM & ".pdf"
         frmPDF.Show
         frmPDF.StartProcess m_AttachPath, strPdfName
         Load frmacc14k0
         With frmacc14k0
            .m_bolCalled = True
            .MaskEdBox1 = CFDate(TransDate(YM & "01", 1))  '入帳日期起
            .MaskEdBox2 = CFDate(TransDate(CompDate(2, -1, YM2 & "01"), 1)) '入帳日期迄
            .Text4 = "" '身分
            .Text2 = adoRst.Fields("od03") '員工號起
            .Text3 = adoRst.Fields("od03") '員工號迄
            .Text1 = "N" '印証明單
            .Command1.Value = True
         End With
         Unload frmacc14k0
         frmPDF.EndtProcess
         Unload frmPDF
         'Modified by Morgan 2024/6/18 +CC給所內郵件收件者
         If PUB_SalarySendMail(stSubject, .Fields("sm01"), m_AttachPath & "\" & strPdfName, strMsg, Check4.Value, strContent, True, True) = True Then
            strMsgOK = strMsgOK & .Fields("sm01") & .Fields("st02") & IIf(strMsg <> "", "(" & strMsg & ")", "") & vbCrLf
         Else
            strMsgErr = strMsgErr & .Fields("sm01") & .Fields("st02") & ":" & strMsg & vbCrLf
         End If
         .MoveNext
      Loop
      End With
      MsgBoxU "EMail完成，清單如下：" & vbCrLf & "成功:" & vbCrLf & strMsgOK & IIf(strMsgErr <> "", vbCrLf & "失敗：" & strMsgErr, ""), vbExclamation
   Else
      MsgBox "無資料可EMail!"
   End If
End Sub
Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         Screen.MousePointer = vbHourglass
         If TxtValidate = True Then
            Me.Enabled = False
            YM = 100 * Val(Text1) + Val(Text2) + 191100
            
            'Added by Morgan 2024/5/23
            If Check5.Value = vbChecked Then
               Mail2Translator
            Else
            'end 2024/5/23
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
   
   'Added by Morgan 2023/10/5
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   Else
      PUB_KillAttach m_AttachPath
   End If
   'end 2023/10/5
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_KillAttach m_AttachPath 'Added by Morgan 2023/10/5
   Set frm170208 = Nothing
End Sub


Private Sub FormReset()
   Dim stDate As String
   If Val(Right(strSrvDate(2), 2)) < 11 Then
      stDate = CompDate("1", -1, strSrvDate(1)) - 19110000
   Else
      stDate = strSrvDate(2)
   End If
   Text1.Text = stDate \ 10000
   Text2.Text = Val(Right(stDate \ 100, 2))
   Text3.Text = ""
   lblName.Caption = ""
End Sub

Private Sub Form_Activate()
   If m_Actived = False Then
      FormReset
      Text2.SetFocus
      Text2_GotFocus
      m_Actived = True
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 <> "" Then
      If Val(Text2) < 1 Or Val(Text2) > 12 Then
         MsgBox "月份輸入錯誤 !"
         Text2_GotFocus
         Cancel = True
      End If
   End If
End Sub

Private Sub Text3_Change()
   lblName = ""
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Added by Morgan 2015/9/24
'畫表格
Private Sub PrintTable(pCompNo As String)
   Dim strTmp As String
   Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long
   
   'Printer.Font = "標楷體"
   'Modified by Morgan 2023/10/5 改印公司名稱--婉莘
   'strTmp = "台一關係企業"
   strTmp = CompNameQuery(pCompNo)
   'end 2023/10/5
   Printer.FontSize = 16
   Printer.FontBold = True
   'Modified by Morgan 2023/10/5
   'xi = Xo + 5100 - Printer.TextWidth(strTmp)
   xi = Xo + 5600 - Printer.TextWidth(strTmp)
   'end 2023/5/10
   yi = Yo + 570
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print strTmp
   
   Printer.FontSize = 12
   'Modified by Morgan 2023/10/5
   'xi = Xo + 5200
   xi = Xo + 5700
   'end 2023/10/5
   yi = Yo + 620
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "　　年　　月份　員工薪資明細表"
   '表格1
   yi = Yo + 1100
   '橫
   x1 = Xo + 850: x2 = x1 + 17 * 567
   y1 = yi - 150: y2 = y1
   Printer.Line (x1, y1)-(x2, y2)
   
   xi = Xo + 850 + 100
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "員工編號"
   xi = xi + 3.8 * 567
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "姓名"
   xi = xi + 3.85 * 567
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "部門"
   xi = xi + 5.75 * 567
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "支給日數"
   
   y1 = y1 + 0.9 * 567: y2 = y1
   Printer.Line (x1, y1)-(x2, y2)
   '直
   y1 = Yo + 950: y2 = y1 + 0.9 * 567
   '1
   x1 = Xo + 850: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '2
   x1 = x1 + 2.1 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '3
   x1 = x1 + 1.7 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '4
   x1 = x1 + 1.35 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '5
   x1 = x1 + 2.5 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '6
   x1 = x1 + 1.35 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '7
   x1 = x1 + 4.4 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '8
   x1 = x1 + 2.1 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '9
   x1 = Xo + 850 + 17 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   
   '表格2
   yi = Yo + 2320
   '橫
   '1
   x1 = Xo + 850: x2 = x1 + 17 * 567
   y1 = yi - 580: y2 = y1
   Printer.Line (x1, y1)-(x2, y2)
   
   Printer.FontSize = 11
   yi = y1 + 110
   xi = Xo + 850 + 1 * 567 + 210
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "基本薪資"
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "伙食津貼"
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "差旅津貼"
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "加班時數"
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "加 班 費"
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "職務津貼"
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi - 40: Printer.CurrentY = yi
   Printer.FontSize = 12
   Printer.Print "應發總額"
   
   yi = yi + 0.7 * 567
   xi = xi + 300
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "(A)"
   
   '2
   x1 = Xo + 850 + 1 * 567: x2 = x1 + 16 * 567
   y1 = y1 + 0.8 * 567: y2 = y1
   Printer.Line (x1, y1)-(x2, y2)
   '3
   x2 = x1 + 13.8 * 567
   y1 = y1 + 0.9 * 567: y2 = y1
   Printer.Line (x1, y1)-(x2, y2)
   
   'Printer.FontSize = 9
   Printer.FontSize = 11
   yi = y1 + 110
   'modify by sonia 2020/5/6 技術津貼改技術/證照津貼,並-80字縮小
   xi = Xo + 850 + 1 * 567 + 210 - 100
   Printer.CurrentX = xi: Printer.CurrentY = yi
   'Modify By Sindy 2020/6/22
   'Printer.Print "技術/證照津貼"
   Printer.Print "證照津貼"
   'Printer.FontSize = 11
   '2020/6/22 END
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "房租津貼"
   'Modify By Sindy 2020/6/22
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "技術津貼"
   '2020/6/22 END
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "特 支 費"
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "其他所得"
   
   '4
   x2 = x1 + 13.8 * 567
   y1 = y1 + 0.8 * 567: y2 = y1
   Printer.Line (x1, y1)-(x2, y2)
   '5
   x1 = Xo + 850: x2 = x1 + 17 * 567
   y1 = y1 + 0.9 * 567: y2 = y1
   Printer.Line (x1, y1)-(x2, y2)
   
   '直
   yi = Yo + 2320
   '1
   y1 = yi - 580: y2 = y1 + 3.4 * 567
   x1 = Xo + 850: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   
   Printer.FontSize = 12
   xi = x1 + 150
   yi = y1 + 300
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "應"
   yi = yi + 380
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "發"
   yi = yi + 380
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "薪"
   yi = yi + 380
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "資"
   
   '2
   x1 = x1 + 1 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '3
   x1 = x1 + 2.3 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '4
   x1 = x1 + 2.3 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '5
   x1 = x1 + 2.3 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '6
   x1 = x1 + 2.3 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '7
   x1 = x1 + 2.3 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '8
   x1 = Xo + 850 + 14.8 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '9
   x1 = Xo + 850 + 17 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
      
   Printer.FontSize = 11
   xi = Xo + 10730
   yi = y1 + 100
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "退"
   yi = yi + 250
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "休"
   yi = yi + 250
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "金"
   yi = yi + 250
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "適"
   yi = yi + 250
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "用"
   yi = yi + 250
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "："
   
   '表格3
   yi = Yo + 4470
   '橫
   '1
   x1 = Xo + 850: x2 = x1 + 17 * 567
   y1 = yi - 580: y2 = y1
   Printer.Line (x1, y1)-(x2, y2)

   Printer.FontSize = 11
   yi = y1 + 110
   xi = Xo + 850 + 1 * 567 + 210
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "勞 保 費"
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "健 保 費"
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi - 30: Printer.CurrentY = yi
   Printer.FontSize = 10
   Printer.Print "退休金自提"
   Printer.FontSize = 11
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "所 得 稅"
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi - 30: Printer.CurrentY = yi
   Printer.FontSize = 10
   Printer.Print "互助會會款"
   Printer.FontSize = 11
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "員工貸款"
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi - 40: Printer.CurrentY = yi
   Printer.FontSize = 12
   Printer.Print "應扣總額"
   
   yi = yi + 0.7 * 567
   xi = xi + 300
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "(B)"
   
   
   '2
   x1 = Xo + 850 + 1 * 567: x2 = x1 + 16 * 567
   y1 = y1 + 0.8 * 567: y2 = y1
   Printer.Line (x1, y1)-(x2, y2)
   '3
   x2 = x1 + 13.8 * 567
   y1 = y1 + 0.9 * 567: y2 = y1
   Printer.Line (x1, y1)-(x2, y2)
   
   Printer.FontSize = 11
   yi = y1 + 110
   xi = Xo + 850 + 1 * 567 + 210
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "借　　支"
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi - 30: Printer.CurrentY = yi
   Printer.FontSize = 10
   Printer.Print "未打卡扣款"
   Printer.FontSize = 11
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "缺勤扣薪"
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "婚喪喜慶"
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "其他扣款"
   xi = xi + 2.3 * 567
   Printer.CurrentX = xi - 30: Printer.CurrentY = yi
   Printer.FontSize = 10
   Printer.Print "補充健保費"
   
   '4
   x2 = x1 + 13.8 * 567
   y1 = y1 + 0.8 * 567: y2 = y1
   Printer.Line (x1, y1)-(x2, y2)
   '5
   x1 = Xo + 850: x2 = x1 + 17 * 567
   y1 = y1 + 0.9 * 567: y2 = y1
   Printer.Line (x1, y1)-(x2, y2)
   
   xi = Xo + 960
   yi = y1 + 110
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.FontSize = 12
   Printer.Print "實發金額(C)＝(A)－(B)"
   
   '6
   x1 = Xo + 850: x2 = x1 + 17 * 567
   y1 = y1 + 0.8 * 567: y2 = y1
   Printer.Line (x1, y1)-(x2, y2)
   
   '直
   yi = Yo + 4470
   '1
   y1 = yi - 580: y2 = y1 + 4.2 * 567
   x1 = Xo + 850: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   
   Printer.FontSize = 12
   xi = x1 + 150
   yi = y1 + 300
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "應"
   yi = yi + 380
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "扣"
   yi = yi + 380
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "項"
   yi = yi + 380
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "目"
   
   '2
   y2 = y1 + 3.4 * 567
   x1 = x1 + 1 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '3
   x1 = x1 + 2.3 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '4
   x1 = x1 + 2.3 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '5
   x1 = x1 + 2.3 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '6
   x1 = x1 + 2.3 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '7
   x1 = x1 + 2.3 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '8
   y2 = y1 + 4.2 * 567
   x1 = Xo + 850 + 14.8 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   '9
   x1 = Xo + 850 + 17 * 567: x2 = x1
   Printer.Line (x1, y1)-(x2, y2)
   
   Printer.FontSize = 11
   xi = Xo + 850
   yi = Yo + 6350
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "☆本月事務所提繳退休金金額："
   
   xi = Xo + 850
   yi = Yo + 6590
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "☆互助會明細："
   
   xi = Xo + 850
   yi = Yo + 6830
   Printer.CurrentX = xi: Printer.CurrentY = yi
   Printer.Print "☆婚喪喜慶當事人："
   Printer.FontBold = False
   
End Sub

Private Sub PrintSheet()
   Dim stCon As String
   Dim strDesc1 As String, strDesc2 As String
   Dim strPdfName As String, strMsg As String, strMsgOK As String, strMsgErr As String 'Added by Morgan 2023/10/4
   Dim stSubject As String 'Added by Morgan 2024/2/2
   
   Xo = 0 + Val(txtSet(1)) * 567
   Yo = -240 + Val(txtSet(2)) * 567
   dblUnitWidth = 1300 '欄位寬
   dblUnitHeight = 480 '欄位高
   
   stCon = ""
   If Text3 <> "" Then
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      'stCon = stCon & " and replace(sm01,'A','0')>='" & Text3 & "'"
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      stCon = stCon & " and substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)='" & Text3 & "'"
   End If
   
   'Added by Morgan 2015/10/7
   If Text5 <> "" Then
      stCon = stCon & " and st06='" & Text5 & "'"
   End If
   'end 2015/10/7
   
   'Added by Morgan 2015/12/10
   If Check1.Value = 1 Then
      stCon = stCon & " and sd50='Y'"
   End If
   'end 2015/12/10
   
   'Added by Morgan 2023/10/4
   If Check3.Value = vbChecked Then
      '離職日抓隔月5號以前的
      stCon = stCon & " and (sm03='R04' or (st04='2' and st51<=" & CompDate(1, 1, YM & "05") & " and sd05 in ('2','4','5','6')))"
   End If
   'end 2023/10/4
   
   '排序:所別,部門,員工編號
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modify By Sindy 2020/6/23 + sm45
   'Modified by Morgan 2023/10/5 +sm37
   'Modified by Morgan 2023/12/26 +新部門
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   strExc(0) = "select st06,sm03,st01,max(sd16) sd16,max(st02) st01N,max(decode(sign(sm02-" & Left(新部門啟用日, 6) & "),-1,a0902,a0922)) sm03N,max(sm27) sm27" & _
      ",sum(sm04) sm04,sum(sm05) sm05,sum(sm06) sm06,sum(sm07) sm07,sum(sm08) sm08" & _
      ",sum(sm09) sm09,sum(sm10) sm10,sum(sm11) sm11,sum(sm12) sm12,sum(sm13) sm13" & _
      ",sum(sm14) sm14,sum(sm15) sm15,sum(sm16) sm16,sum(sm17) sm17,sum(sm18) sm18,sum(sm19) sm19" & _
      ",sum(sm20) sm20,sum(sm21) sm21,sum(sm22) sm22,sum(sm23) sm23,sum(sm24) sm24,sum(sm45) sm45" & _
      ",sum(nvl(sm04,0)+nvl(sm05,0)+nvl(sm06,0)+nvl(sm07,0)+nvl(sm08,0)" & _
      "+nvl(sm09,0)+nvl(sm10,0)+nvl(sm12,0)+nvl(sm13,0)+nvl(sm45,0)) s1" & _
      ",sum(nvl(sm14,0)+nvl(sm15,0)+nvl(sm16,0)+nvl(sm17,0)+nvl(sm18,0)+nvl(sm19,0)" & _
      "+nvl(sm20,0)+nvl(sm21,0)+nvl(sm22,0)+nvl(sm23,0)+nvl(sm24,0)+nvl(sm43,0)) s2" & _
      ",sum(sm30) sm30,sum(sm43) sm43,sm37" & _
      " from salarymonth,staff,acc090,acc090new,salarydata" & _
      " where st01(+)=substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) and a0901(+)=sm03 and a0921(+)=sm03 and sd01(+)=sm01" & _
      " and sm02=" & YM & " and sm01<'F'" & stCon & " group by st06,sm03,st01,sm37"
   'end 2023/12/26
   'Removed by Morgan 2015/12/10
   '將改個人可查詢，財務處不列印，取消有加班費條件
   'If Check1.Value = 1 Then
   '   strExc(0) = strExc(0) & " having sum(sm11)>0"
   'End If
   'end 2015/12/10
   
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Printer.EndDoc
      If Check3.Value = vbUnchecked And Check4.Value = vbUnchecked Then 'Added by Morgan 2023/10/5
         'Added by Morgan 2015/9/24
         If Check2.Value = 1 Then
            Printer.PaperSize = 9
            PrintTable adoRst("sm37")
         Else
         'end 2015/9/24
            Printer.PaperSize = PUB_GetPaperSize(3)
         End If
         Printer.Font = "新細明體"
         Printer.FontSize = 12
      End If 'Added by Morgan 2023/10/5
      
      With adoRst
      Do While Not .EOF
      
         'Added by Morgan 2023/10/4
         If Check3.Value = vbChecked Or Check4.Value = vbChecked Then
            stSubject = Val(Text1) & "年" & Val(Text2) & "月份薪資單(" & .Fields("st01") & ")"
            strPdfName = .Fields("st01") & "_" & YM & ".pdf"
            frmPDF.Show
            frmPDF.StartProcess m_AttachPath, strPdfName
            
            If Check2.Value = 1 Then
               Printer.PaperSize = 9
               PrintTable adoRst("sm37")
            Else
               Printer.PaperSize = PUB_GetPaperSize(3)
            End If
            Printer.Font = "新細明體"
            Printer.FontSize = 12
         Else
         'end 2023/10/4
         
            If .AbsolutePosition > 1 Then
               Printer.NewPage
               If Check2.Value = 1 Then PrintTable adoRst("sm37") 'Added by Morgan 2015/10/7
            End If
            
         End If 'Added by Morgan 2023/10/4
         
         '年
         'Added by Morgan 2023/10/5
         If Check2.Value = 1 Then
            xi = Xo + 5700
         Else
         'end 2023/10/5
            xi = Xo + 5200
         End If
         yi = Yo + 620
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print Val(Text1)
         '月
         xi = xi + 850
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print Val(Text2)
         
         '員工編號
         strExc(1) = "" & .Fields("st01")
         xi = Xo + 2555 - Printer.TextWidth(strExc(1)) / 2
         yi = Yo + 1100
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '姓名
         strExc(1) = "" & .Fields("st01N")
         xi = Xo + 4450 - Printer.TextWidth(strExc(1)) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         'Modified by Morgan 2023/10/4
         'Printer.Print strExc(1)
         PUB_PrintUnicodeText strExc(1), xi, yi
         'end 2023/10/4
         
         '部門
         strExc(1) = "" & .Fields("sm03N")
         xi = Xo + 7115 - Printer.TextWidth(strExc(1)) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '支給日數
         strExc(1) = "" & .Fields("sm27")
         xi = Xo + 9945
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '基本薪資
         strExc(1) = Format("" & .Fields("sm04"), "#,###")
         xi = Xo + 2500 - Printer.TextWidth(strExc(1))
         yi = Yo + 2320
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '伙食津貼
         strExc(1) = Format("" & .Fields("sm07"), "#,###")
         xi = Xo + 2500 + dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '交通津貼
         strExc(1) = Format("" & .Fields("sm08"), "#,###")
         xi = Xo + 2500 + 2 * dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '加班時數
         strExc(1) = "" & .Fields("sm11")
         xi = Xo + 2500 + 3 * dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '加班費
         strExc(1) = Format("" & .Fields("sm12"), "#,###")
         xi = Xo + 2500 + 4 * dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '職務津貼
         strExc(1) = Format("" & .Fields("sm05"), "#,###")
         xi = Xo + 2500 + 5 * dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '應發總額
         strExc(1) = Format("" & .Fields("s1"), "#,###")
         xi = Xo + 2450 + 6 * dblUnitWidth - Printer.TextWidth(strExc(1))
         yi = Yo + 2300 + dblUnitHeight
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         'Add By Sindy 2020/6/23
         '證照津貼
         'strExc(1) = Format("" & .Fields("sm06"), "#,###")
         strExc(1) = Format("" & .Fields("sm45"), "#,###")
         '2020/6/23 END
         xi = Xo + 2500 - Printer.TextWidth(strExc(1))
         yi = Yo + 2320 + 2 * dblUnitHeight
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print ""; strExc(1)
         
         '房租津貼
         strExc(1) = Format("" & .Fields("sm09"), "#,###")
         xi = Xo + 2500 + dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '技術津貼
         strExc(1) = Format("" & .Fields("sm06"), "#,###")
         'xi = Xo + 2500 - Printer.TextWidth(strExc(1))
         xi = Xo + 2500 + 2 * dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print ""; strExc(1)
         
         '特支費
         strExc(1) = Format("" & .Fields("sm10"), "#,###")
         xi = Xo + 2500 + 3 * dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '其他所得
         strExc(1) = Format("" & .Fields("sm13"), "#,###")
         xi = Xo + 2500 + 4 * dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '勞保費
         strExc(1) = Format("" & .Fields("sm14"), "#,###")
         xi = Xo + 2500 - Printer.TextWidth(strExc(1))
         yi = Yo + 4470
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '健保費
         strExc(1) = Format("" & .Fields("sm15"), "#,###")
         xi = Xo + 2500 + dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '退休金自提
         strExc(1) = Format("" & .Fields("sm16"), "#,###")
         xi = Xo + 2500 + 2 * dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '所得稅
         strExc(1) = Format("" & .Fields("sm24"), "#,###")
         xi = Xo + 2500 + 3 * dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '互助會款
         strExc(1) = Format("" & .Fields("sm18"), "#,###")
         xi = Xo + 2500 + 4 * dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '員工貸款
         strExc(1) = Format("" & .Fields("sm19"), "#,###")
         xi = Xo + 2500 + 5 * dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '應扣總額
         strExc(1) = Format("" & .Fields("s2"), "#,###")
         xi = Xo + 2450 + 6 * dblUnitWidth - Printer.TextWidth(strExc(1))
         yi = Yo + 4450 + dblUnitHeight
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '借支
         strExc(1) = Format("" & .Fields("sm20"), "#,###")
         xi = Xo + 2500 - Printer.TextWidth(strExc(1))
         yi = Yo + 4470 + 2 * dblUnitHeight
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '未打卡扣款
         strExc(1) = Format("" & .Fields("sm22"), "#,###")
         xi = Xo + 2500 + dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '缺勤扣薪
         strExc(1) = Format("" & .Fields("sm21"), "#,###")
         xi = Xo + 2500 + 2 * dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '婚喪喜慶
         strExc(1) = Format("" & .Fields("sm17"), "#,###")
         xi = Xo + 2500 + 3 * dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '其他扣款
         strExc(1) = Format("" & .Fields("sm23"), "#,###")
         xi = Xo + 2500 + 4 * dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         'Added by Morgan 2013/1/30
         '補充保費
         strExc(1) = Format("" & .Fields("sm43"), "#,###")
         xi = Xo + 2500 + 5 * dblUnitWidth - Printer.TextWidth(strExc(1))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         'end 2013/1/30
         
         '實發金額
         strExc(1) = Format(Val("" & .Fields("s1")) - Val("" & .Fields("s2")), "#,###")
         xi = Xo + 2450 + 6 * dblUnitWidth - Printer.TextWidth(strExc(1))
         yi = Yo + 4470 + 3 * dblUnitHeight
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         '勞退公司提撥
         If "" & .Fields("sd16") = "Y" Then
            xi = Xo + 10730
            yi = Yo + 3280
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print "新"
            
            strExc(1) = Format("" & .Fields("sm30"), "#,###")
            xi = Xo + 4200
            yi = Yo + 6350
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
         Else
            xi = Xo + 10730
            yi = Yo + 3280
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print "舊"
         End If
         
         xi = Xo + 10730
         yi = Yo + 3280 + Printer.TextHeight("制")
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print "制"
         
         strDesc1 = ""
         strDesc2 = ""
         '婚喪戶助有金額
         If Val("" & .Fields("sm17")) + Val("" & .Fields("sm18")) > 0 Then
            strExc(0) = "select wfa05,wfa02,st02,sum(wfa04) x1" & _
               " from WFAmount,staff where substr(wfa01,1,6)=" & YM & " and wfa03='" & .Fields("st01") & "' and st01(+)=wfa02 group by wfa05,wfa02,st02"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               With RsTemp
               Do While Not .EOF
                  '婚喪當事人
                  If "" & .Fields("wfa05") = "1" Then
                     strDesc1 = strDesc1 & "  " & .Fields("st02")
                  Else
                     strDesc2 = strDesc2 & "  (" & .Fields("wfa02") & "  $" & Format(Val("" & .Fields("x1")), "#,###") & ")"
                  End If
                  .MoveNext
               Loop
               End With
            End If
         End If
         
         '互助會明細
         If strDesc2 <> "" Then
            xi = Xo + 2400
            yi = Yo + 6590
            Printer.CurrentX = xi: Printer.CurrentY = yi
            'Modified by Morgan 2023/10/4
            'Printer.Print strDesc2
            PUB_PrintUnicodeText strDesc2, xi, yi
            'end 2023/10/4
         End If
         
         '婚喪當事人
         If strDesc1 <> "" Then
            xi = Xo + 2900
            yi = Yo + 6830
            Printer.CurrentX = xi: Printer.CurrentY = yi
            'Modified by Morgan 2023/10/4
            'Printer.Print strDesc1
            PUB_PrintUnicodeText strDesc1, xi, yi
            'end 2023/10/4
         End If
         
         'Added by Morgan 2023/10/4
         If Check3.Value = vbChecked Or Check4.Value = vbChecked Then
            Printer.EndDoc
            frmPDF.EndtProcess
            Unload frmPDF
            If PUB_SalarySendMail(stSubject, .Fields("st01"), m_AttachPath & "\" & strPdfName, strMsg, Check4.Value) = True Then
               strMsgOK = strMsgOK & .Fields("st01") & .Fields("st01N") & IIf(strMsg <> "", "(" & strMsg & ")", "") & vbCrLf
            Else
               strMsgErr = strMsgErr & .Fields("st01") & .Fields("st01N") & ":" & strMsg & vbCrLf
            End If
         End If
         'end 2023/10/4
         .MoveNext
      Loop
      End With
      If Check3.Value = vbChecked Or Check4.Value = vbChecked Then
         MsgBoxU "EMail完成，清單如下：" & vbCrLf & "成功:" & vbCrLf & strMsgOK & IIf(strMsgErr <> "", vbCrLf & "失敗：" & strMsgErr, ""), vbExclamation
      Else
         Printer.EndDoc
         MsgBox "列印結束 !"
      End If
      
   Else
      MsgBox "無資料可列印 !"
   End If
   Set adoRst = Nothing
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   
   If Text1 = "" Then
      MsgBox "年度不可空白 !"
      Text1.SetFocus
      Exit Function
   End If
   If Text2 = "" Then
      MsgBox "月份不可空白 !"
      Text2.SetFocus
      Exit Function
   End If
   
   Text2_Validate bCancel
   If bCancel = True Then
      Text2.SetFocus
      Text2_GotFocus
      Exit Function
   End If
   
   'Added by Morgan 2015/12/10
   'Modified by Morgan 2024/5/23 +And Check5.Value = 0
   If Check2.Value = 0 And Check5.Value = 0 Then
      'Added by Morgan 2023/10/5
      If Check3.Value = vbChecked Or Check4.Value = vbChecked Then
         MsgBox "勾選EMail必須為【A4格式】！", vbExclamation
         Exit Function
      End If
      'end 2023/10/5
      If MsgBox("未勾選【A4格式】，是否確定要繼續？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Exit Function
      End If
   End If
   
   'Modified by Morgan 2023/10/65 +Check3.Value = 0
   'Modified by Morgan 2024/5/23 +Check5
   If Text3.Text = "" And (Check1.Value = 0 And Check3.Value = 0 And Check5.Value = 0) Then
      If MsgBox("未勾選任一下列選項，是否確定要繼續？" & vbCrLf & "【" & Check1.Caption & " 】" & vbCrLf & "【" & Check3.Caption & "】" & vbCrLf & "【" & Check5.Caption & "】", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
         Exit Function
      End If
   End If
   'end 2015/12/10
   
   TxtValidate = True
End Function

Private Sub Text3_Validate(Cancel As Boolean)
   If Text3 <> "" Then
      lblName = GetStaffName(Text3, True)
      If lblName = "" Then
         MsgBox "員工編號輸入錯誤！", vbCritical
         Cancel = True
      Else
         Check1.Value = 0
      End If
   End If
End Sub

Private Sub txtSet_GotFocus(Index As Integer)
   TextInverse txtSet(Index)
End Sub

