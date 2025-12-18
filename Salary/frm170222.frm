VERSION 5.00
Begin VB.Form frm170222 
   BorderStyle     =   1  '單線固定
   Caption         =   "勞、健、補充保費扣繳證明書"
   ClientHeight    =   2760
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3432
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   3432
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   4
      Left            =   1710
      MaxLength       =   1
      TabIndex        =   2
      Top             =   960
      Width           =   345
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   1125
      MaxLength       =   1
      TabIndex        =   1
      Top             =   960
      Width           =   345
   End
   Begin VB.TextBox txtSet 
      Height          =   285
      Index           =   2
      Left            =   1485
      TabIndex        =   7
      Top             =   2340
      Width           =   705
   End
   Begin VB.TextBox txtSet 
      Height          =   285
      Index           =   1
      Left            =   1485
      TabIndex        =   6
      Top             =   2040
      Width           =   705
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   810
      Style           =   2  '單純下拉式
      TabIndex        =   5
      Top             =   1650
      Width           =   2460
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Left            =   1245
      TabIndex        =   8
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdeXit 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   2325
      TabIndex        =   9
      Top             =   90
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Index           =   0
      Left            =   1110
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "96"
      Top             =   630
      Width           =   435
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Index           =   1
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1290
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Index           =   2
      Left            =   2115
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1290
      Width           =   735
   End
   Begin VB.Line Line2 
      X1              =   1845
      X2              =   2115
      Y1              =   1410
      Y2              =   1410
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   1710
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "公司別："
      Height          =   180
      Left            =   330
      TabIndex        =   15
      Top             =   990
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "縱軸偏移值(Y)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   5
      Left            =   45
      TabIndex        =   14
      Top             =   2400
      Width           =   3240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "橫軸偏移值(X)：　　　　　　(單位公分)"
      Height          =   180
      Index           =   4
      Left            =   45
      TabIndex        =   13
      Top             =   2100
      Width           =   3240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      Height          =   255
      Index           =   3
      Left            =   75
      TabIndex        =   12
      Top             =   1740
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "扣繳年度："
      Height          =   180
      Left            =   150
      TabIndex        =   11
      Top             =   660
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   10
      Top             =   1320
      Width           =   900
   End
End
Attribute VB_Name = "frm170222"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Memo by Morgan 2024/1/31 新部門已修改
'Create by Morgan 2009/1/17
Option Explicit

Dim m_Actived As Boolean
Dim m_DefaultPrinter As String

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
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
End Sub

Private Function TxtValidate() As Boolean
   If Text1(0) = "" Then
      MsgBox "年度不可空白 !"
      Text1(0).SetFocus
      Exit Function
   End If
   TxtValidate = True
End Function

Private Sub PrintSheet()
   Dim stConSm As String, stConNhi As String
   Dim strFontSize As String, strFontName As String
   Dim Xo As Integer, Yo As Integer, xi As Long, yi As Long
   Dim iWidth As Integer, iLineHeight As Integer
   
   If Text1(0) <> "" Then
      stConSm = stConSm & " and substr(sm02,1,4)=" & (Val(Text1(0)) + 1911)
      stConNhi = stConNhi & " and substr(nhi02,1,4)=" & (Val(Text1(0)) + 1911)
   End If

   If Text1(1) <> "" Then
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      stConSm = stConSm & " and substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)>='" & Text1(1) & "'"
   End If

   If Text1(2) <> "" Then
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      stConSm = stConSm & " and substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)<='" & Text1(2) & "'"
   End If
   
   'Add by Morgan 2010/1/18 加可輸公司別
   If Text1(3) <> "" Then
      stConSm = stConSm & " and sm37>='" & Text1(3) & "'"
      stConNhi = stConNhi & " and nhi11>='" & Text1(3) & "'"
   End If
   
   If Text1(4) <> "" Then
      stConSm = stConSm & " and sm37<='" & Text1(4) & "'"
      stConNhi = stConNhi & " and nhi11<='" & Text1(3) & "'"
   End If
   'end 2010/1/18
   
   'Modify by Morgan 2010/1/18 第二家公司也有勞保費
   'Modify by Morgan 2010/1/29 排序改跟扣單一樣--婧瑄
   'strExc(0) = "select s1.st02 name,s1.st26 id,s1.st23 birth,a0802 comp,s2.st02 inCharge,x.*" & _
      " from (SELECT sm01,sm37" & _
      ",MIN(decode(sign(sm14),1,SM02)) dfL,MAX(decode(sign(sm14),1,SM02)) dtL,SUM(SM14) sL" & _
      ",MIN(decode(sign(sm15),1,SM02)) dfH,MAX(decode(sign(sm15),1,SM02)) dtH,SUM(SM15) sH" & _
      " FROM SALARYMONTH WHERE NVL(SM14,0)+NVL(SM15,0)>0" & stConSm & _
      " GROUP BY sm01,sm37) x,Staff s1,acc080,staff s2" & _
      " where s1.st01(+)=replace(sm01,'A','0') and a0801(+)=sm37 and s2.st01(+)=a0806" & _
      " order by a0802,s1.st06,s1.st03,s1.st01"
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modified by Morgan 2014/1/23 +補充保費
   'strExc(0) = "select s1.st04,s1.st06,s1.st03,s1.st26 id,s1.st02 name,s1.st23 birth,a0802 comp,s2.st02 inCharge,x.*" & _
      " from (SELECT sm01,sm37" & _
      ",MIN(decode(sign(sm14),1,SM02)) dfL,MAX(decode(sign(sm14),1,SM02)) dtL,SUM(SM14) sL" & _
      ",MIN(decode(sign(sm15),1,SM02)) dfH,MAX(decode(sign(sm15),1,SM02)) dtH,SUM(SM15) sH" & _
      " FROM SALARYMONTH WHERE NVL(SM14,0)+NVL(SM15,0)>0" & stConSm & _
      " GROUP BY sm01,sm37) x,Staff s1,acc080,staff s2" & _
      " where s1.st01(+)=substr(sm01,1,1)||replace(substr(sm01,2),'A','0') and a0801(+)=sm37 and s2.st01(+)=a0806" & _
      " order by 1,2,3,4"
   'Modified by Morgan 2014/2/26 73014,73017,73035 103/1轉換公司別,月薪資第3碼為A的員工號要轉為原來的編號
   'Modified by Morgan 2016/1/14 +勞健保還要抓其他所得扣款資料
   'strExc(0) = "select s1.st04,s1.st06,s1.st03,s1.st26 id,s1.st02 name,s1.st23 birth,a0802 comp,s2.st02 inCharge,x.*,y.*" & _
      " from (SELECT sm01,substr(sm01,1,1)||replace(substr(sm01,2),'A','0') sid,sm37" & _
      ",MIN(decode(sign(sm14),1,SM02)) dfL,MAX(decode(sign(sm14),1,SM02)) dtL,SUM(SM14) sL" & _
      ",MIN(decode(sign(sm15),1,SM02)) dfH,MAX(decode(sign(sm15),1,SM02)) dtH,SUM(SM15) sH" & _
      " FROM SALARYMONTH WHERE NVL(SM14,0)+NVL(SM15,0)>0" & stConSm & " GROUP BY sm01,sm37" & _
      ") x,Staff s1,(select min(st01) y0,nvl(st26,st01) y1,nhi11 y2,min(substr(nhi02,1,6)) dfH2,max(substr(nhi02,1,6)) dtH2,sum(nhi06) sH2" & _
      " from nhi2nd,staff where st01(+)=nhi01 and st01>'6'" & stConNhi & " group by nvl(st26,st01),nhi11) y,acc080,staff s2" & _
      " where y0(+)=sid and y2(+)=sm37 and s1.st01(+)=sid" & _
      " and a0801(+)=sm37 and s2.st01(+)=a0806 order by 1,2,3,4"
   'Modified by Morgan 2024/1/31 +新部門判斷
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   strExc(0) = "select s1.st04,s1.st06," & IIf(Val(Text1(0)) + 1911 >= Left(新部門啟用日, 4), "s1.st93", "s1.st03") & " st03,s1.st26 id,s1.st02 name,s1.st23 birth,a0802 comp,s2.st02 inCharge,x.*,y.*" & _
      " from (SELECT sm01,substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) sid,sm37" & _
      ",MIN(decode(sign(sm14),1,SM02)) dfL,MAX(decode(sign(sm14),1,SM02)) dtL,SUM(SM14) sL" & _
      ",MIN(decode(sign(sm15),1,SM02)) dfH,MAX(decode(sign(sm15),1,SM02)) dtH,SUM(SM15) sH" & _
      " from (select sm01,sm02,sm14,sm15,sm37 FROM SALARYMONTH WHERE NVL(SM14,0)+NVL(SM15,0)>0" & stConSm & _
      " union all select od03,od14,decode(od04,'31',od05,'32',-1*od05) sm14,decode(od04,'35',od05,'36',-1*od05) sm15,sm37" & _
      " from salarymonth,othersalarydata where od03(+)=sm01 and od14(+)=sm02 and od04 in ('31','32','35','36')" & stConSm & _
      ") GROUP BY sm01,sm37) x,Staff s1,(select min(st01) y0,nvl(st26,st01) y1,nhi11 y2,min(substr(nhi02,1,6)) dfH2,max(substr(nhi02,1,6)) dtH2,sum(nhi06) sH2" & _
      " from nhi2nd,staff where st01(+)=nhi01 and st01>'6'" & stConNhi & " group by nvl(st26,st01),nhi11) y,acc080,staff s2" & _
      " where y0(+)=sid and y2(+)=sm37 and s1.st01(+)=sid" & _
      " and a0801(+)=sm37 and s2.st01(+)=a0806 order by 1,2,3,4"
      
   'end 2014/2/26
   'end 2014/1/23
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strFontSize = Printer.FontSize
      strFontName = Printer.FontName
      Printer.EndDoc
      Printer.PaperSize = PUB_GetPaperSize(3)
      Printer.Font = "細明體"
      Printer.FontSize = 14
            
      Xo = Val(txtSet(1)) * 567
      Yo = Val(txtSet(2)) * 567
            
      iWidth = 21 * 567 - (Printer.Width - Printer.ScaleWidth)
      iLineHeight = Printer.TextHeight("字") + 30
      With RsTemp
      Do While Not .EOF
         If .AbsolutePosition > 1 Then Printer.NewPage
         
         Printer.FontSize = 18
         Printer.FontBold = True
         yi = Yo + 500 - (Printer.Height - Printer.ScaleHeight) / 2
         strExc(1) = "自付保險費證明書"
         xi = Xo + iWidth / 2 - Printer.TextWidth(strExc(1)) / 2
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         yi = yi + Printer.TextHeight(strExc(1))
         strExc(2) = String(GetTextLength(strExc(1)) + 2, "=")
         xi = xi - Printer.TextWidth("=")
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(2)
         
         Printer.FontSize = 14
         Printer.FontBold = False

         yi = yi + 2 * iLineHeight
         'Modify by Morgan 2010/1/18 台一國際要印公司不印事務所
         If InStr("" & .Fields("comp"), "事務所") > 0 Then
            strExc(1) = "查本事務所職員 " & .Fields("name") & " ，身份證字號為 ： " & .Fields("id") & "，"
         Else
            strExc(1) = "查本公司職員 " & .Fields("name") & " ，身份證字號為 ： " & .Fields("id") & "，"
         End If
         xi = Xo + 850
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         yi = yi + 1.5 * iLineHeight
         strExc(2) = "" & .Fields("birth")
         strExc(1) = "民國 " & (Val(Left(strExc(2), 4)) - 1911) & " 年  " & Val(Mid(strExc(2), 5, 2)) & _
            " 月 " & Val(Right(strExc(2), 2)) & " 日出生。 " & Val(Text1(0)) & " 年度自付勞保費、健保費及補充保費金額如下 ："
         xi = Xo + 850
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         yi = yi + 1.5 * iLineHeight
         strExc(1) = "　勞保費　 ： 自 " & Val(Text1(0)) & " 年 " & Right(" " & .Fields("dfL"), 2) & _
            " 月至 " & Val(Text1(0)) & " 年 " & Right(" " & .Fields("dtL"), 2) & " 月共 " & Right(String(8, " ") & Format(Val("" & .Fields("sL")), "$#,##0"), 8) & " 元整。"
         xi = Xo + 850
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         'Added by Morgan 2014/1/23
         '同健保費,不必帶扣繳月份起迄--辜
         yi = yi + iLineHeight
         strExc(1) = "　補充保費 ： 自 " & Val(Text1(0)) & " 年 " & Right(IIf(IsNull(.Fields("dfH")) = True, " " & .Fields("dfH2"), " " & .Fields("dfH")), 2) & _
            " 月至 " & Val(Text1(0)) & " 年 " & Right(IIf(IsNull(.Fields("dtH")) = True, " " & .Fields("dtH2"), " " & .Fields("dtH")), 2) & " 月共 " & Right(String(8, " ") & Format(Val("" & .Fields("sH2")), "$#,##0"), 8) & " 元整。"
         xi = Xo + 850
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         'end 2014/1/23

         yi = yi + iLineHeight
         strExc(1) = "　健保費　 ： 自 " & Val(Text1(0)) & " 年 " & Right(" " & .Fields("dfH"), 2) & _
            " 月至 " & Val(Text1(0)) & " 年 " & Right(" " & .Fields("dtH"), 2) & " 月共 " & Right(String(8, " ") & Format(Val("" & .Fields("sH")), "$#,##0"), 8) & " 元整。"
         xi = Xo + 850
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         yi = yi + iLineHeight
         'Modify by Morgan 2010/9/14
'         strExc(1) = "　健保費含眷口數    人，眷口明細如下 ："
'         xi = Xo + 850
'         Printer.CurrentX = xi: Printer.CurrentY = yi
'         Printer.Print strExc(1)
'
'         yi = yi + 1.5 * iLineHeight
'         strExc(1) = "　共 " & String(7, " ") & " 元整。"
'         xi = Xo + 850 + Printer.TextWidth(String(38, " "))
'         Printer.CurrentX = xi: Printer.CurrentY = yi
'         Printer.Print strExc(1)
'
'         yi = yi + iLineHeight
'         strExc(1) = "　共 " & String(7, " ") & " 元整。"
'         xi = Xo + 850 + Printer.TextWidth(String(38, " "))
'         Printer.CurrentX = xi: Printer.CurrentY = yi
'         Printer.Print strExc(1)
'
'         yi = yi + iLineHeight
'         strExc(1) = "　共 " & String(7, " ") & " 元整。"
'         xi = Xo + 850 + Printer.TextWidth(String(38, " "))
'         Printer.CurrentX = xi: Printer.CurrentY = yi
'         Printer.Print strExc(1)
'
'         yi = yi + iLineHeight
'         strExc(1) = "　共 " & String(7, " ") & " 元整。"
'         xi = Xo + 850 + Printer.TextWidth(String(38, " "))
'         Printer.CurrentX = xi: Printer.CurrentY = yi
'         Printer.Print strExc(1)
         strExc(0) = "select sr04,net from (select hm01,hm02,sum(hm04) net from salarymonth,himonth" & _
            " where sm01='" & .Fields("sm01") & "' and hm01(+)=sm01 and hm03(+)=sm02 and hm02>0" & stConSm & _
            " group by hm01,hm02),staff_relation where sr01(+)=hm01 and sr02(+)=hm02"
         intI = 1
         Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With AdoRecordSet3
            strExc(1) = "　健保費含眷口數 " & .RecordCount & " 人，眷口明細如下 ："
            xi = Xo + 850
            Printer.CurrentX = xi: Printer.CurrentY = yi
            Printer.Print strExc(1)
            intI = 0
            Do While Not .EOF
               intI = intI + 1
               strExc(1) = Left(.Fields(0) & "　　", 4) & " 共 " & Right(String(8, " ") & Format(Val("" & .Fields(1)), "$#,##0"), 8) & " 元整。"
               'Modified by Morgan 2014/1/23 明細改依序左右列印
               If intI Mod 2 = 1 Then
                  yi = yi + 1.5 * iLineHeight
                  xi = Xo + 850 + Printer.TextWidth(String(4, "　"))
               Else
                  xi = Xo + 850 + Printer.TextWidth(String(38, " "))
               End If
               Printer.CurrentX = xi: Printer.CurrentY = yi
               Printer.Print strExc(1)
               .MoveNext
            Loop
            End With
         End If
         'end 2010/9/14
         
'Remove by Morgan 2009/2/25 不印合計--婧瑄
'         yi = yi + 1.5 * iLineHeight
'         strExc(1) = "合計 " & Right(String(7, " ") & Format(Val("" & .Fields("sL")) + Val("" & .Fields("sH")), "$#,###"), 7) & " 元整。"
'         xi = Xo + 850 + Printer.TextWidth(String(38, " "))
'         Printer.CurrentX = xi: Printer.CurrentY = yi
'         Printer.Print strExc(1)
         
         yi = yi + 1.5 * iLineHeight
         strExc(1) = "謹此證明"
         xi = Xo + 850
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         yi = yi + 1.5 * iLineHeight
         strExc(1) = "" & .Fields("comp")
         xi = Xo + 850 + Printer.TextWidth(String(38, " "))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         yi = yi + 1.5 * iLineHeight
         strExc(1) = "負責人 ： " & .Fields("inCharge")
         xi = Xo + 850 + Printer.TextWidth(String(38, " "))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
         yi = yi + 1.5 * iLineHeight
         strExc(1) = "中華民國 " & (strSrvDate(2) \ 10000) & " 年 " & Mid(strSrvDate(1), 5, 2) & " 月 " & Mid(strSrvDate(1), 7) & " 日"
         xi = Xo + 850 + Printer.TextWidth(String(38, " "))
         Printer.CurrentX = xi: Printer.CurrentY = yi
         Printer.Print strExc(1)
         
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
End Sub

Private Sub FormReset()
   Text1(0) = strSrvDate(2) \ 10000
   Text1(1) = ""
   Text1(2) = ""
End Sub

Private Sub Form_Activate()
   If m_Actived = False Then
      FormReset
      Text1(0).SetFocus
      Text1_GotFocus 0
      m_Actived = True
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cmbPrinter, m_DefaultPrinter, , , Me.txtSet(1), Me.txtSet(2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170222 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   If Index = 2 Then
      If Text1(1) <> "" Then Text1(2) = Text1(1)
   End If
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
   End Select
End Sub
