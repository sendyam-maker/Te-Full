VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060405 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專程序季獎金報表"
   ClientHeight    =   2172
   ClientLeft      =   48
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2172
   ScaleWidth      =   4680
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   2
      Left            =   1335
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "1"
      Top             =   1320
      Width           =   285
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   3
      Left            =   1335
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "1"
      Top             =   1680
      Width           =   285
   End
   Begin VB.TextBox txtNo 
      Height          =   285
      Left            =   1335
      MaxLength       =   6
      TabIndex        =   2
      Top             =   930
      Width           =   750
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   0
      Left            =   1335
      MaxLength       =   5
      TabIndex        =   0
      Top             =   570
      Width           =   615
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   1
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   1
      Top             =   570
      Width           =   615
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3675
      TabIndex        =   6
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2880
      TabIndex        =   5
      Top             =   90
      Width           =   756
   End
   Begin MSForms.Label lbl1 
      Height          =   225
      Left            =   2130
      TabIndex        =   13
      Top             =   960
      Width           =   1395
      VariousPropertyBits=   27
      Size            =   "2461;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      X1              =   2010
      X2              =   2270
      Y1              =   705
      Y2              =   705
   End
   Begin VB.Label Label4 
      Caption         =   "(1.統計 2.明細)"
      Height          =   180
      Left            =   1665
      TabIndex        =   12
      Top             =   1365
      Width           =   1290
   End
   Begin VB.Label Label2 
      Caption         =   "報表內容："
      Height          =   180
      Index           =   1
      Left            =   315
      TabIndex        =   11
      Top             =   1350
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "輸出方式："
      Height          =   180
      Index           =   2
      Left            =   315
      TabIndex        =   10
      Top             =   1710
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "(1.Word 2.印表機)"
      Height          =   180
      Left            =   1665
      TabIndex        =   9
      Top             =   1725
      Width           =   1515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   4
      Left            =   315
      TabIndex        =   8
      Top             =   975
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "年月："
      Height          =   180
      Index           =   0
      Left            =   315
      TabIndex        =   7
      Top             =   615
      Width           =   570
   End
End
Attribute VB_Name = "frm060405"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Morgan 2018/6/14
Option Explicit

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
   Case 0
      If TxtValidate() = False Then Exit Sub
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      Process
      Me.Enabled = True
      Screen.MousePointer = vbDefault
   Case 1
      Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   '預設系統日之前一季
   'Modified by Morgan 2019/8/2
   ''預設系統日之前一月
   'strExc(1) = Format(DateAdd("m", -3, ChangeWStringToWDateString(Left(strSrvDate(1), 6) & "01")), "YYYYMM")
   'Txt1(0) = Left(strExc(1), 4) - 1911
   'Txt1(1) = (Right(strExc(1), 2) + 2) \ 3
   txt1(0) = Format(DateAdd("m", -1, ChangeWStringToWDateString(Left(strSrvDate(1), 6) & "01")), "YYYYMM") - 191100
   txt1(1) = txt1(0)
   txt1(0).Tag = txt1(0)
   txt1(1).Tag = txt1(1)
   'end 2019/8/2
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060405 = Nothing
End Sub

Private Sub Process()
   Dim strDate1 As String, strDate2 As String, stCon As String
   Dim strVTB As String
      
   'Modified by Morgan 2019/8/2 統計區間改年月起迄(原為年季)
   'strDate1 = (Txt1(0) + 1911) & Format((Txt1(1) * 3 - 2), "00") & "01"
   'strDate2 = (Txt1(0) + 1911) & Format((Txt1(1) * 3), "00") & "31"
   strDate1 = (txt1(0) + 191100) & "01"
   strDate2 = (txt1(1) + 191100) & "31"
   'end 2019/8/2
   
   If txtNo <> "" Then stCon = " and cp83='" & txtNo & "'"
   'Modified by Morgan 2019/8/2 +C類(1003通知補文件要排除承辦人為程序人員者)--淑華
   'FCP  (910其他->FCP案有經發文室的才算--敏莉)
   '+605年費->FCP案有經發文室的才算--淑華
   'Modified by Morgan 2020/6/20 605,910 電子送件有智慧局收文號的也要算--敏莉
   strVTB = " select cp01,cp02,cp03,cp04,cp09,cp10,cp27,cp83,'1' Typ" & _
            " From caseprogress, staff s1,staff s2" & _
            " where cp27>=" & strDate1 & " and cp27<=" & strDate2 & " and cp159=0 and not (cp10 in ('605','910') and cp123 is null and instr(cp64||' ','智慧局收文文號')=0)" & _
            " and cp01='FCP' and s1.st01(+)=cp83 and s1.st03='F22'" & stCon & _
            " and s2.st01(+)=cp14 and not (cp10='1003' and s2.st03='F22')"

   '寰華-發文人為FCP程序
   'Modified by Morgan 2019/9/18 FMP案C類發文人也是FCP程序
   'strVTB = strVTB & " union all " & _
            " select cp01,cp02,cp03,cp04,cp09,cp10,cp27,cp83,'2' Typ" & _
            " From caseprogress, staff s1,staff s2" & _
            " where cp27>=" & strDate1 & " and cp27<" & strDate2 & " and cp159=0" & _
            " and cp01='P' and cp12 like 'F%' and s1.st01(+)=cp83 and s1.st03='F22'" & stCon & _
            " and s2.st01(+)=cp14 and not (cp10='1003' and s2.st03='F22')"
   strVTB = strVTB & " union all " & _
            " select cp01,cp02,cp03,cp04,cp09,cp10,cp27,cp83,'2' Typ" & _
            " From caseprogress c1, staff s1,staff s2" & _
            " where cp27>=" & strDate1 & " and cp27<=" & strDate2 & " and cp159=0" & _
            " and cp01='P' and cp12 like 'F%' and s1.st01(+)=cp83 and s1.st03='F22'" & stCon & _
            " and s2.st01(+)=cp14 and not (cp10='1003' and s2.st03='F22')" & _
            " and exists(select * from caseprogress x,staff y where x.cp01=c1.cp01 and x.cp02=c1.cp02 and x.cp03=c1.cp03 and x.cp04=c1.cp04 and x.cp31='Y' and y.st01(+)=cp83 and y.st03='F22')"
   'FMP-發文人為FCP程序
   strVTB = strVTB & " union all " & _
            " select cp01,cp02,cp03,cp04,cp09,cp10,cp27,cp83,'3' Typ" & _
            " From caseprogress c1, staff s1,staff s2" & _
            " where cp27>=" & strDate1 & " and cp27<=" & strDate2 & " and cp159=0" & _
            " and cp01='P' and cp12 like 'F%' and s1.st01(+)=cp83 and s1.st03='F22'" & stCon & _
            " and s2.st01(+)=cp14 and not (cp10='1003' and s2.st03='F22')" & _
            " and not exists(select * from caseprogress x,staff y where x.cp01=c1.cp01 and x.cp02=c1.cp02 and x.cp03=c1.cp03 and x.cp04=c1.cp04 and x.cp31='Y' and y.st01(+)=cp83 and y.st03='F22')"
   'end 2019/9/18

   'FMP-發文人非FCP程序
   'Modified by Morgan 2019/4/1 改以完稿日統計並判斷工時>0(不需程序處理工時輸0 Ex.P-120948,949) --敏莉
   'Modified by Morgan 2025/9/19 改算給管制人--敏莉
   'strVTB = strVTB & " union all " & _
            " select cp01,cp02,cp03,cp04,cp09,cp10,cp27,substr(max(fvl05||lpad(fvl06,6,'0')||fvl04),-5) cp83,'3' Typ" & _
            " from engineerprogress,caseprogress,staff s1,fieldvaluelog,staff s2" & _
            " Where ep09>=" & strDate1 & " and ep09<=" & strDate2 & " and cp09(+)=ep02 And cp159 = 0" & _
            " and cp01='P' and cp12 like 'F%' and ep09>0 and cp113>0" & _
            " and s1.st01(+)=cp83 and s1.st03<>'F22'" & _
            " and fvl03(+)=cp09 and fvl01='ENGINEERPROGRESS' and fvl02='EP09'" & _
            " and s2.st01(+)=FVL04 and s2.st03='F22'" & _
            " group by cp01,cp02,cp03,cp04,cp09,cp10,cp27"
   strVTB = strVTB & " union all " & _
            " select cp01,cp02,cp03,cp04,cp09,cp10,cp27,na79 cp83,'3' Typ" & _
            " from engineerprogress,caseprogress,staff s1,patent,fagent,nation,staff s2" & _
            " Where ep09>=" & strDate1 & " and ep09<=" & strDate2 & " and cp09(+)=ep02 And cp159 = 0" & _
            " and cp01='P' and cp12 like 'F%' and ep09>0 and cp113>0" & _
            " and s1.st01(+)=cp83 and s1.st03<>'F22'" & _
            " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
            " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9)" & _
            " and na01(+)=fa10 and s2.st01(+)=na79 and s2.st03='F22'"
   '統計
   If txt1(2) = "1" Then
      'Modified by Morgan 2019/8/5 +不計算點數人員排在後面(只抓在職的統計--淑華)
      strSql = "select nvl(instr(max(x2.oman),st01),0) man2,nvl(instr(max(x1.oman),st01),0) man,st01 cp83,sum(decode(Typ,'2',cpm32,cpm31)) ss,max(st02) st02,max(st20) st20" & _
            " from staff,(" & strVTB & ") T, casepropertymap,setSpecMan x1,setSpecMan x2" & _
            " where st03='F22' and st04='1' and cp83(+)=st01 and cpm01(+)=cp01 and cpm02(+)=cp10 and decode(Typ,'1',cpm31,cpm32)>0" & stCon & _
            " and x1.ocode(+)='外專程序不計發文點數人員' and x2.ocode(+)='外專程序代核稿人員'" & _
            " group by st01 order by man2 asc,man asc, ss desc"
   '明細
   Else
      '改以案件性質統計--敏莉
      'strSql = "select cp83,st02,sqldatet(cp27) cp27,cp01||'-'||cp02 CNo,decode(cp01,'P',cpm04,cpm03) cpm03,cp09,decode(Typ,'2',cpm32,cpm31) cpm31" & _
            " from (" & strVTB & ") T, casepropertymap,staff" & _
            " where cpm01(+)=cp01 and cpm02(+)=cp10 and decode(Typ,'1',cpm31,cpm32)>0 and st01(+)=cp83" & stCon & _
            " order by cp83,cp27,cp09"
      strSql = "select cp83,max(st02) st02,decode(Typ,'1','FCP','2','寰華','3','FMP') TypName,max(decode(cp01,'P',cpm04,cpm03)) cpm03,count(*) ss1,sum(decode(Typ,'2',cpm32,cpm31)) ss2" & _
            " from (" & strVTB & ") T, casepropertymap,staff" & _
            " where cpm01(+)=cp01 and cpm02(+)=cp10 and decode(Typ,'1',cpm31,cpm32)>0 and st01(+)=cp83" & stCon & _
            " group by cp83,Typ,cp10"
   End If
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If txt1(2) = "1" Then
         Report1
      Else
         Report2
      End If
   Else
      MsgBox "無資料！"
   End If
End Sub
'統計
Private Sub Report1()
      
   Dim strManList As String
   Dim intCount As Integer, intHalf As Integer
   Dim dblTot As Double, dblHTot As Double, dblLTot As Double
   Dim strAvg As String, strHigh As String, strLow As String
   Dim intLevel As Integer
      
   With RsTemp
   'Modified by Morgan 2019/8/5
   '總點數及高均低表剔除不計點數人員(考慮可能有季中離職情形,抓實際發文人員)
   'intCount = .RecordCount
   'intHalf = intCount \ 2
   'Do While Not .EOF
   '   dblTot = dblTot + Val("" & RsTemp("ss"))
   '   If .AbsolutePosition <= intHalf Then
   '      dblHTot = dblHTot + Val("" & RsTemp("ss"))
   '   ElseIf .AbsolutePosition > intCount - intHalf Then
   '      dblLTot = dblLTot + Val("" & RsTemp("ss"))
   '   End If
   '   '判發人員(抓主任)
   '   'Modified by Morgan 2019/5/31 +副理(何淑華)
   '   If RsTemp("st20") = "51" Or RsTemp("st20") = "44" Then
   '      strManList = strManList & " " & RsTemp("st02")
   '   End If
   '   .MoveNext
   'Loop
   
   '統計人數
   intCount = 0: dblTot = 0
   strManList = ""
   .MoveFirst
   Do While Not .EOF
      If RsTemp("man") = 0 Then
         dblTot = dblTot + Val("" & RsTemp("ss"))
         intCount = intCount + 1
      Else
         strManList = strManList & " " & RsTemp("st02")
      End If
      .MoveNext
   Loop
   
   intHalf = Round(intCount / 2)
   intI = 0: dblHTot = 0: dblLTot = 0
   .MoveFirst
   Do While Not .EOF
      If RsTemp("man") = 0 Then
         intI = intI + 1
         If intI <= intHalf Then
            dblHTot = dblHTot + Val("" & RsTemp("ss"))
         End If
         
         If intI > intCount - intHalf Then
            dblLTot = dblLTot + Val("" & RsTemp("ss"))
         End If
      End If
      .MoveNext
   Loop
   End With
   strManList = strManList & vbCrLf & "(發文點數不列入計算)"
   'end 2019/8/5
   
   strAvg = Format(dblTot / intCount, "#.##")
   strHigh = Format(dblHTot / intHalf, "#.##")
   strLow = Format(dblLTot / intHalf, "#.##")
        
   If NewDoc() = False Then Exit Sub
     
   With g_WordAp.Application
   '版面設定
   .Selection.PageSetup.PaperSize = wdPaperA4
   .Selection.PageSetup.Orientation = wdOrientPortrait
   .Selection.Orientation = wdTextOrientationHorizontal
   .Selection.Font.Name = "標楷體"
   .Selection.Font.Size = 14
   .Selection.TypeParagraph
   
   '行距
   With .Selection.ParagraphFormat
     .SpaceBefore = 0
     .SpaceAfter = 0
     .LineSpacingRule = wdLineSpaceSingle
     .DisableLineHeightGrid = True
   End With
  
   '新增表格(1*4)
   .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=4
   
   '設定表格高度欄寬
   .Selection.SelectRow
   .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
   .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
   .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(3.8), RulerStyle:=wdAdjustProportional
   .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
   .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.8), RulerStyle:=wdAdjustProportional
   
   .Selection.SelectRow
   .Selection.InsertRows 1
   '框線
   With .Selection.Tables(1)
     .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
     .Borders(wdBorderLeft).LineWidth = wdLineWidth050pt
     .Borders(wdBorderLeft).ColorIndex = wdAuto
     .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
     .Borders(wdBorderRight).LineWidth = wdLineWidth050pt
     .Borders(wdBorderRight).ColorIndex = wdAuto
     .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
     .Borders(wdBorderTop).LineWidth = wdLineWidth050pt
     .Borders(wdBorderTop).ColorIndex = wdAuto
     .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
     .Borders(wdBorderBottom).LineWidth = wdLineWidth050pt
     .Borders(wdBorderBottom).ColorIndex = wdAuto
     .Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
     .Borders(wdBorderHorizontal).LineWidth = wdLineWidth050pt
     .Borders(wdBorderHorizontal).ColorIndex = wdAuto
     .Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
     .Borders(wdBorderVertical).LineWidth = wdLineWidth050pt
     .Borders(wdBorderVertical).ColorIndex = wdAuto
     .Borders.Shadow = False
   End With
   
   .Selection.Collapse Direction:=wdCollapseStart
   'Modified by Morgan 2019/8/2
   'strExc(0) = Txt1(0) & "." & (Txt1(1) * 3 - 2) & "-" & (Txt1(1) * 3) & "月"
   strExc(0) = Left(txt1(0), 3) & "." & Val(Right(txt1(0), 2)) & "-" & Val(Right(txt1(1), 2)) & "月"
   'end 2019/8/2
   strExc(1) = strExc(0) & vbCrLf & "基本季獎金" & vbCrLf & _
               "判發/核稿工作"
   .Selection.TypeText Text:=strExc(1)
   
   .Selection.MoveRight Unit:=wdCharacter, Count:=1
   .Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
   .Selection.Cells.Merge
   .Selection.TypeText Text:=Trim(strManList)
   
   .Selection.MoveDown Unit:=wdLine, Count:=1
   .Selection.SelectRow
   .Selection.Collapse Direction:=wdCollapseStart
   
   'Modified by Morgan 2019/8/2
   'strExc(1) = (txt1(1) * 3 - 2) & "-" & (txt1(1) * 3) & "月績效獎金"
   strExc(1) = Val(Right(txt1(0), 2)) & "-" & Val(Right(txt1(1), 2)) & "月績效獎金"
   'end 2019/8/2
   .Selection.TypeText Text:=strExc(1)
   
   .Selection.MoveRight Unit:=wdCharacter, Count:=2
   strExc(0) = "新案及中間程序發文"
   .Selection.TypeText Text:=strExc(0)
   
   .Selection.MoveRight Unit:=wdCharacter, Count:=1
   .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
   
   strExc(0) = "總點數: " & dblTot & "/" & intCount & "人" & vbCrLf & _
               "高標: " & strHigh & vbCrLf & _
               "均標: " & strAvg & vbCrLf & _
               "低標: " & strLow
   .Selection.TypeText Text:=strExc(0)
   .Selection.MoveRight Unit:=wdCharacter, Count:=1
   RsTemp.MoveFirst
   Do While Not RsTemp.EOF
      .Selection.InsertRows 1
      .Selection.Collapse Direction:=wdCollapseStart
      If RsTemp("ss") >= Val(strHigh) Then
         If intLevel <> 1 Then
            strExc(0) = "75-100% " & vbCrLf & "(2 個基數)"
            .Selection.TypeText Text:=strExc(0)
            intLevel = 1: intI = 0
         End If
      ElseIf RsTemp("ss") >= Val(strAvg) Then
         If intLevel <> 2 Then
            If intI > 1 Then
               .Selection.MoveUp Unit:=wdLine, Count:=intI + 1
               .Selection.MoveDown Unit:=wdLine, Count:=intI, Extend:=wdExtend
               .Selection.Cells.Merge
               .Selection.SelectRow
               .Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdMove
            End If
            strExc(0) = "50-75% " & vbCrLf & "(1.5 個基數)"
            .Selection.TypeText Text:=strExc(0)
            intLevel = 2: intI = 0
         End If
      ElseIf RsTemp("ss") >= Val(strLow) Then
         If intLevel <> 3 Then
            If intI > 1 Then
               .Selection.MoveUp Unit:=wdLine, Count:=intI + 1
               .Selection.MoveDown Unit:=wdLine, Count:=intI, Extend:=wdExtend
               .Selection.Cells.Merge
               .Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdMove
            End If
            strExc(0) = "25-50% " & vbCrLf & "(1 個基數)"
            .Selection.TypeText Text:=strExc(0)
            intLevel = 3: intI = 0
         End If
         
      ElseIf RsTemp("man") = 0 Then
         If intLevel <> 4 Then
            If intI > 1 Then
               .Selection.MoveUp Unit:=wdLine, Count:=intI + 1, Extend:=wdMove
               .Selection.MoveDown Unit:=wdLine, Count:=intI, Extend:=wdExtend
               .Selection.Cells.Merge
               .Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdMove
            End If
            strExc(0) = "0-25% " & vbCrLf & "(不發獎金)"
            .Selection.TypeText Text:=strExc(0)
            intLevel = 4: intI = 0
         End If
      
      'Added by Morgan 2019/8/5
      ElseIf RsTemp("man2") = 0 Then
         If intLevel <> 5 Then
            If intI > 1 Then
               .Selection.MoveUp Unit:=wdLine, Count:=intI + 1, Extend:=wdMove
               .Selection.MoveDown Unit:=wdLine, Count:=intI, Extend:=wdExtend
               .Selection.Cells.Merge
               .Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdMove
            End If
            strExc(0) = "判發主管" & vbCrLf & "(1.5 個基數)"
            .Selection.TypeText Text:=strExc(0)
            intLevel = 5: intI = 0
         End If
         
      Else
         If intLevel <> 6 Then
            If intI > 1 Then
               .Selection.MoveUp Unit:=wdLine, Count:=intI + 1, Extend:=wdMove
               .Selection.MoveDown Unit:=wdLine, Count:=intI, Extend:=wdExtend
               .Selection.Cells.Merge
               .Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdMove
            End If
            strExc(0) = "代核稿案件" & vbCrLf & "(1.5 個基數)"
            .Selection.TypeText Text:=strExc(0)
            intLevel = 6: intI = 0
         End If
      'end 2019/8/5
      End If
      intI = intI + 1
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.TypeText Text:=RsTemp("st02")
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.TypeText Text:=RsTemp("ss")
      If intI = 1 Then
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.Font.ColorIndex = wdRed
         If intLevel = 1 Then
            .Selection.TypeText Text:=">= 高標"
         ElseIf intLevel = 2 Then
            .Selection.TypeText Text:="介於高標和均標中間"
         ElseIf intLevel = 3 Then
            .Selection.TypeText Text:="介於均標和低標中間"
         ElseIf intLevel = 4 Then
            .Selection.TypeText Text:="<= 低標"
         End If
         .Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
      Else
         .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdMove
      End If
      RsTemp.MoveNext
   Loop
   
   If intI > 1 Then
      .Selection.SelectRow
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.MoveUp Unit:=wdLine, Count:=intI, Extend:=wdMove
      .Selection.MoveDown Unit:=wdLine, Count:=intI, Extend:=wdExtend
      .Selection.Cells.Merge
      .Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdMove
   End If
   
   If txt1(3) = "1" Then
      .Activate
   Else
      .ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
      .Quit wdDoNotSaveChanges
   End If
   End With
   Set g_WordAp = Nothing
End Sub

'明細
Private Sub Report2()
   Dim dblTot As Double
   Dim stNo As String, stName As String
   
   intI = 0
   RsTemp.MoveFirst
   stNo = RsTemp("cp83")
   stName = RsTemp("st02")
   If NewDoc() = False Then Exit Sub
   PrintHead stName
   Do While Not RsTemp.EOF
      If stNo <> RsTemp("cp83") Then
         PrintSum dblTot
         
         If NewDoc() = False Then Exit Sub
         intI = 0
         dblTot = 0
         stNo = RsTemp("cp83")
         stName = RsTemp("st02")
         PrintHead stName
      End If
      
      intI = intI + 1
      With g_WordAp.Application
      If intI = 1 Then
         AddTable
      Else
         .Selection.InsertRows 1
      End If
      .Selection.Collapse Direction:=wdCollapseStart
      '.Selection.TypeText Text:=RsTemp("cp27")
      .Selection.TypeText Text:=RsTemp("TypName")
      .Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
      '.Selection.TypeText Text:=RsTemp("CNo")
      .Selection.TypeText Text:=RsTemp("cpm03")
      .Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
      '.Selection.TypeText Text:=RsTemp("cpm03")
      .Selection.TypeText Text:=RsTemp("ss1")
      .Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
      '.Selection.TypeText Text:=RsTemp("cpm31")
      .Selection.TypeText Text:=RsTemp("ss2")
      .Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
      End With
      'dblTot = dblTot + RsTemp("cpm31")
      dblTot = dblTot + RsTemp("ss2")
      RsTemp.MoveNext
   Loop
   
   PrintSum dblTot, True
End Sub

Private Sub AddTable()
   With g_WordAp.Application
   '新增表格(1*4)
   .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=4
   '設定表格高度欄寬
   .Selection.SelectRow
   .Selection.Font.Name = "標楷體"
   .Selection.Font.Size = 12
   .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
   .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
   '.Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(2.7), RulerStyle:=wdAdjustProportional
   '.Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(3.6), RulerStyle:=wdAdjustProportional
   '.Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(6.4), RulerStyle:=wdAdjustProportional
   .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(2.7), RulerStyle:=wdAdjustProportional
   .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(6.4), RulerStyle:=wdAdjustProportional
   .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.7), RulerStyle:=wdAdjustProportional
   With .Selection.Tables(1)
     .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
     .Borders(wdBorderRight).LineStyle = wdLineStyleNone
     .Borders(wdBorderTop).LineStyle = wdLineStyleNone
     .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
     .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
     .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
     .Borders.Shadow = False
   End With
   End With
End Sub
'表頭
Private Sub PrintHead(pName As String)
   With g_WordAp.Application
   .Visible = False
   '版面設定
   .Selection.PageSetup.PaperSize = wdPaperA4
   .Selection.PageSetup.Orientation = wdOrientPortrait
   .Selection.Orientation = wdTextOrientationHorizontal
   .Selection.Font.Name = "標楷體"
   .Selection.Font.Size = 12
   '.Selection.TypeParagraph
   
   '行距
   With .Selection.ParagraphFormat
     .SpaceBefore = 0
     .SpaceAfter = 0
     .LineSpacingRule = wdLineSpaceSingle
     .DisableLineHeightGrid = True
   End With
   
   '切換至頁首
   .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
   
   AddTable
   
   .Selection.InsertRows 1
   .Selection.Cells.Merge
   .Selection.Font.Bold = True
   .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
   'Modified by Morgan 2019/9/18
   'strExc(0) = txt1(0) & "年第" & txt1(1) & "季點數明細(" & pName & ")"
   strExc(0) = Left(txt1(0), 3) & "年" & Val(Right(txt1(0), 2)) & "-" & Val(Right(txt1(1), 2)) & "月點數明細(" & pName & ")"
   'end 2019/9/18
   .Selection.Font.Size = 16
   .Selection.ParagraphFormat.SpaceAfter = 8
   .Selection.TypeText Text:=strExc(0)
   .Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdMove
   .Selection.SelectRow
   
   .Selection.InsertRows 1
   .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
   .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(10.2), RulerStyle:=wdAdjustProportional
   .Selection.Collapse Direction:=wdCollapseStart
   .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
   .Selection.TypeText Text:="製表人員：" & strUserName
   
   .Selection.MoveRight Unit:=wdCharacter, Count:=1
   .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
   .Selection.TypeText Text:="製表日期：" & Format(strSrvDate(2), "##/##/##")
   
   .Selection.MoveRight Unit:=wdCharacter, Count:=1
   .Selection.InsertRows 1
   .Selection.Collapse Direction:=wdCollapseStart
   .Selection.MoveRight Unit:=wdCharacter, Count:=1
   .Selection.TypeText Text:="頁    次：    "
   .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldEmpty, Text:="Page", PreserveFormatting:=True
      
   .Selection.MoveRight Unit:=wdCharacter, Count:=2
   .Selection.SelectRow
   With .Selection.Cells.Borders(wdBorderBottom)
      .LineStyle = wdLineStyleSingle
      .LineWidth = wdLineWidth050pt
      .ColorIndex = wdAuto
   End With
   .Selection.ParagraphFormat.SpaceBefore = 6
   .Selection.Collapse Direction:=wdCollapseStart
   '.Selection.TypeText Text:="發文日"
   .Selection.TypeText Text:="案件別"
   .Selection.MoveRight Unit:=wdCharacter, Count:=1
   '.Selection.TypeText Text:="本所案號"
   .Selection.TypeText Text:="案件性質"
   .Selection.MoveRight Unit:=wdCharacter, Count:=1
   '.Selection.TypeText Text:="案件性質"
   .Selection.TypeText Text:="件數"
   .Selection.MoveRight Unit:=wdCharacter, Count:=1
   .Selection.TypeText Text:="點數"
   .Selection.SelectRow
   .Selection.Collapse Direction:=wdCollapseStart
   .Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdMove
   .Selection.ParagraphFormat.LineSpacing = .LinesToPoints(0.06)
   .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
   End With
End Sub
'表尾
Private Sub PrintSum(pSum As Double, Optional pShow As Boolean)
   With g_WordAp.Application
   .Selection.InsertRows 1
   With .Selection.Cells.Borders(wdBorderTop)
      .LineStyle = wdLineStyleSingle
      .LineWidth = wdLineWidth050pt
      .ColorIndex = wdAuto
   End With
   .Selection.Collapse Direction:=wdCollapseStart
   .Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
   .Selection.Cells.Merge
   .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
   .Selection.Font.Bold = True
   .Selection.TypeText Text:="合計"
   .Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
   .Selection.Font.Bold = True
   .Selection.TypeText Text:=Format(pSum, "#.##")
   
   If txt1(3) = "1" Then
      .Visible = True
      If pShow Then .Activate
   Else
      .ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
      .Quit wdDoNotSaveChanges
   End If
   End With
   Set g_WordAp = Nothing
End Sub
Private Function TxtValidate() As Boolean
   
   If Trim(txt1(0)) = "" Then
      'Modified by Morgan 2019/8/2
      'MsgBox "請輸入年！", vbExclamation
      MsgBox "請輸起始年月！", vbExclamation
      txt1(0).SetFocus
      Exit Function
   End If
   
   If Trim(txt1(1)) = "" Then
      'Modified by Morgan 2019/8/2
      'MsgBox "請輸入季！", vbExclamation
      MsgBox "請輸迄止年月！", vbExclamation
      txt1(1).SetFocus
      Exit Function
   End If
   
   'Added by Morgan 2019/8/2
   If Left(txt1(0), 3) <> Left(txt1(1), 3) Then
      MsgBox "起迄年度必須相同！", vbExclamation
      txt1(1).SetFocus
      Exit Function
   End If
   'end 2019/8/2
   
   If txtNo <> "" Then
      If lbl1 = "" Then
         MsgBox "員工編號錯誤！", vbExclamation
         txtNo.SetFocus
         Exit Function
      ElseIf GetStaffDepartment(txtNo) <> "F22" Then
         MsgBox "該員工編號不為外專程序！", vbExclamation
         txtNo.SetFocus
         Exit Function
      End If
   End If
   
   If Trim(txt1(2)) = "" Then
      If txt1(2).Enabled Then
         MsgBox "請輸入報表內容！", vbExclamation
         txt1(2).SetFocus
         Exit Function
      End If
   End If
   
   If Trim(txt1(3)) = "" Then
      MsgBox "請輸入顯示方式！", vbExclamation
      txt1(3).SetFocus
      Exit Function
   End If
   
   TxtValidate = True
End Function

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
   If Index = 1 Then
      If txt1(0).Text <> txt1(0).Tag Then
         txt1(1).Text = txt1(0).Text
      End If
   End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 Then
      If Not IsNumeric(Chr(KeyAscii)) Then
         KeyAscii = 0
         Beep
         Exit Sub
      End If
      
      Select Case Index
         'Modified by Morgan 2019/8/2
         'Case 1 '季
         '   If Chr(KeyAscii) < "1" Or Chr(KeyAscii) > "4" Then
         '      KeyAscii = 0
         '      Beep
         '   End If
         Case 0, 2
            If Not IsNumeric(Chr(KeyAscii)) Then
               KeyAscii = 0
               Beep
            End If
         'end 2019/8/2
         Case 2, 3 '報表內容, 顯示方式
            If Chr(KeyAscii) < "1" Or Chr(KeyAscii) > "2" Then
               KeyAscii = 0
               Beep
            End If
      End Select
   End If
End Sub

Private Sub txtNo_Change()
   If txtNo <> "" Then
      txt1(2) = "2"
      txt1(2).Enabled = False
   Else
      txt1(2) = "1"
      txt1(2).Enabled = True
   End If
End Sub

Private Sub txtNo_GotFocus()
   TextInverse txtNo
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtNo_Validate(Cancel As Boolean)
   lbl1 = GetStaffName(txtNo)
End Sub

Private Function NewDoc() As Boolean

   Dim iResumeCnt As Integer
   
On Error GoTo ErrHnd
   
   If TypeName(g_WordAp) <> "Application" Then
      Set g_WordAp = New Word.Application
   End If
   g_WordAp.Documents.add
   '不顯示可能會有問題
   g_WordAp.Visible = True
   'g_WordAp.ActiveWindow.ActivePane.View.Zoom.Percentage = 100
   NewDoc = True
   Exit Function
   
ErrHnd:
   'Resume
   If Err.Number <> 0 Then
      If iResumeCnt > 3 Then
         MsgBox "錯誤 : " & Err.Description, vbCritical
      Else
         iResumeCnt = iResumeCnt + 1
         Select Case Err.Number
            Case 91:
               g_WordAp.Documents.add
               Resume Next
            Case 462:
               Set g_WordAp = New Word.Application
               Resume
            Case Else:
               MsgBox "錯誤 : " & Err.Description, vbCritical
         End Select
      End If
   End If
End Function
