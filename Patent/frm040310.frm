VERSION 5.00
Begin VB.Form frm040310 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文簿"
   ClientHeight    =   3825
   ClientLeft      =   2925
   ClientTop       =   2820
   ClientWidth     =   3720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   3720
   Begin VB.TextBox txtPA46 
      Height          =   285
      Left            =   1710
      TabIndex        =   8
      Top             =   2100
      Width           =   285
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1788
      MaxLength       =   1
      TabIndex        =   10
      Top             =   2760
      Width           =   330
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2808
      TabIndex        =   12
      Top             =   24
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1992
      TabIndex        =   11
      Top             =   24
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1104
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2460
      Width           =   210
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   2256
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1752
      Width           =   1065
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1104
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1752
      Width           =   1065
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1104
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1440
      Width           =   1065
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2256
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1116
      Width           =   1065
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1104
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1116
      Width           =   1065
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2256
      MaxLength       =   7
      TabIndex        =   2
      Top             =   804
      Width           =   1065
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1104
      MaxLength       =   7
      TabIndex        =   1
      Top             =   804
      Width           =   1065
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1104
      TabIndex        =   0
      Top             =   480
      Width           =   2235
   End
   Begin VB.Label Label1 
      Caption         =   "PS : 承辦人條件未輸入時，列印資料不含　　   承辦人為陳玲玲及莊敏惠的發文資料"
      ForeColor       =   &H00008000&
      Height          =   540
      Index           =   8
      Left            =   120
      TabIndex        =   23
      Top             =   3120
      Width           =   3420
   End
   Begin VB.Label Label1 
      Caption         =   "PCT進入國家階段：　（Y：國家階段）"
      Height          =   180
      Index           =   7
      Left            =   180
      TabIndex        =   22
      Top             =   2160
      Width           =   3180
   End
   Begin VB.Label Label1 
      Caption         =   "是否依承辦人跳頁：            (Y : 是)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   9
      Left            =   120
      TabIndex        =   21
      Top             =   2820
      Width           =   3180
   End
   Begin VB.Line Line3 
      X1              =   2172
      X2              =   2262
      Y1              =   1872
      Y2              =   1872
   End
   Begin VB.Line Line2 
      X1              =   2112
      X2              =   2262
      Y1              =   1224
      Y2              =   1224
   End
   Begin VB.Line Line1 
      X1              =   2172
      X2              =   2262
      Y1              =   948
      Y2              =   948
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   2268
      TabIndex        =   20
      Top             =   1476
      Width           =   1392
   End
   Begin VB.Label Label1 
      Caption         =   "(1.承辦人 2.發文日 3.所別 )"
      Height          =   180
      Index           =   6
      Left            =   1410
      TabIndex        =   19
      Top             =   2490
      Width           =   2235
   End
   Begin VB.Label Label1 
      Caption         =   "列印順序："
      Height          =   180
      Index           =   5
      Left            =   150
      TabIndex        =   18
      Top             =   2490
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   4
      Left            =   144
      TabIndex        =   17
      Top             =   1800
      Width           =   936
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   3
      Left            =   324
      TabIndex        =   16
      Top             =   1488
      Width           =   756
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   2
      Left            =   156
      TabIndex        =   15
      Top             =   1164
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "發文日："
      Height          =   180
      Index           =   1
      Left            =   312
      TabIndex        =   14
      Top             =   852
      Width           =   768
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   156
      TabIndex        =   13
      Top             =   528
      Width           =   912
   End
End
Attribute VB_Name = "frm040310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'整理 by Morgan 2005/9/28
Option Explicit
Dim i As Integer, j As Integer, s As VbMsgBoxResult
Dim strTemp(0 To 14) As String, strTemp1 As Variant, strTemp2 As Variant
Dim PLeft(0 To 15) As Integer, iPrint As Integer, iPage As Integer, strTemp3(0 To 3) As String
Dim blnClkSure As Boolean '判斷是否按下確定按鈕
Dim m_iMode As Integer '報表別

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
        If Len(txt1(0)) = 0 Then
           s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
           txt1(0).SetFocus
           Exit Sub
        Else
            If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
               Me.txt1(1).SetFocus
               txt1_GotFocus 1
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
               Me.txt1(2).SetFocus
               txt1_GotFocus 2
               Exit Sub
            End If
            If Me.txt1(1).Text <> "" And Me.txt1(2).Text <> "" Then
               If Val(Me.txt1(1).Text) > Val(Me.txt1(2).Text) Then
                  MsgBox "發文日範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(1).SetFocus
                  txt1_GotFocus 1
                  Exit Sub
               End If
            End If
            
            If Len(txt1(2)) = 0 Then
               s = MsgBox("發文日期區間不可空白!!", , "USER 輸入錯誤")
               txt1(1).SetFocus
               txt1_GotFocus (1)
               Exit Sub
            Else
               If Me.txt1(3).Text <> "" And Me.txt1(4).Text <> "" Then
                  If Me.txt1(3).Text > Me.txt1(4).Text Then
                     MsgBox "案件性質範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(3).SetFocus
                     txt1_GotFocus 3
                     Exit Sub
                  End If
               End If
               If Me.txt1(6).Text <> "" And Me.txt1(7).Text <> "" Then
                  If Me.txt1(6).Text > Me.txt1(7).Text Then
                     MsgBox "申請國家範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(6).SetFocus
                     txt1_GotFocus 6
                     Exit Sub
                  End If
               End If
               lbl1.Caption = GetPrjSalesNM(txt1(5))
               If Trim(lbl1.Caption) = "" And Trim(txt1(5)) <> "" Then
                   MsgBox "承辦人代號錯誤，請重新輸入 !"
                   txt1(5).SetFocus
                   txt1_GotFocus 5
                   Exit Sub
               End If
                
                If Len(txt1(8)) = 0 Then
                   s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
                   txt1(8).SetFocus
                   Exit Sub
                Else
                    Screen.MousePointer = vbHourglass
                    Me.Enabled = False
                    ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/2 清除查詢印表記錄檔欄位
                    Process
                    Me.Enabled = True
                    Screen.MousePointer = vbDefault
                End If
            End If
        End If
      Case 1
         Unload Me
      Case Else
   End Select
End Sub

'發文簿
Sub Process()
   Dim stVTB1 As String, stVTB2 As String
   Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String
   
   strSQL1 = ""
   strSQL2 = ""
   If Len(Trim(txt1(1))) <> 0 Then
      strSQL1 = strSQL1 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(1))) & " "
      StrSQL3 = StrSQL3 + " AND SH01>=" & Val(ChangeTStringToWString(txt1(1))) & " "
   End If
   If Len(Trim(txt1(2))) <> 0 Then
      strSQL1 = strSQL1 & " AND CP27<=" & Val(ChangeTStringToWString(txt1(2))) & " "
      StrSQL3 = StrSQL3 + " AND SH01<=" & Val(ChangeTStringToWString(txt1(2))) & " "
   End If
   If Len(Trim(txt1(1))) <> 0 Or Len(Trim(txt1(2))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/2
   End If
   If Len(txt1(3)) <> 0 Then
       strSQL1 = strSQL1 + " AND CP10>='" & txt1(3) & "' "
   End If
   If Len(txt1(4)) <> 0 Then
       strSQL1 = strSQL1 + " AND CP10<='" & txt1(4) & "' "
   End If
   If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/12/2
   End If
   If Len(txt1(5)) <> 0 Then
       strSQL1 = strSQL1 + " AND CP14='" & txt1(5) & "' "
       StrSQL3 = StrSQL3 + " AND SH02='" & txt1(5) & "' "
       pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(5) & lbl1 'Add By Sindy 2010/12/2
   'Modify by Morgan 2005/11/1 加控制不指定承辦人時不印部門為P12(94003除外),P13,M71的
   '除非有指定承辦人, 否則承辦人為陳玲玲及莊敏惠的資料不印
   Else
      'strSQL1 = strSQL1 + " AND CP14<>'81002' AND CP14<>'73017' "
       '2006/4/3 MODIFY BY SONIA 加入95008
       'strSQL1 = strSQL1 + " AND (ST01='94003' OR ST03 NOT IN ('P12','P13','M71')) "
       '2006/7/10 MODIFY BY SONIA P13,M71不印,81002,73017不印
       'strSQL1 = strSQL1 + " AND (ST01='94003' OR ST01='95008' OR ST03 NOT IN ('P12','P13','M71')) "
       'Modify by Morgan 2006/10/4 又改成P12不印(等級73除外) --郭,CFP的程序也不要印
       'strSQL1 = strSQL1 + " AND (ST01<>'81002' AND ST01<>'73017' AND ST03 NOT IN ('P13','M71')) "
       strSQL1 = strSQL1 + " AND (ST05='73' OR ST03 NOT IN ('P12','P13','M71')) "
   End If

   strSQL2 = strSQL1
   If Len(txt1(0)) <> 0 Then
      strSQL1 = strSQL1 + " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
      strSQL2 = strSQL2 + " and CP01 in (" & SQLGrpStr(txt1(0), 5) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0)  'Add By Sindy 2010/12/2
   End If
   If Len(txt1(6)) <> 0 Then
       strSQL1 = strSQL1 + " AND SUBSTR(pa09,1,3)>='" & txt1(6) & "' "
       strSQL2 = strSQL2 + " AND SUBSTR(sp09,1,3)>='" & txt1(6) & "' "
   End If
   If Len(txt1(7)) <> 0 Then
       strSQL1 = strSQL1 + " AND SUBSTR(pa09,1,3)<='" & txt1(7) & "' "
       strSQL2 = strSQL2 + " AND SUBSTR(sp09,1,3)<='" & txt1(7) & "' "
   End If
   If Len(txt1(6)) <> 0 Or Len(txt1(7)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(6) & "-" & txt1(7) 'Add By Sindy 2010/12/2
   End If
   If txtPA46 = "Y" Then
      strSQL1 = strSQL1 & " And PA09<>'056' AND PA46='Y' "
      pub_QL05 = pub_QL05 & ";" & Left(Label1(7), 10) & txtPA46 'Add By Sindy 2010/12/2
   End If
   'Add by Morgan 2007/8/2 重新委任928不要印--慧汶
   strSQL1 = strSQL1 & " And CP10<>'928'"
   
   'Modify by Morgan 2005/10/20 為了要與承辦人達成情形一致不再作限制
   '只印A,B類, 案件性質為閉卷(913)或不續辦(907)的不印
   'Modify by Morgan 2011/5/30 若有建點數分配資料時點數改分配點數(目前有225提供書狀意見及226配合開庭)
   stVTB1 = "SELECT CP14,PA09,CP27,CP01,CP02,CP03,CP04,NVL(PA05,NVL(PA06,PA07)) PA06" & _
      ",PA08,CP10,SUBSTR(PA26,1,8) C01,SUBSTR(PA26,9,1) C02,CP05,CP15,CP26,nvl(a0n03/1000,CP18) cp18,CP97,CP98,SCR02" & _
      " FROM CASEPROGRESS,PATENT,SpecialCaseRecord,STAFF,acc0n0" & _
      " WHERE ST01(+)=CP14" & strSQL1 & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
      " And SCR01(+)=CP09 and SCR03(+)='V' and a0n02(+)=cp09" & _
      " UNION ALL" & _
      " SELECT CP14,sp09,CP27,CP01,CP02,CP03,CP04,NVL(SP05,NVL(SP06,SP07)) PA05" & _
      ",NULL,CP10,SUBSTR(sP08,1,8) CU01,SUBSTR(sP08,9,1) CU02,CP05,CP15,CP26,nvl(a0n03/1000,CP18) CP18,CP97,CP98,SCR02" & _
      " FROM CASEPROGRESS,SERVICEPRACTICE,SpecialCaseRecord,STAFF,acc0n0" & _
      " WHERE ST01(+)=CP14" & strSQL2 & _
      " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
      " And SCR01(+)=CP09 and SCR03(+)='V' and a0n02(+)=cp09"
      
   'Modified by Morgan 2014/3/20 2014/4/1 起支援改每小時折計0.2基數
   'stVTB2 = "SELECT SH02,Sum(Round(Decode(SH06, 'CFP', Nvl(SH05, 0)/3, Nvl(SH05, 0)/4) ,2)) SH05" & _
      " FROM SupportHour WHERE 1=1" & StrSQL3 & _
      " GROUP BY SH02"
   'Modified by Morgan 2019/4/9 108考核支援時數轉換要除組別參數
   'stVTB2 = "SELECT SH02,Sum(Round(" & Sh2EPtCode & " ,2)) SH05" & _
      " FROM SupportHour,staff WHERE st01(+)=sh02 " & StrSQL3 & _
      " GROUP BY SH02"
   stVTB2 = "SELECT SH02,Sum(Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2)) SH05" & _
      " FROM SupportHour,staff WHERE st01(+)=sh02 " & StrSQL3 & _
      " GROUP BY SH02"
   'end 2014/3/20
   
   strSql = "SELECT NVL(ST02,CP14) C00,Decode(ST06,'1','北所','2','中所','3','南所','4','高所','其他') C01" & _
      "," & SQLDate("CP27") & " C02,CP01||'-'||CP02||'-'||CP03||'-'||CP04 C03" & _
      ",PA06 C04,decode(pa09,'000',ptm03,ptm04) C05,decode(PA09,'000',cpm03,cpm04) C06" & _
      ",NVL(N1.NA03,N1.NA04) C07,NVL(N2.NA03,N2.NA04) C08," & SQLDate("CP05") & " C09" & _
      ",NVL(CP26,'Y') CP26,CP18,SCR02,CP98,CP97,SH05,CP98*CP97 XC1" & _
      ",CP10,ST06,CP14,PA09,CP27,CP01,CP02,CP03,CP04" & _
      " FROM (" & stVTB1 & ") X,(" & stVTB2 & ") Y,CASEPROPERTYMAP,PATENTTRADEMARKMAP" & _
      ",CUSTOMER, NATION N1,NATION N2,STAFF" & _
      " WHERE cpm01(+)=CP01 AND cpm02(+)=CP10 AND PTM01(+)=1 AND PTM02(+)=PA08" & _
      " AND ST01(+)=CP14 AND N1.NA01(+)=CU10 AND CU01(+)=C01 AND CU02(+)=C02" & _
      " AND N2.NA01(+)=PA09 AND SH02(+)=CP14"
      
   m_iMode = Val(txt1(8))
   
   Select Case m_iMode
      Case 1 '列印順序依承辦人
         strSql = strSql & " ORDER BY ST06,CP14,PA09,CP27 DESC,CP01,CP02,CP03,CP04,CP10"
         pub_QL05 = pub_QL05 & ";" & Label1(5) & "1.承辦人" 'Add By Sindy 2010/12/2
      Case 2 '列印順序依發文日
         strSql = strSql & " ORDER BY CP27,CP01,CP02,CP03,CP04,CP10"
         pub_QL05 = pub_QL05 & ";" & Label1(5) & "2.發文日" 'Add By Sindy 2010/12/2
      Case 3 '列印順序依所別
         strSql = strSql & " ORDER BY ST06,CP14,PA09,CP27 DESC,CP01,CP02,CP03,CP04,CP10"
         pub_QL05 = pub_QL05 & ";" & Label1(5) & "3.所別" 'Add By Sindy 2010/12/2
   End Select
   
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If .RecordCount = 0 Then
         InsertQueryLog (0) 'Add By Sindy 2010/12/2
         ShowNoData
      Else
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/2
         Select Case m_iMode
            Case 1 '列印順序依承辦人
               PrintData1 adoRecordset
            Case 2 '列印順序依發文日
               PrintData2 adoRecordset
            Case 3 '列印順序依所別
               PrintData3 adoRecordset
         End Select
         ShowPrintOk
      End If
   End With
End Sub

'列印順序依承辦人排序
Sub PrintData1(ByVal p_Rst As ADODB.Recordset)

   Dim dblVal As Double '加權計件值小計
   Dim dblValTot As Double '加權計件值總計
   Dim dblSh As Double '支援時數
   Dim dblShTot As Double '支援時數總計

   iPage = 1
   dblVal = 0
   dblValTot = 0
   dblSh = 0
   dblShTot = 0
   With p_Rst
      .MoveFirst
      PrintTitle
      '前筆承辦人
      strTemp3(0) = CheckStr(.Fields(0))
      dblSh = Val("" & .Fields("SH05"))
      dblShTot = dblShTot + dblSh
      '前筆所別
      strTemp3(1) = CheckStr(.Fields(1))
      '前筆承辦人代碼
      strTemp3(3) = "" & .Fields("CP14")
      PrintSubTitle
       
      Do While .EOF = False
         For i = 0 To 14
            strTemp(i) = CheckStr(.Fields(i))
         Next i
         '案件名稱
         strTemp(4) = StrToStr(strTemp(4), 14)
         '專利種類
         strTemp(5) = StrToStr(strTemp(5), 4)
         '案件性質
         strTemp(6) = StrToStr(strTemp(6), 6)
         '客戶國籍
         strTemp(7) = StrToStr(strTemp(7), 4)
         '申請國家
         strTemp(8) = StrToStr(strTemp(8), 4)
         '點數
         strTemp(11) = Format(strTemp(11), "###.0")
         '特殊案件
         strTemp(12) = Format(strTemp(12), "###.0")
         '加乘註記
         strTemp(13) = Format(strTemp(13), "###.0")
         '計件值
         strTemp(14) = Format(strTemp(14), "###.0")
           
         '若承辦人不同
         If strTemp3(0) <> strTemp(0) Then
            PrintSubTot p_Rst, dblVal, strTemp3(3), dblSh
            dblValTot = dblValTot + dblVal
            dblSh = Val("" & .Fields("SH05"))
            dblShTot = dblShTot + dblSh
            '若設定依承辦人跳頁且非第一筆資料時
            If Me.txt1(9).Text = "Y" Then
               iPage = iPage + 1
               Printer.NewPage
               PrintTitle
            '若所別不同
            ElseIf strTemp3(1) <> strTemp(1) Then
               iPage = iPage + 1
               Printer.NewPage
               PrintTitle
            End If
            '前筆承辦人
            strTemp3(0) = strTemp(0)
            '前筆所別
            strTemp3(1) = strTemp(1)
            '前筆承辦人代碼
            strTemp3(3) = "" & .Fields("CP14")
            dblVal = 0
            dblSh = 0
            PrintSubTitle
         End If
         NewLine False
         PrintDetail
         dblVal = dblVal + Val("" & .Fields("XC1")) + Val("" & .Fields("SCR02"))
         .MoveNext
      Loop
      PrintSubTot p_Rst, dblVal, strTemp3(3), dblSh
      dblValTot = dblValTot + dblVal
   End With
   If Me.txt1(5).Text = "" Then
      '最後一頁
      iPage = iPage + 1
      Printer.NewPage
      strTemp3(0) = ""
      strTemp3(1) = ""
      PrintTitle
      PrintSubTitle
      PrintSubTot p_Rst, dblValTot, Empty, dblShTot
   End If
   Printer.EndDoc
   
End Sub

'列印順序依發文日排序
Sub PrintData2(ByVal p_Rst As ADODB.Recordset)
   
   Dim dblVal As Double '加權計件值小計
   Dim dblValTot As Double '加權計件值總計
   Dim bolPrint As Boolean '是否列印發文日
   
   iPage = 1
   dblVal = 0
   dblValTot = 0
   bolPrint = True
   With p_Rst
      .MoveFirst
      PrintTitle
      
      '前筆發文日
      strTemp3(3) = CheckStr(.Fields("CP27"))
      PrintSubTitle
       
      Do While .EOF = False
         For i = 0 To 14
            strTemp(i) = CheckStr(.Fields(i))
         Next i
         '案件名稱
         strTemp(4) = StrToStr(strTemp(4), 14)
         '專利種類
         strTemp(5) = StrToStr(strTemp(5), 4)
         '案件性質
         strTemp(6) = StrToStr(strTemp(6), 6)
         '客戶國籍
         strTemp(7) = StrToStr(strTemp(7), 4)
         '申請國家
         strTemp(8) = StrToStr(strTemp(8), 4)
         '點數
         strTemp(11) = Format(strTemp(11), "###.0")
         '特殊案件
         strTemp(12) = Format(strTemp(12), "###.0")
         '加乘註記
         strTemp(13) = Format(strTemp(13), "###.0")
         '計件值
         strTemp(14) = Format(strTemp(14), "###.0")
         '若發文日不同
         If strTemp3(3) <> CheckStr(.Fields("CP27")) Then
            PrintSubTot p_Rst, dblVal, strTemp3(3)
            dblValTot = dblValTot + dblVal
            '前筆發文日
            strTemp3(3) = CheckStr(.Fields("CP27"))
            dblVal = 0
            PrintSubTitle
            bolPrint = True
         End If
         NewLine False
         PrintDetail bolPrint
         'dblVal = dblVal + Val(strTemp(13)) * Val(strTemp(14))
         dblVal = dblVal + Val("" & .Fields("XC1")) + Val("" & .Fields("SCR02"))
         .MoveNext
         bolPrint = False
      Loop
      PrintSubTot p_Rst, dblVal, strTemp3(3)
      dblValTot = dblValTot + dblVal
   End With
   '最後一頁
   iPage = iPage + 1
   Printer.NewPage
   strTemp3(0) = ""
   strTemp3(1) = ""
   PrintTitle
   PrintSubTitle
   PrintSubTot p_Rst, dblValTot
   Printer.EndDoc
   
End Sub

'列印順序依所別排序
Sub PrintData3(ByVal p_Rst As ADODB.Recordset)
   
   Dim dblVal As Double '加權計件值小計
   Dim dblValTot As Double '加權計件值總計
   
   iPage = 1
   dblVal = 0
   dblValTot = 0
   With p_Rst
      .MoveFirst
      PrintTitle
      '前所別
      strTemp3(1) = CheckStr(.Fields(1))
      strTemp3(3) = CheckStr(.Fields("ST06"))
      PrintSubTitle
       
      Do While .EOF = False
           For i = 0 To 14
               strTemp(i) = CheckStr(.Fields(i))
           Next i
           '案件名稱
           strTemp(4) = StrToStr(strTemp(4), 14)
           '專利種類
           strTemp(5) = StrToStr(strTemp(5), 4)
           '案件性質
           strTemp(6) = StrToStr(strTemp(6), 6)
           '客戶國籍
           strTemp(7) = StrToStr(strTemp(7), 4)
           '申請國家
           strTemp(8) = StrToStr(strTemp(8), 4)
           '點數
           strTemp(11) = Format(strTemp(11), "###.0")
           '特殊案件
           strTemp(12) = Format(strTemp(12), "###.0")
           '加乘註記
           strTemp(13) = Format(strTemp(13), "###.0")
           '計件值
           strTemp(14) = Format(strTemp(14), "###.0")
           
           '若所別不同
           If strTemp3(1) <> strTemp(1) Then
               PrintSubTot p_Rst, dblVal, strTemp3(3)
               dblValTot = dblValTot + dblVal
               
              '前所別
               strTemp3(1) = strTemp(1)
               strTemp3(3) = CheckStr(.Fields("ST06"))
               dblVal = 0
               PrintSubTitle
           End If
           NewLine False
           PrintDetail
           'dblVal = dblVal + Val(strTemp(13)) * Val(strTemp(14))
           dblVal = dblVal + Val("" & .Fields("XC1")) + Val("" & .Fields("SCR02"))
           .MoveNext
       Loop
       PrintSubTot p_Rst, dblVal, strTemp3(3)
       dblValTot = dblValTot + dblVal
   End With
   '最後一頁
   iPage = iPage + 1
   Printer.NewPage
   strTemp3(0) = ""
   strTemp3(1) = ""
   PrintTitle
   PrintSubTitle
   PrintSubTot p_Rst, dblValTot
   Printer.EndDoc
   
End Sub

Sub PrintTitle()         '最上面
   iPrint = 500
   Printer.Orientation = 2
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6500
   Printer.CurrentY = iPrint
   Printer.Print "內專  發文簿"
   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 6400
   Printer.CurrentY = iPrint
   Printer.Print "發文日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(iPage)
   iPrint = iPrint + 300
   
   Printer.FontSize = 10
End Sub

Sub PrintSubTitle()

   If iPrint >= 9000 Then
      iPage = iPage + 1
      Printer.NewPage
      PrintTitle
   End If
      
   Select Case m_iMode
      Case 1
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print "承辦人：" & strTemp3(0)
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print "所別：" & strTemp3(1)
         iPrint = iPrint + 300
      Case 3
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print "所別：" & strTemp3(1)
         iPrint = iPrint + 300
   End Select
   
   PLine
   
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "發文日"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "專利種類"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "申請人國籍"
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = iPrint
   Printer.Print "申請國家"
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = iPrint
   Printer.Print "收文日"
   Printer.CurrentX = PLeft(10)
   Printer.CurrentY = iPrint
   Printer.Print "算件數"
   Printer.CurrentX = PLeft(11)
   Printer.CurrentY = iPrint
   Printer.Print "點數"
   Printer.CurrentX = PLeft(12)
   Printer.CurrentY = iPrint
   Printer.Print "特殊案件"
   Printer.CurrentX = PLeft(13)
   Printer.CurrentY = iPrint
   Printer.Print "加乘註記"
   Printer.CurrentX = PLeft(14)
   Printer.CurrentY = iPrint
   Printer.Print "計件值"
   PLine True
   
End Sub
 
Sub GetPleft()
   Erase PLeft
   PLeft(0) = 500
   PLeft(1) = 3000
   
   PLeft(2) = 500
   PLeft(3) = PLeft(2) + 1000 '1600
   PLeft(4) = PLeft(3) + 1700 '3200
   PLeft(5) = PLeft(4) + 3000 '6300
   PLeft(6) = PLeft(5) + 1000 '7200
   PLeft(7) = PLeft(6) + 1400 '8600
   PLeft(8) = PLeft(7) + 1200 '9800
   PLeft(9) = PLeft(8) + 1000 '10800
   PLeft(10) = PLeft(9) + 1000 '11800
   PLeft(11) = PLeft(10) + 800 '12600
   PLeft(12) = PLeft(11) + 800 '13400
   PLeft(13) = PLeft(12) + 1000 '14400
   PLeft(14) = PLeft(13) + 1000 '15400
   PLeft(15) = PLeft(14) + 800 '16200
   
End Sub

Sub PrintDetail(Optional p_bC02 As Boolean = True)
   If p_bC02 = True Then
      Printer.CurrentX = PLeft(2)
      Printer.CurrentY = iPrint
      Printer.Print strTemp(2)
   End If
   For i = 3 To 10
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
   Next i
   For i = 11 To 14
      Printer.CurrentX = PLeft(i + 1) - 200 - Printer.TextWidth(strTemp(i))
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
   Next i
   iPrint = iPrint + 300
End Sub
 
Private Sub Form_Load()
   MoveFormToCenter Me
   txt1(0) = GetSystemKindByNick
   txt1(8) = 1
   GetPleft
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040310 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 8 '列印順序
         If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 Then
            KeyAscii = 0
         End If
      Case 9 '是否依承辦人跳頁
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
         End If
   End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Select Case Index
      Case 0
           strTemp1 = Split(UCase(GetSystemKindByNick), ",")
           strTemp2 = Split(UCase(txt1(0)), ",")
           For i = 0 To UBound(strTemp2)
              s = 0
              For j = 0 To UBound(strTemp1)
                  If strTemp1(j) = strTemp2(i) Then
                      s = 1
                      Exit For
                  End If
              Next j
              If s = 0 Then
                  s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
                  txt1(0).SetFocus
              End If
           Next i
      Case 2
         If blnClkSure = False Then
            If Not nickChgRan(txt1(1), txt1(2), "發文日") Then
                txt1(1).SetFocus
                TextInverse txt1(1)
                Exit Sub
            End If
         Else
            blnClkSure = False
         End If
      Case 4
         If blnClkSure = False Then
            If Not nickChgRan(txt1(3), txt1(4), "案件性質") Then
               txt1(3).SetFocus
               TextInverse txt1(3)
               Exit Sub
            End If
         Else
            blnClkSure = False
         End If
      Case 7
         If blnClkSure = False Then
            If Not nickChgRan(txt1(6), txt1(7), "申請國家") Then
               txt1(6).SetFocus
               TextInverse txt1(6)
               Exit Sub
            End If
         End If
      
      Case 5
         lbl1.Caption = GetPrjSalesNM(txt1(5))
         If Trim(lbl1.Caption) = "" And Trim(txt1(5)) <> "" Then
             MsgBox "承辦人代號錯誤，請重新輸入 !"
             txt1(5).SetFocus
             Exit Sub
         End If
      Case 8
         Select Case Val(txt1(8))
            Case 1, 2, 3
            Case Else
               s = MsgBox("列印別只能1 或 2 或 3 !!", , "USER 輸入錯誤")
               txt1(8).SetFocus
               Exit Sub
         End Select
      Case Else
   End Select

End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 1, 2 '發文日起, 迄
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Cancel = True
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
         End If
   End Select
End Sub

'PCT進入國家階段條件
Private Sub txtPA46_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtPA46.IMEMode = 2
   CloseIme
   TextInverse txtPA46
End Sub

Private Sub txtPA46_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And Chr(KeyAscii) <> "Y" Then
      KeyAscii = 0
      Beep
   End If
End Sub

'依照P台->P大->CFP->PS->CPS原則抓案件性質中文
Private Function GetProperty(ByVal p_CP10 As String) As String

On Error GoTo ErrHnd

   strSql = "SELECT NVL(NVL(NVL(MAX( DECODE(CPM01,'P', DECODE(CPM03,'（無）',CPM04,CPM03)))" & _
      ",MAX( DECODE(CPM01,'CFP', CPM04 ))),MAX( DECODE(CPM01,'PS', CPM04 ))),MAX( DECODE(CPM01,'CPS', CPM04 )))" & _
      " FROM CASEPROPERTYMAP WHERE CPM01 IN ('P','CFP','PS','CPS') AND CPM02='" & p_CP10 & "'"
      
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If Not (.EOF And .BOF) Then
         GetProperty = "" & .Fields(0)
      End If
   End With
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

'印橫虛線
Private Sub PLine(Optional ByVal p_bSkip As Boolean = False)
   If p_bSkip = True Then
      iPrint = iPrint + 300
   End If
   If iPrint >= 10000 Then
      iPage = iPage + 1
      Printer.NewPage
      PrintTitle
      PrintSubTitle
   Else
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      Printer.Print String(200, "-")
      iPrint = iPrint + 300
   End If
End Sub

'跳行
Private Sub NewLine(Optional ByVal p_bSkip As Boolean = True)
   If p_bSkip = True Then
      iPrint = iPrint + 300
   End If
   If iPrint >= 10000 Then
      iPage = iPage + 1
      Printer.NewPage
      PrintTitle
      PrintSubTitle
   End If
End Sub

'印小計
Private Sub PrintSubTot(ByRef p_Rst As ADODB.Recordset, ByVal p_Value As Double, Optional ByRef p_Filter As String = Empty, Optional ByVal p_Sh As Double)
   
   Dim adoRst As ADODB.Recordset, i As Integer, stPreCP10 As String
   Dim iCnt As Integer, iCP10Cnt As Integer, dblPoint As Double, pTemp As String
   
   Set adoRst = p_Rst.Clone
   With adoRst
      '計件
      If p_Filter = Empty Then
         .Filter = "CP26='Y'"
      Else
         PLine
         Select Case m_iMode
            Case 1
               .Filter = "CP14='" & p_Filter & "' and CP26='Y'"
            Case 2
               .Filter = "CP27='" & p_Filter & "' and CP26='Y'"
            Case 3
               .Filter = "ST06='" & p_Filter & "' and CP26='Y'"
         End Select
      End If
      .Sort = "CP10"
      
      '案件性質種類數
      i = 0
      '件數合計
      iCnt = 0
      '點數合計
      dblPoint = 0
      '案件性質件數統計值
      iCP10Cnt = 0
      If .RecordCount > 0 Then
         .MoveFirst
         '案件性質
         stPreCP10 = "" & .Fields("CP10")
         Do While Not .EOF
            If stPreCP10 <> "" & .Fields("CP10") Then
               If i > 4 Then
                  NewLine
                  i = 0
               End If
               '印案件性質
               pTemp = StrToStr(GetProperty(stPreCP10), 6)
               Printer.CurrentX = 500 + (i * 2500)
               Printer.CurrentY = iPrint
               Printer.Print pTemp
               '印案件性質件數統計值
               pTemp = Format(iCP10Cnt)
               Printer.CurrentX = 2500 + (i * 2500) - Printer.TextWidth(pTemp)
               Printer.CurrentY = iPrint
               Printer.Print pTemp
               i = i + 1
               stPreCP10 = "" & .Fields("CP10")
               iCP10Cnt = 0
            End If
            iCnt = iCnt + 1
            iCP10Cnt = iCP10Cnt + 1
            dblPoint = dblPoint + Val("" & .Fields("CP18"))
            .MoveNext
         Loop
         '印最後一個案件性質資料
         If i > 4 Then
            NewLine
            i = 0
         End If
         '印案件性質
         pTemp = StrToStr(GetProperty(stPreCP10), 6)
         Printer.CurrentX = 500 + (i * 2500)
         Printer.CurrentY = iPrint
         Printer.Print pTemp
         '印案件性質件數統計值
         pTemp = Format(iCP10Cnt)
         Printer.CurrentX = 2500 + (i * 2500) - Printer.TextWidth(pTemp)
         Printer.CurrentY = iPrint
         Printer.Print pTemp
         NewLine
      End If
      
      pTemp = "計件合計： 案件數 " & iCnt & " 件"
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      Printer.Print pTemp
      
      pTemp = "點數 " & Right(Space(7) & Format(dblPoint, "###.0"), 7)
      Printer.CurrentX = PLeft(4) + 800
      Printer.CurrentY = iPrint
      Printer.Print pTemp
      
      '印計件值合計值
      If p_Sh > 0 Then
         pTemp = "支援時數：" & Right(Space(7) & Format(p_Sh, "###.00"), 7)
         Printer.CurrentX = PLeft(13) - 200 - Printer.TextWidth(pTemp)
         Printer.CurrentY = iPrint
         Printer.Print pTemp
      End If
      '印計件值合計值
      pTemp = "合計：" & Right(Space(7) & Format(p_Value + p_Sh, "###.00"), 7)
      Printer.CurrentX = PLeft(15) - 200 - Printer.TextWidth(pTemp)
      Printer.CurrentY = iPrint
      Printer.Print pTemp
      
            
      '不計件
      PLine True
      
      If p_Filter = Empty Then
         .Filter = "CP26='N'"
      Else
         Select Case m_iMode
            Case 1
               .Filter = "CP14='" & p_Filter & "' and CP26='N'"
            Case 2
               .Filter = "CP27='" & p_Filter & "' and CP26='N'"
            Case 3
               .Filter = "ST06='" & p_Filter & "' and CP26='N'"
         End Select
      End If
      .Sort = "CP10"
      
      '案件性質種類數
      i = 0
      '件數合計
      iCnt = 0
      '點數合計
      dblPoint = 0
      '案件性質件數統計值
      iCP10Cnt = 0
      If .RecordCount > 0 Then
         .MoveFirst
         '案件性質
         stPreCP10 = "" & .Fields("CP10")
         Do While Not .EOF
            If stPreCP10 <> "" & .Fields("CP10") Then
               If i > 4 Then
                  NewLine
                  i = 0
               End If
               '印案件性質
               pTemp = StrToStr(GetProperty(stPreCP10), 6)
               Printer.CurrentX = 500 + (i * 2500)
               Printer.CurrentY = iPrint
               Printer.Print pTemp
               '印案件性質件數統計值
               pTemp = Format(iCP10Cnt)
               Printer.CurrentX = 2500 + (i * 2500) - Printer.TextWidth(pTemp)
               Printer.CurrentY = iPrint
               Printer.Print pTemp
               i = i + 1
               stPreCP10 = "" & .Fields("CP10")
               iCP10Cnt = 0
            End If
            iCP10Cnt = iCP10Cnt + 1
            iCnt = iCnt + 1
            dblPoint = dblPoint + Val("" & .Fields("CP18"))
            .MoveNext
         Loop
         '印最後一個案件性質資料
         If i > 4 Then
            NewLine
            i = 0
         End If
         '印案件性質
         pTemp = StrToStr(GetProperty(stPreCP10), 6)
         Printer.CurrentX = 500 + (i * 2500)
         Printer.CurrentY = iPrint
         Printer.Print pTemp
         '印案件性質件數統計值
         pTemp = Format(iCP10Cnt)
         Printer.CurrentX = 2500 + (i * 2500) - Printer.TextWidth(pTemp)
         Printer.CurrentY = iPrint
         Printer.Print pTemp
         NewLine
      End If
      
      pTemp = "不計件合計： 案件數 " & iCnt & " 件"
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      Printer.Print pTemp
      
      pTemp = "點數 " & Right(Space(7) & Format(dblPoint, "###.0"), 7)
      Printer.CurrentX = PLeft(4) + 800
      Printer.CurrentY = iPrint
      Printer.Print pTemp
      
      iPrint = iPrint + 600
   End With
   Set adoRst = Nothing
End Sub
