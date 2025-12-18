VERSION 5.00
Begin VB.Form frm020402 
   BorderStyle     =   1  '單線固定
   Caption         =   "商申案智權人員准駁統計表"
   ClientHeight    =   2832
   ClientLeft      =   4476
   ClientTop       =   1740
   ClientWidth     =   3876
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2832
   ScaleWidth      =   3876
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   12
      TabIndex        =   11
      Top             =   1680
      Width           =   3825
      Begin VB.ComboBox Combo1 
         Height          =   276
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   5
         Top             =   180
         Width           =   2880
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   180
         Left            =   105
         TabIndex        =   12
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   972
      TabIndex        =   0
      Top             =   705
      Width           =   1920
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   972
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1365
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1968
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1365
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   972
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1035
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   4
      Left            =   1980
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1035
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2328
      TabIndex        =   6
      Top             =   24
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3120
      TabIndex        =   7
      Top             =   24
      Width           =   756
   End
   Begin VB.Label Label3 
      Caption         =   "註：以A4報表列印"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   180
      TabIndex        =   13
      Top             =   2520
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   10
      Top             =   720
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "公告日："
      Height          =   180
      Index           =   2
      Left            =   90
      TabIndex        =   9
      Top             =   1410
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   90
      TabIndex        =   8
      Top             =   1065
      Width           =   915
   End
   Begin VB.Line Line1 
      X1              =   1410
      X2              =   2670
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   1650
      X2              =   2400
      Y1              =   1140
      Y2              =   1140
   End
End
Attribute VB_Name = "frm020402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, SavDay5 As String, SavDay6 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 5) As String, strTemp3 As String, StrTemp4 As String
Dim PLeft(0 To 4) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, StrTemp7(0 To 5) As String, SeekPrint As Integer, SeekPrintL As Integer
Dim BolEndThisPage As Boolean
'Add By Cheng 2002/05/07
Dim m_strSaleZone As String '業務區

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     PUB_RestorePrinter Combo1 'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除
     Printer.EndDoc 'Add By Sindy 2011/11/1
      'Add By Sindy 2013/8/14
      If Left(Trim(Pub_StrUserSt03), 1) = "F" Or Pub_StrUserSt03 = "M51" Then
         Printer.PaperSize = 9
      Else
      '2013/8/14 END
         'Modified by Moran 2015/6/1
         'Printer.PaperSize = 39
         Printer.PaperSize = PUB_GetPaperSize(15, 2)
         'end 2015/6/1
      End If
     DoEvents

     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         'Add By Cheng 2002/03/21
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
         
         If Len(txt1(2)) = 0 Then
             s = MsgBox("公告日區間不可空白!!", , "USER 輸入錯誤")
             txt1(1).SetFocus
             txt1_GotFocus (1)
             Exit Sub
         Else
             Screen.MousePointer = vbHourglass
             Me.Enabled = False
             ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/19 清除查詢印表記錄檔欄位
             Process
             Me.Enabled = True
             Screen.MousePointer = vbDefault
         End If
     End If
Case 1
    'Add By Cheng 2004/04/30
    '若印表機變動, 則更新列印設定
    If Me.Combo1.Text <> Me.Combo1.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
    End If
     Unload Me
Case Else
End Select
End Sub

Sub Process()
'Add By Cheng 2002/02/05
Dim rsTmp1 As New ADODB.Recordset
Dim strSales As String, strDept As String 'Add By Sindy 2013/8/15

cnnConnection.Execute "DELETE FROM R020402 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
StrSQL6 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND TM01 IN (" & SQLGrpStr(txt1(0), 2) & ")"
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/10/19
End If
'Add By Cheng 2003/07/11
'若非商申收發文時, 若為TF案則不抓後三碼為"000"的資料
strSQL1 = strSQL1 + " AND TM03 <> Decode(CP01,'TF','0','z') AND TM04 <> Decode(CP01,'TF','00','zz') "
StrSQL6 = ""
'Modify By Cheng 2003/12/03
'若申請國家非大陸
If Me.txt1(3).Text <> "020" Then
    If Len(txt1(1)) <> 0 Then
       StrSQL6 = StrSQL6 + " AND TM14>=" & Val(ChangeTStringToWString(txt1(1))) & ""
    End If
    If Len(Trim(txt1(2))) <> 0 Then
       StrSQL6 = StrSQL6 + " AND TM14<=" & Val(ChangeTStringToWString(txt1(2))) & " "
    End If
'若申請國家為大陸
Else
    If Len(txt1(1)) <> 0 Then
       StrSQL6 = StrSQL6 + " AND C3.CP25>=" & Val(ChangeTStringToWString(txt1(1))) & ""
    End If
    If Len(Trim(txt1(2))) <> 0 Then
       StrSQL6 = StrSQL6 + " AND C3.CP25<=" & Val(ChangeTStringToWString(txt1(2))) & " "
    End If
End If
If Len(txt1(1)) <> 0 Or Len(Trim(txt1(2))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/10/19
End If
If Len(txt1(3)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND TM10>='" & txt1(3) & "' "
End If
If Len(txt1(4)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND TM10<='" & txt1(4) & "' "
End If
If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/10/19
End If
'Modify By Cheng 2002/05/07
''Modify By Cheng 2002/02/05
''StrSQL6 = StrSQL6 + " AND (CP10='1001' OR CP10='1002' OR CP10='1403') "
'StrSQL6 = StrSQL6 + " AND (C1.CP10='1001' OR C1.CP10='1002' OR C1.CP10='1403') "
'93.6.14 CANCEL BY SONIA 僅收/發文統計表才控制不算案件數
''Add By Cheng 2004/04/29
''抓計件的資料
'StrSQL6 = StrSQL6 & " And CP26 Is Null "
''End
CheckOC
'StrSQL = "SELECT NVL(A0902,A0903),ST02,DECODE(TM16,'1',1,'2',0,0),DECODE(TM16,'1',0,'2',1,1),0,ST03,TM14,MAX(CP05) FROM TRADEMARK,CASEPROGRESS,ACC090,STAFF WHERE CP01=TM01 AND CP02=TM02 AND CP03=TM03 AND CP04=TM04 AND CP13=ST01(+) AND ST03=A0901(+) " & StrSQL1 & StrSQL6
'Modify By Cheng 2002/02/05
'注意：業務區是CASEPROGRESS的CP12
'strSQL = "SELECT NVL(A0902,A0903),cp13,DECODE(TM16,'1',1,'2',0,0),DECODE(TM16,'1',0,'2',1,1),0,ST03,'" & strUserNum & "' " & _
         " FROM TRADEMARK,(SELECT CP01,CP02,CP03,CP04,CP10,MAX(CP05) AS AA,CP13 FROM CASEPROGRESS GROUP BY CP01,CP02,CP03,CP04,CP10,CP13) NEW1,ACC090,STAFF " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=ST01(+) AND AA<=TM14 AND ST03=A0901(+) " & strSQL1 & StrSQL6
'Modify By Cheng 2002/05/07
''先設定相關總收文號
'strSQL = " SELECT C2.CP43 " & _
'         " FROM TRADEMARK,(SELECT CP01,CP02,CP03,CP04,CP10,MAX(CP05) AS AA,CP13 FROM CASEPROGRESS GROUP BY CP01,CP02,CP03,CP04,CP10,CP13) C1,(SELECT CP01,CP02,CP03,CP04,CP10,MAX(CP05) AS AA,CP13,CP43 FROM CASEPROGRESS GROUP BY CP01,CP02,CP03,CP04,CP10,CP13,CP43) C2 " & _
'         " WHERE C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C1.AA<=TM14 " & _
'         " AND C1.CP01=C2.CP01(+) AND C1.CP02=C2.CP02(+) AND C1.CP03=C2.CP03(+) AND C1.CP04=C2.CP04(+) AND C1.CP10=C2.CP10(+) AND C1.AA=C2.AA(+) AND C1.CP13=C2.CP13(+) " & strSQL1 & StrSQL6
'strSQL = "SELECT NVL(A0902,A0903),C3.CP13,DECODE(TM16,'1',1,'2',0,0),DECODE(TM16,'1',0,'2',1,1),0,C3.CP12,'" & strUserNum & "' " & _
'         " FROM TRADEMARK,CASEPROGRESS C3,ACC090,STAFF " & _
'         " WHERE C3.CP01=TM01(+) AND C3.CP02=TM02(+) AND C3.CP03=TM03(+) AND C3.CP04=TM04(+) AND C3.CP12=A0901(+) AND C3.CP14=ST01(+) AND C3.CP09 IN ( " & strSQL & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR C3.CP14 IS NULL ) ", " AND (SUBSTR(ST03,1,2)='F1' OR C3.CP14 IS NULL ) ")
strSql = " Select Min(CP09) From CaseProgress C1 Where C3.CP01=C1.CP01 AND C3.CP02=C1.CP02 AND C3.CP03=C1.CP03 AND C3.CP04=C1.CP04 AND C1.CP10<>'001' AND C1.CP09<'B' "
'910521 nick
'strSQL = " SELECT NVL(A0902,A0903),C3.CP13,DECODE(TM16,'1',1,'2',0,0),DECODE(TM16,'1',0,'2',1,1),0,C3.CP12,'" & strUserNum & "' " & _
         " FROM TRADEMARK,CASEPROGRESS C3,ACC090,STAFF " & _
         " WHERE TM01=C3.CP01(+) AND TM02=C3.CP02(+) AND TM03=C3.CP03(+) AND TM04=C3.CP04(+) AND C3.CP12=A0901(+) AND C3.CP14=ST01(+) AND C3.CP09 IN ( " & strSQL & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR C3.CP14 IS NULL ) ", " AND (SUBSTR(ST03,1,2)='F1' OR C3.CP14 IS NULL ) ") & strSQL1 & StrSQL6
'92.3.6 MODIFY BY SONIA 若從外商的統計報表進入, 不管承辦人部門別
'strSQL = " SELECT c3.cp12,C3.CP13,DECODE(TM16,'1',1,'2',0,0),DECODE(TM16,'1',0,'2',1,1),0,C3.CP12,'" & strUserNum & "' " & _
'         " FROM TRADEMARK,CASEPROGRESS C3,ACC090,STAFF " & _
'         " WHERE TM01=C3.CP01(+) AND TM02=C3.CP02(+) AND TM03=C3.CP03(+) AND TM04=C3.CP04(+) AND C3.CP12=A0901(+) AND C3.CP14=ST01(+) AND C3.CP09 IN ( " & strSQL & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR C3.CP14 IS NULL ) ", " AND (SUBSTR(ST03,1,2)='F1' OR C3.CP14 IS NULL ) ") & strSQL1 & StrSQL6
'Add By Sindy 2013/8/14 +,C3.CP01,C3.CP02,C3.CP03,C3.CP04
strSql = " SELECT c3.cp12,C3.CP13,DECODE(TM16,'1',1,'2',0,0),DECODE(TM16,'1',0,'2',1,1),0,C3.CP12,'" & strUserNum & "',C3.CP01,C3.CP02,C3.CP03,C3.CP04" & _
         " FROM TRADEMARK,CASEPROGRESS C3,ACC090,STAFF " & _
         " WHERE TM01=C3.CP01(+) AND TM02=C3.CP02(+) AND TM03=C3.CP03(+) AND TM04=C3.CP04(+) AND C3.CP12=A0901(+) AND C3.CP14=ST01(+) AND C3.CP09 IN ( " & strSql & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR C3.CP14 IS NULL ) ", "") & strSQL1 & StrSQL6
'92.3.6 END

cnnConnection.Execute "INSERT INTO R020402 " & strSql

'Fields--0:智權人員部門名稱, 1:智權人員代號 , 2:目前准/駁, 3:目前准/駁, 4:???, 5:智權人員部門代號, 6:使用者代號, 7,8,9,10.本所案號
strSql = "SELECT * FROM R020402 WHERE ID='" & strUserNum & "' "
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/19
        'Add By Sindy 2013/8/14 FCT案件之智權人員改以 PUB_GetFCTSalesNo 抓最新智權人員
        .MoveFirst
        Do While Not .EOF
         If .Fields("R069007") = "FCT" Then
            strSales = PUB_GetFCTSalesNo(.Fields("R069007"), .Fields("R069008"), .Fields("R069009"), .Fields("R069010"))
            strDept = GetSalesArea(strSales)
            strSql = "update R020402 set R069001='" & strDept & "',R069002='" & strSales & "',R069006='" & strDept & "'" & _
                     " where ID='" & strUserNum & "'" & _
                       " and R069007='" & .Fields("R069007") & "' and R069008='" & .Fields("R069008") & "'" & _
                       " and R069009='" & .Fields("R069009") & "' and R069010='" & .Fields("R069010") & "'"
            cnnConnection.Execute strSql
         End If
         .MoveNext
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/19
        ShowNoData
        Screen.MousePointer = vbDefault
        CheckOC
        Exit Sub
    End If
End With
CheckOC
PrintData
Screen.MousePointer = vbDefault
End Sub

Sub PrintData()
BolEndThisPage = False
strSql = "SELECT NVL(A0902,A0903),nvl(st02,R069002),SUM(R069003),SUM(R069004),SUM(R069005),R069006,r069001,r069002 FROM R020402,staff,acc090 WHERE ID='" & strUserNum & "' and r069001=a0901(+) and r069002=st01(+) GROUP BY R069006,R069001,R069002,nvl(st02,R069002),NVL(A0902,A0903) ORDER BY R069006,R069001,R069002 "
CheckOC
Page = 1
'Add By Cheng 2002/05/07
m_strSaleZone = ""
strTemp3 = ""
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        PrintTitle
        strTemp3 = CheckStr(.Fields(0))
        StrTemp4 = CheckStr(.Fields(5))
        SavDay5 = CheckStr(.Fields(6))
        Do While .EOF = False
            For i = 0 To 5
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            
            If strTemp3 <> strTemp(0) Then
                'Add By Cheng 2002/05/07
                m_strSaleZone = ""
                ShowLine
                PrintEnd (0)
                If StrToStr(StrTemp4, 1) <> StrToStr(strTemp(5), 1) Then
                    ShowLine
                    PrintEnd (1)
                    
                    StrTemp4 = strTemp(5)
                End If
                strTemp3 = strTemp(0)
                SavDay5 = CheckStr(.Fields(6))
            End If
            If Val(strTemp(3)) = 0 Then
                If Val(strTemp(2)) = 0 Then
                    strTemp(4) = "0"
                Else
                    strTemp(4) = "100"
                End If
            Else
                strTemp(4) = Trim(str(Val(strTemp(2)) / (Val(strTemp(2)) + Val(strTemp(3))) * 100))
            End If
            PrintDatil
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    Else
         'Add By Sindy 2011/3/1
         ShowNoData
         Exit Sub
         '2011/3/1 End
    End If
End With
CheckOC
ShowLine
PrintEnd (0)
ShowLine
PrintEnd (1)
ShowLine
PrintEnd (2)
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintEnd(Strindex As Integer)
Select Case Strindex
Case 0
     strSql = "select '各區小計','',SUM(R069003),SUM(R069004),SUM(R069005),'' FROM R020402 WHERE ID='" & strUserNum & "' AND R069001='" & SavDay5 & "' GROUP BY R069001 "
Case 1
     strSql = "select '各所小計','',sum(r069003),sum(r069004),sum(r069005),'' FROM R020402 WHERE ID='" & strUserNum & "' AND SUBSTR(R069001,1,2)='" & Mid(SavDay5, 1, 2) & "' GROUP BY SUBSTR(R069001,1,2)  "
Case 2
     strSql = "select '全所總計','',sum(r069003),sum(r069004),sum(r069005),'' FROM R020402 WHERE ID='" & strUserNum & "' "
     BolEndThisPage = True
Case Else
     Exit Sub
End Select
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 5
                StrTemp7(i) = CheckStr(.Fields(i))
            Next i
            If Val(StrTemp7(3)) = 0 Then
                If Val(StrTemp7(2)) = 0 Then
                    StrTemp7(4) = "0"
                Else
                    StrTemp7(4) = "100"
                End If
            Else
                StrTemp7(4) = Trim(str(Val(StrTemp7(2)) / (Val(StrTemp7(2)) + Val(StrTemp7(3))) * 100))
            End If
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            For i = 2 To 3
                Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(StrTemp7(i))
                Printer.CurrentY = iPrint
                Printer.Print StrTemp7(i)
            Next i
            Printer.CurrentX = PLeft(4) + 500 - Printer.TextWidth(Format(StrTemp7(4), "###.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(4), "###.00") & "%"
            iPrint = iPrint + 300
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
iPrint = iPrint + 300
End Sub

Sub PrintTitle()
Dim strText As String
GetPleft
iPrint = 0
'Printer.Orientation = 1 'Removed by Morgan 2015/6/3
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
'Add By Sindy 2013/8/14
If Left(Trim(Pub_StrUserSt03), 1) = "F" Or Pub_StrUserSt03 = "M51" Then
   Printer.CurrentX = 3400
Else
'2013/8/14 END
   Printer.CurrentX = 5500
End If
'Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("商申智權人員准駁統計表") / 2)
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "商申智權人員准駁統計表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
strText = IIf(Me.txt1(3).Text = "020", "　准駁日：", "　公告日：") & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
'Add By Sindy 2013/8/14
If Left(Trim(Pub_StrUserSt03), 1) = "F" Or Pub_StrUserSt03 = "M51" Then
   Printer.CurrentX = 4200
Else
'2013/8/14 END
   Printer.CurrentX = 7000
End If
Printer.CurrentY = iPrint
'Modify By Cheng 2003/12/03
'Printer.Print "公告日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
Printer.Print strText
'End
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
'Add By Sindy 2013/8/14
If Left(Trim(Pub_StrUserSt03), 1) = "F" Or Pub_StrUserSt03 = "M51" Then
   Printer.CurrentX = 9000
Else
'2013/8/14 END
   Printer.CurrentX = 16500
End If
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300

'Add By Cheng 2002/02/05
Printer.CurrentX = 250
Printer.CurrentY = iPrint
Printer.Print "申請國家：" & Me.txt1(3).Text & " － " & Me.txt1(4).Text
'Add By Sindy 2013/8/14
If Left(Trim(Pub_StrUserSt03), 1) = "F" Or Pub_StrUserSt03 = "M51" Then
   Printer.CurrentX = 4200
Else
'2013/8/14 END
   Printer.CurrentX = 6750
End If
Printer.CurrentY = iPrint
Printer.Print "系統類別：" & Me.txt1(0).Text

'Add By Sindy 2013/8/14
If Left(Trim(Pub_StrUserSt03), 1) = "F" Or Pub_StrUserSt03 = "M51" Then
   Printer.CurrentX = 9000
Else
'2013/8/14 END
   Printer.CurrentX = 16500
End If
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "業務區別"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "核准"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "核駁"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "核准率"
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
End Sub

Sub PrintDatil()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
'Modify By Cheng 2002/05/07
If m_strSaleZone <> strTemp(0) Then
   Printer.Print strTemp(0)
   m_strSaleZone = strTemp(0)
Else
   Printer.Print ""
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
For i = 2 To 3
    Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(strTemp(i))
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
Printer.CurrentX = PLeft(4) + 500 - Printer.TextWidth(Format(strTemp(4), "###.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(4), "###.00") + "%"
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 2500
PLeft(2) = 4500
PLeft(3) = 6500
PLeft(4) = 8500
End Sub

Private Sub Form_Load()

MoveFormToCenter Me
txt1(0) = GetSystemKindByNickTnoS

SeekPrintL = Printer.Orientation
PUB_SetPrinter Me.Name, Combo1, , , SeekPrint     'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除

   'Add By Sindy 2013/8/14
   Label3.Visible = False
   If Left(Trim(Pub_StrUserSt03), 1) = "F" Or Pub_StrUserSt03 = "M51" Then
      Label3.Visible = True
   End If
   '2013/8/14 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Printer = Printers(SeekPrint)
Printer.Orientation = SeekPrintL
Set frm020402 = Nothing
End Sub

Private Sub txt1_Change(Index As Integer)
    'Add By Cheng 2003/12/03
    Select Case Index
    Case 3
        Me.txt1(4).Text = Me.txt1(3).Text
        If Me.txt1(3).Text = "020" Then
            Me.Label1(2).Caption = "准駁日："
        Else
            Me.Label1(2).Caption = "公告日："
        End If
    End Select
    'End
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdok(0).SetFocus
End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
     strTemp1 = Split(Replace(UCase(GetSystemKindByNickTnoS), ",,", ""), ",")
     strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp2(i) = strTemp1(j) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限!! ", , "USER 權限問題")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
     Next i
Case 1, 2
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 2 Then
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
    End If
Case 4
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
Case Else
End Select
End Sub

Sub ShowLine()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
   If BolEndThisPage = False Then
      Page = Page + 1
      Printer.NewPage
      PrintTitle
   Else
      BolEndThisPage = False
   End If
End If
End Sub

