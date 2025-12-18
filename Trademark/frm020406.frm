VERSION 5.00
Begin VB.Form frm020406 
   BorderStyle     =   1  '單線固定
   Caption         =   "商申案承辦人准駁統計表"
   ClientHeight    =   2460
   ClientLeft      =   4020
   ClientTop       =   2745
   ClientWidth     =   3915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3915
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   615
      Left            =   36
      TabIndex        =   11
      Top             =   1635
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
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3150
      TabIndex        =   7
      Top             =   36
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2352
      TabIndex        =   6
      Top             =   36
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   4
      Left            =   2505
      MaxLength       =   4
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1485
      MaxLength       =   4
      TabIndex        =   1
      Top             =   960
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2505
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1305
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1485
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1305
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1485
      TabIndex        =   0
      Top             =   600
      Width           =   1956
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   1995
      X2              =   2745
      Y1              =   1065
      Y2              =   1065
   End
   Begin VB.Line Line1 
      X1              =   1995
      X2              =   3255
      Y1              =   1455
      Y2              =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   390
      TabIndex        =   10
      Top             =   975
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "註冊公告日："
      Height          =   180
      Index           =   2
      Left            =   390
      TabIndex        =   9
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   390
      TabIndex        =   8
      Top             =   660
      Width           =   915
   End
End
Attribute VB_Name = "frm020406"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 9) As String, strTemp3 As String, StrTemp4 As String
Dim PLeft(0 To 7) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, StrTemp7(0 To 7) As String, SeekPrint As Integer, SeekPrintL As Integer
Dim BolEndThisPage As Boolean
'Add By Cheng 2002/05/07
Dim m_strPromoter As String '承辦人

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
     PUB_RestorePrinter Combo1 'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除
     Printer.EndDoc 'Add By Sindy 2011/11/1
     'Modified by Moran 2015/6/1
     'Printer.PaperSize = 39
     Printer.PaperSize = PUB_GetPaperSize(15, 2)
     'end 2015/6/1
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
             'Modify By Sindy 2011/11/1
             's = MsgBox("公告日區間不可空白!!", , "USER 輸入錯誤")
             s = MsgBox(Left(Label1(2), Len(Label1(2)) - 1) & "區間不可空白!!", , "USER 輸入錯誤")
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
'Add By Cheng 2002/02/06
Dim rsTmp1 As New ADODB.Recordset

Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM r020406 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
StrSQL6 = ""
If Len(txt1(0)) <> 0 Then
   'Modify By Cheng 2002/05/07
'   'Modify By Cheng 2002/02/06
''   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
'   strSQL1 = strSQL1 + " AND C1.CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   strSQL1 = strSQL1 + " AND TM01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/10/19
End If
'Add By Cheng 2003/07/11
'若非商申收發文時, 若為TF案則不抓後三碼為"000"的資料
strSQL1 = strSQL1 + " AND TM03 <> Decode(CP01,'TF','0','z') AND TM04 <> Decode(CP01,'TF','00','zz') "
StrSQL6 = ""
'Modify By Cheng 2003/12/02
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
If Len(txt1(3)) <> 0 Or Len(Trim(txt1(4))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/10/19
End If
'Modify By Cheng 2002/05/07
''Modify By Cheng 2002/02/05
''StrSQL6 = StrSQL6 + " AND (CP10='1001' OR CP10='1002' OR CP10='1403') "
'StrSQL6 = StrSQL6 + " AND (C1.CP10='1001' OR C1.CP10='1002' OR C1.CP10='1403') "
''''''StrSQL = "SELECT NVL(A0902,A0903),ST02,DECODE(TM16,'1',1,'2',0,0),DECODE(TM16,'1',0,'2',1,1),0,ST03,TM14,MAX(CP05) FROM TRADEMARK,CASEPROGRESS,ACC090,STAFF WHERE CP01=TM01 AND CP02=TM02 AND CP03=TM03 AND CP04=TM04 AND CP13=ST01(+) AND ST03=A0901(+) " & StrSQL1 & StrSQL6
'93.6.14 CANCEL BY SONIA 僅收/發文統計表才控制不算案件數
''Add By Cheng 2004/04/29
''抓計件的資料
'StrSQL6 = StrSQL6 & " And CP26 Is Null "
''End
'Modify By Cheng 2002/05/07
''Modify By Cheng 2002/02/05
''注意：業務區是CASEPROGRESS的CP12
''strSQL = "SELECT S2.ST02,NVL(A0902,A0903),DECODE(TM16,'1',1,'2',0,0),DECODE(TM16,'1',0,'2',1,1),0,S2.ST01,'" & strUserNum & "' " & _
''         " FROM TRADEMARK,(SELECT CP01,CP02,CP03,CP04,CP10,MAX(CP05) AS AA,CP13,cp14 FROM CASEPROGRESS GROUP BY CP01,CP02,CP03,CP04,CP10,CP13,CP14) NEW1,ACC090,STAFF S1,STAFF S2 " & _
''         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND AA<=TM14 AND S1.ST03=A0901(+) AND CP14=S2.ST01(+) " & strSQL1 & StrSQL6
''先設定相關總收文號
'strSQL = " SELECT C2.CP43 " & _
'         " FROM TRADEMARK,(SELECT CP01,CP02,CP03,CP04,CP10,MAX(CP05) AS AA,CP13,cp14 FROM CASEPROGRESS GROUP BY CP01,CP02,CP03,CP04,CP10,CP13,CP14) C1,(SELECT CP01,CP02,CP03,CP04,CP10,MAX(CP05) AS AA,CP13,cp14,CP43 FROM CASEPROGRESS GROUP BY CP01,CP02,CP03,CP04,CP10,CP13,CP14,CP43) C2,ACC090 " & _
'         " WHERE C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C1.AA<=TM14 " & _
'         " AND C1.CP01=C2.CP01(+) AND C1.CP02=C2.CP02(+) AND C1.CP03=C2.CP03(+) AND C1.CP04=C2.CP04(+) AND C1.CP10=C2.CP10(+) AND C1.AA=C2.AA(+) AND C1.CP13=C2.CP13(+) AND C1.CP14=C2.CP14(+) " & strSQL1 & StrSQL6
''90/03/12 ******  nick
''strSQL = "SELECT ST02,NVL(A0902,A0903),DECODE(TM16,'1',1,'2',0,0),DECODE(TM16,'1',0,'2',1,1),0,ST01,'" & strUserNum & "' " & _
'         " FROM TRADEMARK,CASEPROGRESS C3,ACC090,STAFF " & _
'         " WHERE C3.CP01=TM01(+) AND C3.CP02=TM02(+) AND C3.CP03=TM03(+) AND C3.CP04=TM04(+) AND C3.CP12=A0901(+) AND C3.CP14=ST01(+) AND C3.CP09 IN ( " & strSQL & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR C3.CP14 IS NULL ) ", " AND (SUBSTR(ST03,1,2)='F1' OR C3.CP14 IS NULL ) ")
'strSQL = "SELECT c3.cp14,NVL(A0902,A0903),DECODE(TM16,'1',1,'2',0,0),DECODE(TM16,'1',0,'2',1,1),0,ST01,'" & strUserNum & "' " & _
'         " FROM TRADEMARK,CASEPROGRESS C3,ACC090,STAFF " & _
'         " WHERE C3.CP01=TM01(+) AND C3.CP02=TM02(+) AND C3.CP03=TM03(+) AND C3.CP04=TM04(+) AND C3.CP12=A0901(+) AND C3.CP14=ST01(+) AND C3.CP09 IN ( " & strSQL & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR C3.CP14 IS NULL ) ", " AND (SUBSTR(ST03,1,2)='F1' OR C3.CP14 IS NULL ) ")
strSql = " SELECT MIN(C1.CP09) FROM CASEPROGRESS C1 WHERE C3.CP01=C1.CP01 AND C3.CP02=C1.CP02 AND C3.CP03=C1.CP03 AND C3.CP04=C1.CP04 AND C1.CP10<>'001' AND C1.CP09<'B' "
'strSQL = " SELECT c3.cp14,NVL(A0902,A0903),DECODE(TM16,'1',1,'2',0,0),DECODE(TM16,'1',0,'2',1,1),0,ST01,'" & strUserNum & "' " & _
         " FROM TRADEMARK,CASEPROGRESS C3,ACC090,STAFF " & _
         " WHERE TM01=C3.CP01(+) AND TM02=C3.CP02(+) AND TM03=C3.CP03(+) AND TM04=C3.CP04(+) AND C3.CP12=A0901(+) AND C3.CP14=ST01(+) AND C3.CP09 IN ( " & strSQL & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR C3.CP14 IS NULL ) ", " AND (SUBSTR(ST03,1,2)='F1' OR C3.CP14 IS NULL ) ") & strSQL1 & StrSQL6
'**************** 將業務區改成抓案件進度檔   91.08.15  nick
'Modify By Sindy 2011/11/1 +,0,TM01,TM02,TM03,TM04
strSql = " SELECT c3.cp14,c3.cp12,DECODE(TM16,'1',1,'2',0,0),DECODE(TM16,'1',0,'2',1,1),0,ST01,'" & strUserNum & "',0,TM01,TM02,TM03,TM04 " & _
         " FROM TRADEMARK,CASEPROGRESS C3,ACC090,STAFF " & _
         " WHERE TM01=C3.CP01 AND TM02=C3.CP02 AND TM03=C3.CP03 AND TM04=C3.CP04 AND C3.CP12=A0901(+) AND C3.CP14=ST01(+) AND C3.CP09 IN ( " & strSql & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR C3.CP14 IS NULL ) ", " AND (SUBSTR(ST03,1,2)='F1' OR C3.CP14 IS NULL ) ") & strSQL1 & StrSQL6

'*********************
'Fields--0:承辦人姓名, 1:智權人員所屬部門名稱, 2:核准, 3:核駁, 4:???(0), 5:承辦人員工代號, 6:使用者代號
'Modify By Sindy 2011/11/1 +,核准(本所註冊費),TM01,TM02,TM03,TM04
cnnConnection.Execute "INSERT INTO R020406 " & strSql
'Add By Sindy 2011/11/1 計算核准件數是否為本所繳註冊費
cnnConnection.Execute "UPDATE R020406 Set R075007 = 1 " & _
                      "WHERE R075008||R075009||R075010||R075011 in ( " & _
                      "SELECT CP01||CP02||CP03||CP04 FROM R020406,Caseprogress " & _
                      "WHERE ID='" & strUserNum & "' AND R075003=1 " & _
                      "AND R075008=CP01 AND R075009=CP02 AND R075010=CP03 AND R075011=CP04 " & _
                      "AND CP10 in ('715','717') AND CP27 is not null) "
'2011/11/1 End
CheckOC
strSql = "SELECT * FROM R020406 WHERE ID='" & strUserNum & "' "
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/19
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/19
        ShowNoData
        Screen.MousePointer = vbDefault
        CheckOC
        Exit Sub
    End If
End With
CheckOC
'若申請國家為大陸
If Me.txt1(3).Text = "020" Then
   PrintData
Else
   PrintData2 'Add By Sindy 2011/11/1
End If
Screen.MousePointer = vbDefault
End Sub

Sub PrintData()
BolEndThisPage = False
strSql = "SELECT st02,NVL(A0902,A0903),SUM(r075003),SUM(r075004),SUM(r075005),r075006,r075001,r075002 FROM r020406,staff,acc090 WHERE r075002=a0901(+) and ID='" & strUserNum & "' and r075001=st01(+) GROUP BY st03,r075006,r075001,r075002,st02,NVL(A0902,A0903) ORDER BY st03,r075006,r075001,r075002 "
CheckOC
Page = 1
'Add By Cheng 2002/05/07
m_strPromoter = ""
strTemp3 = ""
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        PrintTitle
        strTemp3 = CheckStr(.Fields(6))
        Do While .EOF = False
            For i = 0 To 6
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If strTemp3 <> strTemp(6) Then
                'Add By Cheng 2002/05/07
                m_strPromoter = ""
                ShowLine
                PrintEnd (0)
                strTemp3 = strTemp(6)
                If iPrint > 2000 Then
                  ShowLine
                End If
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
         CheckOC
         ShowNoData
         Exit Sub
         '2011/3/1 End
    End If
End With
CheckOC
ShowLine
strTemp3 = strTemp(6)
PrintEnd (0)
If iPrint > 2000 Then
   ShowLine
End If
PrintEnd (1)
If iPrint > 2000 Then
   ShowLine
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintEnd(Strindex As Integer)
Select Case Strindex
Case 0
     If Len(Trim(strTemp3)) = 0 Then
         strSql = "select '個人小計','',SUM(r075003),SUM(r075004),SUM(r075005),'' FROM r020406 WHERE ID='" & strUserNum & "' AND (r075001='' OR R075001 IS NULL) GROUP BY r075001 "
     Else
         strSql = "select '個人小計','',SUM(r075003),SUM(r075004),SUM(r075005),'' FROM r020406 WHERE ID='" & strUserNum & "' AND r075001='" & strTemp3 & "' GROUP BY r075001 "
     End If
Case 1
     strSql = "select '全所總計','',sum(r075003),sum(r075004),sum(r075005),'' FROM r020406 WHERE ID='" & strUserNum & "' "
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
End Sub

Sub PrintTitle()
GetPleft
iPrint = 0
'Printer.Orientation = 1 'Removed by Morgan 2015/6/3
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "商申承辦人准駁統計表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 7000
Printer.CurrentY = iPrint
'Modify By Cheng 2003/12/02
'Printer.Print "公告日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
'Printer.Print IIf(Me.txt1(3).Text = "020", "准駁日：", "公告日：") & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
Printer.Print "准駁日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
'End
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 16500
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300

'Add By Cheng 2002/02/05
Printer.CurrentX = 250
Printer.CurrentY = iPrint
Printer.Print "申請國家：" & Me.txt1(3).Text & " － " & Me.txt1(4).Text
Printer.CurrentX = 6750
Printer.CurrentY = iPrint
Printer.Print "系統類別：" & Me.txt1(0).Text

Printer.CurrentX = 16500
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
'For I = 2 To 18 Step 2
'    Printer.Line (PLeft(I) - 50, iPrint + 150)-(PLeft(I) - 50, iPrint + 1350)
'Next I
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "業務區別"
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
If m_strPromoter <> strTemp(0) Then
   Printer.Print strTemp(0)
   m_strPromoter = strTemp(0)
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

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Printer = Printers(SeekPrint)
'Printer.Orientation = SeekPrintL
Set frm020406 = Nothing
End Sub

Private Sub txt1_Change(Index As Integer)
    'Add By Cheng 2003/12/02
    Select Case Index
    Case 3 '申請國家
        Me.txt1(4).Text = Me.txt1(3).Text
        If Me.txt1(3).Text = "020" Then
            Me.Label1(2).Caption = "准駁日："
        Else
            Me.Label1(2).Caption = "註冊公告日："
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
Case 4
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
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

Sub PrintData2()
BolEndThisPage = False
'strSql = "SELECT st02,NVL(A0902,A0903),sum(r075007),SUM(r075004),SUM(r075005),r075006,r075001,SUM(r075003),0,r075002 FROM r020406,staff,acc090 WHERE r075002=a0901(+) and ID='" & strUserNum & "' and r075001=st01(+) GROUP BY st03,r075006,r075001,r075002,st02,NVL(A0902,A0903) ORDER BY st03,r075006,r075001,r075002 "
strSql = "SELECT st02,NVL(A0902,A0903),sum(r075007),SUM(r075004),0,SUM(r075003),SUM(r075004),0,r075006,r075001,r075002 FROM r020406,staff,acc090 WHERE r075002=a0901(+) and ID='" & strUserNum & "' and r075001=st01(+) GROUP BY st03,r075006,r075001,r075002,st02,NVL(A0902,A0903) ORDER BY st03,r075006,r075001,r075002 "
CheckOC
Page = 1
m_strPromoter = ""
strTemp3 = ""
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        PrintTitle2
        strTemp3 = CheckStr(.Fields(9))
        Do While .EOF = False
            For i = 0 To 9
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If strTemp3 <> strTemp(9) Then
                m_strPromoter = ""
                ShowLine2
                PrintEnd2 (0)
                strTemp3 = strTemp(9)
                If iPrint > 2300 Then
                  ShowLine2
                End If
            End If
            '核准率
            If Val(strTemp(3)) = 0 Then
                If Val(strTemp(2)) = 0 Then
                    strTemp(4) = "0"
                Else
                    strTemp(4) = "100"
                End If
            Else
                strTemp(4) = Trim(str(Val(strTemp(2)) / (Val(strTemp(2)) + Val(strTemp(3))) * 100))
            End If
            If Val(strTemp(6)) = 0 Then
                If Val(strTemp(5)) = 0 Then
                    strTemp(7) = "0"
                Else
                    strTemp(7) = "100"
                End If
            Else
                strTemp(7) = Trim(str(Val(strTemp(5)) / (Val(strTemp(5)) + Val(strTemp(6))) * 100))
            End If
            PrintDatil2
            If iPrint >= 11000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle2
            End If
            .MoveNext
        Loop
    Else
         CheckOC
         ShowNoData
         Exit Sub
    End If
End With
CheckOC
ShowLine2
strTemp3 = strTemp(9)
PrintEnd2 (0)
If iPrint > 2300 Then
   ShowLine2
End If
PrintEnd2 (1)
If iPrint > 2300 Then
   ShowLine2
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintEnd2(Strindex As Integer)
Select Case Strindex
Case 0
     If Len(Trim(strTemp3)) = 0 Then
         strSql = "select '個人小計','',SUM(r075007),SUM(r075004),0,sum(r075003),SUM(r075004),0 FROM r020406 WHERE ID='" & strUserNum & "' AND (r075001='' OR R075001 IS NULL) GROUP BY r075001 "
     Else
         strSql = "select '個人小計','',SUM(r075007),SUM(r075004),0,sum(r075003),SUM(r075004),0 FROM r020406 WHERE ID='" & strUserNum & "' AND r075001='" & strTemp3 & "' GROUP BY r075001 "
     End If
Case 1
     strSql = "select '全所總計','',SUM(r075007),sum(r075004),0,sum(r075003),SUM(r075004),0 FROM r020406 WHERE ID='" & strUserNum & "' "
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
            For i = 0 To 7
                StrTemp7(i) = CheckStr(.Fields(i))
            Next i
            '核准率
            If Val(StrTemp7(3)) = 0 Then
                If Val(StrTemp7(2)) = 0 Then
                    StrTemp7(4) = "0"
                Else
                    StrTemp7(4) = "100"
                End If
            Else
                StrTemp7(4) = Trim(str(Val(StrTemp7(2)) / (Val(StrTemp7(2)) + Val(StrTemp7(3))) * 100))
            End If
            If Val(StrTemp7(6)) = 0 Then
                If Val(StrTemp7(5)) = 0 Then
                    StrTemp7(7) = "0"
                Else
                    StrTemp7(7) = "100"
                End If
            Else
                StrTemp7(7) = Trim(str(Val(StrTemp7(5)) / (Val(StrTemp7(5)) + Val(StrTemp7(6))) * 100))
            End If
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            For i = 2 To 6
               If i = 2 Or i = 3 Or i = 5 Or i = 6 Then
                  Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(StrTemp7(i))
                  Printer.CurrentY = iPrint
                  Printer.Print StrTemp7(i)
               End If
            Next i
            For i = 4 To 7
               If i = 4 Or i = 7 Then
                  Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(StrTemp7(i), "###.00"))
                  Printer.CurrentY = iPrint
                  Printer.Print Format(StrTemp7(i), "###.00") & "%"
               End If
            Next i
            iPrint = iPrint + 300
            If iPrint >= 11000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle2
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

Sub PrintTitle2()
GetPleft2
iPrint = 300
'Printer.Orientation = 1 'Removed by Morgan 2015/6/3
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "商申承辦人准駁統計表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 7000
Printer.CurrentY = iPrint
Printer.Print "註冊公告日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 14000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300

Printer.CurrentX = 250
Printer.CurrentY = iPrint
Printer.Print "申請國家：" & Me.txt1(3).Text & " － " & Me.txt1(4).Text
Printer.CurrentX = 7000
Printer.CurrentY = iPrint
Printer.Print "　系統類別：" & Me.txt1(0).Text

Printer.CurrentX = 14000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 11000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle2
    Exit Sub
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "業務區別"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "核准(本所註冊費)"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "核駁"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "核准率"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "核准(公告)"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "核駁"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "核准率"
iPrint = iPrint + 300
If iPrint >= 11000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle2
    Exit Sub
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 11000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle2
    Exit Sub
End If
End Sub

Sub PrintDatil2()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
If m_strPromoter <> strTemp(0) Then
   Printer.Print strTemp(0)
   m_strPromoter = strTemp(0)
Else
   Printer.Print ""
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
For i = 2 To 6
   If i = 2 Or i = 3 Or i = 5 Or i = 6 Then
      Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(strTemp(i))
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
   End If
Next i
For i = 4 To 7
   If i = 4 Or i = 7 Then
      Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(strTemp(i), "###.00"))
      Printer.CurrentY = iPrint
      Printer.Print Format(strTemp(i), "###.00") + "%"
   End If
Next
iPrint = iPrint + 300
End Sub

Sub GetPleft2()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 2500
PLeft(2) = 4500
PLeft(3) = 6500
PLeft(4) = 8500
PLeft(5) = 10500
PLeft(6) = 12500
PLeft(7) = 14500
End Sub

Sub ShowLine2()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 11000 Then
   If BolEndThisPage = False Then
      Page = Page + 1
      Printer.NewPage
      PrintTitle2
   Else
      BolEndThisPage = False
   End If
End If
End Sub
