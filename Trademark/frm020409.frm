VERSION 5.00
Begin VB.Form frm020409 
   BorderStyle     =   1  '單線固定
   Caption         =   "商爭案承辦人預估勝敗統計表"
   ClientHeight    =   2565
   ClientLeft      =   3495
   ClientTop       =   1635
   ClientWidth     =   3870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3870
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   540
      Left            =   45
      TabIndex        =   11
      Top             =   1650
      Width           =   3825
      Begin VB.ComboBox Combo1 
         Height          =   300
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
      Left            =   3060
      TabIndex        =   7
      Top             =   24
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2268
      TabIndex        =   6
      Top             =   24
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2145
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1320
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1125
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1320
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2145
      MaxLength       =   7
      TabIndex        =   2
      Top             =   960
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1125
      MaxLength       =   7
      TabIndex        =   1
      Top             =   960
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1125
      TabIndex        =   0
      Top             =   600
      Width           =   1740
   End
   Begin VB.Line Line2 
      X1              =   1665
      X2              =   2415
      Y1              =   1410
      Y2              =   1410
   End
   Begin VB.Line Line1 
      X1              =   1575
      X2              =   2835
      Y1              =   1110
      Y2              =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   10
      Top             =   1350
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "勝敗日期："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   9
      Top             =   990
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   615
      Width           =   915
   End
End
Attribute VB_Name = "frm020409"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/10 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, SavDay4 As String, SavDay5 As String, SavDay6 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 41) As String, strTemp3 As String, TestOk As Boolean, StrTemp7(0 To 40) As String
Dim PLeft(0 To 39) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, PLeft1(1 To 7) As Integer, SeekPrint As Integer, SeekPrintL As Integer, k As Integer
'Add By Cheng 2002/05/06
Dim m_strPromoter As String '承辦人
'Add By Cheng 2003/03/12
Dim strSQL6_1 As String
Dim strSQL6_2 As String     '2009/5/15 add by sonia
Dim bolNotData As Boolean 'Add By Sindy 2011/3/1
'Modify By Sindy 2015/2/4
Dim System_ID As String
Dim bolIsChina As Boolean
Dim Title_601 As String
Dim Title_603 As String
Dim Title_605 As String
Dim Title_401 As String
Dim Title_403 As String
Dim Title_404 As String
Dim Title_408 As String
Dim Title_602 As String
Dim Title_604 As String
Dim Title_606 As String
Dim Title_406 As String
Dim Title_407 As String
'2015/2/4 END


Private Sub cmdOK_Click(Index As Integer)
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
             s = MsgBox("勝敗日期區間不可空白!!", , "USER 輸入錯誤")
             txt1(1).SetFocus
             txt1_GotFocus (1)
             Exit Sub
         Else
            'Modify By Cheng 2004/03/11
'             If Format(ChangeTStringToTDateString(txt1(1)), "yy") = Format(ChangeTStringToTDateString(txt1(2)), "yy") Then
'                 If Format(ChangeTStringToTDateString(txt1(1)), "mm") = Format(ChangeTStringToTDateString(txt1(2)), "mm") Then
'                    If Val(Format(ChangeTStringToTDateString(txt1(1)), "dd")) = 1 Then
             If Format(ChangeTStringToWDateString(txt1(1)), "yyyy") = Format(ChangeTStringToWDateString(txt1(2)), "yyyy") Then
                 If Format(ChangeTStringToWDateString(txt1(1)), "mm") = Format(ChangeTStringToWDateString(txt1(2)), "mm") Then
                    If Val(Format(ChangeTStringToWDateString(txt1(1)), "dd")) = 1 Then
            'End
                        'Modify By Cheng 2003/03/26
'                        If Format(ChangeTStringToTDateString(txt1(1)), "mm") <> Format(DateAdd("d", 1, ChangeTStringToTDateString(txt1(2))), "mm") Then
                        If Format(ChangeTStringToWDateString(txt1(1)), "mm") <> Format(DateAdd("d", 1, ChangeTStringToWDateString(txt1(2))), "mm") Then
                            TestOk = True
                        Else
                            TestOk = False
                        End If
                    Else
                        TestOk = False
                    End If
                 Else
                    TestOk = False
                 End If
             Else
                 TestOk = False
             End If

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
'Add By Cheng 2003/03/11
Dim StrSQLa As String

Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R020409 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL6 = ""
strSQL6_1 = "": strSQL6_2 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
    'Modify By Cheng 2003/03/11
'   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/10/19
End If
'Add By Cheng 2003/07/11
'若非商申收發文時, 若為TF案則不抓後三碼為"000"的資料
strSQL1 = strSQL1 + " AND CP03 <> Decode(CP01,'TF','0','z') AND CP04 <> Decode(CP01,'TF','00','zz') "
StrSQL6 = ""
'2008/11/18 modify by sonia 申請意見書不抓公告日改抓審定來函日,大陸案已無此案件性質
strSQL6_2 = strSQL6_1
If Len(txt1(1)) <> 0 Then
   StrSQL6 = StrSQL6 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(1))) & ""
   strSQL6_1 = strSQL6_1 + " AND TM14>=" & Val(ChangeTStringToWString(txt1(1))) & ""
   strSQL6_2 = strSQL6_2 + " AND TM13>=" & Val(ChangeTStringToWString(txt1(1))) & ""
End If
If Len(Trim(txt1(2))) <> 0 Then
   StrSQL6 = StrSQL6 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(2))) & " "
   strSQL6_1 = strSQL6_1 + " AND TM14<=" & Val(ChangeTStringToWString(txt1(2))) & " "
   strSQL6_2 = strSQL6_2 + " AND TM13<=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
If Len(txt1(1)) <> 0 Or Len(Trim(txt1(2))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/10/19
End If
If Len(txt1(3)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND TM10>='" & txt1(3) & "' "
    strSQL6_1 = strSQL6_1 + " AND TM10>='" & txt1(3) & "' "
End If
If Len(txt1(4)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND TM10<='" & txt1(4) & "' "
    strSQL6_1 = strSQL6_1 + " AND TM10<='" & txt1(4) & "' "
End If
If Len(txt1(3)) <> 0 Or Len(Trim(txt1(4))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/10/19
End If
'Add By Cheng 2002/02/07
'Modify By Sindy 98/04/16
'strSQL1 = strSQL1 + " And (CP10='1003' OR CP10='1004') "
strSQL1 = strSQL1 + " And (CP10='1003' OR CP10='1004' OR CP10='1006') "
'Add By Cheng 2003/03/11
'2011/12/7 modify by sonia
'strSQL2 = strSQL2 + " And CP10='202' "
strSQL2 = strSQL2 + " And CP10 in ('202','210') "
'2011/12/7 end
'93.6.14 CANCEL BY SONIA 僅收/發文統計表才控制不算案件數
''Add By Cheng 2004/04/29
''抓計件的資料
'StrSQL6 = StrSQL6 & " And CP26 Is Null "
'strSQL6_1 = strSQL6_1 & " And CP26 Is Null "
''End
CheckOC
'Modify By Cheng 2002/02/07
'strSQL = "SELECT s2.st02,NVL(A0902,A0903),cp23,CP24,S2.ST01,CP10,decode(tm10,'000','*','') FROM CASEPROGRESS,STAFF S1,STAFF S2,TRADEMARK,ACC090 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') AND (S2.ST05='97' OR S2.ST05='17') AND CP09>'C' AND S1.ST03=A0901(+) " & strSQL1 + StrSQL6
'取消抓承辦人之ST05為"97"或"17"及CP09>"C"的條件
'先設定相關總收文號
strSql = "SELECT CP43 " & _
         " FROM CASEPROGRESS,TRADEMARK " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND (CP57 IS NULL OR CP57='') " & strSQL1 + StrSQL6
'Add By Cheng 2003/03/11
'抓案件性質為申請意見書的本所案號
'strSQLA = "Select DISTINCT CP09 From CaseProgress , ( SELECT CP01 C1, CP02 C2, CP03 C3 , CP04 C4 " & _
'         " FROM CASEPROGRESS,TRADEMARK " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND (CP57 IS NULL OR CP57='') And (CP22 IS NULL OR CP22 <>'N') " & strSQL2 + StrSQL6 & " GROUP BY CP01,CP02, CP03, CP04 ) C WHERE CP01=C.C1 AND CP02=C.C2 AND CP03=C.C3 AND CP04=C.C4 "
StrSQLa = "SELECT DISTINCT CP09 " & _
         " FROM CASEPROGRESS,TRADEMARK " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND (CP57 IS NULL OR CP57='') " & strSQL2 + strSQL6_2

'910521 nick
'strSQL = "SELECT st02,NVL(A0902,A0903),cp23,CP24,ST01,CP10,decode(tm10,'000','*','') " & _
         " FROM CASEPROGRESS,STAFF,TRADEMARK,ACC090 " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND CP09 IN ( " & strSQL & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(ST03,1,2)='F1' OR CP14 IS NULL) ")
'**************** 將業務區改成抓案件進度檔   91.08.15  nick
'Modify By Cheng 2003/06/03
'預估結果(CP23)要有值
'strSQL = "SELECT CP14,CP12,CP23,CP24,ST01,CP10,decode(tm10,'000','*','') " & _
'         " FROM CASEPROGRESS,STAFF,TRADEMARK,ACC090 " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND CP09 IN ( " & strSQL & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(ST03,1,2)='F1' OR CP14 IS NULL) ")
'93.6.14 MODIFY BY SONIA 不管是否計件
'strSQL = "SELECT CP14,CP12,CP23,CP24,ST01,CP10,decode(tm10,'000','*','') " & _
'         " FROM CASEPROGRESS,STAFF,TRADEMARK,ACC090 " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') And CP23 Is Not Null AND CP12=A0901(+) AND CP09 IN ( " & strSQL & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(ST03,1,2)='F1' OR CP14 IS NULL) ")
'edit by nickc 2005/05/12
'StrSql = "SELECT CP14,CP12,CP23,CP24,ST01,CP10,decode(tm10,'000','*','') " & _
         " FROM CASEPROGRESS,STAFF,TRADEMARK,ACC090 " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND (CP57 IS NULL OR CP57='') And CP23 Is Not Null AND CP12=A0901(+) AND CP09 IN ( " & StrSql & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(ST03,1,2)='F1' OR CP14 IS NULL) ")
strSql = "SELECT CP14,CP12,CP23,CP24,S1.ST01,CP10,decode(substr(S2.St15,1,1),'F','','*') " & _
         " FROM CASEPROGRESS,STAFF S1,ACC090,staff S2 " & _
         " WHERE cp13=S2.st01(+) AND CP14=S1.ST01(+) AND (CP57 IS NULL OR CP57='') And CP23 Is Not Null AND CP12=A0901(+) AND CP09 IN ( " & strSql & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(S1.ST03,1,2)='P2' OR CP14 IS NULL Or CP14='A6015') ", " AND (SUBSTR(S1.ST03,1,2)='F1' OR CP14 IS NULL) ")

'93.6.14 END
'Add By Cheng 2003/03/11
'與申請意見書相同本所案號且案件性質為申請的資料
'strSQL = strSQL & " UNION ALL SELECT CP14,CP12,CP23,CP24,ST01,CP10,decode(tm10,'000','*','') " & _
'         " FROM CASEPROGRESS,STAFF,TRADEMARK,ACC090 " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND CP10='101' AND ( CP23 IS NOT NULL OR CP24 IS NOT NULL ) AND CP09 IN ( " & strSQLA & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(ST03,1,2)='F1' OR CP14 IS NULL) ")
'93.6.14 MODIFY BY SONIA 不管是否計件
'strSQL = strSQL & " UNION ALL SELECT CP14,CP12,CP23,TM16,ST01,CP10,decode(tm10,'000','*','') " & _
'         " FROM CASEPROGRESS,STAFF,TRADEMARK,ACC090 " & _
'         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND (CP26 IS NULL OR CP26='') AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND (CP22 IS NULL OR CP22 <> 'N') AND (CP23 IS NOT NULL OR TM16 IS NOT NULL) AND CP09 IN ( " & strSQLA & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(ST03,1,2)='F1' OR CP14 IS NULL) ")
'edit by nickc 2005/05/12
'StrSql = StrSql & " UNION ALL SELECT CP14,CP12,CP23,TM16,ST01,CP10,decode(tm10,'000','*','') " & _
         " FROM CASEPROGRESS,STAFF,TRADEMARK,ACC090 " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=ST01(+) AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND (CP22 IS NULL OR CP22 <> 'N') AND (CP23 IS NOT NULL OR TM16 IS NOT NULL) AND CP09 IN ( " & strSQLA & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(ST03,1,2)='P2' OR CP14 IS NULL) ", " AND (SUBSTR(ST03,1,2)='F1' OR CP14 IS NULL) ")
strSql = strSql & " UNION ALL SELECT CP14,CP12,CP23,TM16,S1.ST01,CP10,decode(substr(S2.St15,1,1),'F','','*') " & _
         " FROM CASEPROGRESS,STAFF S1,TRADEMARK,ACC090,staff S2 " & _
         " WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) and cp13=s2.st01(+) AND (CP57 IS NULL OR CP57='') AND CP12=A0901(+) AND (CP22 IS NULL OR CP22 <> 'N') AND (CP23 IS NOT NULL OR TM16 IS NOT NULL) AND CP09 IN ( " & StrSQLa & " ) " & IIf(intPWhere = 國內, " AND (SUBSTR(S1.ST03,1,2)='P2' OR CP14 IS NULL Or CP14='A6015') ", " AND (SUBSTR(S1.ST03,1,2)='F1' OR CP14 IS NULL) ")

'93.6.14 END

With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/19
        .MoveFirst
        DoEvents
        Do While .EOF = False
            For i = 0 To 6
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            'Modify By Sindy 98/04/13 增加3.部分勝部分敗
            If Val(strTemp(2)) = Val(strTemp(3)) Or Val(strTemp(2)) = 3 Then
                If Val(strTemp(2)) = 1 Or Val(strTemp(2)) = 3 Then '勝
                    SavDay1 = "1"
                    SavDay2 = "0"
                    SavDay3 = "0"
                    SavDay4 = "0"
                Else
                    SavDay1 = "0"
                    SavDay2 = "1"
                    SavDay3 = "0"
                    SavDay4 = "0"
                End If
            Else
                If Val(strTemp(3)) = 1 Then '勝
                    SavDay1 = "0"
                    SavDay2 = "0"
                    SavDay3 = "1"
                    SavDay4 = "0"
                Else
                    SavDay1 = "0"
                    SavDay2 = "0"
                    SavDay3 = "0"
                    SavDay4 = "1"
                End If
            End If
            Select Case Val(strTemp(5))
'******表(1)格式1***********
            Case 601, 627 '異議,Add by Sindy 2019/8/15 +部分異議
                 cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081003,r081004,r081005,r081006,r081007,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',1,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
            Case 603, 629 '評定,Add by Sindy 2019/8/15 +部分評定
                 cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081008,r081009,r081010,r081011,r081012,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',1,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
            Case 605, 623 '廢止,Add by Sindy 2019/8/15 +部分廢止
                 cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081013,r081014,r081015,r081016,r081017,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',1,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
'            'add by nickc 2007/07/27 商標處改格式
'            Case 618
'                If txt1(3) = "020" And txt1(4) = "020" Then
'                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081018,r081019,r081020,r081021,r081022,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',1,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
'                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
'                End If
            'Modify By Cheng 2002/05/06
'            Case 401 '訴願
            'Modify By Cheng 2003/03/11
            '參加訴願移至表(2)
'            Case 401, 406 '訴願, 參加訴願
            Case 401 '訴願
                'edit by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',1,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                Else
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081018,r081019,r081020,r081021,r081022,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',1,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                End If
            'Modify By Cheng 2002/05/06
            '再訴願欄以行政訴訟取代之
'            Case 402 '再訴願
            'Modify By Cheng 2003/03/11
            '參加訴訟移至表(2)
'            Case 403, 407 '行政訴訟,參加訴訟
            Case 403 '行政訴訟
                'edit by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081028,r081029,r081030,r081031,r081032,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',1,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                Else
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',1,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                End If
            'Modify By Cheng 2002/05/06
            '本欄以行政訴訟上訴取代
'            Case 403 '行政訴訟
            Case 408 '行政訴訟上訴
                'edit by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081028,r081029,r081030,r081031,r081032,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',1,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                End If
            'Modify By Chneg 2002/04/12
'            Case 405
            Case 404 '再審之訴
                'edit by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081033,r081034,r081035,r081036,r081037,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',1,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                End If
            'Move By Cheng 2003/03/11
'            Case 101 '申請意見書(<--申請)
            Case 202 '申請意見書
                'edit by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081041,r081042,r081043,r081044,r081045,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',1,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                End If
'************************************
'******表(2)*********
            Case 602, 628 '異議答辯,Add by Sindy 2019/8/15 +部分異議答辯
                 cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081003,r081004,r081005,r081006,r081007,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
            Case 604, 630 '評定答辯,Add by Sindy 2019/8/15 +部分評定答辯
                 cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081008,r081009,r081010,r081011,r081012,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
            Case 606, 624 '廢止答辯,Add by Sindy 2019/8/15 +部分廢止答辯
                 cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081013,r081014,r081015,r081016,r081017,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                 cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
'            'add by nickc 2007/07/27 商標處改格式
'            Case 619
'                If txt1(3) = "020" And txt1(4) = "020" Then
'                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081018,r081019,r081020,r081021,r081022,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
'                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
'                End If
            'Modify By Cheng 2003/03/11
            '申請意見書移至表(1)
'            Case 202 '申請意見書
'                 cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081018,r081019,r081020,r081021,r081022,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
'                 cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
            Case 406 '參加訴願
                'edit by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081028,r081029,r081030,r081031,r081032,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                Else
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081018,r081019,r081020,r081021,r081022,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                End If
            Case 407 '參加訴訟
                'edit by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081028,r081029,r081030,r081031,r081032,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                End If
            Case 410 '上訴答辯
                'edit by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" Then
                Else
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081033,r081034,r081035,r081036,r081037,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                    cnnConnection.Execute "INSERT INTO r020409 (r081001,r081002,r081023,r081024,r081025,r081026,r081027,r081038,r081039,r081040,ID) VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "'," & Val(SavDay1) & "," & Val(SavDay2) & "," & Val(SavDay3) & "," & Val(SavDay4) & ",0,'" & ChgSQL(strTemp(6)) & "',2,'" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                End If
            Case Else
            End Select
            DoEvents
            .MoveNext
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/19
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End With
CheckOC
'Modify By Sindy 2015/2/4
If Trim(txt1(0)) = "" Then
   strTemp1 = Split(GetSystemKindByNickT, ",")
Else
   strTemp1 = Split(txt1(0), ",")
End If
System_ID = strTemp1(0)
bolIsChina = False
If txt1(3) = "020" And txt1(4) = "020" Then
   bolIsChina = True
End If
Call ClsPDGetCaseProperty("T", "601", Title_601, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "603", Title_603, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "605", Title_605, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "401", Title_401, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "403", Title_403, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "404", Title_404, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "408", Title_408, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "602", Title_602, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "604", Title_604, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "606", Title_606, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "406", Title_406, bolIsChina, False)
Call ClsPDGetCaseProperty("T", "407", Title_407, bolIsChina, False)
'2015/2/4 END
bolNotData = True 'Add By Sindy 2011/3/1
PrintData
'Add By Sindy 2011/3/1
If bolNotData = True Then
   ShowNoData
'2011/3/1 End
Else
   ShowPrintOk
End If
Screen.MousePointer = vbDefault
End Sub

Sub PrintData()

'列印表(1)格式1
'Modify By Cheng 2003/03/11
'strSQL = "select st02,NVL(A0902,A0903),sum(r081003),sum(r081004),sum(r081005),sum(r081006),sum(r081007),sum(r081008),sum(r081009),sum(r081010),sum(r081011)" & _
'         ",sum(r081012),sum(r081013),sum(r081014),sum(r081015),sum(r081016),sum(r081017),sum(r081018),sum(r081019),sum(r081020),sum(r081021),sum(r081022)" & _
'         ",sum(r081023),SUM(r081024),SUM(R081025),SUM(R081026),SUM(R081027),SUM(R081028),SUM(R081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081033)" & _
'         ",SUM(R081034),SUM(R081035),SUm(R081036),SUM(R081037),'',R081039,R081040,r081001,r081002 from r020409,staff,acc090 where r081001=st01(+) and r081002=a0901(+) and id='" & strUserNum & "' group by r081001,r081002,R081039,R081040,st02,NVL(A0902,A0903) ORDER BY r081039,r081040, R081001,R081002"
strSql = "select st02,NVL(A0902,A0903),sum(r081003),sum(r081004),sum(r081005),sum(r081006),sum(r081007),sum(r081008),sum(r081009),sum(r081010),sum(r081011)" & _
         ",sum(r081012),sum(r081013),sum(r081014),sum(r081015),sum(r081016),sum(r081017),sum(r081018),sum(r081019),sum(r081020),sum(r081021),sum(r081022)" & _
         ",sum(r081023),SUM(r081024),SUM(R081025),SUM(R081026),SUM(R081027),SUM(R081028),SUM(R081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081033)" & _
         ",SUM(R081034),SUM(R081035),SUm(R081036),SUM(R081037),'',R081039,R081040,r081001,r081002,SUM(R081041),SUM(R081042),SUm(R081043),SUM(R081044),SUM(R081045) from r020409,staff,acc090 where r081001=st01(+) and r081002=a0901(+) and r081039='1' and id='" & strUserNum & "' group by r081001,r081002,R081039,R081040,st02,NVL(A0902,A0903) ORDER BY r081039,r081040, R081001,R081002"
CheckOC
'Add By Cheng 2002/05/06
m_strPromoter = ""
SavDay1 = ""
SavDay2 = ""
Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        bolNotData = False 'Add By Sindy 2011/3/1
        .MoveFirst
        SavDay1 = CheckStr(.Fields(38))
        SavDay2 = CheckStr(.Fields(39))
        SavDay5 = CheckStr(.Fields(40))
        SavDay6 = CheckStr(.Fields(41))
        PrintTitle
        If Val(SavDay1) = 1 Then
            PrintTitle1
        Else
            PrintTitle2
        End If
        Do While .EOF = False
            For i = 0 To 41
                strTemp(i) = CheckStr(.Fields(i))
                'Modify By Sindy 2012/6/22 +And i <> 40 員工編號是英數字
                If Val(strTemp(i)) = 0 And i <> 0 And i <> 1 And i <> 37 And i <> 38 And i <> 39 And i <> 40 Then
                    strTemp(i) = "0"
                End If
            Next i
            If Val(strTemp(2)) + Val(strTemp(3)) + Val(strTemp(4)) + Val(strTemp(5)) = 0 Then
                strTemp(6) = "0"
            Else
                strTemp(6) = Trim(str((Val(strTemp(2)) + Val(strTemp(3))) / (Val(strTemp(2)) + Val(strTemp(3)) + Val(strTemp(4)) + Val(strTemp(5))) * 100))
            End If
            If Val(strTemp(7)) + Val(strTemp(8)) + Val(strTemp(9)) + Val(strTemp(10)) = 0 Then
                strTemp(11) = "0"
            Else
                strTemp(11) = Trim(str((Val(strTemp(8)) + Val(strTemp(7))) / (Val(strTemp(8)) + Val(strTemp(9)) + Val(strTemp(10)) + Val(strTemp(7))) * 100))
            End If
            If Val(strTemp(12)) + Val(strTemp(13)) + Val(strTemp(14)) + Val(strTemp(15)) = 0 Then
                strTemp(16) = "0"
            Else
                strTemp(16) = Trim(str((Val(strTemp(12)) + Val(strTemp(13))) / (Val(strTemp(12)) + Val(strTemp(13)) + Val(strTemp(14)) + Val(strTemp(15))) * 100))
            End If
            If Val(strTemp(17)) + Val(strTemp(18)) + Val(strTemp(19)) + Val(strTemp(20)) = 0 Then
                strTemp(21) = "0"
            Else
                strTemp(21) = Trim(str((Val(strTemp(17)) + Val(strTemp(18))) / (Val(strTemp(17)) + Val(strTemp(18)) + Val(strTemp(19)) + Val(strTemp(20))) * 100))
            End If
            If Val(strTemp(22)) + Val(strTemp(23)) + Val(strTemp(24)) + Val(strTemp(25)) = 0 Then
                strTemp(26) = "0"
            Else
                strTemp(26) = Trim(str((Val(strTemp(22)) + Val(strTemp(23))) / (Val(strTemp(22)) + Val(strTemp(23)) + Val(strTemp(24)) + Val(strTemp(25))) * 100))
            End If
            If Val(strTemp(27)) + Val(strTemp(28)) + Val(strTemp(29)) + Val(strTemp(30)) = 0 Then
                strTemp(31) = "0"
            Else
                strTemp(31) = Trim(str((Val(strTemp(27)) + Val(strTemp(28))) / (Val(strTemp(27)) + Val(strTemp(28)) + Val(strTemp(29)) + Val(strTemp(30))) * 100))
            End If
            If Val(strTemp(32)) + Val(strTemp(33)) + Val(strTemp(34)) + Val(strTemp(35)) = 0 Then
                strTemp(36) = "0"
            Else
                strTemp(36) = Trim(str((Val(strTemp(32)) + Val(strTemp(33))) / (Val(strTemp(32)) + Val(strTemp(33)) + Val(strTemp(34)) + Val(strTemp(35))) * 100))
            End If
'91/05/21  nick
            If SavDay5 <> strTemp(40) Then
                'Add By Cheng 2002/05/06
                m_strPromoter = ""
                ShowLine
                PrintEnd (2)
                ShowLine
                PrintEnd (3)
                ShowLine
                PrintEnd (0)
                ShowLine
                SavDay5 = strTemp(40)
            End If
            If SavDay1 <> strTemp(38) Then
                PrintEnd (1)
                ShowLine
                Page = Page + 1
                Printer.NewPage
                SavDay1 = strTemp(38)
                PrintTitle
                If Val(SavDay1) = 1 Then
                    PrintTitle1
                Else
                    PrintTitle2
                End If
                SavDay2 = strTemp(39)
            End If
            strTemp(0) = StrToStr(strTemp(0), 4)
            strTemp(1) = StrToStr(strTemp(1), 4)
            PrintDatil
            If iPrint >= 13900 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                If Val(SavDay1) = 1 Then
                    PrintTitle1
                Else
                    PrintTitle2
                End If
            End If
            .MoveNext
        Loop
    Else
         GoTo ReadNext1: 'Add By Sindy 2011/3/1
    End If
End With
CheckOC
ShowLine
PrintEnd (2)
ShowLine
PrintEnd (3)
ShowLine
PrintEnd (0)
ShowLine
PrintEnd (1)
ShowLine
Page = Page + 1
Printer.NewPage
ReadNext1: 'Add By Sindy 2011/3/1

'edit by nickc 2007/07/27 商標處改格式
If txt1(3) = "020" And txt1(4) = "020" Then
Else
    '列印表(1)格式2
    strSql = "select st02,NVL(A0902,A0903),sum(r081003),sum(r081004),sum(r081005),sum(r081006),sum(r081007),sum(r081008),sum(r081009),sum(r081010),sum(r081011)" & _
             ",sum(r081012),sum(r081013),sum(r081014),sum(r081015),sum(r081016),sum(r081017),sum(r081018),sum(r081019),sum(r081020),sum(r081021),sum(r081022)" & _
             ",sum(r081023),SUM(r081024),SUM(R081025),SUM(R081026),SUM(R081027),SUM(R081028),SUM(R081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081033)" & _
             ",SUM(R081034),SUM(R081035),SUm(R081036),SUM(R081037),'',R081039,R081040,r081001,r081002,SUM(R081041),SUM(R081042),SUm(R081043),SUM(R081044),SUM(R081045) from r020409,staff,acc090 where r081001=st01(+) and r081002=a0901(+) and r081039='1' and id='" & strUserNum & "' group by r081001,r081002,R081039,R081040,st02,NVL(A0902,A0903) ORDER BY r081039,r081040, R081001,R081002"
    CheckOC
    'Add By Cheng 2002/05/06
    m_strPromoter = ""
    SavDay1 = ""
    SavDay2 = ""
    'Page = 1
    With adoRecordset
        .CursorLocation = adUseClient
        .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If .RecordCount <> 0 And .RecordCount > 0 Then
        
            bolNotData = False 'Add By Sindy 2011/3/1
            .MoveFirst
            SavDay1 = CheckStr(.Fields(38))
            SavDay2 = CheckStr(.Fields(39))
            SavDay5 = CheckStr(.Fields(40))
            SavDay6 = CheckStr(.Fields(41))
            PrintTitle
            If Val(SavDay1) = 1 Then
                PrintTitle1_1
            Else
                PrintTitle2_2
            End If
            Do While .EOF = False
                For i = 0 To 41
                    strTemp(i) = CheckStr(.Fields(i))
                    'Modify By Sindy 2012/6/22 +And i <> 40 員工編號是英數字
                    If Val(strTemp(i)) = 0 And i <> 0 And i <> 1 And i <> 37 And i <> 38 And i <> 39 And i <> 40 Then
                        strTemp(i) = "0"
                    End If
                Next i
                strTemp(2) = Val("0" & .Fields(42))
                strTemp(3) = Val("0" & .Fields(43))
                strTemp(4) = Val("0" & .Fields(44))
                strTemp(5) = Val("0" & .Fields(45))
                strTemp(6) = Val("0" & .Fields(46))
                If Val(strTemp(2)) + Val(strTemp(3)) + Val(strTemp(4)) + Val(strTemp(5)) = 0 Then
                    strTemp(6) = "0"
                Else
                    strTemp(6) = Trim(str((Val(strTemp(2)) + Val(strTemp(3))) / (Val(strTemp(2)) + Val(strTemp(3)) + Val(strTemp(4)) + Val(strTemp(5))) * 100))
                End If
    '91/05/21  nick
                If SavDay5 <> strTemp(40) Then
                    'Add By Cheng 2002/05/06
                    m_strPromoter = ""
                    ShowLine_1
                    PrintEnd_1 (2)
                    ShowLine_1
                    PrintEnd_1 (3)
                    ShowLine_1
                    PrintEnd_1 (0)
                    ShowLine_1
                    SavDay5 = strTemp(40)
                End If
                If SavDay1 <> strTemp(38) Then
                    PrintEnd_1 (1)
                    ShowLine_1
                    Page = Page + 1
                    Printer.NewPage
                    SavDay1 = strTemp(38)
                    PrintTitle
                    If Val(SavDay1) = 1 Then
                        PrintTitle1_1
                    Else
                        PrintTitle2_2
                    End If
                    SavDay2 = strTemp(39)
                End If
                strTemp(0) = StrToStr(strTemp(0), 4)
                strTemp(1) = StrToStr(strTemp(1), 4)
                PrintDatil_1
                If iPrint >= 13900 Then
                    Page = Page + 1
                    Printer.NewPage
                    PrintTitle
                    If Val(SavDay1) = 1 Then
                        PrintTitle1_1
                    Else
                        PrintTitle2_2
                    End If
                End If
                .MoveNext
            Loop
        Else
            GoTo ReadNext2: 'Add By Sindy 2011/3/1
        End If
    End With
    CheckOC
    ShowLine_1
    PrintEnd_1 (2)
    ShowLine_1
    PrintEnd_1 (3)
    ShowLine_1
    PrintEnd_1 (0)
    ShowLine_1
    PrintEnd_1 (1)
    ShowLine_1
    Page = Page + 1
    'Printer.EndDoc 'Removed by Morgan 2015/6/15
End If
ReadNext2: 'Add By Sindy 2011/3/1

'列印表(2)格式1
strSql = "select st02,NVL(A0902,A0903),sum(r081003),sum(r081004),sum(r081005),sum(r081006),sum(r081007),sum(r081008),sum(r081009),sum(r081010),sum(r081011)" & _
         ",sum(r081012),sum(r081013),sum(r081014),sum(r081015),sum(r081016),sum(r081017),sum(r081018),sum(r081019),sum(r081020),sum(r081021),sum(r081022)" & _
         ",sum(r081023),SUM(r081024),SUM(R081025),SUM(R081026),SUM(R081027),SUM(R081028),SUM(R081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081033)" & _
         ",SUM(R081034),SUM(R081035),SUm(R081036),SUM(R081037),'',R081039,R081040,r081001,r081002,SUM(R081041),SUM(R081042),SUm(R081043),SUM(R081044),SUM(R081045) from r020409,staff,acc090 where r081001=st01(+) and r081002=a0901(+) and r081039='2' and id='" & strUserNum & "' group by r081001,r081002,R081039,R081040,st02,NVL(A0902,A0903) ORDER BY r081039,r081040, R081001,R081002"
CheckOC
'Add By Cheng 2002/05/06
m_strPromoter = ""
SavDay1 = ""
SavDay2 = ""
'Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        'Modified by Morgan 2015/11/10
        'bolNotData = False 'Add By Sindy 2011/3/1
        If bolNotData = False Then Printer.NewPage
        'end 2015/11/10
        
        .MoveFirst
        SavDay1 = CheckStr(.Fields(38))
        SavDay2 = CheckStr(.Fields(39))
        SavDay5 = CheckStr(.Fields(40))
        SavDay6 = CheckStr(.Fields(41))
        PrintTitle
        If Val(SavDay1) = 1 Then
            PrintTitle1
        Else
            PrintTitle2
        End If
        Do While .EOF = False
            For i = 0 To 41
                strTemp(i) = CheckStr(.Fields(i))
                'Modify By Sindy 2012/6/22 +And i <> 40 員工編號是英數字
                If Val(strTemp(i)) = 0 And i <> 0 And i <> 1 And i <> 37 And i <> 38 And i <> 39 And i <> 40 Then
                    strTemp(i) = "0"
                End If
            Next i
            strTemp(22) = Val("0" & .Fields(27))
            strTemp(23) = Val("0" & .Fields(28))
            strTemp(24) = Val("0" & .Fields(29))
            strTemp(25) = Val("0" & .Fields(30))
            strTemp(26) = Val("0" & .Fields(31))
            If Val(strTemp(2)) + Val(strTemp(3)) + Val(strTemp(4)) + Val(strTemp(5)) = 0 Then
                strTemp(6) = "0"
            Else
                strTemp(6) = Trim(str((Val(strTemp(2)) + Val(strTemp(3))) / (Val(strTemp(2)) + Val(strTemp(3)) + Val(strTemp(4)) + Val(strTemp(5))) * 100))
            End If
            If Val(strTemp(7)) + Val(strTemp(8)) + Val(strTemp(9)) + Val(strTemp(10)) = 0 Then
                strTemp(11) = "0"
            Else
                strTemp(11) = Trim(str((Val(strTemp(8)) + Val(strTemp(7))) / (Val(strTemp(8)) + Val(strTemp(9)) + Val(strTemp(10)) + Val(strTemp(7))) * 100))
            End If
            If Val(strTemp(12)) + Val(strTemp(13)) + Val(strTemp(14)) + Val(strTemp(15)) = 0 Then
                strTemp(16) = "0"
            Else
                strTemp(16) = Trim(str((Val(strTemp(12)) + Val(strTemp(13))) / (Val(strTemp(12)) + Val(strTemp(13)) + Val(strTemp(14)) + Val(strTemp(15))) * 100))
            End If
            If Val(strTemp(17)) + Val(strTemp(18)) + Val(strTemp(19)) + Val(strTemp(20)) = 0 Then
                strTemp(21) = "0"
            Else
                strTemp(21) = Trim(str((Val(strTemp(17)) + Val(strTemp(18))) / (Val(strTemp(17)) + Val(strTemp(18)) + Val(strTemp(19)) + Val(strTemp(20))) * 100))
            End If
            If Val(strTemp(22)) + Val(strTemp(23)) + Val(strTemp(24)) + Val(strTemp(25)) = 0 Then
                strTemp(26) = "0"
            Else
                strTemp(26) = Trim(str((Val(strTemp(22)) + Val(strTemp(23))) / (Val(strTemp(22)) + Val(strTemp(23)) + Val(strTemp(24)) + Val(strTemp(25))) * 100))
            End If
'91/05/21  nick
            If SavDay5 <> strTemp(40) Then
                'Add By Cheng 2002/05/06
                m_strPromoter = ""
                ShowLine
                PrintEnd2 (2)
                ShowLine
                PrintEnd2 (3)
                ShowLine
                PrintEnd2 (0)
                ShowLine
                SavDay5 = strTemp(40)
            End If
            If SavDay1 <> strTemp(38) Then
                PrintEnd2 (1)
                ShowLine
                Page = Page + 1
                Printer.NewPage
                SavDay1 = strTemp(38)
                PrintTitle
                If Val(SavDay1) = 1 Then
                    PrintTitle1
                Else
                    PrintTitle2
                End If
                SavDay2 = strTemp(39)
            End If
            strTemp(0) = StrToStr(strTemp(0), 4)
            strTemp(1) = StrToStr(strTemp(1), 4)
            PrintDatil
            If iPrint >= 13900 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                If Val(SavDay1) = 1 Then
                    PrintTitle1
                Else
                    PrintTitle2
                End If
            End If
            .MoveNext
        Loop
    Else
         GoTo ReadNext3: 'Add By Sindy 2011/3/1
    End If
End With
CheckOC
ShowLine
PrintEnd2 (2)
ShowLine
PrintEnd2 (3)
ShowLine
PrintEnd2 (0)
ShowLine
PrintEnd2 (1)
ShowLine
Page = Page + 1
Printer.NewPage
ReadNext3: 'Add By Sindy 2011/3/1

'列印表(2)格式2
strSql = "select st02,NVL(A0902,A0903),sum(r081003),sum(r081004),sum(r081005),sum(r081006),sum(r081007),sum(r081008),sum(r081009),sum(r081010),sum(r081011)" & _
         ",sum(r081012),sum(r081013),sum(r081014),sum(r081015),sum(r081016),sum(r081017),sum(r081018),sum(r081019),sum(r081020),sum(r081021),sum(r081022)" & _
         ",sum(r081023),SUM(r081024),SUM(R081025),SUM(R081026),SUM(R081027),SUM(R081028),SUM(R081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081033)" & _
         ",SUM(R081034),SUM(R081035),SUm(R081036),SUM(R081037),'',R081039,R081040,r081001,r081002,SUM(R081041),SUM(R081042),SUm(R081043),SUM(R081044),SUM(R081045) from r020409,staff,acc090 where r081001=st01(+) and r081002=a0901(+) and r081039='2' and id='" & strUserNum & "' group by r081001,r081002,R081039,R081040,st02,NVL(A0902,A0903) ORDER BY r081039,r081040, R081001,R081002"
CheckOC
'Add By Cheng 2002/05/06
m_strPromoter = ""
SavDay1 = ""
SavDay2 = ""
'Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        bolNotData = False 'Add By Sindy 2011/3/1
        .MoveFirst
        SavDay1 = CheckStr(.Fields(38))
        SavDay2 = CheckStr(.Fields(39))
        SavDay5 = CheckStr(.Fields(40))
        SavDay6 = CheckStr(.Fields(41))
        PrintTitle
        If Val(SavDay1) = 1 Then
            PrintTitle1_1
        Else
            PrintTitle2_2
        End If
        Do While .EOF = False
            For i = 0 To 41
                strTemp(i) = CheckStr(.Fields(i))
                'Modify By Sindy 2012/6/22 +And i <> 40 員工編號是英數字
                If Val(strTemp(i)) = 0 And i <> 0 And i <> 1 And i <> 37 And i <> 38 And i <> 39 And i <> 40 Then
                    strTemp(i) = "0"
                End If
            Next i
            'edit by nickc 2007/07/27 商標處改格式
            If txt1(3) = "020" And txt1(4) = "020" Then
                strTemp(2) = Val("0" & .Fields(22))
                strTemp(3) = Val("0" & .Fields(23))
                strTemp(4) = Val("0" & .Fields(24))
                strTemp(5) = Val("0" & .Fields(25))
                strTemp(6) = Val("0" & .Fields(26))
                strTemp(7) = Val("0")
                strTemp(8) = Val("0")
                strTemp(9) = Val("0")
                strTemp(10) = Val("0")
                strTemp(11) = Val("0")
            Else
                strTemp(2) = Val("0" & .Fields(32))
                strTemp(3) = Val("0" & .Fields(33))
                strTemp(4) = Val("0" & .Fields(34))
                strTemp(5) = Val("0" & .Fields(35))
                strTemp(6) = Val("0" & .Fields(36))
                strTemp(7) = Val("0" & .Fields(22))
                strTemp(8) = Val("0" & .Fields(23))
                strTemp(9) = Val("0" & .Fields(24))
                strTemp(10) = Val("0" & .Fields(25))
                strTemp(11) = Val("0" & .Fields(26))
            End If
            If Val(strTemp(2)) + Val(strTemp(3)) + Val(strTemp(4)) + Val(strTemp(5)) = 0 Then
                strTemp(6) = "0"
            Else
                strTemp(6) = Trim(str((Val(strTemp(2)) + Val(strTemp(3))) / (Val(strTemp(2)) + Val(strTemp(3)) + Val(strTemp(4)) + Val(strTemp(5))) * 100))
            End If
            If Val(strTemp(7)) + Val(strTemp(8)) + Val(strTemp(9)) + Val(strTemp(10)) = 0 Then
                strTemp(11) = "0"
            Else
                strTemp(11) = Trim(str((Val(strTemp(8)) + Val(strTemp(7))) / (Val(strTemp(8)) + Val(strTemp(9)) + Val(strTemp(10)) + Val(strTemp(7))) * 100))
            End If
'91/05/21  nick
            If SavDay5 <> strTemp(40) Then
                'Add By Cheng 2002/05/06
                m_strPromoter = ""
                ShowLine_2
                PrintEnd2_2 (2)
                ShowLine_2
                PrintEnd2_2 (3)
                ShowLine_2
                PrintEnd2_2 (0)
                ShowLine_2
                SavDay5 = strTemp(40)
            End If
            If SavDay1 <> strTemp(38) Then
                PrintEnd2_2 (1)
                ShowLine_2
                Page = Page + 1
                Printer.NewPage
                SavDay1 = strTemp(38)
                PrintTitle
                If Val(SavDay1) = 1 Then
                    PrintTitle1_1
                Else
                    PrintTitle2_2
                End If
                SavDay2 = strTemp(39)
            End If
            strTemp(0) = StrToStr(strTemp(0), 4)
            strTemp(1) = StrToStr(strTemp(1), 4)
            PrintDatil_2
            If iPrint >= 13900 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                If Val(SavDay1) = 1 Then
                    PrintTitle1_1
                Else
                    PrintTitle2_2
                End If
            End If
            .MoveNext
        Loop
    Else
         GoTo ReadNext4: 'Add By Sindy 2011/3/1
    End If
End With
CheckOC
ShowLine_2
PrintEnd2_2 (2)
ShowLine_2
PrintEnd2_2 (3)
ShowLine_2
PrintEnd2_2 (0)
ShowLine_2
PrintEnd2_2 (1)
ShowLine_2
Page = Page + 1
ReadNext4: 'Add By Sindy 2011/3/1
If bolNotData = False Then Printer.EndDoc
End Sub

Sub PrintEnd(Strindex As Integer)
Select Case Strindex
Case 0
     strSql = "select '個人小計','',SUM(r081003),sum(r081004),sum(r081005),sum(r081006),sum(r081007),sum(r081008),sum(r081009),sum(r081010),sum(r081011),sum(r081012),sum(r081013),sum(r081014),sum(r081015),sum(r081016),sum(r081017),sum(r081018),sum(r081019),sum(r081020),sum(r081021),sum(r081022),sum(r081023),SUM(r081024),SUM(R081025),SUM(R081026),SUM(R081027),SUM(R081028),SUM(R081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081033),SUM(R081034),SUM(R081035),SUm(R081036),SUM(R081037) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and r081040='" & SavDay5 & "' "
Case 1
     strSql = "select '全所總計','',SUM(r081003),sum(r081004),sum(r081005),sum(r081006),sum(r081007),sum(r081008),sum(r081009),sum(r081010),sum(r081011),sum(r081012),sum(r081013),sum(r081014),sum(r081015),sum(r081016),sum(r081017),sum(r081018),sum(r081019),sum(r081020),sum(r081021),sum(r081022),sum(r081023),SUM(r081024),SUM(R081025),SUM(R081026),SUM(R081027),SUM(R081028),SUM(R081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081033),SUM(R081034),SUM(R081035),SUm(R081036),SUM(R081037) from r020409 where id='" & strUserNum & "' and r081039=" & Val(SavDay1) & " "
Case 2
     'edit by nickc 2005/05/12
     'StrSql = "select '國內小計','',SUM(r081003),sum(r081004),sum(r081005),sum(r081006),sum(r081007),sum(r081008),sum(r081009),sum(r081010),sum(r081011),sum(r081012),sum(r081013),sum(r081014),sum(r081015),sum(r081016),sum(r081017),sum(r081018),sum(r081019),sum(r081020),sum(r081021),sum(r081022),sum(r081023),SUM(r081024),SUM(R081025),SUM(R081026),SUM(R081027),SUM(R081028),SUM(R081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081033),SUM(R081034),SUM(R081035),SUm(R081036),SUM(R081037) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and r081038='*' and r081040='" & SavDay5 & "' "
     strSql = "select '國內業務小計','',SUM(r081003),sum(r081004),sum(r081005),sum(r081006),sum(r081007),sum(r081008),sum(r081009),sum(r081010),sum(r081011),sum(r081012),sum(r081013),sum(r081014),sum(r081015),sum(r081016),sum(r081017),sum(r081018),sum(r081019),sum(r081020),sum(r081021),sum(r081022),sum(r081023),SUM(r081024),SUM(R081025),SUM(R081026),SUM(R081027),SUM(R081028),SUM(R081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081033),SUM(R081034),SUM(R081035),SUm(R081036),SUM(R081037) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and r081038='*' and r081040='" & SavDay5 & "' "
Case 3
     'edit by nickc 2005/05/12
     'StrSql = "select '國外小計','',SUM(r081003),sum(r081004),sum(r081005),sum(r081006),sum(r081007),sum(r081008),sum(r081009),sum(r081010),sum(r081011),sum(r081012),sum(r081013),sum(r081014),sum(r081015),sum(r081016),sum(r081017),sum(r081018),sum(r081019),sum(r081020),sum(r081021),sum(r081022),sum(r081023),SUM(r081024),SUM(R081025),SUM(R081026),SUM(R081027),SUM(R081028),SUM(R081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081033),SUM(R081034),SUM(R081035),SUm(R081036),SUM(R081037) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and (r081038='' or r081038 is null ) and r081040='" & SavDay5 & "' "
     strSql = "select '國外業務小計','',SUM(r081003),sum(r081004),sum(r081005),sum(r081006),sum(r081007),sum(r081008),sum(r081009),sum(r081010),sum(r081011),sum(r081012),sum(r081013),sum(r081014),sum(r081015),sum(r081016),sum(r081017),sum(r081018),sum(r081019),sum(r081020),sum(r081021),sum(r081022),sum(r081023),SUM(r081024),SUM(R081025),SUM(R081026),SUM(R081027),SUM(R081028),SUM(R081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081033),SUM(R081034),SUM(R081035),SUm(R081036),SUM(R081037) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and (r081038='' or r081038 is null ) and r081040='" & SavDay5 & "' "
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
            For i = 2 To 32 Step 5
                Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
            Next i
            For i = 0 To 36
                StrTemp7(i) = CheckStr(.Fields(i))
                If Val(StrTemp7(i)) = 0 And i <> 0 And i <> 1 And i <> 37 And i <> 39 Then
                    StrTemp7(i) = "0"
                End If
            Next i
            If Val(StrTemp7(2)) + Val(StrTemp7(3)) + Val(StrTemp7(4)) + Val(StrTemp7(5)) = 0 Then
                StrTemp7(6) = "0"
            Else
                StrTemp7(6) = Trim(str((Val(StrTemp7(2)) + Val(StrTemp7(3))) / (Val(StrTemp7(2)) + Val(StrTemp7(3)) + Val(StrTemp7(4)) + Val(StrTemp7(5))) * 100))
            End If
            If Val(StrTemp7(7)) + Val(StrTemp7(8)) + Val(StrTemp7(9)) + Val(StrTemp7(10)) = 0 Then
                StrTemp7(11) = "0"
            Else
                StrTemp7(11) = Trim(str((Val(StrTemp7(8)) + Val(StrTemp7(7))) / (Val(StrTemp7(8)) + Val(StrTemp7(9)) + Val(StrTemp7(10)) + Val(StrTemp7(7))) * 100))
            End If
            If Val(StrTemp7(12)) + Val(StrTemp7(13)) + Val(StrTemp7(14)) + Val(StrTemp7(15)) = 0 Then
                StrTemp7(16) = "0"
            Else
                StrTemp7(16) = Trim(str((Val(StrTemp7(12)) + Val(StrTemp7(13))) / (Val(StrTemp7(12)) + Val(StrTemp7(13)) + Val(StrTemp7(14)) + Val(StrTemp7(15))) * 100))
            End If
            If Val(StrTemp7(17)) + Val(StrTemp7(18)) + Val(StrTemp7(19)) + Val(StrTemp7(20)) = 0 Then
                StrTemp7(21) = "0"
            Else
                StrTemp7(21) = Trim(str((Val(StrTemp7(17)) + Val(StrTemp7(18))) / (Val(StrTemp7(17)) + Val(StrTemp7(18)) + Val(StrTemp7(19)) + Val(StrTemp7(20))) * 100))
            End If
            If Val(StrTemp7(22)) + Val(StrTemp7(23)) + Val(StrTemp7(24)) + Val(StrTemp7(25)) = 0 Then
                StrTemp7(26) = "0"
            Else
                StrTemp7(26) = Trim(str((Val(StrTemp7(22)) + Val(StrTemp7(23))) / (Val(StrTemp7(22)) + Val(StrTemp7(23)) + Val(StrTemp7(24)) + Val(StrTemp7(25))) * 100))
            End If
            If Val(StrTemp7(27)) + Val(StrTemp7(28)) + Val(StrTemp7(29)) + Val(StrTemp7(30)) = 0 Then
                StrTemp7(31) = "0"
            Else
                StrTemp7(31) = Trim(str((Val(StrTemp7(27)) + Val(StrTemp7(28))) / (Val(StrTemp7(27)) + Val(StrTemp7(28)) + Val(StrTemp7(29)) + Val(StrTemp7(30))) * 100))
            End If
            If Val(StrTemp7(32)) + Val(StrTemp7(33)) + Val(StrTemp7(34)) + Val(StrTemp7(35)) = 0 Then
                StrTemp7(36) = "0"
            Else
                StrTemp7(36) = Trim(str((Val(StrTemp7(32)) + Val(StrTemp7(33))) / (Val(StrTemp7(32)) + Val(StrTemp7(33)) + Val(StrTemp7(34)) + Val(StrTemp7(35))) * 100))
            End If
            '2012/4/6 MODIFY BY SONIA 只做台灣案的預估準確率
            'If TestOk = True And Strindex = 0 And Val(SavDay1) = 2 Then
            If TestOk = True And Strindex = 0 And Val(SavDay1) = 2 And txt1(3) = "000" And txt1(4) = "000" Then
               'Modify By Cheng 2002/05/06
'                cnnConnection.Execute "update performance set pe17=" & Val(StrTemp7(26)) & " where pe01='" & CheckStr(.Fields(39)) & "' and pe02='T' and pe03=" & Val(Format(ChangeTStringToTDateString(txt1(1)), "mm")) & " "
                cnnConnection.Execute "update performance set pe17=" & Val(StrTemp7(26)) & " where pe01='" & CheckStr(SavDay2) & "' and pe02='T' and pe03=" & Val(Format(ChangeTStringToTDateString(txt1(1)), "mm")) & " "
            End If
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(1)
            For i = 2 To 36
                'edit by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" And i > 31 Then
                Else
                        Select Case i
                        Case 6, 11, 16, 21, 26
                            Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(StrTemp7(i), "##0.00"))
                            Printer.CurrentY = iPrint
                            Printer.Print Format(StrTemp7(i), "##0.00") & "%"
                        Case 31, 36
                            If Val(SavDay1) = 1 Then
                                Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(StrTemp7(i), "##0.00"))
                                Printer.CurrentY = iPrint
                                Printer.Print Format(StrTemp7(i), "##0.00") & "%"
                            End If
                        Case 27, 28, 29, 30, 32, 33, 34, 35
                            If Val(SavDay1) = 1 Then
                                Printer.CurrentX = PLeft(i) + 200 - Printer.TextWidth(Format(StrTemp7(i), "####0"))
                                Printer.CurrentY = iPrint
                                Printer.Print Format(StrTemp7(i), "####0")
                            End If
                        Case Else
                            Printer.CurrentX = PLeft(i) + 200 - Printer.TextWidth(Format(StrTemp7(i), "####0"))
                            Printer.CurrentY = iPrint
                            Printer.Print Format(StrTemp7(i), "####0")
                        End Select
                End If
            Next i
            iPrint = iPrint + 300
            If iPrint >= 13900 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                If Val(SavDay1) = 1 Then
                    PrintTitle1
                Else
                    PrintTitle2
                End If
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

Sub PrintEnd_1(Strindex As Integer)
Select Case Strindex
Case 0
     strSql = "select '個人小計','',SUM(r081041),sum(r081042),sum(r081043),sum(r081044),sum(r081045) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and r081040='" & SavDay5 & "' "
Case 1
     strSql = "select '全所總計','',SUM(r081041),sum(r081042),sum(r081043),sum(r081044),sum(r081045) from r020409 where id='" & strUserNum & "' and r081039=" & Val(SavDay1) & " "
Case 2
     'edit by nickc 2005/05/12
     'StrSql = "select '國內小計','',SUM(r081041),sum(r081042),sum(r081043),sum(r081044),sum(r081045) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and r081038='*' and r081040='" & SavDay5 & "' "
     strSql = "select '國內業務小計','',SUM(r081041),sum(r081042),sum(r081043),sum(r081044),sum(r081045) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and r081038='*' and r081040='" & SavDay5 & "' "
Case 3
     'edit by nickc 2005/05/12
     'StrSql = "select '國外小計','',SUM(r081041),sum(r081042),sum(r081043),sum(r081044),sum(r081045) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and (r081038='' or r081038 is null ) and r081040='" & SavDay5 & "' "
     strSql = "select '國外業務小計','',SUM(r081041),sum(r081042),sum(r081043),sum(r081044),sum(r081045) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and (r081038='' or r081038 is null ) and r081040='" & SavDay5 & "' "
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
            For i = 2 To 6 Step 5
                Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
            Next i
            For i = 0 To 6
                StrTemp7(i) = CheckStr(.Fields(i))
                If Val(StrTemp7(i)) = 0 And i <> 0 And i <> 1 And i <> 37 And i <> 39 Then
                    StrTemp7(i) = "0"
                End If
            Next i
            If Val(StrTemp7(2)) + Val(StrTemp7(3)) + Val(StrTemp7(4)) + Val(StrTemp7(5)) = 0 Then
                StrTemp7(6) = "0"
            Else
                StrTemp7(6) = Trim(str((Val(StrTemp7(2)) + Val(StrTemp7(3))) / (Val(StrTemp7(2)) + Val(StrTemp7(3)) + Val(StrTemp7(4)) + Val(StrTemp7(5))) * 100))
            End If
'            If TestOk = True And Strindex = 0 And Val(SavDay1) = 2 Then
               'Modify By Cheng 2002/05/06
'                cnnConnection.Execute "update performance set pe17=" & Val(StrTemp7(26)) & " where pe01='" & CheckStr(.Fields(39)) & "' and pe02='T' and pe03=" & Val(Format(ChangeTStringToTDateString(txt1(1)), "mm")) & " "
'                cnnConnection.Execute "update performance set pe17=" & Val(StrTemp7(26)) & " where pe01='" & CheckStr(SavDay2) & "' and pe02='T' and pe03=" & Val(Format(ChangeTStringToTDateString(txt1(1)), "mm")) & " "
'            End If
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(1)
            For i = 2 To 6
                Select Case i
                Case 6, 11, 16, 21, 26
                    Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(StrTemp7(i), "##0.00"))
                    Printer.CurrentY = iPrint
                    Printer.Print Format(StrTemp7(i), "##0.00") & "%"
                Case 31, 36
                    If Val(SavDay1) = 1 Then
                        Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(StrTemp7(i), "##0.00"))
                        Printer.CurrentY = iPrint
                        Printer.Print Format(StrTemp7(i), "##0.00") & "%"
                    End If
                Case 27, 28, 29, 30, 32, 33, 34, 35
                    If Val(SavDay1) = 1 Then
                        Printer.CurrentX = PLeft(i) + 200 - Printer.TextWidth(Format(StrTemp7(i), "####0"))
                        Printer.CurrentY = iPrint
                        Printer.Print Format(StrTemp7(i), "####0")
                    End If
                Case Else
                    Printer.CurrentX = PLeft(i) + 200 - Printer.TextWidth(Format(StrTemp7(i), "####0"))
                    Printer.CurrentY = iPrint
                    Printer.Print Format(StrTemp7(i), "####0")
                End Select
            Next i
            iPrint = iPrint + 300
            If iPrint >= 13900 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                If Val(SavDay1) = 1 Then
                    PrintTitle1_1
                Else
                    PrintTitle2_2
                End If
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

'Add By Cheng 2003/03/11
Sub PrintEnd2(Strindex As Integer)
Select Case Strindex
Case 0
     strSql = "select '個人小計','',SUM(r081003),sum(r081004),sum(r081005),sum(r081006),sum(r081007),sum(r081008),sum(r081009),sum(r081010),sum(r081011),sum(r081012),sum(r081013),sum(r081014),sum(r081015),sum(r081016),sum(r081017),sum(r081018),sum(r081019),sum(r081020),sum(r081021),sum(r081022),sum(r081028),SUM(r081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081028),SUM(R081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081033),SUM(R081034),SUM(R081035),SUm(R081036),SUM(R081037) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and r081040='" & SavDay5 & "' "
Case 1
     strSql = "select '全所總計','',SUM(r081003),sum(r081004),sum(r081005),sum(r081006),sum(r081007),sum(r081008),sum(r081009),sum(r081010),sum(r081011),sum(r081012),sum(r081013),sum(r081014),sum(r081015),sum(r081016),sum(r081017),sum(r081018),sum(r081019),sum(r081020),sum(r081021),sum(r081022),sum(r081028),SUM(r081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081028),SUM(R081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081033),SUM(R081034),SUM(R081035),SUm(R081036),SUM(R081037) from r020409 where id='" & strUserNum & "' and r081039=" & Val(SavDay1) & " "
Case 2
     'edit by nickc 2005/05/12
     'StrSql = "select '國內小計','',SUM(r081003),sum(r081004),sum(r081005),sum(r081006),sum(r081007),sum(r081008),sum(r081009),sum(r081010),sum(r081011),sum(r081012),sum(r081013),sum(r081014),sum(r081015),sum(r081016),sum(r081017),sum(r081018),sum(r081019),sum(r081020),sum(r081021),sum(r081022),sum(r081028),SUM(r081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081028),SUM(R081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081033),SUM(R081034),SUM(R081035),SUm(R081036),SUM(R081037) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and r081038='*' and r081040='" & SavDay5 & "' "
     strSql = "select '國內業務小計','',SUM(r081003),sum(r081004),sum(r081005),sum(r081006),sum(r081007),sum(r081008),sum(r081009),sum(r081010),sum(r081011),sum(r081012),sum(r081013),sum(r081014),sum(r081015),sum(r081016),sum(r081017),sum(r081018),sum(r081019),sum(r081020),sum(r081021),sum(r081022),sum(r081028),SUM(r081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081028),SUM(R081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081033),SUM(R081034),SUM(R081035),SUm(R081036),SUM(R081037) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and r081038='*' and r081040='" & SavDay5 & "' "
Case 3
     'edit by nickc 2005/05/12
     'StrSql = "select '國外小計','',SUM(r081003),sum(r081004),sum(r081005),sum(r081006),sum(r081007),sum(r081008),sum(r081009),sum(r081010),sum(r081011),sum(r081012),sum(r081013),sum(r081014),sum(r081015),sum(r081016),sum(r081017),sum(r081018),sum(r081019),sum(r081020),sum(r081021),sum(r081022),sum(r081028),SUM(r081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081028),SUM(R081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081033),SUM(R081034),SUM(R081035),SUm(R081036),SUM(R081037) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and (r081038='' or r081038 is null ) and r081040='" & SavDay5 & "' "
     strSql = "select '國外業務小計','',SUM(r081003),sum(r081004),sum(r081005),sum(r081006),sum(r081007),sum(r081008),sum(r081009),sum(r081010),sum(r081011),sum(r081012),sum(r081013),sum(r081014),sum(r081015),sum(r081016),sum(r081017),sum(r081018),sum(r081019),sum(r081020),sum(r081021),sum(r081022),sum(r081028),SUM(r081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081028),SUM(R081029),SUM(R081030),SUM(R081031),SUM(R081032),SUM(R081033),SUM(R081034),SUM(R081035),SUm(R081036),SUM(R081037) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and (r081038='' or r081038 is null ) and r081040='" & SavDay5 & "' "
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
            For i = 2 To 32 Step 5
                Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
            Next i
            For i = 0 To 36
                StrTemp7(i) = CheckStr(.Fields(i))
                If Val(StrTemp7(i)) = 0 And i <> 0 And i <> 1 And i <> 37 And i <> 39 Then
                    StrTemp7(i) = "0"
                End If
            Next i
            If Val(StrTemp7(2)) + Val(StrTemp7(3)) + Val(StrTemp7(4)) + Val(StrTemp7(5)) = 0 Then
                StrTemp7(6) = "0"
            Else
                StrTemp7(6) = Trim(str((Val(StrTemp7(2)) + Val(StrTemp7(3))) / (Val(StrTemp7(2)) + Val(StrTemp7(3)) + Val(StrTemp7(4)) + Val(StrTemp7(5))) * 100))
            End If
            If Val(StrTemp7(7)) + Val(StrTemp7(8)) + Val(StrTemp7(9)) + Val(StrTemp7(10)) = 0 Then
                StrTemp7(11) = "0"
            Else
                StrTemp7(11) = Trim(str((Val(StrTemp7(8)) + Val(StrTemp7(7))) / (Val(StrTemp7(8)) + Val(StrTemp7(9)) + Val(StrTemp7(10)) + Val(StrTemp7(7))) * 100))
            End If
            If Val(StrTemp7(12)) + Val(StrTemp7(13)) + Val(StrTemp7(14)) + Val(StrTemp7(15)) = 0 Then
                StrTemp7(16) = "0"
            Else
                StrTemp7(16) = Trim(str((Val(StrTemp7(12)) + Val(StrTemp7(13))) / (Val(StrTemp7(12)) + Val(StrTemp7(13)) + Val(StrTemp7(14)) + Val(StrTemp7(15))) * 100))
            End If
            If Val(StrTemp7(17)) + Val(StrTemp7(18)) + Val(StrTemp7(19)) + Val(StrTemp7(20)) = 0 Then
                StrTemp7(21) = "0"
            Else
                StrTemp7(21) = Trim(str((Val(StrTemp7(17)) + Val(StrTemp7(18))) / (Val(StrTemp7(17)) + Val(StrTemp7(18)) + Val(StrTemp7(19)) + Val(StrTemp7(20))) * 100))
            End If
            If Val(StrTemp7(22)) + Val(StrTemp7(23)) + Val(StrTemp7(24)) + Val(StrTemp7(25)) = 0 Then
                StrTemp7(26) = "0"
            Else
                StrTemp7(26) = Trim(str((Val(StrTemp7(22)) + Val(StrTemp7(23))) / (Val(StrTemp7(22)) + Val(StrTemp7(23)) + Val(StrTemp7(24)) + Val(StrTemp7(25))) * 100))
            End If
'            If TestOk = True And Strindex = 0 And Val(SavDay1) = 2 Then
               'Modify By Cheng 2002/05/06
'                cnnConnection.Execute "update performance set pe17=" & Val(StrTemp7(26)) & " where pe01='" & CheckStr(.Fields(39)) & "' and pe02='T' and pe03=" & Val(Format(ChangeTStringToTDateString(txt1(1)), "mm")) & " "
'                cnnConnection.Execute "update performance set pe17=" & Val(StrTemp7(26)) & " where pe01='" & CheckStr(SavDay2) & "' and pe02='T' and pe03=" & Val(Format(ChangeTStringToTDateString(txt1(1)), "mm")) & " "
'            End If
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(1)
            For i = 2 To 36
                Select Case i
                Case 6, 11, 16, 21, 26
                    Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(StrTemp7(i), "##0.00"))
                    Printer.CurrentY = iPrint
                    Printer.Print Format(StrTemp7(i), "##0.00") & "%"
                Case 31, 36
                    If Val(SavDay1) = 1 Then
                        Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(StrTemp7(i), "##0.00"))
                        Printer.CurrentY = iPrint
                        Printer.Print Format(StrTemp7(i), "##0.00") & "%"
                    End If
                Case 27, 28, 29, 30, 32, 33, 34, 35
                    If Val(SavDay1) = 1 Then
                        Printer.CurrentX = PLeft(i) + 200 - Printer.TextWidth(Format(StrTemp7(i), "####0"))
                        Printer.CurrentY = iPrint
                        Printer.Print Format(StrTemp7(i), "####0")
                    End If
                Case Else
                    Printer.CurrentX = PLeft(i) + 200 - Printer.TextWidth(Format(StrTemp7(i), "####0"))
                    Printer.CurrentY = iPrint
                    Printer.Print Format(StrTemp7(i), "####0")
                End Select
            Next i
            iPrint = iPrint + 300
            If iPrint >= 13900 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                If Val(SavDay1) = 1 Then
                    PrintTitle1
                Else
                    PrintTitle2
                End If
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

'Add By Cheng 2003/03/11
Sub PrintEnd2_2(Strindex As Integer)
Select Case Strindex
Case 0
    'edit by nickc 2007/07/27 商標處改格式
    If txt1(3) = "020" And txt1(4) = "020" Then
        strSql = "select '個人小計','',sum(r081023),sum(r081024),sum(r081025),sum(r081026),sum(r081027),'','','','','' from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and r081040='" & SavDay5 & "' "
    Else
        strSql = "select '個人小計','',SUM(r081033),sum(r081034),sum(r081035),sum(r081036),sum(r081037),sum(r081023),sum(r081024),sum(r081025),sum(r081026),sum(r081027) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and r081040='" & SavDay5 & "' "
    End If
Case 1
    'edit by nickc 2007/07/27 商標處改格式
    If txt1(3) = "020" And txt1(4) = "020" Then
        strSql = "select '全所總計','',sum(r081023),sum(r081024),sum(r081025),sum(r081026),sum(r081027),'','','','','' from r020409 where id='" & strUserNum & "' and r081039=" & Val(SavDay1) & " "
    Else
        strSql = "select '全所總計','',SUM(r081033),sum(r081034),sum(r081035),sum(r081036),sum(r081037),sum(r081023),sum(r081024),sum(r081025),sum(r081026),sum(r081027) from r020409 where id='" & strUserNum & "' and r081039=" & Val(SavDay1) & " "
    End If
Case 2
     'edit by nickc 2005/05/12
     'StrSql = "select '國內小計','',SUM(r081033),sum(r081034),sum(r081035),sum(r081036),sum(r081037),sum(r081023),sum(r081024),sum(r081025),sum(r081026),sum(r081027) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and r081038='*' and r081040='" & SavDay5 & "' "
    'edit by nickc 2007/07/27 商標處改格式
    If txt1(3) = "020" And txt1(4) = "020" Then
        strSql = "select '國內業務小計','',sum(r081023),sum(r081024),sum(r081025),sum(r081026),sum(r081027),'','','','','' from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and r081038='*' and r081040='" & SavDay5 & "' "
    Else
        strSql = "select '國內業務小計','',SUM(r081033),sum(r081034),sum(r081035),sum(r081036),sum(r081037),sum(r081023),sum(r081024),sum(r081025),sum(r081026),sum(r081027) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and r081038='*' and r081040='" & SavDay5 & "' "
    End If
Case 3
     'edit by nickc 2005/05/12
     'StrSql = "select '國外小計','',SUM(r081033),sum(r081034),sum(r081035),sum(r081036),sum(r081037),sum(r081023),sum(r081024),sum(r081025),sum(r081026),sum(r081027) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and (r081038='' or r081038 is null ) and r081040='" & SavDay5 & "' "
         'edit by nickc 2007/07/27 商標處改格式
    If txt1(3) = "020" And txt1(4) = "020" Then
        strSql = "select '國外業務小計','',sum(r081023),sum(r081024),sum(r081025),sum(r081026),sum(r081027),'','','','','' from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and (r081038='' or r081038 is null ) and r081040='" & SavDay5 & "' "
    Else
        strSql = "select '國外業務小計','',SUM(r081033),sum(r081034),sum(r081035),sum(r081036),sum(r081037),sum(r081023),sum(r081024),sum(r081025),sum(r081026),sum(r081027) from r020409 where id='" & strUserNum & "' AND r081039=" & Val(SavDay1) & " and (r081038='' or r081038 is null ) and r081040='" & SavDay5 & "' "
    End If
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
            For i = 2 To 11 Step 5
                Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
            Next i
            For i = 0 To 11
                StrTemp7(i) = CheckStr(.Fields(i))
                If Val(StrTemp7(i)) = 0 And i <> 0 And i <> 1 And i <> 37 And i <> 39 Then
                    StrTemp7(i) = "0"
                End If
            Next i
            If Val(StrTemp7(2)) + Val(StrTemp7(3)) + Val(StrTemp7(4)) + Val(StrTemp7(5)) = 0 Then
                StrTemp7(6) = "0"
            Else
                StrTemp7(6) = Trim(str((Val(StrTemp7(2)) + Val(StrTemp7(3))) / (Val(StrTemp7(2)) + Val(StrTemp7(3)) + Val(StrTemp7(4)) + Val(StrTemp7(5))) * 100))
            End If
            If Val(StrTemp7(7)) + Val(StrTemp7(8)) + Val(StrTemp7(9)) + Val(StrTemp7(10)) = 0 Then
                StrTemp7(11) = "0"
            Else
                StrTemp7(11) = Trim(str((Val(StrTemp7(8)) + Val(StrTemp7(7))) / (Val(StrTemp7(8)) + Val(StrTemp7(9)) + Val(StrTemp7(10)) + Val(StrTemp7(7))) * 100))
            End If
            '2012/4/6 MODIFY BY SONIA 只做台灣案的預估準確率
            'If TestOk = True And Strindex = 0 And Val(SavDay1) = 2 Then
            If TestOk = True And Strindex = 0 And Val(SavDay1) = 2 And txt1(3) = "000" And txt1(4) = "000" Then
               'Modify By Cheng 2002/05/06
'                cnnConnection.Execute "update performance set pe17=" & Val(StrTemp7(26)) & " where pe01='" & CheckStr(.Fields(39)) & "' and pe02='T' and pe03=" & Val(Format(ChangeTStringToTDateString(txt1(1)), "mm")) & " "
                'Modify By Cheng 2003/03/26
'                cnnConnection.Execute "update performance set pe17=" & Val(StrTemp7(11)) & " where pe01='" & CheckStr(SavDay2) & "' and pe02='T' and pe03=" & Mid(txt1(1) + 19110000, 1, 6) & " "
                cnnConnection.Execute "update performance set pe17=" & Val(StrTemp7(11)) & " where pe01='" & CheckStr(SavDay5) & "' and pe02='T' and pe03=" & Mid(txt1(1) + 19110000, 1, 6) & " "
            End If
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(1)
            For i = 2 To 11
                'edit by nickc 2007/07/27 商標處改格式
                If txt1(3) = "020" And txt1(4) = "020" And i > 6 Then
                Else
                    Select Case i
                    Case 6, 11, 16, 21, 26
                        Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(StrTemp7(i), "##0.00"))
                        Printer.CurrentY = iPrint
                        Printer.Print Format(StrTemp7(i), "##0.00") & "%"
                    Case 31, 36
                        If Val(SavDay1) = 1 Then
                            Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(StrTemp7(i), "##0.00"))
                            Printer.CurrentY = iPrint
                            Printer.Print Format(StrTemp7(i), "##0.00") & "%"
                        End If
                    Case 27, 28, 29, 30, 32, 33, 34, 35
                        If Val(SavDay1) = 1 Then
                            Printer.CurrentX = PLeft(i) + 200 - Printer.TextWidth(Format(StrTemp7(i), "####0"))
                            Printer.CurrentY = iPrint
                            Printer.Print Format(StrTemp7(i), "####0")
                        End If
                    Case Else
                        Printer.CurrentX = PLeft(i) + 200 - Printer.TextWidth(Format(StrTemp7(i), "####0"))
                        Printer.CurrentY = iPrint
                        Printer.Print Format(StrTemp7(i), "####0")
                    End Select
                End If
            Next i
            iPrint = iPrint + 300
            If iPrint >= 13900 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                If Val(SavDay1) = 1 Then
                    PrintTitle1_1
                Else
                    PrintTitle2_2
                End If
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

Sub PrintTitle()
iPrint = 0
'Printer.Orientation = 1 'Removed by Morgan 2015/6/3
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "商爭案承辦人預估勝敗統計表(" & SavDay1 & ") "
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 7000
Printer.CurrentY = iPrint
Printer.Print "勝敗日期：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 16800
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300

'Add By Cheng 2002/02/07
Printer.CurrentX = 250
Printer.CurrentY = iPrint
Printer.Print "申請國家：" & Me.txt1(3).Text & " － " & Me.txt1(4).Text
Printer.CurrentX = 6750
Printer.CurrentY = iPrint
Printer.Print "系統類別：" & Me.txt1(0).Text

Printer.CurrentX = 16800
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
End Sub

Sub PrintTitle2()
GetPleft
Printer.Font.Size = 10
'Printer.CurrentX = 0
'Printer.CurrentY = iPrint
'Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
ShowLine
iPrint = iPrint - 300
For i = 2 To 32 Step 5
    Printer.Line (PLeft(i) - 50, iPrint + 150)-(PLeft(i) - 50, iPrint + 1750)
Next i
iPrint = iPrint + 300
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle2
    Exit Sub
End If
'add by nickc 2007/07/27 商標處改格式
If txt1(3) = "020" And txt1(4) = "020" Then
    Printer.CurrentX = PLeft1(1) - (Printer.TextWidth(Title_602) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_602 'Modify By Sindy 2015/2/4"異議答辯"
    Printer.CurrentX = PLeft1(2) - (Printer.TextWidth(Title_604) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_604 'Modify By Sindy 2015/2/4"裁定答辯"
    Printer.CurrentX = PLeft1(3) - (Printer.TextWidth(Title_606) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_606 'Modify By Sindy 2015/2/4"撤銷答辯"
    Printer.CurrentX = PLeft1(4) - (Printer.TextWidth("") / 2)
    Printer.CurrentY = iPrint
    Printer.Print "" 'Modify By Sindy 2015/2/4"註冊不當撤銷答辯"
    Printer.CurrentX = PLeft1(5) - (Printer.TextWidth(Title_406) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_406 'Modify By Sindy 2015/2/4"復審答辯"
Else
    Printer.CurrentX = PLeft1(1) - (Printer.TextWidth(Title_602) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_602 'Modify By Sindy 2015/2/4"異議答辯"
    Printer.CurrentX = PLeft1(2) - (Printer.TextWidth(Title_604) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_604 'Modify By Sindy 2015/2/4"評定答辯"
    Printer.CurrentX = PLeft1(3) - (Printer.TextWidth(Title_606) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_606 'Modify By Sindy 2015/2/4"廢止答辯"
    Printer.CurrentX = PLeft1(4) - (Printer.TextWidth(Title_406) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_406 'Modify By Sindy 2015/2/4"參加訴願"
    Printer.CurrentX = PLeft1(5) - (Printer.TextWidth(Title_407) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_407 'Modify By Sindy 2015/2/4"參加訴訟"
End If
iPrint = iPrint + 300
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle2
    Exit Sub
End If
Printer.Font.Underline = True
For k = 2 To 22 Step 5
    Printer.CurrentX = PLeft(k)
    Printer.CurrentY = iPrint
    Printer.Print "正    確"
    Printer.CurrentX = PLeft(k + 2)
    Printer.CurrentY = iPrint
    Printer.Print "錯    誤"
    Printer.Font.Underline = False
    Printer.CurrentX = PLeft(k + 4) - 50
    Printer.CurrentY = iPrint
    Printer.Print "準確率"
    Printer.Font.Underline = True
Next k
Printer.Font.Underline = False
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "業務區別"
For k = 2 To 22 Step 5
    Printer.CurrentX = PLeft(k)
    Printer.CurrentY = iPrint
    Printer.Print "勝"
    Printer.CurrentX = PLeft(k + 1)
    Printer.CurrentY = iPrint
    Printer.Print "敗"
    Printer.CurrentX = PLeft(k + 2)
    Printer.CurrentY = iPrint
    Printer.Print "勝"
    Printer.CurrentX = PLeft(k + 3)
    Printer.CurrentY = iPrint
    Printer.Print "敗"
Next k

iPrint = iPrint + 300
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle2
    Exit Sub
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle2
    Exit Sub
End If
End Sub

'Add By Cheng 2003/03/11
Sub PrintTitle2_2()
GetPleft
Printer.Font.Size = 10
'Printer.CurrentX = 0
'Printer.CurrentY = iPrint
'Printer.Line (0, iPrint + 150)-(19000 / 11 * 5, iPrint + 150)
ShowLine_2
iPrint = iPrint - 300
For i = 2 To 11 Step 5
    Printer.Line (PLeft(i) - 50, iPrint + 150)-(PLeft(i) - 50, iPrint + 1750)
Next i
iPrint = iPrint + 300
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle2_2
    Exit Sub
End If
'edit by nickc 2007/07/27 商標處改格式
If txt1(3) = "020" And txt1(4) = "020" Then
    Printer.CurrentX = PLeft1(1) - (Printer.TextWidth("總計") / 2)
    Printer.CurrentY = iPrint
    Printer.Print "總計"
Else
    Printer.CurrentX = PLeft1(1) - (Printer.TextWidth("上訴答辯") / 2)
    Printer.CurrentY = iPrint
    Printer.Print "上訴答辯"
    Printer.CurrentX = PLeft1(2) - (Printer.TextWidth("總計") / 2)
    Printer.CurrentY = iPrint
    Printer.Print "總計"
End If
iPrint = iPrint + 300
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle2
    Exit Sub
End If
Printer.Font.Underline = True
For k = 2 To 11 Step 5
    'edit by nickc 2007/07/27 商標處改格式
    If txt1(3) = "020" And txt1(4) = "020" And k > 6 Then
    Else
            Printer.CurrentX = PLeft(k)
            Printer.CurrentY = iPrint
            Printer.Print "正    確"
            Printer.CurrentX = PLeft(k + 2)
            Printer.CurrentY = iPrint
            Printer.Print "錯    誤"
            Printer.Font.Underline = False
            Printer.CurrentX = PLeft(k + 4) - 50
            Printer.CurrentY = iPrint
            Printer.Print "準確率"
            Printer.Font.Underline = True
    End If
Next k
Printer.Font.Underline = False
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "業務區別"
For k = 2 To 11 Step 5
    'edit by nickc 2007/07/27 商標處改格式
    If txt1(3) = "020" And txt1(4) = "020" And k > 6 Then
    Else
        Printer.CurrentX = PLeft(k)
        Printer.CurrentY = iPrint
        Printer.Print "勝"
        Printer.CurrentX = PLeft(k + 1)
        Printer.CurrentY = iPrint
        Printer.Print "敗"
        Printer.CurrentX = PLeft(k + 2)
        Printer.CurrentY = iPrint
        Printer.Print "勝"
        Printer.CurrentX = PLeft(k + 3)
        Printer.CurrentY = iPrint
        Printer.Print "敗"
    End If
Next k

iPrint = iPrint + 300
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle2_2
    Exit Sub
End If
'Printer.CurrentX = 0
'Printer.CurrentY = iPrint
'Printer.Line (0, iPrint + 150)-(19000 / 11 * 5, iPrint + 150)
'iPrint = iPrint + 300
ShowLine_2
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle2_2
    Exit Sub
End If
End Sub

Sub PrintTitle1()
GetPleft
Printer.Font.Size = 8
'Printer.CurrentX = 0
'Printer.CurrentY = iPrint
'Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
ShowLine
iPrint = iPrint - 300
For i = 2 To 32 Step 5
    Printer.Line (PLeft(i) - 50, iPrint + 150)-(PLeft(i) - 50, iPrint + 1350)
Next i
iPrint = iPrint + 300
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle1
    Exit Sub
End If
'add by nickc 2007/07/27 商標處改格式
If txt1(3) = "020" And txt1(4) = "020" Then
    Printer.CurrentX = PLeft1(1) - (Printer.TextWidth(Title_601) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_601 'Modify By Sindy 2015/2/4"異議"
    Printer.CurrentX = PLeft1(2) - (Printer.TextWidth(Title_603) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_603 'Modify By Sindy 2015/2/4"裁定"
    Printer.CurrentX = PLeft1(3) - (Printer.TextWidth(Title_605) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_605 'Modify By Sindy 2015/2/4"撤銷"
    Printer.CurrentX = PLeft1(4) - (Printer.TextWidth("") / 2)
    Printer.CurrentY = iPrint
    Printer.Print "" 'Modify By Sindy 2015/2/4"註冊不當撤銷"
    Printer.CurrentX = PLeft1(5) - (Printer.TextWidth(Title_401) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_401 'Modify By Sindy 2015/2/4"復審"
    Printer.CurrentX = PLeft1(6) - (Printer.TextWidth(Title_408) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_408 'Modify By Sindy 2015/2/4"大陸上訴"
    Printer.CurrentX = PLeft1(7) - (Printer.TextWidth("") / 2)
    Printer.CurrentY = iPrint
    Printer.Print ""
Else
    Printer.CurrentX = PLeft1(1) - (Printer.TextWidth(Title_601) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_601 'Modify By Sindy 2015/2/4"異議"
    Printer.CurrentX = PLeft1(2) - (Printer.TextWidth(Title_603) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_603 'Modify By Sindy 2015/2/4"評定"
    Printer.CurrentX = PLeft1(3) - (Printer.TextWidth(Title_605) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_605 'Modify By Sindy 2015/2/4"廢止"
    Printer.CurrentX = PLeft1(4) - (Printer.TextWidth(Title_401) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_401 'Modify By Sindy 2015/2/4"訴願"
    'Printer.CurrentX = PLeft1(5) - (Printer.TextWidth("再訴願") / 2)
    Printer.CurrentX = PLeft1(5) - (Printer.TextWidth(Title_403) / 2)
    Printer.CurrentY = iPrint
    'Printer.Print "再訴願"
    Printer.Print Title_403 'Modify By Sindy 2015/2/4"行政訴訟"
    'Printer.CurrentX = PLeft1(6) - (Printer.TextWidth("行政訴訟") / 2)
    Printer.CurrentX = PLeft1(6) - (Printer.TextWidth(Title_408) / 2)
    Printer.CurrentY = iPrint
    'Printer.Print "行政訴訟"
    Printer.Print Title_408 'Modify By Sindy 2015/2/4"行政訴訟上訴"
    Printer.CurrentX = PLeft1(7) - (Printer.TextWidth(Title_404) / 2)
    Printer.CurrentY = iPrint
    Printer.Print Title_404 'Modify By Sindy 2015/2/4"再審之訴"
End If
iPrint = iPrint + 300
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle1
    Exit Sub
End If
Printer.Font.Underline = True
For k = 2 To 32 Step 5
    'edit by nickc 2007/07/27 商標處改格式
    If txt1(3) = "020" And txt1(4) = "020" And k > 27 Then
    Else
        Printer.CurrentX = PLeft(k)
        Printer.CurrentY = iPrint
        Printer.Print "正    確"
        Printer.CurrentX = PLeft(k + 2)
        Printer.CurrentY = iPrint
        Printer.Print "錯    誤"
        Printer.Font.Underline = False
        Printer.CurrentX = PLeft(k + 4) - 50
        Printer.CurrentY = iPrint
        Printer.Print "準確率"
        Printer.Font.Underline = True
    End If
Next k
Printer.Font.Underline = False
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "業務區別"
For k = 2 To 32 Step 5
    'edit by nickc 2007/07/27 商標處改格式
    If txt1(3) = "020" And txt1(4) = "020" And k > 27 Then
    Else
        Printer.CurrentX = PLeft(k)
        Printer.CurrentY = iPrint
        Printer.Print "勝"
        Printer.CurrentX = PLeft(k + 1)
        Printer.CurrentY = iPrint
        Printer.Print "敗"
        Printer.CurrentX = PLeft(k + 2)
        Printer.CurrentY = iPrint
        Printer.Print "勝"
        Printer.CurrentX = PLeft(k + 3)
        Printer.CurrentY = iPrint
        Printer.Print "敗"
    End If
Next k
iPrint = iPrint + 300
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle1
    Exit Sub
End If
'Printer.CurrentX = 0
'Printer.CurrentY = iPrint
'Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
'iPrint = iPrint + 300
ShowLine
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle1
    Exit Sub
End If
End Sub

'Add By Cheng 2003/03/11
Sub PrintTitle1_1()
GetPleft
Printer.Font.Size = 8
'Printer.CurrentX = 0
'Printer.CurrentY = iPrint
'Printer.Line (0, iPrint + 150)-(19000 / 8 * 2, iPrint + 150)
ShowLine_1
iPrint = iPrint - 300
For i = 2 To 6 Step 5
    Printer.Line (PLeft(i) - 50, iPrint + 150)-(PLeft(i) - 50, iPrint + 1350)
Next i
iPrint = iPrint + 300
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle1
    Exit Sub
End If
Printer.CurrentX = PLeft1(1) - (Printer.TextWidth("申請意見書") / 2)
Printer.CurrentY = iPrint
Printer.Print "申請意見書"
iPrint = iPrint + 300
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle1_1
    Exit Sub
End If
Printer.Font.Underline = True
For k = 2 To 6 Step 5
    Printer.CurrentX = PLeft(k)
    Printer.CurrentY = iPrint
    Printer.Print "正    確"
    Printer.CurrentX = PLeft(k + 2)
    Printer.CurrentY = iPrint
    Printer.Print "錯    誤"
    Printer.Font.Underline = False
    Printer.CurrentX = PLeft(k + 4) - 50
    Printer.CurrentY = iPrint
    Printer.Print "準確率"
    Printer.Font.Underline = True
Next k
Printer.Font.Underline = False
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "業務區別"
For k = 2 To 6 Step 5
    Printer.CurrentX = PLeft(k)
    Printer.CurrentY = iPrint
    Printer.Print "勝"
    Printer.CurrentX = PLeft(k + 1)
    Printer.CurrentY = iPrint
    Printer.Print "敗"
    Printer.CurrentX = PLeft(k + 2)
    Printer.CurrentY = iPrint
    Printer.Print "勝"
    Printer.CurrentX = PLeft(k + 3)
    Printer.CurrentY = iPrint
    Printer.Print "敗"
Next k
iPrint = iPrint + 300
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle1_1
    Exit Sub
End If
'Printer.CurrentX = 0
'Printer.CurrentY = iPrint
'Printer.Line (0, iPrint + 150)-(19000 / 8 * 2, iPrint + 150)
'iPrint = iPrint + 300
ShowLine_1
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle1_1
    Exit Sub
End If
End Sub

Sub PrintDatil()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
'Modify By Cheng 2002/05/06
'若承辦人相同則不印
If m_strPromoter <> strTemp(0) Then
   Printer.Print strTemp(0)
   m_strPromoter = strTemp(0)
Else
   Printer.Print ""
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
For i = 2 To 36
    'edit by nickc 2007/07/27 商標處改格式
    If txt1(3) = "020" And txt1(4) = "020" And i > 31 Then
    Else
        Select Case i
        Case 6, 11, 16, 21, 26
            Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "##0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(strTemp(i), "##0.00") & "%"
        Case 31, 36
            If Val(SavDay1) = 1 Then
                Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "##0.00"))
                Printer.CurrentY = iPrint
                Printer.Print Format(strTemp(i), "##0.00") & "%"
            End If
        Case 27, 28, 29, 30, 32, 33, 34, 35
            If Val(SavDay1) = 1 Then
                Printer.CurrentX = PLeft(i) + 200 - Printer.TextWidth(Format(strTemp(i), "####0"))
                Printer.CurrentY = iPrint
                Printer.Print Format(strTemp(i), "####0")
            End If
        Case Else
            Printer.CurrentX = PLeft(i) + 200 - Printer.TextWidth(Format(strTemp(i), "####0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(strTemp(i), "####0")
        End Select
    End If
Next i
For i = 2 To 32 Step 5
    'edit by nickc 2007/07/27 商標處改格式
    If txt1(3) = "020" And txt1(4) = "020" And i > 32 Then
    Else
        Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
    End If
Next i
iPrint = iPrint + 300
End Sub

'Add By Cheng 2003/03/11
Sub PrintDatil_1()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
'Modify By Cheng 2002/05/06
'若承辦人相同則不印
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
    Select Case i
    Case 6, 11, 16, 21, 26
        Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "##0.00"))
        Printer.CurrentY = iPrint
        Printer.Print Format(strTemp(i), "##0.00") & "%"
    Case 31, 36
        If Val(SavDay1) = 1 Then
            Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "##0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(strTemp(i), "##0.00") & "%"
        End If
    Case 27, 28, 29, 30, 32, 33, 34, 35
        If Val(SavDay1) = 1 Then
            Printer.CurrentX = PLeft(i) + 200 - Printer.TextWidth(Format(strTemp(i), "####0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(strTemp(i), "####0")
        End If
    Case Else
        Printer.CurrentX = PLeft(i) + 200 - Printer.TextWidth(Format(strTemp(i), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(strTemp(i), "####0")
    End Select
Next i
For i = 2 To 6 Step 5
    Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
Next i
iPrint = iPrint + 300
End Sub

'Add By Cheng 2003/03/11
Sub PrintDatil_2()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
'Modify By Cheng 2002/05/06
'若承辦人相同則不印
If m_strPromoter <> strTemp(0) Then
   Printer.Print strTemp(0)
   m_strPromoter = strTemp(0)
Else
   Printer.Print ""
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
For i = 2 To 11
    'edit by nickc 2007/07/27 商標處改格式
    If txt1(3) = "020" And txt1(4) = "020" And i > 6 Then
    Else
        Select Case i
        Case 6, 11, 16, 21, 26
            Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "##0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(strTemp(i), "##0.00") & "%"
        Case 31, 36
            If Val(SavDay1) = 1 Then
                Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "##0.00"))
                Printer.CurrentY = iPrint
                Printer.Print Format(strTemp(i), "##0.00") & "%"
            End If
        Case 27, 28, 29, 30, 32, 33, 34, 35
            If Val(SavDay1) = 1 Then
                Printer.CurrentX = PLeft(i) + 200 - Printer.TextWidth(Format(strTemp(i), "####0"))
                Printer.CurrentY = iPrint
                Printer.Print Format(strTemp(i), "####0")
            End If
        Case Else
            Printer.CurrentX = PLeft(i) + 200 - Printer.TextWidth(Format(strTemp(i), "####0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(strTemp(i), "####0")
        End Select
    End If
Next i
For i = 2 To 11 Step 5
    Printer.Line (PLeft(i) - 50, iPrint - 150)-(PLeft(i) - 50, iPrint + 450)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft()
If Val(SavDay1) = 1 Then
    Erase PLeft
    Erase PLeft1
    PLeft(0) = 0
    PLeft(1) = 1000
    PLeft(2) = 2200
    For i = 3 To 36
        PLeft(i) = 2200 + ((i - 2) * 480)
    Next i
    PLeft1(1) = PLeft(4) + 480 / 2
    PLeft1(2) = PLeft(9) + 480 / 2
    PLeft1(3) = PLeft(14) + 480 / 2
    PLeft1(4) = PLeft(19) + 480 / 2
    PLeft1(5) = PLeft(24) + 480 / 2
    PLeft1(6) = PLeft(29) + 480 / 2
    PLeft1(7) = PLeft(34) + 480 / 2
Else
    Erase PLeft
    Erase PLeft1
    PLeft(0) = 0
    PLeft(1) = 1000
    PLeft(2) = 2200
    For i = 3 To 36
        PLeft(i) = 2200 + ((i - 2) * 660)
    Next i
    PLeft1(1) = PLeft(4) + 660 / 2
    PLeft1(2) = PLeft(9) + 660 / 2
    PLeft1(3) = PLeft(14) + 660 / 2
    PLeft1(4) = PLeft(19) + 660 / 2
    PLeft1(5) = PLeft(24) + 660 / 2
'    PLeft1(6) = PLeft(29) + 480 / 2
'    PLeft1(7) = PLeft(34) + 480 / 2
End If
End Sub

Private Sub Form_Load()

MoveFormToCenter Me
txt1(0) = GetSystemKindByNickTnoS
SeekPrintL = Printer.Orientation
PUB_SetPrinter Me.Name, Combo1, , , SeekPrint 'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Printer = Printers(SeekPrint)
Printer.Orientation = SeekPrintL
Set frm020409 = Nothing
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
'edit by nickc 2007/07/27 商標處改格式
If txt1(3) = "020" And txt1(4) = "020" And Page <= 1 Then
    Printer.Line (0, iPrint + 150)-(19000 / 8 * 7, iPrint + 150)
Else
    Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
End If
iPrint = iPrint + 300
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    If Val(SavDay1) = 1 Then
        PrintTitle1
    Else
        PrintTitle2
    End If
End If
End Sub

'Add By Cheng 2003/03/11
Sub ShowLine_1()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000 / 8 * 2, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    If Val(SavDay1) = 1 Then
        PrintTitle1_1
    Else
        PrintTitle2_2
    End If
End If
End Sub

'Add By Cheng 2003/03/11
Sub ShowLine_2()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
'edit by nickc 2007/07/27 商標處改格式
If txt1(3) = "020" And txt1(4) = "020" And Page > 2 Then
    Printer.Line (0, iPrint + 150)-(19000 / 11 * 3.2, iPrint + 150)
Else
    Printer.Line (0, iPrint + 150)-(19000 / 11 * 5, iPrint + 150)
End If
iPrint = iPrint + 300
If iPrint >= 13900 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    If Val(SavDay1) = 1 Then
        PrintTitle1_1
    Else
        PrintTitle2_2
    End If
End If
End Sub

