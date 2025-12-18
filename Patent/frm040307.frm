VERSION 5.00
Begin VB.Form frm040307 
   BorderStyle     =   1  '單線固定
   Caption         =   "催審函/催審表"
   ClientHeight    =   4065
   ClientLeft      =   3150
   ClientTop       =   2700
   ClientWidth     =   3945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   3945
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   13
      Left            =   1215
      MaxLength       =   4
      TabIndex        =   13
      Top             =   2880
      Width           =   1155
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   14
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   14
      Top             =   2880
      Width           =   1155
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   10
      Left            =   2970
      MaxLength       =   2
      TabIndex        =   10
      Top             =   2250
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   9
      Left            =   2580
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2250
      Width           =   330
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   8
      Left            =   1725
      MaxLength       =   6
      TabIndex        =   8
      Top             =   2250
      Width           =   795
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   7
      Left            =   1230
      MaxLength       =   3
      TabIndex        =   7
      Top             =   2250
      Width           =   435
   End
   Begin VB.OptionButton opt 
      Caption         =   "本所案號："
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   25
      Top             =   2280
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   11
      Left            =   1230
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1920
      Width           =   1155
   End
   Begin VB.OptionButton opt 
      Caption         =   "發文日期："
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   1950
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1230
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1590
      Width           =   1155
   End
   Begin VB.OptionButton opt 
      Caption         =   "催審期限："
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   1620
      Value           =   -1  'True
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   12
      Left            =   2520
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1920
      Width           =   1155
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   350
      Index           =   1
      Left            =   2805
      TabIndex        =   16
      Top             =   12
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   1980
      TabIndex        =   15
      Top             =   12
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   12
      Top             =   2580
      Width           =   1155
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   1215
      MaxLength       =   4
      TabIndex        =   11
      Top             =   2580
      Width           =   1155
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   2520
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1590
      Width           =   1155
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   888
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1224
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   1188
      MaxLength       =   1
      TabIndex        =   1
      Top             =   816
      Width           =   300
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1176
      TabIndex        =   0
      Top             =   444
      Width           =   2370
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質條件僅適用於發文日期及本所案號"
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   7
      Left            =   120
      TabIndex        =   29
      Top             =   3600
      Width           =   3705
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   28
      Top             =   2970
      Width           =   990
   End
   Begin VB.Line Line5 
      X1              =   2385
      X2              =   2505
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "92/02/01"
      Height          =   180
      Index           =   4
      Left            =   1380
      TabIndex        =   27
      Top             =   3330
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "新系統上線日："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   26
      Top             =   3360
      Width           =   1305
   End
   Begin VB.Line Line3 
      X1              =   1665
      X2              =   3075
      Y1              =   2385
      Y2              =   2385
   End
   Begin VB.Line Line4 
      X1              =   2355
      X2              =   2475
      Y1              =   2085
      Y2              =   2085
   End
   Begin VB.Line Line2 
      X1              =   2385
      X2              =   2505
      Y1              =   2715
      Y2              =   2715
   End
   Begin VB.Line Line1 
      X1              =   2355
      X2              =   2475
      Y1              =   1770
      Y2              =   1770
   End
   Begin VB.Label Label3 
      Caption         =   "(1.定稿2.報表)"
      Height          =   180
      Left            =   1296
      TabIndex        =   22
      Top             =   1260
      Width           =   1308
   End
   Begin VB.Label Label2 
      Caption         =   "(1.催實審 2.催審)"
      Height          =   180
      Left            =   1656
      TabIndex        =   21
      Top             =   840
      Width           =   1488
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   20
      Top             =   1272
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   2670
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "催審函性質："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   888
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   504
      Width           =   996
   End
End
Attribute VB_Name = "frm040307"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim s As Integer, strSql As String, iPrint As Integer, Page As Integer, strSQL1 As String, strSQL2 As String
Dim strTemp2 As Variant, StrTest2 As String, i As Integer, j As Integer, strTemp1 As Variant, StrTest As String
Dim strTemp(0 To 19) As String, StrYear1 As String, StrYear2 As String, PLeft(0 To 9) As Integer
Dim Area As String, SavDay1 As String, SavDay2 As String, SavDay(0 To 1) As String
'Add By Cheng 2002/09/11
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0 '確定
      blnClkSure = False
     
     If Len(txt1(0)) = 0 Then
        s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        Exit Sub
     Else
        If Len(txt1(1)) = 0 Then
            s = MsgBox("催審函性質不可空白!!", , "USER 輸入錯誤")
            txt1(1).SetFocus
            Exit Sub
        Else
            If Len(txt1(2)) = 0 Then
                s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
                txt1(2).SetFocus
                Exit Sub
            Else
               'Add By Cheng 2002/08/05
               If Me.opt(0).Value Then
                  If CheckIsTaiwanDate(Me.txt1(3).Text) = False Then
                     Me.txt1(3).SetFocus
                     Exit Sub
                  End If
                  If CheckIsTaiwanDate(Me.txt1(4).Text) = False Then
                     Me.txt1(4).SetFocus
                     Exit Sub
                  End If
                  If Val(Me.txt1(3).Text) > Val(Me.txt1(4).Text) Then
                     MsgBox "催審期限區間輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(3).SetFocus
                     Exit Sub
                  End If
               ElseIf Me.opt(1).Value Then
                  If CheckIsTaiwanDate(Me.txt1(11).Text) = False Then
                     Me.txt1(11).SetFocus
                     Exit Sub
                  End If
                  If CheckIsTaiwanDate(Me.txt1(12).Text) = False Then
                     Me.txt1(12).SetFocus
                     Exit Sub
                  End If
                  If Val(Me.txt1(11).Text) > Val(Me.txt1(12).Text) Then
                     MsgBox "發文日期區間輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(11).SetFocus
                     Exit Sub
                  End If
               Else
                  If Me.txt1(7).Text <> "P" And Me.txt1(7).Text <> "PS" Then
                     MsgBox "系統類別輸入錯誤!!!", vbExclamation + vbOKOnly
                     Me.txt1(7).SetFocus
                     Exit Sub
                  End If
                  If Me.txt1(8).Text = "" Then
                     MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
                     Me.txt1(8).SetFocus
                     Exit Sub
                  End If
               End If
               'Add By Cheng 2002/09/11
               If Me.txt1(5).Text <> "" And Me.txt1(6).Text <> "" Then
                  If Val(Me.txt1(5).Text) > Val(Me.txt1(6).Text) Then
                     MsgBox "申請國家範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(5).SetFocus
                     txt1_GotFocus 5
                     Exit Sub
                  End If
               End If
               '2005/10/18 ADD BY SONIA 加入案件性質條件
               If Me.txt1(13).Text <> "" And Me.txt1(14).Text <> "" Then
                  If Val(Me.txt1(13).Text) > Val(Me.txt1(14).Text) Then
                     MsgBox "案件性質範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(13).SetFocus
                     txt1_GotFocus 13
                     Exit Sub
                  End If
               End If
               '2005/10/18 END

'               'Add By Cheng 2002/03/19
'               If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
'                  Me.txt1(3).SetFocus
'                  txt1_GotFocus 3
'                  Exit Sub
'               End If
'               If PUB_CheckKeyInDate(Me.txt1(4)) = -1 Then
'                  Me.txt1(4).SetFocus
'                  txt1_GotFocus 4
'                  Exit Sub
'               End If
'                If Len(txt1(4)) = 0 Then
'                    s = MsgBox("催審期限區間不可空白!!", , "USER 輸入錯誤")
'                    txt1(3).SetFocus
'                    txt1_GotFocus (3)
'                    Exit Sub
'                Else
                    ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/30 清除查詢印表記錄檔欄位
                    If Val(txt1(2)) = 2 Then
                        pub_QL05 = pub_QL05 & ";" & Label1(5) & "2.報表" 'Add By Sindy 2010/11/30
                        StrTest = StrTest2
                        strTemp1 = Split(UCase(StrTest), ",")
                        strTemp2 = Split(UCase(txt1(0)), ",")
                        For i = 0 To UBound(strTemp1)
                            s = 0
                            For j = 0 To UBound(strTemp2)
                                If strTemp2(j) = strTemp1(i) Then
                                    s = 1
                                    Exit For
                                End If
                            Next j
                            If s = 0 Then
                                StrTest = Replace(StrTest, strTemp1(i), "")
                            End If
                        Next i
                        Screen.MousePointer = vbHourglass
                        Me.Enabled = False
                        Process
                        Me.Enabled = True
                        Screen.MousePointer = vbDefault
                    Else
                        pub_QL05 = pub_QL05 & ";" & Label1(5) & "1.定稿" 'Add By Sindy 2010/11/30
                        Screen.MousePointer = vbHourglass
                        Me.Enabled = False
                        'Add by Morgan 2009/9/18 個案定稿改跑新程式
                        If txt1(2) = "1" And opt(2).Value = True Then
                           ProcessNew
                        Else
                           Process
                        End If
                        Me.Enabled = True
                        Screen.MousePointer = vbDefault
                    End If
'                End If
            End If
        End If
     End If
Case 1 '結束
     Unload Me
End Select
End Sub

'Add by Morgan 2009/9/18 個案定稿(指定本所案號)
Sub ProcessNew()
Dim stSQL As String, iR As Integer
Dim adoRst As ADODB.Recordset
Dim ET01 As String, ET02 As String, ET03 As String
Dim strCP130 As String   '2009/5/15 ADD BY SONIA
Dim iMonth As Integer
Dim stRefCP10 As String '來函案件性質
Dim stCon As String
Dim strCP09New As String 'Add by Amy 2014/09/09 B類總收文號 for P案電子化
Dim bolInTrans As Boolean 'Added by Morgan 2016/5/18
Dim m_Subject As String, strCP27 As String 'Added by Morgan 2016/5/20
   
On Error GoTo ErrHnd

   ET01 = "16"
   stCon = ""
   If txt1(5) <> "" Then
      stCon = stCon & " and pa09>='" & txt1(5) & "'"
   End If
   If txt1(6) <> "" Then
      stCon = stCon & " and pa09<='" & txt1(6) & "'"
   End If
   If txt1(5) <> "" Or txt1(6) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/11/30
   End If
   If txt1(13) <> "" Then
      stCon = stCon & " and cp10>='" & txt1(13) & "'"
   End If
   If txt1(14) <> "" Then
      stCon = stCon & " and cp10<='" & txt1(14) & "'"
   End If
   If txt1(13) <> "" Or txt1(14) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1(6) & txt1(13) & "-" & txt1(14) 'Add By Sindy 2010/11/30
   End If
   pub_QL05 = pub_QL05 & ";" & opt(2).Caption & txt1(7) & "-" & txt1(8) & "-" & Right("0" & txt1(9), 1) & "-" & Right("00" & txt1(10), 2) 'Add By Sindy 2010/11/30
   If txt1(1) = "1" Then
      pub_QL05 = pub_QL05 & ";" & Label1(1) & "1.催實審" 'Add By Sindy 2010/11/30
   Else
      pub_QL05 = pub_QL05 & ";" & Label1(1) & "2.催審" 'Add By Sindy 2010/11/30
   End If
   
   'Modify by Amy 2014/09/09 +pa26,pa75
   'Modify by Morgan 2010/5/18 若發文日為111111時改抓基本檔申請日
   'Modified by Lydia 2018/01/03 +別名f0, 串接FMP2openSQL
   stSQL = "select decode(pa09,'000',cpm03,cpm04) C1,decode(cp27,19221111,pa10,nvl(cp47,cp27)) C2,cp09,cp10,CP43,pa09,pa26,pa75,cp45,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) C3,CP44,CP45" & _
      " from nextprogress,caseprogress f0,patent,casepropertymap" & _
      " where np02='" & txt1(7) & "' and np03='" & txt1(8) & "'" & _
      " and np04='" & Right("0" & txt1(9), 1) & "' and np05='" & Right("00" & txt1(10), 2) & "'" & _
      " and np06 is null and np07='411' and cp09(+)=np01 AND CP57 IS NULL AND CP24 IS NULL" & _
      " and PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57 IS NULL AND PA108 IS NULL" & _
      " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & stCon & FMP2openSQL
   iR = 1
   Set adoRst = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      With adoRst
      InsertQueryLog (.RecordCount) 'Add By Sindy 2010/11/30
      .MoveFirst
      Do While Not .EOF
         ET02 = .Fields("cp09")
         If .Fields("pa09") = "000" Then
            stRefCP10 = "1204"
            'Modified by Morgan 2016/9/26
            'If txt1(2) = "1" Then
            If txt1(2) = "2" Then
            'end 2016/9/26
               ET03 = "02"
            Else
               Select Case .Fields("cp10")
               '舉發,舉發答辯
               Case "803", "804"
                  ET03 = "04"
               '技術報告
               Case "421", "807"
                  ET03 = "05"
                  stRefCP10 = "1405"
               '訴願
               Case "501"
                  ET03 = "06"
               Case Else
                  ET03 = "01"
               End Select
            End If
         Else
            ET03 = "03"
         End If
         
         EndLetter ET01, ET02, ET03, strUserNum
         
         '舉發答辯要抓相關收文號的對造號數
         If .Fields("cp10") = "804" And Not IsNull(.Fields("CP43")) Then
            stSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " SELECT '" & ET01 & "','" & ET02 & "','" & ET03 & "','" & _
               strUserNum & "','對造號數',CP36 FROM CASEPROGRESS WHERE CP09='" & .Fields("CP43") & "'"
            cnnConnection.Execute stSQL, iR
         '技術報告
         ElseIf ET03 = "05" Then
            stSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " select '" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "'" & _
               ",'對造號數',CP36 FROM CASEPROGRESS WHERE CP43='" & .Fields("CP09") & "' AND CP10='" & stRefCP10 & "'"
            cnnConnection.Execute stSQL, iR
         End If
         
         If .Fields("pa09") = "000" Then
            stSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " select '" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "'" & _
               ",'實審通知書',SUBSTR(CP05,1,4)-1911||'年'||SUBSTR(CP05,5,2)||'月'||SUBSTR(CP05,7,2)||'日'||CP08" & _
               " FROM CASEPROGRESS WHERE CP43='" & .Fields("CP09") & "' AND CP10='" & stRefCP10 & "'"
            cnnConnection.Execute stSQL, iR
         End If
         
         iMonth = Left(strSrvDate(1), 4) * 12 + Val(Mid(strSrvDate(1), 5, 2)) - (Left(.Fields("C2"), 4) * 12 + Val(Mid(.Fields("C2"), 5, 2)))
         stSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & _
            strUserNum & "','已發文時間','" & iMonth & "')"
         cnnConnection.Execute stSQL, iR
         
         '案件性質為申請結尾的不印申請2字
         If Right(.Fields("C1"), 2) = "申請" Then
            stSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & _
               strUserNum & "','申請不印','♀')"
            cnnConnection.Execute stSQL, iR
         End If
         
         'Modify by Morgan 2009/9/18 自Process移來
         '2009/3/24 ADD BY SONIA 產生內部收文-催審,同時上發文日
         If Me.opt(2).Value Then
            '2009/5/15 ADD BY SONIA 抓該催審程序的主管機關
            strCP130 = ""
            strExc(0) = "SELECT NVL(A.CP130,CF10),CF02 FROM CASEFEE,(SELECT CP01,PA09,CP10,CP130 FROM CASEPROGRESS,PATENT WHERE CP09='" & .Fields("CP09") & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)) A " & _
                        "WHERE A.CP01=CF01(+) AND A.PA09=CF02(+) AND A.CP10=CF03(+) "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If Not IsNull(RsTemp.Fields(0)) Then strCP130 = RsTemp.Fields(0)
            End If
            '2009/5/15 END
            
            'Added by Morgan 2016/5/18
            cnnConnection.BeginTrans
            bolInTrans = True
            'end 2016/5/18
            
            'Added by Morgan 2016/5/20
            '指示信電子化
            If .Fields("pa09") <> "000" And Left(Pub_StrUserSt03, 1) <> "F" Then
               strCP27 = "NULL"
            Else
               strCP27 = strSrvDate(1)
            End If
            'end 2016/5/20
            
            'Modify by Amy 2014/09/09 +strCP09New
            strCP09New = AutoNo("B", 6)
            'Modified by Morgan 2016/5/20 指示信電子化
            If txt1(1) = "1" Then
               '2009/5/15 MODIFY BY SONIA 加CP130
               strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10," & _
                  "CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP44,CP45,CP123,CP64,CP130) VALUES ('" & txt1(7) & "','" & txt1(8) & "','" & _
                  Left(Me.txt1(9).Text & "0", 1) & "','" & Left(Me.txt1(10).Text & "00", 2) & "'," & strSrvDate(1) & ",'" & strCP09New & "','411'," & CNULL(GetSalesArea(PUB_GetAKindSalesNo(txt1(7), txt1(8), Left(Me.txt1(9).Text & "0", 1), Left(Me.txt1(10).Text & "00", 2)))) & "," & _
                  CNULL(PUB_GetAKindSalesNo(txt1(7), txt1(8), Left(Me.txt1(9).Text & "0", 1), Left(Me.txt1(10).Text & "00", 2))) & ",'" & strUserNum & "','N','N'," & strCP27 & ",'N','" & .Fields("CP09") & "','" & .Fields("CP44") & "','" & .Fields("CP45") & "','" & IIf(RsTemp.Fields(1) = "000", "Y", Null) & "','催實審','" & strCP130 & "')"
            Else
               '2009/5/15 MODIFY BY SONIA 加CP130
               strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10," & _
                  "CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP44,CP45,CP123,CP64,CP130) VALUES ('" & txt1(7) & "','" & txt1(8) & "','" & _
                  Left(Me.txt1(9).Text & "0", 1) & "','" & Left(Me.txt1(10).Text & "00", 2) & "'," & strSrvDate(1) & ",'" & strCP09New & "','411'," & CNULL(GetSalesArea(PUB_GetAKindSalesNo(txt1(7), txt1(8), Left(Me.txt1(9).Text & "0", 1), Left(Me.txt1(10).Text & "00", 2)))) & "," & _
                  CNULL(PUB_GetAKindSalesNo(txt1(7), txt1(8), Left(Me.txt1(9).Text & "0", 1), Left(Me.txt1(10).Text & "00", 2))) & ",'" & strUserNum & "','N','N'," & strCP27 & ",'N','" & .Fields("CP09") & "','" & .Fields("CP44") & "','" & .Fields("CP45") & "','" & IIf(RsTemp.Fields(1) = "000", "Y", Null) & "','催審','" & strCP130 & "')"
            End If
            cnnConnection.Execute strSql, intI
         End If
         '2009/3/24 END
         
         'Add by Amy 2014/09/09 P案電子化
         If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
             If txt1(7) = "P" And .Fields("pa09") = 台灣國家代號 Then
                '新增LetterProgress
                'Modified by Morgan 2018/8/1
                'strExc(1) = PUB_GetLetterJudge(txt1(7), "411", , , txt1(7), txt1(8), Left(Me.txt1(9).Text & "0", 1), Left(Me.txt1(10).Text & "00", 2))
                strExc(1) = PUB_GetLetterJudgeNew("1", txt1(7), "411")
                PUB_AddLetterProgress strCP09New, 1, False, strExc(1), False, .Fields("pa26"), "411", "" & .Fields("pa75"), True
                If ExistCheck("AppForm", "AF01", strCP09New, "", False) = False Then
                    '新增申請書轉檔記錄(B類)
                    PUB_AddAppForm strCP09New
                End If
            End If
         End If
         'end 2014/09/09
         
         'Added by Morgan 2016/5/18
         '指示信電子化
         If .Fields("pa09") <> "000" And Left(Pub_StrUserSt03, 1) <> "F" Then
            m_Subject = "請代為查詢" & .Fields("C1") & "進度，Y/R：" & IIf(Trim("" & .Fields("cp45")) = "", "(請提供)", "" & .Fields("cp45")) & "，O/R：" & .Fields("C3") & "。"
            'Modified by Morgan 2018/7/30 指示信判發人改抓設定檔
            'strExc(2) = Pub_GetSpecMan("PS4")
            strExc(2) = PUB_GetLetterJudgeNew("2", txt1(7), "411", .Fields("pa09"), .Fields("cp10"))
            PUB_AddAppForm strCP09New, True, strExc(2), m_Subject
         End If
         
         cnnConnection.CommitTrans
         bolInTrans = False
         'end 2016/5/18
         
         'Modify by Amy 2014/09/09 +傳strLetterRecNo
         'Add by Morgan 2009/10/15 非台灣固定印兩張--玲玲
         If .Fields("pa09") <> "000" Then
            'Modified by Lydia 2016/03/04 e化後,大陸指示信只需1份(由basLetter決定)
            'NowPrint ET02, ET01, ET03, False, strUserNum, 0, , , , 2, , , , , , , , strCP09New
            'Modified by Morgan 2016/5/20 指示信電子化
            If Left(Pub_StrUserSt03, 1) = "F" Then
               NowPrint ET02, ET01, ET03, False, strUserNum
            Else
               NowPrint ET02, ET01, ET03, True, strUserNum, , , , , , , , , , , , , strCP09New
               frm1105_1.m_RecNo = strCP09New
               frm1105_1.m_PdfName = PUB_CaseNo2FileName(txt1(7), txt1(8), txt1(9), txt1(10)) & ".411.DATA.PDF"
               frm1105_1.m_Subject = m_Subject
               frm1105_1.Show
            End If
            'end 2016/5/20
         Else
            '2014/09/09 +if P台灣案電子化後申請書設2份
            If P台灣案電子化啟用日 <= Val(strSrvDate(1)) And txt1(7) = "P" Then
                NowPrint ET02, ET01, ET03, False, strUserNum, 0, , , , 2, , , , , , , , strCP09New
            Else
                NowPrint ET02, ET01, ET03, False, strUserNum, 0, , , , , , , , , , , , strCP09New
            End If
         End If
         'end 2014/09/09
         .MoveNext
      Loop
      End With
      If strCP27 = strSrvDate(1) Then ShowPrintOk
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/11/30
      ShowNoData
   End If
   
'Add by Morgan 2016/5/18
ErrHnd:
   If Err.NUMBER <> 0 Then
      If bolInTrans Then cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
'end 2016/5/18
   
   Set adoRst = Nothing
End Sub

Sub Process()
Dim strSysDate As String
Dim strCaseType As String
Dim strCaseSitu As String
'Add By Cheng 2002/08/05
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'Add By Cheng 2002/08/07
Dim blnUpdate As Boolean '是否要更新下一程序的期限
'92.1.20 ADD BY SONIA
Dim DataYN As String   '有無符合條件資料

Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R040307 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""

'催審案性質--催實審
If Val(txt1(1)) = 1 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & "1.催實審" 'Add By Sindy 2010/11/30
   'Modify By Cheng 2002/08/21
   If Me.opt(0).Value Then
      strSQL1 = strSQL1 + " and NP07=1204 AND (NP06 IS NULL OR NP06='') "
      If Len(Trim(txt1(3))) <> 0 Then
        strSQL1 = strSQL1 & " AND NP08>=" & Val(ChangeTStringToWString(txt1(3))) & " "
      End If
      If Len(Trim(txt1(4))) <> 0 Then
        strSQL1 = strSQL1 & " AND NP08<=" & Val(ChangeTStringToWString(txt1(4))) & " "
      End If
      If Len(Trim(txt1(3))) <> 0 Or Len(Trim(txt1(4))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & opt(0).Caption & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/11/30
      End If
   End If
'催審案性質--催審
Else
   pub_QL05 = pub_QL05 & ";" & Label1(1) & "2.催審" 'Add By Sindy 2010/11/30
   If Me.opt(0).Value Then
      strSQL1 = strSQL1 + " AND NP07=411 AND (NP06 IS NULL OR NP06='') "
      If Len(Trim(txt1(3))) <> 0 Then
        strSQL1 = strSQL1 & " AND NP08>=" & Val(ChangeTStringToWString(txt1(3))) & " "
      End If
      If Len(Trim(txt1(4))) <> 0 Then
        strSQL1 = strSQL1 & " AND NP08<=" & Val(ChangeTStringToWString(txt1(4))) & " "
      End If
      If Len(Trim(txt1(3))) <> 0 Or Len(Trim(txt1(4))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & opt(0).Caption & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/11/30
      End If
   End If
End If

'Add By Cheng 2002/08/05
If Me.opt(1).Value Then
   If Len(Trim(txt1(11))) <> 0 Then
     strSQL1 = strSQL1 & " AND CP27>=" & Val(ChangeTStringToWString(txt1(11))) & " "
   End If
   If Len(Trim(txt1(12))) <> 0 Then
     strSQL1 = strSQL1 & " AND CP27<=" & Val(ChangeTStringToWString(txt1(12))) & " "
   End If
   If Len(Trim(txt1(11))) <> 0 Or Len(Trim(txt1(12))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & opt(1).Caption & txt1(11) & "-" & txt1(12) 'Add By Sindy 2010/11/30
   End If
End If

'若選擇本所案號
If Me.opt(2).Value Then
   pub_QL05 = pub_QL05 & ";" & opt(2).Caption & txt1(7) & "-" & txt1(8) 'Add By Sindy 2010/11/30
   If Len(txt1(7)) <> 0 Then
       strSQL1 = strSQL1 + " AND CP01='" & txt1(7) & "' "
   End If
   If Len(txt1(8)) <> 0 Then
       strSQL1 = strSQL1 + " AND CP02='" & txt1(8) & "' "
   End If
   If Len(txt1(9)) <> 0 Then
       strSQL1 = strSQL1 + " AND CP03='" & txt1(9) & "' "
       pub_QL05 = pub_QL05 & "-" & txt1(9) 'Add By Sindy 2010/11/30
   End If
   If Len(txt1(10)) <> 0 Then
       strSQL1 = strSQL1 + " AND CP04='" & txt1(10) & "' "
       pub_QL05 = pub_QL05 & "-" & txt1(10) 'Add By Sindy 2010/11/30
   End If
End If

'Add By Cheng 2002/08/05
If Me.opt(1).Value Or Me.opt(2).Value Then
   strSQL1 = strSQL1 & " AND CP27 IS NOT NULL AND CP24 IS NULL AND CP57 IS NULL AND CP09<'C' "
End If
'是否出名
strSQL1 = strSQL1 + " AND (CP22='' OR CP22 IS NULL) "
strSQL2 = strSQL1
'系統類別
If Len(StrTest) <> 0 Then
   'Modify By Cheng 2002/08/21
   If Me.opt(0).Value Then
      strSQL1 = strSQL1 + " and NP02 in (" & SQLGrpStr(StrTest, 1) & ") "
      strSQL2 = strSQL2 + " and NP02 in (" & SQLGrpStr(StrTest, 5) & ") "
   Else
      strSQL1 = strSQL1 + " and CP01 in (" & SQLGrpStr(StrTest, 1) & ") "
      strSQL2 = strSQL2 + " and CP01 in (" & SQLGrpStr(StrTest, 5) & ") "
   End If
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/11/30
End If

'申請國家
If Len(txt1(5)) <> 0 Then
    strSQL1 = strSQL1 & " AND SUBSTR(pa09,1,3)>='" & txt1(5) & "' "
    strSQL2 = strSQL2 & " and SUBSTR(sp09,1,3)>='" & txt1(5) & "' "
End If
If Len(txt1(6)) <> 0 Then
    'Modify By Cheng 2003/02/21
    '加單引號
'    strSQL1 = strSQL1 + " AND SUBSTR(pa09,1,3)<=" & txt1(6) & "' "
'    strSQL2 = strSQL2 & " and SUBSTR(sp09,1,3)<=" & txt1(6) & "' "
    strSQL1 = strSQL1 + " AND SUBSTR(pa09,1,3)<='" & txt1(6) & "' "
    strSQL2 = strSQL2 & " and SUBSTR(sp09,1,3)<='" & txt1(6) & "' "
End If
If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/11/30
End If

'2005/10/18 ADD BY SONIA
'案件性質
If Not Me.opt(0).Value Then
   If Len(txt1(13)) <> 0 Then
       strSQL1 = strSQL1 & " AND CP10>='" & txt1(13) & "' "
       strSQL2 = strSQL2 & " and CP10>='" & txt1(13) & "' "
   End If
   If Len(txt1(14)) <> 0 Then
       strSQL1 = strSQL1 + " AND CP10<='" & txt1(14) & "' "
       strSQL2 = strSQL2 & " and CP10<='" & txt1(14) & "' "
   End If
   If Len(txt1(13)) <> 0 Or Len(txt1(14)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(6) & txt1(13) & "-" & txt1(14) 'Add By Sindy 2010/11/30
   End If
End If
'2005/10/18 END

'是否閉卷
strSQL1 = strSQL1 + " AND PA57 IS NULL  "
strSQL2 = strSQL2 + " AND SP15 IS NULL  "
CheckOC
'Modify By Cheng 2002/08/21
If Me.opt(0).Value Then '若選擇催審期限(要考慮下一程序的資料)
   'Modify By Cheng 2002/08/05
   '多顯示申請國家代號
   'Modified by Lydia 2018/01/03 +別名f0, 串接FMP2openSQL
   strSql = "SELECT CP27,NP08,PA11,NP02||'-'||NP03||'-'||NP04||'-'||NP05," & _
      "NVL(PA05,NVL(PA06,PA07)),decode(pa09,'000',cpm03,cpm04) C5,NVL(NA03,NA04)," & _
      "NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),S1.ST02,S2.ST02," & _
      "NP01,CP10,NP02,NP03,NP04,NP05,NP22,NP07,NP09,NP01,CP09,PA09 FROM NEXTPROGRESS,CASEPROGRESS f0," & _
      "PATENT,STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,CUSTOMER WHERE np01=CP09(+) AND " & _
      "np10=S2.ST01(+) AND cp14=S1.ST01(+) AND np02=PA01(+) AND np03=PA02(+) AND np04=PA03(+) AND " & _
      "np05=PA04(+) AND SUBSTR(PA26,1,8)=cu01(+) AND " & _
      "decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND pa09=na01(+) AND " & _
      "CP01=cpm01(+) AND cp10=cpm02(+) " & strSQL1 & FMP2openSQL
      
   'Modified by Lydia 2018/01/03 +別名f0, 串接FMP2openSQL
   strSql = strSql + " UNION ALL SELECT CP27,NP08,SP11," & _
      "NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07))," & _
      "decode(sp09,'000',cpm03,cpm04) C5,NVL(NA03,NA04)," & _
      "NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),S1.ST02,S2.ST02,NP01,CP10," & _
      "NP02,NP03,NP04,NP05,NP22,NP07,NP09,NP01,CP09,SP09 FROM NEXTPROGRESS,CASEPROGRESS f0,SERVICEPRACTICE," & _
      "STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,CUSTOMER WHERE np01=CP09(+) AND " & _
      "np10=S2.ST01(+) AND cp14=S1.ST01(+) AND np02=SP01(+) AND np03=SP02(+) AND " & _
      "np04=SP03(+) AND np05=SP04(+) AND SUBSTR(sP08,1,8)=cu01(+) AND " & _
      "decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND sp09=na01(+) AND " & _
      "CP01=cpm01(+) AND cp10=cpm02(+) " & strSQL2 & FMP2openSQL
Else '若非選擇催審期限(不考慮下一程序的資料)
   'Modified by Lydia 2018/01/03 +別名f0, 串接FMP2openSQL
   strSql = "SELECT CP27,'',PA11,CP01||'-'||CP02||'-'||CP03||'-'||CP04," & _
      "NVL(PA05,NVL(PA06,PA07)),decode(pa09,'000',cpm03,cpm04) C5,NVL(NA03,NA04)," & _
      "NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),S1.ST02,S2.ST02," & _
      "CP09,CP10,CP01,CP02,CP03,CP04,'','','',CP09,CP09,PA09 FROM CASEPROGRESS f0," & _
      "PATENT,STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,CUSTOMER WHERE " & _
      "Cp13=S2.ST01(+) AND cp14=S1.ST01(+) AND Cp01=PA01(+) AND Cp02=PA02(+) AND Cp03=PA03(+) AND " & _
      "Cp04=PA04(+) AND SUBSTR(PA26,1,8)=cu01(+) AND " & _
      "decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND pa09=na01(+) AND " & _
      "CP01=cpm01(+) AND cp10=cpm02(+) " & strSQL1 & FMP2openSQL
      
    'Modified by Lydia 2018/01/03 +別名f0, 串接FMP2openSQL
   strSql = strSql + " UNION ALL SELECT CP27,'',SP11," & _
      "CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07))," & _
      "decode(sp09,'000',cpm03,cpm04) C5,NVL(NA03,NA04)," & _
      "NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),S1.ST02,S2.ST02,CP09,CP10," & _
      "CP01,CP02,CP03,CP04,'','','',CP09,CP09,SP09 FROM CASEPROGRESS f0,SERVICEPRACTICE," & _
      "STAFF S1,STAFF S2,NATION,CASEPROPERTYMAP,CUSTOMER WHERE " & _
      "Cp13=S2.ST01(+) AND cp14=S1.ST01(+) AND Cp01=SP01(+) AND Cp02=SP02(+) AND " & _
      "Cp03=SP03(+) AND Cp04=SP04(+) AND SUBSTR(sP08,1,8)=cu01(+) AND " & _
      "decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND sp09=na01(+) AND " & _
      "CP01=cpm01(+) AND cp10=cpm02(+) " & strSQL2 & FMP2openSQL
End If
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/11/30
         'Remove by Morgan 2009/9/16 都不再詢問,改由每日批次報表控制週期
         ''Add By Cheng 2002/08/06
         ''若選擇催審期限, 加詢問使用者是否要更新
         'If Me.opt(0).Value Then
         '   If MsgBox("是否再管制三個月?", vbExclamation + vbYesNo) = vbYes Then
         '      blnUpdate = True
         '   Else
         '      blnUpdate = False
         '   End If
         ''其他選項不必詢問一律不更新
         'Else
         '   blnUpdate = False
         'End If
         
        DataYN = "N"
        .MoveFirst
        DoEvents
        Do While .EOF = False
            '92.1.20 ADD BY SONIA 舊系統發文之修正,申請優先權證明,申請英文證明,讓與,領證,補換發證書不催審
            If Not Me.opt(0).Value Then
               If .Fields(0) < ChangeTStringToWString(ChangeTDateStringToTString(Label1(4).Caption)) And (.Fields(11) = 修正 Or .Fields(11) = 申請優先權證明 Or .Fields(11) = 申請英文證明 Or .Fields(11) = 讓與 Or .Fields(11) = 領證及繳年費 Or .Fields(11) = 補換發證書) Then
                  GoTo NextRecord
               End If
               If .Fields(0) = 19221111 Then
                  GoTo NextRecord
               End If
            End If
            '92.1.20 END
            'Add By Cheng 2002/08/05
            '檢查案件國家檔的CF05, 若無資料或CF05 IS NULL,則不管制
            If Me.opt(1).Value Or Me.opt(2).Value Then
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               StrSQLa = "Select CF05 FROM CASEFEE WHERE CF01='" & .Fields(12).Value & "' AND CF02='" & .Fields(21).Value & "' AND CF03='" & .Fields(11) & "' AND CF05 IS NOT NULL "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount <= 0 Then
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
                  GoTo NextRecord
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            End If
            
            s = 0
            For i = 0 To 19
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            'Modify By Cheng 2002/08/21
            If Me.opt(0).Value Then
               '催審函性質--催審
               If Val(txt1(1)) = 2 Then
                   If strTemp(14) <> "0" Then
                       strSql = "SELECT PA16 FROM PATENT WHERE PA01='" & ChgSQL(strTemp(12)) & "' AND PA02='" & ChgSQL(strTemp(13)) & "' AND PA03='0' AND PA04='" & ChgSQL(strTemp(15)) & "' "
                       CheckOC2
                       adoRecordset1.CursorLocation = adUseClient
                       adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                       If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                           If CheckStr(adoRecordset1.Fields(0)) = "1" Then
                           Else
                               s = 1
                           End If
                       End If
                       CheckOC2
                   End If
                   If s <> 1 Then
                        'Modify By Cheng 2002/04/12
   '                    strSQL = "SELECT CP05 FROM CASEPROGRESS WHERE CP01='" & ChgSQL(strTemp(12)) & "' AND CP02='" & ChgSQL(strTemp(13)) & "' AND CP03='" & ChgSQL(strTemp(14)) & "' AND CP04='" & ChgSQL(strTemp(15)) & "' AND CP10='1201' AND SUBSTR(CP09,1,1)='C' ORDER BY CP05"
                       strSql = "SELECT CP05 FROM CASEPROGRESS WHERE CP01='" & ChgSQL(strTemp(12)) & "' AND CP02='" & ChgSQL(strTemp(13)) & "' AND CP03='" & ChgSQL(strTemp(14)) & "' AND CP04='" & ChgSQL(strTemp(15)) & "' AND CP10='1201' AND CP09>'C' ORDER BY CP05"
                       CheckOC2
                       adoRecordset1.CursorLocation = adUseClient
                       adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                       If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                           adoRecordset1.MoveLast
                           'Modify By Cheng 2002/04/12
   '                        strSQL = "SELECT CP09,CP05,CP27 FROM CASEPROGRESS WHERE CP01='" & ChgSQL(strTemp(12)) & "' AND CP02='" & ChgSQL(strTemp(13)) & "' AND CP03='" & ChgSQL(strTemp(14)) & "' AND CP04='" & ChgSQL(strTemp(15)) & "' AND CP10='204' AND (SUBSTR(CP09,1,1)='B' OR SUBSTR(CP09,1,1)='A') AND CP05>=" & adoRecordset1.Fields(0) & " ORDER BY CP05 "
                           strSql = "SELECT CP09,CP05,CP27 FROM CASEPROGRESS WHERE CP01='" & ChgSQL(strTemp(12)) & "' AND CP02='" & ChgSQL(strTemp(13)) & "' AND CP03='" & ChgSQL(strTemp(14)) & "' AND CP04='" & ChgSQL(strTemp(15)) & "' AND CP10='204' AND ( CP09<'C' ) AND CP05>=" & adoRecordset1.Fields(0) & " ORDER BY CP05 "
                           CheckOC2
                           adoRecordset1.CursorLocation = adUseClient
                           adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                           If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                               adoRecordset1.MoveLast
                               If IsNull(adoRecordset1.Fields(2)) Then
                                   s = 1
                               End If
                           Else
                               s = 1
                           End If
                       Else
                           s = 1
                       End If
                       CheckOC2
                   End If
                   If s <> 1 Then
                        'Modify By Cheng 2002/04/12
   '                    strSQL = "SELECT CP09,CP05,CP27 FROM CASEPROGRESS WHERE CP01='" & ChgSQL(strTemp(12)) & "' AND CP02='" & ChgSQL(strTemp(13)) & "' AND CP03='" & ChgSQL(strTemp(14)) & "' AND CP04='" & ChgSQL(strTemp(15)) & "' AND CP10='1905' AND SUBSTR(CP09,1,1)='C' ORDER BY CP05 "
                       strSql = "SELECT CP09,CP05,CP27 FROM CASEPROGRESS WHERE CP01='" & ChgSQL(strTemp(12)) & "' AND CP02='" & ChgSQL(strTemp(13)) & "' AND CP03='" & ChgSQL(strTemp(14)) & "' AND CP04='" & ChgSQL(strTemp(15)) & "' AND CP10='1905' AND CP09>'C' ORDER BY CP05 "
                       CheckOC2
                       adoRecordset1.CursorLocation = adUseClient
                       adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                       If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                           If DateDiff("M", GetTaiwanTodayDate, CheckStr(adoRecordset1.Fields(1))) < 3 Then
                               s = 1
                           End If
                       End If
                   End If
                   CheckOC2
               End If
            End If
            If s = 0 Then
               '列印定稿
               If txt1(2) = "1" Then
                  '910624 Sieh
                  strCaseType = "16"
                  If txt1(1).Text = "1" Then
                     '催實審
                     strCaseSitu = "02"
                  Else
                     strCaseSitu = "01"
                  End If
                  EndLetter strCaseType, .Fields("CP09"), strCaseSitu, strUserNum
                  
                  If strCaseSitu = "01" Then
                     strExc(0) = "SELECT DECODE(CP05,'','',SUBSTR(CP05,1,4)-1911||'年'||SUBSTR(CP05,5,2)||'月'||SUBSTR(CP05,7,2))||'日'||CP08 FROM CASEPROGRESS WHERE CP43='" & .Fields("CP09") & "' AND CP10='1204'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                     
                     If intI = 1 Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & strCaseType & "','" & .Fields("CP09") & "','" & strCaseSitu & "','" & _
                           strUserNum & "','實審通知書','" & RsTemp.Fields(0) & "')"
                        cnnConnection.Execute strSql
                     End If
                  End If
                  strSysDate = Format(ServerDate)
                  'str((Val(Mid(strSysDate, 1, 4)) * 12 + Val(Mid(strSysDate, 5, 2))) - (Val(Mid(strCP27, 1, 4)) * 12 + Val(Mid(strCP27, 5, 2))))
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                     "('" & strCaseType & "','" & .Fields("CP09") & "','" & strCaseSitu & "','" & strUserNum & "','已發文時間','" & _
                     str((Val(Mid(strSysDate, 1, 4)) * 12 + Val(Mid(strSysDate, 5, 2))) - (Val(Mid(.Fields("CP27"), 1, 4)) * 12 + Val(Mid(.Fields("CP27"), 5, 2)))) & "')"
                  cnnConnection.Execute strSql
                  
                  'Modify by Amy 2014/08/15 +strLetterRecNo
                  NowPrint .Fields("CP09"), strCaseType, strCaseSitu, False, strUserNum, 0, , , , , , , , , , , , .Fields("cp09")
                  
               '列印報表
               Else
                  strTemp(0) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(0)))
                  strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                  cnnConnection.Execute "INSERT INTO R040307 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & strUserNum & "','" & CheckStr(.Fields(19)) & "') "
                  If Len(CheckStr(.Fields(1))) <> 8 Then
                      SavDay(0) = CheckStr(.Fields(1))
                  Else
                      SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
                  End If
                  If Len(CheckStr(.Fields(18))) <> 8 Then
                      SavDay(1) = CheckStr(.Fields(18))
                  Else
                      SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(18)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
                  End If
                  
                  'Remove by Morgan 2009/9/16 改由每日批次報表控制週期
                  ''Modify By Cheng 2002/08/07
                  'If blnUpdate Then
                  '   cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(19)) & "' AND NP07=" & Val(CheckStr(.Fields(17))) & " AND NP22=" & Val(CheckStr(.Fields(16)))
                  'End If
                  
               End If
               DataYN = "Y"
            End If
NextRecord:
            DoEvents
            .MoveNext
        Loop
    End With
    If DataYN = "N" Then ShowNoData
Else
    InsertQueryLog (0) 'Add By Sindy 2010/11/30
    ShowNoData
    Screen.MousePointer = vbDefault
    CheckOC
    Exit Sub
End If
'若催審函性質為催審函, 則列印
If txt1(2) <> "1" And DataYN = "Y" Then PrintData
Screen.MousePointer = vbDefault
If DataYN = "Y" Then ShowPrintOk
End Sub

Sub PrintData()
strSql = "SELECT * FROM R040307 WHERE ID='" & strUserNum & "' ORDER BY R024002,R024001,R024004 "
CheckOC
Page = 1
SavDay1 = " "
SavDay2 = " "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            For i = 0 To 9
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If strTemp(1) = SavDay2 Then
                strTemp(1) = ""
                If strTemp(0) = SavDay1 Then
                    strTemp(0) = ""
                Else
                    SavDay1 = strTemp(0)
                End If
            Else
                SavDay2 = strTemp(1)
                SavDay1 = strTemp(0)
            End If
            strTemp(4) = StrToStr(strTemp(4), 10)
            strTemp(5) = StrToStr(strTemp(5), 4)
            strTemp(6) = StrToStr(strTemp(6), 4)
            strTemp(7) = StrToStr(strTemp(7), 8)
            strTemp(8) = StrToStr(strTemp(8), 3)
            strTemp(9) = StrToStr(strTemp(9), 4)
            If iPrint > 10000 Then
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                Printer.NewPage
                Page = Page + 1
                PrintTitle
            End If
            PrintDatil
            .MoveNext
        Loop
    End With
End If

Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
Printer.EndDoc
CheckOC
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "內 專 催 審 表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
'Modify By Cheng 2002/08/22
'Printer.Print "催審期限：" & Format(ChangeTStringToTDateString(txt1(3)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(4))
Printer.Print IIf(Me.opt(0).Value, "催審期限：" & Format(ChangeTStringToTDateString(txt1(3)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(4)), _
               IIf(Me.opt(1).Value, "發文日期：" & Format(ChangeTStringToTDateString(txt1(11)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(12)), _
               "本所案號：" & Me.txt1(7).Text & "-" & Left(Me.txt1(8).Text & "000000", 6) & "-" & Left(Me.txt1(9).Text & "0", 1) & "-" & Left(Me.txt1(10).Text & "00", 2)))
               
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人　：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
If txt1(1) = "1" Then
    Printer.Print "催審性質：催實審"
Else
    Printer.Print "催審性質：催審"
End If
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "發文日"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "催審期限"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "申請案號/審定號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "申請國家"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "申請人"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Sub PrintDatil()
For i = 0 To 9
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1400 + 250
PLeft(2) = 2500 + 250
PLeft(3) = 4400 + 250
PLeft(4) = 6000 + 250
PLeft(5) = 8900 + 250
PLeft(6) = 10000 + 250
PLeft(7) = 11200 + 250
PLeft(8) = 13400 + 250
PLeft(9) = 14300 + 250
End Sub

Private Sub Form_Activate()
   'Add by Morgan 2009/9/18
   txt1(1) = "2"
   txt1(2) = "1"
   opt(2).Value = True
End Sub

Private Sub Form_Load()
MoveFormToCenter Me

'Add by Lydia 2018/01/03 開放外專程序人員操作FMP寰華案件。當非FMP寰華權限,不可看寰華案=>回傳SQL
FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05, "INVERSE_SQL")
If Pub_StrUserSt03 = "M51" Then
   If MsgBox("電腦中心人員請注意你現在是要看FMP寰華案嗎?", vbYesNo + vbInformation + vbDefaultButton1) = vbYes Then
      FMP2openSQL = Replace(FMP2openSQL, "not", "")
      FMP2open = True
   Else
      MsgBox "現在本報表不可查FMP寰華案"
   End If
End If
'end 2018/01/03

StrTest2 = "P,PS"
StrTest = StrTest2
'Added by Lydia 2018/01/03 因為外專程序沒有開P,PS權限
If FMP2open = True Then
    txt1(0).Locked = True
Else
'end 2018/01/03
    strTemp1 = Split(UCase(StrTest), ",")
    strTemp2 = Split(UCase(GetSystemKindByNick), ",")
    For i = 0 To UBound(strTemp1)
        s = 0
        For j = 0 To UBound(strTemp2)
            If strTemp2(j) = strTemp1(i) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            StrTest = Replace(StrTest, strTemp1(i), "")
        End If
    Next i
End If  'end 2018/01/03
txt1(0) = StrTest

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm040307 = Nothing
End Sub

Private Sub opt_Click(Index As Integer)
'Add By Cheng 2002/08/05
Select Case Index
Case 0 '催審期限
   Me.opt(0).Value = True
   Me.txt1(3).Enabled = True
   Me.txt1(4).Enabled = True
   Me.txt1(3).SetFocus
   
   Me.opt(1).Value = False
   Me.txt1(11).Enabled = False
   Me.txt1(12).Enabled = False
   
   Me.opt(2).Value = False
   Me.txt1(7).Enabled = False
   Me.txt1(8).Enabled = False
   Me.txt1(9).Enabled = False
   Me.txt1(10).Enabled = False
   
Case 1 '發文日期
   Me.opt(0).Value = False
   Me.txt1(3).Enabled = False
   Me.txt1(4).Enabled = False
   
   Me.opt(1).Value = True
   Me.txt1(11).Enabled = True
   Me.txt1(12).Enabled = True
   Me.txt1(11).SetFocus
   
   Me.opt(2).Value = False
   Me.txt1(7).Enabled = False
   Me.txt1(8).Enabled = False
   Me.txt1(9).Enabled = False
   Me.txt1(10).Enabled = False

Case 2 '本所案號
   Me.opt(0).Value = False
   Me.txt1(3).Enabled = False
   Me.txt1(4).Enabled = False
   
   Me.opt(1).Value = False
   Me.txt1(11).Enabled = False
   Me.txt1(12).Enabled = False
   
   Me.opt(2).Value = True
   Me.txt1(7).Enabled = True
   Me.txt1(8).Enabled = True
   Me.txt1(9).Enabled = True
   Me.txt1(10).Enabled = True
   Me.txt1(7).SetFocus

End Select
End Sub

'Add by Amy 2014/09/09 選定稿只能輸本所案號
Private Sub txt1_Change(Index As Integer)
    If Index <> 2 Then Exit Sub
    
    If txt1(2) = "1" Then
        opt(0).Enabled = False
        opt(1).Enabled = False
        opt(2).Enabled = True
        opt_Click (2)
    Else
        opt(0).Enabled = True
        opt(1).Enabled = True
        opt(2).Enabled = True
        opt_Click (0)
    End If
    
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
'Add By Cheng 2002/09/11
Select Case Index
Case 1 '催審函性質
   If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
Case 2 '列印別
   If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
      'Added by Lydia 2018/01/03 因為外專程序沒有開P,PS權限
      If FMP2open = True Then
      Else
      'end 2018/01/03
        strTemp1 = Split(GetSystemKindByNick, ",")
        strTemp2 = Split(txt1(0), ",")
        For i = 0 To UBound(strTemp2)
          s = 0
          For j = 0 To UBound(strTemp1)
              If strTemp1(j) = strTemp2(i) Then
                  s = 1
                  Exit For
              End If
          Next j
          If s = 0 Then
              s = MsgBox("USER : " & strUserNum & " 沒有 " & strTemp2(i) & " 的使用權限 ", , "USER 輸入錯誤")
              txt1(0).SetFocus
          End If
        Next i
      End If 'end 2018/01/03
Case 1
      'Modify By Cheng 2002/09/11
'     Select Case Val(txt1(Index))
'     Case 1, 2
'     Case Else
'          s = MsgBox("催審函性質只能1 或 2 !!", , "USER 輸入錯誤")
'          txt1(Index).SetFocus
'     End Select
Case 2
      'Modify By Cheng 2002/09/11
'     Select Case Val(txt1(Index))
'     Case 1, 2
'     Case Else
'          s = MsgBox("列印別只能1 或 2 !!", , "USER 輸入錯誤")
'          txt1(Index).SetFocus
'     End Select
'Modify By Cheng 2002/08/05
'Case 4, 6
Case 4, 6, 12
      If blnClkSure = False Then
         If RunNick(txt1(Index - 1), txt1(Index)) Then
            txt1(Index - 1).SetFocus
            txt1_GotFocus Index - 1
            Exit Sub
         End If
      Else
         blnClkSure = False
      End If
Case 7
   'Modify By Cheng 2002/09/11
   If Me.txt1(7).Text <> "" Then
      Select Case UCase(txt1(7))
      Case "P", "PS"
      Case Else
           s = MsgBox("系統類別只能 P 或 PS !!", , "USER 輸入錯誤")
           txt1(7).SetFocus
      End Select
   End If
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
'Add By Cheng 2002/08/05
Case 1 '催審函性質
   Select Case Val(txt1(Index))
   Case 1, 2
   Case Else
      s = MsgBox("催審函性質只能1 或 2 !!", , "USER 輸入錯誤")
      Cancel = True
      txt1(Index).SetFocus
      txt1_GotFocus Index
   End Select
Case 2 '列印別
   Select Case Val(txt1(Index))
   Case 1, 2
   Case Else
      s = MsgBox("列印別只能1 或 2 !!", , "USER 輸入錯誤")
      Cancel = True
      txt1(Index).SetFocus
      txt1_GotFocus Index
   End Select

Case 3, 4 '催審期限起, 迄
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
   
'Add By Cheng 2002/08/05
Case 11, 12 '發文日期起, 迄
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
   
End Select
End Sub
