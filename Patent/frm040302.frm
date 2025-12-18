VERSION 5.00
Begin VB.Form frm040302 
   BorderStyle     =   1  '單線固定
   Caption         =   "公告期滿通知函"
   ClientHeight    =   2136
   ClientLeft      =   1956
   ClientTop       =   4260
   ClientWidth     =   5064
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2136
   ScaleWidth      =   5064
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1560
      MaxLength       =   7
      TabIndex        =   0
      Top             =   456
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   5
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "1"
      Top             =   810
      Width           =   405
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   4
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   5
      Top             =   1275
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1275
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1275
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "P"
      Top             =   1275
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   1320
      Width           =   1395
   End
   Begin VB.OptionButton Option1 
      Caption         =   "公告期滿日："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   9
      Top             =   480
      Value           =   -1  'True
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4020
      TabIndex        =   8
      Top             =   20
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3195
      TabIndex        =   7
      Top             =   20
      Width           =   800
   End
   Begin VB.Label lbl 
      Caption         =   "列印種類：              (1.當期公告資料 2.歷史未領證資料)"
      Height          =   225
      Left            =   630
      TabIndex        =   10
      Top             =   840
      Width           =   4305
   End
End
Attribute VB_Name = "frm040302"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
'Modified by Morgan 2018/5/14
'Dim intWhere As Integer, PLeft(0 To 2) As Integer, strReceiveNo As String
Dim intWhere As Integer, PLeft(0 To 3) As Integer, strReceiveNo As String

Private Type strPrint
   No As String
   Name As String
   Case As String
End Type

Dim sField() As strPrint
Dim Lprint As Integer

'Add By Cheng 20025/10/24
'列印公告中已收文領證案件明細表
Private Type strPrint1
   CPNo As String '本所案號
   CASENAME As String '案件名稱
   ReceiveDate As String '收文日期
   SaleName As String '智權人員
End Type
Dim sField1() As strPrint1
Dim Lprint1 As Integer

Const ET01 As String = "11"
'Add By Cheng 2002/10/25
Dim m_strPA09 As String
Dim m_strPA08 As String
'92.5.8 add by sonia
Dim m_strPA11 As String

'Add By Cheng 2003/01/14
Dim m_CP09 As String '收文號

Private Sub cmdok_Click(Index As Integer)
 Dim strTmp As String, rsTemp1 As New ADODB.Recordset, rsTemp2 As New ADODB.Recordset
'Add By Cheng 2002/10/24
Dim blnExist601 As Boolean
'Add By Cheng 2003/01/14
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   Select Case Index
      Case 0 '確定
         ' 設定滑鼠游標為等待狀態
         Screen.MousePointer = vbHourglass
         'Add By Cheng 2002/10/24
         '預設符合條件的所有案件皆無領證資料
         blnExist601 = False
         Lprint1 = 0
         '選擇公告期滿日
         If Option1(0).Value = True Then
            'Add By Cheng 2002/03/19
            If Text1(Index).Text = "" Then
               MsgBox "公告期滿日不得空白，請重新輸入 !", vbCritical
               Me.Text1(0).SetFocus
               Text1_GotFocus 0
               ' 設定滑鼠游標為預設值
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.Text1(0)) = -1 Then
               Me.Text1(0).SetFocus
               Text1_GotFocus 0
               ' 設定滑鼠游標為預設值
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            'Add By Cheng 2002/05/15
            If Len(Me.Text1(5).Text) <= 0 Then
               MsgBox "列印種類不得空白，請重新輸入 !", vbCritical
               Me.Text1(5).SetFocus
               Text1_GotFocus 5
               ' 設定滑鼠游標為預設值
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            
            ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/29 清除查詢印表記錄檔欄位
            'Modify By Cheng 2002/05/15
'            strExc(0) = "SELECT PA01||PA02||PA03||PA04,TPB08,PA16," & ChgPatent("", 1) & ",NVL(PA06,NVL(PA07,PA05)),PA11 FROM PATENT,TPBULLETIN WHERE PA01='P' AND PA09='" & 台灣國家代號 & _
'               "' AND PA14<=" & TransDate(Text1(0).Text, 2) & " AND (PA57<>'Y' OR PA57 IS NULL) AND PA24 IS NULL AND PA20 IS NOT NULL AND PA23='1' AND PA11=TPB01(+) ORDER BY PA14,PA01,PA02,PA03,PA04"
            'Modify By Cheng 2002/05/20
'            strExc(0) = "SELECT PA01||PA02||PA03||PA04,TPB08,PA16," & ChgPatent("", 1) & ",NVL(PA06,NVL(PA07,PA05)),PA11 FROM PATENT,TPBULLETIN " & _
'                        " WHERE PA01='P' AND PA09='" & 台灣國家代號 & "' " & _
'                        " AND PA14" & IIf(Me.Text1(5).Text = "1", "=", "<") & TransDate(Text1(0).Text, 2) & " AND PA16='1' AND PA20 IS NOT NULL AND (PA24 IS NULL AND PA25 IS NULL) AND (PA57<>'Y' OR PA57 IS NULL) AND PA11=TPB01(+) " & _
'                        " ORDER BY PA14,PA01,PA02,PA03,PA04"
            '列印當期公告資料
            If Me.Text1(5).Text = "1" Then
                pub_QL05 = pub_QL05 & ";" & Left(Lbl, 5) & "1.當期公告資料" 'Add By Sindy 2010/11/29
                'Modify By Cheng 2002/10/25
'               strExc(0) = "SELECT PA01||PA02||PA03||PA04,TPB08,PA16," & ChgPatent("", 1) & ",NVL(PA06,NVL(PA07,PA05)),PA11 FROM PATENT,TPBULLETIN " & _
'                           " WHERE PA01='P' AND PA09='" & 台灣國家代號 & "' " & _
'                           " AND PA14=" & TransDate(Text1(0).Text, 2) & " AND PA16='1' AND PA20 IS NOT NULL AND (PA24 IS NULL AND PA25 IS NULL) AND (PA57<>'Y' OR PA57 IS NULL) AND PA11=TPB01(+) " & _
'                           " ORDER BY PA14,PA01,PA02,PA03,PA04"
                'Modify By Cheng 2002/11/26
                '以輸入的公告期滿日減三個月抓資料
'               strExc(0) = "SELECT PA01||PA02||PA03||PA04,TPB08,PA16," & ChgPatent("", 1) & ",NVL(PA06,NVL(PA07,PA05)),PA11,PA01,PA02,PA03,PA04,PA09,PA08 FROM PATENT,TPBULLETIN " & _
'                           " WHERE PA01='P' AND PA09='" & 台灣國家代號 & "' " & _
'                           " AND PA14=" & TransDate(Text1(0).Text, 2) & " AND PA16='1' AND PA20 IS NOT NULL AND (PA24 IS NULL AND PA25 IS NULL) AND (PA57<>'Y' OR PA57 IS NULL) AND PA11=TPB01(+) " & _
'                           " ORDER BY PA14,PA01,PA02,PA03,PA04"
               strExc(0) = "SELECT PA01||PA02||PA03||PA04,TPB08,PA16," & ChgPatent("", 1) & ",NVL(PA06,NVL(PA07,PA05)),PA11,PA01,PA02,PA03,PA04,PA09,PA08 FROM PATENT,TPBULLETIN " & _
                           " WHERE PA01='P' AND PA09='" & 台灣國家代號 & "' " & _
                           " AND PA14=" & DBDATE(DateAdd("M", -3, ChangeTStringToWDateString(Text1(0).Text))) & " AND PA16='1' AND PA20 IS NOT NULL AND (PA24 IS NULL AND PA25 IS NULL) AND (PA57<>'Y' OR PA57 IS NULL) AND PA11=TPB01(+) " & _
                           " ORDER BY PA14,PA01,PA02,PA03,PA04"
            '列印歷史未領證資料
            Else
               pub_QL05 = pub_QL05 & ";" & Left(Lbl, 5) & "2.歷史未領證資料" 'Add By Sindy 2010/11/29
               'Modify By Cheng 2002/05/21
'               strExc(0) = "SELECT PA01||PA02||PA03||PA04,TPB08,PA16," & ChgPatent("", 1) & ",NVL(PA06,NVL(PA07,PA05)),PA11,PA14,DECODE(SUM(DECODE(CP24,'',0,1)),COUNT(*),0,1) FROM CASEPROGRESS,PATENT,TPBULLETIN " & _
'                           " WHERE PA01='P' AND PA09='" & 台灣國家代號 & "' " & _
'                           " AND PA14<" & TransDate(Text1(0).Text, 2) & " AND PA16='1' AND PA20 IS NOT NULL AND (PA24 IS NULL AND PA25 IS NULL) AND (PA57<>'Y' OR PA57 IS NULL) AND PA11=TPB01(+) " & _
'                           " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP10='" & 異議答辯 & "' " & _
'                           " GROUP BY PA01||PA02||PA03||PA04,TPB08,PA16," & ChgPatent("", 1) & ",NVL(PA06,NVL(PA07,PA05)),PA11,PA14 " & _
'                           " ORDER BY PA14,PA01||PA02||PA03||PA04"
                'Modify By Cheng 2002/11/26
                '以輸入的公告期滿日減三個月抓資料
'               strExc(0) = "SELECT PA14," & ChgPatent("", 1) & ",PA16,NVL(PA06,NVL(PA07,PA05)),PA11,DECODE(SUM(DECODE(CP24,'',0,1)),COUNT(*),0,1) " & _
'                           " FROM CASEPROGRESS,PATENT " & _
'                           " WHERE PA01='P' AND PA09='" & 台灣國家代號 & "' " & _
'                           " AND PA14<" & TransDate(Text1(0).Text, 2) & " AND PA16='1' AND PA20 IS NOT NULL AND (PA24 IS NULL AND PA25 IS NULL) AND (PA57<>'Y' OR PA57 IS NULL) " & _
'                           " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP10='" & 異議答辯 & "' " & _
'                           " GROUP BY PA14," & ChgPatent("", 1) & ",PA16,NVL(PA06,NVL(PA07,PA05)),PA11 " & _
'                           " ORDER BY PA14," & ChgPatent("", 1)
               strExc(0) = "SELECT PA14," & ChgPatent("", 1) & ",PA16,NVL(PA06,NVL(PA07,PA05)),PA11,DECODE(SUM(DECODE(CP24,'',0,1)),COUNT(*),0,1) " & _
                           " FROM CASEPROGRESS,PATENT " & _
                           " WHERE PA01='P' AND PA09='" & 台灣國家代號 & "' " & _
                           " AND PA14<" & DBDATE(DateAdd("M", -3, ChangeTStringToWDateString(Text1(0).Text))) & " AND PA16='1' AND PA20 IS NOT NULL AND (PA24 IS NULL AND PA25 IS NULL) AND (PA57<>'Y' OR PA57 IS NULL) " & _
                           " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP10='" & 異議答辯 & "' " & _
                           " GROUP BY PA14," & ChgPatent("", 1) & ",PA16,NVL(PA06,NVL(PA07,PA05)),PA11 " & _
                           " ORDER BY PA14," & ChgPatent("", 1)
            
            End If
            pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & Text1(0) 'Add By Sindy 2010/11/29
            
            intI = 0
            Set rsTemp2 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            'Modify By Cheng 2002/05/20
'            If intI = 1 Then
            '列印種類為當期公告資料
            If intI = 1 And Me.Text1(5).Text = "1" Then
               With rsTemp2
                  Lprint = 0
                  InsertQueryLog (.RecordCount) 'Add By Sindy 2010/11/29
                  Do While Not .EOF
                     intI = 1
                     'Add By Cheng 2002/05/15
                     '列印種類為當期公告資料
                     If Me.Text1(5).Text = "1" Then
                        strExc(0) = "SELECT COUNT(*) FROM NEXTPROGRESS WHERE " & ChgNextProgress(.Fields(0)) & " AND NP07='" & 異議答辯 & "'"
                        Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                        '若下一程序無資料
                        If rsTemp1.Fields(0) = 0 Then
                            'Add/Modify By Cheng 2002/10/24
                           '若此案號案件進度檔有領證(601)資料
                           If Exist601("" & .Fields(0).Value) Then
                                blnExist601 = True
                           '若此案號案件進度檔無領證(601)資料
                           Else
                                'Add By Cheng 2002/10/25
                                m_strPA09 = "" & .Fields("PA09").Value
                                m_strPA08 = "" & .Fields("PA08").Value
                                m_strPA11 = "" & .Fields("PA11").Value
                                strReceiveNo = .Fields(0)
                                'Add By Cheng 2003/01/14
                                '抓收文號
                                m_CP09 = ""
                                StrSQLa = "Select CP09 From CaseProgress Where " & ChgCaseprogress(strReceiveNo) & " And CP09 < 'B' Order By CP05 Desc, CP09 Desc "
                                rsA.CursorLocation = adUseClient
                                rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                                If rsA.RecordCount > 0 Then
                                    m_CP09 = "" & rsA.Fields(0).Value
                                    '92.5.7 MODIFY BY SONIA
                                    'StartLetter ET01, "01"
                                    'NowPrint m_CP09, ET01, "01", False, strUserNum, 0
                                    'Modify by Morgan 2010/12/27 申請案號改碼數
                                    'If Len(m_strPA11) < 9 Then
                                    If Len(m_strPA11) < 10 Then
                                       StartLetter ET01, "01"
                                       NowPrint m_CP09, ET01, "01", False, strUserNum, 0
                                    Else
                                       'Modify by Morgan 2010/12/28 申請案號改碼數
                                       'If Mid(m_strPA11, 3, 1) <> "3" Then
                                       If Mid(m_strPA11, 4, 1) <> "3" Then
                                          StartLetter ET01, "02"
                                          NowPrint m_CP09, ET01, "02", False, strUserNum, 0
                                       Else
                                          StartLetter ET01, "03"
                                          NowPrint m_CP09, ET01, "03", False, strUserNum, 0
                                       End If
                                    End If
                                    '92.5.7 END
                                End If
                                If rsA.State <> adStateClosed Then rsA.Close
                                Set rsA = Nothing
                                'Modify By Cheng 2003/01/14
                                '移至抓到收文號時出定稿
'                                StartLetter ET01, "01"
'                                NowPrint .Fields(0) & "&000", ET01, "01", False, strUserNum, 0
                                 'Add By Cheng 2002/10/24
                                 '列印接洽結案單
                                g_PrtForm001.PrintFormCP "" & .Fields("PA01").Value, "" & .Fields("PA02").Value, "" & .Fields("PA03").Value, "" & .Fields("PA04").Value, "" & m_strPA11
                                'Modify By Cheng 2002/11/26
                                GoTo NextRecord
                            End If
                            'Modify By Cheng 2002/11/26
'                           GoTo NextRecord
                        End If
                     End If
                     
                     strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE " & ChgCaseprogress(.Fields(0)) & " AND CP10='" & 異議答辯 & "'"
                     Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                    '若案件進度檔無資料
                     If rsTemp1.Fields(0) = 0 Then
                        'Modify By Cheng 2002/05/15
'                        strReceiveNo = .Fields(0)
'                        ' 90.08.22 modify by louis
'                        'StartLetter ET01, "00"
'                        'NowPrint .Fields(0) & "&000", ET01, "00", False, strUserNum, 0
'                        StartLetter ET01, "01"
'                        NowPrint .Fields(0) & "&000", ET01, "01", False, strUserNum, 0
                        
                        '不印通知函
                    '若案件進度檔有資料
                     Else
                        'Modify By Cheng 2002/05/15
'                        intI = 1
'                        strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE " & ChgCaseprogress(.Fields(0)) & " AND CP10='" & 異議答辯 & "'"
'                        Set rsTemp1 = clslawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'                        If rsTemp1.Fields(0) = 0 Then
'                           '不印通知函
'                        Else
                           intI = 1
                           strExc(0) = "SELECT DECODE(SUM(DECODE(CP24,'',0,1)),COUNT(*),0,1) FROM CASEPROGRESS WHERE " & ChgCaseprogress(.Fields(0)) & " AND CP10='" & 異議答辯 & "'"
                           Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                           '若有無結果的資料
                           If rsTemp1.Fields(0) = 1 Then
                              '不印通知函
                           '若資料皆有結果
                           Else
                              '若專利基本檔的目前准駁欄為"准"
                              If .Fields(2) = "1" Then
                                 ReDim Preserve sField(Lprint + 1)
                                 If Not IsNull(.Fields(3)) Then sField(Lprint).No = .Fields(3)
                                 If Not IsNull(.Fields(4)) Then sField(Lprint).Name = .Fields(4)
                                 If Not IsNull(.Fields(5)) Then sField(Lprint).Case = .Fields(5)
                                 Lprint = Lprint + 1
                              End If
                           End If
'                        End If
                     End If
NextRecord:
                     .MoveNext
                  Loop
               End With
               If Lprint > 0 Then PrintCase
               MsgBox "列印結束 !", vbInformation
               ' 清除暫存陣列
               If Lprint > 0 Then
                  Erase sField
                  Lprint = 0
               End If
            
            'Add By Cheng 2002/05/20
            '列印種類為歷史未領證資料
            ElseIf intI = 1 And Me.Text1(5).Text = "2" Then
               Lprint = 0
               InsertQueryLog (rsTemp2.RecordCount) 'Add By Sindy 2010/11/29
               Do While Not rsTemp2.EOF
                  '若有無結果的資料
                  If rsTemp2.Fields(5) = 1 Then
                     '不印通知函
                  '若資料皆有結果
                  Else
                     '若專利基本檔的目前准駁欄為"准"
                     If rsTemp2.Fields(2) = "1" Then
                        ReDim Preserve sField(Lprint + 1)
                        If Not IsNull(rsTemp2.Fields(1)) Then sField(Lprint).No = rsTemp2.Fields(1)
                        If Not IsNull(rsTemp2.Fields(3)) Then sField(Lprint).Name = rsTemp2.Fields(3)
                        If Not IsNull(rsTemp2.Fields(4)) Then sField(Lprint).Case = rsTemp2.Fields(4)
                        Lprint = Lprint + 1
                     End If
                  End If
                  rsTemp2.MoveNext
               Loop
            
               If Lprint > 0 Then PrintCase
               'Add By Cheng 2002/10/25
               PrintCase1
               MsgBox "列印結束 !", vbInformation
               ' 清除暫存陣列
               If Lprint > 0 Then
                  Erase sField
                  Lprint = 0
               End If
                'Add By Cheng 2002/10/25
                Erase sField1
                Lprint1 = 0
            
            Else
               InsertQueryLog (0) 'Add By Sindy 2010/11/29
               'Modify By Cheng 2002/05/15
'               MsgBox "無符合條件之資料可列印 !", vbInformation
            End If
         
         '選擇本所案號
         Else
            strTmp = Text1(1) & Text1(2)
            pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & Text1(1) & "-" & Text1(2) 'Add By Sindy 2010/11/29
            If Text1(3).Text = "" Then
               strTmp = strTmp & "0"
               pub_QL05 = pub_QL05 & "-0" 'Add By Sindy 2010/11/29
            Else
               strTmp = strTmp & Text1(3).Text
               pub_QL05 = pub_QL05 & "-" & Text1(3) 'Add By Sindy 2010/11/29
            End If
            If Text1(4).Text = "" Then
               strTmp = strTmp & "00"
               pub_QL05 = pub_QL05 & "-00" 'Add By Sindy 2010/11/29
            Else
               strTmp = strTmp & Text1(4).Text
               pub_QL05 = pub_QL05 & "-" & Text1(4) 'Add By Sindy 2010/11/29
            End If
            
            strReceiveNo = strTmp
            'Modify By Cheng 2002/05/15
'            strExc(0) = "SELECT PA09 FROM PATENT WHERE " & ChgPatent(strReceiveNo) & " AND PA09='" & 台灣國家代號 & _
'                        "' AND (PA57<>'Y' OR PA57 IS NULL) AND PA21 IS NULL"
            'Modify By Cheng 2002/10/25
'            strExc(0) = "SELECT PA09 FROM PATENT WHERE " & ChgPatent(strReceiveNo) & " AND PA01='P' AND PA09='" & 台灣國家代號 & _
'                        "' AND PA16='1' AND PA20 IS NOT NULL AND (PA24 IS NULL AND PA25 IS NULL) AND (PA57<>'Y' OR PA57 IS NULL) "
            strExc(0) = "SELECT PA09,PA08,PA11 FROM PATENT WHERE " & ChgPatent(strReceiveNo) & " AND PA01='P' AND PA09='" & 台灣國家代號 & _
                        "' AND PA16='1' AND PA20 IS NOT NULL AND (PA24 IS NULL AND PA25 IS NULL) AND (PA57<>'Y' OR PA57 IS NULL) "
            
            'Modify By Cheng 2002/05/15
'            intI = 1
            intI = 0
            'Add By Cheng 2002/05/15
            Set rsTemp2 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            
            If intI = 1 Then
                InsertQueryLog (rsTemp2.RecordCount) 'Add By Sindy 2010/11/29
                '若此案號案件進度檔有領證(601)資料
                If Exist601("" & strReceiveNo) Then
                     blnExist601 = True
                '若此案號案件進度檔無領證(601)資料
                Else
                    'Add By Cheng 2002/10/25
                    m_strPA09 = "" & rsTemp2.Fields("PA09").Value
                    m_strPA08 = "" & rsTemp2.Fields("PA08").Value
                    m_strPA11 = "" & rsTemp2.Fields("PA11").Value
                    'Add By Cheng 2003/01/14
                    '抓收文號
                    m_CP09 = ""
                    StrSQLa = "Select CP09 From CaseProgress Where " & ChgCaseprogress(strReceiveNo) & " And CP09 < 'B' Order By CP05 Desc, CP09 Desc "
                    rsA.CursorLocation = adUseClient
                    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                    If rsA.RecordCount > 0 Then
                        m_CP09 = "" & rsA.Fields(0).Value
                        '92.5.7 MODIFY BY SONIA
                        'StartLetter ET01, "01"
                        'NowPrint m_CP09, ET01, "01", False, strUserNum, 0
                        'Modify by Morgan 2010/12/28 申請案號改碼數
                        'If Len(m_strPA11) < 9 Then
                        If Len(m_strPA11) < 10 Then
                           StartLetter ET01, "01"
                           NowPrint m_CP09, ET01, "01", False, strUserNum, 0
                        Else
                           'Modify by Morgan 2010/12/28 申請案號改碼數
                           'If Mid(m_strPA11, 3, 1) <> "3" Then
                           If Mid(m_strPA11, 4, 1) <> "3" Then
                              StartLetter ET01, "02"
                              NowPrint m_CP09, ET01, "02", False, strUserNum, 0
                           Else
                              StartLetter ET01, "03"
                              NowPrint m_CP09, ET01, "03", False, strUserNum, 0
                           End If
                        End If
                        '92.5.7 END
                    End If
                    If rsA.State <> adStateClosed Then rsA.Close
                    Set rsA = Nothing
                    
                    ' 90.08.22 modify by louis
                    'StartLetter ET01, "00"
                    'NowPrint strReceiveNo & "&000", ET01, "00", False, strUserNum, 0
                    'Modify By Cheng 2003/01/14
                    '移至抓到總收文號時再出定稿
'                    StartLetter ET01, "01"
'                    NowPrint strReceiveNo & "&000", ET01, "01", False, strUserNum, 0
                      'Add By Cheng 2002/10/24
                      '列印接洽結案單
                     g_PrtForm001.PrintFormCP "" & Me.Text1(1).Text, "" & Me.Text1(2).Text, "" & Me.Text1(3).Text, "" & Me.Text1(4).Text, "" & m_strPA11
                End If
                'Add By Cheng 2002/10/25
                Erase sField1
                Lprint1 = 0
            
            Else
               'Modify By Cheng 2002/05/15
'               MsgBox "無符合條件之資料可列印 !", vbInformation
               InsertQueryLog (0) 'Add By Sindy 2010/11/29
               MsgBox "此案號不符合列印公告期滿的條件 !", vbInformation
            End If
         End If
         ' 設定滑鼠游標為預設
         Screen.MousePointer = vbDefault
      Case 1 '結束
         Unload Me
   End Select
End Sub

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
Dim strTxt(1 To 5) As String, i As Integer, j As Integer, strTmp As String
Dim DATE1 As String, DATE2 As String
'Add By Cheng 2003/01/14
Dim ii As Integer
   
   ii = 0
   EndLetter ET01, m_CP09, ET03, strUserNum
   strExc(0) = "SELECT NP08,NP09 FROM NEXTPROGRESS WHERE " & ChgNextProgress(strReceiveNo) & " AND NP07=" & 領證及繳年費 & " AND NP06 IS NULL ORDER BY NP08"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "','本所期限'," & CNULL(RsTemp.Fields(0)) & ")"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "','法定期限'," & CNULL(RsTemp.Fields(1)) & ")"
   End If
   'Add By Cheng 2002/10/25
   ii = ii + 1
   If Len(m_strPA11) < 9 Then '92.5.7 ADD BY SONIA
        'Modify By Cheng 2003/01/14
       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
          "('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "','服務費','" & Val(PUB_GetYF06(m_strPA09, m_strPA08, "Y00000001", "601", "1", "1")) & "')"
         ii = ii + 1
        'Modify By Cheng 2003/01/14
       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
          "('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "','規費','" & Val(PUB_GetYF07(m_strPA09, m_strPA08, "Y00000001", "601", "1", "1")) & "')"
         ii = ii + 1
        'Modify By Cheng 2003/01/14
       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
          "('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "','費用','" & Val(PUB_GetYF0607(m_strPA09, m_strPA08, "Y00000001", "601", "1", "1")) & "')"
   '92.5.7 ADD BY SONIA
   Else
      'STRTMP = GetNation '抓專用年度
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "','專用年度','" & Val(GetNation) & "')"
      ii = ii + 1
      'Modify by Morgan 2010/12/28 申請案號改碼數
      'If Mid(m_strPA11, 3, 1) <> "3" Then
      If Mid(m_strPA11, 4, 1) <> "3" Then
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "','服務費','" & Val(PUB_GetYF06(m_strPA09, m_strPA08, "Y00000001", "602", "1", "1")) & "')"
      Else
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "','服務費','" & Val(PUB_GetYF06(m_strPA09, m_strPA08, "Y00000001", "603", "1", "1")) & "')"
      End If
   End If
   '92.5.7 END
   
    If Not ClsLawExecSQL(ii, strTxt) Then 'edit by nickc 2007/02/05 不用 dll 了   If Not objLawDll.ExecSQL(ii, strTxt) Then
       MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
    End If

End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   Option1_Click 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040302 = Nothing
End Sub

Private Sub PrintCase()
 Dim i As Integer, Page As Integer, iPrint As Integer
On Error GoTo HndErr
   GetPrintLeft
   Page = 1
   CaseTitle Page
   iPrint = 2700
   For i = 0 To Lprint - 1
      Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
      Printer.Print sField(i).No
      Printer.CurrentX = PLeft(1):      Printer.CurrentY = iPrint
      Printer.Print sField(i).Case
      Printer.CurrentX = PLeft(2):      Printer.CurrentY = iPrint
      Printer.Print sField(i).Name
      'Add By Cheng 2002/05/20
      iPrint = iPrint + 300
      
      If i < Lprint Then
         If (i Mod 37 = 0 And i <> 0) Then
            Printer.NewPage
            Page = Page + 1
            CaseTitle Page
            iPrint = 2700
         End If
         'Modify By Cheng 2002/05/20
'         iPrint = iPrint + 300
      End If
   Next
   Printer.EndDoc
   Exit Sub
HndErr:
   MsgBox Err.Description
End Sub

'Add By Cheng 2002/10/25
Private Sub PrintCase1()
Dim i As Integer, Page As Integer, iPrint As Integer
On Error GoTo HndErr
    GetPrintLeft1
    Page = 1
    CaseTitle1 Page
    iPrint = 2700
    If Lprint1 = 0 Then
            Printer.CurrentX = PLeft(1) + 2000:    Printer.CurrentY = iPrint + 600
            Printer.Print "無已收文領證案件"
    Else
        For i = 1 To Lprint1
            Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
            Printer.Print sField1(i).CPNo
            Printer.CurrentX = PLeft(1):      Printer.CurrentY = iPrint
            Printer.Print sField1(i).CASENAME
            Printer.CurrentX = PLeft(2):      Printer.CurrentY = iPrint
            Printer.Print sField1(i).ReceiveDate
            Printer.CurrentX = PLeft(3):      Printer.CurrentY = iPrint
            Printer.Print sField1(i).SaleName
            iPrint = iPrint + 300
            
            If i < Lprint Then
                If (i Mod 20 = 0 And i <> 0) Then
                    Printer.NewPage
                    Page = Page + 1
                    CaseTitle1 Page
                    iPrint = 2700
                End If
            End If
        Next
    End If
    Printer.EndDoc
    Exit Sub
HndErr:
    MsgBox Err.Description
End Sub

Private Sub GetPrintLeft()
   PLeft(0) = 500:     PLeft(1) = 3000
   PLeft(2) = 4500
End Sub

'Add By Cheng 2002/10/25
Private Sub GetPrintLeft1()
   PLeft(0) = 500:     PLeft(1) = 3000
   PLeft(2) = 8000:     PLeft(3) = 10000
End Sub

Private Sub CaseTitle(ByVal Page As String)
 Dim i As Integer
   i = 500
   Printer.Font.Size = 22
   Printer.Font.Name = "細明體"
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4000:         Printer.CurrentY = i
   Printer.Print "被異議不成立清單"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 4500:         Printer.CurrentY = i + 500
   'Modify By Cheng 2002/05/20
'   Printer.Print "公告日 : " & Text1(0).Text
   Printer.Print "公告期滿日" & IIf(Me.Text1(5).Text = "1", " = ", " < ") & Text1(0).Text
   Printer.Font.Bold = False
   Printer.CurrentX = 500:              Printer.CurrentY = i + 800
   Printer.Print "列印人 : " & strUserName
   'Modify By Cheng 2002/05/20
'   Printer.CurrentX = 9000:            Printer.CurrentY = i + 800
   Printer.CurrentX = 8500:            Printer.CurrentY = i + 800
   Printer.Print "列印日期 : " & Format(GetTaiwanTodayDate, "##/##/##")
   'Modify By Cheng 2002/05/20
'   Printer.CurrentX = 9000:            Printer.CurrentY = i + 1100
   Printer.CurrentX = 8500:            Printer.CurrentY = i + 1100
   Printer.Print "頁　　次 : " & Page
   Printer.CurrentX = 500:              Printer.CurrentY = i + 1400
   Printer.Print String(205, "-")
   Printer.CurrentX = PLeft(0):         Printer.CurrentY = i + 1700
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(1):         Printer.CurrentY = i + 1700
   Printer.Print "申請案號"
   Printer.CurrentX = PLeft(2):         Printer.CurrentY = i + 1700
   Printer.Print "案件名稱"
   Printer.CurrentX = 500:          Printer.CurrentY = i + 2000
   Printer.Print String(205, "-")
End Sub

'Add By Cheng 2002/10/25
Private Sub CaseTitle1(ByVal Page As String)
 Dim i As Integer
   i = 500
   Printer.Font.Size = 22
   Printer.Font.Name = "細明體"
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4000:         Printer.CurrentY = i
   Printer.Print "公告中已收文領證案件明細表"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 4500:         Printer.CurrentY = i + 500
'   Printer.Print "公告日" & IIf(Me.Text1(5).Text = "1", " = ", " < ") & Text1(0).Text
   Printer.Font.Bold = False
   Printer.CurrentX = 500:              Printer.CurrentY = i + 800
   Printer.Print "列印人 : " & strUserName
   'Modify By Cheng 2002/05/20
'   Printer.CurrentX = 9000:            Printer.CurrentY = i + 800
   Printer.CurrentX = 8500:            Printer.CurrentY = i + 800
   Printer.Print "列印日期 : " & Format(GetTaiwanTodayDate, "##/##/##")
   'Modify By Cheng 2002/05/20
'   Printer.CurrentX = 9000:            Printer.CurrentY = i + 1100
   Printer.CurrentX = 8500:            Printer.CurrentY = i + 1100
   Printer.Print "頁　　次 : " & Page
   Printer.CurrentX = 500:              Printer.CurrentY = i + 1400
   Printer.Print String(205, "-")
   Printer.CurrentX = PLeft(0):         Printer.CurrentY = i + 1700
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(1):         Printer.CurrentY = i + 1700
   Printer.Print "案　　件　　名　　稱"
   Printer.CurrentX = PLeft(2):         Printer.CurrentY = i + 1700
   Printer.Print "收文日期"
   Printer.CurrentX = PLeft(3):         Printer.CurrentY = i + 1700
   Printer.Print "智權人員"
   Printer.CurrentX = 500:          Printer.CurrentY = i + 2000
   Printer.Print String(205, "-")
End Sub

Private Sub Option1_Click(Index As Integer)
 Dim txt As TextBox, i As Integer
On Error Resume Next
   For Each txt In Text1
      txt.Enabled = False
   Next
   Select Case Index
      Case 0
         Text1(0).Enabled = True
         'Add By Cheng 2002/05/15
         Me.Text1(5).Enabled = True
      Case 1
         For i = 2 To 4
            Text1(i).Enabled = True
         Next
   End Select

End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If Index = 5 Then
      If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   '選擇公告日
   If Option1(0).Value = True Then
      If Index = 0 Then
         If Text1(Index).Text <> "" Then
            Cancel = Not ChkDate(Text1(Index).Text)
         Else
'            MsgBox "公告日不得空白，請重新輸入 !", vbCritical
'            Cancel = True
         End If
      End If
      'Add By Cheng 2002/05/15
      '檢查列印種類
      If Index = 5 Then
         If Len(Me.Text1(5).Text) <= 0 Then
            MsgBox "列印種類不得空白，請重新輸入 !", vbCritical
            Cancel = True
         End If
      End If
   '選擇本所案號
   Else
      If Index = 1 Then
         If Text1(Index).Text = "" Then
'            MsgBox "本所案號不得空白，請重新輸入 !", vbCritical
'            Cancel = True
         End If
      End If
   End If
   If Cancel Then TextInverse Text1(Index)
End Sub

'Add By Cheng 2002/10/24
'判斷某案號的案件進度檔中有無領證資料
Private Function Exist601(strCP0104) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

Exist601 = False
StrSQLa = "SELECT * FROM PATENT,CASEPROGRESS WHERE PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND " & ChgPatent(strCP0104) & " AND CP10='601'"
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'若有領證資料
If rsA.RecordCount > 0 Then
    Exist601 = True
    Lprint1 = Lprint1 + 1
    ReDim Preserve sField1(Lprint1)
    sField1(Lprint1).CPNo = strCP0104
    sField1(Lprint1).CASENAME = IIf("" & rsA.Fields("PA05").Value <> "", rsA.Fields("PA05").Value, IIf("" & rsA.Fields("PA06").Value <> "", rsA.Fields("PA06").Value, "" & rsA.Fields("PA07").Value))
    sField1(Lprint1).ReceiveDate = ChangeTStringToTDateString(rsA.Fields("CP05") - 19110000)
    sField1(Lprint1).SaleName = GetStaffName(rsA.Fields("CP13").Value)
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function
'92.5.8 add by sonia
Private Function GetNation() As String

   Select Case m_strPA08
      Case 1
         strExc(0) = "na07"
      Case 2
         strExc(0) = "na09"
      Case 3
         strExc(0) = "na11"
   End Select

   GetNation = ""
   strExc(0) = "SELECT " & strExc(0) & " FROM NATION WHERE NA01=" + CNULL(m_strPA09)
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp.Fields(0)) Then
         GetNation = RsTemp.Fields(0)
      End If
   End If
   
End Function
'92.5.8 end
