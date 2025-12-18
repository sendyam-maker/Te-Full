VERSION 5.00
Begin VB.Form frm060303 
   BorderStyle     =   1  '單線固定
   Caption         =   "公告期滿通知函"
   ClientHeight    =   3015
   ClientLeft      =   3585
   ClientTop       =   3735
   ClientWidth     =   3495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3495
   Begin VB.Frame Frame2 
      Caption         =   "設定請款單"
      Height          =   660
      Left            =   0
      TabIndex        =   15
      Top             =   2310
      Width           =   3435
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   8
         Top             =   240
         Width           =   2520
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   0
         Left            =   105
         TabIndex        =   16
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   5
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1230
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定地址條"
      Height          =   660
      Left            =   0
      TabIndex        =   12
      Top             =   1590
      Width           =   3435
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   7
         Top             =   240
         Width           =   2520
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   13
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1950
      TabIndex        =   9
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   2730
      TabIndex        =   10
      Top             =   60
      Width           =   756
   End
   Begin VB.OptionButton Option1 
      Caption         =   "公告日："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   11
      Top             =   564
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   900
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   0
      Left            =   1308
      MaxLength       =   7
      TabIndex        =   0
      Top             =   528
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   1
      Left            =   1308
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "FCP"
      Top             =   864
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   2
      Left            =   1788
      MaxLength       =   6
      TabIndex        =   3
      Top             =   864
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   3
      Left            =   2628
      MaxLength       =   1
      TabIndex        =   4
      Top             =   864
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   4
      Left            =   2868
      MaxLength       =   2
      TabIndex        =   5
      Top             =   864
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Index           =   0
      Left            =   330
      TabIndex        =   14
      Top             =   1290
      Width           =   900
   End
End
Attribute VB_Name = "frm060303"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit
Dim intWhere As Integer, PLeft(0 To 2) As Integer, strReceiveNo As String
Private Type strPrint
   No1 As String
   No As String
   Name As String
   Case As String
End Type
Dim sField() As strPrint
Dim Lprint As Integer
Const ET01 As String = "06"
'Add By Cheng 2003/01/28
Dim m_OriPrinterName As String, SeekPrint As Integer, SeekPrintL As Integer, j As Integer, i As Integer

Private Sub PrintPI(ByVal strTmp As String)
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String, A1K(1 To T_1K0) As String, lTmp As Long
Dim pa() As String, A1K() As String, lTmp As Long
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
ReDim A1K(1 To TF_1K0) As String

'Add By Cheng 2003/02/25
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   ChgCaseNo strTmp, pa
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      If CU72FA39(pa(26), pa(75)) Then
         If CU73FA40(pa(26), pa(75)) Then
            GoTo A0
         Else
            'Modify By Cheng 2003/01/30
'            A1K(1) = AutoNo("X", 6)
            A1K(1) = AccAutoNo(MsgText(815), 5)
            '更新流水號
            AccSaveAutoNo MsgText(815), Right(A1K(1), 5)
            A1K(2) = strSrvDate(2)
            'Modify By Cheng 2003/01/30
            '代理人-->申請人
'            If pa(26) = "" Then
'               A1K(3) = pa(75)
'            Else
'               A1K(3) = pa(26)
'            End If
'            If pa(75) <> "" Then
'               A1K(3) = pa(75)
'            Else
'               A1K(3) = pa(26)
'            End If
            A1K(3) = PUB_GetA1K03(pa(1), pa(2), pa(3), pa(4))
            
            'Add By Cheng 2003/01/30
            '折讓金額預存零
            A1K(6) = 0
            'Modify By Cheng 2003/01/30
            '取得領證及繳年費規費
'            'Modify By Cheng 2002/01/09
''            strExc(0) = "SELECT nvl(YF06,0) FROM PATENTYEARFEE WHERE YF01=" & CNULL(pa(9)) & " AND YF02=" & CNULL(pa(8)) & " AND " & _
''               "YF03='Y99999000' AND YF04='601' AND YF05=1"
'            strExc(0) = "SELECT nvl(YF06,0) FROM PATENTYEARFEE WHERE YF01=" & CNULL(pa(9)) & " AND YF02=" & CNULL(pa(8)) & " AND " & _
'               "YF03='Y00000000' AND YF04='601' AND YF05=1"
'            intI = 1
'            Set rsTemp = clslawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'            lTmp = 0
'            If intI = 1 Then lTmp = rsTemp.Fields(0)
'
'            'Modify By Cheng 2002/01/09
''            strExc(0) = "SELECT nvl(YF06,0) FROM PATENTYEARFEE WHERE YF01=" & CNULL(pa(9)) & " AND YF02=" & CNULL(pa(8)) & " AND " & _
''               "YF03='Y99999000' AND YF04='99601' AND YF05=1"
'            strExc(0) = "SELECT nvl(YF06,0) FROM PATENTYEARFEE WHERE YF01=" & CNULL(pa(9)) & " AND YF02=" & CNULL(pa(8)) & " AND " & _
'               "YF03='Y00000000' AND YF04='99601' AND YF05=1"
'            intI = 1
'            Set rsTemp = clslawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'            A1K(9) = 0
'            If intI = 1 Then A1K(9) = rsTemp.Fields(0)
'            A1K(9) = "3600"
            A1K(9) = Val(PUB_GetYF07(pa(9), pa(8), "Y00000000", "601", 1, 1)) + Val(PUB_GetYF07(pa(9), pa(8), "Y00000000", "605", 1, 1))
            'Modify By Cheng 2003/01/30
'            strExc(0) = "SELECT USXR02 FROM USXRATE WHERE RSXR01<=" & strSrvDate(2) & " ORDER BY USXR01 DESC"
            strExc(0) = "SELECT USXR02 FROM USXRATE WHERE USXR01<=" & strSrvDate(2) & " ORDER BY USXR01 DESC"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               A1K(10) = RsTemp.Fields(0)
            End If
            Dim strDisc As String '折扣
            strDisc = 1 - (PUB_GetA1L07Disc(pa(1), pa(2), pa(3), pa(4), "601", A1K(2)) / 100)
            'Modify By Cheng 2003/01/30
            '取得領證及繳年費費用
'            A1K(11) = "7900"
            'Modify By Cheng 2004/01/07
            'A1K11要先扣除折扣後才存檔
'            A1K(11) = Val(PUB_GetYF0607(pa(9), pa(8), "Y00000000", "601", 1, 1)) + Val(PUB_GetYF0607(pa(9), pa(8), "Y00000000", "605", 1, 1))
            A1K(11) = Val(PUB_GetYF0607(pa(9), pa(8), "Y00000000", "601", 1, 1)) + Val(PUB_GetYF0607(pa(9), pa(8), "Y00000000", "605", 1, 1)) - Val(A1K(9) * Val(strDisc))
            'End
            If A1K(10) <> "" Then
                'Modify By Cheng 2004/04/26
                '美金取至整數位
'               A1K(8) = Format(Val(A1K(11)) / Val(A1K(10)), "0.00")
               A1K(8) = Fix(Val(A1K(11)) / Val(A1K(10)))
                'End
            Else
                'Modify By Cheng 2004/04/26
                '美金取至整數位(無條件捨去)
'               A1K(8) = Format(Val(A1K(11)), "0.00")
               A1K(8) = Fix(Val(A1K(11)))
                'End
            End If
            A1K(13) = pa(1)
            A1K(14) = pa(2)
            A1K(15) = pa(3)
            A1K(16) = pa(4)
            A1K(18) = "USD"
            A1K(19) = strSrvDate(2)
            A1K(20) = Format(time, "HHMMSS")
            A1K(21) = strUserNum
            'Add By Cheng 2003/01/30
            '列印對象及請款對象皆為代理人編號
'            A1K(27) = ChangeCustomerL(A1K(3))
'            A1K(28) = ChangeCustomerL(A1K(3))
            A1K(27) = PUB_GetA1K27(pa(1), pa(2), pa(3), pa(4), "601")
            If A1K(27) = "" Then A1K(27) = A1K(3)
            A1K(28) = PUB_GetA1K28(pa(1), pa(2), pa(3), pa(4), "601")
            If A1K(28) = "" Then A1K(28) = A1K(3)
            
            'Add By Cheng 2003/02/24
            '是否列印申請人
            'Modify by Morgan 2004/12/16 改規則
            'A1K(4) = PUB_GetA1K04(pa(1), pa(2), pa(3), pa(4))
            A1K(4) = PUB_GetA1K04(pa(1), pa(2), pa(3), pa(4), A1K(28), "601")
            '2004/12/16 end
            
            '已收金額預存零
            A1K(30) = 0
            'Modify By Cheng 2003/01/30
            '若存檔成功
'            If Not SaveNew1K0(A1K) Then
            If SaveNew1K0(A1K) = True Then
'                Dim strDisc As String '折扣
'                strDisc = 1 - (PUB_GetA1L07Disc(pa(1), pa(2), pa(3), pa(4), "601", A1K(2)) / 100)
               strExc(1) = "INSERT INTO  ACC1L0  (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L08,A1L09,A1L10) VALUES " & _
                  "('" & A1K(1) & "','001','" & pa(1) & "','601'," & A1K(9) & ",NULL," & A1K(9) * Val(strDisc) & "," & A1K(19) & "," & A1K(20) & ",'" & A1K(21) & "')"
               strExc(2) = "INSERT INTO ACC1L0  (A1L01,A1L02,A1L03,A1L04,A1L05,A1L06,A1L07,A1L08,A1L09,A1L10) VALUES " & _
                  "('" & A1K(1) & "','002','" & pa(1) & "','60199'," & Val(PUB_GetYF06(pa(9), pa(8), "Y00000000", "601", 1, 1)) + Val(PUB_GetYF06(pa(9), pa(8), "Y00000000", "605", 1, 1)) & ",NULL,0 ," & A1K(19) & "," & A1K(20) & ",'" & A1K(21) & "')"
               'edit by nickc 2007/02/05 不用 dll 了
               'If objLawDll.ExecSQL(2, strExc) Then
               If ClsLawExecSQL(2, strExc) Then
                  
                  PUB_UpdateA1k08 A1K(1) 'Added by Morgan 2012/11/2 更新請款單外幣金額
               
                    'Add By Cheng 2003/02/25
                    '更新進度檔的請款編號欄(CP60)
                    StrSQLa = "Select * From CaseProgress Where CP01='" & pa(1) & "' And CP02='" & pa(2) & "' And CP03='" & pa(3) & "' And CP04='" & pa(4) & "' And CP10='" & 領證及繳年費 & "' Order By CP05 Desc, CP09 Desc "
                    rsA.CursorLocation = adUseClient
                    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                    If rsA.RecordCount > 0 Then
                        StrSQLa = "Update CaseProgress Set CP60='" & A1K(1) & "' Where CP09='" & rsA("CP09").Value & "' "
                        cnnConnection.Execute StrSQLa
                    End If
                    
                    PUB_PointAutoassign A1K(1), True 'Add by Morgan 2010/4/21 自動分配點數
                    
                    If rsA.State <> adStateClosed Then rsA.Close
                    Set rsA = Nothing
                    '新增請款單列表資料
                    pub_AddressListSN = pub_AddressListSN + 1
                    PUB_AddNewDebitNoteList strUserNum, A1K(1), "" & pub_AddressListSN
                    
                    'Added by Lydia 2016/11/21 整批列印:以請款對象檢查是否存在於國外固定寄催款單代理人檔(ACC225)且下次寄發日期＞系統日，若存在則新增列印清單
                    If PUB_ChkAcc225MsgList(A1K(1), A1K(28), pa(1), pa(2), pa(3), pa(4), IIf(Option1(0).Value = True, Me.Caption, "")) Then
                    End If
                    'end 2016/11/21
                    
A0:               '列印 P/I
                  
               End If
            End If
         End If
      End If
   End If

End Sub

Private Sub cmdok_Click(Index As Integer)
Dim strTmp As String, rsTemp1 As New ADODB.Recordset, rsTemp2 As New ADODB.Recordset
Dim i As Integer
'Add By Cheng 2002/12/24
Dim strSitu As String '定稿的處理狀況

   Select Case Index
      Case 0 '確定
        'Add By Cheng 2003/01/30
        '檢查本所期限
        If Me.Text1(5).Text = "" Then
            MsgBox "請輸入本所期限!!!", vbExclamation + vbOKOnly
            Me.Text1(5).SetFocus
            Text1_GotFocus 5
            Exit Sub
        End If
        If CheckIsTaiwanDate(Me.Text1(5).Text) = False Then
            Me.Text1(5).SetFocus
            Text1_GotFocus 5
            Exit Sub
        End If
        
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 清除查詢印表記錄檔欄位
         '公告日
         If Option1(0).Value = True Then
            If Text1(0).Text <> "" Then
               If Not ChkDate(Text1(0).Text) Then
                  Text1(0).SetFocus
                  TextInverse Text1(0)
                  Exit Sub
               End If
            Else
               MsgBox "公告日不得空白，請重新輸入 !", vbCritical
               Text1(0).SetFocus
               Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            'Modify By Cheng 2002/12/27
            '印公告日當天的資料
'            strExc(0) = "SELECT PA01||PA02||PA03||PA04,TPB08,PA16," & ChgPatent("", 1) & ",NVL(PA06,NVL(PA07,PA05)),PA11 FROM PATENT,TPBULLETIN WHERE PA01='FCP' AND PA09='" & 台灣國家代號 & _
'               "' AND PA14<=" & TransDate(Text1(0).Text, 2) & " AND (PA57<>'Y' OR PA57 IS NULL) AND PA21 IS NULL AND PA11=TPB01(+)"
            'Modify By Cheng 2003/05/21
            '抓卷宗性質為申請(1)的資料
'            strExc(0) = "SELECT PA01||PA02||PA03||PA04,TPB08,PA16," & ChgPatent("", 1) & ",NVL(PA06,NVL(PA07,PA05)),PA11,PA01,PA02,PA03,PA04 FROM PATENT,TPBULLETIN WHERE PA01='FCP' AND PA09='" & 台灣國家代號 & _
'               "' AND PA14=" & TransDate(Text1(0).Text, 2) & " AND (PA57<>'Y' OR PA57 IS NULL) AND PA21 IS NULL AND PA11=TPB01(+)"
            
            pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & Text1(0) 'Add By Sindy 2010/12/7
            
            strExc(0) = "SELECT PA01||PA02||PA03||PA04,TPB08,PA16," & ChgPatent("", 1) & ",NVL(PA06,NVL(PA07,PA05)),PA11,PA01,PA02,PA03,PA04 FROM PATENT,TPBULLETIN WHERE PA01='FCP' AND PA09='" & 台灣國家代號 & _
               "' AND PA14=" & TransDate(Text1(0).Text, 2) & " AND PA20 IS NOT NULL AND (PA24 IS NULL AND PA25 IS NULL) AND (PA57<>'Y' OR PA57 IS NULL) AND PA21 IS NULL AND PA11=TPB01(+) And PA23 = 1 "
            intI = 1
            Set rsTemp2 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               With rsTemp2
                  InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/7
                  Lprint = 0
                  Do While Not .EOF
                     intI = 1
                     strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE " & ChgCaseprogress(.Fields(0)) & " AND CP10='" & 被異議理由 & "'"
                     Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                     If rsTemp1.Fields(0) = 0 Then
                        strReceiveNo = .Fields(0)
                        'Add By Cheng 2002/12/24
                        strSitu = GetSitu(strReceiveNo)
'                        StartLetter ET01, "00"
'                        NowPrint .Fields(0) & "&000", ET01, "00", False, strUserNum, 0
                        StartLetter ET01, strSitu
                        NowPrint .Fields(0) & "&000", ET01, strSitu, False, strUserNum, 0
                        'Add By Cheng 2003/01/29
                        '新增地址條列表資料
                        pub_AddressListSN = pub_AddressListSN + 1
                        'Modify By Cheng 2003/02/07
                        '加傳入綠皮貼紙的份數
'                        PUB_AddNewAddressList strUserNum, "" & .Fields("PA01").Value, "" & .Fields("PA02").Value, "" & .Fields("PA03").Value, "" & .Fields("PA04").Value, "" & pub_AddressListSN
                        PUB_AddNewAddressList strUserNum, "" & .Fields("PA01").Value, "" & .Fields("PA02").Value, "" & .Fields("PA03").Value, "" & .Fields("PA04").Value, "" & pub_AddressListSN, "0"
                        'Add By Cheng 2003/09/10
                        '新增整批定稿列印清單資料
                        PUB_AddNewLetterList "公告期滿通知函", Me.Text1(0).Text, "" & .Fields("PA01").Value, "" & .Fields("PA02").Value, "" & .Fields("PA03").Value, "" & .Fields("PA04").Value
                     Else
                        intI = 1
                        strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE " & ChgCaseprogress(.Fields(0)) & " AND CP10='" & 異議答辯 & "'"
                        Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                        If rsTemp1.Fields(0) = 0 Then
                           '不印通知函
                        Else
                           intI = 1
                           strExc(0) = "SELECT DECODE(SUM(DECODE(CP24,'',0,1)),COUNT(*),0,1) FROM CASEPROGRESS WHERE " & ChgCaseprogress(.Fields(0)) & " AND CP10='" & 異議答辯 & "'"
                           Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                           If rsTemp1.Fields(0) = 1 Then
                              '不印通知函
                           Else
                              If .Fields(2) = "1" Then
                                 ReDim Preserve sField(Lprint) As strPrint  'Added by Lydia 2016/11/21
                                 If Not IsNull(.Fields(0)) Then sField(Lprint).No1 = .Fields(0)
                                 If Not IsNull(.Fields(3)) Then sField(Lprint).No = .Fields(3)
                                 If Not IsNull(.Fields(4)) Then sField(Lprint).Name = .Fields(4)
                                 If Not IsNull(.Fields(5)) Then sField(Lprint).Case = .Fields(5)
                                 Lprint = Lprint + 1
                              End If
                           End If
                        End If
                     End If
                     .MoveNext
                  Loop
               End With
               If Lprint > 0 Then
                  PrintCase
                  For i = 0 To Lprint - 1
                     PrintPI sField(i).No1
                  Next
               End If
               MsgBox "列印結束 !", vbInformation
            Else
               InsertQueryLog (0) 'Add By Sindy 2010/12/7
               MsgBox "無符合條件之資料可列印 !", vbInformation
            End If
            Screen.MousePointer = vbDefault
         '本所案號
         Else
            strTmp = Text1(1) & Text1(2)
            pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & Text1(1) & "-" & Text1(2) 'Add By Sindy 2010/12/7
            If Text1(3).Text = "" Then
               strTmp = strTmp & "0"
            Else
               strTmp = strTmp & Text1(3).Text
               pub_QL05 = pub_QL05 & "-" & Text1(3) 'Add By Sindy 2010/12/7
            End If
            If Text1(4).Text = "" Then
               strTmp = strTmp & "00"
            Else
               strTmp = strTmp & Text1(4).Text
               pub_QL05 = pub_QL05 & "-" & Text1(4) 'Add By Sindy 2010/12/7
            End If
            Screen.MousePointer = vbHourglass
            strReceiveNo = strTmp
            'Modify By Cheng 2003/05/21
            '抓卷宗性質為申請(1)的資料
'            strExc(0) = "SELECT PA09,PA01,PA02,PA03,PA04 FROM PATENT WHERE " & ChgPatent(strReceiveNo) & " AND PA09='" & 台灣國家代號 & _
'               "' AND (PA57<>'Y' OR PA57 IS NULL) AND PA21 IS NULL"
            strExc(0) = "SELECT PA09,PA01,PA02,PA03,PA04 FROM PATENT WHERE " & ChgPatent(strReceiveNo) & " AND PA09='" & 台灣國家代號 & _
               "' AND PA20 IS NOT NULL AND (PA24 IS NULL AND PA25 IS NULL) AND (PA57<>'Y' OR PA57 IS NULL) AND PA21 IS NULL And PA23=1 "
            intI = 1
            Set rsTemp2 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                With rsTemp2
                    InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/7
                    While Not rsTemp2.EOF
                        'Add By Cheng 2002/12/24
                        strSitu = GetSitu(strReceiveNo)
        '               StartLetter ET01, "00"
        '               NowPrint strReceiveNo & "&000", ET01, "00", False, strUserNum, 0
                       StartLetter ET01, strSitu
                       NowPrint strReceiveNo & "&000", ET01, strSitu, False, strUserNum, 0
                        'Add By Cheng 2003/01/29
                        '新增地址條列表資料
                        pub_AddressListSN = pub_AddressListSN + 1
                        'Modify By Cheng 2003/02/07
                        '加傳入綠皮貼紙的份數
'                        PUB_AddNewAddressList strUserNum, "" & .Fields("PA01").Value, "" & .Fields("PA02").Value, "" & .Fields("PA03").Value, "" & .Fields("PA04").Value, "" & pub_AddressListSN
                        PUB_AddNewAddressList strUserNum, "" & .Fields("PA01").Value, "" & .Fields("PA02").Value, "" & .Fields("PA03").Value, "" & .Fields("PA04").Value, "" & pub_AddressListSN, "0"
                       PrintPI strReceiveNo
                        .MoveNext
                    Wend
                End With
               MsgBox "列印結束 !", vbInformation
               
            Else
               InsertQueryLog (0) 'Add By Sindy 2010/12/7
               MsgBox "無符合條件之資料可列印 !", vbInformation
            End If
            Screen.MousePointer = vbDefault
         End If
      Case 1 '結束
            Me.Enabled = False
            Unload Me
   End Select
End Sub

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
Dim strTxt(1 To 20) As String, i As Integer, j As Integer, strTmp As String
'Add By Cheng 2003/01/19
Dim ii As Integer
'Add By Cheng 2003/02/12
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

    ii = 0
    EndLetter ET01, strReceiveNo & "&000", ET03, strUserNum
    'Modify By Cheng 2003/01/24
    '不論是否續辦與否
'    strExc(0) = "SELECT NVL(np08,''),NVL(np09,'') FROM NEXTPROGRESS WHERE " & ChgNextProgress(strReceiveNo) & " AND NP07=" & 領證及繳年費 & " AND NP06 IS NULL ORDER BY NP08"
    strExc(0) = "SELECT NVL(np08,''),NVL(np09,'') FROM NEXTPROGRESS WHERE " & ChgNextProgress(strReceiveNo) & " AND NP07=" & 領證及繳年費 & " ORDER BY NP08 Desc "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        'Modify By Cheng 2003/01/30
'        ii = ii + 1
'       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'          "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','本所期限'," & CNULL(rsTemp.Fields(0)) & ")"
'        ii = ii + 1
'       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'          "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','法定期限'," & CNULL(rsTemp.Fields(1)) & ")"
        'Modify By Cheng 2003/01/19
'       If Not objLawDll.ExecSQL(2, strTxt) Then
'          MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
'       End If
    End If
    'Add By Cheng 2003/01/24
    '抓母案期限資料
    If intI <> 1 Then
        strExc(0) = "SELECT NVL(np08,''),NVL(np09,'') FROM NEXTPROGRESS WHERE " & ChgNextProgress(Left(strReceiveNo, Len(strReceiveNo) - 3) & "000") & " AND NP07=" & 領證及繳年費 & " ORDER BY NP08 Desc "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
        If intI = 1 Then
'            ii = ii + 1
'           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'              "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','本所期限'," & CNULL(rsTemp.Fields(0)) & ")"
'            ii = ii + 1
'           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'              "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','法定期限'," & CNULL(rsTemp.Fields(1)) & ")"
        End If
    End If
    'Add By Cheng 2003/01/30
    '例外欄位--本所期限
     ii = ii + 1
    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
       "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','本所期限'," & CNULL(DBDATE(Me.Text1(5).Text)) & ")"
    'Add By Cheng 2003/01/24
    '中文定稿例外欄位
     ii = ii + 1
     '規費
    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
       "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','規費','" & Val(PUB_GetYF07(台灣國家代號, "1", "Y00000000", "601", 1, 1)) + Val(PUB_GetYF07(台灣國家代號, "1", "Y00000000", "605", 1, 1)) & "')"
     ii = ii + 1
     '服務費
    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
       "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','服務費','" & Val(PUB_GetYF06(台灣國家代號, "1", "Y00000000", "601", 1, 1)) + Val(PUB_GetYF06(台灣國家代號, "1", "Y00000000", "605", 1, 1)) & "')"
    'Add By Cheng 2003/01/19
     ii = ii + 1
     '領證費
    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
       "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','領證費','" & Val(PUB_GetYF0607(台灣國家代號, "1", "Y00000000", "601", 1, 1)) + Val(PUB_GetYF0607(台灣國家代號, "1", "Y00000000", "605", 1, 1)) & "')"
     ii = ii + 1
     '費用
    'Modify By Cheng 2004/04/27
    '美金取至整數位(無條件捨去)
'    strTxt(iI) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'       "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','費用','" & Format((Val(PUB_GetYF0607(台灣國家代號, "1", "Y00000000", "601", 1, 1)) + Val(PUB_GetYF0607(台灣國家代號, "1", "Y00000000", "605", 1, 1))) / PUB_GetUSXRate, "0.00") & "')"
    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
       "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','費用','" & Format(Fix((Val(PUB_GetYF0607(台灣國家代號, "1", "Y00000000", "601", 1, 1)) + Val(PUB_GetYF0607(台灣國家代號, "1", "Y00000000", "605", 1, 1))) / PUB_GetUSXRate), "0.00") & "')"
    'End
     ii = ii + 1
     '領證及繳年費費用
    'Modify By Cheng 2004/04/27
    '美金取至整數位(無條件捨去)
'    strTxt(iI) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'       "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','領證及繳年費費用','" & Format(5000 / PUB_GetUSXRate, "0.00") & "')"
    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
       "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','領證及繳年費費用','" & Format(Fix(5000 / PUB_GetUSXRate), "0.00") & "')"
    'End
     ii = ii + 1
     '年費費用1
    'Modify By Cheng 2004/04/27
    '美金取整數位(無條件捨去)
'    strTxt(iI) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'       "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','年費費用1','" & Format(4800 / PUB_GetUSXRate, "0.00") & "')"
    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
       "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','年費費用1','" & Format(Fix(4800 / PUB_GetUSXRate), "0.00") & "')"
    'End
     ii = ii + 1
     '年費費用2
    'Modify By Cheng 2004/04/27
    '美金取整數位(無條件捨去)
'    strTxt(iI) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'       "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','年費費用2','" & Format(7500 / PUB_GetUSXRate, "0.00") & "')"
    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
       "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','年費費用2','" & Format(Fix(7500 / PUB_GetUSXRate), "0.00") & "')"
    'End
     ii = ii + 1
     '年費費用3
    'Modify By Cheng 2004/04/27
    '美金取整數位(無條件捨去)
'    strTxt(iI) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'       "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','年費費用3','" & Format(12800 / PUB_GetUSXRate, "0.00") & "')"
    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
       "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','年費費用3','" & Format(Fix(12800 / PUB_GetUSXRate), "0.00") & "')"
    'End
     ii = ii + 1
     '年費費用4
    'Modify By Cheng 2004/04/27
    '美金取整數位(無條件捨去)
'    strTxt(iI) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'       "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','年費費用4','" & Format(23200 / PUB_GetUSXRate, "0.00") & "')"
    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
       "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','年費費用4','" & Format(Fix(23200 / PUB_GetUSXRate), "0.00") & "')"
    'End
    'Add By Cheng 2003/02/12
    '判斷是否不續辦但准通知
    StrSQLa = "Select PA89 From Patent Where " & ChgPatent(strReceiveNo)
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        If "" & rsA.Fields(0).Value = "Y" Then
             ii = ii + 1
             '附註
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','附註','P.S. : This case has been allowed. If your client(s) want(s) to maintain this case, please notify us immediately.')"
        End If
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    'edit by nickc 2007/02/05 不用 dll 了
    'If Not objLawDll.ExecSQL(ii, strTxt) Then
    If Not ClsLawExecSQL(ii, strTxt) Then
       MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
    End If
End Sub

Private Sub Form_Load()

MoveFormToCenter Me
intWhere = 國外_FC
Option1_Click 0

'設定印表機
SeekPrintL = Printer.Orientation
PUB_SetPrinter Me.Name, Combo1, m_OriPrinterName, False, SeekPrint       'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除
PUB_SetPrinter Me.Name, Combo2

End Sub

Private Sub Form_Unload(Cancel As Integer)

   'Copy from cmdok_Click by Morgan 2004/10/26
   '列印定稿整批列印清單
   'Modified by Lydia 2020/09/24 +程式名稱
   'PUB_PrintLetterList strUserNum
   PUB_PrintLetterList strUserNum, , , , " and LL02='公告期滿通知函' "
   '刪除定稿整批列印資料
   'Modified by Lydia +傳入刪除條件
   'PUB_DeleteLetterList strUserNum
   PUB_DeleteLetterList strUserNum, " and LL02='公告期滿通知函' "
   
   '列印請款單
   PUB_PrintDebitNote strUserNum, Me.Combo2.Text
   '刪除請款單列表資料
   PUB_DeleteDebitNoteList strUserNum
   
   'Added by Lydia 2016/11/21
   '列印:國外固定寄催款單清單
   PUB_PrintAcc225List strUserNum, Me.Combo2.Text
   '刪除:國外固定寄催款單清單
   PUB_DeleteAcc225List strUserNum
   'end 2016/11/21
   
   '列印地址條
   PUB_PrintAddressList strUserNum, Me.Combo1.Text
   '刪除地址條列表資料
   PUB_DeleteAddressList strUserNum
   
   '初始化序號
   pub_AddressListSN = 0
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '若請款單印表機變動, 則更新列印設定
   If Me.Combo2.Text <> Me.Combo2.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
   End If
   '2004/10/26 end

'Add By Cheng 2003/01/28
'還原預設印表機
Set Printer = Printers(SeekPrint)
Printer.Orientation = SeekPrintL
Set frm060303 = Nothing
End Sub

Private Sub PrintCase()
 Dim i As Integer, Page As Integer, iPrint As Integer
On Error GoTo ErrHnd
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
      If i < Lprint Then
         If (i Mod 37 = 0 And i <> 0) Then
            Printer.NewPage
            Page = Page + 1
            CaseTitle Page
            iPrint = 2700
         End If
         iPrint = iPrint + 300
      End If
   Next
   Printer.EndDoc
   Exit Sub
ErrHnd:
   MsgBox Err.Description
End Sub

Private Sub GetPrintLeft()
   PLeft(0) = 500:     PLeft(1) = 2000
   PLeft(2) = 3500
End Sub

Private Sub CaseTitle(ByVal Page As String)
 Dim i As Integer
   i = 500
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4000:         Printer.CurrentY = i
   Printer.Print "被異議不成立清單"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 4500:         Printer.CurrentY = i + 500
   Printer.Print "公告日 : " & Text1(0).Text
   Printer.Font.Bold = False
   Printer.CurrentX = 500:              Printer.CurrentY = i + 800
   Printer.Print "列印人 : " & strUserName
   Printer.CurrentX = 9000:            Printer.CurrentY = i + 800
   Printer.Print "列印日期 : " & Format(GetTaiwanTodayDate, "##/##/##")
   Printer.CurrentX = 9000:            Printer.CurrentY = i + 1100
   Printer.Print "頁次 : " & Page
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

Private Sub Option1_Click(Index As Integer)
 Dim txt As TextBox, i As Integer
On Error Resume Next
   For Each txt In Text1
      txt.Enabled = False
   Next
   Select Case Index
      Case 0
         Text1(0).Enabled = True
         Text1(0).SetFocus
      Case 1
         For i = 2 To 4
            Text1(i).Enabled = True
         Next
         Text1(2).SetFocus
   End Select
   'Add By Cheng 2003/01/30
    '設定本所期限欄作用中
   Me.Text1(5).Enabled = True
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   If Text1(Index) = "" Then Exit Sub
   If Option1(0).Value = True Then
      If Index = 0 Then
         If Text1(Index).Text <> "" Then
            If Not ChkDate(Text1(Index).Text) Then
               Text1(Index).SetFocus
               TextInverse Text1(Index)
            End If
         Else
            MsgBox "公告日不得空白，請重新輸入 !", vbCritical
            Text1(Index).SetFocus
         End If
      End If
   Else
      If Index = 1 Then
         If Text1(Index).Text = "" Then
            MsgBox "本所案號不得空白，請重新輸入 !", vbCritical
            Text1(Index).SetFocus
         End If
      End If
   End If
End Sub

'Add By Cheng 2002/12/24
'取得定稿處理方式
Private Function GetSitu(strPA0104 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset

GetSitu = "00"
StrSQLa = "Select * From PATENT WHERE " & ChgPatent(strPA0104)
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    '若基本檔有設定定稿語文
    If "" & rsA("PA85").Value <> "" Then
        Select Case "" & rsA("PA85").Value
        Case "1" '中文
            GetSitu = "01"
        Case "2" '英文
'            'FCP領證自動代繳欄
'            If "" & rsA("PA71").Value = "Y" Then
'                GetSitu = "03"
'            Else
                GetSitu = "02"
'            End If
        Case "3" '日文
            GetSitu = "06"
        End Select
    '若基本檔未設定定稿語文
    Else
        '若基本檔有代理人
        If "" & rsA("PA75").Value <> "" Then
            StrSqlB = "Select * From FAGENT WHERE FA01='" & Mid(rsA("PA75").Value, 1, 8) & "' AND FA02='" & Mid(rsA("PA75").Value, 9, 1) & "'"
            rsB.CursorLocation = adUseClient
            rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
            If rsB.RecordCount > 0 Then
                Select Case "" & rsB("FA31").Value
                Case "1" '中文
                    GetSitu = "01"
                Case "2" '英文
'                    '申請案號的第九碼非NULL
'                    If "" & Mid("" & rsA("PA11").Value, 9, 1) <> "" Then
'                        GetSitu = "05"
'                    'FCP領證自動代繳欄
'                    ElseIf "" & rsB("FA42").Value = "Y" Then
'                        GetSitu = "03"
'                    '收款後辦案有值
'                    ElseIf "" & rsB("FA39").Value <> "" Then
'                        GetSitu = "04"
'                    Else
                        GetSitu = "02"
'                    End If
                Case "3" '日文
                    GetSitu = "06"
                End Select
            End If
        '若基本檔無代理人
        Else
            StrSqlB = "Select * From CUSTOMER WHERE CU01='" & Mid(rsA("PA26").Value, 1, 8) & "' AND CU02='" & Mid(rsA("PA26").Value, 9, 1) & "'"
            rsB.CursorLocation = adUseClient
            rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
            If rsB.RecordCount > 0 Then
                Select Case "" & rsB("CU64").Value
                Case "1" '中文
                    GetSitu = "01"
                Case "2" '英文
'                    '申請案號的第九碼非NULL
'                    If "" & Mid("" & rsA("PA11").Value, 9, 1) <> "" Then
'                        GetSitu = "05"
'                    'FCP領證自動代繳欄
'                    ElseIf "" & rsB("CU75").Value = "Y" Then
'                        GetSitu = "03"
'                    '收款後辦案有值
'                    ElseIf "" & rsB("CU72").Value <> "" Then
'                        GetSitu = "04"
'                    Else
                        GetSitu = "02"
'                    End If
                Case "3" '日文
                    GetSitu = "06"
                End Select
            End If
        End If
    End If
End If
If rsB.State <> adStateClosed Then rsB.Close
Set rsB = Nothing
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
'Add By Cheng 2003/05/21
'若為英文定稿, 再判斷領證是否自動代繳
If GetSitu = "02" Then
    StrSQLa = "Select * From PATENT WHERE " & ChgPatent(strPA0104)
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        '若基本檔有設定FCP領證自動代繳欄
        If "" & rsA("PA71").Value = "Y" Then
            GetSitu = "03"
        '若基本檔無設定FCP領證自動代繳欄
        Else
            '若基本檔有代理人
            If "" & rsA("PA75").Value <> "" Then
                StrSqlB = "Select * From FAGENT WHERE FA01='" & Mid(rsA("PA75").Value, 1, 8) & "' AND FA02='" & Mid(rsA("PA75").Value, 9, 1) & "'"
                rsB.CursorLocation = adUseClient
                rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
                If rsB.RecordCount > 0 Then
                    '申請案號的第九碼非NULL
                    'Modify by Morgan 2010/12/27 申請案號改碼數
                    'If "" & Mid("" & rsA("PA11").Value, 9, 1) <> "" Then
                    If "" & Mid("" & rsA("PA11").Value, 10, 1) <> "" Then
                        GetSitu = "05"
                    'FCP領證自動代繳欄
                    ElseIf "" & rsB("FA42").Value = "Y" Then
                        GetSitu = "03"
                    '收款後辦案有值
                    ElseIf "" & rsB("FA39").Value <> "" Then
                        GetSitu = "04"
                    End If
                End If
            '若基本檔有申請人
            ElseIf "" & rsA("PA26").Value <> "" Then
                StrSqlB = "Select * From CUSTOMER WHERE CU01='" & Mid(rsA("PA26").Value, 1, 8) & "' AND CU02='" & Mid(rsA("PA26").Value, 9, 1) & "'"
                rsB.CursorLocation = adUseClient
                rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
                If rsB.RecordCount > 0 Then
                    '申請案號的第九碼非NULL
                    'Memo by Morgan2010/12/27 申請案號欄已修改
                    'If "" & Mid("" & rsA("PA11").Value, 9, 1) <> "" Then
                    If "" & Mid("" & rsA("PA11").Value, 10, 1) <> "" Then
                        GetSitu = "05"
                    'FCP領證自動代繳欄
                    ElseIf "" & rsB("CU75").Value = "Y" Then
                        GetSitu = "03"
                    '收款後辦案有值
                    ElseIf "" & rsB("CU72").Value <> "" Then
                        GetSitu = "04"
                    End If
                End If
            End If
            If rsB.State <> adStateClosed Then rsB.Close
            Set rsB = Nothing
        End If
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End If
'若無任何設定, 預設為英文一般定稿
If GetSitu = "00" Then GetSitu = "02"
End Function

''Add By Cheng 2003/02/24
'Private Function GetPrintCust(strTmp As String) As String
'Dim rsA As New ADODB.Recordset
'Dim strSQLA As String
'
'GetPrintCust = ""
''取得專利基本檔的"D/N是否列印申請人"
'strSQLA = "SELECT PA78 FROM PATENT WHERE " & ChgPatent(strTmp) & " AND PA78 IS NOT NULL "
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'rsA.CursorLocation = adUseClient
'rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'   GetPrintCust = rsA.Fields(0).Value
'Else
'   '取得國外代理人檔的"D/N是否列印申請人"
'   strSQLA = "SELECT FA44 FROM PATENT, FAGENT WHERE " & ChgPatent(strTmp) & _
'            " AND SUBSTR(PA75,1,8)=FA01 AND SUBSTR(PA75,9,1)=FA02 AND FA44 IS NOT NULL "
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
'   rsA.CursorLocation = adUseClient
'   rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsA.RecordCount > 0 Then
'      GetPrintCust = rsA.Fields(0).Value
'   Else
'      '取得客戶基本檔的"D/N是否列印申請人"
'      strSQLA = "SELECT CU77 FROM PATENT, CUSTOMER WHERE " & ChgPatent(strTmp) & _
'               " AND SUBSTR(PA26,1,8)=CU01 AND SUBSTR(PA26,9,1)=CU02 AND CU77 IS NOT NULL "
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'      rsA.CursorLocation = adUseClient
'      rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'         GetPrintCust = rsA.Fields(0).Value
'      End If
'   End If
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'End Function



