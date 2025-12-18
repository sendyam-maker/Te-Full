VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040325 
   BorderStyle     =   1  '單線固定
   Caption         =   "公開通知函"
   ClientHeight    =   5136
   ClientLeft      =   2796
   ClientTop       =   3948
   ClientWidth     =   6372
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5136
   ScaleWidth      =   6372
   Begin VB.TextBox txtByte 
      Height          =   315
      Left            =   1620
      TabIndex        =   18
      Text            =   "30000"
      Top             =   4350
      Width           =   975
   End
   Begin VB.TextBox txtMinSec 
      Height          =   315
      Left            =   1620
      TabIndex        =   17
      Text            =   "5"
      Top             =   4710
      Width           =   555
   End
   Begin VB.TextBox txtMaxSec 
      Height          =   315
      Left            =   4200
      TabIndex        =   16
      Text            =   "45"
      Top             =   4710
      Width           =   555
   End
   Begin VB.TextBox txtFirstAdd 
      Height          =   315
      Left            =   4200
      TabIndex        =   15
      Text            =   "3"
      Top             =   4350
      Width           =   555
   End
   Begin VB.TextBox txtPDFPath 
      Height          =   315
      Left            =   1860
      TabIndex        =   8
      Text            =   "C:\Program Files\Adobe\Reader 8.0\Reader\AcroRd32.exe"
      Top             =   1272
      Width           =   4395
   End
   Begin VB.ComboBox cmbPrinter2 
      Height          =   276
      Left            =   1860
      TabIndex        =   9
      Top             =   1668
      Width           =   4395
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   4
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   7
      Top             =   900
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   3
      Left            =   2628
      MaxLength       =   1
      TabIndex        =   6
      Top             =   900
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   2
      Left            =   1788
      MaxLength       =   6
      TabIndex        =   5
      Top             =   900
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   1
      Left            =   1308
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "P"
      Top             =   900
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   0
      Left            =   1308
      MaxLength       =   7
      TabIndex        =   2
      Top             =   564
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   936
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "　公開日："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   612
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5052
      TabIndex        =   11
      Top             =   180
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4164
      TabIndex        =   10
      Top             =   180
      Width           =   756
   End
   Begin VB.ListBox List1 
      Height          =   1668
      ItemData        =   "frm040325.frx":0000
      Left            =   60
      List            =   "frm040325.frx":0007
      TabIndex        =   14
      Top             =   2400
      Width           =   6225
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1296
      TabIndex        =   0
      Top             =   192
      Visible         =   0   'False
      Width           =   2088
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3678;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      Caption         =   "程序人員："
      Height          =   240
      Left            =   216
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "幾Byte算1秒："
      Height          =   180
      Index           =   5
      Left            =   450
      TabIndex        =   22
      Top             =   4350
      Width           =   1140
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "1個檔至少幾秒："
      Height          =   180
      Left            =   240
      TabIndex        =   21
      Top             =   4710
      Width           =   1350
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "1個檔最多幾秒："
      Height          =   180
      Left            =   2820
      TabIndex        =   20
      Top             =   4710
      Width           =   1350
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "第1個檔多加幾秒："
      Height          =   180
      Left            =   2640
      TabIndex        =   19
      Top             =   4350
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PDF執行檔路徑："
      Height          =   180
      Left            =   96
      TabIndex        =   13
      Top             =   1332
      Width           =   1560
   End
   Begin VB.Label Label6 
      Caption         =   "列印公報PDF印表機："
      Height          =   180
      Left            =   96
      TabIndex        =   12
      Top             =   1728
      Width           =   1752
   End
End
Attribute VB_Name = "frm040325"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modified by Morgan 2025/1/15 增加程序人員選單並刪除不再使用的物件及部分舊程式碼
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim intWhere As Integer
Dim strReceiveNo As String '本所案號
'Add By Sindy 2011/12/22
Dim strTPG04 As String, strTPG05 As String
Dim i As Integer, j As Integer
'Modify By Sindy 2014/9/3
Dim m_DefaultPrinter As String
'Dim m_DefaultPrinter2 As String
Dim strPrinter As String
'2014/9/3 END
'Dim SeekPrint As Integer
Dim strTime As String
'2011/12/22 End
Dim m_bolELetter As Boolean 'Added by Morgan 2014/6/19 是否有存電子信函
Dim m_AttachPath As String 'Added by Morgan 2022/7/12 公報PDF暫存路徑

Private Sub cmdok_Click(Index As Integer)
   'edit by nickc 2007/02/06 不用 dll 了
   'Dim objPrintDllPublic As New clsPrintPublic
   Dim strTmp As String, rsTemp1 As New ADODB.Recordset, rsTemp2 As New ADODB.Recordset
   Dim stET03 As String 'Add by Morgan 2004/11/12 處理狀況
   'Add By Sindy 2011/12/21
   Dim int_Copys As Integer
   '2011/12/21 End
   Dim strCP09 As String 'Add By Sindy 2014/6/18
   Dim strLP26 As String 'Added by Morgan 2016/1/11
   'Added by Morgan 2025/1/15
   Dim rsQuery As ADODB.Recordset
   Dim mSeqNo As String, stVTB0 As String
   'end 2025/1/15
   
    Select Case Index
    Case 0 '確定
         'Add By Sindy 2011/12/27
         List1.Clear
'         If cmbPrinter2.ListIndex >= 0 Then
'             Set Printer = Printers(cmbPrinter2.ListIndex)
''             Printer.EndDoc
'         End If
         '2011/12/27 End
         'Modify By Sindy 2014/9/3
         '系統印表機
         PUB_RestorePrinter cmbPrinter2
         '設定控制台預設印表機
         PUB_SetOsDefaultPrinter cmbPrinter2
         '2014/9/3 END
         
        ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/18 清除查詢印表記錄檔欄位
        '公開日
        If Option1(0).Value = True Then
            If Text1(0).Text <> "" Then
                If Not ChkDate(Text1(0).Text) Then
                    Text1(0).SetFocus
                    TextInverse Text1(0)
                    Exit Sub
                End If
            Else
                MsgBox "公開日不得空白，請重新輸入 !", vbCritical
                Text1(0).SetFocus
                Exit Sub
            End If
            
            'Add By Sindy 2011/12/27
            'Removed by Morgan 2022/7/12 公報改抓卷宗區
            'If GetFilePath(DBDATE(Text1(0))) = False Then
            '   Me.txtPath2.SetFocus
            '   Exit Sub
            'End If
            'end 2022/7/12
            '2011/12/27 End
            
            Screen.MousePointer = vbHourglass
            pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & Text1(0) 'Add By Sindy 2010/11/18
            'Modify By Sindy 2011/12/27 +PA11
            'Modified by Morgan 2025/1/15 +PID
            'Modified by Morgan 2025/8/6 排除公開公報已有信函進度的案件以避免重複執行
            strExc(0) = "SELECT PA01||PA02||PA03||PA04 C01,DECODE(TPG08,'台一國際',1,0) C02,PA01,PA02,PA03,PA04,CU12,PA26,PA13,PA14,PA11,pa75,'' PID" & _
               " FROM PATENT,TPGAZETTE, Customer WHERE PA01='P' And PA09='000' AND PA12=" & TransDate(Text1(0).Text, 2) & _
               " AND (PA57<>'Y' OR PA57 IS NULL) AND PA11=TPG01 And Substr(PA26,1,8)=CU01(+) And Substr(PA26,9,1)=CU02(+)" & _
               " and not exists(select * from caseprogress,letterprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04" & _
               " and cp10='1229' and lp01(+)=cp09 and lp01 is not null)" & _
               " order by PA13"
            intI = 1
            Set rsTemp2 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            
            'Added by Morgan 2025/1/15
            If intI = 1 And strSrvDate(1) >= P業務區劃分啟用日 And Combo1 <> "" Then
               Combo1.Tag = ""
               Set rsQuery = PUB_CreateRecordset(rsTemp2, , , 300, Me.Name, mSeqNo)
               With rsQuery
                  .MoveFirst
                  Do While Not .EOF
                     .Fields("PID") = PUB_GetPHandler(.Fields("PA01") & "-" & .Fields("PA02") & "-" & .Fields("PA03") & "-" & .Fields("PA04"))
                     .MoveNext
                  Loop
                  .UpdateBatch
                  
                  stVTB0 = "select R001 as " & .Fields(0).Name
                  For intI = 2 To .Fields.Count
                     stVTB0 = stVTB0 & ", R" & Format(intI, "000") & " as " & .Fields(intI - 1).Name
                  Next
                  stVTB0 = stVTB0 & " from Rdatafactory Where Id='" & strUserNum & "' And Formname='" & Me.Name & "'  And Seqno='" & mSeqNo & "'"
               End With
               strSql = "Select X.* From (" & stVTB0 & ") X where PID='" & Left(Combo1, 5) & "' order by PA13"
               intI = 1
               Set rsTemp2 = ClsLawReadRstMsg(intI, strSql)
               Combo1.Tag = Combo1
            End If
            'end 2025/1/15
               
            If intI = 1 Then
                With rsTemp2
                    InsertQueryLog (.RecordCount) 'Add By Sindy 2010/11/18
                    Do While Not .EOF
                        '處理定稿例外欄位
                        strReceiveNo = "" & .Fields(0).Value
                        '大-->台 定稿 2008/09/15 ADD BY TONI
                        If PUB_CheckCuNation(.Fields("PA26"), .Fields("PA01"), .Fields("PA02"), .Fields("PA03"), .Fields("PA04")) = "1" Then
                           stET03 = "02"
                        Else
                           'Add by Morgan 2004/11/12
                           stET03 = "00"
                           
                           'Added by Morgan 2022/7/12 寶齡富錦 Y55435 案件
                           If "" & .Fields("PA75") = "Y55435000" Then
                              stET03 = "99"
                           End If
                           'end 2022/7/12
                            
                           If Val("" & .Fields("PA14")) > 0 Then stET03 = "01"
                           '2004/11/12 end
                        End If
                        
                        m_bolELetter = False 'Added by Morgan 2014/6/19
                        
                        'Add By Sindy 2014/6/18
                        strCP09 = ""
                        strSql = "SELECT cp09 FROM caseprogress " & _
                                 "WHERE CP01='" & .Fields("PA01") & "' AND CP02='" & .Fields("PA02") & "' AND CP03='" & .Fields("PA03") & "' AND CP04='" & .Fields("PA04") & "' " & _
                                  " AND CP10 = '1229'"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           strCP09 = RsTemp.Fields("cp09")
                           
                           'Added by Morgan 2025/1/20
                           strSql = "update caseprogress set cp27=" & strSrvDate(1) & ",cp82=to_char(sysdate,'hh24miss'),cp83='" & strUserNum & "' where cp09='" & strCP09 & "' and cp127 is null"
                           cnnConnection.Execute strSql, intI
                           'end 2025/1/20
                           
                           'Modified by Morgan 2014/7/22 +傳FC代理人(pa75)
                           'Modified by Morgan 2015/12/2 要傳非掛號
                           'Modified by Morgan 2016/1/11 +strLP26
                           If intI = 1 Then 'Added by Morgan 2025/8/6 發文室未發文才可新增信函進度
                              Call PUB_AddLetterProgress(RsTemp.Fields("cp09"), 1, True, "", False, .Fields("PA26"), "1229", "" & .Fields("pa75"), , strLP26)
                           End If
                           
                           m_bolELetter = True 'Added by Morgan 2014/6/19
                        End If
                        '2014/6/18 END
                        
                        StartLetter "20", stET03
                        'Modify By Sindy 2014/6/18
                        'NowPrint "" & .Fields(0) & "&000", "20", stET03, False, strUserNum, 0
                        NowPrint "" & .Fields(0) & "&000", "20", stET03, False, strUserNum, 0, , , , , , , , , , , , strCP09
                        '2014/6/18 END
                        'Add By Sindy 2011/12/21
                        'Modified by Morgan 2016/1/11 e化不印公報
                        'Modified by Morgan 2022/2/14 全E化也不要印
                        'If Not (m_bolELetter And strLP26 = "Y") Then
                        If Not (m_bolELetter And strLP26 <> "") Then
                        'end 2022/2/14
                           Call GetPDFCopys(.Fields("PA01"), .Fields("PA02"), .Fields("PA03"), .Fields("PA04"), "" & .Fields("PA11"), int_Copys)
                        End If
                        '2011/12/21 End
                        
                        .MoveNext
                    Loop
                End With
                
               'Add By Sindy 2011/12/21 列印PDF
               If List1.ListCount > 0 Then
                  Call PrintPDF
                  MsgBox "定稿產生完成 ! (列印PDF花費時間：" & strTime & "  " & time() & ")", vbInformation
               Else
               '2011/12/21 End
                  MsgBox "定稿產生完成 !", vbInformation
               End If
            Else
                InsertQueryLog (0) 'Add By Sindy 2010/11/18
                MsgBox "無符合條件之資料 !", vbInformation
            End If
            Screen.MousePointer = vbDefault
        '本所案號
        Else
            If Text1(2) = "" Then
                MsgBox "本所案號不得空白，請重新輸入 !", vbCritical
                Text1(2).SetFocus
                Exit Sub
            End If
            strTmp = Text1(1) & Text1(2)
            If Text1(3).Text = "" Then
                strTmp = strTmp & "0"
            Else
                strTmp = strTmp & Text1(3).Text
            End If
            If Text1(4).Text = "" Then
                strTmp = strTmp & "00"
            Else
                strTmp = strTmp & Text1(4).Text
            End If
            Screen.MousePointer = vbHourglass
            pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & Text1(1) & "-" & Text1(2) & "-" & IIf(Text1(3) = "", "0", Text1(3)) & "-" & IIf(Text1(4) = "", "00", Text1(4)) 'Add By Sindy 2010/11/18
            'Modify By Cheng 2003/07/04
'            strExc(0) = "SELECT PA01||PA02||PA03||PA04,DECODE(TPB08,'台一國際',1,0),PA01,PA02,PA03,PA04 FROM PATENT,TPBULLETIN WHERE " & ChgPatent(strTmp) & _
'                            " AND (PA57<>'Y' OR PA57 IS NULL) AND PA11=TPB01(+) And PA01='P' And PA09='000' "
            'Modify By Cheng 2003/08/05
'            strExc(0) = "SELECT PA01||PA02||PA03||PA04,DECODE(TPB08,'台一國際',1,0),PA01,PA02,PA03,PA04, CU12, PA26 FROM PATENT,TPBULLETIN, Customer WHERE " & ChgPatent(strTmp) & _
'                            " AND (PA57<>'Y' OR PA57 IS NULL) AND PA11=TPB01(+) And PA01='P' And PA09='000' And Substr(PA26,1,8)=CU01(+) And Substr(PA26,9,1)=CU02(+) "
            'Modify by Morgan 2004/11/12 加公告日 PA14
            'Modify By Sindy 2011/12/27 +PA11
            strExc(0) = "SELECT PA01||PA02||PA03||PA04,DECODE(TPG08,'台一國際',1,0),PA01,PA02,PA03,PA04,CU12,PA26,PA14,PA11,pa12,pa75 FROM PATENT,TPGAZETTE, Customer WHERE " & ChgPatent(strTmp) & _
                            " AND (PA57<>'Y' OR PA57 IS NULL) AND PA11=TPG01 And PA01='P' And PA09='000' And Substr(PA26,1,8)=CU01(+) And Substr(PA26,9,1)=CU02(+) "
            intI = 1
            Set rsTemp2 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                With rsTemp2
                    InsertQueryLog (.RecordCount) 'Add By Sindy 2010/11/18
                    Do While Not .EOF
                        'Add By Sindy 2011/12/27
                        If "" & .Fields("PA12") > "" Then
                           'Modified by Morgan 2022/7/12 公報改抓卷宗區
                           'If GetFilePath("" & .Fields("PA12")) = False Then
                           '   Me.txtPath2.SetFocus
                           If PUB_GetGazettePDF(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"), , , , True) = False Then
                              MsgBox .Fields(0) & "案公告公報卷宗區PDF檔讀取失敗！", vbExclamation
                           'end 2022/7/12
                              Screen.MousePointer = vbDefault
                              Exit Sub
                           End If
                        End If
                        '2011/12/27 End
                        
                        '處理定稿例外欄位
                        strReceiveNo = "" & .Fields(0).Value
                         '大-->台 定稿 2008/09/15 ADD BY TONI
                        If PUB_CheckCuNation(.Fields("PA26"), .Fields("PA01"), .Fields("PA02"), .Fields("PA03"), .Fields("PA04")) = "1" Then
                           stET03 = "02"
                        Else
                           'Add by Morgan 2004/11/12
                           stET03 = "00"
                           
                           'Added by Morgan 2022/7/12 寶齡富錦 Y55435 案件
                           If "" & .Fields("PA75") = "Y55435000" Then
                              stET03 = "99"
                           End If
                           'end 2022/7/12
                           
                           If Val("" & .Fields("PA14")) > 0 Then stET03 = "01"
                           '2004/11/12 end
                        End If
                        
                        m_bolELetter = False 'Added by Morgan 2014/6/19
                        
                        'Add By Sindy 2014/6/18
                        strCP09 = ""
                        strSql = "SELECT cp09 FROM caseprogress " & _
                                 "WHERE CP01='" & .Fields("PA01") & "' AND CP02='" & .Fields("PA02") & "' AND CP03='" & .Fields("PA03") & "' AND CP04='" & .Fields("PA04") & "' " & _
                                  " AND CP10 = '1229'"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           strCP09 = RsTemp.Fields("cp09")
                           
                           'Added by Morgan 2025/1/20
                           If strSrvDate(1) >= P業務區劃分啟用日 Then
                              strSql = "update caseprogress set cp27=" & strSrvDate(1) & ",cp82=to_char(sysdate,'hh24miss'),cp83='" & strUserNum & "' where cp09='" & strCP09 & "' and cp127 is null"
                              cnnConnection.Execute strSql, intI
                           End If
                           'end 2025/1/20
                           
                           'Modified by Morgan 2014/7/22 +傳FC代理人(pa75)
                           'Modified by Morgan 2015/12/2 要傳非掛號
                           'Modified by Morgan 2016/1/11 +strLP26
                           Call PUB_AddLetterProgress(RsTemp.Fields("cp09"), 1, True, "", False, .Fields("PA26"), "1229", "" & .Fields("pa75"), , strLP26)
                           m_bolELetter = True 'Added by Morgan 2014/6/19
                        End If
                        '2014/6/18 END
                        
                        StartLetter "20", stET03
                        'Modify By Sindy 2014/6/18
                        'NowPrint "" & .Fields(0) & "&000", "20", stET03, False, strUserNum, 0
                        NowPrint "" & .Fields(0) & "&000", "20", stET03, False, strUserNum, 0, , , , , , , , , , , , strCP09
                        '2014/6/18 END
                        'Add By Sindy 2011/12/21
                        'Modified by Morgan 2016/1/11 e化不印公報
                        'Modified by Morgan 2022/2/14 全E化也不要印
                        'If "" & .Fields("PA12") > "" And Not (m_bolELetter And strLP26 = "Y") Then
                        If "" & .Fields("PA12") > "" And Not (m_bolELetter And strLP26 <> "") Then
                        'end 2022/2/14
                           Call GetPDFCopys(.Fields("PA01"), .Fields("PA02"), .Fields("PA03"), .Fields("PA04"), "" & .Fields("PA11"), int_Copys)
                        End If
                        '2011/12/21 End
                        
                        'Remove by Morgan 2008/8/13 改開窗定稿
                        '北所的客戶才要印地址條
                        'If "" & .Fields("CU12").Value <> "" And ("" & .Fields("CU12").Value < "S20" Or "" & .Fields("CU12").Value > "S49") And "" & .Fields("PA26").Value <> "" Then
                        '    '新增地址條列表資料
                        '    pub_AddressListSN = pub_AddressListSN + 1
                        '    PUB_AddNewAddressList strUserNum, "" & .Fields("PA01").Value, "" & .Fields("PA02").Value, "" & .Fields("PA03").Value, "" & .Fields("PA04").Value, "" & pub_AddressListSN, "0"
                        'End If
                        
                        .MoveNext
                    Loop
                End With
                
               'Add By Sindy 2011/12/21 列印PDF
               If List1.ListCount > 0 Then
                  Call PrintPDF
                  MsgBox "定稿產生完成 ! (列印PDF花費時間：" & strTime & "  " & time() & ")", vbInformation
               Else
               '2011/12/21 End
                  MsgBox "定稿產生完成 !", vbInformation
               End If
            Else
               InsertQueryLog (0) 'Add By Sindy 2010/11/18
               MsgBox "無符合條件之資料 !", vbInformation
            End If
            Screen.MousePointer = vbDefault
        End If
        
         'Modify By Sindy 2014/9/3
         '還原系統中預設印表機
         PUB_RestorePrinter m_DefaultPrinter
         '還原控制台預設印表機
         PUB_SetOsDefaultPrinter strPrinter
         '2014/9/3 END
         
    Case 1 '結束
        Unload Me
    End Select
End Sub

'Add By Sindy 2011/12/22
Private Sub GetPDFCopys(strPA01 As String, strPA02 As String, strPA03 As String, strPA04 As String, StrPA11 As String, ByRef int_Copys As Integer)
Dim strFileName As String
   
   int_Copys = 0
   
   'Added by Morgan 2014/6/19
   '103/7/1 起有存電子信函的只要印 1 份
   If Val(strSrvDate(1)) >= 20140701 And m_bolELetter = True Then
      int_Copys = 1
   Else
   'end 2014/6/19
   
      '由員工檔取得列印份數 (北部的員工印2份, 其它地區的員工印3份)
      'Modified by Morgan 2014/6/4
      'Modified by Morgan 2014/5/27 +特殊設定A7所有編號視為北所人員
      'strExc(0) = "SELECT ST06 FROM STAFF WHERE ST01='" & PUB_GetAKindSalesNo(strPA01, strPA02, strPA03, strPA04) & "' "
      strExc(0) = "SELECT DECODE(instr(';'||replace(oMan,',',';')||';',';'||ST01||';'),0,ST06,'1') FROM STAFF, SetSpecMan WHERE ST01='" & PUB_GetAKindSalesNo(strPA01, strPA02, strPA03, strPA04) & "' and ocode(+)='A7'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      int_Copys = 3
      If intI = 1 Then
         If RsTemp.Fields(0).Value = "1" Then
            int_Copys = 2
         Else
            int_Copys = 3
         End If
      End If
      
   End If 'Added by Morgan 2014/6/19
   
   'Modify By Sindy 2013/1/4
   'strFileName = txtPath2 & "\img_1\pub0" & strTPG04 & "0" & strTPG05 & "\" & StrPA11 & "-P01.pdf"
   'Modified by Morgan 2022/7/12 公報改抓卷宗區，不再往pat3讀取
   'strFileName = txtPath2 & "\img_1\pub0" & strTPG04 & "0" & strTPG05 & "\" & StrPA11 & ".pdf"
   ''2013/1/4 End
   'List1.AddItem strFileName & " " & int_Copys
   If PUB_GetGazettePDF(strPA01, strPA02, strPA03, strPA04, True, m_AttachPath, strFileName, True) Then
      List1.AddItem strFileName & " " & int_Copys
   End If
   'end 2022/7/12
End Sub

'Add By Sindy 2011/12/22
Private Sub PrintPDF()
Dim i As Integer, k As Integer
'Modified by Morgan 2022/7/18
'Dim strTemp As Variant
Dim strTemp(1) As String
'end 2022/7/18
Dim RetVal, intFileCnt As Integer
Dim ff1 As Integer
Dim MySize, dblSec As Double, dblCntSec As Double
'Add By Sindy 2014/9/3
Dim process_id As Long
Dim process_handle_PDF As Long
'2014/9/3 END
   
   strTime = time()
   intFileCnt = 0

   'Add By Sindy 2014/9/3
   '因為第 2 個以後開啟的 Reader 才會印完後自動關閉,所以固定先開一個空的程式,全部印完後再關閉
   process_id = Shell(txtPDFPath, vbHide)
   process_handle_PDF = OpenProcess(PROCESS_TERMINATE, 0, process_id)
   '2014/9/3 END
   
   If ff1 > 0 Then Close #ff1
   ff1 = FreeFile
   'Modified by Morgan 2022/7/12
   'Open txtPath2 & "\專利公開通知函" & strTPG04 & "卷" & strTPG05 & "期" & "列印PDF時間資訊.txt" For Output As ff1
   Open m_AttachPath & "\專利公開通知函" & strTPG04 & "卷" & strTPG05 & "期" & "列印PDF時間資訊.txt" For Output As ff1
   'end 2022/7/12
   
   For i = 0 To List1.ListCount - 1
      'Modified by Morgan 2022/7/18
      'strTemp = Split(List1.List(i), " ")
      intI = InStrRev(List1.List(i), " ")
      strTemp(0) = Left(List1.List(i), intI - 1)
      strTemp(1) = Mid(List1.List(i), intI + 1)
      'end 2022/7/18
      For k = 0 To Val(strTemp(1)) - 1 '列印份數
         intFileCnt = intFileCnt + 1
         
'Modified by Morgan 2022/7/18
         PUB_PrintOnePdf txtPDFPath, " /n /t """ & strTemp(0) & """ """ & cmbPrinter2 & """"
         If k = 0 Then
            MySize = FileLen(strTemp(0))   '傳回檔案長度 (以 Byte 為單位)
            Print #ff1, Left(i + 1 & "     ", 5) & List1.List(i) & " " & MySize
         End If
'end 2022/7/18
         
      Next k
   Next i
   
'   '還原控制台預設印表機
'   If cmbPrinter2.ListIndex >= 0 Then
'      PUB_SetOsDefaultPrinter m_DefaultPrinter
'   End If
   
   Print #ff1, "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   Print #ff1, "列印時間：" & strTime & "  " & time()
   Print #ff1, "檔案數量：" & intFileCnt
   Close ff1
   
   'Add By Sindy 2014/9/3
   TerminateProcess process_handle_PDF, 0&
   CloseHandle process_handle_PDF
   DoEvents
   '2014/9/3 END
End Sub

Private Sub Form_Activate()
   Option1_Click 0 'Modified by Morgan 2025/1/15 預設在公開日欄位(從form_load移來)
End Sub

Private Sub Form_Load()
'Dim SeekPrintL As Integer
'Dim i As Integer, j As Integer
   
   MoveFormToCenter Me
   intWhere = 國外_FC
   
   'Add By Sindy 2011/12/27
   'Modify By Sindy 2014/9/3
   PUB_SetPrinter Me.Name, cmbPrinter2, m_DefaultPrinter
   strPrinter = PUB_GetOsDefaultPrinter '抓控制台目前預設的印表機
   '2014/9/3 END
   
   List1.Clear
   If Pub_StrUserSt03 = "M51" Then
      'txtPDFPath = "C:\Program Files\Adobe\Reader 8.0\Reader\AcroRd32.exe"
      Me.Height = 6120
   Else
      'C:\Program Files\Adobe\Acrobat 8.0\Acrobat\Acrobat.exe
      'C:\Program Files\Adobe\Acrobat 7.0\Reader\AcroRd32.exe
      'txtPDFPath = "C:\Program Files\Adobe\Acrobat 8.0\Acrobat\Acrobat.exe"
      Me.Height = 2700
   End If
   '2011/12/27 End
   txtPDFPath = PUB_SetFileAssociation 'Add By Sindy 2014/9/3
   
   'Added by Morgan 2025/1/15
   If strSrvDate(1) >= P業務區劃分啟用日 Then
      Combo1.Visible = True
      Label4.Visible = True
      Call SetPatentP12Combo(Combo1, "P", Label4)
   End If
   'end 2025/1/15

   'Added by Morgan 2022/7/12 公報PDF暫存路徑
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   'end 2021/6/24
End Sub

Private Sub Form_Unload(Cancel As Integer)

   'Modify By Sindy 2014/9/3
   '若印表機變動, 則更新列印設定
   If Me.cmbPrinter2.Text <> Me.cmbPrinter2.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter2.Name, "0", "0", Me.cmbPrinter2.Text
   End If
   '2014/9/3 END
   
   Set frm040325 = Nothing
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
            MsgBox "公開日不得空白，請重新輸入 !", vbCritical
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

'定稿例外欄位
Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)

    EndLetter ET01, strReceiveNo & "&000", ET03, strUserNum

End Sub
