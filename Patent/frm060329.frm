VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm060329 
   BorderStyle     =   1  '單線固定
   Caption         =   "年費逾期補繳通知函(整批)"
   ClientHeight    =   2550
   ClientLeft      =   120
   ClientTop       =   945
   ClientWidth     =   4530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4530
   Begin VB.TextBox textPA01 
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "FCP"
      Top             =   855
      Width           =   495
   End
   Begin VB.TextBox textPA02 
      Height          =   315
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   4
      Top             =   855
      Width           =   855
   End
   Begin VB.TextBox textPA03 
      Height          =   315
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   5
      Top             =   855
      Width           =   255
   End
   Begin VB.TextBox textPA04 
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   6
      Top             =   855
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號"
      Height          =   225
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   900
      Width           =   1125
   End
   Begin VB.OptionButton Option1 
      Caption         =   "來函收文日"
      Height          =   225
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   465
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   1425
      TabIndex        =   7
      Top             =   1230
      Width           =   2985
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1425
      Style           =   2  '單純下拉式
      TabIndex        =   8
      Top             =   1560
      Width           =   2985
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   11
      Top             =   540
      Visible         =   0   'False
      Width           =   330
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtDate 
      Height          =   315
      Left            =   1440
      MaxLength       =   9
      TabIndex        =   1
      Top             =   420
      Width           =   870
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   2745
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   3570
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "清單及定稿："
      Height          =   180
      Index           =   11
      Left            =   120
      TabIndex        =   16
      Top             =   1290
      Width           =   1080
   End
   Begin VB.Label Label2 
      Caption         =   "地址條印表機："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   1590
      Width           =   1260
   End
   Begin VB.Label lblCount 
      Alignment       =   2  '置中對齊
      Caption         =   "( 0 / 0 )"
      Height          =   165
      Left            =   120
      TabIndex        =   14
      Top             =   2340
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "是否含已閉卷案件：         ( Y:是 )"
      Height          =   180
      Left            =   2820
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   2580
   End
End
Attribute VB_Name = "frm060329"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
'Create by Morgan 2010/3/8
Option Explicit

Dim strPrinter2 As String 'Add By Sindy 2015/9/25
Dim m_PrintRpt1 As Boolean, ff1 As Integer, m_strFileName1 As String 'Add By Sindy 2016/1/27


Private Sub cmdOK_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 0
         cmdOK(Index).Enabled = False
         If TxtValidate Then
            ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/8 清除查詢印表記錄檔欄位
            Process
         End If
         cmdOK(Index).Enabled = True
      Case 2
         Unload Me
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtDate = strSrvDate(2)
   
   PUB_SetPrinter Me.Name, Combo1
   'Add By Sindy 2015/9/25
   PUB_SetPrinter Me.Name, Combo2, strPrinter2
   '2015/9/25 END
      
   'Added by Lydia 2019/06/17 預設Option顯示
    EnableTextBox txtDate, True
    EnableTextBox textPA01, False
    EnableTextBox textPA02, False
    EnableTextBox textPA03, False
    EnableTextBox textPA04, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2015/9/25 +清單
   '列印定稿整批列印清單
   PUB_PrintLetterList strUserNum, "8", Me.Combo2.Text, strPrinter2
   '刪除定稿整批列印資料
   'Modified by Lydia +傳入刪除條件
   'PUB_DeleteLetterList strUserNum
   PUB_DeleteLetterList strUserNum, " and LL02='年費逾期補繳通知函' "
   '2015/9/25 END
   
   '列印地址條
   PUB_PrintAddressList strUserNum, Me.Combo1.Text
   '刪除地址條列表資料
   PUB_DeleteAddressList strUserNum
   '初始化序號
   pub_AddressListSN = 0
   
   '若地址條印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   
   'Add By Sindy 2015/9/25
   '若請款單印表機變動, 則更新列印設定
   If Me.Combo2.Text <> Me.Combo2.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
   End If
   '2015/9/25 END
   
   Set frm060329 = Nothing
End Sub

Private Sub Text1_GotFocus()
   CloseIme
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And Chr(KeyAscii) <> "Y" Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtDate_GotFocus()
   CloseIme
   TextInverse txtDate
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
   If txtDate = "" Then
      MsgBox "收文日期不可空白！"
      Cancel = True
   ElseIf ChkDate(txtDate) = False Then
      Cancel = True
   End If
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   Dim strTmp As String 'Added by Lydia 2019/06/17
   
   If Option1(0).Value = True Then 'Added by Lydia 2019/06/17 整批檢查
        txtDate_Validate Cancel
        If Cancel = True Then
           txtDate.SetFocus
           Exit Function
        End If
   'Added by Lydia 2019/06/17 個案檢查
   Else
      ' 本所案號
      If IsEmptyText(textPA01) = True Then
         MsgBox "本所案號系統類別不可空白", vbCritical, "檢核資料"
         textPA01.SetFocus
         Exit Function
      End If
      ' 本所案號
      If IsEmptyText(textPA02) = True Then
         MsgBox "本所案號流水號不可空白", vbCritical, "檢核資料"
         textPA02.SetFocus
         Exit Function
      End If
      If IsEmptyText(textPA03) = True Then textPA03 = "0"
      If IsEmptyText(textPA04) = True Then textPA04 = "00"

      If IsExistRecord(strTmp) = False Then
         MsgBox "本所案號不存在", vbOKOnly + vbCritical, "查詢資料"
         textPA02.SetFocus
         Exit Function
      ElseIf strTmp <> "" Then
           If MsgBox("本案已閉卷/銷卷，是否繼續列印定稿？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Function
           End If
      End If
   End If
   'end 2019/06/17
   
   TxtValidate = True
End Function

Private Sub Process()
   Dim adoRst As ADODB.Recordset
   Dim stCP01 As String, stCP02 As String, stCP03 As String, stCP04 As String
   Dim stCP09 As String, stCP06 As String, stCP07 As String, stCP64 As String, stCP27 As String
   Dim ET01 As String, ET02 As String, ET03 As String
   Dim iPos As Integer, stNextYear As String, stOldDate As String
   Dim iLang As Integer, iCopy As Integer
   Dim bolEmail As Boolean, bolPlusPaper As Boolean
   Dim iLetter As Integer
   Dim stPA08 As String 'Added by Morgan 2015/5/21
   'Add By Sindy 2016/1/26
   Dim strFileName As String, strFullFileName As String
   Dim oFileSys As New FileSystemObject
   Dim oFile As File
   Dim strMsg As String
   Dim strNewCP09 As String
   '2016/1/26 END
   Dim strExpired As String 'Added by Morgan 2019/1/17
   
   ET01 = "13"
   
   If Option1(0).Value = True Then 'Added by Lydia 2019/06/17 整批
        'Modified by Lydia 2019/06/17  Label1=>Option1(0).Caption & ":"
        pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & ":" & txtDate  'Add By Sindy 2010/12/8
        If Text1.Text = "Y" Then
           pub_QL05 = pub_QL05 & ";" & Left(Label3, 9) & Text1 'Add By Sindy 2010/12/8
        End If
        
        '2011/11/28 MODIFY BY SONIA Y30150也不寄發,來函的cp27也上1922/11/11故加pa75
        'Modified by Morgan 2015/5/27 +pa167
        '1605.通知年費逾期
        'Modify By Sindy 2015/9/25 +,GetEmailFlag(CP09) eMail
        'Modified by Morgan 2019/1/17 +PA25
        'Modified by Lydia 2019/06/17 +CP27,CP05,PA108
        'Modified by Morgan 2019/6/28 +pa26
        'Modified by Lydia 2019/07/03 已閉卷的案件不用印定稿,但是要印在清單與內專提供的清單比對by江如玉
        'strExc(0) = "select cp01,cp02,cp03,cp04,cp09,cp06,cp07,cp64,pa57,pa75,pa08,pa167,GetEmailFlag(CP09) eMail,pa25,CP27,CP05,PA57,PA108,PA26 from caseprogress,patent" & _
           " where cp01='FCP' and cp10='1605' and cp27 is null and cp05=" & DBDATE(txtDate) & _
           " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
           " order by eMail,CP01,CP02,CP03,CP04"
        strExc(0) = "select cp01,cp02,cp03,cp04,cp09,cp06,cp07,cp64,pa57,pa75,pa08,pa167,GetEmailFlag(CP09) eMail,pa25,CP27,CP05,PA108,PA26 from caseprogress,patent" & _
           " where cp01='FCP' and cp10='1605' and cp159=0 and cp05=" & DBDATE(txtDate) & _
           " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
           " order by eMail,CP01,CP02,CP03,CP04"
   'Added by Lydai 2019/06/17
   Else '個案
        pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & ":" & textPA01 & "-" & textPA02 & IIf(textPA03 & textPA04 <> "000", "-" & textPA03 & "-" & textPA04, "")
      
        'Modified by Morgan 2019/6/28 +pa26
        strExc(0) = "select cp01,cp02,cp03,cp04,cp09,cp06,cp07,cp64,pa57,pa75,pa08,pa167,GetEmailFlag(CP09) eMail,pa25,CP27,CP05,PA108,PA26 from caseprogress,patent" & _
           " where cp01='FCP' and cp10='1605' and cp01 = '" & textPA01 & "' and cp02 = '" & textPA02 & "' and cp03 = '" & textPA03 & "' and cp04 = '" & textPA04 & "' " & _
           " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
           " and cp09 in (select max(cp09) mno from caseprogress where CP01='" & textPA01 & "' And CP02='" & textPA02 & "' And CP03='" & textPA03 & "' And CP04='" & textPA04 & "' and cp10='1605' and cp159=0) " & _
           " order by eMail,CP01,CP02,CP03,CP04"
   'end 2019/06/17
   End If
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoRst
      InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/8
      ProgressBar1.Min = 0
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      ProgressBar1.Visible = True
      lblCount.Visible = True
      iLetter = 0
      DoEvents
      
      'Add by Sindy 2015/9/25
      pub_OsPrinter = PUB_GetOsDefaultPrinter
      PUB_SetOsDefaultPrinter Combo2.Text
      PUB_SetWordActivePrinter
      PUB_RestorePrinter Combo2.Text
      '2015/9/25 END
      
      .MoveFirst
      Do While Not .EOF
         stCP01 = .Fields("cp01")
         stCP02 = .Fields("cp02")
         stCP03 = .Fields("cp03")
         stCP04 = .Fields("cp04")
         stCP09 = .Fields("cp09")
         stCP06 = "" & .Fields("cp06")
         stCP07 = "" & .Fields("cp07")
         stCP64 = "" & .Fields("cp64")
         stPA08 = "" & .Fields("pa08") 'Added by Morgan 2015/5/21
         stCP27 = ""  'Added by Lydia 2019/06/1 7
         
         'Added by Morgan 2019/1/17
         '專利期間屆滿日, 落於原年費期限(含)及六個月補繳期限(含)則年費逾期函不發--David
         strExpired = ""
         If .Fields("CP07") > .Fields("PA25") Then
            strExpired = "Ｘ"
            stCP27 = "19221111"
         'end 2019/1/17

         '2011/11/28 add by sonia Y30150也不寄發
         'Modified by Morgan 2013/5/3 +Y52013
         ElseIf .Fields("pa75") = "Y30150000" Or .Fields("pa75") = "Y52013000" Then 'FC代理人
            stCP27 = "19221111"
         '2011/11/28 end
         
         'Added by Morgan 2019/6/28 --楊映慈
         ElseIf .Fields("pa75") = "Y20656000" And InStr("X70722010,X70286000,X70762010", .Fields("pa26")) > 0 Then
            stCP27 = "19221111"
         'end 2019/6/28
         
         'Added by Morgan 2013/5/15 FCP-23191,FCP-23378,FCP-26506,FCP-30733不要通知--Helen,Susan
         'MODIFY BY SONIA 2014/6/5 FCP-037844也不通知
         'Modified by Morgan 2014/9/22 +FCP-41852
         'Modified by Lydia 2014/10/27 +FCP-45401
         'Modified by Morgan 2015/2/4 +FCP-30914
         'Modified by Morgan 2015/5/25 +FCP-42952 -- 江如玉
         'Modified by Morgan 2015/5/27 改判斷PA167
         'ElseIf .Fields("cp01") = "FCP" And InStr("023191,023378,026506,030733,037844,041852,045401,030914,042952", .Fields("cp02")) > 0 Then
         'Mark by Lydia 2019/06/17 在內專-智慧局年費通知核對清單(frm040331),凡閉卷/銷卷/PA167=N預上假發文
         'ElseIf .Fields("pa167") = "N" Then '是否寄發年費逾期補繳通知單
         '   stCP27 = "19221111"
         'end 2013/5/15
         
         'ElseIf IsNull(.Fields("pa57")) Then '是否閉卷
         '   stCP27 = strSrvDate(1)
            
         'ElseIf Text1.Text = "Y" Then '是否含已閉卷案件
         '   stCP27 = strSrvDate(1)
         
         'Else
         '   stCP27 = "19221111"
         'end 2019/06/17
         End If
         
         If stCP27 <> "" Then 'Added by Lydia 2019/06/17 判斷
            strSql = "update caseprogress set cp14='" & strUserNum & "',cp27=" & stCP27 & " where cp09='" & stCP09 & "'"
            cnnConnection.Execute strSql, intI
         'Added by Lydia 2019/06/17
         Else
            stCP27 = "" & .Fields("cp27")
            '有特殊案件已閉卷尚需報告，則程序會個案下定稿，到時自動將發文日"111111"拿掉，如此案件又可自大批上發文
            If Option1(1).Value = True And "" & .Fields("CP27") = "19221111" And Trim("" & .Fields("PA57") & .Fields("PA108")) <> "" Then
                 '將發文日"111111"拿掉並且上承辦期限
                 strExc(1) = CompDate(2, 10, "" & .Fields("cp05"))
                 strSql = "Update Caseprogress set cp27=null,cp48=" & IIf(Val(strExc(1)) < strSrvDate(1), strSrvDate(1), strExc(1)) & _
                             " where cp09='" & .Fields("cp09") & "' and cp10='1605' "
                 Pub_SeekTbLog strSql
                 cnnConnection.Execute strSql, intI
            End If
         End If
         
         'Modified by Lydia 2019/06/17 整批區分出假發文不用列印
         'If stCP27 <> "19221111" Then
         If (Option1(0).Value = True And stCP27 <> "19221111") Or Option1(1).Value = True Then
'            'Add By Sindy 2015/9/25 列印FCP承辦單
'            Call PUB_PrintFCPEmpBill(stCP01, stCP02, stCP03, stCP04, ET01, stCP09, , , "2")
            
            ET02 = stCP01 & stCP02 & stCP03 & stCP04 & "&1605"
            ET03 = "01"
            iLang = PUB_GetLanguage(stCP01, stCP02, stCP03, stCP04, "1605", "1")
            If iLang = 3 Then ET03 = "03" '日文
            
            iPos = InStr(stCP64, "未繳年度:")
            If iPos > 0 Then
               stNextYear = Val(Mid(stCP64, iPos + 5))
            Else
               stNextYear = ""
            End If
            
            iPos = InStr(stCP64, "原繳費期限:")
            If iPos > 0 Then
               stOldDate = Val(Mid(stCP64, iPos + 6))
            Else
               stOldDate = ""
            End If
            
            EndLetter ET01, ET02, ET03, strUserNum
            
            If stOldDate <> "" Then
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                  "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','年費法定期限','" & DBDATE(stOldDate) & "')"
               cnnConnection.Execute strSql, intI
            End If
            
            If stNextYear <> "" Then
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                  "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','第幾年至幾年費','" & stNextYear & "')"
               cnnConnection.Execute strSql, intI
            End If
   
            If stCP07 <> "" Then
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                  "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','年費延展繳費日','" & DBDATE(stCP07) & "')"
               cnnConnection.Execute strSql, intI
            End If
            
            'Added by Morgan 2015/5/21
            '一案兩請提醒
            If stPA08 = "2" Then
               'Modified by Morgan 2017/1/24 +判斷發明無證書號才帶(+ and pa22 is null) --David
               strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa11,pa77,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CNo" & _
                  " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & stCP01 & "' and cm02='" & stCP02 & "' and cm03='" & stCP03 & "' and cm04='" & stCP04 & "'" & _
                  " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & stCP01 & "' and cm06='" & stCP02 & "' and cm07='" & stCP03 & "' and cm08='" & stCP04 & "') X" & _
                  ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null and pa22 is null"
               intI = 1
               Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                     "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','一案兩請新型案要印','♀')"
                  cnnConnection.Execute strSql, intI
                  
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                     "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','發明案申請號','" & adoRecordset("pa11") & "')"
                  cnnConnection.Execute strSql, intI
                  
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','發明案彼所案號','" & IIf(IsNull(adoRecordset("pa77")), "", "" & adoRecordset("pa77")) & "')"
                  cnnConnection.Execute strSql, intI
                  
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                     "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','發明案本所案號','" & adoRecordset("CNo") & "')"
                  cnnConnection.Execute strSql, intI
               End If
            End If
            'end 2015/5/21
            
            bolEmail = PUB_GetEMailFlag(stCP01 & stCP02 & stCP03 & stCP04, True, , bolPlusPaper)
            '判斷是否EMail同時寄紙本
            If bolPlusPaper Then
               iCopy = 0
            Else
               iCopy = 1
            End If
            
            If bolEmail Then
               NowPrint ET02, ET01, ET03, False, strUserNum, , , , , iCopy, , True, True
               'MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(stCP01) & " ]！"
            Else
               NowPrint ET02, ET01, ET03, False, strUserNum
            End If
            
            'PUB_PrintLetter ET02 '直接印出定稿 Add By Sindy 2015/9/25
            'Modify By Sindy 2016/1/27 定稿轉PDF存卷宗區
            strNewCP09 = stCP09
            strFileName = .Fields("CP01") & .Fields("CP02") & IIf(.Fields("CP04") <> "00", "-" & .Fields("CP03") & "-" & .Fields("CP04"), IIf(.Fields("CP03") <> "0", "-" & .Fields("CP03"), "")) & ".1605.CUS.PDF"
            PUB_DelFtpFile2 strNewCP09, " and cpp02='" & strFileName & "'" '檔案改放 FTP,必須在DB資料刪除前執行
            strSql = "delete from CasePaperPDF where cpp01='" & strNewCP09 & "' and cpp02='" & strFileName & "'"
            cnnConnection.Execute strSql
            If PUB_PrintLetter(ET02, , , True, strFullFileName) = True Then
               Call PUB_ChkFileStatus(strFullFileName, False, strMsg)  'Added by Lydia 2022/10/31 判斷檔案是否存在, 超過時間就繼續; ex. 10.27整批列印,因為沒有檔案才發生上傳錯誤
               Set oFile = oFileSys.GetFile(strFullFileName)
               If SaveAttFile_PDF(strNewCP09, strFullFileName, strFileName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False) = False Then
                  'Modified by Lydia 2022/10/28 +& ";" & strMsg
                  Call ReadTxt1(.Fields("CP01") & "-" & .Fields("CP02") & "-" & .Fields("CP03") & "-" & .Fields("CP04"), strNewCP09, "定稿轉PDF失敗" & ";" & strMsg)
               End If
               Kill strFullFileName
            End If
            '2016/1/27 END
            
            If Not bolEmail Or bolPlusPaper Then
'               'Add By Sindy 2015/9/21 日文定稿才要印地址條
'               If iLang = 3 Or Val(外專開窗信函啟用日) >= Val(strSrvDate(1)) Then
'               '2015/9/21 END
                  '新增地址條列表資料
                  pub_AddressListSN = pub_AddressListSN + 1
                  PUB_AddNewAddressList strUserNum, stCP01, stCP02, stCP03, stCP04, "" & pub_AddressListSN, "0", "605"
'               End If
            End If
            
            iLetter = iLetter + 1
         End If
         
         '新增整批定稿列印清單資料 : N.不寄逾繳函
         PUB_AddNewLetterList "年費逾期補繳通知函", "收文日期:" & ChangeTStringToTDateString(txtDate), stCP01, stCP02, stCP03, stCP04, IIf(bolEmail, IIf(bolPlusPaper, "Ｅ", "ｅ"), "") & IIf(.Fields("pa167") = "N", "Ν", "") & strExpired
         
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblCount.Caption = "( " & ProgressBar1.Value & " / " & ProgressBar1.max & " )"
         
         DoEvents
         .MoveNext
      Loop
      
      'Add by Sindy 2015/9/25
      PUB_SetOsDefaultPrinter pub_OsPrinter
      PUB_RestorePrinter strPrinter2
      '2015/9/25 END
      
      End With
      'MsgBox "定稿已產生，共 " & iLetter & " 筆！"
      'Modify By Sindy 2016/1/27
      If m_PrintRpt1 = True Then
         Close ff1
         strMsg = "請至下列位置列印檢核表：" & PUB_Getdesktop & "\" & m_strFileName1
      End If
      MsgBox "定稿列印完畢，共 " & iLetter & " 筆！" & strMsg, vbInformation
      '2016/1/27 END
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/8
      MsgBox "無未發文之年費逾期補繳通知程序！"
   End If
   ProgressBar1.Visible = False
   lblCount.Visible = False
   Set adoRst = Nothing
End Sub

'Add By Sindy 2016/1/27
'資料檢核表
Private Sub ReadTxt1(strCaseNo As String, strRecvNo As String, strNote As String)
Dim i As Integer
Dim strTemp(1 To 7) As String
   
   If m_PrintRpt1 = False Then
      m_PrintRpt1 = True
      If ff1 > 0 Then Close #ff1
      ff1 = FreeFile
      m_strFileName1 = Me.Caption & txtDate & "資料檢核表.txt"
      Open PUB_Getdesktop & "\" & m_strFileName1 For Output As ff1
      'Print #ff1, "備註：改字型Fixedsys標準11號字以橫式上下左右各10MM列印"
      Print #ff1, "本所案號        總收文號   原因"
      Print #ff1, "=============== ========== ============================================="
   End If
   For i = 1 To 3
      strTemp(i) = ""
   Next i
   strTemp(1) = convForm(CheckStr(Trim(strCaseNo)), 15)
   strTemp(2) = convForm(CheckStr(Trim(strRecvNo)), 10)
   strTemp(3) = Trim(strNote)
   Print #ff1, strTemp(1) & " " & strTemp(2) & " " & strTemp(3)
End Sub

'Added by Lydia 2019/06/17
Private Sub Option1_Click(Index As Integer)
   If Option1(0).Value = True Then
      txtDate.SetFocus
      EnableTextBox txtDate, True
      EnableTextBox textPA01, False
      EnableTextBox textPA02, False
      EnableTextBox textPA03, False
      EnableTextBox textPA04, False
   Else
      textPA02.SetFocus
      EnableTextBox txtDate, False
      EnableTextBox textPA01, True
      EnableTextBox textPA02, True
      EnableTextBox textPA03, True
      EnableTextBox textPA04, True
   End If
End Sub
Private Sub textPA01_GotFocus()
   TextInverse textPA01
End Sub

Private Sub textPA02_GotFocus()
   TextInverse textPA02
End Sub

Private Sub textPA03_GotFocus()
   TextInverse textPA03
End Sub

Private Sub textPA04_GotFocus()
   TextInverse textPA04
End Sub

Private Function IsExistRecord(ByRef CaseType As String) As Boolean
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   
   CaseType = ""
   
   IsExistRecord = False
   strSql = "SELECT PA01,PA02,PA03,PA04, PA57||PA108||PA167 AS CTYPE FROM PATENT " & _
            "WHERE PA01 = '" & textPA01 & "' AND " & _
                  "PA02 = '" & textPA02 & "' AND " & _
                  "PA03 = '" & textPA03 & "' AND " & _
                  "PA04 = '" & textPA04 & "' "
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsExistRecord = True
      CaseType = "" & rsTmp.Fields("CTYPE")
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function



