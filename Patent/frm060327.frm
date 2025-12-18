VERSION 5.00
Begin VB.Form frm060327 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利權消滅函"
   ClientHeight    =   2160
   ClientLeft      =   3090
   ClientTop       =   3255
   ClientWidth     =   4365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4365
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   1365
      TabIndex        =   8
      Top             =   1380
      Width           =   2985
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1365
      Style           =   2  '單純下拉式
      TabIndex        =   9
      Top             =   1710
      Width           =   2985
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3228
      TabIndex        =   11
      Top             =   60
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2220
      TabIndex        =   10
      Top             =   60
      Width           =   972
   End
   Begin VB.TextBox textPA04 
      Height          =   264
      Left            =   3144
      MaxLength       =   2
      TabIndex        =   7
      Top             =   990
      Width           =   375
   End
   Begin VB.TextBox textPA03 
      Height          =   264
      Left            =   2904
      MaxLength       =   1
      TabIndex        =   6
      Top             =   990
      Width           =   255
   End
   Begin VB.TextBox textPA02 
      Height          =   264
      Left            =   2064
      MaxLength       =   6
      TabIndex        =   5
      Top             =   990
      Width           =   855
   End
   Begin VB.TextBox textPA01 
      Height          =   264
      Left            =   1584
      MaxLength       =   3
      TabIndex        =   4
      Top             =   990
      Width           =   495
   End
   Begin VB.TextBox txtCP05_2 
      Height          =   264
      Left            =   3144
      MaxLength       =   7
      TabIndex        =   1
      Top             =   630
      Width           =   1035
   End
   Begin VB.TextBox txtCP05_1 
      Height          =   264
      Left            =   1584
      MaxLength       =   7
      TabIndex        =   0
      Top             =   630
      Width           =   1035
   End
   Begin VB.OptionButton optSel 
      Caption         =   "本所案號："
      CausesValidation=   0   'False
      Height          =   180
      Index           =   1
      Left            =   96
      TabIndex        =   3
      Top             =   1035
      Width           =   1455
   End
   Begin VB.OptionButton optSel 
      Caption         =   "來函收文日："
      CausesValidation=   0   'False
      Height          =   180
      Index           =   0
      Left            =   96
      TabIndex        =   2
      Top             =   675
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "清單及定稿："
      Height          =   180
      Index           =   11
      Left            =   60
      TabIndex        =   13
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label Label2 
      Caption         =   "地址條印表機："
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   12
      Top             =   1740
      Width           =   1260
   End
   Begin VB.Line Line1 
      X1              =   2850
      X2              =   2970
      Y1              =   795
      Y2              =   795
   End
End
Attribute VB_Name = "frm060327"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Dim m_LetterLanguage As String
Dim strPrinter2 As String 'Add By Sindy 2015/9/25
Dim m_PrintRpt1 As Boolean, ff1 As Integer, m_strFileName1 As String 'Add By Sindy 2016/1/27

Private Sub cmdExit_Click()
   Me.Enabled = False
   Unload Me
End Sub

Private Sub cmdOK_Click()
Dim strMsg As String 'Added by Lydia 2019/06/17

   If CheckDataValid() = True Then
      Screen.MousePointer = vbHourglass
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/8 清除查詢印表記錄檔欄位
      'Modified by Lydia 2019/06/17
      'If QueryLetterData() = False Then
      '   MsgBox "沒有符合條件的資料", vbOKOnly + vbCritical, "查詢資料"
      If QueryLetterData(strMsg) = False Then
         If strMsg <> "" Then MsgBox strMsg, vbOKOnly + vbCritical, "查詢資料"
      'end 2019/06/17
      End If
      Clear
      Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub Form_Load()
   
   MoveFormToCenter Me
   
   TextBoxControl
   
   PUB_SetPrinter Me.Name, Combo1
   'Add By Sindy 2015/9/25
   PUB_SetPrinter Me.Name, Combo2, strPrinter2
   '2015/9/25 END
End Sub

Public Sub Clear()
   txtCP05_1 = Empty
   txtCP05_2 = Empty
   textPA01 = Empty
   textPA02 = Empty
   textPA03 = Empty
   textPA04 = Empty
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '列印定稿整批列印清單
   PUB_PrintLetterList strUserNum, "3", Me.Combo2.Text, strPrinter2
   '刪除定稿整批列印資料
   'Modified by Lydia +傳入刪除條件
   'PUB_DeleteLetterList strUserNum
   PUB_DeleteLetterList strUserNum, " and LL02='專利權消滅函' "

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
   
   'Add By Sindy 2015/9/25
   '若請款單印表機變動, 則更新列印設定
   If Me.Combo2.Text <> Me.Combo2.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
   End If
   '2015/9/25 END
   
   Set frm060327 = Nothing
End Sub

Private Sub optSel_Click(Index As Integer)
   TextBoxControl
End Sub

Private Sub TextBoxControl()
   If optSel(0).Value = True Then
      EnableTextBox txtCP05_1, True
      EnableTextBox txtCP05_2, True
      EnableTextBox textPA01, False
      EnableTextBox textPA02, False
      EnableTextBox textPA03, False
      EnableTextBox textPA04, False
      If txtCP05_1.Visible = True Then
         txtCP05_1.SetFocus
      End If
   Else
      EnableTextBox txtCP05_1, False
      EnableTextBox txtCP05_2, False
      EnableTextBox textPA01, True
      EnableTextBox textPA02, True
      EnableTextBox textPA03, True
      EnableTextBox textPA04, True
      If textPA01.Visible = True Then
         textPA01.SetFocus
      End If
   End If
End Sub

Private Sub textPA03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 發證日(起)
Private Sub txtcp05_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(txtCP05_1) = False Then
      If CheckIsTaiwanDate(txtCP05_1, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "來函收文日(起)日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         txtCP05_1_GotFocus
      End If
   End If
End Sub

' 發證日(迄)
Private Sub txtcp05_2_LostFocus()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   If IsEmptyText(txtCP05_2) = False Then
      If CheckIsTaiwanDate(txtCP05_2, False) = False Then
         strTit = "檢核資料"
         strMsg = "來函收文日(迄)日期格式不正確"
         txtCP05_2.SetFocus
         InverseTextBox txtCP05_2
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      Else
         If Not ChkRange(txtCP05_1, txtCP05_2, "來函收文日") Then
            txtCP05_1.SetFocus
            InverseTextBox txtCP05_1
         End If
      End If
   End If
End Sub

Private Sub textPA01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 本所案號的系統別
Private Sub textPA01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textPA01) = False Then
      Select Case textPA01
         Case "FCP":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "本所案號中的系統別不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPA01_GotFocus
      End Select
   End If
End Sub

Private Sub txtCP05_1_GotFocus()
   InverseTextBox txtCP05_1
End Sub

Private Sub txtCP05_2_GotFocus()
   If txtCP05_2 = "" Then txtCP05_2 = txtCP05_1
   InverseTextBox txtCP05_2
End Sub

Private Sub textPA01_GotFocus()
   InverseTextBox textPA01
End Sub

Private Sub textPA02_GotFocus()
   InverseTextBox textPA02
End Sub

Private Sub textPA03_GotFocus()
   InverseTextBox textPA03
End Sub

Private Sub textPA04_GotFocus()
   InverseTextBox textPA04
End Sub

Private Function CheckDataValid() As Boolean

   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
CheckDataValid = False
   
   ' 選項
   If optSel(0).Value = True Then
      ' 來函收文日不可空白
      If IsEmptyText(txtCP05_1) = True Or IsEmptyText(txtCP05_2) = True Then
         strTit = "檢核資料"
         strMsg = "來函收文日不可空白"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         If IsEmptyText(txtCP05_1) = True Then
            txtCP05_1.SetFocus
         Else
            txtCP05_2.SetFocus
         End If
         GoTo EXITSUB
      End If
      If PUB_CheckKeyInDate(txtCP05_1) = -1 Then
         txtCP05_1.SetFocus
         txtCP05_1_GotFocus
         GoTo EXITSUB
      End If
      If PUB_CheckKeyInDate(txtCP05_2) = -1 Then
         txtCP05_2.SetFocus
         txtCP05_2_GotFocus
         GoTo EXITSUB
      End If
      
      ' 範圍
      If Val(txtCP05_1) > Val(txtCP05_2) Then
         strTit = "檢核資料"
         strMsg = "來函收文日範圍不正確"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         txtCP05_1.SetFocus
         GoTo EXITSUB
      End If
   Else
      ' 本所案號
      If IsEmptyText(textPA01) = True Then
         strTit = "檢核資料"
         strMsg = "本所案號系統類別不可空白"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         textPA01.SetFocus
         GoTo EXITSUB
      End If
      ' 本所案號
      If IsEmptyText(textPA02) = True Then
         strTit = "檢核資料"
         strMsg = "本所案號流水號不可空白"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         textPA02.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

'Modified by Lydia 2019/06/17
'Private Function QueryLetterData() As Boolean
Private Function QueryLetterData(ByRef pMsg As String) As Boolean
Dim ET(1 To 3) As String, bolEdit As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'Add By Sindy 2016/1/27
Dim strFileName As String, strFullFileName As String
Dim oFileSys As New FileSystemObject
Dim oFile As File
Dim strMsg As String
Dim strNewCP09 As String
'2016/1/27 END
   
   pMsg = "" 'Added by Lydia 2019/06/17
   ET(1) = "13"
   '1604.專利權消滅
   'Modify By Sindy 2015/9/25 +,GetEmailFlag(CP09) eMail
   'Modified by Lydia 2019/06/17 +CP27,CP05,PA57,PA108
   strExc(0) = "SELECT CP01,CP02,CP03,CP04,CP09,CP25,PA25,PA59,GetEmailFlag(CP09) eMail,CP27,CP05,PA57,PA108 FROM CASEPROGRESS,PATENT WHERE CP01='FCP' AND CP10='1604' AND CP57 IS NULL AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04"
   '收文日
   If optSel(0).Value = True Then
      strExc(0) = strExc(0) & " AND CP05>='" & DBDATE(txtCP05_1.Text) & "' AND CP05<='" & DBDATE(txtCP05_2.Text) & "'"
      bolEdit = False
      pub_QL05 = pub_QL05 & ";" & optSel(0).Caption & txtCP05_1 & "-" & txtCP05_2 'Add By Sindy 2010/12/8
   '本所案號
   Else
      strExc(0) = strExc(0) & " AND CP01='" & textPA01.Text & "' AND CP02='" & textPA02.Text & "' AND CP03='" & Right("0" & textPA03.Text, 1) & "' AND CP04='" & Right("00" & textPA04, 2) & "'"
      'Added by Lydia 2019/06/17 抓最新一道
      strExc(0) = strExc(0) & " and cp09 in (select max(cp09) mno from caseprogress where CP01='" & textPA01.Text & "' AND CP02='" & textPA02.Text & "' AND CP03='" & Right("0" & textPA03.Text, 1) & "' AND CP04='" & Right("00" & textPA04, 2) & "' and cp10='1604' and cp159=0) "
      bolEdit = True
      pub_QL05 = pub_QL05 & ";" & optSel(1).Caption & textPA01 & "-" & textPA02 & "-" & textPA03 & "-" & textPA04 'Add By Sindy 2010/12/8
   End If
   strExc(0) = strExc(0) & " order by eMail,CP01,CP02,CP03,CP04" 'Add By Sindy 2015/9/25
   intI = 1
   Set adoRecordset1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoRecordset1
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/8
        'Added by Lydia 2019/06/17 個案提醒閉卷
        If optSel(1).Value = True Then
            If Trim("" & adoRecordset1.Fields("PA57") & adoRecordset1.Fields("PA108")) <> "" Then
                If MsgBox("本案已閉卷/銷卷，是否繼續列印定稿？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                    QueryLetterData = False
                    Exit Function
                End If
            End If
        End If
            
         'Add by Sindy 2015/9/25
         pub_OsPrinter = PUB_GetOsDefaultPrinter
         PUB_SetOsDefaultPrinter Combo2.Text
         PUB_SetWordActivePrinter
         PUB_RestorePrinter Combo2.Text
         '2015/9/25 END
         
         Do While Not .EOF
            'Modified by Morgan 2014/11/6 +FCP-25875 一律通知
            If "" & .Fields("PA59") = "89" Or bolEdit = True Or .Fields("CP01") & .Fields("CP02") & .Fields("CP03") & .Fields("CP04") = "FCP025875000" Then
'               'Add By Sindy 2015/9/25 列印FCP承辦單
'               Call PUB_PrintFCPEmpBill(.Fields("CP01"), .Fields("CP02"), .Fields("CP03"), .Fields("CP04"), ET(1), .Fields("CP09"), , , "1")
               
               '定稿語文
               m_LetterLanguage = PUB_GetLanguage(.Fields("CP01"), .Fields("CP02"), .Fields("CP03"), .Fields("CP04"))
               ET(2) = .Fields("CP09")
               '消滅日<專用期止日 => 年費逾期未繳
               If Val("" & .Fields("CP25")) < Val("" & .Fields("PA25")) Then
                  ET(3) = "00"
                  'Add by Morgan 2006/11/23 加日文
                  If m_LetterLanguage = "3" Then
                     ET(3) = "02"
                  End If
               Else
                  ET(3) = "01"
                  'Add by Morgan 2006/11/23 加日文
                  If m_LetterLanguage = "3" Then
                     ET(3) = "03"
                  End If
               End If
               'Modify by Morgan 2008/3/21 判斷是否產生電子檔
               'NowPrint ET(2), ET(1), ET(3), bolEdit, strUserNum, 0
               bolEmail = PUB_GetEMailFlag(.Fields("CP01") & .Fields("CP02") & .Fields("CP03") & .Fields("CP04"), , , bolPlusPaper)
               'Add by Morgan 2009/10/20 +判斷是否EMail同時寄紙本
               If bolPlusPaper Then
                  iCopy = 0
               Else
                  iCopy = 1
               End If
               'end 2009/10/20
               If bolEmail Then
                  'Modify By Sindy 2015/10/5
                  'NowPrint ET(2), ET(1), ET(3), bolEdit, strUserNum, 0, , , , iCopy, , True, True
                  NowPrint ET(2), ET(1), ET(3), False, strUserNum, 0, , , , iCopy, , True, True
                  '2015/10/5 END
                  If optSel(1).Value = True Then
                     MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(.Fields("CP01")) & " ]！"
                  End If
               Else
                  'Modify By Sindy 2015/10/5
                  'NowPrint ET(2), ET(1), ET(3), bolEdit, strUserNum, 0
                  NowPrint ET(2), ET(1), ET(3), False, strUserNum, 0
                  '2015/10/5 END
               End If
               'end 2008/3/20
               
               'PUB_PrintLetter ET(2) '直接印出定稿 Add By Sindy 2015/9/25
               'Modify By Sindy 2016/1/27 定稿轉PDF存卷宗區
               strNewCP09 = ET(2)
               strFileName = .Fields("CP01") & .Fields("CP02") & IIf(.Fields("CP04") <> "00", "-" & .Fields("CP03") & "-" & .Fields("CP04"), IIf(.Fields("CP03") <> "0", "-" & .Fields("CP03"), "")) & ".1604.CUS.PDF"
               PUB_DelFtpFile2 strNewCP09, " and cpp02='" & strFileName & "'" '檔案改放 FTP,必須在DB資料刪除前執行
               strSql = "delete from CasePaperPDF where cpp01='" & strNewCP09 & "' and cpp02='" & strFileName & "'"
               cnnConnection.Execute strSql
               If PUB_PrintLetter(ET(2), , , True, strFullFileName) = True Then
                  Call PUB_ChkFileStatus(strFullFileName, False, strMsg)  'Added by Lydia 2022/10/31 判斷檔案是否存在, 超過時間就繼續;
                  Set oFile = oFileSys.GetFile(strFullFileName)
                  If SaveAttFile_PDF(strNewCP09, strFullFileName, strFileName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False) = False Then
                     'Modified by Lydia 2022/10/31 + & ";" & strMsg
                     Call ReadTxt1(.Fields("CP01") & "-" & .Fields("CP02") & "-" & .Fields("CP03") & "-" & .Fields("CP04"), strNewCP09, "定稿轉PDF失敗" & ";" & strMsg)
                  End If
                  Kill strFullFileName
               End If
               '2016/1/27 END
            End If
            
            If bolEdit = False Then '整批
               'Modified by Morgan 2014/11/6 +FCP-25875 一律通知
               If "" & .Fields("PA59") = "89" Or .Fields("CP01") & .Fields("CP02") & .Fields("CP03") & .Fields("CP04") = "FCP025875000" Then
                  If Not bolEmail Or bolPlusPaper Then
'                     'Add By Sindy 2015/9/21 日文定稿才要印地址條
'                     If m_LetterLanguage = "3" Or Val(外專開窗信函啟用日) >= Val(strSrvDate(1)) Then
'                     '2015/9/21 END
                        '新增地址條列表資料
                        pub_AddressListSN = pub_AddressListSN + 1
                        PUB_AddNewAddressList strUserNum, .Fields("CP01"), .Fields("CP02"), .Fields("CP03"), .Fields("CP04"), "" & pub_AddressListSN, "0"
'                     End If
                  End If
               End If
               '新增整批定稿列印清單資料
               'Modified by Lydia 2020/09/24 指定名稱Me.Caption=>專利權消滅函
               PUB_AddNewLetterList "專利權消滅函", txtCP05_1.Text & "-" & Me.txtCP05_2.Text, "" & .Fields("CP01"), "" & .Fields("CP02"), "" & .Fields("CP03"), "" & .Fields("CP04"), IIf(bolEmail, IIf(bolPlusPaper, "Ｅ", "ｅ"), "")
            'Added by Lydia 2019/06/17 有特殊案件已閉卷尚需報告，則程序會個案下定稿，到時自動將發文日"111111"拿掉，如此案件又可自大批上發文
            ElseIf textPA01 <> "" And textPA02 <> "" And "" & .Fields("CP27") = "19221111" And Trim("" & .Fields("PA57") & .Fields("PA108")) <> "" Then
                    '將發文日"111111"拿掉並且上承辦期限
                    strExc(1) = CompDate(2, 10, "" & .Fields("cp05"))
                    strSql = "Update Caseprogress set cp27=null,cp48=" & IIf(Val(strExc(1)) < strSrvDate(1), strSrvDate(1), strExc(1)) & _
                                " where cp09='" & .Fields("cp09") & "' and cp10='1604' "
                    Pub_SeekTbLog strSql
                    cnnConnection.Execute strSql, intI
            End If
            .MoveNext
         Loop
         
         'Add by Sindy 2015/9/25
         PUB_SetOsDefaultPrinter pub_OsPrinter
         PUB_RestorePrinter strPrinter2
         '2015/9/25 END
         
      End With
      QueryLetterData = True
      'Modify By Sindy 2016/1/27
      If m_PrintRpt1 = True Then
         Close ff1
         strMsg = "請至下列位置列印檢核表：" & PUB_Getdesktop & "\" & m_strFileName1
      End If
      MsgBox "定稿列印完畢！ " & strMsg, vbInformation
      '2016/1/27 END
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/8
      pMsg = "沒有符合條件的資料" 'Added by Lydia 2019/06/17
   End If
End Function

'Add By Sindy 2016/1/27
'資料檢核表
Private Sub ReadTxt1(strCaseNo As String, strRecvNo As String, strNote As String)
Dim i As Integer
Dim strTemp(1 To 7) As String
   
   If m_PrintRpt1 = False Then
      m_PrintRpt1 = True
      If ff1 > 0 Then Close #ff1
      ff1 = FreeFile
      m_strFileName1 = Me.Caption & txtCP05_1 & "-" & txtCP05_2 & "資料檢核表.txt"
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

Public Sub SetInputFocus()
   If optSel(0).Value = True Then
      txtCP05_1.SetFocus
   Else
      textPA01.SetFocus
   End If
End Sub
