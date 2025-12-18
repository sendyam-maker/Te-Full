VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040324_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "年費逾期補繳通知函"
   ClientHeight    =   4572
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4572
   ScaleWidth      =   7980
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   3
      Left            =   1395
      MaxLength       =   8
      TabIndex        =   2
      Top             =   3210
      Width           =   1290
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   1
      Left            =   1395
      MaxLength       =   2
      TabIndex        =   1
      Top             =   2835
      Width           =   615
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   2
      Left            =   1395
      MaxLength       =   1
      TabIndex        =   3
      Top             =   3990
      Width           =   375
   End
   Begin VB.TextBox text1 
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   1395
      MaxLength       =   20
      TabIndex        =   0
      Top             =   510
      Width           =   1635
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   5520
      TabIndex        =   4
      Top             =   30
      Width           =   972
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6525
      TabIndex        =   5
      Top             =   30
      Width           =   1215
   End
   Begin MSForms.ComboBox cbo 
      Height          =   300
      Index           =   1
      Left            =   1395
      TabIndex        =   7
      Top             =   1695
      Width           =   6465
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "11404;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cbo 
      Height          =   300
      Index           =   0
      Left            =   1395
      TabIndex        =   6
      Top             =   1290
      Width           =   6465
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "11404;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日："
      Height          =   180
      Index           =   10
      Left            =   225
      TabIndex        =   24
      Top             =   3255
      Width           =   1080
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   4
      Left            =   1395
      TabIndex        =   23
      Top             =   3600
      Width           =   1995
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "已繳年度："
      Height          =   180
      Index           =   9
      Left            =   405
      TabIndex        =   22
      Top             =   2490
      Width           =   900
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   1395
      TabIndex        =   21
      Top             =   930
      Width           =   2235
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   1
      Left            =   405
      TabIndex        =   20
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   180
      Index           =   0
      Left            =   405
      TabIndex        =   19
      Top             =   570
      Width           =   900
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   18
      Top             =   3600
      Width           =   1995
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   3
      Left            =   1395
      TabIndex        =   17
      Top             =   2460
      Width           =   6435
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "是否修改定稿：             (Y : Word )"
      Height          =   180
      Index           =   8
      Left            =   75
      TabIndex        =   16
      Top             =   4035
      Width           =   2670
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   1395
      TabIndex        =   15
      Top             =   2100
      Width           =   1755
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "最後期限："
      Height          =   180
      Index           =   7
      Left            =   4320
      TabIndex        =   14
      Top             =   3637
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "原年費期限："
      Height          =   180
      Index           =   6
      Left            =   225
      TabIndex        =   13
      Top             =   3630
      Width           =   1080
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "未繳年度："
      Height          =   180
      Index           =   5
      Left            =   405
      TabIndex        =   12
      Top             =   2880
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "專利種類："
      Height          =   180
      Index           =   4
      Left            =   405
      TabIndex        =   11
      Top             =   2130
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "此案已閉卷"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   10
      Top             =   840
      Width           =   1485
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "申請人１："
      Height          =   180
      Index           =   3
      Left            =   405
      TabIndex        =   9
      Top             =   1755
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱："
      Height          =   180
      Index           =   2
      Left            =   405
      TabIndex        =   8
      Top             =   1350
      Width           =   900
   End
End
Attribute VB_Name = "frm040324_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (cbo)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'整理 by Morgan 2005/7/11
Option Explicit

Dim ET01 As String '定稿別
Dim ET02 As String '總收文號(或本所案號&案件性質)
Dim ET03 As String '處理狀況
Dim m_strPA01 As String '本所案號
Dim m_strPA02 As String '本所案號
Dim m_strPA03 As String '本所案號
Dim m_strPA04 As String '本所案號
Dim m_strPA08 As String '專利種類
Dim m_strPA09 As String '申請國家
Dim m_strPA14 As String 'Add by Morgan 2004/9/6 公告日
Dim m_strPA26 As String '申請人1
Dim m_strPA72 As String '已繳年度
Dim m_strPA75 As String 'Added by Morgan 2014/7/22
Dim m_strPA57 As String 'Added by Lydia 2023/08/29 是否閉卷
Dim m_strMaxPA72 As String '目前繳費年度
Dim strNextTime As String '下次期限天數或月數
Dim strSql As String
'Add By Cheng 2003/04/23
Dim m_blnFirstShow As Boolean '第一次顯示
Dim m_NewCP09 As String 'Add by Morgan 2008/12/12
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/7 END
Dim m_bolFMP As Boolean, m_CP13 As String, m_CP12 As String 'Added by Morgan 2022/10/17
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/08/29 是否為寰華

Private Sub cmdExit_Click()
    Unload Me
    frm040324.Show
End Sub

Private Sub cmdOK_Click()
Dim bolChk As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   If TxtValidate = False Then Exit Sub
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   '取得收文號
   'Modify by Morgan 2008/12/12 改要新增來函"通知年費逾期"
   'ET02 = ""
   'StrSQLa = "Select CP09 From CaseProgress Where " & ChgCaseprogress(m_strPA01 & m_strPA02 & m_strPA03 & m_strPA04) & " And CP09 <'C' Order By CP05 Desc,CP09 Desc "
   'rsA.CursorLocation = adUseClient
   'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   'If rsA.RecordCount > 0 Then
   '    ET02 = "" & rsA.Fields(0).Value
   'End If
   'If rsA.State <> adStateClosed Then rsA.Close
   'Set rsA = Nothing
   
   'Add By Sindy 2022/7/1
   If m_strIR01 <> "" And Left(Pub_StrUserSt03, 2) = "F2" Then
      If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
   End If
   '2022/7/1 END
   
   If FormSave = False Then
      MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
      Me.Enabled = True
      Screen.MousePointer = vbDefault
   Else
      
      If text1(2) = "Y" Then
         bolChk = True
      Else
         bolChk = False
      End If
      
      '列印定稿
      Select Case m_strPA09
      Case "000"
         ET03 = "01"
         'Added by Morgan 2012/9/19
         '新法每月+20%到100%(2倍)止
         'Modified by Morgan 2012/12/6 用系統日判斷--郭
         'If Val(Label1(5)) >= 1020101 Then
         'Modified by Morgan 2013/6/17 +大->台定稿
         'If strSrvDate(1) >= "20130101" Then
         If PUB_CheckCuNation(m_strPA26, m_strPA01, m_strPA02, m_strPA03, m_strPA04) = "1" Then
            ET03 = "05"
         Else
            ET03 = "04"
         End If
         
         
      Case "020"
         ET03 = "02"
      Case "013"
         ET03 = "03"
      End Select
      
      ET02 = m_NewCP09 'Add by Morgan 2008/12/12
      StartLetter "19", ET02, ET03
      
      'Modified by Morgan 2016/12/16
      'NowPrint ET02, "19", ET03, bolChk, strUserNum, , , , , , , , , , , , , m_NewCP09
      If Left(Pub_StrUserSt03, 1) = "F" Then
         NowPrint ET02, "19", ET03, bolChk, strUserNum
      Else
         'Modified by Morgan 2023/4/11 +m_bolFMP(FMP案不再列印紙本)
         NowPrint ET02, "19", ET03, bolChk, strUserNum, , , , , , , , , , , , , m_NewCP09, , , , , m_bolFMP
         If bolChk Then
            frm1105_1.m_RecNo = m_NewCP09
            frm1105_1.m_PdfName = PUB_CaseNo2FileName(m_strPA01, m_strPA02, m_strPA03, m_strPA04) & ".1605.CUS.PDF"
            frm1105_1.Show
         End If
      End If
      'end 2016/12/16
      
      'Added by Morgan 2022/10/17
      If m_bolFMP Then
         strUserNum = strFMPNum
         StartLetter2 "19", ET02, "51"
         NowPrint ET02, "19", "51", False, strUserNum
         strUserNum = strUser1Num
      End If
      'end 2022/10/17
                  
   'Remove by Morgan 2008/8/13 改開窗定稿
   '    'Add By Cheng 2003/04/23
   '    '新增地址條列表資料
   '    pub_AddressListSN = pub_AddressListSN + 1
   '    PUB_AddNewAddressList strUserNum, m_strPA01, m_strPA02, m_strPA03, m_strPA04, "" & pub_AddressListSN, "0"
      frm040324.Show
      frm040324.Clear
      Me.Enabled = True
      Screen.MousePointer = vbDefault
      'Add By Sindy 2017/12/29
      If Me.m_strIR01 <> "" Then
         Unload frm040324
         'Modify By Sindy 2022/5/20
         'frm04010519.GoNext
         Forms(0).Tmpfrm04010519.GoNext
         Set Forms(0).Tmpfrm04010519 = Nothing
         '2022/5/20 END
      End If
      '2017/12/29 END
      Unload Me
   End If
   
End Sub

Private Sub StartLetter(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)
   Dim strTxt(1 To 10) As String, i As Integer, j As Integer, strTmp As String
   
   EndLetter ET01, ET02, ET03, strUserNum
   i = 0
   
   'Added by Morgan 2012/9/19
   If ET03 = "04" Or ET03 = "05" Then
   'Modified by Lydia 2015/01/07 採共用模組
'      strExc(0) = "Select YF06,YF07 From PatentYearFee Where YF01='" & m_strPA09 & "' AND YF02='" & m_strPA08 & "' AND YF03='Y00000001' AND YF04='605' AND YF05=" & Val(Text1(1))
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         '服務費
'         strExc(1) = "" & RsTemp("YF06")
'         '規費
'         strExc(2) = "" & RsTemp("YF07")
      strExc(0) = PUB_GetYF0607(m_strPA09, m_strPA08, m_strPA26, "605", text1(1), text1(1), "1", strExc(1), strExc(2))
      If strExc(0) = "0" Then strExc(1) = "": strExc(2) = ""
      
      If Val(strExc(0)) > 0 Then
         If Val(text1(1)) < 7 Then
            '可減免
            If PUB_GetCaseDiscStat(m_strPA01 & m_strPA02 & m_strPA03 & m_strPA04) = "Y" Then
               If Val(text1(1)) > 3 Then
                  strExc(2) = Val(strExc(2)) - 1200
               Else
                  strExc(2) = Val(strExc(2)) - 800
               End If
            End If
         End If
'      End If
      End If
   'end 2015/01/07
   
      'Added by Morgan 2013/2/21
      '專利處大對台年費服務費+500 --郭雅娟
      strExc(4) = PUB_GetStaffST15(PUB_GetAKindSalesNo(m_strPA01, m_strPA02, m_strPA03, m_strPA04), "1")
      If Left(strExc(4), 2) = "P1" Then
         strExc(1) = Val(strExc(1)) + 500
      End If
      'end 2013/2/21
      
      'Modified by Morgan 2013/2/21 要加服務費
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','原年費','" & Format(Format(Val(strExc(1)) + Val(strExc(2)), "#,###"), String(6, "@")) & "')"
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','年費1','" & Format(Format(Val(strExc(1)) + Val(strExc(2)) * 1.2, "#,###"), String(6, "@")) & "')"
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','年費2','" & Format(Format(Val(strExc(1)) + Val(strExc(2)) * 1.4, "#,###"), String(6, "@")) & "')"
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','年費3','" & Format(Format(Val(strExc(1)) + Val(strExc(2)) * 1.6, "#,###"), String(6, "@")) & "')"
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','年費4','" & Format(Format(Val(strExc(1)) + Val(strExc(2)) * 1.8, "#,###"), String(6, "@")) & "')"
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','年費5','" & Format(Format(Val(strExc(1)) + Val(strExc(2)) * 2, "#,###"), String(6, "@")) & "')"
         
      'end 2013/2/21
   End If
   
   '年費法定期限
   i = i + 1
   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','年費法定期限'," & CNULL(DBDATE(Val(Me.Label1(4).Caption))) & ")"
   '未繳年度
   i = i + 1
   '2008/11/6 modify by sonia P-073430
   'strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
   '   "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','第幾年至幾年費'," & CNULL(Val(Me.Text1(1).Text)) & ")"
   If m_strPA09 = "013" Then
      If m_strPA08 = "2" Then
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','第幾年至幾年費'," & CNULL(Val(Me.text1(1).Text)) & ")"
      Else
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','第幾年至幾年費'," & CNULL(Val(Me.text1(1).Text) - 1) & ")"
      End If
   Else
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','第幾年至幾年費'," & CNULL(Val(Me.text1(1).Text)) & ")"
   End If
   '2008/11/6 end
   '下次天數或月數
   i = i + 1
   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','列印備註'," & CNULL(strNextTime) & ")"
   '最後期限(本所)
   i = i + 1
   Dim strDate(0 To 3) As String
   strDate(1) = m_strPA01     '系統別
   strDate(2) = m_strPA09     '申請國家
   strDate(3) = ChangeTStringToWString(Me.Label1(5).Caption)  '下次法定期限
   GetCtrlDT strDate()
   strTmp = ChangeWStringToTString(strDate(0))
    'Modify By Cheng 2003/04/23
'   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','下次繳年費日'," & DBDATE(strTmp) & ")"
   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','下次繳年費日','" & DBDATE(strTmp) & "')"
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(i, strTxt) Then
   If Not ClsLawExecSQL(i, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If

End Sub

'Added by Morgan 2022/10/17
Private Sub StartLetter2(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)
   Dim strTxt(1 To 10) As String, i As Integer, j As Integer, strTmp As String
   
   EndLetter ET01, ET02, ET03, strUserNum
   i = 0
   '年費法定期限
   i = i + 1
   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','年費法定期限'," & CNULL(DBDATE(Val(Me.Label1(4).Caption))) & ")"
   '未繳年度
   i = i + 1
   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','第幾年至幾年費'," & CNULL(Val(Me.text1(1).Text)) & ")"
   '補繳法定期限
   i = i + 1
   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','補繳法定期限'," & CNULL(DBDATE(Val(Me.Label1(5).Caption))) & ")"
    
   If Not ClsLawExecSQL(i, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Function Search() As Boolean
     Search = False
     If text1(0) = "" Then
        MsgBox "申請案號不得空白，請重新輸入 !", vbCritical
        Me.text1(0).SetFocus
        TextInverse Me.text1(0)
        Exit Function
     End If
    intI = 0
    'Modified by Lydia 2023/08/29 +pa57
    strExc(0) = "select pa01,pa02,pa03,pa04,pa09,pa26,pa72,pa08,pa14,pa75,pa57 from patent where PA23='1' AND " & _
                    "pa11='" & text1(0) & "' And " & ChgPatent(Replace(Me.Label1(0).Caption, "-", ""))
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
    If intI <> 1 Then Exit Function
    If intI = 1 And RsTemp.Fields(0) = "P" Then
       cmdOK.Default = True
       m_strPA01 = RsTemp.Fields(0)
       m_strPA02 = RsTemp.Fields(1)
       m_strPA03 = RsTemp.Fields(2)
       m_strPA04 = RsTemp.Fields(3)
       m_strPA09 = RsTemp.Fields(4)
       m_strPA26 = RsTemp.Fields(5)
       m_strPA75 = "" & RsTemp.Fields("pa75") 'Added by Morgan 2014/7/22
       m_strPA14 = "" & RsTemp.Fields("pa14") 'Add by Morgan 2004/9/6
       m_strPA57 = "" & RsTemp.Fields("pa57")  'Added by Lydia 2023/08/29
       m_strPA72 = ""
       If IsNull(RsTemp.Fields(6)) = False Then m_strPA72 = RsTemp.Fields(6)
       m_strPA08 = RsTemp.Fields(7)
       ClearControl
       ReadData
       GetNextPayData
       'Added by Morgan 2022/10/17
       m_CP13 = PUB_GetAKindSalesNo(m_strPA01, m_strPA02, m_strPA03, m_strPA04)
       m_CP12 = GetSalesArea(m_CP13)
       If Left(m_CP12, 1) = "F" And m_strPA09 <> "000" Then
         m_bolFMP = True
       Else
         m_bolFMP = False
       End If
       'end 2022/10/17
       'Added by Lydia 2023/08/29 判斷寰華案
       m_bolFMP2 = False
       If m_bolFMP = True Then
          m_bolFMP2 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, m_strPA01, m_strPA02, m_strPA03, m_strPA04)
       End If
       'end 2023/08/29
       Search = True
    End If
    
'Removed by Morgan 2016/12/16 改可呼叫維護視窗
'   'Added by Morgan 2014/7/2
'   '台灣案定稿電子化只能到維護程式修改
'   If m_strPA09 = "000" Then
'      Text1(2).Enabled = False
'
'   'Added by Morgan 2016/6/16
'   ElseIf 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
'      Text1(2).Enabled = False
'   'end 2016/6/16
'   End If
'   'end 2014/7/2
'end 2016/12/16
End Function

Private Sub Form_Activate()
   'Add By Cheng 2003/04/23
   If m_blnFirstShow = True Then
      m_blnFirstShow = False
      
      'Added by Sindy 2017/12/29
      If m_strIR01 <> "" Then
         text1(3) = frm040324.m_RDate
      Else
      '2017/12/29 END
         text1(3) = strSrvDate(2) 'Added by Morgan 2016/12/16
      End If
      
      If Search = False Then
          frm040324.Show
          Unload Me
          
'Removed by Morgan 2018/8/7 未繳年度跳開有檢查
'      'Add by Morgan 2004/2/2
'      '控制年費期限未到期不可列印
'      'Modify by Morgan 2010/8/11 百年蟲
'      'ElseIf Me.Label1(4) >= strSrvDate(2) Then
'      ElseIf Val(Label1(4)) >= Val(strSrvDate(2)) Then
'          MsgBox "年費期限尚未到期！！", vbExclamation
'          frm040324.Show
'          Unload Me
      End If
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   m_blnFirstShow = True
   
   'Add By Sindy 2017/12/29
   m_strIR01 = frm040324.m_strIR01
   m_strIR02 = frm040324.m_strIR02
   m_strIR03 = frm040324.m_strIR03
   m_strIR04 = frm040324.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/29 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Morgan 2020/4/10
    Set frm040324_1 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Me.text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
   Case 2 '是否修改定稿
      KeyAscii = UpperCase(KeyAscii)
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         MsgBox "是否修改定稿只能輸入 Y !!!", vbExclamation + vbOKOnly
         KeyAscii = 0
      End If
   End Select
End Sub

Private Function ReadData() As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

    ReadData = False
    
    StrSQLa = "SELECT * FROM PATENT,CUSTOMER,PATENTTRADEMARKMAP WHERE " & ChgPatent(m_strPA01 & m_strPA02 & m_strPA03 & m_strPA04) & " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        '本所案號
        Me.Label1(0).Caption = m_strPA01 & "-" & m_strPA02 & "-" & m_strPA03 & "-" & m_strPA04
        '是否閉卷
        If "" & rsA.Fields("PA57") = "Y" Then Me.Label1(1).Caption = "此案已閉卷"
        '專利名稱(中-->英-->日)
        Me.cbo(0).AddItem "中：" & rsA.Fields("PA05").Value
        Me.cbo(0).AddItem "英：" & rsA.Fields("PA06").Value
        Me.cbo(0).AddItem "日：" & rsA.Fields("PA07").Value
        Me.cbo(0).ListIndex = 0
        '申請人1(中-->英-->日)
        Me.cbo(1).AddItem "中：" & rsA.Fields("CU04").Value
        Me.cbo(1).AddItem "英：" & Trim("" & rsA.Fields("CU05").Value & " " & rsA.Fields("CU88").Value & " " & rsA.Fields("CU89").Value & " " & rsA.Fields("CU90").Value & " ")
        Me.cbo(1).AddItem "日：" & rsA.Fields("CU06").Value
        Me.cbo(1).ListIndex = 0
        '專利種類
        Me.Label1(2).Caption = "" & rsA("PTM03").Value
        '已繳年度
        Me.Label1(3).Caption = m_strPA72
        ReadData = True
    Else
        MsgBox "基本檔無此案號資料!!!", vbExclamation + vbOKOnly
        Me.text1(0).SetFocus
        Text1_GotFocus 0
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End Function

Private Sub ClearControl()
   Me.cbo(0).Clear
   Me.cbo(1).Clear
   Me.Label1(0).Caption = Empty
   Me.Label1(1).Caption = Empty
   Me.Label1(2).Caption = Empty
   Me.Label1(3).Caption = Empty
   Me.Label1(4).Caption = Empty
   Me.Label1(5).Caption = Empty
   Me.text1(1).Text = Empty
   Me.text1(2).Text = Empty
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim strTmp(0 To 5) As String, strTmp1(0 To 5) As String, varTmp As Variant
   
    On Error GoTo ErrorHandler
   Select Case Index
   Case 1 '未繳年度
      If Me.text1(Index).Text <> "" Then
         If Val(Me.text1(Index).Text) <= Val(m_strMaxPA72) Then
            MsgBox "未繳年度不可小於或等於已繳費年度!!!", vbExclamation + vbOKOnly
            Cancel = True
         Else
            StrSQLa = "SELECT * FROM NATION WHERE NA01='" & m_strPA09 & "'"
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               If m_strPA08 = "1" Then
                  If Val(Me.text1(Index).Text) > Val("" & rsA("NA07").Value) Then
                     MsgBox "未繳年度不可大於需繳費的年度(" & Val("" & rsA("NA07").Value) & ")", vbExclamation + vbOKOnly
                     Cancel = True
                  End If
               ElseIf m_strPA08 = "2" Then
               
                  'Add by Morgan 2004/9/6 93.7.1以前公告的台灣新型專用期為12年
                  If m_strPA09 = "000" And Val(m_strPA14) < 20040701 Then
                     If Val(Me.text1(Index).Text) > 12 Then
                        MsgBox "未繳年度不可大於需繳費的年度(12)", vbExclamation + vbOKOnly
                        Cancel = True
                     End If
                  Else
                  '2004/9/6 end
                  
                     If Val(Me.text1(Index).Text) > Val("" & rsA("NA09").Value) Then
                        MsgBox "未繳年度不可大於需繳費的年度(" & Val("" & rsA("NA09").Value) & ")", vbExclamation + vbOKOnly
                        Cancel = True
                     End If
                     
                  End If
               ElseIf m_strPA08 = "3" Then
                  'Modify by Morgan 2004/10/21
                  'If Val(Me.Text1(Index).Text) > Val("" & rsA("NA09").Value) Then
                  If Val(Me.text1(Index).Text) > Val("" & rsA("NA11").Value) Then
                     MsgBox "未繳年度不可大於需繳費的年度(" & Val("" & rsA("NA11").Value) & ")", vbExclamation + vbOKOnly
                     Cancel = True
                  End If
               End If
            Else
               MsgBox "無此專利種類的繳費年度資料!!!", vbExclamation + vbOKOnly
               Cancel = True
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '若修改未繳年度, 需計算該年度之年費期限及最後期限
            If Cancel <> True Then
               strTmp1(0) = ""
               strTmp1(1) = m_strPA01: strTmp1(2) = m_strPA02: strTmp1(3) = m_strPA03: strTmp1(4) = m_strPA04:
               If GetMoneyDate(m_strPA08, m_strPA09, strTmp1, strTmp(1), strTmp(2), strTmp(3)) = True Then
                  varTmp = Split(strTmp(2), ",")
                  '92.12.23 MODIFY BY SONIA
                  'Me.Label1(4).Caption = ChangeWStringToTString(CompDate(0, Val(varTmp(Val(Text1(1).Text) - 2)), TransDate(strTmp(3), 2)))
                  '2008/11/6 MODIFY BY SONIA P-073430 香港新型,設計
                  'If (Val(Text1(1).Text) - 2) > UBound(varTmp) Then
                  '   Me.Label1(4).Caption = ChangeWStringToTString(CompDate(0, Val(varTmp(UBound(varTmp))) - 1, TransDate(strTmp(3), 2)))
                  'ElseIf (Val(Text1(1).Text) - 2) < LBound(varTmp) Then
                  '   Me.Label1(4).Caption = ChangeWStringToTString(CompDate(0, Val(varTmp(LBound(varTmp))) - 1, TransDate(strTmp(3), 2)))
                  'Else
                  '   Me.Label1(4).Caption = ChangeWStringToTString(CompDate(0, Val(varTmp(Val(Text1(1).Text) - 2)), TransDate(strTmp(3), 2)))
                  'End If
                  If m_strPA09 = "013" And m_strPA08 <> "1" Then
                     Me.Label1(4).Caption = ChangeWStringToTString(CompDate(0, Val(Me.text1(Index).Text), strTmp(3)))
                  Else
                     Me.Label1(4).Caption = ChangeWStringToTString(CompDate(0, Val(Me.text1(Index).Text) - 1, strTmp(3)))
                  End If
                  '2008/11/6 END
                  If Val(Me.Label1(4).Caption) + 19110000 > strSrvDate(1) Then
                     MsgBox "原年費期限不可大於系統日!!!", vbExclamation + vbOKOnly
                     Cancel = True
                  End If
                  If GetCF12(m_strPA01, m_strPA09, 年費) <> 0 Then
                      Me.Label1(5).Caption = ChangeWStringToTString(CompDate(2, (GetCF12(m_strPA01, m_strPA09, 年費)), Format(Me.Label1(4).Caption)))
                      strNextTime = GetCF12(m_strPA01, m_strPA09, 年費) 'Added by Morgan 2018/8/14
                  Else
                      Me.Label1(5).Caption = ChangeWStringToTString(CompDate(1, (GetCF28(m_strPA01, m_strPA09, 年費)), Format(Me.Label1(4).Caption)))
                      strNextTime = GetCF28(m_strPA01, m_strPA09, 年費) 'Added by Morgan 2018/8/14
                  End If
                  If Cancel <> True And Val(Me.Label1(5).Caption) + 19110000 <= strSrvDate(1) Then
                     MsgBox "最後期限不可小於或等於系統日!!!", vbExclamation + vbOKOnly
                     Cancel = True
                  End If
               Else
                  Me.Label1(4).Caption = ""
                  Me.Label1(5).Caption = ""
               End If
            Else
               Me.Label1(4).Caption = ""
               Me.Label1(5).Caption = ""
            End If
         End If
      End If
   End Select
   If Cancel = True Then Text1_GotFocus Index
   Exit Sub
ErrorHandler:
    MsgBox Err.Description
    Cancel = True
    Text1_GotFocus Index
End Sub

Private Function TxtValidate() As Boolean
   Dim objTxt As Object
   Dim ii As Integer
   Dim Cancel As Boolean
   
   TxtValidate = False
   If Me.text1(1).Text = "" Then
      MsgBox "請輸入未繳年度!!!", vbExclamation + vbOKOnly
      Me.text1(1).SetFocus
      TextInverse Me.text1(1)
      Exit Function
   Else
      Cancel = False
      Text1_Validate (1), Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.Label1(4).Caption = "" Then
      MsgBox "無原年費期限資料!!!", vbExclamation + vbOKOnly
      Me.text1(1).SetFocus
      TextInverse Me.text1(1)
      Exit Function
   End If
   
   'Added by Morgan 2016/12/16
   If text1(3) = "" Then
      MsgBox "請輸入來函收文日!!", vbExclamation
      text1(3).SetFocus
      Exit Function
   ElseIf ChkDate(text1(3)) = False Then
      text1(3).SetFocus
      Exit Function
   ElseIf DBDATE(text1(3)) > strSrvDate(1) Then
      MsgBox "來函收文日不可晚於系統日!!", vbExclamation
      text1(3).SetFocus
      Exit Function
   ElseIf DBDATE(text1(3)) <= (strSrvDate(1) - 10000) Then
      MsgBox "來函收文日不可為1年以前!!", vbExclamation
      text1(3).SetFocus
      Exit Function
   End If
   'end 2016/12/16
   
   'Add by Morgan 2008/12/12 重複通知檢查
   If DupeCheck = False Then
      Exit Function
   End If
   
   TxtValidate = True
End Function

Private Function DupeCheck() As Boolean
   DupeCheck = True
   strExc(0) = "select cp10, cp27 from caseprogress where cp01='" & m_strPA01 & "' and cp02='" & m_strPA02 & "' and cp03='" & m_strPA03 & "'  and cp04='" & m_strPA04 & "' and cp57 is null order by cp27 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.Fields("cp10") = "1605" Then
         If MsgBox("本案件已於 " & ChangeWStringToTDateString("" & RsTemp.Fields(1)) & " 通知逾期補繳，是否再次通知？", vbYesNo + vbDefaultButton2) = vbNo Then
            DupeCheck = False
         End If
      End If
   End If
End Function
'Add By Cheng 2003/01/03
'取得下次繳費年度資料
Private Sub GetNextPayData()
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim arrPA72
Dim ArrYear
Dim strNextPA72 '下次繳費年度
Dim ii As Integer
Dim jj As Integer
    
    strNextTime = ""
    '取得下次繳費年度
    StrSqlB = "Select DECODE(" & m_strPA08 & ", '1', NA21, '2', NA23, NA25 ) From Nation Where NA01='" & m_strPA09 & "' "
    rsB.CursorLocation = adUseClient
    rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
    If rsB.RecordCount > 0 Then
        If "" & rsB.Fields(0).Value <> "" Then
            If "" & m_strPA72 <> "" Then
                arrPA72 = Split(m_strPA72, ",")
                m_strMaxPA72 = arrPA72(UBound(arrPA72))
                'Modify by Morgan 2005/7/11 新型舊法12年
                'ArrYear = Split(rsB.Fields(0).Value, ",")
                '2008/11/6 modify by sonia
                'If m_strPA14 < "20040701" Then
                If m_strPA09 = "000" And m_strPA08 = "2" And Val(m_strPA14) < "20040701" Then
                  ArrYear = Split("1,2,3,4,5,6,7,8,9,10,11,12", ",")
                Else
                  ArrYear = Split(rsB.Fields(0).Value, ",")
                End If
               
               jj = -100
               strNextPA72 = ""
               For ii = LBound(ArrYear) To UBound(ArrYear)
                   strNextPA72 = ArrYear(ii)
                   If ii = jj + 1 Then Exit For
                   If m_strMaxPA72 = ArrYear(ii) Then jj = ii
               Next ii
               Me.text1(1).Text = strNextPA72
            Else
               m_strMaxPA72 = 0
               '92.12.23 MODIFY BY SONIA
               ArrYear = Split(rsB.Fields(0).Value, ",")
               Me.text1(1).Text = ArrYear(LBound(ArrYear))
            End If
        Else
            Me.text1(1).Text = ""
        End If
    Else
        Me.text1(1).Text = ""
    End If
    If rsB.State <> adStateClosed Then rsB.Close
    Set rsB = Nothing
    
'Modified by Morgan 2018/8/7 下一程序期限可能有+6個月管制期 Ex:P-89002
'    '取得原年費期限及最後期限
'    StrSqlB = "Select MAX(NP09) From NextProgress Where " & ChgNextProgress(m_strPA01 & m_strPA02 & m_strPA03 & m_strPA04) & " And NP07='" & 年費 & "'"
'    rsB.CursorLocation = adUseClient
'    rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
'    If rsB.RecordCount > 0 Then
'        Me.Label1(4).Caption = ChangeWStringToTString("" & rsB.Fields(0).Value)
'        If GetCF12(m_strPA01, m_strPA09, 年費) <> 0 Then
'            Me.Label1(5).Caption = ChangeWStringToTString(CompDate(2, (GetCF12(m_strPA01, m_strPA09, 年費)), Format(Me.Label1(4).Caption)))
'            strNextTime = GetCF12(m_strPA01, m_strPA09, 年費)
'        Else
'            Me.Label1(5).Caption = ChangeWStringToTString(CompDate(1, (GetCF28(m_strPA01, m_strPA09, 年費)), Format(Me.Label1(4).Caption)))
'            strNextTime = GetCF28(m_strPA01, m_strPA09, 年費)
'        End If
'    Else
'        Me.Label1(4).Caption = ""
'        Me.Label1(5).Caption = ""
'    End If
'    If rsB.State <> adStateClosed Then rsB.Close
'    Set rsB = Nothing
   Text1_Validate 1, False
'end 2018/8/7

End Sub

'取得案件收費表的下次期限天數
Private Function GetCF12(strCF01 As String, strCF02 As String, strCF03 As String) As String
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetCF12 = "0"
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
'91.10.31 MODIFY BY SONIA
'strSQLA = "Select CF12,CF28 From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "' AND CF12 IS NOT NULL"
StrSQLa = "Select CF12 From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "'"
'91.10.31 END
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount <> 0 Then
   If Not IsNull(rsA.Fields(0).Value) Then
      GetCF12 = rsA.Fields(0).Value
   End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Add By SONIA 2002/10/31
'取得案件收費表的下次期限月份
Private Function GetCF28(strCF01 As String, strCF02 As String, strCF03 As String) As String
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetCF28 = "0"
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
StrSQLa = "Select CF28 From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "'"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount <> 0 Then
   If Not IsNull(rsA.Fields(0).Value) Then
      GetCF28 = rsA.Fields(0).Value
   End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

Private Function FormSave() As Boolean
   'Dim stCP13 As String, stCP12 As String 'Removed by Morgan 2022/10/17
   Dim cp() As String
   ReDim cp(1 To TF_CP) As String
   Dim stUpdatePA As String
 
On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
On Error GoTo ErrorHandler1
   
   'Removed by Morgan 2022/10/17
   'stCP13 = PUB_GetAKindSalesNo(m_strPA01, m_strPA02, m_strPA03, m_strPA04)
   'stCP12 = GetSalesArea(stCP13)
   'end 2022/10/17
   cp(1) = m_strPA01
   cp(2) = m_strPA02
   cp(3) = m_strPA03
   cp(4) = m_strPA04
   'Modified by Morgan 2016/12/16
   'cp(5) = strSrvDate(1)
   cp(5) = DBDATE(text1(3))
   'end 2016/12/16
   cp(9) = 主管機關來函
   cp(10) = "1605"
   'Modified by Morgan 2022/10/17
   'cp(12) = stCP12
   'cp(13) = stCP13
   cp(12) = m_CP12
   cp(13) = m_CP13
   'end 2022/10/17
   cp(14) = strUserNum
   cp(27) = strSrvDate(1)
   cp(20) = "N"
   cp(26) = "N"
   cp(32) = "N"
   cp(64) = "未繳年度:" & text1(1)
   'Modified by Morgan 2016/12/16
   'cp(119) = strSrvDate(1) 'Added by Morgan 2012/11/5
   cp(119) = DBDATE(text1(3))
   'end 2016/12/16
   
   strSql = GetCPSQL(cp(), False)
   cnnConnection.Execute strSql, intI
   
   m_NewCP09 = cp(9)
   If m_strPA09 <> "000" Then
      '抓最新的AB類發文代理人更新
      Pub_UpdateFromMaxCP27 cp(1), cp(2), cp(3), cp(4)
      'Added by Morgan 2016/6/16
      If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
         PUB_AddLetterProgress m_NewCP09, 2, True, "", True, m_strPA26, "1605", m_strPA75
      End If
      'end 2016/6/16
      
   'Added by Morgan 2014/7/2
   Else
      '新增信函進度
      'Modified by Morgan 2014/7/22 +傳FC代理人(pa75)
      PUB_AddLetterProgress m_NewCP09, 0, True, "", True, m_strPA26, "1605", m_strPA75
   'end 2014/7/2
   End If
   
   'Add by Sindy 2017/12/29
   If m_strIR01 <> "" Then
      'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", m_NewCP09, "")
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm040324", IIf(Pub_StrUserSt03 = "F22", m_NewCP09, "")
   End If
   '2017/12/29 END
   
   'Added by Morgan 2020/4/10
   'FMP有期限之案件EMAIL通知
   'Modified by Morgan 2022/10/17
   'If Left(stCP12, 1) = "F" Then
   'Modified by Lydia 2023/08/29 寰華案應比照FCP案的設定，若該案已閉卷，後續相關來函皆不須報告客戶，故無需再收到email通知。(Phoebe+Anny)
   'If m_bolFMP Then
   'end 2022/10/17
   If (m_bolFMP = True And m_bolFMP2 = False) Or (m_bolFMP2 = True And m_strPA57 <> "Y") Then
      'Modified by Morgan 2020/9/15 +寰華案,改通知智權人員
      'If Left(Pub_StrUserSt03, 1) <> "F" Then
      '   PUB_FMPCaseInform m_NewCP09, False
      'End If
      PUB_FMPCaseInform m_NewCP09, False, True, Left(Pub_StrUserSt03, 1) = "F"
      'end 2020/9/15
   End If
   'end 2020/4/10
   PUB_DualCaseInform m_NewCP09 'Added by Morgan 2022/4/7
   
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrorHandler1:
   cnnConnection.RollbackTrans
   
ErrorHandler:

End Function

