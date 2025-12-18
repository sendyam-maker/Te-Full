VERSION 5.00
Begin VB.Form frm060318 
   BorderStyle     =   1  '單線固定
   Caption         =   "領證逾期通知函"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6765
   Begin VB.Frame Frame1 
      Caption         =   "設定地址條"
      Height          =   660
      Left            =   330
      TabIndex        =   24
      Top             =   3000
      Width           =   4365
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   25
         Top             =   240
         Width           =   3390
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   26
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   4
      Left            =   1260
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1890
      Width           =   1275
   End
   Begin VB.ComboBox cbo 
      Height          =   300
      Index           =   1
      Left            =   1260
      Style           =   2  '單純下拉式
      TabIndex        =   8
      Top             =   1200
      Width           =   5325
   End
   Begin VB.TextBox text1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   0
      Left            =   1260
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "FCP"
      Top             =   570
      Width           =   495
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   1
      Left            =   1740
      MaxLength       =   6
      TabIndex        =   1
      Top             =   570
      Width           =   855
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   2
      Left            =   2580
      MaxLength       =   1
      TabIndex        =   2
      Top             =   570
      Width           =   255
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   3
      Left            =   2820
      MaxLength       =   2
      TabIndex        =   3
      Top             =   570
      Width           =   375
   End
   Begin VB.ComboBox cbo 
      Height          =   300
      Index           =   0
      Left            =   1260
      Style           =   2  '單純下拉式
      TabIndex        =   7
      Top             =   870
      Width           =   5325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   4560
      TabIndex        =   5
      Top             =   30
      Width           =   972
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   5565
      TabIndex        =   6
      Top             =   30
      Width           =   972
   End
   Begin VB.Label Label1 
      Height          =   300
      Index           =   5
      Left            =   4770
      TabIndex        =   23
      Top             =   2610
      Width           =   1305
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "下次法定期限："
      Height          =   180
      Index           =   9
      Left            =   3510
      TabIndex        =   22
      Top             =   2610
      Width           =   1260
   End
   Begin VB.Label Label1 
      Height          =   300
      Index           =   4
      Left            =   1620
      TabIndex        =   21
      Top             =   2610
      Width           =   1305
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "下次本所期限："
      Height          =   180
      Index           =   8
      Left            =   330
      TabIndex        =   20
      Top             =   2610
      Width           =   1260
   End
   Begin VB.Label Label1 
      Height          =   300
      Index           =   3
      Left            =   4770
      TabIndex        =   19
      Top             =   2220
      Width           =   1305
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "原法定期限："
      Height          =   180
      Index           =   7
      Left            =   3510
      TabIndex        =   18
      Top             =   2220
      Width           =   1080
   End
   Begin VB.Label Label1 
      Height          =   300
      Index           =   2
      Left            =   1620
      TabIndex        =   17
      Top             =   2220
      Width           =   1305
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "原本所期限："
      Height          =   180
      Index           =   6
      Left            =   330
      TabIndex        =   16
      Top             =   2220
      Width           =   1080
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "催辦日期："
      Height          =   180
      Index           =   5
      Left            =   330
      TabIndex        =   15
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label Label1 
      Height          =   300
      Index           =   1
      Left            =   1260
      TabIndex        =   14
      Top             =   1560
      Width           =   5325
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "此案已閉卷"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   3510
      TabIndex        =   13
      Top             =   570
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Index           =   4
      Left            =   330
      TabIndex        =   12
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "申請人１："
      Height          =   180
      Index           =   3
      Left            =   330
      TabIndex        =   11
      Top             =   1230
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱："
      Height          =   180
      Index           =   2
      Left            =   330
      TabIndex        =   10
      Top             =   900
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   1
      Left            =   330
      TabIndex        =   9
      Top             =   600
      Width           =   900
   End
End
Attribute VB_Name = "frm060318"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Dim m_CF07 As String
Dim m_CF08 As String
'edit by nickc 2007/02/02
'Dim m_NP(1 To T_NP) As String
'Dim m_NP_Temp(1 To T_NP) As String
Dim m_NP() As String
Dim m_NP_Temp() As String

Dim ET01 As String '定稿別
Dim ET02 As String '總收文號(或本所案號&案件性質)
Dim ET03 As String '處理狀況
'Add By Cheng 2003/01/01
Dim m_strOfficalFee As String '規費
Dim m_strServiceFee As String '服務費
Dim m_strPoints As String '點數
Dim m_PA08 As String '專利種類
Dim m_PA09 As String '申請國家
'Add By Sindy 2016/5/11
Dim m_bolEmail As Boolean, m_bolPlusPaper As Boolean, m_iCopy As Integer
Dim m_bolDNEmail As Boolean, m_bolDNPlusPaper As Boolean
'2016/5/11 END


Private Sub cmdExit_Click()
   Me.Enabled = False
   
   'Move to Unload by Morgan 2004/10/26
'    'Add By Cheng 2003/01/29
'    '列印地址條
'    PUB_PrintAddressList strUserNum, Me.Combo1.Text
'    '刪除地址條列表資料
'    PUB_DeleteAddressList strUserNum
'    '初始化序號
'    pub_AddressListSN = 0
'    'Add By Cheng 2003/02/05
'    '若印表機變動, 則更新列印設定
'    If Me.Combo1.Text <> Me.Combo1.Tag Then
'        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
'    End If
   '2004/10/26 end
   
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim iLang As Integer 'Added by Morgan 2013/11/27
'Add By Sindy 2016/5/11
Dim strNewCP09 As String, strCP10 As String
Dim strFileName As String, strFullFileName As String
Dim oFileSys As New FileSystemObject
Dim oFile As File
Dim bol1921 As Boolean 'Add by Amy 2018/11/27
Dim tmpErr As String 'Added by Lydia 2022/10/31

   If TxtValidate = False Then Exit Sub
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   'Add by Amy 2018/11/27 增加D類「通知領證逾期1921] ,自動上發文日
   If PUB_AddCaseProgressD("1921", m_NP(2), m_NP(3), m_NP(4), m_NP(5), "", "", "", m_NP(1), , strNewCP09) = False Then
        MsgBox "新增通知領證逾期進度檔失敗！作業中斷！", vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'Modify By Cheng 2003/02/17
    '取消更新資料
'   '更新資料庫
'   If SaveData Then
      '列印定稿
      ET01 = "12"
      ET02 = m_NP(2) & m_NP(3) & m_NP(4) & m_NP(5) & "&601"
      ET03 = "00"
      
      'Added by Morgan 2013/11/27
      iLang = PUB_GetLanguage(m_NP(2), m_NP(3), m_NP(4), m_NP(5), "601", "1")
      If iLang = 3 Then ET03 = "02" '日文
      'end 2013/11/27
      
      StartLetter ET01, ET02, ET03
      
      'Modify by Sindy 2016/5/11 判斷是否產生電子檔
      'NowPrint ET02, ET01, ET03, False, strUserNum, 0
      m_bolEmail = PUB_GetEMailFlag(m_NP(2) & m_NP(3) & m_NP(4) & m_NP(5), True, , m_bolPlusPaper)
      If m_bolEmail = False Then
         m_bolDNEmail = PUB_GetEMailFlag(m_NP(2) & m_NP(3) & m_NP(4) & m_NP(5), True, , m_bolDNPlusPaper, , True)
      Else
         m_bolDNEmail = m_bolEmail
         m_bolDNPlusPaper = m_bolPlusPaper
      End If
      '判斷是否EMail同時寄紙本
      If m_bolPlusPaper Then
         m_iCopy = 0
      Else
         m_iCopy = 1
      End If
      If m_bolEmail Then
         NowPrint ET02, ET01, ET03, False, strUserNum, 0, , , , m_iCopy, , True, True
         MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_NP(2)) & " ]！"
      Else
         NowPrint ET02, ET01, ET03, False, strUserNum, 0
      End If
      '2016/5/11 END
      
      'Modify By Sindy 2016/5/11 定稿轉PDF存卷宗區
      'Modify by Amy 2018/11/27 定稿轉PDF存卷宗區,放於「通知領證逾期1921] 那道
      'strSql = "select np01 From NextProgress Where NP02='" & m_NP(2) & "' And NP03='" & m_NP(3) & "' And NP04='" & m_NP(4) & "' And NP05='" & m_NP(5) & "' And NP07 ='" & 領證及繳年費 & "' And (NP06 IS NULL or NP06='N') order by NP09 desc"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         RsTemp.MoveFirst
'         strNewCP09 = RsTemp.Fields("np01")
         strSql = "select cp10 From caseProgress Where CP09='" & strNewCP09 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strCP10 = RsTemp.Fields("cp10")
         End If
         strFileName = m_NP(2) & m_NP(3) & IIf(m_NP(5) <> "00", "-" & m_NP(4) & "-" & m_NP(5), IIf(m_NP(4) <> "0", "-" & m_NP(4), "")) & "." & strCP10 & ".CUS.PDF"
         PUB_DelFtpFile2 strNewCP09, " and cpp02='" & strFileName & "'" '檔案改放 FTP,必須在DB資料刪除前執行
         strSql = "delete from CasePaperPDF where cpp01='" & strNewCP09 & "' and cpp02='" & strFileName & "'"
         cnnConnection.Execute strSql
         If PUB_PrintLetter(ET02, , , True, strFullFileName) = True Then
            Call PUB_ChkFileStatus(strFullFileName, False, tmpErr)  'Added by Lydia 2022/10/31 判斷檔案是否存在, 超過時間就繼續;
            Set oFile = oFileSys.GetFile(strFullFileName)
            If SaveAttFile_PDF(strNewCP09, strFullFileName, strFileName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False) = False Then
               MsgBox m_NP(2) & "-" & m_NP(3) & "-" & m_NP(4) & "-" & m_NP(5), strNewCP09, "定稿轉PDF失敗"
            Else
               Kill strFullFileName
            End If
         End If
'      End If
      '2016/5/11 END
      
      'Modify By Sindy 2016/5/11 產生電子檔時不印地址條
      If Not m_bolEmail Or m_bolPlusPaper Then
      '2016/5/11 END
      
'      'Add By Sindy 2015/9/21 日文定稿才要印地址條
'      If iLang = 3 Or Val(外專開窗信函啟用日) >= Val(strSrvDate(1)) Then
'      '2015/9/21 END
         'Add By Cheng 2003/01/29
         '新增地址條列表資料
         pub_AddressListSN = pub_AddressListSN + 1
         'Modify By Cheng 2003/02/07
         '加傳入綠皮貼紙的份數
   '     PUB_AddNewAddressList strUserNum, m_NP(2), m_NP(3), m_NP(4), m_NP(5), "" & pub_AddressListSN
         PUB_AddNewAddressList strUserNum, m_NP(2), m_NP(3), m_NP(4), m_NP(5), "" & pub_AddressListSN, "0"
      End If
      
      Me.Enabled = True
      ClearControl
      Me.Text1(1).Text = Empty
      Me.Text1(2).Text = Empty
      Me.Text1(3).Text = Empty
      Me.Text1(1).SetFocus
'   End If
   If Me.Enabled = False Then Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Function SaveData() As Boolean
Dim strUpdStatus As String '0:none 1:Begin 2:Commit
Dim ii As Integer

   On Error GoTo ErrorHandler
   
   SaveData = False
   strUpdStatus = "0"
   cnnConnection.BeginTrans
   strUpdStatus = "1"
   For ii = LBound(m_NP()) To UBound(m_NP())
      m_NP_Temp(ii) = m_NP(ii)
   Next ii
   '若原下一程序檔的"是否序辦"欄為NULL
   If m_NP(6) = "" Then
      m_NP_Temp(6) = "N"
      '更新下一程序檔
      If PUB_UpdateNextProgress(m_NP_Temp(), m_NP(1), m_NP(7), m_NP(22)) = False Then
         GoTo ErrorHandler
      End If
   End If
   For ii = LBound(m_NP()) To UBound(m_NP())
      m_NP_Temp(ii) = m_NP(ii)
   Next ii
    'Add By Cheng 2003/01/08
    '刪除重覆產生的下一程序檔
    strSql = "Delete From NextProgress Where NP02='" & m_NP_Temp(2) & "' And NP03='" & m_NP_Temp(3) & "' And NP04='" & m_NP_Temp(4) & "' And NP05='" & m_NP_Temp(5) & "' And NP07 ='" & 領證及繳年費 & "' And NP06 IS NULL And NP09 > " & strSrvDate(1)
    cnnConnection.Execute strSql
   '取得新的序號
   'edit by nickc 2007/02/02 不用 dll 了
   'm_NP_Temp(22) = objPublicData.GetNextProgressNo
   m_NP_Temp(22) = GetNextProgressNo
   '本所期限
   If Me.Label1(4).Caption <> "" Then
      m_NP_Temp(8) = Replace(Me.Label1(4).Caption, "/", "") + 19110000
   End If
   '法定期限
   If Me.Label1(5).Caption <> "" Then
      m_NP_Temp(9) = Replace(Me.Label1(5).Caption, "/", "") + 19110000
   End If
   '是否續辦
    'Modify By Cheng 2002/12/24
'   m_NP(6) = ""
   m_NP_Temp(6) = ""
   '解除期限日期
    'Modify By Cheng 2002/12/24
'   m_NP(11) = ""
   m_NP_Temp(11) = ""
   '解除期限原因
    'Modify By Cheng 2002/12/24
'   m_NP(12) = ""
   m_NP_Temp(12) = ""
   '新增下一程序檔
   If PUB_AddNewNextProgress(m_NP_Temp()) = False Then
      GoTo ErrorHandler
   End If
   
   cnnConnection.CommitTrans
   strUpdStatus = "2"
   
   SaveData = True
   Exit Function
ErrorHandler:
   If strUpdStatus = "1" Then
      cnnConnection.RollbackTrans
      If Err.Number <> 0 Then MsgBox "(" & Err.Number & ")" & Err.Description, vbExclamation + vbOKOnly, "更新動作失敗"
   End If
      
End Function

Private Sub StartLetter(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)
 Dim strTxt(1 To 10) As String, i As Integer, j As Integer, strTmp As String
Dim strFee As String '領證費
   
   EndLetter ET01, ET02, ET03, strUserNum
   i = 0
   'Modify By Sindy 2021/4/27
   If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
      '約定期限
      If Me.Label1(2).Tag <> "" Then
         i = i + 1
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','約定期限'," & CNULL(Replace(Me.Label1(2).Tag, "/", "") + 19110000) & ")"
      End If
   Else
   '2021/4/27 END
      '本所期限
      If Me.Label1(2).Caption <> "" Then
         i = i + 1
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','本所期限'," & CNULL(Replace(Me.Label1(2).Caption, "/", "") + 19110000) & ")"
      End If
   End If
   
   i = i + 1
'Modified by Morgan 2013/1/18
'   '其他日期
'   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','其他日期'," & CNULL(Replace(Me.text1(4).Text, "/", "") + 19110000) & ")"
   '原法定期限
   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','法定期限'," & CNULL(Replace(Me.Label1(3).Caption, "/", "") + 19110000) & ")"
'end 2013/1/18

   '下次領證日
   If Me.Label1(4).Caption <> "" Then
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','下次領證日'," & CNULL(Replace(Me.Label1(4).Caption, "/", "") + 19110000) & ")"
   End If
    'Add By Cheng 2003/01/01
    '取得領證及繳年費相關費用
    GetPatentYearFee m_PA09, m_PA08, "Y00000000", 領證及繳年費, 1, 1, True
   '領證費
   i = i + 1
    'Modify By Cheng 2003/01/01
'   strFee = Val(m_CF07) + (m_CF08 * 1.5)
'   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','領證費'," & CNULL(strFee) & ")"
   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','領證費'," & CNULL(Val(m_strServiceFee) + Val(m_strOfficalFee)) & ")"
   '費用
   i = i + 1
    'Modify By Cheng 2003/01/01
'   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','費用'," & CNULL(Round((strFee / PUB_GetUSXRate) + 0.00000000001, 2)) & ")"
   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','費用'," & CNULL(Round(((Val(m_strServiceFee) + Val(m_strOfficalFee)) / PUB_GetUSXRate) + 0.00000000001, 2)) & ")"
   '服務費
   i = i + 1
    'Modify By Cheng 2003/01/01
'   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','服務費'," & CNULL(Val(m_CF07)) & ")"
   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','服務費'," & CNULL(Val(m_strServiceFee)) & ")"
   '規費
   i = i + 1
'   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','規費'," & CNULL(Val(m_CF08) * 1.5) & ")"
   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','規費'," & CNULL(Val(m_strOfficalFee)) & ")"
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(i, strTxt) Then
   If Not ClsLawExecSQL(i, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If

End Sub

'Private Sub Form_Initialize()
''add by nickc 2007/02/02
'ReDim m_NP(1 To TF_NP) As String
'ReDim m_NP_Temp(1 To TF_NP) As String
'End Sub
'
Private Sub Form_Load()
   'Add by Amy 2018/11/28 從Form_Initialize搬下來否則會error(mdiMain設ShowFrm060318)
   'add by nickc 2007/02/02
    ReDim m_NP(1 To TF_NP) As String
    ReDim m_NP_Temp(1 To TF_NP) As String
    'end 2018/11/28
   MoveFormToCenter Me

'Modify by Morgan 2011/3/15 改共用且不要排除預設印表機
   PUB_SetPrinter Me.Name, Combo1
'end 2011/3/1

End Sub

Private Sub Form_Unload(Cancel As Integer)

   'Copy from cmdExit_Click by Morgan 2004/10/26
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
   '2004/10/26 end
    Set frm060318 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Me.Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
   Case 3
      If Me.Text1(0).Text <> "" And Me.Text1(1).Text <> "" Then
         ClearControl
         ReadData
      Else
         MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
         Me.Text1(1).SetFocus
         TextInverse Me.Text1(1)
         Exit Sub
      End If
   End Select
End Sub

Private Function ReadData() As Boolean
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim strPA09 As String
   
   ReadData = False
   'Add By Cheng 2003/01/01
   m_PA08 = ""
   m_PA09 = ""
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 清除查詢印表記錄檔欄位
   pub_QL05 = pub_QL05 & ";" & Label17(1) & Text1(0) & "-" & Text1(1) & "-" & Text1(2) & "-" & Text1(3) 'Add By Sindy 2010/12/7
   
   m_CF07 = Empty
   m_CF08 = Empty
   StrSQLa = "SELECT * FROM PATENT,CUSTOMER,FAGENT WHERE " & ChgPatent(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text) & " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
        'Add By Cheng 2003/01/01
        m_PA08 = "" & rsA("PA08").Value
        m_PA09 = "" & rsA("PA09").Value
        
      strPA09 = "" & rsA.Fields("PA09").Value
      If "" & rsA.Fields("PA57") = "Y" Then Me.Label1(0).Visible = True
      '專利名稱(中-->英-->日)
      Me.cbo(0).AddItem "中：" & rsA.Fields("PA05").Value
      Me.cbo(0).AddItem "英：" & rsA.Fields("PA06").Value
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Me.cbo(0).AddItem "外：" & rsA.Fields("PA07").Value
      Me.cbo(0).ListIndex = 0
      '申請人1(中-->英-->日)
      Me.cbo(1).AddItem "中：" & rsA.Fields("CU04").Value
      Me.cbo(1).AddItem "英：" & Trim("" & rsA.Fields("CU05").Value & " " & rsA.Fields("CU88").Value & " " & rsA.Fields("CU89").Value & " " & rsA.Fields("CU90").Value & " ")
      Me.cbo(1).AddItem "日：" & rsA.Fields("CU06").Value
      Me.cbo(1).ListIndex = 0
      '代理人(英-->中-->日)
      Me.Label1(1).Caption = IIf(Not IsNull(rsA("FA05").Value), rsA("FA05").Value & " " & rsA("FA63").Value & " " & rsA("FA64").Value & " " & rsA("FA65").Value, IIf(Not IsNull(rsA("FA04")), rsA("FA04"), rsA("FA06")))
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      StrSQLa = "Select * From NextProgress Where " & ChgNextProgress(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text) & " AND NP07='601' AND (NP09 IS NOT NULL AND NP09<" & ServerDate & " ) ORDER BY NP09 DESC "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      '若下一程序檔無符合的資料
      If rsA.RecordCount <= 0 Then
         InsertQueryLog (0) 'Add By Sindy 2010/12/7
         MsgBox "此案件無領證逾期情形!!!", vbExclamation + vbOKOnly
         Me.Text1(1).SetFocus
         ReadData = False
      Else
         '讀取下一程序檔資料
         PUB_ReadNextProgressData m_NP(), rsA.Fields("NP01"), rsA.Fields("NP07"), rsA.Fields("NP22")
         '若不為續辦
         If "" & rsA.Fields("NP06") <> "Y" Then
            '不必再檢查案件進度檔
            '原本所期限
            Me.Label1(2).Caption = IIf(IsNumeric(rsA.Fields("NP08")), ChangeTStringToTDateString(rsA.Fields("NP08") - 19110000), "")
            'Add By Sindy 2021/4/27 原約定期限
            Me.Label1(2).Tag = IIf(IsNumeric(rsA.Fields("NP23")), ChangeTStringToTDateString(rsA.Fields("NP23") - 19110000), "")
            '2021/4/27 END
            '原法定期限
            Me.Label1(3).Caption = IIf(IsNumeric(rsA.Fields("NP09")), ChangeTStringToTDateString(rsA.Fields("NP09") - 19110000), "")
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            StrSQLa = "Select * From CaseFee Where CF01='" & Me.Text1(0).Text & "' AND CF02='" & strPA09 & "' AND CF03='601' "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               m_CF07 = rsA.Fields("CF07")
               m_CF08 = rsA.Fields("CF08")
               '下期期限(天)
               If Not IsNull(rsA.Fields("CF12")) Then
                  '下次法定期限
                  If Me.Label1(3).Caption <> "" Then
                    'Modify By Cheng 2002/12/23
'                     Me.Label1(5).Caption = "" & ChangeTStringToTDateString(Replace(DateAdd("D", rsA.Fields("CF12").Value, ChangeTStringToWDateString(Replace(Me.Label1(3).Caption, "/", ""))), "/", "") - 19110000)
                     Me.Label1(5).Caption = ChangeTStringToTDateString((DBDATE(DateAdd("D", rsA.Fields("CF12").Value, ChangeWStringToWDateString(ChangeTDateStringToTString(Me.Label1(3).Caption) + 19110000))) - 19110000))
                  End If
                  '下次本所期限=下次本所期限-2天
                  If Me.Label1(5).Caption <> "" Then
                    'Modify By Cheng 2002/12/23
'                     Me.Label1(4).Caption = "" & ChangeTStringToTDateString(Replace(DateAdd("D", -2, ChangeTStringToWDateString(Replace(Me.Label1(5).Caption, "/", ""))), "/", "") - 19110000)
                     Me.Label1(4).Caption = ChangeTStringToTDateString((DBDATE(DateAdd("D", -2, ChangeWStringToWDateString(ChangeTDateStringToTString(Me.Label1(5).Caption) + 19110000))) - 19110000))
                  End If
               '下期期限(月)
               ElseIf Not IsNull(rsA.Fields("CF28")) Then
                  '下次法定期限
                  If Me.Label1(3).Caption <> "" Then
'                     Me.Label1(5).Caption = "" & ChangeTStringToTDateString(Replace(DateAdd("M", rsA.Fields("CF28").Value, ChangeTStringToWDateString(Replace(Me.Label1(3).Caption, "/", ""))), "/", "") - 19110000)
                     Me.Label1(5).Caption = ChangeTStringToTDateString((DBDATE(DateAdd("M", rsA.Fields("CF28").Value, ChangeWStringToWDateString(ChangeTDateStringToTString(Me.Label1(3).Caption) + 19110000))) - 19110000))
                  End If
                  '下次本所期限=下次本所期限-2天
                  If Me.Label1(5).Caption <> "" Then
                    'Modify By Cheng 2002/12/23
'                     Me.Label1(4).Caption = "" & ChangeTStringToTDateString(Replace(DateAdd("D", -2, ChangeTStringToWDateString(Replace(Me.Label1(5).Caption, "/", ""))), "/", "") - 19110000)
                     Me.Label1(4).Caption = ChangeTStringToTDateString((DBDATE(DateAdd("D", -2, ChangeWStringToWDateString(ChangeTDateStringToTString(Me.Label1(5).Caption) + 19110000))) - 19110000))
                  End If
               '92.2.22 ADD BY SONIA 案件國家收費表未定義時抓6個月
               Else
                  '下次法定期限
                  If Me.Label1(3).Caption <> "" Then
                     Me.Label1(5).Caption = ChangeTStringToTDateString((DBDATE(DateAdd("M", 6, ChangeWStringToWDateString(ChangeTDateStringToTString(Me.Label1(3).Caption) + 19110000))) - 19110000))
                  End If
                  '下次本所期限=下次本所期限-2天
                  If Me.Label1(5).Caption <> "" Then
                     Me.Label1(4).Caption = ChangeTStringToTDateString((DBDATE(DateAdd("D", -2, ChangeWStringToWDateString(ChangeTDateStringToTString(Me.Label1(5).Caption) + 19110000))) - 19110000))
                  End If
               '92.2.22 END
               End If
            End If
            InsertQueryLog (rsA.RecordCount) 'Add By Sindy 2010/12/7
            ReadData = True
         '若為續辦
         Else
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text) & " AND CP10='601' AND CP27 IS NULL AND (CP07 IS NOT NULL AND CP07 <" & ServerDate & ") ORDER BY CP07 DESC "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            '若案件進度檔無符合資料
            If rsA.RecordCount <= 0 Then
               InsertQueryLog (0) 'Add By Sindy 2010/12/7
               MsgBox "此案件無領證逾期情形!!!", vbExclamation + vbOKOnly
               Me.Text1(1).SetFocus
               ReadData = False
            Else
               '原本所期限
               Me.Label1(2).Caption = IIf(IsNumeric(rsA.Fields("NP08")), ChangeTStringToTDateString(rsA.Fields("NP08") - 19110000), "")
               'Add By Sindy 2021/4/27 原約定期限
               Me.Label1(2).Tag = IIf(IsNumeric(rsA.Fields("NP23")), ChangeTStringToTDateString(rsA.Fields("NP23") - 19110000), "")
               '2021/4/27 END
               '原法定期限
               Me.Label1(3).Caption = IIf(IsNumeric(rsA.Fields("NP09")), ChangeTStringToTDateString(rsA.Fields("NP09") - 19110000), "")
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               StrSQLa = "Select * From CaseFee Where CF01='" & Me.Text1(0).Text & "' AND CF02='" & strPA09 & "' AND CF03='601' "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  m_CF07 = rsA.Fields("CF07")
                  m_CF08 = rsA.Fields("CF08")
                  '下期期限(天)
                  If Not IsNull(rsA.Fields("CF12")) Then
                     If Me.Label1(2).Caption <> "" Then
                        'Modify By Cheng 2002/12/23
'                        Me.Label1(4).Caption = "" & ChangeTStringToTDateString(Replace(DateAdd("D", rsA.Fields("CF12").Value, ChangeTStringToWDateString(Replace(Me.Label1(2).Caption, "/", ""))), "/", "") - 19110000)
                        Me.Label1(4).Caption = ChangeTStringToTDateString((DBDATE(DateAdd("D", rsA.Fields("CF12").Value, ChangeWStringToWDateString(ChangeTDateStringToTString(Me.Label1(2).Caption) + 19110000))) - 19110000))
                     End If
                     If Me.Label1(3).Caption <> "" Then
                        'Modify By Cheng 2002/12/23
'                        Me.Label1(5).Caption = "" & ChangeTStringToTDateString(Replace(DateAdd("D", rsA.Fields("CF12").Value, ChangeTStringToWDateString(Replace(Me.Label1(3).Caption, "/", ""))), "/", "") - 19110000)
                         Me.Label1(5).Caption = ChangeTStringToTDateString((DBDATE(DateAdd("D", rsA.Fields("CF12").Value, ChangeWStringToWDateString(ChangeTDateStringToTString(Me.Label1(3).Caption) + 19110000))) - 19110000))
                     End If
                  '下期期限(月)
                  ElseIf Not IsNull(rsA.Fields("CF28")) Then
                     If Me.Label1(2).Caption <> "" Then
                        'Modify By Cheng 2002/12/23
'                        Me.Label1(4).Caption = "" & ChangeTStringToTDateString(Replace(DateAdd("M", rsA.Fields("CF28").Value, ChangeTStringToWDateString(Replace(Me.Label1(2).Caption, "/", ""))), "/", "") - 19110000)
                        Me.Label1(4).Caption = ChangeTStringToTDateString((DBDATE(DateAdd("M", rsA.Fields("CF28").Value, ChangeWStringToWDateString(ChangeTDateStringToTString(Me.Label1(2).Caption) + 19110000))) - 19110000))
                     End If
                     If Me.Label1(3).Caption <> "" Then
                        'Modify By Cheng 2002/12/23
'                        Me.Label1(5).Caption = "" & ChangeTStringToTDateString(Replace(DateAdd("M", rsA.Fields("CF28").Value, ChangeTStringToWDateString(Replace(Me.Label1(3).Caption, "/", ""))), "/", "") - 19110000)
                         Me.Label1(5).Caption = ChangeTStringToTDateString((DBDATE(DateAdd("M", rsA.Fields("CF28").Value, ChangeWStringToWDateString(ChangeTDateStringToTString(Me.Label1(3).Caption) + 19110000))) - 19110000))
                     End If
                  End If
               End If
               InsertQueryLog (rsA.RecordCount) 'Add By Sindy 2010/12/7
               ReadData = True
            End If
         End If
      End If
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/7
      MsgBox "基本檔無此案號資料!!!", vbExclamation + vbOKOnly
      Me.Text1(1).SetFocus
      Text1_GotFocus 1
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
End Function

Private Sub ClearControl()
   Me.Label1(0).Visible = False
   Me.cbo(0).Clear
   Me.cbo(1).Clear
   Me.Label1(1).Caption = Empty
   Me.Text1(4).Text = Empty
   Me.Label1(2).Caption = Empty
   Me.Label1(2).Tag = Empty 'Add By Sindy 2021/4/27
   Me.Label1(3).Caption = Empty
   Me.Label1(4).Caption = Empty
   Me.Label1(5).Caption = Empty
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 4 '催辦日期
      If Me.Text1(Index).Text <> "" Then
         If ChkDate(Me.Text1(Index).Text) = False Then
            Cancel = True
         End If
      End If
   End Select
   If Cancel = True Then Text1_GotFocus Index
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.Text1(1).Text = "" Then
   MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
   Me.Text1(1).SetFocus
   TextInverse Me.Text1(1)
   Exit Function
End If
If Me.Text1(4).Text = "" Then
   MsgBox "請輸入催辦日期!!!", vbExclamation + vbOKOnly
   Me.Text1(4).SetFocus
   TextInverse Me.Text1(4)
   Exit Function
End If
For Each objTxt In Text1
   If objTxt.Enabled = True Then
      Cancel = False
      Text1_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next
If ReadData = False Then
   ClearControl
   Me.Text1(1).SetFocus
   Text1_GotFocus 1
   Exit Function
End If
'Add by Amy 2019/01/23 閉卷不可操作
If Me.Label1(0).Visible = True Then
    MsgBox "此案號已閉卷不可操作!!!", vbExclamation + vbOKOnly
    Me.Text1(1).SetFocus
    TextInverse Me.Text1(1)
   Exit Function
End If

TxtValidate = True
End Function

'Add By Cheng 2002/12/31
'計算相關費用
Private Sub GetPatentYearFee( _
    strYF01 As String, strYF02 As String, strYF03 As String, _
    strYF04 As String, strYF05From As String, strYF05To As String, blnDouble As Boolean)
'strYF01  申請國家
'strYF02  專利種類
'strYF03  代理人
'strYF04  案件性質
'strYF05From  起始年度
'strYF05To  終止年度
'blnDouble  規費是否雙倍

Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

    m_strOfficalFee = 0
    m_strServiceFee = 0
    m_strPoints = 0
    '若案件性質為領證及繳年費, 則先取得領證相關費用
    If strYF04 = 領證及繳年費 Then
        StrSQLa = "Select * From PatentYearFee Where YF01='" & strYF01 & "' AND YF02='" & strYF02 & "' AND YF03='" & strYF03 & "' AND YF04='" & strYF04 & "' "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            m_strOfficalFee = Val(m_strOfficalFee) + Val(rsA.Fields("YF07").Value)
            '93.7.14 CANCEL BY SONIA 應為年費規費雙倍
            'If blnDouble = True Then m_strOfficalFee = 2 * Val(m_strOfficalFee)
            '93.7.14 END
            m_strServiceFee = Val(m_strServiceFee) + Val(rsA.Fields("YF06").Value)
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End If
    '取得案件性質為年費的相關費用
    StrSQLa = "Select * From PatentYearFee Where YF01='" & strYF01 & "' AND YF02='" & strYF02 & "' AND YF03='" & strYF03 & "' AND YF04='" & 年費 & "' AND YF05>=" & Val(strYF05From) & " AND YF05<=" & Val(strYF05To) & " Order By YF05 "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    While Not rsA.EOF
        '93.7.14 modify BY SONIA 應為年費規費雙倍
        'm_strOfficalFee = Val(m_strOfficalFee) + Val(rsA.Fields("YF07").Value)
        If blnDouble = True Then
           m_strOfficalFee = Val(m_strOfficalFee) + 2 * Val(rsA.Fields("YF07").Value)
        Else
           m_strOfficalFee = Val(m_strOfficalFee) + Val(rsA.Fields("YF07").Value)
        End If
        '93.7.14 END
        m_strServiceFee = Val(m_strServiceFee) + Val(rsA.Fields("YF06").Value)
        rsA.MoveNext
    Wend
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    m_strOfficalFee = m_strOfficalFee + 3000 'Added by Morgan 2013/1/18 申請復權(NTD3000)
    m_strPoints = Val(m_strServiceFee) / 1000
End Sub
