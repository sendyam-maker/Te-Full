VERSION 5.00
Begin VB.Form frm060319 
   BorderStyle     =   1  '單線固定
   Caption         =   "年費逾期／檢還證據樣品請款通知函"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   7215
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   8
      Left            =   1260
      MaxLength       =   1
      TabIndex        =   8
      Top             =   3480
      Width           =   345
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   4500
      Style           =   2  '單純下拉式
      TabIndex        =   28
      Top             =   2790
      Width           =   2505
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   4500
      TabIndex        =   9
      Top             =   2460
      Width           =   2505
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   7
      Left            =   1260
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1170
      Width           =   345
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   6
      Left            =   4500
      MaxLength       =   7
      TabIndex        =   7
      Top             =   3150
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   5
      Left            =   1260
      MaxLength       =   7
      TabIndex        =   6
      Top             =   3090
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   4
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   5
      Top             =   2790
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.ComboBox cbo 
      Height          =   300
      Index           =   1
      Left            =   1260
      Style           =   2  '單純下拉式
      TabIndex        =   13
      Top             =   1800
      Width           =   5745
   End
   Begin VB.TextBox text1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   0
      Left            =   1260
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "FCP"
      Top             =   870
      Width           =   495
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   1
      Left            =   1740
      MaxLength       =   6
      TabIndex        =   1
      Top             =   870
      Width           =   855
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   2
      Left            =   2580
      MaxLength       =   1
      TabIndex        =   2
      Top             =   870
      Width           =   255
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   3
      Left            =   2820
      MaxLength       =   2
      TabIndex        =   3
      Top             =   870
      Width           =   375
   End
   Begin VB.ComboBox cbo 
      Height          =   300
      Index           =   0
      Left            =   1260
      Style           =   2  '單純下拉式
      TabIndex        =   12
      Top             =   1470
      Width           =   5745
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5040
      TabIndex        =   10
      Top             =   30
      Width           =   972
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6045
      TabIndex        =   11
      Top             =   30
      Width           =   972
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "新增請款單：          (Y:新增請款單)"
      Height          =   180
      Index           =   9
      Left            =   150
      TabIndex        =   30
      Top             =   3510
      Width           =   2715
   End
   Begin VB.Label Label2 
      Caption         =   "地址條印表機："
      Height          =   180
      Index           =   1
      Left            =   3240
      TabIndex        =   29
      Top             =   2805
      Width           =   1260
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "請款單印表機："
      Height          =   180
      Index           =   11
      Left            =   3240
      TabIndex        =   27
      Top             =   2490
      Width           =   1260
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "通知函類別：          (1:年費逾期   2:檢還證據樣品請款)"
      Height          =   180
      Index           =   10
      Left            =   150
      TabIndex        =   26
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label Label1 
      Height          =   300
      Index           =   2
      Left            =   1260
      TabIndex        =   25
      Top             =   2490
      Width           =   1395
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "最後 　　期限："
      Height          =   180
      Index           =   8
      Left            =   3240
      TabIndex        =   24
      Top             =   3180
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "原年費期限："
      Height          =   180
      Index           =   7
      Left            =   150
      TabIndex        =   23
      Top             =   3150
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "未繳年度："
      Height          =   180
      Index           =   6
      Left            =   330
      TabIndex        =   22
      Top             =   2820
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "專利種類："
      Height          =   180
      Index           =   5
      Left            =   330
      TabIndex        =   21
      Top             =   2490
      Width           =   900
   End
   Begin VB.Label Label1 
      Height          =   300
      Index           =   1
      Left            =   1260
      TabIndex        =   20
      Top             =   2160
      Width           =   5745
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
      Height          =   252
      Index           =   0
      Left            =   3684
      TabIndex        =   19
      Top             =   876
      Visible         =   0   'False
      Width           =   1968
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Index           =   4
      Left            =   330
      TabIndex        =   18
      Top             =   2160
      Width           =   720
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "申請人１："
      Height          =   180
      Index           =   3
      Left            =   330
      TabIndex        =   17
      Top             =   1830
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱："
      Height          =   180
      Index           =   2
      Left            =   330
      TabIndex        =   16
      Top             =   1500
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   1
      Left            =   330
      TabIndex        =   15
      Top             =   900
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "PS：列印通知函後，會同時列印請款單！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   0
      Left            =   330
      TabIndex        =   14
      Top             =   510
      Width           =   4320
   End
End
Attribute VB_Name = "frm060319"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
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
Dim m_strPA72 As String '已繳年度
Dim m_strPA75 As String 'FC代理人
Dim m_strPA26 As String '申請人1
Dim m_strPA27 As String '申請人2
Dim m_strPA28 As String '申請人3
Dim m_strPA29 As String '申請人4
Dim m_strPA30 As String '申請人5
Dim m_strUSXR02 As String '美金匯率
Dim m_strCF07 As String '費用(迄)
Dim m_strCF08 As String '規費
'edit by nickc 2007/02/02
'Dim m_CP(1 To T_CP) As String
Dim m_CP() As String

Dim m_strCP09 As String '總收文號

Dim m_strSerialNo As String '請款單號
Dim strSql As String
Dim strNo As String
Dim lngAmount As Long
Dim douAmount As Double
Dim strAmount As String
Dim intLength As Integer
Dim intCounter As Integer
Dim douUSDollar As Double
Dim strLanguage As String
Dim strMaxNo As String
Dim strDiscount As String
Private Const intDefault As Integer = 500
Private Const intTop As Integer = 1000
Dim strNewPage As String
Dim m_strPA14 As String '公告日
Dim m_bolEmail As Boolean, m_bolPlusPaper As Boolean, m_iCopy As Integer
'Added by Morgan 2014/6/3
Dim m_bolDNEmail As Boolean, m_bolDNPlusPaper As Boolean
Dim m_PrintRpt1 As Boolean, ff1 As Integer, m_strFileName1 As String 'Add By Sindy 2016/1/27

Private Sub cmdExit_Click()
   Me.Enabled = False
   Unload Me
End Sub

Private Sub cmdOK_Click()
Dim iLang As Integer
'Add By Sindy 2016/1/27
Dim strFileName As String, strFullFileName As String
Dim oFileSys As New FileSystemObject
Dim oFile As File
Dim strMsg As String
Dim strNewCP09 As String
'2016/1/27 END

   If TxtValidate = False Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   
   Me.Enabled = False
   m_bolEmail = False
   '更新資料庫
   If SaveData Then
      '列印定稿
      Select Case Me.Text1(7).Text
      Case "1"
         ET01 = "13"
         ET02 = m_strPA01 & m_strPA02 & m_strPA03 & m_strPA04 & "&1605"
         ET03 = "01"
         
         'Add by Morgan 2005/12/14 抓定稿語言
         'Modify by Morgan 2006/5/30
         'iLang = GetLetterLanguage(m_strPA01, m_strPA02, m_strPA03, m_strPA04)
         iLang = PUB_GetLanguage(m_strPA01, m_strPA02, m_strPA03, m_strPA04, "1605", "1")
         'end 2006/5/30
         
         'Modify by Morgan 2005/12/14 加日文
         If iLang = 3 Then
            ET03 = "03"
         End If
         
         StartLetter ET01, ET02, ET03
         
         'Modify by Morgan 2008/3/21 判斷是否產生電子檔
         'NowPrint ET02, ET01, ET03, False, strUserNum, 0
         m_bolEmail = PUB_GetEMailFlag(m_strPA01 & m_strPA02 & m_strPA03 & m_strPA04, True, , m_bolPlusPaper)
         'Added by Morgan 2014/6/3
         If m_bolEmail = False Then
            m_bolDNEmail = PUB_GetEMailFlag(m_strPA01 & m_strPA02 & m_strPA03 & m_strPA04, True, , m_bolDNPlusPaper, , True)
         Else
            m_bolDNEmail = m_bolEmail
            m_bolDNPlusPaper = m_bolPlusPaper
         End If
         'end 2014/6/3
         
         'Add by Morgan 2009/10/20 +判斷是否EMail同時寄紙本
         If m_bolPlusPaper Then
            m_iCopy = 0
         Else
            m_iCopy = 1
         End If
         'end 2009/10/20
         If m_bolEmail Then
            NowPrint ET02, ET01, ET03, False, strUserNum, 0, , , , m_iCopy, , True, True
            MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_strPA01) & " ]！"
         Else
            NowPrint ET02, ET01, ET03, False, strUserNum, 0
         End If
         'end 2008/3/20
         
      Case "2"
         ET01 = "13"
         ET02 = m_strPA01 & m_strPA02 & m_strPA03 & m_strPA04 & "&1904"
         ET03 = "02"
         StartLetter ET01, ET02, ET03
         NowPrint ET02, ET01, ET03, False, strUserNum, 0
      End Select
      
      'Modify By Sindy 2016/1/27 定稿轉PDF存卷宗區
      strNewCP09 = m_strCP09
      strFileName = m_strPA01 & m_strPA02 & IIf(m_strPA04 <> "00", "-" & m_strPA03 & "-" & m_strPA04, IIf(m_strPA03 <> "0", "-" & m_strPA03, "")) & "." & IIf(Me.Text1(7).Text = "1", "1605", "1904") & ".CUS.PDF"
      PUB_DelFtpFile2 strNewCP09, " and cpp02='" & strFileName & "'" '檔案改放 FTP,必須在DB資料刪除前執行
      strSql = "delete from CasePaperPDF where cpp01='" & strNewCP09 & "' and cpp02='" & strFileName & "'"
      cnnConnection.Execute strSql
      If PUB_PrintLetter(ET02, , , True, strFullFileName) = True Then
         Call PUB_ChkFileStatus(strFullFileName, False, strMsg)  'Added by Lydia 2022/10/31 判斷檔案是否存在, 超過時間就繼續;
         Set oFile = oFileSys.GetFile(strFullFileName)
         If SaveAttFile_PDF(strNewCP09, strFullFileName, strFileName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False) = False Then
            'Modified by Lydia 2022/10/31 +& ";" & strMsg
            Call ReadTxt1(m_strPA01 & "-" & m_strPA02 & "-" & m_strPA03 & "-" & m_strPA04, strNewCP09, "定稿轉PDF失敗" & ";" & strMsg)
         End If
         Kill strFullFileName
      End If
      '2016/1/27 END
      
      'Modify by Morgan 2008/3/20 產生電子檔時不印地址條
      If Not m_bolEmail Or m_bolPlusPaper Then
'         'Add By Sindy 2015/9/21 日文定稿才要印地址條
'         If iLang = 3 Or Val(外專開窗信函啟用日) >= Val(strSrvDate(1)) Then
'         '2015/9/21 END
            '新增地址條列表資料
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewAddressList strUserNum, m_strPA01, m_strPA02, m_strPA03, m_strPA04, "" & pub_AddressListSN, "0", IIf(Me.Text1(7).Text = "1", "605", "")
'         End If
      End If
      
      '92.7.15 ADD BY SONIA
      If Text1(8) = "Y" Then
      '92.7.15 END
         '6:列印新增的請款資料
         ProcessPrint
      End If
      Me.Enabled = True
      
      ClearControl
      Me.Text1(1).Text = Empty
      Me.Text1(2).Text = Empty
      Me.Text1(3).Text = Empty
      Me.Text1(8).Text = Empty
      Me.Text1(1).SetFocus
   End If
   If Me.Enabled = False Then Me.Enabled = True
   Screen.MousePointer = vbDefault
   
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
      m_strFileName1 = Me.Caption & Text1(0) & "資料檢核表.txt"
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

Private Function SaveData() As Boolean
Dim strUpdStatus As String '0:none 1:Begin 2:Commit
Dim ii As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strAutoNumber As String
Dim strAgentNo As String '代理人編號
Dim strPrintCust  As String '是否列印申請人
Dim dblUSRate As Double '美金匯率
Dim strA1K27 As String '列印對象
Dim strA1K28 As String '請款對象

   On Error GoTo ErrorHandler
   SaveData = False
   strUpdStatus = "0"
   cnnConnection.BeginTrans
   strUpdStatus = "1"
   m_strUSXR02 = PUB_GetUSXRate
   m_strCF07 = GetCFFieldData("CF07", "FCP", m_strPA09, IIf(Me.Text1(7).Text = "1", "1605", "1904"))
   m_strCF08 = GetCFFieldData("CF08", "FCP", m_strPA09, IIf(Me.Text1(7).Text = "1", "1605", "1904"))
   If Me.Text1(7).Text = "1" Then
      Erase m_CP
      ReDim m_CP(1 To TF_CP) As String
      m_CP(1) = m_strPA01
      m_CP(2) = m_strPA02
      m_CP(3) = m_strPA03
      m_CP(4) = m_strPA04
      m_CP(5) = strSrvDate(1)
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetAutoNumber("C", strAutoNumber, True, True) Then
      If ClsPDGetAutoNumber("C", strAutoNumber, True, True) Then
         'Modify By Sindy 2010/8/18 比對自動編號年度
         'm_CP(9) = "C" & Left(strSrvDate(1) - 19110000, 2) & strAutoNumber
         m_CP(9) = "C" & CompAutoNumberYear(Val(Mid(strSrvDate(1), 1, 4)) - 1911) & strAutoNumber
         m_strCP09 = m_CP(9)
      Else
         GoTo ErrorHandler
      End If
      m_CP(10) = "1605"
      m_CP(13) = PUB_GetFCPSalesNo(m_strPA01, m_strPA02, m_strPA03, m_strPA04)
      m_CP(12) = GetSalesArea(PUB_GetFCPSalesNo(m_strPA01, m_strPA02, m_strPA03, m_strPA04))
      m_CP(14) = strUserNum
      m_CP(27) = ServerDate
      m_CP(32) = "N"
      m_CP(64) = "年費逾期請款"
      m_CP(20) = PUB_GetCP20(m_CP(1), m_CP(10), m_CP(16), m_strPA26 & m_strPA27 & m_strPA28 & m_strPA29 & m_strPA30, m_strPA75, m_CP(1) & m_CP(2) & m_CP(3) & m_CP(4))
      If Not PUB_AddNewCaseProgress(m_CP) Then GoTo ErrorHandler
      
      PUB_DualCaseInform m_CP(9) 'Added by Morgan 2022/4/7
   End If
   
   '92.7.15 ADD BY SONIA
   If Text1(8) <> "Y" Then
      cnnConnection.CommitTrans
      strUpdStatus = "2"
      SaveData = True
      Exit Function
   End If
   '92.7.15 END
   
   '開始新增國外請款資料
   '1:先以"X"抓ACC1R0之國外請款單的自動編號, 並更新其流水號
   m_strSerialNo = AccAutoNo(MsgText(815), 5)
   AccSaveAutoNo MsgText(815), Right(m_strSerialNo, 5)
   '2:新增ACC1K0
'   strAgentNo = GetAgentNO '代理人編號
   strAgentNo = PUB_GetA1K03(m_strPA01, m_strPA02, m_strPA03, m_strPA04)
  ' dblUSRate = GetUSRate '美金匯率
   
    strA1K27 = PUB_GetA1K27(m_strPA01, m_strPA02, m_strPA03, m_strPA04, IIf(Me.Text1(7).Text = "1", "1605", "1904"))
    If strA1K27 = "" Then strA1K27 = strAgentNo
    strA1K28 = PUB_GetA1K28(m_strPA01, m_strPA02, m_strPA03, m_strPA04, IIf(Me.Text1(7).Text = "1", "1605", "1904"))
    If strA1K28 = "" Then strA1K28 = strAgentNo
    
'   strPrintCust = GetPrintCust '是否列印申請人
   'Modify by Morgan 2004/12/16 改規則
   'strPrintCust = PUB_GetA1K04(m_strPA01, m_strPA02, m_strPA03, m_strPA04)
   strPrintCust = PUB_GetA1K04(m_strPA01, m_strPA02, m_strPA03, m_strPA04, strA1K28, IIf(Me.Text1(7).Text = "1", "1605", "1904"))
   '2004/12/16 end
    
    'Added by Lydia 2014/12/15 請款單請改為依代理人或客戶檔設定的請款幣別
      Dim strA1K33 As String, strA1K18 As String
      'Modified by Sindy 2016/11/30
      'strA1K33 = PUB_GetInitCurrPrintType(m_strPA01, strA1K28, strA1K18, dblUSRate)
      'Modified by Morgan 2018/4/27 +strA1K27
      strA1K33 = PUB_GetInitCurrPrintType(m_strPA01, strA1K28, strA1K18, dblUSRate, m_strPA02, m_strPA03, m_strPA04, strA1K27)
      '2016/11/30 END
      
    Dim strDisc As String '折扣
    strDisc = 1 - (PUB_GetA1L07Disc(m_strPA01, m_strPA02, m_strPA03, m_strPA04, IIf(Me.Text1(7).Text = "1", "1605", "1904"), strSrvDate(2)) / 100)
    'A1K11要先扣除折扣才存檔
    '美金取整數位(無條件捨去)
    'Added by Lydia 2014/12/15 請款單請改為依代理人或客戶檔設定的請款幣別
'   strSql = "INSERT INTO ACC1K0 (A1K01,            A1K02,                     A1K06,A1K07,A1K09,A1K10,           A1K11,                                  A1K12,A1K13,              A1K14,          A1K15,                  A1K16,          A1K18,A1K30,A1K08,                                                                                                 A1K03,A1K27,A1K28,A1K04,A1K21,A1K19,A1K20 ) " & _
            "VALUES  ('" & m_strSerialNo & "'," & (ServerDate - 19110000) & ",0,    NULL,    0," & dblUSRate & "," & Val(m_strCF07) + Val(m_strCF08) - Val((Val(m_strCF07) + Val(m_strCF08)) * Val(strDisc)) & ",NULL,'" & m_strPA01 & "','" & m_strPA02 & "','" & m_strPA03 & "','" & m_strPA04 & "','USD',0, " & Fix(Val("" & IIf(dblUSRate = 0, (Val(m_strCF07) + Val(m_strCF08) - Val((Val(m_strCF07) + Val(m_strCF08)) * Val(strDisc))), (Val(m_strCF07) + Val(m_strCF08) - Val((Val(m_strCF07) + Val(m_strCF08)) * Val(strDisc))) / dblUSRate))) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "','" & strUserNum & "'," & ServerDate - 19110000 & "," & ServerTime & " )"
    strSql = "INSERT INTO ACC1K0 (A1K01,            A1K02,                     A1K06,A1K07,A1K09,A1K10,           A1K11,                                  A1K12,A1K13,              A1K14,          A1K15,                  A1K16,          A1K18,A1K30,A1K08,                                                                                                 A1K03,A1K27,A1K28,A1K04,A1K21,A1K19,A1K20,A1K33 ) " & _
            "VALUES  ('" & m_strSerialNo & "'," & (ServerDate - 19110000) & ",0,    NULL,    0," & dblUSRate & "," & Val(m_strCF07) + Val(m_strCF08) - Val((Val(m_strCF07) + Val(m_strCF08)) * Val(strDisc)) & ",NULL,'" & m_strPA01 & "','" & m_strPA02 & "','" & m_strPA03 & "','" & m_strPA04 & "','" & strA1K18 & "',0, " & Fix(Val("" & IIf(dblUSRate = 0, (Val(m_strCF07) + Val(m_strCF08) - Val((Val(m_strCF07) + Val(m_strCF08)) * Val(strDisc))), (Val(m_strCF07) + Val(m_strCF08) - Val((Val(m_strCF07) + Val(m_strCF08)) * Val(strDisc))) / dblUSRate))) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "','" & strUserNum & "'," & ServerDate - 19110000 & "," & ServerTime & ",'" & strA1K33 & "' )"
  
   cnnConnection.Execute strSql
   '3:新增一筆ACC1L0
   strSql = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L06,A1L07,A1L02,A1L04,A1L05,A1L10,A1L08,A1L09) " & _
            "VALUES  ('" & m_strSerialNo & "','FCP',''," & (Val(m_strCF07) + Val(m_strCF08)) * Val(strDisc) & ",'001','" & IIf(Me.Text1(7).Text = "1", "1605", "1904") & "'," & Val(m_strCF07) + Val(m_strCF08) & ",'" & strUserNum & "'," & ServerDate - 19110000 & "," & ServerTime & " )"
   cnnConnection.Execute strSql
   
   PUB_UpdateA1k08 m_strSerialNo 'Added by Morgan 2012/11/2 更新請款單外幣金額
   
   '4:新增ACC1W0
   strSql = "INSERT INTO ACC1W0 " & _
            "VALUES  ('" & m_strSerialNo & "','" & m_strCP09 & "')"
   cnnConnection.Execute strSql
   '5:更新新增的C類收文號
   strSql = "UPDATE CASEPROGRESS SET CP60='" & m_strSerialNo & "' WHERE CP09='" & m_strCP09 & "'"
   cnnConnection.Execute strSql
   
   PUB_PointAutoassign m_strSerialNo, True 'Add by Morgan 2010/4/21 自動分配點數
   
   cnnConnection.CommitTrans
   strUpdStatus = "2"
   SaveData = True
   
    'Added by Lydia 2016/11/17 以請款對象檢查是否存在於國外固定寄催款單代理人檔(ACC225)且下次寄發日期＞系統日，若存在則顯示訊息提醒操作人員
    If m_strSerialNo <> "" And strA1K28 <> "" Then
       If PUB_ChkAcc225MsgList(m_strSerialNo, strA1K28, m_strPA01, m_strPA02, m_strPA03, m_strPA04) Then
       End If
    End If
    'end 2016/11/17
    
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
   Select Case Me.Text1(7).Text
   Case "1"
      '第幾年至幾年費
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','第幾年至幾年費'," & CNULL(Me.Text1(4).Text) & ")"
      '年費法定期限
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','年費法定期限'," & CNULL(Val(Me.Text1(5).Text) + 19110000) & ")"
      '年費延展繳費日欄
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','年費延展繳費日'," & CNULL(Val(Me.Text1(6).Text) + 19110000) & ")"
      
      'Added by Morgan 2015/5/21
      '一案兩請提醒
'Removed by Morgan 2022/4/7 消滅函也要用，改寫在basLetter公用
'      If m_strPA08 = "2" Then
'         'Modified by Morgan 2017/1/24 +判斷發明無證書號才帶(+ and pa22 is null) --David
'         strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa11,pa77,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CNo" & _
'            " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & m_strPA01 & "' and cm02='" & m_strPA02 & "' and cm03='" & m_strPA03 & "' and cm04='" & m_strPA04 & "'" & _
'            " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & m_strPA01 & "' and cm06='" & m_strPA02 & "' and cm07='" & m_strPA03 & "' and cm08='" & m_strPA04 & "') X" & _
'            ",patent a where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null and pa22 is null"
'         intI = 1
'         Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            i = i + 1
'            strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'               "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','一案兩請新型案要印','♀')"
'
'            i = i + 1
'            strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'               "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','發明案申請號','" & adoRecordset("pa11") & "')"
'
'            i = i + 1
'            strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'                "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','發明案彼所案號','" & IIf(IsNull(adoRecordset("pa77")), "", "" & adoRecordset("pa77")) & "')"
'
'            i = i + 1
'            strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'               "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','發明案本所案號','" & adoRecordset("CNo") & "')"
'
'         End If
'      End If
'end 2022/4/7
      'end 2015/5/21
      
      '新增請款單
      If Text1(8) = "Y" Then
         i = i + 1
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','列印備註','" & vbCrLf & "We take the liberty of enclosing our debit note for your kind settlement." & "')"
      End If
   End Select
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(i, strTxt) Then
   If Not ClsLawExecSQL(i, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If

End Sub

Private Function GetCFFieldData(strFieldName As String, strCF01 As String, strCF02 As String, strCF03 As String) As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   
   GetCFFieldData = ""
   StrSQLa = "Select " & strFieldName & " From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "'"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      GetCFFieldData = "" & rsA.Fields(0).Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim m_CP(1 To TF_CP) As String
End Sub

Private Sub Form_Load()
   Dim ii As Integer
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
    
   MoveFormToCenter Me
   
   'Modify by Morgan 2011/3/15 改共用且不要排除預設印表機
   PUB_SetPrinter Me.Name, Combo1
   PUB_SetPrinter Me.Name, Combo2
'end 2011/3/15

End Sub

Private Sub Form_Unload(Cancel As Integer)

   'Copy from cmdExit_Click by Morgan 2004/10/26
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
    '若請款單印表機變動, 則更新列印設定
    If Me.Combo2.Text <> Me.Combo2.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
    End If
   '2004/10/26 end
   
    PUB_SendMailCache 'Added by Morgan 2022/4/7
   
    Set frm060319 = Nothing
End Sub

Private Sub Text1_Change(Index As Integer)
    Select Case Index
    Case 7 '通知函類別
        '年費逾期
        If Me.Text1(7).Text = "1" Then
           Me.Label17(6).Visible = True
           Me.Label17(7).Visible = True
           Me.Label17(8).Visible = True
           Me.Text1(4).Visible = True
           Me.Text1(5).Visible = True
           Me.Text1(6).Visible = True
           Me.Text1(8).Visible = True
        '檢還證據樣品請款
        Else
           Me.Label17(6).Visible = False
           Me.Label17(7).Visible = False
           Me.Label17(8).Visible = False
           Me.Text1(4).Visible = False
           Me.Text1(5).Visible = False
           Me.Text1(6).Visible = False
           Me.Text1(8).Visible = False
        End If
    End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Me.Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 7 '通知函類別
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
            MsgBox "通知函類別只能輸入 1 或 2 !!!", vbExclamation + vbOKOnly
            KeyAscii = 0
         End If
      Case 8 '新增請款單
         KeyAscii = UpperCase(KeyAscii)
            If KeyAscii <> 89 And KeyAscii <> 8 Then
            MsgBox "新增請款單只能輸入 Y 或空白 !!!", vbExclamation + vbOKOnly
            KeyAscii = 0
         End If
      Case 0, 2
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
   'Modified by Lydia 2017/01/25 跳離案件流水號(PA02)後，帶出案件資料
   'Case 3
   Case 3, 1
      If Me.Text1(0).Text <> "" And Me.Text1(1).Text <> "" Then
         ClearControl
         ReadData
         
      'Removed by Morgan 2022/4/7 確定再檢查
      'Else
      '   MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
      '   Me.Text1(1).SetFocus
      '   TextInverse Me.Text1(1)
      '   Exit Sub
      'end 2022/4/7
      End If
   Case 7 '通知函類別
      If Me.Text1(7).Text = "2" Then
         If Not CheckCaseProperty Then
            Me.Text1(1).SetFocus
            TextInverse Me.Text1(1)
            Exit Sub
         End If
      End If
      'Add By Cheng 2003/01/03
      If Me.Text1(7).Text = "1" Then
        '取得下次繳費資料
        GetNextPayData
      End If
   End Select
End Sub

Private Function ReadData() As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
    
    ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 清除查詢印表記錄檔欄位
    ReadData = False
    m_strPA01 = Me.Text1(0).Text
    m_strPA02 = Left(Me.Text1(1).Text & "000000", 6)
    m_strPA03 = Left(Me.Text1(2).Text & "0", 1)
    m_strPA04 = Left(Me.Text1(3).Text & "00", 2)
    pub_QL05 = pub_QL05 & ";" & Label17(1) & m_strPA01 & "-" & m_strPA02 & "-" & m_strPA03 & "-" & m_strPA04 'Add By Sindy 2010/12/7
    m_strPA08 = Empty
    m_strPA09 = Empty
    Me.cbo(0).Clear
    Me.cbo(1).Clear
    
    If Text1(7) = "1" Then
       pub_QL05 = pub_QL05 & ";" & Left(Label17(10), 6) & "1:年費逾期" 'Add By Sindy 2010/12/7
    ElseIf Text1(7) = "2" Then
       pub_QL05 = pub_QL05 & ";" & Left(Label17(10), 6) & "2:檢還證據樣品請款" 'Add By Sindy 2010/12/7
    End If
    
    StrSQLa = "SELECT * FROM PATENT,CUSTOMER,FAGENT,PATENTTRADEMARKMAP WHERE " & ChgPatent(m_strPA01 & m_strPA02 & m_strPA03 & m_strPA04) & " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        InsertQueryLog (rsA.RecordCount) 'Add By Sindy 2010/12/7
        m_strPA08 = "" & rsA.Fields("PA08").Value
        m_strPA09 = "" & rsA.Fields("PA09").Value
        m_strPA26 = "" & rsA.Fields("PA26").Value
        m_strPA27 = "" & rsA.Fields("PA27").Value
        m_strPA28 = "" & rsA.Fields("PA28").Value
        m_strPA29 = "" & rsA.Fields("PA29").Value
        m_strPA30 = "" & rsA.Fields("PA30").Value
        m_strPA72 = "" & rsA.Fields("PA72").Value
        m_strPA75 = "" & rsA.Fields("PA75").Value
        m_strPA14 = "" & rsA.Fields("PA14").Value 'Add by Morgan 2006/2/22
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
        Me.Label1(1).Caption = IIf(Not IsNull(rsA("FA05").Value), ("" & rsA("FA05").Value) & " " & rsA("FA63").Value & " " & rsA("FA64").Value & " " & rsA("FA65").Value, IIf(Not IsNull(rsA("FA04")), "" & rsA("FA04"), "" & rsA("FA06")))
        '專利種類
        Me.Label1(2).Caption = "" & rsA("PTM03").Value
        ReadData = True
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/7
        MsgBox "基本檔無此案號資料!!!", vbExclamation + vbOKOnly
        Me.Text1(1).SetFocus
        Text1_GotFocus 1
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End Function

Private Function CheckCaseProperty() As Boolean
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset

   CheckCaseProperty = True
   m_strCP09 = Empty
   StrSQLa = "SELECT * FROM CASEPROGRESS WHERE " & ChgCaseprogress(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text) & " AND CP10='1904' AND CP05 IS NOT NULL AND CP09>='C' ORDER BY CP05 DESC, CP09 DESC "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount <= 0 Then
      MsgBox "此案件無 檢還樣品證據 的來函資料!!!", vbExclamation + vbOKOnly
      CheckCaseProperty = False
   Else
      m_strCP09 = "" & rsA("CP09").Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

Private Sub ClearControl()
   Me.Label1(0).Visible = False
   Me.cbo(0).Clear
   Me.cbo(1).Clear
   Me.Label1(1).Caption = Empty
   Me.Label1(2).Caption = Empty
   Me.Text1(4).Text = Empty
   Me.Text1(5).Text = Empty
   Me.Text1(6).Text = Empty
   Me.Text1(7).Text = Empty
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Dim iFeeYear As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   
   Select Case Index
   Case 4 '未繳年度
      If Me.Text1(Index).Text <> "" Then
         StrSQLa = "SELECT * FROM NATION WHERE NA01='" & m_strPA09 & "'"
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            If m_strPA08 = "1" Then
               If Val(Me.Text1(Index).Text) > Val("" & rsA("NA07").Value) Then
                  MsgBox "未繳年度不可大於繳費年度(" & Val("" & rsA("NA07").Value) & ")", vbExclamation + vbOKOnly
                  Cancel = True
               End If
            ElseIf m_strPA08 = "2" Then
               'Modify by Morgan 2006/2/22 控制新型舊法專用年度為12年
               If Val(m_strPA14) < 20040701 Then
                  iFeeYear = 12
               Else
                  iFeeYear = Val("" & rsA("NA09").Value)
               End If
               If Val(Me.Text1(Index).Text) > iFeeYear Then
                  MsgBox "未繳年度不可大於繳費年度(" & iFeeYear & ")", vbExclamation + vbOKOnly
                  Cancel = True
               End If
            ElseIf m_strPA08 = "3" Then
               If Val(Me.Text1(Index).Text) > Val("" & rsA("NA11").Value) Then
                  MsgBox "未繳年度不可大於繳費年度(" & Val("" & rsA("NA11").Value) & ")", vbExclamation + vbOKOnly
                  Cancel = True
               End If
            End If
         Else
            MsgBox "無專利種類的繳費年度資料!!!", vbExclamation + vbOKOnly
            Cancel = True
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
      End If
   Case 5 '原年費期限
      If Me.Text1(Index).Text <> "" Then
         If ChkDate(Me.Text1(Index).Text) = False Then
            Cancel = True
         ElseIf Val(Me.Text1(Index).Text) + 19110000 > strSrvDate(1) Then
            MsgBox "原年費期限不可大於系統日!!!", vbExclamation + vbOKOnly
            Cancel = True
         End If
      End If
   Case 6 '最後期限
      If Me.Text1(Index).Text <> "" Then
         If ChkDate(Me.Text1(Index).Text) = False Then
            Cancel = True
            'Modify By Cheng 2003/01/01
'         ElseIf Val(Me.Text1(Index).Text) + 19110000 > ServerDate Then
'            MsgBox "最後期限不可大於系統日!!!", vbExclamation + vbOKOnly
'            Cancel = True
'         End If
         ElseIf Val(Me.Text1(Index).Text) + 19110000 < strSrvDate(1) Then
            MsgBox "最後期限不可小於系統日!!!", vbExclamation + vbOKOnly
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
If Me.Text1(7).Text = "" Then
   MsgBox "請輸入通知函類別!!!", vbExclamation + vbOKOnly
   Me.Text1(7).SetFocus
   TextInverse Me.Text1(7)
   Exit Function
End If
If Me.Text1(7).Text = "2" Then
   If Not CheckCaseProperty Then
      Me.Text1(1).SetFocus
      TextInverse Me.Text1(1)
      Exit Function
   End If
End If
If Me.Text1(4).Text = "" And Me.Text1(4).Visible Then
   MsgBox "請輸入未繳年度!!!", vbExclamation + vbOKOnly
   Me.Text1(4).SetFocus
   TextInverse Me.Text1(4)
   Exit Function
End If
If Me.Text1(5).Text = "" And Me.Text1(5).Visible Then
   MsgBox "請輸入原年費期限!!!", vbExclamation + vbOKOnly
   Me.Text1(5).SetFocus
   TextInverse Me.Text1(5)
   Exit Function
End If
If Me.Text1(6).Text = "" And Me.Text1(6).Visible Then
   MsgBox "請輸入最後期限!!!", vbExclamation + vbOKOnly
   Me.Text1(6).SetFocus
   TextInverse Me.Text1(6)
   Exit Function
End If
For Each objTxt In Text1
   If objTxt.Enabled = True And objTxt.Visible = True Then
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
TxtValidate = True
End Function


Private Function GetUSRate() As Double
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetUSRate = 0
StrSQLa = "SELECT USXR02 FROM USXRATE WHERE USXR01<=" & (ServerDate - 19110000) & " ORDER BY USXR01 DESC "
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   GetUSRate = rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

Private Sub ProcessPrint()
   MsgBox "列印請款單，請更換紙張!!!", vbExclamation + vbOKOnly
   Screen.MousePointer = vbHourglass
   Load Frmacc2480
   With Frmacc2480
      .Text1.Text = m_strSerialNo
      .Text2.Text = m_strSerialNo
      .Combo1.Text = Me.Combo2.Text
      'Add by Morgan 2008/5/23 +傳是否存電子檔參數
      .m_bBeCalled = True
      .m_CallPrevForm = Me.Name  'Added by Lydia 2020/01/06 呼叫請款單的程式名稱
      'Modified by Morgan 2014/6/3
      '.m_bEMail = m_bolEmail
      '.m_bPaper = m_bolPlusPaper
      .m_bEMail = m_bolDNEmail
      .m_bPaper = m_bolDNPlusPaper
      'end 2014/6/3
      'end 2008/5/23
      .Command2_Click: DoEvents
   End With
   Unload Frmacc2480
   Screen.MousePointer = vbDefault
End Sub

'取得下次繳費年度資料
Private Sub GetNextPayData()
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim arrPA72
Dim ArrYear
Dim strMaxPA72 '目前繳費年度
Dim strNextPA72 '下次繳費年度
Dim ii As Integer
Dim jj As Integer
    
   'Add by Morgan 2005/7/20
   Me.Text1(4).Text = ""
   Me.Text1(5).Text = ""
   Me.Text1(6).Text = ""
   
    '取得下次繳費年度
    'Modify by Morgan 2005/7/20
    'strSQLB = "Select DECODE(" & m_strPA08 & ", '1', NA21, '2', NA23, NA25 ) From Nation Where NA01='" & m_strPA09 & "' "
    ''
    StrSqlB = "Select DECODE(PA08, '1', NA21, '2', NA23, NA25 ) C1,DECODE(DECODE(PA08,'1',NA06,'2',NA08,NA10),'2',PA10,'4',PA20,'5',PA14,'6',PA21) C2 From Patent,Nation Where PA01='" & m_strPA01 & "' AND PA02='" & m_strPA02 & "' AND PA03='" & m_strPA03 & "' AND PA04='" & m_strPA04 & "' AND NA01=PA09"
    rsB.CursorLocation = adUseClient
    rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
    If rsB.RecordCount > 0 Then
        If "" & rsB.Fields(0).Value <> "" Then
            If "" & m_strPA72 <> "" Then
                arrPA72 = Split(m_strPA72, ",")
                strMaxPA72 = arrPA72(UBound(arrPA72))
            Else
                strMaxPA72 = ""
            End If
            'Modify by Morgan 2006/2/22 控制新型舊法專用年度為12年
            'ArrYear = Split(rsB.Fields(0).Value, ",")
            If m_strPA08 = "2" And Val(m_strPA14) < 20040701 Then
               ArrYear = Split("1,2,3,4,5,6,7,8,9,10,11,12", ",")
            Else
               ArrYear = Split(rsB.Fields(0).Value, ",")
            End If
            '2006/2/22 end
            
            jj = -100
            strNextPA72 = ""
            'Add by Morgan 2005/7/20 改以系統日+起算日判斷
            strExc(0) = "" & rsB.Fields(1).Value '年費起算日
            If strExc(0) <> "" Then
               For ii = LBound(ArrYear) To UBound(ArrYear)
                  strExc(1) = CompDate(2, -1, CompDate(0, ArrYear(ii), strExc(0)))
                  If strExc(1) > strSrvDate(1) Then
                     strNextPA72 = ArrYear(ii)
                     Me.Text1(5).Text = ChangeWStringToTString(strExc(2))
                     '先+1天,再+6月,再-1天這樣才可控制月底
                     Me.Text1(6).Text = ChangeWStringToTString(CompDate(2, -1, CompDate(1, 6, CompDate(2, 1, strExc(2)))))
                     Exit For
                  End If
                  strExc(2) = strExc(1)
               Next ii
            Else
            '2005/7/20 end
               For ii = LBound(ArrYear) To UBound(ArrYear)
                   strNextPA72 = ArrYear(ii)
                   If ii = jj + 1 Then Exit For
                   If strMaxPA72 = ArrYear(ii) Then jj = ii
               Next ii
            End If
            Me.Text1(4).Text = strNextPA72
        Else
            Me.Text1(4).Text = ""
        End If
    Else
        Me.Text1(4).Text = ""
    End If
    If rsB.State <> adStateClosed Then rsB.Close
    Set rsB = Nothing
    'Modify by Morgan 2005/7/20 加控制未取得期限時才跑
    If Me.Text1(5).Text = "" Then
      '取得期限
      StrSqlB = "Select NP09 From NextProgress Where " & ChgNextProgress(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text) & " And NP07='" & 年費 & "' And NP06 IS NULL "
      rsB.CursorLocation = adUseClient
      rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
      If rsB.RecordCount > 0 Then
          Me.Text1(5).Text = ChangeWStringToTString("" & rsB.Fields(0).Value)
          Me.Text1(6).Text = ChangeWDateStringToTString(DateAdd("m", 6, ChangeWStringToWDateString("" & rsB.Fields(0).Value)))
      End If
      If rsB.State <> adStateClosed Then rsB.Close
      Set rsB = Nothing
   End If
End Sub
