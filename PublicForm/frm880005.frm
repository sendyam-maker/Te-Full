VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm880005 
   BorderStyle     =   1  '單線固定
   Caption         =   "發E-Mail"
   ClientHeight    =   5760
   ClientLeft      =   168
   ClientTop       =   996
   ClientWidth     =   9168
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9168
   Begin VB.ListBox lstMailCC 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      ItemData        =   "frm880005.frx":0000
      Left            =   840
      List            =   "frm880005.frx":0002
      Style           =   1  '項目包含核取方塊
      TabIndex        =   23
      Top             =   600
      Width           =   4000
   End
   Begin VB.CheckBox chkSentItem 
      Caption         =   "要留寄件備份"
      Height          =   195
      Left            =   7470
      TabIndex        =   22
      Top             =   5490
      Width           =   1500
   End
   Begin VB.CheckBox chkShowPic 
      Caption         =   "崁圖請勾，不勾一律帶附件"
      Height          =   225
      Left            =   4200
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   2880
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   330
      Top             =   3390
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin MSMAPI.MAPIMessages MAPIMessages 
      Left            =   300
      Top             =   5100
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession 
      Left            =   300
      Top             =   4560
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.CommandButton cmdCards 
      Caption         =   "通訊錄(&B)"
      Height          =   300
      Left            =   8088
      TabIndex        =   8
      Top             =   528
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   7872
      TabIndex        =   7
      Top             =   12
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "傳送(&S)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6936
      TabIndex        =   6
      Top             =   12
      Width           =   912
   End
   Begin MSForms.TextBox Text7 
      Height          =   500
      Left            =   9390
      TabIndex        =   16
      Top             =   3615
      Width           =   1250
      VariousPropertyBits=   671105051
      Size            =   "2205;882"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text6 
      Height          =   500
      Left            =   9390
      TabIndex        =   15
      Top             =   3015
      Width           =   1250
      VariousPropertyBits=   671105051
      Size            =   "2205;882"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   500
      Left            =   9375
      TabIndex        =   14
      Top             =   2400
      Width           =   1250
      VariousPropertyBits=   671105051
      Size            =   "2205;882"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   500
      Left            =   9360
      TabIndex        =   13
      Top             =   1800
      Width           =   1250
      VariousPropertyBits=   671105051
      Size            =   "2205;882"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   500
      Left            =   9360
      TabIndex        =   12
      Top             =   1200
      Width           =   1250
      VariousPropertyBits=   671105051
      Size            =   "2205;882"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   500
      Left            =   9375
      TabIndex        =   11
      Top             =   615
      Width           =   1250
      VariousPropertyBits=   671105051
      Size            =   "2205;882"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Left            =   9375
      TabIndex        =   10
      Top             =   225
      Width           =   705
      VariousPropertyBits=   671105051
      Size            =   "1244;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtEmail 
      Height          =   300
      Index           =   5
      Left            =   30
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
      VariousPropertyBits=   671105051
      Size            =   "2619;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtEmail 
      Height          =   330
      Index           =   4
      Left            =   1560
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   5355
      VariousPropertyBits=   671105051
      Size            =   "9446;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtEmail 
      Height          =   300
      Index           =   3
      Left            =   4920
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   3075
      VariousPropertyBits=   671105053
      Size            =   "5424;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtEmail 
      Height          =   3495
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   1860
      Width           =   8175
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "14420;6165"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtEmail 
      Height          =   330
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   1500
      Width           =   8175
      VariousPropertyBits=   671105051
      Size            =   "14420;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtEmail 
      Height          =   330
      Index           =   0
      Left            =   4920
      TabIndex        =   0
      Top             =   600
      Width           =   3075
      VariousPropertyBits=   671105051
      Size            =   "5424;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "收件者："
      Height          =   180
      Left            =   120
      TabIndex        =   25
      Top             =   240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "(可複選)"
      Height          =   180
      Left            =   120
      TabIndex        =   24
      Top             =   990
      Width           =   660
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "附件位置，中間加* 號："
      Height          =   180
      Left            =   1590
      TabIndex        =   20
      Top             =   30
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "以下欄位不會秀"
      Height          =   480
      Left            =   9300
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "主旨："
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   1542
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "內容："
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收件者："
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   690
      Width           =   720
   End
End
Attribute VB_Name = "frm880005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/03/11 改成Form2.0 ; txtEmail(index)、因為txtEmail(index)會傳到Text1~Text7所以也要修改 ;
                                        'Pub_SendMail依然是使用這支發送,所以追加修改
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit

Public bolLeave As Boolean
'add by nickc 2005/03/17 改成用 winsock 發
Dim Mailing As Boolean
Dim Result$, Sec%
'Modify by Morgan 2009/6/26 改IP比較直接
'Const Server$ = "exchange"
'Modifed by Morgan 2012/9/20
'Const Server$ = "192.168.1.2"
Const Server$ = "192.168.1.10"

Const Domain$ = "Taie"
Const TimeOut% = 20
'edit by nickc 2007/01/04
'Const MailBefore$ = "IMCEAEX-_O=TAIE_OU=DOMAIN_CN=RECIPIENTS_CN="
Const MailBefore$ = ""
'add by nickc 2007/01/23 加入，若是 北寄北 且收件者部門別為'F' 開頭的走 x400
Const MailBeforeF$ = "IMCEAEX-_O=TAIE_OU=DOMAIN_CN=RECIPIENTS_CN="
Const MailAfter$ = "@taie.com.tw"
Const MailReplace$ = "EX:/O=TAIE/OU=DOMAIN/CN=RECIPIENTS/CN="
Public bolByMsa As Boolean
Dim SeekMailErr As String

'Add by Morgan 2010/12/17
Public m_FromID As String '寄件帳號
Public m_FromName As String '寄件名稱
Public m_ReplyTo As String '回覆信箱 Add by Morgan 2011/4/22

Dim m_ErrStat As String '郵件錯誤階段

'Added by Morgan 2013/3/26
'Dim m_SMTP_IP_System As String, m_SMTP_IP_EPAPER As String
Dim m_SMTP_IP_System As String, m_SMTP_IP_BK As String

'Dim m_SMTP_IP_OUT As String 'Modified by Morgan 2015/1/22 +m_SMTP_IP_OUT 'Removed by Morgan 2015/6/8
Dim m_FinalResult As String 'Added by Morgan 2013/9/4
Dim m_MailLog As String 'Added by Morgan 2013/10/18
Dim m_MailSub As String 'Added by Morgan 2017/2/14
'Added by Lydia 2015/08/03
Public m_BCCto As String '密件副本收件人
'Added by Lydia 2016/05/17 有預設收件人,設定副本收件人可勾選
Public bolCCList As Boolean
Public m_EEP01 As String 'Add By Sindy 2018/7/9 文號
Public bolShowErrMsg As Boolean 'Add By Sindy 2016/12/12 是否顯示傳送失敗的訊息
Public bolHTML As Boolean 'Added by Morgan 2017/11/2 是否HTML格式
Public m_strUpdWhere As String, m_strTableName As String 'Add By Sindy 2017/11/17 記錄寄送資訊
Public bolImportant As Boolean 'Added by Morgan 2018/9/11 是否設定高重要性
'Modify By Sindy 2019/4/25 因董事長無收到作業系統發的E-Mail，請協助處理。
'                          發現,因董事長在系統裡是設定不寄信
Public bol63001CanSend As Boolean
Public m_iSignatureID As Integer 'Added by Morgan 2019/8/27
Public bolGetReceipt As Boolean 'Added by Morgan 2021/4/9 是否要求讀取回條

'Added by Morgan 2022/3/23
Public m_ByUTF8 As Boolean '是否指定編碼方式為Utf-8
Public m_SMTP_TEST As String 'Added by Morgan 2024/4/17
Public m_SaveMailCP09 As String 'Added by Morgan 2025/6/27

Dim strTemp As String 'Added by Morgan 2025/5/9 原strExc(1)改為用此變數

Private Sub cmdCards_Click()
Dim i  As Integer
Dim IsPutUserNum As Boolean
On Error GoTo ErrHand

'add by nickc 2005/03/17 改用 winsock
   IsPutUserNum = False
   MAPISession.LogonUI = False
   MAPISession.SignOn
   MAPIMessages.SessionID = MAPISession.SessionID

MAPIMessages.MsgIndex = -1
MAPIMessages.Show False
txtEmail(0) = ""
txtEmail(3) = ""
For i = 0 To MAPIMessages.RecipCount - 1
       MAPIMessages.RecipIndex = i
       'edit by nickc 2005/03/17
       'txtEmail(0) = txtEmail(0) + MAPIMessages.RecipDisplayName + " ;"
       'Modified by Morgan 2015/8/3 新版 Exchange 格式不同無法抽出員工號
       If InStr(MAPIMessages.RecipAddress, MailReplace) > 0 Then
         txtEmail(3) = txtEmail(3) + MAPIMessages.RecipDisplayName + " ;"
         txtEmail(0) = txtEmail(0) + Replace(UCase(MAPIMessages.RecipAddress), MailReplace, "") + " ;"
      Else
         MsgBox "收件者 [ " & MAPIMessages.RecipDisplayName & " ] 帳號無法確認，請自行輸入員工號!!", vbExclamation
      End If
      
Next
MAPISession.SignOff
Exit Sub
ErrHand:
If Err.NUMBER = 32003 And IsPutUserNum = False Then
   MAPISession.UserName = strUserNum
   IsPutUserNum = True
   Resume
End If
If Err.NUMBER <> 32001 Then
   ErrorMsg
End If
End Sub

'Add by Morgan 2010/4/28
Private Function SendMAPIMail() As Boolean
   Dim stSubject As String
   Dim stContent As String
   Dim strTo As String
   Dim IsPutUserNum As Boolean
   
On Error GoTo ErrHnd

   strTo = txtEmail(0)
   stSubject = txtEmail(1)
   stContent = txtEmail(2)
   
   DoEvents
   MAPISession.LogonUI = False
   MAPISession.UserName = strUserNum
   MAPISession.SignOn
   MAPIMessages.SessionID = MAPISession.SessionID
   MAPIMessages.MsgIndex = -1
   MAPIMessages.Compose
   'Modify By Sindy 2014/1/16
   'MAPIMessages.MsgSubject = "◎系統代發◎" & stSubject
   MAPIMessages.MsgSubject = "◎" & IIf(Pub_StrUserSt03 = "M51" And PUB_GetST05(strUserNum) <> "" And UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱), PUB_GetDbTerminal, "") & stSubject
   '2014/1/16 END
   MAPIMessages.MsgNoteText = stContent
   MAPIMessages.RecipIndex = 0
   MAPIMessages.RecipDisplayName = strTo
   MAPIMessages.ResolveName
   MAPIMessages.Send
   MAPISession.SignOff
   
   SendMAPIMail = True
   Exit Function
   
ErrHnd:
   If Err.NUMBER = 32003 And IsPutUserNum = False Then
      MAPISession.UserName = strUserNum
      IsPutUserNum = True
      Resume
   End If
   If Err.NUMBER <> 32001 Then
      ErrorMsg
   End If
   
End Function

'Modify By Cheng 2003/06/30
'Private Sub cmdok_Click(Index As Integer)
Public Sub cmdok_Click(Index As Integer)
On Error GoTo ErrHand
'edit by nickc 2005/03/17 改用 winsock
'If Index = 0 Then
'   MAPIMessages.MsgIndex = -1
'   MAPIMessages.Compose
'   MAPIMessages.MsgSubject = txtEmail(1)
'   MAPIMessages.MsgNoteText = txtEmail(2)
'   SetAdress txtEmail(0)
'   MAPIMessages.Send
'   bolLeave = True
'Else
'   bolLeave = False
'End If
Dim ArrMail As Variant
Dim MailCnt As Integer
Dim strOfficeKind As String
Dim bolSMTPFail As Boolean
'Add By Sindy 2009/09/23
Dim stToName As String, stToMail As String
Dim stCCToName As String, stCCToMail As String
Dim arrName As Variant
'Add by Morgan 2010/7/26
Dim strLocalIP As String
'Added by Lydia 2015/11/25
Dim jj As Integer
Dim intMaxEEP02 As String 'Add By Sindy 2018/7/9
Dim strDeptName As String 'Added by Morgan 2024/1/12

If Index = 0 Then
  'Added by Lydia 2015/11/25 勾選收件者
  'Modified by Lydia 2016/05/17 有預設收件人,設定副本收件人可勾選
  If bolCCList Then
      For jj = 0 To lstMailCC.ListCount - 1
         If lstMailCC.Selected(jj) = True Then
            strTemp = Trim(Mid(lstMailCC.List(jj), InStr(lstMailCC.List(jj), " ") + 1, 5))
            If txtEmail(5) = "" Then
               txtEmail(5) = strTemp
            Else
               txtEmail(5) = txtEmail(5) & ";" & strTemp
            End If
         End If
      Next
      If Trim(txtEmail(0)) = "" Then
         txtEmail(0) = txtEmail(5)
         txtEmail(5) = ""
      End If
  Else
    If Trim(txtEmail(0)) = "" Then
      For jj = 0 To lstMailCC.ListCount - 1
         If lstMailCC.Selected(jj) = True Then
            strTemp = Trim(Mid(lstMailCC.List(jj), InStr(lstMailCC.List(jj), " ") + 1, 5))
            If txtEmail(0) = "" Then
               txtEmail(0) = strTemp
            Else
               txtEmail(0) = txtEmail(0) & ";" & strTemp
            End If
         End If
      Next
    End If
  End If
  'end 2016/05/17
  'end 2015/11/25

   'add by nickc 2005/08/11 沒有收件者，要秀訊息
   'Modify By Sindy 2016/5/11 + 增加顯示IIf(Pub_strLogText <> "", vbCrLf & "Pub_strLogText=" & Pub_strLogText, "")
   'Modify By Sindy 2021/10/7 + 秀玲:把主旨也顯示出來，否則不知道怎麼去追蹤程式是哪裡出錯
   If Trim(txtEmail(0)) = "" Then
      MsgBox "沒有收件者，請問要寄給誰？" & _
      IIf(Pub_strLogText <> "", vbCrLf & "Pub_strLogText=" & Pub_strLogText, "") & _
      vbCrLf & "(主旨=" & txtEmail(1) & ")", vbInformation, "寄信錯誤！"
      Exit Sub
   End If
         
  'Add by Morgan
  If chkSentItem.Value = 1 Then
      If SendMAPIMail = False Then
         bolMailSendOk = False
         Exit Sub
      Else
         bolMailSendOk = True
      End If
  'End 2010/4/28
  ElseIf Mailing = False Then

         strOfficeKind = PUB_GetST06(strUserNum)
         
         'Add by Morgan 2009/1/10 對外信件一律不顯示名字 --Robert
         If bolByMsa = True Then
            'Modified by Morgan 2024/1/26 改也顯示名字 --Robert
            'Text2.Text = ""
            Text2.Text = strUserName
         Else
            Text2.Text = strUserName
            'Added by Morgan 2024/1/12
            strDeptName = GetDeptNameA0922(strUserNum)
            If strDeptName <> "" Then Text2.Text = Text2.Text & "(" & strDeptName & ")"
            'end 2024/1/12
         End If

'Modified by Morgan 2012/9/20 統一改用 192.168.1.15 若失敗再改 192.168.1.10
'
'         '若寄件者為北所人員, 則伺服器後面不加@taie.com.tw
'         'edit by nickc 2006/12/28
'         If bolByMsa = True Then
'             'Modify by Morgan 2009/1/10 改用台一的 spam firewall
'             'Text1.Text = "168.95.4.211"
'             Text1.Text = "192.168.1.10"
'         ElseIf strOfficeKind = "1" Then
'            'Modify by Morgan 2010/8/16 北所也改統一送往防火牆(有Log可查且X400也已能接受)
'            'Text1.Text = Server$
'            Text1.Text = "192.168.1.10"
'         Else
'            'Modify by Morgan 2010/9/27 VPN已架設完畢,統一送往防火牆
'            ''Text1.Text = Server$ & Trim(".taie.com.tw")
'            ''Text1.Text = "168.95.4.211"     'msa.hinet.net
'            'Text1.Text = "211.75.113.68"    'exchange 外部
'            ''Add by Morgan 2010/8/16 分所改判斷若 Ping 得到防火牆時改送往防火牆
'            'If SocketsInitialize() Then
'            '   If ping("192.168.1.10") = 0 Then
'            '      Text1.Text = "192.168.1.10"
'            '   End If
'            'End If
'            ''end 2010/8/16
'            Text1.Text = "192.168.1.10"
'            'end 2010/9/27
'         End If
'
'         'Added by Morgan 2012/9/5
'         If InStr(LCase(txtEmail(0)), LCase("@taie.com.tw")) = 0 And InStr(txtEmail(0), "@") > 0 Then
'            Text1.Text = "192.168.1.15"
'         End If
'         'end 2012/9/5
         'Modified by Morgan 2013/3/26 改抓特殊設定
         'Text1.Text = "192.168.1.15"
         'Modified by Morgan 2015/1/23 財務處改用有備份的 SMTP
         'Modified by Morgan 2015/6/8 預設 SMTP 已設定也會經郵件備份系統
         'If m_FromID = strAccMailBox Then
         '   Text1.Text = m_SMTP_IP_OUT
         'Else
            'Text1.Text = m_SMTP_IP_System 'Removed by Morgan 2024/1/10 改依收件人是否所外決定要使用的SMTP
        ' End If
         'end 2015/1/23
'end 2012/9/20

         Text3.Text = strUserNum
         
         If Mid(Text3.Text, 4, 1) = "9" Then Text3.Text = PUB_GetRealUserNo(Text3.Text) '虛編號要改為真實員工號，否則無法回信
         'Add by Morgan 2010/12/17
         '寄件人
         If m_FromID <> "" Then
            Text3.Text = m_FromID
            If m_FromName <> "" Then
               Text2.Text = m_FromName
            Else
               'Added by Morgan 2024/1/24 對外改也顯示名字 --Robert
               If InStr(m_FromID, "@") = 0 Then
                  Text2.Text = GetStaffName(m_FromID)
               Else
               'end 2024/1/24
                  Text2.Text = m_FromID
               End If
            End If
         End If
         
         'Add By Sindy 2017/2/7 主旨後面+分機資訊
         If InStr(txtEmail(1), "[寄件者：") = 0 And InStr(txtEmail(1), "#") = 0 And InStr(txtEmail(1), "]") = 0 Then
            If Len(Text3) = 5 And (InStr(txtEmail(0), "@") = 0 Or InStr(UCase(txtEmail(0)), "@TAIE.COM.TW") > 0) Then
               If Text3 > "6" And Text3 < "F" Then
                  strSql = "select ed01,decode(ed03,'1','北所','2','中所','3','南所','4','高所','其他') ed03,nvl(ed05,'') ed05 from extensiondata where ed02='" & Text3 & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     txtEmail(1) = txtEmail(1) & " [寄件者：" & _
                                   RsTemp.Fields("ed03") & _
                                   IIf(Trim("" & RsTemp.Fields("ed05")) = "", "", "(" & Trim(RsTemp.Fields("ed05")) & ")") & _
                                   "#" & RsTemp.Fields("ed01") & "]"
                  End If
               End If
            End If
         End If
         '2017/2/7 END
         
         Text6.Text = txtEmail(1).Text
         Text7.Text = txtEmail(2).Text
         
         'Add By Sindy 2009/09/23 先整理收件人及副本資訊
         ArrMail = Split(txtEmail(0), ";")
         For MailCnt = 0 To UBound(ArrMail)
            If ArrMail(MailCnt) <> "" Then
               'add by nickc 2006/08/18 檢查離職者發該區最小編號
               'Modify By Sindy 2019/4/25 + bol63001CanSend
               ArrMail(MailCnt) = ChkMailId(ArrMail(MailCnt), bol63001CanSend)
               'Modified by Lydia 2017/03/28 可能不只一人
               'If Replace(UCase(ArrMail(MailCnt)), "@TAIE.COM.TW", "") <> "99999" And Replace(UCase(ArrMail(MailCnt)), "@TAIE.COM.TW", "") <> "99997" Then 'Added by Morgan 2013/5/20 從 SendMail 移來
               '   stToName = stToName & GetPrjSalesNM(Trim(ArrMail(MailCnt))) & ";"
               If InStr(Replace(UCase(ArrMail(MailCnt)), "@TAIE.COM.TW", ""), "99999") = 0 And InStr(Replace(UCase(ArrMail(MailCnt)), "@TAIE.COM.TW", ""), "99997") = 0 Then
               'end 2017/03/28
                  stToMail = stToMail & Trim(ArrMail(MailCnt)) & ";"
               End If 'Added by Morgan 2013/5/20
            End If
         Next MailCnt
         Call PUB_CompRepeatReple(stToMail, "") 'Add By Sindy 2025/2/6 過濾重覆值
         'Added by Lydia 2017/03/28 因為收件人可能不等於txtEmail(0),所以最後抓姓名
         ArrMail = Empty
         ArrMail = Split(stToMail, ";")
         For MailCnt = 0 To UBound(ArrMail)
            If ArrMail(MailCnt) <> "" Then
               stToName = stToName & GetPrjSalesNM(Trim(ArrMail(MailCnt))) & ";"
            End If
         Next MailCnt
         'end 2017/03/28
         If Trim(txtEmail(5).Text) <> "" Then
            ArrMail = Split(txtEmail(5), ";")
            For MailCnt = 0 To UBound(ArrMail)
               If ArrMail(MailCnt) <> "" Then
                  'Modify By Sindy 2019/4/25 + bol63001CanSend
                  ArrMail(MailCnt) = ChkMailId(ArrMail(MailCnt), bol63001CanSend)
                  'Modified by Lydia 2017/03/28 可能不只一人
                  'If Replace(UCase(ArrMail(MailCnt)), "@TAIE.COM.TW", "") <> "99999" And Replace(UCase(ArrMail(MailCnt)), "@TAIE.COM.TW", "") <> "99997" Then 'Added by Morgan 2013/5/20 從 SendMail 移來
                  '   stCCToName = stCCToName & GetPrjSalesNM(Trim(ArrMail(MailCnt))) & ";"
                  If InStr(Replace(UCase(ArrMail(MailCnt)), "@TAIE.COM.TW", ""), "99999") = 0 And InStr(Replace(UCase(ArrMail(MailCnt)), "@TAIE.COM.TW", ""), "99997") = 0 Then
                  'end 2017/03/28
                     stCCToMail = stCCToMail & Trim(ArrMail(MailCnt)) & ";"
                  End If 'Added by Morgan 2013/5/20
               End If
            Next MailCnt
         End If
         Call PUB_CompRepeatReple(stToMail, stCCToMail) 'Add By Sindy 2025/2/6 過濾重覆值
         'Added by Lydia 2017/03/28 因為收件人可能不等於txtEmail(0),所以最後抓姓名
         If stCCToMail <> "" Then
            ArrMail = Empty
            ArrMail = Split(stCCToMail, ";")
            For MailCnt = 0 To UBound(ArrMail)
               If ArrMail(MailCnt) <> "" Then
                   stCCToName = stCCToName & GetPrjSalesNM(Trim(ArrMail(MailCnt))) & ";"
               End If
            Next MailCnt
         End If
         'end 2017/03/28
         
         'Added by Morgan 2024/4/17
         If m_SMTP_TEST <> "" Then
            Text1.Text = m_SMTP_TEST
         Else
         'end 2024/4/17
            'Added by Morgan 2024/1/10
            '為避免重複進備份系統，改依收件人是否所外決定要使用的SMTP
            '收件人有所內同仁時用防火牆否則用備份系統(預設的SMTP)
            Text1.Text = m_SMTP_IP_System
            'Removed by Morgan 2024/6/7 備份系統轉信會有問題(收件人的domain會無故被更動導致無法寄達)
            'ArrMail = Split(stToMail, ";")
            'For MailCnt = 0 To UBound(ArrMail)
            '   If ArrMail(MailCnt) <> "" Then
            '      If InStr(ArrMail(MailCnt), "@") = 0 Or InStr(LCase(ArrMail(MailCnt)), "@taie.com.tw") > 0 Then
            '         Text1.Text = m_SMTP_IP_BK
            '         Exit For
            '      End If
            '   End If
            'Next
            'end 2024/6/7
         End If
         'Added by Morgan 2024/3/21 +CC和BCC也要檢查
         'Removed by Morgan 2024/3/21 先維持原狀，因為這樣的話對外信會不經雲端直接出去，備份系統用收件者就會查不到
'         If Text1.Text <> m_SMTP_IP_BK Then
'            ArrMail = Split(stCCToMail, ";")
'            For MailCnt = 0 To UBound(ArrMail)
'               If ArrMail(MailCnt) <> "" Then
'                  If InStr(ArrMail(MailCnt), "@") = 0 Or InStr(LCase(ArrMail(MailCnt)), "@taie.com.tw") > 0 Then
'                     Text1.Text = m_SMTP_IP_BK
'                     Exit For
'                  End If
'               End If
'            Next
'         End If
'         If Text1.Text <> m_SMTP_IP_BK Then
'            ArrMail = Split(m_BCCto, ";")
'            For MailCnt = 0 To UBound(ArrMail)
'               If ArrMail(MailCnt) <> "" Then
'                  If InStr(ArrMail(MailCnt), "@") = 0 Or InStr(LCase(ArrMail(MailCnt)), "@taie.com.tw") > 0 Then
'                     Text1.Text = m_SMTP_IP_BK
'                     Exit For
'                  End If
'               End If
'            Next
'         End If
         'end 2024/3/21
         'end 2024/1/10

         'Modify By Sindy 2009/09/23
         '寄收件者
         arrName = Split(stToName, ";")
         ArrMail = Split(stToMail, ";")
         For MailCnt = 0 To UBound(ArrMail)
            If ArrMail(MailCnt) <> "" Then
               Text4.Text = arrName(MailCnt)
               Text5.Text = ArrMail(MailCnt)
               'Removed by Morgan 2012/4/18 財務有需要顯示多收件人,不用再判斷是否有副本
               ''Morgan說若無副本時, 收件人還是維持單人顯示
               'If Trim(txtEmail(5).Text) = "" Then
               '   stToName = ArrName(MailCnt)
               '   stToMail = ArrMail(MailCnt)
               'End If
               'end 2012/4/18
               'Modified by Lydia 2015/08/03 +密件副本收件人
              ' If Not SendMail(Text1.text, Text2.text, Text3.text, Text4.text, _
                  Text5.text, Text6.text, Text7.text, stToName, stToMail, bolSMTPFail, _
                  stCCToName, stCCToMail, m_ReplyTo) Then
                If Not SendMail(Text1.Text, Text2.Text, Text3.Text, Text4.Text, _
                  Text5.Text, Text6.Text, Text7.Text, stToName, stToMail, bolSMTPFail, _
                  stCCToName, stCCToMail, m_ReplyTo, m_BCCto) Then
                  
'Removed by Morgan 2013/5/3 訊息改統一在 SendMail 彈,並加提供重試選擇
'                  'MsgBox "發送失敗！", vbCritical
'                  'Add by Morgan 2007/12/20 加判斷是否顯示錯誤訊息
'                  If bolMailFailNoAlert = False Then
'                     MsgBox "發 Mail 錯誤！請馬上電話連絡相關人員！" & vbCrLf & vbCrLf & "收件人：" & Text4.Text & vbCrLf & vbCrLf & "主旨：" & Text6.Text & vbCrLf & vbCrLf & "內文：" & Text7.Text, vbCritical, "寄信失敗！" & m_ErrStat
'                  End If
'                  Unload frmpic002
'                  Exit Sub
'
'               Else
'                  'Add by Morgan 2009/7/16 北所到北所郵件若失敗會改寄到防火牆
'                  If bolSMTPFail Then
'                     MsgBox "發 Mail 錯誤！郵件會延遲送達，請先電話通知相關人員！" & vbCrLf & vbCrLf & "收件人：" & Text4.Text & vbCrLf & vbCrLf & "主旨：" & Text6.Text & vbCrLf & vbCrLf & "內文：" & Text7.Text, vbCritical, "寄信失敗！" & m_ErrStat
'                  End If
                  Exit For
'end 2013/5/3

               End If
               Exit For 'Added by Morgan 2013/5/20 多收件人(含CC)改只要寄一次
            End If
         Next MailCnt
         
'Removed by Morgan 2013/5/20 併到上面程式作(正常都應該會有收件人)
'
'         '寄副本
'         ArrMail = Split(stCCToMail, ";")
'         ArrName = Split(stCCToName, ";")
'         For MailCnt = 0 To UBound(ArrMail)
'            If ArrMail(MailCnt) <> "" Then
'               Text4.Text = ArrName(MailCnt)
'               Text5.Text = ArrMail(MailCnt)
'               If Not SendMail(Text1.Text, Text2.Text, Text3.Text, Text4.Text, _
'                  Text5.Text, Text6.Text, Text7.Text, stToName, stToMail, bolSMTPFail, _
'                  stCCToName, stCCToMail, m_ReplyTo) Then
'
''Removed by Morgan 2013/5/3 訊息改統一在 SendMail 彈,並加提供重試選擇
''                  'MsgBox "發送失敗！", vbCritical
''                  'Add by Morgan 2007/12/20 加判斷是否顯示錯誤訊息
''                  If bolMailFailNoAlert = False Then
''                     MsgBox "發 Mail 錯誤！" & vbCrLf & "收件人：" & Text4.Text & vbCrLf & "主旨：" & Text6.Text & vbCrLf & "內文：" & Text7.Text & vbCrLf & "請馬上電話連絡相關人員！", vbCritical, "寄信失敗！" & m_ErrStat
''                  End If
''                  Unload frmpic002
''                  Exit Sub
''
''               Else
''                  'Add by Morgan 2009/7/16 北所到北所郵件若失敗會改寄到防火牆
''                  If bolSMTPFail Then
''                     MsgBox "發 Mail 錯誤！" & vbCrLf & vbCrLf & "收件人：" & Text4.Text & vbCrLf & "主旨：" & Text6.Text & vbCrLf & "內文：" & Text7.Text & vbCrLf & vbCrLf & "郵件會延遲送達，請先電話通知相關人員！", vbCritical, "寄信失敗！" & m_ErrStat
''                  End If
'                  Exit Sub
''end 2013/5/3
'
'               End If
'            End If
'         Next MailCnt
'
'end 2013/5/20
         
'         ArrMail = Split(txtEmail(0), ";")
'         For MailCnt = 0 To UBound(ArrMail)
'            If ArrMail(MailCnt) <> "" Then
'               'add by nickc 2006/08/18 檢查離職者發該區最小編號
'               ArrMail(MailCnt) = ChkMailId(ArrMail(MailCnt))
'               Text4.Text = GetPrjSalesNM(Trim(ArrMail(MailCnt)))
'               Text5.Text = Trim(ArrMail(MailCnt))
'               If SendMail(Text1.Text, Text2.Text, Text3.Text, Text4.Text, _
'                  Text5.Text, Text6.Text, Text7.Text, bolSMTPFail) Then
'                  'Add by Morgan 2009/7/16 北所到北所郵件若失敗會改寄到防火牆
'                  If bolSMTPFail Then
'                     MsgBox "發 Mail 錯誤！" & vbCrLf & vbCrLf & "收件人：" & Text4.Text & vbCrLf & "主旨：" & Text6.Text & vbCrLf & "內文：" & Text7.Text & vbCrLf & vbCrLf & "郵件會延遲送達，請先電話通知相關人員！", vbCritical, "寄信失敗！"
'                  End If
'               Else
'                  'MsgBox "發送失敗！", vbCritical
'                  'Add by Morgan 2007/12/20 加判斷是否顯示錯誤訊息
'                  If bolMailFailNoAlert = False Then
'                     MsgBox "發 Mail 錯誤！" & vbCrLf & "收件人：" & Text4.Text & vbCrLf & "主旨：" & Text6.Text & vbCrLf & "內文：" & Text7.Text & vbCrLf & "請馬上電話連絡相關人員！", vbCritical, "寄信失敗！"
'                  End If
'                  Unload frmpic002
'                  Exit Sub
'               End If
'            End If
'         Next MailCnt
         
  Else
    'MsgBox "發生錯誤！" & Err.Description, vbCritical
    '若是發信中，則一直等待，並秀出等待的訊息 edit by nickc 2007/12/14
    cmdok_Click 0
    Exit Sub
  End If
  
  bolLeave = True
Else
   bolLeave = False
End If

'Added by Lydia 2016/09/02 CFP調整點數時並超過點數上限發E-mail給主管批示
'Modified by Lydia 2017/12/20 +命名作業的退回通知 Or txtEmail(0).Tag = "frm090902_2"
If txtEmail(0).Tag = "CFPdg" Or txtEmail(0).Tag = "frm090902_2" Then
     Pub_Send_CFPdg = bolLeave
End If
'end 2016/09/02

'Add By Sindy 2018/7/9 提醒智權人員E-Mail留下歷程記錄
If bolMailSendOk = True And bolCCList = True And m_EEP01 <> "" Then
   '取得最大序號
   intMaxEEP02 = 0
   strSql = "select eep02 From empelectronprocess where eep01='" & m_EEP01 & "' order by eep02 desc"
   intI = 1
   CheckOC3
   Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      AdoRecordSet3.MoveFirst
      If AdoRecordSet3.RecordCount > 0 Then
         intMaxEEP02 = AdoRecordSet3.Fields(0)
      End If
   End If
   strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08,eep10) values(" & _
            CNULL(m_EEP01) & "," & intMaxEEP02 + 1 & ",'" & strUserNum & "'," & _
            CNULL(EMP_聯絡) & "," & _
            CNULL(txtEmail(0)) & "," & _
            strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & "," & CNULL(ChgSQL(txtEmail(2))) & "," & CNULL(txtEmail(5)) & ")"
   cnnConnection.Execute strSql
End If

Unload Me
Exit Sub
ErrHand:

'Modified by Morgan 2017/10/5 發生錯誤要 Unload 否則後面再呼叫 PUB_SendMail 都會錯
'ErrorMsg
If Err.NUMBER <> 0 Then
   MsgBox "錯誤號碼：" & Err.NUMBER & vbCrLf & "錯誤敘述：" & Err.Description, vbCritical, "寄信失敗！"
   Unload Me
End If
'end 2017/10/5
End Sub

'edit by nickc 2005/03/17 改用 winsock
'Sub SetAdress(ByVal strAdress As String)
'Dim varAdressTemp As Variant, i As Integer
'
'on error GoTo Err
'If strAdress = "" Then
'    Exit Sub
'End If
'varAdressTemp = Split(strAdress, ";")
'For i = 0 To UBound(varAdressTemp) - IIf(Right(strAdress, 1) = ";", 1, 0)
'       MAPIMessages.RecipIndex = i
'       MAPIMessages.RecipDisplayName = varAdressTemp(i)
'       MAPIMessages.ResolveName
'Next
'Exit Sub
'Err:
'ErrorMsg
'End Sub

'Modify By Cheng 2003/08/06
'Private Sub Form_Activate()
Public Sub Form_Activate()
Dim varSaveCursor

varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
MoveFormToCenter Me


On Error GoTo ErrHand

   'Add by Morgan 2005/3/9 控制不彈LoginUI
'edit by nickc 2005/03/17 改用 winsock
'   MAPISession.UserName = strUserNum
'   MAPISession.LogonUI = False
   '2005/3/9
'edit by nickc 2005/03/17 改用 winsock
'MAPISession.SignOn
'MAPIMessages.SessionID = MAPISession.SessionID

'Added by Lydia 2015/11/25 原"通訊錄"按鈕的程式會出錯,現在收件者改用勾選的方式
    Me.lstMailCC.Clear
    'Added by Lydia 2016/05/17 有預設收件人,設定副本收件人可勾選
    If bolCCList Then
        Label1.Caption = "副　本："
        Label7.Visible = True
        txtEmail(0).Left = lstMailCC.Left
        txtEmail(0).Top = 240
        GoTo JumpList
    Else
        If txtEmail(0).Text = "" Then
            txtEmail(0).Visible = False 'Added by Lydia 2015/12/01
JumpList:
            'Added by Lydia 2023/12/25
            If strSrvDate(1) >= 新部門啟用日 Then
               strSql = "SELECT nvl(a0923,a0902) as a0902,st01,st02 FROM staff,acc090,acc090new WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and st93=a0921(+) and substr(st01,4,1)<>'9' order by nvl(st93,st03),st01 asc"
            Else
            'end 2023/12/25
               strSql = "SELECT a0902,st01,st02 FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9' order by st03,st01 asc"
            End If
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               With RsTemp
                  RsTemp.MoveFirst
                  Do While RsTemp.EOF = False
                     lstMailCC.AddItem Trim(RsTemp.Fields("a0902")) & " " & Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
                     RsTemp.MoveNext
                  Loop
               End With
            End If
        'Modified by Lydia 2015/12/01
        Else
            lstMailCC.Visible = False: Label6.Visible = False
            txtEmail(0).Left = lstMailCC.Left
            txtEmail(0).Top = lstMailCC.Top
        End If
    End If
    'end 2016/05/17
'end 2015/11/25

Screen.MousePointer = varSaveCursor
Exit Sub
ErrHand:
Screen.MousePointer = varSaveCursor
ErrorMsg
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand 'Add By Sindy 2012/10/24

'add by nickc 2006/12/28
bolByMsa = False
'add by nickc 2007/12/14 加入秀訊息
Load frmpic002
frmpic002.Label1.Caption = "寄信中...請稍候..."
frmpic002.Show

If frmpic002.Visible Then 'Added by Morgan 2023/5/18
   frmpic002.ZOrder 0
End If

SetSMTP 'Added by Morgan 2013/3/26

'Add By Sindy 2012/10/24
Exit Sub
ErrHand:
   If Err.NUMBER = "401" Then '當有強制回應表單顯示時，無法再顯示非強制回應表單
      Resume Next
   End If
'2012/10/24 End
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHand
'edit by nickc 2005/03/17 改用 winsock
'MAPISession.SignOff
   'Add By Cheng 2002/07/18
   Unload frmpic002
   Set frm880005 = Nothing
   
Exit Sub
ErrHand:
ErrorMsg
   'Add By Cheng 2002/07/18
   Unload frmpic002
   Set frm880005 = Nothing
End Sub


Private Sub Timer1_Timer()
  Sec = Sec + 1
  DoEvents
End Sub

Private Sub txtEmail_GotFocus(Index As Integer)
txtEmail(Index).SelStart = 0
txtEmail(Index).SelLength = Len(txtEmail(0))
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
   Winsock1.GetData Result
   
   m_MailLog = m_MailLog & Format(Now, "hh:mm:ss") & " -> " & Result 'Added by Morgan 2013/10/18
End Sub

'Modify by Morgan 2009/7/15 +bolPrimarySmtpFail:先發的SMTP是否寄件失敗
'Modify By Sindy 2009/09/23 加副本
'Modify by Morgan 2011/4/22 +回覆信箱
'Modified by Lydia 2015/08/03 + 密件副本(BCC)
Private Function SendMail(SMTP$, FromName$, FromMail$, RcptName$, _
                          RcptMail$, Subj$, Body$, ByVal ToName As String, ByVal ToMail As String, Optional bolPrimarySmtpFail As Boolean, _
                          Optional ByVal CCToName As String = "", Optional ByVal CCToMail As String = "", Optional ByVal ReplyToMail As String = "", _
                          Optional ByVal BCCtoMail As String = "") As Boolean
   
   Dim bolByX400 As Boolean
   Dim ArrTmpName As Variant, ArrTmpMail As Variant, MailCnt As Integer 'Add By Sindy 2009/09/23
   Dim ArrRcpt 'Added by Morgan 2013/5/20
   Dim strErrMsg As String 'Added by Morgan 2013/9/4
   Dim iRetry As Integer 'Added by Morgan 2014/1/23
   Dim strSignFile As String, strSignHead As String, strSignBody As String, strSignAtt As String 'Added by Morgan 2019/8/27
   Dim bolExist As Boolean 'Added by Morgan 2023/12/29
   Dim strTo As String, strCC As String, strSendDate As String, strSendTime As String, strSql As String, intR As Integer 'Added by Morgan 2025/6/26
   
   m_MailSub = Subj 'Added by Morgan 2017/2/14
   m_ErrStat = "" 'Added by Morgan 2013/2/4
   
   bolPrimarySmtpFail = False
   
   'Added by Morgan 2015/10/5
   '剔除跳行符號，否則可能會導致郵件變亂碼
   ToMail = Replace(ToMail, vbCrLf, "")
   ToName = Replace(ToName, vbCrLf, "")
   CCToMail = Replace(CCToMail, vbCrLf, "")
   CCToName = Replace(CCToName, vbCrLf, "")
   BCCtoMail = Replace(BCCtoMail, vbCrLf, "")
   'end 2015/10/5
   
   
   'Removed by Morgan 2013/5/20
   ''add by nickc 2007/01/29 99999 不發
   ''Modified by Morgan 2012/9/5 +99997 也不要寄
   'If Replace(UCase(RcptMail), "@TAIE.COM.TW", "") = "99999" Or Replace(UCase(RcptMail), "@TAIE.COM.TW", "") = "99997" Then
   '      SendMail = True
   '      bolMailSendOk = True
   '      Exit Function
   'End If
   'end 2013/5/20
    
   'Add by Morgan 2006/6/2
   '如果是執行VB時再作一次確認
   'Modify by Morgan 2009/12/21 改判斷部門以便測試執行檔
   'If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
   'Modify By Sindy 2011/8/30 測試出缺勤簽核系統
   'Modify By Sindy 2015/5/19 連線測試資料庫時,也是彈訊息詢問
   'If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
   If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Then
   '2015/5/19 END
      'Modified by Morgan 2013/5/20 多收件人改一次寄送
      'If MsgBox("是否確定要發Mail給【" & IIf(RcptName$ <> "", RcptName$, RcptMail$) & "】？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      'Modify By Sindy 2015/9/8 增加顯示主旨及內容，供測試人員參考用
      'If MsgBox("是否確定要發Mail給【" & ToMail & IIf(Replace(ToName, ";", "") <> "", "(" & ToName & ")", "") & CCToMail & IIf(CCToName <> "", "(" & CCToName & ")", "") & "】？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      'Modified by Lydia 2022/03/11 改成可支援Unicode
      'If MsgBox("是否確定要發Mail給" & vbCrLf & vbCrLf & "【收件者：" & ToMail & IIf(Replace(ToName, ";", "") <> "", "(" & ToName & ")", "") & IIf(CCToMail <> "", vbCrLf & "　副　本：", "") & CCToMail & IIf(CCToName <> "", "(" & CCToName & ")", "") & "】？" & _
                vbCrLf & vbCrLf & "主旨：" & Subj$ & _
                vbCrLf & vbCrLf & "內容：" & vbCrLf & Body$, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      'Modified by Morgan 2025/5/9
      'strTEMP = "是否確定要發Mail給" & vbCrLf & vbCrLf & "【收件者：" & ToMail & IIf(Replace(ToName, ";", "") <> "", "(" & ToName & ")", "") & IIf(CCToMail <> "", vbCrLf & "　副　本：", "") & CCToMail & IIf(CCToName <> "", "(" & CCToName & ")", "") & "】？" & _
                vbCrLf & vbCrLf & "主旨：" & Subj$ & _
                vbCrLf & vbCrLf & "內容：" & vbCrLf & Body$
      'If UniMsgBox(strTEMP, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      'Added by Morgan 2025/6/18
      If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Then
         Subj$ = "<測試>" & Subj$
      End If
      'end 2025/6/18
      strTemp = "【主　旨】：" & Subj$ & _
               vbCrLf & vbCrLf & "【收件者】：" & ToMail & IIf(Replace(ToName, ";", "") <> "", "(" & ToName & ")", "") & IIf(CCToMail <> "", vbCrLf & "【副　本】：", "") & CCToMail & IIf(CCToName <> "", "(" & CCToName & ")", "") & _
               vbCrLf & vbCrLf & "【內　容】：" & vbCrLf & Body$
               
      'Modified by Morgan 2025/6/17 改電腦中心或用VB才詢問，其他都只顯示訊息但不寄發，以免給User測試時誤寄出而衍生後續問題
      If UCase(pub_DbTerminalName) <> "ORATEST" And (InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or Pub_StrUserSt03 = "M51") Then
         strTemp = "是否確定要發以下EMail？" & vbCrLf & vbCrLf & strTemp
         If UniMsgBox(strTemp, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            SendMail = True
            bolMailSendOk = True
            Exit Function
         End If
      Else
         strTemp = "原本系統會發以下EMail，因目前為測試，故不會真實寄發。" & vbCrLf & vbCrLf & strTemp
         Call UniMsgBox(strTemp)
         SendMail = True
         bolMailSendOk = True
         Exit Function
      End If
      'end 2025/6/17
   End If
   'end 2006/6/2
  
   Dim MAIL$, outTO$, outFR$
   
    If Mailing = True Then Exit Function
    Mailing = True
    
    'add by nickc 2006/12/06 加入可以崁圖
    'Modified by Morgan 2019/11/21 sCharset,sCTEnc 要共用改全域常數並更名為 cCharset,cCTEnc
    'Dim iType As Long, i As Long, sID As String, sCharset As String, sCTEnc As String, tmpMailStr As String, outSF As String, sDat
    'sCharset = "charset=" & Chr$(34) & "Big5" & Chr$(34) & vbCrLf
    'sCTEnc = "Content-Transfer-Encoding: quoted-printable" & vbCrLf
    Dim iType As Long, i As Long, sID As String, sPicBody As String, tmpMailStr As String, outSF As String, sDat
    'end 2019/11/21
    
    Screen.MousePointer = vbHourglass

    If Winsock1.State = sckClosed Then
      'add by nickc 2007/05/22 紀錄錯誤
      
      On Error GoTo ERRORMail
      Winsock1.LocalPort = 0
      'Reply-To
      'edit by nickc 2006/10/13 用 x400 退信會無法退
      'outFR = "mail from: " & MailBefore & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & MailAfter & vbCrLf
      
      'Modify by Morgan 2011/8/19 帳號不要轉大寫否則對方白名單可能會無效
      'outFR = "mail from: " & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & MailAfter & vbCrLf
      If InStr(UCase(FromMail), "@TAIE.COM.TW") > 0 Then
         outFR = "mail from: " & Left(FromMail, InStr(UCase(FromMail), "@TAIE.COM.TW") - 1) & MailAfter & vbCrLf
      Else
         outFR = "mail from: " & FromMail & MailAfter & vbCrLf
      End If
      'end 2011/8/19

      'Modify by Morgan 2006/12/8 對外的信不用
      'outTO = "rcpt to: " & MailBefore & Replace(UCase(RcptMail), "@TAIE.COM.TW", "") & MailAfter & vbCrLf '& "data" & vbCrLf
      If InStr(RcptMail, "@") > 0 Then
         outTO = "rcpt to: " & RcptMail & vbCrLf  '& "data" & vbCrLf
      Else
         'Modified by Morgan 2012/2/24 統一都建立對外帳號(設備不支援X400)
         ''add by nickc 2007/01/23 北所寄北所，且收件者部門為 F 開頭
         'If PUB_GetST06(Replace(UCase(FromMail), "@TAIE.COM.TW", "")) = "1" And PUB_GetST06(Replace(UCase(RcptMail), "@TAIE.COM.TW", "")) = "1" And UCase(Mid(PUB_GetST03(Replace(UCase(RcptMail), "@TAIE.COM.TW", "")), 1, 1)) = "F" Then
         '   outTO = "rcpt to: " & MailBeforeF & Replace(UCase(RcptMail), "@TAIE.COM.TW", "") & MailAfter & vbCrLf '& "data" & vbCrLf
         '   bolByX400 = True 'Add by Morgan 2009/7/17
         'Else
            outTO = "rcpt to: " & MailBefore & Replace(UCase(RcptMail), "@TAIE.COM.TW", "") & MailAfter & vbCrLf '& "data" & vbCrLf
         'End If
         'End 2012/2/24
      End If
      'end 2006/12/8
            
      MAIL = ""
'edit by nickc 2006/12/06  加入可以崁圖
'      MAIL = MAIL & "From: """ & FromName & """ <" & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & "@taie.com.tw>"    ' <" & MailBefore & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & MailAfter & ">"
'      MAIL = MAIL & vbCrLf & "Date: " & Format(Date, "Ddd")
'      MAIL = MAIL & ", " & Format(Date, "dd Mmm YYYY") & " "
'      MAIL = MAIL & Format(time, "hh:mm:ss") & " +0800" & vbCrLf
'      MAIL = MAIL & "X-Mailer: Taie "
'      MAIL = MAIL & vbCrLf & "To: """ & RcptName & """ <" & Replace(UCase(RcptMail), "@TAIE.COM.TW", "") & "@taie.com.tw>"  '<" & MailBefore & Replace(UCase(RcptMail), "@TAIE.COM.TW", "") & MailAfter & ">"
'      'edit by nickc 2006/02/24 秀玲說要區別系統發的 mail
'      'MAIL = MAIL & vbCrLf & "Subject: " & Subj & vbCrLf
'      MAIL = MAIL & vbCrLf & "Subject: ◎系統代發◎" & Subj & vbCrLf
'      MAIL = MAIL & vbCrLf & Body & vbCrLf & vbCrLf & "." & vbCrLf
      sDat = GetGUID()   '-- create a messageID using a GUID.
      '郵件 id
      MAIL = MAIL & "Message-ID: <" & sDat & "@Exchange>" & vbCrLf
      
      '郵件來源
      'Modified by Morgan 2022/3/23 +UTF8
      'tmpMailStr = ConvertToBase64(FromName, False, False)
      tmpMailStr = ConvertToBase64(FromName, False, False, m_ByUTF8)
      'end 2022/3/23
      If tmpMailStr <> FromName And FromName <> "" Then
         'Modified by Morgan 2022/3/23 預設編碼改UTF8格式
         'MAIL = MAIL & "From: =?Big5?B?" & tmpMailStr & "?= <" & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
         If m_ByUTF8 Then
            MAIL = MAIL & "From: =?utf-8?B?" & tmpMailStr & "?= <" & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
         Else
            MAIL = MAIL & "From: =?Big5?B?" & tmpMailStr & "?= <" & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
         End If
         'end 2022/3/23
      Else
         MAIL = MAIL & "From: """ & FromName & """ <" & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
         
      End If
      
      '郵件目的
'      tmpMailStr = ConvertToBase64(RcptName, False, False)
'      If tmpMailStr <> RcptName And RcptName <> "" Then
'        MAIL = MAIL & "To: =?Big5?B?" & tmpMailStr & "?= <" & Replace(UCase(RcptMail), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
'        'MAIL = MAIL & "To: =?Big5?B?" & tmpMailStr & "?= <" & RcptMail & ">" & vbCrLf
'      Else
'         'Modify by Morgan 2006/12/8
'         'MAIL = MAIL & "To: """ & RcptName & """ <" & Replace(UCase(RcptMail), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
'         If InStr(UCase(RcptMail), "@TAIE.COM.TW") > 0 Then
'            MAIL = MAIL & "To: """ & RcptName & """ <" & Replace(UCase(RcptMail), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
'         Else
'            MAIL = MAIL & "To: """ & RcptName & """ <" & RcptMail & ">" & vbCrLf
'         End If
'        'MAIL = MAIL & "To: """ & RcptName & """ <" & RcptMail & ">" & vbCrLf
'      End If
      
      'Modify By Sindy 2009/09/23 組收件者資訊
      ArrTmpMail = Split(ToMail, ";")
      If ToName = "" Then ToName = ";" 'Add by Morgan 2009/9/30
      ArrTmpName = Split(ToName, ";")
      bolExist = False 'Added by Morgan 2023/12/29
      For MailCnt = 0 To UBound(ArrTmpMail)
         If ArrTmpMail(MailCnt) <> "" Then
            'Added by Morgan 2025/6/26
            strTo = strTo & ArrTmpMail(MailCnt)
            If ArrTmpName(MailCnt) <> "" Then
               strTo = strTo & " (" & ArrTmpName(MailCnt) & ")"
            End If
            strTo = strTo & ";"
            'end 2025/6/26
            
            If InStr(ArrTmpMail(MailCnt), "@") = 0 Then ArrTmpMail(MailCnt) = ArrTmpMail(MailCnt) & "@taie.com.tw" 'Added by Morgan 2024/1/10 非員工信箱(ex. backup)要補domain，否則outlook可能會顯示不正確
            'Modified by Morgan 2022/3/23 +UTF8格式
            'tmpMailStr = ConvertToBase64(CStr(ArrTmpName(MailCnt)), False, False)
            tmpMailStr = ConvertToBase64(CStr(ArrTmpName(MailCnt)), False, False, m_ByUTF8)
            'end 2022/3/23
            'Modified by Morgan 2023/12/29 多收件人改用TAB+逗號串接(Header重複可能會被退信)
            If tmpMailStr <> ArrTmpName(MailCnt) And ArrTmpName(MailCnt) <> "" Then
               'Modified by Morgan 2022/3/23 +UTF8格式
               'MAIL = MAIL & "To: =?Big5?B?" & tmpMailStr & "?= <" & Replace(UCase(ArrTmpMail(MailCnt)), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
               If m_ByUTF8 Then
                  MAIL = MAIL & IIf(bolExist, vbTab & ",", "To:") & " =?utf-8?B?" & tmpMailStr & "?= <" & Replace(UCase(ArrTmpMail(MailCnt)), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
               Else
                  MAIL = MAIL & IIf(bolExist, vbTab & ",", "To:") & " =?Big5?B?" & tmpMailStr & "?= <" & Replace(UCase(ArrTmpMail(MailCnt)), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
               End If
               'end 2022/3/23
            Else
               If InStr(UCase(ArrTmpMail(MailCnt)), "@TAIE.COM.TW") > 0 Then
                  MAIL = MAIL & IIf(bolExist, vbTab & ",", "To:") & " """ & ArrTmpName(MailCnt) & """ <" & Replace(UCase(ArrTmpMail(MailCnt)), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
               Else
                  MAIL = MAIL & IIf(bolExist, vbTab & ",", "To:") & " """ & ArrTmpName(MailCnt) & """ <" & ArrTmpMail(MailCnt) & ">" & vbCrLf
               End If
            End If
            bolExist = True 'Added by Morgan 2023/12/29
         End If
      Next MailCnt
      
      'Add By Sindy 2009/09/23 組副本資訊
      bolExist = False 'Added by Morgan 2023/12/29
      If CCToMail <> "" Then
         ArrTmpMail = Split(CCToMail, ";")
         If CCToName = "" Then CCToName = ";" 'Add by Morgan 2009/9/30
         ArrTmpName = Split(CCToName, ";")
         For MailCnt = 0 To UBound(ArrTmpMail)
            If ArrTmpMail(MailCnt) <> "" Then
               'Added by Morgan 2025/6/26
               strCC = strCC & ArrTmpMail(MailCnt)
               If ArrTmpName(MailCnt) <> "" Then
                  strCC = strCC & " (" & ArrTmpName(MailCnt) & ")"
               End If
               strCC = strCC & ";"
               'end 2025/6/26

               If InStr(ArrTmpMail(MailCnt), "@") = 0 Then ArrTmpMail(MailCnt) = ArrTmpMail(MailCnt) & "@taie.com.tw" 'Added by Morgan 2024/1/10 非員工信箱(ex. backup)要補domain，否則outlook可能會顯示不正確
               'Modified by Morgan 2022/3/23 +UTF8格式
               'tmpMailStr = ConvertToBase64(CStr(ArrTmpName(MailCnt)), False, False)
               tmpMailStr = ConvertToBase64(CStr(ArrTmpName(MailCnt)), False, False, m_ByUTF8)
               'end 2022/3/23
               'Modified by Morgan 2023/12/29 多收件人改用TAB+逗號串接(Header重複可能會被退信)
               If tmpMailStr <> ArrTmpName(MailCnt) And ArrTmpName(MailCnt) <> "" Then
                  'Modified by Morgan 2022/3/23 +UTF8格式
                  'MAIL = MAIL & "CC: =?Big5?B?" & tmpMailStr & "?= <" & Replace(UCase(ArrTmpMail(MailCnt)), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
                  If m_ByUTF8 Then
                     MAIL = MAIL & IIf(bolExist, vbTab & ",", "CC:") & " =?utf-8?B?" & tmpMailStr & "?= <" & Replace(UCase(ArrTmpMail(MailCnt)), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
                  Else
                     MAIL = MAIL & IIf(bolExist, vbTab & ",", "CC:") & " =?Big5?B?" & tmpMailStr & "?= <" & Replace(UCase(ArrTmpMail(MailCnt)), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
                  End If
                  'end 2022/3/23
               Else
                  If InStr(UCase(ArrTmpMail(MailCnt)), "@TAIE.COM.TW") > 0 Then
                     MAIL = MAIL & IIf(bolExist, vbTab & ",", "CC:") & " """ & ArrTmpName(MailCnt) & """ <" & Replace(UCase(ArrTmpMail(MailCnt)), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
                  Else
                     MAIL = MAIL & IIf(bolExist, vbTab & ",", "CC:") & " """ & ArrTmpName(MailCnt) & """ <" & ArrTmpMail(MailCnt) & ">" & vbCrLf
                  End If
               End If
               bolExist = True 'Added by Morgan 2023/12/29
            End If
         Next MailCnt
      End If
      
      'Modified by Lydia 2015/08/03 +密件副本
      'Removed by Morgan 2023/12/29 內容不該有BCC否則可能會在郵件原始碼內看見
      'If BCCtoMail <> "" Then
      '   ArrTmpMail = Split(BCCtoMail, ";")
      '   For MailCnt = 0 To UBound(ArrTmpMail)
      '      If ArrTmpMail(MailCnt) <> "" Then
      '          If InStr(UCase(ArrTmpMail(MailCnt)), "@TAIE.COM.TW") > 0 Then
      '             MAIL = MAIL & "BCC:" & " " & Replace(UCase(ArrTmpMail(MailCnt)), "@TAIE.COM.TW", "") & "@taie.com.tw" & vbCrLf
      '          Else
      '             MAIL = MAIL & "BCC:" & " " & ArrTmpMail(MailCnt) & "@taie.com.tw" & vbCrLf
      '          End If
      '      End If
      '   Next MailCnt
      'End If
      'end 2023/12/29
      
      '主旨
      'Modify by Morgan 2006/12/8 對外的信不用
      'tmpMailStr = ConvertToBase64("◎系統代發◎" & Subj, False, False)
      If InStr(UCase(outTO), "@TAIE.COM.TW") > 0 Then
         'Modify By Sindy 2014/1/16
         'tmpMailStr = ConvertToBase64("◎系統代發◎" & Subj, False, False)
         'Modified by Morgan 2022/3/23 +UTF8格式
         'tmpMailStr = ConvertToBase64("◎" & IIf(Pub_StrUserSt03 = "M51" And PUB_GetST05(strUserNum) <> "" And UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱), PUB_GetDbTerminal, "") & Subj, False, False)
         tmpMailStr = ConvertToBase64("◎" & IIf(Pub_StrUserSt03 = "M51" And PUB_GetST05(strUserNum) <> "" And UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱), PUB_GetDbTerminal, "") & Subj, False, False, m_ByUTF8)
         'end 2022/3/23
         '2014/1/16 END
      Else
         'Modified by Morgan 2022/3/23 +UTF8格式
         'tmpMailStr = ConvertToBase64(Subj, False, False)
         tmpMailStr = ConvertToBase64(Subj, False, False, m_ByUTF8)
         'end 2022/3/23
      End If
      'end 2006/12/8
      
      If tmpMailStr <> Subj Then
        'Modified by Morgan 2022/3/23 +UTF8格式
        'MAIL = MAIL & "Subject: =?Big5?B?" & tmpMailStr & "?=" & vbCrLf
        If m_ByUTF8 Then
            MAIL = MAIL & "Subject: =?utf-8?B?" & tmpMailStr & "?=" & vbCrLf
        Else
            MAIL = MAIL & "Subject: =?Big5?B?" & tmpMailStr & "?=" & vbCrLf
        End If
        'end 2022/3/23
      Else
        MAIL = MAIL & "Subject: " & Subj & vbCrLf
      End If
      '時間
      MAIL = MAIL & "Date: " & Format(Date, "Ddd")
      MAIL = MAIL & ", " & Format(Date, "dd Mmm YYYY") & " "
      MAIL = MAIL & Format(time, "hh:mm:ss") & " +0800" & vbCrLf
      
      'Modified by Morgan 2012/6/15 財務處對外信不要標示重要(目前只有財務處會發對外信)-- 婧瑄
      'Modified by Morgan 2018/9/12 +bolImportant
      If InStr(UCase(outTO), "@TAIE.COM.TW") > 0 Or bolImportant Then
         '重要性
         MAIL = MAIL & "Importance: high" & vbCrLf
         '優先
         MAIL = MAIL & "X-Priority: 1" & vbCrLf
      End If
      
      '按回覆用的收件者
      'Modified by Morgan 2022/3/23 +UTF8格式
      'tmpMailStr = ConvertToBase64(FromName, False, False)
      tmpMailStr = ConvertToBase64(FromName, False, False, m_ByUTF8)
      'end 2022/3/23
      
      'Add by Morgan 2011/4/22 若有指定回覆信箱時
      If ReplyToMail <> "" Then
         'Modified by Morgan 2022/3/23 +UTF8格式
         'MAIL = MAIL & "Reply-To: =?Big5?B?" & tmpMailStr & "?= <" & MailBefore & Replace(UCase(ReplyToMail), "@TAIE.COM.TW", "") & MailAfter & ">" & vbCrLf
         If m_ByUTF8 Then
            MAIL = MAIL & "Reply-To: =?utf-8?B?" & tmpMailStr & "?= <" & MailBefore & Replace(UCase(ReplyToMail), "@TAIE.COM.TW", "") & MailAfter & ">" & vbCrLf
         Else
            MAIL = MAIL & "Reply-To: =?Big5?B?" & tmpMailStr & "?= <" & MailBefore & Replace(UCase(ReplyToMail), "@TAIE.COM.TW", "") & MailAfter & ">" & vbCrLf
         End If
         'end 2022/3/23
      'end 2011/4/22
      'Added by Morgan 2018/8/20
      '因用QPGMR寄信會被Outlook規則歸類為垃圾信(信箱不存在),故經理將信箱實體設為administrator
      '為避免誤回覆,故意將回覆信箱設定為noreply@taie.com.tw(不存在),這樣 User 就會收到退信知道不能回覆
      ElseIf UCase(FromMail) = "QPGMR" Then
         MAIL = MAIL & "Reply-To: <noreply@taie.com.tw>" & vbCrLf
      'end 2018/8/20
      Else
         If tmpMailStr <> FromName Then
            'Modified by Morgan 2022/3/23 +UTF8格式
            'MAIL = MAIL & "Reply-To: =?Big5?B?" & tmpMailStr & "?= <" & MailBefore & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & MailAfter & ">" & vbCrLf
            If m_ByUTF8 Then
               MAIL = MAIL & "Reply-To: =?utf-8?B?" & tmpMailStr & "?= <" & MailBefore & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & MailAfter & ">" & vbCrLf
            Else
               MAIL = MAIL & "Reply-To: =?Big5?B?" & tmpMailStr & "?= <" & MailBefore & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & MailAfter & ">" & vbCrLf
            End If
            'end 2022/3/23
         Else
            'add by nickc 2007/01/23 北所寄北所，且收件者部門為 F 開頭
            If PUB_GetST06(Replace(UCase(FromMail), "@TAIE.COM.TW", "")) = "1" And PUB_GetST06(Replace(UCase(RcptMail), "@TAIE.COM.TW", "")) = "1" And UCase(Mid(PUB_GetST03(Replace(UCase(RcptMail), "@TAIE.COM.TW", "")), 1, 1)) = "F" Then
               MAIL = MAIL & "Reply-To: """ & FromName & """ <" & MailBeforeF & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & MailAfter & ">" & vbCrLf
            Else
               MAIL = MAIL & "Reply-To: """ & FromName & """ <" & MailBefore & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & MailAfter & ">" & vbCrLf
            End If
         End If
      End If 'Add by Morgan 2011/4/22
      
      'Added by Morgan 2021/4/1
      '讀取回條(會因收件方郵件軟體而不同,Outlook 要用 Disposition-Notification-To)
      If bolGetReceipt = True Then
         'Modified by Morgan 2022/3/23 +UTF8格式
         'tmpMailStr = ConvertToBase64(FromName, False, False)
         tmpMailStr = ConvertToBase64(FromName, False, False, m_ByUTF8)
         'end 2022/3/23
         If tmpMailStr <> FromName Then
            'Modified by Morgan 2022/3/23 +UTF8格式
            'MAIL = MAIL & "Return-Receipt-To: =?Big5?B?" & tmpMailStr & "?= <" & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
            'MAIL = MAIL & "Disposition-Notification-To: =?Big5?B?" & tmpMailStr & "?= <" & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
            If m_ByUTF8 Then
               MAIL = MAIL & "Return-Receipt-To: =?utf-8?B?" & tmpMailStr & "?= <" & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
               MAIL = MAIL & "Disposition-Notification-To: =?utf-8?B?" & tmpMailStr & "?= <" & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
            Else
               MAIL = MAIL & "Return-Receipt-To: =?Big5?B?" & tmpMailStr & "?= <" & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
               MAIL = MAIL & "Disposition-Notification-To: =?Big5?B?" & tmpMailStr & "?= <" & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
            End If
            'end 2022/3/23
         Else
            MAIL = MAIL & "Return-Receipt-To: """ & FromName & """ <" & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
            MAIL = MAIL & "Disposition-Notification-To: """ & FromName & """ <" & Replace(UCase(FromMail), "@TAIE.COM.TW", "") & "@taie.com.tw>" & vbCrLf
         End If
      End If
      'end 2021/4/1

      '郵件編碼設定
      MAIL = MAIL & "MIME-Version: 1.0" & vbCrLf
      sID = GetGUID()
      'Modified by Morgan 2017/11/2
      'PrepareEmail Body$, Me.txtEmail(4).Text, sID, IIf(chkShowPic.Value = 1, False, True)
      'Modified by Morgan 2017/11/20
      'PrepareEmail Body$, Me.txtEmail(4).Text, sID, Not bolHTML
      'Modified by Morgan 2022/3/23 +m_ByUTF8
      PrepareEmail Body$, Me.txtEmail(4).Text, sID, IIf(chkShowPic.Value = 1, False, True), bolHTML, m_ByUTF8
      'end 2017/11/20
      'end 2017/11/2
  '-- email type: 0-plain. 1-html. 2-plain with attachments. 3-html with attachments.
  
'Added by Morgan 2019/8/27 有簽名檔時格式固定用 3
If m_iSignatureID > 0 Then
   iType = 3
Else
'end 2019/8/27

       iType = 0
       If (UBound(APics) > 0) Then iType = iType + 2
       'If (BooPlain = False) Then iType = iType + 1
       'Modified by Morgan 2017/11/2
       'If (chkShowPic.Value = 1) Then iType = iType + 1
       If bolHTML Then iType = iType + 1
       'end 2017/11/2
       
End If 'Added by Morgan 2019/8/27

         Select Case iType
   '------------- send plain text email ------------------------------------
           Case 0   '--plain text.
'                               Content-Type: text/plain;
'                  Charset = "iso-8859-1"
'                Content-Transfer-Encoding: 7bit
'                x -Mailer:
                MAIL = MAIL & "Content-Type: text/plain;" & vbCrLf
                
                'Modify by Morgan 2011/5/5 純文字並沒有編碼,加編碼敘述反而會造成資料顯示不正確
                'MAIL = MAIL & vbTab & cCharset & cCTEnc & vbCrLf
                'Modified by Morgan 2022/3/23 +UTF8格式
                'MAIL = MAIL & vbTab & cCharset & vbCrLf
                If m_ByUTF8 Then
                     MAIL = MAIL & vbTab & "charset=""utf-8""" & vbCrLf & "Content-Transfer-Encoding: base64" & vbCrLf & vbCrLf
                Else
                     MAIL = MAIL & vbTab & cCharset & vbCrLf
                End If
                'end 2022/3/23
                'end 2011/5/5
                
'                MAIL = MAIL & "X-Mailer: Exchange" & vbCrLf & vbCrLf
                 ' LRet = SendData(sDat)
                  'LRet = SendData(sBodyText & vbCrLf & cDOT & vbCrLf)
                  
                'Modified by Morgan 2022/3/23 +UTF8格式
                'MAIL = MAIL & sBodyText & vbCrLf & cDOT & vbCrL
                If m_ByUTF8 Then
                  MAIL = MAIL & ConvertToBase64(sBodyText, False, True, True) & vbCrLf & cDOT & vbCrLf
                Else
                  MAIL = MAIL & sBodyText & vbCrLf & cDOT & vbCrLf
                End If
                'end 2022/3/23
   '----------- send HTML email, no pictures -----------------------------
           Case 1   '-- html, no attchments.
             '-- send the rest of main header and text header.
                MAIL = MAIL & "Content-Type: multipart/alternative;" & vbCrLf
                MAIL = MAIL & vbTab & "boundary=" & Chr$(34) & cBoundaryA & Chr$(34) & vbCrLf   'Boundary_A_3435FE2_6617A_AA"
                MAIL = MAIL & "X-Mailer: Taie" & vbCrLf & vbCrLf
                '  LRet = SendData(sDat)
                MAIL = MAIL & cDASH2 & cBoundaryA & vbCrLf & "Content-Type: text/plain;" & vbCrLf
                'Modified by Morgan 2022/3/23 +UTF8格式
                'MAIL = MAIL & vbTab & cCharset & cCTEnc & vbCrLf
                If m_ByUTF8 Then
                     MAIL = MAIL & vbTab & "charset=""utf-8""" & vbCrLf & "Content-Transfer-Encoding: base64" & vbCrLf & vbCrLf
                Else
                     MAIL = MAIL & vbTab & cCharset & cCTEnc & vbCrLf
                End If
                'end 2022/3/23
                  'LRet = SendData(sDat) '--text header sent. text data next.
                  
             '-- send text blurb for non-HTML email readers:
                'Modified by Morgan 2022/3/23 +UTF8格式
                'MAIL = MAIL & TextBlurb()
                If m_ByUTF8 Then
                     MAIL = MAIL & ConvertToBase64(TextBlurb(), False, True, True) & vbCrLf & vbCrLf
                Else
                     MAIL = MAIL & TextBlurb()
                End If
                'end 2022/3/23
                
                   'LRet = SendData(sDat) '-- "apparently your program can't read this...."
                   
            '-- send html section header:
                MAIL = MAIL & cDASH2 & cBoundaryA & vbCrLf & "Content-Type: text/html;" & vbCrLf
                'Modified by Morgan 2022/3/23 +UTF8格式
                'MAIL = MAIL & vbTab & cCharset & cCTEnc & vbCrLf
                If m_ByUTF8 Then
                     MAIL = MAIL & vbTab & "charset=""utf-8""" & vbCrLf & "Content-Transfer-Encoding: base64" & vbCrLf & vbCrLf
                Else
                     MAIL = MAIL & vbTab & cCharset & cCTEnc & vbCrLf
                End If
                'end 2022/3/23
                  'LRet = SendData(sDat)
                  
            '-- send HTML message body:
                  'LRet = SendData(sBodyText & vbCrLf & vbCrLf)
                'Modified by Morgan 2020/10/22 +指定細明體,否則outlook預設的新細明體無法顯示空白
                'MAIL = MAIL & sBodyText & vbCrLf & vbCrLf
                'Modified by Morgan 2022/3/23 +UTF8格式
                'MAIL = MAIL & "<span style=3D""font-family:" & ConvertToQp("細明體") & "," & ConvertToQp("Arial") & """>" & vbCrLf & PUB_FixFirstDot(sBodyText) & vbCrLf & "</span>" & vbCrLf & vbCrLf
                If m_ByUTF8 Then
                     MAIL = MAIL & ConvertToBase64("<span style=""font-family:細明體,Arial"">" & vbCrLf & sBodyText & vbCrLf & "</span>", False, True, True) & vbCrLf & vbCrLf
                Else
                     MAIL = MAIL & "<span style=3D""font-family:" & ConvertToQp("細明體") & "," & ConvertToQp("Arial") & """>" & vbCrLf & PUB_FixFirstDot(sBodyText) & vbCrLf & "</span>" & vbCrLf & vbCrLf
                End If
                'end 2022/3/23
                'end 2020/10/22
                  
            '-- send END signal:
                  'LRet = SendData(cDASH2 & cBoundaryA & cDASH2 & vbCrLf & cDOT & vbCrLf)  ' Boundary_A_3435FE2_6617A_AA--
                MAIL = MAIL & cDASH2 & cBoundaryA & cDASH2 & vbCrLf & cDOT & vbCrLf
 '---------------- send email with plain text and attachments ----------------
           Case 2 'plain with atttachments.
             '-- send rest of main header:
                MAIL = MAIL & "Content-Type: multipart/mixed;" & vbCrLf
                MAIL = MAIL & vbTab & "boundary=" & Chr$(34) & cBoundaryA & Chr$(34) & vbCrLf
                MAIL = MAIL & "X-Mailer: Taie" & vbCrLf & vbCrLf
                   'LRet = SendData(sDat)
                   
              '-- send text part header:
                MAIL = MAIL & cDASH2 & cBoundaryA & vbCrLf & "Content-Type: text/plain;" & vbCrLf
                'Modified by Morgan 2022/3/23 +UTF8格式
                'MAIL = MAIL & vbTab & cCharset
                'MAIL = MAIL & "Content-Transfer-Encoding: 7bit" & vbCrLf & vbCrLf
                If m_ByUTF8 Then
                  MAIL = MAIL & vbTab & "charset=""utf-8""" & vbCrLf & "Content-Transfer-Encoding: base64" & vbCrLf & vbCrLf
                Else
                  MAIL = MAIL & vbTab & cCharset
                  MAIL = MAIL & "Content-Transfer-Encoding: 7bit" & vbCrLf & vbCrLf
                End If
                'end 2022/3/23
                   'LRet = SendData(sDat)
                   
              '-- send message body:
                    'LRet = SendData(sBodyText & vbCrLf & vbCrLf)
               'Modified by Morgan 2022/3/23 +UTF8格式
               'MAIL = MAIL & sBodyText & vbCrLf & vbCrLf
               If m_ByUTF8 Then
                  MAIL = MAIL & ConvertToBase64(sBodyText, False, False, True) & vbCrLf & vbCrLf
               Else
                  MAIL = MAIL & sBodyText & vbCrLf & vbCrLf
               End If
               'end 2022/3/23
              '-- do attachments:
                For i = 1 To UBound(APics)
                    'sDat = APics(i)           '-- get base64 encoded file.
                    'LRet = SendData(sDat)   '-- send base64 encoded file.
                    MAIL = MAIL & APics(i)
                Next
                
              '-- send END signal:
                  'LRet = SendData(cDASH2 & cBoundaryA & cDASH2 & vbCrLf & cDOT & vbCrLf)
               'Modified by Morgan 2017/11/10 修正大附件會觸發字串空間不足(#14)錯誤問題
               'MAIL = MAIL & cDASH2 & cBoundaryA & cDASH2 & vbCrLf & cDOT & vbCrLf
               MAIL = MAIL & cDASH2
               MAIL = MAIL & cBoundaryA
               MAIL = MAIL & cDASH2
               MAIL = MAIL & vbCrLf
               MAIL = MAIL & cDOT
               MAIL = MAIL & vbCrLf
               'end 2017/11/10
    '---------- send HTML email with pictures. ------------------------
           Case 3  '--html with pics.
             '--
                MAIL = MAIL & "Content-Type: multipart/mixed;" 'multipart/related;"
'                End If
                MAIL = MAIL & vbCrLf & vbTab & "boundary=" & Chr$(34) & cBoundaryA & Chr$(34) & cSEMIC & vbCrLf
                
                If chkShowPic.Value = 0 Then
                    MAIL = MAIL & vbTab & "type=" & Chr$(34) & "multipart/alternative" & Chr$(34) & vbCrLf
                End If
                MAIL = MAIL & "X-Mailer: Exchange" & vbCrLf & vbCrLf
                    'LRet = SendData(sDat)
                         
            '-- send main header (A - "multipart alternative") and plain text header (B)
               MAIL = MAIL & cDASH2 & cBoundaryA & vbCrLf
               
'Removed by Morgan 2019/11/21 預覽要用，改寫成函數
'               'Added by Morgan 2019/8/27
'               '注意:若加下面純文字內容時,OUTLOOK顯示會不正常,故有簽名時不帶
'               If m_iSignatureID > 0 Then
'                  MAIL = MAIL & "Content-Type: multipart/related;" & vbCrLf
'                  MAIL = MAIL & vbTab & "boundary=" & Chr$(34) & cBoundaryB & Chr$(34) & ";" & vbCrLf
'                  MAIL = MAIL & vbTab & "type=""text/html""" & vbCrLf & vbCrLf
'               Else
'               'end 2019/8/27
'
'                  MAIL = MAIL & "Content-Type: multipart/alternative;" & vbCrLf
'                  MAIL = MAIL & vbTab & "boundary=" & Chr$(34) & cBoundaryB & Chr$(34) & vbCrLf & vbCrLf
'                  MAIL = MAIL & cDASH2 & cBoundaryB & vbCrLf & "Content-Type: text/plain;" & vbCrLf
'                  MAIL = MAIL & vbTab & cCharset & cCTEnc & vbCrLf
'                  'LRet = SendData(sDat)
'
'            '-- send text blurb:
'                  MAIL = MAIL & TextBlurb()
'                  'LRet = SendData(sDat)
'
'               End If 'Added by Morgan 2019/8/27
'
'            '-- send html header (B):
'
'               'MAIL = MAIL & cDASH2 & cBoundaryB & vbCrLf & "Content-Type: text/html;" & vbCrLf     '"--Boundary_B_3435FE2_6617B_BB"
'               'MAIL = MAIL & vbTab & cCharset & cCTEnc & vbCrLf
'                    'LRet = SendData(sDat)
'
'            '-- send HTML message body:
'               'LRet = SendData(sBodyText & vbCrLf & vbCrLf)
''                MAIL = MAIL & sBodyText & vbCrLf & vbCrLf
'            '-- send end of alternative boundary:
'                'MAIL = MAIL & cDASH2 & cBoundaryB & cDASH2 & vbCrLf & vbCrLf
'                   'LRet = SendData(sDat)
'
'                MAIL = MAIL & cDASH2 & cBoundaryB & vbCrLf & "Content-Type: text/html;" & vbCrLf
'                MAIL = MAIL & vbTab & cCharset & cCTEnc & vbCrLf & vbCrLf
'                MAIL = MAIL & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" & vbCrLf
'                MAIL = MAIL & "<HTML><HEAD>" & vbCrLf
'                MAIL = MAIL & "<META HTTP-EQUIV=3D""Content-Type"" CONTENT=3D""text/html; =" & vbCrLf & "charset=3DBig5"">" & vbCrLf
'                'Modify by Morgan 2008/10/30
'                'MAIL = MAIL & "<TITLE id=3DridTitle>N a t i o n a l   H o l i d a y</TITLE>" & vbCrLf & vbCrLf
'                MAIL = MAIL & "<TITLE id=3DridTitle>TAI E INTERNATIONAL PATENT & LAW OFFICE</TITLE>" & vbCrLf & vbCrLf
'
'                MAIL = MAIL & "<STYLE>BODY {" & vbCrLf
'                'Modified by Morgan 2014/2/13 代理人可能沒有字形會導致空白不見故多+Arial
'                'MAIL = MAIL & vbTab & "MARGIN-TOP: 25px; FONT-SIZE: 10pt; MARGIN-LEFT: 10px; COLOR: #0033cc; FONT-FAMILY: " & ConvertToQp("細明體") & ", " & ConvertToQp("新細明體") & " " & vbCrLf
'                MAIL = MAIL & vbTab & "MARGIN-TOP: 25px; FONT-SIZE: 10pt; MARGIN-LEFT: 10px; COLOR: #0033cc; FONT-FAMILY: " & ConvertToQp("細明體") & ", " & ConvertToQp("細明體") & "," & ConvertToQp("Arial") & " " & vbCrLf
'                MAIL = MAIL & "}" & vbCrLf
'                MAIL = MAIL & "</STYLE>" & vbCrLf & vbCrLf
'                MAIL = MAIL & "<META content=3D""MSHTML 6.00.2800.1578"" name=3DGENERATOR>"
'
'                'Added by Morgan 2019/8/27
'                If m_iSignatureID > 0 Then
'                  strSignFile = App.path & "\$SignHead.txt"
'                  If PUB_ReadDB2File(strSignFile, 60) = True Then
'                     strSignHead = PUB_ReadTextFile(strSignFile)
'                     strSignHead = PUB_FixFirstDot(strSignHead)
'                     MAIL = MAIL & strSignHead & vbCrLf
'                  End If
'                End If
'                'end 2019/8/27
'
'                MAIL = MAIL & "</HEAD>" & vbCrLf
'
'                'Modified by Morgan 2017/11/21
'                'MAIL = MAIL & "<BODY id=3DridBody bgColor=3D#ffffff =" & vbCrLf & "background=3D#ffffff>" & vbCrLf
'                MAIL = MAIL & "<BODY id=3DridBody bgColor=3D#ffffff>" & vbCrLf
'
'                'Modified by Morgan 2014/2/13 代理人可能沒有字形會導致空白不見故改Arial
'                'MAIL = MAIL & "<DIV><FONT face=3D" & ConvertToQp("新細明體") & " color=3D#000000></FONT>&nbsp;</DIV>" & vbCrLf
'                'Removed by Morgan 2019/11/20 應該沒用,取消
'                'MAIL = MAIL & "<DIV><FONT face=3D" & ConvertToQp("Arial") & " color=3D#000000></FONT>&nbsp;</DIV>" & vbCrLf
'                'end 2019/11/20
'                MAIL = MAIL & "<DIV>" & sBodyText   'edit by nickc 2007/01/07  不編碼 ConvertToQp(sBodyText)
'                MAIL = MAIL & "</DIV>" & vbCrLf & vbCrLf
'
''                'add by nickc 2006/12/28
''                'MAIL = Clip70(MAIL)
'                'edit by nickc 2006/12/06
'                'If BooPlain = False Then
'                If chkShowPic.Value = 1 Then
'                    MAIL = MAIL & "<DIV>"
'                    For i = 1 To UBound(APics)
''有鎖定圖大小
''                        MAIL = MAIL & "<IMG style=3D""WIDTH: 972px; HEIGHT: 215px"" height=3D234 alt=3D"""" =" & vbCrLf
''                        MAIL = MAIL & "hspace=3D0=20" & vbCrLf
''                        MAIL = MAIL & "src=3D""cid:" & sID & "." & Format(i, "000") & """ width=3D836 align=3Dbaseline =" & vbCrLf
''                        MAIL = MAIL & "border=3D0>"
''不鎖定圖大小
'                        MAIL = MAIL & "<IMG alt=3D"""" =" & vbCrLf
'                        MAIL = MAIL & "hspace=3D0=20" & vbCrLf
'                        MAIL = MAIL & "src=3D""cid:" & sID & "." & Format(i, "000") & """ align=3Dbaseline =" & vbCrLf
'                        MAIL = MAIL & "border=3D0>"
'                    Next i
'                    MAIL = MAIL & "</DIV>" & vbCrLf
'                End If
'
'                'Added by Morgan 2019/8/27
'                If m_iSignatureID > 0 Then
'                  strSignFile = App.path & "\$SignBody.txt"
'                  If PUB_ReadDB2File(strSignFile, 60 + m_iSignatureID) = True Then
'                     strSignBody = PUB_ReadTextFile(strSignFile)
'                     strSignBody = PUB_FixFirstDot(strSignBody)
'                     MAIL = MAIL & strSignBody & vbCrLf
'                  End If
'                End If
'                'end 2019/8/27
'
'                MAIL = MAIL & "</BODY></HTML>" & vbCrLf & vbCrLf
'
'                'Added by Morgan 2019/8/27
'                '落款圖檔
'                If m_iSignatureID > 0 Then
'                  strSignFile = App.path & "\$SignAttPic1.txt"
'                  If PUB_ReadDB2File(strSignFile, 57) = True Then
'                     strSignAtt = PUB_ReadTextFile(strSignFile)
'                     MAIL = MAIL & cDASH2 & cBoundaryB & vbCrLf
'                     MAIL = MAIL & strSignAtt & vbCrLf & vbCrLf
'                  End If
'
'                  strSignFile = App.path & "\$SignAttPic2.txt"
'                  If PUB_ReadDB2File(strSignFile, 58) = True Then
'                     strSignAtt = PUB_ReadTextFile(strSignFile)
'                     MAIL = MAIL & cDASH2 & cBoundaryB & vbCrLf
'                     MAIL = MAIL & strSignAtt & vbCrLf & vbCrLf
'                  End If
'                End If
'                'end 2019/8/27
'
'                MAIL = MAIL & cDASH2 & cBoundaryB & cDASH2 & vbCrLf & vbCrLf

                sPicBody = ""
                If chkShowPic.Value = 1 Then
                    sPicBody = sPicBody & "<DIV>"
                    For i = 1 To UBound(APics)
                        sPicBody = sPicBody & "<IMG alt=3D"""" =" & vbCrLf
                        sPicBody = sPicBody & "hspace=3D0=20" & vbCrLf
                        sPicBody = sPicBody & "src=3D""cid:" & sID & "." & Format(i, "000") & """ align=3Dbaseline =" & vbCrLf
                        sPicBody = sPicBody & "border=3D0>"
                    Next i
                    sPicBody = sPicBody & "</DIV>" & vbCrLf
                End If
                
                
                MAIL = MAIL & PUB_GetContentMIME(sBodyText, m_iSignatureID, sPicBody, m_ByUTF8)
'end 2019/11/21

            '-- do pictures:
                For i = 1 To UBound(APics)
                    'sDat = APics(i)           '-- get base64 encoded pic.
                    'LRet = SendData(sDat)   '-- send base64 encoded pic.
                    MAIL = MAIL & APics(i)
                Next
                
             '-- send END signal:
               'LRet = SendData(cDASH2 & cBoundaryA & cDASH2 & vbCrLf & cDOT & vbCrLf)  ' Boundary_A_3435FE2_6617A_AA--
               'Modified by Morgan 2017/11/10 修正大附件會觸發字串空間不足(#14)錯誤問題
               'MAIL = MAIL & cDASH2 & cBoundaryA & cDASH2 & vbCrLf & cDOT & vbCrLf
               MAIL = MAIL & cDASH2
               MAIL = MAIL & cBoundaryA
               MAIL = MAIL & cDASH2
               MAIL = MAIL & vbCrLf
               MAIL = MAIL & cDOT
               MAIL = MAIL & vbCrLf
               'end 2017/11/10
          End Select
'edit end
      
Retry:
      m_ErrStat = ""
      SeekMailErr = ""
      m_FinalResult = "" 'Added by Morgan 2013/9/4
      m_MailLog = Format(Now, "hh:mm:ss") & vbCrLf & "寄件人：" & FromName & "(" & FromMail & ")" & vbCrLf & "收件人：" & ToName & "(" & ToMail & ")" & vbCrLf & IIf(CCToMail <> "", "副本：" & CCToName & "(" & CCToMail & ")" & vbCrLf, "") & "主旨：" & Subj 'Added by Morgan 2013/10/18 Modified by Morgan 2016/12/14
      
      DoEvents
      Winsock1.Protocol = sckTCPProtocol
      Winsock1.RemoteHost = SMTP
      Winsock1.RemotePort = 25
      Winsock1.Connect
      'edit by nickc 2006/04/04 修改成若第一次錯誤，再往另一個伺服器發
      'If Not Response("220") Then GoTo ERRORMail
      
      If Not Response("220", False, 2) Then
          'Modified by Morgan 2013/3/26
'          'add by nickc 2006/12/28 msa 找不到試
'          If bolByMsa = False Then
'              Winsock1.Close
'              Winsock1.LocalPort = 0
'              Winsock1.Protocol = sckTCPProtocol
'              If SMTP = Server$ Then
'                'Modify by Morgan 2009/7/17 失敗改寄 192.168.1.10,若是用X400方式則不再重試
'                'Winsock1.RemoteHost = "exchange" 'edit by nick 2007/10/16 連不到 "211.75.113.68"
'                If bolByX400 = True Then
'                  GoTo ERRORMail
'                Else
'                  Winsock1.RemoteHost = "192.168.1.10"
'                  bolPrimarySmtpFail = True
'                End If
'                'end 2009/7/17
'              ElseIf SMTP = "211.75.113.68" Then
'                Winsock1.RemoteHost = "168.95.4.211"
'              Else
'                Winsock1.RemoteHost = Server$
'              End If
'              Winsock1.RemotePort = 25
'              Winsock1.Connect
'              If Not Response("220") Then GoTo ERRORMail
'           Else
'              Winsock1.Close
'              Winsock1.LocalPort = 0
'              Winsock1.Protocol = sckTCPProtocol
'              Winsock1.RemotePort = 25
'              Winsock1.Connect
'              If Not Response("220") Then GoTo ERRORMail
'           End If
         Winsock1.Close
         Winsock1.LocalPort = 0
         Winsock1.Protocol = sckTCPProtocol
         'Modified by Morgan 2024/1/10
         'If SMTP <> m_SMTP_IP_ePaper Then
         '   Winsock1.RemoteHost = m_SMTP_IP_ePaper
         If SMTP <> m_SMTP_IP_BK Then
            Winsock1.RemoteHost = m_SMTP_IP_BK
         'end 2024/1/10
         ElseIf SMTP <> m_SMTP_IP_System Then
            Winsock1.RemoteHost = m_SMTP_IP_System
         End If
         Winsock1.RemotePort = 25
         Winsock1.Connect
         If Not Response("220") Then GoTo ERRORMail
         'end 2013/3/26
      End If
      
      DoEvents
      
      'Modify by Morgan 2009/8/21 改用電腦名稱
      'Winsock1.SendData ("HELO " & Domain & vbCrLf)
      Winsock1.SendData ("HELO " & Winsock1.LocalHostName & vbCrLf)
      If Not Response("250") Then m_ErrStat = "Err#01": GoTo ERRORMail
      
      DoEvents
      
      Winsock1.SendData (outFR)
      If Not Response("250") Then m_ErrStat = "Err#02 " & outFR: GoTo ERRORMail
      
      'Modified by Morgan 2013/5/20 多收件人改只要寄一次
      'Winsock1.SendData (outTO)
      'If Not Response("250") Then m_ErrStat = "Err#03": GoTo ERRORMail
      'Modified by Lydia 2015/08/03 +密件副本
     ' ArrRcpt = Split(ToMail & ";" & CCToMail, ";")
      ArrRcpt = Split(ToMail & ";" & CCToMail & ";" & BCCtoMail, ";")
      For MailCnt = LBound(ArrRcpt) To UBound(ArrRcpt)
         RcptMail = ArrRcpt(MailCnt)
         If RcptMail <> "" Then
            If InStr(RcptMail, "@") > 0 Then
               outTO = "rcpt to: " & RcptMail & vbCrLf
            Else
               outTO = "rcpt to: " & MailBefore & Replace(UCase(RcptMail), "@TAIE.COM.TW", "") & MailAfter & vbCrLf
            End If
            Winsock1.SendData (outTO)
            If Not Response("250") Then m_ErrStat = "Err#03 " & outTO: GoTo ERRORMail
         End If
      Next
      'end 2013/5/20
      
      Winsock1.SendData ("DATA" & vbCrLf)
      If Not Response("354") Then m_ErrStat = "Err#04": GoTo ERRORMail
      
      Winsock1.SendData (MAIL)
      If Not Response("250") Then m_ErrStat = "Err#05": GoTo ERRORMail
      
      DoEvents
      
      Winsock1.SendData ("quit" & vbCrLf)
      If Not Response("221") Then m_ErrStat = "Err#06": GoTo ERRORMail
      
      DoEvents
      
      SendMail = True
      'add by nickc 2007/01/04
      bolMailSendOk = True
      Mailing = False
      Winsock1.Close
      
      'Add By Sindy 2017/11/17 為了記錄寄送資訊
      If m_strUpdWhere <> "" And m_strTableName <> "" Then
         Select Case UCase(m_strTableName)
            Case "IPDEPTINPUT"
               strSql = "update " & m_strTableName & _
                       " set II20='T',II21=" & strSrvDate(1) & ",II22=" & Right("000000" & ServerTime, 6) & " where " & m_strUpdWhere
               cnnConnection.Execute strSql
         End Select
         m_strUpdWhere = ""
         m_strTableName = ""
      End If
      '2017/11/17 END
      
      'Added by Morgan 2025/6/26
      If m_SaveMailCP09 <> "" Then
         strSendDate = strSrvDate(1)
         strSendTime = Right("000000" & ServerTime, 6)
         strSql = "insert into smailbackup(smb01,smb02,smb03,smb04,smb05,smb06,smb07,smb08,smb09,smb10)" & _
                     " values('" & m_SaveMailCP09 & "'," & strSendDate & "," & strSendTime & _
                          ",'" & ChgSQL(FromMail & " " & FromName) & "'" & _
                          ",'" & ChgSQL(strTo) & "'" & _
                          ",'" & ChgSQL(strCC) & "'" & _
                          ",'" & ChgSQL(Trim(Subj)) & "'" & _
                          ",'" & ChgSQL(Trim(txtEmail(4).Text)) & "'" & _
                          ",'" & ChgSQL(Trim(Body)) & "'" & _
                          ",'" & ChgSQL(Trim(BCCtoMail)) & "')"
                          
         cnnConnection.Execute strSql, intR
         
         strSql = "insert into casepaperpdf(cpp01,cpp02,cpp03,cpp08,cpp09,cpp10)" & _
            " select cp09,fnCaseNo(cp01,cp02,cp03,cp04,1)||'.'||cp10||'." & strSendDate & strSendTime & ".INCOM." & EMP_Email & ".menu'" & _
            ",0," & strSendDate & "," & strSendTime & ",'Y' from caseprogress where cp09='" & m_SaveMailCP09 & "'"
                          
         cnnConnection.Execute strSql, intR
         m_SaveMailCP09 = ""
      End If
      'end 2025/6/26
      
      Screen.MousePointer = vbDefault
      'add by nickc 2007/01/04
      Exit Function
      
    End If
    
ERRORMail:

   'add by nickc 2007/01/04
   bolMailSendOk = False
   'Add By Sindy 2017/11/20 為了記錄寄送資訊
   If m_strUpdWhere <> "" And m_strTableName <> "" Then
      Select Case UCase(m_strTableName)
         Case "IPDEPTINPUT"
            strSql = "update " & m_strTableName & _
                    " set II20='F',II21=" & strSrvDate(1) & ",II22=" & Right("000000" & ServerTime, 6) & " where " & m_strUpdWhere
            cnnConnection.Execute strSql
      End Select
      m_strUpdWhere = ""
      m_strTableName = ""
   End If
   '2017/11/20 END
   
   'add by nickc 2007/05/22 紀錄錯誤
   If Err.NUMBER <> 0 Then
      m_ErrStat = "Err#" & Err.NUMBER
      SeekMailErr = Err.Description
   End If
   Mailing = False
   
   m_MailLog = m_ErrStat & " " & SeekMailErr & " " & Format(Now, "hh:mm:ss") & " <- " & m_MailLog
   WriteMailLog 'Added by Morgan 2013/10/18
   
   Winsock1.Close
   
   'Modified by Morgan 2014/1/23 不彈訊息則自動重試3次
   'Modified by Morgan 2014/6/12 改都自動重試3次
   'If bolMailFailNoAlert = True Then
   'Modified by Morgan 2015/1/16 若有錯誤碼時都要彈訊息
   If iRetry < 3 And SeekMailErr = "" Then
         iRetry = iRetry + 1
         GoTo Retry
      'End If
   'Added by Morgan 2015/11/2
   'Modify By Sindy 2016/12/12 不顯示傳送失敗的訊息，寫Log
   'ElseIf InStr(LCase(App.EXEName), "auto") > 0 Then
   ElseIf InStr(LCase(App.EXEName), "auto") > 0 Or bolShowErrMsg = False Then
   '2016/12/12 END
      Pub_WriteSysLog "系統發信失敗:" & "SMTP:" & Winsock1.RemoteHost & "(" & m_ErrStat & ":" & SeekMailErr & ",RCode:" & m_FinalResult & ")" & vbCrLf & strErrMsg
   'end 2015/11/2
   Else
      'Added by Morgan 2013/5/3
      strErrMsg = "寄件人：" & FromName & "(" & FromMail & ")" & vbCrLf & "收件人：" & ToName & "(" & ToMail & ")" & vbCrLf & IIf(CCToMail <> "", "副本：" & CCToName & "(" & CCToMail & ")" & vbCrLf, "") & "主旨：" & Subj & vbCrLf & "內文：" & Body & IIf(iType > 1, vbCrLf & "**有附件**", "")
      If MsgBox(vbCrLf & "系統發信失敗!是否重送??" & vbCrLf & vbCrLf & vbCrLf & strErrMsg & vbCrLf & vbCrLf & "【若選""否""請人工通知！】", vbCritical + vbYesNo + vbDefaultButton1, "SMTP:" & Winsock1.RemoteHost & "(" & m_ErrStat & ":" & SeekMailErr & ",RCode:" & m_FinalResult & ")") = vbYes Then
         iRetry = 0
         GoTo Retry
      Else
         Pub_WriteSysLog "系統發信失敗:" & "SMTP:" & Winsock1.RemoteHost & "(" & m_ErrStat & ":" & SeekMailErr & ",RCode:" & m_FinalResult & ")" & vbCrLf & strErrMsg
      End If
      'end 2013/5/3
   End If
   
   Screen.MousePointer = vbDefault
End Function

'Added by Morgan 2013/10/18
Private Sub WriteMailLog()
   Dim stSQL As String, intR As Integer
On Error Resume Next
   stSQL = "insert into MailFailLog(MFL01,MFL02,MFL03,MFL04,MFL05,MFL06) values(sysdate,'" & Winsock1.LocalIP & "','" & Winsock1.RemoteHostIP & "','" & ChgSQL(StrToStr(m_MailLog, 1500)) & "','" & strUserNum & "','" & ChgSQL(StrToStr(m_MailSub, 100)) & "')"
   cnnConnection.Execute stSQL, intR
End Sub

'edit by nickc 2006/04/04 加入若第一次錯誤，改發另一個伺服器
'Private Function Response(RCode$) As Boolean
Private Function Response(RCode$, Optional IsShow As Boolean = True, Optional iTimeDivider As Integer = 1) As Boolean
  Sec = 0
  
  'edit by nickc 2006/04/04 薛說改成 3 秒
  'edit by nickc 2006/04/13 再改回 50 秒
  'Modify by Morgan 2009/7/16 改10秒
  'Modified by Morgan 2011/11/16 改 3 秒
  'Modified by Morgan 2014/6/13 改 6 秒
  'Modified by Morgan 2014/7/17 改回 60 秒
  'Modified by Morgan 2024/7/23 改 10 秒
  Timer1.Interval = 1000
  Timer1.Enabled = True
  
  Response = True
  
  Do While Left$(Result, 3) <> RCode
    DoEvents
    If Sec > TimeOut / iTimeDivider Then
    
'Removed by Morgan 2013/5/3 訊息改統一在 SendMail 彈,並加提供重試選擇
'    'If Sec > 6 Then
'      If Len(Result) Then
'        'add by nickc 2006/12/28
'        If bolByMsa = False Then
'            'add by nickc 2007/01/04
'            If bolMailFailNoAlert = False Then
'                'add by nickc 2006/04/04
'                If IsShow = True Then
'                    'MsgBox "伺服器錯誤！", vbCritical
'                    'edit by nickc 2007/01/04
'                    'MsgBox "發 Mail 錯誤，請通知電腦中心！", vbCritical
'                    MsgBox "發 Mail 錯誤！" & vbCrLf & "收件人：" & Text4.Text & vbCrLf & "主旨：" & Text6.Text & vbCrLf & "內文：" & Text7.Text & vbCrLf & "請馬上電話連絡相關人員！", vbCritical, "寄信失敗！"
'                    'add by nickc 2008/05/12
'                    Response = False 'Add by Morgan 2009/9/11
'                    Exit Do
'                End If
'            End If
'        End If
'      Else
'        'add by nickc 2006/12/28
'        If bolByMsa = False Then
'            'add by nickc 2007/01/04
'            If bolMailFailNoAlert = False Then
'                'add by nickc 2006/04/04
'                If IsShow = True Then
'                    'MsgBox "伺服器逾時！", vbCritical
'                    'edit by nickc 2007/01/04
'                    'MsgBox "發 Mail 逾時，請通知電腦中心！", vbCritical
'                    MsgBox "發 Mail 錯誤！" & vbCrLf & "收件人：" & Text4.Text & vbCrLf & "主旨：" & Text6.Text & vbCrLf & "內文：" & Text7.Text & vbCrLf & "請馬上電話連絡相關人員！", vbCritical, "寄信失敗！"
'                    'add by nickc 2008/05/12
'                    Response = False 'Add by Morgan 2009/9/11
'                    Exit Do
'                End If
'            End If
'        End If
'      End If
'end 2013/5/3
      m_FinalResult = Result
      Response = False
      Exit Do
    End If
  Loop
  
  Result = ""
  Timer1.Enabled = False
  Err.Clear 'Add by Morgan 2009/9/11
End Function

'Added by Morgan 2013/3/26
'Modified by Morgan 2015/1/22 +SMTP_IP_ACC
Private Sub SetSMTP()
   Dim intQ As Integer, stSQL As String
   Dim rsQuery As ADODB.Recordset
   
   'Modified by Morgan 2015/6/8 取消 SMTP_IP_OUT
   'stSQL = "select ocode,oman from SetSpecMan where ocode in ('SMTP_IP','SMTP_IP_EPAPER','SMTP_IP_OUT')"
   'Modified by Morgan 2024/1/10
   'stSQL = "select ocode,oman from SetSpecMan where ocode in ('SMTP_IP','SMTP_IP_EPAPER')"
   'Modified by Morgan 2024/7/23 SMTP_IP_FW->SMTP_IP_BK
   stSQL = "select ocode,oman from SetSpecMan where ocode in ('SMTP_IP','SMTP_IP_BK')"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      Do While Not rsQuery.EOF
         If rsQuery("ocode") = "SMTP_IP" Then
            m_SMTP_IP_System = rsQuery("oman")
         'Modified by Morgan 2024/1/10
         'ElseIf rsQuery("ocode") = "SMTP_IP_EPAPER" Then
         '   m_SMTP_IP_ePaper = rsQuery("oman")
         ElseIf rsQuery("ocode") = "SMTP_IP_BK" Then
            m_SMTP_IP_BK = rsQuery("oman")
         'end 2024/1/10
         'ElseIf rsQuery("ocode") = "SMTP_IP_OUT" Then
         '   m_SMTP_IP_OUT = rsQuery("oman")
         End If
         rsQuery.MoveNext
      Loop
   End If
   Set rsQuery = Nothing
   
   If m_SMTP_IP_System = "" Then m_SMTP_IP_System = "192.168.1.15"
   'Modified by Morgan 2024/1/10
   'If m_SMTP_IP_ePaper = "" Then m_SMTP_IP_ePaper = "192.168.1.10"
   If m_SMTP_IP_BK = "" Then m_SMTP_IP_BK = "192.168.1.10"
   'end 2024/1/10
   'If m_SMTP_IP_OUT = "" Then m_SMTP_IP_OUT = "192.168.0.6"
End Sub


