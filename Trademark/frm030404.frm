VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030404 
   BorderStyle     =   1  '單線固定
   Caption         =   "FCT催延展(日文組)"
   ClientHeight    =   4810
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4810
   ScaleWidth      =   8240
   Begin VB.Frame Frame1 
      Height          =   1300
      Left            =   1560
      TabIndex        =   26
      Top             =   3030
      Width           =   2950
      Begin VB.OptionButton Option1 
         Caption         =   "委任狀不可援用+更址"
         Height          =   220
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   2770
      End
      Begin VB.OptionButton Option1 
         Caption         =   "委任狀可援用+更址"
         Height          =   220
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   690
         Width           =   2770
      End
      Begin VB.OptionButton Option1 
         Caption         =   "委任狀不可援用"
         Height          =   220
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   2770
      End
      Begin VB.OptionButton Option1 
         Caption         =   "委任狀可援用"
         Height          =   220
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   150
         Value           =   -1  'True
         Width           =   2770
      End
   End
   Begin VB.TextBox textTM01 
      Height          =   264
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   0
      Top             =   660
      Width           =   732
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2100
      Width           =   6492
   End
   Begin VB.TextBox textCP13 
      BorderStyle     =   0  '沒有框線
      Height          =   252
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2430
      Width           =   550
   End
   Begin VB.TextBox textTM44 
      BorderStyle     =   0  '沒有框線
      Height          =   252
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1030
   End
   Begin VB.TextBox textTM44_2 
      BorderStyle     =   0  '沒有框線
      Height          =   252
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2760
      Width           =   5410
   End
   Begin VB.TextBox textCP13_2 
      BorderStyle     =   0  '沒有框線
      Height          =   252
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2430
      Width           =   1540
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6210
      TabIndex        =   8
      Top             =   72
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7200
      TabIndex        =   9
      Top             =   72
      Width           =   912
   End
   Begin VB.TextBox textTM03 
      Height          =   264
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   2
      Top             =   660
      Width           =   372
   End
   Begin VB.TextBox textTM04 
      Height          =   264
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   3
      Top             =   660
      Width           =   732
   End
   Begin VB.TextBox textTM02 
      Height          =   264
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   1
      Top             =   660
      Width           =   1092
   End
   Begin MSForms.TextBox textTM07 
      Height          =   285
      Left            =   1560
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1740
      Width           =   6495
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "11456;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM05 
      Height          =   285
      Left            =   1560
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1020
      Width           =   6495
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "11456;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM06 
      Height          =   285
      Left            =   1560
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1380
      Width           =   6495
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "11456;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      Caption         =   "定稿內容："
      Height          =   250
      Left            =   240
      TabIndex        =   19
      Top             =   3090
      Width           =   1280
   End
   Begin VB.Label Label9 
      Caption         =   "FC代理人："
      Height          =   250
      Left            =   240
      TabIndex        =   18
      Top             =   2760
      Width           =   1280
   End
   Begin VB.Label Label7 
      Caption         =   "智權人員："
      Height          =   250
      Left            =   240
      TabIndex        =   16
      Top             =   2430
      Width           =   1280
   End
   Begin VB.Label Label2 
      Caption         =   "商品類別："
      Height          =   250
      Left            =   240
      TabIndex        =   14
      Top             =   2100
      Width           =   1280
   End
   Begin VB.Label Label1 
      Caption         =   "案件中文名稱："
      Height          =   252
      Left            =   240
      TabIndex        =   13
      Top             =   1020
      Width           =   1275
   End
   Begin VB.Label Label4 
      Caption         =   "案件英文名稱："
      Height          =   252
      Left            =   240
      TabIndex        =   12
      Top             =   1380
      Width           =   1275
   End
   Begin VB.Label Label5 
      Caption         =   "案件日文名稱："
      Height          =   252
      Left            =   240
      TabIndex        =   11
      Top             =   1740
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "本所案號："
      Height          =   252
      Left            =   240
      TabIndex        =   10
      Top             =   660
      Width           =   1275
   End
End
Attribute VB_Name = "frm030404"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Sindy 2023/9/12
Option Explicit

Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
Dim m_TM05 As String
Dim m_TM44 As String
Dim m_fa76 As String
Dim m_CP09 As String
Dim m_NP22 As String
'申請人
Dim m_TM23 As String
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String
Dim m_strSubject As String
Dim m_CaseAddr As String
Dim objOutLook As Object
Dim objMail As Object
Dim m_strContent As String


Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Clear()
   textTM01 = Empty
   textTM02 = Empty
   textTM03 = Empty
   textTM04 = Empty
   textTM05 = Empty
   textTM06 = Empty
   textTM07 = Empty
   textTM09 = Empty
   textCP13 = Empty
   textCP13_2 = Empty
   textTM44 = Empty
   textTM44_2 = Empty
   m_TM44 = Empty
   m_fa76 = Empty
   m_CP09 = Empty
   m_NP22 = Empty
   '申請人
   m_TM23 = Empty
   m_TM78 = Empty
   m_TM79 = Empty
   m_TM80 = Empty
   m_TM81 = Empty
   m_strSubject = Empty
   m_CaseAddr = Empty
End Sub

Private Sub cmdok_Click()
   If CheckDataValid() = True Then
      '設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      If OnProcess = True Then
         textTM01.SetFocus
         Call OpenOutLook
         Clear
      End If
      '設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
'    textTM05.BackColor = &H8000000F
'    textTM06.BackColor = &H8000000F
'    textTM07.BackColor = &H8000000F
'    textTM09.BackColor = &H8000000F
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm030404 = Nothing
End Sub

Private Sub textTM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'本所案號的系統別
Private Sub textTM01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM01) = False Then
      '檢查系統類別
      If IsCorrectSysKind(textTM01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "本所案號中的系統別不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM01_GotFocus
         GoTo EXITSUB
      End If
      '檢查使用者權限
      If IsUserHasRightOfSystem(strUserNum, textTM01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "您沒有使用該系統類別的權限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM01_GotFocus
         GoTo EXITSUB
      End If
      
      Select Case textTM01
         Case "FCT":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "本所案號中的系統別不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM01_GotFocus
      End Select
   End If
EXITSUB:
End Sub

Private Sub textTM04_LostFocus()
Dim strSql As String
Dim rsTmp As ADODB.Recordset
Dim strTit As String
Dim strMsg As String
Dim nResponse As Integer
   
   '清除資料
   textTM05 = Empty
   textTM06 = Empty
   textTM07 = Empty
   textTM09 = Empty
   '設定本所案號
   m_TM01 = textTM01
   m_TM02 = textTM02
   m_TM03 = textTM03
   If IsEmptyText(m_TM03) = True Then
      m_TM03 = "0"
      textTM03 = "0"
   End If
   m_TM04 = textTM04
   If IsEmptyText(m_TM04) = True Then
      m_TM04 = "00"
      textTM04 = "00"
   End If
   
   '查詢商標基本檔
   strSql = "SELECT * FROM TRADEMARK " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "'"
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      '案件中文名稱
      m_TM05 = ""
      If IsNull(rsTmp.Fields("TM05")) = False Then
         textTM05 = rsTmp.Fields("TM05")
         m_TM05 = rsTmp.Fields("TM05")
      End If
      '案件英文名稱
      If IsNull(rsTmp.Fields("TM06")) = False Then
         textTM06 = rsTmp.Fields("TM06")
      End If
      '案件日文名稱
      If IsNull(rsTmp.Fields("TM07")) = False Then
         textTM07 = rsTmp.Fields("TM07")
      End If
      '商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      m_TM44 = "" & rsTmp.Fields("TM44")
      textTM44 = m_TM44
      '申請人
      m_TM23 = "" & rsTmp.Fields("TM23")
      m_TM78 = "" & rsTmp.Fields("TM78")
      m_TM79 = "" & rsTmp.Fields("TM79")
      m_TM80 = "" & rsTmp.Fields("TM80")
      m_TM81 = "" & rsTmp.Fields("TM81")
      '個案日文地址
      m_CaseAddr = "" & rsTmp.Fields("TM26")
      If "" & rsTmp.Fields("TM90") <> "" Then
         m_CaseAddr = m_CaseAddr & "、" & rsTmp.Fields("TM90")
      End If
      If "" & rsTmp.Fields("TM91") <> "" Then
         m_CaseAddr = m_CaseAddr & "、" & rsTmp.Fields("TM91")
      End If
      If "" & rsTmp.Fields("TM92") <> "" Then
         m_CaseAddr = m_CaseAddr & "、" & rsTmp.Fields("TM92")
      End If
      If "" & rsTmp.Fields("TM93") <> "" Then
         m_CaseAddr = m_CaseAddr & "、" & rsTmp.Fields("TM93")
      End If
      '主旨
      m_strSubject = " " & PUB_GetUniText(Me.Name, "台灣") & "商標" & PUB_GetUniText(Me.Name, "登錄") & " No." & "" & rsTmp.Fields("TM15") & _
                     "「" & "" & rsTmp.Fields("TM05") & "" & rsTmp.Fields("TM131") & "」" & _
                     "第" & "" & rsTmp.Fields("TM09") & "類 " & _
                     "(貴Ref:" & "" & rsTmp.Fields("TM45") & ") " & _
                     "(TaiE." & m_TM01 & "-" & m_TM02 & IIf(m_TM03 & m_TM04 = "000", "", "-" & m_TM03 & "-" & m_TM04) & ")" & _
                     "【更新期限通知】[renew]"
      cmdOK.Enabled = True
      
      '檢查有下一程序延展
      strSql = "SELECT * FROM NEXTPROGRESS,caseprogress,staff" & _
               " WHERE NP02 = '" & m_TM01 & "' AND NP03 = '" & m_TM02 & "' AND NP04 = '" & m_TM03 & "' AND NP05 = '" & m_TM04 & "'" & _
                 " AND NP06 is null AND NP07='102'" & _
                 " AND NP01=CP09(+) AND CP13=ST01(+)"
      Set rsTmp = New ADODB.Recordset
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount = 0 Then
         strTit = "資料檢核"
         strMsg = "無下一程序延展"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.textTM01.SetFocus
         textTM01_GotFocus
         cmdOK.Enabled = False
         Exit Sub
      Else
         m_CP09 = rsTmp.Fields("NP01")
         m_NP22 = rsTmp.Fields("NP22")
         textCP13 = "" & rsTmp.Fields("CP13")
         textCP13_2 = "" & rsTmp.Fields("ST02")
      End If
      
      '檢查FC代理人為日本
      strSql = "SELECT * FROM fagent WHERE FA01='" & Left(m_TM44, 8) & "' AND FA02='" & Mid(m_TM44, 9, 1) & "'" & _
               " and substr(FA10,1,3)='011'"
      Set rsTmp = New ADODB.Recordset
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount = 0 Then
         strTit = "資料檢核"
         strMsg = "FC代理人( " & m_TM44 & " )必須為日本籍"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.textTM01.SetFocus
         textTM01_GotFocus
         cmdOK.Enabled = False
         Exit Sub
      Else
         If "" & rsTmp.Fields("FA04") <> "" Then
            textTM44_2 = "" & rsTmp.Fields("FA04")
         ElseIf "" & rsTmp.Fields("FA05") & "" & rsTmp.Fields("FA63") & "" & rsTmp.Fields("FA64") & "" & rsTmp.Fields("FA65") <> "" Then
            textTM44_2 = "" & rsTmp.Fields("FA05") & "" & rsTmp.Fields("FA63") & "" & rsTmp.Fields("FA64") & "" & rsTmp.Fields("FA65")
         Else
            textTM44_2 = "" & rsTmp.Fields("FA06")
         End If
         m_fa76 = "" & rsTmp.Fields("FA76")
      End If
   Else
      strTit = "資料檢核"
      strMsg = "本所案號不存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Me.textTM01.SetFocus
      textTM01_GotFocus
      Clear
   End If
End Sub

'檢查資料輸入是否正確
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   
   '本所案號
   If IsEmptyText(textTM01) = True Or IsEmptyText(textTM02) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
      
   CheckDataValid = True
EXITSUB:
End Function

Private Function OnProcess() As Boolean
Dim ET01 As String, ET03 As String
Dim ii As Integer, intTM09Cnt As Integer, dblRate As Double
Dim strTxt(200) As String
Dim dblOfficeFee As Double, strOfficeFee As String
Dim dblTaieFee As Double, strTaieFee As String, intCount As Integer
Dim intTotFee As Integer, strTotFee As String
Dim strNo As String, strCUAddr As String
Dim rsC As New ADODB.Recordset
Dim strCompDate As String, m_bolInsCP As Boolean, iRtn As Integer
Dim strNewCP09 As String
   
   If m_TM01 = "" Or m_TM02 = "" Or m_TM03 = "" Or m_TM04 = "" Then Exit Function
   strCompDate = CompDate(0, -2, strSrvDate(1)) '兩年內
   m_bolInsCP = True
   '檢查進度是否已有1717本所通知延展
   strSql = "SELECT * FROM caseprogress" & _
            " WHERE CP01 = '" & m_TM01 & "' AND CP02 = '" & m_TM02 & "' AND CP03 = '" & m_TM03 & "' AND CP04 = '" & m_TM04 & "'" & _
            " AND CP10='1717'" & _
            " AND CP05>=" & strCompDate
   rsC.CursorLocation = adUseClient
   rsC.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsC.RecordCount > 0 Then
      iRtn = MsgBox("兩年內進度檔已有 " & rsC.RecordCount & " 筆本所通知延展" & vbCrLf & vbCrLf & "請確認，此次是否要產生本所通知延展進度？" & vbCrLf & vbCrLf & _
                    "是:產生進度　否:不產生進度　取消:放棄執行此作業", vbYesNoCancel + vbDefaultButton3 + vbExclamation)
      If iRtn = vbCancel Then
         Screen.MousePointer = vbDefault
         Exit Function
      ElseIf iRtn = vbNo Then
         m_bolInsCP = False
      End If
   End If
   If rsC.State <> adStateClosed Then rsC.Close
   Set rsC = Nothing
     
   OnProcess = False
   
   ET01 = "10"
   
   m_strContent = ""
   '委任狀可援用
   If Option1(0).Value = True Then
      ET03 = "09"
   ElseIf Option1(1).Value = True Then
      ET03 = "10"
   '委任狀可援用+更址
   ElseIf Option1(2).Value = True Then
      ET03 = "11"
   ElseIf Option1(3).Value = True Then
      ET03 = "12"
   End If
   
   '更址
   If ET03 = "11" Or ET03 = "12" Then
      'X編號日文地址
      For ii = 1 To 5
         If ii = 1 Then strNo = m_TM23
         If ii = 2 Then strNo = m_TM78
         If ii = 3 Then strNo = m_TM79
         If ii = 4 Then strNo = m_TM80
         If ii = 5 Then strNo = m_TM81
         If strNo <> "" Then
            strSql = "SELECT CU29 FROM customer WHERE CU01='" & Left(strNo, 8) & "' AND CU02='" & Mid(strNo, 9, 1) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If ii > 1 Then strCUAddr = strCUAddr & "、"
               If "" & RsTemp.Fields("CU29") <> "" Then
                  strCUAddr = strCUAddr & RsTemp.Fields("CU29")
               End If
            End If
         End If
      Next ii
   End If
   
   intTM09Cnt = GetTMKindCnt(m_TM01, m_TM02, m_TM03, m_TM04)
   '請款匯率對台幣匯率
   'Modify By Sindy 2024/2/2 May:催延展定稿修改匯率計算方式為固定匯率
   'FCT日文催延展的定稿,報價的匯率計算如採現行匯率時,辦理延展時會有差價,如較報價USD為高時,客戶就會不高興
   '故請修改匯率計算方式為固定的,不抓電腦的機動匯率,請固定以USD1 = NTD28
   'dblRate = PUB_GetUSXRate_1(strSrvDate(2), "USD")
   'Modify By Sindy 2025/1/21 FCT日文組催延展定稿的匯率請改為NT$29＝US$1.00
   'Modify By Sindy 2025/5/7 改寫為抓系統特殊設定
   'dblRate = 29 '28
   dblRate = Pub_GetSpecMan("FCT日文組催延展定稿固定匯率")
   '2025/5/7 END
   '2024/2/2 END
   
   '政府料金=規費
   dblOfficeFee = (Val(4000) * intTM09Cnt) 'Modify By Sindy 2024/2/26 +val 就沒有出現溢位錯誤訊息了
   strOfficeFee = "NT$4,000 x " & intTM09Cnt & PUB_GetUniText(Me.Name, "區") & "分 = NT$" & Format(dblOfficeFee, "###,###,##0")
   '本所手續費+雜費=服務費
   If intTM09Cnt = 1 Then
      '代理人編號=Y48804000、Y48804010、Y48840時
      If m_TM44 = "Y48804000" Or m_TM44 = "Y48804010" Or Left(m_TM44, 6) = "Y48840" Then
         dblTaieFee = (8000 * (80 / 100))
         strTaieFee = "NT$8,000 x 80% = NT$" & Format(dblTaieFee, "###,###,##0")
      '客戶直接來所
      ElseIf m_fa76 = "B" Then
         dblTaieFee = (8000 + 500)
         strTaieFee = "NT$8,000 + " & PUB_GetUniText(Me.Name, "雜") & "費NT$500 = NT$" & Format(dblTaieFee, "###,###,##0")
      '一般
      Else
         dblTaieFee = (8000 * (90 / 100))
         strTaieFee = "NT$8,000 x 90% = NT$" & Format(dblTaieFee, "###,###,##0")
      End If
   Else '跨類
      '客戶直接來所
      If m_fa76 = "B" Then
         dblTaieFee = (8000 + (2000 * (intTM09Cnt - 1)) + 500)
         strTaieFee = "NT$8,000 + (二" & PUB_GetUniText(Me.Name, "區") & "分目以降NT$2,000 x " & (intTM09Cnt - 1) & PUB_GetUniText(Me.Name, "區") & "分) + " & PUB_GetUniText(Me.Name, "雜") & "費NT$500 = NT$" & Format(dblTaieFee, "###,###,##0")
      Else
         '代理人折扣:
         '代理人編號=Y48804000、Y48804010、Y48840時，為80%(0.8)
         If m_TM44 = "Y48804000" Or m_TM44 = "Y48804010" Or Left(m_TM44, 6) = "Y48840" Then
            intCount = 80
         Else
            '代理人編號為上述以外時 , 為90%
            intCount = 90
         End If
         dblTaieFee = ((8000 * (intCount / 100)) + (1000 * (intTM09Cnt - 1)))
         strTaieFee = "(NT$8,000 x " & intCount & "%) + (二" & PUB_GetUniText(Me.Name, "區") & "分目以降NT$1,000 x " & (intTM09Cnt - 1) & PUB_GetUniText(Me.Name, "區") & "分) = NT$" & Format(dblTaieFee, "###,###,##0")
      End If
   End If
   '合計
   intTotFee = (dblOfficeFee / dblRate) + (dblTaieFee / dblRate)
   strTotFee = "NT$" & Format((dblTaieFee + dblOfficeFee), "###,###,##0") & _
               " (US$" & Format(intTotFee, "###,###,##0") & ")"
   
   '處理定稿中的變數:
   ii = 0
   EndLetter ET01, m_CP09, ET03, strUserNum
   '商品類別數
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "','商品類別數','" & intTM09Cnt & "')"
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "','規費公式','" & strOfficeFee & "')"
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "','服務費公式','" & strTaieFee & "')"
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "','合計公式','" & strTotFee & "')"
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "','USD請款匯率','" & dblRate & "')"
   
   '更址
   If ET03 = "11" Or ET03 = "12" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "','個案日文地址','" & m_CaseAddr & "')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "','X編號日文地址','" & strCUAddr & "')"
      '客戶直接來所
      If m_fa76 = "B" Then
         strTaieFee = Format(3000, "###,###,##0")
         strTotFee = Format(3000 / dblRate, "###,###,##0")
      Else
         strTaieFee = Format(2000, "###,###,##0")
         strTotFee = Format(2000 / dblRate, "###,###,##0")
      End If
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "','變更地址費用','" & strTaieFee & "')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "','變更地址美金','" & strTotFee & "')"
   End If
   
   '***
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
      Exit Function
   End If
   '***
   
   '產生定稿內容
   NowPrint m_CP09, ET01, ET03, False, strUserNum, , , True, m_strContent
   If m_strContent = "" Then
      MsgBox "無郵件內容！", vbExclamation
      Exit Function
   Else
      m_strContent = Replace(m_strContent, "≦", "<")
      m_strContent = Replace(m_strContent, "≧", ">")
   End If
   
   If m_bolInsCP = True Then
      strNewCP09 = AutoNo("D", 6)
      'Modify By Sindy 2023/11/27 有關FCT日文組之催延展
      '請依May 指示:
      '於Danny輸入的同時一併將進度檔之智權人員改為Danny
      '另,下一程序檔之延展期限智權人員亦改為Danny
      'IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) => strUserNum
      strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14," & _
               "CP20,CP26,CP32," & _
               "CP43,CP27,CP30) " & _
      "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strSrvDate(1) & "," & _
              "'" & strNewCP09 & "','1717','" & GetSalesArea(strUserNum) & "','" & strUserNum & "','" & strUserNum & "'," & _
              "'N','N','N'," & _
              "'" & m_CP09 & "'," & strSrvDate(1) & "," & CNULL(m_NP22) & ")"
      cnnConnection.Execute strSql
      'Add By Sindy 2023/11/27
      strSql = "UPDATE nextProgress SET np10='" & strUserNum & "'" & _
               " WHERE np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "'" & _
               " and np07='102' and np06 is null"
      cnnConnection.Execute strSql
      '2023/11/27 END
   End If
   
   OnProcess = True
   
EXITSUB:
End Function

Private Sub OpenOutLook()
   '呼叫新郵件：
   Set objOutLook = CreateObject("Outlook.Application")
   Set objMail = objOutLook.CreateItem(0)
   
   '轉HTML格式:
   '先把換行+空白行轉換成<BR>
   m_strContent = Replace(m_strContent, vbCrLf & vbCrLf, vbCrLf & "<BR>")
   'm_strContent = "<p class=MsoNormal>" & Replace(m_strContent, vbCrLf, "<p><p class=MsoNormal>") & "</p>" '引用class
   '再把換行轉換成<P>
   m_strContent = Replace(m_strContent, vbCrLf, "<p style=""margin:0cm;"">")
   
'      If TypeName(objOutLook.Assistant) <> "Nothing" Then
'         objOutLook.ActiveWindow.WindowState = 1 '0.最大化 1.視窗小點
'      End If
   With objMail
      .BodyFormat = 2 '2=olFormatHTML 1=olFormatPlain 3=olFormatRichText
      '.To = strTo
      '.cc = strCC
'         If strAttach <> "" Then
'            .Attachments.add strAttach '加附件
'         End If
      .Subject = m_strSubject
      '字型:MS Gothic
      '字型大小:10.5
      .HTMLBody = "<HTML>" & _
                  "<head>" & _
                  "<style>" & _
                  "<!-- " & _
                  "/* Style Definitions */ " & _
                  "p.MsoNormal , li.MsoNormal, div.MsoNormal " & _
                     "{margin:0cm; " & _
                     "margin-bottom:.0001pt; " & _
                     "font-size:10.5pt; " & _
                     "font-family:""MS Gothic"",serif;} " & _
                  "--> " & _
                  "</style></head>" & _
                  "<BODY><font face=""MS Gothic""><div style=""font-size:14px;"">" & m_strContent & "</div></font></BODY></HTML>"
      .Display
   End With
   
   Set objMail = Nothing
   Set objOutLook = Nothing
End Sub

Private Sub textTM01_GotFocus()
   InverseTextBox textTM01
End Sub

Private Sub textTM02_GotFocus()
   InverseTextBox textTM02
End Sub

Private Sub textTM03_GotFocus()
   InverseTextBox textTM03
End Sub

Private Sub textTM04_GotFocus()
   InverseTextBox textTM04
End Sub
