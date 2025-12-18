VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030301_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "發E-Mail對象"
   ClientHeight    =   3585
   ClientLeft      =   2130
   ClientTop       =   2895
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form18"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4695
   Begin VB.OptionButton Option1 
      Caption         =   "直屬主管"
      Height          =   180
      Index           =   4
      Left            =   780
      TabIndex        =   3
      Top             =   2130
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CheckBox chkByOutLook 
      Caption         =   "密件副本：操作者本人"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   3120
      Width           =   2580
   End
   Begin VB.OptionButton Option1 
      Caption         =   "核稿人"
      Height          =   180
      Index           =   3
      Left            =   90
      TabIndex        =   4
      Top             =   1710
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.CommandButton cmdGoInput 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3825
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3000
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   60
      Width           =   800
   End
   Begin VB.OptionButton Option1 
      Caption         =   "承辦人"
      Height          =   180
      Index           =   2
      Left            =   90
      TabIndex        =   2
      Top             =   1140
      Width           =   852
   End
   Begin VB.OptionButton Option1 
      Caption         =   "智權人員"
      Height          =   180
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   855
      Width           =   1035
   End
   Begin VB.OptionButton Option1 
      Caption         =   "管制人"
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   570
      Value           =   -1  'True
      Width           =   972
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   3690
      Top             =   3030
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   3060
      Top             =   3030
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSForms.TextBox txt1 
      Height          =   615
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   2370
      Visible         =   0   'False
      Width           =   4575
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "12621;1147"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   615
      Index           =   0
      Left            =   60
      TabIndex        =   10
      Top             =   2370
      Width           =   4575
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "12621;1147"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "直屬主管"
      Height          =   180
      Left            =   360
      TabIndex        =   17
      Top             =   1425
      Width           =   720
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   16
      Top             =   1425
      Width           =   3420
      VariousPropertyBits=   27
      Size            =   "6032;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "完稿日："
      Height          =   180
      Left            =   360
      TabIndex        =   15
      Top             =   270
      Visible         =   0   'False
      Width           =   720
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   14
      Top             =   1710
      Visible         =   0   'False
      Width           =   3420
      VariousPropertyBits=   27
      Size            =   "6032;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   13
      Top             =   855
      Width           =   3420
      VariousPropertyBits=   27
      Size            =   "6032;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   12
      Top             =   1140
      Width           =   3420
      VariousPropertyBits=   27
      Size            =   "6032;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   11
      Top             =   570
      Width           =   3420
      VariousPropertyBits=   27
      Size            =   "6032;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備註："
      Height          =   180
      Left            =   60
      TabIndex        =   9
      Top             =   2130
      Width           =   540
   End
End
Attribute VB_Name = "frm030301_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/23 改成Form2.0 ; lbl1(index)、txt1(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
Option Explicit

Dim i As Integer, j As Integer, strSql As String, s As Integer
Dim Str01 As String, Str02 As String, Str03 As String, strTemp As String
'因為 mail 解析，中文會有問題，所以用員工編號
Public StrMailNum1 As String        '管制人
Public StrMailNum2 As String        '智權人員
Public StrMailNum3 As String        '承辦人
'紀錄作用按鍵
Public cmdState As Integer
Public strCP09 As String                '收文號
Public StrMailNum4 As String        '核稿人
Public StrMailNum5 As String        '直屬主管
Public strEvents As String               '事件
Public strLimitKind As String          '期限
Public strNP22 As String 'Add By Sindy 2015/4/9
Dim strFCP201State As String 'FCP翻譯控制狀態 0:無 1:外譯且已完稿 2:外譯且未完稿


Public Sub PubShowNextData()
   Dim strTo As String '收件者員工編號
   Dim strToName As String
   Dim bolByOutLook As Boolean
   Dim stContent As String, stSubject As String
   Dim stContentPrint As String
   
   Select Case cmdState
      Case 0 '發E-Mail
         
         If Option1(0).Value = True Then
            strTo = StrMailNum1
            strToName = lbl1(0).Caption
         ElseIf Option1(1).Value = True Then
            strTo = StrMailNum2
            strToName = lbl1(1).Caption
         ElseIf Option1(2).Value = True Then
'            Select Case strFCP201State
'               Case "1"
'                  MsgBox "本案承辦為外譯人員且已完稿！"
'                  Exit Sub
'               Case "2"
'                  If MsgBox("本案承辦為外譯人員是否要改通知靜芳？", vbYesNo + vbDefaultButton1) = vbYes Then
'                     strTo = "73023"
'                  Else
'                     Exit Sub
'                  End If
'               Case Else
                  strTo = StrMailNum3
'            End Select
            strToName = lbl1(2).Caption
         ElseIf Option1(3).Value = True Then
            strTo = StrMailNum4
            strToName = lbl1(3).Caption
         ElseIf Option1(4).Value = True Then
            strTo = StrMailNum5
            strToName = lbl1(4).Caption
         End If
         
         If strTo = "" Then
            MsgBox "收件人空白，無法寄送！"
            Exit Sub
         End If
         
         Screen.MousePointer = vbHourglass
         
         stSubject = "通知" & strLimitKind & "期限"
         stContentPrint = "                收受人姓名：" + strToName + vbCrLf + vbCrLf + txt1(1) + vbCrLf + vbCrLf + "                " + txt1(0) + vbCrLf + vbCrLf
         stContent = "TO：收受人姓名：" + strToName + vbCrLf + vbCrLf + txt1(1) + vbCrLf + vbCrLf + "                " + txt1(0) + vbCrLf + vbCrLf + "FROM：" & strTemp + vbCrLf
         
         '加可選擇要有寄件備份
         If chkByOutLook.Value = 1 Then
            'Modified by Lydia 2020/03/13 因為OutLook過去和現在版本不同,所以改用密件副本保留
'            DoEvents
'            MAPISession1.LogonUI = False
'            MAPISession1.UserName = strUserNum
'            Err.Clear
'On Error Resume Next
'            MAPISession1.SignOn
'            If Err.Number <> 0 Then
'               MsgBox "EMail發送失敗!!請啟動 OutLook 後重試!!"
'               Screen.MousePointer = vbDefault
'               Exit Sub
'            End If
'            MAPIMessages1.SessionID = MAPISession1.SessionID
'            MAPIMessages1.MsgIndex = -1
'            MAPIMessages1.Compose
'            'Modify By Sindy 2014/1/16
'            'MAPIMessages1.MsgSubject = "◎系統代發◎" & stSubject
'            MAPIMessages1.MsgSubject = "◎" & IIf(Pub_StrUserSt03 = "M51" And PUB_GetST05(strUserNum) <> "" And UCase(PUB_GetDbTerminal) = "(M51-1)", PUB_GetDbTerminal, "") & stSubject
'            '2014/1/16 END
'            MAPIMessages1.MsgNoteText = stContent
'            MAPIMessages1.RecipIndex = 0
'            MAPIMessages1.RecipType = 1 '收件者是主要收件者
'            MAPIMessages1.RecipDisplayName = ChkMailId(strTo)
'            MAPIMessages1.ResolveName
'            MAPIMessages1.RecipIndex = 1
'            MAPIMessages1.RecipType = 2 '收件者屬於「副本」收件者
'            MAPIMessages1.RecipDisplayName = ChkMailId(StrMailNum5)
'            MAPIMessages1.ResolveName
'            MAPIMessages1.Send
'            MAPISession1.SignOff
            PUB_SendMail strUserNum, strTo, "", stSubject, stContent, "", "", , False, False, StrMailNum5, , , , , , strUserNum
            'end 2020/03/13
         Else
            '無寄件備份,加副本功能
            'Modified by Morgan 2014/1/24 改用預設(Html)格式，因純文字格式遇造字會全部內容變亂碼
            'PUB_SendMail strUserNum, strTo, "", stSubject, stContent, "", "", False, False, False, StrMailNum5
            PUB_SendMail strUserNum, strTo, "", stSubject, stContent, "", "", , False, False, StrMailNum5
         End If
         
         '會有寄信失敗又秀出成功的問題
         's = MsgBox("郵件已送出", , "MAIL!!")
         Screen.MousePointer = vbDefault
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
         
'         '列印
'         Printer.Font.Size = 12
'         Printer.Font.Underline = False
'         Printer.FontBold = False
'         Printer.CurrentX = 600
'         Printer.CurrentY = 900
'         Printer.Print "操作人員：" & strTemp
'         Printer.CurrentX = 600
'         Printer.CurrentY = 1200
'         Printer.Print "列印日期：" & ChangeTStringToTDateString(Format(Now(), "yyyymmdd") - 19110000)
'         Printer.CurrentX = 0
'         Printer.CurrentY = 1800
'         Printer.Print stContentPrint
'         Printer.EndDoc
         
      Case 1 '結束
        tmpBol = fnCancelNowFormAndShowParentForm(Me)
        
      Case Else
   End Select
End Sub

Private Sub cmdGoInput_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   For i = 0 To 4
      lbl1(i).Caption = ""
      lbl1(i).Visible = False
      Option1(i).Visible = False
      Option1(i).Value = False
   Next i
   txt1(0).Text = "本案" & strLimitKind & "期限將至，請儘速作業，以利後續作業。"
   cmdState = -1
'   If pub_strUserOffice <> "1" Then
'      chkByOutLook.Visible = False
'   End If
End Sub

Sub StrMenu()
   Str01 = SystemNumber(Me.Tag, 1)
   Str02 = SystemNumber(Me.Tag, 2)
   Str03 = SystemNumber(Me.Tag, 3)
   'lbl1(0).Caption = Str01
   'lbl1(1).Caption = Str02
   'lbl1(2).Caption = Str03
   CheckOC
   strSql = "SELECT ST02 FROM STAFF WHERE ST01='" & strUserNum & "'"
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       If Not IsNull(adoRecordset.Fields(0)) Then
           strTemp = adoRecordset.Fields(0)
       Else
           strTemp = strUserNum
       End If
   Else
       strTemp = strUserNum
   End If
   CheckOC
   
   strFCP201State = "0"
   If strCP09 <> "" Then
      '未收文的智權人員，要抓下一程序智權人員
      'Modify By Sindy 2015/4/9 +" & IIf(strNP22 <> "", " and np22=" & strNP22, "") & "
      'strExc(0) = "select cp01,cp10,cp13,st3.st02 cp13n,cp14,st1.st02 cp14n,st1.st03 cp14d,ep04,st2.st02 ep04n,ep09,np10,st4.st02 np10n from caseprogress, engineerprogress,staff st1,staff st2,staff st3,nextprogress,staff st4 where cp09='" & strCP09 & "' and ep02(+)=cp09 and st1.st01(+)=cp14 and st2.st01(+)=ep04 and st3.st01(+)=cp13 and cp09=np01(+) and np10=st4.st01(+) "
      strExc(0) = "select cp01,cp10,cp13,st3.st02 cp13n,cp14,st1.st02 cp14n,st1.st03 cp14d,ep04,st2.st02 ep04n,ep09,np10,st4.st02 np10n from caseprogress, engineerprogress,staff st1,staff st2,staff st3,nextprogress,staff st4 where cp09='" & strCP09 & "' and ep02(+)=cp09 and st1.st01(+)=cp14 and st2.st01(+)=ep04 and st3.st01(+)=cp13 and cp09=np01(+)" & IIf(strNP22 <> "", " and np22=" & strNP22, "") & " and np10=st4.st01(+) "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         '未收文，應抓下一程序 : 智權人員
         If strEvents = "未收文" Then
            StrMailNum2 = "" & .Fields("np10")
            lbl1(1).Caption = "" & .Fields("np10n")
         Else
            StrMailNum2 = "" & .Fields("cp13")
            lbl1(1).Caption = "" & .Fields("cp13n")
         End If
         '承辦人
         StrMailNum3 = "" & .Fields("cp14")
         lbl1(2).Caption = "" & .Fields("cp14n")
         
'         '未收文-寄信對象智權人員
'         If strEvents = "未收文" Then
            Option1(1).Visible = True
            lbl1(1).Visible = True
            Option1(1).Value = True
            '直屬主管
            strExc(0) = "select st1.st01,st1.st02,st1.st52,st2.st01,st2.st02 from staff st1,staff st2 where st1.st01='" & StrMailNum2 & "' and st1.st52=st2.st01(+) "
            intI = 1
            Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               StrMailNum5 = "" & adoRecordset.Fields(2)
               lbl1(4).Caption = "" & adoRecordset.Fields(4)
            End If
            lbl1(4).Visible = True
'         '其他-寄信對象承辦人
'         Else
'            Option1(2).Visible = True
'            lbl1(2).Visible = True
'            Option1(2).Value = True
'            '直屬主管
'            strExc(0) = "select st1.st01,st1.st02,st1.st52,st2.st01,st2.st02 from staff st1,staff st2 where st1.st01='" & StrMailNum3 & "' and st1.st52=st2.st01(+) "
'            intI = 1
'            Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               StrMailNum5 = "" & adoRecordset.Fields(2)
'               lbl1(4).Caption = "" & adoRecordset.Fields(4)
'            End If
'            lbl1(4).Visible = True
'         End If
         
         End With
      End If
   Else
      Option1(1).Visible = True
      lbl1(1).Visible = True
      Option1(1).Value = True
      '直屬主管
      strExc(0) = "select st1.st01,st1.st02,st1.st52,st2.st01,st2.st02 from staff st1,staff st2 where st1.st01='" & StrMailNum2 & "' and st1.st52=st2.st01(+) "
      intI = 1
      Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         StrMailNum5 = "" & adoRecordset.Fields(2)
         lbl1(4).Caption = "" & adoRecordset.Fields(4)
      End If
      lbl1(4).Visible = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm030301_1 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   'Add By Cheng 2002/05/01
   Select Case Index
   Case 0
      '切換至中文輸入模式
      'edit by nickc 2007/06/06 切換輸入法改用API
      'Me.txt1(Index).IMEMode = 1
      OpenIme
   Case 1
      '切換至中文輸模式
      'edit by nickc 2007/06/06 切換輸入法改用API
      'Me.txt1(Index).IMEMode = 1
      OpenIme
   End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   'Add By Cheng 2002/05/01
   Select Case Index
   Case 0
      '取消中文輸入模式
      'edit by nickc 2007/06/06 切換輸入法改用API
      'Me.txt1(Index).IMEMode = 2
      CloseIme
   Case 1
      '取消中文輸入模式
      'edit by nickc 2007/06/06 切換輸入法改用API
      'Me.txt1(Index).IMEMode = 2
      CloseIme
   End Select
End Sub
