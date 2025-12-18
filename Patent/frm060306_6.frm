VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060306_6 
   BorderStyle     =   1  '單線固定
   Caption         =   "請款通知函-專利權消滅請款函"
   ClientHeight    =   5340
   ClientLeft      =   750
   ClientTop       =   1545
   ClientWidth     =   7140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7140
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   3570
      TabIndex        =   20
      Top             =   5010
      Width           =   3465
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   " 回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4890
      TabIndex        =   2
      Top             =   36
      Width           =   1155
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Left            =   1488
      MaxLength       =   1
      TabIndex        =   0
      Top             =   5010
      Width           =   255
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6084
      TabIndex        =   3
      Top             =   36
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4110
      TabIndex        =   1
      Top             =   36
      Width           =   756
   End
   Begin VB.Frame Frame1 
      Height          =   2400
      Left            =   105
      TabIndex        =   4
      Top             =   468
      Width           =   6975
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   7
         Left            =   1230
         TabIndex        =   32
         Top             =   2145
         Width           =   5565
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9816;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   6
         Left            =   1230
         TabIndex        =   31
         Top             =   1905
         Width           =   5565
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9816;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   5
         Left            =   1230
         TabIndex        =   30
         Top             =   1680
         Width           =   5565
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9816;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   4
         Left            =   1230
         TabIndex        =   29
         Top             =   1440
         Width           =   5565
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9816;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   3
         Left            =   1230
         TabIndex        =   28
         Top             =   1185
         Width           =   5565
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9816;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   2
         Left            =   1230
         TabIndex        =   27
         Top             =   945
         Width           =   5565
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9816;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   1
         Left            =   1230
         TabIndex        =   26
         Top             =   405
         Width           =   5565
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9816;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   0
         Left            =   1230
         TabIndex        =   25
         Top             =   150
         Width           =   5565
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9816;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   285
         Left            =   1230
         TabIndex        =   23
         Top             =   600
         Width           =   5475
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "9657;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(英)"
         Height          =   180
         Index           =   8
         Left            =   48
         TabIndex        =   13
         Top             =   1932
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(日)"
         Height          =   180
         Index           =   7
         Left            =   48
         TabIndex        =   12
         Top             =   2172
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(中)"
         Height          =   180
         Index           =   6
         Left            =   48
         TabIndex        =   11
         Top             =   1692
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(英)"
         Height          =   180
         Index           =   5
         Left            =   48
         TabIndex        =   10
         Top             =   1212
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(日)"
         Height          =   180
         Index           =   4
         Left            =   48
         TabIndex        =   9
         Top             =   1452
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(中)"
         Height          =   180
         Index           =   3
         Left            =   48
         TabIndex        =   8
         Top             =   972
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利名稱"
         Height          =   180
         Index           =   2
         Left            =   48
         TabIndex        =   7
         Top             =   612
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "請款函日期"
         Height          =   180
         Index           =   1
         Left            =   48
         TabIndex        =   6
         Top             =   372
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號"
         Height          =   180
         Index           =   0
         Left            =   48
         TabIndex        =   5
         Top             =   132
         Width           =   720
      End
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   12
      Left            =   1350
      TabIndex        =   37
      Top             =   3840
      Width           =   5685
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10028;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   11
      Left            =   1350
      TabIndex        =   36
      Top             =   3600
      Width           =   5685
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10028;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   10
      Left            =   1350
      TabIndex        =   35
      Top             =   3360
      Width           =   5685
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10028;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   9
      Left            =   1350
      TabIndex        =   34
      Top             =   3120
      Width           =   5685
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10028;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   8
      Left            =   1350
      TabIndex        =   33
      Top             =   2880
      Width           =   5685
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10028;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   900
      Index           =   0
      Left            =   1320
      TabIndex        =   24
      Top             =   4050
      Width           =   5715
      VariousPropertyBits=   -1466939365
      MaxLength       =   2000
      ScrollBars      =   3
      Size            =   "10081;1587"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利權消滅日"
      Height          =   180
      Index           =   15
      Left            =   150
      TabIndex        =   22
      Top             =   3840
      Width           =   1080
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "請款單印表機"
      Height          =   180
      Index           =   11
      Left            =   2370
      TabIndex        =   21
      Top             =   5040
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否修改請款函        (Y)"
      Height          =   180
      Index           =   14
      Left            =   165
      TabIndex        =   19
      Top             =   5040
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "副本聯絡人"
      Height          =   180
      Index           =   13
      Left            =   168
      TabIndex        =   18
      Top             =   3600
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "副本收受人"
      Height          =   180
      Index           =   12
      Left            =   168
      TabIndex        =   17
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "P. S."
      Height          =   180
      Index           =   11
      Left            =   180
      TabIndex        =   16
      Top             =   4140
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "彼此案號"
      Height          =   180
      Index           =   10
      Left            =   168
      TabIndex        =   15
      Top             =   3120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶案件案號"
      Height          =   180
      Index           =   9
      Left            =   168
      TabIndex        =   14
      Top             =   2880
      Width           =   1080
   End
End
Attribute VB_Name = "frm060306_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/16 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Dim intWhere As Integer, strReceiveNo As String
Dim m_CP25 As String, m_CP10 As String, m_A1J17 As Long
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String
Dim strPrinter As String
Const ET01 As String = "09"
Dim m_bolEmail As Boolean, m_bolPlusPaper As Boolean, m_iCopy As Integer
'Added by Morgan 2014/6/3
Dim m_bolDNEmail As Boolean, m_bolDNPlusPaper As Boolean
Dim m_LetterLanguage As String 'Add By Sindy 2015/9/21


Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
'Modified by Lydia 2016/11/17 strTxt(1 To 2) => strTxt(1 To 9)
 Dim strTxt(1 To 9) As String, i As Integer, j As Integer, strTmp As String
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   If m_CP25 = "N" Then
      strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','列印備註','lapsed')"
   Else
      strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','列印備註','expired')"
   End If
   i = 2
    'Add By Cheng 2003/02/24
   '請款函日期
   If frm060306.Text5.Text <> "" Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','請款函日期','" & DBDATE(frm060306.Text5.Text) & "')"
      i = i + 1
   End If
   '請款函備註
   If Text1(0).Text <> "" Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','請款函備註','P.S. " & ChgSQL(Text1(0).Text) & "')"
      i = i + 1
   End If
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(i - 1, strTxt) Then
   If Not ClsLawExecSQL(i - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Dim bolChk As Boolean
   
   
   Select Case Index
      Case 2 '結束
         Unload frm060306
         Unload Me
      Case 0 '確定
        Screen.MousePointer = vbHourglass
         m_LetterLanguage = PUB_GetLanguage(pa(1), pa(2), pa(3), pa(4)) 'Add By Sindy 2015/9/21
         If Text2.Text = "Y" Then bolChk = True
         '請再區分英文02或日文03
         StartLetter ET01, "02"
         'Add by Morgan 2008/3/31 判斷是否產生電子檔
         m_bolEmail = PUB_GetEMailFlag(pa(1) & pa(2) & pa(3) & pa(4), , , m_bolPlusPaper)
         'Added by Morgan 2014/6/3
         If m_bolEmail = False Then
            m_bolDNEmail = PUB_GetEMailFlag(pa(1) & pa(2) & pa(3) & pa(4), , , m_bolDNPlusPaper, , True)
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
            NowPrint strReceiveNo, ET01, "02", bolChk, strUserNum, , , , , m_iCopy, , True, True
            If bolChk = False Then
               MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(pa(1)) & " ]！"
            End If
         Else
            NowPrint strReceiveNo, ET01, "02", bolChk, strUserNum, 0
         End If
         
         If Not m_bolEmail Or m_bolPlusPaper Then
'            'Add By Sindy 2015/9/21 日文定稿才要印地址條
'            If m_LetterLanguage = "3" Or Val(外專開窗信函啟用日) >= Val(strSrvDate(1)) Then
'            '2015/9/21 END
            'Add By Sindy 2017/3/20 日文定稿才要印地址條
            If frm060306.m_FCna01 = "101" Or m_LetterLanguage = "3" Then '美國 或 日文定稿才要印地址條
            '2017/3/20 END
               '新增地址條列表資料
               pub_AddressListSN = pub_AddressListSN + 1
               PUB_AddNewAddressList strUserNum, frm060306.Text1.Text, frm060306.Text2.Text, frm060306.Text3.Text, frm060306.Text4.Text, "" & pub_AddressListSN, "0"
            End If
         End If

         '新增並列印請款單
         If ProcessPrint = False Then
            MsgBox "新增請款單資料錯誤 !", vbCritical
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         
         frm060306.Show
         frm060306.Clear
        Screen.MousePointer = vbDefault
         Unload Me
      Case 1 '回前畫面
        'Add By Cheng 2003/02/05
        '若請款單印表機變動, 則更新列印設定
        If Me.Combo2.Text <> Me.Combo2.Tag Then
            PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
        End If
         frm060306.Show
         Unload Me
   End Select
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   '若請款單印表機變動, 則更新列印設定
   If Me.Combo2.Text <> Me.Combo2.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
   End If
   Set frm060306_6 = Nothing
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Form_Load()
Dim ii As Integer
'Add By Cheng 2003/02/05
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   MoveFormToCenter Me
   intWhere = 國外_FC
   ReadPatent
   
'Modify by Morgan 2011/3/15 改共用且不要排除預設印表機
   PUB_SetPrinter Me.Name, Combo2, strPrinter
'end 2011/3/15
   
   '先抓固定請款金額
   strExc(0) = "SELECT A1J17 FROM ACC1J0 WHERE A1J01='" & pa(1) & "' AND A1J02='" & m_CP10 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If IsNull(RsTemp.Fields(0)) Then
         MsgBox "未建立固定請款金額，請先建資料才能產生請款單!!!"
         cmdOK_Click (1)
         Exit Sub
      Else
         m_A1J17 = RsTemp.Fields(0)
      End If
   Else
      MsgBox "未建立固定請款金額，請先建資料才能產生請款單!!!"
      cmdOK_Click (1)
      Exit Sub
   End If

End Sub

Private Sub ReadPatent()
 Dim Lbl As Object, i As Integer, strTmp As String
   For Each Lbl In Label2
      Lbl = ""
   Next
   strReceiveNo = frm060306.Tag
   pa(1) = frm060306.Text1.Text
   pa(2) = frm060306.Text2.Text
   pa(3) = frm060306.Text3.Text
   pa(4) = frm060306.Text4.Text
   Label2(0).Caption = GiveSymbol(pa(1), pa(2), pa(3), pa(4))
   Label2(1).Caption = frm060306.Text5.Text
   SetComboToCombo Combo1, frm060306.Combo1
   
   Select Case pa(1)
      Case "FCP"
         If ClsPDReadPatentDatabase(pa, intWhere) Then  'edit by nickc 2007/02/02 不用 dll 了  If objPublicData.ReadPatentDatabase(pA, intWhere) Then
            If PA51CU58FA07(pa) Then
               For i = 1 To 6
                  Label2(i + 1) = pa(50 + i)
               Next
            End If
            Label2(8) = pa(48)
            Label2(9) = pa(77)
            If pa(86) <> "" Then
               'edit by nickc 2007/02/05 不用 dll 了
               'If objLawDll.LawGetName(pa(86), strTmp) Then Label2(10) = strTmp
               If ClsLawLawGetName(pa(86), strTmp) Then Label2(10) = strTmp
            End If
            Label2(11) = pa(87)
         End If
      Case "FG"
         If ClsPDReadServicePracticeDatabase(pa, intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadServicePracticeDatabase(pA, intWhere) Then
            If PA51CU58FA07(pa) Then Label2(2) = pa(30)
            Label2(8) = pa(29)
            Label2(9) = pa(27)
            'edit by nickc 2007/02/05 不用 dll 了
            'If objLawDll.LawGetName(pa(35), strTmp) Then Label2(10) = strTmp
            If ClsLawLawGetName(pa(35), strTmp) Then Label2(10) = strTmp
            Label2(11) = pa(36)
         End If
   End Select
   
   m_CP25 = "": pa(1) = "": m_CP10 = ""
   If pa(25) = "" Then pa(25) = 0
   strExc(0) = "SELECT CP25,CP01,CP10 FROM CASEPROGRESS,CASEPROPERTYMAP WHERE CP09='" & strReceiveNo & "' AND CP01=CPM01(+) AND CP10=CPM02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If IsNull(RsTemp.Fields(0)) Then RsTemp.Fields(0) = pa(25)
      If RsTemp.Fields(0) < ChangeTStringToWString(pa(25)) Then
         m_CP25 = "N"  '專利未到期
      End If
      pa(1) = RsTemp.Fields(1)
      m_CP10 = RsTemp.Fields(2)
   End If
   
End Sub

Private Function ProcessPrint() As Boolean
Dim m_strSerialNo As String '請款單號
Dim strAgentNo As String '代理人編號
Dim strPrintCust  As String '是否列印申請人
Dim dblUSRate As Double '美金匯率
Dim strA1K27 As String '列印對象
Dim strA1K28 As String '請款對象
   
   On Error GoTo ErrorHandler
   
   ProcessPrint = False
   cnnConnection.BeginTrans
   
   '開始新增國外請款資料
   '1:先以"X"抓ACC1R0之國外請款單的自動編號, 並更新其流水號
   m_strSerialNo = AccAutoNo(MsgText(815), 5)
   AccSaveAutoNo MsgText(815), Right(m_strSerialNo, 5)
   '2:新增ACC1K0
'   strAgentNo = GetAgentNO '代理人編號
   strAgentNo = PUB_GetA1K03(pa(1), pa(2), pa(3), pa(4))
  ' dblUSRate = GetUSRate '美金匯率
   
    strA1K27 = PUB_GetA1K27(pa(1), pa(2), pa(3), pa(4), m_CP10)
    If strA1K27 = "" Then strA1K27 = strAgentNo
    strA1K28 = PUB_GetA1K28(pa(1), pa(2), pa(3), pa(4), m_CP10)
    If strA1K28 = "" Then strA1K28 = strAgentNo
    
'   strPrintCust = GetPrintCust '是否列印申請人
   'Modify by Morgan 2004/12/16 改規則
   'strPrintCust = PUB_GetA1K04(pa(1), pa(2), pa(3), pa(4))
   strPrintCust = PUB_GetA1K04(pa(1), pa(2), pa(3), pa(4), strA1K28, m_CP10)
   '2004/12/16 end
    
   'Added by Lydia 2014/12/15 請款單請改為依代理人或客戶檔設定的請款幣別
    Dim strA1K33 As String, strA1K18 As String
    'Modify By Sindy 2016/11/30
    'strA1K33 = PUB_GetInitCurrPrintType(pa(1), strA1K28, strA1K18, dblUSRate)
    'Modified by Morgan 2018/4/27 +strA1K27
    strA1K33 = PUB_GetInitCurrPrintType(pa(1), strA1K28, strA1K18, dblUSRate, pa(2), pa(3), pa(4), strA1K27)
    '2016/11/30 END
        
    Dim strDisc As String '折扣
    strDisc = 1 - (PUB_GetA1L07Disc(pa(1), pa(2), pa(3), pa(4), m_CP10, strSrvDate(2)) / 100)
    'Modify By Cheng 2004/01/07
    'A1K11要先扣除折扣才存檔
'   strSQL = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K18,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K19,A1K20,A1K21) " & _
'            "VALUES  ('" & m_strSerialNo & "'," & (ServerDate - 19110000) & ",0,NULL,0," & dblUSRate & "," & Val(m_A1J17) & ",NULL,'" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','USD',0, " & IIf(dblUSRate = 0, Val(m_A1J17), (Val(m_A1J17) / dblUSRate)) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "'," & (ServerDate - 19110000) & "," & ServerTime & ",'" & strUserNum & "')"
    'Modify By Cheng 2004/04/26
    '美金取整數位(無條件捨去)
'   strSQL = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K18,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K19,A1K20,A1K21) " & _
'            "VALUES  ('" & m_strSerialNo & "'," & (ServerDate - 19110000) & ",0,NULL,0," & dblUSRate & "," & Val(m_A1J17) - Val(Val(m_A1J17) * Val(strDisc)) & ",NULL,'" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','USD',0, " & IIf(dblUSRate = 0, (Val(m_A1J17) - Val(Val(m_A1J17) * Val(strDisc))), ((Val(m_A1J17) - Val(Val(m_A1J17) * Val(strDisc))) / dblUSRate)) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "'," & (ServerDate - 19110000) & "," & ServerTime & ",'" & strUserNum & "')"
   'Added by Lydia 2014/12/15 請款單請改為依代理人或客戶檔設定的請款幣別
'   strSql = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K18,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K19,A1K20,A1K21) " & _
            "VALUES  ('" & m_strSerialNo & "'," & (ServerDate - 19110000) & ",0,NULL,0," & dblUSRate & "," & Val(m_A1J17) - Val(Val(m_A1J17) * Val(strDisc)) & ",NULL,'" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','USD',0, " & Fix(Val("" & IIf(dblUSRate = 0, (Val(m_A1J17) - Val(Val(m_A1J17) * Val(strDisc))), ((Val(m_A1J17) - Val(Val(m_A1J17) * Val(strDisc))) / dblUSRate)))) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "'," & (ServerDate - 19110000) & "," & ServerTime & ",'" & strUserNum & "')"
    strSql = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K18,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K19,A1K20,A1K21,A1K33) " & _
            "VALUES  ('" & m_strSerialNo & "'," & (ServerDate - 19110000) & ",0,NULL,0," & dblUSRate & "," & Val(m_A1J17) - Val(Val(m_A1J17) * Val(strDisc)) & ",NULL,'" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','" & strA1K18 & "',0, " & Fix(Val("" & IIf(dblUSRate = 0, (Val(m_A1J17) - Val(Val(m_A1J17) * Val(strDisc))), ((Val(m_A1J17) - Val(Val(m_A1J17) * Val(strDisc))) / dblUSRate)))) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "'," & (ServerDate - 19110000) & "," & ServerTime & ",'" & strUserNum & "','" & strA1K33 & "')"
   
    'End
   cnnConnection.Execute strSql
   '3:新增ACC1L0
'    Dim strDisc As String '折扣
'    strDisc = 1 - (PUB_GetA1L07Disc(pa(1), pa(2), pa(3), pa(4), m_CP10, strSrvDate(2)) / 100)
   strSql = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L06,A1L07,A1L02,A1L04,A1L05,A1L08,A1L09,A1L10) " & _
            "VALUES  ('" & m_strSerialNo & "','FCP',''," & Val(m_A1J17) * Val(strDisc) & ",'001','" & m_CP10 & "'," & Val(m_A1J17) & "," & (ServerDate - 19110000) & "," & ServerTime & ",'" & strUserNum & "')"
   cnnConnection.Execute strSql
   
   PUB_UpdateA1k08 m_strSerialNo 'Added by Morgan 2012/11/2 更新請款單外幣金額
   
   '4:新增ACC1W0
   strSql = "INSERT INTO ACC1W0 " & _
            "VALUES  ('" & m_strSerialNo & "','" & strReceiveNo & "')"
   cnnConnection.Execute strSql
   '5:更新請款單號
   strSql = "UPDATE CASEPROGRESS SET CP16='" & m_A1J17 & "',CP18='" & Round(m_A1J17 / 1000) & "',CP60='" & m_strSerialNo & "' WHERE CP09='" & strReceiveNo & "'"
   cnnConnection.Execute strSql
   
   PUB_PointAutoassign m_strSerialNo, True 'Add by Morgan 2010/4/21 自動分配點數
   
   cnnConnection.CommitTrans
    ProcessPrint = True
         
    'Added by Lydia 2016/11/17 以請款對象檢查是否存在於國外固定寄催款單代理人檔(ACC225)且下次寄發日期＞系統日，若存在則顯示訊息提醒操作人員
    If m_strSerialNo <> "" And strA1K28 <> "" Then
       If PUB_ChkAcc225MsgList(m_strSerialNo, strA1K28, pa(1), pa(2), pa(3), pa(4)) Then
       End If
    End If
    'end 2016/11/17
         
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
   Exit Function

ErrorHandler:
      
      cnnConnection.RollbackTrans
      ProcessPrint = False
End Function

Private Function GetUSRate() As Double
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetUSRate = 0
'strSQLA = "SELECT USXR02 FROM USXRATE WHERE USXR01<=" & (ServerDate - 19110000) & " AND ROWNUM = 1 ORDER BY USXR01 "
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
