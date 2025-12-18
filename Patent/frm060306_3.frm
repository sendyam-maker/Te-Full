VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060306_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "請款通知函-讓與、合併、繼承、授權、變更、自撤、請款函"
   ClientHeight    =   5160
   ClientLeft      =   915
   ClientTop       =   990
   ClientWidth     =   7035
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   7035
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   " 回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4770
      TabIndex        =   2
      Top             =   48
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   2472
      Left            =   12
      TabIndex        =   4
      Top             =   516
      Width           =   6975
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   7
         Left            =   1260
         TabIndex        =   30
         Top             =   2205
         Width           =   5535
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9763;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   6
         Left            =   1260
         TabIndex        =   29
         Top             =   1965
         Width           =   5535
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9763;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   5
         Left            =   1260
         TabIndex        =   28
         Top             =   1740
         Width           =   5535
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9763;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   4
         Left            =   1260
         TabIndex        =   27
         Top             =   1500
         Width           =   5535
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9763;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   3
         Left            =   1260
         TabIndex        =   26
         Top             =   1245
         Width           =   5535
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9763;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   2
         Left            =   1260
         TabIndex        =   25
         Top             =   1005
         Width           =   5535
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9763;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   1
         Left            =   1260
         TabIndex        =   24
         Top             =   405
         Width           =   5535
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9763;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   0
         Left            =   1260
         TabIndex        =   23
         Top             =   150
         Width           =   5535
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9763;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   285
         Left            =   1260
         TabIndex        =   21
         Top             =   630
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
         Left            =   84
         TabIndex        =   13
         Top             =   1968
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(日)"
         Height          =   180
         Index           =   7
         Left            =   84
         TabIndex        =   12
         Top             =   2208
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(中)"
         Height          =   180
         Index           =   6
         Left            =   84
         TabIndex        =   11
         Top             =   1728
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(英)"
         Height          =   180
         Index           =   5
         Left            =   84
         TabIndex        =   10
         Top             =   1248
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(日)"
         Height          =   180
         Index           =   4
         Left            =   84
         TabIndex        =   9
         Top             =   1488
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(中)"
         Height          =   180
         Index           =   3
         Left            =   84
         TabIndex        =   8
         Top             =   1008
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利名稱"
         Height          =   180
         Index           =   2
         Left            =   84
         TabIndex        =   7
         Top             =   648
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "請款函日期"
         Height          =   180
         Index           =   1
         Left            =   84
         TabIndex        =   6
         Top             =   408
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號"
         Height          =   180
         Index           =   0
         Left            =   84
         TabIndex        =   5
         Top             =   168
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3990
      TabIndex        =   1
      Top             =   48
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   5976
      TabIndex        =   3
      Top             =   48
      Width           =   800
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Left            =   1404
      MaxLength       =   1
      TabIndex        =   0
      Top             =   4836
      Width           =   255
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   12
      Left            =   1290
      TabIndex        =   35
      Top             =   3990
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
      Left            =   1290
      TabIndex        =   34
      Top             =   3750
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
      Left            =   1290
      TabIndex        =   33
      Top             =   3525
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
      Left            =   1290
      TabIndex        =   32
      Top             =   3285
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
      Left            =   1290
      TabIndex        =   31
      Top             =   3060
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
      Height          =   600
      Index           =   0
      Left            =   1260
      TabIndex        =   22
      Top             =   4200
      Width           =   5715
      VariousPropertyBits=   -1466939365
      MaxLength       =   2000
      ScrollBars      =   3
      Size            =   "10081;1058"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質"
      Height          =   180
      Index           =   15
      Left            =   96
      TabIndex        =   20
      Top             =   4020
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶案件案號"
      Height          =   180
      Index           =   9
      Left            =   96
      TabIndex        =   19
      Top             =   3060
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "彼此案號"
      Height          =   180
      Index           =   10
      Left            =   96
      TabIndex        =   18
      Top             =   3300
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "P. S."
      Height          =   180
      Index           =   11
      Left            =   96
      TabIndex        =   17
      Top             =   4260
      Width           =   312
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "副本收受人"
      Height          =   180
      Index           =   12
      Left            =   96
      TabIndex        =   16
      Top             =   3540
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "副本聯絡人"
      Height          =   180
      Index           =   13
      Left            =   96
      TabIndex        =   15
      Top             =   3780
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否修改請款函        (Y)"
      Height          =   180
      Index           =   14
      Left            =   96
      TabIndex        =   14
      Top             =   4836
      Width           =   1860
   End
End
Attribute VB_Name = "frm060306_3"
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
Const ET01 As String = "09"
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String
Dim m_LetterLanguage As String 'Add By Sindy 2015/9/21


Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
 Dim strTxt(1 To 2) As String, i As Integer, j As Integer, strTmp As String
   EndLetter ET01, strReceiveNo, ET03, strUserNum
    'Add By Cheng 2003/02/24
    j = 1
   '請款函日期
   If frm060306.Text5.Text <> "" Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','請款函日期','" & DBDATE(frm060306.Text5.Text) & "')"
      j = j + 1
   End If
   '請款函備註
   If Text1(0).Text <> "" Then
        'Modify By Cheng 2003/12/23
        '避免單引號發生的錯誤
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','請款函備註','P.S. " & Text1(0).Text & "')"
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','請款函備註','P.S. " & ChgSQL(Text1(0).Text) & "')"
        'End
      j = j + 1
   End If
    If j <> 1 Then
        'edit by nickc 2007/02/05 不用 dll 了
        'If Not objLawDll.ExecSQL(j - 1, strTxt) Then
        If Not ClsLawExecSQL(j - 1, strTxt) Then
           MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
        End If
    End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Dim bolChk As Boolean
   Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
   
   Select Case Index
      Case 1
         frm060306.Show
         Unload Me
      Case 0
         Screen.MousePointer = vbHourglass
         m_LetterLanguage = PUB_GetLanguage(m_CP01, m_CP02, m_CP03, m_CP04) 'Add By Sindy 2015/9/21
         If Text2.Text = "Y" Then bolChk = True
         '請再區分英文02或日文03
         StartLetter ET01, "02"
         'Add by Morgan 2008/3/24 判斷是否產生電子檔
         bolEmail = PUB_GetEMailFlag(m_CP01 & m_CP02 & m_CP03 & m_CP04, , , bolPlusPaper)
         'Add by Morgan 2009/10/20 +判斷是否EMail同時寄紙本
         If bolPlusPaper Then
            iCopy = 0
         Else
            iCopy = 1
         End If
         'end 2009/10/20
         If bolEmail Then
            NowPrint strReceiveNo, ET01, "02", bolChk, strUserNum, , , , , iCopy, , True, True
            If bolChk = False Then
               MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_CP01) & " ]！"
            End If
         Else
         'end 2008/3/31
            NowPrint strReceiveNo, ET01, "02", bolChk, strUserNum
         End If
         
         If Not bolEmail Or bolPlusPaper Then
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
         
         frm060306.Show
         frm060306.Clear
        Screen.MousePointer = vbDefault
         Unload Me
      Case 2
         Unload frm060306
         Unload Me
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060306_3 = Nothing
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   ReadPatent
End Sub

Private Sub ReadPatent()
 'edit by nickc 2007/02/02
 'Dim lbl As Label, pA(1 To T_PA) As String, i As Integer, strTmp As String
 Dim Lbl As Object, pa() As String, i As Integer, strTmp As String
 'add by nickc 2007/02/02
 ReDim pa(1 To TF_PA) As String
 
   For Each Lbl In Label2
      Lbl = ""
   Next
   strReceiveNo = frm060306.Tag
   pa(1) = frm060306.Text1.Text
   pa(2) = frm060306.Text2.Text
   pa(3) = frm060306.Text3.Text
   pa(4) = frm060306.Text4.Text
   
   m_CP01 = pa(1): m_CP02 = pa(2): m_CP03 = pa(3): m_CP04 = pa(4) 'Add by Morgan 2008/3/31
   
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
   
   strExc(0) = "SELECT CPM03 FROM CASEPROGRESS,CASEPROPERTYMAP WHERE CP09='" & strReceiveNo & "' AND CP01=CPM01(+) AND CP10=CPM02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then Label2(12) = RsTemp.Fields(0)
   
End Sub
