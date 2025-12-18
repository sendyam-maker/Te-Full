VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060306_8 
   BorderStyle     =   1  '單線固定
   Caption         =   "請款通知函-優先審查請款函"
   ClientHeight    =   5280
   ClientLeft      =   750
   ClientTop       =   1545
   ClientWidth     =   7140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7140
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1590
      MaxLength       =   1
      TabIndex        =   1
      Top             =   4380
      Width           =   255
   End
   Begin VB.TextBox text4 
      Height          =   270
      Left            =   2310
      MaxLength       =   1
      TabIndex        =   2
      Top             =   4665
      Width           =   255
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   " 回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4890
      TabIndex        =   5
      Top             =   36
      Width           =   1155
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1590
      MaxLength       =   1
      TabIndex        =   3
      Top             =   4950
      Width           =   255
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6084
      TabIndex        =   6
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
      TabIndex        =   4
      Top             =   36
      Width           =   756
   End
   Begin VB.Frame Frame1 
      Height          =   2400
      Left            =   105
      TabIndex        =   7
      Top             =   468
      Width           =   6975
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   7
         Left            =   1200
         TabIndex        =   33
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
         Left            =   1200
         TabIndex        =   32
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
         Left            =   1200
         TabIndex        =   31
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
         Left            =   1200
         TabIndex        =   30
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
         Left            =   1200
         TabIndex        =   29
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
         Left            =   1200
         TabIndex        =   28
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
         Left            =   1200
         TabIndex        =   27
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
         Left            =   1200
         TabIndex        =   26
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
         Left            =   1200
         TabIndex        =   25
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
         TabIndex        =   16
         Top             =   1932
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(日)"
         Height          =   180
         Index           =   7
         Left            =   48
         TabIndex        =   15
         Top             =   2172
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(中)"
         Height          =   180
         Index           =   6
         Left            =   48
         TabIndex        =   14
         Top             =   1692
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(英)"
         Height          =   180
         Index           =   5
         Left            =   48
         TabIndex        =   13
         Top             =   1212
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(日)"
         Height          =   180
         Index           =   4
         Left            =   48
         TabIndex        =   12
         Top             =   1452
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(中)"
         Height          =   180
         Index           =   3
         Left            =   48
         TabIndex        =   11
         Top             =   972
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利名稱"
         Height          =   180
         Index           =   2
         Left            =   48
         TabIndex        =   10
         Top             =   612
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "請款函日期"
         Height          =   180
         Index           =   1
         Left            =   48
         TabIndex        =   9
         Top             =   372
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號"
         Height          =   180
         Index           =   0
         Left            =   48
         TabIndex        =   8
         Top             =   132
         Width           =   720
      End
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   11
      Left            =   1350
      TabIndex        =   37
      Top             =   3600
      Width           =   5715
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10081;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   10
      Left            =   1350
      TabIndex        =   36
      Top             =   3360
      Width           =   5715
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10081;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   9
      Left            =   1350
      TabIndex        =   35
      Top             =   3120
      Width           =   5715
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10081;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   8
      Left            =   1350
      TabIndex        =   34
      Top             =   2880
      Width           =   5715
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10081;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   540
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   3810
      Width           =   5745
      VariousPropertyBits=   -1466939365
      MaxLength       =   2000
      ScrollBars      =   3
      Size            =   "10134;952"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否含主動修正：　　(Y)"
      Height          =   180
      Index           =   15
      Left            =   135
      TabIndex        =   24
      Top             =   4410
      Width           =   2040
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "是否與實體審查同函請款：　　(Y)"
      Height          =   180
      Index           =   9
      Left            =   120
      TabIndex        =   23
      Top             =   4710
      Width           =   2760
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否修改請款函：　　(Y)"
      Height          =   180
      Index           =   14
      Left            =   135
      TabIndex        =   22
      Top             =   4980
      Width           =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "副本聯絡人"
      Height          =   180
      Index           =   13
      Left            =   168
      TabIndex        =   21
      Top             =   3600
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "副本收受人"
      Height          =   180
      Index           =   12
      Left            =   168
      TabIndex        =   20
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "P. S."
      Height          =   180
      Index           =   11
      Left            =   180
      TabIndex        =   19
      Top             =   3780
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "彼此案號"
      Height          =   180
      Index           =   10
      Left            =   168
      TabIndex        =   18
      Top             =   3120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶案件案號"
      Height          =   180
      Index           =   9
      Left            =   168
      TabIndex        =   17
      Top             =   2880
      Width           =   1080
   End
End
Attribute VB_Name = "frm060306_8"
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
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String

Dim m_bolContinue As Boolean '是否繼續
Const ET01 As String = "09"  '定稿別
'Add by Morgan 2004/12/29
Dim m_iLanguage As Integer '定稿語文
Dim m_bolEmail As Boolean, m_bolPlusPaper As Boolean, m_iCopy As Integer


Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)

   Dim strTxt(1 To 3) As String, i As Integer, j As Integer, strTmp As String
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   i = 1
   If frm060306.Text5.Text <> "" Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','請款函日期','" & DBDATE(frm060306.Text5.Text) & "')"
      i = i + 1
   End If
   If Text1(0).Text <> "" Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','請款函備註','P.S. " & ChgSQL(Text1(0).Text) & "')"
      i = i + 1
   End If

   If Text3 = "Y" Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','主動修正','    As per your instructions, we have also filed voluntary amendments of the invention patent application. A copy of amendments is enclosed." & Chr(13) & "')"
      i = i + 1
   End If
   
   If i <> 1 Then
       'edit by nickc 2007/02/05 不用 dll 了
       'If Not objLawDll.ExecSQL(i - 1, strTxt) Then
       If Not ClsLawExecSQL(i - 1, strTxt) Then
           MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
       End If
   End If
End Sub

Private Sub Process()
   Dim bolEdit As Boolean '是否修改請款函
   Dim stET03 As String
   Dim stLanguage As String '定稿語文
   
   If Text2.Text = "Y" Then bolEdit = True
   'Modify by Morgan 2006/6/2
   'm_iLanguage = GetLetterLanguage(pA(1), pA(2), pA(3), pA(4))
   m_iLanguage = PUB_GetLanguage(pa(1), pa(2), pa(3), pa(4))
   Select Case m_iLanguage
      Case 3 '日文
         stET03 = "03"
         If Text4 = "Y" Then
            stET03 = "04"
         End If
      
      Case Else '英文
         stET03 = "00"
         '與實體審查同函請款
         If Text4 = "Y" Then
            stET03 = "01"
         End If
   End Select
   
   StartLetter ET01, stET03
   'Add by Morgan 2008/3/31 判斷是否產生電子檔
   m_bolEmail = PUB_GetEMailFlag(pa(1) & pa(2) & pa(3) & pa(4), , , m_bolPlusPaper)
   'Add by Morgan 2009/10/20 +判斷是否EMail同時寄紙本
   If m_bolPlusPaper Then
      m_iCopy = 0
   Else
      m_iCopy = 1
   End If
   'end 2009/10/20
   If m_bolEmail Then
      NowPrint strReceiveNo, ET01, stET03, bolEdit, strUserNum, , , , , m_iCopy, , True, True
      If bolEdit = False Then
         MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(pa(1)) & " ]！"
      End If
   Else
      NowPrint strReceiveNo, ET01, stET03, bolEdit, strUserNum
   End If
   
   If Not m_bolEmail Or m_bolPlusPaper Then
'      'Add By Sindy 2015/9/21 日文定稿才要印地址條
'      If m_iLanguage = "3" Or Val(外專開窗信函啟用日) >= Val(strSrvDate(1)) Then
'      '2015/9/21 END
      'Add By Sindy 2017/3/20 日文定稿才要印地址條
      If frm060306.m_FCna01 = "101" Or m_iLanguage = "3" Then '美國 或 日文定稿才要印地址條
      '2017/3/20 END
         '新增地址條列表資料
         pub_AddressListSN = pub_AddressListSN + 1
         PUB_AddNewAddressList strUserNum, pa(1), pa(2), pa(3), pa(4), "" & pub_AddressListSN, "0"
      End If
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0 '確定
         Screen.MousePointer = vbHourglass
         Process
         Screen.MousePointer = vbDefault
         frm060306.Show
         frm060306.Clear
      Case 1 '回前畫面
         frm060306.Show
      Case 2 '結束
         Unload frm060306
   End Select
   Unload Me
End Sub

Private Sub Form_Activate()
   If m_bolContinue = False Then
      Unload Me
   End If
End Sub

Private Sub Form_Initialize()
   'add by nickc 2007/02/02
   ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   InverseTextBox Text1(Index)
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

   Dim Lbl As Object, i As Integer, strTmp As String
   For Each Lbl In Label2
      Lbl = ""
   Next
   m_bolContinue = True
   strReceiveNo = frm060306.Tag '總收文號
   strExc(0) = "SELECT CP01,CP02,CP03,CP04,CP10 FROM CASEPROGRESS,CASEPROPERTYMAP WHERE CP09='" & strReceiveNo & "' AND CP01=CPM01(+) AND CP10=CPM02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      pa(1) = "" & RsTemp.Fields("CP01")
      pa(2) = "" & RsTemp.Fields("CP02")
      pa(3) = "" & RsTemp.Fields("CP03")
      pa(4) = "" & RsTemp.Fields("CP04")
   Else
      m_bolContinue = False
      MsgBox "無法讀取收文資料！", vbCritical
      Exit Sub
   End If
   '檢查是否有實體審查已發文且未請款
   strExc(0) = "SELECT 1 FROM CASEPROGRESS WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10='416' AND CP27 IS NOT NULL AND CP57 IS NULL AND CP60 IS NULL"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Text4 = "Y"
   End If
   
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
   
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text4_GotFocus()
   InverseTextBox Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
