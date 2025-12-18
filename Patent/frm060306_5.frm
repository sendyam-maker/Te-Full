VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060306_5 
   BorderStyle     =   1  '單線固定
   Caption         =   "請款通知函-結案請款函"
   ClientHeight    =   5016
   ClientLeft      =   1440
   ClientTop       =   1056
   ClientWidth     =   7032
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5016
   ScaleWidth      =   7032
   Begin VB.TextBox Text3 
      Height          =   264
      Left            =   4530
      MaxLength       =   1
      TabIndex        =   1
      Top             =   4710
      Width           =   255
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   " 回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4770
      TabIndex        =   3
      Top             =   24
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   2400
      Left            =   0
      TabIndex        =   5
      Top             =   435
      Width           =   6975
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   7
         Left            =   1260
         TabIndex        =   31
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
         Left            =   1260
         TabIndex        =   30
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
         Left            =   1260
         TabIndex        =   29
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
         Left            =   1260
         TabIndex        =   28
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
         Left            =   1260
         TabIndex        =   27
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
         Left            =   1260
         TabIndex        =   26
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
         Left            =   1260
         TabIndex        =   25
         Top             =   375
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
         Left            =   1260
         TabIndex        =   24
         Top             =   120
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
         Left            =   1260
         TabIndex        =   22
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
         Caption         =   "本所案號"
         Height          =   180
         Index           =   0
         Left            =   84
         TabIndex        =   14
         Top             =   132
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "請款函日期"
         Height          =   180
         Index           =   1
         Left            =   84
         TabIndex        =   13
         Top             =   372
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利名稱"
         Height          =   180
         Index           =   2
         Left            =   84
         TabIndex        =   12
         Top             =   612
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(中)"
         Height          =   180
         Index           =   3
         Left            =   84
         TabIndex        =   11
         Top             =   972
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(日)"
         Height          =   180
         Index           =   4
         Left            =   84
         TabIndex        =   10
         Top             =   1452
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(英)"
         Height          =   180
         Index           =   5
         Left            =   84
         TabIndex        =   9
         Top             =   1212
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(中)"
         Height          =   180
         Index           =   6
         Left            =   84
         TabIndex        =   8
         Top             =   1692
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(日)"
         Height          =   180
         Index           =   7
         Left            =   84
         TabIndex        =   7
         Top             =   2172
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(英)"
         Height          =   180
         Index           =   8
         Left            =   84
         TabIndex        =   6
         Top             =   1932
         Width           =   936
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3990
      TabIndex        =   2
      Top             =   24
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   5976
      TabIndex        =   4
      Top             =   24
      Width           =   800
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   0
      Top             =   4710
      Width           =   255
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   11
      Left            =   1290
      TabIndex        =   35
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
      Left            =   1290
      TabIndex        =   34
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
      Left            =   1290
      TabIndex        =   33
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
      Left            =   1290
      TabIndex        =   32
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
      Height          =   840
      Index           =   0
      Left            =   1260
      TabIndex        =   23
      Top             =   3810
      Width           =   5715
      VariousPropertyBits=   -1466939365
      MaxLength       =   2000
      ScrollBars      =   3
      Size            =   "10081;1482"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否獲指示不續辦        (Y)"
      Height          =   180
      Index           =   15
      Left            =   3030
      TabIndex        =   21
      Top             =   4710
      Width           =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶案件案號"
      Height          =   180
      Index           =   9
      Left            =   90
      TabIndex        =   20
      Top             =   2880
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "彼此案號"
      Height          =   180
      Index           =   10
      Left            =   90
      TabIndex        =   19
      Top             =   3120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "P. S."
      Height          =   180
      Index           =   11
      Left            =   90
      TabIndex        =   18
      Top             =   3840
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "副本收受人"
      Height          =   180
      Index           =   12
      Left            =   90
      TabIndex        =   17
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "副本聯絡人"
      Height          =   180
      Index           =   13
      Left            =   90
      TabIndex        =   16
      Top             =   3600
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否修改請款函        (Y)"
      Height          =   180
      Index           =   14
      Left            =   90
      TabIndex        =   15
      Top             =   4710
      Width           =   1860
   End
End
Attribute VB_Name = "frm060306_5"
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
Dim m_CP10 As String '案件性質
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String
Dim m_LetterLanguage As String 'Add By Sindy 2015/9/21

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
Dim strTxt(1 To 2) As String, i As Integer, j As Integer, strTmp As String
   EndLetter ET01, strReceiveNo, ET03, strUserNum
    j = 1
    'Add By Cheng 2003/02/24
   '請款函日期
   If frm060306.Text5.Text <> "" Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','請款函日期','" & DBDATE(frm060306.Text5.Text) & "')"
      j = j + 1
   End If
   '請款函備註
    'Modify By Cheng 2003/02/26
'   If Text1(j).Text <> "" Then
   If Text1(0).Text <> "" Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','請款函備註','P.S. " & ChgSQL(Text1(0).Text) & "')"
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

'Modify by Amy 2025/05/28 原:Private
Public Sub cmdok_Click(Index As Integer)
   Dim bolChk As Boolean
   Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
   Dim stET03 As String
    
    Select Case Index
    Case 2 '結束
        Unload frm060306
        Unload Me
    Case 0 '確定
        Screen.MousePointer = vbHourglass
        m_LetterLanguage = PUB_GetLanguage(m_CP01, m_CP02, m_CP03, m_CP04) 'Add By Sindy 2015/9/21
        If Text2.Text = "Y" Then bolChk = True
        '請再區分英文12或日文13
        '若案件性質為不續辦
        If m_CP10 = "907" Then
            If Me.Text3.Text = "Y" Then
               stET03 = "01"
            Else
               stET03 = "00"
            End If
        '其他案件性質
        Else
            stET03 = "12"
        End If
         StartLetter ET01, stET03
         'Add by Morgan 2008/3/31 判斷是否產生電子檔
         bolEmail = PUB_GetEMailFlag(m_CP01 & m_CP02 & m_CP03 & m_CP04, , , bolPlusPaper)
         'Add by Morgan 2009/10/20 +判斷是否EMail同時寄紙本
         If bolPlusPaper Then
            iCopy = 0
         Else
            iCopy = 1
         End If
         'end 2009/10/20
         If bolEmail Then
            NowPrint strReceiveNo, ET01, stET03, bolChk, strUserNum, , , , , iCopy, , True, True
            If bolChk = False Then
               MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_CP01) & " ]！"
            End If
         Else
         'end 2008/3/31
            NowPrint strReceiveNo, ET01, stET03, bolChk, strUserNum, 0
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
    Case 1 '回前畫面
        frm060306.Show
        Unload Me
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060306_5 = Nothing
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

Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
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
    'Add By Cheng 2003/10/15
    '取得案件性質
    StrSQLa = "Select * From CaseProgress Where CP09='" & strReceiveNo & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        m_CP10 = "" & rsA("CP10").Value
    Else
        m_CP10 = ""
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    '若案件性質為不續辦(907)
    If m_CP10 = "907" Then
        Me.Label1(15).Visible = True
        Me.Text3.Visible = True
    '其他案件性質
    Else
        Me.Label1(15).Visible = False
        Me.Text3.Visible = False
    End If
    'End
End Sub

Private Sub Text3_GotFocus()
    TextInverse Me.Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> 89 Then
        KeyAscii = 0
    End If
End Sub
