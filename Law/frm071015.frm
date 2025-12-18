VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm071015 
   BorderStyle     =   1  '單線固定
   Caption         =   "其他來函"
   ClientHeight    =   6192
   ClientLeft      =   456
   ClientTop       =   504
   ClientWidth     =   8892
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6192
   ScaleWidth      =   8892
   Begin VB.CommandButton cmdEnd 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7908
      TabIndex        =   45
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6780
      TabIndex        =   44
      Top             =   70
      Width           =   1100
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5952
      TabIndex        =   43
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOther 
      Caption         =   "相關卷號(&F)"
      Height          =   400
      Left            =   4824
      TabIndex        =   42
      Top             =   70
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   4524
      Left            =   180
      TabIndex        =   24
      Top             =   1608
      Width           =   8580
      Begin VB.TextBox TextCF30 
         Height          =   288
         Left            =   7290
         TabIndex        =   4
         Top             =   508
         Width           =   405
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1092
         Left            =   72
         TabIndex        =   11
         Top             =   2424
         Width           =   8292
         _ExtentX        =   14626
         _ExtentY        =   1926
         _Version        =   393216
         Cols            =   12
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   12
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.ComboBox cboGov 
         Height          =   324
         Left            =   1080
         TabIndex        =   5
         Top             =   816
         Width           =   2844
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "5016;572"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   288
         Index           =   9
         Left            =   4920
         TabIndex        =   10
         Top             =   1824
         Width           =   492
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "868;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   288
         Index           =   6
         Left            =   1080
         TabIndex        =   7
         Top             =   1492
         Width           =   855
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1508;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   288
         Index           =   2
         Left            =   1080
         TabIndex        =   2
         Top             =   508
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1931;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   288
         Index           =   0
         Left            =   1080
         TabIndex        =   0
         Top             =   180
         Width           =   615
         VariousPropertyBits=   671105051
         MaxLength       =   4
         Size            =   "1085;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   288
         Index           =   5
         Left            =   1080
         TabIndex        =   6
         Top             =   1164
         Width           =   6495
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "11456;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   288
         Index           =   7
         Left            =   4920
         TabIndex        =   8
         Top             =   1492
         Width           =   975
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1720;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   288
         Index           =   8
         Left            =   1080
         TabIndex        =   9
         Top             =   1824
         Width           =   852
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1503;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   288
         Index           =   3
         Left            =   3840
         TabIndex        =   3
         Top             =   508
         Width           =   972
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1714;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   588
         Index           =   10
         Left            =   96
         TabIndex        =   12
         Top             =   3840
         Width           =   8340
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "14711;1032"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   288
         Index           =   1
         Left            =   4920
         TabIndex        =   1
         Top             =   180
         Width           =   735
         VariousPropertyBits=   671105051
         MaxLength       =   4
         Size            =   "1296;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCF30 
         Caption         =   "是否自動內部收文下一程序：            (Y:是)"
         Height          =   195
         Left            =   4920
         TabIndex        =   46
         Top             =   555
         Width           =   3585
      End
      Begin VB.Label Label29 
         Caption         =   "(Y:閉卷)"
         Height          =   180
         Left            =   5520
         TabIndex        =   41
         Top             =   1878
         Width           =   852
      End
      Begin VB.Label Label21 
         Caption         =   "是否閉卷："
         Height          =   180
         Index           =   1
         Left            =   3960
         TabIndex        =   40
         Top             =   1878
         Width           =   972
      End
      Begin VB.Label Label19 
         Caption         =   "本案期限："
         Height          =   204
         Left            =   120
         TabIndex        =   39
         Top             =   2184
         Width           =   996
      End
      Begin MSForms.Label lbe 
         Height          =   288
         Index           =   0
         Left            =   1800
         TabIndex        =   38
         Top             =   180
         Width           =   2052
         VariousPropertyBits=   27
         Size            =   "3619;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label14 
         Caption         =   "來函性質："
         Height          =   180
         Left            =   120
         TabIndex        =   37
         Top             =   234
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "承  辦  人："
         Height          =   180
         Left            =   120
         TabIndex        =   36
         Top             =   1546
         Width           =   975
      End
      Begin MSForms.Label lbe 
         Height          =   288
         Index           =   8
         Left            =   2040
         TabIndex        =   35
         Top             =   1824
         Width           =   1812
         VariousPropertyBits=   27
         Size            =   "3196;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbe 
         Height          =   288
         Index           =   6
         Left            =   2040
         TabIndex        =   34
         Top             =   1492
         Width           =   1812
         VariousPropertyBits=   27
         Size            =   "3196;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label18 
         Caption         =   "進度備註："
         Height          =   204
         Left            =   96
         TabIndex        =   33
         Top             =   3624
         Width           =   1092
      End
      Begin VB.Label Label12 
         Caption         =   "機關文號："
         Height          =   180
         Left            =   120
         TabIndex        =   32
         Top             =   1218
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "協辦人員："
         Height          =   180
         Left            =   120
         TabIndex        =   31
         Top             =   1878
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "承辦期限："
         Height          =   180
         Left            =   3960
         TabIndex        =   30
         Top             =   1546
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "法定期限："
         Height          =   180
         Left            =   2880
         TabIndex        =   29
         Top             =   562
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "下一程序："
         Height          =   180
         Left            =   3960
         TabIndex        =   28
         Top             =   234
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "本所期限："
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   562
         Width           =   975
      End
      Begin MSForms.Label lbe 
         Height          =   288
         Index           =   1
         Left            =   5760
         TabIndex        =   26
         Top             =   180
         Width           =   2652
         VariousPropertyBits=   27
         Size            =   "4678;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label28 
         Caption         =   "機關代號："
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   906
         Width           =   975
      End
   End
   Begin MSForms.Label lbeCusName 
      Height          =   288
      Left            =   2520
      TabIndex        =   47
      Top             =   1348
      Width           =   6375
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "11245;508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lbeCus 
      Height          =   285
      Left            =   1368
      TabIndex        =   23
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "當  事  人："
      Height          =   180
      Left            =   204
      TabIndex        =   22
      Top             =   1402
      Width           =   972
   End
   Begin VB.Label Label21 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   0
      Left            =   4080
      TabIndex        =   21
      Top             =   582
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "收  文  號："
      Height          =   180
      Left            =   200
      TabIndex        =   20
      Top             =   582
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號："
      Height          =   180
      Left            =   200
      TabIndex        =   19
      Top             =   855
      Width           =   972
   End
   Begin VB.Label Label5 
      Caption         =   "來函收文日："
      Height          =   180
      Left            =   200
      TabIndex        =   18
      Top             =   1128
      Width           =   1092
   End
   Begin VB.Label lbeNum 
      Height          =   288
      Left            =   1368
      TabIndex        =   17
      Top             =   528
      Width           =   1932
   End
   Begin VB.Label lbeAccept 
      Height          =   288
      Left            =   1368
      TabIndex        =   16
      Top             =   1076
      Width           =   1572
   End
   Begin VB.Label lbeCaseNum 
      Height          =   288
      Left            =   1368
      TabIndex        =   15
      Top             =   802
      Width           =   1572
   End
   Begin VB.Label lbeProperty 
      Height          =   288
      Left            =   5160
      TabIndex        =   14
      Top             =   528
      Width           =   612
   End
   Begin VB.Label lbePropertyName 
      Height          =   288
      Left            =   5880
      TabIndex        =   13
      Top             =   528
      Width           =   2820
   End
End
Attribute VB_Name = "frm071015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/14 改成Form2.0 ; lbeCusName、MSHFlexGrid1改字型=新細明體-ExtB、Text(index)、lbe(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim PaperNum As String, t As Integer
Dim cp01 As String, cp02 As String, cp03 As String, cp04 As String, Strsale As String
Dim blnIsSave As Boolean, rsRecordset As New ADODB.Recordset, intPoint As Integer
Dim intLastRow As Integer, blnOKtoShow As Boolean, intCols As Integer, strOldLc As String
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Dim m_CP09 As String
Dim m_Nation As String
'Add By Sindy 2012/3/5
Dim m_bolClose As String
Dim m_strCloseDT As String
Dim m_strCloseReason As String
'2012/3/5 End
Dim m_GovListNew As String, m_GovListDef As String  'Added by Lydia 2025/11/18

Private Sub cmdBack_Click()
   If blnIsSave = False Then
      If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then: Exit Sub
   End If
   Unload Me
   frm071014.Show
End Sub

Private Sub cmdEnd_Click()
   If blnIsSave = False Then
      If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then: Exit Sub
   End If
   Unload frm071014
   Unload frm071015
End Sub

Private Sub cmdOther_Click()
  Dim i As Integer
   Dim strNum As String
   Dim strTmp As String
   If m_CP02 = "" Or m_CP02 = "" Then
      MsgBox "請輸入本所案號", vbInformation, "其他來函"
      Exit Sub
   End If
   Set frm1103_2.m_form = Me
   frm1103_2.intWhereComeFrom = 1
   frm1103_2.lblSystem = m_CP01
   frm1103_2.lblCode(0) = m_CP02
   frm1103_2.lblCode(1) = m_CP03
   frm1103_2.lblCode(2) = m_CP04
   frm1103_2.Show
   Me.Hide
End Sub

Private Sub Form_Activate()
   Text(0).SetFocus
End Sub

Private Sub Form_Load()
 Dim i As Integer, n As Integer
   MoveFormToCenter Me
   blnIsSave = False
   With frm071014.MSHFlexGrid1
      'PaperNum = ""
      For i = 1 To .Rows - 1
         .col = 0
         .row = i
         If .Text = "v" Then
            .col = 2
           ' PaperNum = "'" & .Text & "'"
            m_CP09 = .Text
            Exit For
         End If
      Next
   End With
   With frm071014
      m_CP01 = .txtcp01.Text
      m_CP02 = .txtcp02.Text
      If .txtcp03.Text <> "" Then
         m_CP03 = .txtcp03.Text
      Else
         m_CP03 = "0"
      End If
      If .txtcp04.Text <> "" Then
         m_CP04 = .txtcp04.Text
      Else
         m_CP04 = "0"
      End If
      
      cp01 = .txtcp01
      cp02 = .txtcp02
      cp03 = IIf(.txtcp03 = "", "0", .txtcp03)
      cp04 = IIf(.txtcp04 = "", "00", .txtcp04)
      
      If .txtcp03 = "" Then
         lbeCaseNum = cp01 + "-" + cp02
      Else
         If .txtcp04 = "" Then
            cp03 = .txtcp03
            lbeCaseNum = cp01 + "-" + cp02 + "-" + cp03 + "-" + "00"
         Else
            cp04 = .txtcp04
            lbeCaseNum = cp01 + "-" + cp02 + "-" + cp03 + "-" + cp04
         End If
      End If
      lbeAccept = ChangeTStringToTDateString(.txtAccept)
      lbeCus = .lbeCusNum
      lbeCusName = .lbeCusName
   End With
   
   GetData
   
   'add by sonia 2019/8/6
   If cp01 = "ACS" Then
      Me.Caption = "一般來函" 'Added by Lydia 2023/03/17
      'Modified by Lydia 2025/11/18 改成下拉選單
      'Text(4).Enabled = False
      'Text(4).Visible = False
      cboGov.Visible = False
      'end 2025/11/18
      Text(5).Enabled = False
      Text(5).Visible = False
      Text(8).Enabled = False
      Text(8).Visible = False
      Label28.Visible = False
      Label12.Visible = False
      Label7.Visible = False
   'Added by Lydia 2020/11/06
   Else
      lblCF30.Visible = False  '自動內部收文下一程序
      textCF30.Visible = False
      Label6.Left = Label9.Left
      Text(3).Left = Text(1).Left
   'end 2020/11/06
   End If
   'end 2019/8/6
End Sub

Private Sub GetData()
  Dim i As Integer
  Dim strName As String
   
   If cp01 <> "LA" Then 'lawcase
      'Add By Sindy 2012/3/5 +,lc09,lc10
      'Modified by Lydia 2025/11/18 +CP71
      strExc(1) = "select cp09,cp10,cp46,cp25,cp06,cp07,cp71,cp30,cp08,cp13," + _
         "cp14,cp48,cp29,cp64,lc15,lc27,lc08,lc09,lc10,CP71 from caseprogress,lawcase where " + _
         "cp09 = '" + m_CP09 + "' and CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04"
   Else 'hirecase
      'Add By Sindy 2012/3/5 +,hc10,hc11
      'Modified by Lydia 2025/11/18 +CP71
      strExc(1) = "select cp09,cp10,cp46,cp25,cp06,cp07,cp71,cp30,cp08,cp13," + _
         "cp14,cp48,cp29,cp64,' ',hc12,hc09,hc10,hc11,CP71 from caseprogress,hirecase where " + _
         "cp09 ='" + m_CP09 + "' and CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04"
   End If
   intI = 1
   'edit by nickc 2007/02/07 不用 dll 了
   'Set rsTemp = objLawDll.ReadRstMsg(intI, strExc(1))
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
   If intI = 1 Then
      lbeNum = RsTemp.Fields!CP09
      If Not IsNull(RsTemp.Fields!cp13) Then Strsale = RsTemp.Fields!cp13
      If Not IsNull(RsTemp.Fields!CP10) Then lbeProperty = RsTemp.Fields!CP10: lbePropertyName = ChgType(1, RsTemp.Fields!CP10)
      If m_CP01 <> "LA" Then
         If Not IsNull(RsTemp.Fields!lc15) Then
            m_Nation = RsTemp.Fields("LC15")
         End If
      ElseIf m_CP01 = "LA" Then
             m_Nation = "000"
      End If
      
      strName = ""
      '承辦人
      If Not IsNull(RsTemp.Fields("CP14")) Then
         Text(6).Text = RsTemp.Fields("CP14")
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetStaff(Text(6).Text, strName) Then
         If ClsPDGetStaff(Text(6).Text, strName) Then
            lbe(6).Caption = strName
         End If
      End If
      'Added by Lydia 2020/11/06 ACS案承辦人改為預設操作人員
      If m_CP01 = "ACS" Then
          If strUserNum <> Text(6).Text Then
                Text(6).Text = strUserNum
                If ClsPDGetStaff(Text(6).Text, strName) Then
                   lbe(6).Caption = strName
                End If
          End If
      End If
      'end 2020/11/06
      
      strName = ""
      '協辦人員
      If Not IsNull(RsTemp.Fields("CP29")) Then
         Text(8).Text = RsTemp.Fields("CP29")
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetStaff(Text(8).Text, strName) Then
         If ClsPDGetStaff(Text(8).Text, strName) Then
            lbe(8).Caption = strName
         End If
      End If
      
      'Add By Sindy 2012/3/5
      '是否閉卷
      m_bolClose = ""
      If Not IsNull(RsTemp.Fields(16)) Then
         Text(9).Text = RsTemp.Fields(16)
         m_bolClose = RsTemp.Fields(16)
      End If
      '閉卷日期
      m_strCloseDT = ""
      If Not IsNull(RsTemp.Fields(17)) Then
         m_strCloseDT = RsTemp.Fields(17)
      End If
      '閉卷原因
      m_strCloseReason = ""
      If Not IsNull(RsTemp.Fields(18)) Then
         m_strCloseReason = RsTemp.Fields(18)
      End If
      '2012/3/5 End
      'Added by Lydia 2025/11/18 改成下拉選選
      Call PUB_SetGovCmb(Me.cboGov, m_GovListNew, "" & RsTemp.Fields("cp71"))
      m_GovListDef = m_GovListNew
      'end 2025/11/18
   End If
   Getrs
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm071015 = Nothing
End Sub

Private Sub Text_Change(Index As Integer)
   Select Case Index
      Case 0, 1, 4, 6, 8
         If Text(Index) = "" Then lbe(Index) = ""
   End Select
End Sub

Private Sub Text_GotFocus(Index As Integer)
   TextInverse Text(Index)
   Select Case Index
          Case 5, 10
               'edit by nickc 2007/06/11  切換輸入法改用API
               'Text(Index).IMEMode = 1
               OpenIme
          Case Else
               'edit by nickc 2007/06/11  切換輸入法改用API
               'Text(Index).IMEMode = 2
               CloseIme
   End Select
End Sub

'Modified by Lydia 2021/09/14 改成Form 2.0
'Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 4, 5, 6, 8, 9
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub Text_LostFocus(Index As Integer)
   Select Case Index
          Case 5, 10
               'edit by nickc 2007/06/11  切換輸入法改用API
               'Text(Index).IMEMode = 2
               CloseIme
          Case 0
               If ChkUData(m_Nation, Text(0).Text) Then
               End If
   End Select

End Sub


Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
 Dim strTemp1 As String, strTemp2 As String
   Select Case Index
      Case 0, 1
         If Text(Index) <> "" Then
            'edit by nickc 2007/02/07 不用 dll 了
            'If objPublicData.GetCaseProperty(CP01, Text(Index), strTemp1, False) Then
            If ClsPDGetCaseProperty(cp01, Text(Index), strTemp1, False) Then
               lbe(Index) = strTemp1
            Else
               Cancel = True
            End If
            'add by sonia 2019/8/6
            If Index = 0 And Text(1) = "" Then
               'Modified by Lydia 2020/11/06 +CF30自動內部收文下一程序之設定
               'Text(1) = GetNextProgress(cp01, m_Nation, Text(Index))   '取得下一程序
               Text(1) = GetNextProgress(cp01, m_Nation, Text(Index), strExc(1))
               If textCF30.Visible = True Then textCF30 = strExc(1)
               'end 2020/11/06
               lbe(1) = GetCaseTypeName(cp01, Text(1), 0)
            End If
            'end 2019/8/6
         Else
            If Index = 0 Then
               DataErrorMessage 5, "來函性質"
               Cancel = True
            End If
            'Add By Cheng 2002/03/11
            '若下一程序為空白時, 本所期限及法定期限必須為空白
            If Index = 1 And Me.Text(Index).Text = "" Then
               Me.Text(2).Text = ""
               Me.Text(3).Text = ""
            End If
         End If
      '本所期限、法定期限：當下一程序為空白時，此二欄必須同時為空白，但下一程序不為空白時，則此二欄必須不可空白，檢查日期且本所期限必須≦法定期限。。
      Case 2 '本所期限
         'Add By Cheng 2002/03/11
         '若下一程序為空白時, 本所期限及法定期限必須為空白
         If Me.Text(1).Text = "" And Me.Text(Index).Text <> "" Then
            MsgBox "當沒輸入下一程序時, 本所期限必須為空白!!!"
            Me.Text(Index).Text = ""
            Cancel = True
            Exit Sub
         End If
         
         '若有輸入本所期限
         If Text(Index) <> "" Then
            If CheckIsTaiwanDate(Text(Index)) Then
               'Add By Cheng 2002/03/11
               If (Val(Me.Text(Index).Text + 19110000)) < ServerDate Then
                  DataErrorMessage 10, "本所期限"
                  Cancel = True
                  Exit Sub
               End If
               'add by sonia 2019/8/6 ACS若本所期限非工作天則直接調整至最近的工作天
               If cp01 = "ACS" Then Text(Index) = TransDate(PUB_GetWorkDay1(Text(Index), True), 1)
               If Text(3) <> "" Then
                  If Val(Text(3)) - Val(Text(Index)) < 0 Then
                     DataErrorMessage 13
                     Cancel = True
                  End If
               End If
            Else
               DataErrorMessage 2, "本所期限"
               Cancel = True
            End If
         'Add By Cheng 2002/03/11
         '若有輸入下一程序, 但未輸入本所期限
         ElseIf Me.Text(1).Text <> "" And Me.Text(Index).Text = "" Then
            DataErrorMessage 5, "本所期限"
            Cancel = True
         End If
      Case 3 '法定期限
         'Add By Cheng 2002/03/11
         '若下一程序為空白時, 本所期限及法定期限必須為空白
         If Me.Text(1).Text = "" And Me.Text(Index).Text <> "" Then
            MsgBox "當沒輸入下一程序時, 法定期限必須為空白!!!"
            Me.Text(Index).Text = ""
            Cancel = True
            Exit Sub
         End If
         
         '若有輸入法定期限
         If Text(Index) <> "" Then
            If CheckIsTaiwanDate(Text(Index)) Then
               If Text(2) <> "" Then
                  If Val(Text(Index)) - Val(Text(2)) < 0 Then
                     DataErrorMessage 12
                     Cancel = True
                  End If
               End If
            Else
               Cancel = True
            End If
         'Add By Cheng 2002/03/11
         '若有輸入下一程序, 但未輸入本所期限
         ElseIf Me.Text(1).Text <> "" And Me.Text(Index).Text = "" Then
            DataErrorMessage 5, "法定期限"
            Cancel = True
         End If
      Case 4
         If Text(Index) <> "" Then
            Text(Index) = UCase(Text(Index))
            'edit by nickc 2007/02/07 不用 dll 了
            'If objLawDll.GetGovName(Text(Index), strTemp1) Then Lbe(Index) = strTemp1 Else Cancel = True
            If ClsPDGetGovName(Text(Index), strTemp1) Then lbe(Index) = strTemp1 Else Cancel = True
         End If
      Case 5
         If Text(Index) <> "" Then Text(Index) = UCase(Text(Index))
         If CheckLengthIsOK(Text(5), 40) = False Then
            Cancel = True
            Text(5).SetFocus
            TextInverse Text(5)
         End If
      Case 7
         If Text(Index) <> "" Then
            If CheckIsTaiwanDate(Text(Index)) Then
               If Text(2).Text <> "" Then
                  If Val(Text(Index).Text) > Val(Text(2).Text) Then
                     MsgBox "承辦期限不可大於本所期限!", vbExclamation, "其他來函"
                     Cancel = True
                   End If
               End If
            Else
               Cancel = True
            End If
         End If
      Case 6, 8
         If Text(Index) <> "" Then
            Text(Index) = UCase(Text(Index))
            'edit by nickc 2007/02/07 不用 dll 了
            'If objPublicData.GetStaff(Text(Index), strTemp1) Then Lbe(Index) = strTemp1 Else Cancel = True
            If ClsPDGetStaff(Text(Index), strTemp1) Then lbe(Index) = strTemp1 Else Cancel = True
         End If
      Case 9
         If Text(Index) <> "" Then
            Text(Index) = UCase(Text(Index))
            If Text(Index) <> "Y" Then
               DataErrorMessage 1, "是否閉卷"
               Cancel = True
            End If
         End If
      Case 10
          If Text(Index) <> "" Then
             If CheckLengthIsOK(Text(Index), 2000) = False Then
                Cancel = True
             End If
          End If

   End Select
   If Cancel Then TextInverse Text(Index)
End Sub

Private Function ChgType(i As Integer, strText As String) As String
 Dim strTemp As String
   Select Case i
      Case 3, 4, 10
         ChgType = ChangeWStringToTString(strText)
      Case 1, 2
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetCaseProperty(CP01, StrText, strTemp, False) Then ChgType = strTemp
         If ClsPDGetCaseProperty(cp01, strText, strTemp, False) Then ChgType = strTemp
      Case 5
         'edit by nickc 2007/02/07 不用 dll 了
         'If objLawDll.GetGovName(StrText, strTemp) Then ChgType = strTemp
         If ClsPDGetGovName(strText, strTemp) Then ChgType = strTemp
      Case 9, 11
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetStaff(StrText, strTemp) Then ChgType = strTemp
         If ClsPDGetStaff(strText, strTemp) Then ChgType = strTemp
   End Select
End Function

Private Function ChkUData(strNation As String, strProperty As String) As Boolean
 Dim dblTempDays As Double, strDate As Variant
    'Added by Lydia 2023/03/17 ACS案一般來函：計算所限、法限
    If Text(0) <> "" Then
        strExc(0) = "select cpm07,cpm08,cpm09 from casepropertymap where cpm01='" & m_CP01 & "' and cpm02='" & Text(0) & "' "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
           strExc(2) = "": strExc(3) = ""
           If "" & RsTemp.Fields("cpm08") <> "" Then
              strExc(3) = CompDate(2, Val(RsTemp.Fields("cpm08")), TransDate(Replace(lbeAccept, "/", ""), 2))
           ElseIf "" & RsTemp.Fields("cpm09") <> "" Then
              strExc(3) = CompDate(1, Val(RsTemp.Fields("cpm09")), TransDate(Replace(lbeAccept, "/", ""), 2))
           End If
           If "" & RsTemp.Fields("cpm07") = "1" Then
              strExc(3) = CompDate(2, -1, strExc(3))
           End If
           If strExc(3) <> "" Then
               strExc(2) = PUB_GetOurDeadline(strExc(3))
               '所限：限制工作天
               strExc(2) = PUB_GetWorkDay1(strExc(2), True)
               If strExc(2) < strSrvDate(1) Then strExc(2) = strSrvDate(1)
               If strExc(3) < strSrvDate(1) Then strExc(3) = strSrvDate(1)
               Text(2) = TransDate(strExc(2), 1)
               Text(3) = TransDate(strExc(3), 1)
           End If
        End If
    End If
    'end 2023/03/17
    
   'edit by nickc 2007/02/07 不用 dll 了
   'If objLawDll.GetCaseFee(CP01, strNation, strProperty, dblTempDays) Then
''''edit by nickc 2007/10/12 改抓有時效性的
''''   If ClsLawGetCaseFee(CP01, strNation, strProperty, dblTempDays) Then
''''      strDate = DateAdd("d", dblTempDays, CDate(DateSerial(Val(Left(lbeAccept, 2)) + 1911, Val(Mid(lbeAccept, 4, 2)), Val(Right(lbeAccept, 2)))))
''''      Text(7) = ChangeWDateStringToTString(str(strDate))
''''   End If
    'Modified by Lydia 2018/04/30
    'Text(7) = Pub_GetHandleDay(cp01, strNation, strProperty, lbeAccept, Text(2))
    Text(7) = TransDate(Pub_GetHandleDay(cp01, strNation, strProperty, TransDate(Replace(lbeAccept, "/", ""), 2), TransDate(Text(2), 2)), 1)
End Function

Private Sub Getrs()
 Dim LcTmp As String
   LcTmp = cp01 + cp02 + cp03 + cp04
   strExc(1) = "select decode(np02||np07,cpm01||CPM02,CPM03,CPM04)," + _
      "decode(np08,null,'',SUBSTR(np08,1,4)-1911||'/'||SUBSTR(np08,5,2)||'/'||SUBSTR(np08,7,2))," + _
      "decode(np09,null,'',SUBSTR(np09,1,4)-1911||'/'||SUBSTR(np09,5,2)||'/'||SUBSTR(np09,7,2))," + _
      "np13,np14,np01,np07,np22,'' from nextprogress,CASEPROPERTYMAP where " + _
      "" + ChgNextProgress(LcTmp) + " and np02=cpm01(+) and np07=cpm02(+) and (np06='N' or np06 is null)"
   intI = 1
   'edit by nickc 2007/02/07 不用 dll 了
   'Set MSHFlexGrid1.Recordset = objLawDll.ReadRstMsg(intI, strExc(1))
   Set MSHFlexGrid1.Recordset = ClsLawReadRstMsg(intI, strExc(1))
   GridHead
End Sub

Private Sub GridHead()
 Dim i As Integer
   With MSHFlexGrid1
      blnOKtoShow = False
      .row = 0
      .col = 0: .ColWidth(0) = 1200: .Text = "下一程序"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1000: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1500: .Text = "機關文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1500: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
      For i = 5 To 8
         .col = i: .ColWidth(i) = 0
      Next
      intLastRow = 0
      blnOKtoShow = True
'判斷是否有資料
   End With
End Sub

Private Function SaveNextProgress() As Boolean
 Dim i As Integer
 Dim np01 As String, NP07 As String, np22 As String
   SaveNextProgress = True
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         .row = i
         .col = 8
         If .Text = "Y" Then
            .col = 5
            np01 = .Text
            .col = 6
            NP07 = .Text
            .col = 7
            np22 = .Text
            strExc(1) = "update nextprogress set np06='Y' where np01='" & np01 & _
               "' and np07='" & NP07 & "' and np22='" & np22 & "'"
            'edit by nickc 2007/02/07 不用 dll 了
            'SaveNextProgress = objLawDll.ExecSQL(1, strExc)
            SaveNextProgress = ClsLawExecSQL(1, strExc)
         End If
      Next
   End With
End Function

Private Sub MSHFlexGrid1_Click()
 Dim i As Integer
   With MSHFlexGrid1
      .col = 0
      If MSHFlexGrid1.CellBackColor = &HFFC0C0 Then
         For i = 0 To 8
            .col = i
            .CellBackColor = MSHFlexGrid1.BackColor
         Next
         .col = 8: .Text = ""
      Else
         For i = 0 To 8
            .col = i
            .CellBackColor = &HFFC0C0
         Next
         .col = 8: .Text = "Y"
      End If
   End With
End Sub

Private Sub cmdSure_Click()
Dim strDay1 As String, strDay2 As String
Dim strDate As String
Dim m_StrTo As String, m_contents As String
   
'2011/5/17 cancel by sonia
'   If Not AllTextBeforeSaveCheck Then Exit Sub
'      strDate = GetMailRecField(m_CP01, m_CP02, m_CP03, m_CP04, DBDATE(lbeAccept.Caption), "MR17")
'   If TAIWANDATE(Text(3).Text) <> TAIWANDATE(strDate) Then
'         If MsgBox("輸入的法定期限與來函記錄中的法定期限日期不同", vbYesNo, "資料檢核") = vbNo Then
'            Text(3).SetFocus
'            Exit Sub
'         End If
'   End If
'   strDate = ""
'   strDate = GetMailRecField(m_CP01, m_CP02, m_CP03, m_CP04, DBDATE(lbeAccept.Caption), "MR16")
'   If TAIWANDATE(Text(2).Text) <> TAIWANDATE(strDate) Then
'         If MsgBox("輸入的本所期限與來函記錄中的本所期限日期不同", vbYesNo, "資料檢核") = vbNo Then
'            Text(2).SetFocus
'            Exit Sub
'         End If
'   End If
'2011/5/17 end
   
   'Add By Cheng 2002/05/24
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
   
   Screen.MousePointer = 11
   If SaveData = False Then Exit Sub
   
   If Not blnIsSave Then
      DataErrorMessage (3)
   Else
      'Add By Sindy 2011/10/20 發E-Mail給智權人員
      If cp01 = "FCL" Or cp01 = "LIN" Then
         m_StrTo = PUB_GetFCLSalesNo(cp01, cp02, cp03, cp04)
      Else
         m_StrTo = PUB_GetAKindSalesNo(cp01, cp02, cp03, cp04)
      End If
      'Modify By Sindy 2025/10/15 原只發EMAIL給案件最新智權人員，請修改為若有案源則發給案源介紹人(可能多個)，無案源才發給案件最新智權人員。
      strExc(0) = "select * from LawOfficeSource where LOS06='" & lbeNum & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_StrTo = Replace(RsTemp.Fields("LOS04"), ",", ";")
      End If
      '2025/10/15 END
      If m_StrTo > "" Then
         'Modified by Lydia 2015/10/05 '法務人員'改為'協辦人員'
         'Modified by Lydia 2025/11/18 Text(4) & lbe(4) 改為 Trim(cboGov.Text)
         m_contents = "本所案號：" & cp01 & "-" & cp02 & "-" & cp03 & "-" & cp04 & vbCrLf & _
                      "智權人員：" & GetPrjSalesNM(m_StrTo) & vbCrLf & _
                      "當 事 人：" & lbeCusName & vbCrLf & _
                      "來函性質：" & Text(0) & lbe(0) & vbCrLf & _
                      "下一程序：" & Text(1) & lbe(1) & vbCrLf & _
                      "本所期限：" & ChangeTStringToTDateString(Text(2)) & vbCrLf & _
                      "法定期限：" & ChangeTStringToTDateString(Text(3)) & vbCrLf & _
                      "機關代號：" & Trim(cboGov.Text) & vbCrLf & _
                      "機關文號：" & Text(5) & vbCrLf & _
                      "承 辦 人：" & Text(6) & lbe(6) & vbCrLf & _
                      "承辦期限：" & ChangeTStringToTDateString(Text(7)) & vbCrLf & _
                      "協辦人員：" & Text(8) & lbe(8) & vbCrLf & _
                      "進度備註：" & Text(10) & vbCrLf
                     PUB_SendMail strUserNum, m_StrTo, "", cp01 & "-" & cp02 & "-" & cp03 & "-" & cp04 & "其他來函", m_contents
      End If
      '2011/10/20 End
   End If
   
   Screen.MousePointer = 0
   Unload Me
   Unload frm071014
   frm071014.Show
End Sub

'Mark by Lydia 2025/11/18 不使用的程式
'Private Function AllTextBeforeSaveCheck() As Boolean
'   Dim strTemp  As String, yn As Integer
'   Dim strDay1 As String, strDay2 As String
'
'   AllTextBeforeSaveCheck = True
'   If Text(0) <> "" Then
'      'edit by nickc 2007/02/07 不用 dll 了
'      'If Not objPublicData.GetCaseProperty(CP01, Text(0), strTemp, False) Then
'      If Not ClsPDGetCaseProperty(cp01, Text(0), strTemp, False) Then
'         AllTextBeforeSaveCheck = False
'         Text(0).SetFocus
'         TextInverse Text(0)
'         Exit Function
'       End If
'    Else
'       MsgBox "來函性質不可空白", vbCritical
'       Text(0).SetFocus
'       AllTextBeforeSaveCheck = False
'       Exit Function
'    End If
'
'   '檢查本所期限
'   '若沒輸入下一程序時, 本所期限必須為空白
'   If Me.Text(1).Text = "" And Me.Text(2).Text <> "" Then
'      MsgBox "當沒輸入下一程序時, 本所期限必須為空白!!!"
'      AllTextBeforeSaveCheck = False
'      Text(2).SetFocus
'      TextInverse Text(2)
'      Exit Function
'   End If
'   '若有輸入本所期限
'   If Text(2) <> "" Then
'      If CheckIsTaiwanDate(Text(2)) = False Then
'         AllTextBeforeSaveCheck = False
'         Text(2).SetFocus
'         TextInverse Text(2)
'         Exit Function
'      End If
'      If (Val(Me.Text(2).Text) + 19110000) < ServerDate Then
'         DataErrorMessage 10, "本所期限"
'         AllTextBeforeSaveCheck = False
'         Text(2).SetFocus
'         TextInverse Text(2)
'         Exit Function
'      End If
'   'Add By Cheng 2002/03/11
'   '若有輸入下一程序, 但未輸入本所期限
'   ElseIf Me.Text(1).Text <> "" And Me.Text(2).Text = "" Then
'      DataErrorMessage 5, "本所期限"
'      AllTextBeforeSaveCheck = False
'      Text(2).SetFocus
'      TextInverse Text(2)
'      Exit Function
'   End If
'
'   '檢查法定期限
'   '若沒輸入下一程序時, 法定期限必須為空白
'   If Me.Text(1).Text = "" And Me.Text(3).Text <> "" Then
'      MsgBox "當沒輸入下一程序時, 法定期限必須為空白!!!"
'      AllTextBeforeSaveCheck = False
'      Text(3).SetFocus
'      TextInverse Text(3)
'      Exit Function
'   End If
'   '若有輸入法定期限
'   If Text(3) <> "" Then
'      If CheckIsTaiwanDate(Text(3)) = False Then
'         AllTextBeforeSaveCheck = False
'         Text(3).SetFocus
'         TextInverse Text(3)
'         Exit Function
'      End If
'   End If
'
'    If Text(1) <> "" Then
'       If Text(2) = "" Or Text(3) = "" Then
'          MsgBox "本所期限與法定期限不可為空", vbCritical
'          AllTextBeforeSaveCheck = False
'          If Text(2).Text = "" Then
'             Text(2).SetFocus
'             TextInverse Text(2)
'             Exit Function
'          End If
'          If Text(3).Text = "" Then
'             Text(3).SetFocus
'             TextInverse Text(3)
'             Exit Function
'          End If
'       End If
'    Else
'        If Text(2) <> "" Or Text(3) <> "" Then
'
'           MsgBox "本所期限與法定期限須為空", vbCritical
'           AllTextBeforeSaveCheck = False
'           Text(2).SetFocus
'           Exit Function
'        End If
'    End If
'    strExc(2) = cp01 + cp02 + cp03 + cp04
''   If Not objLawDll.ChkMRec(ChangeTStringToWString(Replace(lbeAccept, "/", "")), strExc(2), strDay1, strDay2) Then
''      MsgBox "本所案號'" + lbeCaseNum + "'與收受日'" + lbeAccept + "'不存在於來函記錄檔中", vbCritical
''      Exit Function
''   ElseIf strDay1 <> "" Then
''      If Text(2) <> ChangeWStringToTString(strDay1) Or Text(3) <> ChangeWStringToTString(strDay2) Then
''       If MsgBox("來函記錄檔之本所期限 '" + ChangeWStringToTString(strDay1) + "' 與法定期限 '" + ChangeWStringToTString(strDay2) + "' 與畫面輸入不同，是否儲存 ? ", vbCritical + vbYesNo) = vbNo Then
''         Text(2).SetFocus
''         Exit Function
''       End If
''      End If
''   End If
'         If Text(4) <> "" Then
'            strTemp = ""
'            Text(4) = UCase(Text(4))
'            'edit by nickc 2007/02/07 不用 dll 了
'            'If objLawDll.GetGovName(Text(4), strTemp) Then
'            If ClsPDGetGovName(Text(4), strTemp) Then
'               lbe(4) = strTemp
'            Else
'               AllTextBeforeSaveCheck = False
'               Text(4).SetFocus
'               Exit Function
'            End If
'         End If
'         If Text(7).Text <> "" Then
'            If CheckIsTaiwanDate(Text(7)) = False Then
'               AllTextBeforeSaveCheck = False
'               Text(7).SetFocus
'               TextInverse Text(7)
'               Exit Function
'          End If
'        End If
'   If Text(7).Text <> "" And Text(2).Text <> "" Then
'
'     If Val(Text(7).Text) > Val(Text(2).Text) Then
'        MsgBox "承辦期限不可大於本所期限!", vbExclamation, "法院文書"
'        Text(7).SetFocus
'        AllTextBeforeSaveCheck = False
'        Exit Function
'     End If
'   End If
'   If Text(9) = "Y" Then
'      If MsgBox("確定要閉卷嗎？", vbYesNo + vbCritical + vbDefaultButton1) = vbNo Then
'         Text(9) = ""
'         AllTextBeforeSaveCheck = False
'         Exit Function
'      End If
'   End If
'
'   If Text(6) <> "" Then
'      Text(6) = UCase(Text(6))
'      strTemp = ""
'      'edit by nickc 2007/02/07 不用 dll 了
'      'If objPublicData.GetStaff(Text(6), strTemp) Then
'      If ClsPDGetStaff(Text(6), strTemp) Then
'         lbe(6) = strTemp
'      Else
'         Text(6).SetFocus
'         TextInverse Text(6)
'         AllTextBeforeSaveCheck = False
'         Exit Function
'      End If
'  End If
'
'   If Text(8) <> "" Then
'      Text(8) = UCase(Text(8))
'      strTemp = ""
'      'edit by nickc 2007/02/07 不用 dll 了
'      'If objPublicData.GetStaff(Text(8), strTemp) Then
'      If ClsPDGetStaff(Text(8), strTemp) Then
'         lbe(8) = strTemp
'      Else
'         Text(8).SetFocus
'         TextInverse Text(8)
'         AllTextBeforeSaveCheck = False
'         Exit Function
'      End If
'  End If
'
'   If Text(10) <> "" Then
'      If CheckLengthIsOK(Text(10), 2000) = False Then
'         Text(10).SetFocus
'         TextInverse Text(10)
'         AllTextBeforeSaveCheck = False
'         Exit Function
'      End If
'   End If
'End Function

Private Function SaveData() As Boolean
Dim strNewNum As String, strNum As String, strSaleArea As String, strTemp As String
Dim strID As String, strdt As String, strTM As String, yn As Boolean
Dim j As Long, iStep As Integer
Dim strCP27 As String 'Modify By Sindy 2020/10/12
Dim strCP13 As String  'Added by Lydia 2020/11/06

'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler

SaveData = True
cnnConnection.BeginTrans
   
   'edit by nickc 2007/02/07 不用 dll 了
   'If objPublicData.GetAutoNumber("C", strNewNum, 1, 1) Then
   'Modified by Lydia 2020/11/06 改用模組
   'If ClsPDGetAutoNumber("C", strNewNum, 1, 1) Then
   '   'Modify By Sindy 2010/8/18 比對自動編號年度
   '   'strNum = "C" + CStr(Year(Date) - 1911) + strNewNum
   '   strNum = "C" + CompAutoNumberYear(CStr(Val(Mid(strSrvDate(1), 1, 4)) - 1911)) + strNewNum
   'End If
   strNum = AutoNo("C", 6)
   'end 2020/11/06
   
   yn = GetCreateUpdateDate(strID, strdt, strTM)
   '2009/9/9 MODIFY BY SONIA
   'strSaleArea = GetST15(Strsale)
   'Modified by Lydia 2020/11/06
   'strSaleArea = GetSalesArea(IIf(cp01 = "FCL", PUB_GetFCLSalesNo(cp01, cp02, cp03, cp04), IIf(cp01 = "LIN", PUB_GetFCLSalesNo(cp01, cp02, cp03, cp04), PUB_GetAKindSalesNo(cp01, cp02, cp03, cp04))))
   Select Case cp01
       Case "FCL", "LIN"
           strCP13 = PUB_GetFCLSalesNo(cp01, cp02, cp03, cp04)
       Case Else
           strCP13 = PUB_GetAKindSalesNo(cp01, cp02, cp03, cp04)
   End Select
   strSaleArea = GetSalesArea(strCP13)
   'end 2020/11/06
   
   ' 91.04.04 modify by louis (修改單引號)
   'Modify By Cheng 2003/04/07
   '智權人員存最近收文A類接洽記錄單的智權人員
   'Modify By Sindy 2011/10/20 存檔時同時將此C類來文上發文日
   'Modify By Sindy 2020/10/12 ACS案件有輸入承辦期限時不上發文日，沒輸才上發文日。
   'modify by sonia 2023/10/12 若已自動內部收文下一程序時也要上發文日ACS-000160
   If cp01 = "ACS" And Val(Text(7)) > 0 And textCF30 <> "Y" Then
      strCP27 = ""
   Else
   '2020/10/12 END
      strCP27 = strSrvDate(1)
   End If
   '2020/10/12 END
   'Modified by Lydia 2020/11/06 改語法
   'strExc(iStep) = "insert into caseprogress(cp09,cp01,cp02,cp03,cp04,cp05,cp06,cp07," & _
      "cp12,cp13,cp20,cp32,cp43,cp26,cp10,cp08,cp65,cp66,cp67,CP14,CP48,CP29,CP64,cp27) values (" + _
      CNULL(strNum) + "," + CNULL(cp01) + "," + CNULL(cp02) + "," + _
      CNULL(cp03) + "," + CNULL(cp04) + "," + CNULL(ChangeTStringToWString(Replace(lbeAccept, "/", ""))) + "," + _
      CNULL(IIf(Text(2) = "", "", ChangeTStringToWString(Text(2)))) + "," + CNULL(IIf(Text(3) = "", "", ChangeTStringToWString(Text(3)))) + "," + CNULL(strSaleArea) + "," + _
      IIf(cp01 = "FCL", CNULL(PUB_GetFCLSalesNo(cp01, cp02, cp03, cp04)), IIf(cp01 = "LIN", CNULL(PUB_GetFCLSalesNo(cp01, cp02, cp03, cp04)), CNULL(PUB_GetAKindSalesNo(cp01, cp02, cp03, cp04)))) + ",'N','N'," + CNULL(lbeNum) + ",'N'," + CNULL(Text(0)) + _
      "," + CNULL(Text(5)) + "," + CNULL(strID) + "," + CNULL(strdt) + "," + _
      CNULL(strTM) + "," + CNULL(Text(6)) + "," + CNULL(IIf(Text(7) = "", "", ChangeTStringToWString(Text(7)))) + "," + CNULL(Text(8)) + "," + CNULL(ChgSQL(Text(10))) + "," + CNULL(strCP27, True) + ")"
   'Modified by Lydia 2025/11/18 +CP71
   strExc(iStep) = "insert into caseprogress(cp09,cp01,cp02,cp03,cp04,cp05,cp06,cp07," & _
      "cp12,cp13,cp20,cp32,cp43,cp26,cp10,cp08,CP14,CP48,CP29,CP64,cp27,CP71) values (" + _
      CNULL(strNum) + "," + CNULL(cp01) + "," + CNULL(cp02) + "," + _
      CNULL(cp03) + "," + CNULL(cp04) + "," + CNULL(ChangeTStringToWString(Replace(lbeAccept, "/", ""))) + "," + _
      CNULL(IIf(Text(2) = "", "", ChangeTStringToWString(Text(2)))) + "," + CNULL(IIf(Text(3) = "", "", ChangeTStringToWString(Text(3)))) + "," + CNULL(strSaleArea) + "," + _
      CNULL(strCP13) + ",'N','N'," + CNULL(lbeNum) + ",'N'," + CNULL(Text(0)) + _
      "," + CNULL(Text(5)) + "," + CNULL(Text(6)) + "," + CNULL(IIf(Text(7) = "", "", ChangeTStringToWString(Text(7)))) + "," + CNULL(Text(8)) + "," + CNULL(ChgSQL(Text(10))) + "," + CNULL(strCP27, True) + "," + CNULL(Trim(Left(cboGov.Text, 3))) + ")"
   'Add By Cheng 2002/11/07
   cnnConnection.Execute strExc(iStep)
   
   'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
   Pub_UpdateFromMaxCP27 cp01, cp02, cp03, cp04
    
   iStep = 2
   If Text(1).Text <> "" Then   '下一程序
      'Added by Lydia 2020/11/06 自動內部收文下一程序
      strExc(9) = ""
      If textCF30.Visible = True And textCF30 = "Y" Then
         strExc(9) = AutoNo("B", 6)
         strExc(7) = "": strExc(8) = ""
         strExc(7) = PUB_GetAKindSalesNo(cp01, cp02, cp03, cp04)
         strExc(8) = strExc(7)
         '收文日為系統日，案件性質為畫面上的下一程序，本所期限及法定期限為畫面上的期限，CP08為畫面上的機關文號、智權人員依PUB_GetAKindSalesNo帶出，承辦人預設為操作人員，CP11='90'
         'Modified by Lydia 2023/03/17 +cp48承辦期限 Text(7)
         'Modify By Sindy 2023/10/20 CP14=CNULL(strUserNum) 要改存畫面上的承辦人CNULL(Text(6)) EX:ACS-000176(1006=驗證通過)
         strSql = "insert into caseprogress(cp09,cp01,cp02,cp03,cp04,cp05,cp06,cp07," & _
                      "cp11,cp12,cp13,cp20,cp32,cp43,cp26,cp10,cp08,CP14,CP64,CP48) values (" + _
                      CNULL(strExc(9)) + "," + CNULL(cp01) + "," + CNULL(cp02) + "," + CNULL(cp03) + "," + CNULL(cp04) + "," + CNULL(strSrvDate(1)) + "," + _
                      CNULL(DBDATE(Text(2))) + "," + CNULL(DBDATE(Text(3))) + ",'90'," + CNULL(strSaleArea) + "," + CNULL(strCP13) + ",'N','N'," + CNULL(strNum) + ",'N'," + CNULL(Text(1)) + _
                     "," + CNULL(Text(5)) + "," + CNULL(Text(6)) + "," + CNULL(ChgSQL(Text(10))) + "," + CNULL(DBDATE(Text(7)), True) + ")"
         cnnConnection.Execute strSql
      End If
      'end 2020/11/06
      
      'edit by nickc 2007/02/07 不用 dll 了
      'j = objLawDll.GetMax
      j = ClsLawGetMax
      ' 91.04.04 modify by louis (修改單引號)
        'Modify By Cheng 2003/04/07
        '智權人員存最近收文A類接洽記錄單的智權人員
      'Modified by Lydia 2020/11/06 改變數
      'strExc(iStep) = "insert into nextprogress(np01,np02,np03,np04,np05,np07,np08,np09," & _
         "np10,np13,np15,np16,np17,np18,np22) values (" + CNULL(strNum) + "," + _
         CNULL(cp01) + "," + CNULL(cp02) + "," + CNULL(cp03) + "," + _
         CNULL(cp04) + "," + CNULL(Text(1)) + "," + CNULL(IIf(Text(2) = "", "", ChangeTStringToWString(Text(2)))) + "," + _
         CNULL(IIf(Text(3) = "", "", ChangeTStringToWString(Text(3)))) + "," + IIf(cp01 = "FCL", CNULL(PUB_GetFCLSalesNo(cp01, cp02, cp03, cp04)), IIf(cp01 = "LIN", CNULL(PUB_GetFCLSalesNo(cp01, cp02, cp03, cp04)), CNULL(PUB_GetAKindSalesNo(cp01, cp02, cp03, cp04)))) + "," + CNULL(Text(5)) + "," + _
         CNULL(ChgSQL(Text(10))) + "," + CNULL(strID) + "," + CNULL(strdt) + "," + CNULL(strTM) + "," & j & ")"
      strExc(iStep) = "insert into nextprogress(np01,np02,np03,np04,np05,np06,np07,np08,np09,np10,np13,np15,np22,np24) values (" & CNULL(strNum) & "," & _
                        CNULL(cp01) & "," & CNULL(cp02) & "," & CNULL(cp03) & "," & CNULL(cp04) & "," & CNULL(IIf(strExc(9) <> "", "Y", "")) & "," & CNULL(Text(1)) & "," & CNULL(DBDATE(Text(2))) & "," & _
                        CNULL(DBDATE(Text(3))) & "," & CNULL(strCP13) & "," & CNULL(Text(5)) & "," & CNULL(ChgSQL(Text(10))) & "," & j & "," & CNULL(strExc(9)) & ")"
        'Add By Cheng 2002/11/07
        cnnConnection.Execute strExc(iStep)
      iStep = iStep + 1
   End If
   
   'Modify By Sindy 2012/3/5 增加判斷及更新閉卷日期及閉卷原因
'   strExc(iStep) = "update lawcase set lc08=" + CNULL(Text(9)) + " where lc01=" + _
'      CNULL(CP01) + " and lc02=" + CNULL(cp02) + " and " + _
'      " lc03=" + CNULL(cp03) + " and lc04=" + CNULL(cp04) + ""
   If cp01 <> "LA" Then 'lawcase
      If Trim(Text(9)) = "" Then
         strExc(iStep) = "update lawcase set lc08=null,lc09=null,lc10=null where lc01=" + _
           CNULL(cp01) + " and lc02=" + CNULL(cp02) + " and " + _
           " lc03=" + CNULL(cp03) + " and lc04=" + CNULL(cp04) + ""
      Else
         '原基本檔非閉卷時才更新
         If Trim(Text(9)) = "Y" And m_bolClose <> "Y" Then
            strExc(iStep) = "update lawcase set lc08=" + CNULL(Text(9)) + ",lc09=" & strSrvDate(1) & ",lc10='99' where lc01=" + _
              CNULL(cp01) + " and lc02=" + CNULL(cp02) + " and " + _
              " lc03=" + CNULL(cp03) + " and lc04=" + CNULL(cp04) + ""
         End If
      End If
   Else 'hirecase
      If Trim(Text(9)) = "" Then
         strExc(iStep) = "update hirecase set hc09=null,hc10=null,hc11=null where hc01=" + _
           CNULL(cp01) + " and hc02=" + CNULL(cp02) + " and " + _
           " hc03=" + CNULL(cp03) + " and hc04=" + CNULL(cp04) + ""
      Else
         '原基本檔非閉卷時才更新
         If Trim(Text(9)) = "Y" And m_bolClose <> "Y" Then
            strExc(iStep) = "update hirecase set hc09=" + CNULL(Text(9)) + ",hc10=" & strSrvDate(1) & ",hc11='99' where hc01=" + _
              CNULL(cp01) + " and hc02=" + CNULL(cp02) + " and " + _
              " hc03=" + CNULL(cp03) + " and hc04=" + CNULL(cp04) + ""
         End If
      End If
   End If
   '2012/3/5 End
   'Add By Cheng 2002/11/07
   cnnConnection.Execute strExc(iStep)
   iStep = iStep + 1
   
'   SaveData = objLawDll.ExecSQL(iStep - 1, strExc)
   blnIsSave = False
   If SaveData = True Then
        If SaveNextProgress = True Then
            blnIsSave = True
        'Add By Cheng 2002/11/07
        Else
            GoTo ErrorHandler
        End If
   End If
'Add By Cheng 2002/11/07
cnnConnection.CommitTrans
Exit Function

ErrorHandler:
    cnnConnection.RollbackTrans
    SaveData = False
    'Added by Lydia 2020/11/06
    Screen.MousePointer = 0
    MsgBox "存檔失敗：" & vbCrLf & Err.Description, vbCritical, "存檔作業失敗"
    'end 2020/11/06
End Function

'Add By Cheng 2002/05/24
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
For Each objTxt In Text
   If objTxt.Enabled = True Then
      Cancel = False
      Text_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

'Added by Lydia 2020/11/06
If textCF30.Visible = True And textCF30 = "Y" Then
    If Text(1) = "" Then
        MsgBox "請輸入下一程序！", vbExclamation, "檢核資料"
        Text(1).SetFocus
        Call Text_GotFocus(1)
        Exit Function
    End If
    'Remove by Lydia 2023/03/17 需要存入承辦期限
    'If Trim(Text(7)) <> "" Then
    '    MsgBox "不可輸入承辦期限！", vbExclamation, "檢核資料"
    '    Text(7).SetFocus
    '    Call Text_GotFocus(7)
    '    Exit Function
    'End If
    'end 2023/03/17
End If
'end 2020/11/06

'Added by Lydia 2025/11/18
If Trim(cboGov.Text) <> "" Then
   Call cboGov_Validate(Cancel)
   If Cancel = True Then
      Exit Function
   End If
End If
'end 2025/11/18

'Added by Lydia 2021/09/14 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
     Exit Function
End If

TxtValidate = True
End Function

'Added by Lydia 2020/11/06
Private Sub textCF30_GotFocus()
    TextInverse textCF30
End Sub

Private Sub textCF30_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCF30_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse

   Cancel = False
   If IsEmptyText(textCF30) = False Then
      Select Case textCF30
         Case "Y", "", " ":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "是否自動內部收文下一程序請輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCF30_GotFocus
      End Select
   End If
End Sub

'Added by Lydia 2025/11/18
Private Sub cboGov_GotFocus()
   TextInverse cboGov
End Sub

'Added by Lydia 2025/11/18
Private Sub cboGov_Validate(Cancel As Boolean)

   If Trim(cboGov.Text) <> "" And cboGov.Tag <> cboGov.Text Then
      If PUB_ChkGovIsExist(IIf(Val(Trim(Left(cboGov.Text, 3))) > 0, Trim(Left(cboGov.Text, 3)), Trim(cboGov.Text)), strExc(3), strExc(4)) = True Then
         cboGov.Text = strExc(3) & " " & strExc(4)
      Else
         Cancel = True
         cboGov.SetFocus
         cboGov_GotFocus
         Exit Sub
      End If
   End If
   cboGov.Tag = cboGov.Text
End Sub

'Added by Lydia 2025/11/18
Private Sub cboGov_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Added by Lydia 2025/11/18
Private Sub cboGov_DropButtonClick()
   If cboGov.Text <> "" Then
      If Val(Trim(Left(cboGov.Text, 3))) > 0 Then
      Else  '依輸入文字模糊比對
         Call PUB_SetGovCmb(cboGov, m_GovListNew, , Trim(cboGov.Text))
         If m_GovListNew = "" Then
            Call PUB_SetGovCmb(cboGov, m_GovListNew)
         End If
      End If
   Else
      If m_GovListNew <> m_GovListDef Then
         Call PUB_SetGovCmb(cboGov, m_GovListNew)
      End If
   End If
End Sub

