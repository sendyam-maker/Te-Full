VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040103_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書／書表"
   ClientHeight    =   6168
   ClientLeft      =   -4296
   ClientTop       =   1176
   ClientWidth     =   9348
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6168
   ScaleWidth      =   9348
   Begin VB.TextBox txtKind 
      Height          =   300
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   5
      Top             =   5784
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   3060
      TabIndex        =   3
      Top             =   600
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1020
      MaxLength       =   3
      TabIndex        =   12
      Text            =   "P"
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   0
      Top             =   660
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   2340
      MaxLength       =   1
      TabIndex        =   1
      Top             =   660
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   2580
      MaxLength       =   2
      TabIndex        =   2
      Top             =   660
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   6120
      MaxLength       =   7
      TabIndex        =   8
      Top             =   660
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   405
      Index           =   0
      Left            =   7530
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   8364
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   4
      Top             =   5370
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3735
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   9075
      _ExtentX        =   16002
      _ExtentY        =   6583
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
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
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label11 
      Height          =   252
      Left            =   7008
      TabIndex        =   24
      Top             =   1248
      Width           =   1692
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "申請國家:"
      Height          =   180
      Left            =   6168
      TabIndex        =   23
      Top             =   1260
      Width           =   768
   End
   Begin VB.Label lblKind 
      AutoSize        =   -1  'True
      Caption         =   "(1.解聘書  2.大陸案委託書 3.委任書  4.讓與契約書  5.簽章切結書)"
      Height          =   180
      Index           =   1
      Left            =   1680
      TabIndex        =   22
      Top             =   5844
      Width           =   5076
   End
   Begin VB.Label lblKind 
      Caption         =   "書表種類:"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   5808
      Width           =   972
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1020
      TabIndex        =   9
      Top             =   1200
      Width           =   4935
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "8705;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   180
      TabIndex        =   20
      Top             =   660
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期:"
      Height          =   180
      Left            =   5160
      TabIndex        =   19
      Top             =   660
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號"
      Height          =   180
      Left            =   180
      TabIndex        =   18
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   180
      Left            =   1020
      TabIndex        =   17
      Top             =   960
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   5160
      TabIndex        =   16
      Top             =   960
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Label6"
      Height          =   180
      Left            =   6120
      TabIndex        =   15
      Top             =   960
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   180
      TabIndex        =   14
      Top             =   1260
      Width           =   765
   End
   Begin VB.Label Label9 
      Caption         =   "特殊申請書:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "(1.詢進度 2.其他申請書 3.電子送件 4.更正已繳年度 5.延期)"
      Height          =   180
      Left            =   1680
      TabIndex        =   11
      Top             =   5400
      Width           =   4575
   End
End
Attribute VB_Name = "frm040103_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/10 改成Form2.0 (Combo1,MSHFlexGrid1)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
Option Explicit

'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String
Dim intWhere As Integer
Dim intLastRow As Integer
'Public iFrom As Integer '0=內專,1=承辦人 Add by Morgan 2011/9/22 'Mark by Lydia 2023/09/22 承辦人改用frm04010301_1
Public stCP09 As String 'Added by Morgan 2020/4/8
Dim stCP10 As String 'Added by Lydia 2023/09/22

'Added by Lydia 2023/09/22
Private Function TxtValidate() As Boolean
Dim intQ As Integer, bolChk As Boolean
   
   If pa(1) & pa(2) & pa(3) & pa(4) <> Text1 & Text2 & Text3 & Text4 Or Trim(pa(5) & pa(6) & pa(7)) = "" Then
      MsgBox "請先尋找本所案號的資料！", vbExclamation
      Exit Function
   End If
   If Me.Text6.Text <> "" And Me.txtKind <> "" Then
      MsgBox "不可同時產生申請書和書表 !", vbInformation
      Exit Function
   End If
      
   Me.Tag = ""
   stCP10 = ""
   For intQ = 1 To MSHFlexGrid1.Rows - 1
      If MSHFlexGrid1.TextMatrix(intQ, 0) = "v" Then
         bolChk = True
         Me.Tag = MSHFlexGrid1.TextMatrix(intQ, 2)
         stCP10 = MSHFlexGrid1.TextMatrix(intQ, 7)
         Exit For
      End If
   Next
   
   If pa(9) = 台灣國家代號 Then
      '若特殊申請書不為"2"時, 須檢查是否有勾選資料
      If Me.Text6.Text <> "2" And Me.Text6.Text <> "" Then
         If bolChk = False Then
            MsgBox "請選擇資料 !", vbInformation
            Exit Function
         End If
      End If
      If InStr("1,2,", Me.txtKind) > 0 And Me.txtKind <> "" Then
         MsgBox "解聘書和大陸案委託書僅限於非臺灣案可以使用！", vbInformation
         Exit Function
      End If
   ElseIf pa(1) = "P" Then
      Me.Text6.Text = ""
      If Trim(Me.txtKind) = "" Then
         MsgBox "非臺灣案請輸入書表種類！", vbExclamation
         txtKind.SetFocus
         txtkind_GotFocus
         Exit Function
      End If
   End If
   If pa(1) <> "P" Then Me.txtKind = ""
   
   TxtValidate = True
End Function

Private Sub cmdok_Click(Index As Integer)
 'Dim i As Integer, bolChk As Boolean 'Mark by Lydia 2023/09/22
   Select Case Index
      Case 0 '確定
         'Memo by Lydia 2023/09/22 cmdOK_Click(0)確定按鈕模組的pa(10)改為stCP10
         'Modified by Lydia 2023/09/22 改成模組
         'Me.Tag = ""
         'stCP10 = ""
         'For i = 1 To MSHFlexGrid1.Rows - 1
         '   If MSHFlexGrid1.TextMatrix(i, 0) = "v" Then
         '      bolChk = True
         '      Me.Tag = MSHFlexGrid1.TextMatrix(i, 2)
         '      stCP10 = MSHFlexGrid1.TextMatrix(i, 7)
         '      Exit For
         '   End If
         'Next
         '
         ''Add By Cheng 2002/07/12
         ''若特殊申請書不為"2"時, 須檢查是否有勾選資料
         'If Me.Text6.Text <> "2" Then
         '   If bolChk = False Then
         '      MsgBox "請選擇資料 !", vbInformation
         '      Exit Sub
         '   End If
         'End If
         If TxtValidate = False Then Exit Sub
         If Me.txtKind <> "" Then '書表種類
            'Modified by Lydia 2024/03/21 +收文號Me.Tag
            Call frm040110.SetParent(Me, Me.txtKind, Me.Tag, pa)
            frm040110.Show
         Else
         'end 2023/09/22
         
            If Me.Text6 = "1" Then
               frm04010309_1.Show
            'Add By Cheng 2002/07/12
            ElseIf Me.Text6.Text = "2" Then
               Set frm04010310_1.oParentForm = Me 'Add by Morgan 2011/9/22
               frm04010310_1.m_strCP10 = stCP10
               frm04010310_1.Show
   'Remove by Morgan 2005/8/1 不再使用,已併入領證繳年費
   '         'Add By Cheng 2002/07/12
            ElseIf Me.Text6.Text = "3" Then 'Add By Sindy 2018/6/29 + 電子送件申請書
   '            frm04010305_1.Show
               Select Case stCP10
                  'Modified by Morgan 2023/5/18 +244補中文說明書,232補優先權證明
                  Case 補文件, 實體審查, "244", "232"
                     Set frm04010304_1.oParentForm = Me
                     frm04010304_1.Show
                  'Add By Sindy 2019/1/3
                  Case 讓與, 合併, 專利權讓與
                     frm04010302_1.m_CP118isY = True
                     frm04010302_1.Caption = "各式申請書-電子送件-讓與, 合併"
                     frm04010302_1.Show
                  'Add By Sindy 2019/1/17
                  'Modify By Sindy 2025/6/18 +案件性質444委任代理人 雅娟說,在各式申請書產生電子送件申請書(同變更申請書)
                  Case 變更, "444"
                     Set frm06010303_1.oParent = Me
                     frm06010303_1.m_CP118isY = "Y" '電子送件申請書
                     frm06010303_1.Caption = "各式申請書-電子送件-變更"
                     frm06010303_1.LoadMe Me.Tag, Text1, Text2, Text3, Text4, 41
                  'Add By Sindy 2018/7/4
                  Case 延期
                     frm04010307_1.Show
                  '2018/7/4 END
                  'Add By Sindy 2020/3/20
                  'Modified by Morgan 2023/8/21 +405申請優先權證明書,436申請優先權存取碼,437申請優先權電子交換,421申請技術報告,807第三人申請技術報告
                  Case 領證及繳年費, 年費, 延緩公告, 補換發證書, "443", "405", "436", "437", "421", "807"
                     'Modified by Morgan 2022/3/8 被舉發准後會變核駁但還是要繳 Ex:P-121502，年費誤繳也能退費不必管控 --郭
                     'If stcp10 = 領證及繳年費 Or stcp10 = 年費 Then
                     If stCP10 = 領證及繳年費 Then
                     'end 2022/3/8
                        If PUB_ApproveCheck(Me.Tag, "不可產生申請書") = False Then
                           Exit Sub
                        End If
                     
                     'Added by Morgan 2025/9/12
                     '申請台灣優先權證明書若申請人都不是台灣籍提醒
                     ElseIf stCP10 = "405" Then
                        If PUB_ChkNoTWApp(pa) = True Then
                           If PUB_TWPriCertMsg() = vbYes Then
                              Exit Sub
                           End If
                        End If
                     'end 2025/9/12
                     
                     End If
                     Set frm04010310_1.oParentForm = Me
                     frm04010310_1.m_CP118isY = True
                     frm04010310_1.m_strCP10 = stCP10
                     frm04010310_1.Show
                     frm04010310_1.Caption = "各式申請書-電子送件-其他"
                  '2020/3/20 END
                  'Added by Morgan 2023/8/23
                  Case 退費
                     frm04010306_1.m_CP118isY = True
                     frm04010306_1.Show
                     frm04010306_1.Caption = "各式申請書-電子送件-代辦退費"
                  'end 2023/8/23
                  Case Else
                     MsgBox "無申請書!", vbCritical
                     Exit Sub
               End Select
            '2009/3/24 MODIFY BY SONIA 加4.更正已繳年度
            ElseIf Me.Text6.Text = "4" Then
               If stCP10 <> 年費 Then
                  MsgBox "請選擇最近發文之年費程序 !", vbInformation
                  Exit Sub
               End If
               frm04010308_1.Show
            '2009/3/24 END
            'Add By Sindy 2018/7/4
            ElseIf Me.Text6.Text = "5" Then '5.延期
               frm04010307_1.Show
            Else
               Select Case stCP10
                  Case 讓與
                     frm04010302_1.Show
                  'Add By Cheng 2002/01/11
                  'Begin
                  Case 專利權讓與
                     frm04010302_1.Show
                  'End
                  Case 變更
                     frm06010303_1.m_CP118isY = "N" '非電子送件
                     Set frm06010303_1.oParent = Me 'Add by Morgan 2011/10/5
                     frm06010303_1.LoadMe Me.Tag, Text1, Text2, Text3, Text4, 41
                  Case 補文件
                     Set frm04010304_1.oParentForm = Me 'Add by Morgan 2011/9/22
                     frm04010304_1.Show
   'Remove by Morgan 2005/8/1 不再使用,已併入領證繳年費
   '               Case 延緩公告
   '                  frm04010305_1.Show
                  Case 退費
                     frm04010306_1.Show
                  'Modify by Amy 2014/08/15
                  'Modified by Morgan 2015/6/2 +232補優先權證明
                  Case 申請英文證明, "232"
                     'Add By Cheng 2002/07/12
                     'Begin
                     Set frm04010310_1.oParentForm = Me 'Add by Morgan 2011/9/22
                     frm04010310_1.m_strCP10 = stCP10
                     'End
                     frm04010310_1.Show
                  Case Else
                     MsgBox "選擇錯誤!", vbCritical
                     Exit Sub
               End Select
            End If
         End If 'If Me.txtKind <> "" Then    'Added by Lydia 2023/09/22
         cmdOK(1).SetFocus
         Me.Hide
      Case 1 '尋找
         Label4 = ""
         Label6 = ""
         Combo1.Clear
         MSHFlexGrid1.Clear
         GridHead
         txtKind = "": Label11 = "" 'Added by Lydia 2023/09/22
         If Text3 = "" Then Text3 = "0"
         If Text4 = "" Then Text4 = "00"
         pa(1) = Text1
         pa(2) = Text2
         pa(3) = Text3
         pa(4) = Text4
         
         If pa(1) = "P" Then
            If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
               'Modified by Lydia 2023/09/22 新增P案各式書表：不限台灣案
               'If pa(9) = 台灣國家代號 Then 'Mark by Lydia 2023/09/22 新增P案各式書表：不限台灣案
                  AddCboName Combo1, pa(5), pa(6), pa(7)
                  Text5.Text = pa(10)
                  Label4.Caption = pa(11)
                  Label6.Caption = pa(22)
                  'Added by Lydia 2023/09/22
               'Add by Morgan 2010/3/18
               'Modified by Lydia 2023/09/22 新增P案各式書表：不限台灣案
               'Else
               '   MsgBox "本案件非台灣案！"
               '   Text2.SetFocus
               '   Exit Sub
               'End If
               'end 2023/09/22
            Else
               Text2.SetFocus
               Exit Sub
            End If
         ElseIf pa(1) = "PS" Then
            If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
               AddCboName Combo1, pa(5), pa(6), pa(7)
               Text5.Text = pa(10)
               Label4.Caption = pa(11)
            Else
               Text2.SetFocus
               Exit Sub
            End If
         End If
         '申請國家
         If pa(9) <> "" Then
            If ClsPDGetNation(pa(9), strExc(1)) Then Label11.Caption = strExc(1)
         End If
         
         'Modified by Lydia 2023/09/22 cpm03 => IIf(pa(9) = "000", " cpm03 ", " cpm04 as cpm03 ") & "
         strExc(0) = "select ''," & SQLDate("CP05") & ",cp09," & IIf(pa(9) = "000", " cpm03 ", " cpm04 as cpm03 ") & ",staff.st02 as st1,staff1.st02 as st2," & _
            "cp64,cp10 from caseprogress, casepropertymap,staff,staff staff1 where " & _
            ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and " & _
            "( cp09<'C' ) and cp01=cpm01(+) and " & _
            "cp10=cpm02(+) and cp14=staff.st01(+) and cp13=staff1.st01(+)" & _
            " order by cp05 desc,cp09 desc"
         intI = 0
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
         GridHead
         'Add By Cheng 2002/05/10
         '若只搜尋到一筆時直接勾選
         'Modified by Lydia 2023/09/24 限臺灣案
         If Me.MSHFlexGrid1.Rows = 2 And pa(9) = 台灣國家代號 Then
            MSHFlexGrid1_Click
         End If
      Case 2
         Unload Me
   End Select
End Sub

Private Sub Form_Activate()
 Dim i As Integer, j As Integer
 'Added by Morgan 2020/4/8
 Static bolDone As Boolean
 
   If bolDone = False Then
      If Text2 <> "" Then
         cmdOK(1).Value = True
         intI = 0
         With MSHFlexGrid1
         For i = 1 To .Rows - 1
            If .TextMatrix(i, 2) = stCP09 Then
               .TextMatrix(i, 0) = "v"
               intI = 1
               Exit For
            End If
         Next
         End With
         If intI = 1 Then
            Text6 = "3"
            cmdOK(0).Value = True
         End If
      End If
      bolDone = True
      Exit Sub
   ElseIf stCP09 <> "" Then
      Unload Me
      Exit Sub
   End If
'end 2020/4/8
 
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = ""
         For j = 0 To .Cols - 1
            .col = j
            .CellBackColor = .BackColor
         Next
      Next
   End With
   'Text6.Text = ""
   
   'Add by Morgan 2011/9/22
   'Mark by Lydia 2023/09/22 承辦人改用frm04010301_1
   'If iFrom = 1 Then
   '   Label9.Visible = False
   '   Text6.Visible = False
   '   Label10.Visible = False
   'End If
   'end 2023/09/22
   
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   intWhere = 國內
'   Combo1.ListIndex = 0
   Label4 = ""
   Label6 = ""
   InitGrid 8, MSHFlexGrid1
   GridHead
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040103_1 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 1) <> "" Then
      GridClick MSHFlexGrid1, intLastRow, 0
      'Add By Sindy 2018/6/29
      If MSHFlexGrid1.TextMatrix(intLastRow, 0) = "v" Then
         Text6.Text = ""
         If GetCP10(MSHFlexGrid1.TextMatrix(intLastRow, 2), "CP118") = "Y" Or _
            GetCP10(MSHFlexGrid1.TextMatrix(intLastRow, 2), "CP118") = "A" Then
            Text6.Text = "3" '3.電子送件
         End If
      End If
      '2018/6/29 END
      cmdOK(0).SetFocus
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 <> "P" And Text1 <> "PS" Then
      MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
      TextInverse Text1
      Cancel = True
   End If
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1200: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1400: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1200: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1400: .Text = "智權人員"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1400: .Text = "案件備註"
      .col = 7: .ColWidth(7) = 0
      .Visible = True
      If .Rows > 1 Then .row = 1
   End With
End Sub

'讀取案件性質
Private Function GetCP10(p_CP09 As String, Optional strCol As String = "CP10") As String
   Dim stSQL As String, iRtn As Integer
   
   GetCP10 = ""
   If p_CP09 <> "" Then
      stSQL = "select " & strCol & " from caseprogress where cp09='" & p_CP09 & "'"
      iRtn = 1
      Set AdoRecordSet3 = ClsLawReadRstMsg(iRtn, stSQL)
      If iRtn = 1 Then
         GetCP10 = "" & AdoRecordSet3.Fields(0)
      End If
   End If
End Function

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If Text3 = "" Then Text3 = "0"
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If Text4 = "" Then Text4 = "00"
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modify By Cheng 2002/07/12
'   If KeyAscii <> 49 And KeyAscii <> 8 Then
    'Modify By Cheng 2002/12/17
'   If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
'Remove by Morgan 2005/8/1 延緩公告"3"不再使用,已併入領證繳年費
'   If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 Then
'2009/3/24 MODIFY BY SONIA 加4.更正已繳年度
   'Modify By Sindy 2018/6/29 + And KeyAscii <> 51
   'Modify By Sindy 2018/8/1 + And KeyAscii <> 53
   If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text1_GotFocus()
   InverseTextBox Text1
End Sub

Private Sub Text2_GotFocus()
   InverseTextBox Text2
End Sub

Private Sub Text3_GotFocus()
   InverseTextBox Text3
End Sub

Private Sub Text4_GotFocus()
   InverseTextBox Text4
End Sub

Private Sub Text5_GotFocus()
   InverseTextBox Text5
End Sub

Private Sub Text6_GotFocus()
   InverseTextBox Text6
End Sub

Public Sub ClearForm()
   'Text1 = Empty 'Mark by Lydia 2023/09/22
   Text2 = Empty
   Text3 = Empty
   Text4 = Empty
   Text5 = Empty
   Label4 = Empty
   Label6 = Empty
   Text6 = Empty
   Combo1.Clear
   InitGrid 8, MSHFlexGrid1
   GridHead
   Text1.SetFocus
   
   'Added by Lydia 2023/09/22
   txtKind = Empty
   Label11 = Empty
End Sub

'Added by Lydia 2023/09/22
Private Sub txtkind_GotFocus()
   InverseTextBox txtKind
End Sub

Private Sub txtKind_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub
