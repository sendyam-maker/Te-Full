VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm110103_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "閉卷"
   ClientHeight    =   3888
   ClientLeft      =   1356
   ClientTop       =   1608
   ClientWidth     =   6672
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3888
   ScaleWidth      =   6672
   Begin VB.TextBox txtNumber1 
      Height          =   285
      Left            =   2130
      TabIndex        =   12
      Top             =   3150
      Width           =   3165
   End
   Begin VB.TextBox txtNumber2 
      Height          =   285
      Left            =   1455
      TabIndex        =   11
      Top             =   2820
      Width           =   3840
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   4
      Left            =   1890
      MaxLength       =   6
      TabIndex        =   1
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Frame fraChoose 
      BorderStyle     =   0  '沒有框線
      Height          =   345
      Index           =   2
      Left            =   1452
      TabIndex        =   24
      Top             =   2430
      Width           =   3150
      Begin VB.Frame fraElse 
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   735
         TabIndex        =   25
         Top             =   24
         Width           =   2625
         Begin VB.TextBox txtCode 
            Height          =   288
            Index           =   0
            Left            =   0
            MaxLength       =   6
            TabIndex        =   8
            Top             =   0
            Width           =   1212
         End
         Begin VB.TextBox txtCode 
            Height          =   288
            Index           =   1
            Left            =   1320
            MaxLength       =   1
            TabIndex        =   9
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtCode 
            Height          =   288
            Index           =   2
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   10
            Top             =   0
            Width           =   492
         End
      End
      Begin VB.TextBox txtSystem 
         Height          =   288
         Left            =   0
         MaxLength       =   3
         TabIndex        =   7
         Top             =   30
         Width           =   732
      End
      Begin VB.Frame fraTF 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   780
         TabIndex        =   26
         Top             =   15
         Width           =   2652
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   0
            Left            =   0
            MaxLength       =   5
            TabIndex        =   13
            Top             =   0
            Width           =   972
         End
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   1
            Left            =   1080
            MaxLength       =   1
            TabIndex        =   14
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   2
            Left            =   1560
            MaxLength       =   1
            TabIndex        =   15
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   3
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   16
            Top             =   0
            Width           =   492
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5688
      TabIndex        =   19
      Top             =   96
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4860
      TabIndex        =   18
      Top             =   96
      Width           =   800
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   0
      Left            =   1140
      TabIndex        =   0
      Top             =   648
      Width           =   3612
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "本所案號："
      CausesValidation=   0   'False
      Height          =   180
      Index           =   2
      Left            =   144
      TabIndex        =   6
      Top             =   2520
      Width           =   1212
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "申請人："
      CausesValidation=   0   'False
      Height          =   180
      Index           =   1
      Left            =   144
      TabIndex        =   4
      Top             =   2070
      Width           =   1095
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "代理人："
      CausesValidation=   0   'False
      Height          =   180
      Index           =   0
      Left            =   132
      TabIndex        =   2
      Top             =   1620
      Width           =   1095
   End
   Begin VB.Frame fraChoose 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Index           =   1
      Left            =   1260
      TabIndex        =   28
      Top             =   2040
      Width           =   1332
      Begin VB.TextBox txtCaseField 
         Height          =   264
         Index           =   2
         Left            =   0
         MaxLength       =   9
         TabIndex        =   5
         Top             =   0
         Width           =   1212
      End
   End
   Begin VB.Frame fraChoose 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Index           =   0
      Left            =   1248
      TabIndex        =   27
      Top             =   1590
      Width           =   1332
      Begin VB.TextBox txtCaseField 
         Height          =   264
         Index           =   1
         Left            =   0
         MaxLength       =   9
         TabIndex        =   3
         Top             =   0
         Width           =   1212
      End
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1440
      TabIndex        =   17
      Top             =   3465
      Width           =   5130
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "9049;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblNumber 
      Caption         =   "審定號數/證書號數："
      Height          =   180
      Left            =   405
      TabIndex        =   32
      Top             =   3195
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "申請案號："
      Height          =   180
      Left            =   405
      TabIndex        =   31
      Top             =   2835
      Width           =   975
   End
   Begin MSForms.Label lblSName 
      Height          =   180
      Left            =   2970
      TabIndex        =   30
      Top             =   1140
      Width           =   1935
      VariousPropertyBits=   27
      Size            =   "3625;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "閉卷指示智權人員："
      Height          =   180
      Left            =   180
      TabIndex        =   29
      Top             =   1140
      Width           =   1635
   End
   Begin MSForms.Label lblCustomer 
      Height          =   180
      Left            =   2595
      TabIndex        =   22
      Top             =   2085
      Width           =   3870
      VariousPropertyBits=   27
      Size            =   "3625;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblAgent 
      Height          =   180
      Left            =   2595
      TabIndex        =   21
      Top             =   1635
      Width           =   3870
      VariousPropertyBits=   27
      Size            =   "3625;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   0
      Left            =   405
      TabIndex        =   23
      Top             =   3510
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Left            =   180
      TabIndex        =   20
      Top             =   684
      Width           =   1092
   End
End
Attribute VB_Name = "frm110103_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/5/10 改成Form2.0(lblSName,lblAgent,lblCustomer,cboCaseName)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'sonia 2010/8/19 日期欄已修改
Option Explicit
'edit by nickc 2007/02/06 不用 dll 了
'Dim obj011 As Object

Private Sub cmdok_Click(Index As Integer)
Dim varSaveCursor, i As Integer

Select Case Index
             Case 0
                        'Add By Cheng 2002/01/08
                        txtCaseField_LostFocus 0
                        varSaveCursor = Screen.MousePointer
                        Screen.MousePointer = vbHourglass
                        For i = 0 To 2
                               If txtCaseField(i).Enabled Then
                                  If CheckKeyIn(i) <> 1 Then
                                     txtCaseField(i).SetFocus
                                     txtCaseField_GotFocus (i)
                                     'Modify By Cheng 2002/05/29
'                                     Exit For
                                    Screen.MousePointer = varSaveCursor
                                    Exit Sub
                                  End If
                               End If
                        Next
                        If optChoose(2).Value Then
                           If CheckKeyIn(i) <> 1 Then
                               txtSystem.SetFocus
                               txtSystem_GotFocus
                               'Add By Cheng 2002/05/29
                               Screen.MousePointer = varSaveCursor
                               Exit Sub
                           Else
                               i = i + 1
                           End If
                        Else
                           i = i + 1
                        End If
                        'Add By Cheng 2002/05/29
                        If CheckKeyIn(4) <> 1 Then
                           txtCaseField(4).SetFocus
                           txtCaseField_GotFocus (4)
                           Screen.MousePointer = varSaveCursor
                           Exit Sub
                        End If
                        
                        If i = 4 Then
                           If optChoose(2).Value Then
                           
                              'Add by Morgan 2010/7/15
                              If txtNumber1.Locked = False And txtNumber1 <> "" And txtNumber1 <> txtNumber1.Tag Then
                                 MsgBox "證書號數錯誤，請重新輸入！"
                                 txtNumber1_GotFocus
                                 txtNumber1.SetFocus
                                 Screen.MousePointer = varSaveCursor
                                 Exit Sub
                              ElseIf txtNumber2.Locked = False And txtNumber2 <> "" And txtNumber2 <> txtNumber2.Tag Then
                                 MsgBox "申請案號錯誤，請重新輸入！"
                                 txtNumber2_GotFocus
                                 txtNumber2.SetFocus
                                 Screen.MousePointer = varSaveCursor
                                 Exit Sub
                              ElseIf txtNumber1.Locked = False And txtNumber1 & txtNumber2 = "" Then
                                 MsgBox "請輸入申請案號或證書號數！"
                                 txtNumber2.SetFocus
                                 Screen.MousePointer = varSaveCursor
                                 Exit Sub
                              End If
                              'end 2010/7/15
                              
                              frm110103_3.intWhereComeFrom = 1
                              Set frm110103_3.mPrev01 = Me 'Add By Sindy 2015/2/13
                              frm110103_3.Show
                           Else
                              If optChoose(0).Value Then
                                 frm110103_2.lblTitle = "代理人："
                                 frm110103_2.LblNo = txtCaseField(1)
                                 frm110103_2.lblName = lblAgent
                              Else
                                 frm110103_2.lblTitle = "申請人："
                                 frm110103_2.LblNo = txtCaseField(2)
                                 frm110103_2.lblName = lblCustomer
                              End If
                              frm110103_2.Show
                           End If
                           Me.Hide
                        End If
                        Screen.MousePointer = varSaveCursor
             Case 1
                        Unload Me
End Select
End Sub
Private Sub Form_Load()
Dim strSystemKind As String

MoveFormToCenter Me
optChoose(0).Value = True
'edit by nickc 2007/02/06 不用 dll 了
'If obj011 Is Nothing Then
'   Set obj011 = CreateObject("prjTaieDll011.cls011")
'   Set obj011.Connection = cnnConnection
'End If
'edit by nickc 2007/02/02 不用 dll 了
'objPublicData.GetGroupSystemKind strGroup, strSystemKind
ClsPDGetGroupSystemKind strGroup, strSystemKind
txtCaseField(0) = strSystemKind
'Added by Lydia 2016/01/04 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)

End Sub
Private Sub Form_Unload(Cancel As Integer)
'edit by nickc 2007/02/06 不用 dll 了
'Set obj011 = Nothing
   'Add By Cheng 2002/07/18
   Set frm110103_1 = Nothing
End Sub
Private Sub optChoose_Click(Index As Integer)
fraChoose(Index).Enabled = True
fraChoose((Index + 1) Mod 3).Enabled = False
fraChoose((Index + 2) Mod 3).Enabled = False
If Index = 2 Then
   txtCaseField(0).Enabled = False
   cboCaseName.Enabled = True
   txtSystem.SetFocus
Else
   txtCaseField(0).Enabled = True
   cboCaseName.Enabled = False
   If txtCaseField(Index + 1).Visible Then txtCaseField(Index + 1).SetFocus
End If
End Sub
Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
             Case 0, 1, 2, 4
                       KeyAscii = UpperCase(KeyAscii)
End Select
End Sub
Private Sub txtCaseField_Change(Index As Integer)
Select Case Index
             Case 1
                        lblAgent = ""
             Case 2
                        lblCustomer = ""
End Select
End Sub

Private Sub txtCaseField_LostFocus(Index As Integer)
'Add By Cheng 2002/01/08
Select Case Index
Case 0 '系統類別
   Me.txtCaseField(Index).Text = GetAllSysKind(Me.txtCaseField(Index))
Case 4 '閉卷指示智權人員
   If Len("" & Me.txtCaseField(Index).Text) > 0 Then
      If CheckKeyIn(Index) = -1 Then
         Me.txtCaseField(Index).SetFocus
         txtCaseField_GotFocus (Index)
      End If
   End If
End Select
End Sub

Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
'Add By Cheng 2002/05/29
'閉卷指示智權人員欄位有效性不在此段檢查
If Index = 4 Then Exit Sub

If CheckKeyIn(Index) = -1 Then
   Cancel = True
End If
If Cancel Then txtCaseField_GotFocus (Index)
End Sub

Private Sub txtNumber1_GotFocus()
   TextInverse txtNumber1
End Sub

Private Sub txtNumber2_GotFocus()
   TextInverse txtNumber2
End Sub

Private Sub txtSystem_Change()
If txtSystem.Text = 馬德里案 Then
   fraTF.Visible = True
   fraElse.Visible = False
Else
   fraTF.Visible = False
   fraElse.Visible = True
End If
If cboCaseName.ListCount > 0 Then cboCaseName.Clear
End Sub
Private Sub txtSystem_GotFocus()
txtSystem.SelStart = 0
txtSystem.SelLength = Len(txtSystem.Text)
End Sub
Private Sub txtSystem_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txtSystem_Validate(Cancel As Boolean)
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.GetGroupCase(txtSystem, strGroup) = False Then
'Added by Lydia 2016/01/04 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件(P)，但非此類案件時外專程序人員不可操作。
If Not (FMP2open = True And (txtSystem = "P" Or txtSystem = "PS")) Then
    '原程式
    If ClsPDGetGroupCase(txtSystem, strGroup) = False Then
       ShowMsg MsgText(1056)
       Cancel = True
       txtSystem_GotFocus
    End If
End If
'end 2016/01/04
End Sub
Private Sub txtCode_Change(Index As Integer)
If cboCaseName.ListCount > 0 Then cboCaseName.Clear
End Sub
Private Sub txtTFCode_Change(Index As Integer)
If cboCaseName.ListCount > 0 Then cboCaseName.Clear
End Sub
Private Sub txtTFCode_GotFocus(Index As Integer)
txtTFCode(Index).SelStart = 0
txtTFCode(Index).SelLength = Len(txtTFCode(Index).Text)
End Sub
Private Sub txtTFCode_Validate(Index As Integer, Cancel As Boolean)
CheckKeyIn1 (Index)
End Sub
Private Sub txtCaseField_GotFocus(Index As Integer)
txtCaseField(Index).SelStart = 0
txtCaseField(Index).SelLength = Len(txtCaseField(Index).Text)
End Sub
Private Sub txtCode_GotFocus(Index As Integer)
txtCode(Index).SelStart = 0
txtCode(Index).SelLength = Len(txtCode(Index).Text)
End Sub
Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
CheckKeyIn2 (Index)
End Sub
Private Function CheckKeyIn1(ByRef intIndex As Integer) As Boolean
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String
Dim strNumber1 As String, strNumber2 As String

'2011/3/30 modify by sonia
'If Len(txtTFCode(intIndex)) > 0 And Len(txtTFCode(intIndex)) < txtTFCode(intIndex).MaxLength Then
'   ShowMsg MsgText(9)
'ElseIf intIndex = 3 Then
If intIndex = 3 Then
'2011/3/30 end
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
         IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3) Then
  'Added by Lydia 2016/01/04 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   If FMP2open = True Then
      If PUB_FMPtoCheck(0, 1, Pub_strUserST05, txtSystem, txtTFCode(0), txtTFCode(1), txtTFCode(2)) = False Then
        txtTFCode(0).SetFocus
        Exit Function
      End If
   End If
   If ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
         IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, , , strNumber1, strNumber2) Then
      SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
      'Add by Morgan 2010/7/15 FCP,P,CFP 要輸入申請號或證書號
      txtNumber1.Tag = strNumber1
      txtNumber2.Tag = strNumber2
      If txtSystem = "FCP" Or txtSystem = "P" Or txtSystem = "CFP" Then
         txtNumber1.Locked = False
         txtNumber2.Locked = False
      Else
         txtNumber1.Locked = True
         txtNumber2.Locked = True
         txtNumber1 = strNumber1
         txtNumber2 = strNumber2
      End If
      'end 2010/7/15
      
      CheckKeyIn1 = True
   End If
Else
   CheckKeyIn1 = True
End If
End Function
Private Function CheckKeyIn2(ByRef intIndex As Integer) As Boolean
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String
Dim strNumber1 As String, strNumber2 As String

'2011/3/30 modify by sonia
'If Len(txtCode(intIndex)) > 0 And Len(txtCode(intIndex)) < txtCode(intIndex).MaxLength Then
'   ShowMsg MsgText(9)
'ElseIf intIndex = 2 Then
If intIndex = 2 Then
'2011/3/30 end
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3) Then
  'Added by Lydia 2016/01/04 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   If FMP2open = True Then
      If PUB_FMPtoCheck(0, 1, Pub_strUserST05, txtSystem, txtCode(0), txtCode(1), txtCode(2)) = False Then
        txtCode(0).SetFocus
        Exit Function
      End If
   End If
   If ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, , , strNumber1, strNumber2) Then
      SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
      'Add by Morgan 2010/7/15 FCP,P,CFP 要輸入申請號或證書號
      txtNumber1.Tag = strNumber1
      txtNumber2.Tag = strNumber2
      If txtSystem = "FCP" Or txtSystem = "P" Or txtSystem = "CFP" Then
         txtNumber1.Locked = False
         txtNumber2.Locked = False
      Else
         txtNumber1.Locked = True
         txtNumber2.Locked = True
         txtNumber1 = strNumber1
         txtNumber2 = strNumber2
      End If
      'end 2010/7/15
      
      CheckKeyIn2 = True
   End If
Else
   CheckKeyIn2 = True
End If
End Function
Private Function CheckKeyIn(intIndex As Integer) As Integer
Dim strTemp As String, strCusTemp As String, bolRt As Boolean, varSystemTemp As Variant, i As Integer
Dim Rs As New ADODB.Recordset
'Add By Cheng 2002/07/09
Dim strTemp1

CheckKeyIn = -1
Select Case intIndex
             Case 0 '系統類別
                        If txtCaseField(intIndex) <> "" Then
                           varSystemTemp = Split(txtCaseField(intIndex), ",")
                           For i = 0 To UBound(varSystemTemp)
                                  'edit by nickc 2007/02/02 不用 dll 了
                                  'If objPublicData.GetGroupCase(CStr(varSystemTemp(i)), strGroup) = False Then
                                  'Added by Lydia 2016/01/04 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件(P)，但非此類案件時外專程序人員不可操作。
                                  If Not (FMP2open = True And (CStr(varSystemTemp(i)) = "P" Or CStr(varSystemTemp(i)) = "PS")) Then
                                    '原程式
                                    If ClsPDGetGroupCase(CStr(varSystemTemp(i)), strGroup) = False Then
                                       ShowMsg MsgText(1056)
                                       Exit For
                                    Else
                                       CheckKeyIn = 1
                                    End If
                                  Else
                                    CheckKeyIn = 1
                                  End If
                                  'end 2016/01/04
                            Next
                        End If
             Case 1 '代理人代號
                        strCusTemp = txtCaseField(intIndex)
                        'Modify By Cheng 2002/07/09
                        strTemp1 = Split(Me.txtCaseField(0).Text & " ", ",")
'                        If objPublicData.GetAgent(strCusTemp, strTemp) Then
                        If PUB_GetAgentName(IIf(Me.txtCaseField(0).Text = "", "", strTemp1(0)), strCusTemp, strTemp) Then
                           txtCaseField(intIndex) = strCusTemp
                           lblAgent.Caption = strTemp
                           CheckKeyIn = 1
                        End If
             Case 2 '申請人代號
                        strCusTemp = txtCaseField(intIndex)
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetCustomer(strCusTemp, strTemp) Then
                        If ClsPDGetCustomer(strCusTemp, strTemp) Then
                           txtCaseField(intIndex) = strCusTemp
                           lblCustomer.Caption = strTemp
                           CheckKeyIn = 1
                        End If
             Case 3
                        If txtSystem = 馬德里案 Then
                           bolRt = CheckKeyIn1(3)
                        Else
                           bolRt = CheckKeyIn2(2)
                        End If
                        If bolRt Then CheckKeyIn = 1
             Case 4 '閉卷指示智權人員
                        If Len(Me.txtCaseField(intIndex).Text) <= 0 Then
                           MsgBox "請輸入閉卷指示智權人員!!!", vbExclamation + vbOKOnly
                           Exit Function
                        End If
                        If Rs.State <> adStateClosed Then Rs.Close
                        Set Rs = Nothing
                        Rs.CursorLocation = adUseClient
                        Rs.Open "Select st02 From Staff Where ST01='" & Me.txtCaseField(4).Text & "' And ST04='1' ", cnnConnection, adOpenStatic, adLockReadOnly
                        If Rs.RecordCount > 0 Then
                           Me.lblSName.Caption = "" & Rs.Fields(0).Value
                           CheckKeyIn = 1
                        Else
                           MsgBox "閉卷指示智權人員輸入錯誤!!!", vbExclamation + vbOKOnly
                           Me.lblSName.Caption = ""
                        End If
                        If Rs.State <> adStateClosed Then Rs.Close
                        Set Rs = Nothing
End Select
End Function

Public Sub Cleartxt()
txtCaseField(1) = ""
txtCaseField(2) = ""
txtSystem = ""
txtCode(0) = ""
txtCode(1) = ""
txtCode(2) = ""
txtNumber1 = "" 'Added by Morgan 2025/3/28
txtNumber2 = "" 'Added by Morgan 2025/3/28
cboCaseName.Clear
'Modify By Cheng 2002/05/29
'If optChoose(0).Value = True Or optChoose(1).Value = True Then
'   txtCaseField(1).SetFocus
If optChoose(0).Value = True Then
   txtCaseField(1).SetFocus
ElseIf Me.optChoose(1).Value Then
   txtCaseField(2).SetFocus
Else
   txtSystem.SetFocus
End If
End Sub
