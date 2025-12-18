VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm072001 
   BorderStyle     =   1  '單線固定
   Caption         =   "庭期資料查詢"
   ClientHeight    =   4560
   ClientLeft      =   1176
   ClientTop       =   1776
   ClientWidth     =   5916
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5916
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   4860
      Style           =   1  '圖片外觀
      TabIndex        =   30
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   4032
      Style           =   1  '圖片外觀
      TabIndex        =   29
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox txtcp04 
      Height          =   288
      Left            =   3144
      MaxLength       =   2
      TabIndex        =   3
      Top             =   564
      Width           =   375
   End
   Begin VB.TextBox txtcp03 
      Height          =   288
      Left            =   2832
      MaxLength       =   1
      TabIndex        =   2
      Top             =   564
      Width           =   255
   End
   Begin VB.TextBox txtcp02 
      Height          =   288
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   1
      Top             =   564
      Width           =   855
   End
   Begin VB.TextBox txtcp01 
      Height          =   288
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   0
      Top             =   564
      Width           =   550
   End
   Begin MSForms.ComboBox cboGov 
      Height          =   300
      Index           =   1
      Left            =   3720
      TabIndex        =   9
      Top             =   2040
      Width           =   2004
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3535;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboGov 
      Height          =   300
      Index           =   0
      Left            =   1320
      TabIndex        =   8
      Top             =   2040
      Width           =   2000
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3528;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboPerson 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   1275
      Width           =   1575
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2778;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   13
      Left            =   1320
      TabIndex        =   18
      Top             =   3864
      Width           =   4476
      VariousPropertyBits=   671105051
      MaxLength       =   80
      Size            =   "7895;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   12
      Left            =   1320
      TabIndex        =   17
      Top             =   3496
      Width           =   2775
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "4895;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   11
      Left            =   1320
      TabIndex        =   16
      Top             =   3131
      Width           =   2775
      VariousPropertyBits=   671105051
      MaxLength       =   80
      Size            =   "4895;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   10
      Left            =   3840
      TabIndex        =   15
      Top             =   2766
      Width           =   1200
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "2117;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   9
      Left            =   2568
      TabIndex        =   14
      Top             =   2766
      Width           =   1200
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "2117;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   8
      Left            =   1320
      TabIndex        =   13
      Top             =   2766
      Width           =   1200
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "2117;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   7
      Left            =   3840
      TabIndex        =   12
      Top             =   2401
      Width           =   1200
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "2117;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   6
      Left            =   2568
      TabIndex        =   11
      Top             =   2401
      Width           =   1200
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "2117;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   5
      Left            =   1320
      TabIndex        =   10
      Top             =   2401
      Width           =   1200
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "2117;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   1
      Left            =   1320
      TabIndex        =   6
      Top             =   1671
      Width           =   1332
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "2350;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   0
      Left            =   1320
      TabIndex        =   4
      Top             =   929
      Width           =   1812
      VariousPropertyBits=   671105051
      MaxLength       =   32
      Size            =   "3196;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   2
      Left            =   3000
      TabIndex        =   7
      Top             =   1671
      Width           =   1335
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "2355;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbeName 
      Height          =   285
      Left            =   3000
      TabIndex        =   31
      Top             =   1290
      Width           =   1935
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "3408;508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      Caption         =   "當事人(日)："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   28
      Top             =   3918
      Width           =   1100
   End
   Begin VB.Label Label10 
      Caption         =   "當事人(英)："
      Height          =   180
      Left            =   120
      TabIndex        =   27
      Top             =   3546
      Width           =   1100
   End
   Begin VB.Label Label9 
      Caption         =   "當事人(中)："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   3180
      Width           =   1100
   End
   Begin VB.Label Label8 
      Caption         =   "檢  察  官："
      Height          =   180
      Left            =   120
      TabIndex        =   25
      Top             =   2814
      Width           =   972
   End
   Begin VB.Label Label7 
      Caption         =   "法       官："
      Height          =   180
      Left            =   120
      TabIndex        =   24
      Top             =   2448
      Width           =   852
   End
   Begin VB.Label Label6 
      Caption         =   "機關代號："
      Height          =   180
      Left            =   120
      TabIndex        =   23
      Top             =   2082
      Width           =   972
   End
   Begin VB.Label Label5 
      Caption         =   "開庭日期："
      Height          =   180
      Left            =   120
      TabIndex        =   22
      Top             =   1716
      Width           =   1092
   End
   Begin VB.Label Label4 
      Caption         =   "開庭人員："
      Height          =   180
      Left            =   120
      TabIndex        =   21
      Top             =   1350
      Width           =   1092
   End
   Begin VB.Label Label3 
      Caption         =   "法院案號："
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   984
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號："
      Height          =   180
      Left            =   120
      TabIndex        =   19
      Top             =   618
      Width           =   972
   End
   Begin VB.Line Line2 
      X1              =   3408
      X2              =   3648
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      X1              =   2712
      X2              =   2952
      Y1              =   1824
      Y2              =   1824
   End
End
Attribute VB_Name = "frm072001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/15 改成Form2.0 ; lbeName、Text(index)、cboPerson
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim blnCom1 As Boolean, blnCom2 As Boolean, blnCom3 As Boolean, blnCom4 As Boolean
'Add By Cheng 2002/09/09
Dim blnClkSure As Boolean '是否按下確定按鈕
Dim m_GovListNew(0 To 1) As String, m_GovListDef(0 To 1) As String  'Added by Lydia 2025/11/18

'Add By Sindy 2010/11/26
'Modified by Lydia 2021/09/15 改成Form 2.0
'Private Sub cboPerson_KeyPress(KeyAscii As Integer)
Private Sub cboPerson_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cmdBack_Click()
   Unload Me
   Set frm072001 = Nothing
End Sub

Private Sub cmdSure_Click()
  Dim i As Integer
  Dim n As Integer
  Dim tmpBol As Boolean 'Added by Lydia 2025/11/18
  
  'Add By Cheng 2002/09/09
  blnClkSure = False
  Screen.MousePointer = 11
  n = 0
  If txtCP01.Text = "" And txtCP01.Text = "" And txtCP01.Text = "" And txtCP01.Text = "" And cboPerson.Text = "" Then
     For i = 0 To 13
         If Text(i).Text <> "" Then
            n = 1
            Exit For
         Else
           n = 2
         End If
     Next
  End If
  If n = 2 Then
     MsgBox "至少輸入一個條件!", vbInformation, "庭期資料查詢"
     Screen.MousePointer = vbDefault
     Exit Sub
  End If
   'Add By Cheng 2002/03/21
   If PUB_CheckKeyInDate(Me.Text(1)) = -1 Then
      Me.Text(1).SetFocus
      Text_GotFocus 1
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   If PUB_CheckKeyInDate(Me.Text(2)) = -1 Then
      Me.Text(2).SetFocus
      Text_GotFocus 2
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   'Add By Cheng 2002/09/09
   If Me.Text(1).Text <> "" And Me.Text(2).Text <> "" Then
      If Val(Me.Text(1).Text) > Val(Me.Text(2).Text) Then
         MsgBox "開庭日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         blnClkSure = True
         Me.Text(1).SetFocus
         Text_GotFocus 1
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
   End If
   'Modified by Lydia 2025/11/18 改成下拉選單
   'If Me.Text(3).Text <> "" And Me.Text(4).Text <> "" Then
   '   If Me.Text(3).Text > Me.Text(4).Text Then
   '      MsgBox "機關代號範圍輸入錯誤!!!", vbExclamation + vbOKOnly
   '      blnClkSure = True
   '      Me.Text(3).SetFocus
   '      Text_GotFocus 3
   '      Screen.MousePointer = vbDefault
   '      Exit Sub
   '   End If
   'End If
   For n = 0 To 1
      If Me.cboGov(n).Text <> "" Then
         Call cboGov_Validate(n, tmpBol)
         If tmpBol = True Then
            MsgBox "機關代號範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            blnClkSure = True
            cboGov_GotFocus n
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
      End If
   Next n
   
   TestSub
   Screen.MousePointer = 0
End Sub

Private Sub TestSub()

   If Text(1).Text <> "" And Text(2).Text <> "" Then
      If Val(Text(1).Text) > Val(Text(2).Text) Then
         MsgBox "開庭日期範圍不正確 !", vbCritical
         Text(2).SetFocus
         Exit Sub
      End If
   End If
   'Modified by Lydia 2025/11/18 改成下拉選單
   'If Text(3).Text <> "" And Text(4).Text <> "" Then
   '   If Text(3).Text > Text(4).Text Then
   '      MsgBox "機關代號範圍不正確 !", vbCritical
   '      Text(4).SetFocus
    '     Exit Sub
    '  End If
   'End If
   If Trim(cboGov(0).Text) <> "" And Trim(cboGov(1).Text) <> "" Then
      If Trim(Left(cboGov(0).Text, 3)) > Trim(Left(cboGov(1).Text, 3)) Then
         MsgBox "機關代號範圍不正確 !", vbCritical
         cboGov(1).SetFocus
         Exit Sub
      End If
   End If
   'end 2025/11/18
   
   'Modify By Sindy 2011/6/17 +其他出庭律師資料檔
   '若已閉卷, 則在本所案號後加"*"號
   'Modify By Sindy 2016/6/27 + '' V,
   'Modify by Amy 2018/01/24 開庭別+調解庭,開庭種類+調解
   strExc(0) = "select '' V,cp01||'-'||cp02||decode(cp03,'0','','-'||cp03)||decode(cp04,'00','','-'||cp04)||DECODE(HC09,'Y','＊','')," & _
               "hc06,decode(cp05,null,'',substr(cp05,1,4)-1911||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2))," & _
               "decode(cdp02,st01,st02),decode(cdp05,or01,or02),cp35," & _
               "decode(cdp18,null,'','**')||sqldatet(cdp03) cdp03," + _
               "substr(sqltime(cdp04||'00'),1,5),decode(cdp17,'1','民事庭','2','偵查庭','3','刑事庭','4','刑附民庭','5','行政庭','6','調解庭'),decode(cdp06,'1','偵查','2','審查','3','言詞辯論','4','調查','5','調解'),cdp08,cdp09,cdp07,cp09,'' " + _
               "from hirecase,courtyardperiod,customer,caseprogress,staff,organization,caselawer " & _
               "where cdp01=cp09(+) and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) " & _
               "and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0','','0',substr(hc05,9,1))=cu02(+) " & _
               "and cp09>='C' and cp01='LA' and cdp02=st01(+) and cdp05=or01(+) and cdp01=cl01(+)" & ReadSQL
   'Modify By Sindy 2009/07/24 增加LIN系統類別
   'Modify By Sindy 2016/6/27 + '' V,
   'Modify by Amy 2018/01/24 開庭別+調解庭,開庭種類+調解
   'modify by sonia 2019/7/29 +ACS系統類別
   strExc(0) = strExc(0) & " union select '' V,cp01||'-'||cp02||decode(cp03,'0','','-'||cp03)||decode(cp04,'00','','-'||cp04)||DECODE(LC08,'Y','＊','')," & _
               "nvl(lc05,nvl(lc06,lc07)),decode(cp05,null,'',substr(cp05,1,4)-1911" + _
               " ||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2)),decode(cdp02,st01,st02)," & _
               "decode(cdp05,or01,or02),cp35,decode(cdp18,null,'','**')||sqldatet(cdp03) cdp03," + _
               "substr(sqltime(cdp04||'00'),1,5),decode(cdp17,'1','民事庭','2','偵查庭','3','刑事庭','4','刑附民庭','5','行政庭','6','調解庭'),decode(cdp06,'1','偵查','2','審查','3','言詞辯論','4','調查','5','調解'),cdp08,cdp09,cdp07,cp09,'' " + _
               "From lawcase,courtyardperiod,customer,caseprogress,staff,organization,caselawer " & _
               "where cdp01=cp09(+) and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) " & _
               "and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0','','0',substr(lc11,9,1))=cu02(+) " & _
               "and cp09>='C' and cp01 in ('L','FCL','CFL','LIN','ACS') and cdp02=st01(+) and cdp05=or01(+) and cdp01=cl01(+)" & ReadSQL
   '2011/6/29 modify by sonia 改排序條件
   'strExc(0) = strExc(0) & " order by cp09 desc "
   strExc(0) = strExc(0) & " order by 1,3,cp09 "
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))    'edit by nickc 2007/02/07 不用 dll 了 Set rstemp = objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      frm072002.Show
      Me.Hide
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   GetStaff
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm072001 = Nothing
End Sub

Private Sub Text_GotFocus(Index As Integer)
   TextInverse Text(Index)
   Select Case Index
      Case 5, 6, 7, 8, 9, 10, 11, 13
           'edit by nickc 2007/06/11  切換輸入法改用API
           'Text(Index).IMEMode = 1
           OpenIme
      Case Else
           'edit by nickc 2007/06/11  切換輸入法改用API
           'Text(Index).IMEMode = 2
           CloseIme
   End Select
End Sub

'Modified by Lydia 2021/09/15 改成Form 2.0
'Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text_LostFocus(Index As Integer)
   Select Case Index
          Case 5, 6, 7, 8, 9, 10, 11
               'edit by nickc 2007/06/11  切換輸入法改用API
               'Text(Index).IMEMode = 2
               CloseIme
   Case 2, 4
      'Add/Modify By Cheng 2002/09/09
      If blnClkSure = False Then
         If RunNick(Text(Index - 1), Text(Index)) Then
            Text(Index - 1).SetFocus
         End If
      Else
         blnClkSure = False
      End If
   End Select
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 1, 2
      If Text(Index) <> "" Then
         If Not CheckIsTaiwanDate(Text(Index)) Then Cancel = True
      End If
Case 3, 4
'    If Text(Index) <> "" Then
'        If ChgType(2, Text(Index)) Then
'
'        Else
'           Cancel = True
'        End If
'    End If
Case 11, 13
     If Text(Index) <> "" Then
        If CheckLengthIsOK(Text(Index), 80) = False Then
           Cancel = True
        End If
     End If
End Select
If Cancel Then TextInverse Text(Index)
End Sub

Private Sub txtcp01_GotFocus()
   TextInverse txtCP01
   CloseIme
End Sub

Private Sub txtcp01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtcp01_Validate(Cancel As Boolean)
  Dim strTit As String
  Dim strMsg As String
  
  txtCP01.Text = UCase(txtCP01.Text)
  If IsEmptyText(txtCP01) = False Then
      '2011/5/20 MODIFY BY SONIA
      'If Not IsCorrectSysKindLaw(txtcp01) Then
      If CheckSys(txtCP01) <> "3" And CheckSys(txtCP01) <> "4" Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "系統類別不正確"
         MsgBox strMsg, vbOKOnly, strTit
         txtcp01_GotFocus
         Exit Sub
      End If
      
'      ' 檢查使用者是否有使用該系統類別的權限
'      If IsUserHasRightOfSystem(strUserNum, txtcp01) = False Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "您沒有使有此系統別的權限"
'         MsgBox strMsg, vbOKOnly, strTit
'         txtcp01_GotFocus
'         Exit Sub
'      End If
   End If

'   If txtcp01 <> "" Then
'      txtcp01 = UCase(txtcp01)
'      If txtcp01 = "L" Or txtcp01 = "LA" Or txtcp01 = "FCL" Then
'         blnCom1 = True
'      Else
'         DataErrorMessage 1, "系統類別"
'         blnCom1 = False
'         Cancel = True
'      End If
'   End If
'   ChkCmd
   If Cancel Then TextInverse txtCP01
End Sub

Private Sub txtcp02_GotFocus()
   TextInverse txtCP02
   CloseIme
End Sub

Private Sub txtcp02_Validate(Cancel As Boolean)
   If txtCP02 <> "" Then
      blnCom2 = True
   End If
   If Cancel Then TextInverse txtCP02
End Sub

Private Sub txtcp03_GotFocus()
   TextInverse txtCP03
   CloseIme
End Sub

Private Sub txtcp03_Validate(Cancel As Boolean)
   If txtCP03 <> "" Then
      blnCom3 = True
   End If
   ChkCmd
End Sub

Private Sub ChkCmd()
   If txtCP03 = "" Then blnCom3 = True: blnCom4 = True
 '  If blnCom1 And blnCom2 And blnCom3 And blnCom4 Then cmdSure.Enabled = True
End Sub

Private Sub cboPerson_Click()
  Dim nPos As Integer
  Dim strPerson As String
  
   If cboPerson.Text <> "" Then
       nPos = InStr(cboPerson.Text, ",")
       If nPos <> 0 Then
          strPerson = Left(cboPerson.Text, nPos - 1)
          CheckStaff (strPerson)
       End If
   End If
    '   If Not ChgType(3, cboPerson) Then lbeName = ""
End Sub

Private Sub cboPerson_Validate(Cancel As Boolean)
   If cboPerson = "" Then lbeName = ""
End Sub

Private Function ChgType(i As Integer, strText As String) As Boolean
 Dim strTemp As String
   Select Case i
      Case 1
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetCaseProperty(txtcp01, StrText, strTemp, False) Then ChgType = True Else ChgType = False
         If ClsPDGetCaseProperty(txtCP01, strText, strTemp, False) Then ChgType = True Else ChgType = False
      Case 2
         'edit by nickc 2007/02/07 不用 dll 了
         'If objLawDll.GetGovName(StrText, strTemp) Then ChgType = True Else ChgType = False
         If ClsPDGetGovName(strText, strTemp) Then ChgType = True Else ChgType = False
      Case 3
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetStaff(StrText, strTemp) Then lbeName = strTemp: ChgType = True Else ChgType = False
         If ClsPDGetStaff(strText, strTemp) Then lbeName = strTemp: ChgType = True Else ChgType = False
   End Select
End Function

Private Function ReadSQL() As String
  Dim Stt(1 To 3) As String
  Dim i As Integer
  Dim strSql As String
  Dim strCP01 As String
  Dim strCP02 As String
  Dim strCP03 As String
  Dim strCP04 As String
  Dim nPos As Integer
  Dim strPerson As String
  
  nPos = 0

  strSql = ""
  strCP01 = txtCP01.Text
  strCP02 = txtCP02.Text
  If txtCP03.Text <> "" Then
     strCP03 = txtCP03.Text
  Else
     strCP03 = "0"
  End If
  If txtCP04.Text <> "" Then
     strCP04 = txtCP04.Text
  Else
     strCP04 = "00"
  End If
  
  '本所案號
  If strCP01 <> "" And strCP02 <> "" Then
     strSql = " AND CP01 ='" & strCP01 & "'" & _
              " AND CP02 ='" & strCP02 & "'" & _
              " AND CP03 ='" & strCP03 & "'" & _
              " AND CP04 ='" & strCP04 & "'"
  End If
   
  '法院案號
  If Text(0) <> "" Then
     strSql = strSql & " AND cp35='" + Text(0).Text + "'"
  End If
   
  '開庭人員
  If cboPerson.Text <> "" Then
     nPos = InStr(cboPerson.Text, ",")
     If nPos <> 0 Then
        strPerson = Left(cboPerson, nPos - 1)
     Else
        strPerson = cboPerson
     End If
     'Modify By Sindy 2011/6/17
     'strSql = strSql & " and cdp02='" + strPerson + "'"
     strSql = strSql & " and (cdp02='" + strPerson + "' or cl02='" + strPerson + "')"
  End If
   
  '開庭日期
  If Text(1) <> "" And Text(2) <> "" Then
     strSql = strSql & " and CDP03 BETWEEN '" + ChangeTStringToWString(Text(1).Text) + "' AND '" + ChangeTStringToWString(Text(2).Text) + "'"
  ElseIf Text(1).Text <> "" And Text(2).Text = "" Then
      'Modify By Cheng 2002/03/21
'      strSQL = strSQL & " AND CDP03 ='" & ChangeTStringToWString(Text(1).Text) & "'"
      strSql = strSql & " AND CDP03 >= '" & ChangeTStringToWString(Text(1).Text) & "' AND CDP03 <= '" & ChangeTStringToWString(ServerDate - 19110000) & "' "
  ElseIf Text(1) = "" And Text(2) <> "" Then
      strSql = strSql & " and cdp03 <= '" + ChangeTStringToWString(Text(2)) + "'"
  End If
   
  '機關代號
  'Modified by Lydia 2025/11/18 改成下拉選單
'  If Text(3) <> "" And Text(4) <> "" Then
'      strSql = strSql & " and (cp71 between '" + Text(3) + "' and '" + Text(4) + "')"
'  ElseIf Text(3) <> "" And Text(4) = "" Then
'      strSql = strSql & " AND CP71 ='" & Text(3).Text & "'"
'  ElseIf Text(3) = "" And Text(4) <> "" Then
'      strSql = strSql & " and cp71 <= '" + Text(4) + "'"
'  End If
  If Trim(Left(cboGov(0).Text, 3)) <> "" And Trim(Left(cboGov(1).Text, 3)) <> "" Then
      strSql = strSql & " and (cp71 between '" + Trim(Left(cboGov(0).Text, 3)) + "' and '" + Trim(Left(cboGov(1).Text, 3)) + "')"
  ElseIf Trim(Left(cboGov(0).Text, 3)) <> "" And Trim(Left(cboGov(1).Text, 3)) = "" Then
      strSql = strSql & " AND CP71 ='" & Trim(Left(cboGov(0).Text, 3)).Text & "'"
  ElseIf Trim(Left(cboGov(0).Text, 3)) = "" And Trim(Left(cboGov(1).Text, 3)) <> "" Then
      strSql = strSql & " and cp71 <= '" + Trim(Left(cboGov(1).Text, 3)) + "'"
  End If
  'end 2025/11/18
  
  '法官
   If Text(5).Text <> "" Then
      strSql = strSql & " and ((cdp08 like '" & Text(5).Text & ",%' OR" & _
                        " CDP08 LIKE '%," & Text(5).Text & ",%' OR " & _
                        " CDP08 LIKE '%," & Text(5).Text & "' OR " & _
                        " CDP08 ='" & Text(5).Text & "')"
      If Text(6).Text = "" And Text(7).Text = "" Then
         strSql = strSql & ")"
      End If
   End If
   
   If Text(6).Text <> "" Then
      If Text(5).Text <> "" Then
         strSql = strSql & " OR (cdp08 like '" & Text(6).Text & ",%' OR" & _
                           " CDP08 LIKE '%," & Text(6).Text & ",%' OR " & _
                           " CDP08 LIKE '%," & Text(6).Text & "' OR " & _
                           " CDP08 ='" & Text(6).Text & "')"
      Else
         strSql = strSql & " AND ((cdp08 like '" & Text(6).Text & ",%' OR" & _
                           " CDP08 LIKE '%," & Text(6).Text & ",%' OR " & _
                           " CDP08 LIKE '%," & Text(6).Text & "' OR " & _
                           " CDP08 ='" & Text(6).Text & "')"
      End If
      If Text(7).Text = "" Then
         strSql = strSql & ")"
      End If
   End If
   
   If Text(7).Text <> "" Then
      If Text(6).Text <> "" Or Text(5).Text <> "" Then
         strSql = strSql & " OR (cdp08 like '" & Text(7).Text & ",%' OR" & _
                           " CDP08 LIKE '%," & Text(7).Text & ",%' OR " & _
                           " CDP08 LIKE '%," & Text(7).Text & "' OR " & _
                           " CDP08 ='" & Text(7).Text & "')"
      ElseIf Text(6).Text = "" And Text(5).Text = "" Then
         strSql = strSql & " AND ((cdp08 like '" & Text(7).Text & ",%' OR" & _
                           " CDP08 LIKE '%," & Text(7).Text & ",%' OR " & _
                           " CDP08 LIKE '%," & Text(7).Text & "' OR " & _
                           " CDP08 ='" & Text(7).Text & "')"
      End If
      strSql = strSql & ")"
   End If

   '檢察官
   If Text(8).Text <> "" Then
      strSql = strSql & " and ((cdp09 like '" & Text(8).Text & ",%' OR" & _
                        " CDP09 LIKE '%," & Text(8).Text & ",%' OR " & _
                        " CDP09 LIKE '%," & Text(8).Text & "' OR " & _
                        " CDP09 ='" & Text(8).Text & "')"
      If Text(9).Text = "" And Text(10).Text = "" Then
         strSql = strSql & ")"
      End If
   End If
   
   If Text(9).Text <> "" Then
      If Text(8).Text <> "" Then
         strSql = strSql & " OR (cdp09 like '" & Text(9).Text & ",%' OR" & _
                           " CDP09 LIKE '%," & Text(9).Text & ",%' OR " & _
                           " CDP09 LIKE '%," & Text(9).Text & "' OR " & _
                           " CDP09 ='" & Text(9).Text & "')"
      Else
         strSql = strSql & " AND ((cdp09 like '" & Text(9).Text & ",%' OR" & _
                           " CDP09 LIKE '%," & Text(9).Text & ",%' OR " & _
                           " CDP09 LIKE '%," & Text(9).Text & "' OR " & _
                           " CDP09 ='" & Text(9).Text & "')"
      End If
      If Text(10).Text = "" Then
         strSql = strSql & ")"
      End If
   End If
   
   If Text(10).Text <> "" Then
      If Text(9).Text <> "" Or Text(8).Text <> "" Then
         strSql = strSql & " OR (cdp09 like '" & Text(10).Text & ",%' OR" & _
                           " CDP09 LIKE '%," & Text(10).Text & ",%' OR " & _
                           " CDP09 LIKE '%," & Text(10).Text & "' OR " & _
                           " CDP09 ='" & Text(10).Text & "')"
      ElseIf Text(9).Text = "" And Text(8).Text = "" Then
         strSql = strSql & " AND ((cdp09 like '" & Text(10).Text & ",%' OR" & _
                           " CDP09 LIKE '%," & Text(10).Text & ",%' OR " & _
                           " CDP09 LIKE '%," & Text(10).Text & "' OR " & _
                           " CDP09 ='" & Text(10).Text & "')"
      End If
      strSql = strSql & ")"
   End If
   
   '當事人(中)：
   If Text(11) <> "" Then
      strSql = strSql & " and cu04='" + Text(11) + "'"
   End If
   '當事人(英)
   If Text(12) <> "" Then
      strSql = strSql & " and cu05='" + Text(12) + "'"
   End If
   '當事人(日)：
   If Text(13) <> "" Then
      strSql = strSql & " and cu06='" + Text(13) + "'"
   End If

   ReadSQL = strSql
End Function

Private Sub txtcp04_GotFocus()
   TextInverse txtCP04
   CloseIme
End Sub

Private Sub txtcp04_Validate(Cancel As Boolean)
   If txtCP04 <> "" Then
      blnCom4 = True
   End If
   ChkCmd
   If Cancel Then TextInverse txtCP04
End Sub

Private Sub GetStaff()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
  
'modify by sonia 2019/2/14 改用共用Function GetLawerList
'  '2011/11/8 modify by sonia 加投資法務的律師,改排序條件
'  'strSql = "SELECT ST01,ST02 FROM STAFF WHERE ST03 ='L01' ORDER BY ST01"
'  strSql = "SELECT ST01,ST02 FROM STAFF WHERE (ST03 ='L01' OR ST20='13') ORDER BY ST04,ST03,ST51 DESC,ST01"
'  rsTmp.CursorLocation = adUseClient
'  rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'
'  If rsTmp.EOF = False Then
'     Do While rsTmp.EOF = False
'        If Not IsNull(rsTmp.Fields("ST01")) Then
'           cboPerson.AddItem rsTmp.Fields("ST01") & "," & IIf(IsNull(rsTmp.Fields("ST02")), "", rsTmp.Fields("ST02"))
'        End If
'        rsTmp.MoveNext
'     Loop
'  End If
'  rsTmp.Close
'  Set rsTmp = Nothing
Dim i As Integer, varTmp1 As Variant, strTmp As String

   strSql = GetLawerList()
   varTmp1 = Split(strSql, ";")
   For i = 0 To UBound(varTmp1)
      strTmp = varTmp1(i)
      cboPerson.AddItem strTmp
   Next
'end 2019/2/14

   'Added by Lydia 2025/11/18
   For i = 0 To 1
      Call PUB_SetGovCmb(Me.cboGov(i), m_GovListNew(i))
      m_GovListDef(i) = m_GovListNew(i)
   Next i
   'end 2025/11/18
End Sub

Private Sub CheckStaff(ByVal strNum As String)
  Dim rsTmp As New ADODB.Recordset
  Dim strSql As String
  
  strSql = "SELECT ST02 FROM STAFF WHERE ST01 ='" & strNum & "'"
  rsTmp.CursorLocation = adUseClient
  rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
  
  If rsTmp.EOF = False Then
        If Not IsNull(rsTmp.Fields("ST02")) Then
           lbeName.Caption = rsTmp.Fields("ST02")
        Else
           lbeName.Caption = ""
        End If
  Else
      lbeName.Caption = ""
  End If
  rsTmp.Close
  Set rsTmp = Nothing
End Sub

'Added by Lydia 2025/11/18
Private Sub cboGov_DropButtonClick(Index As Integer)
   If cboGov(Index).Text <> "" Then
      If Val(Trim(Left(cboGov(Index).Text, 3))) > 0 Then
      Else  '依輸入文字模糊比對
         Call PUB_SetGovCmb(cboGov(Index), m_GovListNew(Index), , Trim(cboGov(Index).Text))
         If m_GovListNew(Index) = "" Then
            Call PUB_SetGovCmb(cboGov(Index), m_GovListNew(Index))
         End If
      End If
   Else
      If m_GovListNew(Index) <> m_GovListDef(Index) Then
         Call PUB_SetGovCmb(cboGov(Index), m_GovListNew(Index))
      End If
   End If
End Sub

'Added by Lydia 2025/11/18
Private Sub cboGov_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Added by Lydia 2025/11/18
Private Sub cboGov_Validate(Index As Integer, Cancel As Boolean)
   If Trim(cboGov(Index).Text) <> "" And cboGov(Index).Tag <> cboGov(Index).Text Then
      If PUB_ChkGovIsExist(IIf(Val(Trim(Left(cboGov(Index).Text, 3))) > 0, Trim(Left(cboGov(Index).Text, 3)), Trim(cboGov(Index).Text)), strExc(3), strExc(4)) = True Then
         cboGov(Index).Text = strExc(3) & " " & strExc(4)
      Else
         Cancel = True
         cboGov(Index).SetFocus
         cboGov_GotFocus Index
         Exit Sub
      End If
   End If
   cboGov(Index).Tag = cboGov(Index).Text
End Sub

'Added by Lydia 2025/11/18
Private Sub cboGov_GotFocus(Index As Integer)
   TextInverse cboGov(Index)
End Sub

