VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140104 
   BorderStyle     =   1  '單線固定
   Caption         =   "分所銷卷維護"
   ClientHeight    =   3200
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   9120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3200
   ScaleWidth      =   9120
   Begin VB.CommandButton cmdok 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   5760
      TabIndex        =   6
      Top             =   30
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "存檔(&S)"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   6765
      TabIndex        =   7
      Top             =   30
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   2
      Left            =   7770
      TabIndex        =   8
      Top             =   30
      Width           =   975
   End
   Begin MSForms.TextBox txt1 
      Height          =   285
      Index           =   5
      Left            =   1020
      TabIndex        =   5
      Top             =   2880
      Width           =   8025
      VariousPropertyBits=   679493659
      MaxLength       =   40
      Size            =   "14155;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   285
      Index           =   3
      Left            =   2700
      TabIndex        =   4
      Top             =   840
      Width           =   375
      VariousPropertyBits=   679493659
      MaxLength       =   2
      Size            =   "661;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   285
      Index           =   2
      Left            =   2430
      TabIndex        =   3
      Top             =   840
      Width           =   255
      VariousPropertyBits=   679493659
      MaxLength       =   1
      Size            =   "450;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   285
      Index           =   1
      Left            =   1590
      TabIndex        =   2
      Top             =   840
      Width           =   855
      VariousPropertyBits=   679493659
      MaxLength       =   6
      Size            =   "1508;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   495
      VariousPropertyBits=   679493659
      MaxLength       =   3
      Size            =   "873;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   285
      Index           =   4
      Left            =   1080
      TabIndex        =   0
      Top             =   540
      Width           =   1395
      VariousPropertyBits=   679493659
      MaxLength       =   12
      Size            =   "2461;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtC 
      Height          =   555
      Left            =   1590
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1170
      Width           =   7485
      VariousPropertyBits=   -1400881131
      BackColor       =   -2147483633
      ScrollBars      =   2
      Size            =   "8555;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblD2 
      Height          =   195
      Left            =   5400
      TabIndex        =   32
      Top             =   2610
      Width           =   1905
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   180
      Left            =   4410
      TabIndex        =   31
      Top             =   2625
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   60
      TabIndex        =   30
      Top             =   900
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "分所案號："
      Height          =   180
      Left            =   60
      TabIndex        =   29
      Top             =   570
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Left            =   60
      TabIndex        =   28
      Top             =   2010
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "銷卷備註："
      Height          =   180
      Left            =   30
      TabIndex        =   27
      Top             =   2910
      Width           =   900
   End
   Begin VB.Label lblJ 
      Height          =   180
      Left            =   1605
      TabIndex        =   26
      Top             =   1740
      Width           =   4665
   End
   Begin VB.Label Label13 
      Caption         =   "案件名稱（中）："
      Height          =   180
      Left            =   60
      TabIndex        =   25
      Top             =   1185
      Width           =   1455
   End
   Begin VB.Label Label18 
      Caption         =   "案件名稱（日）："
      Height          =   180
      Left            =   60
      TabIndex        =   24
      Top             =   1740
      Width           =   1455
   End
   Begin VB.Label Label19 
      Caption         =   "案件名稱（英）："
      Height          =   180
      Left            =   60
      TabIndex        =   23
      Top             =   1470
      Width           =   1455
   End
   Begin VB.Label lblE 
      Height          =   180
      Left            =   1605
      TabIndex        =   22
      Top             =   1455
      Width           =   7455
   End
   Begin VB.Label lblC 
      Height          =   180
      Left            =   1605
      TabIndex        =   21
      Top             =   1185
      Width           =   7425
   End
   Begin VB.Label Label22 
      Caption         =   "案件名稱（中）："
      Height          =   180
      Left            =   60
      TabIndex        =   20
      Top             =   1185
      Width           =   1455
   End
   Begin MSForms.Label lblCU 
      Height          =   195
      Left            =   1020
      TabIndex        =   19
      Top             =   2010
      Width           =   7965
      VariousPropertyBits=   27
      Size            =   "14049;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblST 
      Height          =   195
      Left            =   5400
      TabIndex        =   18
      Top             =   2370
      Width           =   1905
      VariousPropertyBits=   27
      Size            =   "3360;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   4440
      TabIndex        =   17
      Top             =   2370
      Width           =   900
   End
   Begin VB.Label lblLastCP10 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1050
      TabIndex        =   16
      Top             =   2370
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "最後程序："
      Height          =   180
      Left            =   60
      TabIndex        =   15
      Top             =   2370
      Width           =   900
   End
   Begin VB.Label lblD1 
      Height          =   195
      Left            =   1020
      TabIndex        =   14
      Top             =   2610
      Width           =   1905
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Left            =   30
      TabIndex        =   13
      Top             =   2625
      Width           =   900
   End
   Begin VB.Label lblst06 
      Height          =   255
      Left            =   780
      TabIndex        =   11
      Top             =   300
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(2:中;3:南;4:高)"
      Height          =   180
      Left            =   1230
      TabIndex        =   10
      Top             =   330
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "所別："
      Height          =   180
      Left            =   90
      TabIndex        =   9
      Top             =   330
      Width           =   540
   End
End
Attribute VB_Name = "frm140104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/03 Form2.0 txtC/lblCU/lblST/txt1(5)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
Option Explicit


Private Sub cmdOK_Click(Index As Integer)
Dim RetCnt1 As Integer
Dim RetCnt2 As Integer
Dim RetCnt3 As Integer
Dim RetCnt4 As Integer
Dim RetCnt5 As Integer
Dim strTit As String
Dim strMsg As String
Dim nResponse

Select Case Index
Case 0
     If Trim(txt1(0)) = "" And Trim(txt1(4)) = "" Then
        MsgBox "條件最少要有一個要輸入！", vbExclamation
        txt1(4).SetFocus
        Exit Sub
     End If
     If Trim(txt1(0)) <> "" And Trim(txt1(1)) = "" And Trim(txt1(4)) = "" Then
        MsgBox "條件要明確！", vbExclamation
        txt1(0).SetFocus
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     Me.Enabled = False
     StrMenu
     Me.Enabled = True
     Screen.MousePointer = vbDefault
Case 1
    'Add by Amy 2021/12/03 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me) = False Then
        Exit Sub
    End If
    
     StrMenu
     'Add By Sindy 2012/1/4
     strDate = DBDATE(DateAdd("yyyy", -1, ChangeWStringToWDateString(strSrvDate(1))))
     strSql = "SELECT count(*) FROM CaseProgress WHERE cp01='" & txt1(0) & "' and cp02='" & txt1(1) & "' and cp03='" & txt1(2) & "' and cp04='" & txt1(3) & "' and cp05>" & strDate & " and substr(cp09,1,1)='A' "
     intI = 1
     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
     If intI = 1 Then
         If RsTemp.Fields(0) > 0 Then
            strTit = "詢問"
            strMsg = "此案(" & txt1(0) & "-" & txt1(1) & "-" & txt1(2) & "-" & txt1(3) & ")一年內仍有收文，是否確定銷卷？"
            nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
            If nResponse = vbNo Then Exit Sub
         End If
     End If
     '2012/1/4 End
     'Add By Sindy 2015/3/9 檢查財務的規費資料是否不平
     '2022/5/12 MODIFY BY SONIA (CFP-018656)TF母案領土延伸分開算,CFP僅EPC子案與母案一起算,接續案或集體設計個別算
     'strSql = "select sum(ax206),sum(ax207) from acc021 where ax214='" & txt1(0) & txt1(1) & txt1(2) & txt1(3) & "' and substr(ax205,1,4)='2201'"
     If txt1(0) = "TF" Then
         strSql = "select sum(ax206),sum(ax207) from acc021 where ax214>='" & txt1(0) & txt1(1) & "000' and ax214<='" & txt1(0) & txt1(1) & "999' and substr(ax205,1,4)='2201'"
     Else
         strSql = "select sum(ax206),sum(ax207) from acc021 where ax214>='" & txt1(0) & txt1(1) & txt1(2) & "00' and ax214<='" & txt1(0) & txt1(1) & txt1(2) & "99' and substr(ax205,1,4)='2201'"
     End If
     'end 2022/5/12
     intI = 1
     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
     If intI = 1 Then
         If Val("" & RsTemp.Fields(0)) <> Val("" & RsTemp.Fields(1)) Then
            MsgBox "財務的規費資料不平, 尚不可銷卷！", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
     End If
     '2015/3/9 END
     Screen.MousePointer = vbHourglass
     If cmdOK(1).Enabled = True Then
         cnnConnection.Execute "update patent set pa136=to_number(to_char(sysdate,'YYYYMMDD')),pa137='" & strUserNum & "',pa138='" & ChgSQL(txt1(5)) & "' where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " pa01='" & txt1(0) & "' and pa02='" & txt1(1) & "' and pa03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and pa04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' " & IIf(txt1(4) = "", " and pa47 is null ", " and pa47='" & txt1(4) & "' "), "pa47='" & txt1(4) & "' "), RetCnt1
         cnnConnection.Execute "update trademark set tm73=to_number(to_char(sysdate,'YYYYMMDD')),tm74='" & strUserNum & "',tm75='" & ChgSQL(txt1(5)) & "' where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " tm01='" & txt1(0) & "' and tm02='" & txt1(1) & "' and tm03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and tm04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' " & IIf(txt1(4) = "", " and tm34 is null ", " and tm34='" & txt1(4) & "' "), " tm34='" & txt1(4) & "' "), RetCnt2
         cnnConnection.Execute "update servicepractice set sp68=to_number(to_char(sysdate,'YYYYMMDD')),sp69='" & strUserNum & "',sp70='" & ChgSQL(txt1(5)) & "' where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " sp01='" & txt1(0) & "' and sp02='" & txt1(1) & "' and sp03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and sp04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' " & IIf(txt1(4) = "", " and sp28 is null ", " and sp28='" & txt1(4) & "' "), " sp28='" & txt1(4) & "' "), RetCnt3
         cnnConnection.Execute "update lawcase set lc36=to_number(to_char(sysdate,'YYYYMMDD')),lc37='" & strUserNum & "',lc38='" & ChgSQL(txt1(5)) & "' where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " lc01='" & txt1(0) & "' and lc02='" & txt1(1) & "' and lc03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and lc04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' " & IIf(txt1(4) = "", " and lc16 is null ", " and lc16='" & txt1(4) & "' "), " lc16='" & txt1(4) & "' "), RetCnt4
         cnnConnection.Execute "update hirecase set hc20=to_number(to_char(sysdate,'YYYYMMDD')),hc21='" & strUserNum & "',hc22='" & ChgSQL(txt1(5)) & "' where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " hc01='" & txt1(0) & "' and hc02='" & txt1(1) & "' and hc03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and hc04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' " & IIf(txt1(4) = "", " and hc07 is null ", " and hc07='" & txt1(4) & "' "), " hc07='" & txt1(4) & "' "), RetCnt5
     Else
         MsgBox "此本所案號不可銷卷！", vbExclamation
         Screen.MousePointer = vbDefault
         Exit Sub
     End If
     If RetCnt1 + RetCnt2 + RetCnt3 + RetCnt4 + RetCnt5 <> 0 Then
        MsgBox "銷卷成功！", vbInformation
        txt1(0) = ""
        txt1(1) = ""
        txt1(2) = ""
        txt1(3) = ""
        txt1(4) = ""
        txt1(5) = ""
        cmdOK(1).Enabled = False
        txtC = ""
        lblC = "'"
        lblE = ""
        lblJ = ""
        lblCU = ""
        lblLastCP10 = ""
        lblD1 = ""
        lblD2 = ""
        lblST = ""
        txt1(4).SetFocus
     Else
        MsgBox "銷卷失敗！", vbInformation
     End If
     Screen.MousePointer = vbDefault
Case 2
     Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    lblst06 = pub_strUserOffice
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm140104 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
End Sub

'Modify by Amy 2021/12/03 原:Integer
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 0
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If txt1(Index) = "" Then Exit Sub
         strExc(0) = "SELECT SK02 FROM SYSTEMKIND WHERE SK01='" & txt1(Index) & "'"
         intI = 1
         'edit by nickc 2007/02/08 不用 dll 了
         'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI <> 1 Then
            MsgBox "無此系統類別，請重新輸入 !", vbCritical
            Cancel = True
         End If
      Case 5
        Cancel = False
        If CheckLengthIsOK(txt1(Index).Text, txt1(Index).MaxLength) = False Then
            txt1(Index).SetFocus
            txt1_GotFocus Index
            Cancel = True
        End If
   End Select
   If Cancel = True Then TextInverse txt1(Index)
End Sub

Sub StrMenu()
Dim Tmprss As New ADODB.Recordset
Dim strMsg As String 'Add By Sindy 2014/9/4
If Screen.MousePointer = vbHourglass Then
    txtC = ""
    lblC = "'"
    lblE = ""
    lblJ = ""
    lblCU = ""
    txt1(5) = ""
    lblD1 = ""
    lblD2 = ""
    lblST = ""
    lblLastCP10 = ""
End If
'Modify By Sindy 2014/9/4 +案件備註是否有不銷卷字樣者
'+,pa91
'+,tm58
'+,lc27
'+,hc12
'+,sp18
'Select Case txt1(0)
'Case "CFP", "FCP", "P"   '專利
      strSql = "select PA05,PA06,PA07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),pa108,pa136,pa137,pa138,pa01,pa02,pa03,pa04,pa47,st06,pa91 " & _
               " From patent,CUSTOMER,staff " & _
               " Where " & IIf(Trim(txt1(4)) = "", "Pa01='" & txt1(0) & "' and Pa02='" & txt1(1) & "' and Pa03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and Pa04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' ", " pa47='" & txt1(4) & "' ") & " and SUBSTR(pa26,1,8)=CU01(+) AND SUBSTR(pa26,9,1)=CU02(+) and cu13=st01(+)  "
'Case "CFT", "FCT", "T", "TF"   '商標
      strSql = strSql & " union Select TM05,TM06,TM07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),tm57,tm73,tm74,tm75,tm01,tm02,tm03,tm04,tm34,st06,tm58 " & _
               " From TradeMark,CUSTOMER,staff " & _
               " Where " & IIf(Trim(txt1(4)) = "", "tm01='" & txt1(0) & "' and tm02='" & txt1(1) & "' and tm03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and tm04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' ", " tm34='" & txt1(4) & "' ") & " and SUBSTR(tm23,1,8)=CU01(+) AND SUBSTR(tm23,9,1)=CU02(+) and cu13=st01(+)  "
'Case "CFL", "FCL", "L"          '法務
      strSql = strSql & " union select LC05,LC06,LC07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),lc34,lc36,lc37,lc38,lc01,lc02,lc03,lc04,lc16,st06,lc27 " & _
               " From lawcase,CUSTOMER,staff " & _
               " Where " & IIf(Trim(txt1(4)) = "", "lc01='" & txt1(0) & "' and lc02='" & txt1(1) & "' and lc03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and lc04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' ", " lc16='" & txt1(4) & "' ") & " and SUBSTR(lc11,1,8)=CU01(+) AND SUBSTR(lc11,9,1)=CU02(+) and cu13=st01(+)  "
'Case "LA"                      '顧問
      strSql = strSql & " union select HC06,'','',NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),hc19,hc20,hc21,hc22,hc01,hc02,hc03,hc04,hc07,st06,hc12 " & _
               " From hirecase,CUSTOMER,staff " & _
               " Where " & IIf(Trim(txt1(4)) = "", "hc01='" & txt1(0) & "' and hc02='" & txt1(1) & "' and hc03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and hc04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' ", " hc07='" & txt1(4) & "' ") & " and SUBSTR(hc05,1,8)=CU01(+) AND SUBSTR(hc05,9,1)=CU02(+) and cu13=st01(+)  "
'Case Else                  '服務
      strSql = strSql & " union select SP05,SP06,SP07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),sp61,sp68,sp69,sp70,sp01,sp02,sp03,sp04,sp28,st06,sp18 " & _
               " From servicepractice,CUSTOMER,staff " & _
               " Where " & IIf(Trim(txt1(4)) = "", "sp01='" & txt1(0) & "' and sp02='" & txt1(1) & "' and sp03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and sp04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' ", " sp28='" & txt1(4) & "' ") & " and SUBSTR(sp08,1,8)=CU01(+) AND SUBSTR(sp08,9,1)=CU02(+) and cu13=st01(+)  "
'End Select
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 Then
    'add by nickc 2007/08/15  先檢查該所有的資料
    adoRecordset.MoveFirst
    Do While Not adoRecordset.EOF
        If pub_strUserOffice = CheckStr(adoRecordset.Fields("st06").Value) Then
            Exit Do
        End If
        adoRecordset.MoveNext
    Loop
    If adoRecordset.EOF Then
        adoRecordset.MoveFirst
    End If
    lblCU = "" & adoRecordset.Fields(3).Value
    txt1(0) = "" & adoRecordset.Fields(8).Value
    txt1(1) = "" & adoRecordset.Fields(9).Value
    txt1(2) = "" & adoRecordset.Fields(10).Value
    txt1(3) = "" & adoRecordset.Fields(11).Value
    txt1(4) = "" & adoRecordset.Fields(12).Value
    If Screen.MousePointer = vbHourglass Then
        txt1(5) = "" & adoRecordset.Fields(7).Value
    End If
    Select Case txt1(0)
    Case "T", "FCT", "CFT", "TF", "TS", "S"
        txtC.Text = "" & adoRecordset.Fields(0).Value
        txtC.Visible = True
        lblC.Visible = False
        lblE.Visible = False
        lblJ.Visible = False
    Case Else
        txtC.Visible = False
        lblC.Visible = True
        lblE.Visible = True
        lblJ.Visible = True
        If IsNull(adoRecordset.Fields(0)) Then
            lblC.Caption = ""
        Else
            lblC.Caption = Replace(adoRecordset.Fields(0), "&", "&&")
        End If
        If IsNull(adoRecordset.Fields(1)) Then
            lblE.Caption = ""
        Else
            lblE.Caption = Replace(adoRecordset.Fields(1), "&", "&&")
        End If
        If IsNull(adoRecordset.Fields(2)) Then
            lblJ.Caption = ""
        Else
            lblJ.Caption = Replace(adoRecordset.Fields(2), "&", "&&")
        End If
    End Select
    If CheckStr(adoRecordset.Fields(5).Value) = "" Then
        If pub_strUserOffice = CheckStr(adoRecordset.Fields("st06").Value) Then
            cmdOK(1).Enabled = True
        Else
            cmdOK(1).Enabled = False
            MsgBox "此案與你的所別不同！", vbInformation
        End If
    Else
        cmdOK(1).Enabled = False
        MsgBox "此案已銷卷！", vbInformation
    End If
    'Add By Sindy 2014/9/4 若該案的案件備註有'不銷卷'字樣者,或是該案號有應收帳款(ACC1K0及ACC0K0)未收者,要顯示訊息,不可進行銷卷
    strMsg = ""
    If InStr("" & adoRecordset.Fields(14).Value, "不銷卷") > 0 Then
      strMsg = "該案號為不可銷卷"
    End If
    If PUB_ChkReceivables(txt1(0) & "-" & txt1(1) & "-" & txt1(2) & "-" & txt1(3)) = True Then
      If strMsg <> "" Then strMsg = strMsg & " 且 "
      strMsg = strMsg & "該案號有未收款"
    End If
    If strMsg <> "" Then
      cmdOK(1).Enabled = False
      MsgBox strMsg & "，不可進行銷卷作業", vbInformation
      Exit Sub
    End If
    '2014/9/4 END
    If Screen.MousePointer = vbHourglass Then
        '帶最後程序資料
        '2010/3/23  MODIFY BY SONIA 剔除下一程序非智權人員掌控之案件性質改以strNpSqlOfNoSalesDuty控制
        'strSql = "select * from nextprogress,staff,casepropertymap where np02='" & txt1(0) & "' and np03='" & txt1(1) & "' and np04='" & txt1(2) & "' and np05='" & txt1(3) & "' and (np06='N' or np06 is null) and np10=st01(+) and np02=cpm01(+) and to_char(np07)=CPM02(+) and np08=(select max(np08) from nextprogress where np02='" & txt1(0) & "' and np03='" & txt1(1) & "' and np04='" & txt1(2) & "' and np05='" & txt1(3) & "' and (np06='N' or np06 is null) and np08<to_number(to_char(sysdate,'YYYYMMDD')) and not (np02 in ('L','FCL','CFL','LA') and np07='6001') and not (np02 in ('P','PS','CFP','CPS','FCP','FG') and np07 IN ('997','998','999','411','1204','1503'))  and not (np02 NOT in ('L','FCL','CFL','LA','P','PS','CFP','CPS','FCP','FG' ) and np07 IN ('997','998','999','305','1403')) " & _
                 "   ) and not exists (select np01 from nextprogress where np02='" & txt1(0) & "' and np03='" & txt1(1) & "' and np04='" & txt1(2) & "' and np05='" & txt1(3) & "' and np06='Y' and np08>(select max(np08) from nextprogress where np02='" & txt1(0) & "' and np03='" & txt1(1) & "' and np04='" & txt1(2) & "' and np05='" & txt1(3) & "' and (np06='N' or np06 is null) and np08<to_number(to_char(sysdate,'YYYYMMDD')) and not (np02 in ('L','FCL','CFL','LA') and np07='6001') and not (np02 in ('P','PS','CFP','CPS','FCP','FG') and np07 IN ('997','998','999','411','1204','1503'))  and not (np02 NOT in ('L','FCL','CFL','LA','P','PS','CFP','CPS','FCP','FG' ) and np07 IN ('997','998','999','305','1403')) ) and not (np02 in ('L','FCL','CFL','LA') and np07='6001') and not (np02 in ('P','PS','CFP','CPS','FCP','FG') and np07 IN ('997','998','999','411','1204','1503'))  and not (np02 NOT in ('L','FCL','CFL','LA','P','PS','CFP','CPS','FCP','FG' ) and np07 IN ('997','998','999','305','1403')) )"
        strSql = "select * from nextprogress,staff,casepropertymap where np02='" & txt1(0) & "' and np03='" & txt1(1) & "' and np04='" & txt1(2) & "' and np05='" & txt1(3) & "' and (np06='N' or np06 is null) and np10=st01(+) and np02=cpm01(+) and to_char(np07)=CPM02(+) and np08=(select max(np08) from nextprogress where np02='" & txt1(0) & "' and np03='" & txt1(1) & "' and np04='" & txt1(2) & "' and np05='" & txt1(3) & "' and (np06='N' or np06 is null) and np08<to_number(to_char(sysdate,'YYYYMMDD')) " & strNpSqlOfNoSalesDuty & _
                 "   ) and not exists (select np01 from nextprogress where np02='" & txt1(0) & "' and np03='" & txt1(1) & "' and np04='" & txt1(2) & "' and np05='" & txt1(3) & "' and np06='Y' and np08>(select max(np08) from nextprogress where np02='" & txt1(0) & "' and np03='" & txt1(1) & "' and np04='" & txt1(2) & "' and np05='" & txt1(3) & "' and (np06='N' or np06 is null) and np08<to_number(to_char(sysdate,'YYYYMMDD')) " & strNpSqlOfNoSalesDuty & " ) " & strNpSqlOfNoSalesDuty & " )"
        CheckOC
        With adoRecordset
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If .EOF And .BOF Then
                '2010/3/23  MODIFY BY SONIA 剔除下一程序非智權人員掌控之案件性質改以strNpSqlOfNoSalesDuty控制
                'strSql = "select * from nextprogress,staff,casepropertymap where np02='" & txt1(0) & "' and np03='" & txt1(1) & "' and np04='" & txt1(2) & "' and np05='" & txt1(3) & "' and (np06='N' or np06 is null) and np10=st01(+) and np02=cpm01(+) and to_char(np07)=CPM02(+) and np08=(select min(np08) from nextprogress where np02='" & txt1(0) & "' and np03='" & txt1(1) & "' and np04='" & txt1(2) & "' and np05='" & txt1(3) & "' and (np06='N' or np06 is null) and np08>to_number(to_char(sysdate,'YYYYMMDD')) and not (np02 in ('L','FCL','CFL','LA') and np07='6001') and not (np02 in ('P','PS','CFP','CPS','FCP','FG') and np07 IN ('997','998','999','411','1204','1503'))  and not (np02 NOT in ('L','FCL','CFL','LA','P','PS','CFP','CPS','FCP','FG' ) and np07 IN ('997','998','999','305','1403')) ) " & _
                         " and not (np02 in ('L','FCL','CFL','LA') and np07='6001') and not (np02 in ('P','PS','CFP','CPS','FCP','FG') and np07 IN ('997','998','999','411','1204','1503'))  and not (np02 NOT in ('L','FCL','CFL','LA','P','PS','CFP','CPS','FCP','FG' ) and np07 IN ('997','998','999','305','1403')) "
                strSql = "select * from nextprogress,staff,casepropertymap where np02='" & txt1(0) & "' and np03='" & txt1(1) & "' and np04='" & txt1(2) & "' and np05='" & txt1(3) & "' and (np06='N' or np06 is null) and np10=st01(+) and np02=cpm01(+) and to_char(np07)=CPM02(+) and np08=(select min(np08) from nextprogress where np02='" & txt1(0) & "' and np03='" & txt1(1) & "' and np04='" & txt1(2) & "' and np05='" & txt1(3) & "' and (np06='N' or np06 is null) and np08>to_number(to_char(sysdate,'YYYYMMDD')) " & strNpSqlOfNoSalesDuty & " ) " & strNpSqlOfNoSalesDuty
                CheckOC
                .CursorLocation = adUseClient
                .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If .EOF And .BOF Then
                    strSql = "select * from caseprogress ,staff,casepropertymap where cp01='" & txt1(0) & "' and cp02='" & txt1(1) & "' and cp03='" & txt1(2) & "' and cp04='" & txt1(3) & "' and cp05=(select max(cp05) from caseprogress where cp01='" & txt1(0) & "' and cp02='" & txt1(1) & "' and cp03='" & txt1(2) & "' and cp04='" & txt1(3) & "' ) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) "
                    CheckOC
                    .CursorLocation = adUseClient
                    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                    If .EOF And .BOF Then
                        MsgBox "查無案件程序資料！" & vbCrLf & "請檢查一下案件資料！", vbExclamation
                        cmdOK(1).Enabled = False
                        Exit Sub
                    Else
                        lblD1 = ChangeWStringToTDateString("" & .Fields("cp06").Value)
                        lblD2 = ChangeWStringToTDateString("" & .Fields("cp07").Value)
                        lblST = CheckStr(.Fields("st02"))
                        lblLastCP10 = IIf(GetPrjNation1(txt1(0) & "-" & txt1(1) & "-" & txt1(2) & "-" & txt1(3)) = "000", CheckStr(.Fields("cpm03")), CheckStr(.Fields("cpm04"))) & PUB_GetRelateCasePropertyName(CheckStr(.Fields("cp09")), "1")
                        If txt1(5) = "" Then
                            txt1(5) = lblLastCP10
                        End If
                        Exit Sub
                    End If
                End If
            Else
                If "" & .Fields("np09").Value <> "" Then
                    'add by nickc 2006/08/25 在檢查有無大收文日於np 的法定的程序
                    strSql = "select * from caseprogress ,staff,casepropertymap where cp01='" & txt1(0) & "' and cp02='" & txt1(1) & "' and cp03='" & txt1(2) & "' and cp04='" & txt1(3) & "' and cp05=(select max(cp05) from caseprogress where cp05>" & "" & .Fields("np09").Value & " and cp01='" & txt1(0) & "' and cp02='" & txt1(1) & "' and cp03='" & txt1(2) & "' and cp04='" & txt1(3) & "' and not ((cp01 in ('CPS','CFP','FG','FCP','PS','P') and cp10 in ('907','925','903')) or (cp01 in ('L','LA','FCL','CFL') and cp10 in ('991','992','993')) or (cp01 in ('CFT','CFC','FCT','S','T','TB','TC','TD','TF','TM','TR','TS','TT') and cp10 in ('703','718','704')))) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) "
                    Set Tmprss = New ADODB.Recordset
                    If Tmprss.State = 1 Then Tmprss.Close
                    Tmprss.CursorLocation = adUseClient
                    Tmprss.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                    If Tmprss.EOF And Tmprss.BOF Then
                    Else
                        lblD1 = ChangeWStringToTDateString("" & Tmprss.Fields("cp06").Value)
                        lblD2 = ChangeWStringToTDateString("" & Tmprss.Fields("cp07").Value)
                        lblST = CheckStr(Tmprss.Fields("st02"))
                        lblLastCP10 = IIf(GetPrjNation1(txt1(0) & "-" & txt1(1) & "-" & txt1(2) & "-" & txt1(3)) = "000", CheckStr(Tmprss.Fields("cpm03")), CheckStr(Tmprss.Fields("cpm04"))) & PUB_GetRelateCasePropertyName(CheckStr(Tmprss.Fields("cp09")), "1")
                        If txt1(5) = "" Then
                            txt1(5) = lblLastCP10
                        End If
                        Exit Sub
                    End If
                End If
            End If
            lblD1 = ChangeWStringToTDateString("" & .Fields("np08").Value)
            lblD2 = ChangeWStringToTDateString("" & .Fields("np09").Value)
            lblST = CheckStr(.Fields("st02"))
            lblLastCP10 = IIf(GetPrjNation1(txt1(0) & "-" & txt1(1) & "-" & txt1(2) & "-" & txt1(3)) = "000", CheckStr(.Fields("cpm03")), CheckStr(.Fields("cpm04")))
            If txt1(5) = "" Then
                txt1(5) = lblLastCP10 & "不續辦"
            End If
        End With
    End If
Else
    MsgBox "查無此案號資料！", vbExclamation
End If
End Sub
