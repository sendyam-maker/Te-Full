VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140107 
   BorderStyle     =   1  '單線固定
   Caption         =   "銷卷後復原維護"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   9075
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   1050
      MaxLength       =   12
      TabIndex        =   4
      Top             =   690
      Width           =   1395
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1050
      MaxLength       =   3
      TabIndex        =   0
      Top             =   420
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1530
      MaxLength       =   6
      TabIndex        =   1
      Top             =   420
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2370
      MaxLength       =   1
      TabIndex        =   2
      Top             =   420
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2610
      MaxLength       =   2
      TabIndex        =   3
      Top             =   420
      Width           =   375
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   2
      Left            =   7800
      TabIndex        =   7
      Top             =   30
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "復原(&S)"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   6795
      TabIndex        =   6
      Top             =   30
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   5790
      TabIndex        =   5
      Top             =   30
      Width           =   975
   End
   Begin MSForms.TextBox txtC 
      Height          =   555
      Left            =   1560
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   990
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
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "分所銷卷員："
      Height          =   180
      Left            =   30
      TabIndex        =   27
      Top             =   2745
      Width           =   1080
   End
   Begin MSForms.Label lblST 
      Height          =   195
      Left            =   1410
      TabIndex        =   26
      Top             =   2750
      Width           =   7545
      VariousPropertyBits=   27
      Size            =   "14049;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "北所銷卷日："
      Height          =   180
      Left            =   30
      TabIndex        =   25
      Top             =   2190
      Width           =   1080
   End
   Begin VB.Label lblD1 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1410
      TabIndex        =   24
      Top             =   2190
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "分所銷卷日："
      Height          =   180
      Left            =   30
      TabIndex        =   23
      Top             =   2475
      Width           =   1080
   End
   Begin VB.Label lblD2 
      Height          =   195
      Left            =   1410
      TabIndex        =   22
      Top             =   2470
      Width           =   7545
   End
   Begin MSForms.Label lblMemo 
      Height          =   195
      Left            =   1410
      TabIndex        =   21
      Top             =   3030
      Width           =   7545
      VariousPropertyBits=   27
      Size            =   "14049;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCU 
      Height          =   195
      Left            =   990
      TabIndex        =   20
      Top             =   1860
      Width           =   7965
      VariousPropertyBits=   27
      Size            =   "14049;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label22 
      Caption         =   "案件名稱（中）："
      Height          =   180
      Left            =   30
      TabIndex        =   19
      Top             =   1005
      Width           =   1455
   End
   Begin VB.Label lblC 
      Height          =   180
      Left            =   1575
      TabIndex        =   18
      Top             =   1005
      Width           =   7425
   End
   Begin VB.Label lblE 
      Height          =   180
      Left            =   1575
      TabIndex        =   17
      Top             =   1275
      Width           =   7455
   End
   Begin VB.Label Label19 
      Caption         =   "案件名稱（英）："
      Height          =   180
      Left            =   30
      TabIndex        =   16
      Top             =   1290
      Width           =   1455
   End
   Begin VB.Label Label18 
      Caption         =   "案件名稱（日）："
      Height          =   180
      Left            =   30
      TabIndex        =   15
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "案件名稱："
      Height          =   180
      Left            =   30
      TabIndex        =   14
      Top             =   1005
      Width           =   1455
   End
   Begin VB.Label lblJ 
      Height          =   180
      Left            =   1575
      TabIndex        =   13
      Top             =   1560
      Width           =   4665
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "分所銷卷備註："
      Height          =   180
      Left            =   30
      TabIndex        =   11
      Top             =   3030
      Width           =   1260
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Left            =   30
      TabIndex        =   10
      Top             =   1860
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "分所案號："
      Height          =   180
      Left            =   30
      TabIndex        =   9
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   30
      TabIndex        =   8
      Top             =   480
      Width           =   900
   End
End
Attribute VB_Name = "frm140107"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/06 Form2.0已修改 txtC/lblCU/lblST/lblMemo
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
Option Explicit

Dim strSql As String, DoIt As Boolean, IsSaveOk As Boolean
Dim m_Na01 As String 'Added by Lydia 2015/07/17

Private Sub cmdOK_Click(Index As Integer)
Dim Pcount As Long, Tcount As Long, Scount As Long, Lcount As Long, Hcount As Long
Dim updMemo As String 'Added by Lydia 2015/07/17
Select Case Index
Case 0
     If Trim(txt1(0)) = "" And Trim(txt1(4)) = "" Then
        MsgBox "條件最少要有一個要輸入！", vbExclamation
        Exit Sub
     End If
     If Trim(txt1(0)) <> "" And Trim(txt1(1)) = "" Then
        MsgBox "條件要明確！", vbExclamation
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     Me.Enabled = False
     StrMenu
     Me.Enabled = True
     Screen.MousePointer = vbDefault
Case 1
     Screen.MousePointer = vbHourglass
     IsSaveOk = False
     Pcount = 0
     Tcount = 0
     Scount = 0
     Lcount = 0
     Hcount = 0
     If pub_strUserOffice = "1" Then
        StrMenu
        If cmdOK(1).Enabled = True Then
            'Added by Lydia 2015/07/17 依操作人員的所別再更新案件備註欄
            If txt1(0) = "P" And m_Na01 = "000" Then
               updMemo = ChangeWStringToWDateString(strSrvDate(1)) & "復卷但不再立卷,北所原銷卷日期:" & ChangeTStringToWDateString(ChangeTDateStringToTString(lblD1.Caption)) & ";"
            Else
               updMemo = ChangeWStringToWDateString(strSrvDate(1)) & "復卷,北所原銷卷日期:" & ChangeTStringToWDateString(ChangeTDateStringToTString(lblD1.Caption)) & ";"
            End If
            
            'Modified by Lydia 2015/07/17 +PA91,TM58,LC27,HC12,SP18
'            cnnConnection.Execute "update patent set pa108=null where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " pa01='" & txt1(0) & "' and pa02='" & txt1(1) & "' and pa03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and pa04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' ", " pa47='" & txt1(4) & "' ") & " and pa108 is not null ", Pcount
'            cnnConnection.Execute "update trademark set tm57=null where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " tm01='" & txt1(0) & "' and tm02='" & txt1(1) & "' and tm03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and tm04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' ", " tm34='" & txt1(4) & "' ") & " and tm57 is not null ", Tcount
'            cnnConnection.Execute "update servicepractice set sp61=null where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " sp01='" & txt1(0) & "' and sp02='" & txt1(1) & "' and sp03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and sp04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "'   ", " sp28='" & txt1(4) & "' ") & " and sp61 is not null ", Scount
'            cnnConnection.Execute "update lawcase set lc34=null where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " lc01='" & txt1(0) & "' and lc02='" & txt1(1) & "' and lc03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and lc04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "'  ", " lc16='" & txt1(4) & "' ") & " and lc34 is not null ", Lcount
'            cnnConnection.Execute "update hirecase set hc19=null where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " hc01='" & txt1(0) & "' and hc02='" & txt1(1) & "' and hc03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and hc04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "'  ", " hc07='" & txt1(4) & "' ") & " and hc19 is not null ", Hcount
            cnnConnection.Execute "update patent set pa108=null,PA91='" & updMemo & "'||PA91 where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " pa01='" & txt1(0) & "' and pa02='" & txt1(1) & "' and pa03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and pa04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' ", " pa47='" & txt1(4) & "' ") & " and pa108 is not null ", Pcount
            cnnConnection.Execute "update trademark set tm57=null,TM58='" & updMemo & "'||TM58 where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " tm01='" & txt1(0) & "' and tm02='" & txt1(1) & "' and tm03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and tm04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' ", " tm34='" & txt1(4) & "' ") & " and tm57 is not null ", Tcount
            cnnConnection.Execute "update servicepractice set sp61=null,SP18='" & updMemo & "'||SP18 where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " sp01='" & txt1(0) & "' and sp02='" & txt1(1) & "' and sp03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and sp04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "'   ", " sp28='" & txt1(4) & "' ") & " and sp61 is not null ", Scount
            cnnConnection.Execute "update lawcase set lc34=null,LC27='" & updMemo & "'||LC27 where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " lc01='" & txt1(0) & "' and lc02='" & txt1(1) & "' and lc03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and lc04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "'  ", " lc16='" & txt1(4) & "' ") & " and lc34 is not null ", Lcount
            cnnConnection.Execute "update hirecase set hc19=null,HC12='" & updMemo & "'||HC12 where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " hc01='" & txt1(0) & "' and hc02='" & txt1(1) & "' and hc03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and hc04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "'  ", " hc07='" & txt1(4) & "' ") & " and hc19 is not null ", Hcount
            'end 2015/07/17
            If Pcount + Tcount + Scount + Lcount + Hcount > 0 Then
                IsSaveOk = True
            End If
        Else
            MsgBox "無本所案號可復原！", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
     Else
        StrMenu
        If cmdOK(1).Enabled = True Then
            'Added by Lydia 2015/07/17 依操作人員的所別再更新案件備註欄
            If txt1(0) = "P" And m_Na01 = "000" Then
               updMemo = ChangeWStringToWDateString(strSrvDate(1)) & "復卷但不再立卷,分所原銷卷日期:" & ChangeTStringToWDateString(ChangeTDateStringToTString(lblD2.Caption)) & ";"
            Else
               updMemo = ChangeWStringToWDateString(strSrvDate(1)) & "復卷,分所原銷卷日期:" & ChangeTStringToWDateString(ChangeTDateStringToTString(lblD2.Caption)) & ";"
            End If
            
            'Modified by Lydia 2015/07/17 +PA91,TM58,LC27,HC12,SP18
'            cnnConnection.Execute "update patent set pa136=null,pa137=null,pa138=null where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " pa01='" & txt1(0) & "' and pa02='" & txt1(1) & "' and pa03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and pa04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' ", "pa47='" & txt1(4) & "' ") & " and pa136 is not null ", Pcount
'            cnnConnection.Execute "update trademark set tm73=null,tm74=null,tm75=null where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " tm01='" & txt1(0) & "' and tm02='" & txt1(1) & "' and tm03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and tm04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' ", " tm34='" & txt1(4) & "' ") & " and tm73 is not null ", Tcount
'            cnnConnection.Execute "update servicepractice set sp68=null,sp69=null,sp70=null where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " sp01='" & txt1(0) & "' and sp02='" & txt1(1) & "' and sp03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and sp04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "'   ", " sp28='" & txt1(4) & "' ") & " and sp68 is not null ", Scount
'            cnnConnection.Execute "update lawcase set lc36=null,lc37=null,lc38=null where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " lc01='" & txt1(0) & "' and lc02='" & txt1(1) & "' and lc03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and lc04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "'  ", " lc16='" & txt1(4) & "' ") & " and lc36 is not null ", Lcount
'            cnnConnection.Execute "update hirecase set hc20=null,hc21=null,hc22=null where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " hc01='" & txt1(0) & "' and hc02='" & txt1(1) & "' and hc03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and hc04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "'  ", " hc07='" & txt1(4) & "' ") & " and hc20 is not null ", Hcount
            cnnConnection.Execute "update patent set pa136=null,pa137=null,pa138=null,PA91='" & updMemo & "'||PA91 where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " pa01='" & txt1(0) & "' and pa02='" & txt1(1) & "' and pa03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and pa04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' ", "pa47='" & txt1(4) & "' ") & " and pa136 is not null ", Pcount
            cnnConnection.Execute "update trademark set tm73=null,tm74=null,tm75=null,TM58='" & updMemo & "'||TM58 where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " tm01='" & txt1(0) & "' and tm02='" & txt1(1) & "' and tm03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and tm04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' ", " tm34='" & txt1(4) & "' ") & " and tm73 is not null ", Tcount
            cnnConnection.Execute "update servicepractice set sp68=null,sp69=null,sp70=null,SP18='" & updMemo & "'||SP18 where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " sp01='" & txt1(0) & "' and sp02='" & txt1(1) & "' and sp03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and sp04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "'   ", " sp28='" & txt1(4) & "' ") & " and sp68 is not null ", Scount
            cnnConnection.Execute "update lawcase set lc36=null,lc37=null,lc38=null,LC27='" & updMemo & "'||LC27 where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " lc01='" & txt1(0) & "' and lc02='" & txt1(1) & "' and lc03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and lc04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "'  ", " lc16='" & txt1(4) & "' ") & " and lc36 is not null ", Lcount
            cnnConnection.Execute "update hirecase set hc20=null,hc21=null,hc22=null,HC12='" & updMemo & "'||HC12 where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", " hc01='" & txt1(0) & "' and hc02='" & txt1(1) & "' and hc03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and hc04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "'  ", " hc07='" & txt1(4) & "' ") & " and hc20 is not null ", Hcount
            If Pcount + Tcount + Scount + Lcount + Hcount > 0 Then
                IsSaveOk = True
            End If
        Else
            MsgBox "無本所案號可復原！", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
     End If
     If IsSaveOk = True Then
        MsgBox "復原成功！", vbInformation
        '2012/9/19 CANCEL BY SONIA 劉經理說不要清
        'txt1(0) = ""
        'txt1(1) = ""
        'txt1(2) = ""
        'txt1(3) = ""
        'txt1(4) = ""
        '2012/9/19 END
        cmdOK(1).Enabled = False
        txtC = ""
        lblC = "'"
        lblE = ""
        lblJ = ""
        lblCU = ""
        lblMemo = ""
        lblD1 = ""
        lblD2 = ""
        lblST = ""
     ElseIf Pcount + Tcount + Scount + Lcount + Hcount = 0 Then
        MsgBox "沒有可以復原的資料！", vbInformation
     End If
     txt1(0).SetFocus
     Screen.MousePointer = vbDefault
Case 2
     Unload Me
Case Else
End Select
End Sub

Private Sub Form_Activate()
If DoIt = False Then
    If pub_strUserOffice = "1" Then txt1(0).SetFocus Else txt1(4).SetFocus
    DoIt = True
End If
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    DoIt = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm140107 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
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
   End Select
   If Cancel = True Then TextInverse txt1(Index)
End Sub

Sub StrMenu()
txtC = ""
lblC = "'"
lblE = ""
lblJ = ""
lblCU = ""
lblMemo = ""
lblD1 = ""
lblD2 = ""
lblST = ""
m_Na01 = "" 'Added by Lydia 2015/07/17
'Select Case txt1(0)
'Case "CFP", "FCP", "P"   '專利
      'Added by Lydia 2015/07/17 +PA09
      strSql = "select PA05,PA06,PA07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),pa108,pa136,pa137,pa138,st06,pa01,pa02,pa03,pa04,pa47,pa09 as na01 " & _
               " From patent,CUSTOMER,staff " & _
               " Where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", "Pa01='" & txt1(0) & "' and Pa02='" & txt1(1) & "' and Pa03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and Pa04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' ", " pa47='" & txt1(4) & "' ") & " and SUBSTR(pa26,1,8)=CU01(+) AND SUBSTR(pa26,9,1)=CU02(+) and cu13=st01(+)  "
'Case "CFT", "FCT", "T", "TF"   '商標
      'Added by Lydia 2015/07/17 +TM10
      strSql = strSql & " union Select TM05,TM06,TM07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),tm57,tm73,tm74,tm75,st06,tm01,tm02,tm03,tm04,tm34,tm10 as na01 " & _
               " From TradeMark,CUSTOMER,staff " & _
               " Where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", "tm01='" & txt1(0) & "' and tm02='" & txt1(1) & "' and tm03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and tm04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' ", " tm34='" & txt1(4) & "' ") & " and SUBSTR(tm23,1,8)=CU01(+) AND SUBSTR(tm23,9,1)=CU02(+) and cu13=st01(+)  "
'Case "CFL", "FCL", "L"          '法務
      'Added by Lydia 2015/07/17 +LC15
      strSql = strSql & " union select LC05,LC06,LC07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),lc34,lc36,lc37,lc38,st06,lc01,lc02,lc03,lc04,lc16,lc15 as na01 " & _
               " From lawcase,CUSTOMER,staff " & _
               " Where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", "lc01='" & txt1(0) & "' and lc02='" & txt1(1) & "' and lc03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and lc04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' ", " lc16='" & txt1(4) & "' ") & " and SUBSTR(lc11,1,8)=CU01(+) AND SUBSTR(lc11,9,1)=CU02(+) and cu13=st01(+)  "
'Case "LA"                      '顧問
      strSql = strSql & " union select HC06,'','',NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),hc19,hc20,hc21,hc22,st06,hc01,hc02,hc03,hc04,hc07,'000' as na01 " & _
               " From hirecase,CUSTOMER,staff " & _
               " Where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", "hc01='" & txt1(0) & "' and hc02='" & txt1(1) & "' and hc03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and hc04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' ", " hc07='" & txt1(4) & "' ") & " and SUBSTR(hc05,1,8)=CU01(+) AND SUBSTR(hc05,9,1)=CU02(+) and cu13=st01(+)  "
'Case Else                  '服務
      'Added by Lydia 2015/07/17 +SP09
      strSql = strSql & " union select SP05,SP06,SP07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),sp61,sp68,sp69,sp70,st06,sp01,sp02,sp03,sp04,sp28,sp09 as na01 " & _
               " From servicepractice,CUSTOMER,staff " & _
               " Where " & IIf(Trim(txt1(0)) & Trim(txt1(1)) <> "", "sp01='" & txt1(0) & "' and sp02='" & txt1(1) & "' and sp03='" & IIf(Trim(txt1(2)) = "", "0", txt1(2)) & "' and sp04='" & IIf(Trim(txt1(3)) = "", "00", txt1(3)) & "' ", " sp28='" & txt1(4) & "' ") & " and SUBSTR(sp08,1,8)=CU01(+) AND SUBSTR(sp08,9,1)=CU02(+) and cu13=st01(+)  "
'End Select
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 Then

    lblCU = "" & adoRecordset.Fields(3).Value
    lblD1 = ChangeWStringToTDateString("" & adoRecordset.Fields(4).Value)
    lblD2 = ChangeWStringToTDateString("" & adoRecordset.Fields(5).Value)
    lblST = GetStaffName("" & adoRecordset.Fields(6).Value, True)
    lblMemo = "" & adoRecordset.Fields(7).Value
    txt1(0) = "" & adoRecordset.Fields(9).Value
    txt1(1) = "" & adoRecordset.Fields(10).Value
    txt1(2) = "" & adoRecordset.Fields(11).Value
    txt1(3) = "" & adoRecordset.Fields(12).Value
    txt1(4) = "" & adoRecordset.Fields(13).Value
    m_Na01 = "" & adoRecordset.Fields("na01").Value 'Added by Lydia 2015/07/17
    Select Case txt1(0)
    Case "T", "FCT", "CFT", "TF", "TS", "S"
        txtC.Text = "" & adoRecordset.Fields(0).Value
        lblC.Visible = False
        lblE.Visible = False
        lblJ.Visible = False
    Case Else
        txtC.Visible = False
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
    If lblD1 <> "" Then
        cmdOK(1).Enabled = True
    ElseIf lblD2 <> "" Then
        If pub_strUserOffice = CheckStr(adoRecordset.Fields("st06").Value) Then
            cmdOK(1).Enabled = True
        Else
            cmdOK(1).Enabled = False
        End If
    Else
        cmdOK(1).Enabled = False
    End If
Else
    MsgBox "查無此案號資料！", vbExclamation
End If
End Sub
