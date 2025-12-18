VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210134_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "發E-Mail"
   ClientHeight    =   5916
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5916
   ScaleWidth      =   8952
   Begin VB.OptionButton Option1 
      Caption         =   "所有同仁："
      Height          =   180
      Index           =   2
      Left            =   1350
      TabIndex        =   4
      Top             =   3525
      Width           =   1395
   End
   Begin VB.OptionButton Option1 
      Caption         =   "與承辦人同部門之其他人："
      Height          =   180
      Index           =   1
      Left            =   1350
      TabIndex        =   2
      Top             =   3165
      Width           =   2475
   End
   Begin VB.OptionButton Option1 
      Caption         =   "承辦人"
      Height          =   180
      Index           =   0
      Left            =   1350
      TabIndex        =   1
      Top             =   2910
      Value           =   -1  'True
      Width           =   972
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "取消(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   7860
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   30
      Width           =   780
   End
   Begin VB.TextBox txtSC05 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1350
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "發E-Mail(&O)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6720
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   30
      Width           =   1080
   End
   Begin MSForms.ListBox lstMailCC 
      Height          =   900
      Index           =   1
      Left            =   4440
      TabIndex        =   7
      Top             =   4050
      Width           =   2040
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "3598;1587"
      MatchEntry      =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstMailCC 
      Height          =   900
      Index           =   0
      Left            =   2160
      TabIndex        =   6
      Top             =   4050
      Width           =   2040
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "3598;1587"
      MatchEntry      =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   330
      Index           =   1
      Left            =   3870
      TabIndex        =   5
      Top             =   3450
      Width           =   2145
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3784;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   330
      Index           =   0
      Left            =   3870
      TabIndex        =   3
      Top             =   3090
      Width           =   2145
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3784;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   825
      Left            =   1350
      TabIndex        =   8
      Top             =   4980
      Width           =   7485
      VariousPropertyBits=   -1466941413
      MaxLength       =   500
      ScrollBars      =   2
      Size            =   "13203;1455"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtSC06 
      Height          =   825
      Left            =   1350
      TabIndex        =   37
      Top             =   2010
      Width           =   7455
      VariousPropertyBits=   -1467989985
      BackColor       =   -2147483633
      MaxLength       =   500
      ScrollBars      =   2
      Size            =   "13150;1455"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   330
      Left            =   1320
      TabIndex        =   0
      Top             =   450
      Width           =   7455
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13150;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LabCP10 
      Height          =   300
      Left            =   1350
      TabIndex        =   36
      Top             =   780
      Width           =   2220
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3916;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LabCP14 
      Height          =   300
      Left            =   1350
      TabIndex        =   35
      Top             =   1080
      Width           =   1560
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2752;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "(擇一)"
      Height          =   255
      Index           =   6
      Left            =   390
      TabIndex        =   34
      Top             =   3128
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "副本：(可複選)    與承辦人同部門之其他人       所有同仁"
      Height          =   195
      Index           =   9
      Left            =   810
      TabIndex        =   33
      Top             =   3810
      Width           =   6900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "收件人："
      Height          =   255
      Index           =   8
      Left            =   390
      TabIndex        =   32
      Top             =   2910
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "內容："
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   31
      Top             =   5010
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "備註："
      Height          =   255
      Index           =   5
      Left            =   390
      TabIndex        =   30
      Top             =   2010
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "業務管制期限："
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   29
      Top             =   1710
      Width           =   1260
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "齊備日："
      Height          =   255
      Index           =   25
      Left            =   6210
      TabIndex        =   28
      Top             =   1080
      Width           =   930
   End
   Begin VB.Label LabEP06 
      Height          =   255
      Left            =   7170
      TabIndex        =   27
      Top             =   1080
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "會稿日："
      Height          =   255
      Index           =   23
      Left            =   6210
      TabIndex        =   26
      Top             =   1380
      Width           =   930
   End
   Begin VB.Label LabEP07 
      Height          =   255
      Left            =   7170
      TabIndex        =   25
      Top             =   1380
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "本所期限："
      Height          =   255
      Index           =   21
      Left            =   390
      TabIndex        =   24
      Top             =   1380
      Width           =   930
   End
   Begin VB.Label LabCP06 
      Height          =   255
      Left            =   1350
      TabIndex        =   23
      Top             =   1380
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "法定期限："
      Height          =   255
      Index           =   19
      Left            =   3390
      TabIndex        =   22
      Top             =   1380
      Width           =   930
   End
   Begin VB.Label LabCP07 
      Height          =   255
      Left            =   4350
      TabIndex        =   21
      Top             =   1380
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "承辦人："
      Height          =   255
      Index           =   17
      Left            =   390
      TabIndex        =   20
      Top             =   1080
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "承辦期限："
      Height          =   255
      Index           =   10
      Left            =   3390
      TabIndex        =   19
      Top             =   1080
      Width           =   930
   End
   Begin VB.Label LabCP48 
      Height          =   255
      Left            =   4350
      TabIndex        =   18
      Top             =   1080
      Width           =   1560
   End
   Begin VB.Label LabID 
      Height          =   255
      Left            =   1350
      TabIndex        =   17
      Top             =   150
      Width           =   1920
   End
   Begin VB.Label LabCP05 
      Height          =   255
      Left            =   4350
      TabIndex        =   16
      Top             =   150
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "收文日："
      Height          =   255
      Index           =   4
      Left            =   3390
      TabIndex        =   15
      Top             =   150
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件性質："
      Height          =   255
      Index           =   2
      Left            =   390
      TabIndex        =   14
      Top             =   780
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   11
      Left            =   90
      TabIndex        =   13
      Top             =   480
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   390
      TabIndex        =   12
      Top             =   150
      Width           =   930
   End
End
Attribute VB_Name = "frm210134_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/03 改成Form2.0 ; LabCP14、LabCP10、cmbTM05、txtSC06、Text1、Combo1(index)、lstMailCC(index)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/23 日期欄已修改
Option Explicit

'紀錄作用按鍵
Public cmdState As Integer
Public m_strSC01 As String
Public m_strSC02 As String
Public m_strSC03 As String
Dim m_strCP14 As String
Dim m_strCP14ST03 As String
Dim s As Integer


Private Sub cmdOK_Click(Index As Integer)
Dim strTo As String, strSubject As String, strContent As String
Dim ii As Integer, strToCC As String

On Error GoTo ErrHnd

cmdState = Index
Select Case cmdState
Case 1
   strTo = ""
   strToCC = ""
   '正本
   If Option1(0).Value = True Then strTo = m_strCP14
   If Option1(1).Value = True Then strTo = Left(Trim(Combo1(0).Text), 5)
   If Option1(2).Value = True Then strTo = Left(Combo1(1).Text, 5)
   If strTo = "" Then
      MsgBox "收件人空白，無法寄送！"
      Exit Sub
   End If
   '副本
   For ii = 0 To lstMailCC(0).ListCount - 1
      If lstMailCC(0).Selected(ii) = True Then
         If strToCC = "" Then
            strToCC = Left(Trim(lstMailCC(0).List(ii)), 5)
         Else
            strToCC = strToCC & ";" & Left(Trim(lstMailCC(0).List(ii)), 5)
         End If
      End If
   Next
   For ii = 0 To lstMailCC(1).ListCount - 1
      If lstMailCC(1).Selected(ii) = True Then
         If strToCC = "" Then
            strToCC = Left(Trim(lstMailCC(1).List(ii)), 5)
         Else
            strToCC = strToCC & ";" & Left(Trim(lstMailCC(1).List(ii)), 5)
         End If
      End If
   Next
   Screen.MousePointer = vbHourglass
   
   strSubject = LabID & "　智權人員同仁管制通知"
   strContent = "本所案號：" + LabID + vbCrLf + _
                       "案件名稱：" + Mid(cmbTM05.Text, 5, Len(cmbTM05)) + vbCrLf + _
                       "案件性質：" + LabCP10 + vbCrLf + _
                       "收文日　：" + LabCP05 + vbCrLf + _
                       "承辦人　：" + LabCP14 + vbCrLf + _
                       "承辦期限：" + LabCP48 + vbCrLf + _
                       "本所期限：" + LabCP06 + vbCrLf + _
                       "法定期限：" + LabCP07 + vbCrLf + _
                       "齊備日　：" + LabEP06 + vbCrLf + _
                       "會稿日　：" + LabEP07 + vbCrLf + vbCrLf + _
                       "智權人員管制期限：" + txtSC05 + vbCrLf + vbCrLf + _
                       "備　　註：" + txtSC06 + vbCrLf + vbCrLf + _
                       "內　　容：" + Text1 + vbCrLf
   
   PUB_SendMail strUserNum, strTo, "", strSubject, strContent, "", , , , , strToCC
   's = MsgBox("郵件已送出", , "MAIL!!")
   Screen.MousePointer = vbDefault
Case 0
Case Else
End Select
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
tmpBol = fnCancelNowFormAndShowParentForm(Me)
End Sub

Private Sub Combo1_Click(Index As Integer)
   Select Case Index
      Case 0
         Option1(1).Value = True
      Case 1
         Option1(2).Value = True
   End Select
End Sub

Private Sub Form_Load()
bolToEndByNick = False
MoveFormToCenter Me
cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm210134_2 = Nothing
End Sub

Public Function Process() As Boolean
On Error GoTo ErrHnd
   Process = True
   cmbTM05.Clear
   strSql = "select decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,cp09 as 總收文號,NVL(DECODE(TM10,'000',CPM03,CPM04),cp10) as 案件性質,s1.ST02 as 智權人員,s2.ST02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp48) as 承辦期限,sqldatet(ep06) as 齊備日,sqldatet(ep07) as 會稿日,TM05,TM06,TM07,s2.st03 as cp14st03,cp14" & _
                "  from caseprogress,staff s1,staff s2,acc090,casepropertymap,trademark,engineerprogress" & _
                " where cp09='" & m_strSC01 & "' and cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and cp01=cpm01(+) and cp10=cpm02(+) and cp09=ep02(+) " & _
                " union " & _
                "select decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,cp09 as 總收文號,NVL(DECODE(PA09,'000',CPM03,CPM04),cp10) as 案件性質,s1.ST02 as 智權人員,s2.ST02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp48) as 承辦期限,sqldatet(ep06) as 齊備日,sqldatet(ep07) as 會稿日,pa05 as TM05,pa06 as TM06,pa07 as TM07,s2.st03 as cp14st03,cp14" & _
                "  from caseprogress,staff s1,staff s2,acc090,casepropertymap,patent,engineerprogress" & _
                " where cp09='" & m_strSC01 & "' and cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp01=cpm01(+) and cp10=cpm02(+) and cp09=ep02(+) " & _
                " union " & _
                "select decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,cp09 as 總收文號,NVL(DECODE(SP09,'000',CPM03,CPM04),cp10) as 案件性質,s1.ST02 as 智權人員,s2.ST02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp48) as 承辦期限,sqldatet(ep06) as 齊備日,sqldatet(ep07) as 會稿日,sp05 as TM05,sp06 as TM06,sp07 as TM07,s2.st03 as cp14st03,cp14" & _
                "  from caseprogress,staff s1,staff s2,acc090,casepropertymap,servicepractice,engineerprogress" & _
                " where cp09='" & m_strSC01 & "' and cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 and cp01=cpm01(+) and cp10=cpm02(+) and cp09=ep02(+) " & _
                " union " & _
                "select decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,cp09 as 總收文號,NVL(CPM03,cp10) as 案件性質,s1.ST02 as 智權人員,s2.ST02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp48) as 承辦期限,sqldatet(ep06) as 齊備日,sqldatet(ep07) as 會稿日,hc06 as TM05,'' as TM06,'' as TM07,s2.st03 as cp14st03,cp14" & _
                "  from caseprogress,staff s1,staff s2,acc090,casepropertymap,hirecase,engineerprogress" & _
                " where cp09='" & m_strSC01 & "' and cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=hc01 and cp02=hc02 and cp03=hc03 and cp04=hc04 and cp01=cpm01(+) and cp10=cpm02(+) and cp09=ep02(+) " & _
                " union " & _
                "select decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 部門,sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,cp09 as 總收文號,NVL(DECODE(lc15,'000',CPM03,CPM04),cp10) as 案件性質,s1.ST02 as 智權人員,s2.ST02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp48) as 承辦期限,sqldatet(ep06) as 齊備日,sqldatet(ep07) as 會稿日,lc05 as TM05,lc06 as TM06,lc07 as TM07,s2.st03 as cp14st03,cp14" & _
                "  from caseprogress,staff s1,staff s2,acc090,casepropertymap,lawcase,engineerprogress" & _
                " where cp09='" & m_strSC01 & "' and cp13=s1.st01(+) and cp14=s2.st01(+) and cp12=a0901(+) and cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04 and cp01=cpm01(+) and cp10=cpm02(+) and cp09=ep02(+) " & _
                " order by 總收文號"
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         LabID.Caption = "" & .Fields("本所案號")
         ' 案件名稱
         If IsNull(.Fields("TM05")) = False Then
            cmbTM05.AddItem "中 : " & .Fields("TM05")
         End If
         If IsNull(.Fields("TM06")) = False Then
            cmbTM05.AddItem "英 : " & .Fields("TM06")
         End If
         If IsNull(.Fields("TM07")) = False Then
            cmbTM05.AddItem "日 : " & .Fields("TM07")
         End If
         ' 顯示案件名稱
         If cmbTM05.ListCount > 0 Then
            cmbTM05.ListIndex = 0
         End If
         LabCP10.Caption = "" & .Fields("案件性質")
         LabCP05.Caption = "" & .Fields("收文日")
         LabCP14.Caption = "" & .Fields("承辦人")
         LabCP48.Caption = "" & .Fields("承辦期限")
         LabCP06.Caption = "" & .Fields("本所期限")
         LabCP07.Caption = "" & .Fields("法定期限")
         LabEP06.Caption = "" & .Fields("齊備日")
         LabEP07.Caption = "" & .Fields("會稿日")
         m_strCP14ST03 = "" & .Fields("cp14st03")
         m_strCP14 = "" & .Fields("cp14")
      Else
         LabID.Caption = ""
         LabCP10.Caption = ""
         LabCP05.Caption = ""
         LabCP14.Caption = ""
         LabCP48.Caption = ""
         LabCP06.Caption = ""
         LabCP07.Caption = ""
         LabEP06.Caption = ""
         LabEP07.Caption = ""
         m_strCP14ST03 = ""
         m_strCP14 = ""
      End If
   End With
   txtSC05.Text = "": txtSC06.Text = ""
   If m_strSC02 <> "" And m_strSC03 <> "" Then
      strSql = "select sqldatet(sc02) as 異動日期,st02 as 異動人員,sqldatet(sc05) as 管制期限,sc08 as 已完成,sc06 as 備註 from salescontroldate,staff where sc01='" & m_strSC01 & "' and sc02=" & m_strSC02 & " and sc03=" & m_strSC03 & " and sc04=st01(+) "
      CheckOC3
      With AdoRecordSet3
         .CursorLocation = adUseClient
         .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If .RecordCount > 0 Then
            txtSC05.Text = "" & .Fields("管制期限")
            txtSC06.Text = "" & .Fields("備註")
         End If
      End With
   End If
   '與承辦人同部門之其他人
   Combo1(0).Clear
   strSql = "SELECT st01,st02 FROM staff,SalaryData WHERE substr(st03,1,2)='" & Left(m_strCP14ST03, 2) & "' and st04='1' and st01>'6' and st01<'F' and st01=sd01 order by st01 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         RsTemp.MoveFirst
         Do While RsTemp.EOF = False
            Combo1(0).AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
            RsTemp.MoveNext
         Loop
      End With
   End If
   '另選收件人
   Combo1(1).Clear
   strSql = "SELECT st01,st02 FROM staff,SalaryData WHERE st04='1' and st01>'6' and st01<'F' and st01=sd01 order by st03 asc,st01 asc "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         RsTemp.MoveFirst
         Do While RsTemp.EOF = False
            Combo1(1).AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
            RsTemp.MoveNext
         Loop
      End With
   End If
   Option1(0).Value = True
   Text1.Text = ""
   lstMailCC(0).Clear
   lstMailCC(1).Clear
   '副本：與承辦人同部門之其他人
   strSql = "SELECT st01,st02 FROM staff,SalaryData WHERE substr(st03,1,2)='" & Left(m_strCP14ST03, 2) & "' and st04='1' and st01>'6' and st01<'F' and st01=sd01 order by st01 asc "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         RsTemp.MoveFirst
         Do While RsTemp.EOF = False
            lstMailCC(0).AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
            RsTemp.MoveNext
         Loop
      End With
   End If
   '副本：另選收件人
   strSql = "SELECT st01,st02 FROM staff,SalaryData WHERE st04='1' and st01>'6' and st01<'F' and st01=sd01 order by st03 asc,st01 asc "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         RsTemp.MoveFirst
         Do While RsTemp.EOF = False
            lstMailCC(1).AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
            RsTemp.MoveNext
         Loop
      End With
   End If
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical: Process = False
End Function

Private Sub Combo1_GotFocus(Index As Integer)
   InverseTextBox Combo1(Index)
End Sub

'Modified by Lydia 2022/01/03 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Index As Integer, Cancel As Boolean)
   Dim i As Integer, strText As String
   Cancel = False
   Select Case Index
      Case 0
         If Trim(Combo1(0).Text) <> "" Then
            Option1(1).Value = True
            Combo1(1).Text = ""
            For i = 0 To Combo1(0).ListCount
               If Left(Trim(Combo1(0).Text), 5) = Left(Trim(Combo1(0).List(i)), 5) Then Exit Sub
            Next i
            MsgBox "此人不在下拉式選單裡!!!", vbExclamation
            Combo1(0).SetFocus
            Cancel = True
            Exit Sub
         End If
      Case 1
         If Trim(Combo1(1).Text) <> "" Then
            Option1(2).Value = True
            Combo1(0).Text = ""
            For i = 0 To Combo1(1).ListCount
               If Left(Trim(Combo1(1).Text), 5) = Left(Trim(Combo1(1).List(i)), 5) Then Exit Sub
            Next i
            MsgBox "此人不在下拉式選單裡!!!", vbExclamation
            Combo1(1).SetFocus
            Cancel = True
            Exit Sub
         End If
   End Select
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
         Combo1(0).Text = ""
         Combo1(1).Text = ""
      Case 1
         Combo1(1).Text = ""
      Case 2
         Combo1(0).Text = ""
   End Select
End Sub
