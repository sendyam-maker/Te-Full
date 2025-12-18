VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210136_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "發E-Mail"
   ClientHeight    =   5730
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8950
   Begin VB.OptionButton Option1 
      Caption         =   "所有同仁："
      Height          =   180
      Index           =   2
      Left            =   1350
      TabIndex        =   4
      Top             =   2430
      Width           =   1395
   End
   Begin VB.OptionButton Option1 
      Caption         =   "與承辦人同部門之其他人："
      Height          =   180
      Index           =   1
      Left            =   1350
      TabIndex        =   2
      Top             =   2190
      Width           =   2475
   End
   Begin VB.OptionButton Option1 
      Caption         =   "承辦人"
      Height          =   180
      Index           =   0
      Left            =   1350
      TabIndex        =   1
      Top             =   1950
      Value           =   -1  'True
      Width           =   972
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "取消(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   0
      Left            =   7890
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   120
      Width           =   780
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "發E-Mail(&O)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   6750
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   120
      Width           =   1080
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   300
      Left            =   1350
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   7455
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "13150;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstMailCC 
      Height          =   900
      Index           =   1
      Left            =   4440
      TabIndex        =   7
      Top             =   3000
      Width           =   2040
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "3598;1587"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
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
      Top             =   3000
      Width           =   2040
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "3598;1587"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Index           =   1
      Left            =   3870
      TabIndex        =   5
      Top             =   2370
      Width           =   2145
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3784;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Index           =   0
      Left            =   3870
      TabIndex        =   3
      Top             =   2040
      Width           =   2145
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3784;529"
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
      Top             =   3930
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
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "(擇一)"
      Height          =   255
      Index           =   6
      Left            =   390
      TabIndex        =   33
      Top             =   2190
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "副本：(可複選)    與承辦人同部門之其他人        所有同仁"
      Height          =   195
      Index           =   9
      Left            =   780
      TabIndex        =   32
      Top             =   2760
      Width           =   6900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "收件人："
      Height          =   255
      Index           =   8
      Left            =   390
      TabIndex        =   31
      Top             =   1950
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "內容："
      Height          =   255
      Index           =   1
      Left            =   390
      TabIndex        =   30
      Top             =   3930
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "齊備日："
      Height          =   255
      Index           =   25
      Left            =   6210
      TabIndex        =   29
      Top             =   1350
      Width           =   930
   End
   Begin VB.Label LabEP06 
      Height          =   255
      Left            =   7170
      TabIndex        =   28
      Top             =   1350
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "會稿日："
      Height          =   255
      Index           =   23
      Left            =   6210
      TabIndex        =   27
      Top             =   1650
      Width           =   930
   End
   Begin VB.Label LabEP07 
      Height          =   255
      Left            =   7170
      TabIndex        =   26
      Top             =   1650
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "本所期限："
      Height          =   255
      Index           =   21
      Left            =   390
      TabIndex        =   25
      Top             =   1650
      Width           =   930
   End
   Begin VB.Label LabCP06 
      Height          =   255
      Left            =   1350
      TabIndex        =   24
      Top             =   1650
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "法定期限："
      Height          =   255
      Index           =   19
      Left            =   3390
      TabIndex        =   23
      Top             =   1650
      Width           =   930
   End
   Begin VB.Label LabCP07 
      Height          =   255
      Left            =   4350
      TabIndex        =   22
      Top             =   1650
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "承辦人："
      Height          =   255
      Index           =   17
      Left            =   390
      TabIndex        =   21
      Top             =   1350
      Width           =   930
   End
   Begin MSForms.Label LabCP14 
      Height          =   255
      Left            =   1350
      TabIndex        =   20
      Top             =   1350
      Width           =   1560
      VariousPropertyBits=   27
      Size            =   "2752;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "承辦期限："
      Height          =   255
      Index           =   10
      Left            =   3390
      TabIndex        =   19
      Top             =   1350
      Width           =   930
   End
   Begin VB.Label LabCP48 
      Height          =   255
      Left            =   4350
      TabIndex        =   18
      Top             =   1350
      Width           =   1560
   End
   Begin VB.Label LabID 
      Height          =   255
      Left            =   1350
      TabIndex        =   17
      Top             =   420
      Width           =   1920
   End
   Begin VB.Label LabCP05 
      Height          =   255
      Left            =   4350
      TabIndex        =   16
      Top             =   420
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "收文日："
      Height          =   255
      Index           =   4
      Left            =   3390
      TabIndex        =   15
      Top             =   420
      Width           =   930
   End
   Begin VB.Label LabCP10 
      Height          =   255
      Left            =   1350
      TabIndex        =   14
      Top             =   1050
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件性質："
      Height          =   255
      Index           =   2
      Left            =   390
      TabIndex        =   13
      Top             =   1050
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   11
      Left            =   390
      TabIndex        =   12
      Top             =   720
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   390
      TabIndex        =   11
      Top             =   420
      Width           =   930
   End
End
Attribute VB_Name = "frm210136_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/18 改成Form2.0 (cmbTM05,LabCP14,Combo1,lstMailCC,Text1)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Create By Sindy 2012/5/9
Option Explicit

'紀錄作用按鍵
Public cmdState As Integer
Dim m_strCP14 As String
Dim m_strCP14ST03 As String
'Dim s As Integer


Private Sub cmdOK_Click(Index As Integer)
Dim strTo As String, strSubject As String, strContent As String
Dim ii As Integer, strToCC As String

On Error GoTo ErrHnd

cmdState = Index
Select Case cmdState
Case 1
   'Added by Morgan 2022/1/18 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Sub
   End If
   'end 2022/1/18
   
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
   
   'Add By Sindy 2024/8/1
   If Trim(Text1) = "" Then
      MsgBox "內容不可空白！"
      Text1.SetFocus
      Exit Sub
   End If
   '2024/8/1 END
   
   Screen.MousePointer = vbHourglass
   'Modified by Lydia 2018/12/10 台灣商標爭議案=>台灣商標案
   'Modified by Lydia 2022/07/15 台灣商標案=>商標著作權案件
   strSubject = LabID & "　商標著作權案件齊備日維護通知"
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
Set frm210136_2 = Nothing
End Sub

Public Function Process(strText As String) As Boolean
On Error GoTo ErrHnd
   Process = True
   cmbTM05.Clear
   'Modified by Lydia 2018/12/10 開放T台灣案管控文件齊備
   'strSql = "select sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,cpm03 as 案件性質,tm05 as 案件名稱," & _
            "s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(ep06) as 齊備日," & _
            "sqldatet(cp48) as 承辦期限,sqldatet(ep28) as 指定會稿日,ep34 as 是否會稿,sqldatet(ep07) as 會稿日,sqldatet(cp27) as 發文日," & _
            "cp16 As 費用, cp18 As 點數, cp64 As 進度備註, cp09 As 總收文號,cp122,cp14,s2.st03 as cp14st03" & _
            " from caseprogress,engineerprogress,trademark,casepropertymap,staff s1,staff s2" & _
            " where cp01 in('T','FCT') and cp10 in (" & TMdebate & ")" & _
            " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)" & _
            " and cp09=ep02(+)" & _
            " and tm10='000'" & _
            " and cp01=cpm01(+) and cp10=cpm02(+)" & _
            " and cp13=s1.st01(+)" & _
            " and cp14=s2.st01(+)" & _
            " and cp09='" & strText & "'"
   'Modified by Lydia 2022/07/15 T大陸案之齊備日管控: tm10='000' => tm10 in ('000','020')、cpm03 => decode(tm10,'000',cpm03,cpm04)
   strSql = "select sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,decode(tm10,'000',cpm03,cpm04) as 案件性質,tm05 as 案件名稱," & _
            "s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(ep06) as 齊備日," & _
            "sqldatet(cp48) as 承辦期限,sqldatet(ep28) as 指定會稿日,ep34 as 是否會稿,sqldatet(ep07) as 會稿日,sqldatet(cp27) as 發文日," & _
            "cp16 As 費用, cp18 As 點數, cp64 As 進度備註, cp09 As 總收文號,cp122,cp14,s2.st03 as cp14st03" & _
            " from caseprogress,engineerprogress,trademark,casepropertymap,staff s1,staff s2" & _
            " where cp01 in('T','FCT') " & _
            " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)" & _
            " and cp09=ep02(+)" & _
            " and tm10 in ('000','020')" & _
            " and cp01=cpm01(+) and cp10=cpm02(+)" & _
            " and cp13=s1.st01(+)" & _
            " and cp14=s2.st01(+)" & _
            " and cp09='" & strText & "'"
   'Added by Lydia 2022/07/15 TC案之文件齊備日管控: 臺灣、大陸
   strSql = strSql & " Union select sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,decode(sp09,'000',cpm03,cpm04) as 案件性質,sp05 as 案件名稱," & _
            "s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(ep06) as 齊備日," & _
            "sqldatet(cp48) as 承辦期限,sqldatet(ep28) as 指定會稿日,ep34 as 是否會稿,sqldatet(ep07) as 會稿日,sqldatet(cp27) as 發文日," & _
            "cp16 As 費用, cp18 As 點數, cp64 As 進度備註, cp09 As 總收文號,cp122,cp14,s2.st03 as cp14st03" & _
            " from caseprogress,engineerprogress,servicepractice,casepropertymap,staff s1,staff s2" & _
            " where cp01 in('TC') " & _
            " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)" & _
            " and cp09=ep02(+)" & _
            " and sp09 in ('000','020')" & _
            " and cp01=cpm01(+) and cp10=cpm02(+)" & _
            " and cp13=s1.st01(+)" & _
            " and cp14=s2.st01(+)" & _
            " and cp09='" & strText & "'"
   'end 2022/06/23
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         LabID.Caption = "" & .Fields("本所案號")
         ' 案件名稱
         If IsNull(.Fields("案件名稱")) = False Then
            cmbTM05.AddItem .Fields("案件名稱")
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
         MsgBox "查無資料！", vbExclamation
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
         Exit Function
      End If
   End With
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
            'Modified by Morgan 2022/1/18
            'For i = 0 To Combo1(0).ListCount
            For i = 0 To Combo1(0).ListCount - 1
            'end 2022/1/18
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
