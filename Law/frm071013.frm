VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm071013 
   BorderStyle     =   1  '單線固定
   Caption         =   "開庭通知"
   ClientHeight    =   5652
   ClientLeft      =   456
   ClientTop       =   456
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5652
   ScaleWidth      =   8952
   Begin VB.CommandButton cmdBrief 
      Caption         =   "上傳開庭紀要"
      Height          =   375
      Left            =   1440
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   75
      Width           =   1485
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "上傳開庭通知"
      Height          =   375
      Left            =   2970
      TabIndex        =   18
      Top             =   75
      Width           =   1485
   End
   Begin VB.CheckBox Check1 
      Caption         =   "是否取消前次庭期"
      Height          =   315
      Left            =   5700
      TabIndex        =   8
      Top             =   3515
      Width           =   2565
   End
   Begin VB.CommandButton Command4 
      Caption         =   "其他出庭律師(&L)"
      Height          =   375
      Left            =   4485
      TabIndex        =   19
      Top             =   75
      Width           =   1485
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8050
      TabIndex        =   22
      Top             =   75
      Width           =   765
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   21
      Top             =   75
      Width           =   1170
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   20
      Top             =   75
      Width           =   800
   End
   Begin MSForms.ComboBox cboGov 
      Height          =   324
      Left            =   4128
      TabIndex        =   0
      Top             =   1824
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
      Height          =   285
      Index           =   8
      Left            =   6120
      TabIndex        =   5
      Top             =   3195
      Width           =   492
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "868;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   285
      Index           =   7
      Left            =   1230
      TabIndex        =   3
      Top             =   2860
      Width           =   492
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "868;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   285
      Index           =   2
      Left            =   3750
      TabIndex        =   15
      Top             =   4560
      Width           =   1200
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "2117;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   2475
      TabIndex        =   14
      Top             =   4560
      Width           =   1200
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "2117;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   1230
      TabIndex        =   13
      Top             =   4560
      Width           =   1200
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "1931;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3750
      TabIndex        =   12
      Top             =   4215
      Width           =   1200
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "2117;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2475
      TabIndex        =   11
      Top             =   4215
      Width           =   1200
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "2117;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1230
      TabIndex        =   10
      Top             =   4215
      Width           =   1200
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "2117;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   285
      Index           =   2
      Left            =   1230
      TabIndex        =   2
      Top             =   2525
      Width           =   3948
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "6964;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   285
      Index           =   3
      Left            =   1230
      TabIndex        =   4
      Top             =   3195
      Width           =   492
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "868;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   585
      Index           =   6
      Left            =   1230
      TabIndex        =   16
      Top             =   4890
      Width           =   7335
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "12938;1032"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   285
      Index           =   5
      Left            =   4125
      TabIndex        =   7
      Top             =   3530
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   4
      Size            =   "1931;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   285
      Index           =   4
      Left            =   1230
      TabIndex        =   6
      Top             =   3530
      Width           =   1215
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "2143;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   285
      Index           =   1
      Left            =   1230
      TabIndex        =   1
      Top             =   2190
      Width           =   3948
      VariousPropertyBits=   671105051
      MaxLength       =   32
      Size            =   "6964;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboPerson 
      Height          =   300
      Left            =   1230
      TabIndex        =   9
      Top             =   3865
      Width           =   1935
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3413;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label9 
      Height          =   285
      Left            =   4920
      TabIndex        =   50
      Top             =   515
      Width           =   1455
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "2566;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbeName 
      Height          =   285
      Left            =   3210
      TabIndex        =   49
      Top             =   3873
      Width           =   1545
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "2725;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbeTitle 
      Height          =   285
      Left            =   4920
      TabIndex        =   48
      Top             =   850
      Width           =   2775
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "4895;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbeCusName 
      Height          =   285
      Left            =   2130
      TabIndex        =   47
      Top             =   1520
      Width           =   6375
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "11245;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   285
      Left            =   1290
      TabIndex        =   46
      Top             =   1185
      Width           =   7485
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13203;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label13 
      Caption         =   "通知方式：             (1.書面通知 2.口頭告知)"
      Height          =   255
      Left            =   5235
      TabIndex        =   45
      Top             =   3210
      Width           =   3495
   End
   Begin VB.Label Label12 
      Caption         =   "開  庭  別：              (1.民事庭  2.偵查庭  3.刑事庭  4.刑附民庭  5.行政庭  6.調解庭)"
      Height          =   255
      Left            =   255
      TabIndex        =   44
      Top             =   2875
      Width           =   7005
   End
   Begin VB.Label Label10 
      Caption         =   "智權人員："
      Height          =   255
      Left            =   3885
      TabIndex        =   43
      Top             =   530
      Width           =   960
   End
   Begin VB.Label Label4 
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   255
      TabIndex        =   42
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lbeAccpet 
      Height          =   285
      Left            =   1305
      TabIndex        =   41
      Top             =   1855
      Width           =   1455
   End
   Begin VB.Label Label19 
      Caption         =   "檢  察  官："
      Height          =   255
      Left            =   255
      TabIndex        =   40
      Top             =   4575
      Width           =   975
   End
   Begin VB.Label Label17 
      Caption         =   "法        官："
      Height          =   255
      Left            =   255
      TabIndex        =   39
      Top             =   4230
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "股　　別："
      Height          =   255
      Left            =   255
      TabIndex        =   38
      Top             =   2540
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "當事人稱謂："
      Height          =   255
      Left            =   3720
      TabIndex        =   37
      Top             =   865
      Width           =   1125
   End
   Begin VB.Label lbeCaseNum 
      Height          =   285
      Left            =   1305
      TabIndex        =   26
      Top             =   850
      Width           =   1815
   End
   Begin VB.Label lbePaperNum 
      Height          =   285
      Left            =   1305
      TabIndex        =   24
      Top             =   515
      Width           =   1695
   End
   Begin VB.Label lbeCus 
      Height          =   285
      Left            =   1305
      TabIndex        =   25
      Top             =   1520
      Width           =   750
   End
   Begin VB.Label Label6 
      Caption         =   "當  事  人："
      Height          =   255
      Left            =   255
      TabIndex        =   36
      Top             =   1535
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "開庭種類：              (1.偵查  2.審理  3.言詞辯論  4.調查  5.調解)"
      Height          =   255
      Left            =   255
      TabIndex        =   35
      Top             =   3210
      Width           =   4905
   End
   Begin VB.Label Label3 
      Caption         =   "開庭日期："
      Height          =   255
      Left            =   255
      TabIndex        =   34
      Top             =   3545
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   "備        註："
      Height          =   255
      Left            =   255
      TabIndex        =   33
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "時間："
      Height          =   195
      Left            =   3510
      TabIndex        =   32
      Top             =   3575
      Width           =   585
   End
   Begin VB.Label Label8 
      Caption         =   "法院案號："
      Height          =   255
      Left            =   255
      TabIndex        =   31
      Top             =   2205
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "開庭人員："
      Height          =   255
      Left            =   255
      TabIndex        =   30
      Top             =   3888
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號：  "
      Height          =   255
      Left            =   255
      TabIndex        =   29
      Top             =   865
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "收  文  號："
      Height          =   255
      Index           =   0
      Left            =   255
      TabIndex        =   28
      Top             =   530
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "收  受  日："
      Height          =   255
      Left            =   255
      TabIndex        =   27
      Top             =   1870
      Width           =   975
   End
   Begin VB.Label Label25 
      Caption         =   "機關代號："
      Height          =   255
      Left            =   3120
      TabIndex        =   23
      Top             =   1870
      Width           =   975
   End
End
Attribute VB_Name = "frm071013"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/14 改成Form2.0 ; cboCaseName、lbeCusName、lbeTitle、lbeName、Label9、Text(index)、Text1(index)、Text2(index)、cboPerson
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim strCP09() As String, t As Integer
Dim cp01 As String, cp02 As String, cp03 As String, cp04 As String, Strsale As String, strSaleArea As String
Dim blnIsSave As Boolean, PaperNum As String
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Dim m_Nation As String
Dim strTemp As Variant 'Add By Sindy 2011/6/9
Dim strNum As String 'Modify By Sindy 2011/6/9
Dim m_StrTo As String, m_contents As String 'Add By Sindy 2011/7/21
Dim strOpposite As String   '2011/9/15 add by sonia對造名稱
Public m_strCancelRecv 'Add By Sindy 2011/10/20 要取消前次庭期的收文文號
'Added By Lydia 2015/10/30 上傳開庭通知
'Public m_strSaveFiles As String '新增附件
Public m_strSaveFileType As String '欲新增的檔案種類:1.開庭通知 2.開庭紀要 Modify By Sindy 2016/6/24
'Dim m_AttachPath As String '附件暫存區
Public m_strSaveFilesOA As String '新增開庭通知附件 Add By Sindy 2016/6/24
Public m_strSaveFilesBRIEF As String '新增開庭紀要附件 Add By Sindy 2016/6/24
Dim m_CP14 As String '承辦人
Dim m_CP29 As String '協辦人員
Dim m_CL02 As String 'Added by Lydia 2022/12/08 出庭律師
Dim m_GovListNew As String, m_GovListDef As String  'Added by Lydia 2025/11/18

Private Sub cboPerson_Click()
Dim nPos As Integer
Dim strPerson As String
  
  nPos = 0
   If cboPerson.Text <> "" Then
       nPos = InStr(cboPerson.Text, ",")
       If nPos <> 0 Then
          strPerson = Left(cboPerson.Text, nPos - 1)
          lbeName.Caption = ChgType(2, strPerson)
       End If
   End If
End Sub

Private Sub cboPerson_Validate(Cancel As Boolean)
   If cboPerson = "" Then Cancel = True
End Sub

Private Sub cmdBack_Click()
   If blnIsSave = False Then
      If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then Exit Sub
   End If
   frm071012.Show
   Unload Me
End Sub

Private Sub cmdEnd_Click()
   If blnIsSave = False Then
      If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then Exit Sub
   End If
   Unload frm071012
   Unload Me
End Sub

Private Sub cmdSure_Click()
Dim strDay1 As String, strDay2 As String, LcTmp As String, strcp1 As String
Dim strDate As String
Dim i As Integer
Dim strCC As String
   
   If AllTextBeforeSaveCheck Then Exit Sub
   
   'Add By Sindy 2011/10/20 若有勾選取消前次庭期時,必須顯示該案號尚未到期的庭期資料供使用者做勾選
   m_strCancelRecv = ""
   If Check1.Value = 1 Then
      frm071013_1.m_CP01 = m_CP01
      frm071013_1.m_CP02 = m_CP02
      frm071013_1.m_CP03 = m_CP03
      frm071013_1.m_CP04 = m_CP04
      If frm071013_1.doQuery Then
         frm071013_1.Show vbModal
         Unload frm071013_1
         Set frm071013_1 = Nothing
      Else
         Unload frm071013_1
         Set frm071013_1 = Nothing
         MsgBox "該案號無未到期的庭期資料可供使用者做勾選取消！", vbCritical
         Check1.SetFocus
         Exit Sub
      End If
      
      If m_strCancelRecv = "" Then
         MsgBox "若要取消前次庭期，必須勾選欲取消那幾筆庭期資料！", vbCritical
         Check1.SetFocus
         Exit Sub
      End If
   End If
   
   strcp1 = GiveSymbol(cp01, cp02, cp03, cp04, LcTmp)
'   If Not objLawDll.ChkMRec(ChangeTStringToWString(Replace(lbeAccpet, "/", "")), LcTmp, strDay1, strDay2) Then
'      MsgBox "本所案號'" + lbeCaseNum + "'與收受日'" + lbeAccpet + "'不存在於來函記錄檔中", vbCritical
'      Exit Sub
'   End If
'   If Text(4) <> strDay2 Then
'      If MsgBox("來函記錄檔之開庭日期 '" + ChangeWStringToTString(strDay2) + "' 與畫面輸入不同，是否儲存 ? ", vbCritical + vbYesNo) = vbNo Then
'         Text(3).SetFocus
'         Exit Sub
'      End If
'   End If
   GetNation
'2011/5/17 cancel by sonia
'   If m_CP01 = "LA" Or m_Nation = "000" Then
'         strDate = GetMailRecField(m_CP01, m_CP02, m_CP03, m_CP04, DBDATE(lbeAccpet.Caption), "MR17")
'     '    If IsEmptyText(strDate) = False Then
'            If (TAIWANDATE(Text(4).Text) <> TAIWANDATE(strDate)) Or strDate = "" Then
'               If MsgBox("輸入的法定期限與來函記錄中的法定期限日期不同", vbYesNo, "資料檢核") = vbNo Then
'                  Text(4).SetFocus
'                  Exit Sub
'               End If
'            End If
'      '   End If
'   End If
'2011/5/17 end
   'Add By Cheng 2002/05/24
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
   
   If Not SaveData Then
      DataErrorMessage (3)
   Else
      'Add By Sindy 2011/7/21 寄Mail通知最新智權人員
      If cp01 = "FCL" Or cp01 = "LIN" Then
         m_StrTo = PUB_GetFCLSalesNo(cp01, cp02, cp03, cp04)
      Else
         m_StrTo = PUB_GetAKindSalesNo(cp01, cp02, cp03, cp04)
      End If
      'Modify By Sindy 2020/6/10 原只發EMAIL給案件最新智權人員，請修改為若有案源則發給案源介紹人(可能多個)，無案源才發給案件最新智權人員。
      strExc(0) = "select * from LawOfficeSource where LOS06='" & lbePaperNum & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_StrTo = Replace(RsTemp.Fields("LOS04"), ",", ";")
      End If
      '2020/6/10 END
      
      'Added by Lydia 2015/10/30 上傳開庭通知附件存在卷宗區
      If m_strSaveFilesOA <> "" Then
          If PUB_UpdReplyFile(m_strSaveFilesOA, strNum, m_CP01, m_CP02, m_CP03, m_CP04, 通知開庭, , "OA") = False Then Exit Sub
          '刪除匯入來源檔
          Call PUB_DelPCOrgFile(m_strSaveFilesOA): m_strSaveFilesOA = ""
      End If
      'end 2015/10/30
      
      'Add By Sindy 2016/6/24 上傳開庭紀要附件存到卷宗區
      If m_strSaveFilesBRIEF <> "" Then
          If PUB_UpdReplyFile(m_strSaveFilesBRIEF, strNum, m_CP01, m_CP02, m_CP03, m_CP04, 通知開庭, , UCase("BRIEF")) = False Then Exit Sub
          '刪除匯入來源檔
          Call PUB_DelPCOrgFile(m_strSaveFilesBRIEF): m_strSaveFilesBRIEF = ""
      End If
      'end 2015/10/30
   
      If m_StrTo > "" Then
         '開庭日期大於系統日才寄發Mail
         If Val(Text(4)) - Val(GetTaiwanTodayDate) > 0 Then
            '2011/9/15 modify by sonia 楊世安通知改內容
            'm_contents = Label21 & lbeAccpet & vbCrLf & _
                         Label25 & Text(0) & lbeGov & vbCrLf & _
                         Label8 & Text(1) & vbCrLf & _
                         Left(Label5, 5) & Text(3) & "(1.偵查 2.審理 3.辯論)" & vbCrLf & _
                         Label3 & Text(4) & "　　　" & Label11 & Text(5) & vbCrLf & _
                         Label7 & cboPerson.Text & vbCrLf & _
                         Label17 & Text1(0) & "  " & Text1(1) & "  " & Text1(2) & vbCrLf & _
                         Label19 & Text2(0) & "  " & Text2(1) & "  " & Text2(2) & vbCrLf & _
                         Label18 & Text(6)
            m_contents = "本所案號：" & cp01 & "-" & cp02 & "-" & cp03 & "-" & cp04 & vbCrLf & _
                         "智權人員：" & Label9 & vbCrLf & _
                         "當 事 人：" & lbeCusName & vbCrLf
            If strOpposite <> "" Then m_contents = m_contents & "對造名稱：" & strOpposite & vbCrLf
            'Modify By Sindy 2011/10/20 增加開庭別,修改開庭種類
            'Modify By Sindy 2018/01/24 增加開庭別-調解庭,增加開庭種類-調解
            'Modified by Lydia 2025/11/18 lbeGov 改為 Trim(Mid(cboGov.Text, 5))
            m_contents = m_contents & vbCrLf & _
                         "收 受 日：" & lbeAccpet & vbCrLf & _
                         "開庭機關：" & Trim(Mid(cboGov.Text, 5)) & vbCrLf & _
                         "開 庭 別：" & IIf(Text(7) = "1", "民事庭", IIf(Text(7) = "2", "偵查庭", IIf(Text(7) = "3", "刑事庭", IIf(Text(7) = "4", "刑附民庭", IIf(Text(7) = "5", "行政庭", "調解庭"))))) & vbCrLf & _
                         "法院案號：" & Text(1) & vbCrLf & _
                         "開庭日期：" & ChangeTStringToTDateString(Text(4)) & "　　　" & "時間：" & Format(Text(5), "##:##") & vbCrLf & _
                         "開庭種類：" & IIf(Text(3) = "1", "偵查", IIf(Text(3) = "2", "審理", IIf(Text(3) = "3", "言詞辯論", IIf(Text(3) = "4", "調查", "調解")))) & vbCrLf & _
                         "開庭人員：" & Mid(cboPerson.Text, 7) & vbCrLf
            If Text1(0) <> "" Or Text1(1) <> "" Or Text1(2) <> "" Then
               m_contents = m_contents & _
                         "法    官：" & Text1(0) & "  " & Text1(1) & "  " & Text1(2) & vbCrLf
            End If
            If Text2(0) <> "" Or Text2(1) <> "" Or Text2(2) <> "" Then
               m_contents = m_contents & _
                         "檢 察 官：" & Text2(0) & "  " & Text2(1) & "  " & Text2(2) & vbCrLf
            End If
            m_contents = m_contents & "備    註：" & Text(6) & vbCrLf
            'Add By Sindy 2011/10/20 告知同時取消的開庭日期及時間
            If Check1.Value = 1 Then
               strTemp = Split(m_strCancelRecv, ",")
               For i = 0 To UBound(strTemp) - 1
                  strExc(0) = "select cdp03,cdp04 from courtyardperiod where cdp01='" & strTemp(i) & "' order by cdp03 asc"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If i = 0 Then
                        m_contents = m_contents & "取消前次開庭：" & ChangeWStringToTDateString(RsTemp.Fields("cdp03")) & " " & Format(RsTemp.Fields("cdp04"), "##:##") & vbCrLf
                     Else
                        m_contents = m_contents & "　　　　　　　" & ChangeWStringToTDateString(RsTemp.Fields("cdp03")) & " " & Format(RsTemp.Fields("cdp04"), "##:##") & vbCrLf
                     End If
                  End If
               Next i
            End If
            'Modify By Sindy 2020/6/10
            '1. 原只發EMAIL給案件最新智權人員，請修改為若有案源則發給案源介紹人(可能多個)，無案源才發給案件最新智權人員。
            '2. 加發副本給該C類收文號之相關總收文號之承辦人、協辦人員及所有的出庭律師(以收文號讀取caselawer)。
            strCC = ""
            If strPublicTemp <> "" Then '出庭律師
               strCC = strCC & ";" & Replace(strPublicTemp, ",", ";")
            End If
            If m_CP14 <> "" Then strCC = strCC & ";" & m_CP14
            If m_CP29 <> "" Then strCC = strCC & ";" & m_CP29
            If strCC <> "" Then
               strCC = Replace(strCC, ";;", ";")
               strCC = Mid(strCC, 2)
            End If
            '2020/6/10 END
            PUB_SendMail strUserNum, m_StrTo, "", cp01 & "-" & cp02 & "-" & cp03 & "-" & cp04 & "開庭通知：開庭日期：" & ChangeTStringToTDateString(Text(4)) & " " & "時間：" & Format(Text(5), "##:##"), m_contents, , , , , , strCC
         End If
      End If
      '2011/7/21 End
      
      Unload frm071013
      Set frm071013 = Nothing
      Unload frm071012
      frm071012.Show
   End If
End Sub

'Add By Sindy 2011/6/9
Private Sub Command4_Click()
   'Modified by Lydia 2022/12/08
   'frm071018.Hide
   'Set frm071018.UpForm = frm071013
   'frm071018.lbePaperNum = Me.lbePaperNum
   'frm071018.lbeNumber = Me.lbeCaseNum
   'Modified by Lydia 2025/03/19 + AddCP=True >> , , , True
   Call frm071018.SetParent(Me, Me.lbePaperNum, IIf(Me.Tag = "", True, False), Replace(Trim(Left(cboPerson, 6)), ",", ""), , , True)
   'end 2022/12/08
   Me.Hide
   frm071018.Show vbModal
End Sub

Private Sub Form_Activate()
   Text(3).SetFocus 'Add By Sindy 2011/10/20
End Sub

Private Sub Form_Load()
Dim i As Integer, n As Integer
Dim strCP09 As String
  
  m_CP01 = frm071012.txtcp01.Text
  m_CP02 = frm071012.txtcp02.Text
  If frm071012.txtcp03.Text <> "" Then
     m_CP03 = frm071012.txtcp03.Text
  Else
     m_CP03 = "0"
  End If
  If frm071012.txtcp04.Text <> "" Then
     m_CP04 = frm071012.txtcp04.Text
  Else
     m_CP04 = "00"
  End If

   MoveFormToCenter Me
   n = 0
   PaperNum = ""
   With frm071012.MSHFlexGrid1
      For i = 1 To .Rows - 1
         '.Col = 0
         .row = i
         If .Text = "v" Then
            .col = 2
            'If n = 0 Then
             PaperNum = .Text
             strCP09 = .Text
           ' Else
           '    PaperNum = PaperNum + "," + "'" + .Text + "'"
            'End If
           ' ReDim Preserve strCP09(n)
           'strCP09(n) = .Text
           'n = n + 1
         End If
      Next
   End With
   lbePaperNum.Caption = strCP09
   cp01 = frm071012.txtcp01
   cp02 = frm071012.txtcp02
   cp03 = frm071012.txtcp03
   cp04 = frm071012.txtcp04
   lbeCaseNum = GiveSymbol(frm071012.txtcp01, frm071012.txtcp02, frm071012.txtcp03, frm071012.txtcp04)
   lbeCus = frm071012.lbeCusNum
   lbeCusName = frm071012.lbeCusName
   cboCaseName = frm071012.cboCaseName
   GetStaff
   lbeAccpet = ChangeTStringToTDateString(frm071012.txtAccept)
   ClearText
   ReadCaseprogress '點選進來的該筆收文號資料
   QueryCaseProgress '此案號通知開庭的最大收文日庭期資料
   
   Me.Tag = "" 'Added by Lydia 2022/12/08
   
'   m_AttachPath = App.path & "\" & strUserNum 'Added by Lydia 2015/10/30
End Sub

Private Sub ReadCaseprogress()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   strOpposite = "" '2011/9/15 add by sonia
   
   strSql = "SELECT * FROM CASEPROGRESS WHERE CP09 = '" & lbePaperNum.Caption & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.EOF = False Then
      If Not IsNull(rsTmp.Fields("CP49")) Then
         lbeTitle.Caption = rsTmp.Fields("CP49")
      End If
     ' If Not IsNull(rsTmp.Fields("CP12")) Then
     '    lbeTitle.Caption = rsTmp.Fields("CP12")
     ' End If
      If Not IsNull(rsTmp.Fields("CP13")) Then
         Strsale = rsTmp.Fields("CP13")
      End If
      '2011/9/15 add by sonia
      If "" & rsTmp.Fields("CP40") & rsTmp.Fields("CP41") & rsTmp.Fields("CP42") <> "" Then
         strOpposite = "" & rsTmp.Fields("CP40") & " " & rsTmp.Fields("CP41") & " " & rsTmp.Fields("CP42")
      End If
      '2011/9/15 end
      'Add By Sindy 2020/6/10
      m_CP14 = "" & rsTmp.Fields("CP14") '承辦人
      m_CP29 = "" & rsTmp.Fields("CP29") '協辦人員
      '2020/6/10 END
      'Added by Lydia 2025/11/18 改成下拉選單
      Call PUB_SetGovCmb(Me.cboGov, m_GovListNew, "" & rsTmp.Fields("cp71"))
      m_GovListDef = m_GovListNew
      'end 2025/11/18
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   'Modify By Sindy 2011/6/24
   '最新智權人員
   If m_CP01 = "FCL" Or m_CP01 = "LIN" Then
      Label9 = GetPrjSalesNM(PUB_GetFCLSalesNo(m_CP01, m_CP02, m_CP03, m_CP04))
   Else
      Label9 = GetPrjSalesNM(PUB_GetAKindSalesNo(m_CP01, m_CP02, m_CP03, m_CP04))
   End If
End Sub

Private Function ChgType(i As Integer, strText As String) As String
Dim strTemp As String
   
   Select Case i
      Case 0    '  Change format to 880505
         ChgType = ChangeWStringToTString(strText)
      Case 1
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetCaseProperty(CP01, StrText, strTemp, False) Then ChgType = strTemp
         If ClsPDGetCaseProperty(cp01, strText, strTemp, False) Then ChgType = strTemp
      Case 2
         If strText <> "" Then
            'edit by nickc 2007/02/07 不用 dll 了
            'If objPublicData.GetStaff(StrText, strTemp) Then ChgType = strTemp
            If ClsPDGetStaff(strText, strTemp) Then ChgType = strTemp
         Else
            MsgBox "開庭人員不可為空", vbCritical
            ChgType = ""
         End If
      Case 3     'Change 880501 format to 88/05/01
         ChgType = ChangeTStringToTDateString(strText)
      Case 4     'Change 880501 from to 19990501
         ChgType = ChangeTStringToWString(strText)
      Case 5
         'edit by nickc 2007/02/07 不用 dll 了
         'If objLawDll.GetGovName(StrText, strTemp) Then ChgType = strTemp
         If ClsPDGetGovName(strText, strTemp) Then ChgType = strTemp
   End Select
End Function

Private Function SaveData() As Boolean
Dim strNewNum As String, strName1 As String, strName2 As String
Dim CDT As String, CTM As String, i As Integer
Dim nPos As Integer
Dim strPerson As String
Dim iErr As Integer, sErrMsg As String, bolRemove As Boolean
Dim arrFile1, ii As Integer
   
   'Add By Cheng 2002/11/07
   On Error GoTo ErrorHandler
   SaveData = True
   cnnConnection.BeginTrans
   
   nPos = 0
   nPos = InStr(cboPerson, ",")
   If nPos <> 0 Then
      strPerson = Left(cboPerson, nPos - 1)
   ElseIf cboPerson <> "" Then
      strPerson = cboPerson
   End If
   
   CDT = Format(Date, "YYYYMMDD")
   CTM = Format(time, "HHMM")
   '2009/9/9 MODIFY BY SONIA
   'strSaleArea = GetST15(Strsale)
   strSaleArea = GetSalesArea(IIf(cp01 = "FCL", PUB_GetFCLSalesNo(cp01, cp02, cp03, cp04), IIf(cp01 = "LIN", PUB_GetFCLSalesNo(cp01, cp02, cp03, cp04), PUB_GetAKindSalesNo(cp01, cp02, cp03, cp04))))
   
   'edit by nickc 2007/02/07 不用 dll 了
   'If objPublicData.GetAutoNumber("C", strNewNum, 1, 1) Then
   If ClsPDGetAutoNumber("C", strNewNum, 1, 1) Then
      'Modify By Sindy 2010/8/18 比對自動編號年度
      'strNum = "C" + CStr(Year(Date) - 1911) + strNewNum
      strNum = "C" + CompAutoNumberYear(CStr(Val(Mid(strSrvDate(1), 1, 4)) - 1911)) + strNewNum
   End If
   
   'Modify By Sindy 2011/10/20 +Text(7).開庭別
   'Modified by Lydia 2015/10/30 備註的前面+通知方式
   'Modified by Lydia 2025/11/18 Text(0) 改為 Trim(Left(cboGov.Text, 3))
   blnIsSave = insertdata(strNum, cp01, cp02, cp03, cp04, strSaleArea, Strsale, lbePaperNum, _
      ChangeTStringToWString(Replace(lbeAccpet, "/", "")), Trim(Left(cboGov.Text, 3)), Text(1), Text(2), _
      Text(3), ChangeTStringToWString(Text(4)), Text(5), strPerson, Mergerstring(1), Mergerstring(2), _
      IIf(Text(8).Text = "1", "書面通知;", "口頭告知;") & Text(6), strUserNum, CDT, CTM, strPerson, Text(7))
   
   'Add By Cheng 2002/11/07
   If blnIsSave = False Then GoTo ErrorHandler
   
   'Add By Sindy 2011/10/20 勾選取消前次庭期
   If Check1.Value = 1 Then
      strTemp = Split(m_strCancelRecv, ",")
      For i = 0 To UBound(strTemp) - 1
         strExc(0) = "update courtyardperiod set cdp18=" & strSrvDate(1) & " where cdp01='" & strTemp(i) & "'"
         cnnConnection.Execute strExc(0)
         strExc(0) = "update caseprogress set cp57=" & strSrvDate(1) & ",cp58='99' where cp09='" & strTemp(i) & "'"
         cnnConnection.Execute strExc(0)
      Next i
   End If
   
   'Add By Sindy 2011/6/9
   'Modified by Lydia 2022/12/08 配合輸入出庭費,改成先存暫存檔再寫入正式Table
   'strExc(0) = "delete from caselawer where cl01='" & strNum & "'"
   'cnnConnection.Execute strExc(0)
   'If strPublicTemp <> "" Then
   '   strTemp = Split(strPublicTemp, ",")
   '   For i = 0 To UBound(strTemp) - 1
   '      strExc(0) = "insert into caselawer values('" & strNum & "','" & strTemp(i) & "')"
   '      cnnConnection.Execute strExc(0)
   '   Next i
   'End If
   If Me.Tag <> "" And InStr(Me.Tag, "|") > 0 Then '有點選「出庭律師」
      If PUB_SaveCaseLawer(strNum, Mid(Me.Tag, InStr(Me.Tag, "|") + 1), strPublicTemp, True) = True Then
      End If
   End If
   'end 2022/12/08
   
   'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
   Pub_UpdateFromMaxCP27 cp01, cp02, cp03, cp04
   
   If blnIsSave Then SaveData = True
   
   'Add By Cheng 2002/11/07
   cnnConnection.CommitTrans
Exit Function
   
ErrorHandler:
   cnnConnection.RollbackTrans
   If iErr <> 0 Then
      MsgBox sErrMsg, vbCritical
   End If
   SaveData = False
End Function

'Private Function DisHuman(n As Integer, strData As String) As Boolean
'Dim i As Integer, j As Integer, strTemp() As String, t As Integer, DisData As Variant
'
'   j = 1
'   DisData = Split(strData, ",")
'   Select Case n
'      Case 1
'         For t = 0 To UBound(DisData)
'            Text1(t) = DisData(t)
'         Next
'      Case 2
'         For t = 0 To UBound(DisData)
'            Text2(t) = DisData(t)
'         Next
'   End Select
'End Function

Private Function Mergerstring(t As Integer) As String
Dim i As Integer, strName1 As String, strName2 As String
   
   Select Case t
      Case 1
         For i = 0 To 2
            If Text1(i) <> "" Then
               'If i = 0 Then Mergerstring = Text1(i) Else Mergerstring = Mergerstring + "," + Text1(i)
               Mergerstring = Mergerstring & Text1(i) & ","
            End If
         Next
      Case 2
         For i = 0 To 2
            If Text2(i) <> "" Then
               'If i = 0 Then Mergerstring = Text2(i) Else Mergerstring = Mergerstring + "," + Text2(i)
               Mergerstring = Mergerstring & Text2(i) & ","
            End If
         Next
   End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm071013 = Nothing
   'Add By Sindy 2011/6/9
   strPublicTemp = ""
   Unload frm071018
   '2011/6/9 End
End Sub

Private Sub Text_Change(Index As Integer)
   'If Index = 0 Then If Text(0) = "" Then lbeGov = "" 'Mark by Lydia 2025/11/18
End Sub

Private Sub Text_GotFocus(Index As Integer)
   TextInverse Text(Index)
   Select Case Index
         Case 1, 2, 6
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
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text_LostFocus(Index As Integer)
  'edit by nickc 2007/06/11  切換輸入法改用API
  'Text(Index).IMEMode = 2
  CloseIme
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
Dim strTemp As String
   
   Select Case Index
      'Mark by Lydia 2025/11/18
      'Case 0
      '      lbeGov = ""
      '    If Text(Index) <> "" Then
      '      'edit by nickc 2007/02/07 不用 dll 了
      '      'If objLawDll.GetGovName(Text(Index), strTemp) Then
      '      If ClsPDGetGovName(Text(Index), strTemp) Then
      '         lbeGov = strTemp
      '      Else
      '         Cancel = True
      '      End If
      '    End If
      'end 2025/11/18
      Case 1
         If Text(Index) <> "" Then
            If CheckLengthIsOK(Text(Index), 32) = False Then
                Cancel = True
            End If
         End If
      Case 3 '開庭種類
           If Text(Index) <> "" Then
               'Modify By Sindy 2011/10/20
               'If Not (Text(index) = "1" Or Text(index) = "2" Or Text(index) = "3") Then
               'Modify by Amy 2018/01/24 加5.調解
               If Not (Text(Index) = "1" Or Text(Index) = "2" Or Text(Index) = "3" Or Text(Index) = "4" Or Text(Index) = "5") Then
                  DataErrorMessage 1, "開庭種類"
                  Cancel = True
               End If
           End If
      '開庭日期: 不可空白, 不可小於系統日，檢查日期。
      Case 4 '開庭日期
         '若有輸入開庭日期
         If Text(Index) <> "" Then
            If CheckIsTaiwanDate(Text(Index)) Then
'               If Val(GetTaiwanTodayDate) - Val(Text(index)) > 0 Then
'                  MsgBox "輸入日期小於系統日", vbCritical
'                  Cancel = True
'               End If
            Else
              DataErrorMessage 2, "開庭日期"
              Cancel = True
            End If
         '若未輸入開庭日期
         Else
            DataErrorMessage 5, "開庭日期"
            Cancel = True
         End If
      Case 5
           If Text(Index) <> "" Then
              If Len(Text(Index)) = 4 Then
                If Not ChkTime(Text(Index)) Then Cancel = True
              Else
                 DataErrorMessage 1, "時間"
                 Cancel = True
              End If
           End If
        Case 6
           If Text(Index) <> "" Then
              If CheckLengthIsOK(Text(Index), 2000) = False Then
                Cancel = True
              End If
           End If
        'Add By Sindy 2011/10/20
        Case 7 '開庭別
           If Text(Index) <> "" Then
               'Modify by Amy 2018/01/24 加6.調解庭
               If Not (Text(Index) = "1" Or Text(Index) = "2" Or Text(Index) = "3" Or Text(Index) = "4" Or Text(Index) = "5" Or Text(Index) = "6") Then
                  DataErrorMessage 1, "開庭別"
                  Cancel = True
               End If
           End If
        'Added by Lydia 2015/10/30 通知方式合併備註
        Case 8
           If Text(Index) <> "" Then
               If Not (Text(Index) = "1" Or Text(Index) = "2") Then
                  DataErrorMessage 1, "通知方式"
                  Cancel = True
               End If
           Else
               DataErrorMessage 5, "通知方式"
               Cancel = True
           End If
   End Select
   If Cancel Then TextInverse Text(Index)
End Sub

Private Function AllTextBeforeSaveCheck() As Boolean
Dim tmpBol As Boolean 'Added by Lydia 2025/11/18

   'Modified by Lydia 2025/11/18
   'If Text(0) = "" Then
   '   MsgBox "機關代號不可為空", vbCritical
   '   AllTextBeforeSaveCheck = True
   '   Text(0).SetFocus
   '   Exit Function
   'End If
   If Trim(cboGov.Text) = "" Then
      MsgBox "機關代號不可為空", vbCritical
      AllTextBeforeSaveCheck = True
      Exit Function
   Else
      tmpBol = False
      Call cboGov_Validate(tmpBol)
      If tmpBol = True Then
         AllTextBeforeSaveCheck = True
         Exit Function
      End If
   End If
   'Add By Sindy 2011/10/20
   
   If Text(7) = "" Then
      MsgBox "開庭別不可為空", vbCritical
      Text(7).SetFocus
      AllTextBeforeSaveCheck = True
      Exit Function
   End If
   '2011/10/20 End
   If Text(3) = "" Then
      MsgBox "開庭種類不可為空", vbCritical
      Text(3).SetFocus
      AllTextBeforeSaveCheck = True
      Exit Function
   End If
   'Added By Lydia 2015/10/30 通知方式合併備註
   If Text(8) = "" Then
      MsgBox "通知方式不可為空", vbCritical
      Text(8).SetFocus
      AllTextBeforeSaveCheck = True
      Exit Function
   End If
   'end 2015/10/30
   
   '開庭日期: 不可空白, 不可小於系統日，檢查日期。
   '若未輸入開庭日期
   If Text(4) = "" Then
      MsgBox "開庭日期不可為空", vbCritical
      Text(4).SetFocus
      AllTextBeforeSaveCheck = True
      Exit Function
   'Add By Cheng 2002/03/11
   '若有輸入開庭日期
   Else
      If CheckIsTaiwanDate(Text(4)) Then
'         If Val(GetTaiwanTodayDate) - Val(Text(4)) > 0 Then
'            'Modify By Sindy 2011/7/20
'            'MsgBox "輸入日期小於系統日", vbCritical
'            If MsgBox("開庭日期小於系統日，確定日期是否正確？", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
'               Text(4).SetFocus
'               AllTextBeforeSaveCheck = True
'               Exit Function
'            End If
'         End If
      Else
         DataErrorMessage 2, "開庭日期"
         Text(4).SetFocus
         AllTextBeforeSaveCheck = True
         Exit Function
      End If
   End If
   
   If Text(5) = "" Then
      MsgBox "時間不可為空", vbCritical
      Text(5).SetFocus
      AllTextBeforeSaveCheck = True
      Exit Function
   End If
   
   'Added by Lydia 2021/09/14 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
  
   AllTextBeforeSaveCheck = False
End Function

Private Sub Text1_GotFocus(Index As Integer)
   Select Case Index
      Case 0, 1, 2
          'edit by nickc 2007/06/11  切換輸入法改用API
          'Text1(Index).IMEMode = 1
          OpenIme
   End Select
End Sub

'Modified by Lydia 2021/09/14 改成Form 2.0
'Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
  KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
      Case 0, 1, 2
          'edit by nickc 2007/06/11  切換輸入法改用API
          'Text1(Index).IMEMode = 2
          CloseIme
   End Select
End Sub

Private Sub Text2_GotFocus(Index As Integer)
   Select Case Index
      Case 0, 1, 2
          'edit by nickc 2007/06/11  切換輸入法改用API
          'Text2(Index).IMEMode = 1
          OpenIme
   End Select
End Sub

'Modified by Lydia 2021/09/14 改成Form 2.0
'Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
  KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub ClearText()
'Modified by Lydia 2021/09/14
'Dim txt As TextBox
Dim txt As Control

   For Each txt In frm071013.Text
      txt.Text = ""
   Next
   For Each txt In frm071013.Text1
      txt.Text = ""
   Next
   For Each txt In frm071013.Text2
      txt.Text = ""
   Next
   Check1.Value = 0
   
   'm_strSaveFiles = "" 'Added by Lydia 2015/10/30
   m_strSaveFileType = "" 'Add By Sindy 2016/6/24
   m_strSaveFilesOA = "" 'Add By Sindy 2016/6/24
   m_strSaveFilesBRIEF = "" 'Add By Sindy 2016/6/24
End Sub

Private Sub Text2_LostFocus(Index As Integer)
   Select Case Index
      Case 0, 1, 2
          'edit by nickc 2007/06/11  切換輸入法改用API
          'Text2(Index).IMEMode = 2
          CloseIme
   End Select
End Sub

Private Sub GetNation()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String

    If m_CP01 = "LA" Then
       m_Nation = "000"
    Else
   
         strSql = "SELECT LC15 FROM LAWCASE WHERE LC01 ='" & m_CP01 & "'" & _
                  " AND LC02 ='" & m_CP02 & "' AND LC03 ='" & m_CP03 & "'" & _
                  " AND LC04 ='" & m_CP04 & "'"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.EOF = False Then
            If Not IsNull(rsTmp.Fields("LC15")) Then
               m_Nation = rsTmp.Fields("LC15")
            Else
               m_Nation = ""
            End If
         Else
            m_Nation = ""
         End If
   End If
End Sub

Private Sub GetStaff()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
  
'modify by sonia 2019/2/14 改用共用Function GetLawerList
'   '2009/8/19 modify by sonia 加外法律師
'   'strSQL = "SELECT ST01,ST02 FROM STAFF WHERE ST03 ='L01' AND ST04 = '1' ORDER BY ST01"
'   strSql = "SELECT ST01,ST02 FROM STAFF WHERE (ST03 ='L01' OR ST20='13') AND ST04 = '1' ORDER BY ST01"
'   '2009/8/19 end
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'
'   If rsTmp.EOF = False Then
'      Do While rsTmp.EOF = False
'         If Not IsNull(rsTmp.Fields("ST01")) Then
'            cboPerson.AddItem rsTmp.Fields("ST01") & "," & IIf(IsNull(rsTmp.Fields("ST02")), "", rsTmp.Fields("ST02"))
'         End If
'         rsTmp.MoveNext
'      Loop
'   End If
'   rsTmp.Close
'   Set rsTmp = Nothing
Dim i As Integer, varTmp1 As Variant, strTmp As String

   strSql = GetLawerList("1")
   varTmp1 = Split(strSql, ";")
   For i = 0 To UBound(varTmp1)
      strTmp = varTmp1(i)
      cboPerson.AddItem strTmp
   Next
'end 2019/2/14
End Sub

Private Sub QueryCaseProgress()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   'Add By Sindy 2016/9/7
   strSql = "SELECT cp05 FROM CASEPROGRESS WHERE CP01='" & m_CP01 & "'" & _
            " AND CP02 ='" & m_CP02 & "' AND CP03 ='" & m_CP03 & "'" & _
            " AND CP04 ='" & m_CP04 & "' AND CP10 ='" & 通知開庭 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.Close
   '2016/9/7 END
      strSql = "SELECT MAX(CP09) FROM CASEPROGRESS WHERE CP05 IN(" & _
               "SELECT MAX(CP05) FROM CASEPROGRESS WHERE CP01='" & m_CP01 & "'" & _
               " AND CP02 ='" & m_CP02 & "' AND CP03 ='" & m_CP03 & "'" & _
               " AND CP04 ='" & m_CP04 & "' AND CP10 ='" & 通知開庭 & "')" & _
               " AND CP01 ='" & m_CP01 & "' AND CP02 ='" & m_CP02 & "'" & _
               " AND CP03 ='" & m_CP03 & "' AND CP04 ='" & m_CP04 & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.EOF = False Then
         rsTmp.MoveFirst
         If Not IsNull(rsTmp.Fields(0)) Then
            ReadCourtyardPeriod (rsTmp.Fields(0))
         End If
      End If
   End If
   
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub ReadCourtyardPeriod(strCDP01 As String)
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, i As Integer
   
   '庭期資料檔
   strSql = "SELECT * FROM COURTYARDPERIOD WHERE CDP01='" & strCDP01 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.EOF = False Then
      '開庭人員
      If Not IsNull(rsTmp.Fields("CDP02")) Then
         cboPerson.Text = rsTmp.Fields("CDP02") & "," & ChgType(2, rsTmp.Fields("CDP02"))
         lbeName.Caption = ChgType(2, rsTmp.Fields("CDP02"))
      End If
      '機關代號
      'Modified by Lydia 2025/11/18 改成下拉清單
      'If Not IsNull(rsTmp.Fields("CDP05")) Then
      '   Text(0).Text = rsTmp.Fields("CDP05")
      '   lbeGov = IIf(Text(0) = "", "", ChgType(5, Text(0)))
      'End If
      Call PUB_SetGovCmb(Me.cboGov, m_GovListNew, "" & rsTmp.Fields("cdp05"))
      m_GovListDef = m_GovListNew
      'Add By Sindy 2011/10/20
      '備註
      If Not IsNull(rsTmp.Fields("CDP07")) Then
         Text(6).Text = rsTmp.Fields("CDP07")
         'Added by Lydia 2015/10/30 通知方式合併在備註
         If InStr("書面通知;口頭告知;", Left(Text(6).Text, 5)) > 0 Then
            Text(6).Text = Mid(Text(6).Text, 6)
         End If
         'end 2015/10/30
      End If
      '法官
      If Not IsNull(rsTmp.Fields("CDP08")) Then
         strTemp = Split(rsTmp.Fields("CDP08"), ",")
         For i = 0 To UBound(strTemp) - 1
            Text1(i) = strTemp(i)
         Next i
      End If
      '檢察官
      If Not IsNull(rsTmp.Fields("CDP09")) Then
         strTemp = Split(rsTmp.Fields("CDP09"), ",")
         For i = 0 To UBound(strTemp) - 1
            Text2(i) = strTemp(i)
         Next i
      End If
      '開庭別
      If Not IsNull(rsTmp.Fields("CDP17")) Then
         Text(7).Text = rsTmp.Fields("CDP17")
      End If
      '開庭種類,開庭日期,時間欄位預設為空白,之前已有先清欄位值
      '2011/10/20 End
   End If
      
   'Add By Sindy 2011/10/20 讀取收文資料
   strExc(0) = "select cp30,cp35 from caseprogress where cp09='" + strCDP01 + "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      '股別
      If Not IsNull(RsTemp.Fields("cp30")) Then
         Text(2).Text = RsTemp.Fields("cp30")
      End If
      '法院案號
      If Not IsNull(RsTemp.Fields("cp35")) Then
         Text(1).Text = RsTemp.Fields("cp35")
      End If
   End If
   
   'Add By Sindy 2011/6/9
   '其他出庭律師
   strPublicTemp = ""
   'Modify By Sindy 2011/10/20
   'strExc(0) = "select cl02 from caselawer where cl01='" + lbePaperNum + "' order by cl02 asc "
   strExc(0) = "select cl02 from caselawer where cl01='" + strCDP01 + "' order by cl02 asc "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If IsNull(RsTemp.Fields(0).Value) = False Then
            strPublicTemp = strPublicTemp & RsTemp.Fields(0).Value & ","
         End If
         RsTemp.MoveNext
      Loop
   End If
   m_CL02 = strPublicTemp  'Added by Lydia 2022/12/08
   '2011/6/9 End
End Sub

Private Function insertdata(ByRef CP09 As String, ByRef cp01 As String, ByRef cp02 As String, ByRef cp03 As String, ByRef cp04 As String, _
   ByRef cp12 As String, ByRef cp13 As String, ByRef CP43 As String, ByRef cp05 As String, ByRef cp71 As String, ByRef cp35 As String, ByRef cp30 As String, _
   ByRef cdp06 As String, ByRef cdp03 As String, ByRef cdp04 As String, ByRef cdp02 As String, ByRef cdp08 As String, ByRef cdp09 As String, _
   ByRef CP64 As String, ByRef cID As String, ByRef Cd As String, ByRef Ctime As String, cp14 As String, cdp17 As String) As Boolean
   
Dim cdp05 As String
Dim strSql As String
Dim strSQL1 As String
Dim i As Integer
   
   insertdata = True
   Err.Clear
   On Error Resume Next
   ' 91.04.04 modify by louis (修改單引號)
   'Modify By Cheng 2003/04/07
   '智權人員存最近收文A類接洽記錄單的智權人員
   'Modify By Sindy 2011/7/6 存檔時cp48不要存
   strSql = "INSERT INTO CASEPROGRESS (cp09,cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp12,cp13,cp32,cp43,cp20,cp26," + _
      " cp10,cp71,cp35,cp14,cp64,cp30) values (" + CNULL(CP09) + "," + CNULL(cp01) + "," + CNULL(cp02) + "," + CNULL(cp03) + "," + CNULL(cp04) + _
      "," + CNULL(cp05) + "," + CNULL(cdp03) + "," + CNULL(cdp03) + "," + CNULL(cp12) + "," + IIf(cp01 = "FCL", CNULL(PUB_GetFCLSalesNo(cp01, cp02, cp03, cp04)), IIf(cp01 = "LIN", CNULL(PUB_GetFCLSalesNo(cp01, cp02, cp03, cp04)), CNULL(PUB_GetAKindSalesNo(cp01, cp02, cp03, cp04)))) + _
      ",'N'," + CNULL(CP43) + ",'N','N'," + CNULL(通知開庭) + "," + CNULL(cp71) + "," + CNULL(cp35) + "," + CNULL(cp14) + "," + CNULL(ChgSQL(CP64)) + "," + CNULL(ChgSQL(cp30)) + ")"
   'Modify By Sindy 2011/10/20 +CDP17
   strSQL1 = "insert into courtyardperiod(cdp01,cdp02,cdp03,cdp04,cdp05,cdp06,cdp07,cdp08,cdp09,cdp10,cdp11,cdp12,cdp17) values " + _
      " (" + CNULL(CP09) + "," + CNULL(cdp02) + "," + CNULL(cdp03) + "," + CNULL(cdp04) + "," + CNULL(cp71) + _
      "," + CNULL(cdp06) + "," + CNULL(CP64) + "," + CNULL(cdp08) + "," + CNULL(cdp09) + "," + CNULL(cID) + "," + CNULL(Cd) + _
      "," + CNULL(Ctime) + "," + CNULL(cdp17) + ")"
   
   cnnConnection.Execute strSql
   If Err.Number <> 0 Then
      insertdata = False
      Exit Function
   End If
   cnnConnection.Execute strSQL1
   If Err.Number <> 0 Then
      insertdata = False
      Exit Function
   End If
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
            Text(objTxt.Index).SetFocus
            Exit Function
         End If
      End If
   Next
   
   'Added by Lydia 2022/12/08 修改承辦人檢查; 若已有CaseLawer資料,但是修改承辦人後沒有再進入維護
   If m_CL02 <> "" And InStr(strPublicTemp, Replace(Trim(Left(cboPerson, 6)), ",", "")) = 0 Then
      MsgBox "請進入出庭律師資料輸入作業，檢查出庭律師是否正確！", vbExclamation
      Command4.SetFocus
      Exit Function
   End If
   'end 2022/12/08
   
   TxtValidate = True
End Function

'Added by Lydia 2015/10/30 上傳開庭通知
Private Sub cmdFile_Click()
   Call frm090801_8.SetParent(Me)
   frm090801_8.m_strSaveFiles = Me.m_strSaveFilesOA 'Modify By Sindy 2016/6/24
   frm090801_8.lblCaseNo = m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04
   frm090801_8.Caption = "新增開庭通知附件" 'Add By Sindy 2016/6/24
   Me.m_strSaveFileType = "1" 'Add By Sindy 2016/6/24 開庭通知
   frm090801_8.Show vbModal
End Sub

'Add By Sindy 2016/6/24 上傳開庭紀要
Private Sub cmdBrief_Click()
   Call frm090801_8.SetParent(Me)
   frm090801_8.m_strSaveFiles = Me.m_strSaveFilesBRIEF
   frm090801_8.lblCaseNo = m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04
   frm090801_8.Caption = "新增開庭紀要附件"
   Me.m_strSaveFileType = "2" 'Add By Sindy 2016/6/24 開庭紀要
   frm090801_8.bolNotPDF = True '開啟附件的檔案類型*.*
   frm090801_8.Show vbModal
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

