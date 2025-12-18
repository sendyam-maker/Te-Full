VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010515 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "來電記錄"
   ClientHeight    =   5880
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   9310
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9310
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option2 
      Caption         =   "本所案號："
      Height          =   180
      Left            =   3735
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   795
      Width           =   1230
   End
   Begin VB.OptionButton Option1 
      Caption         =   "申請案號："
      Height          =   180
      Left            =   405
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   795
      Value           =   -1  'True
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7065
      TabIndex        =   5
      Top             =   690
      Width           =   800
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1365
      MaxLength       =   20
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1704
      MaxLength       =   20
      TabIndex        =   0
      Top             =   750
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   7470
      TabIndex        =   14
      Top             =   105
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8310
      TabIndex        =   13
      Top             =   105
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   4
      Left            =   6675
      MaxLength       =   2
      TabIndex        =   4
      Top             =   750
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   3
      Left            =   6435
      MaxLength       =   1
      TabIndex        =   3
      Top             =   750
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   5595
      MaxLength       =   6
      TabIndex        =   2
      Top             =   750
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   5115
      MaxLength       =   3
      TabIndex        =   1
      Top             =   750
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frm04010515.frx":0000
      Height          =   1650
      Left            =   195
      TabIndex        =   29
      Top             =   4170
      Width           =   8925
      _ExtentX        =   15752
      _ExtentY        =   2893
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
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
      _Band(0).Cols   =   18
   End
   Begin MSForms.TextBox Text2 
      Height          =   285
      Left            =   1365
      TabIndex        =   10
      Top             =   2775
      Width           =   885
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "1561;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text8 
      Height          =   285
      Left            =   1365
      TabIndex        =   9
      Top             =   2460
      Width           =   885
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "1561;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text6 
      Height          =   285
      Left            =   5175
      TabIndex        =   8
      Top             =   2160
      Width           =   1470
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "2593;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text9 
      Height          =   810
      Left            =   195
      TabIndex        =   11
      Top             =   3330
      Width           =   8910
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "15716;1429"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1365
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1500
      Width           =   7755
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "13679;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "來電內容："
      Height          =   180
      Index           =   0
      Left            =   225
      TabIndex        =   33
      Top             =   3120
      Width           =   900
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   5
      Left            =   2280
      TabIndex        =   32
      Top             =   2805
      Width           =   1470
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2593;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Shape Shape1 
      Height          =   585
      Left            =   135
      Top             =   570
      Width           =   8880
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   2
      Left            =   1365
      TabIndex        =   28
      Top             =   1830
      Width           =   7755
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "13679;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   4
      Left            =   5175
      TabIndex        =   27
      Top             =   2490
      Visible         =   0   'False
      Width           =   1470
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2593;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   26
      Top             =   2490
      Width           =   1470
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2593;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "存檔後電話連絡單進入歷程作業（輸入資料後，按下確定鍵，操作聯絡歷程送出）。"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   530
      Index           =   11
      Left            =   180
      TabIndex        =   25
      Top             =   60
      Width           =   5010
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "副本收受者："
      Height          =   180
      Index           =   10
      Left            =   225
      TabIndex        =   24
      Top             =   2805
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "最後工程師："
      Height          =   180
      Index           =   9
      Left            =   4005
      TabIndex        =   23
      Top             =   2490
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "處理人員："
      Height          =   180
      Index           =   8
      Left            =   225
      TabIndex        =   22
      Top             =   2490
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "分機號碼："
      Height          =   180
      Index           =   7
      Left            =   225
      TabIndex        =   21
      Top             =   2220
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "來電人員："
      Height          =   180
      Index           =   6
      Left            =   4005
      TabIndex        =   20
      Top             =   2205
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Index           =   5
      Left            =   225
      TabIndex        =   19
      Top             =   1830
      Width           =   720
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   1
      Left            =   5175
      TabIndex        =   18
      Top             =   1200
      Width           =   1470
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2593;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "分所案號："
      Height          =   180
      Index           =   3
      Left            =   4005
      TabIndex        =   17
      Top             =   1200
      Width           =   900
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   0
      Left            =   1365
      TabIndex        =   16
      Top             =   1200
      Width           =   1470
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2593;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   2
      Left            =   225
      TabIndex        =   15
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   4
      Left            =   225
      TabIndex        =   12
      Top             =   1560
      Width           =   900
   End
End
Attribute VB_Name = "frm04010515"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/25 Form2.0已修改
'Memo By Morgan 2012/12/13 智權人員欄已修改
'Created by Morgan 2012/4/16
Option Explicit

Dim m_CP13 As String
Dim m_CaseName As String
Dim m_CP36 As String
Dim m_bolFMP As Boolean, m_PA150 As String  'Added by Lydia 2024/05/09


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         cmdOK(0).Enabled = False
         If TxtValidate = True Then
            'Add by Sindy 2021/11/25 檢查畫面上的物件是否含有Unicode文字
            If PUB_ChkUniText(Me, True, True) = False Then
               Exit Sub
            End If

            If FormSave = True Then
               'Modify By Sindy 2025/4/7
'               PUB_SendMailCache
'               FormClear
'               If Option1.Value Then
'                  Text5.SetFocus
'               Else
'                  Text1(1).SetFocus
'               End If
               '2025/4/7 END
            Else
               cmdOK(0).Enabled = True
            End If
         Else
            cmdOK(0).Enabled = True
         End If
      Case 2
         Unload Me
   End Select
End Sub

Private Sub Command1_Click()
   If GetCaseData = False Then
      If Option1.Value Then
         Text5.SetFocus
         Text5_GotFocus
      Else
         Text1(1).SetFocus
         Text1_GotFocus 1
      End If
   Else
      cmdOK(0).Enabled = True
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   FormClear True
   
   'Add By Sindy 2025/5/5
   If Left(Pub_StrUserSt03, 1) = "P" Then
      Label1(11).Caption = "存檔後會自動發郵件給處理人員。"
   Else
      Label1(11).Caption = "存檔後電話連絡單進入歷程作業（輸入資料後，按下確定鍵，操作聯絡歷程送出）。"
   End If
   '2025/5/5 END
End Sub

'Added by Morgan 2013/5/30 游經理跟秀玲反應不知道按了什麼畫面被關掉了
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If MsgBox("是否確定要結束？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Cancel = 1
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010515 = Nothing
End Sub

Private Sub Option1_Click()
   Text1(1).Enabled = False
   Text1(2).Enabled = False
   Text1(3).Enabled = False
   Text1(4).Enabled = False
   Text5.Enabled = True
   Text5.SetFocus
End Sub

Private Sub Option2_Click()
   Text1(1).Enabled = True
   Text1(2).Enabled = True
   Text1(3).Enabled = True
   Text1(4).Enabled = True
   Text5.Enabled = False
   Text1(1).SetFocus
End Sub

Private Sub Text1_Change(Index As Integer)
   If Text1(Index).Tag <> "" Then
      FormClear False
   End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Change()
   Label2(5) = ""
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
   CloseIme
End Sub

Private Sub Text2_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'Modified by Morgan 2014/8/8 修正輸入中文轉員工編號後中文名稱不會顯示問題
Private Sub Text2_Validate(Cancel As Boolean)
   Text2 = Trim(Text2) 'Added by Morgan 2015/4/21
   If Text2 <> "" And Text2.Tag <> Text2 Then
      'Modified by Morgan 2017/10/25 不必抓第1碼否則6字頭的員工號會被剔除
      'If Left(Text2, 1) > "6" And Left(Text2, 1) < "F" Then
      If Text2 > "6" And Text2 < "F" Then
         If ClsPDGetStaff(Text2, strExc(1)) Then
            Label2(5) = strExc(1)
         Else
            Cancel = True
            Text2_GotFocus
         End If
      Else
         If GetIdFromName(Text2, strExc(1)) Then
            strExc(0) = Text2
            Text2 = strExc(1)
            Label2(5) = strExc(0)
         Else
            Cancel = True
            Text2_GotFocus
         End If
      End If
      Text2.Tag = Text2
   End If
End Sub

Private Sub Text5_Change()
   If Text5.Tag <> "" Then
      FormClear False
   End If
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
   CloseIme
End Sub

Private Function FormSave() As Boolean
   Dim strCP09 As String, strCP12 As String
   Dim strCP64 As String, strDate As String
   Dim strReceiver As String, strCC As String
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd

   
   strCP09 = AutoNo("B", 6)
   strDate = CompWorkDay(3, strSrvDate(1))
   strCP12 = GetSalesArea(m_CP13)
   strCP64 = "來電人員：" & Text6 & ", 分機號碼：" & Text7 & ", 來電內容：" & Text9
   
   If Text2 <> "" Then
      strCC = Text2 & ";"
   End If
   
   'ENGINEERPROGRESS_BEFORE5(Trigger)+排除 945 條件 (新增B類收文自動上齊備日時觸發,會清除或計算承辦期限)
   'Modified by Morgan 2012/6/29 不要掛法限--郭
   strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06," & _
      "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP36,cp48,cp64) VALUES " & _
      "('" & Text1(1) & "','" & Text1(2) & "','" & Text1(3) & "','" & Text1(4) & "'," & strSrvDate(1) & "," & strDate & _
      ",'" & strCP09 & "','945','90'," & CNULL(strCP12) & "," & CNULL(m_CP13) & _
      ",'" & Text8 & "','N','N','N','" & ChgSQL(m_CP36) & "'," & strDate & ",'" & ChgSQL(strCP64) & "') "
   cnnConnection.Execute strSql, intI
   
   strReceiver = Text8
   'Modify By Sindy 2025/4/23 電話連絡單進入歷程作業,不用再CC程序人員及程序主管和工程師主管
'   'FCP 寄給承辦人,程序管制人;副本給承辦人組別主管,程序主管(抓未分案管制人)
'   If Text1(1) = "FCP" Then
'      strExc(0) = "select na16 from patent,fagent,nation where pa01='" & Text1(1) & "' and pa02='" & Text1(2) & "' and pa03='" & Text1(3) & "' and pa04='" & Text1(4) & "'" & _
'                  " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10 and na16<>'" & Text8 & "'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If InStr(strReceiver & strCC, RsTemp(0)) = 0 Then
'            'Modify By Sindy 2025/4/7 歷程收受者只能一位
'            'strReceiver = strReceiver & ";" & RsTemp(0)
'            strCC = strCC & RsTemp(0) & ";"
'            '2025/4/7 END
'         End If
'         'add by sonia 2016/7/15 原發程序主管(抓未分案管制人),改發程序之第二級主管,改寫在此處,下面取消
'         strExc(0) = "select st52 from staff where st01='" & RsTemp(0) & "' and st52 is not null and st52<>'" & Text8 & "'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            If InStr(strReceiver & strCC, RsTemp(0)) = 0 Then
'               strCC = strCC & RsTemp(0) & ";"
'            End If
'         End If
'         'end 2016/7/15
'      End If
'
'      'modify by sonia 2016/7/15 原發程序主管(抓未分案管制人),改發程序之第二級主管,改寫在上面,此處取消
'      'strExc(0) = "select oMan from staff,SetSpecMan where st01='" & Text8 & "' and OCODE=decode(st16,'1','T','2','R','3','S','4','T1')" & _
'      '            " union select oMan from SetSpecMan where OCODE='N'"
''      strExc(0) = "select oMan from staff,SetSpecMan where st01='" & Text8 & "' and OCODE=decode(st16,'1','T','2','R','3','S','4','T1')"
''      intI = 1
''      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''      If intI = 1 Then
''         Do While Not RsTemp.EOF
''            If InStr(strReceiver & strCC, RsTemp(0)) = 0 Then
''               strCC = strCC & RsTemp(0) & ";"
''            End If
''            RsTemp.MoveNext
''         Loop
''      End If
'      'Modify By Sindy 2020/3/27 日文組主管機關來電處理,Email cc對象改為第二,三級主管
'      strExc(10) = PUB_GetFCPEngSup(Text8.Text)
'      If strExc(10) <> "" Then
'         If InStr(strReceiver & strCC, strExc(10)) = 0 Then
'            strCC = strCC & strExc(10) & ";"
'         End If
'      End If
'      '2020/3/27 END
'   End If
   
   'Added by Lydia 2024/05/09 +加註【機械設計組】
   strExc(1) = IIf((Text1(1) = "FCP" Or m_bolFMP = True) And m_PA150 = "4", "【機械設計組】", "") & Text1(1) & "-" & Text1(2) & IIf(Text1(3) & Text1(4) = "000", "", "-" & Text1(3) & "-" & Text1(4)) & " (" & strCP09 & ") 電話聯絡單!!"
   'Modified by Lydia 2017/07/27 改變文字大小
   'strExc(2) = GetMailText
   strExc(2) = GetMailTextNew
   'Added by Lydia 2015/11/05 P案-內文和主旨中的電話聯絡單改為來電聯絡單
   If Text1(1) = "P" Then
        strExc(1) = Replace(strExc(1), "電話聯絡單", "來電聯絡單")
        strExc(2) = Replace(strExc(2), "電話聯絡單", "來電聯絡單")
   End If
   'end 2015/11/05
   strExc(3) = PUB_LeftB(strExc(2), 4000)
   strExc(4) = Mid(strExc(2), Len(strExc(3)) + 1)

   'Add By Sindy 2025/5/5 FCP改用聯絡歷程
   If Left(Pub_StrUserSt03, 1) = "P" Then
   '2025/5/5 END
      strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc10,mc11)" & _
         " values ('" & strUserNum & "','" & strReceiver & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
         ",'" & ChgSQL(strExc(1)) & "','" & ChgSQL(strExc(3)) & "','" & strCC & "','Y','" & ChgSQL(strExc(4)) & "')"
      cnnConnection.Execute strSql, intI
   End If
   
   cnnConnection.CommitTrans
   FormSave = True
   
   PUB_SendMailCache
   FormClear
   If Option1.Value Then
      Text5.SetFocus
   Else
      Text1(1).SetFocus
   End If
   
   'Add By Sindy 2025/4/7 聯絡承辦歷程
   If Left(Pub_StrUserSt03, 1) = "F" Then
      Screen.MousePointer = vbHourglass
      frm090202_2.Hide
      frm090202_2.m_EEP01 = strCP09 '總收文號
      frm090202_2.intReceiveKind = 99 '聯絡
      frm090202_2.SetParent Me
      'frm090202_2.Caption = frm090202_2.Caption ' & " - " & GrdDataList.TextMatrix(i, 3)
   '   frm090202_2.cmdOK(0).Visible = False
   '   frm090202_2.cmdOK(1).Visible = False
      If frm090202_2.QueryData = True Then
         frm090202_2.ShowNextData = True
         frm090202_2.cmdAdd_Click '自動啟動聯絡
         frm090202_2.cmdCancel.Enabled = False '以免人員誤按取消
         frm090202_2.cmdExit.Enabled = False '以免人員誤按結束
         frm090202_2.CboEEP05 = strReceiver '收受者
         frm090202_2.CboEEP05_Validate False
         If strCC <> "" Then strCC = Replace(Mid(strCC, 1, Len(strCC) - 1), ";", ",")
         frm090202_2.txtEEP10_2 = strCC '副本
         frm090202_2.txtEEP10_2_Validate False
         frm090202_2.txtEEP08 = ChgSQL(strCP64) & vbCrLf & "*若需另外收文告代，請勾選【需收文告代】" '內容
         frm090202_2.Show
      End If
      frm090202_2.cmd1(0).Visible = False
      frm090202_2.cmd1(1).Visible = False '相似案
      frm090202_2.cmdOK(2).Visible = False '接洽單
      frm090202_2.cmdOutlook.Visible = False
      Screen.MousePointer = vbDefault
   End If
   '2025/4/7 END
   
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
   
End Function

Private Function TxtValidate() As Boolean
  
'Removed by Morgan 2012/6/18 改讀取資料時控制
'   If Option2.Value And Text5 = "" Then
'      MsgBox "請先輸入正確的本所案號後按查詢！", vbExclamation
'      Text1(1).SetFocus
'      Text1_GotFocus 1
'      Exit Function
'
'   ElseIf Option1.Value And Text1(1) = "" Then
'      MsgBox "請先輸入正確的申請案號後按查詢！", vbExclamation
'      Text5.SetFocus
'      Text5_GotFocus
'      Exit Function
'   End If
'end 2012/6/18

   If Text7 = "" Then
      MsgBox "請輸入分機號碼!", vbInformation
      Text7.SetFocus
      Exit Function
   ElseIf Text6 = "" Then
      MsgBox "請輸入來電人員!", vbInformation
      Text6.SetFocus
      Exit Function
   End If
   If Text9 = "" Then
      MsgBox "請輸入來電內容!", vbInformation
      Text9.SetFocus
      Exit Function
   End If
   If Text8 = "" Then
      MsgBox "請輸入處理人員!", vbInformation
      Text8.SetFocus
      Exit Function
   ElseIf Label2(3) = "" Then
      MsgBox "處理人員輸入錯誤!", vbExclamation
      Text8.SetFocus
      Exit Function
   End If
   
   'Added by Morgan 2015/4/21
   If Text2 <> "" Then
      If Label2(5) = "" Then
         MsgBox "副本收受者輸入錯誤!", vbExclamation
         Text2.SetFocus
         Exit Function
      End If
   End If
   'end 2015/4/21
   
   'Added by Lydia 2024/05/09
   If Option2.Value = True Then
      If Text5.Tag <> Text5.Text Then
         MsgBox "請先執行尋找！", vbExclamation
         Exit Function
      End If
   Else
      If Text1(1) & Text1(2) & Text1(3) & Text1(4) <> Text1(1).Tag & Text1(2).Tag & Text1(3).Tag & Text1(4).Tag Then
         MsgBox "請先執行尋找！", vbExclamation
         Exit Function
      End If
   End If
   'end 2024/05/09
   
   TxtValidate = True
End Function
Private Function GetCaseData() As Boolean
   Dim bolFound As Boolean
   Dim ii As Integer
   
   Label1(9).Visible = False
   Label2(4).Visible = False
   m_CaseName = ""
   m_CP36 = ""
   m_PA150 = "": m_bolFMP = False 'Added by Lydia 2024/05/09
   
   'Modified by Morgan 2012/5/8 +可用本所案號尋找
   If Option1.Value Then
      'Modified by Lydia 2025/05/09 +pa150
      strExc(0) = "select pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa11,pa47,nvl(cu04,nvl(cu06,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90))) CuName,pa150 " & _
         " from patent,customer where pa11='" & ChgSQL(Text5) & "' and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)"
   Else
      If Text1(3) = "" Then Text1(3) = "0"
      If Text1(4) = "" Then Text1(4) = "00"
      'Modified by Lydia 2025/05/09 +pa150
      strExc(0) = "select pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa11,pa47,nvl(cu04,nvl(cu06,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90))) CuName,pa150 " & _
         " from patent,customer where pa01='" & Text1(1) & "' and pa02='" & Text1(2) & "' and pa03='" & Text1(3) & "' and pa04='" & Text1(4) & "' and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      bolFound = True
   ElseIf intI = 0 And Option1.Value Then
      '抓對造號數
      strExc(0) = "select pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa10,pa47,nvl(cu04,nvl(cu06,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90))) CuName" & _
         " from caseprogress,patent,customer where cp36='" & ChgSQL(Text5) & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_CP36 = Text5
         bolFound = True
      End If
   End If
   
   If Not bolFound Then
      MsgBox "申請案號輸入錯誤！", vbExclamation
   
   ElseIf Pub_StrUserSt03 = "F22" And RsTemp.Fields("pa01") = "P" Then
      MsgBox "不可輸入 P 案！", vbExclamation
      
   Else
      With RsTemp
      
         'Added by Morgan 2012/5/8
         If Option2.Value Then Text5 = "" & .Fields("pa11")
         'end 2012/5/8
         Text1(1) = "" & .Fields("pa01")
         Text1(1).Tag = Text1(1)
         Text1(2) = "" & .Fields("pa02")
         Text1(2).Tag = Text1(2)
         Text1(3) = "" & .Fields("pa03")
         Text1(3).Tag = Text1(3)
         Text1(4) = "" & .Fields("pa04")
         Text1(4).Tag = Text1(4)
         Text5.Tag = Text5
         
         If Not IsNull(.Fields("pa05")) Then
            m_CaseName = .Fields("pa05")
         ElseIf Not IsNull(.Fields("pa06")) Then
            m_CaseName = .Fields("pa06")
         ElseIf Not IsNull(.Fields("pa07")) Then
            m_CaseName = .Fields("pa07")
         End If
         AddCboName Combo1, "" & .Fields("pa05"), "" & .Fields("pa06"), "" & .Fields("pa07")
         m_CP13 = PUB_GetAKindSalesNo(Text1(1).Text, Text1(2).Text, Text1(3).Text, Text1(4).Text)
         Label2(0) = GetPrjSalesNM(m_CP13)
         Label2(1) = "" & .Fields("pa47")
         Label2(2) = "" & .Fields("CuName")
         
         'Added by Lydia 2024/05/09
         m_PA150 = "" & .Fields("pa150")
         If Text1(1) = "P" Then
            If PUB_ChkIsFMP(Text1(1), Text1(2), Text1(3), Text1(4)) = True Then
               m_bolFMP = True
            End If
         End If
         'end 2024/05/09
      End With
      
      '處理人員:
      '預設該案號最後的工程師(部門為 F21、F81、P10、P11)；若最後工程師為離職人員則不預設由人工輸入；
      '可修改不可空白且必須為在職人員
      'Modified by Lydia 2019/08/02 排除F4102 (FCP年費不續辦)
      'Modified by Lydia 2024/05/09 +cp09,cp10
      strExc(0) = "select cp14,st04,st02,cp09,cp10 from caseprogress,staff where cp01='" & Text1(1) & "' and cp02='" & Text1(2) & "' and cp03='" & Text1(3) & "' and cp04='" & Text1(4) & "' " & _
                       "and st01(+)=cp14 and cp14<>'F4102' and cp57 is null and st03 in ('F21','F81','P10','P11') order by cp05 desc,cp09 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '最後工程師:
         If RsTemp("st04") = "1" Then
            Text8 = RsTemp("cp14")
            Label2(3) = RsTemp("st02")
            Label1(9).Visible = True
            Label2(4).Visible = True
            Label2(4) = RsTemp("st02")
         Else
            'P案離職才顯示,FCP案都要顯示
            If Text1(1) = "FCP" Then
               Label1(9).Visible = True
               Label2(4).Visible = True
               Label2(4) = RsTemp("st02")
            End If
         End If
         'Added by Lydia 2024/05/09 處理人員若抓最後一道工程師為已離職，預設為主管
         If Text1(1) = "FCP" Or m_bolFMP = True Then
            Label1(9).Visible = True
            Label2(4).Visible = True
            If "" & RsTemp("st04") <> "1" Then
               strExc(0) = PUB_GetFCPPromoterNo("" & RsTemp.Fields("cp09"), "" & RsTemp.Fields("cp10"))
               If strExc(0) <> "" Then
                  Text8 = strExc(0)
                  Label2(3) = GetStaffName(strExc(0))
               End If
            End If
         End If
         'end 2024/05/09
      End If
      
      MSHFlexGrid1.Visible = False
      'MODIFY BY SONIA 2014/9/3 所有排序條件加DESC,改同共同查詢(王副總)
      strExc(0) = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10) as 案件性質" & _
         ",CP43 as 相關總收文號,NVL(S1.ST02,CP14) as 承辦人,NVL(S2.ST02,CP13) as 智權人員" & _
         ",SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日" & _
         " from caseprogress,PATENT,casepropertymap,staff s1,staff s2" & _
         " WHERE cp01='" & Text1(1) & "' and cp02='" & Text1(2) & "' and cp03='" & Text1(3) & "' and cp04='" & Text1(4) & "'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and cp01=cpm01(+) and cp10=cpm02(+) and s1.st01(+)=cp14 and s2.st01(+)=cp13" & _
         " ORDER BY SQLDatet2(CP05) DESC,CP66 DESC,CP67 DESC,DECODE(SUBSTR(CP09,1,1),'A','1','C','2','3'),CP09 DESC"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
      GridHead
      If intI = 1 Then
         With MSHFlexGrid1
         For ii = 1 To .Rows - 1
            If .TextMatrix(ii, 4) <> "" Then
                .TextMatrix(ii, 3) = .TextMatrix(ii, 3) & PUB_GetRelateCasePropertyName(.TextMatrix(ii, 2), "1")
            End If
         Next
         
         End With
      End If
      MSHFlexGrid1.Visible = True
      
      GetCaseData = True
   End If
End Function

Private Sub FormClear(Optional pbolAll As Boolean)
   Dim oLabel As Object
   If pbolAll Then
      Text1(1) = ""
      Text1(2) = ""
      Text1(3) = ""
      Text1(4) = ""
      Text5 = ""
   Else
      If Option1.Value Then
         Text1(1) = ""
         Text1(2) = ""
         Text1(3) = ""
         Text1(4) = ""
      Else
         Text5 = ""
      End If
   End If
   Text1(1).Tag = ""
   Text1(2).Tag = ""
   Text1(3).Tag = ""
   Text1(4).Tag = ""
   Text2 = ""
   Text2.Tag = ""
   Text5.Tag = ""
   Text6 = ""
   Text7 = ""
   Text8 = ""
   Text8.Tag = ""
   Text9 = ""
   Combo1.Clear
   For Each oLabel In Label2
      oLabel.Caption = ""
   Next
   MSHFlexGrid1.Clear
   MSHFlexGrid1.Rows = 2
   GridHead
   Me.cmdOK(0).Enabled = False
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
'   If Text5 <> "" Then
'      Command1_Click
'   End If
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
   OpenIme
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
   CloseIme
End Sub

Private Sub Text8_Change()
   Label2(3) = ""
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
   CloseIme
End Sub

Private Sub Text8_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'Modified by Morgan 2014/8/8 修正輸入中文轉員工編號後中文名稱不會顯示問題
Private Sub Text8_Validate(Cancel As Boolean)
   Text8 = Trim(Text8) 'Added by Morgan 2015/4/21
   If Text8 <> "" And Text8.Tag <> Text8 Then
      If Text8 > "6" And Text8 < "F" Then
         If ClsPDGetStaff(Text8, strExc(1)) Then
            Label2(3) = strExc(1)
         Else
            Cancel = True
            Text8_GotFocus
         End If
      Else
         If GetIdFromName(Text8, strExc(1)) Then
            strExc(0) = Text8
            Text8 = strExc(1)
            Label2(3) = strExc(0)
         Else
            Cancel = True
            Text8_GotFocus
         End If
      End If
      Text8.Tag = Text8
   End If
End Sub

Private Function GetIdFromName(ByVal pName As String, ByRef pID As String) As Boolean
   strExc(0) = "select st01,st02 from staff where st02='" & ChgSQL(pName) & "' and st04='1' and st01>'6' and st01<'F'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.RecordCount = 1 Then
         pID = RsTemp.Fields("st01")
         GetIdFromName = True
      Else
         MsgBox "員工名稱重複，請直接輸入員工編號！"
      End If
   Else
      MsgBox "該員工名稱不存在！"
   End If
End Function

Private Sub Text9_GotFocus()
   TextInverse Text9
   OpenIme
End Sub

Private Function GetMailText() As String
   Dim strText As String
   '要有 &nbsp; 字串空白才不會被再轉換一次
   strText = ""
   If Text1(1) = "FCP" Then
      strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">請程序人員協助調卷！</FONT><BR><BR>"
   End If
   strText = strText & "<TABLE BORDER CELLSPACING=2 CELLPADDING=2 WIDTH=600 STYLE=""border:2px solid;"">"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" COLSPAN=4 HEIGHT=37>"
   'Added by Lydia 2015/11/05 P案-內文和主旨中的電話聯絡單改為來電聯絡單
   If Text1(1) = "P" Then
      strText = strText & "<FONT FACE=""標楷體"" SIZE=5><P ALIGN=""CENTER"">主&nbsp;管&nbsp;機&nbsp;關&nbsp;處&nbsp;理&nbsp;記&nbsp;錄&nbsp;單(來電聯絡單)"
   Else
      strText = strText & "<FONT FACE=""標楷體"" SIZE=5><P ALIGN=""CENTER"">主&nbsp;管&nbsp;機&nbsp;關&nbsp;來&nbsp;電&nbsp;處&nbsp;理&nbsp;記&nbsp;錄&nbsp;單(電話聯絡單)"
   End If
   'end 2015/11/05
   
   strText = strText & "</FONT></TD></TR>"
   strText = strText & "<TR><TD WIDTH=""15%"" VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">日　期</FONT></TD>"
   strText = strText & "<TD WIDTH=""30%"" VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & ChangeTStringToTDateString(strSrvDate(2)) & "</FONT></TD>"
   strText = strText & "<TD WIDTH=""15%"" VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">接話人員</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & strUserName & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">申請案號</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Text5 & "</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">申請人</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Label2(2) & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">本所案號</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Text1(1) & "-" & Text1(2) & IIf(Text1(3) & Text1(4) = "000", "", "-" & Text1(3) & "-" & Text1(4)) & "</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P>分所案號</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & IIf(Label2(1) = "", "　", Label2(1)) & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">案件名稱</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"" COLSPAN=3>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & m_CaseName & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">來電人員</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Text6 & "</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">分機號碼</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Text7 & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">處理人員</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Label2(3) & "</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">智權人員</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Label2(0) & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""TOP"" COLSPAN=4 HEIGHT=200>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P>來電內容：<BR>" & Text9 & "</FONT></TD></TR>"
   strText = strText & "</TABLE>"
   
   'Removed by Morgan 2020/3/30 取消公司名稱
   'strText = strText & "<TABLE WIDTH=600>"
   'strText = strText & "<TR border=0><TD VALIGN=""TOP"" HEIGHT=21>"
   'strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""RIGHT"">台一國際專利法律事務所</FONT></TD></TR>"
   'strText = strText & "</TABLE>"
   'end 2020/3/30
   
   GetMailText = strText
End Function

Private Sub GridHead()
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .Cols = 11
      .row = 0
      .col = 0: .Text = "V"
      .ColWidth(0) = 0 '180
      .col = 1: .Text = "收文日"
      .ColWidth(1) = 800 '788
      .CellAlignment = flexAlignRightCenter
      .col = 2: .Text = "總收文號"
      .ColWidth(2) = 1000 '938
      .CellAlignment = flexAlignLeftCenter
      .col = 3: .Text = "案件性質"
      .ColWidth(3) = 2000 ' 950
      .CellAlignment = flexAlignLeftCenter
      .col = 4: .Text = "相關收文號"
      .ColWidth(4) = 0
      .CellAlignment = flexAlignLeftCenter
      .col = 5: .Text = "承辦人"
      .ColWidth(5) = 650 ' 593
      .CellAlignment = flexAlignLeftCenter
      .col = 6: .Text = "智權人員"
      .ColWidth(6) = 650 ' 593
      .CellAlignment = flexAlignLeftCenter
      .col = 7: .Text = "本所期限"
      .ColWidth(7) = 820 '788
      .CellAlignment = flexAlignRightCenter
      .col = 8: .Text = "法定期限"
      .ColWidth(8) = 820 '788
      .CellAlignment = flexAlignRightCenter
      .col = 9: .Text = "發文日"
      .ColWidth(9) = 800 '788
      .CellAlignment = flexAlignRightCenter
      .col = 10: .Text = "取消收文日"
      .ColWidth(10) = 1000 '788
      .CellAlignment = flexAlignLeftCenter
      .Visible = True
   End With
End Sub

'Added by Lydia 2017/07/26 因為W7的解析度不同,改變字體
Private Function GetMailTextNew() As String
   Dim strText As String
   '要有 &nbsp; 字串空白才不會被再轉換一次
   strText = ""
   If Text1(1) = "FCP" Then
      strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">請程序人員協助調卷！</FONT><BR><BR>"
   End If
   strText = strText & "<TABLE BORDER CELLSPACING=2 CELLPADDING=2 WIDTH=620px STYLE=""border:2px solid;"">"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" COLSPAN=4 HEIGHT=37><span style=""font-size:22px"">"
   If Text1(1) = "P" Then
      strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">主&nbsp;管&nbsp;機&nbsp;關&nbsp;處&nbsp;理&nbsp;記&nbsp;錄&nbsp;單(來電聯絡單)"
   Else
      'W7時字太長,折行
      strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">主&nbsp;管&nbsp;機&nbsp;關&nbsp;來&nbsp;電&nbsp;處&nbsp;理&nbsp;記&nbsp;錄&nbsp;單(電話聯絡單)"
   End If
   
   strText = strText & "</FONT></span></TD></TR>"
   strText = strText & "<TR><TD WIDTH=""15%"" VALIGN=""MIDDLE"" HEIGHT=30><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">日　期</FONT></span></TD>"
   strText = strText & "<TD WIDTH=""30%"" VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & ChangeTStringToTDateString(strSrvDate(2)) & "</FONT></span></TD>"
   strText = strText & "<TD WIDTH=""15%"" VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">接話人員</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & strUserName & "</FONT></span></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">申請案號</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & Text5 & "</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">申請人</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & Label2(2) & "</FONT></span></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">本所案號</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & Text1(1) & "-" & Text1(2) & IIf(Text1(3) & Text1(4) = "000", "", "-" & Text1(3) & "-" & Text1(4)) & "</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">分所案號</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & IIf(Label2(1) = "", "　", Label2(1)) & "</FONT></span></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">案件名稱</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"" COLSPAN=3><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & m_CaseName & "</FONT></span></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">來電人員</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & Text6 & "</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">分機號碼</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & Text7 & "</FONT></span></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">處理人員</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & Label2(3) & "</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""CENTER"">智權人員</FONT></span></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE""><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""LEFT"">" & Label2(0) & "</FONT></span></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""TOP"" COLSPAN=4 HEIGHT=200><span style=""font-size:18px"">"
   strText = strText & "<FONT FACE=""標楷體""><P>來電內容：<BR>" & Text9 & "</FONT></span></TD></TR>"
   strText = strText & "</TABLE>"
   
   'Removed by Morgan 2020/3/30 取消公司名稱
   'strText = strText & "<TABLE WIDTH=620px>"
   'strText = strText & "<TR border=0><TD VALIGN=""TOP"" HEIGHT=21><span style=""font-size:18px"">"
   'strText = strText & "<FONT FACE=""標楷體""><P ALIGN=""RIGHT"">台一國際專利法律事務所</FONT></span></TD></TR>"
   'strText = strText & "</TABLE>"
   'end 2020/3/30
   
   GetMailTextNew = strText
End Function
