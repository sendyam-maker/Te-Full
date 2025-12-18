VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010603_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "核駁函輸入"
   ClientHeight    =   5760
   ClientLeft      =   -1440
   ClientTop       =   1176
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9336
   Begin VB.TextBox txtDelivery 
      Height          =   270
      Left            =   6990
      MaxLength       =   7
      TabIndex        =   62
      Top             =   3420
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      Height          =   270
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   4
      Top             =   2550
      Width           =   705
   End
   Begin VB.ComboBox cboOrg 
      Height          =   300
      Left            =   7350
      Style           =   2  '單純下拉式
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1935
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   4515
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   54
      Top             =   1935
      Width           =   255
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2250
      Width           =   7932
   End
   Begin VB.TextBox Text14 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   6990
      MaxLength       =   7
      TabIndex        =   16
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   4920
      TabIndex        =   41
      Top             =   570
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   40
      Top             =   570
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   39
      Top             =   570
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   38
      Top             =   570
      Width           =   255
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   37
      Top             =   570
      Width           =   375
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm06010603_3.frx":0000
      Left            =   1200
      List            =   "frm06010603_3.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   36
      Top             =   870
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7140
      TabIndex        =   22
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6312
      TabIndex        =   21
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8364
      TabIndex        =   23
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox Text18 
      Enabled         =   0   'False
      Height          =   270
      Left            =   5250
      MaxLength       =   1
      TabIndex        =   19
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox Text17 
      Height          =   270
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   18
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      Height          =   270
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   17
      Top             =   4020
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   4110
      MaxLength       =   7
      TabIndex        =   15
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      Height          =   270
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   13
      Top             =   3420
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   552
      Left            =   4125
      TabIndex        =   28
      Top             =   2808
      Width           =   4692
      Begin VB.TextBox Text12 
         Height          =   270
         Left            =   2940
         MaxLength       =   7
         TabIndex        =   12
         Top             =   200
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "                        日"
         Height          =   225
         Index           =   2
         Left            =   2700
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text11 
         Height          =   270
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   10
         Top             =   200
         Width           =   375
      End
      Begin VB.OptionButton Option4 
         Caption         =   "          月"
         Height          =   180
         Index           =   1
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   1092
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   816
         MaxLength       =   2
         TabIndex        =   8
         Top             =   180
         Width           =   375
      End
      Begin VB.OptionButton Option4 
         Caption         =   "文到           天"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1452
      End
   End
   Begin VB.Frame Frame1 
      Height          =   552
      Left            =   1215
      TabIndex        =   27
      Top             =   2808
      Width           =   2535
      Begin VB.OptionButton Option1 
         Caption         =   "文到次日"
         Height          =   180
         Index           =   1
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "文到當日"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1590
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   0
      Top             =   5340
      Width           =   255
   End
   Begin VB.Label lblDelivery 
      AutoSize        =   -1  'True
      Caption         =   "送達日期:"
      Height          =   180
      Left            =   6105
      TabIndex        =   61
      Top             =   3465
      Width           =   768
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "約定期限:"
      Height          =   180
      Left            =   120
      TabIndex        =   60
      Top             =   3750
      Width           =   765
   End
   Begin MSForms.TextBox Text19 
      Height          =   675
      Left            =   1200
      TabIndex        =   20
      Top             =   4620
      Width           =   7995
      VariousPropertyBits=   -1466939365
      ScrollBars      =   2
      Size            =   "14102;1191"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   195
      Left            =   1830
      TabIndex        =   59
      Top             =   900
      Width           =   7455
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "13150;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "國際分類:"
      Height          =   180
      Left            =   120
      TabIndex        =   58
      Top             =   2550
      Width           =   765
   End
   Begin VB.Label lblOrg 
      AutoSize        =   -1  'True
      Caption         =   "來函機關:"
      Height          =   180
      Left            =   6495
      TabIndex        =   57
      Top             =   1980
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "申請案核駁日"
      Height          =   180
      Left            =   120
      TabIndex        =   56
      Top             =   1980
      Width           =   1080
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "案件目前准駁:         (1:准 , 2:駁)"
      Height          =   180
      Left            =   3315
      TabIndex        =   55
      Top             =   1980
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   60
      X2              =   9120
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   9180
      Y1              =   1890
      Y2              =   1890
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   6
      Left            =   4065
      TabIndex        =   53
      Top             =   1530
      Width           =   2790
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4921;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   5
      Left            =   2460
      TabIndex        =   52
      Top             =   3450
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3757;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   4
      Left            =   2460
      TabIndex        =   51
      Top             =   4050
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3757;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   50
      Top             =   870
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   4065
      TabIndex        =   49
      Top             =   570
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   48
      Top             =   570
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   120
      TabIndex        =   47
      Top             =   1230
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "申請日:"
      Height          =   180
      Left            =   4065
      TabIndex        =   46
      Top             =   1230
      Width           =   585
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   45
      Top             =   1530
      Width           =   945
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   44
      Top             =   1200
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3757;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   2
      Left            =   4920
      TabIndex        =   43
      Top             =   1200
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3387;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   3
      Left            =   1200
      TabIndex        =   42
      Top             =   1530
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3757;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Top             =   4620
      Width           =   765
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "是否算案件數:            (N:不算)"
      Height          =   180
      Index           =   0
      Left            =   4065
      TabIndex        =   34
      Top             =   4380
      Width           =   2970
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "承辦期限:"
      Height          =   180
      Left            =   120
      TabIndex        =   33
      Top             =   4320
      Width           =   765
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   120
      TabIndex        =   32
      Top             =   4020
      Width           =   585
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Left            =   6105
      TabIndex        =   31
      Top             =   3750
      Width           =   765
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Left            =   3180
      TabIndex        =   30
      Top             =   3750
      Width           =   765
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "下一程序:"
      Height          =   180
      Left            =   120
      TabIndex        =   29
      Top             =   3420
      Width           =   765
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "來函期限:"
      Height          =   180
      Left            =   120
      TabIndex        =   26
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   120
      TabIndex        =   25
      Top             =   2250
      Width           =   765
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "專利權是否存在:             (Y/N)"
      Enabled         =   0   'False
      Height          =   180
      Left            =   120
      TabIndex        =   24
      Top             =   5370
      Width           =   3015
   End
End
Attribute VB_Name = "frm06010603_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/22 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strReceiveNo As String '點選的收文號
Dim strTemp As String
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String
Dim sp() As String    'add by sonia 2024/11/21

Dim intWhere As Integer, strSales As String
Dim m_NewReceiveNo As String '總收文號
'原案件性質
Dim m_CP10 As String
Dim m_CP14 As String
Dim bolFinalCheck As Boolean
'Add by Morgan 2006/4/21 來函案件性質
Dim m_NewCP10 As String
Const 裁定駁回 As String = "1007"
Const 部分准駁 As String = "1009"  'add by sonia 2025/4/22
Const 抗告 As String = "509"

Dim m_928Upd As Boolean '是否更新重新委任准駁
Dim m_928CP09 As String '重新委任收文號
Dim m_b307Plus107 As Boolean '分割案有提再審
Dim m_strMemo As String 'C類來函接洽單備註 ADD BY SONIA 2014/5/28
'Added by Morgan 2017/5/10 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_DocDate As String
Public m_AppNo As String
Public m_DeadLine As String
'end 2017/5/10
Dim stCP133 As String  'Added by Morgan 2020/11/13
Dim m_pAgreeOnDate As String 'Modify By Sindy 2021/4/23
'Added by Lydia 2023/09/25
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_bolReKeyInOK As Boolean '是否與2次確認期限一致
'end 2023/09/25

'Add by Morgan 2006/4/24
Private Sub cboOrg_Click()
   '上訴最高行政法院的裁定駁回,下一程序為聲請再審(行政再審504)
   '聲請再審最高行政法院的裁定駁回沒有下一程序
   
   '高等行政法院
   If cboOrg.ListIndex = 0 Then
      Select Case m_NewCP10
         Case 裁定駁回
            Text13 = 抗告: Call ChgType(13)
            SetDeadline
         'modify by sonia 2025/4/22 +4部分准駁
         Case 核駁, 部分准駁
            '507
            If m_CP10 = 行政再審 Then
               'Modify by Morgan 2006/7/31
               '改為法定20日內提上訴，不得延期
               'Text13 = 行政再審:
               Text13 = "507"
               'end 2006/7/31
               Call ChgType(13)
               SetDeadline
            End If
      End Select
   '最高行政法院
   ElseIf cboOrg.ListIndex = 1 Then
      Select Case m_CP10
         Case 行政訴訟上訴 '507
            If m_NewCP10 = 裁定駁回 Then
               Text13 = 行政再審: Call ChgType(13)
               SetDeadline
            End If
         Case 行政再審 '504
            Text13 = "": Label3(5) = ""
            Text14(0) = "": Text14(1) = ""
            Text10 = ""
      End Select
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
        '若來函性質為行政再審or行政訴訟上, 可不輸入期限
        If m_CP10 <> 行政再審 And m_CP10 <> 行政訴訟上訴 Then
            If Text14(0) = "" Or Text14(1) = "" Then
               MsgBox "本所期限、法定期限不可空白 !", vbCritical
               Exit Sub
            End If
        End If
        '若有輸入本所期限
        If Me.Text14(0).Text <> "" Then
            If DBDATE(Me.Text14(0).Text) < strSrvDate(1) Then
                MsgBox "本所期限不可小於系統日期!!!", vbExclamation
                Me.Text14(0).SetFocus
                Me.Text14(0).SelStart = 0
                Me.Text14(0).SelLength = Len(Me.Text14(0).Text)
                Exit Sub
            End If
        End If
         '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
         If Len(Me.Text14(0).Text) > 0 And Len(Me.Text17.Text) > 0 Then
            If Val(Me.Text14(0).Text) < Val(Me.Text17.Text) Then
               MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
               Exit Sub
            End If
         End If
         
         'Add By Sindy 2012/3/7 內外專都只做申請案號第四碼<>'3'之新申請案件性質
         'modify by sonia 2024/11/21 +pa(1)="FCP"
         If pa(1) = "FCP" And Mid(Trim(Text1), 4, 1) <> "3" And InStr(NewCasePtyList, m_CP10) > 0 Then
            If Text15.Text = "" Then
               MsgBox "國際分類不可空白！"
               Text15.SetFocus
               Exit Sub
            End If
         End If
         '2012/3/7 End
         
         'Add by Morgan 2004/9/1 '設定是否為最後檢查旗標
         bolFinalCheck = True
         
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then
            bolFinalCheck = False
            Exit Sub
         End If
         
         '2006/3/29 ADD BY SONIA 已發文請求面詢407但無通知面詢1401且無面詢408之收文者提示訊息
         CHECKFCP407 pa(1), pa(2), pa(3), pa(4)
         '2006/2/29 END
         
         'Add by Sindy 2021/11/22 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If
         
         'Added by Lydia 2023/09/25
         If m_strIR01 <> "" Then
            '下載信件檔
            If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, "", , , True) = False Then
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            If Left(Pub_StrUserSt03, 2) = "F2" Then
               If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
         End If
         'end 2023/09/25
         
         bolFinalCheck = False
         'Add by Morgan 2004/7/28
         '加漏斗
         Screen.MousePointer = vbHourglass
         If FormSave = False Then
            Screen.MousePointer = vbDefault
            MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         End If
         
         'Added by Morgan 2012/11/5
         If Text13 = "" And Left(pa(75), 6) = "Y53309" Then
            MsgBox "本案需調卷轉承辦組報告並寄代！", vbInformation
         End If
         'end 2012/11/5
         
         Screen.MousePointer = vbDefault
         '列印C類接洽記錄單
         If PUB_CaseClosed_1(pa(1), pa(2), pa(3), pa(4)) = False Then
            'Add by Morgan 2008/7/4
            If m_b307Plus107 Then
               'Modified by Lydia 2018/12/17 FCP案C類接洽單同時列印並且上傳到卷宗區
               'g_PrtForm001.PrintCForm m_NewReceiveNo, "請客戶提供重新簽署的委任書以俾辦理訴願程序" & vbCrLf & m_strMemo
               'Modified by Lydia 2019/03/18 改成開啟Word
               'g_PrtForm001.PrintCForm m_NewReceiveNo, "請客戶提供重新簽署的委任書以俾辦理訴願程序" & vbCrLf & m_strMemo, , True
               g_PrtForm001.PrintCFormNew m_NewReceiveNo, "請客戶提供重新簽署的委任書以俾辦理訴願程序" & vbCrLf & m_strMemo, , True
            Else
            'end 2008/7/4
               'Modified by Lydia 2018/12/17 FCP案C類接洽單同時列印並且上傳到卷宗區
               'g_PrtForm001.PrintCForm m_NewReceiveNo, m_strMemo
               'Modified by Lydia 2019/03/18 改成開啟Word
               'g_PrtForm001.PrintCForm m_NewReceiveNo, m_strMemo, , True
               g_PrtForm001.PrintCFormNew m_NewReceiveNo, m_strMemo, , True
            End If
         End If
         
         Unload frm06010603_2
         Unload Me
         
         'Added by Lydia 2023/09/25
         If Me.m_strIR01 <> "" Then
            Unload frm06010603_1
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
         'end 2023/09/25
         'Modified by Morgan 2017/5/10 電子公文
         'frm06010603_1.Show
         'Modified by Lydia 2023/09/25 +Else
         ElseIf m_DocNo <> "" Then
            Unload frm06010603_1
            frm060119.GoNext
         Else
            frm06010603_1.Show
         End If
         'end 2017/5/10

      Case 1
         frm06010603_2.Show
         Unload Me
      Case 2
         Unload frm06010603_1
         Unload frm06010603_2
         Unload Me
   End Select
End Sub

Private Function FormSave() As Boolean
 Dim intStep As Integer, strTxt(1 To 20) As String, strTmp As String, bolChk As Boolean
 Dim i As Integer, strCe(99) As String
 Dim strNP22 As String
 Dim strNP02 As String
 Dim strNP03 As String
 Dim strNP04 As String
 Dim strNP05 As String
 'Add By Cheng 2002/07/03
 Dim rsA As New ADODB.Recordset
 Dim StrSQLa As String
 Dim strCP20 As String, strCP16 As String
 Dim strCP12 As String, strCP13 As String 'Added by Lydia 2023/07/07
 Dim bolReKeyInCase As Boolean 'Added by Lydia 2023/09/25
 
   'Add by Morgan 2007/7/17
   If m_CP10 <> "928" Then
      m_928Upd = PUB_928Check(pa, m_928CP09)
   End If
   
   bolReKeyInCase = False 'Added by Lydia 2023/09/25
   
   '911106 nick transation
   FormSave = True
   
On Error GoTo CheckingErr

cnnConnection.BeginTrans

   'Add by Morgan 2007/7/17
   If m_928Upd = True And m_928CP09 <> "" Then
      PUB_928Update pa, m_928CP09
   End If
   'end 2007/7/17
 
   intStep = 1
   
   'Added by Lydia 2023/07/07 改成變數
   strCP13 = PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))
   strCP12 = GetSalesArea(strCP13)
   'end 2023/07/07
   
   '2
   If pa(1) = "FCP" Then 'add by sonia 2024/11/21
      'If Text8 = "" Then MsgBox "專利權是否存在不可空白，請重新輸入 !", vbCritical: Exit Function
      strExc(0) = "PA17='" & Text8 & "',"
      'If Text7 = "Y" Then strExc(0) = strExc(0) & "PA16='2',PA20=" & CNULL(TransDate(Text6, 2)) & ","
   '   If Text7 = "Y" Then strExc(0) = "PA16='2',PA20=" & CNULL(TransDate(Text6, 2)) & ","
      'Modified by Morgan 2012/3/7 排除802, 804
      'If (m_CP10 >= "101" And m_CP10 <= "105") Or m_CP10 = "107" Or (m_CP10 >= "301" And m_CP10 <= "307") Or m_CP10 = "802" Or m_CP10 = "804" Then
      '2013/10/24 MODIFY BY SONIA 再加入卷宗性質判斷pa(23) = "1",P-083407的503不可更新,否則後續改變原處分也不會更新
      'Modified by Morgan 2014/6/25 +125 衍生設計
      If pa(23) = "1" And ((m_CP10 >= "101" And m_CP10 <= "105") Or m_CP10 = "107" Or m_CP10 = "125" Or (m_CP10 >= "301" And m_CP10 <= "307")) Then
         strExc(0) = strExc(0) & "PA16='" & Me.Text7.Text & "',"
         'Modify by Morgan 2004/12/1 爭議程序不更新基本檔准駁日
         'If IsEmptyText(Text6.Text) = False Then
         If IsEmptyText(Text6.Text) = False And Not (Val(m_CP10) >= 802 And Val(m_CP10) <= 804) Then
            strExc(0) = strExc(0) & "PA20=" & CNULL(TransDate(Text6, 2)) & ","
         End If
      End If
      'Add By Sindy 2012/3/7 +國際分類更新
      strExc(0) = strExc(0) & "PA160=" & CNULL(Text15.Text) & ","
      '2012/3/7 End
      If Right(strExc(0), 1) = "," Then strExc(0) = Left(strExc(0), Len(strExc(0)) - 1)
      strTxt(intStep) = "UPDATE PATENT SET " & strExc(0) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      
      '911106 nick transation
      cnnConnection.Execute strTxt(intStep)
      
      intStep = intStep + 1
   End If   'add by sonia 2024/11/21
      
   '1
'Modify by Morgan 2006/4/21 改用全域變數
'   If frm06010603_2.Text6 = "1" Then
'      i = 核駁
    'modify by sonia 2025/4/22 +4部分准駁
    If m_NewCP10 = 核駁 Or m_NewCP10 = 裁定駁回 Or m_NewCP10 = 部分准駁 Then
      'Modify by Morgan 2005/2/15
      'If Left(m_CP10, 1) = "1" Or Left(m_CP10, 1) = "3" Then
      'Modified by Morgan 2014/6/25 +125 衍生設計
      'modify by sonia 2024/11/21 +FG的120植物新品種保護(FG-001323)
      If Len(m_CP10) = 3 And ((m_CP10 >= "101" And m_CP10 <= "105") Or m_CP10 = "107" Or m_CP10 = "125" Or m_CP10 = "120" Or (m_CP10 >= "301" And m_CP10 <= "307")) Then
         strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='2',CP25=" & TransDate(Text6, 2) & _
            " WHERE CP09='" & strReceiveNo & "'"
            
         cnnConnection.Execute strTxt(intStep)
         intStep = intStep + 1
      Else
         strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='2',CP25=" & TransDate(Label3(3).Caption, 2) & _
            " WHERE CP09='" & strReceiveNo & "'"
                 
         cnnConnection.Execute strTxt(intStep)
         intStep = intStep + 1
      End If
'   Else
'      i = 改變原處分
   End If
   
   'Added by Morgan 2025/3/4
   'Y2099001Murgitroyd+Meta集團(X80668000、 X80669000、X80670000)案件，1002核駁、1202審查意見通知函、1227最後通知預設不請款並備註"簡單報告"
   If pa(1) = "FCP" And pa(75) = "Y2099001" And InStr("X80668000,X80669000,X80670000", pa(26)) > 0 Then
      Text19 = "簡單報告;" & Text19
   End If
   'end 2025/3/4
      
      
   '3
   m_NewReceiveNo = AutoNo("C", 6)
   'Modify by Morgan 2006/4/21 改 IIf(i <> 核駁, ServerDate, "NULL") -->  IIf(m_NewCP10 = 改變原處分, strsrvdate(2), "NULL")
   'Modified by Lydia 2023/07/07 改成變數strCP12,strCP13
   strTxt(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP08," & _
      "CP09,CP10,CP12,CP13,CP14,CP48,CP20,CP32,CP26,CP27,CP43,CP64,CP16,CP17,CP18,CP133) VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & _
      "','" & TransDate(Label3(3).Caption, 2) & "'," & CNULL(TransDate(Text14(0), 2)) & "," & CNULL(TransDate(Text14(1), 2)) & _
      "," & CNULL(Text9) & ",'" & m_NewReceiveNo & "','" & m_NewCP10 & "','" & strCP12 & "','" & strCP13 & _
      "','" & Text16 & "','" & TransDate(Text17, 2) & _
      "','N','N'," & CNULL(Text18) & "," & IIf(m_NewCP10 = 改變原處分, strSrvDate(1), "NULL") & ",'" & strReceiveNo & "'," & CNULL(ChgSQL(Text19)) & ",4000,0,4," & CNULL(stCP133, True) & ")"
   cnnConnection.Execute strTxt(intStep)
   intStep = intStep + 1
   
   'ADD BY SONIA 2014/5/28 Intersil及其子公司的案件在C類接洽單加印
   m_strMemo = ""
   'Remove by Morgan 2017/3/9 改從備註維護功能自行設定(與其他的內容合併)--敏莉 Ex.FCP-49174
   'Select Case Left(pa(26) & "000", 8)
   '   Case "X6217700", "X5272200", "X5422700", "X5819500", "X6380100", "X6554500", "X6036001", "X4899100", "X4899101"
   '      m_strMemo = "若有報價請一併CC給Intersil ！"
   'End Select
   'end 2017/3/9
   'END 2014/5/9
   
   'Modified by Lydia 2024/05/28 改成模組
   ''added by Lydia 2022/05/03 FCP-067004核准前不收費控制：申請至核准(暫不包含領證)不收任何收費 (包含規費及服務費、若客戶提AEP也不收費)
   'If pa(16) <> "1" And InStr("FCP067004000", pa(1) & pa(2) & pa(3) & pa(4)) > 0 Then
   If PUB_GetCP20forSpec(pa(1), pa(2), pa(3), pa(4), pa(16)) = "N" Then
   'end 2024/05/28
         strSql = "update caseprogress set cp20='N', cp16=null, cp17=null, cp18=null where cp09='" & m_NewReceiveNo & "'"
         cnnConnection.Execute strSql
   Else
      'Add by Morgan 2007/7/23 CP20改抓CPM的設定
      'Modify by Morgan 2008/3/27 +pa75
      'Modify by Morgan 2008/4/10 +本所案號
      strCP20 = PUB_GetCP20(pa(1), m_NewCP10, strCP16, pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4))
      If strCP20 = "" Then
         strSql = "update caseprogress set cp20=NULL,cp16=" & Val(strCP16) & ",cp17=0,cp18=" & Val(strCP16) / 1000 & _
            " where cp09='" & m_NewReceiveNo & "'"
         cnnConnection.Execute strSql
      End If
   'end 2007/7/23
   End If 'added by Lydia 2022/05/03
   
   'Added by Lydia 2025/08/19 輸入C類來函時，去檢查上一道承辦人掛工程師，是否為未請款，若是，則發Mail通知工程師；
                  '核駁比照核准，指定特定案件性質
   If m_NewCP10 = 核駁 And pa(1) = "FCP" And InStr("101,102,103,107,307,308", m_CP10) > 0 And Text16 <> "" Then
      If PUB_ChkFCPtoCP14CP60(pa(1), pa(2), pa(3), pa(4), m_NewCP10, m_NewReceiveNo, Text16) = True Then
      End If
   End If
   'end 2025/08/19
   
   '4
'   If i = 改變原處分 Then
   If m_NewCP10 = 改變原處分 Then
      '紀錄改變原處分的結果
      strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='2' WHERE CP09='" & m_NewReceiveNo & "'"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
      
      '2010/9/27 modify by sonia 改抓本所號
      'strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & strReceiveNo & "' AND NP07='" & 改變原處分 & "'"
      strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP02='" & pa(1) & "' and NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "' AND NP06 IS NULL AND NP07='" & 改變原處分 & "'"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   
   strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & strReceiveNo & "' AND NP07='" & 催審 & "'"
   cnnConnection.Execute strTxt(intStep)
   intStep = intStep + 1
   
   '6
   If Text13 <> "" Then
      strNP02 = pa(1)
      strNP03 = pa(2)
      strNP04 = pa(3)
      strNP05 = pa(4)
      strNP22 = GetNextProgressNo  'edit by nickc 2007/02/02 不用 dll 了  objPublicData.GetNextProgressNo
      'Modify By Sindy 2021/4/23 + ,NP23=" & CNULL(TransDate(Text14(2), 2)):約定期限
      strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
         "NP09,NP10,NP13,NP14,NP15,NP22,NP23) VALUES ('" & m_NewReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & _
         "'," & Text13 & "," & TransDate(Text14(0), 2) & "," & TransDate(Text14(1), 2) & _
         "," & CNULL(PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))) & "," & CNULL(Text9) & "," & CNULL(ChgSQL(strSales)) & _
         "," & CNULL(Text19) & "," & strNP22 & "," & CNULL(TransDate(Text14(2), 2)) & ")"
         
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   
   '7
   strExc(0) = "SELECT * FROM CHANGEEVENT WHERE CE01='" & strReceiveNo & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         For i = 1 To 99
            If IsNull(.Fields(i - 1)) Then
               strCe(i) = ""
            Else
               strCe(i) = .Fields(i - 1)
            End If
         Next
      End With
      strExc(1) = ""
      
      '申請日
      If strCe(2) <> "" Then strExc(1) = strExc(1) & "CE03='2',"
      
      '申請人
      For i = 4 To 8
         If strCe(i) <> "" Then
            strExc(1) = strExc(1) & "CE09='1',"
            Exit For
         End If
      Next
      
      '代表人
      bolChk = False
      For i = 10 To 15
         If strCe(i) <> "" Then
            bolChk = True
            Exit For
         End If
      Next
      If Not bolChk Then
         For i = 68 To 91
            If strCe(i) <> "" Then
               bolChk = True
               Exit For
            End If
         Next
      End If
      If bolChk Then strExc(1) = strExc(1) & "CE16='1',"
      
      
      '申請地址
      For i = 23 To 37
         If strCe(i) <> "" Then
            strExc(1) = strExc(1) & "CE38='1',"
            Exit For
         End If
      Next
      
      '專利商標種類代號
      If strCe(39) <> "" Then strExc(1) = strExc(1) & "CE40='1',"
      
      '案件名稱
      For i = 41 To 43
         If strCe(i) <> "" Then
            strExc(1) = strExc(1) & "CE44='1',"
            Exit For
         End If
      Next
      
      '代表人中譯文
      bolChk = False
      For i = 63 To 64
         If strCe(i) <> "" Then
            bolChk = True
            Exit For
         End If
      Next
      If Not bolChk Then
         For i = 92 To 99
            If strCe(i) <> "" Then
               bolChk = True
               Exit For
            End If
         Next
      End If
      
      If bolChk Then strExc(1) = strExc(1) & "CE65='1',"
      
      If strExc(1) <> "" Then
         If Right(strExc(1), 1) = "," Then strExc(1) = Left(strExc(1), Len(strExc(1)) - 1)
         strTxt(intStep) = "UPDATE CHANGEEVENT SET " & strExc(1) & " WHERE CE01='" & strReceiveNo & "'"
         
         cnnConnection.Execute strTxt(intStep)
         intStep = intStep + 1
      End If
   End If
   
   If (m_CP10 >= "101" And m_CP10 <= "105") Or (m_CP10 >= "301" And m_CP10 <= "307") Or (m_CP10 >= "501" And m_CP10 <= "508") Or (m_CP10 >= "801" And m_CP10 <= "805") Then
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " And CP10>='203' AND CP10<='206' "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         While Not rsA.EOF
            strTxt(intStep) = "Update NextProgress Set NP06 ='N' Where NP01='" & rsA.Fields("CP09").Value & "' AND " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP07='411' AND NP06 IS NULL "
            
            cnnConnection.Execute strTxt(intStep)
            intStep = intStep + 1
            rsA.MoveNext
         Wend
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
   
   '2012/10/19 ADD BY SONIA Y53309審查意見通知1202或核駁要內部收文901,承辦期限為系統日起7天(日曆天)--吳若芬
   '2013/1/24 MODIFY BY SONIA 加 Y51542
   'Modified by Morgan 2013/3/5 取消 Y51542 --吳彩菱
   'Modified by Morgan 2013/8/28 ,加 Y34210 + X51446 --邱子瑜,Y51542 --吳彩菱
   'Modified by Morgan 2013/8/30 ,+ Y47453 & X55778 --羅惠蓮
   'Modified by Morgan 2013/9/6 + Y20065 --邱子瑜
'Modified by Morgan 2013/9/18 改呼叫共用函數
'   If Left(pa(75), 6) = "Y53309" Or Left(pa(75), 6) = "Y51542" Or Left(pa(75), 6) = "Y20065" Or _
'      (Left(pa(75), 6) = "Y34210" And Left(pa(26), 6) = "X51446") Or _
'      (Left(pa(75), 6) = "Y47453" And Left(pa(26), 6) = "X55778") Then
'
'      strExc(1) = AutoNo("B", 6)
'      strExc(2) = "901"
'
'      'Added by Morgan 2103/8/28
'      'Y51542 改收其他翻譯 --吳彩菱
'      If Left(pa(75), 6) = "Y51542" Then
'         strExc(2) = "927"
'      End If
'      'Y34210 + X51446 14天 --邱子瑜
'      If Left(pa(75), 6) = "Y34210" And Left(pa(26), 6) = "X51446" Then
'         strExc(5) = Val(CompDate(2, 14, strSrvDate(1)))
'      'Added by Morgan 2013/9/6
'      'Y20065 15天 --邱子瑜
'      ElseIf Left(pa(75), 6) = "Y20065" Then
'         strExc(5) = Val(CompDate(2, 15, strSrvDate(1)))
'
'      Else
'      'end 2013/8/28
'         strExc(5) = Val(CompDate(2, 7, strSrvDate(1)))
'      End If 'Added by Morgan 2103/8/28
'Add by Lydia 2014/12/3 核駁及審查意見通知函備註
   'If PUB_ChkAutoRec(pa(1), pa(75), pa(26), , strExc(2), strExc(5), , , pa(27), pa(28), pa(29), pa(30)) = True Then

       Dim sMemo As String
        'Remove by Lydia 2021/11/05
        'strExc(2) = "": strExc(5) = ""
        'strExc(7) = "": strExc(3) = "": strExc(4) = "": strExc(10) = ""
        'If Not IsNull(pa(27)) Then strExc(7) = ChangeCustomerL(pa(27))
        'If Not IsNull(pa(28)) Then strExc(3) = ChangeCustomerL(pa(28))
        'If Not IsNull(pa(29)) Then strExc(4) = ChangeCustomerL(pa(29))
        'If Not IsNull(pa(30)) Then strExc(10) = ChangeCustomerL(pa(30))
        'end 2021/11/05
        'Modified by Morgan 2020/11/12 +stCP133
        'Modified by Lydia 2021/11/05 分別傳回B類收文(承辦期限、所限)和C類來函(承辦期限和指定送件日期)
        'sMemo = PUB_GetIncomMemoNew(pa(1) & pa(2) & pa(3) & pa(4), pa(1), strExc(2), ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)), , strExc(5), stCP133 _
                    , strExc(7), strExc(3), strExc(4), strExc(10))
        'strExc(7) = "": strExc(3) = "": strExc(4) = "": strExc(10) = ""
        Dim stBCP10 As String, stBCP48   As String, stBCP06 As String, stCCP48 As String, stCCP142 As String
        sMemo = PUB_GetIncomMemoNew(pa(1) & pa(2) & pa(3) & pa(4), pa(1), ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)) & "," & ChangeCustomerL(pa(27)) & "," & ChangeCustomerL(pa(28)) & "," & ChangeCustomerL(pa(29)) & "," & ChangeCustomerL(pa(30)), _
                       "", stCP133, m_NewCP10, stCCP48, stCCP142, stBCP10, stBCP48, stBCP06)
                       
        'Added by Lydia 2021/11/05 更新C類來函的承辦期限和指定送件日期，一併更新指定送件日期之前CP164=2
        If stCCP48 <> "" Then
            'Modified by Lydia 2021/11/16 加註cp64
            strSql = "Update CaseProgress set cp48=" & stCCP48 & ", cp141='3', cp142=" & stCCP142 & ", cp164='2' " & _
                        ", cp64='客戶指定" & ChangeWStringToTDateString(stCCP142) & "之前送件;'||cp64 where cp09='" & m_NewReceiveNo & "' "
            cnnConnection.Execute strSql, intI
        End If
        'end 2021/11/05
        
      'Modified by Lydia 2021/11/05 PUB_GetIncomMemoNew已有另外抓B類收文設定
      'If Len(sMemo) > 0 Then
      '   If strExc(2) = "" Then strExc(2) = "901"
      If stBCP10 <> "" Then
'end 'Add by Lydia 2014/12/3
         strExc(1) = AutoNo("B", 6)
'end 2013/9/18
         
         'Add By Sindy 2021/6/17 非智慧局期限，要掛本所期限
         'Remove by Lydia 2021/11/05 改從PUB_GetIncomMemoNew取得
         'Call GetPrjState6HM(pa(1), strExc(2), "cpm34", strExc(0))
         'strExc(6) = "" '本所期限
         ''110/7/21 淑華改, 本所期限=承辦期限; 因為這是算客戶指定的期限(為應該出給客戶的期限)
'         'If Val(strExc(5)) > 0 And strExc(0) = "N" And strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
'         '   strExc(6) = PUB_GetFCPOurDeadline(DBDATE(strExc(5)), , , , "N")
         'If Val(strExc(5)) > 0 And strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
         '  strExc(6) = DBDATE(strExc(5))
         'End If
         ''2021/6/17 END
         'end 2021/11/05
         
         'Modified by Morgan 2019/8/8 +CP20(抓設定)
         'Modified by Lydia 2022/01/05 改抓變數 strExc(2)=> stBCP10 ; ex.FCP-64282的告代CP20不等於N
         strCP20 = PUB_GetCP20(pa(1), stBCP10, strCP16, pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4))
         'Modify By Sindy 2021/6/17 + ,cp06
         'Modified by Lydia 2021/11/05 改變數strExc(2)=>stBCP10, strExc(5)=> stBCP48, strExc(6)=> stBCP06
         'Modified by Morgan 2021/11/9 CP43改放核駁函收文號(原放點選的收文號,改和其他的OA來函一致)
         'Modified by Lydia 2023/07/07 改成變數strCP12,strCP13
         strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp14,cp20,cp43,cp48,cp06,cp16)" & _
            " values('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & DBDATE(Label3(3)) & ",'" & strExc(1) & "','" & stBCP10 & "','90','" & strCP12 & "','" & strCP13 & "','" & Text16 & "'" & _
            ",'" & strCP20 & "','" & m_NewReceiveNo & "'," & CNULL(stBCP48, True) & "," & CNULL(stBCP06, True) & "," & CNULL(strCP16, True) & " )"
      
'end 'Add by Lydia 2014/12/3
         cnnConnection.Execute strSql, intI
      End If
   '2012/10/19 END
   
   'Added by Lydia 2025/02/05 輸入中間程序來函時自動產生行事曆
   If PUB_AddSCforIncomMemo(pa(1), pa(2), pa(3), pa(4), m_NewReceiveNo, m_NewCP10, ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)) & "," & ChangeCustomerL(pa(27)) & "," & ChangeCustomerL(pa(28)) & "," & ChangeCustomerL(pa(29)) & "," & ChangeCustomerL(pa(30))) = False Then
       GoTo CheckingErr
   End If
   'end 2025/02/05
   
   'Added by Morgan 2017/5/10 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, m_NewReceiveNo, pa(1), pa(2), pa(3), pa(4), m_NewCP10, "2"
   'Added by Morgan 2021/6/11 紙本公文--何淑華
   Else
      PUB_FCPOAInform m_NewReceiveNo, pa(1), pa(2), pa(3), pa(4), m_NewCP10
   End If
   'end 2017/5/10
   
   'Added by Lydia 2023/09/25
   If m_strIR01 <> "" Then
      '核駁函輸入後請將整封郵件存入系統
      If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, m_NewReceiveNo, IIf(Pub_StrUserSt03 = "F22", "ALTR", IIf(pa(9) <> 台灣國家代號, "PAT", "RX"))) = False Then 'PAT.陸代郵件
         GoTo CheckingErr
      End If
      If Left(Pub_StrUserSt03, 2) = "F2" Or Text14(1) = "" Then
         PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm06010603_1", IIf(Pub_StrUserSt03 = "F22", m_NewReceiveNo, "")
         bolReKeyInCase = True
      Else
         PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm06010603_1", m_NewReceiveNo, m_bolReKeyInOK
      End If
   End If
   'end 2023/09/25
   
   'Added by Morgan 2022/7/20 --Anny
   '收到TIPO來函 "再審107之核駁1002" or "訴願501之核駁1002",申請人編號為舊名字(X___001)時，則系統自動發email通知相關人員警示本案申請人已有新名字
   If m_CP10 = "107" Or m_CP10 = "501" Then
      For intI = 1 To 5
         If (Len(pa(25 + intI)) = 9 And Right(pa(25 + intI), 1) <> "0") Then
            'Modified by Lydia 2023/09/25 +, bolReKeyInCase
            PUB_POAInform pa(1), pa(2), pa(3), pa(4), m_NewReceiveNo, bolReKeyInCase
            Exit For
         End If
      Next
   End If
   'end 2022/7/20
   
   'Added by Lydia 2023/07/28 外專-FCP專利連結案管制：若有舉發成立確定輸入來函時，發一封Email給承辦工程師
      '舉發成立確定==>1.舉發答辯804駁
                     '2.舉發答辯804訴願501駁(經過收文->舉發答辯->核駁->訴願->核駁(現在來函m_NewReceiveNo)的流程)
                     '3.行政訴訟503駁
                     '4.行政訴訟上訴507駁
   'modify by sonia 2025/4/22 +4部分准駁
   'If pa(177) = "Y" And m_NewCP10 = 核駁 Then
   If pa(177) = "Y" And (m_NewCP10 = 核駁 Or m_NewCP10 = 部分准駁) Then
      If PUB_GetFCPlinkMC("1", TransDate(Label3(3).Caption, 2), pa, strReceiveNo, m_CP10, m_NewCP10, strCP12, strCP13, Text16.Text) = True Then
      End If
   End If
   'end 2023/07/28
   
   cnnConnection.CommitTrans
   Exit Function
   
CheckingErr:
   cnnConnection.RollbackTrans
   FormSave = False
   
End Function

Private Sub Form_Initialize()
'add by nickc 2007/02/02
   ReDim pa(1 To TF_PA) As String
   ReDim sp(1 To tf_SP) As String   'add by sonia 2024/11/21
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm06010603_2
      pa(1) = .Text2
      pa(2) = .Text3
      pa(3) = .Text4
      pa(4) = .Text5
      strReceiveNo = .Tag
      strSales = strExc(5)
      'Add by Morgan 2006/4/21
      Select Case .Text6
         Case "1"
            m_NewCP10 = 核駁
            Me.Caption = Me.Caption & "(核駁)"
         Case "2"
            m_NewCP10 = 改變原處分
            Me.Caption = Me.Caption & "(改變原處分)"
         Case "3"
            m_NewCP10 = 裁定駁回
            Me.Caption = Me.Caption & "(裁定駁回)"
         'add by sonia 2025/4/22 +4部分准駁
         Case "4"
            m_NewCP10 = 部分准駁
            Me.Caption = Me.Caption & "(部分准駁)"
         'end 2025/4/22
      End Select
      '2006/4/21 end
      If pa(1) = "FG" Then 'add by sonia 2024/11/21
         sp(1) = pa(1)
         sp(2) = pa(2)
         sp(3) = pa(3)
         sp(4) = pa(4)
         ReadServicePractice
      Else
         ReadPatent
         'Add by Morgan 2006/4/24
         'modify by sonia 2025/4/22 行政再審+4部分准駁
         If ((m_CP10 = 行政訴訟上訴 Or m_CP10 = 行政再審) And m_NewCP10 = 裁定駁回) Or (m_CP10 = 行政再審 And (m_NewCP10 = 核駁 Or m_NewCP10 = 部分准駁)) Then
            Me.lblOrg.Visible = True
            Me.cboOrg.Visible = True
            Me.cboOrg.AddItem "高等行政法院"
            Me.cboOrg.AddItem "最高行政法院"
            Me.cboOrg.ListIndex = 0
         End If
         '2006/4/24 end
      End If     'add by sonia 2024/11/21
   End With
   Combo2.ListIndex = 0
   
   'Added by Lydia 2023/09/25
   m_strIR01 = frm06010603_2.m_strIR01
   m_strIR02 = frm06010603_2.m_strIR02
   m_strIR03 = frm06010603_2.m_strIR03
   m_strIR04 = frm06010603_2.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
      lblDelivery.Visible = True: txtDelivery.Visible = True
      'modify by sonia 2025/4/22 +4部分准駁
      'If pa(1) = "FCP" And InStr("503,", m_CP10) > 0 And m_NewCP10 = 核駁 Then
      If pa(1) = "FCP" And InStr("503,", m_CP10) > 0 And (m_NewCP10 = 核駁 Or m_NewCP10 = 部分准駁) Then
         '行政訴訟核駁
      Else
          cmdOK(0).Enabled = False
      End If
   Else
      lblDelivery.Visible = False: txtDelivery.Visible = False
   End If
   'end 2023/09/25
   
   Dim strTmp As String
   
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
   Text9.Text = "（" & strTmp & "）智專一（二）字第號"
   
   'Added by Morgan 2017/5/10 電子公文
   If m_DocNo <> "" Then
      If m_DocWord <> "" Then
         Text9 = m_DocWord & "字第" & m_DocNo & "號"
      ElseIf m_DocNo <> "" Then
         Text9 = Replace(Text9, "第號", "第" & m_DocNo & "號")
      End If
      If m_DocDate <> "" And Text6.Locked = False Then
         Text6 = TransDate(m_DocDate, 1)
      End If
      
      '國際分類
      If PUB_GetEDocData(m_DocNo, strExc(1), strExc(2)) Then
         If Text15 = "" Then Text15 = Left(strExc(2), 4)
      End If
   
      '期限
      If m_DeadLine <> "" Then
         Option1(1).Value = True
         If Len(m_DeadLine) >= 7 Then
            Option4(2).Value = True
            Text12 = m_DeadLine
            Text12_Validate False
         ElseIf Right(m_DeadLine, 1) = "日" Then
            Option4(0).Value = True
            Text10 = Val(m_DeadLine)
            Text10_Validate False
         ElseIf Right(m_DeadLine, 1) = "月" Then
            Option4(1).Value = True
            Text11 = Val(m_DeadLine)
            Text11_Validate False
         End If
      End If
   End If
   'end 2017/5/10
   
   Check908 pa 'Add by Morgan 2009/10/1
   
   'Add By Sindy 2021/5/7
   If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
      Label8.Visible = True
      Text14(2).Visible = True
   Else
      Label8.Visible = False
      Text14(2).Visible = False
   End If
   '2021/5/7 END
End Sub

Private Sub ReadPatent()
   Dim Lbl As Object, i As Integer, rsTemp1 As New ADODB.Recordset
   'Modify by Morgan 2006/4/21 改用全域變數
   'Dim strTmp As String
   
   For Each Lbl In Label3
      Lbl.Caption = ""
   Next
   Label3(3).Caption = frm06010603_1.Text5.Text
   Text2 = pa(1)
   Text3 = pa(2)
   Text4 = pa(3)
   Text5 = pa(4)
   
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      LblFM2 = pa(5)
      Label3(2) = pa(10)
      Text1 = pa(11)
      If pa(16) = "1" Then
         Label3(6) = "基本檔目前准駁 : 准"
      ElseIf pa(16) = "2" Then
         Label3(6) = "基本檔目前准駁 : 駁"
      Else
         Label3(6) = "基本檔目前准駁 : 無"
      End If
      Text8 = Empty
   End If
   
   m_CP10 = ""
   strExc(0) = "SELECT CP10,CPM03,CP12,CP13,CP14 FROM CASEPROGRESS,CASEPROPERTYMAP WHERE " & _
      "CP09='" & strReceiveNo & "' AND CP01=CPM01(+) AND CP10=CPM02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      m_CP10 = "" & .Fields(0).Value
      m_CP14 = "" & .Fields(4).Value
      Label3(1) = "" & .Fields(1)
      'Modify by Morgan 2006/4/21 改用全域變數
      'Text16.Text = PUB_GetFCPPromoterNo(strReceiveNo, IIf(frm06010603_2.Text6.Text = "1", 核駁, 改變原處分), "" & .Fields(4))
      Text16.Text = PUB_GetFCPPromoterNo(strReceiveNo, m_NewCP10, "" & .Fields(4))
      ChgType 16
   End If
   End With
   
   '承辦期限
   'Add by Morgan 2004/11/11 FCP初審核駁承辦期限不抓CaseFee(CaseFee設定為6天)，設定為10天。
   '2005/9/29 MODIFY BY SONIA 加入 501
   '2006/3/24 MODIFY BY SONIA 加入 503,504,507 但為5天
   'Modify by Morgan 2006/4/21 改用全域變數
   'If strTmp = 核駁 And InStr("101,102,103,104,105,501,503,504,507", m_CP10) > 0 Then
   '2006/9/15 MODIFY BY SONIA 加分割307
   'modify by sonia 2025/4/22 +4部分准駁
   'If m_NewCP10 = 核駁 And InStr("101,102,103,104,105,501,503,504,507,307", m_CP10) > 0 Then
   If (m_NewCP10 = 核駁 Or m_NewCP10 = 部分准駁) And InStr("101,102,103,104,105,501,503,504,507,307", m_CP10) > 0 Then
      Select Case m_CP10
         '2006/9/15 MODIFY BY SONIA 加分割307
         Case "101", "102", "103", "104", "105", "501", "307"
            Text17 = TransDate(CompWorkDay(10, TransDate(Label3(3).Caption, 2), 0), 1)
         '2006/3/24 ADD BY SONIA 503,504,507之核駁設定為5天
         Case "503", "504", "507"
            Text17 = TransDate(CompWorkDay(5, TransDate(Label3(3).Caption, 2), 0), 1)
         '2006/3/24 END
      End Select
   Else
      'Modify by Morgan 2007/10/11 承辦期限改呼叫共用函數計算
      'Modify by Morgan 2006/4/21 改用全域變數
      'strExc(0) = "SELECT CF04 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & strTmp & "'"
      'strExc(0) = "SELECT CF04 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & m_NewCP10 & "'"
      'intI = 1
      'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      'If intI = 1 Then
      '   With RsTemp
      '      If Not IsNull(.Fields(0)) Then
      '         Text17 = TransDate(CompWorkDay(Val(.Fields(0)), TransDate(Label3(3).Caption, 2), 0), 1)
      '      End If
      '   End With
      'End If
      Text17 = TransDate(Pub_GetHandleDay(pa(1), pa(9), m_NewCP10, TransDate(Label3(3).Caption, 2)), 1)
      'end 2007/10/11
   End If
   
   'Modified by Lydia 2016/08/15 改成共用模組,移到下方
'   'Added by Morgan 2015/4/20
'   '先正達OA承辦期限設7個工作天,若下一程序為 804,501-509時設2個工作天(24Hr)
'   If InStr("Y4830900,Y4830901,Y4830902,Y4830903,Y4830904,Y4830905,Y4830908,Y5132600", Left(pa(75) & "000", 8)) > 0 Then
'      If Text13 = "804" Or (Text13 >= "501" And Text13 <= "509") Then
'         Text17 = TransDate(CompWorkDay(2, TransDate(Label3(3).Caption, 2), 0), 1)
'      Else
'         Text17 = TransDate(CompWorkDay(7, TransDate(Label3(3).Caption, 2), 0), 1)
'      End If
'   'Added by Morgan 2015/7/3 --吳彩菱
'   'Y51753+X45149010 承辦天數:14 起算日期:系統日
'   ElseIf Left(pa(75) & "000", 8) = "Y5175300" And Left(pa(26) & "000", 8) = "X4514901" Then
'      Text17 = TransDate(CompDate(2, 14, strSrvDate(1)), 1)
'   End If
'   'end 2015/4/20
   
   '下一程序
   'Modify by Morgan 2006/4/21 裁定駁回用1007抓下一程序
   'strExc(0) = "SELECT CF15 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & m_CP10 & "'"
   If m_NewCP10 = 裁定駁回 Then
      strExc(0) = "SELECT CF15 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & 裁定駁回 & "'"
   'Add by Morgan 2007/4/19 分割案的核駁若母案已提再審時下一程序改帶訴願(再審的下一程序)
   ElseIf m_CP10 = 分割 And m_NewCP10 = 核駁 Then
      'Modify by Morgan 2008/11/25 改判斷母案有再審且發文日且早於分割案的發文日
      'strExc(0) = "select 1 from caseprogress where cp10='107' and cp27>0 and cp57 is null and (cp01,cp02,cp03,cp04) in (select dc05,dc06,dc07,dc08 from divisioncase where dc01='" & pa(1) & "' and dc02='" & pa(2) & "' and dc03='" & pa(3) & "' and dc04='" & pa(4) & "')"
      'Modified by Morgan 2014/8/19 同一天發文也算
      '2015/8/24 MODIFY BY SONIA 加判斷分割案若有收435續行母案再審則分割核駁下一程序帶訴願 FCP-049062
      'strExc(0) = "select 1 from caseprogress a where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'  and cp10='307' and exists(select dc05,dc06,dc07,dc08 from divisioncase,caseprogress b  where dc01=a.cp01 and dc02=a.cp02 and dc03=a.cp03 and dc04=a.cp04 and b.cp01=dc05 and b.cp02=dc06 and b.cp03=dc07 and b.cp04=dc08 and b.cp10='107' and b.cp27<=a.cp27)"
      strExc(0) = "select 1 from caseprogress a where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'  and cp10='307' and exists(select dc05,dc06,dc07,dc08 from divisioncase,caseprogress b  where dc01=a.cp01 and dc02=a.cp02 and dc03=a.cp03 and dc04=a.cp04 and b.cp01=dc05 and b.cp02=dc06 and b.cp03=dc07 and b.cp04=dc08 and b.cp10='107' and b.cp27<=a.cp27)" & _
                  " union select 2 from caseprogress a where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'  and cp10='435' and nvl(cp27,0)>0 "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_b307Plus107 = True 'Add by Morgan 2008/7/4
         strExc(0) = "SELECT CF15 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='107'"
      Else
         strExc(0) = "SELECT CF15 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & m_CP10 & "'"
      End If
   Else
      strExc(0) = "SELECT CF15 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & m_CP10 & "'"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         If Not IsNull(.Fields(0)) Then Text13 = .Fields(0): ChgType 13
      End With
   End If
   
   If m_DocNo <> "" Then stCP133 = PUB_GetEDocDate(m_DocNo) 'Added by Morgan 2020/11/13 官方發文日
   'Added by Lydia 2016/08/15
   'Modify By Sindy 2017/5/9 + , TransDate(Text14(0), 2)
   'Modified by Morgan 2018/5/22  +m_CP10
   'Modified by Morgan 2020/11/13 +CP133
   Call Pub_SetExceptCP48(pa(75), pa(26), m_NewCP10, TransDate(Label3(3).Caption, 2), Text17, Text13.Text, TransDate(Text14(0), 2), m_CP10, stCP133)
   
   '6預設來函期限(下一程序期限)
   SetDeadline

   '顯示專用權是否在存
   Me.Text8.Text = "" & pa(17)

   EnableTextBox Text7, False
   '顯示目前准駁
   Me.Text7.Text = "" & pa(16)
   
   'Add By Sindy 2012/3/7 +國際分類
   Me.Text15.Text = "" & pa(160)
   '2012/3/7 End
   
   '控制案件性質3碼
   If Len(m_CP10) = 3 Then
      'Modified by Morgan 2012/3/7 排除 802, 804
      'If (m_CP10 >= "101" And m_CP10 <= "105") Or m_CP10 = "107" Or (m_CP10 >= "301" And m_CP10 <= "307") Or m_CP10 = "802" Or m_CP10 = "804" Then
      'Modified by Morgan 2014/6/25 +125 衍生設計
      If (m_CP10 >= "101" And m_CP10 <= "105") Or m_CP10 = "107" Or m_CP10 = "125" Or (m_CP10 >= "301" And m_CP10 <= "307") Then
         Me.Text7.Text = "2"
      End If
   End If
   
   'Modified by Morgan 2014/6/25 +125 衍生設計
   If Len(m_CP10) = 3 And ((m_CP10 >= "101" And m_CP10 <= "105") Or m_CP10 = "107" Or m_CP10 = "125" Or (m_CP10 >= "301" And m_CP10 <= "307")) Then
      EnableTextBox Text6, True
   Else
      EnableTextBox Text6, False
   End If
   
   'Remove by Morgan 2007/7/20 不再預設"N" --靜芳
   'If m_CP10 = "804" Then
   '   Me.Text8.Text = "N"
   'End If
   'end 2007/7/20
   
End Sub

'add by sonia 2024/11/21 +服務業務 FG-001323植物新品種保護120
Private Sub ReadServicePractice()
Dim Lbl As Object, i As Integer, rsTemp1 As New ADODB.Recordset
   
   For Each Lbl In Label3
      Lbl.Caption = ""
   Next
   Label3(3).Caption = frm06010603_1.Text5.Text
   Text2 = sp(1)
   Text3 = sp(2)
   Text4 = sp(3)
   Text5 = sp(4)
   
   If ClsPDReadServicePracticeDatabase(sp(), intWhere) Then
      LblFM2 = sp(5)
      Label3(2) = sp(10)
      Text1 = sp(11)
      Label3(6) = "基本檔目前准駁 : 無"
      Text8 = Empty
   End If
   
   m_CP10 = ""
   strExc(0) = "SELECT CP10,CPM03,CP12,CP13,CP14 FROM CASEPROGRESS,CASEPROPERTYMAP WHERE " & _
      "CP09='" & strReceiveNo & "' AND CP01=CPM01(+) AND CP10=CPM02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   With RsTemp
      If intI = 1 Then
         m_CP10 = "" & .Fields(0).Value
         m_CP14 = "" & .Fields(4).Value
         Label3(1) = "" & .Fields(1)
         Text16.Text = PUB_GetFCPPromoterNo(strReceiveNo, m_NewCP10, "" & .Fields(4))
         ChgType 16
      End If
   End With
   
   '承辦期限
   Text17 = TransDate(Pub_GetHandleDay(sp(1), sp(9), m_NewCP10, TransDate(Label3(3).Caption, 2)), 1)
   
   '下一程序
   strExc(0) = "SELECT CF15 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & sp(9) & "' AND CF03='" & m_CP10 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         If Not IsNull(.Fields(0)) Then Text13 = .Fields(0): ChgType 13
      End With
   End If
   
   Call Pub_SetExceptCP48(pa(75), pa(26), m_NewCP10, TransDate(Label3(3).Caption, 2), Text17, Text13.Text, TransDate(Text14(0), 2), m_CP10, stCP133)
   
   '6預設來函期限(下一程序期限)
   SetDeadline
  
   EnableTextBox Text6, True
   
End Sub
'end 2024/11/21

Private Function ChgType(i As Integer) As Boolean
 Dim strTempName As String
   ChgType = False
   Select Case i
      Case 13
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty("FCP", Text13, strTempName, False) Then
         If ClsPDGetCaseProperty("FCP", Text13, strTempName, False) Then
            Label3(5) = strTempName
            ChgType = True
         Else
            Label3(5) = ""
         End If
      Case 16
        'Modify By Cheng 2003/04/08
        '若有輸入承辦人
        If Me.Text16.Text <> "" Then
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetStaff(Text16.Text, strTempName) Then
            If ClsPDGetStaff(Text16.Text, strTempName) Then
               Label3(4) = strTempName
               ChgType = True
            Else
               Label3(4) = ""
            End If
        '若未輸入承辦人
        Else
            Label3(4) = ""
            ChgType = True
        End If
   End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Lydia 2022/01/07
   Set frm06010603_3 = Nothing
End Sub

Private Sub Combo2_Click()
   Select Case Combo2
      Case "中"
         LblFM2 = pa(5)
      Case "英"
         LblFM2 = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         LblFM2 = pa(7)
   End Select
End Sub

Private Sub Text10_GotFocus()
   InverseTextBox Text10
End Sub
'Add by Morgan 2004/12/1
Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then GetTime
End Sub

Private Sub Text11_GotFocus()
   InverseTextBox Text11
End Sub
'Add by Morgan 2004/12/1
Private Sub Text11_Validate(Cancel As Boolean)
   If Text11 <> "" Then GetTime
End Sub

Private Sub Text12_GotFocus()
   InverseTextBox Text12
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   If Option4(2).Value = False Then Exit Sub
   If Text12 = "" Then
      MsgBox "來函期限不可空白 !", vbCritical
      Cancel = True
   Else
      If ChkDate(Text12) Then
         If Val(Text12) < Val(strSrvDate(2)) Then
            MsgBox "來函期限不可小於系統日 !", vbCritical
            Cancel = True
         Else
            Text14(1) = Text12
            'Modified by Morgan 2014/11/20 外專改回舊規則
            ''Added by Morgan 2014/10/9
            'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
            '   Text14(0) = TransDate(PUB_GetOurDeadline(Text14(1)), 1)
            'Else
            ''end 2014/10/9
            
            'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
            If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
               'Modify By Sindy 2021/4/23 + m_pAgreeOnDate
               'Modify By Sindy 2023/6/7 + Text13, pa(1), pa(2), pa(3), pa(4)
               'Modify By Sindy 2025/2/12 傳入C
               Text14(0) = TransDate(PUB_GetFCPOurDeadline(Text14(1), 2, , m_pAgreeOnDate, , Text13, pa(1), pa(2), pa(3), pa(4), "C"), 1)
               Text14(2) = TransDate(m_pAgreeOnDate, 1) 'Add By Sindy 2021/5/7
            Else
            'end 2019/7/11
         
               Text14(0) = TransDate(CompDate(2, -2, TransDate(Text14(1), 2)), 1)
               
            End If 'Added by Morgan 2019/7/11
            
            'End If 'Added by Morgan 2014/10/9
            'end 2014/11/20
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text12
End Sub

Private Sub Text13_Change()
   If Text13 = 抗告 Then
      SetDeadline
   End If
End Sub

Private Sub Text13_GotFocus()
    InverseTextBox Text13
End Sub
'Modify by Morgan 2006/4/24 整理
Private Sub Text13_Validate(Cancel As Boolean)
   '若來函性質為行政再審,行政訴訟上訴則可不輸入下一程序
   Label3(5) = ""
   If m_CP10 <> 行政再審 And m_CP10 <> 行政訴訟上訴 Then
      If Text13 = "" Then
         MsgBox "下一程序不可空白 !", vbCritical
         Cancel = True: Exit Sub
      End If
   End If
   If Text13 <> "" Then
      If Len(Me.Text13.Text) <> 3 Then
         MsgBox "下一程序欄位值必須為三碼 !", vbCritical
         Text13_GotFocus
         Cancel = True: Exit Sub
      End If
      If ChgType(13) = False Then
         Cancel = True: Exit Sub
      End If
   Else
      Text14(0) = "": Text14(1) = ""
   End If
End Sub

Private Sub Text14_GotFocus(Index As Integer)
   InverseTextBox Text14(Index)
End Sub

Private Sub Text14_Validate(Index As Integer, Cancel As Boolean)
 Static iTime As Integer
 Static strStatic(0 To 1) As String
   If Text14(Index) <> "" Then
      If Not ChkDate(Text14(Index)) Then
         Cancel = True
      Else
         'Add By Cheng 2002/10/15
         '若有輸入本所期限時, 不可小於系統日
         If Index = 0 Then
            If Len(Me.Text14(0).Text) = 8 Then
               If Val(Me.Text14(0).Text) < strSrvDate(1) Then
                  MsgBox "本所期限不可小於系統日!!!", vbExclamation + vbOKOnly
                  Cancel = True
               End If
            ElseIf Len(Me.Text14(0).Text) = 7 Or Len(Me.Text14(0).Text) = 6 Then
               If Val(Me.Text14(0).Text) + 19110000 < strSrvDate(1) Then
                  MsgBox "本所期限不可小於系統日!!!", vbExclamation + vbOKOnly
                  Cancel = True
               End If
            End If
         End If
         
         If Index = 1 Then
            If Not ChkRange(Text14(0), Text14(1), "本所期限、法定期限") Then
               Cancel = True
            Else
               iTime = iTime + 1
               If iTime = 1 Then
                  strStatic(0) = Text14(0)
                  strStatic(1) = Text14(1)
               End If
               'Modify by Morgan 2004/9/1
               '離開時不檢查,存檔前檢查
               'If Text14(0) <> strStatic(0) Or Text14(1) <> strStatic(1) Then
               '   Cancel = Not CheckDueDate
               'End If
            End If
         End If
      End If
   End If
   If Cancel = True Then TextInverse Text14(Index)
End Sub
'與來函期限比對
Private Function CheckDueDate() As Boolean
   If ClsLawChkMRec(TransDate(Label3(3).Caption, 2), pa(1) & pa(2) & pa(3) & pa(4), strExc(1), strExc(2)) Then
      If Text14(0) <> TransDate(strExc(1), 1) Then
         If MsgBox("與櫃台之來函收文記錄本所期限 ( " & TransDate(strExc(1), 1) & ") 不符，請確認 !", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
            If Text14(0).Enabled = True Then Text14(0).SetFocus
            Exit Function
         End If
      ElseIf Text14(1) <> TransDate(strExc(2), 1) Then
         If MsgBox("與櫃台之來函收文記錄法定期限 ( " & TransDate(strExc(2), 1) & ") 不符，請確認 !", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
            If Text14(1).Enabled = True Then Text14(1).SetFocus
            Exit Function
         End If
      End If
   'Added by Morgan 2017/5/10 電子公文
   ElseIf m_DocNo <> "" Then
      If m_DeadLine <> "" Then
         If Len(m_DeadLine) >= 7 Then
            strExc(2) = m_DeadLine
         ElseIf Right(m_DeadLine, 1) = "日" Then
            strExc(2) = CompDate(2, Val(m_DeadLine), Label3(3))
         ElseIf Right(m_DeadLine, 1) = "月" Then
            strExc(2) = CompDate(1, Val(m_DeadLine), Label3(3))
         End If
         If Text14(1) <> TransDate(strExc(2), 1) Then
            If MsgBox("與電子公文之法定期限 ( " & TransDate(strExc(2), 1) & ") 不符，請確認 !", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Function
            End If
         End If
      End If
   'end 2017/5/10
   Else
      If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Function
      End If
   End If
   CheckDueDate = True
End Function

'Add By Sindy 2012/3/7
Private Sub Text15_GotFocus()
   InverseTextBox Text15
End Sub

'Add By Sindy 2012/3/7
Private Sub Text15_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text16_GotFocus()
   InverseTextBox Text16
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text17_GotFocus()
   InverseTextBox Text17
End Sub

Private Sub Text17_Validate(Cancel As Boolean)
   If Text17 <> "" Then
      If ChkWorkDay(TransDate(Text17, 2)) Then
         'Modify by Morgan 2010/11/9
         'If Text17 > Text14(0) And Me.Text14(0).Text <> "" Then
         If Val(Text17) > Val(Text14(0)) And Me.Text14(0).Text <> "" Then
            MsgBox "承辦期限不可大於本所期限，請重新輸入 !", vbCritical
            Cancel = True
         End If
      Else
         MsgBox "承辦期限不正確，請重新輸入 !", vbCritical
         Cancel = True
      End If
   Else
      If Text13 <> "" Then
         MsgBox "有下一程序且有定義工作天數時不可空白 !", vbCritical
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text17
End Sub

Private Sub Text18_GotFocus()
   InverseTextBox Text18
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text16_Validate(Cancel As Boolean)
   If Text16 <> "" Then
      If ChgType(16) = False Then
         Cancel = True
         TextInverse Text16
      End If
   Else
      Label3(4) = ""
   End If
End Sub

Private Sub Text19_GotFocus()
   InverseTextBox Text19
End Sub

Private Sub Text6_GotFocus()
   InverseTextBox Text6
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If Text6 = "" Then
      'Modify by Morgan 2005/2/14 新申請案或再審或改請程序才控制
      'Modified by Morgan 2014/6/25 +125 衍生設計
      'modify by sonia 2024/11/21  +FG的120植物新品種保護(FG-001323)
      If Len(m_CP10) = 3 And ((m_CP10 >= "101" And m_CP10 <= "105") Or m_CP10 = "107" Or m_CP10 = "120" Or (m_CP10 >= "301" And m_CP10 <= "307")) Then
         MsgBox "申請案核駁日不可空白 !", vbCritical
         Cancel = True
      End If
   Else
      If ChkDate(Text6) Then
         If Val(Text6) > Val(strSrvDate(2)) Then
            MsgBox "申請案核駁日不可大於系統日 !", vbCritical
            Cancel = True
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text6
End Sub

Private Sub Text7_GotFocus()
   'Modify By Cheng 2002/07/23
'   InverseTextBox Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   'Modify By Cheng 2002/07/23
'   If KeyAscii <> 89 And KeyAscii <> 78 And KeyAscii <> 8 Then
'      KeyAscii = 0
'      Beep
'   End If
End Sub

Private Sub Text8_GotFocus()
   InverseTextBox Text8
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text9_GotFocus()
'   InverseTextBox Text9
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text9.IMEMode = 1
   OpenIme
Dim intPos As Integer
'Modify By Cheng 2002/04/22
'將游標設定在機關文號欄的"專"的後面
With Me.Text9
   If Len("" & .Text) > 0 Then
      intPos = InStr("" & .Text, "專")
      If intPos > 0 Then
         .SelStart = intPos
         .SelLength = 0
      End If
   End If
End With
End Sub

Private Sub Text9_LostFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text9.IMEMode = 2
   CloseIme
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 = "" Then
      MsgBox "機關文號不可空白 !", vbCritical
      Cancel = True
    Else
      'Modify by Morgan 2011/1/5 機關文號欄位改長度(百年問題)改抓MaxLength屬性控制
      If CheckLengthIsOK(Text9, Text9.MaxLength) = False Then
          Cancel = True
          Text9_GotFocus
      End If
   End If
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.Text12.Enabled = True Then
   Cancel = False
   Text12_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text13.Enabled = True Then
   Cancel = False
   Text13_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

For Each objTxt In Text14
   If objTxt.Enabled = True Then
      Cancel = False
      Text14_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

'Add by Morgan 2008/7/4
If CheckDueDate = False Then Exit Function

If Me.Text16.Enabled = True Then
   Cancel = False
   Text16_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text17.Enabled = True Then
   Cancel = False
   Text17_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text6.Enabled = True Then
   Cancel = False
   Text6_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text9.Enabled = True Then
   Cancel = False
   Text9_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Added by Morgan 2015/10/14
If Text14(1) <> "" Then
   If DBDATE(Text14(1)) > CompDate(1, 6, Label3(3)) Then
      MsgBox "法定期限大於來函收文日6個月!!", vbCritical
      Exit Function
   End If
End If
'end 2015/10/14

'Added by Lydia 2023/09/25 若為來函期限2次確認退回時需檢查法限是否一致
If m_strIR01 <> "" Then
   If PUB_ChkReKeyInOk(m_strIR01, m_strIR02, m_strIR03, m_strIR04, Text14(1).Text, m_bolReKeyInOK) = False Then
      Text14(1).SetFocus
      Exit Function
   End If
   If txtDelivery.Enabled = True Then
      If Trim(txtDelivery) = "" Then
         MsgBox "送達日期不可空白！", vbExclamation
         txtDelivery.SetFocus
         txtDelivery_GotFocus
         Exit Function
      Else
         If TransDate(Label3(3), 2) <> TransDate(txtDelivery, 2) Then
            MsgBox "送達日期與來函收文日不一致，請確認！", vbExclamation
            txtDelivery.SetFocus
            txtDelivery_GotFocus
            Exit Function
         End If
      End If
   End If
End If
'end 2023/09/25
   
'Add by Sindy 2021/4/27 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me) = False Then
   Exit Function
End If
'2021/4/27 END
                  
TxtValidate = True
End Function

'92.1.19 cancel by sonia
'Add By Cheng 2002/07/03
'Private Function GetPromoterNO(strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String) As String
'Dim rsA As New ADODB.Recordset
'Dim strSQLA As String
'Dim strMaxCP09 As String
'
'GetPromoterNO = ""
'strMaxCP09 = ""
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'strSQLA = "Select CP09,CP14 From CaseProgress Where CP01='" & strCP01 & "' AND CP02='" & strCP02 & "' AND CP03='" & strCP03 & "' AND CP04='" & strCP04 & "' AND (CP10='201' OR CP10='209' OR CP10='210' ) ORDER BY CP09 DESC"
'rsA.CursorLocation = adUseClient
'rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'   rsA.MoveFirst
'   strMaxCP09 = "" & rsA.Fields(0).Value
'   GetPromoterNO = "" & rsA.Fields(1).Value
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'If strMaxCP09 <> "" Then
'   strSQLA = "SELECT EP04 FROM ENGINEERPROGRESS WHERE EP02='" & strMaxCP09 & "'"
'   rsA.CursorLocation = adUseClient
'   rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsA.RecordCount > 0 Then'
'      If Not IsNull(rsA.Fields(0).Value) Then GetPromoterNO = "" & rsA.Fields(0).Value
'   End If
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
'End If
'End Function

'92.1.19 copy by sonia from frm06010602_3
Private Function GetPromoterNO(strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String) As String
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim strMaxCP09 As String
'92.1.19 modify by sonia 僅申請案號201,209,210之核稿人, 無核稿人抓承辦人,其他案件性質抓原承辦人
GetPromoterNO = m_CP14
If m_CP10 = "101" Or m_CP10 = "102" Or m_CP10 = "103" Or m_CP10 = "104" Or m_CP10 = "105" Then
   strMaxCP09 = ""
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   StrSQLa = "Select CP09,CP14 From CaseProgress Where CP01='" & strCP01 & "' AND CP02='" & strCP02 & "' AND CP03='" & strCP03 & "' AND CP04='" & strCP04 & "' AND (CP10='201' OR CP10='209' OR CP10='210' ) ORDER BY CP09 DESC"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      rsA.MoveFirst
      strMaxCP09 = "" & rsA.Fields(0).Value
      GetPromoterNO = "" & rsA.Fields(1).Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   If strMaxCP09 <> "" Then
      StrSQLa = "SELECT EP04 FROM ENGINEERPROGRESS WHERE EP02='" & strMaxCP09 & "'"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         If Not IsNull(rsA.Fields(0).Value) Then GetPromoterNO = "" & rsA.Fields(0).Value
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
End If
End Function
'Add by Morgan 2004/12/1 參考 frm0010604_3
Private Sub GetTime()
 Dim i As Integer
   If Option4(0).Value = True Then
      Text14(1) = TransDate(CompDate(2, Val(Text10), TransDate(Label3(3), 2)), 1)
      If Option1(0).Value = True Then Text14(1) = TransDate(CompDate(2, -1, TransDate(Text14(1), 2)), 1)
      If Text10 = "60" Or Text10 = "90" Then
         i = -4
      Else
         i = -2
      End If
   ElseIf Option4(1).Value = True Then
      Text14(1) = TransDate(CompDate(1, Val(Text11), TransDate(Label3(3), 2)), 1)
      If Option1(0).Value = True Then Text14(1) = TransDate(CompDate(2, -1, TransDate(Text14(1), 2)), 1)
      If Text11 = "2" Then
         i = -4
      Else
         i = -2
      End If
   End If
   If Text14(1) <> "" Then
      'Modified by Morgan 2014/11/20 外專改回舊規則
      ''Added by Morgan 2014/10/9
      'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
      '   Text14(0) = TransDate(PUB_GetOurDeadline(Text14(1)), 1)
      'Else
      ''end 2014/10/9
      
      'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
      If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
         'Modify By Sindy 2021/4/23 + m_pAgreeOnDate
         'Modify By Sindy 2023/6/7 + Text13, pa(1), pa(2), pa(3), pa(4)
         'Modify By Sindy 2025/2/12 傳入C
         Text14(0) = TransDate(PUB_GetFCPOurDeadline(Text14(1), Abs(i), , m_pAgreeOnDate, , Text13, pa(1), pa(2), pa(3), pa(4), "C"), 1)
         Text14(2) = TransDate(m_pAgreeOnDate, 1) 'Add By Sindy 2021/5/7
      Else
      'end 2019/7/11
            
         Text14(0) = TransDate(CompDate(2, i, TransDate(Text14(1), 2)), 1)
         
      End If 'Added by Morgan 2019/7/11
      'End If 'Added by Morgan 2014/10/9
      'end 2014/11/20
   End If
End Sub

'Add by Morgan 2006/4/24 設定來函期限&承辦期限(從ReadPatent抽出來)
Private Sub SetDeadline()
   Dim stCPM02 As String
   If Text13 = 抗告 Then
      stCPM02 = m_NewCP10
   'Add by Morgan 2008/7/4 分割案核駁且母案有提再審時抓再審的來函期限
   ElseIf m_b307Plus107 = True Then
      stCPM02 = "107"
   Else
      stCPM02 = m_CP10
   End If
   Dim i As Integer
   strExc(0) = "SELECT CPM07,CPM08,CPM09 FROM CASEPROPERTYMAP WHERE CPM01='" & pa(1) & "' AND CPM02='" & stCPM02 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
      If intI = 1 Then
         If Not IsNull(.Fields(1)) Then
            Option4(0).Value = True
            Text10 = .Fields(1)
            Text14(1) = TransDate(CompDate(2, .Fields(1), TransDate(Label3(3).Caption, 2)), 1)
         ElseIf Not IsNull(.Fields(2)) Then
            Option4(1).Value = True
            Text11 = .Fields(2)
            Text14(1) = TransDate(CompDate(1, .Fields(2), TransDate(Label3(3).Caption, 2)), 1)
         Else
            Text10 = ""
            Text11 = ""
            Option4(0).Value = True
         End If
         If Text14(1) <> "" And Not IsNull(.Fields(0)) Then
            If .Fields(0) = "1" Then
               Option1(0).Value = True
               Text14(1) = TransDate(CompDate(2, -1, TransDate(Text14(1), 2)), 1)
            Else
               Option1(1).Value = True
            End If
         End If
         If Not IsNull(.Fields(1)) Then
            If .Fields(1) = 60 Or .Fields(1) = 90 Then
               i = -4
            Else
               i = -2
            End If
         ElseIf Not IsNull(.Fields(2)) Then
            If .Fields(2) = 2 Then
               i = -4
            Else
               i = -2
            End If
         End If
         If Text14(1) <> "" Then
            'Modified by Morgan 2014/11/20 外專改回舊規則
            ''Added by Morgan 2014/10/9
            'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
            '   Text14(0) = TransDate(PUB_GetOurDeadline(Text14(1)), 1)
            'Else
            ''end 2014/10/9
            
            'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
            If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
               'Modify By Sindy 2021/4/23 + m_pAgreeOnDate
               'Modify By Sindy 2023/6/7 + Text13, pa(1), pa(2), pa(3), pa(4)
               'Modify By Sindy 2025/2/12 傳入C
               Text14(0) = TransDate(PUB_GetFCPOurDeadline(Text14(1), Abs(i), , m_pAgreeOnDate, , Text13, pa(1), pa(2), pa(3), pa(4), "C"), 1)
               Text14(2) = TransDate(m_pAgreeOnDate, 1) 'Add By Sindy 2021/5/7
            Else
            'end 2019/7/11
      
               Text14(0) = TransDate(CompDate(2, i, TransDate(Text14(1), 2)), 1)
               
            End If 'Added by Morgan 2019/7/11
            'End If 'Added by Morgan 2014/10/9
            'end 2014/11/20
         End If
      End If
   End With

End Sub

'Added by Lydia 2023/09/25
Private Sub txtDelivery_GotFocus()
   TextInverse txtDelivery
End Sub

Private Sub txtDelivery_Validate(Cancel As Boolean)
   If Trim(txtDelivery) <> "" Then
      If Not ChkDate(txtDelivery) Then
         txtDelivery.SetFocus
         txtDelivery_GotFocus
         Cancel = True
      End If
   End If
End Sub
