VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03020606_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-變更,移轉,授權"
   ClientHeight    =   4632
   ClientLeft      =   72
   ClientTop       =   996
   ClientWidth     =   8676
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4632
   ScaleWidth      =   8676
   Begin VB.Frame Frame2 
      Caption         =   "附送書件"
      Height          =   1125
      Left            =   7020
      TabIndex        =   56
      Top             =   2730
      Visible         =   0   'False
      Width           =   1545
      Begin VB.CheckBox chkAtt1 
         Caption         =   "授權契約書"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Tag             =   ".license.pdf"
         Top             =   735
         Width           =   1400
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "委任書"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Tag             =   ".poa.pdf"
         Top             =   495
         Value           =   1  '核取
         Width           =   1215
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "基本資料表"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Tag             =   ".contact.pdf"
         Top             =   240
         Value           =   1  '核取
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1845
      Left            =   2730
      TabIndex        =   45
      Top             =   2730
      Width           =   4215
      Begin MSForms.Label lblCName 
         Height          =   300
         Index           =   4
         Left            =   2100
         TabIndex        =   55
         Top             =   1493
         Width           =   2000
         Caption         =   "Form2.0"
         Size            =   "3528;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCName 
         Height          =   300
         Index           =   3
         Left            =   2100
         TabIndex        =   54
         Top             =   1169
         Width           =   2000
         Caption         =   "Form2.0"
         Size            =   "3528;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCName 
         Height          =   300
         Index           =   2
         Left            =   2100
         TabIndex        =   53
         Top             =   847
         Width           =   2000
         Caption         =   "Form2.0"
         Size            =   "3528;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCName 
         Height          =   300
         Index           =   1
         Left            =   2100
         TabIndex        =   52
         Top             =   525
         Width           =   2000
         Caption         =   "Form2.0"
         Size            =   "3528;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCName 
         Height          =   300
         Index           =   0
         Left            =   2100
         TabIndex        =   51
         Top             =   203
         Width           =   2000
         Caption         =   "Form2.0"
         Size            =   "3528;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFM2 
         Height          =   300
         Index           =   4
         Left            =   990
         TabIndex        =   9
         Top             =   1470
         Width           =   1095
         VariousPropertyBits=   679495707
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFM2 
         Height          =   300
         Index           =   3
         Left            =   990
         TabIndex        =   8
         Top             =   1140
         Width           =   1095
         VariousPropertyBits=   679495707
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFM2 
         Height          =   300
         Index           =   2
         Left            =   990
         TabIndex        =   7
         Top             =   825
         Width           =   1095
         VariousPropertyBits=   679495707
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFM2 
         Height          =   300
         Index           =   1
         Left            =   990
         TabIndex        =   6
         Top             =   495
         Width           =   1095
         VariousPropertyBits=   679495707
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFM2 
         Height          =   300
         Index           =   0
         Left            =   990
         TabIndex        =   5
         Top             =   180
         Width           =   1095
         VariousPropertyBits=   679495707
         MaxLength       =   9
         Size            =   "1940;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblC 
         Caption         =   "申請人5:"
         Height          =   165
         Index           =   4
         Left            =   120
         TabIndex        =   50
         Top             =   1538
         Width           =   885
      End
      Begin VB.Label lblC 
         Caption         =   "申請人4:"
         Height          =   165
         Index           =   3
         Left            =   120
         TabIndex        =   49
         Top             =   1214
         Width           =   885
      End
      Begin VB.Label lblC 
         Caption         =   "申請人3:"
         Height          =   165
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   892
         Width           =   885
      End
      Begin VB.Label lblC 
         Caption         =   "申請人2:"
         Height          =   165
         Index           =   1
         Left            =   120
         TabIndex        =   47
         Top             =   570
         Width           =   885
      End
      Begin VB.Label lblC 
         Caption         =   "申請人1:"
         Height          =   165
         Index           =   0
         Left            =   120
         TabIndex        =   46
         Top             =   255
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7770
      TabIndex        =   2
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5820
      TabIndex        =   0
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6660
      TabIndex        =   1
      Top             =   45
      Width           =   1080
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm03020606_1.frx":0000
      Left            =   1260
      List            =   "frm03020606_1.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   17
      Top             =   859
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1020
      MaxLength       =   3
      TabIndex        =   13
      Top             =   210
      Width           =   550
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1575
      MaxLength       =   6
      TabIndex        =   14
      Top             =   210
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2415
      MaxLength       =   1
      TabIndex        =   15
      Top             =   210
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2655
      MaxLength       =   2
      TabIndex        =   16
      Top             =   210
      Width           =   375
   End
   Begin MSForms.CheckBox CheckBox2 
      Height          =   300
      Left            =   5304
      TabIndex        =   61
      Top             =   2400
      Width           =   1836
      BackColor       =   -2147483633
      ForeColor       =   16711680
      DisplayStyle    =   4
      Size            =   "3238;529"
      Value           =   "0"
      Caption         =   "變更地址"
      FontName        =   "新細明體"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   1200
      TabIndex        =   60
      Top             =   3150
      Width           =   1500
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;980"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textFM2_Start 
      Height          =   315
      Left            =   4080
      TabIndex        =   4
      Top             =   2100
      Visible         =   0   'False
      Width           =   975
      VariousPropertyBits=   679495707
      MaxLength       =   7
      Size            =   "1720;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textFM2_09 
      Height          =   315
      Left            =   1200
      TabIndex        =   59
      Top             =   2760
      Width           =   975
      VariousPropertyBits=   679495705
      MaxLength       =   7
      Size            =   "1720;556"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox textFM2_05 
      Height          =   315
      Left            =   1200
      TabIndex        =   58
      Top             =   2430
      Visible         =   0   'False
      Width           =   975
      VariousPropertyBits=   679495707
      MaxLength       =   7
      Size            =   "1720;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      Caption         =   "授權起始日期："
      Height          =   195
      Left            =   2760
      TabIndex        =   57
      Top             =   2160
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSForms.CheckBox CheckBox1 
      Height          =   225
      Left            =   4410
      TabIndex        =   3
      Top             =   2445
      Width           =   615
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "1085;397"
      Value           =   "0"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "是否要變更申請人:"
      Height          =   180
      Left            =   2760
      TabIndex        =   44
      Top             =   2460
      Width           =   1635
   End
   Begin MSForms.Label lblData 
      Height          =   255
      Index           =   10
      Left            =   7560
      TabIndex        =   43
      Top             =   570
      Width           =   1065
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "1879;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "商標種類:"
      Height          =   180
      Left            =   6750
      TabIndex        =   42
      Top             =   570
      Width           =   765
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   8600
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   8600
      Y1              =   2355
      Y2              =   2355
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   9
      Left            =   4710
      TabIndex        =   41
      Top             =   1800
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "3149;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   8
      Left            =   1260
      TabIndex        =   40
      Top             =   1800
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "3149;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   7
      Left            =   4710
      TabIndex        =   39
      Top             =   1485
      Width           =   3615
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "6376;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   6
      Left            =   1260
      TabIndex        =   38
      Top             =   1485
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "3149;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   5
      Left            =   4710
      TabIndex        =   37
      Top             =   1172
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "3149;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   4
      Left            =   1260
      TabIndex        =   36
      Top             =   1172
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "3149;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   35
      Top             =   859
      Width           =   6615
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "11668;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   255
      Index           =   2
      Left            =   4710
      TabIndex        =   34
      Top             =   570
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "3149;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   255
      Index           =   1
      Left            =   1260
      TabIndex        =   33
      Top             =   570
      Width           =   1665
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "2937;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   0
      Left            =   4710
      TabIndex        =   32
      Top             =   233
      Width           =   1035
      VariousPropertyBits=   27
      Caption         =   "Form2.0"
      Size            =   "1826;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "規費 :                                      元"
      Height          =   180
      Left            =   180
      TabIndex        =   31
      Top             =   2835
      Width           =   2340
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人 :"
      Height          =   180
      Left            =   180
      TabIndex        =   30
      Top             =   3180
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期 :"
      Height          =   180
      Left            =   180
      TabIndex        =   29
      Top             =   2535
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   3810
      TabIndex        =   28
      Top             =   255
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   3810
      TabIndex        =   27
      Top             =   1485
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   210
      TabIndex        =   26
      Top             =   1485
      Width           =   945
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3810
      TabIndex        =   25
      Top             =   1172
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   210
      TabIndex        =   24
      Top             =   1172
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   210
      TabIndex        =   23
      Top             =   210
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   210
      TabIndex        =   22
      Top             =   570
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "審定號數:"
      Height          =   180
      Left            =   3810
      TabIndex        =   21
      Top             =   570
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "商標名稱:"
      Height          =   180
      Left            =   210
      TabIndex        =   20
      Top             =   859
      Width           =   765
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   0
      Left            =   3810
      TabIndex        =   19
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Left            =   210
      TabIndex        =   18
      Top             =   1800
      Width           =   765
   End
End
Attribute VB_Name = "frm03020606_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/02/05 申請書日期和規費改成Form2.0 ; (Text5=>textFM2_05, Text9=>textFM2_09)
'Memo by Lydia 2020/10/16 改成Form2.0 (lblData、lstNameAgent、lblC、textFM2、lblCName)
'Create by Lydia 2020/10/16 各式申請書:變更, 移轉
Option Explicit
Dim tm() As String '商標基本檔
Dim intWhere As Integer, intLastRow As Integer
Dim strReceiveNo As String '收文號
Dim m_CP110 As String, m_AgentName As String  '出名代理人
Dim m_CP10 As String '案件性質
Dim m_CP17 As String '收文規費
Dim m_CP118  As String '是否電子送件
Dim m_CaseNo As String '電子送件-本所案號
Dim m_F21st07 As String 'FCT程序分機

Private Sub CheckBox1_Click()
    If CheckBox1.Value = True Then
        Frame1.Enabled = True
    Else
        Frame1.Enabled = False
    End If
    For intI = 0 To 4
        textFM2(intI).Text = ""
    Next intI

End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim bolChk As Boolean
Dim i As Integer
Dim strFolder As String, strFileName As String
Dim ET01 As String, ET03 As String, ET03_1 As String
Dim strContent As String
Dim strKind As String
Dim m_Input As String 'Added by Lydia 2023/07/05

   Select Case Index
      Case 0 '確定
         
         If TxtValidate = False Then Exit Sub
         'Added by Lydia 2023/07/05
         If m_CP118 = "" And (m_CP10 = "501" Or m_CP10 = "301") Then
            If MsgBox("是否為「一文多案」申請書？", vbYesNo, "詢問") = vbYes Then
JumpToRe:
               m_Input = InputBox("請輸入一文多案的總件數：", "一文多案輸入件數", Val(m_Input))
               If Val(m_Input) <= 0 Then
                  If MsgBox("輸入件數為" & Val(m_Input) & "，是否為「一文多案」申請書？", vbInformation + vbYesNo + vbDefaultButton1, "詢問") = vbYes Then
                     GoTo JumpToRe
                  End If
               Else
                  'Modify By Sindy 2023/7/17 規費金額抓畫面上的
                  'm_Input = "M" & (Val(m_Input) * 2000)
                  m_Input = "M" & (Val(m_Input) * Val(Format(textFM2_09, "###0")))
                  '2023/7/17 END
               End If
            End If
         End If
         'end 2023/07/05
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         If m_CP118 = "" Then  '紙本
            If m_CP10 = "501" Then
                strKind = "A1" '移轉=附中文譯本勾選
            ElseIf m_CP10 = "301" And tm(15) = "" Then
                strKind = "A1,B2" '註冊前變更=委任書勾選,具結書不勾選
            ElseIf m_CP10 = "301" And tm(15) <> "" Then
                strKind = "A1" '註冊變更=委任書勾選
            End If
            'Modified by Lydia 2023/07/05 傳入一文多案的規費 +m_Input
            If PUB_GetApplBook(Text1 & "-" & Text2 & "-" & Text3 & "-" & Text4, m_CP10, _
            IIf(textFM2(0).Enabled = True And Trim(textFM2(0)) <> "", textFM2(0), ""), _
            IIf(textFM2(1).Enabled = True And Trim(textFM2(1)) <> "", textFM2(1), ""), _
            IIf(textFM2(2).Enabled = True And Trim(textFM2(2)) <> "", textFM2(2), ""), _
            IIf(textFM2(3).Enabled = True And Trim(textFM2(3)) <> "", textFM2(3), ""), _
            IIf(textFM2(4).Enabled = True And Trim(textFM2(4)) <> "", textFM2(4), ""), strReceiveNo, strKind, m_Input) = True Then
               Call cmdOK_Click(3)
            End If
         Else
            m_CaseNo = PUB_FCPCaseNo2FileName(tm(1), tm(2), tm(3), tm(4))
            '桌面上建立案號資料夾
            strFolder = PUB_Getdesktop
            strFolder = strFolder & "\" & m_CaseNo
            If Dir(strFolder, vbDirectory) = "" Then
                MkDir strFolder
            End If
            
            strLetterDate = textFM2_05.Text
            
            ET01 = "90" '定稿別
            ET03 = ""
            If m_CP10 = "301" Then '變更
                'Modified by Lydia 2022/12/21 註冊以核准+審定號為準
                'If tm(15) = "" Then '註冊前變更
                If Not (tm(16) = "1" And tm(15) <> "") Then '註冊前變更
                    'Modified by Lyda 2021/01/21 變更處理狀況: 小於20保留給補正申請書用; ex.FCT-46637的(紙本)補正->移轉
                    'ET03 = "01"     '申請書
                    'ET03_1 = "00"  '基本資料表(註冊前變更)
                    ET03 = "21"     '申請書
                    ET03_1 = "20"  '基本資料表(註冊前變更)
                    strFileName = strFolder & "\" & m_CaseNo & ".註冊前變更申請書"
                Else                    '註冊變更
                    'Modified by Lyda 2021/01/21 變更處理狀況
                    'ET03 = "02"     '申請書
                    ET03 = "22"     '申請書
                    'Modified by Lydia 2020/11/23 基本資料表不同一般,增加【身分類別】
                    'ET03_1 = "03"  '基本資料表(一般)
                    'Modified by Lyda 2021/01/21 變更處理狀況:
                    'ET03_1 = "04"  '基本資料表(一般)
                    ET03_1 = "24"  '基本資料表(一般)
                    strFileName = strFolder & "\" & m_CaseNo & ".註冊變更申請書"
                End If
            ElseIf m_CP10 = "501" Then '移轉
                    'Modified by Lyda 2021/01/21 變更處理狀況: 小於20保留給補正申請書用; ex.FCT-46637的(紙本)補正->移轉
                    'ET03 = "01"     '申請書
                    'ET03_1 = "02"  '基本資料表(移轉)
                    ET03 = "21"     '申請書
                    ET03_1 = "22"  '基本資料表(移轉)
                    strFileName = strFolder & "\" & m_CaseNo & ".移轉登記申請書"
            'Added by Lydia 2021/02/05
            ElseIf m_CP10 = "502" Then '授權
                    ET03 = "21"     '申請書
                    ET03_1 = "22"  '基本資料表(授權)
                    strFileName = strFolder & "\" & m_CaseNo & ".授權登記申請書"
            End If
            
            If ET03 <> "" Then
                 '申請書
                 If StartLetter2(m_CaseNo, ET01, ET03, strReceiveNo, "2") = False Then Exit Sub
                 NowPrint strReceiveNo, ET01, ET03, False, strUserNum, , strContent, True, strContent
            End If
            
            'Added by Lydia 2021/02/05 判斷勾選項是否要含基本資料表
            If chkAtt1(0).Value = False Then
                Call PUB_MakeDoc(strContent, strFileName)
            Else
            'end 2021/02/05
                '基本資料表
                If ET03_1 <> "" Then
                    If StartLetter2(m_CaseNo, ET01, ET03_1, strReceiveNo, "1") = False Then Exit Sub
                    NowPrint strReceiveNo, ET01, ET03_1, False, strUserNum, , strContent, True, strContent
                End If
                strContent = Replace(strContent, vbCrLf & Chr(12), vbCrLf & "|#(分節)#|")    '換頁符號Chr(12)替換為分節符號 "|#(分節)#|"
                Call PUB_MakeDoc(strContent, strFileName, , , , , True)  '分節處理頁碼
            End If 'Added by Lydia 2021/02/05
         End If
         
         frm030206_1.Show
         '回到原畫面要清除畫面
         frm030206_1.ClearForm
         
      Case 1 '回前畫面
         frm030206_1.Show
   End Select
   Unload Me
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         lblData(3) = tm(5)
      Case "英"
         lblData(3) = tm(6)
      Case "日"
         lblData(3) = tm(7)
   End Select
End Sub

Private Sub Form_Load()
Dim tKind As String '特殊申請書

   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm030206_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      tKind = .Text6
      strReceiveNo = .Tag
      If tKind = "2" Then m_CP118 = "Y"
   End With
   ReDim tm(TF_TM)
   ReadTradeMark
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   PUB_SetOurAgent lstNameAgent, tm(), m_CP110, m_CP10, True
   'Added by Lydia 2021/04/20 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1300
   lstNameAgent.Width = 1300
   
   Combo1.ListIndex = 0
   textFM2_05.Text = strSrvDate(2)
    
   If m_CP10 = "301" Then '變更
       Frame1.Enabled = False
       Label10.Caption = "是否要變更申請人:"
       CheckBox1.Visible = True
       CheckBox2.Visible = True 'Added by Lydia 2023/11/08
   ElseIf m_CP10 = "501" Then '移轉
       Frame1.Enabled = True
       Label10.Caption = "請輸入受讓人:"
       CheckBox1.Visible = False
       CheckBox2.Visible = False 'Added by Lydia 2023/11/08
   'Added by Lydia 2021/02/05
   ElseIf m_CP10 = "502" Then '授權
       Frame1.Enabled = True
       Frame2.Visible = True
       Label10.Visible = False
       CheckBox1.Visible = False
       CheckBox2.Visible = False 'Added by Lydia 2023/11/08
       Label12.Visible = True: Label12.Top = 2040
       textFM2_Start.Visible = True:   textFM2_Start.Top = 1980
       For intI = 0 To 4
           lblC(intI).Caption = "被授權人" & intI + 1 & ":"
       Next intI
   'end 2021/02/05
   End If
   
   For intI = 0 To 4
       textFM2(intI).Text = ""
       textFM2(intI).Tag = ""
       lblCName(intI).Caption = ""
   Next intI
       
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm03020606_1 = Nothing
End Sub

Private Sub ReadTradeMark()
Dim rsRd As New ADODB.Recordset
Dim Lbl
   
   For Each Lbl In lblData
      Lbl.Caption = ""
   Next
   tm(1) = Text1
   tm(2) = Text2
   tm(3) = Text3
   tm(4) = Text4
   If ClsPDReadTrademarkDatabase(tm(), intWhere) Then
      textFM2_05 = tm(11)
      lblData(1) = tm(12)
      lblData(2) = tm(15)
      lblData(3) = tm(5)
   End If
   '商標種類
   If ClsPDGetPatentTrademarkKind(商標, tm(8), strExc(0), False) = 1 Then
       lblData(10) = strExc(0)
   End If
   
   strExc(0) = "select cpm03,s1.st02 as st1,s2.st02 as st2,cp43,cp10,cp05,cp06,cp07,cp84,cp110,cp64,cp27,cp17,s3.st07  " & _
                    "from caseprogress,casepropertymap,staff s1 ,staff s2,staff s3 " & _
                    "where cp09='" & strReceiveNo & "' " & _
                    "and cp01=cpm01(+) and cp10=cpm02(+) and cp14=s1.st01(+) " & _
                    "and cp13=s2.st01(+) and s2.st57=s3.st01(+) "
   intI = 1
   Set rsRd = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
       With rsRd
          m_CP110 = "" & .Fields("CP110")
          m_CP10 = "" & .Fields("CP10")
          m_CP17 = Val("" & .Fields("cp17"))
          '收文規費
          If m_CP118 <> "" Then '電子送件的規費有千分位,會造成轉檔錯誤
               textFM2_09.Text = Val(m_CP17)
          Else
               textFM2_09.Text = Format(Val("" & .Fields("cp17")), "#,##0")
          End If
          
          lblData(0) = "" & .Fields("cpm03") '案件性質
          lblData(4) = "" & .Fields("st1") '承辦人
          lblData(5) = "" & .Fields("st2") '智權人員
          m_F21st07 = "" & .Fields("st07")
          If "" & .Fields("cp43") <> "" Then
             strExc(0) = "SELECT * FROM CASEPROGRESS WHERE CP09='" & .Fields("cp43") & "'"
             intI = 1
             Set rsRd = ClsLawReadRstMsg(intI, strExc(0))
             If intI = 1 Then
                lblData(6) = TransDate("" & rsRd.Fields("CP05"), 1) '來函收文日
                lblData(7) = "" & rsRd.Fields("CP08") '機關文號
             End If
          End If
          lblData(8) = TransDate("" & .Fields("cp06"), 1) '本所期限
          lblData(9) = TransDate("" & .Fields("cp07"), 1) '法定期限
       End With
   End If

   'FCT向智慧局提出之各式申請書上之分機號碼，請將日本區設定為011國家檔管制人分機
   strExc(0) = "select fa10,st07 from fagent, nation, staff where fa01||fa02='" & ChangeCustomerL(tm(44)) & "' and substr(fa10,1,3)=na01(+) and na55=st01(+) "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Left("" & RsTemp.Fields("fa10"), 3) = "011" Then
         m_F21st07 = "" & RsTemp.Fields("st07")
      End If
   End If
   
   Set rsRd = Nothing
End Sub

Private Sub textfm2_05_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(textFM2_05.Text)
   If Cancel = True Then TextInverse textFM2_05
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   
   For intI = 0 To 4
       If textFM2(intI).Text <> "" And lblCName(intI).Caption = "" Then
            MsgBox "申請人代碼<" & textFM2(intI).Text & ">不正確", vbExclamation
            textFM2(intI).SetFocus
            textFM2_GotFocus intI
            Exit Function
       End If
   Next intI
   
   If m_CP10 = "301" And CheckBox1.Value = True And Trim(textFM2(0)) = "" Then
       MsgBox "請輸入欲變更之申請人 !!!"
       textFM2(0).SetFocus
       textFM2_GotFocus 0
       Exit Function
   End If
   If m_CP10 = "501" And textFM2(0).Text = "" Then
        MsgBox "請輸入受讓人 !!!"
        textFM2(0).SetFocus
        textFM2_GotFocus 0
        Exit Function
   End If
   'Added by Lydia 2021/02/05
   If m_CP10 = "502" Then
      If Trim(textFM2_Start) = "" Then
          MsgBox "請輸入授權起始日期 !!!"
          textFM2_Start.SetFocus
          Exit Function
      End If
      If textFM2(0) = "" Then
          MsgBox "請輸入被授權人 !!!"
          textFM2(0).SetFocus
          Exit Function
      End If
   End If
   'end 2021/02/05
   
   '申請人必須依序輸入
   If textFM2(0) <> "" Or textFM2(1) <> "" Or textFM2(2) <> "" Or textFM2(3) <> "" Or textFM2(4) <> "" Then
      If (Trim(textFM2(1)) <> "" And Trim(textFM2(0)) = "") Or _
         (Trim(textFM2(2)) <> "" And Trim(textFM2(1)) = "") Or _
         (Trim(textFM2(3)) <> "" And Trim(textFM2(2)) = "") Or _
         (Trim(textFM2(4)) <> "" And Trim(textFM2(3)) = "") Then
         MsgBox "請依序輸入！", vbExclamation
         Exit Function
      End If
      If (textFM2(1) <> "" And Trim(textFM2(1)) = Trim(textFM2(0))) Or _
         (textFM2(2) <> "" And Trim(textFM2(2)) = Trim(textFM2(1))) Or _
         (textFM2(3) <> "" And Trim(textFM2(3)) = Trim(textFM2(2))) Or _
         (textFM2(4) <> "" And Trim(textFM2(4)) = Trim(textFM2(3))) Then
         MsgBox "資料重覆！", vbExclamation
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   m_CP110 = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'Modified by Lydia 2021/04/23 改模組
         'm_CP110 = m_CP110 & "," & PUB_GetListBoxTagVal(lstNameAgent, ii)
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         Cancel = False
      End If
   Next
   If Cancel = True Then
      MsgBox "出名代理人不可空白！", vbExclamation
   Else
      If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
      m_AgentName = Mid(m_AgentName, 2)
   End If
End Sub

'電子送件-申請書
Private Function StartLetter2(ByVal iCaseNo As String, ByVal iET01 As String, ByVal iET03 As String, ByVal iCp09 As String, ByVal iKind As String) As Boolean
   
   Dim strTxt(1 To 30) As String
   Dim ii As Integer, jj As Integer
   Dim tmpArr1 As Variant, tmpArr2 As Variant
   Dim intA As Integer
   Dim bolReadTM As Boolean 'Added by Lydia 2023/11/08
   
   EndLetter iET01, iCp09, iET03, strUserNum
   
   ii = 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
   
   '申請人資料
   strExc(0) = ""
    If textFM2(0).Enabled = True And textFM2(0).Text <> "" Then strExc(0) = strExc(0) & "@" & ChangeCustomerL(textFM2(0).Text)
    If textFM2(1).Enabled = True And textFM2(1).Text <> "" Then strExc(0) = strExc(0) & "@" & ChangeCustomerL(textFM2(1).Text)
    If textFM2(2).Enabled = True And textFM2(2).Text <> "" Then strExc(0) = strExc(0) & "@" & ChangeCustomerL(textFM2(2).Text)
    If textFM2(3).Enabled = True And textFM2(3).Text <> "" Then strExc(0) = strExc(0) & "@" & ChangeCustomerL(textFM2(3).Text)
    If textFM2(4).Enabled = True And textFM2(4).Text <> "" Then strExc(0) = strExc(0) & "@" & ChangeCustomerL(textFM2(4).Text)
    If strExc(0) <> "" Then strExc(0) = Mid(strExc(0), 2)
   'Modified by Lydia 2023/11/08
   'Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, m_CP10, tm(), False, strExc(0), , tm(1))
   '原本預設抓申請人基本檔之地址;現在改成預設抓案件申請人資料之地址，當勾選「變更地址」，基本資料表之申請人地址改抓申請人基本檔之地址。
   'Modified by Lydia 2023/12/29 單純勾選「變更地址」，只有地址改抓申請人基本檔之地址。ex. FCT-033079
   'If CheckBox1.Value = True Or CheckBox2.Value = True Then
   strExc(1) = ""
   If CheckBox2.Value = True Then
      strExc(1) = strExc(1) & ",申請人地址"
   End If
   If CheckBox1.Value = True Then
      strExc(1) = ""
   'end 2023/12/29
      bolReadTM = False
   Else
      bolReadTM = True
   End If
   If m_CP10 <> "301" Then bolReadTM = False 'Added by Lydia 2023/11/09 排除非變更案; 移轉,授權也要輸入申請人
   'Modified by Lydia 2023/12/29 +指定讀取申請人的資料
   'Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, m_CP10, tm(), bolReadTM, strExc(0), , tm(1))
   Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, m_CP10, tm(), bolReadTM, strExc(0), , tm(1), , strExc(1))
   'end 2023/11/08
   
   '出名代理人: 改成共用模組取得資料
   strExc(0) = PUB_GetAgentCP110(iCp09, m_CP110, "FCT", "4")
   If strExc(0) <> "" Then
       tmpArr1 = Empty
       tmpArr1 = Split(strExc(0), "|")
       'Added by Lydia 2021/02/05
       strExc(1) = "代理人"
       If m_CP10 = "502" Then
           strExc(1) = "授權人之代理人"
           strExc(2) = "被授權人之代理人"
       End If
       'end 2021/02/05
       For jj = 0 To UBound(tmpArr1)
           If Trim(tmpArr1(jj)) <> "" Then
               tmpArr2 = Empty
               tmpArr2 = Split(tmpArr1(jj), ",")
                ii = ii + 1
                'Modified by Lydia 2021/02/05 代理人=>改用變數strExc(1)
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','" & strExc(1) & jj + 1 & "-證書字號','" & tmpArr2(0) & "')"
                ii = ii + 1
                'Modified by Lydia 2021/02/05 代理人=>改用變數strExc(1)
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','" & strExc(1) & jj + 1 & "-ID','" & tmpArr2(1) & "')"
                ii = ii + 1
                'Modified by Lydia 2021/02/05 代理人=>改用變數strExc(1)
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','" & strExc(1) & jj + 1 & "-中文姓名','" & PUB_ConvertNameFormat("" & tmpArr2(2)) & "')"
                'Added by Lydia 2021/02/05 授權502，預設被授權人之代理人=授權人之代理人
                If m_CP10 = "502" Then
                    ii = ii + 1
                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','" & strExc(2) & jj + 1 & "-證書字號','" & tmpArr2(0) & "')"
                    ii = ii + 1
                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','" & strExc(2) & jj + 1 & "-ID','" & tmpArr2(1) & "')"
                    ii = ii + 1
                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','" & strExc(2) & jj + 1 & "-中文姓名','" & PUB_ConvertNameFormat("" & tmpArr2(2)) & "')"
                End If
                'end 2021/02/05
           End If
       Next jj
   End If
   
   If iKind = "1" Then '基本資料表
        ii = ii + 1
        'FCT程序分機
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','FCT程序分機','" & m_F21st07 & "')"
   End If
   
   If iKind = "2" Then '電子送件申請書
        'Added by Lydia 2021/02/05 授權502：商品服務名稱
        If m_CP10 = "502" Then
            strExc(1) = "": strExc(2) = "": strExc(3) = ""
            strExc(0) = BeforePrintGetDBData("TMGoods:" & tm(1) & "-" & tm(2) & "-" & tm(3) & "-" & tm(4) & "-||區隔", True)
            If Trim(strExc(0)) <> "" Then
                tmpArr1 = Empty
                tmpArr1 = Split(strExc(0), "||")
                jj = 1
                For intA = 0 To UBound(tmpArr1)
                    strExc(1) = Trim(tmpArr1(intA))
                    If strExc(1) <> "" Then
                        strExc(2) = strExc(2) & _
                                         "【部分授權" & jj & "】  " & vbCrLf & _
                                         "　　【類別】　　　　　　　　　" & Mid(strExc(1), 1, InStr(strExc(1), "：") - 1) & vbCrLf & _
                                         "　　【商品服務名稱】　　　　　" & Mid(strExc(1), InStr(strExc(1), "：") + 1) & vbCrLf
                        jj = jj + 1
                    End If
                Next intA
            ElseIf tm(9) <> "" Then
                 tmpArr1 = Empty
                 tmpArr1 = Split(tm(9), ",")
                 jj = 1
                 For intA = 0 To UBound(tmpArr1)
                     strExc(1) = Trim(tmpArr1(intA))
                     If strExc(1) <> "" Then
                          strExc(2) = strExc(2) & _
                                           "【部分授權" & jj & "】  " & vbCrLf & _
                                           "　　【類別】　　　　　　　　　" & strExc(1) & vbCrLf & _
                                           "　　【商品服務名稱】　　　　　" & vbCrLf
                          jj = jj + 1
                     End If
                 Next intA
            Else
                     strExc(2) = "【部分授權1】  " & vbCrLf & _
                                      "　　【類別】　　　　　　　　　" & vbCrLf & _
                                      "　　【商品服務名稱】　　　　　" & vbCrLf
            End If
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','部分授權','" & ChgSQL(strExc(2)) & "')"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','授權起始日期','" & ChangeTStringToTDateString(textFM2_Start) & "')"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','授權終止日期','" & ChangeTStringToTDateString(tm(22)) & "')"
        End If
        'end 2021/02/05
        ii = ii + 1
        '繳費金額
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','繳費金額','" & textFM2_09.Text & "')"
        '附送書件
        'Modified by Lydia 2021/02/05 改成核取項
        ' ii = ii + 1
        ' strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        '    " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-基本資料表', '" & iCaseNo & ".contact.pdf')"
        ' ii = ii + 1
        ' strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        '    " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-委任書', '" & iCaseNo & ".poa.pdf')"
        For intI = 0 To 2
             If chkAtt1(intI).Value = 1 Then
                 ii = ii + 1
                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                    " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-" & chkAtt1(intI).Caption & "', '" & m_CaseNo & chkAtt1(intI).Tag & "')"
             End If
        Next intI
        'end 2021/02/05
   End If
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function

Private Function FormSave() As Boolean
Dim strSqlText As String
Dim rsA As New ADODB.Recordset

On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
   '出名代理人
   If lstNameAgent.Visible = True Then
      strSql = " UPDATE CASEPROGRESS SET "
      If strSqlText = "" Then
         strSqlText = " cp110 = " & CNULL(m_CP110)
      Else
         strSqlText = strSqlText & " ,cp110 = " & CNULL(m_CP110)
      End If
      strSql = strSql & strSqlText & " WHERE CP09='" & strReceiveNo & "'"
      cnnConnection.Execute strSql
   End If
   '預設為電子送件
   If m_CP118 = "Y" Then
      '目前FCT尚未自動扣款
      strSql = " UPDATE CASEPROGRESS SET CP118='Y' WHERE CP09='" & strReceiveNo & "' AND CP158=0 AND CP118 IS NULL"
      cnnConnection.Execute strSql, intI
   End If
   
   '更新變更事項檔
   If CheckBox1.Visible = True And CheckBox1.Value = True Then
      '檢查是否有此筆文號變更資料
      strExc(0) = "select ce01 from changeevent where ce01='" & strReceiveNo & "'"
      intI = 1
      Set rsA = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strSql = "update changeevent" & _
                  " set ce04='" & textFM2(0) & "'" & _
                  " ,ce05='" & textFM2(1) & "'" & _
                  " ,ce06='" & textFM2(2) & "'" & _
                  " ,ce07='" & textFM2(3) & "'" & _
                  " ,ce08='" & textFM2(4) & "'" & _
                  " WHERE CE01='" & strReceiveNo & "'"
      Else
         strSql = "insert into changeevent(ce01,ce04,ce05,ce06,ce07,ce08) values(" & _
                  CNULL(strReceiveNo) & "," & CNULL(textFM2(0)) & "," & CNULL(textFM2(1)) & "," & _
                  CNULL(textFM2(2)) & "," & CNULL(textFM2(3)) & "," & CNULL(textFM2(4)) & ")"
      End If
      cnnConnection.Execute strSql
   End If
   Set rsA = Nothing
   cnnConnection.CommitTrans
   FormSave = True
   
ErrorHandler:
   If Err.Number <> 0 Then
        cnnConnection.RollbackTrans
   End If
End Function

Private Sub textFM2_GotFocus(Index As Integer)
    TextInverse textFM2(Index)
End Sub

Private Sub textFM2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFM2_Validate(Index As Integer, Cancel As Boolean)
Dim strTemp As String
   
   If textFM2(Index).Tag <> textFM2(Index).Text Then
       lblCName(Index).Caption = ""
       If textFM2(Index).Text <> "" Then
           textFM2(Index).Text = ChangeCustomerL(textFM2(Index).Text)
           strTemp = GetCustomerName(textFM2(Index).Text)
           If strTemp = "" Then
               MsgBox "申請人代碼<" & textFM2(Index).Text & ">不正確", vbExclamation
               Cancel = True
           Else
               lblCName(Index).Caption = strTemp
           End If
           textFM2(Index).Tag = textFM2(Index).Text
       End If
   End If
   
   If Cancel = True Then
      textFM2(Index).SetFocus
      textFM2_GotFocus Index
   End If
End Sub

'Added by Lydia 2021/02/05
Private Sub textFM2_Start_GotFocus()
    TextInverse textFM2_Start
End Sub

Private Sub textFM2_Start_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(textFM2_Start.Text)
   If Cancel = True Then TextInverse textFM2_Start
End Sub

