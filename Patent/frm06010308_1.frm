VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010308_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-延期"
   ClientHeight    =   6200
   ClientLeft      =   80
   ClientTop       =   990
   ClientWidth     =   9150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6200
   ScaleWidth      =   9150
   Begin VB.Frame FraPA174 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   525
      Left            =   8310
      TabIndex        =   54
      Top             =   750
      Visible         =   0   'False
      Width           =   825
      Begin VB.CommandButton CmdPA174 
         BackColor       =   &H00C0FFFF&
         Caption         =   "特殊字"
         Height          =   280
         Left            =   0
         Style           =   1  '圖片外觀
         TabIndex        =   55
         Top             =   210
         Width           =   800
      End
      Begin VB.Label lblPA174 
         Caption         =   "有特殊字"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   35
         TabIndex        =   56
         Top             =   0
         Width           =   765
      End
   End
   Begin VB.TextBox txtCP07 
      Height          =   270
      Left            =   6450
      MaxLength       =   7
      TabIndex        =   53
      Top             =   3870
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtCP06 
      Height          =   270
      Left            =   5130
      MaxLength       =   7
      TabIndex        =   52
      Top             =   3870
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text11 
      Height          =   270
      Left            =   3450
      MaxLength       =   15
      TabIndex        =   51
      Text            =   "一(二)"
      Top             =   3090
      Width           =   1485
   End
   Begin VB.TextBox Text12 
      Height          =   270
      Left            =   5430
      MaxLength       =   11
      TabIndex        =   50
      Top             =   3090
      Width           =   1260
   End
   Begin VB.TextBox txtCP84 
      Height          =   270
      Left            =   1305
      MaxLength       =   7
      TabIndex        =   6
      Top             =   3660
      Width           =   1140
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   2220
      MaxLength       =   7
      TabIndex        =   5
      Top             =   2805
      Width           =   1005
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   225
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   4860
      MaxLength       =   1
      TabIndex        =   4
      Top             =   2520
      Width           =   300
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8220
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6240
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7080
      TabIndex        =   9
      Top             =   70
      Width           =   1110
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1305
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2205
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm06010308_1.frx":0000
      Left            =   1170
      List            =   "frm06010308_1.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   19
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   960
      MaxLength       =   3
      TabIndex        =   18
      Top             =   180
      Width           =   550
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1515
      MaxLength       =   6
      TabIndex        =   17
      Top             =   180
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   2355
      MaxLength       =   1
      TabIndex        =   16
      Top             =   180
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   2595
      MaxLength       =   2
      TabIndex        =   15
      Top             =   180
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   4860
      MaxLength       =   1
      TabIndex        =   1
      Top             =   2205
      Width           =   300
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   1305
      MaxLength       =   4
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1935
      Left            =   30
      TabIndex        =   7
      Top             =   4185
      Width           =   9045
      _ExtentX        =   15946
      _ExtentY        =   3404
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   7590
      TabIndex        =   2
      Top             =   2190
      Width           =   1500
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;556"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "繳費金額:"
      Height          =   180
      Left            =   480
      TabIndex        =   49
      Top             =   3690
      Width           =   765
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "辦理依據:"
      Height          =   180
      Left            =   480
      TabIndex        =   48
      Top             =   2850
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "發文日期："
      Height          =   180
      Left            =   1275
      TabIndex        =   47
      Top             =   2850
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "發文字號：（          ）智專                                     字第                               號"
      Height          =   180
      Left            =   1275
      TabIndex        =   46
      Top             =   3150
      Width           =   5670
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "範例：（  102  ）智專    一(二)15172     字第    10241450220    號"
      Height          =   180
      Left            =   1635
      TabIndex        =   45
      Top             =   3390
      Width           =   4935
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   6690
      TabIndex        =   43
      Top             =   2205
      Width           =   900
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "再審第          次"
      Height          =   180
      Index           =   2
      Left            =   4230
      TabIndex        =   42
      Top             =   2580
      Width           =   1170
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   30
      X2              =   8970
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   30
      X2              =   8970
      Y1              =   2130
      Y2              =   2130
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   10
      Left            =   4800
      TabIndex        =   41
      Top             =   1770
      Width           =   1410
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2487;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   9
      Left            =   1170
      TabIndex        =   40
      Top             =   1770
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3387;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   8
      Left            =   2100
      TabIndex        =   39
      Top             =   2550
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期　:"
      Height          =   180
      Left            =   120
      TabIndex        =   38
      Top             =   2205
      Width           =   1125
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   3960
      TabIndex        =   37
      Top             =   180
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   3960
      TabIndex        =   36
      Top             =   1470
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   120
      TabIndex        =   35
      Top             =   1470
      Width           =   945
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   0
      Left            =   4800
      TabIndex        =   34
      Top             =   180
      Width           =   1410
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2487;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3960
      TabIndex        =   33
      Top             =   1170
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   120
      TabIndex        =   32
      Top             =   1170
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   31
      Top             =   180
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   30
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   3960
      TabIndex        =   29
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Width           =   765
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   1
      Left            =   1170
      TabIndex        =   27
      Top             =   510
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3387;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   2
      Left            =   4800
      TabIndex        =   26
      Top             =   510
      Width           =   1410
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2487;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   3
      Left            =   1860
      TabIndex        =   25
      Top             =   840
      Width           =   6420
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   4
      Left            =   1170
      TabIndex        =   24
      Top             =   1170
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3387;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   5
      Left            =   4800
      TabIndex        =   23
      Top             =   1170
      Width           =   1410
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2487;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   6
      Left            =   1170
      TabIndex        =   22
      Top             =   1470
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3387;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   7
      Left            =   4800
      TabIndex        =   21
      Top             =   1470
      Width           =   3420
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6032;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "是否修改申請書內容          (Y:WORD)"
      Height          =   180
      Index           =   1
      Left            =   3180
      TabIndex        =   20
      Top             =   2205
      Width           =   2880
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "延期案件性質:"
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   2550
      Width           =   1125
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "未收文期限："
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   0
      Left            =   3960
      TabIndex        =   12
      Top             =   1770
      Width           =   765
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   1770
      Width           =   765
   End
End
Attribute VB_Name = "frm06010308_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/5 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strReceiveNo As String
'Modify by Morgan 2005/8/8 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String, m_CP110 As String, m_AgentName As String
Dim intWhere As Integer, intLastRow As Integer
'Add By Cheng 2003/06/24
Dim m_strNPReceiveNo As String '點選未收的期限的收文號
Public m_CP118isY As Boolean 'Add By Sindy 2018/7/4 是否為電子送件申請書:True.是
Dim m_CaseNo As String 'Add By Sindy 2018/7/4
Dim cp() As String 'Add By Sindy 2018/7/4


Private Sub cmdok_Click(Index As Integer)
Dim bolChk As Boolean, strTmp As String
Dim strFolder As String, strFileName As String 'Add By Sindy 2018/7/4
   
   Select Case Index
      Case 0 '確定
         'Add By Cheng 2003/06/24
         If Me.Text9.Text = "" Then
             MsgBox "請輸入延期案件性質或點選未收文期限資料!!!", vbExclamation + vbOKOnly
             Me.Text9.SetFocus
             Text9_GotFocus
             Exit Sub
         End If
         '若延期的案件性質為再審(107)
         If Me.Text9.Text = "107" Then
            If Me.Text6.Text = "" Then
                MsgBox "請輸入再審的次數!!!", vbExclamation + vbOKOnly
                Me.Text6.SetFocus
                Text6_GotFocus
                Exit Sub
            End If
         End If
         'Add by Morgan 2005/8/8
         If TxtValidate = False Then Exit Sub
         'Added by Lydia 2020/02/17 產生各式申請書時，若基本檔「名稱有特殊字」已勾選，彈訊息提醒，並一併開啟原始檔。
         If (pa(1) = "FCP" Or pa(1) = "P") And pa(174) = "Y" Then
             MsgBox MsgText(1111), vbInformation
             If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = False Then
                  Exit Sub
             End If
         End If
         'end 2020/02/17
         
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         'Add By Sindy 2018/7/4 電子送件申請書
         If m_CP118isY = True Then
            m_CaseNo = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4))
            'If Pub_StrUserSt03 = "M51" Then
            If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Or Pub_StrUserSt03 = "M51" Then
               strFolder = PUB_Getdesktop
            Else
               strFolder = FCP電子送件檔案存放路徑
            End If
            strFolder = strFolder & "\" & m_CaseNo
            If Dir(strFolder, vbDirectory) = "" Then
               MkDir strFolder
            End If
            
            '1.基本資料
            StartLetterPA_EData "01", "14", strReceiveNo, pa, cp, False
            NowPrint strReceiveNo, "01", "14", False, strUserNum, , , True, strExc(9)
            strFileName = strFolder & "\" & m_CaseNo & ".contact"
            Call PUB_MakeDoc(strExc(9), strFileName)
            '2.申請書
            If Me.Text9.Text = "107" And Val(Text6) = 1 Then '若延期的案件性質為再審
               If StartLetter2("01", "16") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "16", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "專利再審查申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            Else
               If StartLetter2("01", "24") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "24", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "專利申請延展指定期間申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            End If
         Else
         '2018/7/4 END
            If Text7 = "Y" Then
               bolChk = True
            Else
               bolChk = False
            End If
            Select Case Text9.Text
               Case "107"  '再審
                   If Me.Text6.Text = "1" Then
                       strTmp = "01"
                       'Added by Morgan 2015/7/27
                       '104/8/3起延期發文不再繳規費以免客戶不辦又要退費 --靜芳
                       If strSrvDate(1) >= "20150803" Then
                        strTmp = "09"
                       End If
                   Else
                       strTmp = "11"
                       'Added by Morgan 2015/7/27
                       '104/8/3起延期發文不再繳規費以免客戶不辦又要退費 --靜芳
                       If strSrvDate(1) >= "20150803" Then
                        strTmp = "10"
                       End If
                   End If
               Case 訴願
                  strTmp = "02"
               Case 改請新型
                  strTmp = "03"
                  
               '92.7.6 CANCEL BY SONIA
               'Case 異議_專
               '   strTmp = "04"
               Case 異議答辯
                  strTmp = "05"
               '92.7.6 CANCEL BY SONIA
               'Case 舉發
               '   strTmp = "06"
               Case 舉發答辯
                  strTmp = "07"
                  
               'Modify by Morgan 2004/9/23
               '加檢視中說, 製作中說
               'Modified by Morgan 2013/11/6 +235核對中說格式
               Case 補文件, 翻譯, 檢視中說, 製作中說, "235"
                  strTmp = "08"
               'Add By Cheng 2003/11/03
               '修正, 申復, 補充
               Case "204", "205", "206"
                  strTmp = "12"
               Case Else
            End Select
            strLetterDate = Text5.Text
            If strTmp = "" Then
               MsgBox "該性質並無申請書！"
            Else
               StartLetter "01", Text1 & Text2 & Text3 & Text4 & "&404", strTmp
               NowPrint Text1 & Text2 & Text3 & Text4 & "&404", "01", strTmp, bolChk, strUserNum
            End If
         End If
         frm060103_1.Show
         ' 90.08.27 modify by louis (回到原畫面要清除畫面)
         frm060103_1.ClearForm
      Case 1 '回前畫面
         frm060103_1.Show
      Case 2 '結束
         Unload frm060103_1
   End Select
   Unload Me
End Sub

'Add By Cheng 2003/06/24
Private Sub StartLetter(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)
Dim strTxt(1 To 10) As String, strTmp As String
Dim ii As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim iDays As Integer
Dim strFee As String 'Added by Morgan 2013/1/15
Dim iMonths As Integer, stLawDate As String 'Added by Morgan 2014/6/30
         
    EndLetter ET01, ET02, ET03, strUserNum
    ii = 0
    Select Case Me.Text9.Text
    Case "107" '再審申請
        '第一次再審申請
        If Me.Text6.Text = "1" Then
            strFee = GetPatentOfficialFee(pa(1), "107", "", pa(8), pa(9), pa(16))
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                "','申請書法條','" & IIf(pa(8) = "1", "第46條", IIf(pa(8) = "2", "第105條準用第40條", "第129條準用第46條")) & "')"
            '點選非延期案進入
            If Me.Text9.Enabled = False Then
                StrSQLa = "Select * From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09='" & strReceiveNo & "')) "
            '點選延期案進入
            Else
                '若有點選未收文期限
                If m_strNPReceiveNo <> "" Then
                    StrSQLa = "Select * From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09='" & m_strNPReceiveNo & "') "
                '若未點選未收文期限
                Else
                    StrSQLa = "Select * From Caseprogress Where CP09='" & GetNewCP09(strReceiveNo, "1002", "-1") & "' "
                End If
            End If
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                     "','准駁日','" & rsA("CP25").Value & "')"
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '點選非延期案進入
            If Me.Text9.Enabled = False Then
                StrSQLa = "Select * From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09='" & strReceiveNo & "') "
                rsA.CursorLocation = adUseClient
                rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                If rsA.RecordCount > 0 Then
                    ii = ii + 1
                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                        "','機關文號','" & rsA("CP08").Value & "')"
                    ii = ii + 1
                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                        "','來函收文日','" & rsA("CP05").Value & "')"
                    ii = ii + 1
                    If txtCP84 <> "" Then
                       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                           "','規費','" & txtCP84 & "')"
                    Else
                      '93.7.19 modify by sonia 調整規費
                       'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       '    "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                       '    "','規費','" & IIf(pa(8) = "1", "6000", IIf(pa(8) = "2", "4500", "3500")) & "')"
                       'Modified by Morgan 2013/1/15
                       ''strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                           "','規費','" & IIf(pa(8) = "1", "8000", IIf(pa(8) = "2", "0", "3500")) & "')"
                       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                           "','規費','" & strFee & "')"
                       '93.7.19 end
                    End If
                End If
                If rsA.State <> adStateClosed Then rsA.Close
                Set rsA = Nothing
            '點選延期案進入
            Else
                '若有點選未收文期限
                If m_strNPReceiveNo <> "" Then
                    StrSQLa = "Select * From Caseprogress Where CP09='" & m_strNPReceiveNo & "' "
                    rsA.CursorLocation = adUseClient
                    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                    If rsA.RecordCount > 0 Then
                        ii = ii + 1
                        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                            "','機關文號','" & rsA("CP08").Value & "')"
                        ii = ii + 1
                        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                            "','來函收文日','" & rsA("CP05").Value & "')"
                        ii = ii + 1
                        If txtCP84 <> "" Then
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                               "','規費','" & txtCP84 & "')"
                        Else
                           '93.7.19 modify by sonia 調整規費
                           'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           '    "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                           '    "','規費','" & IIf(pa(8) = "1", "6000", IIf(pa(8) = "2", "4500", "3500")) & "')"
                           'Modified by Morgan 2013/1/15
                           'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                               "','規費','" & IIf(pa(8) = "1", "8000", IIf(pa(8) = "2", "0", "3500")) & "')"
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                               "','規費','" & strFee & "')"
                           '93.7.19 end
                        End If
                    End If
                    If rsA.State <> adStateClosed Then rsA.Close
                    Set rsA = Nothing
                '若未點選未收文期限
                Else
                    StrSQLa = "Select * From Caseprogress Where CP09='" & GetNewCP09(strReceiveNo, "404", "-1") & "' "
                    rsA.CursorLocation = adUseClient
                    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                    If rsA.RecordCount > 0 Then
                        ii = ii + 1
                        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                            "','機關文號','" & rsA("CP08").Value & "')"
                        ii = ii + 1
                        'Modify by Morgan 2005/6/27 改抓延期發文規費
'                        If "" & rsA("CP17").Value <> "" Then
'                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'                               "','規費','" & rsA("CP17").Value & "')"
                        If txtCP84 <> "" Then
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                               "','規費','" & txtCP84 & "')"
                        '2005/6/27 end
                        Else
                           '93.7.19 modify by sonia 調整規費
                           'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                               "','規費','" & IIf(pa(8) = "1", "6000", IIf(pa(8) = "2", "4500", "3500")) & "')"
                           'Modified by Morgan 2013/1/15
                           'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                               "','規費','" & IIf(pa(8) = "1", "8000", IIf(pa(8) = "2", "0", "3500")) & "')"
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                               "','規費','" & strFee & "')"
                           '93.7.19 end
                        End If
                    End If
                    If rsA.State <> adStateClosed Then rsA.Close
                    Set rsA = Nothing
                    StrSQLa = "Select * From Caseprogress Where CP09='" & GetNewCP09(strReceiveNo, "1002", "0") & "' "
                    rsA.CursorLocation = adUseClient
                    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                    If rsA.RecordCount > 0 Then
                        ii = ii + 1
                        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                            "','來函收文日','" & rsA("CP05").Value & "')"
                    End If
                    If rsA.State <> adStateClosed Then rsA.Close
                    Set rsA = Nothing
                End If
            End If
        '第二次再審申請
        Else
            'Modify By Cheng 2003/07/02
'            '一定是點延期進入的
''            '點選非延期案進入
''            If Me.Text9.Enabled = False Then
''                strSQLA = "Select * From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09='" & strReceiveNo & "') "
''            '點選延期案進入
''            Else
'                '若有點選未收文期限
'                If m_strNPReceiveNo <> "" Then
'                    strSQLA = "Select * From Caseprogress Where CP09='" & m_strNPReceiveNo & "' "
'                '若未點選未收文期限
'                Else
'                    strSQLA = "Select * From Caseprogress Where CP43='" & strReceiveNo & "' "
'                End If
''            End If
            'StrSQLa = "Select * From Caseprogress Where CP43='" & strReceiveNo & "' "
            'Modify By Sindy 2018/2/14 抓函號原已經Mark起來,敏莉提抓1004延期受理的函號
            StrSQLa = "Select cp08,cp05,ed08 From Caseprogress,edocument Where CP10='1004' and cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                      " AND ed11(+)=cp09" & _
                      " ORDER BY CP05 DESC"
            '2018/2/14 END
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               If "" & rsA.Fields("CP08") <> "" Then
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                      "','機關文號','" & rsA("CP08").Value & "')"
               Else
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                      "','機關文號','（" & Left(strSrvDate(1), 4) - 1911 & "）智專（）字第號')"
               End If
               If "" & rsA.Fields("ed08") <> "" Then
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                      "','來函收文日','" & Mid("" & rsA("ed08").Value, 1, 4) - 1911 & "年" & Mid("" & rsA("ed08").Value, 5, 2) & "月" & Mid("" & rsA("ed08").Value, 7, 2) & "日" & "')"
               Else
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                      "','來函收文日','年月日')"
               End If
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
        End If
    'Add By Cheng 2003/11/04
    'Modify By Sindy 2019/8/8 + 239.擇一申復 同申復
    Case "204", "205", "206", "239" '修正, 申復, 補充, 擇一申復
        '點選非延期案進入
        If Me.Text9.Enabled = False Then
            StrSQLa = "Select * From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09='" & strReceiveNo & "') "
        '點選延期案進入
        Else
            '若有點選未收文期限
            If m_strNPReceiveNo <> "" Then
                StrSQLa = "Select * From Caseprogress Where CP09='" & m_strNPReceiveNo & "' "
            '若未點選未收文期限
            Else
                'modify by sonia 2014/6/26 FCP-043107
                'StrSQLa = "Select * From Caseprogress Where CP43='" & strReceiveNo & "' "
                StrSQLa = "Select * From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09='" & strReceiveNo & "')) "
            End If
        End If
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                "','機關文號','" & rsA("CP08").Value & "')"
            'ADD BY SONIA 2014/6/25 FCP-041158
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                "','來函收文日','" & rsA("CP05").Value & "')"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                "','法定期限','" & "" & rsA("CP07").Value & "')"
            'END 2014/6/25
            stLawDate = "" & rsA("CP07").Value 'Added by Morgan 2014/6/30
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        
        
        'Modify by Morgan 2008/5/30 '改抓設定檔
        ''Add by Morgan 2008/1/11
        ''申復要判斷沒有外國申請人時只可延60天
        'If Text9 = "205" Then
        '    If PUB_ExistForeigner(pa(1) & pa(2) & pa(3) & pa(4)) = False Then
        '       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        '          "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
        '          "','延期天數','60')"
        '    End If
        'End If
        
        'Modified by Morgan 2014/6/30 +延期後法定期限 Ex.FCP-40258 (原申請書固定為法限+3個月)
        'If ClsLawGetCaseFeeDelay(pa(1), pa(9), Text9.Text, strExc, pa(1) & pa(2) & pa(3) & pa(4), iDays) Then
        '    ii = ii + 1
        '    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        '       "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
        '       "','延期天數','" & iDays & "')"
        'End If
        ''end 2008/5/30
        '
        ''Add by Morgan 2008/9/5
        'If ClsLawGetCaseFeeDelay(pa(1), pa(9), Text9.Text, strExc, pa(1) & pa(2) & pa(3) & pa(4), , iDays) Then
        '    ii = ii + 1
        '    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        '       "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
        '       "','延期月數','" & iDays & "')"
        'End If
        If ClsLawGetCaseFeeDelay(pa(1), pa(9), Text9.Text, strExc, pa(1) & pa(2) & pa(3) & pa(4), iDays, iMonths) Then
            strExc(1) = ""
            If iMonths > 0 Then
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','延期月數','" & iMonths & "')"
               strExc(1) = CompDate(1, iMonths, stLawDate)
            ElseIf iDays > 0 Then
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','延期天數','" & iDays & "')"
               strExc(1) = CompDate(2, iDays, stLawDate)
            End If
            If strExc(1) <> "" Then
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','延期後法定期限','" & strExc(1) & "')"
            End If
        End If
        'end 2014/6/30
        
        'end 2008/5/30
        
        
    Case "501", "302" '訴願，改請新型
        '點選非延期案進入
        If Me.Text9.Enabled = False Then
            StrSQLa = "Select * From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09='" & strReceiveNo & "') "
        '點選延期案進入
        Else
            '若有點選未收文期限
            If m_strNPReceiveNo <> "" Then
                StrSQLa = "Select * From Caseprogress Where CP09='" & m_strNPReceiveNo & "' "
            '若未點選未收文期限
            Else
                StrSQLa = "Select * From Caseprogress Where CP43='" & strReceiveNo & "' "
            End If
        End If
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                "','機關文號','" & rsA("CP08").Value & "')"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                "','來函收文日','" & rsA("CP05").Value & "')"
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        
         'Add by Morgan 2008/3/3
         strExc(0) = "select cp17 from caseprogress where cp09='" & strReceiveNo & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','規費','" & RsTemp(0).Value & "')"
         End If
         'end 2008/3/3
         
    '92.7.6 CANCEL BY SONIA
    'Case "801" '異議
    '    ii = ii + 1
    '    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
    '        "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
    '        "','申請書法條','" & IIf(pa(8) = "1", "第41條", IIf(pa(8) = "2", "第102條", "第115條")) & "')"
    '    ii = ii + 1
    '    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
    '        "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
    '        "','規費','" & IIf(pa(8) = "1", "6000", IIf(pa(8) = "2", "4500", "3500")) & "')"
    Case "802", "804" '異議答辯, 舉發答辯
        '點選非延期案進入
        If Me.Text9.Enabled = False Then
            StrSQLa = "Select * From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09='" & strReceiveNo & "')) "
        '點選延期案進入
        Else
            StrSQLa = "Select * From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09='" & m_strNPReceiveNo & "') "
        End If
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                 "','准駁日','" & rsA("CP25").Value & "')"
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        '點選非延期案進入
        If Me.Text9.Enabled = False Then
            StrSQLa = "Select * From Caseprogress Where CP09=(Select CP43 From Caseprogress Where CP09='" & strReceiveNo & "') "
        '點選延期案進入
        Else
            '若有點選未收文期限
            If m_strNPReceiveNo <> "" Then
                StrSQLa = "Select * From Caseprogress Where CP09='" & m_strNPReceiveNo & "' "
            '若未點選未收文期限
            Else
                StrSQLa = "Select * From Caseprogress Where CP43='" & strReceiveNo & "' "
            End If
        End If
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                "','機關文號','" & rsA("CP08").Value & "')"
            '2009/5/19 add by sonia
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                "','列印備註','" & rsA("CP36").Value & "')"
            '2009/5/19 end
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    '92.7.6 CANCEL BY SONIA
    'Case "803" '舉發
    '    ii = ii + 1
    '    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
    '        "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
    '        "','申請書法條','" & IIf(pa(8) = "1", "第71條", IIf(pa(8) = "2", "第104條", "第121條")) & "')"
    '    ii = ii + 1
    '    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
    '        "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
    '        "','規費','" & IIf(pa(8) = "1", "9000", IIf(pa(8) = "2", "8500", "8000")) & "')"
   End Select
   
'Remove by Morgan 2009/2/2 因為cp110有存檔所以改抓共用的例外欄位就好
'   'Add by Morgan 2005/8/10
'   ii = ii + 1
'   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'       "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'       "','出名代理人','" & m_AgentName & "')"

    If ii <> 0 Then
        'edit by nickc 2007/02/05 不用 dll 了
        'If Not objLawDll.ExecSQL(ii, strTxt) Then
        If Not ClsLawExecSQL(ii, strTxt) Then
            MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
        End If
    End If
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label12(3) = pa(5)
      Case "英"
         Label12(3) = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         Label12(3) = pa(7)
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm060103_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
   'Add by Morgan 2005/8/8
   ReDim pa(TF_PA)
   ReDim cp(TF_CP) 'Add By Sindy 2018/7/4
   ReadPatent
   'Add by Morgan 2005/8/8
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   PUB_SetOurAgent lstNameAgent, pa(), m_CP110, , True
   'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1100
   lstNameAgent.Width = 1300

   Combo1.ListIndex = 0
   Text5.Text = strSrvDate(2)
   '92.7.16 ADD BY SONIA
   'Modify by Morgan 2004/8/6
   '加翻譯
   'Modify by Morgan 2004/8/6
   '加檢視中說,製作中說
   'Modified by Morgan 2013/11/6 +235核對中說格式
   If Text9 = 補文件 Or Text9 = 翻譯 Or Text9 = 檢視中說 Or Text9 = "235" Or Text9 = 製作中說 Then
      Text7 = "Y"
   End If
   '92.7.16 END
   
   'Add By Sindy 2018/7/31
   If m_CP118isY = True Then Text7.Enabled = False
   If Pub_StrUserSt03 <> "M51" Then
      txtCP06.Visible = False
      txtCP07.Visible = False
   End If
   
   FraPA174.BackColor = &H8000000F 'Added by Lydia 2020/02/21
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010308_1 = Nothing
End Sub

Private Sub ReadPatent()
Dim rsTemp1 As New ADODB.Recordset, Lbl As Object
Dim Cancel As Boolean    'ADD BY SONIA 2014/6/26
Dim strCP43CP09 As String '讀取辦理依據的總收文號
   
   For Each Lbl In Label12
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   If ClsPDReadPatentDatabase(pa(), intWhere) Then
      Text5 = pa(10)
      Label12(1) = pa(11)
      Label12(2) = pa(22)
      Label12(3) = pa(5)
   End If
   Call Text9_Change 'Add By Sindy 2019/1/3
   
   'Add By Sindy 2018/7/4
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
   End If
   '2018/7/4 END
   
   'Modify by Morgan 2004/9/9 若無收文規費時抓發文規費
   'Modify by Morgan 2005/6/23 改都抓發文規費
   'strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2,cp43,cp10,CP06,CP07,nvl(CP17,CP84) from caseprogress,casepropertymap,staff," & _
      "staff staff1 where cp09='" & strReceiveNo & "' AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and " & _
      "cp13=staff1.st01(+)"
   strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2,cp43,cp10,CP06,CP07,CP84,CP110" & _
               " from caseprogress,casepropertymap,staff,staff staff1" & _
               " where cp09='" & strReceiveNo & "' AND cp01=cpm01(+) and cp10=cpm02(+)" & _
               " and cp14=staff.st01(+) and cp13=staff1.st01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      m_CP110 = "" & .Fields("CP110")
      If Not IsNull(.Fields(0)) Then
         Label12(0) = .Fields(0)
         If Label12(0).Caption <> "延期" Then Text9.Enabled = False '點選非延期案進入
         Text9 = .Fields(4) '延期案件性質
         Label12(8).Caption = .Fields(0)
         If Me.Text9.Text = "404" Then '404.延期
            Me.Text9.Text = ""
            Label12(8).Caption = ""
         End If
      End If
      If Not IsNull(.Fields(1)) Then Label12(4) = .Fields(1)
      If Not IsNull(.Fields(2)) Then Label12(5) = .Fields(2)
      If Not IsNull(.Fields(5)) Then Label12(9) = .Fields(5) - 19110000 '本所期限
      If Not IsNull(.Fields(6)) Then Label12(10) = .Fields(6) - 19110000 '法定期限
      If Not IsNull(.Fields(7)) Then txtCP84 = .Fields(7) '發文規費
      txtCP06 = Label12(9) '延期案件性質的本所期限
      txtCP07 = Label12(10) '延期案件性質的法定期限
      strCP43CP09 = strReceiveNo '讀取辦理依據的總收文號
      If Not IsNull(.Fields(3)) Then '有相關總收文號
         'MODIFY BY SONIA 2014/6/26 +CP10,FCP-043107 先申復按延期按鈕發文,再至各式申請書點選延期
         strExc(0) = "SELECT CP05,CP08,CP10,CP06,CP07 FROM CASEPROGRESS WHERE CP09='" & .Fields(3) & "'"
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(rsTemp1.Fields(0)) Then Label12(6) = TransDate(rsTemp1.Fields(0), 1)
            If Not IsNull(rsTemp1.Fields(1)) Then Label12(7) = rsTemp1.Fields(1)
            'ADD BY SONIA 2014/6/26 +CP10,FCP-043107
            If Me.Text9.Text = "" And .Fields(3) < "C" Then
               If Not IsNull(rsTemp1.Fields(2)) Then Me.Text9.Text = rsTemp1.Fields(2)
               Text9_Validate Cancel
               'Add By Sindy 2018/7/31
               If Val("" & rsTemp1.Fields("CP06")) > 0 Then
                  txtCP06 = rsTemp1.Fields("CP06") - 19110000 '延期案件性質的本所期限
               End If
               If Val("" & rsTemp1.Fields("CP07")) > 0 Then
                  txtCP07 = rsTemp1.Fields("CP07") - 19110000 '延期案件性質的法定期限
               End If
               strCP43CP09 = .Fields(3) '讀取辦理依據的總收文號
               '2018/7/31 END
               'Add By Sindy 2018/11/8 ex:FCP-51397
               If Not IsNull(rsTemp1("cp08")) Then
                  strExc(0) = Mid(rsTemp1("cp08"), InStr(rsTemp1("cp08"), "智專") + 2, Len(rsTemp1("cp08")))
                  Text11 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
                  strExc(0) = Replace(strExc(0), Text11 & "字第", "")
                  Text12 = Mid(strExc(0), 1, InStr(strExc(0), "號") - 1)
               End If
               '2018/11/8 END
            End If
            'END 2014/6/26
         End If
      End If
'      'Add by Morgan 2005/6/23 若有延期過則抓延期發文規費(延期先發文,後出申請書)
'      If Label12(0).Caption <> "延期" Then
'         strExc(0) = "select CP84,cp09 from caseprogress" & _
'            " where cp43='" & strReceiveNo & "' AND cp10='404' and cp27 is not null and cp57 is null"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            txtCP84 = RsTemp.Fields("CP84")
'         End If
'      End If
'      '2005/6/23 end
      'Add By Sindy 2018/7/30
      If Me.Text9.Text = "" And Label12(0).Caption = "延期" Then
         strExc(0) = "select np01,np07,np08,np09 from caseprogress,nextprogress" & _
            " where cp09='" & strReceiveNo & "' AND cp30 is not null" & _
            " and cp01=np02 and cp02=np03 and cp03=np04 and cp04=np05 and cp30=np22"
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(rsTemp1.Fields("np07")) Then Me.Text9.Text = rsTemp1.Fields("np07")
            Text9_Validate Cancel
            If Val("" & rsTemp1.Fields("np08")) > 0 Then
               txtCP06 = rsTemp1.Fields("np08") - 19110000 '延期案件性質的本所期限
            End If
            If Val("" & rsTemp1.Fields("np09")) > 0 Then
               txtCP07 = rsTemp1.Fields("np09") - 19110000 '延期案件性質的法定期限
            End If
            strCP43CP09 = rsTemp1.Fields("np01") '讀取辦理依據的總收文號
            strExc(0) = "SELECT CP05,CP08 FROM CASEPROGRESS WHERE CP09='" & rsTemp1.Fields("np01") & "'"
            intI = 1
            Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If Not IsNull(rsTemp1.Fields(0)) Then Label12(6) = TransDate(rsTemp1.Fields("CP05"), 1)
               If Not IsNull(rsTemp1.Fields(1)) Then Label12(7) = rsTemp1.Fields("CP08")
               'Add By Sindy 2018/11/8 ex:FCP-51397
               If Not IsNull(rsTemp1("cp08")) Then
                  strExc(0) = Mid(rsTemp1("cp08"), InStr(rsTemp1("cp08"), "智專") + 2, Len(rsTemp1("cp08")))
                  Text11 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
                  strExc(0) = Replace(strExc(0), Text11 & "字第", "")
                  Text12 = Mid(strExc(0), 1, InStr(strExc(0), "號") - 1)
               End If
               '2018/11/8 END
            End If
         End If
      End If
      '2018/7/30 END
   End If
   End With
    'Modify By Cheng 2003/06/24
    '抓本所案號相同且是否續辦為NULL的下一程序資料
'   strExc(0) = "SELECT '',CPM03," & SQLDate("NP08") & "," & SQLDate("NP09") & _
'      ",NP13,NP14," & SQLDate("NP11") & " FROM NEXTPROGRESS,CASEPROPERTYMAP" & _
'      " WHERE NP01='" & strReceiveNo & "' AND " & _
'      "(NP06<>'Y' OR NP06 IS NULL) AND NP02=CPM01(+) AND NP07=CPM02(+)"
   strExc(0) = "SELECT '', CPM03, " & SQLDate("NP08") & ", " & SQLDate("NP09") & _
      ", NP13, NP14, " & SQLDate("NP11") & ", NP01, NP07 FROM NEXTPROGRESS,CASEPROPERTYMAP" & _
      " WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND " & _
      " NP06 IS NULL AND NP02=CPM01(+) AND NP07=CPM02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   
   'Add By Sindy 2018/7/4
   '來函文號:
   'Modify By Sindy 2018/10/22 新案翻譯(中說)抓新申請案 ex:FCP-058914
   '201.新案翻譯, 209.檢視中說, 235.核對中說格式, 210.製作中說
   If cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "210" Or _
      Text9 = "201" Or Text9 = "209" Or Text9 = "235" Or Text9 = "210" Then
      strExc(0) = "select cp05,cp08,ed08 from caseprogress,edocument where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND ed11(+)=cp09 AND cp09 in(select cp09 from caseprogress where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' and cp10 in(" & NewCasePtyList & ")) ORDER BY CP05 DESC"
   Else
   '2018/10/22 END
      'CP05 Desc -> NVL(ED08,CP05) Desc 歸卷公文會有一筆以上，考慮可能有紙本公文保留CP05的判斷
      strExc(0) = "select cp08,ed08 from caseprogress,edocument where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP09='" & strCP43CP09 & "' AND ed11(+)=cp09 ORDER BY NVL(ED08,CP05) Desc"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp("ED08")) Then
         Text10 = RsTemp("ED08") - 19110000
         Text8 = Val(Text10) \ 10000
         If Not IsNull(RsTemp("cp08")) Then
            strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
            Text11 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
            strExc(0) = Replace(strExc(0), Text11 & "字第", "")
            Text12 = Mid(strExc(0), 1, InStr(strExc(0), "號") - 1)
         End If
      End If
   End If
   If Text10 = "" Then
      '先抓相關總收文號
      'strExc(0) = "select cp08,ed08 from caseprogress,edocument where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP09='" & cp(9) & "' AND ed11(+)=cp43 and cp43 is not null ORDER BY NVL(ED08,CP05) Desc"
      'Modify By Sindt 2018/10/29 ex:P-118713 延期
      strExc(0) = "select cp08,ed08 from caseprogress,edocument" & _
                  " where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                  " AND CP09=(select cp43 from caseprogress" & _
                  " where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                  " AND CP09='" & strCP43CP09 & "' and cp43 is not null)" & _
                  " and ed11(+)=cp09 and cp09 is not null" & _
                  " ORDER BY NVL(ED08,CP05) Desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Not IsNull(RsTemp("ED08")) Then
            Text10 = RsTemp("ED08") - 19110000
            Text8 = Val(Text10) \ 10000
            If Not IsNull(RsTemp("cp08")) Then
               strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
               Text11 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
               strExc(0) = Replace(strExc(0), Text11 & "字第", "")
               Text12 = Mid(strExc(0), 1, InStr(strExc(0), "號") - 1)
            End If
         End If
      End If
   End If
   '2018/7/4 END
   
   'Added by Lydia 2020/02/21 預設「名稱有特殊字」
   FraPA174.Visible = False
   If pa(1) = "FCP" Or pa(1) = "P" Then
       If pa(174) = "Y" Then
          FraPA174.Visible = True
       End If
   End If
   'end 2020/02/21
End Sub

Private Sub MSHFlexGrid1_Click()
Dim ii As Integer
    GridClick MSHFlexGrid1, intLastRow, 0
    'Add By Cheng 2003/06/24
    If Me.Text9.Enabled = True Then
        Me.Text9.Text = ""
        Me.Label12(8).Caption = ""
        m_strNPReceiveNo = ""
        For ii = 1 To Me.MSHFlexGrid1.Rows - 1
            If Me.MSHFlexGrid1.TextMatrix(ii, 0) <> "" Then
               Me.Text9.Text = Me.MSHFlexGrid1.TextMatrix(ii, 8)
               Me.Label12(8).Caption = Me.MSHFlexGrid1.TextMatrix(ii, 1)
               m_strNPReceiveNo = Me.MSHFlexGrid1.TextMatrix(ii, 7)
               'Add By Sindy 2018/7/31
               If Val(Me.MSHFlexGrid1.TextMatrix(ii, 2)) > 0 Then
                  txtCP06 = Replace(Me.MSHFlexGrid1.TextMatrix(ii, 2), "/", "") '延期案件性質的本所期限
               End If
               If Val(Me.MSHFlexGrid1.TextMatrix(ii, 3)) > 0 Then
                  txtCP07 = Replace(Me.MSHFlexGrid1.TextMatrix(ii, 3), "/", "") '延期案件性質的法定期限
               End If
               '2018/7/31 END
               Exit For
            End If
        Next ii
    End If
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
   CloseIme
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then
      If Not ChkDate(Text10) Then
         Cancel = True
      ElseIf Val(Text10) > Val(strSrvDate(2)) Then
         MsgBox "發文日期不可大於系統日！"
         Cancel = True
      Else
         Text8 = Val(Text10) \ 10000
      End If
   End If
End Sub

Private Sub Text11_GotFocus()
   Dim iPos As Integer
   
   iPos = InStr(Text11, "一(二)")
   If iPos > 0 Then
      Text11.SelStart = iPos + 3
      Text11.SelLength = Len(Text11) - 4
   Else
      TextInverse Text11
   End If
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
   CloseIme
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(Text5.Text)
   If Cancel = True Then TextInverse Text5
End Sub

Private Sub Text6_GotFocus()
    TextInverse Me.Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2003/06/24
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 Then
        KeyAscii = 0
    End If
End Sub

'Add By Sindy 2018/7/31
Private Sub Text6_Validate(Cancel As Boolean)
   If Val(Text6) > 1 Then '再審第2次延期
      txtCP84 = 0
      '1004.延期受理
      strExc(0) = "select cp09,cp06,cp07" & _
                  " from caseprogress" & _
                  " where cp10='1004' and cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                  " ORDER BY CP05 DESC"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Val("" & RsTemp.Fields("cp06")) > 0 Then
            txtCP06 = RsTemp.Fields("cp06") - 19110000
         End If
         If Val("" & RsTemp.Fields("cp07")) > 0 Then
            txtCP07 = "" & RsTemp.Fields("cp07") - 19110000
         End If
         strExc(0) = "select cp08,ed08 from caseprogress,edocument where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP09='" & RsTemp.Fields("cp09") & "' AND ed11(+)=cp09 ORDER BY NVL(ED08,CP05) Desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(RsTemp("ED08")) Then
               Text10 = RsTemp("ED08") - 19110000
               Text8 = Val(Text10) \ 10000
               If Not IsNull(RsTemp("cp08")) Then
                  strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
                  Text11 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
                  strExc(0) = Replace(strExc(0), Text11 & "字第", "")
                  Text12 = Mid(strExc(0), 1, InStr(strExc(0), "號") - 1)
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub GridHead()
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1000: .Text = "下一程序"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1200: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1200: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1500: .Text = "機關文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1400: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1200: .Text = "解除期限日期"
        'Add By Cheng 2003/06/24
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 0: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .ColWidth(8) = 0: .Text = "下一程序"
      '判斷是否有資料
      .Visible = True
   End With
End Sub

Private Sub Text9_Change()
    'Add By Cheng 2003/06/24
    '若延期的案件性質為再審
    If Me.Text9.Text = "107" Then
        Me.Text6.Visible = True
        Me.Text6.Enabled = True
        Me.Label18(2).Caption = "再審第          次"
        Me.Label18(2).Visible = True
    '若延期的案件性質非再審
    Else
      'Add By Sindy 2019/1/3
      If pa(8) = "2" Then '新型
         Me.Text6.Visible = True
         Me.Text6.Enabled = True
         Me.Label18(2).Caption = "延期第          次"
         Me.Label18(2).Visible = True
      Else
      '2019/1/3 END
         Me.Text6.Visible = False
         Me.Text6.Enabled = False
         Me.Label18(2).Visible = False
      End If
    End If
End Sub

Private Sub Text9_GotFocus()
  TextInverse Text9
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
Dim strTempName As String
   
    If Me.Text9.Text = "" Then Exit Sub
    'edit by nickc 2007/02/02 不用 dll 了
    'If objPublicData.GetCaseProperty("FCP", Text9, strTempName, False) Then
    If ClsPDGetCaseProperty(pa(1), Text9, strTempName, False) Then
        Label12(8) = strTempName
    Else
        Label12(8) = ""
        Cancel = True
    End If
    If Cancel = True Then TextInverse Text9
End Sub

'Add By Cheng 2004/01/20
Private Function GetNewCP09(ByVal strCP09 As String, strCP10 As String, strKind As String) As String
'strKind : -1 前一個, 0, 本身
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   GetNewCP09 = ""
   Do While GetNewCP09 = ""
      StrSQLa = "Select * From Caseprogress Where CP09='" & strCP09 & "' "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         strCP09 = "" & rsA("CP43").Value
         If strCP10 = "" & rsA("CP10").Value Then
            If strKind = "0" Then
               GetNewCP09 = "" & rsA("CP09").Value
            Else
               GetNewCP09 = "" & rsA("CP43").Value
            End If
         End If
      Else
         Exit Do
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   Loop
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

'Add by Morgan 2005/8/8
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
   TxtValidate = True
End Function

'Add by Morgan 2005/8/8
Private Function FormSave() As Boolean
Dim strCon As String
   
On Error GoTo ErrorHandler
   
   If Label12(0).Caption = "延期" Then 'Add By Sindy 2018/7/31 + if
      cnnConnection.BeginTrans
      
      cp(84) = Val(txtCP84)
      strCon = strCon & ",cp84=" & cp(84)
      If m_CP118isY = True Then
         cp(118) = "A"
      Else
         cp(118) = ""
      End If
      strCon = strCon & ",cp118=" & CNULL(cp(118))
   '   If lstNameAgent.Visible = True Then
   '      cp(110) = m_CP110
   '      strSql = " UPDATE CASEPROGRESS SET cp110=" & CNULL(cp(110)) & strCon & " WHERE CP09='" & strReceiveNo & "'"
   '      cnnConnection.Execute strSql
   '   End If
      cp(110) = m_CP110
      strSql = " UPDATE CASEPROGRESS SET cp110=" & CNULL(cp(110)) & strCon & " WHERE CP09='" & strReceiveNo & "' and cp158=0 and cp159=0"
      cnnConnection.Execute strSql
      
      cnnConnection.CommitTrans
   End If
   FormSave = True
   
ErrorHandler:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
   End If
End Function

'Add by Morgan 2005/8/8
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   m_CP110 = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modify By Sindy 2021/5/10
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         '2021/5/10 END

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

'申請書
Private Function StartLetter2(ByVal ET01 As String, ByVal ET03 As String) As Boolean
Dim strTxt(200) As String, strTmp As String
Dim ii As Integer, jj As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strCP07Add2M As String
Dim strPA10Add2M As String, strPA10Add4M As String, strPA10Add6M As String
Dim strCP27Add4M As String, strCP27Add6M As String
Dim intMonth As Integer
   
   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
   
   Call PUB_GetApplPA_EData(ET01, ET03, strReceiveNo, pa())

   '出名代理人
   strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      jj = 1
      Do While Not .EOF
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
         jj = jj + 1
         .MoveNext
      Loop
      End With
   End If
   
   '辦理依據
   If Text10 <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文日期','" & ChangeTStringToTDateString(Text10) & "')"
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','智專字','" & Text11 & "')"
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文號','" & Text12 & "')"
   End If
   '申請內容:
   '1.貴局原指定期間為：[　年　月　日前/文到次日起　內。｛<延期前期限>｝]
   '2.[因「尚在準備中，不克於期限內補呈」原因，「　　　」文件申請延期「至　年　月　日/2個月內」補正。｛<申請文件>｝]
   If pa(10) > 0 Then
      '申請日+2個月
      strPA10Add2M = CompDate(1, 2, DBDATE(pa(10)))
      strPA10Add2M = Left(strPA10Add2M, 4) - 1911 & "年" & Mid(strPA10Add2M, 5, 2) & "月" & Right(strPA10Add2M, 2) & "日"
      '申請日+4個月
      strPA10Add4M = CompDate(1, 4, DBDATE(pa(10)))
      strPA10Add4M = Left(strPA10Add4M, 4) - 1911 & "年" & Mid(strPA10Add4M, 5, 2) & "月" & Right(strPA10Add4M, 2) & "日"
      '申請日+6個月
      strPA10Add6M = CompDate(1, 6, DBDATE(pa(10)))
      strPA10Add6M = Left(strPA10Add6M, 4) - 1911 & "年" & Mid(strPA10Add6M, 5, 2) & "月" & Right(strPA10Add6M, 2) & "日"
   End If
   If Val(txtCP07) > 0 Then
      '法定期限
      strExc(1) = Left(txtCP07, Len(txtCP07) - 4) & "年" & Mid(txtCP07, Len(txtCP07) - 4 + 1, 2) & "月" & Right(txtCP07, 2) & "日"
      '法定期限+2個月
      strCP07Add2M = CompDate(1, 2, DBDATE(txtCP07))
      strCP07Add2M = Left(strCP07Add2M, 4) - 1911 & "年" & Mid(strCP07Add2M, 5, 2) & "月" & Right(strCP07Add2M, 2) & "日"
   End If
   '2.申請文件:
   If strCP07Add2M = "" Then strCP07Add2M = "　年　月　日"
   strTmp = ""
   Select Case Text9
      Case 翻譯, 檢視中說, 核對中說格式
         'Add By Sindy 2019/1/3 新型
         If pa(8) = 2 Then
            If Text6 = "2" Then '第2次延期
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','延期前期限','" & strPA10Add4M & "前。')"
               strTmp = "因「尚在準備中，不克於期限內補呈」之原因，「摘要、專利說明書、申請專利範圍及圖式」文件申請延期「至" & strPA10Add6M & "」補正。"
            Else
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','延期前期限','" & strPA10Add2M & "前。')"
               strTmp = "因「尚在準備中，不克於期限內補呈」之原因，「摘要、專利說明書、申請專利範圍及圖式」文件申請延期「至" & strPA10Add4M & "」補正。"
            End If
         Else
         '2019/1/3 END
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','延期前期限','" & strPA10Add4M & "前。')"
            strTmp = "因「尚在準備中，不克於期限內補呈」之原因，「摘要、專利說明書、申請專利範圍及圖式」文件申請延期「至" & strPA10Add6M & "」補正。"
         End If
      Case 製作中說
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','延期前期限','" & strPA10Add4M & "前。')"
         strTmp = "因「尚在準備中，不克於期限內補呈」之原因，「設計專利說明書及圖式」文件申請延期「至" & strPA10Add6M & "」補正。"
      Case 補文件
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','延期前期限','" & strPA10Add4M & "前。')"
         '請抓延期進度檔的備註
         strTmp = "因「文件尚在準備中，不克於期限內補呈」之原因，「" & cp(64) & "」文件申請延期「至" & strPA10Add6M & "」補正。"
      'Modify By Sindy 2019/8/8 + 239.擇一申復 同申復
      Case 申復, "239" '申復, 擇一申復
         'Add By Sindy 2020/4/20
         '抓相關總收文號的收文日做計算
         'Ex:FCP-060972;故依專利法施行細則第六條規定申請延期3個月，並將於近期內補呈相關資料或修正，務使本專利申請案所請內容儘量符合專利法之規定」之原因，「相關資料或修正」文件申請延期「至109年07月31日」補正。
         strExc(0) = "select cp05,cp43,cp10 from caseprogress" & _
                     " where CP09='" & strReceiveNo & "' AND CP43 is not null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(0) = "select cp05,cp43,cp10 from caseprogress" & _
                        " where CP09='" & RsTemp.Fields("cp43") & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               txtCP07 = RsTemp.Fields("cp05") - 19110000
               txtCP07.Tag = "CP05"
            End If
         End If
         '2020/4/20 END
         'Add By Sindy 2019/8/15 Ex:FCP-060454
         If pa(8) = "2" Then '新型
            intMonth = "2"
            If Val(txtCP07) > 0 Then
               'Add By Sindy 2020/4/20 Ex:FCP-060972
               If txtCP07.Tag = "CP05" Then
                  strExc(10) = CompDate(1, 4, DBDATE(txtCP07))
               Else
               '2020/4/20 END
                  strExc(10) = CompDate(1, 2, DBDATE(txtCP07))
               End If
            End If
         Else
         '2019/8/15 END
            intMonth = "3"
            If Val(txtCP07) > 0 Then
               'Add By Sindy 2020/4/20 Ex:FCP-060972
               If txtCP07.Tag = "CP05" Then
                  strExc(10) = CompDate(1, 6, DBDATE(txtCP07))
               Else
               '2020/4/20 END
                  strExc(10) = CompDate(1, 3, DBDATE(txtCP07))
               End If
            End If
         End If
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','延期前期限','文到次日起" & intMonth & "個月內。')"
         '原審查意見函的法定期限+3個月
         strExc(1) = "　年　月　日"
         If Val(txtCP07) > 0 Then
            'strExc(1) = CompDate(1, 3, DBDATE(txtCP07))
            'strExc(1) = Left(strExc(1), 4) - 1911 & "年" & Mid(strExc(1), 5, 2) & "月" & Right(strExc(1), 2) & "日"
            strExc(1) = Left(strExc(10), 4) - 1911 & "年" & Mid(strExc(10), 5, 2) & "月" & Right(strExc(10), 2) & "日"
         End If
         strTmp = "因「申請人正積極準備相關資料，由於本案有其複雜性且事關權益，申請人正與公司內部人員研討，實難於　鈞局指定期限內提呈相關資料及修正。故依專利法施行細則第六條規定申請延期" & intMonth & "個月，並將於近期內補呈相關資料或修正，務使本專利申請案所請內容儘量符合專利法之規定」之原因，「相關資料或修正」文件申請延期「至" & strExc(1) & "」補正。"
      Case "107" '再審查
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請修正','否')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請誤譯訂正','否')"
         '第1次延期出再審申請書
         If Val(Text6) = 1 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','延期前期限','" & strCP07Add2M & "前。')"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','再審查申請規費及再審查理由書請准予容後補呈。')"
            
            '1002.核駁-新申請案
            Dim strED08 As String, strSendData1 As String, strSendData2 As String, strCP05 As String
            strExc(0) = "select cp05,cp08,ed08 from caseprogress,edocument" & _
                        " where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                        " AND ed11(+)=cp09 AND cp10='1002'" & _
                        " AND cp43 in(select cp09 from caseprogress where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' and cp10 in(" & NewCasePtyList & ")) ORDER BY CP05 DESC"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If Not IsNull(RsTemp("ED08")) Then
                  strED08 = RsTemp("ED08") - 19110000
                  If Not IsNull(RsTemp("cp08")) Then
                     strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
                     strSendData1 = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
                     strExc(0) = Replace(strExc(0), strSendData1 & "字第", "")
                     strSendData2 = Mid(strExc(0), 1, InStr(strExc(0), "號") - 1)
                  End If
               End If
               If Not IsNull(RsTemp("cp05")) Then
                  strCP05 = RsTemp("cp05") - 19110000
               End If
            End If
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','初審官方發文日','" & ChangeTStringToTDateString(strED08) & "')"
      
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','初審智專字','" & strSendData1 & "')"
      
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','初審發文號','" & strSendData2 & "')"
            
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','初審收文日期','" & ChangeTStringToTDateString(strCP05) & "')"
         '再審第2次延期:請抓1004.延期受理+2個月
         Else
            strExc(0) = "select CP27 from caseprogress c1" & _
               " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
               " and cp10='404' and cp27 is not null and cp57 is null" & _
               " and CP43 in(select c2.CP09 from caseprogress c2 where c2.cp09=c1.cp43 and c2.cp10='107')"
            strExc(0) = strExc(0) & " union " & _
               " select c2.CP27 from caseprogress c1,nextProgress,caseprogress c2" & _
               " where c1.cp09='" & strReceiveNo & "' and c1.cp10='404' and c1.cp43 is not null and c1.cp30 is not null" & _
               " and np01(+)=c1.cp43 and np22(+)=c1.cp30 and np07='107'" & _
               " and np01=c2.cp43(+) and c2.cp10='404' and c2.cp27 is not null and c2.cp57 is null"
            strExc(0) = strExc(0) & " order by CP27 asc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               '第1次延期發文日+4個月
               strCP27Add4M = CompDate(1, 4, RsTemp.Fields("cp27"))
               strCP27Add4M = Left(strCP27Add4M, 4) - 1911 & "年" & Mid(strCP27Add4M, 5, 2) & "月" & Right(strCP27Add4M, 2) & "日"
               '第1次延期發文日+6個月
               strCP27Add6M = CompDate(1, 6, RsTemp.Fields("cp27"))
               strCP27Add6M = Left(strCP27Add6M, 4) - 1911 & "年" & Mid(strCP27Add6M, 5, 2) & "月" & Right(strCP27Add6M, 2) & "日"
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','延期前期限','" & strCP27Add4M & "前。')"
            End If
            strTmp = "因「相關資料或修正，不克於期限內補呈」之原因，「再審查理由書及再審查申請規費」文件申請延期「至" & strCP27Add6M & "」補正。"
         End If
   End Select
   If strTmp <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請文件','" & strTmp & "')"
   End If
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳費金額','" & Val(txtCP84) & "')"
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-基本資料表','" & m_CaseNo & ".contact.pdf')"
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function

Private Sub txtCP84_GotFocus()
   TextInverse txtCP84
   CloseIme
End Sub

Private Sub txtCP84_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Private Sub txtCP84_Validate(Cancel As Boolean)
'   '台灣
'   If pa(9) = "000" Then
'      If Val(txtCP84.Text) <> Val(cp(17)) And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
'         If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & cp(17) & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
'            txtCP84.Tag = txtCP84.Text
'         Else
'            txtCP84_GotFocus
'            Cancel = True
'         End If
'      End If
'   End If
'End Sub

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟FCP0xxxxx.新案性質.案件名稱.doc
Private Sub CmdPA174_Click()

    If pa(1) = "" Or pa(2) = "" Or pa(3) = "" Or pa(4) = "" Then Exit Sub
    If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = True Then
    End If
    
End Sub
