VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm1106 
   BorderStyle     =   1  '單線固定
   Caption         =   "聯絡單列印及E-Mail"
   ClientHeight    =   5750
   ClientLeft      =   1180
   ClientTop       =   3300
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8950
   Begin VB.CommandButton cmdInput 
      Caption         =   "多案案號輸入"
      Height          =   350
      Left            =   7380
      TabIndex        =   52
      Top             =   2820
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Frame Frame2 
      Caption         =   "4:專利缺文件"
      Height          =   1575
      Left            =   6960
      TabIndex        =   46
      Top             =   4170
      Width           =   1965
      Begin VB.CheckBox Check1 
         Caption         =   "身份證影本"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   51
         Top             =   1290
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "優先權證明文件"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   50
         Top             =   1035
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "讓與文件"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   49
         Top             =   780
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "簽署文件"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   48
         Top             =   525
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "委任書"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   47
         Top             =   270
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   435
      Left            =   2880
      TabIndex        =   41
      Top             =   750
      Width           =   6045
      Begin VB.TextBox txtRecvNo 
         Height          =   270
         Left            =   810
         MaxLength       =   9
         TabIndex        =   43
         Top             =   60
         Width           =   1365
      End
      Begin VB.CommandButton cmdSelCp09 
         Caption         =   "選擇總收文號"
         Height          =   300
         Left            =   2190
         TabIndex        =   42
         Top             =   60
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   $"frm1106.frx":0000
         ForeColor       =   &H000000FF&
         Height          =   420
         Index           =   16
         Left            =   3570
         TabIndex        =   45
         Top             =   30
         Width           =   2280
      End
      Begin VB.Label Label3 
         Caption         =   "總收文號:"
         Height          =   225
         Left            =   30
         TabIndex        =   44
         Top             =   90
         Width           =   795
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   0
      Left            =   4980
      MaxLength       =   1
      TabIndex        =   37
      Text            =   "2"
      Top             =   150
      Width           =   225
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "清除(&C)"
      Height          =   350
      Index           =   3
      Left            =   6870
      TabIndex        =   15
      Top             =   105
      Width           =   800
   End
   Begin VB.OptionButton Option1 
      Caption         =   "小姐"
      Height          =   225
      Index           =   1
      Left            =   3420
      TabIndex        =   10
      Top             =   4140
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.OptionButton Option1 
      Caption         =   "先生"
      Height          =   225
      Index           =   0
      Left            =   2640
      TabIndex        =   9
      Top             =   4140
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "發E-Mail(&S)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   2
      Left            =   5670
      TabIndex        =   14
      Top             =   105
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   7
      Left            =   5250
      TabIndex        =   11
      Top             =   4110
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   5
      Left            =   1365
      MaxLength       =   8
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   8
      Left            =   1365
      MaxLength       =   30
      TabIndex        =   4
      Top             =   510
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Left            =   1365
      MaxLength       =   1
      TabIndex        =   6
      Top             =   2520
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1365
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   2850
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   1
      Left            =   1365
      MaxLength       =   3
      TabIndex        =   0
      Top             =   180
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   2
      Left            =   1845
      MaxLength       =   6
      TabIndex        =   1
      Top             =   180
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   3
      Left            =   2685
      MaxLength       =   1
      TabIndex        =   2
      Top             =   180
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   4
      Left            =   2925
      MaxLength       =   2
      TabIndex        =   3
      Top             =   180
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   7710
      TabIndex        =   16
      Top             =   105
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印(&P)　　份"
      Enabled         =   0   'False
      Height          =   350
      Index           =   0
      Left            =   4200
      TabIndex        =   13
      Top             =   105
      Width           =   1425
   End
   Begin MSForms.TextBox txtReviewer 
      Height          =   300
      Left            =   1365
      TabIndex        =   8
      Top             =   4110
      Visible         =   0   'False
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstMailCC 
      Height          =   900
      Left            =   1350
      TabIndex        =   38
      Top             =   3180
      Width           =   4035
      VariousPropertyBits=   746586139
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "7117;1587"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   960
      Left            =   1365
      TabIndex        =   12
      Top             =   4740
      Width           =   5520
      VariousPropertyBits=   -1466939365
      ScrollBars      =   2
      Size            =   "9737;1693"
      FontName        =   "細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "PS：非ＦＣ案件且非文件公簽證聯絡單，受文者預設為最後收文智權人員及承辦人！"
      ForeColor       =   &H000000FF&
      Height          =   540
      Index           =   15
      Left            =   5880
      TabIndex        =   40
      Top             =   3240
      Width           =   2520
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "(可複選)"
      Height          =   255
      Index           =   14
      Left            =   540
      TabIndex        =   39
      Top             =   3420
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "副本："
      Height          =   210
      Index           =   10
      Left            =   540
      TabIndex        =   36
      Top             =   3180
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "審查員："
      Height          =   255
      Index           =   12
      Left            =   345
      TabIndex        =   35
      Top             =   4140
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "申請人："
      Height          =   255
      Index           =   4
      Left            =   345
      TabIndex        =   34
      Top             =   2190
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "註冊號數："
      Height          =   255
      Index           =   6
      Left            =   345
      TabIndex        =   33
      Top             =   840
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "申請案號："
      Height          =   255
      Index           =   5
      Left            =   345
      TabIndex        =   32
      Top             =   510
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "本所案號："
      Height          =   270
      Index           =   0
      Left            =   330
      TabIndex        =   31
      Top             =   180
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "受文者："
      Height          =   270
      Index           =   8
      Left            =   540
      TabIndex        =   30
      Top             =   2850
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "聯絡單備註："
      Height          =   270
      Index           =   9
      Left            =   150
      TabIndex        =   29
      Top             =   4770
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "種　類："
      Height          =   270
      Index           =   2
      Left            =   600
      TabIndex        =   28
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "主旨："
      Height          =   255
      Index           =   13
      Left            =   780
      TabIndex        =   27
      Top             =   4440
      Width           =   4830
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "分機："
      Height          =   255
      Index           =   11
      Left            =   4590
      TabIndex        =   26
      Top             =   4140
      Visible         =   0   'False
      Width           =   630
   End
   Begin MSForms.Label LabTM23 
      Height          =   255
      Left            =   1380
      TabIndex        =   25
      Top             =   2190
      Width           =   7380
      VariousPropertyBits=   27
      Size            =   "13017;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "（1:空白聯絡單 2:信件退回聯絡單 3:FCT審查員電話通知 4:專利缺文件5.文件公簽證聯絡單）"
      Height          =   180
      Index           =   3
      Left            =   1770
      TabIndex        =   24
      Top             =   2550
      Width           =   7245
   End
   Begin VB.Label lblSaleZone 
      Height          =   270
      Left            =   5730
      TabIndex        =   22
      Top             =   2850
      Width           =   1620
   End
   Begin MSForms.Label lblSaleName 
      Height          =   270
      Left            =   2580
      TabIndex        =   23
      Top             =   2850
      Width           =   1620
      VariousPropertyBits=   27
      Size            =   "2857;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "受文者部門："
      Height          =   270
      Index           =   7
      Left            =   4470
      TabIndex        =   21
      Top             =   2850
      Width           =   1185
   End
   Begin MSForms.Label lblCaseName 
      Height          =   270
      Index           =   2
      Left            =   855
      TabIndex        =   20
      Top             =   1875
      Width           =   7905
      VariousPropertyBits=   27
      Caption         =   "(日)："
      Size            =   "13944;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   270
      Index           =   1
      Left            =   855
      TabIndex        =   19
      Top             =   1590
      Width           =   7905
      VariousPropertyBits=   27
      Caption         =   "(英)："
      Size            =   "13944;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   270
      Index           =   0
      Left            =   855
      TabIndex        =   18
      Top             =   1290
      Width           =   7905
      VariousPropertyBits=   27
      Caption         =   "(中)："
      Size            =   "13944;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱"
      Height          =   270
      Index           =   1
      Left            =   60
      TabIndex        =   17
      Top             =   1290
      Width           =   1185
   End
End
Attribute VB_Name = "frm1106"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/5/19 改成Form2.0 (lblSaleName,txtReviewer...)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Dim intWhere As Integer, strReceiveNo As String
Dim m_blnTxtValidate As Boolean
'Add By Cheng 2003/04/07
Dim m_PrtOrientation As Integer '列印方向
Dim m_PrtScaleMode As Integer '列印座標單位
Dim m_dblTop As Double '上邊界
Dim m_dblLeft As Double '左邊界
Dim m_dblTitleHeight As Double '表頭高度
Dim m_dblLine As Double '行數
Dim m_dblLineHeight As Double '行高
Dim m_dblBetweenLine As Double '行間空隙
Dim m_dblLineHeight1 As Double '行高
Dim m_dblBetweenLine1 As Double '行間空隙
Dim m_strSQLA As String
Dim m_rsA As New ADODB.Recordset
Dim m_Combo1ST06 As String     '2008/12/18 add by sonia 受文者之所別
Dim m_strToCCNo As String      'Add By Sindy 2010/11/19
Dim s_MailCC As String         'add by sonai 2014/6/3

'Add By Sindy 2020/1/16
Private Sub Check1_Click(Index As Integer)
   Call Text2_Validate(False)
End Sub

Private Sub cmdok_Click(Index As Integer)
'edit by nickc 2007/02/06 不用 dll 了
'Dim objPrintDllPublic As New clsPrintPublic
Dim rsA  As New ADODB.Recordset
Dim StrSQLa As String
Dim strDept As String
Dim strTo As String, strSubject As String, strContent As String
Dim strToCC As String, strFAX As String, strTEL As String
Dim i As Integer
Dim intMaxEEP02 As Integer
   
On Error GoTo ErrorHandler
   
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   If (Index = 0 Or Index = 2) And FMP2open = True Then
     If PUB_FMPtoCheck(0, 1, Pub_strUserST05, Text1(1), Text1(2), Text1(3), Text1(4)) = False Then
        Me.Text1(2).SetFocus
        TextInverse Me.Text1(2)
        Screen.MousePointer = vbDefault
        Exit Sub
     End If
   End If
    
   'Add By Sindy 2020/1/16
   If (Index = 0 Or Index = 2) And Text2 = "4" And _
      Check1(0).Value = 0 And _
      Check1(1).Value = 0 And Check1(2).Value = 0 And _
      Check1(3).Value = 0 And Check1(4).Value = 0 Then
      If MsgBox("確定不勾選專利缺文件事項嗎？", vbQuestion + vbYesNo + vbDefaultButton2, "詢問") = vbNo Then
         Exit Sub
      End If
   End If
   '2020/1/16 END
         
   Select Case Index
      Case 0 '確定
         ' 設定滑鼠游標為等待狀態
         Screen.MousePointer = vbHourglass
         If Me.Text1(1).Text = "" Then
            MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
            Me.Text1(1).SetFocus
            TextInverse Me.Text1(1)
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         If Me.Text1(2).Text = "" Then
            MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
            Me.Text1(2).SetFocus
            TextInverse Me.Text1(2)
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
        
         If Me.Combo1.Text = "" Then
            'Modify By Cheng 2003/04/18
'            MsgBox "此本所案號無智權人員資料, 無法列印聯絡單!!!", vbExclamation + vbOKOnly
            MsgBox "此本所案號無受文者資料, 無法列印聯絡單!!!", vbExclamation + vbOKOnly
            Me.Text1(1).SetFocus
            Text1_GotFocus 1
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Combo1_Change
         '列印聯絡單
         'Added by Lydia 2023/12/25
         If strSrvDate(1) >= 新部門啟用日 Then
             StrSQLa = "Select NVL(A0923,A0902) AS A0902 From Staff,ACC090,ACC090NEW WHERE ST03=A0901 AND ST01='" & strUserNum & "' AND ST93=A0921(+) "
         Else
         'end 2023/12/25
             StrSQLa = "Select A0902 From Staff,ACC090 WHERE ST03=A0901 AND ST01='" & strUserNum & "'"
         End If
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            strDept = "" & rsA.Fields(0).Value
         Else
            strDept = ""
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         
        If m_rsA.RecordCount > 0 Then
            'Add By Cheng 2003/04/03
            '取得預設印表機設定值
            m_PrtOrientation = Printer.Orientation
            m_PrtScaleMode = Printer.ScaleMode
            '重新設定印表機
            Printer.PaperSize = vbPRPSA4
            Printer.Orientation = vbPRORPortrait
            Printer.ScaleMode = vbCentimeters
            '列印聯絡單
            InitPrtPosition 0.5, 0.5
            PrintContactSheet strDept
'            '2008/12/18 ADD BY SONIA CFP需求,專利處程序印則分所印二份,不分受文者之所別的
'            If GetStaffDepartment(strUserNum) = "P12" Then
'                '列印聯絡單
'                Select Case Text1(1)  '2009/9/15 CFP的二份分二頁,P不變
'                  Case "P", "PS"
'                     InitPrtPosition 13.5, 0.5
'                     PrintContactSheet strDept
'                  Case Else
'                     Printer.NewPage
'                     InitPrtPosition 0.5, 0.5
'                     PrintContactSheet strDept
'               End Select
'            End If
'            '2008/12/18 end
            'Modify By Sindy 2010/10/14
            If Val(Text1(0).Text) > 1 Then '列印多份
               If Text1(1) = "P" Or Text1(1) = "PS" Then
                  For i = 2 To Val(Text1(0))
                     If (i Mod 2) = 0 Then
                        InitPrtPosition 13.5, 0.5 '偶數
                     Else
                        Printer.NewPage
                        InitPrtPosition 0.5, 0.5 '單數
                     End If
                     PrintContactSheet strDept
                  Next i
               Else
                  For i = 2 To Val(Text1(0))
                     Printer.NewPage
                     InitPrtPosition 0.5, 0.5
                     PrintContactSheet strDept
                  Next i
               End If
            End If
            '2010/10/14 End
            Printer.EndDoc
            '還原預設印表機設值
            Printer.Orientation = m_PrtOrientation
            Printer.ScaleMode = m_PrtScaleMode
        End If
         
'         objPrintDllPublic.PrtMail_1 Me.Text1(1).Text, Right("000000" & Me.Text1(2).Text, 6), Right("0" & Me.Text1(3).Text, 1), Right("00" & Me.Text1(4).Text, 2), _
'                                    strUserName, _
'                                    IIf(Me.Text3.Text = "Y", True, False), _
'                                    Me.Text4.Text, _
'                                    strDept, _
'                                    DBYEAR(ServerDate) - 1911 & "年" & Format(DBMONTH(ServerDate), "00") & "月" & Format(DBDAY(ServerDate), "00") & "日", _
'                                    Me.lblSaleZone.Caption, Me.lblSaleName.Caption, _
'                                    Right(Me.lblCaseName(0).Caption, Len(Me.lblCaseName(0).Caption) - 4), _
'                                    Right(Me.lblCaseName(1).Caption, Len(Me.lblCaseName(1).Caption) - 4), _
'                                    Right(Me.lblCaseName(2).Caption, Len(Me.lblCaseName(2).Caption) - 4)
'                                    DoEvents
'         Set objPrintDllPublic = Nothing: DoEvents
         ' 設定滑鼠游標為預設
         Screen.MousePointer = vbDefault
         Me.Text1(1).SetFocus
         Text1_GotFocus 1
         
      Case 1 '結束
         Unload Me
         
      'Add By Sindy 2010/10/13
      Case 2 '發E-Mail
         '2014/1/20 add by sonia
         If Text2 = "" Then
            MsgBox "種類不可空白！"
            Text2_GotFocus
            Exit Sub
         '2014/1/20 end
         ElseIf Text2 = "3" Then
            If Trim(txtReviewer) = "" Then
               MsgBox "審查員不可空白！"
               Text1_GotFocus 6
               Exit Sub
            End If
            If Trim(Text1(7)) = "" Then
               MsgBox "分機不可空白！"
               Text1_GotFocus 7
               Exit Sub
            End If
         End If
        
         strTo = ""
         strToCC = ""
         '正本
         strTo = Trim(Combo1.Text)
         If Trim(strTo) = "" Then
            MsgBox "收件人空白，無法寄送！"
            Exit Sub
         End If
         If strTo = "F4103" Then
            If IsNull(m_rsA.Fields(1).Value) Then
               strTo = "80030" '洪琬姿
            Else
               If Trim(m_rsA.Fields(1).Value) = "日本" Then
                  strTo = "78011" '葉易雲
               Else
                  strTo = "80030" '洪琬姿
               End If
            End If
         End If
         '副本
         For i = 0 To lstMailCC.ListCount - 1
            If lstMailCC.Selected(i) = True Then
               'MODIFY BY SONIA 2014/6/3 因陳經理要求要顯示部門名稱,故需截取員工編號
               'If strToCC = "" Then
               '   strToCC = Left(Trim(lstMailCC.List(i)), 5)
               'Else
               '   strToCC = strToCC & ";" & Left(Trim(lstMailCC.List(i)), 5)
               'End If
               s_MailCC = Trim(Mid(lstMailCC.List(i), InStr(lstMailCC.List(i), " ") + 1, 5))
               If strToCC = "" Then
                  strToCC = s_MailCC
               Else
                  strToCC = strToCC & ";" & s_MailCC
               End If
               'END 2014/6/3
            End If
         Next
         
         'Add By Sindy 2020/1/10 若聯絡單未選擇總收文號，執行動作前請再彈訊息讓user確認
         '是否確定此聯絡單"不"產生承辦歷程，選項要有是、否，此訊息選項預設值請設在否。
         If Frame1.Visible = True And Trim(txtRecvNo) = "" Then
            If MsgBox("是否確定此聯絡單【不】產生承辦歷程？", vbQuestion + vbYesNo + vbDefaultButton2, "詢問") = vbNo Then
               txtRecvNo.SetFocus
               Exit Sub
            End If
         End If
         '2020/1/10 END
         
         Screen.MousePointer = vbHourglass
        
        strContent = ""
        strFAX = ""
        strTEL = ""
        If Not IsNull(m_rsA.Fields("CU18").Value) Then
            strFAX = m_rsA.Fields("CU18").Value
        End If
        If Not IsNull(m_rsA.Fields("CU16").Value) Then
            strTEL = m_rsA.Fields("CU16").Value
        End If
        If m_rsA.RecordCount > 0 Then
            strSubject = Mid(Trim(Label1(13)), 4, Len(Trim(Label1(13))))
            If Text2 = "3" Then
               strContent = strContent & _
                                   "審查員：" + txtReviewer + "  " + IIf(Option1(0).Value = True, "先生", IIf(Option1(1).Value = True, "小姐", "")) + vbCrLf + _
                                   "分　機：" + Text1(7) + vbCrLf + vbCrLf
            End If
               strContent = strContent & _
                                   "本所案號：" + Text1(1) + "-" + Text1(2) + "-" + Text1(3) + "-" + Text1(4) + vbCrLf + _
                                   "案件名稱(中)：" + IIf(Len(Trim(lblCaseName(0))) > 4, Mid(Trim(lblCaseName(0)), 5, Len(Trim(lblCaseName(0)))), "") + vbCrLf + _
                                   "案件名稱(英)：" + IIf(Len(Trim(lblCaseName(1))) > 4, Mid(Trim(lblCaseName(1)), 5, Len(Trim(lblCaseName(1)))), "") + vbCrLf + _
                                   "案件名稱(日)：" + IIf(Len(Trim(lblCaseName(2))) > 4, Mid(Trim(lblCaseName(2)), 5, Len(Trim(lblCaseName(2)))), "") + vbCrLf + _
                                   "申請人：" + LabTM23 + vbCrLf + _
                                   "FAX：" + strFAX + vbCrLf + _
                                   "TEL：" + strTEL + vbCrLf + _
                                   "申請國家：" + "" + m_rsA.Fields(1).Value + vbCrLf
            If Trim(Text4) <> "" Then
               strContent = strContent + "備　　註：" + vbCrLf
               strContent = strContent + Text4 + vbCrLf
            End If
            
            'Added by Lydia 2022/08/12
            If Text2 = "3" Then
                  If FormSave = False Then
                      Screen.MousePointer = vbDefault
                      Exit Sub
                  End If
            Else
            'end 2022/08/12
               PUB_SendMail strUserNum, strTo, "", strSubject, strContent, "", , , , , strToCC
               's = MsgBox("郵件已送出", , "MAIL!!")
               'Add By Sindy 2020/1/10 新增聯絡歷程
               If Frame1.Visible = True And Trim(txtRecvNo) <> "" And bolMailSendOk = True Then
                  '取得最大序號
                  intMaxEEP02 = 0
                  strSql = "select eep02 From empelectronprocess where eep01='" & txtRecvNo & "' order by eep02 desc"
                  intI = 1
                  CheckOC3
                  Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     AdoRecordSet3.MoveFirst
                     If AdoRecordSet3.RecordCount > 0 Then
                        intMaxEEP02 = AdoRecordSet3.Fields(0)
                     End If
                  End If
                  strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08,eep10) values(" & _
                           CNULL(txtRecvNo) & "," & intMaxEEP02 + 1 & ",'" & strUserNum & "'," & _
                           CNULL(EMP_聯絡) & "," & _
                           CNULL(strTo) & "," & _
                           strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & "," & CNULL(ChgSQL(strContent)) & "," & CNULL(Replace(strToCC, ";", ",")) & ")"
                  cnnConnection.Execute strSql
               End If
               '2020/1/10 END
            End If 'Added by Lydia 2022/08/12
        End If
         
         '2011/10/21 add by sonia 發完清畫面
         Call ClearCol(False)
         Text4 = ""
         Me.cmdInput.Tag = "" 'Added by Lydia 2022/08/12
         '2011/10/21 end
         Screen.MousePointer = vbDefault
         Me.Text1(1).SetFocus
         Text1_GotFocus 1
         
      'Add By Sindy 2010/10/13
      Case 3 '清除
         Call ClearCol(False)
         Me.cmdInput.Tag = "" 'Added by Lydia 2022/08/12
   End Select
   
   Exit Sub
   
ErrorHandler:
    MsgBox Err.Number & Err.Description
    'Add By Cheng 2002/12/12
    ' 設定滑鼠游標為預設
    Screen.MousePointer = vbDefault
End Sub

'Add By Sindy 2020/1/10
Private Sub cmdSelCp09_Click()
Dim rsRead As New ADODB.Recordset
Dim sqlB As String
    
   If Trim(Text1(1)) <> "" And Trim(Text1(2)) <> "" Then
      Me.Tag = ""
      Text1(3).Text = IIf(Text1(3) = "", "0", Text1(3))
      Text1(4).Text = IIf(Text1(4) = "", "00", Text1(4))
      '專利
      sqlB = "select '' V," & SQLDate("CP05") & " as 收文日,cp09 as 總收文號,decode(pa09,'000',cpm03,cpm04) as 案件性質,s1.st02 as 承辦人,s2.st02 as 智權人員," & SQLDate("CP27") & " as 發文日,cp05,cp66,cp67,cp09 " & _
             "from caseprogress,casepropertymap,staff s1,staff s2,patent " & _
             "where cp01='" & Text1(1) & "' and cp02='" & Text1(2) & "' and cp03='" & Text1(3) & "' and cp04='" & Text1(4) & "' " & _
             "and cp159=0 and cp01=cpm01(+) and cp10=cpm02(+) " & _
             "and cp14=s1.st01(+) and cp13=s2.st01(+) " & _
             "and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 "
      '商標
      sqlB = sqlB & " union " & _
             "select '' V," & SQLDate("CP05") & " as 收文日,cp09 as 總收文號,decode(tm10,'000',cpm03,cpm04) as 案件性質,s1.st02 as 承辦人,s2.st02 as 智權人員," & SQLDate("CP27") & " as 發文日,cp05,cp66,cp67,cp09 " & _
             "from caseprogress,casepropertymap,staff s1,staff s2,trademark " & _
             "where cp01='" & Text1(1) & "' and cp02='" & Text1(2) & "' and cp03='" & Text1(3) & "' and cp04='" & Text1(4) & "' " & _
             "and cp159=0 and cp01=cpm01(+) and cp10=cpm02(+) " & _
             "and cp14=s1.st01(+) and cp13=s2.st01(+) " & _
             "and cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 "
      '服務
      sqlB = sqlB & " union " & _
             "select '' V," & SQLDate("CP05") & " as 收文日,cp09 as 總收文號,decode(sp09,'000',cpm03,cpm04) as 案件性質,s1.st02 as 承辦人,s2.st02 as 智權人員," & SQLDate("CP27") & " as 發文日,cp05,cp66,cp67,cp09 " & _
             "from caseprogress,casepropertymap,staff s1,staff s2,servicepractice " & _
             "where cp01='" & Text1(1) & "' and cp02='" & Text1(2) & "' and cp03='" & Text1(3) & "' and cp04='" & Text1(4) & "' " & _
             "and cp159=0 and cp01=cpm01(+) and cp10=cpm02(+) " & _
             "and cp14=s1.st01(+) and cp13=s2.st01(+) " & _
             "and cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 "
      sqlB = sqlB & " ORDER BY CP05 DESC, CP66 DESC, CP67 DESC, CP09 DESC"
      intI = 0
      Set rsRead = ClsLawReadRstMsg(intI, sqlB)
      If intI = 1 Then
         Set frm880012.grdDataList.Recordset = rsRead
         Set frm880012.fmParent = Me
         frm880012.iTyp = "1"
         frm880012.Show vbModal
         If Me.Tag <> "" Then
            txtRecvNo.Text = Me.Tag
            txtRecvNo.SetFocus
         End If
      End If
   Else
      MsgBox "請先輸入本所案號！", vbExclamation, "警告！"
      If Me.Text1(1).Enabled = True Then Me.Text1(1).SetFocus
   End If
End Sub

Private Sub Combo1_Change()
   Me.lblSaleName.Caption = GetStaffName("" & Me.Combo1.Text)
   
   If Me.lblSaleName.Caption <> "" Then
      'Added by Lydia 2023/12/25
      If strSrvDate(1) >= 新部門啟用日 Then
         Me.lblSaleZone.Caption = GetDeptNameA0922("" & Me.Combo1.Text)
      Else
      'end 2023/12/25
          Me.lblSaleZone.Caption = A0902Query(GetStaffDepartment("" & Me.Combo1.Text))
      End If
      m_Combo1ST06 = PUB_GetST06(Me.Combo1.Text)
      'Added by Lydia 2019/10/30 受文者的分機
      If Me.Text2 = "3" Then
           Me.lblSaleName.Caption = Me.lblSaleName.Caption & " ＃ " & Pub_GetStaffExtn("" & Me.Combo1.Text)
      End If
   Else
      Me.lblSaleZone.Caption = ""
   End If
End Sub

Private Sub Combo1_Click()
   If Me.Combo1.Text <> "" Then
      Me.lblSaleName.Caption = GetStaffName("" & Me.Combo1.Text)
      If Me.lblSaleName.Caption <> "" Then
         Me.lblSaleZone.Caption = A0902Query(GetStaffDepartment("" & Me.Combo1.Text))
         'Added by Lydia 2019/10/30 受文者的分機
         If Me.Text2 = "3" Then
             Me.lblSaleName.Caption = Me.lblSaleName.Caption & " ＃ " & Pub_GetStaffExtn("" & Me.Combo1.Text)
         End If
      Else
         Me.lblSaleZone.Caption = ""
      End If
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Me.Combo1.Text = ""
   
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
   
   'Add By Sindy 2020/1/10 P1.專利處 P2.商標處
   If Left(Pub_StrUserSt03, 2) = "P1" Or Left(Pub_StrUserSt03, 2) = "P2" Then
      Frame1.Visible = True
   Else
      Frame1.Visible = False
   End If
   '2020/1/10 END
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call PUB_SendMailCache
   
   Set frm1106 = Nothing
End Sub

'Add By Sindy 2010/10/13
Private Sub ClearCol(bolQuery As Boolean)
   If bolQuery = False Then
      Me.Text1(1) = ""
      Me.Text1(2) = ""
      Me.Text1(3) = ""
      Me.Text1(4) = ""
      Me.Text1(5) = ""
      Me.Text1(8) = ""
   End If
   Me.lblCaseName(0).Caption = "(中)："
   Me.lblCaseName(1).Caption = "(英)："
   Me.lblCaseName(2).Caption = "(日)："
   Me.Label1(13).Caption = "主旨："
   Me.Combo1.Clear
   Me.lblSaleZone.Caption = ""
   Me.lblSaleName.Caption = ""
   Me.lstMailCC.Clear
   Me.Text2 = ""
   Me.cmdOK(0).Enabled = False
   Me.cmdOK(2).Enabled = False
   Me.Text1(0).Enabled = False
   Me.txtRecvNo = "" 'Add By Sindy 2020/3/20
   HideCol
End Sub

'Add By Sindy 2010/10/13
Private Sub HideCol()
   '審查員
   Label1(12).Visible = False
   txtReviewer.Visible = False
   Option1(0).Visible = False
   Option1(1).Visible = False
   '分機
   Label1(11).Visible = False
   Text1(7).Visible = False
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Select Case Index
Case 4
   If m_blnTxtValidate = False Then
      Me.Text1(1).SetFocus
      m_blnTxtValidate = True
      Exit Sub
   End If
End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String, strSql As String
Dim strTemp1
Dim strTemp2
Dim ii As Integer
Dim jj As Integer
Dim ss As Integer
Dim m_Dept As String
Dim tmpObj As Object 'Added by Lydia 2022/08/12

Select Case Index
Case 1 '系統類別
   If Text1(Index) <> "" Then
      'Modify By Sindy 2010/10/29
      strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
      strTemp2 = Split(Replace(UCase(Text1(Index).Text), ",,", ""), ",")
      For ii = 0 To UBound(strTemp2)
          ss = 0
          For jj = 0 To UBound(strTemp1)
              If strTemp2(ii) = strTemp1(jj) Then
                  ss = 1
                  Exit For
              End If
          Next jj
          'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件(P)，但非此類案件時外專程序人員不可操作。
          If (Trim(Text1(1).Text) = "P" Or Trim(Text1(1).Text) = "PS") And PUB_FMPtoCheck(1, 0, Pub_strUserST05) = True Then Exit For
          
          If ss = 0 Then
             '開放FF案件之權限
             m_Dept = GetStaffDepartment(strUserNum)
             Select Case m_Dept
                Case "F21", "F23", "F61", "F81"  '開放F21,F23使用P,PS,CFP,CPS權限
                   If Text1(Index).Text = "P" Or Text1(Index).Text = "PS" Or _
                      Text1(Index).Text = "CFP" Or Text1(Index).Text = "CPS" Then
                      Exit For
                   End If
                Case "F10", "F11"    '開放F10,F11使用T權限
                   If Text1(Index).Text = "T" Then
                      Exit For
                   End If
             End Select
             '檢查跨部門權限
             If CheckSR09(strUserNum, Text1(Index), "Y", False, Text1(1), Text1(2), Text1(3), Text1(4)) = True Then
                Exit For
             End If
             ss = MsgBox(strUserName & " 沒有 " & strTemp2(ii) & " 的權限!! ", , "USER 權限問題")
             Text1(Index).SetFocus
             Call Text1_GotFocus(Index)
             Cancel = True
          End If
      Next ii
   End If
   
Case 4, 5, 8 '4.本所案號 5.註冊號數 8.申請案號
   'Added by Lydia 2022/08/12
   If Text1(1).Tag & Text1(2).Tag & Text1(3).Tag & Text1(4).Tag & Text1(5).Tag & Text1(8).Tag = Text1(1) & Text1(2) & Text1(3) & Text1(4) & Text1(5) & Text1(8) Then
        Exit Sub
   End If
   cmdInput.Tag = ""
   'end 2022/08/12
   'Add By Sindy 2010/10/13
   If Text1(1) <> "" And Text1(2) <> "" Then
      If Text1(3) = "" Then Text1(3) = "0"
      If Text1(4) = "" Then Text1(4) = "00"
        'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
        If FMP2open = True Then
          If PUB_FMPtoCheck(0, 1, Pub_strUserST05, Text1(1), Text1(2), Text1(3), Text1(4)) = False Then Exit Sub
        End If
   End If
   If Text1(Index) = "" Then Exit Sub
   Call ClearCol(True)
   '2010/10/13 End
   m_blnTxtValidate = True
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   strSql = ""
   
   'Add By Sindy 2010/10/13
   If Text1(1) <> "" And Text1(2) <> "" Then
      If strSql <> "" Then strSql = strSql & " and "
      strSql = "TM01='" & Text1(1) & "' and TM02='" & Text1(2) & "' and TM03='" & Text1(3) & "' and TM04='" & Text1(4) & "'"
   '2010/10/13 End
      StrSQLa = "Select PA05,PA06,PA07,PA01,PA02,PA03,PA04,'','' From PATENT Where PA01='" & Me.Text1(1).Text & "' AND PA02='" & Me.Text1(2).Text & "' AND PA03='" & IIf(Me.Text1(3).Text = "", "0", Me.Text1(3).Text) & "' AND PA04='" & IIf(Me.Text1(4).Text = "", "00", Me.Text1(4).Text) & "' "
      StrSQLa = StrSQLa & " union Select TM05,TM06,TM07,TM01,TM02,TM03,TM04,TM12,TM15 From TRADEMARK Where " & strSql
      StrSQLa = StrSQLa & " union Select LC05,LC06,LC07,LC01,LC02,LC03,LC04,'','' From LAWCASE Where LC01='" & Me.Text1(1).Text & "' AND LC02='" & Me.Text1(2).Text & "' AND LC03='" & IIf(Me.Text1(3).Text = "", "0", Me.Text1(3).Text) & "' AND LC04='" & IIf(Me.Text1(4).Text = "", "00", Me.Text1(4).Text) & "' "
      StrSQLa = StrSQLa & " union Select HC06,'','',HC01,HC02,HC03,HC04,'','' From HIRECASE Where HC01='" & Me.Text1(1).Text & "' AND HC02='" & Me.Text1(2).Text & "' AND HC03='" & IIf(Me.Text1(3).Text = "", "0", Me.Text1(3).Text) & "' AND HC04='" & IIf(Me.Text1(4).Text = "", "00", Me.Text1(4).Text) & "' "
      StrSQLa = StrSQLa & " union Select SP05,SP06,SP07,SP01,SP02,SP03,SP04,'','' From SERVICEPRACTICE Where SP01='" & Me.Text1(1).Text & "' AND SP02='" & Me.Text1(2).Text & "' AND SP03='" & IIf(Me.Text1(3).Text = "", "0", Me.Text1(3).Text) & "' AND SP04='" & IIf(Me.Text1(4).Text = "", "00", Me.Text1(4).Text) & "' "
   End If
   'Add By Sindy 2010/10/13
   If Text1(5) <> "" Then
      If strSql <> "" Then strSql = strSql & " and "
      strSql = "TM15='" & Text1(5) & "'"
   End If
   If Text1(8) <> "" Then
      If strSql <> "" Then strSql = strSql & " and "
      strSql = "TM12='" & Text1(8) & "'"
   End If
   If Text1(5) <> "" Or Text1(8) <> "" Then
      StrSQLa = "Select TM05,TM06,TM07,TM01,TM02,TM03,TM04,TM12,TM15 From TRADEMARK Where " & strSql
   End If
   '2010/10/13 End
   
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount <= 0 Then
      MsgBox "資料庫無此案號資料!!!", vbExclamation + vbOKOnly
      Me.lblCaseName(0).Caption = "(中)："
      Me.lblCaseName(1).Caption = "(英)："
      Me.lblCaseName(2).Caption = "(日)："
      Me.Combo1.Clear
      Me.lblSaleZone.Caption = ""
      Me.lblSaleName.Caption = ""
      m_blnTxtValidate = False
   Else
      Me.lblCaseName(0).Caption = "(中)：" & rsA.Fields(0).Value
      Me.lblCaseName(1).Caption = "(英)：" & rsA.Fields(1).Value
      Me.lblCaseName(2).Caption = "(日)：" & rsA.Fields(2).Value
      'Add By Sindy 2010/10/13
      If Text1(5) <> "" Or Text1(8) <> "" Then
         Text1(1).Text = "" & rsA.Fields(3).Value
         'Add By Sindy 2010/10/29
         Cancel = False
         Call Text1_Validate(1, Cancel)
         If Cancel = True Then
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            Exit Sub
         End If
         '2010/10/29 End
         Text1(2).Text = "" & rsA.Fields(4).Value
         Text1(3).Text = "" & rsA.Fields(5).Value
         Text1(4).Text = "" & rsA.Fields(6).Value
         Text1(5).Text = "" & rsA.Fields(8).Value
         Text1(8).Text = "" & rsA.Fields(7).Value
      End If
      '2010/10/13 End
      'Added by Lydia 2022/08/12
      For Each tmpObj In Text1
          tmpObj.Tag = tmpObj.Text
      Next
      'end 2022/08/12
      QueryData
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Select

If Cancel Then TextInverse Text1(Index)
End Sub

Private Sub QueryData()
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim bolTmp As Boolean
Dim strTempName As String, strApp As String, ii As Integer
Dim strSalesNo As String 'Add By Sindy 2011/5/20
Dim Cancel As Boolean 'Add By Sindy 2011/8/8
Dim strST04 As String 'ADD BY SONIA 2014/5/30 在職, 離職
Dim strCP12 As String 'ADD BY SONIA 2014/5/30 收文目前業務區
   
   m_strToCCNo = ""
   Me.Combo1.Clear
   Me.lblSaleZone.Caption = ""
   Me.lblSaleName.Caption = ""
   'Add By Sindy 2010/10/13
   Me.cmdOK(0).Enabled = False
   Me.cmdOK(2).Enabled = False
   Me.Text1(0).Enabled = False
'cancel by sonia 2016/10/19 移到下來去
'   If Left(Trim(Pub_StrUserSt03), 2) = "F1" Then
'      Text2 = "3"
'      'Add By Sindy 2011/8/8
'      Cancel = False
'      Call Text2_Validate(Cancel)
'      '2011/8/8 End
'   End If
'end 2016/10/19

   '2010/10/13 End
'2009/12/21 modify by sonia 智權人員改抓依各系統類別規則,此處只抓所有A類接洽單之在職承辦人
'   '若不為"FCP"及"FG"案時
'   If Me.Text1(1).Text <> "FCP" And Me.Text1(1).Text <> "FG" Then
'       'Modify By Cheng 2004/02/03
'       '只抓在職人員
'   '   strSQLA = "SELECT DISTINCT NP10,'1' FROM NEXTPROGRESS WHERE " & ChgNextProgress(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4).Text) & " AND NP06 IS NULL "
'   '   strSQLA = strSQLA & "UNION SELECT DISTINCT CP13,'1' FROM CASEPROGRESS WHERE " & ChgCaseprogress(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4).Text) & " AND CP09 < 'B' "
'   '    'Add By Cheng 2003/04/18
'   '   strSQLA = strSQLA & "UNION SELECT DISTINCT CP14,'2' FROM CASEPROGRESS WHERE " & ChgCaseprogress(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4).Text) & " AND CP09 < 'B' "
'       StrSQLa = "SELECT DISTINCT NP10,'1' FROM NEXTPROGRESS, Staff WHERE NP10=ST01(+) And " & ChgNextProgress(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4).Text) & " AND NP06 IS NULL And ST04='1' "
'       StrSQLa = StrSQLa & "UNION SELECT DISTINCT CP13,'1' FROM CASEPROGRESS, Staff WHERE CP13=ST01(+) And " & ChgCaseprogress(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4).Text) & " AND CP09 < 'B' And ST04='1' "
'       StrSQLa = StrSQLa & "UNION SELECT DISTINCT CP14,'2' FROM CASEPROGRESS, Staff WHERE CP14=ST01(+) And " & ChgCaseprogress(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4).Text) & " AND CP09 < 'B' And ST04='1' "
'       'End
'      StrSQLa = StrSQLa & " ORDER BY 2"
'   '若為"FCP"或"FG"案時
'   Else
'       'Modify By Cheng 2004/02/03
'       '只抓在職人員
'   '   strSQLA = "SELECT DISTINCT NA51 FROM PATENT,FAGENT,NATION WHERE SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA10=NA01(+) AND " & ChgPatent(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4)) & " AND PA75 IS NOT NULL AND NA51 IS NOT NULL "
'   '   strSQLA = strSQLA & " UNION SELECT DISTINCT NA51 FROM PATENT,CUSTOMER,NATION WHERE SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND CU10=NA01(+) AND " & ChgPatent(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4)) & " AND PA75 IS NULL AND PA26 IS NOT NULL AND NA51 IS NOT NULL "
'   '   strSQLA = strSQLA & " UNION SELECT DISTINCT NA51 FROM TRADEMARK,FAGENT,NATION WHERE SUBSTR(TM44,1,8)=FA01(+) AND SUBSTR(TM44,9,1)=FA02(+) AND FA10=NA01(+) AND " & ChgTradeMark(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4)) & " AND TM44 IS NOT NULL AND NA51 IS NOT NULL "
'   '   strSQLA = strSQLA & " UNION SELECT DISTINCT NA51 FROM TRADEMARK,CUSTOMER,NATION WHERE SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) AND CU10=NA01(+) AND " & ChgTradeMark(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4)) & " AND TM44 IS NULL AND TM23 IS NOT NULL AND NA51 IS NOT NULL "
'   '   strSQLA = strSQLA & " UNION SELECT DISTINCT NA51 FROM LAWCASE,FAGENT,NATION WHERE SUBSTR(LC22,1,8)=FA01(+) AND SUBSTR(LC22,9,1)=FA02(+) AND FA10=NA01(+) AND " & ChgLawcase(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4)) & " AND LC22 IS NOT NULL AND NA51 IS NOT NULL "
'   '   strSQLA = strSQLA & " UNION SELECT DISTINCT NA51 FROM LAWCASE,CUSTOMER,NATION WHERE SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND CU10=NA01(+) AND " & ChgLawcase(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4)) & " AND LC22 IS NULL AND LC11 IS NOT NULL AND NA51 IS NOT NULL "
'   '   strSQLA = strSQLA & " UNION SELECT DISTINCT NA51 FROM HIRECASE,CUSTOMER,NATION WHERE SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) AND CU10=NA01(+) AND " & ChgHirecase(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4)) & " AND HC05 IS NOT NULL AND NA51 IS NOT NULL "
'   '   strSQLA = strSQLA & " UNION SELECT DISTINCT NA51 FROM SERVICEPRACTICE,FAGENT,NATION WHERE SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND FA10=NA01(+) AND " & ChgService(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4)) & " AND SP26 IS NOT NULL AND NA51 IS NOT NULL "
'   '   strSQLA = strSQLA & " UNION SELECT DISTINCT NA51 FROM SERVICEPRACTICE,CUSTOMER,NATION WHERE SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND CU10=NA01(+) AND " & ChgService(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4)) & " AND SP26 IS NULL AND SP08 IS NOT NULL AND NA51 IS NOT NULL "
'      StrSQLa = "SELECT DISTINCT NA51 FROM PATENT,FAGENT,NATION, Staff WHERE SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA10=NA01(+) AND NA51=ST01(+) And " & ChgPatent(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4)) & " AND PA75 IS NOT NULL AND NA51 IS NOT NULL And ST04='1' "
'      StrSQLa = StrSQLa & " UNION SELECT DISTINCT NA51 FROM PATENT,CUSTOMER,NATION, Staff WHERE SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND CU10=NA01(+) AND NA51=ST01(+) And " & ChgPatent(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4)) & " AND PA75 IS NULL AND PA26 IS NOT NULL AND NA51 IS NOT NULL And ST04='1' "
'      StrSQLa = StrSQLa & " UNION SELECT DISTINCT NA51 FROM TRADEMARK,FAGENT,NATION, Staff WHERE SUBSTR(TM44,1,8)=FA01(+) AND SUBSTR(TM44,9,1)=FA02(+) AND FA10=NA01(+) AND NA51=ST01(+) And " & ChgTradeMark(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4)) & " AND TM44 IS NOT NULL AND NA51 IS NOT NULL And ST04='1' "
'      StrSQLa = StrSQLa & " UNION SELECT DISTINCT NA51 FROM TRADEMARK,CUSTOMER,NATION, Staff WHERE SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) AND CU10=NA01(+) AND NA51=ST01(+) And " & ChgTradeMark(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4)) & " AND TM44 IS NULL AND TM23 IS NOT NULL AND NA51 IS NOT NULL And ST04='1' "
'      StrSQLa = StrSQLa & " UNION SELECT DISTINCT NA51 FROM LAWCASE,FAGENT,NATION, Staff WHERE SUBSTR(LC22,1,8)=FA01(+) AND SUBSTR(LC22,9,1)=FA02(+) AND FA10=NA01(+) AND NA51=ST01(+) And " & ChgLawcase(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4)) & " AND LC22 IS NOT NULL AND NA51 IS NOT NULL And ST04='1' "
'      StrSQLa = StrSQLa & " UNION SELECT DISTINCT NA51 FROM LAWCASE,CUSTOMER,NATION, Staff WHERE SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND CU10=NA01(+) AND NA51=ST01(+) And " & ChgLawcase(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4)) & " AND LC22 IS NULL AND LC11 IS NOT NULL AND NA51 IS NOT NULL And ST04='1' "
'      StrSQLa = StrSQLa & " UNION SELECT DISTINCT NA51 FROM HIRECASE,CUSTOMER,NATION, Staff WHERE SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) AND CU10=NA01(+) AND NA51=ST01(+) And " & ChgHirecase(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4)) & " AND HC05 IS NOT NULL AND NA51 IS NOT NULL And ST04='1' "
'      StrSQLa = StrSQLa & " UNION SELECT DISTINCT NA51 FROM SERVICEPRACTICE,FAGENT,NATION, Staff WHERE SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND FA10=NA01(+) AND NA51=ST01(+) And " & ChgService(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4)) & " AND SP26 IS NOT NULL AND NA51 IS NOT NULL And ST04='1' "
'      StrSQLa = StrSQLa & " UNION SELECT DISTINCT NA51 FROM SERVICEPRACTICE,CUSTOMER,NATION, Staff WHERE SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND CU10=NA01(+) AND NA51=ST01(+) And " & ChgService(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4)) & " AND SP26 IS NULL AND SP08 IS NOT NULL AND NA51 IS NOT NULL And ST04='1' "
'       'End
'      StrSQLa = StrSQLa & " ORDER BY 1 "
'   End If
   
   'Modify By Sindy 2011/5/20 針對F4103抓原直屬主管
   strSalesNo = ""
   Select Case Me.Text1(1).Text
      Case "FCP", "FG"
         strSalesNo = PUB_GetFCPSalesNo(Me.Text1(1).Text, Me.Text1(2).Text, Left(Me.Text1(3).Text & "0", 1), Left(Me.Text1(4).Text & "00", 2))
      Case "FCT"
         'Modified by Lydia 2019/06/10 FCT案件之收受者改用原規則，抓案件最後收文之智權人員(傳1101案件性質才能走原規則)
         'strSalesNo = PUB_GetFCTSalesNo(Me.Text1(1).Text, Me.Text1(2).Text, Left(Me.Text1(3).Text & "0", 1), Left(Me.Text1(4).Text & "00", 2))
         strSalesNo = PUB_GetFCTSalesNo(Me.Text1(1).Text, Me.Text1(2).Text, Left(Me.Text1(3).Text & "0", 1), Left(Me.Text1(4).Text & "00", 2), "1101")
      Case "FCL", "LIN"
         strSalesNo = PUB_GetFCLSalesNo(Me.Text1(1).Text, Me.Text1(2).Text, Left(Me.Text1(3).Text & "0", 1), Left(Me.Text1(4).Text & "00", 2))
      Case Else
         'MODIFY BY SONIA 2014/5/30 改抓最後收文智權人員
         'strSalesNo = PUB_GetAKindSalesNo(Me.Text1(1).Text, Me.Text1(2).Text, Left(Me.Text1(3).Text & "0", 1), Left(Me.Text1(4).Text & "00", 2))
         strSalesNo = "": strST04 = "": strCP12 = ""
         'modify by sonia 2025/7/31 改為抓進度檔只抓未發文者故加入CP158=0，CP57 Is Null改CP159=0
         StrSQLa = "Select * From CaseProgress, Staff Where CP13=ST01(+) And CP01='" & Me.Text1(1).Text & "' And CP02='" & Me.Text1(2).Text & "' And CP03='" & Left(Me.Text1(3).Text & "0", 1) & "' And CP04='" & Left(Me.Text1(4).Text & "00", 2) & "' And CP05 Is Not Null And CP09 <'B' And CP158=0 and cp159=0 Order By CP05 Desc, CP09 Desc "
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            strSalesNo = "" & rsA("CP13").Value
            strST04 = "" & rsA("ST04").Value
            strCP12 = "" & rsA("CP12").Value
         Else
            '沒有A類智權人員改抓B類智權人員
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            'modify by sonia 2025/7/31 改為抓進度檔只抓未發文者故加入CP158=0，CP57 Is Null改CP159=0
            StrSQLa = "Select * From CaseProgress, Staff Where CP13=ST01(+) And CP01='" & Me.Text1(1).Text & "' And CP02='" & Me.Text1(2).Text & "' And CP03='" & Left(Me.Text1(3).Text & "0", 1) & "' And CP04='" & Left(Me.Text1(4).Text & "00", 2) & "' And CP05 Is Not Null And CP09 <'C' And CP158=0 and cp159=0 Order By CP05 Desc, CP09 Desc "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               strSalesNo = "" & rsA("CP13").Value
               strST04 = "" & rsA("ST04").Value
               strCP12 = "" & rsA("CP12").Value
            'cancel by sonia 2025/7/31
            ''A及B類都已取消收文則抓最後一道進度
            'Else
            '   If rsA.State <> adStateClosed Then rsA.Close
            '   Set rsA = Nothing
            '   StrSQLa = "Select * From CaseProgress, Staff Where CP13=ST01(+) And CP01='" & Me.Text1(1).Text & "' And CP02='" & Me.Text1(2).Text & "' And CP03='" & Left(Me.Text1(3).Text & "0", 1) & "' And CP04='" & Left(Me.Text1(4).Text & "00", 2) & "' And CP05 Is Not Null And CP09 <'C' Order By CP05 Desc, CP09 Desc "
            '   rsA.CursorLocation = adUseClient
            '   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            '   If rsA.RecordCount > 0 Then
            '      strSalesNo = "" & rsA("CP13").Value
            '      strST04 = "" & rsA("ST04").Value
            '      strCP12 = "" & rsA("CP12").Value
            '   End If
            End If
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         '若已離職
         'modify by sonia 2025/7/31 改為抓進度檔只抓未發文者故可能抓不到智權人員
         'If strST04 <> "1" Then
         If strST04 <> "1" Or strSalesNo = "" Then
            '2015/2/25 modify by sonia 若離職再改為原通用抓法 CFT-010077,否則會與進度檔不符
            ''抓部門檔的主管
            'StrSQLa = "Select * From acc090 Where a0901='" & strCP12 & "'  "
            'rsA.CursorLocation = adUseClient
            'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            'If rsA.RecordCount > 0 Then
            '   strSalesNo = "" & rsA("A0909").Value
            'End If
            'If rsA.State <> adStateClosed Then rsA.Close
            'Set rsA = Nothing
            strSalesNo = PUB_GetAKindSalesNo(Me.Text1(1).Text, Me.Text1(2).Text, Left(Me.Text1(3).Text & "0", 1), Left(Me.Text1(4).Text & "00", 2))
            '2015/2/25 END
         End If
         'END 2014/5/30
   End Select
   
   If strSalesNo = "F4103" Then
      '日本地區-葉易雲主任 78011
      '其他地區-洪琬姿副理 80030
      StrSQLa = "select tm01 from trademark where tm01='" & Me.Text1(1).Text & "' and tm02='" & Me.Text1(2).Text & "'" & _
                " AND ((TM44 is not null AND exists (select * from fagent where fa01=substr(tm44,1,8) and fa02=substr(tm44,9) and substr(fa10,1,3)='011'))" & _
                " or (TM44 is null AND exists (select * from customer where cu01=substr(tm23,1,8) and cu02=substr(tm23,9) and substr(cu10,1,3)='011')))" & _
                " union select sp01 from SERVICEPRACTICE where sp01='" & Me.Text1(1).Text & "' and sp02='" & Me.Text1(2).Text & "'" & _
                " AND ((sp26 is not null AND exists (select * from fagent where fa01=substr(sp26,1,8) and fa02=substr(sp26,9) and substr(fa10,1,3)='011'))" & _
                " or (sp26 is null AND exists (select * from customer where cu01=substr(sp08,1,8) and cu02=substr(sp08,9) and substr(cu10,1,3)='011')))"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount <= 0 Then
         Me.Combo1.AddItem "80030"
      Else
         Me.Combo1.AddItem "78011"
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   ElseIf Trim(strSalesNo) <> "" Then
      Me.Combo1.AddItem strSalesNo
   End If
   '2011/5/20 End
   
   StrSQLa = "SELECT DISTINCT CP14,'2' FROM CASEPROGRESS, Staff WHERE CP14=ST01(+) And " & ChgCaseprogress(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4).Text) & " AND CP09 < 'B' And ST04='1' "
   StrSQLa = StrSQLa & " ORDER BY 2"
'2009/12/21 end
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount <= 0 Then
       'Modify By Cheng 2003/04/18
   '   MsgBox "資料庫無智權人員資料!!!", vbExclamation + vbOKOnly
'2009/12/21 CANCEL BY SONIA
'      MsgBox "本案無在職智權人員或承辦人資料!!!", vbExclamation + vbOKOnly
   '   Me.Combo1.Clear
   '   Me.lblSaleZone.Caption = ""
   '   Me.lblSaleName.Caption = ""
   '   If rsA.State <> adStateClosed Then rsA.Close
   '   Set rsA = Nothing
   '   Exit Sub
   Else
      While Not rsA.EOF
         Me.Combo1.AddItem "" & rsA.Fields(0).Value
         rsA.MoveNext
      Wend
'2010/4/19 CANCEL BY SONIA 移到下面,否則CFP-19571承辦離職則不會預設
'      Me.Combo1.ListIndex = 0
'      Combo1_Change
'2010/4/19 END
   End If
   'add by sonia 2016/10/19 從上面移下來
   If Left(Trim(Pub_StrUserSt03), 2) = "F1" Then
      Text2 = "3"
      'Add By Sindy 2011/8/8
      Cancel = False
      Call Text2_Validate(Cancel)
      '2011/8/8 End
   End If
   'end 2016/10/19
   '2010/4/19 ADD BY SONIA 自IF條件中拉出來
   Me.Combo1.ListIndex = 0
   Combo1_Change
   '2010/4/49 END
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
   'Add By Sindy 2010/10/13
   '直屬主管
   'MODIFY BY SONIA 2014/6/3 陳經理要求副本要顯示部門名稱
   strSql = "SELECT st52 FROM staff WHERE st01='" & Combo1.Text & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      m_strToCCNo = "" & RsTemp.Fields("st52")
   End If
   '副本：全部員工
   '2013/9/30 modify by sonia 外商陳經理說依部門+員工編號排序,故排序加入st03條件
   'MODIFY BY SONIA 2014/6/3 陳經理要求要顯示部門名稱
   'strSql = "SELECT st01,st02 FROM staff,SalaryData WHERE st04='1' and st01>'6' and st01<'F' and st01=sd01 order by st03,st01 asc"
   'Added by Lydia 2023/12/25
   If strSrvDate(1) >= 新部門啟用日 Then
       strSql = "SELECT NVL(A0923,A0902) AS a0902,st01,st02 FROM staff,acc090,ACC090NEW WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) AND ST93=A0921(+) and substr(st01,4,1)<>'9' order by nvl(st93,st03),st01 asc"
   Else
   'end 2023/12/25
       strSql = "SELECT a0902,st01,st02 FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9' order by st03,st01 asc"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         RsTemp.MoveFirst
         Do While RsTemp.EOF = False
            'MODIFY BY SONIA 2014/6/3 陳經理要求要顯示部門名稱
            'lstMailCC.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
            lstMailCC.AddItem Trim(RsTemp.Fields("a0902")) & " " & Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
            RsTemp.MoveNext
         Loop
      End With
   End If
   If Text2 = "3" And m_strToCCNo <> "" Then '3:FCT審查員電話通知
      For ii = 0 To lstMailCC.ListCount - 1
         'MODIFY BY SONIA 2014/6/3 因陳經理要求要顯示部門名稱,故截取員工編號要改方式
         'If Left(Trim(lstMailCC.List(ii)), 5) = m_strToCCNo Then lstMailCC.Selected(ii) = True
         s_MailCC = Trim(Mid(lstMailCC.List(ii), InStr(lstMailCC.List(ii), " ") + 1, 5))
         If s_MailCC = m_strToCCNo Then lstMailCC.Selected(ii) = True
         'END 2014/6/3
      Next
   End If
   Me.cmdOK(0).Enabled = True
   Me.cmdOK(2).Enabled = True
   Me.Text1(0).Enabled = True
   If GetStaffDepartment(strUserNum) = "P12" Then
      Text1(0) = "2"
   Else
      Text1(0) = "1"
   End If
   '2010/10/13 End
   
   'Modify by Morgan 2005/9/4加申請人2~5
   'Modify By Sindy 2009/08/14 加CU16,CU18
   'Modify By Sindy 2010/10/13 將此段程式移進來查詢時執行
   If m_rsA.State <> adStateClosed Then m_rsA.Close
   Set m_rsA = Nothing
   'Modify By Sindy 2011/2/18 增加LC43,LC44,LC45,LC46,HC24,HC25,HC26,HC27,SP65,SP66
   m_strSQLA = "select nvl(CU04,decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CU04,nvl(NA03,NA04) as NA03,PA27,PA28,PA29,PA30,CU16,CU18 FROM PATENT,CUSTOMER,NATION " & _
      "WHERE PA01='" & Me.Text1(1).Text & "' AND PA02='" & Me.Text1(2).Text & "' AND PA03='" & Left(Me.Text1(3).Text & "0", 1) & "' AND PA04='" & Left(Me.Text1(4).Text & "00", 2) & "' " & _
      "AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND PA09=NA01(+) " & _
      " UNION select nvl(CU04,decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CU04,nvl(NA03,NA04) as NA03,TM78,TM79,TM80,TM81,CU16,CU18 FROM TRADEMARK,CUSTOMER,NATION " & _
      "WHERE TM01='" & Me.Text1(1).Text & "' AND TM02='" & Me.Text1(2).Text & "' AND TM03='" & Left(Me.Text1(3).Text & "0", 1) & "' AND TM04='" & Left(Me.Text1(4).Text & "00", 2) & "' " & _
      "AND SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) AND TM10=NA01(+) " & _
      " UNION select nvl(CU04,decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CU04,nvl(NA03,NA04) as NA03,LC43,LC44,LC45,LC46,CU16,CU18 FROM LAWCASE,CUSTOMER,NATION " & _
      "WHERE LC01='" & Me.Text1(1).Text & "' AND LC02='" & Me.Text1(2).Text & "' AND LC03='" & Left(Me.Text1(3).Text & "0", 1) & "' AND LC04='" & Left(Me.Text1(4).Text & "00", 2) & "' " & _
      "AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND LC15=NA01(+) " & _
      " UNION select nvl(CU04,decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CU04,nvl(NA03,NA04) as NA03,HC24,HC25,HC26,HC27,CU16,CU18 FROM HIRECASE,CUSTOMER,NATION " & _
      "WHERE HC01='" & Me.Text1(1).Text & "' AND HC02='" & Me.Text1(2).Text & "' AND HC03='" & Left(Me.Text1(3).Text & "0", 1) & "' AND HC04='" & Left(Me.Text1(4).Text & "00", 2) & "' " & _
      "AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) AND '000' = NA01(+) " & _
      " UNION select nvl(CU04,decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CU04,nvl(NA03,NA04) as NA03,SP58,SP59,SP65,SP66,CU16,CU18 FROM SERVICEPRACTICE,CUSTOMER,NATION " & _
      "WHERE SP01='" & Me.Text1(1).Text & "' AND SP02='" & Me.Text1(2).Text & "' AND SP03='" & Left(Me.Text1(3).Text & "0", 1) & "' AND SP04='" & Left(Me.Text1(4).Text & "00", 2) & "' " & _
      "AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND SP09=NA03(+) "
   m_rsA.CursorLocation = adUseClient
   m_rsA.Open m_strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
   If m_rsA.RecordCount > 0 Then
      strApp = m_rsA.Fields(0).Value
      For ii = 2 To 5
         If Not IsNull(m_rsA.Fields(ii).Value) Then
           strApp = strApp & "、" & GetCustomerName(m_rsA.Fields(ii).Value)
         End If
      Next ii
      LabTM23 = strApp
   End If
End Sub

'2008/8/28 add by sonia 加種類
Private Sub Text2_GotFocus()
   TextInverse Me.Text2
   
   'Added by Morgan 2021/6/15
   If Text1(2) <> "" And Text1(4) = "" Then
      Text1_Validate 4, False
   End If
   'end 2021/6/15
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'modify by sonia 2016/10/18 加5文件公簽證選項
   If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
   
   'Add By Sindy 2020/1/16
   If Text1(1) = "CFP" Then
      Check1(0).Value = 0
      Check1(1).Value = 1 '簽署文件
      Check1(2).Value = 0
      Check1(3).Value = 0
      Check1(4).Value = 0
   ElseIf Text1(1) = "P" Then
      Check1(0).Value = 1 '委任書
      Check1(1).Value = 0
      Check1(2).Value = 0
      Check1(3).Value = 0
      Check1(4).Value = 0
   End If
   '2020/1/16 END
End Sub

Private Sub Text2_LostFocus()
   If m_blnTxtValidate = False Then
      Me.Text1(1).SetFocus
      m_blnTxtValidate = True
   End If
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
Dim ii As Integer
Dim strTmp As String, strTmp1 As String 'add by sonia 2016/10/18
Dim strText As String 'Add By Sindy 2020/1/16
   
   HideCol 'Add By Sindy 2010/10/13
   
   'add by sonia 2016/10/18
   cmdOK(2).Enabled = True
   If Me.Text2.Text <> "5" And Combo1.ListCount > 0 Then 'Modified by Morgan 2021/6/15 判斷有受文者才可預設,否則會有380錯誤
      'Add By Sindy 2021/1/28 玲玲說空白才要預設
      If Trim(Combo1.Text) = "" Then
      '2021/1/28 END
         Me.Combo1.ListIndex = 0
         Combo1_Change
      End If
   End If
   'end 2016/10/18

   If Me.Text2.Text = "" Then
      Label1(13).Caption = "主旨："
      Text4.Text = ""
      Exit Sub
   End If
   
   cmdInput.Visible = False 'Added by Lydia 2022/08/12
   
   If Me.Text2.Text = "1" Then
      Me.Text4.Text = ""
      Label1(13).Caption = "主旨：聯絡單"
   ElseIf Me.Text2.Text = "2" Then
      Text4.Text = "　　上述之申請人地址有誤，以致於 XX通知函遭退件，煩請寫異動表更正之。"
      Label1(13).Caption = "主旨：信件退回聯絡單"
   'Add By Sindy 2010/10/13
   ElseIf Me.Text2.Text = "3" Then '3:FCT審查員電話通知
      
      Me.Text4.Text = ""
      'Modified by Lydia 2022/08/12 FCT審查機關來電與回覆流程管制
      'Label1(13).Caption = "主旨：請回電審查員"
      Label1(13).Caption = "主旨：OUR REF: " & Text1(1) & "-" & Text1(2) & IIf(Text1(3) <> "0", "-" & Text1(3), "") & IIf(Text1(4) <> "00", "-" & Text1(4), "") & _
                                          " 電話通知 [INCOM.1727]" '與email一致
      'cmdInput.Visible = True '原本設計多案案號輸入，後來阿蓮說不要，所以隱藏
      'end 2022/08/12
      '審查員
      Label1(12).Visible = True
      txtReviewer.Visible = True
      Option1(0).Visible = True
      Option1(1).Visible = True
      '分機
      Label1(11).Visible = True
      Text1(7).Visible = True
      'Add By Sindy 2010/11/19
      If Text2 = "3" And m_strToCCNo <> "" Then
         For ii = 0 To lstMailCC.ListCount - 1
            'MODIFY BY SONIA 2014/6/3 因陳經理要求要顯示部門名稱,故截取員工編號要改方式
            'If Left(Trim(lstMailCC.List(ii)), 5) = m_strToCCNo Then lstMailCC.Selected(ii) = True
            s_MailCC = Trim(Mid(lstMailCC.List(ii), InStr(lstMailCC.List(ii), " ") + 1, 5))
            If s_MailCC = m_strToCCNo Then lstMailCC.Selected(ii) = True
            'END 2014/6/3
         Next
      End If
      '2010/11/19 End
   ElseIf Me.Text2.Text = "4" Then
      'modify by sonia 2014/10/2 P案也可以用-陳玲玲
      'Text4.Text = "　　此案件缺簽署文件，請儘速提供，謝謝！"
      'Label1(13).Caption = "主旨：CFP案件缺文件聯絡單"
      
      'Modify By Sindy 2020/1/16
'      Select Case Text1(1)
'         Case "CFP"
'            Text4.Text = "　　此案件缺簽署文件，請儘速提供，謝謝！"
'         Case "P"
'            Text4.Text = "　　此案件缺委任書，請儘速提供，謝謝！"
'      End Select
      If Check1(0).Value = 1 Then strText = IIf(strText <> "", strText & "、", "") & Check1(0).Caption
      If Check1(1).Value = 1 Then strText = IIf(strText <> "", strText & "、", "") & Check1(1).Caption
      If Check1(2).Value = 1 Then strText = IIf(strText <> "", strText & "、", "") & Check1(2).Caption
      If Check1(3).Value = 1 Then strText = IIf(strText <> "", strText & "、", "") & Check1(3).Caption
      If Check1(4).Value = 1 Then strText = IIf(strText <> "", strText & "、", "") & Check1(4).Caption
      Text4.Text = "　　此案件缺" & strText & "，請儘速提供，謝謝！"
      '2020/1/16 END
      'Modified by Morgan 2020/8/11 顯示完整案號--蕭茹曣
      'Label1(13).Caption = "主旨：" & Text1(1) & " 案件缺文件聯絡單"
      Label1(13).Caption = "主旨：" & Text1(1) & "-" + Text1(2) & "-" + Text1(3) & "-" + Text1(4) & " 案件缺文件聯絡單"
      'end 2020/8/11
      'end 2014/10/2
   '2010/10/13 End
   'add by sonia 2016/10/18 加5文件公簽證選項
   ElseIf Me.Text2.Text = "5" Then
      cmdOK(2).Enabled = False
      Combo1.Text = "77047" '預設謝碩儒
      Combo1_Change
      Call ClsPDGetStaff(strUserNum, strTmp, strTmp1)   '取得操作人部門名稱+姓名
      Label1(13).Caption = "主旨：" & Replace(Me.Text1(1).Text & "-" & Me.Text1(2).Text & "-" & Left(Me.Text1(3).Text & "0", 1) & "-" & Left(Me.Text1(4).Text & "00", 2), "-0-00", "") & " 文件公簽證"
      'modify by sonia 2017/6/19 禧佩說收據公司加專利法律選項
      'Modified by Lydia 2022/11/18 禧佩: 改為智慧所和智權公司
      'Text4 = "日期： " & ChangeWStringToWDateString(strSrvDate(1)) & vbCrLf & _
              "正本份數：  _____ 份" & vbCrLf & _
              "□ 公證 □外交部 □ _______________ 辦事處" & vbCrLf & vbCrLf & _
              "□ 其他_____________________" & vbCrLf & _
              "收據公司： □專利商標 □專利法律 □智權公司" & vbCrLf & _
              "承辦人： " & strTmp1 & "　" & strTmp
      Text4 = "日期： " & ChangeWStringToWDateString(strSrvDate(1)) & vbCrLf & _
              "正本份數：  _____ 份" & vbCrLf & _
              "□ 公證 □外交部 □ _______________ 辦事處" & vbCrLf & vbCrLf & _
              "□ 其他_____________________" & vbCrLf & _
              "收據公司： □智慧所　 □智權公司" & vbCrLf & _
              "承辦人： " & strTmp1 & "　" & strTmp
   'end 2016/10/18
   End If
End Sub

'2008/8/28 end
Private Sub Text4_GotFocus()
   TextInverse Me.Text4
   'Modifie by Lydia 2021/06/17
   'OpenIme
   CloseIme
End Sub

'Add By Cheng 2003/04/03
Private Sub InitPrtPosition(dblTop As Double, dblLeft As Double)
    m_dblTop = dblTop
    m_dblLeft = dblLeft
    m_dblTitleHeight = 0
    m_dblLine = 0
    m_dblLineHeight = 1
    m_dblBetweenLine = 0.2
    m_dblLineHeight1 = 0.6
    m_dblBetweenLine1 = 0.1
End Sub

'Add By Cheng 2003/04/03
Private Sub PrintContactSheet(strDept As String)
Dim dblPrtX As Double
Dim dblPrtY As Double
Dim ii As Integer
Dim jj As Integer
Dim strTxt  As String
Dim intTxtLeng As Integer
Dim arrLineString 'Add by Morgan 2004/11/2
    
    Printer.Font.Name = "標楷體"
    Printer.Font.Size = 16
    'Mark by Amy 2020/03/27 取消印公司名稱
'    dblPrtX = m_dblLeft + (19 - Printer.TextWidth("台一國際專利商標事務所")) / 2
'    dblPrtY = m_dblTop + m_dblBetweenLine + 0
'    Printer.CurrentX = dblPrtX
'    Printer.CurrentY = dblPrtY
'    Printer.Print "台一國際專利商標事務所"
    dblPrtX = m_dblLeft + (19 - Printer.TextWidth("簡易聯絡單")) / 2
    dblPrtY = m_dblTop + m_dblBetweenLine + 1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "簡易聯絡單"
        
    m_dblTitleHeight = 2.2
    
    m_dblLine = 0
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 19, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight + 9 * m_dblLineHeight)
    Printer.Line (m_dblLeft + 4.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 4.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight + 3 * m_dblLineHeight)
    Printer.Line (m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight + 9 * m_dblLineHeight)
    Printer.Line (m_dblLeft + 19, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 19, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight + 9 * m_dblLineHeight)
    Printer.Font.Size = 14
    dblPrtX = m_dblLeft + (4.5 - Printer.TextWidth("受文者")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "受文者"
    dblPrtX = m_dblLeft + 4.5 + (4 - Printer.TextWidth("發文者")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "發文者"
    
    m_dblLine = m_dblLine + 1
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)
    '受文者部門
    dblPrtX = m_dblLeft + m_dblBetweenLine
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "" & Me.lblSaleZone.Caption
    '受文者
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + (4.5 - Printer.TextWidth(Me.lblSaleName.Caption)) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight - 0.3 * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "" & Me.lblSaleName.Caption
    '發文者
    dblPrtX = m_dblLeft + 4.5 + (4 - Printer.TextWidth(GetStaffName(strUserNum, True))) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight - (m_dblLineHeight / 2)
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print GetStaffName(strUserNum, True)
    
    m_dblLine = m_dblLine + 1
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)
    Printer.Line (m_dblLeft + 2.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 2.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight + 6 * m_dblLineHeight)
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("發文時間")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "發文時間"
    dblPrtX = m_dblLeft + 2.5 + (6 - Printer.TextWidth(Mid(ServerDate, 1, 4) & "年" & Mid(ServerDate, 5, 2) & "月" & Mid(ServerDate, 7, 2) & "日   " & Mid(Right("000000" & ServerTime, 6), 1, 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2))) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    'Modify By Cheng 2003/04/18
'    Printer.Print Mid(ServerDate, 1, 4) & "年" & Mid(ServerDate, 5, 2) & "月" & Mid(ServerDate, 7, 2) & "日   " & Mid(ServerTime, 1, 2) & ":" & Mid(ServerTime, 3, 2)
    Printer.Print Mid(ServerDate, 1, 4) & "年" & Mid(ServerDate, 5, 2) & "月" & Mid(ServerDate, 7, 2) & "日   " & Mid(Right("000000" & ServerTime, 6), 1, 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)
    
    m_dblLine = m_dblLine + 1
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("答覆")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "答覆"
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("□否 □要")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "□否 □要"
    dblPrtX = m_dblLeft + 2.5 + (6 - Printer.TextWidth("用□電話 □口頭  回覆")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight - (m_dblLineHeight / 2)
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "用□電話 □口頭  回覆"
    
    m_dblLine = m_dblLine + 1
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("　限　　")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "　限　　"
    dblPrtX = m_dblLeft + 2.5 + (6 - Printer.TextWidth("            AM")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "            AM"
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("時　要求")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight + 0.5 * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "時　要求"
    dblPrtX = m_dblLeft + 2.5 + (6 - Printer.TextWidth("    月    日    ")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight + 0.5 * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "    月    日    "
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("　間　　")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "　間　　"
    dblPrtX = m_dblLeft + 2.5 + (6 - Printer.TextWidth("            PM")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "            PM"
    
    m_dblLine = m_dblLine + 1
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 8.5, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)
    dblPrtX = m_dblLeft + (2.5 - Printer.TextWidth("發文地點")) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "發文地點"
    dblPrtX = m_dblLeft + 2.5 + (6 - Printer.TextWidth(strDept)) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine + m_dblLine * m_dblLineHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strDept
    
    m_dblLine = m_dblLine + 1
    Printer.Line (m_dblLeft, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)-(m_dblLeft + 19, m_dblTop + m_dblTitleHeight + m_dblLine * m_dblLineHeight)

    Printer.Font.Size = 13
    m_dblLine = 0
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    'modify by sonia 2016/10/18
    'Printer.Print "本所案號：" & Replace(Me.Text1(1).Text & "-" & Me.Text1(2).Text & "-" & Left(Me.Text1(3).Text & "0", 1) & "-" & Left(Me.Text1(4).Text & "00", 2), "-0-00", "")
    If Me.Text2.Text = "5" Then
      Printer.Print "本所案號：" & Replace(Me.Text1(1).Text & "-" & Me.Text1(2).Text & "-" & Left(Me.Text1(3).Text & "0", 1) & "-" & Left(Me.Text1(4).Text & "00", 2), "-0-00", "") & "  文件公簽證"
    Else
      Printer.Print "本所案號：" & Replace(Me.Text1(1).Text & "-" & Me.Text1(2).Text & "-" & Left(Me.Text1(3).Text & "0", 1) & "-" & Left(Me.Text1(4).Text & "00", 2), "-0-00", "")
    End If
    'end 2016/10/18
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "案件名稱" & Me.lblCaseName(0).Caption
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    'Modify By Sindy 2009/09/04
    'Printer.Print "案件名稱" & Me.lblCaseName(1).Caption
    If Len(Trim(Me.lblCaseName(1).Caption)) > 30 Then
      Printer.Print "案件名稱" & Left(Trim(Me.lblCaseName(1).Caption), 30) & "..."
    Else
      Printer.Print "案件名稱" & Trim(Me.lblCaseName(1).Caption)
    End If
    '2009/09/04 End
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "案件名稱" & Me.lblCaseName(2).Caption
    m_dblLine = m_dblLine + 1
    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print "申請人：" & m_rsA.Fields(0).Value
    'Add by Morgan 2005/9/4 加申請人2~5
    For ii = 1 To 4
      If Not IsNull(m_rsA.Fields("PA" & Format(26 + ii))) Then
        m_dblLine = m_dblLine + 1
        dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
        dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
        Printer.CurrentX = dblPrtX
        Printer.CurrentY = dblPrtY
        Printer.Print "　　　　" & GetCustomerName(m_rsA.Fields("PA" & Format(26 + ii)).Value)
      End If
    Next
    
    If Me.Text2.Text <> "5" Then   'add by sonia 2016/10/18
      'Add By Sindy 2009/08/14
      m_dblLine = m_dblLine + 1
      dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
      dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
      Printer.CurrentX = dblPrtX
      Printer.CurrentY = dblPrtY
      Printer.Print "FAX：" & m_rsA.Fields("CU18").Value
      m_dblLine = m_dblLine + 1
      dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
      dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
      Printer.CurrentX = dblPrtX
      Printer.CurrentY = dblPrtY
      Printer.Print "TEL：" & m_rsA.Fields("CU16").Value
      '2009/08/14 End
      m_dblLine = m_dblLine + 1
      dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
      dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
      Printer.CurrentX = dblPrtX
      Printer.CurrentY = dblPrtY
      Printer.Print "申請國家：" & m_rsA.Fields(1).Value
    'add by sonia 2016/10/18
    Else
      m_dblLine = m_dblLine + 1
    End If
    'end 2016/10/18
'Modify by Morgan 2004/11/2 顯示與列印控制一致
'    m_dblLine = m_dblLine + 1
'    If Me.Text4.Text <> "" Then
'        strTxt = ""
'        intTxtLeng = 0
'        For ii = 1 To Len(Me.Text4.Text)
'            If Mid(Me.Text4.Text, ii, 1) = vbCr Then
'                m_dblLine = m_dblLine + 1
'                If Len(strTxt) > 0 Then
'                    dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
'                    dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
'                    Printer.CurrentX = dblPrtX
'                    Printer.CurrentY = dblPrtY
'                    Printer.Print strTxt
'                    strTxt = ""
'                    intTxtLeng = 0
'                End If
'                GoTo NextLine
'            ElseIf Mid(Me.Text4.Text, ii, 1) = vbLf Then
'                GoTo NextLine
'            End If
'            If Asc(Mid(Me.Text4.Text, ii, 1)) >= 0 And Asc(Mid(Me.Text4.Text, ii, 1)) < 128 Then
'                intTxtLeng = intTxtLeng + 1
'            Else
'                intTxtLeng = intTxtLeng + 2
'            End If
'            strTxt = strTxt & Mid(Me.Text4.Text, ii, 1)
'            If intTxtLeng >= 39 Then
'                m_dblLine = m_dblLine + 1
'                dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
'                dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
'                Printer.CurrentX = dblPrtX
'                Printer.CurrentY = dblPrtY
'                Printer.Print strTxt
'                strTxt = ""
'                intTxtLeng = 0
'            End If
'NextLine:
'        Next ii
'        m_dblLine = m_dblLine + 1
'        dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
'        dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
'        Printer.CurrentX = dblPrtX
'        Printer.CurrentY = dblPrtY
'        Printer.Print strTxt
'        strTxt = ""
'        intTxtLeng = 0
'   End If
   If Text4.Text <> "" Then
      'Printer.Font.Name = "MS Mincho"
      'Modify by Morgan 2006/7/28
      'strTxt = formatMemo(Text4.Text)
      strTxt = formatMemo1(Text4.Text)
      'end 2006/7/28
      '拆行
      arrLineString = Split(strTxt, vbCrLf)
      For ii = LBound(arrLineString) To UBound(arrLineString)
         m_dblLine = m_dblLine + 1
         dblPrtX = m_dblLeft + 8.5 + m_dblBetweenLine1
         dblPrtY = m_dblTop + m_dblTitleHeight + m_dblBetweenLine1 + m_dblLine * m_dblLineHeight1
         Printer.CurrentX = dblPrtX
         Printer.CurrentY = dblPrtY
         Printer.Print arrLineString(ii)
      Next
   End If
End Sub

'Add by Morgan 2006/7/28
'依照可列印寬度插入跳行符號
Private Function formatMemo1(p_Text) As String
   Dim ii As Integer, stPreChar As String, stCurChar As String, stTemp As String
   For ii = 1 To Len(p_Text)
      stCurChar = Mid(p_Text, ii, 1)
      If Not (stPreChar = " " And stCurChar = " ") Then
         If stCurChar = vbCr Then
            formatMemo1 = formatMemo1 & stTemp
            stTemp = stCurChar
         ElseIf Printer.TextWidth(stTemp & stCurChar) > 10.25 Then
            formatMemo1 = formatMemo1 & stTemp
            stTemp = vbCrLf
            If stCurChar <> " " Then
               stTemp = stTemp & stCurChar
            End If
         Else
            stTemp = stTemp & stCurChar
         End If
         stPreChar = stCurChar
      End If
   Next
   formatMemo1 = formatMemo1 & stTemp
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 插入分隔符號以便將文章拆字 Create by Morgan 2004/11/2
' stText:待拆文章
' iSepSign:分隔字元ASCII碼 預設30
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function getWords(ByVal stText As String, Optional ByVal iSepSign As Integer = 30) As String
   Dim stWords As String
   Dim ii As Integer, stPreChar As String, stCurChar As String
   For ii = 1 To Len(stText)
      stCurChar = Mid(stText, ii, 1)
      If stPreChar <> " " And stCurChar <> " " Then
         stWords = stWords & stCurChar
      Else
         stWords = stWords & Chr(iSepSign) & stCurChar
      End If
      stPreChar = stCurChar
   Next
   getWords = stWords
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 插入分隔符號以便將文章拆字 Create by Morgan 2004/11/2
' stText:待拆文章
' iMaxLen:行寬(半形字數) 預設40
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function formatMemo(ByVal stText As String, Optional iMaxLen As Integer = 40, Optional ByVal iPreLen As Integer) As String
   Dim iTLen As Integer '字串總長度
   Dim arrWords '單字陣列
   Dim stRtn As String '回傳字串
   Dim iWordLen As Integer '單字長度
   Dim iLineLen As Integer '一列字串長度
   Dim ii As Integer, jj As Integer
   Dim stChar As String
   Dim iRestLen As Integer '列剩餘可印長度
   Dim iWordPos As Integer '列剩餘可印字串位置
   Dim iCharLen As Integer '字元長度
   
   iTLen = Len(stText)
   If iTLen > 0 Then
      '拆字
      arrWords = Split(getWords(stText), Chr(30))
      iLineLen = iPreLen: stRtn = ""
      
      For ii = LBound(arrWords) To UBound(arrWords)
         iRestLen = iMaxLen - iLineLen
         If arrWords(ii) <> "" Then
            iWordLen = 0
            For jj = 1 To Len(arrWords(ii))
               stChar = Mid(arrWords(ii), jj, 1)
               iWordLen = iWordLen + 1
               '全形字算2個
               If Asc(stChar) < 0 Then iWordLen = iWordLen + 1
            Next jj
            '若字串超過列可印長度時斷字(例外情形)
            If iWordLen > iMaxLen - iPreLen Then
               iWordLen = 0: iWordPos = 0
               For jj = 1 To Len(arrWords(ii))
                  stChar = Mid(arrWords(ii), jj, 1)
                  iCharLen = 1
                  '全形字算2個
                  If Asc(stChar) < 0 Then iCharLen = 2
                  iWordLen = iWordLen + iCharLen
                  If iLineLen + iWordLen > iMaxLen Then
                     iWordPos = iWordPos + 1
                     If jj - iWordPos > 0 Then
                        stRtn = stRtn & Mid(arrWords(ii), iWordPos, jj - iWordPos)
                     End If
                     iWordPos = jj
                     stRtn = stRtn & vbCrLf & String(iPreLen, " ") & Mid(arrWords(ii), iWordPos, 1)
                     iLineLen = iPreLen + iCharLen
                     iWordLen = 0
                  End If
               Next jj
               arrWords(ii) = Mid(arrWords(ii), iWordPos + 1)
            End If
            
            '若超過每行固定長度則跳行
            If iLineLen + iWordLen > iMaxLen Then
               stRtn = stRtn & vbCrLf & String(iPreLen, " ") & arrWords(ii)
               iLineLen = iPreLen + iWordLen
            Else
               stRtn = stRtn & arrWords(ii)
               iLineLen = iLineLen + iWordLen
            End If
         End If
      Next ii
   End If
   formatMemo = stRtn
End Function

'Add By Sindy 2020/1/10
Private Sub txtRecvNo_GotFocus()
   TextInverse txtRecvNo
End Sub

'Added by Lydia 2021/06/17
Private Sub txtReviewer_GotFocus()
   TextInverse txtReviewer
   CloseIme
End Sub

'Added by Lydia 2022/08/12 FCT審查機關來電與回覆流程管制：與Frm030209_02共用
Private Sub cmdInput_Click()
   Set frm880004.mPreForm = Me
   frm880004.iStiu = 8
   frm880004.m_LCV01 = Text1(1) & Text1(2) & Text1(3) & Text1(4) & "," & Text1(5) & "," & Text1(8)
   frm880004.m_TempList = Me.cmdInput.Tag
   frm880004.Show vbModal
End Sub

'Added by Lydia 2022/08/12
Private Function FormSave() As Boolean
Dim tmpArr1 As Variant
Dim strTmpA As String, strCase(1 To 4) As String
Dim intA As Integer
Dim m_CP12 As String, m_CP13 As String, m_CP64 As String
Dim strSub As String, strCont As String

   If Text1(1) & Text1(2) & Text1(3) & Text1(4) = "" Then
       FormSave = True
       Exit Function
   End If
   
On Error GoTo ErrHandle

   m_CP64 = "審查員：" & txtReviewer & " " & IIf(Option1(0).Value = True, "先生", "小姐") & ", 分機號碼：" & Text1(7) & ", 備註：" & Text4
   m_CP64 = Replace(m_CP64, vbCrLf, " ")
   
   tmpArr1 = Split(Text1(1) & Text1(2) & Text1(3) & Text1(4) & "," & cmdInput.Tag, ",")
   m_CP13 = Trim(Left(Combo1.Text, 6))
   m_CP12 = GetST15(m_CP13)
   cnnConnection.BeginTrans
   For intA = 0 To UBound(tmpArr1)
      If Trim(tmpArr1(intA)) <> "" Then
          Call ChgCaseNo(tmpArr1(intA), strCase)
          If strCase(1) <> "" And strCase(2) <> "" Then
              strTmpA = "select nvl(tm05,nvl(tm06,tm07)) tm05,tm23, nvl(cu05,nvl(cu04,cu06)) cname1 From trademark, customer where substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) " & _
                      "and tm01='" & strCase(1) & "' and tm02='" & strCase(2) & "' and tm03='" & strCase(3) & "' and tm04='" & strCase(4) & "' "
              intI = 1
              Set RsTemp = ClsLawReadRstMsg(intI, strTmpA)
              If intI = 1 Then
                  strTmpA = AutoNo("C", 6)
                  '7/27 電話通知之承辦人及智權人員皆為受文者 ---- 阿蓮
                  strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP26,CP32, CP64,CP20) " & _
                            "VALUES ('" & strCase(1) & "','" & strCase(2) & "','" & strCase(3) & "','" & strCase(4) & "'," & strSrvDate(1) & "," & _
                            "'" & strTmpA & "','1727','" & m_CP12 & "','" & m_CP13 & "','" & m_CP13 & "','N','N','" & ChgSQL(m_CP64) & "','N') "
                  cnnConnection.Execute strSql
                  strSub = "Our Ref:" & strCase(1) & "-" & strCase(2) & IIf(strCase(3) <> "0", "-" & strCase(3), "") & IIf(strCase(4) <> "00", "-" & strCase(4), "") & " 電話通知 [INCOM.1727]"
                  'Modified by Lydia 2022/09/27 移除本所期限
                  'strCont = "審查員：" & txtReviewer & " " & IIf(Option1(0).Value = True, "先生", "小姐") & vbCrLf & _
                                "分　機：" & Text1(7) & vbCrLf & _
                                "本所案號：" & strCase(1) & "-" & strCase(2) & "-" & strCase(3) & "-" & strCase(4) & vbCrLf & _
                                "案件名稱：" & RsTemp.Fields("tm05") & vbCrLf & _
                                "申請人：" & RsTemp.Fields("tm23") & String(2, " ") & RsTemp.Fields("cname1") & vbCrLf & _
                                "備　註：" & vbCrLf & Text4 & vbCrLf & vbCrLf & _
                                "有關旨述電話通知 , 經與審查員聯繫後, 敬請協助:" & vbCrLf & _
                                "(   ) 1. 輸入期限" & vbCrLf & _
                                "　　     本所期限： 年  月 日" & vbCrLf & _
                                "　　     法定期限： 年  月 日" & vbCrLf & _
                                "   　    下一程序：(  )201補正/ (  )202申請意見書/ (  )206放棄專用權/(  )303延期/ (  )211檢送同意書/ (  )" & vbCrLf & _
                                "　　     輸入發文 : (  ) 是，已於  年  月  日通知代理人" & vbCrLf & _
                                "　　   　　    　　 (  ) 否，無需通知代理人" & vbCrLf & _
                                "　　     (   )此通知為多案，請輸入其他本所案號：" & vbCrLf & _
                                "(  ) 2. 無期限案件" & vbCrLf & _
                                "　　     輸入發文 : (  ) 是，已於  年  月  日通知代理人" & vbCrLf & _
                                "       　      　　 (  ) 否，無需通知代理人" & vbCrLf & vbCrLf & _
                                "電話內容紀錄(智權人員填具):" & vbCrLf
                  strCont = "審查員：" & txtReviewer & " " & IIf(Option1(0).Value = True, "先生", "小姐") & vbCrLf & _
                                "分　機：" & Text1(7) & vbCrLf & _
                                "本所案號：" & strCase(1) & "-" & strCase(2) & "-" & strCase(3) & "-" & strCase(4) & vbCrLf & _
                                "案件名稱：" & RsTemp.Fields("tm05") & vbCrLf & _
                                "申請人：" & RsTemp.Fields("tm23") & String(2, " ") & RsTemp.Fields("cname1") & vbCrLf & _
                                "備　註：" & vbCrLf & Text4 & vbCrLf & vbCrLf & _
                                "有關旨述電話通知 , 經與審查員聯繫後, 敬請協助:" & vbCrLf & _
                                "(   ) 1. 輸入期限" & vbCrLf & _
                                "　　     法定期限： 年  月 日" & vbCrLf & _
                                "   　    下一程序：(  )201補正/ (  )202申請意見書/ (  )206放棄專用權/(  )303延期/ (  )211檢送同意書/ (  )" & vbCrLf & _
                                "　　     輸入發文 : (  ) 是，已於  年  月  日通知代理人" & vbCrLf & _
                                "　　   　　    　　 (  ) 否，無需通知代理人" & vbCrLf & _
                                "　　     (   )此通知為多案，請輸入其他本所案號：" & vbCrLf & _
                                "(  ) 2. 無期限案件" & vbCrLf & _
                                "　　     輸入發文 : (  ) 是，已於  年  月  日通知代理人" & vbCrLf & _
                                "       　      　　 (  ) 否，無需通知代理人" & vbCrLf & vbCrLf & _
                                "電話內容紀錄(智權人員填具):" & vbCrLf
                  strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                        " VALUES ( '" & strUserNum & "','" & m_CP13 & "',to_char(sysdate,'yyyymmdd')" & _
                        ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strSub) & "','" & ChgSQL(strCont) & "')"
                  cnnConnection.Execute strSql
                  Sleep 100
              End If
          End If
      End If
   Next intA
   cnnConnection.CommitTrans
   
   FormSave = True
   Set RsTemp = Nothing
   Exit Function
   
ErrHandle:
   If Err.Number <> 0 Then
       MsgBox "存檔失敗: " & Err.Description, vbCritical
       cnnConnection.RollbackTrans
   End If
End Function

