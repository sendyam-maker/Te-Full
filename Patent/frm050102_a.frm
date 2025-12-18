VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050102_a 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文（授權）"
   ClientHeight    =   5100
   ClientLeft      =   456
   ClientTop       =   996
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   8520
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   6120
      MaxLength       =   4
      TabIndex        =   7
      Top             =   2727
      Width           =   540
   End
   Begin VB.TextBox txtChkRltDate 
      Height          =   270
      Left            =   5130
      MaxLength       =   8
      TabIndex        =   14
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   1050
      MaxLength       =   9
      TabIndex        =   8
      Top             =   3060
      Width           =   1095
   End
   Begin VB.TextBox Text12 
      Height          =   270
      Index           =   0
      Left            =   1050
      MaxLength       =   8
      TabIndex        =   12
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text12 
      Height          =   270
      Index           =   1
      Left            =   2490
      MaxLength       =   8
      TabIndex        =   13
      Top             =   3840
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   5040
      TabIndex        =   1
      Top             =   1740
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   7644
      TabIndex        =   20
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5592
      TabIndex        =   18
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   6420
      TabIndex        =   19
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   405
      Index           =   3
      Left            =   4368
      TabIndex        =   17
      Top             =   70
      Width           =   1200
   End
   Begin MSForms.TextBox Text8 
      Height          =   300
      Index           =   2
      Left            =   2790
      TabIndex        =   11
      Top             =   3510
      Width           =   5415
      VariousPropertyBits=   671107099
      MaxLength       =   60
      Size            =   "9551;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text8 
      Height          =   300
      Index           =   1
      Left            =   2790
      TabIndex        =   10
      Top             =   3270
      Width           =   5415
      VariousPropertyBits=   671107099
      MaxLength       =   60
      Size            =   "9551;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text8 
      Height          =   300
      Index           =   0
      Left            =   2790
      TabIndex        =   9
      Top             =   3030
      Width           =   5415
      VariousPropertyBits=   671107099
      MaxLength       =   60
      Size            =   "9551;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   4
      Left            =   735
      TabIndex        =   6
      Top             =   2730
      Width           =   1695
      VariousPropertyBits=   671107099
      Size            =   "2990;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   3
      Left            =   1560
      TabIndex        =   4
      Top             =   2400
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   615
      Index           =   8
      Left            =   135
      TabIndex        =   16
      Top             =   4410
      Width           =   8250
      VariousPropertyBits=   -1467987941
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "14552;1085"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   10
      Left            =   6120
      TabIndex        =   5
      Top             =   2430
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   9
      Left            =   6120
      TabIndex        =   3
      Top             =   2100
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   1755
      Width           =   1095
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Top             =   2100
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCP113 
      AutoSize        =   -1  'True
      Caption         =   "工作時數："
      Height          =   180
      Index           =   18
      Left            =   5190
      TabIndex        =   52
      Top             =   2775
      Width           =   900
   End
   Begin VB.Label lblChkRltDate 
      AutoSize        =   -1  'True
      Caption         =   "催審期限:"
      Height          =   180
      Left            =   4320
      TabIndex        =   50
      Top             =   3855
      Width           =   765
   End
   Begin VB.Label lblCaseFee 
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   6120
      TabIndex        =   15
      Tag             =   "Y"
      Top             =   3810
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "隨函附                                         資料"
      Height          =   180
      Left            =   135
      TabIndex        =   49
      Top             =   2730
      Width           =   3405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "被授權人:"
      Height          =   180
      Left            =   120
      TabIndex        =   48
      Top             =   3060
      Width           =   765
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "中:"
      Height          =   180
      Left            =   2430
      TabIndex        =   47
      Top             =   3060
      Width           =   225
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "英:"
      Height          =   180
      Left            =   2430
      TabIndex        =   46
      Top             =   3300
      Width           =   225
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "日:"
      Height          =   180
      Left            =   2430
      TabIndex        =   45
      Top             =   3540
      Width           =   225
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "授權期間:"
      Height          =   180
      Left            =   120
      TabIndex        =   44
      Top             =   3840
      Width           =   765
   End
   Begin VB.Label Label28 
      Caption         =   "~"
      Height          =   255
      Left            =   2250
      TabIndex        =   43
      Top             =   3840
      Width           =   135
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "是否修改通知函內容：        （Y：Word）"
      Height          =   180
      Left            =   4320
      TabIndex        =   42
      Top             =   2430
      Width           =   3225
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "是否修改指示信內容：        （Y:Word）"
      Height          =   180
      Index           =   1
      Left            =   4320
      TabIndex        =   41
      Top             =   2100
      Width           =   3090
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "進度備註："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   40
      Top             =   4185
      Width           =   900
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "是否列印通知函：        （N：不印）"
      Height          =   180
      Left            =   120
      TabIndex        =   39
      Top             =   2430
      Width           =   2820
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "發文日："
      Height          =   180
      Left            =   120
      TabIndex        =   38
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Left            =   4320
      TabIndex        =   37
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "是否列印指示信：        （N:不印）"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   36
      Top             =   2100
      Width           =   2685
   End
   Begin MSForms.Label lblAgent 
      Height          =   255
      Left            =   6300
      TabIndex        =   35
      Top             =   1770
      Width           =   2175
      VariousPropertyBits=   27
      Size            =   "3836;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblTrademarkKind 
      Height          =   255
      Left            =   5880
      TabIndex        =   28
      Top             =   720
      Width           =   2535
   End
   Begin MSForms.Label lblSalesName 
      Height          =   255
      Left            =   6000
      TabIndex        =   27
      Top             =   1080
      Width           =   2415
      VariousPropertyBits=   27
      Size            =   "4260;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   26
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   5280
      TabIndex        =   25
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   24
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   23
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   22
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   21
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   0
      Left            =   4320
      TabIndex        =   34
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   33
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利種類："
      Height          =   180
      Index           =   0
      Left            =   4320
      TabIndex        =   32
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   31
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號："
      Height          =   180
      Left            =   120
      TabIndex        =   30
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   180
      Index           =   1
      Left            =   4320
      TabIndex        =   29
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label lblCaseFees 
      BackColor       =   &H80000010&
      Height          =   255
      Left            =   6165
      TabIndex        =   51
      Top             =   3870
      Width           =   255
   End
End
Attribute VB_Name = "frm050102_a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/6 改成Form2.0 (txtCaseField,Text8,lblSalesName...)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'Create by Morgan 2009/7/23
Option Explicit

'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
Dim intCaseKind As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String
'intLeaveKind離開時，是0:結束  1:回上一畫面
Dim intLeaveKind As Integer
Dim m_bolActive As Boolean 'Active事件是否已觸發
'Add By Sindy 2018/1/8
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2018/1/8 END
Dim m_strAF01 As String, m_strLD18 As String 'Added by Morgan 2018/8/22

Private Sub cmdOK_Click(Index As Integer)

   Dim i As Integer, strTmp As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset

   Select Case Index
      Case 0, 3 '確定, 同時發文
         If PUB_ChkFileNP(cp(9)) Then
            MsgBox "下一程序已有提申或收達期限，不可發文！"
            Exit Sub
         End If
         
         Screen.MousePointer = vbHourglass
         '重新檢查欄位有效性
         If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
         
         If SaveDatabase Then
            '檢查代理人Email(需考慮可能為FF案件)
            PUB_CheckEMail cp(44), cp(116)
            PUB_CheckEMail field(75), field(144)
            If field(145) <> "" Then
               PUB_CheckEMail field(75), field(145)
            End If
            
            '指示信
            If txtCaseField(2) <> "N" Then
               strTmp = "30"
               NowPrint cp(9), "01", strTmp, IIf(txtCaseField(9).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_strAF01
               
               'Added by Morgan 2018/8/22 CFP電子化
               If txtCaseField(9).Text = "Y" And m_strAF01 <> "" Then
                  frm1105_1.m_RecNo = m_strAF01
                  frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & cp(10) & ".DATA.PDF"
                  frm1105_1.Show
                  If txtCaseField(10).Text = "Y" Then
                     MsgBox "指示信編輯中，客戶函請至定稿維護修改！", vbExclamation
                     txtCaseField(10).Text = ""
                  End If
               End If
               'end 2018/8/22
            End If
            
            '通知函
            If txtCaseField(3) <> "N" Then
               strTmp = "00"
               StartLetter "01", strTmp
               NowPrint cp(9), "01", strTmp, IIf(txtCaseField(10).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_strLD18
               
               'Added by Morgan 2018/8/22 CFP電子化
               If txtCaseField(10).Text = "Y" And m_strLD18 <> "" Then
                  frm1105_1.m_RecNo = m_strLD18
                  frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & cp(10) & ".CUS.PDF"
                  frm1105_1.Show
               End If
               'end 2018/8/22
            End If
            
            
            bolLeave = True
            intLeaveKind = 1
            '若有未發文資料顯示警告
            PUB_GetCPunIssueDatas "" & Me.lblCaseField(1).Caption
            ' 發文回前畫面時
            Select Case Index
               Case 0:
                  '回發文主畫面並清除畫面
                  'Add By Sindy 2013/5/28
                  If frm050102_1.bolIsEMPFlow = True Then
                     intLeaveKind = 0
                     'Unload frm050102_1
                     frm090202_4.Show
                     frm090202_4.QueryData
                  '2013/5/28 End
                  'Add By Sindy 2018/1/8
                  ElseIf Me.m_strIR01 <> "" Then
                     intLeaveKind = 0
                     'Modify By Sindy 2022/5/20
                     'frm04010519.GoNext
                     Forms(0).Tmpfrm04010519.GoNext
                     Set Forms(0).Tmpfrm04010519 = Nothing
                     '2022/5/20 END
                  '2018/1/8 END
                  Else
                     frm050102_1.Show
                     frm050102_1.Clear
                  End If
               Case 3:
                    '若尚有未發文資料
                    If PUB_ChkUnissueDatas(Me.lblCaseField(1).Caption) = True Then
                        '回發文主畫面並重新查詢
                        'Add By Sindy 2013/5/28
                        If frm050102_1.bolIsEMPFlow = True Then
                           frm090202_4.QueryData
                        'End If
                        '2013/5/28 End
                        'Add By Sindy 2018/1/8
                        ElseIf Me.m_strIR01 <> "" Then
                           'intLeaveKind = 0
                           'Modify By Sindy 2022/5/20
                           'frm04010519.GoNext
                           Forms(0).Tmpfrm04010519.GoNext
                           Set Forms(0).Tmpfrm04010519 = Nothing
                           '2022/5/20 END
                        '2018/1/8 END
                        End If
                        frm050102_1.Show
                        frm050102_1.ReQuery
                    '若無未發文資料
                    Else
                        '回發文主畫面並清除畫面
                        'Add By Sindy 2013/5/28
                        If frm050102_1.bolIsEMPFlow = True Then
                           intLeaveKind = 0
                           'Unload frm050102_1
                           frm090202_4.Show
                           frm090202_4.QueryData
                        '2013/5/28 End
                        'Add By Sindy 2018/1/8
                        ElseIf Me.m_strIR01 <> "" Then
                           intLeaveKind = 0
                           'Modify By Sindy 2022/5/20
                           'frm04010519.GoNext
                           Forms(0).Tmpfrm04010519.GoNext
                           Set Forms(0).Tmpfrm04010519 = Nothing
                           '2022/5/20 END
                        '2018/1/8 END
                        Else
                           frm050102_1.Show
                           frm050102_1.Clear
                        End If
                    End If
            End Select
            Unload Me
         Else
             MsgBox "存檔失敗, 請洽電腦中心人員!!!", vbExclamation + vbOKOnly
         End If
         Screen.MousePointer = vbDefault
      Case 1, 2
         'Add By Sindy 2013/5/28
         If frm050102_1.bolIsEMPFlow = True Then
            intLeaveKind = 0
            'Unload frm050102_1
            frm090202_4.Show
            frm090202_4.QueryData
         '2013/5/28 End
         'Add By Sindy 2018/1/8
         ElseIf Me.m_strIR01 <> "" Then
            intLeaveKind = 0
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         '2018/1/8 END
         Else
            If Index = 2 Then
               intLeaveKind = 0
            Else
               intLeaveKind = 1
            End If
         End If
         Unload Me
   End Select
End Sub

Private Function SaveDatabase() As Boolean
   Dim strTxt(1 To 10) As String, iStep As Integer, iIdx As Integer
   Dim StrSQLa As String
   Dim strTemp As String
   Dim strLetterJudge As String, strSubject As String '指示信判發人/主旨 Added by Morgan 2018/8/22
   
On Error GoTo CheckingErr

   cnnConnection.BeginTrans

   cp(27) = txtCaseField(0)
   
   '代理人
   intI = InStr(Combo1, "-")
   If intI > 0 Then
      cp(44) = Left(Combo1, intI - 1)
      cp(116) = Mid(Combo1, intI + 1)
   Else
      cp(44) = Combo1
      cp(116) = ""
   End If
   cp(44) = ChangeCustomerL(cp(44))
   
   '彼所案號
   'Modified by Morgan 2012/2/15 改呼叫共用函數
   'strExc(0) = "select cp45 from caseprogress where cp01=" + CNULL(cp(1)) + _
   '   " and cp02=" + CNULL(cp(2)) + " and cp03=" + CNULL(cp(3)) + _
   '   " and cp04=" + CNULL(cp(4)) + " and cp44=" + CNULL(cp(44)) + " order by cp27 desc"
   'intI = 1
   'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'cp(45) = ""
   'If intI = 1 And Not IsNull(RsTemp.Fields("CP45")) Then cp(45) = RsTemp.Fields("CP45")
   If Not ClsPDGetCaseThatCode(cp) Then cp(45) = ""
   'end 2012/2/15
   
   If Text6 <> "" Then
      cp(72) = ChangeCustomerL(Text6)
   End If
   
   cp(50) = Text8(0)
   cp(51) = Text8(1)
   cp(52) = Text8(2)
   cp(53) = DBDATE(Text12(0))
   cp(54) = DBDATE(Text12(1))
   cp(113) = txtCP113 'Added by Lydia 2021/05/25 工作時數
   
   strTxt(1) = GetCPSQL(cp())
   
   cnnConnection.Execute strTxt(1)

   '若案件國家收費表存在代理人收達天數則新增一筆收達的下一程序檔
   '判斷發文日非111111的才要
   If txtCaseField(0) <> "111111" Then
'Modify by Morgan 2009/11/11 收達期限管控改呼叫公用函式
'      Dim strCF23 As String
'      Dim strNPDate As String
'      Dim strNPSerial As String
'      If IsExistCasefee(cp(1), field(9), cp(10), strCF23) Then
'         strNPDate = DBDATE(Format(DateSerial(Val(DBYEAR(cp(27))), Val(DBMONTH(cp(27))), Val(DBDAY(cp(27))) + Val(strCF23))))
'         strNPSerial = InsertNextProgress_997(cp(9), cp(1), cp(2), cp(3), cp(4), strNPDate)
'      End If
      PUB_SetArriveDate cp(9)
'end 2009/11/11
   End If

  
   '提申管制
   'Modified by Morgan 2015/8/7 改呼叫共用
   PUB_SetApplyDate cp(1), cp(2), cp(3), cp(4), cp(7), cp(9), cp(10), txtCaseField(0), field(9)
   'end 2015/8/7
   
   'Add by Morgan 2009/8/18
   If txtChkRltDate <> "" Then
      PUB_UpdateChkResultDate txtChkRltDate, cp, cp(9), cp(10), cp(43)
   End If
   
   'Add by Sindy 2018/1/8
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm050102_1"
   End If
   '2018/1/8 END
   
   'Added by Morgan 2018/8/22 CFP電子化
   If strSrvDate(1) >= CFP指示信電子化啟用日 Then
      If txtCaseField(2) <> "N" Then
         strLetterJudge = PUB_GetLetterJudgeNew("2", cp(1), cp(10), field(9))
         strSubject = PUB_GetSubject(cp(1), cp(2), cp(3), cp(4), cp(10), field(11), cp(45), field(9))
         PUB_AddAppForm cp(9), True, strLetterJudge, strSubject
         m_strAF01 = cp(9)
      End If
   End If
   If strSrvDate(1) >= CFP第一階段電子化啟用日 Then
      If txtCaseField(3) <> "N" Then
         strLetterJudge = PUB_GetLetterJudgeNew("1", field(1), cp(10), field(9))
         PUB_AddLetterProgress cp(9), 0, True, strLetterJudge, False, field(26), cp(10), field(75)
         m_strLD18 = cp(9)
      End If
   End If
   'end 2018/8/22
   
   cnnConnection.CommitTrans
   SaveDatabase = True
   Exit Function
   
CheckingErr:
   cnnConnection.RollbackTrans
   
End Function

Private Sub ReadAllData()

   Dim rt As Boolean, i As Integer, varSaveCursor, strTemp As String, strTemp1 As String, j As Integer
   Dim adoRecord As Object, strSameName As String

On Error GoTo ErrHnd

   varSaveCursor = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   
   ReDim cp(TF_CP) As String
   cp(9) = frm050102_1.grdDataList.TextMatrix(frm050102_1.grdDataList.row, 5)
   If PUB_ReadAllData(cp(), field(), intCaseKind, intPWhere) Then

      lblCaseField(0) = cp(9)
      lblCaseField(1) = cp(1) + " - " + cp(2) + _
      IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
      IIf(cp(4) = "00", "", " - " + cp(4))
      lblCaseField(2) = TransDate(cp(6), 1)
      lblCaseField(4) = cp(13)
      lblCaseField(5) = TransDate(cp(7), 1)
      lblCaseField(3) = field(8)
      
      '代理人
        'Added by Lydia 2016/10/27 +新案有申請人指定國外代理人檔則預設
        If cp(31) = "Y" Then
           AddAgent Combo1, cp, , , , cp(9), field(9), field(26)
           If Combo1 <> "" Then CheckKeyIn 1
           
        Else '非新案照原本
           strSql = "select cp44||decode(cp116,null,null,'-'||cp116) from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "' and cp09<'C' and cp44 is not null order by cp27 desc"
           intI = 1
           Set adoRecord = ClsLawReadRstMsg(intI, strSql)
           If intI = 1 Then
              Do While Not adoRecord.EOF
                 If IsNull(adoRecord.Fields(0).Value) = False Then
                    If strSameName <> adoRecord.Fields(0).Value Then
                       Combo1.AddItem adoRecord.Fields(0).Value
                       strSameName = adoRecord.Fields(0).Value
                    End If
                 End If
                 adoRecord.MoveNext
              Loop
              Combo1 = Combo1.List(0)
           End If
           
         'Added by Morgan 2023/10/30 已有設定時不必再重新設定(IDS分案會先設,且抓預設代理人時也會剔除)
         If cp(44) <> "" Then
            Combo1 = cp(44) & IIf(cp(116) <> "", "-" & cp(116), "")
            CheckKeyIn 1
         Else
         'end 2023/10/30
         
           If ClsPDGetCasePreAgent(cp(), strTemp) Then
              Combo1 = strTemp
              CheckKeyIn 1
           End If
           
         End If 'Added by Morgan 2023/10/30
         
        End If
        'end 2016/10/27
   
      If ClsPDGetCaseProperty(cp(1), cp(10), strTemp) Then
         txtCaseField(4) = strTemp
      End If
    
      txtCaseField(2) = ""
      txtCaseField(9) = "Y"
      txtCaseField(3) = ""
      txtCaseField(10) = ""
      
      txtCaseField(8) = cp(64)
      
      'Add by Morgan 2009/8/18
      If txtCaseField(0).Tag <> txtCaseField(0) Then
         PUB_SetChkResultDate cp(1), field(9), cp(10), txtCaseField(0), txtChkRltDate, cp, field(8)
         txtCaseField(0).Tag = txtCaseField(0)
      End If
      'Added by Lydia 2021/05/25
      txtCP113 = ""
      If cp(113) <> "" Then txtCP113 = cp(113)
      'end 2021/05/25
         
   Else
      bolLeave = True
      intLeaveKind = 1
      Unload Me
   End If

ErrHnd:
   ErrorMsg
   Screen.MousePointer = varSaveCursor
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   Dim strNo As String, iPos As Integer
   If Combo1.Text <> "" Then
      If CheckKeyIn(1) = -1 Then
         Cancel = True
      End If
      '檢查客戶/代理人是否不再使用
      If Cancel = False Then
         strNo = Combo1.Text
         '聯絡人判斷
         iPos = InStr(Combo1.Text, "-")
         If iPos > 0 Then
            strNo = Left(Combo1.Text, iPos - 1)
         End If
         
         If PUB_CheckStatus(strNo) = False Then
            Cancel = True
         'Added by Morgan 2012/3/7 發文都要顯示代理人備註--甄妮
         Else
            strExc(0) = "select FA29 from Fagent where " & ChgFagent(strNo) & " and FA29 is not null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "" & RsTemp(0), vbExclamation, "代理人備註"
            End If
         'end 2012/3/7
         End If
      End If
      
      If Cancel Then Combo1.SetFocus
   End If
End Sub

Private Sub lblCaseField_Change(Index As Integer)
   Dim strTemp As String
   
   Select Case Index
   Case 3
      If ClsPDGetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , 台灣國家代號) = 1 Then
         lblTrademarkKind = strTemp
      End If
   Case 4
      If ClsPDGetStaffN(lblCaseField(Index), strTemp) Then
         lblSalesName = strTemp
      Else
         lblSalesName = ""
      End If
   End Select
End Sub
Private Sub Form_Activate()
   If m_bolActive = False Then
      txtCaseField(0) = strSrvDate(2)
      ReadAllData
      txtCaseField(0).SetFocus
      If PUB_ChkFileNP(cp(9)) Then MsgBox "下一程序已有提申或收達期限，若為重新發文時需要先刪除後才可作業！"
      m_bolActive = True
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   
   'Add By Sindy 2018/1/8
   m_strIR01 = frm050102_1.m_strIR01
   m_strIR02 = frm050102_1.m_strIR02
   m_strIR03 = frm050102_1.m_strIR03
   m_strIR04 = frm050102_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2018/1/8 END
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If bolLeave = False Then
      If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
         Cancel = 1
      End If
   End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Add by Morgan 2009/8/18
   If intLeaveKind = 1 Then
      frm050102_1.Show
   ElseIf intLeaveKind = 0 Then
     Unload frm050102_1
   End If
   ShowEditForm 'Added by Morgan 2018/8/22
    
   Set frm050102_a = Nothing
End Sub

Private Sub Text12_GotFocus(Index As Integer)
   TextInverse Text12(Index)
End Sub

Private Sub Text12_Validate(Index As Integer, Cancel As Boolean)
   If Text12(Index) <> "" Then
      If Not ChkDate(Text12(Index)) Then
         MsgBox "授權期間不正確，請重新輸入 !", vbCritical
         Cancel = True
      Else
         If Index = 1 Then
            If field(25) = "" Then
               MsgBox "專用期間止日不正確，請重新輸入 !", vbCritical
               Cancel = True
            Else
               If Val(Text12(1)) > Val(TransDate(field(25), 2)) Then
                  MsgBox "授權期間止日大於專用期間止日，請重新輸入 !", vbCritical
                  Cancel = True
               Else
                  If ChkRange(Text12(0), Text12(1), "授權期間") = False Then Cancel = True
               End If
            End If
         Else
            If field(24) = "" Then
               MsgBox "專用期間起日不正確，請重新輸入 !", vbCritical
               Cancel = True
            Else
               If Val(TransDate(field(24), 2)) > Val(Text12(0)) Then
                  MsgBox "授權期間起日小於專用期間起日，請重新輸入 !", vbCritical
                  Cancel = True
               End If
            End If
         End If
      End If
   End If
   If Cancel = True Then TextInverse Text12(Index)
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
   CloseIme
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   Dim i As Integer
   If Text6 <> "" Then
      If ClsLawGetCusCAJnam(Text6.Text, strExc(0), strExc(1), strExc(2)) = True Then
         For i = 0 To 2
            Text8(i) = strExc(i)
         Next
      Else
         MsgBox "被授權人編號不錯誤！"
         Cancel = True
      End If
   End If
End Sub

Private Sub Text8_GotFocus(Index As Integer)
   TextInverse Text8(Index)
End Sub

Private Sub txtCaseField_Change(Index As Integer)
   Select Case Index
   Case 1: lblAgent = ""
   End Select
End Sub
Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
   Case 1, 2, 3, 4, 5
      KeyAscii = UpperCase(KeyAscii)
   Case 9, 10
      KeyAscii = UpperCase(KeyAscii)
      If KeyAscii <> 8 And KeyAscii <> 89 Then
         KeyAscii = 0
      End If
   End Select
End Sub
Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
   If CheckKeyIn(Index) = -1 Then
      Cancel = True
   End If
   '檢查客戶/代理人是否不再使用
   If Cancel = False And Index = 6 Then
      If PUB_CheckStatus(txtCaseField(Index).Text) = False Then Cancel = True
   End If
   If Cancel Then txtCaseField_GotFocus (Index)
End Sub

Private Function CheckKeyIn(intIndex As Integer) As Integer
   Dim strTemp As String, strTemp1 As String, strCusTemp As String, j As Integer

   CheckKeyIn = -1
   Select Case intIndex
   Case 0 '發文日
      If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
         CheckKeyIn = 1
         'Add by Morgan 2009/8/18
         If txtCaseField(0).Tag <> txtCaseField(0) Then
            PUB_SetChkResultDate field(1), field(9), cp(10), txtCaseField(0), txtChkRltDate, cp, field(8)
            txtCaseField(0).Tag = txtCaseField(0)
         End If
      End If
   Case 1 '代理人
      lblAgent = ""
      If Combo1.Text = "" Then
         MsgBox "代理人欄不可空白!!!", vbExclamation
      Else
         strCusTemp = Combo1
         '判斷是否為聯絡人
         If InStr(strCusTemp, "-") > 0 Then
            If ClsPDGetContact(strCusTemp, strTemp) Then
               Combo1 = strCusTemp
               lblAgent.Caption = strTemp
               CheckKeyIn = 1
            End If
         
         ElseIf ClsPDGetAgent(strCusTemp, strTemp) Then
            Combo1 = strCusTemp
            lblAgent.Caption = strTemp
            CheckKeyIn = 1
         End If
      End If
      
   Case 2, 3 '是否列印指示信, 是否列印通知函
      If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
         CheckKeyIn = 1
      Else
         ShowMsg MsgText(1038)
      End If

   Case Else
      CheckKeyIn = 1
   End Select
End Function

Private Sub txtCaseField_GotFocus(Index As Integer)
   TextInverse txtCaseField(Index)
End Sub

Private Function TxtValidate() As Boolean
   Dim objTxt As Object
   Dim ii As Integer
   Dim Cancel As Boolean

   TxtValidate = False
   
   'Added by Morgan 2021/12/6 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   
   If IsDebt(field(9), cp(9)) Then
        MsgBox "未收款且無 預定收款日 請轉告智權同仁！！", vbOKOnly, "警告！禁止發文！"
        Exit Function
   End If
   
   If Text8(0) = "" And Text8(1) = "" And Text8(2) = "" Then
      MsgBox "被授權人名稱不可同時空白 !"
      Text8(0).SetFocus
      Exit Function
   End If
      
   For Each objTxt In Me.txtCaseField
      If objTxt.Enabled = True Then
         Cancel = False
         txtCaseField_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next

   If Combo1.Enabled = True Then
      If Combo1.Text = "" Then
         MsgBox "代理人欄不可空白!!!", vbExclamation
         Exit Function
      End If
      Cancel = False
      Combo1_Validate Cancel
      If Cancel = True Then
         Combo1.SetFocus
         Exit Function
      End If
   End If

'Added by Morgan 2018/9/12 CFP電子化-接洽單檢查
If strSrvDate(1) >= CFP第一階段電子化啟用日 Then
   If cp(9) < "B" And Left(cp(12), 1) <> "F" Then
      If PUB_CheckPDF3(cp(1), cp(2), cp(3), cp(4)) = False Then
         Exit Function
      End If
   End If
End If
'end 2018/9/12

'Added by Lydia 2021/05/25 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
If Pub_ChkACS112isNull(field(1), field(2), field(3), field(4), txtCP113) = True Then
    txtCP113.SetFocus
    txtCP113_GotFocus
    Exit Function
End If
'end 2021/05/25

   TxtValidate = True
End Function

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt(1 To 5) As String, strTmp As String, intStep As Integer, i As Integer
   EndLetter ET01, cp(9), ET03, strUserNum
   intStep = 1
   strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
      "','隨函附資料','" & txtCaseField(4) & "')"
   intStep = intStep + 1
   If Not ClsLawExecSQL(intStep - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub
'Add by Morgan 2009/8/18
Private Sub lblCaseFee_Click()
   frm12040102_2.txtCF(1) = cp(1)
   frm12040102_2.txtCF(2) = field(9)
   frm12040102_2.txtCF(3) = cp(10)
   frm12040102_2.Show vbModal
   If Val(txtCaseField(0)) > 0 Then
      PUB_SetChkResultDate cp(1), field(9), cp(10), txtCaseField(0), txtChkRltDate, cp, field(8)
   End If
End Sub

Private Sub lblCaseFee_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseDown lblCaseFee, lblCaseFees
End Sub

Private Sub lblCaseFee_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseUp lblCaseFee, lblCaseFees
End Sub

Private Sub txtChkRltDate_Validate(Cancel As Boolean)
   If txtChkRltDate <> "" Then
      If ChkDate(txtChkRltDate) = False Then
         Cancel = True
      End If
   End If
End Sub

'Added by Lydia 2021/05/25
Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

'Added by Lydia 2021/05/25
Private Sub txtCP113_Validate(Cancel As Boolean)
   If txtCP113 <> "" Then
      If Not IsNumeric(txtCP113) Then
         MsgBox "請輸入數字！", vbExclamation
         txtCP113.SetFocus
         txtCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub
