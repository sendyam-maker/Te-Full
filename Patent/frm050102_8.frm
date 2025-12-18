VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050102_8 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文（領證）"
   ClientHeight    =   5208
   ClientLeft      =   252
   ClientTop       =   1440
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5208
   ScaleWidth      =   8520
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   7680
      MaxLength       =   4
      TabIndex        =   17
      Top             =   3780
      Width           =   540
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "微個體"
      Height          =   255
      Index           =   2
      Left            =   2550
      TabIndex        =   12
      Top             =   3420
      Width           =   1605
   End
   Begin VB.TextBox txtChkRltDate 
      Height          =   270
      Left            =   5160
      MaxLength       =   8
      TabIndex        =   18
      Top             =   3780
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   960
      MaxLength       =   2
      TabIndex        =   2
      Top             =   1875
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   1830
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1875
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   5715
      MaxLength       =   9
      TabIndex        =   16
      Top             =   3450
      Width           =   1092
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "小個體"
      Height          =   255
      Index           =   1
      Left            =   1185
      TabIndex        =   11
      Top             =   3420
      Width           =   1335
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "大個體"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   3420
      Width           =   1035
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   5100
      TabIndex        =   1
      Top             =   1530
      Width           =   1215
   End
   Begin VB.CommandButton cmdCountry 
      Caption         =   "指定國家"
      Height          =   300
      Left            =   624
      TabIndex        =   14
      Top             =   3705
      Width           =   852
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   7644
      TabIndex        =   22
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5616
      TabIndex        =   20
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   6420
      TabIndex        =   21
      Top             =   70
      Width           =   1200
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   9
      Left            =   5880
      TabIndex        =   4
      Top             =   1920
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
      Index           =   1
      Left            =   3195
      TabIndex        =   15
      Top             =   3705
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
      Index           =   8
      Left            =   6240
      TabIndex        =   13
      Top             =   3150
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
      Index           =   7
      Left            =   6210
      TabIndex        =   6
      Top             =   2190
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
      Index           =   3
      Left            =   1260
      TabIndex        =   7
      Top             =   2535
      Width           =   1035
      VariousPropertyBits=   671107099
      MaxLength       =   9
      Size            =   "1826;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   4
      Left            =   1710
      TabIndex        =   8
      Top             =   2880
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
      Index           =   5
      Left            =   5790
      TabIndex        =   9
      Top             =   2850
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
      Index           =   2
      Left            =   1680
      TabIndex        =   5
      Top             =   2190
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
      Left            =   960
      TabIndex        =   0
      Top             =   1530
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
      Height          =   870
      Index           =   6
      Left            =   120
      TabIndex        =   19
      Top             =   4290
      Width           =   8295
      VariousPropertyBits=   -1467987941
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "14631;1535"
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
      Left            =   6720
      TabIndex        =   58
      Top             =   3825
      Width           =   900
   End
   Begin VB.Label Label12 
      Caption         =   "是否列印通知函：            （N:不印）"
      Height          =   180
      Left            =   4320
      TabIndex        =   57
      Top             =   1980
      Width           =   2895
   End
   Begin VB.Label lblChkRltDate 
      AutoSize        =   -1  'True
      Caption         =   "催審期限:"
      Height          =   180
      Left            =   4320
      TabIndex        =   55
      Top             =   3825
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
      Left            =   6180
      TabIndex        =   54
      Tag             =   "Y"
      Top             =   3750
      Width           =   255
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "是否含指定國註冊費        （Y/N）"
      Height          =   180
      Index           =   5
      Left            =   1530
      TabIndex        =   53
      Top             =   3765
      Width           =   2625
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "繳納"
      Height          =   180
      Left            =   120
      TabIndex        =   52
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "第                    至                  年年費"
      Height          =   180
      Index           =   4
      Left            =   495
      TabIndex        =   51
      Top             =   1920
      Width           =   2610
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "下次繳費期限："
      Height          =   180
      Left            =   4320
      TabIndex        =   50
      Top             =   3495
      Width           =   1260
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "美國發明是否含公開費            （Y:含公開費）"
      Height          =   180
      Index           =   3
      Left            =   4320
      TabIndex        =   49
      Top             =   3210
      Width           =   3585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "( 美、加、法國案 )"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   48
      Top             =   3180
      Width           =   1470
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "是否修改指示信內容：            （Y:Word）"
      Height          =   180
      Index           =   0
      Left            =   4320
      TabIndex        =   47
      Top             =   2250
      Width           =   3270
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "EPC："
      Height          =   180
      Left            =   120
      TabIndex        =   46
      Top             =   3765
      Width           =   495
   End
   Begin MSForms.Label lblNotify 
      Height          =   255
      Left            =   2355
      TabIndex        =   45
      Top             =   2565
      Width           =   6030
      VariousPropertyBits=   27
      Size            =   "1879;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "年費通知人："
      Height          =   180
      Left            =   120
      TabIndex        =   44
      Top             =   2595
      Width           =   1080
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "美國是否需先付款            （Y/N）"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   43
      Top             =   2940
      Width           =   2625
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "美國是否需附圖            （Y/N）"
      Height          =   180
      Index           =   1
      Left            =   4320
      TabIndex        =   42
      Top             =   2910
      Width           =   2490
   End
   Begin MSForms.Label lblAgent 
      Height          =   255
      Left            =   6330
      TabIndex        =   41
      Top             =   1560
      Width           =   2115
      VariousPropertyBits=   27
      Size            =   "3731;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "是否列印指示信：            （N:不印）"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   40
      Top             =   2250
      Width           =   2910
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Left            =   4320
      TabIndex        =   39
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "發文日："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   38
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "進度備註："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   37
      Top             =   4050
      Width           =   900
   End
   Begin VB.Label lblTrademarkKind 
      Height          =   255
      Left            =   5880
      TabIndex        =   30
      Top             =   570
      Width           =   2535
   End
   Begin MSForms.Label lblSalesName 
      Height          =   255
      Left            =   6000
      TabIndex        =   29
      Top             =   870
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
      TabIndex        =   28
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   5280
      TabIndex        =   27
      Top             =   870
      Width           =   615
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   26
      Top             =   570
      Width           =   495
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   25
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   24
      Top             =   870
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   23
      Top             =   540
      Width           =   3135
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   0
      Left            =   4320
      TabIndex        =   36
      Top             =   870
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利種類："
      Height          =   180
      Index           =   0
      Left            =   4320
      TabIndex        =   34
      Top             =   570
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   33
      Top             =   870
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號："
      Height          =   180
      Left            =   120
      TabIndex        =   32
      Top             =   570
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   180
      Index           =   1
      Left            =   4320
      TabIndex        =   31
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label lblCaseFees 
      BackColor       =   &H80000010&
      Height          =   255
      Left            =   6165
      TabIndex        =   56
      Top             =   3810
      Width           =   255
   End
End
Attribute VB_Name = "frm050102_8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/6 改成Form2.0 (txtCaseField,lblSalesName...)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'2005/7/11整理
Option Explicit

'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
Dim intCaseKind As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String
'intLeaveKind離開時，是0:結束  1:回上一畫面
Dim intLeaveKind As Integer
'StrCountry存放指定國家  strLicenceCountry存放領證國家
Dim strCountry As String, strLicenceCountry As String
'Add By Cheng 2002/10/02
Dim m_strCountryEngName As String '指定國家的英文名稱
'92.1.12 add by sonia
Dim old_Entity As String   '原大小個體
Dim new_Entity As String   '新大小個體
'Add By Cheng 2003/04/01
Dim m_blnFormFirstShow As Boolean '註記表單第一次顯示
'Add by Morgan 2007/8/15
Dim varFeeYears As Variant '繳費年度陣列
Dim m_FeeProperty As String '年費案件性質
Dim m_StartDate As String '年費起算日
'Add By Sindy 2018/1/8
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2018/1/8 END
Dim m_strAF01 As String, m_strLD18 As String 'Added by Morgan 2018/8/22
Dim m_strCP81 As String, m_strJpMemo As String 'Added by Morgan 2019/4/30
Dim m_strReduceOne As String 'Added by Morgan 2019/12/11
Dim strCP09List  As String 'Added by Morgan 2020/8/14 子案指示信-子案總收文號
Dim str249Msg As String 'Added by Morgan 2025/6/20

Private Sub cmdCountry_Click()
   'Modified by Morgan 2023/3/7 +傳入案件性質
   ModifyLicenceCountry strCountry, strLicenceCountry, , cp(10)
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim stLetter As String 'Add by Morgan 2004/9/27
Dim i As Integer, strTmp As String
'Add By Cheng 2002/07/31
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim arrAF01() As String 'Added by Morgan 2010/8/14 子案指示信-子案總收文號

   Select Case Index
      Case 0
      
         'Added by Morgan 2015/8/7
         If PUB_ChkFileNP(cp(9)) Then
            MsgBox "下一程序已有提申或收達期限，不可發文！"
            Exit Sub
         End If
         'end 2015/8/7
   
         'Modify by Morgan 2007/12/27
         'If field(9) = EPC指定國家 And strLicenceCountry = "" Then
         '   MsgBox "未輸入領證之指定國家 !", vbCritical
         '   Exit Sub
         'End If
         If field(9) = EPC指定國家 Then
            If txtCaseField(1) = "" Then
               MsgBox "是否含指定國註冊費不可空白 !", vbCritical
               txtCaseField(1).SetFocus
               Exit Sub
            ElseIf txtCaseField(1) = "Y" Then
               If strLicenceCountry = "" Then
                  MsgBox "未輸入領證之指定國家 !", vbCritical
                  Exit Sub
               End If
            End If
         End If
         'end 2007/12/27
         
         '92.1.12 add by sonia
         'Modify by Morgan 2006/9/20 加法國且申請日>=20050901
         If field(9) = "101" Or field(9) = "102" Or (field(9) = "203" And (field(10) = "" Or DBDATE(field(10)) >= "20050901")) Then
            'Modified by Morgan 2013/3/20 +微個體
            If Not optChoose(0).Value And Not optChoose(1).Value And Not optChoose(2).Value Then
               If optChoose(2).Enabled = True Then
                  MsgBox "請選擇" & optChoose(0).Caption & "、" & optChoose(1).Caption & "或" & optChoose(2).Caption & "資料 !", vbCritical
               Else
                  MsgBox "請選擇" & optChoose(0).Caption & "或" & optChoose(1).Caption & "資料 !", vbCritical
               End If
               Exit Sub
            End If
         End If
         '92.1.12 end
         'Add By Cheng 2002/10/02
         m_strCountryEngName = ""
         
         'Add by Morgan 2008/10/8
         If Text5(0).Enabled = True And Trim(Text5(0)) = "" Then
            'Add by Morgan 2008/10/15 加判斷年費期限有續辦的才要 CFP-18959
            strExc(0) = "select np01 from nextprogress where np02='" & field(1) & "' and np03='" & field(2) & "' and np04='" & field(3) & "' and np05='" & field(4) & "' and np06='Y' and np07='" & m_FeeProperty & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "准後繳且年費期限有續辦，繳費年度不可空白！"
               Text5(0).SetFocus
               Exit Sub
            End If
         End If
         
         Screen.MousePointer = vbHourglass
         'Modify By Cheng 2002/10/02
'         For i = 0 To 6
         For i = 0 To 7
            If i <> 1 Then
                If txtCaseField(i).Enabled Then
                      If CheckKeyIn(i) = -1 Then
                         txtCaseField(i).SetFocus
                         txtCaseField_GotFocus (i)
                         Exit For
                      End If
                End If
            'Add By Cheng 2002/08/19
            Else
               If CheckKeyIn(i) <> 1 Then
                  Me.Combo1.SetFocus
                  Exit For
               End If
            End If
         Next
         If i = 8 Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
            str249Msg = "" 'Added by Morgan 2025/6/20
            If SaveDatabase Then
               If str249Msg <> "" Then MsgBox str249Msg, vbInformation 'Added by Morgan 2025/6/20
               'Add by Morgan 2008/2/20 檢查代理人Email(需考慮可能為FF案件)
               PUB_CheckEMail cp(44), cp(116)
               PUB_CheckEMail field(75), field(144)
               If field(145) <> "" Then
                  PUB_CheckEMail field(75), field(145)
               End If
               'end 2008/2/20
               
               '若美國需附圖, 則列印TNT
               If Me.txtCaseField(5).Text = "Y" Then
                    'Modified by Lydia 2014/12/27 + DHL
                    If MsgBox("請選擇Yes(TNT列印)或No(DHL列印)", vbYesNo + vbInformation + vbDefaultButton1) = vbYes Then
                        '列印TNT
                        Screen.MousePointer = vbDefault
                        frm060321.Show
                        bolToEndByNick = False
                        frm060321.Hide
                        frm060321.GetCP09 = cp(9)
                        frm060321.txt1(0).Text = cp(1)
                        frm060321.txt1(1).Text = cp(2)
                        frm060321.txt1(2).Text = cp(3)
                        frm060321.txt1(3).Text = cp(4)
                        frm060321.txt1(0).Enabled = False
                        frm060321.txt1(1).Enabled = False
                        frm060321.txt1(2).Enabled = False
                        frm060321.txt1(3).Enabled = False
                        Me.Enabled = False
                        frm060321.Show
                        Do
                            DoEvents
                            If bolToEndByNick = True Then Exit Do
                        Loop Until Not frm060321.Visible
                        Unload frm060321
                        Me.Enabled = True
                    Else
                        '列印DHL
                        Screen.MousePointer = vbDefault
                        frm060330.Show
                        bolToEndByNick = False
                        frm060330.Hide
                        'frm060330.GetCP09 = cp(9) 'mark by Lydia 2022/03/28
                        frm060330.txt1(0).Text = cp(1)
                        frm060330.txt1(1).Text = cp(2)
                        frm060330.txt1(2).Text = cp(3)
                        frm060330.txt1(3).Text = cp(4)
                        frm060330.txt1(0).Enabled = False
                        frm060330.txt1(1).Enabled = False
                        frm060330.txt1(2).Enabled = False
                        frm060330.txt1(3).Enabled = False
                        Me.Enabled = False
                        frm060330.Show
                        Do
                            DoEvents
                            If bolToEndByNick = True Then Exit Do
                        Loop Until Not frm060330.Visible
                        Unload frm060330
                        Me.Enabled = True
                    
                    End If
               End If
               '指示信
               If txtCaseField(2) <> "N" Then
                  Select Case field(9)
                     'Added by Morgan 2020/9/18
                     Case "011" '日本
                        strTmp = "38"
                     'end 2020/9/18
                     Case "102" '加拿大領證
                        strTmp = "36"
                     Case "101" '美國
                        If txtCaseField(4) = "Y" And txtCaseField(5) = "Y" Then
                           '美國需先付款、附圖領證
                           strTmp = "35"
                        ElseIf txtCaseField(4) = "Y" And txtCaseField(5) = "N" Then
                           '美國需先付款、不附圖領證
                           strTmp = "34"
                        ElseIf txtCaseField(4) = "N" And txtCaseField(5) = "Y" Then
                           '美國不付款、附圖領證
                           strTmp = "33"
                        ElseIf txtCaseField(4) = "N" And txtCaseField(5) = "N" Then
                           '美國不付款、不附圖領證
                           strTmp = "32"
                        End If
                     Case "221" 'EPC領證
                        'Modify by Morgan 2007/12/27
                        'strTmp = "31"
                        If Me.txtCaseField(1) = "N" Then
                           strTmp = "37"
                        Else
                           strTmp = "31"
                        End If
                        'end 2007/12/27
                     Case Else
                         '一般領證
                        strTmp = "30"
                  End Select
                  
                  StartLetter "01", strTmp
                  
                  'Modify by Morgan 2004/9/27
                  '領證加印傳真封面
                  'Removed by Morgan 2018/10/22 取消傳真封面--慧汶
                  'If txtCaseField(7).Text = "Y" Then
                  '   NowPrint cp(9), "01", "89", False, strUserNum, , , True, stLetter, , , , , , , , , m_strAF01
                  'Else
                  '   NowPrint cp(9), "01", "89", False, strUserNum, , , , , , , , , , , , , m_strAF01
                  'End If
                  'If m_strAF01 <> "" Then Sleep 1000 '等1秒以確保letterdemand不會發生dupe錯誤 Added by Morgan 2018/8/20
                  'end 2018/10/22
                  NowPrint cp(9), "01", strTmp, IIf(txtCaseField(7).Text = "Y", True, False), strUserNum, 0, stLetter, , , , , , , , , , , m_strAF01
                  '2004/9/27 end
                  
                  'Added by Morgan 2018/8/22 CFP電子化
                  If Me.txtCaseField(7).Text = "Y" And m_strAF01 <> "" Then
                     frm1105_1.m_RecNo = m_strAF01
                     frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & cp(10) & ".DATA.PDF"
                     frm1105_1.Show
                  End If
                  'end 2018/8/22
                  
                  'Added by Morgan 2020/8/14 EPC指定國註冊費子案指示信
                  If field(9) = EPC指定國家 And txtCaseField(1) = "Y" And Len(strCP09List) > 0 Then
                     arrAF01 = Split(strCP09List, ",")
                     For i = 0 To UBound(arrAF01)
                        '沒有例外欄位
                        NowPrint arrAF01(i), "01", "28", False, strUserNum, , , , , , , , , , , , , arrAF01(i)
                     Next
                     MsgBox "子案指示信請至待處理區作業！", vbExclamation
                  End If
                  'end 2020/8/14
                  
'Removed by Morgan 2012/3/7 不必詢問是否列印代理人小信封,需要時程序自行列印--甄妮

               End If
               
               'Added by Morgan 2019/6/17
               '通知函
               If txtCaseField(9) <> "N" Then
                  strTmp = "00"
                  NowPrint cp(9), "01", strTmp, False, strUserNum, 0, , , , , , , , , , , , m_strLD18
               End If
               'end 2019/6/17
            
               bolLeave = True
               intLeaveKind = 1
               'Add By Cheng 2002/04/30
               '若有未發文資料顯示警告
               PUB_GetCPunIssueDatas "" & Me.lblCaseField(1).Caption
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
                  frm050102_1.Clear
               End If
               Unload Me
            '911202 nick
            Else
                MsgBox "存檔失敗, 請洽電腦中心人員!!!", vbExclamation + vbOKOnly
            End If
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
   
    ' 發文回前畫面時
   Select Case Index
      Case 0:
         ' 90.07.12 modify by louis (回發文主畫面並清除畫面)
         
   End Select
End Sub

'Add By Cheng 2002/10/02
Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
 Dim strTxt(1 To 7) As String, iStep As Integer, strTmp As String
   EndLetter ET01, cp(9), ET03, strUserNum
   Dim Jjj As Integer
   Jjj = 1
   If m_strCountryEngName <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
         "','指定國家','" & m_strCountryEngName & "')"
      Jjj = Jjj + 1
   End If
   '92.1.12 add by sonia
   'Modify by Morgan 2006/9/21 加法國且申請日>=20050901
   If field(9) = "101" Or field(9) = "102" Or (field(9) = "203" And (field(10) = "" Or DBDATE(field(10)) >= "20050901")) Then
      If optChoose(0).Value = True Then
         strTmp = "(Large Entity)"
      ElseIf optChoose(1).Value = True Then
         strTmp = "(Small Entity)"
      'Added by Morgan 2013/3/20
      ElseIf optChoose(2).Value = True Then
         strTmp = "(Micro Entity)"
      'end 2013/3/20
      End If
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
         "','大小個體','" & strTmp & "')"
      Jjj = Jjj + 1
   End If
   
   If field(9) = "101" And field(8) = "1" And txtCaseField(8) = "Y" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
         "','公開與否','and publication fee ')"
      Jjj = Jjj + 1
   End If
   '92.1.12 end
   Select Case field(9)
      Case "011", "012" '日、韓領證
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','案件性質分類','registration fee and 1-3 annuities fee')"
         Jjj = Jjj + 1
      Case Else
         'Add by Morgan 2009/6/16
         '新加坡
         If field(9) = "014" Then
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','案件性質分類','grant fee')"
            Jjj = Jjj + 1
         'end 2009/6/16
         ElseIf field(9) >= "013" And field(9) <= "099" Then
            '東南亞領證
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','案件性質分類','registration fee')"
            Jjj = Jjj + 1
         ElseIf field(9) >= "201" And field(9) <= "299" Then
            '歐洲領證
            'Modify By Cheng 2003/12/30
            'grant-->granting
'            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
'               "','案件性質分類','grant and printing fees')"
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','案件性質分類','granting and printing fees')"
            'End
            Jjj = Jjj + 1
         Else
             '一般領證
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','案件性質分類','issue fee')"
            Jjj = Jjj + 1
         End If
   End Select
   
   'Add by Morgan 2007/8/27
   If Text5(0).Enabled = True And Text5(0) <> "" Then
      'Modify by Moran 2008/12/19 印尼發明不用說明繳費年度
      If field(9) = "017" And field(8) = "1" Then
         strExc(0) = " and annuities"
      Else
      'end 2008/12/19
         If Text5(1).Enabled = True And Text5(1) <> Text5(0) Then
            strExc(0) = " and " & TranslateKeyWord(incCNV_ENGLISH_FREQUENCY, Text5(0), "") & "-" & TranslateKeyWord(incCNV_ENGLISH_FREQUENCY, Text5(1), "") & " annuities"
         Else
            strExc(0) = " and " & TranslateKeyWord(incCNV_ENGLISH_FREQUENCY, Text5(0), "") & " annuity"
         End If
      End If
      '一般領證
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
         "','繳年費期間','" & strExc(0) & "')"
      Jjj = Jjj + 1
   End If
   'end 2007/8/27
   
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(Jjj - 1, strTxt) Then
   If Not ClsLawExecSQL(Jjj - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
   '*******************************************
End Sub

Private Function SaveDatabase() As Boolean

   Dim i As Integer
   Dim varTmp As Variant, varTmp1 As Variant, pa04 As String
   Dim str930Date As String 'Added by Lydia 2017/11/02
   Dim strLetterJudge As String, strSubject As String '指示信判發人/主旨 Added by Morgan 2018/8/22
   Dim strAgentList As String 'Added by Morgan 2020/8/14
   Dim bolAdd250NP As Boolean 'Added by Morgan 2023/3/9 是否管制 UPC選擇退出
   
On Error GoTo CheckingErr

   cnnConnection.BeginTrans
   
   cp(27) = txtCaseField(0)
   
   'Modify by Morgan 2008/2/21
   'cp(44) = Combo1
   intI = InStr(Combo1, "-")
   If intI > 0 Then
      cp(44) = Left(Combo1, intI - 1)
      cp(116) = Mid(Combo1, intI + 1)
   Else
      cp(44) = Combo1
      cp(116) = ""
   End If
   'end 2008/2/21
   cp(44) = ChangeCustomerL(cp(44))
   
   'Modify by Morgan 2006/9/21 加法國且申請日>=20050901
   'Modified by Morgan 2023/3/25
   'If field(9) = "101" Or field(9) = "102" Or (field(9) = "203" And (field(10) = "" Or DBDATE(field(10)) >= "20050901")) Then
   If InStr(CFP_ChkEntity, field(9)) > 0 Or field(179) <> "" Then
      If strSrvDate(1) >= PA179啟用日 Then
         If optChoose(0).Value = True Then
            new_Entity = optChoose(0).Caption
         ElseIf optChoose(1).Value = True Then
            new_Entity = optChoose(1).Caption
         ElseIf optChoose(2).Value = True Then
            new_Entity = optChoose(2).Caption
         End If
         
      Else
   'end 2023/3/25
   
         If optChoose(0).Value = True Then
            new_Entity = "大個體"
         ElseIf optChoose(1).Value = True Then
            new_Entity = "小個體"
         'Added by Morgan 2013/3/20
         ElseIf optChoose(2).Value = True Then
            new_Entity = "微個體"
         'end 2013/3/20
         End If
         
      End If 'Added by Morgan 2023/3/25
      
      If old_Entity <> new_Entity And old_Entity <> "" Then  '改大小個體時
         If txtCaseField(6) = "" Then
            cp(64) = "原大小個體為" & old_Entity
         Else
            cp(64) = "原大小個體為" & old_Entity & "，" & Me.txtCaseField(6).Text
         End If
      Else
         cp(64) = txtCaseField(6)
      End If
   Else
      cp(64) = txtCaseField(6)
   End If
   
   If m_strJpMemo <> "" Then cp(64) = m_strJpMemo & ";" & cp(64) 'Added by Morgan 2019/4/30 日本領證可減免備註加減免身分
   cp(81) = m_strCP81 'Added by Morgan 2019/4/30
   cp(113) = txtCP113 'Added by Lydia 2021/05/25 工作時數
   
   'Modified by Morgan 2012/2/15 改呼叫共用函數
   'strExc(0) = "select cp45 from caseprogress where cp01=" + CNULL(cp(1)) + _
   '   " and cp02=" + CNULL(cp(2)) + " and cp03=" + CNULL(cp(3)) + _
   '   " and cp04=" + CNULL(cp(4)) + " and cp44=" + CNULL(cp(44)) + " order by cp27 desc"
   'intI = 1
   'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   'cp(45) = ""
   'If intI = 1 And Not IsNull(RsTemp.Fields("CP45")) Then cp(45) = RsTemp.Fields("CP45")
   If Not ClsPDGetCaseThatCode(cp) Then cp(45) = ""
   'end 2012/2/15

   '2012/11/6 ADD BY SONIA CFP-023832
   cp(53) = Text5(0): cp(54) = Text5(1)
   '2012/11/6 END

   'Add by Morgan 2007/12/27
   If field(9) = EPC指定國家 And txtCaseField(1) = "N" Then
      cp(64) = cp(64) & ";未繳指定國註冊費;"
   End If
   'end 2007/12/27
   
   strSql = GetCPSQL(cp())
   
   cnnConnection.Execute strSql, intI
   
   '92.1.12 add by sonia 改大小個體時
   'Modify by Morgan 2006/9/21 加法國且申請日>=20050901
   'Modified by Morgan 2023/3/25
   'If field(9) = "101" Or field(9) = "102" Or (field(9) = "203" And (field(10) = "" Or DBDATE(field(10)) >= "20050901")) Then
   If InStr(CFP_ChkEntity, field(9)) > 0 Or field(179) <> "" Then
      If strSrvDate(1) >= PA179啟用日 Then
         If optChoose(0).Value = True Then
            field(179) = "1"
         ElseIf optChoose(1).Value = True Then
            field(179) = "2"
         ElseIf optChoose(2).Value = True Then
            field(179) = "3"
         End If
      Else
   'end 2023/3/25
   
         If old_Entity <> new_Entity Then
            If InStr(1, field(91), old_Entity, 1) > 0 Then
               field(91) = Replace(field(91), old_Entity, new_Entity, InStr(1, field(91), old_Entity, 1), , 1)
            Else
               If field(91) = "" Then
                  field(91) = new_Entity
               Else
                  field(91) = new_Entity & "，" & field(91)
               End If
            End If
         End If
         
      End If 'Added by Morgan 2023/3/25
   End If
   '92.1.12 end
   field(76) = txtCaseField(3)
   strSql = GetPASQL(field())
   
   cnnConnection.Execute strSql, intI
   
   'Modify by Morgan 2007/12/27
   'If cmdCountry.Visible Then
   If cmdCountry.Enabled = True Then
      
      'Add by Morgan 2008/9/19 要先刪除舊的指定國領證資料(因若重發文會有多筆，如此發證時可能會無法確定原領證國 Ex.CFP-15807)
      strSql = "delete from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04<>'00' and cp10='601'"
      cnnConnection.Execute strSql, intI
      'Modify by Morgan 2006/12/25
      'If objPublicData.SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strLicenceCountry) Then
      'Modified by Morgan 2020/8/14 領證不必寫子案，若有含指定國註冊費時子案要新增224指定國註冊費程序
      'If PUB_SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strLicenceCountry) Then
      'If strLicenceCountry <> "" Then 'Removed by Morgan 2020/9/3 不含指定國註冊費時不會點子案
      'end 2020/8/14
      'end 2006/12/25
      
         'Add by Morgan 2007/12/27
         If txtCaseField(1) = "N" Then
            '新增[224指定國註冊費]下一程序
            '法限=發文日+3個月
            strExc(3) = CompDate(1, 3, strSrvDate(1))
            '所限=發文日+1個月
            strExc(4) = PUB_GetWorkDay1(CompDate(1, 1, strSrvDate(1)), True)
            '智權人員
            strExc(5) = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
            '流水號
            strExc(6) = GetNextProgressNo()
           
            strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','224'," & strExc(4) & "," & strExc(3) & ",'" & strExc(5) & "'," & strExc(6) & ") "
            cnnConnection.Execute strSql, intI
            
            'Added by Morgan 2023/3/7
            '新增[249UP註冊]下一程序
            '法限=發文日+2個月
            strExc(3) = CompDate(1, 2, strSrvDate(1))
            '所限=發文日+1個月
            strExc(4) = PUB_GetWorkDay1(CompDate(1, 1, strSrvDate(1)), True)
            '智權人員
            strExc(5) = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
            '流水號
            strExc(6) = GetNextProgressNo()
           
            strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','249'," & strExc(4) & "," & strExc(3) & ",'" & strExc(5) & "'," & strExc(6) & ") "
         
            cnnConnection.Execute strSql, intI
            'end 2023/3/7
         
         'Modified by Morgan 2020/9/3
         'Else
         ElseIf strLicenceCountry <> "" Then
         'end 2020/9/3
         'end 2007/12/27
            varTmp = Split(strCountry, ",")
            varTmp1 = Split(strLicenceCountry, ",")
            For i = 0 To UBound(varTmp)
               If InStr(strLicenceCountry, Format(varTmp(i))) = 0 Then
                  pa04 = GetPA04(field(1), field(2), field(3), Format(varTmp(i)))
                  strSql = "UPDATE PATENT SET PA57='Y' WHERE PA01='" & field(1) & "' AND PA02='" & field(2) & "' AND PA03='" & field(3) & "' AND PA04='" & pa04 & "'"
                  cnnConnection.Execute strSql, intI
               Else
                   m_strCountryEngName = m_strCountryEngName & ", " & PUB_GetNationEngName("" & varTmp(i))
               End If
            Next
            If m_strCountryEngName <> "" Then
               m_strCountryEngName = Right(m_strCountryEngName, Len(m_strCountryEngName) - 2)
            End If
               
            'Added by Morgan 2020/8/14
            strAgentList = PUB_GetAgentList(field(1), field(2), field(3), strLicenceCountry)
            If PUB_SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strLicenceCountry, strAgentList, strCP09List, "224") Then
               Dim ArrCP09() As String, arrCP(4) As String
               ArrCP09 = Split(strCP09List, ",")
               For i = 0 To UBound(varTmp1)
                  pa04 = GetPA04(field(1), field(2), field(3), Format(varTmp1(i)))
                  strLetterJudge = PUB_GetLetterJudgeNew("2", cp(1), "224", Format(varTmp1(i)))
                  strSubject = PUB_GetSubject(cp(1), cp(2), cp(3), pa04, "224", field(11), , Format(varTmp1(i)))
                  PUB_AddAppForm ArrCP09(i), True, strLetterJudge, strSubject
                  
                  arrCP(1) = cp(1): arrCP(2) = cp(2): arrCP(3) = cp(3): arrCP(4) = pa04
                  'Added by Morgan 2023/3/9
                  'UP子案案件性質改為249UP註冊,收達、一般提申走通用規則,催審期限=領證期限(發證日)+2個月、最終提申=領證期限(發證日)+1個月
                  If Format(varTmp1(i)) = "224" Then
                     bolAdd250NP = False
                     strSql = "update caseprogress set cp10='249' where cp43='" & cp(9) & "' and cp10='224' and cp09='" & ArrCP09(i) & "'"
                     cnnConnection.Execute strSql, intI
                     
                     '催審期限
                     strExc(3) = cp(7)
                     strExc(1) = PUB_Get224CtrlDate(1, strExc(3), cp, True)
                     PUB_UpdateChkResultDate strExc(1), arrCP, ArrCP09(i), "249"
                     
                     '最終提申期限
                     strExc(2) = PUB_Get224CtrlDate(2, strExc(3), cp, True)
                     PUB_SetApplyDate cp(1), cp(2), cp(3), pa04, strExc(2), ArrCP09(i), "249", txtCaseField(0), Format(varTmp1(i))
                     
                     'Added by Morgan 2025/6/20
                     '若有收文249UP註冊則一併上發文
                     strExc(1) = "": strExc(2) = ""
                     If PUB_ChkCPExist(cp, "249", 1, strExc(1)) = True Then
                        strSql = "update caseprogress set (cp27,cp44,cp45)=(select cp27,cp44,cp45 from caseprogress where cp09='" & ArrCP09(i) & "') where cp09='" & strExc(1) & "' and cp27 is null"
                        cnnConnection.Execute strSql, intI
                        str249Msg = "本次發文含指定國註冊費且有指定UP，UP註冊已自動發文！"
                     End If
                     'end 2025/6/20
                  Else
                     If InStr(UPMember, Format(varTmp1(i))) > 0 Then
                        bolAdd250NP = True
                     End If
                  'end 2023/3/9
                  
                     '催審期限
                     strExc(3) = cp(7)
                     strExc(1) = PUB_Get224CtrlDate(1, strExc(3), cp)
                     PUB_UpdateChkResultDate strExc(1), arrCP, ArrCP09(i), "224"
                     
                     '提申期限
                     strExc(2) = PUB_Get224CtrlDate(2, strExc(3), cp)
                     strExc(1) = PUB_Get224CtrlDate(3, strExc(3), cp)
                     PUB_SetApplyDate cp(1), cp(2), cp(3), pa04, strExc(2), ArrCP09(i), "224", txtCaseField(0), Format(varTmp1(i)), strExc(1)
                  
                  End If
                  'end 2023/3/9
                  
                  '收達
                  PUB_SetArriveDate ArrCP09(i)
               Next
               
               'Added by Morgan 2023/3/9
               '管制 UPC選擇退出
               If bolAdd250NP Then
                  If DBDATE(txtCaseField(0)) <= 20300501 Then
                      '法限=2030/5/1
                      strExc(3) = "20300501"
                      '所限=2030/4/1
                      strExc(4) = "20300401"
                      If strExc(4) < strSrvDate(1) Then strExc(4) = strSrvDate(1)
                      strExc(4) = PUB_GetWorkDay1(strExc(4), True)
                      '智權人員
                      strExc(5) = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
                      '流水號
                      strExc(6) = GetNextProgressNo()
                     
                      strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                         "VALUES ('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','250'," & strExc(4) & "," & strExc(3) & ",'" & strExc(5) & "'," & strExc(6) & ") "
                      cnnConnection.Execute strSql, intI
                  End If
               End If
               'end 2023/3/9
            End If
            'end 2020/8/14
            
         End If
      'End If 'Removed by Morgan 2020/9/3
   End If
   
   'Add by Morgan 2007/8/16
   'Modify by Morgan 2010/3/17 +PA74
   If Text5(0) <> "" Then
      strExc(1) = varFeeYears(LBound(varFeeYears))
      strExc(2) = strSrvDate(1)
      strExc(3) = ""
      If Text5(1).Visible = True And Text5(1) <> "" Then
         If Val(Text5(1)) > Val(Text5(0)) Then
            For intI = LBound(varFeeYears) + 1 To UBound(varFeeYears)
               strExc(1) = strExc(1) & "," & varFeeYears(intI)
               strExc(2) = strExc(2) & "," & strSrvDate(1)
               strExc(3) = strExc(3) & ","
               If Val(varFeeYears(intI)) = Val(Text5(1)) Then
                  Exit For
               End If
            Next
         End If
      End If
      
      '更新基本檔
      strSql = "update patent set pa72='" & strExc(1) & "',pa73='" & strExc(2) & "',pa74='" & strExc(3) & "'" & _
         " where pa01='" & cp(1) & "' and pa02='" & cp(2) & "' and pa03='" & cp(3) & "' and pa04='" & cp(4) & "'"
      cnnConnection.Execute strSql, intI
      
      '新增下一程序檔
      If Text5(2) <> "" Then
         strExc(1) = field(1)
         strExc(2) = field(9)
         strExc(3) = DBDATE(Text5(2))
         GetCtrlDT strExc
         strExc(4) = PUB_GetWorkDay1(strExc(0), True)
         strExc(5) = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
         strExc(6) = GetNextProgressNo()
        
         strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
            "VALUES ('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','" & m_FeeProperty & "'," & strExc(4) & "," & strExc(3) & ",'" & strExc(5) & "'," & strExc(6) & ") "
      
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2007/8/16
   
'Removed by Morgan 2020/12/10 取消,印度為自動發證不會有領證發文
'   'Added by Lydia 2017/06/01 印度催商業使用聲明,自動產生繳費期間內的商業使用聲明(930)期限
'   If field(9) = "040" And Trim(Text5(0)) <> "" Then
'      If Trim(Text5(1)) = "" Or Trim(Text5(0)) = Trim(Text5(1)) Then
'         strExc(0) = "1"
'      Else
'         strExc(0) = Val(Text5(1)) - Val(Text5(0)) + 1
'      End If
'      'Added by Lydia 2017/11/02 抓專利起用日期來計算商業使用聲明期限的年度
'      If GetMoneyDate(Val(field(8)), field(9), field, str930Date, strExc(1), , "605") Then
'         If Trim(Text5(0)) > "1" Then
'            str930Date = CompDate(0, Val(Text5(0)) - 1, str930Date)
'         End If
'      End If
'      'end 2017/11/02
'      If str930Date <> "" Then 'Added by Lydia 2017/11/02
'        For intI = 1 To Val(strExc(0))
'            If intI = 1 Then
'               'Modified by Lydia 2017/11/02 從要繳年度起算
'               'strExc(1) = Mid(CompDate(0, 1, strSrvDate(1)), 1, 4) & "0331" '法限=明年3/31
'               strExc(1) = Mid(CompDate(0, 1, str930Date), 1, 4) & "0331" '法限=明年3/31
'            Else
'               strExc(1) = CompDate(0, 1, strExc(1))
'            End If
'            If strExc(1) > strSrvDate(1) Then 'Added by Lydia 2017/11/02 若計算結果的法定期限<系統日則該期限不產生
'                'Added by Lydia 2018/06/05 判斷下一程序期限是否存在
'                strSql = "select np01,np22,np06,np07 from nextprogress where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' " & _
'                            "and np07='930' and np06 is null and np09=" & CNULL(strExc(1), True)
'                intI = 1
'                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                If intI = 0 Then
'                'end 2018/06/05
'                    strExc(2) = PUB_GetWorkDay1(Mid(strExc(1), 1, 4) & "0131", True) '所限=明年1/31(抓工作天)
'                    strExc(6) = GetNextProgressNo()
'
'                    strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                         "VALUES ('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','930'," & strExc(2) & "," & strExc(1) & ",'" & PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4)) & "'," & strExc(6) & ") "
'                    cnnConnection.Execute strSql
'                End If 'end 2018/06/05
'            End If 'end 2017/11/02
'        Next intI
'      End If 'end 2017/11/02
'   End If
'   'end 2017/06/01
'end 2020/12/10
   
   '若案件國家收費表存在代理人收達天數則新增一筆收達的下一程序檔
   'Modify by Morgan 2015/8/7 發文收達期限管控改呼叫公用函式
   PUB_SetArriveDate cp(9)
   'end 2015/8/7
   
   'Added by Morgan 2015/8/7
   '提申管制
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
   '沒有客戶函
   'end 2018/8/22
   
   'Added by Morgan 2019/6/17
   '客戶函
   If txtCaseField(9) <> "N" Then
      strLetterJudge = PUB_GetLetterJudgeNew("1", field(1), cp(10), field(9))
      PUB_AddLetterProgress cp(9), 0, True, strLetterJudge, False, field(26), cp(10), field(75)
      m_strLD18 = cp(9)
   End If
   'end 2019/6/17
   
   'Added by Morgan 2023/6/13
   If field(9) = "023" Then
      strSql = "update caseprogress set cp27=" & DBDATE(cp(27)) & ",cp44='" & cp(44) & "',cp116='" & cp(116) & "' where cp01='" & field(1) & "' and cp02='" & field(2) & "' and cp03='" & field(3) & "' and cp04='" & field(4) & "' and cp10='443' and cp158=0 and cp159=0"
      cnnConnection.Execute strSql, intI
   End If
   'end 2023/6/13
   
   cnnConnection.CommitTrans
   SaveDatabase = True
   Exit Function
   
CheckingErr:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Sub ReadAllData()
   Dim rt As Boolean, i As Integer, varSaveCursor, strTemp As String, strTemp1 As String, j As Integer
   Dim adoRecord As Object, strSameName As String
   Dim stFeeYears As String '繳費年度
   Dim stFeeType  As String '年費起算日種類

On Error GoTo HndErr
varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
'Modify by Morgan 2006/10/19 改不Call Dll
'If objPublicData.ReadAllData(frm050102_1.grdDataList.TextMatrix(frm050102_1.grdDataList.Row, 5), cp(), field(), intCaseKind, intPWhere) Then
ReDim cp(TF_CP) As String
cp(9) = frm050102_1.grdDataList.TextMatrix(frm050102_1.grdDataList.row, 5)
If PUB_ReadAllData(cp(), field(), intCaseKind, intPWhere) Then
'end 2006/10/19

   Call ClsPDGetNationTax(Val(field(8)), field(9), , , , , , m_strReduceOne) 'Added by Morgan 2019/12/11

   lblCaseField(0) = cp(9)
   lblCaseField(1) = cp(1) + " - " + cp(2) + _
      IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
      IIf(cp(4) = "00", "", " - " + cp(4))
   lblCaseField(2) = TransDate(cp(6), 1)
   lblCaseField(4) = cp(13)
   lblCaseField(5) = TransDate(cp(7), 1)
   lblCaseField(3) = field(8)
   '2005/7/8 MODIFY BY SONIA
   'If field(76) = "" Then
   '   txtCaseField(3) = field(26)
   'Else
   '   txtCaseField(3) = field(76)
   'End If
   txtCaseField(3) = field(76)
   '2005/7/8 END
   CheckKeyIn 3
   txtCaseField(6) = cp(64)
   'Modify By Cheng 2002/08/19
'   If objPublicData.GetCasePreAgent(cp(), strTemp) Then
'      txtCaseField(1) = strTemp
'      CheckKeyIn 1
'   End If
   
   '92.1.12 add by sonia
   'Modify by Morgan 2006/9/21 加法國
   'Modified by Morgan 2023/3/25
   'If field(9) = "101" Or field(9) = "102" Or field(9) = "203" Then
   '   'Added by Morgan 2013/3/20
   '   If field(9) = "101" Then
   '      optChoose(2).Enabled = True
   '   Else
   '      optChoose(2).Enabled = False
   '   End If
   '   'end 2013/3/20
   PUB_SetEntityOpt field(1), field(9), field(8), optChoose
   If InStr(CFP_ChkEntity, field(9)) > 0 Or field(179) <> "" Then
      If strSrvDate(1) >= PA179啟用日 Then
         If field(179) = "1" Then
            optChoose(0).Value = True
            old_Entity = optChoose(0).Caption
         ElseIf field(179) = "2" Then
            optChoose(1).Value = True
            old_Entity = optChoose(1).Caption
         ElseIf field(179) = "3" Then
            optChoose(2).Value = True
            old_Entity = optChoose(2).Caption
         Else
            old_Entity = ""
         End If
      Else
   'end 2023/3/25
   
         If InStr(1, field(91), "大個體", 1) > 0 Then
            optChoose(0).Value = True
            old_Entity = "大個體"
         ElseIf InStr(1, field(91), "小個體", 1) > 0 Then
            optChoose(1).Value = True
            old_Entity = "小個體"
         'Added by Morgan 2013/3/20
         ElseIf InStr(1, field(91), "微個體", 1) > 0 Then
            optChoose(2).Value = True
            old_Entity = "微個體"
         'end 2013/3/20
         Else
            old_Entity = ""
         End If
         
      End If 'Added by Morgan 2023/3/25
   
      'Add by Morgan 2005/1/4 美國領證發文預設不先付款
      If field(9) = "101" Then txtCaseField(4) = "N"
      
      'Added by Morgan 2024/12/10 個體別順序會因國家有所不同,且客戶設定目前只設定是否可減免,原預設規則只適用於1,2選項為大小個體時
      If optChoose(0).Caption = "大個體" And optChoose(1).Caption = "小個體" Then
      'end 2024/12/10
         
         'Add by Morgan 2004/9/24
         'Modified by Morgan 2013/3/20
         'If optChoose(0).Value = False And optChoose(1).Value = False Then
         If optChoose(0).Value = False And optChoose(1).Value = False And optChoose(2).Value = False Then
            Dim stAD03 As String
            For i = 1 To 5
               If field(25 + i) <> "" Then
                  stAD03 = PUB_GetAD03(field(25 + i), field(9))
                  If stAD03 = "N" Then
                     optChoose(0).Value = True
                     Exit For
                  '只要有未設定減免身分的公司申請人則不預設大小個體
                  ElseIf stAD03 = "" Then
                     Exit For
                  End If
               End If
            Next
            '若五個申請人檢查完都不是大個體則為小個體
            If optChoose(2).Enabled = False Then 'Added by Morgan 2013/3/20 不可選微個體時才預設
               If optChoose(0).Value = False And i = 6 Then optChoose(1).Value = True
            End If
         
         End If
         
       End If 'Added by Morgan 2024/12/10
       
   End If
   '92.1.12 end
   
   Set adoRecord = CreateObject("ADODB.Recordset")
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.SelectTable("select cp44 from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "'", adoRecord) Then
   '2007/4/23 MODIFY BY SONIA 加發文日降冪排序
   'If ClsPDSelectTable("select cp44 from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "'", adoRecord) Then
   'Modify by Morgan 2008/2/21 加聯絡人
   'Added by Lydia 2016/10/27 +新案有申請人指定國外代理人檔則預設
   If cp(31) = "Y" Then
      AddAgent Combo1, cp, , , , cp(9), field(9), field(26)
      If Combo1 <> "" Then CheckKeyIn 1
      
   Else '非新案照原本
        If ClsPDSelectTable("select cp44||decode(cp116,null,null,'-'||cp116) from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "' and cp09<'C' and cp44 is not null order by cp27 desc", adoRecord) Then
        '2007/4/23 END
           Do While adoRecord.EOF = False
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
                
        'edit by nickc 2007/02/02 不用 dll 了
        'If objPublicData.GetCasePreAgent(cp(), strTemp) Then
        If ClsPDGetCasePreAgent(cp(), strTemp) Then
           Combo1 = strTemp
           CheckKeyIn 1
        End If
        
      End If 'Added by Morgan 2023/10/30
   End If
   'end 2016/10/27
   
   If field(9) <> EPC指定國家 Then
      cmdCountry.Enabled = False
   End If
   If cmdCountry.Enabled Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.ReadCountry(intCaseKind, cp(), strCountry, True) = False Then GoTo Err
      If ClsPDReadCountry(intCaseKind, cp(), strCountry, True) = False Then GoTo HndErr
      If strCountry = "" Then
         MsgBox "所有子案已閉卷 !", vbCritical
         bolLeave = True
         intLeaveKind = 1
         Unload Me
      End If
   Else
      strCountry = ""
   End If
   txtCaseField(7) = "Y"
   
   'Add by Morgan 2008/12/17
   PUB_GetNextYearFeeDate cp, field, , , m_StartDate
   
   'Add by Morgan 2007/8/27 准後繳且不是自動發證的國家可輸年費
   strExc(0) = "select decode('" & field(8) & "','1',NA06,'2',NA08,NA10) C1" & _
      ",decode('" & field(8) & "','1',NA20,'2',NA22,NA24) C2" & _
      ",decode('" & field(8) & "','1',NA21,'2',NA23,NA25) C3" & _
      " from nation where na01='" & field(9) & "'" & _
      " and decode('" & field(8) & "','1',NA56,'2',NA57,NA58)='Y'" & _
      " and decode('" & field(8) & "','1',NA49,'2',NA53,NA54) is null"
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      stFeeType = "" & RsTemp.Fields("C1")
      
      'Remove by Morgan 2008/12/17 移到上面改呼叫共用(配合 CFP-15830 案暫先改本程式,其他程式將來有時間或遇到再修正)
      ''其他日期不用考慮
      'Select Case stFeeType
      '   Case "2" '申請日
      '      m_StartDate = field(10)
      '   Case "4" '准駁日
      '      m_StartDate = field(20)
      '   Case "7" '公開日
      '      m_StartDate = field(12)
      '   Case Else
      '      m_StartDate = ""
      'End Select
      'end 2008/12/17
      
      If m_StartDate <> "" Then
         m_FeeProperty = RsTemp.Fields("C2")
         stFeeYears = "" & RsTemp.Fields("C3")
         varFeeYears = Split(stFeeYears, ",")
         Text5(0).Enabled = True
         If m_FeeProperty = "605" Then
            Text5(1).Enabled = True
         Else
            Text5(1).Visible = False
            Label11(4) = "第                    次" & GetCaseTypeName("P", m_FeeProperty, 1)
         End If
         '下次繳費日
         If field(9) = "013" Then
            Text5(2).Enabled = True
         End If
      End If
   End If
   'end 2007/8/27
   
   'Add by Morgan 2009/8/18
   If txtCaseField(0).Tag <> txtCaseField(0) Then
      PUB_SetChkResultDate cp(1), field(9), cp(10), txtCaseField(0), txtChkRltDate, cp, field(8)
      txtCaseField(0).Tag = txtCaseField(0)
   End If
   
   'Add by Morgan 2010/2/2 美國案檢查是否有收文提早公開
   If field(9) = "101" And field(8) = "1" Then
      If PUB_ChkCPExist(field, "417") Then
         txtCaseField(8).Enabled = False
      End If
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
HndErr:
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
      
      'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
      If Cancel = False Then
         strNo = Combo1.Text
         'Add by Morgan 2008/2/21 加聯絡人判斷
         iPos = InStr(Combo1.Text, "-")
         If iPos > 0 Then
            strNo = Left(Combo1.Text, iPos - 1)
         End If
         'end 2008/2/21
         
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
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , 台灣國家代號) = 1 Then
      If ClsPDGetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , 台灣國家代號) = 1 Then
         lblTrademarkKind = strTemp
      End If
   Case 4
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetStaffN(lblCaseField(Index), strTemp) Then
      If ClsPDGetStaffN(lblCaseField(Index), strTemp) Then
         lblSalesName = strTemp
      Else
         lblSalesName = ""
      End If
End Select
End Sub
Private Sub Form_Activate()
   Dim bCancel As Boolean
   '若表單第一次顯示
   If m_blnFormFirstShow = True Then
       m_blnFormFirstShow = False
       txtCaseField(0) = strSrvDate(2)
       ReadAllData
       txtCaseField(0).SetFocus
       
      'Add by Morgan 2008/1/10
      '墨西哥領證需同時繳納"繳納領證費"當年度起算5年之年費
      '因代理人繳納日期無法預知故以發文日估算待代理人通知後入有跨年問題再更新已繳費年度及下次期限
      If field(9) = "104" And txtCaseField(0) <> "" And m_StartDate <> "" Then
         strExc(1) = DateDiff("yyyy", ChangeWStringToWDateString(DBDATE(m_StartDate)), ChangeWStringToWDateString(CompDate(0, 5, txtCaseField(0))))
         MsgBox "墨西哥領證需同時繳納""繳納領證費""當年度起算5年之年費，故繳費年將預設為 1-" & strExc(1) & " 年。"
         Text5(0) = "1"
         Text5(1) = strExc(1)
         Text5_Validate 1, bCancel
      End If
      If PUB_ChkFileNP(cp(9)) Then MsgBox "下一程序已有提申或收達期限，若為重新發文時需要先刪除後才可作業！" 'Added by Morgan 2015/8/7
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   strLicenceCountry = ""
   m_blnFormFirstShow = True
   
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
   
   'Set frm050102_8 = Nothing'Removed by Morgan 2021/12/10 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub


Private Sub Text5_GotFocus(Index As Integer)
   TextInverse Text5(Index)
End Sub

Private Sub Text5_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 Then
      If Text5(0) <> "" Then
         If Val(Text5(0)) <> varFeeYears(0) Then
            MsgBox "請輸入正確的繳費年度 !", vbCritical
            Text5_GotFocus Index
            Cancel = True
         '延展費,維持費
         ElseIf Text5(1).Visible = False Then
            Text5(2) = ""
            If Val(Text5(0)) < UBound(varFeeYears) Then
               'Modified by Morgan 2019/12/11
               'Text5(2) = CompDate(0, varFeeYears(Val(Text5(0)) + 1), m_StartDate)
               Text5(2) = PUB_GetEndDate(m_StartDate, varFeeYears(Val(Text5(0)) + 1), m_strReduceOne, field(9))
               'end 2019/12/11
            End If
            If field(25) <> "" And Val(TransDate(Text5(2), 2)) >= Val(field(25)) Then
               Text5(2) = ""
            End If
         End If
      End If
      
   ElseIf Index = 1 Then
      If Text5(1) <> "" Then
         If Val(Text5(0)) > Val(Text5(1)) Then
            MsgBox "繳費年度錯誤，請重新輸入 !", vbCritical
            Text5_GotFocus Index
            Cancel = True
         'Modify by Morgan 2010/2/24
         'ElseIf Val(Text5(1)) > varFeeYears(UBound(varFeeYears)) Then
         ElseIf InStr("," & Join(varFeeYears, ",") & ",", "," & Text5(1) & ",") = 0 Then
            MsgBox "繳費年度錯誤，請重新輸入 !", vbCritical
            Text5_GotFocus Index
            Cancel = True
         '年費
         Else
            Text5(2) = ""
            If Val(Text5(1)) < varFeeYears(UBound(varFeeYears)) Then
               'Modify by Morgan 2010/2/24 期限改用下次繳費年減1年計算 CFP-020862
               'Text5(2) = CompDate(0, Val(Text5(1)), m_StartDate)
               strExc(1) = ""
               For intI = LBound(varFeeYears) To UBound(varFeeYears)
                  If Val(varFeeYears(intI)) > Val(Text5(1)) Then
                     strExc(1) = Val(varFeeYears(intI)) - 1
                     Exit For
                  End If
               Next
               If strExc(1) <> "" Then
                  'Modified by Morgan 2019/12/11
                  'Text5(2) = CompDate(0, Val(strExc(1)), m_StartDate)
                  Text5(2) = PUB_GetEndDate(m_StartDate, Val(strExc(1)), m_strReduceOne, field(9))
                  'end 2019/12/11
               End If
               'end 2010/2/24
               
               '2008/3/21 ADD BY SONIA沙烏地阿拉伯,年費期限為每年3/31
               If field(9) = "021" And Text5(2) <> "" Then
                  Text5(2) = ChangeWStringToTString(Mid(ChangeTStringToWString(Text5(2)), 1, 4) + "0331")
               End If
               '2008/3/21 END
            End If
            If field(25) <> "" And Val(TransDate(Text5(2), 2)) >= Val(field(25)) Then
               Text5(2) = ""
            End If
         End If
      End If
   End If
End Sub

Private Sub txtCaseField_Change(Index As Integer)
   Select Case Index
      Case 3
         lblNotify = ""
   End Select
End Sub
Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
Select Case Index
             Case 1, 2, 3, 4, 5
                       KeyAscii = UpperCase(KeyAscii)
            'Add By Cheng 2002/07/31
            Case 7, 8
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

'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
If Cancel = False And Index = 3 Then
   If PUB_CheckStatus(txtCaseField(Index).Text) = False Then Cancel = True
End If

If Cancel Then txtCaseField_GotFocus (Index)
End Sub
Private Function CheckKeyIn(intIndex As Integer) As Integer
   Dim strTemp As String, strCusTemp As String

   CheckKeyIn = -1
   Select Case intIndex
      Case 0
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
            'Add by Morgan 2008/2/21 加判斷是否為聯絡人
            If InStr(strCusTemp, "-") > 0 Then
               If ClsPDGetContact(strCusTemp, strTemp) Then
                  Combo1 = strCusTemp
                  lblAgent.Caption = strTemp
                  CheckKeyIn = 1
               End If
            
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetAgent(strCusTemp, strTemp) Then
            ElseIf ClsPDGetAgent(strCusTemp, strTemp) Then
               Combo1 = strCusTemp
               lblAgent.Caption = strTemp
               CheckKeyIn = 1
            End If
         End If
         
      Case 2
         If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
            CheckKeyIn = 1
         Else
            ShowMsg MsgText(1038)
         End If
             
      Case 3
         '2005/7/8 MODIFY BY SONIA 加判斷有值才做
         If txtCaseField(intIndex) <> "" Then
            strCusTemp = txtCaseField(intIndex)
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetCustomer(strCusTemp, strTemp) Then
            If ClsPDGetCustomer(strCusTemp, strTemp) Then
               txtCaseField(intIndex) = strCusTemp
               lblNotify = strTemp
               CheckKeyIn = 1
            End If
         Else
            CheckKeyIn = 1
         End If
         
      Case 4
         If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Or txtCaseField(intIndex) = "Y" Then
            If txtCaseField(intIndex) = "" And field(9) = 美國國家代號 And txtCaseField(2) = "" Then
               MsgBox "因為為美國案且列印指示信，所以美國是否需先付款不可空白 !", vbCritical
            Else
               CheckKeyIn = 1
            End If
         Else
            ShowMsg MsgText(9177)
         End If
         
      Case 5
         If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Or txtCaseField(intIndex) = "Y" Then
            If txtCaseField(intIndex) = "" And field(9) = 美國國家代號 And txtCaseField(2) = "" Then
               MsgBox "美國案且列印指示信，請輸入美國是否需附圖 !", vbCritical
            Else
               CheckKeyIn = 1
            End If
         Else
            ShowMsg MsgText(9177)
         End If
         
      Case Else
         CheckKeyIn = 1
   End Select
End Function
Private Sub txtCaseField_GotFocus(Index As Integer)
txtCaseField(Index).SelStart = 0
txtCaseField(Index).SelLength = Len(txtCaseField(Index).Text)
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

   'Added by Morgan 2021/12/6 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   
Cancel = False
   'add by nickc 2008/05/01
   If IsDebt(field(9), cp(9)) Then
        MsgBox "未收款且無 預定收款日 請轉告智權同仁！！", vbOKOnly, "警告！禁止發文！"
        Exit Function
   End If
For Each objTxt In Me.txtCaseField
   If objTxt.Enabled = True Then
      txtCaseField_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

'Add by Morgan 2004/9/14
If Combo1.Enabled = True Then
   Combo1_Validate Cancel
   If Cancel = True Then
      Combo1.SetFocus
      Exit Function
   End If
End If

'Add by Morgan 2007/8/16
If Text5(0).Enabled = True Then
   Text5_Validate 0, Cancel
   If Cancel = True Then
      Text5_GotFocus 0
      Text5(0).SetFocus
      Exit Function
   End If
End If
If Text5(1).Enabled = True Then
   If Text5(0) <> "" And Text5(1) = "" Then
      MsgBox "迄年輸入錯誤！", vbExclamation
      Text5_GotFocus 1
      Text5(1).SetFocus
      Exit Function
   End If

   Text5_Validate 1, Cancel
   If Cancel = True Then
      Text5_GotFocus 1
      Text5(1).SetFocus
      Exit Function
   End If
End If
'end 2007/8/16

'Added by Morgan 2018/9/12 CFP電子化-接洽單檢查
If strSrvDate(1) >= CFP第一階段電子化啟用日 Then
   If cp(9) < "B" And Left(cp(12), 1) <> "F" Then
      If PUB_CheckPDF3(cp(1), cp(2), cp(3), cp(4)) = False Then
         Exit Function
      End If
   End If
End If
'end 2018/9/12

'Added by Morgan 2019/4/30
'日本發明案領證發文要設定減免身分
m_strCP81 = ""
m_strJpMemo = ""
If field(9) = "011" And field(8) = "1" And cp(10) = "601" Then
   If PUB_ChkJpDiscount(cp(1), cp(2), cp(3), cp(4), True) = True Then
      Dim stAD10 As String, stAD15 As String
      For ii = 1 To 5
         If field(25 + ii) <> "" Then
            strExc(1) = PUB_GetAD03(field(25 + ii), "011", stAD10, , stAD15)
            m_strJpMemo = m_strJpMemo & PUB_GetJpDiscountDesc(stAD10, stAD15) & ";"
            If strExc(1) = "" Then
               'Modified by Morgan 2019/6/19 改詢問是否不可減免,若是則系統自動設定--禧佩
               'MsgBox "申請人【" & field(25 + ii) & " " & GetCustomerNameAndState(field(25 + ii)) & "】尚未設定減免身分不可發文！", vbCritical, "日本領證發文減免身分檢查"
               'Exit Function
               If MsgBox("申請人【" & field(25 + ii) & " " & GetCustomerNameAndState(field(25 + ii)) & "】尚未設定減免身分！" & vbCrLf & vbCrLf & "是否要設定為【不可減免】？", vbYesNo + vbDefaultButton2 + vbExclamation, "日本實審發文減免身分檢查") = vbYes Then
                  PUB_SetNoDisc field(25 + ii), field(9)
                  m_strCP81 = "N"
               Else
                  Exit Function
               End If
               'end 2019/6/19
            ElseIf m_strCP81 <> "N" Then
               m_strCP81 = strExc(1)
            End If
         End If
      Next
      If m_strCP81 <> "Y" Then m_strJpMemo = ""
   End If
End If
'end 2019/4/30

'Added by Lydia 2021/05/25 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
If Pub_ChkACS112isNull(field(1), field(2), field(3), field(4), txtCP113) = True Then
    txtCP113.SetFocus
    txtCP113_GotFocus
    Exit Function
End If
'end 2021/05/25

TxtValidate = True
End Function
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
