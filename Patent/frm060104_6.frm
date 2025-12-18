VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060104_6 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專發文-授權, 終止授權"
   ClientHeight    =   5736
   ClientLeft      =   -1068
   ClientTop       =   1008
   ClientWidth     =   8724
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   8724
   Begin VB.TextBox txtRecDate 
      Height          =   270
      Left            =   1140
      MaxLength       =   1
      TabIndex        =   12
      Top             =   5370
      Width           =   375
   End
   Begin VB.TextBox txtEmail 
      Height          =   270
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   13
      Text            =   "Y"
      Top             =   5370
      Width           =   375
   End
   Begin VB.TextBox txtPayToday 
      Height          =   264
      Left            =   7815
      MaxLength       =   1
      TabIndex        =   5
      Top             =   2943
      Width           =   255
   End
   Begin VB.TextBox txtCP118 
      Height          =   270
      Left            =   4800
      MaxLength       =   1
      TabIndex        =   4
      Top             =   2940
      Width           =   375
   End
   Begin VB.TextBox Text12 
      Height          =   270
      Index           =   0
      Left            =   1140
      MaxLength       =   7
      TabIndex        =   7
      Top             =   4005
      Width           =   1335
   End
   Begin VB.TextBox Text12 
      Height          =   270
      Index           =   1
      Left            =   2820
      MaxLength       =   7
      TabIndex        =   8
      Top             =   4005
      Width           =   1335
   End
   Begin VB.TextBox txtCP84 
      Height          =   288
      Left            =   7380
      TabIndex        =   2
      Top             =   2631
      Width           =   1092
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   4425
      MaxLength       =   7
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2580
      MaxLength       =   2
      TabIndex        =   33
      Top             =   1020
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2340
      MaxLength       =   1
      TabIndex        =   32
      Top             =   1020
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   31
      Top             =   1020
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1020
      MaxLength       =   3
      TabIndex        =   30
      Top             =   1020
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      ItemData        =   "frm060104_6.frx":0000
      Left            =   1020
      List            =   "frm060104_6.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   15
      Top             =   2220
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   400
      Index           =   3
      Left            =   4452
      TabIndex        =   19
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6504
      TabIndex        =   20
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5676
      TabIndex        =   14
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7728
      TabIndex        =   21
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox Text11 
      Height          =   270
      Left            =   1140
      MaxLength       =   9
      TabIndex        =   10
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   1140
      MaxLength       =   9
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   1140
      MaxLength       =   9
      TabIndex        =   3
      Top             =   2940
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   1140
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblRecDate 
      AutoSize        =   -1  'True
      Caption         =   "當天報告:             (Y:是)"
      Height          =   180
      Left            =   270
      TabIndex        =   63
      Top             =   5415
      Width           =   1815
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "Email維護:             (Y:是)"
      Height          =   180
      Left            =   3510
      TabIndex        =   62
      Top             =   5415
      Width           =   1860
   End
   Begin MSForms.TextBox Text8 
      Height          =   285
      Index           =   2
      Left            =   3060
      TabIndex        =   18
      Top             =   3720
      Width           =   5415
      VariousPropertyBits=   679493661
      BackColor       =   -2147483648
      MaxLength       =   60
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text8 
      Height          =   285
      Index           =   1
      Left            =   3060
      TabIndex        =   17
      Top             =   3480
      Width           =   5415
      VariousPropertyBits=   679493661
      BackColor       =   -2147483648
      MaxLength       =   60
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text14 
      Height          =   705
      Left            =   1140
      TabIndex        =   11
      Top             =   4620
      Width           =   5985
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "10557;1244"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text8 
      Height          =   285
      Index           =   0
      Left            =   3060
      TabIndex        =   16
      Top             =   3240
      Width           =   5415
      VariousPropertyBits=   679493661
      BackColor       =   -2147483648
      MaxLength       =   60
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   7170
      TabIndex        =   9
      Top             =   4005
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
   Begin VB.Label lblPayToday 
      AutoSize        =   -1  'True
      Caption         =   "電子送件是否當日扣款:         (Y/N)"
      Height          =   180
      Left            =   5880
      TabIndex        =   61
      Top             =   2985
      Width           =   2655
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "是否電子送件:          (Y: 是)"
      Height          =   180
      Left            =   3645
      TabIndex        =   60
      Top             =   2985
      Width           =   2085
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   6240
      TabIndex        =   59
      Top             =   4050
      Width           =   900
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "授權期間:"
      Height          =   180
      Left            =   180
      TabIndex        =   58
      Top             =   4050
      Width           =   765
   End
   Begin VB.Label Label28 
      Caption         =   "~"
      Height          =   255
      Left            =   2580
      TabIndex        =   57
      Top             =   4020
      Width           =   135
   End
   Begin VB.Label lblCP84 
      AutoSize        =   -1  'True
      Caption         =   "發文規費:"
      Height          =   180
      Left            =   6525
      TabIndex        =   56
      Top             =   2685
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "催審期限:"
      Height          =   180
      Left            =   3435
      TabIndex        =   55
      Top             =   2685
      Width           =   765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   180
      X2              =   8650
      Y1              =   2580
      Y2              =   2580
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   180
      X2              =   8650
      Y1              =   2640
      Y2              =   2640
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   10
      Left            =   2550
      TabIndex        =   54
      Top             =   4365
      Width           =   4500
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "7937;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   9
      Left            =   2550
      TabIndex        =   53
      Top             =   2985
      Width           =   1050
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "1852;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   180
      TabIndex        =   52
      Top             =   2220
      Width           =   768
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "申請人5:"
      Height          =   180
      Left            =   180
      TabIndex        =   51
      Top             =   1920
      Width           =   672
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "申請人4:"
      Height          =   180
      Left            =   4020
      TabIndex        =   50
      Top             =   1620
      Width           =   672
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "申請人3:"
      Height          =   180
      Left            =   180
      TabIndex        =   49
      Top             =   1620
      Width           =   672
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "申請人2:"
      Height          =   180
      Left            =   4020
      TabIndex        =   48
      Top             =   1320
      Width           =   672
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請人1:"
      Height          =   180
      Left            =   180
      TabIndex        =   47
      Top             =   1320
      Width           =   672
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Left            =   4020
      TabIndex        =   46
      Top             =   1020
      Width           =   768
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   180
      TabIndex        =   45
      Top             =   1020
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   4020
      TabIndex        =   44
      Top             =   720
      Width           =   768
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Left            =   180
      TabIndex        =   43
      Top             =   720
      Width           =   588
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   0
      Left            =   1020
      TabIndex        =   42
      Top             =   720
      Width           =   2460
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4339;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   1
      Left            =   4860
      TabIndex        =   41
      Top             =   720
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5186;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   2
      Left            =   4860
      TabIndex        =   40
      Top             =   1020
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5186;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   3
      Left            =   1020
      TabIndex        =   39
      Top             =   1320
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5186;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   4
      Left            =   4860
      TabIndex        =   38
      Top             =   1320
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5186;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   5
      Left            =   1020
      TabIndex        =   37
      Top             =   1620
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5186;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   6
      Left            =   4860
      TabIndex        =   36
      Top             =   1620
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5186;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   7
      Left            =   1020
      TabIndex        =   35
      Top             =   1920
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5186;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   8
      Left            =   1680
      TabIndex        =   34
      Top             =   2250
      Width           =   6750
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "11906;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Left            =   180
      TabIndex        =   29
      Top             =   4650
      Width           =   765
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "年費代理人:"
      Height          =   180
      Left            =   180
      TabIndex        =   28
      Top             =   4365
      Width           =   945
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "日:"
      Height          =   180
      Left            =   2700
      TabIndex        =   27
      Top             =   3720
      Width           =   225
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "英:"
      Height          =   180
      Left            =   2700
      TabIndex        =   26
      Top             =   3480
      Width           =   225
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "中:"
      Height          =   180
      Left            =   2700
      TabIndex        =   25
      Top             =   3240
      Width           =   225
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "被授權人:"
      Height          =   180
      Left            =   180
      TabIndex        =   24
      Top             =   3285
      Width           =   765
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   180
      TabIndex        =   23
      Top             =   2985
      Width           =   585
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Left            =   180
      TabIndex        =   22
      Top             =   2685
      Width           =   585
   End
End
Attribute VB_Name = "frm060104_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/16 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strReceiveNo As String
'Modify by Morgan 2005/8/4 改用動態陣列
'Dim pa(1 To T_PA) As String, cp(1 To T_CP) As String
Dim pa() As String, cp() As String
Dim intWhere As Integer
' 案件性質
Dim m_CP10 As String
Dim m_PA143 As String 'Add by Morgan 2008/3/18
Dim m_CP09s As String, m_CP123s As String 'Add by Morgan 2009/3/20 收文號,是否算發文室案件
Dim m_CP130 As String 'Add by Morgan 2009/4/28 發文-主管機關
Dim m_AgentName As String 'Add By Sindy 2021/5/10


Private Sub cmdok_Click(Index As Integer)
'Added by Lydia 2018/09/11
Dim strFilePath As String '記錄智慧局收文文號
Dim strNewCP64 As String '保留進度備註
'end 2018/09/11

   Select Case Index
      'Modify by Morgan 2009/3/26 將同時發文併入
      'Case 0 '確定
      Case 0, 3
         'Add By Cheng 2002/07/03
         If TxtValidate = False Then Exit Sub
         
         'Add by Morgan 2008/3/18
         '若未辦或不辦重新委任時不可發文
         If PUB_Check928NotOk(pa) = True Then
            MsgBox "本案下一程序有重新委任之補文件未辦理，不可發文！"
            Exit Sub
         End If
         '若基本檔年費申請人是否出名為N時提醒存檔將取消
         m_PA143 = pa(143)
         If pa(143) = "N" Then
            MsgBox "年費申請人是否出名現為【N】，存檔時將自動取消！"
            m_PA143 = ""
         End If
         'end 2008/3/18
         
         'Added by Lydia 2018/09/11 是否電子送件
         strNewCP64 = Text14
         If txtCP118 = "Y" Then
            '電子送件也要記錄主管機關
            If ModifyDispatchCp130(strReceiveNo, m_CP09s, m_CP123s, m_CP130, Text9, , True) = False Then
               Exit Sub
            End If
            strExc(0) = InputBox("請輸入智慧局收文文號!!")
            If strExc(0) = "" Then
               Exit Sub
            Else
               strFilePath = strExc(0)  '記錄智慧局收文文號
               strNewCP64 = "智慧局收文文號:" & strExc(0) & ";" & Text14
            End If
         Else
         'end 2018/09/11
            'Add by Morgan 2009/4/28
            If ModifyDispatchCp130(strReceiveNo, m_CP09s, m_CP123s, m_CP130, Text9) = False Then
               Exit Sub
            End If
            If m_CP123s = "Y" Then
            'end 2009/4/28
               'Add by Morgan 2009/3/20 設定是否算發文室案件
               'modify by sonia 2014/6/23 加傳發文規費, P-108903
               If ModifyDispatch(strReceiveNo, m_CP09s, m_CP123s, txtCP84, Text9) = False Then
                   Exit Sub
               End If
               'end 2009/3/20
            End If
         End If 'end 2018/09/11

         'Added by Lydia 2018/09/11 依據輸入的智慧局收文號(受理號,ex: 1073066637-0)，將本機C:\E-SET\RdcDocDir\(收文號ex: 1073066637-0)的pdf檔自動搬移到卷宗區(by Phoebe);
         If txtCP118.Text = "Y" And strFilePath <> "" Then
             strExc(1) = cp(82)
             If Val(cp(82)) > 0 Then
                 If MsgBox("重新發文是否上傳檔案到卷宗區？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                      strExc(1) = ""
                 End If
             End If
             If Val(strExc(1)) = 0 Then
                'Modified by Lydia 2019/03/22 +傳入發文日
                If Pub_AutoEsetToCpp(True, pa(1), pa(2), pa(3), pa(4), pa(8), Label2(0).Caption, m_CP10, strFilePath, Text9.Text) = False Then
                      Exit Sub
                End If
             End If
         End If
         'end 2018/09/11
         
         'Added by Lydia 2018/09/11 檢查完畢，更新備註欄位
         Text14.Text = strNewCP64
         
         'Add by Sindy 2021/11/16 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If
         
         ' 設定滑鼠游標為等待狀態
         Screen.MousePointer = vbHourglass
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
         
            'Add by Morgan 2008/2/20 檢查代理人Email
            PUB_CheckEMail pa(75), pa(144)
            If pa(145) <> "" Then
               PUB_CheckEMail pa(75), pa(145)
            End If
            'end 2008/2/20
            
         ' 設定滑鼠游標為預設
         Screen.MousePointer = vbDefault
         
         If pa(1) = "FCP" Then
'Modified by Morgan 2020/3/3 改呼叫共用
'            'Add By Sindy 2016/7/7 + 代理人為Y4829203Hewlett-Packard Company Intellectual Property Administration
'            '承辦人為工程師(ST03 IN ('F21','F51','F52))時,於存檔後彈訊息
'            If ChangeCustomerL(pa(75)) = "Y48292030" And _
'               (PUB_GetST03(Text10.Text) = "F21" Or PUB_GetST03(Text10.Text) = "F51" Or PUB_GetST03(Text10.Text) = "F52") Then
'               'Add By Sindy 2016/7/18
'               strExc(0) = "select cp130 from caseprogress where cp09='" & strReceiveNo & "'"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  If "" & RsTemp.Fields(0) <> "" Then
'               '2016/7/18 END
'                     MsgBox "請優先請款並且在提申當天上傳報告!!"
'                  End If
'               End If
'            End If
'            '2016/7/7 END
'
'            'Add By Sindy 2016/10/17 凡代理人Y33844   KLARQUIST SPARKMAN, LLP的案件，
'            '若是工程師中間程序(例: 申復、再審、訴願、補充說明、...)發文時，
'            '彈訊息"請在送件後3天內並且要當月優先請款"，請排除901.告代、902.回代、1202.審查意見、1002.核駁.....。
'            If (PUB_GetST03(Text10.Text) = "F21" Or PUB_GetST03(Text10.Text) = "F51" Or PUB_GetST03(Text10.Text) = "F52") And _
'               Not (m_CP10 = "901" And m_CP10 = "902" And m_CP10 = "1202" And m_CP10 = "1002") Then
'               strExc(0) = "select cp130 from caseprogress where cp09='" & strReceiveNo & "'"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  If "" & RsTemp.Fields(0) <> "" Then
'                     If ChangeCustomerL(pa(75)) = "Y33844000" Then
'                        MsgBox "請在送件後3天內並且要當月優先請款!!"
'                     'Add By Sindy 2016/10/20
'                     ElseIf ChangeCustomerL(pa(75)) = "Y51982000" Then
'                        MsgBox "申復/再審/修正等送智慧局的案件，於收到指示後7天內送程序送件同時請款(由Wilson指示備註)!!"
'                     '2016/10/20 END
'                     'Add By Sindy 2016/10/24
'                     ElseIf ChangeCustomerL(pa(75)) = "Y20272000" Then
'                        MsgBox "中間程序送件當日簡單報告!!"
'                     '2016/10/24 END
'                     'Add By Sindy 2016/11/16
'                     ElseIf ChangeCustomerL(pa(75)) = "Y34440B30" Then
'                        MsgBox "請當日優先請款報告!!"
'                     '2016/11/16 END
'                     End If
'                  End If
'               End If
'            End If
'            '2016/10/17 END
            PUB_FCPAlert strReceiveNo
'end 2020/3/3
         End If
         
         'Add By Cheng 2002/04/30
         If Index = 0 Then
            'Add By Sindy 2023/11/9
            If frm060104_1.bolIsEMPFlow = True Then
               frm090202_4.QueryData
            End If
            '2023/11/9 End
            '若有未發文資料顯示警告
            'Modify By Sindy 2023/11/9
            If PUB_GetCPunIssueDatas("" & Me.Text1.Text & "-" & Me.Text2.Text & "-" & IIf(Len("" & Me.Text3.Text) <= 0, "0", Me.Text3.Text) & "-" & IIf(Len("" & Me.Text4.Text) <= 0, "00", Me.Text4.Text)) Then
               frm060104_1.Show
               frm060104_1.ReQuery
            Else
               'Add By Sindy 2023/11/9
               If frm060104_1.bolIsEMPFlow = True Then
                  Unload frm060104_1
               Else
               '2023/11/9 End
                  frm060104_1.Show
                  frm060104_1.Clear
               End If
            End If
         Else
            frm060104_1.Show
            frm060104_1.ReQuery
         End If
         
         'Add By Sindy 2022/5/12
         If txtEmail.Text = "Y" Then
            frm060104_k.m_CP09 = strReceiveNo 'cp(9)
            frm060104_k.m_strRecDate = txtRecDate
            frm060104_k.Hide
            frm060104_k.cmdOK(0) = 1
            Unload frm060104_k
         End If
         '2022/5/12 END
         
         Unload Me
      Case 1
         frm060104_1.Show
         Unload Me
      Case 2
         Unload frm060104_1
         Unload Me
'Remove by Morgan 2009/3/26
'      Case 3 '同時發文
'         'Add By Cheng 2002/07/03
'         If TxtValidate = False Then Exit Sub
'         'Add by Morgan 2009/3/20 設定是否算發文室案件
'         If ModifyDispatch(strReceiveNo, m_CP09s, m_CP123s, Text9) = False Then
'             Exit Sub
'         End If
'         'end 2009/3/20
'
'         ' 設定滑鼠游標為等待狀態
'         Screen.MousePointer = vbHourglass
'         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
'         ' 設定滑鼠游標為預設
'         Screen.MousePointer = vbDefault
'
'         'frm060104_1.Text1 = pa(1)
'         'frm060104_1.Text2 = pa(2)
'         'frm060104_1.Text3 = pa(3)
'         'frm060104_1.Text4 = pa(4)
'         'frm060104_1.Command1_Click
'         frm060104_1.Show
'         ' 90.08.06 modify by louis
'         frm060104_1.ReQuery
'         Unload Me
   End Select
End Sub

Private Function FormSave() As Boolean
   'Add By Cheng 2002/07/03
   Dim ii As Integer
   Dim intMax As Long
   Dim stCP118 As String, stCP152 As String 'Added by Lydia 2018/09/11
   
'911105 nick transation
FormSave = True
 On Error GoTo CheckingErr
 cnnConnection.BeginTrans
    'Modify by Morgan 2004/3/16
    '代理人補滿九碼
    'strExc(1) = "UPDATE PATENT SET PA76=" & CNULL(Text11) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
    'Modify by Morgan 2008/3/18 +pa143
   strExc(1) = "UPDATE PATENT SET PA76=" & CNULL(ChangeCustomerL(Text11)) & ",PA143='" & m_PA143 & "' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   
 '911105 nick transation
 cnnConnection.Execute strExc(1), intI
 
   ii = 1
   
   'Added by Lydia 2018/09/11
   '電子送件有規費的一律設自動扣款(同內專) --敏莉
   stCP118 = txtCP118
   stCP152 = ""
   If txtCP118 = "Y" And Val(txtCP84) > 0 Then
      stCP118 = "A"
      stCP152 = Pub_FcpSetPayToday("2", Text9.Text, txtPayToday.Text)
   End If
   'end 2018/09/11
   
   If m_CP10 = "705" Then
      'Modify by morgan 2005/8/4 加 cp110
      'Modified by Lydia 2018/09/11 +CP118,CP152
      strExc(2) = "UPDATE CASEPROGRESS SET CP27=" & TransDate(Text9, 2) & ",cp14=" & CNULL(Text10) & "," & _
         "cp64=" & CNULL(ChgSQL(Text14)) & ",cp72=" & CNULL(ChgSQL(Text6)) & "," & _
         "cp50='" & Text8(0) & "'," & "cp51='" & ChgSQL(Text8(1)) & "'," & _
         "cp52='" & Text8(2) & "',cp84=" & Format(Val(txtCP84.Text)) & _
         ", CP16=NVL(CP16,0)-NVL(CP17,0)+" & Format(Val(txtCP84.Text)) & ", CP17=" & Format(Val(txtCP84.Text)) & _
         ", CP18=NVL(CP18,0),cp110=" & CNULL(cp(110)) & ",CP22=NULL,CP118='" & stCP118 & "',CP152=" & CNULL(stCP152, True) & " where cp09='" & strReceiveNo & "'"
   Else
      'Modify by morgan 2005/8/4 加 cp110
      'Modified by Lydia 2018/09/11 +CP118,CP152
      strExc(2) = "UPDATE CASEPROGRESS SET CP27=" & TransDate(Text9, 2) & ",cp14=" & CNULL(Text10) & "," & _
         "cp64=" & CNULL(ChgSQL(Text14)) & ",cp72=" & CNULL(ChgSQL(Text6)) & ",cp53=" & TransDate(Text12(0), 2) & "," & _
         "cp54=" & TransDate(Text12(1), 2) & "," & "cp50='" & Text8(0) & "'," & "cp51='" & ChgSQL(Text8(1)) & "'," & _
         "cp52='" & Text8(2) & "',cp84=" & Format(Val(txtCP84.Text)) & _
         ", CP16=NVL(CP16,0)-NVL(CP17,0)+" & Format(Val(txtCP84.Text)) & ", CP17=" & Format(Val(txtCP84.Text)) & _
         ", CP18=NVL(CP18,0),cp110=" & CNULL(cp(110)) & ",CP22=NULL,CP118='" & stCP118 & "',CP152='" & stCP152 & "' where cp09='" & strReceiveNo & "'"
   End If
   
 '911105 nick transation
 cnnConnection.Execute strExc(2)
   
   ii = 2
   'Add By Cheng 2002/07/03
   '若有輸入催審期限
   If Me.Text5.Text <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
   'intMax = objPublicData.GetNextProgressNo
   intMax = GetNextProgressNo
      'Modified by Lydia 2025/11/12 改抓最近工作天+PUB_GetWorkDay1
      strExc(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
         "NP07,NP08,NP09,NP10,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & _
         "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & 催審 & "," & _
         PUB_GetWorkDay1(TransDate(Text5.Text, 2), True) & "," & TransDate(Text5.Text, 2) & ",'" & strUserNum & "'," & intMax & ")"
        '911105 nick transation
        cnnConnection.Execute strExc(3)
         
      intMax = intMax + 1
      ii = 3
   End If
   
   PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130 'Add by Morgan 2009/3/20
   
    'Added by Lydia 2015/02/26 若已開請款單則換承辦人或核稿人時發Mail通知靜芳
   If cp(60) > "X" Then
         'Modified by Lydia 2019/10/17 本所案號+"-"
         'PUB_PointReAssignInform Text1 & Text2 & Text3 & Text4, cp(60), Text10.Tag, Text10.Text
         PUB_PointReAssignInform pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4)), cp(60), Text10.Tag, Text10.Text
   End If
   
   'Added by Lydia 2018/09/11 FCP電子送件若發文時若有規費，則自動產生行事曆。
   If txtCP118 = "Y" And Val(txtCP84) > 0 And stCP152 <> "" Then
       If Pub_AddReceiptCalendar1(pa(1), pa(2), pa(3), pa(4), cp(10), stCP152) = True Then
       End If
   End If
   'end 2018/09/11
   
   cnnConnection.CommitTrans
'911105 nick
   Exit Function
CheckingErr:
   cnnConnection.RollbackTrans
   FormSave = False
   
End Function

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label2(8) = pa(5)
      Case "英"
         Label2(8) = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         Label2(8) = pa(7)
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm060104_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
   'Add by Morgan 2005/8/4
   ReDim pa(TF_PA)
   ReDim cp(TF_CP)
   ReadPatent
   'Add by Morgan 2005/8/4
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   PUB_SetOurAgent lstNameAgent, pa(), cp(110), , True
   'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1300
   lstNameAgent.Width = 1300

   Label2(0) = strReceiveNo
   Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060104_6 = Nothing
End Sub

Private Sub ReadPatent()
 Dim Lbl As Object, txt As Object, i As Integer
   For Each Lbl In Label2
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      For i = 3 To 7
         If pa(i + 23) <> "" Then ChgType (i)
      Next
      Label2(8) = pa(5)
      If pa(76) <> "" Then Text11 = pa(76): ChgType (11)
   End If
   
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      If cp(13) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(cp(13), strExc(0)) Then Label2(1) = strExc(0)
         If ClsPDGetStaff(cp(13), strExc(0)) Then Label2(1) = strExc(0)
      End If
      ' 90.06.27 modify by louis 暫存案件性質
      m_CP10 = cp(10)
      
      Label2(2) = cp(6)
      If cp(27) = "" Then
         Text9 = strSrvDate(2)
      Else
         Text9 = cp(27)
      End If
      If cp(14) <> "" Then Text10 = cp(14): ChgType (10)
      'Added by Lydia 2015/02/26
      Text10.Tag = Text10.Text
      
      Text14 = cp(64)
      If cp(72) <> "" Then Text6 = cp(72): ChgType (8)
      Text12(0) = cp(53)
      Text12(1) = cp(54)
   End If
   
   ' 90.06.27 modify by louis
   If m_CP10 = "705" Then
      EnableTextBox Text12(0), False
      EnableTextBox Text12(1), False
   Else
      EnableTextBox Text12(0), True
      EnableTextBox Text12(1), True
   End If
   
   'Add by Morgan 2004/8/12
   txtCP84.Tag = cp(17)
   txtCP84.Text = txtCP84.Tag
   
   'Added by Lydia 2018/09/11 電子送件
   If cp(118) <> "" Then
       txtCP118 = "Y"
   End If
End Sub

Private Function ChgType(i As Integer) As Boolean
 Dim strTempName As String, j As Integer
   ChgType = False
   Select Case i
      Case 0 '發文日
         'Modify/Add By Cheng 2002/07/03
'         If ChkDate(Text9) Or Val(Text9.Text) > Val(strSrvDate(2)) Then
         If Not ChkDate(Text9) Then
         ElseIf Val(Text9.Text) > PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1))) Then
            MsgBox "發文日大於系統日下一個工作日, 請重新輸入!!!", vbExclamation + vbOKOnly
         Else
            ChgType = True
         End If
      Case 555 '催審期限
         If ChkDate(Text5) Then
            ChgType = True
         End If
      Case 3, 4, 5, 6, 7
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.LawGetName(pa(i + 23), strTempName) Then
         If ClsLawLawGetName(pa(i + 23), strTempName) Then
            Label2(i) = strTempName
            ChgType = True
         End If
      Case 8
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.GetCusCAJnam(Text6.Text, strExc(1), strExc(2), strExc(3)) = True Then
         If ClsLawGetCusCAJnam(Text6.Text, strExc(1), strExc(2), strExc(3)) = True Then
            ChgType = True
            For j = 1 To 3
               Text8(j - 1) = strExc(j)
            Next
         End If
      Case 10
         'ADD BY SONIA 2015/9/21 承辦人為外專程序時,改為操作人員
         Text10 = GetFCPUser(Text10)
         'END 2015/9/21
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(Text10, strTempName) Then
         If ClsPDGetStaff(Text10, strTempName) Then
            Label2(9) = strTempName
            ChgType = True
         End If
      Case 11
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.LawGetName(Text11, strTempName) Then
         If ClsLawLawGetName(Text11, strTempName) Then
            Label2(10) = strTempName
            ChgType = True
         End If
   End Select
End Function

Private Sub Text10_GotFocus()
  TextInverse Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then
      If Not ChgType(10) Then Cancel = True
   Else
      MsgBox "承辦人不可空白 !", vbCritical
      Cancel = True
   End If
End Sub

Private Sub Text11_GotFocus()
  TextInverse Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
    'Add By Cheng 2004/01/15
    If Me.Text11.Text = "" Then
        Me.Label2(10).Caption = ""
        Exit Sub
    End If
    'End
'   Text11 = UCase(Text11)
    If Not ChgType(11) Then
        Cancel = True
        Text11_GotFocus
    End If
    
   'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
   If Cancel = False Then
      If PUB_CheckStatus(Text11.Text) = False Then Cancel = True
   End If
End Sub

Private Sub Text12_GotFocus(Index As Integer)
  TextInverse Text12(Index)
End Sub

Private Sub Text12_Validate(Index As Integer, Cancel As Boolean)
   If Not ChkDate(Text12(Index)) Then
      Cancel = True
      TextInverse Text12(Index)
   Else
      If Index = 1 Then
         If ChkRange(Text12(0), Text12(1), "授權期間") = False Then
            TextInverse Text12(0)
            Cancel = True
         End If
      End If
   End If
End Sub

Private Sub Text14_GotFocus()
  TextInverse Text14
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 <> "" Then
      If Not ChgType(555) Then Cancel = True
   End If
   If Cancel = True Then Text5_GotFocus
End Sub

Private Sub Text6_GotFocus()
  TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
    'Add By Cheng 2004/01/15
    If Me.Text6.Text = "" Then
        Me.Text8(0).Text = ""
        Me.Text8(1).Text = ""
        Me.Text8(2).Text = ""
        Exit Sub
    End If
    'End
'   Text6 = UCase(Text6)
   ' 90.06.22 modify by Louis 被授權人編號補滿九碼
   If Len(Text6) < 9 Then
      Text6 = Text6 & String(9 - Len(Text6), "0")
   End If
   'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
   If Cancel = False Then
      If PUB_CheckStatus(Text6.Text) = False Then Cancel = True
   End If
   
   If Not ChgType(8) Then
      Cancel = True
        Text6_GotFocus
   End If
End Sub

Private Sub Text9_GotFocus()
  TextInverse Text9
End Sub

Private Sub Text9_LostFocus()
'Add By Cheng 2002/07/04
' 重新計算催審期限
ReCaculateSpecDate
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 <> "" Then
      If Not ChgType(0) Then
            Cancel = True
      'Added by Lydia 2018/09/11 當發文日有改時
      Else
            If Text9.Tag <> Text9 Then
                  Text9.Tag = Text9
                  txtPayToday.Text = Pub_FcpSetPayToday("1", Text9.Text, txtCP118.Text)
            End If
      'end 2018/09/11
      End If
   Else
      MsgBox "發文日不可空白 !", vbCritical
      Cancel = True
   End If
   'Add By Cheng 2002/07/03
   If Cancel = True Then Text9_GotFocus
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

If Me.Text9.Enabled = True Then
   Cancel = False
   Text9_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text5.Enabled = True Then
   Cancel = False
   Text5_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text10.Enabled = True Then
   Cancel = False
   Text10_Validate Cancel
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

For Each objTxt In Text12
   If objTxt.Enabled = True Then
      Cancel = False
      Text12_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

If Me.Text11.Enabled = True Then
   Cancel = False
   Text11_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

   'Add by Morgan 2004/8/12
   If txtCP84.Enabled = True Then
      Cancel = False
      txtCP84_Validate Cancel
      If Cancel = True Then
         txtCP84.SetFocus
         txtCP84_GotFocus
         Exit Function
      End If
   End If
   
   'Add by Morgan 2005/8/4
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   
   'Add By Sindy 2015/12/17 檢查是否有指定送件日期,若有不可小於指定日期送件
   If cp(142) <> "" Then
      'Modify By Sindy 2021/11/11 淑華說之後可以含當天發文
      'If cp(142) >= strSrvDate(1) Then
      If cp(142) > strSrvDate(1) Then
      '2021/11/11 END
         'Add By Sindy 2021/4/20
         'Modify By Sindy 2021/10/20 + 3.之後
         If ((cp(164) = "1" Or cp(164) = "") And cp(142) > strSrvDate(1)) Or _
            cp(164) = "3" Then '1.當天 3.之後
         '2021/4/20 END
            MsgBox "有指定送件日期（" & ChangeWStringToTDateString(cp(142)) & "），不可提前送件!!!"
            Exit Function
         End If
      End If
   End If
   '2015/12/17 END
   
   'Added by Lydia 2018/09/11
   If txtCP118 = "Y" And Val(txtCP84) > 0 Then
      If txtPayToday = "" Then
         MsgBox "電子送件請輸入是否當日扣款(Y/N)！", vbExclamation
         txtPayToday.SetFocus
         Exit Function
      End If
   End If
   'end 2018/09/11
      
TxtValidate = True
End Function

' 重新計算催審期限
Private Sub ReCaculateSpecDate()
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   
   Text5.Text = ""
   strSql = "SELECT CF05 FROM CASEFEE " & _
            "WHERE CF01='" & pa(1) & "' AND " & _
                  "CF02='" & pa(9) & "' AND " & _
                  "CF03='" & m_CP10 & "'"
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("CF05")) And "" & rsTmp.Fields("CF05") <> 0 Then
         Text5.Text = TransDate(CompDate(2, Val("" & rsTmp.Fields("CF05")), TransDate(Text9, 2)), 1)
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub
'Add by Morgan 2004/8/11
Private Sub txtCP84_GotFocus()
   TextInverse txtCP84
End Sub
'Add by Morgan 2004/8/11
Private Sub txtCP84_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub
'Add by Morgan 2004/8/11
Private Sub txtCP84_Validate(Cancel As Boolean)
   '台灣
   If pa(9) = "000" Then
      If Val(txtCP84.Text) <> Val(cp(17)) And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
         If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & cp(17) & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
            txtCP84.Tag = txtCP84.Text
         Else
            txtCP84_GotFocus
            Cancel = True
         End If
      End If
   End If
End Sub
'Add by Morgan 2005/8/4
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   cp(110) = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
         'Modify By Sindy 2021/5/10
         'cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         cp(110) = cp(110) & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         '2021/5/10 END
         Cancel = False
      End If
   Next
   If Cancel = True Then
      MsgBox "出名代理人不可空白！", vbExclamation
   Else
      If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
      m_AgentName = Mid(m_AgentName, 2) 'Add By Sindy 2021/5/10
   End If
End Sub

'Added by Lydia 2018/09/11
Private Sub txtCP118_GotFocus()
   TextInverse txtCP118
   CloseIme
End Sub

Private Sub txtCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP118_Change()
    txtPayToday.Text = Pub_FcpSetPayToday("1", Text9.Text, txtCP118.Text)
End Sub

Private Sub txtPayToday_GotFocus()
   TextInverse txtPayToday
   CloseIme
End Sub

Private Sub txtPayToday_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub
'end 2018/09/11

'Add By Sindy 2022/5/17
Private Sub txtRecDate_GotFocus()
   TextInverse txtRecDate
End Sub
Private Sub txtRecDate_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub
Private Sub txtEmail_GotFocus()
   TextInverse txtEmail
End Sub
Private Sub txtEmail_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      Beep
      KeyAscii = 0
   End If
End Sub
Private Sub txtRecDate_Validate(Cancel As Boolean)
   If txtRecDate.Tag <> txtRecDate.Text Then
      If txtRecDate = "Y" Then
         txtEmail = "Y"
      End If
   End If
   txtRecDate.Tag = txtRecDate.Text
End Sub
'2022/5/17 END
