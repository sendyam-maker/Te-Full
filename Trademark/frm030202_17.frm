VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030202_17 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(補理由書)"
   ClientHeight    =   5400
   ClientLeft      =   4060
   ClientTop       =   2460
   ClientWidth     =   9200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   9200
   Begin VB.TextBox textPrint 
      Height          =   285
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   8
      Top             =   4071
      Width           =   492
   End
   Begin VB.TextBox textCP118 
      Height          =   285
      Left            =   4905
      MaxLength       =   1
      TabIndex        =   9
      Top             =   4071
      Width           =   375
   End
   Begin VB.TextBox textCP113 
      Height          =   285
      Left            =   5715
      MaxLength       =   4
      TabIndex        =   2
      Top             =   3189
      Width           =   600
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   6480
      MaxLength       =   1
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3189
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox textCP84 
      Alignment       =   1  '靠右對齊
      Height          =   285
      Left            =   3330
      TabIndex        =   1
      Top             =   3189
      Width           =   1425
   End
   Begin VB.CommandButton cmdCaseProgress 
      Caption         =   "案件進度(&C)"
      Height          =   400
      Left            =   4608
      TabIndex        =   12
      Top             =   50
      Width           =   1272
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8100
      TabIndex        =   15
      Top             =   50
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5940
      TabIndex        =   13
      Top             =   50
      Width           =   912
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6888
      TabIndex        =   14
      Top             =   50
      Width           =   1152
   End
   Begin VB.TextBox textCP43 
      Height          =   285
      Left            =   1710
      MaxLength       =   9
      TabIndex        =   6
      Top             =   3780
      Width           =   1932
   End
   Begin VB.TextBox textCP08 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   4905
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   2888
      Width           =   2532
   End
   Begin VB.TextBox textCP23 
      Height          =   285
      Left            =   4905
      MaxLength       =   1
      TabIndex        =   7
      Top             =   3781
      Width           =   372
   End
   Begin VB.TextBox textCP27 
      Height          =   285
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   0
      Top             =   3189
      Width           =   1092
   End
   Begin VB.TextBox textCP18 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   4905
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3490
      Width           =   2532
   End
   Begin VB.TextBox textDN 
      Height          =   285
      Left            =   1200
      MaxLength       =   9
      TabIndex        =   5
      Top             =   3490
      Width           =   492
   End
   Begin VB.TextBox textCP49 
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Top             =   4372
      Width           =   7752
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1082
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   4905
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1082
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1383
      Width           =   1245
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   480
      Width           =   1245
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   781
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   4905
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   781
      Width           =   2532
   End
   Begin VB.TextBox textTM20 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   7035
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   480
      Width           =   1245
   End
   Begin VB.TextBox textCP12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   4155
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   480
      Width           =   1245
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   4905
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2286
      Width           =   2532
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "(N:不印)"
      Height          =   180
      Left            =   1785
      TabIndex        =   69
      Top             =   4123
      Width           =   645
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      Caption         =   "是否電子送件:          (Y: 是)"
      Height          =   180
      Left            =   3750
      TabIndex        =   68
      Top             =   4123
      Width           =   2085
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1200
      TabIndex        =   67
      Top             =   2587
      Width           =   7755
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13679;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM44 
      Height          =   285
      Left            =   1200
      TabIndex        =   66
      Top             =   2888
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM79 
      Height          =   285
      Left            =   1200
      TabIndex        =   65
      Top             =   1985
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM80 
      Height          =   285
      Left            =   4905
      TabIndex        =   64
      Top             =   1985
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM81 
      Height          =   285
      Left            =   1200
      TabIndex        =   63
      Top             =   2286
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1200
      TabIndex        =   62
      Top             =   1684
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM78 
      Height          =   285
      Left            =   4905
      TabIndex        =   61
      Top             =   1684
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   285
      Left            =   4155
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   1383
      Width           =   1545
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "2725;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   7035
      TabIndex        =   59
      Top             =   1383
      Width           =   1545
      VariousPropertyBits=   671105055
      Size            =   "2725;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   7650
      TabIndex        =   4
      Top             =   3210
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
   Begin MSForms.TextBox textCP64 
      Height          =   585
      Left            =   1200
      TabIndex        =   10
      Top             =   4680
      Width           =   7755
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13679;1032"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作時數:"
      Height          =   180
      Index           =   12
      Left            =   4920
      TabIndex        =   58
      Top             =   3241
      Width           =   765
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "申請人5 :"
      Height          =   180
      Left            =   120
      TabIndex        =   57
      Top             =   2340
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "申請人4 :"
      Height          =   180
      Left            =   4035
      TabIndex        =   56
      Top             =   2037
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "申請人3 :"
      Height          =   180
      Left            =   120
      TabIndex        =   55
      Top             =   2037
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "申請人2 :"
      Height          =   180
      Left            =   4035
      TabIndex        =   54
      Top             =   1736
      Width           =   720
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   6720
      TabIndex        =   53
      Top             =   3240
      Width           =   900
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      Caption         =   "發文規費："
      Height          =   180
      Left            =   2400
      TabIndex        =   52
      Top             =   3241
      Width           =   900
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "進度備註 :"
      Height          =   180
      Left            =   120
      TabIndex        =   51
      Top             =   4770
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "補理由書總收文號 :"
      Height          =   180
      Index           =   7
      Left            =   120
      TabIndex        =   50
      Top             =   3855
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "機關文號 :"
      Height          =   180
      Index           =   5
      Left            =   3945
      TabIndex        =   49
      Top             =   2940
      Width           =   810
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "列印定稿 :"
      Height          =   180
      Left            =   120
      TabIndex        =   47
      Top             =   4158
      Width           =   810
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "(1:勝 2:敗)"
      Height          =   180
      Left            =   5340
      TabIndex        =   46
      Top             =   3833
      Width           =   795
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "預估勝敗 :"
      Height          =   180
      Left            =   3945
      TabIndex        =   45
      Top             =   3833
      Width           =   810
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "發文日 :"
      Height          =   180
      Left            =   120
      TabIndex        =   44
      Top             =   3249
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "點　　數 :"
      Height          =   180
      Index           =   10
      Left            =   3945
      TabIndex        =   43
      Top             =   3542
      Width           =   810
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      Caption         =   "是否輸入D/N:"
      Height          =   180
      Left            =   120
      TabIndex        =   42
      Top             =   3552
      Width           =   1050
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      Caption         =   "(Y:輸入)"
      Height          =   180
      Left            =   1785
      TabIndex        =   41
      Top             =   3542
      Width           =   645
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "條　　款 :"
      Height          =   180
      Left            =   120
      TabIndex        =   40
      Top             =   4461
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱 :"
      Height          =   180
      Left            =   120
      TabIndex        =   38
      Top             =   2643
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "審定號數 :"
      Height          =   180
      Left            =   120
      TabIndex        =   37
      Top             =   1128
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員 :"
      Height          =   180
      Index           =   11
      Left            =   6075
      TabIndex        =   36
      Top             =   1435
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "彼所案號 :"
      Height          =   180
      Index           =   9
      Left            =   3945
      TabIndex        =   35
      Top             =   1134
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質 :"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   34
      Top             =   1431
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號 :"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   33
      Top             =   522
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號 :"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   32
      Top             =   825
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號 :"
      Height          =   180
      Left            =   3945
      TabIndex        =   31
      Top             =   833
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發證日 :"
      Height          =   180
      Index           =   3
      Left            =   6075
      TabIndex        =   30
      Top             =   532
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業務區別 :"
      Height          =   180
      Index           =   2
      Left            =   3195
      TabIndex        =   29
      Top             =   532
      Width           =   810
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "承辦人 :"
      Height          =   180
      Left            =   3195
      TabIndex        =   28
      Top             =   1435
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "商標種類 :"
      Height          =   180
      Index           =   4
      Left            =   3945
      TabIndex        =   27
      Top             =   2338
      Width           =   810
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "申請人1 :"
      Height          =   180
      Left            =   120
      TabIndex        =   26
      Top             =   1734
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "代理人 :"
      Height          =   180
      Left            =   120
      TabIndex        =   25
      Top             =   2946
      Width           =   630
   End
End
Attribute VB_Name = "frm030202_17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/10 改成Form2.0 ;cmbTM05、textCP13、textCP14、textCP64、textTM44、textTM23、textTM78~81、lstNameAgent
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 收文號
Dim m_CP09 As String
' 申請國家
Dim m_TM10 As String
' 案件性質代號
Dim m_CP10 As String
' 承辦人員
Dim m_CP14 As String
Dim m_CP82 As String 'Added by Lydia 2018/08/10 發文時間

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
' 儲存商標基本檔或服務業務基本檔檔案欄位的串列
Dim m_TMSPList() As FIELDITEM
Dim m_TMSPCount As Integer
' 儲存案件進度檔檔案欄位的串列
Dim m_CPList() As FIELDITEM
Dim m_CPCount As Integer
'add by nick 2004/08/13
Dim m_CP84 As String       '發文規費
'add by nickc 2006/01/25
Dim m_CP110 As String
'add by nickc 2008/02/22
Dim m_CP44 As String
Dim m_CP116 As String
Dim m_TM44 As String
Dim m_TM119 As String
Dim m_TM120 As String
Dim m_CP09s As String, m_CP123s As String 'Add by Sindy 98/3/24 收文號,是否算發文室案件
Dim m_CP130s As String 'Add by Sindy 2009/4/24 發文-主管機關


Private Sub cmdCancel_Click()
   frm030202_01.Show
   Unload Me
End Sub

' 案件進度查詢
Private Sub cmdCaseProgress_Click()
   frm030202_04.SetData 0, m_TM01, True
   frm030202_04.SetData 1, m_TM02, False
   frm030202_04.SetData 2, m_TM03, False
   frm030202_04.SetData 3, m_TM04, False
   frm030202_04.SetData 4, m_CP09, False
   frm030202_04.SetParent Me
   Me.Hide
   frm030202_04.Show
   frm030202_04.QueryData
End Sub

Private Sub cmdExit_Click()
   Unload frm030202_01
   Unload Me
   'frm030202_01.Show
End Sub

Private Sub cmdOK_Click()
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      
      'Add by Sindy 98/3/24 設定是否算發文室案件
      If m_TM10 = "000" Then
         'Modify By Sindy 2012/12/20 若為電子送件則不經發文室
         'Modify By Sindy 2023/8/1 電子送件欄位值不是空白者,即為電子送件
         If (textCP118.Visible = True And textCP118 <> "") Then
            'Added by Morgan 2016/5/16 電子送件也要記錄主管機關
            If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27, , True) = False Then
               Exit Sub
            End If
            'end 2016/5/16
            'add by sonia 2016/3/31
            strExc(0) = Trim(InputBox("請輸入智慧局收文文號!!"))
            If strExc(0) = "" Then
               Exit Sub
            Else
               textCP64 = "智慧局收文文號:" & strExc(0) & ";" & Trim(textCP64)
            End If
            'end 2016/3/31
         Else
            'Add by Sindy 2009/4/24
            If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27) = False Then
               Exit Sub
            Else
               If m_CP123s = "Y" Then
                  'modify by sonia 2014/6/23 加傳發文規費, P-108903
                  If ModifyDispatch(textCP09, m_CP09s, m_CP123s, textCP84, textCP27) = False Then
                      Exit Sub
                  End If
               End If
            End If
         End If '2012/12/20 End
      End If
      
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 更新欄位輸入的內容
      OnUpdateField
      ' 存檔
      'edit by  nick 2004/11/03
      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      'Add By Sindy 2012/4/5 CFT,FCT所有案件性質發文時,檢查代表圖是否存在
      'Mark by Amy 2018/07/31 因ChkIsExistImg不使用,與Sindy確認FCT不彈Msg故拿掉
      'Call ChkIsExistImg(m_TM01, m_TM02, m_TM03, m_TM04)

        'Added by Lyddia 2018/08/10 增加重新發文判斷
        strExc(1) = m_CP82
        If Val(m_CP82) > 0 Then
             If MsgBox("重新發文是否上傳檔案到卷宗區？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                 strExc(1) = ""
             End If
        End If
        If Val(strExc(1)) = 0 Then
        'end 2018/08/10
            'Added by Lydia 2018/07/19 FCT發文自動將下載的PDF檔,上傳到卷宗區
            If Pub_AutoSavePdf_FCT(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, m_CP10) = False Then
            End If
            'end 2018/07/19
        End If 'end 2018/08/10
      
      '************   90.11.23 nick   清畫面
      'frm030202_01.radio(0).Value = True
      'frm030202_01.textCP09.Enabled = True
      'frm030202_01.textCP09.Text = ""
      'frm030202_01.textTM01.Enabled = False
      'frm030202_01.textTM01.Text = ""
      'frm030202_01.textTM02.Enabled = False
      'frm030202_01.textTM02.Text = ""
      'frm030202_01.textTM02_2.Enabled = False
      'frm030202_01.textTM02_2.Text = ""
      'frm030202_01.textTM03.Enabled = False
      'frm030202_01.textTM03.Text = ""
      'frm030202_01.textTM04.Enabled = False
      'frm030202_01.textTM04.Text = ""
      'frm030202_01.grdList.Clear
      'frm030202_01.grdList.Rows = 2
      'frm030202_01.QueryData
      'frm030202_01.Show
      '*************************************
      
      Call PUB_FCTSendRecvMail(m_CP09) 'Add By Sindy 2024/10/30 外商發文時,增加發Mail通知承辦人及副本給判發主管
      'Add By Sindy 2024/8/19
      If frm030202_01.bolIsEMPFlow = True Then
         frm090202_4.QueryData
      End If
      '2024/8/19 End
      'Ken 91.04.09 -- Start
      If textDN = "Y" Then
        'Add By Cheng 2003/03/19
        '新增地址條列表資料
'edit by nick 2004/11/17  因為請款已經有產生了
'        pub_AddressListSN = pub_AddressListSN + 1
'        PUB_AddNewAddressList strUserNum, m_TM01, m_TM02, m_TM03, m_TM04, "" & pub_AddressListSN, "0"
         Screen.MousePointer = vbHourglass
         Frmacc21h0.Show
         mdiMain.ToolShow
         mdiMain.tool1_enabled
         Screen.MousePointer = vbDefault
         Set Frmacc21h0.frmlink = frm030202_01
         'add by nick 2004/11/24
         Frmacc21h0.IsPrintAddress = False
      Else
         'Add By Cheng 2002/04/30
         '若有未發文資料顯示警告
         If PUB_GetCPunIssueDatas("" & Me.textTMKey.Text) = False Then
            'Add By Sindy 2024/8/19
            If frm030202_01.bolIsEMPFlow = True Then
               Unload frm030202_01
               frm090202_4.Show
               Unload Me
               Exit Sub
            End If
            '2024/8/19 End
         End If
         frm030202_01.Show
         ' 90.12.07 modify by louis
         frm030202_01.Clear1
      End If
      'Ken 91.04.09 -- End
      Unload Me
   End If
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
End Sub

Private Sub Form_Activate()
'add by nickc 2005/08/23
If (pub_ModifyCaseNum = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 And pub_ModifyCaseNum <> "") Then
   pub_ModifyCaseNum = ""
   QueryData
End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM08.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM20.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   'add by nickc 2007/01/30
   textTM78.BackColor = &H8000000F
   textTM79.BackColor = &H8000000F
   textTM80.BackColor = &H8000000F
   textTM81.BackColor = &H8000000F
   
   textTM44.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
      
   textCP08.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP12.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP18.BackColor = &H8000000F
   
   MoveFormToCenter Me
   'Add by nickc 2006/01/25
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   Text7.Visible = False
   lstNameAgent.Clear
   lstNameAgent.Visible = True
   lblNameAgent.Visible = True
   'Added by Lydia 2021/09/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1055
   lstNameAgent.Width = 1500
   
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP09 = Empty
   End If
   
   Select Case nType
      ' 收文號
      Case 0: m_CP09 = strData
      ' 相關總收文號
      Case 99: textCP43 = strData
   End Select
End Sub

' 清除商標基本檔檔案欄位串列
Private Sub ClearTMSPFieldList()
   If m_TMSPCount > 0 Then
      Erase m_TMSPList
   End If
   m_TMSPCount = 0
End Sub

' 設定商標基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetTMSPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TMSPCount - 1
      If m_TMSPList(nPos).fiName = strFieldName Then
         bFind = True
         m_TMSPList(nPos).fiOldData = strFieldData
         m_TMSPList(nPos).fiNewData = strFieldData
         m_TMSPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_TMSPList(m_TMSPCount + 1)
      m_TMSPList(m_TMSPCount).fiName = strFieldName
      m_TMSPList(m_TMSPCount).fiOldData = strFieldData
      m_TMSPList(m_TMSPCount).fiNewData = strFieldData
      m_TMSPList(m_TMSPCount).fiType = nFieldType
      m_TMSPCount = m_TMSPCount + 1
   End If
End Sub

' 設定商標基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetTMSPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TMSPCount - 1
      If m_TMSPList(nPos).fiName = strFieldName Then
         bFind = True
         m_TMSPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

' 清除案件進度檔檔案欄位串列
Private Sub ClearCPFieldList()
   If m_CPCount > 0 Then
      Erase m_CPList
   End If
   m_CPCount = 0
End Sub

' 設定案件進度檔欄位串列中的欄位內容
Private Sub SetCPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiOldData = strFieldData
         m_CPList(nPos).fiNewData = strFieldData
         m_CPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_CPList(m_CPCount + 1)
      m_CPList(m_CPCount).fiName = strFieldName
      m_CPList(m_CPCount).fiOldData = strFieldData
      m_CPList(m_CPCount).fiNewData = strFieldData
      m_CPList(m_CPCount).fiType = nFieldType
      m_CPCount = m_CPCount + 1
   End If
End Sub

' 設定案件進度檔欄位串列中的欄位內容
Private Sub SetCPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

' 取得商標基本檔的欄位內容
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'add by nickc 2008/02/22
      m_TM44 = CheckStr(rsTmp.Fields("TM44"))
      m_TM119 = CheckStr(rsTmp.Fields("TM119"))
      m_TM120 = CheckStr(rsTmp.Fields("TM120"))
      ' 審定號數
      If IsNull(rsTmp.Fields("TM15")) = False Then: textTM15 = rsTmp.Fields("TM15")
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then: textTM12 = rsTmp.Fields("TM12")
      ' 發證日
      If IsNull(rsTmp.Fields("TM20")) = False Then: textTM20 = TAIWANDATE(rsTmp.Fields("TM20"))
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then: cmbTM05.AddItem rsTmp.Fields("TM05")
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("TM06")) = False Then: cmbTM05.AddItem rsTmp.Fields("TM06")
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("TM07")) = False Then: cmbTM05.AddItem rsTmp.Fields("TM07")
      ' 顯示案件名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 商標種類
      textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then: m_TM10 = rsTmp.Fields("TM10")
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then: textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      'add by nickc 2007/01/30
      If IsNull(rsTmp.Fields("TM78")) = False Then: textTM78 = GetCustomerName("" & rsTmp.Fields("TM78"), 0)
      If IsNull(rsTmp.Fields("TM79")) = False Then: textTM79 = GetCustomerName("" & rsTmp.Fields("TM79"), 0)
      If IsNull(rsTmp.Fields("TM80")) = False Then: textTM80 = GetCustomerName("" & rsTmp.Fields("TM80"), 0)
      If IsNull(rsTmp.Fields("TM81")) = False Then: textTM81 = GetCustomerName("" & rsTmp.Fields("TM81"), 0)
      
      ' FC代理人
      If IsNull(rsTmp.Fields("TM44")) = False Then: textTM44 = GetFAgentName(rsTmp.Fields("TM44"))
      ' 彼所案號
      If IsNull(rsTmp.Fields("TM45")) = False Then: textTM45 = rsTmp.Fields("TM45")
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得案件進度檔的欄位內容
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim strSubSQL As String
   Dim rsTmp As New ADODB.Recordset
   Dim rsSubTmp As New ADODB.Recordset
   Dim strDate As String
   Dim strCP27 As String
   Dim strCP44 As String
   Dim strCP45 As String
   Dim nIndex As Integer
   Dim bFind As Boolean
   
   ' 系統日
   strDate = DBDATE(SystemDate())
   ' 收文號
   textCP09 = m_CP09
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'add by nickc 2008/02/22
      m_CP116 = CheckStr(rsTmp.Fields("CP116"))
      m_CP44 = CheckStr(rsTmp.Fields("CP44"))
      m_CP82 = "" & rsTmp.Fields("CP82")  'Added by Lydia 2018/08/10 發文時間
      ' 機關文號
      If IsNull(rsTmp.Fields("CP08")) = False Then: textCP08 = rsTmp.Fields("CP08")
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 業務區別
      '910718 Sieg
      If IsNull(rsTmp.Fields("CP12")) = False Then
         textCP12 = GetDepartmentName(rsTmp.Fields("CP12"))
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      End If
      
      'Add By Sindy 98/03/11
      '工作時數
      textCP113 = "" & rsTmp.Fields("CP113")
      SetCPFieldOldData "CP113", textCP113, 1
      '98/03/11 End
      
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         m_CP14 = rsTmp.Fields("CP14")
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      ' 發文日(預設為系統日)
      strCP27 = Empty
      textCP27 = TAIWANDATE(SystemDate())
      If IsNull(rsTmp.Fields("CP27")) = False Then: strCP27 = rsTmp.Fields("CP27")
      SetCPFieldOldData "CP27", strCP27, 1
      ' 點數
      textCP18 = Empty
      If IsNull(rsTmp.Fields("CP18")) = False Then: textCP18 = rsTmp.Fields("CP18")
      ' 預估結果(預估勝敗)
      textCP23 = Empty
      If IsNull(rsTmp.Fields("CP23")) = False Then: textCP23 = rsTmp.Fields("CP23")
      SetCPFieldOldData "CP23", textCP23, 0
      ' 相關總收文號
      textCP43 = Empty
      If IsNull(rsTmp.Fields("CP43")) = False Then: textCP43 = rsTmp.Fields("CP43")
      SetCPFieldOldData "CP43", textCP43, 0
      ' 條款
      textCP49 = Empty
      If IsNull(rsTmp.Fields("CP49")) = False Then: textCP49 = rsTmp.Fields("CP49")
      SetCPFieldOldData "CP49", textCP49, 0
      ' 進度備註
      textCP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then: textCP64 = rsTmp.Fields("CP64")
      SetCPFieldOldData "CP64", textCP64, 0
      
      'Add By Sindy 2012/12/20
      ' 是否電子送件
      textCP118 = Empty
      If IsNull(rsTmp.Fields("CP118")) = False Then
         textCP118 = rsTmp.Fields("CP118")
      End If
      SetCPFieldOldData "CP118", textCP118, 0
      
      'add by nick 2004/08/13 發文規費
      If IsNull(rsTmp.Fields("CP17")) = False And textCP84.Enabled = True Then
          m_CP84 = CheckStr(rsTmp.Fields("CP17"))
      End If
      'Add By Sindy 2012/12/20 電子送件發文規費預設為承辦人已輸入的金額
      If rsTmp.Fields("cp118") = "Y" Then
         textCP84 = Val("" & rsTmp.Fields("cp84"))
      End If
      'end 2012/12/20
      
      'add by nickc 2006/02/10
      Text7 = CheckStr(rsTmp.Fields("CP22"))
      SetCPFieldOldData "CP22", Text7, 0
   End If
   
   'add by nickc 2006/01/25
   'SetCPFieldOldData "CP110", m_CP110, 0
   'Modify By Sindy 2010/9/20
   If m_CP110 = "" Then m_CP110 = CheckStr(rsTmp.Fields("cp110"))
   SetCPFieldOldData "CP110", CheckStr(rsTmp.Fields("cp110")), 0
   '2010/9/20 End
   
   rsTmp.Close
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
End Sub

' 讀取資料庫
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   'add by nickc 2006/01/25
   Dim tm(1 To 4) As String
   ' 先清除商標基本檔或服務業務基本檔欄位串列
   ClearTMSPFieldList
   ' 先清除案件進度檔欄位串列
   ClearCPFieldList
   
   ' 先取得本所案號
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 本所案號
      If IsNull(rsTmp.Fields("CP01")) = False Then: m_TM01 = rsTmp.Fields("CP01")
      If IsNull(rsTmp.Fields("CP02")) = False Then: m_TM02 = rsTmp.Fields("CP02")
      If IsNull(rsTmp.Fields("CP03")) = False Then: m_TM03 = rsTmp.Fields("CP03")
      If IsNull(rsTmp.Fields("CP04")) = False Then: m_TM04 = rsTmp.Fields("CP04")
   End If
   rsTmp.Close
   
   ' 本所案號
   'Modify By Cheng 2002/04/30
'   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   textTMKey.Text = m_TM01 & "-" & m_TM02 & "-" & IIf(Len("" & m_TM03) <= 0, "0", m_TM03) & "-" & IIf(Len("" & m_TM04) <= 0, "00", m_TM04)
   
   'add by nickc 2006/01/25
   tm(1) = m_TM01
   tm(2) = m_TM02
   tm(3) = m_TM03
   tm(4) = m_TM04
   'Modify By Sindy 2010/9/20 預設出名代理人,移到下面讀完CP再做
   'PUB_SetOurAgent lstNameAgent, tm(), m_CP110
   '2010/9/20 End
   
   '讀取商標基本檔
   QueryTradeMark
   
   ' 讀取案件進度檔
   QueryCaseProgress
   'Modified by Lydia 2021/09/10 + Form 2.0 = True
   'PUB_SetOurAgent lstNameAgent, tm(), m_CP110 'Modify By Sindy 2010/9/20
   PUB_SetOurAgent lstNameAgent, tm(), m_CP110, m_CP10, True
   
   'Add By Sindy 2012/12/20 外商000台灣案所有案件性質加電子送件功能
   If m_TM01 = "FCT" And m_TM10 = "000" Then
      Label43.Visible = True
      textCP118.Visible = True
   Else
      Label43.Visible = False
      textCP118.Visible = False
   End If
   '2012/12/20 End
   
   Set rsTmp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm030202_17 = Nothing
End Sub

'add by nickc 2006/01/25
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   m_CP110 = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/5 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modified by Lydia 2021/09/10 改模組
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         bolCheck = True
      End If
   Next
   If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
   If bolCheck = True Then
      Text7 = ""
   Else
      Text7 = "N"
      MsgBox "未勾選代理人!", vbInformation, "必要欄位！"
      Cancel = True
   End If
End Sub

' 預估勝敗
Private Sub textCP23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP23) = False Then
      Select Case textCP23
         Case "1", "2":
         Case Else
            strTit = "檢核資料"
            strMsg = "預估勝敗只可輸入1或2"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP23_GotFocus
      End Select
   End If
End Sub

' 發文日
Private Sub textCP27_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strDate As String
   Cancel = False
   
   If IsEmptyText(textCP27) = False Then
      ' 發文日日期不正確
      If CheckIsTaiwanDate(textCP27, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的發文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      
      ' 發文日日期不可超過系統日
      'edit by nick 2004/08/31 系統日加一天
      'If Val(DBDATE(textCP27)) > Val(DBDATE(SystemDate())) Then
      If Val(DBDATE(textCP27)) > Val(DBDATE(PUB_GetWorkDay(2))) Then
         Cancel = True
         strTit = "資料檢核"
         'edit by nick 2004/08/31
         'strMsg = "發文日不可超過系統日"
         strMsg = "發文日不可超過系統日加一天"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

' 補理由書總收文號
Private Sub textCP43_Validate(Cancel As Boolean)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP43) = False Then
      If textCP43 = m_CP09 Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "相關總收文號不可為本案之收文號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP43_GotFocus
         GoTo EXITSUB
      End If
      
      strSql = "SELECT * FROM CaseProgress " & _
               "WHERE CP01 = '" & m_TM01 & "' AND " & _
                     "CP02 = '" & m_TM02 & "' AND " & _
                     "CP03 = '" & m_TM03 & "' AND " & _
                     "CP04 = '" & m_TM04 & "' AND " & _
                     "CP09 = '" & textCP43 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount <= 0 Then
         rsTmp.Close
         Cancel = True
         strTit = "資料檢核"
         strMsg = "相關總收文號資料不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP43_GotFocus
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub textCP49_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 條款
Private Sub textCP49_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim strTemp As String
   Dim nResponse
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Cancel = False
   ' 無資料時不做任何檢查
   If IsEmptyText(textCP49) = True Then
      GoTo EXITSUB
   End If
   
   nCount = GetSubStringCount(textCP49)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textCP49, nIndex)
      'Modify By Cheng 2002/07/22
      '條款每項必須輸入4碼
'      If Len(strTemp) > 4 Then
      'Modify By Sindy 2012/7/5
      'If Len(strTemp) <> 4 Then
      If Len(strTemp) <> 4 And Len(strTemp) <> 5 Then
      '2012/7/5 End
         Cancel = True
         strTit = "條款"
         strMsg = "條款內容<" & strTemp & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP49_GotFocus
         GoTo EXITSUB
      End If
      
      ' 檢查主張內容分類表
      strSql = "SELECT * FROM ClaimContents " & _
               "WHERE CC01 = '" & Right(strTemp, 1) & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount <= 0 Then
         Cancel = True
         strTit = "條款"
         strMsg = "條款內容<" & strTemp & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP49_GotFocus
         rsTmp.Close
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      ' 檢查
      'Modify By Sindy 2012/7/5
'      strSql = "SELECT * FROM LAW " & _
'               "WHERE LW01 = '" & Mid(strTemp, 1, 3) & "' "
      strSql = "SELECT * FROM LAW " & _
               "WHERE LW01 = '" & Mid(strTemp, 1, Len(strTemp) - 1) & "' "
      '2012/7/5 End
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount <= 0 Then
         Cancel = True
         strTit = "條款"
         strMsg = "條款代號<" & strTemp & ">不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP49_GotFocus
         rsTmp.Close
         GoTo EXITSUB
      End If
      rsTmp.Close
   Next nIndex
EXITSUB:
   Set rsTmp = Nothing
End Sub

'edit by nickc 2006/01/25 刪除
'Private Sub textCP64_2_GotFocus()
'   TextInverse textCP64_2
'End Sub

'add by nick 2004/08/13
Private Sub textCP84_GotFocus()
   InverseTextBox textCP84
End Sub

Private Sub textCP84_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
If IsEmptyText(textCP84) = False Then
    If IsNumeric(textCP84) = False Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "請輸入數字"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP84_GotFocus
    Else
        textCP84.Text = Trim(Val(textCP84.Text))
    End If
End If
End Sub

Private Sub textDN_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否輸入D/N
Private Sub textDN_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textDN) = False Then
      Select Case textDN
         Case " ", "Y":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textDN_GotFocus
      End Select
   End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 列印定稿
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         Case " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

' 更新欄位的內容
Private Sub OnUpdateField()
   ' 預估結果
   SetCPFieldNewData "CP23", textCP23
   ' 發文日
   SetCPFieldNewData "CP27", DBDATE(textCP27)
   ' 相關總收文號
   SetCPFieldNewData "CP43", textCP43
   ' 條款
   SetCPFieldNewData "CP49", textCP49
   ' 進度備註
   '910801 Sieg 602
'edit by nickc 2006/01/25 改 cp110
'   If textCP64_2 <> "" Then
'      If textCP64 = "" Then
'         textCP64 = textCP64_2
'      Else
'         textCP64 = textCP64 & "," & textCP64_2
'      End If
'   End If
   SetCPFieldNewData "CP64", textCP64
   'add by nickc 2006/01/25
   SetCPFieldNewData "CP110", m_CP110
   'add by nickc 2006/02/10
   SetCPFieldNewData "CP22", Text7
   ' Add By Sindy 98/03/11
   SetCPFieldNewData "CP113", textCP113
   
   'Add By Sindy 2012/12/20
   ' 是否電子送件
   SetCPFieldNewData "CP118", textCP118
End Sub

' 更新商標基本檔的相關欄位
Private Sub OnUpdateTradeMark()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
   ' 更新案件進度檔
   strSql = "UPDATE TradeMark SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            ' 91.03.25 modify by louis (單引號)
            'strTmp = m_TMSPList(nIndex).fiName & " = '" & m_TMSPList(nIndex).fiNewData & "'"
            strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
         Else
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = " & 0
            Else
               strTmp = m_TMSPList(nIndex).fiName & " = " & m_TMSPList(nIndex).fiNewData
            End If
         End If
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   ' 設定SQL語法更新的條件
   strSql = strSql & " " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
End Sub

' 更新商標基本檔的相關欄位
Private Sub OnUpdateCaseProperty()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   
   ' 更新案件進度檔
   strSql = "UPDATE CaseProgress SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_CPCount - 1
      strTmp = Empty
      If m_CPList(nIndex).fiOldData <> m_CPList(nIndex).fiNewData Then
         If m_CPList(nIndex).fiType = 0 Then
            ' 91.03.25 modify by louis (單引號)
            'strTmp = m_CPList(nIndex).fiName & " = '" & m_CPList(nIndex).fiNewData & "'"
            strTmp = m_CPList(nIndex).fiName & " = '" & ChgSQL(m_CPList(nIndex).fiNewData) & "'"
         Else
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = " & 0
            Else
               strTmp = m_CPList(nIndex).fiName & " = " & m_CPList(nIndex).fiNewData
            End If
         End If
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   ' 設定SQL語法更新的條件
   strSql = strSql & " " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
   
End Sub

'edit by nick 2004/11/03
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim strSql As String
   Dim strNP07 As String
   Dim strNP08 As String
   Dim strNP22 As String
   Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2012/9/10
   
 '911107 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   ' 更新案件進度檔
   OnUpdateCaseProperty
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新相關總收文號資料中的條款
   If IsNull(textCP43) = False Then
      strSql = "UPDATE CaseProgress SET CP49 = '" & textCP49 & "' " & _
               "WHERE CP09 = '" & textCP43 & "' "
      cnnConnection.Execute strSql
   End If
   
   'add by nick 2004/08/13 更新實際發文規費
   If textCP84.Enabled = True Then
      strSql = "Update CaseProgress Set CP84=" & Trim(Val(textCP84.Text)) & " Where CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   'Add By Sindy 2012/12/20 若為電子送件則自動設定為不經發文室
   '以防動作為重新發文, 所以一併把發文室相關欄位清空
   If textCP118.Visible = True And textCP118 = "Y" Then
      strSql = "Update CaseProgress Set CP123=null" & _
                                                          ",CP124=null" & _
                                                          ",CP125=null" & _
                                                          ",CP28=null" & _
                                                          ",CP131=null" & _
                                                          ",CP132=null" & _
                   " Where CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   'Add by Sindy 98/3/24
   If m_TM10 = "000" Then
      'Modify By Sindy 2009/04/24
      'PUB_UpdateDispatch m_CP09s, m_CP123s
      PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130s
   End If
   
   'Add By Sindy 2010/7/8 檢查商品資料與基本檔商品類別是否一致
   Call CheckTMGoodsErr(m_TM01, m_TM02, m_TM03, m_TM04, False, True, m_CP14)
   
   'Add By Sindy 2012/9/10
   ' 若有審查天數, 新增一筆催審期限的記錄到下一程序檔
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & m_TM01 & "' AND " & _
                  "CF02 = '" & m_TM10 & "' AND " & _
                  "CF03 = '" & m_CP10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CF05")) = False Then
         strNP08 = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
         'Add By Sindy 2023/5/5 FCT重新發文，若下一程序已有該收文號未續辦之催審期限，則更新期限即可，不要另新增期限
         strExc(0) = "SELECT NP01,NP22 from NextProgress" & _
                     " Where NP01='" & m_CP09 & "' and NP07='305' and NP06 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strSql = "UPDATE NextProgress SET NP08=" & PUB_GetWorkDay1(strNP08, True) & ",NP09=" & strNP08 & _
                     " Where NP01='" & m_CP09 & "' and NP07='305' and NP06 is null"
            cnnConnection.Execute strSql
         Else
         '2023/5/5 END
            strNP07 = "305"
            strNP22 = GetNextProgressNo()
            'Modified by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
            'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                     "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                               strNP08 & "," & strNP08 & ",'" & m_CP14 & "'," & strNP22 & ")"
            strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                     "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                               PUB_GetWorkDay1(strNP08, True) & "," & strNP08 & ",'" & m_CP14 & "'," & strNP22 & ")"
            cnnConnection.Execute strSql
         End If
      End If
   End If
   rsTmp.Close
   '2012/9/10 End
   
 '911107 nick transation
  cnnConnection.CommitTrans
  
     'Add by nickc 2008/02/22 檢查代理人Email(需考慮可能為FF案件)
    PUB_CheckEMail m_CP44, m_CP116
    PUB_CheckEMail m_TM44, m_TM119
    If m_TM120 <> "" Then
       PUB_CheckEMail m_TM44, m_TM120
    End If
    'end 2008/02/22
  
     Exit Function
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
     'edit by nick 2004/11/03
     OnSaveData = False
End Function

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   ' 發文日
   If IsEmptyText(textCP27) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入發文日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP27.SetFocus
      GoTo EXITSUB
   End If
   ' 預估勝敗不可為空白
   'If IsEmptyText(textCP23) = True Then
   '   strTit = "檢核資料"
   '   strMsg = "請輸入預估勝敗"
   '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '   GoTo ExitSub
   'End If
   
   'Add By Sindy 2011/01/06
   '外商(S)申請人1或FC代理人至少要輸入一個
   '其他的一定要輸入申請人1
   If m_TM01 = "S" Then
        If textTM23 = "" And m_TM44 = "" Then
            MsgBox "申請人1或FC代理人至少要輸入一個!!!", vbExclamation + vbOKOnly
            GoTo EXITSUB
        End If
   Else
        If textTM23 = "" Then
            MsgBox "申請人1不可空白!!!", vbExclamation + vbOKOnly
            GoTo EXITSUB
        End If
   End If
   
   'Added by Lydia 2021/09/10 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textDN_GotFocus()
   InverseTextBox textDN
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textCP23_GotFocus()
   InverseTextBox textCP23
End Sub

Private Sub textCP27_GotFocus()
   InverseTextBox textCP27
End Sub

Private Sub textCP43_GotFocus()
   InverseTextBox textCP43
End Sub

Private Sub textCP49_GotFocus()
   InverseTextBox textCP49
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
'add by nick 2004/08/13 發文規費，申請國家台灣才檢查
If Me.textCP84.Enabled = True Then
   Cancel = False
   textCP84_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If textCP84.Enabled = True And m_TM10 = "000" Then
    If Val(textCP84.Text) <> Val(m_CP84) Then
        MsgBox "發文規費[" & Trim(Val(m_CP84)) & "] 與實際發文規費[" & Trim(Val(textCP84.Text)) & "]不同", , "警告！"
        textCP84_GotFocus
        Exit Function
    End If
End If

If Me.textCP23.Enabled = True Then
   Cancel = False
   textCP23_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP27.Enabled = True Then
   Cancel = False
   textCP27_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP43.Enabled = True Then
   Cancel = False
   textCP43_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP49.Enabled = True Then
   Cancel = False
   textCP49_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Sindy 98/03/11
If Me.textCP113.Enabled = True Then
   Cancel = False
   textCP113_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'98/03/11 End

If Me.textDN.Enabled = True Then
   Cancel = False
   textDN_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPrint.Enabled = True Then
   Cancel = False
   textPrint_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'add by nickc 2006/01/25
'edit by nickc 2006/02/07
If m_TM01 = "FCT" Then
    If Me.lstNameAgent.Enabled = True Then
        Cancel = False
        lstNameAgent_Validate Cancel
        If Cancel = True Then
            lstNameAgent.SetFocus
            Exit Function
        End If
    End If
End If
TxtValidate = True
End Function

'Add By Sindy 98/03/11
Private Sub textCP113_GotFocus()
   TextInverse textCP113
End Sub
Private Sub textCP113_Validate(Cancel As Boolean)
   If textCP113 <> "" Then
      If Not IsNumeric(textCP113) Then
         MsgBox "請輸入數字！", vbExclamation
         textCP113.SetFocus
         textCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   If GetPrjNation1(textTMKey) = "000" Then
      Cancel = Not PUB_CheckCP113(textCP113, m_TM01, m_CP10, m_CP14)
   End If
End Sub
'98/03/11 End

'Add By Sindy 2012/12/20
Private Sub textCP118_GotFocus()
   TextInverse textCP118
   CloseIme
End Sub

'Add By Sindy 2012/12/20
Private Sub textCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub
