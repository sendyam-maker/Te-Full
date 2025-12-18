VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03010409_03 
   BorderStyle     =   1  '單線固定
   Caption         =   "爭議案部分勝部分敗輸入"
   ClientHeight    =   6108
   ClientLeft      =   96
   ClientTop       =   1008
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6108
   ScaleWidth      =   8952
   Begin VB.TextBox textPrint 
      Height          =   285
      Left            =   5670
      MaxLength       =   1
      TabIndex        =   1
      Top             =   3123
      Width           =   732
   End
   Begin VB.CommandButton Command2 
      Caption         =   "商品及服務資料輸入(&I)"
      Height          =   375
      Left            =   3990
      TabIndex        =   12
      Top             =   60
      Width           =   1965
   End
   Begin VB.TextBox textTM16 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   876
      Width           =   2532
   End
   Begin VB.TextBox textCF15 
      Height          =   285
      Left            =   1230
      MaxLength       =   4
      TabIndex        =   0
      Top             =   3123
      Width           =   732
   End
   Begin VB.TextBox textCF15_2 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   2073
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   3123
      Width           =   1692
   End
   Begin VB.TextBox textCP06 
      Height          =   285
      Left            =   1230
      MaxLength       =   8
      TabIndex        =   2
      Top             =   3444
      Width           =   1095
   End
   Begin VB.TextBox textCP07 
      Height          =   285
      Left            =   5670
      MaxLength       =   8
      TabIndex        =   3
      Top             =   3444
      Width           =   1095
   End
   Begin VB.TextBox textCP26_S 
      Height          =   285
      Left            =   7740
      MaxLength       =   1
      TabIndex        =   8
      Top             =   4095
      Width           =   372
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1230
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   876
      Width           =   2532
   End
   Begin VB.TextBox textTM32 
      Height          =   285
      Left            =   1230
      MaxLength       =   699
      TabIndex        =   9
      Top             =   4470
      Width           =   7575
   End
   Begin VB.TextBox textTM17 
      Height          =   285
      Left            =   4770
      MaxLength       =   1
      TabIndex        =   7
      Top             =   4095
      Width           =   372
   End
   Begin VB.TextBox textCP26 
      Height          =   285
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   6
      Top             =   4095
      Width           =   372
   End
   Begin VB.TextBox textCP14 
      Height          =   285
      Left            =   1230
      MaxLength       =   6
      TabIndex        =   4
      Top             =   3765
      Width           =   732
   End
   Begin VB.TextBox textCP48 
      Height          =   285
      Left            =   5670
      MaxLength       =   8
      TabIndex        =   5
      Top             =   3765
      Width           =   1095
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1230
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   555
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1230
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1839
      Width           =   2532
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2802
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2481
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1230
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2481
      Width           =   2532
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   555
      Width           =   2532
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8070
      TabIndex        =   15
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   375
      Left            =   5985
      TabIndex        =   13
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   375
      Left            =   6840
      TabIndex        =   14
      Top             =   60
      Width           =   1200
   End
   Begin VB.Label Label23 
      Caption         =   "(N:不印)"
      Height          =   255
      Left            =   6540
      TabIndex        =   61
      Top             =   3123
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "列印定稿 :"
      Height          =   255
      Left            =   4710
      TabIndex        =   60
      Top             =   3120
      Width           =   945
   End
   Begin MSForms.TextBox textTM58 
      Height          =   600
      Left            =   1230
      TabIndex        =   11
      Top             =   5430
      Width           =   7575
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13361;1058"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   600
      Left            =   1230
      TabIndex        =   10
      Top             =   4800
      Width           =   7575
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13361;1058"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14_2 
      Height          =   285
      Left            =   2070
      TabIndex        =   59
      Top             =   3765
      Width           =   1695
      VariousPropertyBits=   671105055
      Size            =   "2990;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14_Src 
      Height          =   285
      Left            =   1230
      TabIndex        =   58
      Top             =   2802
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5670
      TabIndex        =   57
      Top             =   2160
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP40 
      Height          =   285
      Left            =   1230
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   1830
      Width           =   2535
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1200
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   1515
      Width           =   7485
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "13203;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1200
      TabIndex        =   54
      Top             =   1200
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
   Begin VB.Label Label12 
      Caption         =   "目前准駁 :"
      Height          =   225
      Left            =   4710
      TabIndex        =   53
      Top             =   897
      Width           =   915
   End
   Begin VB.Label Label22 
      Caption         =   "下一程序 :"
      Height          =   255
      Left            =   150
      TabIndex        =   51
      Top             =   3123
      Width           =   1035
   End
   Begin VB.Label Label18 
      Caption         =   "本所期限 :"
      Height          =   255
      Left            =   150
      TabIndex        =   50
      Top             =   3444
      Width           =   1035
   End
   Begin VB.Label Label25 
      Caption         =   "法定期限 :"
      Height          =   255
      Left            =   4710
      TabIndex        =   49
      Top             =   3444
      Width           =   915
   End
   Begin VB.Label Label14 
      Caption         =   "案件備註 :"
      Height          =   255
      Left            =   150
      TabIndex        =   46
      Top             =   5430
      Width           =   1035
   End
   Begin VB.Label Label13 
      Caption         =   "商品群組 :"
      Height          =   255
      Left            =   150
      TabIndex        =   45
      Top             =   4485
      Width           =   1035
   End
   Begin VB.Label Label11 
      Caption         =   "申請案號 :"
      Height          =   255
      Left            =   150
      TabIndex        =   44
      Top             =   901
      Width           =   1035
   End
   Begin VB.Label Label10 
      Caption         =   "(Y / N)"
      Height          =   255
      Left            =   8160
      TabIndex        =   43
      Top             =   4095
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "是否計算勝訴率 :"
      Height          =   255
      Left            =   6210
      TabIndex        =   42
      Top             =   4095
      Width           =   1455
   End
   Begin VB.Label Label20 
      Caption         =   "(Y / N)"
      Height          =   255
      Left            =   5190
      TabIndex        =   41
      Top             =   4095
      Width           =   615
   End
   Begin VB.Label Label19 
      Caption         =   "專用權是否存在 :"
      Height          =   255
      Left            =   3240
      TabIndex        =   40
      Top             =   4095
      Width           =   1455
   End
   Begin VB.Label Label15 
      Caption         =   "(N:不算)"
      Height          =   255
      Left            =   1860
      TabIndex        =   39
      Top             =   4095
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "是否算案件數 :"
      Height          =   255
      Left            =   150
      TabIndex        =   38
      Top             =   4095
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "承辦人 :"
      Height          =   255
      Left            =   150
      TabIndex        =   37
      Top             =   3765
      Width           =   1035
   End
   Begin VB.Label Label4 
      Caption         =   "對照名稱 :"
      Height          =   255
      Left            =   150
      TabIndex        =   36
      Top             =   1849
      Width           =   1035
   End
   Begin VB.Label Label21 
      Caption         =   "進度備註 :"
      Height          =   255
      Left            =   150
      TabIndex        =   35
      Top             =   4785
      Width           =   1035
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   255
      Left            =   150
      TabIndex        =   34
      Top             =   2802
      Width           =   1035
   End
   Begin VB.Label Label26 
      Caption         =   "承辦期限 :"
      Height          =   255
      Left            =   4710
      TabIndex        =   33
      Top             =   3765
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   32
      Top             =   585
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   150
      TabIndex        =   31
      Top             =   1217
      Width           =   1035
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   255
      Left            =   150
      TabIndex        =   30
      Top             =   1533
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   29
      Top             =   2165
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   255
      Index           =   6
      Left            =   4710
      TabIndex        =   28
      Top             =   1863
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   255
      Index           =   8
      Left            =   4710
      TabIndex        =   27
      Top             =   2802
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   255
      Index           =   9
      Left            =   4710
      TabIndex        =   26
      Top             =   2487
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   255
      Index           =   10
      Left            =   150
      TabIndex        =   25
      Top             =   2481
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   255
      Index           =   11
      Left            =   4710
      TabIndex        =   24
      Top             =   2175
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "審定號 :"
      Height          =   255
      Left            =   4710
      TabIndex        =   23
      Top             =   585
      Width           =   915
   End
End
Attribute VB_Name = "frm03010409_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/16 改成Form2.0 ;textTM23、cmbTM05、textCP13、textCP40、textCP14_Src、textCP14_2、textCP64、textTM58
'CREATE BY SINDY 2014/8/7
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 申請國家
Dim m_TM10 As String
Dim m_TM09 As String
' 來函收文日
Dim m_CP05 As String
' 原本所期限
Dim m_CP06 As String
' 機關文號
Dim m_CP08 As String
' 所選取的收文號
Dim m_CP09 As String
' 案件性質
Dim m_CP10 As String
' 智權人員
Dim m_CP13 As String
Dim m_CP12 As String
' 對照號數
Dim m_CP36 As String
' 對照案件名稱(中)
Dim m_CP37 As String
' 對照案件名稱(英)
Dim m_CP38 As String
' 對照案件名稱(日)
Dim m_CP39 As String
' 對照名稱(中)
Dim m_cp40 As String
' 對照名稱(英)
Dim m_CP41 As String
' 對照名稱(日)
Dim m_CP42 As String
' 預估結果
Dim m_CP23 As String

Dim m_intNumBegin As Integer
Dim m_intNumEnd As Integer

'檢查是否已經有商品及服務
Public ChkTG As Boolean
Dim BolPrintCaseCheck As Boolean
Dim m_TM28 As String 'Add By Sindy 2015/6/3
'add by sonia 2017/10/16
Dim IsAppNpData As Boolean
Dim SeekNewCp09 As String
Dim oClsPrtForm001 As New ClsPrtForm001
'end 2017/10/16
'Add By Sindy 2023/5/3
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2023/5/3 END


'Add By Sindy 2023/5/3
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frm03010409_02.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm03010409_02
   Unload frm03010409_01
   Unload Me
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   BolPrintCaseCheck = CaseCheck(m_TM01, m_TM02, m_TM03, m_TM04, m_TM10)
   
   If CheckDataValid() = True Then
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      
      'Added by Lydia 2021/12/08
      ' 列印定稿
      If textPrint <> "N" Then
         PrintLetter
      End If
      'end 2021/12/08
      
      'add by sonia 2017/10/16
      If IsAppNpData Then
         If MsgBox("準備列印回覆單!!!", vbExclamation + vbOKCancel) = vbOK Then
            Call oClsPrtForm001.PrintReturnSheet(SeekNewCp09, textCF15, DBDATE(textCP07), False)
         End If
      End If
      'end 2017/10/16
      
      ' 設定滑鼠游標為預設
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      'Add By Sindy 2023/5/3
      If Me.m_strIR01 <> "" Then
         Unload frm03010409_02
         Unload frm03010409_01
         If Not m_PrevForm Is Nothing Then
            Call m_PrevForm.GoNext
         End If
         Unload Me
      Else
      '2023/5/3 END
         Unload frm03010409_02
         Unload Me
         frm03010409_01.Show
      End If
   End If
End Sub

Private Sub Command2_Click()
   frm03010303_04.Hide
   Set frm03010303_04.UpForm = Me
   frm03010303_04.TGKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   frm03010303_04.AllClass = m_TM09
   frm03010303_04.cmdOK(2).Visible = True
   
   If m_TM09 <> "" Then  '有商品類別才可進入 T-113511團體標章
      Me.Hide
      frm03010303_04.QueryData
      frm03010303_04.Show vbModal '強制回應表單
   Else
      MsgBox ("無商品類別，不可使用此按鈕 !")
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM16.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   
   textCP05S.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14_Src.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCP40.BackColor = &H8000000F
   textCF15_2.BackColor = &H8000000F
   
   'Added by Lydia 2021/12/08 CFT預設出通用定稿
   If m_TM01 = "CFT" Then
      textPrint = ""
   Else
      textPrint = "N"
   End If
   'end 2021/12/08
   
   MoveFormToCenter Me
   
   'Add By Sindy 2023/5/3
   m_strIR01 = frm03010409_02.m_strIR01
   m_strIR02 = frm03010409_02.m_strIR02
   m_strIR03 = frm03010409_02.m_strIR03
   m_strIR04 = frm03010409_02.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2023/5/3 END
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP05 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_TM01 = strData
      ' 本所案號 欄位2
      Case 1: m_TM02 = strData
      ' 本所案號 欄位3
      Case 2: m_TM03 = strData
      ' 本所案號 欄位4
      Case 3: m_TM04 = strData
      ' 來函收文日
      Case 4: m_CP05 = strData
      ' 收文號
      Case 5: m_CP09 = strData
   End Select
End Sub

' 取得商標基本檔
Private Sub QueryTradeMark()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         textTM10 = GetNationName(rsTmp.Fields("TM10"), 0)
      End If
      '商品類別
      m_TM09 = ""
      If IsNull(rsTmp.Fields("TM09")) = False Then
         m_TM09 = rsTmp.Fields("TM09")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      ' 審定號
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
      End If
      ' 目前准駁
      If IsNull(rsTmp.Fields("TM16")) = False Then
         Select Case rsTmp.Fields("TM16")
            Case 1: textTM16 = "准"
            Case 2: textTM16 = "駁"
         End Select
      End If
      ' 專用權是否存在
      If IsNull(rsTmp.Fields("TM17")) = False Then
         textTM17 = rsTmp.Fields("TM17")
      End If
      m_TM28 = "" 'Add By Sindy 2015/6/3
      If IsNull(rsTmp.Fields("TM28")) = False Then
         m_TM28 = rsTmp.Fields("TM28") 'Add By Sindy 2015/6/3
         If rsTmp.Fields("TM28") <> "1" Then
            textTM17 = "N"
         End If
      End If
      ' 彼所案號
      If IsNull(rsTmp.Fields("TM45")) = False Then
         textTM45 = rsTmp.Fields("TM45")
      End If
      ' 商品組群
      If IsNull(rsTmp.Fields("TM32")) = False Then
         textTM32 = rsTmp.Fields("TM32")
      End If
      ' 案件備註
      If IsNull(rsTmp.Fields("TM58")) = False Then
         textTM58 = rsTmp.Fields("TM58")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
   Dim strDay As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim bCP40 As Boolean
   Dim strTemp As String
   
   m_TM10 = Empty
   m_CP13 = Empty
   m_CP12 = Empty
   
   m_CP36 = Empty
   m_CP37 = Empty
   m_CP38 = Empty
   m_CP39 = Empty
   m_cp40 = Empty
   m_CP41 = Empty
   m_CP42 = Empty
   m_CP23 = Empty
   
   ' 本所案號
   textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   ' 來函收文日
   textCP05S = m_CP05
   ' 收文號
   textCP09 = m_CP09
   
   ' 先設定承辦期限為可輸入
   textCP48.BackColor = &H80000005
   textCP48.Locked = False
   textCP48.TabStop = True
   
   ' 取得商標基本檔的相關項目
   QueryTradeMark
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 取得案件進度檔
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' AND " & _
                  "CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 原本所期限
      If IsNull(rsTmp.Fields("CP06")) = False Then
         m_CP06 = rsTmp.Fields("CP06")
      End If
      ' 機關文號
      If IsNull(rsTmp.Fields("CP08")) = False Then
         m_CP08 = rsTmp.Fields("CP08")
      End If
      ' 案件性質
      m_CP10 = Empty
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"), True) 'Modified by Lydia 2016/03/25 離職人員也顯示
      End If
      '業務區
      If IsNull(rsTmp.Fields("cp12")) = False Then
        m_CP12 = rsTmp.Fields("cp12")
      End If
      ' 承辦人員
      'Modified by Lydia 2016/03/11 CFT改成模組判斷
      'Modified by Lydia 2016/03/25 全部套用
      'If m_TM01 = "CFT" Then
        Dim strNA69 As String
        'Modified by Lydia 2016/08/17 改抓NA69
        'Call GetNP69("", "", "", strNA69, m_TM01, m_TM02, m_TM03, m_TM04)
        'Modified by Lydia 2016/11/21 傳入智權人員代號
        'Call GetNP69("", m_TM10, "", strNA69, m_TM01, m_TM02, m_TM03, m_TM04)
        'Modified by Lydia 2017/05/12 GetNP69更名為GetNA69
        Call GetNA69("", m_TM10, "" & rsTmp.Fields("CP13"), strNA69, m_TM01, m_TM02, m_TM03, m_TM04)
        
        textCP14_Src = GetStaffName("" & rsTmp.Fields("CP14"), True)
        textCP14 = strNA69
        textCP14_2 = GetStaffName(strNA69)
'      Else
'      'end 2016/03/11
'        If IsNull(rsTmp.Fields("CP14")) = False Then
'           textCP14 = rsTmp.Fields("CP14")
'           textCP14_Src = GetStaffName(rsTmp.Fields("CP14"))
'           textCP14_2 = GetStaffName(rsTmp.Fields("CP14"))
'        End If
'      End If
      'end 2016/03/25
      ' 對照名稱 (無中文取英文, 無英文取日文)
      bCP40 = False
      If bCP40 = False Then
         If IsNull(rsTmp.Fields("CP40")) = False Then
            If IsEmptyText(rsTmp.Fields("CP40")) = False Then
               textCP40 = rsTmp.Fields("CP40")
               bCP40 = True
            End If
         End If
      End If
      If bCP40 = False Then
         If IsNull(rsTmp.Fields("CP41")) = False Then
            If IsEmptyText(rsTmp.Fields("CP41")) = False Then
               textCP40 = rsTmp.Fields("CP41")
               bCP40 = True
            End If
         End If
      End If
      If bCP40 = False Then
         If IsNull(rsTmp.Fields("CP42")) = False Then
            If IsEmptyText(rsTmp.Fields("CP42")) = False Then
               textCP40 = rsTmp.Fields("CP42")
               bCP40 = True
            End If
         End If
      End If
      ' 預估結果
      If IsNull(rsTmp.Fields("CP23")) = False Then
         m_CP23 = rsTmp.Fields("CP23")
      End If
      ' 程式存檔用資料
      ' 對造號數
      If IsNull(rsTmp.Fields("CP36")) = False Then
         m_CP36 = rsTmp.Fields("CP36")
      End If
      ' 對造案件名稱(中)
      If IsNull(rsTmp.Fields("CP37")) = False Then
         m_CP37 = rsTmp.Fields("CP37")
      End If
      ' 對造案件名稱(英)
      If IsNull(rsTmp.Fields("CP38")) = False Then
         m_CP38 = rsTmp.Fields("CP38")
      End If
      ' 對造案件名稱(日)
      If IsNull(rsTmp.Fields("CP39")) = False Then
         m_CP39 = rsTmp.Fields("CP39")
      End If
      ' 對造名稱(中)
      If IsNull(rsTmp.Fields("CP40")) = False Then
         m_cp40 = rsTmp.Fields("CP40")
      End If
      ' 對造名稱(英)
      If IsNull(rsTmp.Fields("CP41")) = False Then
         m_CP41 = rsTmp.Fields("CP41")
      End If
      ' 對造名稱(日)
      If IsNull(rsTmp.Fields("CP42")) = False Then
         m_CP42 = rsTmp.Fields("CP42")
      End If
   End If
   rsTmp.Close
   
   ' 承辦期限(若計算結果超過本所期限), 則設定為本所期限且不可輸入
   strTemp = GetCP48()
   If IsEmptyText(strTemp) = False And IsEmptyText(m_CP06) = False Then
      If Val(strTemp) > Val(m_CP06) Then
         textCP48 = m_CP06
         textCP48.BackColor = &H8000000F
         textCP48.Locked = True
         textCP48.TabStop = False
      End If
   End If
   
   ' 非A類收文其預設為不可算案件數
   textCP26 = "N"
   
   Set rsTmp = Nothing
End Sub

' 取得承辦期限
Private Function GetCP48() As String
   GetCP48 = Empty
   ' 承辦期限的日期應為來函收文日加上工作天數
   ' 工作天數由系統別+國家代碼+案件性質(勝訴)搜尋案件收費表的工作天數
   ' 若有值才做檢查
   If IsEmptyText(textCP48) = False Then
      textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1003", DBDATE(m_CP05), DBDATE(textCP06), textCP09)
   End If
End Function

Public Function OnSaveData() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   Dim strSql As String
   Dim strSubTMSQL As String
   Dim bUpdate As Boolean
   Dim strCP09 As String
   Dim strCP10 As String
   Dim strCP12 As String
   Dim strCP27 As String
   Dim strNP07 As String
   Dim strNP14 As String
   Dim strNP22 As String
   Dim strCP06 As String
   Dim strCP07 As String
   Dim oClsPrtForm001 As New ClsPrtForm001
   
On Error GoTo ErrorHandler
   
OnSaveData = True
cnnConnection.BeginTrans
   
   strSubTMSQL = "WHERE TM01 = '" & m_TM01 & "' AND " & _
                       "TM02 = '" & m_TM02 & "' AND " & _
                       "TM03 = '" & m_TM03 & "' AND " & _
                       "TM04 = '" & m_TM04 & "' "
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 判斷是否更新實際結果 (無實際結果才更新)
   bUpdate = True
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CP24")) = False Then
         If IsEmptyText(rsTmp.Fields("CP24")) = False Then
            bUpdate = False
         End If
      End If
   End If
   rsTmp.Close
   ' 更新原案件進度檔的收文資料其實際結果為准, 准駁日為來函收文日
   'Modify By Sindy 2024/11/1 2024/8/2進度維護已開放可以輸入為3=部分勝敗
   '                          此處也要更新為3
   If bUpdate = True Then
      strSql = "UPDATE CaseProgress SET CP24 = '3'," & _
                                       "CP25 = " & DBDATE(m_CP05) & _
               "WHERE CP09 = '" & m_CP09 & "'"
      cnnConnection.Execute strSql
   End If
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 依是否計算勝訴率來更新原案件進度檔資料的是否算案件數欄位
   Select Case textCP26_S
      Case "Y":
         strSql = "UPDATE CaseProgress SET CP26 = NULL " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
         cnnConnection.Execute strSql
      Case "N":
         strSql = "UPDATE CaseProgress SET CP26 = 'N' " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
         cnnConnection.Execute strSql
   End Select
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新商標基本檔的專用權是否存在, 商品組群, 案件備註欄
   strSql = "UPDATE TradeMark SET TM17 = '" & textTM17 & "', " & _
                                 "TM32 = '" & textTM32 & "', " & _
                                 "TM58 = '" & ChgSQL(textTM58) & "' "
   strSql = strSql & strSubTMSQL
   cnnConnection.Execute strSql
   
   ' 新增一筆資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   
   strCP06 = Empty
   strCP07 = Empty
   If IsEmptyText(textCP06) = False Then: strCP06 = DBDATE(textCP06)
   If IsEmptyText(textCP07) = False Then: strCP07 = DBDATE(textCP07)
   ' 案件性質為部分勝部分敗
   strCP10 = "1006"
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43,CP48,CP64,CP06,CP07) " & _
                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                          "'" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & textCP14 & "'," & _
                          "'" & "N" & "','" & textCP26 & "','" & "N" & "'," & _
                          "'" & m_CP36 & "','" & ChgSQL(m_CP37) & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
                          "'" & m_cp40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "'," & DBDATE(textCP48) & "," & _
                          "'" & ChgSQL(Trim(textCP64)) & "'," & CNULL(strCP06) & "," & CNULL(strCP07) & ")"
   cnnConnection.Execute strSql
   
   Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新下一程序檔中下一程序為催審且將其是否續辦欄位設為Y
   ' 更新下一程序檔案件性質為997.收達998.提申的資料
   strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
            "WHERE NP01 = '" & m_CP09 & "' AND " & _
                  "NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "' AND " & _
                  "NP07 in (305,997,998) "
   cnnConnection.Execute strSql
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'Added by Lydia 2025/09/12 TF基礎案號設定：基礎案狀態通知Email
   Dim strTFcase As String
   If m_TM01 = "CFT" Then
      strTFcase = PUB_GetTFbaseInfo(m_TM01, m_TM02, m_TM03, m_TM04, textTM15, m_TM10, "2", textTM12, strCP09)
   End If
   'end 2025/09/12
   
   ' 有輸入下一程序時
   If IsEmptyText(textCF15) = False Then
      strNP22 = GetNextProgressNo()
      strNP14 = Empty
      strNP14 = GetRelatedPerson(m_CP09)
      '智權人員存最近收文A類接洽記錄單的智權人員
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP14,NP22) " & _
               "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
                        DBDATE(textCP06) & "," & DBDATE(textCP07) & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & ChgSQL(strNP14) & "'," & strNP22 & ")"
      cnnConnection.Execute strSql
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
      Select Case textCF15
         Case "102", "105", "702", "708", "305", "998", "997":
         Case Else:
            'add by sonia 2017/10/16
            IsAppNpData = True
            SeekNewCp09 = strCP09
            'end 2017/10/16
            '改成整批列印
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
'2017/10/16原已不印,陳金蓮來電通知仍應列印,故改在COMMIT之後再印
'            If MsgBox("準備列印回覆單!!!", vbExclamation + vbOKCancel) = vbOK Then
'               Call oClsPrtForm001.PrintReturnSheet(strCP09, textCF15, DBDATE(textCP07), False)
'            End If
      End Select
   End If
   
   'Add by Sindy 2023/5/3
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm03010409_01", strCP09
   End If
   '2023/5/3 END
   
   Set rsTmp = Nothing
   cnnConnection.CommitTrans
   Exit Function
   
ErrorHandler:
   cnnConnection.RollbackTrans
   OnSaveData = False
End Function

Private Function CheckDataValid()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   
   ' 申請國家為台灣時且案件性質非(行政訴訟,行政訴訟上訴,行政上訴答辯), 下一程序不可為空白
   If m_TM10 < "010" And m_CP10 <> "403" And m_CP10 <> "408" And m_CP10 <> "410" Then
      If IsEmptyText(textCF15) = True Then
         strTit = "檢核資料"
         strMsg = "申請國家為台灣時, 下一程序不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF15.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   ' 有輸入下一程序時, 本所期限與法定期限不可為空白
   If IsEmptyText(textCF15) = False Then
      If IsEmptyText(textCP06) = True Or IsEmptyText(textCP07) = True Then
         strTit = "檢核資料"
         strMsg = "有下一程序時, 本所期限與法定期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
      If Me.textCP06.Text <> "" Then
         If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
            MsgBox "本所期限不可小於系統日期!!!", vbExclamation
            Me.textCP06.SetFocus
            textCP06_GotFocus
            GoTo EXITSUB
         End If
      End If
      ' 本所期限必須小於法定期限
      If Val(textCP06) > Val(textCP07) Then
         strTit = "檢核資料"
         strMsg = "本所期限的日期不可超過法定期限的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
   Else
      strTit = "檢核資料"
      strMsg = "下一程序不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCF15.SetFocus
      GoTo EXITSUB
   End If
   
   ' 承辦期限不可空白
   If IsEmptyText(textCP48) = True Then
      strTit = "檢核資料"
      strMsg = "承辦期限不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP48.SetFocus
      GoTo EXITSUB
   End If
   '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
   If Len(Me.textCP06.Text) > 0 And Len(Me.textCP48.Text) > 0 Then
      If Val(Me.textCP06.Text) < Val(Me.textCP48.Text) Then
         strTit = "檢核資料"
         strMsg = "承辦期限不得大於本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48.SetFocus
         GoTo EXITSUB
      End If
   End If
   '承辦期限不可小於系統日
   If Val(DBDATE(Me.textCP48.Text)) < strSrvDate(1) Then
      strTit = "檢核資料"
      strMsg = "承辦期限不可小於系統日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP48.SetFocus
      GoTo EXITSUB
   End If
   
   ' 是否計算勝訴率不可空白
   If Not IsEmptyText(m_CP23) Then
      If IsEmptyText(textCP26_S) = True Then
         strTit = "檢核資料"
         strMsg = "是否計算勝訴率不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP26_S.SetFocus
         GoTo EXITSUB
      End If
   End If
   ' 專用權是否存在不可為空白
   'Modify By Sindy 2015/6/3
   'If IsEmptyText(textTM17) = True Then
   If IsEmptyText(textTM17) = True And m_TM28 = "1" Then
   '2015/6/3 END
      strTit = "檢核資料"
      strMsg = "專用權是否存在不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM17.SetFocus
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub Form_Unload(Cancel As Integer)
   Call PUB_SendMailCache 'Added by Lydia 2025/09/12
   
   Set frm03010409_03 = Nothing
End Sub

' 下一程序
Private Sub textCF15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCF15) = False Then
      ' 只取得國內的案件性質名稱
      If m_TM10 < "010" Then
         textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 0)
      Else
         textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 1)
      End If
      If IsEmptyText(textCF15_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件性質代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF15_GotFocus
      End If
   End If
End Sub

' 本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP06) = False Then
      ' 檢查是否為西元日期
      If CheckIsDate(textCP06, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06_GotFocus
         Exit Sub
      'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textCP06.Text = PUB_GetWorkDay1(textCP06, True)
      'end 2020/07/09
      End If
      'Add By Cheng 2002/03/11
      'Modify By Sindy 2009/09/18
      'If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
      If Val(Me.textCP06.Text) < ServerDate Then
         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
         Cancel = True
         textCP06_GotFocus
         Exit Sub
      End If
   End If
End Sub

' 法定期限
Private Sub textCP07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP07) = False Then
      ' 檢查是否為民國年
      If CheckIsDate(textCP07, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的法定期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07_GotFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub textCP14_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 承辦人
Private Sub textCP14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   
   Cancel = False
   textCP14_2 = Empty
   If IsEmptyText(textCP14) = False Then
      textCP14_2 = GetStaffName(textCP14)
      If IsEmptyText(textCP14_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "承辦人代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP14_GotFocus
      End If
   End If
End Sub

Private Sub textCP26_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP26_S_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否計算勝訴率
Private Sub textCP26_S_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP26_S) = False Then
      Select Case textCP26_S
         Case "Y", "N"
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入Y或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP26_S_GotFocus
      End Select
   End If
End Sub

' 是否算案件數
Private Sub textCP26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP26) = False Then
      Select Case textCP26
         Case " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP26_GotFocus
      End Select
   End If
End Sub

' 承辦期限
Private Sub textCP48_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP48) = False Then
      ' 檢查是否為民國年
      If CheckIsDate(textCP48, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的承辦期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48_GotFocus
         Exit Sub
      End If
   End If
End Sub

' 進度備註
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP64, 2020) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "進度備註欄位內容太長"
      textCP64_GotFocus
   End If
End Sub

Private Sub textTM17_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 案件備註
Private Sub textTM58_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM58, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件備註欄位內容太長"
      textTM58_GotFocus
   End If
End Sub

' 專用權是否存在
Private Sub textTM17_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM17) = False Then
      Select Case textTM17
         Case "Y", "N"
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入Y或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM17_GotFocus
      End Select
   End If
End Sub

' 商品組群
Private Sub textTM32_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim strTemp As String
   Dim nResponse
   Dim lst() As String
   Dim lstCount As Integer
   Dim nIndex As Integer
   Dim bFind As Boolean
   Dim nPos As Integer
   
   lstCount = 0
   
   Cancel = False
   If IsEmptyText(textTM32) = False Then
      ' 檢查欄位是否太長
      If CheckLengthIsOK(textTM32, textTM32.MaxLength) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "商品組群欄位內容太長"
         textTM32_GotFocus
      End If
      
      For nIndex = 1 To GetSubStringCount(textTM32)
         strTemp = GetSubString(textTM32, nIndex)
         
         If IsEmptyText(strTemp) = True Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "商品組群的資料不可有空白的內容"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM32_GotFocus
            GoTo EXITSUB
         End If
         
         bFind = False
         For nPos = 0 To lstCount - 1
            If lst(nPos) = strTemp Then
               bFind = True
               Exit For
            End If
         Next nPos
         
         If bFind = True Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "商品組群的資料不可重覆"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM32_GotFocus
            GoTo EXITSUB
         Else
            ReDim Preserve lst(lstCount + 1)
            lst(lstCount) = strTemp
            lstCount = lstCount + 1
         End If
      Next nIndex
   End If
   
EXITSUB:
   If lstCount > 0 Then: Erase lst
   lstCount = 0
End Sub

Private Sub textCF15_GotFocus()
   InverseTextBox textCF15
End Sub

Private Sub textCP06_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP07
End Sub

Private Sub textTM17_GotFocus()
   InverseTextBox textTM17
End Sub

Private Sub textTM32_GotFocus()
   InverseTextBox textTM32
End Sub

Private Sub textTM58_GotFocus()
   InverseTextBox textTM58
End Sub

Private Sub textCP14_GotFocus()
   InverseTextBox textCP14
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

Private Sub textCP26_S_GotFocus()
   InverseTextBox textCP26_S
End Sub

Private Sub textCP48_GotFocus()
   InverseTextBox textCP48
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse

   TxtValidate = False
   If Me.textCP14.Enabled = True Then
      Cancel = False
      textCP14_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP26.Enabled = True Then
      Cancel = False
      textCP26_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP26_S.Enabled = True Then
      Cancel = False
      textCP26_S_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP48.Enabled = True Then
      Cancel = False
      textCP48_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP64.Enabled = True Then
      Cancel = False
      textCP64_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP06.Enabled = True Then
      Cancel = False
      textCP06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP07.Enabled = True Then
      Cancel = False
      textCP07_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM17.Enabled = True Then
      Cancel = False
      textTM17_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM32.Enabled = True Then
      Cancel = False
      textTM32_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM58.Enabled = True Then
      Cancel = False
      textTM58_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   ' 申請國家為台灣時需檢查來函記錄檔
   If m_TM10 < "010" Then
      strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, DBDATE(m_CP05), "MR16")
      If IsEmptyText(strDate) = False Then
         If TAIWANDATE(textCP06) <> TAIWANDATE(strDate) Then
            strTit = "資料檢核"
            strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
            If nResponse = vbCancel Then
               Cancel = True
               textCP06_GotFocus
               Exit Function
            End If
         End If
      Else
        If IsMailRecNoTermExist(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05) = False Then
        Else
           If TAIWANDATE(textCP06) <> TAIWANDATE(strDate) Then
              strTit = "資料檢核"
              strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
              nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
              If nResponse = vbCancel Then
                 Cancel = True
                 textCP06_GotFocus
                 Exit Function
              End If
           End If
        End If
      End If
   End If
   ' 申請國家為台灣時需檢查來函記錄檔
   If m_TM10 < "010" Then
      strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, DBDATE(m_CP05), "MR17")
      If IsEmptyText(strDate) = False Then
         If TAIWANDATE(textCP07) <> TAIWANDATE(strDate) Then
            strTit = "資料檢核"
            strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
            If nResponse = vbCancel Then
               Cancel = True
               textCP07_GotFocus
               Exit Function
            End If
         End If
      Else
         If IsMailRecNoTermExist(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05) = False Then  '2011/6/15 ADD BY SONIA
            strTit = "資料檢核"
            strMsg = "來函記錄中無該筆記錄"
            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
            If nResponse = vbCancel Then
               Cancel = True
               textCP07_GotFocus
               Exit Function
            End If
         Else
            If TAIWANDATE(textCP07) <> TAIWANDATE(strDate) Then
               strTit = "資料檢核"
               strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
               If nResponse = vbCancel Then
                  Cancel = True
                  textCP07_GotFocus
                  Exit Function
               End If
            End If
         End If
      End If
   End If
   
    'Added by Lydia 2021/12/08
    If Me.textPrint.Enabled = True Then
       Cancel = False
       textPrint_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    
    'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
         Exit Function
    End If
    
   TxtValidate = True
End Function

'Added by Lydia 2021/12/08
Private Sub textPrint_GotFocus()
   TextInverse textPrint
End Sub

'Added by Lydia 2021/12/08
' 是否列印定稿
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         Case "", " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

'Added by Lydia 2021/12/08
' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
   If m_TM01 = "CFT" Then
      ' 清除定稿例外欄位檔原有資料(定稿別 /收文號或本所案號000 /處理狀況 /使用者編號)
      EndLetter "04", m_CP09, "00", strUserNum
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   'Added by Lydia 2021/12/08 CFT 沒設定都出通用定稿(不列印)
   If m_TM01 = "CFT" Then
      NowPrint m_CP09, "04", "00", False, strUserNum, 0, , , , , , , , , , , , , , , , , True
   End If
End Sub

