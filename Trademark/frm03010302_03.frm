VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03010302_03 
   BorderStyle     =   1  '單線固定
   Caption         =   "審查報告輸入"
   ClientHeight    =   5748
   ClientLeft      =   -3420
   ClientTop       =   4356
   ClientWidth     =   9156
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9156
   Begin VB.TextBox textPermit 
      Height          =   264
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   6
      Top             =   2940
      Width           =   372
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   552
      Width           =   2532
   End
   Begin VB.TextBox textPrint 
      Height          =   264
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   7
      Top             =   4680
      Width           =   732
   End
   Begin VB.TextBox textCP26 
      Height          =   264
      Left            =   6060
      MaxLength       =   1
      TabIndex        =   8
      Top             =   4680
      Width           =   372
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6888
      TabIndex        =   10
      Top             =   70
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5940
      TabIndex        =   9
      Top             =   70
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8160
      TabIndex        =   11
      Top             =   70
      Width           =   912
   End
   Begin VB.TextBox textCP08 
      Height          =   264
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   0
      Top             =   2040
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1740
      Width           =   2355
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2532
   End
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1140
      Width           =   2532
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
   End
   Begin VB.TextBox textCF15 
      Height          =   264
      Left            =   5700
      MaxLength       =   4
      TabIndex        =   1
      Top             =   2040
      Width           =   732
   End
   Begin VB.TextBox textCF15_2 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   6540
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1692
   End
   Begin VB.TextBox textCP06 
      Height          =   264
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   2
      Top             =   2340
      Width           =   2532
   End
   Begin VB.TextBox textCP14 
      Height          =   264
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   4
      Top             =   2640
      Width           =   732
   End
   Begin VB.TextBox textCP07 
      Height          =   264
      Left            =   5700
      MaxLength       =   8
      TabIndex        =   3
      Top             =   2340
      Width           =   2532
   End
   Begin VB.TextBox textCP48 
      Height          =   264
      Left            =   5700
      MaxLength       =   8
      TabIndex        =   5
      Top             =   2640
      Width           =   2532
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   1080
      Left            =   120
      TabIndex        =   43
      Top             =   3555
      Width           =   8910
      _ExtentX        =   15706
      _ExtentY        =   1905
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
   Begin MSForms.TextBox textCP64 
      Height          =   525
      Left            =   1200
      TabIndex        =   49
      Top             =   5040
      Width           =   7815
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13785;926"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5700
      TabIndex        =   48
      Top             =   1740
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
   Begin MSForms.TextBox textCP14_2 
      Height          =   285
      Left            =   1980
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2535
      VariousPropertyBits=   671105055
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
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1110
      Width           =   2565
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "4524;503"
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
      TabIndex        =   45
      Top             =   810
      Width           =   7875
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13891;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblClose 
      Caption         =   "lblClose"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3750
      TabIndex        =   44
      Top             =   570
      Width           =   645
   End
   Begin VB.Label Label10 
      Caption         =   "(Y / N)"
      Height          =   225
      Left            =   2040
      TabIndex        =   42
      Top             =   2970
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "(N:不算)"
      Height          =   252
      Left            =   6540
      TabIndex        =   41
      Top             =   4680
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "是否為核駁 :"
      Height          =   252
      Left            =   120
      TabIndex        =   40
      Top             =   2940
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號 :"
      Height          =   252
      Index           =   13
      Left            =   4740
      TabIndex        =   39
      Top             =   540
      Width           =   852
   End
   Begin VB.Label Label21 
      Caption         =   "進度備註 :"
      Height          =   252
      Left            =   120
      TabIndex        =   37
      Top             =   5040
      Width           =   972
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   252
      Left            =   120
      TabIndex        =   36
      Top             =   4680
      Width           =   972
   End
   Begin VB.Label Label23 
      Caption         =   "(N:不印)"
      Height          =   252
      Left            =   2040
      TabIndex        =   35
      Top             =   4680
      Width           =   852
   End
   Begin VB.Label Label16 
      Caption         =   "是否算案件數 :"
      Height          =   252
      Left            =   4740
      TabIndex        =   34
      Top             =   4680
      Width           =   1212
   End
   Begin VB.Label Label15 
      Caption         =   "(N:不算)"
      Height          =   252
      Left            =   6480
      TabIndex        =   33
      Top             =   5040
      Width           =   972
   End
   Begin VB.Label Label7 
      Caption         =   "本案期限 :"
      Height          =   252
      Left            =   120
      TabIndex        =   32
      Top             =   3240
      Width           =   852
   End
   Begin VB.Label Label8 
      Caption         =   "機關文號 :"
      Height          =   252
      Left            =   120
      TabIndex        =   31
      Top             =   2040
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4740
      TabIndex        =   30
      Top             =   1740
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   29
      Top             =   1740
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   28
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   252
      Index           =   3
      Left            =   4740
      TabIndex        =   27
      Top             =   1140
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "商標種類 :"
      Height          =   252
      Index           =   2
      Left            =   4740
      TabIndex        =   26
      Top             =   1440
      Width           =   852
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   25
      Top             =   1140
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   24
      Top             =   840
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   540
      Width           =   852
   End
   Begin VB.Label Label4 
      Caption         =   "下一程序 :"
      Height          =   252
      Left            =   4740
      TabIndex        =   22
      Top             =   2040
      Width           =   852
   End
   Begin VB.Label Label5 
      Caption         =   "本所期限 :"
      Height          =   252
      Left            =   120
      TabIndex        =   21
      Top             =   2340
      Width           =   852
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   852
   End
   Begin VB.Label Label25 
      Caption         =   "法定期限 :"
      Height          =   252
      Left            =   4740
      TabIndex        =   19
      Top             =   2340
      Width           =   852
   End
   Begin VB.Label Label26 
      Caption         =   "承辦期限 :"
      Height          =   252
      Left            =   4740
      TabIndex        =   18
      Top             =   2640
      Width           =   852
   End
End
Attribute VB_Name = "frm03010302_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/12 改成Form2.0 ;textTM23、cmbTM05、textCP13、textCP14_2、textCP64、grdList改字型=新細明體-ExtB
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
' 來函收文日
Dim m_CP05 As String
' 收文號
Dim m_CP09 As String
' 原案件性質
Dim m_CP10 As String
' 原業務區
Dim m_CP12 As String
' 原智權人員代號
Dim m_CP13 As String
' 原授權期間(迄)
Dim m_CP54 As String
' 原移轉申請人代號
Dim m_CP56 As String
' 商標種類代碼
Dim m_TM08 As String
' 國家代碼
Dim m_TM10 As String
' 原專用期限起日
Dim m_TM21 As String
' 原專用期限止日
Dim m_TM22 As String
' 原申請人代號
Dim m_TM23 As String
' 申請國家的延展年度
Dim m_NA14 As Integer
'Add By Sindy 2009/06/22
Dim m_CP24 As String

Dim m_CurrSel As Integer
'add by nickc 2005/05/31
Dim IsAppNpData As Boolean
Dim SeekNewCp09 As String
Dim oClsPrtForm001 As New ClsPrtForm001
'Add By Sindy 2023/4/27
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2023/4/27 END


'Add By Sindy 2023/4/27
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

' 原資料是否有實際結果
Private Sub cmdCancel_Click()
   Unload Me
   frm03010302_02.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm03010302_02
   Unload frm03010302_01
   Unload Me
End Sub

Private Sub cmdok_Click()
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
   
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 存檔
      'edit by nick 2004/11/03
      'OnSaveData
   'add by nickc 2005/05/31
   IsAppNpData = False
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub

      'Added by Morgan 2021/12/7
      ' 列印定稿
      If textPrint <> "N" Then
         PrintLetter
      End If
      'end 2021/12/7
   
      'add by nickc 2005/05/31
      If IsAppNpData Then
         'add by nickc 2005/09/27
         If MsgBox("準備列印回覆單!!!", vbExclamation + vbOKCancel) = vbOK Then
            Call oClsPrtForm001.PrintReturnSheet(SeekNewCp09, textCF15, DBDATE(textCP07), False)
         End If
      End If

      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      'Add By Sindy 2023/4/27
      If Me.m_strIR01 <> "" Then
         Unload frm03010302_02
         Unload frm03010302_01
         If Not m_PrevForm Is Nothing Then
            Call m_PrevForm.GoNext
         End If
         Unload Me
      Else
      '2023/4/27 END
         Unload Me
         Unload frm03010302_02
         frm03010302_01.Show
         frm03010302_01.textTM01 = ""
         frm03010302_01.textTM02 = ""
         frm03010302_01.textTM02_2 = ""
         frm03010302_01.textTM03 = ""
         frm03010302_01.textTM04 = ""
      End If
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM08.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   
   textCP05.BackColor = &H8000000F
   textCP05S.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   
   textCF15_2.BackColor = &H8000000F
   'Added by Morgan 2021/12/7 CFT預設出通用定稿
   If m_TM01 = "CFT" Then
      textPrint = ""
   Else
   'end 2021/12/7
      textPrint = "N"
   End If
   MoveFormToCenter Me
   InitialGrdList
   
   'Add By Sindy 2023/4/27
   m_strIR01 = frm03010302_02.m_strIR01
   m_strIR02 = frm03010302_02.m_strIR02
   m_strIR03 = frm03010302_02.m_strIR03
   m_strIR04 = frm03010302_02.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2023/4/27 END
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP05 = Empty
      m_CP09 = Empty
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

' 讀取商標基本檔
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim strSub As String
   Dim rsTmp As New ADODB.Recordset
   Dim rsSub As ADODB.Recordset
   
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
      textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
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
      ' 商標種類
      If IsNull(rsTmp.Fields("TM08")) = False Then
         m_TM08 = rsTmp.Fields("TM08")
         If m_TM10 < "010" Then
            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
         Else
            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 1)
         End If
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"))
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      'add by nickc 2006/05/29 加入閉卷提示
      If IsNull(rsTmp.Fields("TM29")) Then
         Me.lblClose.Caption = ""
      Else
         Me.lblClose.Caption = "已閉卷"
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得服務頁務基本檔的欄位內容
Private Sub QueryServicePractice()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("SP05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP05")
      End If
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("SP06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP06")
      End If
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("SP07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("SP08"))
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         textTM12 = rsTmp.Fields("SP11")
      End If
      'add by nickc 2006/05/29 加入閉卷提示
      If IsNull(rsTmp.Fields("sp15")) Then
         Me.lblClose.Caption = ""
      Else
         Me.lblClose.Caption = "已閉卷"
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取案件進度檔
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 來函收文日
   textCP05S = m_CP05
   
   ' 取得案件進度檔檔案中欄位
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then
         textCP05 = rsTmp.Fields("CP05")
      End If
      ' 機關文號
      If IsNull(rsTmp.Fields("CP08")) = False Then
         'Modify By Sindy 2012/5/31 Mark
         'textCP08 = rsTmp.Fields("CP08")
      End If
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
         If m_CP10 = "303" Then
            grdList.Enabled = True
         Else
'            grdList.Enabled = False
         End If
      End If
      ' 業務區
      If IsNull(rsTmp.Fields("CP12")) = False Then
         m_CP12 = rsTmp.Fields("CP12")
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"), True) 'Modified by Lydia 2016/03/25 離職人員也顯示
      End If
      ' Add By Sindy 2009/06/22 CP24
      If IsNull(rsTmp.Fields("CP24")) = False Then
         m_CP24 = rsTmp.Fields("CP24")
      End If
      '2009/06/22 End
   End If
   rsTmp.Close
   
    'Modified by Lydia 2016/03/11 CFT改成模組判斷
    'If m_TM01 = "CFT" Or m_TM01 = "CFC" Then 'Mark by Lydia 2022/09/21 全部套用; 增加CFC案
        Dim strNA69 As String
        'Modified by Lydia 2016/08/17 改抓NA69
        'Call GetNP69("", "", "", strNA69, m_TM01, m_TM02, m_TM03, m_TM04)
        'Modified by Lydia 2016/11/21 傳入智權人員代號
        'Call GetNP69("", m_TM10, "", strNA69, m_TM01, m_TM02, m_TM03, m_TM04)
        'Modified by Lydia 2017/05/12 GetNP69更名為GetNA69
        Call GetNA69("", m_TM10, m_CP13, strNA69, m_TM01, m_TM02, m_TM03, m_TM04)
        textCP14 = strNA69
        textCP14_2 = GetStaffName(textCP14)
    'Mark by Lydia 2022/09/21 全部套用; 增加CFC案
'    Else
'    'end 2016/03/11
'        '2010/10/22 ADD BY SONIA 承辦人預設最後發文承辦人
'        ' 取得案件進度檔檔案中欄位
'        strSql = "SELECT * FROM CASEPROGRESS " & _
'                 "WHERE CP01 = '" & m_TM01 & "' AND " & _
'                       "CP02 = '" & m_TM02 & "' AND " & _
'                       "CP03 = '" & m_TM03 & "' AND " & _
'                       "CP04 = '" & m_TM04 & "' AND CP27 IS NOT NULL ORDER BY CP27 DESC,CP82 DESC,CP09 DESC"
'        rsTmp.CursorLocation = adUseClient
'        rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'        If rsTmp.RecordCount > 0 Then
'           If IsNull(rsTmp.Fields("CP14")) = False Then
'              textCP14 = rsTmp.Fields("CP14")
'              textCP14_2 = GetStaffName(textCP14, True)
'           End If
'        End If
'        rsTmp.Close
'        '2010/10/22 END
'    End If
    'end 2022/09/21
   Set rsTmp = Nothing
End Sub

Public Sub QueryData()
   Dim strDay As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   
   ' 讀取基本檔
   Select Case m_TM01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "CFT":
         QueryTradeMark
      ' 系統類別為CFC的為讀取服務業務基本檔
      Case Else:
         QueryServicePractice
   End Select
     
   ' 本案期限
   InitialGrdList
   ' 取得下一程序檔案中的資料列表在 Grid List 中
   strSql = "SELECT * FROM NextProgress " & _
            "WHERE NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         ' 是否續辦欄位必須為空白
         If IsNull(rsTmp.Fields("NP06")) = False Then
            If IsEmptyText(rsTmp.Fields("NP06")) = False Then
               GoTo NextRecord
            End If
         End If
         
         grdList.Rows = grdList.Rows + 1
         grdList.row = grdList.Rows - 1
         
         ' 收文號
         If IsNull(rsTmp.Fields("NP01")) = False Then
            grdList.TextMatrix(grdList.row, 7) = rsTmp.Fields("NP01")
         End If
         ' 下一程序
         If IsNull(rsTmp.Fields("NP07")) = False Then
            grdList.TextMatrix(grdList.row, 1) = GetCaseTypeName(m_TM01, rsTmp.Fields("NP07"))
            grdList.TextMatrix(grdList.row, 8) = rsTmp.Fields("NP07")
         End If
         ' 本所期限
         If IsNull(rsTmp.Fields("NP08")) = False Then
            If IsEmptyText(rsTmp.Fields("NP08")) = False Then
               grdList.TextMatrix(grdList.row, 2) = rsTmp.Fields("NP08")
            End If
         End If
         ' 法定期限
         If IsNull(rsTmp.Fields("NP09")) = False Then
            If IsEmptyText(rsTmp.Fields("NP09")) = False Then
               grdList.TextMatrix(grdList.row, 3) = rsTmp.Fields("NP09")
            End If
         End If
         ' 機關文號
         If IsNull(rsTmp.Fields("NP13")) = False Then
            grdList.TextMatrix(grdList.row, 4) = rsTmp.Fields("NP13")
         End If
         ' 相關人
         If IsNull(rsTmp.Fields("NP14")) = False Then
            grdList.TextMatrix(grdList.row, 5) = rsTmp.Fields("NP14")
         End If
         ' 備註
         If IsNull(rsTmp.Fields("NP15")) = False Then
            grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("NP15")
         End If
         ' 序號
         If IsNull(rsTmp.Fields("NP22")) = False Then
            grdList.TextMatrix(grdList.row, 9) = rsTmp.Fields("NP22")
         End If
NextRecord:
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   
   ' 讀取案件進度檔
   QueryCaseProgress
      
   ' 以前畫面輸入的案件性質計算承辦期限
''''   strDay = Empty
   Select Case frm03010302_02.GetSelectResult
      Case "1":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1201")
        textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1201", DBDATE(m_CP05), DBDATE(textCP06))
      Case "2":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1403")
        textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1403", DBDATE(m_CP05), DBDATE(textCP06))
   End Select
''''   If IsEmptyText(strDay) = False Then
''''      ' 90.07.03 modify by louis (承辦期限以實際的工作天數來計算)
''''      'textCP48 = DBDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''      textCP48 = DBDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
''''      If IsEmptyText(textCP06) = False Then
''''         If Val(DBDATE(textCP48)) > Val(DBDATE(textCP06)) Then
''''            textCP48 = DBDATE(textCP06)
''''         End If
''''      End If
''''   End If
   
   ' 非A類收文其預設為不可算案件數
   textCP26 = "N"
      
   Set rsTmp = Nothing
End Sub

'edit b nick 2004/11/03
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim nIndex As Integer
   Dim strSql As String
   Dim strCP09 As String
   Dim strCP10 As String
   'Dim strCP12 As String
   Dim strCP27 As String
   Dim strNP07 As String
   Dim strNP08 As String
   Dim strNP09 As String
   Dim strNP14 As String
   Dim strNP15 As String
   Dim strNP22 As String
      
 '911106 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '  新增資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   ' 案件性質
   strCP10 = "1201"
   Select Case frm03010302_02.GetSelectResult
      Case "1": strCP10 = "1201"
      Case "2": strCP10 = "1403"
   End Select
   ' 業務區別 91.8.26 MODIFY BY SONIA
   'strCP12 = GetStaffDepartment(m_CP13)
   '92.6.14 MODIFY BY SONIA
   ' 發文日
   'strCP27 = DBDATE(Date)
   '2008/12/11 MODIFY BY SONIA 有期限則不上發文日
   'strCP27 = DBDATE(SystemDate())
   If IsEmptyText(textCP06) = False Then
      strCP27 = ""
   Else
      strCP27 = DBDATE(SystemDate())
   End If
   '2008/12/11 END
   '92.6.14 END
   ' 91.03.25 modify by louis (單引號)
    'Modify By Cheng 2003/04/03
    '智權人員存最近收文A類接洽記錄單的智權人員
   '92.6.14 MODIFY BY SONIA 加發文日
   'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP32,CP43,CP64) " & _
   '         "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
   '                 "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
   '                 "'','" & textCP26 & "','" & "N" & "','" & m_CP09 & "'," & CNULL(ChgSQL(textCP64)) & ") "
    'Modify By Cheng 2004/02/04
    '業務區為最近收文A類接洽記錄單智權人員的業務區
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP27,CP32,CP43,CP64) " & _
'            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                    "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
'                    "'','" & textCP26 & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "'," & CNULL(ChgSQL(textCP64)) & ") "
   '2008/12/11 modify by sonia 合併CP14,CP48,CP06,CP07
   'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP27,CP32,CP43,CP64) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                    "'','" & textCP26 & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "'," & CNULL(ChgSQL(textCP64)) & ") "
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP27,CP32,CP43,CP64,CP14,CP48,CP06,CP07) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                    "'','" & textCP26 & "'," & CNULL(DBDATE(strCP27)) & ",'" & "N" & "','" & m_CP09 & "'," & CNULL(ChgSQL(textCP64)) & "," & CNULL(textCP14) & "," & CNULL(DBDATE(textCP48)) & "," & CNULL(DBDATE(textCP06)) & "," & CNULL(DBDATE(textCP07)) & ")"
   '2008/12/11 END
    'End
   '92.6.14 END
   cnnConnection.Execute strSql
   
    'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
    Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
   
   '2008/12/11 CANCEL BY SONIA 移至INSERT INTO CaseProgress
   '' 若有輸入承辦人時
   'If IsEmptyText(textCP14) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP14 = '" & textCP14 & "' " & _
   '            "WHERE CP09 = '" & strCP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '' 若有輸入承辦期限時
   'If IsEmptyText(textCP48) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP48 = " & DBDATE(textCP48) & " " & _
   '            "WHERE CP09 = '" & strCP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '' 本所期限
   'If IsEmptyText(textCP06) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP06 = " & DBDATE(textCP06) & " " & _
   '            "WHERE CP09 = '" & strCP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '' 法定期限
   'If IsEmptyText(textCP07) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP07 = " & DBDATE(textCP07) & " " & _
   '            "WHERE CP09 = '" & strCP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '2008/12/11 END
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若案件性質為改變原處分時, 更新下一程序檔中的是否續辦欄位為Y
   'If m_CP10 = "1403" Then
   If frm03010302_02.GetSelectResult = "2" Then '2:改變原處分
      strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
               "WHERE NP01 = '" & m_CP09 & "' AND " & _
                     "NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07 = 1403 "
      cnnConnection.Execute strSql
   End If
   
   'Add By Sindy 2009/06/10 更新下一程序檔案件性質為997.收達998.提申的資料
   If frm03010302_02.GetSelectResult = "1" Then '1:審查報告
      strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
               "WHERE NP01 = '" & m_CP09 & "' AND " & _
                     "NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07 in (997,998) "
      cnnConnection.Execute strSql
   End If
   '2009/06/10 End
   
   'Add By Sindy 2009/06/22
   '以「審查報告」之法定期限加1年更新至「商申」的催審期限
   If frm03010302_02.GetSelectResult = "1" And _
      m_CP10 = "101" And (IsNull(m_CP24) Or Trim(m_CP24) = "") Then
      ' 催審期限=審查報告之法定期限+1年
      ' 本所期限=法定期限(催審期限)
      'Modified by Lydia 2025/09/01 若法限為空白代入系統日；textCP07.Text=>IIf(textCP07.Text = "", strSrvDate(1), textCP07.Text)
      'Modified by Lydia 2025/09/12 CFT審查報告、通知公告日更新商申催審期限：除已有國家檔設定NA63=>CFT審查報告更新商申催審月數，以審查報告之法定期限加6個月更新至「商申」的催審期限
      'strNP09 = DBDATE(DateAdd("m", 12, ChangeWStringToWDateString(DBDATE(IIf(textCP07.Text = "", strSrvDate(1), textCP07.Text)))))
      intI = Val(Pub_GetField("NATION", " NA01='" & m_TM10 & "' ", "NA63"))
      If intI = 0 Then intI = 6
      strNP09 = DBDATE(DateAdd("m", intI, ChangeWStringToWDateString(DBDATE(IIf(textCP07.Text = "", strSrvDate(1), DBDATE(textCP07.Text))))))
      'end 2025/09/12
      'Modified by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天 ; strNP09=> PUB_GetWorkDay1(strNP09, True)
      'Modified by Lydia 2025/09/12 +AND NP06 IS NULL
      strSql = "UPDATE NextProgress SET NP08 = " & PUB_GetWorkDay1(strNP09, True) & ", NP09 = " & strNP09 & " " & _
               "WHERE NP01 = '" & m_CP09 & "' AND NP02 = '" & m_TM01 & "' AND NP03 = '" & m_TM02 & "' AND NP04 = '" & m_TM03 & "' AND NP05 = '" & m_TM04 & "' AND NP07 ='305' AND NP06 IS NULL "
      cnnConnection.Execute strSql
   End If
   '2009/06/22 End
   
   'Added by Lydia 2025/09/12 TF基礎案號設定：基礎案狀態通知Email
   Dim strTFcase As String
   If m_TM01 = "CFT" Then
      strTFcase = PUB_GetTFbaseInfo(m_TM01, m_TM02, m_TM03, m_TM04, "", m_TM10, "2", textTM12, strCP09)
   End If
   'end 2025/09/12
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若有輸入下一程序時, 新增資料到下一程序檔
   If IsEmptyText(textCF15) = False Then
      strNP22 = GetNextProgressNo()
      strNP14 = Empty
      strNP14 = GetRelatedPerson(m_CP09)
      'Modify By Cheng 2002/09/25
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13," & _
'         "NP14,NP22,NP15) VALUES " & _
'         "('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & _
'            textCF15 & "," & DBDATE(textCP06) & "," & DBDATE(textCP07) & ",'" & m_CP13 & "','" & _
'            textCP08 & "','" & strNP14 & "'," & strNP22 & "," & CNULL(ChgSQL(textCP64)) & ")"
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13," & _
         "NP14,NP22,NP15) VALUES " & _
         "('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & _
            textCF15 & "," & DBDATE(textCP06) & "," & DBDATE(textCP07) & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & _
            textCP08 & "','" & ChgSQL(strNP14) & "'," & strNP22 & "," & CNULL(ChgSQL(textCP64)) & ")"
      cnnConnection.Execute strSql
      
      '910729 Sieg
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
      Select Case textCF15
         Case "102", "105", "702", "708", "305", "998", "997":
         Case Else:
            'add by nickc 2005/05/31
            IsAppNpData = True
            SeekNewCp09 = strCP09
            'Modify By Cheng 2003/07/09
            '改成整批列印
'            ' 列印國內案件接洽及結案記錄單
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
      End Select
   End If
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若是否為核駁為Y時, 更新商標基本檔的審定來函日為來函收文日, 目前准駁為駁
   If textPermit = "Y" Then
      strSql = "UPDATE TradeMark SET TM13 = " & DBDATE(m_CP05) & ", " & _
                                    "TM16 = '2' " & _
               "WHERE TM01 = '" & m_TM01 & "' AND " & _
                     "TM02 = '" & m_TM02 & "' AND " & _
                     "TM03 = '" & m_TM03 & "' AND " & _
                     "TM04 = '" & m_TM04 & "' "
      cnnConnection.Execute strSql
        'Add By Cheng 2004/02/03
        '更新原點選資料的CP24實際結果, CP25准駁日
        strSql = "Update CaseProgress Set CP24='2', CP25=" & DBDATE(m_CP05) & " Where CP09='" & m_CP09 & "' "
        cnnConnection.Execute strSql
        'End
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新使用者所選取的本案期限資料
   For nIndex = 1 To grdList.Rows - 1
      ' 判斷該列是否有被選取
      If grdList.TextMatrix(nIndex, 0) = "v" Then
         strNP07 = grdList.TextMatrix(nIndex, 8)
         strNP22 = grdList.TextMatrix(nIndex, 9)
         strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
                  "WHERE NP02 = '" & m_TM01 & "' AND " & _
                        "NP03 = '" & m_TM02 & "' AND " & _
                        "NP04 = '" & m_TM03 & "' AND " & _
                        "NP05 = '" & m_TM04 & "' AND " & _
                        "NP07 = " & strNP07 & " AND " & _
                        "NP22 = " & strNP22 & " "
         cnnConnection.Execute strSql
      End If
   Next nIndex
   '911106 nick transation
 
   'Add by Sindy 2023/4/27
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm03010302_01", strCP09
   End If
   '2023/4/27 END
   
   cnnConnection.CommitTrans
  
   Exit Function

CheckingErr:
   MsgBox (Err.Description)
   cnnConnection.RollbackTrans
   OnSaveData = False
End Function

'Add By Sindy 2021/12/8
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
   'Add By Sindy 2021/12/8
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   'Added by Morgan 2021/12/7 CFT 沒設定都出通用定稿(不列印)
   If m_TM01 = "CFT" Then
      NowPrint m_CP09, "04", "00", False, strUserNum, 0, , , , , , , , , , , , , , , , , True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call PUB_SendMailCache 'Added by Lydia 2025/09/12
   
   Set frm03010302_03 = Nothing
End Sub

'910718 Sieg
Private Sub grdList_Click()
 Static intLastRow As Integer
   GridClick grdList, intLastRow, 0, 1
End Sub

' 下一程序
Private Sub textCF15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textCF15_2 = Empty
   If IsEmptyText(textCF15) = False Then
      '2014/11/26 ADD BY SONIA
      If Len(Me.textCF15.Text) <> 3 Then
         Cancel = True
         MsgBox "下一程序欄位值必須為三碼!!!", vbExclamation
         textCF15_GotFocus
         Exit Sub
      End If
      '2014/11/26 END
      If m_TM10 < "010" Then
         textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 0)
      Else
         textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 1)
      End If
      If IsEmptyText(textCF15_2) = True Then
         strTit = "檢核資料"
         strMsg = "下一程序代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse textCF15
EXITSUB:
End Sub

' 本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
   Dim strDay As String
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
      
      ' 以前畫面輸入的案件性質計算承辦期限
''''edit by nickc 2007/10/11    改抓有時效性的
''''      strDay = Empty
      '2008/12/11 MODIFY BY SONIA加入textCP48 =""條件,否則輸入後又會被改掉
      If textCP48 = "" Then
         Select Case frm03010302_02.GetSelectResult
            Case "1":
   ''''            strDay = GetWorkDays(m_TM01, m_TM10, "1201")
                   textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1201", DBDATE(m_CP05), DBDATE(textCP06))
            Case "2":
   ''''            strDay = GetWorkDays(m_TM01, m_TM10, "1403")
                   textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1403", DBDATE(m_CP05), DBDATE(textCP06))
         End Select
      End If
''''      If IsEmptyText(strDay) = False Then
''''         ' 90.07.03 modify by louis (承辦期限以實際的工作天數來計算)
''''         'textCP48 = DBDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''         textCP48 = DBDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
''''         If IsEmptyText(textCP06) = False Then
''''            If Val(DBDATE(textCP48)) > Val(DBDATE(textCP06)) Then
''''               textCP48 = DBDATE(textCP06)
''''            End If
''''         End If
''''      End If
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
      End If
   End If
End Sub

'Add By Sindy 2010/11/29
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

' 承辦人期限
Private Sub textCP48_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP48) = False Then
      ' 檢查是否為民國日期
      If CheckIsDate(textCP48, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的承辦期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48_GotFocus
         Exit Sub
      End If
   End If
   'Add By Cheng 2002/05/07
   '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
   If Len(Me.textCP06.Text) > 0 And Len(Me.textCP48.Text) > 0 Then
      If Val(Me.textCP06.Text) < Val(Me.textCP48.Text) Then
         Cancel = True
         MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
         textCP48_GotFocus
         Exit Sub
      End If
   End If
   
End Sub

Private Sub textPermit_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

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

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   ' 下一程序
   If IsEmptyText(textCF15) = True Then
      If IsEmptyText(textCP06) = False Or IsEmptyText(textCP07) = False Then
         strTit = "資料檢核"
         strMsg = "無下一程序時, 本所期限及法定期限不可輸入"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
   Else
      If IsEmptyText(textCP06) = True Or IsEmptyText(textCP07) = True Then
         strTit = "資料檢核"
         strMsg = "有下一程序時, 本所期限及法定期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
      'Add By Cheng 2002/03/11
      If Me.textCP06.Text <> "" Then
         If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
            MsgBox "本所期限不可小於系統日期!!!", vbExclamation
            Me.textCP06.SetFocus
            textCP06_GotFocus
            GoTo EXITSUB
         Else
            If Val(Me.textCP06.Text) > Val(Me.textCP07.Text) Then
               MsgBox "本所期限不可大於法定期限!!!", vbExclamation
               Me.textCP06.SetFocus
               textCP06_GotFocus
               GoTo EXITSUB
            End If
         End If
      End If
   End If
   ' 承辦期限不可超過本所期限
   If IsEmptyText(textCP48) = False And IsEmptyText(textCP06) = False Then
      If Val(textCP48) > Val(textCP06) Then
         strTit = "資料檢核"
         strMsg = "承辦期限的日期不可超過本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Rows = 1
   grdList.Cols = 10
   grdList.ColWidth(0) = 300
   
   grdList.row = 0
   grdList.col = 0
   grdList.ColWidth(0) = 300
   grdList.Text = ""
   grdList.col = 1
   grdList.Text = "下一程序"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "本所期限"
   grdList.ColWidth(2) = 1000
   grdList.col = 3
   grdList.Text = "法定期限"
   grdList.ColWidth(3) = 1000
   grdList.col = 4
   grdList.Text = "機關文號"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "相關人"
   grdList.ColWidth(5) = 1200
   grdList.col = 6
   grdList.Text = "備註"
   grdList.ColWidth(6) = 1200
   grdList.col = 7
   grdList.Text = "收文號"
   grdList.ColWidth(7) = 0
   grdList.col = 8
   grdList.Text = "下一程序代號"
   grdList.ColWidth(8) = 0
   grdList.col = 9
   grdList.Text = "序號"
   grdList.ColWidth(9) = 0
End Sub

Private Sub grdList_KeyDown(KeyCode As Integer, Shift As Integer)
'910718 Sieg
Exit Sub
   ' 案件性質必須為延期的才可以選取
   If KeyCode = vbKeySpace Then
      If grdList.row > 0 Then
         grdList.col = 0
         If grdList.Text = "V" Then
            grdList.Text = Empty
         Else
            grdList.Text = "V"
         End If
      End If
   End If
EXITSUB:
End Sub

Private Sub grdList_SelChange()
'910718 Sieg
   Exit Sub
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 0
      grdList.Text = ""
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
         Next nCol
      End If
      grdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 0
      grdList.Text = "V"
      grdList.col = 1
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      grdList.col = 0
   End If
EXITSUB:
End Sub

Private Sub textCP08_GotFocus()
   InverseTextBox textCP08
End Sub

Private Sub textCF15_GotFocus()
   InverseTextBox textCF15
End Sub

Private Sub textCP06_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP14_GotFocus()
   InverseTextBox textCP14
End Sub

Private Sub textCP48_GotFocus()
   InverseTextBox textCP48
End Sub

Private Sub textPermit_GotFocus()
   InverseTextBox textPermit
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
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
If Me.textCF15.Enabled = True Then
   Cancel = False
   textCF15_Validate Cancel
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

If Me.textCP48.Enabled = True Then
   Cancel = False
   textCP48_Validate Cancel
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

'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
     Exit Function
End If
   
TxtValidate = True
End Function

