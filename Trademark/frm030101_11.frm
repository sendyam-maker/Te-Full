VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030101_11 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(延期)"
   ClientHeight    =   6060
   ClientLeft      =   5080
   ClientTop       =   1540
   ClientWidth     =   9140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   9140
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   4752
      MaxLength       =   4
      TabIndex        =   7
      Top             =   3840
      Width           =   540
   End
   Begin VB.TextBox Text9 
      Height          =   264
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   3
      Top             =   3555
      Width           =   372
   End
   Begin VB.TextBox textCP07 
      Height          =   264
      Left            =   7830
      MaxLength       =   8
      TabIndex        =   5
      Top             =   3555
      Width           =   1200
   End
   Begin VB.TextBox textCP06 
      Height          =   264
      Left            =   4380
      MaxLength       =   8
      TabIndex        =   4
      Top             =   3555
      Width           =   1200
   End
   Begin VB.TextBox textDL02 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2532
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8220
      TabIndex        =   11
      Top             =   60
      Width           =   852
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6180
      TabIndex        =   9
      Top             =   60
      Width           =   852
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7080
      TabIndex        =   10
      Top             =   60
      Width           =   1092
   End
   Begin VB.TextBox textCP27 
      Height          =   264
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2952
      Width           =   1092
   End
   Begin VB.TextBox textPrint 
      Height          =   264
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   6
      Top             =   3840
      Width           =   372
   End
   Begin VB.TextBox textCP18 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2928
      Width           =   2532
   End
   Begin VB.ComboBox textCP44 
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   3240
      Width           =   1668
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2532
   End
   Begin VB.TextBox textCP12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
   End
   Begin VB.TextBox textTM20 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   840
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1140
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   840
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2532
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1140
      Width           =   2532
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   1212
      Left            =   1200
      TabIndex        =   52
      Top             =   4128
      Width           =   7872
      _ExtentX        =   13882
      _ExtentY        =   2134
      _Version        =   393216
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
      _Band(0).Cols   =   2
   End
   Begin MSForms.TextBox textCP64 
      Height          =   555
      Left            =   1200
      TabIndex        =   8
      Top             =   5400
      Width           =   7815
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13785;979"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP44_2 
      Height          =   270
      Left            =   2940
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   3240
      Width           =   6060
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "10689;476"
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
      TabIndex        =   50
      Top             =   2310
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
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1200
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2040
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
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5700
      TabIndex        =   48
      Top             =   1710
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
   Begin MSForms.TextBox textCP14 
      Height          =   285
      Left            =   1200
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   1740
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
   Begin VB.Label lblCP113 
      AutoSize        =   -1  'True
      Caption         =   "工作時數:"
      Height          =   180
      Index           =   18
      Left            =   3630
      TabIndex        =   46
      Top             =   3882
      Width           =   765
   End
   Begin VB.Label Label12 
      Caption         =   "延期月數 :"
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   3570
      Width           =   810
   End
   Begin VB.Label Label10 
      Caption         =   "延期後法定期限 :"
      Height          =   255
      Left            =   6270
      TabIndex        =   44
      Top             =   3570
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "延期後本所期限 :"
      Height          =   255
      Left            =   2820
      TabIndex        =   43
      Top             =   3570
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "上次延期日 :"
      Height          =   252
      Left            =   120
      TabIndex        =   41
      Top             =   2640
      Width           =   1092
   End
   Begin VB.Label Label28 
      Caption         =   "進度備註 :"
      Height          =   252
      Left            =   120
      TabIndex        =   40
      Top             =   5400
      Width           =   972
   End
   Begin VB.Label Label25 
      Caption         =   "發文日 :"
      Height          =   252
      Left            =   120
      TabIndex        =   39
      Top             =   2940
      Width           =   852
   End
   Begin VB.Label Label4 
      Caption         =   "代理人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   38
      Top             =   3240
      Width           =   972
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   252
      Left            =   120
      TabIndex        =   37
      Top             =   3840
      Width           =   972
   End
   Begin VB.Label Label23 
      Caption         =   "(N:不印)"
      Height          =   252
      Left            =   1680
      TabIndex        =   36
      Top             =   3840
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "點數 :"
      Height          =   252
      Index           =   10
      Left            =   4740
      TabIndex        =   35
      Top             =   2928
      Width           =   732
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   34
      Top             =   2040
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "商標種類 :"
      Height          =   252
      Index           =   4
      Left            =   4740
      TabIndex        =   33
      Top             =   2040
      Width           =   852
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   32
      Top             =   1740
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "業務區別 :"
      Height          =   252
      Index           =   2
      Left            =   4740
      TabIndex        =   31
      Top             =   540
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "發證日 :"
      Height          =   252
      Index           =   3
      Left            =   4740
      TabIndex        =   30
      Top             =   840
      Width           =   732
   End
   Begin VB.Label Label3 
      Caption         =   "申請案號 :"
      Height          =   252
      Left            =   4740
      TabIndex        =   29
      Top             =   1140
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   27
      Top             =   540
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   26
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   252
      Index           =   9
      Left            =   4740
      TabIndex        =   25
      Top             =   1440
      Width           =   920
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4740
      TabIndex        =   24
      Top             =   1740
      Width           =   920
   End
   Begin VB.Label Label2 
      Caption         =   "審定號數 :"
      Height          =   252
      Left            =   120
      TabIndex        =   23
      Top             =   1140
      Width           =   732
   End
   Begin VB.Label Label5 
      Caption         =   "案件名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   22
      Top             =   2340
      Width           =   972
   End
   Begin VB.Label Label9 
      Caption         =   "欲延期期限:"
      Height          =   252
      Left            =   120
      TabIndex        =   21
      Top             =   4140
      Width           =   1092
   End
End
Attribute VB_Name = "frm030101_11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/17 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Lydia 2021/08/11 改成Form2.0 ; textCP13、textCP14、textTM23、cmbTM05、textCP44_2、textCP64、grdList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

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
' 承辦人代號
Dim m_CP14 As String
' 本所期限
Dim m_CP06 As String
' 法定期限
Dim m_CP07 As String
' 智權人員代號
Dim m_CP13 As String
' 業務區
Dim m_CP12 As String
' 代理人
Dim m_CP44 As String
'彼所案號
Dim m_CP45 As String

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

' 宣告代理人內容結構
Private Type AGENTITEM
   aiCode As String
   aiName As String
End Type
Dim m_AgentList() As AGENTITEM
Dim m_AgentCount As Integer
'
Dim m_CurrSel As Integer
'Add By Cheng 2002/08/19
Dim m_strDL05 As String
'add by nickc 2008/02/22
Dim m_CP44New As String
Dim m_CP116 As String
Dim m_TM44 As String
Dim m_TM119 As String
Dim m_TM120 As String
Dim m_990CP09 As String 'Add By Sindy 2016/12/20


Private Sub cmdCancel_Click()
   frm030101_01.Show
   Unload Me
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm030101_01
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
   
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 更新欄位輸入的內容
      OnUpdateField
      ' 存檔
      'edit by nick 2004/11/03
      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub

      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      'Add By Sindy 2012/4/5 CFT,FCT所有案件性質發文時,檢查代表圖是否存在
      'Modify by Amy 2018/07/31 ChkIsExistImg不使用
      'Call ChkIsExistImg(m_TM01, m_TM02, m_TM03, m_TM04)
      If ChkImgByteFile(m_TM01, m_TM02, m_TM03, m_TM04) = False Then MsgBox "本案尚未放代表圖至系統！"
      
      '*********** 90.11.23   nick  清畫面
      'frm030101_01.radio(0).Value = True
      'frm030101_01.textCP09.Enabled = True
      'frm030101_01.textCP09.Text = ""
      'frm030101_01.textTM01.Enabled = False
      'frm030101_01.textTM01.Text = ""
      'frm030101_01.textTM02.Enabled = False
      'frm030101_01.textTM02.Text = ""
      'frm030101_01.textTM02_2.Enabled = False
      'frm030101_01.textTM02_2.Text = ""
      'frm030101_01.textTM03.Enabled = False
      'frm030101_01.textTM03.Text = "'"
      'frm030101_01.textTM04.Enabled = False
      'frm030101_01.textTM04.Text = ""
      'frm030101_01.grdList.Clear
      'frm030101_01.grdList.Rows = 2
      'frm030101_01.RefreshData
      '***********************************
      'Add By Sindy 2024/8/19
      If frm030101_01.bolIsEMPFlow = True Then
         frm090202_4.QueryData
      End If
      '2024/8/19 End
      'Add By Cheng 2002/04/30
      '若有未發文資料顯示警告
      If m_CP10 = "303" Then
         If PUB_GetCPunIssueDatas("" & Me.textTMKey.Text) = False Then
            'Add By Sindy 2024/8/19
            If frm030101_01.bolIsEMPFlow = True Then
               Unload frm030101_01
               frm090202_4.Show
               Unload Me
               Exit Sub
            End If
            '2024/8/19 End
         End If
      End If
      frm030101_01.Show
      ' 90.12.07 modify by louis
'      frm030101_01.Clear
      'Add By Cheng 2002/01/10
      frm030101_01.Clear1
      Unload Me
   End If
End Sub

Private Sub Form_Activate()
'add by nickc 2005/08/23
'Modify By Sindy 2012/10/1 下列程式無意義Mark
'If (pub_ModifyCaseNum = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 And pub_ModifyCaseNum <> "") Then
'   pub_ModifyCaseNum = ""
'   QueryData
'End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM08.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM20.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP12.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP18.BackColor = &H8000000F
   textCP44_2.BackColor = &H8000000F
   
   textDL02.BackColor = &H8000000F
   
   MoveFormToCenter Me
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
   End Select
End Sub

Private Sub ClearAgentList()
   If m_AgentCount > 0 Then
      Erase m_AgentList
   End If
   m_AgentCount = 0
End Sub

Private Sub AddAgent(ByVal strAgentCode As String, ByVal strAgentName As String)
   Dim nIndex As Integer
   Dim bFind As Boolean
   bFind = False
   For nIndex = 0 To m_AgentCount - 1
      If m_AgentList(nIndex).aiCode = strAgentCode Then
         bFind = True
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      ReDim Preserve m_AgentList(m_AgentCount + 1)
      m_AgentList(m_AgentCount).aiCode = strAgentCode
      m_AgentList(m_AgentCount).aiName = strAgentName
      m_AgentCount = m_AgentCount + 1
   End If
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
   Dim strSubSQL As String
   Dim rsSubTmp As New ADODB.Recordset
   
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
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      ' 發證日
      If IsNull(rsTmp.Fields("TM20")) = False Then
         textTM20 = DBDATE(rsTmp.Fields("TM20"))
      End If
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM05")
      End If
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM06")
      End If
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM07")
      End If
      ' 顯示案件名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 商標種類
      If IsNull(rsTmp.Fields("TM08")) = False Then
         textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
End Sub

' 取得服務業務基本檔的欄位內容
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
        'add by nickc 2008/02/22
        m_TM44 = CheckStr(rsTmp.Fields("SP26"))
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
      ' 顯示案件名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         textTM12 = rsTmp.Fields("SP11")
      End If
      ' 發證日
      If IsNull(rsTmp.Fields("SP12")) = False Then
         textTM20 = rsTmp.Fields("SP12")
      End If
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
   Dim strCP27 As String
   Dim strCP44 As String
   Dim strCP45 As String
   Dim nIndex As Integer
   Dim bFind As Boolean
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'add by nickc 2008/02/22
      m_CP116 = CheckStr(rsTmp.Fields("CP116"))
      
      m_CP10 = Empty: m_CP14 = Empty
      ' 本所期限
      If IsNull(rsTmp.Fields("CP06")) = False Then
         m_CP06 = rsTmp.Fields("CP06")
      End If
      '2011/5/19 ADD BY SONIA
      If m_CP10 <> "303" Then
         SetCPFieldOldData "CP06", textCP06, 1
      End If
      '2011/5/19 END
      ' 法定期限
      If IsNull(rsTmp.Fields("CP07")) = False Then
         m_CP07 = rsTmp.Fields("CP07")
      End If
      '2011/5/19 ADD BY SONIA
      If m_CP10 <> "303" Then
         SetCPFieldOldData "CP07", textCP07, 1
      End If
      '2011/5/19 END
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      
      'Add By Sindy 2010/12/27 判斷有相關總收文號才做
      ' 相關總收文號
      If IsNull(rsTmp.Fields("CP43")) = False Then
         '案件性質
         textCP10 = textCP10 & PUB_GetRelateCasePropertyName(m_CP09, "1")
      End If
      '2010/12/27 End
      ' 業務區別
      If IsNull(rsTmp.Fields("CP12")) = False Then
         m_CP12 = rsTmp.Fields("CP12")
         textCP12 = GetDepartmentName(rsTmp.Fields("CP12"))
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      End If
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         '92.10.6 ADD BY SONIA
         m_CP14 = rsTmp.Fields("CP14")
         '92'10'6 END
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      ' 發文日(預設為系統日)
      'edit by nickc 2006/03/17
      'textCP27 = DBDATE(Date)
      textCP27 = strSrvDate(1)
      strCP27 = Empty
      If IsNull(rsTmp.Fields("CP27")) = False Then
         strCP27 = rsTmp.Fields("CP27")
      End If
      CaculateNP08NP09
      SetCPFieldOldData "CP27", strCP27, 1
      ' 代理人
      textCP44 = Empty
      If IsNull(rsTmp.Fields("CP44")) = False Then
         textCP44 = rsTmp.Fields("CP44")
      End If
      SetCPFieldOldData "CP44", textCP44, 0
      ' 點數
      textCP18 = Empty
      If IsNull(rsTmp.Fields("CP18")) = False Then
         textCP18 = rsTmp.Fields("CP18")
      End If
      ' 代理人
      m_CP44 = Empty
      If IsNull(rsTmp.Fields("CP44")) = False Then
         m_CP44 = rsTmp.Fields("CP44")
      End If
      ' 彼所案號
      strCP45 = Empty
      m_CP45 = Empty
      If IsNull(rsTmp.Fields("CP45")) = False Then
         strCP45 = rsTmp.Fields("CP45")
         m_CP45 = rsTmp.Fields("CP45")
      End If
      SetCPFieldOldData "CP45", strCP45, 0
      ' 進度備註
      textCP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then
         textCP64 = rsTmp.Fields("CP64")
      End If
      SetCPFieldOldData "CP64", textCP64, 0
      'Added by Lydia 2021/06/04 工作時數
       txtCP113 = "" & rsTmp.Fields("CP113")
       SetCPFieldOldData "CP113", txtCP113, 1
      'end 2021/06/04
      
      ' 代理人
      ClearAgentList
      'Add By Sindy 2013/5/23 若是原先有，也要加入
      If textCP44.Text <> "" Then
'         If InStr(textCP44, "-") > 0 Then
'            If ClsPDGetContact(textCP44, strCP44) Then
'               AddAgent textCP44, strCP44
'            End If
'         Else
            strCP44 = GetFAgentName(textCP44)
            AddAgent textCP44, strCP44
'         End If
      End If
      '2013/5/23 End
      '2009/2/3 modify by sonia B類收文之文件簽證711及申請英文證明304不要列入
      '2010/9/7 Modify by Sindy 文件簽證711及申請英文證明304不要列入
      strSubSQL = "SELECT CP44, MAX(CP27) AS CP27 FROM CASEPROGRESS " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND " & _
                        "CP04 = '" & m_TM04 & "' AND " & _
                        "CP09 <> '" & m_CP09 & "' And CP09<'C' And CP44 Is Not Null " & _
                        "AND CP10 NOT IN ('711','304') " & _
                  "GROUP BY CP44 " & _
                  "ORDER BY CP27 DESC "
      rsSubTmp.CursorLocation = adUseClient
      rsSubTmp.Open strSubSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If rsSubTmp.RecordCount > 0 Then
         rsSubTmp.MoveFirst
         ' 依序將代理人加入到系統串列中
         Do While rsSubTmp.EOF = False
            If IsNull(rsSubTmp.Fields("CP44")) = False Then
               strCP44 = GetFAgentName(rsSubTmp.Fields("CP44"))
               AddAgent rsSubTmp.Fields("CP44"), GetFAgentName(rsSubTmp.Fields("CP44"))
            End If
            rsSubTmp.MoveNext
         Loop
      End If
      rsSubTmp.Close
      ' 從系統串列中取得所有代理人並放入Combo Box中
      For nIndex = 0 To m_AgentCount - 1
         textCP44.AddItem m_AgentList(nIndex).aiCode
      Next nIndex
      ' 設定顯示為第一筆
      If textCP44.ListCount > 0 Then
         textCP44.ListIndex = 0
         textCP44_Validate False
      End If
   End If
   rsTmp.Close
   
   ' 未收文期限
   InitialGrdList
   ' 取得下一程序檔案中的資料列表在 Grid List 中
   'Modify by Morgan 2009/12/29 下一程序要排除程序管制的案件性質
   '2010/3/23 MODIFY BY SONIA 剔除下一程序非智權人員掌控之案件性質改以strNpSqlOfNoSalesDuty控制
   strSql = "SELECT NP01,NP07,NP08,NP09,NP10,NP11,NP12,NP13,NP14,NP15,NP22 FROM NextProgress " & _
            "WHERE NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "' AND NP06 IS NULL " & strNpSqlOfNoSalesDuty
      
   'add by sonia 2017/6/16 CFT案加可選下一程序緩衝期限CFT-017235
   strSql = strSql & " UNION SELECT NP01,NP07,NP08,NP09,NP10,NP11,NP12,NP13,NP14,NP15,NP22 FROM NextProgress " & _
            "WHERE NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "' AND NP06 IS NULL AND NP07='312'"
   'end 2017/6/16
   
   'Add by Morgan 2009/12/29 延期+AB類未發文未取消收文的程序
   strSql = strSql & " UNION SELECT CP09,CP10,CP06,CP07,CP13,CP57,CP58,CP08,CP40,CP64,0" & _
      " FROM CASEPROGRESS WHERE CP01 = '" & m_TM01 & "' AND CP02 = '" & m_TM02 & "'" & _
      " AND CP03 = '" & m_TM03 & "' AND CP04 = '" & m_TM04 & "'" & _
      " AND CP09<'C' and cp10<>'303' and cp07>0 AND CP27 IS NULL AND CP57 IS NULL"
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         'Remove by Morgan 2009/12/29 改語法加條件控制
         '' 是否續辦欄位必須為空白
         'If IsNull(rsTmp.Fields("NP06")) = False Then
         '   If IsEmptyText(rsTmp.Fields("NP06")) = False Then
         '      GoTo NextRecord
         '   End If
         'End If
         
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
               grdList.TextMatrix(grdList.row, 2) = DBDATE(rsTmp.Fields("NP08"))
            End If
         End If
         ' 法定期限
         If IsNull(rsTmp.Fields("NP09")) = False Then
            If IsEmptyText(rsTmp.Fields("NP09")) = False Then
               grdList.TextMatrix(grdList.row, 3) = DBDATE(rsTmp.Fields("NP09"))
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
      'Added by Lydia 2023/10/17
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/17
   End If
   rsTmp.Close
   
   ' 上次延期日(取最後一筆)
   strSql = "SELECT * FROM DateLimit " & _
            "WHERE DL01 = '" & m_CP09 & "' AND " & _
                  "DL02 IN (SELECT MAX(DL02) FROM DateLimit " & _
                           "WHERE DL01 = '" & m_CP09 & "') "
            
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("DL02")) = False Then
         If rsTmp.Fields("DL02") <> "0" Then
            textDL02 = rsTmp.Fields("DL02")
         End If
      End If
   End If
   rsTmp.Close
      
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
   
End Sub

' 讀取資料庫
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
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
'   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   textTMKey.Text = m_TM01 & "-" & m_TM02 & "-" & IIf(Len("" & m_TM03) <= 0, "0", m_TM03) & "-" & IIf(Len("" & m_TM04) <= 0, "00", m_TM04)

   ' 收文號
   textCP09 = m_CP09
   
   ' 取得國家代碼
   m_TM10 = GetNationNo(m_TM01, m_TM02, m_TM03, m_TM04)
   
   ' 取得案件進度檔的欄位
   QueryCaseProgress
   
   Select Case m_TM01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "CFT":
         QueryTradeMark
      Case Else:
         QueryServicePractice
   End Select
   
   ' 大約?可接獲回音(欄位)
   'textCF09 = Empty
   'strSQL = "SELECT * FROM CaseFee " & _
   '         "WHERE CF01 = '" & m_TM01 & "' AND " & _
   '               "CF02 = '" & m_TM10 & "' AND " & _
   '               "CF03 = '" & m_CP10 & "' "
   'rsTmp.CursorLocation = adUseClient
   'rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
   'If rsTmp.RecordCount > 0 Then
   '   If IsNull(rsTmp.Fields("CF09")) = False Then
   '      textCF09 = rsTmp.Fields("CF09")
   '   End If
   'End If
   'rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Add By Cheng 2002/07/19
   Set frm030101_11 = Nothing
End Sub

'Add by Morgan 2009/12/29
Private Sub Text9_GotFocus()
   TextInverse Text9
   CloseIme
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Val(Text9) > 0 Then
      strExc(1) = m_TM01
      strExc(2) = m_TM10
      strExc(3) = CompDate("1", Val(Text9), m_CP07)
      GetCtrlDT strExc
      textCP07 = TransDate(strExc(3), 2) '延期後法定期限
      '延期後本所期限
      'Modify By Sindy 2011/5/20 韓國延期發文時，請設定本所期限為法定期限減7天
      If m_TM10 = "012" Then '韓國
        textCP06 = TransDate(PUB_GetWorkDay1(CompDate("2", -7, textCP07), True), 2)
      '2011/5/20 End
      Else
        textCP06 = TransDate(PUB_GetWorkDay1(strExc(0), True), 2)
      End If
   End If
End Sub
'end 2009/12/29

' 延期後本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP06) = False Then
      If CheckIsDate(textCP06, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的延期後本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         TextInverse textCP06
      'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textCP06.Text = PUB_GetWorkDay1(textCP06, True)
      'end 2020/07/09
      End If
   End If
End Sub

' 延期後法定期限
Private Sub textCP07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP07) = False Then
      If CheckIsDate(textCP07, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的延期後法定期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         TextInverse textCP07
      End If
   End If
End Sub

' 發文日
Private Sub textCP27_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP27) = False Then
      ' 發文日日期不正確
      If CheckIsDate(textCP27, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的發文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      
      ' 發文日日期不可超過系統日
      'edit by nickc 2006/03/17
      'If Val(DBDATE(textCP27)) > Val(DBDATE(Date)) Then
      If Val(DBDATE(textCP27)) > Val(strSrvDate(1)) Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "發文日不可超過系統日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
        'Modify/Add By Cheng 2002/12/17
'        CaculateNP08NP09
'        If m_TM10 <> 美國國家代號 And m_CP10 <> "303" Then
'            CaculateNP08NP09
'        End If
   End If
EXITSUB:
End Sub

' 計算本所期限及法定期限
Private Sub CaculateNP08NP09()
   If IsEmptyText(textCP27) = False Then
      strExc(0) = TransDate(textCP27.Text, 2)
      'edit by nickc 2007/02/06 不用 dll 了
      'If objLawDll.GetCaseFeeDelay(m_TM01, m_TM10, m_CP10, strExc) Then
      If ClsLawGetCaseFeeDelay(m_TM01, m_TM10, m_CP10, strExc) Then
        'Modify By Cheng 2002/11/29
        '以西元日期顯示
'         textCP07 = TransDate(strExc(1), 1)
'         textCP06 = TransDate(strExc(2), 1)
         textCP07 = TransDate(strExc(1), 2)
         textCP06 = TransDate(strExc(2), 2)
         textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 2) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      End If
   End If
End Sub

Private Sub textCP44_Click()
   textCP44_2 = m_AgentList(textCP44.ListIndex).aiName
End Sub

' 當使用者按向下鍵時, 將ComboBox顯示成下拉式的樣子
Private Sub textCP44_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then
      SendMessage textCP44.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
   End If
End Sub

Private Sub textCP44_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 代理人
Private Sub textCP44_Validate(Cancel As Boolean)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strTempName As String   '2010/11/24 add by sonia
   
   Cancel = False
   'Add By Cheng 2002/03/08
   If m_TM10 <> 台灣國家代號 Then
      If Len(Me.textCP44.Text) <= 0 Then
         MsgBox "當申請國家非台灣時, 代理人欄不可為空白!!!", vbExclamation
         Cancel = True
         Exit Sub
      End If
   End If
   
   If textCP44.ListIndex >= 0 Then
      textCP44 = m_AgentList(textCP44.ListIndex).aiCode
   End If
   textCP44_2 = Empty
   If IsEmptyText(textCP44) = False Then
      'edit by 2004/07/22 nick  檢查該申請人或代理人狀態，若為不再使用則停在原地
      '2010/11/24 modify by sonia 取消basQuery的GetFAgentNameAndState
      'Dim oState As Boolean
      'oState = True
      ''textCP44_2 = GetFAgentName(textCP44)
      'textCP44_2 = GetFAgentNameAndState(textCP44, oState)
      'If oState = False Then
      '      Cancel = True
      '      Exit Sub
      'End If
      If PUB_GetAgentNameAndState(m_TM01, textCP44.Text, strTempName) Then
         textCP44_2 = strTempName
      Else
         textCP44_2.Text = ""
         If strTempName <> "" Then
            Cancel = True
            Exit Sub
         End If
      End If
      '2010/11/24 end
      If IsEmptyText(textCP44_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "代理人不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP44_GotFocus
      Else
         ' 依所選擇的代理人找出案件進度檔中其收文日最大的一筆其彼所案號更新到畫面上的彼所案號欄位
         strSql = "SELECT CP45 FROM CaseProgress " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND " & _
                        "CP04 = '" & m_TM04 & "' AND " & _
                        "CP44 = '" & textCP44 & "' AND " & _
                        "CP05 IN (SELECT MAX(CP05) FROM CASEPROGRESS " & _
                                 "WHERE CP01 = '" & m_TM01 & "' AND " & _
                                       "CP02 = '" & m_TM02 & "' AND " & _
                                       "CP03 = '" & m_TM03 & "' AND " & _
                                       "CP04 = '" & m_TM04 & "' AND " & _
                                       "CP44 = '" & textCP44 & "')"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("CP45")) = False Then
               textTM45 = rsTmp.Fields("CP45")
            End If
         End If
         rsTmp.Close
      End If
   End If
   Set rsTmp = Nothing
End Sub

' 進度備註
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP64, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "進度備註內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
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
   Dim nIndex As Integer
    'Add By Cheng 2002/12/17
    Dim ii As Integer
       
   ' 更新案件進度檔的欄位
   '若延期發文是按確定按鈕進來者(點選的案件進度資料為延期)
   If m_CP10 = "303" Then
      ' 發文日
        'Modify By Cheng 2002/12/17
'      SetCPFieldNewData "CP27", DBDATE(textDL02)
      SetCPFieldNewData "CP27", DBDATE(Me.textCP27.Text)
      ' 代理人
      If IsEmptyText(textCP44) = False Then
         SetCPFieldNewData "CP44", textCP44 & String(9 - Len(textCP44), "0")
        'add by nickc 2008/02/22
        m_CP44New = textCP44 & String(9 - Len(textCP44), "0")
      Else
         SetCPFieldNewData "CP44", textCP44
        'add by nickc 2008/02/22
        m_CP44New = textCP44
      End If
      ' 彼所案號
      SetCPFieldNewData "CP45", textTM45
      ' 進度備註
      SetCPFieldNewData "CP64", textCP64
        'Add By Cheng 2002/12/17
        If Me.grdList.Rows > 1 Then
            For ii = 1 To Me.grdList.Rows - 1
                ' 相關總收文號
                If Me.grdList.TextMatrix(ii, 0) <> "" Then SetCPFieldNewData "CP43", Me.grdList.TextMatrix(ii, 7): Exit For
            Next ii
        End If
   '若延期發文是按延期按鈕進來者(點選的案件進度資料非延期)
   Else
      ' 本所期限
      If IsEmptyText(textCP06) = False Then
         SetCPFieldNewData "CP06", DBDATE(textCP06)
      Else
         SetCPFieldNewData "CP06", textCP06
      End If
      ' 法定期限
      If IsEmptyText(textCP07) = False Then
         SetCPFieldNewData "CP07", DBDATE(textCP07)
      Else
         SetCPFieldNewData "CP07", textCP07
      End If
   End If
   
   'Added by Lydia 2021/06/04 工作時數
   SetCPFieldNewData "CP113", txtCP113
   
End Sub

'edit b nick 2004/11/03
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim strTmp As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   Dim strCP05 As String
   Dim strCP09 As String
   Dim strCP10 As String
   'Dim strCP12 As String
   Dim strCP27 As String
   Dim strCP44 As String
   Dim strNP01 As String
   Dim strNP08 As String
   Dim strNP07 As String
   Dim strNP22 As String
   Dim strDL01 As String
   Dim strDL03 As String
   Dim strDL04 As String
   Dim NP08602 As String  '2005/12/5 ADD BY SONIA
   Dim NP09602 As String  '2005/12/5 ADD BY SONIA
   Dim str303CP10 As String, str303CP09 As String 'Add By Sindy 2012/3/23
   
'911106 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   ' 新增資料到延期記錄檔
   strDL01 = Empty
   strDL03 = Empty
   strDL04 = Empty
   If m_CP10 = "303" Then
      ' 案件性質為延期時, 總收文號, 本所期限及法定期限為未收文期限所選取的收文資料
      For nIndex = 1 To grdList.Rows - 1
         ' 判斷該列是否有被選取
         If grdList.TextMatrix(nIndex, 0) = "V" Then
            strDL01 = grdList.TextMatrix(nIndex, 7)
            strDL03 = DBDATE(grdList.TextMatrix(grdList.row, 2))
            strDL04 = DBDATE(grdList.TextMatrix(grdList.row, 3))
            'Add By Cheng 2002/08/19
            strNP22 = grdList.TextMatrix(grdList.row, 9)
            
            ' 先刪除舊的資料
            strSql = "DELETE FROM DATELIMIT " & _
                     "WHERE DL01 = '" & strDL01 & "' AND " & _
                           "DL02 = " & DBDATE(textCP27) & " "
            cnnConnection.Execute strSql
            ' 新增一筆
            'Modify By Cheng 2002/08/19
'            'Modify By Cheng 2002/06/20
'            strSQL = "INSERT INTO DateLimit (DL01, DL02, DL03, DL04) " & _
'                     "VALUES ('" & strDL01 & "'," & _
'                              DBDATE(textCP27) & "," & _
'                              strDL03 & "," & strDL04 & ")"
'            strSQL = "INSERT INTO DateLimit (DL01, DL02, DL03, DL04, DL05) " & _
'                     "VALUES ('" & strDL01 & "'," & _
'                              DBDATE(textCP27) & "," & _
'                              strDL03 & "," & strDL04 & ",'" & IIf(m_CP10 = "303", "2", "1") & "')"
            m_strDL05 = IIf(m_CP10 = "303", "2", "1")
            strSql = "INSERT INTO DateLimit (DL01, DL02, DL03, DL04, DL05, DL06) " & _
                     "VALUES ('" & strDL01 & "'," & _
                              DBDATE(textCP27) & "," & _
                              strDL03 & "," & strDL04 & ",'" & m_strDL05 & "','" & IIf(m_strDL05 = "1", "", strNP22) & "' )"
            cnnConnection.Execute strSql
         End If
      Next nIndex
   Else
      ' 案件性質不為延期時, 總收文號, 本所期限及法定期限為該案本身
      strDL01 = m_CP09
      strDL03 = m_CP06
      strDL04 = m_CP07
      'Add By Cheng 2002/08/19
      strNP22 = ""
      'Modify By Cheng 2002/08/19
'      'Modify By Cheng 2002/06/20
'      strSQL = "INSERT INTO DateLimit (DL01, DL02, DL03, DL04) " & _
'               "VALUES ('" & strDL01 & "'," & _
'                        DBDATE(textCP27) & "," & _
'                        strDL03 & "," & strDL04 & ")"
'      strSQL = "INSERT INTO DateLimit (DL01, DL02, DL03, DL04, DL05) " & _
'               "VALUES ('" & strDL01 & "'," & _
'                        DBDATE(textCP27) & "," & _
'                        strDL03 & "," & strDL04 & ",'" & IIf(m_CP10 = "303", "2", "1") & "')"
      m_strDL05 = IIf(m_CP10 = "303", "2", "1")
      strSql = "INSERT INTO DateLimit (DL01, DL02, DL03, DL04, DL05, DL06) " & _
               "VALUES ('" & strDL01 & "'," & _
                        DBDATE(textCP27) & "," & _
                        strDL03 & "," & strDL04 & ",'" & m_strDL05 & "','" & IIf(m_strDL05 = "1", "", strNP22) & "' )"
      cnnConnection.Execute strSql
   End If
      
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
               strTmp = m_CPList(nIndex).fiName & " = " & "NULL"
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
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若案件性質非延期時時新增一筆資料到案件進度檔中
   If m_CP10 <> "303" Then
      ' 收文號
      strCP09 = Empty
      strCP09 = AutoNo("B", 6)
      ' 收文日
      'edit by nickc 2006/03/17
      'strCP05 = DBDATE(Date)
      strCP05 = strSrvDate(1)
      ' 案件性質
      strCP10 = "303"
      ' 業務區別 MODIFY BY SONIA 91.8.24
      'strCP12 = GetStaffDepartment(m_CP13)
      ' 發文日
      strCP27 = DBDATE(textCP27)
      
      strCP44 = Empty
      If IsEmptyText(textCP44) = False Then
         strCP44 = textCP44 & String(9 - Len(textCP44), "0")
      End If
      
      'Modified by Lydia 2018/03/31 CFT資料處理＼發文由點「延期」發文所產生之「延期」進度，請設定承辦人為其欲延期案件性質之承辦人。
      'strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP27,CP43,CP44,CP45) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & "," & _
                       m_CP06 & "," & m_CP07 & ",'" & strCP09 & "','" & strCP10 & "','" & m_CP12 & "','" & m_CP13 & "','" & strUserNum & "'," & _
                       strCP27 & ",'" & m_CP09 & "','" & strCP44 & "','" & textTM45 & "') "
      strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP27,CP43,CP44,CP45) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & "," & _
                       m_CP06 & "," & m_CP07 & ",'" & strCP09 & "','" & strCP10 & "','" & m_CP12 & "','" & m_CP13 & "','" & IIf(m_CP14 <> "", m_CP14, strUserNum) & "'," & _
                       strCP27 & ",'" & m_CP09 & "','" & strCP44 & "','" & textTM45 & "') "
      cnnConnection.Execute strSql
      ' 更新原案件進度檔資料中的本所期限及法定期限
      strSql = "UPDATE CASEPROGRESS SET CP06 = " & DBDATE(textCP06) & ", " & _
                                       "CP07 = " & DBDATE(textCP07) & " " & _
               "WHERE CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新使用者所選取的未收文期限資料
   'edit by nick 2004/09/27 加入將本案期限選取的第一個收文號存入相關總收文號
   Dim TheFirstV As String
   TheFirstV = ""
   If m_CP10 = "303" Then
      For nIndex = 1 To grdList.Rows - 1
         grdList.row = nIndex
         ' 判斷該列是否有被選取
         'Modify By Sindy 2009/11/03
         'If grdList.Text = "V" Then
         If grdList.TextMatrix(nIndex, 0) = "V" Then
         '2009/11/03 End
            strNP01 = grdList.TextMatrix(grdList.row, 7)
            'edit by nick 2004/09/27 加入將本案期限選取的第一個收文號存入相關總收文號
            If TheFirstV = "" Then TheFirstV = strNP01
            strNP07 = grdList.TextMatrix(grdList.row, 8)
            strNP22 = grdList.TextMatrix(grdList.row, 9)
            
            If Val(strNP22) > 0 Then
               strSql = "UPDATE NextProgress SET NP08 = " & textCP06 & "," & _
                                             "NP09 = " & textCP07 & " " & _
                     "WHERE NP01 = '" & strNP01 & "' AND " & _
                           "NP07 = " & strNP07 & " AND " & _
                           "NP22 = " & strNP22 & " "
                           
            'Add by Morgan 2009/12/29
            Else
               strSql = "UPDATE CaseProgress SET CP06 = " & textCP06 & "," & _
                                             "CP07 = " & textCP07 & " " & _
                     "WHERE CP09 = '" & strNP01 & "'"
            End If
            cnnConnection.Execute "begin user_data.user_notrigger:=1; end;" 'Add by Morgan 2010/7/13 +控制來函期限通知的 Trigger 不被觸發
            cnnConnection.Execute strSql, intI
            cnnConnection.Execute "begin user_data.user_notrigger:=0; end;" 'Add by Morgan 2010/7/13 +控制來函期限通知的 Trigger 不被觸發
            '2005/12/5 ADD BY SONIA 歐盟緩衝期限延期,異議答辯同時延至緩衝期限的後四個月
            If strNP07 = "312" Then
               NP09602 = DateAdd("m", 4, ChangeWStringToWDateString(DBDATE(textCP07)))
               '延期後本所期限為延期後法定期限 - 1個月
               'Modified by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
               'NP08602 = DateAdd("m", -1, NP09602)
               NP08602 = PUB_GetWorkDay1(DateAdd("m", -1, NP09602), True)
               strSql = "UPDATE NextProgress SET NP08 = " & NP08602 & "," & _
                                                "NP09 = " & ChangeWDateStringToWString(NP09602) & " " & _
                        "WHERE NP01 = '" & strNP01 & "' AND " & _
                              "NP07 = 602 AND NP06 IS NULL "
               cnnConnection.Execute "begin user_data.user_notrigger:=1; end;" 'Add by Morgan 2010/7/13 +控制來函期限通知的 Trigger 不被觸發
               cnnConnection.Execute strSql
               cnnConnection.Execute "begin user_data.user_notrigger:=0; end;" 'Add by Morgan 2010/7/13 +控制來函期限通知的 Trigger 不被觸發
               'add by sonia 2017/6/16 若異議答辯已收文未發文也要同時更新CFT-017235
               strSql = " update caseprogress SET CP06 = " & NP08602 & ",CP07 = " & ChangeWDateStringToWString(NP09602) & _
                        " Where CP01='" & m_TM01 & "' AND CP02='" & m_TM02 & "' AND CP03='" & m_TM03 & "' AND CP04='" & m_TM04 & "' AND CP10='602' AND CP158=0 AND CP159=0 and cp43 = '" & strNP01 & "'"
               cnnConnection.Execute strSql
               'end 2017/6/16
            End If
            '2005/12/5 END
         End If
      Next nIndex
      'edit by nick 2004/09/27 加入將本案期限選取的第一個收文號存入相關總收文號
      If TheFirstV <> "" Then
            strSql = " update caseprogress set cp43='" & TheFirstV & "' Where CP09='" & m_CP09 & "' "
            cnnConnection.Execute strSql
      End If
   End If

   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若有定義代理人收達天數時, 新增一筆收達的記錄到下一程序檔
   'Modify By Sindy 2012/3/23 收達期限的m_CP10及m_CP09要依延期的狀況帶入不同資料
   '按確定按鈕進入
   If m_CP10 = "303" Then
      str303CP10 = strNP07 '點未收文期限的那一筆案件性質
      str303CP09 = m_CP09
   '按延期按鈕進入
   Else
      str303CP10 = m_CP10
      str303CP09 = strCP09 'B類收文
   End If
   '2012/3/23 End
   
   'Added by Lydia 2016/03/23 下一程序997的np10用模組取得
   Dim strNP10 As String
   'Modified by Lydia 2017/05/12 GetNP69更名為GetNA69
   Call GetNA69("", m_TM10, m_CP13, strNP10, m_TM01, m_TM02, m_TM03, m_TM04)
      
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & m_TM01 & "' AND " & _
                  "CF02 = '" & m_TM10 & "' AND " & _
                  "CF03 = '" & str303CP10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CF23")) = False Then
         strNP07 = "997"
         strNP08 = DBDATE(textCP27)
        'Modify By Cheng 2003/09/02
'         strNP08 = DBDATE(Format(DateSerial(Val(DBYEAR(strNP08)), Val(DBMONTH(strNP08)), Val(DBDAY(strNP08)) + Val(rsTmp.Fields("CF23")))))
         strNP08 = DBDATE(DateAdd("d", Val(rsTmp.Fields("CF23")), ChangeWStringToWDateString(DBDATE(strNP08))))
         'Add By Sindy 2019/6/11 檢查期限是否正確
         strNP08 = PUB_T997998LimitDate(strNP08, m_CP07, 1)
         '2019/6/11 END
         strNP22 = GetNextProgressNo()
         '92.10.6 MODIFY BY SONIA
         'strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
         '         "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
         '                   strNP08 & "," & strNP08 & ",'" & strUserNum & "'," & strNP22 & ")"
         'Modified by Lydia 2016/03/23
         'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & str303CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            strNP08 & "," & strNP08 & ",'" & m_CP14 & "'," & strNP22 & ")"
         'Modified by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & str303CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            strNP08 & "," & strNP08 & ",'" & strNP10 & "'," & strNP22 & ")"
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & str303CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            PUB_GetWorkDay1(strNP08, True) & "," & strNP08 & ",'" & strNP10 & "'," & strNP22 & ")"
         '92.10.6 END
         cnnConnection.Execute strSql
         ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
         Select Case strNP07
            Case "102", "105", "702", "708", "305", "998", "997":
            Case Else:
               ' 列印國內案件接洽及結案記錄單
'               g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
                'Add By Cheng 2004/04/08
                '新增列印接洽結案單資料
                pub_AddressListSN = pub_AddressListSN + 1
                PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
         End Select
      End If
      'Modify By Sindy 2013/3/25 內外商的延期發文取消掛催審期限
'      'Add By Sindy 2012/9/10
'      If IsNull(rsTmp.Fields("CF05")) = False Then
'         strNP07 = "305"
'         strNP08 = GetUrgeDate(m_TM01, m_TM10, str303CP10, textCP27)
'         strNP22 = GetNextProgressNo()
'         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                  "VALUES ('" & str303CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
'                            strNP08 & "," & strNP08 & ",'" & m_CP14 & "'," & strNP22 & ")"
'         cnnConnection.Execute strSql
'      End If
'      '2012/9/10 End
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   'Added by Lydia 2024/07/09 判斷案件國家收費表內有設定提申期限(天)CF11，要加掛提申(998)期限；
   Call Pub_GetCF11to998(m_TM10, m_TM01, m_TM02, m_TM03, m_TM04, m_CP07, m_CP09, m_CP10, m_CP14, textCP27)
   
   'Add By Sindy 2010/7/8 檢查商品資料與基本檔商品類別是否一致
   Call CheckTMGoodsErr(m_TM01, m_TM02, m_TM03, m_TM04, False, True, m_CP14)
   
   'Add By Sindy 2016/12/20
   If m_990CP09 <> "" Then
      strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_990CP09 & "' and cp27 is null"
      cnnConnection.Execute strSql
   End If
   '2016/12/20 END
   
'911106 nick transation
    cnnConnection.CommitTrans
   
     'Add by nickc 2008/02/22 檢查代理人Email(需考慮可能為FF案件)
    PUB_CheckEMail m_CP44New, m_CP116
    PUB_CheckEMail m_TM44, m_TM119
    If m_TM120 <> "" Then
       PUB_CheckEMail m_TM44, m_TM120
    End If
    'end 2008/02/22
    
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 列印定稿
   If textPrint <> "N" Then
      PrintLetter
   End If
   
   Exit Function
CheckingErr:
   MsgBox (Err.Description)
   cnnConnection.RollbackTrans
   OnSaveData = False
End Function

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 10
   grdList.ColWidth(0) = 300
   grdList.row = 0
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
   ' 案件性質必須為延期的才可以選取
   If m_CP10 <> "303" Then: GoTo EXITSUB
   If KeyCode = vbKeySpace Then
      If grdList.row > 0 Then
         grdList.col = 0
         If grdList.Text = "V" Then
            grdList.Text = Empty
            'Add By Cheng 2002/12/17
            '若延期是按確定進來者, 點選未收文期限時, 清空期限
            If m_CP10 = "303" Then
                Me.textCP06.Text = ""
                Me.textCP07.Text = ""
            End If
         Else
            grdList.Text = "V"
            'Add By Cheng 2002/12/17
            '若延期是按確定進來者, 點選未收文期限時, 重新計算期限
            If m_CP10 = "303" Then
                ComputeDeadLine m_TM10, Me.grdList.TextMatrix(Me.grdList.row, 8), Me.grdList.TextMatrix(Me.grdList.row, 2), Me.grdList.TextMatrix(Me.grdList.row, 3), Me.textCP27.Text
            End If
         End If
      End If
   End If
EXITSUB:
End Sub

Private Sub grdList_Click()
   If grdList.Rows > 1 Then
      If grdList.row > 0 Then
         If grdList.TextMatrix(grdList.row, 0) = "V" Then
            grdList.TextMatrix(grdList.row, 0) = Empty
            'Add By Cheng 2002/12/17
            '若延期是按確定進來者, 點選未收文期限時, 清空期限
            If m_CP10 = "303" Then
                Me.textCP06.Text = ""
                Me.textCP07.Text = ""
            End If
         Else
            grdList.TextMatrix(grdList.row, 0) = "V"
            'Add By Cheng 2002/12/17
            '若延期是按確定進來者, 點選未收文期限時, 重新計算期限
            If m_CP10 = "303" Then
               'Modify by Morgan 2009/12/29 檢查期限並改用原法限計算
               'ComputeDeadLine m_TM10, Me.grdList.TextMatrix(Me.grdList.row, 8), Me.grdList.TextMatrix(Me.grdList.row, 2), Me.grdList.TextMatrix(Me.grdList.row, 3), Me.textCP27.Text
               strExc(2) = Replace(grdList.TextMatrix(grdList.row, 3), "/", "")
               If DBDATE(strExc(2)) <> m_CP07 Then
                  MsgBox "所點選案件性質的法定期限與延期程序不同，不可點選！"
                  grdList.TextMatrix(grdList.row, 0) = ""
               Else
                  ComputeDeadLine m_TM10, Me.grdList.TextMatrix(Me.grdList.row, 8), Me.grdList.TextMatrix(Me.grdList.row, 2), Me.grdList.TextMatrix(Me.grdList.row, 3), m_CP07
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub grdList_SelChange()
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      Dim nOldCol As Integer
      nOldCol = grdList.col
      grdList.col = 1
      If grdList.CellBackColor <> &H8000000D Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H8000000D Then grdList.CellBackColor = &H8000000D
            If grdList.CellForeColor <> &H80000005 Then grdList.CellForeColor = &H80000005
         Next nCol
      End If
      grdList.col = nOldCol
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
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

Private Function CheckDataValid() As Boolean
   Dim nIndex As Integer
   Dim bFind As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'add by nickc 2008/05/01
   If IsDebt(m_TM10, textCP09) Then
        strTit = "警告！禁止發文！"
        strMsg = "未收款且無 預定收款日 請轉告智權同仁！！"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        GoTo EXITSUB
   End If
   ' 發文日
   If IsEmptyText(textCP27) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入發文日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP27.SetFocus
      GoTo EXITSUB
   End If
   ' 代理人
   If IsEmptyText(textCP44) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入代理人"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP44.SetFocus
      GoTo EXITSUB
   End If
   ' 延期後本所期限範圍
   If Val(textCP06) > Val(textCP07) Then
      strTit = "檢核資料"
      strMsg = "延期後本所期限的日期不可超過延期後法定期限的日期"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP06.SetFocus
      GoTo EXITSUB
   End If
   'Add By Sindy 2011/5/20 延期後本所期限不可<=系統日
   If Val(DBDATE(textCP06)) <= Val(strSrvDate(1)) Then
      strTit = "檢核資料"
      strMsg = "延期後本所期限不可小於等於系統日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP06.SetFocus
      GoTo EXITSUB
   End If
   ' 當案件性質為延期時, 未收文期限至少要選取一筆
   If m_CP10 = "303" Then
      If grdList.Rows <= 1 Then
         strTit = "檢核資料"
         strMsg = "未收文期限無資料, 無法執行延期的處理"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
      
      bFind = False
      For nIndex = 1 To grdList.Rows - 1
         If grdList.TextMatrix(nIndex, 0) = "V" Then
            bFind = True
            Exit For
         End If
      Next nIndex
      If bFind = False Then
         strTit = "檢核資料"
         strMsg = "請先選取未收文期限的資料來做延期的處理"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   
    'add by nickc 2006/03/17 加入驗證
    Dim Cancel As Boolean
    Cancel = False
    textCP27_Validate Cancel
    If Cancel = True Then GoTo EXITSUB
    
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textCP06_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP07
End Sub

Private Sub textCP27_GotFocus()
   InverseTextBox textCP27
End Sub

Private Sub textCP44_GotFocus()
   InverseTextBox textCP44
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
   Dim strSql As String
   
   ' 系統類別為CFT
   If m_TM01 = "CFT" Then
      Select Case m_CP10
         ' 申請
         Case "101":
            ' 清除定稿例外欄位檔原有資料
            EndLetter "01", m_CP09, "01", strUserNum
            ' 回音
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "01" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
                     "','回音','" & "" & "')"
            cnnConnection.Execute strSql
         ' 不續辦
         Case "703":
            ' 清除定稿例外欄位檔原有資料
            EndLetter "01", m_CP09, "02", strUserNum
         ' 其它
         Case Else:
            ' 清除定稿例外欄位檔原有資料
            EndLetter "01", m_CP09, "03", strUserNum
      End Select
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   ' 系統類別為CFT
   If m_TM01 = "CFT" Then
      Select Case m_CP10
         ' 申請
         Case "101":
            ' 列印定稿
            NowPrint m_CP09, "01", "01", False, strUserNum, 0
         ' 不續辦
         Case "703":
            ' 列印定稿
            NowPrint m_CP09, "01", "02", False, strUserNum, 0
         ' 其它
         Case Else:
            ' 列印定稿
            NowPrint m_CP09, "01", "03", False, strUserNum, 0
      End Select
   End If
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
   
   TxtValidate = False
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
   
   If Me.textCP27.Enabled = True Then
      Cancel = False
      textCP27_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP44.Enabled = True Then
      Cancel = False
      textCP44_Validate Cancel
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
   
   If Me.textPrint.Enabled = True Then
      Cancel = False
      textPrint_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2016/12/20
   '檢查有設定副本收受人需提醒並新增信函副本B類收文
   m_990CP09 = ""
   If textPrint = "N" Then '不印定稿
      If PUB_ChkCC(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, m_990CP09) = False Then
         Exit Function
      End If
   End If
   '2016/12/20 END
    'Added by Lydia 2021/06/04 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
    If Pub_ChkACS112isNull(m_TM01, m_TM02, m_TM03, m_TM04, txtCP113) = True Then
        txtCP113.SetFocus
        txtCP113_GotFocus
        Exit Function
    End If
    'end 2021/06/04
    
   'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
   End If
       
   TxtValidate = True
End Function

'Add By Cheng 2002/12/17
Private Sub ComputeDeadLine(strTM10 As String, strNP07 As String, strNP08 As String, strNP09 As String, strCP27 As String)
'strTM10 申請國家, strNP07 下一程序案件性質, strNP08 本所期限, strNP09 法定期限, strCP27 發文日
Dim strDate0 As String '延期後本所期限
Dim strDate1 As String '延期後本所期限

'若為美國使用宣誓
If strTM10 = 美國國家代號 And strNP07 = "105" Then
    If strNP09 <> "" Then
        '延期後法定期限為原法定期限 + 6個月
        strDate1 = DateAdd("m", 6, ChangeWStringToWDateString(DBDATE(strNP09)))
        '延期後本所期限為延期後法定期限 - 2個月
        strDate0 = DateAdd("m", -2, strDate1)
        Me.textCP07.Text = ChangeWDateStringToWString(strDate1)
        Me.textCP06.Text = ChangeWDateStringToWString(strDate0)
    Else
        Me.textCP07.Text = ""
        Me.textCP06.Text = ""
    End If
'2005/12/13 ADD BY SONIA
'若為歐盟緩衝期限
ElseIf strTM10 = "239" And strNP07 = "312" Then
    If strNP09 <> "" Then
        '延期後法定期限為原法定期限 + 2個月
        strDate1 = DateAdd("m", 2, ChangeWStringToWDateString(DBDATE(strNP09)))
        '延期後本所期限為原本所期限 + 2個月
        strDate0 = DateAdd("m", 2, ChangeWStringToWDateString(DBDATE(strNP08)))
        Me.textCP07.Text = ChangeWDateStringToWString(strDate1)
        Me.textCP06.Text = ChangeWDateStringToWString(strDate0)
    Else
        Me.textCP07.Text = ""
        Me.textCP06.Text = ""
    End If
'2005/12/13 END
'其他
Else
    If strCP27 <> "" Then
        '延期後法定期限發文日 + 6個月
        strDate1 = DateAdd("m", 6, ChangeWStringToWDateString(DBDATE(strCP27)))
        '延期後本所期限為延期後法定期限 - 2個月
        strDate0 = DateAdd("m", -2, strDate1)
        Me.textCP07.Text = ChangeWDateStringToWString(strDate1)
        Me.textCP06.Text = ChangeWDateStringToWString(strDate0)
    Else
        Me.textCP07.Text = ""
        Me.textCP06.Text = ""
    End If
End If

'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
If textCP06.Text <> "" Then textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06, True), 2)

End Sub

'Added by Lydia 2021/06/04
Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

'Added by Lydia 2021/06/04
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
