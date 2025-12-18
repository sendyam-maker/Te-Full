VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030101_18 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(領證, 使用宣誓)"
   ClientHeight    =   4680
   ClientLeft      =   4720
   ClientTop       =   2030
   ClientWidth     =   9130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   9130
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   3870
      MaxLength       =   4
      TabIndex        =   5
      Top             =   3237
      Width           =   540
   End
   Begin VB.TextBox textUargeDate 
      Height          =   264
      Left            =   3480
      MaxLength       =   8
      TabIndex        =   1
      Top             =   2640
      Width           =   1092
   End
   Begin VB.TextBox textCF09 
      Height          =   264
      Left            =   6180
      MaxLength       =   3
      TabIndex        =   6
      Top             =   3240
      Width           =   612
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1140
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   540
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
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1140
      Width           =   2532
   End
   Begin VB.TextBox textTM20 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   840
      Width           =   2532
   End
   Begin VB.TextBox textCP12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2532
   End
   Begin VB.TextBox textCP27 
      Height          =   264
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2640
      Width           =   1092
   End
   Begin VB.TextBox textPrint 
      Height          =   264
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   4
      Top             =   3240
      Width           =   372
   End
   Begin VB.TextBox textCP18 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2652
      Width           =   2532
   End
   Begin VB.ComboBox textCP44 
      Height          =   276
      Left            =   1200
      TabIndex        =   3
      Top             =   2940
      Width           =   1764
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7080
      TabIndex        =   10
      Top             =   60
      Width           =   1092
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
   Begin MSForms.TextBox textPS 
      Height          =   525
      Left            =   1200
      TabIndex        =   7
      Top             =   3540
      Width           =   7815
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13779;926"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   525
      Left            =   1200
      TabIndex        =   8
      Top             =   4080
      Width           =   7812
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13779;926"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP44_2 
      Height          =   270
      Left            =   3000
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2940
      Width           =   6030
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "10636;476"
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
      TabIndex        =   48
      Top             =   2340
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
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5700
      TabIndex        =   46
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
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1200
      TabIndex        =   45
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
   Begin VB.Label lblCP113 
      AutoSize        =   -1  'True
      Caption         =   "工作時數:"
      Height          =   180
      Index           =   18
      Left            =   2910
      TabIndex        =   44
      Top             =   3282
      Width           =   765
   End
   Begin VB.Label Label14 
      Caption         =   "催審期限 :"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   43
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "列印備註 :"
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   3570
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "大約"
      Height          =   255
      Index           =   12
      Left            =   5520
      TabIndex        =   41
      Top             =   3245
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "可接獲回音"
      Height          =   255
      Left            =   6900
      TabIndex        =   40
      Top             =   3245
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "案件名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   39
      Top             =   2340
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "審定號數 :"
      Height          =   252
      Left            =   120
      TabIndex        =   38
      Top             =   1140
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4740
      TabIndex        =   37
      Top             =   1740
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   252
      Index           =   9
      Left            =   4740
      TabIndex        =   36
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   35
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   34
      Top             =   540
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   33
      Top             =   840
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "申請案號 :"
      Height          =   252
      Left            =   4740
      TabIndex        =   32
      Top             =   1140
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "發證日 :"
      Height          =   252
      Index           =   3
      Left            =   4740
      TabIndex        =   31
      Top             =   840
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "業務區別 :"
      Height          =   252
      Index           =   2
      Left            =   4740
      TabIndex        =   30
      Top             =   540
      Width           =   900
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   29
      Top             =   1740
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "商標種類 :"
      Height          =   252
      Index           =   4
      Left            =   4740
      TabIndex        =   28
      Top             =   2040
      Width           =   852
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   27
      Top             =   2040
      Width           =   852
   End
   Begin VB.Label Label28 
      Caption         =   "進度備註 :"
      Height          =   252
      Left            =   120
      TabIndex        =   26
      Top             =   4080
      Width           =   972
   End
   Begin VB.Label Label25 
      Caption         =   "發文日 :"
      Height          =   252
      Left            =   120
      TabIndex        =   25
      Top             =   2640
      Width           =   852
   End
   Begin VB.Label Label4 
      Caption         =   "代理人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   24
      Top             =   2940
      Width           =   972
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   252
      Left            =   120
      TabIndex        =   23
      Top             =   3240
      Width           =   972
   End
   Begin VB.Label Label23 
      Caption         =   "(N:不印)"
      Height          =   252
      Left            =   1680
      TabIndex        =   22
      Top             =   3240
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "點數 :"
      Height          =   252
      Index           =   10
      Left            =   4740
      TabIndex        =   21
      Top             =   2652
      Width           =   732
   End
End
Attribute VB_Name = "frm030101_18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/11 改成Form2.0 ; textCP13、textCP14、textTM23、cmbTM05、textCP44_2、textCP64、textPS
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
' 智權人員
Dim m_CP13 As String
' 承辦人
Dim m_CP14 As String

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

'Add By Cheng 2002/06/05
Dim m_TM21 As Double '專用期起日
'Add By Cheng 2002/06/06
Dim m_TM22 As Double '專用期止日
Dim m_TM20 As Double 'ADD BY SONAI 2015/12/4 發證日
Dim m_CP07 As Double '法定期限
'add by nickc 2008/02/22
Dim m_CP44New As String
Dim m_CP116 As String
Dim m_TM44 As String
Dim m_TM119 As String
Dim m_TM120 As String
Dim m_TM11 As String
Dim m_990CP09 As String 'Add By Sindy 2016/12/20

Private Sub cmdCancel_Click()
   frm030101_01.Show
   Unload Me
End Sub

'Private Sub cmdCaseProgress_Click()
'   frm030101_04.SetData 0, m_TM01, True
'   frm030101_04.SetData 1, m_TM02, False
'   frm030101_04.SetData 2, m_TM03, False
'   frm030101_04.SetData 3, m_TM04, False
'   frm030101_04.SetData 4, m_CP09, False
'   frm030101_04.SetParent Me
'   Me.Hide
'   frm030101_04.Show
'   frm030101_04.QueryData
'End Sub

Private Sub cmdExit_Click()
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
      frm030101_01.Show
      ' 90.12.07 modify by louis
'      frm030101_01.Clear
      'Add By Cheng 2002/01/10
      frm030101_01.Clear1
      Unload Me
   End If
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
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
'add by sonia 2015/12/14
Dim strTit As String
Dim strMsg As String
'end 2015/12/14
   
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
         textTM20 = rsTmp.Fields("TM20")
      End If
      'Add By Sindy 2009/08/19
      ' 申請日
      If IsNull(rsTmp.Fields("TM11")) = False Then
         m_TM11 = rsTmp.Fields("TM11")
      End If
      '2009/08/19 End
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
      'Add By Cheng 2002/06/05
      '專用期起日
      If IsNull(rsTmp.Fields("TM21")) = False Then
         m_TM21 = rsTmp.Fields("TM21").Value
      Else
         m_TM21 = 0
      End If
      'Add By Cheng 2002/06/06
      '專用期止日
      If IsNull(rsTmp.Fields("TM22")) = False Then
         m_TM22 = rsTmp.Fields("TM22").Value
      Else
         m_TM22 = 0
      End If
      'ADD BY SONIA 2015/12/4
      '發證日
      If IsNull(rsTmp.Fields("TM20")) = False Then
         m_TM20 = rsTmp.Fields("TM20").Value
      Else
         m_TM20 = 0
      End If
      'END 2015/12/4
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
   
   'add by sonia 2015/12/14 菲律賓使用宣誓案一定要有申請日才可檢查是否為申請日+3年期限(不必掛下一次)
   If m_CP10 = "105" And m_TM10 = "030" And Val(m_TM11) = 0 Then
      MsgBox "菲律賓使用宣誓案一定要有申請日才可檢查是否為申請日+3年期限!!!", vbExclamation
   End If
   'end if
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
      'Add By Cheng 2002/06/05
      '專用期起日
      If IsNull(rsTmp.Fields("SP20")) = False Then
         m_TM21 = rsTmp.Fields("SP20").Value
      Else
         m_TM21 = 0
      End If
      'Add By Cheng 2002/06/06
      '專用期止日
      If IsNull(rsTmp.Fields("SP21")) = False Then
         m_TM22 = rsTmp.Fields("SP21").Value
      Else
         m_TM22 = 0
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
   Dim strCP43 As String
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
      If IsNull(rsTmp.Fields("CP12")) = False Then
         textCP12 = GetDepartmentName(rsTmp.Fields("CP12"))
      End If
      ' 智權人員
      m_CP13 = "" 'Add By Sindy 2014/9/11
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      End If
      ' 承辦人員
      m_CP14 = Empty
      If IsNull(rsTmp.Fields("CP14")) = False Then
         m_CP14 = rsTmp.Fields("CP14")
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      End If
     ' 發文日(預設為系統日)
      textCP27 = strSrvDate(1)  'edit by nickc 2006/03/17 原先為 DBDATE(Date)
      strCP27 = Empty
      If IsNull(rsTmp.Fields("CP27")) = False Then
         strCP27 = rsTmp.Fields("CP27")
      End If
      SetCPFieldOldData "CP27", strCP27, 1
      ' 代理人
      textCP44 = Empty
      If IsNull(rsTmp.Fields("CP44")) = False Then
         textCP44 = rsTmp.Fields("CP44")
      End If
      SetCPFieldOldData "CP44", textCP44, 0
      ' 彼所案號
      strCP45 = Empty
      If IsNull(rsTmp.Fields("CP45")) = False Then
         strCP45 = rsTmp.Fields("CP45")
      End If
      SetCPFieldOldData "CP45", strCP45, 0
      ' 點數
      textCP18 = Empty
      If IsNull(rsTmp.Fields("CP18")) = False Then
         textCP18 = rsTmp.Fields("CP18")
      End If
      ' 進度備註
      textCP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then
         textCP64 = rsTmp.Fields("CP64")
      End If
      'Added by Lydia 2021/06/04 工作時數
       txtCP113 = "" & rsTmp.Fields("CP113")
       SetCPFieldOldData "CP113", txtCP113, 1
      'end 2021/06/04
      
      'Add By Cheng 2002/06/06
      '法定期限
      m_CP07 = 0
      If IsNull(rsTmp.Fields("CP07")) = False Then
         m_CP07 = rsTmp.Fields("CP07")
      End If
      
      SetCPFieldOldData "CP64", textCP64, 0
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
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
End Sub

' 讀取資料庫
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strDay As String
   
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
   
   'Add By Sindy 2009/08/19
   ' 計算催審期限
   strDay = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
   If IsEmptyText(strDay) = False Then
      textUargeDate = strDay
   End If
   Me.textCP27.Tag = Me.textCP27.Text 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
   
   ' 大約?可接獲回音(欄位)
   textCF09 = Empty
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & m_TM01 & "' AND " & _
                  "CF02 = '" & m_TM10 & "' AND " & _
                  "CF03 = '" & m_CP10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CF09")) = False Then
         textCF09 = rsTmp.Fields("CF09")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   'Add By Cheng 2002/11/05
   '若案件性質為領證, 隱藏大約多久回音
   If m_CP10 = "701" Then
        Me.Label1(12).Visible = False
        Me.textCF09.Visible = False
        Me.textCF09.Enabled = False
        Me.Label11.Visible = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm030101_18 = Nothing
End Sub

' 發文日
Private Sub textCP27_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strDay As String
   
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
      If Val(DBDATE(textCP27)) > Val(strSrvDate(1)) Then 'edit by nickc 2006/03/17 原先為 DBDATE(Date)) Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "發文日不可超過系統日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      
      'Add By Sindy 2009/08/19
      ' 計算催審期限
      If Me.textCP27.Tag <> Me.textCP27.Text Then 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
            strDay = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
            If IsEmptyText(strDay) = False Then
               textUargeDate = strDay
            End If
      End If
      Me.textCP27.Tag = Me.textCP27.Text 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
   End If
EXITSUB:
End Sub

Private Sub textCP44_Click()
   If textCP44.ListIndex >= 0 Then
      textCP44_2 = m_AgentList(textCP44.ListIndex).aiName
   End If
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

' 列印備註
Private Sub textPS_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textPS, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "列印備註內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textPS_GotFocus
   End If
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
   ' 發文日
   SetCPFieldNewData "CP27", DBDATE(textCP27)
   ' 代理人代號
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
   Dim strNP07 As String
   Dim strNP08 As String
   Dim strNP09 As String
   Dim strNP22 As String
   Dim bolIns305 As Boolean 'Modify By Sindy 2012/8/20
   Dim strNP10 As String 'Add By Sindy 2014/9/11
   
'911106 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   'Modified by Lydia 2016/03/11 +案號
   'Call GetNP69("", m_TM10, m_CP13, strNP10) 'Add By Sindy 2014/9/11
   'Modified by Lydia 2017/05/12 GetNP69更名為GetNA69
   Call GetNA69("", m_TM10, m_CP13, strNP10, m_TM01, m_TM02, m_TM03, m_TM04)
   
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
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若有定義代理人收達天數時, 新增一筆收達的記錄到下一程序檔
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & m_TM01 & "' AND " & _
                  "CF02 = '" & m_TM10 & "' AND " & _
                  "CF03 = '" & m_CP10 & "' "
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
         strNP08 = PUB_T997998LimitDate(strNP08, CStr(m_CP07), 1)
         '2019/6/11 END
         strNP22 = GetNextProgressNo()
         '92.10.6 MODIFY BY SONIA
'         strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
'                              strNP08 & "," & strNP08 & ",'" & strUserNum & "'," & strNP22 & ")"
         'Modify By Sindy 2014/9/11 m_CP14=>strNP10
         'Modified by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                              strNP08 & "," & strNP08 & ",'" & strNP10 & "'," & strNP22 & ")"
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
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
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   'Added by Lydia 2024/07/09 判斷案件國家收費表內有設定提申期限(天)CF11，要加掛提申(998)期限；
   Call Pub_GetCF11to998(m_TM10, m_TM01, m_TM02, m_TM03, m_TM04, m_CP07, m_CP09, m_CP10, m_CP14, textCP27)
   
'2015/12/4 modify by sonia 菲律賓還有原用期為20年的案件,故使用宣誓發文仍要掛下一次期限
'2011/9/23 CANCEL BY SONIA 因為現行3個國家都是專用10年延展10年,在每10年內只要提一次使用宣誓,在發註冊證掛期限即可
'   'Modify By Cheng 2002/06/05
'   '若案件性質為"使用宣誓"(105), 先抓國家檔的"下次使用宣誓年度"欄, 若為NULL則不產生下一程序檔的期限
'    'Modify By Cheng 2003/09/03
'    '若案件性質為使用宣誓且有專用期限
''   If m_CP10 = "105" Then
'   If m_CP10 = "105" And Val("" & m_TM21) <> "0" And Val("" & m_TM22) <> "0" Then
'      Dim rsA As New ADODB.Recordset
'      Dim StrSQLa As String
'      Dim ii As Integer
'      Dim dblNewDate As Double
'      ii = 0
'      StrSQLa = "Select NA39,NA38 From Nation Where NA01='" & m_TM10 & "' AND NA39 IS NOT NULL "
'      rsA.CursorLocation = adUseClient
'      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'ReDo:
'         ii = ii + 1
'         'Modify By Cheng 2003/09/02
''         dblNewDate = DBDATE(Format(DateSerial(Val(DBYEAR(m_TM21)) + Val(rsA.Fields(1).Value) + Val(rsA.Fields(0).Value * ii), Val(DBMONTH(m_TM21)), Val(DBDAY(m_TM21)))))
'         dblNewDate = DBDATE(DateAdd("yyyy", Val(rsA.Fields(1).Value) + Val(rsA.Fields(0).Value * ii), ChangeWStringToWDateString(DBDATE(m_TM21))))
'         '若大於專用期止日
'         If dblNewDate > m_TM22 Then
'            '無動作
'         '若小於等於專用期止日
'         Else
'            '若小於此筆資料的法定期限+2年
'            'Modify By Cheng 2003/09/02
''            If dblNewDate < DBDATE(Format(DateSerial(Val(DBYEAR(m_CP07)) + 2, Val(DBMONTH(m_CP07)), Val(DBDAY(m_CP07))))) Then
'            If dblNewDate < DBDATE(DateAdd("yyyy", 2, ChangeWStringToWDateString(DBDATE(m_CP07)))) Then
'               GoTo ReDo
'            '若大於等於此筆資料的法定期限+2年
'            Else
'               strNP22 = GetNextProgressNo()
'               '法定期限
'               strNP09 = dblNewDate
'               '本所期限
'                'Modify By Cheng 2003/09/02
''               strNP08 = IIf(m_TM10 = "030", DBDATE(Format(DateSerial(Val(DBYEAR(dblNewDate)) - 1, Val(DBMONTH(dblNewDate)), Val(DBDAY(dblNewDate))))), _
''                              DBDATE(Format(DateSerial(Val(DBYEAR(dblNewDate)), Val(DBMONTH(dblNewDate)) - 6, Val(DBDAY(dblNewDate))))))
'               'edit by nickc 2007/04/16 本所改成法定-2個月
'               'strNP08 = IIf(m_TM10 = "030", DBDATE(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(dblNewDate)))), _
'                              DBDATE(DateAdd("m", -6, ChangeWStringToWDateString(DBDATE(dblNewDate)))))
'               strNP08 = DBDATE(DateAdd("m", -2, ChangeWStringToWDateString(DBDATE(dblNewDate))))
'               '新增下一程序檔
'               strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                        "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',105," & _
'                                    strNP08 & "," & strNP09 & ",'" & m_CP13 & "'," & strNP22 & ")"
'               cnnConnection.Execute strSql
'            End If
'         End If
'      End If
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'   End If
   
   '2015/12/4 add by sonia 菲律賓還有原用期為20年的案件,故使用宣誓發文仍要掛下一次期限 CFT-005951
   '若案件性質為使用宣誓且有專用期限
   If m_CP10 = "105" And Val("" & m_TM21) <> "0" And Val("" & m_TM22) <> "0" Then
      'add by sonia 2015/12/14 若為菲律賓的申請日+3年使用宣誓期限則不必掛下一次
      If m_TM10 = "030" Then
         If m_TM11 + 30000 = m_CP07 Then
            GoTo Nextstep
         End If
      End If
      'end if
   
      Dim rsA As New ADODB.Recordset
      Dim StrSQLa As String
      Dim ii As Integer
      Dim dblNewDate As Double
      ii = 0
      StrSQLa = "Select NA38,NA39 From Nation Where NA01='" & m_TM10 & "' AND NA39 IS NOT NULL "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
ReDo:
         ii = ii + 1
         'modify by sonia 2022/10/21 318莫三比克以申請日計算算,無申請日才抓專用期起日
         If m_TM10 = "318" Then
            If m_TM11 > 0 Then
               dblNewDate = DBDATE(DateAdd("yyyy", Val(rsA.Fields("NA38").Value) + Val(rsA.Fields("NA39").Value * ii), ChangeWStringToWDateString(DBDATE(m_TM11))))
            Else
               dblNewDate = DBDATE(DateAdd("yyyy", Val(rsA.Fields("NA38").Value) + Val(rsA.Fields("NA39").Value * ii), ChangeWStringToWDateString(DBDATE(m_TM21))))
            End If
         'end 2022/10/21
         '以發證日計算,無發證日才抓專用期起日
         ElseIf m_TM20 > 0 Then
            dblNewDate = DBDATE(DateAdd("yyyy", Val(rsA.Fields("NA38").Value) + Val(rsA.Fields("NA39").Value * ii), ChangeWStringToWDateString(DBDATE(m_TM20))))
         Else
            dblNewDate = DBDATE(DateAdd("yyyy", Val(rsA.Fields("NA38").Value) + Val(rsA.Fields("NA39").Value * ii), ChangeWStringToWDateString(DBDATE(m_TM21))))
         End If
         '若大於專用期止日無動作
         'modify by sonia 2021/9/17 葉易雲於2021/8/27提出菲律賓2017/8/1新法延展核准後一年方再提出「延展使用宣誓」，故菲律賓不檢查專用期止日
         'If dblNewDate > m_TM22 Then
         If dblNewDate > m_TM22 And m_TM10 <> "030" Then
         '若小於等於專用期止日
         Else
            '若小於此筆資料的法定期限+2年,再計算下一次
            If dblNewDate < DBDATE(DateAdd("yyyy", 2, ChangeWStringToWDateString(DBDATE(m_CP07)))) Then
               GoTo ReDo
            '若大於等於此筆資料的法定期限+2年
            Else
               '法定期限
               strNP09 = dblNewDate
               '本所期限 業務說改成本所=法定-2個月 不管任何國家
               strNP08 = DBDATE(DateAdd("m", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
               strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
               'modify by sonia 2022/6/10 有期限則更新，無才新增
               'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               '         "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',105," & _
               '                     strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & GetNextProgressNo() & ")"
               strSql = "select * from NEXTPROGRESS where NP02='" & m_TM01 & "' and NP03='" & m_TM02 & "' and NP04='" & m_TM03 & "' and NP05='" & m_TM04 & "' AND NP07='105' AND NP06 IS NULL and NP09>" & m_CP07 + 20000
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If rsTmp.RecordCount <> 0 Then
                  strSql = "UPDATE NEXTPROGRESS SET NP08=" & strNP08 & ",NP09=" & strNP09 & ",NP10='" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "' where NP02='" & m_TM01 & "' and NP03='" & m_TM02 & "' and NP04='" & m_TM03 & "' and NP05='" & m_TM04 & "' AND NP07='105' AND NP06 IS NULL and NP09>" & m_CP07 + 20000
               Else
                  strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                           "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',105," & _
                                       strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & GetNextProgressNo() & ")"
               End If
               'end 2022/6/10
               cnnConnection.Execute strSql
            End If
         End If
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
   'end 2015/12/4
   
Nextstep:
   'Add By Sindy 2009/08/19
   ' 若有輸入催審期限時, 新增一筆催審的記錄到下一程序檔
   If IsEmptyText(textUargeDate) = False Then
      'Modify By Sindy 2012/8/20
      bolIns305 = True
'cancel by sonia 2023/3/23 也要催
'      'Modify By Sindy 2011/2/15
'      If m_TM10 = "030" And m_CP10 = "105" And Val(m_TM11) > 0 Then
'         If m_CP07 = DBDATE(DateAdd("yyyy", 3, ChangeWStringToWDateString(DBDATE(m_TM11)))) Then
'            ' 菲律賓030使用宣誓105發文, 法定期限為申請日+3年者則不掛催審期限
'            bolIns305 = False
'         End If
'      End If
'      '2011/2/15 End
'end 2023/3/23
      If bolIns305 = True Then
      '2012/8/20 End
         strNP07 = "305"
         'Modify By Sindy 2014/9/11 m_CP14=>strNP10
         'Modified by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "SELECT '" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                             DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & strNP10 & "',NP22 from (select nvl(max(np22),0)+1 NP22 from nextprogress)"
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "SELECT '" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                             PUB_GetWorkDay1(textUargeDate, True) & "," & DBDATE(textUargeDate) & ",'" & strNP10 & "',NP22 from (select nvl(max(np22),0)+1 NP22 from nextprogress)"
         cnnConnection.Execute strSql
      End If
   End If
   
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

' 檢查欄位是否都已輸入或是輸入的值是否正確
Private Function CheckDataValid() As Boolean
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
   'add by nickc 2006/03/17 加入驗證
   Dim Cancel As Boolean
   Cancel = False
   textCP27_Validate Cancel
   If Cancel = True Then GoTo EXITSUB
   
   'add by sonia 2015/12/14 菲律賓使用宣誓案一定要有申請日才可檢查是否為申請日+3年期限(不必掛下一次)
   If m_CP10 = "105" And m_TM10 = "030" And Val(m_TM11) = 0 Then
      strTit = "檢核資料"
      strMsg = "菲律賓使用宣誓案一定要有申請日才可檢查是否為申請日+3年期限"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   'end if
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textPS_GotFocus()
   InverseTextBox textPS
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

Private Sub textCF09_GotFocus()
   InverseTextBox textCF09
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
                     "','回音','" & textCF09 & "')"
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
   
   If Me.textPS.Enabled = True Then
      Cancel = False
      textPS_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2009/08/19
   If Me.textUargeDate.Enabled = True Then
      Cancel = False
      textUargeDate_Validate Cancel
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

'Add By Sindy 2009/08/19
' 催審期限
Private Sub textUargeDate_GotFocus()
   InverseTextBox textUargeDate
End Sub

Private Sub textUargeDate_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textUargeDate) = False Then
      If CheckIsDate(textUargeDate, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "催審期限日期不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textUargeDate_GotFocus
      End If
   End If
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
