VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010409_6 
   BorderStyle     =   1  '單線固定
   Caption         =   "服務業務結果輸入(監視系統)"
   ClientHeight    =   5250
   ClientLeft      =   140
   ClientTop       =   990
   ClientWidth     =   9130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   9130
   Begin VB.TextBox textCP54 
      Height          =   264
      Left            =   6840
      TabIndex        =   1
      Top             =   3600
      Width           =   1092
   End
   Begin VB.TextBox textCP53 
      Height          =   264
      Left            =   5400
      TabIndex        =   0
      Top             =   3600
      Width           =   1092
   End
   Begin VB.TextBox textSP50 
      Height          =   264
      Left            =   5400
      MaxLength       =   15
      TabIndex        =   3
      Top             =   3960
      Width           =   2532
   End
   Begin VB.CommandButton cmdCCC 
      Caption         =   "CCC Code(&C)"
      Height          =   400
      Left            =   3720
      TabIndex        =   6
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "變更事項(&R)"
      Height          =   400
      Left            =   4944
      TabIndex        =   7
      Top             =   70
      Width           =   1200
   End
   Begin VB.TextBox textSP21 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   2940
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1092
   End
   Begin VB.TextBox textSP20 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1092
   End
   Begin VB.TextBox textPrint 
      Height          =   264
      Left            =   1500
      MaxLength       =   1
      TabIndex        =   4
      Top             =   4320
      Width           =   492
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8220
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6168
      TabIndex        =   8
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6996
      TabIndex        =   9
      Top             =   70
      Width           =   1200
   End
   Begin VB.TextBox textCP08 
      Height          =   264
      Left            =   1500
      MaxLength       =   40
      TabIndex        =   2
      Top             =   3960
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2532
   End
   Begin VB.TextBox textSPKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   720
      Width           =   2532
   End
   Begin VB.TextBox textSP06 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1440
      Width           =   7512
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2532
   End
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5400
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2532
   End
   Begin MSForms.TextBox textCP64 
      Height          =   300
      Left            =   1500
      TabIndex        =   5
      Top             =   4680
      Width           =   7512
      VariousPropertyBits=   -1467989989
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13250;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5400
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2532
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP08 
      Height          =   264
      Left            =   1500
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2160
      Width           =   7512
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "13250;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP05 
      Height          =   264
      Left            =   1500
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1080
      Width           =   7512
      VariousPropertyBits=   679493663
      Size            =   "13250;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP07 
      Height          =   264
      Left            =   1500
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1800
      Width           =   7512
      VariousPropertyBits=   679493663
      Size            =   "13250;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line2 
      X1              =   6630
      X2              =   6750
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label1 
      Caption         =   "授權期間 :"
      Height          =   255
      Index           =   7
      Left            =   4560
      TabIndex        =   39
      Top             =   3630
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "BTTM :"
      Height          =   252
      Index           =   5
      Left            =   4560
      TabIndex        =   38
      Top             =   3984
      Width           =   1092
   End
   Begin VB.Line Line1 
      X1              =   2700
      X2              =   2820
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label1 
      Caption         =   "使用期間 :"
      Height          =   252
      Index           =   2
      Left            =   180
      TabIndex        =   37
      Top             =   3600
      Width           =   1092
   End
   Begin VB.Label Label21 
      Caption         =   "進度備註 :"
      Height          =   252
      Left            =   180
      TabIndex        =   36
      Top             =   4680
      Width           =   972
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   252
      Left            =   180
      TabIndex        =   35
      Top             =   4320
      Width           =   972
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
      Height          =   180
      Left            =   2070
      TabIndex        =   34
      Top             =   4350
      Width           =   2745
   End
   Begin VB.Label Label1 
      Caption         =   "機關文號 :"
      Height          =   252
      Index           =   4
      Left            =   180
      TabIndex        =   33
      Top             =   3960
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4560
      TabIndex        =   30
      Top             =   2880
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   252
      Index           =   10
      Left            =   180
      TabIndex        =   29
      Top             =   3240
      Width           =   1092
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   180
      TabIndex        =   28
      Top             =   2160
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   180
      TabIndex        =   27
      Top             =   720
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "案件中文名稱 :"
      Height          =   252
      Left            =   180
      TabIndex        =   26
      Top             =   1080
      Width           =   1332
   End
   Begin VB.Label Label9 
      Caption         =   "案件英文名稱 :"
      Height          =   252
      Left            =   180
      TabIndex        =   25
      Top             =   1440
      Width           =   1212
   End
   Begin VB.Label Label10 
      Caption         =   "案件日文名稱 :"
      Height          =   252
      Left            =   180
      TabIndex        =   24
      Top             =   1800
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   180
      TabIndex        =   23
      Top             =   2880
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   252
      Index           =   3
      Left            =   4560
      TabIndex        =   22
      Top             =   2520
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   1
      Left            =   180
      TabIndex        =   21
      Top             =   2520
      Width           =   732
   End
End
Attribute VB_Name = "frm02010409_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/29 Form2.0已修改 textSP05/textSP07/textSP08/textCP13/textCP64
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

' 本所案號
Dim m_SP01 As String
Dim m_SP02 As String
Dim m_SP03 As String
Dim m_SP04 As String
' 申請國家
Dim m_SP09 As String
' 商標審定號
Dim m_SP32 As String
' 來函收文日
Dim m_CP05 As String
' 所選取的收文號
Dim m_CP09 As String
' 案件性質
Dim m_CP10 As String
' 智權人員
Dim m_CP13 As String
Dim m_CP12 As String
Dim m_TM23 As String
'原承辦人  2015/1/14 add by sonia
Dim m_CP14 As String
'Add By Sindy 2019/5/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/22 END
Dim strLD18 As String 'Add By Sindy 2019/12/20 信函總收文號
Dim m_TM44 As String 'Add By Sindy 2019/12/20 FC代理人

'Add By Sindy 2019/5/22
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdCancel_Click()
   frm02010409_2.Show
   Unload Me
End Sub

Private Sub cmdCCC_Click()
   'frm02010409_10.SetData 0, m_SP01, True
   'frm02010409_10.SetData 1, m_SP02, False
   'frm02010409_10.SetData 2, m_SP03, False
   'frm02010409_10.SetData 3, m_SP04, False
   'frm02010409_10.Show
   'frm02010409_10.QueryData
   frmCCCCode.SetData 0, m_SP01, True
   frmCCCCode.SetData 1, m_SP02, False
   frmCCCCode.SetData 2, m_SP03, False
   frmCCCCode.SetData 3, m_SP04, False
   frmCCCCode.QueryData
   frmCCCCode.Show vbModal
   Unload frmCCCCode
   'Me.Hide
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm02010409_2
   Unload frm02010409_1
   Unload Me
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid() = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
          'add by nickc 2005/04/22
          Pub_EndModCashMsg m_SP09
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
        'Modify By Cheng 2002/11/07
'      'OnSaveData
        If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
        'Add By Cheng 2002/11/08
        ' 列印定稿
        If textPrint <> "N" Then
           PrintLetter
        End If
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      Unload frm02010409_2
      'Add By Sindy 2019/5/22
      If Me.m_strIR01 <> "" Then
         Unload frm02010409_1
         If Not m_PrevForm Is Nothing Then
            Call m_PrevForm.GoNext
         End If
      '2019/5/22 END
      Else
         frm02010409_1.Show
      End If
      Unload Me
   End If
End Sub

Private Sub cmdRelate_Click()
   frm02010409_8.SetData 0, m_SP01, True
   frm02010409_8.SetData 1, m_SP02, False
   frm02010409_8.SetData 2, m_SP03, False
   frm02010409_8.SetData 3, m_SP04, False
   frm02010409_8.SetData 4, m_CP09, False
   frm02010409_8.Show
   frm02010409_8.QueryData
   Me.Hide
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textSPKey.BackColor = &H8000000F
   textSP05.BackColor = &H8000000F
   textSP06.BackColor = &H8000000F
   textSP07.BackColor = &H8000000F
   textSP08.BackColor = &H8000000F
   textCP05.BackColor = &H8000000F
   textCP05S.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textSP20.BackColor = &H8000000F
   textSP21.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/22
   m_strIR01 = frm02010409_1.m_strIR01
   m_strIR02 = frm02010409_1.m_strIR02
   m_strIR03 = frm02010409_1.m_strIR03
   m_strIR04 = frm02010409_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/22 END
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_SP01 = Empty
      m_SP02 = Empty
      m_SP03 = Empty
      m_SP04 = Empty
      m_CP05 = Empty
      m_CP09 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_SP01 = strData
      ' 本所案號 欄位2
      Case 1: m_SP02 = strData
      ' 本所案號 欄位3
      Case 2: m_SP03 = strData
      ' 本所案號 欄位4
      Case 3: m_SP04 = strData
      ' 來函收文日
      Case 4: m_CP05 = strData
      ' 收文號
      Case 5: m_CP09 = strData
   End Select
End Sub

' 取得服務業務基本檔
Private Sub QueryServicePractice()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
      
   m_SP32 = Empty
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_SP01 & "' AND " & _
                  "SP02 = '" & m_SP02 & "' AND " & _
                  "SP03 = '" & m_SP03 & "' AND " & _
                  "SP04 = '" & m_SP04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      'Add By Cheng 2002/07/17
      m_SP09 = Empty
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_SP09 = rsTmp.Fields("SP09")
      End If
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("SP05")) = False Then
         textSP05 = rsTmp.Fields("SP05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("SP06")) = False Then
         textSP06 = rsTmp.Fields("SP06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("SP07")) = False Then
         textSP07 = rsTmp.Fields("SP07")
      End If
      ' 申請人
      'Add By Cheng 2002/07/17
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textSP08 = GetCustomerName(rsTmp.Fields("SP08"), 0)
         m_TM23 = rsTmp.Fields("SP08")
      End If
      
      'Add By Sindy 2019/12/20
      ' FC代理人
      m_TM44 = Empty
      If IsNull(rsTmp.Fields("SP26")) = False Then
         m_TM44 = rsTmp.Fields("SP26")
      End If
      '2019/12/20 END
      
      ' 專用期間(起)
      If IsNull(rsTmp.Fields("SP20")) = False Then
         textSP20 = TAIWANDATE(rsTmp.Fields("SP20"))
      End If
      ' 專用期間(迄)
      If IsNull(rsTmp.Fields("SP21")) = False Then
         textSP21 = TAIWANDATE(rsTmp.Fields("SP21"))
      End If
      ' 商標審定號
      If IsNull(rsTmp.Fields("SP32")) = False Then
         m_SP32 = rsTmp.Fields("SP32")
      End If
      ' 91.09.02 modify by louis
      If IsNull(rsTmp.Fields("SP50")) = False Then
         textSP50 = rsTmp.Fields("SP50")
      End If
      'add by nickc 2006/11/21
      textPrint = CheckStr(rsTmp.Fields("SP72"))
   End If

   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strCP53 As String
   Dim strCP54 As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   m_CP13 = Empty
   m_CP12 = Empty
   m_CP14 = Empty   '2015/1/14 add by sonia
   
   ' 來函收文日
   textCP05S = m_CP05
   ' 收文號
   textCP09 = m_CP09
   ' 讀取服務業務基本檔檔案
   QueryServicePractice
   
   ' 本所案號
   textSPKey = m_SP01 & m_SP02 & m_SP03 & m_SP04
   
   ' 取得案件進度檔
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_SP01 & "' AND " & _
                  "CP02 = '" & m_SP02 & "' AND " & _
                  "CP03 = '" & m_SP03 & "' AND " & _
                  "CP04 = '" & m_SP04 & "' AND " & _
                  "CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then
         textCP05 = TAIWANDATE(rsTmp.Fields("CP05"))
      End If
      ' 收文號
      If IsNull(rsTmp.Fields("CP09")) = False Then
         textCP09 = rsTmp.Fields("CP09")
      End If
      ' 案件性質
      m_CP10 = Empty
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_SP09 < "010" Then
            textCP10 = GetCaseTypeName(m_SP01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_SP01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"), True)
      End If
      '業務區   nick 91.08.22
      If IsNull(rsTmp.Fields("cp12")) = False Then
        m_CP12 = rsTmp.Fields("cp12")
      End If
      'Add By Sindy 2012/3/30
      '2014/1/27 MODIFY BY SONIA 加809
      If m_CP10 = "502" Or m_CP10 = "809" Then '授權502 進出口監視備案809
         textCP53.Visible = True
         textCP53.Enabled = True
         If m_CP10 = "809" Then
            Label1(7).Caption = "保護期間 :"
         Else
            Label1(7).Caption = "授權期間 :"
         End If
      Else
         textCP53.Visible = False
         textCP53.Enabled = False
         'add by sonia 2023/11/23
         textCP54.Visible = False
         textCP54.Enabled = False
         'end 2023/11/23
      End If
      '2012/3/30 End
      ' 授權期間(起)
      If IsNull(rsTmp.Fields("CP53")) = False Then
         strCP53 = DBDATE(rsTmp.Fields("CP53"))
         textCP53 = TAIWANDATE(rsTmp.Fields("CP53")) 'Add By Sindy 2012/3/30
      End If
      ' 授權期間(迄)
      If IsNull(rsTmp.Fields("CP54")) = False Then
         strCP54 = DBDATE(rsTmp.Fields("CP54"))
         textCP54 = TAIWANDATE(rsTmp.Fields("CP54")) 'Add By Sindy 2012/3/30
      End If
      m_CP14 = "" & rsTmp("CP14").Value  '2015/1/14 add by sonia
   End If
   rsTmp.Close
   
   Select Case m_CP10
      ' 監視系統申請
      '2014/1/27 MODIFY BY SONIA 加809
      Case "801", "809":
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM15 = '" & m_SP32 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            ' 使用期間 (起)
            If IsNull(rsTmp.Fields("TM21")) = False Then
               textSP20 = TAIWANDATE(rsTmp.Fields("TM21"))
            End If
            ' 使用期間 (迄)
            If IsNull(rsTmp.Fields("TM22")) = False Then
               textSP21 = TAIWANDATE(rsTmp.Fields("TM22"))
            End If
         End If
         rsTmp.Close
      ' 變更
      Case "301":
         If strCP53 <> "" Then '2011/9/1 add by sonia 加入有值才做條件 TM-000053
            textSP20 = TAIWANDATE(strCP53)
            textSP21 = TAIWANDATE(strCP54)
         End If
      'add by sonia 2023/11/23
      ' 監視系統TM延展102
      Case "102"
         If strCP53 <> "" Then '2011/9/1 add by sonia 加入有值才做條件 TM-000053
            textSP20 = DBDATE(strCP53)
            textSP21 = DBDATE(strCP54)
         End If
      'end 2023/11/23
   End Select
   
   ' 設定變更事項的按紐狀態
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      cmdRelate.Enabled = True
   Else
      cmdRelate.Enabled = False
   End If
   rsTmp.Close
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 檢查專用期間
   If IsEmptyText(textSP20) = True Or IsEmptyText(textSP21) = True Then
      strTit = "錯誤"
      strMsg = "專用期間資料不存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
   
   Set rsTmp = Nothing
   
   'add by nickc 2006/06/30 帶列印定稿預設值
   'edit by nickc 2006/11/21
   If textPrint = "" Then
        textPrint = GetTWordLng(m_SP01, m_SP02, m_SP03, m_SP04)
   End If
End Sub

' 儲存資料
'Modify By Cheng 2002/11/07
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strCP09 As String
   Dim strCP10 As String
   Dim strCP12 As String
   Dim strCP27 As String
   Dim strNP07 As String
   Dim strNP08 As String
   Dim strNP09 As String
   Dim strNP22 As String
   Dim strCon As String 'Add By Sindy 2012/3/30
   
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   
   ' 91.09.02 modify by louis
   ' 更新服務業務基本檔
   strSql = Empty
   If IsEmptyText(textSP50) Then
      'Modify By Sindy 2009/05/07
'      strSQL = "UPDATE SERVICEPRACTICE SET SP50 = 'NULL' " & _
'               "WHERE SP01 = '" & m_SP01 & "' AND " & _
'                     "SP02 = '" & m_SP02 & "' AND " & _
'                     "SP03 = '" & m_SP03 & "' AND " & _
'                     "SP04 = '" & m_SP04 & "'"
   Else
      strSql = "UPDATE SERVICEPRACTICE SET SP50 = '" & textSP50 & "' " & _
               "WHERE SP01 = '" & m_SP01 & "' AND " & _
                     "SP02 = '" & m_SP02 & "' AND " & _
                     "SP03 = '" & m_SP03 & "' AND " & _
                     "SP04 = '" & m_SP04 & "'"
      cnnConnection.Execute strSql
   End If
   
   ' 新增一筆資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   ' 案件性質為服務業務結果
   strCP10 = "1801"
   ' 業務區別
   'strCP12 = GetST15(m_CP13)
   ' 發文日為系統日
   strCP27 = DBDATE(SystemDate())
   ' 91.03.25 modify by louis (單引號)
    '承辦人為使用者, 發文日為系統日
    'Modify By Cheng 2003/04/03
    '智權人員存最近收文A類接洽記錄單的智權人員
    'Modify By Cheng 2004/02/03
    '業務區為最近收文A類接洽記錄單智權人員的業務區
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64) " & _
'            "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
'                    "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & strUserNum & "'," & _
'                    "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
   '2015/1/14 modify by sonia 所有服務業務結果的承辦人改放原承辦人TM-000067(宋若蘭),否則期限表帶出之承辦人會是程序
   'strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64) " & _
            "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04)) & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & strUserNum & "'," & _
                    "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64) " & _
            "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04)) & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & m_CP14 & "'," & _
                    "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
    'End
   '2014/1/27 ADD BY SONIA
   If textCP53.Enabled = True And m_CP10 = "809" Then
      '2015/1/14 modify by sonia 所有服務業務結果的承辦人改放原承辦人TM-000067(宋若蘭),否則期限表帶出之承辦人會是程序
      'strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64,CP53,CP54) " & _
               "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04)) & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & strUserNum & "'," & _
                    "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "'," & IIf(textCP53 = "", Null, DBDATE(textCP53)) & "," & IIf(textCP54 = "", Null, DBDATE(textCP54)) & ")"
      strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64,CP53,CP54) " & _
               "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04)) & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & m_CP14 & "'," & _
                    "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "'," & IIf(textCP53 = "", Null, DBDATE(textCP53)) & "," & IIf(textCP54 = "", Null, DBDATE(textCP54)) & ")"
   End If
   '2014/1/27 END
   cnnConnection.Execute strSql
         
   'Add By Sindy 2019/12/20 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
      strLD18 = strCP09
      PUB_AddLetterProgress strLD18, 1, IIf(textPrint = "N", False, True), "", False, m_TM23, strCP10, m_TM44
   End If
   '2019/12/20 END
   
    'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
    Pub_UpdateFromMaxCP27 m_SP01, m_SP02, m_SP03, m_SP04
                       
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新服務業務基本檔的使用期間欄位
   'modify by sonia 2023/11/23 加非延展102條件
   If IsEmptyText(textSP20) = False And m_CP10 <> "102" Then
      strSql = "UPDATE ServicePractice SET SP20 = " & DBDATE(textSP20) & " " & _
            "WHERE SP01 = '" & m_SP01 & "' AND " & _
                  "SP02 = '" & m_SP02 & "' AND " & _
                  "SP03 = '" & m_SP03 & "' AND " & _
                  "SP04 = '" & m_SP04 & "' "
      cnnConnection.Execute strSql
   End If
   If IsEmptyText(textSP21) = False Then
      strSql = "UPDATE ServicePractice SET SP21 = " & DBDATE(textSP21) & " " & _
            "WHERE SP01 = '" & m_SP01 & "' AND " & _
                  "SP02 = '" & m_SP02 & "' AND " & _
                  "SP03 = '" & m_SP03 & "' AND " & _
                  "SP04 = '" & m_SP04 & "' "
      cnnConnection.Execute strSql
   End If
   
   'add by nickc 2006/11/21
   If textPrint <> "N" Then
        strSql = "UPDATE ServicePractice SET SP72 = '" & textPrint & "' " & _
            "WHERE SP01 = '" & m_SP01 & "' AND " & _
                  "SP02 = '" & m_SP02 & "' AND " & _
                  "SP03 = '" & m_SP03 & "' AND " & _
                  "SP04 = '" & m_SP04 & "' "
        cnnConnection.Execute strSql
   End If
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新案件進度檔所選取收文資料的實際結果為 1
   'Modify By Sindy 2012/3/30 +授權期間
   strCon = ""
   '2014/1/27 MODIFY BY SONIA 加m_CP10 = "502"
   If textCP53.Enabled = True And m_CP10 = "502" Then
      strCon = ",CP53=" & IIf(textCP53 = "", Null, DBDATE(textCP53)) & ",CP54=" & IIf(textCP54 = "", Null, DBDATE(textCP54))
   End If
   '2012/3/30 End
   strSql = "UPDATE CaseProgress SET CP24='1'" & strCon & _
            " WHERE CP09='" & m_CP09 & "' "
   cnnConnection.Execute strSql
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 新增一筆資料到下一程序檔
   If IsEmptyText(textSP21) = False And m_CP10 = "801" Then
      ' 下一程序為延展
      strNP07 = "102"
      ' 法定期限為專用期限截止日
      strNP09 = DBDATE(textSP21)
      ' 本所期限為法定期限-2天
        'Modify By Cheng 2003/09/02
'      strNP08 = DBDATE(DateSerial(Val(DBYEAR(strNP09)), Val(DBMONTH(strNP09)), Val(DBDAY(strNP09)) - 2))
      'Modify By Sindy 2014/10/6 台灣案之本所期限設定
      If m_SP09 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
         strNP08 = PUB_GetOurDeadline(DBDATE(strNP09))
      Else
      '2014/10/6 END
         strNP08 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
      End If
      strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      ' 序號
      strNP22 = GetNextProgressNo()
      ' SQL 語法
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & strCP09 & "','" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & _
                       "'" & strNP07 & "'," & strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "'," & strNP22 & ")"
      cnnConnection.Execute strSql
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
'      '92.6.8 SONIA 加 言詞辯論, 準備程序
      Select Case strNP07
'         Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
         Case "102", "105", "702", "708", "305", "998", "997"
         Case Else:
            ' 列印國內案件接洽及結案記錄單
'             g_PrtForm001.PrintForm strNP22, m_SP01, m_SP02, m_SP03, m_SP04
            'Add By Cheng 2004/04/08
            '新增列印接洽結案單資料
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_SP01, "" & m_SP02, "" & m_SP03, "" & m_SP04
      End Select
   End If
   
   '2014/1/27 ADD BY SONIA 809 新增一筆資料到下一程序檔
   If IsEmptyText(textSP21) = False And m_CP10 = "809" Then
      ' 下一程序為809
      strNP07 = "809"
      ' 法定期限為專用期限截止日
      strNP09 = DBDATE(textCP54)
      ' 本所期限為法定期限-2天
      'Modify By Sindy 2014/10/6 台灣案之本所期限設定
      If m_SP09 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
         strNP08 = PUB_GetOurDeadline(DBDATE(strNP09))
      Else
      '2014/10/6 END
         strNP08 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
      End If
      strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      ' 序號
      strNP22 = GetNextProgressNo()
      ' SQL 語法
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & strCP09 & "','" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & _
                       "'" & strNP07 & "'," & strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "'," & strNP22 & ")"
      cnnConnection.Execute strSql
   End If
   
   '更新下一程序催審期限為Y
   strSql = "UPDATE NEXTProgress SET NP06='Y' WHERE NP01='" & m_CP09 & "' AND NP07='305' AND NP06 IS NULL"
   cnnConnection.Execute strSql
   '2014/1/27 END
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Modify By Cheng 2002/11/08
'   ' 列印定稿
'   If textPrint <> "N" Then
'      PrintLetter
'   End If

   'add by nickc 2005/04/22
   Pub_UpdateEndModCash m_SP01, m_SP02, m_SP03, m_SP04
   
   'Add by Sindy 2019/5/22
   Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010409_1"
   End If
   '2019/5/22 END
   
   'Add By Cheng 2002/11/07
   cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False
End Function

Private Sub Form_Unload(Cancel As Integer)
'edit by nickc 2008/04/25 改整批印
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   
   'Add By Sindy 2019/5/22
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   'Add By Cheng 2002/07/18
   Set frm02010409_6 = Nothing
End Sub

'Add By Sindy 2012/3/30
Private Sub textCP53_GotFocus()
   InverseTextBox textCP53
End Sub

'Add By Sindy 2012/3/30
Private Sub textCP53_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP53) = False Then
      ' 檢查是否為民國日期
      If CheckIsTaiwanDate(textCP53, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的授權期間"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP53_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2012/3/30
Private Sub textCP54_GotFocus()
   InverseTextBox textCP54
End Sub

'Add By Sindy 2012/3/30
Private Sub textCP54_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP54) = False Then
      ' 檢查是否為民國日期
      If CheckIsTaiwanDate(textCP54, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的授權期間"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP54_GotFocus
      End If
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
      strMsg = "進度備註欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'add by nickc 2006/06/29
   If KeyAscii <> 78 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 And KeyAscii <> 13 Then
       KeyAscii = 0
   End If
End Sub

' 檢查是否列印定稿欄位
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         'edit by nickc 2006/06/29
         'Case " ", "N":
         Case "N", "1", "2", "3":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            'edit by nickc 2006/06/29
            'strMsg = "只可輸入空白或N"
            strMsg = "只可輸入空白 或 N 或 1-3 "
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

' 檢查該輸入的資料是否已完成
Private Function CheckDataValid()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'Add by Amy 2021/12/29檢查畫面的 TextBox是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        GoTo EXITSUB
   End If

   ' 使用期間
   If IsEmptyText(textSP20) = True Or IsEmptyText(textSP21) = True Then
      strTit = "資料檢核"
      strMsg = "使用期間不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP20.SetFocus
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textCP08_GotFocus()
   InverseTextBox textCP08
   OpenIme
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
   Dim strTM23Nation As String
   Dim strSql As String
   Dim strTmp As String
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   If m_SP01 = "TM" Then
      Select Case m_CP10
         'Add By Sindy 2012/3/30
         Case "502" '授權
            If m_SP09 = "020" Then
               If textPrint = "1" Then '1:台->大
                  EndLetter "06", m_CP09, "01", strUserNum
                  '授權人資料
                  strSql = "SELECT * FROM CaseProgress WHERE CP09='" & m_CP09 & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     If "" & RsTemp.Fields("CP72") = "" Then
                        strTmp = "" & RsTemp("CP50") & "" & RsTemp("CP51") & "" & RsTemp("CP52")
                     Else
                        strTmp = "" & RsTemp("CP72")
                        strSql = "SELECT * FROM customer WHERE CU01='" & Left(Trim(strTmp), 8) & "' and CU02='" & Mid(Trim(strTmp), 9, 1) & "'"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           If "" & RsTemp("CU104") <> "" Then
                              strTmp = "" & RsTemp("CU104")
                           ElseIf "" & RsTemp("CU04") <> "" Then
                              strTmp = "" & RsTemp("CU04")
                           ElseIf "" & RsTemp("CU05") <> "" Then
                              strTmp = "" & RsTemp("CU05") & " " & RsTemp("CU88") & " " & RsTemp("CU89") & " " & RsTemp("CU90")
                           Else
                              strTmp = "" & RsTemp("CU06")
                           End If
                        End If
                     End If
                  End If
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "06" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
                        "','授權人資料','" & strTmp & "')"
                  cnnConnection.Execute strSql
               End If
            End If
            
         ' 案件性質為監視系統申請
         Case "801"
            'add by nickc 2006/05/16 新增申請國家是台灣的定稿
            If m_SP09 = "000" Then
               'add by nickc 2006/06/30
               If textPrint = "1" Then '1:台->台
                  EndLetter "06", m_CP09, "07", strUserNum
               'add by nickc 2006/06/30
               ElseIf textPrint = "2" Then '2:外->台
                  EndLetter "06", m_CP09, "08", strUserNum
               End If
            ElseIf m_SP09 = "020" Then
               'Add By Sindy 2009/05/07
               If textPrint = "1" Then '1.台->大
                  EndLetter "06", m_CP09, "09", strUserNum
               End If
            End If
            
         Case "602" '侵權處理
            If m_SP09 <> "000" Then
               If textPrint = "1" Then
                  EndLetter "06", m_CP09, "00", strUserNum
               End If
            End If
         '2014/1/27 ADD BY SONIA
         Case "809"
            If m_SP09 = "000" Then
               If textPrint = "1" Then '1:台->台
                  EndLetter "06", m_CP09, "07", strUserNum
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "06" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
                        "','補文件 V 1','" & textCP53 & "')"
                  cnnConnection.Execute strSql
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "06" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
                        "','補文件 V 2','" & textCP54 & "')"
                  cnnConnection.Execute strSql
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "06" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
                        "','機關文號','" & textCP08 & "')"
                  cnnConnection.Execute strSql
               ElseIf textPrint = "2" Then '2:外->台
                  EndLetter "06", m_CP09, "08", strUserNum
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "06" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & _
                        "','補文件 V 1','" & textCP53 & "')"
                  cnnConnection.Execute strSql
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "06" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & _
                        "','補文件 V 2','" & textCP54 & "')"
                  cnnConnection.Execute strSql
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "06" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & _
                        "','機關文號','" & textCP08 & "')"
                  cnnConnection.Execute strSql
               End If
            End If
         '2014/1/27 END
      End Select
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
Dim strTM23Nation As String
'Add By Sindy 2012/1/13
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/13 End
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   'Add By Sindy 2012/1/13
   ET01 = "06"
   ET02 = m_CP09
   bolEdit = False
   '2012/1/13 End
   
   ' 系統別為TB
   If m_SP01 = "TM" Then
      Select Case m_CP10
         'Add By Sindy 2012/3/30
         Case "502" '授權
            If m_SP09 = "020" Then
               If textPrint = "1" Then '1:台->大
                  ET03 = "01"
               End If
            End If
            
         ' 監視系統申請
         '2014/1/27 modify by sonia 加809
         Case "801", "809"
            'add by nickc 2006/05/16 新增大->台灣的定稿
            If m_SP09 = "000" Then
               'add by nickc 2006/06/30
               If textPrint = "1" Then '1:台->台
'                  NowPrint m_CP09, "06", "07", False, strUserNum, 0
                  ET03 = "07" 'Modify By Sindy 2012/1/13
               ElseIf textPrint = "2" Then '2:外->台
'                  NowPrint m_CP09, "06", "08", False, strUserNum, 0
                  ET03 = "08" 'Modify By Sindy 2012/1/13
               End If
            ElseIf m_SP09 = "020" Then
               'Add By Sindy 2009/05/07
               If textPrint = "1" Then '1:台->大
'                  NowPrint m_CP09, "06", "09", False, strUserNum, 0
                  ET03 = "09" 'Modify By Sindy 2012/1/13
               End If
            End If
         '2011/6/8 ADD BY SONIA 新增台->大侵權處理的定稿
         Case "602" '侵權處理
            If m_SP09 <> "000" Then
               If textPrint = "1" Then
'                  NowPrint m_CP09, "06", "00", False, strUserNum, 0
                  ET03 = "00" 'Modify By Sindy 2012/1/13
               End If
            End If
         '2011/6/8 END
      End Select
   End If
   
   'Add By Sindy 2012/1/13
   If ET03 <> "" Then
      bolEmail = PUB_GetEMailFlag(m_SP01 & m_SP02 & m_SP03 & m_SP04, , , bolPlusPaper)
      If bolEmail Then
         '判斷是否EMail同時寄紙本
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'Add By Sindy 2020/1/7 + 信函總收文號
         If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , strLD18
         Else
         '2020/1/7 END
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True
            MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_SP01) & " ]！"
         End If
      Else
         'Add By Sindy 2019/12/20 + strLD18.信函總收文號
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
      End If
   'Add By Sindy 2021/1/5 沒有系統產出的定稿
   Else
      'Add By Sindy 2021/2/1 詢問有沒有客戶函
      If strLD18 <> "" Then
         Call PUB_TCaseAskIsPost_C(strLD18)
      End If
   '2021/1/5 EMD
   End If
   '2012/1/13 End
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
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

TxtValidate = True
End Function
