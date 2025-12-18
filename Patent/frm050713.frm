VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050713 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件聯絡人修改作業"
   ClientHeight    =   5565
   ClientLeft      =   240
   ClientTop       =   990
   ClientWidth     =   8400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8400
   Begin VB.TextBox textSys 
      BorderStyle     =   0  '沒有框線
      Height          =   492
      Left            =   1500
      MultiLine       =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   540
      Width           =   6732
   End
   Begin VB.TextBox textCUFA_2 
      Height          =   300
      Left            =   3180
      MaxLength       =   1
      TabIndex        =   3
      Top             =   2100
      Width           =   252
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6360
      TabIndex        =   10
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   7320
      TabIndex        =   11
      Top             =   60
      Width           =   912
   End
   Begin VB.TextBox textCUFA 
      Height          =   300
      Left            =   2220
      MaxLength       =   8
      TabIndex        =   2
      Top             =   2100
      Width           =   972
   End
   Begin VB.OptionButton optCUFA 
      Caption         =   "代理人編號"
      Height          =   252
      Index           =   1
      Left            =   1500
      TabIndex        =   1
      Top             =   1500
      Width           =   1452
   End
   Begin VB.OptionButton optCUFA 
      Caption         =   "客戶編號"
      Height          =   252
      Index           =   0
      Left            =   1500
      TabIndex        =   0
      Top             =   1140
      Value           =   -1  'True
      Width           =   1332
   End
   Begin MSForms.TextBox textNewJP 
      Height          =   300
      Left            =   2220
      TabIndex        =   9
      Top             =   4980
      Width           =   3375
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "5953;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNewEN 
      Height          =   300
      Left            =   2220
      TabIndex        =   8
      Top             =   4660
      Width           =   3375
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "5953;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNewCH 
      Height          =   300
      Left            =   2220
      TabIndex        =   7
      Top             =   4340
      Width           =   3375
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "5953;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textOldJP 
      Height          =   300
      Left            =   2220
      TabIndex        =   6
      Top             =   4020
      Width           =   3375
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "5953;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textOldEN 
      Height          =   300
      Left            =   2220
      TabIndex        =   5
      Top             =   3700
      Width           =   3375
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "5953;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textOldCH 
      Height          =   300
      Left            =   2220
      TabIndex        =   4
      Top             =   3380
      Width           =   3375
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "5953;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCU06 
      Height          =   300
      Left            =   2220
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3060
      Width           =   6015
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "10610;529"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCU05 
      Height          =   300
      Left            =   2220
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2740
      Width           =   6015
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "10610;529"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCU04 
      Height          =   300
      Left            =   2220
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2420
      Width           =   6015
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "10610;529"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      Caption         =   "系統類別 : "
      Height          =   252
      Left            =   300
      TabIndex        =   23
      Top             =   660
      Width           =   972
   End
   Begin VB.Line Line2 
      X1              =   270
      X2              =   8190
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   300
      X2              =   8220
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label11 
      Caption         =   "新聯絡人(日) :"
      Height          =   255
      Left            =   300
      TabIndex        =   22
      Top             =   5003
      Width           =   1332
   End
   Begin VB.Label Label10 
      Caption         =   "新聯絡人(英) :"
      Height          =   255
      Left            =   300
      TabIndex        =   21
      Top             =   4683
      Width           =   1332
   End
   Begin VB.Label Label9 
      Caption         =   "新聯絡人(中) :"
      Height          =   255
      Left            =   300
      TabIndex        =   20
      Top             =   4363
      Width           =   1332
   End
   Begin VB.Label Label8 
      Caption         =   "原聯絡人(日) :"
      Height          =   255
      Left            =   300
      TabIndex        =   19
      Top             =   4043
      Width           =   1332
   End
   Begin VB.Label Label7 
      Caption         =   "原聯絡人(英) :"
      Height          =   255
      Left            =   300
      TabIndex        =   18
      Top             =   3723
      Width           =   1332
   End
   Begin VB.Label Label6 
      Caption         =   "原聯絡人(中) :"
      Height          =   255
      Left            =   300
      TabIndex        =   17
      Top             =   3403
      Width           =   1332
   End
   Begin VB.Label Label5 
      Caption         =   "日文名稱 :"
      Height          =   255
      Left            =   300
      TabIndex        =   16
      Top             =   3083
      Width           =   1092
   End
   Begin VB.Label Label4 
      Caption         =   "英文名稱 :"
      Height          =   255
      Left            =   300
      TabIndex        =   15
      Top             =   2763
      Width           =   1092
   End
   Begin VB.Label Label3 
      Caption         =   "中文名稱 :"
      Height          =   255
      Left            =   300
      TabIndex        =   14
      Top             =   2443
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "客戶 / 代理人編號 :"
      Height          =   255
      Left            =   300
      TabIndex        =   13
      Top             =   2123
      Width           =   1812
   End
   Begin VB.Label Label1 
      Caption         =   "修改對象 :"
      Height          =   252
      Left            =   300
      TabIndex        =   12
      Top             =   1344
      Width           =   972
   End
End
Attribute VB_Name = "frm050713"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/23 改成Form2.0 ;textCU04、textCU05、textCU06、textOldCH、textOldEN、textOldJP、textNewCH、textNewEN、textNewJP
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit
Dim m_QuerySys As String
Dim m_SysCVS As String
Dim m_CUFASel As Integer
Dim m_CUFANo As String

Dim m_RecordAffect As Integer

Private Sub Form_Load()
   MoveFormToCenter Me

   textSys.BackColor = &H8000000F
   textCU04.BackColor = &H8000000F
   textCU05.BackColor = &H8000000F
   textCU06.BackColor = &H8000000F

   QueryLimit
End Sub

Private Sub cmdok_Click()
'add by nickc 2007/02/08
Dim strTit As String
Dim strMsg As String
Dim nResponse
   If CheckDataValid() = True Then
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 設定客戶/FC代理人代碼
      m_CUFANo = textCUFA & String(8 - Len(textCUFA), "0") & textCUFA_2 & String(1 - Len(textCUFA_2), "0")
      ' 設定更新的是客戶還是FC代理人
      If optCUFA(0) = True Then
         m_CUFASel = 0
      Else
         m_CUFASel = 1
      End If
      ' 清除筆數
      m_RecordAffect = 0
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
      
      ' 更新資料
      'edit by nick 2004/11/03
      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub

      ' 判斷筆數
      If m_RecordAffect <= 0 Then
         strTit = "更新資料"
         strMsg = "無符合條件的聯絡人可修改"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Else
         ' 清除欄位資料
         ClearField
      End If
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      ' 設定輸入欄
      textCUFA.SetFocus
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub ClearField()
   textCUFA = Empty
   textCUFA_2 = Empty
   textCU04 = Empty
   textCU05 = Empty
   textCU06 = Empty
   textOldCH = Empty
   textOldEN = Empty
   textOldJP = Empty
   textNewCH = Empty
   textNewEN = Empty
   textNewJP = Empty
End Sub

Private Sub QueryLimit()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strGroup As String
   Dim strSys As String
   
   strSql = "SELECT ST11 FROM Staff " & _
            "WHERE ST01 = '" & strUserNum & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("ST11")) = False Then
         strGroup = rsTmp.Fields("ST11")
      End If
   End If
   rsTmp.Close
   
   strSys = Empty
   m_QuerySys = Empty
   strSql = "SELECT DISTINCT(SG02) FROM STAFF_GROUP " & _
            "WHERE SG01 = '" & strGroup & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Do While rsTmp.EOF = False
         If IsEmptyText(strSys) = False Then: strSys = strSys & ","
         If IsEmptyText(m_QuerySys) = False Then: m_QuerySys = m_QuerySys & ","
         strSys = strSys & rsTmp.Fields("SG02")
         m_QuerySys = m_QuerySys & "'" & rsTmp.Fields("SG02") & "'"
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close

   m_SysCVS = strSys
   textSys = strSys
   
   Set rsTmp = Nothing
End Sub

' 更新檔案
'edit b nick 2004/11/03
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   
'911106 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   ' 更新商標基本檔
   If CheckUpdateSystem(0) = True Then
      OnUpdateTradeMark
   End If
   ' 更新專利基本檔
   If CheckUpdateSystem(1) = True Then
      OnUpdatePatent
   End If
   ' 更新法務基本檔
   If CheckUpdateSystem(2) = True Then
      OnUpdateLawCase
   End If
   ' 更新服務業務基本檔
   If CheckUpdateSystem(3) = True Then
      OnUpdateServicePractice
   End If
'911106 nick transation
   cnnConnection.CommitTrans
Exit Function
CheckingErr:
     MsgBox (Err.Description)
     cnnConnection.RollbackTrans
     OnSaveData = False
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 檢查是否需要更新檔案
' Input : strSys ==> 所要檢查的系統
'                0 : 表是否要更新商標基本檔
'                1 : 表是否要更新專利基本檔
'                2 : 表是否要更新法務基本檔
'                3 : 表是否要更新服務業務基本檔
' Output : True 表示要更新
'          False 表示不更新
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckUpdateSystem(ByVal strSys As Integer) As Boolean
   'add by nickc 2007/02/08
   Dim nIndex, strTemp
   
   CheckUpdateSystem = False
   For nIndex = 1 To GetSubStringCount(m_SysCVS)
      strTemp = GetSubString(m_SysCVS, nIndex)
      Select Case strTemp
         ' 讀取商標基本檔
         Case "T", "TF", "CFT", "FCT":
            If strSys = 0 Then
               CheckUpdateSystem = True
               GoTo EXITSUB
            End If
         ' 讀取專利基本檔
         Case "P", "CFP", "FCP":
            If strSys = 1 Then
               CheckUpdateSystem = True
               GoTo EXITSUB
            End If
         ' 讀取法務基本檔
         'Modify By Sindy 2009/07/24 增加LIN系統類別
         'modify by sonia 2019/7/29 +ACS系統類別
         Case "L", "CFL", "FCL", "LIN", "ACS":
            If strSys = 2 Then
               CheckUpdateSystem = True
               GoTo EXITSUB
            End If
         ' 讀取服務業務基本檔
         Case Else:
            If strSys = 3 Then
               CheckUpdateSystem = True
               GoTo EXITSUB
            End If
      End Select
   Next nIndex
   
EXITSUB:
End Function

' 更新商標基本檔
Private Sub OnUpdateTradeMark()
   Dim strSql As String
   Dim nRecords As Integer
   
   nRecords = 0
   ' 更新商標基本檔
   If m_CUFASel = 0 Then
      'Modify By Sindy 2011/2/24 增加TM78,TM79,TM80,TM81
      strSql = "UPDATE TRADEMARK SET TM38 = DECODE(TM38, '" & textOldCH & "','" & textNewCH & "',TM38) , " & _
                                    "TM39 = DECODE(TM39, '" & textOldEN & "','" & textNewEN & "',TM39) , " & _
                                    "TM40 = DECODE(TM40, '" & textOldJP & "','" & textNewJP & "',TM40) , " & _
                                    "TM41 = DECODE(TM41, '" & textOldCH & "','" & textNewCH & "',TM41) , " & _
                                    "TM42 = DECODE(TM42, '" & textOldEN & "','" & textNewEN & "',TM42) , " & _
                                    "TM43 = DECODE(TM43, '" & textOldJP & "','" & textNewJP & "',TM43) , " & _
                                    "TM55 = DECODE(TM55, '" & textOldEN & "','" & textNewEN & "',TM55) " & _
               "WHERE TM01 IN (" & m_QuerySys & ") AND " & _
                     "(TM23 = '" & m_CUFANo & "' OR " & _
                     "TM78 = '" & m_CUFANo & "' OR " & _
                     "TM79 = '" & m_CUFANo & "' OR " & _
                     "TM80 = '" & m_CUFANo & "' OR " & _
                     "TM81 = '" & m_CUFANo & "') AND " & _
                     "(TM38 = '" & textOldCH & "' OR " & _
                      "TM39 = '" & textOldEN & "' OR " & _
                      "TM40 = '" & textOldJP & "' OR " & _
                      "TM41 = '" & textOldCH & "' OR " & _
                      "TM42 = '" & textOldEN & "' OR " & _
                      "TM43 = '" & textOldJP & "' OR " & _
                      "TM55 = '" & textOldEN & "') "
      cnnConnection.Execute strSql, nRecords
   Else
      strSql = "UPDATE TRADEMARK SET TM38 = DECODE(TM38, '" & textOldCH & "','" & textNewCH & "',TM38) , " & _
                                    "TM39 = DECODE(TM39, '" & textOldEN & "','" & textNewEN & "',TM39) , " & _
                                    "TM40 = DECODE(TM40, '" & textOldJP & "','" & textNewJP & "',TM40) , " & _
                                    "TM41 = DECODE(TM41, '" & textOldCH & "','" & textNewCH & "',TM41) , " & _
                                    "TM42 = DECODE(TM42, '" & textOldEN & "','" & textNewEN & "',TM42) , " & _
                                    "TM43 = DECODE(TM43, '" & textOldJP & "','" & textNewJP & "',TM43) , " & _
                                    "TM55 = DECODE(TM55, '" & textOldEN & "','" & textNewEN & "',TM55) " & _
               "WHERE TM01 IN (" & m_QuerySys & ") AND " & _
                     "TM44 = '" & m_CUFANo & "' AND " & _
                     "(TM38 = '" & textOldCH & "' OR " & _
                      "TM39 = '" & textOldEN & "' OR " & _
                      "TM40 = '" & textOldJP & "' OR " & _
                      "TM41 = '" & textOldCH & "' OR " & _
                      "TM42 = '" & textOldEN & "' OR " & _
                      "TM43 = '" & textOldJP & "' OR " & _
                      "TM55 = '" & textOldEN & "') "
      cnnConnection.Execute strSql, nRecords
   End If
   
   m_RecordAffect = m_RecordAffect + nRecords
End Sub

' 更新服務業務基本檔
Private Sub OnUpdateServicePractice()
   Dim strSql As String
   Dim nRecords As Integer
   
   nRecords = 0

   ' 客戶資料
   If m_CUFASel = 0 Then
      'Modify By Sindy 2011/2/24 增加SP65,SP66
      strSql = "UPDATE SERVICEPRACTICE SET SP30 = DECODE(SP30,'" & textOldEN & "','" & textNewEN & "',SP30), " & _
                                          "SP36 = DECODE(SP36,'" & textOldEN & "','" & textNewEN & "',SP36) " & _
               "WHERE SP01 IN (" & m_QuerySys & ") AND " & _
                     "(SP08 = '" & m_CUFANo & "' OR " & _
                     "SP58 = '" & m_CUFANo & "' OR " & _
                     "SP59 = '" & m_CUFANo & "' OR " & _
                     "SP65 = '" & m_CUFANo & "' OR " & _
                     "SP66 = '" & m_CUFANo & "') AND " & _
                     "(SP30 = '" & textOldEN & "' OR " & _
                      "SP36 = '" & textOldEN & "') "
      cnnConnection.Execute strSql, nRecords
   ' FC代理人資料
   Else
      strSql = "UPDATE SERVICEPRACTICE SET SP30 = DECODE(SP30,'" & textOldEN & "','" & textNewEN & "',SP30), " & _
                                          "SP36 = DECODE(SP36,'" & textOldEN & "','" & textNewEN & "',SP36) " & _
               "WHERE SP01 IN (" & m_QuerySys & ") AND " & _
                     "SP26 = '" & m_CUFANo & "' AND " & _
                     "(SP30 = '" & textOldEN & "' OR " & _
                      "SP36 = '" & textOldEN & "') "
      cnnConnection.Execute strSql, nRecords
   End If

   m_RecordAffect = m_RecordAffect + nRecords
End Sub

' 更新專利基本檔
Private Sub OnUpdatePatent()
   Dim strSql As String
   Dim nRecords As Integer
   
   nRecords = 0

   ' 客戶資料
   If m_CUFASel = 0 Then
      strSql = "UPDATE PATENT SET PA51 = DECODE(PA51,'" & textOldCH & "','" & textNewCH & "',PA51), " & _
                                 "PA52 = DECODE(PA52,'" & textOldEN & "','" & textNewEN & "',PA52), " & _
                                 "PA53 = DECODE(PA53,'" & textOldJP & "','" & textNewJP & "',PA53), " & _
                                 "PA54 = DECODE(PA54,'" & textOldCH & "','" & textNewCH & "',PA54), " & _
                                 "PA55 = DECODE(PA55,'" & textOldEN & "','" & textNewEN & "',PA55), " & _
                                 "PA56 = DECODE(PA56,'" & textOldJP & "','" & textNewJP & "',PA56), " & _
                                 "PA87 = DECODE(PA87,'" & textOldEN & "','" & textNewEN & "',PA87), " & _
                                 "PA98 = DECODE(PA98,'" & textOldCH & "','" & textNewCH & "',PA98), " & _
                                 "PA99 = DECODE(PA99,'" & textOldEN & "','" & textNewEN & "',PA99), " & _
                                 "PA102 = DECODE(PA102,'" & textOldEN & "','" & textNewEN & "',PA102), " & _
                                 "PA100 = DECODE(PA100,'" & textOldJP & "','" & textNewJP & "',PA100) " & _
      "WHERE PA01 IN (" & m_QuerySys & ") AND " & _
                     "(PA26 = '" & m_CUFANo & "' OR " & _
                     "PA27 = '" & m_CUFANo & "' OR " & _
                     "PA28 = '" & m_CUFANo & "' OR " & _
                     "PA29 = '" & m_CUFANo & "' OR " & _
                     "PA30 = '" & m_CUFANo & "') AND " & _
                     "(PA51 = '" & textOldCH & "' OR PA52 = '" & textOldEN & "' OR PA53 = '" & textOldJP & "' OR " & _
                      "PA54 = '" & textOldCH & "' OR PA55 = '" & textOldEN & "' OR PA56 = '" & textOldJP & "' OR " & _
                      "PA87 = '" & textOldEN & "' OR PA98 = '" & textOldCH & "' OR PA99 = '" & textOldEN & "' OR " & _
                      "PA102 = '" & textOldEN & "' OR PA100 = '" & textOldJP & "') "
      cnnConnection.Execute strSql, nRecords
   ' FC代理人資料
   Else
      strSql = "UPDATE PATENT SET PA51 = DECODE(PA51,'" & textOldCH & "','" & textNewCH & "',PA51), " & _
                                 "PA52 = DECODE(PA52,'" & textOldEN & "','" & textNewEN & "',PA52), " & _
                                 "PA53 = DECODE(PA53,'" & textOldJP & "','" & textNewJP & "',PA53), " & _
                                 "PA54 = DECODE(PA54,'" & textOldCH & "','" & textNewCH & "',PA54), " & _
                                 "PA55 = DECODE(PA55,'" & textOldEN & "','" & textNewEN & "',PA55), " & _
                                 "PA56 = DECODE(PA56,'" & textOldJP & "','" & textNewJP & "',PA56), " & _
                                 "PA87 = DECODE(PA87,'" & textOldEN & "','" & textNewEN & "',PA87), " & _
                                 "PA98 = DECODE(PA98,'" & textOldCH & "','" & textNewCH & "',PA98), " & _
                                 "PA99 = DECODE(PA99,'" & textOldEN & "','" & textNewEN & "',PA99), " & _
                                 "PA102 = DECODE(PA102,'" & textOldEN & "','" & textNewEN & "',PA102), " & _
                                 "PA100 = DECODE(PA100,'" & textOldJP & "','" & textNewJP & "',PA100) " & _
      "WHERE PA01 IN (" & m_QuerySys & ") AND " & _
                     "PA75 = '" & m_CUFANo & "' AND " & _
                     "(PA51 = '" & textOldCH & "' OR PA52 = '" & textOldEN & "' OR PA53 = '" & textOldJP & "' OR " & _
                      "PA54 = '" & textOldCH & "' OR PA55 = '" & textOldEN & "' OR PA56 = '" & textOldJP & "' OR " & _
                      "PA87 = '" & textOldEN & "' OR PA98 = '" & textOldCH & "' OR PA99 = '" & textOldEN & "' OR " & _
                      "PA102 = '" & textOldEN & "' OR PA100 = '" & textOldJP & "') "
      cnnConnection.Execute strSql, nRecords
   End If
   
   m_RecordAffect = m_RecordAffect + nRecords
End Sub

' 更新專利基本檔
Private Sub OnUpdateLawCase()
   Dim strSql As String
   Dim nRecords As Integer
   
   nRecords = 0

   ' 客戶資料
   If m_CUFASel = 0 Then
      'Modify By Sindy 2011/2/24 增加LC43,LC44,LC45,LC46
      strSql = "UPDATE LAWCASE SET LC18 = DECODE(LC18,'" & textOldCH & "','" & textNewCH & "',LC18), " & _
                                  "LC19 = DECODE(LC19,'" & textOldEN & "','" & textNewEN & "',LC19), " & _
                                  "LC20 = DECODE(LC20,'" & textOldJP & "','" & textNewJP & "',LC20), " & _
                                  "LC21 = DECODE(LC21,'" & textOldEN & "','" & textNewEN & "',LC21) " & _
               "WHERE LC01 IN (" & m_QuerySys & ") AND " & _
                     "(LC11 = '" & m_CUFANo & "' OR " & _
                     "LC43 = '" & m_CUFANo & "' OR " & _
                     "LC44 = '" & m_CUFANo & "' OR " & _
                     "LC45 = '" & m_CUFANo & "' OR " & _
                     "LC46 = '" & m_CUFANo & "') AND " & _
                     "(LC18 = '" & textOldCH & "' OR " & _
                      "LC19 = '" & textOldEN & "' OR " & _
                      "LC20 = '" & textOldJP & "' OR " & _
                      "LC21 = '" & textOldEN & "') "
      cnnConnection.Execute strSql, nRecords
   Else
      strSql = "UPDATE LAWCASE SET LC18 = DECODE(LC18,'" & textOldCH & "','" & textNewCH & "',LC18), " & _
                                  "LC19 = DECODE(LC19,'" & textOldEN & "','" & textNewEN & "',LC19), " & _
                                  "LC20 = DECODE(LC20,'" & textOldJP & "','" & textNewJP & "',LC20), " & _
                                  "LC21 = DECODE(LC21,'" & textOldEN & "','" & textNewEN & "',LC21) " & _
               "WHERE LC01 IN (" & m_QuerySys & ") AND " & _
                     "LC22 = '" & m_CUFANo & "' AND " & _
                     "(LC18 = '" & textOldCH & "' OR " & _
                      "LC19 = '" & textOldEN & "' OR " & _
                      "LC20 = '" & textOldJP & "' OR " & _
                      "LC21 = '" & textOldEN & "') "
      cnnConnection.Execute strSql, nRecords
   End If
   
   m_RecordAffect = m_RecordAffect + nRecords
End Sub

' 檢查輸入的資料是否完整
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   
   ' 客戶/代理人編號
   If IsEmptyText(textCUFA) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入客戶/代理人編號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCUFA.SetFocus
      GoTo EXITSUB
   End If
   
   ' 原聯絡人(中)(英)(日)不可同時空白
   If IsEmptyText(textOldCH) = True And IsEmptyText(textOldEN) = True And IsEmptyText(textOldJP) = True Then
      strTit = "檢核資料"
      strMsg = "原聯絡人(中)(英)(日)不可同時空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textOldCH.SetFocus
      GoTo EXITSUB
   End If
   
   ' 新聯絡人(中)(英)(日)不可同時空白
   If IsEmptyText(textNewCH) = True And IsEmptyText(textNewEN) = True And IsEmptyText(textNewJP) = True Then
      strTit = "檢核資料"
      strMsg = "新聯絡人(中)(英)(日)不可同時空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textOldCH.SetFocus
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm050713 = Nothing
End Sub

Private Sub textCUFA_2_Validate(Cancel As Boolean)
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strCUFA As String
   Dim strCUFA_2 As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCUFA_2 = textCUFA_2 & String(1 - Len(textCUFA_2), "0")
   
   textCU04 = Empty
   textCU05 = Empty
   textCU06 = Empty
   
   If IsEmptyText(textCUFA_2) = False And IsEmptyText(textCUFA) = False Then
      strCUFA = textCUFA & String(8 - Len(textCUFA), "0")
      strCUFA_2 = textCUFA_2 & String(1 - Len(textCUFA_2), "0")
      If optCUFA(0).Value = True Then
         If Mid(textCUFA, 1, 1) = "X" Then
            strSql = "SELECT * FROM CUSTOMER " & _
                     "WHERE CU01 = '" & strCUFA & "' AND " & _
                           "CU02 = '" & strCUFA_2 & "' "
            Set rsTmp = New ADODB.Recordset
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               rsTmp.MoveFirst
               If IsNull(rsTmp.Fields("CU04")) = False Then
                  textCU04 = rsTmp.Fields("CU04")
               End If
               If IsNull(rsTmp.Fields("CU05")) = False Then
                  textCU05 = rsTmp.Fields("CU05")
               End If
               If IsEmptyText(textCU05) = True And IsNull(rsTmp.Fields("CU88")) = False Then
                  textCU05 = rsTmp.Fields("CU88")
               End If
               If IsEmptyText(textCU05) = True And IsNull(rsTmp.Fields("CU89")) = False Then
                  textCU05 = rsTmp.Fields("CU89")
               End If
               If IsEmptyText(textCU05) = True And IsNull(rsTmp.Fields("CU90")) = False Then
                  textCU05 = rsTmp.Fields("CU90")
               End If
               If IsNull(rsTmp.Fields("CU06")) = False Then
                  textCU06 = rsTmp.Fields("CU06")
               End If
            Else
               rsTmp.Close
               Set rsTmp = Nothing
               Cancel = True
               strTit = "檢核資料"
               strMsg = "客戶/代理人編號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCUFA_2_GotFocus
               GoTo EXITSUB
            End If
            rsTmp.Close
            Set rsTmp = Nothing
         Else
            Cancel = True
            strTit = "檢核資料"
            strMsg = "客戶/代理人編號不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCUFA_2_GotFocus
            GoTo EXITSUB
         End If
      Else
         If Mid(textCUFA, 1, 1) = "Y" Then
            strSql = "SELECT * FROM FAGENT " & _
                     "WHERE FA01 = '" & strCUFA & "' AND " & _
                           "FA02 = '" & strCUFA_2 & "' "
            Set rsTmp = New ADODB.Recordset
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               rsTmp.MoveFirst
               If IsNull(rsTmp.Fields("FA04")) = False Then
                  textCU04 = rsTmp.Fields("FA04")
               End If
               If IsNull(rsTmp.Fields("FA05")) = False Then
                  textCU05 = rsTmp.Fields("FA05")
               End If
               If IsEmptyText(textCU05) = True And IsNull(rsTmp.Fields("FA63")) = False Then
                  textCU05 = rsTmp.Fields("FA63")
               End If
               If IsEmptyText(textCU05) = True And IsNull(rsTmp.Fields("FA64")) = False Then
                  textCU05 = rsTmp.Fields("FA64")
               End If
               If IsEmptyText(textCU05) = True And IsNull(rsTmp.Fields("FA65")) = False Then
                  textCU05 = rsTmp.Fields("FA65")
               End If
               If IsNull(rsTmp.Fields("FA06")) = False Then
                  textCU06 = rsTmp.Fields("FA06")
               End If
            Else
               rsTmp.Close
               Set rsTmp = Nothing
               Cancel = True
               strTit = "檢核資料"
               strMsg = "客戶/代理人編號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCUFA_2_GotFocus
               GoTo EXITSUB
            End If
            rsTmp.Close
            Set rsTmp = Nothing
         Else
            Cancel = True
            strTit = "檢核資料"
            strMsg = "客戶/代理人編號不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCUFA_2_GotFocus
            GoTo EXITSUB
         End If
      End If
   End If
EXITSUB:
End Sub

Private Sub textCUFA_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 客戶/代理人編號
Private Sub textCUFA_Validate(Cancel As Boolean)
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strCUFA As String
   Dim strCUFA_2 As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCU04 = Empty
   textCU05 = Empty
   textCU06 = Empty
   If IsEmptyText(textCUFA) = False Then
      strCUFA = textCUFA & String(8 - Len(textCUFA), "0")
      strCUFA_2 = textCUFA_2 & String(1 - Len(textCUFA_2), "0")
      If optCUFA(0).Value = True Then
         If Mid(textCUFA, 1, 1) = "X" Then
            strSql = "SELECT * FROM CUSTOMER " & _
                     "WHERE CU01 = '" & strCUFA & "' AND " & _
                           "CU02 = '" & strCUFA_2 & "' "
            Set rsTmp = New ADODB.Recordset
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               rsTmp.MoveFirst
               If IsNull(rsTmp.Fields("CU04")) = False Then
                  textCU04 = rsTmp.Fields("CU04")
               End If
               If IsNull(rsTmp.Fields("CU05")) = False Then
                  textCU05 = rsTmp.Fields("CU05")
               End If
               If IsEmptyText(textCU05) = True And IsNull(rsTmp.Fields("CU88")) = False Then
                  textCU05 = rsTmp.Fields("CU88")
               End If
               If IsEmptyText(textCU05) = True And IsNull(rsTmp.Fields("CU89")) = False Then
                  textCU05 = rsTmp.Fields("CU89")
               End If
               If IsEmptyText(textCU05) = True And IsNull(rsTmp.Fields("CU90")) = False Then
                  textCU05 = rsTmp.Fields("CU90")
               End If
               If IsNull(rsTmp.Fields("CU06")) = False Then
                  textCU06 = rsTmp.Fields("CU06")
               End If
            Else
               rsTmp.Close
               Set rsTmp = Nothing
               Cancel = True
               strTit = "檢核資料"
               strMsg = "客戶/代理人編號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCUFA_GotFocus
               GoTo EXITSUB
            End If
            rsTmp.Close
            Set rsTmp = Nothing
         Else
            Cancel = True
            strTit = "檢核資料"
            strMsg = "客戶/代理人編號不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCUFA_GotFocus
            GoTo EXITSUB
         End If
      Else
         If Mid(textCUFA, 1, 1) = "Y" Then
            strSql = "SELECT * FROM FAGENT " & _
                     "WHERE FA01 = '" & strCUFA & "' AND " & _
                           "FA02 = '" & strCUFA_2 & "' "
            Set rsTmp = New ADODB.Recordset
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               rsTmp.MoveFirst
               If IsNull(rsTmp.Fields("FA04")) = False Then
                  textCU04 = rsTmp.Fields("FA04")
               End If
               If IsNull(rsTmp.Fields("FA05")) = False Then
                  textCU05 = rsTmp.Fields("FA05")
               End If
               If IsEmptyText(textCU05) = True And IsNull(rsTmp.Fields("FA63")) = False Then
                  textCU05 = rsTmp.Fields("FA63")
               End If
               If IsEmptyText(textCU05) = True And IsNull(rsTmp.Fields("FA64")) = False Then
                  textCU05 = rsTmp.Fields("FA64")
               End If
               If IsEmptyText(textCU05) = True And IsNull(rsTmp.Fields("FA65")) = False Then
                  textCU05 = rsTmp.Fields("FA65")
               End If
               If IsNull(rsTmp.Fields("FA06")) = False Then
                  textCU06 = rsTmp.Fields("FA06")
               End If
            Else
               rsTmp.Close
               Set rsTmp = Nothing
               Cancel = True
               strTit = "檢核資料"
               strMsg = "客戶/代理人編號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCUFA_GotFocus
               GoTo EXITSUB
            End If
            rsTmp.Close
            Set rsTmp = Nothing
         Else
            Cancel = True
            strTit = "檢核資料"
            strMsg = "客戶/代理人編號不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCUFA_GotFocus
            GoTo EXITSUB
         End If
      End If
   End If
EXITSUB:
End Sub

' 原聯絡人(中)
Private Sub textOldCH_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Modified by Lydia 2017/06/14 聯絡人(中)改為30字
   'If CheckLengthIsOK(textOldCH, 10) = False Then
   If CheckLengthIsOK(textOldCH, 30) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "原聯絡人(中)內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textOldCH_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textOldCH.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 原聯絡人(英)
Private Sub textOldEN_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Modified by Lydia 2017/06/14 聯絡人(英)改為35字
   'If CheckLengthIsOK(textOldEN, textOldEN.MaxLength) = False Then
   If CheckLengthIsOK(textOldEN, 35) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "原聯絡人(英)內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textOldEN_GotFocus
   End If
End Sub

' 原聯絡人(日)
Private Sub textOldJP_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Modified by Lydia 2017/06/14 聯絡人(日)改為60字
   'If CheckLengthIsOK(textOldJP, 20) = False Then
   If CheckLengthIsOK(textOldJP, 60) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "原聯絡人(日)內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textOldJP_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textOldJP.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 新聯絡人(中)
Private Sub textNewCH_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Modified by Lydia 2017/06/14 聯絡人(中)改為30字
   'If CheckLengthIsOK(textNewCH, 10) = False Then
   If CheckLengthIsOK(textNewCH, 30) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "新聯絡人(中)內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textNewCH_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textNewCH.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 新聯絡人(英)
Private Sub textNewEN_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Modified by Lydia 2017/06/14 聯絡人(英)改為35字
   'If CheckLengthIsOK(textNewEN, textNewEN.MaxLength) = False Then
   If CheckLengthIsOK(textNewEN, 35) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "新聯絡人(英)內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textNewEN_GotFocus
   End If
End Sub

' 新聯絡人(日)
Private Sub textNewJP_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Modified by Lydia 2017/06/14 聯絡人(日)改為30字
   'If CheckLengthIsOK(textNewJP, 20) = False Then
   If CheckLengthIsOK(textNewJP, 60) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "新聯絡人(日)內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textNewJP_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textNewJP.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textCUFA_GotFocus()
   InverseTextBox textCUFA
End Sub

Private Sub textCUFA_2_GotFocus()
   InverseTextBox textCUFA_2
End Sub

Private Sub textOldCH_GotFocus()
   InverseTextBox textOldCH
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textOldCH.IMEMode = 1
   OpenIme
End Sub

Private Sub textOldEN_GotFocus()
   InverseTextBox textOldEN
End Sub

Private Sub textOldJP_GotFocus()
   InverseTextBox textOldJP
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textOldJP.IMEMode = 1
   OpenIme
End Sub

Private Sub textNewCH_GotFocus()
   InverseTextBox textNewCH
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textNewCH.IMEMode = 1
   OpenIme
End Sub

Private Sub textNewEN_GotFocus()
   InverseTextBox textNewEN
End Sub

Private Sub textNewJP_GotFocus()
   InverseTextBox textNewJP
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textNewJP.IMEMode = 1
   OpenIme
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textCUFA.Enabled = True Then
   Cancel = False
   textCUFA_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCUFA_2.Enabled = True Then
   Cancel = False
   textCUFA_2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textNewCH.Enabled = True Then
   Cancel = False
   textNewCH_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textNewEN.Enabled = True Then
   Cancel = False
   textNewEN_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textNewJP.Enabled = True Then
   Cancel = False
   textNewJP_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textOldCH.Enabled = True Then
   Cancel = False
   textOldCH_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textOldEN.Enabled = True Then
   Cancel = False
   textOldEN_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textOldJP.Enabled = True Then
   Cancel = False
   textOldJP_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Added by Lydia 2021/09/23 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
    Exit Function
End If

TxtValidate = True
End Function

