VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm880010 
   BorderStyle     =   1  '單線固定
   Caption         =   "資料刪除記錄"
   ClientHeight    =   4785
   ClientLeft      =   1185
   ClientTop       =   945
   ClientWidth     =   7665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7665
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6384
      TabIndex        =   3
      Top             =   50
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5448
      TabIndex        =   2
      Top             =   50
      Width           =   912
   End
   Begin VB.TextBox textDD23 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1380
      MaxLength       =   6
      TabIndex        =   0
      Top             =   3360
      Width           =   1212
   End
   Begin VB.TextBox textDD07 
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Left            =   1380
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2760
      Width           =   6132
   End
   Begin VB.TextBox textHit 
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Left            =   180
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   540
      Width           =   7332
   End
   Begin VB.TextBox textDD14 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1380
      MaxLength       =   9
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1212
   End
   Begin VB.TextBox textDD01 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1380
      MaxLength       =   3
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   960
      Width           =   492
   End
   Begin VB.TextBox textDD02 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1860
      MaxLength       =   6
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   960
      Width           =   732
   End
   Begin VB.TextBox textDD03 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2580
      MaxLength       =   1
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   252
   End
   Begin VB.TextBox textDD04 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2820
      MaxLength       =   2
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   372
   End
   Begin MSForms.TextBox textDD15 
      Height          =   300
      Left            =   5100
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2412
      VariousPropertyBits=   671105055
      Size            =   "4254;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDD23_2 
      Height          =   300
      Left            =   2700
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3360
      Width           =   4812
      VariousPropertyBits=   671105055
      Size            =   "8488;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDD24 
      Height          =   930
      Left            =   1380
      TabIndex        =   1
      Top             =   3720
      Width           =   6135
      VariousPropertyBits=   -1467989989
      BackColor       =   16777215
      MaxLength       =   60
      ScrollBars      =   2
      Size            =   "10821;1640"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDD12 
      Height          =   300
      Left            =   1380
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2400
      Width           =   6132
      VariousPropertyBits=   671105055
      Size            =   "10816;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDD06 
      Height          =   300
      Left            =   1380
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2040
      Width           =   6132
      VariousPropertyBits=   671105055
      Size            =   "10816;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDD05 
      Height          =   300
      Left            =   1380
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1680
      Width           =   6132
      VariousPropertyBits=   671105055
      Size            =   "10816;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   7560
      Y1              =   3180
      Y2              =   3180
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   7560
      Y1              =   3204
      Y2              =   3204
   End
   Begin VB.Label Label9 
      Caption         =   "案件性質："
      Height          =   252
      Left            =   3900
      TabIndex        =   23
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label Label8 
      Caption         =   "刪除備註："
      Height          =   252
      Left            =   180
      TabIndex        =   21
      Top             =   3720
      Width           =   972
   End
   Begin VB.Label Label7 
      Caption         =   "失誤人員："
      Height          =   252
      Left            =   180
      TabIndex        =   20
      Top             =   3360
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "申請國家："
      Height          =   252
      Left            =   180
      TabIndex        =   18
      Top             =   2760
      Width           =   972
   End
   Begin VB.Label Label5 
      Caption         =   "FC代理人："
      Height          =   252
      Left            =   180
      TabIndex        =   16
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   "申請人："
      Height          =   252
      Left            =   180
      TabIndex        =   14
      Top             =   2040
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱："
      Height          =   252
      Left            =   180
      TabIndex        =   12
      Top             =   1680
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "收文號："
      Height          =   252
      Left            =   180
      TabIndex        =   9
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號："
      Height          =   252
      Left            =   180
      TabIndex        =   8
      Top             =   960
      Width           =   972
   End
End
Attribute VB_Name = "frm880010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/16 改成Form2.0 ; textDD15、textDD05、textDD06、textDD12、textDD23_2、textDD24
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan2010/8/19 日期欄已修改
Option Explicit

Const MAX_FIELD = 28

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList(MAX_FIELD) As FIELDITEM

' 本所案號
Dim m_DD01 As String
Dim m_DD02 As String
Dim m_DD03 As String
Dim m_DD04 As String
' 收文號
Dim m_DD14 As String
' 下一程序檔的序號
Dim m_NP22 As String
' 資料來源種類
Dim m_DataSource As Integer
' 國家代號
Dim m_Nation As String
' 確定或取消
Dim m_OKCancel As Integer

Private Sub cmdCancel_Click()
   Me.Hide
End Sub

Private Sub cmdOK_Click()
   If CheckDataValid() = True Then
      cmdOK.Enabled = False 'Add By Sindy 2016/6/2 防止操作人員誤按二下
      UpdateFieldNewData
      OnSaveDataDeleteRecord
      'If m_DataSource = 1 Then
      '   If ExistProgress(m_DD01, m_DD02, m_DD03, m_DD04) = False Then
      '      DeleteMainTable m_DD01, m_DD02, m_DD03, m_DD04
      '   End If
      'End If
      m_OKCancel = 1
      Me.Hide
      Me.cmdOK.Enabled = True 'Add By Sindy 2016/6/2 恢復
   End If
End Sub

Private Sub Form_Load()
   m_OKCancel = 0
   textHit.BackColor = &H8000000F
   textDD01.BackColor = &H8000000F
   textDD02.BackColor = &H8000000F
   textDD03.BackColor = &H8000000F
   textDD04.BackColor = &H8000000F
   textDD05.BackColor = &H8000000F
   textDD06.BackColor = &H8000000F
   textDD07.BackColor = &H8000000F
   textDD12.BackColor = &H8000000F
   textDD14.BackColor = &H8000000F
   textDD15.BackColor = &H8000000F
   textDD23_2.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   InitialField
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ClearFieldList
   'Add By Cheng 2002/07/18
   'Set frm880010 = Nothing
End Sub

' 傳回使用者按下OK還是Cancel
Public Function IsOK() As Boolean
   IsOK = False
   If m_OKCancel = 1 Then
      IsOK = True
   End If
End Function

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_DD01 = Empty
      m_DD02 = Empty
      m_DD03 = Empty
      m_DD04 = Empty
      m_DD14 = Empty
      m_NP22 = Empty
      m_DataSource = 0
   End If

   Select Case nType
      ' 本所案號
      Case 0: m_DD01 = strData
      Case 1: m_DD02 = strData
      Case 2: m_DD03 = strData
      Case 3: m_DD04 = strData
      Case 4:
         m_DD14 = strData
         m_DataSource = 1
      Case 5:
         m_NP22 = strData
         m_DataSource = 2
      Case 6:
         textDD23 = strData
      Case 7:
         textDD24 = strData
   End Select
End Sub

Public Function GetData(ByVal nType As Integer) As String
   Select Case nType
      Case 0: GetData = textDD23
      Case 1: GetData = textDD24
   End Select
End Function

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "DD" & strTmp
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
      Select Case nIndex
         Case 16, 17, 18, 20, 21, 25, 27, 28:
            m_FieldList(nIndex - 1).fiType = 1
      End Select
   Next nIndex
End Sub

Private Sub ClearFieldList()
   Erase m_FieldList
End Sub

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To MAX_FIELD - 1
      If strName = m_FieldList(nIndex).fiName Then
         m_FieldList(nIndex).fiNewData = strData
         Exit For
      End If
   Next nIndex
End Sub

' 更新欄位的內容
Private Sub UpdateFieldNewData()
   SetFieldNewData "DD23", textDD23
   SetFieldNewData "DD24", textDD24
   SetFieldNewData "DD28", CStr(GetMaxNumber())
End Sub

Private Function GetMaxNumber() As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   strSql = "SELECT MAX(DD28) FROM DATADELETERECORD "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields(0)) = False Then
         GetMaxNumber = rsTmp.Fields(0)
         GetMaxNumber = Val(GetMaxNumber) + 1
      End If
   End If
   If IsEmptyText(GetMaxNumber) = True Then
      GetMaxNumber = 1
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 取得國家代碼
Private Function GetNationNo(ByVal strKey1 As String, ByVal StrKey2 As String, ByVal strKey3 As String, ByVal strKey4 As String) As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   GetNationNo = "000"
   
   Select Case m_DD01
      ' 讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         strSql = "SELECT TM10 FROM TRADEMARK WHERE TM01 = '" & strKey1 & "' AND TM02 = '" & StrKey2 & "' AND TM03 = '" & strKey3 & "' AND TM04 = '" & strKey4 & "' "
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         strSql = "SELECT PA09 FROM PATENT WHERE PA01 = '" & strKey1 & "' AND PA02 = '" & StrKey2 & "' AND PA03 = '" & strKey3 & "' AND PA04 = '" & strKey4 & "' "
      ' 讀取法務基本檔
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS":
         strSql = "SELECT LC15 FROM LAWCASE WHERE LC01 = '" & strKey1 & "' AND LC02 = '" & StrKey2 & "' AND LC03 = '" & strKey3 & "' AND LC04 = '" & strKey4 & "' "
      ' 讀取顧問案件基本檔
      Case "LA":
         strSql = Empty
      ' 讀取服務業務基本檔
      Case Else:
         strSql = "SELECT SP09 FROM SERVICEPRACTICE WHERE SP01 = '" & strKey1 & "' AND SP02 = '" & StrKey2 & "' AND SP03 = '" & strKey3 & "' AND SP04 = '" & strKey4 & "' "
   End Select
   If IsEmptyText(strSql) = False Then
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then
            If IsEmptyText(rsTmp.Fields(0)) = False Then
               GetNationNo = rsTmp.Fields(0)
            End If
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

' 讀取商標基本檔
Private Function QueryTradeMark(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryTradeMark = False
   strSql = "SELECT * FROM TRADEMARK " & _
            "WHERE TM01 = '" & strTM01 & "' AND " & _
                  "TM02 = '" & strTM02 & "' AND " & _
                  "TM03 = '" & strTM03 & "' AND " & _
                  "TM04 = '" & strTM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryTradeMark = True
      rsTmp.MoveFirst
      ' 案件名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then
         textDD05 = rsTmp.Fields("TM05")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textDD06 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_Nation = rsTmp.Fields("TM10")
         textDD07 = GetNationName(rsTmp.Fields("TM10"), 0)
      End If
      If m_DataSource = 0 Then
         ' 案件名稱
         If IsNull(rsTmp.Fields("TM05")) = False Then
            SetFieldNewData "DD05", rsTmp.Fields("TM05")
         End If
         ' 申請人
         If IsNull(rsTmp.Fields("TM23")) = False Then
            SetFieldNewData "DD06", rsTmp.Fields("TM23")
         End If
         ' 申請國家
         If IsNull(rsTmp.Fields("TM10")) = False Then
            SetFieldNewData "DD07", rsTmp.Fields("TM10")
         End If
         ' 申請案號
         If IsNull(rsTmp.Fields("TM12")) = False Then
            SetFieldNewData "DD08", rsTmp.Fields("TM12")
         End If
         ' 分所案號
         If IsNull(rsTmp.Fields("TM34")) = False Then
            SetFieldNewData "DD09", rsTmp.Fields("TM34")
         End If
         ' 商標種類
         If IsNull(rsTmp.Fields("TM08")) = False Then
            SetFieldNewData "DD10", rsTmp.Fields("TM08")
         End If
         ' 目前准駁
         If IsNull(rsTmp.Fields("TM16")) = False Then
            SetFieldNewData "DD11", rsTmp.Fields("TM16")
         End If
         ' FC代理人
         If IsNull(rsTmp.Fields("TM44")) = False Then
            SetFieldNewData "DD12", rsTmp.Fields("TM44")
         End If
         ' 延展通知人
         If IsNull(rsTmp.Fields("TM33")) = False Then
            SetFieldNewData "DD13", rsTmp.Fields("TM33")
         End If
         ' 原資料產生日期
         If IsNull(rsTmp.Fields("TM60")) = False Then
            SetFieldNewData "DD25", rsTmp.Fields("TM60")
         End If
         ' 原資料產生人員
         If IsNull(rsTmp.Fields("TM59")) = False Then
            SetFieldNewData "DD26", rsTmp.Fields("TM59")
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取服務業務基本檔
Private Function QueryServicePractice(ByVal strSP01 As String, ByVal strSP02 As String, ByVal strSP03 As String, ByVal strSP04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryServicePractice = False
   strSql = "SELECT * FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & strSP01 & "' AND " & _
                  "SP02 = '" & strSP02 & "' AND " & _
                  "SP03 = '" & strSP03 & "' AND " & _
                  "SP04 = '" & strSP04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryServicePractice = True
      rsTmp.MoveFirst
      ' 案件名稱
      If IsNull(rsTmp.Fields("SP05")) = False Then
         textDD05 = rsTmp.Fields("SP05")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textDD06 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_Nation = rsTmp.Fields("SP09")
         textDD07 = GetNationName(rsTmp.Fields("SP09"), 0)
      End If
      If m_DataSource = 0 Then
         ' 案件名稱
         If IsNull(rsTmp.Fields("SP05")) = False Then
            SetFieldNewData "DD05", rsTmp.Fields("SP05")
         End If
         ' 申請人
         If IsNull(rsTmp.Fields("SP08")) = False Then
            SetFieldNewData "DD06", rsTmp.Fields("SP08")
         End If
         ' 申請國家
         If IsNull(rsTmp.Fields("SP09")) = False Then
            SetFieldNewData "DD07", rsTmp.Fields("SP09")
         End If
         ' 申請案號
         If IsNull(rsTmp.Fields("SP11")) = False Then
            SetFieldNewData "DD08", rsTmp.Fields("SP11")
         End If
         ' 分所案號
         If IsNull(rsTmp.Fields("SP28")) = False Then
            SetFieldNewData "DD09", rsTmp.Fields("SP28")
         End If
         ' FC代理人
         If IsNull(rsTmp.Fields("SP26")) = False Then
            SetFieldNewData "DD12", rsTmp.Fields("SP26")
         End If
         ' 原資料產生日期
         If IsNull(rsTmp.Fields("SP53")) = False Then
            SetFieldNewData "DD25", rsTmp.Fields("SP53")
         End If
         ' 原資料產生人員
         If IsNull(rsTmp.Fields("SP52")) = False Then
            SetFieldNewData "DD26", rsTmp.Fields("SP52")
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取專利基本檔
Private Function QueryPatent(ByVal strPA01 As String, ByVal strPA02 As String, ByVal strPA03 As String, ByVal strPA04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryPatent = False
   strSql = "SELECT * FROM PATENT " & _
            "WHERE PA01 = '" & strPA01 & "' AND " & _
                  "PA02 = '" & strPA02 & "' AND " & _
                  "PA03 = '" & strPA03 & "' AND " & _
                  "PA04 = '" & strPA04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryPatent = True
      rsTmp.MoveFirst
      ' 案件名稱
      If IsNull(rsTmp.Fields("PA05")) = False Then
         textDD05 = rsTmp.Fields("PA05")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("PA26")) = False Then
         textDD06 = GetCustomerName(rsTmp.Fields("PA26"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("PA09")) = False Then
         m_Nation = rsTmp.Fields("PA09")
         textDD07 = GetNationName(rsTmp.Fields("PA09"), 0)
      End If
      If m_DataSource = 0 Then
         ' 案件名稱
         If IsNull(rsTmp.Fields("PA05")) = False Then
            SetFieldNewData "DD05", rsTmp.Fields("PA05")
         End If
         ' 申請人
         If IsNull(rsTmp.Fields("PA26")) = False Then
            SetFieldNewData "DD06", rsTmp.Fields("PA26")
         End If
         ' 申請國家
         If IsNull(rsTmp.Fields("PA09")) = False Then
            SetFieldNewData "DD07", rsTmp.Fields("PA09")
         End If
         ' 申請案號
         If IsNull(rsTmp.Fields("PA11")) = False Then
            SetFieldNewData "DD08", rsTmp.Fields("PA11")
         End If
         ' 分所案號
         If IsNull(rsTmp.Fields("PA47")) = False Then
            SetFieldNewData "DD09", rsTmp.Fields("PA47")
         End If
         ' 商標種類
         If IsNull(rsTmp.Fields("PA08")) = False Then
            SetFieldNewData "DD10", rsTmp.Fields("PA08")
         End If
         ' 目前准駁
         If IsNull(rsTmp.Fields("PA16")) = False Then
            SetFieldNewData "DD11", rsTmp.Fields("PA16")
         End If
         ' FC代理人
         If IsNull(rsTmp.Fields("PA75")) = False Then
            SetFieldNewData "DD12", rsTmp.Fields("PA75")
         End If
         ' 年費通知人
         If IsNull(rsTmp.Fields("PA76")) = False Then
            SetFieldNewData "DD13", rsTmp.Fields("PA76")
         End If
         ' 原資料產生日期
         If IsNull(rsTmp.Fields("PA93")) = False Then
            SetFieldNewData "DD25", rsTmp.Fields("PA93")
         End If
         ' 原資料產生人員
         If IsNull(rsTmp.Fields("PA92")) = False Then
            SetFieldNewData "DD26", rsTmp.Fields("PA92")
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取法務基本檔
Private Function QueryLawCase(ByVal strLC01 As String, ByVal strLC02 As String, ByVal strLC03 As String, ByVal strLC04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryLawCase = False
   strSql = "SELECT * FROM LAWCASE " & _
            "WHERE LC01 = '" & strLC01 & "' AND " & _
                  "LC02 = '" & strLC02 & "' AND " & _
                  "LC03 = '" & strLC03 & "' AND " & _
                  "LC04 = '" & strLC04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryLawCase = True
      rsTmp.MoveFirst
      ' 案件名稱
      If IsNull(rsTmp.Fields("LC05")) = False Then
         textDD05 = rsTmp.Fields("LC05")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("LC11")) = False Then
         textDD06 = GetCustomerName(rsTmp.Fields("LC11"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("LC15")) = False Then
         m_Nation = rsTmp.Fields("LC15")
         textDD07 = GetNationName(rsTmp.Fields("LC15"), 0)
      End If
      If m_DataSource = 0 Then
         ' 案件名稱
         If IsNull(rsTmp.Fields("LC05")) = False Then
            SetFieldNewData "DD05", rsTmp.Fields("LC05")
         End If
         ' 申請人
         If IsNull(rsTmp.Fields("LC11")) = False Then
            SetFieldNewData "DD06", rsTmp.Fields("LC11")
         End If
         ' 申請國家
         If IsNull(rsTmp.Fields("LC15")) = False Then
            SetFieldNewData "DD07", rsTmp.Fields("LC15")
         End If
         ' 分所案號
         If IsNull(rsTmp.Fields("LC16")) = False Then
            SetFieldNewData "DD09", rsTmp.Fields("LC16")
         End If
         ' FC代理人
         If IsNull(rsTmp.Fields("LC22")) = False Then
            SetFieldNewData "DD12", rsTmp.Fields("LC22")
         End If
         ' 原資料產生日期
         If IsNull(rsTmp.Fields("LC29")) = False Then
            SetFieldNewData "DD25", rsTmp.Fields("LC29")
         End If
         ' 原資料產生人員
         If IsNull(rsTmp.Fields("LC28")) = False Then
            SetFieldNewData "DD26", rsTmp.Fields("LC28")
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取顧問案件基本檔
Private Function QueryHireCase(ByVal strHC01 As String, ByVal strHC02 As String, ByVal strHC03 As String, ByVal strHC04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryHireCase = False
   strSql = "SELECT * FROM HIRECASE " & _
            "WHERE HC01 = '" & strHC01 & "' AND " & _
                  "HC02 = '" & strHC02 & "' AND " & _
                  "HC03 = '" & strHC03 & "' AND " & _
                  "HC04 = '" & strHC04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryHireCase = True
      rsTmp.MoveFirst
      ' 案件名稱
      If IsNull(rsTmp.Fields("HC06")) = False Then
         textDD05 = rsTmp.Fields("HC06")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("HC05")) = False Then
         textDD06 = GetCustomerName(rsTmp.Fields("HC05"), 0)
      End If
      If m_DataSource = 0 Then
         ' 案件名稱
         If IsNull(rsTmp.Fields("HC06")) = False Then
            SetFieldNewData "DD05", rsTmp.Fields("HC06")
         End If
         ' 申請人
         If IsNull(rsTmp.Fields("HC05")) = False Then
            SetFieldNewData "DD06", rsTmp.Fields("HC05")
         End If
         ' 申請國家
         SetFieldNewData "DD07", "000"
         ' 分所案號
         If IsNull(rsTmp.Fields("HC07")) = False Then
            SetFieldNewData "DD09", rsTmp.Fields("HC07")
         End If
         ' 原資料產生日期
         If IsNull(rsTmp.Fields("HC14")) = False Then
            SetFieldNewData "DD25", rsTmp.Fields("HC14")
         End If
         ' 原資料產生人員
         If IsNull(rsTmp.Fields("HC13")) = False Then
            SetFieldNewData "DD26", rsTmp.Fields("HC13")
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取案件進度檔
Private Function QueryCaseProgress(ByVal strCP09 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryCaseProgress = False
   strSql = "SELECT * FROM CASEPROGRESS " & _
            "WHERE CP09 = '" & strCP09 & "' "
   ' 若有輸入本所案號則需設定本所案號以加快搜尋的速度
   If IsEmptyText(m_DD01) = False Then
      strSql = strSql & " AND " & _
               "CP01 = '" & m_DD01 & "' AND " & _
               "CP02 = '" & m_DD02 & "' AND " & _
               "CP03 = '" & m_DD03 & "' AND " & _
               "CP04 = '" & m_DD04 & "' "
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryCaseProgress = True
      rsTmp.MoveFirst
      ' 若資料來源不為從基本檔且本所案號為空的時則需設定本所案號
      If m_DataSource <> 0 And IsEmptyText(m_DD01) = True Then
         m_DD01 = rsTmp.Fields("CP01")
         m_DD02 = rsTmp.Fields("CP02")
         m_DD03 = rsTmp.Fields("CP03")
         m_DD04 = rsTmp.Fields("CP04")
      End If
      ' 取得國家代碼
      m_Nation = GetNationNo(m_DD01, m_DD02, m_DD03, m_DD04)
      ' 收文號
      If IsNull(rsTmp.Fields("CP09")) = False Then
         SetFieldNewData "DD14", rsTmp.Fields("CP09")
      End If
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         SetFieldNewData "DD15", rsTmp.Fields("CP10")
         If m_Nation < "010" Then
            textDD15 = GetCaseTypeName(m_DD01, rsTmp.Fields("CP10"), 0)
         Else
            textDD15 = GetCaseTypeName(m_DD01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 本所期限
      If IsNull(rsTmp.Fields("CP06")) = False Then
         SetFieldNewData "DD16", rsTmp.Fields("CP06")
      End If
      ' 法定期限
      If IsNull(rsTmp.Fields("CP07")) = False Then
         SetFieldNewData "DD17", rsTmp.Fields("CP07")
      End If
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then
         SetFieldNewData "DD18", rsTmp.Fields("CP05")
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         SetFieldNewData "DD19", rsTmp.Fields("CP13")
      End If
      ' 費用
      If IsNull(rsTmp.Fields("CP16")) = False Then
         SetFieldNewData "DD20", rsTmp.Fields("CP16")
      End If
      ' 規費
      If IsNull(rsTmp.Fields("CP17")) = False Then
         SetFieldNewData "DD21", rsTmp.Fields("CP17")
      End If
      ' 收據編號/請款編號
      If IsNull(rsTmp.Fields("CP60")) = False Then
         SetFieldNewData "DD22", rsTmp.Fields("CP60")
      End If
      If m_DataSource = 1 Then
         ' 原資料產生日期
         If IsNull(rsTmp.Fields("CP66")) = False Then
            SetFieldNewData "DD25", rsTmp.Fields("CP66")
         End If
         ' 原資料產生人員
         If IsNull(rsTmp.Fields("CP65")) = False Then
            SetFieldNewData "DD26", rsTmp.Fields("CP65")
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取下一程序檔
Private Function QueryNextProgress(ByVal strNP22 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   QueryNextProgress = False
   strSql = "SELECT * FROM NEXTPROGRESS " & _
            "WHERE NP22 = " & strNP22 & " "
   ' 若有輸入本所案號則需設定本所案號以加快搜尋的速度
   If IsEmptyText(m_DD01) = False Then
      strSql = strSql & " AND " & _
               "NP02 = '" & m_DD01 & "' AND " & _
               "NP03 = '" & m_DD02 & "' AND " & _
               "NP04 = '" & m_DD03 & "' AND " & _
               "NP05 = '" & m_DD04 & "' "
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryNextProgress = True
      rsTmp.MoveFirst
      ' 若資料來源不為從基本檔且本所案號為空的時則需設定本所案號
      If m_DataSource <> 0 And IsEmptyText(m_DD01) = True Then
         m_DD14 = rsTmp.Fields("NP01")
         m_DD01 = rsTmp.Fields("NP02")
         m_DD02 = rsTmp.Fields("NP03")
         m_DD03 = rsTmp.Fields("NP04")
         m_DD04 = rsTmp.Fields("NP05")
      End If
      ' 取得國家代碼
      m_Nation = GetNationNo(m_DD01, m_DD02, m_DD03, m_DD04)
      ' 收文號
      If IsNull(rsTmp.Fields("NP01")) = False Then
         SetFieldNewData "DD14", rsTmp.Fields("NP01")
      End If
      ' 案件性質
      If IsNull(rsTmp.Fields("NP07")) = False Then
         SetFieldNewData "DD15", rsTmp.Fields("NP07")
         If m_Nation < "010" Then
            textDD15 = GetCaseTypeName(m_DD01, rsTmp.Fields("NP07"), 0)
         Else
            textDD15 = GetCaseTypeName(m_DD01, rsTmp.Fields("NP07"), 1)
         End If
      End If
      ' 本所期限
      If IsNull(rsTmp.Fields("NP08")) = False Then
         SetFieldNewData "DD16", rsTmp.Fields("NP08")
      End If
      ' 法定期限
      If IsNull(rsTmp.Fields("NP09")) = False Then
         SetFieldNewData "DD17", rsTmp.Fields("NP09")
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("NP10")) = False Then
         SetFieldNewData "DD19", rsTmp.Fields("NP10")
      End If
      If m_DataSource = 2 Then
         ' 原資料產生日期
         If IsNull(rsTmp.Fields("NP17")) = False Then
            SetFieldNewData "DD25", rsTmp.Fields("NP17")
         End If
         ' 原資料產生人員
         If IsNull(rsTmp.Fields("NP16")) = False Then
            SetFieldNewData "DD26", rsTmp.Fields("NP16")
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
EXITSUB:
End Function

Public Function QueryData() As Boolean
   Dim bQuery As Boolean
   
   QueryData = False
   
   Select Case m_DataSource
      Case 1:
         textHit = "刪除案件進度檔"
      Case 2:
         textHit = "刪除下一程序檔"
      Case Else:
         textHit = "刪除基本檔"
   End Select
   
   ' 依資料的來源方式讀取檔案
   Select Case m_DataSource
      Case 1:
         ' 讀取下一程序檔
         If QueryCaseProgress(m_DD14) = False Then
            GoTo EXITSUB
         End If
      Case 2:
         ' 讀取案件進度檔
         If QueryNextProgress(m_NP22) = False Then
            GoTo EXITSUB
         End If
   End Select
   
   ' 更新本所案號
   SetFieldNewData "DD01", m_DD01
   SetFieldNewData "DD02", m_DD02
   SetFieldNewData "DD03", m_DD03
   SetFieldNewData "DD04", m_DD04
   ' 刪除日期
   SetFieldNewData "DD27", DBDATE(SystemDate())
   
   textDD01 = m_DD01
   textDD02 = m_DD02
   textDD03 = m_DD03
   textDD04 = m_DD04
   textDD14 = m_DD14
   'Add By Cheng 2002/07/17
   m_Nation = ""
   
   Select Case m_DD01
      ' 讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         bQuery = QueryTradeMark(m_DD01, m_DD02, m_DD03, m_DD04)
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         bQuery = QueryPatent(m_DD01, m_DD02, m_DD03, m_DD04)
      ' 讀取法務基本檔
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS":
         bQuery = QueryLawCase(m_DD01, m_DD02, m_DD03, m_DD04)
      ' 讀取顧問案件基本檔
      Case "LA":
         bQuery = QueryHireCase(m_DD01, m_DD02, m_DD03, m_DD04)
      ' 讀取服務業務基本檔
      Case Else:
         bQuery = QueryServicePractice(m_DD01, m_DD02, m_DD03, m_DD04)
   End Select
   
   ' 當刪除的是基本檔時, 若基本檔不存在則傳回False
   If m_DataSource = 0 Then
      If bQuery = False Then
         GoTo EXITSUB
      End If
   End If
   
   QueryData = True
EXITSUB:
End Function

Public Sub OnSaveData()
   UpdateFieldNewData
   OnSaveDataDeleteRecord
   'If m_DataSource = 1 Then
   '   If ExistProgress(m_DD01, m_DD02, m_DD03, m_DD04) = False Then
   '      DeleteMainTable m_DD01, m_DD02, m_DD03, m_DD04
   '   End If
   'End If
   m_OKCancel = 1
   Me.Hide
End Sub

' 新增記錄
Private Sub OnSaveDataDeleteRecord()
   Dim strSql As String
   Dim strTmp As String
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   
   ' 先設定為第一個欄位
   bFirst = True
   ' 先設定為不執行
   bDifference = False
   ' 組成SQL語法
   strSql = "INSERT INTO DATADELETERECORD ("
   For nIndex = 0 To MAX_FIELD - 1
      If IsEmptyText(m_FieldList(nIndex).fiNewData) = False Then
         strTmp = m_FieldList(nIndex).fiName
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   
   bFirst = True
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If IsEmptyText(m_FieldList(nIndex).fiNewData) = False Then
         If m_FieldList(nIndex).fiType = 0 Then
            'strTmp = "'" & m_FieldList(nIndex).fiNewData & "'"
            strTmp = "'" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
         Else
            strTmp = m_FieldList(nIndex).fiNewData
         End If
      End If
      If strTmp <> Empty Then
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ")"
   
   ' 檢查是否要執行
   If bDifference = True Then
      cnnConnection.Execute strSql
   End If
End Sub

' 刪除基本檔
Private Sub DeleteMainTable(ByVal strDD01 As String, ByVal strDD02 As String, ByVal strDD03 As String, ByVal strDD04 As String)
   Dim bQuery As Boolean
   
   ' 設定刪除基本檔
   m_DataSource = 0
   ' 清除欄位串列的內容
   InitialField
   
   ' 更新本所案號
   SetFieldNewData "DD01", strDD01
   SetFieldNewData "DD02", strDD02
   SetFieldNewData "DD03", strDD03
   SetFieldNewData "DD04", strDD04
   ' 刪除日期
   SetFieldNewData "DD27", DBDATE(SystemDate())
   
   textDD01 = strDD01
   textDD02 = strDD02
   textDD03 = strDD03
   textDD04 = strDD04
   
   Select Case m_DD01
      ' 讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         bQuery = QueryTradeMark(strDD01, strDD02, strDD03, strDD04)
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         bQuery = QueryPatent(strDD01, strDD02, strDD03, strDD04)
      ' 讀取法務基本檔
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS":
         bQuery = QueryLawCase(strDD01, strDD02, strDD03, strDD04)
      ' 讀取顧問案件基本檔
      Case "LA":
         bQuery = QueryHireCase(strDD01, strDD02, strDD03, strDD04)
      ' 讀取服務業務基本檔
      Case Else:
         bQuery = QueryServicePractice(strDD01, strDD02, strDD03, strDD04)
   End Select
   
   If bQuery = True Then
      ' 更新欄位
      UpdateFieldNewData
      
      ' 寫刪除記錄檔
      OnSaveDataDeleteRecord
      
      Select Case m_DD01
         ' 讀取商標基本檔
         Case "T", "TF", "CFT", "FCT":
            strSql = "DELETE FROM TRADEMRK " & _
                     "WHERE TM01 = '" & strDD01 & "' AND " & _
                           "TM02 = '" & strDD02 & "' AND " & _
                           "TM03 = '" & strDD03 & "' AND " & _
                           "TM04 = '" & strDD04 & "' "
         ' 讀取專利基本檔
         Case "P", "CFP", "FCP":
            strSql = "DELETE FROM PATENT " & _
                     "WHERE PA01 = '" & strDD01 & "' AND " & _
                           "PA02 = '" & strDD02 & "' AND " & _
                           "PA03 = '" & strDD03 & "' AND " & _
                           "PA04 = '" & strDD04 & "' "
         ' 讀取法務基本檔
         'Modify By Sindy 2009/07/24 增加LIN系統類別
         'modify by sonia 2019/7/29 +ACS系統類別
         Case "L", "CFL", "FCL", "LIN", "ACS":
            strSql = "DELETE FROM LAWCASE " & _
                     "WHERE LC01 = '" & strDD01 & "' AND " & _
                           "LC02 = '" & strDD02 & "' AND " & _
                           "LC03 = '" & strDD03 & "' AND " & _
                           "LC04 = '" & strDD04 & "' "
         ' 讀取顧問案件基本檔
         Case "LA":
            strSql = "DELETE FROM HIRECASE " & _
                     "WHERE HC01 = '" & strDD01 & "' AND " & _
                           "HC02 = '" & strDD02 & "' AND " & _
                           "HC03 = '" & strDD03 & "' AND " & _
                           "HC04 = '" & strDD04 & "' "
         ' 讀取服務業務基本檔
         Case Else:
            strSql = "DELETE FROM SERVICEPRACTICE " & _
                     "WHERE SP01 = '" & strDD01 & "' AND " & _
                           "SP02 = '" & strDD02 & "' AND " & _
                           "SP03 = '" & strDD03 & "' AND " & _
                           "SP04 = '" & strDD04 & "' "
      End Select
      ' 執行刪除的指令
      cnnConnection.Execute strSql
   End If
   
   ' 刪除優先權檔
   strSql = "DELETE FROM PRIDATE " & _
            "WHERE PD01 = '" & strDD01 & "' AND " & _
                  "PD02 = '" & strDD02 & "' AND " & _
                  "PD03 = '" & strDD03 & "' AND " & _
                  "PD04 = '" & strDD04 & "' "
   cnnConnection.Execute strSql
   
   ' 刪除相關卷號檔
   strSql = "DELETE FROM CASERELATION " & _
            "WHERE (CR01 = '" & strDD01 & "' AND " & _
                   "CR02 = '" & strDD02 & "' AND " & _
                   "CR03 = '" & strDD03 & "' AND " & _
                   "CR04 = '" & strDD04 & "') OR " & _
                  "(CR05 = '" & strDD01 & "' AND " & _
                   "CR06 = '" & strDD02 & "' AND " & _
                   "CR07 = '" & strDD03 & "' AND " & _
                   "CR08 = '" & strDD04 & "') "
   cnnConnection.Execute strSql
   
   'add by nickc 2006/06/22 刪除相關卷號檔  留 2006/11/21
   strSql = "DELETE FROM CASERELATION1 " & _
            "WHERE (CR01 = '" & strDD01 & "' AND " & _
                   "CR02 = '" & strDD02 & "' AND " & _
                   "CR03 = '" & strDD03 & "' AND " & _
                   "CR04 = '" & strDD04 & "') OR " & _
                  "(CR05 = '" & strDD01 & "' AND " & _
                   "CR06 = '" & strDD02 & "' AND " & _
                   "CR07 = '" & strDD03 & "' AND " & _
                   "CR08 = '" & strDD04 & "') "
   cnnConnection.Execute strSql
End Sub

' 檢查案件進度檔是否存在
Private Function ExistProgress(ByVal strCP01 As String, ByVal strCP02 As String, ByVal strCP03 As String, ByVal strCP04 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String

   ExistProgress = False
   strSql = "SELECT CP09 FROM CASEPROGRESS " & _
            "WHERE CP01 = '" & strCP01 & "' AND " & _
                  "CP02 = '" & strCP02 & "' AND " & _
                  "CP03 = '" & strCP03 & "' AND " & _
                  "CP04 = '" & strCP04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ExistProgress = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'Add By Sindy 2010/11/25
Private Sub textDD23_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 失誤人員代號
Private Sub textDD23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textDD23_2 = Empty
   If IsEmptyText(textDD23) = False Then
      textDD23_2 = GetStaffName(textDD23, False)
      If IsEmptyText(textDD23_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "失誤人員代號錯誤"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDD23_GotFocus
      End If
   End If
End Sub

' 刪除備註
Private Sub textDD24_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textDD24, 60) = False Then
      Cancel = True
      textDD24_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textDD24.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 檢查輸入的資料是否完整
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   
   ' 失誤人員
   If IsEmptyText(textDD23) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入失誤人員"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textDD23.SetFocus
      GoTo EXITSUB
   End If
   
   ' 刪除備註
   If IsEmptyText(textDD24) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入刪除備註"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textDD24.SetFocus
      GoTo EXITSUB
   End If
   
    'Added by Lydia 2022/02/16 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        GoTo EXITSUB
    End If

   CheckDataValid = True
EXITSUB:
End Function

Private Sub textDD23_GotFocus()
   InverseTextBox textDD23
End Sub

Private Sub textDD24_GotFocus()
   InverseTextBox textDD24
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textDD24.IMEMode = 1
   OpenIme
End Sub

