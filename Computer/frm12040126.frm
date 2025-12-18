VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040126 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件改號作業"
   ClientHeight    =   7170
   ClientLeft      =   110
   ClientTop       =   1850
   ClientWidth     =   7070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   7070
   Begin VB.TextBox textCP13 
      Height          =   264
      Left            =   1710
      MaxLength       =   6
      TabIndex        =   11
      Top             =   4728
      Width           =   1092
   End
   Begin VB.TextBox textOEName 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Index           =   1
      Left            =   1728
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4032
      Width           =   5052
   End
   Begin VB.TextBox textDelete 
      Height          =   264
      Left            =   2280
      MaxLength       =   9
      TabIndex        =   12
      Top             =   5172
      Width           =   495
   End
   Begin VB.TextBox textStatus 
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5532
      Width           =   6612
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   5988
      TabIndex        =   14
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5160
      TabIndex        =   13
      Top             =   120
      Width           =   800
   End
   Begin VB.TextBox textOEName 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Index           =   0
      Left            =   1728
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1392
      Width           =   5052
   End
   Begin VB.TextBox textNEW01 
      Height          =   264
      Left            =   1728
      MaxLength       =   3
      TabIndex        =   6
      Top             =   3264
      Width           =   732
   End
   Begin VB.TextBox textNEW03 
      Height          =   264
      Left            =   3552
      MaxLength       =   1
      TabIndex        =   9
      Top             =   3264
      Width           =   372
   End
   Begin VB.TextBox textNEW04 
      Height          =   264
      Left            =   3912
      MaxLength       =   2
      TabIndex        =   10
      Top             =   3264
      Width           =   732
   End
   Begin VB.TextBox textNEW02_2 
      Height          =   264
      Left            =   3168
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3264
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textOLD01 
      Height          =   264
      Left            =   1728
      MaxLength       =   3
      TabIndex        =   0
      Top             =   672
      Width           =   732
   End
   Begin VB.TextBox textOLD03 
      Height          =   264
      Left            =   3552
      MaxLength       =   1
      TabIndex        =   3
      Top             =   672
      Width           =   372
   End
   Begin VB.TextBox textOLD04 
      Height          =   264
      Left            =   3912
      MaxLength       =   2
      TabIndex        =   4
      Top             =   672
      Width           =   732
   End
   Begin VB.TextBox textOLD02_2 
      Height          =   264
      Left            =   3168
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   672
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textOLD02 
      Height          =   264
      Left            =   2448
      MaxLength       =   6
      TabIndex        =   1
      Top             =   672
      Width           =   1092
   End
   Begin VB.TextBox textNEW02 
      Height          =   264
      Left            =   2448
      MaxLength       =   6
      TabIndex        =   7
      Top             =   3264
      Width           =   1092
   End
   Begin MSForms.TextBox textOMemo 
      Height          =   1068
      Left            =   1728
      TabIndex        =   5
      Top             =   2088
      Width           =   5076
      VariousPropertyBits=   -1467987939
      ScrollBars      =   2
      Size            =   "8954;1884"
      FontName        =   "新細明體"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label16 
      Caption         =   "案件備註 :"
      Height          =   252
      Left            =   168
      TabIndex        =   39
      Top             =   2064
      Width           =   1332
   End
   Begin MSForms.TextBox textCP13_2 
      Height          =   300
      Left            =   2856
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4692
      Width           =   1692
      VariousPropertyBits=   671105055
      Size            =   "2984;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textOJName 
      Height          =   300
      Index           =   1
      Left            =   1728
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4368
      Width           =   5052
      VariousPropertyBits=   671105055
      Size            =   "8911;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textOJName 
      Height          =   300
      Index           =   0
      Left            =   1728
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1710
      Width           =   5050
      VariousPropertyBits=   671105055
      Size            =   "8911;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textOCName 
      Height          =   300
      Index           =   1
      Left            =   1728
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3648
      Width           =   5052
      VariousPropertyBits=   671105055
      Size            =   "8911;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textOCName 
      Height          =   300
      Index           =   0
      Left            =   1728
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1020
      Width           =   5050
      VariousPropertyBits=   671105055
      Size            =   "8911;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   0
      Left            =   168
      TabIndex        =   38
      Top             =   4764
      Width           =   900
   End
   Begin VB.Label Label15 
      Caption         =   "　　 5.下一程序未過期期限之智權人員要一併修改"
      ForeColor       =   &H000000C0&
      Height          =   192
      Left            =   156
      TabIndex        =   36
      Top             =   6912
      Width           =   6156
   End
   Begin VB.Label Label14 
      Caption         =   "　　 4. 檢查下一程序是否有未到期但已不續辦期限, 同時通知專業部程序注意"
      ForeColor       =   &H000000C0&
      Height          =   192
      Left            =   156
      TabIndex        =   35
      Top             =   6672
      Width           =   6156
   End
   Begin VB.Label Label13 
      Caption         =   "　　 3. FC案轉國內案, 要詢問專業部, 基本檔FC資料是否取消?"
      ForeColor       =   &H000000C0&
      Height          =   192
      Left            =   156
      TabIndex        =   34
      Top             =   6432
      Width           =   6156
   End
   Begin VB.Label Label12 
      Caption         =   "　　 2. P案轉FCP案要詢問阮威立此案工程師組別,並通知自行處理說明書檔案"
      ForeColor       =   &H000000C0&
      Height          =   192
      Left            =   156
      TabIndex        =   33
      Top             =   6192
      Width           =   6156
   End
   Begin VB.Label Label11 
      Caption         =   "注意 : 只有台灣案才可轉至國外部"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label10 
      Caption         =   "PS :  1. 跨部門改案號時, 請通知""財務處總帳人員"", 自行做傳票調整規費科目"
      ForeColor       =   &H000000C0&
      Height          =   192
      Left            =   156
      TabIndex        =   31
      Top             =   5952
      Width           =   6156
   End
   Begin VB.Label Label9 
      Caption         =   "案件名稱(中) :"
      Height          =   252
      Left            =   168
      TabIndex        =   30
      Top             =   3672
      Width           =   1332
   End
   Begin VB.Label Label8 
      Caption         =   "案件名稱(英) :"
      Height          =   252
      Left            =   168
      TabIndex        =   29
      Top             =   4032
      Width           =   1332
   End
   Begin VB.Label Label7 
      Caption         =   "案件名稱(日) :"
      Height          =   252
      Index           =   1
      Left            =   168
      TabIndex        =   27
      Top             =   4392
      Width           =   1332
   End
   Begin VB.Label Label6 
      Caption         =   "原案件基本資料是否刪除:                 (Y: 刪除)"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   156
      TabIndex        =   24
      Top             =   5172
      Width           =   3948
   End
   Begin VB.Label Label5 
      Caption         =   "案件名稱(日) :"
      Height          =   252
      Left            =   168
      TabIndex        =   22
      Top             =   1752
      Width           =   1332
   End
   Begin VB.Label Label4 
      Caption         =   "案件名稱(英) :"
      Height          =   252
      Left            =   168
      TabIndex        =   21
      Top             =   1392
      Width           =   1332
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱(中) :"
      Height          =   252
      Left            =   168
      TabIndex        =   18
      Top             =   1032
      Width           =   1332
   End
   Begin VB.Label Label2 
      Caption         =   "新本所案號 :"
      Height          =   252
      Left            =   168
      TabIndex        =   16
      Top             =   3264
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "原本所案號 :"
      Height          =   252
      Left            =   168
      TabIndex        =   15
      Top             =   672
      Width           =   1092
   End
End
Attribute VB_Name = "frm12040126"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/27 Form2.0已修改(textOCName(0),textOJName(0),textOCName(1),textOJName(1),textCP13_2)
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

'Modify by Amy 2018/07/24 改陣列
'Dim m_OLD01 As String
'Dim m_OLD02 As String
'Dim m_OLD03 As String
'Dim m_OLD04 As String
'
'Dim m_NEW01 As String
'Dim m_NEW02 As String
'Dim m_NEW03 As String
'Dim m_NEW04 As String
Dim m_OLD(1 To 4) As String
Dim m_NEW(1 To 4) As String

Private Sub Form_Load()
   textOCName(0).BackColor = &H8000000F
   textOEName(0).BackColor = &H8000000F
   textOJName(0).BackColor = &H8000000F
   textOCName(1).BackColor = &H8000000F
   textOEName(1).BackColor = &H8000000F
   textOJName(1).BackColor = &H8000000F
   textStatus.BackColor = &H8000000F
   MoveFormToCenter Me
End Sub

Private Sub cmdok_Click()
   If CheckDataValid() = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      
      m_OLD(1) = textOLD01
      m_OLD(2) = textOLD02
      If m_OLD(1) = "TF" Then: m_OLD(2) = m_OLD(2) & textOLD02_2
      m_OLD(3) = textOLD03 & String(1 - Len(textOLD03), "0")
      m_OLD(4) = textOLD04 & String(2 - Len(textOLD04), "0")
      
      m_NEW(1) = textNEW01
      m_NEW(2) = textNEW02
      If m_NEW(1) = "TF" Then: m_NEW(2) = m_NEW(2) & textNEW02_2
      m_NEW(3) = textNEW03 & String(1 - Len(textNEW03), "0")
      m_NEW(4) = textNEW04 & String(2 - Len(textNEW04), "0")
      
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 執行更新的作業
        'Modify By Cheng 2002/11/08
'      OnSaveData
        If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 設定滑鼠游標為預設
      'Modify By Sindy 2017/3/23 有關卷宗區,原始檔,歷程附件檔的電子檔名更改是寫在Trigger:caseprogress_before1
      Screen.MousePointer = vbDefault
      
      ' 清除欄位內容
      textOLD01 = Empty
      textOLD02 = Empty
      textOLD02_2 = Empty
      textOLD03 = Empty
      textOLD04 = Empty
      textOCName(0) = Empty
      textOEName(0) = Empty
      textOJName(0) = Empty
      textOCName(1) = Empty
      textOEName(1) = Empty
      textOJName(1) = Empty
      textNEW01 = Empty
      textNEW02 = Empty
      textNEW02_2 = Empty
      textNEW03 = Empty
      textNEW04 = Empty
      textStatus = Empty
      textOMemo = Empty 'Added by Lydia 2025/10/15
      ' 設定輸入欄位
      textOLD01.SetFocus
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

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
      ' 案件名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then
         textOCName(0) = rsTmp.Fields("TM05")
      End If
      ' 案件名稱
      If IsNull(rsTmp.Fields("TM06")) = False Then
         textOEName(0) = rsTmp.Fields("TM06")
      End If
      ' 案件名稱
      If IsNull(rsTmp.Fields("TM07")) = False Then
         textOJName(0) = rsTmp.Fields("TM07")
      End If
      textOMemo = "" & rsTmp.Fields("TM58") 'Added by Lydia 2025/10/15
      
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
      ' 案件名稱
      If IsNull(rsTmp.Fields("SP05")) = False Then
         textOCName(0) = rsTmp.Fields("SP05")
      End If
      If IsNull(rsTmp.Fields("SP06")) = False Then
         textOEName(0) = rsTmp.Fields("SP06")
      End If
      If IsNull(rsTmp.Fields("SP07")) = False Then
         textOJName(0) = rsTmp.Fields("SP07")
      End If
      textOMemo = "" & rsTmp.Fields("SP18") 'Added by Lydia 2025/10/15
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
      ' 案件名稱
      If IsNull(rsTmp.Fields("PA05")) = False Then
         textOCName(0) = rsTmp.Fields("PA05")
      End If
      If IsNull(rsTmp.Fields("PA06")) = False Then
         textOEName(0) = rsTmp.Fields("PA06")
      End If
      If IsNull(rsTmp.Fields("PA07")) = False Then
         textOJName(0) = rsTmp.Fields("PA07")
      End If
      textOMemo = "" & rsTmp.Fields("PA91") 'Added by Lydia 2025/10/15
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
      ' 案件名稱
      If IsNull(rsTmp.Fields("LC05")) = False Then
         textOCName(0) = rsTmp.Fields("LC05")
      End If
      If IsNull(rsTmp.Fields("LC06")) = False Then
         textOEName(0) = rsTmp.Fields("LC06")
      End If
      If IsNull(rsTmp.Fields("LC07")) = False Then
         textOJName(0) = rsTmp.Fields("LC07")
      End If
      textOMemo = "" & rsTmp.Fields("LC27") 'Added by Lydia 2025/10/15
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
      ' 案件名稱
      If IsNull(rsTmp.Fields("HC06")) = False Then
         textOCName(0) = rsTmp.Fields("HC06")
      End If
      textOMemo = "" & rsTmp.Fields("HC12") 'Added by Lydia 2025/10/15
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'Modify By Cheng 2002/11/0
'Private Sub OnSaveData()
Private Function OnSaveData() As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim StrSQLa As String   'add by sonia 2014/12/27
'Add by Amy 2018/07/24
Dim i As Integer, intCount As Integer
Dim bolDelFile As Boolean
Dim strNewFileName As String
Dim strOldFileName As String
Dim stCP09 As String, stCP12 As String, stCP14 As String 'Add By Sindy 2019/12/30
   
'Add By Cheng 2002/11/08
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   
   '2010/10/12 自最下面移上來,因trigger"PGMID.CASEPROGRESS_BEFORE2"要抓基本檔才能更新cp139所以要先改基本檔
   '2011/3/30 MODIFY BY SONIA 原本所案號之備註加在該欄之前面並加註日期
   If textDelete <> "Y" Then
      ' 更新基本檔
      Select Case m_OLD(1)
         ' 商標基本檔
         Case "T", "TF", "CFT", "FCT":
            ShowStatus "變更商標基本檔中, 請稍候 . . ."
            'add by sonia 2014/12/27 T案改FCT案時,TM53為1者清除,TM77為1或2者清除
            'strSql = "UPDATE TRADEMARK SET TM01 = '" & m_NEW01 & "', " & _
                                          "TM02 = '" & m_NEW02 & "', " & _
                                          "TM03 = '" & m_NEW03 & "', " & _
                                          "TM04 = '" & m_NEW04 & "', " & _
                                          "TM58 = '" & ChangeTStringToTDateString(strSrvDate(2)) & "改本所案號,原本所案號" & m_OLD01 & "-" & m_OLD02 & "-" & m_OLD03 & "-" & m_OLD04 & ";' || TM58 " & _
                     "WHERE TM01 = '" & m_OLD01 & "' AND " & _
                           "TM02 = '" & m_OLD02 & "' AND " & _
                           "TM03 = '" & m_OLD03 & "' AND " & _
                           "TM04 = '" & m_OLD04 & "' "
            StrSQLa = ""
            If m_OLD(1) = "T" And m_NEW(1) = "FCT" Then
               StrSQLa = ", TM53 = DECODE(TM53,'1',NULL,TM53), TM77 = DECODE(TM77,'1',NULL,'2',NULL,TM77) "
            End If
            strSql = "UPDATE TRADEMARK SET TM01 = '" & m_NEW(1) & "', " & _
                                          "TM02 = '" & m_NEW(2) & "', " & _
                                          "TM03 = '" & m_NEW(3) & "', " & _
                                          "TM04 = '" & m_NEW(4) & "', " & _
                                          "TM58 = '" & ChangeTStringToTDateString(strSrvDate(2)) & "改本所案號,原本所案號" & m_OLD(1) & "-" & m_OLD(2) & "-" & m_OLD(3) & "-" & m_OLD(4) & ";' || TM58 " & _
                     "" & StrSQLa & "WHERE TM01 = '" & m_OLD(1) & "' AND " & _
                                          "TM02 = '" & m_OLD(2) & "' AND " & _
                                          "TM03 = '" & m_OLD(3) & "' AND " & _
                                          "TM04 = '" & m_OLD(4) & "' "
            'end 2014/12/27
         ' 專利基本檔
         Case "P", "CFP", "FCP":
            ShowStatus "變更專利基本檔中, 請稍候 . . ."
            strSql = "UPDATE PATENT SET PA01 = '" & m_NEW(1) & "', " & _
                                       "PA02 = '" & m_NEW(2) & "', " & _
                                       "PA03 = '" & m_NEW(3) & "', " & _
                                       "PA04 = '" & m_NEW(4) & "', " & _
                                       "PA91 = '" & ChangeTStringToTDateString(strSrvDate(2)) & "改本所案號,原本所案號" & m_OLD(1) & "-" & m_OLD(2) & "-" & m_OLD(3) & "-" & m_OLD(4) & ";' || PA91 " & _
                     "WHERE PA01 = '" & m_OLD(1) & "' AND " & _
                           "PA02 = '" & m_OLD(2) & "' AND " & _
                           "PA03 = '" & m_OLD(3) & "' AND " & _
                           "PA04 = '" & m_OLD(4) & "' "
         ' 法務基本檔
         Case "L", "CFL", "FCL":
            ShowStatus "變更法務基本檔中, 請稍候 . . ."
            strSql = "UPDATE LAWCASE SET LC01 = '" & m_NEW(1) & "', " & _
                                        "LC02 = '" & m_NEW(2) & "', " & _
                                        "LC03 = '" & m_NEW(3) & "', " & _
                                        "LC04 = '" & m_NEW(4) & "', " & _
                                        "LC27 = '" & ChangeTStringToTDateString(strSrvDate(2)) & "改本所案號,原本所案號" & m_OLD(1) & "-" & m_OLD(2) & "-" & m_OLD(3) & "-" & m_OLD(4) & ";' || LC27 " & _
                     "WHERE LC01 = '" & m_OLD(1) & "' AND " & _
                           "LC02 = '" & m_OLD(2) & "' AND " & _
                           "LC03 = '" & m_OLD(3) & "' AND " & _
                           "LC04 = '" & m_OLD(4) & "' "
         ' 顧問案件基本檔
         Case "LA":
            ShowStatus "變更顧問基本檔中, 請稍候 . . ."
            strSql = "UPDATE HIRECASE SET HC01 = '" & m_NEW(1) & "', " & _
                                         "HC02 = '" & m_NEW(2) & "', " & _
                                         "HC03 = '" & m_NEW(3) & "', " & _
                                         "HC04 = '" & m_NEW(4) & "', " & _
                                         "HC12 = '" & ChangeTStringToTDateString(strSrvDate(2)) & "改本所案號,原本所案號" & m_OLD(1) & "-" & m_OLD(2) & "-" & m_OLD(3) & "-" & m_OLD(4) & ";' || HC12 " & _
                     "WHERE HC01 = '" & m_OLD(1) & "' AND " & _
                           "HC02 = '" & m_OLD(2) & "' AND " & _
                           "HC03 = '" & m_OLD(3) & "' AND " & _
                           "HC04 = '" & m_OLD(4) & "' "
         ' 服務業務基本檔
         Case Else:
            ShowStatus "變更服務業務基本檔中, 請稍候 . . ."
            strSql = "UPDATE SERVICEPRACTICE SET SP01 = '" & m_NEW(1) & "', " & _
                                                "SP02 = '" & m_NEW(2) & "', " & _
                                                "SP03 = '" & m_NEW(3) & "', " & _
                                                "SP04 = '" & m_NEW(4) & "', " & _
                                                "SP18 = '" & ChangeTStringToTDateString(strSrvDate(2)) & "改本所案號,原本所案號" & m_OLD(1) & "-" & m_OLD(2) & "-" & m_OLD(3) & "-" & m_OLD(4) & ";' || SP18 " & _
                     "WHERE SP01 = '" & m_OLD(1) & "' AND " & _
                           "SP02 = '" & m_OLD(2) & "' AND " & _
                           "SP03 = '" & m_OLD(3) & "' AND " & _
                           "SP04 = '" & m_OLD(4) & "' "
      End Select
      Pub_SeekTbLog strSql   '2009/3/27 ADD BY SONIA 新增維護記錄檔
      cnnConnection.Execute strSql
      strSql = "UPDATE DML_Log SET DL01 = '" & m_NEW(1) & "', " & _
                                  "DL02 = '" & m_NEW(2) & "', " & _
                                  "DL03 = '" & m_NEW(3) & "', " & _
                                  "DL04 = '" & m_NEW(4) & "' " & _
                            "WHERE DL01 = '" & m_OLD(1) & "' AND " & _
                                  "DL02 = '" & m_OLD(2) & "' AND " & _
                                  "DL03 = '" & m_OLD(3) & "' AND " & _
                                  "DL04 = '" & m_OLD(4) & "' "
      cnnConnection.Execute strSql
   '2011/6/15 因計件值等trigger會錯誤,故將刪除基本檔改到最後再刪
   End If
   
   ' 更新案件進度檔
   ShowStatus "變更案件進度檔中, 請稍候 . . ."
   'Modify By Sindy 2023/4/6 + 更新A類的北所分案日期=19221111
   strSql = "UPDATE CASEPROGRESS SET CP01 = '" & m_NEW(1) & "', " & _
                                    "CP02 = '" & m_NEW(2) & "', " & _
                                    "CP03 = '" & m_NEW(3) & "', " & _
                                    "CP04 = '" & m_NEW(4) & "', " & _
                                    "CP157=decode(substr(CP09,1,1),'A',decode(CP157,null,19221111,0,19221111,CP157),cp157)," & _
                                    "CP64 = '" & ChangeTStringToTDateString(strSrvDate(2)) & "改本所案號,原本所案號" & m_OLD(1) & "-" & m_OLD(2) & "-" & m_OLD(3) & "-" & m_OLD(4) & ";' || CP64 " & _
                  "WHERE CP01 = '" & m_OLD(1) & "' AND " & _
                        "CP02 = '" & m_OLD(2) & "' AND " & _
                        "CP03 = '" & m_OLD(3) & "' AND " & _
                        "CP04 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   ' 更新下一程序檔
   ShowStatus "變更下一程序檔中, 請稍候 . . ."
   strSql = "UPDATE NEXTPROGRESS SET NP02 = '" & m_NEW(1) & "', " & _
                                    "NP03 = '" & m_NEW(2) & "', " & _
                                    "NP04 = '" & m_NEW(3) & "', " & _
                                    "NP05 = '" & m_NEW(4) & "' " & _
                  "WHERE NP02 = '" & m_OLD(1) & "' AND " & _
                        "NP03 = '" & m_OLD(2) & "' AND " & _
                        "NP04 = '" & m_OLD(3) & "' AND " & _
                        "NP05 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   ' 更新相關卷號檔
   ShowStatus "變更相關卷號檔中, 請稍候 . . ."
   strSql = "UPDATE CASERELATION SET CR01 = '" & m_NEW(1) & "', " & _
                                    "CR02 = '" & m_NEW(2) & "', " & _
                                    "CR03 = '" & m_NEW(3) & "', " & _
                                    "CR04 = '" & m_NEW(4) & "' " & _
                  "WHERE CR01 = '" & m_OLD(1) & "' AND " & _
                        "CR02 = '" & m_OLD(2) & "' AND " & _
                        "CR03 = '" & m_OLD(3) & "' AND " & _
                        "CR04 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE CASERELATION SET CR05 = '" & m_NEW(1) & "', " & _
                                    "CR06 = '" & m_NEW(2) & "', " & _
                                    "CR07 = '" & m_NEW(3) & "', " & _
                                    "CR08 = '" & m_NEW(4) & "' " & _
                  "WHERE CR05 = '" & m_OLD(1) & "' AND " & _
                        "CR06 = '" & m_OLD(2) & "' AND " & _
                        "CR07 = '" & m_OLD(3) & "' AND " & _
                        "CR08 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   ' 更新相關卷號1檔
   ShowStatus "變更相關卷號檔中, 請稍候 . . ."
   strSql = "UPDATE CASERELATION1 SET CR01 = '" & m_NEW(1) & "', " & _
                                    "CR02 = '" & m_NEW(2) & "', " & _
                                    "CR03 = '" & m_NEW(3) & "', " & _
                                    "CR04 = '" & m_NEW(4) & "' " & _
                  "WHERE CR01 = '" & m_OLD(1) & "' AND " & _
                        "CR02 = '" & m_OLD(2) & "' AND " & _
                        "CR03 = '" & m_OLD(3) & "' AND " & _
                        "CR04 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE CASERELATION1 SET CR05 = '" & m_NEW(1) & "', " & _
                                    "CR06 = '" & m_NEW(2) & "', " & _
                                    "CR07 = '" & m_NEW(3) & "', " & _
                                    "CR08 = '" & m_NEW(4) & "' " & _
                  "WHERE CR05 = '" & m_OLD(1) & "' AND " & _
                        "CR06 = '" & m_OLD(2) & "' AND " & _
                        "CR07 = '" & m_OLD(3) & "' AND " & _
                        "CR08 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   ' 更新國內外案件關聯表
   ShowStatus "變更國內外案件關聯表中, 請稍候 . . ."
   strSql = "UPDATE CASEMAP SET CM01 = '" & m_NEW(1) & "', " & _
                               "CM02 = '" & m_NEW(2) & "', " & _
                               "CM03 = '" & m_NEW(3) & "', " & _
                               "CM04 = '" & m_NEW(4) & "' " & _
                  "WHERE CM01 = '" & m_OLD(1) & "' AND " & _
                        "CM02 = '" & m_OLD(2) & "' AND " & _
                        "CM03 = '" & m_OLD(3) & "' AND " & _
                        "CM04 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE CASEMAP SET CM05 = '" & m_NEW(1) & "', " & _
                               "CM06 = '" & m_NEW(2) & "', " & _
                               "CM07 = '" & m_NEW(3) & "', " & _
                               "CM08 = '" & m_NEW(4) & "' " & _
                  "WHERE CM05 = '" & m_OLD(1) & "' AND " & _
                        "CM06 = '" & m_OLD(2) & "' AND " & _
                        "CM07 = '" & m_OLD(3) & "' AND " & _
                        "CM08 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   ' 更新優先權資料檔
   ShowStatus "變更優先權資料中, 請稍候 . . ."
   strSql = "UPDATE PRIDATE SET PD01 = '" & m_NEW(1) & "', " & _
                               "PD02 = '" & m_NEW(2) & "', " & _
                               "PD03 = '" & m_NEW(3) & "', " & _
                               "PD04 = '" & m_NEW(4) & "' " & _
                  "WHERE PD01 = '" & m_OLD(1) & "' AND " & _
                        "PD02 = '" & m_OLD(2) & "' AND " & _
                        "PD03 = '" & m_OLD(3) & "' AND " & _
                        "PD04 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   'Added by Morgan 2014/10/17
   strSql = "UPDATE PRIDATE SET PD06 = '" & m_NEW(1) & m_NEW(2) & m_NEW(3) & m_NEW(4) & "' " & _
                  "WHERE PD06 = '" & m_OLD(1) & m_OLD(2) & m_OLD(3) & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   'end 2014/10/17
   
   'Add By Sindy 2022/10/21 更新發明人資料檔
   ShowStatus "變更發明人資料檔中, 請稍候 . . ."
   strSql = "UPDATE patentInventor SET pi01 = '" & m_NEW(1) & "', " & _
                               "pi02 = '" & m_NEW(2) & "', " & _
                               "pi03 = '" & m_NEW(3) & "', " & _
                               "pi04 = '" & m_NEW(4) & "' " & _
                  "WHERE pi01 = '" & m_OLD(1) & "' AND " & _
                        "pi02 = '" & m_OLD(2) & "' AND " & _
                        "pi03 = '" & m_OLD(3) & "' AND " & _
                        "pi04 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   '2022/10/21 END
   
   ' 更新CFP案准駁記錄檔
   ShowStatus "變更CFP案核准記錄檔中, 請稍候 . . ."
   strSql = "UPDATE PERMITRECORD SET PR01 = '" & m_NEW(1) & "', " & _
                                    "PR02 = '" & m_NEW(2) & "', " & _
                                    "PR03 = '" & m_NEW(3) & "', " & _
                                    "PR04 = '" & m_NEW(4) & "' " & _
                  "WHERE PR01 = '" & m_OLD(1) & "' AND " & _
                        "PR02 = '" & m_OLD(2) & "' AND " & _
                        "PR03 = '" & m_OLD(3) & "' AND " & _
                        "PR04 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   ' 更新來函記錄檔
   ShowStatus "變更來函記錄檔中, 請稍候 . . ."
   strSql = "UPDATE MAILREC SET MR12 = '" & m_NEW(1) & "', " & _
                               "MR13 = '" & m_NEW(2) & "', " & _
                               "MR14 = '" & m_NEW(3) & "', " & _
                               "MR15 = '" & m_NEW(4) & "' " & _
                  "WHERE MR12 = '" & m_OLD(1) & "' AND " & _
                        "MR13 = '" & m_OLD(2) & "' AND " & _
                        "MR14 = '" & m_OLD(3) & "' AND " & _
                        "MR15 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   ' 更新資料刪除記錄檔
   ShowStatus "變更資料刪除記錄檔中, 請稍候 . . ."
   strSql = "UPDATE DATADELETERECORD SET DD01 = '" & m_NEW(1) & "', " & _
                                        "DD02 = '" & m_NEW(2) & "', " & _
                                        "DD03 = '" & m_NEW(3) & "', " & _
                                        "DD04 = '" & m_NEW(4) & "' " & _
                  "WHERE DD01 = '" & m_OLD(1) & "' AND " & _
                        "DD02 = '" & m_OLD(2) & "' AND " & _
                        "DD03 = '" & m_OLD(3) & "' AND " & _
                        "DD04 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   ' 更新智慧局送件清單明細檔
   ShowStatus "變更智慧局送件清單明細檔中, 請稍候 . . ."
   strSql = "UPDATE applistdetail SET ald05 = '" & m_NEW(1) & "', " & _
                                     "ald06 = '" & m_NEW(2) & "', " & _
                                     "ald07 = '" & m_NEW(3) & "', " & _
                                     "ald08 = '" & m_NEW(4) & "' " & _
                  "WHERE ald05 = '" & m_OLD(1) & "' AND " & _
                        "ald06 = '" & m_OLD(2) & "' AND " & _
                        "ald07 = '" & m_OLD(3) & "' AND " & _
                        "ald08 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   ' 更新轉傳票分錄資料
   ShowStatus "變更轉傳票分錄資料中, 請稍候 . . ."
   strSql = "UPDATE ACC1P0 SET A1P17 = '" & m_NEW(1) & m_NEW(2) & m_NEW(3) & m_NEW(4) & "' " & _
                  "WHERE A1P17 = '" & m_OLD(1) & m_OLD(2) & m_OLD(3) & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   '2007/12/17 ADD BY SONIA
   ' 更新國外請款資料(明細檔)
   strSql = "UPDATE ACC1L0 SET A1L03 = '" & m_NEW(1) & "'  WHERE A1L01 IN ( " & _
                  "SELECT A1K01 FROM ACC1K0 " & _
                  "WHERE A1K13 = '" & m_OLD(1) & "' AND " & _
                        "A1K14 = '" & m_OLD(2) & "' AND " & _
                        "A1K15 = '" & m_OLD(3) & "' AND " & _
                        "A1K16 = '" & m_OLD(4) & "') "
   cnnConnection.Execute strSql
   '2007/12/17 END
   
   ' 更新國外請款資料(主檔)
   ShowStatus "變更國外請款資料中, 請稍候 . . ."
   '2013/10/17 MODIFY BY SONIA A1K05改為A1K34
   strSql = "UPDATE ACC1K0 SET A1K13 = '" & m_NEW(1) & "', " & _
                              "A1K14 = '" & m_NEW(2) & "', " & _
                              "A1K15 = '" & m_NEW(3) & "', " & _
                              "A1K16 = '" & m_NEW(4) & "', " & _
                              "A1K34 = '" & ChangeTStringToTDateString(strSrvDate(2)) & "改本所案號,原本所案號" & m_OLD(1) & "-" & m_OLD(2) & "-" & m_OLD(3) & "-" & m_OLD(4) & ";' || A1K34 " & _
                  "WHERE A1K13 = '" & m_OLD(1) & "' AND " & _
                        "A1K14 = '" & m_OLD(2) & "' AND " & _
                        "A1K15 = '" & m_OLD(3) & "' AND " & _
                        "A1K16 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   ' 更新國外抵帳單資料(交易檔)
   ShowStatus "變更國外抵帳單資料中, 請稍候 . . ."
   strSql = "UPDATE ACC161 SET AXG03 = '" & m_NEW(1) & m_NEW(2) & m_NEW(3) & m_NEW(4) & "' " & _
                  "WHERE AXG03 = '" & m_OLD(1) & m_OLD(2) & m_OLD(3) & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   ' 更新國外帳單資料(交易檔)
   ShowStatus "變更國外帳單資料中, 請稍候 . . ."
   strSql = "UPDATE ACC151 SET AXF03 = '" & m_NEW(1) & m_NEW(2) & m_NEW(3) & m_NEW(4) & "' " & _
                  "WHERE AXF03 = '" & m_OLD(1) & m_OLD(2) & m_OLD(3) & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   ' 更新國外暫收款資料
   ShowStatus "變更國外暫收款資料中, 請稍候 . . ."
   strSql = "UPDATE ACC120 SET A1208 = '" & m_NEW(1) & m_NEW(2) & m_NEW(3) & m_NEW(4) & "' " & _
                  "WHERE A1208 = '" & m_OLD(1) & m_OLD(2) & m_OLD(3) & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   ' 更新國內未開收據案件資料(暫存檔)
   ShowStatus "變更國內未開收據案件資料中, 請稍候 . . ."
   strSql = "UPDATE ACC0J0 SET A0J02 = '" & m_NEW(1) & m_NEW(2) & m_NEW(3) & m_NEW(4) & "' " & _
                  "WHERE A0J02 = '" & m_OLD(1) & m_OLD(2) & m_OLD(3) & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   '93.5.6 瑞婷說不更新傳票對沖資料 '94/4/1取消此控制
   '' 更新傳票資料(交易檔-財務)
   ShowStatus "變更傳票資料中, 請稍候 . . ."
   strSql = "UPDATE ACC021 SET AX214 = '" & m_NEW(1) & m_NEW(2) & m_NEW(3) & m_NEW(4) & "' " & _
                  "WHERE AX214 = '" & m_OLD(1) & m_OLD(2) & m_OLD(3) & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   '
   '' 更新傳票資料(交易檔-帳務)
   ShowStatus "變更傳票資料中, 請稍候 . . ."
   strSql = "UPDATE ACC031 SET AX314 = '" & m_NEW(1) & m_NEW(2) & m_NEW(3) & m_NEW(4) & "' " & _
                  "WHERE AX314 = '" & m_OLD(1) & m_OLD(2) & m_OLD(3) & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   '93.5.6 END
   
   ' 更新國外結匯資料
   ShowStatus "變更國外結匯資料中, 請稍候 . . ."
   strSql = "UPDATE ACC170 SET A1707 = '" & m_NEW(1) & m_NEW(2) & m_NEW(3) & m_NEW(4) & "' " & _
                  "WHERE A1707 = '" & m_OLD(1) & m_OLD(2) & m_OLD(3) & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   ' 更新結餘資料
   ShowStatus "變更結餘資料中, 請稍候 . . ."
   strSql = "UPDATE acc240 SET a240005 = '" & m_NEW(1) & "', " & _
                              "a240006 = '" & m_NEW(2) & "', " & _
                              "a240007 = '" & m_NEW(3) & "', " & _
                              "a240008 = '" & m_NEW(4) & "' " & _
                  "WHERE a240005 = '" & m_OLD(1) & "' AND " & _
                        "a240006 = '" & m_OLD(2) & "' AND " & _
                        "a240007 = '" & m_OLD(3) & "' AND " & _
                        "a240008 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   'Moidfy  by Amy 2018/07/24
   '判斷舊代表圖檔是否存在
   If ChkImgByteFile(m_OLD(1), m_OLD(2), m_OLD(3), m_OLD(4), intCount) = True Then
        bolDelFile = True
   End If
   
   '2009/3/27 ADD BY SONIA
   ShowStatus "變更卷和代表圖資料中, 請稍候 . . ."
   strSql = "SELECT COUNT(*) FROM ImgByteFile WHERE IBF01 = '" & m_NEW(1) & "' AND " & _
                                   "IBF02 = '" & m_NEW(2) & "' AND " & _
                                   "IBF03 = '" & m_NEW(3) & "' AND " & _
                                   "IBF04 = '" & m_NEW(4) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.Fields(0) > 0 Then
      '刪除舊代表圖檔
      If bolDelFile = True Then
        For i = 1 To intCount
              PUB_DelFtpFile2 m_OLD(1) & "-" & m_OLD(2) & "-" & m_OLD(3) & "-" & m_OLD(4) & "-" & i, , UCase("ImgByteFile")
        Next i
      End If
      'Memo 原刪除ImgByteFile 舊案號改至下方刪
   Else
'      strSql = "UPDATE ImgByteFile SET IBF01 = '" & m_NEW(1) & "', " & _
'                                      "IBF02 = '" & m_NEW(2) & "', " & _
'                                      "IBF03 = '" & m_NEW(3) & "', " & _
'                                      "IBF04 = '" & m_NEW(4) & "' " & _
'                                "WHERE IBF01 = '" & m_OLD(1) & "' AND " & _
'                                      "IBF02 = '" & m_OLD(2) & "' AND " & _
'                                      "IBF03 = '" & m_OLD(3) & "' AND " & _
'                                      "IBF04 = '" & m_OLD(4) & "' "
'      cnnConnection.Execute strSql
      '複製舊代表圖至新案中
      If PUB_CopyImgFile(m_OLD, m_NEW) = True Then
        '刪除舊代表圖檔
        For i = 1 To intCount
            PUB_DelFtpFile2 m_OLD(1) & "-" & m_OLD(2) & "-" & m_OLD(3) & "-" & m_OLD(4) & "-" & i, , UCase("ImgByteFile")
        Next i
      End If
   End If
   If bolDelFile = True Then
        strSql = "DELETE ImgByteFile WHERE IBF01 = '" & m_OLD(1) & "' AND " & _
                                      "IBF02 = '" & m_OLD(2) & "' AND " & _
                                      "IBF03 = '" & m_OLD(3) & "' AND " & _
                                      "IBF04 = '" & m_OLD(4) & "' "
        cnnConnection.Execute strSql
   End If
   'end 2018/07/24
   rsTmp.Close
   
   ShowStatus "變更商品及服務資料中, 請稍候 . . ."
   strSql = "SELECT COUNT(*) FROM TMGoods WHERE TG01 = '" & m_NEW(1) & "' AND " & _
                                   "TG02 = '" & m_NEW(2) & "' AND " & _
                                   "TG03 = '" & m_NEW(3) & "' AND " & _
                                   "TG04 = '" & m_NEW(4) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.Fields(0) > 0 Then
      strSql = "DELETE TMGoods WHERE TG01 = '" & m_OLD(1) & "' AND " & _
                                      "TG02 = '" & m_OLD(2) & "' AND " & _
                                      "TG03 = '" & m_OLD(3) & "' AND " & _
                                      "TG04 = '" & m_OLD(4) & "' "
      cnnConnection.Execute strSql
   Else
      strSql = "UPDATE TMGoods SET TG01 = '" & m_NEW(1) & "', " & _
                                  "TG02 = '" & m_NEW(2) & "', " & _
                                  "TG03 = '" & m_NEW(3) & "', " & _
                                  "TG04 = '" & m_NEW(4) & "' " & _
                            "WHERE TG01 = '" & m_OLD(1) & "' AND " & _
                                  "TG02 = '" & m_OLD(2) & "' AND " & _
                                  "TG03 = '" & m_OLD(3) & "' AND " & _
                                  "TG04 = '" & m_OLD(4) & "' "
      cnnConnection.Execute strSql
   End If
   rsTmp.Close
   
   ShowStatus "變更國外代理人信件記錄資料中, 請稍候 . . ."
   strSql = "UPDATE FagentMail SET FM03 = '" & m_NEW(1) & "', " & _
                                  "FM04 = '" & m_NEW(2) & "', " & _
                                  "FM05 = '" & m_NEW(3) & "', " & _
                                  "FM06 = '" & m_NEW(4) & "' " & _
                            "WHERE FM03 = '" & m_OLD(1) & "' AND " & _
                                  "FM04 = '" & m_OLD(2) & "' AND " & _
                                  "FM05 = '" & m_OLD(3) & "' AND " & _
                                  "FM06 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   ShowStatus "變更維護資料記錄中, 請稍候 . . ."
   strSql = "UPDATE DML_Log SET DL01 = '" & m_NEW(1) & "', " & _
                               "DL02 = '" & m_NEW(2) & "', " & _
                               "DL03 = '" & m_NEW(3) & "', " & _
                               "DL04 = '" & m_NEW(4) & "' " & _
                         "WHERE DL01 = '" & m_OLD(1) & "' AND " & _
                               "DL02 = '" & m_OLD(2) & "' AND " & _
                               "DL03 = '" & m_OLD(3) & "' AND " & _
                               "DL04 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   '2009/3/27 END
   
   '2010/8/6 add by sonia
   ShowStatus "變更重新委任客戶資料中, 請稍候 . . ."
   strSql = "UPDATE LINREASIGNREC SET LR02 = '" & m_NEW(1) & "', " & _
                                     "LR03 = '" & m_NEW(2) & "', " & _
                                     "LR04 = '" & m_NEW(3) & "', " & _
                                     "LR05 = '" & m_NEW(4) & "' " & _
                               "WHERE LR02 = '" & m_OLD(1) & "' AND " & _
                                     "LR03 = '" & m_OLD(2) & "' AND " & _
                                     "LR04 = '" & m_OLD(3) & "' AND " & _
                                     "LR05 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   '2010/8/6 end
      
   'add by sonia 2014/6/30
   ShowStatus "變更定稿資料中, 請稍候 . . ."
   strSql = "UPDATE LetterDemand SET LD05 = '" & m_NEW(1) & "', " & _
                                    "LD06 = '" & m_NEW(2) & "', " & _
                                    "LD07 = '" & m_NEW(3) & "', " & _
                                    "LD08 = '" & m_NEW(4) & "' " & _
                              "WHERE LD05 = '" & m_OLD(1) & "' AND " & _
                                    "LD06 = '" & m_OLD(2) & "' AND " & _
                                    "LD07 = '" & m_OLD(3) & "' AND " & _
                                    "LD08 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   '2014/6/30 end
   
   'add by sonia 2015/6/25
   ShowStatus "變更分割案件關係資料中, 請稍候 . . ."
   strSql = "UPDATE DIVISIONCASE SET DC01 = '" & m_NEW(1) & "', " & _
                                    "DC02 = '" & m_NEW(2) & "', " & _
                                    "DC03 = '" & m_NEW(3) & "', " & _
                                    "DC04 = '" & m_NEW(4) & "' " & _
                              "WHERE DC01 = '" & m_OLD(1) & "' AND " & _
                                    "DC02 = '" & m_OLD(2) & "' AND " & _
                                    "DC03 = '" & m_OLD(3) & "' AND " & _
                                    "DC04 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   strSql = "UPDATE DIVISIONCASE SET DC05 = '" & m_NEW(1) & "', " & _
                                    "DC06 = '" & m_NEW(2) & "', " & _
                                    "DC07 = '" & m_NEW(3) & "', " & _
                                    "DC08 = '" & m_NEW(4) & "' " & _
                              "WHERE DC05 = '" & m_OLD(1) & "' AND " & _
                                    "DC06 = '" & m_OLD(2) & "' AND " & _
                                    "DC07 = '" & m_OLD(3) & "' AND " & _
                                    "DC08 = '" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   '2015/6/25 end
      
   'Added by Lydia 2016/11/30 各項備註檔
   strSql = "UPDATE INSTRUCTIONS SET ITS01='" & Pub_GetITS01Type(m_NEW(1) & m_NEW(2) & m_NEW(3) & m_NEW(4)) & "'," & _
                                    "ITS02='" & m_NEW(1) & m_NEW(2) & m_NEW(3) & m_NEW(4) & "' " & _
                              "WHERE ITS01='" & Pub_GetITS01Type(m_OLD(1) & m_OLD(2) & m_OLD(3) & m_OLD(4)) & "' AND " & _
                                    "ITS02='" & m_OLD(1) & m_OLD(2) & m_OLD(3) & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   'end 2016/11/30
   
   'Add By Sindy 2019/12/30
   '改系統類別時，畫面增加 '智權人員'欄且檢查一定要輸入。
   '存檔時同時新增B類993進度，收文日發文日都存系統日，
   '承辦人：新系統類別P存P1001、FCP存F4102、T存P2001、FCT存F4103。
   If textOLD01 <> textNEW01 And textOLD01 <> "" And textNEW01 <> "" Then
      stCP09 = AutoNo("B", 6)
      'stCP13 = PUB_GetAKindSalesNo(pCP01, pCP02, pCP03, pCP04)
      stCP12 = GetSalesArea(textCP13)
      If m_NEW(1) = "P" Then
         stCP14 = "P1001"
      ElseIf m_NEW(1) = "FCP" Then
         stCP14 = "F4102"
      ElseIf m_NEW(1) = "T" Then
         stCP14 = "P2001"
      ElseIf m_NEW(1) = "FCT" Then
         stCP14 = "F4103"
      End If
      strSql = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
         ",cp12,cp13,cp14,cp20,cp26,cp27,cp32) values ('" & m_NEW(1) & "'" & _
         ",'" & m_NEW(2) & "','" & m_NEW(3) & "','" & m_NEW(4) & "'," & strSrvDate(1) & _
         ",'" & stCP09 & "','993','" & stCP12 & "'" & _
         ",'" & textCP13 & "','" & stCP14 & "','N','N'," & strSrvDate(1) & ",'N')"
      cnnConnection.Execute strSql, intI
   End If
   '2019/12/30 END
   
   'Added by Lydia 2018/11/28 更新備註檔
   '下一程序固定備註(NpMemo)
   strSql = "UPDATE NPMEMO SET NM03='" & m_NEW(1) & m_NEW(2) & m_NEW(3) & m_NEW(4) & "' WHERE NM03='" & m_OLD(1) & m_OLD(2) & m_OLD(3) & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   '核准函輸入備註(ApprovalMemo2)
   strSql = "UPDATE APPROVALMEMO2 SET AM03='" & m_NEW(1) & m_NEW(2) & m_NEW(3) & m_NEW(4) & "' WHERE AM03='" & m_OLD(1) & m_OLD(2) & m_OLD(3) & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   '核駁及審查意見通知函備註(IncomMemo)
   strSql = "UPDATE INCOMMEMO SET IM03='" & m_NEW(1) & m_NEW(2) & m_NEW(3) & m_NEW(4) & "' WHERE IM03='" & m_OLD(1) & m_OLD(2) & m_OLD(3) & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   '請款函預設備註維護檔(DebitNotePS)
   strSql = "UPDATE DEBITNOTEPS SET DNPS03='" & m_NEW(1) & m_NEW(2) & m_NEW(3) & m_NEW(4) & "' WHERE DNPS03='" & m_OLD(1) & m_OLD(2) & m_OLD(3) & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   'end 2018/11/28
   'Added by Lydia 2019/03/11 FCP承辦單設定維護(FcpEMPbill)
   strSql = "UPDATE FCPEMPBILL SET FEB03='" & m_NEW(1) & m_NEW(2) & m_NEW(3) & m_NEW(4) & "' WHERE FEB03='" & m_OLD(1) & m_OLD(2) & m_OLD(3) & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   'Added by Lydia 2019/03/11 通知告准加註(ApprovalPS)
   strSql = "UPDATE APPROVALPS SET APS03='" & m_NEW(1) & m_NEW(2) & m_NEW(3) & m_NEW(4) & "' WHERE APS03='" & m_OLD(1) & m_OLD(2) & m_OLD(3) & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   'Added by Lydia 2023/06/28 國外部：藥證號數對照檔 MedicineCodeMap
   strSql = "UPDATE MedicineCodeMap SET MCM02='" & m_NEW(1) & "', MCM03='" & m_NEW(2) & "', MCM04='" & m_NEW(3) & "', MCM05='" & m_NEW(4) & "' WHERE MCM02='" & m_OLD(1) & "'AND MCM03='" & m_OLD(2) & "'AND MCM04='" & m_OLD(3) & "'AND MCM05='" & m_OLD(4) & "' "
   cnnConnection.Execute strSql
   
   '2011/6/15 自上方移下來
   If textDelete = "Y" Then
      ' 刪除基本檔
      '2011/3/30 MODIFY BY SONIA 併入本所案號之備註加在該欄之前面並加註日期
      Select Case m_OLD(1)
         ' 商標基本檔
         Case "T", "TF", "CFT", "FCT":
            ShowStatus "刪除商標基本檔中, 請稍候 . . ."
            strSql = "DELETE TRADEMARK " & _
                     "WHERE TM01 = '" & m_OLD(1) & "' AND " & _
                           "TM02 = '" & m_OLD(2) & "' AND " & _
                           "TM03 = '" & m_OLD(3) & "' AND " & _
                           "TM04 = '" & m_OLD(4) & "' "
            Pub_SeekTbLog strSql   '2009/3/27 ADD BY SONIA 新增維護記錄檔
            cnnConnection.Execute strSql
            '2005/8/30 ADD BY SONIA 更新新案號之案件備註
            strSql = "UPDATE TRADEMARK SET TM58 = '" & ChangeTStringToTDateString(strSrvDate(2)) & "由" & m_OLD(1) & "-" & m_OLD(2) & "-" & m_OLD(3) & "-" & m_OLD(4) & "併入;' || TM58 " & _
                     "WHERE TM01 = '" & m_NEW(1) & "' AND " & _
                           "TM02 = '" & m_NEW(2) & "' AND " & _
                           "TM03 = '" & m_NEW(3) & "' AND " & _
                           "TM04 = '" & m_NEW(4) & "' "
            Pub_SeekTbLog strSql   '2009/3/27 ADD BY SONIA 新增維護記錄檔
            cnnConnection.Execute strSql
            '2005/8/30 END
         ' 專利基本檔
         Case "P", "CFP", "FCP":
            ShowStatus "刪除專利基本檔中, 請稍候 . . ."
            strSql = "DELETE PATENT " & _
                     "WHERE PA01 = '" & m_OLD(1) & "' AND " & _
                           "PA02 = '" & m_OLD(2) & "' AND " & _
                           "PA03 = '" & m_OLD(3) & "' AND " & _
                           "PA04 = '" & m_OLD(4) & "' "
            Pub_SeekTbLog strSql   '2009/3/27 ADD BY SONIA 新增維護記錄檔
            cnnConnection.Execute strSql
            '2005/8/30 ADD BY SONIA 更新新案號之案件備註
            strSql = "UPDATE PATENT SET PA91 = '" & ChangeTStringToTDateString(strSrvDate(2)) & "由" & m_OLD(1) & "-" & m_OLD(2) & "-" & m_OLD(3) & "-" & m_OLD(4) & "併入;' || PA91 " & _
                     "WHERE PA01 = '" & m_NEW(1) & "' AND " & _
                           "PA02 = '" & m_NEW(2) & "' AND " & _
                           "PA03 = '" & m_NEW(3) & "' AND " & _
                           "PA04 = '" & m_NEW(4) & "' "
            Pub_SeekTbLog strSql   '2009/3/27 ADD BY SONIA 新增維護記錄檔
            cnnConnection.Execute strSql
            '2005/8/30 END
         ' 法務基本檔
         Case "L", "CFL", "FCL":
            ShowStatus "刪除法務基本檔中, 請稍候 . . ."
            strSql = "DELETE LAWCASE " & _
                     "WHERE LC01 = '" & m_OLD(1) & "' AND " & _
                           "LC02 = '" & m_OLD(2) & "' AND " & _
                           "LC03 = '" & m_OLD(3) & "' AND " & _
                           "LC04 = '" & m_OLD(4) & "' "
            Pub_SeekTbLog strSql   '2009/3/27 ADD BY SONIA 新增維護記錄檔
            cnnConnection.Execute strSql
            '2005/8/30 ADD BY SONIA 更新新案號之案件備註
            strSql = "UPDATE LAWCASE SET LC27 = '" & ChangeTStringToTDateString(strSrvDate(2)) & "由" & m_OLD(1) & "-" & m_OLD(2) & "-" & m_OLD(3) & "-" & m_OLD(4) & "併入;' || LC27 " & _
                     "WHERE LC01 = '" & m_NEW(1) & "' AND " & _
                           "LC02 = '" & m_NEW(2) & "' AND " & _
                           "LC03 = '" & m_NEW(3) & "' AND " & _
                           "LC04 = '" & m_NEW(4) & "' "
            Pub_SeekTbLog strSql   '2009/3/27 ADD BY SONIA 新增維護記錄檔
            cnnConnection.Execute strSql
            '2005/8/30 END
         ' 顧問案件基本檔
         Case "LA":
            ShowStatus "刪除顧問基本檔中, 請稍候 . . ."
            strSql = "DELETE HIRECASE " & _
                     "WHERE HC01 = '" & m_OLD(1) & "' AND " & _
                           "HC02 = '" & m_OLD(2) & "' AND " & _
                           "HC03 = '" & m_OLD(3) & "' AND " & _
                           "HC04 = '" & m_OLD(4) & "' "
            Pub_SeekTbLog strSql   '2009/3/27 ADD BY SONIA 新增維護記錄檔
            cnnConnection.Execute strSql
            '2005/8/30 ADD BY SONIA 更新新案號之案件備註
            strSql = "UPDATE HIRECASE SET HC12 = '" & ChangeTStringToTDateString(strSrvDate(2)) & "由" & m_OLD(1) & "-" & m_OLD(2) & "-" & m_OLD(3) & "-" & m_OLD(4) & "併入;' || HC12 " & _
                     "WHERE HC01 = '" & m_NEW(1) & "' AND " & _
                           "HC02 = '" & m_NEW(2) & "' AND " & _
                           "HC03 = '" & m_NEW(3) & "' AND " & _
                           "HC04 = '" & m_NEW(4) & "' "
            Pub_SeekTbLog strSql   '2009/3/27 ADD BY SONIA 新增維護記錄檔
            cnnConnection.Execute strSql
            '2005/8/30 END
         ' 服務業務基本檔
         Case Else:
            ShowStatus "刪除服務業務基本檔中, 請稍候 . . ."
            strSql = "DELETE SERVICEPRACTICE " & _
                     "WHERE SP01 = '" & m_OLD(1) & "' AND " & _
                           "SP02 = '" & m_OLD(2) & "' AND " & _
                           "SP03 = '" & m_OLD(3) & "' AND " & _
                           "SP04 = '" & m_OLD(4) & "' "
            Pub_SeekTbLog strSql   '2009/3/27 ADD BY SONIA 新增維護記錄檔
            cnnConnection.Execute strSql
            '2005/8/30 ADD BY SONIA 更新新案號之案件備註
            strSql = "UPDATE SERVICEPRACTICE SET SP18 = '" & ChangeTStringToTDateString(strSrvDate(2)) & "由" & m_OLD(1) & "-" & m_OLD(2) & "-" & m_OLD(3) & "-" & m_OLD(4) & "併入;' || SP18 " & _
                     "WHERE SP01 = '" & m_NEW(1) & "' AND " & _
                           "SP02 = '" & m_NEW(2) & "' AND " & _
                           "SP03 = '" & m_NEW(3) & "' AND " & _
                           "SP04 = '" & m_NEW(4) & "' "
            Pub_SeekTbLog strSql   '2009/3/27 ADD BY SONIA 新增維護記錄檔
            cnnConnection.Execute strSql
            '2005/8/30 END
      End Select
   End If
   '2011/6/15 end
   ShowStatus Empty

   'Add By Sindy 2019/4/11 檢查卷宗區掛個案未分至文號的電子檔,也要一併移檔
   '新案號
   strNewFileName = m_NEW(1) & m_NEW(2)
   If m_NEW(3) & m_NEW(4) <> "000" Then
      strNewFileName = strNewFileName & "-" & m_NEW(3)
   End If
   If m_NEW(4) <> "00" Then
      strNewFileName = strNewFileName & "-" & m_NEW(4)
   End If
   '舊案號
   strOldFileName = m_OLD(1) & m_OLD(2)
   If m_OLD(3) & m_OLD(4) <> "000" Then
      strOldFileName = strOldFileName & "-" & m_OLD(3)
   End If
   If m_OLD(4) <> "00" Then
      strOldFileName = strOldFileName & "-" & m_OLD(4)
   End If
   strSql = "update casepaperpdf" & _
            " set cpp02=replace(cpp02,'" & strOldFileName & "','" & strNewFileName & "'),cpp13=0" & _
            " WHERE cpp01='" & m_OLD(1) & m_OLD(2) & m_OLD(3) & m_OLD(4) & "' and cpp02<>replace(cpp02,'" & strOldFileName & "','" & strNewFileName & "')"
   Pub_SeekTbLog strSql
   strSql = "begin user_data.user_enabled:=1; " & strSql & "; end; " 'Add by Sindy 2022/5/26
   cnnConnection.Execute strSql
   '2019/4/11 END
'Add By Cheng 2002/11/08
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False
End Function

'Private Function IsDBExist(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String, ByVal strKEY04 As String) As Boolean
'   Dim strSQL As String
'   Dim rsTmp As New ADODB.Recordset
'
'   If IsEmptyText(strKEY01) = True Then: GoTo EXITSUB
'   If IsEmptyText(strKEY02) = True Then: GoTo EXITSUB
'   strKEY02 = Mid(strKEY02, 1, 6)
'   If IsEmptyText(strKEY03) = True Then: strKEY03 = "0"
'   If IsEmptyText(strKEY04) = True Then: strKEY04 = "00"
'
'   IsDBExist = False
'   Select Case strKEY01
'      ' 讀取商標基本檔
'      Case "T", "TF", "CFT", "FCT":
'         strSQL = "SELECT * FROM TRADEMARK " & _
'                  "WHERE TM01 = '" & strKEY01 & "' AND " & _
'                        "TM02 = '" & strKEY02 & "' AND " & _
'                        "TM03 = '" & strKEY03 & "' AND " & _
'                        "TM04 = '" & strKEY04 & "' "
'      ' 讀取專利基本檔
'      Case "P", "CFP", "FCP":
'         strSQL = "SELECT * FROM PATENT " & _
'                  "WHERE PA01 = '" & strKEY01 & "' AND " & _
'                        "PA02 = '" & strKEY02 & "' AND " & _
'                        "PA03 = '" & strKEY03 & "' AND " & _
'                        "PA04 = '" & strKEY04 & "' "
'      ' 讀取法務基本檔
'      Case "L", "CFL", "FCL":
'         strSQL = "SELECT * FROM LAWCASE " & _
'                  "WHERE LC01 = '" & strKEY01 & "' AND " & _
'                        "LC02 = '" & strKEY02 & "' AND " & _
'                        "LC03 = '" & strKEY03 & "' AND " & _
'                        "LC04 = '" & strKEY04 & "' "
'      ' 讀取顧問案件基本檔
'      Case "LA":
'         strSQL = "SELECT * FROM HIRECASE " & _
'                  "WHERE HC01 = '" & strKEY01 & "' AND " & _
'                        "HC02 = '" & strKEY02 & "' AND " & _
'                        "HC03 = '" & strKEY03 & "' AND " & _
'                        "HC04 = '" & strKEY04 & "' "
'      ' 讀取服務業務基本檔
'      Case Else:
'         strSQL = "SELECT * FROM SERVICEPRACTICE " & _
'                  "WHERE SP01 = '" & strKEY01 & "' AND " & _
'                        "SP02 = '" & strKEY01 & "' AND " & _
'                        "SP03 = '" & strKEY01 & "' AND " & _
'                        "SP04 = '" & strKEY01 & "' "
'   End Select
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      IsDBExist = True
'   End If
'   rsTmp.Close
''EXITSUB:
'   Set rsTmp = Nothing
'End Function

Private Function ClearName()
   textOCName(0) = Empty
   textOEName(0) = Empty
   textOJName(0) = Empty
   textOCName(1) = Empty
   textOEName(1) = Empty
   textOJName(1) = Empty
   textOMemo = Empty 'Added by Lydia 2025/10/15
End Function

' 取得案件的名稱
Private Function QueryName(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String, ByVal strKEY04 As String) As Boolean
   Dim bQuery As Boolean
   
   QueryName = False
   If IsEmptyText(strKEY01) = True Then: GoTo EXITSUB
   If IsEmptyText(strKEY02) = True Then: GoTo EXITSUB
   strKEY02 = Mid(strKEY02, 1, 6)
   If IsEmptyText(strKEY03) = True Then: strKEY03 = "0"
   If IsEmptyText(strKEY04) = True Then: strKEY04 = "00"
   
   ' 更新基本檔
   Select Case strKEY01
      ' 商標基本檔
      Case "T", "TF", "CFT", "FCT":
         bQuery = QueryTradeMark(strKEY01, strKEY02, strKEY03, strKEY04)
      ' 專利基本檔
      Case "P", "CFP", "FCP":
         bQuery = QueryPatent(strKEY01, strKEY02, strKEY03, strKEY04)
      ' 法務基本檔
      'modify by sonia 2023/9/19 加入LIN及ACS
      Case "L", "CFL", "FCL", "LIN", "ACS":
         bQuery = QueryLawCase(strKEY01, strKEY02, strKEY03, strKEY04)
      ' 顧問案件基本檔
      Case "LA":
         bQuery = QueryHireCase(strKEY01, strKEY02, strKEY03, strKEY04)
      ' 服務業務基本檔
      Case Else:
         bQuery = QueryServicePractice(strKEY01, strKEY02, strKEY03, strKEY04)
   End Select
   
   QueryName = bQuery
   
   'Added by Lydia 2025/10/15
   If bQuery = True And Trim(textOMemo) <> "" Then
       MsgBox "請注意原本所案號之案件備註欄是否曾經有改本所案號的記錄，若欲改回原系統類別則請改回原案號 !", vbInformation + vbOKOnly
   End If
   'end 2025/10/15
   
EXITSUB:
End Function

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18       維護資料記錄中
   'Modify By Sindy 2024/8/30 mark
   'MsgBox "維護記錄資料，請人工更新! ", vbCritical + vbOKOnly, "檢核資料"
   Set frm12040126 = Nothing
End Sub

'Add By Sindy 2019/12/27
Private Sub textCP13_GotFocus()
   InverseTextBox textCP13
End Sub
Private Sub textCP13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
' 智權人員代號
Private Sub textCP13_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse

   Cancel = False
   textCP13_2 = Empty
   If IsEmptyText(textCP13) = False Then
      textCP13_2 = GetStaffName(textCP13, True)
      If IsEmptyText(textCP13_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "智權人員代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP13_GotFocus
      End If
   Else
      '改系統類別時，畫面增加的'智權人員'欄,檢查一定要輸入
      If textOLD01 <> textNEW01 And textOLD01 <> "" And textNEW01 <> "" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "智權人員不可空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP13_GotFocus
      End If
   End If
   If Cancel = False Then
      textCP13.Tag = textCP13
   End If
End Sub
'2019/12/27 END

Private Sub TextDelete_GotFocus()
   InverseTextBox textDelete
End Sub

Private Sub textDelete_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub
'92.8.11 ADD BY SONIA
Private Sub textDelete_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   If textOCName(1) = "" And textOEName(1) = "" And textOJName(1) = "" And textDelete = "Y" Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "新本所案號不存在, 不可刪除原案件基本資料 !!"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      TextDelete_GotFocus
      Exit Sub
   End If
End Sub
'92.8.11 END
Private Sub textNEW02_Validate(Cancel As Boolean)
   Dim bQuery As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If Len(textNEW02) < 5 Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "新本所案號的流水號不正確"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textNEW02_GotFocus
      Exit Sub
   End If

End Sub

Private Sub textNEW03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textOLD01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textOLD01_Validate(Cancel As Boolean)
   Dim bQuery As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textOLD01) = False Then
      If textOLD01 = "TF" Then
         textOLD02.MaxLength = 5
         EnableTextBox textOLD02_2, True
         textOLD02_2.Visible = True
      Else
         textOLD02.MaxLength = 6
         textOLD02_2 = Empty
         EnableTextBox textOLD02_2, False
         textOLD02_2.Visible = False
      End If
      'If textOLD01 = "TF" Then
      '   If IsEmptyText(textOLD02) = False And IsEmptyText(textOLD02_2) = False And IsEmptyText(textOLD03) = False And IsEmptyText(textOLD04) = False Then
      '      bQuery = QueryName(textOLD01, textOLD02 & textOLD02_2, textOLD03, textOLD04)
      '   End If
      'Else
      '   If IsEmptyText(textOLD02) = False And IsEmptyText(textOLD03) = False And IsEmptyText(textOLD04) = False Then
      '      bQuery = QueryName(textOLD01, textOLD02, textOLD03, textOLD04)
      '   End If
      'End If
   Else
      textOLD02.MaxLength = 6
      textOLD02_2 = Empty
      EnableTextBox textOLD02_2, False
      textOLD02_2.Visible = False
   End If
   
End Sub

' 本所案號第二欄後方的欄位
Private Sub textOLD02_Validate(Cancel As Boolean)
   Dim bQuery As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'If IsEmptyText(textOLD02) = False Then
   '   If textOLD01 = "TF" Then
   '      If IsEmptyText(textOLD01) = False And IsEmptyText(textOLD02_2) = False And IsEmptyText(textOLD03) = False And IsEmptyText(textOLD04) = False Then
   '         bQuery = QueryName(textOLD01, textOLD02 & textOLD02_2, textOLD03, textOLD04)
   '      End If
   '   Else
   '      If IsEmptyText(textOLD01) = False And IsEmptyText(textOLD03) = False And IsEmptyText(textOLD04) = False Then
   '         bQuery = QueryName(textOLD01, textOLD02, textOLD03, textOLD04)
   '      End If
   '   End If
   'End If
      
End Sub

' 本所案號第二欄後方的欄位
Private Sub textOLD02_2_Validate(Cancel As Boolean)
   Dim bQuery As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If textOLD01 = "TF" Then
      textOLD02_2 = textOLD02_2 & String(1 - Len(textOLD02_2), "0")
   Else
      textOLD02_2 = Empty
   End If

   'If IsEmptyText(textOLD02_2) = False Then
   '   If textOLD01 = "TF" Then
   '      If IsEmptyText(textOLD01) = False And IsEmptyText(textOLD02) = False And IsEmptyText(textOLD03) = False And IsEmptyText(textOLD04) = False Then
   '         bQuery = QueryName(textOLD01, textOLD02 & textOLD02_2, textOLD03, textOLD04)
   '      End If
   '   Else
   '      If IsEmptyText(textOLD01) = False And IsEmptyText(textOLD03) = False And IsEmptyText(textOLD04) = False Then
   '         bQuery = QueryName(textOLD01, textOLD02, textOLD03, textOLD04)
   '      End If
   '   End If
   'End If
      
End Sub

Private Sub textOLD03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 本所案號第三欄
Private Sub textOLD03_Validate(Cancel As Boolean)
   Dim bQuery As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'textOLD03 = textOLD03 & String(1 - Len(textOLD03), "0")
   
   'If IsEmptyText(textOLD03) = False Then
      'If textOLD01 = "TF" Then
      '   If IsEmptyText(textOLD01) = False And IsEmptyText(textOLD02) = False And IsEmptyText(textOLD02_2) = False And IsEmptyText(textOLD04) = False Then
      '      bQuery = QueryName(textOLD01, textOLD02 & textOLD02_2, textOLD03, textOLD04)
      '   End If
      'Else
      '   If IsEmptyText(textOLD01) = False And IsEmptyText(textOLD02) = False And IsEmptyText(textOLD04) = False Then
      '      bQuery = QueryName(textOLD01, textOLD02, textOLD03, textOLD04)
      '   End If
      'End If
   'End If
   
End Sub

Private Sub textOLD04_LostFocus()
   Dim bQuery As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   'textOLD04 = textOLD04 & String(2 - Len(textOLD04), "0")
   ' 清除案件名稱
   ClearName
   'If IsEmptyText(textOLD04) = False Then
      If textOLD01 = "TF" Then
         If IsEmptyText(textOLD01) = False And IsEmptyText(textOLD02) = False And IsEmptyText(textOLD02_2) = False Then
            bQuery = QueryName(textOLD01, textOLD02 & textOLD02_2, textOLD03, textOLD04)
            If bQuery = False Then
               strTit = "檢核資料"
               strMsg = "該筆案件不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textOLD01.SetFocus
            End If
         End If
      Else
         If IsEmptyText(textOLD01) = False And IsEmptyText(textOLD02) = False Then
            bQuery = QueryName(textOLD01, textOLD02, textOLD03, textOLD04)
            If bQuery = False Then
               strTit = "檢核資料"
               strMsg = "該筆案件不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textOLD01.SetFocus
            End If
         End If
      End If
   'End If
End Sub

' 本所案號第四欄
'Private Sub textOLD04_Validate(Cancel As Boolean)
'   Dim bQuery As Boolean
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'
'   Cancel = False
'   'textOLD04 = textOLD04 & String(2 - Len(textOLD04), "0")
'   ' 清除案件名稱
'   ClearName
'   'If IsEmptyText(textOLD04) = False Then
'      If textOLD01 = "TF" Then
'         If IsEmptyText(textOLD01) = False And IsEmptyText(textOLD02) = False And IsEmptyText(textOLD02_2) = False Then
'            bQuery = QueryName(textOLD01, textOLD02 & textOLD02_2, textOLD03, textOLD04)
'            If bQuery = False Then
'               Cancel = True
'               strTit = "檢核資料"
'               strMsg = "該筆案件不存在"
'               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'               textOLD01_GotFocus
'            End If
'         End If
'      Else
'         If IsEmptyText(textOLD01) = False And IsEmptyText(textOLD02) = False Then
'            bQuery = QueryName(textOLD01, textOLD02, textOLD03, textOLD04)
'            If bQuery = False Then
'               Cancel = True
'               strTit = "檢核資料"
'               strMsg = "該筆案件不存在"
'               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'               textOLD01_GotFocus
'            End If
'         End If
'      End If
'   'End If
'End Sub

Private Sub textNEW01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textNEW01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textNEW01) = False Then
      If IsEmptyText(textOLD01) = False Then
         Select Case textOLD01
            ' 商標基本檔
            Case "T", "TF", "CFT", "FCT":
               Select Case textNEW01
                  Case "T", "TF", "CFT", "FCT":
                  Case Else:
                     Cancel = True
                     strTit = "檢核資料"
                     strMsg = "新本所案號的系統別不正確"
                     nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                     textNEW01_GotFocus
                     GoTo EXITSUB
               End Select
            ' 專利基本檔
            Case "P", "CFP", "FCP":
               Select Case textNEW01
                  Case "P", "CFP", "FCP":
                  Case Else:
                     Cancel = True
                     strTit = "檢核資料"
                     strMsg = "新本所案號的系統別不正確"
                     nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                     textNEW01_GotFocus
                     GoTo EXITSUB
               End Select
            ' 法務基本檔
            Case "L", "CFL", "FCL", "LIN":
               Select Case textNEW01
                  Case "L", "CFL", "FCL", "LIN":
                  Case Else:
                     Cancel = True
                     strTit = "檢核資料"
                     strMsg = "新本所案號的系統別不正確"
                     nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                     textNEW01_GotFocus
                     GoTo EXITSUB
               End Select
            ' 顧問案件基本檔
            Case "LA":
               Select Case textNEW01
                  Case "LA":
                  Case Else:
                     Cancel = True
                     strTit = "檢核資料"
                     strMsg = "新本所案號的系統別不正確"
                     nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                     textNEW01_GotFocus
                     GoTo EXITSUB
               End Select
            ' 服務業務基本檔
            Case Else:
               If textNEW01 <> textOLD01 Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "新本所案號的系統別不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textNEW01_GotFocus
                  GoTo EXITSUB
               End If
         End Select
      End If
      
      If textNEW01 = "TF" Then
         textNEW02.MaxLength = 5
         EnableTextBox textNEW02_2, True
         textNEW02_2.Visible = True
      Else
         textOLD02.MaxLength = 6
         textNEW02_2 = Empty
         EnableTextBox textNEW02_2, False
         textNEW02_2.Visible = False
      End If
   Else
      textOLD02.MaxLength = 6
      textNEW02_2 = Empty
      EnableTextBox textNEW02_2, False
      textNEW02_2.Visible = False
   End If
EXITSUB:
End Sub

Private Sub textNEW02_2_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If textNEW01 = "TF" Then
      textNEW02_2 = textNEW02_2 & String(1 - Len(textNEW02_2), "0")
   Else
      textNEW02_2 = Empty
   End If
End Sub

Private Sub textNEW03_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   'textNEW03 = textNEW03 & String(1 - Len(textNEW03), "0")
End Sub

Private Function IsDBExist(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String, ByVal strKEY04 As String, Optional ByRef strName As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strName = Empty
   textOCName(1) = "": textOEName(1) = "": textOJName(1) = ""
   If IsEmptyText(strKEY01) = True Then: GoTo EXITSUB
   If IsEmptyText(strKEY02) = True Then: GoTo EXITSUB
   If IsEmptyText(strKEY03) = True Then: strKEY03 = "0"
   If IsEmptyText(strKEY04) = True Then: strKEY04 = "00"
   
   IsDBExist = False
   Select Case strKEY01
      ' 讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         strSql = "SELECT * FROM TRADEMARK " & _
                  "WHERE TM01 = '" & strKEY01 & "' AND " & _
                        "TM02 = '" & strKEY02 & "' AND " & _
                        "TM03 = '" & strKEY03 & "' AND " & _
                        "TM04 = '" & strKEY04 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            IsDBExist = True
            If IsNull(rsTmp.Fields("TM05")) = False Then
               textOCName(1) = rsTmp.Fields("TM05")
            End If
            ' 案件名稱
            If IsNull(rsTmp.Fields("TM06")) = False Then
               textOEName(1) = rsTmp.Fields("TM06")
            End If
            ' 案件名稱
            If IsNull(rsTmp.Fields("TM07")) = False Then
               textOJName(1) = rsTmp.Fields("TM07")
            End If
         End If
         rsTmp.Close
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         strSql = "SELECT * FROM PATENT " & _
                  "WHERE PA01 = '" & strKEY01 & "' AND " & _
                        "PA02 = '" & strKEY02 & "' AND " & _
                        "PA03 = '" & strKEY03 & "' AND " & _
                        "PA04 = '" & strKEY04 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            IsDBExist = True
            If IsNull(rsTmp.Fields("PA05")) = False Then
               textOCName(1) = rsTmp.Fields("PA05")
            End If
            ' 案件名稱
            If IsNull(rsTmp.Fields("PA06")) = False Then
               textOEName(1) = rsTmp.Fields("PA06")
            End If
            ' 案件名稱
            If IsNull(rsTmp.Fields("PA07")) = False Then
               textOJName(1) = rsTmp.Fields("PA07")
            End If
         End If
         rsTmp.Close
      ' 讀取法務基本檔
      Case "L", "CFL", "FCL":
         strSql = "SELECT * FROM LAWCASE " & _
                  "WHERE LC01 = '" & strKEY01 & "' AND " & _
                        "LC02 = '" & strKEY02 & "' AND " & _
                        "LC03 = '" & strKEY03 & "' AND " & _
                        "LC04 = '" & strKEY04 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            IsDBExist = True
            If IsNull(rsTmp.Fields("LC05")) = False Then
               textOCName(1) = rsTmp.Fields("LC05")
            End If
            ' 案件名稱
            If IsNull(rsTmp.Fields("LC06")) = False Then
               textOEName(1) = rsTmp.Fields("LC06")
            End If
            ' 案件名稱
            If IsNull(rsTmp.Fields("LC07")) = False Then
               textOJName(1) = rsTmp.Fields("LC07")
            End If
         End If
         rsTmp.Close
      ' 讀取顧問案件基本檔
      Case "LA":
         strSql = "SELECT * FROM HIRECASE " & _
                  "WHERE HC01 = '" & strKEY01 & "' AND " & _
                        "HC02 = '" & strKEY02 & "' AND " & _
                        "HC03 = '" & strKEY03 & "' AND " & _
                        "HC04 = '" & strKEY04 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            IsDBExist = True
            ' 案件名稱
            If IsNull(rsTmp.Fields("HC06")) = False Then
               textOCName(1) = rsTmp.Fields("HC06")
            End If
            textOEName(1) = ""
            textOJName(1) = ""
         End If
         rsTmp.Close
      ' 讀取服務業務基本檔
      Case Else:
         strSql = "SELECT * FROM SERVICEPRACTICE " & _
                  "WHERE SP01 = '" & strKEY01 & "' AND " & _
                        "SP02 = '" & strKEY02 & "' AND " & _
                        "SP03 = '" & strKEY03 & "' AND " & _
                        "SP04 = '" & strKEY04 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            IsDBExist = True
            If IsNull(rsTmp.Fields("SP05")) = False Then
               textOCName(1) = rsTmp.Fields("SP05")
            End If
            ' 案件名稱
            If IsNull(rsTmp.Fields("SP06")) = False Then
               textOEName(1) = rsTmp.Fields("SP06")
            End If
            ' 案件名稱
            If IsNull(rsTmp.Fields("SP07")) = False Then
               textOJName(1) = rsTmp.Fields("SP07")
            End If
         End If
         rsTmp.Close
   End Select
   
EXITSUB:
   Set rsTmp = Nothing
End Function

Private Sub textNEW04_LostFocus()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strKEY01 As String
   Dim strKEY02 As String
   Dim strKEY03 As String
   Dim strKEY04 As String
   Dim strName As String
   
   If IsEmptyText(textNEW01) = False And IsEmptyText(textNEW02) = False Then
      If textNEW01 = "TF" And IsEmptyText(textNEW02_2) = True Then
         GoTo EXITSUB
      End If
      strKEY01 = textNEW01
      strKEY02 = textNEW02
      If strKEY01 = "TF" Then strKEY02 = strKEY02 & textNEW02_2
      strKEY03 = textNEW03
      If IsEmptyText(strKEY03) Then: strKEY03 = "0"
      strKEY04 = textNEW04
      If IsEmptyText(strKEY04) Then: strKEY04 = "00"
      If IsDBExist(strKEY01, strKEY02, strKEY03, strKEY04, strName) = True Then
         strMsg = "本所案號" & strKEY01 & "-" & strKEY02 & "-" & strKEY03 & "-" & strKEY04 & "已存在"
         'If IsEmptyText(strName) = False Then
         '   strMsg = strMsg & ", 案件名稱 : " & strName
         'End If
         If MsgBox(strMsg, vbYesNo + vbCritical) = vbNo Then
            textNEW01.SetFocus
         End If
      End If
   End If
EXITSUB:
End Sub

'Private Sub textNEW04_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Dim strKEY01 As String
'   Dim strKEY02 As String
'   Dim strKEY03 As String
'   Dim strKEY04 As String
'   Dim strName As String
'   Dim strMsg As String
'
'   If IsEmptyText(textNEW01) = False And IsEmptyText(textNEW02) = False Then
'      If textNEW01 = "TF" And IsEmptyText(textNEW02_2) = True Then
'         GoTo EXITSUB
'      End If
'      strKEY01 = textNEW01
'      strKEY02 = textNEW02
'      If strKEY01 = "TF" Then strKEY02 = strKEY02 & textNEW02_2
'      streky03 = textNEW03
'      If IsEmptyText(strKEY03) Then: strKEY03 = "0"
'      strKEY04 = textNEW04
'      If IsEmptyText(strKEY04) Then: strKEY04 = "00"
'      If QueryRecord(strKEY01, strKEY02, strKEY03, strKEY04, strName) = True Then
'         strMsg = "本所案號" & strKEY01 & "-" & strKEY02 & "-" & strKEY03 & "-" & strKEY04 & "已存在"
'         If IsEmptyText(strName) = False Then
'            strMsg = strMsg & ", 案件名稱 : " & strName
'         End If
'         MsgBox strName, vbOKOnly + vbCritical, "檢核資料"
'         GoTo extisub
'      End If
'   End If
'
'   'textNEW04 = textNEW04 & String(2 - Len(textNEW04), "0")
'EXITSUB:
'End Sub

' 檢查輸入的資料是否完整
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strName As String
   CheckDataValid = False
   
   ' 原本所案號輸入不完整
   If textOLD01 = "TF" Then
      If IsEmptyText(textOLD01) Or IsEmptyText(textOLD02) Or IsEmptyText(textOLD02_2) Then
         strTit = "檢核資料"
         strMsg = "原本所案號輸入不完整"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textOLD01.SetFocus
         GoTo EXITSUB
      End If
   Else
      If IsEmptyText(textOLD01) Or IsEmptyText(textOLD02) Then
         strTit = "檢核資料"
         strMsg = "原本所案號輸入不完整"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textOLD01.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   ' 新本所案號輸入不完整
   If textNEW01 = "TF" Then
      If IsEmptyText(textNEW01) Or IsEmptyText(textNEW02) Or IsEmptyText(textNEW02_2) Then
         strTit = "檢核資料"
         strMsg = "新本所案號輸入不完整"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNEW01.SetFocus
         GoTo EXITSUB
      End If
   Else
      If IsEmptyText(textNEW01) Or IsEmptyText(textNEW02) Then
         strTit = "檢核資料"
         strMsg = "新本所案號輸入不完整"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNEW01.SetFocus
         GoTo EXITSUB
      End If
   End If

   ' 原本所案號
   If IsDBExist(textOLD01, textOLD02 & textOLD02_2, textOLD03 & String(1 - Len(textOLD03), "0"), textOLD04 & String(2 - Len(textOLD04), "0"), strName) = False Then
      strMsg = "本所案號" & textOLD01 & "-" & textOLD02 & textOLD02_2 & "-" & textOLD03 & String(1 - Len(textOLD03), "0") & "-" & textOLD04 & String(2 - Len(textOLD04), "0") & "不存在"
      strTit = "檢核資料"
      strMsg = "原本所案號不存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textOLD01.SetFocus
      GoTo EXITSUB
   End If
   
   ' 新本所案號
   If IsDBExist(textNEW01, textNEW02 & textNEW02_2, textNEW03 & String(1 - Len(textNEW03), "0"), textNEW04 & String(2 - Len(textNEW04), "0"), strName) = True Then
      strMsg = "本所案號" & textNEW01 & "-" & textNEW02 & textNEW02_2 & "-" & textNEW03 & String(1 - Len(textNEW03), "0") & "-" & textNEW04 & String(2 - Len(textNEW04), "0") & "已經存在資料庫中"
      'If IsEmptyText(strName) = False Then
      '   strMsg = strMsg & ", 案件名稱 : " & strName
      'End If
      If MsgBox(strMsg, vbYesNo + vbCritical) = vbNo Then
         textNEW01.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   Select Case textOLD01
      ' 商標基本檔
      Case "T", "TF", "CFT", "FCT":
         Select Case textNEW01
            Case "T", "TF", "CFT", "FCT":
            Case Else:
               strTit = "檢核資料"
               strMsg = "新本所案號的系統別不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textNEW01.SetFocus
               GoTo EXITSUB
         End Select
      ' 專利基本檔
      Case "P", "CFP", "FCP":
         Select Case textNEW01
            Case "P", "CFP", "FCP":
            Case Else:
               strTit = "檢核資料"
               strMsg = "新本所案號的系統別不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textNEW01.SetFocus
               GoTo EXITSUB
         End Select
      ' 法務基本檔
      Case "L", "CFL", "FCL", "LIN":
         Select Case textNEW01
            Case "L", "CFL", "FCL", "LIN":
            Case Else:
               strTit = "檢核資料"
               strMsg = "新本所案號的系統別不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textNEW01.SetFocus
               GoTo EXITSUB
         End Select
      ' 顧問案件基本檔
      Case "LA":
         Select Case textNEW01
            Case "LA":
            Case Else:
               strTit = "檢核資料"
               strMsg = "新本所案號的系統別不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textNEW01.SetFocus
               GoTo EXITSUB
         End Select
      ' 服務業務基本檔
      Case Else:
         If textNEW01 <> textOLD01 Then
            strTit = "檢核資料"
            strMsg = "新本所案號的系統別不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textNEW01.SetFocus
            GoTo EXITSUB
         End If
   End Select
   
   'Added by Lydia 2018/11/28 國外部：檢查備註檔在更新代號後，是否會造成重複
   If textOLD01 = "FCP" Or textOLD01 = "P" Then
        If frm12040125.CheckMemoDual(textOLD01 & textOLD02 & textOLD02_2 & textOLD03 & String(1 - Len(textOLD03), "0") & textOLD04 & String(2 - Len(textOLD04), "0") _
                   , textNEW01 & textNEW02 & textNEW02_2 & textNEW03 & String(1 - Len(textNEW03), "0") & textNEW04 & String(2 - Len(textNEW04), "0")) = True Then
             textOLD02.SetFocus
             GoTo EXITSUB
        End If
   End If
   'end 2018/11/28
   'Added by Lydia 2020/05/07 檢查各項指示檔在更新代號後，是否會造成重複
   If frm12040125.CheckInstructionsDual(textOLD01 & textOLD02 & textOLD02_2 & textOLD03 & String(1 - Len(textOLD03), "0") & textOLD04 & String(2 - Len(textOLD04), "0") _
                   , textNEW01 & textNEW02 & textNEW02_2 & textNEW03 & String(1 - Len(textNEW03), "0") & textNEW04 & String(2 - Len(textNEW04), "0")) = True Then
        textOLD02.SetFocus
        GoTo EXITSUB
   End If
   'end 2020/05/07
   
   'Added by Lydia 2023/06/28 國外部：藥證號數對照檔 MedicineCodeMap
   strSql = "select mcm01,count(*) cnt from (" & _
            " select mcm01,mcm02,mcm03,mcm04,mcm05,'1' as ord1 from medicinecodemap where mcm02='" & textNEW01 & "' and mcm03='" & textNEW02 & textNEW02_2 & "' and mcm04='" & textNEW03 & String(1 - Len(textNEW03), "0") & "' and mcm05='" & textNEW04 & String(2 - Len(textNEW04), "0") & "'" & _
            " union select mcm01,'" & textNEW01 & "' as mcm02,'" & textNEW02 & textNEW02_2 & "' as mcm03, '" & textNEW03 & String(1 - Len(textNEW03), "0") & "' as mcm04,'" & textNEW04 & String(2 - Len(textNEW04), "0") & "' as mcm05,'2' as ord1 from medicinecodemap where mcm02='" & textOLD01 & "' and mcm03='" & textOLD02 & textOLD02_2 & "' and mcm04='" & textOLD03 & String(1 - Len(textOLD03), "0") & "' and mcm05='" & textOLD04 & String(2 - Len(textOLD04), "0") & "'" & _
            " ) group by mcm01 having count(*) > 1 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strExc(1) = ""
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         strExc(1) = strExc(1) & vbCrLf & "流水號：" & RsTemp.Fields("mcm01")
         RsTemp.MoveNext
      Loop
      If strExc(1) <> "" Then
         MsgBox "下列藥證號數對照檔在更新案號會產生重複資料，請先行調整：" & strExc(1), vbExclamation
         GoTo EXITSUB
      End If
   End If
   'end 2023/06/28
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textOLD01_GotFocus()
   CloseIme
   InverseTextBox textOLD01
End Sub

Private Sub textOLD02_GotFocus()
   InverseTextBox textOLD02
End Sub

Private Sub textOLD02_2_GotFocus()
   InverseTextBox textOLD02_2
End Sub

Private Sub textOLD03_GotFocus()
   InverseTextBox textOLD03
End Sub

Private Sub textOLD04_GotFocus()
   InverseTextBox textOLD04
End Sub

Private Sub textNEW01_GotFocus()
   CloseIme
   InverseTextBox textNEW01
End Sub

Private Sub textNEW02_GotFocus()
   InverseTextBox textNEW02
End Sub

Private Sub textNEW02_2_GotFocus()
   InverseTextBox textNEW02_2
End Sub

Private Sub textNEW03_GotFocus()
   InverseTextBox textNEW03
End Sub

Private Sub textNEW04_GotFocus()
   InverseTextBox textNEW04
End Sub

Private Sub ShowStatus(ByVal strData As String)
   If IsEmptyText(strData) = False Then
      textStatus = "執行 : " & strData
   Else
      textStatus = strData
   End If
   textStatus.Refresh
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textNEW01.Enabled = True Then
   Cancel = False
   textNEW01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textNEW02.Enabled = True Then
   Cancel = False
   textNEW02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textNEW02_2.Enabled = True Then
   Cancel = False
   textNEW02_2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textNEW03.Enabled = True Then
   Cancel = False
   textNEW03_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textOLD01.Enabled = True Then
   Cancel = False
   textOLD01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textOLD02.Enabled = True Then
   Cancel = False
   textOLD02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textOLD02_2.Enabled = True Then
   Cancel = False
   textOLD02_2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textOLD03.Enabled = True Then
   Cancel = False
   textOLD03_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Sindy 2019/12/27
If Me.textCP13.Enabled = True Then
   Cancel = False
   textCP13_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'2019/12/27 END

If Me.textDelete.Enabled = True Then
   Cancel = False
   textDelete_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function

'Added by Lydia 2025/10/15
Private Sub textomemo_GotFocus()
   TextInverse textOMemo
End Sub
