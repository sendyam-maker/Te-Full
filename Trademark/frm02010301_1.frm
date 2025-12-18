VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010301_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標申請案號輸入"
   ClientHeight    =   5748
   ClientLeft      =   276
   ClientTop       =   960
   ClientWidth     =   9324
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9324
   Begin VB.TextBox textTM02_2 
      Height          =   264
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   660
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.CommandButton cmdNoReg 
      Caption         =   "未申請(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5325
      TabIndex        =   5
      Top             =   70
      Width           =   1020
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "所有資料(&L)"
      Height          =   400
      Left            =   6372
      TabIndex        =   6
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7596
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8424
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   6060
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   660
      Width           =   2292
   End
   Begin VB.TextBox textTM06 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1260
      Width           =   7512
   End
   Begin VB.TextBox textTM01 
      Height          =   264
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   0
      Top             =   660
      Width           =   612
   End
   Begin VB.TextBox textTM02 
      Height          =   264
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   1
      Top             =   660
      Width           =   972
   End
   Begin VB.TextBox textTM03 
      Height          =   264
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   3
      Top             =   660
      Width           =   372
   End
   Begin VB.TextBox textTM04 
      Height          =   264
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   4
      Top             =   660
      Width           =   612
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   3492
      Left            =   72
      TabIndex        =   23
      Top             =   2184
      Width           =   9132
      _ExtentX        =   16108
      _ExtentY        =   6160
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
   Begin MSForms.TextBox textTM05_1 
      Height          =   840
      Left            =   1680
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   960
      Width           =   7512
      VariousPropertyBits=   679493663
      ScrollBars      =   2
      Size            =   "13250;1482"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   264
      Left            =   1680
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1860
      Width           =   7512
      VariousPropertyBits=   679493663
      Size            =   "13250;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM07 
      Height          =   264
      Left            =   1680
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1560
      Width           =   7512
      VariousPropertyBits=   679493663
      Size            =   "13250;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM05 
      Height          =   264
      Left            =   1680
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   960
      Width           =   7512
      VariousPropertyBits=   679493663
      Size            =   "13250;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "注意：分割案申請案號請由核准進入"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   180
      TabIndex        =   22
      Top             =   300
      Width           =   2880
   End
   Begin VB.Label Label7 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   19
      Top             =   1860
      Width           =   1572
   End
   Begin VB.Label Label5 
      Caption         =   "案件日文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   1572
   End
   Begin VB.Label Label4 
      Caption         =   "案件英文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   17
      Top             =   1260
      Width           =   1572
   End
   Begin VB.Label Label3 
      Caption         =   "案件中文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   1572
   End
   Begin VB.Label Label2 
      Caption         =   "申請國家 :"
      Height          =   252
      Left            =   4980
      TabIndex        =   15
      Top             =   660
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Left            =   120
      TabIndex        =   14
      Top             =   660
      Width           =   1572
   End
End
Attribute VB_Name = "frm02010301_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/18 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2021/12/28 Form2.0已修改 textTM05_1/textTM05/textTM07/textTM23/grdList
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

Dim m_CP09 As String
Public strTM01 As String
Public strTM02 As String
Public strTM03 As String
Public strTM04 As String
'Add By Sindy 2019/5/10
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_RDate As String
Dim m_Done As Boolean
Dim m_PrevForm As Form
'2019/5/10 END


'Add By Sindy 2019/5/13
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

'Added by Sindy 2019/5/10
Private Sub Form_Activate()
   If m_strIR01 <> "" And m_Done = False Then
      textTM01.Text = strTM01
      If strTM01 = "TF" Then
         textTM02.Text = Left(strTM02, 5)
         textTM02_2.Text = Mid(strTM02, 6)
      Else
         textTM02.Text = strTM02
      End If
      textTM03.Text = strTM03
      textTM04.Text = strTM04
      cmdAll.Value = True
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
End Sub
'2019/5/10 END

Private Sub Form_Load()
    ' 設定只顯示不輸入的控制項其背景顏色
    textTM10.BackColor = &H8000000F
    textTM05.BackColor = &H8000000F
    textTM05_1.BackColor = &H8000000F
    textTM06.BackColor = &H8000000F
    textTM07.BackColor = &H8000000F
    textTM23.BackColor = &H8000000F
    MoveFormToCenter Me
    InitialGrdList
    'Add By Cheng 2003/07/18
    '預設系統類別
    Me.textTM01.Text = "T"
    Select Case Me.textTM01.Text
    Case "T", "FCT", "TF", "TS"
        Me.Label7.Visible = True
        Me.textTM05_1.Visible = True
        Me.Label3.Visible = False
        Me.textTM05.Visible = False
        Me.Label4.Visible = False
        Me.textTM05.Visible = False
        Me.Label5.Visible = False
        Me.textTM07.Visible = False
    Case Else
        Me.Label7.Visible = False
        Me.textTM05_1.Visible = False
        Me.Label3.Visible = True
        Me.textTM05.Visible = True
        Me.Label4.Visible = True
        Me.textTM05.Visible = True
        Me.Label5.Visible = True
        Me.textTM07.Visible = True
    End Select
    SendKeys "{Tab}"
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdNoReg_Click()
   If CheckDataValid = True Then
      ListData 0
   End If
End Sub

Private Sub cmdAll_Click()
   If CheckDataValid = True Then
      ListData 1
   End If
End Sub

Private Sub cmdOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(m_CP09) = False Then
      DisplayNextForm
   Else
      strTit = "請先選取收文號"
      strMsg = "檢核資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
End Sub

' 由案件性質代碼取得案件性質名稱
Private Function GetCaseType(ByVal strKey1 As String, ByVal StrKey2 As String) As String
   Dim rsTmp As ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   GetCaseType = Empty
   If IsEmptyText(strKey1) = False And IsEmptyText(StrKey2) = False Then
      Set rsTmp = New ADODB.Recordset
      strSql = "SELECT * FROM CasePropertyMap " & _
               "WHERE CPM01 = '" & strKey1 & "' AND " & _
                     "CPM02 = '" & StrKey2 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("CPM03")) = False Then
            GetCaseType = rsTmp.Fields("CPM03")
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function
' 由客戶代碼取得客戶名稱
Private Function GetCustomer(ByVal strData As String) As String
   Dim rsTmp As ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   GetCustomer = Empty
   If IsEmptyText(strData) = False Then
      Set rsTmp = New ADODB.Recordset
      If Len(strData) > 8 Then
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                        "CU02 = '" & Mid(strData, 9, 1) & "'"
      Else
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "'"
      End If
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("CU04")) = False Then
            GetCustomer = rsTmp.Fields("CU04")
         ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
            GetCustomer = rsTmp.Fields("CU05")
         ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
            GetCustomer = rsTmp.Fields("CU06")
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function
' 取得國家的名稱
Private Function GetNation(ByVal strNation As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   GetNation = Empty
   If IsEmptyText(strNation) = False Then
      strSql = "SELECT * FROM NATION " & _
               "WHERE NA01 = '" & strNation & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("NA03")) = False Then
            GetNation = rsTmp.Fields("NA03")
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

' 列出所有資料
' Input : nType = 0 表取得的資料是未申請
'                 1 表取得的資料為所有的
Public Sub ListData(ByVal nType As Integer)
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strNationNo As String
Dim m_TM12 As String   '2016/3/21 add by sonia
   
   strNationNo = Empty
   InitialGrdList
   
   'Add By Sindy 2019/5/10
   If m_strIR01 <> "" Then
      If strTM01 & strTM02 & strTM03 & strTM04 <> textTM01 & IIf(textTM01 = "TF", textTM02 & textTM02_2, textTM02) & textTM03 & textTM04 Then
         MsgBox "信件輸入必須與信件本所案號(" & strTM01 & strTM02 & strTM03 & strTM04 & ")一致！"
         Exit Sub
      End If
   End If
   '2019/5/10 END
   
   strTM01 = textTM01
   strTM02 = textTM02
   If strTM01 = "TF" Then
      strTM02 = strTM02 & textTM02_2
   End If
   strTM03 = textTM03
   strTM04 = textTM04
   If IsEmptyText(strTM03) = True Then: strTM03 = "0"
   If IsEmptyText(strTM04) = True Then: strTM04 = "00"
'edit by nickc 2006/10/18  分割案已經不從此處進入
    'add by nick 2004/12/23  檢查分割案
'    If CheckDC = True Then
'        frm02010301_3.SetData strTM01, strTM02, strTM03, strTM04
'        Me.Hide
'        frm02010301_3.Show
'        frm02010301_3.QueryData
'        Exit Sub
'    End If
   Select Case strTM01
   Case "T", "FCT", "TF", "TS"
       Me.Label7.Visible = True
       Me.textTM05_1.Visible = True
       Me.Label3.Visible = False
       Me.textTM05.Visible = False
       Me.Label4.Visible = False
       Me.textTM05.Visible = False
       Me.Label5.Visible = False
       Me.textTM07.Visible = False
   Case Else
       Me.Label7.Visible = False
       Me.textTM05_1.Visible = False
       Me.Label3.Visible = True
       Me.textTM05.Visible = True
       Me.Label4.Visible = True
       Me.textTM05.Visible = True
       Me.Label5.Visible = True
       Me.textTM07.Visible = True
   End Select
   
   m_TM12 = ""   '2016/3/21 add by sonia
   Select Case strTM01
      Case "T", "TF", "FCT", "CFT":
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM01 = '" & strTM01 & "' AND " & _
                        "TM02 = '" & strTM02 & "' AND " & _
                        "TM03 = '" & strTM03 & "' AND " & _
                        "TM04 = '" & strTM04 & "'"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            ' 申請國家
            If IsNull(rsTmp.Fields("TM10")) = False Then
               textTM10 = GetNation(rsTmp.Fields("TM10"))
               strNationNo = rsTmp.Fields("TM10")
            End If
            ' 商標名稱(中)
            If IsNull(rsTmp.Fields("TM05")) = False Then
'               textTM05 = rsTmp.Fields("TM05")
               textTM05_1 = rsTmp.Fields("TM05")
            End If
'            ' 商標名稱(英)
'            If IsNull(rsTmp.Fields("TM06")) = False Then
'               textTM06 = rsTmp.Fields("TM06")
'            End If
'            ' 商標名稱(日)
'            If IsNull(rsTmp.Fields("TM07")) = False Then
'               textTM07 = rsTmp.Fields("TM07")
'            End If
            ' 申請人
            If IsNull(rsTmp.Fields("TM23")) = False Then
               textTM23 = GetCustomer(rsTmp.Fields("TM23"))
            End If
            '2016/3/21 add by sonia
            If IsNull(rsTmp.Fields("TM12")) = False Then
               m_TM12 = rsTmp.Fields("TM12")
            End If
            '2016/3/21 ebd
         End If
         rsTmp.Close
      Case Else
         strSql = "SELECT * FROM ServicePractice " & _
                  "WHERE SP01 = '" & strTM01 & "' AND " & _
                        "SP02 = '" & strTM02 & "' AND " & _
                        "SP03 = '" & strTM03 & "' AND " & _
                        "SP04 = '" & strTM04 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            ' 申請國家
            If IsNull(rsTmp.Fields("SP09")) = False Then
               textTM10 = GetNation(rsTmp.Fields("SP09"))
               strNationNo = rsTmp.Fields("SP09")
            End If
            Select Case strTM01
            Case "TS"
                textTM05_1 = "" & rsTmp.Fields("SP05")
            Case Else
                ' 商標名稱(中)
                If IsNull(rsTmp.Fields("SP05")) = False Then
                   textTM05 = rsTmp.Fields("SP05")
                End If
            End Select
            ' 商標名稱(英)
            If IsNull(rsTmp.Fields("SP06")) = False Then
               textTM06 = rsTmp.Fields("SP06")
            End If
            ' 商標名稱(日)
            If IsNull(rsTmp.Fields("SP07")) = False Then
               textTM07 = rsTmp.Fields("SP07")
            End If
            ' 申請人
            If IsNull(rsTmp.Fields("SP08")) = False Then
               textTM23 = GetCustomer(rsTmp.Fields("SP08"))
            End If
            '2016/3/21 add by sonia
            If IsNull(rsTmp.Fields("sp11")) = False Then
               m_TM12 = rsTmp.Fields("sp11")
            End If
            '2016/3/21 ebd
         End If
         rsTmp.Close
   End Select
      
   ' 案件進度檔列表
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & strTM01 & "' AND " & _
                  "CP02 = '" & strTM02 & "' AND " & _
                  "CP03 = '" & strTM03 & "' AND " & _
                  "CP04 = '" & strTM04 & "' "
   'add by nickc 2006/10/18 加入不抓分割案
   strSql = strSql & " and cp10 <> '308' "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         ' 申請國家為大陸時
         '92.9.10 MODIFY BY SONIA
         'If strNationNo = "020" Then
         If strNationNo <> "000" Then
         '92.9.10 END
            If IsNull(rsTmp.Fields("CP09")) = False Then
               Select Case Mid(rsTmp.Fields("CP09"), 1, 1)
                  ' 收文號必須為A,B類的
                  Case "A", "B"
                     ' 當取得的資料必須是未申請的時, 大陸申請案號欄位必須為空白, 否則不予列入
                     ' 若取得的資料為所有的時, 則不管大陸申請案號欄位
                     If nType = 0 Then
                        If IsNull(rsTmp.Fields("CP30")) = False Then
                           If IsEmptyText(rsTmp.Fields("CP30")) = False Then
                              GoTo NextRecord
                           End If
                        End If
                        '************   90.11.23  nick
                        '*邱小姐說發文日空白的不出現
                        If IsNull(rsTmp.Fields("CP27")) = True Then
                           GoTo NextRecord
                        End If
                        '*******************
                     End If
                  Case Else
                     GoTo NextRecord
               End Select
            Else
               GoTo NextRecord
            End If
         Else
            ' 申請國家為非大陸時只取得案件性質為申請的
            If IsNull(rsTmp.Fields("CP10")) = False Then
               'edit by nick 2004/12/23 加入分割與申請相同
               'If rsTmp.Fields("CP10") <> "101" And rsTmp.Fields("CP10") <> "806" Then
               If rsTmp.Fields("CP10") <> "101" And rsTmp.Fields("CP10") <> "806" And rsTmp.Fields("CP10") <> "308" Then
                  GoTo NextRecord
               '2016/3/21 modify by sonia
               ElseIf nType = 0 And m_TM12 <> "" Then
                  GoTo NextRecord
               '2016/3/21 end
               End If
            Else
               GoTo NextRecord
            End If
         End If
      
         grdList.Rows = grdList.Rows + 1
         grdList.row = grdList.Rows - 1
         ' 收文號
         If IsNull(rsTmp.Fields("CP09")) = False Then
            grdList.TextMatrix(grdList.row, 1) = rsTmp.Fields("CP09")
         End If
         ' 收文日
         If IsNull(rsTmp.Fields("CP05")) = False Then
            grdList.TextMatrix(grdList.row, 2) = TAIWANDATE(rsTmp.Fields("CP05"))
         End If
         ' 案件性質
         If strNationNo < "010" Then
            If IsNull(rsTmp.Fields("CP10")) = False Then
               grdList.TextMatrix(grdList.row, 3) = GetCaseTypeName(strTM01, rsTmp.Fields("CP10"), 0)
            End If
         Else
            If IsNull(rsTmp.Fields("CP10")) = False Then
               grdList.TextMatrix(grdList.row, 3) = GetCaseTypeName(strTM01, rsTmp.Fields("CP10"), 1)
            End If
         End If
         ' 發文日
         If IsNull(rsTmp.Fields("CP27")) = False Then
            grdList.TextMatrix(grdList.row, 4) = TAIWANDATE(rsTmp.Fields("CP27"))
         'add by sonia 2015/12/29
         Else
            MsgBox "此進度尚未發文！請補發文 !", vbCritical
         'end 2015/12/29
         End If
         ' 大陸申請案號
         If IsNull(rsTmp.Fields("CP30")) = False Then
            grdList.TextMatrix(grdList.row, 5) = rsTmp.Fields("CP30")
         End If
NextRecord:
         rsTmp.MoveNext
      Loop
      'Added by Lydia 2023/10/18
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/18
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   ' 列示出的項目只有一筆時, 直接顯示詳細的資料
   If grdList.Rows > 1 Then
      grdList.row = 1
      grdList.col = 1
      m_CP09 = grdList.Text
      grdList_ShowSelection
   End If
   If grdList.Rows = 2 Then
      DisplayNextForm
   End If
   '2016/3/21 modify by sonia
   If grdList.Rows = 1 Then
      MsgBox "無符合條件資料！", vbCritical
   End If
   '2016/3/21 end
End Sub
' 初始化 GridList
Public Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 6
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "收文號"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "收文日"
   grdList.ColWidth(2) = 1000
   grdList.col = 3
   grdList.Text = "案件性質"
   grdList.ColWidth(3) = 1200
   grdList.col = 4
   grdList.Text = "發文日"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "申請案號"
   grdList.ColWidth(5) = 1200
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2019/5/13
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   'Add By Cheng 2002/07/18
   Set frm02010301_1 = Nothing
End Sub

Private Sub grdList_SelChange()
   ' 將選取的收文號儲存起來
   If grdList.row > 0 Then
      grdList.col = 1
      m_CP09 = grdList.Text
      cmdOK.SetFocus
   End If
   grdList_ShowSelection
End Sub
' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nRow As Integer
   Dim nCol As Integer
   Dim nCurrSel As Integer
   nCurrSel = grdList.row
   For nRow = 1 To grdList.Rows - 1
      grdList.row = nRow
      If nRow = nCurrSel Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            grdList.CellBackColor = &H8000000D
            grdList.CellForeColor = &H80000005
         Next nCol
      Else
         grdList.col = 1
         If grdList.CellBackColor <> &H80000005 Then
            For nCol = 1 To grdList.Cols - 1
               grdList.col = nCol
               If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
               If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
            Next nCol
         End If
      End If
   Next nRow
   grdList.row = nCurrSel
   grdList.col = 0
End Sub

' 系統別
Private Sub textTM01_Validate(Cancel As Boolean)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim strGroup As String
   Dim nResponse
   
   If IsEmptyText(textTM01) = False Then
      Cancel = True
      If Mid(textTM01, 1, 1) <> "T" Then
         strMsg = "系統別不正確"
         strTit = "權限檢查"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM01_GotFocus
         GoTo EXITSUB
      End If
      
      Select Case textTM01
         Case "T":
            textTM02_2.Visible = False
            textTM02_2.Locked = True
            textTM02_2.TabStop = False
            textTM02.MaxLength = 6
         Case "TF":
            textTM02_2.Visible = True
            textTM02_2.Locked = False
            textTM02_2.TabStop = True
            textTM02.MaxLength = 5
                     
         '910709 Sieg 0620-2
'         Case Else
'            Cancel = True
'            strTit = "資料檢核"
'            strMsg = "本所案號中的系統別不正確"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textTM01_GotFocus
'            GoTo EXITSUB
      End Select
      
      strGroup = Empty
      strSql = "SELECT * FROM Staff WHERE ST01 = '" & strUserNum & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("ST11")) = False Then
            strGroup = rsTmp.Fields("ST11")
         End If
      End If
      rsTmp.Close
   
      strSql = "SELECT * FROM Staff_Group " & _
               "WHERE SG01 = '" & strGroup & "' AND " & _
                     "SG02 = '" & textTM01 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.Close
         Cancel = False
         GoTo EXITSUB
      End If
            
      rsTmp.Close
      ' 顯示錯誤訊息
      Cancel = True
      strMsg = "您的使用權限無法存取該系統別的資料"
      strTit = "權限檢查"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM01_GotFocus
      GoTo EXITSUB
   Else
      textTM02_2.Visible = False
      textTM02_2.Locked = True
      textTM02_2.TabStop = False
      textTM02.MaxLength = 6
   End If
   
   Cancel = False
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示編輯案件進度檔案的畫面
Private Sub DisplayNextForm()
   'Add By Sindy 2019/5/10
   If m_strIR01 <> "" Then
      If strTM01 & strTM02 & strTM03 & strTM04 <> textTM01 & IIf(textTM01 = "TF", textTM02 & textTM02_2, textTM02) & textTM03 & textTM04 Then
         MsgBox "信件輸入必須與信件本所案號(" & strTM01 & strTM02 & strTM03 & strTM04 & ")一致！"
         Exit Sub
      End If
   End If
   '2019/5/10 END
   
   strTM01 = textTM01
   strTM02 = textTM02
   If strTM01 = "TF" Then
      strTM02 = strTM02 & textTM02_2
   End If
   strTM03 = textTM03
   strTM04 = textTM04
   If IsEmptyText(strTM03) = True Then: strTM03 = "0"
   If IsEmptyText(strTM04) = True Then: strTM04 = "00"
   
   If IsEmptyText(m_CP09) = False Then
      frm02010301_2.SetData strTM01, strTM02, strTM03, strTM04, m_CP09
      'cmdNoReg.SetFocus
      'Add By Sindy 2019/5/10
      If Not m_PrevForm Is Nothing Then
         Call frm02010301_2.SetParent(m_PrevForm)
      End If
      frm02010301_2.m_strIR01 = m_strIR01
      frm02010301_2.m_strIR02 = m_strIR02
      frm02010301_2.m_strIR03 = m_strIR03
      frm02010301_2.m_strIR04 = m_strIR04
      '2019/5/10 END
      Me.Hide
      frm02010301_2.Show
      frm02010301_2.UpdateCtrl
   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   If IsEmptyText(textTM01) = True Then
      strMsg = "請輸入本所案號"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If IsEmptyText(textTM02) = True Then
      strMsg = "請輸入本所案號"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If textTM01 = "TF" Then
      If IsEmptyText(textTM02_2) = True Then
         strMsg = "本所案號輸入不完整"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textTM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM01_GotFocus()
   InverseTextBox textTM01
   CloseIme
End Sub

Private Sub textTM02_GotFocus()
   InverseTextBox textTM02
   CloseIme
End Sub

Private Sub textTM02_2_GotFocus()
   InverseTextBox textTM02_2
   CloseIme
End Sub

Private Sub textTM03_GotFocus()
   InverseTextBox textTM03
   CloseIme
End Sub

Private Sub textTM03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM04_GotFocus()
   InverseTextBox textTM04
   CloseIme
End Sub
'add by nick 2004/12/23 檢查是否有分割
Public Function CheckDC() As Boolean
CheckOC3
Dim strSql As String
strSql = "select count(*) from divisioncase where dc05='" & strTM01 & "' and dc06='" & strTM02 & "' and dc07='" & strTM03 & "' and dc08='" & strTM04 & "' "
With AdoRecordSet3
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .Fields(0).Value <> 0 Then
        CheckDC = True
    Else
        CheckDC = False
    End If
End With
CheckOC3
End Function
