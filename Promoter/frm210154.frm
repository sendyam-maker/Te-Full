VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210154 
   BorderStyle     =   1  '單線固定
   Caption         =   "下一程序接洽單 列印"
   ClientHeight    =   4884
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8604
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4884
   ScaleWidth      =   8604
   Begin VB.TextBox textTM12 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   4890
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1545
      Width           =   2052
   End
   Begin VB.TextBox textTM10 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1545
      Width           =   1932
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   435
      Index           =   2
      Left            =   7620
      TabIndex        =   6
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Height          =   435
      Index           =   1
      Left            =   6600
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   5580
      TabIndex        =   4
      Top             =   90
      Width           =   975
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   3
      Top             =   525
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2580
      MaxLength       =   1
      TabIndex        =   2
      Top             =   525
      Width           =   240
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1650
      MaxLength       =   6
      TabIndex        =   1
      Top             =   525
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   0
      Top             =   525
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   2868
      Left            =   24
      TabIndex        =   20
      Top             =   1920
      Width           =   8556
      _ExtentX        =   15092
      _ExtentY        =   5059
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
   Begin MSForms.ComboBox cmbTM05 
      Height          =   300
      Left            =   1080
      TabIndex        =   19
      Top             =   840
      Width           =   7305
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "12885;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1080
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1215
      Width           =   7275
      VariousPropertyBits=   671105055
      BackColor       =   -2147483633
      MaxLength       =   20
      Size            =   "12832;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      Caption         =   "紅色：與進度檔本所案號不同"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4410
      TabIndex        =   17
      Top             =   630
      Width           =   2565
   End
   Begin VB.Label Label7 
      Caption         =   "綠色：結案中"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7320
      TabIndex        =   16
      Top             =   630
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "申請案號："
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   1545
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "申請國家："
      Height          =   255
      Left            =   90
      TabIndex        =   14
      Top             =   1545
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   90
      TabIndex        =   13
      Top             =   870
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "申  請  人："
      Height          =   255
      Left            =   90
      TabIndex        =   12
      Top             =   1215
      Width           =   975
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
      Height          =   255
      Left            =   6990
      TabIndex        =   11
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbldestroy 
      Caption         =   "北所銷卷"
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
      Height          =   255
      Left            =   7740
      TabIndex        =   10
      Top             =   1560
      Width           =   795
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1500
      X2              =   2955
      Y1              =   660
      Y2              =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   255
      Left            =   90
      TabIndex        =   7
      Top             =   570
      Width           =   975
   End
End
Attribute VB_Name = "frm210154"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/17 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Lydia 2022/01/05 改成Form2.0 ; grdList改字型=新細明體-ExtB、cmbTM05、textTM23
'Create by Amy 2019/05/31
Option Explicit

Dim RsQ As New ADODB.Recordset
Dim strQ As String, m_Nation As String
Dim strNP02 As String, strNP03 As String, strNP04 As String, strNP05 As String, strNP22 As String
Dim m_row As Integer, m_CurrSel As Integer

Private Sub cmdOK_Click(Index As Integer)
    Select Case Index
        Case 0
            If Trim(txt1(0)) = "" Or Trim(txt1(1)) = "" Then
                MsgBox "本所案號不可以空白！", vbCritical, "操作錯誤！"
                Exit Sub
            End If
            Label5.Visible = False
            Label7.Visible = False
            Screen.MousePointer = vbHourglass
            grdList.MousePointer = flexHourglass
            doQuery
            grdList.MousePointer = flexDefault
            Screen.MousePointer = vbDefault
        Case 1
            If IsNull(grdList.TextMatrix(grdList.row, 13)) = True Then Exit Sub
            If Not (strNP02 = "FCT" Or strNP02 = "S") Then
                MsgBox "無列印下一程序接洽單權限！", vbCritical, "操作錯誤！"
                Exit Sub
            End If
            
            strNP22 = "" & grdList.TextMatrix(grdList.row, 12)
            If IsEmptyText(strNP22) = False Then
                g_PrtForm001.PrintForm strNP22, strNP02, strNP03, strNP04, strNP05, , , , , , , , Me.Name
            End If

        Case 2
            Unload Me
    End Select
End Sub

Private Function doQuery() As Boolean
    Dim bQuery As Boolean, bQueryNP As Boolean
    Dim strTit As String, strMsg As String
    Dim nResponse
    
On Error GoTo ErrHnd
    strNP02 = UCase(txt1(0))
    strNP03 = txt1(1)
    strNP04 = Left(txt1(2) & "0", 1)
    strNP05 = Left(txt1(3) & "00", 2)
    
    ClearField
    InitialGridList
    bQuery = False
   
    '權限控制(同frm075007_1)
    If CheckSR09(strUserNum, strNP02, "Y", , strNP02, strNP03, strNP04, strNP05) = False Then
        Exit Function
    End If
    
    'T非台灣案非外商收文之案件不必寫程式控制,因為在系統類別外商人員即不可使用T案件
    'FCT無爭議程序之案件內商人員不可查詢(該案有內商承辦人者為FCT爭議案)
    If strNP02 = "FCT" And Mid(GetStaffDepartment(strUserNum), 1, 2) = "P2" Then
        strQ = "Select * From CaseProgress,Staff Where CP01='" & strNP02 & "' AND CP02='" & strNP03 & "' AND CP03='" & strNP04 & "' AND CP04='" & strNP05 & "' AND CP14=ST01(+) AND SUBSTR(ST03,1,2)='P2' "
        RsQ.CursorLocation = adUseClient
        RsQ.Open strQ, cnnConnection, adOpenStatic
        If RsQ.RecordCount = 0 Then
            strMsg = "非FCT爭議案，您沒有使用該案號資料的權限"
            strTit = "查詢資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            If RsQ.State <> adStateClosed Then RsQ.Close
            Set RsQ = Nothing
            Exit Function
        Else
            If RsQ.State <> adStateClosed Then RsQ.Close
            Set RsQ = Nothing
        End If
    End If
    
    '依本所案號讀取基本檔案
    Select Case strNP02
        ' 讀取商標基本檔
        Case "T", "TF", "CFT", "FCT":
            bQuery = QueryTradeMark(strNP02, strNP03, strNP04, strNP05)
        ' 讀取專利基本檔
        Case "P", "CFP", "FCP":
            bQuery = QueryPatent(strNP02, strNP03, strNP04, strNP05)
        ' 讀取法務基本檔
        Case "L", "CFL", "FCL", "LIN":
            bQuery = QueryLawCase(strNP02, strNP03, strNP04, strNP05)
        ' 讀取顧問案件基本檔
        Case "LA":
            bQuery = QueryHireCase(strNP02, strNP03, strNP04, strNP05)
        ' 讀取服務業務基本檔
        Case Else:
            bQuery = QueryServicePractice(strNP02, strNP03, strNP04, strNP05)
    End Select
    
    ' 讀取下一程序檔
    bQueryNP = False
    bQueryNP = QueryNextProgress(strNP02, strNP03, strNP04, strNP05)
   
    If bQuery = False Then
        strTit = "查詢資料"
        strMsg = "該筆不存在於基本檔中"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
    Else
        If bQueryNP = False Then
            strTit = "查詢資料"
            strMsg = "該筆案件無下一程序的資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        End If
    End If
    doQuery = bQueryNP
  
ErrHnd:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub Form_Load()
    MoveFormToCenter Me
    ClearField
    Label5.Visible = False
    Label7.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm210154 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
    TextInverse txt1(Index)
    CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
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
            cmbTM05.AddItem "中 : " & rsTmp.Fields("TM05")
        End If
        If IsNull(rsTmp.Fields("TM06")) = False Then
            cmbTM05.AddItem "英 : " & rsTmp.Fields("TM06")
        End If
        If IsNull(rsTmp.Fields("TM07")) = False Then
            cmbTM05.AddItem "日 : " & rsTmp.Fields("TM07")
        End If
        ' 顯示商標名稱
        If cmbTM05.ListCount > 0 Then
            cmbTM05.ListIndex = 0
        End If
        ' 申請人
        If IsNull(rsTmp.Fields("TM23")) = False Then
            textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
        End If
        ' 申請國家
        If IsNull(rsTmp.Fields("TM10")) = False Then
            m_Nation = rsTmp.Fields("TM10")
         textTM10 = GetNationName(rsTmp.Fields("TM10"), 0)
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      '顯示是否閉卷
      If rsTmp("TM29") = "Y" Then
         Me.lblClose.Caption = "已閉卷"
      End If
      '顯示北所是否銷卷
      If IsNull(rsTmp("TM57")) = False Then
         Me.lbldestroy.Caption = "北所銷卷"
      End If
      '2012/7/17 END
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
            cmbTM05.AddItem "中 : " & rsTmp.Fields("SP05")
        End If
        If IsNull(rsTmp.Fields("SP06")) = False Then
            cmbTM05.AddItem "英 : " & rsTmp.Fields("SP06")
        End If
        If IsNull(rsTmp.Fields("SP07")) = False Then
            cmbTM05.AddItem "日 : " & rsTmp.Fields("SP07")
        End If
        ' 顯示商標名稱
        If cmbTM05.ListCount > 0 Then
            cmbTM05.ListIndex = 0
        End If
        ' 申請人
        If IsNull(rsTmp.Fields("SP08")) = False Then
            textTM23 = GetCustomerName(rsTmp.Fields("SP08"), 0)
        End If
        ' 申請國家
        If IsNull(rsTmp.Fields("SP09")) = False Then
            m_Nation = rsTmp.Fields("SP09")
            textTM10 = GetNationName(rsTmp.Fields("SP09"), 0)
        End If
        ' 申請案號
        If IsNull(rsTmp.Fields("SP11")) = False Then
            textTM12 = rsTmp.Fields("SP11")
        End If
        '顯示是否閉卷
        If rsTmp("SP15") = "Y" Then
            Me.lblClose.Caption = "已閉卷"
        End If
        '顯示北所是否銷卷
        If IsNull(rsTmp("SP61")) = False Then
            Me.lbldestroy.Caption = "北所銷卷"
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
      ' 案件名稱
      If IsNull(rsTmp.Fields("PA05")) = False Then
         cmbTM05.AddItem "中 : " & rsTmp.Fields("PA05")
      End If
      If IsNull(rsTmp.Fields("PA06")) = False Then
         cmbTM05.AddItem "英 : " & rsTmp.Fields("PA06")
      End If
      If IsNull(rsTmp.Fields("PA07")) = False Then
         cmbTM05.AddItem "日 : " & rsTmp.Fields("PA07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("PA26")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("PA26"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("PA09")) = False Then
         m_Nation = rsTmp.Fields("PA09")
         textTM10 = GetNationName(rsTmp.Fields("PA09"), 0)
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("PA11")) = False Then
         textTM12 = rsTmp.Fields("PA11")
      End If
      'Add By Cheng 2002/04/25
      '顯示是否閉卷
      If rsTmp("PA57") = "Y" Then
         Me.lblClose.Caption = "已閉卷"
      End If
      '2012/7/17 ADD BY SONIA
      '顯示北所是否銷卷
      If IsNull(rsTmp("PA108")) = False Then
         Me.lbldestroy.Caption = "北所銷卷"
      End If
      '2012/7/17 END
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
            cmbTM05.AddItem "中 : " & rsTmp.Fields("LC05")
        End If
        If IsNull(rsTmp.Fields("LC06")) = False Then
            cmbTM05.AddItem "英 : " & rsTmp.Fields("LC06")
        End If
        If IsNull(rsTmp.Fields("LC07")) = False Then
            cmbTM05.AddItem "日 : " & rsTmp.Fields("LC07")
        End If
        ' 顯示商標名稱
        If cmbTM05.ListCount > 0 Then
            cmbTM05.ListIndex = 0
        End If
        ' 申請人
        If IsNull(rsTmp.Fields("LC11")) = False Then
            textTM23 = GetCustomerName(rsTmp.Fields("LC11"), 0)
        End If
        ' 申請國家
        If IsNull(rsTmp.Fields("LC15")) = False Then
            m_Nation = rsTmp.Fields("LC15")
            textTM10 = GetNationName(rsTmp.Fields("LC15"), 0)
        End If
        '顯示是否閉卷
        If rsTmp("LC08") = "Y" Then
            Me.lblClose.Caption = "已閉卷"
        End If
        '顯示北所是否銷卷
        If IsNull(rsTmp("LC34")) = False Then
            Me.lbldestroy.Caption = "北所銷卷"
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
        ' 案件名稱
        If IsNull(rsTmp.Fields("HC06")) = False Then
            cmbTM05.AddItem rsTmp.Fields("HC06")
        End If
        ' 顯示商標名稱
        If cmbTM05.ListCount > 0 Then
            cmbTM05.ListIndex = 0
        End If
        ' 申請人
        If IsNull(rsTmp.Fields("HC05")) = False Then
            textTM23 = GetCustomerName(rsTmp.Fields("HC05"), 0)
        End If
        '顯示是否閉卷
        If rsTmp("HC09") = "Y" Then
            Me.lblClose.Caption = "已閉卷"
        End If
        '顯示北所是否銷卷
        If IsNull(rsTmp("HC19")) = False Then
            Me.lbldestroy.Caption = "北所銷卷"
        End If
        m_Nation = "000"
        textTM10 = GetNationName(m_Nation, 0)
    End If
    rsTmp.Close
    Set rsTmp = Nothing
End Function

' 讀取下一程序檔資料
Private Function QueryNextProgress(ByVal strNP02 As String, ByVal strNP03 As String, ByVal strNP04 As String, ByVal strNP05 As String) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strCP09 As String
    Dim nRow As Integer
   
    QueryNextProgress = False
    strQ = "Select np01,np06,Nvl(" & IIf(m_Nation < "010", "C2.cpm03", "C2.cpm04") & ",np07) AS np07,Nvl(np08 - 19110000, NULL) AS np08, Nvl(np09 - 19110000, NULL) AS np09,Nvl(S1.ST02,NP10) AS NP10,NP22,Nvl(CP05 - 19110000, NULL) AS CP05,CP09,Nvl(S2.ST02,CP14) AS CP14,Nvl(CP27 - 19110000, NULL) AS CP27,NP14,NP15,CP01,CP02,CP03,CP04,NP24 " & _
                "From NextProgress, CaseProgress C1, CasePropertyMap C2, Staff S1, Staff S2 " & _
                "Where np02='" & strNP02 & "' And np03='" & strNP03 & "' And np04='" & strNP04 & "'And np05='" & strNP05 & "' " & _
                "And np01>='C' And np06 is null And np10=S1.st01(+) " & _
                "And np01 = C1.cp09(+) And cp14 = S2.st01(+) And np02 = C2.cpm01(+) And np07 = C2.cpm02(+) " & _
                "Order by cp05 Desc,np01 Desc,np08 Desc "
                
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount > 0 Then
        QueryNextProgress = True
        rsTmp.MoveFirst
        Do While rsTmp.EOF = False
            ' 新增一筆記錄
            grdList.Rows = grdList.Rows + 1
            nRow = grdList.Rows - 1
            ' 收文日
            If IsNull(rsTmp.Fields("CP05")) = False Then
                grdList.TextMatrix(nRow, 1) = rsTmp.Fields("CP05")
            End If
            ' 暫存總收文號
            strCP09 = Empty
            If IsNull(rsTmp.Fields("NP01")) = False Then
                strCP09 = rsTmp.Fields("NP01")
            End If
            ' 總收文號的欄位
            grdList.TextMatrix(nRow, 2) = strCP09
         
            ' 下一程序
            If IsNull(rsTmp.Fields("NP07")) = False Then
                grdList.TextMatrix(nRow, 3) = rsTmp.Fields("NP07")
            End If
            ' 本所期限
            If IsNull(rsTmp.Fields("NP08")) = False Then
                grdList.TextMatrix(nRow, 4) = rsTmp.Fields("NP08")
            End If
            ' 法定期限
            If IsNull(rsTmp.Fields("NP09")) = False Then
                grdList.TextMatrix(nRow, 5) = rsTmp.Fields("NP09")
            End If
            ' 是否續辦欄位
            If IsNull(rsTmp.Fields("NP06")) = False Then
                grdList.TextMatrix(nRow, 6) = rsTmp.Fields("NP06")
            End If
            ' 智權人員
            If IsNull(rsTmp.Fields("NP10")) = False Then
                grdList.TextMatrix(nRow, 7) = rsTmp.Fields("NP10")
            End If
            ' 相關人
            If IsNull(rsTmp.Fields("NP14")) = False Then
                grdList.TextMatrix(nRow, 8) = rsTmp.Fields("NP14")
            End If
            ' 備註
            If IsNull(rsTmp.Fields("NP15")) = False Then
                grdList.TextMatrix(nRow, 9) = rsTmp.Fields("NP15")
            End If
            ' 承辦人
            If IsNull(rsTmp.Fields("CP14")) = False Then
                grdList.TextMatrix(nRow, 10) = rsTmp.Fields("CP14")
            End If
            ' 發文日
            If IsNull(rsTmp.Fields("CP27")) = False Then
                grdList.TextMatrix(nRow, 11) = rsTmp.Fields("CP27")
            End If
            ' 序號
            If IsNull(rsTmp.Fields("NP22")) = False Then
                grdList.TextMatrix(nRow, 12) = rsTmp.Fields("NP22")
            End If
            ' 相關案件性質
            If IsNull(rsTmp.Fields("NP01")) = False Then
                grdList.TextMatrix(nRow, 3) = grdList.TextMatrix(nRow, 3) & PUB_GetNextCasePropertyName(grdList.TextMatrix(nRow, 2), grdList.TextMatrix(nRow, 12), "1")
            End If
            If IsNull(rsTmp.Fields("CP09")) = False Then
                grdList.TextMatrix(nRow, 13) = 0
            
                If rsTmp.Fields("CP01") <> strNP02 Then
                    grdList.TextMatrix(nRow, 13) = 1
                End If
                If rsTmp.Fields("CP02") <> strNP03 Then
                    grdList.TextMatrix(nRow, 13) = 1
                End If
                If rsTmp.Fields("CP03") <> strNP04 Then
                    grdList.TextMatrix(nRow, 13) = 1
                End If
                If rsTmp.Fields("CP04") <> strNP05 Then
                    grdList.TextMatrix(nRow, 13) = 1
                End If
        
            Else
                grdList.TextMatrix(nRow, 13) = 1
            End If
            If IsNull(rsTmp.Fields("NP24")) = False Then
                grdList.TextMatrix(nRow, 14) = rsTmp.Fields("NP24")
            End If
         
            ' 設定顯示的顏色
            Call SetColor(nRow)
            rsTmp.MoveNext
        Loop
        'Added by Lydia 2023/10/17
        If grdList.Rows >= 2 Then
           grdList.FixedRows = 1
        End If
        'end 2023/10/17
    End If
    rsTmp.Close
    Set rsTmp = Nothing
End Function

Private Sub SetColor(nRow As Integer)
    Dim nCol As Integer
   
    ' 設定顯示的顏色
    Select Case grdList.TextMatrix(nRow, 13)
        Case "1":
            For nCol = 1 To grdList.Cols - 1
                grdList.row = nRow
                grdList.col = nCol
                grdList.CellBackColor = &HFF& '紅色
                grdList.CellForeColor = &H80000008
            Next nCol
            Label5.Visible = True
        Case Else:
            If Trim(grdList.TextMatrix(nRow, 6)) = "" And Len(Trim(grdList.TextMatrix(nRow, 14))) = 8 Then '長度8是結案單編號
                For nCol = 1 To grdList.Cols - 1
                    grdList.row = nRow
                    grdList.col = nCol
                    grdList.CellBackColor = &H8000& '綠色
                    grdList.CellForeColor = &H80000008
                Next nCol
                Label7.Visible = True
            End If
    End Select
End Sub

' 初始化列表
Public Sub InitialGridList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 15

   grdList.ColWidth(0) = 300
   grdList.row = 0

   grdList.col = 0
   grdList.ColAlignment(0) = flexAlignCenterCenter
   grdList.col = 1
   grdList.Text = "收文日"
   grdList.ColWidth(1) = 900
   grdList.ColAlignment(1) = flexAlignCenterCenter
   grdList.col = 2
   grdList.Text = "總收文號"
   grdList.ColWidth(2) = 1000
   grdList.ColAlignment(2) = flexAlignCenterCenter
   grdList.col = 3
   grdList.Text = "下一程序"
   grdList.ColWidth(3) = 1400
   grdList.ColAlignment(3) = flexAlignLeftCenter
   grdList.col = 4
   grdList.Text = "本所期限"
   grdList.ColWidth(4) = 800
   grdList.ColAlignment(4) = flexAlignCenterCenter
   grdList.col = 5
   grdList.Text = "法定期限"
   grdList.ColWidth(5) = 800
   grdList.ColAlignment(5) = flexAlignCenterCenter
   grdList.col = 6
   grdList.Text = "續辦"
   grdList.ColWidth(6) = 600
   grdList.ColAlignment(6) = flexAlignCenterCenter
   grdList.col = 7
   grdList.Text = "智權人員"
   grdList.ColWidth(7) = 800
   grdList.ColAlignment(7) = flexAlignLeftCenter
   grdList.col = 8
   grdList.Text = "相關人"
   grdList.ColWidth(8) = 1600
   grdList.ColAlignment(8) = flexAlignLeftCenter
   grdList.col = 9
   grdList.Text = "備　註"
   grdList.ColWidth(9) = 3000
   grdList.ColAlignment(9) = flexAlignLeftCenter
   grdList.col = 10
   grdList.Text = "承辦人"
   grdList.ColWidth(10) = 0
   grdList.ColAlignment(10) = flexAlignLeftCenter
   grdList.col = 11
   grdList.Text = "發文日"
   grdList.ColWidth(11) = 0
   grdList.ColAlignment(11) = flexAlignCenterCenter
   grdList.col = 12
   grdList.Text = "序號"
   grdList.ColWidth(12) = 0
   grdList.ColAlignment(12) = flexAlignLeftCenter
   grdList.col = 13
   grdList.Text = "案件進度檔是否存在"
   grdList.ColWidth(13) = 0
   grdList.ColAlignment(13) = flexAlignLeftCenter
   grdList.col = 14
   grdList.Text = "NP24"
   grdList.ColWidth(14) = 0
   grdList.ColAlignment(14) = flexAlignLeftCenter
End Sub

Private Sub ClearField()
    cmbTM05.Clear
    textTM23 = Empty
    textTM10 = Empty
    textTM12 = Empty
    Me.lblClose.Caption = Empty
    Me.lbldestroy.Caption = Empty
End Sub

Private Sub grdList_Click()
    If grdList.row > 0 Then
        grdList.col = 0
        If grdList.Text = "V" Then
            grdList.Text = Empty
        Else
            grdList.Text = "V"
            cmdok(1).SetFocus
        End If
    End If
End Sub

Private Sub grdList_KeyDown(KeyCode As Integer, Shift As Integer)
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
        Exit Sub
    End If
   
    ' 將原先選取的列回復到正常的顏色
    If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
        grdList.row = m_CurrSel
        grdList.col = 1
        If grdList.CellBackColor <> &H80000005 Then
            Select Case grdList.TextMatrix(grdList.row, 13)
                Case "1":
                    For nCol = 1 To grdList.Cols - 1
                       grdList.col = nCol
                       If grdList.CellBackColor <> &HFF& Then: grdList.CellBackColor = &HFF& '紅色
                       If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
                    Next nCol
                Case Else:
                    If Trim(grdList.TextMatrix(grdList.row, 6)) = "" And Trim(grdList.TextMatrix(grdList.row, 14)) <> "" Then
                        For nCol = 1 To grdList.Cols - 1
                           grdList.col = nCol
                           If grdList.CellBackColor <> &H8000& Then: grdList.CellBackColor = &H8000& '綠色
                           If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
                        Next nCol
                    Else
                        For nCol = 1 To grdList.Cols - 1
                           grdList.col = nCol
                           If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
                           If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
                        Next nCol
                    End If
            End Select
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
End Sub
