VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100134 
   BorderStyle     =   1  '單線固定
   Caption         =   "臺灣地址郵遞區號查詢"
   ClientHeight    =   5460
   ClientLeft      =   6090
   ClientTop       =   1545
   ClientWidth     =   9135
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9135
   Begin VB.Frame Frame2 
      Height          =   350
      Left            =   4080
      TabIndex        =   23
      Top             =   360
      Width           =   2100
      Begin VB.OptionButton Option1 
         Caption         =   "字首比對"
         Height          =   180
         Index           =   0
         Left            =   72
         TabIndex        =   25
         Top             =   144
         Width           =   1020
      End
      Begin VB.OptionButton Option1 
         Caption         =   "模糊比對"
         Height          =   180
         Index           =   1
         Left            =   1050
         TabIndex        =   24
         Top             =   144
         Value           =   -1  'True
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "解析如下"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1300
      Left            =   60
      TabIndex        =   18
      Top             =   825
      Width           =   3700
      Begin VB.CommandButton cmdQuery 
         Caption         =   "解析查詢"
         Height          =   300
         Index           =   1
         Left            =   2520
         TabIndex        =   13
         Top             =   240
         Width           =   1000
      End
      Begin MSForms.TextBox txtAddr 
         Height          =   300
         Left            =   600
         TabIndex        =   11
         Top             =   960
         Width           =   2955
         VariousPropertyBits=   671105051
         Size            =   "5212;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCity 
         Height          =   300
         Left            =   855
         TabIndex        =   7
         Top             =   240
         Width           =   1425
         VariousPropertyBits=   671105051
         Size            =   "2514;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtStreet 
         Height          =   300
         Left            =   855
         TabIndex        =   9
         Top             =   600
         Width           =   2685
         VariousPropertyBits=   671105051
         Size            =   "4736;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         Caption         =   "地址："
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   1005
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "縣市："
         Height          =   195
         Index           =   23
         Left            =   360
         TabIndex        =   20
         Top             =   300
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "路/街名："
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Top             =   660
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   345
      Left            =   7620
      TabIndex        =   4
      Top             =   30
      Width           =   855
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   6720
      TabIndex        =   3
      Top             =   30
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm100134.frx":0000
      Height          =   3000
      Left            =   60
      TabIndex        =   5
      Top             =   2370
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   5292
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "郵遞區號|城市|區/市/鄉/鎮|路／街名|單/雙|號碼|起始巷|起始弄|起始號"
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
      _Band(0).Cols   =   9
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Top             =   45
      Width           =   1530
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2699;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   315
      Left            =   975
      TabIndex        =   0
      Top             =   45
      Width           =   1530
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2699;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtRoad 
      Height          =   300
      Left            =   975
      TabIndex        =   2
      Top             =   405
      Width           =   3030
      VariousPropertyBits=   671105051
      Size            =   "5345;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label10 
      Caption         =   "路/街名："
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   480
      Width           =   1200
   End
   Begin VB.Label Label9 
      Caption         =   "區/市/鄉/鎮："
      Height          =   195
      Left            =   2640
      TabIndex        =   21
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label8 
      Caption         =   "● 釣魚臺列嶼"
      Height          =   300
      Left            =   4725
      TabIndex        =   17
      Top             =   2040
      Width           =   4380
   End
   Begin VB.Label Label4 
      Caption         =   "● 其他縣市格式：XX 縣 XX 市(鄉)(鎮)"
      Height          =   300
      Left            =   4725
      TabIndex        =   16
      Top             =   1725
      Width           =   4380
   End
   Begin VB.Label Label3 
      Caption         =   "● 基隆市、新竹市、嘉義市，格式為 XX 市 XX 區"
      Height          =   300
      Left            =   4725
      TabIndex        =   15
      Top             =   1395
      Width           =   4380
   End
   Begin VB.Label Label2 
      Caption         =   "● 直轄市：臺北市 、新北市、桃園市、臺中市、臺南                         市、高雄市，格式為 XX 市 XX 區"
      Height          =   450
      Index           =   0
      Left            =   4725
      TabIndex        =   14
      Top             =   930
      Width           =   4380
   End
   Begin VB.Label Label1 
      Caption         =   "地址格式："
      Height          =   300
      Index           =   0
      Left            =   3840
      TabIndex        =   12
      Top             =   930
      Width           =   1005
   End
   Begin VB.Label Label7 
      Caption         =   "縣市："
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   105
      Width           =   600
   End
   Begin VB.Label Label6 
      Caption         =   "查詢後點2下,可將資料帶回前畫面"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   105
      TabIndex        =   8
      Top             =   2160
      Width           =   3105
   End
   Begin VB.Label Label5 
      Caption         =   "（請勿輸鄰、里）"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   6180
      TabIndex        =   6
      Top             =   525
      Width           =   1605
   End
End
Attribute VB_Name = "frm100134"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已檢查 (無需修改的物件)
'Create By Sindy 2015/3/20
Option Explicit

'Add by Amy 2015/09/11
Public strPrevFormMon As String '前畫面的母層表單名 for 強制表單用
Public BFormStatus As String '前畫面目前狀態
Public BFormZip As String '前畫面Zip欄位名稱
Dim intQuery As Integer '查詢次數
Dim m_PrevForm As Form '前畫面
Dim strROC As String, strIndArea As String '有輸中華民國文字/有輸xx工業區
Dim stOldAddr As String, stOldCity As String, stOldArea As String '前畫面地址/縣市/地區 for 選到資料與前畫面不同ex:新北市五工一路(新莊區/五股區)
Dim strCity As String, strStreet As String
Dim strST06 As String 'Add by Amy 2018/05/07 所別


Private Sub cmdExit_Click()
   Unload Me
End Sub

Public Function QueryData(Optional ByVal intCmd As Integer = 1) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String

    strSql = MsgText(601)
    'Add by Amy 2018/05/07 判斷必輸欄位
    If intCmd = 0 Then
        If Combo1 = MsgText(601) Then
            MsgBox "縣市 不可空白！", vbExclamation
            Combo1.SetFocus
            Exit Function
        End If
        If txtRoad = MsgText(601) Then
            MsgBox "路/街名 不可空白！", vbExclamation
            txtRoad.SetFocus
            Exit Function
        End If
        If Combo1 <> MsgText(601) Then strSql = strSql & " And pzd02='" & Combo1 & "' "
        If Combo2 <> MsgText(601) Then strSql = strSql & " And pzd03='" & Combo2 & "' "
        If txtRoad <> MsgText(601) Then
            If Option1(1).Value = True Then
                strSql = strSql & " And InStr(pzd04,'" & ReplaceTWNo(txtRoad) & "')>0 "
            Else
                strSql = strSql & " And pzd04='" & ReplaceTWNo(txtRoad) & "' "
            End If
        End If
    '由其他畫面進來
    Else
        If Trim(txtCity) = "" Then
          MsgBox "請輸入城市，不可空白！", vbExclamation
          txtCity.SetFocus
          txtCity_GotFocus
          Exit Function
       Else
          txtCity = ReplaceAddr(txtCity, 3)
       End If
       If Trim(txtStreet) = "" Then
          MsgBox "請輸入路/街名，不可空白！", vbExclamation
          txtStreet.SetFocus
          txtStreet_GotFocus
          Exit Function
       End If
       strSql = strSql & " And InStr(pzd02||pzd03,Replace('" & ChgSQL(txtCity) & "','巿','市'))>0 "
       'Modify by Amy 2015/11/27 +pzd02||pzd03拆成兩個欄位顯示,且+pzd11國籍,rtrim(字串+半型空白) 避免造字當掉
       strSql = strSql & " And InStr(pzd04,RTrim('" & ReplaceTWNo(ChgSQL(txtStreet)) & " '))>0 "
    End If

    intQuery = intQuery + 1
'   'Modify by Amy 2015/09/11 同區多zip查詢時限制zip(如:中山路查新莊時(不帶區)會出現金山資料)
'   If stMZip <> MsgText(601) Then
'        strSql = "And  SubStr(pzd01,1,3) in (" & PUB_ChgNumeralStyle(stMZip) & ")"
'   End If
   strSql = "Select pzd01,pzd02,pzd03,pzd04,pzd05,pzd06,pzd07,pzd08,pzd09,pzd11" & _
                " From postzipdata Where " & Mid(strSql, 6) & _
                " Order by pzd01 asc"
   'end 2018/05/07
   If rsTmp.State <> adStateClosed Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Set GRD1.Recordset = rsTmp
   SetGrd
   If rsTmp.RecordCount = 0 Then
      GRD1.Rows = 2
      GRD1.row = 1
      GRD1.col = 0
      MsgBox "查無資料！", vbOKOnly, "查詢資料"
      QueryData = False
   Else
      QueryData = True
   End If
 
EXITSUB:
   Set rsTmp = Nothing

End Function

'Add by Amy 2016/01/04 依條件查詢郵遞區號檔並回傳筆數
'intQuery:1-縣市+區、鄉、鎮+路名/2-縣、市+路名/3-只傳路/街名
Public Function CountQuery(intQuery As Integer, Optional ByRef stZip As String = "", Optional ByRef stCountry As String = "") As Integer
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strWhere As String, strTemp(1) As String

    Select Case intQuery
        Case 1
            strWhere = " And InStr(pzd02||pzd03,Replace('" & ChgSQL(strCity) & "','巿','市'))>0"
        Case 2
            strWhere = " And InStr(pzd02,Replace('" & ChgSQL(strCity) & "','巿','市'))>0"
    End Select
    strQ = "SubStr(Pzd01,1,3) as Pzd01 "
    If stCountry <> MsgText(601) Then strQ = strQ & ",Pzd11 "
   
    strQ = "Select Distinct " & strQ & _
                "From postzipdata Where InStr(pzd04,rtrim('" & ReplaceTWNo(ChgSQL(strStreet)) & " '))>0 " & strWhere & _
                " Order by pzd01 asc"
    If RsQ.State <> adStateClosed Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
   
    If RsQ.RecordCount > 0 And (stZip <> MsgText(601) Or stCountry <> MsgText(601)) Then
        Do While Not RsQ.EOF
            If stZip <> MsgText(601) Then strTemp(0) = strTemp(0) & "," & PUB_ChangeZIPToSir("" & RsQ.Fields("Pzd01"))
            If stCountry <> MsgText(601) Then strTemp(1) = strTemp(1) & "," & "" & RsQ.Fields("Pzd011")
            RsQ.MoveNext
        Loop
    End If
    stZip = Mid(strTemp(0), 2)
    stCountry = Mid(strTemp(1), 2)
    CountQuery = RsQ.RecordCount
    
EXITSUB:
   Set RsQ = Nothing

End Function

Private Sub cmdQuery_Click(Index As Integer)
    'Modify by Amy 2015/11/18 因切割地址時可能切不正確,故增加解析鈕查詢
    Select Case Index
        Case 0
            'Modify by Amy 2018/05/07 重轉郵局資料,修改顯示頁面
            'If Trim(txtAddr) = MsgText(601) Then MsgBox "請輸入地址！", MsgText(5): Exit Sub
            'Modify by Amy 2015/11/11 +切割街/路
            'GetStreet txtAddr, 2
            QueryData (Index)
            'end 2018/05/07
        Case 1
            txtCity = ReplaceAddrTW(txtCity)
            txtStreet = ReplaceAddrTW(txtStreet) '取代台灣大道的台
            QueryData
    End Select
    'end 2015/11/18
End Sub

'Add by Amy 2018/05/07
Private Sub Combo1_Click()
    If Combo1 = MsgText(601) Then Exit Sub
    SetCombo2
End Sub

Private Sub Form_Load()
    'Add by Amy 2015/09/11
    If BFormZip <> MsgText(601) Then cmdExit.Caption = "回前畫面": cmdExit.Width = 1200
    MoveFormToCenter Me
    SetGrd
    'Add by Amy 2018/05/07 選單依所別排序
    strST06 = PUB_GetST06(strUserNum)
    SetCombo1
    Option1(0).Value = True '預設字首比對
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    'Add by Amy 20415/09/11
    If BFormZip <> MsgText(601) Then
        If UCase(m_PrevForm.Name) = "FRM210101_2" Then
            txtCity = Empty
            txtStreet = Empty
            GRD1.Clear
            SetGrd
            GRD1.Rows = 2
            GRD1.row = 1
            BFormZip = MsgText(601)
            For Each frm In Forms
                If frm.Name = strPrevFormMon Then
                    frm.bolBack = True
                    Me.Hide
                    frm.Show
                End If
            Next
        Else
            m_PrevForm.Show
        End If
    End If
    BFormStatus = MsgText(601)
    BFormZip = MsgText(601)
    Set m_PrevForm = Nothing
    'end 2015/09/11
    Set frm100134 = Nothing
End Sub

'Add by Amy 2015/09/11
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

'Modify by Amy 2016/05/26 查詢後雙擊選取資料可帶郵遞區號回前畫面
Private Sub GRD1_DblClick()
    Dim frm As Form
    Dim strBackVal As String, stNewArea As String
    
    If BFormZip = MsgText(601) Then Exit Sub
   
    If GRD1.row <> 0 Then
        GRD1.col = 2
         stNewArea = GRD1.TextMatrix(GRD1.row, 1) & GRD1.Text
         
        GRD1.col = 0
        strBackVal = PUB_ChangeZIPToSir(GRD1.Text)
        Select Case UCase(m_PrevForm.Name)
            Case "FRM090801", "FRM090801_NEW" 'Modify By Sindy 2022/9/16 +, "FRM090801_NEW"
                Select Case UCase(BFormZip)
                    Case "TEXT1(25)"
                        m_PrevForm.Text1(25) = strBackVal
                        m_PrevForm.BackIntField = 1
                    Case "TEXT1(41)"
                        m_PrevForm.Text1(41) = strBackVal
                        m_PrevForm.BackIntField = 3
                    Case "TEXT1(57)"
                        m_PrevForm.Text1(57) = strBackVal
                        m_PrevForm.BackIntField = 5
                    Case "TEXT1(73)"
                        m_PrevForm.Text1(73) = strBackVal
                        m_PrevForm.BackIntField = 7
                    Case "TEXT1(89)"
                        m_PrevForm.Text1(89) = strBackVal
                        m_PrevForm.BackIntField = 9
                    Case "TEXT1(120)"
                        m_PrevForm.Text1(120) = strBackVal
                        m_PrevForm.BackIntField = 2
                    Case "TEXT1(121)"
                        m_PrevForm.Text1(121) = strBackVal
                        m_PrevForm.BackIntField = 4
                    Case "TEXT1(122)"
                        m_PrevForm.Text1(122) = strBackVal
                        m_PrevForm.BackIntField = 6
                    Case "TEXT1(123)"
                        m_PrevForm.Text1(123) = strBackVal
                        m_PrevForm.BackIntField = 8
                    Case "TEXT1(124)"
                        m_PrevForm.Text1(124) = strBackVal
                        m_PrevForm.BackIntField = 10
                End Select
                '同縣市區別錯誤,自動更正
                If stOldCity & stOldArea <> MsgText(601) And stOldCity & stOldArea <> stNewArea Then
                    If Right(stOldArea, 3) = "XX區" Then
                        If (strIndArea = "科學工業園區" Or strIndArea = "科學園區") And Left(stOldCity, 2) = "新竹" And InStr(strIndArea, "新竹") = 0 Then strIndArea = "新竹" & strIndArea
                        stOldAddr = strROC & stNewArea & IIf(strIndArea = "True", "", strIndArea) & Mid(stOldAddr, Len(stOldCity) + 1)
                    Else
                        stOldAddr = strROC & Replace(stOldAddr, stOldCity & stOldArea, stNewArea)
                    End If
                    Select Case UCase(BFormZip)
                        Case "TEXT1(25)"
                            If m_PrevForm.Text1(26).Enabled = True Then m_PrevForm.Text1(26) = stOldAddr
                        Case "TEXT1(41)"
                             If m_PrevForm.Text1(42).Enabled = True Then m_PrevForm.Text1(42) = stOldAddr
                        Case "TEXT1(57)"
                            If m_PrevForm.Text1(58).Enabled = True Then m_PrevForm.Text1(58) = stOldAddr
                        Case "TEXT1(73)"
                            If m_PrevForm.Text1(74).Enabled = True Then m_PrevForm.Text1(74) = stOldAddr
                        Case "TEXT1(89)"
                            If m_PrevForm.Text1(90).Enabled = True Then m_PrevForm.Text1(90) = stOldAddr
                        Case "TEXT1(120)"
                            If m_PrevForm.Text1(27).Enabled = True Then m_PrevForm.Text1(27) = stOldAddr
                        Case "TEXT1(121)"
                            If m_PrevForm.Text1(43).Enabled = True Then m_PrevForm.Text1(43) = stOldAddr
                        Case "TEXT1(122)"
                            If m_PrevForm.Text1(59).Enabled = True Then m_PrevForm.Text1(59) = stOldAddr
                        Case "TEXT1(123)"
                            If m_PrevForm.Text1(75).Enabled = True Then m_PrevForm.Text1(75) = stOldAddr
                        Case "TEXT1(124)"
                            If m_PrevForm.Text1(91).Enabled = True Then m_PrevForm.Text1(91) = stOldAddr
                    End Select
                    If Val(m_PrevForm.BackIntField) <> 0 Then m_PrevForm.Tag = stOldAddr
                End If
            Case "FRM140401"
                If (Val(BFormStatus) = 1 Or Val(BFormStatus) = 2) Then
                    Select Case UCase(BFormZip)
                        Case "TEXTCU30"
                            If m_PrevForm.textCU30 <> strBackVal Then
                                If m_PrevForm.textCU30 <> MsgText(601) Then MsgBox "聯絡地址郵遞區號有誤,系統將自動更正！", , MsgText(5)
                                m_PrevForm.textCU30 = strBackVal
                            End If
                            If m_PrevForm.textCU87 <> GRD1.TextMatrix(GRD1.row, 9) Then
                                If m_PrevForm.textCU87 <> MsgText(601) Then MsgBox "聯絡地址國籍有誤,系統將自動更正！", , MsgText(5)
                                m_PrevForm.textCU87 = GRD1.TextMatrix(GRD1.row, 9) '國籍
                                m_PrevForm.textCU87_Validate (False)
                            End If
                        Case "TEXTCU112"
                            If m_PrevForm.textCU112 <> strBackVal Then
                                If m_PrevForm.textCU112 <> MsgText(601) Then MsgBox "中文地址郵遞區號有誤,系統將自動更正！", , MsgText(5)
                                m_PrevForm.textCU112 = strBackVal
                            End If
                            If m_PrevForm.textCU10 <> GRD1.TextMatrix(GRD1.row, 9) Then
                                If m_PrevForm.textCU10 <> MsgText(601) Then MsgBox "中文地址國籍有誤,系統將自動更正！", , MsgText(5)
                                m_PrevForm.textCU10 = GRD1.TextMatrix(GRD1.row, 9) '國籍
                                m_PrevForm.textCU10_Validate (False)
                            End If
                    End Select
                    '同縣市區別錯誤,自動更正
                     If stOldCity & stOldArea <> MsgText(601) And stOldCity & stOldArea <> stNewArea Then
                        If Right(stOldArea, 3) = "XX區" Then
                            If (strIndArea = "科學工業園區" Or strIndArea = "科學園區") And Left(stOldCity, 2) = "新竹" And InStr(strIndArea, "新竹") = 0 Then strIndArea = "新竹" & strIndArea
                            stOldAddr = strROC & stNewArea & IIf(strIndArea = "True", "", strIndArea) & Mid(stOldAddr, Len(stOldCity) + 1)
                        Else
                            stOldAddr = strROC & Replace(stOldAddr, stOldCity & stOldArea, stNewArea)
                        End If
                        Select Case UCase(BFormZip)
                            Case "TEXTCU30"
                                m_PrevForm.textCU31 = stOldAddr
                            Case "TEXTCU112"
                                'Modify by Amy 2016/12/29 客戶檔修改時中文地址沒區不加區
                                If Val(BFormStatus) = 1 Then
                                    m_PrevForm.textCU23 = stOldAddr
                                Else
                                    m_PrevForm.textCU23.Tag = stOldAddr
                                End If
                        End Select
                    End If
                End If
            Case "FRM210101_1"
                If m_PrevForm.txtRead(9).Enabled = True Then
                    m_PrevForm.txtRead(9) = strBackVal
                    '同縣市區別錯誤,自動更正
                    If stOldCity & stOldArea <> MsgText(601) And stOldCity & stOldArea <> stNewArea Then
                        If Right(stOldArea, 3) = "XX區" Then
                            If (strIndArea = "科學工業園區" Or strIndArea = "科學園區") And Left(stOldCity, 2) = "新竹" And InStr(strIndArea, "新竹") = 0 Then strIndArea = "新竹" & strIndArea
                            stOldAddr = strROC & stNewArea & IIf(strIndArea = "True", "", strIndArea) & Mid(stOldAddr, Len(stOldCity) + 1)
                        Else
                            stOldAddr = strROC & Replace(stOldAddr, stOldCity & stOldArea, stNewArea)
                        End If
                        m_PrevForm.txtRead(4) = stOldAddr
                    End If
                End If
                If m_PrevForm.txtRead(10) <> GRD1.TextMatrix(GRD1.row, 9) Then
                    If m_PrevForm.txtRead(10) <> MsgText(601) Then MsgBox "地址國籍有誤,系統將自動更正！", , MsgText(5)
                    m_PrevForm.txtRead(10) = GRD1.TextMatrix(GRD1.row, 9) '國籍
                    m_PrevForm.txtRead_LostFocus (10)
                End If
            Case "FRM210101_2"
                If m_PrevForm.txtPCC(21).Enabled = True Then
                    m_PrevForm.txtPCC(21) = strBackVal
                    '同縣市區別錯誤,自動更正
                    If stOldCity & stOldArea <> MsgText(601) And stOldCity & stOldArea <> stNewArea Then
                        If Right(stOldArea, 3) = "XX區" Then
                            If (strIndArea = "科學工業園區" Or strIndArea = "科學園區") And Left(stOldCity, 2) = "新竹" And InStr(strIndArea, "新竹") = 0 Then strIndArea = "新竹" & strIndArea
                            stOldAddr = strROC & stNewArea & IIf(strIndArea = "True", "", strIndArea) & Mid(stOldAddr, Len(stOldCity) + 1)
                        Else
                            stOldAddr = strROC & Replace(stOldAddr, stOldCity & stOldArea, stNewArea)
                        End If
                        m_PrevForm.txtPCC(22) = stOldAddr
                    End If
                    For Each frm In Forms
                        If frm.Name = strPrevFormMon Then
                            Unload Me
                            Exit For
                        End If
                    Next
                    Exit Sub
                End If
        End Select
        Unload Me
    End If
End Sub
'end 2016/05/26

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow GRD1, x, y, nCol, nRow
   If nRow < 0 Then nRow = 0
   GRD1.col = nCol
   GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
Dim tmpMouseRow
Dim i, j
   
   tmpMouseRow = GRD1.row
   If tmpMouseRow <> 0 Then
      GRD1.row = tmpMouseRow
      GRD1.col = 0
      GRD1.Visible = False
      If Trim(GRD1.TextMatrix(tmpMouseRow, 0)) <> "" Then
         If GRD1.CellBackColor = &HFFC0C0 Then '原灰藍變白
            For i = 0 To GRD1.Cols - 1
               GRD1.col = i
               GRD1.CellBackColor = QBColor(15) '白
            Next i
         Else '原白變灰藍
            For i = 0 To GRD1.Cols - 1
               GRD1.col = i
               GRD1.CellBackColor = &HFFC0C0 '灰藍
            Next i
         End If
      End If
      GRD1.Visible = True
   End If
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   'Modfiy by Amy 2015/09/11+國籍不顯示
   'Modify by Amy 2018/05/07 起始巷/起始弄/起始號 不顯示(郵局無此資料)
   arrGridHeadText = Array("郵遞區號", "縣市", "區/市/鄉/鎮", "路/街名", "單/雙", "號碼", "起始巷", "起始弄", "起始號", "國籍")
   arrGridHeadWidth = Array(800, 1000, 1000, 2100, 500, 1600, 0, 0, 0, 0)
   'end 2015/09/11
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub txtAddr_GotFocus()
    TextInverse txtAddr
    OpenIme
End Sub

Private Sub txtAddr_KeyPress(KeyAscii As MSForms.ReturnInteger)
     KeyAscii = ChangeZIP(KeyAscii, txtAddr)
End Sub

Private Sub txtCity_GotFocus()
   TextInverse txtCity
   OpenIme
End Sub

Private Sub txtStreet_GotFocus()
   TextInverse txtStreet
   OpenIme
End Sub

Private Sub txtStreet_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, txtStreet)
End Sub

'Add by Amy 2015/11/11 傳入地址截取街道名
'intMany:0-取過郵遞區號判斷只有一筆 /1-取過郵遞區號判斷為多筆 /2-未取過郵遞區號
Public Sub GetStreet(ByVal strAddr As String, ByVal intMany As Integer, Optional ByRef intGet As Integer = 0, Optional ByRef strZip As String = "", Optional ByVal bolShowForm As Boolean = True)
    Dim strCountry As String, strCityN As String
    Dim bolMany As Boolean
    
    strROC = ""
    If BFormZip <> MsgText(601) Then txtAddr = strAddr
    If Left(strAddr, 4) = "中華民國" Then strROC = strROC & "中華民國"
    If Left(strAddr, 3) = "臺灣省" Or Left(strAddr, 3) = "台灣省" Then
        strROC = strROC & "臺灣省"
    ElseIf Left(strAddr, 2) = "臺灣" Or Left(strAddr, 2) = "台灣" Then
        strROC = strROC & "臺灣"
    End If
    
    '取代 臺 字
    strAddr = ReplaceAddrTW(IIf(strROC <> MsgText(601), Mid(strAddr, Len(strROC) + 1), strAddr))
    
    '取代 xx工業區(ReplaceIndArea不加 臺中工業區/台塑工業園區 因會抓錯zip)
    If InStr(strAddr, "臺中工業區") > 0 Then
        If InStr(strAddr, "市") > 0 Or InStr(strAddr, "縣") > 0 Then
            strIndArea = "臺中工業區"
            strAddr = Replace(strAddr, "臺中工業區", "")
        Else
            strIndArea = "工業區"
            strAddr = Replace(strAddr, "工業區", "")
        End If
    Else
        strIndArea = "True"
        strAddr = ReplaceIndArea(strAddr, strIndArea)
    End If
    If Left(strAddr, 4) = "新竹新竹" And (strIndArea = "科學工業園區" Or strIndArea = "科學園區") Then
        strIndArea = "新竹" & strIndArea
        strAddr = Mid(strAddr, 3)
    End If
    '記錄原始地址 for 選到資料與前畫面不同需取代 ex:新北市五工一路(新莊區/五股區)或舊制鄉/鎮
    stOldAddr = strAddr

    Select Case intMany
        Case 0 '取過郵遞區號判斷只有一筆
            strCity = Mid(strAddr, 1, intGet)
            stOldArea = Mid(stOldAddr, 1, intGet)
        Case 1 '有區/鄉/鎮取過郵遞區號判斷為多筆
            strCity = Mid(strAddr, 1, intGet)
            stOldArea = Mid(stOldAddr, 1, intGet)
        Case 2 '按查詢鈕 進入的
            strAddr = ReplaceAddr(strAddr, 1)  '取代段
            strAddr = ReplaceAddr(strAddr, 3)  '取舊 鄉/鎮/區 名
           
            '第3個字沒 縣/市
            If Mid(strAddr, 3, 1) <> "市" And Mid(strAddr, 3, 1) <> "縣" And Mid(strAddr, 1, 3) <> "釣魚臺" And Mid(strAddr, 1, 3) <> "海南島" Then
                '傳入地址前2個字判斷是否有其縣/市
                strCityN = "Pzd02"
                strZip = GetPostZip(Left(strAddr, 2), 2, 1, , bolMany, "Pzd02", strCityN)
                If strZip <> MsgText(601) Then
                    stOldCity = Left(strAddr, 2)
                    If bolMany = False Then strAddr = strCityN & Mid(strAddr, 3) '只有一筆(新竹、嘉義會有2筆)
                End If
            Else
                '傳入地址前3個字判斷是否有其縣/市
                strCityN = "Pzd02"
                strZip = GetPostZip(Left(strAddr, 3), 3, 1, , bolMany, "Pzd02")
                If strZip <> MsgText(601) Then
                    stOldCity = Left(strAddr, 3)
                End If
            End If
            
            '** 有縣/市字 or 已加縣/市字
            If strZip <> MsgText(601) Then
                strZip = MsgText(601)
                '*** 有縣/市 只有1筆
                If bolMany = False Then
                    'Modify by Amy 2018/12/19 傳入地址前7個字取郵遞區號  ex:嘉義縣阿里山鄉 X80024
                    If Mid(strAddr, 7, 1) = "市" Or Mid(strAddr, 7, 1) = "區" Or Mid(strAddr, 7, 1) = "鄉" Or Mid(strAddr, 7, 1) = "鎮" Then
                        strZip = GetPostZip(Left(strAddr, 7), 7, , strCountry, bolMany)
                        intGet = 7
                    '傳入地址前6個字取郵遞區號
                    ElseIf Mid(strAddr, 6, 1) = "市" Or Mid(strAddr, 6, 1) = "區" Or Mid(strAddr, 6, 1) = "鄉" Or Mid(strAddr, 6, 1) = "鎮" Then
                        strZip = GetPostZip(Left(strAddr, 6), 6, , strCountry, bolMany)
                        intGet = 6
                    '傳入地址前5個字取郵遞區號
                    ElseIf Mid(strAddr, 5, 1) = "市" Or Mid(strAddr, 5, 1) = "區" Or Mid(strAddr, 5, 1) = "鄉" Or Mid(strAddr, 5, 1) = "鎮" Then
                        strZip = GetPostZip(Left(strAddr, 5), 5, , strCountry, bolMany)
                        intGet = 5
                    End If
                    
                    '有抓到 縣市||鄉鎮 相對應的zip
                    If strZip <> MsgText(601) Then
                        stOldArea = Mid(strAddr, 4, IIf(Len(stOldCity) = 2, intGet - 1, intGet - 3))
                        If bolMany = False Then
                            strCity = Mid(strAddr, 1, intGet)  '有ZipCode可以直接切割區/鄉/鎮
                        Else
                            strCity = Mid(strAddr, 1, 3)          '郵遞區號多筆查詢時不帶區/鄉/鎮 查
                        End If
                    '有 區/鄉/鎮 字,但未抓到zip
                    ElseIf intGet > 0 Then
                        '判斷是否有此區/鄉/鎮(ex:新竹園區2路會抓錯)
                         strZip = GetPostZip(Mid(strAddr, Len(stOldCity) + 1, intGet - Len(stOldCity)), intGet - Len(stOldCity), , strCountry, bolMany, "Pzd03")
                        If strZip <> MsgText(601) Then
                            stOldArea = Mid(strAddr, Len(stOldCity) + 1, intGet - Len(stOldCity))
                            strCity = Mid(strAddr, 1, 3)
                            intMany = 3
                        Else
                            intGet = 0
                            intMany = 4
                        End If
                    '沒有區/鄉/鎮
                    Else
                        intMany = 4
                    End If
                '*** 沒縣/市字,但多筆的縣市(新竹、嘉義會有2筆)
                Else
                    Call CityMany(strAddr, strZip, bolMany, intGet, intMany)
                End If
                
            '**  沒此縣/市
            Else
                intMany = 99
            End If
        Case 3 '有區別,但無此區
            strCity = Mid(strAddr, 1, 3)   '只帶縣/市查
            stOldCity = strCity
            'Moidfy by Amy 2018/12/19 +判斷第七個字 ex:嘉義縣阿里山鄉 X80024
            If Mid(strAddr, 7, 1) = "市" Or Mid(strAddr, 7, 1) = "區" Or Mid(strAddr, 7, 1) = "鄉" Or Mid(strAddr, 7, 1) = "鎮" Then
                intGet = 7
            ElseIf Mid(strAddr, 6, 1) = "市" Or Mid(strAddr, 6, 1) = "區" Or Mid(strAddr, 6, 1) = "鄉" Or Mid(strAddr, 6, 1) = "鎮" Then
                intGet = 6
            ElseIf Mid(strAddr, 5, 1) = "市" Or Mid(strAddr, 5, 1) = "區" Or Mid(strAddr, 5, 1) = "鄉" Or Mid(strAddr, 5, 1) = "鎮" Then
                intGet = 5
            End If
            stOldArea = Mid(strAddr, 4, intGet)
        Case 4 '無區別,只有路是多筆
            stOldCity = Mid(strAddr, 1, 3) '只帶縣/市查
        Case 5 '抓到2個字縣市,但多筆
            strCity = Mid(strAddr, 1, 2)   '只帶縣/市查
            stOldCity = strCity
            Call CityMany(strAddr, strZip, bolMany, intGet, intMany)
    End Select
    
    '記錄前畫面傳入的區鄉鎮
    If intGet > 0 Then strAddr = Mid(strAddr, intGet + 1)
     
    If intMany = 2 And intGet = 0 Then
        '未抓到ZipCode
        Call CutStreet(strAddr)
    ElseIf intMany = 4 Then
        '未抓到ZipCode-只有路
        strCity = Mid(strAddr, 1, Len(stOldCity)) '只帶縣/市查
        stOldArea = "XX區"             '沒區帶XX區返回時取代選的
        strAddr = Mid(strAddr, Len(stOldCity) + 1)
        Call CutStreet(strAddr)
    Else
        Call CutStreet(strAddr)
    End If
    If bolShowForm = True Then txtCity = strCity: txtStreet = strStreet
End Sub

Private Sub CutStreet(ByVal strAddr As String)
    Dim strChk1, strChk2
    Dim strTemp As String
    Dim i As Integer, intGet As Integer
    
    strChk1 = Array("大道", "榮總", "市場", "家園", "社區", "山莊", "別莊", "新城", "商場", "工業", "農場", "營區")
    strChk2 = Array("里", "村", "巷", "嶼", "坑", "寮", "厝", "湖", "溪", "潭", "港", "嶺", "腳")

    
    '檢查地址有下列字時需截取的字元
    If InStr(strAddr, "段") > 0 And IsNumeric(Mid(InStr(strAddr, "段") - 1, 1)) Then
         intGet = InStr(strAddr, "段")
    ElseIf InStr(strAddr, "路") > 0 Then
        intGet = InStr(strAddr, "路")
    ElseIf InStr(strAddr, "街") > 0 Then
       intGet = InStr(strAddr, "街")
    Else
        For i = 0 To UBound(strChk1)
            If InStr(strAddr, strChk1(i)) > 0 Then
                intGet = InStr(strAddr, strChk1(i)) + 1
                Exit For
            End If
        Next i
        If intGet = 0 Then
            For i = 0 To UBound(strChk2)
                If InStr(strAddr, strChk2(i)) > 0 Then
                    intGet = InStr(strAddr, strChk2(i))
                Exit For
            End If
            Next i
        End If
    End If
    
    If intGet > 0 Then
        strStreet = Mid(strAddr, 1, intGet)
    ElseIf InStr(strAddr, "號") > 0 Then
        strStreet = Mid(strAddr, 1, InStr(strAddr, "號"))
    Else
        '無法判斷帶全部
        strStreet = strAddr
    End If
    
End Sub

Private Sub CityMany(ByVal strAddr As String, ByRef strZip As String, ByRef bolMany As Boolean, ByRef intGet As Integer, ByRef intMany As Integer)
    Dim strCityN As String
    
    strCityN = "Pzd02"
    'Modify by Amy 2018/12/19 +判斷第6個字 ex:嘉義縣阿里山鄉 X80024
    If Mid(strAddr, 6, 1) = "市" Or Mid(strAddr, 6, 1) = "區" Or Mid(strAddr, 6, 1) = "鄉" Or Mid(strAddr, 6, 1) = "鎮" Then
        intGet = 6
        strZip = GetPostZip(Left(strAddr, 6), 6, , , bolMany, "SubStr(pzd02,1,2)||pzd03", strCityN)
    ElseIf Mid(strAddr, 5, 1) = "市" Or Mid(strAddr, 5, 1) = "區" Or Mid(strAddr, 5, 1) = "鄉" Or Mid(strAddr, 5, 1) = "鎮" Then
        intGet = 5
        strZip = GetPostZip(Left(strAddr, 5), 5, , , bolMany, "SubStr(pzd02,1,2)||pzd03", strCityN)
    ElseIf Mid(strAddr, 4, 1) = "市" Or Mid(strAddr, 4, 1) = "區" Or Mid(strAddr, 4, 1) = "鄉" Or Mid(strAddr, 4, 1) = "鎮" Then
        intGet = 4
        strZip = GetPostZip(Left(strAddr, 4), 4, , , bolMany, "SubStr(pzd02,1,2)||pzd03", strCityN)
    Else
        intMany = 4
    End If
                
    '有抓到 縣市||鄉鎮 相對應的zip
    If strZip <> MsgText(601) Then
        stOldArea = Mid(strAddr, 3, intGet)
        If bolMany = False Then
            strCity = strCityN & Mid(strAddr, 3, intGet)  '有ZipCode可以直接切割區/鄉/鎮
        Else
            strCity = Mid(strAddr, 1, 3)          '郵遞區號多筆or只有路名 查詢時不帶區/鄉/鎮 查
        End If
    '有 區/鄉/鎮 字,但未抓到zip
    ElseIf intGet > 0 Then
        '判斷是否有此區/鄉/鎮
        strZip = GetPostZip(Mid(strAddr, Len(stOldCity) + 1, intGet - Len(stOldCity)), intGet - Len(stOldCity), , , bolMany, "Pzd03")
        If strZip <> MsgText(601) Then
            stOldArea = Mid(strAddr, Len(strCity) + 1, intGet)
            intMany = 3
        Else
            intGet = 0
            intMany = 4
        End If
    '未抓到相對應的zip
    ElseIf intMany <> 4 Then
        intMany = 3
    End If

End Sub
'end 2015/11/11

'Add by Amy 2018/05/07
'依所別排序縣市
Private Sub SetCombo1()
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    strQ = "Select Distinct pzd02,Decode(pzd10,'" & strST06 & "',' ',pzd10)||Decode(pzd02,'臺北市','11',Decode(pzd02,'新北市','12',SubStr(pzd01,1,1))) as pzd01 " & _
               "From PostZipData " & _
                "Order by pzd01 "
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        RsQ.MoveFirst
        Combo1.AddItem ""
        Do While Not RsQ.EOF
            Combo1.AddItem RsQ.Fields("pzd02")
            RsQ.MoveNext
        Loop
    End If
    RsQ.Close
End Sub

'依縣市抓取對應區/市/鄉/鎮
Private Sub SetCombo2()
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    strQ = "Select Distinct pzd03,SubStr(pzd01,1,3) " & _
               "From PostZipData Where pzd02='" & Combo1 & "' " & _
                "Order by SubStr(pzd01,1,3) "
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    Combo2.Clear 'Add By Sindy 2021/12/16
    If intQ = 1 Then
        RsQ.MoveFirst
        Combo2.AddItem ""
        Do While Not RsQ.EOF
            Combo2.AddItem RsQ.Fields("pzd03")
            RsQ.MoveNext
        Loop
    End If
    RsQ.Close
End Sub

