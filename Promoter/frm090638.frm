VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090638 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "商標未發文原因註記"
   ClientHeight    =   4980
   ClientLeft      =   1695
   ClientTop       =   3105
   ClientWidth     =   8760
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8760
   Begin VB.CommandButton cmdSelect 
      Caption         =   "全選"
      Height          =   345
      Left            =   4450
      TabIndex        =   11
      Top             =   10
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   1
      Top             =   480
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "顯示未註記"
      Height          =   345
      Left            =   6585
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   10
      Width           =   1300
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "原因註記"
      Height          =   345
      Left            =   5265
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   15
      Width           =   1300
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢"
      Default         =   -1  'True
      Height          =   350
      Left            =   2130
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   495
      Width           =   550
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   7900
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   10
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   3795
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   6694
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
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
   End
   Begin MSForms.ComboBox CboEmp 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1635
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2884;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "目前狀態：全部"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   6480
      TabIndex        =   10
      Top             =   600
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "管制年月："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   510
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "　承辦人："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   150
      Width           =   900
   End
   Begin VB.Label lbl1 
      Height          =   210
      Index           =   1
      Left            =   2640
      TabIndex        =   3
      Top             =   2670
      Width           =   3270
   End
End
Attribute VB_Name = "frm090638"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/22 改成Form2.0 ; CboEmp、GrdDataList改字型=新細明體-ExtB
'Memo by Lydia 2019/07/01 表單名稱:商標未發文案件原因註記=>商標未發文原因註記
'Create by Amy 2015/09/04
Option Explicit

Dim i As Integer
Public intPeople As Integer '操作人身份 1:承辦人 2.智權人員
Dim m_QueryType As String '1:所有未發文 2.未填寫資料
Dim stST05 As String, stST15 As String
Dim stSalesArea As String, stSalesArea1 As String '業務區
Public strVCP09 As String 'Modify by Amy 2016/03/03
'Dim bolOnlySelf As Boolean '只能看自己
Dim stIdList As String
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
'Add by Amy 2015/10/01
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
'Add by Amy 2016/01/06
Dim intSpecSet As Integer '特殊設定:1-特殊人員ID /2-ST05特殊設定 /3-區主管/4.特殊設定檔/5.帶人主管/6.SalesList人員

Private Sub Form_Load()
Dim stST52List As String '帶人主管所帶人員list
   
   MoveFormToCenter Me
   If intPeople = 2 Then
       Label1(0).Caption = "智權人員："
       stST15 = PUB_GetStaffST15(strUserNum, 1)
   Else
       stST15 = Pub_StrUserSt03
   End If
   SetDataListWidth
   
   stST05 = PUB_GetST05(strUserNum)
   
   'Modify By Sindy 2020/7/28 設定員編,部門,所別權限
   Call PUB_SetFormSaleDept(strUserNum, , , , , bolSpecMan, strSpecCode, stSalesArea, stSalesArea1, , , intSpecSet)
   
'    Select Case strUserNum
'        '副總預設所有智權人員
'        Case "68006"
'            stSalesArea = "S"
'            stSalesArea1 = "S99"
'            intSpecSet = 1
'        '杜燕文,劉大愛可看S31
'        'modify by sonia 2016/8/22劉大愛78007改為蘇嫄媛79053
'        Case "74018", "79053"
'            stSalesArea = "S31"
'            stSalesArea1 = "S31"
'            intSpecSet = 1
'        '王協理可看專利處
'        Case "71011"
'            stSalesArea = "P10"
'            stSalesArea1 = "P19"
'            intSpecSet = 1
'        '葉經理可看商標處
'        Case "67002", "69008"
'            stSalesArea = "P20"
'            stSalesArea1 = "P29"
'            intSpecSet = 1
'        '外商陳經理可看外商
'        Case "68005"
'            stSalesArea = "F10"
'            stSalesArea1 = "F19"
'            intSpecSet = 1
'        Case Else
'            Select Case stST05
'                '電腦中心,財務,總經理,主任秘書(等級08)看全部
'                Case "00", "01", "08"
'                    If intPeople = 1 Then
'                        '商標承辦
'                        stSalesArea = "F1"
'                        stSalesArea1 = "P2"
'                    Else
'                        '智權
'                        stSalesArea = "S"
'                        stSalesArea1 = "S99"
'                    End If
'                    intSpecSet = 2
'                'Mark by Amy 2016/01/06 各區主管,因有承辦資料不可抓st05
''                Case "SM"
''                    '73009可看中所全部,
''                    If strUserNum = "73009" Then
''                        stSalesArea = "S20"
''                        stSalesArea1 = "S29"
''                    '69010可看中南高全部及自己區(SetCboEmp設定)的
''                    ElseIf strUserNum = "69010" Then
''                        stSalesArea = "S20"
''                        stSalesArea1 = "S49"
''                    '簡協理可看北所全部
''                    ElseIf strUserNum = "69005" Then
''                        stSalesArea = "S10"
''                        stSalesArea1 = "S19"
''                    Else
''                        stSalesArea = stST15
''                        stSalesArea1 = stST15
''                    End If
'                '其他判斷特殊權限智權人員清單,非特殊則只能看自己
'                Case Else
'                    stSalesArea = stST15
'                    stSalesArea1 = stST15
'
'                    'Add by Amy 2016/01/06 各區主管 抓A0908
'                    If GetDeptMan(stST15) = strUserNum Then
'                        intSpecSet = 3
'                        '73009可看中所全部,
'                        If strUserNum = "73009" Then
'                            stSalesArea = "S20"
'                            stSalesArea1 = "S29"
'                        '69010可看中南高全部及自己區(SetCboEmp設定)的
'                        ElseIf strUserNum = "69010" Then
'                            stSalesArea = "S20"
'                            stSalesArea1 = "S49"
'                        '簡協理可看北所全部
'                        ElseIf strUserNum = "69005" Then
'                           'Modified by Morgan 2019/12/30 杜主秘說加開簡協理69005可看全所
'                            'stSalesArea = "S10"
'                            'stSalesArea1 = "S19"
'                            stSalesArea = "S"
'                            stSalesArea1 = "S99"
'                            intSpecSet = 1
'                            'end 2019/12/30
'                        End If
'                    '總經理業務工作代理人員
'                    ElseIf CheckLevel(strUserNum, "總經理業務工作代理人員") = True Then
'                        bolSpecMan = True: intSpecSet = 4
'                        strSpecCode = "總經理業務工作代理人員"
'                    '特殊設定A8人員可操作專利處部份智權同仁(A7)
'                    ElseIf CheckLevel(strUserNum, "A8") = True Then
'                        bolSpecMan = True: intSpecSet = 4
'                        strSpecCode = "A8"
'                    Else
'                        stIdList = PUB_GetSalesList(strUserNum, stSalesArea, stSalesArea1)
'                        stST52List = GetST52List(strUserNum)
'                        If (stIdList = MsgText(601) Or sCount(Replace(stIdList, "'", ""), strUserNum, ",") = 1) And _
'                           stST52List = MsgText(601) Then
'                            bolOnlySelf = True
'                        ElseIf stIdList <> MsgText(601) Then
'                            intSpecSet = 6
'                        End If
'                        If stST52List <> MsgText(601) Then stIdList = stIdList & "," & stST52List: intSpecSet = 5
'                    End If
'            End Select
'    End Select
    
   Select Case stST05
      '電腦中心,財務,總經理,主任秘書(等級08)看全部
      Case "00", "01", "08"
          If intPeople = 1 Then
              '商標承辦
              stSalesArea = "F1"
              stSalesArea1 = "P2"
          Else
              '智權
              stSalesArea = "S"
              stSalesArea1 = "S99"
          End If
          intSpecSet = 2
   End Select
   
   Text1 = Left(strSrvDate(1), 6) - 191100
    
'    CboEmp.Clear
'    'Modify by Amy 2016/01/07 自已檔月沒資料也要出現自己名字下拉
'    If bolOnlySelf = True And bolSpecMan = False And InStr(stIdList, ",") = 0 Then
'        CboEmp.AddItem strUserNum & " " & StaffQuery(strUserNum)
'        CboEmp = strUserNum & " " & StaffQuery(strUserNum)
'        CboEmp.Enabled = False
'    Else
        SetCboEmp
'    End If
   '抓出個人當月未發文原因中仍在管制中的資料
   If cboEmp <> MsgText(601) Then QueryData
End Sub

'Add by Amy 2015/10/01 人員選擇完自動執行查詢
Private Sub CboEmp_Click()
    Dim bCancel As Boolean
    
    If cboEmp = MsgText(601) Then Exit Sub
    
    If Trim(Text1) <> MsgText(601) Then
        'Modify by Amy 2016/01/06
        cboEmp.Tag = "Y"
        Call Text1_Validate(bCancel)
        cboEmp.Tag = MsgText(601)
        'end 2016/01/06
        If bCancel = True Then Exit Sub
        Screen.MousePointer = vbHourglass
        cmdSelect.Caption = "全選"
        QueryData
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdMod_Click()
   PubShowNextData
End Sub

Private Sub cmdOK_Click()
    If cmdOK.Caption = "顯示全部" Then
        m_QueryType = 1
        cmdOK.Caption = "顯示未註記"
        Label3.Caption = Replace(Label3.Caption, "未註記", "全部")
    Else
        m_QueryType = 2
        cmdOK.Caption = "顯示全部"
        Label3.Caption = Replace(Label3.Caption, "全部", "未註記")
    End If
    Screen.MousePointer = vbHourglass
    QueryData
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
    Dim bCancel As Boolean
    
    'Modify by Amy 2016/01/06
    cboEmp.Tag = "Y"
    Call Text1_Validate(bCancel)
    cboEmp.Tag = MsgText(601)
    'end 2016/01/06
    If bCancel = True Then Exit Sub
    Screen.MousePointer = vbHourglass
    cmdSelect.Caption = "全選"
    QueryData
    Screen.MousePointer = vbDefault
End Sub

'Add by Amy 2015/10/01 +全選/全取消 鈕
Private Sub cmdSelect_Click()
    Dim j As Integer
    
    grdDataList.Visible = False
    If cmdSelect.Caption = "全選" Then
        cmdSelect.Caption = "全取消"
        For i = 1 To grdDataList.Rows - 1
            grdDataList.TextMatrix(i, 0) = "V"
            For j = 0 To grdDataList.Cols - 1
                grdDataList.row = i
                grdDataList.col = j
                grdDataList.CellBackColor = &HFFC0C0
            Next j
        Next i
    Else
        cmdSelect.Caption = "全選"
        For i = 1 To grdDataList.Rows - 1
            grdDataList.TextMatrix(i, 0) = ""
            For j = 0 To grdDataList.Cols - 1
                grdDataList.row = i
                grdDataList.col = j
                grdDataList.CellBackColor = QBColor(15)
            Next j
        Next i
    End If
     grdDataList.Visible = True
End Sub

'設定承辦人/智權人員下拉式選單
Private Sub SetCboEmp()
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
'    Dim bolHasSelf As Boolean '下拉選單有自己
    Dim strWhere As String 'Add by Amy 2016/01/06
   
    '總經理業務工作代理人員
    If InStr(strSpecCode, "總經理業務工作代理人員") > 0 Then
        strWhere = " And st01='94007' "
    '特殊設定A8人員可操作專利處部份智權同仁(A7)
    ElseIf InStr(strSpecCode, "A8") > 0 Then
         strWhere = " And st01 in ('" & Replace(Pub_GetSpecMan("A7"), ";", "','") & "') "
    '電腦中心,財務,總經理,主任秘書(等級08)看全部商標承辦(從商標->承辦人作業進入)
    ElseIf stSalesArea = "F1" And stSalesArea1 = "P2" Then
         strWhere = "And SubStr(st03,1,2) in ('" & stSalesArea & "','" & stSalesArea1 & "')"
    'Add by Amy 2015/10/07 蘇特助可看自己區及中南高全部
    ElseIf strUserNum = "69010" Then
        strWhere = " And (st15='" & stST15 & "' Or (st15>='" & stSalesArea & "' And st15<='" & stSalesArea1 & "')) "
    ElseIf intSpecSet > 0 And intSpecSet < 5 Then
        strWhere = "And st15>='" & stSalesArea & "' And st15<='" & stSalesArea1 & "' "
    Else
        stIdList = PUB_GetSalesList(strUserNum, stSalesArea, stSalesArea1)
        
        'Modify by Amy 2016/01/07stIdList人員另外抓,因可能跨區
        'strWhere = "And st15>='" & stSalesArea & "' And st15<='" & stSalesArea1 & "' "
        If InStr(stIdList, ",") = 0 Then
            If stIdList = MsgText(601) Then stIdList = strUserNum
            strWhere = " And st01 =" & stIdList & ""
        Else
            strWhere = " And st01 in (" & stIdList & ") "
        End If
        '外商主管  王宗珮、洪琬姿、葉易雲
        If stST05 = "21" Or stST05 = "26" Or stST05 = "28" Then strWhere = strWhere & " And st16=(Select st16 From Staff where st01='" & strUserNum & "') "
    End If
    'Modify by Amy 2016/01/07離職人員下拉選單沒有但未發文原因仍有資料,離職人員需出現
    'strQ = "Select st01,st02 From Staff Where st04='1' And st01>'6' And SubStr(st01,1,1)<'F' And SubStr(st01,4,1)<>'9' " & strQ & " Order by st03,st01"
    'Modified by Lydia 2019/08/08 +創新業務部W1001或W2001
    '    strQ = "Select Distinct st01||' '||st02 as P,st03,st01 From Staff,NotComplete,CaseProgress Where NC01=CP09(+) And NC02=" & Val(Text1) + 191100 & _
                " And " & IIf(intPeople = 2, "CP13", "CP14") & "=st01 And st01>'6' And SubStr(st01,1,1)<'F' And SubStr(st01,4,1)<>'9' " & strWhere & _
                " Order by st03,st01"
    strWhere = "and ((st01>'6' And SubStr(st01,1,1)<'F' And SubStr(st01,4,1)<>'9') or substr(st01,1,1)='W') " & strWhere
    strQ = "Select Distinct st01||' '||st02 as P,st03,st01 From Staff,NotComplete,CaseProgress Where NC01=CP09(+) And NC02=" & Val(Text1) + 191100 & _
            " And " & IIf(intPeople = 2, "CP13", "CP14") & "=st01 " & strWhere & _
            " And st01<>'" & strUserNum & "'"
    strQ = strQ & " Order by st03,st01"
    
    cboEmp.Clear 'Add By Sindy 2020/7/29
'    '當月沒資料,下拉選單中也要有自己
'    If Pub_StrUserSt03 <> "M51" Then
'        CboEmp.AddItem strUserNum & " " & StaffQuery(strUserNum)
'        CboEmp = strUserNum & " " & StaffQuery(strUserNum)
'    End If
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    'Add By Sindy 2020/7/29
    If RsQ.RecordCount > 0 Then
    '2020/7/29 END
      cboEmp.AddItem ""
      cboEmp.ListIndex = 0
    'Add By Sindy 2020/7/29
    End If
    cboEmp.AddItem strUserNum & " " & StaffQuery(strUserNum)
    '2020/7/29 END
    If RsQ.RecordCount > 0 Then
        RsQ.MoveFirst
        Do While Not RsQ.EOF
'            If strUserNum = Left(RsQ.Fields("P"), 5) Then
'                'bolHasSelf = True
'            Else
                cboEmp.AddItem RsQ.Fields("P")
'            End If
            RsQ.MoveNext
        Loop
    'Add By Sindy 2020/7/29
    Else
      cboEmp.Enabled = False
      '2020/7/29 END
    End If
    
    cboEmp = strUserNum & " " & StaffQuery(strUserNum)
'    'If bolHasSelf = True And Pub_StrUserSt03 <> "M51" Then cboEmp = strUserNum & " " & StaffQuery(strUserNum)
'    If Pub_StrUserSt03 <> "M51" Then CboEmp = strUserNum & " " & StaffQuery(strUserNum)
'    'end 2016/01/07
End Sub

Private Sub SetDataListWidth()
    Dim stTitle, intWidth
    Dim strTmp As String
    
    strTmp = "智權人員"
    If intPeople = 2 Then strTmp = "承辦人"
    
    stTitle = Array("V", strTmp, "申請人", "本所案號", "收文日", "申請國家", "案件名稱" _
                         , "案件性質", "專業部原因", "智權部原因", "CP10", "NC01")
    
    intWidth = Array(200, 800, 1000, 1200, 800, 800, 1000 _
                            , 800, 1500, 1500, 0, 0)
                            
    For i = 0 To UBound(stTitle)
        grdDataList.row = 0
        grdDataList.col = i
        grdDataList.Text = stTitle(i)
        grdDataList.ColWidth(i) = intWidth(i)
        grdDataList.CellAlignment = flexAlignLeftCenter
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intPeople = 0
    Set frm090638 = Nothing
End Sub

Private Sub GrdDataList_Click()
    grdDataList.Visible = False
    grdDataList.col = 0
     If grdDataList.row <> 0 Then
        If grdDataList.Text = "V" Then
            grdDataList.Text = ""
            For i = 0 To grdDataList.Cols - 1
                grdDataList.col = i
                grdDataList.CellBackColor = QBColor(15)
            Next i
        Else
            '勾選
            grdDataList.Text = "V"
            For i = 0 To grdDataList.Cols - 1
                grdDataList.col = i
                grdDataList.CellBackColor = &HFFC0C0
            Next i
        End If
     End If
     grdDataList.Visible = True
End Sub

'Add by Amy 2015/10/01 +欄位排序
Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim nCol As Long, nRow As Long
    
    getGrdColRow grdDataList, x, y, nCol, nRow
    If nCol < 0 Or nRow < 0 Then Exit Sub
    
    With Me.grdDataList
        .col = nCol
        .row = nRow
        If .row < 1 And .Text <> "V" Then
            If m_blnColOrderAsc = True Then
                .Sort = 5 '字串昇冪
                m_blnColOrderAsc = False
            Else
                .Sort = 6 '字串昇冪
                m_blnColOrderAsc = True
            End If
        End If
    End With
End Sub

Private Sub Text1_GotFocus()
    TextInverse Text1
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If Text1 = MsgText(601) Then Exit Sub
    
    If ChkDate(Text1 & "01") = False Then
        Text1_GotFocus
        Cancel = True
    End If
    '查非當月預設查所有未發文資料(顯示顯示未註記鈕)
    If Val(Text1) <> Val(Left(strSrvDate(1), 6) - 191100) Then
        cmdOK.Caption = "顯示未註記"
    End If
    'Add by Amy 2016/01/06
    If cboEmp.Tag = MsgText(601) And cboEmp.Enabled = True Then
'        CboEmp.Clear
        SetCboEmp
    End If
    'end 2016/01/06
End Sub

Public Sub QueryData(Optional ByVal bolMsg As Boolean = True)
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strWhere As String, strQuser As String
    
    m_blnColOrderAsc = True
    
    grdDataList.Clear
    strQuser = Left(cboEmp, 5)
    
    If cmdOK.Caption = "顯示全部" Then
        m_QueryType = 2
    Else
        m_QueryType = 1
    End If
    
    '1.承辦人登入
    If intPeople = 1 Then
        strWhere = "And CP14='" & strQuser & "' And CP13=ST01(+) "
        If m_QueryType = 2 Then strWhere = strWhere & "And NC03 is null "
    '2.智權人員登入
    Else
        strWhere = "And CP13='" & strQuser & "' And CP14=ST01(+) "
        If m_QueryType = 2 Then strWhere = strWhere & "And NC07 is null "
    End If
    
    '抓取當月未發文原因仍在管制中的資料
    strQ = "Select '' as V,ST02 as " & IIf(intPeople = 1, "智權人員", "承辦人") & ",Nvl(cu04,Nvl(cu05,cu06)) as 申請人,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,SQLDATET2(CP05) as 收文日, na03 as 申請國家," & _
                "CaseName as 案件名稱,Nvl(Decode(tm10,'000',CPM03,CPM04),cp10) as 案件性質,NC03 as 專業部原因,NC07 as 智權部原因,cp10,nc01 From " & _
                "(Select nc01,nc03,nc07,cp01,cp02,cp03,cp04,cp05, cp10,cp12,cp13,cp14,Decode(tm05,null,Decode(sp05,null,Decode( sp06, null,sp07,sp06),sp05),tm05) as CaseName ,nvl(tm10,sp09) as tm10,nvl(tm23,sp08) as tm23 " & _
                "From NotComplete,CaseProgress,TradeMark ,ServicePractice Where NC02=" & Val(Text1) + 191100 & " And (NC11<>'N' Or NC11 is null) And NC01=CP09(+) " & _
                "And cp01=tm01(+) And cp02=tm02(+) And cp03=tm03(+) And cp04=tm04(+) And cp01=sp01(+) And cp02=sp02(+) And cp03=sp03(+) And cp04=sp04(+) )," & _
                "Staff,Customer,Nation, CasePropertyMap " & _
                "Where SubStr(tm23,1,8)=cu01(+) And SubStr(tm23,9,1)=cu02(+) And tm10=na01(+) And cp01=CPM01(+) And cp10=CPM02(+) " & strWhere & _
                "Order by cp14,cp12,cp13,tm23,cp01,cp02,cp03,cp04,cp05,nc01"
               
    If RsQ.State <> 0 Then RsQ.Close
    
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    SetDataListWidth
    If RsQ.RecordCount = 0 Then
        grdDataList.Rows = 2
        If bolMsg = True Then MsgBox "查無資料！", vbOKOnly, "查詢資料"
    Else
        Set grdDataList.Recordset = RsQ
        grdDataList.Visible = False
    End If
    grdDataList.Visible = True
End Sub

Public Sub PubShowNextData()
    Dim j As Integer
    If Me.Tag = "Save" Then
        QueryData
        SetGrdV
        Me.Tag = MsgText(601)
    End If
    
    Me.Enabled = False
    For i = 1 To grdDataList.Rows - 1
        grdDataList.col = 0
        grdDataList.row = i
        If Trim(grdDataList.Text) = "V" Then
            grdDataList.Text = ""
            For j = 0 To grdDataList.Cols - 1
                grdDataList.col = j
                grdDataList.CellBackColor = QBColor(15)
            Next j
            
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            
            Call frm090638_1.SetParent(Me)
            frm090638_1.BFormPeople = intPeople
            frm090638_1.intModSet = intSpecSet 'Add by Amy 2016/01/07 傳特殊設定
            frm090638_1.m_NC01 = grdDataList.TextMatrix(i, 11)
            frm090638_1.m_NC02 = Text1
            frm090638_1.Show
            If frm090638_1.QueryData = False Then
                ShowNoData
                frm090638_1.cmdState = 2
                frm090638_1.PubShowNextData
            End If
            
            'Modfiy by Amy 2016/03/03 畫面選3筆後跳至存檔畫面存2筆後,跳回此畫面bug修正-陳金蓮
            'If m_QueryType = 2 Then
            GetGrdV '記錄Grid選取的cp09
            Me.Enabled = True
            Exit Sub
        End If
    Next i
    Me.Enabled = True
End Sub

Private Sub GetGrdV()
    Dim strCP09 As String
    
    For i = 1 To grdDataList.Rows - 1
        If Trim(grdDataList.TextMatrix(i, 0)) = "V" Then
            strCP09 = strCP09 & ";" & grdDataList.TextMatrix(i, 11)
        End If
    Next i
    If strCP09 <> MsgText(601) Then strVCP09 = strCP09 'Modfiy by Amy 2016/03/03 Mid(strCP09, 2)
End Sub

Private Sub SetGrdV()
    Dim strCP09() As String
    Dim n As Integer, j As Integer
    
    If strVCP09 = MsgText(601) Then Exit Sub
    
    strCP09 = Split(Mid(strVCP09, 2), ";"): n = 0 'Modfiy by Amy 2016/03/03
    For i = 1 To grdDataList.Rows - 1
        If n > UBound(strCP09) Then Exit For
        If grdDataList.TextMatrix(i, 11) = strCP09(n) Then
            grdDataList.TextMatrix(i, 0) = "V"
            grdDataList.row = i
            For j = 0 To grdDataList.Cols - 1
                If j <> 1 Then
                    grdDataList.col = j
                    grdDataList.CellBackColor = &HFFC0C0
                End If
            Next j
            n = n + 1
        End If
    Next i
End Sub

'計算某文字在字串中出現次數
'stStr:字串/ stFind:尋找字串/ stTag:切割字元
Private Function sCount(ByVal stStr As String, ByVal stFind As String, ByVal stTag As String) As Integer
    Dim stSplit() As String
    Dim intFindCount As Integer
    Dim ii As Integer
    sCount = 0
    If stStr = "" Then Exit Function
    
    stSplit = Split(stStr, stTag)
    For ii = 0 To UBound(stSplit)
        If stSplit(ii) = stFind Then
            intFindCount = intFindCount + 1
        End If
    Next ii
    sCount = intFindCount
 End Function



