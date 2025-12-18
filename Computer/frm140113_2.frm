VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140113_2 
   Caption         =   "教育訓練登錄作業-人員設定"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6072
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   6072
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "刪除"
      Height          =   500
      Index           =   6
      Left            =   5740
      TabIndex        =   18
      Top             =   1620
      Width           =   300
   End
   Begin VB.Frame Frame1 
      Height          =   1000
      Left            =   40
      TabIndex        =   7
      Top             =   570
      Width           =   6000
      Begin VB.CommandButton Command1 
         Caption         =   "全選"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   1665
         TabIndex        =   12
         Top             =   660
         Width           =   1005
      End
      Begin VB.CommandButton Command1 
         Caption         =   "全選"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4800
         TabIndex        =   10
         Top             =   660
         Width           =   780
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3045
         MaxLength       =   6
         TabIndex        =   9
         Top             =   390
         Width           =   950
      End
      Begin VB.CommandButton Command1 
         Caption         =   "加入"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4800
         TabIndex        =   8
         Top             =   330
         Width           =   780
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   315
         Left            =   555
         TabIndex        =   11
         Top             =   180
         Width           =   2145
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3784;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "來源："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   30
         TabIndex        =   17
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "已選人員："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3045
         TabIndex        =   16
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "待選清單："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   30
         TabIndex        =   15
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "編號/名稱："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3060
         TabIndex        =   14
         Top             =   180
         Width           =   1020
      End
      Begin MSForms.Label lbl_Name 
         Height          =   264
         Left            =   4020
         TabIndex        =   13
         Top             =   450
         Width           =   705
         VariousPropertyBits=   27
         Caption         =   "lbl_Name"
         Size            =   "1244;466"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "->"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   2730
      TabIndex        =   3
      Top             =   2910
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3465
      TabIndex        =   1
      Top             =   5040
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4545
      TabIndex        =   0
      Top             =   5040
      Width           =   1005
   End
   Begin MSForms.ListBox List2 
      Height          =   3405
      Left            =   3120
      TabIndex        =   4
      Top             =   1590
      Width           =   2570
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "4533;6006"
      MatchEntry      =   0
      MultiSelect     =   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox List1 
      Height          =   3405
      Left            =   30
      TabIndex        =   2
      Top             =   1590
      Width           =   2655
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "4674;5997"
      MatchEntry      =   0
      MultiSelect     =   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblMemo 
      Height          =   555
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   6000
      ForeColor       =   16711680
      VariousPropertyBits=   27
      Caption         =   "          "
      Size            =   "10583;979"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "下列為各別發信人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   210
      TabIndex        =   5
      Top             =   1380
      Width           =   1755
   End
   Begin VB.Menu mnuPop 
      Caption         =   "彈跳選單"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuPopItem 
         Caption         =   "刪除"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frm140113_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/06 Form2.0已修改 combo3/list1/list2/lbl_Name
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Created by Morgan 2012/5/1
Option Explicit

Public m_stNumList As String
Public fmParent As Form '呼叫的表單
'Add by Amy 2018/09/18
Public bolPublic As Boolean '是公開
Dim m_JoinList As String  '參加人員
Public strDeptNo As String 'Add by Amy 2019/11/12 目前登入者部門
Dim stGroup As String, arrGroup 'Add by Amy 2020/11/27 特殊設定Mail Group

'Add by Amy 2018/09/18 設定參加人員list
'Memo 2020/11/27 不使用,但先保留
Public Sub SetJoinList2(ByRef strJoinList() As String)
    Dim i As Integer, intQ As Integer
    Dim strQ As String, strTmp As String
    
'    List2.Clear '已選人員
'    '參加人員
'    For i = LBound(strJoinList) To UBound(strJoinList)
'        If strJoinList(i) <> MsgText(601) Then
'            m_JoinList = m_JoinList & strJoinList(i)
'        End If
'    Next i
'    'Moidfy by Amy 2019/11/12 原:st03 改st15因文雄需操作S部門-秀玲
'    'Modify by Amy 2020/11/27 +st03
'    strQ = "select st01,st02||'('||decode(st15,'P11','工程師','P13','繪圖','P12','程序','P14','英文顧問',decode(st03,'P11','工程師','P13','繪圖','P12','程序','P14','英文顧問',ST01))||'.'||ac03||')' C02" & _
'                    " from staff,allcode" & _
'                    " where instr('" & m_JoinList & "',st01)>0 and ac02(+)=st20 and ac01(+)='01' order by st02 desc"
'    intQ = 1
'    Set RsTemp = ClsLawReadRstMsg(intQ, strQ)
'    If intQ = 1 Then
'        Do While Not RsTemp.EOF
'            strTmp = ""
'            For i = LBound(strJoinList) To UBound(strJoinList)
'                If strJoinList(i) <> MsgText(601) Then
'                    If InStr(strJoinList(i), RsTemp.Fields("st01")) > 0 Then
'                        strTmp = strTmp & "," & i
'                    End If
'                End If
'            Next i
'            If strTmp <> MsgText(601) Then strTmp = Mid(strTmp, 2)
'            List2.AddItem RsTemp.Fields(1) & "-議題 " & strTmp, 0
'            List2.ItemData(0) = PUB_Id2Num(RsTemp.Fields(0))
'            RsTemp.MoveNext
'        Loop
'    End If
'    RsTemp.Close
'    '人員確認加入之人員
'    If m_stNumList <> MsgText(601) Then
'        'Moidfy by Amy 2019/11/12 原:st03 改st15因文雄需操作S部門-秀玲
'        'Modify by Amy 2020/11/27 +st03
'        strQ = "select st01,st02||'('||decode(st15,'P11','工程師','P13','繪圖','P12','程序','P14','英文顧問',decode(st03,'P11','工程師','P13','繪圖','P12','程序','P14','英文顧問',ST01))||'.'||ac03||')' C02" & _
'                    " from staff,allcode" & _
'                    " where instr('" & m_stNumList & "',st01)>0 and ac02(+)=st20 and ac01(+)='01' order by st02 desc"
'    intQ = 1
'    Set RsTemp = ClsLawReadRstMsg(intQ, strQ)
'    If intQ = 1 Then
'        RsTemp.MoveFirst
'        Do While Not RsTemp.EOF
'            List2.AddItem RsTemp.Fields(1), 0
'            List2.ItemData(0) = PUB_Id2Num(RsTemp.Fields(0))
'            RsTemp.MoveNext
'        Loop
'    End If
'    RsTemp.Close
'    End If
'    SetListScroll List2
'    '2019/01/02 不可選收信人員-經理
'    List1.Visible = True
'    Frame1.Visible = True
'    Command1(2).Visible = True
'    Label2.Visible = False
'    If InStr(Me.Caption, "(收信人員)") > 0 Then
'         List1.Visible = False
'         Frame1.Visible = False
'         List2.Width = 5450
'         List2.Left = 170
'         lblMemo.Height = 1300
'         Command1(2).Visible = False
'         Label2.Visible = True
'    End If
End Sub

Private Sub Combo3_Click()
   SetBookInList
End Sub

Private Sub SetBookInList()
   Dim stCon As String
   Dim stField As String  'Add by Amy 2018/09/18
   
   'Modify by Amy 2018/09/18 增加其他部門顯示,排除虛設編號(第4碼為9)及巨京(st06='5')
   List1.Clear
   'Moidfy by Amy 2019/11/12 原:Pub_StrUserSt03 改為strDeptNo因文雄需操作S部門-秀玲
   '專利國內部
   If Left(strDeptNo, 2) = "P1" Then
        'Modify by Amy 2020/11/27 +st03 +經副理級以上人員 原:Case 1 ...
        Select Case Combo3.ListIndex
           Case GetGroup(True, "經副理級以上人員"): stCon = " and (SubStr(st15,1,2)='P1'  or SubStr(st03,1,2)='P1') And st20<='44'"
           Case GetGroup(True, "工程師"): stCon = " and (st15<>'P12' and st15<>'P13' and st15<>'P14' Or st03<>'P12' and st03<>'P13' and st03<>'P14')"
           Case GetGroup(True, "工程師-北所"): stCon = " and (st15<>'P12' and st15<>'P13' and st15<>'P14' Or st03<>'P12' and st03<>'P13' and st03<>'P14') and st06='1'"
           Case GetGroup(True, "工程師-中所"): stCon = " and (st15<>'P12' and st15<>'P13' and st15<>'P14' Or st03<>'P12' and st03<>'P13' and st03<>'P14') and st06='2'"
           Case GetGroup(True, "工程師-南所"): stCon = " and (st15<>'P12' and st15<>'P13' and st15<>'P14' Or st03<>'P12' and st03<>'P13' and st03<>'P14') and st06='3'"
           Case GetGroup(True, "工程師-高所"): stCon = " and (st15<>'P12' and st15<>'P13' and st15<>'P14' Or st03<>'P12' and st03<>'P13' and st03<>'P14') and st06='4'"
           Case GetGroup(True, "英文顧問"): stCon = " and (st15='P14' Or st03='P14')"
           Case GetGroup(True, "程序"): stCon = " and (st15='P12' Or st03='P12')"
           Case GetGroup(True, "繪圖"): stCon = " and (st15='P13' Or st03='P13')"
           Case GetGroup(True, "繪圖-北所"): stCon = " and (st15='P13' Or st03='P13') and st06='1'"
           Case GetGroup(True, "繪圖-中所"): stCon = " and (st15='P13' Or st03='P13') and st06='2'"
           Case GetGroup(True, "繪圖-南所"): stCon = " and (st15='P13' Or st03='P13') and st06='3'"
           Case GetGroup(True, "繪圖-高所"): stCon = " and (st15='P13' Or st03='P13') and st06='4'"
           Case Else
                stCon = ChgCombo3Sql
        End Select
        If Combo3.ListIndex <= 12 Then stCon = stCon & " and (st15 like 'P1%' Or st03 like 'P1%')"
        'end 2020/11/27
   Else
        If Combo3.ListIndex = 0 Then
            If Left(strDeptNo, 1) = "S" Then
                stCon = stCon & " and (st15 like 'S%' Or st03 like 'S%' ) "
            Else
                stCon = stCon & " and (st15 like '" & Left(strDeptNo, 2) & "%' Or st03 like '" & Left(strDeptNo, 2) & "%')"
            End If
        Else
            stCon = ChgCombo3Sql
        End If
   End If
   'Add by Amy 2018/09/18 有登記議題者不重覆顯示
   If m_JoinList <> MsgText(601) Then
        'stCon = stCon & " And st01 not in ('"& replace(m_joinlist,",","','") &')"
        stCon = stCon & " And InStr('" & m_JoinList & "',st01)=0"
   End If
   
   'Modify by Amy 2022/01/06 避免未改到,改至funcion
   strExc(0) = GetListSql(1, stCon)
   'end 2018/09/18
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         'Modify by Amy 2022/01/06 改Form2.0,使用PUB_Num2Id會錯
         'List1.AddItem .Fields(1), 0
         'List1.ItemData(0) = PUB_Id2Num(.Fields(0))
         List1.AddItem .Fields("C02")
         .MoveNext
      Loop
      End With
   End If
   'Add by Amy 2018/09/18 只有一筆自動反白
   If List1.ListCount = 1 Then
      List1.Selected(0) = True
   End If
End Sub

'以名稱 or 員編加入參加人員
'Modify by Amy 2022/01/06 原:As ListBox->object
Private Sub AddOneMan(oList2 As Object)
   'Modified by Morgan 2022/5/31
   'Dim jj As Integer, lngItemData As Long
   Dim jj As Integer, lngItemData As String
   'end 2022/5/31
   
   If Text1 <> "" Then
      'Moidfy by Amy 2019/11/12 原:st03 改st15因文雄需操作S部門-秀玲
      'Modify by Amy 2022/01/06 避免未改到,改至funcion
      strExc(0) = GetListSql(2)
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         If .Fields("st04") = "1" Then
            'Modify by Amy 2021/01/06 改Form2.0 無法使用ItemData
            'lngItemData = PUB_Id2Num(.Fields("st01"))
            lngItemData = .Fields("C02")
            For jj = 0 To oList2.ListCount - 1
               'If oList2.ItemData(jj) = lngItemData Then
               If oList2.List(jj) = lngItemData Then
                  Exit For
               End If
            Next
            If jj = oList2.ListCount Then
               oList2.AddItem .Fields("C02"), oList2.ListCount
               'oList2.ItemData(oList2.ListCount - 1) = lngItemData
               oList2.List(oList2.ListCount - 1) = lngItemData
               oList2.Selected(oList2.ListCount - 1) = True
               oList2.ListIndex = oList2.ListCount - 1
            End If
            'end 2022/01/06
            Text1 = ""
         Else
            MsgBox "員工已離職！", vbExclamation
            Text1.SetFocus
         End If
         End With
      Else
         MsgBox "員工編號輸入錯誤！", vbExclamation
         Text1.SetFocus
     End If
   End If
End Sub

Private Sub Command1_Click(Index As Integer)
   Select Case Index
   Case 0 '加入
      AddOneMan List2
      lbl_Name.Caption = ""
   Case 1 '取消
      Unload Me
   Case 2 '確認
      If SaveList Then
         Unload Me
      End If
   Case 3 '全選-待選清單
      SelectAll List1
   Case 4 '全選-已選人員
      SelectAll List2
   Case 5 '->
      Add2List List1, List2
   Case 6 '刪除 'Add by Amy 2021/01/06
      mnuPopItem_Click (0)
   End Select
End Sub

Private Function SaveList() As Boolean
   Dim ii As Integer, stList As String
   Dim strData As String 'Add by Amy 2022/01/06
   
   If List2.ListCount > 0 Then
      'Modify by Amy 2022/01/06 改Form2.0,使用PUB_Num2Id會錯,故改寫法
      'Add by Amy 2018/09/18 +if 有議題字樣不需回傳
      If InStr(List2.List(ii), "-議題") = 0 Then
            'stList = PUB_Num2Id(List2.ItemData(0))
            strData = List2.List(ii)
            strData = Mid(strData, Val(InStr(strData, "(")) + 1)
            strData = Mid(strData, 1, 5)
            stList = strData
      End If
      
      For ii = 1 To List2.ListCount - 1
         If InStr(List2.List(ii), "-議題") = 0 Then
            'stList = stList & "," & PUB_Num2Id(List2.ItemData(ii))
            strData = List2.List(ii)
            strData = Mid(strData, Val(InStr(strData, "(")) + 1)
            strData = Mid(strData, 1, 5)
            stList = stList & "," & strData
         End If
      Next
      'end 2022/01/06
      If Left(stList, 1) = "," Then stList = Mid(stList, 2)
      'end 2018/09/18
      fmParent.m_stNumList = stList
   Else
      MsgBox "尚未選擇人員！", vbExclamation
   End If
   SaveList = True
End Function

'Modify by Amy 2022/01/06 原:As ListBox->object
Private Sub SelectAll(oList As Object)
   Dim ii As Integer
   For ii = 0 To oList.ListCount - 1
      oList.Selected(ii) = True
   Next
End Sub

'Modify by Amy 2022/01/06 原:As ListBox->object
Private Sub Add2List(oList1 As Object, oList2 As Object)
   Dim ii As Integer, jj As Integer
   If oList1.ListCount = 0 Then Exit Sub
   For ii = oList1.ListCount - 1 To 0 Step -1
      If oList1.Selected(ii) Then
         For jj = 0 To oList2.ListCount - 1
            'Modify byAmy 2022/01/06 改Form2.0 無法使用ItemData
            'If oList2.ItemData(jj) = oList1.ItemData(ii) Then
            If oList2.List(jj) = oList1.List(ii) Then
               Exit For
            End If
         Next
         If jj = oList2.ListCount Then
            oList2.AddItem oList1.List(ii), oList2.ListCount
            'Modify byAmy 2022/01/06 改Form2.0 無法使用ItemData
            'oList2.ItemData(oList2.ListCount - 1) = oList1.ItemData(ii)
            oList2.List(oList2.ListCount - 1) = oList1.List(ii)
            oList2.Selected(oList2.ListCount - 1) = True
            oList2.ListIndex = oList2.ListCount - 1
         End If
         oList1.RemoveItem ii
      End If
   Next
   
End Sub

Private Sub InitData()
   Dim strSB01 As String
   Dim intIdx As Integer 'Add by Amy 2018/09/18
   Dim i As Integer, stGroupN As String, stSpace As String 'Add by Amy 2020/11/27
 
   Combo3.Clear
   lbl_Name.Caption = "" 'Add by Amy 2019/01/24
   'Modify by Amy 2018/09/18
   intIdx = 0
   '內專
   'Moidfy by Amy 2019/11/12 原:Pub_StrUserSt03 改為strDeptNo/st03 改為st15 因文雄需操作S部門-秀玲
   If Left(strDeptNo, 2) = "P1" Then
        'Modify by Amy 2020/11/27 改動態
'        Combo3.AddItem "專利處", 0
'        Combo3.AddItem "　工程師", 1
'        Combo3.AddItem "　　北所", 2
'        Combo3.AddItem "　　中所", 3
'        Combo3.AddItem "　　南所", 4
'        Combo3.AddItem "　　高所", 5
'        Combo3.AddItem "　英文顧問", 6
'        Combo3.AddItem "　程序", 7
'        Combo3.AddItem "　繪圖", 8
'        Combo3.AddItem "　　北所", 9
'        Combo3.AddItem "　　中所", 10
'        Combo3.AddItem "　　南所", 11
'        Combo3.AddItem "　　高所", 12
'        intIdx = 13
        'Memo 數字:表階層
        stGroup = "0專利國內部;1經副理級以上人員;1工程師;2工程師-北所;2工程師-中所;2工程師-南所;2工程師-高所;" & _
                        "1英文顧問;1程序;1繪圖;2繪圖-北所;2繪圖-中所;2繪圖-南所;2繪圖-高所"
        arrGroup = Split(stGroup, ";")
        For i = LBound(arrGroup) To UBound(arrGroup)
            stGroupN = arrGroup(i)
            stSpace = Replace(Space(Val(Mid(stGroupN, 1))), " ", "　")
            If InStr(arrGroup(i), "所") > 0 Then
                stGroupN = Mid(stGroupN, Val(InStr(stGroupN, "-")))
            End If
            Combo3.AddItem stSpace & Mid(stGroupN, 2), intIdx: intIdx = intIdx + 1
        Next i
        'end 2020/11/27
        'Add by Amy 2020/12/25 從下面搬上來
        If bolPublic = True Then
            Combo3.AddItem "專利國外部", intIdx: intIdx = intIdx + 1
            Call SetCombo3Item("F2", intIdx)
            Combo3.AddItem "智權部", intIdx: intIdx = intIdx + 1
            Call SetCombo3Item("S", intIdx)
        End If
   Else
        intIdx = 0
        'Modify by Amy 2020/11/27 組織調動部門名稱修改
        Select Case Left(strDeptNo, 2)
          '內商
          Case "P2"
            Combo3.AddItem "商標部內商", intIdx: intIdx = intIdx + 1
            Call SetCombo3Item("P2", intIdx)
            If bolPublic = True Then
                Combo3.AddItem "商標部外商", intIdx: intIdx = intIdx + 1
                Call SetCombo3Item("F1", intIdx)
                Combo3.AddItem "智權部", intIdx: intIdx = intIdx + 1
                Call SetCombo3Item("S", intIdx)
            End If
          '外商
          Case "F1"
            Combo3.AddItem "商標部外商", intIdx: intIdx = intIdx + 1
            Call SetCombo3Item("F1", intIdx)
            If bolPublic = True Then
                Combo3.AddItem "商標部內商", intIdx: intIdx = intIdx + 1
                Call SetCombo3Item("P2", intIdx)
                Combo3.AddItem "智權部", intIdx: intIdx = intIdx + 1
                Call SetCombo3Item("S", intIdx)
            End If
          '外專
          Case "F2"
            Combo3.AddItem "專利國外部", intIdx: intIdx = intIdx + 1
            Call SetCombo3Item("F2", intIdx)
            If bolPublic = True Then
                Combo3.AddItem "專利國內部", intIdx: intIdx = intIdx + 1
                Call SetCombo3Item("P1", intIdx)
                Combo3.AddItem "智權部", intIdx: intIdx = intIdx + 1
                Call SetCombo3Item("S", intIdx)
            End If
          'Add by Amy 2019/01/24 +業務拓展
          Case "F4"
            Combo3.AddItem "業務拓展", intIdx: intIdx = intIdx + 1
            Call SetCombo3Item("F4", intIdx)
            If bolPublic = True Then
                Combo3.AddItem "專利國內部", intIdx: intIdx = intIdx + 1
                Call SetCombo3Item("P1", intIdx)
                Combo3.AddItem "專利國外部", intIdx: intIdx = intIdx + 1
                Call SetCombo3Item("F2", intIdx)
            End If
          '法務
          Case "P3", "L0"
            Combo3.AddItem "法律所", intIdx: intIdx = intIdx + 1
            Call SetCombo3Item("L", intIdx)
            If bolPublic = True Then
                Combo3.AddItem "智權部", intIdx: intIdx = intIdx + 1
                Call SetCombo3Item("S", intIdx)
                Combo3.AddItem "專利國內部", intIdx: intIdx = intIdx + 1
                Call SetCombo3Item("P1", intIdx)
                Combo3.AddItem "商標部內商", intIdx: intIdx = intIdx + 1
                Call SetCombo3Item("P2", intIdx)
                Combo3.AddItem "商標部外商", intIdx: intIdx = intIdx + 1
                Call SetCombo3Item("F1", intIdx)
                Combo3.AddItem "專利國外部", intIdx: intIdx = intIdx + 1
                Call SetCombo3Item("F2", intIdx)
            End If
          'Add by Amy 2019/01/24 創新業務客服組/顧問組
          '創新業務客服組
          Case "W1"
            Combo3.AddItem "創新業務客服組", intIdx: intIdx = intIdx + 1
            Call SetCombo3Item("W1", intIdx)
          '創新業務顧問組
          Case "W2"
            Combo3.AddItem "創新業務顧問組", intIdx: intIdx = intIdx + 1
            Call SetCombo3Item("W2", intIdx)
            If bolPublic = True Then
                Combo3.AddItem "專利國內部", intIdx: intIdx = intIdx + 1
                Call SetCombo3Item("P1", intIdx)
                Combo3.AddItem "業務拓展", intIdx: intIdx = intIdx + 1
                Call SetCombo3Item("F4", intIdx)
            End If
          'end 2019/01/24
          '研發
          Case "D0"
            Call SetCombo3Item("Other", intIdx)
          Case Else
            '智權部
            If Left(strDeptNo, 1) = "S" Then
                Combo3.AddItem "智權部", intIdx: intIdx = intIdx + 1
                Call SetCombo3Item("S", intIdx)
                If bolPublic = True Then
                    Combo3.AddItem "專利國內部", intIdx: intIdx = intIdx + 1
                    Call SetCombo3Item("P1", intIdx)
                    Combo3.AddItem "商標部內商", intIdx: intIdx = intIdx + 1
                    Call SetCombo3Item("P2", intIdx)
                    Combo3.AddItem "商標部外商", intIdx: intIdx = intIdx + 1
                    Call SetCombo3Item("F1", intIdx)
                    Combo3.AddItem "專利國外部", intIdx: intIdx = intIdx + 1
                    Call SetCombo3Item("F2", intIdx)
                    Combo3.AddItem "法律所", intIdx: intIdx = intIdx + 1
                    Call SetCombo3Item("L", intIdx)
                End If
            End If
      End Select
      'end 2020/11/27
   End If
   'end 2018/09/18
   
   List1.Clear '待選清單
   
   If m_stNumList <> "" Then
      'Modify by Amy 2022/01/06 避免未改到,改至funcion
      strExc(0) = GetListSql(0)
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         Do While Not .EOF
            List2.AddItem .Fields(0), 0
            'Mark by Amy 2022/01/06 改Form2.0無法使用ItemData
            'List2.ItemData(0) = PUB_Id2Num(.Fields(0))
            .MoveNext
         Loop
         End With
      End If
   End If
   'end 2019/11/12
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me, True
   InitData
   'Modify by Amy 2022/01/06一開始將ListBox拉到需要的大小,字型會自動放大；所以畫面預設為一列高度,Form_Load才放大到需要的大小
   List1.Height = 3300
   List1.Width = 2650
   List2.Height = 3300
   List2.Width = 2570
End Sub

Private Sub Form_Unload(Cancel As Integer)
   bolPublic = False 'Add by Amy 2018/09/18
   strDeptNo = "" 'Add by Amy 2019/11/12
   Set frm140113_2 = Nothing
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then
      PopupMenu Me.mnuPop
   End If
End Sub

Private Sub mnuPopItem_Click(Index As Integer)
   Dim ii As Integer
   If Index = 0 Then
      If List2.ListCount = 0 Then Exit Sub
      For ii = List2.ListCount - 1 To 0 Step -1
         'Modify by Amy 2018/09/18 議題字樣不可刪
         If List2.Selected(ii) And InStr(List2.List(ii), "-議題") = 0 Then
            List2.RemoveItem ii
         End If
      Next
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add by Amy 2018/09/18
Private Sub SetCombo3Item(ByVal stDeptNo As String, ByRef intIdx As Integer)
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    strQ = "And Substr(a0901,1,2)='" & stDeptNo & "' "
    '智權
    If stDeptNo = "S" Then
        'Modify by Amy 2020/12/22 S00改S01
        strQ = " And SubStr(a0901,1,1)='S' And a0901 not in ('S01','S12','S10','S20','S29','S91') "
    '法務
    ElseIf stDeptNo = "L" Then
        'Modify by Amy 2020/12/22 拿掉P31
        strQ = " And a0901 in ('L01','L02') "
    '研發
    ElseIf stDeptNo = "Other" Then
        'Modify by Amy 2020/12/22 加P31,S00改S01
        strQ = " And a0901<>'" & strDeptNo & "' And a0901 not in ('F31','F51','F52','P31','P41','S01','S10','S20','S29','S91','M41')" & _
                  " And a0901>'D' And SubStr(a0901,1,1)<>'R' "
    End If
    strQ = "Select A0901,A0902 From ACC090 Where a0904 <> 'Y' and a0901<>'P29' And a0901<>'P19' " & strQ & _
              " Order by a0901"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        Do While Not RsQ.EOF
            Combo3.AddItem "　" & RsQ.Fields("A0902"), intIdx
            intIdx = intIdx + 1
            If "" & RsQ.Fields("A0901") = "F21" Then
                Combo3.AddItem "　　電機組", intIdx
                intIdx = intIdx + 1
                Combo3.AddItem "　　化學組", intIdx
                intIdx = intIdx + 1
                Combo3.AddItem "　　日文組", intIdx
                intIdx = intIdx + 1
                Combo3.AddItem "　　機械組", intIdx
                intIdx = intIdx + 1
            End If
            RsQ.MoveNext
        Loop
    End If
    RsQ.Close
End Sub

'來源下拉選單對映之部門語法
Private Function ChgCombo3Sql() As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    'Moidfy by Amy 2019/11/12 原:st03 改st15因文雄需操作S部門-秀玲
    'Modify by Amy 2020/12/22 +st03
    If Left(Combo3.Text, 1) <> "　" Or Left(Combo3.Text, 2) = "　　" Then
        Select Case Trim(Combo3)
            Case "智權部"
                ChgCombo3Sql = "And (SubStr(st15,1,1)='S' Or SubStr(st03,1,1)='S') "
            'Modify by Amy 2021/01/22 bug-原:Or SubStr(st03,1,1)
            Case "專利國內部" 'Modify by Amy 2020/11/27 原:專利處
                ChgCombo3Sql = "And (SubStr(st15,1,2)='P1' Or SubStr(st03,1,2)='P1') "
            Case "商標部內商" 'Modify by Amy 2020/11/27 原:商標處
                ChgCombo3Sql = "And (SubStr(st15,1,2)='P2' Or SubStr(st03,1,2)='P2') "
            Case "商標部外商" 'Modify by Amy 2020/11/27 原:外商人員
                ChgCombo3Sql = "And (SubStr(st15,1,2)='F1' Or SubStr(st03,1,2)='F1') "
            Case "專利國外部" 'Modify by Amy 2020/11/27 原:外專工程師
                ChgCombo3Sql = "And (SubStr(st15,1,2)='F2' Or SubStr(st03,1,2)='F2') "
            'end 2021/01/22
            Case "法律所" 'Modify by Amy 2020/11/27 原:法務人員,拿掉P3
                ChgCombo3Sql = "And (SubStr(st15,1,2) in ('L0') Or SubStr(st03,1,2) in ('L0')) "
            Case "管理部"
                ChgCombo3Sql = "And (SubStr(st15,1,1)='M' And st15<>'M41' Or SubStr(st03,1,1)='M' And st03<>'M41') "
            Case "專利國外部工程師" 'Modify by Amy 2020/11/27 原:外專工程師
                ChgCombo3Sql = "And (st15='F21' Or SubStr(st03,1,1)='F21') "
            Case "電機組"
                ChgCombo3Sql = "And (st15='F21' Or SubStr(st03,1,1)='F21') And st16='1' "
            Case "化學組"
                ChgCombo3Sql = "And (st15='F21' Or SubStr(st03,1,1)='F21') And st16='2' "
            Case "日文組"
                ChgCombo3Sql = "And (st15='F21' Or SubStr(st03,1,1)='F21') And st16='3' "
            Case "機械組"
                ChgCombo3Sql = "And (st15='F21' Or SubStr(st03,1,1)='F21') And st16='4' "
        End Select
        Exit Function
    End If
    'end 2019/11/12
    If Trim(Combo3.Text) = "法律所" Then strQ = " And a0901<>'L' "
    
    strQ = "Select a0901 From Acc090 Where a0902='" & Trim(Combo3.Text) & "' " & strQ
    
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        'Modify by Amy 2020/11/27 +st03
        ChgCombo3Sql = " And (st15='" & RsQ.Fields("a0901") & "' Or st03='" & RsQ.Fields("a0901") & "' ) "
    End If
    RsQ.Close
End Function

'Modify by Amy 2022/01/06 原:As ListBox->object
Private Sub SetListScroll(oList As Object)
   Dim ii As Integer
   Dim lWnow As Long, lWmax As Long
   
   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next
  
   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

'Add by Amy 2019/01/24 可輸員編或姓名
Private Sub Text1_Validate(Cancel As Boolean)
    Dim stST01 As String, stST02 As String
    
    Text1 = Trim(Text1)
    If Text1 = MsgText(601) Then Exit Sub
    
    If ByInputGetST01or02(Text1, stST01, stST02) = False Then
        Cancel = True
        Text1.SetFocus
        Exit Sub
    End If
    Text1 = stST01
    lbl_Name = stST02
End Sub

'Add by Amy 2020/11/27 取得Mail Group
'IsNo:True-回傳對應 Index/False-回傳對應字串
Private Function GetGroup(ByVal iSNo As Boolean, ByVal stFindStr As String) As String
    Dim ii As Integer
    
    For ii = LBound(arrGroup) To UBound(arrGroup)
        If iSNo = True Then
            If UCase(Mid(arrGroup(ii), 2)) = UCase(stFindStr) Then
                GetGroup = ii
                Exit For
            End If
        Else
            GetGroup = arrGroup(Val(stFindStr))
        End If
    Next ii
   
End Function

Public Sub SetMailInfo()
    List1.Visible = False
    Frame1.Visible = False
    Command1(1).Top = 10 '取消鈕
    Command1(1).Left = 7800
    lblMemo.Top = 10
    lblMemo.Left = 0
    lblMemo.Width = Me.Width
    lblMemo.Height = 5700
    Command1(2).Visible = False
    Label2.Visible = False
    List2.Visible = False
    Command1(5).Visible = False
End Sub

'Add by Amy 2022/01/06 避免有未改到,故改成Funion
'intChoose:0-初始設定(InitData用)/1-設定已選人員(SetBookInList用)/2-以名稱 or 員編加入參加人員(AddOneMan用)
Private Function GetListSql(intChoose As Integer, Optional ByVal stCon As String = "") As String
    Dim stField As String, stWhere As String, stOrder As String
    
    GetListSql = ""
    'Modify by Amy 2020/11/27 +st03
    'strExc(0) = "select st01,st02||'('||decode(st15,'P11','工程師','P13','繪圖','P12','程序','P14','英文顧問',decode(st03,'P11','工程師','P13','繪圖','P12','程序','P14','英文顧問',ST01))||'.'||ac03||')' C02" & _
                       " from staff,allcode" & _
                       " where instr('" & m_stNumList & "',st01)>0 and ac02(+)=st20 and ac01(+)='01' order by st02 desc"
    'Modify by Amy 2024/04/19 職稱第一個字為[代] 拿掉 ex:李柏翰 代經理
    'stField = "st02||'('||Decode(st15,'P11',st01||' '||'工程師','P13',st01||' '||'繪圖','P12',st01||' '||'程序','P14',st01||' '||'英文顧問',Decode(st03,'P11',st01||' '||'工程師','P13',st01||' '||'繪圖','P12',st01||' '||'程序','P14',st01||' '||'英文顧問',ST01||' '))||Decode(ac03,null,ac03,'.'||ac03)||')' C02,st01"
    stField = "st02||'('||Decode(st15,'P11',st01||' '||'工程師','P13',st01||' '||'繪圖','P12',st01||' '||'程序','P14',st01||' '||'英文顧問',Decode(st03,'P11',st01||' '||'工程師','P13',st01||' '||'繪圖','P12',st01||' '||'程序','P14',st01||' '||'英文顧問',ST01||' '))||Decode(ac03,null,ac03,'.'||Decode(SubStr(AC03,1,1),'代',SubStr(AC03,2,length(AC03)),AC03))||')' C02,st01"
    Select Case intChoose
        Case 0 '初始設定(InitData用)
            stWhere = "And InStr('" & m_stNumList & "',st01)>0 And ac02(+)=st20 And ac01(+)='01' "
            stOrder = "Order by st02 Desc"
        Case 1 '設定已選人員(SetBookInList用)
                stWhere = "And st04='1' And ac02(+)=st20 And ac01(+)='01' And st04='1' And st01>'6' And st01<'F' And substr(st01,1,3)<>'999' And st06<>'5' And SubStr(st01,4,1)<>'9' "
                stOrder = "Order by st01 Asc"
        Case 2 '以名稱 or 員編加入參加人員(AddOneMan用)
            stField = stField & ",st04"
            stWhere = "And st01='" & Text1 & "' And ac02(+)=st20 And ac01(+)='01' "
    End Select
    
    GetListSql = "Select " & stField & " " & _
                        "From Staff,AllCode Where 1=1 " & stWhere & stCon & " " & _
                        stOrder
    
End Function
