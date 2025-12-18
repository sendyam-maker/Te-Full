VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140101 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶變更名稱作業"
   ClientHeight    =   5808
   ClientLeft      =   660
   ClientTop       =   636
   ClientWidth     =   8100
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5808
   ScaleWidth      =   8100
   Begin VB.CommandButton CmdAddr 
      Caption         =   "變更地址(&A)"
      Height          =   400
      Left            =   4200
      TabIndex        =   26
      Top             =   70
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消(&U)"
      Height          =   400
      Index           =   3
      Left            =   6264
      TabIndex        =   15
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   7092
      TabIndex        =   16
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   5436
      TabIndex        =   14
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   350
      Index           =   2
      Left            =   2472
      TabIndex        =   1
      Top             =   600
      Width           =   800
   End
   Begin MSForms.TextBox Text1 
      Height          =   525
      Index           =   12
      Left            =   1320
      TabIndex        =   13
      Top             =   5190
      Width           =   6615
      VariousPropertyBits=   -1467989989
      MaxLength       =   80
      ScrollBars      =   2
      Size            =   "11668;926"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   11
      Left            =   1320
      TabIndex        =   12
      Top             =   4830
      Width           =   3855
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "6800;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   10
      Left            =   1320
      TabIndex        =   11
      Top             =   4540
      Width           =   3855
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "6800;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   9
      Left            =   1320
      TabIndex        =   10
      Top             =   4250
      Width           =   3855
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "6800;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   8
      Left            =   1320
      TabIndex        =   9
      Top             =   3960
      Width           =   3855
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "6800;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   525
      Index           =   7
      Left            =   1320
      TabIndex        =   8
      Top             =   3390
      Width           =   6615
      VariousPropertyBits=   -1467989989
      MaxLength       =   80
      ScrollBars      =   2
      Size            =   "11668;926"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   5
      Left            =   1320
      TabIndex        =   6
      Top             =   2460
      Width           =   3855
      VariousPropertyBits=   671105055
      BackColor       =   -2147483638
      MaxLength       =   30
      Size            =   "6800;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   4
      Left            =   1320
      TabIndex        =   5
      Top             =   2160
      Width           =   3855
      VariousPropertyBits=   671105055
      BackColor       =   -2147483638
      MaxLength       =   30
      Size            =   "6800;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   3
      Left            =   1320
      TabIndex        =   4
      Top             =   1860
      Width           =   3855
      VariousPropertyBits=   671105055
      BackColor       =   -2147483638
      MaxLength       =   30
      Size            =   "6800;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   525
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   990
      Width           =   6615
      VariousPropertyBits=   -1467989985
      BackColor       =   -2147483638
      MaxLength       =   80
      ScrollBars      =   2
      Size            =   "11668;926"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   3855
      VariousPropertyBits=   671105055
      BackColor       =   -2147483638
      MaxLength       =   30
      Size            =   "6800;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   630
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   525
      Index           =   6
      Left            =   1320
      TabIndex        =   7
      Top             =   2820
      Width           =   6615
      VariousPropertyBits=   -1467989985
      BackColor       =   -2147483638
      MaxLength       =   80
      ScrollBars      =   2
      Size            =   "11668;917"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "變更後客戶編號："
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   4680
      TabIndex        =   25
      Top             =   630
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   6240
      TabIndex        =   24
      Top             =   630
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "新中文名稱："
      Height          =   180
      Index           =   6
      Left            =   150
      TabIndex        =   23
      Top             =   3420
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "新英文名稱："
      Height          =   180
      Index           =   5
      Left            =   150
      TabIndex        =   22
      Top             =   4020
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "新日文名稱："
      Height          =   180
      Index           =   3
      Left            =   150
      TabIndex        =   21
      Top             =   5220
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶編號："
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   20
      Top             =   690
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "原中文名稱："
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   19
      Top             =   1050
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "原英文名稱："
      Height          =   180
      Index           =   2
      Left            =   150
      TabIndex        =   18
      Top             =   1620
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "原日文名稱："
      Height          =   180
      Index           =   4
      Left            =   150
      TabIndex        =   17
      Top             =   2850
      Width           =   1080
   End
End
Attribute VB_Name = "frm140101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/10/24 接洽單申請人資料是獨立的資料檔,不更名
'Memo by Lydia 2021/09/23 改成Form2.0 ; Text1(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
Option Explicit

' 90.07.16 modify by Ken (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
'93.11.15 ADD BY SONIA
Dim m_CU32 As String
Dim m_CU79 As String 'Add By Sindy 2011/1/18
Dim m_bolDesc As Boolean 'Added by Lydia 2025/10/30

Private Sub Form_Load()
   ' 90.07.16 modify by Ken (取得使用者執行各項功能的權限)
   m_bUpdate = IsUserHasRightOfFunction("frm140101", strEdit, False)
   ' Ken 90.07.16 -- End
   
   MoveFormToCenter Me
   CmdLock 1
   Label2(0) = "": Label2(1) = ""
   
   ' Ken 90.07.16 -- start
   'Remove by Lydia 2018/10/24 改在CmdLock控制
   'If m_bUpdate Then
   '     'Modify By Cheng 2002/10/25
   '     '畫面一載入時確定按鈕預設不能作用
'   '    Command1(0).Enabled = True
   'Else
   '    Command1(0).Enabled = False
   'End If
   'end 2018/10/24
   ' Ken 90.07.16 -- End
   
   'Added by Lydia 2018/10/24 檢查權限
   If IsUserHasRightOfFunction("frm140401", strEdit, False) = True Then
        CmdAddr.Visible = True
   Else
        CmdAddr.Visible = False
   End If
   CmdAddr.Enabled = False '確定變更名稱後,才可執行
   'end 2018/10/24
   
End Sub

Private Sub Command1_Click(Index As Integer)
Dim strTxt(1 To 25) As String
Dim i As Integer, St(1 To 6) As String, j As Integer, strTmp As String
'add by nickc 2006/06/15
Dim CuNation As String
Dim orsTmp As New ADODB.Recordset
Dim strOldName As String
Dim strCU13 As String 'Add By Sindy 2016/4/25
Dim strCRL01 As String 'Add By Sindy 2023/1/12
   
On Error GoTo ErrHand
   Select Case Index
      Case 0 '確定
         'Remove by Lydia 2021/01/07 名稱長度直接以TextBox.MaxLength控制; ex.R15419的英文名稱"Patentanwaelte · Rechtsanwaelt"字數30字，中英文長度31
'         If Not CheckLengthIsOK(Text1(7).Text, 80) Then
'            Text1(7).SetFocus
'            Exit Sub
'         End If
'         For i = 8 To 11
'            If Not CheckLengthIsOK(Text1(i).Text, 30) Then
'               Text1(i).SetFocus
'               Exit Sub
'            End If
'         Next
'         If Not CheckLengthIsOK(Text1(12).Text, 80) Then
'            Text1(12).SetFocus
'            Exit Sub
'         End If
         'end 2021/01/07
         
         For i = 1 To 6
            St(i) = Text1(i + 6).Text
         Next i
         If St(1) = "" And St(2) = "" And St(3) = "" _
            And St(4) = "" And St(5) = "" And St(6) = "" Then
               MsgBox "至少要有一種名稱，輸入錯誤 !", vbCritical
               Exit Sub
         End If
         If St(1) = Text1(1).Text And St(2) = Text1(2).Text And _
            St(3) = Text1(3).Text And St(4) = Text1(4).Text And _
            St(5) = Text1(5).Text And St(6) = Text1(6).Text Then
               MsgBox "新名稱和舊名稱完全相同，輸入錯誤 !", vbCritical
               Exit Sub
         End If
         'add by nickc 2006/06/15
         CuNation = ""
         
         'Add By Cheng 2002/05/23
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         Screen.MousePointer = vbHourglass
         
         'Added by Lydia 2025/10/30 改用模組判斷
         m_bolDesc = PUB_FilterSeekSQL("", Me)
         
        'Add By Cheng 2002/11/06
        On Error GoTo ErrorHandler
        cnnConnection.BeginTrans
         strTmp = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
         'add by nickc 2006/06/15
         strSql = "select * from customer WHERE CU01='" & strTmp & "' AND CU02='0' "
         Set orsTmp = New ADODB.Recordset
         orsTmp.CursorLocation = adUseClient
         orsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If orsTmp.RecordCount <> 0 Then
            CuNation = CheckStr(orsTmp.Fields("cu10"))
            strCU13 = orsTmp.Fields("cu13") 'Add By Sindy 2016/4/25
'edit by nickc 2008/05/08 改成 function
'            If Mid(CuNation, 1, 3) = "101" Then
'                '2008/2/29 MODIFY BY SONIA 原為A~I為101,J~Z為1011,2008年改為分四段
'                If Mid(UCase(Text1(8)), 1, 1) >= "A" And Mid(UCase(Text1(8)), 1, 1) <= "E" Then
'                     CuNation = "101"
'                ElseIf Mid(UCase(Text1(8)), 1, 1) >= "F" And Mid(UCase(Text1(8)), 1, 1) <= "I" Then
'                     CuNation = "1011"
'                ElseIf Mid(UCase(Text1(8)), 1, 1) >= "J" And Mid(UCase(Text1(8)), 1, 1) <= "N" Then
'                     CuNation = "1012"
'                ElseIf Mid(UCase(Text1(8)), 1, 1) >= "O" And Mid(UCase(Text1(8)), 1, 1) <= "Z" Then
'                     CuNation = "1013"
'                Else
'                     CuNation = "1013"
'                End If
'            ElseIf Mid(Text1(12), 1, 3) = "011" Then
'                '2008/4/21 MODIFY BY SONIA 原為A~L為011,M~Z為0111,2008/4/22改為分三段(將M~Z再細分成二段)
'                If Mid(UCase(Text1(8)), 1, 1) >= "A" And Mid(UCase(Text1(8)), 1, 1) <= "L" Then
'                     CuNation = "011"
'                ElseIf Mid(UCase(Text1(8)), 1, 1) >= "M" And Mid(UCase(Text1(8)), 1, 1) <= "O" Then
'                     CuNation = "0111"
'                ElseIf Mid(UCase(Text1(8)), 1, 1) >= "P" And Mid(UCase(Text1(8)), 1, 1) <= "Z" Then
'                     CuNation = "0112"
'                ElseIf Trim(Text1(8)) = "" Then
'                     CuNation = "0112"
'                End If
'            End If
            CuNation = pub_NationByName(Text1(8) & Text1(9) & Text1(10) & Text1(11), CuNation)
         End If
         
         '92.9.24 modify by sonia
         'strTxt(1) = "UPDATE CUSTOMER SET CU32='N',CU02='" & Right(Label2(0), 1) & "' WHERE CU01='" & strTmp & "' AND CU02='0'"
         'Modify By Sindy 2011/1/18 調整CU79
         strTxt(1) = "UPDATE CUSTOMER SET CU32='N',CU02='" & Right(Label2(0), 1) & "',CU79='" & ChangeTStringToTDateString(strSrvDate(2)) & "更名'||decode(CU79,'',';',decode(substr(CU79,1,1),';',CU79,';'||CU79)) WHERE CU01='" & strTmp & "' AND CU02='0'"
         '92.9.24 end
         'ADD BY NICKC 2007/01/03
         'Modified by Lydia 2018/10/24 +詳細記錄(True)
         'Modified by Lydai 2025/10/30 +strTmp 將改前客戶編號寫入log; 詳細記錄(True)改用模組判斷 True=>m_bolDesc
         Pub_SeekTbLog strTxt(1), , m_bolDesc, , , strTmp
        'Add By Cheng 2002/11/06
        cnnConnection.Execute strTxt(1)
        '92.6.9 取消CU32
        '93.11.15 MODIFY BY SONIA 保留原記錄之CU32
         'edit by nickc 2005/12/14 cu111,cu110 帶過去
         'edit by nickc 2006/06/15 加入存入國籍
         'Modify by Morgan 2006/10/23 改用欄位數 TF_CU 跑回圈方式做，這樣新增欄位時才不用改
         strTxt(2) = "INSERT INTO CUSTOMER (CU01"
         strSql = "SELECT CU01"
         For intI = 2 To TF_CU
            '除Create(Update) ID, Date, Time 以外都要
            If intI < 81 Or intI > 86 Then
               strTxt(2) = strTxt(2) & ",CU" & Format(intI, "0#")
               Select Case intI
                  Case 2 '變更碼
                     strSql = strSql & ",'0'"
                  Case 4 '中文名稱
                     strSql = strSql & "," & CNULL(ChgSQL(St(1)))
                  Case 5 '英文名稱
                     strSql = strSql & "," & CNULL(ChgSQL(St(2)))
                  Case 88 '英文名稱
                     strSql = strSql & "," & CNULL(ChgSQL(St(3)))
                  Case 89 '英文名稱
                     strSql = strSql & "," & CNULL(ChgSQL(St(4)))
                  Case 90 '英文名稱
                     strSql = strSql & "," & CNULL(ChgSQL(St(5)))
                  Case 6 '日文名稱
                     strSql = strSql & "," & CNULL(ChgSQL(St(6)))
                  Case 10 '國籍
                     strSql = strSql & "," & CNULL(CuNation)
                  Case 32 '是否寄雜誌
                     strSql = strSql & "," & CNULL(m_CU32)
                  Case 79 '客戶備註 Add By Sindy 98/02/16
                     '舊名稱
                     strOldName = ""
                     For i = 1 To 6
                        If Trim(Text1(i).Text) <> "" And Not IsNull(Text1(i).Text) And _
                           Trim(Text1(i).Text) <> Trim(St(i)) Then
                           If i = 2 Or i = 3 Or i = 4 Or i = 5 Then
                              If strOldName = "" Then
                                 strOldName = Trim(Text1(2).Text) & " " & Trim(Text1(3).Text) & " " & Trim(Text1(4).Text) & " " & Trim(Text1(5).Text)
                              Else
                                 strOldName = strOldName & " " & Trim(Text1(2).Text) & " " & Trim(Text1(3).Text) & " " & Trim(Text1(4).Text) & " " & Trim(Text1(5).Text)
                              End If
                              i = 5
                           Else
                              If strOldName = "" Then
                                 strOldName = Trim(Text1(i).Text)
                              Else
                                 strOldName = strOldName & " " & Trim(Text1(i).Text)
                              End If
                           End If
                        End If
                     Next i
                     'Modify By Sindy 2011/1/18 調整CU79
                     If m_CU79 = "" Or Left(Trim(m_CU79), 1) <> ";" Then
                        m_CU79 = ";" & m_CU79
                     End If
                     'Modified by Morgan 2016/2/2 名稱會有單引號
                     strSql = strSql & ",'" & ChgSQL(ChangeTStringToTDateString(strSrvDate(2)) & "更名(" & "舊名稱：" & ChgSQL(Trim(strOldName)) & ")" & m_CU79) & "'"
                  Case Else
                     strSql = strSql & ",CU" & Format(intI, "0#")
               End Select
            End If
         Next
         strTxt(2) = strTxt(2) & ") " & strSql & " FROM CUSTOMER WHERE CU01='" & strTmp & "' AND CU02='" & Right(Label2(0), 1) & "'"
         'end 2006/10/23
         
         'ADD BY NICKC 2007/01/03
         'Modified by Lydia 2018/10/24 +詳細記錄(True)
         'Modify by Amy 2025/09/19 +strTmp 將改前客戶編號寫入log
         'Modified by Lydai 2025/10/30 詳細記錄(True)改用模組判斷 True=>m_bolDesc
         Pub_SeekTbLog strTxt(2), , m_bolDesc, , , strTmp
         
        'Add By Cheng 2002/11/06
        cnnConnection.Execute strTxt(2)
          
         strTxt(3) = "UPDATE PATENT SET PA26 = DECODE(PA26,'" & strTmp & "0','" & Label2(0) & "',PA26), " & _
            "PA27 = DECODE(PA27,'" & strTmp & "0','" & Label2(0) & "',PA27), " & _
            "PA28 = DECODE(PA28,'" & strTmp & "0','" & Label2(0) & "',PA28), " & _
            "PA29 = DECODE(PA29,'" & strTmp & "0','" & Label2(0) & "',PA29), " & _
            "PA30 = DECODE(PA30,'" & strTmp & "0','" & Label2(0) & "',PA30) WHERE PA26='" & strTmp & "0' OR PA27='" & strTmp & "0' OR PA28='" & strTmp & "0' OR PA29='" & strTmp & "0' OR PA30='" & strTmp & "0'"
         
         'ADD BY NICKC 2007/01/03
         'Modified by Lydia 2018/10/24 +詳細記錄(True)
         'Modified by Lydai 2025/10/30 +strTmp 將改前客戶編號寫入log; 詳細記錄(True)改用模組判斷 True=>m_bolDesc
         Pub_SeekTbLog strTxt(3), , m_bolDesc, , , strTmp
         
         
        'Add By Cheng 2002/11/06
        cnnConnection.Execute strTxt(3)
         
         'Modify By Sindy 2011/2/24 增加TM78,TM79,TM80,TM81
         'strTxt(4) = "UPDATE TRADEMARK SET TM23='" & Label2(0) & "' WHERE TM23='" & strTmp & "0'"
         strTxt(4) = "UPDATE TRADEMARK SET TM23 = DECODE(TM23,'" & strTmp & "0','" & Label2(0) & "',TM23), " & _
            "TM78 = DECODE(TM78,'" & strTmp & "0','" & Label2(0) & "',TM78), " & _
            "TM79 = DECODE(TM79,'" & strTmp & "0','" & Label2(0) & "',TM79), " & _
            "TM80 = DECODE(TM80,'" & strTmp & "0','" & Label2(0) & "',TM80), " & _
            "TM81 = DECODE(TM81,'" & strTmp & "0','" & Label2(0) & "',TM81) WHERE TM23='" & strTmp & "0' OR TM78='" & strTmp & "0' OR TM79='" & strTmp & "0' OR TM80='" & strTmp & "0' OR TM81='" & strTmp & "0'"
         
         'ADD BY NICKC 2007/01/03
         'Modified by Lydia 2018/10/24 +詳細記錄(True)
         'Modified by Lydai 2025/10/30 +strTmp 將改前客戶編號寫入log; 詳細記錄(True)改用模組判斷 True=>m_bolDesc
         Pub_SeekTbLog strTxt(4), , m_bolDesc, , , strTmp
         
        'Add By Cheng 2002/11/06
        cnnConnection.Execute strTxt(4)
         
         'Modify By Sindy 2011/2/24 增加LC43,LC44,LC45,LC46
         'strTxt(5) = "UPDATE LAWCASE SET LC11='" & Label2(0) & "' WHERE LC11='" & strTmp & "0'"
         strTxt(5) = "UPDATE LAWCASE SET LC11 = DECODE(LC11,'" & strTmp & "0','" & Label2(0) & "',LC11), " & _
            "LC43 = DECODE(LC43,'" & strTmp & "0','" & Label2(0) & "',LC43), " & _
            "LC44 = DECODE(LC44,'" & strTmp & "0','" & Label2(0) & "',LC44), " & _
            "LC45 = DECODE(LC45,'" & strTmp & "0','" & Label2(0) & "',LC45), " & _
            "LC46 = DECODE(LC46,'" & strTmp & "0','" & Label2(0) & "',LC46) WHERE LC11='" & strTmp & "0' OR LC43='" & strTmp & "0' OR LC44='" & strTmp & "0' OR LC45='" & strTmp & "0' OR LC46='" & strTmp & "0'"
         
         'ADD BY NICKC 2007/01/03
         'Modified by Lydia 2018/10/24 +詳細記錄(True)
         'Modified by Lydai 2025/10/30 +strTmp 將改前客戶編號寫入log; 詳細記錄(True)改用模組判斷 True=>m_bolDesc
         Pub_SeekTbLog strTxt(5), , m_bolDesc, , , strTmp
         
        'Add By Cheng 2002/11/06
        cnnConnection.Execute strTxt(5)
         
        '92.9.15 CANCEL BY SONIA
         'strTxt(6) = "UPDATE HIRECASE SET HC05='" & Label2(0) & "' WHERE HC05='" & strTmp & "0'"
        'Add By Cheng 2002/11/06
        'cnnConnection.Execute strTxt(6)
        '92.9.15 END
         
         'Modify By Sindy 2011/2/24 增加SP65,SP66
'         strTxt(7) = "UPDATE SERVICEPRACTICE SET SP08 = DECODE(SP08,'" & strTmp & "0','" & Label2(0) & "',SP08), " & _
'            "SP58 = DECODE(SP58,'" & strTmp & "0','" & Label2(0) & "',SP58), " & _
'            "SP59 = DECODE(SP59,'" & strTmp & "0','" & Label2(0) & "',SP59) WHERE SP08='" & strTmp & "0' OR SP58='" & strTmp & "0' OR SP59='" & strTmp & "0'"
         strTxt(7) = "UPDATE SERVICEPRACTICE SET SP08 = DECODE(SP08,'" & strTmp & "0','" & Label2(0) & "',SP08), " & _
            "SP58 = DECODE(SP58,'" & strTmp & "0','" & Label2(0) & "',SP58), " & _
            "SP59 = DECODE(SP59,'" & strTmp & "0','" & Label2(0) & "',SP59), " & _
            "SP65 = DECODE(SP65,'" & strTmp & "0','" & Label2(0) & "',SP65), " & _
            "SP66 = DECODE(SP66,'" & strTmp & "0','" & Label2(0) & "',SP66) WHERE SP08='" & strTmp & "0' OR SP58='" & strTmp & "0' OR SP59='" & strTmp & "0' OR SP65='" & strTmp & "0' OR SP66='" & strTmp & "0'"
            
         'ADD BY NICKC 2007/01/03
         'Modified by Lydia 2018/10/24 +詳細記錄(True)
         'Modified by Lydai 2025/10/30 +strTmp 將改前客戶編號寫入log; 詳細記錄(True)改用模組判斷 True=>m_bolDesc
         Pub_SeekTbLog strTxt(7), , m_bolDesc, , , strTmp
         
         'Add By Cheng 2002/11/06
         cnnConnection.Execute strTxt(7)
         
         '2010/5/13 MODIFY BY SONIA 加CP89~CP96
         strTxt(8) = "UPDATE CASEPROGRESS SET CP55 = DECODE(CP55,'" & strTmp & "0','" & Label2(0) & "',CP55), " & _
            "CP56 = DECODE(CP56,'" & strTmp & "0','" & Label2(0) & "',CP56), " & _
            "CP72 = DECODE(CP72,'" & strTmp & "0','" & Label2(0) & "',CP72), " & _
            "CP89 = DECODE(CP89,'" & strTmp & "0','" & Label2(0) & "',CP89), " & _
            "CP90 = DECODE(CP90,'" & strTmp & "0','" & Label2(0) & "',CP90), " & _
            "CP91 = DECODE(CP91,'" & strTmp & "0','" & Label2(0) & "',CP91), " & _
            "CP92 = DECODE(CP92,'" & strTmp & "0','" & Label2(0) & "',CP92), " & _
            "CP93 = DECODE(CP93,'" & strTmp & "0','" & Label2(0) & "',CP93), " & _
            "CP94 = DECODE(CP94,'" & strTmp & "0','" & Label2(0) & "',CP94), " & _
            "CP95 = DECODE(CP95,'" & strTmp & "0','" & Label2(0) & "',CP95), " & _
            "CP96 = DECODE(CP96,'" & strTmp & "0','" & Label2(0) & "',CP96) " & _
            " WHERE CP55='" & strTmp & "0' OR CP56='" & strTmp & "0' OR CP72='" & strTmp & "0' OR CP89='" & strTmp & "0' OR CP90='" & strTmp & "0' OR CP91='" & strTmp & "0' OR CP92='" & strTmp & "0' OR CP93='" & strTmp & "0' OR CP94='" & strTmp & "0' OR CP95='" & strTmp & "0' OR CP96='" & strTmp & "0'"
            
         'ADD BY NICKC 2007/01/03
         'Modified by Lydia 2018/10/24 +詳細記錄(True)
         'Modified by Lydai 2025/10/30 +strTmp 將改前客戶編號寫入log; 詳細記錄(True)改用模組判斷 True=>m_bolDesc
         Pub_SeekTbLog strTxt(8), , m_bolDesc, , , strTmp
         
        'Add By Cheng 2002/11/06
        cnnConnection.Execute strTxt(8)
         
         strTxt(9) = "UPDATE CHANGEEVENT SET CE04 = DECODE(CE04,'" & strTmp & "0','" & Label2(0) & "',CE04), " & _
            "CE05 = DECODE(CE05,'" & strTmp & "0','" & Label2(0) & "',CE05), " & _
            "CE06 = DECODE(CE06,'" & strTmp & "0','" & Label2(0) & "',CE06), " & _
            "CE07 = DECODE(CE07,'" & strTmp & "0','" & Label2(0) & "',CE07), " & _
            "CE08 = DECODE(CE08,'" & strTmp & "0','" & Label2(0) & "',CE08) WHERE CE04='" & strTmp & "0' OR CE05='" & strTmp & "0' OR CE06='" & strTmp & "0' OR CE07='" & strTmp & "0' OR CE08='" & strTmp & "0'"
            
         'ADD BY NICKC 2007/01/03
         'Modified by Lydia 2018/10/24 +詳細記錄(True)
         'Modified by Lydai 2025/10/30 +strTmp 將改前客戶編號寫入log; 詳細記錄(True)改用模組判斷 True=>m_bolDesc
         Pub_SeekTbLog strTxt(9), , m_bolDesc, , , strTmp
         
        'Add By Cheng 2002/11/06
        cnnConnection.Execute strTxt(9)
         
      'add by sonia 2025/7/24
      strExc(10) = "UPDATE Datadeleterecord SET DD06='" & Label2(0) & "' WHERE DD06='" & strTmp & "0'"
      cnnConnection.Execute strExc(10)
      'end 2025/7/24
         
        'Modify By Cheng 2002/11/06
'         Screen.MousePointer = vbHourglass
'         If Not objLawDll.ExecSQL(9, strTxt) Then
'            Screen.MousePointer = vbDefault
'            MsgBox "更新新名稱失敗，請洽系統管理員 !", vbCritical
'            Exit Sub
'         End If

         'Add By Sindy 2022/10/24 尚有新案接洽單未收文,更名及加註簽核備註
         strExc(0) = "SELECT * FROM consultrecordlist,ConsultRecApp,ConsultRecCMP" & _
                     " WHERE CRA05='" & Left(ChangeCustomerL(Text1(0).Text), 8) & "'" & _
                     " AND CRA01=CRC01 AND CRC08 is null AND CRA01=CRL01 AND CRL06='Y'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
'            MsgBox "此客戶編號尚有【新案】接洽單未收文，" & vbCrLf & vbCrLf & "請通知智權人員刪除接洽單，待更名後再收文！", vbCritical
'            CmdLock 1
'            TextInverse Text1(0)
'            Text1(0).SetFocus
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               strCRL01 = RsTemp.Fields("CRL01")
               
               '流程備註檔
               strSql = GetInsertFLOW004Sql(strCRL01, strUserNum, strSrvDate(1), Right("000000" & ServerTime, 6), "", _
                  ChangeTStringToTDateString(strSrvDate(2)) & "更名;原名稱=" & RsTemp.Fields("CRA07") & ";" & IIf("" & RsTemp.Fields("CRA08") <> "", "" & RsTemp.Fields("CRA08") & ";", ""))
               cnnConnection.Execute strSql
               '更名
               strSql = "UPDATE ConsultRecApp SET CRA07='" & Trim(Text1(1)) & "'" & _
                        ",CRA08='" & Trim(Text1(8) & " " & Text1(9) & " " & Text1(10) & " " & Text1(11)) & "'" & _
                        " WHERE CRA01='" & strCRL01 & "' AND CRA02=" & RsTemp.Fields("CRA02")
               cnnConnection.Execute strSql
               
               RsTemp.MoveNext
            Loop
         End If
         '2022/10/24 END
         
        cnnConnection.CommitTrans
        
         'Add by Morgan 2006/1/11
         '2011/12/26 MODIFY BY SONIA 加入中日文名稱欄位
         'strExc(0) = Trim(Text1(2) & " " & Text1(3) & " " & Text1(4) & " " & Text1(5))
         'strExc(1) = Trim(Text1(8) & " " & Text1(9) & " " & Text1(10) & " " & Text1(11))
         'If (strExc(1) <> "" And strExc(1) <> strExc(0)) Then
         '   PUB_AccDataCheck Left(Text1(0) & "0000", 9), "英：" & strExc(0) & " --> " & strExc(1)
         'End If
         strExc(0) = Trim(Text1(1) & " " & Text1(6) & " " & Text1(2) & " " & Text1(3) & " " & Text1(4) & " " & Text1(5))
         strExc(1) = Trim(Text1(7) & " " & Text1(12) & " " & Text1(8) & " " & Text1(9) & " " & Text1(10) & " " & Text1(11))
         If (strExc(1) <> "" And strExc(1) <> strExc(0)) Then
            PUB_AccDataCheck Left(Text1(0) & "0000", 9), _
            "中：" & Trim(Text1(1)) & " --> " & Trim(Text1(7)) & Chr(13) & _
            "英：" & Trim(Text1(2) & " " & Text1(3) & " " & Text1(4) & " " & Text1(5)) & " --> " & Trim(Text1(8) & " " & Text1(9) & " " & Text1(10) & " " & Text1(11)) & Chr(13) & _
            "日：" & Trim(Text1(6)) & " --> " & Trim(Text1(12))
         End If
         '2011/12/26 END
         
         Call ChkCustName(strCU13) 'Add By Sindy 2016/4/25 比對收據抬頭檔名稱相同者寄Mail通知財務處
         
         Screen.MousePointer = vbDefault
         CmdLock 1
         CmdAddr.Enabled = True 'Added by Lydia 2018/10/24
         'Modified by Lydia 2018/10/24  取消清空
         'Label2(0) = "": Label2(1) = ""
         'For i = 0 To 12
         '   Text1(i).Text = ""
         'Next
         MsgBox "變更作業完成 !", vbInformation
         'end 2018/10/24
         Text1(0).SetFocus
      Case 1 '結束
         Unload frm140101
         Set frm140101 = Nothing
      Case 2 '尋找
         Label2(0) = "": Label2(1) = ""
         For i = 1 To 12
            Text1(i).Text = ""
         Next
         intI = 1
         '93.11.15 MODIFY BY SONIA
         'strExc(0) = "SELECT CU04,CU05,CU88,CU89,CU90,CU06 FROM CUSTOMER WHERE " & ChgCustomer(Text1(0).Text)
         'Modify By Sindy 2011/1/18 +CU79
         strExc(0) = "SELECT CU04,CU05,CU88,CU89,CU90,CU06,CU32,CU79 FROM CUSTOMER WHERE " & ChgCustomer(Text1(0).Text)
         '93.11.15 END
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For i = 0 To 5
               If Not IsNull(RsTemp.Fields(i).Value) Then
                  Text1(i + 1).Text = RsTemp.Fields(i).Value
                  Text1(i + 7).Text = RsTemp.Fields(i).Value
               End If
            Next
            
            '93.11.15 ADD BY SONIA
            m_CU32 = ""
            If Not IsNull(RsTemp.Fields(6).Value) Then m_CU32 = RsTemp.Fields(6).Value
            '93.11.15 END
            
            'Add By Sindy 2011/1/18
            m_CU79 = ""
            If Not IsNull(RsTemp.Fields("CU79").Value) Then m_CU79 = RsTemp.Fields("CU79").Value
            '2011/1/18 End
            
            CmdLock 0
            strTmp = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
            strExc(0) = "SELECT MAX(CU02) FROM CUSTOMER WHERE CU01='" & strTmp & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            'Modified by Lydia 2018/04/10
            'If intI = 1 Then Label2(0) = strTmp & Format(Val(RsTemp.Fields(0).Value) + 1)
            If intI = 1 Then Label2(0) = strTmp & Format(Val(RsTemp.Fields(0).Value) + 1): Label2(1) = "變更後客戶編號："
         Else
            MsgBox "客戶編號錯誤，請重新輸入 !", vbCritical
            'Add By Cheng 2002/11/18
            CmdLock 1
            TextInverse Text1(0)
            Text1(0).SetFocus
         End If
         
      Case 3 '取消
         If MsgBox("你並未存檔，確定離開 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
         CmdLock 1
         Label2(0) = "": Label2(1) = ""
         For i = 0 To 12
            Text1(i).Text = ""
         Next
         Text1(0).SetFocus
   End Select
   Exit Sub
ErrHand:
   MsgBox "錯誤 : " & Err.Description, vbInformation
'Add By Cheng 2002/11/06
   Exit Sub
ErrorHandler:
    cnnConnection.RollbackTrans
    Screen.MousePointer = vbDefault
    MsgBox "更新新名稱失敗，請洽系統管理員 !", vbCritical
End Sub

'Add By Sindy 2016/4/25 比對收據抬頭檔名稱相同者寄Mail通知財務處
Private Sub ChkCustName(strCU13 As String)
Dim rsTmp As New ADODB.Recordset
Dim strCustID As String, strContext As String
   
   strCustID = "": strContext = ""
   '比對名稱
   If Text1(7) <> "" Then
      strSql = "SELECT a4201" & _
                " FROM acc420" & _
               " WHERE a4201='" & ChgSQL(Text1(7)) & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If IsNull(rsTmp.Fields(0)) = False Then
               strCustID = strCustID & ",'" & rsTmp.Fields(0) & "'"
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   If Text1(8) <> "" Then
      strSql = "SELECT a4201" & _
                " FROM acc420" & _
               " WHERE a4201='" & ChgSQL(UCase(Trim(Trim(Text1(8)) & " " & Trim(Text1(9)) & " " & Trim(Text1(10)) & " " & Trim(Text1(11))))) & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If IsNull(rsTmp.Fields(0)) = False Then
               strCustID = strCustID & ",'" & rsTmp.Fields(0) & "'"
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   If Text1(12) <> "" Then
      strSql = "SELECT a4201" & _
                " FROM acc420" & _
               " WHERE a4201='" & ChgSQL(Text1(12)) & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If IsNull(rsTmp.Fields(0)) = False Then
               strCustID = strCustID & ",'" & rsTmp.Fields(0) & "'"
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   '有相同者寄Mail通知財務處
   If strCustID <> "" Then
      strCustID = Mid(strCustID, 2, Len(strCustID))
      strCustID = Replace(strCustID, "'", "")
      
      strContext = Left(Text1(0) & "00000000", 8) & "0" & vbCrLf
      strContext = strContext & "　　中文名稱：" & Text1(1) & vbCrLf
      strContext = strContext & "　　英文名稱：" & Trim(Trim(Text1(2)) & " " & Trim(Text1(3)) & " " & Trim(Text1(4)) & " " & Trim(Text1(5))) & vbCrLf
      strContext = strContext & "　　日文名稱：" & Text1(6) & vbCrLf
      strContext = strContext & "　　智權人員：" & strCU13 & " " & GetPrjSalesNM(strCU13) & IIf(ChkStaffST04(strCU13, False) = True, "　(已離職)", "") & vbCrLf & vbCrLf
      
      strContext = strContext & "收據抬頭檔：" & strCustID & vbCrLf & vbCrLf
      strContext = strContext & "請向智權人員確認是否相同, 並做後續處理."
      
      'Ｍodify by Amy 2024/08/08 原:Pub_GetSpecMan("財務處總帳人員") 2024/05/15 改財務2個特殊設定拆成3個-未上線
      PUB_SendMail strUserNum, Pub_GetSpecMan("財務處應收處理人員"), "", Left(Text1(0) & "00000000", 8) & "0" & " 與 " & strCustID & "名稱相同 通知 !", strContext
   End If
   Set rsTmp = Nothing
End Sub
'2016/4/25 END

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm140101 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   If Index = 0 Then
      CloseIme
   Else
      OpenIme
   End If
End Sub

'Modified by Lydia 2021/09/23 改成Form 2.0
'Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   If Index = 0 Then KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CmdLock(TF As Integer)
   Select Case TF
      Case 0 '按下尋找時
         Command1(2).Enabled = False
         Command1(0).Enabled = True
         'Add By Cheng 2002/10/25
         '依是否能修改權限以控制按鈕是否可作用
         Command1(0).Enabled = m_bUpdate

         Command1(3).Enabled = True
         Text1(0).Locked = True
         '92.6.9 ADD BY SONIA
         Command1(0).Default = True
         '92.6.9 END
      Case 1 '起始, 按下確定時
         Command1(2).Enabled = True
         Command1(0).Enabled = False
         Command1(3).Enabled = False
         Text1(0).Locked = False
         '92.6.9 ADD BY SONIA
         Command1(2).Default = True
         '92.6.9 END
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   'Remove by Lydia 2021/01/07 名稱長度直接以TextBox.MaxLength控制; ex.R15419的英文名稱"Patentanwaelte · Rechtsanwaelt"字數30字，中英文長度31
'   Select Case Index
'      Case 7
'         If Not CheckLengthIsOK(Text1(Index).Text, 80) Then
'            Text1(7).SetFocus
'            Cancel = True
'         End If
'      Case 8, 9, 10, 11
'         If Not CheckLengthIsOK(Text1(Index).Text, 30) Then
'            Text1(Index).SetFocus
'            Cancel = True
'         End If
'      Case 12
'         If Not CheckLengthIsOK(Text1(Index).Text, 80) Then
'            Text1(Index).SetFocus
'            Cancel = True
'         End If
'   End Select
   'end 2021/01/07
   If Cancel = True Then TextInverse Text1(Index)
   'edit by nickc 2007/06/06 切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
For Each objTxt In Text1
   If objTxt.Enabled = True Then
      Cancel = False
      Text1_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

'Added by Lydia 2021/09/23 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
    Exit Function
End If

TxtValidate = True
End Function

'Added by Lydia 2018/10/24 切換到客戶檔維護(確定變更名稱後,才可執行)
Private Sub CmdAddr_Click()
   Call frm140401.SetParent(Me, Text1(0).Text)
   frm140401.Show
   'Call Command1_Click(1) 'Mark by Lydia 2024/02/22
End Sub

