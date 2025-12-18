VERSION 5.00
Begin VB.Form frm880017 
   BorderStyle     =   1  '單線固定
   Caption         =   "補件期限"
   ClientHeight    =   5184
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   6036
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5184
   ScaleWidth      =   6036
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.ComboBox Combo2 
      Height          =   300
      Index           =   2
      ItemData        =   "frm880017.frx":0000
      Left            =   1125
      List            =   "frm880017.frx":0002
      TabIndex        =   32
      Top             =   4140
      Width           =   3480
   End
   Begin VB.TextBox txtCP 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   288
      Index           =   1
      Left            =   1110
      MaxLength       =   3
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   510
      Width           =   492
   End
   Begin VB.TextBox txtCP 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   288
      Index           =   2
      Left            =   1590
      MaxLength       =   6
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   510
      Width           =   732
   End
   Begin VB.TextBox txtCP 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   288
      Index           =   3
      Left            =   2310
      MaxLength       =   1
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   510
      Width           =   252
   End
   Begin VB.TextBox txtCP 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   288
      Index           =   4
      Left            =   2550
      MaxLength       =   2
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   510
      Width           =   372
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Index           =   1
      ItemData        =   "frm880017.frx":0004
      Left            =   1125
      List            =   "frm880017.frx":0006
      TabIndex        =   24
      Top             =   3840
      Width           =   3480
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Index           =   0
      ItemData        =   "frm880017.frx":0008
      Left            =   1125
      List            =   "frm880017.frx":000A
      TabIndex        =   23
      Top             =   3540
      Width           =   3480
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Index           =   4
      ItemData        =   "frm880017.frx":000C
      Left            =   1125
      List            =   "frm880017.frx":000E
      TabIndex        =   21
      Top             =   4740
      Width           =   3480
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Index           =   3
      ItemData        =   "frm880017.frx":0010
      Left            =   1125
      List            =   "frm880017.frx":0012
      TabIndex        =   19
      Top             =   4440
      Width           =   3480
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   5
      ItemData        =   "frm880017.frx":0014
      Left            =   1125
      List            =   "frm880017.frx":0016
      TabIndex        =   17
      Top             =   3000
      Width           =   3480
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   4
      ItemData        =   "frm880017.frx":0018
      Left            =   1125
      List            =   "frm880017.frx":001A
      TabIndex        =   15
      Top             =   2700
      Width           =   3480
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   3
      ItemData        =   "frm880017.frx":001C
      Left            =   1125
      List            =   "frm880017.frx":001E
      TabIndex        =   13
      Top             =   2400
      Width           =   3480
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   2
      ItemData        =   "frm880017.frx":0020
      Left            =   1125
      List            =   "frm880017.frx":0022
      TabIndex        =   11
      Top             =   2100
      Width           =   3480
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   4695
      TabIndex        =   5
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   3870
      TabIndex        =   4
      Top             =   60
      Width           =   800
   End
   Begin VB.TextBox txtDate 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1125
      MaxLength       =   8
      TabIndex        =   3
      Top             =   840
      Width           =   1005
   End
   Begin VB.TextBox txtDate 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1125
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1170
      Width           =   1005
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   1
      ItemData        =   "frm880017.frx":0024
      Left            =   1125
      List            =   "frm880017.frx":0026
      TabIndex        =   1
      Top             =   1800
      Width           =   3480
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   0
      ItemData        =   "frm880017.frx":0028
      Left            =   1125
      List            =   "frm880017.frx":002A
      TabIndex        =   0
      Top             =   1485
      Width           =   3480
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   285
      Index           =   10
      Left            =   4680
      TabIndex        =   35
      Top             =   4770
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   285
      Index           =   9
      Left            =   4680
      TabIndex        =   34
      Top             =   4470
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   285
      Index           =   8
      Left            =   4680
      TabIndex        =   33
      Top             =   4170
      Width           =   1185
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "補正內容"
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   31
      Top             =   3570
      Width           =   720
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   285
      Left            =   3285
      TabIndex        =   30
      Top             =   840
      Width           =   2220
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號"
      Height          =   180
      Left            =   180
      TabIndex        =   29
      Top             =   510
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   285
      Index           =   7
      Left            =   4680
      TabIndex        =   22
      Top             =   3870
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   285
      Index           =   6
      Left            =   4680
      TabIndex        =   20
      Top             =   3570
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   285
      Index           =   5
      Left            =   4680
      TabIndex        =   18
      Top             =   3000
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   285
      Index           =   4
      Left            =   4680
      TabIndex        =   16
      Top             =   2700
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   285
      Index           =   3
      Left            =   4680
      TabIndex        =   14
      Top             =   2400
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   285
      Index           =   2
      Left            =   4680
      TabIndex        =   12
      Top             =   2100
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   285
      Index           =   1
      Left            =   4680
      TabIndex        =   10
      Top             =   1800
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   285
      Index           =   0
      Left            =   4680
      TabIndex        =   9
      Top             =   1515
      Width           =   1185
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "補件內容"
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   8
      Top             =   1485
      Width           =   720
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "法定期限"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   7
      Top             =   1170
      Width           =   735
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "本所期限"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   840
      Width           =   720
   End
End
Attribute VB_Name = "frm880017"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/16 Form2.0已檢查 (無需修改的物件);
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan2010/8/19 日期欄已修改
'Create by Morgan 2009/11/10
'1.已存檔的補件期限不論已收或未收文都不可修改內容只會更新期限。
'2.新增的補件期限如相關收文號為C類來函則因尚未有收文號故設定本畫面不存檔，不論新增或即有的期限都回呼叫畫面於存檔時一併更新。
Option Explicit

Public m_CP43 As String, m_CP06 As String, m_CP07 As String
Public m_siSaveFlag As Single '存檔標記
Public m_stUnSaveData As String, m_bolNoAdd As Boolean
'Added by Morgan 2014/3/5
Public m_bolAddFix As Boolean '是否加補正選單
Public m_stUnSaveData2 As String '補正期限
'Added by Lydia 2015/05/01 (傳) 是否為FMP
Dim m_bolFMP As Boolean
'Added by Morgan 2015/8/26
Dim m_PA09 As String '申請國家

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         Unload Me
         
      Case 1
         If FormSave = True Then
            Unload Me
         End If
   End Select
End Sub

Private Function FormSave() As Boolean
   Dim oCombo As ComboBox
   Dim strCP06 As String 'Added by Morgan 2016/12/28
   'Added by Lydia 2015/05/01
   Dim bolOne As Boolean
   
   m_siSaveFlag = 0 'Added by Morgan 2018/8/14
   
        '防止重複
        bolOne = False
        For Each oCombo In Combo1
          If oCombo.Text = "轉讓證明" Then
             If bolOne = False Then
                bolOne = True
             Else
                If Trim(txtCP(1)) = "P" And oCombo.Text = "轉讓證明" And m_bolFMP = False Then
                   MsgBox "重複輸入轉讓證明!!", vbCritical
                   Exit Function
                End If
             End If
          End If
        Next
   'end 2015/05/01
   
   If m_bolNoAdd Then
      For Each oCombo In Combo1
         If oCombo.Text <> "" Then
            If oCombo.Tag = "" Then
               m_stUnSaveData = m_stUnSaveData & IIf(m_stUnSaveData = "", "", vbCrLf) & oCombo.Text
               
            'Added by Morgan 2018/8/31 已有期限則可不必點選新的補件內容--敏莉　Ex.P-120707
            ElseIf Left(oCombo.Tag, 1) <> "1" Then
               m_siSaveFlag = 1
            'end 2018/8/31
            End If
         End If
      Next
         
      'Modified by Morgan 2018/8/13 一定要輸入補件內容 --玲玲 Ex:P-120272
      'm_siSaveFlag = 1
      If m_stUnSaveData <> "" Then m_siSaveFlag = 1
      'end 2018/8/13
      
      'Added by Morgan 2014/3/5
      m_stUnSaveData2 = ""
      For Each oCombo In Combo2
         If oCombo.Tag = "" And oCombo.Text <> "" Then
            m_stUnSaveData2 = m_stUnSaveData2 & IIf(m_stUnSaveData2 = "", "", vbCrLf) & oCombo.Text
         End If
      Next
      '若只有補正期限時控制不必新增通知補文件來函
      If m_stUnSaveData2 <> "" And Combo1(0) = "" And m_stUnSaveData = "" Then
         m_siSaveFlag = 2
      End If
      'end 2014/3/5
      
      'Modified by Morgan 2018/8/14
      'FormSave = True
      If m_siSaveFlag = 0 Then
         MsgBox "請輸入補件內容！", vbExclamation
      Else
         FormSave = True
      End If
      'end 2018/8/14
      
      Exit Function
   End If
   
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   
   For Each oCombo In Combo1
      'Added by Lydia 2015/05/01 P案國外指示信補件期限點選「轉讓證明」,自動產生一道Ｂ類收文(240補轉讓證明)承辦人掛PS2，期限設定為3個月,本所提前10天。
      If Trim(txtCP(1)) = "P" And oCombo.Text = "轉讓證明" And m_bolFMP = False Then
          '補轉讓證明法限=系統日+3個月,所限提早10天
          strExc(7) = CompDate(1, 3, strSrvDate(1)) 'PUB_GetWorkDay1(CompDate(1, 3, strSrvDate(1)), True)
          strExc(6) = PUB_GetWorkDay1(CompDate(2, -10, strExc(7)), True)
          'Added by Morgan 2025/1/21
          If strSrvDate(1) >= P業務區劃分啟用日 Then
            strExc(2) = PUB_GetPHandler(txtCP(1) & txtCP(2) & txtCP(3) & txtCP(4))
          Else
          'end 2025/1/21
          
            strExc(2) = Pub_GetSpecMan("PS2")
            
          End If 'Added by Morgan 2025/1/21
          strExc(1) = AutoNo("B", 6)
          strExc(0) = "INSERT INTO CASEPROGRESS (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp43,cp64) " & _
                "SELECT cp01,cp02,cp03,cp04," & CNULL(strSrvDate(1), True) & "," & CNULL(strExc(6), True) & "," & CNULL(strExc(7), True) & ",'" & strExc(1) & "','240',cp12,cp13,'" & strExc(2) & "',CP09,'補優先權轉讓證明' " & _
                "FROM CASEPROGRESS WHERE CP09='" & m_CP43 & "' "
          cnnConnection.Execute strExc(0), intI
      'end 2021/05/01
      'Added by Lydia 2021/04/01 FMP寰華國外指示信補件期限點選期限控管：選「優先權轉讓證明」，請自動產生期限於下一程序(240補優先權轉讓證明)，期限先抓新案進度之發文日加3個月。
      ElseIf m_bolFMP = True And oCombo.Text = "優先權轉讓證明" Then
           'Modified by Lydia 2021/05/04
           'strExc(0) = "select cp09,cp158 from caseprogress where cp01='" & txtCP(1) & "' and cp02='" & txtCP(2) & "' and cp03='" & txtCP(3) & "' and cp04='" & txtCP(4) & "' and substr(cp09,1,1) < 'C' and cp31='Y' and cp159=0 and cp158>0 "
           strExc(0) = "select cp09,cp158 from caseprogress where cp01='" & txtCP(1) & "' and cp02='" & txtCP(2) & "' and cp03='" & txtCP(3) & "' and cp04='" & txtCP(4) & "' and substr(cp09,1,1) < 'C' and cp31='Y' and cp159=0 "
           intI = 1
           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
           If intI = 1 Then
           'Modified by Lydia 2021/05/04 新案皆沒有發文日，故不應以新案發文日計算。期限應同委任書一樣，點選補件期限240時請預設本所期限為系統日並+3個月以及往前抓工作天。（正確期限將於輸入通知申請案號時，更新期限為提申日+3個月，本所期限為法定期限前10天。）
           '    strExc(1) = "" & RsTemp.Fields("cp158")
           '    If strExc(1) = "19221111" Then '假發文
           '        strExc(1) = strSrvDate(1)
           '    End If
           '    strExc(2) = CompDate(1, 3, strExc(1))
           '    If strExc(2) <= strSrvDate(1) Then strExc(2) = strSrvDate(1)
               strExc(2) = PUB_GetWorkDay1(CompDate(1, 3, strSrvDate(1)), True)
           'end 2021/05/04
               strExc(1) = GetNextProgressNo
               strSql = "INSERT INTO NEXTPROGRESS(NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP15,NP22) " & _
                           "VALUES (" & CNULL(RsTemp.Fields("cp09")) & ", " & CNULL(txtCP(1)) & ", " & CNULL(txtCP(2)) & ", " & CNULL(txtCP(3)) & ", " & CNULL(txtCP(4)) & ", '240', " & CNULL(strExc(2)) & ", " & CNULL(strExc(2)) & ", " & CNULL(strUserNum) & ", " & CNULL(ChgSQL(oCombo.Text)) & ", " & CNULL(strExc(1)) & ") "
               cnnConnection.Execute strSql, intI
           End If 'Remove by Lydia 2021/05/04
      'end 2021/04/01
      Else
      
            strCP06 = CNULL(txtDate(0), True)
            'Added by Morgan 2016/12/28 若有指定提申時以該日期設定為"線條清晰之圖式"補文件的本所期限 -- 品薇
            If oCombo.Text = "線條清晰之圖式" Then
               strExc(0) = "select np08 from nextprogress where np01='" & m_CP43 & "' and np07='995' and np06 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strCP06 = RsTemp(0)
               End If
            End If
            
            Select Case Left(oCombo.Tag, 1)
               Case "2" '已收文未發文
                  strSql = "UPDATE CASEPROGRESS SET CP06=" & strCP06 & ",CP07=" & CNULL(txtDate(1), True) & _
                     " WHERE CP09=" & CNULL(Mid(oCombo.Tag, 3))
                  cnnConnection.Execute strSql, intI
                  
               Case "3" '未收文
                  If oCombo.Text <> "" Then
                     strSql = "UPDATE NEXTPROGRESS SET NP08=" & strCP06 & ",NP09=" & CNULL(txtDate(1), True) & _
                        " WHERE NP01=" & CNULL(m_CP43) & " AND NP07='202' AND NP22=" & CNULL(Mid(oCombo.Tag, 3), True)
                     cnnConnection.Execute strSql, intI
                  End If
                  
               Case Else
                  If oCombo.Text <> "" Then
                     'Added by Morgan 2015/8/26
                     '非FMP案改自動內部收文
                     If txtCP(1) = "P" And Not m_bolFMP Then
                        strExc(1) = AutoNo("B", 6)
                        'Added by Morgan 2025/1/21
                        If strSrvDate(1) >= P業務區劃分啟用日 Then
                          strExc(2) = PUB_GetPHandler(txtCP(1) & txtCP(2) & txtCP(3) & txtCP(4))
                        Else
                        'end 2025/1/21
                           strExc(2) = Pub_GetSpecMan("PS2")
                        End If 'Added by Morgan 2025/1/21
                        strSql = "INSERT INTO CASEPROGRESS (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp43,cp64) " & _
                           "SELECT cp01,cp02,cp03,cp04," & strSrvDate(1) & "," & strCP06 & "," & CNULL(txtDate(1), True) & ",'" & strExc(1) & "','202',cp12,cp13,'" & strExc(2) & "',cp09,'" & ChgSQL(oCombo.Text) & "' " & _
                           "FROM CASEPROGRESS WHERE CP09='" & m_CP43 & "' "
                        cnnConnection.Execute strSql, intI
                     Else
                     'end 2015/8/26
                        strSql = "INSERT INTO NEXTPROGRESS(NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP15,NP22)" & _
                           " SELECT CP09,CP01,CP02,CP03,CP04,'202'" & _
                           "," & strCP06 & "," & CNULL(txtDate(1), True) & ",CP13," & CNULL(ChgSQL(oCombo.Text)) & _
                           ",NP22 FROM CASEPROGRESS,(SELECT MAX(NP22)+1 NP22 FROM NEXTPROGRESS) WHERE CP09=" & CNULL(m_CP43)
                        cnnConnection.Execute strSql, intI
                     End If
                  End If
            End Select
      End If
   Next
   cnnConnection.CommitTrans
   m_siSaveFlag = 1
   FormSave = True
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If
End Function


Private Sub Form_Load()
   Dim oCombo As ComboBox, oLabel As LABEL
   
   MoveFormToCenter Me
   
   For Each oCombo In Combo1
      oCombo.AddItem "委託書"
      oCombo.AddItem "轉讓證明"
      If m_bolAddFix = False Then 'Added by Morgan 2014/3/5
         oCombo.AddItem "線條清晰之圖式"
      End If
      'oCombo.AddItem "優先權證明"
   Next
   
   'Added by Morgan 2014/3/5
   For Each oCombo In Combo2
      oCombo.AddItem "線條清晰之圖式"
      oCombo.AddItem "摘要譯文(與國際公布文本中記載不一致)"
   Next
   'end 2014/3/5
   
   For Each oLabel In Label1
      oLabel.Caption = "未收文"
   Next
   
   Label3 = ""
   'Added by Lydia 2015/05/01 + FMP 判斷
   'strExc(0) = "select cp01,cp02,cp03,cp04 from caseprogress where cp09='" & m_CP43 & "'"
   strExc(0) = "select cp01,cp02,cp03,cp04,cp12,pa09 from caseprogress,patent where cp09='" & m_CP43 & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      txtCP(1) = RsTemp.Fields(0)
      txtCP(2) = RsTemp.Fields(1)
      txtCP(3) = RsTemp.Fields(2)
      txtCP(4) = RsTemp.Fields(3)
      txtDate(0) = m_CP06
      txtDate(1) = m_CP07
      m_PA09 = RsTemp.Fields("pa09") 'Added by Morgan 2015/8/26
      
      'Added by Lydia 2015/05/01
      'Modified by Morgan 2021/2/2
      'If Trim(txtCP(1)) = "P" And Left(RsTemp.Fields("CP12"), 1) = "F" And RsTemp.Fields("pa09") <> "000" Then
      '   m_bolFMP = True
      'Else
      '   m_bolFMP = False
      'End If
      m_bolFMP = PUB_ChkIsFMP(txtCP(1), txtCP(2), txtCP(3), txtCP(4), m_PA09)
      'end 2021/2/2
      
      ReadData
   End If
   'end 2015/05/01
   
   'Added by Morgan 2014/3/5
   If m_bolAddFix = False Then
      Me.Height = 3870
   End If
   
   'Added by Lydia 2021/04/01 FMP寰華國外指示信補件期限點選期限控管
   If txtCP(1) = "P" And m_bolFMP = True Then
        For Each oCombo In Combo1
           oCombo.AddItem "優先權轉讓證明"
        Next
        For Each oCombo In Combo2
           oCombo.AddItem "補優先權轉讓證明期限"
        Next
   End If
   'end 2021/04/01
End Sub

Private Sub ReadData()
   Dim ii As Integer, arrData() As String
   
   ii = 0
   If m_CP43 <> "" Then
      'Modify by Morgan 2010/3/25 排除優先權證明
      strExc(0) = "select '1' SRC,cp09,cp64,cp06 from caseprogress where cp43='" & m_CP43 & "' and cp10='202' and cp57 is null and cp27>0 and instr(cp64,'優先權證明')=0" & _
           " union select '2' SRC,cp09,cp64,cp06 from caseprogress where cp43='" & m_CP43 & "' and cp10='202' and cp57 is null and cp27 is null and instr(cp64,'優先權證明')=0" & _
           " union select '3' SRC,To_char(np22) cp09,np15 cp64,np08 cp06 from nextprogress where np01='" & m_CP43 & "' and np07||np06='202' and instr(np15,'優先權證明')=0" & _
           " order by SRC,cp09"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         
         With RsTemp
         If .Fields("cp06") <> m_CP06 Then
            Label3 = "原本所期限 " & (.Fields("cp06") - 19110000)
         End If
         Do While Not .EOF
            Combo1(ii).Text = .Fields("CP64")
            If .Fields("SRC") = "1" Then
               Label1(ii).Caption = "已發文"
               Combo1(ii).Enabled = False
            ElseIf .Fields("SRC") = "2" Then
               Label1(ii).Caption = "已收文未發文"
               Combo1(ii).Enabled = False
            'Add by Morgan 2009/11/19 舊的期限都不可修改
            Else
               Combo1(ii).Enabled = False
            End If
            Combo1(ii).Tag = .Fields("SRC") & "," & .Fields("cp09")
            ii = ii + 1
            .MoveNext
         Loop
         End With
      End If
   End If
   
   If m_stUnSaveData <> "" Then
      arrData = Split(m_stUnSaveData, vbCrLf)
      For intI = LBound(arrData) To UBound(arrData)
         If arrData(intI) <> "" Then
            Combo1(ii).Text = arrData(intI)
            ii = ii + 1
         End If
      Next
   End If
   
   'Added by Morgan 2014/3/5
   If m_stUnSaveData2 <> "" Then
      arrData = Split(m_stUnSaveData2, vbCrLf)
      ii = 0
      For intI = LBound(arrData) To UBound(arrData)
         If arrData(intI) <> "" Then
            Combo2(ii).Text = arrData(intI)
            ii = ii + 1
         End If
      Next
   End If
   'end 2014/3/5
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '因為要回傳結束狀態,改由呼叫的視窗清除
   'Set frm880017 = Nothing
End Sub
