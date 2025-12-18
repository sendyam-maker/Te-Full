VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140103 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人變更名稱作業"
   ClientHeight    =   5808
   ClientLeft      =   576
   ClientTop       =   972
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
      Left            =   3880
      TabIndex        =   14
      Top             =   72
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   350
      Index           =   2
      Left            =   2490
      TabIndex        =   1
      Top             =   575
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   5100
      TabIndex        =   15
      Top             =   72
      Width           =   912
   End
   Begin VB.CommandButton Command1 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   7020
      TabIndex        =   17
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消(&C)"
      Height          =   400
      Index           =   3
      Left            =   6060
      TabIndex        =   16
      Top             =   60
      Width           =   912
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   11
      Left            =   1290
      TabIndex        =   12
      Top             =   4861
      Width           =   4095
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   10
      Left            =   1290
      TabIndex        =   11
      Top             =   4535
      Width           =   4095
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   9
      Left            =   1290
      TabIndex        =   10
      Top             =   4209
      Width           =   4095
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   5
      Left            =   1290
      TabIndex        =   6
      Top             =   2455
      Width           =   4095
      VariousPropertyBits=   671105055
      BackColor       =   -2147483648
      MaxLength       =   30
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   4
      Left            =   1290
      TabIndex        =   5
      Top             =   2129
      Width           =   4095
      VariousPropertyBits=   671105055
      BackColor       =   -2147483648
      MaxLength       =   30
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   3
      Left            =   1290
      TabIndex        =   4
      Top             =   1803
      Width           =   4095
      VariousPropertyBits=   671105055
      BackColor       =   -2147483648
      MaxLength       =   30
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   525
      Index           =   6
      Left            =   1290
      TabIndex        =   7
      Top             =   2781
      Width           =   6615
      VariousPropertyBits=   -1467989985
      BackColor       =   -2147483648
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
      Index           =   0
      Left            =   1290
      TabIndex        =   0
      Top             =   600
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
      Height          =   300
      Index           =   2
      Left            =   1290
      TabIndex        =   3
      Top             =   1477
      Width           =   4095
      VariousPropertyBits=   671105055
      BackColor       =   -2147483648
      MaxLength       =   30
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   525
      Index           =   1
      Left            =   1290
      TabIndex        =   2
      Top             =   926
      Width           =   6615
      VariousPropertyBits=   -1467989985
      BackColor       =   -2147483648
      MaxLength       =   80
      ScrollBars      =   2
      Size            =   "11668;926"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   525
      Index           =   7
      Left            =   1290
      TabIndex        =   8
      Top             =   3332
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
      Index           =   8
      Left            =   1290
      TabIndex        =   9
      Top             =   3883
      Width           =   4095
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   525
      Index           =   12
      Left            =   1290
      TabIndex        =   13
      Top             =   5190
      Width           =   6615
      VariousPropertyBits=   -1467989989
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
      Caption         =   "變更後代理人編號："
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   4350
      TabIndex        =   26
      Top             =   600
      Width           =   1620
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   6150
      TabIndex        =   25
      Top             =   600
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "原日文名稱："
      Height          =   180
      Index           =   4
      Left            =   90
      TabIndex        =   24
      Top             =   2820
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "原英文名稱："
      Height          =   180
      Index           =   2
      Left            =   90
      TabIndex        =   23
      Top             =   1537
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "原中文名稱："
      Height          =   180
      Index           =   1
      Left            =   90
      TabIndex        =   22
      Top             =   990
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代理人編號："
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   21
      Top             =   660
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "新日文名稱："
      Height          =   180
      Index           =   3
      Left            =   90
      TabIndex        =   20
      Top             =   5220
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "新英文名稱："
      Height          =   180
      Index           =   5
      Left            =   90
      TabIndex        =   19
      Top             =   3943
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "新中文名稱："
      Height          =   180
      Index           =   6
      Left            =   90
      TabIndex        =   18
      Top             =   3390
      Width           =   1080
   End
End
Attribute VB_Name = "frm140103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim m_FA24 As String
Dim m_FA29 As String 'Add By Sindy 2011/1/18
Dim m_bolDesc As Boolean 'Added by Lydia 2025/10/30

Private Sub Form_Load()
 
   ' 90.07.16 modify by Ken (取得使用者執行各項功能的權限)
   m_bUpdate = IsUserHasRightOfFunction("frm140103", strEdit, False)
   ' Ken 90.07.16 -- End
   
   MoveFormToCenter Me
   CmdLock 1
   Label2(0) = "": Label2(1) = ""
   
   ' Ken 90.07.16 -- start
   'Remove by Lydia 2018/10/24 改在CmdLock控制
   'If m_bUpdate Then
   '    Command1(0).Enabled = True
   'Else
   '    Command1(0).Enabled = False
   'End If
   ' Ken 90.07.16 -- End
   'end 2018/10/24
   
   'Added by Lydia 2018/10/24 檢查權限
   If IsUserHasRightOfFunction("frm050705", strEdit, False) = True Then
        CmdAddr.Visible = True
   Else
        CmdAddr.Visible = False
   End If
   CmdAddr.Enabled = False '確定變更名稱後,才可執行
   'end 2018/10/24
End Sub

Private Sub Command1_Click(Index As Integer)
 Dim i As Integer, St(1 To 6) As String, strTmp As String
 'add by nickc 2006/06/15
 Dim FANation As String
 Dim orsTmp As New ADODB.Recordset
 Dim strOldName As String
 
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
         FANation = ""
         
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         'Added by Lydia 2025/10/30 改用模組判斷
         m_bolDesc = PUB_FilterSeekSQL("", Me)
         
        'Add By Cheng 2002/11/06
        On Error GoTo ErrorHandler
        cnnConnection.BeginTrans
        
         strTmp = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
         
         'add by nickc 2006/06/15
         strSql = "select * from fagent WHERE fa01='" & strTmp & "' AND fa02='0' "
         Set orsTmp = New ADODB.Recordset
         orsTmp.CursorLocation = adUseClient
         orsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If orsTmp.RecordCount <> 0 Then
            FANation = CheckStr(orsTmp.Fields("fa10"))
'edit by nickc 2008/05/08  改成 function
'            If Mid(FANation, 1, 3) = "101" Then
'                '2008/2/29 MODIFY BY SONIA 原為A~I為101,J~Z為1011,2008年改為分四段
'                If Mid(UCase(Text1(8)), 1, 1) >= "A" And Mid(UCase(Text1(8)), 1, 1) <= "E" Then
'                     FANation = "101"
'                ElseIf Mid(UCase(Text1(8)), 1, 1) >= "F" And Mid(UCase(Text1(8)), 1, 1) <= "I" Then
'                     FANation = "1011"
'                ElseIf Mid(UCase(Text1(8)), 1, 1) >= "J" And Mid(UCase(Text1(8)), 1, 1) <= "N" Then
'                     FANation = "1012"
'                ElseIf Mid(UCase(Text1(8)), 1, 1) >= "O" And Mid(UCase(Text1(8)), 1, 1) <= "Z" Then
'                     FANation = "1013"
'                Else
'                     FANation = "1013"
'                End If
'            ElseIf Mid(FANation, 1, 3) = "011" Then
'                '2008/4/21 MODIFY BY SONIA 原為A~L為011,M~Z為0111,2008/4/22改為分三段(將M~Z再細分成二段)
'                If Mid(UCase(Text1(8)), 1, 1) >= "A" And Mid(UCase(Text1(8)), 1, 1) <= "L" Then
'                     FANation = "011"
'                ElseIf Mid(UCase(Text1(8)), 1, 1) >= "M" And Mid(UCase(Text1(8)), 1, 1) <= "O" Then
'                     FANation = "0111"
'                ElseIf Mid(UCase(Text1(8)), 1, 1) >= "P" And Mid(UCase(Text1(8)), 1, 1) <= "Z" Then
'                     FANation = "0112"
'                ElseIf Trim(Text1(8)) = "" Then
'                     FANation = "0112"
'                End If
'            End If
            FANation = pub_NationByName(Text1(8) & Text1(9) & Text1(10) & Text1(11), FANation)
         End If
         
         '92.9.24 MODIFY BY SONIA
         'strExc(1) = "UPDATE FAGENT SET FA24='N',FA02='" & Right(Label2(0), 1) & "' WHERE FA01='" & strTmp & "' AND FA02='0'"
         'Modify by Morgan 2006/10/23 FA24 要放在
         'Modify By Sindy 2011/1/18 調整FA29
         strExc(1) = "UPDATE FAGENT SET FA24='N',FA02='" & Right(Label2(0), 1) & "',FA29='" & ChangeTStringToTDateString(strSrvDate(2)) & "更名'||decode(FA29,'',';',decode(substr(FA29,1,1),';',FA29,';'||FA29)) WHERE FA01='" & strTmp & "' AND FA02='0'"
         '92.9.24 END
         
         'ADD BY NICKC 2007/01/03
         'Modified by Lydia 2018/10/24 +詳細記錄(True)
         'Modified by Lydai 2025/10/30 +strTmp 將改前客戶編號寫入log; 詳細記錄(True)改用模組判斷 True=>m_bolDesc
         Pub_SeekTbLog strExc(1), , m_bolDesc, , , strTmp
         
         'Add By Cheng 2002/11/06
         cnnConnection.Execute strExc(1)
         
         strExc(2) = "INSERT INTO FAGENT (FA01"
         strSql = "SELECT FA01"
         For intI = 2 To TF_FA
            '除Create(Update) ID, Date, Time 以外都要
            If intI < 46 Or intI > 51 Then
               strExc(2) = strExc(2) & ",FA" & Format(intI, "0#")
               Select Case intI
                  Case 2 '變更碼
                     strSql = strSql & ",'0'"
                  Case 4 '中文名稱
                     strSql = strSql & "," & CNULL(ChgSQL(St(1)))
                  Case 5 '英文名稱
                     strSql = strSql & "," & CNULL(ChgSQL(St(2)))
                  Case 63 '英文名稱
                     strSql = strSql & "," & CNULL(ChgSQL(St(3)))
                  Case 64 '英文名稱
                     strSql = strSql & "," & CNULL(ChgSQL(St(4)))
                  Case 65 '英文名稱
                     strSql = strSql & "," & CNULL(ChgSQL(St(5)))
                  Case 6 '日文名稱
                     strSql = strSql & "," & CNULL(ChgSQL(St(6)))
                  Case 10 '國籍
                     strSql = strSql & "," & CNULL(FANation)
                  Case 24 '是否寄雜誌
                     strSql = strSql & "," & CNULL(m_FA24)
                  Case 29 '代理人備註 Add By Sindy 98/02/16
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
                     'Modify By Sindy 2011/1/18 調整FA29
                     If m_FA29 = "" Or Left(Trim(m_FA29), 1) <> ";" Then
                        m_FA29 = ";" & m_FA29
                     End If
                     strSql = strSql & ",'" & ChangeTStringToTDateString(strSrvDate(2)) & "更名(" & "舊名稱：" & ChgSQL(Trim(strOldName)) & ")" & ChgSQL(m_FA29) & "'"
                  Case Else
                     strSql = strSql & ",FA" & Format(intI, "0#")
               End Select
            End If
         Next
         strExc(2) = strExc(2) & ") " & strSql & " FROM FAGENT WHERE FA01='" & strTmp & "' AND FA02='" & Right(Label2(0), 1) & "'"
         'end 2006/10/23
         
         'ADD BY NICKC 2007/01/03
         'Modified by Lydia 2018/10/24 +詳細記錄(True)
         'Modified by Lydai 2025/10/30 +strTmp 將改前客戶編號寫入log; 詳細記錄(True)改用模組判斷 True=>m_bolDesc
         Pub_SeekTbLog strExc(2), , m_bolDesc, , , strTmp
         
         'Add By Cheng 2002/11/06
         cnnConnection.Execute strExc(2)
         '92.3.23 ADD BY SONIA"C" & Left(strSrvDate(2), 2)
         '92.9.29 cancel by sonia
         'strExc(3) = "UPDATE FAGENT SET FA29=FA29||DECODE(FA29,'','',',')||'" & ChangeTStringToTDateString(strSrvDate(2)) & "更名' WHERE FA01='" & strTmp & "' AND FA02='" & Right(Label2(0), 1) & "'"
         'cnnConnection.Execute strExc(3)
         '92.9.29 end
         '92.3.23 END
         
        'Modify By Cheng 2002/11/06
'         If Not objLawDll.ExecSQL(2, strExc) Then
'            MsgBox "更新新名稱失敗，請洽系統管理員 !", vbCritical
'            Exit Sub
'         End If
        cnnConnection.CommitTrans
        
'cancel by sonia 2020/2/14 代理人維護已於2017/08/11取消
'      'add by nickc 2006/12/26 若有改名稱，地址，電話，傳真將列出該代理人國籍，編號，名稱，3個月內期限，6個月內期限，所有案件數
'      Dim strScanFagent As String
'      'add by nickc 2006/12/28
'      Dim intLine As Integer
'      Dim intLineCnt As Integer
'      Dim nowCnt As Integer
'      Dim Seek01 As String
'      Dim strFA01 As String
'      Dim strFA02 As String
'      strFA01 = strTmp
'      strFA02 = "0"
'      intLineCnt = 4
'    strScanFagent = " select fnation,fnum,fname,sum(c3) as C3,sum(c6) as C6,ccount from ("
'    strScanFagent = strScanFagent & " select distinct 'ok' as key, 1 as c3,0 as c6,cp01||cp02||cp03||cp04 from caseprogress where (cp01,cp02,cp03,cp04) in ("
'    strScanFagent = strScanFagent & " select pa01,pa02,pa03,pa04 from patent where pa75='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select tm01,tm02,tm03,tm04 from trademark where tm44='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select sp01,sp02,sp03,sp04 from servicepractice where sp26='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select lc01,lc02,lc03,lc04 from lawcase where lc22='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select cp01,cp02,cp03,cp04 from caseprogress where cp44='" & strFA01 & strFA02 & "')"
'    strScanFagent = strScanFagent & " and cp06>=to_number(to_char(sysdate,'YYYYMMDD')) and cp06<=to_number(to_char(add_months(sysdate,3),'YYYYMMDD')) and cp27 is null and cp57 is null"
'    strScanFagent = strScanFagent & " union select 'ok' as key,1 as c3,0 as c6,np02||np03||np04||np05 from nextprogress where (np02,np03,np04,np05) in ("
'    strScanFagent = strScanFagent & " select pa01,pa02,pa03,pa04 from patent where pa75='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select tm01,tm02,tm03,tm04 from trademark where tm44='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select sp01,sp02,sp03,sp04 from servicepractice where sp26='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select lc01,lc02,lc03,lc04 from lawcase where lc22='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select cp01,cp02,cp03,cp04 from caseprogress where cp44='" & strFA01 & strFA02 & "')"
'    'Modify By Sindy 2009/07/24 增加LIN系統類別
'    '2010/3/23 MODIFY BY SONIA 剔除下一程序非智權人員掌控之案件性質改以strNpSqlOfNoSalesDuty控制
'    strScanFagent = strScanFagent & " and np06 is null and np08>=to_number(to_char(sysdate,'YYYYMMDD')) and np08<=to_number(to_char(add_months(sysdate,3),'YYYYMMDD')) " & strNpSqlOfNoSalesDuty
'    strScanFagent = strScanFagent & " union select 'ok' as key,0 as c3,1 as c6,cp01||cp02||cp03||cp04 from caseprogress where (cp01,cp02,cp03,cp04) in ("
'    strScanFagent = strScanFagent & " select pa01,pa02,pa03,pa04 from patent where pa75='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select tm01,tm02,tm03,tm04 from trademark where tm44='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select sp01,sp02,sp03,sp04 from servicepractice where sp26='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select lc01,lc02,lc03,lc04 from lawcase where lc22='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select cp01,cp02,cp03,cp04 from caseprogress where cp44='" & strFA01 & strFA02 & "')"
'    strScanFagent = strScanFagent & " and cp06>=to_number(to_char(sysdate,'YYYYMMDD')) and cp06<=to_number(to_char(add_months(sysdate,6),'YYYYMMDD')) and cp27 is null and cp57 is null"
'    strScanFagent = strScanFagent & " union select 'ok' as key,0 as c3,1 as c6,np02||np03||np04||np05 from nextprogress where (np02,np03,np04,np05) in ("
'    strScanFagent = strScanFagent & " select pa01,pa02,pa03,pa04 from patent where pa75='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select tm01,tm02,tm03,tm04 from trademark where tm44='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select sp01,sp02,sp03,sp04 from servicepractice where sp26='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select lc01,lc02,lc03,lc04 from lawcase where lc22='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select cp01,cp02,cp03,cp04 from caseprogress where cp44='" & strFA01 & strFA02 & "')"
'    'Modify By Sindy 2009/07/24 增加LIN系統類別
'    '2010/3/23 MODIFY BY SONIA 剔除下一程序非智權人員掌控之案件性質改以strNpSqlOfNoSalesDuty控制
'    strScanFagent = strScanFagent & " and np06 is null and np08>=to_number(to_char(sysdate,'YYYYMMDD')) and np08<=to_number(to_char(add_months(sysdate,6),'YYYYMMDD')) " & strNpSqlOfNoSalesDuty
'    strScanFagent = strScanFagent & " )B,(select 'ok' as key,count(*) as CCount from (select pa01,pa02,pa03,pa04 from patent where pa75='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select tm01,tm02,tm03,tm04 from trademark where tm44='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select sp01,sp02,sp03,sp04 from servicepractice where sp26='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select lc01,lc02,lc03,lc04 from lawcase where lc22='" & strFA01 & strFA02 & "'"
'    strScanFagent = strScanFagent & " union select cp01,cp02,cp03,cp04 from caseprogress where cp44='" & strFA01 & strFA02 & "')CA"
'    strScanFagent = strScanFagent & " )C,(select 'ok' as key,nvl(na03,na04) as Fnation,fa01||fa02 as fnum,nvl(fa05,nvl(na04,na06)) as fname"
'    strScanFagent = strScanFagent & " from fagent,nation where fa01='" & strFA01 & "' and fa02='" & strFA02 & "' and fa10=na01(+)"
'    strScanFagent = strScanFagent & " )D where D.key=B.key(+) and D.key=C.key(+) group by fnation,fnum,fname,ccount"
'    CheckOC3
'    AdoRecordSet3.CursorLocation = adUseClient
'    AdoRecordSet3.Open strScanFagent, cnnConnection, adOpenStatic, adLockReadOnly
'    If AdoRecordSet3.RecordCount <> 0 Then
'        'add by nickc 2006/12/28 沒資料不印
'        'edit by nickc 2007/01/17 有請作單說有 6 個月案件才印
'        'If Val(CheckStr(AdoRecordSet3.Fields("C3"))) = 0 And Val(CheckStr(AdoRecordSet3.Fields("C6"))) = 0 And Val(CheckStr(AdoRecordSet3.Fields("Ccount"))) = 0 Then
'        If Val(CheckStr(AdoRecordSet3.Fields("C3"))) = 0 And Val(CheckStr(AdoRecordSet3.Fields("C6"))) = 0 Then
'            MsgBox "無六個月期限案件！", vbInformation, "不印明細表"
'        Else
'            Printer.Font.Size = "20"
'            Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth("修改代理人資料")) / 2
'            Printer.CurrentY = 0
'            Printer.Print "修改代理人資料"
'            Printer.Font.Size = 12
'            Printer.CurrentX = 0
'            Printer.CurrentY = 600
'            Printer.Print "修改人員：" & GetStaffName(strUserNum)
'            Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 300
'            Printer.CurrentY = 600
'            Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
'            Printer.CurrentX = 0
'            Printer.CurrentY = 900
'            Printer.Print String(150, "=")
'            Printer.CurrentX = 200
'            Printer.CurrentY = 1200
'            Printer.Print "國  籍"
'            Printer.CurrentX = 1500
'            Printer.CurrentY = 1200
'            Printer.Print "編   號"
'            Printer.CurrentX = 3000
'            Printer.CurrentY = 1200
'            Printer.Print "名         稱"
'            Printer.CurrentX = 6500
'            Printer.CurrentY = 1200
'            Printer.Print "3個月期限"
'            Printer.CurrentX = 7700
'            Printer.CurrentY = 1200
'            Printer.Print "6個月期限"
'            Printer.CurrentX = 8900
'            Printer.CurrentY = 1200
'            Printer.Print "所有案件數"
'            Printer.CurrentX = 0
'            Printer.CurrentY = 1500
'            Printer.Print String(150, "=")
'            Printer.CurrentX = 200
'            Printer.CurrentY = 1800
'            Printer.Print StrToStr(CheckStr(AdoRecordSet3.Fields("fnation")), 5)
'            Printer.CurrentX = 1500
'            Printer.CurrentY = 1800
'            Printer.Print StrToStr(CheckStr(AdoRecordSet3.Fields("fnum")), 5)
'            Printer.CurrentX = 3000
'            Printer.CurrentY = 1800
'            Printer.Print StrToStr(CheckStr(AdoRecordSet3.Fields("fname")), 10)
'            Printer.CurrentX = 6500 + Printer.TextWidth("3個月期限") - Printer.TextWidth(Format(Val(CheckStr(AdoRecordSet3.Fields("C3"))), "0"))
'            Printer.CurrentY = 1800
'            Printer.Print Format(Val(CheckStr(AdoRecordSet3.Fields("C3"))), "0")
'            Printer.CurrentX = 7700 + Printer.TextWidth("6個月期限") - Printer.TextWidth(Format(Val(CheckStr(AdoRecordSet3.Fields("C6"))), "0"))
'            Printer.CurrentY = 1800
'            Printer.Print Format(Val(CheckStr(AdoRecordSet3.Fields("C6"))), "0")
'            Printer.CurrentX = 8900 + Printer.TextWidth("所有案件數") - Printer.TextWidth(Format(Val(CheckStr(AdoRecordSet3.Fields("ccount"))), "0"))
'            Printer.CurrentY = 1800
'            Printer.Print CheckStr(AdoRecordSet3.Fields("ccount"))
'
'            strScanFagent = "select Cp01||'-'||cp02||'-'||cp03||'-'||cp04,cp01,cp02,cp03,cp04 from caseprogress where (cp01,cp02,cp03,cp04) in ("
'            strScanFagent = strScanFagent & " select pa01,pa02,pa03,pa04 from patent where pa75='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select tm01,tm02,tm03,tm04 from trademark where tm44='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select sp01,sp02,sp03,sp04 from servicepractice where sp26='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select lc01,lc02,lc03,lc04 from lawcase where lc22='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select cp01,cp02,cp03,cp04 from caseprogress where cp44='" & strFA01 & strFA02 & "')"
'            strScanFagent = strScanFagent & " and cp06>=to_number(to_char(sysdate,'YYYYMMDD')) and cp06<=to_number(to_char(add_months(sysdate,6),'YYYYMMDD')) and cp27 is null and cp57 is null"
'            strScanFagent = strScanFagent & " union select np02||'-'||np03||'-'||np04||'-'||np05,np02 as cp01,np03 as cp02,np04 as cp03,np05 as cp04 from nextprogress where (np02,np03,np04,np05) in ("
'            strScanFagent = strScanFagent & " select pa01,pa02,pa03,pa04 from patent where pa75='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select tm01,tm02,tm03,tm04 from trademark where tm44='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select sp01,sp02,sp03,sp04 from servicepractice where sp26='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select lc01,lc02,lc03,lc04 from lawcase where lc22='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select cp01,cp02,cp03,cp04 from caseprogress where cp44='" & strFA01 & strFA02 & "')"
'            'Modify By Sindy 2009/07/24 增加LIN系統類別
'            '2010/3/23 MODIFY BY SONIA 剔除下一程序非智權人員掌控之案件性質改以strNpSqlOfNoSalesDuty控制
'            strScanFagent = strScanFagent & " and np06 is null and np08>=to_number(to_char(sysdate,'YYYYMMDD')) and np08<=to_number(to_char(add_months(sysdate,6),'YYYYMMDD')) " & strNpSqlOfNoSalesDuty
'            strScanFagent = strScanFagent & " order by cp01,cp02,cp03,cp04  "
'            CheckOC3
'
'            AdoRecordSet3.CursorLocation = adUseClient
'            AdoRecordSet3.Open strScanFagent, cnnConnection, adOpenStatic, adLockReadOnly
'            If AdoRecordSet3.RecordCount <> 0 Then
'                Printer.CurrentX = 0
'                Printer.CurrentY = 2700
'                Printer.Print "六個月本所期限案件明細："
'                intLine = 3000
'                AdoRecordSet3.MoveFirst
'                nowCnt = 0
'                Seek01 = SystemNumber(CheckStr(AdoRecordSet3.Fields(0).Value), 1)
'                Do While Not AdoRecordSet3.EOF
'                    nowCnt = nowCnt + 1
'                    If nowCnt > intLineCnt Then
'                        nowCnt = 1
'                    End If
'                    If Seek01 <> SystemNumber(CheckStr(AdoRecordSet3.Fields(0).Value), 1) Then
'                        nowCnt = 1
'                        intLine = intLine + 600
'                        Seek01 = SystemNumber(CheckStr(AdoRecordSet3.Fields(0).Value), 1)
'                    End If
'                    Select Case nowCnt '(AdoRecordSet3.AbsolutePosition Mod intLineCnt)
'                    Case intLineCnt
'                         Printer.CurrentX = Printer.ScaleWidth - ((1 / intLineCnt) * Printer.ScaleWidth)
'                         Printer.CurrentY = intLine
'                         Printer.Print CheckStr(AdoRecordSet3.Fields(0).Value)
'                         intLine = intLine + 300
'                    Case Else
'                         Printer.CurrentX = Printer.ScaleWidth * (((nowCnt Mod intLineCnt) - 1) / intLineCnt)
'                         Printer.CurrentY = intLine
'                         Printer.Print CheckStr(AdoRecordSet3.Fields(0).Value)
'                    End Select
'                    AdoRecordSet3.MoveNext
'                Loop
''edit by nickc 2007/01/17 填請作單改有 6 個月期限案件才印
'            'add by nickc 2007/01/02
''            Else
''                Printer.CurrentX = 0
''                Printer.CurrentY = 2700
''                Printer.Print "無六個月本所期限案件!!"
'            End If
'        End If
'        Printer.EndDoc
'    End If
'end 2020/2/14
            
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
         Unload frm140103
         Set frm140103 = Nothing
      Case 2 '尋找
         CmdAddr.Enabled = False 'Added by Lydia 2018/10/24
         Label2(0) = "": Label2(1) = ""
         For i = 1 To 12
            Text1(i).Text = ""
         Next
         intI = 1
         '93.11.15 MODIFY BY SONIA
         'strExc(0) = "SELECT FA04,FA05,FA63,FA64,FA65,FA06 FROM FAGENT WHERE " & ChgFagent(Text1(0).Text)
         'Modify By Sindy 2011/1/18 +FA29
         strExc(0) = "SELECT FA04,FA05,FA63,FA64,FA65,FA06,FA24,FA29 FROM FAGENT WHERE " & ChgFagent(Text1(0).Text)
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
            m_FA24 = ""
            If Not IsNull(RsTemp.Fields(6).Value) Then m_FA24 = RsTemp.Fields(6).Value
            '93.11.15 END
            'Add By Sindy 2011/1/18
            m_FA29 = ""
            If Not IsNull(RsTemp.Fields("FA29").Value) Then m_FA29 = RsTemp.Fields("FA29").Value
            '2011/1/18 End
            CmdLock 0
            strTmp = Text1(0).Text & String(8 - Len(Text1(0).Text), "0")
            strExc(0) = "SELECT MAX(FA02) FROM FAGENT WHERE FA01='" & strTmp & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then Label2(0) = strTmp & Format(Val(RsTemp.Fields(0).Value) + 1): Label2(1) = "變更後代理人編號："
         Else
            MsgBox "代理人編號錯誤，請重新輸入 !", vbCritical
            TextInverse Text1(0)
            Text1(0).SetFocus
         End If
      Case 3 '取消
         If MsgBox("你並未存檔，確定離開嗎 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
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
    MsgBox "更新新名稱失敗，請洽系統管理員 !", vbCritical
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm140103 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   'edit by nickc 2007/06/06 切換輸入法改用API
   If Index = 7 Then OpenIme
   If Index = 12 Then OpenIme
End Sub

'Modified by Lydia 2021/09/23 改成Form 2.0
'Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   If Index = 0 Then KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CmdLock(TF As Integer)
   Select Case TF
      Case 0
         Command1(2).Enabled = False
         'Modified by Lydia 2018/10/24 依是否能修改權限以控制按鈕是否可作用
         'Command1(0).Enabled = True
         Command1(0).Enabled = m_bUpdate
         
         Command1(3).Enabled = True
         Text1(0).Locked = True
      Case 1
         Command1(2).Enabled = True
         Command1(0).Enabled = False
         Command1(3).Enabled = False
         Text1(0).Locked = False
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
'         'edit by nickc 2007/06/06 切換輸入法改用API
'         If Cancel = False Then CloseIme
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
'         'edit by nickc 2007/06/06 切換輸入法改用API
'         If Cancel = False Then CloseIme
'   End Select
   'end 2021/01/07
   If Cancel = True Then TextInverse Text1(Index)
End Sub

'Add By Cheng 2002/05/22
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

'Added by Lydia 2018/10/24 切換到代理人檔維護(確定變更名稱後,才可執行)
Private Sub CmdAddr_Click()
   Call frm050705.SetParent(Me, Text1(0).Text)
   frm050705.Show
   'Call Command1_Click(1) 'Mark by Lydia 2024/02/22
End Sub
