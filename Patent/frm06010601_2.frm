VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010601_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "實審通知日輸入"
   ClientHeight    =   4092
   ClientLeft      =   132
   ClientTop       =   936
   ClientWidth     =   7824
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4092
   ScaleWidth      =   7824
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   4
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   3
      Top             =   3660
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Index           =   5
      Left            =   5520
      MaxLength       =   7
      TabIndex        =   4
      Top             =   3660
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   960
      MaxLength       =   3
      TabIndex        =   13
      Top             =   570
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   12
      Top             =   570
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   11
      Top             =   570
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   10
      Top             =   570
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6864
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4824
      TabIndex        =   5
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5640
      TabIndex        =   6
      Top             =   72
      Width           =   1200
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm06010601_2.frx":0000
      Left            =   960
      List            =   "frm06010601_2.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   9
      Top             =   900
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   0
      Left            =   960
      MaxLength       =   50
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2700
      Visible         =   0   'False
      Width           =   5550
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   1
      Left            =   5520
      MaxLength       =   7
      TabIndex        =   0
      Top             =   3045
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   2
      Left            =   960
      MaxLength       =   20
      TabIndex        =   1
      Top             =   3345
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   3
      Left            =   5520
      MaxLength       =   7
      TabIndex        =   2
      Top             =   3345
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "主動修正本所期限:"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   36
      Top             =   3660
      Width           =   1485
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "主動修正法定期限:"
      Height          =   180
      Index           =   2
      Left            =   3945
      TabIndex        =   35
      Top             =   3660
      Width           =   1485
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   7700
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   7700
      Y1              =   2616
      Y2              =   2616
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   7
      Left            =   960
      TabIndex        =   34
      Top             =   2220
      Width           =   1950
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3440;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   6
      Left            =   5130
      TabIndex        =   33
      Top             =   1890
      Width           =   1950
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3440;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   32
      Top             =   2220
      Width           =   588
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Index           =   1
      Left            =   4272
      TabIndex        =   31
      Top             =   1890
      Width           =   768
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   30
      Top             =   900
      Width           =   768
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   4272
      TabIndex        =   29
      Top             =   570
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   28
      Top             =   570
      Width           =   768
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "專利種類:"
      Height          =   180
      Left            =   120
      TabIndex        =   27
      Top             =   1230
      Width           =   768
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "申請國家:"
      Height          =   180
      Index           =   0
      Left            =   4272
      TabIndex        =   26
      Top             =   1230
      Width           =   768
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "代理人:"
      Height          =   180
      Left            =   120
      TabIndex        =   25
      Top             =   1560
      Width           =   588
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   120
      TabIndex        =   24
      Top             =   1890
      Width           =   948
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   120
      TabIndex        =   23
      Top             =   2760
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "實審通知日期:"
      Height          =   180
      Left            =   4320
      TabIndex        =   22
      Top             =   3045
      Width           =   1125
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   21
      Top             =   3345
      Width           =   765
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "申請日:"
      Height          =   180
      Left            =   4320
      TabIndex        =   20
      Top             =   3345
      Width           =   585
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   0
      Left            =   5130
      TabIndex        =   19
      Top             =   570
      Width           =   1950
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3440;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   18
      Top             =   900
      Width           =   5790
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "10213;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   17
      Top             =   1230
      Width           =   1950
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3440;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   3
      Left            =   5130
      TabIndex        =   16
      Top             =   1230
      Width           =   1950
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3440;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   4
      Left            =   960
      TabIndex        =   15
      Top             =   1560
      Width           =   6570
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "11589;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   5
      Left            =   1110
      TabIndex        =   14
      Top             =   1890
      Width           =   1950
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3440;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm06010601_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/18 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String

Dim intWhere As Integer

Dim m_928Upd As Boolean   '是否更新重新委任准駁
Dim m_928CP09 As String   '重新委任收文號
'Public m_203check As Boolean  '第三人提實審是否管制三個月主動修正 Removed by Morgan 2014/10/30 (2013/1/11 取消)
'Modified by Lydia 2021/10/07 加以區隔；CP09=> mCP43 ;推測：使用者中途移到別的作業，影響到共用變數strExc(2)記錄的前畫面總收文號的案件性質
'Dim CP09 As String 'Add by Amy 2013/08/27 前畫面總收文號
Dim mCP43 As String
Dim mpCP10 As String 'Added by Lydia 2021/10/07 前畫面總收文號的案件性質
'Added by Morgan 2017/5/9 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
'end 2017/5/9
Dim m_431CP09 As String, m_422CP09 As String 'Added by Lydia 2020/11/27 431高速審查收文號、422加速審查收文號
Dim m_447CP09 As String 'Added by Morgan 2024/11/14 447再審查加速審查收文號
Dim stCP09 As String 'Modified by Morgan 2023/9/14 改全域變數

Private Sub cmdok_Click(Index As Integer)
Dim s_NIKON As Boolean      '2011/11/14 add by sonia
Dim bolFCMail As Boolean 'Added by Morgan 2023/9/19

   Select Case Index
      Case 0
         If Text5(2) = "" Then MsgBox "申請案號不可空白 !", vbCritical: Exit Sub
         'Modified by Lydia 2021/10/07 前畫面總收文號的案件性質
         'If strExc(3) <> 舉發 Then 'Added by Morgan 2017/9/1 Ex.FCP-56491
         If mpCP10 <> 舉發 Then
            If Text5(3) = "" Then MsgBox "申請日不可空白 !", vbCritical: Exit Sub
         End If
         'Added by Lydia 2021/10/07 增加檢查前畫面總收文號; ex.FCP-059502(110/8/12)的CB0049624的相關總收文號存到A5023
         If Len(mCP43) < 9 Then
            MsgBox "相關總收文號不符，請重新選擇！", vbCritical
            Exit Sub
         End If
         'end 2021/10/07
         
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         'Add by Sindy 2021/11/22 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If
         
         'Added by Morgan 2023/9/14--Sharon
         bolFCMail = False
         '112/9/1 起台灣設計專利新增「加速審查」制度
         'Modified by Morgan 2023/9/21 +125衍生設計--Anny
         'Modified by Morgan 2023/10/17 +307,303,308
         If pa(8) = "3" And (mpCP10 = "103" Or mpCP10 = "125" Or mpCP10 = "307" Or mpCP10 = "303" Or mpCP10 = "308") Then
            strExc(0) = "select * from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in ('245','422') and cp57 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               bolFCMail = True
               'Added by Morgan 2023/10/17
               If mpCP10 = "307" Or mpCP10 = "303" Or mpCP10 = "308" Then
                  If pa(163) = "Y" Then
                     bolFCMail = True
                  ElseIf pa(163) = "N" Then
                     bolFCMail = False
                  Else
                     MsgBox "尚未設定是否為初審階段提" & Label2(6) & "!!" & vbCrLf & "請先確認並於基本檔設定後再輸入本來函!!", vbExclamation
                     Exit Sub
                  End If
               End If
               'end 2023/10/17
               
               If bolFCMail = True Then
                  MsgBox "設計案通知實審來函請通知客戶！", vbInformation
                  'Added by Morgan 2023/9/19
                  If PUB_CheckFormExist("frm090401") Then
                     MsgBox "系統將產生FC郵件，請先關閉【撰寫信函】畫面！", vbExclamation
                     Exit Sub
                  End If
                  'end 2023/9/19
                  
                  'Added by Morgan 2024/11/15
                  '若為初審階段提分割且卷宗區無寄請款函(REPDN)或請款單上傳(DNUPL)時彈訊息(只管制分割,因為是程序請款,承辦請款的狀況比較複雜先不考慮)--敏莉
                  'Modified by Morgan 2025/1/14 +308 --敏莉
                  If mpCP10 = "307" Or mpCP10 = "308" Then
                     'Modified by Lydia 2025/01/16 改成模組取得語法;AND (UPPER(CPP02) LIKE '%.REPDN.%' OR UPPER(CPP02) LIKE '%.DNUPL.%' ) >> PUB_GetFCPforDNsql
                     strExc(0) = "select * from casepaperpdf where cpp01='" & mCP43 & "' AND NVL(CPP10,'N') <> 'D' " & PUB_GetFCPforDNsql
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 0 Then
                        'Modified by Morgan 2025/1/14 改訊息 --敏莉
                        'MsgBox "本案分割尚未寄出請款信，通知實審暫緩寄出，待完成請款流程後再寄出。", vbExclamation
                        MsgBox "本案尚有請款流程尚未完成，通知實審暫緩寄出，待完成請款流程後再寄出。", vbExclamation
                     End If
                  End If
                  'end 2024/11/15
                  
               End If
            End If
         End If
         'end 2023/9/14
   
         'Added by Morgan 2023/12/28
         '發明申請(101),設計申請(103),衍生設計申請(125)的通知實審檢查是否有補文件未收文或未發文 --敏莉
         If (mpCP10 = "101" Or mpCP10 = "103" Or mpCP10 = "125") Then
            PUB_ChkUnAddDoc mCP43
         End If
         'end 2023/12/28
         
         Screen.MousePointer = vbHourglass
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
   
         'Add by Morgan 2010/3/29
         '輸通知實審日時若有加速審查未發文則提醒
         'Modified by Morgan 2012/5/30 +431高速審查
         'Modified by Lydia 2020/11/27 +cp14
         'Modified by Morgan 2024/11/14 +447再審查加速審查
         strExc(0) = "select cp10, cp14, cpm03 from caseprogress,casepropertymap where cp01='" & pa(1) & "'" & _
            " and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
            " and cp10 in ('422','431','447') and cp57 is null and cp27 is null and cpm01(+)=cp01 and cpm02(+)=cp10"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Modifedi by Lydia 2020/11/27 自動發Outlook給承辦工程師ＣＣ工程師主管及程序人員、backup，內容如下
            'MsgBox "本案" & RsTemp("cpm03") & "尚未發文！", vbInformation
            RsTemp.MoveFirst
            strExc(3) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
            Do While Not RsTemp.EOF
                strExc(0) = "": strExc(1) = ""
                strExc(2) = ""
                '主旨
                'Modified by Lydia 2020/12/02 + [INCOM.1204]
                strExc(0) = "Our Ref: " & pa(1) & "-" & pa(2) & IIf(pa(3) <> "0", "-" & pa(3), "") & IIf(pa(4) <> "00", "-" & pa(4), "") & " [INCOM.1204] 已收到即將進行實體審查函，請進行" & RsTemp.Fields("CPM03") & "送件作業"
                'Add by Amy 2025/08/05 後續准駁簡單報告=Y,輸C類來函[主旨]最前面加【請簡單報告】-Winfrey
                If pa(89) = "Y" Then strExc(0) = " 【請簡單報告】" & strExc(0)
                
                '收件人
                strExc(1) = IIf("" & RsTemp.Fields("cp14") <> "", "" & RsTemp.Fields("cp14"), strExc(3))
                '副本=ＣＣ工程師主管及程序人員、backup
                'Modified by Morgan 2024/12/19 不CC的只有管制人=承辦人，其他還是要CC
                'If strExc(1) <> strExc(3) Then
                '    strExc(2) = PUB_GetFCPEngSup(strExc(1))
                '    strExc(2) = strExc(2) & ";" & strExc(3) & ";backup"
                'End If
                strExc(2) = PUB_GetFCPEngSup(strExc(1))
                If strExc(1) <> strExc(3) Then
                  strExc(2) = strExc(2) & ";" & strExc(3)
                End If
                strExc(2) = strExc(2) & ";backup"
                'end 2024/12/19
                
                PUB_SendMail strUserNum, strExc(1), "", strExc(0), vbCrLf & "同主旨", , , , , , strExc(2)
                
                'MsgBox "本案尚有" & RsTemp("cpm03") & "未發文，請將來函會給工程師！", vbInformation 'Removed by Morgan 2024/11/14 不用了--敏莉
                
                RsTemp.MoveNext
            Loop
            'end 2020/11/27
         End If
         
         'Add by Morgan 2008/8/21
         If Left(pa(75), 6) = "Y20304" Then
            MsgBox "本案為 ASAHI 的案件，來函要寄客戶！", vbInformation
         End If
         'END 2008/8/21
         
         'Added by Morgan 2012/11/5
         'modify by sonia 2013/6/10 加入Y20624
         If Left(pa(75), 6) = "Y53309" Or Left(pa(75), 6) = "Y20624" Then
            MsgBox "本案需調卷轉承辦組報告並寄代！", vbInformation
         End If
         'end 2012/11/5
         
         '2011/11/14 add by sonia
         s_NIKON = False
         'Modified by Morgan 2020/2/7 +X60049C10--Landy
         Select Case Left(pa(26), 8)
            Case "X45149", "X56040", "X53310", "X48340", "X47956", "X45148", "X60049C1"
               s_NIKON = True
         End Select
         Select Case Left(pa(27), 8)
            Case "X45149", "X56040", "X53310", "X48340", "X47956", "X45148", "X60049C1"
               s_NIKON = True
         End Select
         Select Case Left(pa(28), 8)
            Case "X45149", "X56040", "X53310", "X48340", "X47956", "X45148", "X60049C1"
               s_NIKON = True
         End Select
         Select Case Left(pa(29), 8)
            Case "X45149", "X56040", "X53310", "X48340", "X47956", "X45148", "X60049C1"
               s_NIKON = True
         End Select
         Select Case Left(pa(30), 8)
            Case "X45149", "X56040", "X53310", "X48340", "X47956", "X45148", "X60049C1"
               s_NIKON = True
         End Select
         'add by sonia 2016/6/17 Y51333010+X74001(北京銀龍+吉佳藍科技) 或 Y34232+X48637 (YASUTOMI+大日本印刷)也要提醒
         If pa(75) = "Y5133301" And (Left(pa(26), 8) = "X74001" Or Left(pa(27), 8) = "X74001" Or Left(pa(28), 8) = "X74001" Or Left(pa(29), 8) = "X74001" Or Left(pa(30), 8) = "X74001") Then
            s_NIKON = True
         End If
         'modify by sonia 2017/8/14  再加Y54116+X48637 也要提醒
         'modify by sonia 2017/10/16 再加Y47649+X48637,Y48651+X48637 也要提醒
         'modify by sonia 2017/12/19 取消Y48651+X48637
         'Modified by Morgan 2018/10/12 +Y55102+X48637 --洪郁嵐
         If (pa(75) = "Y34232" Or pa(75) = "Y54116" Or pa(75) = "Y47649" Or pa(75) = "Y55102") And (Left(pa(26), 8) = "X48637" Or Left(pa(27), 8) = "X48637" Or Left(pa(28), 8) = "X48637" Or Left(pa(29), 8) = "X48637" Or Left(pa(30), 8) = "X48637") Then
            s_NIKON = True
         End If
         'end 2016/6/17
         
         'Added by Morgan 2020/4/10 --Jessica,Lisa
         'Modified by Morgan 2020/6/23 +Y5133301 --Jessica,Kimi
         If pa(75) = "Y54339" Or pa(75) = "Y54339B1" Or pa(75) = "Y54339B2" Or pa(75) = "Y5133301" Then
            s_NIKON = True
         End If
         'end 2020/4/10
         
         'Added by Morgan 2017/12/4 FCP57047 收到智慧局未出接洽單之來函: 歸卷、一般來函(延期受理、通知補文件)、通知實審日、通知公開 提醒 "收到智慧局來函需退承辦報告客戶" -- Sharon
         If pa(1) & pa(2) = "FCP057047" Then
            MsgBox "本案收到智慧局來函需退承辦報告客戶！", vbInformation
         
         'Added by Morgan 2020/1/20--Landy
         ElseIf pa(1) & pa(2) = "FCP062271" Then
            MsgBox "請退工程師報告客戶！", vbInformation
         'end 2020/1/20
         
         'add by sonia 2021/3/4--林芳如
         ElseIf pa(1) & pa(2) = "FCP064092" Then
            MsgBox "請會承辦：通知客戶PPH相關事！", vbInformation
         'end 2021/3/4
         
         'add by sonia 2021/11/8--蘇暐嵐
         ElseIf pa(1) & pa(2) = "FCP060205" Then
            MsgBox "請通知承辦報告客戶！", vbInformation
         'end 2021/3/4
         
         'Added by Morgan 2018/8/14 FCP59431 --Sharon
         'Removed by Morgan 2018/8/31 取消 --Sharon
         'ElseIf pa(1) & pa(2) = "FCP059431" Then
         '   MsgBox "客戶預計提AEP,需調卷退承辦報告！", vbInformation
         'end 2018/8/31
         
         'Added by Morgan 2022/4/12 +Y55105 --蘇暐嵐
         ElseIf pa(75) = "Y55105" And pa(57) = "" Then
            MsgBox "收到智慧局來函,請通知承辦寄代！", vbInformation
         
         'Modified by Morgan 2019/7/24 +FCP-055144--敏莉
         ElseIf s_NIKON = True Or pa(1) & pa(2) = "FCP055144" Then
         
            MsgBox "請調卷退承辦智權同仁通知即將進入實審！", vbInformation  'modify by sonia 2016/6/17 加入'調卷'二字
         End If
         '2011/11/14 end
         
         If bolFCMail Then FCMail 'Added by Morgan 2023/9/14
         
         Screen.MousePointer = vbDefault
         
         Unload Me
         Unload frm06010601_1
         
         'Added by Morgan 2017/5/9 電子公文
         If m_DocNo <> "" Then
            Unload frm06010601
            frm060119.GoNext
         Else
         'end 2017/5/9
            frm06010601.Show
            frm06010601.Clear
         End If 'Added by Morgan 2017/5/9
         
      Case 1
         frm06010601_1.Show
         Unload Me
      Case 2
         Unload frm06010601
         Unload frm06010601_1
         Unload Me
   End Select
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label2(1) = pa(5)
      Case "英"
         Label2(1) = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         Label2(1) = pa(7)
   End Select
End Sub

Private Sub Form_Activate()
   If TransDate(Label2(5), 1) = "111111" Then
      MsgBox "實審通知日期請輸入實審提出日期或系統日！", vbInformation
   End If

End Sub

'add by nickc 2007/02/09
Private Sub Form_Initialize()
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()
Dim strTemp As String, strTemp1 As String

   MoveFormToCenter Me
   intWhere = 國內
   With frm06010601_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      'Modified by Lydia 2021/10/07
      'CP09 = strExc(2) 'Add by Amy 2013/08/27 前畫面總收文號
      mCP43 = strExc(2)
      '2007/8/6 ADD BY SONIA
      mpCP10 = strExc(3) 'Added by Lydia 2021/10/07 前畫面總收文號的案件性質
      
      'Removed by Morgan 2014/10/30 (2013/1/11 取消)
      'If m_203check Then
      '   Label5(1).Visible = True
      '   Label5(2).Visible = True
      '   Text5(4).Visible = True
      '   Text5(4).Enabled = True
      '   Text5(5).Visible = True
      '   Text5(5).Enabled = True
      '   strTemp = CompDate(1, 3, TransDate(frm06010601_1.Label2(5), 2))
      '   strTemp1 = CompDate(2, -7, strTemp)
      '   Text5(4) = ChangeWStringToTString(strTemp1)
      '   Text5(5) = ChangeWStringToTString(strTemp)
      'Else
         Label5(1).Visible = False
         Label5(2).Visible = False
         Text5(4).Visible = False
         Text5(4).Enabled = False
         Text5(5).Visible = False
         Text5(5).Enabled = False
      'End If
      '2007/8/6 END
      'end 2014/10/30
      
   End With
   If strExc(5) = "1" Then
      Text5(2).Enabled = False
      Text5(3).Enabled = False
   End If
'   strExc(2) = .TextMatrix(i, 1) '收文號
'   strExc(3) = .TextMatrix(i, 5) '案件性質代號
'   strExc(4) = .TextMatrix(i, 6) '業務區別
'   strExc(5) = .TextMatrix(i, 7) '智權人員代號
'   strExc(6) = .TextMatrix(i, 2) '案件性質
'   strExc(7) = .TextMatrix(i, 3) '發文日
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Morgan 2021/6/11
   Set frm06010601_2 = Nothing
End Sub

Public Sub QueryData()
   ReadPatent
   'Added by Morgan 2017/5/9 電子公文
   If m_DocWord <> "" Then
      Text5(0) = m_DocWord & "字第" & m_DocNo & "號"
   End If
   'end 2017/5/9
   Combo1.ListIndex = 0

End Sub

'************************************************
' 取回專利基本資料及收文資料
'
'************************************************
Private Sub ReadPatent()
 Dim Lbl As Object, i As Integer, strTempName As String
   For Each Lbl In Label2
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      Label2(1) = pa(5)
      If pa(8) <> "" Then ChgType (2) ' Label2(2)
      If pa(9) <> "" Then ChgType (3) ' Label2(3)
      If pa(75) <> "" Then ChgType (4) ' Label2(4)
      Label2(0) = pa(11)
      Text5(2) = pa(11)
      Text5(3) = pa(10)
   End If
   Label2(5) = frm06010601_1.Label2(5)
   Text5(1) = Label2(5)          'add by sonia 2016/3/24
   Label2(6) = strExc(6)
   Label2(7) = strExc(7)
   '94/2/22 add by sonia
   If pa(10) <> "" Then
      'Modified by Lydia 2021/10/07 前畫面總收文號的案件性質
      'If pa(23) = "1" And (strExc(3) = "102" Or strExc(3) = "302") Then
      If pa(23) = "1" And (mpCP10 = "102" Or mpCP10 = "302") Then
         Label17 = "文件齊備日期:"
      End If
   End If
   '94/2/22 end
   
End Sub

Private Function ChgType(i As Integer) As Boolean
 Dim strTempName As String
   ChgType = False
   Select Case i
      Case 2
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetPatentTrademarkKind(專利, pA(8), strTempName, False, 台灣國家代號) = 1 Then
         If ClsPDGetPatentTrademarkKind(專利, pa(8), strTempName, False, 台灣國家代號) = 1 Then
            Label2(2) = strTempName
         End If
      Case 4
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetAgent(pA(75), strTempName) Then
         If ClsPDGetAgent(pa(75), strTempName) Then
            Label2(4) = strTempName
            ChgType = True
         End If
      Case 3
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNation(pA(9), strTempName) = True Then
         If ClsPDGetNation(pa(9), strTempName) = True Then
            Label2(3) = strTempName
            ChgType = True
         End If
   End Select
End Function

' 儲存資料表
Private Function FormSave() As Boolean

Dim i As Integer, bolChk As Boolean, strTxt(1 To 5) As String
Dim strCP12 As String, strCP13 As String
Dim strCP20 As String, strCP16 As String
Dim stNP09 As String, stNP22 As String
Dim stCP10 As String 'Added by Morgan 2017/5/9

FormSave = True

m_928Upd = PUB_928Check(pa, m_928CP09) 'Add by Morgan 2007/7/17

On Error GoTo ErrHnd

cnnConnection.BeginTrans

   strCP13 = PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))
   strCP12 = GetSalesArea(strCP13)
   
   'Add by Morgan 2007/7/17
   If m_928Upd = True And m_928CP09 <> "" Then
      PUB_928Update pa, m_928CP09, , strCP13, strCP12
   End If
   'end 2007/7/17
   stCP09 = AutoNo("C", 6)
   '94.2.2 MODIFY BY SONIA 不考慮申請日
   'If pa(8) = "2" And pa(23) = "1" And pa(10) >= 920701 Then
   '94.2.22 MODIFY BY SONIA 不判斷專利種類改判斷案件性質為102及302者
   'Modified by Lydia 2021/10/07 前畫面總收文號的案件性質
   'If (strExc(3) = "102" Or strExc(3) = "302") And pa(23) = "1" Then
   If (mpCP10 = "102" Or mpCP10 = "302") And pa(23) = "1" Then
      stCP10 = "1217" 'Added by Morgan 2017/5/9
      'Modified by Lydia 2021/10/07 前畫面總收文號strExc(2)=>mCP43; Ex. FCP-059502(110/8/12)的CB0049624的相關總收文號存到A5023
      strTxt(1) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10," & _
         "CP12,CP13,CP14,CP20,CP26,CP32,CP27,CP43) VALUES ('" & Text1 & "','" & Text2 & "','" & _
         Text3 & "','" & Text4 & "'," & TransDate(Text5(1), 2) & ",'" & _
         stCP09 & "','1217','" & strCP12 & "','" & _
         strCP13 & "','" & strUserNum & "','N','N','N'," & strSrvDate(1) & ",'" & _
         mCP43 & "')"
   Else
      stCP10 = 通知實審日 'Added by Morgan 2017/5/9
      'Add by Morgan 2005/1/11
      '案件性質為發明申請’101’或改請發明’301’下一程序檔新增催審期限，本所期限＝法定期限＝來函收文日＋2年
      '2006/5/16 MODIFY BY SONIA 加發明之分割
      'If strExc(3) = 發明申請 Or strExc(3) = 改請發明 Then
      'Modified by Lydia 2021/10/07 前畫面總收文號的案件性質
      'If strExc(3) = 發明申請 Or strExc(3) = 改請發明 Or (strExc(3) = 分割 And pa(8) = "1") Then
      If mpCP10 = 發明申請 Or mpCP10 = 改請發明 Or (mpCP10 = 分割 And pa(8) = "1") Then
      '2006/5/16 END
         '2009/2/19 modify by sonia 靜芳說改為實審通知日期+3年(同時改資料庫資料)
         'stNP09 = CompDate(0, 2, TransDate(Label2(5), 2))
         stNP09 = CompDate(0, 3, TransDate(Text5(1), 2))
         'add by sonia 2019/11/27 FCP-061886
         If Text5(1) = 111111 Then
            stNP09 = CompDate(0, 3, strSrvDate(1))
         End If
         'end 2019/11/27
         '2009/2/19 end
         stNP22 = GetNextProgressNo  'edit by nickc 2007/02/02 不用 dll 了  objPublicData.GetNextProgressNo
'Modify by Morgan 2005/2/21 改用申請案的總收文號,這樣核准時才會上自動上'Y'--靜芳
'         StrSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,NP22) " & _
'            "VALUES ('" & stCP09 & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
'            "','" & pa(4) & "','" & strUserNum & "'," & 催審 & "," & stNP09 & "," & _
'           stNP09 & "," & stNP22 & ")"
         'Modified by Lydia 2021/10/07 前畫面總收文號strExc(2)=>mCP43; Ex. FCP-059502(110/8/12)的CB0049624的相關總收文號存到A5023
         'Modified by Lydia 2025/11/12 改抓最近工作天+PUB_GetWorkDay1
         strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,NP22) " & _
            "VALUES ('" & mCP43 & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
            "','" & pa(4) & "','" & strUserNum & "'," & 催審 & "," & PUB_GetWorkDay1(stNP09, True) & "," & _
           stNP09 & "," & stNP22 & ")"
'2005/2/21 end
         cnnConnection.Execute strSql
         
         
         'Removed by Morgan 2014/10/30 (2013/1/11 取消)
         ''2007/8/6 ADD BY SONIA 掛主動修正期限
         'If m_203check = True Then
         '   stNP22 = GetNextProgressNo
         '   strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,NP22) " & _
         '      "VALUES ('" & stCP09 & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
         '      "','" & pa(4) & "'," & CNULL(PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))) & ",'203'," & TransDate(Text5(4), 2) & "," & _
         '     TransDate(Text5(5), 2) & "," & stNP22 & ")"
         '   cnnConnection.Execute strSql
         'End If
         ''2007/8/6 END
         'end 2014/10/30
         'Modified by Lydia 2021/10/07 前畫面總收文號strExc(2)=>mCP43; Ex. FCP-059502(110/8/12)的CB0049624的相關總收文號存到A5023
         strTxt(1) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10," & _
            "CP12,CP13,CP14,CP20,CP26,CP32,CP27,CP43) VALUES ('" & Text1 & "','" & Text2 & "','" & _
            Text3 & "','" & Text4 & "'," & TransDate(Text5(1), 2) & ",'" & _
            stCP09 & "','" & 通知實審日 & "','" & strCP12 & "','" & _
            strCP13 & "','" & strUserNum & "','N','N','N'," & strSrvDate(1) & ",'" & _
            mCP43 & "')"
      Else
      '2005/1/11 end
         'Modified by Lydia 2021/10/07 前畫面總收文號strExc(2)=>mCP43; Ex. FCP-059502(110/8/12)的CB0049624的相關總收文號存到A5023
         strTxt(1) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10," & _
            "CP12,CP13,CP14,CP20,CP26,CP32,CP27,CP43) VALUES ('" & Text1 & "','" & Text2 & "','" & _
            Text3 & "','" & Text4 & "'," & TransDate(Text5(1), 2) & ",'" & _
            stCP09 & "','" & 通知實審日 & "','" & strCP12 & "','" & _
            strCP13 & "','" & strUserNum & "','N','N','N'," & strSrvDate(1) & ",'" & _
            mCP43 & "')"
            
      End If
   End If
   '92.2.17 end
   '911105 nick transation
   cnnConnection.Execute strTxt(1)
   
   'Add by Amy 2013/08/27 判斷收文號為B類更新cp05及cp27=19221111
   'Modified by Morgan 2019/10/31 電子公文除外 Ex:FCP-061493
   'Modified by Lydia 2021/10/07 前畫面總收文號strExc(2)=>mCP43; Ex. FCP-059502(110/8/12)的CB0049624的相關總收文號存到A5023
   'If Left(CP09, 1) = "B" And pa(8) = 1 And m_DocNo = "" Then
   If Left(mCP43, 1) = "B" And pa(8) = 1 And m_DocNo = "" Then
        'Modified by Morgan 2014/9/4
        'strTxt(1) = "Update CaseProgress Set cp05=19221111,cp27=19271111 Where cp09='" & CP09 & "' "
        strTxt(1) = "Update CaseProgress Set cp05=19221111,cp27=19221111 Where cp09='" & stCP09 & "' "
        cnnConnection.Execute strTxt(1)
   End If
   'end 2013/08/27
   
   'Modified by Lydia 2024/05/28 改成模組
   ''Added by Lydia 2022/05/03 FCP-062174審定前不收費控制: (補上)判斷基本檔之目前准/駁PA16為空值時，不管任何案件性質都不必預設收文費用、規費、點數。
   'If pa(16) = "" And InStr("FCP062174000", pa(1) & pa(2) & pa(3) & pa(4)) > 0 Then
   '      strSql = "update caseprogress set cp20='N', cp16=null, cp17=null, cp18=null where cp09='" & stCP09 & "'"
   '      cnnConnection.Execute strSql
   ''FCP-067004核准前不收費控制：申請至核准(暫不包含領證)不收任何收費 (包含規費及服務費、若客戶提AEP也不收費)
   'ElseIf pa(16) <> "1" And InStr("FCP067004000", pa(1) & pa(2) & pa(3) & pa(4)) > 0 Then
   If PUB_GetCP20forSpec(pa(1), pa(2), pa(3), pa(4), pa(16)) = "N" Then
   'end 2024/05/28
         strSql = "update caseprogress set cp20='N', cp16=null, cp17=null, cp18=null where cp09='" & stCP09 & "'"
         cnnConnection.Execute strSql
   Else
   'end 2022/05/03
      'Add by Morgan 2007/7/23 CP20改抓CPM的設定
      'Modify by Morgan 2008/3/27 +pa75
      'Modify by Morgan 2008/4/10 +本所案號
      strCP20 = PUB_GetCP20(Text1, 通知實審日, strCP16, pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4))
      If strCP20 = "" Then
         strSql = "update caseprogress set cp20=NULL,cp16=" & strCP16 & ",cp17=0,cp18=" & strCP16 / 1000 & _
            " where cp09='" & stCP09 & "'"
         cnnConnection.Execute strSql
      End If
   'end 2007/7/23
   End If 'Added by Lydia 2022/05/03
   'Modified by Lydia 2021/10/07 前畫面總收文號strExc(2)=>mCP43; Ex. FCP-059502(110/8/12)的CB0049624的相關總收文號存到A5023
   strTxt(2) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & mCP43 & "' AND NP07='" & 通知實審日 & "'"
   
   '911105 nick transation
   cnnConnection.Execute strTxt(2)
   
   i = 3
'2015/1/30 cancel by sonia 前申請案號輸入已更新,此處再更新則會覆蓋掉原申請案號
'   If Left(strExc(3), 1) = "3" And frm06010601.Tag = "1" Then
'      strTxt(3) = "UPDATE CASEPROGRESS SET CP30='" & pa(11) & "' WHERE CP09='" & strExc(2) & "'"
'
'        '911105 nick transation
'        cnnConnection.Execute strTxt(3)
'
'      i = 4
'   End If
'2015/1/30 end

   strTxt(i) = "UPDATE PATENT SET PA10=" & CNULL(TransDate(Text5(3), 2), True) & ",PA11='" & Text5(2) & "' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   '911105 nick transation
   'FormSave = objLawDll.ExecSQL(i, strTxt)
   cnnConnection.Execute strTxt(i)
   
   'Added by Morgan 2012/7/13
   '若有加速審查未發文則更新期限
   'Modified by Lydia 2020/11/27 記錄收文號,並且以來函收文日計算承辦期
   'If PUB_ChkCPExist(pa(), "422", 1, strExc(1)) = True Then
      'strExc(2) = Pub_GetHandleDay("FCP", "000", "422")
   If PUB_ChkCPExist(pa(), "422", 1, m_422CP09) = True Then 'AEP
      strExc(2) = Pub_GetHandleDay("FCP", "000", "422", IIf(TransDate(Label2(5), 1) = "111111", strSrvDate(1), TransDate(Label2(5), 2)))
   'end 2020/11/27
      If Val(strExc(2)) > 0 Then
         'Modified by Lydia 2020/11/27 本所期限(承辦期限+5個工作天--含CP48當天); 承辦期限一律要更新by Sharon, Phoebe
         'strSql = "update caseprogress set cp48=" & strExc(2) & " where cp09='" & strExc(1) & "' and cp48 is null"
         strSql = "update caseprogress set cp48=" & strExc(2) & ", cp06=" & CompWorkDay(5, strExc(2)) & " where cp09='" & m_422CP09 & "' "
         cnnConnection.Execute strSql, intI
         'Added by Lydia 2020/12/18 若有主動修正未發文，則一併回寫上述之期限至主動修正--12/11 email
         'Modified by Lydia 2021/03/18 更新本所期限cp06
         'strSql = "update caseprogress set cp48=" & strExc(2) & " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp01='" & pa(4) & "' and cp10='203' and cp158=0 and cp159=0 "
          'Modified by Lydia 2021/04/28 debug and cp01='" & pa(4) & "'=> and cp04='" & pa(4) & "'
         strSql = "update caseprogress set cp48=" & strExc(2) & ", cp06=" & CompWorkDay(5, strExc(2)) & " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='203' and cp158=0 and cp159=0 "
         cnnConnection.Execute strSql, intI
         'end 2020/12/18
      End If
      
   'Added by Morgan 2024/11/14 447再審查加速審查比照422--敏莉
   ElseIf PUB_ChkCPExist(pa(), "447", 1, m_447CP09) = True Then 'AEPRe
      strExc(2) = Pub_GetHandleDay("FCP", "000", "447", IIf(TransDate(Label2(5), 1) = "111111", strSrvDate(1), TransDate(Label2(5), 2)))
      If Val(strExc(2)) > 0 Then
         strSql = "update caseprogress set cp48=" & strExc(2) & ", cp06=" & CompWorkDay(5, strExc(2)) & " where cp09='" & m_447CP09 & "' "
         cnnConnection.Execute strSql, intI
         strSql = "update caseprogress set cp48=" & strExc(2) & ", cp06=" & CompWorkDay(5, strExc(2)) & " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='203' and cp158=0 and cp159=0 "
         cnnConnection.Execute strSql, intI
      End If
   'end 2024/11/14
   
   End If
   
   
   
   'Added by Lydia 2020/11/27 若有高速審查未發文則更新期限
   If PUB_ChkCPExist(pa(), "431", 1, m_431CP09) = True Then 'PPH
      strExc(2) = Pub_GetHandleDay("FCP", "000", "431", IIf(TransDate(Label2(5), 1) = "111111", strSrvDate(1), TransDate(Label2(5), 2)))
      If Val(strExc(2)) > 0 Then
         strSql = "update caseprogress set cp48=" & strExc(2) & ", cp06=" & CompWorkDay(5, strExc(2)) & " where cp09='" & m_431CP09 & "' "
         cnnConnection.Execute strSql, intI
         'Added by Lydia 2020/12/18 若有主動修正未發文，則一併回寫上述之期限至主動修正--12/11 email
         'Modified by Lydia 2021/03/18 更新本所期限cp06
         'strSql = "update caseprogress set cp48=" & strExc(2) & " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp01='" & pa(4) & "' and cp10='203' and cp158=0 and cp159=0 "
         'Modified by Lydia 2021/04/28 debug and cp01='" & pa(4) & "'=> and cp04='" & pa(4) & "'
         strSql = "update caseprogress set cp48=" & strExc(2) & ", cp06=" & CompWorkDay(5, strExc(2)) & " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='203' and cp158=0 and cp159=0 "
         cnnConnection.Execute strSql, intI
         'end 2020/12/18
      End If
   End If
   'end 2020/11/27
   
   'Added by Morgan 2017/5/9 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, stCP09, pa(1), pa(2), pa(3), pa(4), stCP10
   'Added by Morgan 2021/6/11 紙本公文--何淑華
   Else
      PUB_FCPOAInform stCP09, pa(1), pa(2), pa(3), pa(4), stCP10
   End If
   'end 2017/5/9
   
   cnnConnection.CommitTrans
   Exit Function
ErrHnd:
   cnnConnection.RollbackTrans
   FormSave = False
   '911105 nick transation
   
End Function

Private Sub Text5_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      'Add by Morgan 2011/1/5
      Case 0
            If CheckLengthIsOK(Text5(Index), Text5(Index).MaxLength) = False Then
               Cancel = True
            End If
      Case 1, 3
         If Text5(Index) <> "" Then
            If ChkDate(Text5(Index)) Then
               If Val(Text5(Index)) > Val(strSrvDate(2)) Then
                  MsgBox "日期不可大於系統日 !", vbCritical
                  Cancel = True
               End If
            Else
               Cancel = True
            End If
         'Modified by Morgan 2017/9/1 申請日存檔有控制
         'Else
         ElseIf Index = 1 Then
         'end 2017/9/1
            MsgBox "日期不可空白 !", vbCritical
            Cancel = True
         End If
      Case 2
         If Text5(Index) = "" Then
            MsgBox "申請案號不可空白 !", vbCritical
            Cancel = True
         Else
            '2005/6/14 MODIFY BY SONIA
            'If Not ChkAppNo(Text5(Index).Text, pa(8), 0) Then
            If Not ChkAppNo(Text5(Index).Text, pa(8), 0, Val(pa(23))) Then
            '2005/6/14 END
               Cancel = True
            End If
         End If
      '2007/8/6 ADD BY SONIA
      Case 4, 5
         If Text5(Index).Enabled = True Then
            If Text5(Index) <> "" Then
               If ChkDate(Text5(Index)) Then
                  If Val(Text5(Index)) <= Val(strSrvDate(2)) Then
                     MsgBox "主動修正期限必須大於系統日 !", vbCritical
                     Cancel = True
                  End If
               Else
                  Cancel = True
               End If
            Else
               MsgBox "主動修正期限不可空白 !", vbCritical
               Cancel = True
            End If
         End If
      '2007/8/6 END
   End Select
   If Cancel = True Then TextInverse Text5(Index)
End Sub

Private Sub Text5_GotFocus(Index As Integer)
   InverseTextBox Text5(Index)
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
For Each objTxt In Text5
   If objTxt.Enabled = True Then
      Cancel = False
      Text5_Validate objTxt.Index, Cancel
      If Cancel = True Then
         objTxt.SetFocus
         Exit Function
      End If
   End If
Next

TxtValidate = True
End Function

'Added by Morgan 2023/9/14
'產生FC郵件
Private Sub FCMail()
   Dim strTempFolder As String, strTmpFileName As String
   Dim strFileName As String, strFullFileName As String
   Dim oFileSys As New FileSystemObject
   Dim oFile As File
   Dim strMsg As String, strTmp As String
   Dim strAtt As String
   
On Error GoTo ErrHnd

   strTempFolder = App.path & "\$$LetterTemp"
   If Dir(strTempFolder, vbDirectory) = "" Then
      MkDir strTempFolder
   Else
      If Dir(strTempFolder & "\.") <> "" Then
         Kill strTempFolder & "\*.*"
      End If
   End If
   '下載電子公文
   strFileName = pa(1) & pa(2) & IIf(pa(4) <> "00", "-" & pa(3) & "-" & pa(4), IIf(pa(3) <> "0", "-" & pa(3), "")) & ".1204.PDF"
   strTmpFileName = pa(1) & pa(2) & IIf(pa(3) & pa(4) <> "000", pa(3) & pa(4), "") & "_Official Notice.pdf"
   strFullFileName = strTempFolder & "\" & strTmpFileName
   If PUB_GetAttachFile_CPP(stCP09, strFileName, strFullFileName, True) = False Then
      Err.Raise 999, , "電子公文下載失敗！"
   End If
   strAtt = strFullFileName
   
   strTmpFileName = pa(1) & pa(2) & IIf(pa(3) & pa(4) <> "000", pa(3) & pa(4), "") & "_Letter.pdf"
   strFullFileName = strTempFolder & "\" & strTmpFileName
   strFileName = pa(1) & pa(2) & IIf(pa(4) <> "00", "-" & pa(3) & "-" & pa(4), IIf(pa(3) <> "0", "-" & pa(3), "")) & ".1204.CUS.PDF"
   strExc(1) = PUB_GetLanguage(pa(1), pa(2), pa(3), pa(4))
   strTmp = "01"
   If strExc(1) = "3" Then
      strTmp = "02"
   End If
   strUserLevel = "發FC郵件" 'Added by Lydia 2025/05/21 (參考frm1105)這電子檔是要E給客戶的,因此不要加蓋Confirmation的章 ; Ex. FCP-073579
   NowPrint stCP09, "17", strTmp, True, strUserNum
   strUserLevel = "" 'Added by Lydia 2025/05/21 取消
   g_WordAp.ActiveDocument.ExportAsFixedFormat OutputFileName:=strFullFileName, ExportFormat:=17, OpenAfterExport:=False
   g_WordAp.Quit wdDoNotSaveChanges
   Set g_WordAp = Nothing
   If PUB_ChkFileStatus(strFullFileName, False, strMsg) = False Then
      Err.Raise 999, , strMsg
   Else
      Set oFile = oFileSys.GetFile(strFullFileName)
      '上傳卷宗區
      If SaveAttFile_PDF(stCP09, strFullFileName, strFileName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False) = False Then
         Err.Raise 999, , "上傳卷宗區失敗！"
      End If
   End If
   strAtt = strFullFileName & ";" & strAtt
            
   frm090401.OutCallCP10 = "1204"
   frm090401.Hide
   frm090401.strAttach = strAtt
   frm090401.Text1 = pa(1)
   frm090401.Text2 = pa(2)
   frm090401.Text3 = pa(3)
   frm090401.Text4 = pa(4)
   Call frm090401.Read
   Call frm090401.cmdFCMail_Click(1)
   Unload frm090401
                  
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
   Set g_WordAp = Nothing
   Set oFileSys = Nothing
   Set oFile = Nothing
End Sub
