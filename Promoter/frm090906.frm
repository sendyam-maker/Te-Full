VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090906 
   BorderStyle     =   1  '單線固定
   Caption         =   "確認翻譯人員"
   ClientHeight    =   4476
   ClientLeft      =   420
   ClientTop       =   4416
   ClientWidth     =   8940
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4476
   ScaleWidth      =   8940
   Begin VB.CommandButton cmdBack 
      Caption         =   "退回(&S)"
      Height          =   360
      Left            =   3930
      TabIndex        =   40
      Top             =   120
      Width           =   900
   End
   Begin VB.TextBox txtTFA05 
      Height          =   270
      Left            =   6240
      TabIndex        =   39
      Top             =   4050
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtTF 
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Index           =   26
      Left            =   3840
      TabIndex        =   37
      Top             =   2865
      Width           =   855
   End
   Begin VB.TextBox txtTF 
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Index           =   32
      Left            =   7080
      TabIndex        =   36
      Top             =   2865
      Width           =   855
   End
   Begin VB.TextBox txtTF 
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Index           =   23
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   35
      Top             =   2490
      Width           =   855
   End
   Begin VB.TextBox txtTF 
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Index           =   20
      Left            =   7080
      TabIndex        =   34
      Top             =   2490
      Width           =   1455
   End
   Begin VB.TextBox txtTF 
      Alignment       =   1  '靠右對齊
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Index           =   19
      Left            =   3840
      MaxLength       =   3
      TabIndex        =   33
      Top             =   2490
      Width           =   495
   End
   Begin VB.TextBox txtPA 
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Index           =   10
      Left            =   4920
      MaxLength       =   160
      TabIndex        =   26
      Top             =   645
      Width           =   1065
   End
   Begin VB.TextBox txtPA 
      Height          =   270
      Index           =   1
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   17
      Text            =   "FCP"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtPA 
      Height          =   270
      Index           =   2
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   16
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtPA 
      Height          =   270
      Index           =   3
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   15
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox txtPA 
      Height          =   270
      Index           =   4
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   14
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtPA 
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Index           =   26
      Left            =   1080
      MaxLength       =   9
      TabIndex        =   13
      Top             =   1860
      Width           =   910
   End
   Begin VB.TextBox txtPA 
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Index           =   75
      Left            =   1080
      MaxLength       =   9
      TabIndex        =   12
      Top             =   2175
      Width           =   910
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "上班"
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   1
      Top             =   3960
      Width           =   855
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "下班"
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   0
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "卷宗區(&C)"
      Height          =   300
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   1000
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "原始檔(&P)"
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1000
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "回前畫面(&U)"
      Height          =   360
      Left            =   7485
      TabIndex        =   5
      Top             =   120
      Width           =   1160
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同意(&S)"
      Default         =   -1  'True
      Height          =   360
      Left            =   6390
      TabIndex        =   4
      Top             =   120
      Width           =   900
   End
   Begin MSForms.TextBox txtTF36 
      Height          =   510
      Left            =   1590
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   3180
      Width           =   7095
      VariousPropertyBits=   -1467989985
      BackColor       =   -2147483644
      MaxLength       =   20
      Size            =   "12515;900"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblPA75 
      Height          =   255
      Left            =   2070
      TabIndex        =   47
      Top             =   2175
      Width           =   6660
      BackColor       =   -2147483644
      VariousPropertyBits=   27
      Size            =   "11747;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblPA26 
      Height          =   255
      Left            =   2070
      TabIndex        =   46
      Top             =   1860
      Width           =   6660
      BackColor       =   -2147483644
      VariousPropertyBits=   27
      Size            =   "11747;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPA07 
      Height          =   285
      Left            =   1560
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1560
      Width           =   7200
      VariousPropertyBits=   671105055
      BackColor       =   -2147483644
      MaxLength       =   20
      Size            =   "12700;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPA06 
      Height          =   285
      Left            =   1560
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1230
      Width           =   7200
      VariousPropertyBits=   671105055
      BackColor       =   -2147483644
      MaxLength       =   20
      Size            =   "12700;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPA05 
      Height          =   285
      Left            =   1560
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   900
      Width           =   7200
      VariousPropertyBits=   671105055
      BackColor       =   -2147483644
      MaxLength       =   20
      Size            =   "12700;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   285
      Left            =   1260
      TabIndex        =   42
      Top             =   3930
      Width           =   2160
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3810;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "翻譯特殊指示："
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   41
      Top             =   3210
      Width           =   1305
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   8640
      Y1              =   3780
      Y2              =   3780
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "Label27"
      Height          =   225
      Index           =   1
      Left            =   1200
      TabIndex        =   38
      Top             =   2880
      Width           =   1050
   End
   Begin VB.Label lblTF33 
      AutoSize        =   -1  'True
      Caption         =   "中說4個月不得延"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   4980
      TabIndex        =   32
      Top             =   210
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "相似案號："
      Height          =   180
      Index           =   8
      Left            =   6120
      TabIndex        =   31
      Top             =   2490
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "相似度：             %"
      Height          =   180
      Index           =   5
      Left            =   3120
      TabIndex        =   30
      Top             =   2490
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "原文字數："
      Height          =   180
      Index           =   4
      Left            =   240
      TabIndex        =   29
      Top             =   2490
      Width           =   900
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "Label27"
      Height          =   225
      Index           =   0
      Left            =   7200
      TabIndex        =   28
      Top             =   645
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "類別:"
      Height          =   180
      Index           =   3
      Left            =   6720
      TabIndex        =   27
      Top             =   645
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請日:"
      Height          =   180
      Index           =   2
      Left            =   4320
      TabIndex        =   25
      Top             =   645
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   24
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   23
      Top             =   900
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "(中):"
      Height          =   180
      Index           =   1
      Left            =   1080
      TabIndex        =   22
      Top             =   900
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(英):"
      Height          =   180
      Left            =   1080
      TabIndex        =   21
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(外):"
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   20
      Top             =   1560
      Width           =   345
   End
   Begin VB.Label Lbl 
      Caption         =   "代理人："
      Height          =   225
      Index           =   75
      Left            =   240
      TabIndex        =   19
      Top             =   2175
      Width           =   855
   End
   Begin VB.Label Lbl 
      Caption         =   "申請人1："
      Height          =   225
      Index           =   26
      Left            =   240
      TabIndex        =   18
      Top             =   1860
      Width           =   855
   End
   Begin VB.Label lblTransKind 
      AutoSize        =   -1  'True
      Caption         =   "固定報價、有折扣、有相似度"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   6030
      TabIndex        =   11
      Top             =   4260
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本案不能下班翻譯的原因："
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   0
      Left            =   3870
      TabIndex        =   10
      Top             =   4260
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label Label1 
      Caption         =   "只交Claims期限："
      Height          =   255
      Index           =   7
      Left            =   5640
      TabIndex        =   9
      Top             =   2865
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "交稿期限："
      Height          =   255
      Index           =   6
      Left            =   2880
      TabIndex        =   8
      Top             =   2865
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "承辦期限："
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   2865
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "員工編號："
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   975
   End
End
Attribute VB_Name = "frm090906"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/24 改成Form2.0 ; Combo1、txtPA(5)~(7)=> txtPA05~txtPA07、Label27(26)=>lblPA26 , Label27(75)=>lblPA75、txtTF(36)=>txtTF36
'Memo by Lydia 2019/06/19 依照Sharon和Owen的意見,主要修改如下列
'1.有認領人員則下拉選單出現該人員，沒有才需要列出全組工程師=>在1080603(Owen)工程師認領翻譯案 - 主管確認.doc有另外備註要保留,需要確認
'2.分成"同意or退回"翻譯人員
'3.不可下班翻譯的案件,由工程師發呈報Email給工程師主管, 主管再將Email轉寄何主祕呈報，等待主秘回覆Email；
   '收到主祕回覆後，主管需回系統確認作業按下同意or退回，選擇同意會一併在進度備註:「此為有折扣/固定報價/相似度，經呈報同意承接下班翻譯」；選擇退回只發退回Email並CC給Sharon。
'end 2019/06/19
'Created by Lydia 2018/09/28 工程師認翻譯-主管確認(確認翻譯人員)
Option Explicit
Dim m_PrevForm As Form '前一畫面
Dim m_UserNo As String   '傳入員工編號
Dim m_UserSt16 As String '工程師組別
Dim pa(1 To 4) As String '1~4本所案號
Dim mPA08 As String '專利類型
Dim m_TF01 As String '中說收文號
Dim mTransKind As String '不可下班翻譯的原因
Dim m_TFA06 As String '是否有主管確認
Dim cmdState As String 'Added by Lydia 2019/06/19 存檔的狀態(0-主管指定工程師進行下班翻譯, 1-同意, 2-退回)
'Modified by Lydia 2025/06/05 更改名稱
'Dim m_strBASF As String 'Added by Lydia 2023/04/19 BASF集團的X編號
Dim m_str所內譯 As String
Dim m_str所內譯例外 As String 'Added by Lydia 2025/07/01

Public Sub SetParent(ByRef fm As Form, ByVal pCase As String, ByVal pKeyNo As String, ByVal pUser As String, ByVal pUserGrp As String, Optional ByVal pTransKind As String = "")
   Set m_PrevForm = fm
   m_UserNo = pUser
   m_UserSt16 = pUserGrp
   m_TF01 = pKeyNo
   Call ChgCaseNo(Replace(pCase, "-", ""), pa)
   If pTransKind <> "" Then mTransKind = Mid(pTransKind, 3, Len(pTransKind) - 3)
End Sub

Private Function SaveDatabase() As Boolean

On Error GoTo Err01
    SaveDatabase = False
        cnnConnection.BeginTrans
           Select Case cmdState 'Added by Lydia 2019/06/19 判斷存檔的狀態
                Case "0", "1"
                    If Combo1.Tag <> Trim(Left(Combo1.Text, 6)) Then '有變更人員
                         strSql = "delete from transfeeassign where tfa01='" & m_TF01 & "' "
                         cnnConnection.Execute strSql
                         If Trim(Left(Combo1.Text, 6)) <> "" Then
                             'Added by Lydia 2019/06/19 主管指定工程師進行下班翻譯
                             'Remove by Lydia 2019/09/12 取消下班翻譯控制
                             'If cmdState = "0" Then
                             '   strSql = "insert into transfeeassign (tfa01,tfa02,tfa03,tfa04,tfa05) select " & _
                             '                       "'" & m_TF01 & "', to_char(sysdate, 'YYYYMMDD') , to_char(sysdate, 'HH24MISS'), '" & Trim(Left(Combo1.Text, 6)) & "'," & CNULL(txtTFA05.Text) & _
                            '                        " from dual "
                            '    cnnConnection.Execute strSql
                            ' Else
                             'end 2019/06/19
                                strSql = "insert into transfeeassign (tfa01,tfa02,tfa03,tfa04,tfa05,tfa06,tfa07,tfa08) select " & _
                                                    "'" & m_TF01 & "', to_char(sysdate, 'YYYYMMDD') , to_char(sysdate, 'HH24MISS'), '" & Trim(Left(Combo1.Text, 6)) & "'," & CNULL(txtTFA05.Text) & _
                                                    ", '" & strUserNum & "', to_char(sysdate, 'YYYYMMDD'), to_char(sysdate, 'HH24MISS') from dual "
                                cnnConnection.Execute strSql
                             'End If 'end 2019/09/12
                         End If
                    Else
                         strSql = "update transfeeassign set tfa06='" & strUserNum & "', tfa07=" & strSrvDate(1) & ", tfa08=" & Val(Format(ServerTime, "000000")) & _
                                     IIf(txtTFA05.Text <> txtTFA05.Tag, ", tfa05=" & CNULL(txtTFA05.Text), "") & " where tfa01='" & m_TF01 & "' "
                         cnnConnection.Execute strSql
                    End If
                    'Added by Lydia 2019/06/19 在新案翻譯的進度備註加註
                    'Remove by Lydia 2019/09/12 取消下班翻譯控制
                    'If cmdState = "1" And mTransKind <> "" And Trim(Left(Combo1.Text, 6)) <> "" And Left(Combo1.Text, 1) <> "F" And txtTFA05 = "A" Then
                    '    strSql = "update caseprogress set cp64='" & ChangeWStringToWDateString(strSrvDate(1)) & " 認翻譯備註:此為" & mTransKind & "，經呈報同意承接下班翻譯;'||cp64 where cp09='" & m_TF01 & "' "
                    '    cnnConnection.Execute strSql
                    'End If
                    'end 2019/09/12
                    
                'Added by Lydia 2019/06/19
                Case Else '退回
                         strSql = "delete from transfeeassign where tfa01='" & m_TF01 & "' "
                         cnnConnection.Execute strSql
            End Select
        cnnConnection.CommitTrans
    SaveDatabase = True
    
Err01:
If Err.Number <> 0 Then
   MsgBox Err.Description
End If
End Function
'
Private Sub cmdExit_Click()
   If CheckDiff = True Then
       If MsgBox("你並未存檔，確定離開嗎?", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
           Exit Sub
       End If
   End If
   Unload Me
End Sub

Private Function CheckDiff() As Boolean
    
    CheckDiff = False
    
    If Combo1.Tag & txtTFA05.Tag & IIf(m_TFA06 <> "", "Y", "") <> Trim(Left(Combo1.Text, 6)) & txtTFA05.Text & "Y" Then
        CheckDiff = True
    End If
    'Added by Lydia 2019/08/14
    'Modified by Lydia 2025/03/13 改用模組取得
    'If Combo1.Text <> "" And InStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達, Trim(Left(Combo1.Text, 6))) > 0 Then
    If Combo1.Text <> "" And InStr(Pub_SetF51Order("F", ""), Trim(Left(Combo1.Text, 6))) > 0 Then
        Chk2(0).Value = False
        Chk2(1).Value = False
    End If
End Function

Private Sub Chk2_Click(Index As Integer)
    If Chk2(Index).Value = 1 Then
        If Index = 0 Then
            Chk2(1).Value = 0
            txtTFA05 = "A" '下班
        Else
            Chk2(0).Value = 0
            txtTFA05 = "B" '上班
        End If
    Else
        If Chk2(0).Value = 0 And Chk2(0).Value = 0 Then
             txtTFA05 = ""
        End If
    End If
End Sub

Private Sub cmdok_Click()
Dim tmpBol As Boolean
Dim dbTfRate As Double, bolIsHigher As Boolean  'Added by Lydia 2021/07/29 判斷翻譯費折扣率＞30%
Dim m_strMemo As String 'Added by Lydia 2022/07/12

    If CheckDiff = True Then
        If Trim(Left(Combo1.Text, 6)) <> "" And Left(Combo1.Text, 1) <> "F" And txtTFA05 = "" Then
              MsgBox "請勾選上班或下班翻譯!", vbCritical
              Exit Sub
        End If
        'Added by Lydia 2023/04/19 外專翻譯分案承辦人不得為翻譯社及外譯人員
        If Trim(Left(Combo1.Text, 6)) <> "" And (InStr(m_str所內譯, txtPA(26)) > 0 Or InStr(m_str所內譯, txtPA(75)) > 0) Then
            If Trim(Left(Combo1.Text, 1)) = "F" Then
                strExc(2) = PUB_GetMapID(Trim(Left(Combo1.Text, 6)), 1)
                If strExc(2) = "" Then
                    'Modified by Lydia 2025/06/05 「BASF集團公司為申請人的所有專利案件」改為「本案所有」
                    MsgBox "本案所有相關翻譯事宜（201新案翻譯/927其他翻譯）皆須由本所工程師翻譯/處理，不得委外。", vbExclamation + vbOKOnly
                    Combo1.SetFocus
                    Exit Sub
                End If
            End If
            dbTfRate = 0
        Else
        'end 2023/04/19
            'Added by Lydia 2021/07/30 翻譯費折扣率＞30%客戶只能上班譯; 可排除特別客戶
            dbTfRate = PUB_GetTransFeeRate(txtPA(1), txtPA(2), txtPA(3), txtPA(4), , bolIsHigher, True)
            '控制翻譯費折扣率＞30%客戶案件之承辦人只能為所內人員上班譯編號。
            If dbTfRate > 30 Then
                If Chk2(0).Value = 1 Or Left(Combo1.Text, 1) = "F" Then
                    'Modifed by Lydia 2022/07/12 改在email內加註
                    'MsgBox "該案件之承辦人只能為所內人員上班譯編號！", vbExclamation
                    'Exit Sub
                    m_strMemo = m_strMemo & vbCrLf & "該案件客戶翻譯費折扣率＞30%，請注意是否已經呈報主管。"
                    'end 2022/07/12
                End If
            ElseIf bolIsHigher = True Then  '折扣率＞30%但是例外控制的客戶
                 '不受限
            End If
            'end 2021/07/30
        End If 'Added by Lydia 2023/04/19
        'Added by Lydia 2021/04/14 外專翻譯承辦及核稿期限控管：
        '工程師主管確認時，查詢該認領人員，新案翻譯未上完稿日案件彈訊息：尚未完稿案件FCPxxxx，承辦期限：
        'Modified by Lydia 2025/03/13 改用模組取得
        'If InStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達, Trim(Left(Combo1.Text, 6))) = 0 Then
        If InStr(Pub_SetF51Order("F", ""), Trim(Left(Combo1.Text, 6))) = 0 Then
            strExc(4) = Pub_GetEngEP09List(Trim(Left(Combo1.Text, 6)))
            If strExc(4) <> "" Then
                If MsgBox(Trim(Replace(Mid(Combo1.Text, 6), "△", "")) & " 尚未完稿案件：" & strExc(4) & vbCrLf & vbCrLf & "是否繼續確認？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
        'end 2021/04/14
        
        cmdState = "1" 'Added by Lydia 2019/06/19
        
        'Modified by Lydia 2019/06/19 不可下班翻譯的案件,由工程師發呈報Email給工程師主管, 主管再將Email轉寄何主祕呈報，等待主秘回覆Email；
'        If mTransKind <> "" And Trim(Left(Combo1.Text, 6)) <> "" And Left(Combo1.Text, 1) <> "F" And txtTFA05 = "A" Then
'            If MsgBox("本案不能下班翻譯：" & mTransKind & vbCrLf & "是否下班翻譯？", vbYesNo + vbDefaultButton2) = vbNo Then
'                  Chk2(1).Value = 1 '改為上班
'            Else
'                  frm880019.txtSubject = pa(1) & pa(2) & IIf(pa(3) & pa(4) <> "000", pa(3) & pa(4), "") & " 確認新案翻譯人員(" & Trim(Replace(Combo1.Text, "△", "")) & "-下班翻譯)"
'                  frm880019.txtContent = "※本案不能下班翻譯：" & mTransKind & vbCrLf
'                  frm880019.txtReceiver = "68009" '上呈給何主秘
'                  frm880019.SetParent Me
'                  frm880019.Show vbModal
'                  tmpBol = frm880019.m_bolDone '是否傳送成功
'                  Unload frm880019
'                  If tmpBol = False Then
'                      MsgBox "送信失敗，請重新Email !", vbCritical, "取消待比對"
'                      Exit Sub
'                  End If
'            End If
'        End If
        'Remove by Lydia 2019/09/12 取消下班翻譯控制
'        If ((Combo1.Tag = "" And txtTFA05.Tag = "") Or (Combo1.Tag <> "" And Combo1.Tag <> Trim(Left(Combo1.Text, 6)))) And _
'                mTransKind <> "" And Trim(Left(Combo1.Text, 6)) <> "" And Left(Combo1.Text, 1) <> "F" And txtTFA05 = "A" Then
'            If MsgBox("本案不能下班翻譯：" & mTransKind & vbCrLf & "若選擇下班翻譯需要先呈報主管，是否先發Email？", vbYesNo + vbDefaultButton2) = vbNo Then
'                  Exit Sub
'            Else
'ReEmail:
'                  frm880019.txtSubject = pa(1) & pa(2) & IIf(pa(3) & pa(4) <> "000", pa(3) & pa(4), "") & " 欲認領新案翻譯人員(" & Trim(Replace(Combo1.Text, "△", "")) & "-下班翻譯)，呈報主管"
'                  frm880019.txtContent = "※本案不能下班翻譯：" & mTransKind & vbCrLf & vbCrLf & _
'                                                     "欲承接下班翻譯，請呈報主管"
'                  frm880019.txtReceiver = "68009" '上呈給何主秘
'                  frm880019.cmdAttach.Visible = False
'                  frm880019.SetParent Me
'                  frm880019.Show vbModal
'                  tmpBol = frm880019.m_bolDone '是否傳送成功
'                  Unload frm880019
'                  If tmpBol = False Then
'                      If MsgBox("送信失敗，是否重新Email？ ", vbCritical + vbYesNo + vbDefaultButton2, "呈報主管") = vbYes Then
'                          GoTo ReEmail
'                      Else
'                          Exit Sub
'                      End If
'                  End If
'                  cmdState = "0"
'            End If
'        End If
        'end 2019/06/19
        'end 2019/09/12
        
        If SaveDatabase = True Then
            strExc(2) = ""
            strExc(3) = ""
             '通知所內員工
            'Memo by Lydia 2019/06/04 與Sharon再次確認：1.給所內：cc給她知道案件的情況；2.給翻譯社：不用發Email直接在翻譯分案處理
            'Modified by Lydia 2019/06/19 排除主管指定人員cmdState=0
            If cmdState <> "0" And (Trim(Left(Combo1.Text, 6)) <> "" And Left(Combo1.Text, 1) <> "F") Or Left(Combo1.Tag, 1) <> "F" Then
                 'Modified by Lydia 2023/03/14 主管指定所外譯者F51,不在下拉選單內
                 'strExc(0) = pa(1) & pa(2) & IIf(pa(3) & pa(4) <> "000", pa(3) & pa(4), "") & " 確認新案翻譯人員(" & Trim(Mid(Replace(Combo1.Text, "△", ""), 6)) & IIf(txtTFA05 <> "", IIf(txtTFA05 = "A", "-下班", "-上班"), "") & ")"
                 strExc(0) = pa(1) & pa(2) & IIf(pa(3) & pa(4) <> "000", pa(3) & pa(4), "") & " 確認新案翻譯人員(" & GetStaffName(Trim(Left(Replace(Combo1.Text, "△", ""), 6))) & IIf(txtTFA05 <> "", IIf(txtTFA05 = "A", "-下班", "-上班"), "") & ")"
                 
                 'Remove by Lydia 2019/09/12 取消下班翻譯控制
                 'If mTransKind <> "" And txtTFA05 = "A" Then
                 '    strExc(1) = "※本案不能下班翻譯：" & mTransKind & vbCrLf
                 'Else
                     strExc(1) = "同主旨" & vbCrLf
                 'End If 'end 2019/09/12
                 strExc(1) = strExc(1) & m_strMemo 'Added by Lydia 2022/07/12 改在內文加註：控制翻譯費折扣率＞30%客戶案件之承辦人只能為所內人員上班譯編號
                 
                 strExc(3) = Pub_GetSpecMan("M") 'CC：翻譯分案人員
                  'Remove by Lydia 2019/08/14 若修改人員,則分別通知雙方
                  'If Combo1.Tag & txtTFA05.Tag <> "" And Combo1.Tag & txtTFA05.Tag <> Trim(Left(Combo1.Text, 6)) & txtTFA05.Text Then
                  '    strExc(1) = strExc(1) & vbCrLf & "原認領人員：" & Combo1.Tag & " " & GetStaffName(Combo1.Tag) & IIf(txtTFA05.Tag <> "", IIf(txtTFA05.Tag = "A", "-下班", "-上班"), "")
                  '    strExc(3) = strExc(3) & ";" & Combo1.Tag
                  'End If
                  'end 2019/08/14
                  If Left(Combo1.Text, 1) <> "F" Then
                      strExc(2) = Trim(Left(Combo1.Text, 6))
                  Else
                      strExc(2) = strExc(3)
                      strExc(3) = ""
                  End If
                  'Modified by Lydia 2019/08/14 改成mailcache
                  'If strExc(2) <> "" Then PUB_SendMail strUserNum, strExc(2), "", strExc(0), strExc(1), , , , , , strExc(3)
                  If strExc(2) <> "" Then
                      '同意
                       strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                  " values ('" & strUserNum & "','" & strExc(2) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & strExc(0) & "','" & strExc(1) & "','" & strExc(3) & "')"
                       cnnConnection.Execute strSql
                      '退回(若修改人員,則分別通知雙方)
                      If Combo1.Tag <> "" And Combo1.Tag <> Trim(Left(Combo1.Text, 6)) And Left(Combo1.Tag, 1) <> "F" Then
                          strExc(0) = pa(1) & pa(2) & IIf(pa(3) & pa(4) <> "000", pa(3) & pa(4), "") & " 退回認翻譯"
                          strExc(1) = "同主旨" & vbCrLf
                          strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                                       " values ('" & strUserNum & "','" & Combo1.Tag & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & strExc(0) & "','" & strExc(1) & "')"
                          cnnConnection.Execute strSql
                      End If
                  End If
                  'end 2019/08/14
            End If
        End If
    End If
    Unload Me
End Sub

'Added by Lydia 2019/06/19
Private Sub cmdBack_Click()

    cmdState = "2"
    If SaveDatabase = True Then
        strExc(2) = ""
        strExc(3) = ""
        If (Trim(Left(Combo1.Text, 6)) <> "" And Left(Combo1.Text, 1) <> "F") Or Left(Combo1.Tag, 1) <> "F" Then '通知所內員工
             strExc(0) = pa(1) & pa(2) & IIf(pa(3) & pa(4) <> "000", pa(3) & pa(4), "") & " 退回認翻譯(" & Trim(Replace(Combo1.Text, "△", "")) & IIf(txtTFA05 <> "", IIf(txtTFA05 = "A", "-下班", "-上班"), "") & ")"
             'Remove by Lydia 2019/09/12 取消下班翻譯控制
             'If mTransKind <> "" And txtTFA05 = "A" Then
             '    strExc(1) = "※本案不能下班翻譯：" & mTransKind & vbCrLf
             '    strExc(3) = Pub_GetSpecMan("M") 'CC：翻譯分案人員 (不可下班翻譯之退回才要CC)
             'Else
                 strExc(1) = "同主旨" & vbCrLf
             'End If 'end 2019/09/12
             
             If Left(Combo1.Text, 1) <> "F" Then
                  strExc(2) = Trim(Left(Combo1.Text, 6))
             Else
                 strExc(2) = strExc(3)
             End If
             'Modified by Lydia 2019/08/14 改成mailcache
             'If strExc(2) <> "" Then PUB_SendMail strUserNum, strExc(2), "", strExc(0), strExc(1), , , , , , strExc(3)
             strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                        " values ('" & strUserNum & "','" & strExc(2) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & strExc(0) & "','" & strExc(1) & "','" & strExc(3) & "')"
             cnnConnection.Execute strSql
                
        End If
        Unload Me 'Added by Lydia 2019/11/25
    End If
End Sub

Private Sub cmdOpen_Click(Index As Integer)
Dim hLocalFile As Long

On Error GoTo ErrHand01
    Select Case Index
            Case 0 '外文本
                'Modified by Lydia 2021/04/14 改放在原始檔區
                'strExc(1) = Pub_GetFCPcaseFilePath(pa(2), , pa(1))
                'If Dir(strExc(1) & "\*.*") <> "" Then
                '      ShellExecute hLocalFile, "explore", strExc(1), vbNullString, vbNullString, 1
                'Else
                '     MsgBox pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & "在" & strExc(1) & "的資料夾不存在或無檔案!", vbInformation
                'End If
                If PUB_CheckFormExist("frm100101_M") Then
                    MsgBox "請先關閉共同查詢〔原始檔區〕畫面！"
                    Exit Sub
                End If
                If cmdOpen(Index).Tag = "" Then
                    MsgBox txtPA(1) & "-" & txtPA(2) & "在〔原始檔區〕的English_Vers收文號不存在!", vbInformation
                Else
                    frm100101_M.m_strKey = cmdOpen(Index).Tag '收文號
                    frm100101_M.SetParent Me
                    If frm100101_M.QueryData = True Then
                       frm100101_M.Show
                       Me.Hide
                    End If
                End If
                'end 2021/04/14
             Case 1 '卷宗區
                Me.Enabled = False
                Screen.MousePointer = vbHourglass
                frm100101_L.m_strKey = pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)
                frm100101_L.SetParent Me
                If frm100101_L.QueryData = True Then
                   frm100101_L.Show
                   Me.Hide
                End If
                Screen.MousePointer = vbDefault
                Me.Enabled = True
    End Select
    Exit Sub
    
ErrHand01:
    If Err.Number <> 0 Then
         '全部錯誤訊息統一
         MsgBox "無法讀取，請通知電腦中心！", vbCritical
         Resume Next
    End If

End Sub

'Added by Lydia 2019/08/14
Private Sub Combo1_LostFocus()
    
    'Modified by Lydia 2025/03/13 改用模組取得
    'If Combo1.Text <> "" And InStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達, Trim(Left(Combo1.Text, 6))) > 0 Then
    If Combo1.Text <> "" And InStr(Pub_SetF51Order("F", ""), Trim(Left(Combo1.Text, 6))) > 0 Then
        Chk2(0).Value = False
        Chk2(1).Value = False
    'Added by Lydia 2023/03/14
    Else
        Dim intX As Integer
        intX = -1
         For intI = 0 To Combo1.ListCount - 1
             If InStr(Combo1.List(intI), Trim(Combo1.Text)) > 0 Then
                 intX = intI
                 Exit For
             End If
        Next intI
        If intX = -1 Then
             If ByInputGetST01or02(Trim(Left(Combo1.Text, 6)), strExc(0), strExc(1)) = False Then
                 Combo1.SetFocus
                 'Combo1.Tag = Trim(Left(Combo1.Text, 6)) 'Mark by Lydia 2023/06/15 保留原記錄
                 Exit Sub
             Else
                 Combo1.Text = strExc(0) & " " & strExc(1)
             End If
        Else
             Combo1.ListIndex = intX
        End If
    'end 2023/03/14
    End If
    'Combo1.Tag = Trim(Left(Combo1.Text, 6)) 'Added by Lydia 2023/03/14 'Mark by Lydia 2023/06/15 保留原記錄
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Lydia 2019/08/14
   If UCase(TypeName(m_PrevForm)) <> "NOTHING" Then
        m_PrevForm.Show
        If TypeName(m_PrevForm) = "frm060122" Then
             m_PrevForm.cmdState = 0
             Call m_PrevForm.PubShowNextData
        End If
   End If
   
   Set frm090906 = Nothing
End Sub

Private Sub ClearForm()
Dim oLbl As LABEL
Dim oTxt As TextBox

   For Each oTxt In txtPA
      oTxt.Text = ""
      oTxt.Locked = True
      If oTxt.Index > 4 Then
          oTxt.BackColor = &H8000000F
      End If
   Next
   
   For Each oLbl In Label27
      oLbl.Caption = ""
   Next
   
   For Each oTxt In txtTF
      oTxt.Text = ""
      oTxt.Locked = True
      oTxt.BackColor = &H8000000F
   Next
   lblTF33.Visible = False
   lblTransKind.Caption = ""
   
   'Added by Lydia 2021/09/24
   txtPA05 = "": txtPA06 = "": txtPA07 = ""
   lblPA26 = "": lblPA75 = ""
   txtTF36 = ""
   'end 2021/09/24
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   Combo1.Clear
   ClearForm
   'Modified by Lydia 2025/06/05 更改名稱
   'm_strBASF = Pub_GetSpecMan("外專翻譯分案-BASF") & ","  'Added by Lydia 2023/04/19
   m_str所內譯 = Pub_GetSpecMan("外專翻譯分案-所內譯") & ","
   m_str所內譯例外 = Pub_GetSpecMan("外專翻譯分案-所內譯例外") & "," 'Added by Lydia 2025/07/01
   
   If ReadData = True Then
      Call SetCombo1
   Else
   
   End If

End Sub

Private Function ReadData() As Boolean

   txtPA(1) = pa(1):     txtPA(2) = pa(2)
   txtPA(3) = pa(3):     txtPA(4) = pa(4)
   
   '客戶名稱:中->英->日 ; 代理人名稱: 英->中->日
   strExc(0) = "SELECT PA01,PA02,PA03,PA04,PA05,PA06,PA07,PA08,PA10,PA158,PA26,PA75," & _
                     " NVL(CU04,NVL(CU05,CU06)) CNAME,NVL(FA05,NVL(FA04,FA06)) FNAME" & _
                     " FROM PATENT,CUSTOMER,FAGENT" & _
                     " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4)) & _
                     " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+)" & _
                     " AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)"
    intI = 0
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        'Modified by Lydia 2021/09/24 txtPA(5)~(7)=> txtPA05~txtPA07
        txtPA05 = "" & RsTemp.Fields("PA05")
        txtPA06 = "" & RsTemp.Fields("PA06")
        txtPA07 = "" & RsTemp.Fields("PA07")
        'end 2021/09/24
        mPA08 = "" & RsTemp.Fields("PA08")
        txtPA(10) = ChangeTStringToTDateString(TransDate("" & RsTemp.Fields("PA10"), 1))
        txtPA(26).Text = "" & RsTemp.Fields("PA26")
        txtPA(75).Text = "" & RsTemp.Fields("PA75")
        'Modified by Lydia 2021/09/24 Label27(26)=>lblPA26 , Label27(75)=>lblPA75
        lblPA26.Caption = "" & RsTemp.Fields("CNAME")
        lblPA75.Caption = "" & RsTemp.Fields("FNAME")
        'end 2021/09/24
        lblTransKind.Caption = mTransKind
        
        strExc(0) = "select a.*,b.* from transfee a , caseprogress,transfeeassign b " & _
                          "where tf01='" & m_TF01 & "' and tf01=cp09(+) and cp159=0 and tf01=tfa01(+) "
        intI = 0
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
             '原文字數
             txtTF(23) = "" & RsTemp.Fields("tf23")
             '相似度
             txtTF(19) = "" & RsTemp.Fields("tf19")
             '相似案號
             txtTF(20) = "" & RsTemp.Fields("tf20")
             '交稿期限
             txtTF(26) = ChangeTStringToTDateString(TransDate("" & RsTemp.Fields("tf26"), 1))
             '只交Claims期限
             txtTF(32) = ChangeTStringToTDateString(TransDate("" & RsTemp.Fields("tf32"), 1))
             'Added by Lydia 2019/08/23 翻譯特殊指示
             'Modified by Lydia 2021/09/24 txtTF(36)=>txtTF36
             txtTF36 = "" & RsTemp.Fields("tf36")
             
             '承辦期限:參考PUB_GetFCPsetCP48
             Label27(1) = ChangeTStringToTDateString(TransDate(PUB_GetWorkDay1(CompDate(2, 75, strSrvDate(1)), False), 1))
             '中說4個月不得延
             If "" & "" & RsTemp.Fields("tf33") <> "" Then lblTF33.Visible = True
             '認領人員
             Combo1.Tag = "" & RsTemp.Fields("tfa04")
             txtTFA05 = "" & RsTemp.Fields("tfa05")
             txtTFA05.Tag = txtTFA05.Text
             If txtTFA05 = "A" Then
                 Chk2(0).Value = 1
             ElseIf txtTFA05 = "B" Then
                 Chk2(1).Value = 1
             End If
             m_TFA06 = "" & RsTemp.Fields("tfa06")
        End If
        '命名作業
        strExc(0) = "select tct01,tct10,tct27,tct28,decode(tct25,'1','生醫','2','化學','3','化工','4','材料','5','電子','6','機械','7','電機','8','其他',tct25) tct25n from caseprogress,transcasetitle " & _
                          "where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp31='Y' and cp09=tct01 "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
             If "" & RsTemp.Fields("tct25n") <> "" Then
                 Label27(0).Caption = "" & RsTemp.Fields("tct25n")
             End If
        End If
    End If
    'Added by Lydia 2021/04/14 原始檔區掛的收文號
    cmdOpen(1).Tag = ""
    If PUB_ChkCPExist(pa, cntEnglish_Vers, , strExc(1), , "D") = True Then 'English_Vers992
         cmdOpen(1).Tag = strExc(1)
    End If
    'end 2021/04/14
    
    ReadData = True
        
End Function

Private Sub SetCombo1()
Dim rsB As New ADODB.Recordset
Dim intA As Integer
Dim stCon As String
    
    cmdBack.Visible = False 'Added by Lydia 2019/06/19
    If Combo1.Tag <> "" Then
         cmdBack.Visible = True 'Added by Lydia 2019/06/19 有工程師認領,才需要顯示退回
         stCon = " and st01 <> '" & Combo1.Tag & "' "
    End If
    'Added by Lydia 2019/08/14 排除林信昌和第4碼9
    stCon = stCon & " and st01 not in ('F5162','68007','68092','68091','F5644','F5645') and st01 not like '___9%' "
    
    '先抓主任
    strSql = "select 1 ord1, st01,st02,st16 from staff where st04='1' and st03='F21' and st16='" & m_UserSt16 & "' and st20<='52' " & stCon
    '再抓工程師
    strSql = strSql & "union all select 2 ord1, st01,st02,st16 from staff where st04='1' and st03='F21' and st16='" & m_UserSt16 & "' and nvl(st20,'99') > '52' " & stCon
    '國外翻譯
    'Modified by Lydia 2025/03/13 改用模組取得
    'strSql = strSql & "union all select 3 ord1, st01,st02,st16 from staff where st01 in (" & GetAddStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達) & ") "
    strSql = strSql & "union all select 3 ord1, st01,st02,st16 from staff where st01 in (" & GetAddStr(Pub_SetF51Order("F", "")) & ") "
    'Added by Lydia 2019/08/21 增加所外國內譯者翻譯
    Select Case m_UserSt16
         Case "1" '電子電機組
                       '阮威立 F5220,張元銘 F5267,楊志雄 F5219,范揚達 F5198
             strSql = strSql & "union all select 4 ord1, st01,st02,st16 from staff where st01 in ('F5220','F5267','F5219','F5198') "
         Case "4" '機械組
                        '郭稚艷 F5250,林錫增 F5714,李敦維F5616
             strSql = strSql & "union all select 4 ord1, st01,st02,st16 from staff where st01 in ('F5250','F5714','F5616') "
    End Select
    
    '有認領人員(排前面)
    If Combo1.Tag <> "" Then
        'Modified by Lydia 2019/06/19 有認領人員則下拉選單出現該人員，沒有才需要列出全組工程師
        strSql = "select 0 ord1, st01,st02,st16 from staff where st04='1' and st01='" & Combo1.Tag & "' Union all " & strSql
    End If
    strSql = strSql & " order by ord1,st01 "
    
    If strSql <> "" Then
       intI = 1
       Set rsB = ClsLawReadRstMsg(intI, strSql)
       If intI = 1 Then
          rsB.MoveFirst
          Do While Not rsB.EOF
             '排除94099(總經理外專編號) ; 楊雯芳(99033)屬於兼任,先排除
             If Trim("" & rsB.Fields("ST01")) <> "94099" And Trim("" & rsB.Fields("ST01")) <> "99033" Then
                'Modified by Lydia 2019/06/03 有認領人員在名稱後+△
                Combo1.AddItem Trim("" & rsB.Fields("ST01")) & " " & Trim("" & rsB.Fields("ST02")) & IIf("" & rsB.Fields("ord1") = "0", "△", "")
                If Trim("" & rsB.Fields("ST01")) = Combo1.Tag And intA = 0 Then
                    intA = rsB.AbsolutePosition
                End If
             End If
             rsB.MoveNext
          Loop
       End If
       If intA <> 0 Then
            Combo1.ListIndex = intA - 1
       End If
       Set rsB = Nothing
    End If
End Sub


