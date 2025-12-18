VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090902_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "待分案"
   ClientHeight    =   3264
   ClientLeft      =   420
   ClientTop       =   4416
   ClientWidth     =   8916
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3264
   ScaleWidth      =   8916
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   3420
      TabIndex        =   38
      Text            =   "Combo2"
      Top             =   150
      Width           =   2055
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "外文本(&P)"
      Height          =   300
      Left            =   240
      TabIndex        =   32
      Top             =   120
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      Height          =   360
      Index           =   1
      Left            =   7485
      TabIndex        =   29
      Top             =   120
      Width           =   1160
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "存檔(&S)"
      Default         =   -1  'True
      Height          =   360
      Index           =   0
      Left            =   6480
      TabIndex        =   19
      Top             =   120
      Width           =   900
   End
   Begin VB.TextBox txtData 
      Height          =   270
      Index           =   2
      Left            =   9000
      TabIndex        =   18
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtData 
      Height          =   270
      Index           =   1
      Left            =   5880
      MaxLength       =   4
      TabIndex        =   1
      Top             =   620
      Width           =   615
   End
   Begin VB.TextBox txtData 
      Height          =   270
      Index           =   0
      Left            =   5040
      MaxLength       =   7
      TabIndex        =   0
      Top             =   620
      Width           =   855
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "急件"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   620
      Width           =   735
   End
   Begin VB.Label lblTCN 
      Caption         =   "提申急件組別："
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   2130
      TabIndex        =   37
      Top             =   180
      Width           =   1305
   End
   Begin MSForms.Label lblCM 
      Height          =   255
      Left            =   6210
      TabIndex        =   36
      Top             =   2790
      Width           =   2535
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Caption         =   "命名人員："
      Size            =   "4471;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1560
      TabIndex        =   35
      Top             =   2760
      Width           =   2055
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3625;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCMboth 
      AutoSize        =   -1  'True
      Caption         =   "lblCMboth"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3690
      TabIndex        =   34
      Top             =   2790
      Width           =   2445
   End
   Begin VB.Label Label5 
      Caption         =   "命名人員："
      Height          =   255
      Index           =   1
      Left            =   6120
      TabIndex        =   33
      Top             =   945
      Width           =   975
   End
   Begin MSForms.Label lblData 
      Height          =   255
      Index           =   16
      Left            =   3960
      TabIndex        =   31
      Top             =   945
      Width           =   975
      VariousPropertyBits=   27
      Size            =   "1720;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   255
      Index           =   15
      Left            =   1200
      TabIndex        =   30
      Top             =   1260
      Width           =   1215
      VariousPropertyBits=   27
      Size            =   "2143;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   255
      Index           =   11
      Left            =   7080
      TabIndex        =   28
      Top             =   1980
      Width           =   1575
      VariousPropertyBits=   27
      Size            =   "2778;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   27
      Top             =   945
      Width           =   855
      VariousPropertyBits=   27
      Size            =   "1508;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   26
      Top             =   640
      Width           =   855
      VariousPropertyBits=   27
      Size            =   "1508;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "本所期限："
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   25
      Top             =   945
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "法定期限："
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   24
      Top             =   640
      Width           =   975
   End
   Begin MSForms.Label lblData 
      Height          =   435
      Index           =   3
      Left            =   1320
      TabIndex        =   23
      Top             =   1980
      Width           =   1560
      VariousPropertyBits=   27
      Size            =   "2752;767"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "2.中說類型："
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   22
      Top             =   1980
      Width           =   1095
   End
   Begin MSForms.Label lblData 
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   21
      Top             =   1650
      Width           =   1200
      VariousPropertyBits=   27
      Size            =   "2117;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "1.專利種類："
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   20
      Top             =   1650
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "案件命名人員："
      Height          =   255
      Index           =   4
      Left            =   210
      TabIndex        =   17
      Top             =   2790
      Width           =   1335
   End
   Begin MSForms.Label lblData 
      Height          =   255
      Index           =   9
      Left            =   7080
      TabIndex        =   16
      Top             =   1260
      Width           =   1155
      VariousPropertyBits=   27
      Size            =   "2037;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   255
      Index           =   8
      Left            =   7080
      TabIndex        =   15
      Top             =   960
      Width           =   1125
      VariousPropertyBits=   27
      Size            =   "1984;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   90
      X2              =   8730
      Y1              =   2580
      Y2              =   2580
   End
   Begin MSForms.Label lblData 
      Height          =   255
      Index           =   10
      Left            =   7320
      TabIndex        =   14
      Top             =   1650
      Width           =   1335
      VariousPropertyBits=   27
      Size            =   "2355;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "國　　籍："
      Height          =   255
      Index           =   5
      Left            =   6120
      TabIndex        =   13
      Top             =   1650
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   3
      Left            =   6120
      TabIndex        =   12
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label Label4 
      Caption         =   "分案組別："
      Height          =   255
      Index           =   2
      Left            =   6120
      TabIndex        =   11
      Top             =   1980
      Width           =   975
   End
   Begin MSForms.Label lblData 
      Height          =   255
      Index           =   6
      Left            =   3960
      TabIndex        =   10
      Top             =   1980
      Width           =   1485
      VariousPropertyBits=   27
      Size            =   "2619;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   9
      Top             =   1980
      Width           =   975
   End
   Begin MSForms.Label lblData 
      Height          =   255
      Index           =   5
      Left            =   3960
      TabIndex        =   8
      Top             =   1650
      Width           =   1005
      VariousPropertyBits=   27
      Size            =   "1773;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      Caption         =   "收文日期："
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   7
      Top             =   1650
      Width           =   975
   End
   Begin MSForms.Label lblData 
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   6
      Top             =   1260
      Width           =   2055
      VariousPropertyBits=   27
      Size            =   "3625;661"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      Caption         =   "總收文號："
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   5
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "譯畢期限：                ，請於　　　　　　　　　前譯畢名稱"
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   3
      Top             =   645
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1260
      Width           =   975
   End
End
Attribute VB_Name = "frm090902_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/27 改成Form2.0 ; Combo1、lblData(index)、lblCM
'Created by Lydia 2017/11/14 外專新案未命名區-待分案明細
Option Explicit
Dim m_PrevForm As Form '前一畫面
Dim m_UserNo As String   '傳入員工編號
'Modified by Lydia 2018/04/18 +9-申請日,10-公告日(PA14),11-目前准/駁(PA16)
'Memo by Lydia 2021/09/27 改成Form2.0 ; Combo1、lblCM、lblData(index)
'Modified by Lydia 2020/02/17 +12-名稱有特殊字(PA174)
Dim strCase(1 To 12) As String '1~4本所案號pa01~pa04,5-專利種類pa08,6-申請國家pa09,7-分案組別pa150,8-設計案屬性pa158
Dim m_TCT01 As String  '收文號=PK
Dim m_TCT04 As String  '工程師主管
Dim m_TCT07 As String  '工程師主任
Dim m_TCT10 As String  '命名人員編號
Dim bolEmail As Boolean '是否發Email通知
Dim m_TCT27kind As String '欲翻譯此案件者可輸入的選項
Dim bolEmail2 As Boolean '重新改命名人員或期限,補發Email
Dim m_UserSt16 As String 'Added by Lydia 2017/12/29 傳入員工編號的工程師組別
Dim m_Receiver As String, m_ReMailTxt As String 'Added by Lydia 2018/01/03 通知原本分案的人員和郵件內容
Dim n_CP118 As String 'Added by Lydia 2018/04/18 新申請案：是否電子送件
Dim m_GrpMan As String  'Added by Lydia 2024/03/20 各組工程師主管

Private Function CheckDataValidate() As Boolean
Dim pTime As String
Dim Cancel As Boolean

   If Trim(txtData(0) & txtData(1)) <> "" Or Chk1.Value = 1 Then
       Chk1.Value = 1
       If Trim(txtData(0)) = "" Or Trim(txtData(1)) = "" Then
          MsgBox "急件請輸入譯畢期限!", vbExclamation
          If Trim(txtData(0)) = "" Then
             txtData(0).SetFocus
             Txtdata_GotFocus 0
          Else
             txtData(1).SetFocus
             Txtdata_GotFocus 1
          End If
          Exit Function
       End If
       For intI = 0 To 1
          Txtdata_Validate intI, Cancel
          If Cancel = True Then
             Exit Function
          End If
       Next
   End If
    
   pTime = Mid(Format(ServerTime, "000000"), 1, 4)
   '譯畢期限:新增和修改
   If txtData(0).Tag = "" Or (txtData(0).Text & txtData(1).Text <> txtData(0).Tag & txtData(1).Tag) Then
      If txtData(0).Text <> "" And Val(txtData(0).Text & txtData(1)) - Val(strSrvDate(2) & pTime) < 200 Then
          MsgBox "譯畢期限不可小於系統時間+2小時!! ", vbCritical
          Exit Function
      End If
   End If
   If Trim(txtData(0).Text & txtData(1).Text) <> "" Then
      Chk1.Value = 1
   Else
      Chk1.Value = 0
   End If
  
   CheckDataValidate = True
End Function

Private Function CheckDiff() As Boolean
    If txtData(0).Text <> txtData(0).Tag Then
       CheckDiff = True
    End If
    If txtData(1).Text <> txtData(1).Tag Then
       CheckDiff = True
    End If
    If m_TCT10 <> Trim(Mid(Combo1.Text, 1, 6)) Then
       CheckDiff = True
    End If
    'Added by Lydia 2023/03/02 外專新案認領
    If lblTCN.Visible = True And Combo2.Visible = True And Combo2.Text <> Combo2.Tag Then
       CheckDiff = True
    End If
    'end 2023/03/02
End Function

Private Function SaveDatabase() As Boolean
Dim bolConn As Boolean

On Error GoTo Err01
  
    '譯畢期限
    strExc(4) = IIf(txtData(0).Text <> "", TransDate(txtData(0).Text, 2), "")
    strExc(5) = IIf(txtData(1).Text <> "", txtData(1).Text, "")
    strExc(6) = ""
    If Combo1.Locked = False Then
       strExc(6) = Trim(Mid(Combo1.Text, 1, 6))
    End If
      
    strSql = ""

    '譯畢期限
    If txtData(0).Text <> txtData(0).Tag Then
       strSql = strSql & ", TCT02=" & CNULL(strExc(4), True)
    End If
    If txtData(1).Text <> txtData(1).Tag Then
       strSql = strSql & ", TCT03=" & CNULL(strExc(5), True)
    End If
    '改期限,發mail通知
    If Trim(txtData(0).Text & txtData(1).Text) <> Trim(txtData(0).Tag & txtData(1).Tag) And Trim(txtData(0).Tag & txtData(1).Tag) <> "" Then
        bolEmail2 = True
    End If
    
    '命名人員
    If strExc(6) <> "" And m_TCT10 <> strExc(6) Then
        '主管->主任or工程師
       If m_UserNo = m_TCT04 Then
            'Modified by Lydia 2022/12/01 判斷ST05的權限
            'strExc(0) = "select nvl(st20,'99') st20 from staff where st01='" & strExc(6) & "' "
            strExc(0) = "select st01,nvl(st20,'99') st20,sr08 from staff,(select sr01,sr02,sr08 from staff_right where sr02='frm090902') " & _
                             "where st01='" & strExc(6) & "' and st05=sr01(+) "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               'Modified by Lydia 2019/03/08 排除46專利代理人
               'If Val(RsTemp(0)) <= 52 Then
               'Modified by Lydia 2022/12/01 判斷ST05的權限; ex. 李宗勳95003的實際作業及組織分配裡，他僅有工程師的權限，而非主任。
               'If Val(RsTemp(0)) <= 52 And Val(RsTemp(0)) <> 46 Then
               'Modified by Lydia 2024/03/20 外專機械設計組人員異動調整程式：排除工程師主管 And InStr(m_GrpMan, strExc(6)) = 0
               If "" & RsTemp.Fields("SR08") = "Y" And InStr(m_GrpMan, strExc(6)) = 0 Then
                  strSql = strSql & ", TCT07=" & CNULL(strExc(6)) & ", TCT08=NULL, TCT09=NULL "
                  If m_TCT07 <> "" Then
                     bolEmail2 = True
                     'Added by Lydia 2020/04/27 若修改前or後是主管本人; (取消四組主管不能分案給自己的限制)
                     If strExc(6) = m_UserNo Then
                        m_Receiver = m_Receiver & ";" & m_TCT07
                        m_ReMailTxt = "原分案主任：" & GetStaffName(m_TCT07) & vbCrLf
                     Else
                     'end 2020/04/27
                        'Added by Lydia 2018/01/03
                        m_Receiver = m_Receiver & ";" & m_TCT07
                        m_ReMailTxt = "原分案主任：" & GetStaffName(m_TCT07) & vbCrLf & "新分案主任：" & GetStaffName(strExc(6))
                        'end 2018/01/03
                     End If 'Added by Lydia 2020/04/27
                  End If
                  '已有命名人員,又改主任
                  If m_TCT10 <> "" Then
                     strSql = strSql & ", TCT10=NULL, TCT11=NULL, TCT12=NULL "
                     'Added by Lydia 2020/04/27 若修改前or後是主管本人; (取消四組主管不能分案給自己的限制)
                     If strExc(6) = m_UserNo Then
                        m_Receiver = m_Receiver & ";" & m_TCT10
                        m_ReMailTxt = m_ReMailTxt & vbCrLf & "原命名人員：" & GetStaffName(m_TCT10) & vbCrLf & "新命名人員：" & GetStaffName(strExc(6))
                     Else
                     'end 2020/04/27
                        'Added by Lydia 2018/01/03
                        m_Receiver = m_Receiver & ";" & m_TCT10
                        If m_ReMailTxt = "" Then m_ReMailTxt = "原命名人員：" & GetStaffName(m_TCT10) & vbCrLf & "新分案主任：" & GetStaffName(strExc(6))
                        'end 2018/01/03
                     End If 'Added by Lydia 2020/04/27
                  End If
               Else
                  'Modified by Lydia 2018/03/26 與Jack討論未分命名人員前,要發逾期通知信; 分命名人員後,再發一次通知信
                  'strSql = strSql & ", TCT10=" & CNULL(strExc(6)) & ", TCT11=NULL, TCT12=NULL "
                  strSql = strSql & ", TCT10=" & CNULL(strExc(6)) & ", TCT11=NULL, TCT12=NULL, TCT115=NULL "
                  'Added by Lydia 2018/01/03 曾分發給主任,後又直接改命名人員
                  If m_TCT07 <> "" And m_TCT07 <> strExc(6) Then
                     strSql = strSql & ", TCT07=NULL , TCT08=NULL, TCT09=NULL "
                      m_Receiver = m_Receiver & ";" & m_TCT07
                      m_ReMailTxt = "原分案主任：" & GetStaffName(m_TCT07) & vbCrLf & "新命名人員：" & GetStaffName(strExc(6))
                  End If
                  'end 2018/01/03
                  If m_TCT10 <> "" Then
                      bolEmail2 = True
                      'Added by Lydia 2018/01/03
                      m_Receiver = m_Receiver & ";" & m_TCT10
                      If m_ReMailTxt = "" Then m_ReMailTxt = "原命名人員：" & GetStaffName(m_TCT10) & vbCrLf & "新命名人員：" & GetStaffName(strExc(6))
                      'end 2018/01/03
                  End If
               End If
               bolEmail = True
            End If
            'Added by Lydia 2024/02/27 外專機械設計組人員異動調整程式：內專工程師協助承辦FCP機械案件，另建虛帳號(員工編號第4碼為9)給內專工程師操作外專歷程
            'Modified by Lydia 2024/03/20 機械組案件101,102由外專工程師(可能是電子組或化學組), 103由內專工程師翻譯
            'If m_UserSt16 = "1" Then
             '  strExc(9) = PUB_GetStaffST16(strExc(6))
             strExc(9) = PUB_GetStaffST16(strExc(6))
             If m_UserSt16 <> strExc(9) Then
               '變更工程師主管: 後續控制比照原模式
               If InStr(m_GrpMan, strExc(6)) > 0 Then
                  strExc(8) = "Update TransCaseTitle Set TCT04='" & strExc(6) & "' ,TCT07=NULL,TCT08=NULL,TCT09=NULL,TCT10=NULL,TCT11=NULL,TCT12=NULL where TCT01='" & m_TCT01 & "' "
                  cnnConnection.Execute strExc(8)
                  bolEmail = True
                  bolEmail2 = False
                  m_TCT04 = strExc(6)
                  strSql = ""  '清除其他變更/通知
               End If
             'end 2024/03/20
               If PUB_GetST03(strExc(6)) = "F21" And strCase(7) <> strExc(9) And strExc(9) <> "" Then
                  strExc(8) = "Update Patent Set PA150='" & strExc(9) & "' where pa01='" & strCase(1) & "' and pa02='" & strCase(2) & "' and pa03='" & strCase(3) & "' and pa04='" & strCase(4) & "' "
                  Pub_SeekTbLog strExc(8)
                  cnnConnection.Execute strExc(8)
               End If
            End If
            'end 2024/02/27
       '主任->自己or工程師
       ElseIf m_UserNo = m_TCT07 Then
             'Modified by Lydia 2018/03/26 與Jack討論未分命名人員前,要發逾期通知信; 分命名人員後,再發一次通知信
             'strSql = strSql & ", TCT10=" & CNULL(strExc(6)) & ", TCT11=NULL, TCT12=NULL "
             strSql = strSql & ", TCT10=" & CNULL(strExc(6)) & ", TCT11=NULL, TCT12=NULL, TCT115=NULL "
             If m_UserNo <> strExc(6) Then
                bolEmail = True
                If m_TCT10 <> "" Then
                   bolEmail2 = True
                    'Added by Lydia 2018/01/03
                    m_Receiver = m_Receiver & ";" & m_TCT10
                    m_ReMailTxt = "原命名人員：" & GetStaffName(m_TCT10) & vbCrLf & "新命名人員：" & GetStaffName(strExc(6))
                    'end 2018/01/03
                End If
             End If
       End If
    End If
    
    'Modified by Lydia 2023/03/02
    'If strSql <> "" Then
    If strSql <> "" Or (lblTCN.Visible = True And Combo2.Visible = True And Combo2.Text <> Combo2.Tag) Then
       bolConn = True
       cnnConnection.BeginTrans
       
       If strSql <> "" Then 'Added by Lydia 2023/03/02
          strSql = "UPDATE TransCaseTitle SET" & Mid(strSql, 2) & " WHERE TCT01='" & m_TCT01 & "' "
          cnnConnection.Execute strSql
       End If 'Added by Lydia 2023/03/02
       'Added by Lydia 2023/03/02 外專新案認領
       If lblTCN.Visible = True And Combo2.Visible = True And Combo2.Text <> Combo2.Tag Then
           '已排除變更命名人員
           strSql = "Update TrackingCaseName Set TCN20='" & Left(Combo2.Text, 1) & "' Where TCN05='" & m_TCT01 & "' "
           cnnConnection.Execute strSql
           If PUB_UpdateTCNstate("2", strCase(1) & strCase(2) & strCase(3) & strCase(4)) = False Then
               GoTo Err01
           End If
       Else
       'end 2023/03/02
          'Added by Lydia 2020/04/27 若修改後是主管本人,直接進入主管確認階段
          If m_UserNo = m_TCT04 And m_UserNo = Trim(Left(Combo1.Text, 6)) Then
               strSql = "update transcasetitle set tct08=to_char(sysdate,'YYYYMMDD'), tct09=substr(to_char(sysdate,'HH24MISS'),1,4), " & _
                           " tct10='" & m_TCT04 & "', tct11=to_char(sysdate,'YYYYMMDD'), tct12=substr(to_char(sysdate,'HH24MISS'),1,4) where TCT01='" & m_TCT01 & "' "
               cnnConnection.Execute strSql
          End If
          'end 2020/04/27
       End If 'Added by Lydia 2023/03/02
       cnnConnection.CommitTrans
    End If
    SaveDatabase = True
    Exit Function
    
Err01:
If Err.Number <> 0 Then
   If bolConn = True Then
       cnnConnection.RollbackTrans
   End If
   MsgBox Err.Description
End If
End Function

Private Sub cmdok_Click(Index As Integer)

  Select Case Index
      Case 0 '存檔
            If CheckDataValidate = False Then Exit Sub
            'Added by Lydia 2023/03/02 外專新案認領
            If lblTCN.Visible = True And Combo2.Visible = True And Combo2.Text <> Combo2.Tag Then
                 If m_TCT10 <> Trim(Mid(Combo1.Text, 1, 6)) Then
                    MsgBox "變更提申急件組別不可同時變更命名人員！", vbCritical + vbOKOnly, "稽核資料"
                    Combo1.SetFocus
                    Exit Sub
                 End If
                 If Val(Left(Combo2, 1)) < 1 Or Val(Left(Combo2, 1)) > 4 Then
                    MsgBox "請輸入提申急件組別！", vbCritical + vbOKOnly, "稽核資料"
                    Combo2.SetFocus
                    Exit Sub
                 End If
            End If
            'end 2023/03/02
            
            'Added by Lydia 2020/04/27 檢查
            If m_UserNo = m_TCT04 And m_UserNo = Trim(Left(Combo1.Text, 6)) Then
                If MsgBox("命名人員為自己，是否繼續存檔？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                    Exit Sub
                End If
            End If
            'end 2020/04/27
            
            'Added by Lydia 2022/05/12 避免重複Click
            cmdok(Index).Enabled = False
            Screen.MousePointer = vbHourglass
            'end 2022/05/12
            If SaveDatabase = True Then
               Screen.MousePointer = vbDefault 'Added by Lydia 2022/05/12
               Call SetCaseTitle(False)
               If Trim(Mid(Combo1.Text, 1, 6)) <> "" And (bolEmail = True Or bolEmail2 = True) Then
                  'Modified by Lydia 2018/01/03 增加原分案人員和郵件內容
                  'If PUB_GetTCTmail(False, 1, strCase(1), strCase(2), strCase(3), strCase(4), m_TCT01, "", Trim(Mid(Combo1.Text, 1, 6)), , IIf(bolEmail2 = True, "修改補發: ", "")) = True Then
                  If PUB_GetTCTmail(False, 1, strCase(1), strCase(2), strCase(3), strCase(4), m_TCT01, "", Trim(Mid(Combo1.Text, 1, 6)) & m_Receiver, , IIf(bolEmail2 = True, "修改補發: ", ""), m_ReMailTxt) = True Then
                    GoTo JumpCloseFrm
                  End If
               ElseIf m_UserNo = m_TCT10 Then
                    GoTo JumpCloseFrm
               End If
            Else
               Screen.MousePointer = vbDefault 'Added by Lydia 2022/05/12
               Exit Sub
            End If
      Case 1 '回前畫面
            If cmdok(0).Enabled = False Then GoTo JumpCloseFrm
            If CheckDiff = True Then
               If MsgBox("你並未存檔，確定離開嗎?", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                   Exit Sub
               Else
                   GoTo JumpCloseFrm
               End If
            Else
               GoTo JumpCloseFrm
            End If
  End Select
  Exit Sub
  
JumpCloseFrm:
  Screen.MousePointer = vbDefault 'Added by Lydia 2022/05/12
  Me.Hide
  Unload Me
End Sub

'Added by Lydia 2024/04/20
Private Sub Combo1_Validate(Cancel As Boolean)
Dim strTmpA As String

   strTmpA = GetPrjSalesNM(Left(Combo1.Text, 5))
   If strTmpA <> "" Then
      Combo1.Text = Left(Combo1.Text, 5) & " " & strTmpA
   Else
      MsgBox "此人員已離職！！", , "人員錯誤！！"
      Cancel = True
      Combo1.Text = ""
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Lydia 2023/03/02
   If TypeName(m_PrevForm) = "frm090902" Then
      m_PrevForm.doQuery False
      m_PrevForm.Show
   End If
   Set frm090902_1 = Nothing
End Sub

Public Sub SetParent(ByRef fm As Form, ByVal pCase As String, ByVal pNo As String, ByVal pUser As String)
   Set m_PrevForm = fm
   m_TCT01 = pNo
   m_UserNo = pUser
   Call ChgCaseNo(Replace(pCase, "-", ""), strCase)
   m_UserSt16 = PUB_GetStaffST16(m_UserNo) 'Added by Lydia 2017/12/29
End Sub

Private Sub ClearForm(Optional ByVal bolRest As Boolean)
Dim oLbl As Control
Dim oTxt As Control

   For Each oLbl In lblData
      oLbl.Caption = ""
      If bolRest = True Then oLbl.BackColor = &H8000000F
   Next
   
   Chk1.Value = 0
   Chk1.Tag = ""
   bolEmail = False
   bolEmail2 = False
   Combo1.Locked = False
   'Adde by Lydia 2018/01/03
   m_Receiver = ""
   m_ReMailTxt = ""
   'end 2018/01/03
   
   For Each oTxt In txtData
      oTxt.Text = ""
      oTxt.Tag = ""
   Next
   
   'Added by Lydia 2018/10/22
   lblCMboth.Caption = ""
   lblCMboth.Tag = ""
   lblCM.Tag = ""
   
   'Added by Lydia 2023/03/02
   Combo2.Clear
   Combo2.Tag = ""
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   Combo1.Clear
   ClearForm True
  
   If ReadData = True Then
      Call SetCombo1
   End If
   
   'Added by Lydia 2017/12/29 檢查專利檔和命名記錄檔的工程師組別
   'Modified by Lydia 2024/02/27 外專機械設計組人員異動調整程式：新案認領組別，請取消機械設計組，案件歸電子組; 由Wilson代機械組主管T1
   'If strCase(7) <> m_UserSt16 Then
   If strCase(7) <> m_UserSt16 And strCase(7) <> "1" And strCase(7) <> "4" Then
        MsgBox "命名記錄檔的工程師組別和專利基本檔不一致，請通知程序人員到新案建檔設定工程師組別! "
        cmdok(0).Enabled = False
        cmdOpen.Enabled = False
        Combo1.Locked = True
   End If
   'end 2017/12/29
   
   m_GrpMan = Pub_GetSt16Man(False) 'Added by Lydia 2024/03/20
   
End Sub

Private Function ReadData() As Boolean
Dim rsRd As New ADODB.Recordset
    
    '改成模組控制,若基本資料顯示有變,要注意frm090902_1,frm090902_2,frm090903_1的欄位
    'Modified by Lydia 2018/04/18 +n_cp118
    If PUB_GetTCTread(Me, strCase, m_TCT27kind, n_CP118) = True Then
       ReadData = True
       'Added by Lydia 2018/10/22 顯示相關案
       If lblCMboth.Tag <> "" Then
           Call ChgCaseNo(lblCMboth.Tag, strExc)
           strExc(0) = "select tct10,st02 from caseprogress,transcasetitle,staff where cp01='" & strExc(1) & "' and cp02='" & strExc(2) & "' and cp03='" & strExc(3) & "' and cp04='" & strExc(4) & "' and cp31='Y' and cp09=tct01(+) and tct10=st01(+) "
           intI = 1
           Set rsRd = ClsLawReadRstMsg(intI, strExc(0))
           If intI = 1 Then
                lblCM.Visible = True
                lblCM.Caption = "命名人員：" & rsRd.Fields("tct10") & " " & rsRd.Fields("st02")
                lblCM.Tag = "" & rsRd.Fields("tct10")
           End If
       Else
           lblCM.Visible = False
       End If
       'end 2018/10/22
       Call SetCaseTitle(True)
       
        'Added by Lydia 2020/01/20 專利案件和English_Vers檔案：判斷檔案上傳目的地，外文本->原始檔區
        If PUB_ChkCPExist(strCase, cntEnglish_Vers, , strExc(1), , "D") = True Then 'English_Vers992
            cmdOpen.Caption = Replace(cmdOpen.Caption, "外文本", "原始檔")
            cmdOpen.Tag = strExc(1)
        End If
        'Mark by Lydia 2020/03/18 以收文為準
        'If strSrvDate(1) >= XY特殊權限啟用日by檔案 Then
        '    cmdOpen.Caption = Replace(cmdOpen.Caption, "外文本", "原始檔")
        'End If
        'end 2020/01/20
        'Added by Lydia 2023/03/02 外專新案認領
        lblTCN.Visible = False: Combo2.Visible = False
        If strSrvDate(1) >= 外專新案認領啟用日 And Val(strCase(9)) = 0 Then
           strExc(0) = "select pa26,cu154,tcn23 from patent, customer,trackingcasename " & _
                             "where pa01='" & strCase(1) & "' and pa02='" & strCase(2) & "' and pa03='" & strCase(3) & "' and pa04='" & strCase(4) & "' and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) " & _
                             "and tcn05='" & m_TCT01 & "' "
           intI = 1
           Set rsRd = ClsLawReadRstMsg(intI, strExc(0))
           If intI = 1 Then
              If "" & rsRd.Fields("cu154") = strCase(7) And Val("" & rsRd.Fields("tcn23")) = 0 Then
                  lblTCN.Visible = True: Combo2.Visible = True
                  Combo2.Clear
                  For intI = 1 To 4
                      Combo2.AddItem intI & "." & PUB_GetFCPGrpName("" & intI)
                  Next intI
                  Combo2.ListIndex = Val(strCase(7)) - 1
                  Combo2.Tag = Combo2.Text
              End If
           End If
        End If
        Set rsRd = Nothing
        'end 2023/03/02
    Else
       MsgBox "查無資料 !", vbExclamation
       Unload Me
    End If
    
End Function

'設案件命名欄位
Private Sub SetCaseTitle(ByVal bolCmb As Boolean)
Dim rsA As New ADODB.Recordset
Dim Str01 As String
Dim intA As Integer

    Str01 = "select A.*,B.ST02 " & _
            "FROM TransCaseTitle A,STAFF B WHERE TCT01='" & m_TCT01 & "' AND TCT10=ST01(+) "
    intA = 1
    Set rsA = ClsLawReadRstMsg(intA, Str01)
    If intA = 1 Then
        With rsA
           '急件
           If Val("" & .Fields("TCT02")) > 0 Then
              Chk1.Value = 1
           End If
           '譯畢期限
           If "" & .Fields("TCT02") <> "" Then
              txtData(0).Text = TransDate(.Fields("TCT02"), 1)
           End If
           txtData(0).Tag = txtData(0).Text
           If "" & .Fields("TCT03") <> "" Then
              'Modified by Lydia 2018/03/05 + format
              txtData(1).Text = Format(.Fields("TCT03"), "0000")
           End If
           txtData(1).Tag = txtData(1).Text
           '工程師主管
           m_TCT04 = "" & .Fields("TCT04")
           '工程師主任
           m_TCT07 = "" & .Fields("TCT07")
           '命名人員
           m_TCT10 = Trim("" & .Fields("TCT10"))
           If m_TCT10 <> "" And bolCmb = True Then
              Combo1.AddItem m_TCT10 & " " & Trim("" & .Fields("ST02"))
              Combo1.ListIndex = 0
              lblData(8).Caption = Trim("" & .Fields("ST02")) 'Added by Lydia 2018/03/16 顯示原命名人員
           End If
           '已回報=>不可變更
           If Val("" & .Fields("TCT11")) > 0 Then
              Combo1.Locked = True
           End If
           'Added by Lydia 2017/12/15 限主管可修改譯畢期限
           If m_UserNo <> m_TCT04 Then
               txtData(0).Locked = True: txtData(1).Locked = True
               Chk1.Enabled = False
           End If
           'end 2017/12/15
        End With
    End If
    Set rsA = Nothing
End Sub

Private Sub SetCombo1()
Dim rsB As New ADODB.Recordset
Dim iA As Integer
     
    strSql = ""
    '分案主管
    If m_UserNo = m_TCT04 Then
        'Added by Lydia 2020/04/27 先加入主管自己; (取消四組主管不能分案給自己的限制)
        strSql = "SELECT 0 ORD1, ST01,ST02,ST16 FROM STAFF WHERE ST04='1' AND ST03='F21' AND ST01 = '" & m_UserNo & "' "
        '先抓主任
        'Modified by Lydia 2019/03/08 排除46專利代理人; ex.A8004被認做主任級
        'strSql = "SELECT 0 ORD1, ST01,ST02,ST16 FROM STAFF WHERE ST04='1' AND ST03='F21' AND ST01 <> '" & m_UserNo & "' AND ST16='" & strCase(7) & "' AND ST20<='52' "
        'Modified by Lydia 2020/04/27 +Union, 並且 ORD1 改為 1
        'strSql = "SELECT 0 ORD1, ST01,ST02,ST16 FROM STAFF WHERE ST04='1' AND ST03='F21' AND ST01 <> '" & m_UserNo & "' AND ST16='" & strCase(7) & "' AND ST20<='52'  AND ST20<>'46' "
        'Modified by Lydia 2024/02/29 改用工程師組別ST16判斷同組人員;strCase(7) >> IIf(m_UserSt16 <> "", m_UserSt16, strCase(7))
        strSql = strSql & "UNION SELECT 1 ORD1, ST01,ST02,ST16 FROM STAFF WHERE ST04='1' AND ST03='F21' AND ST01 <> '" & m_UserNo & "' AND ST16='" & IIf(m_UserSt16 <> "", m_UserSt16, strCase(7)) & "' AND ST20<='52'  AND ST20<>'46' "
        '再抓工程師
        'Modified by Lydia 2019/03/08 +46專利代理人
        'strSql = strSql & "UNION SELECT 1 ORD1, ST01,ST02,ST16 FROM STAFF WHERE ST04='1' AND ST03='F21' AND ST01 <> '" & m_UserNo & "' AND ST16='" & strCase(7) & "' AND NVL(ST20,'99') > '52' " & _
                 "ORDER BY ORD1, ST01 "
        'Modified by Lydia 2020/04/27 ORD1 改為 2
        'Modified by Lydia 2022/10/12 排除F4開頭(虛建編號)
        'Modified by Lydia 2024/02/29 改用工程師組別ST16判斷同組人員;strCase(7) >> IIf(m_UserSt16 <> "", m_UserSt16, strCase(7))
        strSql = strSql & "UNION SELECT 2 ORD1, ST01,ST02,ST16 FROM STAFF WHERE ST04='1' AND ST03='F21' AND ST01 <> '" & m_UserNo & "' AND ST16='" & IIf(m_UserSt16 <> "", m_UserSt16, strCase(7)) & "' AND (NVL(ST20,'99') > '52'  or ST20='46' ) AND ST01 NOT LIKE 'F4%' "
        'Added by Lydia 2024/02/27 外專機械設計組人員異動調整程式：內專工程師協助承辦FCP機械案件，另建虛帳號(員工編號第4碼為9)給內專工程師操作外專歷程
        If m_UserSt16 = "1" Then
           strExc(2) = Pub_GetSpecMan("協助機械組工程師")
           If strExc(2) <> "" Then
              strSql = strSql & "UNION SELECT 3 ORD1, ST01,ST02,ST16 FROM STAFF WHERE ST04='1' AND ST03='F21' AND ST01 <> '" & m_UserNo & "' AND INSTR('" & strExc(2) & "',ST01)>0 "
           End If
           'Added by Lydia 2025/09/03 直接列出機械組工程師；特殊設定改成只存內專工程師(員工編號第4碼為9)
           strSql = strSql & "UNION SELECT 3 ORD1, ST01,ST02,ST16 FROM STAFF WHERE ST04='1' AND ST03='F21' AND ST01 <> '" & m_UserNo & "' and st16='4' "
        End If
        'end 2024/02/27
        strSql = strSql & "ORDER BY ORD1, ST01 "
    ElseIf m_UserNo = m_TCT07 Then '分案主任 (抓ST52)
        strSql = "SELECT 0 ORD1, ST01,ST02,ST16 FROM STAFF WHERE ST04='1' AND ST03='F21' AND ST01='" & m_UserNo & "' "
        'Modified by Lydia 2018/11/05 增加第3階主管(ST53)
        'strSql = strSql & "UNION SELECT 1 ORD1, ST01,ST02,ST16 FROM STAFF WHERE ST04='1' AND ST03='F21' AND ST01 <> '" & m_UserNo & "' AND ST16='" & strCase(7) & "' AND ST52='" & m_UserNo & "' " & _
                 "ORDER BY ORD1, ST01 "
        'Modifeid by Lydia 2022/04/13 增加第4階主管(ST54) ; ex. 日文組分成4層,簡偉倫99037要分給A7029陳宇柔(ST54=99037)
        'Modified by Lydia 2024/02/29 改用工程師組別ST16判斷同組人員;strCase(7) >> IIf(m_UserSt16 <> "", m_UserSt16, strCase(7))
        strSql = strSql & "UNION SELECT 1 ORD1, ST01,ST02,ST16 FROM STAFF WHERE ST04='1' AND ST03='F21' AND ST01 <> '" & m_UserNo & "' AND ST16='" & IIf(m_UserSt16 <> "", m_UserSt16, strCase(7)) & "' AND instr(ST52||','||ST53||','||ST54, '" & m_UserNo & "' ) > 0 "
        'Added by Lydia 2022/10/13 特殊情況之指定職代
        strExc(1) = PUB_GetStateForMan(m_UserNo, "A")
        If strExc(1) <> m_UserNo Then
            'Modified by Lydia 2024/02/29 改用工程師組別ST16判斷同組人員;strCase(7) >> IIf(m_UserSt16 <> "", m_UserSt16, strCase(7))
            strSql = strSql & "UNION SELECT 1 ORD1, ST01,ST02,ST16 FROM STAFF WHERE ST04='1' AND ST03='F21' AND ST01 <> '" & m_UserNo & "' AND ST16='" & IIf(m_UserSt16 <> "", m_UserSt16, strCase(7)) & "' AND instr(ST52||','||ST53||','||ST54, '" & strExc(1) & "' ) > 0 "
        End If
        'end 2022/10/13

        strSql = strSql & "ORDER BY ORD1, ST01 "
    End If
    
    If strSql <> "" Then
       iA = 1
       Set rsB = ClsLawReadRstMsg(iA, strSql)
       If iA = 1 Then
          rsB.MoveFirst
          Do While Not rsB.EOF
             'Modified by Lydia 2017/12/27 排除94099(總經理外專編號)
             'Modified by Lydia 2018/01/05 楊雯芳(99033)屬於兼任,先排除
             If Trim("" & rsB.Fields("ST01")) <> m_TCT10 And Trim("" & rsB.Fields("ST01")) <> "94099" And Trim("" & rsB.Fields("ST01")) <> "99033" Then
                Combo1.AddItem Trim("" & rsB.Fields("ST01")) & " " & Trim("" & rsB.Fields("ST02"))
             End If
             rsB.MoveNext
          Loop
       End If
       If m_TCT10 = "" Then
          Combo1.Text = ""
       End If
       Set rsB = Nothing
    End If
End Sub

Private Sub Txtdata_GotFocus(Index As Integer)
    TextInverse txtData(Index)
End Sub

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0 '譯畢期限日期
          If Trim(txtData(Index).Text) = "" Then Exit Sub
          If CheckIsTaiwanDate(txtData(Index)) = False Then
             GoTo ExceptRun
          End If
          strExc(0) = TransDate(txtData(Index), 2)
          If CompWorkDay(1, strExc(0)) <> strExc(0) Then
             MsgBox "請輸入上班日!"
             GoTo ExceptRun
          End If
          If Chk1.Value = 0 Then Chk1.Value = 1
      Case 1 '譯畢期限時間
          If Trim(txtData(0).Text & txtData(1).Text) = "" Then
             Chk1.Value = 0
             Exit Sub
          End If
          If Len(txtData(Index)) < 4 Or Val(txtData(Index)) < 900 Or Val(txtData(Index)) > 1700 Or Val(txtData(Index)) < 0 Or Val(Right(txtData(Index), 2)) > 60 Then
             MsgBox "請輸入正確時間(ex.0900~1700) ! ", vbExclamation
             GoTo ExceptRun
          End If
          If Chk1.Value = 0 Then Chk1.Value = 1
   End Select
   
   Exit Sub
   
ExceptRun:
   txtData(Index).SetFocus
   Txtdata_GotFocus Index
   Cancel = True
End Sub

'Added by Lydai 2017/12/27 外文本
Private Sub cmdOpen_Click()
Dim hLocalFile As Long 'Added by Lydia 2018/06/21

On Error GoTo ErrHand01 'Added by Lydia 2018/03/23 無權限的錯誤要改訊息

    'Added by Lydia 2020/01/20 開啟[原始檔區]
    If InStr(cmdOpen.Caption, "原始檔") > 0 Then
        If PUB_CheckFormExist("frm100101_M") Then
            MsgBox "請先關閉共同查詢〔原始檔區〕畫面！"
            Exit Sub
        Else
            If cmdOpen.Tag = "" Then
                MsgBox strCase(1) & "-" & strCase(2) & "在〔原始檔區〕的English_Vers收文號不存在!", vbInformation
            Else
                strExc(1) = ""
                frm100101_M.m_strKey = cmdOpen.Tag '多筆總收文號
                frm100101_M.SetParent Me
                If frm100101_M.QueryData = True Then
                   frm100101_M.Show
                   Me.Hide
                End If
            End If
        End If
    Else
    'end 2020/01/20
        'Modified by Lydia 2018/05/09 +系統別
        'Modifiede by Lydia 2021/12/06 (109/4/6)已將\\Typing2的"English_Vers"和"專利案件"的案件資料夾，全部搬到原始檔區
        'strExc(1) = Pub_GetFCPcaseFilePath(strCase(2), , strCase(1))
        'If Dir(strExc(1) & "\*.*") <> "" Then
        '      'Modified by Lydia 2018/06/21 用檔案總管開啟放置1~2分鐘後,檔案總管會出錯(ex. A2037, A4041)
        '      'SHELL "Explorer.exe " & strExc(1), vbNormalFocus  '開啟案件資料夾
        '      ShellExecute hLocalFile, "explore", strExc(1), vbNullString, vbNullString, 1
        'Else
        '     MsgBox lblData(6).Caption & "在" & strExc(1) & "的資料夾不存在或無檔案!", vbInformation
        'End If
        strExc(1) = ""
        'end 2021/12/06
    End If 'Added by Lydia 2020/01/20
    
'Added by Lydia 2018/03/23
    Exit Sub
    
ErrHand01:
    If Err.Number <> 0 Then
         '全部錯誤訊息統一
         MsgBox "無法讀取" & strExc(1) & "，請通知電腦中心！", vbCritical
         Resume Next
    End If
'end 2018/03/23
End Sub

