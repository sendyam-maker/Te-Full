VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060312 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專承辦人請款/發文明細表"
   ClientHeight    =   4068
   ClientLeft      =   1296
   ClientTop       =   2952
   ClientWidth     =   6204
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4068
   ScaleWidth      =   6204
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   4065
      ScaleHeight     =   444
      ScaleWidth      =   648
      TabIndex        =   28
      Top             =   2460
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   14
      Left            =   1050
      MaxLength       =   1
      TabIndex        =   14
      Top             =   3690
      Width           =   240
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   720
      MaxLength       =   1
      TabIndex        =   7
      Top             =   2100
      Width           =   240
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   13
      Left            =   1050
      MaxLength       =   1
      TabIndex        =   13
      Top             =   3360
      Width           =   240
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1050
      TabIndex        =   0
      Top             =   504
      Width           =   3270
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1050
      MaxLength       =   1
      TabIndex        =   1
      Top             =   804
      Width           =   240
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1050
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1128
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2145
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1128
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1050
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1440
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   2145
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1440
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1485
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1764
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1050
      MaxLength       =   9
      TabIndex        =   8
      Top             =   2415
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   2145
      MaxLength       =   99
      TabIndex        =   9
      Top             =   2415
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1050
      MaxLength       =   9
      TabIndex        =   10
      Top             =   2745
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   2145
      MaxLength       =   9
      TabIndex        =   11
      Top             =   2745
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   12
      Left            =   1050
      MaxLength       =   1
      TabIndex        =   12
      Top             =   3060
      Width           =   240
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4410
      TabIndex        =   15
      Top             =   90
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   5205
      TabIndex        =   16
      Top             =   90
      Width           =   800
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Left            =   2550
      TabIndex        =   29
      Top             =   1770
      Width           =   1470
      VariousPropertyBits=   27
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "輸出方式：  　   ( 1.螢幕 2.印表機 )"
      Height          =   180
      Index           =   9
      Left            =   150
      TabIndex        =   27
      Top             =   3720
      Width           =   3090
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "組別：          ( 1.電子電機 2.化學 3.日文 4.機械設計 5.其他 )"
      Height          =   180
      Index           =   7
      Left            =   150
      TabIndex        =   26
      Top             =   2145
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "列印內容： 　　( 1.工程師 2.程序人員 3.所有人 4.P案管制人 5.發文操作人 )"
      Height          =   180
      Index           =   12
      Left            =   150
      TabIndex        =   25
      Top             =   3420
      Width           =   5955
   End
   Begin VB.Line Line4 
      X1              =   1980
      X2              =   2145
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Line Line3 
      X1              =   1980
      X2              =   2145
      Y1              =   2535
      Y2              =   2535
   End
   Begin VB.Line Line2 
      X1              =   1980
      X2              =   2145
      Y1              =   1545
      Y2              =   1545
   End
   Begin VB.Line Line1 
      X1              =   1980
      X2              =   2145
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   24
      Top             =   555
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "列印別：            (1.請款 2.發文)"
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   23
      Top             =   870
      Width           =   2550
   End
   Begin VB.Label Label1 
      Caption         =   "日期："
      Height          =   180
      Index           =   2
      Left            =   150
      TabIndex        =   22
      Top             =   1185
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   3
      Left            =   150
      TabIndex        =   21
      Top             =   1500
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦人/核稿人："
      Height          =   180
      Index           =   4
      Left            =   150
      TabIndex        =   20
      Top             =   1806
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "申請人："
      Height          =   180
      Index           =   5
      Left            =   150
      TabIndex        =   19
      Top             =   2460
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "代理人："
      Height          =   180
      Index           =   6
      Left            =   150
      TabIndex        =   18
      Top             =   2790
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "報表別：             ( 1.明細：排除限閱案件　 2.統計 )"
      Height          =   180
      Index           =   8
      Left            =   156
      TabIndex        =   17
      Top             =   3120
      Width           =   4104
   End
End
Attribute VB_Name = "frm060312"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; lbl1 ; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
'Modify by Morgan 2007/7/25 改為承辦人請款/發文明細表
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, strTemp3(0 To 9) As String, SavDay1 As String, SavDay2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String, StrSQL6 As String
Dim strSQL2 As String, iPrint As Long, Page As Integer, strTemp(0 To 10) As String
Dim PLeft(0 To 12) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrTemp4(0 To 4) As String, StrTemp5(0 To 4) As String
Dim strST05 As String
'Add By Cheng 2002/09/17
Dim blnClkSure As Boolean
'Add By Cheng 2003/03/05
Dim dblPoint As Double '點數小計
Dim dblTotPoint As Double '點數總計
'Add by Morgan 2004/1/2 '案件性質與承辦人過濾條件
Dim strCon1 As String
Dim stName As String, stGrp As String, stDep As String, stID As String
Dim dblHour As Double, dblTotHour As Double
Dim m_RptType As Integer, m_GrpType As Integer
Dim m_bPrinter As Boolean, m_iPages As Integer, m_Device
Dim m_strSharePointVTB As String '分配點數語法
'Add by Morgan 2010/11/15
Dim dblPoint2 As Double '核稿點數
Dim dblTotPoint2 As Double '核稿點數總計
Private Const DefSysTxt = "FCP,FG,P,PS,CFP,CPS,FCL,CFL,L" 'Modified by Lydia 2015/01/22
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件
Dim m_AttachPath As String 'Added by Morgan 2021/5/21
Dim stDefCP14 As String 'Added by Lydia 2024/03/05 預設承辦人
Dim strST16 As String, strST70 As String 'Added by Lydia 2025/02/06

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If Len(txt1(0)) = 0 Then
             s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
             txt1(0).SetFocus
             Exit Sub
         Else
            If Len(txt1(1)) = 0 Then
                s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
                txt1(1).SetFocus
                Exit Sub
            Else
               If Len(txt1(3)) = 0 Then
                  s = MsgBox("日期區間不可空白!!", , "USER 輸入錯誤")
                  txt1(2).SetFocus
                  txt1_GotFocus (2)
                  Exit Sub
               Else
                  If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
                     Me.txt1(2).SetFocus
                     txt1_GotFocus 2
                     Exit Sub
                  End If
                  If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
                     Me.txt1(3).SetFocus
                     txt1_GotFocus 3
                     Exit Sub
                  End If
               
                  If Me.txt1(2).Text <> "" And Me.txt1(3).Text <> "" Then
                     If Val(Me.txt1(2).Text) > Val(Me.txt1(3).Text) Then
                        MsgBox "日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                        blnClkSure = True
                        Me.txt1(2).SetFocus
                        txt1_GotFocus 2
                     End If
                  End If
                  If Me.txt1(4).Text <> "" And Me.txt1(5).Text <> "" Then
                     If Me.txt1(4).Text > Me.txt1(5).Text Then
                        MsgBox "案件性質範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                        blnClkSure = True
                        Me.txt1(4).SetFocus
                        txt1_GotFocus 4
                        Exit Sub
                     End If
                  End If
                  'Added by Lydia 2024/03/05 Debug:蔡政村在113/2/5輸入承辦人=各組主管直接按Enter跳過txt1_validate
                  If stDefCP14 <> "" Then
                     txt1(6).Tag = stDefCP14
                     Call txt1_Validate(6, blnClkSure)
                     If blnClkSure = True Then
                        Me.txt1(6).SetFocus
                        txt1_GotFocus 6
                        Exit Sub
                     End If
                  End If
                  'end 2024/03/05
                  If Me.txt1(6).Text <> "" Then
                     '2008/1/8 MODIFY BY SONIA 不檢查離職
                     'If ClsPDGetStaff(txt1(6), strExc(0)) Then
                     '   lbl1 = strExc(0)
                     'Else
                     '   lbl1 = ""
                     '   Me.txt1(6).SetFocus
                     '   txt1_GotFocus 6
                     '   Exit Sub
                     'End If
                     lbl1 = GetPrjSales(txt1(6), "智權人員")
                     '2008/1/8 END
                  'Added by Lydia 2025/02/06
                  'Mark by Lydia 2025/02/25 調整承辦人=空白，查詢同組別+ST52~ST55的下屬工程師
                  'Else
                  '   '外專/日專工程師中級主管(主任)可查詢底下所有人員的資料: 未輸入承辦人/核稿人,預設工程師組別
                  '   If strST05 = "39" Then
                  '      If txt1(11) <> strST16 Then
                  '         txt1(11) = strST16
                  '      End If
                  '   End If
                  'end 2025/02/06
                  'end 2025/02/25
                  End If
                  
                  If Mid(txt1(7), 1, 6) <> Mid(txt1(8), 1, 6) Then
                     s = MsgBox("申請人前六碼必須相同", , "USER 輸入錯誤")
                     txt1(7).SetFocus
                     txt1(7).SelStart = 0
                     txt1(7).SelLength = Len(txt1(7))
                     Exit Sub
                  End If
                  If Me.txt1(7).Text <> "" And Me.txt1(8).Text <> "" Then
                     If Me.txt1(7).Text > Me.txt1(8).Text Then
                        MsgBox "申請人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                        blnClkSure = True
                        Me.txt1(7).SetFocus
                        txt1_GotFocus 7
                        Exit Sub
                     End If
                  End If
                  If Mid(txt1(9), 1, 6) <> Mid(txt1(10), 1, 6) Then
                     s = MsgBox("代理人前六碼必須相同", , "USER 輸入錯誤")
                     txt1(9).SetFocus
                     txt1(9).SelStart = 0
                     txt1(9).SelLength = Len(txt1(9))
                     Exit Sub
                  End If
                  
                  If Me.txt1(9).Text <> "" And Me.txt1(10).Text <> "" Then
                     If Me.txt1(9).Text > Me.txt1(10).Text Then
                        MsgBox "代理人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                        blnClkSure = True
                        Me.txt1(9).SetFocus
                        txt1_GotFocus 9
                        Exit Sub
                     End If
                  End If
                  '檢查報表別
                  If Me.txt1(12).Text = "" Then
                     MsgBox "請輸入報表別!!!", vbExclamation + vbOKOnly
                     Me.txt1(12).SetFocus
                     txt1_GotFocus 12
                     Exit Sub
                  End If
                  
                  '檢查列印內容
                  If Me.txt1(13).Text = "" Then
                     MsgBox "請輸入列印內容!!!", vbExclamation + vbOKOnly
                     Me.txt1(13).SetFocus
                     txt1_GotFocus 13
                     Exit Sub
                  End If
                  
                  If Me.txt1(14).Text = "" Then
                     MsgBox "請選擇輸出方式!!!", vbExclamation + vbOKOnly
                     Me.txt1(14).SetFocus
                     txt1_GotFocus 14
                     Exit Sub
                  End If
                  'Modified by Lydia 2015/01/22
                    If Me.txt1(13) = "4" Then
                       If Me.txt1(0) <> "P" Then
                          MsgBox "系統別超出範圍!", vbCritical
                          Me.txt1(0).SetFocus
                          Exit Sub
                       ElseIf txt1(1) = "1" Then
                          MsgBox "P案管制人只限發文別!", vbCritical
                          Me.txt1(1).SetFocus
                          Exit Sub
                       End If
                    'Add By Sindy 2015/9/22
                    ElseIf Me.txt1(13) = "5" Then '發文操作人
                       If txt1(1) = "1" Then
                          MsgBox "發文操作人只限發文別!", vbCritical
                          Me.txt1(1).SetFocus
                          Exit Sub
                       End If
                    '2015/9/22 END
                    End If
                    
                  Screen.MousePointer = vbHourglass
                  Me.Enabled = False
                  'Modify by Morgan 2010/5/20 改抓分配點數資料
                  'Process
                  'Modified by Lydia 2015/01/23 增加P案管制人,改變SQL
                '  ProcessNew
                  ProcessNew2
                  Me.Enabled = True
                  Screen.MousePointer = vbDefault
               End If
            End If
         End If
      Case 1
         Unload Me
      Case Else
   End Select
End Sub

'Private Sub Process()
'   Dim stVTB As String
'   Dim stConPA As String, stConSP As String, stConLC As String
'   Dim stCon, stCon1K0 As String, stCon0K0 As String, stConCP As String
'   Dim stCon0 As String, stCon1 As String, stCon2 As String, stCon3 As String
'   Dim stSys As String, ii As Integer
'   Dim rsQuery As New ADODB.Recordset
'
'   stCon0 = "": stCon1 = "": stCon2 = "": stCon3 = ""
'   stCon1K0 = "": stCon0K0 = "": stConCP = ""
'
'   stSys = "'" & Join(Split(txt1(0), ","), "','") & "'"
'   If txt1(1) = "1" Then
'      stCon1K0 = stCon1K0 & " and a1k13||'' in (" & stSys & ")"
'   End If
'   stCon = stCon & " and cp01 in (" & stSys & ")"
'
'   '日期
'   If txt1(2) <> "" Then
'      '請款
'      If txt1(1) = "1" Then
'         stCon1K0 = stCon1K0 & " and a1k02>=" & txt1(2)
'         stCon0K0 = stCon0K0 & " and a0k02>=" & txt1(2)
'         stConCP = stConCP & " and CP27>=" & DBDATE(txt1(2))
'      '發文
'      Else
'         stCon = stCon & " and cp27>=" & DBDATE(txt1(2))
'      End If
'   End If
'   If txt1(3) <> "" Then
'      '請款
'      If txt1(1) = "1" Then
'         stCon1K0 = stCon1K0 & " and a1k02<=" & txt1(3)
'         stCon0K0 = stCon0K0 & " and a0k02<=" & txt1(3)
'         stConCP = stConCP & " and CP27<=" & DBDATE(txt1(3))
'      '發文
'      Else
'         stCon = stCon & " and cp27<=" & DBDATE(txt1(3))
'      End If
'   End If
'   '案件性質
'   If txt1(4) <> "" Then
'      stCon = stCon & " and cp10>='" & txt1(4) & "'"
'   End If
'   If txt1(5) <> "" Then
'      stCon = stCon & " and cp10<='" & txt1(5) & "'"
'   End If
'   '承辦人
'   If txt1(6) <> "" Then
'      '一般案件性質
'      stCon0 = stCon0 & " and cp14='" & txt1(6) & "'"
'      '翻譯
'      strExc(0) = "select sim01,sim02 from staff_idmap where '" & txt1(6) & "' in (sim01,sim02)"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         stCon1 = stCon1 & " and cp14 in ('" & RsTemp.Fields(0) & "','" & RsTemp.Fields(1) & "')"
'      Else
'         stCon1 = stCon1 & " and cp14='" & txt1(6) & "'"
'      End If
'      '核稿
'      stCon2 = stCon2 & " and ep04='" & txt1(6) & "'"
'   End If
'
'   '申請人
'   If Len(Trim(txt1(7))) <> 0 And Len(Trim(txt1(8))) <> 0 Then
'       stConPA = stConPA & " AND ((PA26>='" & GetNewFagent(txt1(7)) & "' AND PA26<='" & GetNewFagent(txt1(8)) & "') OR (PA27>='" & GetNewFagent(txt1(7)) & "' AND PA27<='" & GetNewFagent(txt1(8)) & "') OR (PA28>='" & GetNewFagent(txt1(7)) & "' AND PA28<='" & GetNewFagent(txt1(8)) & "') OR (PA29>='" & GetNewFagent(txt1(7)) & "' AND PA29<='" & GetNewFagent(txt1(8)) & "') OR (PA30>='" & GetNewFagent(txt1(7)) & "' AND PA30<='" & GetNewFagent(txt1(8)) & "')) "
'       stConSP = stConSP & " AND ((SP08>='" & GetNewFagent(txt1(7)) & "' AND SP08<='" & GetNewFagent(txt1(8)) & "') OR (SP58<='" & GetNewFagent(txt1(7)) & "' AND SP58<='" & GetNewFagent(txt1(8)) & "') OR (SP59>='" & GetNewFagent(txt1(7)) & "' AND SP59<='" & GetNewFagent(txt1(8)) & "')) "
'       stConLC = stConLC & " AND ((LC11>='" & GetNewFagent(txt1(7)) & "' AND LC11<='" & GetNewFagent(txt1(8)) & "')) "
'   ElseIf Len(Trim(txt1(7))) <> 0 Then
'       stConPA = stConPA & " AND (PA26>='" & GetNewFagent(txt1(7)) & "' OR PA27>='" & GetNewFagent(txt1(7)) & "' OR PA28>='" & GetNewFagent(txt1(7)) & "' OR PA29>='" & GetNewFagent(txt1(7)) & "' OR PA30>='" & GetNewFagent(txt1(7)) & "') "
'       stConSP = stConSP & " AND NOT (SP08>='" & GetNewFagent(txt1(7)) & "' OR SP58>='" & GetNewFagent(txt1(7)) & "' OR SP59>='" & GetNewFagent(txt1(7)) & "') "
'       stConLC = stConLC & " AND ((LC11>='" & GetNewFagent(txt1(7)) & "')) "
'   ElseIf Len(Trim(txt1(8))) <> 0 Then
'       stConPA = stConPA & " AND NOT (PA26<='" & GetNewFagent(txt1(8)) & "' OR PA27<='" & GetNewFagent(txt1(8)) & "' OR PA28<='" & GetNewFagent(txt1(8)) & "' OR PA29<='" & GetNewFagent(txt1(8)) & "' OR PA30<='" & GetNewFagent(txt1(8)) & "') "
'       stConSP = stConSP & " AND NOT (SP08<='" & GetNewFagent(txt1(8)) & "' OR SP58<='" & GetNewFagent(txt1(8)) & "' OR SP59<='" & GetNewFagent(txt1(8)) & "') "
'       stConLC = stConLC & " AND ((LC11<='" & GetNewFagent(txt1(8)) & "')) "
'   End If
'
'   '代理人
'   If Len(Trim(txt1(9))) <> 0 And Len(Trim(txt1(10))) <> 0 Then
'       stConPA = stConPA & " AND PA75>='" & GetNewFagent(txt1(9)) & "' AND PA75<='" & GetNewFagent(txt1(10)) & "' "
'       stConSP = stConSP & " AND SP26>='" & GetNewFagent(txt1(9)) & "' AND SP26<='" & GetNewFagent(txt1(10)) & "' "
'       stConLC = stConLC & " AND LC22>='" & GetNewFagent(txt1(9)) & "' AND LC22<='" & GetNewFagent(txt1(10)) & "' "
'   ElseIf Len(Trim(txt1(9))) <> 0 Then
'       stConPA = stConPA & " AND PA75>='" & GetNewFagent(txt1(9)) & "' "
'       stConSP = stConSP & " AND SP26>='" & GetNewFagent(txt1(9)) & "' "
'       stConLC = stConLC & " AND LC22>='" & GetNewFagent(txt1(9)) & "' "
'   ElseIf Len(Trim(txt1(10))) <> 0 Then
'       stConPA = stConPA & " AND PA75<='" & GetNewFagent(txt1(10)) & "' "
'       stConSP = stConSP & " AND SP26<='" & GetNewFagent(txt1(10)) & "' "
'       stConLC = stConLC & " AND LC22<='" & GetNewFagent(txt1(10)) & "' "
'   End If
'
'   '工程師
'   If txt1(13) = "1" Then
'      stCon3 = stCon3 & " AND ST15='F21'"
'   'Add by Morgan 2007/11/26
'   '程序
'   ElseIf txt1(13) = "2" Then
'      stCon3 = stCon3 & " AND ST15='F22'"
'   '工程師+程序
'   Else
'      stCon3 = stCon3 & " AND ST15 IN ('F21','F22')"  '2008/4/8 加F81
'   End If
'
'   '組別
'   If txt1(11) <> "" Then
'       stCon3 = stCon3 & " AND ST16='" & txt1(11) & "'"
'   End If
'
'   '中間程序(含開收據的製作中說)
'   '請款日(若發文前請款案件則以發文日為日期條件)
'   If txt1(1) = "1" Then
'      '請款單:承辦點數=對客戶請款金額(扣除折扣金額),發明新型的檢視中說=50%
'      stVTB = "select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",(a1l05-a1l07)/1000*(decode(cp10,'209',0.5,1)) C9" & _
'         " from acc1k0,caseprogress c1,acc1l0 where a1k13||''<>'FCL' and a1k12 is null" & stCon1K0 & _
'         " and cp60(+)=a1k01 and cp27<=a1k02+19110000 and cp14 is not null and cp10<>'201' and cp10<>'210'" & stCon & stCon0 & _
'         " and a1l01(+)=cp60 and a1l04(+)=cp10"
'
'      'Add by Morgan 2008/9/8 FCL請款因無法區分工程師點數故以整張計算
'      'Modify by Morgan 2009/10/30 規費要剔除
'      'stVTB = stVTB & " union all select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",sum((a1l05-a1l07)/1000*(decode(cp10,'209',0.5,1))) C9" & _
'         " from acc1k0,caseprogress c1,acc1l0 where a1k13='FCL' and a1k12 is null" & stCon1K0 & _
'         " and cp60(+)=a1k01 and cp27<=a1k02+19110000 and cp14 is not null " & stCon & stCon0 & _
'         " and a1l01(+)=cp60 group by a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113"
'
'      stVTB = stVTB & " union all select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",sum(DECODE(substr(a1l04,length(a1l04)-1),'99',0,(a1l05-a1l07)/1000)) C9" & _
'         " from acc1k0,caseprogress c1,acc1l0 where a1k13||''='FCL' and a1k12 is null" & stCon1K0 & _
'         " and cp60(+)=a1k01 and cp27<=a1k02+19110000 and cp14 is not null " & stCon & stCon0 & _
'         " and a1l01(+)=cp60 group by a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113"
'      'end 2009/10/30
'
'      '請款單:先請款後發文案件
'      stVTB = stVTB & " union all select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",(a1l05-a1l07)/1000*(decode(cp10,'209',0.5,1)) C9" & _
'         " from caseprogress c1,acc1k0,acc1l0 where cp01<>'FCL' and cp60>'X'" & stConCP & _
'         " and cp27>0 and cp14 is not null and cp10<>'201' and cp10<>'210'" & _
'         " and a1k01(+)=cp60 and a1k02+19110000<cp27 and a1k12 is null" & stCon & stCon0 & _
'         " and a1l01(+)=cp60 and a1l04(+)=cp10"
'
'      'Add by Morgan 2008/9/8 FCL請款因無法區分工程師點數故以整張計算
'      'Modify by Morgan 2009/10/30 規費要剔除
'      'stVTB = stVTB & " union all select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",sum((a1l05-a1l07)/1000*(decode(cp10,'209',0.5,1))) C9" & _
'         " from caseprogress c1,acc1k0,acc1l0 where cp01='FCL' and cp60>'X'" & stConCP & _
'         " and cp27>0 and cp14 is not null" & _
'         " and a1k01(+)=cp60 and a1k02+19110000<cp27 and a1k12 is null" & stCon & stCon0 & _
'         " and a1l01(+)=cp60 group by a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113"
'
'      stVTB = stVTB & " union all select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",sum(DECODE(substr(a1l04,length(a1l04)-1),'99',0,(a1l05-a1l07)/1000)) C9" & _
'         " from caseprogress c1,acc1k0,acc1l0 where cp01='FCL' and cp60>'X'" & stConCP & _
'         " and cp27>0 and cp14 is not null" & _
'         " and a1k01(+)=cp60 and a1k02+19110000<cp27 and a1k12 is null" & stCon & stCon0 & _
'         " and a1l01(+)=cp60 group by a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113"
'      'end 2009/10/30
'
'      '收據:承辦點數=業務收文點數
'      stVTB = stVTB & " union all select a0k02 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",cp18 C9" & _
'         " from acc0k0,caseprogress c1 where nvl(a0k09,0)=0" & stCon0K0 & _
'         " and cp60(+)=a0k01 and cp27<=a0k02+19110000 and cp14 is not null and cp10<>'201'" & stCon & stCon0
'
'      '收據:先請款後發文案件
'      stVTB = stVTB & " union all select a0k02 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113,cp18 C9" & _
'         " from caseprogress c1,acc0k0 where cp60<'X'" & stConCP & _
'         " and cp27>0 and cp14 is not null and cp10<>'201'" & _
'         " and a0k01(+)=cp60 and a0k02+19110000<cp27 and nvl(a0k09,0)=0" & stCon & stCon0
'
'   '發文日
'   Else
'      stVTB = "select cp27 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp05 as cp27,cp60,cp64,cp113" & _
'         ",decode(substr(cp60,1,1),'E',cp18,'X',(a1l05-a1l07)/1000*(decode(cp10,'209',0.5,1))) C9" & _
'         " from caseprogress c1,acc1l0 where cp01<>'FCL' and cp14 is not null and cp10<>'201' and cp10<>'210'" & stCon & stCon0 & _
'         " and a1l01(+)=cp60 and a1l04(+)=cp10"
'
'      'Add by Morgan 2008/9/8 FCL請款因無法區分工程師點數故以整張計算
'      stVTB = stVTB & " union all select cp27 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp05 as cp27,cp60,cp64,cp113" & _
'         ",sum(decode(substr(cp60,1,1),'E',cp18,'X',DECODE(substr(a1l04,length(a1l04)-1),'99',0,(a1l05-a1l07)/1000))) C9" & _
'         " from caseprogress c1,acc1l0 where cp01='FCL' and cp14 is not null" & stCon & stCon0 & _
'         " and a1l01(+)=cp60 group by cp27,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp05,cp60,cp64,cp113"
'   End If
'
'   strExc(1) = "select sqldatet(a1k02) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
'      ",substrb(pa05,1,16) C3,substrb(ptm03,1,6) C4,substrb(cpm03,1,8) C5" & _
'      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
'      ",cp113 C8,decode(ST15,'F22',0,C9) C9,substrb(cp64,1,16) C10" & _
'      ",st01 C11,A0902 C12,ST16 C13,st02 C14" & _
'      " From (" & stVTB & ") X,patent, patenttrademarkmap, casepropertymap,staff,acc090" & _
'      " where  pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA & _
'      " and ptm01(+)=1 and ptm02(+)=pa08 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
'      " and st01(+)=cp14 and a0901(+)=ST15" & stCon3
'
'   strExc(1) = strExc(1) & " union all select sqldatet(a1k02) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
'      ",substrb(sp05,1,16) C3,'' C4,substrb(cpm03,1,8) C5" & _
'      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
'      ",cp113 C8,decode(ST15,'F22',0,C9) C9,substrb(cp64,1,16) C10" & _
'      ",st01 C11,A0902 C12,ST16 C13,st02 C14" & _
'      " From (" & stVTB & ") X,servicepractice,casepropertymap,staff,acc090" & _
'      " where  sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP & _
'      " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
'      " and st01(+)=cp14 and a0901(+)=ST15" & stCon3
'
'   'Add by Morgan 2008/9/8 FCL請款因無法區分工程師點數故以整張計算
'   strExc(1) = strExc(1) & " union all select sqldatet(a1k02) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
'      ",substrb(LC05,1,16) C3,'' C4,substrb(cpm03,1,8) C5" & _
'      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
'      ",cp113 C8,decode(ST15,'F22',0,C9) C9,substrb(cp64,1,16) C10" & _
'      ",st01 C11,A0902 C12,ST16 C13,st02 C14" & _
'      " From (" & stVTB & ") X,lawcase,casepropertymap,staff,acc090" & _
'      " where  lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04 and lc01 is not null" & stConLC & _
'      " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
'      " and st01(+)=cp14 and a0901(+)=ST15" & stCon3
'
'   '新案翻譯(內翻的外譯編號要轉成所內員工編號)
'   '請款日
'   'Modify by Morgan 2009/5/8 核稿人<>承辦人時*0.7
'   If txt1(1) = "1" Then
'      '請款單:內翻翻譯承辦點數=對客戶之翻譯請款金額全額,否則=0
'      stVTB = "select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,nvl(sim01,cp14) as cp14" & _
'         ",cp18,cp27,cp60,cp64,cp113,decode(substrb(cp14,1,1),'F',0,(a1l05-a1l07)/1000*decode(cp14,ep04,1,0.7)) C9" & _
'         " from acc1k0,caseprogress c1,acc1l0,staff_idmap,engineerprogress where a1k12 is null" & stCon1K0 & _
'         " and cp60(+)=a1k01 and cp27<=a1k02+19110000 and cp14 is not null and cp10='201'" & stCon & stCon1 & _
'         " and a1l01(+)=cp60 and a1l04(+)=cp10" & _
'         " and sim02(+)=cp14 and ep02(+)=cp09"
'
'      '請款單:先請款後發文案件
'      stVTB = stVTB & " union all select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,nvl(sim01,cp14) as cp14" & _
'         ",cp18,cp27,cp60,cp64,cp113,decode(substrb(cp14,1,1),'F',0,(a1l05-a1l07)/1000*decode(cp14,ep04,1,0.7)) C9" & _
'         " from caseprogress c1,acc1k0,acc1l0,staff_idmap,engineerprogress" & _
'         " where cp60>'X'" & stConCP & _
'         " and cp27>0 and cp14 is not null and cp10='201'" & _
'         " and a1k01(+)=cp60 and a1k02+19110000<cp27 and a1k12 is null" & stCon & stCon1 & _
'         " and a1l01(+)=cp60 and a1l04(+)=cp10" & _
'         " and sim02(+)=cp14 and ep02(+)=cp09"
'
'      '收據:內翻翻譯承辦點數=業務收文點數,否則=0
'      stVTB = stVTB & " union all select a0k02 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,nvl(sim01,cp14) as cp14" & _
'         ",cp18,cp27,cp60,cp64,cp113,decode(substrb(cp14,1,1),'F',0,cp18*decode(cp14,ep04,1,0.7)) C9" & _
'         " from acc0k0,caseprogress c1,staff_idmap,engineerprogress where nvl(a0k09,0)=0" & stCon0K0 & _
'         " and cp60(+)=a0k01 and cp27<=a0k02+19110000 and cp14 is not null and cp10='201'" & stCon & stCon1 & _
'         " and sim02(+)=cp14 and ep02(+)=cp09"
'
'      '收據:先請款後發文案件
'      stVTB = stVTB & " union all select a0k02 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,nvl(sim01,cp14) as cp14" & _
'         ",cp18,cp27,cp60,cp64,cp113,decode(substrb(cp14,1,1),'F',0,cp18*decode(cp14,ep04,1,0.7)) C9" & _
'         " from caseprogress c1,acc0k0,staff_idmap,engineerprogress" & _
'         " where cp60<'X'" & stConCP & _
'         " and cp27>0 and cp14 is not null and cp10='201'" & _
'         " and a0k01(+)=cp60 and a0k02+19110000<cp27 and nvl(a0k09,0)=0" & stCon & stCon1 & _
'         " and sim02(+)=cp14 and ep02(+)=cp09"
'
'   '發文日
'   Else
'      stVTB = "select cp27 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,nvl(sim01,cp14) as cp14,cp18,cp05 as cp27,cp60,cp64,cp113" & _
'         ",decode(substrb(cp14,1,1),'F',0,decode(substr(cp60,1,1),'E',cp18,'X',(a1l05-a1l07)/1000*decode(cp14,ep04,1,0.7))) C9" & _
'         " from caseprogress c1,acc1l0,staff_idmap,engineerprogress where cp14 is not null and cp10='201'" & stCon & stCon1 & _
'         " and a1l01(+)=cp60 and a1l04(+)=cp10" & _
'         " and sim02(+)=cp14 and ep02(+)=cp09"
'
'   End If
'
'   strExc(2) = " union all select sqldatet(a1k02) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
'      ",substrb(pa05,1,16) C3,substrb(ptm03,1,6) C4,substrb(cpm03,1,8) C5" & _
'      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
'      ",cp113 C8,C9,substrb(cp64,1,16) C10,st01 C11,A0902 C12,ST16 C13,st02 C14" & _
'      " From (" & stVTB & ") X,patent, patenttrademarkmap, casepropertymap,staff,acc090" & _
'      " where pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA & _
'      " and ptm01(+)=1 and ptm02(+)=pa08 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
'      " and st01(+)=cp14 and a0901(+)=ST15" & stCon3
'
'   strExc(2) = strExc(2) & " union all select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
'      ",substrb(sp05,1,16) C3,'' C4,substrb(cpm03,1,8) C5" & _
'      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
'      ",cp113 C8,C9,substrb(cp64,1,16) C10" & _
'      ",st01 C11,A0902 C12,ST16 C13,st02 C14" & _
'      " From (" & stVTB & ") X,servicepractice,casepropertymap,staff,acc090" & _
'      " where  sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP & _
'      " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
'      " and st01(+)=cp14 and a0901(+)=ST15" & stCon3
'
'   '核稿
'   '請款日
'   If txt1(1) = "1" Then
'      '請款單:核稿點數=翻譯請款點數(扣除折扣)x(30%+因翻譯瑕疵扣減支付外譯人員翻譯費之百分比x70%)
'      stVTB = "select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,ep04,cp18,cp27,cp60,cp64,cp114" & _
'         ",(a1l05-a1l07)/1000*(0.3+0.7*(1-nvl(TF06,100)/100)) C9" & _
'         " from acc1k0,caseprogress c1,engineerprogress,acc1l0,transfee where a1k12 is null" & stCon1K0 & _
'         " and cp60(+)=a1k01 and cp27<=a1k02+19110000 and cp14 is not null and cp10='201'" & stCon & stCon2 & _
'         " and ep02(+)=cp09 and ep04<>cp14" & _
'         " and a1l01(+)=cp60 and a1l04(+)=cp10 and tf01(+)=cp09"
'
'      '請款單:先請款後發文案件
'      stVTB = stVTB & " union all select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,ep04,cp18,cp27,cp60,cp64,cp114" & _
'         ",(a1l05-a1l07)/1000*(0.3+0.7*(1-nvl(TF06,100)/100)) C9" & _
'         " from caseprogress c1,acc1k0,engineerprogress,acc1l0,transfee" & _
'         " where cp60>'X'" & stConCP & _
'         " and cp27>0 and cp14 is not null and cp10='201'" & _
'         " and a1k01(+)=cp60 and a1k02+19110000<cp27 and a1k12 is null" & stCon & stCon2 & _
'         " and ep02(+)=cp09 and ep04<>cp14" & _
'         " and a1l01(+)=cp60 and a1l04(+)=cp10 and tf01(+)=cp09"
'
'      '收據:核稿點數=業務收文點數x(30%+因翻譯瑕疵扣減支付外譯人員翻譯費之百分比x70%)
'      stVTB = stVTB & " union all select a0k02 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,ep04,cp18,cp27,cp60,cp64,cp114" & _
'         ",cp18*(0.3+0.7*(1-nvl(TF06,100)/100)) C9" & _
'         " from acc0k0,caseprogress c1,engineerprogress,transfee where nvl(a0k09,0)=0" & stCon0K0 & _
'         " and cp60(+)=a0k01 and cp27<=a0k02+19110000 and cp14 is not null and cp10='201'" & stCon & stCon2 & _
'         " and ep02(+)=cp09 and ep04<>cp14 and tf01(+)=cp09"
'
'      '收據:先請款後分案案件
'      stVTB = stVTB & " union all select a0k02 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,ep04,cp18,cp27,cp60,cp64,cp114" & _
'         ",cp18*(0.3+0.7*(1-nvl(TF06,100)/100)) C9" & _
'         " from caseprogress c1,acc0k0,engineerprogress,transfee" & _
'         " where cp60<'X'" & stConCP & _
'         " and cp27>0 and cp14 is not null and cp10='201'" & _
'         " and a0k01(+)=cp60 and a0k02+19110000<cp27 and nvl(a0k09,0)=0" & stCon & stCon2 & _
'         " and ep02(+)=cp09 and ep04<>cp14 and tf01(+)=cp09"
'
'   '發文日
'   Else
'      stVTB = "select cp27 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,ep04,cp18,cp05 as cp27,cp60,cp64,cp114" & _
'         ",decode(substr(cp60,1,1),'E',cp18,'X',(a1l05-a1l07)/1000)*(0.3+0.7*(1-nvl(TF06,100)/100)) C9" & _
'         " from caseprogress c1,engineerprogress,acc1l0,transfee where cp14 is not null and cp10='201'" & stCon & stCon2 & _
'         " and ep02(+)=cp09 and ep04<>cp14" & _
'         " and a1l01(+)=cp60 and a1l04(+)=cp10 and tf01(+)=cp09"
'
'   End If
'
'   strExc(3) = " union all select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
'      ",substrb(pa05,1,16) C3,substrb(ptm03,1,6) C4,substrb(cpm03,1,8)||'-核稿' C5" & _
'      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
'      ",cp114 C8,C9,substrb(cp64,1,16) C10" & _
'      ",st01 C11,A0902 C12,ST16 C13,st02 C14" & _
'      " From (" & stVTB & ") X,patent,patenttrademarkmap,casepropertymap,staff,acc090" & _
'      " where pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA & _
'      " and ptm01(+)=1 and ptm02(+)=pa08 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
'      " and st01(+)=ep04 and a0901(+)=ST15" & stCon3
'
'   strExc(3) = strExc(3) & " union all select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
'      ",substrb(sp05,1,16) C3,'' C4,substrb(cpm03,1,8)||'-核稿' C5" & _
'      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
'      ",cp114 C8,C9,substrb(cp64,1,16) C10" & _
'      ",st01 C11,A0902 C12,ST16 C13,st02 C14" & _
'      " From (" & stVTB & ") X, servicepractice,casepropertymap,staff,acc090" & _
'      " where  sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP & _
'      " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
'      " and st01(+)=ep04 and a0901(+)=ST15" & stCon3
'
'   '製作中說(開收據的併入中間程序)
'   '請款日
'   If txt1(1) = "1" Then
'      '請款單
'      '設計&聯合申請:承辦點數=申請案服務費的67%
'      '發明新型:承辦點數=申請案服務費
'      'Modify by Morgan 2009/10/30
'      'stVTB = "select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",(a1l05-a1l07)/1000*decode(a1l04,'101',1,'102',1,0.67) C9" & _
'         " from acc1k0,caseprogress c1,acc1l0 where a1k12 is null" & stCon1K0 & _
'         " and cp60(+)=a1k01 and cp27<=a1k02+19110000 and cp14 is not null and cp10='210'" & stCon & stCon0 & _
'         " and a1l01(+)=cp60 and (a1l04 in ('101','102','103','105') or a1l04 is null)"
'
'      'stVTB = stVTB & " union all select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",(a1l05-a1l07)/1000*decode(a1l04,'101',1,'102',1,0.67) C9" & _
'         " from caseprogress c1,acc1k0,acc1l0 where cp60>'X'" & stConCP & _
'         " and cp27>0 and cp14 is not null and cp10='210'" & _
'         " and a1k01(+)=cp60 and a1k02+19110000<cp27 and a1k12 is null" & stCon & stCon0 & _
'         " and a1l01(+)=cp60 and (a1l04 in ('101','102','103','105') or a1l04 is null)"
'
'      '設計(先抓製作中說,沒有才抓申請案)
'      stVTB = "select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",decode(L1.a1l04,null,0.67*(L2.a1l05-L2.a1l07),(L1.a1l05-L1.a1l07))/1000 C9" & _
'         " from acc1k0,caseprogress c1,patent,acc1l0 L1,acc1l0 L2 where a1k12 is null" & stCon1K0 & _
'         " and cp60(+)=a1k01 and cp27<=a1k02+19110000 and cp14 is not null and cp10='210'" & stCon & stCon0 & _
'         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa08='3'" & _
'         " and L1.a1l01(+)=cp60 and L1.a1l04(+)=cp10" & _
'         " and L2.a1l01(+)=cp60 and instr('103,105',L2.a1l04(+))>0"
'
'      '發明新型:承辦點數=製作中說服務費
'      stVTB = stVTB & " union all select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",(a1l05-a1l07)/1000 C9" & _
'         " from acc1k0,caseprogress c1,patent,acc1l0 where a1k12 is null" & stCon1K0 & _
'         " and cp60(+)=a1k01 and cp27<=a1k02+19110000 and cp14 is not null and cp10='210'" & stCon & stCon0 & _
'         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa08<>'3'" & _
'         " and a1l01(+)=cp60 and a1l04(+)=cp10"
'
'      '請款單:先請款後發文案件
'
'      '發明新型:承辦點數=製作中說服務費
'      stVTB = stVTB & " union all select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",decode(L1.a1l04,null,0.67*(L2.a1l05-L2.a1l07),(L1.a1l05-L1.a1l07))/1000 C9" & _
'         " from caseprogress c1,patent,acc1k0,acc1l0 L1,acc1l0 L2 where cp60>'X'" & stConCP & _
'         " and cp27>0 and cp14 is not null and cp10='210'" & _
'         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa08='3'" & _
'         " and a1k01(+)=cp60 and a1k02+19110000<cp27 and a1k12 is null" & stCon & stCon0 & _
'         " and L1.a1l01(+)=cp60 and L1.a1l04(+)=cp10" & _
'         " and L2.a1l01(+)=cp60 and instr('103,105',L2.a1l04(+))>0"
'
'      stVTB = stVTB & " union all select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",(a1l05-a1l07)/1000 C9" & _
'         " from caseprogress c1,patent,acc1k0,acc1l0 where cp60>'X'" & stConCP & _
'         " and cp27>0 and cp14 is not null and cp10='210'" & _
'         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa08<>'3'" & _
'         " and a1k01(+)=cp60 and a1k02+19110000<cp27 and a1k12 is null" & stCon & stCon0 & _
'         " and a1l01(+)=cp60 and a1l04(+)=cp10"
'   '發文日
'   Else
'      stVTB = "select cp27 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp05 as cp27,cp60,cp64,cp113" & _
'         ",decode(L1.a1l04,null,0.67*(L2.a1l05-L2.a1l07),(L1.a1l05-L1.a1l07))/1000 C9" & _
'         " from caseprogress c1,patent,acc1l0 L1,acc1l0 L2 where cp14 is not null and cp10='210'" & stCon & stCon0 & _
'         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa08='3'" & _
'         " and L1.a1l01(+)=cp60 and L1.a1l04(+)=cp10" & _
'         " and L2.a1l01(+)=cp60 and instr('103,105',L2.a1l04(+))>0"
'
'      stVTB = stVTB & " union all select cp27 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp05 as cp27,cp60,cp64,cp113" & _
'         ",(a1l05-a1l07)/1000 C9" & _
'         " from caseprogress c1,patent,acc1l0 where cp14 is not null and cp10='210'" & stCon & stCon0 & _
'         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa08<>'3'" & _
'         " and a1l01(+)=cp60 and a1l04(+)=cp10"
'
'   End If
'   'Modify by Morgan 2010/8/16 百年蟲
'   strExc(4) = " union all select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
'      ",substrb(pa05,1,16) C3,substrb(ptm03,1,6) C4,substrb(cpm03,1,8) C5" & _
'      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
'      ",cp113 C8,decode(ST15,'F22',0,C9) C9,substrb(cp64,1,16) C10" & _
'      ",st01 C11,A0902 C12,ST16 C13,st02 C14" & _
'      " From (" & stVTB & ") X,patent,patenttrademarkmap,casepropertymap,staff,acc090" & _
'      " where  pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA & _
'      " and ptm01(+)=1 and ptm02(+)=pa08 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
'      " and st01(+)=cp14 and a0901(+)=ST15" & stCon3
'
'   strExc(0) = strExc(1) & strExc(2) & strExc(3) & strExc(4) & " order by C12,C13,C11,C1,C2,C5"
'
'   intI = 1
'   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      cnnConnection.Execute "DELETE FROM R060312 WHERE ID='" & strUserNum & "' "
'      stName = "": stGrp = "": stDep = "": stID = ""
'      dblPoint = 0: dblTotPoint = 0
'      dblHour = 0: dblTotHour = 0
'      m_RptType = 0: m_GrpType = 0
'      Page = 0: m_iPages = 0
'      iPrint = 0
'
'      If txt1(14) = "1" Then
'         m_bPrinter = False
'         Set m_Device = Picture1
'         m_Device.AutoRedraw = True
'         m_Device.Width = 16836
'         m_Device.Height = 11904
'         DelPic
'      Else
'         m_bPrinter = True
'         Set m_Device = Printer
'         m_Device.Orientation = 2
'      End If
'
'      GetPleft1
'      With rsQuery
'      .MoveFirst
'      Do While Not .EOF
'         If txt1(12) = "2" Then
'            stID = "" & .Fields("C11")
'            stName = "" & .Fields("C14")
'            '2010/1/8 MODIFY BY SONIA
'            'stGrp = "" & .Fields("C13")
'            stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
'            '2010/1/8 END
'            stDep = "" & .Fields("C12")
'            'Add by Morgan 2007/8/23 若沒組別時設為部門
'            If stGrp = "" Then stGrp = stDep
'         Else
'            If stID <> "" & .Fields("C11") Then
'               If stID <> "" Then
'                  PrintSubTot stID
'               End If
'               dblPoint = 0: dblHour = 0
'               stID = "" & .Fields("C11")
'               stName = "" & .Fields("C14")
'               '2010/1/8 MODIFY BY SONIA
'               'stGrp = "" & .Fields("C13")
'               stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
'               '2010/1/8 END
'               stDep = "" & .Fields("C12")
'               'Add by Morgan 2007/8/23 若沒組別時設為部門
'               If stGrp = "" Then stGrp = stDep
'               PrintTitle2
'               iPrint = iPrint + 300
'            Else
'               iPrint = iPrint + 400
'               If iPrint > m_Device.ScaleHeight - 600 Then
'                  PrintTitle2
'                  iPrint = iPrint + 300
'               End If
'            End If
'
'            For ii = 1 To 10
'               strExc(0) = ""
'               Select Case ii
'                  Case 8
'                     If "" & .Fields("C" & ii) <> "" Then
'                        strExc(0) = Format("" & .Fields("C" & ii), "0.0")
'                     End If
'                     m_Device.CurrentX = PLeft(ii) + 960 - m_Device.TextWidth(strExc(0))
'                     dblHour = dblHour + Val(strExc(0))
'                     dblTotHour = dblTotHour + Val(strExc(0))
'                  Case 9
'                     If "" & .Fields("C" & ii) <> "" Then
'                        strExc(0) = Format("" & .Fields("C" & ii), "0.00")
'                     End If
'                     m_Device.CurrentX = PLeft(ii) + 960 - m_Device.TextWidth(strExc(0))
'                     dblPoint = dblPoint + Val(strExc(0))
'                     dblTotPoint = dblTotPoint + Val(strExc(0))
'                  Case Else
'                     strExc(0) = "" & .Fields("C" & ii)
'                     'Add by Morgan 2008/1/3
'                     If ii = 1 And Right("0" & .Fields("C1"), 9) < Right("0" & .Fields("C7"), 9) Then
'                        strExc(0) = strExc(0) & "*"
'                     End If
'                     'end 2008/1/3
'                     m_Device.CurrentX = PLeft(ii)
'               End Select
'               m_Device.CurrentY = iPrint
'               m_Device.Print strExc(0)
'            Next
'         End If
'         strSql = " INSERT INTO R060312 (ID, R045001, R045003, R045005, R045008, R045004, R045006, R045007) values('" & strUserNum & "','" & stID & "','" & .Fields("C5") & "'," & Format(Val("" & .Fields("C8")), "0.0") & "," & Format(Val("" & .Fields("C9")), "0.00") & ",'" & stDep & "','" & stGrp & "','" & stName & "')"
'         cnnConnection.Execute strSql, intI
'         .MoveNext
'      Loop
'
'      If txt1(12) <> "2" Then
'         PrintSubTot stID
'         If m_bPrinter = True Then
'            m_Device.EndDoc
'         End If
'      End If
'
'      '列印統計表
'      If txt1(6) = "" And txt1(12) >= "2" Then
'         If m_bPrinter = True Then
'            m_Device.Orientation = 2
'         End If
'         Page = 0
'         m_GrpType = 1
'         PrintStatistic
'         If m_bPrinter = True Then
'            m_Device.EndDoc
'         End If
'         If txt1(11) = "" Then
'            If m_bPrinter = True Then
'               m_Device.Orientation = 2
'            End If
'            Page = 0
'            m_GrpType = 2
'            PrintStatistic
'            If m_bPrinter = True Then
'               m_Device.EndDoc
'            End If
'            'Modify by Morgan 2007/11/26 加程序人員選項,全部改為3
'            'If txt1(13) = "2" Then
'            If txt1(13) = "3" Then
'               If m_bPrinter = True Then
'                  m_Device.Orientation = 2
'               End If
'               Page = 0
'               m_GrpType = 3
'               PrintStatistic
'               If m_bPrinter = True Then
'                  m_Device.EndDoc
'               End If
'            End If
'         End If
'      End If
'      End With
'      If m_bPrinter = True Then
'         ShowPrintOk
'      Else
'         If m_iPages > 0 Then
'            SetPic m_iPages
'            frm060312_1.m_ImageW = m_Device.Width
'            frm060312_1.m_ImageH = m_Device.Height
'            frm060312_1.m_iPages = m_iPages
'            frm060312_1.Show
'         End If
'      End If
'   Else
'      MsgBox "無可列印資料！"
'   End If
'   Set rsQuery = Nothing
'   Set m_Device = Nothing
'End Sub

Private Sub DelPic()
   Dim strPicFileName As String
   'Modified by Morgan 2021/5/21 修正多人系統問題
   'strPicFileName = App.path & "\$tmp_*.tmp"
   strPicFileName = m_AttachPath & "\$tmp_*.tmp"
   If Dir(strPicFileName) <> "" Then
      Kill strPicFileName
   End If
   m_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
End Sub

Private Sub SetPic(idx As Integer)

   Dim strPicFileName As String
   'Modified by Morgan 2021/5/21 修正多人系統問題
   'strPicFileName = App.path & "\$tmp_" & idx & ".tmp"
   strPicFileName = m_AttachPath & "\$tmp_" & idx & ".tmp"
   
'   Clipboard.Clear
'   Clipboard.SetData Picture1.Image
'   Set m_Pictures(m_iPages - 1) = Clipboard.GetData
'   Set m_Pictures(idx) = Picture1.Image

   SavePicture Picture1.Image, strPicFileName
   '要用覆蓋的否則會錯誤--VB Bug
   'Picture1.Cls
   m_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
   
End Sub

'合計
Private Sub PrintStatistic1()
   Dim stCon As String
   Dim stCon1 As String
   Dim iRows As Integer
   Dim adoRst As ADODB.Recordset
   Dim stSQL As String
   
   Select Case m_GrpType
      Case 1
         stCon = " and R045004='" & stDep & "' and R045006='" & stGrp & "'"
         stCon1 = " and a0902='" & stDep & "' and CST16(ST16)='" & stGrp & "'"
      Case 2
         stCon = " and R045004='" & stDep & "'"
         stCon1 = " and a0902='" & stDep & "'"
   End Select
   'Modify by Morgan 2010/11/16 +統計非核稿點數
   strExc(0) = "select sum(R045005) S1,sum(R045008) S2,sum(Decode(instr(R045003,'核稿'),0,R045008)) S3 from R060312 where ID='" & strUserNum & "'" & stCon
   stSQL = "select NVL(sum(R03),0) C1 from R060312_1,staff,acc090 where ID='" & strUserNum & "' and st01(+)=R01 and a0901(+)=st15" & stCon1
   
   intI = 1
   Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      PrintLine
      NewLine
      With AdoRecordSet3
         
         m_Device.CurrentX = PLeft(1)
         m_Device.CurrentY = iPrint
         strExc(0) = ""
         Select Case m_GrpType
            Case "1": strExc(0) = "組別合計："
            Case "2": strExc(0) = "部門合計："
            Case "3": strExc(0) = "外專合計："
         End Select
         m_Device.Print strExc(0)
      
         strExc(1) = Format("" & .Fields("S1"), "0.0")
         strExc(2) = Format("" & .Fields("S2"), "0.00")
         strExc(3) = Format("" & .Fields("S3"), "0.00") 'Add by Morgan 2010/11/16
         strExc(4) = ""
         strExc(5) = "" '小計 Added by Morgan 2012/2/2
         intI = 1
         Set adoRst = ClsLawReadRstMsg(intI, stSQL)
         If intI = 1 Then
            With adoRst
            If .Fields(0) > 0 Then
               strExc(4) = Format(.Fields(0), "0.00")
               strExc(5) = Format(Val(Format(strExc(2))) + Val(Format(strExc(4))), "0.00") '小計 Added by Morgan 2012/2/2
            End If
            End With
         End If
         
         
         '列印點數小計
         m_Device.CurrentX = PLeft(11) + 1200 - m_Device.TextWidth(strExc(1))
         m_Device.CurrentY = iPrint
         m_Device.Print strExc(1)
         m_Device.CurrentX = PLeft(12) + 1200 - m_Device.TextWidth(strExc(2))
         m_Device.CurrentY = iPrint
         m_Device.Print strExc(2)
         
         iPrint = iPrint - 300
         iRows = 0
         
         'Modify by Morgan 2011/4/12
         'strExc(0) = "select R045003,count(*)  from R060312 where ID='" & strUserNum & "'" & stCon & " group by R045003"
         strExc(0) = "select R045003,count(*)  from R060312 where ID='" & strUserNum & "'" & stCon & " and R045003 is not null group by R045003"
         
         intI = 1
         Set adoRecordset1 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With adoRecordset1
            .MoveFirst
            Do While .EOF = False
               For i = 0 To 4
                  StrTemp4(i) = CheckStr(.Fields(0))
                  StrTemp5(i) = CheckStr(.Fields(1))
                  .MoveNext
                  If .EOF = True Then
                      For j = i + 1 To 4
                          StrTemp4(j) = ""
                          StrTemp5(j) = ""
                      Next j
                      Exit For
                  End If
               Next i
               PrintSubTot1
               'Add by Morgan 2010/5/7 列印分配點數小計
               iRows = iRows + 1
               If iRows = 2 Then
               'Add by Morgan 2010/11/16 +統計非核稿點數
                  m_Device.CurrentX = PLeft(12) + 300 - m_Device.TextWidth("不含核稿點數：")
                  m_Device.CurrentY = iPrint
                  m_Device.Print "不含核稿點數："
                  m_Device.CurrentX = PLeft(12) + 1200 - m_Device.TextWidth(strExc(3))
                  m_Device.CurrentY = iPrint
                  m_Device.Print strExc(3)
               ElseIf iRows = 3 Then
               'end 2010/11/16
                  If strExc(4) <> "" Then
                     m_Device.CurrentX = PLeft(12) + 300 - m_Device.TextWidth("分配點數：")
                     m_Device.CurrentY = iPrint
                     m_Device.Print "分配點數："
                     m_Device.CurrentX = PLeft(12) + 1200 - m_Device.TextWidth(strExc(4))
                     m_Device.CurrentY = iPrint
                     m_Device.Print strExc(4)
                  End If
               'Added by Morgan 2012/2/2
               ElseIf iRows = 4 Then
                  If strExc(5) <> "" Then
                     m_Device.CurrentX = PLeft(12) + 300 - m_Device.TextWidth("小計：")
                     m_Device.CurrentY = iPrint
                     m_Device.Print "小計："
                     m_Device.CurrentX = PLeft(12) + 1200 - m_Device.TextWidth(strExc(5))
                     m_Device.CurrentY = iPrint
                     m_Device.Print strExc(5)
                  End If
               'end 2012/2/2
               
               End If
               'end 2010/5/7
            Loop
            End With
         End If
         'Add by Morgan 2010/5/7 列印分配點數小計
         If iRows < 2 Then
         'Add by Morgan 2010/11/16 +統計非核稿點數
            NewLine 300
            m_Device.CurrentX = PLeft(12) + 300 - m_Device.TextWidth("不含核稿點數：")
            m_Device.CurrentY = iPrint
            m_Device.Print "不含核稿點數："
            m_Device.CurrentX = PLeft(12) + 1200 - m_Device.TextWidth(strExc(3))
            m_Device.CurrentY = iPrint
            m_Device.Print strExc(3)
         End If
         If iRows < 3 And strExc(4) <> "" Then
         'end 2010/11/16
            NewLine 300
            m_Device.CurrentX = PLeft(12) + 300 - m_Device.TextWidth("分配點數：")
            m_Device.CurrentY = iPrint
            m_Device.Print "分配點數："
            m_Device.CurrentX = PLeft(12) + 1200 - m_Device.TextWidth(strExc(4))
            m_Device.CurrentY = iPrint
            m_Device.Print strExc(4)
         End If
         
         'end 2010/5/7
         
         'Added by Morgan 2012/2/2
         If iRows < 4 And strExc(5) <> "" Then
            NewLine 300
            m_Device.CurrentX = PLeft(12) + 300 - m_Device.TextWidth("小計：")
            m_Device.CurrentY = iPrint
            m_Device.Print "小計："
            m_Device.CurrentX = PLeft(12) + 1200 - m_Device.TextWidth(strExc(5))
            m_Device.CurrentY = iPrint
            m_Device.Print strExc(5)
         End If
         'end 2012/2/2
         
      End With
      NewLine
      m_Device.CurrentX = PLeft(1)
      m_Device.CurrentY = iPrint
      m_Device.Print "PS:不含非個人點數"
   End If
   Set adoRst = Nothing
End Sub
'統計表
Private Sub PrintStatistic()
   Dim stCol1 As String, iGrpCount As Integer
   Dim iRows As Integer
   Dim adoRst As ADODB.Recordset
   Dim stSQL As String
   
   m_RptType = 1
   'Modify by Morgan 2010/11/16 +統計非核稿點數
   'Modified by Morgan 2013/1/10 +組別以代碼排序
   Select Case m_GrpType
      Case 1 '組別
         strExc(0) = "select R045004,R045006,R045001,R045007,sum(R045005) S1,sum(R045008) S2,sum(Decode(instr(R045003,'核稿'),0,R045008)) S3 from R060312 where ID='" & strUserNum & "' group by R045004,R045014,R045006,R045001,R045007"
      Case 2 '部門
         strExc(0) = "select R045004,R045006,sum(R045005) S1,sum(R045008) S2,sum(Decode(instr(R045003,'核稿'),0,R045008)) S3 from R060312 where ID='" & strUserNum & "' group by R045004,R045014,R045006"
      Case 3 '全部
         strExc(0) = "select R045004,sum(R045005) S1,sum(R045008) S2,sum(Decode(instr(R045003,'核稿'),0,R045008)) S3 from R060312 where ID='" & strUserNum & "' group by R045004"
   End Select
   
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      stDep = "": stGrp = "": stID = "": stName = ""
      iGrpCount = 0
      With adoRecordset
      
      Do While Not .EOF
         Select Case m_GrpType
            Case 1
               If stDep <> "" & .Fields("R045004") Or stGrp <> "" & .Fields("R045006") Then
                  If stGrp <> "" Then
                     PrintStatistic1
                  End If
                  stDep = "" & .Fields("R045004")
                  stGrp = "" & .Fields("R045006")
                  PrintTitle2
                  iPrint = iPrint + 300
               Else
                  PrintLine
                  NewLine
               End If
               stID = "" & .Fields("R045001")
               stName = "" & .Fields("R045007")
               stCol1 = stName
            Case 2
               If stDep <> "" & .Fields("R045004") Then
                  '兩組以上才印組別合計
                  If stDep <> "" And iGrpCount > 1 Then
                     PrintStatistic1
                  End If
                  stDep = "" & .Fields("R045004")
                  PrintTitle2
                  iPrint = iPrint + 300
                  iGrpCount = 0
               Else
                  PrintLine
                  NewLine
               End If
               iGrpCount = iGrpCount + 1
               stGrp = "" & .Fields("R045006")
               stCol1 = stGrp
            Case 3
               If stDep = "" Then
                  PrintTitle2
                  iPrint = iPrint + 300
                  stDep = "全部"
               Else
                  PrintLine
                  NewLine
               End If
               stDep = "" & .Fields("R045004")
               stCol1 = stDep
         End Select
         
         '列印承辦人
         m_Device.CurrentX = PLeft(1)
         m_Device.CurrentY = iPrint
         m_Device.Print stCol1
         
         strExc(1) = Format("" & .Fields("S1"), "0.0")
         strExc(2) = Format("" & .Fields("S2"), "0.00")
         strExc(3) = Format("" & .Fields("S3"), "0.00") 'Add by Morgan 2010/11/16
         
         '列印點數小計
         m_Device.CurrentX = PLeft(11) + 1200 - m_Device.TextWidth(strExc(1))
         m_Device.CurrentY = iPrint
         m_Device.Print strExc(1)
         m_Device.CurrentX = PLeft(12) + 1200 - m_Device.TextWidth(strExc(2))
         m_Device.CurrentY = iPrint
         m_Device.Print strExc(2)
         
         'Modify by Morgan 2011/4/12 排除只有分配點數的資料 & "and R045003 is not null"
         Select Case m_GrpType
            Case 1 '組別
               strExc(0) = "select R045003,count(*) C1 from R060312 where ID='" & strUserNum & "' and R045004='" & stDep & "' and R045006='" & stGrp & "' and R045001='" & stID & "' and R045003 is not null group by R045003"
               stSQL = "select NVL(sum(R03),0) C1 from R060312_1,staff,acc090 where ID='" & strUserNum & "' and st01(+)=R01 and a0901(+)=st15 and a0902='" & stDep & "' and CST16(ST16)='" & stGrp & "' and R01='" & stID & "'"
               
            Case 2 '部門
               strExc(0) = "select R045003,count(*) C1 from R060312 where ID='" & strUserNum & "' and R045004='" & stDep & "' and R045006='" & stGrp & "' and R045003 is not null group by R045003"
               stSQL = "select NVL(sum(R03),0) C1 from R060312_1,staff,acc090 where ID='" & strUserNum & "' and st01(+)=R01 and a0901(+)=st15 and a0902='" & stDep & "' and CST16(ST16)='" & stGrp & "'"
               
            Case 3 '全部
               strExc(0) = "select R045003,count(*) C1 from R060312 where ID='" & strUserNum & "' and R045004='" & stDep & "' and R045003 is not null group by R045003"
               stSQL = "select NVL(sum(R03),0) C1 from R060312_1,staff,acc090 where ID='" & strUserNum & "' and st01(+)=R01 and a0901(+)=st15 and a0902='" & stDep & "'"
               
         End Select
         
         strExc(4) = ""
         strExc(5) = "" '小計 Added by Morgan 2012/2/2
         intI = 1
         Set adoRst = ClsLawReadRstMsg(intI, stSQL)
         If intI = 1 Then
            With adoRst
            If .Fields(0) > 0 Then
               strExc(4) = Format(.Fields(0), "0.00")
               strExc(5) = Format(Val(Format(strExc(2))) + Val(Format(strExc(4))), "0.00") '小計 Added by Morgan 2012/2/2
            End If
            End With
         End If
         
         iRows = 0
         
         iPrint = iPrint - 300
         
         intI = 1
         Set adoRecordset1 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            
            With adoRecordset1
            .MoveFirst
            Do While .EOF = False
               For i = 0 To 4
                  StrTemp4(i) = CheckStr(.Fields("R045003"))
                  StrTemp5(i) = CheckStr(.Fields("C1"))
                  .MoveNext
                  If .EOF = True Then
                      For j = i + 1 To 4
                          StrTemp4(j) = ""
                          StrTemp5(j) = ""
                      Next j
                      Exit For
                  End If
               Next i
               PrintSubTot1
               'Add by Morgan 2010/5/7 列印分配點數小計
               iRows = iRows + 1
               If iRows = 2 Then
               'Add by Morgan 2010/11/16
                  m_Device.CurrentX = PLeft(12) + 300 - m_Device.TextWidth("不含核稿點數：")
                  m_Device.CurrentY = iPrint
                  m_Device.Print "不含核稿點數："
                  m_Device.CurrentX = PLeft(12) + 1200 - m_Device.TextWidth(strExc(3))
                  m_Device.CurrentY = iPrint
                  m_Device.Print strExc(3)
               ElseIf iRows = 3 And strExc(4) <> "" Then
               'end 2010/11/16
                  m_Device.CurrentX = PLeft(12) + 300 - m_Device.TextWidth("分配點數：")
                  m_Device.CurrentY = iPrint
                  m_Device.Print "分配點數："
                  m_Device.CurrentX = PLeft(12) + 1200 - m_Device.TextWidth(strExc(4))
                  m_Device.CurrentY = iPrint
                  m_Device.Print strExc(4)
                  
               'Added by Morgan 2012/2/2
               ElseIf iRows = 4 And strExc(5) <> "" Then
                     m_Device.CurrentX = PLeft(12) + 300 - m_Device.TextWidth("小計：")
                     m_Device.CurrentY = iPrint
                     m_Device.Print "小計："
                     m_Device.CurrentX = PLeft(12) + 1200 - m_Device.TextWidth(strExc(5))
                     m_Device.CurrentY = iPrint
                     m_Device.Print strExc(5)
               'end 2012/2/2
               
               End If
               'end 2010/5/7
            Loop
            End With
         End If
         
         
         If iRows = 0 Then NewLine 300 'Add by Morgan 2011/4/11 沒資料也要跳行
         
         'Add by Morgan 2010/5/7 列印分配點數小計
         If iRows < 2 Then
         'Add by Morgan 2010/11/16
            NewLine 300
            m_Device.CurrentX = PLeft(12) + 300 - m_Device.TextWidth("不含核稿點數：")
            m_Device.CurrentY = iPrint
            m_Device.Print "不含核稿點數："
            m_Device.CurrentX = PLeft(12) + 1200 - m_Device.TextWidth(strExc(3))
            m_Device.CurrentY = iPrint
            m_Device.Print strExc(3)
         End If
         If iRows < 3 And strExc(4) <> "" Then
         'end 2010/11/16
            NewLine 300
            m_Device.CurrentX = PLeft(12) + 300 - m_Device.TextWidth("分配點數：")
            m_Device.CurrentY = iPrint
            m_Device.Print "分配點數："
            m_Device.CurrentX = PLeft(12) + 1200 - m_Device.TextWidth(strExc(4))
            m_Device.CurrentY = iPrint
            m_Device.Print strExc(4)
         End If
         'end 2010/5/7
         
         'Added by Morgan 2012/2/2
         If iRows < 4 And strExc(5) <> "" Then
            NewLine 300
            m_Device.CurrentX = PLeft(12) + 300 - m_Device.TextWidth("小計：")
            m_Device.CurrentY = iPrint
            m_Device.Print "小計："
            m_Device.CurrentX = PLeft(12) + 1200 - m_Device.TextWidth(strExc(5))
            m_Device.CurrentY = iPrint
            m_Device.Print strExc(5)
         End If
         'end 2012/2/2
         
         .MoveNext
      Loop
      If txt1(6) = "" Then
         If Not (m_GrpType = "2" And iGrpCount < 2) Then
            PrintStatistic1
         End If
      End If
      End With
   End If
   Set adoRst = Nothing
End Sub

Private Sub PrintSubTot(Optional p_ID As String)
   
   Dim stCon As String
   Dim iRows As Integer
   
   PrintLine
   
   NewLine 300
   
   If p_ID <> "" Then
      stCon = " and R045001='" & p_ID & "'"
      strExc(1) = Format(dblHour, "0.0")
      strExc(2) = Format(dblPoint, "0.00")
      'Add by Morgan 2010/11/15
      strExc(3) = Format(dblPoint - dblPoint2, "0.00")
   Else
      strExc(1) = Format(dblTotHour, "0.0")
      strExc(2) = Format(dblTotPoint, "0.00")
      'Add by Morgan 2010/11/15
      strExc(3) = Format(dblTotPoint - dblTotPoint2, "0.00")
   End If
   
   '列印點數小計
   m_Device.CurrentX = PLeft(8) + 960 - m_Device.TextWidth(strExc(1))
   m_Device.CurrentY = iPrint
   m_Device.Print strExc(1)
   m_Device.CurrentX = PLeft(9) + 960 - m_Device.TextWidth(strExc(2))
   m_Device.CurrentY = iPrint
   m_Device.Print strExc(2)
   
   iRows = 0
   
   iPrint = iPrint - 300
   
   strExc(0) = "select R045003,count(*)  from R060312 where ID='" & strUserNum & "'" & stCon & " group by R045003"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      .MoveFirst
      Do While .EOF = False
         For i = 0 To 4
            StrTemp4(i) = CheckStr(.Fields(0))
            StrTemp5(i) = CheckStr(.Fields(1))
            .MoveNext
            If .EOF = True Then
                For j = i + 1 To 4
                    StrTemp4(j) = ""
                    StrTemp5(j) = ""
                Next j
                Exit For
            End If
         Next i
         PrintSubTot1
         'Add by Morgan 2010/11/16
         iRows = iRows + 1
         If iRows = 2 Then
            m_Device.CurrentX = PLeft(9) + 960 - m_Device.TextWidth(strExc(3))
            m_Device.CurrentY = iPrint
            m_Device.Print strExc(3) & " (不含核稿)"
         End If
         'end 2010/11/16
      Loop
      End With
   End If
   
   'Add by Morgan 2010/11/16
   If iRows < 2 Then
      NewLine 300
      m_Device.CurrentX = PLeft(9) + 960 - m_Device.TextWidth(strExc(3))
      m_Device.CurrentY = iPrint
      m_Device.Print strExc(3) & " (不含核稿)"
   End If
   'end 2010/11/16
   
   PrintSharePoint p_ID 'Add by Morgan 2010/4/19
   
End Sub

'Add by Morgan 2010/4/19
'列印分配點數
Private Sub PrintSharePoint(p_ID As String)
   Dim lngX0 As Long, lngX As Long, dblTot As Double
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   strExc(0) = "select sqldatet(R09)||R10 C1" & _
      ",R04||'-'||R05||decode(R06||R07,'000','','-'||R06||'-'||R07) C2" & _
      ",substrb(decode(nvl(pa09,sp09),'000',cpm03,cpm04)||decode(R11,'Y','-核稿'),1,14) C3" & _
      ",R03 C4" & _
      " from R060312_1,patent,servicepractice,casepropertymap" & _
      " WHERE ID='" & strUserNum & "' and R01='" & p_ID & "'" & _
      " AND cpm01(+)=R04 and cpm02(+)=R08" & _
      " and pa01(+)=R04 and pa02(+)=R05 and pa03(+)=R06 and pa04(+)=R07" & _
      " and sp01(+)=R04 and sp02(+)=R05 and sp03(+)=R06 and sp04(+)=R07"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      NewLine
      NewLine
      If m_RptType = 0 Then
         lngX0 = PLeft(1)
      Else
         lngX0 = PLeft(2) + 400
      End If
      m_Device.CurrentX = lngX0
      m_Device.CurrentY = iPrint
      m_Device.Print "請款單分配點數："
      NewLine
      
      lngX = lngX0
      m_Device.CurrentX = lngX
      m_Device.CurrentY = iPrint
      If txt1(1) = "1" Then
         m_Device.Print "請款日"
      Else
         m_Device.Print "發文日"
      End If
            
      lngX = lngX + m_Device.TextWidth(String(11, "@"))
      m_Device.CurrentX = lngX
      m_Device.CurrentY = iPrint
      m_Device.Print "本所案號"
      
      lngX = lngX + m_Device.TextWidth(String(17, "@"))
      m_Device.CurrentX = lngX
      m_Device.CurrentY = iPrint
      m_Device.Print "案件性質"
      
      lngX = lngX + m_Device.TextWidth(String(16, "@"))
      m_Device.CurrentX = lngX
      m_Device.CurrentY = iPrint
      m_Device.Print "點數"
      
      NewLine
      m_Device.CurrentX = lngX0
      m_Device.CurrentY = iPrint
      m_Device.Print String(51, "-")
      
      iPrint = iPrint - 100
      
      With RsTemp
      Do While Not .EOF
         NewLine
         '日期
         lngX = lngX0
         m_Device.CurrentX = lngX
         m_Device.CurrentY = iPrint
         m_Device.Print .Fields("C1")
         '本所案號
         lngX = lngX + m_Device.TextWidth(String(11, "@"))
         m_Device.CurrentX = lngX
         m_Device.CurrentY = iPrint
         m_Device.Print .Fields("C2")
         '案件性質
         lngX = lngX + m_Device.TextWidth(String(17, "@"))
         m_Device.CurrentX = lngX
         m_Device.CurrentY = iPrint
         m_Device.Print .Fields("C3")
         '點數
         lngX = lngX + m_Device.TextWidth(String(16, "@"))
         m_Device.CurrentX = lngX + m_Device.TextWidth(String(7, "@")) - m_Device.TextWidth(Format(.Fields("C4"), "0.00"))
         m_Device.CurrentY = iPrint
         m_Device.Print Format(.Fields("C4"), "0.00")
         dblTot = dblTot + Val("" & .Fields("C4"))
         .MoveNext
      Loop
      End With
      
      NewLine 300
      m_Device.CurrentX = lngX0
      m_Device.CurrentY = iPrint
      m_Device.Print String(51, "-")
      
      NewLine 300
      m_Device.CurrentX = lngX + m_Device.TextWidth(String(7, "@")) - m_Device.TextWidth(Format(dblTot, "0.00"))
      m_Device.CurrentY = iPrint
      m_Device.Print Format(dblTot, "0.00")
   End If
End Sub

Sub PrintSubTot1()
   Dim lngX0 As Long
   
   NewLine 300
   
   If m_RptType = 0 Then
      lngX0 = PLeft(1)
   Else
      lngX0 = PLeft(2) + 400
   End If
   
   For j = 0 To 4
      m_Device.CurrentX = lngX0 + (j * 2300)
      m_Device.CurrentY = iPrint
      m_Device.Print StrConv(MidB(StrConv(StrTemp4(j), vbFromUnicode), 1, 14), vbUnicode)
      m_Device.CurrentX = lngX0 + ((j + 1) * 2300) - 400 - m_Device.TextWidth(StrTemp5(j))
      m_Device.CurrentY = iPrint
      m_Device.Print StrTemp5(j)
   Next j
End Sub

Private Sub NewLine(Optional iHeight As Integer = 400)
   iPrint = iPrint + iHeight
   If iPrint > m_Device.ScaleHeight - 800 Then
      PrintTitle2
      iPrint = iPrint + 300
   End If
End Sub

Private Sub PrintLine(Optional iType As Integer = 0)
   iPrint = iPrint + 300
   m_Device.CurrentX = PLeft(1)
   m_Device.CurrentY = iPrint
   If iType = 1 Then
      m_Device.Line (PLeft(1), iPrint + 150)-(PLeft(1) + m_Device.TextWidth(String(200, "-")), iPrint + 150)
   Else
      m_Device.Print String(200, "-")
   End If
End Sub

Private Sub Form_Load()
      
   MoveFormToCenter Me
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
   '系統類別
   '2008/12/29 modify by sonia 加入L,因為L-004061
   '2009/10/15 modify by sonia 加入CFL,因為魏汎娟97009
   'Modifiec by Lydia 2015/01/22
   txt1(0) = DefSysTxt
   '列印別--預設請款    '2008/2/1改預設
   txt1(1) = "1"
   '報表別--預設明細    '2008/2/1改預設並取消全部1+2
   txt1(12) = "1"
   '列印內容--預設工程師
   txt1(13) = "1"
   
   
   'Add by Morgan 2007/9/21
   '外專工程師要控管權限
   strST05 = PUB_GetST05(strUserNum)
   'Added by Lydia 2025/02/06
   strExc(0) = "select st16,st70 from staff where st01='" & strUserNum & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strST16 = "" & RsTemp.Fields("st16")
      strST70 = "" & RsTemp.Fields("st70")
   End If
   'end 2025/02/06
   
   Select Case strST05
      Case "39" '外專工程師中級主管只可查該組
         'Modified by Lydia 2020/08/14 現在有限制人員範圍
         'txt1(11) = PUB_GetStaffST16(strUserNum)
         'txt1(11).Locked = True
         txt1(6) = strUserNum
         lbl1 = strUserName
         'end 2020/08/14
      Case "40", "49" '外專工程師只可查本人  'modify by sonia 2024/8/15 加入等級49日外專海外工程師
         txt1(6) = strUserNum
         lbl1 = strUserName
         txt1(6).Locked = True
   End Select
   txt1(6).Tag = txt1(6).Text 'Added by Lydia 2020/08/14
   stDefCP14 = txt1(6).Text 'Added by Lydia 2024/03/05
   
   '輸出
   If Pub_StrUserSt15 = "F21" Then
      txt1(14) = "1"
      'Modified by Morgan 2012/5/15
      'txt1(14).Enabled = False
      '各組主管可列印
      'Removed by Morgan 2012/6/1 改用權限控管
      'If strST05 <> "38" And strST05 <> "42" Then
      '   txt1(14).Enabled = False
      'End If
      'end 2012/6/1
      'end 2012/5/15
   Else
      txt1(14) = "2"
   End If
   
   'Added by Morgan 2012/6/1 改用權限控管
   If IsUserHasRightOfFunction(Me.Name, strPrint, False) = False Then
      txt1(14) = "1"
      txt1(14).Enabled = False
   End If
   'end 2012/6/1
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060312 = Nothing
End Sub

Private Sub txt1_Change(Index As Integer)
If Index = 13 Then
   'Modified by Lydia 2015/01/22 +P案管制人
   If txt1(Index) = "4" Then
      txt1(0) = "P"
      If txt1(1) = "1" Then
         MsgBox "P案管制人只限發文別!", vbCritical
         txt1(1) = "2"
      End If
   Else
      If txt1(0) = "P" Then
         txt1(0) = DefSysTxt
      End If
      'Modify by Sindy 2015/9/22 +發文操作人
      If txt1(Index) = "5" Then
         If txt1(1) = "1" Then
            MsgBox "發文操作人只限發文別!", vbCritical
            txt1(1) = "2"
         End If
      End If
   End If
End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 1, 14
         If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
             KeyAscii = 0
         End If
      Case 11
         If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 8 Then
            KeyAscii = 0
         End If
      Case 13    '原為11,12,13於2008/2/1 取消報表別之3(全部1+2),2008/2/22加4德文組
         'Modified by Lydia 2015/01/22 +P案管制人
         'Modify by Sindy 2015/9/22 +發文操作人
         If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 8 Then
            KeyAscii = 0
         End If
      Case 12
         If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
            KeyAscii = 0
         End If
   End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Select Case Index
      Case 3, 5
         If blnClkSure = False Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               txt1(Index - 1).SetFocus
               txt1_GotFocus (Index - 1)
               Exit Sub
            End If
         Else
            blnClkSure = False
         End If
         
      Case 8
         If blnClkSure = False Then
            If Len(txt1(Index - 1)) <> 0 Then
               If Left(txt1(Index - 1), 6) <> Left(txt1(Index), 6) Then
                   s = MsgBox("申請人前 6 碼必須相同", , "USER 輸入錯誤")
                   txt1(Index - 1).SetFocus
                   Exit Sub
               End If
            End If
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               txt1(Index - 1).SetFocus
               txt1_GotFocus (Index - 1)
               Exit Sub
            End If
         Else
            blnClkSure = False
         End If
         
      Case 10
         If blnClkSure = False Then
            If Len(txt1(Index - 1)) <> 0 Then
               If Left(txt1(Index - 1), 6) <> Left(txt1(Index), 6) Then
                   s = MsgBox("代理人前 6 碼必須相同", , "USER 輸入錯誤")
                   txt1(Index - 1).SetFocus
                   Exit Sub
               End If
            End If
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               txt1(Index - 1).SetFocus
               txt1_GotFocus (Index - 1)
               Exit Sub
            End If
         Else
            blnClkSure = False
         End If
      Case Else
      
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 2, 3 '日期
         If txt1(Index) = "" Then Exit Sub
         Cancel = Not ChkDate(txt1(Index))
         If Cancel Then TextInverse txt1(Index)
         
      Case 1 '列印別
           Select Case Val(txt1(1))
           Case 1, 2
           Case Else
              s = MsgBox("列印別只能 1 或 2 !!", , "USER 輸入錯誤")
              Cancel = True
           End Select
           
      Case 6 '承辦人
         'Modified by Lydia 2020/08/14 外專工程師中級主管可輸入自己和下屬的員工編號，該欄位不可空白。
         'lbl1 = ""
         'If txt1(Index) = "" Then Exit Sub
         '2008/1/8 MODIFY BY SONIA 不檢查離職
         'If ClsPDGetStaff(txt1(Index), strExc(0)) Then
         '   lbl1 = strExc(0)
         'Else
         '   lbl1 = ""
         '   Cancel = True
         'End If
         'lbl1 = GetPrjSales(txt1(Index), "承辦人")
         ''2008/1/8 END
         If txt1(Index).Tag <> txt1(Index).Text Then
             If txt1(Index).Text = "" Then
                'Modified by Lydia 2025/02/06 外專/日專工程師中級主管(主任)可查詢底下所有人員的資料:請開放ST05="39"的人員，承辦人/核稿人欄可以空白
                'If strST05 = "39" Or strST05 = "40" Or strST05 = "49" Then     'modify by sonia 2024/8/15 加入等級49日外專海外工程師
                If strST05 = "40" Or strST05 = "49" Then
                    MsgBox "承辦人/核稿人不可空白！", vbCritical, "檢核資料"
                    Cancel = True
                    txt1(Index).Text = txt1(Index).Tag
                End If
             Else
                If (strST05 = "39" Or strST05 = "40" Or strST05 = "49") And txt1(Index).Text <> strUserNum Then    'modify by sonia 2024/8/15 加入等級49日外專海外工程師
                   'Modified by Lydia 2025/02/06 外專/日專工程師中級主管(主任)可查詢底下所有人員的資料:
                               'ST05="39"的人員在查詢資料時只能抓出ST16||ST70與操作人員的相同的承辦人/核稿人，且同時該同仁之各級主管ST52~ST55必須有操作人員。
                   'If PUB_GetST52(txt1(Index).Text, strUserNum) = False Then
                   'Modified by Lydia 2025/02/25 還原
                   'strExc(1) = ""
                   'If ChkST52Range(txt1(Index)) = "" Then
                   '   strExc(1) = "N"
                   'End If
                   'If strExc(1) <> "" Then
                   ''end 2025/02/06
                   If PUB_GetST52(txt1(Index).Text, strUserNum) = False Then
                   'end 2025/02/25
                       MsgBox "無權限查詢：" & txt1(Index), vbCritical, "檢核資料"
                       Cancel = True
                       txt1(Index).Text = txt1(Index).Tag
                   End If
                End If
             End If
             lbl1 = ""
             If txt1(Index).Text <> "" Then lbl1 = GetPrjSales(txt1(Index), "承辦人")
         End If
         txt1(Index).Tag = txt1(Index).Text
         'end 2020/08/14
'2010/1/8 CANCEL BY SONIA KEYPRESS已控制
'      Case 11 '組別   2008/2/22加德文組
'         If txt1(Index) = "" Then Exit Sub
'         Select Case Val(txt1(Index))
'           Case 1, 2, 3, 4
'           Case Else
'              s = MsgBox("組別只能 1, 2, 3 或 4 !!", , "USER 輸入錯誤")
'              Cancel = True
'         End Select
'2010/1/8 END

      Case 12 '報表別
         Select Case Val(txt1(Index))
           Case 1, 2, 3
           Case Else
              s = MsgBox("報表別只能 1, 2 或 3 !!", , "USER 輸入錯誤")
              Cancel = True
         End Select
      
      Case 13 '列印內容
         Select Case Val(txt1(Index))
         'Modified by Lydia 2015/01/22 +4 P案管制人
           Case 1, 2, 3, 4, 5
           Case Else
              s = MsgBox("列印內容只能 1, 2, 3, 4 或 5 !!", , "USER 輸入錯誤")
              Cancel = True
         End Select
   End Select

   If Cancel Then TextInverse txt1(Index)
End Sub

Sub PrintTitle2()
   Dim stCon As String
   Dim stTmp As String
   
   Page = Page + 1
   m_iPages = m_iPages + 1
      
   If m_iPages > 1 Then
      If m_bPrinter = False Then
         SetPic m_iPages - 1
      ElseIf Page > 1 Then
         m_Device.NewPage
      End If
   End If
   
   If Val(txt1(1)) = 1 Then
      stCon = "請款"
   Else
      stCon = "發文"
   End If
   
   If m_RptType = 0 Then
      stCon = stCon & "明細表"
   Else
      stCon = stCon & "統計表"
   End If
   'Modified by Lydia 2015/01/23
   If txt1(13) = "4" Then
      stTmp = "P案管制人" & stCon
   Else
      stTmp = GetTitleNick & "承辦人" & stCon
   End If
   iPrint = 500
   m_Device.FontName = "細明體"
   m_Device.Font.Size = 22
   m_Device.Font.Bold = True
   m_Device.Font.Underline = True
   m_Device.CurrentX = (PLeft(12) - m_Device.TextWidth(stTmp)) / 2
   m_Device.CurrentY = iPrint
   m_Device.Print stTmp
      
   If Val(txt1(1)) = 1 Then
       stTmp = "請款日：" & Format(ChangeTStringToTDateString(txt1(2)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
   Else
       stTmp = "發文日：" & Format(ChangeTStringToTDateString(txt1(2)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
   End If
   
   iPrint = iPrint + 500
   m_Device.Font.Bold = False
   m_Device.Font.Underline = False
   m_Device.Font.Size = 12
   m_Device.CurrentX = (PLeft(12) - m_Device.TextWidth(stTmp)) / 2
   m_Device.CurrentY = iPrint
   m_Device.Print stTmp
   
   iPrint = iPrint + 300
   m_Device.CurrentX = 500
   m_Device.CurrentY = iPrint
   m_Device.Print "列印人：" & strUserName
   
   'Add by Morgan 2008/1/3
   If Val(txt1(1)) = 1 Then
      stTmp = "( 請款日有加 * 號表先請款後發文 )"
      m_Device.Font.Size = 10
      m_Device.CurrentX = (PLeft(12) - m_Device.TextWidth(stTmp)) / 2
      m_Device.CurrentY = iPrint
      m_Device.Print stTmp
   End If
   m_Device.Font.Size = 12
   
   m_Device.CurrentX = 13000
   m_Device.CurrentY = iPrint
   m_Device.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   iPrint = iPrint + 300
   
   If m_RptType = 0 Then
      m_Device.CurrentX = 500
      m_Device.CurrentY = iPrint
      'Modified by Lydia 2015/01/23
      If txt1(13) = "4" Then
         m_Device.Print "管制人：" & stName
      Else
         m_Device.Print "承辦人：" & stName
      End If
      
      m_Device.CurrentX = 4500
      m_Device.CurrentY = iPrint
      m_Device.Print "組別：" & stGrp
      
      m_Device.CurrentX = 8500
      m_Device.CurrentY = iPrint
      m_Device.Print "部門：" & stDep
      
   
   Else
      Select Case m_GrpType
         Case "1"
            m_Device.CurrentX = 500
            m_Device.CurrentY = iPrint
            m_Device.Print "組別：" & stGrp
            
            m_Device.CurrentX = 4500
            m_Device.CurrentY = iPrint
            m_Device.Print "部門：" & stDep
            
         Case "2"
            m_Device.CurrentX = 500
            m_Device.CurrentY = iPrint
            m_Device.Print "部門：" & stDep
            
         Case "3"
            m_Device.CurrentX = 500
            m_Device.CurrentY = iPrint
            m_Device.Print "全部"
      End Select

   End If
            
   m_Device.CurrentX = 13000
   m_Device.CurrentY = iPrint
   m_Device.Print "頁    次：" & str(Page)
   PrintLine 1
   
   iPrint = iPrint + 300
   '明細表
   If m_RptType = 0 Then
      PrintTitle3
   '統計表
   Else
      PrintTitle4
   End If
   PrintLine 1
End Sub

Sub PrintTitle3()
   m_Device.CurrentX = PLeft(1)
   m_Device.CurrentY = iPrint
   If Val(txt1(1)) = 1 Then
       m_Device.Print "請款日"
   Else
       m_Device.Print "發文日"
   End If
   m_Device.CurrentX = PLeft(2)
   m_Device.CurrentY = iPrint
   m_Device.Print "本所案號"
   m_Device.CurrentX = PLeft(3)
   m_Device.CurrentY = iPrint
   m_Device.Print "案件名稱"
   m_Device.CurrentX = PLeft(4)
   m_Device.CurrentY = iPrint
   m_Device.Print "專利種類"
   m_Device.CurrentX = PLeft(5)
   m_Device.CurrentY = iPrint
   m_Device.Print "案件性質"
   m_Device.CurrentX = PLeft(6)
   m_Device.CurrentY = iPrint
   m_Device.Print "本所期限"
   m_Device.CurrentX = PLeft(7)
   m_Device.CurrentY = iPrint
   If Val(txt1(1)) = 1 Then
       m_Device.Print "發文日"
   Else
       m_Device.Print "收文日"
   End If
   m_Device.CurrentX = PLeft(8)
   m_Device.CurrentY = iPrint
   m_Device.Print "工作時數"
   m_Device.CurrentX = PLeft(9)
   m_Device.CurrentY = iPrint
   m_Device.Print "承辦點數"
   m_Device.CurrentX = PLeft(10)
   m_Device.CurrentY = iPrint
   m_Device.Print "備註"
End Sub

Sub PrintTitle4()
   Select Case m_GrpType
      Case 1
         strExc(0) = "承辦人"
      Case 2
         strExc(0) = "組別"
      Case 3
         strExc(0) = "部門"
   End Select
   m_Device.CurrentX = PLeft(1)
   m_Device.CurrentY = iPrint
   m_Device.Print strExc(0)
   m_Device.CurrentX = PLeft(2) + 400
   m_Device.CurrentY = iPrint
   m_Device.Print "各項承辦案件總數"
   m_Device.CurrentX = PLeft(11)
   m_Device.CurrentY = iPrint
   m_Device.Print "工作總時數"
   m_Device.CurrentX = PLeft(12)
   m_Device.CurrentY = iPrint
   m_Device.Print "承辦總點數"
End Sub

Sub GetPleft1()
      
   Erase PLeft
   
   PLeft(0) = 500
   PLeft(1) = 500
   PLeft(2) = 1800
   PLeft(3) = 3800
   PLeft(4) = 6200
   PLeft(5) = 7600
   PLeft(6) = 9700
   PLeft(7) = 10900
   PLeft(8) = 12000
   PLeft(9) = 13100
   PLeft(10) = 14200
   PLeft(11) = 13400
   PLeft(12) = 15000
End Sub

''Add by Morgan 2010/4/15
''改抓分配點數
'Private Sub ProcessNew()
'   Dim stVTB As String
'   Dim stConPA As String, stConSP As String, stConLC As String
'   Dim stCon, stCon1K0 As String, stCon0K0 As String, stConCP As String, stCon1N0 As String
'   Dim stCon0 As String, stCon1 As String, stCon2 As String, stCon3 As String
'   Dim stSys As String, ii As Integer
'   Dim rsQuery As ADODB.Recordset, rsQuery1 As ADODB.Recordset
'   Dim intR As Integer, strKeyNow As String, strKeyLast As String
'   Dim strGrpNo As String 'Added by Morgan 2013/1/10 組別代碼
'
'   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/13 清除查詢印表記錄檔欄位
'
'   stCon0 = "": stCon1 = "": stCon2 = "": stCon3 = ""
'   stCon1K0 = "": stCon0K0 = "": stConCP = ""
'   stCon1N0 = ""
'
'   stSys = "'" & Join(Split(txt1(0), ","), "','") & "'"
'   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/13
'   If txt1(1) = "1" Then
'      stCon1K0 = stCon1K0 & " and a1k13||'' in (" & stSys & ")"
'      pub_QL05 = pub_QL05 & ";" & Left(Label1(1), 4) & "1.請款" 'Add By Sindy 2010/12/13
'   Else
'      pub_QL05 = pub_QL05 & ";" & Left(Label1(1), 4) & "2.發文" 'Add By Sindy 2010/12/13
'   End If
'   stCon = stCon & " and c1.cp01 in (" & stSys & ")"
'
'   '日期
'   If txt1(2) <> "" Then
'      '請款
'      If txt1(1) = "1" Then
'         stCon1K0 = stCon1K0 & " and a1k02>=" & txt1(2)
'         stCon0K0 = stCon0K0 & " and a0k02>=" & txt1(2)
'         stConCP = stConCP & " and c1.cp27>=" & DBDATE(txt1(2))
'      '發文
'      Else
'         stCon = stCon & " and c1.cp27>=" & DBDATE(txt1(2))
'      End If
'   End If
'   If txt1(3) <> "" Then
'      '請款
'      If txt1(1) = "1" Then
'         stCon1K0 = stCon1K0 & " and a1k02<=" & txt1(3)
'         stCon0K0 = stCon0K0 & " and a0k02<=" & txt1(3)
'         stConCP = stConCP & " and c1.cp27<=" & DBDATE(txt1(3))
'      '發文
'      Else
'         stCon = stCon & " and c1.cp27<=" & DBDATE(txt1(3))
'      End If
'   End If
'   If txt1(2) <> "" Or txt1(3) <> "" Then
'      If txt1(1) = "1" Then
'         pub_QL05 = pub_QL05 & ";請款" & Label1(2) & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/12/13
'      Else
'         pub_QL05 = pub_QL05 & ";發文" & Label1(2) & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/12/13
'      End If
'   End If
'
'   '案件性質
'   If txt1(4) <> "" Then
'      stCon = stCon & " and c1.cp10>='" & txt1(4) & "'"
'   End If
'   If txt1(5) <> "" Then
'      stCon = stCon & " and c1.cp10<='" & txt1(5) & "'"
'   End If
'   If txt1(4) <> "" Or txt1(5) <> "" Then
'      pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/12/13
'   End If
'
'   '承辦人
'   If txt1(6) <> "" Then
'      stCon1N0 = stCon1N0 & " and a1n04='" & txt1(6) & "'"
'      pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(6) & lbl1 'Add By Sindy 2010/12/13
'
'      '一般案件性質
'      stCon0 = stCon0 & " and c1.cp14='" & txt1(6) & "'"
'      '翻譯
'      strExc(0) = "select sim01,sim02 from staff_idmap where '" & txt1(6) & "' in (sim01,sim02)"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         stCon1 = stCon1 & " and c1.cp14 in ('" & RsTemp.Fields(0) & "','" & RsTemp.Fields(1) & "')"
'      Else
'         stCon1 = stCon1 & " and c1.cp14='" & txt1(6) & "'"
'      End If
'      '核稿
'      stCon2 = stCon2 & " and ep04='" & txt1(6) & "'"
'   End If
'
'   '申請人
'   If Len(Trim(txt1(7))) <> 0 And Len(Trim(txt1(8))) <> 0 Then
'      stConPA = stConPA & " AND ((PA26>='" & GetNewFagent(txt1(7)) & "' AND PA26<='" & GetNewFagent(txt1(8)) & "') OR (PA27>='" & GetNewFagent(txt1(7)) & "' AND PA27<='" & GetNewFagent(txt1(8)) & "') OR (PA28>='" & GetNewFagent(txt1(7)) & "' AND PA28<='" & GetNewFagent(txt1(8)) & "') OR (PA29>='" & GetNewFagent(txt1(7)) & "' AND PA29<='" & GetNewFagent(txt1(8)) & "') OR (PA30>='" & GetNewFagent(txt1(7)) & "' AND PA30<='" & GetNewFagent(txt1(8)) & "')) "
'      stConSP = stConSP & " AND ((SP08>='" & GetNewFagent(txt1(7)) & "' AND SP08<='" & GetNewFagent(txt1(8)) & "') OR (SP58<='" & GetNewFagent(txt1(7)) & "' AND SP58<='" & GetNewFagent(txt1(8)) & "') OR (SP59>='" & GetNewFagent(txt1(7)) & "' AND SP59<='" & GetNewFagent(txt1(8)) & "')) "
'      stConLC = stConLC & " AND ((LC11>='" & GetNewFagent(txt1(7)) & "' AND LC11<='" & GetNewFagent(txt1(8)) & "')) "
'   ElseIf Len(Trim(txt1(7))) <> 0 Then
'      stConPA = stConPA & " AND (PA26>='" & GetNewFagent(txt1(7)) & "' OR PA27>='" & GetNewFagent(txt1(7)) & "' OR PA28>='" & GetNewFagent(txt1(7)) & "' OR PA29>='" & GetNewFagent(txt1(7)) & "' OR PA30>='" & GetNewFagent(txt1(7)) & "') "
'      stConSP = stConSP & " AND NOT (SP08>='" & GetNewFagent(txt1(7)) & "' OR SP58>='" & GetNewFagent(txt1(7)) & "' OR SP59>='" & GetNewFagent(txt1(7)) & "') "
'      stConLC = stConLC & " AND ((LC11>='" & GetNewFagent(txt1(7)) & "')) "
'   ElseIf Len(Trim(txt1(8))) <> 0 Then
'      stConPA = stConPA & " AND NOT (PA26<='" & GetNewFagent(txt1(8)) & "' OR PA27<='" & GetNewFagent(txt1(8)) & "' OR PA28<='" & GetNewFagent(txt1(8)) & "' OR PA29<='" & GetNewFagent(txt1(8)) & "' OR PA30<='" & GetNewFagent(txt1(8)) & "') "
'      stConSP = stConSP & " AND NOT (SP08<='" & GetNewFagent(txt1(8)) & "' OR SP58<='" & GetNewFagent(txt1(8)) & "' OR SP59<='" & GetNewFagent(txt1(8)) & "') "
'      stConLC = stConLC & " AND ((LC11<='" & GetNewFagent(txt1(8)) & "')) "
'   End If
'   If Len(Trim(txt1(7))) <> 0 Or Len(Trim(txt1(8))) <> 0 Then
'      pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(7) & "-" & txt1(8) 'Add By Sindy 2010/12/13
'   End If
'
'   '代理人
'   If Len(Trim(txt1(9))) <> 0 And Len(Trim(txt1(10))) <> 0 Then
'       stConPA = stConPA & " AND PA75>='" & GetNewFagent(txt1(9)) & "' AND PA75<='" & GetNewFagent(txt1(10)) & "' "
'       stConSP = stConSP & " AND SP26>='" & GetNewFagent(txt1(9)) & "' AND SP26<='" & GetNewFagent(txt1(10)) & "' "
'       stConLC = stConLC & " AND LC22>='" & GetNewFagent(txt1(9)) & "' AND LC22<='" & GetNewFagent(txt1(10)) & "' "
'   ElseIf Len(Trim(txt1(9))) <> 0 Then
'       stConPA = stConPA & " AND PA75>='" & GetNewFagent(txt1(9)) & "' "
'       stConSP = stConSP & " AND SP26>='" & GetNewFagent(txt1(9)) & "' "
'       stConLC = stConLC & " AND LC22>='" & GetNewFagent(txt1(9)) & "' "
'   ElseIf Len(Trim(txt1(10))) <> 0 Then
'       stConPA = stConPA & " AND PA75<='" & GetNewFagent(txt1(10)) & "' "
'       stConSP = stConSP & " AND SP26<='" & GetNewFagent(txt1(10)) & "' "
'       stConLC = stConLC & " AND LC22<='" & GetNewFagent(txt1(10)) & "' "
'   End If
'   If Len(Trim(txt1(9))) <> 0 Or Len(Trim(txt1(10))) <> 0 Then
'      pub_QL05 = pub_QL05 & ";" & Label1(6) & txt1(9) & "-" & txt1(10) 'Add By Sindy 2010/12/13
'   End If
'
'   '組別
'   If txt1(11) <> "" Then
'       stCon3 = stCon3 & " AND ST16='" & txt1(11) & "'"
'       pub_QL05 = pub_QL05 & ";" & Left(Label1(7), 3) & txt1(11) & "( 1.電子電機 2.化學 3.日文 4.機械設計 5.其他 )" 'Add By Sindy 2010/12/13
'   End If
'
'   '工程師
'   If txt1(13) = "1" Then
'      stCon3 = stCon3 & " AND ST15='F21'"
'      pub_QL05 = pub_QL05 & ";" & Left(Label1(12), 5) & "1.工程師" 'Add By Sindy 2010/12/13
'   '程序
'   ElseIf txt1(13) = "2" Then
'      stCon3 = stCon3 & " AND ST15='F22'"
'      pub_QL05 = pub_QL05 & ";" & Left(Label1(12), 5) & "2.程序人員" 'Add By Sindy 2010/12/13
'   '工程師+程序
'   Else
'      stCon3 = stCon3 & " AND ST15 IN ('F21','F22')"
'      pub_QL05 = pub_QL05 & ";" & Left(Label1(12), 5) & "3.所有人" 'Add By Sindy 2010/12/13
'   End If
'
'   If txt1(12) = "1" Then
'      pub_QL05 = pub_QL05 & ";" & Left(Label1(8), 4) & "1.明細" 'Add By Sindy 2010/12/13
'   Else
'      pub_QL05 = pub_QL05 & ";" & Left(Label1(8), 4) & "2.統計" 'Add By Sindy 2010/12/13
'   End If
'   If txt1(14) = "1" Then
'      pub_QL05 = pub_QL05 & ";" & Left(Label1(9), 5) & "1.螢幕" 'Add By Sindy 2010/12/13
'   Else
'      pub_QL05 = pub_QL05 & ";" & Left(Label1(9), 5) & "2.印表機" 'Add By Sindy 2010/12/13
'   End If
'
'   'Add by Morgan 2010/4/20
'   If stConPA <> "" Then
'      stCon1N0 = stCon1N0 & " and (exists(select * from patent where pa01=c1.cp01 and pa02=c1.cp02 and pa03=c1.cp03 and pa04=c1.cp04 " & stConPA & ")" & _
'         " or exists(select * from servicepractice where sp01=c1.cp01 and sp02=c1.cp02 and sp03=c1.cp03 and sp04=c1.cp04 " & stConSP & ")" & _
'         " or exists(select * from lawcase where lc01=c1.cp01 and lc02=c1.cp02 and lc03=c1.cp03 and lc04=c1.cp04 " & stConLC & "))"
'   End If
'   If stCon3 <> "" Then
'      stCon1N0 = stCon1N0 & " and exists(select * from staff where st01=a1n04 " & stCon3 & ")"
'   End If
'
'   'Add by Morgan 2010/4/19 分配點數語法
'   '請款
'   If txt1(1) = "1" Then
'      '先請款後發文
'      m_strSharePointVTB = "select '" & strUserNum & "' id,a1n04,a1n03,a1n05,cp01,cp02,cp03,cp04,cp10,a1k02,'*' Flg,a1n06" & _
'         " From caseprogress c1,acc1k0,acc1n0,engineerprogress" & _
'         " Where cp60>'X' and cp27>0 and cp14 is not null" & stConCP & stCon & _
'         " and a1k01(+)=cp60 and a1k02+19110000<cp27 and a1k12 is null and a1k25 is null" & _
'         " and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n05>0 and a1n04<>cp14" & _
'         " and ep02(+)=cp09 and a1n04||a1n06<>ep04||'Y'" & stCon1N0
'
'      '先發文後請款
'      m_strSharePointVTB = m_strSharePointVTB & _
'         " union select '" & strUserNum & "' id,a1n04,a1n03,a1n05,cp01,cp02,cp03,cp04,cp10,a1k02,'' Flg,a1n06" & _
'         " From acc1k0,caseprogress c1,acc1n0,engineerprogress" & _
'         " Where a1k12 Is Null and a1k25 is null" & stCon1K0 & _
'         " and cp60(+)=a1k01 and cp27<=a1k02+19110000 and cp14 is not null" & stCon & _
'         " and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n05>0 and a1n04<>cp14" & _
'         " and ep02(+)=cp09 and a1n04||a1n06<>ep04||'Y'" & stCon1N0
'   '發文
'   Else
'      m_strSharePointVTB = "select '" & strUserNum & "' id,a1n04,a1n03,a1n05,cp01,cp02,cp03,cp04,cp10,cp27,'' Flg,a1n06" & _
'         " From caseprogress c1,acc1n0,engineerprogress,acc1k0" & _
'         " Where cp60>'X' and cp27>0 and cp14 is not null" & stConCP & stCon & _
'         " and a1k01(+)=cp60 and a1k25 is null and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n05>0" & _
'         " and ep02(+)=cp09 and a1n04<>cp14 and a1n04||a1n06<>ep04||'Y'" & stCon1N0
'   End If
'
'   'Modifyby Morgan 2010/12/22 +重新核稿 229
'   '新案翻譯除外
'   '請款日(若發文前請款案件則以發文日為日期條件)
'   If txt1(1) = "1" Then
'      '請款單:先發文後請款案件
'      stVTB = "select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",decode(a1k25,null,a1n05) C9" & _
'         " from acc1k0,caseprogress c1,acc1n0 where a1k12 is null" & stCon1K0 & _
'         " and cp60(+)=a1k01 and cp27<=a1k02+19110000 and cp14 is not null and cp10<>'201'" & stCon & stCon0 & _
'         " and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n04(+)=cp14 and a1n06(+) is null"
'
'      '請款單:先請款後發文案件
'      stVTB = stVTB & " union all select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",decode(a1k25,null,a1n05) C9" & _
'         " from caseprogress c1,acc1k0,acc1n0 where cp60>'X'" & stConCP & _
'         " and cp27>0 and cp14 is not null" & _
'         " and a1k01(+)=cp60 and a1k02+19110000<cp27 and a1k12 is null and cp10<>'201'" & stCon & stCon0 & _
'         " and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n04(+)=cp14 and a1n06(+) is null"
'
'      '承辦點數 = 業務收文點數
'      '收據:先發文後請款案件
'      stVTB = stVTB & " union all select a0k02 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",cp18 C9" & _
'         " from acc0k0,caseprogress c1 where nvl(a0k09,0)=0" & stCon0K0 & _
'         " and cp60(+)=a0k01 and cp27<=a0k02+19110000 and cp14 is not null and cp10<>'201'" & stCon & stCon0
'
'      '收據:先請款後發文案件
'      stVTB = stVTB & " union all select a0k02 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",cp18 C9" & _
'         " from caseprogress c1,acc0k0 where cp60<'X'" & stConCP & _
'         " and cp27>0 and cp14 is not null and cp10<>'201'" & _
'         " and a0k01(+)=cp60 and a0k02+19110000<cp27 and nvl(a0k09,0)=0" & stCon & stCon0
'
'
'      'Add by Morgan 2010/12/22
'      '重新核稿
'      '翻譯為請款單或未請款
'      stVTB = stVTB & " union all select c1.cp27 as a1k02,c1.cp01,c1.cp02,c1.cp03,c1.cp04,c1.cp05,c1.cp06,c1.cp10,c1.cp12,c1.cp14,c1.cp18,c1.cp27,c1.cp60,c1.cp64,c1.cp113" & _
'         ",round((a1l05-nvl(a1l07,0))*0.3/1000,2) C9" & _
'         " from caseprogress c1,caseprogress c2,acc1l0 where c1.cp10='229'" & stConCP & _
'         " and c1.cp27>0 and c1.cp14 is not null" & stCon & stCon0 & _
'         " and c2.cp01(+)=c1.cp01 and c2.cp02(+)=c1.cp02 and c2.cp03(+)=c1.cp03 and c2.cp04(+)=c1.cp04" & _
'         " and c2.cp10(+)='201' and (c2.cp60 is null or c2.cp60>'X')" & _
'         " and a1l01(+)=c2.cp60 and a1l04(+)=c2.cp10"
'
'      '收據
'      stVTB = stVTB & " union all select c1.cp27 as a1k02,c1.cp01,c1.cp02,c1.cp03,c1.cp04,c1.cp05,c1.cp06,c1.cp10,c1.cp12,c1.cp14,c1.cp18,c1.cp27,c1.cp60,c1.cp64,c1.cp113" & _
'         ",round(nvl(c2.cp18,0)*0.3,2) C9" & _
'         " from caseprogress c1,caseprogress c2 where c1.cp10='229'" & stConCP & _
'         " and c1.cp27>0 and c1.cp14 is not null" & stCon & stCon0 & _
'         " and c2.cp01(+)=c1.cp01 and c2.cp02(+)=c1.cp02 and c2.cp03(+)=c1.cp03 and c2.cp04(+)=c1.cp04" & _
'         " and c2.cp10(+)='201' and c2.cp60<'X'"
'
'   '發文日
'   Else
'      'Modify by Morgan 2010/12/22 +排除重新核稿 229
'      stVTB = "select cp27 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp18,cp05 as cp27,cp60,cp64,cp113" & _
'         ",decode(substr(cp60,1,1),'E',cp18,'X',decode(a1k25,null,a1n05)) C9" & _
'         " from caseprogress c1,acc1n0,acc1k0 where cp14 is not null and cp10<>'201' and cp10<>'229'" & stCon & stCon0 & _
'         " and a1k01(+)=cp60 and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n04(+)=cp14 and a1n06(+) is null"
'
'      'Add by Morgan 2010/12/22
'      '重新核稿 229
'      stVTB = stVTB & " union all select c1.cp27 as a1k02,c1.cp01,c1.cp02,c1.cp03,c1.cp04,c1.cp05,c1.cp06,c1.cp10,c1.cp12,c1.cp14,c1.cp18,c1.cp05 as cp27,c1.cp60,c1.cp64,c1.cp113" & _
'         ",round(decode(substr(c2.cp60,1,1),'E',c2.cp18,(a1l05-nvl(a1l07,0))/1000)*0.3,2) C9" & _
'         " from caseprogress c1,caseprogress c2,acc1l0 where c1.cp14 is not null and c1.cp10='229'" & stCon & stCon0 & _
'         " and c2.cp01(+)=c1.cp01 and c2.cp02(+)=c1.cp02 and c2.cp03(+)=c1.cp03 and c2.cp04(+)=c1.cp04" & _
'         " and c2.cp10(+)='201' and a1l01(+)=c2.cp60 and a1l04(+)=c2.cp10"
'
'   End If
'   'Modify by Morgan 2010/8/16 百年蟲 sqldatet(a1k02)-->substrb(' '||sqldatet(a1k02),-9)
'   strExc(1) = "select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
'      ",substrb(pa05,1,16) C3,substrb(ptm03,1,6) C4,substrb(decode(pa09,'020',cpm04,cpm03),1,8) C5" & _
'      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
'      ",cp113 C8, C9,substrb(cp64,1,16) C10" & _
'      ",st01 C11,A0902 C12,ST16 C13,st02 C14" & _
'      " From (" & stVTB & ") X,patent, patenttrademarkmap, casepropertymap,staff,acc090" & _
'      " where  pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA & _
'      " and ptm01(+)=1 and ptm02(+)=pa08 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
'      " and st01(+)=cp14 and a0901(+)=ST15" & stCon3
'
'   strExc(1) = strExc(1) & " union all select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
'      ",substrb(sp05,1,16) C3,'' C4,substrb(decode(sp09,'020',cpm04,cpm03),1,8) C5" & _
'      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
'      ",cp113 C8, C9,substrb(cp64,1,16) C10" & _
'      ",st01 C11,A0902 C12,ST16 C13,st02 C14" & _
'      " From (" & stVTB & ") X,servicepractice,casepropertymap,staff,acc090" & _
'      " where  sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP & _
'      " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
'      " and st01(+)=cp14 and a0901(+)=ST15" & stCon3
'
'   strExc(1) = strExc(1) & " union all select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
'      ",substrb(LC05,1,16) C3,'' C4,substrb(cpm03,1,8) C5" & _
'      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
'      ",cp113 C8, C9,substrb(cp64,1,16) C10" & _
'      ",st01 C11,A0902 C12,ST16 C13,st02 C14" & _
'      " From (" & stVTB & ") X,lawcase,casepropertymap,staff,acc090" & _
'      " where  lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04 and lc01 is not null" & stConLC & _
'      " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
'      " and st01(+)=cp14 and a0901(+)=ST15" & stCon3
'
'   '新案翻譯(內翻的外譯編號要轉成所內員工編號),原則上所內編號(上班翻)的才會有承辦點數
'   'Modify by Morgan 2011/7/5 排除承辦人為外翻部門的(否則已離職員工也會印出)
'   '請款日
'   If txt1(1) = "1" Then
'      '請款單:先發文後請款案件
'      stVTB = "select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,nvl(sim01,cp14) as cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",decode(a1k25,null,a1n05) C9" & _
'         " from acc1k0,caseprogress c1,staff,acc1n0,staff_idmap where a1k12 is null" & stCon1K0 & _
'         " and cp60(+)=a1k01 and cp27<=a1k02+19110000 and cp14 is not null and cp10='201'" & stCon & stCon1 & _
'         " and st01(+)=cp14 and st03<>'F51' and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n04(+)=cp14 and a1n06(+) is null" & _
'         " and sim02(+)=cp14"
'
'      '請款單:先請款後發文案件
'      stVTB = stVTB & " union all select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,nvl(sim01,cp14) as cp14,cp18,cp27,cp60,cp64,cp113" & _
'         ",decode(a1k25,null,a1n05) C9" & _
'         " from caseprogress c1,staff,acc1k0,acc1n0,staff_idmap where cp60>'X'" & stConCP & _
'         " and cp27>0 and cp14 is not null and st01(+)=cp14 and st03<>'F51'" & _
'         " and a1k01(+)=cp60 and a1k02+19110000<cp27 and a1k12 is null and cp10='201'" & stCon & stCon1 & _
'         " and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n04(+)=cp14 and a1n06(+) is null" & _
'         " and sim02(+)=cp14"
'
'      '內翻翻譯承辦點數=智權人員收文點數-核稿點數(核稿人不是承辦人才會有),否則=0
'      '收據:先發文後請款案件
'      stVTB = stVTB & " union all select a0k02 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,nvl(sim01,cp14) as cp14" & _
'         ",cp18,cp27,cp60,cp64,cp113,decode(substrb(cp14,1,1),'F',0,cp18*decode(cp14,ep04,1,0.7*nvl(TF06,100)/100)) C9" & _
'         " from acc0k0,caseprogress c1,staff,staff_idmap,engineerprogress,transfee where nvl(a0k09,0)=0" & stCon0K0 & _
'         " and cp60(+)=a0k01 and cp27<=a0k02+19110000 and cp14 is not null and cp10='201'" & stCon & stCon1 & _
'         " and st01(+)=cp14 and st03<>'F51' and sim02(+)=cp14 and ep02(+)=cp09 and tf01(+)=cp09"
'
'      '收據:先請款後發文案件
'      stVTB = stVTB & " union all select a0k02 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,nvl(sim01,cp14) as cp14" & _
'         ",cp18,cp27,cp60,cp64,cp113,decode(substrb(cp14,1,1),'F',0,cp18*decode(cp14,ep04,1,0.7*nvl(TF06,100)/100)) C9" & _
'         " from caseprogress c1,staff,acc0k0,staff_idmap,engineerprogress,transfee where cp60<'X'" & stConCP & _
'         " and cp27>0 and cp14 is not null and cp10='201' and st01(+)=cp14 and st03<>'F51'" & _
'         " and a0k01(+)=cp60 and a0k02+19110000<cp27 and nvl(a0k09,0)=0" & stCon & stCon1 & _
'         " and sim02(+)=cp14 and ep02(+)=cp09 and tf01(+)=cp09"
'
'   '發文日
'   Else
'      '請款單或未請款
'      stVTB = "select cp27 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,nvl(sim01,cp14) as cp14,cp18,cp05 as cp27,cp60,cp64,cp113" & _
'         ",decode(a1k25,null,a1n05) C9" & _
'         " from caseprogress c1,staff,acc1k0,acc1n0,staff_idmap where (cp60 is null or cp60>'X') and cp14 is not null and cp10='201'" & stCon & stCon1 & _
'         " and st01(+)=cp14 and st03<>'F51' and a1k01(+)=cp60 and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n04(+)=cp14 and a1n06(+) is null" & _
'         " and sim02(+)=cp14"
'
'      '收據
'      stVTB = stVTB & " union all select cp27 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,nvl(sim01,cp14) as cp14,cp18,cp05 as cp27,cp60,cp64,cp113" & _
'         ",decode(substrb(cp14,1,1),'F',0,cp18*decode(cp14,ep04,1,0.7*nvl(TF06,100)/100)) C9" & _
'         " from caseprogress c1,staff,staff_idmap,engineerprogress,transfee where cp60<'X' and cp14 is not null and cp10='201'" & stCon & stCon1 & _
'         " and st01(+)=cp14 and st03<>'F51' and sim02(+)=cp14 and ep02(+)=cp09 and tf01(+)=cp09"
'   End If
'
'   strExc(2) = " union all select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
'      ",substrb(pa05,1,16) C3,substrb(ptm03,1,6) C4,substrb(decode(pa09,'020',cpm04,cpm03),1,8) C5" & _
'      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
'      ",cp113 C8,C9,substrb(cp64,1,16) C10,st01 C11,A0902 C12,ST16 C13,st02 C14" & _
'      " From (" & stVTB & ") X,patent, patenttrademarkmap, casepropertymap,staff,acc090" & _
'      " where pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA & _
'      " and ptm01(+)=1 and ptm02(+)=pa08 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
'      " and st01(+)=cp14 and a0901(+)=ST15" & stCon3
'
'   strExc(2) = strExc(2) & " union all select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
'      ",substrb(sp05,1,16) C3,'' C4,substrb(decode(sp09,'020',cpm04,cpm03),1,8) C5" & _
'      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
'      ",cp113 C8,C9,substrb(cp64,1,16) C10" & _
'      ",st01 C11,A0902 C12,ST16 C13,st02 C14" & _
'      " From (" & stVTB & ") X,servicepractice,casepropertymap,staff,acc090" & _
'      " where  sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP & _
'      " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
'      " and st01(+)=cp14 and a0901(+)=ST15" & stCon3
'
'   '核稿
'   '核稿點數=翻譯請款點數(扣除折扣)x(30%+因翻譯瑕疵扣減支付外譯人員翻譯費之百分比x70%)
'   '請款日
'   'Modified by Morgan 2013/5/28 +927其他翻譯 Ex.FG-000858
'   If txt1(1) = "1" Then
'      '請款單:
'      stVTB = "select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,ep04,cp18,cp27,cp60,cp64,cp114" & _
'         ",decode(a1k25,null,a1n05) C9" & _
'         " from acc1k0,caseprogress c1,engineerprogress,acc1n0 where a1k12 is null" & stCon1K0 & _
'         " and cp60(+)=a1k01 and cp27<=a1k02+19110000 and cp14 is not null and cp10 in ('201','927')" & stCon & stCon2 & _
'         " and ep02(+)=cp09 and ep04<>cp14" & _
'         " and a1n02(+)='2' and a1n03(+)=ep02 and a1n04(+)=ep04 and a1n06(+)='Y'"
'
'      '請款單:先請款後發文案件
'      stVTB = stVTB & " union all select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,ep04,cp18,cp27,cp60,cp64,cp114" & _
'         ",decode(a1k25,null,a1n05) C9" & _
'         " from caseprogress c1,acc1k0,engineerprogress,acc1n0" & _
'         " where cp60>'X'" & stConCP & _
'         " and cp27>0 and cp14 is not null and cp10 in ('201','927')" & _
'         " and a1k01(+)=cp60 and a1k02+19110000<cp27 and a1k12 is null" & stCon & stCon2 & _
'         " and ep02(+)=cp09 and ep04<>cp14" & _
'         " and a1n02(+)='2' and a1n03(+)=ep02 and a1n04(+)=ep04 and a1n06(+)='Y'"
'
'      '收據:核稿點數=智權人員收文點數x(30%+因翻譯瑕疵扣減支付外譯人員翻譯費之百分比x70%)
'      stVTB = stVTB & " union all select a0k02 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,ep04,cp18,cp27,cp60,cp64,cp114" & _
'         ",cp18*(0.3+0.7*(1-nvl(TF06,100)/100)) C9" & _
'         " from acc0k0,caseprogress c1,engineerprogress,transfee where nvl(a0k09,0)=0" & stCon0K0 & _
'         " and cp60(+)=a0k01 and cp27<=a0k02+19110000 and cp14 is not null and cp10='201'" & stCon & stCon2 & _
'         " and ep02(+)=cp09 and ep04<>cp14 and tf01(+)=cp09"
'
'      '收據:先請款後分案案件
'      stVTB = stVTB & " union all select a0k02 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,ep04,cp18,cp27,cp60,cp64,cp114" & _
'         ",cp18*(0.3+0.7*(1-nvl(TF06,100)/100)) C9" & _
'         " from caseprogress c1,acc0k0,engineerprogress,transfee" & _
'         " where cp60<'X'" & stConCP & _
'         " and cp27>0 and cp14 is not null and cp10='201'" & _
'         " and a0k01(+)=cp60 and a0k02+19110000<cp27 and nvl(a0k09,0)=0" & stCon & stCon2 & _
'         " and ep02(+)=cp09 and ep04<>cp14 and tf01(+)=cp09"
'
'   '發文日
'   Else
'      '請款單
'      stVTB = "select cp27 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,ep04,cp18,cp05 as cp27,cp60,cp64,cp114" & _
'         ",decode(a1k25,null,a1n05) C9" & _
'         " from caseprogress c1,engineerprogress,acc1k0,acc1n0 where cp14 is not null and cp10 in ('201','927')" & stCon & stCon2 & _
'         " and (cp60 is null or cp60>'X') and ep02(+)=cp09 and ep04<>cp14" & _
'         " and a1k01(+)=cp60 and a1n02(+)='2' and a1n03(+)=ep02 and a1n04(+)=ep04 and a1n06(+)='Y'"
'      '收據
'      stVTB = stVTB & " union all select cp27 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,ep04,cp18,cp05 as cp27,cp60,cp64,cp114" & _
'         ",cp18*(0.3+0.7*(1-nvl(TF06,100)/100)) C9" & _
'         " from caseprogress c1,engineerprogress,transfee where cp14 is not null and cp10='201'" & stCon & stCon2 & _
'         " and cp60<'X' and ep02(+)=cp09 and ep04<>cp14" & _
'         " and tf01(+)=cp09"
'   End If
'
'   strExc(3) = " union all select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
'      ",substrb(pa05,1,16) C3,substrb(ptm03,1,6) C4,substrb(decode(pa09,'020',cpm04,cpm03),1,8)||'-核稿' C5" & _
'      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
'      ",cp114 C8,C9,substrb(cp64,1,16) C10" & _
'      ",st01 C11,A0902 C12,ST16 C13,st02 C14" & _
'      " From (" & stVTB & ") X,patent,patenttrademarkmap,casepropertymap,staff,acc090" & _
'      " where pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA & _
'      " and ptm01(+)=1 and ptm02(+)=pa08 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
'      " and st01(+)=ep04 and a0901(+)=ST15" & stCon3
'
'   strExc(3) = strExc(3) & " union all select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
'      ",substrb(sp05,1,16) C3,'' C4,substrb(decode(sp09,'020',cpm04,cpm03),1,8)||'-核稿' C5" & _
'      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
'      ",cp114 C8,C9,substrb(cp64,1,16) C10" & _
'      ",st01 C11,A0902 C12,ST16 C13,st02 C14" & _
'      " From (" & stVTB & ") X, servicepractice,casepropertymap,staff,acc090" & _
'      " where  sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP & _
'      " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
'      " and st01(+)=ep04 and a0901(+)=ST15" & stCon3
'
'   strExc(0) = strExc(1) & strExc(2) & strExc(3) & " order by C12,C13,C11,C1,C2,C5"
'
'   stName = "": stGrp = "": stDep = "": stID = ""
'   dblPoint = 0: dblTotPoint = 0
'   dblPoint2 = 0: dblTotPoint2 = 0 'Add by Morgan 2010/11/15
'   dblHour = 0: dblTotHour = 0
'   m_RptType = 0: m_GrpType = 0
'   Page = 0: m_iPages = 0
'   iPrint = 0
'   strKeyNow = "": strKeyLast = ""
'
'   cnnConnection.Execute "DELETE FROM R060312_1 WHERE ID='" & strUserNum & "'"
'   strSql = "INSERT INTO R060312_1(ID,R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11)" & m_strSharePointVTB
'   cnnConnection.Execute strSql, intR
'
'   intI = 1
'   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
'   If intI <> 1 Then
'      '跑明細報表且有分配點數資料
'      If intR > 0 And txt1(12) = "1" Then
'         strExc(0) = "select distinct st01 C11,A0902 C12,ST16 C13,st02 C14 from R060312_1,staff,acc090 where ID='" & strUserNum & "' and st01(+)=R01 and a0901(+)=ST15 order by C12,C13,C11"
'         intI = 1
'         Set rsQuery1 = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            If txt1(14) = "1" Then
'               m_bPrinter = False
'               Set m_Device = Picture1
'               m_Device.AutoRedraw = True
'               m_Device.Width = 16836
'               m_Device.Height = 11904
'               DelPic
'            Else
'               m_bPrinter = True
'               Set m_Device = Printer
'               m_Device.Orientation = 2
'            End If
'
'            GetPleft1
'            With rsQuery1
'            InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/13
'            .MoveFirst
'            Do While Not .EOF
'               stID = "" & .Fields("C11")
'               stName = "" & .Fields("C14")
'               stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
'               stDep = "" & .Fields("C12")
'               If stGrp = "" Then stGrp = stDep
'               PrintTitle2
'               iPrint = iPrint + 300
'               PrintSharePoint stID
'               .MoveNext
'            Loop
'            End With
'            If m_bPrinter = True Then
'               m_Device.EndDoc
'               ShowPrintOk
'            ElseIf m_iPages > 0 Then
'               SetPic m_iPages
'               frm060312_1.m_ImageW = m_Device.Width
'               frm060312_1.m_ImageH = m_Device.Height
'               frm060312_1.m_iPages = m_iPages
'               frm060312_1.Show
'            End If
'         End If
'      Else
'         InsertQueryLog (0) 'Add By Sindy 2010/12/13
'         MsgBox "無可列印資料！"
'      End If
'   Else
'      cnnConnection.Execute "DELETE FROM R060312 WHERE ID='" & strUserNum & "' "
'      If txt1(14) = "1" Then
'         m_bPrinter = False
'         Set m_Device = Picture1
'         m_Device.AutoRedraw = True
'         m_Device.Width = 16836
'         m_Device.Height = 11904
'         DelPic
'      Else
'         m_bPrinter = True
'         Set m_Device = Printer
'         m_Device.Orientation = 2
'      End If
'
'      GetPleft1
'      With rsQuery
'      InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/13
'      .MoveFirst
'      Do While Not .EOF
'         '列印明細小計
'         If txt1(12) = "1" Then
'            If stID <> "" & .Fields("C11") Then
'               If stID <> "" Then
'                  PrintSubTot stID
'               End If
'            End If
'         End If
'
'         'Modify by Morgan 2011/4/12 本來只考慮明細,改統計也要做
'         If stID <> "" & .Fields("C11") Then
'            'Add by Morgan 2010/4/20 檢查是否有無資料但有分配點數的情形
'            If intR > 0 Then
'               strKeyNow = "" & .Fields("C12") & .Fields("C13") & .Fields("C11")
'               strExc(0) = "select distinct st01 C11,A0902 C12,ST16 C13,st02 C14 from R060312_1,staff,acc090 where ID='" & strUserNum & "' and st01(+)=R01 and a0901(+)=ST15 and A0902||ST16||st01>'" & strKeyLast & " ' and A0902||ST16||st01<'" & strKeyNow & "' order by C12,C13,C11"
'               intI = 1
'               Set rsQuery1 = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  With rsQuery1
'                     Do While Not .EOF
'
'                        If txt1(12) = "1" Then
'                           stID = "" & .Fields("C11")
'                           stName = "" & .Fields("C14")
'                           stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
'                           stDep = "" & .Fields("C12")
'                           If stGrp = "" Then stGrp = stDep
'                           PrintTitle2
'                           iPrint = iPrint + 300
'                           PrintSharePoint stID
'
'                        'Add by Morgan 2011/4/12
'                        Else
'                           stID = "" & .Fields("C11")
'                           stName = "" & .Fields("C14")
'                           strGrpNo = "" & .Fields("C13") 'Added by Morgan 2013/1/10
'                           stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
'                           stDep = "" & .Fields("C12")
'                           If stGrp = "" Then stGrp = stDep
'                           strSql = " INSERT INTO R060312 (ID, R045001, R045003, R045005, R045008, R045004, R045006, R045007,R045014) values('" & strUserNum & "','" & stID & "','',0,0,'" & stDep & "','" & stGrp & "','" & stName & "','" & strGrpNo & "')"
'                           cnnConnection.Execute strSql, intI
'
'                        End If
'                        .MoveNext
'                     Loop
'                  End With
'               End If
'               strKeyLast = strKeyNow
'            End If
'            'end 2010/4/20
'         End If
'
'
'         If txt1(12) = "2" Then
'
'            stID = "" & .Fields("C11")
'            stName = "" & .Fields("C14")
'            strGrpNo = "" & .Fields("C13") 'Added by Morgan 2013/1/10
'            '2010/1/8 MODIFY BY SONIA
'            'stGrp = "" & .Fields("C13")
'            stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
'            '2010/1/8 END
'            stDep = "" & .Fields("C12")
'            'Add by Morgan 2007/8/23 若沒組別時設為部門
'            If stGrp = "" Then stGrp = stDep
'         Else
'
'            If stID <> "" & .Fields("C11") Then
'
'               dblPoint = 0: dblHour = 0
'               dblPoint2 = 0 'Add by Morgan 2010/11/15
'               stID = "" & .Fields("C11")
'               stName = "" & .Fields("C14")
'               '2010/1/8 MODIFY BY SONIA
'               'stGrp = "" & .Fields("C13")
'               stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
'               '2010/1/8 END
'               stDep = "" & .Fields("C12")
'               'Add by Morgan 2007/8/23 若沒組別時設為部門
'               If stGrp = "" Then stGrp = stDep
'               PrintTitle2
'               iPrint = iPrint + 300
'
'            Else
'               iPrint = iPrint + 400
'               If iPrint > m_Device.ScaleHeight - 600 Then
'                  PrintTitle2
'                  iPrint = iPrint + 300
'               End If
'            End If
'
'            For ii = 1 To 10
'               strExc(0) = ""
'               Select Case ii
'                  Case 8
'                     If "" & .Fields("C" & ii) <> "" Then
'                        strExc(0) = Format("" & .Fields("C" & ii), "0.0")
'                     End If
'                     m_Device.CurrentX = PLeft(ii) + 960 - m_Device.TextWidth(strExc(0))
'                     dblHour = dblHour + Val(strExc(0))
'                     dblTotHour = dblTotHour + Val(strExc(0))
'                  Case 9
'                     If "" & .Fields("C" & ii) <> "" Then
'                        strExc(0) = Format("" & .Fields("C" & ii), "0.00")
'                     End If
'                     m_Device.CurrentX = PLeft(ii) + 960 - m_Device.TextWidth(strExc(0))
'                     dblPoint = dblPoint + Val(strExc(0))
'                     dblTotPoint = dblTotPoint + Val(strExc(0))
'                     'Add by Morgan 2010/11/15
'                     If InStr("" & .Fields("C5"), "核稿") > 0 Then
'                        dblPoint2 = dblPoint2 + Val(strExc(0))
'                        dblTotPoint2 = dblTotPoint2 + Val(strExc(0))
'                     End If
'                  Case Else
'                     strExc(0) = "" & .Fields("C" & ii)
'                     'Add by Morgan 2008/1/3
'                     If ii = 1 And Right("0" & .Fields("C1"), 9) < Right("0" & .Fields("C7"), 9) Then
'                        strExc(0) = strExc(0) & "*"
'                     End If
'                     'end 2008/1/3
'                     m_Device.CurrentX = PLeft(ii)
'               End Select
'               m_Device.CurrentY = iPrint
'               'Modify by Morgan 2010/11/15
'               'm_Device.Print strExc(0)
'               If InStr(strExc(0), "核稿") > 0 Then
'                  m_Device.FontBold = True
'                  m_Device.Font.Underline = True
'                  m_Device.Print strExc(0)
'                  m_Device.Font.Underline = False
'                  m_Device.FontBold = False
'               Else
'                  m_Device.Print strExc(0)
'               End If
'               'end 2010/11/15
'            Next
'         End If
'         strSql = " INSERT INTO R060312 (ID, R045001, R045003, R045005, R045008, R045004, R045006, R045007, R045014) values('" & strUserNum & "','" & stID & "','" & .Fields("C5") & "'," & Format(Val("" & .Fields("C8")), "0.0") & "," & Format(Val("" & .Fields("C9")), "0.00") & ",'" & stDep & "','" & stGrp & "','" & stName & "','" & strGrpNo & "')"
'         cnnConnection.Execute strSql, intI
'         .MoveNext
'      Loop
'
'      If txt1(12) = "1" Then
'         PrintSubTot stID
'      End If
'
'      'Modify by Morgan 2011/4/12 本來只考慮明細,改統計也要做
'      'If txt1(12) = "1" Then
'         'Add by Morgan 2010/4/20 檢查是否有無資料但有分配點數的情形
'         If intR > 0 Then
'            strExc(0) = "select distinct st01 C11,A0902 C12,ST16 C13,st02 C14 from R060312_1,staff,acc090 where ID='" & strUserNum & "' and st01(+)=R01 and a0901(+)=ST15 and A0902||ST16||st01>'" & strKeyLast & " ' order by C12,C13,C11"
'            intI = 1
'            Set rsQuery1 = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               With rsQuery1
'                  Do While Not .EOF
'                     If txt1(12) = "1" Then
'                        stID = "" & .Fields("C11")
'                        stName = "" & .Fields("C14")
'                        stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
'                        stDep = "" & .Fields("C12")
'                        If stGrp = "" Then stGrp = stDep
'                        PrintTitle2
'                        iPrint = iPrint + 300
'                        PrintSharePoint stID
'
'                     'Add by Morgan 2011/4/12
'                     Else
'                        stID = "" & .Fields("C11")
'                        stName = "" & .Fields("C14")
'                        strGrpNo = "" & .Fields("C13") 'Added by Morgan 2013/1/10
'                        stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
'                        stDep = "" & .Fields("C12")
'                        If stGrp = "" Then stGrp = stDep
'                        strSql = " INSERT INTO R060312 (ID, R045001, R045003, R045005, R045008, R045004, R045006, R045007,R045014) values('" & strUserNum & "','" & stID & "','',0,0,'" & stDep & "','" & stGrp & "','" & stName & "','" & strGrpNo & "')"
'                        cnnConnection.Execute strSql, intI
'
'                     End If
'                     .MoveNext
'                  Loop
'               End With
'            End If
'         End If
'         'end 2010/4/20
'      'End If
'
'      If txt1(12) = "1" Then
'         If m_bPrinter = True Then
'            m_Device.EndDoc
'         End If
'
'      '列印統計表
'      'Modify by Morgan 2010/10/19 個人也可印統計
'      'ElseIf Txt1(6) = "" Then
'      Else
'         If m_bPrinter = True Then
'            m_Device.Orientation = 2
'         End If
'         Page = 0
'         m_GrpType = 1
'         PrintStatistic
'         If m_bPrinter = True Then
'            m_Device.EndDoc
'         End If
'         If txt1(11) = "" Then
'            If m_bPrinter = True Then
'               m_Device.Orientation = 2
'            End If
'            Page = 0
'            m_GrpType = 2
'            PrintStatistic
'            If m_bPrinter = True Then
'               m_Device.EndDoc
'            End If
'            'Modify by Morgan 2007/11/26 加程序人員選項,全部改為3
'            'If txt1(13) = "2" Then
'            If txt1(13) = "3" Then
'               If m_bPrinter = True Then
'                  m_Device.Orientation = 2
'               End If
'               Page = 0
'               m_GrpType = 3
'               PrintStatistic
'               If m_bPrinter = True Then
'                  m_Device.EndDoc
'               End If
'            End If
'         End If
'      End If
'      End With
'      If m_bPrinter = True Then
'         ShowPrintOk
'      ElseIf m_iPages > 0 Then
'         SetPic m_iPages
'         frm060312_1.m_ImageW = m_Device.Width
'         frm060312_1.m_ImageH = m_Device.Height
'         frm060312_1.m_iPages = m_iPages
'         frm060312_1.Show
'      End If
'   End If
'   Set rsQuery = Nothing
'   Set rsQuery1 = Nothing
'   Set m_Device = Nothing
'End Sub

'Modified by Lydia 2015/01/23 增加P案管制人
Private Sub ProcessNew2()
   Dim stVTB As String
   Dim stConPA As String, stConSP As String, stConLC As String
   Dim stCon, stCon1K0 As String, stCon0K0 As String, stConCP As String, stCon1N0 As String
   Dim stCon0 As String, stCon1 As String, stCon2 As String, stCon3 As String
   Dim stSys As String, ii As Integer
   Dim rsQuery As ADODB.Recordset, rsQuery1 As ADODB.Recordset
   Dim intR As Integer, strKeyNow As String, strKeyLast As String
   Dim strGrpNo As String 'Added by Morgan 2013/1/10 組別代碼
   Dim spartPA As String, spartSP As String, spartLC  As String 'Add by Lydia 2015/01/23
   Dim repMidStr As String  'Add by Lydia 2015/01/23 管制人取代承辦人
   'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
   Dim SeColPA As String
   Dim SeColSP As String
   Dim SeColLC As String
   Dim strExcept As String
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/13 清除查詢印表記錄檔欄位
   
   stCon0 = "": stCon1 = "": stCon2 = "": stCon3 = ""
   stCon1K0 = "": stCon0K0 = "": stConCP = ""
   stCon1N0 = ""
   
   'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
   SeColPA = " ,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,PA26 AS CUST01,PA27 AS CUST02,PA28 AS CUST03,PA29 AS CUST04,PA30 AS CUST05,PA75 AS FCNO "
   SeColSP = " ,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,SP08 AS CUST01,SP58 AS CUST02,SP59 AS CUST03,SP65 AS CUST04,SP66 AS CUST05,SP26 AS FCNO "
   SeColLC = " ,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,LC11 AS CUST01,LC43 AS CUST02,LC44 AS CUST03,LC45 AS CUST04,LC46 AS CUST05,LC22 AS FCNO "
   m_AllSys = IIf(frm060312.txt1(0).Text <> "ALL", frm060312.txt1(0).Text, GetAllSysKind(, frm060312.txt1(0).Text))
   intCufaCnt = 0
   'end 2019/11/01
        
   stSys = "'" & Join(Split(txt1(0), ","), "','") & "'"
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/13
   If txt1(1) = "1" Then
      stCon1K0 = stCon1K0 & " and a1k13||'' in (" & stSys & ")"
      pub_QL05 = pub_QL05 & ";" & Left(Label1(1), 4) & "1.請款" 'Add By Sindy 2010/12/13
   Else
      pub_QL05 = pub_QL05 & ";" & Left(Label1(1), 4) & "2.發文" 'Add By Sindy 2010/12/13
   End If
   stCon = stCon & " and c1.cp01 in (" & stSys & ")"
   
   '日期
   If txt1(2) <> "" Then
      '請款
      If txt1(1) = "1" Then
         stCon1K0 = stCon1K0 & " and a1k02>=" & txt1(2)
         stCon0K0 = stCon0K0 & " and a0k02>=" & txt1(2)
         stConCP = stConCP & " and c1.cp27>=" & DBDATE(txt1(2))
      '發文
      Else
         stCon = stCon & " and c1.cp27>=" & DBDATE(txt1(2))
      End If
   End If
   If txt1(3) <> "" Then
      '請款
      If txt1(1) = "1" Then
         stCon1K0 = stCon1K0 & " and a1k02<=" & txt1(3)
         stCon0K0 = stCon0K0 & " and a0k02<=" & txt1(3)
         stConCP = stConCP & " and c1.cp27<=" & DBDATE(txt1(3))
      '發文
      Else
         stCon = stCon & " and c1.cp27<=" & DBDATE(txt1(3))
      End If
   End If
   If txt1(2) <> "" Or txt1(3) <> "" Then
      If txt1(1) = "1" Then
         pub_QL05 = pub_QL05 & ";請款" & Label1(2) & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/12/13
      Else
         pub_QL05 = pub_QL05 & ";發文" & Label1(2) & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/12/13
      End If
   End If
   
   '案件性質
   If txt1(4) <> "" Then
      stCon = stCon & " and c1.cp10>='" & txt1(4) & "'"
   End If
   If txt1(5) <> "" Then
      stCon = stCon & " and c1.cp10<='" & txt1(5) & "'"
   End If
   If txt1(4) <> "" Or txt1(5) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/12/13
   End If
   
   'Add by Lydia 2015/01/23
    If txt1(13) = "4" Then
       stCon = stCon & " and c1.cp01='P' and substr(c1.cp12,1,1)='F'"
    End If
      
   '承辦人
   If txt1(6) <> "" Then
JumpToCP14: 'Added by Lydia 2025/02/06
      stCon1N0 = stCon1N0 & " and a1n04='" & txt1(6) & "'"
      pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(6) & lbl1 'Add By Sindy 2010/12/13
      
      '一般案件性質
      stCon0 = stCon0 & " and c1.cp14='" & txt1(6) & "'"
      '翻譯
      strExc(0) = "select sim01,sim02 from staff_idmap where '" & txt1(6) & "' in (sim01,sim02)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         stCon1 = stCon1 & " and c1.cp14 in ('" & RsTemp.Fields(0) & "','" & RsTemp.Fields(1) & "')"
      Else
         stCon1 = stCon1 & " and c1.cp14='" & txt1(6) & "'"
      End If
      '核稿
      stCon2 = stCon2 & " and ep04='" & txt1(6) & "'"
   'Added by Lydia 2025/02/06
   Else
      '外專/日專工程師中級主管(主任)可查詢底下所有人員的資料: 未輸入承辦人/核稿人
      If strST05 = "39" Then
         strExc(0) = ChkST52Range
         If strExc(0) = "" Then
            txt1(6) = strUserNum
            GoTo JumpToCP14
         Else
            stCon1N0 = stCon1N0 & " and a1n04 in (" & GetAddStr(strExc(0)) & ") "
            pub_QL05 = pub_QL05 & ";" & Label1(4) & PUB_ReadUserData(Replace(strExc(0), ",", ";"), True)
            
            '一般案件性質
            stCon0 = stCon0 & " and c1.cp14 in (" & GetAddStr(strExc(0)) & ") "
            '翻譯
            strExc(1) = "select sim01,sim02 from staff_idmap where sim01 in (" & GetAddStr(strExc(0)) & ")"
            intI = 1
            strExc(2) = strExc(0)
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
            If intI = 1 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                  If "" & RsTemp.Fields("sim02") <> "" Then
                     strExc(2) = strExc(2) & "," & RsTemp.Fields("sim02")
                  End If
                  RsTemp.MoveNext
               Loop
            End If
            stCon1 = stCon1 & " and c1.cp14 in (" & GetAddStr(strExc(2)) & ") "
            '核稿
            stCon2 = stCon2 & " and ep04 in (" & GetAddStr(strExc(2)) & ") "
         End If
      End If
   'end 2025/02/06
   End If

   '申請人
   If Len(Trim(txt1(7))) <> 0 And Len(Trim(txt1(8))) <> 0 Then
      stConPA = stConPA & " AND ((PA26>='" & GetNewFagent(txt1(7)) & "' AND PA26<='" & GetNewFagent(txt1(8)) & "') OR (PA27>='" & GetNewFagent(txt1(7)) & "' AND PA27<='" & GetNewFagent(txt1(8)) & "') OR (PA28>='" & GetNewFagent(txt1(7)) & "' AND PA28<='" & GetNewFagent(txt1(8)) & "') OR (PA29>='" & GetNewFagent(txt1(7)) & "' AND PA29<='" & GetNewFagent(txt1(8)) & "') OR (PA30>='" & GetNewFagent(txt1(7)) & "' AND PA30<='" & GetNewFagent(txt1(8)) & "')) "
      'Modified by Lydia 2019/11/01 補上申請人1~5
      'stConSP = stConSP & " AND ((SP08>='" & GetNewFagent(txt1(7)) & "' AND SP08<='" & GetNewFagent(txt1(8)) & "') OR (SP58<='" & GetNewFagent(txt1(7)) & "' AND SP58<='" & GetNewFagent(txt1(8)) & "') OR (SP59>='" & GetNewFagent(txt1(7)) & "' AND SP59<='" & GetNewFagent(txt1(8)) & "')) "
      'stConLC = stConLC & " AND ((LC11>='" & GetNewFagent(txt1(7)) & "' AND LC11<='" & GetNewFagent(txt1(8)) & "')) "
      stConSP = stConSP & " AND ((SP08>='" & GetNewFagent(txt1(7)) & "' AND SP08<='" & GetNewFagent(txt1(8)) & "') OR (SP58<='" & GetNewFagent(txt1(7)) & "' AND SP58<='" & GetNewFagent(txt1(8)) & "') OR (SP59>='" & GetNewFagent(txt1(7)) & "' AND SP59<='" & GetNewFagent(txt1(8)) & "') OR (SP65>='" & GetNewFagent(txt1(7)) & "' AND SP65<='" & GetNewFagent(txt1(8)) & "') OR (SP66>='" & GetNewFagent(txt1(7)) & "' AND SP66<='" & GetNewFagent(txt1(8)) & "')) "
      stConLC = stConLC & " AND ((LC11>='" & GetNewFagent(txt1(7)) & "' AND LC11<='" & GetNewFagent(txt1(8)) & "') OR (LC43>='" & GetNewFagent(txt1(7)) & "' AND LC43<='" & GetNewFagent(txt1(8)) & "') OR (LC44>='" & GetNewFagent(txt1(7)) & "' AND LC44<='" & GetNewFagent(txt1(8)) & "') OR (LC45>='" & GetNewFagent(txt1(7)) & "' AND LC45<='" & GetNewFagent(txt1(8)) & "') OR (LC46>='" & GetNewFagent(txt1(7)) & "' AND LC46<='" & GetNewFagent(txt1(8)) & "')) "
   ElseIf Len(Trim(txt1(7))) <> 0 Then
      stConPA = stConPA & " AND (PA26>='" & GetNewFagent(txt1(7)) & "' OR PA27>='" & GetNewFagent(txt1(7)) & "' OR PA28>='" & GetNewFagent(txt1(7)) & "' OR PA29>='" & GetNewFagent(txt1(7)) & "' OR PA30>='" & GetNewFagent(txt1(7)) & "') "
      'Modified by Lydia 2019/11/01 補上申請人1~5
      'stConSP = stConSP & " AND NOT (SP08>='" & GetNewFagent(txt1(7)) & "' OR SP58>='" & GetNewFagent(txt1(7)) & "' OR SP59>='" & GetNewFagent(txt1(7)) & "') "
      'stConLC = stConLC & " AND ((LC11>='" & GetNewFagent(txt1(7)) & "')) "
      stConSP = stConSP & " AND (SP08>='" & GetNewFagent(txt1(7)) & "' OR SP58>='" & GetNewFagent(txt1(7)) & "' OR SP59>='" & GetNewFagent(txt1(7)) & "' OR SP65>='" & GetNewFagent(txt1(7)) & "' OR SP66>='" & GetNewFagent(txt1(7)) & "') "
      stConLC = stConLC & " AND (LC11>='" & GetNewFagent(txt1(7)) & "' OR LC43>='" & GetNewFagent(txt1(7)) & "' OR LC44>='" & GetNewFagent(txt1(7)) & "' OR LC45>='" & GetNewFagent(txt1(7)) & "' OR LC46>='" & GetNewFagent(txt1(7)) & "') "
   ElseIf Len(Trim(txt1(8))) <> 0 Then
      'Modified by Lydia 2019/11/01 補上申請人1~5
      'stConPA = stConPA & " AND NOT (PA26<='" & GetNewFagent(txt1(8)) & "' OR PA27<='" & GetNewFagent(txt1(8)) & "' OR PA28<='" & GetNewFagent(txt1(8)) & "' OR PA29<='" & GetNewFagent(txt1(8)) & "' OR PA30<='" & GetNewFagent(txt1(8)) & "') "
      'stConSP = stConSP & " AND NOT (SP08<='" & GetNewFagent(txt1(8)) & "' OR SP58<='" & GetNewFagent(txt1(8)) & "' OR SP59<='" & GetNewFagent(txt1(8)) & "') "
      'stConLC = stConLC & " AND ((LC11<='" & GetNewFagent(txt1(8)) & "')) "
      stConPA = stConPA & " AND (PA26<='" & GetNewFagent(txt1(8)) & "' OR PA27<='" & GetNewFagent(txt1(8)) & "' OR PA28<='" & GetNewFagent(txt1(8)) & "' OR PA29<='" & GetNewFagent(txt1(8)) & "' OR PA30<='" & GetNewFagent(txt1(8)) & "') "
      stConSP = stConSP & " AND (SP08<='" & GetNewFagent(txt1(8)) & "' OR SP58<='" & GetNewFagent(txt1(8)) & "' OR SP59<='" & GetNewFagent(txt1(8)) & "' OR SP65<='" & GetNewFagent(txt1(8)) & "' OR SP66<='" & GetNewFagent(txt1(8)) & "') "
      stConLC = stConLC & " AND (LC11<='" & GetNewFagent(txt1(8)) & "' OR LC43<='" & GetNewFagent(txt1(8)) & "' OR LC44<='" & GetNewFagent(txt1(8)) & "' OR LC45<='" & GetNewFagent(txt1(8)) & "' OR LC46<='" & GetNewFagent(txt1(8)) & "') "
   End If
   If Len(Trim(txt1(7))) <> 0 Or Len(Trim(txt1(8))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(7) & "-" & txt1(8) 'Add By Sindy 2010/12/13
   End If
   
   '代理人
   If Len(Trim(txt1(9))) <> 0 And Len(Trim(txt1(10))) <> 0 Then
       stConPA = stConPA & " AND PA75>='" & GetNewFagent(txt1(9)) & "' AND PA75<='" & GetNewFagent(txt1(10)) & "' "
       stConSP = stConSP & " AND SP26>='" & GetNewFagent(txt1(9)) & "' AND SP26<='" & GetNewFagent(txt1(10)) & "' "
       stConLC = stConLC & " AND LC22>='" & GetNewFagent(txt1(9)) & "' AND LC22<='" & GetNewFagent(txt1(10)) & "' "
   ElseIf Len(Trim(txt1(9))) <> 0 Then
       stConPA = stConPA & " AND PA75>='" & GetNewFagent(txt1(9)) & "' "
       stConSP = stConSP & " AND SP26>='" & GetNewFagent(txt1(9)) & "' "
       stConLC = stConLC & " AND LC22>='" & GetNewFagent(txt1(9)) & "' "
   ElseIf Len(Trim(txt1(10))) <> 0 Then
       stConPA = stConPA & " AND PA75<='" & GetNewFagent(txt1(10)) & "' "
       stConSP = stConSP & " AND SP26<='" & GetNewFagent(txt1(10)) & "' "
       stConLC = stConLC & " AND LC22<='" & GetNewFagent(txt1(10)) & "' "
   End If
   If Len(Trim(txt1(9))) <> 0 Or Len(Trim(txt1(10))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(6) & txt1(9) & "-" & txt1(10) 'Add By Sindy 2010/12/13
   End If
   
   'Modified by Lydia 2015/01/23 +s1.
   '組別
   If txt1(11) <> "" Then
       stCon3 = stCon3 & " AND s1.ST16='" & txt1(11) & "'"
       pub_QL05 = pub_QL05 & ";" & Left(Label1(7), 3) & txt1(11) & "( 1.電子電機 2.化學 3.日文 4.機械設計 5.其他 )" 'Add By Sindy 2010/12/13
   End If
   '工程師
   If txt1(13) = "1" Then
      stCon3 = stCon3 & " AND s1.ST15='F21'"
      pub_QL05 = pub_QL05 & ";" & Left(Label1(12), 5) & "1.工程師" 'Add By Sindy 2010/12/13
   '程序
   ElseIf txt1(13) = "2" Then
      stCon3 = stCon3 & " AND s1.ST15='F22'"
      pub_QL05 = pub_QL05 & ";" & Left(Label1(12), 5) & "2.程序人員" 'Add By Sindy 2010/12/13
   '工程師+程序
   Else
      'Modified by Lydia 2015/01/23
      If txt1(13) = "4" Then
        '以fmp案條件為準
      Else
        stCon3 = stCon3 & " AND s1.ST15 IN ('F21','F22')"
      End If
      pub_QL05 = pub_QL05 & ";" & Left(Label1(12), 5) & "3.所有人" 'Add By Sindy 2010/12/13
   End If
   
   If txt1(12) = "1" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label1(8), 4) & "1.明細" 'Add By Sindy 2010/12/13
   Else
      pub_QL05 = pub_QL05 & ";" & Left(Label1(8), 4) & "2.統計" 'Add By Sindy 2010/12/13
   End If
   If txt1(14) = "1" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label1(9), 5) & "1.螢幕" 'Add By Sindy 2010/12/13
   Else
      pub_QL05 = pub_QL05 & ";" & Left(Label1(9), 5) & "2.印表機" 'Add By Sindy 2010/12/13
   End If
   
   'Add by Morgan 2010/4/20
   If stConPA <> "" Then
      stCon1N0 = stCon1N0 & " and (exists(select * from patent where pa01=c1.cp01 and pa02=c1.cp02 and pa03=c1.cp03 and pa04=c1.cp04 " & stConPA & ")" & _
         " or exists(select * from servicepractice where sp01=c1.cp01 and sp02=c1.cp02 and sp03=c1.cp03 and sp04=c1.cp04 " & stConSP & ")" & _
         " or exists(select * from lawcase where lc01=c1.cp01 and lc02=c1.cp02 and lc03=c1.cp03 and lc04=c1.cp04 " & stConLC & "))"
   End If
   'Modified by Lydia 2015/01/23 +s3
   If stCon3 <> "" Then
      stCon1N0 = stCon1N0 & " and exists(select * from staff s3 where s3.st01=a1n04 " & Replace(stCon3, "s1", "s3") & ")"
   End If
   'Add by Lydia 2015/01/23
   spartPA = " And substr(PA75,1,8)=FA01 And substr(PA75,9,1)=FA02 And FA10=NA01 and na16=s2.ST01(+) and d2.a0901(+)=s2.ST15"
   spartSP = " And substr(SP26,1,8)=FA01 And substr(SP26,9,1)=FA02 And FA10=NA01 and na16=s2.ST01(+) and d2.a0901(+)=s2.ST15"
   spartLC = " And substr(LC22,1,8)=FA01 And substr(LC22,9,1)=FA02 And FA10=NA01 and na16=s2.ST01(+) and d2.a0901(+)=s2.ST15"
    If txt1(13) = "4" Then
       'Added by Lydia 2017/02/13 FCP管制人又額外區分FMP案和寰華案管制人
       If strSrvDate(1) >= FMP管制人啟用日 Then
            spartPA = " And substr(PA75,1,8)=FA01 And substr(PA75,9,1)=FA02 And FA10=NA01 and nvl(na79,na16)=s2.ST01(+) and d2.a0901(+)=s2.ST15"
            spartSP = " And substr(SP26,1,8)=FA01 And substr(SP26,9,1)=FA02 And FA10=NA01 and nvl(na79,na16)=s2.ST01(+) and d2.a0901(+)=s2.ST15"
       End If
       'end 2017/02/13
       spartPA = " And PA09<>'000'" & spartPA
       spartSP = " And SP09<>'000'" & spartSP
      repMidStr = ",NVL(s2.st01,'000') C11,NVL(d2.A0902,'未分配') C12,decode(s2.st01,null,'N','Y') C13,NVL(s2.st02,'未分配') C14"
    'Add By Sindy 2015/9/22
    ElseIf txt1(13) = "5" Then '發文操作人
      stCon3 = stCon3 & " AND s4.ST15='F22'"
      repMidStr = ",s4.st01 C11,d4.A0902 C12,s4.ST16 C13,s4.st02 C14"
    '2015/9/22 END
    Else
      repMidStr = ",s1.st01 C11,d1.A0902 C12,s1.ST16 C13,s1.st02 C14"
    End If
      
   'Add by Morgan 2010/4/19 分配點數語法
   '請款
   If txt1(1) = "1" Then
      '先請款後發文
      m_strSharePointVTB = "select '" & strUserNum & "' id,a1n04,a1n03,a1n05,cp01,cp02,cp03,cp04,cp10,a1k02,'*' Flg,a1n06" & _
         " From caseprogress c1,acc1k0,acc1n0,engineerprogress" & _
         " Where cp60>'X' and cp27>0 and cp14 is not null" & stConCP & stCon & _
         " and a1k01(+)=cp60 and a1k02+19110000<cp27 and a1k12 is null and a1k25 is null" & _
         " and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n05>0 and a1n04<>cp14" & _
         " and ep02(+)=cp09 and a1n04||a1n06<>ep04||'Y'" & stCon1N0
         
      '先發文後請款
      m_strSharePointVTB = m_strSharePointVTB & _
         " union select '" & strUserNum & "' id,a1n04,a1n03,a1n05,cp01,cp02,cp03,cp04,cp10,a1k02,'' Flg,a1n06" & _
         " From acc1k0,caseprogress c1,acc1n0,engineerprogress" & _
         " Where a1k12 Is Null and a1k25 is null" & stCon1K0 & _
         " and cp60(+)=a1k01 and cp27<=a1k02+19110000 and cp14 is not null" & stCon & _
         " and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n05>0 and a1n04<>cp14" & _
         " and ep02(+)=cp09 and a1n04||a1n06<>ep04||'Y'" & stCon1N0
   '發文
   Else
      m_strSharePointVTB = "select '" & strUserNum & "' id,a1n04,a1n03,a1n05,cp01,cp02,cp03,cp04,cp10,cp27,'' Flg,a1n06" & _
         " From caseprogress c1,acc1n0,engineerprogress,acc1k0" & _
         " Where cp60>'X' and cp27>0 and cp14 is not null" & stConCP & stCon & _
         " and a1k01(+)=cp60 and a1k25 is null and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n05>0" & _
         " and ep02(+)=cp09 and a1n04<>cp14 and a1n04||a1n06<>ep04||'Y'" & stCon1N0
   End If
   
   'Modifyby Morgan 2010/12/22 +重新核稿 229
   '新案翻譯除外
   '請款日(若發文前請款案件則以發文日為日期條件)
   If txt1(1) = "1" Then
      '請款單:先發文後請款案件
      stVTB = "select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp83,cp18,cp27,cp60,cp64,cp113" & _
         ",decode(a1k25,null,a1n05) C9" & _
         " from acc1k0,caseprogress c1,acc1n0 where a1k12 is null" & stCon1K0 & _
         " and cp60(+)=a1k01 and cp27<=a1k02+19110000 and cp14 is not null and cp10<>'201'" & stCon & stCon0 & _
         " and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n04(+)=cp14 and a1n06(+) is null"
      
      '請款單:先請款後發文案件
      stVTB = stVTB & " union all select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp83,cp18,cp27,cp60,cp64,cp113" & _
         ",decode(a1k25,null,a1n05) C9" & _
         " from caseprogress c1,acc1k0,acc1n0 where cp60>'X'" & stConCP & _
         " and cp27>0 and cp14 is not null" & _
         " and a1k01(+)=cp60 and a1k02+19110000<cp27 and a1k12 is null and cp10<>'201'" & stCon & stCon0 & _
         " and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n04(+)=cp14 and a1n06(+) is null"
        
      '承辦點數 = 業務收文點數
      '收據:先發文後請款案件
      stVTB = stVTB & " union all select a0k02 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp83,cp18,cp27,cp60,cp64,cp113" & _
         ",cp18 C9" & _
         " from acc0k0,caseprogress c1 where nvl(a0k09,0)=0" & stCon0K0 & _
         " and cp60(+)=a0k01 and cp27<=a0k02+19110000 and cp14 is not null and cp10<>'201'" & stCon & stCon0
      
      '收據:先請款後發文案件
      stVTB = stVTB & " union all select a0k02 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp83,cp18,cp27,cp60,cp64,cp113" & _
         ",cp18 C9" & _
         " from caseprogress c1,acc0k0 where cp60<'X'" & stConCP & _
         " and cp27>0 and cp14 is not null and cp10<>'201'" & _
         " and a0k01(+)=cp60 and a0k02+19110000<cp27 and nvl(a0k09,0)=0" & stCon & stCon0
         
         
      'Add by Morgan 2010/12/22
      '重新核稿
      '翻譯為請款單或未請款
      stVTB = stVTB & " union all select c1.cp27 as a1k02,c1.cp01,c1.cp02,c1.cp03,c1.cp04,c1.cp05,c1.cp06,c1.cp10,c1.cp12,c1.cp14,c1.cp83,c1.cp18,c1.cp27,c1.cp60,c1.cp64,c1.cp113" & _
         ",round((a1l05-nvl(a1l07,0))*0.3/1000,2) C9" & _
         " from caseprogress c1,caseprogress c2,acc1l0 where c1.cp10='229'" & stConCP & _
         " and c1.cp27>0 and c1.cp14 is not null" & stCon & stCon0 & _
         " and c2.cp01(+)=c1.cp01 and c2.cp02(+)=c1.cp02 and c2.cp03(+)=c1.cp03 and c2.cp04(+)=c1.cp04" & _
         " and c2.cp10(+)='201' and (c2.cp60 is null or c2.cp60>'X')" & _
         " and a1l01(+)=c2.cp60 and a1l04(+)=c2.cp10"
         
      '收據
      stVTB = stVTB & " union all select c1.cp27 as a1k02,c1.cp01,c1.cp02,c1.cp03,c1.cp04,c1.cp05,c1.cp06,c1.cp10,c1.cp12,c1.cp14,c1.cp83,c1.cp18,c1.cp27,c1.cp60,c1.cp64,c1.cp113" & _
         ",round(nvl(c2.cp18,0)*0.3,2) C9" & _
         " from caseprogress c1,caseprogress c2 where c1.cp10='229'" & stConCP & _
         " and c1.cp27>0 and c1.cp14 is not null" & stCon & stCon0 & _
         " and c2.cp01(+)=c1.cp01 and c2.cp02(+)=c1.cp02 and c2.cp03(+)=c1.cp03 and c2.cp04(+)=c1.cp04" & _
         " and c2.cp10(+)='201' and c2.cp60<'X'"
         
   '發文日
   Else
      'Modify by Morgan 2010/12/22 +排除重新核稿 229
      stVTB = "select cp27 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,cp14,cp83,cp18,cp05 as cp27,cp60,cp64,cp113" & _
         ",decode(substr(cp60,1,1),'E',cp18,'X',decode(a1k25,null,a1n05)) C9" & _
         " from caseprogress c1,acc1n0,acc1k0 where cp14 is not null and cp10<>'201' and cp10<>'229'" & stCon & stCon0 & _
         " and a1k01(+)=cp60 and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n04(+)=cp14 and a1n06(+) is null"
   
      'Add by Morgan 2010/12/22
      '重新核稿 229
      stVTB = stVTB & " union all select c1.cp27 as a1k02,c1.cp01,c1.cp02,c1.cp03,c1.cp04,c1.cp05,c1.cp06,c1.cp10,c1.cp12,c1.cp14,c1.cp83,c1.cp18,c1.cp05 as cp27,c1.cp60,c1.cp64,c1.cp113" & _
         ",round(decode(substr(c2.cp60,1,1),'E',c2.cp18,(a1l05-nvl(a1l07,0))/1000)*0.3,2) C9" & _
         " from caseprogress c1,caseprogress c2,acc1l0 where c1.cp14 is not null and c1.cp10='229'" & stCon & stCon0 & _
         " and c2.cp01(+)=c1.cp01 and c2.cp02(+)=c1.cp02 and c2.cp03(+)=c1.cp03 and c2.cp04(+)=c1.cp04" & _
         " and c2.cp10(+)='201' and a1l01(+)=c2.cp60 and a1l04(+)=c2.cp10"
         
   End If
   'Modify by Morgan 2010/8/16 百年蟲 sqldatet(a1k02)-->substrb(' '||sqldatet(a1k02),-9)
   'Modified by Lydia 2015/01/23 +PMA = P案管制人,+spartpa,spartsp,spartLC
                                 '承辦人=管制人 repMidStr
 '  strExc(1) = "select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
      ",substrb(pa05,1,16) C3,substrb(ptm03,1,6) C4,substrb(decode(pa09,'020',cpm04,cpm03),1,8) C5" & _
      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
      ",cp113 C8, C9,substrb(cp64,1,16) C10" & _
      ",s1.st01 C11,d1.A0902 C12,s1.ST16 C13,s1.st02 C14,NVL(NA16,'000') PMA01,NVL(s2.st02,'未分配') PMA02,NVL(d2.A0902,'未分配') PMA03 " & _
      " From (" & stVTB & ") X,patent, patenttrademarkmap, casepropertymap,staff s1,acc090 d1,fagent,staff s2,Nation, acc090 d2 " & _
      " where  pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA & _
      " and ptm01(+)=1 and ptm02(+)=pa08 and cpm01(+)=cp01 and cpm02(+)=cp10" & spartPA & _
      " and s1.st01(+)=cp14 and d1.a0901(+)=s1.ST15" & stCon3
   'Modify By Sindy 2015/9/22 + ,staff s4,acc090 d4 ; and s4.st01(+)=cp83 and d4.a0901(+)=s4.ST15
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   'Modified by Morgan 2018/7/3 因 O12 設為 UTF-8格式中文字不是固定為 2bytese 改用 substr 抓字數
   'Modified by Lydia 2019/11/01 +增加欄位: SeColPA
   strExc(1) = "select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
      ",substr(pa05,1,8) C3,substr(ptm03,1,3) C4,substr(decode(pa09,'000',cpm03,cpm04),1,4) C5" & _
      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
      ",cp113 C8, C9,substr(cp64,1,8) C10" & repMidStr & SeColPA & _
      " From (" & stVTB & ") X,patent, patenttrademarkmap, casepropertymap,staff s1,acc090 d1,fagent,staff s2,Nation, acc090 d2,staff s4,acc090 d4 " & _
      " where  pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA & _
      " and ptm01(+)=1 and ptm02(+)=pa08 and cpm01(+)=cp01 and cpm02(+)=cp10" & spartPA & _
      " and s1.st01(+)=cp14 and d1.a0901(+)=s1.ST15 and s4.st01(+)=cp83 and d4.a0901(+)=s4.ST15" & stCon3
   'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷，列出排除案件
   strExcept = ProcExceptList(strExc(1) & " order by CASENO ")
   If strExcept <> "" Then strExc(1) = strExc(1) & strExcept
   'end 2019/11/21
   
   'Modify By Sindy 2015/9/22 + ,staff s4,acc090 d4 ; and s4.st01(+)=cp83 and d4.a0901(+)=s4.ST15
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   'Modified by Morgan 2018/7/3 因 O12 設為 UTF-8格式中文字不是固定為 2bytese 改用 substr 抓字數
   'Modified by Lydia 2019/11/01 +增加欄位: SeColSP
   strExc(1) = strExc(1) & " union all select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
      ",substr(sp05,1,8) C3,'' C4,substr(decode(sp09,'000',cpm03,cpm04),1,4) C5" & _
      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
      ",cp113 C8, C9,substr(cp64,1,8) C10" & repMidStr & SeColSP & _
      " From (" & stVTB & ") X,servicepractice,casepropertymap,staff s1,acc090 d1,fagent,staff s2,Nation, acc090 d2,staff s4,acc090 d4 " & _
      " where  sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10" & spartSP & _
      " and s1.st01(+)=cp14 and d1.a0901(+)=s1.ST15 and s4.st01(+)=cp83 and d4.a0901(+)=s4.ST15" & stCon3
   'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷，列出排除案件
   strExcept = ProcExceptList(strExc(1) & " order by CASENO ")
   If strExcept <> "" Then strExc(1) = strExc(1) & strExcept
   'end 2019/11/21
   
   'Modify By Sindy 2015/9/22 + ,staff s4,acc090 d4 ; and s4.st01(+)=cp83 and d4.a0901(+)=s4.ST15
   'Modified by Morgan 2018/7/3 因 O12 設為 UTF-8格式中文字不是固定為 2bytese 改用 substr 抓字數
   'Modified by Lydia 2019/11/01 +增加欄位: SeColLC
   strExc(1) = strExc(1) & " union all select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
      ",substr(LC05,1,8) C3,'' C4,substr(cpm03,1,4) C5" & _
      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
      ",cp113 C8, C9,substr(cp64,1,8) C10" & repMidStr & SeColLC & _
      " From (" & stVTB & ") X,lawcase,casepropertymap,staff s1,acc090 d1,fagent,staff s2,Nation, acc090 d2,staff s4,acc090 d4 " & _
      " where  lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04 and lc01 is not null" & stConLC & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10" & spartLC & _
      " and s1.st01(+)=cp14 and d1.a0901(+)=s1.ST15 and s4.st01(+)=cp83 and d4.a0901(+)=s4.ST15" & stCon3
   'end  'Modified by Lydia 2015/01/23 +PMA = P案管制人,+spartpa,spartsp,spartLC
   '新案翻譯(內翻的外譯編號要轉成所內員工編號),原則上所內編號(上班翻)的才會有承辦點數
   'Modify by Morgan 2011/7/5 排除承辦人為外翻部門的(否則已離職員工也會印出)
   '請款日
   If txt1(1) = "1" Then
      '請款單:先發文後請款案件
      stVTB = "select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,nvl(sim01,cp14) as cp14,cp83,cp18,cp27,cp60,cp64,cp113" & _
         ",decode(a1k25,null,a1n05) C9" & _
         " from acc1k0,caseprogress c1,staff bs1,acc1n0,staff_idmap where a1k12 is null" & stCon1K0 & _
         " and cp60(+)=a1k01 and cp27<=a1k02+19110000 and cp14 is not null and cp10='201'" & stCon & stCon1 & _
         " and bs1.st01(+)=cp14 and bs1.st03<>'F51' and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n04(+)=cp14 and a1n06(+) is null" & _
         " and sim02(+)=cp14"
      
      '請款單:先請款後發文案件
      stVTB = stVTB & " union all select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,nvl(sim01,cp14) as cp14,cp83,cp18,cp27,cp60,cp64,cp113" & _
         ",decode(a1k25,null,a1n05) C9" & _
         " from caseprogress c1,staff bs2,acc1k0,acc1n0,staff_idmap where cp60>'X'" & stConCP & _
         " and cp27>0 and cp14 is not null and bs2.st01(+)=cp14 and bs2.st03<>'F51'" & _
         " and a1k01(+)=cp60 and a1k02+19110000<cp27 and a1k12 is null and cp10='201'" & stCon & stCon1 & _
         " and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n04(+)=cp14 and a1n06(+) is null" & _
         " and sim02(+)=cp14"
   
      '內翻翻譯承辦點數=智權人員收文點數-核稿點數(核稿人不是承辦人才會有),否則=0
      '收據:先發文後請款案件
      stVTB = stVTB & " union all select a0k02 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,nvl(sim01,cp14) as cp14" & _
         ",cp83,cp18,cp27,cp60,cp64,cp113,decode(substrb(cp14,1,1),'F',0,cp18*decode(cp14,ep04,1,0.7*nvl(TF06,100)/100)) C9" & _
         " from acc0k0,caseprogress c1,staff bs3,staff_idmap,engineerprogress,transfee where nvl(a0k09,0)=0" & stCon0K0 & _
         " and cp60(+)=a0k01 and cp27<=a0k02+19110000 and cp14 is not null and cp10='201'" & stCon & stCon1 & _
         " and bs3.st01(+)=cp14 and bs3.st03<>'F51' and sim02(+)=cp14 and ep02(+)=cp09 and tf01(+)=cp09"
         
      '收據:先請款後發文案件
      stVTB = stVTB & " union all select a0k02 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,nvl(sim01,cp14) as cp14" & _
         ",cp83,cp18,cp27,cp60,cp64,cp113,decode(substrb(cp14,1,1),'F',0,cp18*decode(cp14,ep04,1,0.7*nvl(TF06,100)/100)) C9" & _
         " from caseprogress c1,staff bs4,acc0k0,staff_idmap,engineerprogress,transfee where cp60<'X'" & stConCP & _
         " and cp27>0 and cp14 is not null and cp10='201' and bs4.st01(+)=cp14 and bs4.st03<>'F51'" & _
         " and a0k01(+)=cp60 and a0k02+19110000<cp27 and nvl(a0k09,0)=0" & stCon & stCon1 & _
         " and sim02(+)=cp14 and ep02(+)=cp09 and tf01(+)=cp09"
         
   '發文日
   Else
      '請款單或未請款
      stVTB = "select cp27 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,nvl(sim01,cp14) as cp14,cp83,cp18,cp05 as cp27,cp60,cp64,cp113" & _
         ",decode(a1k25,null,a1n05) C9" & _
         " from caseprogress c1,staff bs5,acc1k0,acc1n0,staff_idmap where (cp60 is null or cp60>'X') and cp14 is not null and cp10='201'" & stCon & stCon1 & _
         " and bs5.st01(+)=cp14 and bs5.st03<>'F51' and a1k01(+)=cp60 and a1n01(+)=cp60 and a1n02(+)='2' and a1n03(+)=cp09 and a1n04(+)=cp14 and a1n06(+) is null" & _
         " and sim02(+)=cp14"
      
      '收據
      stVTB = stVTB & " union all select cp27 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp10,cp12,nvl(sim01,cp14) as cp14,cp83,cp18,cp05 as cp27,cp60,cp64,cp113" & _
         ",decode(substrb(cp14,1,1),'F',0,cp18*decode(cp14,ep04,1,0.7*nvl(TF06,100)/100)) C9" & _
         " from caseprogress c1,staff bs6,staff_idmap,engineerprogress,transfee where cp60<'X' and cp14 is not null and cp10='201'" & stCon & stCon1 & _
         " and bs6.st01(+)=cp14 and bs6.st03<>'F51' and sim02(+)=cp14 and ep02(+)=cp09 and tf01(+)=cp09"
   End If
   'Modified by Lydia 2015/01/23 +PMA = P案管制人,+spartpa,spartsp
   'Modify By Sindy 2015/9/22 + ,staff s4,acc090 d4 ; and s4.st01(+)=cp83 and d4.a0901(+)=s4.ST15
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   'Modified by Morgan 2018/7/3 因 O12 設為 UTF-8格式中文字不是固定為 2bytese 改用 substr 抓字數
   'Modified by Lydia 2019/11/01 +增加欄位: SeColPA
   strExc(2) = " union all select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
      ",substr(pa05,1,8) C3,substr(ptm03,1,3) C4,substr(decode(pa09,'000',cpm03,cpm04),1,4) C5" & _
      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
      ",cp113 C8,C9,substr(cp64,1,8) C10" & repMidStr & SeColPA & _
      " From (" & stVTB & ") X,patent, patenttrademarkmap, casepropertymap,staff s1,acc090 d1,fagent,staff s2,Nation, acc090 d2,staff s4,acc090 d4 " & _
      " where pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA & _
      " and ptm01(+)=1 and ptm02(+)=pa08 and cpm01(+)=cp01 and cpm02(+)=cp10" & spartPA & _
      " and s1.st01(+)=cp14 and d1.a0901(+)=s1.ST15 and s4.st01(+)=cp83 and d4.a0901(+)=s4.ST15" & stCon3
   'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷，列出排除案件
   strExcept = ProcExceptList(strExc(2) & " order by CASENO ")
   If strExcept <> "" Then strExc(2) = strExc(2) & strExcept
   'end 2019/11/21
         
   'Modify By Sindy 2015/9/22 + ,staff s4,acc090 d4 ; and s4.st01(+)=cp83 and d4.a0901(+)=s4.ST15
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   'Modified by Morgan 2018/7/3 因 O12 設為 UTF-8格式中文字不是固定為 2bytese 改用 substr 抓字數
   'Modified by Lydia 2019/11/01 +增加欄位: SeColSP
   strExc(2) = strExc(2) & " union all select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
      ",substr(sp05,1,8) C3,'' C4,substr(decode(sp09,'000',cpm03,cpm04),1,4) C5" & _
      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
      ",cp113 C8,C9,substr(cp64,1,8) C10" & repMidStr & SeColSP & _
      " From (" & stVTB & ") X,servicepractice,casepropertymap,staff s1,acc090 d1,fagent,staff s2,Nation, acc090 d2,staff s4,acc090 d4 " & _
      " where  sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10" & spartSP & _
      " and s1.st01(+)=cp14 and d1.a0901(+)=s1.ST15 and s4.st01(+)=cp83 and d4.a0901(+)=s4.ST15" & stCon3
   'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷，列出排除案件
   strExcept = ProcExceptList(strExc(2) & " order by CASENO ")
   If strExcept <> "" Then strExc(2) = strExc(2) & strExcept
   'end 2019/11/21
   
    'end 'Modified by Lydia 2015/01/23 +PMA = P案管制人,+spartpa,spartsp
   '核稿
   '核稿點數=翻譯請款點數(扣除折扣)x(30%+因翻譯瑕疵扣減支付外譯人員翻譯費之百分比x70%)
   '請款日
   'Modified by Morgan 2013/5/28 +927其他翻譯 Ex.FG-000858
   If txt1(1) = "1" Then
      '請款單:
      stVTB = "select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,cp14,cp83,ep04,cp18,cp27,cp60,cp64,cp114" & _
         ",decode(a1k25,null,a1n05) C9" & _
         " from acc1k0,caseprogress c1,engineerprogress,acc1n0 where a1k12 is null" & stCon1K0 & _
         " and cp60(+)=a1k01 and cp27<=a1k02+19110000 and cp14 is not null and cp10 in ('201','927')" & stCon & stCon2 & _
         " and ep02(+)=cp09 and ep04<>cp14" & _
         " and a1n02(+)='2' and a1n03(+)=ep02 and a1n04(+)=ep04 and a1n06(+)='Y'"
      
      '請款單:先請款後發文案件
      stVTB = stVTB & " union all select a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,cp14,cp83,ep04,cp18,cp27,cp60,cp64,cp114" & _
         ",decode(a1k25,null,a1n05) C9" & _
         " from caseprogress c1,acc1k0,engineerprogress,acc1n0" & _
         " where cp60>'X'" & stConCP & _
         " and cp27>0 and cp14 is not null and cp10 in ('201','927')" & _
         " and a1k01(+)=cp60 and a1k02+19110000<cp27 and a1k12 is null" & stCon & stCon2 & _
         " and ep02(+)=cp09 and ep04<>cp14" & _
         " and a1n02(+)='2' and a1n03(+)=ep02 and a1n04(+)=ep04 and a1n06(+)='Y'"
         
      '收據:核稿點數=智權人員收文點數x(30%+因翻譯瑕疵扣減支付外譯人員翻譯費之百分比x70%)
      stVTB = stVTB & " union all select a0k02 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,cp14,cp83,ep04,cp18,cp27,cp60,cp64,cp114" & _
         ",cp18*(0.3+0.7*(1-nvl(TF06,100)/100)) C9" & _
         " from acc0k0,caseprogress c1,engineerprogress,transfee where nvl(a0k09,0)=0" & stCon0K0 & _
         " and cp60(+)=a0k01 and cp27<=a0k02+19110000 and cp14 is not null and cp10='201'" & stCon & stCon2 & _
         " and ep02(+)=cp09 and ep04<>cp14 and tf01(+)=cp09"
      
      '收據:先請款後分案案件
      stVTB = stVTB & " union all select a0k02 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,cp14,cp83,ep04,cp18,cp27,cp60,cp64,cp114" & _
         ",cp18*(0.3+0.7*(1-nvl(TF06,100)/100)) C9" & _
         " from caseprogress c1,acc0k0,engineerprogress,transfee" & _
         " where cp60<'X'" & stConCP & _
         " and cp27>0 and cp14 is not null and cp10='201'" & _
         " and a0k01(+)=cp60 and a0k02+19110000<cp27 and nvl(a0k09,0)=0" & stCon & stCon2 & _
         " and ep02(+)=cp09 and ep04<>cp14 and tf01(+)=cp09"
         
   '發文日
   Else
      '請款單
      stVTB = "select cp27 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,cp14,cp83,ep04,cp18,cp05 as cp27,cp60,cp64,cp114" & _
         ",decode(a1k25,null,a1n05) C9" & _
         " from caseprogress c1,engineerprogress,acc1k0,acc1n0 where cp14 is not null and cp10 in ('201','927')" & stCon & stCon2 & _
         " and (cp60 is null or cp60>'X') and ep02(+)=cp09 and ep04<>cp14" & _
         " and a1k01(+)=cp60 and a1n02(+)='2' and a1n03(+)=ep02 and a1n04(+)=ep04 and a1n06(+)='Y'"
      '收據
      stVTB = stVTB & " union all select cp27 as a1k02,cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp12,cp14,cp83,ep04,cp18,cp05 as cp27,cp60,cp64,cp114" & _
         ",cp18*(0.3+0.7*(1-nvl(TF06,100)/100)) C9" & _
         " from caseprogress c1,engineerprogress,transfee where cp14 is not null and cp10='201'" & stCon & stCon2 & _
         " and cp60<'X' and ep02(+)=cp09 and ep04<>cp14" & _
         " and tf01(+)=cp09"
   End If
   'Modified by Lydia 2015/01/23 +PMA = P案管制人,+spartpa,spartsp
   'Modify By Sindy 2015/9/22 + ,staff s4,acc090 d4 ; and s4.st01(+)=cp83 and d4.a0901(+)=s4.ST15
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   'Modified by Morgan 2018/7/3 因 O12 設為 UTF-8格式中文字不是固定為 2bytese 改 substr 抓字數
   'Modified by Lydia 2019/11/01 +增加欄位: SeColPA
   strExc(3) = " union all select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
      ",substr(pa05,1,8) C3,substr(ptm03,1,3) C4,substr(decode(pa09,'000',cpm03,cpm04),1,4)||'-核稿' C5" & _
      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
      ",cp114 C8,C9,substr(cp64,1,8) C10" & repMidStr & SeColPA & _
      " From (" & stVTB & ") X,patent,patenttrademarkmap,casepropertymap,staff s1,acc090 d1,fagent,staff s2,Nation, acc090 d2,staff s4,acc090 d4 " & _
      " where pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA & _
      " and ptm01(+)=1 and ptm02(+)=pa08 and cpm01(+)=cp01 and cpm02(+)=cp10" & spartPA & _
      " and s1.st01(+)=ep04 and d1.a0901(+)=s1.ST15 and s4.st01(+)=cp83 and d4.a0901(+)=s4.ST15" & stCon3
   'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷，列出排除案件
   strExcept = ProcExceptList(strExc(3) & " order by CASENO ")
   If strExcept <> "" Then strExc(3) = strExc(3) & strExcept
   'end 2019/11/21
   
   'Modify By Sindy 2015/9/22 + ,staff s4,acc090 d4 ; and s4.st01(+)=cp83 and d4.a0901(+)=s4.ST15
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   'Modified by Lydia 2019/11/01 +增加欄位: SeColSP
   strExc(3) = strExc(3) & " union all select substrb(' '||sqldatet(a1k02),-9) C1,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||cp04) C2" & _
      ",substr(sp05,1,8) C3,'' C4,substr(decode(sp09,'000',cpm03,cpm04),1,4)||'-核稿' C5" & _
      ",substrb(sqldatet(cp06),1,9) C6,substrb(sqldatet(cp27),1,9) C7" & _
      ",cp114 C8,C9,substr(cp64,1,8) C10" & repMidStr & SeColSP & _
      " From (" & stVTB & ") X, servicepractice,casepropertymap,staff s1,acc090 d1,fagent,staff s2,Nation, acc090 d2,staff s4,acc090 d4 " & _
      " where  sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10" & spartSP & spartSP & _
      " and s1.st01(+)=ep04 and d1.a0901(+)=s1.ST15 and s4.st01(+)=cp83 and d4.a0901(+)=s4.ST15" & stCon3
   'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷，列出排除案件
   strExcept = ProcExceptList(strExc(3) & " order by CASENO ")
   If strExcept <> "" Then strExc(3) = strExc(3) & strExcept
   'end 2019/11/21
   
   strExc(0) = strExc(1) & strExc(2) & strExc(3) & " order by C12,C13,C11,C1,C2,C5"
   'end 'Modified by Lydia 2015/01/23 +PMA = P案管制人,+spartpa,spartsp ,承辦人=管制人 repMidStr
   
   stName = "": stGrp = "": stDep = "": stID = ""
   dblPoint = 0: dblTotPoint = 0
   dblPoint2 = 0: dblTotPoint2 = 0 'Add by Morgan 2010/11/15
   dblHour = 0: dblTotHour = 0
   m_RptType = 0: m_GrpType = 0
   Page = 0: m_iPages = 0
   iPrint = 0
   strKeyNow = "": strKeyLast = ""
   
   cnnConnection.Execute "DELETE FROM R060312_1 WHERE ID='" & strUserNum & "'"
    'Modified by Lydia 2015/01/23 P案管制人不計請款分配點數
    'Modify by Sindy 2015/9/23 發文操作人不計請款分配點數
    If txt1(13) = "4" Or txt1(13) = "5" Then
        intR = 0
    Else
        strSql = "INSERT INTO R060312_1(ID,R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11)" & m_strSharePointVTB
        cnnConnection.Execute strSql, intR
    End If
    
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      'Added by Lydia 2019/11/01
      If intCufaCnt > 0 Then
           MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
      End If
      'end 2019/11/01
      
      '跑明細報表且有分配點數資料
      If intR > 0 And txt1(12) = "1" Then
         strExc(0) = "select distinct st01 C11,A0902 C12,ST16 C13,st02 C14 from R060312_1,staff,acc090 where ID='" & strUserNum & "' and st01(+)=R01 and a0901(+)=ST15 order by C12,C13,C11"
         intI = 1
         Set rsQuery1 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If txt1(14) = "1" Then
               m_bPrinter = False
               Set m_Device = Picture1
               m_Device.AutoRedraw = True
               m_Device.Width = 16836
               m_Device.Height = 11904
               DelPic
            Else
               m_bPrinter = True
               Set m_Device = Printer
               m_Device.Orientation = 2
            End If
            
            GetPleft1
            With rsQuery1
            InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/13
            .MoveFirst
            Do While Not .EOF
               stID = "" & .Fields("C11")
               stName = "" & .Fields("C14")
               'Modified by Lydia 2015/01/23
               If .Fields("C13") = "Y" Then
                  stGrp = "P案管制人"
               ElseIf .Fields("C13") = "N" Then
                  stGrp = "未分配"
               Else
                  stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
               End If
               'end 2015/01/23
               stDep = "" & .Fields("C12")
               If stGrp = "" Then stGrp = stDep
               PrintTitle2
               iPrint = iPrint + 300
               PrintSharePoint stID
               .MoveNext
            Loop
            End With
            If m_bPrinter = True Then
               m_Device.EndDoc
               ShowPrintOk
            ElseIf m_iPages > 0 Then
               SetPic m_iPages
               frm060312_1.m_ImageW = m_Device.Width
               frm060312_1.m_ImageH = m_Device.Height
               frm060312_1.m_iPages = m_iPages
               frm060312_1.Show
            End If
         End If
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/12/13
         MsgBox "無可列印資料！"
      End If
   Else
      cnnConnection.Execute "DELETE FROM R060312 WHERE ID='" & strUserNum & "' "
      'Added by Lydia 2019/11/01
      If intCufaCnt > 0 Then
           MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
      End If
      'end 2019/11/01
      If txt1(14) = "1" Then
         m_bPrinter = False
         Set m_Device = Picture1
         m_Device.AutoRedraw = True
         m_Device.Width = 16836
         m_Device.Height = 11904
         DelPic
      Else
         m_bPrinter = True
         Set m_Device = Printer
         m_Device.Orientation = 2
      End If
      
      GetPleft1
      With rsQuery
      InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/13
      .MoveFirst
      Do While Not .EOF
         '列印明細小計
         If txt1(12) = "1" Then
            If stID <> "" & .Fields("C11") Then
               If stID <> "" Then
                  PrintSubTot stID
               End If
            End If
         End If
         'Modify by Morgan 2011/4/12 本來只考慮明細,改統計也要做
         If stID <> "" & .Fields("C11") Then
            'Add by Morgan 2010/4/20 檢查是否有無資料但有分配點數的情形
            If intR > 0 Then
               strKeyNow = "" & .Fields("C12") & .Fields("C13") & .Fields("C11")
               strExc(0) = "select distinct st01 C11,A0902 C12,ST16 C13,st02 C14 from R060312_1,staff,acc090 where ID='" & strUserNum & "' and st01(+)=R01 and a0901(+)=ST15 and A0902||ST16||st01>'" & strKeyLast & " ' and A0902||ST16||st01<'" & strKeyNow & "' order by C12,C13,C11"
               intI = 1
               Set rsQuery1 = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  With rsQuery1
                     Do While Not .EOF
                        If txt1(12) = "1" Then
                           stID = "" & .Fields("C11")
                           stName = "" & .Fields("C14")
                          ' stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
                            'Modified by Lydia 2015/01/23
                            If .Fields("C13") = "Y" Then
                               stGrp = "P案管制人"
                            ElseIf .Fields("C13") = "N" Then
                               stGrp = "未分配"
                            Else
                               stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
                            End If
                            'end 2015/01/23
                            
                           stDep = "" & .Fields("C12")
                           If stGrp = "" Then stGrp = stDep
                           PrintTitle2
                           iPrint = iPrint + 300
                           PrintSharePoint stID
                           
                        'Add by Morgan 2011/4/12
                        Else
                             stID = "" & .Fields("C11")
                             stName = "" & .Fields("C14")
                             strGrpNo = "" & .Fields("C13") 'Added by Morgan 2013/1/10
                            ' stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
                             'Modified by Lydia 2015/01/23
                             If .Fields("C13") = "Y" Then
                                stGrp = "P案管制人"
                             ElseIf .Fields("C13") = "N" Then
                                stGrp = "未分配"
                             Else
                                stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
                             End If
                              'end 2015/01/23
                             stDep = "" & .Fields("C12")
                           If stGrp = "" Then stGrp = stDep
                           strSql = " INSERT INTO R060312 (ID, R045001, R045003, R045005, R045008, R045004, R045006, R045007,R045014) values('" & strUserNum & "','" & stID & "','',0,0,'" & stDep & "','" & stGrp & "','" & stName & "','" & strGrpNo & "')"
                           cnnConnection.Execute strSql, intI
                           
                        End If
                        .MoveNext
                     Loop
                  End With
               End If
               strKeyLast = strKeyNow
            End If
            'end 2010/4/20
         End If
               
      
         If txt1(12) = "2" Then
            stID = "" & .Fields("C11")
            stName = "" & .Fields("C14")
            strGrpNo = "" & .Fields("C13") 'Added by Morgan 2013/1/10
           ' stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
            'Modified by Lydia 2015/01/23
            If .Fields("C13") = "Y" Then
               stGrp = "P案管制人"
            ElseIf .Fields("C13") = "N" Then
               stGrp = "未分配"
            Else
               stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
            End If
            'end 2015/01/23
            stDep = "" & .Fields("C12")
            'Add by Morgan 2007/8/23 若沒組別時設為部門
            If stGrp = "" Then stGrp = stDep
         Else
         
            If stID <> "" & .Fields("C11") Then
            
               dblPoint = 0: dblHour = 0
               dblPoint2 = 0 'Add by Morgan 2010/11/15
               stID = "" & .Fields("C11")
               stName = "" & .Fields("C14")
              ' stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
               'Modified by Lydia 2015/01/23
               If .Fields("C13") = "Y" Then
                  stGrp = "P案管制人"
               ElseIf .Fields("C13") = "N" Then
                  stGrp = "未分配"
               Else
                  stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
               End If
               'end 2015/01/23
                stDep = "" & .Fields("C12")
                'Add by Morgan 2007/8/23 若沒組別時設為部門
               If stGrp = "" Then stGrp = stDep
               PrintTitle2
               iPrint = iPrint + 300
               
            Else
               iPrint = iPrint + 400
               If iPrint > m_Device.ScaleHeight - 600 Then
                  PrintTitle2
                  iPrint = iPrint + 300
               End If
            End If
            
            For ii = 1 To 10
               strExc(0) = ""
               Select Case ii
                  Case 8
                     If "" & .Fields("C" & ii) <> "" Then
                        strExc(0) = Format("" & .Fields("C" & ii), "0.0")
                     End If
                     m_Device.CurrentX = PLeft(ii) + 960 - m_Device.TextWidth(strExc(0))
                     dblHour = dblHour + Val(strExc(0))
                     dblTotHour = dblTotHour + Val(strExc(0))
                  Case 9
                     If "" & .Fields("C" & ii) <> "" Then
                        strExc(0) = Format("" & .Fields("C" & ii), "0.00")
                     End If
                     m_Device.CurrentX = PLeft(ii) + 960 - m_Device.TextWidth(strExc(0))
                     dblPoint = dblPoint + Val(strExc(0))
                     dblTotPoint = dblTotPoint + Val(strExc(0))
                     'Add by Morgan 2010/11/15
                     If InStr("" & .Fields("C5"), "核稿") > 0 Then
                        dblPoint2 = dblPoint2 + Val(strExc(0))
                        dblTotPoint2 = dblTotPoint2 + Val(strExc(0))
                     End If
                  Case Else
                     strExc(0) = "" & .Fields("C" & ii)
                     'Add by Morgan 2008/1/3
                     If ii = 1 And Right("0" & .Fields("C1"), 9) < Right("0" & .Fields("C7"), 9) Then
                        strExc(0) = strExc(0) & "*"
                     End If
                     'end 2008/1/3
                     m_Device.CurrentX = PLeft(ii)
               End Select
               m_Device.CurrentY = iPrint
               'Modify by Morgan 2010/11/15
               'm_Device.Print strExc(0)
               If InStr(strExc(0), "核稿") > 0 Then
                  m_Device.FontBold = True
                  m_Device.Font.Underline = True
                  m_Device.Print strExc(0)
                  m_Device.Font.Underline = False
                  m_Device.FontBold = False
               Else
                  m_Device.Print strExc(0)
               End If
               'end 2010/11/15
            Next
         End If
         strSql = " INSERT INTO R060312 (ID, R045001, R045003, R045005, R045008, R045004, R045006, R045007, R045014) values('" & strUserNum & "','" & stID & "','" & .Fields("C5") & "'," & Format(Val("" & .Fields("C8")), "0.0") & "," & Format(Val("" & .Fields("C9")), "0.00") & ",'" & stDep & "','" & stGrp & "','" & stName & "','" & strGrpNo & "')"
         cnnConnection.Execute strSql, intI
         .MoveNext
      Loop
      
      If txt1(12) = "1" Then
         PrintSubTot stID
      End If
      
      'Modify by Morgan 2011/4/12 本來只考慮明細,改統計也要做
      'If txt1(12) = "1" Then
         'Add by Morgan 2010/4/20 檢查是否有無資料但有分配點數的情形
         If intR > 0 Then
            strExc(0) = "select distinct st01 C11,A0902 C12,ST16 C13,st02 C14 from R060312_1,staff,acc090 where ID='" & strUserNum & "' and st01(+)=R01 and a0901(+)=ST15 and A0902||ST16||st01>'" & strKeyLast & " ' order by C12,C13,C11"
            intI = 1
            Set rsQuery1 = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               With rsQuery1
                  Do While Not .EOF
                     If txt1(12) = "1" Then
                        stID = "" & .Fields("C11")
                        stName = "" & .Fields("C14")
                      '  stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
                        'Modified by Lydia 2015/01/23
                        If .Fields("C13") = "Y" Then
                           stGrp = "P案管制人"
                        ElseIf .Fields("C13") = "N" Then
                           stGrp = "未分配"
                        Else
                           stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
                        End If
                        'end 2015/01/23
                        stDep = "" & .Fields("C12")
                        If stGrp = "" Then stGrp = stDep
                        PrintTitle2
                        iPrint = iPrint + 300
                        PrintSharePoint stID
                     
                     'Add by Morgan 2011/4/12
                     Else
                        stID = "" & .Fields("C11")
                        stName = "" & .Fields("C14")
                        strGrpNo = "" & .Fields("C13") 'Added by Morgan 2013/1/10
                       ' stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
                        'Modified by Lydia 2015/01/23
                        If .Fields("C13") = "Y" Then
                           stGrp = "P案管制人"
                        ElseIf .Fields("C13") = "N" Then
                           stGrp = "未分配"
                        Else
                           stGrp = PUB_GetFCPGrpName("" & .Fields("C13"), True)
                        End If
                        'end 2015/01/23
                        stDep = "" & .Fields("C12")
                        If stGrp = "" Then stGrp = stDep
                        strSql = " INSERT INTO R060312 (ID, R045001, R045003, R045005, R045008, R045004, R045006, R045007,R045014) values('" & strUserNum & "','" & stID & "','',0,0,'" & stDep & "','" & stGrp & "','" & stName & "','" & strGrpNo & "')"
                        cnnConnection.Execute strSql, intI
                     
                     End If
                     .MoveNext
                  Loop
               End With
            End If
         End If
         'end 2010/4/20
      'End If
      
      If txt1(12) = "1" Then
         If m_bPrinter = True Then
            m_Device.EndDoc
         End If
      
      '列印統計表
      'Modify by Morgan 2010/10/19 個人也可印統計
      'ElseIf Txt1(6) = "" Then
      Else
         If m_bPrinter = True Then
            m_Device.Orientation = 2
         End If
         Page = 0
         m_GrpType = 1
         PrintStatistic
         If m_bPrinter = True Then
            m_Device.EndDoc
         End If
         If txt1(11) = "" Then
            If m_bPrinter = True Then
               m_Device.Orientation = 2
            End If
            Page = 0
            m_GrpType = 2
            PrintStatistic
            If m_bPrinter = True Then
               m_Device.EndDoc
            End If
            'Modified by Lydia 2015/01/23 + 4
            If txt1(13) = "3" Or txt1(14) = "4" Then
               If m_bPrinter = True Then
                  m_Device.Orientation = 2
               End If
               Page = 0
               m_GrpType = 3
               PrintStatistic
               If m_bPrinter = True Then
                  m_Device.EndDoc
               End If
            End If
         End If
      End If
      End With
      If m_bPrinter = True Then
         ShowPrintOk
      ElseIf m_iPages > 0 Then
         SetPic m_iPages
         frm060312_1.m_ImageW = m_Device.Width
         frm060312_1.m_ImageH = m_Device.Height
         frm060312_1.m_iPages = m_iPages
         frm060312_1.Show
      End If
   End If
   Set rsQuery = Nothing
   Set rsQuery1 = Nothing
   Set m_Device = Nothing
End Sub

'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷，列出排除案件
Private Function ProcExceptList(ByVal pSQL As String) As String
Dim intJ As Integer, strGrp As String, strTmp1 As String
Dim rsR1 As New ADODB.Recordset

    ProcExceptList = ""
    'Modified by Lydia 2024/12/31 只有明細排除限閱案件 ---- David
    If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" And txt1(12) = "1" Then
        intJ = 1
        If Left(Trim(UCase(pSQL)), 5) <> "UNION" Then
           Set rsR1 = ClsLawReadRstMsg(intJ, pSQL)
        Else
           Set rsR1 = ClsLawReadRstMsg(intJ, Mid(pSQL, InStr(UCase(pSQL), "SELECT")))
        End If
        If intJ = 1 Then
            With rsR1
                 .MoveFirst
                 Do While Not .EOF
                     If strGrp <> "" & .Fields("CASENO") Then
                        If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & .Fields("CASENO"), "" & .Fields("cust01") & "," & .Fields("cust02") & "," & .Fields("cust03") & "," & .Fields("cust04") & "," & .Fields("cust05"), "" & .Fields("fcno")) = False Then
                            intCufaCnt = intCufaCnt + 1
                            strTmp1 = strTmp1 & "," & .Fields("CASENO")
                        End If
                     End If
                     strGrp = "" & .Fields("CASENO")
                     .MoveNext
                 Loop
            End With
        End If
        Set rsR1 = Nothing
        
        If strTmp1 <> "" Then
            ProcExceptList = " AND CP01||'-'||CP02||'-'||CP03||'-'||CP04 NOT IN (" & GetAddStr(strTmp1) & ") "
        End If
    End If
End Function

'Added by Lydia 2025/02/06 外專/日專工程師中級主管(主任)可查詢底下所有人員的資料
Private Function ChkST52Range(Optional ByVal pUserNo As String) As String
Dim strQ1 As String, intQ As Integer
Dim rsQD As New ADODB.Recordset
   
   'Added by Lydia 2025/02/19 請開放Alina蘇韋寧(99025)可以查詢Ray蔡昀甫(B2042)的資料---Wilison
   'Mark by Lydia 2025/02/25 調整承辦人=空白，查詢同組別+ST52~ST55的下屬工程師
   'If strST16 <> "3" Then
   '   strQ1 = "select st01 from staff where instr(st52||','||st53||','||st54||','||st55,'" & strUserNum & "') > 0 and st03='F21' " & _
   '           IIf(pUserNo <> "", " and st01='" & Trim(pUserNo) & "' ", " and st04='1' ")
   'Else
   'end 2025/02/19
      strQ1 = "select st01 from staff where st16='" & strST16 & "' " & IIf(strST70 <> "", " and st70='" & strST70 & "' ", "") & _
              " and instr(st52||','||st53||','||st54||','||st55,'" & strUserNum & "') > 0 and st03='F21' " & _
              IIf(pUserNo <> "", " and st01='" & Trim(pUserNo) & "' ", " and st04='1' ")
   'End If 'Mark by Lydia 2025/02/25
   intQ = 1
   Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
   If intQ = 1 Then
      ChkST52Range = strUserNum & "," & rsQD.GetString(adClipString, , , ",")
      If Right(ChkST52Range, 1) = "," Then ChkST52Range = Mid(ChkST52Range, 1, Len(ChkST52Range) - 1)
   End If
   Set rsQD = Nothing
End Function
