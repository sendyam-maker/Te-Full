VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060309 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專收文未發文明細查詢/列印"
   ClientHeight    =   4600
   ClientLeft      =   1400
   ClientTop       =   2200
   ClientWidth     =   4980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4600
   ScaleWidth      =   4980
   Begin VB.CheckBox Check1 
      Caption         =   "含核對已准專利無承辦期限資料"
      Height          =   345
      Left            =   360
      TabIndex        =   40
      Top             =   4200
      Width           =   3000
   End
   Begin VB.TextBox txt2 
      Height          =   270
      Left            =   990
      MaxLength       =   1
      TabIndex        =   15
      Text            =   "1"
      Top             =   3840
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   18
      Left            =   1005
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1440
      Width           =   435
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   17
      Left            =   1005
      MaxLength       =   1
      TabIndex        =   14
      Top             =   3501
      Width           =   435
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1005
      TabIndex        =   0
      Top             =   468
      Width           =   3090
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1005
      MaxLength       =   3
      TabIndex        =   1
      Top             =   804
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2325
      MaxLength       =   3
      TabIndex        =   2
      Top             =   804
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1005
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1140
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1008
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1812
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1008
      MaxLength       =   4
      TabIndex        =   6
      Top             =   2148
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2328
      MaxLength       =   4
      TabIndex        =   7
      Top             =   2148
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1005
      MaxLength       =   4
      TabIndex        =   8
      Top             =   2484
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   2325
      MaxLength       =   4
      TabIndex        =   9
      Top             =   2484
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   9
      Left            =   5250
      MaxLength       =   4
      TabIndex        =   21
      Top             =   5655
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   10
      Left            =   6570
      MaxLength       =   4
      TabIndex        =   20
      Top             =   5655
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   11
      Left            =   5700
      MaxLength       =   1
      TabIndex        =   19
      Top             =   6135
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   12
      Left            =   5970
      MaxLength       =   1
      TabIndex        =   18
      Top             =   6615
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   13
      Left            =   1005
      MaxLength       =   9
      TabIndex        =   10
      Top             =   2820
      Width           =   1000
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   14
      Left            =   2325
      MaxLength       =   9
      TabIndex        =   11
      Top             =   2820
      Width           =   1000
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   15
      Left            =   1005
      MaxLength       =   9
      TabIndex        =   12
      Top             =   3165
      Width           =   1000
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   16
      Left            =   2328
      MaxLength       =   9
      TabIndex        =   13
      Top             =   3165
      Width           =   1000
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3180
      TabIndex        =   16
      Top             =   36
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3975
      TabIndex        =   17
      Top             =   36
      Width           =   800
   End
   Begin MSForms.Label LBL1 
      Height          =   300
      Index           =   1
      Left            =   1860
      TabIndex        =   42
      Top             =   1812
      Width           =   1470
      VariousPropertyBits=   27
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LBL1 
      Height          =   300
      Index           =   0
      Left            =   1860
      TabIndex        =   41
      Top             =   1140
      Width           =   1470
      VariousPropertyBits=   27
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(1.查詢  2.印表)"
      Height          =   180
      Left            =   1350
      TabIndex        =   39
      Top             =   3885
      Width           =   1200
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "列印別："
      Height          =   180
      Left            =   135
      TabIndex        =   38
      Top             =   3885
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "組別：                    ( 1電子電機 2化學 3日文 4機械設計)"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   37
      Top             =   1500
      Width           =   4260
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "(1.承辦人2.智權人員)"
      Height          =   180
      Index           =   2
      Left            =   1560
      TabIndex        =   36
      Top             =   3525
      Width           =   1650
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "列印順序："
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   35
      Top             =   3525
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   34
      Top             =   510
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "業務區："
      Height          =   180
      Left            =   135
      TabIndex        =   33
      Top             =   840
      Width           =   720
   End
   Begin VB.Line Line1 
      X1              =   1968
      X2              =   2208
      Y1              =   924
      Y2              =   924
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   135
      TabIndex        =   32
      Top             =   1170
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Left            =   135
      TabIndex        =   31
      Top             =   1845
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "收文天數："
      Height          =   180
      Left            =   135
      TabIndex        =   30
      Top             =   2220
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   1965
      X2              =   2205
      Y1              =   2295
      Y2              =   2295
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   29
      Top             =   2550
      Width           =   900
   End
   Begin VB.Line Line3 
      X1              =   1965
      X2              =   2205
      Y1              =   2625
      Y2              =   2625
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   4410
      TabIndex        =   28
      Top             =   5655
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Line Line4 
      Visible         =   0   'False
      X1              =   6210
      X2              =   6450
      Y1              =   5775
      Y2              =   5775
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "是否列印明細："
      Height          =   180
      Left            =   4410
      TabIndex        =   27
      Top             =   6135
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "(Y:印)"
      Height          =   180
      Left            =   5970
      TabIndex        =   26
      Top             =   6135
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "是否計算多國家："
      Height          =   180
      Left            =   4410
      TabIndex        =   25
      Top             =   6615
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "(Y:計算)"
      Height          =   180
      Left            =   6330
      TabIndex        =   24
      Top             =   6615
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Left            =   135
      TabIndex        =   23
      Top             =   2850
      Width           =   720
   End
   Begin VB.Line Line5 
      X1              =   2040
      X2              =   2280
      Y1              =   2955
      Y2              =   2955
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Left            =   135
      TabIndex        =   22
      Top             =   3195
      Width           =   720
   End
   Begin VB.Line Line6 
      X1              =   2040
      X2              =   2280
      Y1              =   3285
      Y2              =   3285
   End
End
Attribute VB_Name = "frm060309"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; LBL1(index); Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
'Modified by Morgan 2012/5/16 將查詢與列印功能合併(加列印別選項)
Option Explicit
Dim strSql As String, strTemp1 As Variant, strTemp2 As Variant, StrTest1 As String, StrTest2 As String, i As Integer, j As Integer, s As Integer
Dim PLeft(0 To 14) As Integer, k As Integer, TmpArea As String, iLine As Integer, Page As Integer
Dim strTemp3(0 To 14) As String, iPrint As Integer
Dim StrTest3 As String, Day1 As String, Day2 As String, StrTemp4 As String
Dim St As String, iK As Integer, iTatle As Integer
'Add By Cheng 2002/09/16
Dim blnClkSure As Boolean '判斷是否按下確定按鈕
Dim m_Grp As String '組別
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         'Add By Cheng 2002/09/16
         blnClkSure = False
           If Len(txt1(0)) = 0 Then
              s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
              txt1(0).SetFocus
              Exit Sub
           Else
               'Add By Cheng 2002/09/16
               If Me.txt1(1).Text <> "" And Me.txt1(2).Text <> "" Then
                  If Me.txt1(1).Text > Me.txt1(2).Text Then
                     MsgBox "業務區範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(1).SetFocus
                     txt1_GotFocus 1
                     Exit Sub
                  End If
               End If
               'Modify By Cheng 2002/09/27
      '         LBL1(0) = GetPrjSales(txt1(3))
               LBL1(0) = GetPrjSales(txt1(3), "智權人員")
               If Me.txt1(3).Text <> "" Then
                  If Me.txt1(3).Text = Me.LBL1(0).Caption Then
                     Me.LBL1(0).Caption = ""
                     Me.txt1(3).SetFocus
                     txt1_GotFocus 3
                     Exit Sub
                  End If
               End If
               LBL1(1) = GetPrjSales(txt1(4))
               If Me.txt1(4).Text <> "" Then
                  If Me.txt1(4).Text = Me.LBL1(1).Caption Then
                     Me.LBL1(1).Caption = ""
                     Me.txt1(4).SetFocus
                     txt1_GotFocus 4
                     Exit Sub
                  End If
               End If
              
              If Len(txt1(6)) = 0 Then
                  s = MsgBox("收文天數不可空白!!", , "USER 輸入錯誤")
                  txt1(5).SetFocus
                  txt1_GotFocus (5)
                  Exit Sub
              Else
                  'Add By Cheng 2002/09/16
                  If Me.txt1(5).Text <> "" And Me.txt1(6).Text <> "" Then
                     If Val(Me.txt1(5).Text) > Val(Me.txt1(6).Text) Then
                        MsgBox "收文天數範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                        blnClkSure = True
                        Me.txt1(5).SetFocus
                        txt1_GotFocus 5
                        Exit Sub
                     End If
                  End If
                  If Me.txt1(7).Text <> "" And Me.txt1(8).Text <> "" Then
                     If Me.txt1(7).Text > Me.txt1(8).Text Then
                        MsgBox "案件性質範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                        blnClkSure = True
                        Me.txt1(7).SetFocus
                        txt1_GotFocus 7
                        Exit Sub
                     End If
                  End If
                  If Me.txt1(13).Text <> "" And Me.txt1(14).Text <> "" Then
                     If Me.txt1(13).Text > Me.txt1(14).Text Then
                        MsgBox "申請人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                        blnClkSure = True
                        Me.txt1(13).SetFocus
                        txt1_GotFocus 13
                        Exit Sub
                     End If
                  End If
                  If Me.txt1(15).Text <> "" And Me.txt1(16).Text <> "" Then
                     If Me.txt1(15).Text > Me.txt1(16).Text Then
                        MsgBox "代理人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                        blnClkSure = True
                        Me.txt1(15).SetFocus
                        txt1_GotFocus 15
                        Exit Sub
                     End If
                  End If
                  
                  'Added by Morgan 2012/5/15
                  If txt2 = "" Then
                     MsgBox "列印別不可空白！"
                     txt2.SetFocus
                     Exit Sub
                  End If
                  'end 2012/5/15
         
                  Me.Enabled = False
                  Screen.MousePointer = vbHourglass
                  '若未輸入申請人及代理人條件
                  'Modified by Lydia 2019/11/01 +Trim
                  If Len(Trim(txt1(13))) = 0 And Len(Trim(txt1(14))) = 0 And Len(Trim(txt1(15))) = 0 And Len(Trim(txt1(16))) = 0 Then
                      StrMenu
                  Else
                      StrMenu2
                  End If
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
              End If
          End If
      Case 1
           Unload Me
      Case Else
   End Select
End Sub

'若未輸入申請人及代理人條件
Sub StrMenu()
Dim stConCP As String, stConCP1 As String, stConCP2 As String, stConCP3 As String
Dim stVTB As String
Dim iSort As Integer
Dim dblRow As Double 'Add By Sindy 2025/9/3

   Screen.MousePointer = vbHourglass
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/10 清除查詢印表記錄檔欄位
   StrTest1 = ""
   StrTest2 = ""
   StrTest3 = ""
   stConCP = ""
   stConCP1 = ""
   stVTB = ""
   'Added by Lydia 2019/11/01 利益衝突案件
   m_AllSys = IIf(txt1(0) <> "ALL", txt1(0), GetAllSysKind(, txt1(0)))
   intCufaCnt = 0
   'end 2019/11/01

   If Len(txt1(0)) <> 0 Then
       stConCP1 = stConCP1 & " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
       stConCP2 = stConCP2 & " and CP01 in (" & SQLGrpStr(txt1(0), 2) & ") "
       stConCP3 = stConCP3 & " and CP01 in (" & SQLGrpStr(txt1(0), 5) & ") "
       pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(0) 'Add By Sindy 2010/12/10
   End If
   If Len(txt1(1)) <> 0 Then
       stConCP = stConCP & " AND CP12>='" & txt1(1) & "' "
   End If
   If Len(txt1(2)) <> 0 Then
       stConCP = stConCP & " AND CP12<='" & txt1(2) & "' "
   End If
   If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/10
   End If
   If Len(txt1(3)) <> 0 Then
       stConCP = stConCP & " AND CP13='" & txt1(3) & "' "
       pub_QL05 = pub_QL05 & ";" & Label3 & txt1(3) & LBL1(0) 'Add By Sindy 2010/12/10
   End If
   If Len(txt1(4)) <> 0 Then
      'Modify by Morgan 2007/8/6 加外譯編號
      strExc(1) = PUB_GetMapID(txt1(4), 0)
      If strExc(1) <> "" Then
         stConCP = stConCP & " AND CP14 in ('" & txt1(4) & "','" & strExc(1) & "')"
      Else
         stConCP = stConCP & " AND CP14='" & txt1(4) & "' "
      End If
      pub_QL05 = pub_QL05 & ";" & Label5 & txt1(4) & LBL1(1) 'Add By Sindy 2010/12/10
   End If
   If Len(txt1(7)) <> 0 Then
       stConCP = stConCP & " AND CP10>='" & txt1(7) & "' "
   End If
   If Len(txt1(8)) <> 0 Then
       stConCP = stConCP & " AND CP10<='" & txt1(8) & "' "
   End If
   If Len(txt1(7)) <> 0 Or Len(txt1(8)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label8(0) & txt1(7) & "-" & txt1(8) 'Add By Sindy 2010/12/10
   End If
   '2008/2/22 ADD BY SONIA 加組別條件
   If Len(txt1(18)) <> 0 Then
       StrTest1 = StrTest1 & " AND S1.ST16='" & txt1(18) & "' "
       StrTest2 = StrTest2 & " AND S1.ST16='" & txt1(18) & "' "
       StrTest3 = StrTest3 & " AND S1.ST16='" & txt1(18) & "' "
       pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(18) & "( 1電子電機 2化學 3日文 4機械設計 5其他)" 'Add By Sindy 2010/12/10
   End If
   '2008/2/22 END
   If Len(txt1(9)) <> 0 Then
       StrTest1 = StrTest1 & " AND SUBSTR(PA09,1,3)>='" & txt1(9) & "' "
       StrTest2 = StrTest2 & " AND SUBSTR(TM10,1,3)>='" & txt1(9) & "' "
       StrTest3 = StrTest3 & " AND SUBSTR(SP09,1,3)>='" & txt1(9) & "' "
   End If
   If Len(txt1(10)) <> 0 Then
       StrTest1 = StrTest1 & " AND SUBSTR(PA09,1,3)<='" & txt1(10) & "' "
       StrTest2 = StrTest2 & " AND SUBSTR(TM10,1,3)<='" & txt1(10) & "' "
       StrTest3 = StrTest3 & " AND SUBSTR(SP09,1,3)<='" & txt1(10) & "' "
   End If
   If Len(txt1(9)) <> 0 Or Len(txt1(10)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label9 & txt1(9) & "-" & txt1(10) 'Add By Sindy 2010/12/10
   End If
   '申請人
   If Len(Trim(txt1(13))) <> 0 And Len(Trim(txt1(14))) <> 0 Then
       StrTest1 = StrTest1 & " AND ((PA26>='" & GetNewFagent(txt1(13)) & "' AND PA26<='" & GetNewFagent(txt1(14)) & "') OR (PA27>='" & GetNewFagent(txt1(13)) & "' AND PA27<='" & GetNewFagent(txt1(14)) & "') OR (PA28>='" & GetNewFagent(GetNewFagent(txt1(13))) & "' AND PA28<='" & GetNewFagent(txt1(14)) & "') OR (PA29>='" & GetNewFagent(GetNewFagent(txt1(13))) & "' AND PA29<='" & GetNewFagent(txt1(14)) & "') OR (PA30>='" & GetNewFagent(GetNewFagent(txt1(13))) & "' AND PA30<='" & GetNewFagent(txt1(14)) & "')) "
       'Modified by Lydia 2019/11/01 補上申請人1~5
       'StrTest2 = StrTest2 & " AND (TM23>='" & GetNewFagent(Txt1(13)) & "' AND TM23<='" & GetNewFagent(Txt1(14)) & "') "
       'StrTest3 = StrTest3 & " AND ((SP08>='" & GetNewFagent(Txt1(13)) & "' AND SP08<='" & GetNewFagent(Txt1(14)) & "') OR (SP58<='" & GetNewFagent(Txt1(13)) & "' AND SP58<='" & GetNewFagent(Txt1(14)) & "') OR (SP59>='" & GetNewFagent(Txt1(13)) & "' AND SP59<='" & GetNewFagent(Txt1(14)) & "')) "
       StrTest2 = StrTest2 & " AND ((TM23>='" & GetNewFagent(txt1(13)) & "' AND TM23<='" & GetNewFagent(txt1(14)) & "') OR (TM78>='" & GetNewFagent(txt1(13)) & "' AND TM78<='" & GetNewFagent(txt1(14)) & "') OR (TM79>='" & GetNewFagent(txt1(13)) & "' AND TM79<='" & GetNewFagent(txt1(14)) & "') OR (TM80>='" & GetNewFagent(txt1(13)) & "' AND TM80<='" & GetNewFagent(txt1(14)) & "') OR (TM81>='" & GetNewFagent(txt1(13)) & "' AND TM81<='" & GetNewFagent(txt1(14)) & "')) "
       StrTest3 = StrTest3 & " AND ((SP08>='" & GetNewFagent(txt1(13)) & "' AND SP08<='" & GetNewFagent(txt1(14)) & "') OR (SP58<='" & GetNewFagent(txt1(13)) & "' AND SP58<='" & GetNewFagent(txt1(14)) & "') OR (SP59>='" & GetNewFagent(txt1(13)) & "' AND SP59<='" & GetNewFagent(txt1(14)) & "') OR (SP65<='" & GetNewFagent(txt1(13)) & "' AND SP65<='" & GetNewFagent(txt1(14)) & "') OR (SP66>='" & GetNewFagent(txt1(13)) & "' AND SP66<='" & GetNewFagent(txt1(14)) & "')) "
   Else
       If Len(Trim(txt1(13))) <> 0 And Len(Trim(txt1(14))) = 0 Then
           StrTest1 = StrTest1 & " AND (PA26>='" & GetNewFagent(txt1(13)) & "' OR PA27>='" & GetNewFagent(txt1(13)) & "' OR PA28>='" & GetNewFagent(txt1(13)) & "' OR PA29>='" & GetNewFagent(txt1(13)) & "' OR PA30>='" & GetNewFagent(txt1(13)) & "') "
           'Modified by Lydia 2019/11/01 補上申請人1~5
           'StrTest2 = StrTest2 & " AND (TM23>='" & GetNewFagent(txt1(13)) & "' ) "
           'StrTest3 = StrTest3 & " AND (SP08>='" & GetNewFagent(txt1(13)) & "' OR SP58>='" & GetNewFagent(txt1(13)) & "' OR SP59>='" & GetNewFagent(txt1(13)) & "') "
           StrTest2 = StrTest2 & " AND (TM23>='" & GetNewFagent(txt1(13)) & "' OR TM78>='" & GetNewFagent(txt1(13)) & "' OR TM79>='" & GetNewFagent(txt1(13)) & "' OR TM80>='" & GetNewFagent(txt1(13)) & "' OR TM81>='" & GetNewFagent(txt1(13)) & "') "
           StrTest3 = StrTest3 & " AND (SP08>='" & GetNewFagent(txt1(13)) & "' OR SP58>='" & GetNewFagent(txt1(13)) & "' OR SP59>='" & GetNewFagent(txt1(13)) & "' OR SP65>='" & GetNewFagent(txt1(13)) & "' OR SP66>='" & GetNewFagent(txt1(13)) & "') "
       Else
           If Len(Trim(txt1(13))) = 0 And Len(Trim(txt1(14))) <> 0 Then
               StrTest1 = StrTest1 & " AND (PA26<='" & GetNewFagent(txt1(14)) & "' OR PA27<='" & GetNewFagent(txt1(14)) & "' OR PA28<='" & GetNewFagent(txt1(14)) & "' OR PA29<='" & GetNewFagent(txt1(14)) & "' OR PA30<='" & GetNewFagent(txt1(14)) & "') "
               'Modified by Lydia 2019/11/01 補上申請人1~5
               'StrTest2 = StrTest2 & " AND (TM23<='" & GetNewFagent(txt1(14)) & "') "
               'StrTest3 = StrTest3 & " AND (SP08<='" & GetNewFagent(txt1(14)) & "' OR SP58<='" & GetNewFagent(txt1(14)) & "' OR SP59<='" & GetNewFagent(txt1(14)) & "') "
                StrTest2 = StrTest2 & " AND (TM23<='" & GetNewFagent(txt1(14)) & "' OR TM78<='" & GetNewFagent(txt1(14)) & "' OR TM79<='" & GetNewFagent(txt1(14)) & "' OR TM80<='" & GetNewFagent(txt1(14)) & "' OR TM81<='" & GetNewFagent(txt1(14)) & "') "
                StrTest3 = StrTest3 & " AND (SP08<='" & GetNewFagent(txt1(14)) & "' OR SP58<='" & GetNewFagent(txt1(14)) & "' OR SP59<='" & GetNewFagent(txt1(14)) & "' OR SP65<='" & GetNewFagent(txt1(14)) & "' OR SP66<='" & GetNewFagent(txt1(14)) & "') "
           End If
       End If
   End If
   If Len(Trim(txt1(13))) <> 0 Or Len(Trim(txt1(14))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label14 & txt1(13) & "-" & txt1(14) 'Add By Sindy 2010/12/10
   End If
   If Len(Trim(txt1(15))) <> 0 And Len(Trim(txt1(16))) <> 0 Then
       StrTest1 = StrTest1 & " AND ((PA75>='" & GetNewFagent(txt1(15)) & "' AND PA75<='" & GetNewFagent(txt1(16)) & "') OR (CP44>='" & GetNewFagent(txt1(15)) & "' AND CP44<='" & GetNewFagent(txt1(16)) & "')) "
       StrTest2 = StrTest2 & " AND ((TM44>='" & GetNewFagent(txt1(15)) & "' AND TM44<='" & GetNewFagent(txt1(16)) & "') OR (CP44>='" & GetNewFagent(txt1(15)) & "' AND CP44<='" & GetNewFagent(txt1(16)) & "')) "
       StrTest3 = StrTest3 & " AND ((SP26>='" & GetNewFagent(txt1(15)) & "' AND SP26<='" & GetNewFagent(txt1(16)) & "') OR (CP44>='" & GetNewFagent(txt1(15)) & "' AND CP44<='" & GetNewFagent(txt1(16)) & "')) "
   Else
       If Len(Trim(txt1(15))) <> 0 And Len(Trim(txt1(16))) = 0 Then
           StrTest1 = StrTest1 & " AND (PA75>='" & GetNewFagent(txt1(15)) & "' OR CP44>='" & GetNewFagent(txt1(15)) & "') "
           StrTest2 = StrTest2 & " AND (TM44>='" & GetNewFagent(txt1(15)) & "' OR CP44>='" & GetNewFagent(txt1(15)) & "') "
           StrTest3 = StrTest3 & " AND (SP26>='" & GetNewFagent(txt1(15)) & "' OR CP44>='" & GetNewFagent(txt1(15)) & "') "
       Else
           If Len(Trim(txt1(15))) = 0 And Len(Trim(txt1(16))) <> 0 Then
               StrTest1 = StrTest1 & " AND (PA75<='" & GetNewFagent(txt1(16)) & "' OR CP44<='" & GetNewFagent(txt1(16)) & "') "
               StrTest2 = StrTest2 & " AND (TM44<='" & GetNewFagent(txt1(16)) & "' OR CP44<='" & GetNewFagent(txt1(16)) & "') "
               StrTest3 = StrTest3 & " AND (SP26<='" & GetNewFagent(txt1(16)) & "' OR CP44<='" & GetNewFagent(txt1(16)) & "') "
           End If
       End If
   End If
   If Len(Trim(txt1(15))) <> 0 Or Len(Trim(txt1(16))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label15 & txt1(15) & "-" & txt1(16) 'Add By Sindy 2010/12/10
   End If
   'Add by Morgan 2007/8/2
   If Not (txt1(7) = "201" And txt1(8) = txt1(7)) Then
      'Modify by Morgan 2010/3/23 翻譯排除無承辦或承辦為外翻人員者(F開頭且無所內編號或所內編號已離職)
      'stConCP = stConCP & " and not exists(select * from staff where st01=cp14 and ST15='F51')"
      stConCP = stConCP & " AND NOT ( CP10='201' AND (CP14 IS NULL OR (SUBSTRB(CP14,1,1)='F' AND (ST04 IS NULL OR ST04='2'))))"
   End If
   
   StrTemp4 = DateSerial(Year(Now), Month(Now), Day(Now) + (Val(txt1(5)) * -1))
   Day1 = ChangeWDateStringToWString(StrTemp4)
   StrTemp4 = DateSerial(Year(Now), Month(Now), Day(Now) + (Val(txt1(6)) * -1))
   Day2 = ChangeWDateStringToWString(StrTemp4)
   pub_QL05 = pub_QL05 & ";" & Label7 & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/12/10
   
   'Modify by Morgan 2004/4/2
   '加核稿人及不續辦檢查
   'strSQL = "SELECT nvl(S1.ST02,cp14) ,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)),CP06,CP07,decode(pa09,'000',PTM03,ptm04),nvl(decode(pa09,'000',CPM03,cpm04),cp10),nvl(S2.ST02,cp13),CP64,CP14 AS A,CP13 AS D, CP10,PA10 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL  AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) and cp14=s1.st01(+)  and cp13=s2.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND (PA57<>'Y' OR PA57 IS NULL) And (S1.ST15>='F' And S1.ST15<='F99') " & StrTest1
   'strSQL = strSQL & " union all select nvl(S1.ST02,cp14) ,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(TM05,NVL(TM06,TM07)),CP06,CP07,decode(tm10,'000',PTM03,ptm04),nvl(decode(tm10,'000',CPM03,cpm04),cp10),nvl(S2.ST02,cp13),CP64,CP14 AS A,CP13 AS D, CP10, TM11 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) and cp14=s1.st01(+)  and cp13=s2.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) and '2'=ptm01(+)  AND TM08=PTM02(+)  AND (TM29<>'Y' OR TM29 IS NULL) And (S1.ST15>='F' And S1.ST15<='F99') " & StrTest2
   'strSQL = strSQL & " union all select nvl(S1.ST02,cp14) ,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(SP05,NVL(SP06,SP07)),CP06,CP07,'',nvl(decode(sp09,'000',CPM03,cpm04),cp10),nvl(S2.ST02,cp13),CP64,CP14 AS A,CP13 AS D, CP10, SP10 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) and cp14=s1.st01(+)  and cp13=s2.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND SP15<>'Y' And (S1.ST15>='F' And S1.ST15<='F99') " & StrTest3
   
   'Add by Morgan 2007/8/2
   'Modify by Morgan 2010/3/23 無承辦人的也要
   'stVTB = "SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP13,NVL(SIM01,CP14) CP14,CP64" & _
      " FROM CASEPROGRESS,STAFF_IDMAP WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL AND SIM02(+)=CP14" & stConCP
   'Modified by Lydia 2016/12/21 CP27 IS NULL AND CP57 IS NULL => CP158=0 AND CP159=0
   stVTB = "SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP13,NVL(SIM01,CP14) CP14,CP64" & _
      " FROM CASEPROGRESS,STAFF_IDMAP,STAFF,ENGINEERPROGRESS WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP158=0 AND CP159=0" & _
      " AND SIM02(+)=CP14 AND ST01(+)=SIM01 AND EP02(+)=CP09 " & stConCP
   
   'Added by Morgan 2012/5/24 FMP排除翻譯已核稿完成或非翻譯已完稿程序
   'Modify By Sindy 2023/10/30 EP33要回歸用在英文核完日,改抓EP39.核稿完成日
   stVTB = stVTB & " and not (cp01 in ('P','PS','CFP','CPS') and substr(cp12,1,1)='F' and ((cp10<>'201' and ep09 is not null) or (cp10='201' and " & IIf(strSrvDate(1) >= FCP核完日改用EP39, "ep39", "ep33") & " is not null))) "
   '2012/7/30 add by sonia 勾選是否含核對已准專利無承辦期限資料
   If Check1.Value = 0 Then
      stVTB = stVTB & " and not (cp10='926' and cp48 is null) "
      pub_QL05 = pub_QL05 & ";" & Check1.Caption
   Else
      
   End If
   '2012/7/30 end
      
   'Modify by Morgan 2007/5/30 加部門&組別
   'Modify by Morgan 2007/8/2 外譯編號要轉為員工編號,否則因組別不同資料會分開
   'strSQL = "SELECT nvl(S1.ST02,cp14) C01, CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)) C02,CP06,CP07,decode(pa09,'000',PTM03,ptm04) C03,nvl(decode(pa09,'000',CPM03,cpm04),cp10) C04,nvl(S2.ST02,cp13) C05,CP64,CP14 AS A,CP13 AS D, CP10,PA10, CP09,s1.ST15 dep1,s1.st16 grp1,s2.ST15 dep2,s2.st16 grp2 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL  AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) and cp14=s1.st01(+)  and cp13=s2.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND (PA57<>'Y' OR PA57 IS NULL) And (S1.ST15>='F' And S1.ST15<='F99') " & StrTest1
   'strSQL = strSQL & " union all select nvl(S1.ST02,cp14) ,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(SP05,NVL(SP06,SP07)),CP06,CP07,'',nvl(decode(sp09,'000',CPM03,cpm04),cp10),nvl(S2.ST02,cp13),CP64,CP14 AS A,CP13 AS D, CP10, SP10, CP09,s1.ST15 dep1,s1.st16 grp2,s2.ST15 dep2,s2.st16 grp2 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) and cp14=s1.st01(+)  and cp13=s2.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND SP15<>'Y' And (S1.ST15>='F' And S1.ST15<='F99') " & StrTest3
   'Modify by Morgan 2010/3/23 無承辦人的也要
   'strSql = "SELECT nvl(S1.ST02,cp14) C01, CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)) C02,CP06,CP07,decode(pa09,'000',PTM03,ptm04) C03,nvl(decode(pa09,'000',CPM03,cpm04),cp10) C04,nvl(S2.ST02,cp13) C05,CP64,CP14 AS A,CP13 AS D, CP10,PA10, CP09,s1.ST15 dep1,s1.st16 grp1,s2.ST15 dep2,s2.st16 grp2 FROM (" & stVTB & stConCP1 & ") X,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP WHERE PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND (PA57 IS NULL OR PA57<>'Y') AND S1.ST01(+)=CP14 and s2.st01(+)=CP13  AND cpm01(+)=cp01 AND cpm02(+)=cp10 AND PTM01(+)='1' AND PTM02(+)=PA08 And (S1.ST15>='F' And S1.ST15<='F99') " & StrTest1
   'strSql = strSql & " union all select nvl(S1.ST02,cp14) ,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(SP05,NVL(SP06,SP07)),CP06,CP07,'',nvl(decode(sp09,'000',CPM03,cpm04),cp10),nvl(S2.ST02,cp13),CP64,CP14 AS A,CP13 AS D, CP10, SP10, CP09,s1.ST15 dep1,s1.st16 grp2,s2.ST15 dep2,s2.st16 grp2 FROM (" & stVTB & stConCP3 & ") X,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE SP01(+)=cp01 AND SP02(+)=cp02 AND SP03(+)=cp03 AND SP04(+)=cp04 AND (SP15 IS NULL OR SP15<>'Y') AND S1.ST01(+)=CP14 and s2.st01(+)=cp13  AND cpm01(+)=cp01 AND cpm02(+)=cp10  And (S1.ST15>='F' And S1.ST15<='F99') " & StrTest3
   'Modified by Lydia 2015/09/09 國外部收的澳門大陸案關聯,香港案110(無期限,不顯示) ,以及P,FCP之分割案無期限,不顯示
   'Modified by Lydia 2015/09/09 +專利案431(PPH)無期限不出現
'   strSql = "SELECT nvl(S1.ST02,cp14) C01, CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)) C02,CP06,CP07,decode(pa09,'000',PTM03,ptm04) C03,nvl(decode(pa09,'000',CPM03,cpm04),cp10) C04,nvl(S2.ST02,cp13) C05,CP64,CP14 AS A,CP13 AS D, CP10,PA10, CP09,s1.ST15 dep1,s1.st16 grp1,s2.ST15 dep2,s2.st16 grp2 FROM (" & stVTB & stConCP1 & ") X,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP WHERE PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND (PA57 IS NULL OR PA57<>'Y') AND S1.ST01(+)=CP14 and s2.st01(+)=CP13  AND cpm01(+)=cp01 AND cpm02(+)=cp10 AND PTM01(+)='1' AND PTM02(+)=PA08 And (CP14 IS NULL OR (S1.ST15>='F' And S1.ST15<='F99')) " & StrTest1
'   strSql = strSql & " union all select nvl(S1.ST02,cp14) ,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(SP05,NVL(SP06,SP07)),CP06,CP07,'',nvl(decode(sp09,'000',CPM03,cpm04),cp10),nvl(S2.ST02,cp13),CP64,CP14 AS A,CP13 AS D, CP10, SP10, CP09,s1.ST15 dep1,s1.st16 grp2,s2.ST15 dep2,s2.st16 grp2 FROM (" & stVTB & stConCP3 & ") X,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2 WHERE SP01(+)=cp01 AND SP02(+)=cp02 AND SP03(+)=cp03 AND SP04(+)=cp04 AND (SP15 IS NULL OR SP15<>'Y') AND S1.ST01(+)=CP14 and s2.st01(+)=cp13  AND cpm01(+)=cp01 AND cpm02(+)=cp10  And (CP14 IS NULL OR (S1.ST15>='F' And S1.ST15<='F99')) " & StrTest3
   'Modified by Lydia 2019/11/01 增加欄位:申請人1~5(cust01~cust05),FC代理人
   strSql = "SELECT nvl(S1.ST02,cp14) C01, CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)) C02,CP06,CP07,decode(pa09,'000',PTM03,ptm04) C03,nvl(decode(pa09,'000',CPM03,cpm04),cp10) C04,nvl(S2.ST02,cp13) C05,CP64,CP14 AS A,CP13 AS D, CP10,PA10, CP09,s1.ST15 dep1,s1.st16 grp1,s2.ST15 dep2,s2.st16 grp2 " & _
            ",PA26 AS CUST01,PA27 AS CUST02,PA28 AS CUST03,PA29 AS CUST04,PA30 AS CUST05,PA75 AS FCNO " & _
            "FROM (" & stVTB & stConCP1 & ") X,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP " & _
            ",(select cp01 v1c1,cp02 v1c2,cp03 v1c3,cp04 v1c4,cp06 v1c6,cp07 v1c7,cp12 v1c8 from casemap,caseprogress where cm10 in ('4','5') and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and ((cm10='4' and cp10='110') or (cm10='5' and cp10 in (" & CaseMapIn & "))) ) VT1 " & _
            ",(select cp01 v2c1,cp02 v2c2,cp03 v2c3,cp04 v2c4,cp06 v2c6,cp07 v2c7,cp12 v2c8 from divisioncase,caseprogress where dc01 in ('P','FCP') and dc01=cp01(+) and dc02=cp02(+) and dc03=cp03(+) and dc04=cp04(+) and cp10 = '307' ) VT2 " & _
            "WHERE PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND (PA57 IS NULL OR PA57<>'Y') AND S1.ST01(+)=CP14 and s2.st01(+)=CP13  AND cpm01(+)=cp01 AND cpm02(+)=cp10 AND PTM01(+)='1' AND PTM02(+)=PA08 And (CP14 IS NULL OR (S1.ST15>='F' And S1.ST15<='F99')) " & _
            " and cp01=v1c1(+) and cp02=v1c2(+) and cp03=v1c3(+) and cp04=v1c4(+) and cp01=v2c1(+) and cp02=v2c2(+) and cp03=v2c3(+) and cp04=v2c4(+)" & _
            " and decode(v1c1||v2c1,null,1,decode(substr(v1c6||v1c8,1,1),'F',0,decode(substr(v2c6||v2c8,1,1),'F',0,1)))=1 and decode(cp10||cp06,'431',0,1)=1 " & StrTest1
   strSql = strSql & " union all select nvl(S1.ST02,cp14) ,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(SP05,NVL(SP06,SP07)),CP06,CP07,'',nvl(decode(sp09,'000',CPM03,cpm04),cp10),nvl(S2.ST02,cp13),CP64,CP14 AS A,CP13 AS D, CP10, SP10, CP09,s1.ST15 dep1,s1.st16 grp2,s2.ST15 dep2,s2.st16 grp2 " & _
            ",SP08 AS CUST01,SP58 AS CUST02,SP59 AS CUST03,SP65 AS CUST04,SP66 AS CUST05,SP26 AS FCNO " & _
            "FROM (" & stVTB & stConCP3 & ") X,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2 " & _
            "WHERE SP01(+)=cp01 AND SP02(+)=cp02 AND SP03(+)=cp03 AND SP04(+)=cp04 AND (SP15 IS NULL OR SP15<>'Y') AND S1.ST01(+)=CP14 " & _
            "and s2.st01(+)=cp13  AND cpm01(+)=cp01 AND cpm02(+)=cp10  And (CP14 IS NULL OR (S1.ST15>='F' And S1.ST15<='F99')) " & StrTest3
   
   'Remove by Morgan 2007/8/2 不會有商標案
   'strSQL = strSQL & " union all select nvl(S1.ST02,cp14) ,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(TM05,NVL(TM06,TM07)),CP06,CP07,decode(tm10,'000',PTM03,ptm04),nvl(decode(tm10,'000',CPM03,cpm04),cp10),nvl(S2.ST02,cp13),CP64,CP14 AS A,CP13 AS D, CP10, TM11, CP09,s1.ST15 dep1,s1.st16 grp1,s2.ST15 dep2,s2.st16 grp2 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) and cp14=s1.st01(+)  and cp13=s2.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) and '2'=ptm01(+)  AND TM08=PTM02(+)  AND (TM29<>'Y' OR TM29 IS NULL) And (S1.ST15>='F' And S1.ST15<='F99') " & stConCP2 & stConCP & StrTest2
   'Modified by Morgan 2012/12/24 延期有可能1次以上,加 distinct
   If txt2 = "1" Then
      'Modified by Lydia 2016/12/21 +排除D類收文 X.CP09>'C' => SUBSTR(X.CP09,1,1)='C'
      strSql = "SELECT distinct C01 R01,SUBSTR(B,1,4)-1911||'/'||SUBSTR(B,5,2)||'/'||SUBSTR(B,7,2) R02" & _
           ",DECODE(PA10,NULL,'',SUBSTR(PA10,1,4)-1911||'/'||SUBSTR(PA10,5,2)||'/'||SUBSTR(PA10,7,2)) R03" & _
           ",DECODE(INSTR('201,210',X.CP10),0,'  ',DECODE(Y.CP01,NULL,'  ','*'))||C R04" & _
           ",C02 R05" & _
           ",DECODE(X.CP06,NULL,'',SUBSTR(X.CP06,1,4)-1911||'/'||SUBSTR(X.CP06,5,2)||'/'||SUBSTR(X.CP06,7,2)) R06" & _
           ",DECODE(X.CP07,NULL,'',SUBSTR(X.CP07,1,4)-1911||'/'||SUBSTR(X.CP07,5,2)||'/'||SUBSTR(X.CP07,7,2)) R07" & _
           ",C03 R08,C04 R09,C05 R10" & _
           ",DECODE(EP09,NULL,'',SUBSTR(EP09,1,4)-1911||'/'||SUBSTR(EP09,5,2)||'/'||SUBSTR(EP09,7,2)) R11,A R12,cst16(grp1) R13" & _
           ",X.* FROM (" & strSql & ") X, ENGINEERPROGRESS, CASEPROGRESS Y" & _
           " WHERE NOT EXISTS( SELECT * FROM NEXTPROGRESS WHERE NP01=X.CP09 AND NP06='N' AND SUBSTR(X.CP09,1,1)='C')" & _
           " AND EP02(+)=X.CP09 AND Y.CP43(+)=X.CP09 AND Y.CP10(+)='404'"

      '2010/9/16 MODIFY BY SONIA 排序加X.CP09
      If txt1(17) = "1" Then
         pub_QL05 = pub_QL05 & ";" & Label8(1) & "1.承辦人" 'Add By Sindy 2010/12/21
         strSql = strSql & " ORDER BY dep1,grp1,A,B,C,X.CP09 "
         iSort = 1
      Else
         pub_QL05 = pub_QL05 & ";" & Label8(1) & "2.智權人員" 'Add By Sindy 2010/12/21
         strSql = strSql & " ORDER BY dep2,grp2,D,B,C,X.CP09 "
         iSort = 2
      End If
   Else
      'Modify by Morgan 2005/4/20 翻譯,檢視中說若有延期加 * 號
      'strSQL = "SELECT X.*,EP09 FROM (" & strSQL & ") X, ENGINEERPROGRESS" & _
         " WHERE NOT EXISTS( SELECT * FROM NEXTPROGRESS WHERE NP01=CP09 AND NP06='N' AND CP09>'C')" & _
         " AND EP02(+)=CP09"
      'Modified by Lydia 2016/12/21 +排除D類收文 X.CP09>'C' => SUBSTR(X.CP09,1,1)='C'
      strSql = "SELECT distinct X.*,EP09,DECODE(INSTR('201,210',X.CP10),0,NULL,DECODE(Y.CP01,NULL,NULL,'*')) MK FROM (" & strSql & ") X, ENGINEERPROGRESS, CASEPROGRESS Y" & _
         " WHERE NOT EXISTS( SELECT * FROM NEXTPROGRESS WHERE NP01=X.CP09 AND NP06='N' AND SUBSTR(X.CP09,1,1)='C')" & _
         " AND EP02(+)=X.CP09 AND Y.CP43(+)=X.CP09 AND Y.CP10(+)='404'"
      '2005/4/20 END
      
      'Modify by Morgan 2007/5/30 排序加先依部門&組別
      If txt1(17) = "1" Then
         strSql = strSql & " ORDER BY dep1,grp1,A,B,C "
         pub_QL05 = pub_QL05 & ";" & Label8(1) & "1.承辦人" 'Add By Sindy 2010/12/10
      Else
         strSql = strSql & " ORDER BY dep2,grp2,D,B,C "
         pub_QL05 = pub_QL05 & ";" & Label8(1) & "2.智權人員" 'Add By Sindy 2010/12/10
      End If
   End If
   
   intI = 1
   'Modified by Lydia 2019/11/01 改變型態
   'Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
   'If intI = 1 Then
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
      dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3
      
         'Added by Lydia 2019/11/01 逐案號判斷
         If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            adoRecordset.MoveFirst
            Do While adoRecordset.EOF = False
                '利益衝突案件：逐案號判斷
                If PUB_ChkCufaByCase(Me.Name, m_AllSys, Trim("" & adoRecordset.Fields("C")), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
                    intCufaCnt = intCufaCnt + 1
                    adoRecordset.Delete
                End If
                adoRecordset.MoveNext
            Loop
            '利益衝突案件：限閱案件
            If intCufaCnt > 0 Then
               pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
               MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
            End If
            InsertQueryLog (dblRow) 'Add By Sindy 2010/12/10
            If adoRecordset.RecordCount = 0 Then
                  GoTo JumpToNoData
            End If
         Else
            InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/10
         End If
        'end 2019/11/01
       
       If txt2 = "1" Then
         SetGrid adoRecordset, iSort
       Else
         StrPrintDoc       '列印主程式
       End If
       CheckOC
   Else
       InsertQueryLog (0) 'Add By Sindy 2010/12/10
JumpToNoData:   'Added by Lydia 2019/11/01
       ShowNoData
       CheckOC
       Screen.MousePointer = vbDefault
       Exit Sub
   End If
   Screen.MousePointer = vbDefault
End Sub

'輸入申請人及代理人條件
Sub StrMenu2()
Dim dblRow As Double 'Add By Sindy 2025/9/3
   
   Screen.MousePointer = vbHourglass
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/10 清除查詢印表記錄檔欄位
   StrTest1 = ""
   StrTest2 = ""
   StrTest3 = ""
   'Added by Lydia 2019/11/01 利益衝突案件
   m_AllSys = IIf(txt1(0) <> "ALL", txt1(0), GetAllSysKind(, txt1(0)))
   intCufaCnt = 0
   'end 2019/11/01
   
   If Len(txt1(0)) <> 0 Then
      StrTest1 = StrTest1 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
      StrTest2 = StrTest2 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
      StrTest3 = StrTest3 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 3) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(0) 'Add By Sindy 2010/12/10
   End If
   If Len(txt1(1)) <> 0 Then
       StrTest1 = StrTest1 & " AND CP12>='" & txt1(1) & "' "
       StrTest2 = StrTest2 & " AND CP12>='" & txt1(1) & "' "
       StrTest3 = StrTest3 & " AND CP12>='" & txt1(1) & "' "
   End If
   If Len(txt1(2)) <> 0 Then
       StrTest1 = StrTest1 & " AND CP12<='" & txt1(2) & "' "
       StrTest2 = StrTest2 & " AND CP12<='" & txt1(2) & "' "
       StrTest3 = StrTest3 & " AND CP12<='" & txt1(2) & "' "
   End If
   If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/10
   End If
   If Len(txt1(3)) <> 0 Then
       StrTest1 = StrTest1 & " AND CP13>='" & txt1(3) & "' "
       StrTest2 = StrTest2 & " AND CP13>='" & txt1(3) & "' "
       StrTest3 = StrTest3 & " AND CP13>='" & txt1(3) & "' "
       pub_QL05 = pub_QL05 & ";" & Label3 & txt1(3) & LBL1(0) 'Add By Sindy 2010/12/10
   End If
   If Len(txt1(4)) <> 0 Then
      'Modify by Morgan 2007/8/6 加外譯編號
      strExc(1) = PUB_GetMapID(txt1(4), 0)
      If strExc(1) <> "" Then
         StrTest1 = StrTest1 & " AND CP14 in ('" & txt1(4) & "','" & strExc(1) & "')"
         StrTest2 = StrTest2 & " AND CP14 in ('" & txt1(4) & "','" & strExc(1) & "')"
         StrTest3 = StrTest3 & " AND CP14 in ('" & txt1(4) & "','" & strExc(1) & "')"
      Else
         StrTest1 = StrTest1 & " AND CP14='" & txt1(4) & "' "
         StrTest2 = StrTest2 & " AND CP14='" & txt1(4) & "' "
         StrTest3 = StrTest3 & " AND CP14='" & txt1(4) & "' "
      End If
      pub_QL05 = pub_QL05 & ";" & Label5 & txt1(4) & LBL1(1) 'Add By Sindy 2010/12/10
   End If
   If Len(txt1(7)) <> 0 Then
       StrTest1 = StrTest1 & " AND CP10>='" & txt1(7) & "' "
       StrTest2 = StrTest2 & " AND CP10>='" & txt1(7) & "' "
       StrTest3 = StrTest3 & " AND CP10>='" & txt1(7) & "' "
   End If
   If Len(txt1(8)) <> 0 Then
       StrTest1 = StrTest1 & " AND CP10<='" & txt1(8) & "' "
       StrTest2 = StrTest2 & " AND CP10<='" & txt1(8) & "' "
       StrTest3 = StrTest3 & " AND CP10<='" & txt1(8) & "' "
   End If
   If Len(txt1(7)) <> 0 Or Len(txt1(8)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label8(0) & txt1(7) & "-" & txt1(8) 'Add By Sindy 2010/12/10
   End If
   If Len(txt1(9)) <> 0 Then
       StrTest1 = StrTest1 & " AND SUBSTR(PA09,1,3)>='" & txt1(9) & "' "
       StrTest2 = StrTest2 & " AND SUBSTR(TM10,1,3)>='" & txt1(9) & "' "
       StrTest3 = StrTest3 & " AND SUBSTR(SP09,1,3)>='" & txt1(9) & "' "
   End If
   If Len(txt1(10)) <> 0 Then
       StrTest1 = StrTest1 & " AND SUBSTR(PA09,1,3)>='" & txt1(10) & "' "
       StrTest2 = StrTest2 & " AND SUBSTR(TM10,1,3)>='" & txt1(10) & "' "
       StrTest3 = StrTest3 & " AND SUBSTR(SP09,1,3)>='" & txt1(10) & "' "
   End If
   If Len(txt1(9)) <> 0 Or Len(txt1(10)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label9 & txt1(9) & "-" & txt1(10) 'Add By Sindy 2010/12/10
   End If
   '申請人
   If Len(Trim(txt1(13))) <> 0 And Len(Trim(txt1(14))) <> 0 Then
       StrTest1 = StrTest1 & " AND ((PA26>='" & GetNewFagent(txt1(13)) & "' AND PA26<='" & GetNewFagent(txt1(14)) & "') OR (PA27>='" & GetNewFagent(txt1(13)) & "' AND PA27<='" & GetNewFagent(txt1(14)) & "') OR (PA28>='" & GetNewFagent(GetNewFagent(txt1(13))) & "' AND PA28<='" & GetNewFagent(txt1(14)) & "') OR (PA29>='" & GetNewFagent(GetNewFagent(txt1(13))) & "' AND PA29<='" & GetNewFagent(txt1(14)) & "') OR (PA30>='" & GetNewFagent(GetNewFagent(txt1(13))) & "' AND PA30<='" & GetNewFagent(txt1(14)) & "')) "
       'Modified by Lydia 2019/11/01 補上申請人1~5
       'StrTest2 = StrTest2 & " AND (TM23>='" & GetNewFagent(Txt1(13)) & "' AND TM23<='" & GetNewFagent(Txt1(14)) & "') "
       'StrTest3 = StrTest3 & " AND ((SP08>='" & GetNewFagent(Txt1(13)) & "' AND SP08<='" & GetNewFagent(Txt1(14)) & "') OR (SP58<='" & GetNewFagent(Txt1(13)) & "' AND SP58<='" & GetNewFagent(Txt1(14)) & "') OR (SP59>='" & GetNewFagent(Txt1(13)) & "' AND SP59<='" & GetNewFagent(Txt1(14)) & "')) "
       StrTest2 = StrTest2 & " AND ((TM23>='" & GetNewFagent(txt1(13)) & "' AND TM23<='" & GetNewFagent(txt1(14)) & "') OR (TM78>='" & GetNewFagent(txt1(13)) & "' AND TM78<='" & GetNewFagent(txt1(14)) & "') OR (TM79>='" & GetNewFagent(txt1(13)) & "' AND TM79<='" & GetNewFagent(txt1(14)) & "') OR (TM80>='" & GetNewFagent(txt1(13)) & "' AND TM80<='" & GetNewFagent(txt1(14)) & "') OR (TM81>='" & GetNewFagent(txt1(13)) & "' AND TM81<='" & GetNewFagent(txt1(14)) & "')) "
       StrTest3 = StrTest3 & " AND ((SP08>='" & GetNewFagent(txt1(13)) & "' AND SP08<='" & GetNewFagent(txt1(14)) & "') OR (SP58<='" & GetNewFagent(txt1(13)) & "' AND SP58<='" & GetNewFagent(txt1(14)) & "') OR (SP59>='" & GetNewFagent(txt1(13)) & "' AND SP59<='" & GetNewFagent(txt1(14)) & "') OR (SP65<='" & GetNewFagent(txt1(13)) & "' AND SP65<='" & GetNewFagent(txt1(14)) & "') OR (SP66>='" & GetNewFagent(txt1(13)) & "' AND SP66<='" & GetNewFagent(txt1(14)) & "')) "
   Else
       If Len(Trim(txt1(13))) <> 0 And Len(Trim(txt1(14))) = 0 Then
           StrTest1 = StrTest1 & " AND (PA26>='" & GetNewFagent(txt1(13)) & "' OR PA27>='" & GetNewFagent(txt1(13)) & "' OR PA28>='" & GetNewFagent(txt1(13)) & "' OR PA29>='" & GetNewFagent(txt1(13)) & "' OR PA30>='" & GetNewFagent(txt1(13)) & "') "
           'Modified by Lydia 2019/11/01 補上申請人1~5
           'StrTest2 = StrTest2 & " AND (TM23>='" & GetNewFagent(txt1(13)) & "' ) "
           'StrTest3 = StrTest3 & " AND (SP08>='" & GetNewFagent(txt1(13)) & "' OR SP58>='" & GetNewFagent(txt1(13)) & "' OR SP59>='" & GetNewFagent(txt1(13)) & "') "
           StrTest2 = StrTest2 & " AND (TM23>='" & GetNewFagent(txt1(13)) & "' OR TM78>='" & GetNewFagent(txt1(13)) & "' OR TM79>='" & GetNewFagent(txt1(13)) & "' OR TM80>='" & GetNewFagent(txt1(13)) & "' OR TM81>='" & GetNewFagent(txt1(13)) & "') "
           StrTest3 = StrTest3 & " AND (SP08>='" & GetNewFagent(txt1(13)) & "' OR SP58>='" & GetNewFagent(txt1(13)) & "' OR SP59>='" & GetNewFagent(txt1(13)) & "' OR SP65>='" & GetNewFagent(txt1(13)) & "' OR SP66>='" & GetNewFagent(txt1(13)) & "') "
       Else
           If Len(Trim(txt1(13))) = 0 And Len(Trim(txt1(14))) <> 0 Then
               StrTest1 = StrTest1 & " AND (PA26<='" & GetNewFagent(txt1(14)) & "' OR PA27<='" & GetNewFagent(txt1(14)) & "' OR PA28<='" & GetNewFagent(txt1(14)) & "' OR PA29<='" & GetNewFagent(txt1(14)) & "' OR PA30<='" & GetNewFagent(txt1(14)) & "') "
               'Modified by Lydia 2019/11/01 補上申請人1~5
               'StrTest2 = StrTest2 & " AND (TM23<='" & GetNewFagent(txt1(14)) & "') "
               'StrTest3 = StrTest3 & " AND (SP08<='" & GetNewFagent(txt1(14)) & "' OR SP58<='" & GetNewFagent(txt1(14)) & "' OR SP59<='" & GetNewFagent(txt1(14)) & "') "
                StrTest2 = StrTest2 & " AND (TM23<='" & GetNewFagent(txt1(14)) & "' OR TM78<='" & GetNewFagent(txt1(14)) & "' OR TM79<='" & GetNewFagent(txt1(14)) & "' OR TM80<='" & GetNewFagent(txt1(14)) & "' OR TM81<='" & GetNewFagent(txt1(14)) & "') "
                StrTest3 = StrTest3 & " AND (SP08<='" & GetNewFagent(txt1(14)) & "' OR SP58<='" & GetNewFagent(txt1(14)) & "' OR SP59<='" & GetNewFagent(txt1(14)) & "' OR SP65<='" & GetNewFagent(txt1(14)) & "' OR SP66<='" & GetNewFagent(txt1(14)) & "') "
               
           End If
       End If
   End If
   If Len(Trim(txt1(13))) <> 0 Or Len(Trim(txt1(14))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label14 & txt1(13) & "-" & txt1(14) 'Add By Sindy 2010/12/10
   End If
   
   '代理人
   If Len(Trim(txt1(15))) <> 0 And Len(Trim(txt1(16))) <> 0 Then
       StrTest1 = StrTest1 & " AND ((PA75>='" & GetNewFagent(txt1(15)) & "' AND PA75<='" & GetNewFagent(txt1(16)) & "') OR (CP44>='" & GetNewFagent(txt1(15)) & "' AND CP44<='" & GetNewFagent(txt1(16)) & "')) "
       StrTest2 = StrTest2 & " AND ((TM44>='" & GetNewFagent(txt1(15)) & "' AND TM44<='" & GetNewFagent(txt1(16)) & "') OR (CP44>='" & GetNewFagent(txt1(15)) & "' AND CP44<='" & GetNewFagent(txt1(16)) & "')) "
       StrTest3 = StrTest3 & " AND ((SP26>='" & GetNewFagent(txt1(15)) & "' AND SP26<='" & GetNewFagent(txt1(16)) & "') OR (CP44>='" & GetNewFagent(txt1(15)) & "' AND CP44<='" & GetNewFagent(txt1(16)) & "')) "
   Else
       If Len(Trim(txt1(15))) <> 0 And Len(Trim(txt1(16))) = 0 Then
           StrTest1 = StrTest1 & " AND (PA75>='" & GetNewFagent(txt1(15)) & "' OR CP44>='" & GetNewFagent(txt1(15)) & "') "
           StrTest2 = StrTest2 & " AND (TM44>='" & GetNewFagent(txt1(15)) & "' OR CP44>='" & GetNewFagent(txt1(15)) & "') "
           StrTest3 = StrTest3 & " AND (SP26>='" & GetNewFagent(txt1(15)) & "' OR CP44>='" & GetNewFagent(txt1(15)) & "') "
       Else
           If Len(Trim(txt1(15))) = 0 And Len(Trim(txt1(16))) <> 0 Then
               StrTest1 = StrTest1 & " AND (PA75<='" & GetNewFagent(txt1(16)) & "' OR CP44<='" & GetNewFagent(txt1(16)) & "') "
               StrTest2 = StrTest2 & " AND (TM44<='" & GetNewFagent(txt1(16)) & "' OR CP44<='" & GetNewFagent(txt1(16)) & "') "
               StrTest3 = StrTest3 & " AND (SP26<='" & GetNewFagent(txt1(16)) & "' OR CP44<='" & GetNewFagent(txt1(16)) & "') "
           End If
       End If
   End If
   If Len(Trim(txt1(15))) <> 0 Or Len(Trim(txt1(16))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label15 & txt1(15) & "-" & txt1(16) 'Add By Sindy 2010/12/10
   End If
   '2008/2/22 ADD BY SONIA 加組別條件
   If Len(txt1(18)) <> 0 Then
       StrTest1 = StrTest1 & " AND S1.ST16='" & txt1(18) & "' "
       StrTest2 = StrTest2 & " AND S1.ST16='" & txt1(18) & "' "
       StrTest3 = StrTest3 & " AND S1.ST16='" & txt1(18) & "' "
       pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(18) & "( 1電子電機 2化學 3日文 4機械設計 5其他)" 'Add By Sindy 2010/12/10
   End If
   '2008/2/22 END
   StrTemp4 = DateSerial(Year(Now), Month(Now), Day(Now) + (Val(txt1(5)) * -1))
   Day1 = ChangeWDateStringToWString(StrTemp4)
   StrTemp4 = DateSerial(Year(Now), Month(Now), Day(Now) + (Val(txt1(6)) * -1))
   Day2 = ChangeWDateStringToWString(StrTemp4)
   pub_QL05 = pub_QL05 & ";" & Label7 & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/12/10
   
   'Modify By Cheng 2003/02/05
   '承辦人為NULL的資料也要出現
   'strSQL = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(PA05,NVL(PA06,PA07)) AS B,'',PA11,CP05,CPM03,CP06,CP07,PA26,PA27,PA28,PA29,PA30,CP44,PA75 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL  AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (PA57<>'Y' OR PA57 IS NULL) " & StrTest1
   'strSQL = strSQL & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(TM05,NVL(TM06,TM07)) AS B,'',TM12,CP05,CPM03,CP06,CP07,TM23,'','','','',CP44,TM44 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL  AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (TM29<>'Y' OR TM29 IS NULL) " & StrTest2
   'strSQL = strSQL & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)) AS B,'',SP11,CP05,CPM03,CP06,CP07,SP08,SP58,SP59,'','',CP44,SP26 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL  AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (SP15<>'Y' OR SP15 IS NULL) " & StrTest3
   'Modify By Cheng 2003/02/18
   '承辦人不可NULL, 且承辦人的部門別在"F"及"F99"間
   'strSQL = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(PA05,NVL(PA06,PA07)) AS B,'',PA11,CP05,CPM03,CP06,CP07,PA26,PA27,PA28,PA29,PA30,CP44,PA75 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP27 IS NULL AND CP57 IS NULL  AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (PA57<>'Y' OR PA57 IS NULL) " & StrTest1
   'strSQL = strSQL & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(TM05,NVL(TM06,TM07)) AS B,'',TM12,CP05,CPM03,CP06,CP07,TM23,'','','','',CP44,TM44 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP27 IS NULL AND CP57 IS NULL  AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (TM29<>'Y' OR TM29 IS NULL) " & StrTest2
   'strSQL = strSQL & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)) AS B,'',SP11,CP05,CPM03,CP06,CP07,SP08,SP58,SP59,'','',CP44,SP26 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP27 IS NULL AND CP57 IS NULL  AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (SP15<>'Y' OR SP15 IS NULL) " & StrTest3
   'Modify by Morgan 2004/4/2
   '加核稿人及不續辦檢查
   'strSQL = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(PA05,NVL(PA06,PA07)) AS B,'',PA11,CP05,CPM03,CP06,CP07,PA26,PA27,PA28,PA29,PA30,CP44,PA75 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL  AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (PA57<>'Y' OR PA57 IS NULL) And (S1.ST15>='F' And S1.ST15<='F99') " & StrTest1
   'strSQL = strSQL & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(TM05,NVL(TM06,TM07)) AS B,'',TM12,CP05,CPM03,CP06,CP07,TM23,'','','','',CP44,TM44 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL  AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (TM29<>'Y' OR TM29 IS NULL) And (S1.ST15>='F' And S1.ST15<='F99')  " & StrTest2
   'strSQL = strSQL & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)) AS B,'',SP11,CP05,CPM03,CP06,CP07,SP08,SP58,SP59,'','',CP44,SP26 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL  AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (SP15<>'Y' OR SP15 IS NULL) And (S1.ST15>='F' And S1.ST15<='F99')  " & StrTest3
   
   'Modify by Morgan 2005/4/20 翻譯,檢視中說若有延期加 * 號
   'strSQL = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(PA05,NVL(PA06,PA07)) AS B,'' C01,PA11,CP05,CPM03,CP06,CP07,PA26,PA27,PA28,PA29,PA30,CP44,PA75, CP09 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL  AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (PA57<>'Y' OR PA57 IS NULL) And (S1.ST15>='F' And S1.ST15<='F99') " & StrTest1
   'strSQL = strSQL & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(TM05,NVL(TM06,TM07)) AS B,'',TM12,CP05,CPM03,CP06,CP07,TM23,'','','','',CP44,TM44, CP09 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL  AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (TM29<>'Y' OR TM29 IS NULL) And (S1.ST15>='F' And S1.ST15<='F99')  " & StrTest2
   'strSQL = strSQL & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)) AS B,'',SP11,CP05,CPM03,CP06,CP07,SP08,SP58,SP59,'','',CP44,SP26, CP09 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL  AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (SP15<>'Y' OR SP15 IS NULL) And (S1.ST15>='F' And S1.ST15<='F99')  " & StrTest3
   'strSQL = "SELECT X.* FROM (" & strSQL & ") X" & _
      " WHERE NOT EXISTS( SELECT * FROM NEXTPROGRESS WHERE NP01=CP09 AND NP06='N' AND CP09>'C')"
   'Modified by Lydia 2015/09/09 國外部收的澳門大陸案關聯,香港案110(無期限,不顯示) ,以及P,FCP之分割案無期限,不顯示
   'Modified by Lydia 2015/09/09 專利案431(PPH)無期限不出現
'   strSql = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(PA05,NVL(PA06,PA07)) AS B,'' C01,PA11,CP05,CPM03,CP06,CP07,PA26,PA27,PA28,PA29,PA30,CP44,PA75, CP09,CP10 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL  AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (PA57<>'Y' OR PA57 IS NULL) And (S1.ST15>='F' And S1.ST15<='F99') " & StrTest1
'   strSql = strSql & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(TM05,NVL(TM06,TM07)) AS B,'',TM12,CP05,CPM03,CP06,CP07,TM23,'','','','',CP44,TM44, CP09,CP10 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL  AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (TM29<>'Y' OR TM29 IS NULL) And (S1.ST15>='F' And S1.ST15<='F99')  " & StrTest2
'   strSql = strSql & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)) AS B,'',SP11,CP05,CPM03,CP06,CP07,SP08,SP58,SP59,'','',CP44,SP26, CP09,CP10 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL  AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (SP15<>'Y' OR SP15 IS NULL) And (S1.ST15>='F' And S1.ST15<='F99')  " & StrTest3
   'Modified by Lydia 2016/12/21 CP27 IS NULL AND CP57 IS NULL => CP158=0 AND CP159=0
   strSql = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(PA05,NVL(PA06,PA07)) AS B,'' C01,PA11,CP05,CPM03,CP06,CP07,PA26,PA27,PA28,PA29,PA30,CP44,PA75, CP09,CP10 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF S1 " & _
            ",(select cp01 v1c1,cp02 v1c2,cp03 v1c3,cp04 v1c4,cp06 v1c6,cp07 v1c7,cp12 v1c8 from casemap,caseprogress where cm10 in ('4','5') and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and ((cm10='4' and cp10='110') or (cm10='5' and cp10 in (" & CaseMapIn & "))) ) VT1 " & _
               ",(select cp01 v2c1,cp02 v2c2,cp03 v2c3,cp04 v2c4,cp06 v2c6,cp07 v2c7,cp12 v2c8 from divisioncase,caseprogress where dc01 in ('P','FCP') and dc01=cp01(+) and dc02=cp02(+) and dc03=cp03(+) and dc04=cp04(+) and cp10 = '307' ) VT2 " & _
            "WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP158=0 AND CP159=0 AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (PA57<>'Y' OR PA57 IS NULL) And (S1.ST15>='F' And S1.ST15<='F99') " & _
            " and cp01=v1c1(+) and cp02=v1c2(+) and cp03=v1c3(+) and cp04=v1c4(+) and cp01=v2c1(+) and cp02=v2c2(+) and cp03=v2c3(+) and cp04=v2c4(+) " & _
            " and decode(v1c1||v2c1,null,1,decode(substr(v1c6||v1c8,1,1),'F',0,decode(substr(v2c6||v2c8,1,1),'F',0,1)))=1 and decode(cp10||cp06,'431',0,1)=1 " & StrTest1
   
   'Modified by Lydia 2019/11/01 補上申請人1~5
   'strSql = strSql & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(TM05,NVL(TM06,TM07)) AS B,'',TM12,CP05,CPM03,CP06,CP07,TM23,'','','','',CP44,TM44, CP09,CP10 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP158=0 AND CP159=0 AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (TM29<>'Y' OR TM29 IS NULL) And (S1.ST15>='F' And S1.ST15<='F99')  " & StrTest2
   'strSql = strSql & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)) AS B,'',SP11,CP05,CPM03,CP06,CP07,SP08,SP58,SP59,'','',CP44,SP26, CP09,CP10 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP158=0 AND CP159=0 AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (SP15<>'Y' OR SP15 IS NULL) And (S1.ST15>='F' And S1.ST15<='F99')  " & StrTest3
   strSql = strSql & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(TM05,NVL(TM06,TM07)) AS B,'',TM12,CP05,CPM03,CP06,CP07,TM23,TM78,TM79,TM80,TM81,CP44,TM44, CP09,CP10 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP158=0 AND CP159=0 AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (TM29<>'Y' OR TM29 IS NULL) And (S1.ST15>='F' And S1.ST15<='F99')  " & StrTest2
   strSql = strSql & " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)) AS B,'',SP11,CP05,CPM03,CP06,CP07,SP08,SP58,SP59,SP65,SP66,CP44,SP26, CP09,CP10 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1 WHERE (CP05 BETWEEN " & Day2 & " AND " & Day1 & ") AND CP14 IS NOT NULL AND CP158=0 AND CP159=0 AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) and cp14=s1.st01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND (SP15<>'Y' OR SP15 IS NULL) And (S1.ST15>='F' And S1.ST15<='F99')  " & StrTest3
   'end 2019/11/01
   
   'Modified by Morgan 2012/12/24 延期有可能1次以上,加 distinct
   If txt2 = "1" Then
      'Modified by Lydia 2016/12/21 +排除D類收文 X.CP09>'C' => SUBSTR(X.CP09,1,1)='C'
      strSql = "SELECT distinct DECODE(INSTR('201,210',X.CP10),0,'  ',DECODE(Y.CP01,NULL,'  ','*'))||A R01" & _
         ",B R02,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) R03,PA11 R04" & _
         ",SUBSTR(X.CP05,1,4)-1911||'/'||SUBSTR(X.CP05,5,2)||'/'||SUBSTR(X.CP05,7,2) R05" & _
         ",CPM03 R06" & _
         ",DECODE(X.CP06,NULL,'',SUBSTR(X.CP06,1,4)-1911||'/'||SUBSTR(X.CP06,5,2)||'/'||SUBSTR(X.CP06,7,2)) R07" & _
         ",DECODE(X.CP07,NULL,'',SUBSTR(X.CP07,1,4)-1911||'/'||SUBSTR(X.CP07,5,2)||'/'||SUBSTR(X.CP07,7,2)) R08" & _
         ",X.* FROM (" & strSql & ") X, CASEPROGRESS Y,CUSTOMER" & _
         " WHERE NOT EXISTS( SELECT * FROM NEXTPROGRESS WHERE NP01=X.CP09 AND NP06='N' AND SUBSTR(X.CP09,1,1)='C') AND Y.CP43(+)=X.CP09 AND Y.CP10(+)='404' AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1)"
   Else
      'Modified by Lydia 2016/12/21 +排除D類收文 X.CP09>'C' => SUBSTR(X.CP09,1,1)='C'
      strSql = "SELECT distinct X.*,DECODE(INSTR('201,210',X.CP10),0,NULL,DECODE(Y.CP01,NULL,NULL,'*')) MK FROM (" & strSql & ") X, CASEPROGRESS Y" & _
         " WHERE NOT EXISTS( SELECT * FROM NEXTPROGRESS WHERE NP01=X.CP09 AND NP06='N' AND SUBSTR(X.CP09,1,1)='C') AND Y.CP43(+)=X.CP09 AND Y.CP10(+)='404'"
      '2005/4/20 end
   End If
   strSql = strSql & " ORDER BY A,B "
   
   intI = 1
   'Modified by Lydia 2019/11/01 改變型態
   'Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
   'If intI = 1 Then
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
      dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3

         'Added by Lydia 2019/11/01 逐案號判斷
         If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            adoRecordset.MoveFirst
            Do While adoRecordset.EOF = False
                '利益衝突案件：逐案號判斷
                If PUB_ChkCufaByCase(Me.Name, m_AllSys, Trim("" & adoRecordset.Fields("A")), "" & adoRecordset.Fields("pa26") & "," & adoRecordset.Fields("pa27") & "," & adoRecordset.Fields("pa28") & "," & adoRecordset.Fields("pa29") & "," & adoRecordset.Fields("pa30"), "" & adoRecordset.Fields("pa75")) = False Then
                    intCufaCnt = intCufaCnt + 1
                    adoRecordset.Delete
                End If
                adoRecordset.MoveNext
            Loop
            '利益衝突案件：限閱案件
            If intCufaCnt > 0 Then
               pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
               MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
            End If
            InsertQueryLog (dblRow) 'Add By Sindy 2010/12/10
            If adoRecordset.RecordCount = 0 Then
                  GoTo JumpToNoData
            End If
         Else
            InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/10
         End If
        'end 2019/11/01

       If txt2 = "1" Then
         SetGrid adoRecordset, 3
       Else
         adoRecordset.MoveFirst
         StrPrintDoc2       '列印主程式
         ShowPrintOk
      End If
      CheckOC
   Else
       InsertQueryLog (0)  'Add By Sindy 2010/12/10
JumpToNoData:   'Added by Lydia 2019/11/01
       ShowNoData
       CheckOC
       Screen.MousePointer = vbDefault
       Exit Sub
   End If
   Screen.MousePointer = vbDefault
End Sub

Sub StrPrintDoc()
'Add By Cheng 2003/02/19
Dim strGroup As String '記錄承辦人或智權人員

   'Add By Cheng 20032/02/19
   '初始化記錄承辦人員或智權人員變數
   strGroup = ""
   GetPrintLeft
   iLine = 1
   Page = 1
   adoRecordset.MoveFirst
   '依承辦人排序
   If txt1(17) = "1" Then
       If Not IsNull(adoRecordset.Fields(0)) Then
           TmpArea = adoRecordset.Fields(0)
       Else
           TmpArea = ""
       End If
   '依智權人員排序
   Else
       If Not IsNull(adoRecordset.Fields(8)) Then
           TmpArea = adoRecordset.Fields(8)
       Else
           TmpArea = ""
       End If
   End If
   
   'Add by Morgan 2007/6/1
   '依承辦人列印排序
   If txt1(17) = "1" Then
       m_Grp = "" & adoRecordset.Fields("grp1")
   '依智權人員列印排序
   Else
       m_Grp = "" & adoRecordset.Fields("grp2")
   End If
   'end 2007/6/1
   
   StrPrintTital TmpArea, str(Page)
   iPrint = 2700
   iTatle = 0       ' 總筆數
   iK = 0           ' 小計
   With adoRecordset
       .MoveFirst
       Do While .EOF = False
           For j = 0 To 9
               If Not IsNull(.Fields(j)) Then
                   strTemp3(j) = .Fields(j)
               Else
                   strTemp3(j) = ""
               End If
           Next j
           
           '依承辦人列印排序
           If txt1(17) = "1" Then
               St = strTemp3(0)
               m_Grp = "" & .Fields("grp1") 'Add by Morgan 2007/6/1
           '依智權人員列印排序
           Else
               St = strTemp3(8)
               m_Grp = "" & .Fields("grp2") 'Add by Morgan 2007/6/1
           End If
           iK = iK + 1
           iTatle = iTatle + 1
           If Len(strTemp3(1)) > 7 Then
               strTemp3(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(1)))
           End If
           If Len(strTemp3(4)) > 7 Then
               strTemp3(4) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(4)))
           End If
           If Len(strTemp3(5)) > 7 Then
               strTemp3(5) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(5)))
           End If
           Printer.CurrentX = PLeft(0)
           Printer.CurrentY = iPrint
           '依承辦人列印排序
           If txt1(17) = "1" Then
               'Modify By Cheng 2003/02/19
               '若承辦人不同
               If strGroup <> strTemp3(0) Then
                  '2012/1/2 MODIFY BY SONIA 承辦人姓名改抓4個字
                  Printer.Print StrToStr(strTemp3(0), 4)
                  strGroup = strTemp3(0)
               End If
           '依智權人員列印排序
           Else
               'Modify By Cheng 2003/02/19
               '若承辦人不同
               If strGroup <> strTemp3(8) Then
                   Printer.Print StrToStr(strTemp3(8), 3)
                   strGroup = strTemp3(8)
               End If
           End If
           Printer.CurrentX = PLeft(1)
           Printer.CurrentY = iPrint
           Printer.Print strTemp3(1)
           'Add By Cheng 2003/03/04
           '若案件性質為翻譯(201), 檢視中說(209), 製作中說(210), 要顯示申請日
           'Modified by Morgan 2013/11/6 +235核對中說格式
           If "" & .Fields(12).Value = 翻譯 Or "" & .Fields(12).Value = 檢視中說 Or "" & .Fields(12).Value = "235" Or "" & .Fields(12).Value = 製作中說 Then
               If "" & .Fields(13).Value <> "" Then
                   Printer.CurrentX = PLeft(10)
                   Printer.CurrentY = iPrint
                   Printer.Print ChangeTStringToTDateString(ChangeWStringToTString("" & .Fields(13).Value))
               End If
           End If
           Printer.CurrentX = PLeft(2)
           Printer.CurrentY = iPrint
           'Modify by Morgan 2005/4/27
           'Printer.Print strTemp3(2)
           Printer.Print "" & .Fields("MK") & strTemp3(2)
           Printer.CurrentX = PLeft(3)
           Printer.CurrentY = iPrint
           Printer.Print StrToStr(strTemp3(3), 12)
           Printer.CurrentX = PLeft(4)
           Printer.CurrentY = iPrint
           Printer.Print strTemp3(4)
           Printer.CurrentX = PLeft(5)
           Printer.CurrentY = iPrint
           Printer.Print strTemp3(5)
           Printer.CurrentX = PLeft(6)
           Printer.CurrentY = iPrint
           Printer.Print StrToStr(strTemp3(6), 4)
           Printer.CurrentX = PLeft(7)
           Printer.CurrentY = iPrint
           Printer.Print StrToStr(strTemp3(7), 4)
           Printer.CurrentX = PLeft(8)
           Printer.CurrentY = iPrint
           If txt1(17) = "1" Then
               Printer.Print StrToStr(strTemp3(8), 4)
           Else
               Printer.Print StrToStr(strTemp3(0), 4)
           End If
           Printer.CurrentX = PLeft(9)
           Printer.CurrentY = iPrint
           'Modify by Morgan 2004/4/2
           'Printer.Print StrToStr(strTemp3(9), 12)
           Printer.Print StrToStr(ChangeTStringToTDateString(ChangeWStringToTString("" & .Fields("EP09"))), 12)
           .MoveNext
           
           If .EOF = False Then
               If txt1(17) = "1" Then
                   If Not IsNull(.Fields(0)) Then
                       StrTest1 = CheckStr(.Fields(0))
                   Else
                       StrTest1 = ""
                   End If
                   m_Grp = "" & .Fields("grp1") 'Add by Morgan 2007/8/1
               Else
                   If Not IsNull(.Fields(8)) Then
                       StrTest1 = CheckStr(.Fields(8))
                   Else
                       StrTest1 = ""
                   End If
                   m_Grp = "" & .Fields("grp2") 'Add by Morgan 2007/8/1
               End If
               If StrTest1 <> St Then
                   iPrint = iPrint + 300
                   Printer.CurrentX = 500
                   Printer.CurrentY = iPrint
                   Printer.Print String(200, "-")
                   iPrint = iPrint + 300
                   Printer.CurrentX = 1000
                   Printer.CurrentY = iPrint
                   Printer.Print "小計： " & Trim(str(iK)) & " 筆"
                   iK = 0
                   iPrint = iPrint + 300
                   Printer.CurrentX = 500
                   Printer.CurrentY = iPrint
                   Printer.Print String(200, "-")
                   iLine = iLine + 3
                   St = StrTest1
                   'Add by Morgan 2007/5/31 加依照員工編號跳頁
                   'Modify by Morgan 2007/8/2 若沒有指定案件性質201時才跳頁
                   If Not (txt1(7) = "201" And txt1(8) = txt1(7)) Then
                     Printer.NewPage
                     Page = 1
                     StrPrintTital TmpArea, str(Page)
                     iPrint = 2400
                     iLine = 0
                   End If
                   'end 2007/5/31
               End If
               If (iLine >= 23) Then
                   Printer.NewPage
                   Page = Page + 1
                   StrPrintTital TmpArea, str(Page)
                   iPrint = 2400
                   iLine = 0
               End If
               iLine = iLine + 1
               iPrint = iPrint + 300
           End If
       Loop
   End With
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = 1000
   Printer.CurrentY = iPrint
   Printer.Print "小計： " & Trim(str(iK)) & " 筆"
   'Remove by Morgan 2007/5/31 改依照員工編號跳頁故不必再印合計
   '合計
   'iPrint = iPrint + 300
   'Printer.CurrentX = 500
   'Printer.CurrentY = iPrint
   'Printer.Print String(200, "-")
   'iPrint = iPrint + 300
   'Printer.CurrentX = 500
   'Printer.CurrentY = iPrint
   'Printer.Print "合計：共 " & Trim(str(iTatle)) & " 筆"
   'iPrint = iPrint + 300
   'Printer.CurrentX = 500
   'Printer.CurrentY = iPrint
   'Printer.Print String(200, "-")
   'end 2007/5/31
   Printer.EndDoc
   ShowPrintOk
   CheckOC
End Sub

Sub StrPrintDoc2()
   adoRecordset.MoveFirst
   GetPrintLeft2
   iLine = 1
   Page = 1
   If Len(txt1(13)) = 0 And Len(txt1(14)) = 0 Then
       If Len(CheckStr(adoRecordset.Fields(13))) <> 0 Then
         TmpArea = IIf(GetPrjName2(CheckStr(adoRecordset.Fields(13))) = "", CheckStr(adoRecordset.Fields(13)), GetPrjName2(CheckStr(adoRecordset.Fields(13))))
       Else
         If Len(CheckStr(adoRecordset.Fields(14))) <> 0 Then
            TmpArea = IIf(GetPrjName2(CheckStr(adoRecordset.Fields(14))) = "", CheckStr(adoRecordset.Fields(14)), GetPrjName2(CheckStr(adoRecordset.Fields(14))))
         Else
            TmpArea = ""
         End If
      End If
   Else
       If Len(CheckStr(adoRecordset.Fields(8))) <> 0 Then
         TmpArea = IIf(GetPrjPeople1(CheckStr(adoRecordset.Fields(8))) = "", CheckStr(adoRecordset.Fields(8)), GetPrjPeople1(CheckStr(adoRecordset.Fields(8))))
       Else
         If Len(CheckStr(adoRecordset.Fields(9))) <> 0 Then
            TmpArea = IIf(GetPrjPeople1(CheckStr(adoRecordset.Fields(9))) = "", CheckStr(adoRecordset.Fields(9)), GetPrjPeople1(CheckStr(adoRecordset.Fields(9))))
         Else
            If Len(CheckStr(adoRecordset.Fields(10))) <> 0 Then
               TmpArea = IIf(GetPrjPeople1(CheckStr(adoRecordset.Fields(10))) = "", CheckStr(adoRecordset.Fields(10)), GetPrjPeople1(CheckStr(adoRecordset.Fields(10))))
            Else
               If Len(CheckStr(adoRecordset.Fields(11))) <> 0 Then
                  TmpArea = IIf(GetPrjPeople1(CheckStr(adoRecordset.Fields(11))) = "", CheckStr(adoRecordset.Fields(11)), GetPrjPeople1(CheckStr(adoRecordset.Fields(11))))
               Else
                  If Len(CheckStr(adoRecordset.Fields(12))) <> 0 Then
                     TmpArea = IIf(GetPrjPeople1(CheckStr(adoRecordset.Fields(12))) = "", CheckStr(adoRecordset.Fields(12)), GetPrjPeople1(CheckStr(adoRecordset.Fields(12))))
                  Else
                     TmpArea = ""
                  End If
               End If
            End If
         End If
      End If
   End If
   StrPrintTital2 TmpArea, str(Page)
   iPrint = 2700
   iTatle = 0       ' 總筆數
   iK = 0           ' 小計
   
   With adoRecordset
       .MoveFirst
       Do While .EOF = False
           For j = 0 To 14
               If Not IsNull(.Fields(j)) Then
                   strTemp3(j) = .Fields(j)
               Else
                   strTemp3(j) = ""
               End If
           Next j
            If Len(txt1(13)) = 0 And Len(txt1(14)) = 0 Then
               If Len(strTemp3(8)) <> 0 Then
                   
               Else
                   If Len(strTemp3(9)) <> 0 Then
                       strTemp3(8) = strTemp3(9)
                   Else
                       If Len(strTemp3(10)) <> 0 Then
                           strTemp3(8) = strTemp3(10)
                       Else
                           If Len(strTemp3(11)) <> 0 Then
                               strTemp3(8) = strTemp3(11)
                           Else
                               If Len(strTemp3(12)) <> 0 Then
                                   strTemp3(8) = strTemp3(12)
                               Else
                                   strTemp3(8) = ""
                               End If
                           End If
                       End If
                   End If
               End If
               strTemp3(2) = GetPrjPeople1(strTemp3(8))
            Else
               If Len(strTemp3(13)) <> 0 Then
                   strTemp3(2) = strTemp3(13)
               Else
                   If Len(strTemp3(14)) <> 0 Then
                      strTemp3(2) = strTemp3(14)
                   Else
                      strTemp3(2) = ""
                   End If
               End If
               strTemp3(2) = GetPrjName1(strTemp3(2))
            End If
           St = strTemp3(0)
           iK = iK + 1
           iTatle = iTatle + 1
           strTemp3(4) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(4)))
           strTemp3(6) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(6)))
           strTemp3(7) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(7)))
           Printer.CurrentX = PLeft(0)
           Printer.CurrentY = iPrint
           'Modify by Morgan 2005/4/27
           'Printer.Print strTemp3(0)
           Printer.Print "" & .Fields("MK") & strTemp3(0)
           Printer.CurrentX = PLeft(1)
           Printer.CurrentY = iPrint
           Printer.Print StrToStr(strTemp3(1), 18)
           Printer.CurrentX = PLeft(2)
           Printer.CurrentY = iPrint
           Printer.Print StrToStr(strTemp3(2), 12)
           Printer.CurrentX = PLeft(3)
           Printer.CurrentY = iPrint
           Printer.Print StrToStr(strTemp3(3), 8)
           Printer.CurrentX = PLeft(4)
           Printer.CurrentY = iPrint
           Printer.Print strTemp3(4)
           Printer.CurrentX = PLeft(5)
           Printer.CurrentY = iPrint
           Printer.Print StrToStr(strTemp3(5), 4)
           Printer.CurrentX = PLeft(6)
           Printer.CurrentY = iPrint
           Printer.Print strTemp3(6)
           Printer.CurrentX = PLeft(7)
           Printer.CurrentY = iPrint
           Printer.Print strTemp3(7)
           .MoveNext
           If .EOF = False Then
               If Not IsNull(.Fields(0)) Then
                   StrTest1 = .Fields(0)
               Else
                   StrTest1 = ""
               End If
               If StrTest1 <> St Then
                   iPrint = iPrint + 300
                   Printer.CurrentX = 500
                   Printer.CurrentY = iPrint
                   Printer.Print String(200, "-")
                   iPrint = iPrint + 300
                   Printer.CurrentX = 1000
                   Printer.CurrentY = iPrint
                   Printer.Print "小計： " & Trim(str(iK)) & " 筆"
                   iK = 0
                   iPrint = iPrint + 300
                   Printer.CurrentX = 500
                   Printer.CurrentY = iPrint
                   Printer.Print String(200, "-")
                   iLine = iLine + 3
               End If
               If (iLine >= 23) Then
                   Printer.NewPage
                   Page = Page + 1
                   StrPrintTital2 TmpArea, str(Page)
                   iPrint = 2400
                   iLine = 0
               End If
               iLine = iLine + 1
               iPrint = iPrint + 300
           End If
       Loop
   End With
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = 1000
   Printer.CurrentY = iPrint
   Printer.Print "小計： " & Trim(str(iK)) & " 筆"
   '合計
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "合計：共 " & Trim(str(iTatle)) & " 筆"
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   
   Printer.EndDoc
   CheckOC

End Sub

Sub StrPrintTital(ByRef Area As String, ByRef Page As String)
   GetPrintLeft
   k = 500
   Printer.Orientation = 2
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 5000
   Printer.CurrentY = 100
   Printer.Print "FCP 收文未發文明細表"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   
   'Add by Morgan 2007/6/1
   If m_Grp <> "" Then
      Printer.CurrentX = 6000
      Printer.CurrentY = k + 200
      '2010/1/8 MODIFY BY SONIA
      'Printer.Print "組別：" & PUB_GetFCPGrpName(m_Grp)
      Printer.Print "組別：" & PUB_GetFCPGrpName(m_Grp, True)
   End If
   'end 2007/6/1
   Printer.CurrentX = 6000
   Printer.CurrentY = k + 500
   Printer.Print "天數：" & txt1(5) & "-" & txt1(6)
   Printer.Font.Bold = False
   Printer.CurrentX = 500
   Printer.CurrentY = k + 800
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 13000
   Printer.CurrentY = k + 800
   Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   'Add by Morgan 2005/4/20 加已延期符號說明
   Printer.CurrentX = 6000
   Printer.CurrentY = k + 1100
   Printer.Print "* 為已延期"
   '2005/4/20 end
   Printer.CurrentX = 13000
   Printer.CurrentY = k + 1100
   Printer.Print "頁    次：" & Page
   Printer.CurrentX = 500
   Printer.CurrentY = k + 1400
   Printer.Print String(200, "-")
   Printer.Font.Underline = True
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = k + 1700
   If txt1(17) = "1" Then
       Printer.Print "承辦人"
   Else
       Printer.Print "智權人員"
   End If
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = k + 1700
   Printer.Print "收文日"
   'Add By Cheng 2003/03/04
   '加申請日標題
   Printer.CurrentX = PLeft(10)
   Printer.CurrentY = k + 1700
   Printer.Print "申請日"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = k + 1700
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = k + 1700
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = k + 1700
   Printer.Print "本所期限"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = k + 1700
   Printer.Print "法定期限"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = k + 1700
   Printer.Print "種類"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = k + 1700
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = k + 1700
   If txt1(17) = "1" Then
       Printer.Print "智權人員"
   Else
       Printer.Print "承辦人"
   End If
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = k + 1700
   Printer.Print "完稿日"
   Printer.Font.Underline = False
   Printer.CurrentX = 500
   Printer.CurrentY = k + 2000
   Printer.Print String(200, "-")
End Sub

Sub StrPrintTital2(ByRef Area As String, ByRef Page As String)
   GetPrintLeft2
   k = 500
   Printer.Orientation = 2
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 5000
   Printer.CurrentY = 100
   Printer.Print "FCP 收文未發文明細表"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 6000
   Printer.CurrentY = k + 500
   Printer.Print "天數：" & txt1(5) & "-" & txt1(6)
   Printer.Font.Bold = False
   Printer.CurrentX = 500
   Printer.CurrentY = k + 800
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 13000
   Printer.CurrentY = k + 800
   Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   Printer.CurrentX = 500
   Printer.CurrentY = k + 1100
   If Len(txt1(13)) = 0 And Len(txt1(14)) = 0 Then
       Printer.Print "代理人：" & Area
   Else
       Printer.Print "申請人：" & Area
   End If
   'Add by Morgan 2005/4/20 加已延期符號說明
   Printer.CurrentX = 6000
   Printer.CurrentY = k + 1100
   Printer.Print "* 為已延期"
   '2005/4/20 end
   Printer.CurrentX = 13000
   Printer.CurrentY = k + 1100
   Printer.Print "頁    次：" & Page
   Printer.CurrentX = 500
   Printer.CurrentY = k + 1400
   Printer.Print String(200, "-")
   Printer.Font.Underline = True
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = k + 1700
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = k + 1700
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = k + 1700
   If Len(txt1(13)) = 0 And Len(txt1(14)) = 0 Then
       Printer.Print "申請人"
   Else
       Printer.Print "代理人"
   End If
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = k + 1700
   Printer.Print "申請案號"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = k + 1700
   Printer.Print "收文日"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = k + 1700
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = k + 1700
   Printer.Print "本所期限"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = k + 1700
   Printer.Print "法定期限"
   Printer.Font.Underline = False
   Printer.CurrentX = 500
   Printer.CurrentY = k + 2000
   Printer.Print String(200, "-")
End Sub

Sub GetPrintLeft()
   Erase PLeft
   PLeft(0) = 500
   PLeft(1) = 1550
   'Add By Cheng 2003/03/04
   PLeft(10) = 2750 '申請日
   PLeft(2) = 2850 + 1100
   PLeft(3) = 4850 + 1100
   PLeft(4) = 7950 + 1100
   'Modify by Morgan 2004/12/27 本所期限與法定期限間距加寬
   'PLeft(5) = 8700 + 1100
   'PLeft(6) = 9700 + 1100
   PLeft(5) = 9150 + 1100
   PLeft(6) = 10300 + 1100
   '2004/12/27 end
   PLeft(7) = 11100 + 1100
   PLeft(8) = 12200 + 1100
   PLeft(9) = 13300 + 1100
End Sub

Sub GetPrintLeft2()
   Erase PLeft
   PLeft(0) = 500
   PLeft(1) = 2500
   PLeft(2) = 7000
   PLeft(3) = 10000
   PLeft(4) = 11600
   PLeft(5) = 12800
   PLeft(6) = 14000
   PLeft(7) = 15200
End Sub

Private Sub Form_Load()
Dim strST05 As String
   
   MoveFormToCenter Me
   '900803  邱小姐說預設為 FCP,FG 因為報表格式一樣
   txt1(0) = "FCP,FG"
   'Add By Cheng 2003/02/19
   '系統類別再加"CFP","CPS","P","PS"
   txt1(0) = txt1(0) & ",CFP,CPS,P,PS"
   'txt1(0).Enabled = False
   
   'Add by Morgan 2010/3/23 收文業務區預設F23
   txt1(1) = "F23"
   txt1(2) = "F23"

   'Add by Morgan 2007/9/21
   '外專工程師要控管權限
   strST05 = PUB_GetST05(strUserNum)
   Select Case strST05
      Case "39" '外專工程師中級主管只可查該組
         txt1(18) = PUB_GetStaffST16(strUserNum)
         txt1(18).Locked = True
      Case "40", "49" '外專工程師只可查本人  'modify by sonia 2024/8/15 加入等級49日外專海外工程師
         txt1(4) = strUserNum
         LBL1(1) = strUserName
         txt1(4).Locked = True
   End Select

   '輸出
   If Pub_StrUserSt15 = "F21" Then
      txt2 = "1"
      'Modified by Morgan 2012/5/15
      'txt1(14).Enabled = False
      '各組主管可列印
      'Removed by Morgan 2012/6/1 改用權限控管
      'If strST05 <> "38" And strST05 <> "42" Then
      '   txt2.Enabled = False
      'End If
      'end 2012/6/1
      'end 2012/5/15
   Else
      txt2 = "2"
   End If
   
   'Added by Morgan 2012/6/1 改用權限控管
   If IsUserHasRightOfFunction(Me.Name, strPrint, False) = False Then
      txt2 = "1"
      txt2.Enabled = False
   End If
   'end 2012/6/1
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060309 = Nothing
End Sub


Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2002/09/16
   Select Case Index
   Case 17
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   '2008/2/22 ADD BY SONIA 加組別條件
   Case 18
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   '2008/2/22 END
   End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Select Case Index
      Case 0 '系統類別
           strTemp1 = Split(UCase(GetSystemKindByNick), ",")
           strTemp2 = Split(UCase(txt1(0)), ",")
           For i = 0 To UBound(strTemp2)
              s = 0
              For j = 0 To UBound(strTemp1)
                  If strTemp1(j) = strTemp2(i) Then
                      s = 1
                      Exit For
                  End If
              Next j
              If s = 0 And strTemp2(i) <> "CFP" And strTemp2(i) <> "CPS" And strTemp2(i) <> "P" And strTemp2(i) <> "PS" Then
                  s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
                  txt1(0).SetFocus
                  txt1(0).SelStart = 0
                  txt1(0).SelLength = Len(txt1(0))
                  Exit Sub
              End If
              'If UCase(StrTemp2(I)) <> "FCP" Then
              '     S = MsgBox("此功能只能查詢  FCP 的報表，若要查詢 FCP 以外的文件，請從其他系統進入", , "報表格式不同")
              '    TXT1(0).SetFocus
              '    TXT1(0).SelStart = 0
              '    TXT1(0).SelLength = Len(TXT1(0))
              '    Exit Sub
              ' End If
          Next i
      Case 3
            'Modify By Cheng 2002/09/27
      '     lbl1(0) = GetPrjSales(txt1(3))
           LBL1(0) = GetPrjSales(txt1(3), "智權人員")
            'Add By Cheng 2002/09/26
            If Me.txt1(3).Text <> "" Then
               If Me.txt1(3).Text = Me.LBL1(0).Caption Then
                  Me.LBL1(0).Caption = ""
                  Me.txt1(3).SetFocus
                  txt1_GotFocus 3
                  Exit Sub
               End If
            End If
      Case 4
           LBL1(1) = GetPrjSales(txt1(4))
            'Add By Cheng 2002/09/26
            If Me.txt1(4).Text <> "" Then
               If Me.txt1(4).Text = Me.LBL1(1).Caption Then
                  Me.LBL1(1).Caption = ""
                  Me.txt1(4).SetFocus
                  txt1_GotFocus 4
                  Exit Sub
               End If
            End If
      Case 2, 6, 8
         'Modify By Cheng 2002/09/16
         If blnClkSure = False Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               txt1(Index - 1).SetFocus
               txt1_GotFocus (Index - 1)
               Exit Sub
            End If
         Else
            blnClkSure = False
         End If
      Case 14
         'Modify By Cheng 2002/09/16
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
      Case 16
         'Modify By Cheng 2002/09/16
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
      Case 11
           Select Case txt1(11)
           Case "y", "Y", " "
           Case Else
               s = MsgBox("是否列印明細只能輸入 Y 或空白 !!", , "USER 輸入錯誤")
               txt1(11).SetFocus
               txt1(11).SelStart = 0
               txt1(11).SelLength = Len(txt1(11))
               Exit Sub
           End Select
      Case 12
           Select Case txt1(11)
           Case "y", "Y", " "
           Case Else
               s = MsgBox("是否計算多國家只能輸入 Y 或空白 !!", , "USER 輸入錯誤")
               txt1(12).SetFocus
               txt1(12).SelStart = 0
               txt1(12).SelLength = Len(txt1(12))
               Exit Sub
           End Select
      Case 17
           Select Case Val(txt1(17))
           Case 1, 2
           Case Else
              s = MsgBox("列印別只能輸入 1 或 2 !!", , "USER 輸入錯誤")
              txt1(17).SetFocus
              txt1(17).SelStart = 0
              txt1(17).SelLength = Len(txt1(17))
              Exit Sub
           End Select
      '2008/2/22 ADD BY SONIA 加組別條件
      Case 18
           Select Case txt1(18)
           Case "1", "2", "3", "4", ""
           Case Else
              s = MsgBox("組別只能輸入 1, 2, 3 或 4 !!", , "USER 輸入錯誤")
              txt1(18).SetFocus
              txt1(18).SelStart = 0
              txt1(18).SelLength = Len(txt1(18))
              Exit Sub
           End Select
      '2008/2/22 END
      Case Else
   End Select
End Sub

Private Sub txt2_GotFocus()
   TextInverse txt2
   CloseIme
End Sub

Private Sub txt2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
   End If
End Sub

Private Sub SetGrid(p_Rst As ADODB.Recordset, Optional p_iMode As Integer = 1)
   Dim iCols As Integer
   With frm060309_1
      .Show
      .grdDataList.Visible = False
      Set .grdDataList.Recordset = p_Rst.Clone
      Select Case p_iMode
         Case 1  '未輸入申請人及代理人條件: 排列順序 1.承辦人
            .grdDataList.FormatString = "承辦人|收文日　|申請日　|本所案號　　　　|案件名稱　　　　|本所期限|法定期限|種類|案件性質|智權人員|完稿日　|員工代碼|組別　　"
            iCols = 11
         Case 2  '未輸入申請人及代理人條件: 排列順序 2.智權人員
            .grdDataList.FormatString = "智權人員|收文日　|申請日　|本所案號　　　　|案件名稱　　　　|本所期限|法定期限|種類|案件性質|承辦人|完稿日　|員工代碼"
            iCols = 11
         Case 3 '輸入申請人及代理人條件
            .grdDataList.FormatString = "本所案號　　　　|案件名稱　　　　|申請人　　　　　|申請案號|收文日　|案件性質|本所期限|法定期限"
            iCols = 8
      End Select
      Select Case p_iMode
         Case 1, 2
            For intI = 0 To .grdDataList.Cols - 1
               Select Case intI
                  '日期置中
                  Case 3
                     .grdDataList.ColAlignment(intI) = 4
                  '其他靠左
                  Case Else
                     .grdDataList.ColAlignment(intI) = 1
               End Select
               'Added by Morgan 2012/12/24
               If intI > iCols - 1 Then
                  .grdDataList.ColWidth(intI) = 0
               End If
               
            Next
         Case 3
            'Added by Morgan 2012/12/24
            For intI = 0 To .grdDataList.Cols - 1
               If intI > iCols - 1 Then
                  .grdDataList.ColWidth(intI) = 0
               End If
            Next
            
      End Select
      .grdDataList.Visible = True
   End With
End Sub
