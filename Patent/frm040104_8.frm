VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040104_8 
   BorderStyle     =   1  '單線固定
   Caption         =   "內專發文-設定質權 / 終止設定質權"
   ClientHeight    =   4920
   ClientLeft      =   -372
   ClientTop       =   2388
   ClientWidth     =   8604
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   8604
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   4950
      MaxLength       =   4
      TabIndex        =   9
      Top             =   3000
      Width           =   540
   End
   Begin VB.TextBox txtChkRltDate 
      Height          =   270
      Left            =   1095
      MaxLength       =   8
      TabIndex        =   13
      Top             =   4620
      Width           =   975
   End
   Begin VB.TextBox txtCP84 
      Height          =   285
      Left            =   4410
      TabIndex        =   5
      Top             =   2676
      Width           =   1092
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   1320
      TabIndex        =   10
      Top             =   3330
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   5688
      MaxLength       =   1
      TabIndex        =   4
      Top             =   2676
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   1
      Left            =   2580
      MaxLength       =   8
      TabIndex        =   8
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7272
      TabIndex        =   15
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6444
      TabIndex        =   14
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   0
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   3
      Top             =   2676
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   19
      Top             =   744
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   18
      Top             =   744
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   17
      Top             =   744
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   960
      MaxLength       =   3
      TabIndex        =   16
      Top             =   744
      Width           =   495
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   1104
      Left            =   6984
      TabIndex        =   6
      Top             =   2676
      Width           =   1452
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2561;1947"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text15 
      Height          =   300
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   1845
      Width           =   7110
      VariousPropertyBits=   671107099
      MaxLength       =   160
      Size            =   "12541;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text9 
      Height          =   735
      Left            =   3960
      TabIndex        =   12
      Top             =   3855
      Width           =   4470
      VariousPropertyBits=   -1467987941
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "7885;1296"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text8 
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   3735
      VariousPropertyBits=   -1467987941
      MaxLength       =   60
      ScrollBars      =   2
      Size            =   "6588;1296"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text15 
      Height          =   300
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   2085
      Width           =   7110
      VariousPropertyBits=   671107099
      MaxLength       =   250
      Size            =   "12541;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text15 
      Height          =   300
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   2325
      Width           =   7110
      VariousPropertyBits=   671107099
      MaxLength       =   160
      Size            =   "12541;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCP113 
      AutoSize        =   -1  'True
      Caption         =   "工作時數:"
      Height          =   180
      Index           =   18
      Left            =   4110
      TabIndex        =   54
      Top             =   3045
      Width           =   765
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "催審期限:"
      Height          =   180
      Left            =   135
      TabIndex        =   52
      Top             =   4635
      Width           =   765
   End
   Begin VB.Label lblCaseFee 
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   2085
      TabIndex        =   51
      Tag             =   "Y"
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人:"
      Height          =   180
      Left            =   6000
      TabIndex        =   50
      Top             =   2712
      Width           =   948
   End
   Begin VB.Label lblCP84 
      AutoSize        =   -1  'True
      Caption         =   "發文規費:"
      Height          =   180
      Left            =   3510
      TabIndex        =   49
      Top             =   2715
      Width           =   765
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(日):"
      Height          =   180
      Left            =   120
      TabIndex        =   48
      Top             =   2355
      Width           =   1065
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(英):"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   47
      Top             =   2115
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(中):"
      Height          =   180
      Left            =   120
      TabIndex        =   46
      Top             =   1875
      Width           =   1065
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   120
      X2              =   8460
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   8460
      Y1              =   1776
      Y2              =   1776
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家"
      Height          =   180
      Index           =   1
      Left            =   3510
      TabIndex        =   45
      Top             =   540
      Width           =   720
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   12
      Left            =   4395
      TabIndex        =   44
      Top             =   540
      Width           =   1650
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2910;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   14
      Left            =   2880
      TabIndex        =   43
      Top             =   3390
      Width           =   2160
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3810;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "代理人:"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   42
      Top             =   3330
      Width           =   585
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   10
      Left            =   4395
      TabIndex        =   41
      Top             =   1515
      Width           =   1650
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2910;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   9
      Left            =   6975
      TabIndex        =   40
      Top             =   1530
      Width           =   1440
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2540;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   8
      Left            =   960
      TabIndex        =   39
      Top             =   1515
      Width           =   2460
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "4339;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   6
      Left            =   4395
      TabIndex        =   38
      Top             =   1290
      Width           =   1650
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2910;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   5
      Left            =   960
      TabIndex        =   37
      Top             =   1290
      Width           =   2460
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "4339;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   3
      Left            =   960
      TabIndex        =   36
      Top             =   1080
      Width           =   2460
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "4339;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   1
      Left            =   4395
      TabIndex        =   35
      Top             =   1080
      Width           =   1650
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2910;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   0
      Left            =   960
      TabIndex        =   34
      Top             =   540
      Width           =   2460
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "4339;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Left            =   3960
      TabIndex        =   33
      Top             =   3615
      Width           =   765
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "關係人:"
      Height          =   180
      Left            =   120
      TabIndex        =   32
      Top             =   3615
      Width           =   585
   End
   Begin VB.Label Label26 
      Caption         =   "~"
      Height          =   255
      Left            =   2385
      TabIndex        =   31
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "質權設定期間:"
      Height          =   180
      Left            =   120
      TabIndex        =   30
      Top             =   3000
      Width           =   1125
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Left            =   120
      TabIndex        =   29
      Top             =   2715
      Width           =   585
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "業務區:"
      Height          =   180
      Left            =   3510
      TabIndex        =   28
      Top             =   1515
      Width           =   585
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Index           =   0
      Left            =   6120
      TabIndex        =   27
      Top             =   1515
      Width           =   765
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "申請人:"
      Height          =   180
      Left            =   120
      TabIndex        =   26
      Top             =   1515
      Width           =   585
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   3510
      TabIndex        =   25
      Top             =   1080
      Width           =   585
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "專利種類:"
      Height          =   180
      Left            =   3510
      TabIndex        =   24
      Top             =   1290
      Width           =   765
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   120
      TabIndex        =   23
      Top             =   1290
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   21
      Top             =   750
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   540
      Width           =   585
   End
   Begin VB.Label lblCaseFees 
      BackColor       =   &H80000010&
      Height          =   255
      Left            =   2130
      TabIndex        =   53
      Top             =   4620
      Width           =   255
   End
End
Attribute VB_Name = "frm040104_8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/14 改成Form2.0 (Text15,Text8,Text9,lstNameAgent,Label2)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
'整理 by Morgan 2005/7/15
Option Explicit

Dim strReceiveNo As String

'Modify by Morgan 2005/7/15 改用動態陣列
'Dim pa(1 To T_PA) As String, cp(T_CP) As String
Dim pa() As String, cp() As String

Dim intWhere As Integer
Dim m_CP09s As String, m_CP123s As String 'Add by Morgan 2009/3/23 收文號,是否算發文室案件
Dim m_CP130 As String 'Add by Morgan 2009/4/28 發文-主管機關
Dim m_bolFMP As Boolean 'Added by Lydia 2023/06/20 是否為FMP案
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/06/20 是否為寰華

Private Function Process(Index As Integer) As Boolean
   Dim strTmp As String
   '檢查輸入資料的完整性
   If CheckDataIntegrity = False Then Exit Function
   'Add By Cheng 2002/05/22
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Function
   
   'Add by Morgan 2009/3/23 設定是否算發文室案件
   If pa(9) = 台灣國家代號 Then
      'Add by Morgan 2009/4/28
      If ModifyDispatchCp130(cp(9), m_CP09s, m_CP123s, m_CP130, Text7) = False Then
         Exit Function
      End If
      If m_CP123s = "Y" Then
      'end 2009/4/28
         'modify by sonia 2014/6/23 加傳發文規費, P-108903
         If ModifyDispatch(cp(9), m_CP09s, m_CP123s, txtCP84, Text7) = False Then
             Exit Function
         End If
      End If
      'Add by Amy 2014/10/14 P台灣案發文控制
      If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
        If pa(1) = "P" And cp(9) < "C" Then
            If cp(9) < "B" Then
                'A類一定要有接洽單才可發文
                'Modify by Amy 2014/11/27 取消ChkOneDayHasCP27判斷,接洽單改檢查,因考慮可能同時發文其他案件性質情形
                'If PUB_CheckPDF2(cp(9), 0, True, strExc(0)) = False And ChkOneDayHasCP27(pa(1), pa(2), pa(3), pa(4), cp(5) + 19110000) = False Then
                If PUB_CheckPDF3(Text1, Text2, Text3, Text4) = False Then
                    Exit Function
                End If
            End If
            'AB類申請書確認檢查,符合條件才可發文
            'Modified by Morgan 2015/3/17
            'If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" And PUB_CheckPDF2(cp(9), 1, True, strExc(0)) = False Then
            If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" Then
               If PUB_CheckPDF2(cp(9), 1, True, strExc(0)) = False Then
            'end 2015/3/17
                  MsgBox "無申請書PDF檔 ,不可發文!", vbInformation
                  Exit Function
               End If 'Added by Morgan 2015/3/17
            End If
        End If
      End If
      'end 2014/10/14
      
   'Added by Morgan 2016/6/29 非臺灣案電子化
   ElseIf 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
      If cp(9) < "B" And Left(cp(12), 1) <> "F" Then
          If PUB_CheckPDF3(Text1, Text2, Text3, Text4) = False Then
              Exit Function
          End If
      End If
   'end 2016/6/29
   End If
      
   If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical:  Exit Function
   
   Process = True
   
   'Add by Morgan 2008/2/20 檢查代理人Email(需考慮可能為FF案件)
   PUB_CheckEMail Combo2
   PUB_CheckEMail pa(75), pa(144)
   If pa(145) <> "" Then
      PUB_CheckEMail pa(75), pa(145)
   End If
   'end 2008/2/20
      
   '2012/7/23 add by sonia
   '台灣案發文規費與收文規費不符時,mail給智權人員
   If txtCP84.Enabled = True And pa(9) = "000" And Val(Me.txtCP84.Text) <> Val(cp(17)) Then
      '2013/7/2 modify by sonia 改用共用module
      PUB_ChkOfficialFee cp(9), Me.txtCP84.Text
   End If
   '2012/7/23 end
   
   'Add by Morgan 2007/6/14
   If pa(9) = "000" Then
      PUB_ReAsignInform pa(1), pa(2), pa(3), pa(4), strReceiveNo
   End If
      
   If pa(9) = 台灣國家代號 Then '通知函
      strTmp = "00"
   ElseIf pa(9) <> 台灣國家代號 Then
      strTmp = "01"
   End If
   EndLetter "02", strReceiveNo, strTmp, strUserNum
   'Modify by Amy 2014/08/22 台灣案電子化 +傳strLetterRecNo
   NowPrint strReceiveNo, "02", strTmp, False, strUserNum, 0, , , , , , , , , , , , strReceiveNo
   
   'Added by Lydia 2024/03/06 外專機械設計組人員異動調整程式：內專協辦工程師完成送件之後，需通知外專工程師進行請款
   'Move by Lydia 2024/03/12 改使用Outlook草稿，從FormSave移出
   'Mark by Lydia 2024/04/09 FMP案不用通知--- Phoebe
   'If m_bolFMP = True And cp(1) = "P" And Mid(cp(14), 4, 1) = "9" Then
   '   Call Pub_SetEngMail(cp(9))
   'End If
   ''end 2024/03/06
   'end 2024/04/09
   
   '若有未發文資料顯示警告
   PUB_GetCPunIssueDatas "" & Me.Text1.Text & "-" & Me.Text2.Text & "-" & IIf(Len("" & Me.Text3.Text) <= 0, "0", Me.Text3.Text) & "-" & IIf(Len("" & Me.Text4.Text) <= 0, "00", Me.Text4.Text)
End Function

Private Sub cmdok_Click(Index As Integer)
   ' 設定滑鼠游標為等待狀態
   Screen.MousePointer = vbHourglass
   If Index = 0 Then
      'Modify by Morgan 2010/2/10 改呼叫函數方式以便鎖定按鍵
      cmdOK(Index).Enabled = False
      If Not Process(Index) Then
         cmdOK(Index).Enabled = True
      Else
         'Add By Sindy 2013/5/20
         If frm040104_1.bolIsEMPFlow = True Then
            Unload frm040104_1
            frm090202_4.Show
            frm090202_4.QueryData
         Else
         '2013/5/20 End
            frm040104_1.Show
            frm040104_1.Clear
         End If
         Unload Me
      End If
   Else
      'Add By Sindy 2013/5/20
      If frm040104_1.bolIsEMPFlow = True Then
         Unload frm040104_1
         frm090202_4.Show
         frm090202_4.QueryData
      Else
      '2013/5/20 End
         frm040104_1.Show
         frm040104_1.Clear
      End If
      Unload Me
   End If
   ' 設定滑鼠游標為預設
   Screen.MousePointer = vbDefault
End Sub

Private Function FormSave() As Boolean
 Dim i As Integer, strTxt(1 To 20) As String, ii As Integer
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
FormSave = True
cnnConnection.BeginTrans


   'Added by Morgan 2013/6/7 自 lstNameAgent_Validate 移來,否則若觸發 Form_Activate 事件會跑 ReadPatent 導致 cp(110) 被清除
   cp(110) = ""
   If lstNameAgent.Visible = True Then
      For ii = 0 To lstNameAgent.ListCount - 1
         If lstNameAgent.Selected(ii) = True Then
            'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
            'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
            cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         End If
      Next
      If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
   End If
   'end 2013/6/7

   If Combo2 <> "" Then
      'Modify by Morgan 2008/2/22
      'cp(44) = ChangeCustomerL(Combo2)
      intI = InStr(Combo2, "-")
      If intI > 0 Then
         cp(44) = Left(Combo2, intI - 1)
         cp(116) = Mid(Combo2, intI + 1)
      Else
         cp(44) = Combo2
         cp(116) = ""
      End If
      cp(44) = ChangeCustomerL(cp(44))
      'end 2008/2/22
      'edit by nickc 2007/02/02 不用 dll 了
      'If Not objPublicData.GetCaseThatCode(cp) Then cp(45) = ""
      If Not ClsPDGetCaseThatCode(cp) Then cp(45) = ""
   Else
      cp(44) = ""
      cp(116) = ""
      cp(45) = ""
   End If
   
   ' 91.03.25 modify by louis (單引號)
   'Modify by morgan 2004/8/11 加 cp84
   'Modify by Morgan 2005/7/15 加 cp110
   'Modified by Lydia 2021/05/25 +CP113工作時數
   'Modified by Lydia 2023/06/20 +CP14
   strTxt(1) = "UPDATE CASEPROGRESS SET CP27=" & TransDate(Text7, 2) & ",CP14=" & CNULL(cp(14)) & _
      ",CP44=" & CNULL(cp(44)) & ",CP116=" & CNULL(cp(116)) & ",CP45=" & CNULL(ChgSQL(cp(45))) & _
      ",cp53=" & TransDate(Text5(0), 2) & ",cp54=" & TransDate(Text5(1), 2) & _
      ",cp50=" & CNULL(Text8) & ",cp64=" & CNULL(ChgSQL(Text9)) & ", cp84=" & Format(Val(txtCP84.Text)) & _
      ",cp110=" & CNULL(cp(110)) & " ,cp113=" & CNULL(txtCP113, True) & _
      " where cp09='" & strReceiveNo & "'"
    'Add By Cheng 2002/11/06
    cnnConnection.Execute strTxt(1)
      
   strTxt(2) = "UPDATE PATENT SET PA05=" & CNULL(Text15(0)) & ",PA06=" & CNULL(ChgSQL(Text15(1))) & _
      ",PA07=" & CNULL(Text15(2)) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
    'Add By Cheng 2002/11/06
    cnnConnection.Execute strTxt(2)
   
   i = 2
   
   strExc(0) = "SELECT CF23 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & cp(10) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp.Fields(0)) Or RsTemp.Fields(0) <> 0 Then
         i = 3
        '若本所期限非工作天則抓最近的工作天
         'edit by nickc 2007/02/02 不用 dll 了
         'strTxt(i) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
            "NP09,NP10,NP22) VALUES ('" & strReceiveNo & "','" & pA(1) & "','" & pA(2) & _
            "','" & pA(3) & "','" & pA(4) & "'," & 收達 & "," & _
            PUB_GetWorkDay1(CompDate(2, rsTemp.Fields(0), TransDate(Text7, 2)), True) & "," & _
            CompDate(2, rsTemp.Fields(0), TransDate(Text7, 2)) & ",'" & _
            strUserNum & "'," & objPublicData.GetNextProgressNo & ")"
         strTxt(i) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
            "NP09,NP10,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & _
            "','" & pa(3) & "','" & pa(4) & "'," & 收達 & "," & _
            PUB_GetWorkDay1(CompDate(2, RsTemp.Fields(0), TransDate(Text7, 2)), True) & "," & _
            CompDate(2, RsTemp.Fields(0), TransDate(Text7, 2)) & ",'" & _
            strUserNum & "'," & GetNextProgressNo & ")"
        cnnConnection.Execute strTxt(i)
      End If
   End If
   
   'Add by Morgan 2009/3/23
   If pa(9) = 台灣國家代號 Then
      PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130
      'Modify by Amy 2014/09/09 for 台灣案電子化
      If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
        If cp(9) < "C" Then
            'Modified by Morgan 2018/8/1
            'strExc(1) = PUB_GetLetterJudge(pa(1), cp(10), , , pa(1), pa(2), pa(3), pa(4))
            strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), cp(10))
            'Modify by Amy 2015/02/13 此固定出客戶通知函,故不考慮未出客戶函 修改判斷條件
            'PUB_AddLetterProgress strReceiveNo, 1, True, strExc(1), False, pa(26), cp(10), pa(75), True
              '1.　電子送件有規費的有收據；無規費的無回執
              '2.非電子送件要計件的有回執；不計件的無回執
            'Mark by Amy 2015/03/06 回執改至PUB_UpdateLP19做
'            If cp(118) = "Y" Then
'                If Val(txtCP84) > 0 Then
'                    PUB_AddLetterProgress strReceiveNo, 1, True, strExc(1), False, pa(26), cp(10), pa(75), True
'                Else
'                    PUB_AddLetterProgress strReceiveNo, 0, True, strExc(1), False, pa(26), cp(10), pa(75), False
'                End If
'            Else
                If Left(m_CP123s, 1) = "Y" Then
                    PUB_AddLetterProgress strReceiveNo, 1, True, strExc(1), False, pa(26), cp(10), pa(75), True
                Else
                    PUB_AddLetterProgress strReceiveNo, 0, True, strExc(1), False, pa(26), cp(10), pa(75), False
                End If
'            End If
            'end 2015/03/06
        End If
      End If
      'end 2014/09/09
      'Add by Amy 2015/02/13 更新收據/回執設定
      'Modify by Amy 2015/03/06 +發文日參數
      PUB_UpdateLP19 cp(1), cp(2), cp(3), cp(4), m_CP09s, m_CP123s, Text7
   
   'Added by Morgan 2016/5/26 非臺灣案電子化
   ElseIf Left(Pub_StrUserSt03, 1) <> "F" Then
      '客戶通知函
      If 內專全面電子化啟用日 <= Val(strSrvDate(1)) Then
         'Modified by Morgan 2018/8/1
         'strExc(1) = PUB_GetLetterJudge(pa(1), cp(10), , pa(9), pa(1), pa(2), pa(3), pa(4))
         strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), cp(10), pa(9), , , IIf(Left(cp(12), 1) = "F", True, False))
         PUB_AddLetterProgress strReceiveNo, 0, True, strExc(1), False, pa(26), cp(10), pa(75)
      End If
   'end 2016/5/26
   End If
   
   'Add by Morgan 2009/8/17
   If txtChkRltDate <> "" Then
      PUB_UpdateChkResultDate txtChkRltDate, cp, cp(9), cp(10), cp(43)
   End If
   
   cnnConnection.CommitTrans
   Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    FormSave = False
End Function

Private Sub Combo2_Click()
   Combo2_Validate False
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   Dim strNo As String, iPos As Integer
   If Combo2.Text = "" Then
      If pa(9) <> 台灣國家代號 Then
         MsgBox "當申請國家非台灣時, 代理人欄不可為空白!!!", vbExclamation
         Cancel = True
         Exit Sub
      End If
      
   ElseIf Not ChgType(12) Then
      Cancel = True
      
   Else
      strNo = Combo2.Text
      
      'Add by Morgan 2008/2/22 加聯絡人判斷
      iPos = InStr(strNo, "-")
      If iPos > 0 Then
         strNo = Left(strNo, iPos - 1)
      End If
      'end 2008/2/22
      
      'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
      If PUB_CheckStatus(strNo) = False Then Cancel = True
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   With frm040104_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
   'Add by Morgan 2005/7/15
   ReDim pa(1 To TF_PA)
   ReDim cp(TF_CP)
   
   ReadPatent
   
   cp(110) = "" '要清空,否則若重新發文會殘留前次發文資料,當新案有改出名人而本程序未改選將會造成不一致 Added by Morgan 2012/9/7
   
   'Add by Morgan 2005/7/15
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   Text10.Visible = False
   lstNameAgent.Clear
   If pa(9) = "000" Then
      PUB_SetOurAgent lstNameAgent, pa(), cp(110), , True 'Modified by Morgan 2021/12/15 +傳入bForm2=True
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   '2005/7/14 END
   
   Label2(0) = strReceiveNo
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call PUB_SendMailCache 'Added by Lydia 2024/03/06
   
   'Set frm040104_8 = Nothing 'Removed by Morgan 2021/12/15 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub ReadPatent()
Dim Lbl As Object, txt As TextBox, i As Integer, bolTmp As Boolean
Dim m_Fee As String         '銷帳服務費 2012/8/1 add by sonia
Dim m_Official As String    '銷帳規費   2012/8/1 add by sonia
   
   For Each Lbl In Label2
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
         
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      If pa(9) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNation(pA(9), strExc(0)) Then Label2(12) = strExc(0)
         If ClsPDGetNation(pa(9), strExc(0)) Then Label2(12) = strExc(0)
      End If
      Label2(3) = pa(22)
      ChgType (999) 'pa(8)
      Text15(0) = pa(5)
      Text15(1) = pa(6)
      Text15(2) = pa(7)
      If pa(26) <> "" Then ChgType (11) ' Label2(8)
   End If
   
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      If cp(10) <> "" Then
         If pa(9) = 台灣國家代號 Then
            bolTmp = False
         Else
            bolTmp = True
         End If
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(cp(1), cp(10), strExc(0), BolTmp) Then
         If ClsPDGetCaseProperty(cp(1), cp(10), strExc(0), bolTmp) Then
            Label2(5) = strExc(0)
         End If
      End If
      'Added by Lydia 2023/06/20 判斷FCP案,寰華案
      If Left(cp(12), 1) = "F" And pa(9) <> "000" Then
         m_bolFMP = True
      Else
         m_bolFMP = False
      End If
      m_bolFMP2 = False
      If m_bolFMP = True Then
         m_bolFMP2 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, pa(1), pa(2), pa(3), pa(4))
         '寰華案:承辦人為外專程序時,改為操作人員
         If m_bolFMP2 = True Then
            cp(14) = GetFCPUser(cp(14))
         End If
      End If
      'end 2023/06/20
      If cp(14) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(cp(14), strExc(0)) Then Label2(1) = strExc(0)
         If ClsPDGetStaff(cp(14), strExc(0)) Then Label2(1) = strExc(0)
      End If
      If cp(12) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaffDeptName(cp(12), strExc(0)) Then Label2(10) = strExc(0)
         If ClsPDGetStaffDeptName(cp(12), strExc(0)) Then Label2(10) = strExc(0)
      End If
      If cp(13) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(cp(13), strExc(0)) Then Label2(9) = strExc(0)
         If ClsPDGetStaff(cp(13), strExc(0)) Then Label2(9) = strExc(0)
      End If
      If cp(27) = "" Then
         Text7 = strSrvDate(2)
      Else
         Text7 = cp(27)
      End If
      Text5(0) = cp(53)
      Text5(1) = cp(54)
      Text8 = cp(50)
      Text9 = cp(64)
      Text10 = cp(22)
      'Modify by Morgan 2008/10/16 +若進度檔已有代理人則預設
      'Modified by Lydia 2016/10/27 +新案有申請人指定國外代理人檔則預設 => cp(9), pa(9), pa(26)
      AddAgent Combo2, cp, , cp(44), cp(116), cp(9), pa(9), pa(26)
   End If
   
   '2012/8/1 add by sonia 若有銷帳則要扣除銷帳規費
   If Val(cp(77)) > 0 Then
      If GetCP77Detail(cp(9), m_Fee, m_Official) = True Then
         cp(17) = cp(17) - m_Official
      End If
   End If
   '2012/8/1 end
   
   'Add by Morgan 2004/8/11
   txtCP84.Tag = cp(17)
   
   'Add by Morgan 2009/8/17
   If Text7 <> "" Then
      PUB_SetChkResultDate pa(1), pa(9), cp(10), Text7, txtChkRltDate, cp, pa(8)
      Text7.Tag = Text7
   End If
   
    'Added by Lydia 2021/05/25
    txtCP113 = ""
    If cp(113) <> "" Then txtCP113 = cp(113)
    'end 2021/05/25
End Sub

Private Function ChgType(i As Integer) As Boolean
 Dim strTempName As String, bolTmp As Boolean
   ChgType = False
   Select Case i
      Case 0
         '2011/12/8 MODIFY BY SONIA 發文日可輸系統日的下一個工作日
         'If Not ChkDate(Text7) Or Val(Text7.Text) > Val(strSrvDate(2)) Then
         '   MsgBox "發文日期不正確或發文日大於系統日，請重新輸入 !", vbCritical
         If Not ChkDate(Text7) Or DBDATE(Val(Text7.Text)) > DBDATE(PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)))) Then
            MsgBox "發文日期不正確或發文日大於系統日下一個工作日，請重新輸入 !", vbCritical
         '2011/12/8 END
         Else
            ChgType = True
         End If
      Case 11
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.LawGetName(pa(26), strTempName) Then
         If ClsLawLawGetName(pa(26), strTempName) Then
            Label2(8) = strTempName
            ChgType = True
         End If
      Case 12 '代理人
         strExc(1) = Combo2.Text
         'Add by Morgan 2008/2/22 加判斷是否為聯絡人
         If InStr(strExc(1), "-") > 0 Then
            If ClsPDGetContact(strExc(1), strTempName) Then
               Combo2 = strExc(1)
               Label2(14) = strTempName
               ChgType = True
            End If
         
         '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
         ElseIf PUB_GetAgentName(pa(1), strExc(1), strTempName) = True Then
            Combo2.Text = strExc(1)
            Label2(14).Caption = strTempName
            ChgType = True
            
         Else
            Label2(14).Caption = ""
         End If
         
      Case 999
         If pa(9) = 台灣國家代號 Then
            bolTmp = False
         Else
            bolTmp = True
         End If
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetPatentTrademarkKind(專利, pA(8), strTempName, BolTmp, pA(9)) = 1 Then
         If ClsPDGetPatentTrademarkKind(專利, pa(8), strTempName, bolTmp, pa(9)) = 1 Then
            Label2(6) = strTempName
         End If
   End Select
End Function

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text15_GotFocus(Index As Integer)
    'Modify By Cheng 2002/10/28
    Select Case Index
    Case 0
        Me.Text15(Index).SelStart = 0
        Me.Text15(Index).SelLength = 0
    Case Else
        TextInverse Text15(Index)
    End Select
End Sub

Private Sub Text5_GotFocus(Index As Integer)
  TextInverse Text5(Index)
End Sub

Private Sub Text5_Validate(Index As Integer, Cancel As Boolean)

   If Text5(Index) = "" Then
      MsgBox "質權設定期間不可空白，請重新輸入 !", vbCritical
      Cancel = True
   Else
      If Not ChkDate(Text5(Index)) Then
         MsgBox "質權設定期間不正確，請重新輸入 !", vbCritical
         Cancel = True
      Else
         If Index = 1 Then
            If pa(25) = "" Then
               MsgBox "專用期間止日不正確，請重新輸入 !", vbCritical
               Cancel = True
            Else
               If Val(DBDATE(Text5(1))) > Val(DBDATE(pa(25))) Then
                  MsgBox "質權設定期間止日大於專用期間止日，請重新輸入 !", vbCritical
                  Cancel = True
               Else
                  If ChkRange(Text5(0), Text5(1), "質權設定期間") = False Then Cancel = True
               End If
            End If
         Else
            If pa(24) = "" Then
               MsgBox "專用期間起日不正確，請重新輸入 !", vbCritical
               Cancel = True
            Else
               If Val(DBDATE(pa(24))) > Val(DBDATE(Text5(0))) Then
                  MsgBox "質權設定期間起日小於專用期間起日，請重新輸入 !", vbCritical
                  Cancel = True
               End If
            End If
         End If
      End If
   End If
   If Cancel = True Then TextInverse Text5(Index)
End Sub

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   If Text7 <> "" Then
      If ChgType(0) = False Then
         Cancel = True
      Else
         'Add by Morgan 2009/8/17
         If Text7.Tag <> Text7 Then
            PUB_SetChkResultDate pa(1), pa(9), cp(10), Text7, txtChkRltDate, cp, pa(8)
            Text7.Tag = Text7
         End If
      End If
   Else
      MsgBox "發文日不可空白 !", vbCritical
      Cancel = True
   End If
End Sub

Private Sub Text8_GotFocus()
  TextInverse Text8
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   If Text8 = "" Then
      MsgBox "關係人不可空白，請重新輸入 !", vbCritical
      Cancel = True
   End If
End Sub

Private Sub Text9_GotFocus()
  TextInverse Text9
End Sub

'Add By Cheng 2002/03/08
Private Function CheckDataIntegrity() As Boolean
Dim Cancel As Boolean
   'add by nickc 2008/05/01
   If IsDebt(pa(9), cp(9)) Then
        MsgBox "未收款且無 預定收款日 請轉告智權同仁！！", vbOKOnly, "警告！禁止發文！"
        GoTo IntegrityOrNot
   End If
Cancel = False

'檢查代理人欄位
Combo2_Validate Cancel
If Cancel = True Then
   Me.Combo2.SetFocus
   GoTo IntegrityOrNot
End If

CheckDataIntegrity = True
Exit Function

IntegrityOrNot:
CheckDataIntegrity = False
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

   'Added by Morgan 2021/12/15 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/15
   
For Each objTxt In Text5
   If objTxt.Enabled = True Then
      Cancel = False
      Text5_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Me.Text5(objTxt.Index).SetFocus
         Text5_GotFocus objTxt.Index
         Exit Function
      End If
   End If
Next

If Me.Text7.Enabled = True Then
   Cancel = False
   Text7_Validate Cancel
   If Cancel = True Then
      Me.Text7.SetFocus
      Text7_GotFocus
      Exit Function
   End If
End If

If Me.Text8.Enabled = True Then
   Cancel = False
   Text8_Validate Cancel
   If Cancel = True Then
      Me.Text8.SetFocus
      Text8_GotFocus
      Exit Function
   End If
End If

'Add by Morgan 2004/8/11
If txtCP84.Enabled = True Then
   Cancel = False
   txtCP84_Validate Cancel
   If Cancel = True Then
      txtCP84.SetFocus
      txtCP84_GotFocus
      Exit Function
   End If
End If

'Add by Morgan 2004/9/14
If Combo2.Enabled = True Then
   Cancel = False
   Combo2_Validate Cancel
   If Cancel = True Then
      Combo2.SetFocus
      Exit Function
   End If
End If

   'Add by Morgan 2005/7/15
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   '2005/7/14 END
   
   'Added by Lydia 2021/05/25 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
   If Pub_ChkACS112isNull(pa(1), pa(2), pa(3), pa(4), txtCP113) = True Then
         txtCP113.SetFocus
         txtCP113_GotFocus
         Exit Function
   End If
   'end 2021/05/25
TxtValidate = True
End Function

'Added by Lydia 2021/05/25
Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

'Added by Lydia 2021/05/25
Private Sub txtCP113_Validate(Cancel As Boolean)
   If txtCP113 <> "" Then
      If Not IsNumeric(txtCP113) Then
         MsgBox "請輸入數字！", vbExclamation
         txtCP113.SetFocus
         txtCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

'Add by Morgan 2004/8/11
Private Sub txtCP84_GotFocus()
   TextInverse txtCP84
End Sub
'Add by Morgan 2004/8/11
Private Sub txtCP84_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub
'Add by Morgan 2004/8/11
Private Sub txtCP84_Validate(Cancel As Boolean)
   '台灣
   If pa(9) = "000" Then
      If Val(txtCP84.Text) <> Val(cp(17)) And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
         If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & cp(17) & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
            txtCP84.Tag = txtCP84.Text
         Else
            txtCP84_GotFocus
            Cancel = True
         End If
      End If
   End If
End Sub

'Add by Morgan 2005/7/15
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   cp(110) = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
         'Modified by Morgan 2021/12/15 Forms2.0 改用模組
         'cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         cp(110) = cp(110) & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         bolCheck = True
      End If
   Next
   If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
   If bolCheck = True Then
      Text10 = ""
   Else
      Text10 = "N"
      If MsgBox("未勾選代理人，確定不出名？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then
         Cancel = True
      End If
   End If
End Sub

'Add by Morgan 2009/8/17
Private Sub lblCaseFee_Click()
   frm12040102_2.txtCF(1) = cp(1)
   frm12040102_2.txtCF(2) = pa(9)
   frm12040102_2.txtCF(3) = cp(10)
   frm12040102_2.Show vbModal
   If Val(Text7) > 0 Then
      PUB_SetChkResultDate pa(1), pa(9), cp(10), Text7, txtChkRltDate, cp, pa(8)
   End If
End Sub

Private Sub lblCaseFee_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseDown lblCaseFee, lblCaseFees
End Sub

Private Sub lblCaseFee_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseUp lblCaseFee, lblCaseFees
End Sub

Private Sub txtChkRltDate_Validate(Cancel As Boolean)
   If txtChkRltDate <> "" Then
      If ChkDate(txtChkRltDate) = False Then
         Cancel = True
      End If
   End If
End Sub
