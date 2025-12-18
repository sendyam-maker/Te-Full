VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060104_8 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專發文-設定質權/終止設定質權"
   ClientHeight    =   5196
   ClientLeft      =   276
   ClientTop       =   960
   ClientWidth     =   8892
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5196
   ScaleWidth      =   8892
   Begin VB.TextBox txtCP118 
      Height          =   270
      Left            =   1400
      MaxLength       =   1
      TabIndex        =   8
      Top             =   4530
      Width           =   375
   End
   Begin VB.TextBox txtPayToday 
      Height          =   264
      Left            =   4740
      MaxLength       =   1
      TabIndex        =   9
      Top             =   4530
      Width           =   255
   End
   Begin VB.TextBox txtRecDate 
      Height          =   270
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   10
      Top             =   4860
      Width           =   375
   End
   Begin VB.TextBox txtEmail 
      Height          =   270
      Left            =   4350
      MaxLength       =   1
      TabIndex        =   11
      Text            =   "Y"
      Top             =   4860
      Width           =   375
   End
   Begin VB.TextBox txtCP84 
      Height          =   288
      Left            =   7260
      TabIndex        =   2
      Top             =   2760
      Width           =   1140
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   4370
      MaxLength       =   7
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   1
      Left            =   2760
      MaxLength       =   7
      TabIndex        =   4
      Top             =   3090
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      ItemData        =   "frm060104_8.frx":0000
      Left            =   1020
      List            =   "frm060104_8.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   14
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7404
      TabIndex        =   13
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6520
      TabIndex        =   12
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   0
      Left            =   1410
      MaxLength       =   7
      TabIndex        =   3
      Top             =   3090
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   1410
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2490
      MaxLength       =   2
      TabIndex        =   18
      Top             =   930
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2250
      MaxLength       =   1
      TabIndex        =   17
      Top             =   930
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1410
      MaxLength       =   6
      TabIndex        =   16
      Top             =   930
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1020
      MaxLength       =   3
      TabIndex        =   15
      Top             =   930
      Width           =   495
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "是否電子送件:          (Y: 是)"
      Height          =   180
      Left            =   180
      TabIndex        =   53
      Top             =   4580
      Width           =   2240
   End
   Begin VB.Label lblPayToday 
      AutoSize        =   -1  'True
      Caption         =   "電子送件是否當日扣款:         (Y/N)"
      Height          =   180
      Left            =   2780
      TabIndex        =   52
      Top             =   4580
      Width           =   2660
   End
   Begin VB.Label lblRecDate 
      AutoSize        =   -1  'True
      Caption         =   "當天報告:             (Y:是)"
      Height          =   180
      Left            =   180
      TabIndex        =   51
      Top             =   4920
      Width           =   1820
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "Email維護:             (Y:是)"
      Height          =   180
      Left            =   3420
      TabIndex        =   50
      Top             =   4920
      Width           =   1860
   End
   Begin MSForms.TextBox Text9 
      Height          =   830
      Left            =   4020
      TabIndex        =   7
      Top             =   3630
      Width           =   3230
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "5689;1455"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text8 
      Height          =   830
      Left            =   180
      TabIndex        =   6
      Top             =   3630
      Width           =   3740
      VariousPropertyBits=   -1466941413
      MaxLength       =   60
      ScrollBars      =   2
      Size            =   "6588;1455"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   7260
      TabIndex        =   5
      Top             =   3090
      Width           =   1500
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;556"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   6330
      TabIndex        =   49
      Top             =   3135
      Width           =   900
   End
   Begin VB.Label lblCP84 
      AutoSize        =   -1  'True
      Caption         =   "發文規費:"
      Height          =   180
      Left            =   6480
      TabIndex        =   48
      Top             =   2805
      Width           =   770
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "催審期限:"
      Height          =   180
      Left            =   3530
      TabIndex        =   47
      Top             =   2805
      Width           =   770
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   100
      X2              =   8340
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   100
      X2              =   8340
      Y1              =   2670
      Y2              =   2670
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   10
      Left            =   3690
      TabIndex        =   46
      Top             =   2250
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2117;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   290
      Index           =   9
      Left            =   6240
      TabIndex        =   45
      Top             =   2250
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2646;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   290
      Index           =   8
      Left            =   930
      TabIndex        =   44
      Top             =   2250
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2646;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   290
      Index           =   7
      Left            =   1680
      TabIndex        =   43
      Top             =   1920
      Width           =   6990
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "12330;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   6
      Left            =   4860
      TabIndex        =   42
      Top             =   1590
      Width           =   2820
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4974;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   290
      Index           =   5
      Left            =   1020
      TabIndex        =   41
      Top             =   1620
      Width           =   2820
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4974;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   4
      Left            =   4860
      TabIndex        =   40
      Top             =   1260
      Width           =   2820
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4974;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   290
      Index           =   3
      Left            =   1020
      TabIndex        =   39
      Top             =   1260
      Width           =   2820
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4974;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   2
      Left            =   4860
      TabIndex        =   38
      Top             =   930
      Width           =   2820
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4974;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   1
      Left            =   4860
      TabIndex        =   37
      Top             =   600
      Width           =   2820
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4974;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   290
      Index           =   0
      Left            =   1050
      TabIndex        =   36
      Top             =   600
      Width           =   1740
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3069;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Left            =   4020
      TabIndex        =   35
      Top             =   3420
      Width           =   770
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "關係人:"
      Height          =   180
      Left            =   180
      TabIndex        =   34
      Top             =   3420
      Width           =   590
   End
   Begin VB.Label Label26 
      Caption         =   "~"
      Height          =   260
      Left            =   2610
      TabIndex        =   33
      Top             =   3110
      Width           =   140
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "質權設定期間:"
      Height          =   180
      Left            =   180
      TabIndex        =   32
      Top             =   3135
      Width           =   1125
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Left            =   180
      TabIndex        =   31
      Top             =   2805
      Width           =   585
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "業務區:"
      Height          =   180
      Left            =   2970
      TabIndex        =   30
      Top             =   2250
      Width           =   590
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   5430
      TabIndex        =   29
      Top             =   2250
      Width           =   770
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "申請人:"
      Height          =   180
      Left            =   180
      TabIndex        =   28
      Top             =   2250
      Width           =   585
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   180
      TabIndex        =   27
      Top             =   1920
      Width           =   768
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   4020
      TabIndex        =   26
      Top             =   600
      Width           =   585
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "專利種類:"
      Height          =   180
      Left            =   4020
      TabIndex        =   25
      Top             =   1620
      Width           =   768
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   180
      TabIndex        =   24
      Top             =   1620
      Width           =   768
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "彼所案號:"
      Height          =   180
      Left            =   4020
      TabIndex        =   23
      Top             =   1260
      Width           =   768
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   180
      TabIndex        =   22
      Top             =   1260
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發證日:"
      Height          =   180
      Left            =   4020
      TabIndex        =   21
      Top             =   930
      Width           =   588
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   180
      TabIndex        =   20
      Top             =   930
      Width           =   768
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Left            =   180
      TabIndex        =   19
      Top             =   600
      Width           =   585
   End
End
Attribute VB_Name = "frm060104_8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/16 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strReceiveNo As String
'Modify by Morgan 2005/8/4 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String, m_CP110 As String, m_AgentName As String

Dim intWhere As Integer
' 案件性質
Dim m_CP10 As String
'Add by Morgan 2004/8/11
Dim m_CP17 As String '收文規費
Dim m_PA143 As String 'Add by Morgan 2008/3/18
Dim m_CP09s As String, m_CP123s As String 'Add by Morgan 2009/3/20 收文號,是否算發文室案件
Dim m_CP130 As String 'Add by Morgan 2009/4/28 發文-主管機關
Dim m_CP14 As String 'ADD BY SONIA 2015/9/21 承辦人
Dim m_CP142 As String 'Add By Sindy 2015/12/17
Dim m_CP164 As String 'Add By Sindy 2021/4/20
Dim m_CP82 As String 'Add By Sindy 2024/12/2


Private Sub cmdok_Click(Index As Integer)
'Added by Sindy 2024/12/2
Dim strFilePath As String '記錄智慧局收文文號
Dim strNewCP64 As String '保留進度備註
'end 2024/12/2

   Select Case Index
      Case 0
         If CheckDataValid = False Then Exit Sub
         If TxtValidate = False Then Exit Sub
   
         'Add by Morgan 2008/3/18
         '若未辦或不辦重新委任時不可發文
         If PUB_Check928NotOk(pa) = True Then
            MsgBox "本案下一程序有重新委任之補文件未辦理，不可發文！"
            Exit Sub
         End If
         '若基本檔年費申請人是否出名為N時提醒存檔將取消
         m_PA143 = pa(143)
         If pa(143) = "N" Then
            MsgBox "年費申請人是否出名現為【N】，存檔時將自動取消！"
            m_PA143 = ""
         End If
         'end 2008/3/18
         
         'Added by Sindy 2024/12/2 + 是否電子送件
         strNewCP64 = Text9
         If txtCP118 = "Y" Then
            '電子送件也要記錄主管機關
            If ModifyDispatchCp130(strReceiveNo, m_CP09s, m_CP123s, m_CP130, Text7, , True) = False Then
               Exit Sub
            End If
            strExc(0) = InputBox("請輸入智慧局收文文號!!")
            If strExc(0) = "" Then
               Exit Sub
            Else
               strFilePath = strExc(0)  '記錄智慧局收文文號
               strNewCP64 = "智慧局收文文號:" & strExc(0) & ";" & Text9
            End If
         Else
         '2024/12/2 END
         
            'Add by Morgan 2009/4/28
            If ModifyDispatchCp130(strReceiveNo, m_CP09s, m_CP123s, m_CP130, Text7) = False Then
               Exit Sub
            End If
            If m_CP123s = "Y" Then
            'end 2009/4/28
               'Add by Morgan 2009/3/20 設定是否算發文室案件
               'modify by sonia 2014/6/23 加傳發文規費, P-108903
               If ModifyDispatch(strReceiveNo, m_CP09s, m_CP123s, txtCP84, Text7) = False Then
                   Exit Sub
               End If
               'end 2009/3/20
            End If
         End If
         'Added by Sindy 2024/12/2 +
         '依據輸入的智慧局收文號(受理號,ex: 1073066637-0)，將本機C:\E-SET\RdcDocDir\(收文號ex: 1073066637-0)的pdf檔自動搬移到卷宗區(by Phoebe);
         If txtCP118.Text = "Y" And strFilePath <> "" Then
             strExc(1) = m_CP82
             If Val(m_CP82) > 0 Then
                 If MsgBox("重新發文是否上傳檔案到卷宗區？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                      strExc(1) = ""
                 End If
             End If
             If Val(strExc(1)) = 0 Then
                If Pub_AutoEsetToCpp(True, pa(1), pa(2), pa(3), pa(4), pa(8), Label2(0).Caption, m_CP10, strFilePath, Text7.Text) = False Then
                      Exit Sub
                End If
             End If
         End If
         Text9.Text = strNewCP64 '檢查完畢，更新備註欄位
         '2024/12/2 END
         
         'Add by Sindy 2021/11/16 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If
         
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
        
         'Add by Morgan 2008/2/20 檢查代理人Email
         PUB_CheckEMail pa(75), pa(144)
         If pa(145) <> "" Then
            PUB_CheckEMail pa(75), pa(145)
         End If
         'end 2008/2/20
         
         If pa(1) = "FCP" Then
            'Add By Sindy 2016/11/16 特殊代理人彈訊息提醒
            If (PUB_GetST03(m_CP14) = "F21" Or PUB_GetST03(m_CP14) = "F51" Or PUB_GetST03(m_CP14) = "F52") And _
               Not (m_CP10 = "901" And m_CP10 = "902" And m_CP10 = "1202" And m_CP10 = "1002") Then
               strExc(0) = "select cp130 from caseprogress where cp09='" & strReceiveNo & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If "" & RsTemp.Fields(0) <> "" Then
                     If ChangeCustomerL(pa(75)) = "Y34440B30" Then
                        MsgBox "請當日優先請款報告!!"
                     End If
                  End If
               End If
            End If
         End If
         '2016/11/16 END
         
         'Added by Lydia 2024/03/06 外專機械設計組人員異動調整程式：內專協辦工程師完成送件之後，需通知外專工程師進行請款
         'Move by Lydia 2024/03/12 改使用Outlook草稿，從FormSave移出
         'Mark by Lydia 2024/04/18 FCP案直接併入frm060104_k的Outlook，所以也不用---Sharon
         'If pa(1) = "FCP" And Mid(m_CP14, 4, 1) = "9" Then
         '   Call Pub_SetEngMail(strReceiveNo)
         'End If
         ''end 2024/03/06
         'end 2024/04/18
         
         'Add By Sindy 2023/11/9
         If frm060104_1.bolIsEMPFlow = True Then
            frm090202_4.QueryData
         End If
         '2023/11/9 End
         '若有未發文資料顯示警告
         'Modify By Sindy 2023/11/9
         If PUB_GetCPunIssueDatas("" & Me.Text1.Text & "-" & Me.Text2.Text & "-" & IIf(Len("" & Me.Text3.Text) <= 0, "0", Me.Text3.Text) & "-" & IIf(Len("" & Me.Text4.Text) <= 0, "00", Me.Text4.Text)) Then
            frm060104_1.Show
            frm060104_1.ReQuery
         Else
            'Add By Sindy 2023/11/9
            If frm060104_1.bolIsEMPFlow = True Then
               Unload frm060104_1
            Else
            '2023/11/9 End
               frm060104_1.Show
               frm060104_1.Clear
            End If
         End If
            
         'Add By Sindy 2022/5/12
         If txtEmail.Text = "Y" Then
            frm060104_k.m_CP09 = strReceiveNo 'cp(9)
            frm060104_k.m_strRecDate = txtRecDate
            frm060104_k.Hide
            frm060104_k.cmdOK(0) = 1
            Unload frm060104_k
         End If
         '2022/5/12 END
   End Select
   
   frm060104_1.Show
   frm060104_1.Clear
   Unload Me
End Sub

Private Function FormSave() As Boolean
'Add By Cheng 2002/07/04
Dim ii As Integer
Dim intMax As Long
Dim stCP118 As String, stCP152 As String 'Added by Sindy 2024/12/2
   
'911105 nick transation
FormSave = True
 On Error GoTo CheckingErr
cnnConnection.BeginTrans

   ii = 0
   
   'Added by Sindy 2024/12/2
   '電子送件有規費的一律設自動扣款(同內專) --敏莉
   stCP118 = txtCP118
   stCP152 = ""
   If txtCP118 = "Y" And Val(txtCP84) > 0 Then
      stCP118 = "A"
      stCP152 = Pub_FcpSetPayToday("2", Text7.Text, txtPayToday.Text)
   End If
   'end 2024/12/2
   
   ' 90.06.27 modify by louis
   If m_CP10 = "707" Then
      ii = ii + 1
      'Modify by morgan 2005/8/4 加 cp110
      'MODIFY BY SONIA 2015/9/21 加 cp14
      'Modified by Sindy 2024/12/2 +CP118,CP152
      strExc(ii) = "UPDATE CASEPROGRESS " & _
                  "SET CP27=" & TransDate(Text7, 2) & ",cp14=" & CNULL(m_CP14) & "," & _
                  "cp64=" & CNULL(ChgSQL(Text9)) & ",cp84=" & Format(Val(txtCP84.Text)) & ", CP16=NVL(CP16,0)-NVL(CP17,0)+" & Format(Val(txtCP84.Text)) & ", CP17=" & Format(Val(txtCP84.Text)) & _
                  ", CP18=NVL(CP18,0),cp110=" & CNULL(m_CP110) & ",CP22=NULL,CP118='" & stCP118 & "',CP152=" & CNULL(stCP152, True) & _
                  " where cp09='" & strReceiveNo & "'"
                  
       '911105 nick transation
       cnnConnection.Execute strExc(ii)
                  
      'Modify By Cheng 2002/07/04
'      FormSave = objLawDll.ExecSQL(1, strExc)
   Else
      ii = ii + 1
      'Modify by morgan 2005/8/4 加 cp110
      'MODIFY BY SONIA 2015/9/21 加 cp14
      'Modified by Sindy 2024/12/2 +CP118,CP152
      strExc(ii) = "UPDATE CASEPROGRESS " & _
                  "SET CP27=" & TransDate(Text7, 2) & ",cp14=" & CNULL(m_CP14) & "," & _
                  "cp53=" & TransDate(Text5(0), 2) & ",cp54=" & TransDate(Text5(1), 2) & ",cp50=" & CNULL(Text8) & _
                  ",cp64=" & CNULL(ChgSQL(Text9)) & ",cp84=" & Format(Val(txtCP84.Text)) & ", CP16=NVL(CP16,0)-NVL(CP17,0)+" & Format(Val(txtCP84.Text)) & ", CP17=" & Format(Val(txtCP84.Text)) & ", CP18=NVL(CP18,0),cp110=" & CNULL(m_CP110) & ",CP22=NULL,CP118='" & stCP118 & "',CP152=" & CNULL(stCP152, True) & _
                  " where cp09='" & strReceiveNo & "'"
               
       '911105 nick transation
       cnnConnection.Execute strExc(ii)
   End If
   
   'Add By Cheng 2002/07/04
   '若有輸入催審期限
   If Me.Text6.Text <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
   'intMax = objPublicData.GetNextProgressNo
   intMax = GetNextProgressNo
      ii = ii + 1
      'Modified by Lydia 2025/11/12 改抓最近工作天+PUB_GetWorkDay1
      strExc(ii) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
         "NP07,NP08,NP09,NP10,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & _
         "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & 催審 & "," & _
         PUB_GetWorkDay1(TransDate(Text6.Text, 2), True) & "," & TransDate(Text6.Text, 2) & ",'" & strUserNum & "'," & intMax & ")"
         
       '911105 nick transation
       cnnConnection.Execute strExc(ii)
         
      intMax = intMax + 1
   End If
   
   'Add by Morgan 2008/3/18
   If m_PA143 <> pa(143) Then
      ii = ii + 1
      strExc(ii) = "update patent set pa143='" & m_PA143 & "' where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'"
      cnnConnection.Execute strExc(ii), intI
   End If
   
   'Added by Sindy 2024/12/2 FCP電子送件若發文時若有規費，則自動產生行事曆。
   If txtCP118 = "Y" And Val(txtCP84) > 0 And stCP152 <> "" Then
       If Pub_AddReceiptCalendar1(pa(1), pa(2), pa(3), pa(4), m_CP10, stCP152) = True Then
       End If
   End If
   'end 2018/09/11
   
   If ii >= 1 Then
      PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130 'Add by Morgan 2009/3/20
      cnnConnection.CommitTrans
   End If
   
'911105 nick
   Exit Function
CheckingErr:
   cnnConnection.RollbackTrans
   FormSave = False
   
End Function

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label2(7) = pa(5)
      Case "英"
         Label2(7) = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         Label2(7) = pa(7)
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm060104_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
   'Add by Morgan 2005/8/4
   ReDim pa(TF_PA)
   ReadPatent
   'Add by Morgan 2005/8/4
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   PUB_SetOurAgent lstNameAgent, pa(), m_CP110, , True
   'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1300
   lstNameAgent.Width = 1300

   Label2(0) = strReceiveNo
   Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call PUB_SendMailCache 'Added by Lydia 2024/03/06
   
   Set frm060104_8 = Nothing
End Sub

Private Sub ReadPatent()
Dim Lbl As Object, txt As Object, i As Integer
Dim strTempName As String

   For Each Lbl In Label2
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   Select Case pa(1)
      Case "FCP"
         If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
            Label2(2) = pa(21)
            Label2(3) = pa(22)
            Label2(4) = pa(77)
            ChgType (999) 'pa(8)
            Label2(7) = pa(5)
            If pa(26) <> "" Then ChgType (11) ' Label2(8)
         End If
      Case "FG"
      
   End Select
   'strExc(0) = "select cpm03,staff.st02 as st1,a0902,staff1.st02 as st2," & _
   '   "cp27,cp53,cp54,cp50,cp64 from caseprogress,casepropertymap," & _
   '   "staff,staff staff1,acc090 where cp09='" & strReceiveNo & "' and " & _
   '   "cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and " & _
   '   "cp13=staff1.st01(+) and a0901=staff.st03(+)"
   'Modify by Morgan 2004/8/12 Add cp17
   'MODIFY BY SONIA 2015/9/21 add cp14
   strExc(0) = "select cpm03,staff.st02 as st1,a0902,staff1.st02 as st2," & _
      "cp27,cp53,cp54,cp50,cp64,cp10,cp17,CP110,CP14,CP142,CP164,cp82,CP118 from caseprogress,casepropertymap," & _
      "staff,staff staff1,acc090 where cp09='" & strReceiveNo & "' and " & _
      "cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and " & _
      "cp13=staff1.st01(+) and cp12=a0901(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      m_CP110 = "" & .Fields("CP110")
      m_CP142 = "" & .Fields("CP142") 'Add By Sindy 2015/12/17
      m_CP164 = "" & .Fields("CP164") 'Add By Sindy 2021/4/20
      ' 90.06.27 modify by louis
      m_CP10 = .Fields("cp10")
      'Add by Morgan 2004/8/12
      m_CP17 = "" & .Fields("cp17")
      txtCP84.Tag = m_CP17
      txtCP84.Text = txtCP84.Tag
      
      If Not IsNull(.Fields(0)) Then Label2(5) = .Fields(0)
      
      'Add By Sindy 2024/12/2
      m_CP82 = "" & .Fields("cp82")
      '電子送件
      txtCP118 = ""
      If "" & .Fields("CP118") <> "" Then
          txtCP118 = "Y"
      End If
      '2024/12/2 END
      
      'MODIFY BY SONIA 2015/9/21 承辦人為外專程序時,改為操作人員
      'If Not IsNull(.Fields(1)) Then Label2(1) = .Fields(1)
      If Not IsNull(.Fields("cp14")) Then
         m_CP14 = GetFCPUser(.Fields("cp14"))   '存檔要更新
         If ClsPDGetStaff(m_CP14, strTempName) Then Label2(1) = strTempName
      End If
      'END 2015/9/21
      If Not IsNull(.Fields(2)) Then Label2(10) = .Fields(2)
      If Not IsNull(.Fields(3)) Then Label2(9) = .Fields(3)
      If Not IsNull(.Fields(4)) Then
         Text7 = TransDate(.Fields(4), 1)
      Else
         Text7 = strSrvDate(2)
      End If
      If Not IsNull(.Fields(5)) Then Text5(0) = TransDate(.Fields(5), 1)
      If Not IsNull(.Fields(6)) Then Text5(1) = TransDate(.Fields(6), 1)
      If Not IsNull(.Fields(7)) Then Text8 = .Fields(7)
      If Not IsNull(.Fields(8)) Then Text9 = .Fields(8)
   End If
   End With
   
   ' 90.06.27 modify by louis
   If m_CP10 = "707" Then
      EnableTextBox Text5(0), False
      EnableTextBox Text5(1), False
      EnableTextBox Text8, False
   Else
      EnableTextBox Text5(0), True
      EnableTextBox Text5(1), True
      EnableTextBox Text8, True
   End If
End Sub

Private Function ChgType(i As Integer) As Boolean
 Dim strTempName As String
   ChgType = False
   Select Case i
      Case 0 '發文日
         If Not ChkDate(Text7) Then
         ElseIf Val(Text7.Text) > PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1))) Then
            MsgBox "發文日大於系統日下一工作日, 請重新輸入!!!", vbExclamation + vbOKOnly
         Else
            ChgType = True
         End If
      'Add By Cheng 2002/07/04
      Case 6 '催審期限
         If ChkDate(Text6) Then
            ChgType = True
         End If
      Case 11
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.LawGetName(pa(26), strTempName) Then
         If ClsLawLawGetName(pa(26), strTempName) Then
            Label2(8) = strTempName
            ChgType = True
         End If
      Case 999
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetPatentTrademarkKind(專利, pA(8), strTempName, False, 台灣國家代號) = 1 Then
         If ClsPDGetPatentTrademarkKind(專利, pa(8), strTempName, False, 台灣國家代號) = 1 Then
            Label2(6) = strTempName
         End If
   End Select
End Function

Private Sub Text5_GotFocus(Index As Integer)
  TextInverse Text5(Index)
End Sub

Private Sub Text5_Validate(Index As Integer, Cancel As Boolean)
   If Not ChkDate(Text5(Index)) Then
      Cancel = True
      TextInverse Text5(Index)
   Else
      If Index = 1 Then
         If ChkRange(Text5(0), Text5(1), "質權期間") = False Then
            TextInverse Text5(0)
            Cancel = True
         End If
      End If
   End If
End Sub

Private Sub Text6_GotFocus()
TextInverse Text6
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If Text6 <> "" Then
      If Not ChgType(6) Then Cancel = True
   End If
   If Cancel = True Then Text6_GotFocus
End Sub

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub

Private Sub Text7_LostFocus()
'Add By Cheng 2002/07/04
' 重新計算催審期限
ReCaculateSpecDate
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   If Text7 <> "" Then
      'If ChgType(0) = False Then Cancel = True
      If Not ChgType(0) Then
            Cancel = True
      'Added by Sindy 2024/12/2 當發文日有改時
      Else
            If Text7.Tag <> Text7 Then
                  Text7.Tag = Text7
                  txtPayToday.Text = Pub_FcpSetPayToday("1", Text7.Text, txtCP118.Text)
            End If
      '2024/12/2 END
      End If
      
   Else
      MsgBox "發文日不可空白 !", vbCritical
      Cancel = True
   End If
   If Cancel = True Then Text7_GotFocus
End Sub

' 重新計算催審期限
Private Sub ReCaculateSpecDate()
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   
   Text6.Text = ""
   strSql = "SELECT CF05 FROM CASEFEE " & _
            "WHERE CF01='" & pa(1) & "' AND " & _
                  "CF02='" & pa(9) & "' AND " & _
                  "CF03='" & m_CP10 & "'"
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("CF05")) And rsTmp.Fields("CF05") <> 0 Then
         Text6.Text = TransDate(CompDate(2, Val(rsTmp.Fields("CF05")), TransDate(Text7, 2)), 1)
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
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

Private Function CheckDataValid() As Boolean
CheckDataValid = False
'檢查發文日
If Text7 <> "" Then
   If ChgType(0) = False Then
      Me.Text7.SetFocus
      Text7_GotFocus
      Exit Function
   End If
Else
   MsgBox "發文日不可空白 !", vbCritical
   Me.Text7.SetFocus
   Text7_GotFocus
   Exit Function
End If
'檢查質權設定期間
If Not ChkDate(Text5(0)) Then
   Me.Text5(0).SetFocus
   Text5_GotFocus 0
   Exit Function
End If
If Not ChkDate(Text5(1)) Then
   Me.Text5(1).SetFocus
   Text5_GotFocus 1
   Exit Function
Else
   If ChkRange(Text5(0), Text5(1), "質權期間") = False Then
      Me.Text5(0).SetFocus
      Text5_GotFocus 0
      Exit Function
   End If
End If

CheckDataValid = True
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

If Me.Text7.Enabled = True Then
   Cancel = False
   Text7_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text6.Enabled = True Then
   Cancel = False
   Text6_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

For Each objTxt In Text5
   If objTxt.Enabled = True Then
      Cancel = False
      Text5_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next


'Add by Morgan 2004/8/12
If txtCP84.Enabled = True Then
   Cancel = False
   txtCP84_Validate Cancel
   If Cancel = True Then
      txtCP84.SetFocus
      txtCP84_GotFocus
      Exit Function
   End If
End If

   'Add by Morgan 2005/8/4
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   
   'Add By Sindy 2015/12/17 檢查是否有指定送件日期,若有不可小於指定日期送件
   If m_CP142 <> "" Then
      'Modify By Sindy 2021/11/11 淑華說之後可以含當天發文
      'If m_CP142 >= strSrvDate(1) Then
      If m_CP142 > strSrvDate(1) Then
         'Add By Sindy 2021/4/20
         'Modify By Sindy 2021/10/20 + 3.之後
         If ((m_CP164 = "1" Or m_CP164 = "") And m_CP142 > strSrvDate(1)) Or _
            m_CP164 = "3" Then '1.當天 3.之後
         '2021/4/20 END
            MsgBox "有指定送件日期（" & ChangeWStringToTDateString(m_CP142) & "），不可提前送件!!!"
            Exit Function
         End If
      End If
   End If
   '2015/12/17 END
   
   'Added by Lydia 2018/09/11
   If txtCP118 = "Y" And Val(txtCP84) > 0 Then
      If txtPayToday = "" Then
         MsgBox "電子送件請輸入是否當日扣款(Y/N)！", vbExclamation
         txtPayToday.SetFocus
         Exit Function
      End If
   End If
   'end 2018/09/11
   
TxtValidate = True
End Function

Private Sub txtCP118_GotFocus()
   TextInverse txtCP118
   CloseIme
End Sub

Private Sub txtCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP118_Change()
    txtPayToday.Text = Pub_FcpSetPayToday("1", Text7.Text, txtCP118.Text)
End Sub

Private Sub txtPayToday_GotFocus()
   TextInverse txtPayToday
   CloseIme
End Sub

Private Sub txtPayToday_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub
'end 2018/09/11

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
      If Val(txtCP84.Text) <> Val(m_CP17) And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
         If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & m_CP17 & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
            txtCP84.Tag = txtCP84.Text
         Else
            txtCP84_GotFocus
            Cancel = True
         End If
      End If
   End If
End Sub
'Add by Morgan 2005/8/4
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   m_CP110 = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modify By Sindy 2021/5/10
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         '2021/5/10 END
         Cancel = False
      End If
   Next
   If Cancel = True Then
      MsgBox "出名代理人不可空白！", vbExclamation
   Else
      If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
      m_AgentName = Mid(m_AgentName, 2) 'Add By Sindy 2021/5/10
   End If
End Sub

'Add By Sindy 2022/5/17
Private Sub txtRecDate_GotFocus()
   TextInverse txtRecDate
End Sub
Private Sub txtRecDate_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub
Private Sub txtEmail_GotFocus()
   TextInverse txtEmail
End Sub
Private Sub txtEmail_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      Beep
      KeyAscii = 0
   End If
End Sub
Private Sub txtRecDate_Validate(Cancel As Boolean)
   If txtRecDate.Tag <> txtRecDate.Text Then
      If txtRecDate = "Y" Then
         txtEmail = "Y"
      End If
   End If
   txtRecDate.Tag = txtRecDate.Text
End Sub
'2022/5/17 END
