VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040104_h 
   BorderStyle     =   1  '單線固定
   Caption         =   "內專發文-重新委任"
   ClientHeight    =   5580
   ClientLeft      =   -220
   ClientTop       =   1360
   ClientWidth     =   9350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9350
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   5190
      MaxLength       =   4
      TabIndex        =   6
      Top             =   3960
      Width           =   540
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   2580
      MaxLength       =   1
      TabIndex        =   3
      Top             =   2855
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   5290
      TabIndex        =   4
      Top             =   3570
      Width           =   1200
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   5130
      MaxLength       =   1
      TabIndex        =   15
      Top             =   1110
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   4
      Left            =   5130
      MaxLength       =   1
      TabIndex        =   17
      Top             =   1380
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   3
      Left            =   900
      MaxLength       =   1
      TabIndex        =   16
      Top             =   1380
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   5
      Left            =   900
      MaxLength       =   1
      TabIndex        =   18
      Top             =   1650
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   900
      MaxLength       =   1
      TabIndex        =   14
      Top             =   1110
      Width           =   240
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   5205
      MaxLength       =   1
      TabIndex        =   12
      Top             =   4620
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   0
      Left            =   5205
      MaxLength       =   1
      TabIndex        =   10
      Top             =   4290
      Width           =   255
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   1476
      MaxLength       =   1
      TabIndex        =   11
      Top             =   4620
      Width           =   255
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   1
      Left            =   6264
      MaxLength       =   1
      TabIndex        =   7
      Top             =   3960
      Width           =   255
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8412
      TabIndex        =   22
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6360
      TabIndex        =   20
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7188
      TabIndex        =   21
      Top             =   45
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   400
      Index           =   3
      Left            =   5130
      TabIndex        =   19
      Top             =   45
      Width           =   1200
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   900
      MaxLength       =   8
      TabIndex        =   5
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   0
      Left            =   1476
      MaxLength       =   1
      TabIndex        =   9
      Top             =   4290
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1020
      MaxLength       =   3
      TabIndex        =   26
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   25
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2340
      MaxLength       =   1
      TabIndex        =   24
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2580
      MaxLength       =   2
      TabIndex        =   23
      Top             =   720
      Width           =   375
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   984
      Left            =   7620
      TabIndex        =   8
      Top             =   3936
      Width           =   1500
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;1736"
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
      Top             =   2070
      Width           =   7740
      VariousPropertyBits=   671107099
      MaxLength       =   160
      Size            =   "13652;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text12 
      Height          =   465
      Left            =   1110
      TabIndex        =   13
      Top             =   4980
      Width           =   7995
      VariousPropertyBits=   -1467987941
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "14102;820"
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
      Top             =   2310
      Width           =   7740
      VariousPropertyBits=   671107099
      MaxLength       =   250
      Size            =   "13652;529"
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
      Top             =   2550
      Width           =   7740
      VariousPropertyBits=   671107099
      MaxLength       =   160
      Size            =   "13652;529"
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
      Left            =   4410
      TabIndex        =   67
      Top             =   4005
      Width           =   765
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "3.  總委任書正本"
      Height          =   180
      Index           =   1
      Left            =   3120
      TabIndex        =   66
      Top             =   3375
      Width           =   1305
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "1.  委任書正本"
      Height          =   180
      Index           =   0
      Left            =   3135
      TabIndex        =   65
      Top             =   2895
      Width           =   1125
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "申請書附件:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   64
      Top             =   2895
      Width           =   945
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "2.  委任書後補"
      Height          =   180
      Index           =   2
      Left            =   3135
      TabIndex        =   63
      Top             =   3135
      Width           =   1125
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "4.  總委任書影本"
      Height          =   180
      Index           =   3
      Left            =   3120
      TabIndex        =   62
      Top             =   3615
      Width           =   1305
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "正本存於                                號卷內:"
      Height          =   180
      Index           =   4
      Left            =   4500
      TabIndex        =   61
      Top             =   3615
      Width           =   2745
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "委任書(四選一):"
      Height          =   180
      Index           =   1
      Left            =   1320
      TabIndex        =   60
      Top             =   2895
      Width           =   1245
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人:"
      Height          =   180
      Left            =   6588
      TabIndex        =   59
      Top             =   3972
      Width           =   948
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "申請人2:"
      Height          =   180
      Left            =   4380
      TabIndex        =   58
      Top             =   1155
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請人1:"
      Height          =   180
      Left            =   180
      TabIndex        =   57
      Top             =   1155
      Width           =   675
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   3
      Left            =   1200
      TabIndex        =   56
      Top             =   1155
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5186;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   4
      Left            =   5445
      TabIndex        =   55
      Top             =   1155
      Width           =   3120
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5503;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "是否修改通知函內容:       (Y:Word)"
      Height          =   180
      Index           =   4
      Left            =   3525
      TabIndex        =   54
      Top             =   4665
      Width           =   2670
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(日):"
      Height          =   180
      Left            =   165
      TabIndex        =   53
      Top             =   2580
      Width           =   1065
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(英):"
      Height          =   180
      Index           =   1
      Left            =   165
      TabIndex        =   52
      Top             =   2340
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(中):"
      Height          =   180
      Left            =   165
      TabIndex        =   51
      Top             =   2100
      Width           =   1065
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   150
      X2              =   9120
      Y1              =   2010
      Y2              =   2010
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   150
      X2              =   9120
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "是否修改申請書內容:       (Y:Word)"
      Height          =   180
      Left            =   3525
      TabIndex        =   50
      Top             =   4335
      Width           =   2670
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   13
      Left            =   7860
      TabIndex        =   49
      Top             =   720
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2328;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "是否列印申請書:        (N:不印)"
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   48
      Top             =   4335
      Width           =   2310
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "是否列印通知函:       (N:不印)"
      Height          =   180
      Index           =   3
      Left            =   150
      TabIndex        =   47
      Top             =   4665
      Width           =   2925
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   1
      Left            =   7020
      TabIndex        =   46
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家"
      Height          =   180
      Index           =   1
      Left            =   4380
      TabIndex        =   45
      Top             =   510
      Width           =   720
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   12
      Left            =   5220
      TabIndex        =   44
      Top             =   510
      Width           =   1740
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3069;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   9
      Left            =   7860
      TabIndex        =   43
      Top             =   510
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2328;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   7
      Left            =   1200
      TabIndex        =   42
      Top             =   1695
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5186;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   6
      Left            =   5445
      TabIndex        =   41
      Top             =   1425
      Width           =   3120
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5503;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   5
      Left            =   1200
      TabIndex        =   40
      Top             =   1425
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5186;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   2
      Left            =   5220
      TabIndex        =   39
      Top             =   720
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2117;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   1
      Left            =   5220
      TabIndex        =   38
      Top             =   930
      Width           =   1215
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2143;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   0
      Left            =   1020
      TabIndex        =   37
      Top             =   510
      Width           =   480
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "847;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Left            =   150
      TabIndex        =   36
      Top             =   4980
      Width           =   765
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   7020
      TabIndex        =   35
      Top             =   510
      Width           =   585
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Left            =   165
      TabIndex        =   34
      Top             =   4005
      Width           =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Left            =   180
      TabIndex        =   33
      Top             =   510
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Index           =   0
      Left            =   4380
      TabIndex        =   32
      Top             =   930
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   180
      TabIndex        =   31
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Index           =   0
      Left            =   4380
      TabIndex        =   30
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "申請人3:"
      Height          =   180
      Left            =   180
      TabIndex        =   29
      Top             =   1425
      Width           =   675
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "申請人4:"
      Height          =   180
      Left            =   4380
      TabIndex        =   28
      Top             =   1425
      Width           =   675
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "申請人5:"
      Height          =   180
      Left            =   180
      TabIndex        =   27
      Top             =   1695
      Width           =   675
   End
End
Attribute VB_Name = "frm040104_h"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/14 改成Form2.0 (Text15,Text12,lstNameAgent,Label2)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
'2007/6/28 add by sonia
Option Explicit
Dim strReceiveNo As String '總收文號
Dim m_bolActive As Boolean 'Active事件是否已觸發
Dim pa() As String, cp() As String
Dim intWhere As Integer
Dim m_bolReturn As Boolean '共用回傳值
Dim m_CP09s As String, m_CP123s As String 'Add by Morgan 2009/3/23 收文號,是否算發文室案件
Dim m_CP130 As String 'Add by Morgan 2009/4/28 發文-主管機關
Dim m_bolFMP As Boolean 'Added by Lydia 2023/06/20 是否為FMP案
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/06/20 是否為寰華

Private Sub cmdok_Click(Index As Integer)
   ' 設定滑鼠游標為等待狀態
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 0, 3 '確定,同時發文
         'Modify by Morgan 2010/2/10 改呼叫函數方式以便鎖定按鍵
         cmdOK(Index).Enabled = False
         If Not Process(Index) Then
            cmdOK(Index).Enabled = True
         Else
            If Index = 3 Then
               'Add By Sindy 2013/5/20
               If frm040104_1.bolIsEMPFlow = True Then
                  frm090202_4.QueryData
               End If
               '2013/5/20 End
               frm040104_1.Show
               frm040104_1.ReQuery
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
            End If
            Unload Me
         End If
         
      Case 1
         'Add By Sindy 2013/5/20
         If frm040104_1.bolIsEMPFlow = True Then
            Unload frm040104_1
            frm090202_4.Show
            frm090202_4.QueryData
         Else
         '2013/5/20 End
            frm040104_1.Show
         End If
         Unload Me
      Case 2
         'Add By Sindy 2013/5/20
         If frm040104_1.bolIsEMPFlow = True Then
            Unload frm040104_1
            frm090202_4.Show
            frm090202_4.QueryData
         Else
         '2013/5/20 End
            Unload frm040104_1
         End If
         Unload Me
   End Select
   ' 設定滑鼠游標為預設
   Screen.MousePointer = vbDefault
End Sub

Private Function Process(Index As Integer) As Boolean
   
   If TxtValidate = True Then
      'Add by Morgan 2009/4/28
      If ModifyDispatchCp130(cp(9), m_CP09s, m_CP123s, m_CP130, Text9) = False Then
         Exit Function
      End If
      If m_CP123s = "Y" Then
      'end 2009/4/28
         'Add by Morgan 2009/3/23 設定是否算發文室案件
         'modify by sonia 2014/6/23 加傳發文規費, P-108903
         If ModifyDispatch(cp(9), m_CP09s, m_CP123s, 0, Text9) = False Then
             Exit Function
         End If
      End If
      'Add by Amy 2014/10/14 P台灣案發文控制
      If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
        If pa(1) = "P" And cp(9) < "C" And pa(9) = 台灣國家代號 Then
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
            'If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" And Text8(0) = "N" And PUB_CheckPDF2(cp(9), 1, True, strExc(0)) = False Then
            If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" And Text8(0) = "N" Then
               If PUB_CheckPDF2(cp(9), 1, True, strExc(0)) = False Then
            'end 2015/3/17
                  MsgBox "無申請書PDF檔 ,不可發文!", vbInformation
                  Exit Function
               End If 'Added by Morgan 2015/3/17
            End If
        End If
      End If
      'end 2014/10/14
      If FormSave = True Then
         Process = True
         PrintLetter
      End If
   End If
   
End Function
'資料檢查
Private Function TxtValidate() As Boolean

   Dim m_DiscType As String   '減免身分
   Dim i As Integer
   Dim Cancel As Boolean

On Error GoTo ErrHnd

   TxtValidate = False
   
   'Added by Morgan 2021/12/15 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/15
   
   'add by nickc 2008/05/01
   If IsDebt(pa(9), cp(9)) Then
        MsgBox "未收款且無 預定收款日 請轉告智權同仁！！", vbOKOnly, "警告！禁止發文！"
        Exit Function
   End If
   
   If Text6.Text = "" Then
      MsgBox "委任書不可空白，請重新輸入 !", vbCritical
      Text6.SetFocus
      Exit Function
   End If
   
   If Text6 = "4" And Text7 = "" Then
      MsgBox "請輸入委任書正本之申請案號 !", vbCritical
      Text7.SetFocus
      Exit Function
   End If
   
   If Text9.Text = "" Then
      MsgBox "發文日不可空白，請重新輸入 !", vbCritical
      Text9.SetFocus
      Exit Function
   End If
   
   If Text9.Enabled = True Then
      Cancel = False
      Text9_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Text7.Enabled = True Then
      Cancel = False
      Text7_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If

   If pa(9) = "000" Then
      m_DiscType = ""
      For i = 1 To 5
         m_DiscType = m_DiscType & txtAD(i).Text
         If txtAD(i).Enabled = True Then
            If txtAD(i).Text = "" Then
               MsgBox "申請人【" & pa(25 + i) & "-" & Label2(i + 2) & "】減免身分不可空白", vbInformation
               txtAD(i).SetFocus
               txtAD_GotFocus i
               Exit Function
            '學校
            ElseIf (txtAD(i).Text = "2") Then
               '變更
               If (txtAD(i).Tag <> "2" And txtAD(i).Tag <> "") Then
                  If MsgBox("確定要變更申請人【" & pa(25 + i) & "-" & Label2(i + 2) & "】減免身分為【學校】？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                     txtAD(i).SetFocus
                     txtAD_GotFocus i
                     Exit Function
                  End If
               End If
            '公司
            ElseIf (txtAD(i).Text = "3") Then
               '新增或變更
               If (txtAD(i).Tag <> "3") Then
                  If MsgBox("申請人【" & pa(25 + i) & "-" & Label2(i + 2) & "】的減免身分將設定為【中小企業】，確定有【證明文件】存放於本卷？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                     txtAD(i).SetFocus
                     txtAD_GotFocus i
                     Exit Function
                  End If
               End If
            '不可減免
            ElseIf (txtAD(i).Text = "N") Then
               '身分變更
               If (txtAD(i).Tag <> "N" And txtAD(i).Tag <> "") Then
                  If MsgBox("確定要變更申請人【" & pa(25 + i) & "-" & Label2(i + 2) & "】減免身分為【不可減免】？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                     txtAD(i).SetFocus
                     txtAD_GotFocus i
                     Exit Function
                  End If
               End If
            End If
         End If
      Next
      If InStr(m_DiscType, "N") > 0 Then
         cp(81) = "N"
      Else
         cp(81) = "Y"
      End If
   End If
      
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   
   'Added by Lydia 2021/05/25 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
   If Pub_ChkACS112isNull(pa(1), pa(2), pa(3), pa(4), txtCP113) = True Then
         txtCP113.SetFocus
         txtCP113_GotFocus
         Exit Function
   End If
   'end 2021/05/25
   
   TxtValidate = True
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   
End Function
'存檔
Private Function FormSave() As Boolean
   Dim strTmp(0 To 5) As String, ii As Integer
   
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   
   'Added by Morgan 2013/6/7 自 lstNameAgent_Validate 移來,否則若觸發 Form_Activate 事件會跑 ReadPatent 導致 cp(110) 被清除
   cp(110) = ""
   If lstNameAgent.Visible = True Then
      For ii = 0 To lstNameAgent.ListCount - 1
         If lstNameAgent.Selected(ii) = True Then
            'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
            'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
            'Modified by Morgan 2021/12/15f Forms2.0 改用模組
            'cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
            cp(110) = cp(110) & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         End If
      Next
      If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
   End If
   'end 2013/6/7

   '設定客戶減免身分
   If pa(9) = "000" Then
      For ii = 1 To 5
         If txtAD(ii).Enabled = True Then
            '身分有變更才要做
            If txtAD(ii).Tag <> txtAD(ii).Text Then
               '不可減免
               If txtAD(ii).Text = "N" Then
                  strSql = PUB_GetADSQL(pa(25 + ii), pa(9), "N")
               '自然人,學校不用證明
               ElseIf (txtAD(ii).Text = "1" Or txtAD(ii).Text = "2") Then
                  strSql = PUB_GetADSQL(pa(25 + ii), pa(9), "Y", txtAD(ii).Text)
               '公司
               Else
                  '原來沒有減免資料或不可減免
                  If txtAD(ii).Tag = "" Or txtAD(ii).Tag = "N" Then
                     strSql = PUB_GetADSQL(pa(25 + ii), pa(9), "Y", txtAD(ii).Text, pa(1), pa(2), pa(3), pa(4))
                  '修改減免身分別
                  Else
                     strSql = PUB_GetADSQL(pa(25 + ii), pa(9), "Y", txtAD(ii).Text)
                  End If
               End If
               cnnConnection.Execute strSql
            End If
         End If
      Next
   End If
   
   'Modified by Lydia 2021/05/25 +CP113工作時數
   'Modified by Lydia 2023/06/20 +CP14
   strSql = "UPDATE CASEPROGRESS SET CP27=" & CNULL(TransDate(Text9, 2)) & ",CP22=" & CNULL(Text8(1)) & _
               ",cp110=" & CNULL(cp(110)) & " ,cp113=" & CNULL(txtCP113, True) & ",CP14=" & CNULL(cp(14)) & _
               " WHERE CP09='" & strReceiveNo & "'"
   cnnConnection.Execute strSql
   
   'Modify by Amy 2014/09/09 for 台灣案電子化
   If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
    If pa(9) = 台灣國家代號 Then
         cnnConnection.Execute "delete LetterProgress where lp01='" & strReceiveNo & "'", intI 'Added by Morgan 2016/2/26 可能會重新發文
        '*沒客戶通知函
        If Text8(2) = "N" Then
            'Modify by Amy 2015/02/13 原:判斷同一天沒有其他有規費的發文
              '1.    電子送件且規費>0 (此無規費,故不考慮)
              '2.非電子送件且經發文室要計件,有回執
            'Mark by Amy 2015/03/06 回執改至PUB_UpdateLP19做
'            If Left(m_CP123s, 1) = "Y" Then
'                strExc(1) = PUB_GetLetterJudge(pa(1), cp(10))
'                PUB_AddLetterProgress strReceiveNo, 1, False, strExc(1), False, pa(26), cp(10), pa(75), True
'            End If
            
        '*有客戶通知函
        Else
            'Modified by Morgan 2018/8/1
            'strExc(1) = PUB_GetLetterJudge(pa(1), cp(10), , , pa(1), pa(2), pa(3), pa(4))
            strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), cp(10))
            'Modify by Amy 2015/02/13 修改判斷條件
            'PUB_AddLetterProgress strReceiveNo, 0, True, strExc(1), False, pa(26), cp(10), pa(75), False
              '1.　電子送件有規費的有收據(此無規費,故不考慮)；無規費的無回執
              '2.非電子送件要計件的有回執；不計件的無回執
            'Mark by Amy 2015/03/06 回執改至PUB_UpdateLP19做
'            If cp(118) = "Y" Then
'                PUB_AddLetterProgress strReceiveNo, 0, True, strExc(1), False, pa(26), cp(10), pa(75), False
'            Else
                If Left(m_CP123s, 1) = "Y" Then
                  PUB_AddLetterProgress strReceiveNo, 1, True, strExc(1), False, pa(26), cp(10), pa(75), True
               Else
                  PUB_AddLetterProgress strReceiveNo, 0, True, strExc(1), False, pa(26), cp(10), pa(75), False
               End If
'            End If
            'end 2015/03/06
        End If
        '*有申請書
        If Text8(0) <> "N" Then
            If ExistCheck("AppForm", "AF01", strReceiveNo, "", False) = False Then
                 '新增申請書轉檔記錄
                 PUB_AddAppForm strReceiveNo
            End If
        End If
    End If
   End If
   'end 2014/09/09
   
   'Add by Morgan 2009/3/23
   PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130
   'Add by Amy 2015/02/13 更新收據/回執設定
   'Modify by Amy 2015/03/06 +發文日參數
   PUB_UpdateLP19 cp(1), cp(2), cp(3), cp(4), m_CP09s, m_CP123s, Text9
 
   cnnConnection.CommitTrans
   FormSave = True
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Function
'列印通知函及申請書
Private Sub PrintLetter()

   Dim strTmp As String, bolChk As Boolean
   
   '個案收重新委任印回覆單
   If pa(9) = "000" Then
      PUB_ReAsignInform pa(1), pa(2), pa(3), pa(4), strReceiveNo
   End If
         
   If Text8(2) <> "N" Then '通知函
      If Text5(1) = "Y" Then '是否修改通知函
         bolChk = True
      Else
         bolChk = False
      End If
      strTmp = "00"
      'Modify by Amy 2014/08/27 +傳strLetterRecNo
      'cancel by sonia 2025/6/12 杜協理要求取消
      'NowPrint strReceiveNo, "02", strTmp, bolChk, strUserNum, 0, , , , , , , , , , , , strReceiveNo
   End If
         
   If Text8(0) <> "N" Then '申請書
      If Text5(0) = "Y" Then
         bolChk = True
      Else
         bolChk = False
      End If
      strTmp = "01"
      StartLetter strReceiveNo, "01", strTmp
      'Modifyb Amy 2014/08/27 +傳strLetterRecNo 及台灣案申請書修改改開1105_1
      NowPrint strReceiveNo, "01", strTmp, bolChk, strUserNum, 0, , , , , , , , , , , , strReceiveNo
      If P台灣案電子化啟用日 <= Val(strSrvDate(1)) And Text8(0) <> "N" And Text5(0) = "Y" Then
            frm1105_1.m_RecNo = strReceiveNo
            'Modify By Sindy 2022/5/11 流水號要足6碼
            frm1105_1.m_PdfName = Text1 & Text2 & IIf(Text3 & Text4 = "000", "", "-" & Text3 & "-" & Text4) & "." & cp(10) & ".DATA.PDF"
            frm1105_1.Show
      End If
      'end 2014/08/27
   End If
   
End Sub

Private Function StartLetter(ByVal strReceiveNo As String, ByVal ET01 As String, ByVal ET03 As String) As Boolean

   Dim strTxt(1 To 20) As String, strTmp As String, strTmp1 As String
   Dim iAppCnt As Integer
   Dim stAppData(1 To 1, 0 To 3) As String
   Dim ii As Integer, iLen As Integer, i As Integer
   
   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   '1 委任書
   ii = ii + 1
   Select Case Text6
      Case "1"
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','勾選1','■')"
      Case "2"
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','勾選2','■')"
      Case "3"
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','勾選3','■')"
      Case "4"
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','勾選4','■')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','總委任書申請案號'," & CNULL(Text7.Text) & ")"
   End Select
   
    '2 發文日
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      "','其他日期'," & CNULL(Text9.Text) & ")"
  
   If Not ClsLawExecSQL(ii, strTxt) Then
       MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
   
End Function

Private Sub Form_Activate()
   Dim i As Integer
   If m_bolActive = True Then Exit Sub
   m_bolActive = True
   If pa(9) = "000" Then
      For i = 1 To 5
         If txtAD(i).Enabled = True And txtAD(i).Text = "" Then
            txtAD(i).SetFocus
            Exit For
         End If
      Next
      If i = 6 Then Text6.SetFocus
   Else
      Text6.SetFocus
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
   
   ReDim pa(1 To TF_PA)
   ReDim cp(TF_CP)
    
   ReadPatent
   
   cp(110) = "" '要清空,否則若重新發文會殘留前次發文資料,當新案有改出名人而本程序未改選將會造成不一致 Added by Morgan 2012/9/7
   
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   Text8(1).Visible = False
   lstNameAgent.Clear
   If pa(9) = "000" Then
      PUB_SetOurAgent lstNameAgent, pa(), cp(110), , True 'Modified by Morgan 2021/12/15 +傳入bForm2=True
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
      Text5(1).Enabled = False 'Add by Amy 2014/08/27 申請函不可修改 for 台灣案電子化
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   
   Label2(0) = strReceiveNo
End Sub

Private Sub ReadPatent()

   Dim Lbl As Object, txt As TextBox, i As Integer
   Dim strAD10 As String, strCU15 As String
   Dim strTempName As String, arrPA72

   For Each Lbl In Label2
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
         
   If ClsPDReadPatentDatabase(pa(), intWhere) Then
      '申請人名稱
      For i = 3 To 7
         If pa(i + 23) <> "" Then
            If ClsLawLawGetName(pa(i + 23), strTempName) Then
               Label2(i) = strTempName
            End If
         End If
      Next
      '案件名稱
      Text15(0) = pa(5)
      Text15(1) = pa(6)
      Text15(2) = pa(7)
      
      Text8(2).Text = "N"   '預設不印通知函
      '國家
      If pa(9) <> "" Then
         If ClsPDGetNation(pa(9), strExc(0)) Then Label2(12) = strExc(0)
      End If
      
   End If
   
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      If cp(13) <> "" Then
         If ClsPDGetStaff(cp(13), strExc(0)) Then Label2(1) = strExc(0)
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
         If ClsPDGetStaff(cp(14), strExc(0)) Then Label2(9) = strExc(0)
      End If
      Label2(2) = cp(6)
      Label2(13) = cp(7)
      If cp(27) = "" Then
         Text9 = strSrvDate(2)
      Else
         Text9 = cp(27)
      End If
      Text12 = cp(64)   '備註
      Text8(1) = cp(22) '是否出名
   End If
   
   '減免身分
   For i = 1 To 5
      txtAD(i).Enabled = False
      txtAD(i).Tag = ""
      txtAD(i).Text = ""
      If pa(25 + i) <> "" Then
         txtAD(i).Text = PUB_GetAD03(pa(25 + i), pa(9), strAD10, strCU15)
         txtAD(i).Tag = txtAD(i).Text
         '個人只可設定自然人(1)
         If strCU15 = "0" Then
            txtAD(i).Text = "1"
         'Added by Morgan 2014/7/15 學校也預設--玲玲
         ElseIf strCU15 = "2" Then
            txtAD(i).Text = "2"
         'end 2014/7/15
         '公司
         Else
            If txtAD(i).Text = "Y" Then
               txtAD(i).Text = strAD10
               txtAD(i).Tag = txtAD(i).Text
            End If
            txtAD(i).Enabled = True
         End If
      End If
   Next

    'Added by Lydia 2021/05/25
    txtCP113 = ""
    If cp(113) <> "" Then txtCP113 = cp(113)
    'end 2021/05/25

End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Set frm040104_h = Nothing 'Removed by Morgan 2021/12/15 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub Text5_GotFocus(Index As Integer)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text5(Index).IMEMode = 2
   CloseIme
   TextInverse Text5(Index)
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text6_GotFocus()
  TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 49 Or KeyAscii > 52) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
Dim rsTemp1 As New ADODB.Recordset
   
   If Text6 = "4" And Text7 = "" Then
      MsgBox "請輸入委任書正本之申請案號 !", vbCritical
      Cancel = True
   End If
   If Text7 <> "" Then
      strExc(0) = "SELECT PA26 FROM PATENT WHERE PA11='" & Text7 & "'"
      intI = 1
      Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With rsTemp1
            If IsNull(.Fields("PA26")) Then
               MsgBox "該委任書正本之申請案號之無申請人 !", vbCritical
               Cancel = True
            Else
               If ChangeCustomerS(.Fields("PA26")) <> pa(26) Then
                  MsgBox "該委任書正本之申請案號之申請人與本案不符 !", vbCritical
                  Cancel = True
               End If
            End If
         End With
      Else
         MsgBox "輸入委任書正本之申請案號不存在 !", vbCritical
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text7
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(Text9.Text)
   '2011/12/8 ADD BY SONIA
   If DBDATE(Val(Text9)) > DBDATE(PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)))) Then
      MsgBox "發文日期不正確或發文日大於系統日下一個工作日，請重新輸入 !", vbCritical
   End If
   '2011/12/8 END
   If Cancel = True Then TextInverse Text9
End Sub

Private Sub txtAD_GotFocus(Index As Integer)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtAD(Index).IMEMode = 2
   CloseIme
   TextInverse txtAD(Index)
End Sub

Private Sub txtAD_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Morgan 2014/7/15 學校改預設且不可改
   'If Not (KeyAscii = 8 Or KeyAscii = 50 Or KeyAscii = 51 Or KeyAscii = 78) Then
   If Not (KeyAscii = 8 Or KeyAscii = 51 Or KeyAscii = 78) Then
      KeyAscii = 0
   End If
End Sub

Private Sub Text8_GotFocus(Index As Integer)
  'edit by nickc 2007/07/11 切換輸入法改用API
  'Text8(Index).IMEMode = 2
  CloseIme
  TextInverse Text8(Index)
End Sub

Private Sub Text8_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text9_GotFocus()
  'edit by nickc 2007/07/11 切換輸入法改用API
  'Text9.IMEMode = 2
  CloseIme
  TextInverse Text9
End Sub
   
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   cp(110) = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
         'Modified by Morgan 2021/12/15f Forms2.0 改用模組
         'cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         cp(110) = cp(110) & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         bolCheck = True
      End If
   Next
   If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
   If bolCheck = True Then
      Text8(1) = ""
   Else
      Text8(1) = "N"
      If MsgBox("未勾選代理人，確定不出名？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then
         Cancel = True
      End If
   End If
End Sub

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
