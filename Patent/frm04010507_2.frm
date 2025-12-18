VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010507_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人已收達/已提申"
   ClientHeight    =   5088
   ClientLeft      =   -60
   ClientTop       =   2928
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5088
   ScaleWidth      =   7500
   Begin VB.TextBox Text20 
      Height          =   270
      Left            =   4455
      MaxLength       =   1
      TabIndex        =   7
      Top             =   3900
      Width           =   375
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   3690
      TabIndex        =   48
      Top             =   3240
      Width           =   945
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   6165
      TabIndex        =   3
      Top             =   3240
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1710
      MaxLength       =   20
      TabIndex        =   6
      Top             =   3900
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      Height          =   270
      Left            =   1710
      MaxLength       =   15
      TabIndex        =   0
      Top             =   2910
      Width           =   2505
   End
   Begin VB.TextBox Text12 
      Height          =   270
      Left            =   5820
      MaxLength       =   8
      TabIndex        =   1
      Top             =   2910
      Width           =   1095
   End
   Begin VB.TextBox Text13 
      Height          =   270
      Left            =   1710
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   6270
      MaxLength       =   1
      TabIndex        =   9
      Top             =   4200
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   1710
      MaxLength       =   50
      TabIndex        =   8
      Top             =   4200
      Width           =   2835
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   4455
      MaxLength       =   8
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1710
      MaxLength       =   8
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmkok 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   6204
      TabIndex        =   15
      Top             =   90
      Width           =   1200
   End
   Begin VB.CommandButton cmkok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5400
      TabIndex        =   14
      Top             =   90
      Width           =   800
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  '沒有框線
      Height          =   465
      Left            =   210
      TabIndex        =   41
      Top             =   4530
      Visible         =   0   'False
      Width           =   7065
      Begin VB.CheckBox chk 
         Caption         =   "意見陳述書"
         Height          =   345
         Index           =   0
         Left            =   1140
         TabIndex        =   10
         Top             =   30
         Width           =   1335
      End
      Begin VB.CheckBox chk 
         Caption         =   "說明書"
         Height          =   345
         Index           =   1
         Left            =   2490
         TabIndex        =   11
         Top             =   30
         Width           =   945
      End
      Begin VB.CheckBox chk 
         Caption         =   "權利要求書"
         Height          =   345
         Index           =   2
         Left            =   3450
         TabIndex        =   12
         Top             =   30
         Width           =   1335
      End
      Begin VB.CheckBox chk 
         Caption         =   "圖式"
         Height          =   345
         Index           =   3
         Left            =   4800
         TabIndex        =   13
         Top             =   30
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "定稿文件："
         Height          =   180
         Left            =   150
         TabIndex        =   42
         Top             =   90
         Width           =   900
      End
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1320
      TabIndex        =   16
      Top             =   1050
      Width           =   5895
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "10398;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   2
      Left            =   3690
      TabIndex        =   54
      Top             =   2520
      Width           =   3660
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "6456;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "進度備註："
      Height          =   180
      Left            =   2760
      TabIndex        =   53
      Top             =   2520
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "是否收到副本：           (Y:是)"
      Height          =   180
      Left            =   3150
      TabIndex        =   52
      Top             =   3945
      Width           =   2220
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   180
      Left            =   360
      TabIndex        =   51
      Top             =   2520
      Width           =   900
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   10
      Left            =   1320
      TabIndex        =   50
      Top             =   2520
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2117;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label25 
      Caption         =   "幣別:"
      Height          =   180
      Index           =   4
      Left            =   3150
      TabIndex        =   49
      Top             =   3300
      Width           =   405
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "請輸入西元日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   5760
      TabIndex        =   47
      Top             =   3600
      Width           =   1680
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "年費繳費年度:"
      Height          =   180
      Index           =   0
      Left            =   4950
      TabIndex        =   46
      Top             =   3270
      Width           =   1125
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   9
      Left            =   1320
      TabIndex        =   45
      Top             =   2130
      Width           =   5955
      VariousPropertyBits=   27
      Size            =   "10504;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "繳費年度："
      Height          =   180
      Index           =   9
      Left            =   330
      TabIndex        =   44
      Top             =   2130
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   180
      Left            =   390
      TabIndex        =   43
      Top             =   3900
      Width           =   900
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "代理人D/N NO:"
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   40
      Top             =   2940
      Width           =   1155
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "帳單日期:"
      Height          =   180
      Index           =   2
      Left            =   4950
      TabIndex        =   39
      Top             =   2940
      Width           =   765
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "帳單金額:"
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   38
      Top             =   3240
      Width           =   765
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   5160
      TabIndex        =   25
      Top             =   1770
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   270
      X2              =   7350
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   270
      X2              =   7350
      Y1              =   2760
      Y2              =   2760
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   8
      Left            =   6090
      TabIndex        =   37
      Top             =   1770
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2117;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   6
      Left            =   1320
      TabIndex        =   35
      Top             =   1770
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2117;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   5
      Left            =   5910
      TabIndex        =   34
      Top             =   1410
      Width           =   1380
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2434;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   4
      Left            =   3480
      TabIndex        =   33
      Top             =   1410
      Width           =   1620
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2857;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   3
      Left            =   1320
      TabIndex        =   32
      Top             =   1410
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2117;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   1
      Left            =   4920
      TabIndex        =   31
      Top             =   720
      Width           =   2310
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "4075;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   0
      Left            =   1320
      TabIndex        =   30
      Top             =   720
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3757;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "列印客戶通知函         (N:不印)"
      Height          =   180
      Left            =   4950
      TabIndex        =   29
      Top             =   4200
      Width           =   2310
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "彼所案號："
      Height          =   180
      Left            =   390
      TabIndex        =   28
      Top             =   4200
      Width           =   900
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "代理人提申日："
      Height          =   180
      Left            =   3150
      TabIndex        =   27
      Top             =   3600
      Width           =   1260
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "代理人收達日："
      Height          =   180
      Left            =   390
      TabIndex        =   26
      Top             =   3600
      Width           =   1260
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Index           =   0
      Left            =   2760
      TabIndex        =   24
      Top             =   1770
      Width           =   720
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "發文日："
      Height          =   180
      Left            =   360
      TabIndex        =   23
      Top             =   1770
      Width           =   720
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Left            =   5160
      TabIndex        =   22
      Top             =   1410
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   2580
      TabIndex        =   21
      Top             =   1410
      Width           =   900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   360
      TabIndex        =   20
      Top             =   1410
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱"
      Height          =   180
      Left            =   360
      TabIndex        =   19
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "專利號數："
      Height          =   180
      Left            =   3960
      TabIndex        =   18
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   17
      Top             =   720
      Width           =   900
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   7
      Left            =   3480
      TabIndex        =   36
      Top             =   1770
      Width           =   1620
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2857;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm04010507_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (Combo1,Label2)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'整理 by Morgan 2005/7/8
Option Explicit
Dim intWhere As Integer
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String

Dim strReceiveNo As String, m_CP10 As String
'Add By Cheng 2003/04/02
Dim m_LastYearFee As String '目前年費最後已繳年度
Dim m_blnPrtContact As Boolean '是否列印聯絡單
'Add By Cheng 2003/04/16
'edit by nickc 2007/02/02
'Dim Ncp(T_CP) As String
Dim Ncp() As String

Dim strAutoNumber As String
Dim m_CP12 As String '業務區
Dim m_CP13 As String '智權人員
'92.11.3 ADD BY SONIA
Dim m_CP14 As String '承辦人
Dim m_CP44 As String 'CF代理人
Dim m_CP53 As String, m_CP54 As String 'add by sonia 2025/8/13 年費發文年度
Dim m_CP145 As String 'Added by Morgan 2012/2/10 是否已收副本
'Remove by Morgan 2009/10/1
'Dim m_902CP09 As String '香港回代收文號
Dim m_CP07 As String '法定期限 Add by Morgan 2009/11/18
Dim m_bFMP As Boolean 'Add by Morgan 2010/12/27
Dim m_CP47 As String 'Add by Lydia 2015/01/26
Dim m_NewCP09 As String 'Added by Morgan 2016/5/11
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/7 END


'Add By Cheng 2002/11/01
Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
Dim strTxt(1 To 5) As String
Dim strComString As String
Dim ii As Integer
    
    ii = 1
    EndLetter ET01, strReceiveNo, ET03, strUserNum
    If Me.frm.Visible Then
        strComString = ""
        If Me.chk(0).Value = vbChecked Then strComString = strComString & Me.chk(0).Caption & "、"
        If Me.chk(1).Value = vbChecked Then strComString = strComString & Me.chk(1).Caption & "、"
        If Me.chk(2).Value = vbChecked Then strComString = strComString & Me.chk(2).Caption & "、"
        If Me.chk(3).Value = vbChecked Then strComString = strComString & Me.chk(3).Caption & "、"
        If strComString <> "" Then
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','函知附件','" & Left(strComString, Len(strComString) - 1) & "')"
            ii = ii + 1
        End If
    End If
    'Add By Cheng 2002/11/21
    If pa(1) = "P" And (m_CP10 = "601" Or m_CP10 = "605") Then
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
           "','第幾年至幾年費','" & GetPayYear & "')"
        ii = ii + 1
    End If
    
    'Added by Lydia 2016/02/05 陳述意見之相關總收文號為無效宣告(803),用無效宣告取代專利基本檔-卷宗性質
    If pa(1) = "P" And pa(9) = "020" And ET03 = "02" And m_CP10 = 申復 Then
        strExc(0) = "select c3.cp10 from caseprogress c1,caseprogress c2,caseprogress c3 where c1.cp09='" & strReceiveNo & "' and c1.cp43=c2.cp09(+) and c2.cp43=c3.cp09(+) "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
           If "" & RsTemp(0) = "803" Then
              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                    "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                    "','無效宣告','無效宣告')"
                ii = ii + 1
           End If
        End If
    End If
    'end 2016/02/05
    
    If ii <> 1 Then
        'edit by nickc 2007/02/05 不用 dll 了
        'If Not objLawDll.ExecSQL(ii - 1, strTxt) Then
        If Not ClsLawExecSQL(ii - 1, strTxt) Then
           MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
        End If
    End If
End Sub

Private Sub cmkok_Click(Index As Integer)
Dim strTxt(1 To 4) As String, i As Integer, strTmp As String
'Add By Cheng 2002/11/05
Dim strBillNo As String '帳單編號
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim bolNoLetter As Boolean 'Add by Morgan 2010/2/4
'2011/12/19 ADD BY SONIA
Dim oContext As String
Dim oCustName As String
'2011/12/19 END
Dim strErrMsg As String 'Added by Morgan 2016/6/30
Dim iLP02 As Integer, stCP121 As String 'Added by Morgan 2016/7/6
Dim bolReceipt As Boolean 'Added by Morgan 2016/7/26 是否有收據
Dim strFullFileName As String 'Add By Sindy 2019/7/24
   
   Select Case Index
      Case 0 '確定
         'Add by Morgan 2005/10/18
         If TxtValidate = False Then Exit Sub
         
         Screen.MousePointer = vbHourglass
         
         'Add by Morgan 2009/11/18
         'Modify by Morgan 2009/12/25 年費除外--玲玲
         If m_CP07 <> "" And m_CP10 <> "605" Then
            strExc(1) = CompDate(2, 15, m_CP07)
            If Text3 <> "" Then
               strExc(2) = DBDATE(Text3)
               strExc(3) = "提申日"
            Else
               strExc(2) = DBDATE(Text2)
               strExc(3) = "收達日"
            End If
            
            If Val(strExc(2)) > Val(strExc(1)) Then
               'Modify by Morgan 2011/8/22 玲玲要再跟郭確認...
               'Modified by Morgan 2012/2/13 最後法定(+在途15日)遇假日順延 Ex.P-96588(因法限不一定正確,可選擇繼續)--玲玲
               'MsgBox strExc(3) & "已超過真實法限！"
               If MsgBox(strExc(3) & "已超過真實法限[ " & strExc(1) & " ]！是否確定要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               
            ElseIf Val(strExc(2)) > Val(m_CP07) Then
               If Left(m_CP12, 1) = "F" Then
                  MsgBox strExc(3) & "已超過法定期限，請留意！"
               End If
            End If
         End If
         
         '2005/4/6 ADD BY SONIA
         'Modified by Morgan 2016/7/26 +907,913除外
         'Modified by Morgan 2019/2/27 已收達不必控管--玲玲
         'If Text4 = "" And m_CP10 <> "907" And m_CP10 <> "913" Then
         If frm04010507_1.Text3.Text <> "1" And Text4 = "" And m_CP10 <> "907" And m_CP10 <> "913" Then
         'end 2019/2/27
            MsgBox "彼所案號欄不可空白!!!", vbExclamation + vbOKOnly
            Me.Text4.SetFocus
            Text4_GotFocus
            Screen.MousePointer = vbDefault
            Exit Sub
          End If
          '2005/4/6 END
         'Add By Cheng 2002/11/01
         If (Me.Text11.Text = "" Xor Me.Text12.Text = "") Or (Me.Text11.Text = "" Xor Me.Text13.Text = "") Or (Me.Text12.Text = "" Xor Me.Text13.Text = "") Then
             MsgBox "代理人D/N NO , 帳單日期 及 帳單金額 " & Chr(10) & Chr(13) & "三欄位必須同時輸入或不輸入資料!!!", vbExclamation + vbOKOnly
             If Me.Text11.Text = "" Then Screen.MousePointer = vbDefault: Me.Text11.SetFocus:    Text11_GotFocus:     Exit Sub
             If Me.Text12.Text = "" Then Screen.MousePointer = vbDefault: Me.Text12.SetFocus:    Text12_GotFocus:     Exit Sub
             If Me.Text13.Text = "" Then Screen.MousePointer = vbDefault: Me.Text13.SetFocus:    Text13_GotFocus:     Exit Sub
         End If
'        'Add By Cheng 2003/04/02
'        '預設不列印聯絡單
'        m_blnPrtContact = False
         '若案件性質為年費
         'Modify by Morgan 2011/8/19 +判斷已提申才要檢查年度--玲玲
         If m_CP10 = "605" And frm04010507_1.Text3 = "2" Then
            GetLastYearFee pa(1), pa(2), pa(3), pa(4)
            '若輸入的年費繳費年度與基本檔的最後已繳年度不符時
            If Me.Text6.Text <> m_LastYearFee Then
               'Modify By Cheng 2003/04/11
               '不列印聯絡單
'                If MsgBox("您輸入的年費繳費年度與最後已繳年度( " & m_LastYearFee & " )不符，是否要繼續並列印聯絡單???", vbExclamation + vbYesNo) = vbNo Then
               'Modify by Morgan 2005/4/11
               'If MsgBox("您輸入的年費繳費年度與最後已繳年度( " & m_LastYearFee & " )不符，是否要繼續???", vbExclamation + vbYesNo) = vbNo Then
               If MsgBox("您輸入的年費繳費年度與最後已繳年度( " & m_LastYearFee & " )不符，是否要繼續???", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                  Me.Text6.SetFocus
                  Text6_GotFocus
                  Screen.MousePointer = vbDefault
                  Exit Sub
               Else
                  bolNoLetter = True
               End If
'              '列印聯絡單
'              m_blnPrtContact = True
            End If
         End If
         'Add By Cheng 2003/12/25
         'Modify By Sindy 2009/06/17 若為專利處只須以代理人+代理人D/N No.做重覆檢核
         If pa(1) = "P" And Left(Trim(GetStaffDepartment(strUserNum)), 2) = "P1" Then
            '若有輸入代理人D/N No.
            If Me.Text11.Text <> "" Then
               If PUB_ChkDNDup("", ChangeCustomerL(m_CP44), Text11.Text) = True Then
                  Text11.SetFocus
                  Text11_GotFocus
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
         Else
            '若有輸入代理人D/N No.或帳單日期
            If Me.Text11.Text <> "" Or Me.Text12.Text <> "" Then
   'Modify by Morgan 2006/4/26 改Call共用函數
               If PUB_ChkDNDup(Text12.Text, ChangeCustomerL(m_CP44), Text11.Text) = True Then
                  Text11.SetFocus
                  Text11_GotFocus
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
   '2006/4/26 end
            End If
         End If
         
         'Added by Morgan 2016/6/30 非臺灣案電子化
         'Removed by Morgan 2025/8/13 帳單已全部都電子化
         'If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
         'end 2025/8/13
            If Me.Text11.Text <> "" And Me.Text12.Text <> "" And Me.Text13.Text <> "" Then
               '匯入該案帳單電子檔
               If Not PUB_ImportInvoice(pa(1), pa(2), pa(3), pa(4)) Then
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
         'End If
         'end 2016/6/30
         
         'Add by Sindy 2017/10/23 已收達信件要歸卷
         If m_strIR01 <> "" Then
'            'Add By Sindy 2022/7/1
'            If Left(Pub_StrUserSt03, 2) = "F2" Then
'               If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
'                  Screen.MousePointer = vbDefault
'                  Exit Sub
'               End If
'            Else
'            '2022/7/1 END
               If frm04010507_1.Text3.Text = "1" Then
                  '下載信件檔,上傳卷宗區
                  If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, strReceiveNo, IIf(Pub_StrUserSt03 = "F22", "ALTR", "ACK")) = False Then
                     Screen.MousePointer = vbDefault
                     Exit Sub
                  End If
               'Add By Sindy 2020/7/20
               Else
                  '下載信件檔,檢查信件是否開啟中,以免後面上傳卷宗區會無法儲存
                  If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, "", , , True) = False Then
                     Screen.MousePointer = vbDefault
                     Exit Sub
                  End If
               '2020/7/20 END
               End If
'            End If
         End If
         '2017/10/23 END
         
         'Added by Morgan 2016/7/6
         '已收達要檢查有檔案
         If frm04010507_1.Text3.Text = "1" And Left(Pub_StrUserSt03, 1) <> "F" Then
            If PUB_CheckAck(strReceiveNo, pa(1), pa(2), pa(3), pa(4), m_CP10) = False Then
               MsgBox "找不到代理人已收達電子檔(ACK)!!", vbExclamation
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
         End If
         'end 2016/7/6
         
         'Add By Cheng 2002/11/06
         'on error GoTo ErrorHandler
         cnnConnection.BeginTrans
         
On Error GoTo ErrorHandler 'Added by Morgan 2016/6/30
         
         '2005/4/27 MODIFY BY SONIA 改為西元日期
         'strTxt(1) = "UPDATE CASEPROGRESS SET CP45=" & CNULL(Text4) & ",CP46=" & CNULL(TransDate(Text2, 2)) & ",CP47=" & CNULL(TransDate(Text3, 2)) & " WHERE CP09='" & strReceiveNo & "'"
         'Modified by Morgan 2012/2/10 +CP145
         strTxt(1) = "UPDATE CASEPROGRESS SET CP45=" & CNULL(Text4) & ",CP46=" & CNULL(Text2) & ",CP47=" & CNULL(Text3) & ",CP145='" & Text20 & "' WHERE CP09='" & strReceiveNo & "'"
         '2005/4/27 END
         'Add By Cheng 2002/11/06
         cnnConnection.Execute strTxt(1)
         
         i = 2
         '若為收達
         If frm04010507_1.Text3.Text = "1" Then
            strTxt(i) = "UPDATE nextprogress SET NP06='Y' where np01='" & strReceiveNo & "' and np07=997"
            'Add By Cheng 2002/11/06
            cnnConnection.Execute strTxt(i), intI
            i = 3
            'Add by Morgan 2009/11/11 無提申及順稿期限者掛收達日+5個工作天的期限
            'Modify by Morgan 2010/1/14 回覆委任代理人不必掛提申期限
            'Modify by Morgan 2010/1/29 +203也不用
            'Modify by Morgan 2010/5/26 +202也不用
            'Modify by Morgan 2010/12/27 +非FMP 的601,605也不用
            'Modify by Morgan 2011/8/30 FMP 的601,605也不用 -- 玲玲
            'Modify by Morgan 2012/11/30 606也不用 -- 敏惠
            'If m_CP10 <> "936" And m_CP10 <> "203" And m_CP10 <> "202" And Not (m_bFMP = False And (m_CP10 = "601" Or m_CP10 = "605")) Then
            'Modified by Morgan 2014/3/7 FMP或PCT案的232(補優先權證明)也不管制提申
            'Modified by Morgan 2016/7/26 +907,913也不用
            'Modified by Morgan 2017/1/6 +950 --玲玲確認
            'Modified by Morgan 2017/1/9 +外專輸的收達領證或年費是要管制提申--敏莉 2017/2/13 取消--敏莉
            'Modified by Morgan 2017/1/16 +411催審也不要管制提申--茹曣
            'Modified by Morgan 2018/9/7 202補文件改要管制提申--茹曣
            'Modified by Morgan 2019/5/13 +953催提申--茹曣
            'Modified by Morgan 2019/5/30 +957詢問代理人--玲玲
            'Modified by Morgan 2019/6/14 +952催收達,954催公開--茹曣
            'Modified by Morgan 2019/7/26 +958代理人撰稿--茹曣
            If (m_CP10 <> "958" And m_CP10 <> "957" And m_CP10 <> "952" And m_CP10 <> "953" And m_CP10 <> "954" And m_CP10 <> "411" And m_CP10 <> "950" And m_CP10 <> "907" And m_CP10 <> "913" And m_CP10 <> "936" And m_CP10 <> "203" And m_CP10 <> "601" And m_CP10 <> "605" And m_CP10 <> "606" And Not ((m_bFMP Or pa(46) = "Y") And m_CP10 = "232")) Then
            'end 2017/1/9
               strExc(1) = CompWorkDay(5, DBDATE(Text2))
               '2010/7/6 MODIFY BY SONIA 若該收文號已輸入已提申則不必再掛提申期限P-084471(A99023479先輸提申才輸收達)
               'strSql = "insert into nextprogress(np01,np02,np03,np04,np05,np07,np08,np09,np10,np22)" & _
                  "select cp09,cp01,cp02,cp03,cp04,'998'," & strExc(1) & "," & strExc(1) & ",'" & strUserNum & "',np22" & _
                  " from caseprogress,(select max(np22)+1 np22 from nextprogress) x where cp09='" & strReceiveNo & "'" & _
                  " and not exists(select * from nextprogress where np01=cp09 and np07 in ('994','995','996','998') and np06 is null)"
               'Add By Sindy 2014/8/5 檢查此文號是否有案件暫緩,若有不可新增提申
               strExc(0) = "select cp09,cp10 from caseprogress where cp43='" & strReceiveNo & "' and cp10='950'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
               '2014/8/5 END
                  strSql = "insert into nextprogress(np01,np02,np03,np04,np05,np07,np08,np09,np10,np22)" & _
                     "select cp09,cp01,cp02,cp03,cp04,'998'," & strExc(1) & "," & strExc(1) & ",'" & strUserNum & "',np22" & _
                     " from caseprogress,(select max(np22)+1 np22 from nextprogress) x where cp09='" & strReceiveNo & "'" & _
                     " AND CP47 IS NULL and not exists(select * from nextprogress where np01=cp09 and np07 in ('994','995','996','998') and np06 is null)"
                  cnnConnection.Execute strSql, intI
               End If
            End If
            'end 2009/11/11
            
         '若為提申
         ElseIf frm04010507_1.Text3.Text = "2" Then
            'Modify by Morgan 2005/10/18收達也要上'Y'
            'strTxt(i) = "UPDATE nextprogress SET NP06='Y' where np01='" & strReceiveNo & "' and np07=998 "
            'Modify by Morgan 2009/7/13 +995,996
            strTxt(i) = "UPDATE nextprogress SET NP06='Y' where np01='" & strReceiveNo & "' and np07 in ('995','996','997','998')"
            'Add By Cheng 2002/11/06
            cnnConnection.Execute strTxt(i)
            i = 3
            
            'Add By Sindy 2019/7/24 消催審期限
            If m_CP10 = "958" Then '958.代理人撰稿
               strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
                        "WHERE NP01 = '" & strReceiveNo & "' AND " & _
                              "NP07 = '411' AND NP06 IS NULL "
               cnnConnection.Execute strSql
            End If
            '2019/7/24 END
         End If
         'Add By Cheng 2002/11/01
         If pa(1) = "P" And Me.Text3.Text <> "" Then
            '92.5.16 MODIFY BY SONIA
            '2005/4/28 MODIFY BY SONIA
            'If m_CP10 = "101" Or m_CP10 = "102" Or m_CP10 = "103" Or m_CP10 = "104" Then
            If m_CP10 = "101" Or m_CP10 = "102" Or m_CP10 = "103" Or m_CP10 = "104" Or m_CP10 = "109" Or m_CP10 = "110" Or m_CP10 = "112" Then
            '2005/4/28 END
               strTxt(i) = "Update Patent Set PA10=" & DBDATE(Me.Text3.Text) & ",PA11='" & Me.Text1.Text & "' Where " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
            Else
               strTxt(i) = "Update Patent Set PA11='" & Me.Text1.Text & "' Where " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
            End If
            '92.5.16 END
            'Add By Cheng 2002/11/06
            cnnConnection.Execute strTxt(i)
            i = 4
         ElseIf pa(1) = "PS" And Me.Text3.Text <> "" Then
            strTxt(i) = "Update ServicePractice Set SP11='" & Me.Text1.Text & "' Where " & ChgService(pa(1) & pa(2) & pa(3) & pa(4))
            'Add By Cheng 2002/11/06
            cnnConnection.Execute strTxt(i)
            i = 4
         End If
         '2005/12/19 ADD BY SONIA 更新相同本所案號之相同代理人的彼所案號，若是彼所案號空的話
         'Modified by Morgan 2012/2/15 取消 cp09<'C' 條件(C類也會有發文作業,有代理人就要更新彼號,資料才會一致)
         strTxt(i) = "update caseprogress set cp45=" & CNULL(ChgSQL(Text4)) & " where cp09 in (select cp09 from caseprogress where rtrim(cp45) is null and " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND cp44 in (select cp44 from caseprogress where cp09='" & strReceiveNo & "' ))"
         cnnConnection.Execute strTxt(i)
         i = 5
         '2005/12/19 END
         
         'Added by Morgan 2016/5/11
         '有印客戶通知函時需新增已提申來函
         If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
            bolReceipt = False
            '提申有副本或要通知客戶時新增(管制文件或通知函)
            If frm04010507_1.Text3 = "2" And (Text20 = "Y" Or Text5 <> "N") Then
               stCP121 = ""
               'Modified by Morgan 2016/6/30 年費提申有副本時不必檢查altr --玲玲
               'Modified by Morgan 2016/7/6 領證、年費提申都不必檢查altr --玲玲
               If m_CP10 = "601" Or m_CP10 = "605" Then
                  stCP121 = "Y"
                  '年費有收據
                  'Modified by Morgan 2016/7/26 領證有副本也要收據
                  'If Text20 = "Y" And m_CP10 = "605" Then
                  If Text20 = "Y" Then
                     bolReceipt = True
                     iLP02 = 1
                  Else
                     iLP02 = 0
                  End If
                  
               'Add By Sindy 2019/12/11 代理人撰稿已直接存整封信件,不須另外檢查檔案
               ElseIf m_CP10 = "958" Then
                  stCP121 = "Y"
                  iLP02 = 0
               '2019/12/11 END
               
               'Removed by Morgan 2016/7/27 有副本都設兩個pdf--玲玲
               ''不出定稿只存 altr
               'ElseIf Text5 = "N" Then
               '   iLP02 = 1
               'end 2016/7/27
               ElseIf Text20 = "Y" Then
                  iLP02 = 2
               Else
                  iLP02 = 1
               End If
               
               '新增C類收文
               m_NewCP09 = AutoNo("C", 6)
               strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp20,cp26," & _
                  "cp32,cp27,cp43,cp121,cp145) values('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'" & _
                  "," & strSrvDate(1) & ",'" & m_NewCP09 & "','1909','" & m_CP12 & "','" & m_CP13 & "','" & strUserNum & "','N','N','N'," & strSrvDate(1) & _
                  ",'" & strReceiveNo & "','" & stCP121 & "','" & Text20 & "')"
               cnnConnection.Execute strSql, intI
               '新增信函進度
               'Modified by Morgan 2018/8/1
               'strExc(1) = PUB_GetLetterJudge(pa(1), "1909", , pa(9), pa(1), pa(2), pa(3), pa(4))
               strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), "1909", pa(9), , , IIf(Left(m_CP12, 1) = "F", True, False))
               PUB_AddLetterProgress m_NewCP09, iLP02, IIf(Text5 = "N", False, True), strExc(1), False, pa(26), "1909", pa(75), bolReceipt
            End If
         End If
         'end 2016/5/11
         
'2009/9/7 CANCEL BY SONIA 陳玲玲提出,P-085360
'         'Add By Cheng 2003/04/16
'         '若為大陸案, 且已提申陳述意見
'         If pa(9) = 大陸國家代號 And m_CP10 = "205" And frm04010507_1.Text3.Text = "2" Then
'            '新增C類收文
'...
'
'2009/9/7 END
         
'Remove by Morgan 2009/10/1 改輸核准
'         '92.11.3 MODIFY BY SONIA
'         '若為大陸案, 且已提申新穎性調查或申請檢索報告自動收文901告代
'         If pa(9) = 大陸國家代號 And (m_CP10 = "426" Or m_CP10 = "421") And frm04010507_1.Text3.Text = "2" Then
'            '新增B類收文
'         'Add by Morgan 2006/6/21 申請香港檢索報告已提申自動收文902回代,承辦掛操作人
'            'Add by Morgan 2007/1/17 改短期專利發文時會先掛2個月期限
'...
'
'end 2009/10/1

         'Add By Cheng 2002/11/05
         'Modified by Lydia 2021/01/20 改成TextBox
         'frm04010507_1.lblBillNo.Caption = ""
         frm04010507_1.txtBillno.Text = ""
         'Add By Cheng 2002/10/31
         '若有輸入代理人D/N No, 帳單日期 及 帳單金額, 則新增國外帳單資料
         If Me.Text11.Text <> "" And Me.Text12.Text <> "" And Me.Text13.Text <> "" Then
            'Modify by Morgan 2008/5/13 +傳幣別
            'If PUB_AddNewFBillData(strReceiveNo, Me.Text11.Text, Me.Text12.Text, Me.Text13.Text, strBillNo) = False Then
            If PUB_AddNewFBillData(strReceiveNo, Me.Text11.Text, Me.Text12.Text, Me.Text13.Text, strBillNo, Combo2.Text) = False Then
               'Modified by Morgan 2016/6/30 錯誤訊息不可放在 Transaction 內
               'MsgBox "新增國外帳單資料作業失敗!!!", vbExclamation + vbOKOnly
               strErrMsg = "新增國外帳單資料作業失敗!!!"
               'end 2016/6/30
               GoTo ErrorHandler
            Else
            
               'Added by Morgan 2016/6/30 非臺灣案電子化
               'Removed by Morgan 2025/8/13 帳單已全部都電子化
               'If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
               'end 2025/8/13
                  '檢查帳單是否存在
                  If PUB_CheckInvoicePDF(pa(1), pa(2), pa(3), pa(4), m_CP10, strErrMsg, , True, strBillNo) = False Then
                     GoTo ErrorHandler
                  End If
               'End If
               'end 2016/6/30
               
               'add by sonia 2025/2/8 P案809提第三方意見有帳單時自動上可結餘不必詢問
               If m_CP10 = "809" And pa(9) <> "000" Then
                  bolEndModCash = True
                  Pub_UpdateEndModCash pa(1), pa(2), pa(3), pa(4)
               'add by sonia 2025/8/13 第五年，第10年，第15年年費605有帳單時不必詢問直接上可結餘日
               ElseIf m_CP10 = "605" And pa(9) <> "000" And ((Val(m_CP53) <= 5 And Val(m_CP54) >= 5) Or (Val(m_CP53) <= 10 And Val(m_CP54) >= 10) Or (Val(m_CP53) <= 15 And Val(m_CP54) >= 15)) Then
                  bolEndModCash = True
                  Pub_UpdateEndModCash pa(1), pa(2), pa(3), pa(4)
               'end 2025/8/13
               End If
               'end 2025/2/8
            
               'Add By Cheng 2002/11/05
               'Modified by Lydia 2021/01/20 改成TextBox
               'frm04010507_1.lblBillNo.Caption = "" & strBillNo
               frm04010507_1.txtBillno.Text = "" & strBillNo
            End If
         End If
         
         'Add by Sindy 2016/10/7
         If m_strIR01 <> "" Then
            'Add by Sindy 2018/9/7 信件自動歸至卷宗區
            '已提申(領證,年費除外),是否收到副本=Y
            'Add By Sindy 2018/9/17 請設定以下案件性質輸入提申日時,要自動存入整封郵件
            '203 主動補正
            '204 補正
            '205 陳述意見 Sindy 2018/12/3 修改215=>205
            '206 補充說明
            '408 口審
            '501 訴願
            '503 行政訴訟
            '803 無效宣告
            '804 無效宣告答辯
            '906 異同分析
'            If frm04010507_1.Text3.Text = "2" And _
'               Not (m_CP10 = "601" Or m_CP10 = "605") And Text20 = "Y" And m_NewCP09 <> "" Then
            'Modify By Sindy 2018/12/3 + 107.復審 506.參加訴訟
            'Modified by Morgan 2019/2/27 +沒有副本也要歸卷--玲玲
            'If frm04010507_1.Text3.Text = "2" And _
               (m_CP10 = "203" Or m_CP10 = "204" Or m_CP10 = "205" Or m_CP10 = "206" Or _
                m_CP10 = "408" Or m_CP10 = "501" Or m_CP10 = "503" Or m_CP10 = "803" Or _
                m_CP10 = "804" Or m_CP10 = "906" Or m_CP10 = "107" Or m_CP10 = "506") And _
               Text20 = "Y" And m_NewCP09 <> "" Then
            'Modify By Sindy 2019/3/13 + 111 標準專利批准記錄請求
            '                          + 207 提供前案資料
            'Modified by Morgan 2019/3/14 +沒有副本也要歸卷,另111,207取消歸卷--玲玲
            'Modify By Sindy 2019/12/3 + m_CP10 = "958",因發信要夾帶信件所以一定都要存msg電子檔
            If frm04010507_1.Text3.Text = "2" And _
               (Text20 <> "Y" Or _
               (m_CP10 = "203" Or m_CP10 = "204" Or m_CP10 = "205" Or m_CP10 = "206" Or _
                m_CP10 = "408" Or m_CP10 = "501" Or m_CP10 = "503" Or m_CP10 = "803" Or _
                m_CP10 = "804" Or m_CP10 = "906" Or m_CP10 = "107" Or m_CP10 = "506" Or _
                m_CP10 = "958")) Then
            'end 2019/2/27
               '下載信件檔,上傳卷宗區(外來郵件)
               'Modify By Sindy 2018/12/3 RX.外來郵件 改 PAT.陸代郵件
               'Modified by Morgan 2019/2/27
               'If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, m_NewCP09, "PAT") = False Then
               'Modify By Sindy 2022/11/9 + IIf(pa(9) <> 台灣國家代號, "PAT", "RX")
               If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, IIf(m_NewCP09 = "", strReceiveNo, m_NewCP09), IIf(Pub_StrUserSt03 = "F22", "ALTR", IIf(pa(9) <> 台灣國家代號, "PAT", "RX")), strFullFileName) = False Then
               'end 2019/2/27
                  GoTo ErrorHandler
               End If
            End If
            '2018/9/7 END
            
            'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22"...
            PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm04010507_1", IIf(Pub_StrUserSt03 = "F22", IIf(m_NewCP09 = "", strReceiveNo, m_NewCP09), "")
         End If
         '2016/10/7 END
         
         'Added by Morgan 2025/4/11
         '新增 FMP 陳述意見(205)與補正(204) 收到副本要通知工程師及程序--品薇
         'Modified by Morgan 2025/6/24 +復審申請(107)、主動補正(203)--品薇/OWen
         'Modified by Morgan 2025/7/31 +PPH(431)--品薇
         If m_bFMP And Text20 = "Y" And Left(Pub_StrUserSt03, 1) <> "F" Then
            strExc(0) = "": strExc(1) = "": strExc(2) = ""
            If (m_CP10 = "205" Or m_CP10 = "204" Or m_CP10 = "107" Or m_CP10 = "203" Or m_CP10 = "431") Then
               strExc(0) = m_CP14
               'Modified by Morgan 2025/4/14 不用通知程序(工程師寫好請款信後才有程序的事)--敏莉
               'strExc(1) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
               'Modified by Morgan 2025/4/15 +CC給工程師主管--品薇/OWen
               strExc(1) = PUB_GetFCPEngSup(m_CP14)
               'end 2025/4/14
               strExc(2) = pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & "案已提申，且有副本，請工程師報告及請款"
               
            'Added by Morgan 2025/10/16
            '針對年費(605)、領證與繳年費(601)、實體審查(416)、讓與(701)及變更(401)這五種程序，key in時若副本上Y，系統自動發信告知承辦，主旨為「P108724已提申且有收據，請續行後續程序。」--品薇
            ElseIf (m_CP10 = "605" Or m_CP10 = "601" Or m_CP10 = "416" Or m_CP10 = "701" Or m_CP10 = "401") Then
               strExc(0) = m_CP13
               strExc(2) = pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & "案已提申且有收據，請續行後續程序。"
            'end 2025/10/16
            End If
            If strExc(0) <> "" And strExc(2) <> "" Then
               ClsPDGetCustomerNameAndAddress pa(26), oCustName
               oContext = "本所案號：" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & vbCrLf & vbCrLf & "專利名稱：" & Combo1 & vbCrLf & vbCrLf & "申請人　：" & oCustName & vbCrLf & vbCrLf & "案件性質：" & Label2(3) & vbCrLf & vbCrLf & "提申日　：" & ChangeWStringToWDateString(DBDATE(Text3))
               'Modified by Morgan 2025/6/27 +傳收文號以回存寄件備份及卷宗區
               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13,mc14)" & _
                  " values('" & strUserNum & "','" & strExc(0) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                  ",'" & ChgSQL(strExc(2)) & "','" & ChgSQL(oContext) & "','" & strExc(1) & "','" & m_NewCP09 & "','Y')"
               cnnConnection.Execute strSql, intI
            End If
         End If
         'end 2025/4/11
         
         cnnConnection.CommitTrans
         
         'Add By Sindy 2019/7/24 已提申代理人撰稿發E-Mail通知工程師
         '要注意上列電子檔是否有上傳
         If frm04010507_1.Text3.Text = "2" And m_CP10 = "958" Then
            strExc(0) = "select cp14 from caseprogress where cp09=(select cp43 from caseprogress where cp09='" & strReceiveNo & "' and cp43 is not null)"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If ClsPDGetCustomerNameAndAddress(pa(26), oCustName) Then
               End If
               oContext = "本所案號：" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & vbCrLf & vbCrLf & "專利名稱：" & Combo1 & vbCrLf & vbCrLf & "申請人　：" & oCustName & vbCrLf & vbCrLf & "案件性質：" & Label2(3) & vbCrLf & vbCrLf & "提申日　：" & ChangeWStringToWDateString(DBDATE(Text3))
               'Modified by Morgan 2025/6/27 +傳收文號以回存寄件備份及卷宗區
               PUB_SendMail strUserNum, RsTemp.Fields("cp14"), strReceiveNo, "通知 " & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & "「" & Combo1 & "」 " & Label2(3) & "已完成！", oContext, , strFullFileName, , , , , , , , , , , , , , , , , , , , , , , , , strReceiveNo
            End If
         End If
         '2019/7/24 END
         
         'Add by Sindy 2018/1/2
         If m_strIR01 <> "" And strBillNo <> "" Then
            MsgBox "已新增帳單【 " & strBillNo & " 】。", vbInformation
         End If
         '2018/1/2 END
         
         '已提申才有定稿
         If frm04010507_1.Text3 = "2" Then
            If Text5 <> "N" Then
               '預設
               strTmp = "00"
               Select Case m_CP10
                  Case 面詢 '1
                     strTmp = "01"
                  Case 申復 '2(陳述意見)
                     '若有勾選陳述意見書
                     If Me.chk(0).Value Then
                         strTmp = "02"
                     '若未勾選陳述意見書
                     Else
                         strTmp = "03"
                     End If
                     
                     'Added by Morgan 2021/5/28 寶齡富錦 Y55435 案件
                     '不管附件都帶相同內容--玲玲
                     If pa(75) = "Y55435" Then
                        strTmp = "99"
                     End If
                     'end 2021/5/28
                        
                  'Add By Cheng 2002/11/01
                  '修正(補正), 主動修正(主動補正)
                  Case 修正, 主動修正
                     '若有勾選陳述意見書
                     If Me.chk(0).Value Then
                         strTmp = "02"
                     '若未勾選陳述意見書
                     Else
                         strTmp = "03"
                     End If
                  'Add by Morgan 2006/6/21
                  Case "421"
                     '申請香港檢索報告
                     If pa(9) = "013" Then
                        strTmp = "04"
                     End If
               End Select
               
               'Modify by Morgan 2010/2/4 +不出定稿控制
               If Not bolNoLetter Then
               
                  'Add By Cheng 2002/11/01
                  StartLetter "13", strTmp
                  
                  'Add by Morgan 2009/11/19 FMP的中文定稿只要1份
                  If Left(m_CP12, 1) = "F" Then
                     'Modified by Morgan 2025/4/11 FMP不再印紙本--品薇
                     NowPrint strReceiveNo, "13", strTmp, False, strUserNum, 0, , , , 1, , , , , , , , m_NewCP09, , , , , True
                  Else
                  'end 2009/11/19
                     NowPrint strReceiveNo, "13", strTmp, False, strUserNum, 0, , , , , , , , , , , , m_NewCP09
                  End If
                  
               End If
               
'Remove by Morgan 2009/10/1 改輸核准時出
'               'Add by Morgan 2006/7/10 香港短期專利補檢索報告指示信
'               If m_902CP09 <> Empty Then
'                  NowPrint m_902CP09, "02", "40", False, strUserNum, 0
'                  PUB_PrintLetter m_902CP09
'               End If
'end 2009/10/1

            End If
            
            'Add by Morgan 2009/11/19
            'Modify by Morgan 2010/9/8 FMP一律要出英文通知信(從上面移下來)--美珍
            If Left(m_CP12, 1) = "F" Then
               strUserNum = strFMPNum
               StartLetter1 "13", "51" 'Added by Morgan 2019/12/12
               NowPrint strReceiveNo, "13", "51", False, strUserNum, 0
               strUserNum = strUser1Num
               'Add by Lydia 2015/01/26 再次提申不發送mail
               If m_CP47 = "" Then
                    '2011/12/19 ADD BY SONIA 同時發MAIL通知收文智權人員
                    If ClsPDGetCustomerNameAndAddress(pa(26), oCustName) Then
                    End If
                    oContext = "本所案號：" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & vbCrLf & vbCrLf & "專利名稱：" & Combo1 & vbCrLf & vbCrLf & "申請人　：" & oCustName & vbCrLf & vbCrLf & "案件性質：" & Label2(3) & vbCrLf & vbCrLf & "提申日　：" & ChangeWStringToWDateString(DBDATE(Text3))
                    'Modified by Morgan 2021/1/28
                    'PUB_SendMail strUserNum, PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4)), strReceiveNo, "通知 " & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & " 的" & Label2(3) & "程序已提申！", oContext
                    'Modified by Morgan 2025/6/27 +傳收文號以回存寄件備份及卷宗區
                    PUB_SendMail strUserNum, m_CP13, strReceiveNo, "通知 " & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & " 的" & Label2(3) & "程序已提申！", oContext, , , , , , , , , , , , , , , , , , , , , , , , , , , strReceiveNo
                    'end 2021/1/28
                    '2011/12/19 END
                    
               End If
            End If
         End If
         
         'Added by Lydia 2016/03/10 FMP之領證且有輸入帳單金額者額,發E-MAIL給該案件之FCP承辦智權人員
         If m_bFMP And m_CP10 = "601" And Trim(Text13) <> "" Then
            '依FC代理人定稿語言取得名稱
            strExc(0) = "SELECT DECODE(FA31,'2',DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)," & _
                        "'3',NVL(FA06,NVL(FA04,FA05||' '||FA63||' '||FA64||' '||FA65))," & _
                        "NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65))),NA03 from fagent,nation where fa01='" & Mid(ChangeCustomerL(pa(75)), 1, 8) & "' and fa02='" & Mid(ChangeCustomerL(pa(75)), 9, 1) & "' and fa10=na01(+) "
            intI = 1: strExc(1) = "": strExc(2) = ""
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            strExc(1) = "" & RsTemp(0)
            strExc(2) = "" & RsTemp(1)
            oContext = "本所案號：" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & vbCrLf & _
                       "專利名稱：" & Combo1 & vbCrLf & _
                       "FC代理人：" & pa(75) & " " & strExc(1) & vbCrLf & _
                       "代理人國籍：" & strExc(2) & vbCrLf & _
                       "代理人D/N NO：" & Trim(Text11) & vbCrLf & _
                       "帳單日期：" & ChangeWStringToWDateString(DBDATE(Text12)) & vbCrLf & _
                       "幣　　別：" & Combo2 & vbCrLf & _
                       "帳單金額：" & Trim(Text13)
            'Modified by Morgan 2025/6/27 +傳收文號以回存寄件備份及卷宗區
            PUB_SendMail strUserNum, PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)), strReceiveNo, pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & " 的" & Label2(3) & "收到大陸代理人領證程序帳單.", oContext, , , , , , , , , , , , , , , , , , , , , , , , , , , strReceiveNo
         End If
         Screen.MousePointer = vbDefault
         
         'Add By Sindy 2016/10/7
         If Me.m_strIR01 <> "" Then
            Unload frm04010507_1
            Unload Me
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         Else
         '2016/10/7 END
            Unload Me
            frm04010507_1.Show
            ' 91.01.22 modify by louis (清除前一個畫面的欄位)
            frm04010507_1.Clear
            frm04010507_1.SetInputFocus
         End If
      Case 1
         Unload Me
         frm04010507_1.Show
   End Select
'Add By Cheng 2002/11/06
Exit Sub
ErrorHandler:
   cnnConnection.RollbackTrans
   
   If strErrMsg <> "" Then
      MsgBox strErrMsg, vbCritical
   Else
      MsgBox "存檔失敗，請洽系統管理人 !", vbCritical
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
ReDim Ncp(1 To TF_CP) As String
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Morgan 2025/4/11
   Set frm04010507_2 = Nothing
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   
   strReceiveNo = frm04010507_1.Tag
   pa(1) = strExc(1)
   pa(2) = strExc(2)
   pa(3) = strExc(3)
   pa(4) = strExc(4)
   ReadPatent
   
   'Add By Sindy 2017/12/27
   m_strIR01 = frm04010507_1.m_strIR01
   m_strIR02 = frm04010507_1.m_strIR02
   m_strIR03 = frm04010507_1.m_strIR03
   m_strIR04 = frm04010507_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
      If frm04010507_1.Text3 = "1" And Text2 = "" Then Text2 = DBDATE(frm04010507_1.m_RDate)
   Else
   '2017/12/27 END
      '2005/4/27 MODIFY BY SONIA 改為西元日期
      'If frm04010507_1.Text3 = "1" And Text2 = "" Then Text2 = strSrvDate(2)
      'If frm04010507_1.Text3 = "2" And Text3 = "" Then Text3 = strSrvDate(2)
      If frm04010507_1.Text3 = "1" And Text2 = "" Then Text2 = strSrvDate(1)
   End If
   
   'Remove by Morgan 2006/7/10 提申不預設
   'If frm04010507_1.Text3 = "2" And Text3 = "" Then Text3 = strSrvDate(1)
   'Add by Morgan 2006/10/5 提申且案件姓質選年費要預設
   'Removed by Morgan 2016/7/26 取消預設--陳玲玲
   'If frm04010507_1.Text3.Text = "2" And Text3.Text = "" And m_CP10 = "605" Then Text3.Text = strSrvDate(1)
   'end 2016/7/26
   
   '2005/4/27 END
    'Add By Cheng 2003/04/02
    '判斷案件性質
    Select Case m_CP10
    '2005/4/28 MODIFY BY SONIA
    'Case "101", "102", "103", "109"
    'Modify by Morgan 2009/10/1 +421
    'Case "101", "102", "103", "104", "109", "110", "112"
    Case "101", "102", "103", "104", "109", "110", "112", "421"
    
    '2005/4/28 END
        '若前畫面結果選已收達
        If frm04010507_1.Text3 = "1" Then
            'Modify by Morgan 2005/7/8 --敏惠
'            '游標停在彼所案號欄位
'            SendKeys "{Tab}{Tab}{Tab}{Tab}{Tab}{Tab}{Tab}"
            '游標停在代理人收達日欄位
            SendKeys "{Tab}{Tab}{Tab}{Tab}"
            '2005/7/8 end
        '若前畫面結果選已提申
        Else
            '2005/4/20 MODIFY BY SONI
            ''游標停在申請案號欄位
            'SendKeys "{Tab}{Tab}{Tab}{Tab}{Tab}{Tab}"
            '游標停在代理人提申日欄位
            SendKeys "{Tab}{Tab}{Tab}{Tab}{Tab}"
            '2005/4/20 END
        End If
        '不印客戶通知函
        
'        Me.Text5.Text = "N" 'Removed by Morgan 2015/10/19 改都預設不出定稿,副本上Y時才要(下面設定) --玲玲

    '2005/4/20 ADD BY SONIA
    Case Else
        '若前畫面結果選已收達
        If frm04010507_1.Text3 = "1" Then
            '游標停在代理人收達日欄位
            SendKeys "{Tab}{Tab}{Tab}{Tab}"
        
'Removed by Morgan 2015/10/19 改都預設不出定稿,副本上Y時才要(下面設定) --玲玲
'        Else
'            'Add by Morgan 2010/9/8 領證提申預設不印客戶通知函 --玲玲
'            Select Case m_CP10
'               'Modified by Morgan 2012/2/10 +203,204,205,107,803,804 也預設不印,已收副本上Y才改要印
'               'Case "601"
'               Case "601", "203", "204", "205", "107", "803", "804"
'                  Text5.Text = "N"
'
'            End Select
'end 2015/10/19

        End If
    '2005/4/20 END
    End Select
   'Add by Morgan 2004/1/30
   '抓代理人
   m_CP44 = GetCP44()
   'Add end---
   
   'Add by Morgan 2008/5/13
   If m_CP44 <> "" Then
      PUB_Add2Combo Combo2, m_CP44
   End If
   
   'Add by Morgan 2010/4/16 已收達不可輸提申日
   If frm04010507_1.Text3 = "1" Then
      EnableTextBox Text3, False
   End If
   
   '2011//2/8 add by sonia 玲玲說要鎖起來
   Text1.Locked = True
   Text1.Enabled = False
   '2011/2/8 end
   
   'Add by Morgan 2011/4/15
   'B類收文都預設不要出通知函--敏惠
   'Modified by Morgan 2015/10/19 改都預設不出定稿,副本上Y時才要 --玲玲
   'If Left(strReceiveNo, 1) = "B" Then
      Text5 = "N"
   'End If
   'end 2015/10/19
End Sub

Private Sub ReadPatent()
 Dim Lbl As Object, strTmp As String, i As Integer
 'Add By Cheng 2002/07/08
 Dim StrSQLa As String
 
   For Each Lbl In Label2
      Lbl.Caption = ""
   Next
   Label2(0) = MergeString(pa(1), pa(2), pa(3), pa(4))
   If pa(1) = "P" Then
      If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
         Label2(1) = pa(22)
        'Modify By Cheng 2002/11/01
'         Label2(2) = pa(11)
        Me.Text1.Text = "" & pa(11)
         If pa(9) <> "" Then
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetNation(pA(9), strTmp) Then
            If ClsPDGetNation(pa(9), strTmp) Then
               Label2(8) = strTmp
            End If
         End If
      End If
   ElseIf pa(1) = "PS" Then
      If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
        'Modify By Cheng 2002/11/01
'         Label2(2) = pa(11)
        Me.Text1.Text = "" & pa(11)
         If pa(9) <> "" Then
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetNation(pA(9), strTmp) Then
            If ClsPDGetNation(pa(9), strTmp) Then
               Label2(8) = strTmp
            End If
         End If
      End If
   End If
   
    'Modify By Cheng 2002/11/01
'   If pa(10) = 台灣國家代號 Then
   If pa(9) = 台灣國家代號 Then
      strTmp = "CPM03"
   Else
      strTmp = "CPM04"
   End If
   
   '92.3.21 add by sonia
   Label2(9) = ""
   If pa(72) <> "" Then Label2(9) = pa(72)
   '92.3.21 end
   
   'Modify By Cheng 2002/07/08
   '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
'   strExc(0) = "select " & strTmp & ",S1.ST02,S2.ST02," & SQLDate("CP27") & ",NVL(FA05,NVL(FA04,FA06)),CP45,NVL(CP44,''),CP10,CP46" & _
'      " FROM CASEPROGRESS,CASEPROPERTYMAP,STAFF S1,STAFF S2,FAGENT WHERE CP09='" & strReceiveNo & "' AND " & _
'      "CP01=cpm01(+) and cp10=cpm02(+) and SUBSTR(cp44,1,8)=FA01(+) and SUBSTR(cp44,9,1)=FA02(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+)"
   StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
    'Modify By Cheng 2003/04/16
'   strExc(0) = "select " & strTmp & ",S1.ST02,S2.ST02," & SQLDate("CP27") & "," & strSQLA & "CP45,NVL(CP44,''),CP10,CP46" & _
'      " FROM CASEPROGRESS,CASEPROPERTYMAP,STAFF S1,STAFF S2,FAGENT,SystemKind WHERE CP09='" & strReceiveNo & "' AND " & _
'      "CP01=cpm01(+) and cp10=cpm02(+) and SUBSTR(cp44,1,8)=FA01(+) and SUBSTR(cp44,9,1)=FA02(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) and CP01=SK01(+) "
   '92.11.3 MODIFY BY SONIA
   'strExc(0) = "select " & strTmp & ",S1.ST02,S2.ST02," & SQLDate("CP27") & "," & strSQLA & "CP45,NVL(CP44,''),CP10,CP46, CP12, CP13 " & _
   '   " FROM CASEPROGRESS,CASEPROPERTYMAP,STAFF S1,STAFF S2,FAGENT,SystemKind WHERE CP09='" & strReceiveNo & "' AND " & _
   '   "CP01=cpm01(+) and cp10=cpm02(+) and SUBSTR(cp44,1,8)=FA01(+) and SUBSTR(cp44,9,1)=FA02(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) and CP01=SK01(+) "
   '2005/4/20 MODIFY BY SONIA 加入CP47
   'strExc(0) = "select " & strTmp & ",S1.ST02,S2.ST02," & SQLDate("CP27") & "," & strSQLA & "CP45,NVL(CP44,''),CP10,CP46, CP12, CP13, CP14 " & _
   '   " FROM CASEPROGRESS,CASEPROPERTYMAP,STAFF S1,STAFF S2,FAGENT,SystemKind WHERE CP09='" & strReceiveNo & "' AND " & _
   '   "CP01=cpm01(+) and cp10=cpm02(+) and SUBSTR(cp44,1,8)=FA01(+) and SUBSTR(cp44,9,1)=FA02(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) and CP01=SK01(+) "
   'Modified by Morgan 2012/2/10 +CP145
   'Modified by Morgan 2012/3/27 +CP64
   'modify by sonia 2025/8/13 +CP53,CP54
   strExc(0) = "select " & strTmp & ",S1.ST02,S2.ST02," & SQLDate("CP27") & "," & StrSQLa & "CP45,NVL(CP44,''),CP10,CP46, CP12, CP13, CP14, CP47,CP07,CP145,CP64,CP53,CP54 " & _
      " FROM CASEPROGRESS,CASEPROPERTYMAP,STAFF S1,STAFF S2,FAGENT,SystemKind WHERE CP09='" & strReceiveNo & "' AND " & _
      "CP01=cpm01(+) and cp10=cpm02(+) and SUBSTR(cp44,1,8)=FA01(+) and SUBSTR(cp44,9,1)=FA02(+) AND CP13=S1.ST01(+) AND CP14=S2.ST01(+) and CP01=SK01(+) "
   '2005/4/20 END
   '92.11.3 END
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Label2(2) = "" & RsTemp.Fields("cp64") 'Added by Morgan 2012/3/27
      m_CP145 = "" & RsTemp.Fields("cp145") 'Added by Morgan 2012/2/10
      Text20 = m_CP145 'Added by Morgan 2012/2/10
      'Add by Lydia 2015/01/26
        m_CP47 = ""
        If Not IsNull(RsTemp!cp47) Then m_CP47 = RsTemp!cp47
      'end 2015/01/26
      For i = 0 To 4
         If Not IsNull(RsTemp.Fields(i)) Then Label2(i + 3) = RsTemp.Fields(i)
      Next
      'Add by Morgan 2009/11/18
      m_CP07 = "" & RsTemp.Fields("CP07")
      If m_CP07 <> "" Then
         Label2(10) = ChangeWStringToTDateString(m_CP07)
      End If
      'end 2009/11/18
      If Not IsNull(RsTemp.Fields(8)) Then
         '2005/4/27 MODIFY BY SONIA 改為西元日期
         'Text2 = ChangeWStringToTString(rsTemp.Fields(8))
         Text2 = RsTemp.Fields(8)
         '2005/4/27 END
      Else
         Text2 = ""
      End If
      '2005/4/20 ADD BY SONIA
      If Not IsNull(RsTemp.Fields(12)) Then
         '2005/4/27 MODIFY BY SONIA 改為西元日期
         'Text3 = ChangeWStringToTString(rsTemp.Fields(12))
         Text3 = RsTemp.Fields(12)
         '2005/4/27 END
      Else
         Text3 = ""
      End If
      '2005/4/20 END
      m_CP10 = "" & RsTemp.Fields(7)
        'Add By Cheng 2003/04/16
        'Modified by Morgan 2017/6/20
        ''業務區
        'm_CP12 = "" & RsTemp.Fields("CP12").Value
        ''智權人員
        'm_CP13 = "" & RsTemp.Fields("CP13").Value
        m_CP13 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
        m_CP12 = GetSalesArea(m_CP13)
        'end 2017/6/20
        If Left(m_CP12, 1) = "F" Then m_bFMP = True 'Add by Morgan 2010/12/27
        
        '92.11.3 ADD BY SONIA
        '承辦人
        m_CP14 = "" & RsTemp.Fields("CP14").Value
        '92.11.3 END
        'add by sonia 2025/8/13
        m_CP53 = "" & RsTemp.Fields("CP53").Value
        m_CP54 = "" & RsTemp.Fields("CP54").Value
        'end 2025/8/13
      If Not IsNull(RsTemp.Fields(5)) Then
         strExc(0) = "SELECT CP45 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
            " AND CP44=" & CNULL("" & RsTemp.Fields(6)) & " AND CP45 IS NOT NULL"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(RsTemp.Fields(0)) Then Text4 = RsTemp.Fields(0)
         End If
      End If
   End If
    'Add By Cheng 2002/11/01
    If m_CP10 = 修正 Or m_CP10 = 主動修正 Then
       frm.Visible = True
       Me.chk(0).Caption = "補正書"
    End If
    'Add By Cheng 2002/12/12
    '大陸案已提申時, 案件性質為陳述意見(205)
    'Modified by Morgan 2021/5/11 改定稿文件選項說明 意見陳述書-->陳述意見書 --郭
    If pa(9) = 大陸國家代號 And frm04010507_1.Text3 = "2" And m_CP10 = "205" Then frm.Visible = True
End Sub

Private Sub Text1_GotFocus()
    'Add By Cheng 2002/11/01
    TextInverse Me.Text1
End Sub

'2005/6/14 ADD BY SONIA
Private Sub Text1_Validate(Cancel As Boolean)
Dim i As Integer
   i = 2
   If pa(9) = 台灣國家代號 Then
      i = 0
   Else
      If pa(9) = 大陸國家代號 And pa(46) <> "Y" Then
         i = 1
      End If
   End If
   If Text1 <> "" And (i = 0 Or i = 1) Then
      If Not ChkAppNo(Text1, Val(pa(8)), i, Val(pa(23))) Then Cancel = True
   End If
End Sub
'2005/6/14 END

Private Sub Text11_GotFocus()
    TextInverse Me.Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text12_GotFocus()
    TextInverse Me.Text12
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   If Me.Text12.Text <> "" Then
      'Modify by Morgan 2005/10/18統一控制西元年
      'If ChkDate(Me.Text12.Text) = False Then
      If CheckIsDate(Me.Text12.Text) = False Then
         Cancel = True
         TextInverse Me.Text12
      'Add by Morgan 2006/4/25 檢查不可大於系統日
      ElseIf Val(Text12) > Val(strSrvDate(1)) Then
         MsgBox "帳單日期不可大於系統日！", vbExclamation
         Cancel = True
      End If
   End If
End Sub

Private Sub Text13_GotFocus()
    TextInverse Me.Text13
End Sub

Private Sub Text13_Validate(Cancel As Boolean)
    If Me.Text13.Text <> "" Then
        If IsNumeric(Me.Text13.Text) = False Then
            MsgBox "帳單金額輸入錯誤!!!", vbExclamation + vbOKOnly
            Cancel = True
            TextInverse Me.Text13
        'Add by Morgan 2004/1/30
        ElseIf Val(Text13) <> 0 Then
            If m_CP44 = "" Then
                MsgBox "該筆進度資料無代理人，不可輸入帳單!!!", vbExclamation + vbOKOnly
                Cancel = True
                TextInverse Text13
            End If
        'Add end ---------------
        End If
    End If
End Sub
'Add by Morgan 2004/1/30
Private Function GetCP44() As String
    Dim StrSQLa As String
    Dim rsA As New ADODB.Recordset
'on error GoTo ErrHnd
    StrSQLa = "Select CP44 From CaseProgress Where CP09='" & strReceiveNo & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        GetCP44 = "" & rsA.Fields(0).Value
    Else
        GetCP44 = ""
    End If
ErrHnd:
    If Err.NUMBER <> 0 Then
        MsgBox Err.Description
    End If
End Function

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   'Modify by Morgan 2005/10/18統一控制西元年
   'If Text2 <> "" Then If Not ChkDate(Text2.Text) Then Text2.SetFocus
   If Text2 <> "" Then If Not CheckIsDate(Text2.Text) Then Cancel = True
End Sub


'Added by Morgan 2012/2/10
Private Sub Text20_GotFocus()
   TextInverse Text20
   CloseIme
End Sub

'Added by Morgan 2012/2/10
Private Sub Text20_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   ElseIf KeyAscii = Asc("Y") Then
      'Modified by Morgan 2014/3/7 B類205(陳述意見),204(補正)不預設通知定稿
      'Text5 = ""
      'Modified by Morgan 2015/10/19 B類都不出定稿--玲玲
      'If Not (Left(strReceiveNo, 1) = "B" And (m_CP10 = "205" Or m_CP10 = "204")) Then
      If Left(strReceiveNo, 1) <> "B" Then
      'end 2015/10/19
         'Modified by Morgan 2016/10/13 非FMP的領證預設不出定稿Ex:P114331--玲玲
         'Text5 = ""
         If Not (m_CP10 = "601" And m_bFMP <> True) Then
            Text5 = ""
         End If
         'end 2016/10/13
      End If
      'end 2014/3/7
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub
'Modify by Morgan 2005/10/18 LostFocus改Validate事件
Private Sub Text3_Validate(Cancel As Boolean)
   If frm04010507_1.Text3 = "1" Then
      If Text2 = "" And Text3 = "" Then
         MsgBox "代理人收達日及提申日不可同時空白 !", vbCritical
         Cancel = True
      Else
         'Modify by Morgan 2005/10/18統一控制西元年
         'If Text3 <> "" Then If Not ChkDate(Text3.Text) Then Text3.SetFocus
         If Text3 <> "" Then If Not CheckIsDate(Text3.Text) Then Cancel = True
      End If
   ElseIf frm04010507_1.Text3 = "2" Then
      If Text3 = "" Then
         MsgBox "代理人提申日不可空白 !", vbCritical
         Cancel = True
      Else
         'Modify by Morgan 2005/10/18統一控制西元年
         'If Text3 <> "" Then If Not ChkDate(Text3.Text) Then Text3.SetFocus
         If Text3 <> "" Then If Not CheckIsDate(Text3.Text) Then Cancel = True
      End If
   End If
End Sub

Private Sub Text4_GotFocus()
  TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2002/11/22
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Add By Cheng 2002/11/21
Private Function GetPayYear() As String
    Dim ArrYear
    Dim arrDay
    Dim ii As Integer
    Dim intBegin As Integer
    Dim intEnd As Integer
        
    GetPayYear = ""
    intBegin = 0
    intEnd = 0
    '若繳年費日期, 年度及發文日有資料
    If "" & pa(73) <> "" And pa(72) <> "" And Me.Label2(6).Caption <> "" Then
        ArrYear = Split(pa(72), ",")
        arrDay = Split(pa(73), ",")
        '若陣列數相同
        If UBound(ArrYear) = UBound(arrDay) Then
            For ii = LBound(arrDay) To UBound(arrDay)
                '91.11.27 MODIFY BY SONIA 字串與日期永遠不會相等
                'If arrDay(ii) = DBDATE(Me.Label2(6).Caption) Then
                '判斷同一個發文日繳幾年度
                If Val(arrDay(ii)) = DBDATE(Me.Label2(6).Caption) Then
                '91.11.27 END
                    If intBegin = 0 Then intBegin = ii
                    intEnd = ii
                End If
            Next ii
            'Add By Cheng 2003/01/10
            '若無繳費日與發文日相同時, 則直接抓繳費年度最後一年
            If intEnd = 0 Then
                intEnd = Val(ArrYear(UBound(ArrYear)))
                intBegin = intEnd
            '若有繳費日與發文日相同者, 則抓相同日的相對應的繳費年度
            Else
                intBegin = Val(ArrYear(intBegin))
                intEnd = Val(ArrYear(intEnd))
            End If
        End If
    End If
If intBegin = intEnd Then
   '2008/11/12 MODIFY BY SONIA P-074053陳玲玲說香港U,D改為續期費且不通知第幾次
   'GetPayYear = "第" & PUB_ChgNumber2Chinese("" & intBegin) & "年年費"
   If pa(9) = "013" And pa(8) <> "1" Then
      GetPayYear = "續期費"
   Else
      GetPayYear = "第" & PUB_ChgNumber2Chinese("" & intBegin) & "年年費"
   End If
   '2008/11/12 END
Else
    GetPayYear = "第" & PUB_ChgNumber2Chinese("" & intBegin) & "年至第" & PUB_ChgNumber2Chinese("" & intEnd) & "年年費"
End If
End Function

'Add By Cheng 2003/04/02
'取得基本檔最後已繳年度
Private Sub GetLastYearFee(strPA01 As String, strPA02 As String, strPA03 As String, strPA04 As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strArrYearFee
Dim ii As Integer
    
m_LastYearFee = ""
StrSQLa = "Select PA72 From Patent Where " & ChgPatent(strPA01 & strPA02 & strPA03 & strPA04)
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    If Trim("" & rsA.Fields(0).Value) <> "" Then
        strArrYearFee = Split("" & rsA.Fields(0).Value, ",")
        For ii = LBound(strArrYearFee) To UBound(strArrYearFee)
            If Val("0" & strArrYearFee(ii)) > Val("0" & m_LastYearFee) Then
                m_LastYearFee = strArrYearFee(ii)
            End If
        Next ii
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Sub

Private Sub Text6_GotFocus()
    TextInverse Me.Text6
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   Text12_Validate Cancel
   If Cancel = True Then
      Text12_GotFocus
      Text12.SetFocus
      Exit Function
   End If
   Text13_Validate Cancel
   If Cancel = True Then
      Text13_GotFocus
      Text13.SetFocus
      Exit Function
   End If
   Text2_Validate Cancel
   If Cancel = True Then
      Text2_GotFocus
      Text2.SetFocus
      Exit Function
   End If
   Text3_Validate Cancel
   If Cancel = True Then
      Text3_GotFocus
      Text3.SetFocus
      Exit Function
   End If
   'Added by Morgan 2012/5/3
   If Text3 <> "" And Text2 <> "" Then
      If DBDATE(Text3) < DBDATE(Text2) Then
         MsgBox "提申日不可早於收達日！", vbExclamation
         Text3.SetFocus
         Text3_GotFocus
         Exit Function
      End If
   End If
   'end 2012/5/3
   
   'Modify By Sindy 2017/12/22 Move到cmkok
'   'Added by Morgan 2016/7/6
'   '已收達要檢查有檔案
'   If frm04010507_1.Text3.Text = "1" And Left(Pub_StrUserSt03, 1) <> "F" Then
'      If PUB_CheckAck(strReceiveNo, pa(1), pa(2), pa(3), pa(4), m_CP10) = False Then
'         MsgBox "找不到代理人已收達電子檔(ACK)!!", vbExclamation
'         Exit Function
'      End If
'   End If
'   'end 2016/7/6
   
'   'Added by Morgan 2016/7/22
'   '已提申自動匯入代理人來函電子檔
'   If frm04010507_1.Text3.Text = "2" And Left(Pub_StrUserSt03, 1) <> "F" And (Text20 = "Y" Or Text5 <> "N") Then
'      If PUB_CheckAltr(pa(1), pa(2), pa(3), pa(4), "1909") = False Then
'         MsgBox "找不到代理人來函電子檔(ALTR)!!", vbExclamation
'         Exit Function
'      End If
'   End If
'   'end 2016/7/22
   
   TxtValidate = True
End Function

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   If Combo2 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo2, Label25(4)) = False Then
      Cancel = True
      Combo2.SetFocus
   End If
End Sub
'Added by Morgan 2019/12/12
Private Sub StartLetter1(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt(1 To 5) As String
   Dim ii As Integer
   Dim stPS As String
    
   ii = 1
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   '請款函備註
   stPS = PUB_GetDebitNotePS(pa(1) & pa(2) & pa(3) & pa(4), m_CP10, ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)))
   If stPS <> "" Then
      stPS = "P.S. " & stPS
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','有請款函備註時不印','♀')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','請款函備註','" & ChgSQL(stPS) & "')"
      ii = ii + 1
   End If
   
   If ii <> 1 Then
       If Not ClsLawExecSQL(ii - 1, strTxt) Then
          MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
       End If
   End If
End Sub


