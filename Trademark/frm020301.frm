VERSION 5.00
Begin VB.Form frm020301 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權人員期限管制表"
   ClientHeight    =   5790
   ClientLeft      =   220
   ClientTop       =   370
   ClientWidth     =   6130
   ControlBox      =   0   'False
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   6130
   Begin VB.TextBox textTM02_2 
      Height          =   264
      Left            =   2910
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2250
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textTM01 
      Height          =   264
      Left            =   1470
      MaxLength       =   3
      TabIndex        =   12
      Top             =   2250
      Width           =   732
   End
   Begin VB.TextBox textTM03 
      Height          =   264
      Left            =   3270
      MaxLength       =   1
      TabIndex        =   15
      Top             =   2250
      Width           =   372
   End
   Begin VB.TextBox textTM04 
      Height          =   264
      Left            =   3630
      MaxLength       =   2
      TabIndex        =   16
      Top             =   2250
      Width           =   732
   End
   Begin VB.TextBox textMoney 
      Height          =   264
      Left            =   2040
      MaxLength       =   8
      TabIndex        =   30
      Top             =   5370
      Width           =   1155
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   17
      Left            =   3780
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1350
      Width           =   1005
   End
   Begin VB.FileListBox File1 
      Height          =   180
      Left            =   0
      TabIndex        =   62
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CheckBox ChkMail 
      Caption         =   "寄發管制表給智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3600
      TabIndex        =   21
      Top             =   3150
      Width           =   2385
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   22
      Left            =   2955
      MaxLength       =   7
      TabIndex        =   11
      Top             =   1935
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   21
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   10
      Top             =   1935
      Width           =   1005
   End
   Begin VB.OptionButton Option1 
      Caption         =   "大陸撤三期限："
      Height          =   180
      Index           =   1
      Left            =   110
      TabIndex        =   9
      Top             =   1980
      Width           =   2000
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所期限："
      Height          =   180
      Index           =   0
      Left            =   110
      TabIndex        =   6
      Top             =   1680
      Value           =   -1  'True
      Width           =   1245
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   20
      Left            =   2055
      MaxLength       =   1
      TabIndex        =   29
      Top             =   4590
      Width           =   285
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   19
      Left            =   5640
      MaxLength       =   1
      TabIndex        =   32
      Top             =   4830
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   16
      Left            =   1860
      MaxLength       =   1
      TabIndex        =   22
      Text            =   "1"
      Top             =   3420
      Width           =   285
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   18
      Left            =   4590
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1050
      Width           =   315
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   15
      Left            =   1455
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1350
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定地址條"
      Height          =   540
      Left            =   180
      TabIndex        =   48
      Top             =   5850
      Width           =   3585
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         TabIndex        =   33
         Text            =   "Combo1"
         Top             =   240
         Width           =   2580
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   49
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   5790
      MaxLength       =   1
      TabIndex        =   31
      Top             =   4560
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   5040
      TabIndex        =   35
      Top             =   150
      Width           =   756
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4200
      TabIndex        =   34
      Top             =   150
      Width           =   756
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   14
      Left            =   2745
      MaxLength       =   4
      TabIndex        =   28
      Top             =   4280
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   13
      Left            =   1455
      MaxLength       =   4
      TabIndex        =   27
      Top             =   4280
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   12
      Left            =   2745
      MaxLength       =   9
      TabIndex        =   26
      Top             =   3990
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   11
      Left            =   1455
      MaxLength       =   9
      TabIndex        =   25
      Top             =   3990
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   10
      Left            =   2745
      MaxLength       =   9
      TabIndex        =   24
      Top             =   3710
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   9
      Left            =   1455
      MaxLength       =   9
      TabIndex        =   23
      Top             =   3710
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   8
      Left            =   1455
      MaxLength       =   1
      TabIndex        =   20
      Top             =   3140
      Width           =   285
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   7
      Left            =   1455
      MaxLength       =   6
      TabIndex        =   19
      Top             =   2850
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   6
      Left            =   2730
      MaxLength       =   3
      TabIndex        =   18
      Top             =   2540
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   5
      Left            =   1455
      MaxLength       =   3
      TabIndex        =   17
      Top             =   2540
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   4
      Left            =   2730
      MaxLength       =   7
      TabIndex        =   8
      Top             =   1635
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   3
      Left            =   1455
      MaxLength       =   7
      TabIndex        =   7
      Top             =   1635
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   2
      Left            =   1455
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1050
      Width           =   315
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   1
      Left            =   1455
      TabIndex        =   1
      Top             =   750
      Width           =   2355
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   0
      Left            =   1455
      TabIndex        =   0
      Top             =   480
      Width           =   2340
   End
   Begin VB.TextBox textTM02 
      Height          =   264
      Left            =   2190
      MaxLength       =   6
      TabIndex        =   13
      Top             =   2250
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   21
      Left            =   380
      TabIndex        =   65
      Top             =   2280
      Width           =   950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "定稿日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   180
      Index           =   20
      Left            =   2760
      TabIndex        =   64
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "注意：不要使用PDF相關程式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   980
      Left            =   4170
      TabIndex        =   63
      Top             =   3450
      Width           =   1940
   End
   Begin VB.Label lblMsg 
      Caption         =   "lblMsg"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   3840
      TabIndex        =   61
      Top             =   1680
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "(專用期起日滿三年)"
      Height          =   180
      Index           =   3
      Left            =   4080
      TabIndex        =   60
      Top             =   1980
      Width           =   1905
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   2280
      X2              =   3420
      Y1              =   2070
      Y2              =   2070
   End
   Begin VB.Label Label3 
      Caption         =   $"frm020301.frx":0000
      ForeColor       =   &H000000C0&
      Height          =   350
      Left            =   630
      TabIndex        =   59
      Top             =   4890
      Width           =   4490
   End
   Begin VB.Label Label1 
      Caption         =   "台灣案催延展對象："
      Height          =   180
      Index           =   19
      Left            =   380
      TabIndex        =   58
      Top             =   4650
      Width           =   1640
   End
   Begin VB.Label Label1 
      Caption         =   "(1.台->台  2.大->台)"
      Height          =   180
      Index           =   18
      Left            =   2400
      TabIndex        =   57
      Top             =   4650
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "PS：報價定稿，須待智權人員確認"
      Height          =   180
      Index           =   17
      Left            =   3240
      TabIndex        =   56
      Top             =   5430
      Width           =   2870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否含第二期：         (Y：含)"
      Height          =   180
      Index           =   16
      Left            =   4340
      TabIndex        =   54
      Top             =   4880
      Visible         =   0   'False
      Width           =   2270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(1.非智權部同仁  2.全部)"
      Height          =   180
      Index           =   14
      Left            =   2190
      TabIndex        =   53
      Top             =   3480
      Width           =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "管制表列印對象："
      Height          =   180
      Index           =   13
      Left            =   380
      TabIndex        =   52
      Top             =   3480
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否含服務業務繳年費：          (Y：含)"
      Height          =   180
      Index           =   15
      Left            =   2535
      TabIndex        =   51
      Top             =   1065
      Width           =   3030
   End
   Begin VB.Label Label1 
      Caption         =   "製表日期："
      Height          =   180
      Index           =   12
      Left            =   375
      TabIndex        =   50
      Top             =   1410
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否含刊登廣告：        (Y：含)"
      Height          =   180
      Index           =   10
      Left            =   4340
      TabIndex        =   47
      Top             =   4610
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   2010
      X2              =   3150
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2100
      X2              =   3240
      Y1              =   4140
      Y2              =   4140
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2010
      X2              =   3150
      Y1              =   3870
      Y2              =   3870
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2100
      X2              =   3240
      Y1              =   2670
      Y2              =   2670
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2025
      X2              =   3165
      Y1              =   1770
      Y2              =   1770
   End
   Begin VB.Label LBL1 
      Height          =   180
      Left            =   2520
      TabIndex        =   46
      Top             =   2910
      Width           =   1130
   End
   Begin VB.Label Label1 
      Caption         =   "(1.管制表  2.定稿)"
      Height          =   180
      Index           =   11
      Left            =   1800
      TabIndex        =   45
      Top             =   3170
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   9
      Left            =   380
      TabIndex        =   44
      Top             =   4310
      Width           =   950
   End
   Begin VB.Label Label1 
      Caption         =   "代理人："
      Height          =   180
      Index           =   8
      Left            =   380
      TabIndex        =   43
      Top             =   4020
      Width           =   950
   End
   Begin VB.Label Label1 
      Caption         =   "申請人："
      Height          =   180
      Index           =   7
      Left            =   380
      TabIndex        =   42
      Top             =   3750
      Width           =   950
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   6
      Left            =   380
      TabIndex        =   41
      Top             =   3170
      Width           =   950
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   180
      Index           =   5
      Left            =   380
      TabIndex        =   40
      Top             =   2870
      Width           =   950
   End
   Begin VB.Label Label1 
      Caption         =   "業務區："
      Height          =   180
      Index           =   4
      Left            =   380
      TabIndex        =   39
      Top             =   2570
      Width           =   950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否含延展：        (Y：含)"
      Height          =   180
      Index           =   2
      Left            =   375
      TabIndex        =   38
      Top             =   1065
      Width           =   2040
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   1
      Left            =   375
      TabIndex        =   37
      Top             =   810
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   375
      TabIndex        =   36
      Top             =   525
      Width           =   945
   End
   Begin VB.Label Label8 
      Caption         =   "美國使用宣誓　費用："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   180
      Left            =   90
      TabIndex        =   55
      Top             =   5430
      Width           =   2090
   End
End
Attribute VB_Name = "frm020301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 20) As String, strTemp3 As String
Dim PLeft(0 To 13) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String
Dim m_strSales As String '智權人員 Add By Cheng 2002/03/01
'Add By Cheng 2002/11/11
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim m_strTM23Nation As String '申請人國籍Add By Cheng 2004/04/07
Dim m_strLanguage As String   '定稿語文  2009/12/1 ADD BY SONIA
Dim m_TM44 As String 'Add By Sindy 2010/11/12
'Dim boleFileSave As Boolean, m_TM01 As String 'Add By Sindy 2012/1/13
Dim m_CU15 As String 'Add By Sindy 2013/5/29
'Added by Lydia 2017/08/21
Dim strCP10 As String '相關總收文號的案件性質
Dim strCP53 As String '相關總收文號的第?期登記期
Dim strCP84 As String '相關總收文號的發文規費
Dim bolT102Again As Boolean 'Add by Amy 2018/09/11 延展再通知
'Added by Lydia 2019/08/22
Dim m_bWord As Boolean '是否開啟Word
Dim m_WordLeft As Long, m_WordTop As Long 'Word開啟位置
Dim bVisible As Boolean '是否顯示Word
Dim m_AttachPath As String


Private Sub cmdok_Click(Index As Integer)
'Add By Cheng 2003/01/08
Dim ii As Integer
'Add By Cheng 2002/11/21
Dim bolHadData As Boolean

On Error GoTo ErrorHandler

   Select Case Index
   Case 0 '確定
        'Modify by Amy 2018/08/31 原執行前畫面判斷程式改至FormCheck
        If FormCheck = False Then Exit Sub
        ClearQueryLog (Me.Name) 'Add By Sindy 2010/9/30 清除查詢印表記錄檔欄位
        'Add by Amy 2019/03/08 避免按確定再按到 ex:1080307 雅雯新增到非大陸案之撤三(1729) T-001876-刪除記錄
        Option1(0).Enabled = False
        Option1(1).Enabled = False
        '列印管制表
        If txt1(8) = "1" Then
            Screen.MousePointer = vbHourglass
            Me.Enabled = False
            bolHadData = Process
            Me.Enabled = True
            Screen.MousePointer = vbDefault
            '寄發管制表給智權人員
            If ChkMail.Value = 1 And bolHadData = True Then
                If ChkMail.Value = 1 Then MsgBox "寄發管制表，完畢!!", , "寄信"
            End If
            '2019/11/1 END
        '列印定稿
        Else
            'Add By Sindy 2019/12/11
            'Modify By Sindy 2025/4/30
            If textTM01 <> "" And textTM02 <> "" Then
               ChkMail.Value = 0
            Else
            '2025/4/30 END
               If ChkMail.Value = 0 Then
                  If MsgBox("要一併寄發管制表給智權人員嗎？", vbYesNo + vbInformation + vbDefaultButton2) = vbYes Then
                     ChkMail.Value = 1
                  End If
               End If
               '2019/12/11 END
            End If
            
            'Modify By Cheng 2003/02/11
            '列印定稿不要限制案件性質
   '        'Add By Cheng 2003/02/04
   '        '若有選擇延展或刊登廣告, 才產生定稿
'            If Me.TXT1(2).Text = "Y" Or Me.Text1.Text = "Y" Then
                Screen.MousePointer = vbHourglass
                Me.Enabled = False
                bolHadData = ProcessToWord
                Me.Enabled = True
                Screen.MousePointer = vbDefault
   '         Else
   '             MsgBox "請設定含延展或刊登廣告，否則不產生定稿!!!", vbExclamation + vbOKOnly
   '         End If
            'Add By Sindy 2019/11/1
            '寄發管制表給智權人員
            If ChkMail.Value = 1 And bolHadData = True Then
                Screen.MousePointer = vbHourglass
                Me.Enabled = False
                Process
                If ChkMail.Value = 1 Then MsgBox "寄發管制表，完畢!!", , "寄信"
                Me.Enabled = True
                Screen.MousePointer = vbDefault
            End If
            '2019/11/1 END
            
            'Modify By Sindy 2025/5/7
            If textTM01 <> "" And textTM02 <> "" Then
               textTM02 = ""
               textTM02_2 = ""
               txt1(3) = ""
               txt1(4) = ""
            End If
            '2025/5/7 END
        End If
        'Add by Amy 2019/03/08 避免按確定再按到
        Option1(0).Enabled = True
        Option1(1).Enabled = True
   Case 1 '結束
       Me.Enabled = False
       Unload Me
   Case Else
   End Select
   'Add By Cheng 2002/11/21
   Exit Sub

ErrorHandler:
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    MsgBox "(" & Err.Number & ")" & Err.Description
End Sub

'Modify By Sindy 2019/11/4
'Sub ProcessToWord()
Function ProcessToWord() As Boolean
'2019/11/4 END
'Add By Cheng 2002/09/18
'記錄本所案號
Dim strTM01 As String
Dim strTM02 As String
Dim strTM03 As String
Dim strTM04 As String
'Add By Sindy 2011/1/11
Dim strNP02 As String, strNP03 As String, strNP04 As String, strNP05 As String
Dim intRow As Integer
Dim strST15 As String, strNP10 As String
'2011/1/11 End
Dim strDueDate As String    'add by sonia 2014/5/6 不催延展者僅限於本所期限止於大於系統日超過一年者
Dim strCP09_New As String 'Add by Amy 2018/08/31
'Add by Amy 2018/09/11
Dim strWhere As String
Dim strCP13 As String, strErrCaseNo As String
   
bolT102Again = False
'end 2018/09/11
   strDueDate = CompDate(0, 1, strSrvDate(1))   'add by sonia 2014/5/6 不催延展者僅限於本所期限止於大於系統日超過一年者
   
   pub_QL05 = pub_QL05 & ";" & Label1(6) & "定稿" 'Add By Sindy 2010/9/30
   Screen.MousePointer = vbHourglass
   strSQL1 = ""
   strSQL2 = ""
   StrSQL6 = ""
   'Modify by Amy 2018/08/31 +大陸撤三
   If Option1(1).Value = True Then
        If Len(txt1(0)) <> 0 Then
            strSQL1 = strSQL1 + " and TM01 in (" & SQLGrpStr(txt1(0), 2) & ") "
            pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0)
        End If
        'Modify by Amy 2020/01/07 因延展改專用期起日會有錯,故改抓公告日+三個月+1天,無公告日則剔除 ex:T-141365
        If Len(txt1(21)) <> 0 Then
           'strSQL1 = strSQL1 & " And TM21>=" & Val(ChangeTStringToWString(txt1(21))) - 30000
           strSQL1 = strSQL1 & " And to_char(to_date(add_months(to_date(tm14,'YYYYMMDD'),3))+1,'YYYYMMDD')>=" & Val(ChangeTStringToWString(txt1(21))) - 30000
        End If
        If Len(txt1(22)) <> 0 Then
           'strSQL1 = strSQL1 & " And TM21<=" & Val(ChangeTStringToWString(txt1(22))) - 30000
           strSQL1 = strSQL1 & " And to_char(to_date(add_months(to_date(tm14,'YYYYMMDD'),3))+1,'YYYYMMDD')<=" & Val(ChangeTStringToWString(txt1(22))) - 30000
        End If
        If Len(txt1(21)) <> 0 Or Len(txt1(22)) <> 0 Then
           strSQL1 = strSQL1 & " And TM10='020' And (TM29<>'Y' OR TM29 IS NULL) And Nvl(tm14,0)<>0 And Nvl(tm21,0)<>0 " & _
                   " And Not Exists(Select * From CaseProgress b Where tm01=b.cp01 and tm02=b.cp02 and tm03=b.cp03 and tm04=b.cp04 and SubStr(b.CP09,1,1)='D' and b.CP10='1729')"
           pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & txt1(21) & "-" & txt1(22)
           pub_QL05 = pub_QL05 & ";" & Label1(9) & "020"
        End If
        'end 2020/01/07
   '非大陸撤三
   Else
       If Len(txt1(0)) <> 0 Then
          strSQL1 = strSQL1 + " and NP02 in (" & SQLGrpStr(txt1(0), 2) & ") "
          strSQL2 = strSQL2 + " and NP02 in (" & SQLGrpStr(txt1(0), 5) & ") "
          pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/9/30
       End If
       
       'Add By Sindy 2025/4/30
       If textTM01 <> "" And textTM02 <> "" Then
         If textTM01 = "TF" Then
            strSQL1 = strSQL1 + " and NP02='" & textTM01 & "' and NP03='" & textTM02 & textTM02_2 & "' and NP04='" & textTM03 & "' and NP05='" & textTM04 & "' "
            strSQL2 = strSQL2 + " and NP02='" & textTM01 & "' and NP03='" & textTM02 & textTM02_2 & "' and NP04='" & textTM03 & "' and NP05='" & textTM04 & "' "
            pub_QL05 = pub_QL05 & ";" & Label1(21) & textTM01 & textTM02 & textTM02_2 & textTM03 & textTM04
         Else
            strSQL1 = strSQL1 + " and NP02='" & textTM01 & "' and NP03='" & textTM02 & "' and NP04='" & textTM03 & "' and NP05='" & textTM04 & "' "
            strSQL2 = strSQL2 + " and NP02='" & textTM01 & "' and NP03='" & textTM02 & "' and NP04='" & textTM03 & "' and NP05='" & textTM04 & "' "
            pub_QL05 = pub_QL05 & ";" & Label1(21) & textTM01 & textTM02 & textTM03 & textTM04
         End If
       End If
       '2025/4/30 END
       
       'Modify By Cheng 2003/11/20
       ''910725 Sieg
       'If TXT1(2) = "Y" And Text1 = "Y" Then
       '  strExc(0) = "102,702,"
       'ElseIf TXT1(2) = "Y" And Text1 <> "Y" Then
       '  strExc(0) = "102,"
       'ElseIf TXT1(2) <> "Y" And Text1 = "Y" Then
       '  strExc(0) = "702,"
       'Else
       '  strExc(0) = ""
       'End If
       strExc(0) = ""
       If Me.txt1(2).Text = "Y" Then
           '2010/3/22 modify by sonia 加大陸被異議續展109
           strExc(0) = strExc(0) & "102,109,"
           pub_QL05 = pub_QL05 & ";" & Left(Label1(2), 6) & "含"  'Add By Sindy 2010/9/30
       End If
       If Me.Text1.Text = "Y" Then
           strExc(0) = strExc(0) & "702,"
           pub_QL05 = pub_QL05 & ";" & Left(Label1(10), 8) & "含"  'Add By Sindy 2010/9/30
       End If
       'add by nickc 2007/03/29
       If txt1(19).Text = "Y" Then
           strExc(0) = strExc(0) & "716,"
           pub_QL05 = pub_QL05 & ";" & Left(Label1(16), 7) & "含" 'Add By Sindy 2010/9/30
       End If
       'edit by nick 2004/07/20
       'If Me.txt1(16).Text = "Y" Then
       '    strExc(0) = strExc(0) & "715,"
       'End If
       'If Me.txt1(17).Text = "Y" Then
       '    strExc(0) = strExc(0) & "716,"
       'End If
       If Me.txt1(18).Text = "Y" Then
           strExc(0) = strExc(0) & "708,"
           pub_QL05 = pub_QL05 & ";" & Left(Label1(15), 11) & "含" 'Add By Sindy 2010/9/30
       End If
           
       StrSQL6 = strExc(0)
       
       Dim varTmp As Variant
       'Modify By Cheng 2003/02/11
       '列印定稿不要限制案件性質
       ''Modify By Cheng 2003/02/05
       ''案件性質非延展及刊登廣告者不印
       If txt1(1) <> "" Then
          varTmp = Split(txt1(1), ",")
          For i = 0 To UBound(varTmp)
             If varTmp(i) <> "" Then
                StrSQL6 = StrSQL6 & Format(varTmp(i)) & ","
             End If
          Next
          pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1)  'Add By Sindy 2010/9/30
       End If
       If StrSQL6 <> "" Then StrSQL6 = " AND NP07 IN (" & Left(StrSQL6, Len(StrSQL6) - 1) & ") "
       '本所期限
       If Len(txt1(3)) <> 0 Then
          StrSQL6 = StrSQL6 + " AND NP08>=" & Val(ChangeTStringToWString(txt1(3))) & ""
       End If
       If Len(Trim(txt1(4))) <> 0 Then
          StrSQL6 = StrSQL6 + " AND NP08<=" & Val(ChangeTStringToWString(txt1(4))) & " "
       End If
       If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Then
          'Modify by Amy 2019/03/07 原:Label1(3)-bug
          pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & txt1(3) & "-" & txt1(4)    'Add By Sindy 2010/9/30
       End If
       StrSQL6 = StrSQL6 & " AND (NP06 IS NULL OR NP06='') "
    '   '業務區
    '   If Len(TXT1(5)) <> 0 Then
    '       'Modify By Cheng 2003/03/11
    '   '    StrSQL6 = StrSQL6 + " AND s1.ST03>='" & TXT1(5) & "' "
    '       StrSQL6 = StrSQL6 + " AND s1.ST15>='" & TXT1(5) & "' "
    '   End If
    '   If Len(TXT1(6)) <> 0 Then
    '       'Modify By Cheng 2003/03/15
    '   '    StrSQL6 = StrSQL6 + " AND s1.ST03<='" & TXT1(6) & "' "
    '       StrSQL6 = StrSQL6 + " AND s1.ST15<='" & TXT1(6) & "' "
    '   End If
       If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
          pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(5) & "-" & txt1(6)    'Add By Sindy 2010/9/30
       End If
       '智權人員
       If Len(txt1(7)) <> 0 Then
    '       StrSQL6 = StrSQL6 + " AND NP10='" & TXT1(7) & "' "
           pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(7) & lbl1    'Add By Sindy 2010/9/30
       End If
      '申請人
      If Len(txt1(9)) <> 0 Then
          strSQL1 = strSQL1 + " AND (TM23>='" & GetNewFagent(txt1(9)) & "') "
          strSQL2 = strSQL2 + " AND (SP08>='" & GetNewFagent(txt1(9)) & "' OR SP58>='" & GetNewFagent(txt1(9)) & "' OR SP59>='" & GetNewFagent(txt1(9)) & "') "
      End If
      If Len(txt1(10)) <> 0 Then
          strSQL1 = strSQL1 + " AND (TM23<='" & GetNewFagent(txt1(10)) & "') "
          strSQL2 = strSQL2 + " AND (SP08<='" & GetNewFagent(txt1(10)) & "' OR SP58<='" & GetNewFagent(txt1(10)) & "' OR SP59<='" & GetNewFagent(txt1(10)) & "') "
      End If
      If Len(Trim(txt1(9))) <> 0 Or Len(Trim(txt1(10))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & Label1(7) & txt1(9) & "-" & txt1(10)  'Add By Sindy 2010/9/30
      End If
       '代理人
       If Len(txt1(11)) <> 0 Then
           strSQL1 = strSQL1 + " and TM44>='" & GetNewFagent(txt1(11)) & "' "
           strSQL2 = strSQL2 + " and SP26>='" & GetNewFagent(txt1(11)) & "' "
       End If
       If Len(txt1(12)) <> 0 Then
           strSQL1 = strSQL1 + " and TM44<='" & GetNewFagent(txt1(12)) & "' "
           strSQL2 = strSQL2 + " and SP26<='" & GetNewFagent(txt1(12)) & "' "
       End If
       If Len(Trim(txt1(11))) <> 0 Or Len(Trim(txt1(12))) <> 0 Then
          pub_QL05 = pub_QL05 & ";" & Label1(8) & txt1(11) & "-" & txt1(12)   'Add By Sindy 2010/9/30
       End If
       '申請國家
       If Len(txt1(13)) <> 0 Then
           strSQL1 = strSQL1 + " AND TM10>='" & txt1(13) & "' "
           strSQL2 = strSQL2 + " AND SP09>='" & txt1(13) & "' "
       End If
       strSQL1 = strSQL1 & " AND (TM29<>'Y' OR TM29 IS NULL) "
       strSQL2 = strSQL2 & " AND (SP15<>'Y' OR SP15 IS NULL) "
       If Len(txt1(14)) <> 0 Then
           strSQL1 = strSQL1 + " AND TM10<='" & txt1(14) & "' "
           strSQL2 = strSQL2 + " AND SP09<='" & txt1(14) & "' "
       End If
       If Len(Trim(txt1(13))) <> 0 Or Len(Trim(txt1(14))) <> 0 Then
          pub_QL05 = pub_QL05 & ";" & Label1(9) & txt1(13) & "-" & txt1(14)   'Add By Sindy 2010/9/30
       End If
       
       'Add by Amy 2018/09/11 +延展再通知
       Call ShowT102Again
       If lblMsg.Caption <> MsgText(601) Then
            If txt1(13) = "000" And txt1(14) = "000" Then
                '通知北、中所
                bolT102Again = True
                strSQL1 = strSQL1 & " And s3.ST06 in('1','2') And CU13=s3.ST01(+)"
                strSQL2 = strSQL2 & " And s3.ST06 in('1','2') And CU13=s3.ST01(+)"
                pub_QL05 = pub_QL05 & ";通知北、中所"
            End If
            If txt1(13) = "020" And txt1(14) = "020" Then
                If Val(txt1(3)) + 19110000 < Val(GetPreMonLastDate(DBDATE(DateAdd("m", 6, Format(strSrvDate(1), "####/##/##"))), False)) Then
                    '到期前2個月通知北、中、南所
                    bolT102Again = True
                    strSQL1 = strSQL1 & " And s3.ST06 <>'4' And CU13=s3.ST01(+)"
                    strSQL2 = strSQL2 & " And s3.ST06 <>'4' And CU13=s3.ST01(+)"
                    pub_QL05 = pub_QL05 & ";通知北、中、南所"
                ElseIf Val(txt1(3)) + 19110000 < Val(GetPreMonLastDate(DBDATE(DateAdd("m", 18, Format(strSrvDate(1), "####/##/##"))), False)) Then
                    '通知北、中所
                    bolT102Again = True
                    strSQL1 = strSQL1 & " And s3.ST06 in('1','2') And CU13=s3.ST01(+)"
                    strSQL2 = strSQL2 & " And s3.ST06 in('1','2') And CU13=s3.ST01(+)"
                    pub_QL05 = pub_QL05 & ";通知北、中所"
                End If
            End If
       End If
       
       'Add By Sindy 2015/4/17 台灣案催延展對象
       '大->台係指葉經理及巨京收文且申請人國籍非台灣者,
       '其他案件皆屬台->台範圍
       If Trim(txt1(13)) = "000" And Trim(txt1(14)) = "000" Then
          '系統別有T者
          strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
          s = 0
          For i = 0 To UBound(strTemp2)
             If strTemp2(i) = "T" Then
                s = 1
                Exit For
             End If
          Next i
          If s = 1 And txt1(2) = "Y" Then 'T含延展
             If Trim(txt1(20)) = "1" Then '1.台->台
                'Modify By Sindy 2020/8/13 改判斷 TM44.FC代理人
                'StrSQL6 = StrSQL6 + " AND not(Substr(NP10,1,5) in('67002','96029','96030','MCTF0') AND CU10>'010') "
                strSQL1 = strSQL1 + " AND not(Substr(NP10,1,5) in('67002','96029','96030','MCTF0') AND TM44 is not null) "
                strSQL2 = strSQL2 + " AND not(Substr(NP10,1,5) in('67002','96029','96030','MCTF0') AND SP26 is not null) "
             ElseIf Trim(txt1(20)) = "2" Then '2.大->台
                'Modify by Amy 2017/01/10 +MCTF特殊人員
                'Modify By Sindy 2020/8/13 改判斷 TM44.FC代理人
                'StrSQL6 = StrSQL6 + " AND SubStr(NP10,1,5) in('67002','96029','96030','MCTF0') AND CU10>'010' "
                strSQL1 = strSQL1 + " AND SubStr(NP10,1,5) in('67002','96029','96030','MCTF0') AND TM44 is not null "
                strSQL2 = strSQL2 + " AND SubStr(NP10,1,5) in('67002','96029','96030','MCTF0') AND SP26 is not null "
             End If
             pub_QL05 = pub_QL05 & ";" & Label1(19) & txt1(20) & Label1(18)
          End If
       End If
       '2015/4/17 END
       
       'add by nickc 2006/05/30
       If txt1(16) = "1" Then
    '       StrSQL6 = StrSQL6 + " AND substr(s1.ST15,1,1)<>'S' "
           pub_QL05 = pub_QL05 & ";" & Label1(13) & "非智權部同仁"   'Add By Sindy 2010/9/30
       End If
       'add by nickc 2006/05/30 延展和第二期專用權須存在
       If InStr(1, txt1(1), "716") <> 0 Or InStr(1, txt1(1), "102") <> 0 Then
           strSQL1 = strSQL1 & " and decode(np07,716,tm17,102,tm17,'Y')='Y' "
       End If
   End If
   'end 2018/08/31
    
   'Modify By Cheng 2003/01/10 取消限制收文號類別
   'Modify By Cheng 2003/02/04 案件進度檔的收文號不可為NULL
   'Modify By Cheng 2003/03/11 智權人員的部門別抓ST15
   'Modify By Cheng 2003/10/02 排序依業務區+下一程序業務員,再加客戶編號
   'Modify By Cheng 2004/04/07 加判斷可使用的系統類別+案件性質
   '2008/9/26 MODIFY BY SONIA 改以NP01抓CP09
   '2009/12/2 modify by sonia 加TM77畫面上的定稿語文
   'Modify By Sindy 2010/6/11 增加NP22
   'Modify By Sindy 2010/11/12  增加TM44
   'Modfiy By Sindy 2012/6/28 刪除where條件裡的AND NP02=CP01(+) AND NP03=CP02(+) AND NP04=CP03(+) AND NP05=CP04(+)因ex.1020301-1020331的TF使用宣誓會抓不到資料
'   strSql = "SELECT CP09,NP07,TM01,TM10,TM22,TM01,TM02,TM03,TM04,S1.ST15,CP13,TM23,NP10,TM77,NP22,TM44,S1.ST04 as ST04 FROM Staff_Group, NEXTPROGRESS,TRADEMARK,CASEPROGRESS,NATION,STAFF s1,STAFF S2,ACC090,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=CP01(+) AND NP03=CP02(+) AND NP04=CP03(+) AND NP05=CP04(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=s1.ST01(+) AND NP02=CPM01(+) AND NP07=(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) and s1.ST15=a0901(+) And CP09 IS NOT NULL And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 " & StrSQL6 & strSQL1
'   strSql = strSql + " union all select CP09,NP07,SP01,SP09,SP21,SP01,SP02,SP03,SP04,S1.ST15,CP13,SP08,NP10,'' TM77,NP22,'' TM44,S1.ST04 as ST04 FROM Staff_Group, NEXTPROGRESS,SERVICEPRACTICE,CASEPROGRESS,NATION,STAFF s1,STAFF S2,ACC090,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=CP01(+) AND NP03=CP02(+) AND NP04=CP03(+) AND NP05=CP04(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=s1.ST01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND SP09=NA01(+) AND s1.ST15=a0901(+) And CP09 IS NOT NULL And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 " & StrSQL6 & strSQL2
'   strSql = strSql & " ORDER BY 10, 13, 12, 6, 7, 8, 9 "
   'Modify By Sindy 2013/5/29 +CU15
   'Modify By Sindy 2013/9/16 +TM129
   'Modify By Sindy 2015/3/2 +,NP08,NP09,NP01
   'Modified by Lydia 2017/08/21 +CP10,CP53,CP84
   'Modify by Amy 2018/09/11 +CP05/iif
   strSql = "SELECT CP09,NP07,TM01,TM10,TM22,TM01,TM02,TM03,TM04,S1.ST15,CP13,TM23,NP10,TM77,NP22,TM44,S1.ST04 as ST04,CU15,TM129,NP08,NP09,NP01,CP10,CP53,CP84,CP05 FROM Staff_Group, NEXTPROGRESS,TRADEMARK,CASEPROGRESS,NATION,STAFF s1,STAFF S2,ACC090,CASEPROPERTYMAP,CUSTOMER" & IIf(bolT102Again = True, ",Staff s3", "") & " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=s1.ST01(+) AND NP02=CPM01(+) AND NP07=(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) and s1.ST15=a0901(+) And CP09 IS NOT NULL And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 " & StrSQL6 & strSQL1
   strSql = strSql + " union all select CP09,NP07,SP01,SP09,SP21,SP01,SP02,SP03,SP04,S1.ST15,CP13,SP08,NP10,'' TM77,NP22,'' TM44,S1.ST04 as ST04,CU15,'' TM129,NP08,NP09,NP01,CP10,CP53,CP84,CP05 FROM Staff_Group, NEXTPROGRESS,SERVICEPRACTICE,CASEPROGRESS,NATION,STAFF s1,STAFF S2,ACC090,CASEPROPERTYMAP,CUSTOMER" & IIf(bolT102Again = True, ",Staff s3", "") & " WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=s1.ST01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND SP09=NA01(+) AND s1.ST15=a0901(+) And CP09 IS NOT NULL And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 " & StrSQL6 & strSQL2
   'end 2018/09/11
   strSql = strSql & " ORDER BY 10, 13, 12, 6, 7, 8, 9 "
   '2012/6/28 End
   'Add by Amy 2018/08/31 +大陸撤三
   If Option1(1).Value = True Then
      strSql = ",to_number(SubStr(TM21,1,4))+3||SubStr(TM21,5) NP08,to_number(SubStr(TM21,1,4))+3||SubStr(TM21,5) NP09"
      'Moidfy by Amy 2018/11/02 因造成多筆,count筆數有問題,故不抓CaseProgress
      'strSql = "SELECT CP09,'1729' NP07,TM01,TM10,TM22,TM01,TM02,TM03,TM04,S1.ST15,CP13,TM23,CU13 NP10,TM77,'' NP22,TM44,S1.ST04 as ST04,CU15,TM129" & strSql & ",'' NP01,CP10,CP53,CP84,CP05 FROM Staff_Group,TradeMark,CaseProgress,Nation,Staff s1,Staff S2,ACC090,CasePropertyMap,Customer WHERE TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND CU13=s1.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP14=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) and s1.ST15=a0901(+) And CP09 IS NOT NULL And '" & strGroup & "'=SG01 And SG02=CP01 And SG03=CP10 And TM21 is not null " & strSQL1
      strSql = "SELECT '' CP09,'1729' NP07,TM01,TM10,TM22,TM01,TM02,TM03,TM04,S1.ST15,'' CP13,TM23,CU13 NP10,TM77,'' NP22,TM44,S1.ST04 as ST04,CU15,TM129" & strSql & ",'' NP01,'' CP10,'' CP53,'' CP84,'' CP05 FROM TradeMark,Nation,Staff s1,ACC090,Customer WHERE CU13=s1.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) and s1.ST15=a0901(+) And TM21 is not null " & strSQL1
      strSql = strSql & " ORDER BY 10, 13, 12, 6, 7, 8, 9 "
   End If
   'end 2018/09/11
   intRow = 0 'Add By Sindy 2011/1/11
   Dim m_rs As New ADODB.Recordset
   Set m_rs = New ADODB.Recordset
   'CheckOC
   With m_rs
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
            'Add By Sindy 2023/4/11 檢查沒有目前智權人員時,後續新增CP會有錯
            If txt1(8) = "2" Then '2.定稿
               .MoveFirst
               Do While .EOF = False
                  strCP13 = PUB_GetAKindSalesNo(.Fields(5), .Fields(6), .Fields(7), .Fields(8))
                  If strCP13 = "" Then
                     strErrCaseNo = strErrCaseNo & "、" & .Fields(5) & .Fields(6) & .Fields(7) & .Fields(8)
                  End If
                  .MoveNext
               Loop
               If strErrCaseNo <> "" Then
                  strErrCaseNo = Mid(strErrCaseNo, 2)
                  MsgBox "案號 " & strErrCaseNo & " 沒有讀取到目前智權人員，" & vbCrLf & "後續新增進度會有錯！請檢查~", vbExclamation
                  Screen.MousePointer = vbDefault
                  Exit Function
               End If
            End If
            '2023/4/11 END
            
            ProcessToWord = True 'Add By Sindy 2019/11/4
'            InsertQueryLog (.RecordCount) 'Add By Sindy 2010/9/30
            'Add By Cheng 2002/09/18
            strTM01 = ""
            strTM02 = ""
            strTM03 = ""
            strTM04 = ""
'            boleFileSave = False 'Add By Sindy 2012/1/13
           .MoveFirst
           Do While .EOF = False
               strST15 = "" & .Fields(9)
               strNP10 = "" & .Fields(12)
               'Add By Sindy 2011/1/11 抓未收文資料時, 若np10智權人員為離職時, 改用PUB_GetAKindSalesNo抓智權人員
               '2011/4/13 modify by sonia 不下智權人員才做,否則下離職智權人員條件時會因已轉換而無資料
               'If CheckStr(.Fields("ST04")) <> "1" Then
               If CheckStr(.Fields("ST04")) <> "1" And Len(txt1(7)) = 0 Then
                  strNP02 = .Fields(5)
                  strNP03 = .Fields(6)
                  strNP04 = .Fields(7)
                  strNP05 = .Fields(8)
                  strST15 = PUB_GetStaffST15(PUB_GetAKindSalesNo(strNP02, strNP03, strNP04, strNP05), "1")
                  strNP10 = PUB_GetAKindSalesNo(strNP02, strNP03, strNP04, strNP05)
               End If
               If Len(txt1(5)) <> 0 Then
                  If strST15 < txt1(5) Then GoTo GoToExit3
               End If
               If Len(txt1(6)) <> 0 Then
                  If strST15 > txt1(6) Then GoTo GoToExit3
               End If
               If Len(txt1(7)) <> 0 Then
                  If strNP10 <> Trim(txt1(7)) Then GoTo GoToExit3
               End If
               intRow = intRow + 1 '記錄筆數
               '2011/1/11 End
               
               '相同案號只印一份定稿
               'Modify By Sindy 2012/6/29 馬德里申請國家相同(因申請多個商品類別時會多案號)只需要出一份定稿
               If (strTM01 <> "" & .Fields(5) Or strTM02 <> "" & .Fields(6) Or strTM03 <> "" & .Fields(7) Or strTM04 <> "" & .Fields(8)) And _
                  Not ("" & .Fields(5) = "TF" And "" & .Fields(1) = "105" And strTM02 = "" & .Fields(6) And strTM04 = "" & .Fields(8)) Then
                  'modify by sonia 2015/9/21 改用共用函數GetTWordLng
                  ''取得申請人國籍
                  ''2009/12/1 MODIFY BY SONIA 不以申請人國籍判斷,改以定稿語文判斷
                  ''m_strTM23Nation = GetCustomerNation("" & .Fields(11).Value)
                  'm_strLanguage = GetLetterLanguage(.Fields(5), .Fields(6), .Fields(7), .Fields(8))
                  ''2009/12/24 ADD BY SONIA 非中文定稿都以英文列印,中文定稿時需先判斷國籍T-115998,否則若案件未設會抓成台->各國
                  'If m_strLanguage <> "1" Then
                  '   m_strLanguage = "3"
                  'ElseIf m_strLanguage = "1" Then
                  '   m_strTM23Nation = GetCustomerNation("" & .Fields(11).Value)
                  '   '2015/7/14 MODIFY BY SONIA T-119695 僅葉經理及巨京的國外客戶以外->台列印,其他人的國內外客戶皆以台->台列印
                  '   'If m_strTM23Nation > "010" Then m_strLanguage = "2"
                  '   If m_strTM23Nation > "010" And (strNP10 = "67002" Or strNP10 = "96029" Or strNP10 = "96030") Then m_strLanguage = "2"
                  'End If
                  ''2009/12/24 END
                  'If IsNull(.Fields("TM77")) = False Then
                  '   If IsEmptyText(.Fields("TM77")) = False Then
                  '      m_strLanguage = .Fields("TM77")
                  '   End If
                  'End If
                  If IsNull(.Fields("TM77")) = False Then
                     m_strLanguage = .Fields("TM77")
                  Else
                     m_strLanguage = GetTWordLng(.Fields(5), .Fields(6), .Fields(7), .Fields(8))
                  End If
                  '2009/12/1 END
                  'end 2015/9/21
                  
                  'Add By Sindy 2010/11/12
                  m_TM44 = ""
                  If IsNull(.Fields("TM44")) = False Then
                     If IsEmptyText(.Fields("TM44")) = False Then
                        m_TM44 = .Fields("TM44")
                     End If
                  End If
                  
                  'Modify By Sindy 2010/6/11 增加NP22
                  m_CU15 = Trim("" & .Fields("CU15")) 'Add By Sindy 2013/5/29
                  'Modify By Sindy 2013/9/16 檢查個案有設定不催延展者的案件不出定稿
                  'modify by sonia 2014/5/6 再加入控制, 不催延展者僅限於本所期限止於大於系統日超過一年者
                  If "" & .Fields(1) = "102" And "" & .Fields("TM129") = "Y" And Val(ChangeTStringToWString(txt1(4))) >= Val(strDueDate) Then
                     strTM01 = "" & .Fields(5)
                     strTM02 = "" & .Fields(6)
                     strTM03 = "" & .Fields(7)
                     strTM04 = "" & .Fields(8)
                     intRow = intRow - 1 '記錄筆數
                     GoTo GoToExit3
                  End If
                  '2013/9/16 END
                  
                  'Added by Lydia 2017/08/21 相關總收文號的案件性質
                  strCP10 = "" & .Fields("CP10")
                  strCP53 = "" & .Fields("CP53")
                  strCP84 = "" & .Fields("CP84")
                  '發文-服務業務結果會將NP01由A類改為C類
                  If Left(.Fields("NP01"), 1) = "C" Then
                     strSql = "select c2.cp10,c2.cp53,c2.cp84 from caseprogress c1,caseprogress c2 where c1.cp09='" & .Fields("NP01") & "' and c1.cp43=c2.cp09(+) and c2.cp10='708' "
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        strCP10 = "" & RsTemp.Fields("CP10")
                        strCP53 = "" & RsTemp.Fields("CP53")
                        strCP84 = "" & RsTemp.Fields("CP84")
                     End If
                  End If
                  'end 2017/08/21
  
                  'Add by Amy 2018/08/31 +大陸撤三
                  If Option1(1).Value = True Then
                    'Modify by Amy 2018/11/06 期限拿掉 原:.Fields("NP08")=tm21
                    'Modify By Sindy 2019/10/22 + , .Fields("TM10"), .Fields("TM23"), m_TM44
                    'Modify By Sindy 2021/1/11 + , TXT1(17) 加定稿日期
                    'Modify By Sindy 2025/4/30 有輸入個案則示為非大宗 +IIf(textTM01 <> "" And textTM02 <> "", False, True)
                    If PUB_AddCaseProgressD("1729", .Fields(5), .Fields(6), .Fields(7), .Fields(8), "", "", "", "", , strCP09_New, .Fields("TM10"), .Fields("TM23"), m_TM44, txt1(17), IIf(textTM01 <> "" And textTM02 <> "", False, True)) = False Then
                        MsgBox "大陸撤三新增進度檔失敗！作業中斷！", vbCritical
                        Screen.MousePointer = vbDefault
                        Exit Function
                    End If
                    PrintLetter "1729", CheckStr(.Fields(2)), CheckStr(.Fields(3)), CheckStr(.Fields(4)), GetTodayDate, CheckStr(strCP09_New), CheckStr(.Fields(9)), "" & .Fields(12).Value, "" & .Fields(6), "" & .Fields(7), "" & .Fields(8), "" & .Fields("NP22"), strCP09_New

                  '非大陸撤三
                  Else
                    'Add by Amy 2018/09/11 +延展再通知
                    If bolT102Again = True Then
                        strWhere = " And CP05>=" & Val(DBDATE(DateAdd("m", -5, Format(strSrvDate(1), "####/##/##")))) & " And CP05<=" & Val(strSrvDate(1))
                        'Modify By Sindy 2019/10/22 + , .Fields("TM10"), .Fields("TM23"), m_TM44
                        'Modify By Sindy 2021/1/11 + , TXT1(17) 加定稿日期
                        'Modify By Sindy 2025/4/30 有輸入個案則示為非大宗 +IIf(textTM01 <> "" And textTM02 <> "", False, True)
                        If PUB_AddCaseProgressD("1725", .Fields(5), .Fields(6), .Fields(7), .Fields(8), .Fields("NP08"), .Fields("NP09"), .Fields("NP22"), .Fields("NP01"), strWhere, strCP09_New, .Fields("TM10"), .Fields("TM23"), m_TM44, txt1(17), IIf(textTM01 <> "" And textTM02 <> "", False, True)) = False Then
                            MsgBox "延展再通知增進度檔失敗！作業中斷！", vbCritical
                            Screen.MousePointer = vbDefault
                            Exit Function
                        End If
                    Else
                        'Add By Sindy 2015/3/2
                        'Modify By Sindy 2019/10/22 + , .Fields("TM10"), .Fields("TM23"), m_TM44
                        'Modify By Sindy 2021/1/11 + , TXT1(17) 加定稿日期
                        'Modify By Sindy 2025/4/30 有輸入個案則示為非大宗 +IIf(textTM01 <> "" And textTM02 <> "", False, True)
                        'Modify By Sindy 2025/4/17 統一使用相同函數
                        'If PUB_AddCP1725(.Fields(5), .Fields(6), .Fields(7), .Fields(8), .Fields("NP08"), .Fields("NP09"), .Fields("NP01"), .Fields("NP22"), .Fields("TM10"), .Fields("TM23"), strCP09_New, m_TM44, TXT1(17)) = False Then
                        strWhere = " and cp30='" & .Fields("NP22") & "' and cp05>=" & Format(DateAdd("M", -1, ChangeWStringToWDateString(strSrvDate(1))), "YYYYMMDD")
                        If PUB_AddCaseProgressD("1725", .Fields(5), .Fields(6), .Fields(7), .Fields(8), .Fields("NP08"), .Fields("NP09"), .Fields("NP22"), .Fields("NP01"), strWhere, strCP09_New, .Fields("TM10"), .Fields("TM23"), m_TM44, txt1(17), IIf(textTM01 <> "" And textTM02 <> "", False, True)) = False Then
                        '2025/4/17 END
                           MsgBox "新增進度檔【通知期限】失敗！作業中斷！", vbCritical
                           Screen.MousePointer = vbDefault
                           Exit Function
                        End If
                        '2015/3/2 END
                    End If
                    'end 2018/09/11
                    PrintLetter CheckStr(.Fields(1)), CheckStr(.Fields(2)), CheckStr(.Fields(3)), CheckStr(.Fields(4)), GetTodayDate, CheckStr(.Fields(0)), CheckStr(.Fields(9)), "" & .Fields(12).Value, "" & .Fields(6), "" & .Fields(7), "" & .Fields(8), "" & .Fields("NP22"), strCP09_New
                  End If
                  'end 2018/08/31
                  
                  'Remove by Morgan 2008/11/25 改用開窗信封不必再印
                  ''新增地址條列表資料
                  'pub_AddressListSN = pub_AddressListSN + 1
                  'PUB_AddNewAddressList strUserNum, "" & .Fields(5).Value, "" & .Fields(6).Value, "" & .Fields(7).Value, "" & .Fields(8).Value, "" & pub_AddressListSN, "0"
                  'end 2008/11/25
                  
                  strTM01 = "" & .Fields(5)
                  strTM02 = "" & .Fields(6)
                  strTM03 = "" & .Fields(7)
                  strTM04 = "" & .Fields(8)
               End If
GoToExit3:
               .MoveNext
           Loop
           
'            'Add By Sindy 2012/1/13
'            If boleFileSave = True Then
'               MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
'            End If
'            '2012/1/13 End
            
            'Added by Lydia 2019/08/22 關閉Word
            If m_bWord = True Then
                Pub_RePosWord g_WordAp, bVisible, m_WordLeft, m_WordTop '還原Word位置
                g_WordAp.ActiveDocument.Close wdDoNotSaveChanges
                g_WordAp.Quit wdDoNotSaveChanges
                Set g_WordAp = Nothing
                m_bWord = False
            End If
            
            'Add By Sindy 2011/1/11
            If intRow = 0 Then
              InsertQueryLog (0) 'Add By Sindy 2010/9/30
              ProcessToWord = False 'Add By Sindy 2019/11/4
              ShowNoData
              Screen.MousePointer = vbDefault
              Exit Function
            Else
              InsertQueryLog (intRow) 'Add By Sindy 2010/9/30
            End If
            '2011/1/11 End
       Else
         InsertQueryLog (0) 'Add By Sindy 2010/9/30
         ProcessToWord = False 'Add By Sindy 2019/11/4
         ShowNoData
         Screen.MousePointer = vbDefault
         Exit Function
       End If
       'CheckOC
   End With
   'Modify By Sindy 2019/11/19
   'ShowPrintOk
   If ChkMail.Value = 0 Then ShowPrintOk
   '2019/11/19 END
   Screen.MousePointer = vbDefault
End Function

'Modify By Sindy 2010/6/11 增加NP22
'Modified by Lydia 2019/05/20 +處理狀況strET03
'Modify By Sindy 2019/11/6 + , ByVal strLD18 As String : 信函收文號
Private Sub InsExpField(ByVal strNP07 As String, ByVal strTM01 As String, ByVal strTM10 As String, _
   ByVal strDate As String, ByVal strSysDate As String, ByVal strCP09 As String, ByVal strCP12 As String, _
   ByVal strNP10 As String, ByVal strTM02 As String, ByVal strTM03 As String, _
   ByVal strTM04 As String, ByVal strNP22 As String, ByVal strET03 As String, _
   ByVal strLD18 As String)
Dim o71713 As Double  '點數即服務費
Dim o71708 As Double  '規費
Dim o71613 As Double
Dim o71608 As Double
Dim o71513 As Double
Dim o71508 As Double
Dim intCnt As Integer      '商品類別數
Dim dbl_usxrate As Double, dbl_fee As Double 'Add By Sindy 2010/11/12
Dim strNA2 As String, strTmp As String 'Add By Sindy 2010/12/6
Dim dbl_official As Double  'add by sonia 2014/6/27
Dim dblTMKindCnt As Double, varTmp As Variant 'Add By Sindy 2014/10/28
   
'Dim strSQL As String
' 下一程序
'Dim StrNP07 As Strubg
' 系統別
'Dim strTM01 As String
' 申請國家
'Dim strTM10 As String
' 專用期限止日
'Dim strDate As String
' 系統日
'Dim strSysDate As String
' 總收文號
'Dim strCP09 As String
   
   'Add By Sindy 2010/12/6 取得馬德里指定國家
   strNA2 = Empty
   If strTM01 = "TF" Then
      'modify by sonia 2017/12/12 +TM29條件
      strSql = "SELECT DISTINCT(TM10) FROM TradeMark " & _
                    "WHERE TM01 = '" & strTM01 & "' AND " & _
                    "SUBSTR(TM02,1,5) = '" & Mid(strTM02, 1, 5) & "' AND " & _
                    "TM04 <> '00' AND (TM16 IS NULL OR TM16<>'2') AND TM29 IS NULL"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         RsTemp.MoveFirst
         Do While RsTemp.EOF = False
            If IsNull(RsTemp.Fields("TM10")) = False Then
               strTmp = GetNationName(RsTemp.Fields("TM10"), 0)
               If IsEmptyText(strTmp) = False Then
                  If strNA2 <> Empty Then: strNA2 = strNA2 & "、"
                  strNA2 = strNA2 & strTmp
               End If
            End If
            RsTemp.MoveNext
         Loop
      End If
   End If
   '2010/12/6 End
   
   'Add By Sindy 2014/10/28
   '取得商品類別數
   strSql = "SELECT tm09 from trademark " & _
            "where tm01='" & strTM01 & "' and tm02='" & strTM02 & "' and tm03='" & strTM03 & "' and tm04='" & strTM04 & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   dblTMKindCnt = 1
   If intI = 1 Then
      If "" & RsTemp.Fields("TM09").Value <> "" Then
         If InStr(RsTemp.Fields("TM09").Value, ",") > 0 Then
            varTmp = Split(RsTemp.Fields("TM09").Value, ",")
            dblTMKindCnt = UBound(varTmp) + 1
         End If
      End If
   End If
   '2014/10/28 END
   '2009/12/1 modify BY SONIA 不以申請人國籍m_strTM23Nation判斷,改以定稿語文m_strLanguage判斷
   Select Case strNP07
      ' 延展
      Case "102":
        'add by nickc 2006/05/05
        If strTM01 = "TF" Then
'              ' 清除定稿例外欄位檔原有資料
'              'Modified by Lydia 2019/05/20 "00"=>strET03
'              EndLetter "10", strCP09, strET03, strUserNum
'               'add by nickc 2008/04/25 案件回覆單
'               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                        "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
'                        "'下一程序','" & strNP07 & "')"
'               cnnConnection.Execute strSql
'               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                        "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
'                        "'服務專線','" & PUB_GetServiceLine(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "')"
'               cnnConnection.Execute strSql
''cancel by sonia 2016/2/24 定稿設計改用<業務員>,因中所新人特別+主管名字
''               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                        "VALUES ('" & "10" & "','" & strCP09 & "','" & "00" & "','" & strUserNum & "'," & _
''                        "'下一程序業務員','" & GetStaffName(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "')"
''               cnnConnection.Execute strSql
'               'Add By Sindy 2010/12/6
'               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                        "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
'                        "'馬德里指定國家','" & strNA2 & "')"
'               cnnConnection.Execute strSql
'               '2010/12/6 End
'               'end 2019/05/20
               'Modify By Sindy 2019/11/29 改為報價定稿
               strET03 = "00"
               'Modify By Sindy 2021/3/18 TF續展的報價定稿取消,改為單純續展期限通知函
               ' 清除定稿例外欄位檔原有資料
               EndLetter "10", strCP09, strET03, strUserNum
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                        "'下一程序','" & strNP07 & "')"
               cnnConnection.Execute strSql
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                        "'服務專線','" & PUB_GetServiceLine(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "')"
               cnnConnection.Execute strSql
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                        "'馬德里指定國家','" & strNA2 & "')"
               cnnConnection.Execute strSql
               
'               If Val(Trim(Me.textMoney.Text)) <> 0 Then
'                  ' 清除定稿例外欄位檔原有資料
'                  EndLetter "10", strCP09, strET03, strUserNum
'                  '+ 信函收文號
'                  PUB_AddLetterCache strCP09, strNP22, strCP09, "10", strET03, , IIf(strSrvDate(1) >= T商標電子化啟用日, strLD18, "")
'                  '費用
'                  strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
'                               "VALUES ('" & strCP09 & "'," & strNP22 & ",'費用','" & Me.textMoney.Text & "','Y')"
'                  cnnConnection.Execute strSql
'                  '案件回覆單
'                  'LCV05拿掉Y (影響報價備註CP64)
'                  strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
'                               "VALUES ('" & strCP09 & "'," & strNP22 & ",'下一程序','" & strNP07 & "','')"
'                  cnnConnection.Execute strSql
'                  strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
'                               "VALUES ('" & strCP09 & "'," & strNP22 & ",'服務專線','" & PUB_GetServiceLine(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "','')"
'                  cnnConnection.Execute strSql
'                  strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
'                               "VALUES ('" & strCP09 & "'," & strNP22 & ",'馬德里指定國家','" & strNA2 & "','')"
'                  cnnConnection.Execute strSql
'
'                  strExc(0) = CompWorkDay(5, strSrvDate(1))
'                  strExc(1) = GetNP08(strCP09, strNP07, strTM01, strTM02, strTM03, strTM04)
'                  '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
'                  If Val(strExc(1)) <= Val(strExc(0)) Then
'                     PUB_Cache2Letter strCP09, strNP22, False, False
'                  End If
'               End If
               '2019/11/29 END
        Else
                ' 申請國家為台灣
                If strTM10 < "010" Then
                    ' 申請人國籍為台灣
                    '2009/12/1 modify BY SONIA 改以定稿語文m_strLanguage判斷
                    'If m_strTM23Nation < "010" Then
                    If m_strLanguage = "1" Then
                       If strDate <= strSysDate Then '逾期延展
                           ' 清除定稿例外欄位檔原有資料
                           'Modify By Sindy 2009/04/27 處理狀況原為03改為12
                           'Modified by Lydia 2019/05/20 "12"=>strET03
                           EndLetter "10", strCP09, strET03, strUserNum
                           'add by nickc 2008/04/25 案件回覆單
                           'Modify By Sindy 2009/04/27 處理狀況原為03改為12
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                                    "'下一程序','" & strNP07 & "')"
                           cnnConnection.Execute strSql
                          
                           'Add By Cheng 2002/09/18
                           ' 服務專線
                            'Modify By Cheng 2003/11/21
            '               strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            '                        "VALUES ('" & "10" & "','" & strCP09 & "','" & "03" & "','" & strUserNum & "'," & _
            '                        "'服務專線','" & IIf(strCP12 >= "S2" And strCP12 <= "S29", "04-3270288", IIf(strCP12 >= "S3" And strCP12 <= "S39", "06-2743866", IIf(strCP12 >= "S4" And strCP12 <= "S49", "07-2363602", "02-25061023 轉"))) & "')"
                           'Modify By Sindy 2009/04/27 處理狀況原為03改為12
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                                    "'服務專線','" & PUB_GetServiceLine(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "')"
                           cnnConnection.Execute strSql
                           'end 2019/05/20
                           'Add By Cheng 2002/11/21
'cancel by sonia 2016/2/24 定稿設計改用<業務員>,因中所新人特別+主管名字
'                           '下一程序業務員
'                            'Modify By Cheng 2003/07/03
'            '               strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            '                        "VALUES ('" & "10" & "','" & strCP09 & "','" & "03" & "','" & strUserNum & "'," & _
'            '                        "'下一程序業務員','" & GetStaffName(strNP10) & "')"
'                           'Modify By Sindy 2009/04/27 處理狀況原為03改為12
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "10" & "','" & strCP09 & "','" & "12" & "','" & strUserNum & "'," & _
'                                    "'下一程序業務員','" & GetStaffName(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "')"
'                           cnnConnection.Execute strSql

                           'Added by Lydia 2019/08/22 計算雙面列印定稿的可印範圍;
                           Call SetTMGoodsDetail(strTM01, strTM02, strTM03, strTM04, strCP09, "10", strET03)
                           
                       Else
                          'Modify By Sindy 2009/04/17 定稿別10處理狀況01改為08
                          ' 清除定稿例外欄位檔原有資料
                          'Modified by Lydia 2019/05/20 "08"=>strET03
                          EndLetter "10", strCP09, strET03, strUserNum
                          'add by nickc 2008/04/25 案件回覆單
                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                                 "'下一程序','" & strNP07 & "')"
                          cnnConnection.Execute strSql
                          '2009/04/17 End
                          
                           'Add By Cheng 2002/09/18
                           ' 服務專線
                            'Modify By Cheng
            '               strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            '                        "VALUES ('" & "10" & "','" & strCP09 & "','" & "01" & "','" & strUserNum & "'," & _
            '                        "'服務專線','" & IIf(strCP12 >= "S2" And strCP12 <= "S29", "04-3270288", IIf(strCP12 >= "S3" And strCP12 <= "S39", "06-2743866", IIf(strCP12 >= "S4" And strCP12 <= "S49", "07-2363602", "02-25061023 轉"))) & "')"
                           'Modify By Sindy 2009/04/17 定稿別10處理狀況01改為08
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                                    "'服務專線','" & PUB_GetServiceLine(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "')"
                           cnnConnection.Execute strSql
                           'Add By Cheng 2002/11/21
'cancel by sonia 2016/2/24 定稿設計改用<業務員>,因中所新人特別+主管名字
'                           '下一程序業務員
'                            'Modify By Cheng 2003/07/03
'            '               strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            '                        "VALUES ('" & "10" & "','" & strCP09 & "','" & "01" & "','" & strUserNum & "'," & _
'            '                        "'下一程序業務員','" & GetStaffName(strNP10) & "')"
'                           'Modify By Sindy 2009/04/17 定稿別10處理狀況01改為08
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "10" & "','" & strCP09 & "','" & "08" & "','" & strUserNum & "'," & _
'                                    "'下一程序業務員','" & GetStaffName(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "')"
'                           cnnConnection.Execute strSql
                           'Add By Sindy 2011/2/1
                           '法定期限
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                                    "'法定期限','" & GetNP09(strCP09, strNP07) & "')"
                           cnnConnection.Execute strSql
                           '2011/2/1 End
                           'end 2019/05/20
                           
                           'Added by Lydia 2019/08/22 計算雙面列印定稿的可印範圍;
                           Call SetTMGoodsDetail(strTM01, strTM02, strTM03, strTM04, strCP09, "10", strET03)
                       End If
                    ElseIf m_strLanguage = "2" Then
'                        ' 清除定稿例外欄位檔原有資料
'                        EndLetter "10", strCP09, "00", strUserNum
'                        'add by nickc 2008/04/25 案件回覆單
'                        strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "10" & "','" & strCP09 & "','" & "00" & "','" & strUserNum & "'," & _
'                                 "'下一程序','" & strNP07 & "')"
'                        cnnConnection.Execute strSQL
                        
                        'Modify By Sindy 2009/04/16
                        ' 清除定稿例外欄位檔原有資料
                        'Modified by Lydia 2019/05/20 "09"=>strET03
                        EndLetter "10", strCP09, strET03, strUserNum
                        'add by nickc 2008/04/25 案件回覆單
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                                 "'下一程序','" & strNP07 & "')"
                        cnnConnection.Execute strSql
                        '2009/04/16 End
                        'Add By Sindy 2013/5/29
                        If m_CU15 = "0" Then '個人
                           strTmp = "身分證"
                        Else
                           strTmp = "營業執照"
                        End If
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                                 "'證照文件','" & strTmp & "')"
                        cnnConnection.Execute strSql
                        '2013/5/29 End
                        'Add By Sindy 2014/10/28
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                                 "'費用','" & Format(1800 + (1500 * (dblTMKindCnt - 1)), "##,##0") & "')"
                        cnnConnection.Execute strSql
                        '2014/10/28 END
                        'end 2019/05/20
                    End If
                    
                ' 申請國家為大陸
                ElseIf strTM10 = "020" Then
                   If m_strLanguage = "1" Then '中文
                     ' 清除定稿例外欄位檔原有資料
                     'Modified by Lydia 2019/05/20 "02"=>strET03
                     EndLetter "10", strCP09, strET03, strUserNum
                     'Modify By Sindy 2025/7/14
                     If strTM01 <> "TM" Then
                     '2025/7/14 END
                        'add by nickc 2008/04/25 案件回覆單
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                                 "'下一程序','" & strNP07 & "')"
                        cnnConnection.Execute strSql
                        'Add By Cheng 2002/09/18
                        ' 服務專線
                        'Modify By Cheng 2003/11/21
            '            strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            '                     "VALUES ('" & "10" & "','" & strCP09 & "','" & "02" & "','" & strUserNum & "'," & _
            '                     "'服務專線','" & IIf(strCP12 >= "S2" And strCP12 <= "S29", "04-3270288", IIf(strCP12 >= "S3" And strCP12 <= "S39", "06-2743866", IIf(strCP12 >= "S4" And strCP12 <= "S49", "07-2363602", "02-25061023 轉"))) & "')"
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                                 "'服務專線','" & PUB_GetServiceLine(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "')"
                        cnnConnection.Execute strSql
                        'Add By Cheng 2002/11/21
'cancel by sonia 2016/2/24 定稿設計改用<業務員>,因中所新人特別+主管名字
'                        '下一程序業務員
'                        'Modify By Cheng 2003/07/03
'            '            strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            '                     "VALUES ('" & "10" & "','" & strCP09 & "','" & "02" & "','" & strUserNum & "'," & _
'            '                     "'下一程序業務員','" & GetStaffName(strNP10) & "')"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "10" & "','" & strCP09 & "','" & "02" & "','" & strUserNum & "'," & _
'                                 "'下一程序業務員','" & GetStaffName(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "')"
'                        cnnConnection.Execute strSql
                        'Add By Sindy 2012/1/9 以申請人抓所有未閉卷未銷卷之大陸商標案,若有申請地址與客戶檔中文地址不同,才要印此段
                        'modify by sonia 2017/11/16 +CU15個人或公司
                        'Remove by Lydia 2020/05/19 大陸案改與台灣案一致化
                        'strSql = "select T1.TM01,T1.TM02,T1.TM03,T1.TM04,T1.TM23,T1.TM24,CU23,T1.TM10,T1.TM29,T1.TM57,T1.TM73,DECODE(CU15,'0','台端','1','貴公司','貴單位') CU15 " & _
                                 "from trademark T1,customer,(select TM23 from trademark where TM01='" & strTM01 & "' and TM02='" & strTM02 & "' and TM03='" & strTM03 & "' and TM04='" & strTM04 & "') T2 " & _
                                 "Where t1.tm23=t2.tm23 and T1.TM10='020' " & _
                                 "and T1.TM29 is null and T1.TM57 is null and T1.TM73 is null " & _
                                 "and CU01=substr(rtrim(ltrim(T2.TM23)),1,8) and CU02=substr(rtrim(ltrim(T2.TM23)),9,1) " & _
                                 "and T1.TM24 is not null " & _
                                 "and CU23 is not null " & _
                                 "and replace(replace(replace(rtrim(ltrim(T1.TM24)),'　',null),chr(10),null),chr(13),null) <> replace(rtrim(ltrim(CU23)),'　',null) " & _
                                 "and replace(replace(replace(rtrim(ltrim(T1.TM24)),'　',null),chr(10),null),chr(13),null) <> rtrim(CU112)||replace(rtrim(ltrim(CU23)),'　',null) " & _
                                 "order by T1.TM01,T1.TM02,T1.TM03,T1.TM04 asc "
                        'intI = 1
                        'Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        'If intI = 1 Then
                        '   If RsTemp.RecordCount > 0 Then
                        '      'modify by sonia 2017/11/16桂英通知修改內容,並加入判斷CU15個人或公司
                        '      'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        '               "VALUES ('" & "10" & "','" & strCP09 & "','" & "02" & "','" & strUserNum & "'," & _
                        '               "'申請人地址變更告知','依本所資料顯示，貴公司的大陸商標有以不同的地址申請註冊。依大陸商標法規定，若申請人的地址變更，必須所申請的大陸商標全部皆須辦理商標變更地址登記，否則將構成商標得撤銷的事由，特此告知。')"
                        '      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        '               "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                        '               "'申請人地址變更告知','" & Chr(13) & "　　如　" & RsTemp("CU15") & "的地址已變更，可在辦理延展時，另案申請變更地址。依大陸商標法規定，若申請人的地址變更，必須所申請的大陸商標全部皆須辦理商標變更地址登記，否則將構成商標得撤銷的事由，特此告知。')"
                        '      cnnConnection.Execute strSql
                        '   End If
                        'End If
                        ''2012/1/9 End
                        ''end 2019/05/20
                        'end 2020/05/19
                        
                        'Addeded by Lydia 2020/05/19 大陸案改與台灣案一致化
                        '法定期限
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                                 "'法定期限','" & GetNP09(strCP09, strNP07) & "')"
                        cnnConnection.Execute strSql
                        '計算雙面列印定稿的可印範圍;
                        Call SetTMGoodsDetail(strTM01, strTM02, strTM03, strTM04, strCP09, "10", strET03)
                        'end 2020/05/19
                     End If
                        
                   'Modify By Sindy 2010/11/12
                   Else '外->大英文定稿
                        '取得美金匯率
                        strSql = "SELECT usxr02 FROM usxrate order by usxr01 desc "
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           dbl_usxrate = RsTemp("usxr02")
                        End If
                        '取得費用
                        'Add By Sindy 2012/8/22 加註 frm210138 也有此費用的計算,若有異動時,須一併改寫
                        strSql = "SELECT cf08+(cf13*1000) FROM casefee where cf01='" & strTM01 & "' and cf02='" & strTM10 & "' and cf03='" & strNP07 & "' "
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           dbl_fee = RsTemp.Fields(0)
                        End If
                        If m_TM44 = "" Then '無代理人
                           ' 清除定稿例外欄位檔原有資料
                           'Modified by Lydia 2019/05/20 "15"=>strET03
                           EndLetter "10", strCP09, strET03, strUserNum
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                                    "'費用','" & Format(dbl_fee, "##,##0") & "')"
                           cnnConnection.Execute strSql
                           '2011/5/6 modify by sonia 美金取整數,小數捨去
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                                    "'費用美金','" & IIf(dbl_usxrate = 0, 0, Format(dbl_fee / dbl_usxrate, "##,##0")) & "')"
                           cnnConnection.Execute strSql
                           'end 2019/05/20
                        Else '有代理人
                           ' 清除定稿例外欄位檔原有資料
                           'Modified by Lydia 2019/05/20 "16"=>strET03
                           EndLetter "10", strCP09, strET03, strUserNum
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                                    "'費用','" & Format(dbl_fee, "##,##0") & "')"
                           cnnConnection.Execute strSql
                           '2011/5/6 modify by sonia 美金取整數,小數捨去
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                                    "'費用美金','" & IIf(dbl_usxrate = 0, 0, Format(dbl_fee / dbl_usxrate, "##,##0")) & "')"
                           cnnConnection.Execute strSql
                           'end 2019/05/20
                        End If
                   '2010/11/12 End
                   End If
                End If
        End If
        
      ' 刊登廣告
      Case "702":
        ' 申請國家為大陸
        If strTM10 = "020" Then
           ' 清除定稿例外欄位檔原有資料
           'Modified by Lydia 2019/05/20 "04"=>strET03
           EndLetter "10", strCP09, strET03, strUserNum
            'add by nickc 2008/04/25 案件回覆單
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                     "'下一程序','" & strNP07 & "')"
            cnnConnection.Execute strSql
           
            'Add By Cheng 2002/09/18
            ' 服務專線
            'Modify By Cheng 2003/11/21
'            strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                     "VALUES ('" & "10" & "','" & strCP09 & "','" & "04" & "','" & strUserNum & "'," & _
'                     "'服務專線','" & IIf(strCP12 >= "S2" And strCP12 <= "S29", "04-3270288", IIf(strCP12 >= "S3" And strCP12 <= "S39", "06-2743866", IIf(strCP12 >= "S4" And strCP12 <= "S49", "07-2363602", "02-25061023 轉"))) & "')"
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                     "'服務專線','" & PUB_GetServiceLine(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "')"
            cnnConnection.Execute strSql
            'end 2019/05/20
            'Add By Cheng 2002/11/21
'cancel by sonia 2016/2/24 定稿設計改用<業務員>,因中所新人特別+主管名字
'            '下一程序業務員
'            'Modify By Cheng 2003/07/03
''            strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                     "VALUES ('" & "10" & "','" & strCP09 & "','" & "04" & "','" & strUserNum & "'," & _
''                     "'下一程序業務員','" & GetStaffName(strNP10) & "')"
'            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                     "VALUES ('" & "10" & "','" & strCP09 & "','" & "04" & "','" & strUserNum & "'," & _
'                     "'下一程序業務員','" & GetStaffName(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "')"
'            cnnConnection.Execute strSql
        End If
        
      'Add By Sindy 2009/06/15
      ' 被異議續展
      Case "109":
        ' 申請國家為大陸
        If strTM10 = "020" Then
'           If TXT1(9) = "97038" Then '臨時用
'                ' 清除定稿例外欄位檔原有資料
'                EndLetter "10", strCP09, "14", strUserNum
'                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                         "VALUES ('" & "10" & "','" & strCP09 & "','" & "14" & "','" & strUserNum & "'," & _
'                         "'下一程序','" & strNP07 & "')"
'                cnnConnection.Execute strSql
'                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                         "VALUES ('" & "10" & "','" & strCP09 & "','" & "14" & "','" & strUserNum & "'," & _
'                         "'服務專線','" & PUB_GetServiceLine(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "')"
'                cnnConnection.Execute strSql
'                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                         "VALUES ('" & "10" & "','" & strCP09 & "','" & "14" & "','" & strUserNum & "'," & _
'                         "'下一程序業務員','" & GetStaffName(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "')"
'                cnnConnection.Execute strSql
'           Else
                ' 清除定稿例外欄位檔原有資料
                'Modified by Lydia 2019/05/20 "13"=>strET03
                EndLetter "10", strCP09, strET03, strUserNum
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                         "'下一程序','" & strNP07 & "')"
                cnnConnection.Execute strSql
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                         "'服務專線','" & PUB_GetServiceLine(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "')"
                cnnConnection.Execute strSql
                'end 2019/05/20
'cancel by sonia 2016/2/24 定稿設計改用<業務員>,因中所新人特別+主管名字
'                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                         "VALUES ('" & "10" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & "'," & _
'                         "'下一程序業務員','" & GetStaffName(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "')"
'                cnnConnection.Execute strSql
'           End If
        End If
      '2009/06/15 End
      
      ' 繳年費
      Case "708":
        ' 系統別為TB
        If strTM01 = "TB" Then
            ' 清除定稿例外欄位檔原有資料
            'Modified by Lydia 2019/05/20 "05"=>strET03
            EndLetter "10", strCP09, strET03, strUserNum
            'add by nickc 2008/04/25 案件回覆單
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                        "'下一程序','" & strNP07 & "')"
            cnnConnection.Execute strSql
            ' 本所期限
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                     "'本所期限','" & GetNP08(strCP09, strNP07) & "')"
            cnnConnection.Execute strSql
            '法定期限
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                     "'法定期限','" & GetNP09(strCP09, strNP07) & "')"
            cnnConnection.Execute strSql
            
            'Added by Lydia 2017/08/21 相關總收文號若為繳年費708時,年費年度=CP53+1,規費=發文規費
            If strCP10 = "708" Then
                If Val(strCP53) = 0 Then
                   strExc(1) = "2"
                Else
                   strExc(1) = Val(strCP53) + 1
                End If
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                         "'條碼年費年度','" & strExc(1) & "')"
                cnnConnection.Execute strSql
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                         "'條碼年費規費','" & Format(IIf(Val(strExc(1)) >= 4, Val(strCP84), Val(strCP84) * 0.9), "0") & "')"
                cnnConnection.Execute strSql
                Select Case strExc(1)
                    Case "2"
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                                 "'年費折扣說明','該會為回饋先進廠商，年費之收取，第2期遞減百分之十，即')"
                        cnnConnection.Execute strSql
                    Case "3"
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                                 "'年費折扣說明','該會為回饋先進廠商，年費之收取，第3期遞減百分之二十，即')"
                        cnnConnection.Execute strSql
                    Case Else
                End Select
            Else
                '非繳年費708時,年費年度=2,規費=19845
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                         "'條碼年費年度','2')"
                cnnConnection.Execute strSql
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                         "'條碼年費規費','19845')"
                cnnConnection.Execute strSql
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                         "'年費折扣說明','該會為回饋先進廠商，年費之收取，第2期遞減百分之十，即')"
                cnnConnection.Execute strSql
            End If
            'end 2017/08/21
        End If
        
      'add by nickc 2006/05/18 新增第二期註冊費定稿
      Case "716":
           'add by nickc 2006/05/30 加入新的大陸定稿 申請人國籍為大陸
           Dim oState As String
           '2009/12/1 modify BY SONIA 改以定稿語文m_strLanguage判斷
           'If m_strTM23Nation > "010" Then '大-->台
           'Modified by Lydia 2019/05/20 改成傳入處理狀況
           'If m_strLanguage = "2" Then '大-->台
           '     'Modify By Sindy 2009/04/22
           '     'oState = "07"
           '     oState = "11"
           'Else
           '     'Modify By Sindy 2009/04/22
           '     'oState = "06"
           '     oState = "10"
           'End If
           oState = strET03
           'end 2019/05/20
           ' 清除定稿例外欄位檔原有資料
           EndLetter "10", strCP09, oState, strUserNum
           
            'Add By Sindy 2009/04/22
            '2009/3/4 ADD BY SONIA 抓商品類別數
            intCnt = GetTMKindCnt(strTM01, strTM02, strTM03, strTM04)
            CheckOC3
            strSql = "select * from casefee where cf01='" & strTM01 & "' and cf02='" & strTM10 & "' and cf03 in ('715','716','717') order by cf03 "
            AdoRecordSet3.CursorLocation = adUseClient
            AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If AdoRecordSet3.RecordCount <> 0 Then
                AdoRecordSet3.MoveFirst
                Do While Not AdoRecordSet3.EOF
                    '2009/3/4 MODIFY BY SONIA CF08規費要*商品類別數
                    Select Case AdoRecordSet3.Fields("cf03").Value
                    Case "715"
                        o71508 = AdoRecordSet3.Fields("cf08").Value * intCnt
                        o71513 = AdoRecordSet3.Fields("cf13").Value
                    Case "716"
                        o71608 = AdoRecordSet3.Fields("cf08").Value * intCnt
                        o71613 = AdoRecordSet3.Fields("cf13").Value
                    Case "717"
                        o71708 = AdoRecordSet3.Fields("cf08").Value * intCnt
                        o71713 = AdoRecordSet3.Fields("cf13").Value
                    Case Else
                    End Select
                    AdoRecordSet3.MoveNext
                Loop
            End If
            CheckOC3
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "10" & "','" & strCP09 & "','" & oState & "','" & strUserNum & _
                     "','錢715','" & Format(o71513 * 1000 + o71508, "###,###,##0") & "')"
            cnnConnection.Execute strSql
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "10" & "','" & strCP09 & "','" & oState & "','" & strUserNum & _
                     "','錢716','" & Format(o71613 * 1000 + o71608, "###,###,##0") & "')"
            cnnConnection.Execute strSql
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "10" & "','" & strCP09 & "','" & oState & "','" & strUserNum & _
                     "','錢717','" & Format(o71713 * 1000 + o71708, "###,###,##0") & "')"
            cnnConnection.Execute strSql
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "10" & "','" & strCP09 & "','" & oState & "','" & strUserNum & _
                     "','錢71508','" & Format(o71508, "###,###,##0") & "')"
            cnnConnection.Execute strSql
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "10" & "','" & strCP09 & "','" & oState & "','" & strUserNum & _
                     "','錢71608','" & Format(o71608, "###,###,##0") & "')"
            cnnConnection.Execute strSql
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "10" & "','" & strCP09 & "','" & oState & "','" & strUserNum & _
                     "','錢71708','" & Format(o71708, "###,###,##0") & "')"
            cnnConnection.Execute strSql
            '2009/04/22 End
           
            'add by nickc 2008/04/25 案件回覆單
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "10" & "','" & strCP09 & "','" & oState & "','" & strUserNum & "'," & _
                     "'下一程序','" & strNP07 & "')"
            cnnConnection.Execute strSql
            ' 服務專線
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "10" & "','" & strCP09 & "','" & oState & "','" & strUserNum & "'," & _
                     "'服務專線','" & PUB_GetServiceLine(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "')"
            cnnConnection.Execute strSql
'cancel by sonia 2016/2/24 定稿設計改用<業務員>,因中所新人特別+主管名字
'            '下一程序業務員
'            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                     "VALUES ('" & "10" & "','" & strCP09 & "','" & oState & "','" & strUserNum & "'," & _
'                     "'下一程序業務員','" & GetStaffName(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "')"
'            cnnConnection.Execute strSql
            'Add By Sindy 2012/4/20
            '本所期限
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "10" & "','" & strCP09 & "','" & oState & "','" & strUserNum & "'," & _
                     "'本所期限','" & GetNP08(strCP09, strNP07) & "')"
            cnnConnection.Execute strSql
            '法定期限
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "10" & "','" & strCP09 & "','" & oState & "','" & strUserNum & "'," & _
                     "'法定期限','" & GetNP09(strCP09, strNP07) & "')"
            cnnConnection.Execute strSql
            '2012/4/20 End
            
      '2010/6/9 add by sonia TF美國使用宣誓
      'Modify By Sindy 2010/6/11
      Case "105":
            If strTM10 = "101" Then '美國
               strET03 = "01" 'Add by Sindy 2019/11/6
               If Val(Trim(Me.textMoney.Text)) <> 0 Then
                  'Add By Sindy 2012/6/29
                  strTmp = ""
                  If strTM01 = "TF" Then
                     '取得商品類別
                     strSql = "SELECT tm09 from trademark " & _
                              "where tm01='" & strTM01 & "' and tm02='" & strTM02 & "' and tm04='" & strTM04 & "' " & _
                              "and tm29 is null "
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        RsTemp.MoveFirst
                        Do While Not RsTemp.EOF
                           strTmp = strTmp & "," & RsTemp.Fields(0)
                           RsTemp.MoveNext
                        Loop
                     End If
                     If strTmp <> "" Then strTmp = Right(strTmp, Len(strTmp) - 1)
                  End If
                  '2012/6/29 End
                  ' 清除定稿例外欄位檔原有資料
                  'Modified by Lydia 2019/05/20 "01"=>strET03
                  EndLetter "10", strCP09, strET03, strUserNum
                  'Modify By Sindy 2019/10/18 + 信函收文號
                  PUB_AddLetterCache strCP09, strNP22, strCP09, "10", strET03, , IIf(strSrvDate(1) >= T商標電子化啟用日, strLD18, "")
                  'end 2019/05/20
                  '費用
                  strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                               "VALUES ('" & strCP09 & "'," & strNP22 & ",'費用','" & Me.textMoney.Text & "','Y')"
                  cnnConnection.Execute strSql
                  '本所期限
                  'Modified by Lydia 2017/10/23 LCV05拿掉Y (影響報價備註CP64)
                  strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                               "VALUES ('" & strCP09 & "'," & strNP22 & ",'本所期限','" & GetNP08(strCP09, strNP07, strTM01, strTM02, strTM03, strTM04) & "','')"
                  cnnConnection.Execute strSql
                  '法定期限
                  'Modified by Lydia 2017/10/23 LCV05拿掉Y (影響報價備註CP64)
                  strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                               "VALUES ('" & strCP09 & "'," & strNP22 & ",'法定期限','" & GetNP09(strCP09, strNP07, strTM01, strTM02, strTM03, strTM04) & "','')"
                  cnnConnection.Execute strSql
                  'Add By Sindy 2012/6/29
                  'Modified by Lydia 2017/10/23 LCV05拿掉Y (影響報價備註CP64)
                  strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                               "VALUES ('" & strCP09 & "'," & strNP22 & ",'馬德里商品類別','" & strTmp & "','')"
                  cnnConnection.Execute strSql
                  '2012/6/29 End
                  strExc(0) = CompWorkDay(5, strSrvDate(1))
                  strExc(1) = GetNP08(strCP09, strNP07, strTM01, strTM02, strTM03, strTM04)
                  '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                  If Val(strExc(1)) <= Val(strExc(0)) Then
                     PUB_Cache2Letter strCP09, strNP22, False, False
                  End If
               End If
            'Add By Sindy 2015/2/25 要報價但沒有定稿時提醒
            Else
               If Val(Trim(Me.textMoney.Text)) <> 0 Then
                  MsgBox "本案要報價但沒有系統的定稿，請注意！", vbExclamation
               End If
            '2015/2/25 END
            End If
      'Add by Amy 2018/08/31
      Case "1729"
            'Moidfy by Amy 2018/10/03 +非英文定稿
            If m_strLanguage <> "3" Then
                ' 清除定稿例外欄位檔原有資料
                'Modified by Lydia 2019/05/20 "00"=>strET03
                EndLetter "10", strCP09, strET03, strUserNum
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                            "'服務專線','" & PUB_GetServiceLine(PUB_GetAKindSalesNo(strTM01, strTM02, strTM03, strTM04)) & "')"
                cnnConnection.Execute strSql
                'end 2091/05/20
            End If
   
      'Add By Sindy 2025/7/15
      Case "809" '進出口監視備案
         ' 清除定稿例外欄位檔原有資料
         EndLetter "10", strCP09, strET03, strUserNum
         ' 申請人國籍為台灣
         If m_strLanguage = "1" Then
         ' 申請人國籍非台灣
         Else
         End If
      '2025/7/15 END
   End Select
   
   'Add By Sindy 2020/12/14
   m_MySt(8) = DBDATE(txt1(17))
   If strET03 <> "" Then
      If strET03 = "09" Then '大->台延展
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                     "'定稿日期','" & ExceptFieldData("定稿日期/中西") & "')"
      ElseIf strET03 = "15" Or strET03 = "16" Or strET03 = "18" Or strET03 = "20" Then '英文定稿
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                     "'定稿日期','" & ExceptFieldData("定稿日期/英") & "')"
      Else
         'Modify By Sindy 2025/4/21 ExceptFieldData("定稿日期/中") 改為 ExceptFieldData("定稿日期/中2")
         'Modify by Sindy 2025/7/1 大陸撤三定稿日期維持完全顯示
         If Option1(1).Value = True Then
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                        "'定稿日期','" & ExceptFieldData("定稿日期/中") & "')"
         Else
         '2025/7/1 END
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "10" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                        "'定稿日期','" & ExceptFieldData("定稿日期/中2") & "')"
         End If
      End If
      cnnConnection.Execute strSql
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Modify By Sindy 2010/6/11 增加NP22
'Modify By Sindy 2019/10/24 + , ByVal strLD18 As String : 信函收文號
Private Sub PrintLetter(ByVal strNP07 As String, ByVal strTM01 As String, ByVal strTM10 As String, _
   ByVal strDate As String, ByVal strSysDate As String, ByVal strCP09 As String, _
   ByVal strCP12 As String, ByVal strNP10 As String, _
   ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String, ByVal strNP22 As String, _
   ByVal strLD18 As String)
   
'Add By Sindy 2012/1/13
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/13 End
Dim intR As Integer, stLP26 As String 'Add By Sindy 2025/5/21
   
   'Add By Sindy 2012/1/13
   ET01 = "10"
   ET02 = strCP09
   bolEdit = False
   iCopy = 0
   '2012/1/13 End
   
   '2009/12/1 modify BY SONIA 不以申請人國籍m_strTM23Nation判斷,改以定稿語文m_strLanguage判斷
   Select Case strNP07
      ' 延展
      Case "102":
        'add by nickc 2006/05/05
        If strTM01 = "TF" Then
            'Modify By Sindy 2019/11/29 改為報價定稿
            ' 清除定稿例外欄位檔原有資料
'            NowPrint strCP09, "10", "00", False, strUserNum, 0
            'Modify By Sindy 2021/3/18 TF續展的報價定稿取消,改為單純續展期限通知函
            ET03 = "00" 'Modify By Sindy 2012/1/13
        Else
            ' 申請國家為台灣
            If strTM10 < "010" Then
                ' 申請人國籍為台灣
                '2009/12/1 modify BY SONIA 改以定稿語文m_strLanguage判斷
                'If m_strTM23Nation < "010" Then
                If m_strLanguage = "1" Then
                    If strDate <= strSysDate Then '逾期延展
                       ' 列印定稿
                       'Modify By Sindy 2009/04/27 處理狀況原為03改為12
'                           NowPrint strCP09, "10", "12", False, strUserNum, 0
                       'Modified by Lydia 2019/05/20 雙面列印=橫式雙面
                       'ET03 = "12" 'Modify By Sindy 2012/1/13
                       ET03 = "22"
                    Else
                       ' 列印定稿
                       'NowPrint strCP09, "10", "01", False, strUserNum, 0
'                           NowPrint strCP09, "10", "08", False, strUserNum, 0 'Modify By Sindy 2009/04/17
                       'Modified by Lydia 2019/05/20 雙面列印=橫式雙面
                       'ET03 = "08" 'Modify By Sindy 2012/1/13
                       ET03 = "21"
                    End If
                ' 申請人國籍非台灣
                ElseIf m_strLanguage = "2" Then
                    ' 列印定稿
                    '2007/11/2 MODIFY BY SONIA 改為二份
                    'NowPrint strCP09, "10", "05", False, strUserNum, 0
                    'Modify By Sindy 2009/04/17
                    'NowPrint strCP09, "10", "05", False, strUserNum, 0, "", False, "", 2
'                        NowPrint strCP09, "10", "09", False, strUserNum, 0, "", False, "", 2
                    ET03 = "09" 'Modify By Sindy 2012/1/13
                    'Modify By Sindy 2019/11/21 Mark
                    'iCopy = 2
                    '2019/11/21 END
                'ADD BY SONIA 2014/6/27 T-140539 英文無代理人
                Else
                    'Modify By Sindy 2015/3/3
'                        If m_TM44 = "" Then '無代理人
'                           ET03 = "17"
'                           iCopy = 2
'                        Else
                       ET03 = "18"
                       'Modify By Sindy 2019/11/21 Mark
                       'iCopy = 2
                       '2019/11/21 END
                    '2015/3/3 END
'                        End If
                'end 2014/6/27
                End If
            ' 申請國家為大陸
            ElseIf strTM10 = "020" Then
               If m_strLanguage = "1" Then '中文
                    ' 列印定稿
'                        NowPrint strCP09, "10", "02", False, strUserNum, 0
                    'Modified by Lydia 2020/05/19 大陸案改與台灣案一致化
                    'ET03 = "02" 'Modify By Sindy 2012/1/13
                    'Modify By Sindy 2025/7/14
                    If strTM01 = "TM" Then '中國 102.海關備案延展
                       ET03 = "01"
                    Else
                    '2025/7/14 END
                       ET03 = "23"
                    End If
               'Modify By Sindy 2010/11/12
               Else '外->大英文定稿
                    If m_TM44 = "" Then '無代理人
'                           NowPrint strCP09, "10", "15", False, strUserNum, 0
                       ET03 = "15" 'Modify By Sindy 2012/1/13
                    Else '有代理人
'                           NowPrint strCP09, "10", "16", False, strUserNum, 0
                       ET03 = "16" 'Modify By Sindy 2012/1/13
                    End If
               '2010/11/12 End
               End If
            End If
        End If
        
      ' 刊登廣告
      Case "702":
        ' 申請國家為大陸
        If strTM10 = "020" Then
           ' 列印定稿
'           NowPrint strCP09, "10", "04", False, strUserNum, 0
            ET03 = "04" 'Modify By Sindy 2012/1/13
        End If
      
      'Add By Sindy 2009/06/15
      ' 被異議續展
      Case "109":
         ' 申請國家為大陸
         If strTM10 = "020" Then
           ' 列印定稿
'           If TXT1(9) = "97038" Then '臨時用
'               NowPrint strCP09, "10", "14", False, strUserNum, 0
'           Else
'               NowPrint strCP09, "10", "13", False, strUserNum, 0
               ET03 = "13" 'Modify By Sindy 2012/1/13
'           End If
         End If
         '2009/06/15 End
            
      ' 繳年費
      Case "708":
         ' 系統別為TB
         If strTM01 = "TB" Then
            ' 列印定稿
            'Modify By Cheng 2003/01/14
            '開Word修改
'           NowPrint strCP09, "10", "05", False, strUserNum, 0
'           NowPrint strCP09, "10", "05", True, strUserNum, 0
            ET03 = "05" 'Modify By Sindy 2012/1/13
            'bolEdit = True 'Add By Sindy 2012/1/13 'Mark by Lydia 2017/08/24 因為定稿內已計算好規費,預設不用開Word
         End If
         
'      'add by nickc 2006/05/18 新增第二期註冊費定稿
'      Case "716":
'        'add by nickc 2006/05/30 加入新的大陸定稿 申請人國籍為大陸
'        '2007/11/2 MODIFY BY SONIA 改判斷申請人國籍非台灣
'        'If m_strTM23Nation = "020" Then
'        '2009/12/1 modify BY SONIA 改以定稿語文m_strLanguage判斷
'        'If m_strTM23Nation > "010" Then '大-->台
'        If m_strLanguage = "2" Then
'           '2007/11/2 MODIFY BY SONIA 改為二份
'           'NowPrint strCP09, "10", "07", False, strUserNum, 0
'           'Modify By Sindy 2009/04/22
'           'NowPrint strCP09, "10", "07", False, strUserNum, 0, "", False, "", 2
''           NowPrint strCP09, "10", "11", False, strUserNum, 0, "", False, "", 2
'            ET03 = "11" 'Modify By Sindy 2012/1/13
'            iCopy = 2
'        Else
'           'Modify By Sindy 2009/04/22
'           'NowPrint strCP09, "10", "06", False, strUserNum, 0
''           NowPrint strCP09, "10", "10", False, strUserNum, 0
'            ET03 = "10" 'Modify By Sindy 2012/1/13
'        End If
      '2010/6/9 ADD BY SONIA 使用宣誓有定稿但程式沒寫
      'Modify By Sindy 2010/6/11
'cancel by sonia 2014/6/26 已改報價定稿TF-000580-1-03
'      Case "105"
'         If strTM10 = "101" Then '美國
'            If Val(Trim(Me.textMoney.Text)) <> 0 Then
''               NowPrint strCP09, "10", "01", False, strUserNum, 0
'               ET03 = "01" 'Modify By Sindy 2012/1/13
'            End If
'         End If
'end 2014/6/26
      '2010/6/9 END
      
      'Add by Amy 2018/08/31 +大陸撤三
      Case "1729"
         'Modify by Amy 2018/10/03 +英文定稿
         If m_strLanguage = "3" Then
             ET03 = "20"
         Else
             ET03 = "19"
         End If
         
      'Add By Sindy 2025/7/15
      Case "809" '進出口監視備案
         ' 申請人國籍為台灣
         If m_strLanguage = "1" Then
            ET03 = "02" '台灣 809.進出口監視備案(台至台)
         ' 申請人國籍非台灣
         Else
            ET03 = "03" '台灣 809.進出口監視備案(台至大)MCT
         End If
      '2025/7/15 END
   End Select
   
   'Added by Lydia 2019/08/22  開啟Word
   'Mark by Lydia 2019/08/22 直接算和丟Word算105份約2分鐘,Word算略慢一些
'   If m_bWord = False And strNP07 = "102" And strTM01 <> "TF" And strTM10 < "010" And (ET03 = "21" Or ET03 = "22") Then
'        If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = False Then Exit Sub
'
'        With g_WordAp.Application
'           ' 橫式雙面LD12=8的目前版面,若有調整需要一併修改
'            .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
'            .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
'            .Selection.PageSetup.TopMargin = .CentimetersToPoints(4)
'            '下邊界加大
'            .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2.5)
'            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
'            .Selection.WholeStory
'
'            '固定行高
'            .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
'            .Selection.ParagraphFormat.LineSpacing = 15
'            .Selection.Font.Name = "細明體"
'            .Selection.Font.Size = 14
'        End With
'        m_bWord = True
'   End If
   
   'Modify By Sindy 2010/6/11 增加NP22
   'Modified by Lydia 2019/05/20 +處理狀況ET03 ; 並且從模組最上方移下來
   'Modify By Sindy 2019/11/6 + strLD18 : 信函收文號
   InsExpField strNP07, strTM01, strTM10, strDate, strSysDate, strCP09, strCP12, strNP10, strTM02, strTM03, strTM04, strNP22, ET03, strLD18
   
   'Add By Sindy 2012/1/13
   If ET03 <> "" Then
      bolEmail = PUB_GetEMailFlag(strTM01 & strTM02 & strTM03 & strTM04, strNP07 = "102", , bolPlusPaper)
      If bolEmail Then
         'Add By Sindy 2025/5/21 檢查是否為全E化,全E化不用列印出定稿
         If PUB_ChkECust(GetPrjPeopleNum1(strTM01 & "-" & strTM02 & "-" & strTM03 & "-" & strTM04), strTM01, , intR) = True Then
            '有指定信箱
            If intR = 1 Then
               stLP26 = "E"
            End If
         End If
         '2025/5/21 END
         '判斷是否EMail同時寄紙本
         If Not bolPlusPaper And stLP26 <> "E" Then 'Modify By Sindy 2025/5/21 + And stLP26 <> "E"
            iCopy = 1
         End If
         'Modify By Sindy 2019/10/18 + 信函收文號
         'NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True, , , , , IIf(strSrvDate(1) >= T商標電子化啟用日, strLD18, "")
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , IIf(strSrvDate(1) >= T商標電子化啟用日, strLD18, "")
'         boleFileSave = True
'         m_TM01 = strTM01
'         MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
      Else
         'Modify By Sindy 2019/10/18 + 信函收文號
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , IIf(strSrvDate(1) >= T商標電子化啟用日, strLD18, "")
      End If
   End If
   '2012/1/13 End
End Sub

'Modify By Sindy 2020/1/2
'Sub Process()
Function Process() As Boolean
'2020/1/2 END
   'Add By Sindy 2019/11/1
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   Else
      'ChDir App.path 'Add By Sindy 2020/2/25 釋放資料夾權限
      'Modified by Lydia 2020/03/16 遇到有PDF檔設唯讀,改用模組
      'If Dir(m_AttachPath & "\.") <> "" Then
      '   Kill m_AttachPath & "\*.*"
      '   DoEvents
      'End If
      Call PUB_KillTempFile(Mid(Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum, 2) & "\.") '去掉\
      'end 2020/03/16
   End If
   '2019/11/1 END
   
   pub_QL05 = pub_QL05 & ";" & Label1(6) & "管制表" 'Add By Sindy 2010/9/30
   Screen.MousePointer = vbHourglass
   '智權人員
   cnnConnection.Execute "DELETE FROM R020301_1 WHERE ID='" & strUserNum & "' "
   '延展, 刊登廣告, 第一期註冊費, 第二期註冊費, 繳年費
   cnnConnection.Execute "DELETE FROM R020301_2 WHERE ID='" & strUserNum & "1" & "' or id='" & strUserNum & "2" & "' "
   strSQL1 = ""
   strSQL2 = ""
   StrSQL6 = ""
   bolT102Again = False 'Add by Amy 2018/09/11
   'Modify by Amy 2018/08/31 +大陸撤三
   If Option1(1).Value = True Then
        '系統類別
        If Len(txt1(0)) <> 0 Then
            strSQL1 = strSQL1 + " and TM01 in (" & SQLGrpStr(txt1(0), 2) & ") "
            pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0)
        End If
        '大陸撤三期限
        If Len(txt1(21)) <> 0 Then
           'modify by sonia 2020/6/1 因延展改專用期起日會有錯,故改抓公告日+三個月+1天,無公告日則剔除
           'strSQL1 = strSQL1 & " And TM21>=" & Val(ChangeTStringToWString(txt1(21))) - 30000
           strSQL1 = strSQL1 & " And to_char(to_date(add_months(to_date(tm14,'YYYYMMDD'),3))+1,'YYYYMMDD')>=" & Val(ChangeTStringToWString(txt1(21))) - 30000
        End If
        If Len(txt1(22)) <> 0 Then
           'modify by sonia 2020/6/1 因延展改專用期起日會有錯,故改抓公告日+三個月+1天,無公告日則剔除
           'strSQL1 = strSQL1 & " And TM21<=" & Val(ChangeTStringToWString(txt1(22))) - 30000
           strSQL1 = strSQL1 & " And to_char(to_date(add_months(to_date(tm14,'YYYYMMDD'),3))+1,'YYYYMMDD')<=" & Val(ChangeTStringToWString(txt1(22))) - 30000
        End If
        If Len(txt1(21)) <> 0 Or Len(txt1(22)) <> 0 Then
           'modify by sonia 2020/6/1 因延展改專用期起日會有錯,故改抓公告日+三個月+1天,無公告日則剔除 And Nvl(tm14,0)<>0 And Nvl(tm21,0)<>0
           'Modify By Sindy 2020/8/3 管制表不鎖是否已跑過進度,均可以產出管制表或電子檔
'           strSQL1 = strSQL1 & " And TM10='020' And (TM29<>'Y' OR TM29 IS NULL) And Nvl(tm14,0)<>0 And Nvl(tm21,0)<>0  " & _
'                   " And Not Exists(Select * From CaseProgress b Where tm01=b.cp01 and tm02=b.cp02 and tm03=b.cp03 and tm04=b.cp04 and SubStr(b.CP09,1,1)='D' and b.CP10='1729')"
           strSQL1 = strSQL1 & " And TM10='020' And (TM29<>'Y' OR TM29 IS NULL) And Nvl(tm14,0)<>0 And Nvl(tm21,0)<>0 "
           pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & txt1(21) & "-" & txt1(22)
           pub_QL05 = pub_QL05 & ";" & Label1(9) & "020"
        End If
        
   '非大陸撤三
   Else
       '系統類別
       If Len(txt1(0)) <> 0 Then
          strSQL1 = strSQL1 + " and NP02 in (" & SQLGrpStr(txt1(0), 2) & ") "
          strSQL2 = strSQL2 + " and NP02 in (" & SQLGrpStr(txt1(0), 5) & ") "
          pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/9/30
       End If
       
       'add by nickc 2007/06/13
       If txt1(19).Text = "Y" Then
           If InStr(1, txt1(1), "716") = 0 Then
               txt1(1) = txt1(1) & ",716"
           End If
           pub_QL05 = pub_QL05 & ";" & Left(Label1(16), 7) & "含"  'Add By Sindy 2010/9/30
       End If
       
       StrSQL6 = ""
       Dim varTmp As Variant
       '案件性質
       If txt1(1) <> "" Then
          varTmp = Split(txt1(1), ",")
          For i = 0 To UBound(varTmp)
             If varTmp(i) <> "" Then
                StrSQL6 = StrSQL6 & Format(varTmp(i)) & ","
             End If
          Next
          pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1)  'Add By Sindy 2010/9/30
       End If
       If StrSQL6 <> "" Then StrSQL6 = " AND NP07 IN (" & Left(StrSQL6, Len(StrSQL6) - 1) & ") "
       '本所期限
       If Len(txt1(3)) <> 0 Then
          StrSQL6 = StrSQL6 + " AND NP08>=" & Val(ChangeTStringToWString(txt1(3))) & " "
       End If
       If Len(txt1(4)) <> 0 Then
          StrSQL6 = StrSQL6 + " AND NP08<=" & Val(ChangeTStringToWString(txt1(4))) & " "
       End If
       If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Then
          'Modify by Amy 2018/08/31 原:Label1(3)
          pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & txt1(3) & "-" & txt1(4)    'Add By Sindy 2010/9/30
       End If
       StrSQL6 = StrSQL6 & " AND (NP06 IS NULL OR NP06='') "
    '   '業務區
    '   If Len(TXT1(5)) <> 0 Then
    '       'Modify By Cheng 2003/03/11
    '   '    StrSQL6 = StrSQL6 + " AND s1.ST03>='" & TXT1(5) & "' "
    '       StrSQL6 = StrSQL6 + " AND s1.ST15>='" & TXT1(5) & "' "
    '   End If
    '   If Len(TXT1(6)) <> 0 Then
    '       'Modify By Cheng 2003/03/15
    '   '    StrSQL6 = StrSQL6 + " AND s1.ST03<='" & TXT1(6) & "' "
    '       StrSQL6 = StrSQL6 + " AND s1.ST15<='" & TXT1(6) & "' "
    '   End If
       If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
          pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(5) & "-" & txt1(6)    'Add By Sindy 2010/9/30
       End If
       '智權人員
       If Len(txt1(7)) <> 0 Then
    '       StrSQL6 = StrSQL6 + " AND NP10='" & TXT1(7) & "' "
           pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(7) & lbl1    'Add By Sindy 2010/9/30
       End If
       '申請人
       'If TXT1(9) = "97038" Then '臨時用
       'Else
          If Len(Trim(txt1(9))) <> 0 And Len(Trim(txt1(10))) <> 0 Then
              strSQL1 = strSQL1 & " AND (TM23>='" & GetNewFagent(txt1(9)) & "' AND TM23<='" & GetNewFagent(txt1(10)) & "') "
              strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(txt1(9)) & "' AND SP08<='" & GetNewFagent(txt1(10)) & "') OR (SP58<='" & GetNewFagent(txt1(9)) & "' AND SP58<='" & GetNewFagent(txt1(10)) & "') OR (SP59>='" & GetNewFagent(txt1(9)) & "' AND SP59<='" & GetNewFagent(txt1(10)) & "')) "
          Else
              If Len(Trim(txt1(9))) <> 0 And Len(Trim(txt1(10))) = 0 Then
                  strSQL1 = strSQL1 & " AND (TM23>='" & GetNewFagent(txt1(9)) & "' ) "
                  strSQL2 = strSQL2 + " AND (SP08>='" & GetNewFagent(txt1(9)) & "' OR SP58>='" & GetNewFagent(txt1(9)) & "' OR SP59>='" & GetNewFagent(txt1(9)) & "') "
              Else
                  If Len(Trim(txt1(9))) = 0 And Len(Trim(txt1(10))) <> 0 Then
                      strSQL1 = strSQL1 & " AND (TM23<='" & GetNewFagent(txt1(10)) & "') "
                      strSQL2 = strSQL2 + " AND (SP08<='" & GetNewFagent(txt1(10)) & "' OR SP58<='" & GetNewFagent(txt1(10)) & "' OR SP59<='" & GetNewFagent(txt1(10)) & "') "
                  End If
              End If
          End If
          If Len(Trim(txt1(9))) <> 0 Or Len(Trim(txt1(10))) <> 0 Then
             pub_QL05 = pub_QL05 & ";" & Label1(7) & txt1(9) & "-" & txt1(10)  'Add By Sindy 2010/9/30
          End If
       'End If
       '代理人
       If Len(Trim(txt1(11))) <> 0 And Len(Trim(txt1(12))) <> 0 Then
           strSQL1 = strSQL1 + " AND (TM44>='" & GetNewFagent(txt1(11)) & "' AND TM44<='" & GetNewFagent(txt1(12)) & "') "
           strSQL2 = strSQL2 + " AND (SP26>='" & GetNewFagent(txt1(11)) & "' AND SP26<='" & GetNewFagent(txt1(12)) & "') "
       Else
           If Len(Trim(txt1(11))) <> 0 And Len(Trim(txt1(12))) = 0 Then
               strSQL1 = strSQL1 + " AND (TM44>='" & GetNewFagent(txt1(11)) & "' ) "
               strSQL2 = strSQL2 + " AND (SP26>='" & GetNewFagent(txt1(11)) & "' ) "
           Else
               If Len(Trim(txt1(11))) = 0 And Len(Trim(txt1(12))) <> 0 Then
                   strSQL1 = strSQL1 + " AND (TM44<='" & GetNewFagent(txt1(12)) & "' ) "
                   strSQL2 = strSQL2 + " AND (SP26<='" & GetNewFagent(txt1(12)) & "' ) "
               End If
           End If
       End If
       If Len(Trim(txt1(11))) <> 0 Or Len(Trim(txt1(12))) <> 0 Then
          pub_QL05 = pub_QL05 & ";" & Label1(8) & txt1(11) & "-" & txt1(12)   'Add By Sindy 2010/9/30
       End If
       '申請國家
       '93.3.25 MODIFY BY SONIA 是否閉卷條件不管有沒有輸申請國家條件都要控管
       'If Len(txt1(13)) <> 0 Then
       '    strSQL1 = strSQL1 + " AND TM10>='" & txt1(13) & "' AND (TM29<>'Y' OR TM29 IS NULL) "
       '    strSQL2 = strSQL2 + " AND SP09>='" & txt1(13) & "' AND (SP15<>'Y' OR SP15 IS NULL) "
       'End If
       'If Len(txt1(14)) <> 0 Then
       '    strSQL1 = strSQL1 + " AND TM10<='" & txt1(14) & "' AND (TM29<>'Y' OR TM29 IS NULL) "
       '    strSQL2 = strSQL2 + " AND SP09<='" & txt1(14) & "' AND (SP15<>'Y' OR SP15 IS NULL) "
       'End If
       If Len(txt1(13)) <> 0 Then
           strSQL1 = strSQL1 + " AND TM10>='" & txt1(13) & "' "
           strSQL2 = strSQL2 + " AND SP09>='" & txt1(13) & "' "
       End If
       If Len(txt1(14)) <> 0 Then
           strSQL1 = strSQL1 + " AND TM10<='" & txt1(14) & "' "
           strSQL2 = strSQL2 + " AND SP09<='" & txt1(14) & "' "
       End If
       If Len(Trim(txt1(13))) <> 0 Or Len(Trim(txt1(14))) <> 0 Then
          pub_QL05 = pub_QL05 & ";" & Label1(9) & txt1(13) & "-" & txt1(14)   'Add By Sindy 2010/9/30
       End If
       strSQL1 = strSQL1 + " AND (TM29<>'Y' OR TM29 IS NULL) "
       strSQL2 = strSQL2 + " AND (SP15<>'Y' OR SP15 IS NULL) "
       
       'Add by Amy 2018/09/11 +延展再通知
       Call ShowT102Again
       If lblMsg.Caption <> MsgText(601) Then
            If txt1(13) = "000" And txt1(14) = "000" Then
                '通知北、中所
                bolT102Again = True
                strSQL1 = strSQL1 & " And s3.ST06 in('1','2') And CU13=s3.ST01(+)"
                strSQL2 = strSQL2 & " And s3.ST06 in('1','2') And CU13=s3.ST01(+)"
                pub_QL05 = pub_QL05 & ";通知北、中所"
            End If
            If txt1(13) = "020" And txt1(14) = "020" Then
                If Val(txt1(3)) + 19110000 < Val(GetPreMonLastDate(DBDATE(DateAdd("m", 6, Format(strSrvDate(1), "####/##/##"))), False)) Then
                    '到期前2個月通知北、中、南所
                    bolT102Again = True
                    strSQL1 = strSQL1 & " And s3.ST06 <>'4' And CU13=s3.ST01(+)"
                    strSQL2 = strSQL2 & " And s3.ST06 <>'4' And CU13=s3.ST01(+)"
                    pub_QL05 = pub_QL05 & ";通知北、中、南所"
                ElseIf Val(txt1(3)) + 19110000 < Val(GetPreMonLastDate(DBDATE(DateAdd("m", 18, Format(strSrvDate(1), "####/##/##"))), False)) Then
                    '通知北、中所
                    bolT102Again = True
                    strSQL1 = strSQL1 & " And s3.ST06 in('1','2') And CU13=s3.ST01(+)"
                    strSQL2 = strSQL2 & " And s3.ST06 in('1','2') And CU13=s3.ST01(+)"
                    pub_QL05 = pub_QL05 & ";通知北、中所"
                End If
            End If
       End If
       
       'Add By Sindy 2015/4/17 台灣案催延展對象
       '大->台係指葉經理及巨京收文且申請人國籍非台灣者,
       '其他案件皆屬台->台範圍
       If Trim(txt1(13)) = "000" And Trim(txt1(14)) = "000" Then
          '系統別有T者
          strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
          s = 0
          For i = 0 To UBound(strTemp2)
             If strTemp2(i) = "T" Then
                s = 1
                Exit For
             End If
          Next i
          If s = 1 And txt1(2) = "Y" Then 'T含延展
             If Trim(txt1(20)) = "1" Then '1.台->台
                'Modify By Sindy 2020/8/13 改判斷 TM44.FC代理人
                'StrSQL6 = StrSQL6 + " AND not(Substr(NP10,1,5) in('67002','96029','96030','MCTF0') AND CU10>'010') "
                strSQL1 = strSQL1 + " AND not(Substr(NP10,1,5) in('67002','96029','96030','MCTF0') AND TM44 is not null) "
                strSQL2 = strSQL2 + " AND not(Substr(NP10,1,5) in('67002','96029','96030','MCTF0') AND SP26 is not null) "
             ElseIf Trim(txt1(20)) = "2" Then '2.大->台
                'Modify by Amy 2017/01/10 +MCTF特殊人員
                'Modify By Sindy 2020/8/13 改判斷 TM44.FC代理人
                'StrSQL6 = StrSQL6 + " AND SubStr(NP10,1,5) in('67002','96029','96030','MCTF0') AND CU10>'010' "
                strSQL1 = strSQL1 + " AND SubStr(NP10,1,5) in('67002','96029','96030','MCTF0') AND TM44 is not null "
                strSQL2 = strSQL2 + " AND SubStr(NP10,1,5) in('67002','96029','96030','MCTF0') AND SP26 is not null "
             End If
             pub_QL05 = pub_QL05 & ";" & Label1(19) & txt1(20) & Label1(18)
          End If
       End If
       '2015/4/17 END
       
       'add by nickc 2006/05/30
       If txt1(16) = "1" Then
    '       StrSQL6 = StrSQL6 + " AND substr(s1.ST15,1,1)<>'S' "
           pub_QL05 = pub_QL05 & ";" & Label1(13) & "非智權部同仁"   'Add By Sindy 2010/9/30
       End If
       'add by nickc 2006/05/30 延展和第二期專用權須存在
       If InStr(1, txt1(1), "716") <> 0 Or InStr(1, txt1(1), "102") <> 0 Then
           strSQL1 = strSQL1 & " and decode(np07,716,tm17,102,tm17,'Y')='Y' "
       End If
   End If
   'end 2018/08/31
   'Add By Sindy 2020/8/3
   If ChkMail.Value = 0 Then
      pub_QL05 = pub_QL05 & ";「不」" & ChkMail.Caption
   ElseIf ChkMail.Value = 1 Then
      pub_QL05 = pub_QL05 & ";「要」" & ChkMail.Caption
   End If
   '2020/8/3 END
   
   '93.3.25 END
   '智權人員
   If Process1 = True Then
      Process = True 'Add By Sindy 2020/1/2
      PrintData1
   End If
   '延展, 刊登廣告
   'Add By Cheng 2003/02/04
   '若有選擇含延展, 刊登廣告, 第一期註冊費, 第二期註冊費, 繳年費時
   'edit by nick 2004/07/20
   'If Me.txt1(2).Text = "Y" Or Me.Text1.Text = "Y" Or Me.txt1(16).Text = "Y" Or Me.txt1(17).Text = "Y" Or Me.txt1(18).Text = "Y" Then
   'If Me.txt1(2).Text = "Y" Or Me.Text1.Text = "Y" Or Me.txt1(18).Text = "Y" Then
   'edit by nickc 2007/03/29
   'If Me.txt1(2).Text = "Y" Or Me.Text1.Text = "Y" Or Me.txt1(18).Text = "Y" Or InStr(1, txt1(1), "716") <> 0 Then
   If Me.txt1(2).Text = "Y" Or Me.Text1.Text = "Y" Or Me.txt1(18).Text = "Y" Or InStr(1, txt1(1), "716") <> 0 Or txt1(19).Text = "Y" Then
      If Process2 = True Then
         Process = True 'Add By Sindy 2020/1/2
         PrintData2
      End If
   End If
   Screen.MousePointer = vbDefault
End Function

Sub PrintData1()
Dim stFileName As String
Dim dblFCnt As Double
Dim strTo As String
   
   '91/03/12 日期排序不能用符號
   'nick
   'Modify By Cheng 2003/03/11
   'strSQL = "select nvl(a0902,a0903),nvl(st02,r053002),r053003,r053004,r053005,r053006,r053007,r053008,r053009,r053010,r053011,r053012,r053013,r053014,r053001,r053002 from r020301_1,staff,acc090  where r053002=st01(+) and st03=a0901(+) and id='" & strUserNum & "' order by r053001,r053002,decode(r053003,substr(r053003,1,1),'#',substr(r053003,2,10),'V',substr(r053003,2,10),'*',substr(r053003,2,10),r053003),r053005 "
   strSql = "select nvl(a0902,a0903),nvl(st02,r053002),r053003,r053004,r053005,r053006,r053007,r053008,r053009,r053010,r053011,r053012,r053013,r053014,r053001,r053002 from r020301_1,staff,acc090  where r053002=st01(+) and ST15=a0901(+) and id='" & strUserNum & "' order by r053001,r053002,decode(r053003,substr(r053003,1,1),'#',substr(r053003,2,10),'V',substr(r053003,2,10),'*',substr(r053003,2,10),r053003),r053005 "
   CheckOC
   'Modify By Cheng 2002/03/01
   'Page = 1
   Page = 0
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           'Add By Cheng 2002/03/01
           m_strSales = ""
           strTemp3 = CheckStr(.Fields(0))
            'Modify By Cheng 2002/03/01
   '        PrintTitle1
         'Add By Sindy 2019/11/1
         If ChkMail.Value = 1 Then
            '產生PDF
            'Load frmPDF
            frmPDF.Show
            If Option1(1).Value = True Then
               stFileName = .Fields(15) & "-大陸撤三開拓函寄發清單" & txt1(21) & "~" & txt1(22) 'strSrvDate(2) '& "-" & CheckStr(.Fields(1))
            Else
               stFileName = .Fields(15) & "-智權人員期限管制表" & txt1(3) & "~" & txt1(4) 'strSrvDate(2) '& "-" & CheckStr(.Fields(1))
            End If
            frmPDF.StartProcess m_AttachPath, stFileName
         End If
         '2019/11/1 END
         
           Do While .EOF = False
               For i = 0 To 15 '13
                   strTemp(i) = CheckStr(.Fields(i))
               Next i
               'Modify By Cheng 2002/03/01
               '智權人員不同則跳頁, 業務區不同則跳頁
               If m_strSales <> strTemp(1) Then
                  If Option1(1).Value = True And m_strSales <> "" Then PrintMemo 'Add by Amy 2018/08/31
                  
                  'Add By Sindy 2019/11/1
                  If ChkMail.Value = 1 And m_strSales <> "" Then
                     Printer.EndDoc
                     frmPDF.EndtProcess
                     Unload frmPDF
                     '產生PDF
                     'Load frmPDF
                     frmPDF.Show
                     If Option1(1).Value = True Then
                        stFileName = CheckStr(.Fields(15)) & "-大陸撤三開拓函寄發清單" & txt1(21) & "~" & txt1(22) 'strSrvDate(2) '& "-" & CheckStr(.Fields(1))
                     Else
                        stFileName = CheckStr(.Fields(15)) & "-智權人員期限管制表" & txt1(3) & "~" & txt1(4) 'strSrvDate(2) '& "-" & CheckStr(.Fields(1))
                     End If
                     frmPDF.StartProcess m_AttachPath, stFileName
                     Page = 0
                  End If
                  '2019/11/1 END
                  
                  strTemp3 = strTemp(0)
                  Page = Page + 1
                  If Page <> 1 Then
                    Printer.NewPage
                  End If
                  PrintTitle1
'               ElseIf strTemp3 <> strTemp(0) Then
'                   strTemp3 = strTemp(0)
'                   If Option1(1).Value = True Then PrintMemo  'Add by Amy 2018/08/31
'                   Page = Page + 1
'                   If Page <> 1 Then
'                     Printer.NewPage
'                   End If
'                   PrintTitle1
'                   m_strSales = ""
               End If
               strTemp(1) = StrToStr(strTemp(1), 4)
               strTemp(5) = StrToStr(strTemp(5), 9)
               If Option1(1).Value = False Then strTemp(6) = StrToStr(strTemp(6), 4)  'Modify by Amy 2018/08/31 大陸撤三不印
               'Modify by Sindy 2012/10/18
               'strTemp(7) = StrToStr(strTemp(7), 4)
               strTemp(7) = strTemp(7)
               '2012/10/18 End
               'Modify By Cheng 2003/06/26
               '案件性質
   '            strTemp(8) = StrToStr(strTemp(8), 4)
               If Option1(1).Value = False Then strTemp(8) = StrToStr(strTemp(8), 7) 'Modify by Amy 2018/08/31 大陸撤三不印
               strTemp(9) = StrToStr(strTemp(9), 8)
               strTemp(10) = ""
               strTemp(11) = StrToStr(strTemp(11), 4)
               strTemp(12) = StrToStr(strTemp(12), 5)
               strTemp(13) = StrToStr(strTemp(13), 4)
               PrintDatil1
               If iPrint >= 10000 Then
                   'If Option1(1).Value = True Then PrintMemo 'Add by Amy 2018/08/31
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle1
               End If
               .MoveNext
           Loop
           If Option1(1).Value = True Then PrintMemo  'Add by Amy 2018/08/31
           CheckOC
           Printer.EndDoc
           
           'Add By Sindy 2019/11/1
           If ChkMail.Value = 1 Then
               frmPDF.EndtProcess
               Unload frmPDF
           End If
           '2019/11/1 END
       End With
   Else
       CheckOC
       Exit Sub
   End If
'   If Option1(1).Value = True Then PrintMemo  'Add by Amy 2018/08/31
'   CheckOC
'   Printer.EndDoc
   
   '若不列印延展, 刊登廣告, 第一期註冊費, 第二期註冊費
   'edit by nick 2004/07/20
   'If txt1(2) = "" And Text1 = "" And Me.txt1(16).Text = "" And Me.txt1(17).Text = "" Then
   'edit by nickc 2007/03/29
   'If txt1(2) = "" And Text1 = "" Then
   If txt1(2) = "" And Text1 = "" And txt1(19) = "" Then
       If ChkMail.Value = 0 Then ShowPrintOk
   End If
   
   'Add By Sindy 2019/11/1
   If ChkMail.Value = 1 Then
      '寄Mail
      File1.path = m_AttachPath & "\"
      File1.Refresh
      If File1.ListCount > 0 Then
         If Option1(1).Value = True Then
            stFileName = "大陸撤三開拓函寄發清單"
         Else
            stFileName = "智權人員期限管制表"
         End If
         For dblFCnt = 0 To File1.ListCount - 1
            If InStr(UCase(Trim(File1.List(dblFCnt))), stFileName) > 0 And _
               UCase(Right(Trim(File1.List(dblFCnt)), 4)) = ".PDF" Then
               strTo = Left(Trim(File1.List(dblFCnt)), InStr(Trim(File1.List(dblFCnt)), "-") - 1)
               PUB_SendMail strUserNum, strTo, "", Left(File1.List(dblFCnt), Len(File1.List(dblFCnt)) - 4), "請參附件！", , m_AttachPath & "\" & File1.List(dblFCnt), , , , , , , , True
               Kill m_AttachPath & "\" & File1.List(dblFCnt)
            End If
         Next dblFCnt
      End If
   End If
   '2019/11/1 END
End Sub

'列印延展, 刊登廣告, 註冊費, 繳年費管制表
Sub PrintData2()
Dim strSalesNo As String '智權人員 Add By Cheng 2003/02/27
'Add By Sindy 2019/11/4
Dim stFileName As String
Dim dblFCnt As Double
Dim strTo As String
'2019/11/4 END
   
   'Modify By Cheng 2002/02/18
   'strSQL = "select nvl(a0902,a0903),nvl(st02,r054002),r054003,r054004,r054005,r054006,r054007,r054008,r054009,r054010,r054011,r054012,r054001,r054002 from r020301_2,staff ,acc090 where r054002=st01(+) and st03=a0901(+) and  id='" & strUserNum & "1" & "' order by r054001,r054002,r054005 "
   'Modify By Cheng 2003/03/11
   'strSQL = "select nvl(a0902,a0903),nvl(st02,r054002),r054003,r054004,r054005,r054006,r054007,r054008,r054009,r054010,r054011,r054012,r054001,r054002,r054013,r054014,r054015,r054016,r054017 " & _
   '         " from r020301_2,staff ,acc090 " & _
   '         " where r054002=st01(+) and st03=a0901(+) and id='" & strUserNum & "1" & "' " & _
   '         " order by r054001,r054002,r054005 "
   'Modify By Cheng 2003/05/22
   'strSQL = "select nvl(a0902,a0903),nvl(st02,r054002),r054003,r054004,r054005,r054006,r054007,r054008,r054009,r054010,r054011,r054012,r054001,r054002,r054013,r054014,r054015,r054016,r054017 " & _
   '         " from r020301_2,staff ,acc090 " & _
   '         " where r054002=st01(+) and ST15=a0901(+) and id='" & strUserNum & "1" & "' " & _
   '         " order by r054001,r054002,r054005 "
   'edit by nickc 2007/06/13 加入申請案號
   'strSQL = "select nvl(a0902,a0903),nvl(st02,r054002),r054003,r054004,r054005,r054006,r054007,r054008,r054009,r054010,r054011,r054012,r054001,r054002,r054013,r054014,r054015,r054016,r054017 " & _
            " from r020301_2,staff ,acc090 " & _
            " where r054002=st01(+) and ST15=a0901(+) and id='" & strUserNum & "1" & "' " & _
            " order by r054001, r054002, r054013, r054005 "
   'Modify By Sindy 2013/9/17 +tm129
   strSql = "select nvl(a0902,a0903),nvl(st02,r054002),r054003,r054004,r054005,r054006,r054007,r054008,r054009,r054010,r054011,r054012,r054001,r054002,r054013,r054014,r054015,r054016,r054017,tm12,tm129 " & _
            " from r020301_2,staff,acc090,trademark " & _
            " where r054002=st01(+) and ST15=a0901(+) and id='" & strUserNum & "1" & "' and r054005=tm01||'-'||tm02||'-'||TM03||'-'||tm04  "
   strSql = strSql & " union all select nvl(a0902,a0903),nvl(st02,r054002),r054003,r054004,r054005,r054006,r054007,r054008,r054009,r054010,r054011,r054012,r054001,r054002,r054013,r054014,r054015,r054016,r054017,sp11 as tm12,'' as tm129 " & _
            " from r020301_2,staff,acc090,servicepractice " & _
            " where r054002=st01(+) and ST15=a0901(+) and id='" & strUserNum & "1" & "' and r054005=sp01||'-'||sp02||'-'||sp03||'-'||sp04  " & _
            " order by r054001, r054002, r054013, r054005 "
   
   CheckOC
   Page = 1
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           'Add By Cheng 2003/02/27
           '記錄智權人員
           strSalesNo = "" & .Fields(13).Value
           strTemp3 = CheckStr(.Fields(0))
           
           'Add By Sindy 2019/11/1
           If ChkMail.Value = 1 Then
               '產生PDF
               'Load frmPDF
               frmPDF.Show
               If InStr(1, txt1(1), "716") <> 0 Then
                   stFileName = strSalesNo & "-內商註冊費管制表" & txt1(3) & "~" & txt1(4) 'strSrvDate(2) '& "-" & CheckStr(.Fields(1))
               Else
                   stFileName = strSalesNo & "-內商延展及繳年費管制表" & txt1(3) & "~" & txt1(4) 'strSrvDate(2) '& "-" & CheckStr(.Fields(1))
               End If
               frmPDF.StartProcess m_AttachPath, stFileName
           End If
           '2019/11/1 END
           
           PrintTitle2
           Do While .EOF = False
               For i = 0 To 11
                   strTemp(i) = CheckStr(.Fields(i))
               Next i
               'Add By Cheng 2003/02/27
               '若智權人員不同
               If strSalesNo <> "" & .Fields(13).Value Then
                   'Add By Sindy 2012/2/20 加註解
                   iPrint = iPrint + 300
                   Printer.CurrentX = PLeft(1)
                   Printer.CurrentY = iPrint
                   'Modify By Sindy 2013/9/16 加不催延展註解
                   Printer.Print "PS：本所案號前有※符號者表示案件已轉他所 , x 表示不催延展"
                   '2012/2/20 End
                   
                   'Add By Sindy 2019/11/1
                   If ChkMail.Value = 1 Then
                     Printer.EndDoc
                     frmPDF.EndtProcess
                     Unload frmPDF
                     '產生PDF
                     'Load frmPDF
                     frmPDF.Show
                     If InStr(1, txt1(1), "716") <> 0 Then
                         stFileName = "" & .Fields(13).Value & "-內商註冊費管制表" & txt1(3) & "~" & txt1(4) 'strSrvDate(2) '& "-" & CheckStr(.Fields(1))
                     Else
                         stFileName = "" & .Fields(13).Value & "-內商延展及繳年費管制表" & txt1(3) & "~" & txt1(4) 'strSrvDate(2) '& "-" & CheckStr(.Fields(1))
                     End If
                     frmPDF.StartProcess m_AttachPath, stFileName
                     Page = 0
                   End If
                   '2019/11/1 END
                   
                   strSalesNo = "" & .Fields(13).Value
                   strTemp3 = strTemp(0)
                   Page = Page + 1
                   If Page <> 1 Then
                     Printer.NewPage
                   End If
                   PrintTitle2
               End If
'               '若業務區不同
'               If strTemp3 <> strTemp(0) Then
'                   strTemp3 = strTemp(0)
'                   'Add By Sindy 2012/2/20 加註解
'                   iPrint = iPrint + 300
'                   Printer.CurrentX = PLeft(1)
'                   Printer.CurrentY = iPrint
'                   'Modify By Sindy 2013/9/16 加不催延展註解
'                   Printer.Print "PS：本所案號前有※符號者表示案件已轉他所 , x 表示不催延展"
'                   '2012/2/20 End
'                   Page = Page + 1
'                   Printer.NewPage
'                   PrintTitle2
'               End If
               'Move By Cheng 2003/03/07
               '避免發生只列印表頭而無內容的情況
               If iPrint >= 10000 Then
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle2
               End If
               strTemp(1) = StrToStr(strTemp(1), 4)
               strTemp(2) = StrToStr(strTemp(2), 5)
               strTemp(3) = StrToStr(strTemp(3), 12)
               'Add By Sindy 2012/2/20 T*案件若該案號進度檔有728的進度時,在本所期限前加※符號
               If Left(Trim(strTemp(4)), 1) = "T" Then
                  strTemp(4) = PUB_GetCP10ValueAttachText(strTemp(4), "728", "※", strTemp(4))
               End If
               '2012/2/20 End
               'Add By Sindy 2013/9/17
               If "" & .Fields("tm129").Value = "Y" Then
                  strTemp(4) = "x" & strTemp(4)
               End If
               '2013/9/17 END
               strTemp(5) = StrToStr(strTemp(5), 7)
               'Modify By Sindy 2012/10/18
               'strTemp(6) = StrToStr(strTemp(6), 4)
               strTemp(6) = strTemp(6)
               '2012/10/18 End
               strTemp(7) = StrToStr(strTemp(7), 12)
               strTemp(9) = StrToStr(strTemp(9), 4)
               strTemp(10) = StrToStr(strTemp(10), 6)
               strTemp(11) = StrToStr(strTemp(11), 27)
               'Add By Cheng 2002/02/18
               strTemp(12) = CheckStr(.Fields("R054013"))
               'Modify By Cheng 2003/05/29
   '            strTemp(13) = CheckStr(.Fields("R054014"))
   '            strTemp(14) = CheckStr(.Fields("R054015"))
               strTemp(13) = CheckStr(.Fields("R054014")) & IIf("" & .Fields("R054015").Value <> "", ",", "") & "" & .Fields("R054015").Value
               strTemp(14) = ""
               'Modify By Cheng 2003/05/29
   '            strTemp(15) = CheckStr(.Fields("R054016"))
   '            strTemp(16) = CheckStr(.Fields("R054017"))
               strTemp(15) = CheckStr(.Fields("R054016")) & IIf("" & .Fields("R054017").Value <> "", ",", "") & "" & .Fields("R054017").Value
               strTemp(16) = ""
               'add by nickc 2007/06/13
               strTemp(17) = CheckStr(.Fields("tm12"))
               
               PrintDatil2
               'Modify By Cheng 2003/03/07
   '            If iPrint >= 10000 Then
   '                Page = Page + 1
   '                Printer.NewPage
   '                PrintTitle2
   '            End If
               .MoveNext
           Loop
           'Add By Sindy 2012/2/20 加註解
           iPrint = iPrint + 300
           Printer.CurrentX = PLeft(1)
           Printer.CurrentY = iPrint
           'Modify By Sindy 2013/9/16 加不催延展註解
           Printer.Print "PS：本所案號前有※符號者表示案件已轉他所 , x 表示不催延展"
           '2012/2/20 End
           Printer.EndDoc
           
            'Add By Sindy 2019/11/1
            If ChkMail.Value = 1 Then
               frmPDF.EndtProcess
               Unload frmPDF
            End If
            '2019/11/1 END
       End With
   Else
       CheckOC
       Exit Sub
   End If
   CheckOC
   'strSQL = "select * from r020301_2 where id='" & strUserNum & "2" & "' order by r054001,r054002,r054005 "
   'Modify By Cheng 2002/02/18
   'strSQL = "select nvl(a0902,a0903),nvl(st02,r054002),r054003,r054004,r054005,r054006,r054007,r054008,r054009,r054010,r054011,r054012,r054001,r054002 from r020301_2,staff ,acc090 where r054002=st01(+) and st03=a0901(+) and  id='" & strUserNum & "2" & "' order by r054001,r054002,r054005 "
   'Modify By Sindy 2013/9/17 +tm129
   strSql = "select nvl(a0902,a0903),nvl(st02,r054002),r054003,r054004,r054005,r054006,r054007,r054008,r054009,r054010,r054011,r054012,r054001,r054002,r054013,r054014,r054015,r054016,r054017,tm129 " & _
            " from r020301_2,staff,acc090,trademark " & _
            " where r054002=st01(+) and ST15=a0901(+) and id='" & strUserNum & "2" & "' and r054005=tm01||'-'||tm02||'-'||TM03||'-'||tm04 " & _
            " order by r054001,r054002,r054005 "
   CheckOC
   Page = 1
   'Printer.NewPage
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           'Add By Cheng 2003/02/27
           '記錄智權人員
           strSalesNo = "" & .Fields(13).Value
           strTemp3 = CheckStr(.Fields(0))
           
           'Add By Sindy 2019/11/1
           If ChkMail.Value = 1 Then
               '產生PDF
               'Load frmPDF
               frmPDF.Show
               If InStr(1, txt1(1), "716") <> 0 Then
                   stFileName = strSalesNo & "-內商註冊費管制表" & txt1(3) & "~" & txt1(4) 'strSrvDate(2) '& "-" & CheckStr(.Fields(1))
               Else
                   stFileName = strSalesNo & "-內商延展及繳年費管制表" & txt1(3) & "~" & txt1(4) 'strSrvDate(2) '& "-" & CheckStr(.Fields(1))
               End If
               frmPDF.StartProcess m_AttachPath, stFileName
           End If
           '2019/11/1 END
           
           PrintTitle2
           Do While .EOF = False
               For i = 0 To 11
                   strTemp(i) = CheckStr(.Fields(i))
               Next i
               'Add By Cheng 2003/02/27
               '若智權人員不同
               If strSalesNo <> "" & .Fields(13).Value Then
                   'Add By Sindy 2012/2/20 加註解
                   iPrint = iPrint + 300
                   Printer.CurrentX = PLeft(1)
                   Printer.CurrentY = iPrint
                   'Modify By Sindy 2013/9/16 加不催延展註解
                   Printer.Print "PS：本所案號前有※符號者表示案件已轉他所 , x 表示不催延展"
                   '2012/2/20 End
                   
                   'Add By Sindy 2019/11/1
                   If ChkMail.Value = 1 Then
                     Printer.EndDoc
                     frmPDF.EndtProcess
                     Unload frmPDF
                     '產生PDF
                     'Load frmPDF
                     frmPDF.Show
                     If InStr(1, txt1(1), "716") <> 0 Then
                         stFileName = "" & .Fields(13).Value & "-內商註冊費管制表" & txt1(3) & "~" & txt1(4) 'strSrvDate(2) '& "-" & CheckStr(.Fields(1))
                     Else
                         stFileName = "" & .Fields(13).Value & "-內商延展及繳年費管制表" & txt1(3) & "~" & txt1(4) 'strSrvDate(2) '& "-" & CheckStr(.Fields(1))
                     End If
                     frmPDF.StartProcess m_AttachPath, stFileName
                     Page = 0
                   End If
                   '2019/11/1 END
                   
                   strSalesNo = "" & .Fields(13).Value
                   strTemp3 = strTemp(0)
                   Page = Page + 1
                   If Page <> 1 Then
                     Printer.NewPage
                   End If
                   PrintTitle2
               End If
'               '若業務區不同
'               If strTemp3 <> strTemp(0) Then
'                   strTemp3 = strTemp(0)
'                   'Add By Sindy 2012/2/20 加註解
'                   iPrint = iPrint + 300
'                   Printer.CurrentX = PLeft(1)
'                   Printer.CurrentY = iPrint
'                   'Modify By Sindy 2013/9/16 加不催延展註解
'                   Printer.Print "PS：本所案號前有※符號者表示案件已轉他所 , x 表示不催延展"
'                   '2012/2/20 End
'                   Page = Page + 1
'                   Printer.NewPage
'                   PrintTitle2
'               End If
               'Move By Cheng 2003/03/07
               '避免發生只列印表頭而無內容的情況
               If iPrint >= 10000 Then
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle2
               End If
               strTemp(1) = StrToStr(strTemp(1), 4)
               strTemp(2) = StrToStr(strTemp(2), 4)
               strTemp(3) = StrToStr(strTemp(3), 12)
               'Add By Sindy 2012/2/20 T*案件若該案號進度檔有728的進度時,在本所期限前加※符號
               If Left(Trim(strTemp(4)), 1) = "T" Then
                  strTemp(4) = PUB_GetCP10ValueAttachText(strTemp(4), "728", "※", strTemp(4))
               End If
               '2012/2/20 End
               'Add By Sindy 2013/9/17
               If "" & .Fields("tm129").Value = "Y" Then
                  strTemp(4) = "x" & strTemp(4)
               End If
               '2013/9/17 END
               strTemp(5) = StrToStr(strTemp(5), 7)
               strTemp(6) = StrToStr(strTemp(6), 4)
               strTemp(7) = StrToStr(strTemp(7), 12)
               strTemp(9) = StrToStr(strTemp(9), 4)
               strTemp(10) = StrToStr(strTemp(10), 6)
               strTemp(11) = StrToStr(strTemp(11), 27)
               'Add By Cheng 2002/02/18
               strTemp(12) = CheckStr(.Fields("R054013"))
               'Modify By Cheng 2003/05/29
   '            strTemp(13) = CheckStr(.Fields("R054014"))
   '            strTemp(14) = CheckStr(.Fields("R054015"))
               strTemp(13) = CheckStr(.Fields("R054014")) & IIf("" & .Fields("R054015").Value <> "", ",", "") & "" & .Fields("R054015").Value
               strTemp(14) = ""
               'Modify By Cheng 2003/05/29
   '            strTemp(15) = CheckStr(.Fields("R054016"))
   '            strTemp(16) = CheckStr(.Fields("R054017"))
               strTemp(15) = CheckStr(.Fields("R054016")) & IIf("" & .Fields("R054017").Value <> "", ",", "") & "" & .Fields("R054017").Value
               strTemp(16) = ""
               
               PrintDatil2
               'Modify By Cheng 2003/03/07
   '            If iPrint >= 10000 Then
   '                Page = Page + 1
   '                Printer.NewPage
   '                PrintTitle2
   '            End If
               .MoveNext
           Loop
           'Add By Sindy 2012/2/20 加註解
           iPrint = iPrint + 300
           Printer.CurrentX = PLeft(1)
           Printer.CurrentY = iPrint
           'Modify By Sindy 2013/9/16 加不催延展註解
           Printer.Print "PS：本所案號前有※符號者表示案件已轉他所 , x 表示不催延展"
           '2012/2/20 End
           Printer.EndDoc
            
            'Add By Sindy 2019/11/1
            If ChkMail.Value = 1 Then
               frmPDF.EndtProcess
               Unload frmPDF
            End If
            '2019/11/1 END
       End With
   Else
       CheckOC
       'Printer.EndDoc
       '若有列印延展, 刊登廣告, 第一期註冊費, 第二期註冊費
       'edit by nick 2004/07/20
       'If txt1(2) = "Y" Or Text1 = "Y" Or txt1(16) = "Y" Or txt1(17) = "Y" Or txt1(18) = "Y" Then
       'edit by nickc 2007/03/29
       'If txt1(2) = "Y" Or Text1 = "Y" Or txt1(18) = "Y" Then
       If txt1(2) = "Y" Or Text1 = "Y" Or txt1(18) = "Y" Or txt1(19) = "Y" Then
           If ChkMail.Value = 0 Then ShowPrintOk
       End If
       'Exit Sub
       GoTo RunNext
   End If
   CheckOC
   'Printer.EndDoc
   '若有列印延展, 刊登廣告, 第一期註冊費, 第二期註冊費
   'edit by nick 2004/07/20
   'If txt1(2) = "Y" Or Text1 = "Y" Or Me.txt1(16).Text = "Y" Or Me.txt1(17).Text = "Y" Or txt1(18) = "Y" Then
   'edit by nickc 2007/03/29
   'If txt1(2) = "Y" Or Text1 = "Y" Or txt1(18) = "Y" Then
   If txt1(2) = "Y" Or Text1 = "Y" Or txt1(18) = "Y" Or txt1(19) = "Y" Then
       If ChkMail.Value = 0 Then ShowPrintOk
   End If
      
   'Add By Sindy 2019/11/1
RunNext:
   If ChkMail.Value = 1 Then
      '寄Mail
      File1.path = m_AttachPath & "\"
      File1.Refresh
      If File1.ListCount > 0 Then
         If InStr(1, txt1(1), "716") <> 0 Then
             stFileName = "內商註冊費管制表"
         Else
             stFileName = "內商延展及繳年費管制表"
         End If
         For dblFCnt = 0 To File1.ListCount - 1
            If InStr(UCase(Trim(File1.List(dblFCnt))), stFileName) > 0 And _
               UCase(Right(Trim(File1.List(dblFCnt)), 4)) = ".PDF" Then
               strTo = Left(Trim(File1.List(dblFCnt)), InStr(Trim(File1.List(dblFCnt)), "-") - 1)
               PUB_SendMail strUserNum, strTo, "", Left(File1.List(dblFCnt), Len(File1.List(dblFCnt)) - 4), "請參附件！", , m_AttachPath & "\" & File1.List(dblFCnt), , , , , , , , True
               Kill m_AttachPath & "\" & File1.List(dblFCnt)
            End If
         Next dblFCnt
      End If
   End If
   '2019/11/1 END
End Sub

Sub PrintTitle1()
   GetPleft1
   iPrint = 500
   
   Printer.Orientation = 2
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 5700
   Printer.CurrentY = iPrint
   'Modify by Amy 2018/08/31 +大陸撤三
   If Option1(1).Value = True Then
      Printer.Print "大陸撤三開拓函寄發清單" 'Modify by Amy 2018/11/06 原:大陸預告撤三管制表
   Else
      Printer.Print GetTitleNick & "智權人員期限管制表"
   End If
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   iPrint = iPrint + 500
   Printer.CurrentY = iPrint
   'Modify by Amy 2018/08/31 +if 大陸撤三
   If Option1(1).Value = True Then
      Printer.CurrentX = 6000
      Printer.Print "大陸撤三期限：" & Format(ChangeTStringToTDateString(txt1(21)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(22))
   Else
      Printer.CurrentX = 6700
      Printer.Print "本所期限：" & Format(ChangeTStringToTDateString(txt1(3)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(4))
   End If
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "業務區：" & strTemp3
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(Page)
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.Font.Size = 10
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "智權人員"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "本所期限"
   'Modify by Amy 2018/08/31 +if 大陸撤三不印
   If Option1(1).Value = False Then
      Printer.CurrentX = PLeft(3)
      Printer.CurrentY = iPrint
      Printer.Print "法定期限"
   End If
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "商品類別"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "申請案號/審定號"
   'Modify by Amy 2018/08/31 +if 大陸撤三不印
   If Option1(1).Value = False Then
      Printer.CurrentX = PLeft(8)
      Printer.CurrentY = iPrint
      Printer.Print "案件性質"
   End If
   'end 2018/08/31
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = iPrint
   'edit by nickc 2006/05/30
   'Printer.Print "備  註"
   Printer.Print "分所案號"
   'Printer.CurrentX = PLeft(10)
   'Printer.CurrentY = iPrint
   'Printer.Print "是否出名"
   Printer.CurrentX = PLeft(11)
   Printer.CurrentY = iPrint
   Printer.Print "承辦人"
   Printer.CurrentX = PLeft(12)
   Printer.CurrentY = iPrint
   Printer.Print "申請人"
   Printer.CurrentX = PLeft(13)
   Printer.CurrentY = iPrint
   Printer.Print "申請國家"
   iPrint = iPrint + 300
   Printer.Font.Size = 12
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.Font.Size = 10
End Sub

Sub PrintDatil1()
   'Modify By Cheng 2002/03/01
   '若智權人員與上筆相同時不印
   If m_strSales <> strTemp(1) Then
      m_strSales = strTemp(1)
      For i = 1 To 13
         Printer.CurrentX = PLeft(i)
         Printer.CurrentY = iPrint
         Printer.Print strTemp(i)
      Next i
   Else
      For i = 2 To 13
         Printer.CurrentX = PLeft(i)
         Printer.CurrentY = iPrint
         Printer.Print strTemp(i)
      Next i
   End If
   iPrint = iPrint + 300
End Sub

Sub GetPleft1()
   Erase PLeft
   PLeft(0) = 500
   PLeft(1) = 500
   PLeft(2) = 1500
   PLeft(3) = 2500
   PLeft(4) = 3500
   PLeft(5) = 5000
   PLeft(6) = 6000 + 1100
   PLeft(7) = 7000 + 1100
   PLeft(8) = 8800 + 1100 - 200
   PLeft(9) = 9800 + 1100 + 200 + 200
   PLeft(10) = 10800 + 1100 + 200
   PLeft(11) = 11800 + 1100 + 200
   PLeft(12) = 12800 + 1100 + 200
   PLeft(13) = 15000 + 200
End Sub

Sub PrintTitle2()
   GetPleft2
   iPrint = 500
   Printer.Orientation = 2
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   'edit by nick 2004/07/06 刪除註冊費
   'Printer.Print "內商延展、刊登廣告、註冊費及繳年費管制表"
   'edit by nickc 2006/07/07
   If InStr(1, txt1(1), "716") <> 0 Then
       Printer.CurrentX = 6700
       Printer.CurrentY = iPrint
       Printer.Print "內商註冊費管制表"
   Else
       Printer.CurrentX = 4000
       Printer.CurrentY = iPrint
       Printer.Print "內商延展及繳年費管制表"
   End If
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   iPrint = iPrint + 500
   Printer.CurrentX = 6700
   Printer.CurrentY = iPrint
   Printer.Print "本所期限：" & Format(ChangeTStringToTDateString(txt1(3)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(4))
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   'Modify By Cheng 2002/11/29
   'Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   Printer.Print "列印日期：" & Format(Me.txt1(15).Text, "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "業務區：" & strTemp3
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(Page)
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.Font.Size = 10
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "智權人員"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "申請人"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "審定號數"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "商品類別"
   
   'edit by nickc 2006/08/25
   If InStr(1, txt1(1), "716") <> 0 Then
       'add by nickc 2007/06/13  加入申請案號
       Printer.CurrentX = PLeft(7) + 1500
       Printer.CurrentY = iPrint
       Printer.Print "申請案號"
       
       Printer.CurrentX = PLeft(8)
       Printer.CurrentY = iPrint
       Printer.Print "法定期限"
   Else
       Printer.CurrentX = PLeft(8) + 500
       Printer.CurrentY = iPrint
       Printer.Print "專用期間"
   End If
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = iPrint
   Printer.Print "申請國家"
   Printer.CurrentX = PLeft(10)
   Printer.CurrentY = iPrint
   'edit by nickc 2006/05/30
   'Printer.Print "正商標號數"
   Printer.Print "分所案號"
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(11)
   Printer.CurrentY = iPrint
   Printer.Print "申請人聯絡地址"
   
   'Add By Cheng 2002/02/18
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "申請人編號"
   Printer.CurrentX = PLeft(6) + 1200
   Printer.CurrentY = iPrint
   Printer.Print "TEL1"
   Printer.CurrentX = PLeft(6) + 2800 + 600
   Printer.CurrentY = iPrint
   Printer.Print "TEL2"
   Printer.CurrentX = PLeft(6) + 4400 + 600
   Printer.CurrentY = iPrint
   Printer.Print "FAX1"
   Printer.CurrentX = PLeft(6) + 6000 + 600
   Printer.CurrentY = iPrint
   Printer.Print "FAX2"
   
   iPrint = iPrint + 300
   Printer.Font.Size = 12
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.Font.Size = 10
End Sub

Sub PrintDatil2()
   For i = 1 To 10
       Printer.CurrentX = PLeft(i)
       Printer.CurrentY = iPrint
       Printer.Print strTemp(i)
   Next i
   'add by nickc 2007/06/13
   If InStr(1, txt1(1), "716") <> 0 Then
       Printer.CurrentX = PLeft(7) + 1500
       Printer.CurrentY = iPrint
       Printer.Print strTemp(17)
   End If
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(11)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(11)
   
   'Add By Cheng 2002/02/18
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(12)
   Printer.CurrentX = PLeft(6) + 1200
   Printer.CurrentY = iPrint
   Printer.Print strTemp(13)
   Printer.CurrentX = PLeft(6) + 2800
   Printer.CurrentY = iPrint
   Printer.Print strTemp(14)
   Printer.CurrentX = PLeft(6) + 4400 + 600
   Printer.CurrentY = iPrint
   Printer.Print strTemp(15)
   Printer.CurrentX = PLeft(6) + 6000 + 600
   Printer.CurrentY = iPrint
   Printer.Print strTemp(16)
   
   iPrint = iPrint + 300
End Sub

Sub GetPleft2()
   Erase PLeft
   PLeft(0) = 500
   PLeft(1) = 500
   'Modified by Lydia 2023/12/04 智權人員與案件性質的抬頭有重疊
   'PLeft(2) = 1200
   PLeft(2) = 1400
   PLeft(3) = 2500
   PLeft(4) = 4900
   PLeft(5) = 6500
   PLeft(6) = 8000
   PLeft(7) = 9000
   'edit by nickc 2007/06/13
   If InStr(1, txt1(1), "716") <> 0 Then
       PLeft(8) = 12500
       PLeft(9) = 14000
       PLeft(10) = 15000
       PLeft(11) = 1500
   Else
       PLeft(8) = 11500
       PLeft(9) = 13500
       PLeft(10) = 14500
       'Modify By Cheng 2002/02/19
       'PLeft(11) = 2000
       'Modified by Lydia 2023/12/04 智權人員與案件性質的抬頭有重疊
       'PLeft(11) = 1200
       PLeft(11) = 1400
   End If
End Sub

Function Process1() As Boolean          '智權人員    不含102 ,702, 715, 716, 708  2010/3/22sonia再加不含109
'Add By Sindy 2011/1/11
Dim strTmp, strNP02 As String, strNP03 As String, strNP04 As String, strNP05 As String
Dim intRow As Integer
'2011/1/11 End
'Add By Sindy 2012/6/29
'記錄本所案號
Dim strTM01 As String
Dim strTM02 As String
Dim strTM03 As String
Dim strTM04 As String
Dim arrCaseNo
'2012/6/29 End
   
   'Modify By Cheng 2003/02/05
   '若無審定號數抓申請案號, 承辦人員為內商人員
   'strSQL = "SELECT s1.st03,np10,NP08,NP09,NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM09,TM15,decode(tm10,'000',cpm03,cpm04),NP15,DECODE(CP22,'Y','是','N','否',NULL,'是'),S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(NA03,NA04),CP09,CP27,np22,s1.st03 as a FROM NEXTPROGRESS,TRADEMARK,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=s1.ST01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) AND NP07<>102 AND NP07<>702  " & StrSQL6 & strSQL1
   'strSQL = strSQL + " union all select s1.st03,np10,NP08,NP09,NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),' ',SP11,decode(sp09,'000',CPM03,CPM04),NP15,DECODE(CP22,'Y','是','N','否',NULL,'是'),S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(NA03,NA04),CP09,CP27,np22,s1.st03 as a FROM NEXTPROGRESS,SERVICEPRACTICE,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=s1.ST01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND SP09=NA01(+) AND NP07<>102 AND NP07<>702 " & StrSQL6 & strSQL2
   'Modify By Cheng 2003/03/11
   'strSQL = "SELECT s1.st03,np10,NP08,NP09,NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(TM15,TM12),decode(tm10,'000',cpm03,cpm04),NP15,DECODE(CP22,'Y','是','N','否',NULL,'是'),S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(NA03,NA04),CP09,CP27,np22,s1.st03 as a FROM NEXTPROGRESS,TRADEMARK,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=s1.ST01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) And (S2.ST03>='P20' AND S2.ST03<='P29') AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) AND ( NP07<>102 AND NP07<>702 AND NP07<>305 AND NP07<>997 AND NP07<>998 ) " & StrSQL6 & strSQL1
   'strSQL = strSQL + " union all select s1.st03,np10,NP08,NP09,NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),' ',SP11,decode(sp09,'000',CPM03,CPM04),NP15,DECODE(CP22,'Y','是','N','否',NULL,'是'),S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(NA03,NA04),CP09,CP27,np22,s1.st03 as a FROM NEXTPROGRESS,SERVICEPRACTICE,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=s1.ST01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) And (S2.ST03>='P20' AND S2.ST03<='P29') AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND SP09=NA01(+) AND ( NP07<>102 AND NP07<>702 AND NP07<>305 AND NP07<>997 AND NP07<>998 ) " & StrSQL6 & strSQL2
   'Modify By Cheng 2003/10/02
   'Begin
   'strSQL = "SELECT s1.ST15,np10,NP08,NP09,NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(TM15,TM12),decode(tm10,'000',cpm03,cpm04),NP15,DECODE(CP22,'Y','是','N','否',NULL,'是'),S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(NA03,NA04),CP09,CP27,np22,s1.ST15 as a FROM NEXTPROGRESS,TRADEMARK,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=s1.ST01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) And (S2.ST15>='P20' AND S2.ST15<='P29') AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) AND ( NP07<>102 AND NP07<>702 AND NP07<>305 AND NP07<>997 AND NP07<>998 ) " & StrSQL6 & strSQL1
   'strSQL = strSQL + " union all select s1.ST15,np10,NP08,NP09,NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),' ',SP11,decode(sp09,'000',CPM03,CPM04),NP15,DECODE(CP22,'Y','是','N','否',NULL,'是'),S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(NA03,NA04),CP09,CP27,np22,s1.ST15 as a FROM NEXTPROGRESS,SERVICEPRACTICE,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=s1.ST01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) And (S2.ST15>='P20' AND S2.ST15<='P29') AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND SP09=NA01(+) AND ( NP07<>102 AND NP07<>702 AND NP07<>305 AND NP07<>997 AND NP07<>998 ) " & StrSQL6 & strSQL2
   'edit by nick 2004/07/06 將 715和716 加入處理
   'strSQL = "SELECT s1.ST15,np10,NP08,NP09,NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(TM15,TM12),decode(tm10,'000',cpm03,cpm04),NP15,DECODE(CP22,'Y','是','N','否',NULL,'是'),S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(NA03,NA04),CP09,CP27,np22,s1.ST15 as a FROM Staff_Group, NEXTPROGRESS,TRADEMARK,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=s1.ST01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) And (S2.ST15>='P20' AND S2.ST15<='P29') AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 AND ( NP07<>102 AND NP07<>702 And NP07<>715 And NP07<>716 And NP07<>708 AND NP07<>305 AND NP07<>997 AND NP07<>998 ) " & StrSQL6 & strSQL1
   'strSQL = strSQL + " union all select s1.ST15,np10,NP08,NP09,NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),' ',SP11,decode(sp09,'000',CPM03,CPM04),NP15,DECODE(CP22,'Y','是','N','否',NULL,'是'),S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(NA03,NA04),CP09,CP27,np22,s1.ST15 as a FROM Staff_Group, NEXTPROGRESS,SERVICEPRACTICE,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=s1.ST01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) And (S2.ST15>='P20' AND S2.ST15<='P29') AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND SP09=NA01(+) And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 And  ( NP07<>102 AND NP07<>702 And NP07<>715 And NP07<>716 And NP07<>708 AND NP07<>305 AND NP07<>997 AND NP07<>998 ) " & StrSQL6 & strSQL2
   'edit by nickc 2006/05/30 將備註改成印分所案號 葉大說的
   'strSQL = "SELECT s1.ST15,np10,NP08,NP09,NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(TM15,TM12),decode(tm10,'000',cpm03,cpm04),NP15,DECODE(CP22,'Y','是','N','否',NULL,'是'),S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(NA03,NA04),CP09,CP27,np22,s1.ST15 as a FROM Staff_Group, NEXTPROGRESS,TRADEMARK,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=s1.ST01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) And (S2.ST15>='P20' AND S2.ST15<='P29') AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 AND ( NP07<>102 AND NP07<>702 And NP07<>708 AND NP07<>305 AND NP07<>997 AND NP07<>998 ) " & StrSQL6 & strSQL1
   'strSQL = strSQL + " union all select s1.ST15,np10,NP08,NP09,NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),' ',SP11,decode(sp09,'000',CPM03,CPM04),NP15,DECODE(CP22,'Y','是','N','否',NULL,'是'),S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(NA03,NA04),CP09,CP27,np22,s1.ST15 as a FROM Staff_Group, NEXTPROGRESS,SERVICEPRACTICE,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=s1.ST01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) And (S2.ST15>='P20' AND S2.ST15<='P29') AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND SP09=NA01(+) And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 And  ( NP07<>102 AND NP07<>702  And NP07<>708 AND NP07<>305 AND NP07<>997 AND NP07<>998 ) " & StrSQL6 & strSQL2
   'edit by nickc 2006/07/07
   'strSQL = "SELECT s1.ST15,np10,NP08,NP09,NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(TM15,TM12),decode(tm10,'000',cpm03,cpm04),tm34,DECODE(CP22,'Y','是','N','否',NULL,'是'),S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(NA03,NA04),CP09,CP27,np22,s1.ST15 as a FROM Staff_Group, NEXTPROGRESS,TRADEMARK,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=s1.ST01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) And (S2.ST15>='P20' AND S2.ST15<='P29') AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 AND ( NP07<>102 AND NP07<>702 And NP07<>708 AND NP07<>305 AND NP07<>997 AND NP07<>998 ) " & StrSQL6 & strSQL1
   'strSQL = strSQL + " union all select s1.ST15,np10,NP08,NP09,NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),' ',SP11,decode(sp09,'000',CPM03,CPM04),sp28,DECODE(CP22,'Y','是','N','否',NULL,'是'),S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(NA03,NA04),CP09,CP27,np22,s1.ST15 as a FROM Staff_Group, NEXTPROGRESS,SERVICEPRACTICE,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=s1.ST01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) And (S2.ST15>='P20' AND S2.ST15<='P29') AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND SP09=NA01(+) And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 And  ( NP07<>102 AND NP07<>702  And NP07<>708 AND NP07<>305 AND NP07<>997 AND NP07<>998 ) " & StrSQL6 & strSQL2
   '2010/3/22 modify by sonia 下一程序的案件性質加入不含大陸被異議續展(109)
   '2010/3/23 MODIFY BY SONIA 剔除下一程序非智權人員掌控之案件性質改以strNpSqlOfNoSalesDuty控制
   'Modify by Amy 2018/09/11 + iif
   strSql = "SELECT s1.ST15,np10,NP08,NP09,NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(TM15,TM12),decode(tm10,'000',cpm03,cpm04),tm34,DECODE(CP22,'Y','是','N','否',NULL,'是'),S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(NA03,NA04),CP09,CP27,np22,s1.ST04 as ST04 FROM Staff_Group, NEXTPROGRESS,TRADEMARK,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER" & IIf(bolT102Again = True, ",Staff s3", "") & " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=s1.ST01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) And (S2.ST15>='P20' AND S2.ST15<='P29') AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 AND ( NP07<>102 AND NP07<>702 And NP07<>708 and np07<>716 and np07<>109 ) " & StrSQL6 & strSQL1 & strNpSqlOfNoSalesDuty
   strSql = strSql + " union all select s1.ST15,np10,NP08,NP09,NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),' ',SP11,decode(sp09,'000',CPM03,CPM04),sp28,DECODE(CP22,'Y','是','N','否',NULL,'是'),S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(NA03,NA04),CP09,CP27,np22,s1.ST04 as ST04 FROM Staff_Group, NEXTPROGRESS,SERVICEPRACTICE,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER" & IIf(bolT102Again = True, ",Staff s3", "") & " WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=s1.ST01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) And (S2.ST15>='P20' AND S2.ST15<='P29') AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND SP09=NA01(+) And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 And  ( NP07<>102 AND NP07<>702  And NP07<>708 and np07<>716 and np07<>109 ) " & StrSQL6 & strSQL2 & strNpSqlOfNoSalesDuty
   'end 2018/09/11
   'Add by Amy 2018/08/31 +大陸撤三
   If Option1(1).Value = True Then
      strSql = "Select s1.ST15,CU13 np10,to_number(SubStr(TM21,1,4))+3||SubStr(TM21,5) NP08,'' NP09,Decode(cp09,null,'','*')||tm01||'-'||tm02||'-'||tm03||'-'||tm04,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(TM15,TM12),'',tm34,'','',NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(NA03,NA04),'','','' np22,s1.ST04 as ST04 FROM TradeMark,CaseProgress,Nation,Staff s1,Customer " & _
                  "Where  TM01=cp01(+) And tm02=cp02(+) And TM03=cp03(+) And TM04=cp04(+) And cu13=s1.ST01(+)  And SUBSTR(TM23,1,8)=CU01(+) And decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) And TM10=NA01(+)  And TM21 is not null And '1729'=CP10(+) " & strSQL1
   End If
   'end 2018/08/31
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   intRow = 0 'Add By Sindy 2011/1/11
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'       InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/9/30
       With adoRecordset
           .MoveFirst
           'Add By Sindy 2012/6/29
           strTM01 = ""
           strTM02 = ""
           strTM03 = ""
           strTM04 = ""
           '2012/6/29 End
           Do While .EOF = False
               For i = 0 To 15
                   strTemp(i) = "" & CheckStr(.Fields(i))
               Next i
               
               'Add By Sindy 2011/1/11 抓未收文資料時, 若np10智權人員為離職時, 改用PUB_GetAKindSalesNo抓智權人員
               '2011/4/13 modify by sonia 不下智權人員才做,否則下離職智權人員條件時會因已轉換而無資料
               'If CheckStr(.Fields("ST04")) <> "1" Then
               If CheckStr(.Fields("ST04")) <> "1" And Len(txt1(7)) = 0 Then
                  strTmp = Split(strTemp(4), "-")
                  strNP02 = strTmp(0)
                  strNP03 = strTmp(1)
                  strNP04 = strTmp(2)
                  strNP05 = strTmp(3)
                  strTemp(0) = PUB_GetStaffST15(PUB_GetAKindSalesNo(strNP02, strNP03, strNP04, strNP05), "1")
                  strTemp(1) = PUB_GetAKindSalesNo(strNP02, strNP03, strNP04, strNP05)
               End If
               If Len(txt1(5)) <> 0 Then
                  If strTemp(0) < txt1(5) Then GoTo GoToExit2
               End If
               If Len(txt1(6)) <> 0 Then
                  If strTemp(0) > txt1(6) Then GoTo GoToExit2
               End If
               If Len(txt1(7)) <> 0 Then
                  If strTemp(1) <> Trim(txt1(7)) Then GoTo GoToExit2
               End If
               If txt1(16) = "1" Then
                  If Left(strTemp(0), 1) = "S" Then GoTo GoToExit2
               End If
               intRow = intRow + 1 '記錄筆數
               '2011/1/11 End
               
               'Add by Amy 2018/08/31 +if 非大陸撤三
               If Option1(1).Value = True Then
                    strTemp(3) = ""
                    If "" & strTemp(2) <> MsgText(601) Then
                        strTemp(2) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(2)))
                    End If
               Else
                    strTemp(3) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
                    If strTemp(2) < GetTodayDate Then
                        strTemp(2) = "*" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(2)))
                    Else
                        If strTemp(2) = GetTodayDate Then
                            strTemp(2) = "V" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(2)))
                        Else
                            If Mid(strTemp(14), 1, 1) = "C" And Len(strTemp(15)) = 0 Then
                                strTemp(2) = "#" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(2)))
                            Else
                                strTemp(2) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(2)))
                            End If
                        End If
                    End If
               End If
               'end 2018/08/31
               
               'Modify By Sindy 2012/6/29 馬德里申請國家相同(因申請多個商品類別時會多案號)只需要出一份定稿
               arrCaseNo = Split(strTemp(4), "-")
               If Not (arrCaseNo(0) = "TF" And strTemp(8) = "使用宣誓" And strTM02 = arrCaseNo(1) And strTM04 = arrCaseNo(3)) Then
                  'If arrCaseNo(0) = "TF" And strTemp(8) = "使用宣誓" Then strTemp(4) = arrCaseNo(0) & Left(arrCaseNo(1), 5) & "0000" '顯示母案
               '2012/6/29 End
                  strSql = "INSERT INTO R020301_1 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & strUserNum & "') "
                  cnnConnection.Execute strSql
               End If
               'Add By Sindy 2012/6/29
               strTM01 = arrCaseNo(0)
               strTM02 = arrCaseNo(1)
               strTM03 = arrCaseNo(2)
               strTM04 = arrCaseNo(3)
               '2012/6/29 End
GoToExit2:
               .MoveNext
               DoEvents
           Loop
       End With
       'Add By Sindy 2011/1/11
       If intRow = 0 Then
         InsertQueryLog (0) 'Add By Sindy 2010/9/30
         ShowNoData
         CheckOC
         Process1 = False
         Exit Function
       Else
         InsertQueryLog (intRow) 'Add By Sindy 2010/9/30
       End If
       '2011/1/11 End
   Else
       '若不列印延展, 刊登廣告, 第一期註冊費, 第二期註冊費
       'edit by nick 2004/07/20
       'If txt1(2) <> "Y" And Text1 <> "Y" And Me.txt1(16).Text <> "Y" And Me.txt1(17).Text <> "Y" And Me.txt1(18).Text <> "Y" Then
       'edit by nickc 2007/03/29
       'If txt1(2) <> "Y" And Text1 <> "Y" And Me.txt1(18).Text <> "Y" And InStr(1, txt1(1), "716") = 0 Then
       If txt1(2) <> "Y" And Text1 <> "Y" And Me.txt1(18).Text <> "Y" And InStr(1, txt1(1), "716") = 0 And txt1(19).Text <> "Y" Then
           InsertQueryLog (0) 'Add By Sindy 2010/9/30
           ShowNoData
       End If
       CheckOC
       Process1 = False
       Exit Function
   End If
   CheckOC
   Process1 = True
End Function

'延展,刊登廣告,第一期註冊費,第二期註冊費,繳年費     只有 102,702,715,716,708
'2010/3/22 再加大陸被異議續展109
Function Process2() As Boolean
'Add By Sindy 2011/1/11
Dim strTmp, strNP02 As String, strNP03 As String, strNP04 As String, strNP05 As String
Dim intRow As Integer
'2011/1/11 End
   
   'Modify By Cheng 2003/11/20
   ''910725 Sieg
   'If TXT1(2) = "Y" And Text1 = "Y" Then
   '  StrSQL6 = "AND NP07 IN (102,702)"
   'ElseIf TXT1(2) = "Y" And Text1 <> "Y" Then
   '  StrSQL6 = "AND NP07=102"
   'ElseIf TXT1(2) <> "Y" And Text1 = "Y" Then
   '  StrSQL6 = "AND NP07=702"
   'Else
   '    StrSQL6 = ""
   'End If
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/9/30 清除查詢印表記錄檔欄位
   pub_QL05 = pub_QL05 & ";" & Label1(6) & "管制表" 'Add By Sindy 2010/9/30
   StrSQL6 = ""
   If Me.txt1(2).Text = "Y" Then
       '2010/3/22 modify by sonia 加大陸被異議續展109
       StrSQL6 = StrSQL6 & "102,109,"
       pub_QL05 = pub_QL05 & ";" & Left(Label1(2), 6) & "含"  'Add By Sindy 2010/9/30
   End If
   If Me.Text1.Text = "Y" Then
       StrSQL6 = StrSQL6 & "702,"
       pub_QL05 = pub_QL05 & ";" & Left(Label1(10), 8) & "含"  'Add By Sindy 2010/9/30
   End If
   'add by nickc 2007/03/29
   If txt1(19).Text = "Y" Then
       StrSQL6 = StrSQL6 & "716,"
       pub_QL05 = pub_QL05 & ";" & Left(Label1(16), 7) & "含" 'Add By Sindy 2010/9/30
   End If
   'edit by nick 2004/07/06 移到 process1
   'If Me.txt1(16).Text = "Y" Then
   '    StrSQL6 = StrSQL6 & "715,"
   'End If
   'If Me.txt1(17).Text = "Y" Then
   '    StrSQL6 = StrSQL6 & "716,"
   'End If
   If Me.txt1(18).Text = "Y" Then
       StrSQL6 = StrSQL6 & "708,"
       pub_QL05 = pub_QL05 & ";" & Left(Label1(15), 11) & "含" 'Add By Sindy 2010/9/30
   End If
   
   'add by nickc 2006/07/07
   If InStr(1, txt1(1), "716") <> 0 Then
       StrSQL6 = StrSQL6 & "716,"
   End If
   If StrSQL6 <> "" Then
       StrSQL6 = Left(StrSQL6, Len(StrSQL6) - 1)
       StrSQL6 = " And NP07 In (" & StrSQL6 & ") "
   End If
   If Len(txt1(1)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1)  'Add By Sindy 2010/9/30
   End If
   
   If Len(txt1(3)) <> 0 Then
      StrSQL6 = StrSQL6 + " AND NP08>=" & Val(ChangeTStringToWString(txt1(3))) & " "
   End If
   If Len(txt1(4)) <> 0 Then
      StrSQL6 = StrSQL6 + " AND NP08<=" & Val(ChangeTStringToWString(txt1(4))) & " "
   End If
   If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Then
      'Modify by Amy 2019/03/08 原:Label1(3)-bug
      pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & txt1(3) & "-" & txt1(4)  'Add By Sindy 2010/9/30
   End If
   StrSQL6 = StrSQL6 & " AND (NP06 IS NULL OR NP06='') "
'   If Len(TXT1(5)) <> 0 Then
'       'Modify By Cheng 2003/03/11
'   '    StrSQL6 = StrSQL6 + " AND s1.ST03>='" & TXT1(5) & "' "
'       StrSQL6 = StrSQL6 + " AND s1.ST15>='" & TXT1(5) & "' "
'   End If
'   If Len(TXT1(6)) <> 0 Then
'       'Modify By Cheng 2003/03/15
'   '    StrSQL6 = StrSQL6 + " AND s1.ST03<='" & TXT1(6) & "' "
'       StrSQL6 = StrSQL6 + " AND s1.ST15<='" & TXT1(6) & "' "
'   End If
   If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(5) & "-" & txt1(6)  'Add By Sindy 2010/9/30
   End If
   If Len(txt1(7)) <> 0 Then
'       StrSQL6 = StrSQL6 + " AND NP10='" & TXT1(7) & "' "
       pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(7) & lbl1  'Add By Sindy 2010/9/30
   End If
   
   'Add By Sindy 2015/4/17 台灣案催延展對象
   '大->台係指葉經理及巨京收文且申請人國籍非台灣者,
   '其他案件皆屬台->台範圍
   If Trim(txt1(13)) = "000" And Trim(txt1(14)) = "000" Then
      '系統別有T者
      strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
      s = 0
      For i = 0 To UBound(strTemp2)
         If strTemp2(i) = "T" Then
            s = 1
            Exit For
         End If
      Next i
      If s = 1 And txt1(2) = "Y" Then 'T含延展
         If Trim(txt1(20)) = "1" Then '1.台->台
            'Modify By Sindy 2020/8/13 改判斷 TM44.FC代理人
            'StrSQL6 = StrSQL6 + " AND not(Substr(NP10,1,5) in('67002','96029','96030','MCTF0') AND CU10>'010') "
            strSQL1 = strSQL1 + " AND not(Substr(NP10,1,5) in('67002','96029','96030','MCTF0') AND TM44 is not null) "
            strSQL2 = strSQL2 + " AND not(Substr(NP10,1,5) in('67002','96029','96030','MCTF0') AND SP26 is not null) "
         ElseIf Trim(txt1(20)) = "2" Then '2.大->台
            'Modify by Amy 2017/01/10 +MCTF特殊人員
            'Modify By Sindy 2020/8/13 改判斷 TM44.FC代理人
            'StrSQL6 = StrSQL6 + " AND SubStr(NP10,1,5) in('67002','96029','96030','MCTF0') AND CU10>'010' "
            strSQL1 = strSQL1 + " AND SubStr(NP10,1,5) in('67002','96029','96030','MCTF0') AND TM44 is not null "
            strSQL2 = strSQL2 + " AND SubStr(NP10,1,5) in('67002','96029','96030','MCTF0') AND SP26 is not null "
         End If
         pub_QL05 = pub_QL05 & ";" & Label1(19) & txt1(20) & Label1(18)
      End If
   End If
   '2015/4/17 END
   
   'add by nickc 2006/05/30
   If txt1(16) = "1" Then
'       StrSQL6 = StrSQL6 + " AND substr(s1.ST15,1,1)<>'S' "
       pub_QL05 = pub_QL05 & ";" & Label1(13) & "非智權部同仁"   'Add By Sindy 2010/9/30
   End If
   'Modify By Cheng 2003/03/15
   '智權人員業務區改抓ST15
   'strSQL = "SELECT s1.st03,np10,decode(tm10,'000',CPM03,CPM04),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06))," & _
   '   "NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM15,TM09,TM21,TM22,NVL(NA03,NA04),TM27," & _
   '   "NVL(CU23,NVL(CU24||CU25||CU26||CU27||CU28,CU29)),CP09,CP27,np22,CU01||CU02,CU16,CU17,CU18,CU19 " & _
   '         " FROM NEXTPROGRESS,TRADEMARK,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER " & _
   '         " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=s1.ST01(+) AND " & _
   '         "NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND " & _
   '         "decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) " & StrSQL6 & strSQL1
   'strSQL = strSQL + " union all " & _
   '         "SELECT s1.st03,np10,decode(SP09,'000',CPM03,CPM04),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06))," & _
   '   "NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),SP11,'',SP20,SP21,NVL(NA03,NA04),SP32," & _
   '   "NVL(CU23,NVL(CU24||CU25||CU26||CU27||CU28,CU29)),CP09,CP27,np22,CU01||CU02,CU16,CU17,CU18,CU19 " & _
   '         " FROM NEXTPROGRESS,SERVICEPRACTICE,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER " & _
   '         " WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=s1.ST01(+) AND " & _
   '         "NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND " & _
   '         "decode(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SP09=NA01(+) " & StrSQL6 & strSQL2
   'Modify By Cheng 2003/10/02
   'Begin
   'strSQL = "SELECT s1.ST15,np10,decode(tm10,'000',CPM03,CPM04),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06))," & _
   '   "NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM15,TM09,TM21,TM22,NVL(NA03,NA04),TM27," & _
   '   "NVL(CU23,NVL(CU24||CU25||CU26||CU27||CU28,CU29)),CP09,CP27,np22,CU01||CU02,CU16,CU17,CU18,CU19 " & _
   '         " FROM NEXTPROGRESS,TRADEMARK,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER " & _
   '         " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=s1.ST01(+) AND " & _
   '         "NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND " & _
   '         "decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) " & StrSQL6 & strSQL1
   'strSQL = strSQL + " union all " & _
   '         "SELECT s1.ST15,np10,decode(SP09,'000',CPM03,CPM04),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06))," & _
   '   "NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),SP11,'',SP20,SP21,NVL(NA03,NA04),SP32," & _
   '   "NVL(CU23,NVL(CU24||CU25||CU26||CU27||CU28,CU29)),CP09,CP27,np22,CU01||CU02,CU16,CU17,CU18,CU19 " & _
   '         " FROM NEXTPROGRESS,SERVICEPRACTICE,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER " & _
   '         " WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=s1.ST01(+) AND " & _
   '         "NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND " & _
   '         "decode(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SP09=NA01(+) " & StrSQL6 & strSQL2
   'edit by nickc 2006/05/30 將正商標號數改成印分所案號 葉大說的
   'strSQL = "SELECT s1.ST15,np10,decode(tm10,'000',CPM03,CPM04),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06))," & _
   '   "NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM15,TM09,TM21,TM22,NVL(NA03,NA04),TM27," & _
   '   "NVL(CU23,NVL(CU24||CU25||CU26||CU27||CU28,CU29)),CP09,CP27,np22,CU01||CU02,CU16,CU17,CU18,CU19 " & _
   '         " FROM Staff_Group, NEXTPROGRESS,TRADEMARK,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER " & _
   '         " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=s1.ST01(+) AND " & _
   '         "NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND " & _
   '         "decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 " & StrSQL6 & strSQL1
   'strSQL = strSQL + " union all " & _
   '         "SELECT s1.ST15,np10,decode(SP09,'000',CPM03,CPM04),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06))," & _
   '   "NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),SP11,'',SP20,SP21,NVL(NA03,NA04),SP32," & _
   '   "NVL(CU23,NVL(CU24||CU25||CU26||CU27||CU28,CU29)),CP09,CP27,np22,CU01||CU02,CU16,CU17,CU18,CU19 " & _
   '         " FROM Staff_Group, NEXTPROGRESS,SERVICEPRACTICE,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER " & _
   '         " WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=s1.ST01(+) AND " & _
   '         "NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND " & _
   '         "decode(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SP09=NA01(+) And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 " & StrSQL6 & strSQL2
   
   'edit by nickc 2006/08/25 若是716 時，專用期間改成法定期限
   'strSQL = "SELECT s1.ST15,np10,decode(tm10,'000',CPM03,CPM04),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06))," & _
   '   "NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM15,TM09,TM21,TM22,NVL(NA03,NA04),TM34," & _
   '   "NVL(CU23,NVL(CU24||CU25||CU26||CU27||CU28,CU29)),CP09,CP27,np22,CU01||CU02,CU16,CU17,CU18,CU19 " & _
   '         " FROM Staff_Group, NEXTPROGRESS,TRADEMARK,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER " & _
   '         " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=s1.ST01(+) AND " & _
   '         "NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND " & _
   '         "decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 " & StrSQL6 & strSQL1
   'strSQL = strSQL + " union all " & _
   '         "SELECT s1.ST15,np10,decode(SP09,'000',CPM03,CPM04),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06))," & _
   '   "NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),SP11,'',SP20,SP21,NVL(NA03,NA04),SP28," & _
   '   "NVL(CU23,NVL(CU24||CU25||CU26||CU27||CU28,CU29)),CP09,CP27,np22,CU01||CU02,CU16,CU17,CU18,CU19 " & _
   '         " FROM Staff_Group, NEXTPROGRESS,SERVICEPRACTICE,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER " & _
   '         " WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=s1.ST01(+) AND " & _
   '         "NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND " & _
   '         "decode(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SP09=NA01(+) And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 " & StrSQL6 & strSQL2
   '2011/1/10 modify by sonia地址先抓聯絡地址
   'Modify by Amy 2018/09/11 +iif
   strSql = "SELECT s1.ST15,np10,decode(tm10,'000',CPM03,CPM04),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06))," & _
      "NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM15,TM09," & IIf(InStr(1, txt1(1), "716") <> 0, "'' as tm21,sqldatet(np09) as tm22 ", " TM21,TM22") & ",NVL(NA03,NA04),TM34," & _
      "nvl(cu31,NVL(CU23,NVL(CU24||CU25||CU26||CU27||CU28,CU29))),CP09,CP27,np22,CU01||CU02,CU16,CU17,CU18,CU19,s1.ST04 as ST04 " & _
            " FROM Staff_Group, NEXTPROGRESS,TRADEMARK,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER " & IIf(bolT102Again, ",Staff s3", "") & _
            " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=s1.ST01(+) AND " & _
            "NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND " & _
            "decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 " & StrSQL6 & strSQL1
   'Modify by Sindy 2019/11/4 非大陸撤三
   If Option1(1).Value = False Then
   '2019/11/4 END
      strSql = strSql + " union all " & _
               "SELECT s1.ST15,np10,decode(SP09,'000',CPM03,CPM04),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06))," & _
         "NP02||'-'||nP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),SP11,''," & IIf(InStr(1, txt1(1), "716") <> 0, "'' as sp20 ,sqldatet(np09) as sp21 ", " SP20,SP21") & ",NVL(NA03,NA04),SP28," & _
         "nvl(cu31,NVL(CU23,NVL(CU24||CU25||CU26||CU27||CU28,CU29))),CP09,CP27,np22,CU01||CU02,CU16,CU17,CU18,CU19,s1.ST04 as ST04 " & _
               " FROM Staff_Group, NEXTPROGRESS,SERVICEPRACTICE,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER " & IIf(bolT102Again, ",Staff s3", "") & _
               " WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=s1.ST01(+) AND " & _
               "NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND " & _
               "decode(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SP09=NA01(+) And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 " & StrSQL6 & strSQL2
   End If
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   intRow = 0 'Add By Sindy 2011/1/11
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'       InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/9/30
       With adoRecordset
           .MoveFirst
           Do While .EOF = False
               'Modify By Cheng 2002/02/18
   '            For i = 0 To 14
               For i = 0 To 20
                   strTemp(i) = CheckStr(.Fields(i))
               Next i
               
               'Add By Sindy 2011/1/11 抓未收文資料時, 若np10智權人員為離職時, 改用PUB_GetAKindSalesNo抓智權人員
               '2011/4/13 modify by sonia 不下智權人員才做,否則下離職智權人員條件時會因已轉換而無資料
               'If CheckStr(.Fields("ST04")) <> "1" Then
               If CheckStr(.Fields("ST04")) <> "1" And Len(txt1(7)) = 0 Then
                  strTmp = Split(strTemp(4), "-")
                  strNP02 = strTmp(0)
                  strNP03 = strTmp(1)
                  strNP04 = strTmp(2)
                  strNP05 = strTmp(3)
                  strTemp(0) = PUB_GetStaffST15(PUB_GetAKindSalesNo(strNP02, strNP03, strNP04, strNP05), "1")
                  strTemp(1) = PUB_GetAKindSalesNo(strNP02, strNP03, strNP04, strNP05)
               End If
               If Len(txt1(5)) <> 0 Then
                  If strTemp(0) < txt1(5) Then GoTo gotoExit
               End If
               If Len(txt1(6)) <> 0 Then
                  If strTemp(0) > txt1(6) Then GoTo gotoExit
               End If
               If Len(txt1(7)) <> 0 Then
                  If strTemp(1) <> Trim(txt1(7)) Then GoTo gotoExit
               End If
               If txt1(16) = "1" Then
                  If Left(strTemp(0), 1) = "S" Then GoTo gotoExit
               End If
               intRow = intRow + 1 '記錄筆數
               '2011/1/11 End
               
               SavDay1 = ""
               'edit by nickc 2006/08/25 716 格式與一般不同
               If InStr(1, txt1(1), "716") <> 0 Then
                   SavDay1 = "1"
               Else
                   '若系統日在專用期間
                   If strTemp(8) >= GetTodayDate And strTemp(9) <= GetTodayDate Then
                       SavDay1 = "2"
                   '若系統日不在專用期間
                   Else
                       SavDay1 = "1"
                   End If
               End If
               'edit by nickc 2006/08/25 716 格式與一般不同
               If InStr(1, txt1(1), "716") <> 0 Then
                   strTemp(8) = strTemp(9)
               Else
                   strTemp(8) = strTemp(8) & "-" & strTemp(9)
               End If
               strTemp(9) = strTemp(10)
               strTemp(10) = strTemp(11)
               strTemp(11) = strTemp(12)
               
               'Add By Cheng 2002/02/15
               '多加欄位--申請人編號, TEL1, TEL2, FAX1, TAX2
               strTemp(12) = strTemp(16)
               strTemp(13) = strTemp(17)
               strTemp(14) = strTemp(18)
               strTemp(15) = strTemp(19)
               strTemp(16) = strTemp(20)
               
               'Modify By Cheng 2002/02/15
   '            strSQL = "INSERT INTO R020301_2 VALUES('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "','" & chgsql(strTemp(2)) & "','" & chgsql(strTemp(3)) & "','" & chgsql(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & chgsql(strTemp(6)) & "','" & chgsql(strTemp(7)) & "','" & chgsql(strTemp(8)) & "','" & chgsql(strTemp(9)) & "','" & chgsql(strTemp(10)) & "','" & chgsql(strTemp(11)) & "','" & strUserNum & SavDay1 & "') "
               strSql = "INSERT INTO R020301_2 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & strUserNum & SavDay1 & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & ChgSQL(strTemp(16)) & "' ) "
               cnnConnection.Execute strSql
gotoExit:
               .MoveNext
               DoEvents
           Loop
       End With
       'Add By Sindy 2011/1/11
       If intRow = 0 Then
         InsertQueryLog (0) 'Add By Sindy 2010/9/30
         ShowNoData
         CheckOC
         Process2 = False
         Exit Function
       Else
         InsertQueryLog (intRow) 'Add By Sindy 2010/9/30
       End If
       '2011/1/11 End
   Else
       InsertQueryLog (0) 'Add By Sindy 2010/9/30
       ShowNoData
       CheckOC
       Process2 = False
       Exit Function
   End If
   CheckOC
   Process2 = True
End Function

Private Sub Form_Load()
Dim PrinterIndex As Integer
   
   MoveFormToCenter Me
   txt1(0) = GetSystemKindByNick
   lblMsg.Caption = "" 'Add by Amy 2018/09/11
   
   'Add By Sindy 2019/11/1
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum
   '檢查是否有安裝PDFCreator
   PrinterIndex = -1
   For i = 0 To Printers.Count - 1
    If UCase(Printers(i).DeviceName) = UCase$("PDFCreator") Then
     PrinterIndex = i
     Exit For
    End If
   Next i
   If PrinterIndex < 0 Then
      MsgBox "請通知電腦中心安裝PDFCreator !!!"
      Exit Sub
   End If
   '2019/11/1 END
   
   Me.txt1(17).Text = strSrvDate(2) 'Add By Sindy 2024/21
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm020301 = Nothing
End Sub

'Add by Amy 2018/08/31
Private Sub Option1_Click(Index As Integer)
    If Option1(Index).Value = True Then
        '本所期限
        If Index = 0 Then
            txt1(3).Locked = False
            txt1(4).Locked = False
            txt1(21) = "": txt1(22) = ""
            txt1(21).Locked = True
            txt1(22).Locked = True
            txt1(3).SetFocus
        '大陸撤三
        Else
            txt1(21).Locked = False
            txt1(22).Locked = False
            txt1(3) = "": txt1(4) = ""
            txt1(3).Locked = True
            txt1(4).Locked = True
            txt1(21).SetFocus
        End If
    End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdok(0).SetFocus
End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    'Modify By Cheng 2003/11/20
'   If Index = 2 Then
   'edit by nickc 2007/03/29 加入第二期
   'If Index = 2 Or Index = 17 Or Index = 18 Then
   If Index = 2 Or Index = 18 Or Index = 19 Then
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         KeyAscii = 0
         Beep
      End If
   End If
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Select Case Index
   Case 0
        strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
        strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
        For i = 0 To UBound(strTemp2)
           s = 0
           For j = 0 To UBound(strTemp1)
               If strTemp2(i) = strTemp1(j) Then
                   s = 1
                   Exit For
               End If
           Next j
           If s = 0 Then
               s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限!! ", , "USER 權限問題")
               txt1(0).SetFocus
               txt1(0).SelStart = 0
               txt1(0).SelLength = Len(txt1(0))
               Exit Sub
           End If
        Next i
   Case 1
        If Len(txt1(0)) > 4 And Len(txt1(1)) > 4 Then
            s = MsgBox("當案件性質有多組時, 系統類別只能有 1 組!!")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
        If InStr(1, txt1(1), "102") <> 0 Then
            txt1(2) = "Y"
            ShowT102Again 'Add by Amy 2018/09/11
        End If
        If InStr(1, txt1(1), "702") <> 0 Then
            Text1 = "Y"
        End If
        'add by nickc 2007/03/29
        If InStr(1, txt1(1), "716") <> 0 Then
             txt1(19) = "Y"
        End If
       'Add By Cheng 2003/11/20
       'edit by nick 2004/07/20
   '     If InStr(1, txt1(1), "715") <> 0 Then
   '         txt1(16) = "Y"
   '     End If
   '     If InStr(1, txt1(1), "716") <> 0 Then
   '         txt1(17) = "Y"
   '     End If
       'End
       'Add By Cheng 2004/05/17
        If InStr(1, txt1(1), "708") <> 0 Then
            txt1(18) = "Y"
        End If
       'End
   Case 2
        Select Case txt1(2)
        Case "Y", ""
        Case Else
             s = MsgBox("是否含延展只能輸入 Y 或 空白!!", , "USER 輸入錯誤")
             txt1(2).SetFocus
             txt1(2).SelStart = 0
             txt1(2).SelLength = Len(txt1(2))
             Exit Sub
        End Select
   Case 3
   '    If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
   '       Me.txt1(Index).SetFocus
   '       txt1_GotFocus Index
   '    End If
        ShowT102Again 'Add by Amy 2018/09/11
   Case 4
   '        If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
   '           Me.txt1(Index).SetFocus
   '           txt1_GotFocus Index
   '        End If
         If Not nickChgRan(txt1(3), txt1(4), "本所期限") Then
            txt1(3).SetFocus
            txt1_GotFocus (3)
            Exit Sub
         End If
         ShowT102Again 'Add by Amy 2018/09/11
   Case 6
        If Not nickChgRan(txt1(5), txt1(6), "業務區") Then
            txt1(5).SetFocus
            txt1_GotFocus (5)
            Exit Sub
        End If
   Case 7
        If txt1(7) <> "" Then
         lbl1 = GetPrjSalesNM(txt1(7))
         If lbl1.Caption = "" Then
              s = MsgBox("智權人員錯誤！", , "錯誤！")
              txt1(7).SetFocus
              txt1_GotFocus (7)
              Exit Sub
          End If
         End If
   Case 8
        Select Case txt1(8)
        Case "1", "2", "", " "
             ShowT102Again 'Added by Lydia 2019/05/24 定稿才限制所別
        Case Else
             s = MsgBox("列印別只能輸入 1 或 2 !!", , "USER 輸入錯誤")
             txt1(8).SetFocus
             txt1(8).SelStart = 0
             txt1(8).SelLength = Len(txt1(8))
             Exit Sub
        End Select
   Case 10
        If Not nickChgRan(txt1(9), txt1(10), "申請人") Then
            txt1(9).SetFocus
            txt1_GotFocus (9)
            Exit Sub
        End If
   Case 12
        If Not nickChgRan(txt1(11), txt1(12), "代理人") Then
            txt1(11).SetFocus
            txt1_GotFocus (11)
            Exit Sub
         End If
   'Add by Amy 2018/09/11 延展再通知訊息
   Case 13
         ShowT102Again
   Case 14
         If Not nickChgRan(txt1(13), txt1(14), "申請國家") Then
             txt1(13).SetFocus
             txt1_GotFocus (13)
             Exit Sub
          End If
          ShowT102Again 'Add by Amy 2018/09/11
   Case 15 '製表日期
       If Me.txt1(Index) <> "" Then
           If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
              Me.txt1(Index).SetFocus
              txt1_GotFocus Index
           End If
       End If
   'add by nickc 2006/05/30
   Case 16
        Select Case txt1(16)
        Case "1", "2", "", " "
        Case Else
             s = MsgBox("管制表列印對象只能輸入 1 或 2 !!", , "USER 輸入錯誤")
             txt1(16).SetFocus
             txt1(16).SelStart = 0
             txt1(16).SelLength = Len(txt1(16))
             Exit Sub
        End Select
   'Add By Sindy 2020/12/14
   Case 17 '定稿日期
        If Me.txt1(Index) <> "" Then
           If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
              Me.txt1(Index).SetFocus
              txt1_GotFocus Index
           ElseIf ChkWork(ChangeTStringToWString(Me.txt1(Index))) = False Then
              Me.txt1(Index).SetFocus
              txt1_GotFocus Index
           ElseIf Val(Me.txt1(Index)) < Val(strSrvDate(2)) Then
              MsgBox "定稿日期要大於等於系統日!!!", vbExclamation + vbOKOnly
              Me.txt1(Index).SetFocus
              txt1_GotFocus Index
           End If
       End If
   'add by nickc 2007/03/29
   Case 19
        Select Case txt1(19)
        Case "Y", ""
        Case Else
             s = MsgBox("是否含第二期只能輸入 Y 或 空白!!", , "USER 輸入錯誤")
             txt1(19).SetFocus
             txt1(19).SelStart = 0
             txt1(19).SelLength = Len(txt1(19))
             Exit Sub
        End Select
   Case Else
   End Select
End Sub

''Add By Cheng 2002/11/11
'Private Sub PrintCase(strCustNo As String)
'Dim i As Integer
'Dim St As String
'Dim Page As Integer
'Dim iPrint As Integer
'Dim IntF As Integer
'Dim PriType As Integer
'Dim j As Integer
'Dim Prn As Printer
'Dim nRow As Integer
'
'On Error GoTo ErrHand
'   St = ChgCustomer(strCustNo)
'   strExc(0) = "SELECT CU04,NA03," & _
'      "CU23,CU01||CU02 FROM CUSTOMER,NATION WHERE " & St & " AND CU10=NA01(+)"
'   Page = 5
'   intI = 0
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))   'edit by nickc 2007/02/06 不用 dll 了   = objLawDll.ReadRstMsg(intI, strExc(0))
'   If intI <> 1 Then Exit Sub
'   Select Case Page
'      Case 1
'         IntF = 7
'      Case 2
'         IntF = 11
'      Case 3
'         IntF = 6
'      Case 4
'         IntF = 5
'        'Add By Cheng 2002/11/11
'        Case 5
'         IntF = 4
'   End Select
'   Printer.Font.Size = 10
'   Printer.Height = 2900
'   Printer.Width = 10000
'   RsTemp.MoveFirst
'   iPrint = 1
'   With RsTemp
'      Do While Not .EOF
'         nRow = 0
'         For i = 0 To IntF - 1
'            Printer.CurrentX = 1000
'            If IsNull(.Fields(i)) = False Then
'               If IsEmptyText(.Fields(i)) = False Then
'                  Printer.CurrentY = nRow * 220
'                  nRow = nRow + 1
'               End If
'            End If
'
'            If IsNull(.Fields(i)) = False Then
'               If IsEmptyText(.Fields(i)) = False Then
'                  Printer.Print .Fields(i)
'               End If
'            End If
'         Next
'         Printer.CurrentX = 4200
'         Printer.CurrentY = (nRow - 1) * 220
'         iPrint = iPrint + 1
'         Printer.NewPage
'         .MoveNext
'      Loop
'   End With
'   Printer.EndDoc
'   Exit Sub
'ErrHand:
'   MsgBox Err.Description
'End Sub

'Add By Cheng 2003/01/16
'取得下一程序本所期限
'Modify By Sindy 2012/6/29 +, Optional strTM01 As String = "", Optional strTM02 As String = "", Optional strTM03 As String = "", Optional strTM04 As String = ""
Private Function GetNP08(strNP01 As String, strNP07 As String, Optional strTM01 As String = "", Optional strTM02 As String = "", Optional strTM03 As String = "", Optional strTM04 As String = "") As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   GetNP08 = ""
   'Modify By Cheng 2003/03/14
   'strSQLA = "Select NP08 From NextProgress Where NP01='" & strNP01 & "' And NP07=" & Val(strNP07) & " And NP06 IS NULL "
   '92.04.03 nick add left join
   'strSQLA = "Select NP08 From NextProgress,(Select CP01, CP02, CP03, CP04 FROM CASEPROGRESS Where CP09='" & strNP01 & "' ) CP Where NP02=CP01 AND NP03=CP02 AND NP04=CP03 AND NP05=CP04 And NP07=" & Val(strNP07) & " And NP06 IS NULL "
   'Modify By Cheng 2003/05/09
   'strSQLA = "Select NP08 From NextProgress,(Select CP01, CP02, CP03, CP04 FROM CASEPROGRESS Where CP09='" & strNP01 & "' ) CP Where NP02=CP01(+) AND NP03=CP02(+) AND NP04=CP03(+) AND NP05=CP04(+) And NP07=" & Val(strNP07) & " And NP06 IS NULL "
   'Modify By Sindy 2012/6/29 TF的使用宣誓會抓不到np期限ex.C97062671/105
   If strTM01 = "TF" Then
      StrSQLa = "Select NP08 From NextProgress Where NP01='" & strNP01 & "' AND NP02='" & strTM01 & "' AND NP03=" & strTM02 & " AND NP04=" & strTM03 & " AND NP05=" & strTM04 & " And NP07=" & Val(strNP07) & " And NP06 IS NULL "
   Else
   '2012/6/29 End
      StrSQLa = "Select NP08 From NextProgress,(Select CP01, CP02, CP03, CP04 FROM CASEPROGRESS Where CP09='" & strNP01 & "' ) CP Where NP02=CP01 AND NP03=CP02 AND NP04=CP03 AND NP05=CP04 And NP07=" & Val(strNP07) & " And NP06 IS NULL "
   End If
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       GetNP08 = "" & rsA.Fields(0).Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

'Add By Cheng 2003/01/16
'取得下一程序法定期限
'Modify By Sindy 2012/6/29 +, Optional strTM01 As String = "", Optional strTM02 As String = "", Optional strTM03 As String = "", Optional strTM04 As String = ""
Private Function GetNP09(strNP01 As String, strNP07 As String, Optional strTM01 As String = "", Optional strTM02 As String = "", Optional strTM03 As String = "", Optional strTM04 As String = "") As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   GetNP09 = ""
   'Modify By Cheng 2003/03/14
   'strSQLA = "Select NP09 From NextProgress Where NP01='" & strNP01 & "' And NP07=" & Val(strNP07) & " And NP06 IS NULL "
   '92.04.03 nick add left join
   'strSQLA = "Select NP09 From NextProgress,(Select CP01, CP02, CP03, CP04 FROM CASEPROGRESS Where CP09='" & strNP01 & "' ) CP Where NP02=CP01 AND NP03=CP02 AND NP04=CP03 AND NP05=CP04 And NP07=" & Val(strNP07) & " And NP06 IS NULL "
   'Modify By Cheng 2003/05/09
   'strSQLA = "Select NP09 From NextProgress,(Select CP01, CP02, CP03, CP04 FROM CASEPROGRESS Where CP09='" & strNP01 & "' ) CP Where NP02=CP01(+) AND NP03=CP02(+) AND NP04=CP03(+) AND NP05=CP04(+) And NP07=" & Val(strNP07) & " And NP06 IS NULL "
   'Modify By Sindy 2012/6/29 TF的使用宣誓會抓不到np期限ex.C97062671/105
   If strTM01 = "TF" Then
      StrSQLa = "Select NP09 From NextProgress Where NP01='" & strNP01 & "' AND NP02='" & strTM01 & "' AND NP03=" & strTM02 & " AND NP04=" & strTM03 & " AND NP05=" & strTM04 & " And NP07=" & Val(strNP07) & " And NP06 IS NULL "
   Else
   '2012/6/29 End
      StrSQLa = "Select NP09 From NextProgress,(Select CP01, CP02, CP03, CP04 FROM CASEPROGRESS Where CP09='" & strNP01 & "' ) CP Where NP02=CP01 AND NP03=CP02 AND NP04=CP03 AND NP05=CP04 And NP07=" & Val(strNP07) & " And NP06 IS NULL "
   End If
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       GetNP09 = "" & rsA.Fields(0).Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

'Add by Amy 2018/08/31
Private Sub txt1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
     If Index = 3 Or Index = 4 Or Index = 21 Or Index = 22 Then
        If Index = 3 Or Index = 4 Then
            Option1(0).Value = True
        Else
            Option1(1).Value = True
        End If
    End If
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
    'Add By Cheng 2003/05/13
    Select Case Index
    Case 3 '本所期限起
        If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Cancel = True
        End If
    Case 4 '本所期限迄
        If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Cancel = True
        End If
    End Select
End Sub

'Add By Sindy 2010/6/11
Private Sub textMoney_GotFocus()
   InverseTextBox textMoney
End Sub
' 美國使用宣誓費用
Private Sub textMoney_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(textMoney) = False Then
      If IsNumeric(textMoney) = False Then
         Cancel = True
         strTit = "資料檢核"
         'strMsg = "請輸入正確的美國使用宣誓費用!!!"
         strMsg = "請輸入正確的報價費用!!!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textMoney_GotFocus
      End If
   End If
End Sub

'Add by Amy 2018/08/31
Private Function FormCheck() As Boolean
    FormCheck = False
    If Len(txt1(0)) = 0 Then
        s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        Exit Function
    End If
    
    'Add By Sindy 2025/4/30
    If textTM01 <> "" And textTM02 <> "" Then
      If textTM02_2.Visible = True And textTM02_2.Text = "" Then textTM02_2.Text = "0"
      If textTM03.Text = "" Then textTM03.Text = "0"
      If textTM04.Text = "" Then textTM04.Text = "00"
      If txt1(8) = "1" Then
         s = MsgBox("輸入本所案號時,列印別必須為 2 定稿!!", , "USER 輸入錯誤")
         txt1(8).SetFocus
         Exit Function
      End If
      If Option1(1).Value = True Then
         s = MsgBox("輸入本所案號時,不可選 大陸撤三期限!!", , "USER 輸入錯誤")
         Exit Function
      End If
      '抓下一程序的本所期限
      'Modify By Sindy 2025/7/14
      If Trim(txt1(1)) <> "" Then
      '2025/7/14 END
         If textTM01 = "TF" Then
            strSql = "select * from nextprogress" & _
                     " where NP02='" & textTM01 & "' and NP03='" & textTM02 & textTM02_2 & "' and NP04='" & textTM03 & "' and NP05='" & textTM04 & "'" & _
                     " and NP07 in(" & txt1(1) & ") and (NP06 IS NULL OR NP06='')"
         Else
            strSql = "select * from nextprogress" & _
                     " where NP02='" & textTM01 & "' and NP03='" & textTM02 & "' and NP04='" & textTM03 & "' and NP05='" & textTM04 & "'" & _
                     " and NP07 in(" & txt1(1) & ") and (NP06 IS NULL OR NP06='')"
         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If Val("" & RsTemp.Fields("NP08")) > 0 Then
               txt1(3) = TransDate(RsTemp.Fields("NP08"), 1)
               txt1(4) = txt1(3)
            End If
         End If
      End If
    End If
    '2025/4/30 END
    
    'Add By Cheng 2002/10/01
    strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
    strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
    For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp2(i) = strTemp1(j) Then
                s = 1
                Exit For
            End If
        Next j
        'Add By Sindy 2010/6/11
        'Modify By Sindy 2021/3/18 TF續展的報價定稿取消,改為單純續展期限通知函 拿掉Or TXT1(1) = "102"
        If strTemp2(i) = "TF" And txt1(8) = "2" And _
            txt1(1) = "105" And _
            Val(textMoney) = 0 Then
            s = MsgBox("報價費用不可空白!!", , "USER 輸入錯誤")
            textMoney.SetFocus
            textMoney_GotFocus
            Exit Function
        End If
        '2010/6/11 End
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限!! ", , "USER 權限問題")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Function
        End If
    Next i
    'Add By Cheng 2002/03/21
    'Add by Amy 2018/11/02
    If Option1(1).Value = True Then
        If txt1(0) <> "T" Then
            MsgBox "選大陸撤三,系統別只能輸T"
            txt1(0).SetFocus
            txt1_GotFocus 0
            Exit Function
        End If
        If PUB_CheckKeyInDate(Me.txt1(21)) = -1 Then
            Me.txt1(21).SetFocus
            txt1_GotFocus 21
            Exit Function
        End If
        If PUB_CheckKeyInDate(Me.txt1(22)) = -1 Then
            Me.txt1(22).SetFocus
            txt1_GotFocus 22
            Exit Function
        End If
        If Len(txt1(21)) = 0 Then
            s = MsgBox("大陸撤三期限區間不可空白!!", , "USER 輸入錯誤")
            txt1(21).SetFocus
            txt1_GotFocus 21
            Exit Function
        End If
        If Len(txt1(22)) = 0 Then
            s = MsgBox("大陸撤三期限間不可空白!!", , "USER 輸入錯誤")
            txt1(22).SetFocus
            txt1_GotFocus 22
            Exit Function
        End If
    End If
    'end 2018/11/02
    'Modify by Amy 2018/08/31
     If Option1(0).Value = True Then
        If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
            Me.txt1(3).SetFocus
            txt1_GotFocus 3
            Exit Function
        End If
        If PUB_CheckKeyInDate(Me.txt1(4)) = -1 Then
            Me.txt1(4).SetFocus
            txt1_GotFocus 4
            Exit Function
        End If
        If Len(txt1(4)) = 0 Then
            s = MsgBox("本所期限區間不可空白!!", , "USER 輸入錯誤")
            txt1(3).SetFocus
            txt1_GotFocus (3)
            Exit Function
        End If
    End If
    If Len(txt1(8)) = 0 Then
        s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
        txt1(8).SetFocus
        Exit Function
    End If
    
    'Add By Sindy 2020/12/14
    If txt1(8) = "2" Then '2.定稿
      If Me.txt1(17).Text = "" Then
         MsgBox "請輸入定稿日期!!!", vbExclamation + vbOKOnly
         Me.txt1(17).SetFocus
         txt1_GotFocus 17
         Exit Function
      Else
         'Add By Sindy 2022/8/5 定稿日期不可超過系統日2個月
         If DBDATE(Me.txt1(17).Text) > DBDATE(DateAdd("m", 2, ChangeWStringToWDateString(strSrvDate(1)))) Then
            MsgBox "定稿日期不可超過系統日2個月！", vbInformation, "輸入日期錯誤"
            Me.txt1(17).SetFocus
            txt1_GotFocus 17
            Exit Function
         End If
         '2022/8/5 End
         
         'Add By Sindy 2022/6/6
         If MsgBox("確定此定稿日期（" & Me.txt1(17).Text & "）是對的嗎？" & vbCrLf & vbCrLf & _
                   "注意：它是顯示在定稿內容裡的發文日期。", vbYesNo + vbInformation + vbDefaultButton2) = vbNo Then
            Me.txt1(17).SetFocus
            txt1_GotFocus 17
            Exit Function
         End If
         '2022/6/6 END
'      ElseIf PUB_CheckKeyInDate(Me.txt1(17)) = -1 Then
'          Me.txt1(17).SetFocus
'          txt1_GotFocus 17
'          Exit Function
'      ElseIf ChkWork(ChangeTStringToWString(Me.txt1(17).Text)) = False Then
'          Me.txt1(17).SetFocus
'          txt1_GotFocus 17
'          Exit Function
'      ElseIf Val(Me.txt1(17).Text) < Val(strSrvDate(2)) Then
'          MsgBox "定稿日期要大於等於系統日!!!", vbExclamation + vbOKOnly
'          Me.txt1(17).SetFocus
'          txt1_GotFocus 17
'          Exit Function
      End If
    End If
    
    'Add By Cheng 2002/11/21 列印別為管制表
    'add by nickc 2006/05/30 管制表列印對象不可空白
    If Me.txt1(8).Text = "1" Then
        If Me.txt1(16).Text = "" Then
            s = MsgBox("管制表列印對象不可空白!!", , "USER 輸入錯誤")
            txt1(16).SetFocus
            Exit Function
        End If
        If Mid(txt1(9), 1, 6) <> Mid(txt1(10), 1, 6) Then
            s = MsgBox("申請人代號前六碼必須相同!!", , "USER 輸入錯誤")
            txt1(9).SetFocus
            txt1(9).SelStart = 0
            txt1(9).SelLength = Len(txt1(9))
            Exit Function
        End If
        If Mid(txt1(11), 1, 6) <> Mid(txt1(12), 1, 6) Then
            s = MsgBox("代理人代號前六碼必須相同!!", , "USER 輸入錯誤")
            txt1(11).SetFocus
            txt1(11).SelStart = 0
            txt1(11).SelLength = Len(txt1(11))
            Exit Function
        End If
    End If
    DoEvents
    'Add By Cheng  2002/11/29
    '若要列印延展, 刊登廣告, 第一期註冊費, 第二期註冊費, 繳年費
    'If Me.txt1(2).Text = "Y" Or Me.Text1.Text = "Y" Or Me.txt1(16).Text = "Y" Or Me.txt1(17).Text = "Y" Or Me.txt1(18).Text = "Y" Then
    'edit by nickc 2007/03/29
    'If Me.txt1(2).Text = "Y" Or Me.Text1.Text = "Y" Or Me.txt1(18).Text = "Y" Then
    If Me.txt1(2).Text = "Y" Or Me.Text1.Text = "Y" Or Me.txt1(18).Text = "Y" Or txt1(19) = "Y" Then
        '若列印別為管制表
        If txt1(8).Text = "1" Then
            If Me.txt1(15).Text = "" Then
                MsgBox "請輸入製表日期!!!", vbExclamation + vbOKOnly
                Me.txt1(15).SetFocus
                txt1_GotFocus 15
                Exit Function
            ElseIf PUB_CheckKeyInDate(Me.txt1(15)) = -1 Then
                Me.txt1(15).SetFocus
                txt1_GotFocus 15
                Exit Function
            End If
        End If
    End If
    '定稿
    If txt1(8) = "2" Then
        'Added by Lydia 2018/05/28
        If InStr(UCase(txt1(0)), "FCT") > 0 Then
            MsgBox "FCT沒有定稿，請排除 !", vbExclamation
            txt1_GotFocus 0
            Exit Function
        End If
        'end 2018/05/28
        'Add By Sindy 2015/4/17 台灣案催延展對象
        If Trim(txt1(13)) = "000" And Trim(txt1(14)) = "000" Then
            '系統別有T者
            strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
            s = 0
            For i = 0 To UBound(strTemp2)
                If strTemp2(i) = "T" Then
                    s = 1
                    Exit For
                End If
            Next i
            If s = 1 And txt1(2) = "Y" Then 'T含延展
                If Len(txt1(20)) = 0 Then
                    MsgBox "請輸入台灣案催延展對象!!", vbExclamation + vbOKOnly
                    Me.txt1(20).SetFocus
                    txt1_GotFocus 20
                    Exit Function
                End If
            End If
        End If
        '2015/4/17 END
        
        'Added by Lydia 2019/05/20 T案催延展必須輸入國別
        If InStr("T,", Trim(txt1(0))) > 0 And txt1(1) = "102" And txt1(2) = "Y" Then
            If txt1(13) = txt1(14) And InStr("000,020", txt1(13)) > 0 And Trim(txt1(13)) <> "" And InStr("000,020", txt1(14)) > 0 And Trim(txt1(14)) <> "" Then
            Else
                MsgBox "申請國家請輸入000或020!!", vbExclamation + vbOKOnly
                Me.txt1(13).SetFocus
                txt1_GotFocus 13
                Exit Function
            End If
            
            'Added by Morgan 2024/6/19
            '檢查催延展/續展期限
            If Check102DateRange = False Then
               txt1(3).SetFocus
               txt1_GotFocus 3
               Exit Function
            End If
            'end 2024/6/19
            
        End If
        'end 2019/05/20
    End If
    
    FormCheck = True
End Function

'Add by Amy 2018/08/31 備註
Sub PrintMemo()
    iPrint = iPrint + 300
    Printer.CurrentX = PLeft(1)
    Printer.CurrentY = iPrint
    Printer.Print "PS：本所案號前有*者表示已通知撤三開拓"
End Sub

'Add by Amy 2018/09/11 延展再通知訊息
Sub ShowT102Again()
    lblMsg.Caption = ""
    'Modify by Amy 2019/06/11 +TXT1(2) = "1" 排除大->台
    '(條件:T/案性:102/本所期限:1090401-1090430/智權:96030(巨京)/列印別:2.定稿/管制表列印對象:1/申請國家:000-000/台灣案催延展對象2.大->台) T-165338 會出不來
    'Modify by Amy 2024/06/20 2019/06/11 排除大->台 欄位應為TXT1(20)-bug
    If txt1(8) = "2" And txt1(20) = "1" Then 'Added by Lydia 2019/05/24 定稿才限制所別
        If txt1(2) = "Y" And Len(txt1(3)) <> 0 Then
            If txt1(13) = "000" And Val(txt1(4)) + 19110000 < Val(GetPreMonLastDate(DBDATE(DateAdd("yyyy", 1, Format(strSrvDate(1), "####/##/##"))), False)) Then
                lblMsg.Caption = "南、高所不發"
            End If
            If txt1(13) = "020" Then
                If Val(txt1(3)) + 19110000 < Val(GetPreMonLastDate(DBDATE(DateAdd("m", 6, Format(strSrvDate(1), "####/##/##"))), False)) Then
                    lblMsg.Caption = "高所不發"
                ElseIf Val(txt1(3)) + 19110000 < Val(GetPreMonLastDate(DBDATE(DateAdd("m", 18, Format(strSrvDate(1), "####/##/##"))), False)) Then
                    lblMsg.Caption = "南、高所不發"
                End If
            End If
        End If
    End If
End Sub

'Added by Morgan 2024/6/19
'檢查催延展/續展期限
Private Function Check102DateRange() As Boolean
   Check102DateRange = True
   If txt1(8) = "2" And txt1(2) = "Y" Then
      '台灣案期限超過一年
      If txt1(13) = "000" Then
         If Left(DBDATE(txt1(4)), 6) > Left(CompDate(0, 1, DBDATE(txt1(17))), 6) Then
            'Modify By Sindy 2025/4/21 調整訊息原:繼續跑管制表 改為:繼續執行
            If MsgBox("台灣案延展期限超過一年，確定是否要繼續執行？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
               Check102DateRange = False
            End If
         End If
      '大陸案期限超過一年半
      ElseIf txt1(13) = "020" Then
         'Modified by Morgan 2024/6/26 實際上是催第19個月 Ex: 113/8/1 -> 115/3/1~31 --湘芸/桂英
         'If Left(DBDATE(TXT1(4)), 6) > Left(CompDate(1, 18, DBDATE(TXT1(17))), 6) Then
         If Left(DBDATE(txt1(4)), 6) > Left(CompDate(1, 19, DBDATE(txt1(17))), 6) Then
         'end 2024/6/26
            'Modify By Sindy 2025/4/21 調整訊息原:繼續跑管制表 改為:繼續執行
            If MsgBox("大陸案續展期限超過一年半，確定是否要繼續執行？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
               Check102DateRange = False
            End If
         End If
      End If
   End If
End Function
'end 2024/6/19

'Added by Lydia 2019/08/22 計算雙面列印定稿的可印範圍
'因為原本單面印整批定稿約半小時,改成雙面列印要1小時以上;
'改成先計算商品名稱的可印範圍存入例外欄位記錄,減少印整批定稿的時間
'Memo by Lydia 2020/01/09 Word2013預設.doc檔的第2頁版面，最大行數43；但是Word2013預設.docx檔，在插入代表圖後最大行數小於43，改用變數控制
Private Sub SetTMGoodsDetail(ByVal pTM01 As String, ByVal pTM02 As String, ByVal pTM03 As String, ByVal pTM04 As String, _
                                             ByVal pCP09 As String, ByVal pET01 As String, ByVal pET03 As String)
Dim m_line As Variant
Dim intL As Integer, intQ As Integer, strTemp As String
Dim intA As Integer, intB As Integer, intC As String
Dim strMid As String
Dim r

     'T延展通知之商品名稱調整: 若商品類別名稱超過26個字,其他印到第2頁
     
     '商品名稱第1頁可印完 +  只有一個類別 => 直接顯示
     strExc(0) = "select tg05,tg06,tg07,tg08,tg15,tg16,tg17 from tmgoods where tg01='" & pTM01 & "' and tg02='" & pTM02 & "' and tg03='" & pTM03 & "' and tg04='" & pTM04 & "' and tg18 is null order by tg05 "
     intI = 1
     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
     strMid = ""
     '有商品名稱
     If RsTemp.RecordCount > 1 Then
        strMid = "(詳見本通知函背面)"
     ElseIf RsTemp.RecordCount = 1 Then
        'Modified by Lydia 2023/02/14 +trim()
        If GetTextLength(Trim("" & RsTemp.Fields("tg06") & RsTemp.Fields("tg15"))) > 26 Then
            strMid = "(詳見本通知函背面)"
        Else '只有一個類別,並且不超過可顯示範圍
            strMid = CheckStr("" & RsTemp.Fields("tg06") & RsTemp.Fields("tg15"))
        End If
    '沒有商品名稱
    Else
             strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                      "VALUES ('" & pET01 & "','" & pCP09 & "','" & pET03 & "','" & strUserNum & "'," & _
                      "'TMGoods第2頁',' ')"
             cnnConnection.Execute strSql
     End If
     'TMGoods第2頁
     If strMid <> "" Then
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & pET01 & "','" & pCP09 & "','" & pET03 & "','" & strUserNum & "'," & _
                  "'TMGoods第1頁','" & strMid & "')"
         cnnConnection.Execute strSql
         strMid = ""
         
         strExc(1) = "": strExc(2) = "": strExc(3) = ""
         If m_bWord = True Then '將文字丟到Word
            intI = RsTemp.RecordCount
            RsTemp.MoveFirst
            strExc(2) = "第" & RsTemp.Fields("tg05") & "類："
            Do While Not RsTemp.EOF
                'Modified by Lydia 2023/02/14 +trim()
                strExc(1) = strExc(1) & IIf(strExc(1) <> "", "第" & RsTemp.Fields("tg05") & "類：", "") & CheckStr(Trim("" & RsTemp.Fields("tg06") & RsTemp.Fields("tg15"))) & vbCrLf
                RsTemp.MoveNext
            Loop
             'Modified by Lydia 2023/02/14 去掉 vbCrLf ; 雅雯反應2/1開始第2頁都會印商品類別名稱
            'If Not (intI = 1 And GetTextLength(strExc(1)) <= 26) Then
            If Not (intI = 1 And GetTextLength(Replace(strExc(1), vbCrLf, "")) <= 26) Then
                  '將文字丟到Word
                  strExc(1) = "商品名稱：" & vbCrLf & strExc(2) & strExc(1)
                 g_WordAp.Selection.TypeText Text:=strExc(1)
                 '抓總行數
                 r = g_WordAp.ActiveDocument.BuiltInDocumentProperties(wdPropertyLines)
                 'Modified by  Lydia 2020/01/09 改用變數控制
                 'If r <= 41 Then '未超過第二頁
                 If r <= TMGoods第2頁行數 - 1 Then '未超過第二頁
                     strMid = strExc(1)
                 Else  '超過第二頁,用直接計算
                     GoTo JumpToCount
                 End If
            End If
            
         Else
'------------------------直接計算---------------------------
            If RsTemp.RecordCount >= 1 Then
               intI = RsTemp.RecordCount
               RsTemp.MoveFirst
               strExc(2) = "第" & RsTemp.Fields("tg05") & "類："
               Do While Not RsTemp.EOF
                   strExc(1) = strExc(1) & IIf(strExc(1) <> "", "第" & RsTemp.Fields("tg05") & "類：", "") & CheckStr("" & RsTemp.Fields("tg06") & RsTemp.Fields("tg15")) & vbCrLf
                   RsTemp.MoveNext
               Loop
               'Modified by Lydia 2023/02/14 去掉 vbCrLf ; 雅雯反應2/1開始第2頁都會印商品類別名稱
               'If Not (intI = 1 And GetTextLength(strExc(1)) <= 26) Then '商品名稱第1頁可印完 只有一個類別 >= 直接顯示
               If Not (intI = 1 And GetTextLength(Replace(strExc(1), vbCrLf, "")) <= 26) Then
                  '計算第2頁版面：高40~42行，每一行最多34中文字
                  'Memo by Lydia 2020/01/09 Word2013預設.doc檔的第2頁版面，最大行數43(含未完全列出的備註2行)；
                  '但是Word2013預設.docx檔，在插入代表圖後最大行數小於43，改用變數控制
                  strExc(1) = "商品名稱：" & vbCrLf & strExc(2) & strExc(1)
JumpToCount:
                  m_line = Empty
                  m_line = Split(strExc(1), vbCrLf)
                  intL = 0 '行數
                  For intQ = 0 To UBound(m_line)
                      If Trim(m_line(intQ)) <> "" Then
                          strTemp = m_line(intQ)
                          intI = GetTextLength(strTemp)
                          For i = 0 To intI \ 68
                               strExc(4) = Trim(convForm(strTemp, 68))
                               strExc(3) = strExc(3) & strExc(4)
                               strTemp = MidB(strTemp, LenB(strExc(4)) + 1)
                               If i > 0 Then intL = intL + 1
                               'Modified by Lydia 2020/01/09 改用變數控制
                               'If intL = 40 Then
                               If intL = TMGoods第2頁行數 - 2 Then
                                   '省略號的位置
                                   If GetTextLength(strExc(4)) > 62 Then
                                       intA = InStrRev(strExc(3), "、")
                                       intB = InStrRev(strExc(3), "，")
                                       intC = InStrRev(strExc(3), "；")
                                       If intA + intB + intC > 0 Then
                                           If intA > intB Then
                                              intI = intA
                                           Else
                                              intI = intB
                                           End If
                                           If intI < intC Then
                                              intI = intC
                                           End If
                                           strExc(3) = Mid(strExc(3), 1, intI - 1) & "~~~" & Mid(strExc(3), intI)
                                       Else
                                           strExc(3) = Mid(strExc(3), 1, Len(strExc(3)) - 3) & "~~~" & Mid(strExc(3), Len(strExc(3)) - 3)
                                       End If
                                   Else
                                       If InStr(strExc(4), "。") = 0 Then
                                          strExc(3) = strExc(3) & "~~~"
                                       End If
                                   End If
                                   strExc(3) = strExc(3) & "|||" '第2頁最末行的位置
                               End If
                          Next i
                          If intI \ 68 > 0 And Trim(strTemp) <> "" Then '剩下的字串
                              strExc(3) = strExc(3) & strTemp
                              intL = intL + 1
                              'Modified by Lydia 2020/01/09 改用變數控制
                              'If intL = 40 Then strExc(3) = strExc(3) & "~~~|||" '省略號+第2頁最末行的位置
                              If intL = TMGoods第2頁行數 - 2 Then strExc(3) = strExc(3) & "~~~|||"
                          End If
                          intL = intL + 1
                          strExc(3) = strExc(3) & vbCrLf
                      End If
                  Next intQ
                  'Modified by Lydia 2020/01/09 改用變數控制
                  'If intL <= 42 Then
                  If intL <= TMGoods第2頁行數 Then
                     strExc(3) = Replace(strExc(3), "~~~", "")
                     strExc(3) = Replace(strExc(3), "|||", "")
                  Else '商品名稱第2頁無法印完
                     If InStr(strExc(3), "~~~") > 0 Then
                         strExc(3) = Mid(strExc(3), 1, InStr(strExc(3), "~~~") - 1)
                         strExc(3) = strExc(3) & "‧‧‧"
                     Else
                          strExc(3) = Mid(strExc(3), 1, InStr(strExc(3), "|||") - 1)
                     End If
                     strExc(3) = strExc(3) & vbCrLf & "（限於篇幅，本案之商品名稱未完全列出，欲瞭解完整之商品名稱，請與本所智權人員聯繫）"
                  End If
                  strMid = strExc(3)
               End If
            End If
'---------------------------------------------------
         End If
         
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & pET01 & "','" & pCP09 & "','" & pET03 & "','" & strUserNum & "'," & _
                  "'TMGoods第2頁','" & strMid & "')"
         cnnConnection.Execute strSql
         
         If m_bWord = True Then
            g_WordAp.Selection.WholeStory
            g_WordAp.Selection.Text = "" '清空文字
         End If
    End If
End Sub

'Add By Sindy 2025/4/29
Private Sub textTM01_GotFocus()
   InverseTextBox textTM01
   CloseIme
End Sub
Private Sub textTM02_GotFocus()
   InverseTextBox textTM02
   CloseIme
End Sub
Private Sub textTM02_2_GotFocus()
   InverseTextBox textTM02_2
   CloseIme
End Sub
Private Sub textTM03_GotFocus()
   InverseTextBox textTM03
   CloseIme
End Sub
Private Sub textTM03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textTM04_GotFocus()
   InverseTextBox textTM04
   CloseIme
End Sub
Private Sub textTM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
' 本所案號的系統別
Private Sub textTM01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM01) = False Then
      If InStr(txt1(0), textTM01) = 0 Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "本所案號中的系統別必須存在於畫面上的系統類別欄位中!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
      Select Case textTM01
         Case "TF":
            textTM02_2.Visible = True
            textTM02_2.Locked = False
            textTM02_2.TabStop = True
            textTM02.MaxLength = 5
         Case Else
            textTM02_2.Visible = False
            textTM02_2.Locked = True
            textTM02_2.TabStop = False
            textTM02.MaxLength = 6
      End Select
      If Len(textTM02) <> textTM02.MaxLength Then textTM02 = ""
   Else
      textTM02_2.Visible = False
      textTM02_2.Locked = True
      textTM02_2.TabStop = False
      textTM02.MaxLength = 6
   End If
   
EXITSUB:
End Sub
'2025/4/29 END
