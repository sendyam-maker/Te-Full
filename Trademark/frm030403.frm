VERSION 5.00
Begin VB.Form frm030403 
   BorderStyle     =   1  '單線固定
   Caption         =   "期限管制表"
   ClientHeight    =   6470
   ClientLeft      =   7700
   ClientTop       =   2970
   ClientWidth     =   8230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6470
   ScaleWidth      =   8230
   Begin VB.Frame Frame102 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Height          =   4990
      Left            =   5520
      TabIndex        =   63
      Top             =   780
      Width           =   2620
      Begin VB.ListBox List1 
         Height          =   4360
         Index           =   1
         ItemData        =   "frm030403.frx":0000
         Left            =   0
         List            =   "frm030403.frx":0002
         TabIndex        =   32
         Top             =   300
         Width           =   1940
      End
      Begin VB.CommandButton Command2 
         Caption         =   "刪除"
         Height          =   400
         Index           =   1
         Left            =   1980
         TabIndex        =   31
         Top             =   480
         Width           =   600
      End
      Begin VB.CommandButton Command2 
         Caption         =   "新增"
         Default         =   -1  'True
         Height          =   400
         Index           =   0
         Left            =   1980
         TabIndex        =   30
         Top             =   30
         Width           =   600
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   0
         MaxLength       =   3
         TabIndex        =   26
         Text            =   "FCT"
         Top             =   30
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Index           =   2
         Left            =   510
         MaxLength       =   6
         TabIndex        =   27
         Top             =   30
         Width           =   800
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Index           =   3
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   28
         Top             =   30
         Width           =   255
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Index           =   4
         Left            =   1590
         MaxLength       =   2
         TabIndex        =   29
         Top             =   30
         Width           =   350
      End
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   7
      Left            =   2000
      MaxLength       =   1
      TabIndex        =   14
      Top             =   2910
      Width           =   315
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   19
      Left            =   2570
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1320
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   18
      Left            =   1430
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1320
      Width           =   990
   End
   Begin VB.OptionButton Option1 
      Caption         =   "可辦期限："
      Height          =   180
      Index           =   3
      Left            =   210
      TabIndex        =   56
      Top             =   1350
      Width           =   1200
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   17
      Left            =   3030
      MaxLength       =   1
      TabIndex        =   15
      Text            =   "1"
      Top             =   3180
      Width           =   285
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   16
      Left            =   2660
      MaxLength       =   4
      TabIndex        =   24
      Top             =   4920
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   15
      Left            =   1500
      MaxLength       =   4
      TabIndex        =   23
      Top             =   4920
      Width           =   990
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Index           =   0
      Left            =   1430
      MaxLength       =   1
      TabIndex        =   12
      Top             =   2250
      Width           =   315
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定地址條"
      Height          =   570
      Left            =   240
      TabIndex        =   50
      Top             =   5220
      Width           =   3800
      Begin VB.ComboBox Combo1 
         Height          =   260
         Left            =   765
         TabIndex        =   25
         Text            =   "Combo1"
         Top             =   240
         Width           =   2940
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   260
         Index           =   1
         Left            =   110
         TabIndex        =   51
         Top             =   290
         Width           =   770
      End
   End
   Begin VB.TextBox text1 
      Height          =   285
      Index           =   3
      Left            =   3210
      MaxLength       =   2
      TabIndex        =   10
      Top             =   1620
      Width           =   555
   End
   Begin VB.TextBox text1 
      Height          =   285
      Index           =   2
      Left            =   2910
      MaxLength       =   1
      TabIndex        =   9
      Top             =   1620
      Width           =   255
   End
   Begin VB.TextBox text1 
      Height          =   285
      Index           =   1
      Left            =   1980
      MaxLength       =   6
      TabIndex        =   8
      Top             =   1620
      Width           =   885
   End
   Begin VB.TextBox text1 
      Height          =   285
      Index           =   0
      Left            =   1430
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1620
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   2
      Left            =   210
      TabIndex        =   49
      Top             =   1680
      Width           =   1200
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   14
      Left            =   1500
      TabIndex        =   16
      Top             =   3420
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   9
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   18
      Top             =   4010
      Width           =   1005
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所期限："
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   44
      Top             =   840
      Value           =   -1  'True
      Width           =   1200
   End
   Begin VB.OptionButton Option1 
      Caption         =   "法定期限："
      Height          =   180
      Index           =   1
      Left            =   210
      TabIndex        =   43
      Top             =   1080
      Width           =   1200
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   1
      Left            =   1430
      MaxLength       =   7
      TabIndex        =   1
      Top             =   780
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   2
      Left            =   2570
      MaxLength       =   7
      TabIndex        =   2
      Top             =   780
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   3
      Left            =   1430
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1050
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   4
      Left            =   2570
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1050
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   0
      Left            =   1430
      TabIndex        =   0
      Top             =   495
      Width           =   2340
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   5
      Left            =   1430
      TabIndex        =   11
      Top             =   1965
      Width           =   675
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   6
      Left            =   2000
      MaxLength       =   1
      TabIndex        =   13
      Top             =   2490
      Width           =   315
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   8
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   17
      Top             =   3720
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   10
      Left            =   1500
      MaxLength       =   9
      TabIndex        =   19
      Top             =   4310
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   11
      Left            =   2660
      MaxLength       =   9
      TabIndex        =   20
      Top             =   4310
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   12
      Left            =   1500
      MaxLength       =   9
      TabIndex        =   21
      Top             =   4610
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   13
      Left            =   2660
      MaxLength       =   9
      TabIndex        =   22
      Top             =   4610
      Width           =   990
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   4380
      TabIndex        =   33
      Top             =   80
      Width           =   756
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   5220
      TabIndex        =   34
      Top             =   80
      Width           =   756
   End
   Begin VB.Label Lbl_102 
      AutoSize        =   -1  'True
      Caption         =   "不需定稿通知的案號："
      Height          =   180
      Left            =   3720
      TabIndex        =   62
      Top             =   840
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "4.未催延展檢核表)"
      Height          =   180
      Index           =   16
      Left            =   3780
      TabIndex        =   61
      Top             =   2730
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "報表類別 (1、4)組別："
      Height          =   180
      Index           =   6
      Left            =   210
      TabIndex        =   60
      Top             =   2940
      Width           =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "FCT催延展請務必填案件性質102"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Index           =   15
      Left            =   2460
      TabIndex        =   59
      Top             =   1980
      Width           =   1880
   End
   Begin VB.Label LblNote2 
      ForeColor       =   &H000000C0&
      Height          =   260
      Left            =   120
      TabIndex        =   58
      Top             =   6120
      Width           =   8000
   End
   Begin VB.Label LblNote 
      ForeColor       =   &H000000C0&
      Height          =   260
      Left            =   120
      TabIndex        =   57
      Top             =   5820
      Width           =   8000
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1890
      X2              =   3090
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CFT、CFC 未收文管制表列印對象："
      Height          =   180
      Index           =   9
      Left            =   210
      TabIndex        =   55
      Top             =   3210
      Width           =   2850
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(1.非智權部同仁  2.全部)"
      Height          =   180
      Index           =   14
      Left            =   3360
      TabIndex        =   54
      Top             =   3210
      Width           =   1920
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   13
      Left            =   300
      TabIndex        =   53
      Top             =   4950
      Width           =   920
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2070
      X2              =   3270
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否含延展：           (Y：含)"
      Height          =   180
      Index           =   12
      Left            =   210
      TabIndex        =   52
      Top             =   2270
      Width           =   2180
   End
   Begin VB.Line Line3 
      X1              =   1800
      X2              =   3330
      Y1              =   1770
      Y2              =   1770
   End
   Begin VB.Label Label1 
      Caption         =   "費用(報價)："
      Height          =   180
      Index           =   4
      Left            =   300
      TabIndex        =   48
      Top             =   3440
      Width           =   1160
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   3
      Left            =   300
      TabIndex        =   47
      Top             =   4050
      Width           =   950
   End
   Begin VB.Label LBL1 
      Height          =   180
      Index           =   1
      Left            =   2550
      TabIndex        =   46
      Top             =   4050
      Width           =   1130
   End
   Begin VB.Label Label1 
      Caption         =   "(1.英文組 2.日文組)"
      Height          =   180
      Index           =   10
      Left            =   2370
      TabIndex        =   45
      Top             =   2940
      Width           =   1620
   End
   Begin VB.Line Line2 
      X1              =   2190
      X2              =   2805
      Y1              =   920
      Y2              =   920
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   1890
      X2              =   3090
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   42
      Top             =   540
      Width           =   950
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   1
      Left            =   210
      TabIndex        =   41
      Top             =   2010
      Width           =   950
   End
   Begin VB.Label Label1 
      Caption         =   "延展註冊費報表類別："
      Height          =   180
      Index           =   2
      Left            =   210
      TabIndex        =   40
      Top             =   2520
      Width           =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   180
      Index           =   5
      Left            =   300
      TabIndex        =   39
      Top             =   3750
      Width           =   950
   End
   Begin VB.Label Label1 
      Caption         =   "申請人："
      Height          =   180
      Index           =   7
      Left            =   300
      TabIndex        =   38
      Top             =   4350
      Width           =   950
   End
   Begin VB.Label Label1 
      Caption         =   "代理人："
      Height          =   180
      Index           =   8
      Left            =   300
      TabIndex        =   37
      Top             =   4640
      Width           =   950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(1.管制表  2.定稿  3.傳真定稿"
      Height          =   180
      Index           =   11
      Left            =   2370
      TabIndex        =   36
      Top             =   2520
      Width           =   2240
   End
   Begin VB.Label LBL1 
      Height          =   180
      Index           =   0
      Left            =   2550
      TabIndex        =   35
      Top             =   3750
      Width           =   1130
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2010
      X2              =   3150
      Y1              =   4430
      Y2              =   4430
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2070
      X2              =   3210
      Y1              =   4760
      Y2              =   4760
   End
End
Attribute VB_Name = "frm030403"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/03/24 因為已有每月批次StrMenu15，所以可辦期限彈訊息不執行
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 22) As String, strTemp3 As String
Dim PLeft(0 To 13) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String
'Add By Cheng 2003/02/17
Dim SeekPrint As Integer, SeekPrintL As Integer
'Add By Cheng 2003/03/14
Dim m_blnSingleCase As Boolean '單筆
'add by nickc 2007/02/13
Dim tm() As String
Dim m_blnPrintAddress As Boolean '是否要列印地條
Dim m_ET01 As String, m_ET02 As String, m_ET03 As String, m_ET03_1 As String 'Add By Sindy 2012/11/19
'Add By Sindy 2013/1/4 是否要產生來函通知進度
Dim m_bolInsCP As Boolean
Dim m_TM01 As String, m_TM02 As String, m_TM03 As String, m_TM04 As String
'2013/1/4 End
Dim m_strUpdCP10 As String 'Add By Sindy 2021/9/9


Private Sub cmdok_Click(Index As Integer)
LblNote.Caption = "" 'Add By Sindy 2017/11/20
LblNote2.Caption = ""
Select Case Index
Case 0 '確定
     If Len(TXT1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         TXT1(0).SetFocus
         Exit Sub
     Else
         'Added by Lydia 2023/03/24
         If Option1(3).Value = True Then
             MsgBox "已有每月批次通知可辦期限！", vbInformation
             Exit Sub
         End If
         'end 2023/03/24
         If Option1(0).Value = True Then
            'Add By Cheng 2002/03/21
            If PUB_CheckKeyInDate(Me.TXT1(1)) = -1 Then
               Me.TXT1(1).SetFocus
               txt1_GotFocus 1
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.TXT1(2)) = -1 Then
               Me.TXT1(2).SetFocus
               txt1_GotFocus 2
               Exit Sub
            End If
            If Len(TXT1(2)) = 0 Then
                s = MsgBox("本所期限區間不可空白!!", , "USER 輸入錯誤")
                TXT1(1).SetFocus
                txt1_GotFocus (1)
                Exit Sub
            End If
        'Modify By Cheng 2003/01/08
         ElseIf Me.Option1(1).Value Then
            'Add By Cheng 2002/03/21
            If PUB_CheckKeyInDate(Me.TXT1(3)) = -1 Then
               Me.TXT1(3).SetFocus
               txt1_GotFocus 3
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.TXT1(4)) = -1 Then
               Me.TXT1(4).SetFocus
               txt1_GotFocus 4
               Exit Sub
            End If
            
            If Len(TXT1(4)) = 0 Then
                s = MsgBox("法定期限區間不可空白!!", , "USER 輸入錯誤")
                TXT1(3).SetFocus
                txt1_GotFocus (3)
                Exit Sub
            End If
         'edit by nickc 2007/03/05
         'Else
         ElseIf Option1(2).Value Then
            If Me.text1(0).Text = "" Then
                s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
                text1(0).SetFocus
                Text1_GotFocus (0)
                Exit Sub
            End If
            If Me.text1(1).Text = "" Then
                s = MsgBox("本所案號不可空白!!", , "USER 輸入錯誤")
                text1(1).SetFocus
                Text1_GotFocus (1)
                Exit Sub
            End If
            If Me.text1(2).Text = "" Then Me.text1(2).Text = "0"
            If Me.text1(3).Text = "" Then Me.text1(3).Text = "00"
         'add by nickc 2007/03/05
         ElseIf Option1(3).Value Then
            If PUB_CheckKeyInDate(Me.TXT1(18)) = -1 Then
               Me.TXT1(18).SetFocus
               txt1_GotFocus 18
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.TXT1(19)) = -1 Then
               Me.TXT1(19).SetFocus
               txt1_GotFocus 19
               Exit Sub
            End If
            
            If Len(TXT1(19)) = 0 Then
                s = MsgBox("可辦期限區間不可空白!!", , "USER 輸入錯誤")
                TXT1(18).SetFocus
                txt1_GotFocus (18)
                Exit Sub
            End If
         End If
         If InStr(1, TXT1(5), "102") <> 0 Then
             If Len(TXT1(6)) = 0 Then
                 's = MsgBox("延展報表類別不可空白!!", , "USER 輸入錯誤")
                 s = MsgBox("延展或註冊費報表類別不可空白!!", , "USER 輸入錯誤")
                 TXT1(6).SetFocus
                 Exit Sub
             End If
         End If
         If InStr(1, TXT1(5), "715") <> 0 Then
             If Len(TXT1(6)) = 0 Then
                 s = MsgBox("延展或註冊費報表類別不可空白!!", , "USER 輸入錯誤")
                 TXT1(6).SetFocus
                 Exit Sub
             End If
         End If
'            If InStr(1, TXT1(5), "716") <> 0 Then
'                If Len(TXT1(6)) = 0 Then
'                    s = MsgBox("延展或註冊費報表類別不可空白!!", , "USER 輸入錯誤")
'                    TXT1(6).SetFocus
'                    Exit Sub
'                End If
'            End If
         If Mid(TXT1(10), 1, 6) <> Mid(TXT1(11), 1, 6) Then
             s = MsgBox("申請人代號前六碼必須相同!!", , "USER 輸入錯誤")
             TXT1(11).SetFocus
             TXT1(11).SelStart = 0
             TXT1(11).SelLength = Len(TXT1(11))
             Exit Sub
         End If
         If Mid(TXT1(12), 1, 6) <> Mid(TXT1(13), 1, 6) Then
             s = MsgBox("代理人代號前六碼必須相同!!", , "USER 輸入錯誤")
             TXT1(13).SetFocus
             TXT1(13).SelStart = 0
             TXT1(13).SelLength = Len(TXT1(13))
             Exit Sub
         End If
         'add by nickc 2006/05/30
         If Val(TXT1(6)) = 1 And (InStr(1, TXT1(0), "CFT") <> 0 Or InStr(1, TXT1(0), "CFC") <> 0) Then
             If Trim(TXT1(17).Text) = "" Then
                 s = MsgBox("CFT、CFC 未收文管制表列印對象不可空白!!", , "USER 輸入錯誤")
                 TXT1(17).SetFocus
                 TXT1(17).SelStart = 0
                 TXT1(17).SelLength = Len(TXT1(17))
             Exit Sub
             End If
         End If
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/22 清除查詢印表記錄檔欄位
         pub_QL05 = pub_QL05 & ";" & Label1(2) & TXT1(6) & Label1(11) & " " & Label1(16) 'Add By Sindy 2010/10/22
         Select Case Val(TXT1(6)) '延展報表類別
         Case 2, 3 '定稿   'edit by nickc 2007/02/15  加入傳真定稿
               ProcessToWord
         Case Else '管制表
               'Add By Sindy 2023/9/20
               'Modify By Sindy 2023/10/3 加系統別判斷
               If Len(TXT1(7)) = 0 And InStr(TXT1(0), "FCT") > 0 Then
                  s = MsgBox("報表類別(1、4)組別，不可空白!!", , "USER 輸入錯誤")
                  TXT1(7).SetFocus
                  Exit Sub
               ElseIf InStr(TXT1(0), "FCT") = 0 And Len(TXT1(7)) > 0 Then
                  s = MsgBox("報表類別(1、4)組別，不必輸入!!", , "USER 輸入錯誤")
                  TXT1(7).SetFocus
                  Exit Sub
               End If
               '2023/9/20 END
               Screen.MousePointer = vbHourglass
               Me.Enabled = False
               Process
               Me.Enabled = True
               Screen.MousePointer = vbDefault
         End Select
     End If
Case 1 '結束
    Me.Enabled = False
    'Add By Cheng 2003/02/17
    '列印地址條
'move to unload by nick 2004/10/22
'    PUB_PrintAddressList strUserNum, Me.Combo1.Text
'    '刪除地址條列表資料
'    PUB_DeleteAddressList strUserNum
'    '初始化序號
'    pub_AddressListSN = 0
'    '若印表機變動, 則更新列印設定
'    If Me.Combo1.Text <> Me.Combo1.Tag Then
'        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
'    End If
    Unload Me
Case Else
End Select
End Sub
'定稿
Sub ProcessToWord()
'Add By Cheng 2003/03/14
Dim StrSQLa As String
Dim StrSqlB As String
Dim strSQLc As String
Dim strSQLD As String
Dim strSQLE As String
Dim rsA As New ADODB.Recordset
Dim rsB As New ADODB.Recordset
'Add By Sindy 2011/1/12
Dim intRow As Integer
Dim strST15 As String, strNP10 As String
'2011/1/12 End
'Add By Sindy 2013/1/4
Dim strSQLCP
Dim rsC As New ADODB.Recordset
'2013/1/4 End
Dim bolPrintLetter As Boolean, bolIsNoticeScale As Boolean 'Add By Sindy 2013/8/15
'Dim bolConnect As Boolean 'Add By Sindy 2019/1/25
Dim iRtn As Integer 'Add By Sindy 2019/1/25
Dim strTmp As String 'Add By Sindy 2024/4/16
Dim bolChkTM141 As Boolean, strTM141 As String 'Add By Sindy 2025/3/11

On Error GoTo ErrHand

Screen.MousePointer = vbHourglass

'Add By Sindy 2024/4/16
strTmp = "無;"
For i = 0 To List1(1).ListCount - 1
   strTmp = strTmp & Replace(List1(1).List(i), "-", "") & ","
Next
'2024/4/16 END

strSQL1 = "" 'TM
strSQL2 = "" 'SP
StrSQL6 = "" '共用
If Len(TXT1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND NP02 IN (" & SQLGrpStr(TXT1(0), 2) & ") "
   strSQL2 = strSQL2 + " AND NP02 IN (" & SQLGrpStr(TXT1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & TXT1(0)  'Add By Sindy 2010/10/22
End If
StrSQL6 = ""
If Len(TXT1(5)) <> 0 Then
    pub_QL05 = pub_QL05 & ";" & Label1(1) & TXT1(5)  'Add By Sindy 2010/10/22
    'edit by nickc 2007/04/16
    'StrSQL6 = " AND ("
    strSQL1 = strSQL1 & " AND ("
    strSQL2 = strSQL2 & " AND ("
    If Len(TXT1(5)) <> 0 Then
        strTemp1 = ""
        strTemp1 = Split(Replace(TXT1(5), ",,", ""), ",")
        For i = 0 To UBound(strTemp1)
            'edit by nickc 2007/04/16 加入可辦期限條件沒專用期限的使用宣誓不用
            If Val(strTemp1(i)) = 105 And Option1(3).Value = True Then
                'edit by nickc 2007/07/11 菲律賓 無專用期的要抓                              ****回復   因為與阿蓮確認過，定稿部分，菲律賓的與可辦期限無關 2007/07/11 nickc
                strSQL1 = strSQL1 + " (tm21 is not null and NP07=" & Val(strTemp1(i)) & ") OR "
                'strSQL1 = strSQL1 + " (tm21 is not null and NP07=" & Val(strTemp1(i)) & "  and tm10<>'030') OR "
                strSQL2 = strSQL2 + " NP07=" & Val(strTemp1(i)) & " OR "
            Else
                'edit by nickc 2007/04/16
                'StrSQL6 = StrSQL6 + " NP07=" & Val(strTemp1(i)) & " OR "
                strSQL1 = strSQL1 + " NP07=" & Val(strTemp1(i)) & " OR "
                strSQL2 = strSQL2 + " NP07=" & Val(strTemp1(i)) & " OR "
            End If
        Next i
        'ediit by nickc 2007/04/16
        'StrSQL6 = StrSQL6 + " NP07=0) "
        strSQL1 = strSQL1 + " NP07=0) "
        strSQL2 = strSQL2 + " NP07=0) "
    End If
End If
StrSQL6 = StrSQL6 + " AND (NP06 IS NULL OR NP06='') "
'本所期限
If Option1(0).Value = True Then
    If Len(Trim(TXT1(1))) <> 0 Then
      StrSQL6 = StrSQL6 + " AND NP08>=" & Val(ChangeTStringToWString(TXT1(1))) & " "
    End If
    If Len(Trim(TXT1(2))) <> 0 Then
      StrSQL6 = StrSQL6 + " AND NP08<=" & Val(ChangeTStringToWString(TXT1(2))) & " "
    End If
    If Len(Trim(TXT1(1))) <> 0 Or Len(Trim(TXT1(2))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & TXT1(1) & "-" & TXT1(2)  'Add By Sindy 2010/10/22
    End If
'Modify By Cheng 2003/01/08
'法定期限
ElseIf Me.Option1(1).Value Then
    If Len(Trim(TXT1(3))) <> 0 Then
        StrSQL6 = StrSQL6 + " AND NP09>=" & Val(ChangeTStringToWString(TXT1(3))) & " "
    End If
    If Len(Trim(TXT1(4))) <> 0 Then
        StrSQL6 = StrSQL6 + " AND NP09<=" & Val(ChangeTStringToWString(TXT1(4))) & " "
    End If
    If Len(Trim(TXT1(3))) <> 0 Or Len(Trim(TXT1(4))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & TXT1(3) & "-" & TXT1(4)  'Add By Sindy 2010/10/22
    End If
'本所案號
'Else
ElseIf Option1(2).Value Then
    StrSQL6 = StrSQL6 + " AND NP02='" & Me.text1(0).Text & "' And NP03='" & Me.text1(1).Text & "' And NP04='" & Me.text1(2).Text & "' And NP05='" & Me.text1(3).Text & "' "
    pub_QL05 = pub_QL05 & ";" & Option1(2).Caption & text1(0) & "-" & text1(1) & "-" & text1(2) & "-" & text1(3) 'Add By Sindy 2010/10/22
'add by nickc 2007/03/05 加入可辦，而且只有  延展跟使用宣誓
ElseIf Option1(3).Value Then
'to_char(add_months(to_date(ChangeTStringToWString(TXT1(18)),'YYYYMMDD'),na15 * -1),'YYYYMMDD')
    If Len(Trim(TXT1(18))) <> 0 Then
    StrSQL6 = StrSQL6 + " AND to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') "
    End If
    If Len(Trim(TXT1(19))) <> 0 Then
      StrSQL6 = StrSQL6 + " AND to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') "
    End If
    If Len(Trim(TXT1(18))) <> 0 Or Len(Trim(TXT1(19))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Option1(3).Caption & TXT1(18) & "-" & TXT1(19)  'Add By Sindy 2010/10/22
    End If
End If
If Len(TXT1(8)) <> 0 Then
'    StrSQL6 = StrSQL6 + " AND NP10='" & txt1(8) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(5) & TXT1(8) & LBL1(0) 'Add By Sindy 2010/10/22
End If
If Len(TXT1(9)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP14='" & TXT1(9) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(3) & TXT1(9) & LBL1(1) 'Add By Sindy 2010/10/22
End If
If Len(Trim(TXT1(10))) <> 0 And Len(Trim(TXT1(11))) <> 0 Then
    strSQL1 = strSQL1 & " AND (TM23>='" & GetNewFagent(TXT1(10)) & "' AND TM23<='" & GetNewFagent(TXT1(11)) & "') "
    strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(TXT1(10)) & "' AND SP08<='" & GetNewFagent(TXT1(11)) & "') OR (SP58<='" & GetNewFagent(TXT1(10)) & "' AND SP58<='" & GetNewFagent(TXT1(11)) & "') OR (SP59>='" & GetNewFagent(TXT1(10)) & "' AND SP59<='" & GetNewFagent(TXT1(11)) & "')) "
Else
    If Len(Trim(TXT1(10))) <> 0 And Len(Trim(TXT1(11))) = 0 Then
        strSQL1 = strSQL1 & " AND (TM23>='" & GetNewFagent(TXT1(10)) & "' ) "
        strSQL2 = strSQL2 + " AND (SP08>='" & GetNewFagent(TXT1(10)) & "' OR SP58>='" & GetNewFagent(TXT1(10)) & "' OR SP59>='" & GetNewFagent(TXT1(10)) & "') "
    Else
        If Len(Trim(TXT1(10))) = 0 And Len(Trim(TXT1(11))) <> 0 Then
            strSQL1 = strSQL1 & " AND (TM23<='" & GetNewFagent(TXT1(11)) & "') "
            strSQL2 = strSQL2 + " AND (SP08<='" & GetNewFagent(TXT1(11)) & "' OR SP58<='" & GetNewFagent(TXT1(11)) & "' OR SP59<='" & GetNewFagent(TXT1(11)) & "') "
        End If
    End If
End If
If Len(Trim(TXT1(10))) <> 0 Or Len(Trim(TXT1(11))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(7) & TXT1(10) & "-" & TXT1(11) 'Add By Sindy 2010/10/22
End If
If Len(Trim(TXT1(12))) <> 0 And Len(Trim(TXT1(13))) <> 0 Then
    strSQL1 = strSQL1 + " AND (TM44>='" & GetNewFagent(TXT1(12)) & "' AND TM44<='" & GetNewFagent(TXT1(13)) & "') "
    strSQL2 = strSQL2 + " AND (SP26>='" & GetNewFagent(TXT1(12)) & "' AND SP26<='" & GetNewFagent(TXT1(13)) & "') "
Else
    If Len(Trim(TXT1(12))) <> 0 And Len(Trim(TXT1(13))) = 0 Then
        strSQL1 = strSQL1 + " AND (TM44>='" & GetNewFagent(TXT1(12)) & "' ) "
        strSQL2 = strSQL2 + " AND (SP26>='" & GetNewFagent(TXT1(12)) & "' ) "
    Else
        If Len(Trim(TXT1(12))) = 0 And Len(Trim(TXT1(13))) <> 0 Then
            strSQL1 = strSQL1 + " AND (TM44<='" & GetNewFagent(TXT1(13)) & "' ) "
            strSQL2 = strSQL2 + " AND (SP26<='" & GetNewFagent(TXT1(13)) & "' ) "
        End If
    End If
End If
If Len(Trim(TXT1(12))) <> 0 Or Len(Trim(TXT1(13))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(8) & TXT1(12) & "-" & TXT1(13) 'Add By Sindy 2010/10/22
End If
'add by nick 2005/02/15 加入申請國家
If Len(Trim(TXT1(15))) <> 0 Then
   strSQL1 = strSQL1 & " and tm10>='" & TXT1(15) & "' "
   strSQL2 = strSQL2 & " and sp09>='" & TXT1(15) & "' "
End If
If Len(Trim(TXT1(16))) <> 0 Then
   strSQL1 = strSQL1 & " and tm10<='" & TXT1(16) & "' "
   strSQL2 = strSQL2 & " and sp09<='" & TXT1(16) & "' "
End If
'add end
If Len(Trim(TXT1(15))) <> 0 Or Len(Trim(TXT1(16))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(13) & TXT1(15) & "-" & TXT1(16) 'Add By Sindy 2010/10/22
End If
strSQL1 = strSQL1 + " AND (tm29 is null or tm29 <> 'Y' ) "
strSQL2 = strSQL2 + " AND (SP15 IS NULL OR SP15 <> 'Y' ) "
StrSQL6 = StrSQL6 + " AND NOT(NP02='FCT' AND substr(FA10,1,3)='011') " 'Add By Sindy 2023/9/20 排除FCT日本案

'add by nickc 2006/05/30
If Val(TXT1(6)) = 1 And (InStr(1, TXT1(0), "CFT") <> 0 Or InStr(1, TXT1(0), "CFC") <> 0) Then
    If TXT1(17) = "1" Then
'        StrSQL6 = StrSQL6 + " AND substr(s1.ST15,1,1)<>'S' "
        pub_QL05 = pub_QL05 & ";" & Label1(9) & TXT1(17) & Label1(14) 'Add By Sindy 2010/10/22
    End If
End If
'add by nickc 2006/05/30 延展和第二期專用權須存在  FCT  才做
If InStr(1, TXT1(5), "716") <> 0 Or InStr(1, TXT1(5), "102") <> 0 Then
    'edit by nickc 2007/07/13 加入CFT
    'strSQL1 = strSQL1 & " and np02||decode(np07,716,tm17,102,tm17,'Y')='FCTY' "
    strSQL1 = strSQL1 & " and decode(np02,'CFT','FCTY',np02||decode(np07,716,tm17,102,tm17,'Y'))='FCTY' "
End If

CheckOC
'Modify By Cheng 2003/02/17
'strSQL = "SELECT NVL(A0902,A0903),NP10,NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(TM15,TM12),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NP15,S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(N1.NA03,N1.NA04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),NVL(N2.NA03,N2.NA04),TM22,np07,NP01,CP27,TM10,NP02 FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,FAGENT,CUSTOMER WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND FA10=N2.NA01(+) AND TM10=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND S1.ST03=A0901(+) AND SUBSTR(TM44,1,8)=FA01(+) AND SUBSTR(TM44,9,1)=FA02(+) " & strSQL1 & StrSQL6
'strSQL = strSQL + " union all select NVL(A0902,A0903),NP10,NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),'',SP11,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NP15,S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(N1.NA03,N1.NA04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),NVL(N2.NA03,N2.NA04),SP21,NP07,NP01,CP27,SP09,NP02 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,FAGENT,CUSTOMER WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND FA10=N2.NA01(+) AND SP09=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND S1.ST03=A0901(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) " & strSQL2 & StrSQL6
'Modify By Sindy 2024/4/16 剔除不需定稿通知的案號
'                          + AND instr('" & strTmp & "',NP02||NP03||NP04||NP05)=0
strSql = "SELECT NVL(A0902,A0903),NP10,NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(TM15,TM12),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NP15,S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(N1.NA03,N1.NA04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),NVL(N2.NA03,N2.NA04),TM22,np07,NP01,CP27,TM10,NP02,NP03,NP04,NP05,S1.ST15 as ST15" & _
         " FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,FAGENT,CUSTOMER " & _
         " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND FA10=N2.NA01(+) AND TM10=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND S1.ST03=A0901(+) AND SUBSTR(TM44,1,8)=FA01(+) AND SUBSTR(TM44,9,1)=FA02(+) " & strSQL1 & StrSQL6 & _
         " AND instr('" & strTmp & "',NP02||NP03||NP04||NP05)=0"
strSql = strSql + " union all select NVL(A0902,A0903),NP10,NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),'',SP11,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NP15,S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(N1.NA03,N1.NA04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),NVL(N2.NA03,N2.NA04),SP21,NP07,NP01,CP27,SP09,NP02,NP03,NP04,NP05,S1.ST15 as ST15" & _
         " FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,FAGENT,CUSTOMER " & _
         " WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND FA10=N2.NA01(+) AND SP09=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND S1.ST03=A0901(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) " & strSQL2 & StrSQL6 & _
         " AND instr('" & strTmp & "',NP02||NP03||NP04||NP05)=0"
'Add By Cheng 2003/03/14
'FCT的案件
'edit by nick 2004/07/28 加入定稿語文
'strSQLA = "SELECT TM23, TM44, COUNT(*) FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,FAGENT,CUSTOMER " & _
'                " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND FA10=N2.NA01(+) AND TM10=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND S1.ST03=A0901(+) AND SUBSTR(TM44,1,8)=FA01(+) AND SUBSTR(TM44,9,1)=FA02(+) " & strSQL1 & StrSQL6 & " AND NP02='FCT' GROUP BY TM23, TM44 "
'strSQLA = strSQLA & " UNION SELECT SP08, SP26, COUNT(*) FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,FAGENT,CUSTOMER " & _
'                " WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND FA10=N2.NA01(+) AND SP09=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND S1.ST03=A0901(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) " & strSQL2 & StrSQL6 & " AND NP02='FCT' GROUP BY SP08, SP26 "
'Modify By Sindy 2016/11/22 tm53 ==> decode(tm53,'3','3','')
'                           sp34 ==> decode(sp34,'3','3','')
'ex:X70054000 Y52050020          1 2
'   X70054000 Y52050020          1
'   FCT-34431
'   FCT-34611
'Modify By Sindy 2024/4/16 剔除不需定稿通知的案號
'                          + AND instr('" & strTmp & "',NP02||NP03||NP04||NP05)=0
StrSQLa = "SELECT TM23, TM44, COUNT(*),decode(tm53,'3','3','') FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,FAGENT,CUSTOMER " & _
          " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND FA10=N2.NA01(+) AND TM10=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND S1.ST03=A0901(+) AND SUBSTR(TM44,1,8)=FA01(+) AND SUBSTR(TM44,9,1)=FA02(+) " & strSQL1 & StrSQL6 & " AND NP02='FCT'" & _
          " AND instr('" & strTmp & "',NP02||NP03||NP04||NP05)=0" & _
          " GROUP BY TM23, TM44,decode(tm53,'3','3','') "
StrSQLa = StrSQLa & " UNION SELECT SP08, SP26, COUNT(*),decode(sp34,'3','3','') FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,FAGENT,CUSTOMER " & _
          " WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND FA10=N2.NA01(+) AND SP09=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND S1.ST03=A0901(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) " & strSQL2 & StrSQL6 & " AND NP02='FCT'" & _
          " AND instr('" & strTmp & "',NP02||NP03||NP04||NP05)=0" & _
          " GROUP BY SP08, SP26,decode(sp34,'3','3','') "
'Add By Cheng 2003/04/23
StrSQLa = StrSQLa & " Order By 3 Desc, 2, 1 "
'Modify By Sindy 2015/4/30 +NP22
'Modify By Sindy 2024/4/16 剔除不需定稿通知的案號
'                          + AND instr('" & strTmp & "',NP02||NP03||NP04||NP05)=0
'Modify By Sindy 2025/3/11 + TM141延展折扣
StrSqlB = "SELECT NVL(A0902,A0903),NP10,NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(TM15,TM12),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NP15,S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(N1.NA03,N1.NA04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),NVL(N2.NA03,N2.NA04),TM22,np07,NP01,CP27,TM10,NP02,NP03,NP04,NP05,S1.ST15 as ST15,NP22,TM141" & _
          " FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,FAGENT,CUSTOMER " & _
          " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND FA10=N2.NA01(+) AND TM10=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND S1.ST03=A0901(+) AND SUBSTR(TM44,1,8)=FA01(+) AND SUBSTR(TM44,9,1)=FA02(+) " & strSQL1 & StrSQL6 & " AND NP02='FCT'" & _
          " AND instr('" & strTmp & "',NP02||NP03||NP04||NP05)=0"
strSQLc = "SELECT NVL(A0902,A0903),NP10,NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),'',SP11,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NP15,S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(N1.NA03,N1.NA04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),NVL(N2.NA03,N2.NA04),SP21,NP07,NP01,CP27,SP09,NP02,NP03,NP04,NP05,S1.ST15 as ST15,NP22,0 TM141" & _
          " FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,FAGENT,CUSTOMER " & _
          " WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND FA10=N2.NA01(+) AND SP09=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND S1.ST03=A0901(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) " & strSQL2 & StrSQL6 & " AND NP02='FCT'" & _
          " AND instr('" & strTmp & "',NP02||NP03||NP04||NP05)=0"
'2015/4/30 END
'Add By Sindy 2013/1/4
m_bolInsCP = False
If InStr(TXT1(5), "102") > 0 Or InStr(TXT1(5), "716") > 0 Then
   'Modify By Sindy 2024/4/16 剔除不需定稿通知的案號
   '                          + AND instr('" & strTmp & "',NP02||NP03||NP04||NP05)=0
   strSQLCP = "SELECT C2.CP05,C2.CP01,C2.CP02,C2.CP03,C2.CP04 FROM CASEPROGRESS C2,(" & _
              "SELECT TM01,TM02,TM03,TM04 FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,FAGENT,CUSTOMER " & _
                   " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND FA10=N2.NA01(+) AND TM10=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND S1.ST03=A0901(+) AND SUBSTR(TM44,1,8)=FA01(+) AND SUBSTR(TM44,9,1)=FA02(+) " & strSQL1 & StrSQL6 & " AND NP02='FCT'" & _
                   " AND instr('" & strTmp & "',NP02||NP03||NP04||NP05)=0"
   strSQLCP = strSQLCP & " UNION SELECT SP01,SP02,SP03,SP04 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,ACC090,FAGENT,CUSTOMER " & _
                   " WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND FA10=N2.NA01(+) AND SP09=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND S1.ST03=A0901(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) " & strSQL2 & StrSQL6 & " AND NP02='FCT'" & _
                   " AND instr('" & strTmp & "',NP02||NP03||NP04||NP05)=0" & _
                   ") T" & _
              " WHERE C2.CP01=T.TM01 AND C2.CP02=T.TM02 AND C2.CP03=T.TM03 AND C2.CP04=T.TM04"
   If TXT1(6) = "2" Then '定稿
'      If InStr(TXT1(5), "102") > 0 And InStr(TXT1(5), "716") > 0 Then
'         strSQLCP = strSQLCP & " AND C2.CP10 in('1717','1716')"
'      ElseIf InStr(TXT1(5), "102") > 0 Then
         strSQLCP = strSQLCP & " AND C2.CP10 in('1717')"
'      Else
'         strSQLCP = strSQLCP & " AND C2.CP10 in('1716')"
'      End If
   Else '傳真定稿
'      If InStr(TXT1(5), "102") > 0 And InStr(TXT1(5), "716") > 0 Then
'         strSQLCP = strSQLCP & " AND C2.CP10 in('1722','1721')"
'      ElseIf InStr(TXT1(5), "102") > 0 Then
         strSQLCP = strSQLCP & " AND C2.CP10 in('1722')"
'      Else
'         strSQLCP = strSQLCP & " AND C2.CP10 in('1721')"
'      End If
   End If
   strSQLCP = strSQLCP & " AND C2.CP05>=" & CompWorkDay(30, strSrvDate(1), 1) '一個月內 Add By Sindy 2023/2/8 增加此條件筆數才會正確
   strSQLCP = strSQLCP & " GROUP By C2.CP05,C2.CP01,C2.CP02,C2.CP03,C2.CP04 Order By C2.CP05 desc"
End If
'2013/1/4 End
intRow = 0 'Add By Sindy 2011/1/12
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
'        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/22
        .MoveFirst
        DoEvents
        Do While .EOF = False
            'Add By Sindy 2011/1/12
            strST15 = "" & .Fields("ST15")
            strNP10 = "" & .Fields("NP10")
            '檢查若智權人員離職時, 需要重新取得目前承辦智權人員
            strSql = "select st03 from staff where upper(st01)=" & CNULL(UCase("" & .Fields("NP10"))) & " and st04<>'1' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If Left(RsTemp("st03"), 1) = "F" Then
                  '取得FCT承辦智權人員
                  strNP10 = PUB_GetFCTSalesNo("" & .Fields("NP02"), "" & .Fields("NP03"), "" & .Fields("NP04"), "" & .Fields("NP05"))
                  strST15 = PUB_GetStaffST15(PUB_GetFCTSalesNo("" & .Fields("NP02"), "" & .Fields("NP03"), "" & .Fields("NP04"), "" & .Fields("NP05")), "1")
               Else
                  '取得目前承辦智權人員
                  strNP10 = PUB_GetAKindSalesNo("" & .Fields("NP02"), "" & .Fields("NP03"), "" & .Fields("NP04"), "" & .Fields("NP05"))
                  strST15 = PUB_GetStaffST15(PUB_GetAKindSalesNo("" & .Fields("NP02"), "" & .Fields("NP03"), "" & .Fields("NP04"), "" & .Fields("NP05")), "1")
               End If
            End If
            If Len(TXT1(8)) <> 0 Then
                If strNP10 <> Trim(TXT1(8)) Then GoTo GoToExit2
            End If
            If Val(TXT1(6)) = 1 And (InStr(1, TXT1(0), "CFT") <> 0 Or InStr(1, TXT1(0), "CFC") <> 0) Then
               If TXT1(17) = "1" Then
                  If Left(strST15, 1) = "S" Then GoTo GoToExit2
               End If
            End If
            intRow = intRow + 1 '記錄筆數
            '2011/1/12 End
            
            'Add By Cheng 2003/03/14
            '先排除系統類別為FCT
            If "" & .Fields(20).Value <> "FCT" Then
                'Modify By Cheng 2002/12/29
    '            PrintLetter CheckStr(.Fields(16)), CheckStr(.Fields(20)), CheckStr(.Fields(19)), "", GetTodayDate, CheckStr(.Fields(17)), CheckStr(.Fields(4))
                'Modify By Sindy 2012/11/19
                'PrintLetter CheckStr(.Fields(16)), CheckStr(.Fields(20)), CheckStr(.Fields(19)), "", GetTodayDate, CheckStr(.Fields(17)), CheckStr(.Fields(4)), "" & .Fields("NP08").Value, "" & .Fields("NP09").Value
                m_ET01 = "": m_ET02 = "": m_ET03 = "": m_ET03_1 = ""
                PrintLetter CheckStr(.Fields(16)), CheckStr(.Fields(20)), CheckStr(.Fields(19)), "", GetTodayDate, CheckStr(.Fields(17)), CheckStr(.Fields(4)), "" & .Fields("NP08").Value, "" & .Fields("NP09").Value, m_ET01, m_ET02, m_ET03, m_ET03_1, rsB
                '2012/11/19 End
                'Add By Cheng 2003/02/17
                '新增地址條列表資料
                'add by nickc 2007/03/13
                If TXT1(6) <> "3" Then
                     pub_AddressListSN = pub_AddressListSN + 1
                     If m_blnPrintAddress = True Then
                        '2009/12/10 MODIFY BY SONIA加傳案件性質否則延展代理人名條會抓錯FCT-016425
                        PUB_AddNewAddressList strUserNum, .Fields("NP02").Value, .Fields("NP03").Value, .Fields("NP04").Value, .Fields("NP05").Value, "" & pub_AddressListSN, "0", .Fields("NP07").Value
                     End If
                End If
            End If
         
GoToExit2: 'Add By Sindy 2011/1/12
            '移至下一筆
            .MoveNext
        Loop
        'Add By Sindy 2011/1/12
        If intRow = 0 Then
            InsertQueryLog (0)
            ShowNoData
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            InsertQueryLog (intRow)
        '2011/1/12 End
            '若有下FCT的系統類別
            If InStr(Me.TXT1(0).Text, "FCT") > 0 Then
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  'Add By Sindy 2013/1/4 要產生來函通知進度
                  'If InStr(TXT1(5), "102") > 0 Or InStr(TXT1(5), "716") > 0 Then
                  If InStr(TXT1(5), "102") > 0 Then
                     m_bolInsCP = True
                     
                     'Modify By Sindy 2021/9/9 要更新的案件性質只要判斷一次即可, 不用放在 FCTInsertCP 逐筆判斷;
                     '有發生 6筆 掉了案件性質，110/9/6 15:12同時間產生的(也沒有產生定稿)
                     'FCT -34141
                     'FCT -47384
                     'FCT -32173
                     'FCT -31096
                     'FCT -19254
                     'FCT -36987
                     If TXT1(6) = "2" Then '定稿-延展
                        '本所通知延展
                        m_strUpdCP10 = "1717"
                     ElseIf TXT1(6) = "3" Then '傳真定稿-延展
                        '本所再通知延展
                        m_strUpdCP10 = "1722"
                     Else
                        MsgBox "無讀取到案件性質無法新增進度！", vbExclamation
                        Screen.MousePointer = vbDefault
                        Exit Sub
                     End If
                     '2021/9/9 END
                     
                     rsC.CursorLocation = adUseClient
                     rsC.Open strSQLCP, cnnConnection, adOpenStatic, adLockReadOnly
                     If rsC.RecordCount > 0 Then
                        rsC.MoveFirst
                        '檢查7天內重新以此方式作業,以詢問方式確認是否要產生通知進度
                        'Modify By Sindy 2013/12/3
                        'If CompWorkDay(7, rsC.Fields(0), 0) >= strSrvDate(1) Then
                        If CompWorkDay(30, rsC.Fields(0), 0) >= strSrvDate(1) Then
                        '2013/12/3 END
                           'Modify By Sindy 2019/1/25
                           'If MsgBox("是否產生來函通知進度？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'                              m_bolInsCP = False
'                           End If
                           'Add By Sindy 2024/4/23
                           '列印條件有輸入本所案號或申請人或代理人，維持提醒，其他取消訊息
                           If text1(1) <> "" Or TXT1(10) <> "" Or TXT1(12) <> "" Then
                           '2024/4/23 END
                              iRtn = MsgBox("一個月內進度檔已有 " & rsC.RecordCount & " 筆本所通知延展" & vbCrLf & vbCrLf & "請確認，此次是否要產生本所通知延展進度？" & vbCrLf & vbCrLf & _
                                            "是:產生進度　否:不產生進度　取消:放棄執行此作業", vbYesNoCancel + vbDefaultButton3 + vbExclamation)
                              If iRtn = vbCancel Then
                                 Screen.MousePointer = vbDefault
                                 Exit Sub
                              ElseIf iRtn = vbNo Then
                                 m_bolInsCP = False
                              End If
                              '2019/1/25 END
                           End If
                        End If
                     End If
                     If rsC.State <> adStateClosed Then rsC.Close
                     Set rsC = Nothing
                  End If
                  '2013/1/4 End
'                  cnnConnection.BeginTrans: bolConnect = True 'Add By Sindy 2019/1/25
                  While Not rsA.EOF
                     '若有多筆
                     If rsA.Fields(2).Value > 1 Then
                        'edit by nick 2004/07/26 若是日文定稿，都定為單筆
                        'If GetLetterLanguage(SystemNumber(strTM, 1), SystemNumber(strTM, 2), SystemNumber(strTM, 3), SystemNumber(strTM, 4)) = "3" Then
                        If CheckStr(rsA.Fields(3).Value) = "3" Then
                            m_blnSingleCase = True
                        Else
                            m_blnSingleCase = False
                        End If
                        'm_blnSingleCase = False
                        '2006/5/4 ADD BY SONIA 第二期註冊費都定為單筆
                        If InStr(1, TXT1(5), "716") <> 0 Then
                           m_blnSingleCase = True
                        End If
                        '2006/5/4 END
                     '若只有一筆
                     Else
                        m_blnSingleCase = True
                     End If
                     
                     strSQLD = "": strSQLE = ""
                     '若有申請人
                     If "" & rsA.Fields(0).Value <> "" Then
                         strSQLD = strSQLD & " AND TM23='" & rsA.Fields(0).Value & "' "
                         strSQLE = strSQLE & " AND SP08='" & rsA.Fields(0).Value & "' "
                     Else
                         strSQLD = strSQLD & " AND TM23 IS NULL "
                         strSQLE = strSQLE & " AND SP08 IS NULL "
                     End If
                     '若有代理人
                     If "" & rsA.Fields(1).Value <> "" Then
                         strSQLD = strSQLD & " AND TM44='" & rsA.Fields(1).Value & "' "
                         strSQLE = strSQLE & " AND SP26='" & rsA.Fields(1).Value & "' "
                     Else
                         strSQLD = strSQLD & " AND TM44 IS NULL "
                         strSQLE = strSQLE & " AND SP26 IS NULL "
                     End If
                     rsB.CursorLocation = adUseClient
                     rsB.Open StrSqlB & strSQLD & " UNION ALL " & strSQLc & strSQLE, cnnConnection, adOpenStatic, adLockReadOnly
                     If rsB.RecordCount > 0 Then
                        'Add By Sindy 2011/1/12
                        strST15 = "" & rsB.Fields("ST15")
                        strNP10 = "" & rsB.Fields("NP10")
                        '檢查若智權人員離職時, 需要重新取得目前承辦智權人員
                        strSql = "select st03 from staff where upper(st01)=" & CNULL(UCase("" & rsB.Fields("NP10"))) & " and st04<>'1' "
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           If Left(RsTemp("st03"), 1) = "F" Then
                              '取得FCT承辦智權人員
                              strNP10 = PUB_GetFCTSalesNo("" & rsB.Fields("NP02"), "" & rsB.Fields("NP03"), "" & rsB.Fields("NP04"), "" & rsB.Fields("NP05"))
                              strST15 = PUB_GetStaffST15(PUB_GetFCTSalesNo("" & rsB.Fields("NP02"), "" & rsB.Fields("NP03"), "" & rsB.Fields("NP04"), "" & rsB.Fields("NP05")), "1")
                           Else
                              '取得目前承辦智權人員
                              strNP10 = PUB_GetAKindSalesNo("" & rsB.Fields("NP02"), "" & rsB.Fields("NP03"), "" & rsB.Fields("NP04"), "" & rsB.Fields("NP05"))
                              strST15 = PUB_GetStaffST15(PUB_GetAKindSalesNo("" & rsB.Fields("NP02"), "" & rsB.Fields("NP03"), "" & rsB.Fields("NP04"), "" & rsB.Fields("NP05")), "1")
                           End If
                        End If
                        If Len(TXT1(8)) <> 0 Then
                            If strNP10 <> Trim(TXT1(8)) Then GoTo GoToExit3
                        End If
                        If Val(TXT1(6)) = 1 And (InStr(1, TXT1(0), "CFT") <> 0 Or InStr(1, TXT1(0), "CFC") <> 0) Then
                           If TXT1(17) = "1" Then
                              If Left(strST15, 1) = "S" Then GoTo GoToExit3
                           End If
                        End If
                        '2011/1/12 End
                        
                        'Add By Cheng 2003/04/22
                        '初始化產生定稿的時間
                        pub_strDate = "00000000"
                        pub_strTime = "000000"
                        '新增地址條列表資料
                        'add by nickc 2007/03/13
                        If TXT1(6) <> "3" Then
                          'Modify By Sindy 2015/2/4 列印催延展定稿, 不催延展者之案件地址條也不要印 !
                          bolIsNoticeScale = True
                          If rsB.Fields(20) = "FCT" And rsB.Fields(16) = "102" Then
                             bolIsNoticeScale = PUB_ChkCaseIsNoticeScale(rsB.Fields(20), rsB.Fields(21), rsB.Fields(22), rsB.Fields(23))
                          End If
                          If bolIsNoticeScale = True Then
                          '2015/2/4 END
                             pub_AddressListSN = pub_AddressListSN + 1
                             '2009/12/10 MODIFY BY SONIA加傳案件性質否則延展代理人名條會抓錯FCT-016425
                             PUB_AddNewAddressList strUserNum, rsB.Fields("NP02").Value, rsB.Fields("NP03").Value, rsB.Fields("NP04").Value, rsB.Fields("NP05").Value, "" & pub_AddressListSN, "0", rsB.Fields("NP07").Value
                          End If
                        End If
                        
                        'Add By Sindy 2013/1/4 要產生來函通知進度
                        'If rsB.RecordCount > 0 Then
                          bolPrintLetter = False 'Add By Sindy 2013/8/15
                          rsB.MoveFirst
                          bolChkTM141 = True: strTM141 = "S" 'Add By Sindy 2025/3/11 預設值
                          While Not rsB.EOF
                              'Add By Sindy 2025/3/11 多件時,檢查是否有不一致的折扣,顯示提醒訊息
                              If bolChkTM141 = True And strTM141 <> "S" Then
                                 If strTM141 <> "" & rsB.Fields("TM141") Then
                                    bolChkTM141 = False
                                    strExc(10) = ""
                                    '代理人
                                    If "" & rsA.Fields(1).Value <> "" Then
                                       strExc(10) = strExc(10) & ",代理人= " & rsA.Fields(1).Value
                                    End If
                                    '申請人
                                    If "" & rsA.Fields(0).Value <> "" Then
                                       strExc(10) = strExc(10) & ",申請人= " & rsA.Fields(0).Value
                                    End If
                                    strExc(10) = Mid(strExc(10), 2)
                                    MsgBox strExc(10) & vbCrLf & "多件個案折扣有不一致狀況, 請檢查!", vbExclamation
                                 End If
                              End If
                              strTM141 = "" & rsB.Fields("TM141")
                              '2025/3/11 END
                             'Modify By Sindy 2013/8/15
                             If rsB.Fields(20) = "FCT" And rsB.Fields(16) = "716" And m_bolInsCP = True Then
                                'Modify By Sindy 2015/4/30 +,CheckStr("" & rsB.Fields("np22"))
                                Call FCTInsertCP(CheckStr(rsB.Fields(16)), CheckStr(rsB.Fields(17)), CheckStr(rsB.Fields(4)), CheckStr("" & rsB.Fields("np22")))
                                bolPrintLetter = True
                             ElseIf rsB.Fields(20) = "FCT" And rsB.Fields(16) = "102" Then
                                '不催延展者,不產生進度,不出定稿
                                bolIsNoticeScale = PUB_ChkCaseIsNoticeScale(rsB.Fields(20), rsB.Fields(21), rsB.Fields(22), rsB.Fields(23))
                                If bolIsNoticeScale = True Then
                                   bolPrintLetter = True
                                End If
                                If m_bolInsCP = True And bolIsNoticeScale = True Then
                                   'Modify By Sindy 2015/4/30 +,CheckStr("" & rsB.Fields("np22"))
                                   Call FCTInsertCP(CheckStr(rsB.Fields(16)), CheckStr(rsB.Fields(17)), CheckStr(rsB.Fields(4)), CheckStr("" & rsB.Fields("np22")))
                                End If
                             Else
                                bolPrintLetter = True
                             End If
                             '2013/8/15 END
                             rsB.MoveNext
                          Wend
                          rsB.MoveFirst '***
                        'End If
                        '2013/1/4 END
                        
                        '產生定稿
                        '2006/5/11 ADD BY SONIA 第二期註冊費多筆時要逐筆列印定稿
                        'PrintLetter CheckStr(rsB.Fields(16)), CheckStr(rsB.Fields(20)), CheckStr(rsB.Fields(19)), "", GetTodayDate, CheckStr(rsB.Fields(17)), CheckStr(rsB.Fields(4)), "" & rsB.Fields("NP08").Value, "" & rsB.Fields("NP09").Value
                        'DoEvents
                        If InStr(1, TXT1(5), "716") <> 0 And rsB.RecordCount > 1 Then
                           While Not rsB.EOF
                              '產生定稿
                              'Modify By Sindy 2012/11/19
                              'PrintLetter CheckStr(rsB.Fields(16)), CheckStr(rsB.Fields(20)), CheckStr(rsB.Fields(19)), "", GetTodayDate, CheckStr(rsB.Fields(17)), CheckStr(rsB.Fields(4)), "" & rsB.Fields("NP08").Value, "" & rsB.Fields("NP09").Value
                              m_ET01 = "": m_ET02 = "": m_ET03 = "": m_ET03_1 = ""
                              PrintLetter CheckStr(rsB.Fields(16)), CheckStr(rsB.Fields(20)), CheckStr(rsB.Fields(19)), "", GetTodayDate, CheckStr(rsB.Fields(17)), CheckStr(rsB.Fields(4)), "" & rsB.Fields("NP08").Value, "" & rsB.Fields("NP09").Value, m_ET01, m_ET02, m_ET03, m_ET03_1, rsB
                              '2012/11/19 End
                              DoEvents
                              rsB.MoveNext
                           Wend
                        Else
                           'Modify By Sindy 2012/11/19
                           'PrintLetter CheckStr(rsB.Fields(16)), CheckStr(rsB.Fields(20)), CheckStr(rsB.Fields(19)), "", GetTodayDate, CheckStr(rsB.Fields(17)), CheckStr(rsB.Fields(4)), "" & rsB.Fields("NP08").Value, "" & rsB.Fields("NP09").Value
                           m_ET01 = "": m_ET02 = "": m_ET03 = "": m_ET03_1 = ""
                           If bolPrintLetter = True Then 'Add By Sindy 2013/8/15
                             PrintLetter CheckStr(rsB.Fields(16)), CheckStr(rsB.Fields(20)), CheckStr(rsB.Fields(19)), "", GetTodayDate, CheckStr(rsB.Fields(17)), CheckStr(rsB.Fields(4)), "" & rsB.Fields("NP08").Value, "" & rsB.Fields("NP09").Value, m_ET01, m_ET02, m_ET03, m_ET03_1, rsB
                             '2012/11/19 End
                             DoEvents
                           Else
                             GoTo GoToExit3
                           End If
                           '2013/8/15
                        End If
                        '2006/5/11 END
                        '若有多筆
                        If m_blnSingleCase = False Then
                           'Add by Morgan 2010/4/14
                           pub_OsPrinter = PUB_GetOsDefaultPrinter
                           PUB_SetOsDefaultPrinter Printer.DeviceName
                           'end 2010/4/14
                           
                           '列印定稿
                           PrinterLetterDB strUserNum, "1", pub_strDate, pub_strTime
                           
                           PUB_SetOsDefaultPrinter pub_OsPrinter 'Add by Morgan 2010/4/14
                           
                           DoEvents
                           'Add By Cheng 2003/04/22
                           '上列印註記
                           cnnConnection.Execute "Update Letterdemand Set LD16='*' Where LD01='" & strUserNum & "' And LD02=" & pub_strDate & " And LD03=" & pub_strTime
                           'Modify By Sindy 2012/11/19 阿蓮要求FCT延展多筆時清單加入定稿中
                           'Modify By Sindy 2013/6/27 阿蓮要求FCT延展多筆[傳真]時清單加入定稿中
                           If Left(Trim(m_ET02), 3) = "FCT" And Right(Trim(m_ET02), 3) = "102" And (m_ET03 = "03" Or m_ET03 = "99") Then
                              '已加入定稿中
                           Else
                           '2012/11/19 End
                              '列印延期案件
                              PrinterDetail rsB
                           End If
                           DoEvents
                        End If
GoToExit3:
                     End If
                     If rsB.State <> adStateClosed Then rsB.Close
                     Set rsB = Nothing
                     
                     rsA.MoveNext
                  Wend
'                  cnnConnection.CommitTrans: bolConnect = False 'Add By Sindy 2019/1/25
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            End If
            'Add By Cheng 2002/12/20
            MsgBox "定稿產生完成!!!", vbExclamation + vbOKOnly
        End If '2011/1/12 End
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/22
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End With
CheckOC
Screen.MousePointer = vbDefault

Exit Sub

ErrHand:
   Screen.MousePointer = vbDefault
'   'Add By Sindy 2019/1/25
'   If bolConnect = True Then
'      cnnConnection.RollbackTrans
'   End If
'   '2019/1/25 ENd
   MsgBox " 資料產生失敗！" & vbCrLf & Err.Description & vbCrLf & " strSql=" & strSql
End Sub

'Modify By Cheng 2002/12/29
'Private Sub InsExpField(ByVal strNP07 As String, ByVal strTM01 As String, ByVal strTM10 As String, ByVal strDate As String, ByVal strSysDate As String, ByVal strCP09 As String, ByVal strTM As String)
Private Sub InsExpField(ByVal strNP07 As String, ByVal strTM01 As String, ByVal strTM10 As String, ByVal strDate As String, ByVal strSysDate As String, ByVal strCP09 As String, ByVal strTM As String, ByVal strNP08 As String, ByVal strNP09 As String, rsRS As ADODB.Recordset)
   Dim oRate As Double    '匯率
   Dim o10206 As Double  '費用
   Dim o10208 As Double  '規費
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
   'add by nickc 2007/02/09
   Dim ExFaName As String
   Dim ExFaNa As String
   Dim ExFax As String
   Dim ExTel As String
   Dim ExFaDate As String
   Dim ExSworDate As String
   Dim intFee As Integer 'Add By Sindy 2010/8/25
   Dim strTemp As String 'Add By Sindy 2012/11/19
   Dim strTM09 As String, strTmpTM09 As Variant, strFA76 As String 'Add By Sindy 2014/1/14
   Dim strNote1 As String 'Add By Sindy 2014/2/12 公式1
   Dim strNote2 As String 'Add By Sindy 2014/2/12 公式2
   Dim strDisc As String '折扣
   
   'Add By Sindy 2014/1/14
   tm(1) = SystemNumber(strTM, 1)
   tm(2) = SystemNumber(strTM, 2)
   tm(3) = SystemNumber(strTM, 3)
   tm(4) = SystemNumber(strTM, 4)
   strSql = "select tm09,fa76 from trademark,fagent where tm01='" & tm(1) & "' and tm02='" & tm(2) & "' and tm03='" & tm(3) & "' and tm04='" & tm(4) & "'" & _
            " and substr(tm44,1,8)=fa01(+) and substr(tm44,9,1)=fa02(+)"
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If AdoRecordSet3.RecordCount <> 0 Then
      strTM09 = Trim("" & AdoRecordSet3.Fields("tm09").Value)
      strFA76 = Trim("" & AdoRecordSet3.Fields("fa76").Value)
   End If
   CheckOC3
   strTmpTM09 = Split(strTM09, ",")
   '2014/1/14 END
   
   Select Case strTM01
      Case "FCT"
         '分下一程序
         Select Case strNP07
         '延展
         Case "102"
            'Add By Sindy 2025/3/5 抓催延展的折扣
            strDisc = PUB_GetA1L07Disc(tm(1), tm(2), tm(3), tm(4), strNP07, strSrvDate(2))
            If strDisc = 100 Then strDisc = ""
            
            Select Case GetLetterLanguage(SystemNumber(strTM, 1), SystemNumber(strTM, 2), SystemNumber(strTM, 3), SystemNumber(strTM, 4))
               '中文
               Case "1"
                  'add by nickc 2007/02/15
                  If Val(TXT1(6)) = 2 Then
                        'edit by nickc 2007/01/10 改成與英文同
                        'EndLetter "10", strCP09, "01", strUserNum
                        EndLetter "10", Replace(strTM, "-", "") & "&" & Trim(strNP07), "01", strUserNum
                  End If
                  
               '英文
               Case "2"
                  'add by nickc 2007/02/15
                  If Val(TXT1(6)) = 2 Then
                     'Modify By Cheng 2003/03/14
                     'EndLetter "10", strCP09, "02", strUserNum
                     'edit by nick 2004/12/13
                     'EndLetter "10", strCP09, IIf(m_blnSingleCase = True, "02", "03"), strUserNum
                     EndLetter "10", Replace(strTM, "-", "") & "&" & Trim(strNP07), IIf(m_blnSingleCase = True, "02", "03"), strUserNum
                     'Add By Sindy 2012/11/19
                     If m_blnSingleCase = False Then
                        strTemp = PrinterDetail_Text(rsRS)
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & IIf(m_blnSingleCase = True, "02", "03") & "','" & strUserNum & _
                                 "','多件清單','" & ChgSQL(strTemp) & "')"
                        cnnConnection.Execute strSql
                     End If
                     '2012/11/19 End
                     
                     'Add By Sindy 2025/3/6
                     '詢問過湘嫺多件一樣要抓折扣
                     If strDisc <> "" Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & IIf(m_blnSingleCase = True, "02", "03") & "','" & strUserNum & "','折扣',' x " & strDisc & "％')"
                        cnnConnection.Execute strSql
                     End If
                     '跨類 / 催延展多件(多類顯示方式)
                     If UBound(strTmpTM09) > 0 Or m_blnSingleCase = False Then
                        If strDisc <> "" Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & IIf(m_blnSingleCase = True, "02", "03") & "','" & strUserNum & "','加一類別有折扣','　　　　　　　　　 NT$4,000" & IIf(strDisc <> "", " x " & strDisc & "％", "") & " for each additional class')"
                           cnnConnection.Execute strSql
                        End If
                     Else '一案一類別
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & IIf(m_blnSingleCase = True, "02", "03") & "','" & strUserNum & "','一案一類別','♀')"
                        cnnConnection.Execute strSql
                        '加一類別:不顯示內容
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & IIf(m_blnSingleCase = True, "02", "03") & "','" & strUserNum & "','加一類別有折扣','♀')"
                        cnnConnection.Execute strSql
                     End If
                     '2025/3/6 END
                     
                  ElseIf Val(TXT1(6)) = 3 Then
                     'add by nickc 2007/02/13 加入傳真定稿
                     tm(1) = SystemNumber(strTM, 1)
                     tm(2) = SystemNumber(strTM, 2)
                     tm(3) = SystemNumber(strTM, 3)
                     tm(4) = SystemNumber(strTM, 4)
                     If ClsPDReadTrademarkDatabase(tm(), 國外_FC) Then
                        'Add By Sindy 2013/11/19
                        If Trim(tm(33)) <> "" Then '延展代理人
                              ExFaName = PUB_GetFAgentName(tm(33), 2)
                              ExFaNa = ""
                              ExFax = ""
                              ExTel = ""
                              ExFaDate = strNP09
                              ExSworDate = strNP08
                              strSql = "select * from fagent where fa01='" & Mid(tm(33) & "000000000", 1, 8) & "' and fa02='" & Mid(tm(33) & "000000000", 9, 1) & "' "
                              CheckOC3
                              AdoRecordSet3.CursorLocation = adUseClient
                              AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                              If AdoRecordSet3.RecordCount <> 0 Then
                                  ExFaNa = PUB_GetNationEngName(CheckStr(AdoRecordSet3.Fields("fa10").Value))
                                  ExFax = CheckStr(AdoRecordSet3.Fields("fa14").Value) & IIf(CheckStr(AdoRecordSet3.Fields("fa15").Value) = "", "　　　　　　", "," & CheckStr(AdoRecordSet3.Fields("fa15").Value))
                                  ExTel = CheckStr(AdoRecordSet3.Fields("fa12").Value) & IIf(CheckStr(AdoRecordSet3.Fields("fa13").Value) = "", "　　　　　　", "," & CheckStr(AdoRecordSet3.Fields("fa13").Value))
                              End If
                              CheckOC3
                        '2013/11/19 END
                        ElseIf Trim(tm(44)) <> "" Then
                              ExFaName = PUB_GetFAgentName(tm(44), 2)
                              ExFaNa = ""
                              ExFax = ""
                              ExTel = ""
                              ExFaDate = strNP09
                              ExSworDate = strNP08
                              strSql = "select * from fagent where fa01='" & Mid(tm(44) & "000000000", 1, 8) & "' and fa02='" & Mid(tm(44) & "000000000", 9, 1) & "' "
                              CheckOC3
                              AdoRecordSet3.CursorLocation = adUseClient
                              AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                              If AdoRecordSet3.RecordCount <> 0 Then
                                  ExFaNa = PUB_GetNationEngName(CheckStr(AdoRecordSet3.Fields("fa10").Value))
                                  ExFax = CheckStr(AdoRecordSet3.Fields("fa14").Value) & IIf(CheckStr(AdoRecordSet3.Fields("fa15").Value) = "", "　　　　　　", "," & CheckStr(AdoRecordSet3.Fields("fa15").Value))
                                  ExTel = CheckStr(AdoRecordSet3.Fields("fa12").Value) & IIf(CheckStr(AdoRecordSet3.Fields("fa13").Value) = "", "　　　　　　", "," & CheckStr(AdoRecordSet3.Fields("fa13").Value))
                              End If
                              CheckOC3
                        ElseIf Trim(tm(23)) <> "" Then
                              ExFaName = GetCustomerName(tm(23), 1)
                              ExFaNa = ""
                              ExFax = ""
                              ExTel = ""
                              ExFaDate = strNP09
                              ExSworDate = strNP08
                              strSql = "select * from customer where cu01='" & Mid(tm(23) & "000000000", 1, 8) & "' and cu02='" & Mid(tm(23) & "000000000", 9, 1) & "' "
                              CheckOC3
                              AdoRecordSet3.CursorLocation = adUseClient
                              AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                              If AdoRecordSet3.RecordCount <> 0 Then
                                  ExFaNa = PUB_GetNationEngName(CheckStr(AdoRecordSet3.Fields("cu10").Value))
                                  ExFax = CheckStr(AdoRecordSet3.Fields("cu18").Value) & IIf(CheckStr(AdoRecordSet3.Fields("cu19").Value) = "", "　　　　　　", "," & CheckStr(AdoRecordSet3.Fields("cu19").Value))
                                  ExTel = CheckStr(AdoRecordSet3.Fields("cu16").Value) & IIf(CheckStr(AdoRecordSet3.Fields("cu17").Value) = "", "　　　　　　", "," & CheckStr(AdoRecordSet3.Fields("cu17").Value))
                              End If
                              CheckOC3
                        End If
                        EndLetter "10", Replace(strTM, "-", "") & "&" & Trim(strNP07), IIf(m_blnSingleCase = True, "98", "99"), strUserNum
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & IIf(m_blnSingleCase = True, "98", "99") & "','" & strUserNum & _
                                 "','例-代理人名稱','" & ChgSQL(ExFaName) & "')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & IIf(m_blnSingleCase = True, "98", "99") & "','" & strUserNum & _
                                 "','例-代理人國籍','" & ExFaNa & "')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & IIf(m_blnSingleCase = True, "98", "99") & "','" & strUserNum & _
                                 "','例-傳真','" & ExFax & "')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & IIf(m_blnSingleCase = True, "98", "99") & "','" & strUserNum & _
                                 "','例-電話','" & ExTel & "')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & IIf(m_blnSingleCase = True, "98", "99") & "','" & strUserNum & _
                                 "','例-法定','" & ExFaDate & "')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & IIf(m_blnSingleCase = True, "98", "99") & "','" & strUserNum & _
                                 "','例-本所','" & ExSworDate & "')"
                        cnnConnection.Execute strSql
                        'Add By Sindy 2013/6/27
                        If m_blnSingleCase = False Then
                           strTemp = PrinterDetail_Text(rsRS)
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & IIf(m_blnSingleCase = True, "98", "99") & "','" & strUserNum & _
                                    "','多件清單','" & ChgSQL(strTemp) & "')"
                           cnnConnection.Execute strSql
                        End If
                        '2013/6/27 End
                     End If
                  End If
                   
               'Modify By Sindy 2025/6/13 日文組已不使用此作業定稿,改用frm030404
'               '日文
'               Case "3"
'                  'add by nickc 2007/02/15
'                  If Val(TXT1(6)) = 2 Then
'                     'edit by nickc 2007/01/10 改成與英文同
'                     'EndLetter "10", strCP09, "04", strUserNum
'                     EndLetter "10", Replace(strTM, "-", "") & "&" & Trim(strNP07), "04", strUserNum
'
'                     CheckOC3
'                     strSql = "select * from usxrate where USXR01 in (select max(USXR01) from usxrate where USXR01<=to_number(to_char(sysdate, 'YYYYMMDD'))) "
'                     AdoRecordSet3.CursorLocation = adUseClient
'                     AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                     If AdoRecordSet3.RecordCount <> 0 Then
'                         oRate = AdoRecordSet3.Fields("USXR02").Value
'                     End If
'                     CheckOC3
'                     'Add By Sindy 2011/5/30
'                     'Modify By Sindy 2014/1/14
'                     'o10206 = 12500
'                     If UBound(strTmpTM09) > 0 Then '跨類
'                         o10208 = 4000 * (UBound(strTmpTM09) + 1)
'                         strNote1 = "NT$4,000x" & (UBound(strTmpTM09) + 1) & "�P分 = " 'Add By Sindy 2014/2/12
'
'                         If strFA76 = "B" Then '客戶直接來所
'                            o10206 = 8500 + (4000 * UBound(strTmpTM09))
'                            strNote2 = "NT$8,500+(NT$4,000x" & UBound(strTmpTM09) & "�P分) = " 'Add By Sindy 2014/2/12
'
'                         Else
'                            o10206 = (8500 * (90 / 100)) + (1000 * UBound(strTmpTM09) * (50 / 100))
'                            strNote2 = "(NT$8,500x90%)+" & UBound(strTmpTM09) & "�P分目以降(NT$1,000x" & UBound(strTmpTM09) & "x50%)" & vbCrLf & "　　　　　　　　　　= " 'Add By Sindy 2014/2/12
'                         End If
'                     Else
'                         o10208 = 4000
'                         If strFA76 = "B" Then '客戶直接來所
'                            o10206 = 8500
'
'                         Else
'                            o10206 = (8500 * (90 / 100))
'                            strNote2 = "(NT$8,500x90%) = " 'Add By Sindy 2014/2/12
'                         End If
'                     End If
'                     '2014/1/14 END
'                     '2011/5/30 End
'                     'Add By Sindy 2014/2/12
'                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                              "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & "04" & "','" & strUserNum & _
'                              "','公式1','" & strNote1 & "')"
'                     cnnConnection.Execute strSql
'                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                              "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & "04" & "','" & strUserNum & _
'                              "','公式2','" & strNote2 & "')"
'                     cnnConnection.Execute strSql
'                     '2014/2/12 END
'                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                              "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & "04" & "','" & strUserNum & _
'                              "','錢1','" & Format(o10208, "###,###,##0") & "')"
'                     cnnConnection.Execute strSql
'                     intFee = o10208 / oRate 'Modify By Sindy 2010/8/25 o10208 \ oRate
'                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                              "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & "04" & "','" & strUserNum & _
'                              "','錢2','" & Format((intFee + IIf(o10208 Mod oRate <> 0, 1, 0)), "###,###,##0") & "')"
'                     cnnConnection.Execute strSql
'                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                              "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & "04" & "','" & strUserNum & _
'                              "','錢3','" & Format(o10206, "###,###,##0") & "')"
'                     cnnConnection.Execute strSql
'                     intFee = (o10206 / oRate) 'Modify By Sindy 2010/8/25 ((o10206 - o10208) \ oRate)
'                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                              "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & "04" & "','" & strUserNum & _
'                              "','錢4','" & Format((intFee + IIf(o10206 Mod oRate <> 0, 1, 0)), "###,###,##0") & "')"
'                     cnnConnection.Execute strSql
'                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                              "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & "04" & "','" & strUserNum & _
'                              "','錢5','" & Format(3500, "###,###,##0") & "')"
'                     cnnConnection.Execute strSql
'                     intFee = 3500 / oRate 'Modify By Sindy 2010/8/25 3500 \ oRate
'                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                              "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & "04" & "','" & strUserNum & _
'                              "','錢6','" & Format((intFee + IIf(3500 Mod oRate <> 0, 1, 0)), "###,###,##0") & "')"
'                     cnnConnection.Execute strSql
'                  End If
            End Select

'         '2006/3/10 ADD BY SONIA 第二期註冊費
'         'edit by nickc 2006/06/01
'         Case "716"
'            'add by nickc 2006/06/01  加入日文定稿
'            Select Case GetLetterLanguage(SystemNumber(strTM, 1), SystemNumber(strTM, 2), SystemNumber(strTM, 3), SystemNumber(strTM, 4))
'            Case "3"
'                  'add by nickc 2007/02/15
'                  If Val(TXT1(6)) = 2 Then
'                          EndLetter "10", strCP09, "07", strUserNum
'
'                          CheckOC3
'                          strSql = "select * from usxrate where USXR01 in (select max(USXR01) from usxrate where USXR01<=to_number(to_char(sysdate, 'YYYYMMDD'))) "
'                          AdoRecordSet3.CursorLocation = adUseClient
'                          AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                          If AdoRecordSet3.RecordCount <> 0 Then
'                              oRate = AdoRecordSet3.Fields("USXR02").Value
'                          End If
'                          CheckOC3
''                          strSql = "select * from casefee where cf01='" & strTM01 & "' and cf02='" & strTM10 & "' and cf03='716' order by cf03 "
''                          AdoRecordSet3.CursorLocation = adUseClient
''                          AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
''                          If AdoRecordSet3.RecordCount <> 0 Then
''                              AdoRecordSet3.MoveFirst
''                              Do While Not AdoRecordSet3.EOF
''                                  Select Case AdoRecordSet3.Fields("cf03").Value
''                                  Case "716"
''                                      o10208 = AdoRecordSet3.Fields("cf08").Value
''                                      o10206 = AdoRecordSet3.Fields("cf06").Value
''                                  Case Else
''                                  End Select
''                                  AdoRecordSet3.MoveNext
''                              Loop
''                          End If
''                          CheckOC3
'                          'Add By Sindy 2011/5/30
'                          o10206 = 5000
'                          o10208 = 1500
'                          '2011/5/30 End
'                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                   "VALUES ('" & "10" & "','" & strCP09 & "','" & "07" & "','" & strUserNum & _
'                                   "','錢1','" & Format(o10208, "###,###,##0") & "')"
'                          cnnConnection.Execute strSql
'                          intFee = o10208 / oRate 'Modify By Sindy 2010/8/25 o10208 \ oRate
'                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                   "VALUES ('" & "10" & "','" & strCP09 & "','" & "07" & "','" & strUserNum & _
'                                   "','錢2','" & Format((intFee + IIf(o10208 Mod oRate <> 0, 1, 0)), "###,###,##0") & "')"
'                          cnnConnection.Execute strSql
'                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                   "VALUES ('" & "10" & "','" & strCP09 & "','" & "07" & "','" & strUserNum & _
'                                   "','錢3','" & Format((o10206 - o10208), "###,###,##0") & "')"
'                          cnnConnection.Execute strSql
'                          intFee = ((o10206 - o10208) / oRate) 'Modify By Sindy 2010/8/25 ((o10206 - o10208) \ oRate)
'                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                   "VALUES ('" & "10" & "','" & strCP09 & "','" & "07" & "','" & strUserNum & _
'                                   "','錢4','" & Format(intFee, "###,###,##0") & "')"
'                          cnnConnection.Execute strSql
'                          'edit by nickc 2006/08/09 外商 may  說應該是 1000
'                          'strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                   "VALUES ('" & "10" & "','" & strCP09 & "','" & "07" & "','" & strUserNum & _
'                                   "','錢5','" & Format(3500, "###,###,##0") & "')"
'                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                   "VALUES ('" & "10" & "','" & strCP09 & "','" & "07" & "','" & strUserNum & _
'                                   "','錢5','" & Format(1000, "###,###,##0") & "')"
'                          cnnConnection.Execute strSql
'                          'edit by nickc 2006/08/09 外商 may  說應該是 1000
'                          'strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                   "VALUES ('" & "10" & "','" & strCP09 & "','" & "07" & "','" & strUserNum & _
'                                   "','錢6','" & Format((3500 \ oRate + IIf(3500 Mod oRate <> 0, 1, 0)), "###,###,##0") & "')"
'                          intFee = 1000 / oRate 'Modify By Sindy 2010/8/25 1000 \ oRate
'                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                   "VALUES ('" & "10" & "','" & strCP09 & "','" & "07" & "','" & strUserNum & _
'                                   "','錢6','" & Format((intFee + IIf(1000 Mod oRate <> 0, 1, 0)), "###,###,##0") & "')"
'                          cnnConnection.Execute strSql
'                End If
'            Case Else
'         'Case Else
'                  'add by nickc 2007/02/15
'                  If Val(TXT1(6)) = 2 Then
'                        EndLetter "10", strCP09, "06", strUserNum
'                        ' 本所期限
'                        '2014/12/10 MODIFY BY SONIA 改通知法定期限
'                        'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "10" & "','" & strCP09 & "','" & "06" & "','" & strUserNum & _
'                                 "','本所期限','" & strNP08 & "')"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "10" & "','" & strCP09 & "','" & "06" & "','" & strUserNum & _
'                                 "','法定期限','" & strNP09 & "')"
'                        '2014/12/10 END
'                        cnnConnection.Execute strSql
'                        'add by nickc 2007/05/01 拆成兩個定稿，太長了
'                        EndLetter "10", strCP09, "08", strUserNum
'                        '2014/12/10 MODIFY BY SONIA 改通知法定期限
'                        'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "10" & "','" & strCP09 & "','" & "08" & "','" & strUserNum & _
'                                 "','本所期限','" & strNP08 & "')"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "10" & "','" & strCP09 & "','" & "08" & "','" & strUserNum & _
'                                 "','法定期限','" & strNP09 & "')"
'                        '2014/12/10 END
'                        cnnConnection.Execute strSql
'                  'add by nickc 2007/02/15
'                  ElseIf Val(TXT1(6)) = 3 Then
'                        'add by nickc 2007/02/09 加入 reminder 定稿
'                        If GetLetterLanguage(SystemNumber(strTM, 1), SystemNumber(strTM, 2), SystemNumber(strTM, 3), SystemNumber(strTM, 4)) = "2" Then
'                            tm(1) = SystemNumber(strTM, 1)
'                            tm(2) = SystemNumber(strTM, 2)
'                            tm(3) = SystemNumber(strTM, 3)
'                            tm(4) = SystemNumber(strTM, 4)
'                            If ClsPDReadTrademarkDatabase(tm(), 國外_FC) Then
'                              If Trim(tm(44)) <> "" Then
'                                    ExFaName = PUB_GetFAgentName(tm(44), 2)
'                                    ExFaNa = ""
'                                    ExFax = ""
'                                    ExTel = ""
'                                    ExFaDate = strNP09
'                                    ExSworDate = strNP08
'                                    strSql = "select * from fagent where fa01='" & Mid(tm(44) & "000000000", 1, 8) & "' and fa02='" & Mid(tm(44) & "000000000", 9, 1) & "' "
'                                    CheckOC3
'                                    AdoRecordSet3.CursorLocation = adUseClient
'                                    AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                                    If AdoRecordSet3.RecordCount <> 0 Then
'                                        ExFaNa = PUB_GetNationEngName(CheckStr(AdoRecordSet3.Fields("fa10").Value))
'                                        ExFax = CheckStr(AdoRecordSet3.Fields("fa14").Value) & IIf(CheckStr(AdoRecordSet3.Fields("fa15").Value) = "", "　　　　　　", "," & CheckStr(AdoRecordSet3.Fields("fa15").Value))
'                                        ExTel = CheckStr(AdoRecordSet3.Fields("fa12").Value) & IIf(CheckStr(AdoRecordSet3.Fields("fa13").Value) = "", "　　　　　　", "," & CheckStr(AdoRecordSet3.Fields("fa13").Value))
'                                    End If
'                                    CheckOC3
'                              ElseIf Trim(tm(23)) <> "" Then
'                                    ExFaName = GetCustomerName(tm(23), 1)
'                                    ExFaNa = ""
'                                    ExFax = ""
'                                    ExTel = ""
'                                    ExFaDate = strNP09
'                                    ExSworDate = strNP08
'                                    strSql = "select * from customer where cu01='" & Mid(tm(23) & "000000000", 1, 8) & "' and cu02='" & Mid(tm(23) & "000000000", 9, 1) & "' "
'                                    CheckOC3
'                                    AdoRecordSet3.CursorLocation = adUseClient
'                                    AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                                    If AdoRecordSet3.RecordCount <> 0 Then
'                                        ExFaNa = PUB_GetNationEngName(CheckStr(AdoRecordSet3.Fields("cu10").Value))
'                                        ExFax = CheckStr(AdoRecordSet3.Fields("cu18").Value) & IIf(CheckStr(AdoRecordSet3.Fields("cu19").Value) = "", "　　　　　　", "," & CheckStr(AdoRecordSet3.Fields("cu19").Value))
'                                        ExTel = CheckStr(AdoRecordSet3.Fields("cu16").Value) & IIf(CheckStr(AdoRecordSet3.Fields("cu17").Value) = "", "　　　　　　", "," & CheckStr(AdoRecordSet3.Fields("cu17").Value))
'                                    End If
'                                    CheckOC3
'                              End If
'                                EndLetter "10", Replace(strTM, "-", "") & "&" & Trim(strNP07), "97", strUserNum
'                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                         "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & "97" & "','" & strUserNum & _
'                                         "','例-代理人名稱','" & ChgSQL(ExFaName) & "')"
'                                cnnConnection.Execute strSql
'                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                         "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & "97" & "','" & strUserNum & _
'                                         "','例-代理人國籍','" & ExFaNa & "')"
'                                cnnConnection.Execute strSql
'                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                         "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & "97" & "','" & strUserNum & _
'                                         "','例-傳真','" & ExFax & "')"
'                                cnnConnection.Execute strSql
'                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                         "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & "97" & "','" & strUserNum & _
'                                         "','例-電話','" & ExTel & "')"
'                                cnnConnection.Execute strSql
'                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                         "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & "97" & "','" & strUserNum & _
'                                         "','例-法定','" & ExFaDate & "')"
'                                cnnConnection.Execute strSql
'                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                         "VALUES ('" & "10" & "','" & Replace(strTM, "-", "") & "&" & Trim(strNP07) & "','" & "97" & "','" & strUserNum & _
'                                         "','例-本所','" & ExSworDate & "')"
'                                cnnConnection.Execute strSql
'                            End If
'                        End If
'                End If
'            End Select
'         '2006/3/10 END
         End Select
      
      Case "CFT"
            'add by nickc 2007/02/15
            If Val(TXT1(6)) = 2 Then
                 '分下一程序
                 Select Case strNP07
                 '延展
                 Case "102"
                      Select Case strTM10
                      '香港
                      Case "013"
                          EndLetter "10", strCP09, "02", strUserNum
                          ' 本所期限
                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                   "VALUES ('" & "10" & "','" & strCP09 & "','" & "02" & "','" & strUserNum & _
                                   "','本所期限','" & strNP08 & "')"
                          cnnConnection.Execute strSql
                          ' 費用
                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                   "VALUES ('" & "10" & "','" & strCP09 & "','" & "02" & "','" & strUserNum & _
                                   "','費用','" & Me.TXT1(14).Text & "')"
                          cnnConnection.Execute strSql
                      '德國
                      Case "231"
                          EndLetter "10", strCP09, "03", strUserNum
                          ' 本所期限
                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                   "VALUES ('" & "10" & "','" & strCP09 & "','" & "03" & "','" & strUserNum & _
                                   "','本所期限','" & strNP08 & "')"
                          cnnConnection.Execute strSql
                          ' 費用
                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                   "VALUES ('" & "10" & "','" & strCP09 & "','" & "03" & "','" & strUserNum & _
                                   "','費用','" & Me.TXT1(14).Text & "')"
                          cnnConnection.Execute strSql
        'edit by nickc 2006/05/16 加錯地方了
                      'add by nickc 2006/03/24 美國，因為葉芳如改內容
        '              Case "101"
        '                  EndLetter "10", strCP09, "11", strUserNum
        '                  ' 本所期限
        '                  strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        '                           "VALUES ('" & "10" & "','" & strCP09 & "','" & "11" & "','" & strUserNum & _
        '                           "','本所期限','" & strNP08 & "')"
        '                  cnnConnection.Execute strSQL
        '                  ' 費用
        '                  strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        '                           "VALUES ('" & "10" & "','" & strCP09 & "','" & "11" & "','" & strUserNum & _
        '                           "','費用','" & Me.TXT1(14).Text & "')"
        '                  cnnConnection.Execute strSQL
                      'Add By Sindy 2019/3/4
                      '墨西哥
                      Case "104"
                          EndLetter "10", strCP09, "16", strUserNum
                          ' 本所期限
                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                   "VALUES ('" & "10" & "','" & strCP09 & "','" & "16" & "','" & strUserNum & _
                                   "','本所期限','" & strNP08 & "')"
                          cnnConnection.Execute strSql
                          ' 費用
                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                   "VALUES ('" & "10" & "','" & strCP09 & "','" & "16" & "','" & strUserNum & _
                                   "','費用','" & Me.TXT1(14).Text & "')"
                          cnnConnection.Execute strSql
                      '其他
                      Case Else
                          EndLetter "10", strCP09, "01", strUserNum
                          ' 本所期限
                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                   "VALUES ('" & "10" & "','" & strCP09 & "','" & "01" & "','" & strUserNum & _
                                   "','本所期限','" & strNP08 & "')"
                          cnnConnection.Execute strSql
                          ' 費用
                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                   "VALUES ('" & "10" & "','" & strCP09 & "','" & "01" & "','" & strUserNum & _
                                   "','費用','" & Me.TXT1(14).Text & "')"
                          cnnConnection.Execute strSql
                      End Select
                 '使用宣誓
                 Case "105"
                       Select Case strTM10
                       '柬埔寨
                       Case "046"
                             EndLetter "10", strCP09, "04", strUserNum
                             ' 本所期限
                             strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                      "VALUES ('" & "10" & "','" & strCP09 & "','" & "04" & "','" & strUserNum & _
                                      "','本所期限','" & strNP08 & "')"
                             cnnConnection.Execute strSql
                             ' 費用
                             strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                      "VALUES ('" & "10" & "','" & strCP09 & "','" & "04" & "','" & strUserNum & _
                                      "','費用','" & Me.TXT1(14).Text & "')"
                             cnnConnection.Execute strSql
                       '菲律賓
                       Case "030"
                             EndLetter "10", strCP09, "05", strUserNum
                             ' 本所期限
                             strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                      "VALUES ('" & "10" & "','" & strCP09 & "','" & "05" & "','" & strUserNum & _
                                      "','本所期限','" & strNP08 & "')"
                             cnnConnection.Execute strSql
                             ' 費用
                             strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                      "VALUES ('" & "10" & "','" & strCP09 & "','" & "05" & "','" & strUserNum & _
                                      "','費用','" & Me.TXT1(14).Text & "')"
                             cnnConnection.Execute strSql
                       Case Else
                       End Select
                 '刊登廣告
                 Case "702"
                      EndLetter "10", strCP09, "06", strUserNum
                       ' 本所期限
                       strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                "VALUES ('" & "10" & "','" & strCP09 & "','" & "06" & "','" & strUserNum & _
                                "','本所期限','" & strNP08 & "')"
                       cnnConnection.Execute strSql
                       ' 費用
                       strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                "VALUES ('" & "10" & "','" & strCP09 & "','" & "06" & "','" & strUserNum & _
                                "','費用','" & Me.TXT1(14).Text & "')"
                       cnnConnection.Execute strSql
                 Case Else
                 End Select
           End If
      Case Else
   End Select
           ' 清除定稿例外欄位檔原有資料
           'EndLetter "10", strCP09, "05", strUserNum
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Modify By Cheng 2002/12/29
'Private Sub PrintLetter(ByVal strNP07 As String, ByVal strTM01 As String, ByVal strTM10 As String, ByVal strData As String, ByVal strSysDate As String, ByVal strCP09 As String, ByVal strTM As String)
'Modify By Sindy 2012/11/19 區域變數改為須回傳值 +, ByRef ET01 As String, ByRef ET02 As String, ByRef ET03 As String, ByRef ET03_1 As String
Private Sub PrintLetter(ByVal strNP07 As String, ByVal strTM01 As String, ByVal strTM10 As String, ByVal strData As String, ByVal strSysDate As String, ByVal strCP09 As String, ByVal strTM As String, ByVal strNP08 As String, ByVal strNP09 As String _
                        , ByRef ET01 As String, ByRef ET02 As String, ByRef ET03 As String, ByRef ET03_1 As String, rsRS As ADODB.Recordset)
   'Dim ET01 As String, ET02 As String, ET03 As String, ET03_1 As String
   Dim stContent As String
   Dim bolEmail As Boolean, iCopy As Integer, iCopy1 As Integer, bolPlusPaper As Boolean
   
   m_blnPrintAddress = True
   ET01 = "10"
   
   ' 下一程序
   'Dim StrNP07 As Strubg
   ' 系統別
   'Dim StrTM01 As String
   ' 申請國家
   'Dim StrTM10 As String
   ' 專用期限止日
   'Dim StrDate As String
   ' 系統日
   'Dim StrSysDate As String
   ' 總收文號
   'Dim StrCP09 As String
   ' 下一程序本所期限
   'Dim StrNP08 As String
   ' 下一程序法定期限
   'Dim StrNP09 As String
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   'Modify By Cheng 2002/12/29
'   InsExpField strNP07, strTM01, strTM10, strDate, strSysDate, strCP09, strTM
   InsExpField strNP07, strTM01, strTM10, strDate, strSysDate, strCP09, strTM, strNP08, strNP09, rsRS
   Select Case strTM01
      'fct
      Case "FCT"
         ET02 = Replace(strTM, "-", "") & "&" & Trim(strNP07)
         iCopy = 0
         iCopy1 = 0
         '分下一程序
         Select Case strNP07
            '延展
            Case "102"
               Select Case GetLetterLanguage(SystemNumber(strTM, 1), SystemNumber(strTM, 2), SystemNumber(strTM, 3), SystemNumber(strTM, 4))
                  '中文
                  Case "1"
                      'add by nickc 2007/02/15
                      If Val(TXT1(6)) = 2 Then
                            'edit by nickc 2007/01/10 改成與英文同
                            'NowPrint strCP09, "10", "01", False, strUserNum, 0
                            ET03 = "01"
                      End If
                  '英文
                  Case "2"
                      'add by nickc 2007/02/15
                      If Val(TXT1(6)) = 2 Then
                           'edit by nick 2004/12/13
                           'NowPrint strCP09, "10", IIf(m_blnSingleCase = True, "02", "03"), False, strUserNum, 0
                           If m_blnSingleCase = True Then
                              ET03 = "02"
                           Else
                              ET03 = "03"
                           End If
                      'add by nickc 2007/02/15
                      ElseIf Val(TXT1(6)) = 3 Then
                           'add by nickc 2007/02/13 加入傳真定稿
                           If m_blnSingleCase = True Then
                              ET03 = "98"
                           Else
                              ET03 = "99"
                           End If
                           iCopy = 1
                       End If
                  'Modify By Sindy 2025/6/13 日文組已不使用此作業定稿,改用frm030404
'                  '日文
'                  Case "3"
'                      'add by nickc 2007/02/15
'                      If Val(TXT1(6)) = 2 Then
'                            'edit by nickc 2007/01/10 改成與英文同
'                            'NowPrint strCP09, "10", "04", False, strUserNum, 0
'                            ET03 = "04"
'                      End If
               End Select
            '2006/3/10 ADD BY SONIA 第二期註冊費
            'edit by nickc 2006/06/01
            'Case Else
'             Case "716"
'                'add by nickc 2007/02/15
'                If Val(TXT1(6)) = 2 Then
'                        'add by nickc 2006/08/18 加入第二期接洽結案單 阿蓮的請作單
'                        'edit by nickc 2006/09/05 改整批
'                        'g_PrtForm001.PrintForm GetNp22(SystemNumber(strTM, 1), SystemNumber(strTM, 2), SystemNumber(strTM, 3), SystemNumber(strTM, 4)), SystemNumber(strTM, 1), SystemNumber(strTM, 2), SystemNumber(strTM, 3), SystemNumber(strTM, 4)
'                        pub_AddressListSN = pub_AddressListSN + 1
'                        PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, GetNp22(SystemNumber(strTM, 1), SystemNumber(strTM, 2), SystemNumber(strTM, 3), SystemNumber(strTM, 4)), SystemNumber(strTM, 1), SystemNumber(strTM, 2), SystemNumber(strTM, 3), SystemNumber(strTM, 4)
'                End If
'                'add by nickc 2006/06/01  加入日文定稿
'                Select Case GetLetterLanguage(SystemNumber(strTM, 1), SystemNumber(strTM, 2), SystemNumber(strTM, 3), SystemNumber(strTM, 4))
'                Case "3"
'                      'add by nickc 2007/02/15
'                      If Val(TXT1(6)) = 2 Then
'                           ET02 = strCP09
'                           ET03 = "07"
'                           '2008/8/13 因為紙張不同所拆成兩個定稿
'                           ET03_1 = "09"
'                      End If
'                Case Else
'                      'add by nickc 2007/02/15
'                      If Val(TXT1(6)) = 2 Then
'                           ET02 = strCP09
'                           ET03 = "06"
'                           'add by nickc 2007/05/01 拆成兩個定稿，太長了
'                           ET03_1 = "08"
'                           iCopy1 = 1
'                       'add by nickc 2007/02/15
'                      ElseIf Val(TXT1(6)) = 3 Then
'                               'add by nickc 2007/02/13 加入英文傳真定稿
'                            If GetLetterLanguage(SystemNumber(strTM, 1), SystemNumber(strTM, 2), SystemNumber(strTM, 3), SystemNumber(strTM, 4)) = "2" Then
'                                 ET03 = "97"
'                                 iCopy = 1
'                            End If
'                       End If
'                End Select
'            '2006/3/10 END
         End Select
         
         If ET03 <> "" Then
            'Add by Morgan 2008/6/13
            bolEmail = PUB_GetEMailFlag(Replace(strTM, "-", ""), Trim(strNP07), , bolPlusPaper)
            
            'Modify By Sindy 2021/3/4 防疫,是否停止郵務
            Dim strNA86 As String
            Call GetPrjPeopleNum6(strTM, "NA86", strNA86)
            '2021/3/4 END
            
            'Modify By Sindy 2021/3/4 + Or strNA86 = "Y":針對無收受郵件之國家改以Email通知
            'Modify By Sindy 2024/4/12 以上定稿原只有設定Email通知之定稿會自動存入FCT_workflow，
            '                          請比照其他定稿, 無論是否設定以Email通知, 都要產生定稿存於FCT_workflow
'            If bolEmail Or strNA86 = "Y" Then
            '2024/4/12 END
               'Add by Morgan 2009/10/20 + 判斷是否EMail同時寄紙本
               'Modify By Sindy 2021/3/4 + Or strNA86 = "Y":針對無收受郵件之國家改以Email通知，請管控定稿只需列印一份即可
               If Not bolPlusPaper Or strNA86 = "Y" Then
                  iCopy = 1
                  iCopy1 = 1
               End If
               'end 2009/10/20
               m_blnPrintAddress = False
               If ET03_1 <> "" Then
                  NowPrint ET02, ET01, ET03, False, strUserNum, , , , , iCopy
                  NowPrint ET02, ET01, ET03_1, False, strUserNum, , , , , iCopy1
                  NowPrint ET02, ET01, ET03, False, strUserNum, , , True, stContent, , , , True
                  NowPrint ET02, ET01, ET03_1, False, strUserNum, , stContent, , , , , True, True
               Else
                  NowPrint ET02, ET01, ET03, False, strUserNum, , , , , iCopy, , True, True
               End If
               'MsgBox "電子檔已存於 [ " & FCTeFilePath & " ]！"
               'Add By Sindy 2017/11/20
               'Modify By Sindy 2019/1/25
               'LblNote.Caption = "電子檔已存於 [ " & FCTeFilePath & " ]！"
               LblNote.Caption = "電子檔已存於 [ " & m_strFilePath & " ]！"
               '2019/1/25 END
               DoEvents
'            Else
'            'end 2008/6/13
'               m_blnPrintAddress = True
'               NowPrint ET02, ET01, ET03, False, strUserNum, , , , , iCopy
'               If ET03_1 <> "" Then
'                  NowPrint ET02, ET01, ET03_1, False, strUserNum, , , , , iCopy1
'               End If
'            End If
            'Add By Sindy 2017/11/20
            LblNote2.Caption = ET02 & "-" & ET01 & "-" & ET03
            DoEvents
         End If
         
      'cft
      Case "CFT"
        'add by nickc 2007/02/15
        If Val(TXT1(6)) = 2 Then
                 '分下一程序
                 Select Case strNP07
                 '延展
                 Case "102"
                      Select Case strTM10
                      '香港
                      Case "013"
                          ET03 = "02"
                          'NowPrint strCP09, "10", "02", False, strUserNum, 0
                      '德國
                      Case "231"
                          ET03 = "03"
                          'NowPrint strCP09, "10", "03", False, strUserNum, 0
        'edit by nickc 2006/05/16 加錯地方了
        '              'add by nickc 2006/03/24 美國，因為葉芳如加內容
        '              Case "101"
        '                 ET03="11"
        '                 NowPrint strCP09, "10", "11", False, strUserNum, 0
                      'Add By Sindy 2019/3/4
                      '墨西哥
                      Case "104"
                          ET03 = "16"
                      '其他
                      Case Else
                          ET03 = "01"
                          'NowPrint strCP09, "10", "01", False, strUserNum, 0
                      End Select
                 '使用宣誓
                 Case "105"
                       Select Case strTM10
                       '柬埔寨
                       Case "046"
                             ET03 = "04"
                             'NowPrint strCP09, "10", "04", False, strUserNum, 0
                       '菲律賓
                       Case "030"
                             ET03 = "05"
                             'NowPrint strCP09, "10", "05", False, strUserNum, 0
                       Case Else
                       End Select
                 '刊登廣告
                 Case "702"
                      ET03 = "06"
                      'NowPrint strCP09, "10", "06", False, strUserNum, 0
                 Case Else
                 End Select
                 If ET03 <> "" Then
                     NowPrint strCP09, ET01, ET03, False, strUserNum, 0
                     'Add By Sindy 2017/11/20
                     LblNote2.Caption = strCP09 & "-" & ET01 & "-" & ET03
                     DoEvents
                 End If
           End If
      Case Else
   End Select
End Sub

'Add By Sindy 2013/1/4 要產生來函通知進度
'Modify By Sindy 2015/4/30 +, ByVal strNP22 As String
Private Sub FCTInsertCP(ByVal strNP07 As String, ByVal strCP09 As String, ByVal strTM As String, ByVal strNP22 As String)
   Dim strNewCP09 As String
'   Dim strCP10 As String
   
   m_TM01 = SystemNumber(strTM, 1)
   m_TM02 = SystemNumber(strTM, 2)
   m_TM03 = SystemNumber(strTM, 3)
   m_TM04 = SystemNumber(strTM, 4)
   'If m_bolInsCP = True And m_TM01 = "FCT" And (strNP07 = "102" Or strNP07 = "716") And (TXT1(6) = "2" Or TXT1(6) = "3") Then
      'Modified by Lydia 2016/12/22 本所管控C類進度自2017/01/01起改用D類收文
      'strNewCP09 = AutoNo("C", 6)
      If strSrvDate(1) >= 本所D類收文啟用日 Then
         strNewCP09 = AutoNo("D", 6)
      Else
         strNewCP09 = AutoNo("C", 6)
      End If
      'end 2016/12/22
      
      'Modify By Sindy 2015/4/30 +,CP30
      strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP27,CP30) " & _
      "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strSrvDate(1) & "," & _
              "'" & strNewCP09 & "','" & m_strUpdCP10 & "','" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & strUserNum & "'," & _
              "'N','N','N'," & _
              "'" & strCP09 & "'," & strSrvDate(1) & "," & CNULL(strNP22) & ")"
      cnnConnection.Execute strSql
   'End If
End Sub

'期限管制表-未收文
Sub Process()
Dim intRow As Integer
Dim strST15 As String 'Add By Sindy 2011/1/12
Dim strEmp As String 'Add By Sindy 2015/5/7

Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R030403_1 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R030403_2 WHERE ID='" & strUserNum & "' "
'cnnConnection.Execute "DELETE FROM R030403_3 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL6 = ""
If Len(TXT1(0)) <> 0 Then
   strSQL1 = strSQL1 + " AND NP02 IN (" & SQLGrpStr(TXT1(0), 2) & ") "
   strSQL2 = strSQL2 + " AND NP02 IN (" & SQLGrpStr(TXT1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & TXT1(0)  'Add By Sindy 2010/10/22
End If
'add by nick 2004/08/04 加入是否含延展及是否含註冊費
'2006/1/9 CANCEL BY SONIA 移至下方,否則選Y時只抓102的資料
'If Text2(0).Text = "Y" Then
'    txt1(5).Text = txt1(5).Text & ",102"
'End If
'2006/1/9 END
'2006/1/9 CANCEL BY SONIA
'If Text2(1).Text = "Y" Then
'    txt1(5).Text = txt1(5).Text & ",716"
'End If
'2006/1/9 END
StrSQL6 = ""
If Len(TXT1(5)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & TXT1(5)  'Add By Sindy 2010/10/22
   'edit by nickc 2007/04/16
   'StrSQL6 = " AND ("
   strSQL1 = strSQL1 & " AND ("
   strSQL2 = strSQL2 & " AND ("
   If Len(TXT1(5)) <> 0 Then
      strTemp1 = ""
      strTemp1 = Split(Replace(TXT1(5), ",,", ""), ",")
      For i = 0 To UBound(strTemp1)
         'edit by nickc 2007/04/16 加入可辦期限條件沒專用期限的使用宣誓不用
         If Val(strTemp1(i)) = 105 And Option1(3).Value = True Then
            'edit by nickc 2007/07/11 菲律賓 無專用期的要抓
            'strSQL1 = strSQL1 + " (tm21 is not null and NP07=" & Val(strTemp1(i)) & ") OR "
            'modify by sonia 2014/11/4 CFT-013059波多黎各112無專用期也要抓
            strSQL1 = strSQL1 + " (tm21 is not null and NP07=" & Val(strTemp1(i)) & " and tm10<>'030' and tm10<>'112') OR (NP07=" & Val(strTemp1(i)) & " and tm10 in ('030','112') ) or "
            strSQL2 = strSQL2 + " NP07=" & Val(strTemp1(i)) & " OR "
         Else
            'edit by nickc 2007/04/16
            'StrSQL6 = StrSQL6 + " NP07=" & Val(strTemp1(i)) & " OR "
            strSQL1 = strSQL1 + " NP07=" & Val(strTemp1(i)) & " OR "
            strSQL2 = strSQL2 + " NP07=" & Val(strTemp1(i)) & " OR "
         End If
      Next i
      '2006/1/9 ADD BY SONIA
      If Text2(0).Text = "Y" Then
         'edit by nickc 2007/04/16
         'StrSQL6 = StrSQL6 + " NP07=102 OR "
         strSQL1 = strSQL1 + " NP07=102 OR "
         strSQL2 = strSQL2 + " NP07=102 OR "
      End If
      '2006/1/9 END
      'edit by nickc 2007/04/16
      'StrSQL6 = StrSQL6 + " NP07=0) "
      strSQL1 = strSQL1 + " NP07=0) "
      strSQL2 = strSQL2 + " NP07=0) "
   End If
'Add By Cheng 2004/03/11
'若未輸入案件性質, 催審(305)不印
Else
   StrSQL6 = StrSQL6 + " AND NP07<>'305' "
   'edit by nickc 2007/04/16
   If Option1(3).Value = True Then
      strSQL1 = strSQL1 + " and decode(np07,'105',tm21,'1') is not null  "
   End If
   
   '2006/1/9 ADD BY SONIA
   If Text2(0).Text <> "Y" Then
      StrSQL6 = StrSQL6 + " AND NP07<>'102' "
   Else
      pub_QL05 = pub_QL05 & ";" & Left(Label1(12), 6) & "含" 'Add By Sindy 2010/10/22
   End If
   '2006/1/9 END
'End
End If
'Add By Cheng 2004/03/11
'下一程序的案件性質為收達(997), 提申(998)的不印
StrSQL6 = StrSQL6 + " AND (NP07<>'997' And NP07<>'998') "
'End
StrSQL6 = StrSQL6 + " AND (NP06 IS NULL OR NP06='') "
'本所期限
If Option1(0).Value = True Then
    If Len(Trim(TXT1(1))) <> 0 Then
      StrSQL6 = StrSQL6 + " AND NP08>=" & Val(ChangeTStringToWString(TXT1(1))) & " "
    End If
    If Len(Trim(TXT1(2))) <> 0 Then
      StrSQL6 = StrSQL6 + " AND NP08<=" & Val(ChangeTStringToWString(TXT1(2))) & " "
    End If
    If Len(Trim(TXT1(1))) <> 0 Or Len(Trim(TXT1(2))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & TXT1(1) & "-" & TXT1(2)  'Add By Sindy 2010/10/22
    End If
'Modify By Cheng 2003/01/08
'法定期限
ElseIf Me.Option1(1).Value Then
    If Len(Trim(TXT1(3))) <> 0 Then
    StrSQL6 = StrSQL6 + " AND NP09>=" & Val(ChangeTStringToWString(TXT1(3))) & " "
    End If
    If Len(Trim(TXT1(4))) <> 0 Then
      StrSQL6 = StrSQL6 + " AND NP09<=" & Val(ChangeTStringToWString(TXT1(4))) & " "
    End If
    If Len(Trim(TXT1(3))) <> 0 Or Len(Trim(TXT1(4))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & TXT1(3) & "-" & TXT1(4)  'Add By Sindy 2010/10/22
    End If
'Add By Cheng 2003/01/08
'本所案號
'edit by nickc 2007/03/05
'Else
ElseIf Option1(2).Value Then
    StrSQL6 = StrSQL6 + " AND NP02='" & Me.text1(0).Text & "' And NP03='" & Me.text1(1).Text & "' And NP04='" & Me.text1(2).Text & "' And NP05='" & Me.text1(3).Text & "' "
    pub_QL05 = pub_QL05 & ";" & Option1(2).Caption & text1(0) & "-" & text1(1) & "-" & text1(2) & "-" & text1(3) 'Add By Sindy 2010/10/22
'add by nickc 2007/03/05 加入可辦，但是只有延展跟使用宣誓
ElseIf Option1(3).Value Then
'to_char(add_months(to_date(ChangeTStringToWString(TXT1(18)),'YYYYMMDD'),na15 * -1),'YYYYMMDD')
    'add by nickc 2007/05/01 外商阿蓮跟秀玲研究後，延展(可辦期限 = 法定期限 - 國家檔定義) ，使用宣誓(可辦期限 = 法定期限- 1年)
    If Len(Trim(TXT1(18))) <> 0 And Len(Trim(TXT1(19))) <> 0 Then
        'edot by nickc 2007/07/11 菲律賓無專用期105 要抓 18個月
        'StrSQL6 = StrSQL6 + " AND ((to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and np07=102) or (to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and np07=105)) "
        'Modify By Sindy 2013/11/15 菲律賓無專用期105 要抓 14個月
        'modify by sonia 2014/11/4 CFT-013059波多黎各112要抓6個月
        'strSQL1 = strSQL1 + " AND ((to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(18)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(19)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and np07=102) or (to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(18)) & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(19)) & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and np07=105 and tm21 is not null ) or (to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(18)) & " ,'YYYYMMDD'),14 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(19)) & " ,'YYYYMMDD'),14 ),'YYYYMMDD') and np07=105 and tm21 is null and tm10='030')) "
        'modify by sonia 2021/6/2 CFT-020426墨西哥104使用宣誓(可辦期限 = 法定期限-3個月)
        'strSQL1 = strSQL1 + " AND ((to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(18)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(19)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and np07=102) or (to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(18)) & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(19)) & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and np07=105 and tm21 is not null ) or (to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(18)) & " ,'YYYYMMDD'),14 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(19)) & " ,'YYYYMMDD'),14 ),'YYYYMMDD') and np07=105 and tm21 is null and tm10='030') " & _
                            "   or (to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(18)) & " ,'YYYYMMDD'),6 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(19)) & " ,'YYYYMMDD'),6 ),'YYYYMMDD') and np07=105 and tm21 is null and tm10='112')) "
        strSQL1 = strSQL1 + " AND ((to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and np07=102) " & _
                            " or (to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and np07=105 and tm21 is not null and tm10<>'104') " & _
                            " or (to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),3 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),3 ),'YYYYMMDD') and np07=105 and tm21 is not null and tm10='104') " & _
                            " or (to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),14 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),14 ),'YYYYMMDD') and np07=105 and tm21 is null and tm10='030') " & _
                            " or (to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),6 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),6 ),'YYYYMMDD') and np07=105 and tm21 is null and tm10='112')) "
        'end 2014/11/4
        strSQL2 = strSQL2 + " AND ((to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and np07=102) or (to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and np07=105)) "
    Else
        If Len(Trim(TXT1(18))) <> 0 Then
    'edit by nickc 2007/05/01 外商阿蓮跟秀玲研究後，延展(可辦期限 = 法定期限 - 國家檔定義) ，使用宣誓(可辦期限 = 法定期限- 1年)
    '    StrSQL6 = StrSQL6 + " AND to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and np07 in (102,105) "
        'edot by nickc 2007/07/11 菲律賓無專用期105 要抓 18個月
        'StrSQL6 = StrSQL6 + " AND ((to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and np07=102) or (to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and np07=105)) "
        'Modify By Sindy 2013/11/15 菲律賓無專用期105 要抓 14個月
        'modify by sonia 2014/11/4 CFT-013059波多黎各112要抓6個月
        'strSQL1 = strSQL1 + " AND ((to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(18)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and np07=102) or (to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(18)) & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and np07=105 and tm21 is not null) or (to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(18)) & " ,'YYYYMMDD'),14 ),'YYYYMMDD') and np07=105 and tm21 is null and tm10='030' )  ) "
        strSQL1 = strSQL1 + " AND ((to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and np07=102) or (to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and np07=105 and tm21 is not null) or (to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),14 ),'YYYYMMDD') and np07=105 and tm21 is null and tm10='030' ) or (to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),6 ),'YYYYMMDD') and np07=105 and tm21 is null and tm10='112' ) ) "
        'end 2014/11/4
        strSQL2 = strSQL2 + " AND ((to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and np07=102) or (to_char(NP09)>=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(18)) & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and np07=105)) "
        End If
        If Len(Trim(TXT1(19))) <> 0 Then
           'edit by nickc 2007/05/01 外商阿蓮跟秀玲研究後，延展(可辦期限 = 法定期限 - 國家檔定義) ，使用宣誓(可辦期限 = 法定期限- 1年)
           'StrSQL6 = StrSQL6 + " AND to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and np07 in (102,105)  "
           'edot by nickc 2007/07/11 菲律賓無專用期105 要抓 18個月
           'StrSQL6 = StrSQL6 + " AND ((to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and np07=102) or (to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and np07=105))  "
           'Modify By Sindy 2013/11/15 菲律賓無專用期105 要抓 14個月
           'modify by sonia 2014/11/4 CFT-013059波多黎各112要抓6個月
           'strSQL1 = strSQL1 + " AND ((to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(19)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and np07=102) or (to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(19)) & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and np07=105 and tm21 is not null ) or (to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(19)) & " ,'YYYYMMDD'),14 ),'YYYYMMDD') and np07=105 and tm21 is null and tm10='030'))  "
           strSQL1 = strSQL1 + " AND ((to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and np07=102) or (to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and np07=105 and tm21 is not null ) or (to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),14 ),'YYYYMMDD') and np07=105 and tm21 is null and tm10='030') or (to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),6 ),'YYYYMMDD') and np07=105 and tm21 is null and tm10='112'))  "
           'end 2014/11/4
           strSQL2 = strSQL2 + " AND ((to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and np07=102) or (to_char(NP09)<=to_char(add_months(to_date(" & ChangeTStringToWString(TXT1(19)) & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and np07=105))  "
        End If
    End If
    If Len(Trim(TXT1(18))) <> 0 Or Len(Trim(TXT1(19))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Option1(3).Caption & TXT1(18) & "-" & TXT1(19)  'Add By Sindy 2010/10/22
    End If
End If
' Make By Sindy 98/02/27 下面迴圈才判斷
If Len(TXT1(8)) <> 0 Then
'    'Modify By Cheng 2003/11/05
''    StrSQL6 = StrSQL6 + " AND NP10='" & TXT1(8) & "' "
'    StrSQL6 = StrSQL6 + " AND Decode(NP02, 'FCT', Nvl(N2.NA55, N3.NA55), NP10)='" & TXT1(8) & "' "
   pub_QL05 = pub_QL05 & ";" & Label1(5) & TXT1(8) & LBL1(0)  'Add By Sindy 2010/10/22
End If
If Len(TXT1(9)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND CP14='" & TXT1(9) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(3) & TXT1(9) & LBL1(1) 'Add By Sindy 2010/10/22
End If
If Len(Trim(TXT1(10))) <> 0 And Len(Trim(TXT1(11))) <> 0 Then
    strSQL1 = strSQL1 & " AND (TM23>='" & GetNewFagent(TXT1(10)) & "' AND TM23<='" & GetNewFagent(TXT1(11)) & "') "
    strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(TXT1(10)) & "' AND SP08<='" & GetNewFagent(TXT1(11)) & "') OR (SP58<='" & GetNewFagent(TXT1(10)) & "' AND SP58<='" & GetNewFagent(TXT1(11)) & "') OR (SP59>='" & GetNewFagent(TXT1(10)) & "' AND SP59<='" & GetNewFagent(TXT1(11)) & "')) "
Else
    If Len(Trim(TXT1(10))) <> 0 And Len(Trim(TXT1(11))) = 0 Then
        strSQL1 = strSQL1 & " AND (TM23>='" & GetNewFagent(TXT1(10)) & "' ) "
        strSQL2 = strSQL2 + " AND (SP08>='" & GetNewFagent(TXT1(10)) & "' OR SP58>='" & GetNewFagent(TXT1(10)) & "' OR SP59>='" & GetNewFagent(TXT1(10)) & "') "
    Else
        If Len(Trim(TXT1(10))) = 0 And Len(Trim(TXT1(11))) <> 0 Then
            strSQL1 = strSQL1 & " AND (TM23<='" & GetNewFagent(TXT1(11)) & "') "
            strSQL2 = strSQL2 + " AND (SP08<='" & GetNewFagent(TXT1(11)) & "' OR SP58<='" & GetNewFagent(TXT1(11)) & "' OR SP59<='" & GetNewFagent(TXT1(11)) & "') "
        End If
    End If
End If
If Len(Trim(TXT1(10))) <> 0 Or Len(Trim(TXT1(11))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(7) & TXT1(10) & "-" & TXT1(11) 'Add By Sindy 2010/10/22
End If
If Len(Trim(TXT1(12))) <> 0 And Len(Trim(TXT1(13))) <> 0 Then
    strSQL1 = strSQL1 + " AND (TM44>='" & GetNewFagent(TXT1(12)) & "' AND TM44<='" & GetNewFagent(TXT1(13)) & "') "
    strSQL2 = strSQL2 + " AND (SP26>='" & GetNewFagent(TXT1(12)) & "' AND SP26<='" & GetNewFagent(TXT1(13)) & "') "
Else
    If Len(Trim(TXT1(12))) <> 0 And Len(Trim(TXT1(13))) = 0 Then
        strSQL1 = strSQL1 + " AND (TM44>='" & GetNewFagent(TXT1(12)) & "' ) "
        strSQL2 = strSQL2 + " AND (SP26>='" & GetNewFagent(TXT1(12)) & "' ) "
    Else
        If Len(Trim(TXT1(12))) = 0 And Len(Trim(TXT1(13))) <> 0 Then
            strSQL1 = strSQL1 + " AND (TM44<='" & GetNewFagent(TXT1(13)) & "' ) "
            strSQL2 = strSQL2 + " AND (SP26<='" & GetNewFagent(TXT1(13)) & "' ) "
        End If
    End If
End If
If Len(Trim(TXT1(12))) <> 0 Or Len(Trim(TXT1(13))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(8) & TXT1(12) & "-" & TXT1(13) 'Add By Sindy 2010/10/22
End If
'add by nick 2005/02/15 加入申請國家
If Len(Trim(TXT1(15))) <> 0 Then
   strSQL1 = strSQL1 & " and tm10>='" & TXT1(15) & "' "
   strSQL2 = strSQL2 & " and sp09>='" & TXT1(15) & "' "
End If
If Len(Trim(TXT1(16))) <> 0 Then
   strSQL1 = strSQL1 & " and tm10<='" & TXT1(16) & "' "
   strSQL2 = strSQL2 & " and sp09<='" & TXT1(16) & "' "
End If
'add end
If Len(Trim(TXT1(15))) <> 0 Or Len(Trim(TXT1(16))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(13) & TXT1(15) & "-" & TXT1(16) 'Add By Sindy 2010/10/22
End If
strSQL1 = strSQL1 + " and (tm29 is null or tm29 <> 'Y' ) "
strSQL2 = strSQL2 + " AND (SP15 IS NULL OR SP15 <> 'Y' ) "

'Add By Sindy 2023/9/20 報表類別(1、4)組別
If Len(Trim(TXT1(7))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(6) & TXT1(7) & Label1(10)
   If TXT1(7) = "1" Then '1.英文組
      strSQL1 = strSQL1 + " AND NOT(NP02='FCT' AND substr(FA1.FA10,1,3)='011') "
      strSQL2 = strSQL2 + " AND NOT(NP02='FCT' AND substr(FA10,1,3)='011') "
   Else '2.日文組
      strSQL1 = strSQL1 + " AND (NP02='FCT' AND substr(FA1.FA10,1,3)='011') "
      strSQL2 = strSQL2 + " AND (NP02='FCT' AND substr(FA10,1,3)='011') "
   End If
End If
'2023/9/20 END

'add by nickc 2006/05/30
If (InStr(1, TXT1(0), "CFT") <> 0 Or InStr(1, TXT1(0), "CFC") <> 0) Then
    If TXT1(17) = "1" Then
        ' Modify By Sindy 2011/1/12 下面迴圈才判斷
        'StrSQL6 = StrSQL6 + " AND substr(s1.ST15,1,1)<>'S' "
        pub_QL05 = pub_QL05 & ";" & Label1(9) & TXT1(17) & Label1(14) 'Add By Sindy 2010/10/22
    End If
End If
'Add By Sindy 2023/9/20 4=未催延展檢核表
If Val(TXT1(6)) = 4 Then
'   If TXT1(7) = "1" Then '1.英文組
      '30天之內是否已有催延展的進度
      StrSQL6 = StrSQL6 + " AND NOT exists(select * from caseprogress where cp01=np02 and cp02=np03 and cp03=np04 and cp04=np05" & _
                " and cp10 in('1717','1722') AND CP05>=" & CompWorkDay(30, strSrvDate(1), 1) & ") "
'   Else '2.日文組
'      StrSQL6 = StrSQL6 + " AND NOT exists(select * from caseprogress where cp01=np02 and cp02=np03 and cp03=np04 and cp04=np05" & _
'                " and cp10 in('1717','1722') AND CP05>=" & CompDate(0, -2, strSrvDate(1)) & ") "
'   End If
End If
'2023/9/20 END

'add by nickc 2006/05/30 延展和第二期專用權須存在  FCT  才做
If InStr(1, TXT1(5), "716") <> 0 Or InStr(1, TXT1(5), "102") <> 0 Then
    'edit by nickc 2007/07/13 加入CFT
    'strSQL1 = strSQL1 & " and np02||decode(np07,716,tm17,102,tm17,'Y')='FCTY' "
    strSQL1 = strSQL1 & " and decode(np02,'CFT','FCTY',np02||decode(np07,716,tm17,102,tm17,'Y'))='FCTY' "
End If
CheckOC
'Modify By Cheng 2003/11/05
'strSQL = "SELECT S1.ST03,NP10,NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(TM15,TM12),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NP15,S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(N1.NA03,N1.NA04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),NVL(N2.NA03,N2.NA04),TM22,np07,NP01,CP27 FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,FAGENT,CUSTOMER WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND FA10=N2.NA01(+) AND TM10=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA01(+) AND SUBSTR(TM44,9,1)=FA02(+) " & strSQL1 & StrSQL6
'strSQL = strSQL + " union all select S1.ST03,NP10,NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),'',SP11,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NP15,S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(N1.NA03,N1.NA04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),NVL(N2.NA03,N2.NA04),SP21,NP07,NP01,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,FAGENT,CUSTOMER WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND FA10=N2.NA01(+) AND SP09=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) " & strSQL2 & StrSQL6
'edit by nick 2004/12/07   延展抓延展代理人
'StrSql = "SELECT S1.ST03, Decode(NP02, 'FCT', Nvl(N2.NA55, N3.NA55), NP10),NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(TM15,TM12),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NP15,S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(N1.NA03,N1.NA04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),NVL(N2.NA03,N2.NA04),TM22,np07,NP01,CP27 FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,FAGENT,CUSTOMER, Nation N3 WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND FA10=N2.NA01(+) AND TM10=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA01(+) AND SUBSTR(TM44,9,1)=FA02(+) And CU10=N3.NA01(+) " & strSQL1 & StrSQL6
'StrSql = StrSql + " union all select S1.ST03, Decode(NP02, 'FCT', Nvl(N2.NA55, N3.NA55), NP10),NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),'',SP11,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NP15,S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(N1.NA03,N1.NA04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),NVL(N2.NA03,N2.NA04),SP21,NP07,NP01,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,FAGENT,CUSTOMER, Nation N3 WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND FA10=N2.NA01(+) AND SP09=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) And CU10=N3.NA01(+) " & strSQL2 & StrSQL6
' Modify By Sindy 98/02/27
'strSQL = "SELECT S1.ST03, Decode(NP02, 'FCT', Nvl(N2.NA55, N3.NA55), NP10),NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(TM15,TM12),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NP15,S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(N1.NA03,N1.NA04),decode(np07,'102',decode(tm33,null,NVL(FA1.FA04,NVL(FA1.FA05||FA1.FA63||FA1.FA64||FA1.FA65,FA1.FA06)),NVL(FA2.FA04,NVL(FA2.FA05||FA2.FA63||FA2.FA64||FA2.FA65,FA2.FA06))),NVL(FA1.FA04,NVL(FA1.FA05||FA1.FA63||FA1.FA64||FA1.FA65,FA1.FA06))),decode(np07,'102',decode(tm33,null,NVL(N2.NA03,N2.NA04),NVL(N4.NA03,N4.NA04)),NVL(N2.NA03,N2.NA04)),TM22,np07,NP01,CP27 FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,FAGENT FA1,FAGENT FA2,CUSTOMER, Nation N3,nation N4 " & _
'              " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND Fa2.FA10=N4.na01(+) and FA1.FA10=N2.NA01(+) AND TM10=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA1.FA01(+) AND SUBSTR(TM44,9,1)=FA1.FA02(+) And CU10=N3.NA01(+) and SUBSTR(TM33,1,8)=FA2.FA01(+) AND SUBSTR(TM33,9,1)=FA2.FA02(+)" & strSQL1 & StrSQL6
'strSQL = strSQL + " union all select S1.ST03, Decode(NP02, 'FCT', Nvl(N2.NA55, N3.NA55), NP10),NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),'',SP11,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NP15,S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(N1.NA03,N1.NA04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),NVL(N2.NA03,N2.NA04),SP21,NP07,NP01,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,FAGENT,CUSTOMER, Nation N3 WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND FA10=N2.NA01(+) AND SP09=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) And CU10=N3.NA01(+) " & strSQL2 & StrSQL6
'2009/12/10 MODIFY BY SONIA 延展抓的延展代理人若為X編號則抓不到,FCT-016425
'strSQL = "SELECT S1.ST03, NP10,NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(TM15,TM12),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NP15,S2.ST02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(N1.NA03,N1.NA04),decode(np07,'102',decode(tm33,null,NVL(FA1.FA04,NVL(FA1.FA05||FA1.FA63||FA1.FA64||FA1.FA65,FA1.FA06)),NVL(FA2.FA04,NVL(FA2.FA05||FA2.FA63||FA2.FA64||FA2.FA65,FA2.FA06))),NVL(FA1.FA04,NVL(FA1.FA05||FA1.FA63||FA1.FA64||FA1.FA65,FA1.FA06))),decode(np07,'102',decode(tm33,null,NVL(N2.NA03,N2.NA04),NVL(N4.NA03,N4.NA04)),NVL(N2.NA03,N2.NA04)),TM22,np07,NP01,CP27,TM01,TM02,TM03,TM04 FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,FAGENT FA1,FAGENT FA2,CUSTOMER, Nation N3,nation N4 " & _
              " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND Fa2.FA10=N4.na01(+) and FA1.FA10=N2.NA01(+) AND TM10=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA1.FA01(+) AND SUBSTR(TM44,9,1)=FA1.FA02(+) And CU10=N3.NA01(+) and SUBSTR(TM33,1,8)=FA2.FA01(+) AND SUBSTR(TM33,9,1)=FA2.FA02(+)" & strSQL1 & StrSQL6
'Modify By Sindy 2015/5/6 S2.ST02==>N1.NA69
If TXT1(0) = "CFT" Then
   'Modified by Lydia 2016/04/12 CFT案延展和使用宣誓改抓NA69
   'strEmp = "N1.NA69" '國家檔的CFT承辦人
   strEmp = "DECODE(NP07,'102',N1.NA69,'105',N1.NA69,S2.ST01)"
Else
''2015/5/6 END
   strEmp = "S2.ST01"
End If
strSql = "SELECT S1.ST03, NP10,NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(TM15,TM12),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NP15," & strEmp & ",NVL(CU1.CU04,NVL(CU1.CU05||CU1.CU88||CU1.CU89||CU1.CU90,CU1.CU06)),NVL(N1.NA03,N1.NA04),decode(np07,'102',decode(tm33,null,NVL(FA1.FA04,NVL(FA1.FA05||FA1.FA63||FA1.FA64||FA1.FA65,FA1.FA06)),NVL(NVL(FA2.FA04,NVL(FA2.FA05||FA2.FA63||FA2.FA64||FA2.FA65,FA2.FA06)),NVL(CU2.CU04,NVL(CU2.CU05||CU2.CU88||CU2.CU89||CU2.CU90,CU2.CU06)))),NVL(FA1.FA04,NVL(FA1.FA05||FA1.FA63||FA1.FA64||FA1.FA65,FA1.FA06))),decode(np07,'102',decode(tm33,null,NVL(N2.NA03,N2.NA04),NVL(NVL(N4.NA03,N4.NA04),NVL(N5.NA03,N5.NA04))),NVL(N2.NA03,N2.NA04)),TM22,np07,NP01,CP27,TM01,TM02,TM03,TM04,S1.ST15 as ST15 FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,FAGENT FA1,FAGENT FA2,CUSTOMER CU1,CUSTOMER CU2, Nation N3,nation N4,nation N5 " & _
              " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND Fa2.FA10=N4.na01(+) and FA1.FA10=N2.NA01(+) AND TM10=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(TM23,1,8)=CU1.CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU1.CU02(+) AND SUBSTR(TM44,1,8)=FA1.FA01(+) AND SUBSTR(TM44,9,1)=FA1.FA02(+) And CU1.CU10=N3.NA01(+) And CU2.CU10=N5.NA01(+) and SUBSTR(TM33,1,8)=FA2.FA01(+) AND SUBSTR(TM33,9,1)=FA2.FA02(+) and SUBSTR(TM33,1,8)=CU2.CU01(+) AND SUBSTR(TM33,9,1)=CU2.CU02(+) " & strSQL1 & StrSQL6
'2009/12/10 END
strSql = strSql + " union all select S1.ST03, NP10,NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),'',SP11,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NP15," & strEmp & ",NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),NVL(N1.NA03,N1.NA04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)),NVL(N2.NA03,N2.NA04),SP21,NP07,NP01,CP27,SP01,SP02,SP03,SP04,S1.ST15 as ST15 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,FAGENT,CUSTOMER, Nation N3 WHERE NP01=CP09(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+) AND FA10=N2.NA01(+) AND SP09=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) And CU10=N3.NA01(+) " & strSQL2 & StrSQL6
intRow = 0
' 98/02/27 End
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        DoEvents
        Do While .EOF = False
            For i = 0 To 22 '16
                'edit by nick
                'strTemp(i) = CheckStr(.Fields(i))
                strTemp(i) = "" & .Fields(i)
            Next i
            strST15 = "" & .Fields("ST15") 'Add By Sindy 2011/1/12
            
            'Modify By Sindy 2011/1/12
            'Modify By Sindy 2010/12/16
            ' Add By Sindy 98/02/27
            '檢查若智權人員離職時, 需要重新取得目前承辦智權人員
            strSql = "select st03 from staff where upper(st01)=" & CNULL(UCase(strTemp(1))) & " and st04<>'1' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If Left(RsTemp("st03"), 1) = "F" Then
                  '取得FCT承辦智權人員
                  strTemp(0) = GetStaffDepartment(PUB_GetFCTSalesNo(strTemp(19), strTemp(20), strTemp(21), strTemp(22)))
                  strTemp(1) = PUB_GetFCTSalesNo(strTemp(19), strTemp(20), strTemp(21), strTemp(22))
                  strST15 = PUB_GetStaffST15(PUB_GetFCTSalesNo(strTemp(19), strTemp(20), strTemp(21), strTemp(22)), "1")
               Else
                  '取得目前承辦智權人員
                  strTemp(0) = GetStaffDepartment(PUB_GetAKindSalesNo(strTemp(19), strTemp(20), strTemp(21), strTemp(22)))
                  strTemp(1) = PUB_GetAKindSalesNo(strTemp(19), strTemp(20), strTemp(21), strTemp(22))
                  strST15 = PUB_GetStaffST15(PUB_GetAKindSalesNo(strTemp(19), strTemp(20), strTemp(21), strTemp(22)), "1")
               End If
            End If
            If Len(TXT1(8)) <> 0 Then
                If strTemp(1) <> Trim(TXT1(8)) Then GoTo gotoExit
            End If
            If (InStr(1, TXT1(0), "CFT") <> 0 Or InStr(1, TXT1(0), "CFC") <> 0) Then
               If TXT1(17) = "1" Then
                  If Left(strST15, 1) = "S" Then GoTo gotoExit
               End If
            End If
            intRow = intRow + 1 '記錄筆數
            ' 98/02/27 End
            
            m_TM01 = SystemNumber(strTemp(4), 1)
            m_TM02 = SystemNumber(strTemp(4), 2)
            m_TM03 = SystemNumber(strTemp(4), 3)
            m_TM04 = SystemNumber(strTemp(4), 4)
            
            'Add By Sindy 2013/8/16 是否為不催延展者
            If PUB_ChkCaseIsNoticeScale(m_TM01, m_TM02, m_TM03, m_TM04) = False Then
               strTemp(4) = "x" & strTemp(4)
            '2013/8/16 END
            Else
               If Val(strTemp(2)) < Val(strSrvDate(1)) Then
                   strTemp(4) = "*" & strTemp(4)
               Else
                   If Val(strTemp(2)) = Val(strSrvDate(1)) Then
                       strTemp(4) = "V" & strTemp(4)
                   Else
                       If Mid(CheckStr(.Fields(17)), 1, 1) = "C" And Len(CheckStr(.Fields(18))) = 0 Then
                           strTemp(4) = "#" & strTemp(4)
                       'Add By Sindy 2021/3/19 為排序
                       Else
                           strTemp(4) = " " & strTemp(4)
                       '2021/3/19 END
                       End If
                   End If
               End If
            End If
            'Add By Cheng 2003/07/04
            '本所期限
            strTemp(2) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(2)))
            '法定期限
            strTemp(3) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
            'Modify By Cheng 2003/11/20
'            If Val(strTemp(16)) = 102 Then
            'edit by nick 2004/08/04 將 715 移到第一個報表
            'If Val(strTemp(16)) = 102 Or Val(strTemp(16)) = 715 Or Val(strTemp(16)) = 716 Then
            '2005/12/7 MODIFY BY SONIA 將 716 移到第一個報表
            'If Val(strTemp(16)) = 102 Or Val(strTemp(16)) = 716 Then
            If Val(strTemp(16)) = 102 Then
               '2006/1/9 MODIFY BY SONIA 阿蓮說CFT只有輸入案件性質102時才寫入R030403_2,若未輸入案件性質但Text2(0)="Y"時則寫入R030403_1
               'strSQL = " INSERT INTO R030403_2 VALUES ('" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(15)) & "','" & strUserNum & "') "
               If TXT1(0) = "CFT" And TXT1(5) <> "102" And Text2(0) = "Y" Then
                  strSql = " INSERT INTO R030403_1 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & strUserNum & "') "
               Else
                  'Modify By Sindy 2015/5/6 strTemp(13).代理人==>strTemp(10).CFT承辦人
                  If TXT1(0) = "CFT" Then
                     strSql = " INSERT INTO R030403_2 VALUES ('" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(15)) & "','" & strUserNum & "') "
                  Else
                  '2015/5/6 END
                     strSql = " INSERT INTO R030403_2 VALUES ('" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(15)) & "','" & strUserNum & "') "
                  End If
               End If
               cnnConnection.Execute strSql
            Else
                strSql = " INSERT INTO R030403_1 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & strUserNum & "') "
                cnnConnection.Execute strSql
            End If
gotoExit:
            .MoveNext
            DoEvents
        Loop
        ' Add By Sindy 98/02/27
        If intRow = 0 Then
            InsertQueryLog (0) 'Add By Sindy 2010/10/22
            ShowNoData
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            InsertQueryLog (intRow) 'Add By Sindy 2010/10/22
        End If
        ' 98/02/27 End
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/22
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End With
CheckOC
PrintData  '列印-外商智權人員期限管制表 (R030403_1)
PrintData2 '列印-外商延展管制表(R030403_2)
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

'列印-外商智權人員期限管制表 (R030403_1)
Sub PrintData()
'Add By Cheng 2002/08/28
Dim strSaleName As String '智權人員

'Add By Cheng 2002/08/28
strSaleName = ""

If Option1(0).Value = True Then
   '91/03/12 日期排序不能用符號
   'nick
   'Modify By Sindy 2010/12/16 修改order by百年日期問題
   'strSql = "SELECT DISTINCT NVL(NVL(A0902,A0903),ST03),st02,r092003,r092004,r092005,r092006,r092007,r092008,r092009,r092010,r092011,r092012,r092013,r092002,R092001 FROM R030403_1,staff,ACC090 WHERE ID='" & strUserNum & "' and r092002=st01(+) AND ST03=A0901(+) ORDER BY R092001,R092002,decode(substr(R092003,1,1),'#',substr(R092003,2,10),'V',substr(R092003,2,10),'*',substr(R092003,2,10),R092003),R092005 "
   'Modify By Sindy 2015/5/6 R092011:承辦人姓名改放承辦人員編
   strSql = "SELECT DISTINCT NVL(NVL(A0902,A0903),s1.ST03),s1.st02,r092003,r092004,r092005,r092006,r092007,r092008,r092009,r092010,s2.st02,r092012,r092013,r092002,R092001,R092011 FROM R030403_1,staff s1,staff s2,ACC090 WHERE ID='" & strUserNum & "' and r092002=s1.st01(+) AND s1.ST03=A0901(+) and r092011=s2.st01(+) ORDER BY R092001,R092002,substr('0'||R092003,length(R092003)-8+1,9),R092005"
Else
   'Modify By Sindy 2010/12/16 修改order by百年日期問題
   'strSql = "SELECT DISTINCT NVL(NVL(A0902,A0903),ST03),st02,r092003,r092004,r092005,r092006,r092007,r092008,r092009,r092010,r092011,r092012,r092013,r092002,R092001 FROM R030403_1,staff,ACC090 WHERE ID='" & strUserNum & "' and r092002=st01(+) AND ST03=A0901(+) ORDER BY R092001,R092002,R092004,R092005 "
   'Modify By Sindy 2015/5/6 R092011:承辦人姓名改放承辦人員編
   '                         CFT要改排序
   'ORDER BY R092001,R092002,substr('0'||R092004,length(R092004)-8+1,9),R092005 ==> ORDER BY R092011,R092013,substr('0'||R092004,length(R092004)-8+1,9),R092005 (改以承辦人+申請國家+可辦期限+本所案號排序)
   If TXT1(0) = "CFT" Then
      strSql = "SELECT DISTINCT NVL(NVL(A0902,A0903),s1.ST03),s1.st02,r092003,r092004,r092005,r092006,r092007,r092008,r092009,r092010,s2.st02,r092012,r092013,r092002,R092001,R092011 FROM R030403_1,staff s1,staff s2,ACC090 WHERE ID='" & strUserNum & "' and r092002=s1.st01(+) AND s1.ST03=A0901(+) and r092011=s2.st01(+) ORDER BY R092011,R092013,substr('0'||R092004,length(R092004)-8+1,9),R092005"
   Else
   '2015/5/6 END
      strSql = "SELECT DISTINCT NVL(NVL(A0902,A0903),s1.ST03),s1.st02,r092003,r092004,r092005,r092006,r092007,r092008,r092009,r092010,s2.st02,r092012,r092013,r092002,R092001,R092011 FROM R030403_1,staff s1,staff s2,ACC090 WHERE ID='" & strUserNum & "' and r092002=s1.st01(+) AND s1.ST03=A0901(+) and r092011=s2.st01(+) ORDER BY R092001,R092002,substr('0'||R092004,length(R092004)-8+1,9),R092005"
   End If
End If
CheckOC
Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        SavDay1 = CheckStr(.Fields(0))
        SavDay2 = CheckStr(.Fields(1))
        PrintTitle
        Do While .EOF = False
            For i = 0 To 12
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If Trim(TXT1(5)) = "" Then 'Add By Sindy 2012/3/9 +if 案件性質沒輸入時才要跳頁
               If SavDay1 <> strTemp(0) Or SavDay2 <> strTemp(1) Then
                   Page = Page + 1
                   Printer.CurrentX = 0
                  Printer.CurrentY = iPrint
                  Printer.Print String(200, "-")
                  iPrint = iPrint + 300
                  Printer.CurrentX = 0
                  Printer.CurrentY = iPrint
                  Printer.Print "* 表示逾本所期限 , V 表示當日本所期限 , # 表示承辦人未通知主管機關來函 , x 表示不催延展"
                   Printer.NewPage
                   SavDay1 = strTemp(0)
                   SavDay2 = strTemp(1)
                   PrintTitle
                   'Add By Cheng 2002/08/28
                   strSaleName = ""
               End If
            End If
            strTemp(1) = StrToStr(strTemp(1), 4)
            'Add By Cheng 2002/08/28
            If strTemp(1) = strSaleName Then
               strTemp(1) = ""
            Else
               strSaleName = strTemp(1)
            End If
            'strTemp(2) = IIf("" & strTemp(2) <> "", ChangeTStringToTDateString(strTemp(2) - 19110000), "")
            
            strTemp(5) = StrToStr(strTemp(5), 4)
            strTemp(6) = StrToStr(strTemp(6), 9)
            strTemp(7) = StrToStr(strTemp(7), 7)
            strTemp(8) = StrToStr(strTemp(8), 4)
            strTemp(9) = StrToStr(strTemp(9), 5)
            strTemp(10) = StrToStr(strTemp(10), 4)
            strTemp(11) = StrToStr(strTemp(11), 9)
            strTemp(12) = StrToStr(strTemp(12), 4)
            PrintDatil
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.CurrentX = 0
               Printer.CurrentY = iPrint
               Printer.Print String(200, "-")
               iPrint = iPrint + 300
               Printer.CurrentX = 0
               Printer.CurrentY = iPrint
               Printer.Print "* 表示逾本所期限 , V 表示當日本所期限 , # 表示承辦人未通知主管機關來函 , x 表示不催延展"

                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    Else
        Exit Sub
    End If
End With
CheckOC
Printer.CurrentX = 0
               Printer.CurrentY = iPrint
               Printer.Print String(200, "-")
               iPrint = iPrint + 300
               Printer.CurrentX = 0
               Printer.CurrentY = iPrint
               Printer.Print "* 表示逾本所期限 , V 表示當日本所期限 , # 表示承辦人未通知主管機關來函 , x 表示不催延展"

Printer.EndDoc
End Sub

'列印抬頭-外商智權人員期限管制表 (R030403_1)
Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
DoEvents
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5700
Printer.CurrentY = iPrint
Printer.Print "外商智權人員期限管制表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
If Option1(0).Value = True Then
    Printer.Print "本所期限：" & Format(ChangeTStringToTDateString(TXT1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(TXT1(2))
'edit by nickc 2007/03/05
'else
ElseIf Option1(1).Value Then
    Printer.Print "法定期限：" & Format(ChangeTStringToTDateString(TXT1(3)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(TXT1(4))
'add by nickc 2007/03/05
ElseIf Option1(3).Value Then
    Printer.Print "可辦期限：" & Format(ChangeTStringToTDateString(TXT1(18)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(TXT1(19))
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
If Trim(TXT1(5)) = "" Then 'Add By Sindy 2012/3/9 +if 案件性質沒輸入時才要跳頁
   Printer.Print "業務區：" & SavDay1
End If
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
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
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "法定期限"
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
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "備  註"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "申請人"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "申請國家"
iPrint = iPrint + 300
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
End Sub

Sub PrintDatil()
For i = 1 To 12
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500 - 500
PLeft(1) = 500 - 500
PLeft(2) = 1500 - 500
PLeft(3) = 2500 - 500
PLeft(4) = 3500 - 500
PLeft(5) = 5000 - 500 + 500
PLeft(6) = 6000 - 500 + 500
PLeft(7) = 8000 - 500 + 500
PLeft(8) = 9800 - 500 + 500
PLeft(9) = 10800 - 500 + 500
PLeft(10) = 12000 - 500 + 500
PLeft(11) = 13000 - 500 + 500
PLeft(12) = 15000 - 500 + 500
PLeft(13) = 17000 - 500 + 500
End Sub

'列印-外商延展管制表(R030403_2)
Sub PrintData2()
'92.7.14 modify by sonia
'strSQL = "SELECT DISTINCT * FROM R030403_2 WHERE ID='" & strUserNum & "' ORDER BY R093003 "
'Modify By Cheng 2003/10/07
'strSQL = "SELECT DISTINCT * FROM R030403_2 WHERE ID='" & strUserNum & "' ORDER BY decode(substr(R093003,1,1),'#',substr(R093003,2,10),'V',substr(R093003,2,10),'*',substr(R093003,2,10),R093003) "
'Modify By Sindy 2015/5/6 CFT要將代理人欄位==>承辦人欄位
If TXT1(0) = "CFT" Then
   strSql = "select st02,R093002,R093003,R093004,R093005,R093006,R093007,R093008,R093009,R093010 from staff,(SELECT DISTINCT * FROM R030403_2 WHERE ID='" & strUserNum & "')" & _
            " where R093001=st01(+)" & _
            " ORDER BY R093001,Decode(substr(R093003,Length(R093003)-3,1),'T','T','0'), decode(substr(R093003,1,1),'#',substr(R093003,2,10),'V',substr(R093003,2,10),'*',substr(R093003,2,10),R093003)"
Else
'2015/5/6 END
   'Modify By Sindy 2021/3/22 陳金蓮:請調整先列印所有要催延展之案件，如將FCT-XXXXXX-T案移至不催延展案件前面
   'strSql = "SELECT DISTINCT * FROM R030403_2 WHERE ID='" & strUserNum & "' ORDER BY Decode(substr(R093003,Length(R093003)-3,1),'T','T','0'), decode(substr(R093003,1,1),'#',substr(R093003,2,10),'V',substr(R093003,2,10),'*',substr(R093003,2,10),R093003)"
   strSql = "SELECT DISTINCT * FROM R030403_2 WHERE ID='" & strUserNum & "' ORDER BY decode(substr(R093003,1,1),'x',1,0), Decode(substr(R093003,Length(R093003)-3,1),'T','T','0'), R093003"
   '2021/3/22 END
End If
'92.7.14 end
CheckOC
Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        'SavDay1 = CheckStr(.Fields(0))
        PrintTitle2
        Do While .EOF = False
            For i = 0 To 9
               strTemp(i) = CheckStr(.Fields(i))
            Next i
            'If SavDay1 <> StrTemp(0) Then
            '    Page = Page + 1
            '    Printer.NewPage
            '    SavDay1 = StrTemp(0)
            '    PrintTitle
            'End If
            strTemp(0) = StrToStr(strTemp(0), 9)
            strTemp(1) = StrToStr(strTemp(1), 10)
            strTemp(3) = StrToStr(strTemp(3), 9)
            strTemp(4) = StrToStr(strTemp(4), 8)
            strTemp(5) = StrToStr(strTemp(5), 7)
            'Add By Cheng 2002/08/28
            '列印格式改為民國日期
            If strTemp(9) <> "" Then strTemp(9) = ChangeTStringToTDateString(strTemp(9) - 19110000)
            PrintDatil2
            If iPrint >= 10000 Then
               Page = Page + 1
               Printer.CurrentX = 0
               Printer.CurrentY = iPrint
               Printer.Print String(200, "-")
               iPrint = iPrint + 300
               Printer.CurrentX = 0
               Printer.CurrentY = iPrint
               Printer.Print "* 表示逾本所期限 , V 表示當日本所期限 , # 表示承辦人未通知主管機關來函 , x 表示不催延展"
Printer.NewPage
                PrintTitle2
            End If
            .MoveNext
        Loop
    Else
      Exit Sub
        
    End If
End With
CheckOC
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "* 表示逾本所期限 , V 表示當日本所期限 , # 表示承辦人未通知主管機關來函 , x 表示不催延展"
Printer.EndDoc
End Sub

'列印抬頭-外商延展管制表(R030403_2)
Sub PrintTitle2()
Dim strTitName As String

GetPleft2

iPrint = 500
Printer.Orientation = 2
DoEvents
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
'Add By Sindy 2023/9/20 報表類別=4.未催延展檢核表
If Trim(TXT1(6)) = "4" Then
   strTitName = "未催延展檢核表" & IIf(TXT1(7) = "2", "(日文組)", "(英文組)")
Else
'2023/9/20 END
   strTitName = "外商延展管制表" & IIf(TXT1(7) = "", "", IIf(TXT1(7) = "2", "(日文組)", "(英文組)"))
End If
'Printer.CurrentX = 6700
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strTitName) / 2)
Printer.CurrentY = iPrint
Printer.Print strTitName
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6700
Printer.CurrentY = iPrint
If Option1(0).Value = True Then
    Printer.Print "本所期限：" & Format(ChangeTStringToTDateString(TXT1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(TXT1(2))
'edit by nickc 2007/03/05
'Else
ElseIf Option1(1).Value Then
    Printer.Print "法定期限：" & Format(ChangeTStringToTDateString(TXT1(3)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(TXT1(4))
'add by nickc 2007/03/05
ElseIf Option1(3).Value Then
    Printer.Print "可辦期限：" & Format(ChangeTStringToTDateString(TXT1(18)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(TXT1(19))
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
'Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
'Modify By Sindy 2015/5/6 CFT要將代理人欄位==>承辦人欄位
If TXT1(0) = "CFT" Then
   'Modified by Lydia 2016/04/12
   'Printer.Print "承辦人"
   Printer.Print "該國承辦人"
Else
'2015/5/6 END
   Printer.Print "代理人"
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "申請人"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "商品類別"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "審定號"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
'Modify By Sindy 2015/5/6 CFT取消代理人國籍
If TXT1(0) = "CFT" Then
   Printer.Print ""
Else
'2015/5/6 END
   Printer.Print "代理人國籍"
End If
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "申請國家"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "本所期限"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "專用期止日"
iPrint = iPrint + 300
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
End Sub

Sub PrintDatil2()
For i = 0 To 9
    'Modified by Lydia 2016/04/12 跳過代理人國籍
    If TXT1(0) = "CFT" And i = 6 Then
    Else
        Printer.CurrentX = PLeft(i)
        Printer.CurrentY = iPrint
        'Modified by Lydia 2016/10/06
        'Printer.Print strTemp(i)
        Select Case i
            Case 7: Printer.Print convForm(strTemp(i), 8)
            Case Else: Printer.Print strTemp(i)
        End Select
       
    End If
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft2()
Erase PLeft
PLeft(0) = 500 - 500
PLeft(1) = 2500 - 500 + 125
PLeft(2) = 4700 - 500 + 125
PLeft(3) = 6500 - 500 + 125
PLeft(4) = 8500 - 500 + 125 + 500
PLeft(5) = 10300 - 500 + 125 + 500
PLeft(6) = 11800 - 500 + 125 + 500
PLeft(7) = 13000 - 500 + 125 + 500
PLeft(8) = 14000 - 500 + 125 + 500
PLeft(9) = 15000 - 500 + 125 + 500
End Sub

''期限管制表-已收文
'Sub Process1()
''cnnConnection.Execute "DELETE FROM R030403_1 WHERE ID='" & strUserNum & "' "
''cnnConnection.Execute "DELETE FROM R030403_2 WHERE ID='" & strUserNum & "' "
'Screen.MousePointer = vbHourglass
'cnnConnection.Execute "DELETE FROM R030403_3 WHERE ID='" & strUserNum & "' "
'strSQL1 = ""
'strSQL2 = ""
'StrSQL6 = ""
'If Len(TXT1(0)) <> 0 Then
'   strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(TXT1(0), 2) & ") "
'   strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(TXT1(0), 5) & ") "
'   pub_QL05 = pub_QL05 & ";" & Label1(0) & TXT1(0)  'Add By Sindy 2010/10/22
'End If
'StrSQL6 = ""
'If Len(TXT1(5)) <> 0 Then
'    pub_QL05 = pub_QL05 & ";" & Label1(1) & TXT1(5)  'Add By Sindy 2010/10/22
'    StrSQL6 = " AND ("
'    If Len(TXT1(5)) <> 0 Then
'        strTemp1 = ""
'        strTemp1 = Split(Replace(TXT1(5), ",,", ""), ",")
'        For i = 0 To UBound(strTemp1)
'            StrSQL6 = StrSQL6 + " CP10='" & strTemp1(i) & "' OR "
'        Next i
'        StrSQL6 = StrSQL6 + " CP10='') "
'    End If
'End If
'StrSQL6 = StrSQL6 + " AND (CP27 IS NULL OR CP27='') AND (CP57 IS NULL OR CP57='') "
''本所期限
'If Option1(0).Value = True Then
'    If Len(TXT1(1)) <> 0 Then
'      StrSQL6 = StrSQL6 + " AND CP06>=" & Val(ChangeTStringToWString(TXT1(1))) & " "
'    End If
'    If Len(Trim(TXT1(2))) <> 0 Then
'      StrSQL6 = StrSQL6 + " AND CP06<=" & Val(ChangeTStringToWString(TXT1(2)))
'    End If
'    If Len(Trim(TXT1(1))) <> 0 Or Len(Trim(TXT1(2))) <> 0 Then
'      pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & TXT1(1) & "-" & TXT1(2)  'Add By Sindy 2010/10/22
'    End If
''Modify By Cheng 2003/01/08
''法定期限
'ElseIf Me.Option1(1).Value Then
'    If Len(TXT1(3)) <> 0 Then
'    StrSQL6 = StrSQL6 + " AND CP07>=" & Val(ChangeTStringToWString(TXT1(3))) & " "
'    End If
'    If Len(Trim(TXT1(4))) <> 0 Then
'      StrSQL6 = StrSQL6 + " AND CP07<=" & Val(ChangeTStringToWString(TXT1(4)))
'    End If
'    If Len(Trim(TXT1(3))) <> 0 Or Len(Trim(TXT1(4))) <> 0 Then
'      pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & TXT1(3) & "-" & TXT1(4)  'Add By Sindy 2010/10/22
'    End If
''Add By Cheng 2003/01/08
''本所案號
''edit by nickc 2007/03/05
''Else
'ElseIf Option1(2).Value Then
'    StrSQL6 = StrSQL6 + " AND CP01='" & Me.text1(0).Text & "' And CP02='" & Me.text1(1).Text & "' And CP03='" & Me.text1(2).Text & "' And CP04='" & Me.text1(3).Text & "' "
'    pub_QL05 = pub_QL05 & ";" & Option1(2).Caption & text1(0) & "-" & text1(1) & "-" & text1(2) & "-" & text1(3) 'Add By Sindy 2010/10/22
''edit by nickc 2007/03/05 可辦只有下一程序有
''add by nickc 2007/03/05
''ElseIf Option1(3).Value Then
'''to_char(add_months(to_date(ChangeTStringToWString(TXT1(18)),'YYYYMMDD'),na15 * -1),'YYYYMMDD')
''    If Len(Trim(txt1(18))) <> 0 Then
''    StrSQL6 = StrSQL6 + " AND to_char(cp07)>=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(18)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') "
''    End If
''    If Len(Trim(txt1(19))) <> 0 Then
''      StrSQL6 = StrSQL6 + " AND to_char(cp07)<=to_char(add_months(to_date(" & ChangeTStringToWString(txt1(19)) & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') "
''    End If
'End If
'If Len(TXT1(8)) <> 0 Then
'    StrSQL6 = StrSQL6 + " AND CP13='" & TXT1(8) & "' "
'    pub_QL05 = pub_QL05 & ";" & Label1(5) & TXT1(8) & LBL1(0) 'Add By Sindy 2010/10/22
'End If
'If Len(TXT1(9)) <> 0 Then
'    StrSQL6 = StrSQL6 + " AND CP14='" & TXT1(9) & "' "
'    pub_QL05 = pub_QL05 & ";" & Label1(3) & TXT1(9) & LBL1(1) 'Add By Sindy 2010/10/22
'End If
'If Len(TXT1(10)) <> 0 Then
'    strSQL1 = strSQL1 + " AND (TM23>='" & TXT1(10) & "') "
'    strSQL2 = strSQL2 + " AND (SP08>='" & TXT1(10) & "' OR SP58>='" & TXT1(10) & "' OR SP59>='" & TXT1(10) & "') "
'End If
'If Len(TXT1(11)) <> 0 Then
'    strSQL1 = strSQL1 + " AND (TM23<='" & TXT1(11) & "') "
'    strSQL2 = strSQL2 + " AND (SP08<='" & TXT1(11) & "' OR SP58<='" & TXT1(11) & "' OR SP59<='" & TXT1(11) & "') "
'End If
'If Len(Trim(TXT1(10))) <> 0 Or Len(Trim(TXT1(11))) <> 0 Then
'   pub_QL05 = pub_QL05 & ";" & Label1(7) & TXT1(10) & "-" & TXT1(11) 'Add By Sindy 2010/10/22
'End If
'If Len(TXT1(12)) <> 0 Then
'    strSQL1 = strSQL1 + " and TM44>='" & TXT1(12) & "' "
'    strSQL2 = strSQL2 + " and SP26>='" & TXT1(12) & "' "
'End If
'If Len(TXT1(13)) <> 0 Then
'    strSQL1 = strSQL1 + " and TM44<='" & TXT1(13) & "' "
'    strSQL2 = strSQL2 + " and SP26<='" & TXT1(13) & "' "
'End If
'If Len(Trim(TXT1(12))) <> 0 Or Len(Trim(TXT1(13))) <> 0 Then
'   pub_QL05 = pub_QL05 & ";" & Label1(8) & TXT1(12) & "-" & TXT1(13) 'Add By Sindy 2010/10/22
'End If
''add by nick 2005/02/15 加入申請國家
'If Len(Trim(TXT1(15))) <> 0 Then
'   strSQL1 = strSQL1 & " and tm10>='" & TXT1(15) & "' "
'   strSQL2 = strSQL2 & " and sp09>='" & TXT1(15) & "' "
'End If
'If Len(Trim(TXT1(16))) <> 0 Then
'   strSQL1 = strSQL1 & " and tm10<='" & TXT1(16) & "' "
'   strSQL2 = strSQL2 & " and sp09<='" & TXT1(16) & "' "
'End If
''add end
'If Len(Trim(TXT1(15))) <> 0 Or Len(Trim(TXT1(16))) <> 0 Then
'   pub_QL05 = pub_QL05 & ";" & Label1(13) & TXT1(15) & "-" & TXT1(16) 'Add By Sindy 2010/10/22
'End If
'strSQL1 = strSQL1 + " and (tm29 is null or tm29 <> 'Y' OR TM29='') "
'strSQL2 = strSQL2 + " AND (SP15 IS NULL OR SP15 <> 'Y' OR SP15='') "
''add by nickc 2006/05/30
'If Val(TXT1(6)) = 1 And (InStr(1, TXT1(0), "CFT") <> 0 Or InStr(1, TXT1(0), "CFC") <> 0) Then
'    If TXT1(17) = "1" Then
'        StrSQL6 = StrSQL6 + " AND substr(s1.ST15,1,1)<>'S' "
'    End If
'End If
''add by nickc 2006/05/30 延展和第二期專用權須存在  FCT  才做
'If InStr(1, TXT1(5), "716") <> 0 Or InStr(1, TXT1(5), "102") <> 0 Then
'    'edit by nickc 2007/07/13 加入CFT
'    'strSQL1 = strSQL1 & " and cp01||decode(cp10,'716',tm17,'102',tm17,'Y')='FCTY' "
'    strSQL1 = strSQL1 & " and decode(cp01,'CFT','FCTY',cp01||decode(cp10,'716',tm17,'102',tm17,'Y'))='FCTY' "
'End If
'CheckOC
''edit by nickc 2007/03/05 加入國家
''    strSQL = "SELECT CP14,CP06,CP07,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04,TM09,NVL(TM15,TM12),NVL(TM05,NVL(TM06,TM07)),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),CP64,S1.ST02,Cp10,CP09,CP27 FROM CASEPROGRESS,TRADEMARK,STAFF S1,CASEPROPERTYMAP WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & StrSQL6
''strSQL = strSQL + " union all select CP14,CP06,CP07,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04,'',SP11,NVL(SP05,NVL(SP06,SP07)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),CP64,S1.ST02,CP10,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,CASEPROPERTYMAP WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=S1.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & StrSQL6
'strSql = "SELECT CP14,CP06,CP07,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04,TM09,NVL(TM15,TM12),NVL(TM05,NVL(TM06,TM07)),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),CP64,S1.ST02,Cp10,CP09,CP27 FROM CASEPROGRESS,TRADEMARK,STAFF S1,CASEPROPERTYMAP,nation n1 WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP13=S1.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and tm10=n1.na01(+) " & strSQL1 & StrSQL6
'strSql = strSql + " union all select CP14,CP06,CP07,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04,'',SP11,NVL(SP05,NVL(SP06,SP07)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),CP64,S1.ST02,CP10,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,CASEPROPERTYMAP,nation n1 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP13=S1.ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) and sp09=n1.na01(+) " & strSQL2 & StrSQL6
'With adoRecordset
'    .CursorLocation = adUseClient
'    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/22
'        .MoveFirst
'        DoEvents
'        Do While .EOF = False
'            For i = 0 To 11
'                strTemp(i) = CheckStr(.Fields(i))
'            Next i
'
'            m_TM01 = SystemNumber(strTemp(4), 1)
'            m_TM02 = SystemNumber(strTemp(4), 2)
'            m_TM03 = SystemNumber(strTemp(4), 3)
'            m_TM04 = SystemNumber(strTemp(4), 4)
'            'Add By Sindy 2013/8/16 是否為不催延展者
'            If PUB_ChkCaseIsNoticeScale(m_TM01, m_TM02, m_TM03, m_TM04) = False Then
'               strTemp(4) = "x" & strTemp(4)
'            '2013/8/16 END
'            Else
'               If Val(strTemp(1)) < Val(strSrvDate(1)) Then
'                   strTemp(4) = "*" & strTemp(4)
'               Else
'                   If Val(strTemp(1)) = Val(GetTodayDate) Then
'                       strTemp(4) = "V" & strTemp(4)
'                   Else
'                       If Mid(CheckStr(.Fields(12)), 1, 1) = "C" And Len(CheckStr(.Fields(13))) = 0 Then
'                           strTemp(4) = "#" & strTemp(4)
'                       End If
'                   End If
'               End If
'            End If
'            strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1))) 'Add By Sindy 2010/12/16
'            strTemp(2) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(2)))
'            strTemp(3) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
'            strSql = " INSERT INTO R030403_3 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & strUserNum & "') "
'            cnnConnection.Execute strSql
'            .MoveNext
'            DoEvents
'        Loop
'    Else
'        InsertQueryLog (0) 'Add By Sindy 2010/10/22
'        ShowNoData
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'End With
'CheckOC
'PrintData1 '列印-外商承辦人期限管制表(R030403_3)
'ShowPrintOk
'Screen.MousePointer = vbDefault
'End Sub

'列印-外商承辦人期限管制表(R030403_3)
Sub PrintData1()
If Option1(0).Value = True Then
   'Modify By Sindy 2010/12/16 修改order by百年日期問題
   'strSql = "SELECT DISTINCT st02,r094002,r094003,r094004,r094005,r094006,r094007,r094008,r094009,r094010,r094011,r094001 FROM R030403_3,staff  WHERE ID='" & strUserNum & "' and r094001=st01(+) ORDER BY R094001,R094002,R094004,R094005 "
   strSql = "SELECT DISTINCT st02,r094002,r094003,r094004,r094005,r094006,r094007,r094008,r094009,r094010,r094011,r094001 FROM R030403_3,staff  WHERE ID='" & strUserNum & "' and r094001=st01(+) ORDER BY R094001,substr('0'||R094002,length(R094002)-8+1,9),substr('0'||R094004,length(R094004)-8+1,9),R094005 "
Else
   'Modify By Sindy 2010/12/16 修改order by百年日期問題
   'strSql = "SELECT DISTINCT st02,r094002,r094003,r094004,r094005,r094006,r094007,r094008,r094009,r094010,r094011,r094001 FROM R030403_3,staff WHERE ID='" & strUserNum & "'and r094001=st01(+) ORDER BY R094001,R094003,R094004,R094005 "
   strSql = "SELECT DISTINCT st02,r094002,r094003,r094004,r094005,r094006,r094007,r094008,r094009,r094010,r094011,r094001 FROM R030403_3,staff WHERE ID='" & strUserNum & "'and r094001=st01(+) ORDER BY R094001,substr('0'||R094003,length(R094003)-8+1,9),substr('0'||R094004,length(R094004)-8+1,9),R094005 "
End If
CheckOC
Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        SavDay1 = CheckStr(.Fields(0))
        PrintTitle1
        Do While .EOF = False
            For i = 0 To 10
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If SavDay1 <> strTemp(0) Then
                Page = Page + 1
                Printer.CurrentX = 0
               Printer.CurrentY = iPrint
               Printer.Print String(200, "-")
               iPrint = iPrint + 300
               Printer.CurrentX = 0
               Printer.CurrentY = iPrint
               Printer.Print "* 表示逾本所期限 , V 表示當日本所期限 , # 表示承辦人未通知主管機關來函 , x 表示不催延展"
Printer.NewPage
                SavDay1 = strTemp(0)
                PrintTitle1
            End If
            strTemp(5) = StrToStr(strTemp(5), 9)
            strTemp(6) = StrToStr(strTemp(6), 9)
            strTemp(7) = StrToStr(strTemp(7), 9)
            strTemp(8) = StrToStr(strTemp(8), 10)
            strTemp(9) = StrToStr(strTemp(9), 9)
            strTemp(10) = StrToStr(strTemp(10), 4)
            PrintDatil1
            If iPrint >= 10000 Then '16000
                Page = Page + 1
Printer.CurrentX = 0
               Printer.CurrentY = iPrint
               Printer.Print String(200, "-")
               iPrint = iPrint + 300
               Printer.CurrentX = 0
               Printer.CurrentY = iPrint
               Printer.Print "* 表示逾本所期限 , V 表示當日本所期限 , # 表示承辦人未通知主管機關來函 , x 表示不催延展"
                Printer.NewPage
                PrintTitle1
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC
Printer.CurrentX = 0
               Printer.CurrentY = iPrint
               Printer.Print String(200, "-")
               iPrint = iPrint + 300
               Printer.CurrentX = 0
               Printer.CurrentY = iPrint
               Printer.Print "* 表示逾本所期限 , V 表示當日本所期限 , # 表示承辦人未通知主管機關來函 , x 表示不催延展"
Printer.EndDoc
End Sub

'列印抬頭-外商承辦人期限管制表(R030403_3)
Sub PrintTitle1()
GetPleft1
iPrint = 500
Printer.Orientation = 2
DoEvents
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5700
Printer.CurrentY = iPrint
Printer.Print "外商承辦人期限管制表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6700
Printer.CurrentY = iPrint
If Option1(0).Value = True Then
    Printer.Print "本所期限：" & Format(ChangeTStringToTDateString(TXT1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(TXT1(2))
'edit by nickc 2007/03/05
'Else
ElseIf Option1(1).Value Then
    Printer.Print "法定期限：" & Format(ChangeTStringToTDateString(TXT1(3)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(TXT1(4))
'add by nickc     2007/03/05
ElseIf Option1(3).Value Then
    Printer.Print "可辦期限：" & Format(ChangeTStringToTDateString(TXT1(18)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(TXT1(19))
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "承辦人：" & SavDay1
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "本所期限"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "法定期限"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "收文日"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "商品類別"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "申請案號/審定號"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "備  註"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
iPrint = iPrint + 300
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.Font.Size = 10
End Sub

Sub PrintDatil1()
For i = 1 To 10
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft1()
Erase PLeft
PLeft(0) = 500 - 500
PLeft(1) = 500 - 500
PLeft(2) = 1500 - 500
PLeft(3) = 2500 - 500
PLeft(4) = 3500 - 500
'Modify By Cheng 2002/12/20
'PLeft(5) = 5000 - 500
PLeft(5) = 5000
PLeft(6) = 7000 - 500
PLeft(7) = 9000 - 500
PLeft(8) = 11000 - 500
PLeft(9) = 13000 - 500
PLeft(10) = 15000 - 500
End Sub

Private Sub Form_Initialize()
ReDim tm(1 To TF_TM) As String
End Sub

Private Sub Form_Load()

MoveFormToCenter Me
TXT1(0) = GetSystemKindByNick

SeekPrintL = Printer.Orientation
PUB_SetPrinter Me.Name, Combo1, , False, SeekPrint     'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'add by nickc 2006/09/05 結束時問印接洽結案單
    PUB_PrintCaseCloseSheet strUserNum
    PUB_DeleteCaseCloseSheet strUserNum
    PUB_PrintAddressList strUserNum, Me.Combo1.Text
    '刪除地址條列表資料
    PUB_DeleteAddressList strUserNum
    '初始化序號
    pub_AddressListSN = 0
    '若印表機變動, 則更新列印設定
    If Me.Combo1.Text <> Me.Combo1.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
    End If
    'Add By Cheng 2003/02/17
    '還原預設印表機
    Set Printer = Printers(SeekPrint)
    Printer.Orientation = SeekPrintL
    Set frm030403 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   If Index = 0 Then
      TXT1(1).Enabled = True
      TXT1(2).Enabled = True
      TXT1(3).Enabled = False
      TXT1(4).Enabled = False
      'add by nickc 2007/03/05
      TXT1(18).Enabled = False
      TXT1(19).Enabled = False
      
        Me.text1(0).Enabled = False
        Me.text1(1).Enabled = False
        Me.text1(2).Enabled = False
        Me.text1(3).Enabled = False
      TXT1(1).SetFocus
    'Modify By Cheng 2003/01/08
   ElseIf Index = 1 Then
      TXT1(1).Enabled = False
      TXT1(2).Enabled = False
      TXT1(3).Enabled = True
      TXT1(4).Enabled = True
      'add by nickc 2007/03/05
      TXT1(18).Enabled = False
      TXT1(19).Enabled = False
      
        Me.text1(0).Enabled = False
        Me.text1(1).Enabled = False
        Me.text1(2).Enabled = False
        Me.text1(3).Enabled = False
      TXT1(3).SetFocus
'add by nickc 2007/03/05
   ElseIf Index = 2 Then
        TXT1(1).Enabled = False
        TXT1(2).Enabled = False
        TXT1(3).Enabled = False
        TXT1(4).Enabled = False
        TXT1(18).Enabled = False
        TXT1(19).Enabled = False
        Me.text1(0).Enabled = True
        Me.text1(1).Enabled = True
        Me.text1(2).Enabled = True
        Me.text1(3).Enabled = True
      Me.text1(0).SetFocus
    Else
        TXT1(1).Enabled = False
        TXT1(2).Enabled = False
        TXT1(3).Enabled = False
        TXT1(4).Enabled = False
        'add by nickc 2007/03/05
        TXT1(18).Enabled = True
        TXT1(19).Enabled = True
        
        Me.text1(0).Enabled = False
        Me.text1(1).Enabled = False
        Me.text1(2).Enabled = False
        Me.text1(3).Enabled = False
      Me.TXT1(18).SetFocus
   End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    'Add By Cheng 2003/01/08
    TextInverse Me.text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    'Add By Cheng 2003/01/08
    KeyAscii = UpperCase(KeyAscii)
End Sub

'add by nick 2004/08/04
Private Sub Text2_GotFocus(Index As Integer)
TextInverse Text2(Index)
End Sub

'add by nick 2004/08/04
Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

'add by nick 2004/08/04
Private Sub Text2_LostFocus(Index As Integer)
Select Case Index
    Case 0
     Select Case Text2(0)
     Case "Y", ""
     Case Else
          s = MsgBox("是否含延展只能輸入 Y 或 空白!!", , "USER 輸入錯誤")
          Text2(0).SetFocus
          Text2(0).SelStart = 0
          Text2(0).SelLength = Len(Text2(0))
          Exit Sub
     End Select
End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
TXT1(Index).SelStart = 0
TXT1(Index).SelLength = Len(TXT1(Index))
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   CMDOK(0).SetFocus
End If
End Sub
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
     strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
     strTemp2 = Split(Replace(UCase(TXT1(0)), ",,", ""), ",")
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
            TXT1(0).SetFocus
            TXT1(0).SelStart = 0
            TXT1(0).SelLength = Len(TXT1(0))
            Exit Sub
        End If
     Next i
Case 6
     Select Case Trim(TXT1(Index))
     Case "1", "2", "", "3", "4" 'add by nickc 2007/02/15 加入 3
     Case Else
         'edit by nickc 2007/02/15
         's = MsgBox("延展報表類別只能輸入 1 或 2 !!", , "USER 輸入錯誤")
         'Modify By Sindy 2023/9/20 + 4
         s = MsgBox("延展報表類別只能輸入 1 或 2 或 3 或 4 !!", , "USER 輸入錯誤")
         TXT1(Index).SetFocus
         TXT1(Index).SelStart = 0
         TXT1(Index).SelLength = Len(TXT1(Index))
         Exit Sub
     End Select
Case 7
     Select Case Trim(TXT1(Index))
     Case "1", "2", ""
     Case Else
          s = MsgBox("報表類別(1、4)組別只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          TXT1(Index).SetFocus
          TXT1(Index).SelStart = 0
          TXT1(Index).SelLength = Len(TXT1(Index))
          Exit Sub
     End Select
Case 8
     'lbl1(0) = GetPrjSales(txt1(Index))
     LBL1(0) = GetPrjSalesNM(TXT1(8))
     If Trim(TXT1(Index)) <> "" Then
        If Trim(LBL1(0).Caption) = "" Then
            s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
            TXT1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 9
     'lbl1(1) = GetPrjSales(txt1(Index))
     LBL1(1) = GetPrjSalesNM(TXT1(9))
     If Trim(TXT1(Index)) <> "" Then
        If Trim(LBL1(1).Caption) = "" Then
            s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
            TXT1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 11, 13
     If Trim(TXT1(Index - 1)) <> "" Then
        If Mid(TXT1(Index - 1), 1, 6) <> Mid(TXT1(Index), 1, 6) Then
           s = MsgBox("前6碼必須相同！", , "錯誤！")
           TXT1(Index - 1).SetFocus
           txt1_GotFocus (Index - 1)
           Exit Sub
        End If
     End If
     If RunNick(TXT1(Index - 1), TXT1(Index)) Then
         TXT1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
Case 1, 2, 3, 4, 18, 19   'edit by nickc 2007/03/05 加入 18,19
   If PUB_CheckKeyInDate(Me.TXT1(Index)) = -1 Then
      Me.TXT1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   'edit by nickc 2007/03/05 加入 19
   'If Index = 2 Or Index = 4 Then
   If Index = 2 Or Index = 4 Or Index = 19 Then
     If RunNick(TXT1(Index - 1), TXT1(Index)) Then
         TXT1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
   End If
'add by nick 2005/02/15 加入申請國家
Case 16
     If RunNick(TXT1(Index - 1), TXT1(Index)) Then
         TXT1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
'add by nickc 2006/05/30
Case 17
     Select Case Trim(TXT1(Index))
     Case "1", "2", ""
     Case Else
          s = MsgBox("CFT、CFC 未收文管制表列印對象只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          TXT1(Index).SetFocus
          TXT1(Index).SelStart = 0
          TXT1(Index).SelLength = Len(TXT1(Index))
          Exit Sub
     End Select
Case Else
End Select

'Add By Sindy 2024/4/12 FCT定稿只有英文組
If Index = 0 Or Index = 5 Or Index = 6 Then
   TXT1(7).Enabled = True
   Frame102.Enabled = False
   If TXT1(0) = "FCT" Then
      If Trim(TXT1(6)) = "2" Or Trim(TXT1(6)) = "3" Then
         TXT1(7) = "1" '英文組
         TXT1(7).Enabled = False
         If TXT1(5) = "102" Then
            Frame102.Enabled = True
         End If
      End If
   End If
Else
   If Not (TXT1(0) = "FCT" And (Trim(TXT1(6)) = "2" Or Trim(TXT1(6)) = "3") And TXT1(5) = "102") Then
      TXT1(7).Enabled = True
      Frame102.Enabled = False
   End If
End If
'2024/4/12 END

End Sub

'Add By Cheng 2003/03/14
'列印延展案件
Private Sub PrinterDetail(rsRS As ADODB.Recordset)
Dim intCnt As Integer
Dim intLine As Integer
Dim dblLeft As Double
Dim dblTop As Double
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim intLineHeight As Integer

intLine = 0
intCnt = 0
intLineHeight = 300
dblTop = 3500
If rsRS.RecordCount > 0 Then
    Printer.Font.Size = 12
    Printer.Font.Name = "Times New Roman"
    '列印表頭
    PrintHead dblTop, intLine, intLineHeight
    While Not rsRS.EOF
        If intCnt > 10 Then
            Printer.NewPage
            intLine = 0
            intCnt = 0
            '列印表頭
            PrintHead dblTop, intLine, intLineHeight
        End If
        StrSQLa = "SELECT * FROM TRADEMARK WHERE " & ChgTradeMark(Replace("" & rsRS.Fields(4).Value, "-", ""))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            Printer.CurrentX = 500
            Printer.CurrentY = dblTop + intLine * intLineHeight
            'Modify By Cheng 2003/08/07
            '案件名稱中英日
'            Printer.Print "" & rsA("TM06").Value
            Printer.Print "" & rsA("TM05").Value & rsA("TM06").Value & rsA("TM07").Value
            intLine = intLine + 1
            
            Printer.CurrentX = 500
            Printer.CurrentY = dblTop + intLine * intLineHeight
            Printer.Print "" & rsA("TM15").Value
            
            Printer.CurrentX = 500 + 2000 - 250
            Printer.CurrentY = dblTop + intLine * intLineHeight
            Printer.Print "" & rsA("TM09").Value
            
            Printer.CurrentX = 500 + 3500 - 500
            Printer.CurrentY = dblTop + intLine * intLineHeight
            Printer.Print ChgEngDate("" & rsA("TM22").Value)
            
            Printer.CurrentX = 500 + 5250 '5500
            Printer.CurrentY = dblTop + intLine * intLineHeight
            Printer.Print "" & rsA("TM45").Value
            
            Printer.CurrentX = 500 + 9000 '8500
            Printer.CurrentY = dblTop + intLine * intLineHeight
            Printer.Print Replace("" & rsRS.Fields(4).Value, "-0-00", "")
            intLine = intLine + 2
            intCnt = intCnt + 1
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        rsRS.MoveNext
    Wend
    Printer.EndDoc
    rsRS.MoveFirst
End If
End Sub

'Add By Sindy 2012/11/19
'組合延展多件清單
Private Function PrinterDetail_Text(rsRS As ADODB.Recordset) As String
Dim intCnt As Integer
Dim intLine As Integer
Dim dblLeft As Double
Dim dblTop As Double
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim intLineHeight As Integer
Dim intRow As Integer 'Add By Sindy 2013/1/22
Dim strTmp As String 'Add By Sindy 2013/1/22

PrinterDetail_Text = ""
intLine = 0
intCnt = 0
intLineHeight = 300
dblTop = 3500
If rsRS.RecordCount > 0 Then
   rsRS.MoveFirst
   While Not rsRS.EOF
      'Modify By Sindy 2015/7/1 +tm131
      StrSQLa = "SELECT tm01,tm02,tm03,tm04,tm05,tm09,tm15,tm22,tm45,tm131 FROM TRADEMARK WHERE " & ChgTradeMark(Replace("" & rsRS.Fields(4).Value, "-", ""))
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         'Modify By Sindy 2013/1/22
'         PrinterDetail_Text = PrinterDetail_Text & "" & rsA("TM05").Value & rsA("TM06").Value & rsA("TM07").Value & vbCrLf
'         PrinterDetail_Text = PrinterDetail_Text & IIf("" & rsA("TM15").Value = "", "          ", PUB_StrToStr(Trim("" & rsA("TM15").Value), 10, True))
'         PrinterDetail_Text = PrinterDetail_Text & IIf("" & rsA("TM09").Value = "", "            ", PUB_StrToStr(Trim("" & rsA("TM09").Value), 12, True))
'         PrinterDetail_Text = PrinterDetail_Text & IIf("" & rsA("TM22").Value = "", "                 ", PUB_StrToStr(ChgEngDate("" & rsA("TM22").Value), 17, True))
'         PrinterDetail_Text = PrinterDetail_Text & IIf("" & rsA("TM45").Value = "", "              ", PUB_StrToStr(Trim("" & rsA("TM45").Value), 14, True))
'         PrinterDetail_Text = PrinterDetail_Text & PUB_StrToStr(Trim(Replace("" & rsRS.Fields(4).Value, "-0-00", "")), 15, True) & vbCrLf
         
         'Add By Sindy 2013/8/15 不催延展者,不產生進度,不出定稿
         If PUB_ChkCaseIsNoticeScale(rsA.Fields("tm01"), rsA.Fields("tm02"), rsA.Fields("tm03"), rsA.Fields("tm04")) = False Then
            GoTo ReadNext
         End If
         '2013/8/15 END
         
         intRow = intRow + 1
         If Len(Trim("" & rsA.Fields("tm09"))) >= 11 Then
            PrinterDetail_Text = PrinterDetail_Text & "(" & intRow & ") Reg. No.：" & Trim("" & rsA.Fields("tm15")) & " "
            'modify by sonia 2020/6/11 Class+(es)
            PrinterDetail_Text = PrinterDetail_Text & "Class(es)：" & Trim("" & rsA.Fields("tm09")) & " " & vbCrLf
         Else
            PrinterDetail_Text = PrinterDetail_Text & "(" & intRow & ") Reg. No.：" & Left(("" & rsA.Fields("tm15")) & "            ", 12)
            'modify by sonia 2020/6/11 Class+(es)
            PrinterDetail_Text = PrinterDetail_Text & "Class(es)：" & Left(("" & rsA.Fields("tm09")) & "           ", 11) & vbCrLf
         End If
         '商標名稱
         'Modify By Sindy 2015/7/1 rsA.Fields("tm05") ==> IIf("" & rsA.Fields("tm131") <> "", rsA.Fields("tm131"), rsA.Fields("tm05"))
         'PrinterDetail_Text = PrinterDetail_Text & "   Trademark：" & rsA.Fields("tm05") & vbCrLf
         PrinterDetail_Text = PrinterDetail_Text & "   Trademark：" & IIf("" & rsA.Fields("tm131") <> "", rsA.Fields("tm131"), rsA.Fields("tm05")) & vbCrLf
         '2015/7/1 END
         '專用權止日
         strTmp = Empty
         If "" & rsA.Fields("tm22") = "" Then
            PrinterDetail_Text = PrinterDetail_Text & "   Expiry Date：" & vbCrLf
         Else
            PrinterDetail_Text = PrinterDetail_Text & "   Expiry Date：" & TranslateKeyWord(incCNV_ENGLISH_DATE, rsA.Fields("tm22"), strTmp) & vbCrLf
         End If
         '本所案號
         'Modify By Sindy 2014/9/10 + & vbCrLf 因為Your Ref要獨立放一行
         If rsA.Fields("tm03") = "0" And rsA.Fields("tm04") = "00" Then
            PrinterDetail_Text = PrinterDetail_Text & "   Our Ref：" & rsA.Fields("tm01") & "-" & rsA.Fields("tm02") & vbCrLf
         Else
            PrinterDetail_Text = PrinterDetail_Text & "   Our Ref：" & rsA.Fields("tm01") & "-" & rsA.Fields("tm02") & "-" & rsA.Fields("tm03") & "-" & rsA.Fields("tm04") & vbCrLf
         End If
         '彼所案號
         'PrinterDetail_Text = PrinterDetail_Text & "     Your Ref：" & rsA.Fields("tm45") & vbCrLf & vbCrLf
         PrinterDetail_Text = PrinterDetail_Text & "   Your Ref：" & rsA.Fields("tm45") & vbCrLf & vbCrLf
         '2013/1/22 End
      End If
ReadNext:
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      rsRS.MoveNext
   Wend
   rsRS.MoveFirst
End If
End Function

Private Sub PrintHead(dblTop As Double, intLine As Integer, intLineHeight As Integer)
    Printer.CurrentX = 4000 + 500
    Printer.CurrentY = dblTop + intLine * intLineHeight
    Printer.Print "S C H E D U L E"
    intLine = intLine + 1
    Printer.CurrentX = 4000 + 500
    Printer.CurrentY = dblTop + intLine * intLineHeight
    Printer.Print "============="
    intLine = intLine + 1
    Printer.CurrentX = 500
    Printer.CurrentY = dblTop + intLine * intLineHeight
    Printer.Print "TRADEMARK"
    intLine = intLine + 1
    
    Printer.CurrentX = 500
    Printer.CurrentY = dblTop + intLine * intLineHeight
    Printer.Print "Reg. NO."
    
    Printer.CurrentX = 500 + 2000 - 250
    Printer.CurrentY = dblTop + intLine * intLineHeight
    Printer.Print "CLASS"
    
    Printer.CurrentX = 500 + 3500 - 500
    Printer.CurrentY = dblTop + intLine * intLineHeight
    Printer.Print "EXPIRY DATE"
    
    Printer.CurrentX = 500 + 5250 '5500
    Printer.CurrentY = dblTop + intLine * intLineHeight
    Printer.Print "YOUR REF"
    
    Printer.CurrentX = 500 + 9000 '8500
    Printer.CurrentY = dblTop + intLine * intLineHeight
    Printer.Print "OUR REF"
    intLine = intLine + 1
    
    Printer.CurrentX = 500
    Printer.CurrentY = dblTop + intLine * intLineHeight
    Printer.Print "-----------------------------------------------------------------------------------------------------------------------------------------"
    intLine = intLine + 1
End Sub

Function GetNp22(oNP02 As String, oNP03 As String, oNP04 As String, oNP05 As String) As String
GetNp22 = ""
Dim rrtmp As New ADODB.Recordset
Dim oStrSQL As String
Set rrtmp = New ADODB.Recordset
oStrSQL = "select * from nextprogress where np02='" & oNP02 & "' and np03='" & oNP03 & "' and np04='" & oNP04 & "' and np05='" & oNP05 & "' and (NP06 IS NULL OR NP06='') and np07=716 "
rrtmp.CursorLocation = adUseClient
rrtmp.Open oStrSQL, cnnConnection, adOpenStatic, adLockReadOnly
If rrtmp.RecordCount <> 0 Then
    GetNp22 = CheckStr(rrtmp.Fields("np22"))
End If
Set rrtmp = Nothing
End Function

'Add By Sindy 2024/4/16
Private Sub Command2_Click(Index As Integer)
Dim strTmp As String
Dim ii As Integer
   
   If Index = 0 And Text3(2).Text <> "" Then
      strTmp = Text3(1) & Text3(2)
      If Text3(3).Text = "" Then
         Text3(3).Text = "0"
         strTmp = strTmp & "0"
      Else
         strTmp = strTmp & Text3(3).Text
      End If
      If Text3(4).Text = "" Then
         Text3(4).Text = "00"
         strTmp = strTmp & "00"
      Else
         strTmp = strTmp & Text3(4).Text
      End If
      intI = 1
      strExc(0) = "SELECT tm29 FROM trademark WHERE " & ChgTradeMark(strTmp)
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.Fields(0) = "Y" Then
            MsgBox "必須為未閉卷之案號，請重新輸入 !", vbCritical
         Else
            '檢查是否有案號重覆
            If List1(1).ListCount > 0 Then
               For ii = 0 To List1(1).ListCount - 1
                  If List1(1).List(ii) = Text3(1) & "-" & Text3(2) & "-" & Text3(3) & "-" & Text3(4) Then
                     MsgBox "重覆輸入案號 !", vbCritical
                     Text3(2).SetFocus
                     Exit Sub
                  End If
               Next ii
            End If
            '加入案號
            List1(1).AddItem Text3(1) & "-" & Text3(2) & "-" & Text3(3) & "-" & Text3(4)
            Text3(2).Text = ""
            Text3(3).Text = ""
            Text3(4).Text = ""
         End If
      Else
         MsgBox "案號不存在，請重新輸入 !", vbCritical
      End If
      Text3(2).SetFocus
   Else
      If List1(1).ListIndex > -1 Then
         List1(1).RemoveItem List1(1).ListIndex
      End If
   End If
End Sub
