VERSION 5.00
Begin VB.Form frm050312 
   BorderStyle     =   1  '單線固定
   Caption         =   "期限通知管制表"
   ClientHeight    =   6400
   ClientLeft      =   3050
   ClientTop       =   1510
   ClientWidth     =   4390
   ControlBox      =   0   'False
   LinkTopic       =   "Form11"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6400
   ScaleWidth      =   4390
   Begin VB.CheckBox Check1 
      Caption         =   "非大宗"
      Height          =   195
      Left            =   2415
      TabIndex        =   57
      Top             =   3810
      Width           =   1680
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   3525
      TabIndex        =   31
      Top             =   300
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   2745
      TabIndex        =   30
      Top             =   300
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "EPC年費金額(&D)"
      Height          =   400
      Index           =   0
      Left            =   1215
      TabIndex        =   29
      Top             =   300
      Width           =   1500
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   300
      Left            =   5925
      MaxLength       =   1
      TabIndex        =   34
      Top             =   5265
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Height          =   2325
      Left            =   240
      TabIndex        =   36
      Top             =   4020
      Width           =   3900
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   10
         Left            =   2745
         TabIndex        =   28
         Top             =   1950
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   6
         Left            =   1080
         TabIndex        =   27
         Top             =   1950
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   9
         Left            =   2745
         TabIndex        =   26
         Top             =   1650
         Width           =   810
      End
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   8
         Left            =   1080
         TabIndex        =   25
         Top             =   1650
         Width           =   810
      End
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   7
         Left            =   2745
         TabIndex        =   24
         Top             =   1320
         Width           =   810
      End
      Begin VB.OptionButton Option2 
         Caption         =   "法定期限："
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   50
         Top             =   429
         Width           =   1200
      End
      Begin VB.OptionButton Option2 
         Caption         =   "本所案號："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   125
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.TextBox txt1 
         Enabled         =   0   'False
         Height          =   264
         Index           =   11
         Left            =   1335
         MaxLength       =   7
         TabIndex        =   18
         Top             =   420
         Width           =   800
      End
      Begin VB.TextBox txt1 
         Enabled         =   0   'False
         Height          =   264
         Index           =   13
         Left            =   1335
         MaxLength       =   4
         TabIndex        =   20
         Top             =   720
         Width           =   800
      End
      Begin VB.TextBox txt1 
         Enabled         =   0   'False
         Height          =   264
         Index           =   14
         Left            =   2610
         MaxLength       =   4
         TabIndex        =   21
         Top             =   720
         Width           =   800
      End
      Begin VB.TextBox txt1 
         Enabled         =   0   'False
         Height          =   264
         Index           =   12
         Left            =   2610
         MaxLength       =   7
         TabIndex        =   19
         Top             =   420
         Width           =   800
      End
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   5
         Left            =   1080
         TabIndex        =   23
         Top             =   1320
         Width           =   810
      End
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   4
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   22
         Top             =   1020
         Width           =   810
      End
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   3
         Left            =   3195
         MaxLength       =   2
         TabIndex        =   17
         Top             =   120
         Width           =   375
      End
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   2
         Left            =   2835
         MaxLength       =   1
         TabIndex        =   16
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   1
         Left            =   1905
         MaxLength       =   6
         TabIndex        =   15
         Top             =   120
         Width           =   810
      End
      Begin VB.TextBox txt2 
         Height          =   264
         Index           =   0
         Left            =   1335
         MaxLength       =   3
         TabIndex        =   14
         Top             =   120
         Width           =   465
      End
      Begin VB.Label lblOverFee 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  '透明
         Caption         =   "超頁費："
         Height          =   180
         Index           =   1
         Left            =   2010
         TabIndex        =   61
         Top             =   1995
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblOverFee 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  '透明
         Caption         =   "超項費："
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   60
         Top             =   1995
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblMerge 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "點數："
         Height          =   180
         Index           =   1
         Left            =   2190
         TabIndex        =   54
         Top             =   1695
         Width           =   540
      End
      Begin VB.Label lblMerge 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  '透明
         Caption         =   "另一金額："
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   53
         Top             =   1695
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "點數："
         Height          =   180
         Left            =   2190
         TabIndex        =   51
         Top             =   1365
         Width           =   540
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "申請國家："
         Height          =   180
         Left            =   390
         TabIndex        =   48
         Top             =   762
         Width           =   900
      End
      Begin VB.Line Line8 
         X1              =   2250
         X2              =   2490
         Y1              =   545
         Y2              =   545
      End
      Begin VB.Line Line7 
         X1              =   2250
         X2              =   2490
         Y1              =   845
         Y2              =   845
      End
      Begin VB.Line Line6 
         X1              =   3279
         X2              =   1515
         Y1              =   245
         Y2              =   245
      End
      Begin VB.Label lblFee1 
         Caption         =   "金額："
         Height          =   210
         Left            =   120
         TabIndex        =   45
         Top             =   1347
         Width           =   540
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "lbl1"
         Height          =   180
         Left            =   1950
         TabIndex        =   44
         Top             =   1065
         Width           =   270
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "下一程序："
         Height          =   180
         Left            =   120
         TabIndex        =   43
         Top             =   1062
         Width           =   900
      End
      Begin VB.Label lblFee1s 
         BackColor       =   &H80000010&
         Height          =   180
         Left            =   240
         TabIndex        =   52
         Top             =   1420
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2925
      Left            =   240
      TabIndex        =   35
      Top             =   750
      Width           =   3900
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   16
         Left            =   1236
         MaxLength       =   9
         TabIndex        =   12
         Top             =   2178
         Width           =   800
      End
      Begin VB.CheckBox ChkPerson 
         Caption         =   "列印個人管制部門案件"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   1860
         Width           =   2325
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   15
         Left            =   750
         MaxLength       =   9
         TabIndex        =   13
         Text            =   "1"
         Top             =   2505
         Width           =   315
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   2
         Left            =   2520
         MaxLength       =   7
         TabIndex        =   2
         Top             =   408
         Width           =   800
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   4
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   4
         Top             =   696
         Width           =   800
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   6
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   6
         Top             =   984
         Width           =   800
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   8
         Left            =   2520
         MaxLength       =   9
         TabIndex        =   8
         Top             =   1272
         Width           =   1000
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   10
         Left            =   2520
         MaxLength       =   9
         TabIndex        =   10
         Top             =   1560
         Width           =   1000
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   9
         Left            =   1236
         MaxLength       =   9
         TabIndex        =   9
         Top             =   1560
         Width           =   1000
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   7
         Left            =   1236
         MaxLength       =   9
         TabIndex        =   7
         Top             =   1272
         Width           =   1000
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   5
         Left            =   1236
         MaxLength       =   4
         TabIndex        =   5
         Top             =   984
         Width           =   800
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   3
         Left            =   1236
         MaxLength       =   4
         TabIndex        =   3
         Top             =   696
         Width           =   800
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   1
         Left            =   1236
         MaxLength       =   7
         TabIndex        =   1
         Top             =   408
         Width           =   800
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   0
         Left            =   1236
         TabIndex        =   0
         Top             =   120
         Width           =   2085
      End
      Begin VB.Label lblSalesName 
         Caption         =   "lblSalesName"
         Height          =   225
         Left            =   2130
         TabIndex        =   59
         Top             =   2205
         Width           =   1185
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "管制人："
         Height          =   180
         Left            =   450
         TabIndex        =   58
         Top             =   2220
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "(1.本所案號 2.智權人員)"
         Height          =   180
         Left            =   1110
         TabIndex        =   56
         Top             =   2550
         Width           =   2055
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "排序："
         Height          =   180
         Left            =   150
         TabIndex        =   55
         Top             =   2550
         Width           =   540
      End
      Begin VB.Line Line5 
         X1              =   2310
         X2              =   2460
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line4 
         X1              =   2310
         X2              =   2460
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Line Line3 
         X1              =   2160
         X2              =   2400
         Y1              =   1109
         Y2              =   1109
      End
      Begin VB.Line Line2 
         X1              =   2160
         X2              =   2400
         Y1              =   821
         Y2              =   821
      End
      Begin VB.Line Line1 
         X1              =   2160
         X2              =   2400
         Y1              =   533
         Y2              =   533
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "代理人："
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   1602
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "申請人："
         Height          =   180
         Left            =   120
         TabIndex        =   41
         Top             =   1314
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "下一程序："
         Height          =   180
         Left            =   120
         TabIndex        =   40
         Top             =   1026
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "申請國家："
         Height          =   180
         Left            =   120
         TabIndex        =   39
         Top             =   738
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "法定期限："
         Height          =   180
         Left            =   120
         TabIndex        =   38
         Top             =   450
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "系統類別："
         Height          =   180
         Left            =   120
         TabIndex        =   37
         Top             =   120
         Width           =   900
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "定稿"
      Height          =   255
      Index           =   1
      Left            =   45
      TabIndex        =   33
      Top             =   3780
      Width           =   720
   End
   Begin VB.OptionButton Option1 
      Caption         =   "管制表"
      Height          =   195
      Index           =   0
      Left            =   30
      TabIndex        =   32
      Top             =   510
      Value           =   -1  'True
      Width           =   960
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "(1. 管制表 2. 定稿)"
      Height          =   180
      Left            =   6645
      TabIndex        =   47
      Top             =   5265
      Width           =   1425
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "列印格式:"
      Height          =   180
      Left            =   4860
      TabIndex        =   46
      Top             =   5280
      Width           =   765
   End
End
Attribute VB_Name = "frm050312"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo By Sindy 2010/12/7 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, strSQL2 As String, i As Integer, j As Integer, k As Integer, s As Integer, StrHave416 As String
'Modified by Lydia 2016/08/25
'Dim strTemp1 As Variant, strTemp2 As Variant, strTemp(0 To 14) As String, PLeft(0 To 14) As Integer, Page As Integer, iPrint As Integer
Dim strTemp1 As Variant, strTemp2 As Variant, Page As Integer, iPrint As Integer
'92.1.10 ADD BY SONIA
Dim strTmp As String, m_NP08 As String, m_NP09 As String, m_st02 As String
'92.3.21 MODIFY BY SONIA
Public m_PA21 As String, m_PA08 As String, m_PA09 As String, m_PA91 As String
'92.7.14 ADD BY SONIA
Dim m_PA10 As String, m_PA46 As String
'92.5.4 add by sonia
Dim m_Nexttimes As String
Dim m_NP02 As String, m_NP03 As String, m_NP04 As String, m_NP05 As String
'92.5.4 END
Const ET01 As String = "10"
'92.1.10 END
'Add By Cheng 2002/09/16
Dim blnClkSure As Boolean '判斷是否按下確定按鈕
Dim m_PA12 As String 'Add by Morgan 2005/4/18
Dim m_LC03 As String 'Add by Morgan 2008/5/1
Dim m_PA26 As String
'Add by Morgan 2008/5/30
Dim pa(1 To 4) As String '本所案號
Dim m_dblYear As Double '下次繳費年度
Dim m_strYear As String 'Add By Sindy 2009/07/30
'Added by Lydia 2016/01/04
Dim bolMerge As Boolean '兩個催函是否合併
Dim mNP07_2(0 To 1) As String '被合併的案件性質
Dim min_NP08 As String '最早所限
Dim feeTit1 As String, feeTit2 As String '費用名稱
Dim mNP09_1 As String, mNP09_2 As String '合併催函的各自法限
Dim mNP01_2 As String, mNP08_2 As String, mNP22_2 As String, mCP09_2 As String 'Added by Morgan 2021/10/28
Dim cp() As String 'Add By Sindy 2016/6/20
Dim m_PA25 As String 'Added by Morgan 2016/8/9 專用期止日
'Added by Lydia 2016/08/25
Private Const ciStartX = 400, ciStartY = 500, ciColGap = 100
Private Const ciTitleFontSize = 22 '報表抬頭字型大小
'Modified by Morgan 2023/10/17
'Private Const ciFontSize = 8  '報表內容字型大小
Private Const ciFontSize = 10  '報表內容字型大小
'end 2023/10/17
Private Const cInX = 14
Dim strTitle As String, strTitle2 As String '欄位抬頭/起始位置
Dim PTitle(0 To cInX) As String '欄位抬頭陣列
Dim strTemp(0 To cInX) As String, PLeft(0 To cInX) As Integer
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim strChkEntity As String 'Added by Lydia 2016/09/13 是否已檢查個體
Public m_InputEPC As Boolean 'Added by Lydia 2016/09/29
Dim strPA16 As String  'Added by Lydia 2016/11/11 目前案件准駁
Dim m_LD18 As String 'Added by Morgan 2017/9/5
'Add By Sindy 2017/12/29
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String, m_AppNo As String
Dim m_Done As Boolean
'2017/12/29 END
Dim m_PA179 As String 'Added by Morgan 2023/3/25

Private Sub cmdok_Click(Index As Integer)
Dim bolCancel As Boolean 'Add By Sindy 2016/6/20
ReDim cp(TF_CP) As String 'Add By Sindy 2016/6/20

Select Case Index
Case 0 'EPC年費金額
     '定稿
     If Option1(1).Value = True Then
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/6 清除查詢印表記錄檔欄位
         pub_QL05 = pub_QL05 & ";" & Option1(1).Caption 'Add By Sindy 2010/12/6
         
         'Modify By Cheng 2002/10/11
         If Me.Option2(0).Value Then '若選擇本所案號
            If Len(Trim(txt2(0))) = 0 Or Len(Trim(txt2(1))) = 0 Then
               s = MsgBox("本所案號不可空白!!", , "USER 輸入錯誤")
               txt2(0).SetFocus
               txt2_GotFocus (0)
               Exit Sub
            End If
         Else '若選擇法定期限及申請國家
            If PUB_CheckKeyInDate(Me.txt1(11)) = -1 Then
               Me.txt1(11).SetFocus
               txt1_GotFocus 11
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(12)) = -1 Then
               Me.txt1(12).SetFocus
               txt1_GotFocus 12
               Exit Sub
            End If
            If Me.txt1(11).Text <> "" And Me.txt1(12).Text <> "" Then
               If Val(Me.txt1(11).Text) > Val(Me.txt1(12).Text) Then
                    'Modify By Cheng 2002/12/12
'                  MsgBox "本所期限範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  MsgBox "法定期限範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(11).SetFocus
                  txt1_GotFocus 11
                  Exit Sub
               End If
            End If
            If Len(Trim(txt1(12))) = 0 Then
                'Modify By Cheng 2002/12/12
'                 s = MsgBox("本所期限不可空白!!", , "USER 輸入錯誤")
                 s = MsgBox("法定期限不可空白!!", , "USER 輸入錯誤")
                  txt1(11).SetFocus
                  txt1_GotFocus (11)
                 Exit Sub
            Else
               'Add By Cheng 2002/09/16
               If Me.txt1(13).Text <> "" And Me.txt1(14).Text <> "" Then
                  If Me.txt1(13).Text > Me.txt1(14).Text Then
                     MsgBox "申請國家範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(13).SetFocus
                     txt1_GotFocus 13
                     Exit Sub
                  End If
               End If
            End If
         End If
         
         'Add By Sindy 2017/12/29
         If m_strIR01 <> "" Then
            If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> txt2(0) & txt2(1) & txt2(2) & txt2(3) Then
               MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
               Exit Sub
            End If
         End If
         '2017/12/29 END
         
         If Len(Trim(txt2(4))) = 0 Then
            s = MsgBox("下一程序不可空白!!", , "USER 輸入錯誤")
            txt2(4).SetFocus
            Exit Sub
         Else
            'Add By Cheng 2002/09/16
            lbl1 = GetPrjState6HM("CFP", txt2(4))
            ChkAlert 'Added by Morgan 2025/3/13
            If Me.txt2(4).Text <> "" Then
               If Me.lbl1.Caption = "" Then
                  Me.txt2(4).SetFocus
                  txt2_GotFocus 4
                  Exit Sub
               End If
            End If
            
            If Val(txt2(5)) = 0 Then
               s = MsgBox("金額不可為 0 或 空白！！", , "USER 輸入錯誤")
               txt2(5).SetFocus
               txt2(5).SelStart = 0
               txt2(5).SelLength = Len(txt2(5))
               Exit Sub
            End If

         End If
         
         Screen.MousePointer = vbHourglass
         strSQL1 = ""
         'Modify By Cheng 2002/10/11
         If Me.Option2(0).Value Then '若選擇本所案號
            pub_QL05 = pub_QL05 & ";" & Option2(0).Caption & txt2(0) & "-" & txt2(1) 'Add By Sindy 2010/12/6
            If Len(Trim(txt2(0))) <> 0 Then
               strSQL1 = strSQL1 & " and np02='" & txt2(0) & "' "
            End If
            If Len(Trim(txt2(1))) <> 0 Then
               strSQL1 = strSQL1 & " and np03='" & txt2(1) & "' "
            End If
            If Len(Trim(txt2(2))) <> 0 Then
               strSQL1 = strSQL1 & " and np04='" & txt2(2) & "' "
               pub_QL05 = pub_QL05 & "-" & txt2(2) 'Add By Sindy 2010/12/6
            Else
               strSQL1 = strSQL1 & " and np04='0' "
            End If
            If Len(Trim(txt2(3))) <> 0 Then
               strSQL1 = strSQL1 & " and np05='" & txt2(3) & "' "
               pub_QL05 = pub_QL05 & "-" & txt2(3) 'Add By Sindy 2010/12/6
            Else
               strSQL1 = strSQL1 & " and np05='00' "
            End If
         Else '若選擇法定期限及申請國家
            If Len(Trim(txt1(11))) <> 0 Then
                'Modify By Cheng 2002/12/13
'               strSQL1 = strSQL1 & " and np08>=" & txt1(11) & " "
               strSQL1 = strSQL1 & " and np09>=" & txt1(11) & " "
            End If
            If Len(Trim(txt1(12))) <> 0 Then
                'Modify By Cheng 2002/12/13
'               strSQL1 = strSQL1 & " and np08<=" & txt1(12) & " "
               strSQL1 = strSQL1 & " and np09<=" & txt1(12) & " "
            End If
            If Len(Trim(txt1(11))) <> 0 Or Len(Trim(txt1(12))) <> 0 Then
               pub_QL05 = pub_QL05 & ";" & Option2(1).Caption & txt1(11) & "-" & txt1(12) 'Add By Sindy 2010/12/6
            End If
            If Len(Trim(txt1(13))) <> 0 Then
               strSQL1 = strSQL1 & " and pa09>='" & txt1(13) & "' "
            End If
            If Len(Trim(txt1(14))) <> 0 Then
               strSQL1 = strSQL1 & " and pa09<='" & txt1(14) & "' "
            End If
            If Len(Trim(txt1(13))) <> 0 Or Len(Trim(txt1(14))) <> 0 Then
               pub_QL05 = pub_QL05 & ";" & Label14 & txt1(13) & "-" & txt1(14) 'Add By Sindy 2010/12/6
            End If
         End If
         If Len(Trim(txt2(4))) <> 0 Then
            strSQL1 = strSQL1 & " and np07=" & Val(txt2(4)) & " "
            pub_QL05 = pub_QL05 & ";" & Label8 & txt2(4) & "-" & lbl1 'Add By Sindy 2010/12/6
         End If
         strSQL1 = strSQL1 & " and np06 is null "
         'Modify by Morgan 2008/5/15
         'strSQL = "select pa09,np07,np01,np22 from nextprogress,patent where np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) and np02 in (" & SQLGrpStr("", 1) & ") " & strSQL1
         'Modified by Morgan 2016/8/10 +pa25
         strSql = "select pa09,np07,np02,np03,np04,np05,np22,np08,np09,pa21,a0902,st02,pa08,PA91,PA10,PA46,PA12,np01,pa25 from nextprogress,patent,staff,acc090 " & _
                  "where np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) and np10=st01(+) and st15=a0901(+) and np02 in (" & SQLGrpStr("", 1) & ") " & strSQL1
         CheckOC
         With adoRecordset
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If .RecordCount <> 0 Then
               InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/6
               Me.Hide
               frm050312_1.Show
               frm050312_1.lbl1.Caption = txt2(5)
               frm050312_1.Lbl3.Caption = txt2(7)
               'Add by Morgan 2008/5/15
               frm050312_1.m_NP01 = "" & .Fields("np01")
               frm050312_1.m_NP22 = "" & .Fields("np22")
               frm050312_1.m_st02 = "" & .Fields(10) & "　" & .Fields(11)
               frm050312_1.m_NP08 = "" & .Fields("np08")
               frm050312_1.m_NP09 = "" & .Fields("np09")
               frm050312_1.m_PA25 = "" & .Fields("pa25") 'Added by Morgan 2016/8/10
               'end 2008/5/15
            Else
               InsertQueryLog (0) 'Add By Sindy 2010/12/6
               'Modify By Cheng 2002/10/11
               If Me.Option2(0).Value Then
                  s = MsgBox("此 " & txt2(0) & "-" & txt2(1) & IIf(Len(txt2(2)) <> 0, "-" & txt2(0), "") & IIf(Len(txt2(3)) <> 0, IIf(Len(txt2(2)) <> 0, "-" & txt2(3), "--" & txt2(3)), "") & "，找不到下一程序 " & txt2(4) & " 的資料 或 此案之申請國家非EPC ！！", , "沒有資料")
               Else
                    'Modify By Cheng 2002/12/12
'                  s = MsgBox("此本所期限及申請國家範圍，找不到下一程序 " & txt2(4) & " 的資料 或 此範圍之申請國家非EPC ！！", , "沒有資料")
                  s = MsgBox("此法定期限及申請國家範圍，找不到下一程序 " & txt2(4) & " 的資料 或 此範圍之申請國家非EPC ！！", , "沒有資料")
               End If
               Screen.MousePointer = vbDefault
               CheckOC
               Exit Sub
            End If
         End With
         Screen.MousePointer = vbDefault
         CheckOC
     Else
         s = MsgBox("請選擇定稿！！", , "列印別錯誤")
     End If
Case 1 '確定
      'Add By Cheng 2002/09/16
      blnClkSure = False
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/6 清除查詢印表記錄檔欄位
      'Add By Sindy 2016/6/20
      bolCancel = False
      Call txt1_Validate(15, bolCancel)
      If bolCancel = True Then Exit Sub
      '2016/6/20 END
      
        '管制表
        If Option1(0).Value = True Then
            If Len(Trim(txt1(0))) = 0 Then
                s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
                txt1(0).SetFocus
                txt1(0).SelStart = 0
                txt1(0).SelLength = Len(txt1(0))
                Exit Sub
            Else
               'Add By Cheng 2002/03/20
               If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
                  Me.txt1(1).SetFocus
                  txt1_GotFocus 1
                  Exit Sub
               End If
               If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
                  Me.txt1(2).SetFocus
                  txt1_GotFocus 2
                  Exit Sub
               End If
               'Add By Cheng 2002/09/16
               If Me.txt1(1).Text <> "" And Me.txt1(2).Text <> "" Then
                  If Val(Me.txt1(1).Text) > Val(Me.txt1(2).Text) Then
                    'Modify By Cheng 2002/12/12
'                     MsgBox "本所期限範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     MsgBox "法定期限範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(1).SetFocus
                     txt1_GotFocus 1
                     Exit Sub
                  End If
               'Added by Lydia 2019/08/30
               Else
                     MsgBox "法定期限範圍不可空白!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     If Me.txt1(1) = "" Then
                        Me.txt1(1).SetFocus
                        txt1_GotFocus 1
                     Else
                        Me.txt1(2).SetFocus
                        txt1_GotFocus 2
                     End If
                     Exit Sub
               'end 2019/08/30
               End If
                If Len(Trim(txt1(2))) = 0 Then
                    'Modify By Cheng 2002/12/12
'                    s = MsgBox("本所期限不可空白!!", , "USER 輸入錯誤")
                    s = MsgBox("法定期限不可空白!!", , "USER 輸入錯誤")
                     txt1(1).SetFocus
                     txt1_GotFocus (1)
                    Exit Sub
                Else
                  'Add By Cheng 2002/09/16
                  If Me.txt1(3).Text <> "" And Me.txt1(4).Text <> "" Then
                     If Me.txt1(3).Text > Me.txt1(4).Text Then
                        MsgBox "申請國家範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                        blnClkSure = True
                        Me.txt1(3).SetFocus
                        txt1_GotFocus 3
                        Exit Sub
                     End If
                  End If
                  If Me.txt1(5).Text <> "" And Me.txt1(6).Text <> "" Then
                     If Me.txt1(5).Text > Me.txt1(6).Text Then
                        MsgBox "下一程序範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                        blnClkSure = True
                        Me.txt1(5).SetFocus
                        txt1_GotFocus 5
                        Exit Sub
                     End If
                  End If
                    
                    If Len(Trim(txt1(7))) <> 0 And Len(Trim(txt1(8))) <> 0 Then
                        If Left(txt1(7), 6) <> Left(txt1(8), 6) Then
                            s = MsgBox("申請人前六碼必須相同!!", , "USER 輸入錯誤")
                            blnClkSure = True
                            txt1(7).SetFocus
                            txt1(7).SelStart = 0
                            txt1(7).SelLength = Len(txt1(7))
                            Exit Sub
                        End If
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
                    If Len(Trim(txt1(9))) <> 0 And Len(Trim(txt1(10))) <> 0 Then
                        If Left(txt1(9), 6) <> Left(txt1(10), 6) Then
                            s = MsgBox("代理人前六碼必須相同!!", , "USER 輸入錯誤")
                            blnClkSure = True
                            txt1(9).SetFocus
                            txt1(9).SelStart = 0
                            txt1(9).SelLength = Len(txt1(9))
                            Exit Sub
                        End If
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
                  'Added by Lydia 2020/03/05
                  If ChkPerson.Value = 1 And (txt1(16).Text = "" Or lblSalesName.Caption = "") Then
                        MsgBox "請輸入管制人員代號!!!", vbExclamation + vbOKOnly
                        blnClkSure = True
                        Me.txt1(16).SetFocus
                        txt1_GotFocus 16
                        Exit Sub
                  End If
                  'end 2020/03/05
                  
                    Screen.MousePointer = vbHourglass
                    Me.Enabled = False
                    pub_QL05 = pub_QL05 & ";" & Option1(0).Caption 'Add By Sindy 2010/12/6
                    Process
                    Me.Enabled = True
                    Screen.MousePointer = vbDefault
                End If
            End If
        '定稿
        Else
            'Modify By Cheng 2002/10/11
            If Me.Option2(0).Value Then '若選擇本所案號
               If Len(Trim(txt2(0))) = 0 Or Len(Trim(txt2(1))) = 0 Then
                  s = MsgBox("本所案號不可空白!!", , "USER 輸入錯誤")
                  txt2(0).SetFocus
                  txt2_GotFocus (0)
                  Exit Sub
               End If
            Else '若選擇法定期限及申請國家
               If PUB_CheckKeyInDate(Me.txt1(11)) = -1 Then
                  Me.txt1(11).SetFocus
                  txt1_GotFocus 11
                  Exit Sub
               End If
               If PUB_CheckKeyInDate(Me.txt1(12)) = -1 Then
                  Me.txt1(12).SetFocus
                  txt1_GotFocus 12
                  Exit Sub
               End If
               If Me.txt1(11).Text <> "" And Me.txt1(12).Text <> "" Then
                  If Val(Me.txt1(11).Text) > Val(Me.txt1(12).Text) Then
                        'Modify By Cheng 2002/12/12
'                     MsgBox "本所期限範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     MsgBox "法定期限範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(11).SetFocus
                     txt1_GotFocus 11
                     Exit Sub
                  End If
               End If
               If Len(Trim(txt1(12))) = 0 Then
                    'Modify By Cheng 2002/12/12
'                    s = MsgBox("本所期限不可空白!!", , "USER 輸入錯誤")
                    s = MsgBox("法定期限不可空白!!", , "USER 輸入錯誤")
                     txt1(11).SetFocus
                     txt1_GotFocus (11)
                    Exit Sub
               Else
                  'Add By Cheng 2002/09/16
                  If Me.txt1(13).Text <> "" And Me.txt1(14).Text <> "" Then
                     If Me.txt1(13).Text > Me.txt1(14).Text Then
                        MsgBox "申請國家範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                        blnClkSure = True
                        Me.txt1(13).SetFocus
                        txt1_GotFocus 13
                        Exit Sub
                     End If
                  End If
               End If
            End If
            
            'Add By Sindy 2017/12/29
            If m_strIR01 <> "" Then
               If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> txt2(0) & txt2(1) & txt2(2) & txt2(3) Then
                  MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
                  Exit Sub
               End If
            End If
            '2017/12/29 END
            
            If Len(Trim(txt2(4))) = 0 Then
               s = MsgBox("下一程序不可空白!!", , "USER 輸入錯誤")
               txt2(4).SetFocus
               txt2(4).SelStart = 0
               txt2(4).SelLength = Len(txt2(4))
               Exit Sub
            Else
               'Add By Cheng 2002/09/16
               lbl1 = GetPrjState6HM("CFP", txt2(4))
               ChkAlert 'Added by Morgan 2025/3/13
               If Me.txt2(4).Text <> "" Then
                  If Me.lbl1.Caption = "" Then
                     Me.txt2(4).SetFocus
                     txt2_GotFocus 4
                     Exit Sub
                  End If
               End If
               
               'Add by Morgan 2010/5/31 PCT進各國期限不用
               'Modify by Morgan 2010/6/11 +209
               If Trim(txt2(4)) <> "119" And Trim(txt2(4)) <> "910" Then
               'end 2010/5/31
                  If Val(txt2(5)) = 0 And Me.Option2(0).Value Then
                     s = MsgBox("金額不可為 0 或 空白！！", , "USER 輸入錯誤")
                     txt2(5).SetFocus
                     txt2(5).SelStart = 0
                     txt2(5).SelLength = Len(txt2(5))
                     Exit Sub
                  End If
                  
                  'Added by Morgan 2023/5/12
                  '游本俊〈X75231000〉及游翊〈X75231010〉的CFP案件之領證費及年費在程序人員報價後，系統請另外再多加服務費NT$1,000並跳訊息告知操作人員。
                  'Modified by Morgan 2023/10/18 +606,607 --禧佩
                  'Modified by Morgan 2024/9/10 + And m_PA26 <> ""
                  If txt2(0) = "CFP" And (txt2(4) = "605" Or txt2(4) = "606" Or txt2(4) = "607") And m_PA26 <> "" Then
                     If InStr("X75231000,X75231010", m_PA26) > 0 Then
                        If MsgBox("客戶游本俊〈X75231000〉及游翊〈X75231010〉因長年在中國大陸，所以委辦案件付款都是由本所收據金額直接換算當時美金在匯至本所華南銀行。近日發現華南銀行扣除手續費後會有虧損情況，故此兩客戶的CFP案件之領證費及年費請" & vbCrLf & vbCrLf & "另外再多加 NT$1,000(0點)，本次報價是否已調整？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                           Exit Sub
                        End If
                        
                     'Added by Morgan 2023/11/16
                     ElseIf InStr("X05659010", m_PA26) > 0 Then
                        If MsgBox("此客戶的維持費、年費或延展費為特殊報價需再多加 NT$2,000(2點)，本次報價是否已調整？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                           Exit Sub
                        End If
                     'end 2023/11/16
                     End If
                  End If
                  'end 2023/5/12
               End If
            End If
            
            'Added by Lydia 2016/01/04
            If lblMerge(0).Visible = True Then
               'Added by Morgan 2020/11/16
               If m_PA09 = "239" Then
                  If Val(txt2(8)) = 0 Then
                     s = MsgBox("英國金額不可為 0 或 空白！！", , "USER 輸入錯誤")
                     txt2(8).SetFocus
                     txt2(8).SelStart = 0
                     txt2(8).SelLength = Len(txt2(8))
                     Exit Sub
                  End If
               Else
               'end 2020/11/16
                  For intI = 8 To 9
                      If Val(txt2(intI)) = 0 Then
                          s = MsgBox("合併" & IIf(intI = 8, "金額", "點數") & "不可為 0 或 空白！！", , "USER 輸入錯誤")
                          txt2(intI).SetFocus
                          txt2(intI).SelStart = 0
                          txt2(intI).SelLength = Len(txt2(intI))
                          Exit Sub
                      End If
                  Next intI
               End If 'Added by Morgan 2020/11/16
            End If
            'end 2016/01/04
            
            'Added by Lydia 2016/09/29 EPC進入各國繳年費分別判斷
            'Modified by Lydia 2016/11/11 母案核准後才進入國家階段 + and strpa16 = "1"
            'Modified by Lydia 2017/02/13 進入國家階段以發證日為準 (BY 甄妮)
            'If m_PA09 = "221" And m_InputEPC = False And strPA16 = "1" Then
            If m_PA09 = "221" And m_InputEPC = False And m_PA21 <> "" Then
                s = MsgBox("請輸入正確的EPC年費金額", , "USER 輸入錯誤")
                cmdok(0).SetFocus
                Exit Sub
            End If
            
            'Added by Lydia 2016/08/31 CFP控制年費、延展費及維持費智權同仁可加的點數
            'Modified by Lydia 2016/09/29 +非EPC國
            'If (txt2(4) = "605" And Val(txt2(7)) > CFP_dg605) Or (txt2(4) = "606" And Val(txt2(7)) > CFP_dg606) Or (txt2(4) = "607" And Val(txt2(7)) > CFP_dg607) Then
            'Modified by Lydia 2016/11/11 EPC母案核准後才進入國家階段,尚未核准時以一般個案判斷
            'If m_PA09 <> "221" And ((txt2(4) = "605" And Val(txt2(7)) > CFP_dg605) Or (txt2(4) = "606" And Val(txt2(7)) > CFP_dg606) Or (txt2(4) = "607" And Val(txt2(7)) > CFP_dg607)) Then
            'Modified by Lydia 2017/02/13 進入國家階段以發證日為準 (BY 甄妮)
            'If (m_PA09 <> "221" Or (m_PA09 = "221" And strPA16 <> "1")) And ((txt2(4) = "605" And Val(txt2(7)) > CFP_dg605) Or (txt2(4) = "606" And Val(txt2(7)) > CFP_dg606) Or (txt2(4) = "607" And Val(txt2(7)) > CFP_dg607)) Then
            If (m_PA09 <> "221" Or (m_PA09 = "221" And m_PA21 = "")) And ((txt2(4) = "605" And Val(txt2(7)) > CFP_dg605) Or (txt2(4) = "606" And Val(txt2(7)) > CFP_dg606) Or (txt2(4) = "607" And Val(txt2(7)) > CFP_dg607)) Then
                If MsgBox("請注意！年費(605)超過" & CFP_dg605 & "點，維持費(606)及延展費(607)超過" & CFP_dg606 & "點，是否確定？", vbYesNo + vbDefaultButton2, "控制點數") = vbNo Then
                   txt2(7).SetFocus
                   txt2_GotFocus 7
                   Exit Sub
                End If
            End If
            If txt2(8).Visible = True And ((txt2(8).Tag = "605" And Val(txt2(9)) > CFP_dg605) Or (txt2(8).Tag = "606" And Val(txt2(9)) > CFP_dg606) Or (txt2(8).Tag = "607" And Val(txt2(9)) > CFP_dg607)) Then
                If MsgBox("請注意！年費(605)超過" & CFP_dg605 & "點，維持費(606)及延展費(607)超過" & CFP_dg606 & "點，是否確定？", vbYesNo + vbDefaultButton2, "控制點數") = vbNo Then
                   txt2(9).SetFocus
                   txt2_GotFocus 9
                   Exit Sub
                End If
            End If
            'end 2016/08/31
            
            
            'Added by Morgan 2020/3/17
            If txt2(6).Visible And txt2(6) = "" Then
               s = MsgBox("請輸入超項費！！" & vbCrLf & vbCrLf & "若沒有請輸入 0 ！", vbExclamation, "USER 輸入錯誤")
               txt2_GotFocus 6
               txt2(6).SetFocus
               Exit Sub
            End If
            If txt2(10).Visible And txt2(10) = "" Then
               s = MsgBox("請輸入超頁費！！" & vbCrLf & vbCrLf & "若沒有請輸入 0 ！", vbExclamation, "USER 輸入錯誤")
               txt2_GotFocus 10
               txt2(10).SetFocus
               Exit Sub
            End If
            'end 2020/3/17
            
            'Add by Amy 2018/03/20 CFP案若有智權及承辦人員相同且為P12部門之未發文未取消收文之B類911補收款資料彈訊息
            If Option2(0).Value = True And txt2(0) = "CFP" Then
                If Pub_B911NotPay(txt2(0), txt2(1), IIf(Trim(txt2(2)) = "", "0", txt2(2)), IIf(Trim(txt2(3)) = "", "0", txt2(3))) = True Then
                    MsgBox "此案有未收款！", vbExclamation
                End If
            End If
            'end 2018/03/20
            Screen.MousePointer = vbHourglass
            Me.Enabled = False
            pub_QL05 = pub_QL05 & ";" & Option1(1).Caption 'Add By Sindy 2010/12/6
            ProcessToWord
            Me.Enabled = True
            Screen.MousePointer = vbDefault
            
            'Add By Sindy 2017/12/29
            If m_strIR01 <> "" Then
               'Modify By Sindy 2022/5/20
               'frm04010519.GoNext
               Forms(0).Tmpfrm04010519.GoNext
               Set Forms(0).Tmpfrm04010519 = Nothing
               '2022/5/20 END
               Unload Me
            End If
            '2017/12/29 END
        End If
Case 2 '結束
     Unload Me
Case Else
End Select
End Sub

'列印定稿
Private Sub ProcessToWord()
Dim oText As TextBox 'Added by Lydia 2016/01/04
Dim bolInTran As Boolean, strErrMsg As String 'Added by Morgan 2021/10/29

Screen.MousePointer = vbHourglass
strSQL1 = ""
'Modify By Cheng 2002/10/11
If Me.Option2(0).Value Then
   pub_QL05 = pub_QL05 & ";" & Option2(0).Caption & txt2(0) & "-" & txt2(1) 'Add By Sindy 2010/12/6
   If Len(Trim(txt2(0))) <> 0 Then
      strSQL1 = strSQL1 & " and np02='" & txt2(0) & "' "
   End If
   If Len(Trim(txt2(1))) <> 0 Then
      strSQL1 = strSQL1 & " and np03='" & txt2(1) & "' "
   End If
   If Len(Trim(txt2(2))) <> 0 Then
      strSQL1 = strSQL1 & " and np04='" & txt2(2) & "' "
      pub_QL05 = pub_QL05 & "-" & txt2(2) 'Add By Sindy 2010/12/6
   Else
      strSQL1 = strSQL1 & " and np04='0' "
   End If
   If Len(Trim(txt2(3))) <> 0 Then
      strSQL1 = strSQL1 & " and np05='" & txt2(3) & "' "
      pub_QL05 = pub_QL05 & "-" & txt2(3) 'Add By Sindy 2010/12/6
   Else
      strSQL1 = strSQL1 & " and np05='00' "
   End If
   'Added by Morgan 2017/12/19 單筆只抓未到期期限
   strSQL1 = strSQL1 & " and np09>=" & strSrvDate(1)
Else
   If Len(Trim(txt1(11))) <> 0 Then
        'Modify By Cheng 2002/12/13
'      strSQL1 = strSQL1 & " and np08>=" & txt1(11) & " "
      strSQL1 = strSQL1 & " and np09>=" & DBDATE(txt1(11)) & " "
   End If
   If Len(Trim(txt1(12))) <> 0 Then
        'Modify By Cheng 2002/12/13
'      strSQL1 = strSQL1 & " and np08<=" & txt1(12) & " "
      strSQL1 = strSQL1 & " and np09<=" & DBDATE(txt1(12)) & " "
   End If
   If Len(Trim(txt1(11))) <> 0 Or Len(Trim(txt1(12))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Option2(1).Caption & txt1(11) & "-" & txt1(12) 'Add By Sindy 2010/12/6
   End If
   If Len(Trim(txt1(13))) <> 0 Then
      strSQL1 = strSQL1 & " and pa09>='" & txt1(13) & "' "
   End If
   If Len(Trim(txt1(14))) <> 0 Then
      strSQL1 = strSQL1 & " and pa09<='" & txt1(14) & "' "
   End If
   If Len(Trim(txt1(13))) <> 0 Or Len(Trim(txt1(14))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label14 & txt1(13) & "-" & txt1(14) 'Add By Sindy 2010/12/6
   End If
End If
If Len(Trim(txt2(4))) <> 0 Then
   strSQL1 = strSQL1 & " and np07=" & Val(txt2(4)) & " "
   pub_QL05 = pub_QL05 & ";" & Label8 & txt2(4) & lbl1 'Add By Sindy 2010/12/6
End If
'Added by Lydia 2016/01/04
If Len(Trim(txt2(8))) <> 0 Then
   pub_QL05 = pub_QL05 & ";合併催函"
End If

'Modify by Morgan 2006/6/27 加閉卷不要
'strSQL1 = strSQL1 & " and np06 is null "
strSQL1 = strSQL1 & " and np06 is null and pa57 is null"

'92.1.10 MODIFY BY SONIA 為列印接洽結案單而加欄位
'strSQL = "select pa09,np07,np01 from nextprogress,patent where np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) and np02 in (" & SQLGrpStr("", 1) & ") " & strSQL1
'Modified by Morgan 2016/8/9 +pa25
'Modified by Morgan 2023/3/25 +PA179
strSql = "select pa09,np07,np02,np03,np04,np05,np22,np08,np09,pa21,a0902,st02,pa08,PA91,PA10,PA46,PA12,np01,pa25,pa26,pa75,pa179 from nextprogress,patent,staff,acc090 where np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) and np10=st01(+) and st15=a0901(+) and np02 in (" & SQLGrpStr("", 1) & ") " & strSQL1 & " order by np09 asc"
'92.1.10 END
CheckOC
With adoRecordset
   .CursorLocation = adUseClient
   .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
      InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/6
      Do While .EOF = False
         m_NP08 = .Fields(7): m_NP09 = .Fields(8): m_PA21 = "": m_PA08 = "": m_PA91 = "": m_PA10 = "": m_PA46 = "": m_st02 = ""
         If .Fields(2) <> "" Then m_NP02 = .Fields(2)
         If .Fields(3) <> "" Then m_NP03 = .Fields(3)
         If .Fields(4) <> "" Then m_NP04 = .Fields(4)
         If .Fields(5) <> "" Then m_NP05 = .Fields(5)
         If .Fields(9) <> "" Then m_PA21 = .Fields(9)
         If .Fields(12) <> "" Then m_PA08 = .Fields(12)
         If .Fields(13) <> "" Then m_PA91 = .Fields(13)
         If .Fields(10) <> "" Then m_st02 = .Fields(10) & "　" & .Fields(11)
         If .Fields(14) <> "" Then m_PA10 = .Fields(14)
         If .Fields(15) <> "" Then m_PA46 = .Fields(15)
         m_PA25 = "" & .Fields("pa25") 'Added by Morgan 2016/8/9
         m_LD18 = "" 'Added by Morgan 2017/9/5
        'Add by Lydia 2014/11/24 加判斷未收文恢復權利414
    If CheckCPExists(m_NP02, m_NP03, m_NP04, m_NP05) = False Then
    
'Added by Morgan 2021/10/29
On Error GoTo ErrHnd
         cnnConnection.BeginTrans
         bolInTran = True
'end 2021/10/29

         'Added by Morgan 2013/8/7
         'Modified by Morgan 2016/1/13 改轉定稿時上發文日且要在定稿產生前新增,否則報價即時轉定稿的狀況會沒有進度可上發文日
         'Modified by Morgan 2017/9/5 +信函進度相關欄位
         'Modified by Morgan 2018/10/25 +傳入是否大宗發文參數
         'Modified by Morgan 2021/10/29 +pbolInTrans=True
         If PUB_AddCP1913(.Fields("np02"), .Fields("np03"), .Fields("np04"), .Fields("np05"), .Fields("NP08"), .Fields("NP09"), .Fields("NP01"), .Fields("NP22"), "" & .Fields("pa09"), "" & .Fields("pa26"), m_LD18, "" & .Fields("pa75"), True, True, IIf(Check1.Value = vbChecked, False, True)) = False Then
            'Modified by Morgan 2021/10/29
            'MsgBox "新增進度檔【通知期限】失敗！作業中斷！", vbCritical
            'Exit Sub
            strErrMsg = "新增進度檔【通知期限】失敗！作業中斷！"
            GoTo ErrHnd
            'end 2021/10/29
         End If
         'end 2013/8/7
         
         'Added by Morgan 2021/10/29
         '合併催函也要新增進度
         If lblMerge(0).Visible = True Then
            If PUB_AddCP1913(.Fields("np02"), .Fields("np03"), .Fields("np04"), .Fields("np05"), mNP08_2, mNP09_2, mNP01_2, mNP22_2, "" & .Fields("pa09"), "" & .Fields("pa26"), mCP09_2, "" & .Fields("pa75"), True, True, IIf(Check1.Value = vbChecked, False, True)) = False Then
               strErrMsg = "新增進度檔【通知期限】失敗！作業中斷！"
               GoTo ErrHnd
            Else
               strExc(0) = "已併入通知期限-" & lbl1 & "通知函(" & m_LD18 & ")告知客戶;"
               strSql = "update letterprogress set lp06='" & strUserNum & "',lp07=to_char(sysdate,'yyyymmdd'),lp10='N',lp12='" & strExc(0) & "'||lp12,lp42='" & m_LD18 & "'" & _
                  " where lp01='" & mCP09_2 & "'"
               cnnConnection.Execute strSql, intI
            End If
         End If
         'end 2021/10/28
         
         'Add by Sindy 2017/12/29
         If m_strIR01 <> "" Then
            PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm050312"
         End If
         '2017/12/29 END
         
         'Added by Morgan 2021/10/29
         cnnConnection.CommitTrans
         bolInTran = False
         'end 2021/10/29
         
         'Add by Morgan 2008/5/30 抓繳費年度
         m_dblYear = 0
         If "" & .Fields("np07") = "605" Or "" & .Fields("np07") = "606" Or "" & .Fields("np07") = "607" Then
            pa(1) = m_NP02: pa(2) = m_NP03: pa(3) = m_NP04: pa(4) = m_NP05
            m_dblYear = PUB_GetNextYear(pa)
         End If
         'Add by Morgan 2005/4/18
         m_PA12 = "" & .Fields("PA12")
         PrintLetter CheckStr(.Fields(1)), CheckStr(.Fields(0)), .Fields("NP01"), "", .Fields("NP22")
         
         '92.1.10 ADD BY SONIA
         g_PrtForm001.PrintForm .Fields(6), .Fields(2), .Fields(3), .Fields(4), .Fields(5)
         '92.1.10 END
         If Me.Option2(0).Value Then Exit Do 'Added by Morgan 2017/12/19 單筆只抓最近的期限
      End If  'end 'Add by Lydia 2014/11/24 CheckCPExists
      
         .MoveNext
      Loop
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/6
      ShowNoData
      Exit Sub
   End If
End With
ShowPrintOk
'Modified by Lydia 2016/01/04
'For i = 0 To 7
'   txt2(i) = ""
'Next i
For Each oText In txt2
   oText.Text = ""
   oText.Tag = ""
Next
lblMerge(0).Visible = False: lblMerge(1).Visible = False
txt2(8).Visible = False: txt2(9).Visible = False
lbl1 = ""
lbl1.Tag = lbl1 'Added by Morgan 2025/3/13
Screen.MousePointer = vbDefault

'Added by Morgan 2021/10/29
Exit Sub
ErrHnd:
   If bolInTran Then
      cnnConnection.RollbackTrans
   End If
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description, vbCritical
   ElseIf strErrMsg <> "" Then
      MsgBox strErrMsg, vbCritical
   End If
Screen.MousePointer = vbDefault
'end 2021/10/29
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Modify by Morgan 2008/4/30 加 strNP22
Private Sub PrintLetter(ByVal strNP07 As String, ByVal strPA09 As String, ByVal strCP09 As String, ByVal strTM As String, Optional strNP22 As String)
   Dim bolNoConfirmPrice As Boolean
   Dim strET02 As String   'Added by Morgan 2025/5/21
   
   strET02 = strCP09 'Added by Morgan 2025/5/21
   
   bolNoConfirmPrice = False
   
   '92.1.10 MODIFY BY SONIA
   strTmp = ""   '93.11.15 ADD BY SONIA
   Select Case strNP07
      Case "416"
         Select Case strPA09
            '泰國
            Case "019"
               strTmp = "01"
            '巴西
            Case "117"
               strTmp = "02"
            '日本
            Case "011"
               strTmp = "03"
            '加拿大
            Case "102"
               strTmp = "04"
            '新加坡
            Case "014"
               strTmp = "05"
            '韓國
            Case "012"
               strTmp = "06"
            '英國
            Case "201"
               strTmp = "07"
            '歐盟 epc
            Case "221"
               strTmp = "07"
               'Added by Lydia 2016/01/04
               If bolMerge = True Then
                  strTmp = "34"
               End If
            '澳洲
            Case "015"
               strTmp = "09"
            '馬來西亞
            Case "018"
               If strTM = "Y" Then
                  strTmp = "11"
               Else
                  strTmp = "12"
               End If
            '越南
            Case "042"
               strTmp = "14"
            '印度 92.6.23 add by sonia
            Case "040"
               strTmp = "15"
            '俄羅斯 92.9.20 ADD BY SONIA
            'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
            Case "023"
               strTmp = "16"
            'PCT Add by Morgan 2009/7/29
            Case "056"
               strTmp = "25"
            Case Else
               '93.11.15 add by sonia
               strTmp = "15"
         End Select
      'Add by Morgan 2008/9/2
      Case "427" '合併檢索及實審
          Select Case strPA09
            '新加坡
            Case "014"
               strTmp = "24"
         End Select
      Case "605", "606", "607"
         Select Case strPA09
            Case "221"
               If m_PA21 <> "" Then
                  strTmp = "08"
               Else
                  '92.7.14 MODIFY BY SONIA
                  'strTmp = "13"
                  If m_PA46 = "Y" Then     'PCT
                     strTmp = "21"
                  Else
                     If strNP07 = "605" Then
                        strTmp = "18"
'Remove by Morgan 2011/9/16 EPC 只有年費
'                     ElseIf strNP07 = "606" Then
'                        strTmp = "19"
'                     Else
'                        strTmp = "13"
                     End If
                  End If
                  '92.7.14 END
               End If
            '92.7.14 ADD BY SONIA
            Case "011", "012"
               '2008/11/27 MODIFY BY SONIA 日韓為發證日起算年費,但因韓國新型有修法問題,故取消PCT定稿CFP-019667
               'If m_PA46 = "Y" Then     'PCT
               '   strTmp = "21"
               'Else
                     If strNP07 = "605" Then
                        strTmp = "18"
                     ElseIf strNP07 = "606" Then
                        strTmp = "19"
                     Else
                        strTmp = "13"
                     End If
               'End If
            '92.7.14 END
            '2007/10/2 ADD BY SONIA 德國延展費用不同定稿 22
            Case "231"
               If m_PA46 = "Y" Then     'PCT
                  strTmp = "21"
               Else
                  If strNP07 = "605" Then
                     strTmp = "18"
                  ElseIf strNP07 = "606" Then
                     strTmp = "19"
                  Else
                     strTmp = "22"
                  End If
               End If
               
            'Add by Morgan 2008/1/7
            Case "040"
               If strNP07 = "605" Then
                  'Modified by Lydia 2017/06/01 印度年費定稿改成一般;另外新增印度催商業使用聲明定稿
                  'strTmp = "23"
                  strTmp = "18"
                  'end 2017/06/01
               ElseIf strNP07 = "606" Then
                  strTmp = "19"
               Else
                  strTmp = "13"
               End If
            
            'add by sonia 2018/3/16
            'Removed by Morgan 2020/9/18 改刊登廣告
            'Case "048"  '緬甸延展費
            '   strTmp = "37"
            'end 2020/9/18
            'end 2018/3/16
               
            '2007/10/2 END
            Case Else
                  'Add by Morgan 2004/8/13   PCT
                  If m_PA46 = "Y" And m_PA21 = "" Then
                     strTmp = "21"
                  Else
                  'end 2004/8/13
                  
                     '92.10.22 modify by sonia
                     'strTmp = "13"
                     'If m_PA21 <> "" Then 'Removed by Morgan 2006/12/27 還是要判斷案件性質
                        If strNP07 = "605" Then
                           strTmp = "18"
                        ElseIf strNP07 = "606" Then
                           strTmp = "19"
                        Else
                           'Added by Morgan 2012/4/12
                           '馬來西亞 新型
                           If strPA09 = "018" And m_PA08 = "2" Then
                              strTmp = "31"
                           Else
                           'end 2012/4/12
                              strTmp = "13"
                           End If
                        End If
                        'Added by Lydia 2016/01/04
                        If bolMerge = True And InStr("605,607", strNP07) > 0 Then
                           If strPA09 = "018" And m_PA08 = "2" Then
                               strTmp = "32"
                           'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
                           ElseIf strPA09 = "023" And m_PA08 = "3" Then
                               strTmp = "33"
                           End If
                        End If
                        'end 2016/01/04
                        
                     'Removed by Morgan 2006/12/27 還是要判斷案件性質
                     'Else
                     '   strTmp = "18"
                     'end 2006/12/27
                     'End If
                     'end 2006/12/27
                     '92.10.22 end
                  End If
         End Select
         'Add by Morgan 2008/9/2 PCT 非年費定稿不同
         If strTmp = "21" Then
            If strNP07 <> "605" Then
               strTmp = "20"
            End If
         End If
         
      Case "207"
         Select Case strPA09
            Case "015" '澳洲
               strTmp = "10"

            'Add by Morgan 2009/7/29
            Case "018" '馬來西亞
               strTmp = "27"
            'Add by Amy 2018/03/21
            Case "221" 'EPC
                strTmp = "38"
         End Select
         
      'Add by Morgan 2005/4/18 指定費定稿
      Case "215"
         Select Case strPA09
            Case "221" 'EPC
               'Modified by Morgan 2011/12/2 不會再有舊法的案件定稿刪除--慧汶
               'If Val(m_PA10) < 20090401 Then
               '   strTmp = "17"
               'Else
                  strTmp = "28"
               'End If
               'end 2011/12/2
               If Val(txt2(5)) = 0 Then txt2(5) = "12000"
               'Added by Lydia 2016/01/04
               If bolMerge = True Then strTmp = "34"
         End Select
      'Added by Morgan 2020/9/18
      Case "951" '刊登廣告
         If strPA09 = "048" Then '緬甸
            strTmp = "37"
         End If
      'end 2020/9/18
      
      'Added by Lydia 2017/06/01 催商業使用聲明定稿
      Case "930"
         If strPA09 = "040" Then   '印度
            'modify by sonia 2024/5/28 印度2024/3修法，商業使用聲明改3年一次，故再增加專用期已過期之定稿40
            'strTmp = "35"
            If Val(m_PA25) >= Val(m_NP09) Then
               strTmp = "35"
            Else
               strTmp = "40"
            End If
         End If
         'add by sonia 2017/12/27
         If strPA09 = "235" Then   '土耳其
            'Added by Morgan 2025/5/21 +EPC子案定稿
            If cp(4) <> "00" Then
               strTmp = "17"
               strET02 = m_LD18
            Else
             'end 2025/5/21
             
               strTmp = "36"
               
            End If
         End If
         'end 2017/12/27
      'end 2017/06/01
      
      'Added by Morgan 2018/7/13
      Case "421" '檢索報告
         '土耳其新型
         If strPA09 = "235" And m_PA08 = "2" Then
            strTmp = "39"
         End If
      'end 2018/7/13
      Case Else '非報價定稿
         bolNoConfirmPrice = True 'Added by Morgan 2015/2/9
         
         Select Case strNP07
            Case "119" 'Add by Morgan 2009/7/29 PCT進國家階段
               strTmp = "26"
            Case "910" 'Add by Morgan 2010/6/11 其他(美國暫時申請)
               strTmp = "29"
         End Select
         
         'Modify by Morgan 2010/5/31 不必經過智權人員確認
         If strTmp <> "" Then
            InsExpField strNP07, strPA09, strET02, strTmp
            NowPrint strET02, ET01, strTmp, False, strUserNum, , , , , , , , , , , , , m_LD18
            strTmp = ""
            UpdateDate m_LD18 'Added by Morgan 2017/12/19
         End If
   End Select
   '92.1.10 END
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   'Modified by Morgan 2015/2/9
   'If strTmp <> "" Then
   If bolNoConfirmPrice = False Then
   'end 2015/2/9
   
      'Add by Morgan 2008/5/1 新增報價通知
      If Val(strSrvDate(1)) > 20080810 Then
         'Modified by Morgan 2015/2/9
         '要報價但沒有定稿時提醒
         If strTmp = "" Then
            MsgBox "本案要報價但沒有系統的定稿，請注意！", vbExclamation
         Else
         'end 2015/2/9
         
            PUB_AddLetterCache strCP09, strNP22, strET02, ET01, strTmp, m_dblYear, m_LD18
            InsExpField1 strNP07, strPA09, strTmp, strCP09, strNP22
            
            'Add by Morgan 2008/10/21
            '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
            strExc(0) = CompWorkDay(5, strSrvDate(1))
            strExc(1) = DBDATE(m_NP08)
            'Modify by Morgan 2009/1/6 開放維護功能(因常有需要重新產生定稿)--慧汶
            'Memo 2018/03/20 維護功能已取消--Morgan
            If Val(strExc(1)) <= Val(strExc(0)) Then
               PUB_Cache2Letter strCP09, strNP22, False
            End If
            'end 2008/10/21
            
         End If 'Modified by Morgan 2015/2/9
      Else
         
         InsExpField strNP07, strPA09, strET02, strTmp
         NowPrint strET02, ET01, strTmp, False, strUserNum, , , , , , , , , , , , , m_LD18
         
      End If
      'end 2008/5/1
   End If
End Sub

Public Sub InsExpField(ByVal strNP07 As String, ByVal strPA09 As String, ByVal strCP09 As String, ByVal ET03 As String)
   Dim strTxt(1 To 99) As String, iStep As Integer
   
   iStep = 1

   EndLetter ET01, strCP09, ET03, strUserNum
   
   strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
       "','下一程序業務員','" & m_st02 & "')"
   iStep = iStep + 1
   
   If txt2(4) <> "" Then
      strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
         "','下一程序','" & txt2(4) & "')"
      iStep = iStep + 1
   End If
   
   strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
       "','下一程序名稱','" & lbl1.Caption & "')"
   iStep = iStep + 1

   strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
       "','法定期限','" & m_NP09 & "')"
   iStep = iStep + 1

   'Added by Morgan 2016/8/9
   '若法定期限與專用期止日相差不足半年時定稿帶出將屆滿的句子
   If Val(m_PA25) > 0 And Val(m_NP09) > 0 Then
      If m_PA25 < CompDate(1, 6, m_NP09) Then
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
             "','即將屆滿','♀')"
         iStep = iStep + 1
      End If
   End If
   'end 2016/8/9

   strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
        "','本所期限','" & m_NP08 & "')"
   iStep = iStep + 1
   
   'Add by Morgan 2010/6/11
   'Modify by Morgan 2011/10/25 改所限前4周--慧汶
   'strExc(1) = CompDate(1, -1, m_NP09)
   'strExc(1) = CompDate(2, -5, strExc(1))
   strExc(1) = CompDate(2, -28, m_NP08)
   strExc(1) = PUB_GetWorkDay1(strExc(1), True)
   strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
       "','約定期限','" & strExc(1) & "')"
   iStep = iStep + 1
   'end 2010/6/11
   
   If Me.Option2(0).Value Then
        If txt2(5) <> "" Then
           strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
               "','費用','" & Val(txt2(5)) & "')"
           iStep = iStep + 1
           'EPC 要用
           strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
               "','費用合計','" & Val(txt2(5)) & "')"
           iStep = iStep + 1
        End If
        If txt2(7) <> "" Then
           'modify by Morgan 2008/5/12 變數名稱"點數"改為"費用點數"
           strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
               "','費用點數','" & Val(txt2(7)) & "')"
           iStep = iStep + 1
           'EPC 要用
           strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
               "','點數合計','" & Val(txt2(7)) & "')"
           iStep = iStep + 1
        End If
   Else
      '92.5.4 add by sonia
      '美國發明案小個體的第一次維持費為32000,第二次為55000,第三次為78000
      '          大個體的第一次維持費為51000,第二次為96000,第三次為140000
      '德國新型          第一次延展費23000,第二次為30000,第三次為38000
      '93.1.12 MODIFY BY SONIA
      '德國新型          第一次延展費24000,第二次為31000,第三次為40000
      '德國設計          第一次延展費20000,第二次為22000,第三次為25000
      
      'Modify by Morgan 2004/7/9
      '美國發明案小個體的第一次維持費為26000(5.),第二次為48000(6.),第三次為69000(8.)
      '          大個體的第一次維持費為44000,第二次為86000,第三次為127000
      
      '93.9.14 MODIFY by sonia
      '美國發明案小個體的第一次維持費為27000(5.),第二次為49000(6.),第三次為72000(8.)
      '          大個體的第一次維持費為46000(7.),第二次為89000(8.),第三次為132000(10.)
      
      'Modify by Morgan 2004/12/14
      '美國發明案小個體的第一次維持費為25000(5.),第二次為49000(6.),第三次為76000(8.)
      '          大個體的第一次維持費為42000(7.),第二次為89000(8.),第三次為141000(10.)
      
      '2009/3/4 MODIFY BY SONIA 加EU239設計
      'If (strPA09 = "101" And strNP07 = "606") Or (strPA09 = "231" And strNP07 = "607") Then
      'Add By Sindy 2012/8/22 加註 frm210138 也有此費用的計算,若有異動時,須一併改寫
      If (strPA09 = "101" And strNP07 = "606") Or (strPA09 = "231" And strNP07 = "607") Or (strPA09 = "239" And strNP07 = "607") Then
         'Modify By Sindy 2009/07/29
         'Getnexttimes   '取得下次繳費次數
         m_Nexttimes = PUB_Getnexttimes(pa(1), pa(2), pa(3), pa(4), m_strYear, m_PA91)
         '2009/07/29 End
         
         'Add by Morgan 2008/4/30 大個體改成也抓 PATENTYEARFEE
         'Modified by Morgan 2023/3/29
         'If InStr(1, m_PA91, "大個體", 1) > 0 Then
         m_PA179 = PUB_GetEntityType(pa(1), pa(2), pa(3), pa(4))
         If m_PA179 = "1" Then
         'end 2023/3/29
            'Modify By Sindy 2009/07/29 把m_Nexttimes改為m_strYear
            strExc(0) = "SELECT nvl(YF06,0),nvl(YF07,0) FROM PATENTYEARFEE WHERE YF01=" & CNULL(strPA09) & " AND YF02=" & CNULL(m_PA08) & " AND " & _
               "YF03='Y00000002' AND YF04=" & CNULL(strNP07) & " AND YF05=" & CNULL(m_strYear)
         'Added by Morgan 2013/3/20
         'Modified by Morgan 2023/3/29
         'ElseIf InStr(1, m_PA91, "微個體", 1) > 0 Then
         ElseIf m_PA179 = "3" Then
         'end 2023/3/29
            strExc(0) = "SELECT nvl(YF06,0),nvl(YF07,0) FROM PATENTYEARFEE WHERE YF01=" & CNULL(strPA09) & " AND YF02=" & CNULL(m_PA08) & " AND " & _
               "YF03='Y00000003' AND YF04=" & CNULL(strNP07) & " AND YF05=" & CNULL(m_strYear)
         'end 2013/3/20
         Else
            'Modify By Sindy 2009/07/29 把m_Nexttimes改為m_strYear
            strExc(0) = "SELECT nvl(YF06,0),nvl(YF07,0) FROM PATENTYEARFEE WHERE YF01=" & CNULL(strPA09) & " AND YF02=" & CNULL(m_PA08) & " AND " & _
               "YF03='Y00000000' AND YF04=" & CNULL(strNP07) & " AND YF05=" & CNULL(m_strYear)
         End If
         
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Modify by Morgan 2008/4/30
'            Select Case strPA09
'               Case "101"
'                  strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                      "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
'                      "','費用','" & Val(RsTemp.Fields(0)) + Val(RsTemp.Fields(1)) & "')"
'                  If InStr(1, m_PA91, "大個體", 1) > 0 Then
'                     Select Case m_Nexttimes
'                        Case "1"
'                           strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                               "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
'                               "','費用','42000')"
'                        Case "2"
'                           strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                               "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
'                               "','費用','89000')"
'                        Case "3"
'                           strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                               "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
'                               "','費用','141000')"
'                     End Select
'                  End If
'                  iStep = iStep + 1
'               Case "231"
'                  strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                      "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
'                      "','費用','" & Val(RsTemp.Fields(0)) + Val(RsTemp.Fields(1)) & "')"
'                  iStep = iStep + 1
'            End Select
            strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
                "','費用','" & Val(RsTemp.Fields(0)) + Val(RsTemp.Fields(1)) & "')"
            iStep = iStep + 1
            'modify by Morgan 2008/5/12 變數名稱"點數"改為"費用點數"
            strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
                "','費用點數','" & Val(RsTemp.Fields(0)) / 1000 & "')"
            iStep = iStep + 1
         End If
      End If
      '92.5.4 end
   End If
   
   '92.1.10 end
   
   'Add by Morgan 2005/4/18
   If strNP07 = "215" Then
      If m_PA12 = "" Then
         strExc(0) = "申請日起2年內"
      Else
         strExc(0) = "公開日起半年內"
      End If
      strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
          "','指定期限','" & strExc(0) & "')"
      iStep = iStep + 1
      If GetMemberCountryData(m_PA10, strExc) = True Then
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
             "','生效日','" & strExc(1) & "')"
         iStep = iStep + 1
         
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
             "','成員國數','" & strExc(2) & "')"
         iStep = iStep + 1
         
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
             "','成員國','" & strExc(3) & "')"
         iStep = iStep + 1
         
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
             "','延伸國數','" & strExc(4) & "')"
         iStep = iStep + 1
         
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
             "','延伸國','" & strExc(5) & "')"
         iStep = iStep + 1
      End If
   End If
   
   '92.7.14 ADD BY SONIA
   If ET03 = "03" Then
      If Val(m_PA10) >= 20011001 Then
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
             "','列印備註','3')"
      Else
         strTxt(iStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & _
             "','列印備註','7')"
      End If
      iStep = iStep + 1
   End If

   '92.7.14 END
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(iStep - 1, strTxt) Then
   If Not ClsLawExecSQL(iStep - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If

End Sub

'Add by Morgan 2005/4/18 讀取epc會員國相關資料
'p_Data():1=生效日,2=成員國數,3=成員國,4=延伸國數,5=延伸國
Private Function GetMemberCountryData(ByVal p_AppDate As String, ByRef p_Data() As String) As Boolean

   Dim iCnt1 As Integer, sCountry1 As String, iCnt2 As Integer, sCountry2 As String, iPos As Integer
On Error GoTo ErrHnd
   '若無申請日期則預設系統日
   If p_AppDate = "" Then p_AppDate = strSrvDate(1)
   strSql = "select max(b.mc02) from membercountry b where b.mc01='1' and b.mc02<=" & p_AppDate
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         If Not IsNull(.Fields(0)) Then
            p_Data(1) = .Fields(0)
            strSql = "select nvl(na03,mc03),mc04 from membercountry,nation a" & _
               " where mc01='1' and mc02=" & p_Data(1) & " and na01(+)=mc03 order by mc04,na03"
            CheckOC3
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
               Do While Not .EOF
                  If "" & .Fields(1) = "1" Then
                     iCnt1 = iCnt1 + 1
                     sCountry1 = sCountry1 & IIf(sCountry1 = "", "", "、") & .Fields(0)
                  Else
                     iCnt2 = iCnt2 + 1
                     sCountry2 = sCountry2 & IIf(sCountry2 = "", "", "、") & .Fields(0)
                  End If
                  .MoveNext
               Loop
               iPos = 0
               Do While InStr(iPos + 1, sCountry1, "、") > 0
                  iPos = InStr(iPos + 1, sCountry1, "、")
               Loop
               If iPos > 1 Then
                  sCountry1 = Left(sCountry1, iPos - 1) & "及" & Mid(sCountry1, iPos + 1)
               End If
               iPos = 0
               Do While InStr(iPos + 1, sCountry2, "、") > 0
                  iPos = InStr(iPos + 1, sCountry2, "、")
               Loop
               If iPos > 1 Then
                  sCountry2 = Left(sCountry2, iPos - 1) & "及" & Mid(sCountry2, iPos + 1)
               End If
               p_Data(2) = iCnt1
               p_Data(3) = sCountry1
               p_Data(4) = iCnt2
               p_Data(5) = sCountry2
               GetMemberCountryData = True
            End If
         End If
      End If
   End With
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub Process()                '上半部
Dim strYF15 As String, strFeeType As String 'Add By Sindy 2009/07/31
Dim strText As String 'Add By Sindy 2016/6/20
Dim rsA1 As New ADODB.Recordset 'Added by Lydia 2016/12/08
Dim ArrYear 'Added by Lydia 2017/05/22
Dim strNP01 As String, strNP22 As String 'Add by Amy 2018/04/30
Dim strNP09 As String 'Add by Amy 2018/05/03 法限(子案無期限掛母案 for 子案母案排一起)
'Added by Lydia 2019/08/30 提早催期限
Dim intEx As Integer
Dim m_ExceptCust() As String  '指定客戶編號
Dim m_ExceptNp07() As String '下一程序
Dim m_ExceptCond() As String '組合條件
Dim intQ As Integer, intX As Integer
'Added by Lydia 2025/07/25
Dim m_ExceptMon() As String '提前X個月
Dim tmpArr1 As Variant
Dim tmpArrData As Variant '(0)代表X編號|(1)客戶編號(含關係企業)|提前X個月


Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R050312 WHERE ID='" & strUserNum & "' "

strTitle = "": strTitle2 = ""  'Added by Lydia 2016/08/25

'Added by Lydia 2025/07/25
If cntFrm040303New = "Y" Then
   '取得例外通知的客戶相關控制資料
   intI = Pub_Getfrm040303ExceptNew("CFP", "", "", "ALL", , tmpArr1)
   intEx = UBound(tmpArr1) + 1
   ReDim m_ExceptCust(1 To intEx)
   ReDim m_ExceptNp07(1 To intEx)
   ReDim m_ExceptCond(1 To intEx)
   ReDim m_ExceptMon(1 To intEx)
   For intI = 0 To UBound(tmpArr1)
      If Trim(tmpArr1(intI)) <> "" Then
         tmpArrData = Split(tmpArr1(intI), "|")
         '(0)代表X編號|(1)客戶編號(含關係企業)|(2)提前X個月
         m_ExceptCust(intI + 1) = Trim("" & tmpArrData(1))
         m_ExceptNp07(intI + 1) = "605,606,607,119,416"
         m_ExceptMon(intI + 1) = Val("" & tmpArrData(2))
      End If
   Next intI
Else
'end 2025/07/25
   'Added by Lydia 2019/08/30 客戶X70017000和碩聯合科技股份有限公司及其關係企業(編號X70017010、X70017020、X70017030)
   '                                      所有P及CFP案年費的通知時間均提早為法定期限前6個月，實審的通知時間則是提早為申請日＋１年。
   '                                      若往後此客戶X70017新建關係企業，由智權人員通知設定。
   'Modified by Lydia 2022/08/12 增加信邦案
   'intEx = 1
   'Modified by Lydia 2022/08/25 增加大亞(X60601000)
   'intEx = 2
   'Modified by Lydia 2022/08/25 增加:康舒科技(X00497070) intE=3 => 4
   'Modified by Lydia 2024/04/22 增加:立德電子(X01506000)、江蘇領先(X01506010) intE=4 => 5 'Memo by Lydia 2024/05/02 增加新編號:東莞立德(X01506020)
   intEx = 5
   ReDim m_ExceptCust(1 To intEx)
   ReDim m_ExceptNp07(1 To intEx)
   ReDim m_ExceptCond(1 To intEx)
   'Modified by Lydia 2022/02/23 改成共用模組取得
   'm_ExceptCust(1) = "X70017000,X70017010,X70017020,X70017030"
   ''Added by Lydia 2022/02/08 長庚體系逐漸要將其專利案件回歸到產學中心，顯然之後皆須遵循其規則進行通知，故除顧服組客戶外，本所其他非顧服組也建議需一併設
   'm_ExceptCust(1) = m_ExceptCust(1) & ",X69365020,X75299000,X75299020,X69365060,X69365010,X69365030,X69365000,X69365040,X69365050"
   intI = Pub_Getfrm040303Except("X70017000", m_ExceptCust(1))
   'end 2022/02/23
   
   m_ExceptNp07(1) = "605,606,607,119,416"
   
   'Added by Lydia 2022/08/12 因信邦電子(X39056000)承辦人反應本所專利年費、維持費繳交期限提前通知日期過早(現為三個月)，要求本所在法定期限前一個月通知即可
   intI = Pub_Getfrm040303Except("X39056000", m_ExceptCust(2))
   m_ExceptNp07(2) = "605,606,607,119,416"
   'end 2022/08/12
      
   'Added by Lydia 2022/08/25 增加大亞電線電纜股份有限公司(X60601000)因為有導入TIPS智財管理制度，目前被要求專利案件年費需於法定期限前4個月通知，故客戶來電希望本所協助調整期限通知。
   intI = Pub_Getfrm040303Except("X60601000", m_ExceptCust(3))
   m_ExceptNp07(3) = "605,606,607,119,416"
   'end 2022/08/25
   
   'Added by Lydia 2023/11/09 康舒科技(X00497070)年費期限通知由三個月前改為一個月前; 'Memo by Lydia 2024/04/23 CFP案為法限前兩個月通知
   'Mark by Lydia 2024/12/05 (12/4)現因該公司內部問題，希望本所回復為原先三個月期限通知。
   'intI = Pub_Getfrm040303Except("X00497070", m_ExceptCust(4))
   'm_ExceptNp07(4) = "605,606,607,119,416"
   ''end 2023/11/09
   'end 2024/12/05
   
   'Added b Lydia 2024/04/22 立德電子(X01506000)、江蘇領先(X01506010) 年費期限通知由三個月前改為一個月前; 'Memo by Lydia 2024/04/23 CFP案為法限前兩個月通知
   'Memo by Lydia 2024/05/02 增加新編號:東莞立德(X01506020)
   intI = Pub_Getfrm040303Except("X01506000", m_ExceptCust(5))
   m_ExceptNp07(5) = "605,606,607,119,416"
   'end 2024/04/22
End If 'Added by Lydia 2025/07/25

'系統類別
strSQL1 = ""
strSQL2 = ""
If Len(Trim(txt1(1))) <> 0 Then
    'Modify By Cheng 2002/12/13
'   strSQL1 = strSQL1 + " AND NP08>=" & Val(ChangeTStringToWString(txt1(1))) & " "
   strSQL1 = strSQL1 + " AND NP09>=" & Val(ChangeTStringToWString(txt1(1))) & " "
End If
If Len(Trim(txt1(2))) <> 0 Then
    'Modify By Cheng 2002/12/13
'   strSQL1 = strSQL1 & " AND NP08<=" & Val(ChangeTStringToWString(txt1(2))) & " "
   strSQL1 = strSQL1 & " AND NP09<=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
If Len(Trim(txt1(1))) <> 0 Or Len(Trim(txt1(2))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/6
End If
strSQL2 = strSQL1
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 & " AND NP02 IN (" & SQLGrpStr(txt1(0), 1) & ")"
   strSQL2 = strSQL2 & " AND NP02 IN (" & SQLGrpStr(txt1(0), 5) & ")"
   pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0)  'Add By Sindy 2010/12/6
End If
If Len(txt1(3)) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)>='" & txt1(3) & "' "
    strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)>='" & txt1(3) & "' "
End If
If Len(txt1(4)) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)<='" & txt1(4) & "' "
    strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)<='" & txt1(4) & "' "
End If
If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label3 & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/12/6
End If
If Len(txt1(5)) <> 0 Then
    strSQL1 = strSQL1 + " AND NP07>=" & Val(txt1(5))
    strSQL2 = strSQL2 + " AND NP07>=" & Val(txt1(5))
End If
If Len(txt1(6)) <> 0 Then
    strSQL1 = strSQL1 + " AND NP07<=" & Val(txt1(6))
    strSQL2 = strSQL2 + " AND NP07<=" & Val(txt1(6))
End If

'Added by Lydia 2019/08/30 判斷客戶和下一程序的性質；排除和碩的年費和實審
strExc(1) = ""
strExc(2) = ""
If Trim(txt1(5) & txt1(6)) <> "" Or Trim(txt1(7) & txt1(8)) <> "" Then
    For intX = 1 To intEx
        If Trim(txt1(7) & txt1(8)) <> "" Then '判斷客戶區間
             strExc(0) = "select cu01||cu02 as custno from customer where cu01||cu02 in (" & GetAddStr(m_ExceptCust(intX)) & ") "
             If Trim(txt1(7)) <> "" Then strExc(0) = strExc(0) & " and cu01||cu02>=" & CNULL(GetNewFagent(txt1(7)))
             If Trim(txt1(8)) <> "" Then strExc(0) = strExc(0) & " and cu01||cu02<=" & CNULL(GetNewFagent(txt1(8)))
             strExc(0) = strExc(0) & " order by 1"
             
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
             If intI = 0 Then
                 m_ExceptCust(intX) = ""
             Else
                 strExc(1) = ""
                 RsTemp.MoveFirst
                 Do While Not RsTemp.EOF
                     If InStr(m_ExceptCust(intX), "" & RsTemp.Fields("custno")) > 0 Then
                        strExc(1) = strExc(1) & "," & RsTemp.Fields("custno")
                     End If
                     RsTemp.MoveNext
                 Loop
                 If strExc(1) <> "" Then m_ExceptCust(intX) = Mid(strExc(1), 2)
             End If
        End If
        If m_ExceptCust(intX) <> "" And Trim(txt1(5) & txt1(6)) <> "" Then '判斷下一程序
             strExc(0) = "select cpm02 from casepropertymap where cpm01='CFP' and cpm02 in (" & GetAddStr(m_ExceptNp07(intX)) & ") "
             If Trim(txt1(5)) <> "" Then strExc(0) = strExc(0) & " and cpm02>=" & txt1(5)
             If Trim(txt1(6)) <> "" Then strExc(0) = strExc(0) & " and cpm02<=" & txt1(6)
             strExc(0) = strExc(0) & " order by 1"
             
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
             If intI = 0 Then
                 m_ExceptCust(intX) = ""
             Else
                 strExc(1) = ""
                 RsTemp.MoveFirst
                 Do While Not RsTemp.EOF
                     If InStr(m_ExceptNp07(intX), "" & RsTemp.Fields("cpm02")) > 0 Then
                        strExc(1) = strExc(1) & "," & RsTemp.Fields("cpm02")
                     End If
                     RsTemp.MoveNext
                 Loop
                 If strExc(1) <> "" Then m_ExceptNp07(intX) = Mid(strExc(1), 2)
             End If
        End If
    Next
End If

'提早催期限的條件
strExc(1) = ""
For intX = 1 To intEx
   If Trim(m_ExceptCust(intX)) <> "" Then
      'Added by Lydia 2025/07/25
      If cntFrm040303New = "Y" Then
         m_ExceptCond(intX) = " AND INSTR('" & m_ExceptCust(intX) & "',PA26)>0 AND ("
         '年費+實審：一般提前3個月催期限=> 輸入日期-(3-設定月數)
         If InStr(m_ExceptCust(intX), "X01506") > 0 Then '立德電子(X01506000)、江蘇領先(X01506010):P案年費期限通知由三個月前改為一個月前;CFP案為法限前兩個月通知，由程式控制。
            intQ = -1
         Else
            intQ = (3 - Val(m_ExceptMon(intX))) * -1
         End If
         strExc(2) = Mid(CompDate(1, intQ, DBDATE(txt1(2))), 1, 6) & "31"
         m_ExceptCond(intX) = m_ExceptCond(intX) & " (NP07 IN (" & Replace(m_ExceptNp07(intX), ",416", "") & ") AND NP09 BETWEEN " & CompDate(1, intQ, DBDATE(txt1(1))) & " AND " & strExc(2) & ")"
         '實審416->系統日減一個月＜申請日＋1年＜系統日
         m_ExceptCond(intX) = m_ExceptCond(intX) & " OR (NP07='416' AND PA10+10000 BETWEEN " & Mid(CompDate(1, -1, strSrvDate(1)), 1, 6) & "01" & " AND " & strSrvDate(1) & ")"
         m_ExceptCond(intX) = m_ExceptCond(intX) & " )"
      
         strExc(1) = strExc(1) & " OR (INSTR('" & m_ExceptCust(intX) & "',PA26)>0 AND NP07 IN (" & m_ExceptNp07(intX) & "))"
      Else
      'end 2025/07/25
         'Added by Lydia 2022/08/12
         If intX = 2 Then  '信邦案之指定客戶(晚催期限，法定期限前一個月通知)
            m_ExceptCond(intX) = " AND INSTR('" & m_ExceptCust(intX) & "',PA26)>0 AND ("
            '年費+實審->法定期限前一個月通知；一般提前3個月催期限=> 輸入日期-2個月
            'Modified by Morgan 2023/7/7 改法限前兩個月通知--文雄
            'If Val(Right(txt1(2), 2)) >= 28 Then
            '    strExc(2) = Mid(CompDate(1, -2, DBDATE(txt1(2))), 1, 6) & "31"
            'Else
            '    strExc(2) = CompDate(1, -2, DBDATE(txt1(2)))
            'End If
            strExc(2) = Mid(CompDate(1, -1, DBDATE(txt1(2))), 1, 6) & "31"
            'end 2023/7/7
            'Modified by Lydia 2023/11/09
            'm_ExceptCond(intX) = m_ExceptCond(intX) & " (NP07 IN (" & m_ExceptNp07(intX) & ") AND NP09 BETWEEN " & CompDate(1, -2, DBDATE(txt1(1))) & " AND " & strExc(2) & ")"
            m_ExceptCond(intX) = m_ExceptCond(intX) & " (NP07 IN (" & m_ExceptNp07(intX) & ") AND NP09 BETWEEN " & CompDate(1, -1, DBDATE(txt1(1))) & " AND " & strExc(2) & ")"
            m_ExceptCond(intX) = m_ExceptCond(intX) & " )"
         
            strExc(1) = strExc(1) & " OR (INSTR('" & m_ExceptCust(intX) & "',PA26)>0 AND NP07 IN (" & m_ExceptNp07(intX) & "))"
         'Added by Lydia 2022/08/25
         ElseIf intX = 3 Then  '大亞案(X60601000)：法定期限前4個月
            m_ExceptCond(intX) = " AND INSTR('" & m_ExceptCust(intX) & "',PA26)>0 AND ("
            '年費+實審->法定期限前4個月通知；一般提前3個月催期限=> 輸入日期+1個月
            'Modified by Lydia 2023/11/09
            'If Val(Right(txt1(2), 2)) >= 28 Then
            '    strExc(2) = Mid(CompDate(1, 1, DBDATE(txt1(2))), 1, 6) & "31"
            'Else
            '    strExc(2) = CompDate(1, 1, DBDATE(txt1(2)))
            'End If
            strExc(2) = Mid(CompDate(1, 1, DBDATE(txt1(2))), 1, 6) & "31"
            'end 2023/11/09
            m_ExceptCond(intX) = m_ExceptCond(intX) & " (NP07 IN (" & m_ExceptNp07(intX) & ") AND NP09 BETWEEN " & CompDate(1, 1, DBDATE(txt1(1))) & " AND " & strExc(2) & ")"
            m_ExceptCond(intX) = m_ExceptCond(intX) & " )"
         
            strExc(1) = strExc(1) & " OR (INSTR('" & m_ExceptCust(intX) & "',PA26)>0 AND NP07 IN (" & m_ExceptNp07(intX) & "))"
         'Added by Lydia 2023/11/09
         ElseIf intX = 4 Then  '康舒科技(X00497070)年費期限通知由三個月前改為一個月前
            m_ExceptCond(intX) = " AND INSTR('" & m_ExceptCust(intX) & "',PA26)>0 AND ("
            '年費+實審->法定期限前4個月通知；一般提前3個月催期限=> 輸入日期-2個月
            'Modified by Lydia 2024/04/23 CFP案為法限前兩個月通知
            'strExc(2) = Mid(CompDate(1, -2, DBDATE(txt1(2))), 1, 6) & "31"
            strExc(2) = Mid(CompDate(1, -1, DBDATE(txt1(2))), 1, 6) & "31"
            m_ExceptCond(intX) = m_ExceptCond(intX) & " (NP07 IN (" & m_ExceptNp07(intX) & ") AND NP09 BETWEEN " & CompDate(1, -2, DBDATE(txt1(1))) & " AND " & strExc(2) & ")"
            m_ExceptCond(intX) = m_ExceptCond(intX) & " )"
         
            strExc(1) = strExc(1) & " OR (INSTR('" & m_ExceptCust(intX) & "',PA26)>0 AND NP07 IN (" & m_ExceptNp07(intX) & "))"
         'end 2023/11/09
         'Added by Lydia 2024/04/22
         ElseIf intX = 5 Then  '立德電子(X01506000)、江蘇領先(X01506010) 年費期限通知由三個月前改為一個月前 'Memo by Lydia 2024/05/02 增加新編號:東莞立德(X01506020)
            m_ExceptCond(intX) = " AND INSTR('" & m_ExceptCust(intX) & "',PA26)>0 AND ("
            '年費+實審->法定期限前4個月通知；一般提前3個月催期限=> 輸入日期-2個月
            'Modified by Lydia 2024/04/23 CFP案為法限前兩個月通知
            'strExc(2) = Mid(CompDate(1, -2, DBDATE(txt1(2))), 1, 6) & "31"
            strExc(2) = Mid(CompDate(1, -1, DBDATE(txt1(2))), 1, 6) & "31"
            'Modified by Lydia 2025/07/25 CompDate(1, -2, DBDATE(txt1(1)))=>CompDate(1, -1, DBDATE(txt1(1)))
            m_ExceptCond(intX) = m_ExceptCond(intX) & " (NP07 IN (" & m_ExceptNp07(intX) & ") AND NP09 BETWEEN " & CompDate(1, -1, DBDATE(txt1(1))) & " AND " & strExc(2) & ")"
            m_ExceptCond(intX) = m_ExceptCond(intX) & " )"
         
            strExc(1) = strExc(1) & " OR (INSTR('" & m_ExceptCust(intX) & "',PA26)>0 AND NP07 IN (" & m_ExceptNp07(intX) & "))"
         'end 2024/04/22
         Else
         'end 2022/08/12
            m_ExceptCond(intX) = " AND INSTR('" & m_ExceptCust(intX) & "',PA26)>0 AND ("
            '年費->法定期限前6個月；一般提前3個月催期限=> 輸入日期+3個月
            'Modified by Lydia 2023/11/09
            'If Val(Right(txt1(2), 2)) >= 28 Then
            '    strExc(2) = Mid(CompDate(1, 3, DBDATE(txt1(2))), 1, 6) & "31"
            'Else
            '    strExc(2) = CompDate(1, 3, DBDATE(txt1(2)))
            'End If
            strExc(2) = Mid(CompDate(1, 3, DBDATE(txt1(2))), 1, 6) & "31"
            'end 2023/11/09
            m_ExceptCond(intX) = m_ExceptCond(intX) & " (NP07 IN (" & Replace(m_ExceptNp07(intX), ",416", "") & ") AND NP09 BETWEEN " & CompDate(1, 3, DBDATE(txt1(1))) & " AND " & strExc(2) & ")"
            '實審416->系統日減一個月＜申請日＋1年＜系統日
            m_ExceptCond(intX) = m_ExceptCond(intX) & " OR (NP07='416' AND PA10+10000 BETWEEN " & Mid(CompDate(1, -1, strSrvDate(1)), 1, 6) & "01" & " AND " & strSrvDate(1) & ")"
            m_ExceptCond(intX) = m_ExceptCond(intX) & " )"
         
            strExc(1) = strExc(1) & " OR (INSTR('" & m_ExceptCust(intX) & "',PA26)>0 AND NP07 IN (" & m_ExceptNp07(intX) & "))"
         End If 'Added by Lydia 2022/08/12
      End If 'Added by Lydia 2025/07/25
   End If
Next intX
If strExc(1) <> "" Then
    strSQL1 = strSQL1 + " AND NOT (" & Mid(strExc(1), 4) & ") "
End If
'end 2019/08/30

If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label4 & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/12/6
End If
If Len(Trim(txt1(7))) <> 0 And Len(Trim(txt1(8))) <> 0 Then
    strSQL1 = strSQL1 + " AND PA26>='" & GetNewFagent(txt1(7)) & "' AND PA26<='" & GetNewFagent(txt1(8)) & "' "
    strSQL2 = strSQL2 + " AND SP08>='" & GetNewFagent(txt1(7)) & "' AND SP08<='" & GetNewFagent(txt1(8)) & "' "
Else
    If Len(Trim(txt1(7))) <> 0 And Len(Trim(txt1(8))) = 0 Then
        strSQL1 = strSQL1 + " AND PA26>='" & GetNewFagent(txt1(7)) & "' "
        strSQL2 = strSQL2 + " AND SP08>='" & GetNewFagent(txt1(7)) & "' "
    Else
        If Len(Trim(txt1(7))) = 0 And Len(Trim(txt1(8))) <> 0 Then
            strSQL1 = strSQL1 + " AND PA26<='" & GetNewFagent(txt1(8)) & "' "
            strSQL2 = strSQL2 + " AND SP08<='" & GetNewFagent(txt1(8)) & "' "
        End If
    End If
End If
If Len(Trim(txt1(7))) <> 0 Or Len(Trim(txt1(8))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label5 & txt1(7) & "-" & txt1(8) 'Add By Sindy 2010/12/6
End If
If Len(Trim(txt1(9))) <> 0 And Len(Trim(txt1(10))) <> 0 Then
    strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(9)) & "' AND PA75<='" & GetNewFagent(txt1(10)) & "'"
    strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(9)) & "' AND SP26<='" & GetNewFagent(txt1(10)) & "'"
Else
    If Len(Trim(txt1(9))) <> 0 And Len(Trim(txt1(10))) = 0 Then
        strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(9)) & "' "
        strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(9)) & "' "
    Else
        If Len(Trim(txt1(9))) = 0 And Len(Trim(txt1(10))) <> 0 Then
            strSQL1 = strSQL1 + " AND PA75<='" & GetNewFagent(txt1(10)) & "' "
            strSQL2 = strSQL2 + " AND SP26<='" & GetNewFagent(txt1(10)) & "' "
        End If
    End If
End If
If Len(Trim(txt1(9))) <> 0 Or Len(Trim(txt1(10))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label6 & txt1(9) & "-" & txt1(10) 'Add By Sindy 2010/12/6
End If
'92.03.27 nick add left join
'strSQL = "SELECT NA03,DECODE(PA09,'000',CPM03,CPM04) AS A," & SQLDate("NP08") & "," & SQLDate("NP09") & ",NP02||'-'||NP03||'-'||NP04||'-'||NP05,PA22,NVL(PA05,NVL(PA06,PA07)),ptm03,nvl(CU04,NVL(CU05,CU06))," & SQLDate("PA10") & "," & SQLDate("PA14") & "," & SQLDate("PA21") & ",PA24,PA25 FROM NEXTPROGRESS,PATENT,CASEPROPERTYMAP,NATION,CUSTOMER,PATENTTRADEMARKMAP WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP06 IS NULL AND np02=cpm01(+) AND CPM02=TO_CHAR(NP07) AND pa09=na01(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND (PA57<>'Y' OR PA57 IS NULL) " & strSQL1
'strSQL = strSQL + " union all select NA03,DECODE(SP09,'000',CPM03,CPM04) AS A," & SQLDate("NP08") & "," & SQLDate("NP09") & ",NP02||'-'||NP03||'-'||NP04||'-'||NP05,'',NVL(SP05,NVL(SP06,SP07)),'',NVL(CU04,NVL(CU05,CU06))," & SQLDate("SP10") & ",''," & SQLDate("SP12") & ",SP20,SP21 FROM NEXTPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,NATION,CUSTOMER WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP06 IS NULL AND np02=cpm01(+) AND CPM02=TO_CHAR(NP07) AND sp09=na01(+) AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND (SP15<>'Y' OR SP15 IS NULL) " & strSQL2

'strSql = "SELECT NA03,DECODE(PA09,'000',CPM03,CPM04) AS A," & SQLDate("NP08") & "," & SQLDate("NP09") & ",NP02||'-'||NP03||'-'||NP04||'-'||NP05,PA22,NVL(PA05,NVL(PA06,PA07)),ptm03,nvl(CU04,NVL(CU05,CU06))," & SQLDate("PA10") & "," & SQLDate("PA14") & "," & SQLDate("PA21") & ",PA24,PA25,PA08,PA09 FROM NEXTPROGRESS,PATENT,CASEPROPERTYMAP,NATION,CUSTOMER,PATENTTRADEMARKMAP WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP06 IS NULL AND np02=cpm01(+) AND TO_CHAR(NP07)=cpm02(+) AND pa09=na01(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND (PA57<>'Y' OR PA57 IS NULL) " & strSQL1
'strSql = strSql + " union all select NA03,DECODE(SP09,'000',CPM03,CPM04) AS A," & SQLDate("NP08") & "," & SQLDate("NP09") & ",NP02||'-'||NP03||'-'||NP04||'-'||NP05,'',NVL(SP05,NVL(SP06,SP07)),'',NVL(CU04,NVL(CU05,CU06))," & SQLDate("SP10") & ",''," & SQLDate("SP12") & ",SP20,SP21,'','' FROM NEXTPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,NATION,CUSTOMER WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP06 IS NULL AND np02=cpm01(+) AND TO_CHAR(NP07)=cpm02(+) AND sp09=na01(+) AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND (SP15<>'Y' OR SP15 IS NULL) " & strSQL2
'strSql = strSql + " ORDER BY 1,2,3,4 "
'Modify By Sindy 2016/6/20
'Modified by Lydia 2016/08/25 拿掉PA22,加上PA26,SP08
'strSql = "SELECT NA03,DECODE(PA09,'000',CPM03,CPM04) AS A,st02," & SQLDate("NP09") & ",NP02||'-'||NP03||'-'||NP04||'-'||NP05,PA22,NVL(PA05,NVL(PA06,PA07)),ptm03,nvl(CU04,NVL(CU05,CU06))," & SQLDate("PA10") & "," & SQLDate("PA14") & "," & SQLDate("PA21") & ",PA24,PA25,PA08,PA09,NP01,NP07,NP10 FROM NEXTPROGRESS,PATENT,CASEPROPERTYMAP,NATION,CUSTOMER,PATENTTRADEMARKMAP,staff WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP06 IS NULL AND np10=st01(+) AND np02=cpm01(+) AND TO_CHAR(NP07)=cpm02(+) AND pa09=na01(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND (PA57<>'Y' OR PA57 IS NULL) " & strSQL1
'strSql = strSql + " union all select NA03,DECODE(SP09,'000',CPM03,CPM04) AS A,st02," & SQLDate("NP09") & ",NP02||'-'||NP03||'-'||NP04||'-'||NP05,'',NVL(SP05,NVL(SP06,SP07)),'',NVL(CU04,NVL(CU05,CU06))," & SQLDate("SP10") & ",''," & SQLDate("SP12") & ",SP20,SP21,'','',NP01,NP07,NP10 FROM NEXTPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,NATION,CUSTOMER,staff WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP06 IS NULL AND np10=st01(+) AND np02=cpm01(+) AND TO_CHAR(NP07)=cpm02(+) AND sp09=na01(+) AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND (SP15<>'Y' OR SP15 IS NULL) " & strSQL2
'strSql = strSql + " ORDER BY 1,2,3,4 "
'2016/6/20 END
'Modified by Lydia 2016/12/08 + PA16
'Modified by Lydia 2017/05/22 +PA72/SP09,NA21
'Add by Amy 2018/04/30 +NP22
'Modified by Morgan 2022/6/13 +PA10
strSql = "SELECT NA03,DECODE(PA09,'000',CPM03,CPM04) AS A,st02," & SQLDate("NP09") & " as VS04,NP02||'-'||NP03||'-'||NP04||'-'||NP05 as VS05,NVL(PA05,NVL(PA06,PA07)) as VS06,ptm03 as VS07,nvl(CU04,NVL(CU05,CU06)) as VS08," & SQLDate("PA10") & " as VS09,CU04 as VS10," & SQLDate("PA14") & " as VS11," & SQLDate("PA21") & " as VS12,PA24,PA25,PA26,PA08,PA09,NP01,NP07,NP10,PA16,PA72,NA21,NP22,PA10 FROM NEXTPROGRESS,PATENT,CASEPROPERTYMAP,NATION,CUSTOMER,PATENTTRADEMARKMAP,staff WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP06 IS NULL AND np10=st01(+) AND np02=cpm01(+) AND TO_CHAR(NP07)=cpm02(+) AND pa09=na01(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND (PA57<>'Y' OR PA57 IS NULL) " & strSQL1
strSql = strSql + " union all select NA03,DECODE(SP09,'000',CPM03,CPM04) AS A,st02," & SQLDate("NP09") & " as VS04,NP02||'-'||NP03||'-'||NP04||'-'||NP05 as VS05,NVL(SP05,NVL(SP06,SP07)) as VS06,' ' as VS07,NVL(CU04,NVL(CU05,CU06)) as VS08," & SQLDate("SP10") & " as VS09,' ' as VS10,' ' as VS11," & SQLDate("SP12") & " as VS12,SP20,SP21,SP08,'','',NP01,NP07,NP10,'',SP09,NA21,NP22,SP10 FROM NEXTPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,NATION,CUSTOMER,staff WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP06 IS NULL AND np10=st01(+) AND np02=cpm01(+) AND TO_CHAR(NP07)=cpm02(+) AND sp09=na01(+) AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND (SP15<>'Y' OR SP15 IS NULL) " & strSQL2

'Added by Lydia 2019/08/30 提早催期限的案件
For intX = 1 To intEx
    'Added by Lydia 2019/11/12 系統別
    strExc(1) = ""
    If Trim(txt1(0)) <> "" Then
        strExc(1) = " AND NP02 IN (" & SQLGrpStr(txt1(0), 1) & ")"
    End If
    'end 2019/11/12
    If m_ExceptCust(intX) <> "" Then
        'Modified by Lydia 2019/11/12 debug: + 系統別 strexc(1)
        strSql = strSql & " union all SELECT NA03,DECODE(PA09,'000',CPM03,CPM04) AS A,st02," & SQLDate("NP09") & " as VS04,NP02||'-'||NP03||'-'||NP04||'-'||NP05 as VS05," & _
                    "NVL(PA05,NVL(PA06,PA07)) as VS06,ptm03 as VS07,nvl(CU04,NVL(CU05,CU06)) as VS08," & SQLDate("PA10") & " as VS09,CU04 as VS10," & _
                     SQLDate("PA14") & " as VS11," & SQLDate("PA21") & " as VS12,PA24,PA25,PA26,PA08,PA09,NP01,NP07,NP10,PA16,PA72,NA21,NP22,PA10 " & _
                     "FROM NEXTPROGRESS,PATENT,CASEPROPERTYMAP,NATION,CUSTOMER,PATENTTRADEMARKMAP,staff " & _
                     "WHERE NP02=PA01 AND NP03=PA02 AND NP04=PA03 AND NP05=PA04 AND NP06 IS NULL AND np10=st01(+) " & _
                     "AND np02=cpm01(+) AND TO_CHAR(NP07)=cpm02(+) AND pa09=na01(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) " & _
                     "AND '1'=PTM01(+)  AND PA08=PTM02(+) AND (PA57<>'Y' OR PA57 IS NULL) " & m_ExceptCond(intX) & strExc(1)
    End If
Next intX

strSql = strSql + " ORDER BY 1,2,3,4 "

CheckOC
k = 0
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/7
        .MoveFirst
        DoEvents
        Do While .EOF = False
           'Added by Lydia 2020/03/05 列印個人管制部門案件：只列管制人的案件。
           If ChkPerson.Value = 1 And txt1(16).Text <> "" Then
                strExc(0) = PUB_GetCFPHandler("" & .Fields("VS05"))
                If strExc(0) <> txt1(16).Text Then
                    GoTo JumpToNext
                End If
           End If
           'end 2020/03/05
           For i = 0 To 13
              strTemp(i) = CheckStr(.Fields(i))
           Next i
           'Add by Amy 2018/04/30
           strNP01 = "" & .Fields("NP01")
           strNP22 = "" & .Fields("NP22")
           strNP09 = "" & .Fields("VS04") 'Add by Amy 2018/05/03
           
           'Add by Lydia 2014/11/24 加判斷未收文恢復權利414
           If CheckCPExists(SystemNumber(strTemp(4), 1), SystemNumber(strTemp(4), 2), SystemNumber(strTemp(4), 3), SystemNumber(strTemp(4), 4)) = False Then
               strTemp(12) = strTemp(12) + "-" & strTemp(13)
               'Modify By Sindy 2009/07/31
               'strSQL = "SELECT PA72 FROM PATENT WHERE PA01='" & SystemNumber(strTemp(4), 1) & "' AND PA02='" & SystemNumber(strTemp(4), 2) & "' AND PA03='" & SystemNumber(strTemp(4), 3) & "' AND PA04='" & SystemNumber(strTemp(4), 4) & "' AND PA57<>'Y' "
               strSql = "SELECT PA72 FROM PATENT WHERE PA01='" & SystemNumber(strTemp(4), 1) & "' AND PA02='" & SystemNumber(strTemp(4), 2) & "' AND PA03='" & SystemNumber(strTemp(4), 3) & "' AND PA04='" & SystemNumber(strTemp(4), 4) & "' AND PA57 is null "
               CheckOC2
               adoRecordset1.CursorLocation = adUseClient
               adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                   strTemp1 = Split(CheckStr(adoRecordset1.Fields(0)), ",")
                   If UBound(strTemp1) <> 0 And UBound(strTemp1) > 0 Then
                       For i = UBound(strTemp1) To 0 Step -1
                           If Val(strTemp1(i)) <> 0 Then
                               strTemp(13) = Trim(strTemp1(i))
                               Exit For
                           End If
                       Next i
                   Else
                       If UBound(strTemp1) = 0 Then
                           strTemp(13) = Trim(strTemp1(0))
                       Else
                           strTemp(13) = "0"
                       End If
                   End If
                   'Add By Sindy 2009/07/31
                   If strTemp(13) <> "0" Then
                     'Modified by Morgan 2022/6/13 +field("PA10")
                     strFeeType = PUB_GetNa20Na22Na24("" & .Fields("PA09"), "" & .Fields("PA08"), "" & .Fields("PA10"))
                     strYF15 = PUB_GetYF15("" & .Fields("PA09"), "" & .Fields("PA08"), "Y0000000", strFeeType, CDbl(strTemp(13)))
                     strTemp(13) = strTemp(13) & " " & strYF15
                   End If
                   '2009/07/31 End
               Else
                   strTemp(13) = "0"
               End If
               CheckOC2
               strSql = "SELECT MAX(CP05) FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(4), 1) & "' AND CP02='" & SystemNumber(strTemp(4), 2) & "' AND CP03='" & SystemNumber(strTemp(4), 3) & "' AND CP04='" & SystemNumber(strTemp(4), 4) & "'"
               CheckOC2
               adoRecordset1.CursorLocation = adUseClient
               adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                   strTemp(14) = CheckStr(adoRecordset1.Fields(0))
                   strSql = "SELECT MAX(CP09) FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(4), 1) & "' AND CP02='" & SystemNumber(strTemp(4), 2) & "' AND CP03='" & SystemNumber(strTemp(4), 3) & "' AND CP04='" & SystemNumber(strTemp(4), 4) & "' AND CP05=" & Val(strTemp(14)) & " AND SUBSTR(CP09, 1, 1) = 'A'"
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseServer
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.BOF = False Then
                      If IsNull(AdoRecordSet3.Fields(0)) = False Then
                          strTemp(14) = strTemp(14) + "=" + CheckStr(AdoRecordSet3.Fields(0))
                          CheckOC3
                      Else
                           CheckOC3
                           'Modify By Cheng 2002/04/12
   '                        strSQL = "SELECT MAX(CP09) FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(4), 1) & "' AND CP02='" & SystemNumber(strTemp(4), 2) & "' AND CP03='" & SystemNumber(strTemp(4), 3) & "' AND CP04='" & SystemNumber(strTemp(4), 4) & "' AND CP05=" & Val(strTemp(14)) & " AND SUBSTR(CP09,1,1)<>'A' "
                           strSql = "SELECT MAX(CP09) FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(4), 1) & "' AND CP02='" & SystemNumber(strTemp(4), 2) & "' AND CP03='" & SystemNumber(strTemp(4), 3) & "' AND CP04='" & SystemNumber(strTemp(4), 4) & "' AND CP05=" & Val(strTemp(14)) & " AND CP09>'B' "
                           AdoRecordSet3.CursorLocation = adUseClient
                           AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                           If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                               strTemp(14) = strTemp(14) + "=" + CheckStr(AdoRecordSet3.Fields(0))
                               CheckOC3
                           Else
                               strTemp(14) = "0"
                           End If
                      End If
                   Else
                       CheckOC3
                        'Modify By Cheng 2002/04/12
   '                    strSQL = "SELECT MAX(CP09) FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(4), 1) & "' AND CP02='" & SystemNumber(strTemp(4), 2) & "' AND CP03='" & SystemNumber(strTemp(4), 3) & "' AND CP04='" & SystemNumber(strTemp(4), 4) & "' AND CP05=" & Val(strTemp(14)) & " AND SUBSTR(CP09,1,1)<>'A' "
                       strSql = "SELECT MAX(CP09) FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(4), 1) & "' AND CP02='" & SystemNumber(strTemp(4), 2) & "' AND CP03='" & SystemNumber(strTemp(4), 3) & "' AND CP04='" & SystemNumber(strTemp(4), 4) & "' AND CP05=" & Val(strTemp(14)) & " AND CP09>'B' "
                       AdoRecordSet3.CursorLocation = adUseClient
                       AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                       If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                           strTemp(14) = strTemp(14) + "=" + CheckStr(AdoRecordSet3.Fields(0))
                           CheckOC3
                       Else
                           strTemp(14) = "0"
                       End If
                   End If
               Else
                   strTemp(14) = "0"
               End If
               If strTemp(14) <> "0" Then
                   strSql = "SELECT CP16 FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(4), 1) & "' AND CP02='" & SystemNumber(strTemp(4), 2) & "' AND CP03='" & SystemNumber(strTemp(4), 3) & "' AND CP04='" & SystemNumber(strTemp(4), 4) & "' AND CP05=" & Val(StringTwoString(strTemp(14), 1)) & " AND CP09='" & StringTwoString(strTemp(14), 2) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                       strTemp(14) = CheckStr(AdoRecordSet3.Fields(0))
                   Else
                       strTemp(14) = "0"
                   End If
                   CheckOC3
               End If
               CheckOC2
               
               'Add By Sindy 2016/6/20
               strTemp(14) = "" & .Fields("NP10") '智權人員
               cp(1) = SystemNumber(strTemp(4), 1)
               cp(2) = SystemNumber(strTemp(4), 2)
               cp(3) = SystemNumber(strTemp(4), 3)
               cp(4) = SystemNumber(strTemp(4), 4)
               '抓進度備註欄專利個體
               'Modified by Lydia 2016/08/25 strTemp(9) => strTemp(8)
               strTemp(8) = PUB_GetP_PA91Individual(cp(1), cp(2), cp(3), cp(4))
               '代理人名稱
               'Modified by Lydia 2016/08/25 strTemp(6) => strTemp(5)
               'Modified by Morgan 2024/6/26 年費代理人不可排除年費的發文代理人 Ex:CFP-027733--禧佩
               If ClsPDGetCasePreAgent(cp(), strText, , , , IIf(InStr("605,605,605", .Fields("NP07")) > 0, False, True)) Then
                  strTemp(5) = strText
                  If strTemp(5) <> "" Then
                     '加判斷是否為聯絡人
                     If InStr(strTemp(5), "-") > 0 Then
                        If ClsPDGetContact(strTemp(5), strText) Then
                           strTemp(5) = strText
                        End If
                     Else
                        If ClsPDGetAgent(strTemp(5), strText) Then
                           strTemp(5) = strText
                        End If
                     End If
                  End If
                  strTemp(5) = PUB_StrToStr(strTemp(5), 60) 'Added by Lydia 2016/08/25 等測試完,再將R011006擴大到60字元
               End If
               
               'Added by Lydia 2016/08/25 客戶減免
               strTemp(9) = ""
               If "" & .Fields("PA26") <> "" And "" & .Fields("PA09") <> "" Then
                   strSql = "select AD03 from applicantdiscount where ad01='" & Mid(.Fields("PA26"), 1, 8) & "' and ad02='" & .Fields("PA09") & "' "
                   intI = 1
                   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                   If intI = 1 Then
                      strTemp(9) = "" & RsTemp(0)
                   End If
               End If
               
               '前次收費情形:以案號+案件性質抓進度檔未取消收文且最大收文日之費用(點數),要減acc1u0銷帳
               strTemp(10) = 0
               'Modified by Lydia 2016/08/25 CP57->CP159
               'strSql = "SELECT cp16 - nvl(cp77,0),cp09,cp18" & _
                         " FROM caseprogress WHERE CP01='" & cp(1) & "' and CP02='" & cp(2) & "' and CP03='" & cp(3) & "' and CP04='" & cp(4) & "'" & _
                         " and cp10='" & .Fields("NP07") & "' and cp57 is null and cp16>0" & _
                         " order by cp05 desc"
               strSql = "SELECT cp16 - nvl(cp77,0),cp09,cp18" & _
                         " FROM caseprogress WHERE CP01='" & cp(1) & "' and CP02='" & cp(2) & "' and CP03='" & cp(3) & "' and CP04='" & cp(4) & "'" & _
                         " and cp10='" & .Fields("NP07") & "' and cp159=0 and cp16>0" & _
                         " order by cp05 desc"
               CheckOC3
               AdoRecordSet3.CursorLocation = adUseClient
               AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If AdoRecordSet3.RecordCount > 0 Then
                  strTemp(10) = Format(AdoRecordSet3.Fields(0), "#,##0")
                  cp(9) = AdoRecordSet3.Fields("cp09")
                  cp(18) = AdoRecordSet3.Fields("cp18")
                  strSql = "SELECT sum(nvl(a1u07,0))" & _
                            " FROM acc1u0 WHERE a1u03='" & cp(9) & "'"
                  CheckOC3
                  AdoRecordSet3.CursorLocation = adUseClient
                  AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If AdoRecordSet3.RecordCount > 0 Then
                     If Val("" & AdoRecordSet3.Fields(0)) > 0 Then '銷帳服務費
                        cp(18) = (Val(cp(18)) * 1000) - AdoRecordSet3.Fields(0)
                        cp(18) = Format(cp(18) / 1000, "0.0")
                     End If
                  End If
                  strTemp(10) = strTemp(10) & " (" & cp(18) & ")"
               End If
               CheckOC3
               '2016/6/20 END
               
               'Modified by Lydia 2016/12/08 改成模組
               'strTemp(0) = StrConv(MidB(StrConv(strTemp(0), vbFromUnicode), 1, 8), vbUnicode)
               'strTemp(1) = StrConv(MidB(StrConv(strTemp(1), vbFromUnicode), 1, 8), vbUnicode)
               'strTemp(2) = StrConv(MidB(StrConv(strTemp(2), vbFromUnicode), 1, 10), vbUnicode)
               'strTemp(3) = StrConv(MidB(StrConv(strTemp(3), vbFromUnicode), 1, 10), vbUnicode)
               'strTemp(4) = StrConv(MidB(StrConv(strTemp(4), vbFromUnicode), 1, 15), vbUnicode)
               'strTemp(6) = StrConv(MidB(StrConv(strTemp(6), vbFromUnicode), 1, 8), vbUnicode) 'Modified by Lydia 2016/08/25 strTemp(5) => strTemp(6)
               ''910430 nick 因為轉碼問題
               ''strTemp(6) = StrConv(MidB(StrConv(strTemp(6), vbFromUnicode), 1, 20), vbUnicode)
               ''Modified by Lydia 2016/11/22 先轉全形否則會抓到半個字導致語法錯誤
               ''strTemp(7) = StrConv(MidB(StrConv(strTemp(7), vbFromUnicode), 1, 8), vbUnicode)
               'strTemp(7) = StrConv(MidB(StrConv(StrConv(strTemp(7), vbWide), vbFromUnicode), 1, 8), vbUnicode)
               ''Modified by Morgan 2012/1/10 先轉全形否則會抓到半個字導致語法錯誤
               ''strTemp(8) = StrConv(MidB(StrConv(strTemp(8), vbFromUnicode), 1, 8), vbUnicode)
               'strTemp(8) = StrConv(MidB(StrConv(StrConv(strTemp(8), vbWide), vbFromUnicode), 1, 8), vbUnicode)
               'strTemp(9) = StrConv(MidB(StrConv(strTemp(9), vbFromUnicode), 1, 10), vbUnicode)
               ''Modified by Lydia 2016/08/25 擴大
               ''strTemp(10) = StrConv(MidB(StrConv(strTemp(10), vbFromUnicode), 1, 10), vbUnicode)
               'strTemp(10) = StrConv(MidB(StrConv(strTemp(10), vbFromUnicode), 1, 15), vbUnicode)
               'strTemp(11) = StrConv(MidB(StrConv(strTemp(11), vbFromUnicode), 1, 10), vbUnicode)
               'strTemp(12) = StrConv(MidB(StrConv(strTemp(12), vbFromUnicode), 1, 20), vbUnicode)
               'strTemp(13) = StrConv(MidB(StrConv(strTemp(13), vbFromUnicode), 1, 50), vbUnicode)
               ''strTemp(14) = StrConv(MidB(StrConv(strTemp(14), vbFromUnicode), 1, 12), vbUnicode)
               'strSql = "INSERT INTO R050312 VALUES ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & strUserNum & "') "
               'cnnConnection.Execute strSql
               'Add by Amy 2018/04/30 已有通知期限1913不列
               If bolHas1913(strNP01, strNP22) = False Then
               Call ProcessInsRec(strTemp)
               End If
               'end 2016/12/08
               'Added by Lydia 2016/12/08 EPC(221)已核准(PA16='1')之年費(605)期限，若該案下一程序沒有未續辦的指定國註冊費(NP07=224,NP06 is null)期限時，則請加印未閉卷的子案資料。
               If "" & .Fields("pa09") = "221" And "" & .Fields("pa16") = "1" And "" & .Fields("np07") = "605" Then
                   If PUB_ChkNPExist(cp(), "224") = False Then
                      'Added by Lydia 2017/05/22 EPC進各國要照各國年費規定，故需排除該年度不需繳費的國家
                       strExc(1) = ""
                       If "" & .Fields("PA72") = "" Then
                          If "" & .Fields("NA21") = "" Then
                             strExc(1) = "3"
                          Else
                             strExc(1) = Val("" & .Fields("NA21")) 'VB會回傳第一個數字
                          End If
                       Else
                          ArrYear = Split("" & .Fields("PA72"), ",")
                          strExc(1) = Val(ArrYear(UBound(ArrYear))) + 1
                       End If
                       If Val(strExc(1)) > 0 Then
                          strExc(1) = " and instr(','||na21||',','," & Val(strExc(1)) & ",')>0"
                       End If
                       'end 2017/05/22
                       
                      '抓子案指定國的案號、國家和代理人(指定國註冊費224或領證601)
                      'Modified by Lydia 2017/05/22 EPC進各國要照各國年費規定，故需排除該年度不需繳費的國家 + strexc(1)
                      'Modified by Lydia 2019/01/15 +CP04
                      'Modified by Morgan 2023/5/30 +249 UP註冊
                      'Modified by Morgan 2024/2/22 增加抓年費605(已進各國年費階段來所不會有領證或指定國註冊費程序 Ex:CFP-031728)
                      strSql = "select NA03,'年費' A,ST02,'' VS04,PA01||'-'||PA02||'-'||PA03||'-'||PA04 VS05,NVL(PA05,NVL(PA06,PA07)) as VS06," & _
                               "ptm03 as VS07,nvl(CU04,NVL(CU05,CU06)) as VS08,'' VS09,'' as VS10," & _
                               "'0' VS11,'' VS12,'-' VS13,'0' VS14,CP13,CP44||DECODE(CP116,NULL,'','-'||CP116) CP44,CP04 " & _
                               "From PATENT, Nation, CUSTOMER, PATENTTRADEMARKMAP, CASEPROGRESS, STAFF " & _
                               "WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03||PA04 <> '" & cp(3) & cp(4) & "' AND (PA57<>'Y' OR PA57 IS NULL) " & _
                               "AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) " & _
                               "AND CP09 IN (SELECT MAX(N.CP09) FROM CASEPROGRESS N WHERE N.CP01=PA01 AND N.CP02=PA02 AND N.CP03=PA03 AND N.CP04=PA04 AND N.CP10 in ('224','601','249','605') AND CP158>0) " & _
                               "AND PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND CP13=ST01(+) " & strExc(1) & " ORDER BY VS05 "
                      intI = 1
                      Set rsA1 = ClsLawReadRstMsg(intI, strSql)
                      If intI = 1 Then
                         If rsA1.RecordCount > 0 Then
                            rsA1.MoveFirst
                            Do While Not rsA1.EOF
                                For i = 0 To cInX
                                   strTemp(i) = CheckStr(rsA1(i))
                                Next
                                'Added by Lydia 2017/02/08 因為子案的智權人員有可能離職,所以子案的智權人員抓現在母案的智權人員
                                strTemp(2) = "" & .Fields("ST02")
                                strTemp(14) = "" & .Fields("NP10")
                                'end 2017/02/08
                                strTemp(9) = "" '個體
                                '指定國註冊費的代理人
                                'Modified by Lydia 2019/01/15 與發文一致,統一抓子案最新/最後代理人(ex.CFP-28620-0-08)
                                'strTemp(5) = "" & rsA1.Fields("cp44")
                                'Modified by Lydia 2019/01/31 EPC子案代理人依發文性質排順序
                                'If PUB_GetCP44(cp(1), cp(2), cp(3), "" & rsA1.Fields("CP04"), strTemp(5), strExc(6)) = True Then
                                If PUB_GetEPCtoCP44(cp(1), cp(2), cp(3), "" & rsA1.Fields("CP04"), strTemp(5), strExc(6)) = True Then
                                    If strExc(6) <> "" Then strTemp(5) = strTemp(5) & "-" & strExc(6) '+聯絡人
                                Else
                                    strTemp(5) = "" & rsA1.Fields("cp44")
                                End If
                                'end 2019/01/15
                                
                                If strTemp(5) <> "" And Len(strTemp(5)) >= 9 Then
                                   '加判斷是否為聯絡人
                                   If InStr(strTemp(5), "-") > 0 Then
                                      If ClsPDGetContact(strTemp(5), strText) Then
                                         strTemp(5) = strText
                                      End If
                                   Else
                                      If ClsPDGetAgent(strTemp(5), strText) Then
                                         strTemp(5) = strText
                                      End If
                                   End If
                                End If
                                'Add by Amy 2018/04/30 已有通知期限1913不列
                                If bolHas1913(strNP01, strNP22) = False Then
                                'Modify by Amy 2018/05/03 +strNP09
                                Call ProcessInsRec(strTemp, strNP09)
                                End If
                                rsA1.MoveNext
                            Loop
                         End If
                      End If
                   End If
               End If
               'end 2016/12/08
           End If 'end 'Add by Lydia 2014/11/24 CheckCPExists
           
JumpToNext: 'Added by Lydia 2020/03/05
           .MoveNext
           'k = k + 1 'Remove by Lydia 2016/12/08
           DoEvents
        Loop
    End With
Else
   InsertQueryLog (0) 'Add By Sindy 2010/12/7
   ShowNoData
   Screen.MousePointer = vbDefault
   Exit Sub
End If
CheckOC
PrintData
Screen.MousePointer = vbDefault
End Sub

'Private Sub Process1()               '下半部
'Screen.MousePointer = vbDefault
'cnnConnection.Execute "DELETE FROM R050312 WHERE ID='" & strUserNum & "' "
''系統類別
'strSQL1 = ""
'If Len(txt2(0)) <> 0 Then
'    strSQL1 = strSQL1 + " AND NP02='" & txt2(0) & "' "
'End If
'If Len(txt2(1)) <> 0 Then
'    strSQL1 = strSQL1 + " AND NP03='" & txt2(1) & "' "
'End If
'If Len(txt2(2)) <> 0 Then
'    strSQL1 = strSQL1 + " AND NP04='" & txt2(2) & "' "
'End If
'If Len(txt2(3)) <> 0 Then
'    strSQL1 = strSQL1 + " AND NP05='" & txt2(3) & "' "
'End If
'If Len(txt2(4)) <> 0 Then
'    strSQL1 = strSQL1 + " AND NP07=" & Val(txt2(4))
'End If
''Modify By Sindy 2009/07/31
''strSQL = "SELECT NA03,DECODE(PA09,'000',CPM03,CPM04) AS A," & SQLDate("NP08") & "," & SQLDate("NP09") & ",NP02||'-'||NP03||'-'||NP04||'-'||NP05,PA22,NVL(PA05,NVL(PA06,PA07)),PA08,nvl(CU04,NVL(CU05,CU06))," & SQLDate("PA10") & "," & SQLDate("PA14") & "," & SQLDate("PA21") & ",PA24,PA25 FROM NEXTPROGRESS,PATENT,CASEPROPERTYMAP,NATION WHERE PA01=NP02(+) AND PA02=NP03(+) AND PA03=NP04(+) AND PA04=NP05(+) AND np02=cpm01(+) AND TO_CHAR(NP07)=CPM02(+) AND pa09=na01(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND PA57<>'Y' " & strSQL1
''strSQL = strSQL + " union all select NA03,DECODE(SP09,'000',CPM03,CPM04) AS A," & SQLDate("NP08") & "," & SQLDate("NP09") & ",NP02||'-'||NP03||'-'||NP04||'-'||NP05,'',NVL(SP05,NVL(SP06,SP07)),'',NVL(CU04,NVL(CU05,CU06))," & SQLDate("SP10") & ",''," & SQLDate("SP12") & ",SP20,SP21 FROM NEXTPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,NATION WHERE SP01=NP02(+) AND SP02=NP03(+) AND SP03=NP04(+) AND SP04=NP05(+)  AND np02=cpm01(+) AND TO_CHAR(NP07)=CPM02(+) AND sp09=na01(+) AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND SP15<>'Y' " & strSQL1
'strSQL = "SELECT NA03,DECODE(PA09,'000',CPM03,CPM04) AS A," & SQLDate("NP08") & "," & SQLDate("NP09") & ",NP02||'-'||NP03||'-'||NP04||'-'||NP05,PA22,NVL(PA05,NVL(PA06,PA07)),PA08,nvl(CU04,NVL(CU05,CU06))," & SQLDate("PA10") & "," & SQLDate("PA14") & "," & SQLDate("PA21") & ",PA24,PA25 FROM NEXTPROGRESS,PATENT,CASEPROPERTYMAP,NATION WHERE PA01=NP02(+) AND PA02=NP03(+) AND PA03=NP04(+) AND PA04=NP05(+) AND np02=cpm01(+) AND TO_CHAR(NP07)=CPM02(+) AND pa09=na01(+) AND SUBSTR(PA26,1,8)=cu01(+) AND decode(SUBSTR(Pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) AND (PA57<>'Y' or PA57 is null) " & strSQL1
'strSQL = strSQL + " union all select NA03,DECODE(SP09,'000',CPM03,CPM04) AS A," & SQLDate("NP08") & "," & SQLDate("NP09") & ",NP02||'-'||NP03||'-'||NP04||'-'||NP05,'',NVL(SP05,NVL(SP06,SP07)),'',NVL(CU04,NVL(CU05,CU06))," & SQLDate("SP10") & ",''," & SQLDate("SP12") & ",SP20,SP21 FROM NEXTPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,NATION WHERE SP01=NP02(+) AND SP02=NP03(+) AND SP03=NP04(+) AND SP04=NP05(+)  AND np02=cpm01(+) AND TO_CHAR(NP07)=CPM02(+) AND sp09=na01(+) AND SUBSTR(sP08,1,8)=cu01(+) AND decode(SUBSTR(sP08,9,1),null,'0',substr(sp08,9,1))=cu02(+) AND (SP15<>'Y' or SP15 is null) " & strSQL1
'strSQL = strSQL + " ORDER BY NA03,A,NP08,NP09 "
'CheckOC
'k = 0
'adoRecordset.CursorLocation = adUseClient
'adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'    With adoRecordset
'        .MoveFirst
'        DoEvents
'        Do While .EOF = False
'            For i = 0 To 13
'                strTemp(i) = CheckStr(.Fields(i))
'            Next i
'            strTemp(12) = strTemp(12) + "-" & strTemp(13)
'            'Modify By Sindy 2009/07/31
'            'strSQL = "SELECT PA72 FROM PATENT WHERE PA01='" & SystemNumber(strTemp(4), 1) & "' AND PA02='" & SystemNumber(strTemp(4), 2) & "' AND PA03='" & SystemNumber(strTemp(4), 3) & "' AND PA04='" & SystemNumber(strTemp(4), 4) & "' AND PA57<>'Y' "
'            strSQL = "SELECT PA72 FROM PATENT WHERE PA01='" & SystemNumber(strTemp(4), 1) & "' AND PA02='" & SystemNumber(strTemp(4), 2) & "' AND PA03='" & SystemNumber(strTemp(4), 3) & "' AND PA04='" & SystemNumber(strTemp(4), 4) & "' AND PA57 is null "
'            CheckOC2
'            adoRecordset1.CursorLocation = adUseClient
'            adoRecordset1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'                strTemp1 = Split(CheckStr(adoRecordset1.Fields(0)), ",")
'                If UBound(strTemp1) <> 0 And UBound(strTemp1) > 0 Then
'                    For i = UBound(strTemp1) To 0 Step -1
'                        If Val(strTemp1(i)) <> 0 Then
'                            strTemp(13) = Trim(strTemp1(i))
'                            Exit For
'                        End If
'                    Next i
'                Else
'                    If UBound(strTemp1) = 0 Then
'                        strTemp(13) = Trim(strTemp1(0))
'                    Else
'                        strTemp(13) = "0"
'                    End If
'                End If
'            Else
'                strTemp(13) = "0"
'            End If
'            CheckOC2
'            strSQL = "SELECT MAX(CP05) FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(4), 1) & "' AND CP02='" & SystemNumber(strTemp(4), 2) & "' AND CP03='" & SystemNumber(strTemp(4), 3) & "' AND CP04='" & SystemNumber(strTemp(4), 4) & "'"
'            CheckOC2
'            adoRecordset1.CursorLocation = adUseClient
'            adoRecordset1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'                strTemp(14) = CheckStr(adoRecordset1.Fields(0))
'                strSQL = "SELECT MAX(CP09) FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(4), 1) & "' AND CP02='" & SystemNumber(strTemp(4), 2) & "' AND CP03='" & SystemNumber(strTemp(4), 3) & "' AND CP04='" & SystemNumber(strTemp(4), 4) & "' AND CP05=" & Val(strTemp(14)) & " AND SUBSTR(CP09, 1, 1) = 'A'"
'                CheckOC3
'                AdoRecordSet3.CursorLocation = adUseServer
'                AdoRecordSet3.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'                If AdoRecordSet3.BOF = False Then
'                   If IsNull(AdoRecordSet3.Fields(0)) = False Then
'                       strTemp(14) = strTemp(14) + "=" + CheckStr(AdoRecordSet3.Fields(0))
'                       CheckOC3
'                   Else
'                        CheckOC3
'                        'Modify By Cheng 2002/04/12
''                        strSQL = "SELECT MAX(CP09) FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(4), 1) & "' AND CP02='" & SystemNumber(strTemp(4), 2) & "' AND CP03='" & SystemNumber(strTemp(4), 3) & "' AND CP04='" & SystemNumber(strTemp(4), 4) & "' AND CP05=" & Val(strTemp(14)) & " AND SUBSTR(CP09,1,1)<>'A' "
'                        strSQL = "SELECT MAX(CP09) FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(4), 1) & "' AND CP02='" & SystemNumber(strTemp(4), 2) & "' AND CP03='" & SystemNumber(strTemp(4), 3) & "' AND CP04='" & SystemNumber(strTemp(4), 4) & "' AND CP05=" & Val(strTemp(14)) & " AND CP09>'B' "
'                        AdoRecordSet3.CursorLocation = adUseClient
'                        AdoRecordSet3.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'                        If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
'                            strTemp(14) = strTemp(14) + "=" + CheckStr(AdoRecordSet3.Fields(0))
'                            CheckOC3
'                        Else
'                            strTemp(14) = "0"
'                        End If
'                   End If
'                Else
'                    CheckOC3
'                     'Modify By Cheng 2002/04/12
''                    strSQL = "SELECT MAX(CP09) FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(4), 1) & "' AND CP02='" & SystemNumber(strTemp(4), 2) & "' AND CP03='" & SystemNumber(strTemp(4), 3) & "' AND CP04='" & SystemNumber(strTemp(4), 4) & "' AND CP05=" & Val(strTemp(14)) & " AND SUBSTR(CP09,1,1)<>'A' "
'                    strSQL = "SELECT MAX(CP09) FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(4), 1) & "' AND CP02='" & SystemNumber(strTemp(4), 2) & "' AND CP03='" & SystemNumber(strTemp(4), 3) & "' AND CP04='" & SystemNumber(strTemp(4), 4) & "' AND CP05=" & Val(strTemp(14)) & " AND CP09>'B' "
'                    AdoRecordSet3.CursorLocation = adUseClient
'                    AdoRecordSet3.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'                    If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
'                        strTemp(14) = strTemp(14) + "=" + CheckStr(AdoRecordSet3.Fields(0))
'                        CheckOC3
'                    Else
'                        strTemp(14) = "0"
'                    End If
'                End If
'            Else
'                strTemp(14) = "0"
'            End If
'            If strTemp(14) <> "0" Then
'                strSQL = "SELECT CP16 FROM CASEPROGRESS WHERE CP01='" & SystemNumber(strTemp(4), 1) & "' AND CP02='" & SystemNumber(strTemp(4), 2) & "' AND CP03='" & SystemNumber(strTemp(4), 3) & "' AND CP04='" & SystemNumber(strTemp(4), 4) & "' AND CP05=" & Val(StringTwoString(strTemp(14), 1)) & " AND CP09='" & StringTwoString(strTemp(14), 2) & "' "
'                CheckOC3
'                AdoRecordSet3.CursorLocation = adUseClient
'                AdoRecordSet3.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
'                    strTemp(14) = CheckStr(AdoRecordSet3.Fields(0))
'                Else
'                    strTemp(14) = "0"
'                End If
'                CheckOC3
'            End If
'            CheckOC2
'            strTemp(0) = StrConv(MidB(StrConv(strTemp(0), vbFromUnicode), 1, 8), vbUnicode)
'            strTemp(1) = StrConv(MidB(StrConv(strTemp(1), vbFromUnicode), 1, 8), vbUnicode)
'            strTemp(2) = StrConv(MidB(StrConv(strTemp(2), vbFromUnicode), 1, 10), vbUnicode)
'            strTemp(3) = StrConv(MidB(StrConv(strTemp(3), vbFromUnicode), 1, 10), vbUnicode)
'            strTemp(4) = StrConv(MidB(StrConv(strTemp(4), vbFromUnicode), 1, 15), vbUnicode)
'            strTemp(5) = StrConv(MidB(StrConv(strTemp(5), vbFromUnicode), 1, 8), vbUnicode)
'            '910430 nick 因為轉碼問題
'            'strTemp(6) = StrConv(MidB(StrConv(strTemp(6), vbFromUnicode), 1, 20), vbUnicode)
'            strTemp(7) = StrConv(MidB(StrConv(strTemp(7), vbFromUnicode), 1, 8), vbUnicode)
'            strTemp(8) = StrConv(MidB(StrConv(strTemp(8), vbFromUnicode), 1, 8), vbUnicode)
'            strTemp(9) = StrConv(MidB(StrConv(strTemp(9), vbFromUnicode), 1, 10), vbUnicode)
'            strTemp(10) = StrConv(MidB(StrConv(strTemp(10), vbFromUnicode), 1, 10), vbUnicode)
'            strTemp(11) = StrConv(MidB(StrConv(strTemp(11), vbFromUnicode), 1, 10), vbUnicode)
'            strTemp(12) = StrConv(MidB(StrConv(strTemp(12), vbFromUnicode), 1, 20), vbUnicode)
'            strTemp(13) = StrConv(MidB(StrConv(strTemp(13), vbFromUnicode), 1, 8), vbUnicode)
'            strTemp(14) = StrConv(MidB(StrConv(strTemp(14), vbFromUnicode), 1, 12), vbUnicode)
'            If Len(Trim(txt2(5))) <> 0 Then
'                If Val(txt2(5)) = Val(strTemp(14)) Then
'                    strSQL = "INSERT INTO R050312 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & strUserNum & "') "
'                    cnnConnection.Execute strSQL
'                End If
'            Else
'                strSQL = "INSERT INTO R050312 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & strUserNum & "') "
'                cnnConnection.Execute strSQL
'            End If
'            .MoveNext
'            k = k + 1
'            DoEvents
'        Loop
'    End With
'Else
'   ShowNoData
'   Screen.MousePointer = vbDefault
'   Exit Sub
'End If
'CheckOC
'PrintData
'Screen.MousePointer = vbDefault
'End Sub
 
Private Sub PrintData()
'Modify By Cheng 2002/10/01
'strSQL = "SELECT * FROM R050312 WHERE ID='" & strUserNum & "' ORDER BY 1,2,3,4"
'Modify by Amy 2018/04/30 止日月份的報表與其他月份分開列印
Dim strOldYM As String 'Add by Amy 2018/04/30
Dim strColList As String 'Added by Morgan 2023/10/17

strColList = "R011001,R011002,R011003,R011004,R011005,R011006,R011007,R011008,R011009,R011010,R011011,R011014" 'Added by Morgan 2023/10/17

strSql = ""
If Left(Val(txt1(1) + 19110000), 6) <> Left(Val(txt1(2) + 19110000), 6) Then
    strSql = " Decode(SubStr(Replace(Replace(R011004,'/',''),'*','')+19110000,1,6),'" & Left(Val(txt1(2) + 19110000), 6) & "','2','1'), "
End If

If txt1(15) = "1" Then '依本所案號排序
   'strSql = "SELECT * FROM R050312 WHERE ID='" & strUserNum & "' ORDER BY 5"
   'Modified by Morgan 2023/10/17
   'strSql = "SELECT * FROM R050312 WHERE ID='" & strUserNum & "' ORDER BY " & strSql & "R011005"
   strSql = "SELECT " & strColList & " FROM R050312 WHERE ID='" & strUserNum & "' ORDER BY " & strSql & "R011005"
   'end 2023/10/17
Else '依智權人員排序
   'Modify By Sindy 2016/10/14 排序:所別,區別,智權人員,客戶
   'strSql = "SELECT * FROM R050312 WHERE ID='" & strUserNum & "' ORDER BY R011015"
   'Modified by Lydia 2016/12/08 +本所案號R011005
   'Modified by Morgan 2023/10/17
   'strSql = "SELECT * FROM R050312,staff WHERE ID='" & strUserNum & "' and R011015=st01(+) ORDER BY " & strSql & "st06,st15,st01,R011008,R011005"
   strSql = "SELECT " & strColList & " FROM R050312,staff WHERE ID='" & strUserNum & "' and R011015=st01(+) ORDER BY " & strSql & "st06,st15,st01,R011008,R011005"
   'end 2023/10/17
   '2016/10/14 END
End If
'end 2018/04/30
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        Page = 1
        PrintTitle
        Do While .EOF = False
            'Modify By Sindy 2009/07/31
            'For i = 0 To 14
            'Modified by Morgan 2023/10/17
            'For i = 0 To 13
            For i = 0 To 11
            'end 2023/10/17
               strTemp(i) = CheckStr(.Fields(i))
            Next i
            'StrTemp(9) = "88/01/01"
            'StrTemp(10) = "88/01/01"
            'StrTemp(11) = "88/01/01"
            'StrTemp(12) = "19990101-19991231"
            'Modified by Lydia 2016/08/25
            'strTemp(6) = StrConv(MidB(StrConv(strTemp(6), vbFromUnicode), 1, 28), vbUnicode)
            strTemp(5) = StrConv(MidB(StrConv(strTemp(5), vbFromUnicode), 1, 60), vbUnicode)
            
            'Remove by Lydia 2016/08/25
            'If iPrint > 10000 Then
            '   Printer.Font.Size = 8
            '   Printer.CurrentX = 500
            '   Printer.CurrentY = iPrint
            '   Printer.Print String(190, "-")
            '   Printer.NewPage
            '   Page = Page + 1
            '   PrintTitle
            'End If
            'Modify by Amy 2018/05/04 止月報表分開列印
            If strOldYM <> "" And strOldYM <> Left(Val(Replace(Replace(strTemp(3), "/", ""), "*", "")) + 19110000, 6) _
              And Left(Val(Replace(Replace(strTemp(3), "/", ""), "*", "")) + 19110000, 6) = Left(Val(txt1(2)) + 19110000, 6) Then
                Printer.CurrentX = ciStartX
                Printer.CurrentY = iPrint
            
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            '法限有*改不顯示期限(表示子案無期限帶母案期限 for 排序)
            If InStr(strTemp(3), "*") > 0 Then strTemp(3) = ""
            'end 2018/05/04
            PrintDetil
            strOldYM = Left(Val(Replace(Replace(.Fields(3), "/", ""), "*", "")) + 19110000, 6)   'Add by Amy 2018/04/30 止月報表分開列印
            .MoveNext
        Loop
    End With
    Printer.EndDoc
    ShowPrintOk
'Added by Lydia 2020/03/05
Else  '無暫存檔資料
   InsertQueryLog (0)
   ShowNoData
'end 2020/03/05
End If
CheckOC
End Sub
 
Sub PrintTitle()
'Added by Lydia 2016/8/25
Dim Str01 As String
Dim x1 As Integer
Dim iPos As Integer

GetPleft

'Modified by Lydia 2016/08/25 改寫法
'iPrint = 500
'Printer.Orientation = 2
'Printer.Font.Name = "細明體"
'Printer.Font.Size = 22
'Printer.Font.Bold = True
'Printer.Font.Underline = True
'Printer.CurrentX = 6300
'Printer.CurrentY = iPrint
'Printer.Print "期限通知管制表"
'iPrint = iPrint + 500
'Printer.Font.Size = 12
'Printer.Font.Bold = False
'Printer.Font.Underline = False
'If Option1(0).Value = True Then
'    Printer.CurrentX = 6500
'    Printer.CurrentY = iPrint
'    Printer.Print "日期：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
'End If
'iPrint = iPrint + 300
'Printer.CurrentX = 500
'Printer.CurrentY = iPrint
'Printer.Print "列印人：" & strUserName
'Printer.CurrentX = 13000
'Printer.CurrentY = iPrint
'Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
'iPrint = iPrint + 300
'Printer.CurrentX = 13000
'Printer.CurrentY = iPrint
'Printer.Print "頁    次：" & str(Page)
'iPrint = iPrint + 300
'Printer.Font.Size = 8
'Printer.CurrentX = 500
'Printer.CurrentY = iPrint
'Printer.Print String(190, "-")
'Printer.Font.Size = 12
'iPrint = iPrint + 300
'Printer.Font.Size = 8
'Printer.CurrentX = PLeft(0)
'Printer.CurrentY = iPrint
'Printer.Print "申請國家"
'Printer.CurrentX = PLeft(1)
'Printer.CurrentY = iPrint
'Printer.Print "下一程序"
'Printer.CurrentX = PLeft(2)
'Printer.CurrentY = iPrint
'Printer.Print "智權人員" 'Modify By Sindy 2016/6/20 "本所期限"
'Printer.CurrentX = PLeft(3)
'Printer.CurrentY = iPrint
'Printer.Print "法定期限"
'Printer.CurrentX = PLeft(4)
'Printer.CurrentY = iPrint
'Printer.Print "本所案號"
'Printer.CurrentX = PLeft(5)
'Printer.CurrentY = iPrint
'Printer.Print "專利號數"
'Printer.CurrentX = PLeft(6)
'Printer.CurrentY = iPrint
'Printer.Print "代理人" 'Modify By Sindy 2016/6/20 "專利名稱"
'Printer.CurrentX = PLeft(7)
'Printer.CurrentY = iPrint
'Printer.Print "專利種類"
'Printer.CurrentX = PLeft(8)
'Printer.CurrentY = iPrint
'Printer.Print "客戶名稱"
'Printer.CurrentX = PLeft(9)
'Printer.CurrentY = iPrint
'Printer.Print "個體" 'Modify By Sindy 2016/6/20 "申請日"
'Printer.CurrentX = PLeft(10) - Printer.TextWidth("前次收費情形") 'Modify By Sindy 2016/6/20 "公告日"
'Printer.CurrentY = iPrint
'Printer.Print "前次收費情形"
'Printer.CurrentX = PLeft(11)
'Printer.CurrentY = iPrint
'Printer.Print "發證日"
'Printer.CurrentX = PLeft(12)
'Printer.CurrentY = iPrint
'Printer.Print "專用期間"
'Printer.CurrentX = PLeft(13)
'Printer.CurrentY = iPrint
'Printer.Print "繳費年度"
''Modify By Sindy 2009/07/31
''Printer.CurrentX = PLeft(14)
''Printer.CurrentY = iPrint
''Printer.Print "上次收費金額"
'iPrint = iPrint + 300
'Printer.CurrentX = 500
'Printer.CurrentY = iPrint
'Printer.Print String(190, "-")
'iPrint = iPrint + 300

iPrint = ciStartY

Printer.Font.Size = ciTitleFontSize
Printer.Font.Bold = True
Printer.Font.Underline = True
Str01 = "期限通知管制表"
Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(Str01)) / 2
Printer.CurrentY = iPrint
Printer.Print Str01

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.Font.Bold = False
PrintNewLine
PrintNewLine

Printer.CurrentX = ciStartX
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName

If Option1(0).Value = True Then
   Str01 = "日期：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
   'Modified by Lydia 2020/03/05 記錄位置
   'Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(Str01)) / 2
   iPos = (Printer.ScaleWidth - Printer.TextWidth(Str01)) / 2
   Printer.CurrentX = iPos
   'end 2020/03/05
   Printer.CurrentY = iPrint
   Printer.Print Str01
End If

x1 = Printer.ScaleWidth - Printer.TextWidth(String(12, "　"))
Printer.CurrentX = x1
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

PrintNewLine
'Added by Lydia 2020/03/05
If ChkPerson.Value = 1 And txt1(16).Text <> "" Then
     Printer.CurrentX = iPos
     Printer.CurrentY = iPrint
     Printer.Print "管制人：" & GetStaffName(txt1(16).Text, True)
End If
'end 2020/03/05
Printer.CurrentX = x1
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & Printer.Page

Printer.Font.Size = ciFontSize
PrintNewLine
Printer.CurrentX = ciStartX
Printer.CurrentY = iPrint
Printer.Print String((Printer.ScaleWidth - ciStartX * 2) / Printer.TextWidth("-"), "-")
PrintNewLine

'列印欄位抬頭
For iPos = 0 To cInX
    If PTitle(iPos) <> "" And PTitle(iPos) <> "結束" Then
       Printer.CurrentX = PLeft(iPos)
       Printer.CurrentY = iPrint
       Printer.Print PTitle(iPos)
    ElseIf iPos > 1 Then
        x1 = iPos '結束
        Exit For
    End If
Next

PrintNewLine
Printer.CurrentX = ciStartX
Printer.CurrentY = iPrint
Printer.Print String((Printer.ScaleWidth - ciStartX * 2) / Printer.TextWidth("-"), "-")
PrintNewLine

End Sub

Sub PrintDetil()
 'Added by Lydia 2016/08/25
Dim tmpArr As Variant
Dim strP As String

tmpArr = Empty
tmpArr = Split(strTitle2, ",")

'Modify By Sindy 2009/07/31
'For i = 0 To 14
'Modified by Morgan 2023/10/17
'For i = 0 To 13
For i = 0 To 11
'end 2023/10/17
'    If i = 13 Then
'        Printer.CurrentX = PLeft(13) - Printer.TextWidth(strTemp(i))
'        Printer.CurrentY = iPrint
'        Printer.Print strTemp(i)
'    Else
'        If i = 14 Then
'            Printer.CurrentX = PLeft(14) - Printer.TextWidth(strTemp(i))
'            Printer.CurrentY = iPrint
'            Printer.Print strTemp(i)
'        Else

            'Added by Lydia 2016/08/25 控制字串長度
            strP = convForm(strTemp(i), Val(tmpArr(i + 1)) * 2)
            If PLeft(i) + Printer.TextWidth(strP) > PLeft(i + 1) Then strP = convForm(strTemp(i), (Val(tmpArr(i + 1)) * 2) - 4)
            
            'Add By Sindy 2016/6/20 10.前次收費情形
            If i = 10 Then
               'Modifiedby Lydia 2016/08/25 置右
               'Printer.CurrentX = PLeft(i) - Printer.TextWidth(strTemp(i))
               Printer.CurrentX = PLeft(i + 1) - Printer.TextWidth(strTemp(i)) - ciColGap
               
            'Added by Lydia 2016/08/25 客戶減免 置中
            ElseIf i = 9 Then
               Printer.CurrentX = PLeft(i) + (PLeft(i + 1) - PLeft(i) - ciColGap) / 2
               
            Else
            '2016/6/20 END
               Printer.CurrentX = PLeft(i)
            End If
            Printer.CurrentY = iPrint
            'Modified by Lydia 2016/08/25
            'Printer.Print strTemp(i)
            Printer.Print strP
'        End If
'    End If
Next i

'Modified by Lydia 2016/08/25
'iPrint = iPrint + 300
PrintNewLine

End Sub

Sub GetPleft()

If strTitle = "" Then 'Added by Lydia 2016/08/25
'Modified by Lydia 2016/08/25
'Erase PLeft
'PLeft(0) = 500
'PLeft(1) = 1300
'PLeft(2) = 2200
'PLeft(3) = 3000
'PLeft(4) = 3900
'PLeft(5) = 5200
'PLeft(6) = 5900
'PLeft(7) = 8200
'PLeft(8) = 9000
'PLeft(9) = 9800
'PLeft(10) = 11400
'PLeft(11) = 11500
'PLeft(12) = 12200
'PLeft(13) = 13800

   'Modified by Morgan 2023/10/17
   'strTitle = "申請國家,下一程序,智權人員,法定期限,本所案號,代理人,專利種類,客戶名稱,個體,客戶減免,前次收費情形,發證日,專用期間,繳費年度,結束"
   'strTitle2 = "0,4,4,4,5,8,18,4,4,3,4,6,5,10,15"
   strTitle = "申請國家,下一程序,智權人員,法定期限,本所案號,代理人,專利種類,客戶名稱,個體,客戶減免,前次收費情形,已繳費年度,結束"
   strTitle2 = "0,4,4,4,5,8,18,4,4,3,4,6,15"
   'end 2023/10/17
   Call SettingPrtSet
End If
'end 2016/08/25

End Sub
 
Private Sub Form_Activate()
   'Added by Sindy 2017/12/29
   If m_strIR01 <> "" And m_Done = False Then
      txt2(0).Text = m_strCP01
      txt2(1).Text = m_strCP02
      txt2(2).Text = m_strCP03
      txt2(3).Text = m_strCP04
      Option1(0).Enabled = False
      Option1(1).Value = True
      Option2(0).Value = True
      Option2(1).Enabled = False
      'cmdok(0).Value = True
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/29 END
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
'Added by Lydia 2016/01/04 合併催函-被合併的金額
lblMerge(0).Visible = False: lblMerge(1).Visible = False
txt2(8).Visible = False: txt2(9).Visible = False

'Modified by Morgan 2018/10/22 CFP程序系統別預設CFP,法定期限起日改系統月提前6個月
'txt1(0) = GetSystemKindByNick
'txt1(1) = Val(Left(CompDate(1, -3, strSrvDate(1)), 6)) - 191100 & "01" 'Add by Amy 2018/05/03 預設法定期限起日為系統月前三個月
If Pub_strUserST05 = "83" Or Pub_strUserST05 = "85" Then
   txt1(0) = "CFP"
   txt1(1) = Val(Left(CompDate(1, -6, strSrvDate(1)), 6)) - 191100 & "01"
Else
   txt1(0) = GetSystemKindByNick
   txt1(1) = Val(Left(CompDate(1, -3, strSrvDate(1)), 6)) - 191100 & "01" 'Add by Amy 2018/05/03 預設法定期限起日為系統月前三個月
End If
'end 2018/10/22

lbl1.Caption = "" 'Added by Lydia 2017/05/25
lbl1.Tag = lbl1 'Added by Morgan 2025/3/13
'Added by Lydia 2020/03/05
If strSrvDate(1) >= CFP業務區劃分啟用日 Then
    ChkPerson.Value = 1
    txt1(16).Text = strUserNum
    Call txt1_Validate(16, False)
Else
    ChkPerson.Visible = False
    Label9.Visible = False: txt1(16).Visible = False
    lblSalesName.Visible = False
End If
'end 2020/03/05

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050312 = Nothing
End Sub

Private Sub lblFee1_Click()
   If lblFee1.Tag = "Y" Then
      SetOldPrice True
   End If
End Sub

Private Sub SetOldPrice(bolShow As Boolean)
   'Add by Morgan 2010/6/11
   If txt2(4) = "119" Or txt2(4) = "910" Then Exit Sub
   
   PUB_LabelActive lblFee1, lblFee1s, False
   If txt2(0) & txt2(1) & txt2(2) & txt2(3) <> "" And Len(txt2(4)) = 3 Then
      txt2(2) = Right("0" & txt2(2), 1)
      txt2(3) = Right("00" & txt2(3), 2)
      'Modified by Lydia 2017/05/12 +pa21
      strExc(0) = "SELECT pa26,pa09,pa08,pa21 FROM patent" & _
            " WHERE " & ChgPatent(txt2(0) & txt2(1) & txt2(2) & txt2(3))
            
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_PA26 = "" & RsTemp("pa26")
         m_PA09 = "" & RsTemp("pa09")
         m_PA08 = "" & RsTemp("pa08")
         m_PA21 = "" & RsTemp("pa21") 'Added by Lydia 2017/05/12 發證日
         m_dblYear = 0
         If txt2(4) = "605" Or txt2(4) = "606" Or txt2(4) = "607" Then
            pa(1) = txt2(0): pa(2) = txt2(1): pa(3) = txt2(2): pa(4) = txt2(3)
            'Modified by Lydia 2016/08/22 傳回空白
            'm_dblYear = PUB_GetNextYear(pa)
            m_dblYear = Val(PUB_GetNextYear(pa))
         End If
         If PUB_GetOldPrice(m_PA26, m_PA09, m_PA08, txt2(4), RsTemp, , , , m_dblYear) = True Then
            PUB_LabelActive lblFee1, lblFee1s
            If bolShow = True Then
               Set frm880014.grdDataList.Recordset = RsTemp
               Set frm880014.fmParent = Me
               frm880014.Show vbModal
            End If
         End If
      End If
   End If
End Sub

Private Sub Option1_Click(Index As Integer)
'Add by Amy 2018/04/30 選擇管制表預設法定期限起日為系統月1日
txt1(1) = ""
Select Case Index
Case 0
     txt1(1) = Left(strSrvDate(2), 5) & "01" 'Add by Amy 2018/04/30 預設法定期限起日為系統月1日
     txt1(0).SetFocus
     txt1_GotFocus (0)
Case 1
      'Modify By Cheng 2002/10/11
'     txt2(0).SetFocus
'     txt2_GotFocus (0)
   If Me.Option2(0).Value Then
      Option2_Click 0
   Else
      Option2_Click 1
   End If
Case Else
End Select
End Sub

Private Sub Option2_Click(Index As Integer)
   'Add By Cheng 2002/10/11
   Select Case Index
   Case 0 '本所案號
      Me.txt2(0).Enabled = True
      Me.txt2(1).Enabled = True
      Me.txt2(2).Enabled = True
      Me.txt2(3).Enabled = True
      txt2(0).SetFocus
      txt2_GotFocus (0)
      Me.txt1(11).Enabled = False
      Me.txt1(12).Enabled = False
      Me.txt1(13).Enabled = False
      Me.txt1(14).Enabled = False
   Case 1 '法定期限
      Me.txt1(11).Enabled = True
      Me.txt1(12).Enabled = True
      Me.txt1(13).Enabled = True
      Me.txt1(14).Enabled = True
      Me.txt1(11).SetFocus
      txt1_GotFocus 11
      Me.txt2(0).Enabled = False
      Me.txt2(1).Enabled = False
      Me.txt2(2).Enabled = False
      Me.txt2(3).Enabled = False
   End Select
End Sub

Private Sub txt_GotFocus()
txt.SelStart = 0
txt.SelLength = Len(txt)
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
 KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt_LostFocus()
Select Case Val(txt)
Case 1, 2
Case Else
    s = MsgBox("列印格式錯誤, 只能輸入 1 或 2 !!", , "USER 輸入錯誤")
    txt.SetFocus
    txt.SelStart = 0
    txt.SelLength = Len(txt)
    Exit Sub
End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
 KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
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
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
    Next i
'Modify By Cheng 2002/10/11
'Case 2, 4, 6
Case 2, 4, 6, 12, 14
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
Case 8
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
      End If
   Else
      blnClkSure = False
   End If
Case 10
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
      End If
   Else
      blnClkSure = False
   End If
Case Else
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
'Modify By Cheng 2002/10/11
'Case 1, 2 '本所期限
Case 1, 2, 11, 12 '本所期限
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
'Add By Sindy 2016/6/20
Case 15
   If Me.txt1(Index) <> "1" And Me.txt1(Index) <> "2" Then
      s = MsgBox("排序只能輸入 1 或 2 !!", , "USER 輸入錯誤")
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
'2016/6/20 END
'Added by Lydia 2020/03/05
Case 16   '管制人
   If txt1(Index).Tag <> txt1(Index).Text Then
     strExc(0) = GetStaffName(txt1(Index).Text, True)
     lblSalesName.Caption = strExc(0)
   End If
   txt1(Index).Tag = txt1(Index).Text
'end 2020/03/05
End Select
End Sub

Private Sub txt2_GotFocus(Index As Integer)
   TextInverse txt2(Index)
End Sub

Private Sub txt2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
   'Add by Morgan 2008/4/30
   'Added by Lydia 2016/01/04 +合併金額
   Case 5, 7, 8, 9, 6, 10
      If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape) Then
         KeyAscii = 0
         Beep
      End If
   End Select
End Sub

Private Sub txt2_LostFocus(Index As Integer)
Select Case Index
Case 4
     lbl1 = GetPrjState6HM("CFP", txt2(4))
     ChkAlert 'Added by Morgan 2025/3/13
     If Me.txt2(4).Text <> "" Then
         If Me.lbl1.Caption = "" Then
            MsgBox "下一程序代號輸入錯誤!!!", vbExclamation + vbOKOnly
            Me.txt2(4).SetFocus
            txt2_GotFocus 4
            Exit Sub
         End If
        'Added by Lydia 2016/09/13 美國和法國檢查"個案之個體狀態"與"客戶減免身份"個體是不同
         Call ChkEntity
         'Add by Morgan 2008/4/30
         If Option2(0).Value = True Then
            SetFees
            ChkOverFee 'Added by Morgan 2020/3/17 檢查是否可輸入超項費超頁費
            'Added by Lydia 2016/01/04 檢查是否有兩個需要合併的催函
            Call ChkLetterMerge
            
         Else
            lblMerge(0).Visible = False: lblMerge(1).Visible = False
            txt2(8).Visible = False: txt2(9).Visible = False
         'end 2016/01/04
         End If
     End If
Case 3
   'Added by Lydia 2016/09/13 美國和法國檢查"個案之個體狀態"與"客戶減免身份"個體是不同
   Call ChkEntity
   'Added by Morgan 2025/3/18
   lbl1 = GetPrjState6HM("CFP", txt2(4))
   ChkAlert
   'end 2025/3/18
   'Add by Morgan 2008/4/30
   If Option2(0).Value = True Then
      PUB_CheckCaseBillMemo txt2(0) & txt2(1) & txt2(2) & txt2(3) 'Add by Morgan 2008/6/9
      SetFees
   End If
Case Else
End Select
End Sub

''抓下次繳費次數
'Private Sub Getnexttimes()
'   Dim strPA08 As String
'   Dim strPA09 As String
'   Dim strPA72 As String
'   Dim varPA72 As Variant
'   Dim varRef As Variant
'   Dim strKey(5) As String
'   Dim strCaseFee(1 To 2) As String
'   Dim bFind As Boolean
'   Dim ii As Integer
'
'   m_Nexttimes = ""
'   strExc(0) = "SELECT PA08,PA09,PA72,PA91 FROM PATENT WHERE PA01='" & m_NP02 & "' And PA02='" & m_NP03 & "' AND PA03='" & m_NP04 & "' AND PA04='" & m_NP05 & "'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      If IsNull(RsTemp.Fields("PA08")) = False Then strPA08 = RsTemp.Fields("PA08")
'      If IsNull(RsTemp.Fields("PA09")) = False Then strPA09 = RsTemp.Fields("PA09")
'      If IsNull(RsTemp.Fields("PA72")) = False Then strPA72 = RsTemp.Fields("PA72")
'      m_PA91 = "" & RsTemp.Fields("PA91") 'Add by Morgan 2008/4/30
'      strKey(0) = ""
'      strKey(1) = m_NP02
'      strKey(2) = m_NP03
'      strKey(3) = m_NP04
'      strKey(4) = m_NP05
'      bFind = GetMoneyDate(strPA08, strPA09, strKey, strCaseFee(1), strCaseFee(2))
'      If bFind Then
'         ' 尋找下次繳費的位置
'         '若已有繳費記錄
'         If IsEmptyText(strPA72) = False Then
'              varPA72 = Split(strPA72, ",")
'              If IsEmptyText(strCaseFee(2)) = False Then
'                 varRef = Split(strCaseFee(2), ",")
'                 For ii = LBound(varRef) To UBound(varRef)
'                       If Val(varPA72(UBound(varPA72))) < Val(varRef(ii)) Then
'                            m_Nexttimes = ii + 1
'                            Exit For
'                       End If
'                 Next ii
'              End If
'          '若無繳費記錄
'          Else
'              If IsEmptyText(strCaseFee(2)) = False Then
'                  m_Nexttimes = 1
'              End If
'         End If
'      End If
'   End If
'End Sub

Private Sub SetFees()
   If txt2(0) & txt2(1) & txt2(2) & txt2(3) <> "" Then
      strPA16 = "" 'Added by Lydia 2016/11/11
      '年費或延展費時預設費用,點數
      If (txt2(4) = "605" Or txt2(4) = "606" Or txt2(4) = "607") Then
         m_NP02 = txt2(0)
         m_NP03 = txt2(1)
         m_NP04 = Right("0" & txt2(2), 1)
         m_NP05 = Right("00" & txt2(3), 2)
         
         'Modify By Sindy 2009/07/29
         'Getnexttimes   '取得下次繳費次數
         m_Nexttimes = PUB_Getnexttimes(m_NP02, m_NP03, m_NP04, m_NP05, m_strYear, m_PA91)
         '2009/07/29 End

'         'add by sonia 2024/7/16 英國201只有脫歐後衍生的的再註冊設計案才能帶金額
         If m_PA09 = "201" And InStr(m_PA91, "歐盟案案號：CFP") = 0 Then
            GoTo Nextstep
         End If
'         'end 2024/7/16

         '大個體
         'Modified by Morgan 2023/3/29
         'If InStr(m_PA91, "大個體") > 0 Then
         m_PA179 = PUB_GetEntityType(m_NP02, m_NP03, m_NP04, m_NP05)
         If m_PA179 = "1" Then
         'end 2023/3/29
            'Modify By Sindy 2009/07/29 把m_Nexttimes改為m_strYear
            'Modified by Lydia 2016/11/11 +PA16
            strExc(0) = "SELECT nvl(YF06,0),nvl(YF07,0),PA16 FROM patent,PATENTYEARFEE" & _
               " WHERE " & ChgPatent(txt2(0) & txt2(1) & txt2(2) & txt2(3)) & _
               " and YF01(+)=pa09 AND YF02(+)=pa08 AND " & _
               " YF03='Y00000002' AND YF04=" & txt2(4) & " AND YF05=" & CNULL(m_strYear)
               
         'Added by Morgan 2013/3/20
         'Modified by Lydia 2016/11/11 +PA16
         'Modified by Morgan 2023/3/29
         'ElseIf InStr(m_PA91, "微個體") > 0 Then
         ElseIf m_PA179 = "3" Then
         'end 2023/3/29
            strExc(0) = "SELECT nvl(YF06,0),nvl(YF07,0),PA16 FROM patent,PATENTYEARFEE" & _
               " WHERE " & ChgPatent(txt2(0) & txt2(1) & txt2(2) & txt2(3)) & _
               " and YF01(+)=pa09 AND YF02(+)=pa08 AND " & _
               " YF03='Y00000003' AND YF04=" & txt2(4) & " AND YF05=" & CNULL(m_strYear)
         'end 2013/3/20
         Else
            'Modify By Sindy 2009/07/29 把m_Nexttimes改為m_strYear
            'Modified by Lydia 2016/11/11 +PA16
            strExc(0) = "SELECT nvl(YF06,0),nvl(YF07,0),PA16 FROM patent,PATENTYEARFEE" & _
               " WHERE " & ChgPatent(txt2(0) & txt2(1) & txt2(2) & txt2(3)) & _
               " and YF01(+)=pa09 AND YF02(+)=pa08 AND " & _
               " YF03='Y00000000' AND YF04=" & txt2(4) & " AND YF05=" & CNULL(m_strYear)
         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Modify by Morgan 2009/7/17 專利年費資料檔因為加年費年度說明欄位有新增沒有費用的資料
            txt2(5) = RsTemp.Fields(0) + RsTemp.Fields(1)
            txt2(7) = RsTemp.Fields(0) / 1000
            If txt2(5) = "0" Then txt2(5) = ""
            If txt2(7) = "0" Then txt2(7) = ""
            strPA16 = "" & RsTemp.Fields("PA16") 'Added by Lydia 2016/11/11
         End If
      '2009/3/5 add by sonia
      Else
         txt2(5) = ""
         txt2(7) = ""
      '2009/3/5 end
      End If
Nextstep:
      'Add by Morgan 2008/5/16
      If Len(txt2(4)) = 3 Then
         m_InputEPC = False 'Added by Lydia 2016/09/29
         SetOldPrice False '查詢參考報價資料及設定案件資料的全域變數
      End If
      'end 2008/5/16
   End If
   
End Sub

'Add by Morgan 2008/5/1
Public Sub InsExpField1(ByVal strNP07 As String, ByVal strPA09 As String, ByVal ET03 As String, Optional strNP01 As String, Optional strNP22 As String)
   Dim strTxt(1 To 21) As String, iStep As Integer
   
   iStep = 1
   
   strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
      "VALUES ('" & strNP01 & "'," & strNP22 & ",'下一程序業務員','" & m_st02 & "')"
   iStep = iStep + 1
   
   If txt2(4) <> "" Then
      'Modified by Lydia 2016/01/04
'      strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
'         "VALUES ('" & strNP01 & "'," & strNP22 & ",'下一程序','" & txt2(4) & "')"
'      iStep = iStep + 1
'
'      strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
'         "VALUES ('" & strNP01 & "'," & strNP22 & ",'下一程序名稱','" & LBL1.Caption & "')"
'      iStep = iStep + 1
      strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'下一程序','" & txt2(4) & IIf(bolMerge = True, "," & mNP07_2(0), "") & "')"
      iStep = iStep + 1
      
      strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'下一程序名稱','" & lbl1.Caption & IIf(bolMerge = True, "、" & mNP07_2(1), "") & "')"
      iStep = iStep + 1
   End If
   
   'Added by Lydia 2016/01/04
'   If bolMerge Then
'      strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
'         "VALUES ('" & strNP01 & "'," & strNP22 & ",'下一程序','" & mNP07_2(0) & "')"
'      iStep = iStep + 1
'
'      strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
'         "VALUES ('" & strNP01 & "'," & strNP22 & ",'下一程序名稱','" & mNP07_2(1) & "')"
'      iStep = iStep + 1
'   End If
   strExc(1) = "" 'Added by Morgan 2016/8/10
   '馬來西亞新型合併定稿有2個法限
   If bolMerge = True And strPA09 = "018" Then
      If mNP09_1 <> "" Then
         strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
                "VALUES ('" & strNP01 & "'," & strNP22 & ",'" & feeTit1 & "法定期限" & "','" & mNP09_1 & "')"
         iStep = iStep + 1
      End If
      If mNP09_2 <> "" Then
         strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
                "VALUES ('" & strNP01 & "'," & strNP22 & ",'" & feeTit2 & "法定期限" & "','" & mNP09_2 & "')"
         iStep = iStep + 1
      End If
      strExc(1) = IIf(mNP09_1 = "", mNP09_2, IIf(mNP09_2 = "", mNP09_1, IIf(mNP09_1 > mNP09_2, mNP09_2, mNP09_1))) 'Added by Morgan 2016/8/10
      '合併催函抓最早法限
   ElseIf bolMerge = True Then
      strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
             "VALUES ('" & strNP01 & "'," & strNP22 & ",'法定期限','" & IIf(mNP09_1 > mNP09_2, mNP09_2, mNP09_1) & "')"
      iStep = iStep + 1
      strExc(1) = IIf(mNP09_1 = "", mNP09_2, IIf(mNP09_2 = "", mNP09_1, IIf(mNP09_1 > mNP09_2, mNP09_2, mNP09_1))) 'Added by Morgan 2016/8/10
   Else
      '--非合併催函
      strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
             "VALUES ('" & strNP01 & "'," & strNP22 & ",'法定期限','" & m_NP09 & "')"
      iStep = iStep + 1
      strExc(1) = m_NP09 'Added by Morgan 2016/8/10
   End If
   'end 2016/01/04
      
   'Added by Morgan 2016/8/10
   '若法定期限與專用期止日相差不足半年時定稿帶出將屆滿的句子
   If Val(m_PA25) > 0 And Val(strExc(1)) > 0 Then
      If m_PA25 < CompDate(1, 6, strExc(1)) Then
         strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
             "VALUES ('" & strNP01 & "'," & strNP22 & ",'即將屆滿','♀')"
         iStep = iStep + 1
      End If
   End If
   'end 2016/8/9
   
   'Added by Lydia 2016/01/04 合併催函的最早所限
   If bolMerge = True And min_NP08 <> "" Then
        strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
           "VALUES ('" & strNP01 & "'," & strNP22 & ",'本所期限','" & min_NP08 & "')"
   Else
        strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
           "VALUES ('" & strNP01 & "'," & strNP22 & ",'本所期限','" & m_NP08 & "')"
   End If
   iStep = iStep + 1
   
   If Me.Option2(0).Value Then
      'Added by Lydia 2016/01/04 合併催函的金額和點數
      If lblMerge(0).Visible = True Then
         'Added by Morgan 2020/11/16
         If m_PA09 = "239" Then
            strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05,LCV07) " & _
              "VALUES ('" & strNP01 & "'," & strNP22 & ",'費用','" & Val(txt2(5)) & "','Y','" & lbl1.Caption & "')"
            iStep = iStep + 1
            
           strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
              "VALUES ('" & strNP01 & "'," & strNP22 & ",'費用點數','" & Val(txt2(7)) & "','')"
           iStep = iStep + 1
           
            strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05,LCV07) " & _
              "VALUES ('" & strNP01 & "'," & strNP22 & ",'英國費用','" & Val(txt2(8)) & "','Y','延展費(英國)')"
            iStep = iStep + 1
            
            strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
               "VALUES ('" & strNP01 & "'," & strNP22 & ",'英國費用點數','" & Val(txt2(9)) & "','')"
            iStep = iStep + 1
         Else
         'end 2020/11/16
            If txt2(5) <> "" Then
               strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                  "VALUES ('" & strNP01 & "'," & strNP22 & ",'" & feeTit1 & "','" & Val(txt2(5)) & "','Y')"
               iStep = iStep + 1
            End If
            If txt2(8) <> "" Then
               strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                  "VALUES ('" & strNP01 & "'," & strNP22 & ",'" & feeTit2 & "','" & Val(txt2(8)) & "','Y')"
               iStep = iStep + 1
            End If
            strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
               "VALUES ('" & strNP01 & "'," & strNP22 & ",'費用合計','" & Val(txt2(5)) + Val(txt2(8)) & "')"
            iStep = iStep + 1
            If txt2(7) <> "" Then
               strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                  "VALUES ('" & strNP01 & "'," & strNP22 & ",'" & feeTit1 & "點數" & "','" & Val(txt2(7)) & "','')"
               iStep = iStep + 1
            End If
            If txt2(9) <> "" Then
               strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                  "VALUES ('" & strNP01 & "'," & strNP22 & ",'" & feeTit2 & "點數" & "','" & Val(txt2(9)) & "','')"
               iStep = iStep + 1
            End If
            strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
               "VALUES ('" & strNP01 & "'," & strNP22 & ",'點數合計','" & Val(txt2(7)) + Val(txt2(9)) & "')"
            iStep = iStep + 1
         End If
         
      Else
        '原程式
        If txt2(5) <> "" Then
           strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05,LCV07) " & _
              "VALUES ('" & strNP01 & "'," & strNP22 & ",'費用','" & Val(txt2(5)) & "','Y','" & lbl1.Caption & "')"
           iStep = iStep + 1
           'Added by Morgan 2020/3/17
           '超項費、超頁費
           If Val(txt2(6)) > 0 Or Val(txt2(10)) > 0 Then
               If Val(txt2(6)) > 0 Then
                  strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05,LCV08) " & _
                     "VALUES ('" & strNP01 & "'," & strNP22 & ",'超項費','" & Val(txt2(6)) & "','Y','N')"
                  iStep = iStep + 1
               End If
               If Val(txt2(10)) > 0 Then
                  strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05,LCV08) " & _
                     "VALUES ('" & strNP01 & "'," & strNP22 & ",'超頁費','" & Val(txt2(10)) & "','Y','N')"
                  iStep = iStep + 1
               End If
               strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
                  "VALUES ('" & strNP01 & "'," & strNP22 & ",'要印合計','♀')"
               iStep = iStep + 1
               
               strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
                  "VALUES ('" & strNP01 & "'," & strNP22 & ",'費用合計','" & (Val(txt2(5)) + Val(txt2(6)) + Val(txt2(10))) & "')"
               iStep = iStep + 1
            Else
           'end 2020/3/17
           
               'EPC 用
               strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
                  "VALUES ('" & strNP01 & "'," & strNP22 & ",'費用合計','" & Val(txt2(5)) & "')"
               iStep = iStep + 1
           End If 'Added by Morgan 2020/3/17
        End If
        
        If txt2(7) <> "" Then
           strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
              "VALUES ('" & strNP01 & "'," & strNP22 & ",'費用點數','" & Val(txt2(7)) & "','')"
           iStep = iStep + 1
           'EPC 用
           strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
              "VALUES ('" & strNP01 & "'," & strNP22 & ",'點數合計','" & Val(txt2(7)) & "','')"
           iStep = iStep + 1
        End If
      End If
   Else
      '2009/3/4 MODIFY BY SONIA 加EU239設計
      'If (strPA09 = "101" And strNP07 = "606") Or (strPA09 = "231" And strNP07 = "607") Then
      'modify by sonia 2024/7/16 加入英國201脫歐後衍生的的再註冊設計案
      If (strPA09 = "101" And strNP07 = "606") Or (strPA09 = "231" And strNP07 = "607") Or (strPA09 = "239" And strNP07 = "607") Or (strPA09 = "201" And strNP07 = "607" And InStr(m_PA91, "歐盟案案號：CFP") > 0) Then
         'Modify By Sindy 2009/07/29
         'Getnexttimes   '取得下次繳費次數
         m_Nexttimes = PUB_Getnexttimes(pa(1), pa(2), pa(3), pa(4), m_strYear, m_PA91)
         '2009/07/29 End
         
         '大個體用 Y00000002 抓
         'Modified by Morgan 2023/3/29
         'If InStr(1, m_PA91, "大個體", 1) > 0 Then
         m_PA179 = PUB_GetEntityType(pa(1), pa(2), pa(3), pa(4))
         If m_PA179 = "1" Then
         'end 2023/3/29
            'Modify By Sindy 2009/07/29 把m_Nexttimes改為m_strYear
            strExc(0) = "SELECT nvl(YF06,0),nvl(YF07,0) FROM PATENTYEARFEE WHERE YF01=" & CNULL(strPA09) & " AND YF02=" & CNULL(m_PA08) & " AND " & _
               "YF03='Y00000002' AND YF04=" & CNULL(strNP07) & " AND YF05=" & CNULL(m_strYear)
         'Added by Morgan 2013/3/20
         'Modified by Morgan 2023/3/29
         'ElseIf InStr(1, m_PA91, "微個體", 1) > 0 Then
         ElseIf m_PA179 = "3" Then
         'end 2023/3/29
            strExc(0) = "SELECT nvl(YF06,0),nvl(YF07,0) FROM PATENTYEARFEE WHERE YF01=" & CNULL(strPA09) & " AND YF02=" & CNULL(m_PA08) & " AND " & _
               "YF03='Y00000003' AND YF04=" & CNULL(strNP07) & " AND YF05=" & CNULL(m_strYear)
         'end 2013/3/20
         Else
            'Modify By Sindy 2009/07/29 把m_Nexttimes改為m_strYear
            strExc(0) = "SELECT nvl(YF06,0),nvl(YF07,0) FROM PATENTYEARFEE WHERE YF01=" & CNULL(strPA09) & " AND YF02=" & CNULL(m_PA08) & " AND " & _
               "YF03='Y00000000' AND YF04=" & CNULL(strNP07) & " AND YF05=" & CNULL(m_strYear)
         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
               "VALUES ('" & strNP01 & "'," & strNP22 & ",'費用','" & Val(RsTemp.Fields(0)) + Val(RsTemp.Fields(1)) & "','Y')"
            iStep = iStep + 1
            
            strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
               "VALUES ('" & strNP01 & "'," & strNP22 & ",'費用點數','" & Val(RsTemp.Fields(0)) / 1000 & "','')"
            iStep = iStep + 1
         End If
      End If
   End If
   'Modified by Lydia 2015/01/04 +EPC合併催函
   'If strNP07 = "215" Then
   If strNP07 = "215" Or (bolMerge = True And strPA09 = "221") Then
      If m_PA12 = "" Then
         strExc(0) = "申請日起2年內"
      Else
         strExc(0) = "公開日起半年內"
      End If
      strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'指定期限','" & strExc(0) & "')"
      iStep = iStep + 1
      If GetMemberCountryData(m_PA10, strExc) = True Then
         strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'生效日','" & strExc(1) & "')"
         iStep = iStep + 1

         strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'成員國數','" & strExc(2) & "')"
         iStep = iStep + 1
         
         strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'成員國','" & strExc(3) & "')"
         iStep = iStep + 1
         
         strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'延伸國數','" & strExc(4) & "')"
         iStep = iStep + 1
         
         strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'延伸國','" & strExc(5) & "')"
         iStep = iStep + 1
      End If
   End If
   If ET03 = "03" Then
      If Val(m_PA10) >= 20011001 Then
         strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'列印備註','3')"
      Else
         strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'列印備註','7')"
      End If
      iStep = iStep + 1
   End If
   
   'Added by Morgan 2012/4/11
   '馬來西亞新型延展費次數
   'Modified by Lydia 2016/01/04 合併催函
   'If strPA09 = "018" And m_PA08 = "2" And strNP07 = "607" Then
   '   strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
         "select '" & strNP01 & "'," & strNP22 & ",'大馬新型下次延展費次數','第'||decode(sign(trunc(np09/10000)- trunc(pa10/10000)-15),-1,1,2)||'次'" & _
         " from nextprogress,patent where np01='" & strNP01 & "' and np22=" & strNP22 & " and pa01(+)=np02 and pa02(+)=np03 and pa03(+)=np04 and pa04(+)=np05"
   If strPA09 = "018" And m_PA08 = "2" And (strNP07 = "607" Or bolMerge = True) Then
      strExc(0) = PUB_GetExpYF607(strPA09, pa(1), pa(2), pa(3), pa(4), m_PA10)
      strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
             "VALUES ('" & strNP01 & "'," & strNP22 & ",'大馬新型下次延展費次數','" & strExc(0) & "') "
      iStep = iStep + 1
   End If
   'Added by Lydia 2016/01/04 俄羅斯設計延展費有分新舊制
   'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
   If bolMerge = True And strPA09 = "023" Then
       strExc(0) = PUB_GetExpYF607(strPA09, pa(1), pa(2), pa(3), pa(4), m_PA10)
         strTxt(iStep) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
             "VALUES ('" & strNP01 & "'," & strNP22 & ",'俄羅斯設計下次延展費次數','" & strExc(0) & "') "
         iStep = iStep + 1
   End If
   'end 2016/01/04
   
   If Not ClsLawExecSQL(iStep - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If

End Sub

Private Sub lblFee1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseDown lblFee1, lblFee1s
End Sub

Private Sub lblFee1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseUp lblFee1, lblFee1s
End Sub

Private Function CheckCPExists(ByRef p_CP1 As String, ByRef p_CP2 As String, ByRef p_CP3 As String, ByRef p_CP4 As String) As Boolean
   'Modified by Lydia 2016/08/25 CP27->CP158
   ' strSql = "select * from caseprogress A where cp01='" & p_CP1 & "' and cp02='" & p_CP2 & "' and cp03='" & p_CP3 & "' and cp04='" & p_CP4 & "' and cp10 in ('1604','1606','1907','413','429') and cp27 is not null"
   strSql = "select * from caseprogress A where cp01='" & p_CP1 & "' and cp02='" & p_CP2 & "' and cp03='" & p_CP3 & "' and cp04='" & p_CP4 & "' and cp10 in ('1604','1606','1907','413','429') and cp158 > 0 "
   'Add by Lydia 2014/11/24 加判斷未收文恢復權利414，恢復權利改發文日也要判斷(原來只判斷收文日但有發生例外 Ex.P-78503) from frm050312
   'Modified by Lydia 2016/08/25 CP57->CP159
   'strSql = strSql & " and not exists(select * from caseprogress B where B.cp01=A.cp01 and B.cp02=A.cp02 and B.cp03=A.cp03 and B.cp04=A.cp04 and ((B.cp05>A.cp05 and B.cp27 is null) or B.cp27>A.cp27) and B.cp10='414' and B.cp57 is null)"
   strSql = strSql & " and not exists(select * from caseprogress B where B.cp01=A.cp01 and B.cp02=A.cp02 and B.cp03=A.cp03 and B.cp04=A.cp04 and ((B.cp05>A.cp05 and B.cp158 = 0) or B.cp158>A.cp158) and B.cp10='414' and B.cp159 = 0 )"
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         CheckCPExists = True
      End If
   End With
End Function
'Added by Lydia 2016/01/04 檢查是否有兩個需要合併的催函(只判斷個案)
Private Sub ChkLetterMerge()
Dim strAA As String
   bolMerge = False
   min_NP08 = "": feeTit1 = "": feeTit2 = ""
   mNP09_1 = "": mNP09_2 = ""
   mNP01_2 = "": mNP08_2 = "": mNP22_2 = "" 'Added by Morgan 2022/6/8
   mNP07_2(0) = "": mNP07_2(1) = ""
   lblMerge(0).Visible = False: lblMerge(1).Visible = False
   txt2(8).Visible = False: txt2(9).Visible = False
   txt2(8).Tag = "" 'Added by Lydia 2016/08/31
   lblMerge(0) = "另一金額：" 'Added by Morgan 2020/11/16
   
   If Option1(1).Value = True And Option2(0).Value = True And txt2(0).Text = "CFP" Then
   
   
      If txt2(2).Text = "" Then txt2(2).Text = "0"
      If txt2(3).Text = "" Then txt2(3).Text = "00"
      'Modified by Morgan 2022/6/13 +pa10
      strSql = "select pa09,pa08,pa10 from patent where pa01='" & txt2(0) & "' and pa02='" & txt2(1) & "' and pa03='" & txt2(2) & "' and pa04='" & txt2(3) & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         'Added by Morgan 2020/11/16 歐盟延展費檢查是否有613延展費(英國)要輸金額及點數
         m_PA09 = RsTemp.Fields("pa09")
         If m_PA09 = "239" And txt2(4) = "607" Then
            'modify by sonia 2025/10/7 過期期限不抓CFP-028295
            strExc(0) = "select np07,np08,np09,np01,np22 from nextprogress where np06 is null and np07='613'" & _
                      " and np02='" & txt2(0) & "' and np03='" & txt2(1) & "' and np04='" & txt2(2) & "' and np05='" & txt2(3) & "' and np09>=" & strSrvDate(1)
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               lblMerge(0).Visible = True: lblMerge(1).Visible = True
               txt2(8).Visible = True: txt2(9).Visible = True
               lblMerge(0) = "英國金額："
               
               'Added by Morgan 2022/6/8
               mNP01_2 = "" & RsTemp.Fields("np01")
               mNP08_2 = "" & RsTemp.Fields("np08")
               mNP09_2 = "" & RsTemp.Fields("np09")
               mNP22_2 = "" & RsTemp.Fields("np22")
               'end 2022/6/8
               
               'add by sonia 2024/9/11  抓英國201的設定費用
               strExc(0) = "SELECT nvl(YF06,0),nvl(YF07,0),PA16 FROM patent,PATENTYEARFEE" & _
                  " WHERE " & ChgPatent(txt2(0) & txt2(1) & txt2(2) & txt2(3)) & _
                  " and YF01(+)='201' AND YF02(+)=pa08 AND " & _
                  " YF03='Y00000000' AND YF04=" & txt2(4) & " AND YF05=" & CNULL(m_strYear)
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  txt2(8) = RsTemp.Fields(0) + RsTemp.Fields(1)
                  txt2(9) = RsTemp.Fields(0) / 1000
                  If txt2(8) = "0" Then txt2(8) = ""
                  If txt2(9) = "0" Then txt2(9) = ""
               End If
               'end 2024/9/11
            End If
         Else
         'end 2020/11/16
            strExc(0) = ""
            Select Case RsTemp.Fields("pa09")
               Case "221" 'EPC案件催實審及指定費
                    strAA = "416,215"
               Case "018" '馬來西亞新型年費及延展費
                    If RsTemp.Fields("pa08") = "2" Then
                       strAA = "605,607"
                    End If
               'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
               Case "023" '俄羅斯設計的年費及延展費
                    'Modified by Morgan 2022/6/13
                    '2015/1/1以前提申案件除了延展費外仍要繳年費，2015/1/1以後提申案件僅須繳延展費
                    'If RsTemp.Fields("pa08") = "3" Then
                    If RsTemp.Fields("pa08") = "3" And RsTemp.Fields("pa10") < 20150101 Then
                    'end 2022/6/13
                       strAA = "605,607"
                    End If
            End Select
            If strAA <> "" Then
               If InStr(strAA, txt2(4)) > 0 Then
                  'Modified by Morgan 2021/10/28 +np01,np22
                  strExc(0) = "select np07,np08,np09,np01,np22 from nextprogress where np06 is null and np07 in (" & strAA & ") " & _
                      " and np02='" & txt2(0) & "' and np03='" & txt2(1) & "' and np04='" & txt2(2) & "' and np05='" & txt2(3) & "' "
               End If
            End If
            If strExc(0) <> "" Then
               feeTit1 = lbl1.Caption
               If InStrRev(feeTit1, "（") > 0 Then
                  feeTit1 = Left(feeTit1, InStrRev(feeTit1, "（") - 1)
               End If
               mNP07_2(0) = Replace(Replace(strAA, txt2(4), ""), ",", "")
               feeTit2 = GetPrjState6HM("CFP", mNP07_2(0))
               mNP07_2(1) = feeTit2
               If InStrRev(feeTit2, "（") > 0 Then
                  feeTit2 = Left(feeTit2, InStrRev(feeTit2, "（") - 1)
               End If
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If RsTemp.RecordCount > 1 Then
                  If MsgBox("催函是否要與" & feeTit2 & "合併為一個?", vbYesNo + vbInformation, "合併催函") = vbYes Then
                     bolMerge = True
                     '合併催函->另一金額
                     lblMerge(0).Visible = True: lblMerge(1).Visible = True
                     txt2(8).Visible = True: txt2(9).Visible = True
                     RsTemp.MoveFirst
                     min_NP08 = RsTemp.Fields("np08")
                     Do While Not RsTemp.EOF
                        If min_NP08 > RsTemp.Fields("np08") Then min_NP08 = RsTemp.Fields("np08")
                        If RsTemp.Fields("np07") = txt2(4) Then
                           mNP09_1 = "" & RsTemp.Fields("np09")
                        Else
                           mNP09_2 = "" & RsTemp.Fields("np09")
                           txt2(8).Tag = "" & RsTemp.Fields("np07")  'Added by Lydia 2016/08/31
                           'Added by Morgan 2021/10/29
                           mNP01_2 = "" & RsTemp.Fields("np01")
                           mNP08_2 = "" & RsTemp.Fields("np08")
                           mNP22_2 = "" & RsTemp.Fields("np22")
                           'end 2021/10/29
                        End If
                        RsTemp.MoveNext
                     Loop
                     
                  End If
               End If
            End If
         End If 'Added by Morgan 2020/11/16
         
      End If
   End If
End Sub

'Added by Lydia 2016/08/15 換行判斷
Private Sub PrintNewLine(Optional ByVal mRate As Single = 1, Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 4)
   iPrint = iPrint + lngLineHeight * mRate
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint

      Page = Page + 1
      Printer.NewPage
      PrintTitle
   End If
End Sub
'Added by Lydia 2016/08/25 設定印表機
Private Sub SettingPrtSet()
Dim inX As Integer
Dim tmpArr As Variant, tmpArr2 As Variant

    '設定印表機
     Printer.EndDoc
     Printer.PaperSize = 9  'A4
     Printer.Orientation = 2 '2.橫印

     lngPageHeight = Printer.ScaleHeight
     lngPageWidth = Printer.ScaleWidth
     lngLineHeight = 300
     Printer.Font.Name = "新細明體"
     Printer.Font.Size = ciFontSize
     Erase PLeft
     Erase PTitle
     tmpArr = Empty: tmpArr2 = Empty
     
     '設定欄位抬頭和位置
     If strTitle <> "" And strTitle2 <> "" Then
        tmpArr = Split(strTitle, ",")
        tmpArr2 = Split(strTitle2, ",")
        For inX = 0 To UBound(tmpArr)
            If Trim(tmpArr(inX)) <> "" And Trim(tmpArr2(inX)) <> "" Then
                If Trim(tmpArr(inX)) <> "結束" Then PTitle(inX) = Trim(tmpArr(inX))
                If inX < 1 Then
                   PLeft(inX) = ciStartX
                Else
                   PLeft(inX) = PLeft(inX - 1) + Printer.TextWidth(String(Val(tmpArr2(inX)), "　")) + ciColGap
                End If
                
                If Trim(tmpArr(inX)) = "結束" Then Exit For
            End If
        Next
     End If
     
     Page = 0
End Sub

'Added by Lydia 2016/09/13 檢查"個案之個體狀態"與"客戶減免身份"個體是不同
Private Sub ChkEntity()
Dim StrStr1 As String
Dim rsA As New ADODB.Recordset
Dim arrCF209() As String 'Added by Morgan 2023/3/25
Dim strCF203 As String 'Added by Morgan 2023/3/29

   If Option1(1).Value = True And Option2(0).Value = True And strChkEntity <> txt2(0) & txt2(1) & txt2(2) & txt2(3) And txt2(0).Text <> "" And txt2(1).Text <> "" Then
      'Modified by Morgan 2023/3/25 +pa179
      StrStr1 = "select sk02,pa08,pa09,pa26,pa91,pa179,ad03 from patent,applicantdiscount,systemkind " & _
                "where pa01='" & txt2(0) & "' and pa02='" & txt2(1) & "' and pa03='" & Left(txt2(2) & "0", 1) & "' and pa04='" & Left(txt2(3) & "00", 2) & "' and substr(pa26,1,8)=ad01(+) and pa09=ad02(+) and pa01=sk01(+) "
      intI = 1
      Set rsA = ClsLawReadRstMsg(intI, StrStr1)
      If intI = 1 Then
         m_PA09 = "" & rsA.Fields("pa09")  'add by sonia 2024/9/11
         m_PA08 = "" & rsA.Fields("pa08")  'Added by Morgan 2025/3/13
         If "" & rsA.Fields("sk02") = "1" And InStr(CFP_ChkEntity, "" & rsA.Fields("pa09")) > 0 And "" & rsA.Fields("ad03") <> "" Then
            StrStr1 = ""
            'Modified by Morgan 2023/3/25
            'If InStr(1, "" & rsA.Fields("pa91"), "大個體", 1) > 0 And rsA.Fields("ad03") = "Y" Then
            '   StrStr1 = "大個體"
            'ElseIf InStr(1, "" & rsA.Fields("pa91"), "小個體", 1) > 0 And rsA.Fields("ad03") = "N" Then
            '   StrStr1 = "小個體"
            'ElseIf InStr(1, "" & rsA.Fields("pa91"), "微個體", 1) > 0 And rsA.Fields("ad03") = "N" Then
            '   StrStr1 = "微個體"
            'End If
            ReDim arrCF209(2) As String
            arrCF209(0) = "大個體"
            arrCF209(1) = "小個體"
            arrCF209(2) = "微個體"
            
            PUB_SetEntityOpt "CFP", rsA.Fields("pa09"), rsA.Fields("pa08"), arrCF209
            If strSrvDate(1) >= PA179啟用日 Then
               If rsA.Fields("pa179") = "1" And rsA.Fields("ad03") = "Y" Then
                  StrStr1 = arrCF209(0)
               ElseIf rsA.Fields("pa179") = "2" And rsA.Fields("ad03") = "N" Then
                  StrStr1 = arrCF209(1)
               ElseIf rsA.Fields("pa179") = "3" And rsA.Fields("ad03") = "N" Then
                  StrStr1 = arrCF209(2)
               End If
            Else
               If InStr(1, rsA.Fields("pa91"), "大個體", 1) > 0 And rsA.Fields("ad03") = "Y" Then
                  StrStr1 = arrCF209(0)
               ElseIf InStr(1, rsA.Fields("pa91"), "小個體", 1) > 0 And rsA.Fields("ad03") = "N" Then
                  StrStr1 = arrCF209(1)
               ElseIf InStr(1, rsA.Fields("pa91"), "微個體", 1) > 0 And rsA.Fields("ad03") = "N" Then
                  StrStr1 = arrCF209(2)
               End If
            End If
            'end 2023/3/25
            
            If StrStr1 <> "" Then
               MsgBox "本案客戶減免設定為【" & StrStr1 & "】與基本檔不同！", vbCritical, "客戶減免檢查"
            End If
         End If
      End If
      Set rsA = Nothing
   End If
   
   strChkEntity = txt2(0) & txt2(1) & txt2(2) & txt2(3)
End Sub

'Added by Lydia 2016/12/08 新增資料到暫存檔
'Modify by Amy 2018/05/03 +stNP09
Private Sub ProcessInsRec(ByRef sTemp() As String, Optional stNP09 As String)
    sTemp(0) = StrConv(MidB(StrConv(sTemp(0), vbFromUnicode), 1, 8), vbUnicode)
    sTemp(1) = StrConv(MidB(StrConv(sTemp(1), vbFromUnicode), 1, 8), vbUnicode)
    sTemp(2) = StrConv(MidB(StrConv(sTemp(2), vbFromUnicode), 1, 10), vbUnicode)
    sTemp(3) = StrConv(MidB(StrConv(sTemp(3), vbFromUnicode), 1, 10), vbUnicode)
    'Add by Amy 2018/05/03 法限(子案無期限掛母案 for 子案母案排一起)
    If sTemp(3) = MsgText(601) Then
        sTemp(3) = stNP09 & "*"
    End If
    sTemp(4) = StrConv(MidB(StrConv(sTemp(4), vbFromUnicode), 1, 15), vbUnicode)
    sTemp(6) = StrConv(MidB(StrConv(sTemp(6), vbFromUnicode), 1, 8), vbUnicode)
    sTemp(7) = StrConv(MidB(StrConv(StrConv(sTemp(7), vbWide), vbFromUnicode), 1, 8), vbUnicode)
    sTemp(8) = StrConv(MidB(StrConv(StrConv(sTemp(8), vbWide), vbFromUnicode), 1, 8), vbUnicode)
    sTemp(9) = StrConv(MidB(StrConv(sTemp(9), vbFromUnicode), 1, 10), vbUnicode)
    sTemp(10) = StrConv(MidB(StrConv(sTemp(10), vbFromUnicode), 1, 15), vbUnicode)
    sTemp(11) = StrConv(MidB(StrConv(sTemp(11), vbFromUnicode), 1, 10), vbUnicode)
    sTemp(12) = StrConv(MidB(StrConv(sTemp(12), vbFromUnicode), 1, 20), vbUnicode)
    sTemp(13) = StrConv(MidB(StrConv(sTemp(13), vbFromUnicode), 1, 50), vbUnicode)
    sTemp(5) = StrConv(MidB(StrConv(sTemp(5), vbFromUnicode), 1, 120), vbUnicode) 'Added by Lydia 2024/04/22 配合Table欄位大小
    
    strSql = "INSERT INTO R050312 VALUES ('" & ChgSQL(sTemp(0)) & "','" & ChgSQL(sTemp(1)) & "','" & ChgSQL(sTemp(2)) & "','" & ChgSQL(sTemp(3)) & "','" & ChgSQL(sTemp(4)) & "','" & ChgSQL(sTemp(5)) & "','" & ChgSQL(sTemp(6)) & "','" & ChgSQL(sTemp(7)) & "','" & ChgSQL(sTemp(8)) & "','" & ChgSQL(sTemp(9)) & "','" & ChgSQL(sTemp(10)) & "','" & ChgSQL(sTemp(11)) & "','" & ChgSQL(sTemp(12)) & "','" & ChgSQL(sTemp(13)) & "','" & ChgSQL(sTemp(14)) & "','" & strUserNum & "') "
    cnnConnection.Execute strSql
    k = k + 1
End Sub

'Added by Morgan 2017/12/19
Private Sub UpdateDate(pCP09 As String)
   Dim iRec As Integer
On Error GoTo ErrHnd
   cnnConnection.Execute "UPDATE CASEPROGRESS SET CP27=" & strSrvDate(1) & " WHERE CP09='" & pCP09 & "' AND CP10='1913' AND CP27 IS NULL", iRec
   If strSrvDate(1) >= CFP第一階段電子化啟用日 Or UCase(pub_DbTerminalName) <> 正式資料庫電腦名稱 Then
      PUB_UpdateLP03 pCP09
   End If
   Exit Sub
ErrHnd:
   MsgBox Err.Description, vbCritical
End Sub

'Add by Amy 2018/04/30 進度檔是否已有通知期限1913
Private Function bolHas1913(stCP43, stCP30) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    bolHas1913 = False
    strQ = "Select * From CaseProgress Where cp43='" & stCP43 & "' And cp10='1913' And cp30='" & stCP30 & "' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        bolHas1913 = True
    End If
    RsQ.Close
End Function
'Added by Morgan 2020/3/17
'檢查是否可輸入超項費超頁費
Private Sub ChkOverFee()
   Dim pa(4) As String
   '日本、韓國、俄羅斯、歐亞專利聯盟、阿根廷、巴西及英國(超頁)等國家請跳訊息提醒程序人員是否有超項及超頁費應輸入
   If txt2(4) = "416" And InStr("011,012,023,074,118,117,201", m_PA09) > 0 Then
      pa(1) = txt2(0): pa(2) = txt2(1)
      pa(3) = txt2(2): pa(4) = txt2(3)
      If m_PA09 = "201" Then
         lblOverFee(1).Visible = True
         txt2(10).Visible = True
         If PUB_ChkCPExist(pa(), "938") Then
            MsgBox "本案已收文超頁費，請留意！", vbExclamation
         End If
      Else
         lblOverFee(0).Visible = True
         txt2(6).Visible = True
         If PUB_ChkCPExist(pa(), "939") Then
            MsgBox "本案已收文超項費，請留意！", vbExclamation
         End If
      End If
   Else
      lblOverFee(0).Visible = False
      txt2(6).Visible = False
      txt2(6) = ""
      lblOverFee(1).Visible = False
      txt2(10).Visible = False
      txt2(10) = ""
   End If
End Sub

'Added by Morgan 2025/3/13
'檢查案件性質是否要提醒
Private Sub ChkAlert()
   Dim pa(4) As String
   Dim strDesc As String
   
   If Option2(0).Value Then
      '2024/10/5起俄羅斯發明及新型年費改一次繳5年
      If m_PA09 = "023" And (m_PA08 = "1" Or m_PA08 = "2") And txt2(4) = "605" Then
         pa(1) = txt2(0)
         pa(2) = txt2(1)
         pa(3) = Right("0" & txt2(2), 1)
         pa(4) = Right("00" & txt2(3), 2)
         strExc(0) = PUB_GetNextYear(pa(), strDesc)
         strExc(0) = PUB_GetNextYear(pa(), strDesc)
         If strDesc <> "" Then
            lbl1 = lbl1 & "(" & strDesc & ")"
         End If
         
         If lbl1.Tag <> lbl1 Then
            MsgBox "俄羅斯已修法改為一次繳5年，本次應繳「" & strDesc & "」！", vbExclamation
         End If
      End If
   End If
   lbl1.Tag = lbl1
End Sub
