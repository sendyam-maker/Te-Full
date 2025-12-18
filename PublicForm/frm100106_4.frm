VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100106_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "發E-Mail對象"
   ClientHeight    =   5280
   ClientLeft      =   2136
   ClientTop       =   2892
   ClientWidth     =   5124
   ControlBox      =   0   'False
   LinkTopic       =   "Form18"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5124
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  '沒有框線
      Height          =   4365
      Left            =   3510
      TabIndex        =   21
      Top             =   600
      Width           =   4935
      Begin VB.TextBox TxtAddCC 
         Height          =   270
         Left            =   1080
         TabIndex        =   40
         Top             =   2640
         Width           =   3165
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "收件人"
         Height          =   2115
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   4755
         Begin VB.TextBox TxtAddTo 
            Height          =   300
            Left            =   1380
            TabIndex        =   39
            Top             =   1747
            Width           =   1125
         End
         Begin VB.CheckBox ChkMail 
            Caption         =   "其他收件人"
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   28
            Top             =   1785
            Width           =   1275
         End
         Begin VB.CheckBox ChkMail 
            Caption         =   "核稿人"
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   27
            Top             =   1416
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.CheckBox ChkMail 
            Caption         =   "承辦人"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   26
            Top             =   1049
            Width           =   1005
         End
         Begin VB.CheckBox ChkMail 
            Caption         =   "智權人員"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   25
            Top             =   682
            Width           =   1065
         End
         Begin VB.CheckBox ChkMail 
            Caption         =   "管制人"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   24
            Top             =   315
            Width           =   1005
         End
         Begin MSForms.Label lbl2 
            Height          =   285
            Index           =   4
            Left            =   2520
            TabIndex        =   33
            Top             =   1770
            Width           =   2145
            VariousPropertyBits=   27
            Size            =   "3784;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lbl2 
            Height          =   285
            Index           =   3
            Left            =   1380
            TabIndex        =   32
            Top             =   1416
            Visible         =   0   'False
            Width           =   3285
            VariousPropertyBits=   27
            Size            =   "5794;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lbl2 
            Height          =   285
            Index           =   2
            Left            =   1380
            TabIndex        =   31
            Top             =   1049
            Width           =   3285
            VariousPropertyBits=   27
            Size            =   "5794;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lbl2 
            Height          =   285
            Index           =   1
            Left            =   1380
            TabIndex        =   30
            Top             =   682
            Width           =   3285
            VariousPropertyBits=   27
            Size            =   "5794;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lbl2 
            Height          =   285
            Index           =   0
            Left            =   1380
            TabIndex        =   29
            Top             =   315
            Width           =   3285
            VariousPropertyBits=   27
            Size            =   "5794;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin MSForms.TextBox TxtAddCCName 
         Height          =   615
         Left            =   1080
         TabIndex        =   41
         Top             =   2940
         Width           =   3705
         VariousPropertyBits=   -1476378597
         BackColor       =   -2147483644
         ScrollBars      =   2
         Size            =   "6535;1085"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt2 
         Height          =   675
         Index           =   1
         Left            =   630
         TabIndex        =   37
         Top             =   3630
         Visible         =   0   'False
         Width           =   4575
         VariousPropertyBits=   -1476378597
         ScrollBars      =   2
         Size            =   "8070;1191"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt2 
         Height          =   675
         Index           =   0
         Left            =   300
         TabIndex        =   36
         Top             =   3600
         Width           =   4575
         VariousPropertyBits=   -1476378597
         ScrollBars      =   2
         Size            =   "8064;1191"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註："
         Height          =   180
         Index           =   3
         Left            =   270
         TabIndex        =   35
         Top             =   3180
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "副本："
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   34
         Top             =   2670
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "完稿日："
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   180
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Height          =   4125
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   4935
      Begin VB.OptionButton Option1 
         Caption         =   "管制人"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   135
         Width           =   972
      End
      Begin VB.OptionButton Option1 
         Caption         =   "智權人員"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   8
         Top             =   495
         Width           =   1035
      End
      Begin VB.OptionButton Option1 
         Caption         =   "承辦人"
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   7
         Top             =   855
         Width           =   852
      End
      Begin VB.OptionButton Option1 
         Caption         =   "核稿人"
         Height          =   180
         Index           =   3
         Left            =   90
         TabIndex        =   6
         Top             =   1215
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.OptionButton Option1 
         Caption         =   "其他收件人"
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   5
         Top             =   1575
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   4
         Top             =   1530
         Width           =   735
      End
      Begin MSForms.ListBox lstMailCC 
         Height          =   915
         Left            =   960
         TabIndex        =   15
         Top             =   1950
         Width           =   1815
         VariousPropertyBits=   746586139
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "3201;1614"
         MatchEntry      =   0
         ListStyle       =   1
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   795
         Index           =   1
         Left            =   360
         TabIndex        =   20
         Top             =   3210
         Visible         =   0   'False
         Width           =   4575
         VariousPropertyBits=   -1476378597
         ScrollBars      =   2
         Size            =   "8070;1402"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   795
         Index           =   0
         Left            =   210
         TabIndex        =   18
         Top             =   3120
         Width           =   4572
         VariousPropertyBits=   -1476378597
         ScrollBars      =   2
         Size            =   "8064;1402"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         Caption         =   "txt1(0):可見可輸入, txt1(1):不可見,案件資料"
         Height          =   165
         Left            =   990
         TabIndex        =   38
         Top             =   2940
         Visible         =   0   'False
         Width           =   3405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註："
         Height          =   180
         Index           =   0
         Left            =   210
         TabIndex        =   19
         Top             =   2880
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "(可複選)"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   17
         Top             =   2250
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "副本："
         Height          =   210
         Index           =   10
         Left            =   120
         TabIndex        =   16
         Top             =   2010
         Width           =   795
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   0
         Left            =   1335
         TabIndex        =   14
         Top             =   135
         Width           =   3285
         VariousPropertyBits=   27
         Size            =   "5794;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   2
         Left            =   1335
         TabIndex        =   13
         Top             =   855
         Width           =   3285
         VariousPropertyBits=   27
         Size            =   "5794;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   1
         Left            =   1335
         TabIndex        =   12
         Top             =   495
         Width           =   3285
         VariousPropertyBits=   27
         Size            =   "5794;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   3
         Left            =   1335
         TabIndex        =   11
         Top             =   1215
         Visible         =   0   'False
         Width           =   3285
         VariousPropertyBits=   27
         Size            =   "5794;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   285
         Index           =   4
         Left            =   2145
         TabIndex        =   10
         Top             =   1530
         Width           =   2475
         VariousPropertyBits=   27
         Size            =   "4366;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CheckBox chkByOutLook 
      Caption         =   "密件副本：操作者本人"
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   5010
      Width           =   2580
   End
   Begin VB.CommandButton cmdGoInput 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3852
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3024
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   45
      Width           =   800
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   2280
      Top             =   -120
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   1620
      Top             =   -180
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
End
Attribute VB_Name = "frm100106_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/05/18 Form2.0已修改: lbl1(index)、txt1(index)、lbl2(index)、txt2(index)、TxtAddCCName
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/13 日期欄已修改
Option Explicit

'Added by Lydia 2020/03/11
Dim m_Role As String '依人員區分顯示Frame: A-其他, F-外專
Dim mPrevForm As Form  '前一畫面
Dim m_CP07 As String '點選進度之法定期限
'end 2020/03/11
Dim i As Integer, j As Integer, strSql As String, s As Integer
Dim Str01 As String, Str02 As String, Str03 As String, strTemp As String
'910430  因為 mail 解析，中文會有問題，所以用員工編號
Public StrMailNum1 As String        '管制人
Public StrMailNum2 As String        '智權人員
Public StrMailNum3 As String        '承辦人
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Add by Morgan 2007/1/8
Public strCP09 As String            '收文號
Public strCaseNo As String 'Added by Lydia 2020/05/18 本所案號
Public StrMailNum4 As String        '核稿人
Dim strFCP201State As String 'FCP翻譯控制狀態 0:無 1:外譯且已完稿 2:外譯且未完稿
Public strFRname As String   'Added by Lydia 2017/01/19 呼叫表單名稱
Public strTypeMemo As String 'Added by Lydia 2021/04/29 預設備註(可見可輸入)
Dim oObj 'Added by Lydia 2021/05/18

'Added by Lydia 2020/03/11
Public Sub SetParent(ByRef pForm As Form, ByVal pCP07 As String)
   Set mPrevForm = pForm
   If Left(Pub_StrUserSt03, 2) = "F2" Then
      m_Role = "F"
   Else
      m_Role = "A"
   End If
   m_CP07 = pCP07
End Sub

'Modify by Morgan 2007/1/9
Public Sub PubShowNextData()
   Dim strTo As String '收件者員工編號
   Dim strToName As String
   Dim bolByOutLook As Boolean
   Dim stContent As String, stSubject As String
   Dim strToCC As String
   Dim strF51CC As String         'add by sonia 2016/7/5
   Dim strF51CCNAME As String     'add by sonia 2016/7/5
   Dim tmpArr As Variant 'Added by Lydia 2020/03/11

   Select Case cmdState
      Case 0
         'Added by Lydia 2020/03/11 依人員區分顯示Frame: A-其他, F-外專
         strExc(0) = ""
         If m_Role = "F" Then '外專人員
              txt2(1).Text = txt1(1).Text '複製案件資料
              If ChkMail(0).Value = 1 Then  '管制人
                  strTo = strTo & ";" & StrMailNum1
                  'Added by Lydia 2023/12/22
                  If InStr(lbl1(0), "離職") > 0 Then
                     strExc(0) = "原管制人已離職，請主管重新分案。"
                  End If
                  'end 2023/12/22
              End If
              If ChkMail(1).Value = 1 Then  '智權人員
                  strTo = strTo & ";" & StrMailNum2
                  'Added by Lydia 2023/12/22
                  If InStr(lbl1(1), "離職") > 0 Then
                     strExc(0) = "原智權人員已離職，請主管重新分案。"
                  End If
                  'end 2023/12/22
              End If
              If ChkMail(2).Value = 1 Then  '承辦人
                   Select Case strFCP201State '判斷是否為新案翻譯
                      Case "1"
                         MsgBox "本案承辦為外譯人員且已完稿！"
                         ChkMail(2).Value = False
                         Exit Sub
                      Case "2"
                         strF51CC = Pub_GetSpecMan("M")
                         strF51CCNAME = GetStaffName(strF51CC)
                         If MsgBox("本案承辦為外譯人員是否要改通知 " & strF51CCNAME & " ？", vbYesNo + vbDefaultButton1) = vbYes Then
                            strTo = strTo = strTo & ";" & strF51CC
                         Else
                            Exit Sub
                         End If
                      Case Else
                         strTo = strTo & ";" & StrMailNum3
                   End Select
                  'Added by Lydia 2023/12/22
                  If InStr(lbl1(2), "離職") > 0 Then
                     strExc(0) = "原承辦人已離職，請主管重新分案。"
                  End If
                  'end 2023/12/22
              End If
              If ChkMail(3).Value = 1 Then  '核稿人
                  strTo = strTo & ";" & StrMailNum4
              End If
              If ChkMail(4).Value = 1 Then  '其他收件人
                  If TxtAddTo.Text <> "" Then
                      If lbl2(4).Caption = "" Then
                          MsgBox "請輸入正確的員工編號！", vbCritical, "檢核資料"
                          TxtAddTo.SetFocus
                          TxtAddTo_GotFocus
                          Exit Sub
                      Else
                          strTo = strTo & ";" & TxtAddTo.Text
                      End If
                  End If
              End If
              strTo = Replace(strTo, ";;", ";")
              strToName = PUB_ReadUserData(strTo)
              If TxtAddCC.Text <> "" Then
                  If TxtAddCCName.Text = "" Then
                      MsgBox "請輸入正確的員工編號！", vbCritical, "檢核資料"
                      TxtAddCC.SetFocus
                      TxtAddCC_GotFocus
                      Exit Sub
                  Else
                      strToCC = TxtAddCC.Text
                  End If
              End If
         Else    '其他部門-A
         'end 2020/03/11
                If Option1(0).Value = True Then
                   strTo = StrMailNum1
                   strToName = lbl1(0).Caption
                ElseIf Option1(1).Value = True Then
                   strTo = StrMailNum2
                   strToName = lbl1(1).Caption
                ElseIf Option1(2).Value = True Then
                   Select Case strFCP201State
                      Case "1"
                         MsgBox "本案承辦為外譯人員且已完稿！"
                         Exit Sub
                      Case "2"
                         'MODIFY BY SONIA 2016/7/5 改用系統特殊設定人員
                         'If MsgBox("本案承辦為外譯人員是否要改通知靜芳？", vbYesNo + vbDefaultButton1) = vbYes Then
                         '   strTo = "73023"
                         strF51CC = Pub_GetSpecMan("M")
                         strF51CCNAME = GetStaffName(strF51CC)
                         If MsgBox("本案承辦為外譯人員是否要改通知 " & strF51CCNAME & " ？", vbYesNo + vbDefaultButton1) = vbYes Then
                            strTo = strF51CC
                         'end 2016/7/5
                         Else
                            Exit Sub
                         End If
                      Case Else
                         strTo = StrMailNum3
                   End Select
                   strToName = lbl1(2).Caption
                'Add by Morgan 2010/6/7
                ElseIf Option1(4).Value = True Then
                   If lbl1(4) = "" Then
                      MsgBox "請輸入正確的收件人！"
                      If Text1.Enabled = True Then Text1.SetFocus
                      Exit Sub
                   Else
                      strTo = Text1
                      strToName = lbl1(4).Caption 'Add By Sindy 2012/3/29
                   End If
                Else
                   strTo = StrMailNum4
                   strToName = lbl1(3).Caption
                End If
                'Add By Sindy 2012/3/29
                '副本
                'Move by Lydia 2020/03/11 從”收件人空白”下方移上來
                For i = 0 To lstMailCC.ListCount - 1
                   If lstMailCC.Selected(i) = True Then
                      If strToCC = "" Then
                         strToCC = Left(Trim(lstMailCC.List(i)), 5)
                      Else
                         strToCC = strToCC & ";" & Left(Trim(lstMailCC.List(i)), 5)
                      End If
                   End If
                Next
                'end 2020/03/11
         End If 'Added by Lydia 2020/03/11
         
         If strTo = "" Then
            MsgBox "收件人空白，無法寄送！"
            Exit Sub
         End If
                  
         Screen.MousePointer = vbHourglass
         
         'Modified by Lydia 2020/05/18 +本所案號
         'stSubject = "通知期限"
         stSubject = "通知期限" & IIf(strCaseNo <> "", "：" & strCaseNo, "")
         'Added by Lydia 2020/03/11 依人員區分顯示Frame: A-其他, F-外專
         If m_Role = "F" Then '外專人員
             'Modified by Lydia 2023/12/22 離職人員改發主管加註+ strexc(0)+ IIf(strExc(0) <> "", strExc(0) + vbCrLf + vbCrLf, "")
             stContent = "TO ：收受人姓名：" + strToName + vbCrLf + vbCrLf + txt2(1) + vbCrLf + vbCrLf + Space(11) + txt2(0) + vbCrLf + vbCrLf + IIf(strExc(0) <> "", Space(11) + strExc(0) + vbCrLf + vbCrLf, "") + "FROM ：" & strTemp + vbCrLf
         Else
         'end 2020/03/11
             stContent = "TO ：收受人姓名：" + strToName + vbCrLf + vbCrLf + txt1(1) + vbCrLf + vbCrLf + Space(11) + txt1(0) + vbCrLf + vbCrLf + "FROM ：" & strTemp + vbCrLf
         End If 'Added by Lydia 2020/03/11
         
         'Add by Morgan 2009/3/5 加可選擇要有寄件備份
         If chkByOutLook.Value = 1 Then
            'Modified by Lydia 2020/03/13 因為OutLook過去和現在版本不同,所以改用密件副本保留
'            DoEvents
'            MAPISession1.LogonUI = False
'            MAPISession1.UserName = strUserNum
'            Err.Clear
'on error Resume Next
'            MAPISession1.SignOn
'            If Err.Number <> 0 Then
'               MsgBox "EMail發送失敗!!請啟動 OutLook 後重試!!"
'               Screen.MousePointer = vbDefault
'               Exit Sub
'            End If
'            MAPIMessages1.SessionID = MAPISession1.SessionID
'            MAPIMessages1.MsgIndex = -1
'            MAPIMessages1.Compose
'            'Modify By Sindy 2014/1/16
'            'MAPIMessages1.MsgSubject = "◎系統代發◎" & stSubject
'            MAPIMessages1.MsgSubject = "◎" & IIf(Pub_StrUserSt03 = "M51" And PUB_GetST05(strUserNum) <> "", PUB_GetDbTerminal, "") & stSubject
'            '2014/1/16 END
'            MAPIMessages1.MsgNoteText = stContent
'            MAPIMessages1.RecipIndex = 0
'            MAPIMessages1.RecipType = 1 '收件者是主要收件者 'Add By Sindy 2012/3/29
'            MAPIMessages1.RecipDisplayName = ChkMailId(strTo)
'            'Add By Sindy 2012/3/29
'            '副本
'            j = 0
'            For i = 0 To lstMailCC.ListCount - 1
'               If lstMailCC.Selected(i) = True Then
'                  j = j + 1
'                  MAPIMessages1.RecipIndex = j
'                  MAPIMessages1.RecipType = 2 '收件者屬於「副本」收件者
'                  MAPIMessages1.RecipDisplayName = Left(Trim(lstMailCC.List(i)), 5)
'               End If
'            Next
'            '2012/3/29 End
'            MAPIMessages1.ResolveName
'            MAPIMessages1.Send
'            MAPISession1.SignOff
            PUB_SendMail strUserNum, strTo, "", stSubject, stContent, "", , , , , strToCC, , , , , , strUserNum
            'end 2020/03/13
         Else
         'end 2009/3/5
            'Modify By Sindy 2012/3/29 加寄副本
            'PUB_SendMail strUserNum, strTo, "", stSubject, stContent, ""
            PUB_SendMail strUserNum, strTo, "", stSubject, stContent, "", , , , , strToCC
         End If
         
         'edit by nickc 2008/05/12 會有寄信失敗又秀出成功的問題
         's = MsgBox("郵件已送出", , "MAIL!!")
         Screen.MousePointer = vbDefault
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 1
        tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case Else
      
   End Select

End Sub

Private Sub cmdGoInput_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Private Sub Form_Load()
   
   'Added by Lydia 2024/06/26
   Me.Height = 5700
   Me.Width = 5200
   'end 2024/06/26
   
   bolToEndByNick = False
   MoveFormToCenter Me
   For i = 0 To 2
       lbl1(i).Caption = ""
   Next i
   
   'Added by Lydia 2020/03/11
   If Pub_StrUserSt03 = "M51" Then
      If MsgBox("請問是否以外專人員身份操作？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
         m_Role = "F"
      End If
   End If
   'end 2020/03/11
   
   'Added by Lydia 2021/04/29 預設備註
   If strTypeMemo <> "" Then
        txt1(0).Text = strTypeMemo
   Else
   'end 2021/04/29
        'Modify By Cheng 2002/05/01
        'txt1(0).Text = "本案期限將至，請儘速作業，已利後續作業。"
        txt1(0).Text = "本案期限將至，請儘速作業，以利後續作業。"
   End If 'Added by Lydia 2021/04/29
   
   'Added by Lydia 2020/03/11
   txt2(0).Text = txt1(0).Text
   txt2(1).Text = txt1(1).Text
   If m_Role = "F" And TypeName(mPrevForm) = "frm100106_2" Then '從"以期限管制日查詢by法限"
       txt2(0).Text = "本案法定期限為" & m_CP07 & "，請今日完成送件(或延期)，謝謝。"
   End If
   '依人員區分顯示Frame: A-其他, F-外專
   Frame1.BackColor = &H8000000F
   Frame2.BackColor = &H8000000F
   Frame3.BackColor = &H8000000F
   Frame2.Visible = False
   Frame2.Left = Frame1.Left
   Frame1.Visible = False
   If m_Role = "F" Then
      Frame2.Visible = True
   Else
      Frame1.Visible = True
   End If
   'end 2020/03/11
   
   '92.04.16 nick
   cmdState = -1
   'Add by Morgan 2009//3/5 只有北所可選要有寄件備份
   If pub_strUserOffice <> "1" Then
      chkByOutLook.Visible = False
   End If
   
   'Add By Sindy 2012/3/29
   '副本：全部員工
   Me.lstMailCC.Clear
   strSql = "SELECT st01,st02 FROM staff,SalaryData WHERE st04='1' and st01>'6' and st01<'F' and st01=sd01 order by st01 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         RsTemp.MoveFirst
         Do While RsTemp.EOF = False
            lstMailCC.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
            RsTemp.MoveNext
         Loop
      End With
   End If
End Sub

Sub StrMenu()
'Added by Lydia 2017/01/19
Dim tmpArr As Variant
Dim inJ As Integer

   Str01 = SystemNumber(Me.Tag, 1)
   Str02 = SystemNumber(Me.Tag, 2)
   Str03 = SystemNumber(Me.Tag, 3)
   lbl1(0).Caption = Str01
   lbl1(1).Caption = Str02
   lbl1(2).Caption = Str03
   CheckOC
   strSql = "SELECT ST02 FROM STAFF WHERE ST01='" & strUserNum & "'"
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       If Not IsNull(adoRecordset.Fields(0)) Then
           strTemp = adoRecordset.Fields(0)
       Else
           strTemp = strUserNum
       End If
   Else
       strTemp = strUserNum
   End If
   CheckOC

   'Add by Morgan 2007/1/9
   strFCP201State = "0"
   If strCP09 <> "" Then
      'edit by nickc 2008/01/10 秀玲說未收文的智權人員，要抓下一程序智權人員
      'strExc(0) = "select cp01,cp10,cp13,st3.st02 cp13n,cp14,st1.st02 cp14n,st1.st03 cp14d,ep04,st2.st02 ep04n,ep09 from caseprogress, engineerprogress,staff st1,staff st2,staff st3 where cp09='" & strCP09 & "' and ep02(+)=cp09 and st1.st01(+)=cp14 and st2.st01(+)=ep04 and st3.st01(+)=cp13"
      '2009/11/16 modify by sonia 智權人員離職要抓主管故加是否在職欄
      'strExc(0) = "select cp01,cp10,cp13,st3.st02 cp13n,cp14,st1.st02 cp14n,st1.st03 cp14d,ep04,st2.st02 ep04n,ep09,np10,st4.st02 np10n, from caseprogress, engineerprogress,staff st1,staff st2,staff st3,nextprogress,staff st4 where cp09='" & strCP09 & "' and ep02(+)=cp09 and st1.st01(+)=cp14 and st2.st01(+)=ep04 and st3.st01(+)=cp13 and cp09=np01(+) and np10=st4.st01(+) "
      strExc(0) = "select cp01,cp10,cp13,st3.st02 cp13n,cp14,st1.st02 cp14n,st1.st03 cp14d,ep04,st2.st02 ep04n,ep09,np10,st4.st02 np10n,st1.st04 CP14ST04,st3.st04 CP13ST04,st4.st04 NP10ST04 from caseprogress, engineerprogress,staff st1,staff st2,staff st3,nextprogress,staff st4 where cp09='" & strCP09 & "' and ep02(+)=cp09 and st1.st01(+)=cp14 and st2.st01(+)=ep04 and st3.st01(+)=cp13 and cp09=np01(+) and np10=st4.st01(+) "
      
      'Add by Morgan 2010/5/21 未收文要排除程序管制的案件性質
      If frm100106_1.opt2(0).Value = True Then
         strExc(0) = strExc(0) & strNpSqlOfNoSalesDuty
      End If
      
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         'edit by nickc 2008/01/10 若是未收文，應抓下一程序
         If frm100106_1.opt2(0).Value = True Then
            StrMailNum2 = "" & .Fields("np10")
            lbl1(1).Caption = "" & .Fields("np10n")
            '2009/11/16 add by sonia 離職或虛建智權人員抓主管A0908
            If .Fields("NP10ST04") <> "1" Or StrMailNum2 < "6" Then
               CheckOC
               strSql = "SELECT A0908,S2.ST02 FROM STAFF S1,STAFF S2,ACC090 WHERE S1.ST01='" & StrMailNum2 & "' AND S1.ST04<>'1' AND S1.ST15=A0901(+) AND A0908=S2.ST01(+) AND S2.ST04='1' "
               adoRecordset.CursorLocation = adUseClient
               adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
                   StrMailNum2 = adoRecordset.Fields(0)
                   lbl1(1).Caption = lbl1(1).Caption & "　原智權人員離職改通知　　" & adoRecordset.Fields(1)
               End If
            End If
            '2009/11/16 end
         Else
            StrMailNum2 = "" & .Fields("cp13")
            lbl1(1).Caption = "" & .Fields("cp13n")
            '2009/11/16 add by sonia 離職或虛建智權人員抓主管A0908
            If .Fields("CP13ST04") <> "1" Or StrMailNum2 < "6" Then
               CheckOC
               strSql = "SELECT A0908,S2.ST02 FROM STAFF S1,STAFF S2,ACC090 WHERE S1.ST01='" & StrMailNum2 & "' AND S1.ST04<>'1' AND S1.ST15=A0901(+) AND A0908=S2.ST01(+) AND S2.ST04='1' "
               adoRecordset.CursorLocation = adUseClient
               adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
                   StrMailNum2 = adoRecordset.Fields(0)
                   lbl1(1).Caption = lbl1(1).Caption & "　原智權人員離職改通知　　" & adoRecordset.Fields(1)
               End If
            End If
            '2009/11/16 end
        End If
        StrMailNum3 = "" & .Fields("cp14")
        lbl1(2).Caption = "" & .Fields("cp14n")
         '2009/11/16 add by sonia 離職抓主管A0908
         If .Fields("CP14ST04") <> "1" Then
            CheckOC
            strSql = "SELECT A0908,S2.ST02 FROM STAFF S1,STAFF S2,ACC090 WHERE S1.ST01='" & StrMailNum3 & "' AND S1.ST04<>'1' AND S1.ST15=A0901(+) AND A0908=S2.ST01(+) AND S2.ST04='1' "
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
                'Modify by Morgan 2009/11/18
                'StrMailNum2 = adoRecordset.Fields(0)
                StrMailNum3 = adoRecordset.Fields(0)
                lbl1(2).Caption = lbl1(2).Caption & "　原承辦人離職改通知　" & adoRecordset.Fields(1)
            End If
         End If
         '2009/11/16 end
         'FCP的翻譯時加可選擇核稿人
         If .Fields("cp01") = "FCP" Then
            If .Fields("cp10") = "201" Then
               lbl1(2).Caption = lbl1(2).Caption & "(" & StrMailNum3 & ")"
               Me.Option1(3).Visible = True
               Me.lbl1(3).Visible = True
               'Added by Lydia 2020/03/11
               Me.ChkMail(3).Visible = True
               Me.lbl2(3).Visible = True
               'end 2020/03/11
               Me.lbl1(3).Caption = "" & .Fields("ep04n")
               StrMailNum4 = "" & .Fields("ep04")
               Label2.Visible = True
               Label2.Caption = "完稿日："
               If Not IsNull(.Fields("ep09")) Then
                  Label2.Caption = Label2.Caption & ChangeWStringToTDateString(.Fields("ep09"))
                  '外譯且已完稿
                  If "" & .Fields("cp14d") = "F51" Then
                     strFCP201State = "1"
                  End If
               Else
                  '外譯且未完稿
                  If "" & .Fields("cp14d") = "F51" Then
                     strFCP201State = "2"
                  End If
               End If
            End If
         Else
            'Add by Morgan 2007/3/28
            Option1(1).Value = True
            lbl1(0).Visible = False
            Option1(0).Visible = False
         End If
         End With
      End If
   'Added by Lydia 2017/01/19 增加國外部專利處期限通知及外專非台灣案已達約定期限通知功能
   Else
      If strFRname <> "" Then
         tmpArr = Split(Me.Tag, "-")
         For inJ = 0 To UBound(tmpArr)
            If Trim(tmpArr(inJ)) <> "" Then
               'Added by Lydia 2023/12/22
               If strSrvDate(1) >= 新部門啟用日 Then
                  strSql = "select s1.st02 s1name,s1.st04 s1state,s1.st03,nvl(a0924,a0908) a0908,nvl(s3.st01,s2.st01) s2no,nvl(s3.st02,s2.st02) s2name " & _
                           "from staff s1,staff s2,acc090,acc090new,staff s3 where s1.st01='" & tmpArr(inJ) & "' and s1.st15=a0901(+) and a0908=s2.st01(+) and s1.st93=a0921(+) and a0924=s3.st01(+) "
               Else
               'end 2023/12/22
                  'Modified by Lydia 2023/12/22 +s1.st03
                  strSql = "SELECT S1.ST02 S1NAME,S1.ST04 S1STATE,s1.st03,A0908,S2.ST01 S2NO,S2.ST02 S2NAME FROM STAFF S1,STAFF S2,ACC090 WHERE S1.ST01='" & tmpArr(inJ) & "' AND S1.ST15=A0901(+) AND A0908=S2.ST01(+) "
               End If
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  strExc(1) = Trim(tmpArr(inJ))
                  If InStr("" & RsTemp.Fields("S1NAME"), "林信昌") > 0 Then
                     strExc(1) = "68007"
                  End If
                  Select Case inJ
                      Case 0: StrMailNum1 = strExc(1)
                      Case 1: StrMailNum2 = strExc(1)
                      Case 2: StrMailNum3 = strExc(1)
                  End Select
                  lbl1(inJ).Caption = IIf(strExc(1) = "68007", "林信昌", "" & RsTemp.Fields("S1NAME"))
                  
                  If "" & RsTemp.Fields("S1STATE") <> "1" Then
                     'Added by Lydia 2023/12/22 判斷外專工程師
                     If "" & RsTemp.Fields("ST03") = "F21" Then
                         strSql = PUB_GetFCPEngSup(Trim(tmpArr(inJ)))
                         lbl1(inJ).Caption = lbl1(inJ).Caption & "　原" & Trim(Option1(inJ).Caption) & "離職改通知　" & GetStaffName(strSql)
                         Select Case inJ
                             Case 0: StrMailNum1 = strSql
                             Case 1: StrMailNum2 = strSql
                             Case 2: StrMailNum3 = strSql
                         End Select
                     Else
                     'end 2023/12/22
                         lbl1(inJ).Caption = lbl1(inJ).Caption & "　原" & Trim(Option1(inJ).Caption) & "離職改通知　" & RsTemp.Fields("S2NAME")
                     End If
                  End If
               End If
            End If
         Next inJ
      End If
   End If
   'end 2007/1/9
   
   'Added by Lydia 2021/05/18 Form 2.0的label沒有Change：將lbl1的值,丟給lbl2
   For Each oObj In lbl1
       lbl2(oObj.Index) = oObj.Caption
   Next
   'end 2021/05/18
   
   Call GetPersonMan 'Added by Lydia 2020/03/11
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100106_4 = Nothing
End Sub

'Add by Morgan 2010/6/7
Private Sub Option1_Click(Index As Integer)
   If Index = 4 Then
      Text1.Enabled = True
      Text1.SetFocus
   Else
      Text1.Enabled = False
      Text1 = ""
      lbl1(4) = ""
   End If
End Sub

'Add by Morgan 2010/6/7
Private Sub Text1_Change()
   lbl1(4) = ""
   If Len(Text1) >= 5 Then
      If ClsPDGetStaff(Text1, strExc(0)) Then
         lbl1(4) = strExc(0)
      End If
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

'Add By Sindy 2010/11/25
Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   'Add By Cheng 2002/05/01
   Select Case Index
   Case 0
      '切換至中文輸入模式
      'edit by nickc 2007/06/06 切換輸入法改用API
      'Me.txt1(Index).IMEMode = 1
      OpenIme
   Case 1
      '切換至中文輸模式
      'edit by nickc 2007/06/06 切換輸入法改用API
      'Me.txt1(Index).IMEMode = 1
      OpenIme
   End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   'Add By Cheng 2002/05/01
   Select Case Index
   Case 0
      '取消中文輸入模式
      'edit by nickc 2007/06/06 切換輸入法改用API
      'Me.txt1(Index).IMEMode = 2
      CloseIme
   Case 1
      '取消中文輸入模式
      'edit by nickc 2007/06/06 切換輸入法改用API
      'Me.txt1(Index).IMEMode = 2
      CloseIme
   End Select
End Sub

'Added by Lydia 2020/03/11
'Remove by Lydia 2021/05/18 Form 2.0的label沒有Change
'Private Sub lbl1_Change(Index As Integer)
'  '將lbl1的值,丟給lbl2
'  lbl2(Index).Caption = lbl1(Index).Caption
'End Sub
'end 2021/05/18

'Added by Lydia 2020/03/11 抓第二級管制人
Private Sub GetPersonMan()
Dim intP As Integer
Dim strTmpA As String

  '抓第二級管制人
    For intP = 0 To 4
        If lbl1(intP).Caption = "" Then
             lbl2(intP).Tag = ""
        Else
             Select Case intP
                 Case 0 '管制人
                     lbl2(intP).Tag = PUB_GetFCPProSup(StrMailNum1)
                 Case 1 '智權人員
                     lbl2(intP).Tag = PUB_GetFCPProSup(StrMailNum2)
                 Case 2 '承辦人
                     If InStr("F21,F51", GetST15(StrMailNum3)) > 0 Then '外專工程師
                         'Modified by Lydia 2024/06/26 CC: 第二級管制人(主任)＋主管
                         'Lbl2(intP).Tag = PUB_GetFCPEngSup(StrMailNum3)
                         lbl2(intP).Tag = PUB_GetFCPEngSup(StrMailNum3, , , True)
                     Else
                         lbl2(intP).Tag = PUB_GetFCPProSup(StrMailNum3)
                     End If
                 Case 3 '核稿人
                     If InStr("F21,F51", GetST15(StrMailNum4)) > 0 Then '外專工程師
                         'Modified by Lydia 2024/06/26 CC: 第二級管制人(主任)＋主管
                         'Lbl2(intP).Tag = PUB_GetFCPEngSup(StrMailNum4)
                         lbl2(intP).Tag = PUB_GetFCPEngSup(StrMailNum4, , , True)
                     Else
                         lbl2(intP).Tag = PUB_GetFCPProSup(StrMailNum4)
                     End If
             End Select
             '如果人員已離職則收件人自動帶主管，所以不用再抓期限管制人
             If InStr(lbl2(intP).Caption, "離職") > 0 Then
                 lbl2(intP).Tag = ""
             End If
        End If
    Next intP
End Sub
'Added by Lydia 2020/03/11
Private Sub ChkMail_Click(Index As Integer)
Dim tmpArr As Variant
Dim intP As Integer
Dim strTmp As String

   If Index < 4 And lbl2(Index).Tag <> "" Then
       tmpArr = Split(lbl2(Index).Tag, ";")
       strTmp = ";" & TxtAddCC.Text
       For intP = 0 To UBound(tmpArr)
           '副本:自動帶入點選" 管制人,智權人員,工程師"之各自所屬主管(二級期限管制人),唯"日文組工程師"需要通知所屬的二、三級期限管制人
           If Trim(tmpArr(intP)) <> "" Then
                If ChkMail(Index).Value = 1 Then '帶入
                    If InStr(strTmp, tmpArr(intP)) = 0 Then
                        strTmp = strTmp & ";" & tmpArr(intP)
                    End If
                Else   '取消
                    If InStr(strTmp, tmpArr(intP)) > 0 Then
                        strTmp = Replace(strTmp, ";" & tmpArr(intP), "")
                    End If
                End If
           End If
       Next intP
       If strTmp <> "" Then
           strTmp = Replace(strTmp, ";;", ";")
       End If
       If Left(strTmp, 1) = ";" Then
           strTmp = Mid(strTmp, 2)
       End If
       TxtAddCC.Text = strTmp
       
       Call TxtAddCC_Validate(False)
   End If
End Sub

Private Sub TxtAddCC_GotFocus()
   TextInverse TxtAddCC
End Sub

Private Sub TxtAddCC_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub TxtAddCC_Validate(Cancel As Boolean)
   If TxtAddCC.Tag <> TxtAddCC.Text Then
        TxtAddCC.Text = Replace(TxtAddCC.Text, ",", ";")
        If TxtAddCC.Text = "" Then
            TxtAddCCName.Text = ""
        Else
            '副本:收件人名稱
            TxtAddCCName.Text = PUB_ReadUserData(TxtAddCC.Text)
        End If
   End If
   TxtAddCC.Tag = TxtAddCC.Text
End Sub

Private Sub TxtAddTo_GotFocus()
   TextInverse TxtAddTo
End Sub

Private Sub TxtAddTo_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub TxtAddTo_Validate(Cancel As Boolean)
   If TxtAddTo.Tag <> TxtAddTo.Text Then
        TxtAddTo.Text = Replace(TxtAddTo.Text, ",", ";")
        If TxtAddTo.Text = "" Then
            lbl2(4).Caption = ""
        Else
            '收件人:其他收件人名稱
            lbl2(4).Caption = PUB_ReadUserData(TxtAddTo.Text)
        End If
   End If
   TxtAddTo.Tag = TxtAddTo.Text
End Sub
'end 2020/03/11

