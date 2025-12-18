VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100109_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "以收文日查詢來函"
   ClientHeight    =   3480
   ClientLeft      =   1710
   ClientTop       =   2820
   ClientWidth     =   4970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4970
   Begin VB.CheckBox Check1 
      Caption         =   "排除特定來函"
      Height          =   300
      Left            =   1260
      TabIndex        =   13
      Top             =   2872
      Width           =   1455
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   11
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   11
      Top             =   2520
      Width           =   675
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   12
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   12
      Top             =   2520
      Width           =   675
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   10
      Left            =   2100
      MaxLength       =   4
      TabIndex        =   10
      Top             =   2220
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   9
      Left            =   1260
      MaxLength       =   4
      TabIndex        =   9
      Top             =   2220
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   8
      Left            =   2100
      MaxLength       =   4
      TabIndex        =   8
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   7
      Left            =   1260
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3360
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   0
      Width           =   756
   End
   Begin VB.CommandButton cmdGoInput 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4130
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   0
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   6
      Left            =   2100
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1608
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   2
      Top             =   744
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1032
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   1260
      TabIndex        =   4
      Top             =   1320
      Width           =   1600
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   5
      Left            =   1260
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1608
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1260
      MaxLength       =   7
      TabIndex        =   0
      Top             =   456
      Width           =   1092
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2700
      MaxLength       =   7
      TabIndex        =   1
      Top             =   456
      Width           =   1092
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   1
      Left            =   2016
      TabIndex        =   22
      Top             =   1080
      Width           =   1932
      Size            =   "3408;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   0
      Left            =   2016
      TabIndex        =   21
      Top             =   792
      Width           =   1932
      Size            =   "3408;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCheck2 
      Caption         =   ",通知年費逾繳,專利權消滅,...)"
      Height          =   255
      Left            =   1230
      TabIndex        =   27
      Top             =   3195
      Width           =   2925
   End
   Begin VB.Label lblCheck1 
      Caption         =   "(通知申請案號"
      Height          =   255
      Left            =   2760
      TabIndex        =   26
      Top             =   2895
      Width           =   1365
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   2040
      X2              =   2160
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FCP管制人："
      Height          =   180
      Index           =   7
      Left            =   60
      TabIndex        =   25
      Top             =   2550
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人國籍："
      Height          =   180
      Index           =   6
      Left            =   60
      TabIndex        =   24
      Top             =   2250
      Width           =   1080
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   1860
      X2              =   1980
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   23
      Top             =   1950
      Width           =   900
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   1860
      X2              =   1980
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   1860
      X2              =   1980
      Y1              =   1728
      Y2              =   1728
   End
   Begin VB.Line Line1 
      X1              =   2436
      X2              =   2556
      Y1              =   576
      Y2              =   576
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文日期："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   20
      Top             =   480
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Index           =   2
      Left            =   60
      TabIndex        =   19
      Top             =   765
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   3
      Left            =   60
      TabIndex        =   18
      Top             =   1050
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別：                                           (ALL：全部)"
      Height          =   180
      Index           =   4
      Left            =   60
      TabIndex        =   17
      Top             =   1350
      Width           =   3825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   5
      Left            =   60
      TabIndex        =   16
      Top             =   1635
      Width           =   900
   End
End
Attribute VB_Name = "frm100109_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/20 改成Form2.0(lbl1(0),lbl1(1))
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/14 日期欄已修改
Option Explicit

Dim s As Integer, strSql As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Add by Morgan 2004/4/28
Dim bolDataOk As Boolean

'92.04.16 nick
Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
            cmdState = -1
            'Modify by Morgan 2004/4/28
            
      '      If PUB_CheckKeyInDate(Me.txt1(0)) = -1 Then
      '         Me.txt1(0).SetFocus
      '         txt1_GotFocus 0
      '         Exit Sub
      '      End If
      '      If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
      '         Me.txt1(1).SetFocus
      '         txt1_GotFocus 1
      '         Exit Sub
      '      End If
      '      txt1(0).SetFocus 'Remove by Morgan 2004/9/3
            Dim iIdx As Integer
            bolDataOk = False
            For iIdx = 0 To 12
               txt1_LostFocus iIdx
               If bolDataOk = False Then
                  Exit Sub
               Else
                  bolDataOk = False
               End If
            Next
            'Modify end
            
           If Len(Trim(txt1(0))) = 0 Then
              s = MsgBox("收文日區間起點不可空白", , "USER 輸入錯誤")
              txt1(0).SetFocus
              txt1(0).SelStart = 0
              txt1(0).SelLength = Len(txt1(0))
              Exit Sub
           End If
           If Len(Trim(txt1(1))) = 0 Then
              s = MsgBox("收文日區間終點不可空白", , "USER 輸入錯誤")
              txt1(1).SetFocus
              txt1(1).SelStart = 0
              txt1(1).SelLength = Len(txt1(1))
              Exit Sub
           End If
           Me.Enabled = False
          If fnSaveParentForm(Me) = False Then
              Me.Enabled = True
              Exit Sub
          End If
          ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/3 清除查詢印表記錄檔欄位
           Screen.MousePointer = vbHourglass
           frm100109_2.Show
           frm100109_2.StrMenu
           Screen.MousePointer = vbDefault
           Me.Enabled = True
           Exit Sub
      Case 1
           fnCloseAllFrm100
      Case Else
   End Select
End Sub

Private Sub cmdGoInput_Click(Index As Integer)
   'add by nickc 2007/01/12
   If Len(Trim(Me.txt1(4).Text)) = 0 Then
       Me.txt1(4).Text = "ALL"
   End If
   '92.04.16 nick 紀錄作用按鍵
   cmdState = Index
   PubShowNextData
   Exit Sub
''92.04.16 nick 以下無效
'Select Case Index
'Case 0
'      'Modify By Cheng 2002/03/14
''      'Add By Cheng 2002/01/07
''      txt1_LostFocus 4
'      'Add By Cheng 2002/03/18
'
'      If PUB_CheckKeyInDate(Me.txt1(0)) = -1 Then
'         Me.txt1(0).SetFocus
'         txt1_GotFocus 0
'         Exit Sub
'      End If
'      If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
'         Me.txt1(1).SetFocus
'         txt1_GotFocus 1
'         Exit Sub
'      End If
'
'     If Len(Trim(txt1(0))) = 0 Then
'        s = MsgBox("收文日區間起點不可空白", , "USER 輸入錯誤")
'        txt1(0).SetFocus
'        txt1(0).SelStart = 0
'        txt1(0).SelLength = Len(txt1(0))
'        Exit Sub
'     End If
'     If Len(Trim(txt1(1))) = 0 Then
'        s = MsgBox("收文日區間終點不可空白", , "USER 輸入錯誤")
'        txt1(1).SetFocus
'        txt1(1).SelStart = 0
'        txt1(1).SelLength = Len(txt1(1))
'        Exit Sub
'     End If
'     Me.Enabled = False
'     Screen.MousePointer = vbHourglass
'     frm100109_2.Show
'     'frm100109_2.Hide
'
'     frm100109_2.StrMenu
'     Screen.MousePointer = vbDefault
'     Me.Hide
'     'frm100109_2.Show
'     Do
'     DoEvents
'     If bolToEndByNick = True Then Unload Me: Exit Sub
'     Loop Until Not frm100109_2.Visible
'     Unload frm100109_2
'     Me.Enabled = True
'     Me.Show
'Case 1
'      Unload Me
'Case Else
'End Select
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
      MoveFormToCenter Me
   txt1(4) = Systemkind_g
   '92.04.16 nick
   cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100109_1 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Select Case Index
      Case 0, 1
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Exit Sub
         End If
         If Index = 1 Then
         If RunNick(txt1(Index - 1), txt1(Index)) Then
              txt1(Index - 1).SetFocus
              txt1_GotFocus (Index - 1)
              Exit Sub
         End If
         End If
      Case 2
           If Len(txt1(Index)) <> 0 Then
                 lbl1(0).Caption = GetPrjSalesNM(txt1(Index))
                 If Trim(lbl1(0).Caption) = "" Then
                     s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
                     txt1(Index).SetFocus
                     txt1_GotFocus (Index)
                     Exit Sub
                 End If
            Else
              lbl1(0) = ""
            End If
      Case 3
            If Len(txt1(Index)) <> 0 Then
                 lbl1(1).Caption = GetPrjSalesNM(txt1(Index))
                 If Trim(lbl1(1).Caption) = "" Then
                     s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
                     txt1(Index).SetFocus
                     txt1_GotFocus (Index)
                     Exit Sub
                 End If
            Else
                 lbl1(1) = ""
            End If
      Case 4 '系統類別
            'Modify By Cheng 2002/03/14
      '      'Add By Cheng 2002/01/07
      '      Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
      Case 5
      Case 6, 8, 10, 12
         If RunNick(txt1(Index - 1), txt1(Index)) Then
              txt1(Index - 1).SetFocus
              txt1_GotFocus (Index - 1)
              Exit Sub
         End If
      Case Else
   End Select
   bolDataOk = True
End Sub
