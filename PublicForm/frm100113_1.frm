VERSION 5.00
Begin VB.Form frm100113_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "申請人查詢案件變更紀錄"
   ClientHeight    =   4290
   ClientLeft      =   570
   ClientTop       =   2910
   ClientWidth     =   5460
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5460
   Begin VB.CheckBox chk 
      Caption         =   "所有系統類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame fraTF 
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1980
      TabIndex        =   36
      Top             =   3840
      Width           =   2300
      Begin VB.TextBox txt3 
         Height          =   288
         Index           =   3
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   25
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txt3 
         Height          =   288
         Index           =   2
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   24
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txt3 
         Height          =   288
         Index           =   1
         Left            =   0
         MaxLength       =   6
         TabIndex        =   23
         Top             =   0
         Width           =   972
      End
   End
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   288
      Index           =   0
      Left            =   1980
      MaxLength       =   6
      TabIndex        =   35
      Top             =   3840
      Width           =   1212
   End
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   288
      Index           =   1
      Left            =   3300
      MaxLength       =   1
      TabIndex        =   34
      Top             =   3840
      Width           =   372
   End
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   288
      Index           =   2
      Left            =   3780
      MaxLength       =   2
      TabIndex        =   33
      Top             =   3840
      Width           =   492
   End
   Begin VB.TextBox txt3 
      Height          =   288
      Index           =   0
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   22
      Top             =   3840
      Width           =   732
   End
   Begin VB.CommandButton cmdGoInput 
      Cancel          =   -1  'True
      Caption         =   "全部勾選(&S)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   4020
      Style           =   1  '圖片外觀
      TabIndex        =   32
      Top             =   480
      Width           =   1400
   End
   Begin VB.CheckBox Check1 
      Caption         =   "其他"
      Height          =   180
      Index           =   16
      Left            =   3036
      TabIndex        =   18
      Top             =   2850
      Width           =   1245
   End
   Begin VB.CheckBox Check1 
      Caption         =   "變更圖樣"
      Height          =   180
      Index           =   15
      Left            =   1152
      TabIndex        =   17
      Top             =   2850
      Width           =   1080
   End
   Begin VB.CheckBox Check1 
      Caption         =   "變更正商標號數"
      Height          =   180
      Index           =   14
      Left            =   3036
      TabIndex        =   16
      Top             =   2580
      Width           =   1785
   End
   Begin VB.CheckBox Check1 
      Caption         =   "變更代理人"
      Height          =   180
      Index           =   13
      Left            =   1152
      TabIndex        =   15
      Top             =   2580
      Width           =   1245
   End
   Begin VB.CheckBox Check1 
      Caption         =   "變更代表人印鑑"
      Height          =   180
      Index           =   12
      Left            =   3036
      TabIndex        =   14
      Top             =   2310
      Width           =   1665
   End
   Begin VB.CheckBox Check1 
      Caption         =   "變更申請人印鑑"
      Height          =   180
      Index           =   11
      Left            =   1152
      TabIndex        =   13
      Top             =   2310
      Width           =   1740
   End
   Begin VB.CheckBox Check1 
      Caption         =   "變更商品組群"
      Height          =   180
      Index           =   10
      Left            =   3036
      TabIndex        =   12
      Top             =   2040
      Width           =   1470
   End
   Begin VB.CheckBox Check1 
      Caption         =   "變更商品類別"
      Height          =   180
      Index           =   9
      Left            =   1152
      TabIndex        =   11
      Top             =   2025
      Width           =   1500
   End
   Begin VB.CheckBox Check1 
      Caption         =   "減縮商品"
      Height          =   180
      Index           =   8
      Left            =   3036
      TabIndex        =   10
      Top             =   1770
      Width           =   1305
   End
   Begin VB.CheckBox Check1 
      Caption         =   "變更案件名稱"
      Height          =   180
      Index           =   7
      Left            =   1152
      TabIndex        =   9
      Top             =   1755
      Width           =   1560
   End
   Begin VB.CheckBox Check1 
      Caption         =   "變更專利商標種類"
      Height          =   180
      Index           =   6
      Left            =   3036
      TabIndex        =   8
      Top             =   1515
      Width           =   1830
   End
   Begin VB.CheckBox Check1 
      Caption         =   "變更申請地址"
      Height          =   180
      Index           =   5
      Left            =   1152
      TabIndex        =   7
      Top             =   1500
      Width           =   1485
   End
   Begin VB.CheckBox Check1 
      Caption         =   "變更代表人中譯文"
      Height          =   180
      Index           =   4
      Left            =   3036
      TabIndex        =   6
      Top             =   1275
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "變更申請人中譯文"
      Height          =   180
      Index           =   3
      Left            =   1152
      TabIndex        =   5
      Top             =   1245
      Width           =   1800
   End
   Begin VB.CheckBox Check1 
      Caption         =   "變更代表人"
      Height          =   180
      Index           =   2
      Left            =   4080
      TabIndex        =   4
      Top             =   990
      Width           =   1305
   End
   Begin VB.CheckBox Check1 
      Caption         =   "變更申請人"
      Height          =   180
      Index           =   1
      Left            =   2568
      TabIndex        =   3
      Top             =   990
      Width           =   1305
   End
   Begin VB.CheckBox Check1 
      Caption         =   "變更申請日"
      Height          =   180
      Index           =   0
      Left            =   1152
      TabIndex        =   2
      Top             =   990
      Width           =   1305
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1140
      TabIndex        =   21
      Top             =   3480
      Width           =   2700
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2616
      MaxLength       =   7
      TabIndex        =   20
      Top             =   3120
      Width           =   1212
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1152
      MaxLength       =   7
      TabIndex        =   19
      Top             =   3120
      Width           =   1212
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2520
      MaxLength       =   9
      TabIndex        =   1
      Top             =   492
      Width           =   972
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4620
      Style           =   1  '圖片外觀
      TabIndex        =   27
      Top             =   24
      Width           =   795
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3780
      Style           =   1  '圖片外觀
      TabIndex        =   26
      Top             =   24
      Width           =   795
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1200
      MaxLength       =   9
      TabIndex        =   0
      Top             =   492
      Width           =   972
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   3
      Left            =   90
      TabIndex        =   37
      Top             =   3870
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2430
      X2              =   2550
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "系統類別：                                                                  (ALL：全部)"
      Height          =   180
      Index           =   2
      Left            =   60
      TabIndex        =   31
      Top             =   3510
      Width           =   4860
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "變更收文日："
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   30
      Top             =   3150
      Width           =   1080
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2280
      X2              =   2400
      Y1              =   624
      Y2              =   624
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人編號："
      Height          =   180
      Left            =   120
      TabIndex        =   29
      Top             =   510
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "變更事項："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   28
      Top             =   1035
      Width           =   900
   End
End
Attribute VB_Name = "frm100113_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/01/05 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim s As Integer, i As Integer, j As Integer
'Add By Cheng 2002/02/20
Dim m_blnCheck As Boolean
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

'92.04.16 nick
Public Sub PubShowNextData()
Dim ii As Integer

   Select Case cmdState
      Case 0 '確定
        cmdState = -1
            '不勾選表示全勾
            If Check1(0).Value = 0 And Check1(1).Value = 0 And Check1(2).Value = 0 And Check1(3).Value = 0 And Check1(4).Value = 0 And Check1(5).Value = 0 And Check1(6).Value = 0 And Check1(7).Value = 0 And Check1(8).Value = 0 And Check1(9).Value = 0 And Check1(10).Value = 0 And Check1(11).Value = 0 And Check1(12).Value = 0 And Check1(13).Value = 0 And Check1(14).Value = 0 And Check1(15).Value = 0 And Check1(16).Value = 0 Then
               m_blnCheck = False
               cmdGoInput_Click 2
            End If
         If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
            Me.txt1(2).SetFocus
            txt1_GotFocus 2
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
            Me.txt1(3).SetFocus
            txt1_GotFocus 3
            Exit Sub
         End If
         Me.Enabled = False
       If fnSaveParentForm(Me) = False Then
           Me.Enabled = True
           Exit Sub
       End If
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/4 清除查詢印表記錄檔欄位
         Screen.MousePointer = vbHourglass
         frm100113_2.Show
         frm100113_2.SetLabelEnabled
         frm100113_2.StrMenu
         Screen.MousePointer = vbDefault
         Me.Enabled = True
      Case 1 '結束
         fnCloseAllFrm100
      Case 2 '勾選與否
         m_blnCheck = Not m_blnCheck
         '全部勾選
         If m_blnCheck = True Then
            Me.cmdGoInput(cmdState).Caption = "取消勾選(&U)"
            For ii = 0 To frm100113_1.Check1.Count - 1
               frm100113_1.Check1(ii).Value = vbChecked
            Next ii
         Else
            Me.cmdGoInput(cmdState).Caption = "全部勾選(&S)"
            For ii = 0 To frm100113_1.Check1.Count - 1
               frm100113_1.Check1(ii).Value = vbUnchecked
            Next ii
         End If
         
      Case Else
   End Select
End Sub

'2011/12/6 add by sonia
Private Sub chk_Click()
   If Me.chk.Value = vbChecked Then
       Me.txt1(4).Text = "ALL"
   Else
       Me.txt1(4).Text = Systemkind_g
   End If
End Sub
'2011/12/6 end

Private Sub cmdGoInput_Click(Index As Integer)
   'add by nickc 2007/01/12
   If Len(Trim(Me.txt1(4).Text)) = 0 Then
       Me.txt1(4).Text = "ALL"
   End If
   '92.04.16 nick 紀錄作用按鍵
   cmdState = Index
   PubShowNextData
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   bolToEndByNick = False
   '2011/12/6 modify by sonia
   'txt1(4) = Systemkind_g
   Me.chk.Value = vbChecked
   txt1(4) = "ALL"
   '2011/12/6 end
   If bolFNation = False Then
       Check1(13).Visible = False
   End If
   '92.04.16 nick
   cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100113_1 = Nothing
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
   'Add By Cheng 2002/01/07
   Select Case Index
      Case 0, 1
           If Index = 1 Then
              If Mid(txt1(Index - 1), 1, 6) <> Mid(txt1(Index), 1, 6) Then
                   s = MsgBox("前6碼必須相同！", , "錯誤！")
                   txt1(Index - 1).SetFocus
                   txt1_GotFocus (Index - 1)
                   Exit Sub
               End If
               If RunNick(txt1(Index - 1), txt1(Index)) Then
                   txt1(Index - 1).SetFocus
                   txt1_GotFocus (Index - 1)
                   Exit Sub
               End If
            End If
      Case 2, 3
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Exit Sub
         End If
         If Index = 3 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
                txt1(Index - 1).SetFocus
                txt1_GotFocus (Index - 1)
                Exit Sub
            End If
          End If
      Case 4 '系統類別
         'Modify By Cheng 2002/03/14
      '   Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
   End Select
End Sub

Private Sub TXT3_GotFocus(Index As Integer)
   'Add By Cheng 2002/11/13
   TextInverse Me.txt3(Index)
   CloseIme
End Sub

Private Sub txt3_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   'Add By Cheng 2002/11/13
   TextInverse Me.txtCode(Index)
   CloseIme
End Sub
