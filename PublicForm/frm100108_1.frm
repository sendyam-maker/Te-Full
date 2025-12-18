VERSION 5.00
Begin VB.Form frm100108_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "關聯案件資料及正聯商標查詢"
   ClientHeight    =   3015
   ClientLeft      =   3825
   ClientTop       =   1440
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5760
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
      TabIndex        =   20
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   2070
      MaxLength       =   20
      TabIndex        =   8
      Top             =   1522
      Width           =   2892
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   1140
      TabIndex        =   9
      Top             =   1863
      Width           =   3000
   End
   Begin VB.Frame fraElse 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2040
      TabIndex        =   16
      Top             =   840
      Width           =   2724
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   1
         Left            =   0
         MaxLength       =   6
         TabIndex        =   2
         Top             =   0
         Width           =   1212
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   2
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   3
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   3
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   4
         Top             =   0
         Width           =   492
      End
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
      Height          =   204
      Left            =   2085
      TabIndex        =   15
      Top             =   885
      Width           =   2652
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1245
      MaxLength       =   3
      TabIndex        =   1
      Top             =   840
      Width           =   732
   End
   Begin VB.CommandButton cmdGoInput 
      Cancel          =   -1  'True
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4092
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   48
      Width           =   756
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3300
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   48
      Width           =   756
   End
   Begin VB.OptionButton Option1 
      Caption         =   "審定號數/證書號數："
      Height          =   192
      Index           =   2
      Left            =   165
      TabIndex        =   7
      Top             =   1576
      Width           =   2025
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   1245
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1181
      Width           =   2892
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   192
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   885
      Width           =   1332
   End
   Begin VB.OptionButton Option1 
      Caption         =   "申請案號："
      Height          =   192
      Index           =   1
      Left            =   165
      TabIndex        =   5
      Top             =   1230
      Width           =   1212
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   1125
      MaxLength       =   1
      TabIndex        =   10
      Top             =   2205
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "　3.分割案"
      Height          =   180
      Index           =   1
      Left            =   1500
      TabIndex        =   19
      Top             =   2790
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "　2.正聯商標"
      Height          =   180
      Index           =   0
      Left            =   1500
      TabIndex        =   18
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "系統類別：                                                                              (ALL：全部)"
      Height          =   180
      Left            =   150
      TabIndex        =   17
      Top             =   1923
      Width           =   5400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "查詢內容："
      Height          =   180
      Index           =   2
      Left            =   165
      TabIndex        =   14
      Top             =   2250
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "　1.相關卷號"
      Height          =   180
      Left            =   1500
      TabIndex        =   13
      Top             =   2250
      Width           =   1035
   End
End
Attribute VB_Name = "frm100108_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/12/21 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/14 日期欄已修改
Option Explicit
Dim s As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

'92.04.16 nick
Public Sub PubShowNextData()
      
   Select Case cmdState
      Case 0
           cmdState = -1
           If Len(Trim(txt1(7))) = 0 Then
              s = MsgBox("查詢內容欄請勿空白", , "USER 輸入錯誤")
              txt1(7).SetFocus
              txt1(7).SelStart = 0
              txt1(7).SelLength = Len(txt1(7))
              Exit Sub
           End If
           ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/3 清除查詢印表記錄檔欄位
           '系統類別
           If Len(txt1(6)) <> 0 Then
               pub_QL05 = pub_QL05 & ";" & Left(Label2, 5) & txt1(6) 'Add By Sindy 2010/11/3
           End If
           '查詢內容
           If txt1(7) = "1" Then
               pub_QL05 = pub_QL05 & ";" & Label1(2) & Trim(Label4) 'Add By Sindy 2010/11/3
           ElseIf txt1(7) = "2" Then
               pub_QL05 = pub_QL05 & ";" & Label1(2) & Trim(Label3(0)) 'Add By Sindy 2010/11/3
           ElseIf txt1(7) = "3" Then
               pub_QL05 = pub_QL05 & ";" & Label1(2) & Trim(Label3(1)) 'Add By Sindy 2010/11/3
           End If
           '以本所案號查詢
           If Option1(0).Value = True Then
              If Len(Trim(txt1(0))) = 0 And Len(Trim(txt1(1))) = 0 Then
                  s = MsgBox("本所案號 4 個欄位請勿空白", , "USER 輸入錯誤")
                  txt1(0).SetFocus
                  txt1(0).SelStart = 0
                  txt1(0).SelLength = Len(txt1(0))
                  Exit Sub
              Else
                  'edit by nick 加入分割案
      '            If ((txt1(7) = "2" Or Me.txt1(7).Text = "3") And (txt1(0) = "CFT" Or txt1(0) = "FCT" Or txt1(0) = "T" Or txt1(0) = "TF")) _
                     Or txt1(7) = "1" Then
                  pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & txt1(0) + "-" + txt1(1) + "-" + IIf(Len(Trim(txt1(2))) = 0, "0", txt1(2)) + "-" + IIf(Len(Trim(txt1(3))) = 0, "00", txt1(3)) 'Add By Sindy 2010/11/3
                  If txt1(7) = "1" Then
                      Me.Enabled = False
                      If fnSaveParentForm(Me) = False Then
                          Me.Enabled = True
                          Exit Sub
                      End If
                      Screen.MousePointer = vbHourglass
                      frm100108_3.Show
                      frm100108_3.Tag = txt1(0) + "-" + txt1(1) + "-" + IIf(Len(Trim(txt1(2))) = 0, "0", txt1(2)) + "-" + IIf(Len(Trim(txt1(3))) = 0, "00", txt1(3))
                      frm100108_3.StrMenu
                      Screen.MousePointer = vbDefault
                      Me.Enabled = True
                      Exit Sub
                  Else
                      'add by nick 2004 加入分割案
                      'edit by nick 2004/09/14
                      'If Txt1(7).Text = "4" Or ((Txt1(7) = "2" Or Me.Txt1(7).Text = "3") And (Txt1(0) = "CFT" Or Txt1(0) = "FCT" Or Txt1(0) = "T" Or Txt1(0) = "TF")) Then
                      If txt1(7).Text = "3" Or ((txt1(7) = "2") And (txt1(0) = "CFT" Or txt1(0) = "FCT" Or txt1(0) = "T" Or txt1(0) = "TF")) Then
                          Me.Enabled = False
                          If fnSaveParentForm(Me) = False Then
                              Me.Enabled = True
                              Exit Sub
                          End If
                          Screen.MousePointer = vbHourglass
                          frm100108_4.Show
                          frm100108_4.frm100108_txt_7 = txt1(7).Text
                          frm100108_4.SetDataListWidth
                          frm100108_4.Tag = txt1(0) + "-" + txt1(1) + "-" + IIf(Len(Trim(txt1(2))) = 0, "0", txt1(2)) + "-" + IIf(Len(Trim(txt1(3))) = 0, "00", txt1(3))
                          frm100108_4.StrMenu
                          Screen.MousePointer = vbDefault
                          Me.Enabled = True
                          Exit Sub
                      Else
                          s = MsgBox("此本所案號沒有正聯商標, 無法查詢!!" & txt1(0) + "-" + txt1(1) + "-" + IIf(Len(Trim(txt1(2))) = 0, "0", txt1(2)) + "-" + IIf(Len(Trim(txt1(3))) = 0, "00", txt1(3)) & "  ", , "錯誤")
                      End If
                  End If
              End If
           Else
              '以申請案號查詢
              If Option1(1).Value = True Then
                  If Len(Trim(txt1(4))) = 0 Then
                      s = MsgBox("申請案號欄請勿空白", , "USER 輸入錯誤")
                      txt1(4).SetFocus
                      txt1(4).SelStart = 0
                      txt1(4).SelLength = Len(txt1(4))
                      Exit Sub
                  Else
                      Me.Enabled = False
                      If fnSaveParentForm(Me) = False Then
                          Me.Enabled = True
                          Exit Sub
                      End If
                      pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & txt1(4) 'Add By Sindy 2010/11/3
                      Screen.MousePointer = vbHourglass
                      frm100108_2.Show
                      frm100108_2.StrMenu
                      Screen.MousePointer = vbDefault
                      Me.Enabled = True
                      Exit Sub
                  End If
              '以審定號/證書號查詢
              Else
                  If Len(Trim(txt1(5))) = 0 Then
                      s = MsgBox("審定號數/證書號數欄請勿空白", , "USER 輸入錯誤")
                      txt1(5).SetFocus
                      txt1(5).SelStart = 0
                      txt1(5).SelLength = Len(txt1(5))
                      Exit Sub
                  Else
                      Me.Enabled = False
                      If fnSaveParentForm(Me) = False Then
                          Me.Enabled = True
                          Exit Sub
                      End If
                      pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & txt1(5) 'Add By Sindy 2010/11/3
                      Screen.MousePointer = vbHourglass
                      frm100108_2.Show
                      frm100108_2.StrMenu
                      Screen.MousePointer = vbDefault
                      Me.Enabled = True
                      Exit Sub
                  End If
              End If
           End If
      Case 1
           fnCloseAllFrm100
      Case Else
   End Select
End Sub

'2011/12/6 add by sonia
Private Sub chk_Click()
   If Me.Chk.Value = vbChecked Then
      Me.txt1(6).Text = "ALL"
   Else
      Me.txt1(6).Text = Systemkind_g
   End If
End Sub
'2011/12/6 end

Private Sub cmdGoInput_Click(Index As Integer)
   'add by nickc 2007/01/12
   If Len(Trim(Me.txt1(6).Text)) = 0 Then
       Me.txt1(6).Text = "ALL"
   End If
   '92.04.16 nick 紀錄作用按鍵
   cmdState = Index
   PubShowNextData
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   '2011/12/6 modify by sonia
   'txt1(6) = Systemkind_g
   Me.Chk.Value = vbChecked
   txt1(6) = "ALL"
   '2011/12/6 end
   '92.04.16 nick
   cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100108_1 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   
   Select Case Index
      Case 0
           If Option1(0).Value = True Then
               Option1(1).Value = False
               'txt1(4).Enabled = False
               Option1(2).Value = False
               'txt1(5).Enabled = False
               'txt1(0).Enabled = True
               'txt1(1).Enabled = True
               'txt1(2).Enabled = True
               'txt1(3).Enabled = True
               txt1(0).SetFocus
               txt1_GotFocus (0)
           End If
      Case 1
            If Option1(1).Value = True Then
              Option1(0).Value = False
              'txt1(0).Enabled = False
              'txt1(1).Enabled = False
              'txt1(2).Enabled = False
              'txt1(3).Enabled = False
              Option1(2).Value = False
              'txt1(5).Enabled = False
              'txt1(4).Enabled = True
              txt1(4).SetFocus
               txt1_GotFocus (4)
          End If
      Case 2
           If Option1(2).Value = True Then
              Option1(0).Value = False
              'txt1(0).Enabled = False
              'txt1(1).Enabled = False
              'txt1(2).Enabled = False
              'txt1(3).Enabled = False
              Option1(1).Value = False
              'txt1(4).Enabled = False
              'txt1(5).Enabled = True
              txt1(5).SetFocus
               txt1_GotFocus (5)
           End If
      Case Else
   End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 7 '查詢內容
          If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 Then
              KeyAscii = 0
          End If
   End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   'Add By Cheng 2002/01/07
   Select Case Index
      Case 6
         'Modify By Cheng 2002/03/14
      '   Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
      Case 7
            If InStr(1, "1234", UCase(txt1(Index))) = 0 Then
               s = MsgBox("請輸入 1 或 2 或 3 或 4 !!!", , "輸入錯誤")
               txt1(Index).SetFocus
               txt1(Index).SelStart = 0
               txt1(Index).SelLength = Len(txt1(Index))
               Exit Sub
            End If
   End Select
End Sub

Private Sub txt1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Select Case Index
      Case 0, 1, 2, 3
          Option1(0).Value = True
      Case 4
          Option1(1).Value = True
      Case 5
          Option1(2).Value = True
      Case Else
   End Select

End Sub
