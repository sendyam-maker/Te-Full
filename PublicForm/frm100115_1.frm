VERSION 5.00
Begin VB.Form frm100115_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "以國籍查詢代理人/申請人"
   ClientHeight    =   2055
   ClientLeft      =   10605
   ClientTop       =   1650
   ClientWidth     =   4200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4200
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2610
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   48
      Width           =   765
   End
   Begin VB.CommandButton cmdGoInput 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3405
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   48
      Width           =   765
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   1032
      MaxLength       =   7
      TabIndex        =   2
      Top             =   900
      Width           =   972
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   2352
      MaxLength       =   7
      TabIndex        =   3
      Top             =   900
      Width           =   972
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   696
      MaxLength       =   3
      TabIndex        =   0
      Top             =   528
      Width           =   972
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   2016
      MaxLength       =   3
      TabIndex        =   1
      Top             =   528
      Width           =   972
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   816
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1260
      Width           =   285
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "注意!!   有輸入 ""往來日期"" 區間時會很久 "
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   45
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   3240
   End
   Begin VB.Line Line2 
      X1              =   2115
      X2              =   2235
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Line Line1 
      X1              =   1776
      X2              =   1896
      Y1              =   648
      Y2              =   648
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "往來日期："
      Height          =   180
      Left            =   45
      TabIndex        =   10
      Top             =   945
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國籍："
      Height          =   180
      Left            =   45
      TabIndex        =   9
      Top             =   570
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "查詢別："
      Height          =   180
      Left            =   45
      TabIndex        =   8
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label3 
      Caption         =   "(1.代理人 2.申請人)"
      Height          =   180
      Left            =   1200
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "frm100115_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/12/29 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim s As Integer, i As Integer, j As Integer, strTemp As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

'92.04.16 nick
Public Sub PubShowNextData()
Select Case cmdState
Case 0
     cmdState = -1
      If Len(Trim(txt1(1))) = 0 Then
          s = MsgBox("國籍不可空白", , "USER 輸入錯誤")
          If Len(Trim(txt1(1))) = 0 Then txt1(1).SetFocus
          If Len(Trim(txt1(0))) = 0 Then txt1(0).SetFocus
          Exit Sub
      Else
          If Len(Trim(txt1(4))) = 0 Then
              s = MsgBox("查詢別不可空白", , "USER 輸入錯誤")
              txt1(4).SetFocus
              Exit Sub
          End If
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
      frm100115_2.Show
      If txt1(4) = "1" Then
         pub_QL05 = pub_QL05 & ";" & Label2 & "代理人" 'Add By Sindy 2010/11/4
         frm100115_2.StrMenu
      Else
         pub_QL05 = pub_QL05 & ";" & Label2 & "申請人" 'Add By Sindy 2010/11/4
         frm100115_2.StrMenu1
      End If
      Screen.MousePointer = vbDefault
      
      Me.Enabled = True
Case 1
     fnCloseAllFrm100
Case Else
End Select
End Sub
 
Private Sub cmdGoInput_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
'92.04.16 nick 以下無效
'Select Case Index
'Case 0
'      If Len(Trim(txt1(1))) = 0 Then
'          S = MsgBox("國籍不可空白", , "USER 輸入錯誤")
'          If Len(Trim(txt1(1))) = 0 Then txt1(1).SetFocus
'          If Len(Trim(txt1(0))) = 0 Then txt1(0).SetFocus
'          Exit Sub
'      Else
'          If Len(Trim(txt1(4))) = 0 Then
'              S = MsgBox("查詢別不可空白", , "USER 輸入錯誤")
'              txt1(4).SetFocus
'              Exit Sub
'          End If
'      End If
'      'Add By Cheng 2002/03/18
'      If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
'         Me.txt1(2).SetFocus
'         txt1_GotFocus 2
'         Exit Sub
'      End If
'      If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
'         Me.txt1(3).SetFocus
'         txt1_GotFocus 3
'         Exit Sub
'      End If
'
'      Screen.MousePointer = vbHourglass
'      frm100115_2.Show
'
'      If txt1(4) = "1" Then
'         frm100115_2.StrMenu
'      Else
'         frm100115_2.StrMenu1
'      End If
'      Screen.MousePointer = vbDefault
'      Me.Hide
'      'frm100115_2.Show
'      Do
'      DoEvents
'      If bolToEndByNick = True Then Unload Me: Exit Sub
'      Loop Until Not frm100115_2.Visible
'      Unload frm100115_2
'      'Me.Enabled = True
'      Me.Show
'Case 1
'      Unload Me
'Case Else
'End Select
End Sub

Private Sub Form_Load()
bolToEndByNick = False
   MoveFormToCenter Me
   bolToEndByNick = False
If bolFNation = False Then
    Label3.Caption = "(1.申請人)"
    txt1(4) = "1"
    txt1(4).Enabled = False
    Me.Caption = "以國籍查詢申請人"
Else
    Label3.Caption = "(1.代理人 2.申請人)"
End If
'92.04.16 nick
cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100115_1 = Nothing
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
Case 1
      If RunNick(txt1(Index - 1), txt1(Index)) Then
          txt1(Index - 1).SetFocus
          txt1_GotFocus (Index - 1)
          Exit Sub
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
Case 4
     If InStr(1, "12 ", txt1(4)) = 0 Then
         s = MsgBox("查詢別只可 1 或 2 !!", , "USER 輸入錯誤")
         txt1(4).SetFocus
         txt1(4).SelStart = 0
         txt1(4).SelLength = Len(txt1(4))
         Exit Sub
     End If
Case Else
End Select
End Sub
