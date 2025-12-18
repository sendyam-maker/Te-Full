VERSION 5.00
Begin VB.Form frm040204 
   BorderStyle     =   1  '單線固定
   Caption         =   "審查委員准駁統計"
   ClientHeight    =   1350
   ClientLeft      =   1200
   ClientTop       =   3900
   ClientWidth     =   4110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4110
   Begin VB.CommandButton Cmdok 
      Caption         =   "結束(&X)"
      Height          =   350
      Index           =   1
      Left            =   3204
      TabIndex        =   6
      Top             =   10
      Width           =   800
   End
   Begin VB.CommandButton Cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   2376
      TabIndex        =   5
      Top             =   10
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2940
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1068
      Width           =   1125
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1332
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1068
      Width           =   1125
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2940
      MaxLength       =   7
      TabIndex        =   2
      Top             =   756
      Width           =   1125
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1332
      MaxLength       =   7
      TabIndex        =   1
      Top             =   756
      Width           =   1125
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1332
      TabIndex        =   0
      Top             =   468
      Width           =   2712
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2604
      X2              =   2814
      Y1              =   1212
      Y2              =   1212
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2604
      X2              =   2814
      Y1              =   876
      Y2              =   876
   End
   Begin VB.Label lbl1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   9
      Top             =   1116
      Width           =   1116
   End
   Begin VB.Label lbl1 
      Caption         =   "准駁日："
      Height          =   180
      Index           =   1
      Left            =   168
      TabIndex        =   8
      Top             =   816
      Width           =   1116
   End
   Begin VB.Label lbl1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   132
      TabIndex        =   7
      Top             =   504
      Width           =   1116
   End
End
Attribute VB_Name = "frm040204"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/05 Form2.0已修改 (無需修改)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim strTemp As Variant, strTemp1 As Variant, i As Integer, j As Integer, s As Integer

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0 '確定

     If Len(txt1(0)) = 0 Then
        s = MsgBox("系統類別不可空白", , "USER 輸入錯誤")
        txt1(0).SetFocus
        Exit Sub
     Else
         'Add By Cheng 2002/03/19
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
        
        If Len(txt1(2)) = 0 Then
            s = MsgBox("准駁日區間不可空白", , "USER 輸入錯誤")
            txt1(1).SetFocus
            txt1_GotFocus (1)
            Exit Sub
        End If
     End If
     Me.Hide
     ClearQueryLog (Me.Name) 'Add By Sindy 2010/9/28 清除查詢印表記錄檔欄位
     Screen.MousePointer = vbHourglass
     frm040204a.Show
     Screen.MousePointer = vbDefault
Case 1
     Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = StrStartSystemByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm040204 = Nothing
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
     If Len(Trim(txt1(0))) <> 0 Then
            strTemp = Split("T,FCT,CFT,TF,CFP,P,FCP", ",")
            strTemp1 = Split(txt1(0), ",")
            For i = 0 To UBound(strTemp1)
                s = 0
                For j = 0 To UBound(strTemp)
                    If strTemp1(i) = strTemp(j) Then
                        s = 1
                    End If
                Next j
                If s = 0 Then
                    s = MsgBox(strUserNum + " 沒有 " + strTemp1(i) + " 的使用權限 ", , "USER 權限不足!!!")
                    txt1(0).SetFocus
                    txt1_GotFocus (0)
                    Exit Sub
                End If
            Next i
        End If
Case 1, 2
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
   If Index = 2 Then
        If Not nickChgRan(txt1(1), txt1(2), "准駁日") Then
           txt1(1).SetFocus
           txt1_GotFocus 1
           Exit Sub
        End If
    End If
Case 4
    If Not nickChgRan(txt1(3), txt1(4), "案件性質") Then
       txt1(3).SetFocus
       txt1_GotFocus 3
       Exit Sub
    End If
Case Else
End Select
End Sub


