VERSION 5.00
Begin VB.Form frm06010617 
   BorderStyle     =   1  '單線固定
   Caption         =   "信件沖銷記錄統計(外專)"
   ClientHeight    =   3360
   ClientLeft      =   290
   ClientTop       =   1750
   ClientWidth     =   6290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6290
   Begin VB.ComboBox cboType 
      Height          =   260
      Index           =   1
      ItemData        =   "frm06010617.frx":0000
      Left            =   3600
      List            =   "frm06010617.frx":0002
      TabIndex        =   3
      Text            =   "cboType"
      Top             =   1170
      Width           =   1790
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   2220
      MaxLength       =   3
      TabIndex        =   8
      Top             =   2190
      Width           =   630
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1455
      MaxLength       =   3
      TabIndex        =   7
      Top             =   2190
      Width           =   630
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   5
      Left            =   1455
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1860
      Width           =   890
   End
   Begin VB.ComboBox cboType 
      Height          =   260
      Index           =   0
      ItemData        =   "frm06010617.frx":0004
      Left            =   1455
      List            =   "frm06010617.frx":0006
      TabIndex        =   2
      Text            =   "cboType"
      Top             =   1170
      Width           =   1790
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   3590
      MaxLength       =   100
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   1790
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   3075
      MaxLength       =   7
      TabIndex        =   1
      Top             =   810
      Width           =   1245
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4440
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   180
      Width           =   756
   End
   Begin VB.CommandButton CmdOk 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5280
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   180
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1455
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2520
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1455
      MaxLength       =   7
      TabIndex        =   0
      Top             =   810
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   1455
      MaxLength       =   100
      TabIndex        =   4
      Top             =   1530
      Width           =   1790
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "註：外專信件沖銷啟用日 = 20220810"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   390
      TabIndex        =   20
      Top             =   3030
      Width           =   3720
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   3390
      X2              =   3510
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Label Label8 
      Caption         =   "部  門  別："
      Height          =   180
      Left            =   390
      TabIndex        =   19
      Top             =   2240
      Width           =   1050
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   1995
      X2              =   2340
      Y1              =   2325
      Y2              =   2325
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   3
      Left            =   2400
      TabIndex        =   18
      Top             =   1920
      Width           =   2030
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "承辦同仁："
      Height          =   180
      Left            =   390
      TabIndex        =   17
      Top             =   1920
      Width           =   1020
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "沖銷類別："
      Height          =   180
      Left            =   390
      TabIndex        =   16
      Top             =   1230
      Width           =   1020
   End
   Begin VB.Line Line3 
      Index           =   3
      Visible         =   0   'False
      X1              =   3380
      X2              =   3500
      Y1              =   1790
      Y2              =   1790
   End
   Begin VB.Line Line5 
      X1              =   2850
      X2              =   2970
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "(1.承辦同仁  2.沖銷類別  3.來函對象)"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   2010
      TabIndex        =   15
      Top             =   2580
      Width           =   2960
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "來函對象："
      Height          =   180
      Left            =   390
      TabIndex        =   14
      Top             =   1590
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "轉入日期：                                                                      輸入民國年"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   390
      TabIndex        =   13
      Top             =   870
      Width           =   5300
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "統計條件："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   390
      TabIndex        =   12
      Top             =   2580
      Width           =   1020
   End
End
Attribute VB_Name = "frm06010617"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Sindy 2024/4/25
Option Explicit

Dim s As Integer
'紀錄作用按鍵
Public cmdState As Integer


Public Sub PubShowNextData()
Dim bolCancel As Boolean
   
Select Case cmdState
Case 0 '確定
   Dim oText As TextBox
   For Each oText In txt1
      Call txt1_Validate(oText.Index, bolCancel)
      If bolCancel = True Then
         Exit Sub
      End If
   Next
   
    cmdState = -1
    If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
        txt1_GotFocus 1
        Exit Sub
    End If
    If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
        txt1_GotFocus 2
        Exit Sub
    End If
    
    If txt1(1) = "" Then
        MsgBox "起始日期不可空白 !", vbCritical
        txt1(1).SetFocus
        Exit Sub
    End If
    If txt1(2) = "" Then
        MsgBox "迄止日期不可空白 !", vbCritical
        txt1(2).SetFocus
        Exit Sub
    End If
    If txt1(6) = "" Then
        MsgBox "統計條件不可空白 !", vbCritical
        txt1(6).SetFocus
        Exit Sub
    End If
    'Added by Lydia 2023/12/28
    If DBDATE(txt1(2)) >= 新部門啟用日 And DBDATE(txt1(1)) < 新部門啟用日 And txt1(6) = "1" Then
       MsgBox "轉入日期不可跨過" & TransDate(新部門啟用日, 1) & "！"
       txt1(1).SetFocus
       txt1_GotFocus 1
       Exit Sub
    End If
    'end 2023/12/28
    
    Me.Enabled = False
    If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ClearQueryLog (Me.Name) '清除查詢印表記錄檔欄位
    frm06010617_1.Show
    frm06010617_1.StrMenu
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    
Case 1 '結束
     'fnCloseAllFrm100
     Unload Me
Case Else
End Select
End Sub

Private Sub cboType_Click(Index As Integer)
   Dim iPos As Integer
   iPos = InStr(cboType(Index).Text, Chr(1))
   If iPos > 0 Then
      cboType(Index).Text = Left(cboType(Index).Text, iPos - 1)
   End If
End Sub

Private Sub cboType_GotFocus(Index As Integer)
   If cboType(Index).Locked = False Then
      CloseIme
      SendMessage cboType(Index).hWnd, CB_SHOWDROPDOWN, 1, 0
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
'紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   bolToEndByNick = False
   
   cmdState = -1
   
   Call SetcboType(0)
   Call SetcboType(1)
End Sub

'沖銷類別
Private Sub SetcboType(Index As Integer)
   cboType(Index).Clear
   cboType(Index).AddItem ""
   If Pub_StrUserSt03 = "F22" Then '程序組
      cboType(Index).AddItem "1 輸入"
      cboType(Index).AddItem "8 不處理"
      cboType(Index).AddItem "9 已處理"
      cboType(Index).AddItem "10 回信"
   Else
      cboType(Index).AddItem "1 舊案收文－自沖"
      cboType(Index).AddItem "2 多案收文－自沖"
      cboType(Index).AddItem "3 新案命名－自沖"
      cboType(Index).AddItem "4 客戶提供文件－自沖"
      cboType(Index).AddItem "5 不續辦或閉卷－主核"
      cboType(Index).AddItem "6 往來記錄－主核"
      cboType(Index).AddItem "7 承辦作業－自沖"
      cboType(Index).AddItem "8 不處理－主核"
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm06010617 = Nothing
End Sub

Private Sub txt1_Change(Index As Integer)
   Select Case Index
      Case 5 '承辦人
         Me.lbl1(3).Caption = StaffQuery("" & Me.txt1(Index).Text)
   End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 5, 7, 8
         KeyAscii = UpperCase(KeyAscii)
      Case 1, 2, 6
         KeyAscii = Pub_NumAscii(KeyAscii)
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
       Case 1, 2
            If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
               Me.txt1(Index).SetFocus
               txt1_GotFocus Index
               Cancel = True
               Exit Sub
            End If
            If Index = 2 Then
               If RunNick(txt1(Index - 1), txt1(Index)) Then
                   txt1(Index - 1).SetFocus
                   txt1_GotFocus (Index - 1)
                   Cancel = True
                   Exit Sub
               End If
            End If
       Case 8
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               txt1(Index - 1).SetFocus
               txt1_GotFocus (Index - 1)
               Cancel = True
               Exit Sub
            End If
       Case 5 '承辦同仁
            lbl1(3) = GetPrjSalesNM(txt1(Index))
            If Trim(txt1(Index)) <> "" Then
               If Trim(lbl1(3)) = "" Then
                  s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
                  txt1(Index).SetFocus
                  txt1_GotFocus (Index)
                  Cancel = True
                  Exit Sub
               End If
            End If
       Case 6
            If InStr(1, "123 ", txt1(Index)) = 0 Then
                s = MsgBox("請輸入 1 或 2 或 3 !!", , "輸入錯誤")
                txt1(Index).SetFocus
                txt1(Index).SelStart = 0
                txt1(Index).SelLength = Len(txt1(Index))
                Exit Sub
             End If
       Case Else
   End Select
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub
