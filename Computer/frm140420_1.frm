VERSION 5.00
Begin VB.Form frm140420_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "往來記錄統計"
   ClientHeight    =   3084
   ClientLeft      =   288
   ClientTop       =   1752
   ClientWidth     =   6288
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3084
   ScaleWidth      =   6288
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
      Height          =   264
      Index           =   5
      Left            =   1455
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1860
      Width           =   1245
   End
   Begin VB.ComboBox cboSort 
      Height          =   300
      Index           =   1
      ItemData        =   "frm140420_1.frx":0000
      Left            =   3075
      List            =   "frm140420_1.frx":0002
      TabIndex        =   3
      Text            =   "cboSort"
      Top             =   1170
      Width           =   1320
   End
   Begin VB.ComboBox cboSort 
      Height          =   300
      Index           =   0
      ItemData        =   "frm140420_1.frx":0004
      Left            =   1455
      List            =   "frm140420_1.frx":0006
      TabIndex        =   2
      Text            =   "cboSort"
      Top             =   1170
      Width           =   1320
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   3075
      MaxLength       =   9
      TabIndex        =   5
      Top             =   1530
      Width           =   1245
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
      Height          =   264
      Index           =   3
      Left            =   1455
      MaxLength       =   9
      TabIndex        =   4
      Top             =   1530
      Width           =   1245
   End
   Begin VB.Label Label8 
      Caption         =   "業務區："
      Height          =   180
      Left            =   570
      TabIndex        =   19
      Top             =   2235
      Width           =   720
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
      Left            =   2730
      TabIndex        =   18
      Top             =   1920
      Width           =   1605
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "接洽同仁："
      Height          =   180
      Left            =   390
      TabIndex        =   17
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "往來類別："
      Height          =   180
      Left            =   390
      TabIndex        =   16
      Top             =   1230
      Width           =   900
   End
   Begin VB.Line Line3 
      Index           =   3
      X1              =   2865
      X2              =   2985
      Y1              =   1635
      Y2              =   1635
   End
   Begin VB.Line Line5 
      X1              =   2850
      X2              =   2970
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "(1.接洽同仁  2.往來類別  3.往來對象)"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   2010
      TabIndex        =   15
      Top             =   2580
      Width           =   3585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "往來對象："
      Height          =   180
      Left            =   390
      TabIndex        =   14
      Top             =   1590
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "往來日期：                                                                      輸入民國年"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   390
      TabIndex        =   13
      Top             =   870
      Width           =   4950
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "統計條件："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   390
      TabIndex        =   12
      Top             =   2580
      Width           =   900
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   2865
      X2              =   2985
      Y1              =   1275
      Y2              =   1275
   End
End
Attribute VB_Name = "frm140420_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Sindy 2019/12/27
Option Explicit

Dim s As Integer
'紀錄作用按鍵
Public cmdState As Integer


Public Sub PubShowNextData()
Dim bolCancel As Boolean
   
Select Case cmdState
Case 0 '確定
   Dim oText As TextBox
   For Each oText In Txt1
      Call txt1_Validate(oText.Index, bolCancel)
      If bolCancel = True Then
         Exit Sub
      End If
   Next
   
    cmdState = -1
    If PUB_CheckKeyInDate(Me.Txt1(1)) = -1 Then
        txt1_GotFocus 1
        Exit Sub
    End If
    If PUB_CheckKeyInDate(Me.Txt1(2)) = -1 Then
        txt1_GotFocus 2
        Exit Sub
    End If
    
    If Txt1(1) = "" Then
        MsgBox "起始日期不可空白 !", vbCritical
        Txt1(1).SetFocus
        Exit Sub
    End If
    If Txt1(2) = "" Then
        MsgBox "迄止日期不可空白 !", vbCritical
        Txt1(2).SetFocus
        Exit Sub
    End If
    If Txt1(6) = "" Then
        MsgBox "統計條件不可空白 !", vbCritical
        Txt1(6).SetFocus
        Exit Sub
    End If
    'Added by Lydia 2023/12/28
    If DBDATE(Txt1(2)) >= 新部門啟用日 And DBDATE(Txt1(1)) < 新部門啟用日 And Txt1(6) = "1" Then
       MsgBox "往來日期不可跨過" & TransDate(新部門啟用日, 1) & "！"
       Txt1(1).SetFocus
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
    frm140420_2.Show
    frm140420_2.StrMenu
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    
Case 1 '結束
     'fnCloseAllFrm100
     Unload Me
Case Else
End Select
End Sub

Private Sub cboSort_Click(Index As Integer)
   Dim iPos As Integer
   iPos = InStr(cboSort(Index).Text, Chr(1))
   If iPos > 0 Then
      cboSort(Index).Text = Left(cboSort(Index).Text, iPos - 1)
   End If
End Sub

Private Sub cboSort_GotFocus(Index As Integer)
   If cboSort(Index).Locked = False Then
      CloseIme
      SendMessage cboSort(Index).hWnd, CB_SHOWDROPDOWN, 1, 0
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
   
   Call SetcboSort(0)
   Call SetcboSort(1)
End Sub

'往來類別
Private Sub SetcboSort(Index As Integer)
   cboSort(Index).Clear
   strSql = "select ac02,ac03 from allcode where ac01='11'" & _
            " order by ac02 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   cboSort(Index).AddItem ""
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         cboSort(Index).AddItem RsTemp.Fields("ac02") & " " & RsTemp.Fields("ac03")
         RsTemp.MoveNext
      Loop
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm140420_1 = Nothing
End Sub

Private Sub txt1_Change(Index As Integer)
   Select Case Index
      Case 5 '承辦人
         Me.lbl1(3).Caption = StaffQuery("" & Me.Txt1(Index).Text)
   End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   Txt1(Index).SelStart = 0
   Txt1(Index).SelLength = Len(Txt1(Index))
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 3, 4, 5, 7, 8
         KeyAscii = UpperCase(KeyAscii)
      Case 1, 2, 6
         KeyAscii = Pub_NumAscii(KeyAscii)
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
       Case 1, 2
            If PUB_CheckKeyInDate(Me.Txt1(Index)) = -1 Then
               Me.Txt1(Index).SetFocus
               txt1_GotFocus Index
               Cancel = True
               Exit Sub
            End If
            If Index = 2 Then
               If RunNick(Txt1(Index - 1), Txt1(Index)) Then
                   Txt1(Index - 1).SetFocus
                   txt1_GotFocus (Index - 1)
                   Cancel = True
                   Exit Sub
               End If
            End If
       Case 8
            If RunNick(Txt1(Index - 1), Txt1(Index)) Then
               Txt1(Index - 1).SetFocus
               txt1_GotFocus (Index - 1)
               Cancel = True
               Exit Sub
            End If
       Case 5 '接洽同仁
            lbl1(3) = GetPrjSalesNM(Txt1(Index))
            If Trim(Txt1(Index)) <> "" Then
               If Trim(lbl1(3)) = "" Then
                  s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
                  Txt1(Index).SetFocus
                  txt1_GotFocus (Index)
                  Cancel = True
                  Exit Sub
               End If
            End If
       Case 4
            If Txt1(3) <> "" Or Txt1(Index) <> "" Then
               If RunNick(Txt1(3), Txt1(Index)) Then
                  Txt1(3).SetFocus
                  txt1_GotFocus 3
                  Cancel = True
                  Exit Sub
               End If
               If Mid(Txt1(3), 1, 6) <> Mid(Txt1(Index), 1, 6) Then
                  MsgBox "往來對象前六碼必須相同！", , "發生錯誤！"
                  Txt1(3).SetFocus
                  txt1_GotFocus 3
                  Cancel = True
                  Exit Sub
               End If
            End If
       Case 6
            If InStr(1, "123 ", Txt1(Index)) = 0 Then
                s = MsgBox("請輸入 1 或 2 或 3 !!", , "輸入錯誤")
                Txt1(Index).SetFocus
                Txt1(Index).SelStart = 0
                Txt1(Index).SelLength = Len(Txt1(Index))
                Exit Sub
             End If
       Case Else
   End Select
   Txt1(Index).SelStart = 0
   Txt1(Index).SelLength = Len(Txt1(Index))
End Sub
