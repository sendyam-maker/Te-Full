VERSION 5.00
Begin VB.Form frm04010703 
   BorderStyle     =   1  '單線固定
   Caption         =   "重新委任發文(不印申請書)"
   ClientHeight    =   2970
   ClientLeft      =   255
   ClientTop       =   990
   ClientWidth     =   9225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   9225
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1875
      MaxLength       =   6
      TabIndex        =   0
      Top             =   720
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   1
      Top             =   720
      Width           =   210
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   2
      Top             =   720
      Width           =   315
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1080
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1920
      Width           =   8040
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8220
      TabIndex        =   5
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7410
      TabIndex        =   4
      Top             =   60
      Width           =   800
   End
   Begin VB.TextBox txtCP27 
      Height          =   270
      Left            =   1080
      MaxLength       =   7
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Caption         =   "lblCount"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   1320
      TabIndex        =   20
      Top             =   240
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "本次筆數："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   240
      Width           =   990
   End
   Begin VB.Line Line4 
      X1              =   1680
      X2              =   3540
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   1140
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "中："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   1095
      TabIndex        =   16
      Top             =   1125
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "英："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   1095
      TabIndex        =   15
      Top             =   1350
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "日："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   1095
      TabIndex        =   14
      Top             =   1575
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "申請人："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   780
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   0
      Left            =   1515
      TabIndex        =   12
      Top             =   1140
      Width           =   7620
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   1
      Left            =   1530
      TabIndex        =   11
      Top             =   1350
      Width           =   7620
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   2
      Left            =   1485
      TabIndex        =   10
      Top             =   1575
      Width           =   7620
   End
   Begin VB.Label lblClose 
      Caption         =   "lblClose"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3840
      TabIndex        =   9
      Top             =   780
      Width           =   990
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "發文日期："
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   900
   End
End
Attribute VB_Name = "frm04010703"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan2010/8/11 日期欄已修改
'2007/7/3 ADD BY SONIA
Option Explicit
Dim m_Count As Integer
Dim bolResult As Boolean

Private Sub cmdok_Click(Index As Integer)
   Dim strAppDate As String, iRecs As Integer
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 0 '確定
         If TxtValidate = True Then
            strExc(0) = "是否確定要發文此筆重新委任？"
            '預設要新增
            If MsgBox(strExc(0), vbYesNo + vbDefaultButton1) = vbYes Then
               If FormSave = False Then
                  MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
               Else
                  FormClear
               End If
            End If
            txt1_GotFocus (2)
            txt1(2).SetFocus
         End If
      Case 1 '結束
         Unload Me
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Me.Top = Me.Top - 1354
   FormClear
   txtCP27 = ""
   txt1(1) = "P"
   lblCount = "": m_Count = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010703 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   '游標停在"0"的後面
   txt1(2).SelStart = 1
   txt1(2).SelLength = Len(txt1(2).Text)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Select Case Index
      Case 3
         If txt1(3) = "" Then txt1(3) = "0"
      Case 4
         If txt1(4) = "" Then txt1(4) = "00"
         bolResult = False
         Check928
         If bolResult = False Then
            txt1_GotFocus (2)
            txt1(2).SetFocus
         End If
   End Select
   
End Sub

Private Sub txtCP27_Validate(Cancel As Boolean)
   If txtCP27.Text <> "" Then
      If Not ChkDate(txtCP27.Text) Then
         MsgBox "發文日期不正確，請重新輸入 !", vbCritical
         Cancel = True
      Else
         If Val(txtCP27.Text) < Val(strSrvDate(2)) And Val(txtCP27.Text) <> "111111" Then
            MsgBox "發文日期不可小於系統日，請重新輸入 !", vbCritical
            Cancel = True
         End If
      End If
   Else
      MsgBox "發文日期不可空白 !", vbCritical
      Cancel = True
   End If
   If Cancel = True Then TextInverse txtCP27
End Sub

Private Sub txtCP27_GotFocus()
  'edit by nickc 2007/07/11 切換輸入法改用API
  'txtCP27.IMEMode = 2
  CloseIme
  TextInverse txtCP27
End Sub
   
Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
   
   lbl1(0) = "": lbl1(1) = "": lbl1(2) = ""
   Combo1.Clear
   Me.lblClose.Caption = ""
   
   If txt1(2) = "" Then
      MsgBox "本所案號不可空白！", vbExclamation
      txt1(2).SetFocus
      Exit Function
   End If
   
   If txtCP27 = "" Then
      MsgBox "發文日期不可空白！", vbExclamation
      txtCP27.SetFocus
      Exit Function
   Else
      Cancel = False
      txtCP27_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   bolResult = False
   If txt1(3) = "" Then txt1(3) = "0"
   If txt1(4) = "" Then txt1(4) = "00"
   
   Check928
   If bolResult = False Then
      txt1_GotFocus (2)
   Else
      TxtValidate = True
   End If

End Function

Private Function Check928() As Boolean
   strExc(0) = "select pa05,pa06,pa07,'中 '||cu04,'英 '||cu05||' '||cu88||' '||cu89||' '||cu90,'日 '||cu06,'錯誤!! '||PA26,PA57,CP27,CP57 from patent,customer,CASEPROGRESS where pa01='" & txt1(1) & "' and pa02='" & txt1(2) & "' and pa03='" & txt1(3) & "' and pa04='" & txt1(4) & "' and pa09='000' and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),'','0',substr(pa26,9,1))=cu02(+) AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND '928'=CP10(+) AND CP57 IS NULL and cp09 is not null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         If lbl1(0) = "" Then
            lbl1(0) = "" & .Fields("pa05")
         End If
         If lbl1(1) = "" Then
            lbl1(1) = "" & .Fields("pa06")
         End If
         If lbl1(2) = "" Then
            lbl1(2) = "" & .Fields("pa07")
         End If
         
         If Len(CheckStr(.Fields(3))) = 0 And Len(CheckStr(.Fields(4))) = 0 And Len(CheckStr(.Fields(5))) = 0 Then
            Combo1.AddItem CheckStr(.Fields(6)), 0
         Else
            If Len(CheckStr(.Fields(3))) = 0 Then
               Combo1.AddItem CheckStr(.Fields(6)), 0
            Else
               Combo1.AddItem CheckStr(.Fields(3)), 0
            End If
            If Len(CheckStr(.Fields(4))) = 0 Then
               Combo1.AddItem CheckStr(.Fields(6)), 1
            Else
               Combo1.AddItem CheckStr(.Fields(4)), 1
            End If
            If Len(CheckStr(.Fields(5))) = 0 Then
               Combo1.AddItem CheckStr(.Fields(6)), 2
            Else
               Combo1.AddItem CheckStr(.Fields(5)), 2
            End If
         End If
         Combo1.Text = Combo1.List(0)
         
         If IsNull(.Fields("CP27")) = False Then
            MsgBox "此案件已發文 !", vbCritical
            Exit Function
         End If
         
         bolResult = True
         
         If IsNull(.Fields("pa57")) = False Then
            lblClose.Caption = "已閉卷"
            MsgBox "此案件已閉卷 !", vbCritical
         End If
      End With
   Else
      MsgBox "本所案號輸入錯誤！"
   End If
   
End Function
Private Function FormSave() As Boolean
   
On Error GoTo ErrHnd
   
   strSql = "UPDATE caseprogress SET cp27=" & DBDATE(txtCP27) & " WHERE cp01='" & txt1(1) & "' and cp02='" & txt1(2) & "' and cp03='" & txt1(3) & "' and cp04='" & txt1(4) & "' AND '928'=CP10 AND CP27 IS NULL AND CP57 IS NULL "
   cnnConnection.Execute strSql, intI
   FormSave = True
   m_Count = m_Count + 1
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function

Private Sub FormClear()
Dim intPos As Integer
   
   txt1(2) = "0": txt1(3) = "": txt1(4) = ""
   lbl1(0) = "": lbl1(1) = "": lbl1(2) = ""
   Combo1.Clear
   Me.lblClose.Caption = ""
   lblCount = m_Count

   If m_Count = 0 Then
      txtCP27.Locked = False
      txtCP27.Enabled = True
   Else
      txtCP27.Locked = True
      txtCP27.Enabled = False
   End If

End Sub


