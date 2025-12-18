VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090103_1 
   Caption         =   "查名期限資料查詢"
   ClientHeight    =   2895
   ClientLeft      =   1665
   ClientTop       =   2025
   ClientWidth     =   6075
   ControlBox      =   0   'False
   LinkTopic       =   "frm0903_1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2895
   ScaleWidth      =   6075
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   4992
      TabIndex        =   7
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   4164
      TabIndex        =   6
      Top             =   120
      Width           =   800
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   0
      Left            =   2064
      MaxLength       =   6
      TabIndex        =   0
      Top             =   1080
      Width           =   825
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   2
      Left            =   3228
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1716
      Width           =   825
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   1
      Left            =   2076
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1692
      Width           =   825
   End
   Begin MSForms.Label LblTmq10NM 
      Height          =   255
      Left            =   2940
      TabIndex        =   8
      Top             =   1110
      Width           =   1875
      VariousPropertyBits=   27
      Size            =   "3307;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "－"
      Height          =   288
      Left            =   2976
      TabIndex        =   5
      Top             =   1764
      Width           =   276
   End
   Begin VB.Label Label2 
      Caption         =   "期限日期："
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1152
      TabIndex        =   4
      Top             =   1752
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "查名人："
      Height          =   300
      Left            =   1152
      TabIndex        =   3
      Top             =   1104
      Width           =   900
   End
End
Attribute VB_Name = "frm090103_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/21 改成Form2.0 ; LblTmq10NM ; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim s As Integer, i As Integer, j As Integer, strSql As String
Dim BlnCheck As Boolean

Private Sub cmdOK_Click()
   Txtdata_LostFocus (0)
   If BlnCheck Then Exit Sub
   Txtdata_LostFocus (2)
   If BlnCheck Then Exit Sub
   'Add By Cheng 2002/03/21
   If PUB_CheckKeyInDate(Me.Txtdata(1)) = -1 Then
      Me.Txtdata(1).SetFocus
      Txtdata_GotFocus 1
      Exit Sub
   End If
   If PUB_CheckKeyInDate(Me.Txtdata(2)) = -1 Then
      Me.Txtdata(2).SetFocus
      Txtdata_GotFocus 2
      Exit Sub
   End If
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/13 清除查詢印表記錄檔欄位
   Me.Enabled = False
   Me.Hide
   frm090103_2.Show
   frm090103_2.Hide
   frm090103_2.MousePointer = vbHourglass
   frm090103_2.GridData
   frm090103_2.MousePointer = vbDefault
   If frm090103_2.Enabled = True Then
      frm090103_2.Show
   Else
      s = MsgBox("資料庫中沒有符合的資料!!", , "請檢查條件")
   End If
   Do
      DoEvents
      If bolToEndByNick = True Then
         cmdExit_Click
         Exit Sub
      End If
   Loop Until Not frm090103_2.Visible
   Unload frm090103_2
   Me.Enabled = True
   Me.Show
End Sub

Private Sub cmdExit_Click()
   Unload Me
   Set frm090103_2 = Nothing
End Sub

Private Sub Form_Load()
   Me.Height = 3915
   Me.Width = 6195
   MoveFormToCenter Me
   BlnCheck = False: bolToEndByNick = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm090103_1 = Nothing
End Sub

Private Sub Txtdata_GotFocus(Index As Integer)
   TextInverse Txtdata(Index)
End Sub

Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Txtdata_LostFocus(Index As Integer)
Dim strTemp As String
Dim strTemp1 As String
   BlnCheck = False
   Select Case Index
   Case 0
      If Txtdata(0) <> Empty Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(txtData(0).Text, strTemp, strTemp1) Then
         'Modified by Lydia 2018/07/02
         'If ClsPDGetStaff(Txtdata(0).Text, strTemp, strTemp1) Then
         strTemp = GetStaffName(Txtdata(0).Text, True)
         If strTemp <> "" Then
         'end 2018/07/02
            LblTmq10NM.Caption = strTemp
         Else
            Txtdata(0).SetFocus
            TextInverse Txtdata(0)
            BlnCheck = True
         Exit Sub
         End If
      Else
         LblTmq10NM.Caption = ""
      End If
   Case 2
      If Txtdata(1) = Empty And Txtdata(2) = Empty Then
         s = MsgBox("請輸入期限日期條件", , "使用者輸入錯誤")
         Txtdata(1).SetFocus
         BlnCheck = True
         Exit Sub
      End If
      'Modify by Morgan 2010/8/16 百年蟲
      'If Txtdata(1) > Txtdata(2) Then
      If Val(Txtdata(1)) > Val(Txtdata(2)) Then
         s = MsgBox("期限日期範圍錯誤", , "使用者輸入錯誤")
         Txtdata(1).SetFocus
         TextInverse Txtdata(1)
         BlnCheck = True
         Exit Sub
      End If
   Case Else
   End Select
End Sub

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 1, 2 '期限日期起, 迄
   If PUB_CheckKeyInDate(Me.Txtdata(Index)) = -1 Then
      Cancel = True
      Me.Txtdata(Index).SetFocus
      Txtdata_GotFocus Index
   End If
End Select
End Sub
