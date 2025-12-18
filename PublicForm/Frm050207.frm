VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050207 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工查詢印表記錄資料查詢"
   ClientHeight    =   3180
   ClientLeft      =   3045
   ClientTop       =   1515
   ClientWidth     =   5865
   ControlBox      =   0   'False
   LinkTopic       =   "Form11"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5865
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2580
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1620
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1620
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1440
      TabIndex        =   8
      Top             =   2310
      Width           =   2532
   End
   Begin VB.TextBox txtSystem 
      Height          =   264
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   9
      Top             =   2640
      Width           =   732
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   2
      Left            =   3870
      MaxLength       =   2
      TabIndex        =   12
      Top             =   2640
      Width           =   492
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   1
      Left            =   3450
      MaxLength       =   1
      TabIndex        =   11
      Top             =   2640
      Width           =   372
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   0
      Left            =   2205
      MaxLength       =   6
      TabIndex        =   10
      Top             =   2640
      Width           =   1212
   End
   Begin VB.CheckBox Check1 
      Caption         =   "含離職人員"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   900
      Width           =   1785
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1290
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2580
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1290
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   4905
      TabIndex        =   14
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4125
      TabIndex        =   13
      Top             =   60
      Width           =   756
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Index           =   2
      Left            =   1440
      TabIndex        =   7
      Top             =   1950
      Width           =   4185
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "7382;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   900
      Width           =   2235
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3942;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   540
      Width           =   2235
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3942;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   23
      Top             =   2700
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   2340
      X2              =   2620
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "操作時間："
      Height          =   180
      Left            =   480
      TabIndex        =   22
      Top             =   1650
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "系統類別：                                                          (ALL ：全部)"
      Height          =   180
      Left            =   480
      TabIndex        =   21
      Top             =   2355
      Width           =   4545
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "操作日期："
      Height          =   180
      Left            =   480
      TabIndex        =   20
      Top             =   1320
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2340
      X2              =   2620
      Y1              =   1410
      Y2              =   1410
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "操作程式："
      Height          =   180
      Index           =   0
      Left            =   480
      TabIndex        =   19
      Top             =   1980
      Width           =   900
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "部　　門： "
      Height          =   180
      Index           =   2
      Left            =   480
      TabIndex        =   18
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   3
      Left            =   480
      TabIndex        =   17
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "(1. 管制表 2. 定稿)"
      Height          =   180
      Left            =   5955
      TabIndex        =   16
      Top             =   5325
      Width           =   1425
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "列印格式:"
      Height          =   180
      Left            =   4170
      TabIndex        =   15
      Top             =   5340
      Width           =   765
   End
End
Attribute VB_Name = "frm050207"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ; Combo1(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/2 日期欄已修改
'2010/01/06 CREATE BY Sindy
Option Explicit

Dim s As Integer, i As Integer, j As Integer
Dim StrTag As String, strSql As String
Dim m_bln_KeyinValid As Boolean
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer


Public Sub PubShowNextData()
'If Len(Trim(Me.txt1(4).Text)) = 0 Then
'    Me.txt1(4).Text = "ALL"
'End If
Select Case cmdState
   Case 0
       cmdState = -1
       If Len(Trim(Me.Combo1(0).Text)) = 0 And _
         Len(Trim(Me.Combo1(1).Text)) = 0 And _
         Len(Trim(Me.Combo1(2).Text)) = 0 And _
         Len(Trim(Txt1(0).Text)) = 0 And _
         Len(Trim(Txt1(1).Text)) = 0 And _
         Len(Trim(Txt1(2).Text)) = 0 And _
         Len(Trim(Txt1(3).Text)) = 0 And _
         Len(Trim(Txt1(4).Text)) = 0 And _
         Len(Trim(txtSystem.Text)) = 0 And _
         Len(Trim(txtCode(0).Text)) = 0 And _
         Check1.Value = 0 Then
          s = MsgBox("請檢查是否有必要條件忘了輸入．．．．", , "輸入條件不足")
          Me.Combo1(0).SetFocus
          Exit Sub
       End If
       If Me.Txt1(0).Text = "" Then
         MsgBox "操作日起不可空白!!!", vbExclamation + vbOKOnly
         txt1_LostFocus 0
         Exit Sub
       End If
       If Me.Txt1(0).Text <> "" Then
         txt1_LostFocus 0
         If m_bln_KeyinValid = False Then txt1_GotFocus 0: Exit Sub
       End If
       If Me.Txt1(1).Text = "" Then
         MsgBox "操作日迄不可空白!!!", vbExclamation + vbOKOnly
         txt1_LostFocus 1
         Exit Sub
       End If
       If Me.Txt1(1).Text <> "" Then
         txt1_LostFocus 1
         If m_bln_KeyinValid = False Then txt1_GotFocus 1: Exit Sub
       End If
       If txtSystem <> "" And txtCode(0) <> "" Then
         If txtCode(1) = "" Then txtCode(1) = "0"
         If txtCode(2) = "" Then txtCode(2) = "00"
       End If
       
       Me.Enabled = False
       If fnSaveParentForm(Me) = False Then
           Me.Enabled = True
           Exit Sub
       End If
       Screen.MousePointer = vbHourglass
       frm050207_1.Show
       frm050207_1.StrMenu
       Screen.MousePointer = vbDefault
       Me.Enabled = True
   Case 1
        'fnCloseAllFrm100
        Unload Me
   Case Else
End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
End Sub

Private Sub Combo1_GotFocus(Index As Integer)
Combo1(Index).SelStart = 0
Combo1(Index).SelLength = Len(Combo1(Index))
Select Case Index
   Case 0
      Combo1(0).Tag = Combo1(0).Text
   Case 2 '操作程式
      OpenIme
   Case Else
End Select
End Sub

'Modified by Lydia 2021/09/22 改成Form 2.0
'Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
Select Case Index
   Case 0, 1
      KeyAscii = UpperCase(KeyAscii)
   Case Else
End Select
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
m_bln_KeyinValid = False
Select Case Index
   Case 0
      If Combo1(Index).Text <> "" Then
         SetComboData_Staff
      End If
   Case 2 '操作程式
      CloseIme
   Case Else
End Select
m_bln_KeyinValid = True
End Sub

'Add By Sindy 2012/7/31
Private Sub Combo1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
   Case 1
      If Combo1(Index).Text <> "" And Len(Trim(Combo1(Index).Text)) = 5 Then
         Combo1(Index).Text = Trim(Combo1(Index).Text) & "  " & GetPrjSalesNM(Combo1(Index).Text)
      End If
End Select
End Sub

Private Sub Form_Activate()
Me.Combo1(0).SetFocus
End Sub

Private Sub Form_Load()
bolToEndByNick = False
MoveFormToCenter Me
Me.Check1.Value = 0
'Me.Combo1(0).Text = ""
'Me.Combo1(1).Text = ""
Me.Combo1(2).Text = ""
Me.Txt1(0).Text = strSrvDate(2)
Me.Txt1(1).Text = strSrvDate(2)
SetComboData
SetComboData_Staff
'92.04.16 nick

'txt1(4) = Systemkind_g 'Modify By Sindy 2012/7/31 系統類別不預設欄位值

cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050207_1 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
Txt1(Index).SelStart = 0
Txt1(Index).SelLength = Len(Txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
   Case 0, 1, 2, 3
      KeyAscii = Pub_NumAscii(KeyAscii)
   Case 4
      KeyAscii = UpperCase(KeyAscii)
   Case Else
End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
m_bln_KeyinValid = False
Select Case Index
   Case 0 '操作日起
         If Me.Txt1(Index) <> "" Then
            If CheckIsTaiwanDate(Me.Txt1(Index)) = False Then
               Me.Txt1(Index).SetFocus
               Exit Sub
            End If
         End If
   Case 1 '操作日迄
         If Me.Txt1(Index) <> "" Then
            If CheckIsTaiwanDate(Me.Txt1(Index)) = False Then
               Me.Txt1(Index).SetFocus
               Exit Sub
            End If
            If Val(Me.Txt1(0).Text) > Val(Me.Txt1(1).Text) Then
               MsgBox "操作日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               Me.Txt1(0).SetFocus
               Exit Sub
            End If
         End If
   Case 3 '操作時間迄
         If Me.Txt1(Index) <> "" Then
            If Val(Me.Txt1(2).Text) > Val(Me.Txt1(3).Text) Then
               MsgBox "操作時間範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               Me.Txt1(2).SetFocus
               Exit Sub
            End If
         End If
   Case Else
End Select
m_bln_KeyinValid = True
End Sub

Private Sub SetComboData()
Dim rs As New ADODB.Recordset

'部門
Me.Combo1(0).Clear
rs.CursorLocation = adUseClient
rs.Open "Select A0901,A0902 From ACC090 Where a0904 <> 'Y' Order By A0901", _
         cnnConnection, adOpenStatic, adLockReadOnly
Me.Combo1(0).AddItem ""
While Not rs.EOF
   Me.Combo1(0).AddItem Left(rs.Fields(0).Value & Space(5), 5) & rs.Fields(1).Value
   rs.MoveNext
Wend
If rs.State <> adStateClosed Then rs.Close
Set rs = Nothing

'操作程式
Me.Combo1(2).Clear
rs.CursorLocation = adUseClient
rs.Open "select fo02 from form where fo01<>'frm050207' group by fo02 order by 1", _
         cnnConnection, adOpenStatic, adLockReadOnly
Me.Combo1(2).AddItem ""
While Not rs.EOF
   Me.Combo1(2).AddItem "" & rs.Fields(0).Value
   rs.MoveNext
Wend
If rs.State <> adStateClosed Then rs.Close
Set rs = Nothing
End Sub

Private Sub SetComboData_Staff()
Dim rs As New ADODB.Recordset
Dim strSql As String
Dim strTemp As String 'Add By Sindy 2012/7/31 記錄畫面上員工編號欄值

strTemp = Combo1(1).Text 'Add By Sindy 2012/7/31

strSql = ""
'員工編號
Me.Combo1(1).Clear
Me.Enabled = False
Screen.MousePointer = vbHourglass
rs.CursorLocation = adUseClient
If Check1.Value = 0 Then '不含離職人員
   strSql = strSql & "and st04='1'"
End If
If Trim(Left(Trim(Combo1(0).Text), 5)) <> "" Then '部門
   strSql = strSql & "and st03='" & Trim(Left(Trim(Combo1(0).Text), 5)) & "'"
End If
rs.Open "select st01,st02 from staff,SalaryData where ST01=SD01 and substr(st01,1,1) in ('6','7','8','9') and ((SD02 not in('P','F') or SD02 is null)) " & strSql & " order by 1", _
         cnnConnection, adOpenStatic, adLockReadOnly
Me.Combo1(1).AddItem ""
While Not rs.EOF
   Me.Combo1(1).AddItem Left(rs.Fields(0).Value & Space(7), 7) & rs.Fields(1).Value
   rs.MoveNext
Wend
If rs.State <> adStateClosed Then rs.Close
Set rs = Nothing

'Add By Sindy 2012/7/31
If Combo1(0).Tag = Combo1(0).Text And strTemp <> "" Then
   Combo1(1).Text = strTemp
End If
'2012/7/31 End

Screen.MousePointer = vbDefault
Me.Enabled = True
End Sub

Private Sub txtSystem_GotFocus()
txtSystem.SelStart = 0
txtSystem.SelLength = Len(txtSystem)
CloseIme
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
txtCode(Index).SelStart = 0
txtCode(Index).SelLength = Len(txtCode(Index))
CloseIme
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Then KeyAscii = UpperCase(KeyAscii)
End Sub
