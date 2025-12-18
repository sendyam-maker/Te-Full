VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm880003 
   BorderStyle     =   1  '單線固定
   Caption         =   "補件期限"
   ClientHeight    =   5520
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7335
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6072
      TabIndex        =   8
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   5112
      TabIndex        =   7
      Top             =   70
      Width           =   912
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "新增(&A)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4368
      TabIndex        =   3
      Top             =   1920
      Width           =   912
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "刪除(&D)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5340
      TabIndex        =   4
      Top             =   1920
      Width           =   912
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "清除(&C)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6300
      TabIndex        =   5
      Top             =   1920
      Width           =   912
   End
   Begin MSForms.TextBox txtAddDeadline 
      Height          =   300
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1926;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtAddDeadline 
      Height          =   300
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   810
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtAddDeadline 
      Height          =   735
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   1140
      Width           =   5715
      VariousPropertyBits=   -1467989989
      MaxLength       =   200
      ScrollBars      =   2
      Size            =   "10081;1296"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstAddDeadline 
      Height          =   2220
      Left            =   150
      TabIndex        =   6
      Top             =   2670
      Width           =   7095
      VariousPropertyBits=   746586139
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "12509;3604"
      MatchEntry      =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      Caption         =   "法定期限："
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   810
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "補件本所期限："
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "文件內容："
      Height          =   252
      Left            =   120
      TabIndex        =   10
      Top             =   1140
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "補件本所期限   法定期限      文件內容"
      Height          =   252
      Left            =   180
      TabIndex        =   9
      Top             =   2400
      Width           =   6972
   End
End
Attribute VB_Name = "frm880003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/15 改成Form2.0 ; txtAddDeadline(index)、lstAddDeadline
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/17 日期欄已修改
Option Explicit
Public strAddDeadline1 As String, strAddDeadline2 As String, strAddDeadline3 As String
Public m_txtAddDeadline As String   '2008/11/26 add by sonia
Public PA09 As String, CP10 As String

Private Sub cmdMove_Click(Index As Integer)
Dim i As Integer, intlastIndex As Integer

If Index = 0 Then
   For i = 0 To 2
          If CheckKeyIn(i) = False Then
             txtAddDeadline(i).SetFocus
             txtAddDeadLine_GotFocus i
             Exit For
          End If
   Next
   If i < 3 Then Exit Sub
   For i = 0 To lstAddDeadline.ListCount - 1
          If txtAddDeadline(0) = Left(lstAddDeadline.List(i), 3) Then
             Exit For
          End If
   Next
   If i = lstAddDeadline.ListCount Then
      lstAddDeadline.AddItem txtAddDeadline(0) + vbTab + "          " + txtAddDeadline(1) + vbTab + txtAddDeadline(2)
      txtAddDeadline(0) = ""
      txtAddDeadline(1) = ""
      txtAddDeadline(2) = ""
      '2008/11/26 add by sonia P指示信預設本所期限
      If m_txtAddDeadline <> "" Then
         txtAddDeadline(0) = TransDate(m_txtAddDeadline, 1)
      End If
      '2008/11/26 END
      If lstAddDeadline.ListCount = 1 Then lstAddDeadline.ListIndex = 0
   Else
      ShowMsg MsgText(9200)
   End If
ElseIf Index = 1 Then
   If lstAddDeadline.ListIndex = -1 Then
      ShowMsg MsgText(8006)
   Else
      intlastIndex = lstAddDeadline.ListIndex
      lstAddDeadline.RemoveItem lstAddDeadline.ListIndex
      If lstAddDeadline.ListCount <> 0 Then
         If intlastIndex = lstAddDeadline.ListCount Then
            lstAddDeadline.ListIndex = lstAddDeadline.ListCount - 1
         Else
            lstAddDeadline.ListIndex = intlastIndex
         End If
      End If
   End If
Else
   lstAddDeadline.Clear
End If
txtAddDeadline(0).SetFocus
End Sub

Private Sub cmdOK_Click(Index As Integer)
 Dim i As Integer, varTemp As Variant

If Index = 0 Then
   strAddDeadline1 = ""
   strAddDeadline2 = ""
   strAddDeadline3 = ""
   If lstAddDeadline.ListCount > 1 Then
      For i = 0 To lstAddDeadline.ListCount - 2
             varTemp = Split(lstAddDeadline.List(i), vbTab)
             strAddDeadline1 = strAddDeadline1 + Trim(ChangeTStringToWString(CStr(varTemp(0)))) + ","
             strAddDeadline2 = strAddDeadline2 + Trim(ChangeTStringToWString(CStr(varTemp(1)))) + ","
             strAddDeadline3 = strAddDeadline3 + varTemp(2) + ","
      Next
      varTemp = Split(lstAddDeadline.List(i), vbTab)
      strAddDeadline1 = strAddDeadline1 + Trim(ChangeTStringToWString(CStr(varTemp(0))))
      strAddDeadline2 = strAddDeadline2 + Trim(ChangeTStringToWString(CStr(varTemp(1))))
      strAddDeadline3 = strAddDeadline3 + varTemp(2)
   ElseIf lstAddDeadline.ListCount = 1 Then
      varTemp = Split(lstAddDeadline.List(0), vbTab)
      strAddDeadline1 = ChangeTStringToWString(CStr(varTemp(0)))
      strAddDeadline2 = ChangeTStringToWString(CStr(varTemp(1)))
      strAddDeadline3 = varTemp(UBound(varTemp))
   End If
End If
Unload Me
End Sub
Private Sub Form_Load()
Dim i As Integer, varAddDeadLineTemp1, varAddDeadLineTemp2, varAddDeadLineTemp3, strTemp As String

MoveFormToCenter Me
If intPWhere = 國外_CF Then
   'txtAddDeadline(0).MaxLength = 8
   'txtAddDeadline(1).MaxLength = 8
   txtAddDeadline(0).MaxLength = 7
   txtAddDeadline(1).MaxLength = 7
Else
   txtAddDeadline(0).MaxLength = 7
   txtAddDeadline(1).MaxLength = 7
End If
If strAddDeadline1 <> "" Then
   varAddDeadLineTemp1 = Split(strAddDeadline1, ",")
   varAddDeadLineTemp2 = Split(strAddDeadline2, ",")
   varAddDeadLineTemp3 = Split(strAddDeadline3, ",")
   If strAddDeadline3 = "" Then
      lstAddDeadline.AddItem ChangeWStringToTString(CStr(varAddDeadLineTemp1(i))) + vbTab + "         " + ChangeWStringToTString(CStr(varAddDeadLineTemp2(i)))
   Else
      For i = 0 To UBound(varAddDeadLineTemp1)
         lstAddDeadline.AddItem ChangeWStringToTString(CStr(varAddDeadLineTemp1(i))) + vbTab + "         " + ChangeWStringToTString(CStr(varAddDeadLineTemp2(i))) + vbTab + varAddDeadLineTemp3(i)
      Next
   End If
End If
If lstAddDeadline.ListCount > 0 Then
   lstAddDeadline.ListIndex = 0
End If

'910704 Sieg 402
If PA09 = "020" Then
   Select Case CP10
      Case "101", "102", "103", "104", "105"
         '2008/11/26 MODIFY BY SONIA P-089548測試
         'txtAddDeadline(2).Text = "委託書、優先權證明書、圖式"
         txtAddDeadline(2).Text = "委託書、轉讓證明"
   End Select
End If

'2008/11/26 add by sonia P指示信預設本所期限
If m_txtAddDeadline <> "" Then
   txtAddDeadline(0) = TransDate(m_txtAddDeadline, 1)
End If
'2008/11/26 END

End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
'   Set frm880003 = Nothing
End Sub

Private Sub txtAddDeadLine_Validate(Index As Integer, Cancel As Boolean)
   If CheckKeyIn(Index) = False Then Cancel = True
   If Cancel Then txtAddDeadLine_GotFocus Index
End Sub

Private Function CheckKeyIn(intIndex As Integer) As Boolean
 Dim strTemp As String
   CheckKeyIn = False
   Select Case intIndex
      Case 0
         If CheckIsTaiwanDate(txtAddDeadline(intIndex).Text) Then
            If CheckReKey(txtAddDeadline(intIndex)) Then
               If Val(txtAddDeadline(intIndex).Text) <= Val(strSrvDate(2)) Then
                  MsgBox "不可小於系統日 !", vbCritical
               Else
                  txtAddDeadline(0).Text = TransDate(PUB_GetWorkDay1(txtAddDeadline(0).Text, True), 1) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天

                  '92.1.27 modify by sonia
                  'If txtAddDeadline(1) = "" Then txtAddDeadline(1) = TransDate(CompDate(2, 2, TransDate(txtAddDeadline(0), 2)), 1)
                  If txtAddDeadline(1) = "" And intPWhere = 國外_CF Then
                     txtAddDeadline(1) = TransDate(CompDate(2, 14, TransDate(txtAddDeadline(0), 2)), 1)
                  Else
                     '2008/11/21 modify by sonia 非台灣案改10天
                     'txtAddDeadline(1) = TransDate(CompDate(2, 2, TransDate(txtAddDeadline(0), 2)), 1)
                     If PA09 = "000" Then
                        txtAddDeadline(1) = TransDate(CompDate(2, 2, TransDate(txtAddDeadline(0), 2)), 1)
                     Else
                        txtAddDeadline(1) = TransDate(CompDate(2, 10, TransDate(txtAddDeadline(0), 2)), 1)
                     End If
                     '2008/11/21 end
                 End If
                  CheckKeyIn = 1
               End If
            End If
         End If
      Case 1
         If txtAddDeadline(intIndex) <> "" Then
            If CheckIsTaiwanDate(txtAddDeadline(intIndex).Text) Then
               '2010/8/17 modify by sonia
               'If txtAddDeadline(0) <= txtAddDeadline(1) Then
               If Val(txtAddDeadline(0)) <= Val(txtAddDeadline(1)) Then
                  If CheckReKey(txtAddDeadline(intIndex)) Then
                     CheckKeyIn = 1
                  End If
               Else
                  ShowMsg MsgText(1033)
               End If
            End If
         ElseIf txtAddDeadline(0) <> "" Then
            ShowMsg MsgText(1033)
         Else
            CheckKeyIn = 1
         End If
      Case 2
         If txtAddDeadline(intIndex) = "" Then
            MsgBox "補件內容不得空白 !", vbCritical
         Else
            CheckKeyIn = True
         End If
   End Select
End Function

Private Sub txtAddDeadLine_GotFocus(Index As Integer)
   TextInverse txtAddDeadline(Index)
End Sub
