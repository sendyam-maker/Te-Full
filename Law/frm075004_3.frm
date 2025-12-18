VERSION 5.00
Begin VB.Form frm075004_3 
   BorderStyle     =   1  '虫絬㏕﹚
   Caption         =   "糤计"
   ClientHeight    =   3030
   ClientLeft      =   410
   ClientTop       =   1500
   ClientWidth     =   8320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   8320
   Begin VB.TextBox textCP09 
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   0
      Top             =   90
      Width           =   1572
   End
   Begin VB.TextBox txtDocCh4 
      Height          =   270
      Index           =   7
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   1
      Top             =   450
      Width           =   420
   End
   Begin VB.Frame FrameAddPage 
      Appearance      =   0  'キ
      BackColor       =   &H80000004&
      Caption         =   "糤计"
      ForeColor       =   &H00FF0000&
      Height          =   1845
      Left            =   120
      TabIndex        =   31
      Top             =   1080
      Width           =   2535
      Begin VB.TextBox txtAddPage 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   270
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   6
         Top             =   1395
         Width           =   420
      End
      Begin VB.TextBox txtDocAdd 
         Height          =   270
         Index           =   4
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1125
         Width           =   420
      End
      Begin VB.TextBox txtDocAdd 
         Height          =   270
         Index           =   3
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   4
         Top             =   855
         Width           =   420
      End
      Begin VB.TextBox txtDocAdd 
         Height          =   270
         Index           =   1
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   3
         Top             =   579
         Width           =   420
      End
      Begin VB.TextBox txtDocAdd 
         Height          =   270
         Index           =   0
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   2
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "计羆璸:"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   240
         TabIndex        =   36
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "瓜Α计:"
         Height          =   180
         Left            =   240
         TabIndex        =   35
         Top             =   1170
         Width           =   765
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "ビ叫盡絛瞅计:"
         Height          =   180
         Left            =   240
         TabIndex        =   34
         Top             =   900
         Width           =   1485
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "弧计:"
         Height          =   180
         Left            =   240
         TabIndex        =   33
         Top             =   630
         Width           =   945
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "篕璶计:"
         Height          =   180
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.Frame FrameCP167 
      Appearance      =   0  'キ
      BackColor       =   &H80000004&
      Caption         =   "埃ゼ糵计"
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   2880
      TabIndex        =   25
      Top             =   1080
      Width           =   2535
      Begin VB.TextBox txtDocCp167 
         Height          =   270
         Index           =   0
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   7
         Top             =   315
         Width           =   420
      End
      Begin VB.TextBox txtDocCp167 
         Height          =   270
         Index           =   1
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   8
         Top             =   579
         Width           =   420
      End
      Begin VB.TextBox txtDocCp167 
         Height          =   270
         Index           =   3
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   9
         Top             =   855
         Width           =   420
      End
      Begin VB.TextBox txtDocCp167 
         Height          =   270
         Index           =   4
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   10
         Top             =   1125
         Width           =   420
      End
      Begin VB.TextBox txtCP167 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   270
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   11
         Top             =   1395
         Width           =   420
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "篕璶计:"
         Height          =   180
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "弧计:"
         Height          =   180
         Left            =   240
         TabIndex        =   29
         Top             =   630
         Width           =   945
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "ビ叫盡絛瞅计:"
         Height          =   180
         Left            =   240
         TabIndex        =   28
         Top             =   900
         Width           =   1485
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "瓜Α计:"
         Height          =   180
         Left            =   240
         TabIndex        =   27
         Top             =   1170
         Width           =   765
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "计羆璸:"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   765
      End
   End
   Begin VB.Frame FrameCP168 
      Appearance      =   0  'キ
      BackColor       =   &H80000004&
      Caption         =   "埃糵计"
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   5640
      TabIndex        =   19
      Top             =   1080
      Width           =   2535
      Begin VB.TextBox txtCP168 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   270
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   16
         Top             =   1395
         Width           =   420
      End
      Begin VB.TextBox txtDocCp168 
         Height          =   270
         Index           =   4
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   15
         Top             =   1125
         Width           =   420
      End
      Begin VB.TextBox txtDocCp168 
         Height          =   270
         Index           =   3
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   14
         Top             =   855
         Width           =   420
      End
      Begin VB.TextBox txtDocCp168 
         Height          =   270
         Index           =   1
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   13
         Top             =   579
         Width           =   420
      End
      Begin VB.TextBox txtDocCp168 
         Height          =   270
         Index           =   0
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   12
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "计羆璸:"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   240
         TabIndex        =   24
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "瓜Α计:"
         Height          =   180
         Left            =   240
         TabIndex        =   23
         Top             =   1170
         Width           =   765
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "ビ叫盡絛瞅计:"
         Height          =   180
         Left            =   240
         TabIndex        =   22
         Top             =   900
         Width           =   1485
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "弧计:"
         Height          =   180
         Left            =   240
         TabIndex        =   21
         Top             =   630
         Width           =   945
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "篕璶计:"
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "郎"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   6150
      TabIndex        =   17
      Top             =   90
      Width           =   690
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "玡礶"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   6900
      TabIndex        =   18
      Top             =   90
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Μ  ゅ  腹 "
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   39
      Top             =   135
      Width           =   945
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      Caption         =   "瓜Α瓜计 "
      Height          =   180
      Left            =   120
      TabIndex        =   38
      Top             =   510
      Width           =   945
   End
   Begin VB.Label LblPD20Note 
      AutoSize        =   -1  'True
      Caption         =   "ㄖタ"
      ForeColor       =   &H00FF00FF&
      Height          =   180
      Left            =   120
      TabIndex        =   37
      Top             =   765
      Width           =   750
   End
End
Attribute VB_Name = "frm075004_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Sindy 2023/3/16
Option Explicit

Public m_PrevForm As Form '玡礶
Public strReceiveNo As String
Public bolModify As Boolean 'True:э,False:琩高
Public objAddPage As Object, m_strCP135 As String
Public objCP167 As Object, m_strCP167 As String
Public objCP168 As Object, m_strCP168 As String
Dim pageD() As String


Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim bolUpd As Boolean

   Select Case Index
      Case 0 '郎
         On Error GoTo ErrHand
         cnnConnection.BeginTrans
         '穝盡弧计灿
         If pageD(1) = "" Then
            strSql = "INSERT INTO pagedetail(pd01,pd02,pd03,pd04,pd05,pd06,pd07,pd08,pd09,pd10,pd11,pd12,pd13,pd21)" & _
                     " VALUES('" & textCP09 & "'" & _
                     "," & CNULL(txtDocAdd(0), True) & "," & CNULL(txtDocAdd(1), True) & "," & CNULL(txtDocAdd(3), True) & "," & CNULL(txtDocAdd(4), True) & _
                     "," & CNULL(txtDocCp167(0), True) & "," & CNULL(txtDocCp167(1), True) & "," & CNULL(txtDocCp167(3), True) & "," & CNULL(txtDocCp167(4), True) & _
                     "," & CNULL(txtDocCp168(0), True) & "," & CNULL(txtDocCp168(1), True) & "," & CNULL(txtDocCp168(3), True) & "," & CNULL(txtDocCp168(4), True) & _
                     "," & CNULL(txtDocCh4(7), True) & ")"
            Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
            cnnConnection.Execute strSql
            bolUpd = True
         Else
            strExc(10) = ""
            If txtDocAdd(0) <> pageD(2) Then strExc(10) = strExc(10) & ",pd02=" & CNULL(txtDocAdd(0), True)
            If txtDocAdd(1) <> pageD(3) Then strExc(10) = strExc(10) & ",pd03=" & CNULL(txtDocAdd(1), True)
            If txtDocAdd(3) <> pageD(4) Then strExc(10) = strExc(10) & ",pd04=" & CNULL(txtDocAdd(3), True)
            If txtDocAdd(4) <> pageD(5) Then strExc(10) = strExc(10) & ",pd05=" & CNULL(txtDocAdd(4), True)
            If txtDocCp167(0) <> pageD(6) Then strExc(10) = strExc(10) & ",pd06=" & CNULL(txtDocCp167(0), True)
            If txtDocCp167(1) <> pageD(7) Then strExc(10) = strExc(10) & ",pd07=" & CNULL(txtDocCp167(1), True)
            If txtDocCp167(3) <> pageD(8) Then strExc(10) = strExc(10) & ",pd08=" & CNULL(txtDocCp167(3), True)
            If txtDocCp167(4) <> pageD(9) Then strExc(10) = strExc(10) & ",pd09=" & CNULL(txtDocCp167(4), True)
            If txtDocCp168(0) <> pageD(10) Then strExc(10) = strExc(10) & ",pd10=" & CNULL(txtDocCp168(0), True)
            If txtDocCp168(1) <> pageD(11) Then strExc(10) = strExc(10) & ",pd11=" & CNULL(txtDocCp168(1), True)
            If txtDocCp168(3) <> pageD(12) Then strExc(10) = strExc(10) & ",pd12=" & CNULL(txtDocCp168(3), True)
            If txtDocCp168(4) <> pageD(13) Then strExc(10) = strExc(10) & ",pd13=" & CNULL(txtDocCp168(4), True)
            If txtDocCh4(7) <> pageD(21) Then strExc(10) = strExc(10) & ",pd21=" & CNULL(txtDocCh4(7), True)
            If strExc(10) <> "" Then
               strExc(10) = Mid(strExc(10), 2)
               strSql = "UPDATE pagedetail SET " & strExc(10) & _
                        " WHERE pd01='" & textCP09 & "'"
               Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
               cnnConnection.Execute strSql
               bolUpd = True
            End If
         End If
         '穝玡礶逆
         If bolUpd = True Then
            If TypeName(objAddPage) <> "Nothing" Then objAddPage = txtAddPage.Text
            If TypeName(objCP167) <> "Nothing" Then objCP167 = txtCP167.Text
            If TypeName(objCP168) <> "Nothing" Then objCP168 = txtCP168.Text
            strSql = ""
            If m_strCP135 <> txtAddPage Then
               strSql = strSql & ",cp135=" & CNULL(txtAddPage, True)
            End If
            If m_strCP167 <> txtCP167 Then
               strSql = strSql & ",cp167=" & CNULL(txtCP167, True)
            End If
            If m_strCP168 <> txtCP168 Then
               strSql = strSql & ",cp168=" & CNULL(txtCP168, True)
            End If
            If strSql <> "" Then
               strSql = Mid(strSql, 2)
               strSql = "update caseprogress set " & strSql & " where cp09='" & textCP09 & "'"
               Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
               cnnConnection.Execute strSql
            End If
         End If
         cnnConnection.CommitTrans
      
      Case 1 '玡礶
         'm_PrevForm.Show
   End Select
   Unload Me
   
   Exit Sub
   
ErrHand:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox (Err.Description)
End Sub

Private Sub Form_Load()
Dim oText As Object

   MoveFormToCenter Me
   textCP09 = strReceiveNo
   
   If bolModify = True Then
      For Each oText In txtDocAdd
         oText.Enabled = True
      Next
      For Each oText In txtDocCp167
         oText.Enabled = True
      Next
      For Each oText In txtDocCp168
         oText.Enabled = True
      Next
      cmdOK(0).Visible = True
      txtDocCh4(7).Enabled = True
   Else
      For Each oText In txtDocAdd
         oText.Enabled = False
      Next
      For Each oText In txtDocCp167
         oText.Enabled = False
      Next
      For Each oText In txtDocCp168
         oText.Enabled = False
      Next
      cmdOK(0).Visible = False
      txtDocCh4(7).Enabled = False
   End If
   
   '弄盡弧计灿
   ReDim pageD(1 To 21) As String
   Call PUB_ReadPageDetail(strReceiveNo, pageD)
   '糤计
   txtDocAdd(0) = pageD(2)
   txtDocAdd(1) = pageD(3)
   txtDocAdd(3) = pageD(4)
   txtDocAdd(4) = pageD(5)
   '埃ゼ糵计
   txtDocCp167(0) = pageD(6)
   txtDocCp167(1) = pageD(7)
   txtDocCp167(3) = pageD(8)
   txtDocCp167(4) = pageD(9)
   '埃糵计
   txtDocCp168(0) = pageD(10)
   txtDocCp168(1) = pageD(11)
   txtDocCp168(3) = pageD(12)
   txtDocCp168(4) = pageD(13)
   txtDocCh4(7) = pageD(21)
   '璸衡计璸
   Call CountPage
   If pageD(20) <> "" Then
      strExc(10) = ""
      strSql = "SELECT CP09,CP01,CP10 FROM caseprogress WHERE CP09='" & pageD(20) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strExc(10) = GetPrjState6(RsTemp.Fields("CP01"), RsTemp.Fields("CP10"))
      End If
      LblPD20Note.Caption = pageD(20) & " " & strExc(10) & "ㄖタ"
      LblPD20Note.Visible = True
   Else
      LblPD20Note.Visible = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_PrevForm = Nothing
   Set frm075004_3 = Nothing
End Sub

'璸衡计璸
Private Sub CountPage()
   '璸:
   '糤计:
   If Val(txtDocAdd(0)) + Val(txtDocAdd(1)) + Val(txtDocAdd(3)) + Val(txtDocAdd(4)) > 0 Then
      txtAddPage = Val(txtDocAdd(0)) + Val(txtDocAdd(1)) + Val(txtDocAdd(3)) + Val(txtDocAdd(4))
   Else
      txtAddPage = ""
   End If
   '埃ゼ糵计:
   If Val(txtDocCp167(0)) + Val(txtDocCp167(1)) + Val(txtDocCp167(3)) + Val(txtDocCp167(4)) > 0 Then
      txtCP167 = Val(txtDocCp167(0)) + Val(txtDocCp167(1)) + Val(txtDocCp167(3)) + Val(txtDocCp167(4))
   Else
      txtCP167 = ""
   End If
   '埃糵计:
   If Val(txtDocCp168(0)) + Val(txtDocCp168(1)) + Val(txtDocCp168(3)) + Val(txtDocCp168(4)) > 0 Then
      txtCP168 = Val(txtDocCp168(0)) + Val(txtDocCp168(1)) + Val(txtDocCp168(3)) + Val(txtDocCp168(4))
   Else
      txtCP168 = ""
   End If
End Sub

Private Sub txtDocAdd_GotFocus(Index As Integer)
   TextInverse txtDocAdd(Index)
   CloseIme
End Sub

Private Sub txtDocAdd_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtDocAdd_Validate(Index As Integer, Cancel As Boolean)
   '璸衡计璸
   Call CountPage
End Sub

Private Sub txtDocCh4_GotFocus(Index As Integer)
   TextInverse txtDocCh4(Index)
   CloseIme
End Sub

Private Sub txtDocCh4_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtDocCp167_GotFocus(Index As Integer)
   TextInverse txtDocCp167(Index)
   CloseIme
End Sub

Private Sub txtDocCp167_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtDocCp167_Validate(Index As Integer, Cancel As Boolean)
   '璸衡计璸
   Call CountPage
End Sub

Private Sub txtDocCp168_GotFocus(Index As Integer)
   TextInverse txtDocCp168(Index)
   CloseIme
End Sub

Private Sub txtDocCp168_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtDocCp168_Validate(Index As Integer, Cancel As Boolean)
   '璸衡计璸
   Call CountPage
End Sub
