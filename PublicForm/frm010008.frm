VERSION 5.00
Begin VB.Form frm010008 
   BorderStyle     =   1  '單線固定
   Caption         =   "主管機關來函查詢"
   ClientHeight    =   3510
   ClientLeft      =   1125
   ClientTop       =   1800
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5655
   Begin VB.CheckBox Check1 
      Caption         =   "是否含專業部輸入資料"
      Height          =   255
      Left            =   240
      TabIndex        =   41
      Top             =   240
      Width           =   3000
   End
   Begin VB.Frame fraChoose 
      BorderStyle     =   0  '沒有框線
      Height          =   732
      Index           =   2
      Left            =   1392
      TabIndex        =   34
      Top             =   1410
      Width           =   4644
      Begin VB.TextBox txtKeyIn 
         Height          =   264
         Index           =   5
         Left            =   240
         MaxLength       =   32
         TabIndex        =   8
         Top             =   408
         Width           =   3492
      End
      Begin VB.TextBox txtKeyIn 
         Height          =   264
         Index           =   4
         Left            =   240
         MaxLength       =   32
         TabIndex        =   7
         Top             =   0
         Width           =   3492
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   0
         Y1              =   504
         Y2              =   504
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4548
      Style           =   1  '圖片外觀
      TabIndex        =   30
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3720
      Style           =   1  '圖片外觀
      TabIndex        =   29
      Top             =   70
      Width           =   800
   End
   Begin VB.Frame fraChoose 
      BorderStyle     =   0  '沒有框線
      Height          =   732
      Index           =   4
      Left            =   1380
      TabIndex        =   36
      Top             =   2520
      Width           =   3852
      Begin VB.TextBox txtSystemC 
         Height          =   264
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   21
         Top             =   360
         Width           =   732
      End
      Begin VB.TextBox txtSystem 
         Height          =   264
         Left            =   0
         MaxLength       =   3
         TabIndex        =   13
         Top             =   0
         Width           =   732
      End
      Begin VB.Frame fraElse 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   1
         Left            =   1080
         TabIndex        =   37
         Top             =   360
         Width           =   2892
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   264
            Index           =   5
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   24
            Top             =   0
            Width           =   492
         End
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   264
            Index           =   4
            Left            =   1440
            MaxLength       =   1
            TabIndex        =   23
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtCode 
            Height          =   264
            Index           =   3
            Left            =   0
            MaxLength       =   6
            TabIndex        =   22
            Top             =   0
            Width           =   1332
         End
      End
      Begin VB.Frame fraTF 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   1
         Left            =   1080
         TabIndex        =   38
         Top             =   360
         Visible         =   0   'False
         Width           =   2772
         Begin VB.TextBox txtTFCode 
            Height          =   264
            Index           =   7
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   28
            Top             =   0
            Width           =   492
         End
         Begin VB.TextBox txtTFCode 
            Height          =   264
            Index           =   6
            Left            =   1680
            MaxLength       =   1
            TabIndex        =   27
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtTFCode 
            Height          =   264
            Index           =   5
            Left            =   1200
            MaxLength       =   1
            TabIndex        =   26
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtTFCode 
            Height          =   264
            Index           =   4
            Left            =   0
            MaxLength       =   5
            TabIndex        =   25
            Top             =   0
            Width           =   1092
         End
      End
      Begin VB.Frame fraElse 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   0
         Left            =   840
         TabIndex        =   39
         Top             =   0
         Width           =   2892
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   264
            Index           =   2
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   16
            Top             =   0
            Width           =   492
         End
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   264
            Index           =   1
            Left            =   1440
            MaxLength       =   1
            TabIndex        =   15
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtCode 
            Height          =   264
            Index           =   0
            Left            =   0
            MaxLength       =   6
            TabIndex        =   14
            Top             =   0
            Width           =   1332
         End
      End
      Begin VB.Frame fraTF 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   0
         Left            =   840
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
         Width           =   2772
         Begin VB.TextBox txtTFCode 
            Height          =   264
            Index           =   0
            Left            =   0
            MaxLength       =   5
            TabIndex        =   17
            Top             =   0
            Width           =   1092
         End
         Begin VB.TextBox txtTFCode 
            Height          =   264
            Index           =   1
            Left            =   1200
            MaxLength       =   1
            TabIndex        =   18
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtTFCode 
            Height          =   264
            Index           =   2
            Left            =   1680
            MaxLength       =   1
            TabIndex        =   19
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtTFCode 
            Height          =   264
            Index           =   3
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   20
            Top             =   0
            Width           =   492
         End
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   120
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.Frame fraChoose 
      BorderStyle     =   0  '沒有框線
      Height          =   372
      Index           =   3
      Left            =   1380
      TabIndex        =   35
      Top             =   2160
      Width           =   2892
      Begin VB.TextBox txtKeyIn 
         Height          =   264
         Index           =   7
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   11
         Top             =   0
         Width           =   1212
      End
      Begin VB.TextBox txtKeyIn 
         Height          =   264
         Index           =   6
         Left            =   0
         MaxLength       =   8
         TabIndex        =   10
         Top             =   0
         Width           =   1212
      End
      Begin VB.Line Line4 
         X1              =   1320
         X2              =   1440
         Y1              =   144
         Y2              =   144
      End
   End
   Begin VB.Frame fraChoose 
      BorderStyle     =   0  '沒有框線
      Height          =   372
      Index           =   1
      Left            =   1380
      TabIndex        =   33
      Top             =   1080
      Width           =   2892
      Begin VB.TextBox txtKeyIn 
         Height          =   264
         Index           =   2
         Left            =   0
         MaxLength       =   7
         TabIndex        =   4
         Top             =   0
         Width           =   1212
      End
      Begin VB.TextBox txtKeyIn 
         Height          =   264
         Index           =   3
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   5
         Top             =   0
         Width           =   1212
      End
      Begin VB.Line Line2 
         X1              =   1320
         X2              =   1440
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.Frame fraChoose 
      BorderStyle     =   0  '沒有框線
      Height          =   372
      Index           =   0
      Left            =   1260
      TabIndex        =   31
      Top             =   720
      Width           =   2892
      Begin VB.TextBox txtKeyIn 
         Height          =   264
         Index           =   1
         Left            =   1824
         MaxLength       =   8
         TabIndex        =   2
         Top             =   0
         Width           =   1050
      End
      Begin VB.TextBox txtKeyIn 
         Height          =   264
         Index           =   0
         Left            =   264
         MaxLength       =   8
         TabIndex        =   1
         Top             =   0
         Width           =   1050
      End
      Begin VB.Line Line1 
         X1              =   1440
         X2              =   1560
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label1 
         Caption         =   "D                               D"
         Height          =   252
         Left            =   120
         TabIndex        =   32
         Top             =   48
         Width           =   1650
      End
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "本所案號："
      CausesValidation=   0   'False
      Height          =   252
      Index           =   4
      Left            =   180
      TabIndex        =   12
      Top             =   2520
      Width           =   1212
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "機關文號："
      CausesValidation=   0   'False
      Height          =   252
      Index           =   3
      Left            =   180
      TabIndex        =   9
      Top             =   2160
      Width           =   1212
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "來函號數："
      CausesValidation=   0   'False
      Height          =   252
      Index           =   2
      Left            =   180
      TabIndex        =   6
      Top             =   1440
      Width           =   1212
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "收件日期："
      CausesValidation=   0   'False
      Height          =   252
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   1080
      Width           =   1212
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "收件號："
      CausesValidation=   0   'False
      Height          =   252
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frm010008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/16 Form2.0已修改 (無需修改)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/22 日期欄已修改
Option Explicit

'Option的選項
Dim intOpt As Integer
'Add By Cheng 2002/09/10
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdOK_Click(Index As Integer)
Dim intOptionChoose As Integer, j As Integer
Dim i As Integer, strKind1 As String, strKind2 As String
Dim adoquery As New ADODB.Recordset

   If Index = 0 Then
      For i = 0 To 4
         If OptChoose(i) Then
            Exit For
         End If
      Next
      'Add By Cheng 2002/09/10
      If CheckDataValidate = False Then Exit Sub
      
      If i = 4 Then
         If txtSystem = 馬德里案 Then
            strKind1 = txtSystem + txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)) + IIf(txtTFCode(2) = "", "0", txtTFCode(2)) + IIf(txtTFCode(3) = "", "00", txtTFCode(3))
            strKind2 = txtSystem + txtTFCode(4) + IIf(txtTFCode(5) = "", "0", txtTFCode(5)) + IIf(txtTFCode(6) = "", "0", txtTFCode(6)) + IIf(txtTFCode(7) = "", "00", txtTFCode(7))
         Else
             strKind1 = txtSystem + txtCode(0) + IIf(txtCode(1) = "", "0", txtCode(1)) + IIf(txtCode(2) = "", "00", txtCode(2))
             strKind2 = txtSystem + txtCode(3) + IIf(txtCode(4) = "", "0", txtCode(4)) + IIf(txtCode(5) = "", "00", txtCode(5))
         End If
      Else
         strKind1 = txtKeyIn(i * 2)
         strKind2 = txtKeyIn(i * 2 + 1)
      End If
      'edit by nickc 2007/02/06 不用 dll 了
      'Set adoquery = obj001.ReadOrgRst(i, strKind1, strKind2)
      Set adoquery = Cls001ReadOrgRst(i, strKind1, strKind2, Check1.Value)
      If adoquery.RecordCount = 0 Then
         ShowMsg MsgText(9211)
         Exit Sub
      End If
   End If
   If Index = 0 Then
      If intOpt = 4 Then
         txtSystem_Validate False
         If CheckRange = False Then Exit Sub
      Else
         For j = intOpt * 2 To intOpt * 2 + 1
                If CheckKeyIn(j) <> 1 Then
                    txtKeyIn(j).SetFocus
                    txtKeyIn_GotFocus j
                    Exit Sub
                End If
         Next
      End If
      frm010011.Label1 = "合計共 " & "" & adoquery.RecordCount & " 筆 !"
      frm010011.Show
      Me.Hide
   Else
      Unload Me
   End If
End Sub

Private Function CheckRange() As Boolean
Dim strKind1 As String, strKind2 As String

   If txtSystem = 馬德里案 Then
      strKind1 = txtSystem + txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)) + IIf(txtTFCode(2) = "", "0", txtTFCode(2)) + IIf(txtTFCode(3) = "", "00", txtTFCode(3))
      strKind2 = txtSystem + txtTFCode(4) + IIf(txtTFCode(5) = "", "0", txtTFCode(5)) + IIf(txtTFCode(6) = "", "0", txtTFCode(6)) + IIf(txtTFCode(7) = "", "00", txtTFCode(7))
   Else
      strKind1 = txtSystem + txtCode(0) + IIf(txtCode(1) = "", "0", txtCode(1)) + IIf(txtCode(2) = "", "00", txtCode(2))
      strKind2 = txtSystem + txtCode(3) + IIf(txtCode(4) = "", "0", txtCode(4)) + IIf(txtCode(5) = "", "00", txtCode(5))
   End If
   If strKind1 > strKind2 Then
      ShowMsg MsgText(1043)
      CheckRange = False
   Else
      CheckRange = True
   End If
End Function

'ADD BY SONIA 2014/5/7
Private Sub Form_Activate()
   If Left(Pub_StrUserSt03, 2) = "P1" Then
      OptChoose(2).Value = True
      Check1.Value = 1
   Else
      Check1.Value = 0
   End If
End Sub
'END 2014/5/7

Private Sub Form_Load()
   MoveFormToCenter Me
   'Modify By Sindy 2010/8/17 比對自動編號年度
   'txtKeyIn(0) = GetTaiwanThisYear
   'txtKeyIn(1) = GetTaiwanThisYear
   txtKeyIn(0) = CompAutoNumberYear(GetTaiwanThisYear)
   txtKeyIn(1) = CompAutoNumberYear(GetTaiwanThisYear)
   txtKeyIn(2) = GetTaiwanTodayDate
   txtKeyIn(3) = GetTaiwanTodayDate
   'edit by nickc 2007/02/06 不用 dll 了
   'If obj001 Is Nothing Then
   '   Set obj001 = CreateObject("prjTaieDll001.cls001")
   '   Set obj001.Connection = cnnConnection
   'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'edit by nickc 2007/02/06 不用 dll 了
   'Set obj001 = Nothing
   'Add By Cheng 2002/07/18
   Set frm010008 = Nothing
End Sub

Private Sub optChoose_Click(Index As Integer)
Dim i As Integer

   intOpt = Index
   For i = 0 To 4
          If Index = i Then
             fraChoose(i).Enabled = True
          Else
             fraChoose(i).Enabled = False
             OptChoose(i).Value = False
          End If
   Next
   If Index = 4 Then
      txtSystem.SetFocus
      ' 90.12.19 modify by louis
      txtCode(1) = "0"
      txtCode(1).Locked = True
      txtCode(2) = "00"
      txtCode(2).Locked = True
      txtCode(4) = "9"
      txtCode(4).Locked = True
      txtCode(5) = "99"
      txtCode(5).Locked = True
   Else
      txtKeyIn(Index * 2).SetFocus
      ' 90.12.19 modify by louis
      txtCode(1).Locked = False
      txtCode(2).Locked = False
      txtCode(4).Locked = False
      txtCode(5).Locked = False
   End If
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   txtCode(Index).SelStart = 0
   txtCode(Index).SelLength = Len(txtCode(Index))
End Sub

Private Sub txtCode_LostFocus(Index As Integer)
   Select Case Index
      Case 3
         'Add By Cheng 2002/09/27
         If blnClkSure = False Then
            If Me.txtSystem.Text & Me.txtCode(0).Text > Me.txtSystemC.Text & Me.txtCode(3).Text Then
               MsgBox "本所案號範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               Me.txtSystem.SetFocus
               txtSystem_GotFocus
               Exit Sub
            End If
         Else
            blnClkSure = False
         End If
   End Select
End Sub

Private Sub txtKeyIn_GotFocus(Index As Integer)
   txtKeyIn(Index).SelStart = 0
   txtKeyIn(Index).SelLength = Len(txtKeyIn(Index))
End Sub

Private Sub txtKeyIn_LostFocus(Index As Integer)
   'Add By Cheng 2002/09/10
   Select Case Index
      Case 1 '收件號
         If blnClkSure = False Then
            If Me.txtKeyIn(0).Text <> "" And Me.txtKeyIn(1).Text <> "" Then
               If Me.txtKeyIn(0).Text > Me.txtKeyIn(1).Text Then
                  MsgBox "收件號範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  Me.txtKeyIn(0).SetFocus
                  txtKeyIn_GotFocus 0
                  Exit Sub
               End If
            End If
         Else
            blnClkSure = False
         End If
      Case 3 '收件日期
         If blnClkSure = False Then
            If Me.txtKeyIn(2).Text <> "" And Me.txtKeyIn(3).Text <> "" Then
               If Val(Me.txtKeyIn(2).Text) > Val(Me.txtKeyIn(3).Text) Then
                  MsgBox "收件日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  Me.txtKeyIn(2).SetFocus
                  txtKeyIn_GotFocus 2
                  Exit Sub
               End If
            End If
         Else
            blnClkSure = False
         End If
      'ADD BY SONIA 2014/5/7
      Case 4 '來函號數
         If Me.txtKeyIn(5).Text = "" Then
            Me.txtKeyIn(5).Text = Me.txtKeyIn(4).Text
         End If
      'END 2014/5/7
      Case 5 '來函號數
         If blnClkSure = False Then
            If Me.txtKeyIn(4).Text <> "" And Me.txtKeyIn(5).Text <> "" Then
               If Me.txtKeyIn(4).Text > Me.txtKeyIn(5).Text Then
                  MsgBox "來函號數範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  Me.txtKeyIn(4).SetFocus
                  txtKeyIn_GotFocus 4
                  Exit Sub
               End If
            End If
         Else
            blnClkSure = False
         End If
      Case 7 '機關文號
         If blnClkSure = False Then
            If Me.txtKeyIn(6).Text <> "" And Me.txtKeyIn(7).Text <> "" Then
               If Me.txtKeyIn(6).Text > Me.txtKeyIn(7).Text Then
                  MsgBox "機關文號範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  Me.txtKeyIn(6).SetFocus
                  txtKeyIn_GotFocus 6
                  Exit Sub
               End If
            End If
         Else
            blnClkSure = False
         End If
   End Select

End Sub

Private Sub txtKeyIn_Validate(Index As Integer, Cancel As Boolean)
   If CheckKeyIn(Index) = -1 Then
      Cancel = True
      txtKeyIn_GotFocus Index
   End If
End Sub
Private Sub txtSystem_Change()
Static bolTF As Integer

   If txtSystem.Text = 馬德里案 Then
      fraTF(0).Visible = True
      fraTF(1).Visible = True
      fraElse(0).Visible = False
      fraElse(1).Visible = False
      bolTF = True
   Else
      fraTF(0).Visible = False
      fraTF(1).Visible = False
      fraElse(0).Visible = True
      fraElse(1).Visible = True
      If bolTF Then
         bolTF = False
      End If
   End If
   txtSystemC = txtSystem
End Sub

Private Sub txtSystem_GotFocus()
   txtSystem.SelStart = 0
   txtSystem.SelLength = Len(txtSystem)
End Sub
Private Function CheckKeyIn(ByVal intIndex As Integer) As Integer
Dim bolRight As Boolean

CheckKeyIn = -1
   Select Case intIndex
      Case 0
                 If txtKeyIn(intIndex) = "" Or Len(txtKeyIn(intIndex)) = 8 Then
'                           If txtKeyIn(intIndex) > txtKeyIn(1) Then
'                              CheckKeyIn = 0
'                              ShowMsg MsgText(1043)
'                           Else
                       CheckKeyIn = 1
'                           End If
                 Else
                    ShowMsg MsgText(1044)
                 End If
      Case 1
                 If Len(txtKeyIn(intIndex)) = 8 Then
                    'Modify By Cheng 2002/09/10
'                           If txtKeyIn(0) <= txtKeyIn(intIndex) Then
                       CheckKeyIn = 1
'                           Else
'                              ShowMsg MsgText(1043)
'                              CheckKeyIn = 0
'                           End If
                 Else
                    ShowMsg MsgText(1044)
                 End If
      Case 2
                 If txtKeyIn(intIndex) = "" Then
'                           If txtKeyIn(intIndex) > txtKeyIn(3) Then
'                              CheckKeyIn = 0
'                              ShowMsg MsgText(1043)
'                           Else
                       CheckKeyIn = 1
                       Exit Function
'                           End If
                 End If
                 If CheckIsTaiwanDate(txtKeyIn(intIndex)) Then
                    CheckKeyIn = 1
                    Exit Function
                 End If
      Case 3
                 If CheckIsTaiwanDate(txtKeyIn(intIndex)) Then
                    'Modify By Cheng 2002/09/10
'                           If txtKeyIn(2) <= txtKeyIn(intIndex) Then
                       CheckKeyIn = 1
'                           Else
'                              ShowMsg MsgText(1043)
'                              CheckKeyIn = 0
'                           End If
                 End If
      Case 4
                 If txtKeyIn(5) <> "" Then
'                           If txtKeyIn(intIndex) > txtKeyIn(5) Then
'                              CheckKeyIn = 0
'                              ShowMsg MsgText(1043)
'                           Else
                       CheckKeyIn = 1
'                           End If
                 Else
                    CheckKeyIn = 1
                 End If
      Case 5
                 If txtKeyIn(intIndex) <> "" Then
                    'Modify By Cheng 2002/09/10
'                           If txtKeyIn(4) <= txtKeyIn(intIndex) Then
                       CheckKeyIn = 1
'                           Else
'                              ShowMsg MsgText(1043)
'                              CheckKeyIn = 0
'                           End If
                 Else
                    ShowMsg MsgText(9015)
                 End If
      Case 6
                 If txtKeyIn(intIndex) <> "" Then
'                           If txtKeyIn(intIndex) > txtKeyIn(7) Then
'                              CheckKeyIn = 0
'                              ShowMsg MsgText(1043)
'                           Else
                       CheckKeyIn = 1
'                           End If
                 Else
                    CheckKeyIn = 1
                 End If
      Case 7
                 If txtKeyIn(intIndex) <> "" Then
                    'Modify By Cheng 2002/09/10
'                           If txtKeyIn(7) <= txtKeyIn(intIndex) Then
                       CheckKeyIn = 1
'                           Else
'                              ShowMsg MsgText(1043)
'                              CheckKeyIn = 0
'                           End If
                 Else
                    ShowMsg MsgText(9015)
                 End If
      Case Else
                 CheckKeyIn = 1
   End Select
End Function

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSystem_Validate(Cancel As Boolean)
   If txtSystem <> "" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetSystemKind(txtSystem.Text) = False Then
      If ClsPDGetSystemKind(txtSystem.Text) = False Then
         Cancel = True
         txtSystem_GotFocus
      End If
   End If
End Sub

Private Sub txtTFCode_GotFocus(Index As Integer)
   txtTFCode(Index).SelStart = 0
   txtTFCode(Index).SelLength = Len(txtTFCode(Index))
End Sub

Private Function CheckDataValidate() As Boolean
   'Add By Cheng 2002/09/10
   CheckDataValidate = False
   blnClkSure = False
   If Me.OptChoose(0).Value Then
      If Me.txtKeyIn(0).Text <> "" And Me.txtKeyIn(1).Text <> "" Then
         If Me.txtKeyIn(0).Text > Me.txtKeyIn(1).Text Then
            MsgBox "收件號範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            blnClkSure = True
            Me.txtKeyIn(0).SetFocus
            txtKeyIn_GotFocus 0
            Exit Function
         End If
      End If
   ElseIf Me.OptChoose(1).Value Then
      If Me.txtKeyIn(2).Text <> "" And Me.txtKeyIn(3).Text <> "" Then
         If Val(Me.txtKeyIn(2).Text) > Val(Me.txtKeyIn(3).Text) Then
            MsgBox "收件日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            blnClkSure = True
            Me.txtKeyIn(2).SetFocus
            txtKeyIn_GotFocus 2
            Exit Function
         End If
      End If
   ElseIf Me.OptChoose(2).Value Then
      If Me.txtKeyIn(4).Text <> "" And Me.txtKeyIn(5).Text <> "" Then
         If Me.txtKeyIn(4).Text > Me.txtKeyIn(5).Text Then
            MsgBox "來函號數範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            blnClkSure = True
            Me.txtKeyIn(4).SetFocus
            txtKeyIn_GotFocus 4
            Exit Function
         End If
      End If
   ElseIf Me.OptChoose(3).Value Then
      If Me.txtKeyIn(6).Text <> "" And Me.txtKeyIn(7).Text <> "" Then
         If Me.txtKeyIn(6).Text > Me.txtKeyIn(7).Text Then
            MsgBox "機關文號範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            blnClkSure = True
            Me.txtKeyIn(6).SetFocus
            txtKeyIn_GotFocus 6
            Exit Function
         End If
      End If
   'Add By Cheng 2002/09/27
   ElseIf Me.OptChoose(4).Value Then
      If Me.txtSystem.Text & Me.txtCode(0).Text > Me.txtSystemC.Text & Me.txtCode(3).Text Then
         MsgBox "本所案號範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         blnClkSure = True
         Me.txtSystem.SetFocus
         txtSystem_GotFocus
         Exit Function
      End If
   End If
   CheckDataValidate = True
End Function
