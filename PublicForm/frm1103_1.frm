VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm1103_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "相關卷號"
   ClientHeight    =   1785
   ClientLeft      =   90
   ClientTop       =   1680
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   6930
   Begin VB.Frame fraElse 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   1920
      TabIndex        =   13
      Top             =   720
      Width           =   2652
      Begin VB.TextBox txtCode 
         Height          =   288
         Index           =   2
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   3
         Top             =   60
         Width           =   492
      End
      Begin VB.TextBox txtCode 
         Height          =   288
         Index           =   1
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   2
         Top             =   60
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   288
         Index           =   0
         Left            =   0
         MaxLength       =   6
         TabIndex        =   1
         Top             =   60
         Width           =   1212
      End
   End
   Begin VB.TextBox txtSystem 
      Height          =   288
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   0
      Top             =   780
      Width           =   732
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5964
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5136
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin VB.Frame fraTF 
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1920
      TabIndex        =   14
      Top             =   720
      Width           =   2652
      Begin VB.TextBox txtTFCode 
         Height          =   288
         Index           =   3
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   7
         Top             =   60
         Width           =   492
      End
      Begin VB.TextBox txtTFCode 
         Height          =   288
         Index           =   2
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   6
         Top             =   60
         Width           =   372
      End
      Begin VB.TextBox txtTFCode 
         Height          =   288
         Index           =   1
         Left            =   960
         MaxLength       =   1
         TabIndex        =   5
         Top             =   60
         Width           =   372
      End
      Begin VB.TextBox txtTFCode 
         Height          =   288
         Index           =   0
         Left            =   0
         MaxLength       =   5
         TabIndex        =   4
         Top             =   60
         Width           =   972
      End
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1200
      TabIndex        =   8
      Top             =   1290
      Width           =   5595
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "9869;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   252
      Left            =   240
      TabIndex        =   12
      Top             =   780
      Width           =   1092
   End
   Begin VB.Label Label6 
      Caption         =   "案件名稱："
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   972
   End
End
Attribute VB_Name = "frm1103_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/10/15 改成Form2.0 (cboCaseName)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
Option Explicit

Private Sub cmdok_Click(Index As Integer)
Dim varSaveCursor, bolRt As Boolean

Select Case Index
   Case 0 '確定
      varSaveCursor = Screen.MousePointer
      Screen.MousePointer = vbHourglass
      If txtSystem = 馬德里案 Then
         bolRt = CheckKeyIn1(3)
      Else
         bolRt = CheckKeyIn2(2)
      End If
      If bolRt Then
         frm1103_2.intWhereComeFrom = 1
         frm1103_2.lblSystem.Caption = Me.txtSystem.Text
         If txtSystem = 馬德里案 Then
            frm1103_2.lblTFCode(0) = txtTFCode(0)
            frm1103_2.lblTFCode(1) = txtTFCode(1)
            frm1103_2.lblTFCode(2) = txtTFCode(2)
            frm1103_2.lblTFCode(3) = txtTFCode(3)
         Else
            frm1103_2.lblCode(0) = txtCode(0)
            frm1103_2.lblCode(1) = txtCode(1)
            frm1103_2.lblCode(2) = txtCode(2)
         End If
         frm1103_2.SetFrom1103
         frm1103_2.Show
         Me.Hide
      End If
      Screen.MousePointer = varSaveCursor
   Case 1 '結束
      Unload Me
End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   ClearAll
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm1103_1 = Nothing
End Sub

Private Sub txtSystem_Change()
If txtSystem.Text = 馬德里案 Then
   fraTF.Visible = True
   fraElse.Visible = False
Else
   fraTF.Visible = False
   fraElse.Visible = True
End If
If cboCaseName.ListCount > 0 Then cboCaseName.Clear
End Sub

Private Sub txtSystem_GotFocus()
txtSystem.SelStart = 0
txtSystem.SelLength = Len(txtSystem.Text)
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSystem_Validate(Cancel As Boolean)
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.GetGroupCase(txtSystem, strGroup) = False Then
'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件(P)，但非此類案件時外專程序人員不可操作。
If Not (FMP2open = True And (txtSystem.Text = "P" Or txtSystem.Text = "PS")) Then
    If ClsPDGetGroupCase(txtSystem, strGroup) = False Then
       ShowMsg MsgText(1056)
       Cancel = True
       txtSystem_GotFocus
    End If
End If
End Sub

Private Sub txtCode_Change(Index As Integer)
If cboCaseName.ListCount > 0 Then cboCaseName.Clear
End Sub

Private Sub txtTFCode_Change(Index As Integer)
If cboCaseName.ListCount > 0 Then cboCaseName.Clear
End Sub

Private Sub txtTFCode_GotFocus(Index As Integer)
txtTFCode(Index).SelStart = 0
txtTFCode(Index).SelLength = Len(txtTFCode(Index).Text)
End Sub

Private Sub txtTFCode_Validate(Index As Integer, Cancel As Boolean)
CheckKeyIn1 (Index)
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
txtCode(Index).SelStart = 0
txtCode(Index).SelLength = Len(txtCode(Index).Text)
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
CheckKeyIn2 (Index)
End Sub

Private Function CheckKeyIn1(ByRef intIndex As Integer) As Boolean
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String

If Len(txtTFCode(intIndex)) > 0 And Len(txtTFCode(intIndex)) < txtTFCode(intIndex).MaxLength Then
   ShowMsg MsgText(9)
ElseIf intIndex = 3 Then
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
         IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3) Then
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   Dim mOKchk As Boolean
   If FMP2open = False Then
      mOKchk = ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
              IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3)
   Else
      mOKchk = PUB_FMPtoCheck(0, 1, Pub_strUserST05, txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
              IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)))
      If mOKchk = True Then '借由另一模組取值
         mOKchk = ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
              IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3)
      End If
   End If
   If mOKchk = True Then
   'end 'Add by Lydia 2014/10/31
  '  If ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
          IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3) Then
       SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
       CheckKeyIn1 = True
    End If

Else
   CheckKeyIn1 = True
End If
End Function

Private Function CheckKeyIn2(ByRef intIndex As Integer) As Boolean
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String

If Len(txtCode(intIndex)) > 0 And Len(txtCode(intIndex)) < txtCode(intIndex).MaxLength Then
   ShowMsg MsgText(9)
ElseIf intIndex = 2 Then
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3) Then
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   Dim mOKchk As Boolean
   If FMP2open = False Then
      mOKchk = ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
             IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3)
   Else
      mOKchk = PUB_FMPtoCheck(0, 1, Pub_strUserST05, txtSystem, txtCode(0), _
             IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)))
      If mOKchk = True Then '借由另一模組取值
         mOKchk = ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
             IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3)
      End If
   End If
   If mOKchk = True Then
   'end 'Add by Lydia 2014/10/31
  '  If ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
         IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3) Then
       SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
       CheckKeyIn2 = True
    End If

Else
   CheckKeyIn2 = True
End If
End Function

Private Function ClearAll()
   txtSystem = ""
   txtCode(0) = ""
   txtCode(1) = ""
   txtCode(2) = ""
   txtTFCode(0) = ""
   txtTFCode(1) = ""
   txtTFCode(2) = ""
   txtTFCode(3) = ""
   cboCaseName.Clear
End Function
