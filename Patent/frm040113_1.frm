VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040113_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "內部收文分案"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5265
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox txtExV 
      Height          =   300
      Left            =   1215
      MaxLength       =   5
      TabIndex        =   0
      Top             =   2475
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "取消(&C)"
      CausesValidation=   0   'False
      Height          =   380
      Left            =   4275
      TabIndex        =   10
      Top             =   2010
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   380
      Left            =   3420
      TabIndex        =   9
      Top             =   2010
      Width           =   800
   End
   Begin MSForms.Label lblCaseName 
      Height          =   255
      Left            =   1215
      TabIndex        =   17
      Top             =   1122
      Width           =   4000
      VariousPropertyBits=   27
      Caption         =   "lblCaseName"
      Size            =   "7056;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblAppName 
      Height          =   255
      Left            =   1215
      TabIndex        =   16
      Top             =   825
      Width           =   4000
      VariousPropertyBits=   27
      Caption         =   "lblAppName"
      Size            =   "7056;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboCP14 
      Height          =   330
      Left            =   1215
      TabIndex        =   2
      Top             =   2040
      Width           =   1785
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3149;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "基　　數："
      Height          =   255
      Left            =   180
      TabIndex        =   15
      Top             =   2475
      Width           =   900
   End
   Begin VB.Label lblOurDeadLine 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   255
      Left            =   1215
      TabIndex        =   14
      Top             =   1710
      Width           =   900
   End
   Begin VB.Label lblRecNo 
      AutoSize        =   -1  'True
      Caption         =   "總收文號："
      Height          =   255
      Left            =   1215
      TabIndex        =   13
      Top             =   240
      Width           =   900
   End
   Begin VB.Label lblProperty 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   255
      Left            =   1215
      TabIndex        =   12
      Top             =   1416
      Width           =   900
   End
   Begin VB.Label lblCaseNo 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   255
      Left            =   1215
      TabIndex        =   11
      Top             =   534
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   255
      Left            =   180
      TabIndex        =   8
      Top             =   1710
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "承  辦  人："
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   2100
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "總收文號："
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   240
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   1416
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   1122
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申  請  人："
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   828
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   534
      Width           =   900
   End
End
Attribute VB_Name = "frm040113_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/07 改成Form2.0 ; CboCP14、lblAppName、lblCaseName
'Created by Morgan 2014/10/15
Option Explicit
Public mInputKey As String 'Add by Lydia 2015/01/20
Private Sub CboCP14_Change()
   If Len(CboCP14) <> 5 Then Exit Sub
   CheckCP14
End Sub

Private Sub CboCP14_GotFocus()
InverseTextBox CboCP14
End Sub

'Modified by Lydia 2021/10/07 改成Form 2.0
'Private Sub CboCP14_KeyPress(KeyAscii As Integer)
Private Sub CboCP14_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Function CheckCP14() As Boolean
Dim strText As String
   'Modified by Lydia 2015/01/22 +非離職不可變
   If mInputKey = "" Then
      CheckCP14 = True
   Else
        If CboCP14.ListIndex < 0 Then
           If CboCP14.Text > "" Then
              If CboCP14.Tag = CboCP14 Then
                 CheckCP14 = True
              Else
                 If IsNumeric(Left(CboCP14, 5)) Then
                    strText = Left(CboCP14, 5)
                 Else
                    '依員工姓名抓取員工編號
                    strText = GetPrjSalesNM_2(LTrim(Mid(CboCP14.Text, 6, Len(CboCP14.Text) - 5)))
                 End If
                 
                 If strText <> "" Then
                    For intI = 0 To CboCP14.ListCount - 1
                       If InStr(CboCP14.List(intI), strText) = 1 Then
                          CboCP14.ListIndex = intI
                          CheckCP14 = True
                          CboCP14.Tag = CboCP14
                          Exit For
                       End If
                    Next
                 End If
              End If
           End If
        Else
           CheckCP14 = True
        End If
   End If
End Function

Private Sub CboCP14_LostFocus()
   CheckCP14
End Sub

Private Sub CboCP14_Validate(Cancel As Boolean)
   CheckCP14
End Sub

Private Sub cmdExit_Click()
   frm040113.strCP14 = ""
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If CheckCP14 = False Then
      MsgBox "請輸入在職專利工程師！", vbExclamation
      CboCP14.SetFocus
      CboCP14_GotFocus
   Else
      frm040113.strCP14 = Left(CboCP14, 5)
      'Add by Lydia 2015/01/20
      frm040113.m_EV02 = Trim(txtExV.Text)
      Unload Me
   End If
End Sub

Private Sub Form_Activate()
'Add by Lydia 2015/01/20
If mInputKey <> "" Then
   CboCP14.Locked = False
   CboCP14.SetFocus
Else
   CboCP14.Locked = True
   txtExV.SetFocus
End If
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm040113_1 = Nothing
End Sub

'Add by Lydia 2015/01/20 開放基數(計件值)修改
Private Sub txtExV_KeyPress(KeyAscii As Integer)
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") Then
      KeyAscii = 0
   End If
End Sub
Private Sub txtExV_GotFocus()
InverseTextBox txtExV
End Sub
