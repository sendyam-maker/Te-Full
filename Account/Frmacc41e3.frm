VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc41e3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   3015
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   450
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   1575
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdErase 
      Caption         =   "-"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2475
      TabIndex        =   3
      Top             =   60
      Width           =   300
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      TabIndex        =   1
      Top             =   60
      Width           =   300
   End
   Begin MSForms.ListBox List2 
      Height          =   2205
      Left            =   270
      TabIndex        =   2
      Top             =   450
      Width           =   2490
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "4392;3889"
      MatchEntry      =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   315
      Left            =   270
      TabIndex        =   0
      Top             =   90
      Width           =   885
      VariousPropertyBits=   679493659
      MaxLength       =   20
      Size            =   "1561;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblName 
      Height          =   315
      Left            =   1215
      TabIndex        =   6
      Top             =   120
      Width           =   1000
      VariousPropertyBits=   19
      Caption         =   "lblName"
      Size            =   "1764;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "Frmacc41e3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/08 Form2.0已修改 text2/List2/lblName
'Created by Morgan 2013/4/26
Option Explicit

Private Sub cmdAdd_Click()
   Dim ii As Integer
   If lblName <> "" Then
      strExc(0) = lblName & vbTab & Text2
      For ii = 0 To List2.ListCount - 1
         If List2.List(ii) = strExc(0) Then
            Exit For
         End If
      Next
      If ii = List2.ListCount Then
         List2.AddItem strExc(0)
      End If
      Text2 = ""
      Text2.Tag = ""
   End If
End Sub

Private Sub cmdErase_Click()
   Dim idx As Integer
   
   cmdErase.Enabled = False
   If List2 <> "" Then
      idx = List2.ListIndex
      List2.RemoveItem idx
      If List2.ListCount > idx Then
         List2.Selected(idx) = True
      ElseIf List2.ListCount > 0 Then
         List2.Selected(idx - 1) = True
      End If
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim ii As Integer
   
   If Index = 0 Then
      strExc(1) = "": strExc(2) = ""
      For ii = 0 To List2.ListCount - 1
         If List2.List(ii) <> "" Then
            strExc(1) = strExc(1) & Left(List2.List(ii), InStr(List2.List(ii), vbTab) - 1) & ";"
            strExc(2) = strExc(2) & Mid(List2.List(ii), InStr(List2.List(ii), vbTab) + 1) & ";"
         End If
      Next
      intI = 1
   Else
      intI = 0
   End If
   Unload Me
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height, strBackPicPath1
   lblName = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Frmacc41e3 = Nothing
End Sub

Private Sub List2_Click()
   cmdErase.Enabled = False
   If List2.Text <> "" Then
      'If InStr(List2.Text, "智權部北所全部人員") = 0 Then
         cmdErase.Enabled = True
      'End If
   End If
End Sub

'Modify by Amy 2021/12/08 原:Integer
Private Sub List2_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   If KeyCode = 46 Then
      If cmdErase.Enabled = True Then
         cmdErase.Value = True
      End If
   End If
End Sub

Private Sub Text2_Change()
   lblName = ""
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
   CloseIme
End Sub

'Modify by Amy 2021/12/08 原:Integer
Private Sub Text2_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 <> "" And Text2.Tag <> Text2 Then
      If Left(Text2, 1) >= "6" And Left(Text2, 1) < "F" Then
         If PUB_GetStaffState(Text2, strExc(1), True) = 1 Then
            Text2.Tag = Text2
            lblName = strExc(1)
         Else
            Cancel = True
            Text2_GotFocus
         End If
      Else
         If GetIdFromName(Text2, strExc(1)) Then
            Text2.Tag = Text2
            Text2 = strExc(1)
            lblName = Text2.Tag
         Else
            Cancel = True
            Text2_GotFocus
         End If
      End If
   End If
End Sub

Private Function GetIdFromName(ByVal pName As String, ByRef pID As String) As Boolean
   strExc(0) = "select st01,st02 from staff where st02='" & ChgSQL(pName) & "' and st04='1' and st01>'6' and st01<'F'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.RecordCount = 1 Then
         pID = RsTemp.Fields("st01")
         GetIdFromName = True
      Else
         MsgBox "員工名稱重複，請直接輸入員工編號！"
      End If
   Else
      MsgBox "該員工名稱不存在！"
   End If
End Function


