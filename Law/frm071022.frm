VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm071022 
   BorderStyle     =   1  '單線固定
   Caption         =   "律師執業登錄地區維護"
   ClientHeight    =   2265
   ClientLeft      =   6765
   ClientTop       =   2670
   ClientWidth     =   3105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3105
   Begin VB.TextBox txtDKind 
      Height          =   372
      Left            =   600
      MaxLength       =   1
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox txtDYear 
      Height          =   372
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox txtST71 
      Height          =   855
      Left            =   120
      MaxLength       =   200
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   1395
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   75
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   2220
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   75
      Width           =   800
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   315
      Left            =   840
      TabIndex        =   7
      Top             =   540
      Width           =   2145
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3784;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "執業地區："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "frm071022"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/15 改成Form2.0 ; Combo1
'Create by Amy 2019/01/08
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()

    If Combo1 = MsgText(601) Then
        MsgBox "請選擇人員！"
        Exit Sub
    End If
    
    If SaveForm(Left(Combo1, 5)) = True Then
        Combo1.Tag = "": txtST71.Tag = txtST71
        MsgBox "資料已存檔！"
    End If
    Exit Sub

End Sub

Private Sub Combo1_Click()
    '有修改再選人彈訊息
    If Combo1.Tag <> MsgText(601) And Combo1.Tag <> Combo1 And txtST71.Tag <> txtST71 Then
        If MsgBox(Combo1.Tag & "尚未存檔，要存檔？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
            If SaveForm(Left(Combo1.Tag, 5)) = False Then
                Exit Sub
            End If
        End If
    End If
    
    txtST71 = ""
    Combo1.Tag = Combo1
    If Combo1 = MsgText(601) Then Exit Sub
    Call ReadST71
    
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    Call SetCombo1
End Sub

Private Sub SetCombo1()
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer, intIdx As Integer
    
    Combo1.Clear
    Combo1.AddItem "", intIdx
    intIdx = intIdx + 1
    
'modify by sonia 2019/2/14 改用共用Function GetLawerList
'    strQ = "Select ST01,ST02 From Staff Where (ST03 ='L01' OR ST20='13') and st04='1' " & _
'                "Order by ST03,ST51 DESC,ST01 "
'    intQ = 1
'    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
'    If intQ = 1 Then
'        Do While Not RsQ.EOF
'            Combo1.AddItem RsQ.Fields("ST01") & " " & RsQ.Fields("ST02"), intIdx
'            intIdx = intIdx + 1
'            RsQ.MoveNext
'        Loop
'    End If
'    RsQ.Close
Dim i As Integer, varTmp1 As Variant, strTmp As String

   strSql = GetLawerList("1")
   varTmp1 = Split(strSql, ";")
   For i = 0 To UBound(varTmp1)
      strTmp = varTmp1(i)
      Combo1.AddItem strTmp
   Next
'end 2019/2/14
End Sub

Private Sub ReadST71()
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer

    strQ = "Select ST71 From Staff Where ST01='" & Left(Combo1, 5) & "' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        txtST71 = "" & RsQ.Fields("ST71")
        txtST71.Tag = txtST71
    End If
    RsQ.Close
End Sub

Private Function SaveForm(ByVal StrST01 As String) As Boolean
    Dim strSql As String
On Error GoTo ErrHand

    SaveForm = False
    strSql = "Update Staff Set ST71='" & ChgSQL(txtST71) & "' Where ST01='" & StrST01 & "' "
    cnnConnection.Execute strSql
    SaveForm = True

ErrHand:
     If Err.Number <> 0 Then
        MsgBox "存檔有誤請洽電腦中心-" & Err.Description, vbCritical
     End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frm071022 = Nothing
End Sub
