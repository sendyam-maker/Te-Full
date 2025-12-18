VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060108_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦/核稿期限、會稿日輸入"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "新細明體-ExtB"
      Size            =   9
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   8955
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   7965
      TabIndex        =   43
      Top             =   90
      Width           =   800
   End
   Begin VB.TextBox txtEP34 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7515
      MaxLength       =   1
      TabIndex        =   2
      Top             =   2940
      Width           =   360
   End
   Begin VB.TextBox txtCP48T 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5130
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2610
      Width           =   1125
   End
   Begin VB.TextBox txtEP07T 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7515
      MaxLength       =   7
      TabIndex        =   3
      Top             =   3285
      Width           =   1125
   End
   Begin VB.TextBox txtEP08T 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5130
      MaxLength       =   7
      TabIndex        =   1
      Top             =   2940
      Width           =   1125
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面(&U)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6720
      TabIndex        =   6
      Top             =   90
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5895
      TabIndex        =   5
      Top             =   90
      Width           =   800
   End
   Begin VB.TextBox txtCaseNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   630
      Width           =   495
   End
   Begin VB.TextBox txtCaseNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   1980
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   630
      Width           =   855
   End
   Begin VB.TextBox txtCaseNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   2820
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   630
      Width           =   255
   End
   Begin VB.TextBox txtCaseNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   3060
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   630
      Width           =   375
   End
   Begin MSForms.TextBox txtCP64 
      Height          =   1050
      Left            =   1485
      TabIndex        =   4
      Top             =   3285
      Width           =   4800
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "8467;1852"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblEP09T 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   7515
      TabIndex        =   42
      Top             =   2610
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(Y/N)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   14
      Left            =   8100
      TabIndex        =   41
      Top             =   2940
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否會稿:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   6705
      TabIndex        =   40
      Top             =   2940
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦期限:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   12
      Left            =   4320
      TabIndex        =   39
      Top             =   2610
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "會稿日:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   11
      Left            =   6885
      TabIndex        =   38
      Top             =   3285
      Width           =   585
   End
   Begin VB.Label lblEP04 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   1500
      TabIndex        =   37
      Top             =   2940
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   10
      Left            =   645
      TabIndex        =   36
      Top             =   3285
      Width           =   765
   End
   Begin VB.Label lblPA08T 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   5130
      TabIndex        =   35
      Top             =   2280
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利種類:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   9
      Left            =   4320
      TabIndex        =   34
      Top             =   2280
      Width           =   765
   End
   Begin VB.Label lblCP14 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   1500
      TabIndex        =   33
      Top             =   2610
      Width           =   885
   End
   Begin MSForms.Label lblEP04T 
      Height          =   285
      Left            =   2430
      TabIndex        =   32
      Top             =   2940
      Width           =   1305
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2302;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "核稿人:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   5
      Left            =   825
      TabIndex        =   31
      Top             =   2940
      Width           =   585
   End
   Begin VB.Label lblCP09 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   5130
      TabIndex        =   30
      Top             =   1950
      Width           =   2445
   End
   Begin MSForms.Label lblCP14T 
      Height          =   285
      Left            =   2430
      TabIndex        =   29
      Top             =   2610
      Width           =   1305
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCP10T 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   1500
      TabIndex        =   28
      Top             =   2280
      Width           =   1845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "總收文號:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   8
      Left            =   4320
      TabIndex        =   27
      Top             =   1950
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "完稿日:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   7
      Left            =   6885
      TabIndex        =   26
      Top             =   2610
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "核稿期限:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   4320
      TabIndex        =   25
      Top             =   2940
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   825
      TabIndex        =   24
      Top             =   2610
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   645
      TabIndex        =   23
      Top             =   2280
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文日:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   825
      TabIndex        =   22
      Top             =   1950
      Width           =   585
   End
   Begin VB.Label lblCP05T 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   1500
      TabIndex        =   21
      Top             =   1950
      Width           =   1185
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   3
      Left            =   1500
      TabIndex        =   20
      Top             =   1620
      Width           =   3195
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5636;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   2
      Left            =   1500
      TabIndex        =   19
      Top             =   1290
      Width           =   3195
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5636;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   1
      Left            =   1500
      TabIndex        =   18
      Top             =   960
      Width           =   3195
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblAppDate 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   5130
      TabIndex        =   17
      Top             =   630
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請日:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   4500
      TabIndex        =   16
      Top             =   630
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   270
      TabIndex        =   15
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "(中):"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1065
      TabIndex        =   14
      Top             =   960
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(英):"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1065
      TabIndex        =   13
      Top             =   1290
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(外):"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   1065
      TabIndex        =   12
      Top             =   1620
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   645
      TabIndex        =   11
      Top             =   630
      Width           =   765
   End
End
Attribute VB_Name = "frm060108_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/25 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Dim bolIsValidate As Boolean


Private Sub cmdBack_Click()
    Call frm060108.SetGrid(False)
    frm060108.Show
    Unload Me
End Sub

Public Sub SetData(ByRef rstGrid As ADODB.Recordset, ByVal iRow As Integer)
    
    Dim ii As Integer, stPA08 As String
    
    With frm060108
      For ii = 1 To 4
          txtCaseNo(ii) = .txtCaseNo(ii)
      Next ii
      lblAppDate = .lblAppDate
      For ii = 1 To 3
          lblCaseName(ii) = .lblCaseName(ii)
      Next ii
    End With
    With rstGrid
      .Move iRow - 1, adBookmarkFirst
      lblCP05T = "" & .Fields("CP05T")
      lblCP09 = "" & .Fields("CP09")
      lblCP10T = "" & .Fields("CP10T")
      lblCP14 = "" & .Fields("CP14")
      lblCP14T = "" & .Fields("CP14T")
      lblEP04 = "" & .Fields("EP04")
      lblEP04T = "" & .Fields("EP04T")
      lblPA08T = "" & .Fields("PA08T")
      stPA08 = "" & .Fields("PA08")
      
      '承辦期限
      txtCP48T = "" & .Fields("CP48T")
      '完稿日
      lblEP09T = "" & .Fields("EP09T")
      If lblEP09T <> "" Then
         txtCP48T.Enabled = False
      End If
      '核稿期限
      txtEP08T = "" & .Fields("EP08T")
      '是否會稿
      txtEP34 = "" & .Fields("EP34")
      '會稿日
      txtEP07T = "" & .Fields("EP07T")
      '進度備註
      txtCP64 = "" & .Fields("CP64")
    End With
    
End Sub

Private Function FormSave() As Boolean

    Dim stEP07 As String, stEP08 As String, stEP34 As String, stCP09 As String, stCP48  As String
    
On Error GoTo flgError

cnnConnection.BeginTrans
    
    stCP09 = lblCP09
    stCP48 = TransDate(txtCP48T, 2)
    stEP08 = TransDate(txtEP08T, 2)
    stEP07 = TransDate(txtEP07T, 2)
    stEP34 = txtEP34
    If stEP34 = "" Then stEP34 = "Y" '2012/5/15 ADD BY SONIA
    strSql = " Begin" & _
            " Update ENGINEERPROGRESS Set EP07=" & IIf(stEP07 = "", "Null", stEP07) & ",EP08=" & IIf(stEP08 = "", "Null", stEP08) & ",  EP34=" & CNULL(stEP34) & " Where EP02='" & stCP09 & "'; " & _
            " Update CASEPROGRESS Set CP48=" & IIf(stCP48 = "", "Null", stCP48) & ",CP64=" & CNULL(ChgSQL(txtCP64)) & " Where CP09='" & stCP09 & "';" & _
            " End;"
    
    cnnConnection.Execute strSql
    
cnnConnection.CommitTrans
FormSave = True

flgError:
    If Err.Number <> 0 Then
        cnnConnection.RollbackTrans
        MsgBox Err.Description, vbCritical
    End If

End Function

Private Function TxtValidate() As Boolean

   Dim bolCancel As Boolean
    
   bolCancel = False: bolIsValidate = False
   
   If txtCP48T.Enabled = True Then
      Call txtCP48T_Validate(bolCancel)
      If bolCancel Then GoTo flgFail
   End If
   
   If txtEP08T.Enabled = True Then
      Call txtEP08T_Validate(bolCancel)
      If bolCancel Then GoTo flgFail
   End If
   
   If txtEP07T.Enabled = True Then
      Call txtEP07T_Validate(bolCancel)
      If bolCancel Then GoTo flgFail
   End If
   
   If txtEP08T = "" And txtEP34 <> "N" Then
      txtEP34 = "N"
      If MsgBox("無核稿期限時是否會稿將自動上【N】，確定要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
         txtEP34 = "Y"   '2012/5/15 MODIFY BY SONIA 改上Y原為NULL
         If txtEP08T.Enabled = True Then txtEP08T.SetFocus
         GoTo flgFail
      End If
   End If
   If txtEP08T <> "" And txtEP34 = "N" Then
      If MsgBox("您有設核稿期限，確定不會稿？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
         If txtEP34.Enabled = True Then txtEP34.SetFocus
         GoTo flgFail
      End If
   End If
   TxtValidate = True
    
flgFail:
   bolIsValidate = True
    
End Function

Private Sub cmdExit_Click()
   Unload frm060108
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If TxtValidate Then
      'Add by Sindy 2021/11/25 檢查畫面上的物件是否含有Unicode文字
      If PUB_ChkUniText(Me, True, True) = False Then
         Exit Sub
      End If

      If FormSave() = True Then
         cmdBack_Click
      Else
         MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
      End If
   End If
End Sub

Private Sub Form_Activate()
   If txtEP08T <> "" Then
      txtEP07T.SetFocus
   End If
End Sub

Private Sub Form_Load()

    MoveFormToCenter Me
    bolIsValidate = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm060108_1 = Nothing
End Sub

Private Sub txtCP48T_GotFocus()
    TextInverse txtCP48T
    'edit by nickc 2007/07/11 切換輸入法改用API
    'txtCP48T.IMEMode = 2
    CloseIme
End Sub

Private Sub txtCP48T_Validate(Cancel As Boolean)
   If txtCP48T <> "" Then
      If Not ChkDate(txtCP48T) Then
        Cancel = True
        If Not bolIsValidate Then txtCP48T.SetFocus
        Call txtCP48T_GotFocus
      End If
   End If
End Sub

Private Sub txtCP64_Validate(Cancel As Boolean)
    Cancel = Not CheckLengthIsOK(txtCP64, 2000)
End Sub

Private Sub txtEP07T_GotFocus()
   TextInverse txtEP07T
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtEP07T.IMEMode = 2
   CloseIme
End Sub

Private Sub txtEP07T_Validate(Cancel As Boolean)
   If txtEP07T <> "" Then
      If Not ChkDate(txtEP07T) Then
         Cancel = True
         If Not bolIsValidate Then txtEP07T.SetFocus
         Call txtEP07T_GotFocus
       End If
   End If
End Sub

Private Sub txtEP08T_GotFocus()
   TextInverse txtEP08T
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtEP08T.IMEMode = 2
   CloseIme
End Sub

Private Sub txtEP08T_Validate(Cancel As Boolean)
   If txtEP08T <> "" Then
      If Not ChkDate(txtEP08T) Then
         Cancel = True
         If Not bolIsValidate Then txtEP08T.SetFocus
         Call txtEP08T_GotFocus
       End If
   End If
End Sub

Private Sub txtEP34_GotFocus()
   TextInverse txtEP34
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtEP34.IMEMode = 2
   CloseIme
End Sub

Private Sub txtEP34_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
End Sub
