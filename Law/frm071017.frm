VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm071017 
   BorderStyle     =   1  '單線固定
   Caption         =   "會稿日輸入"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   8955
   Begin VB.TextBox txtEP07 
      Height          =   285
      Left            =   1170
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2958
      Width           =   1305
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面"
      Height          =   400
      Left            =   7575
      TabIndex        =   3
      Top             =   90
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Height          =   400
      Left            =   6750
      TabIndex        =   2
      Top             =   90
      Width           =   800
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   1
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   270
      Width           =   495
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   2
      Left            =   1650
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   270
      Width           =   855
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   3
      Left            =   2490
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   270
      Width           =   255
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   4
      Left            =   2730
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   270
      Width           =   375
   End
   Begin MSForms.Label lblCP29T 
      Height          =   285
      Left            =   6990
      TabIndex        =   36
      Top             =   2616
      Width           =   1485
      VariousPropertyBits=   27
      Size            =   "2619;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCP13T 
      Height          =   285
      Left            =   6990
      TabIndex        =   35
      Top             =   2274
      Width           =   1485
      VariousPropertyBits=   27
      Size            =   "2619;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCP14T 
      Height          =   285
      Left            =   2070
      TabIndex        =   34
      Top             =   2616
      Width           =   1485
      VariousPropertyBits=   27
      Size            =   "2619;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblCustName 
      Height          =   285
      Left            =   2460
      TabIndex        =   33
      Top             =   1590
      Width           =   6135
      VariousPropertyBits=   27
      Size            =   "10821;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   3
      Left            =   1530
      TabIndex        =   31
      Top             =   1215
      Width           =   6915
      VariousPropertyBits=   27
      Size            =   "12197;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   2
      Left            =   1530
      TabIndex        =   30
      Top             =   915
      Width           =   6915
      VariousPropertyBits=   27
      Size            =   "12197;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   1
      Left            =   1530
      TabIndex        =   29
      Top             =   600
      Width           =   6915
      VariousPropertyBits=   27
      Size            =   "12197;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCP64 
      Height          =   525
      Left            =   1170
      TabIndex        =   1
      Top             =   3300
      Width           =   7590
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13388;926"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LblCustNo 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   1170
      TabIndex        =   28
      Top             =   1590
      Width           =   1245
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "當事人"
      Height          =   180
      Left            =   240
      TabIndex        =   27
      Top             =   1642
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Index           =   5
      Left            =   5160
      TabIndex        =   26
      Top             =   2326
      Width           =   765
   End
   Begin VB.Label lblCP13 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   6105
      TabIndex        =   25
      Top             =   2274
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "協辦人員:"
      Height          =   180
      Index           =   1
      Left            =   5160
      TabIndex        =   24
      Top             =   2668
      Width           =   765
   End
   Begin VB.Label lblCP29 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   6105
      TabIndex        =   23
      Top             =   2616
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Index           =   10
      Left            =   225
      TabIndex        =   22
      Top             =   3300
      Width           =   765
   End
   Begin VB.Label lblCP14 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   1170
      TabIndex        =   21
      Top             =   2616
      Width           =   855
   End
   Begin VB.Label lblCP09 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   6120
      TabIndex        =   20
      Top             =   1932
      Width           =   1665
   End
   Begin VB.Label lblCP10T 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   1170
      TabIndex        =   19
      Top             =   2274
      Width           =   3225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "總收文號:"
      Height          =   180
      Index           =   8
      Left            =   5160
      TabIndex        =   18
      Top             =   1984
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "會稿日:"
      Height          =   180
      Index           =   7
      Left            =   240
      TabIndex        =   17
      Top             =   3010
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Index           =   4
      Left            =   225
      TabIndex        =   16
      Top             =   2668
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Index           =   3
      Left            =   225
      TabIndex        =   15
      Top             =   2326
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文日:"
      Height          =   180
      Index           =   2
      Left            =   225
      TabIndex        =   14
      Top             =   1984
      Width           =   585
   End
   Begin VB.Label lblCP05T 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   1170
      TabIndex        =   13
      Top             =   1932
      Width           =   1665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱"
      Height          =   180
      Left            =   225
      TabIndex        =   12
      Top             =   652
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "(中):"
      Height          =   180
      Left            =   1005
      TabIndex        =   11
      Top             =   652
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(英):"
      Height          =   180
      Left            =   1005
      TabIndex        =   10
      Top             =   967
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(日):"
      Height          =   180
      Left            =   1005
      TabIndex        =   9
      Top             =   1267
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   225
      TabIndex        =   8
      Top             =   270
      Width           =   765
   End
End
Attribute VB_Name = "frm071017"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/14 改成Form2.0 ; lblCaseName(index)、txtCP64、LblCustName、lblCP13T、lblCP14T、lblCP29T ; 將「確定」按鈕的defalut拿掉，並且拿掉快速鍵功能；
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim bolIsValidate As Boolean, stST03 As String, stCP10 As String, stCP06 As String

Private Sub cmdBack_Click()
    Call frm071016.SetGrid(False)
    frm071016.Show
    Unload Me
End Sub

Public Sub SetData(ByRef rstGrid As ADODB.Recordset, ByVal iRow As Integer)
    
    Dim ii As Integer, stPA08 As String
    
    With frm071016
        For ii = 1 To 4
            txtCaseNo(ii) = .txtCaseNo(ii)
        Next ii
        For ii = 1 To 3
            lblCaseName(ii) = .lblCaseName(ii)
        Next ii
        LblCustNo = .LblCustNo
        LblCustName = .LblCustName
    End With
    With rstGrid
        .Move iRow - 1, adBookmarkFirst
        lblCP05T = "" & .Fields("CP05T")
        lblCP09 = "" & .Fields("CP09")
        lblCP10T = "" & .Fields("CP10T")
        lblCP14 = "" & .Fields("CP14")
        lblCP14T = "" & .Fields("CP14T")
        lblCP13 = "" & .Fields("CP13")
        lblCP13T = "" & .Fields("CP13T")
        lblCP29 = "" & .Fields("CP29")
        lblCP29T = "" & .Fields("CP29T")
        
        txtEP07 = "" & .Fields("EP07T")
        txtCP64 = "" & .Fields("CP64")
        
    End With
    
End Sub

Private Function FormSave() As Boolean

    Dim stSQL As String
    
On Error GoTo flgError

cnnConnection.BeginTrans
    
    stSQL = " Begin" & _
            " Update ENGINEERPROGRESS Set EP07=" & ChangeTStringToWString(txtEP07) & " Where EP02='" & lblCP09 & "'; " & _
            " Update CASEPROGRESS Set CP64='" & ChgSQL(txtCP64) & "' Where CP09='" & lblCP09 & "';" & _
            " End;"
    
    cnnConnection.Execute stSQL
    
cnnConnection.CommitTrans
FormSave = True

flgError:

    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical
        cnnConnection.RollbackTrans
    End If

End Function

Private Function TxtValidate() As Boolean

   Dim bolCancel As Boolean
    
   bolCancel = False: bolIsValidate = False
   
   If txtEP07 = "" Then
      bolCancel = True
      MsgBox "會稿日不可空白！", vbCritical
      txtEP07.SetFocus
   Else
      Call txtEP07_Validate(bolCancel)
   End If
   
    'Added by Lydia 2021/09/14 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
         bolCancel = True
    End If

   If bolCancel Then GoTo flgFail
   
   TxtValidate = True
    
flgFail:
    bolIsValidate = True
    
End Function

Private Sub cmdok_Click()
    
    If TxtValidate Then
        If FormSave() = True Then
            cmdBack_Click
        Else
            MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
        End If
    End If
    
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    bolIsValidate = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm071017 = Nothing
End Sub

Private Sub txtEP07_GotFocus()
    TextInverse txtEP07
End Sub

Private Sub txtEP07_Validate(Cancel As Boolean)

   If txtEP07 <> "" Then
      If ChkDate(txtEP07) = False Then
         txtEP07.SetFocus
         Call txtEP07_GotFocus
         Cancel = True
      End If
   End If
   
End Sub

Private Sub txtCP64_Validate(Cancel As Boolean)
    Cancel = Not CheckLengthIsOK(txtCP64, 2000)
End Sub

Private Sub txtCP64_GotFocus()
    TextInverse txtCP64
End Sub
