VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090207_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利案例資料查詢"
   ClientHeight    =   4410
   ClientLeft      =   150
   ClientTop       =   1290
   ClientWidth     =   6060
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6060
   Begin VB.CommandButton Command1 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5184
      TabIndex        =   0
      Top             =   36
      Width           =   756
   End
   Begin VB.Label Label10 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "文書日期："
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   1860
      Width           =   900
   End
   Begin MSForms.Label LBL1 
      Height          =   255
      Index           =   9
      Left            =   1080
      TabIndex        =   19
      Top             =   1830
      Width           =   1485
      VariousPropertyBits=   27
      Size            =   "2619;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "文書類型："
      Height          =   180
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   900
   End
   Begin MSForms.Label LBL1 
      Height          =   255
      Index           =   8
      Left            =   1080
      TabIndex        =   17
      Top             =   1547
      Width           =   800
      VariousPropertyBits=   27
      Size            =   "1411;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LBL1 
      Height          =   975
      Index           =   7
      Left            =   1080
      TabIndex        =   16
      Top             =   3360
      Width           =   4935
      Size            =   "8705;1720"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LBL1 
      Height          =   465
      Index           =   6
      Left            =   1080
      TabIndex        =   15
      Top             =   2869
      Width           =   4935
      Size            =   "8705;820"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LBL1 
      Height          =   465
      Index           =   5
      Left            =   1080
      TabIndex        =   14
      Top             =   2381
      Width           =   4935
      Size            =   "8705;820"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LBL1 
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   13
      Top             =   2103
      Width           =   4935
      VariousPropertyBits=   27
      Size            =   "8705;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LBL1 
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   12
      Top             =   1269
      Width           =   800
      VariousPropertyBits=   27
      Size            =   "1411;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LBL1 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   11
      Top             =   991
      Width           =   800
      VariousPropertyBits=   27
      Size            =   "1411;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LBL1 
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   10
      Top             =   713
      Width           =   800
      VariousPropertyBits=   27
      Size            =   "1411;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LBL1 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   9
      Top             =   435
      Width           =   495
      VariousPropertyBits=   27
      Size            =   "873;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "案情摘要："
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   3390
      Width           =   900
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "案例字號："
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   900
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "主旨："
      Height          =   180
      Left            =   480
      TabIndex        =   6
      Top             =   2430
      Width           =   540
   End
   Begin VB.Label Label5 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   2130
      Width           =   900
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "備用類："
      Height          =   180
      Left            =   300
      TabIndex        =   4
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "次次類："
      Height          =   180
      Left            =   300
      TabIndex        =   3
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "次類："
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   750
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "主類："
      Height          =   180
      Left            =   480
      TabIndex        =   1
      Top             =   435
      Width           =   540
   End
End
Attribute VB_Name = "frm090207_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/26 改成Form2.0 ; LBL1(index)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
'Option Explicit
'Dim pemain As New ADODB.Recordset

Private Sub Command1_Click()
frm090207_2.Show
Unload Me
End Sub

Private Sub Form_Load()
'If pemain.State = adStateOpen Then pemain.Close
'pemain.CursorLocation = adUseClient
MoveFormToCenter Me
'Label9.Caption = Mid(frm090207_2.A, 1, 3)
'Label10.Caption = Mid(frm090207_2.A, 5, 2)
'Label11.Caption = Mid(frm090207_2.A, 8, 2)
'Label12.Caption = Mid(frm090207_2.A, 11, 2)
'strExc(0) = "select pc05||'-'||pc06||'-'||pc07||'-'||pc08,pc09,pc10,pc11 from patentcase where pc01='" & Label9.Caption & "' and pc02='" & Label10.Caption & "' and pc03='" & Label11.Caption & "' and pc04='" & Label12.Caption & "'"
'pemain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
'If pemain.BOF And pemain.EOF Then pemain.Close
'If IsNull(pemain.Fields(0).Value) Then
'    Label13.Caption = ""
'Else
'    Label13.Caption = pemain.Fields(0).Value
'End If
'If IsNull(pemain.Fields(1).Value) Then
'    Label14.Caption = ""
'Else
'    Label14.Caption = pemain.Fields(1).Value
'End If
'If IsNull(pemain.Fields(2).Value) Then
'    Label15.Caption = ""
'Else
'    Label15.Caption = pemain.Fields(2).Value
'End If
'If IsNull(pemain.Fields(3).Value) Then
'    Label16.Caption = ""
'Else
'    Label16.Caption = pemain.Fields(3).Value
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090207_3 = Nothing
End Sub
