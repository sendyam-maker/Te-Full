VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090201_b_2 
   BackColor       =   &H80000004&
   BorderStyle     =   1  '單線固定
   Caption         =   "資料記錄明細"
   ClientHeight    =   5745
   ClientLeft      =   1440
   ClientTop       =   2310
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   Begin VB.CommandButton CmdOk1 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   0
      Left            =   8040
      TabIndex        =   8
      Top             =   45
      Width           =   800
   End
   Begin MSForms.TextBox txtTCD 
      Height          =   300
      Index           =   6
      Left            =   1485
      TabIndex        =   5
      Top             =   2448
      Width           =   1092
      VariousPropertyBits=   671105055
      BackColor       =   -2147483633
      MaxLength       =   80
      Size            =   "1926;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtTCD 
      Height          =   300
      Index           =   7
      Left            =   1485
      TabIndex        =   6
      Top             =   2775
      Width           =   3735
      VariousPropertyBits=   671105055
      BackColor       =   -2147483633
      MaxLength       =   20
      Size            =   "6588;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtTCD 
      Height          =   1605
      Index           =   8
      Left            =   1485
      TabIndex        =   7
      Top             =   3150
      Width           =   7245
      VariousPropertyBits=   -1466941409
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "12779;2831"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtTCD 
      Height          =   300
      Index           =   1
      Left            =   1485
      TabIndex        =   0
      Top             =   1140
      Width           =   1092
      VariousPropertyBits=   671105055
      BackColor       =   -2147483633
      MaxLength       =   9
      Size            =   "1926;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtTCD 
      Height          =   300
      Index           =   4
      Left            =   1485
      TabIndex        =   3
      Top             =   2121
      Width           =   1092
      VariousPropertyBits=   671105055
      BackColor       =   -2147483633
      MaxLength       =   80
      Size            =   "1926;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtTCD 
      Height          =   300
      Index           =   3
      Left            =   1485
      TabIndex        =   2
      Top             =   1794
      Width           =   705
      VariousPropertyBits=   671105055
      BackColor       =   -2147483633
      MaxLength       =   30
      Size            =   "1244;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtTCD 
      Height          =   300
      Index           =   2
      Left            =   1485
      TabIndex        =   1
      Top             =   1467
      Width           =   480
      VariousPropertyBits=   671105055
      BackColor       =   -2147483633
      MaxLength       =   1
      Size            =   "847;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtTCD 
      Height          =   300
      Index           =   5
      Left            =   2610
      TabIndex        =   4
      Top             =   2115
      Width           =   1092
      VariousPropertyBits=   671105055
      BackColor       =   -2147483633
      MaxLength       =   80
      Size            =   "1926;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   0
      Left            =   2250
      TabIndex        =   17
      Top             =   1800
      Width           =   1095
      BackColor       =   -2147483644
      VariousPropertyBits=   27
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "(1.齊備日  2.業務補充資料日  3.承辦期限  4.電腦中心  5.通知補充資料)"
      Height          =   210
      Left            =   2220
      TabIndex        =   16
      Top             =   1467
      Width           =   5625
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "輸入日期："
      Height          =   180
      Index           =   3
      Left            =   570
      TabIndex        =   15
      Top             =   2454
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "異動類別："
      Height          =   180
      Index           =   6
      Left            =   555
      TabIndex        =   14
      Top             =   1491
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "操作時間："
      Height          =   180
      Index           =   1
      Left            =   555
      TabIndex        =   13
      Top             =   2133
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "操作人員："
      Height          =   180
      Index           =   2
      Left            =   555
      TabIndex        =   12
      Top             =   1812
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "總收文號："
      Height          =   210
      Index           =   0
      Left            =   540
      TabIndex        =   11
      Top             =   1140
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "通知補充資料："
      Height          =   180
      Index           =   19
      Left            =   195
      TabIndex        =   10
      Top             =   3210
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "操作動作："
      Height          =   180
      Index           =   5
      Left            =   555
      TabIndex        =   9
      Top             =   2775
      Width           =   900
   End
End
Attribute VB_Name = "frm090201_b_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/23 改成Form2.0 ; txtTCD(index)、lbl1(index)
'Create By Sindy 2012/12/14
Option Explicit

Public cmdState As Integer
Dim oText As Object
Dim idx As Integer


Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090201_b_2 = Nothing
End Sub

Private Sub cmdok1_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
         Unload Me
   End Select
End Sub

Function StrMenu(strKey1 As String, StrKey2 As String, strKey3 As String) As Boolean
   strExc(0) = "select * from tmctldate where TCD01='" & strKey1 & "' and TCD04='" & StrKey2 & "' and TCD05='" & strKey3 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      StrMenu = True
      ShowRecord RsTemp
   Else
      StrMenu = False
      Unload Me
      Exit Function
   End If
End Function

' 將資料庫中的資料更新到所有欄位中
Private Sub ShowRecord(ByRef p_Rst As ADODB.Recordset)
   Dim rsTCD As ADODB.Recordset
   
   ClearField
   Set rsTCD = p_Rst.Clone
   With rsTCD
      If .RecordCount > 0 Then
         For Each oText In txtTCD
            idx = oText.Index
            oText.Text = "" & .Fields("TCD" & Format(idx, "0#"))
         Next
         If Trim(txtTCD(4)) <> "" Then txtTCD(4) = ChangeWStringToTDateString(txtTCD(4))
         If Trim(txtTCD(5)) <> "" Then txtTCD(5) = Format(txtTCD(5), "##:##:##")
         If Trim(txtTCD(6)) <> "" Then txtTCD(6) = ChangeWStringToTDateString(txtTCD(6))
         '操作人員
         lbl1(0).Caption = ""
         If Not IsNull(.Fields("TCD03")) Then
            lbl1(0).Caption = GetPrjSalesNM(.Fields("TCD03"))
         End If
      End If
   End With
End Sub

Private Sub ClearField()
   Dim oLabel As Object
   For Each oText In txtTCD
      oText.Text = Empty
   Next
   For Each oLabel In lbl1
      oLabel.Caption = Empty
   Next
End Sub
