VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060104_f 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專發文-第三人提實審"
   ClientHeight    =   3795
   ClientLeft      =   270
   ClientTop       =   960
   ClientWidth     =   8760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   8760
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   9
      Top             =   540
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1620
      MaxLength       =   6
      TabIndex        =   8
      Top             =   540
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2460
      MaxLength       =   1
      TabIndex        =   7
      Top             =   540
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2700
      MaxLength       =   2
      TabIndex        =   6
      Top             =   540
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7404
      TabIndex        =   4
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6576
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   2
      Left            =   1530
      TabIndex        =   2
      Top             =   3345
      Width           =   7095
      VariousPropertyBits=   671105051
      MaxLength       =   160
      Size            =   "12515;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   1
      Left            =   1530
      TabIndex        =   1
      Top             =   3000
      Width           =   7095
      VariousPropertyBits=   671105051
      MaxLength       =   160
      Size            =   "12515;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   0
      Left            =   1530
      TabIndex        =   0
      Top             =   2655
      Width           =   7095
      VariousPropertyBits=   671105051
      MaxLength       =   160
      Size            =   "12515;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1140
      TabIndex        =   5
      Top             =   840
      Width           =   7455
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "13150;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "                       (日):"
      Height          =   180
      Index           =   12
      Left            =   120
      TabIndex        =   30
      Top             =   3360
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "                       (英):"
      Height          =   180
      Index           =   11
      Left            =   120
      TabIndex        =   29
      Top             =   3030
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "第三人名稱: (中):"
      Height          =   180
      Index           =   10
      Left            =   180
      TabIndex        =   28
      Top             =   2700
      Width           =   1335
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   3
      Left            =   4560
      TabIndex        =   27
      Top             =   1830
      Width           =   4020
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "7091;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   4
      Left            =   1140
      TabIndex        =   26
      Top             =   2160
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3757;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   0
      Left            =   1140
      TabIndex        =   25
      Top             =   1170
      Width           =   2130
      VariousPropertyBits=   27
      Size            =   "3757;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   5
      Left            =   4560
      TabIndex        =   24
      Top             =   2160
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3757;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   9
      Left            =   3645
      TabIndex        =   23
      Top             =   2160
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "本所期限:"
      Height          =   180
      Index           =   8
      Left            =   180
      TabIndex        =   22
      Top             =   2160
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Index           =   7
      Left            =   3645
      TabIndex        =   21
      Top             =   1830
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請日:"
      Height          =   180
      Index           =   3
      Left            =   3660
      TabIndex        =   20
      Top             =   1170
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   19
      Top             =   840
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   18
      Top             =   1170
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   17
      Top             =   540
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Index           =   6
      Left            =   180
      TabIndex        =   16
      Top             =   1830
      Width           =   765
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   2
      Left            =   1140
      TabIndex        =   15
      Top             =   1830
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3757;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   1
      Left            =   4560
      TabIndex        =   14
      Top             =   1170
      Width           =   2130
      VariousPropertyBits=   27
      Size            =   "3757;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   13
      Top             =   1500
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文日:"
      Height          =   180
      Index           =   5
      Left            =   3660
      TabIndex        =   12
      Top             =   1500
      Width           =   585
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   0
      Left            =   1140
      TabIndex        =   11
      Top             =   1500
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3757;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   1
      Left            =   4560
      TabIndex        =   10
      Top             =   1500
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3757;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   8520
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   8520
      Y1              =   2580
      Y2              =   2580
   End
End
Attribute VB_Name = "frm060104_f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/18 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
'2007/8/6 ADD BY SONIA
Option Explicit

Dim pa() As String
Dim intWhere As Integer
Dim m_CP09 As String


Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm060104_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      m_CP09 = .Tag
      Label3(0) = m_CP09
   End With
   ReDim pa(TF_PA)
   ReadPatent
   Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060104_f = Nothing
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         frm060104_3.m_CP50 = Text7(0)
         frm060104_3.m_CP51 = Text7(1)
         frm060104_3.m_CP52 = Text7(2)
      Case 1
         frm060104_3.Show
   End Select
   Unload Me
End Sub

Private Sub ReadPatent()
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   Select Case pa(1)
      Case "FCP"
         If ClsPDReadPatentDatabase(pa(), intWhere) Then
            Label2(0) = pa(11)
            Label2(1) = pa(10)
            AddCboName Combo1, pa(5), pa(6), pa(7)
         End If
      Case "FG"
         If PUB_ReadServicePracticeDatabase(pa(), intWhere) Then
            Label2(0) = pa(11)
            Label2(1) = pa(10)
            AddCboName Combo1, pa(5), pa(6), pa(7)
         End If
   End Select
   strExc(0) = "select cp05,cp10,cpm03,cp08,cp06,cp07,cp14,cp113,cp64,st02" & _
      " from caseprogress,casepropertymap,staff where cp09='" & Label3(0) & "'" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp14"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
   With RsTemp
      Label3(1) = .Fields("cp05") - 19110000
      Label3(2) = .Fields("cp10") & " " & .Fields("cpm03")
      If Not IsNull(.Fields("cp08")) Then
         Label3(3) = .Fields("cp08")
      End If
      If Not IsNull(.Fields("cp06")) Then
         Label3(4) = .Fields("cp06") - 19110000
      End If
      If Not IsNull(.Fields("cp07")) Then
         Label3(5) = .Fields("cp07") - 19110000
      End If
   End With
   End If
End Sub
