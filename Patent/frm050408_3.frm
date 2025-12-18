VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050408_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "互惠代理人案件統計表(帳款統計)"
   ClientHeight    =   3888
   ClientLeft      =   168
   ClientTop       =   960
   ClientWidth     =   5304
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3888
   ScaleWidth      =   5304
   Begin VB.TextBox Text7 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   2250
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3140
      Width           =   1410
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   2250
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1900
      Width           =   1410
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   2250
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3450
      Width           =   1410
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   2250
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2520
      Width           =   1410
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   2250
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2830
      Width           =   1410
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   2250
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2210
      Width           =   1410
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   2250
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1590
      Width           =   1410
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3915
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   120
      Width           =   1200
   End
   Begin MSForms.Label lblContact 
      Height          =   300
      Left            =   1170
      TabIndex        =   24
      Top             =   645
      Width           =   2500
      VariousPropertyBits=   27
      Size            =   "4410;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblAgentName 
      Height          =   300
      Left            =   1170
      TabIndex        =   23
      Top             =   360
      Width           =   2500
      VariousPropertyBits=   27
      Size            =   "4410;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "CFP盈虧:"
      Height          =   255
      Left            =   1485
      TabIndex        =   22
      Top             =   3165
      Width           =   705
   End
   Begin VB.Label Label11 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "FCP請款服務費:"
      Height          =   255
      Left            =   945
      TabIndex        =   20
      Top             =   1905
      Width           =   1245
   End
   Begin VB.Label Label8 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "CFP應付金額:"
      Height          =   255
      Left            =   1125
      TabIndex        =   16
      Top             =   2535
      Width           =   1065
   End
   Begin VB.Label Label10 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "FCP請款-CFP應付:"
      Height          =   255
      Left            =   765
      TabIndex        =   12
      Top             =   3480
      Width           =   1425
   End
   Begin VB.Label Label9 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "CFP付款金額:"
      Height          =   255
      Left            =   1125
      TabIndex        =   11
      Top             =   2850
      Width           =   1065
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "FCP收款金額:"
      Height          =   255
      Left            =   1125
      TabIndex        =   10
      Top             =   2220
      Width           =   1065
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "FCP請款金額:"
      Height          =   255
      Left            =   1125
      TabIndex        =   9
      Top             =   1590
      Width           =   1065
   End
   Begin VB.Label lblCondition 
      Height          =   255
      Left            =   1170
      TabIndex        =   8
      Top             =   1170
      Width           =   2505
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "資料區間:"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1170
      Width           =   765
   End
   Begin VB.Label lblNation 
      Height          =   255
      Left            =   1170
      TabIndex        =   6
      Top             =   930
      Width           =   2500
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "國籍:"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   906
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人:"
      Height          =   255
      Left            =   540
      TabIndex        =   4
      Top             =   644
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "代理人名稱:"
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   382
      Width           =   945
   End
   Begin VB.Label lblAgentNo 
      Height          =   255
      Left            =   1170
      TabIndex        =   2
      Top             =   120
      Width           =   2500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "代理人代號:"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "frm050408_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/11 改成Form2.0 ; lblAgentName、lblContact
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'Create by Morgan 2008/3/11
Option Explicit

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   'Added by Lydia 2025/05/22
   If frm050408.Tag <> "" Then
      Me.Caption = "互惠期間統計表(帳款統計)"
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm050408_3 = Nothing
End Sub
