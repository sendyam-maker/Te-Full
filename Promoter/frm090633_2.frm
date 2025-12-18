VERSION 5.00
Begin VB.Form frm090633_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "確認及選擇畫面"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5925
   Begin VB.Frame Frame3 
      Height          =   705
      Left            =   90
      TabIndex        =   10
      Top             =   2250
      Width           =   5730
      Begin VB.CheckBox Check6 
         Caption         =   "刪除該筆內部收文"
         Height          =   255
         Left            =   3780
         TabIndex        =   12
         Top             =   263
         Width           =   1860
      End
      Begin VB.CheckBox Check5 
         Caption         =   "內部收文輸入發文日"
         Height          =   255
         Left            =   1665
         TabIndex        =   11
         Top             =   263
         Width           =   2040
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "請選擇處理方式："
         Height          =   180
         Left            =   135
         TabIndex        =   13
         Top             =   300
         Width           =   1440
      End
   End
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   90
      TabIndex        =   6
      Top             =   1410
      Width           =   5730
      Begin VB.CheckBox Check3 
         Caption         =   "收文列管及通知智權同仁"
         Height          =   255
         Left            =   1665
         TabIndex        =   8
         Top             =   263
         Width           =   2355
      End
      Begin VB.CheckBox Check4 
         Caption         =   "不列管"
         Height          =   255
         Left            =   4230
         TabIndex        =   7
         Top             =   263
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "請選擇處理方式："
         Height          =   180
         Left            =   135
         TabIndex        =   9
         Top             =   300
         Width           =   1440
      End
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   90
      TabIndex        =   2
      Top             =   570
      Width           =   5730
      Begin VB.CheckBox Check2 
         Caption         =   "未呈報"
         Height          =   255
         Left            =   4185
         TabIndex        =   5
         Top             =   270
         Width           =   1005
      End
      Begin VB.CheckBox Check1 
         Caption         =   "已呈報"
         Height          =   255
         Left            =   2925
         TabIndex        =   4
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否已事先呈報地區主管認可"
         Height          =   180
         Left            =   135
         TabIndex        =   3
         Top             =   307
         Width           =   2340
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "取消"
      Height          =   375
      Index           =   1
      Left            =   4950
      TabIndex        =   1
      Top             =   75
      Width           =   870
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   4005
      TabIndex        =   0
      Top             =   75
      Width           =   870
   End
   Begin VB.Label lblAlert 
      AutoSize        =   -1  'True
      Caption         =   "lblAlert"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   180
      TabIndex        =   14
      Top             =   150
      Width           =   660
   End
End
Attribute VB_Name = "frm090633_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/01/03 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/10 智權人員欄已修改
Option Explicit
Public p_Choice As Integer '選項:1 時數,2 核可,3 管制
Public p_Parent As Form

Private Sub Check1_Click()
   If Check1.Value = 1 Then
      Check2.Value = 0
   End If
End Sub

Private Sub Check2_Click()
   If Check2.Value = 1 Then
      Check1.Value = 0
   End If
End Sub

Private Sub Check3_Click()
   If Check3.Value = 1 Then
      Check4.Value = 0
   End If
End Sub

Private Sub Check4_Click()
   If Check4.Value = 1 Then
      Check3.Value = 0
   End If
End Sub

Private Sub Check5_Click()
   If Check5.Value = 1 Then
      Check6.Value = 0
   End If
End Sub

Private Sub Check6_Click()
   If Check6.Value = 1 Then
      Check5.Value = 0
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
   p_Parent.p_iRtn = 0
   If Index = 0 Then
      Select Case p_Choice
         Case 1
            If Check1.Value = 1 Then
               p_Parent.p_iRtn = 1
               
            ElseIf Check2.Value = 1 Then
               p_Parent.p_iRtn = 2
            End If
            
         Case 2
            If Check3.Value = 1 Then
               p_Parent.p_iRtn = 1
               
            ElseIf Check4.Value = 1 Then
               p_Parent.p_iRtn = 2
            End If
            
         Case 3
            If Check5.Value = 1 Then
               p_Parent.p_iRtn = 1
               
            ElseIf Check6.Value = 1 Then
               p_Parent.p_iRtn = 2
            End If
            
      End Select
      If p_Parent.p_iRtn = 0 Then
         MsgBox "尚未勾選不可確認！"
         Exit Sub
      End If
   End If
   Unload Me
End Sub

Private Sub Form_Load()
   Select Case p_Choice
      Case 2
         Frame2.Top = Frame1.Top
         Frame2.ZOrder 0
         'cmdok(1).Visible = False
      Case 3
         Frame3.Top = Frame1.Top
         Frame3.ZOrder 0
         'cmdok(1).Visible = False
   End Select
   Me.Height = 1770
   MoveFormToCenter Me
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090633_2 = Nothing
End Sub
