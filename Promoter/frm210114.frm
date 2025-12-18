VERSION 5.00
Begin VB.Form frm210114 
   BorderStyle     =   1  '單線固定
   Caption         =   "委任契約書"
   ClientHeight    =   1560
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4540
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3135
      TabIndex        =   2
      Top             =   150
      Width           =   1020
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2010
      TabIndex        =   1
      Top             =   150
      Width           =   1020
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm210114.frx":0000
      Left            =   1455
      List            =   "frm210114.frx":0019
      Style           =   2  '單純下拉式
      TabIndex        =   0
      Top             =   795
      Width           =   2880
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統別："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   165
      TabIndex        =   3
      Top             =   780
      Width           =   1200
   End
End
Attribute VB_Name = "frm210114"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/11 Form2.0已檢查 (無需修改的物件)
'Memo by Lydia 2019/07/01 表單名稱:案件委任契約書=>委任契約書
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
    Select Case Trim(Combo1.Text)
    Case "P"
         Me.Hide
         frm210114_1.Show
    'Modified by Lydia 2020/07/20 改為:CFP及P非台灣案
    'Case "CFP 及 P"
    Case "CFP及P非台灣案"
         Me.Hide
         frm210114_2.Show
    Case "T"
         Me.Hide
         frm210114_3.Show
    Case "CFT 及 T大陸案 及 TF馬德里案"
         Me.Hide
         frm210114_4.Show
'cancel by sonia 2022/11/11 杜協理通知關閉
'    'add by nickc 2007/11/14
'    Case "常年顧問聘任書"
'         Me.Hide
'         frm210114_5.Show
'end 2022/11/11
    'add by nickc 2007/11/14
    Case "條碼案件委任契約書"
         frm210114_6.Show
    'Added by Lydia 2022/04/15
    Case "著作權案件委任契約書"
         'Memo by Lydia 2022/04/20 版面已調整
         'Memo by Lydia 2022/04/18 著作權案件委任契約書: 因為尚未有問題,所以先從Combo1清單移除
         frm210114_7.Show
    'Added by Lydia 2022/04/26
    Case "專利申請案保密同意書"
         frm210114_8.Show
    Case Else
    End Select
Case 1
    Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set frm210114 = Nothing
End Sub
