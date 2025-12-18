VERSION 5.00
Begin VB.Form frm030301_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "事件說明"
   ClientHeight    =   6204
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8568
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6204
   ScaleWidth      =   8568
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   7500
      TabIndex        =   1
      Top             =   60
      Width           =   800
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   3500
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frm030301_2.frx":0000
      Top             =   210
      Width           =   7905
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   2412
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frm030301_2.frx":04AF
      Top             =   3750
      Width           =   8340
   End
   Begin VB.Label lblPS 
      Caption         =   "P.S. 有修改說明內容，請到TextHistory留下記錄"
      ForeColor       =   &H00FF00FF&
      Height          =   204
      Left            =   216
      TabIndex        =   3
      Top             =   24
      Width           =   3900
   End
End
Attribute VB_Name = "frm030301_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/15 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/2 日期欄已修改
Option Explicit
 
Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   
   'Added by Lydia 2025/05/06
   If Pub_StrUserSt03 <> "M51" Then
      lblPS.Visible = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm030301_2 = Nothing
End Sub

'Added by Lydia 2025/05/06 留下說明文字的修改記錄
Private Sub TextHistory()
Dim strText1 As String, strText2 As String

   strText1 = "T、FCT、S台灣案：" & vbCrLf
   strText1 = strText1 & "承　辦　組：　 １未收文：含CF案，本所期限＜＝系統日＋７個工作天之未收文案件。" & vbCrLf
   strText1 = strText1 & "（補優先權證明)２未收文：FCT，本所期限＜＝系統日＋３０日曆天之未收文案件。" & vbCrLf
   strText1 = strText1 & "（催 款）           ３未收文：本所期限＜＝系統日＋２個工作天之未收文案件。" & vbCrLf
   strText1 = strText1 & "（非外商人員） ４達本所：本所期限＜＝系統日＋２個工作天之未發文案件。" & vbCrLf
   strText1 = strText1 & "（外 商 人 員）  ５達本所：含CF案，本所期限＜＝系統日＋１個工作天之未發文案件。" & vbCrLf
   strText1 = strText1 & "（外 商 人 員）  ６達承辦：承辦期限＜＝系統日＋２個工作天之未發文案件。" & vbCrLf
   strText1 = strText1 & "（回代、催款） ７達承辦：承辦期限＜＝系統日＋１個工作天之未發文案件。" & vbCrLf
   strText1 = strText1 & "（審查報告）　 ８達承辦：FCT，承辦期限＜＝系統日＋４個工作天之未發文案件。" & vbCrLf
   strText1 = strText1 & "（外 商 人 員）  ９未請款：含CF案，發文日＜系統日＋３個工作天之未請款案件。" & vbCrLf
   strText1 = strText1 & "（內 商 人 員） １０未請款：發文日＜系統日＋１０個工作天之未請款案件。" & vbCrLf
   strText1 = strText1 & "註：未收文延展期限，英文組改為法定期限＋１個工作天＋7個月才開始提醒，日文組維持原規則。" & vbCrLf
   strText1 = strText1 & "FCT、S台灣案：" & vbCrLf
   strText1 = strText1 & "程　序　組：　１達本所：本所期限＜＝系統日＋２個工作天之未發文案件。" & vbCrLf
   strText1 = strText1 & "　　　　　　　２未收文：本所期限＜＝系統日＋１個工作天之未收文案件。" & vbCrLf
   strText1 = strText1 & "　　　　　　　　　　　　(109/7月起,程序不管制未收文延展期限)" & vbCrLf
   strText1 = strText1 & "　　　　　　　３今送件：承辦期限＜＝系統日之未發文案件。" & vbCrLf
   strText1 = strText1 & "　　　　　　　４需請款：本所期限＜＝系統日之未請款案件。" & vbCrLf
'-------------------------------------------------------
   strText2 = "CFT、CFC、S、TF非台灣案：" & vbCrLf
   strText2 = strText2 & "　　　　　　　１達法定：法定期限＜＝系統日＋５個工作天之未發文案件。" & vbCrLf
   'Modified by Lydia 2025/05/06
   'strText2 = strText2 & "　　　　　　　２未收文：法定期限＜＝系統日＋５或６個工作天之未收文案件。" & vbCrLf
   strText2 = strText2 & "　　　　　　　２未收文：法定期限＜＝系統日＋６個工作天／本所期限當日起之未收文案件。" & vbCrLf
   strText2 = strText2 & "　　　　　　　３達本所：本所期限＜＝系統日＋１個工作天之未發文案件。" & vbCrLf
   strText2 = strText2 & "　　　　　　　４達承辦：承辦期限＜＝系統日＋１個工作天之未發文案件。" & vbCrLf
   strText2 = strText2 & "　　　　　　　５達指定：指定日期＜＝系統日＋２個工作天之未發文案件。" & vbCrLf
   'Modified by Lydia 2025/05/06
   'strText2 = strText2 & "　　　　　　　６延展(102)/使用宣誓(105)未收文：法定期限＜＝系統日+30日曆天之未收文案件。" & vbCrLf
   strText2 = strText2 & "　　　　　　　６延展(102)/使用宣誓(105)未收文：系統日+第21,22個工作天之未收文案件。" & vbCrLf
   strText2 = strText2 & "　　　　　　　７催審(305)/收達(997)/提申(998)：本所期限＜＝系統日" & vbCrLf
   strText2 = strText2 & "未分案：１已收文未發文且無承辦人且（本所期限＜＝系統日＋２個工作天）" & vbCrLf
   strText2 = strText2 & "　　　　２承辦期限＜＝系統日＋２個工作天之未發文案件。（第四級主管才會顯示）" & vbCrLf
   strText2 = strText2 & "未發文：已收文未發文且無期限或期限未達管制日期之案件(含未分案)。（按所有未發" & vbCrLf
   strText2 = strText2 & "　　　　文按鈕才會顯示）" & vbCrLf
   strText2 = strText2 & "T、FCT台灣商標爭議案逾承辦期限、逾 指定會稿日。" & vbCrLf

End Sub


