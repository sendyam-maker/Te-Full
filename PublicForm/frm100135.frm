VERSION 5.00
Begin VB.Form frm100135 
   BorderStyle     =   1  '單線固定
   Caption         =   "臺灣地址格式說明畫面"
   ClientHeight    =   2130
   ClientLeft      =   2565
   ClientTop       =   3105
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4695
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   3720
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   60
      Width           =   900
   End
   Begin VB.Label Label5 
      Caption         =   "● 釣魚臺列嶼"
      Height          =   300
      Left            =   0
      TabIndex        =   7
      Top             =   1680
      Width           =   4380
   End
   Begin VB.Label Label4 
      Caption         =   "● 其他縣市格式：XX 縣 XX 市(鄉)(鎮)"
      Height          =   300
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   4380
   End
   Begin VB.Label Label1 
      Caption         =   "例如：「台北」改輸入「臺北」"
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   2
      Left            =   1080
      TabIndex        =   5
      Top             =   360
      Width           =   2595
   End
   Begin VB.Label Label1 
      Caption         =   "縣市名稱之「台」'請改輸「臺」"
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   1
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   2600
   End
   Begin VB.Label Label1 
      Caption         =   "地址格式："
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label3 
      Caption         =   "● 基隆市、新竹市、嘉義市：格式為 XX 市 XX 區"
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   4380
   End
   Begin VB.Label Label2 
      Caption         =   "● 直轄市：臺北市 、新北市、桃園市、臺中市、臺南                         市、高雄市，格式為 XX 市 XX 區"
      Height          =   450
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   4380
   End
End
Attribute VB_Name = "frm100135"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/22 Form2.0已檢查 (無需修改的物件)
'Add by Amy 2015/07/24
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm100135 = Nothing
    'Add by Amy 2016/07/01 解按過臺灣地址格式(frm100135)後,無法開啟 代理人帳目的功能
    strFormName = ""
End Sub

