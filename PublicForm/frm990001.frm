VERSION 5.00
Begin VB.Form frm990001 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "關於TAIE程式 "
   ClientHeight    =   3015
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5895
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   136
   Icon            =   "frm990001.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2071.685
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   5521.665
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "確 定"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   4680
      TabIndex        =   0
      Top             =   2400
      Width           =   1020
   End
   Begin VB.Image imgKao 
      BorderStyle     =   1  '單線固定
      Height          =   945
      Index           =   0
      Left            =   360
      Top             =   120
      Width           =   1560
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "警告 : 本電腦程式受著作權法及國際公約保護。凡未經授權擅自複製或散佈本電腦程式的部份或全部，將遭受最嚴厲的民刑事處份。"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   828
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   4092
   End
   Begin VB.Label lblVersion 
      Caption         =   "Copyright (C)1999-2001 GATEWAY International Co., LTD All rights reserved."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   5148
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.4
      X2              =   5398.024
      Y1              =   1370.129
      Y2              =   1370.129
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "TAIE V1.0"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      BorderWidth     =   2
      Index           =   1
      X1              =   112.4
      X2              =   5398.024
      Y1              =   1353.638
      Y2              =   1353.638
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "傑威資訊股份有限公司協助製作"
      Height          =   300
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   3360
   End
End
Attribute VB_Name = "frm990001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/15 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit
Private Sub cmdOK_Click()
Unload Me
End Sub
Private Sub Form_Load()
MoveFormToCenter Me
End Sub
