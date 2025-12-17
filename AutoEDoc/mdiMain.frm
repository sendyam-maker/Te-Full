VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H80000018&
   Caption         =   "電子公文"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Left            =   2835
      Top             =   1830
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '對齊表單下方
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   3870
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   1085
      ButtonWidth     =   609
      ButtonHeight    =   926
      Appearance      =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Activate()
   Static bolActivated As Boolean
   
   If bolActivated = False Then
      With frm010027
      .Show
      Me.Width = .Width + 150
      Me.Height = .Height + 150
      .WindowState = vbMaximized
      End With
      bolActivated = True
   End If
End Sub

Private Sub MDIForm_Load()
   strUserNum = "QPGMR"
   If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then
      If ConnectToServer_1 Then
         PUB_SetSystemVar
         PUB_SetUserData
         pub_OS = GetVersion32
         pub_HostName = PUB_ReadHostName
      Else
         End
      End If
   Else
      If PUB_Connect2DB() = False Then
         End
      End If
      PUB_SetUserData
      pub_OS = GetVersion32
      pub_HostName = PUB_ReadHostName
   End If
   mdiMain.Caption = mdiMain.Caption & " " & PUB_GetDbTerminal
   ToolHide
   
   Timer1.Interval = 1000
End Sub

Public Sub ToolHide()
   Toolbar1.Visible = False
   StatusBar1.Visible = False
End Sub

Private Sub Timer1_Timer()
   If Me.Visible = True Then
      MDIForm_Activate
      Timer1 = False
   End If
End Sub
