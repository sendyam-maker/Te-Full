VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H80000018&
   Caption         =   "¤H¨Æ¨t²Î"
   ClientHeight    =   6360
   ClientLeft      =   2630
   ClientTop       =   3590
   ClientWidth     =   9750
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  '³Ì¤j¤Æ
   Begin VB.Timer tmrSalary 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1395
      Top             =   3600
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1440
      Top             =   1410
   End
   Begin VB.Timer tmrConnect 
      Left            =   1485
      Top             =   2010
   End
   Begin VB.Timer Timer2 
      Left            =   270
      Top             =   1950
   End
   Begin VB.Timer Timer1 
      Left            =   270
      Top             =   1470
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '¹ï»ôªí³æ¤U¤è
      Height          =   280
      Left            =   0
      TabIndex        =   1
      Top             =   6080
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   494
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   600
      _ExtentX        =   988
      _ExtentY        =   988
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      Height          =   520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   917
      ButtonWidth     =   406
      ButtonHeight    =   811
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   3000
      _ExtentX        =   494
      _ExtentY        =   494
      _Version        =   393216
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "¨t²Î"
      Index           =   0
      Begin VB.Menu mnu00 
         Caption         =   "¤Á´«³s½u"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnu00 
         Caption         =   "µ²§ô"
         Index           =   1
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "¸ê®Æ³B²z"
      Index           =   1
      Begin VB.Menu mnu1 
         Caption         =   "¤é±`§@·~"
         Index           =   1
         Begin VB.Menu mnu11 
            Caption         =   "­û¤u°ò¥»¸ê®Æ"
            Index           =   1
         End
         Begin VB.Menu mnu11 
            Caption         =   "¥X¯Ê¶Ô¸ê®Æ"
            Index           =   2
         End
         Begin VB.Menu mnu11 
            Caption         =   "¥[¯Z¸ê®Æ"
            Index           =   3
         End
         Begin VB.Menu mnu11 
            Caption         =   "¥X®t¸ê®Æ"
            Index           =   4
         End
         Begin VB.Menu mnu11 
            Caption         =   "½Ð°²¸ê®Æ"
            Index           =   5
         End
         Begin VB.Menu mnu11 
            Caption         =   "¥i¸É¥ð¸ê®Æ"
            Index           =   6
         End
         Begin VB.Menu mnu11 
            Caption         =   "±B³à³ß¼y¸ê®Æ"
            Index           =   7
         End
         Begin VB.Menu mnu11 
            Caption         =   "¤H¨Æ²§°Ê¸ê®Æ"
            Index           =   8
         End
         Begin VB.Menu mnu11 
            Caption         =   "¼úÃg¸ê®Æ"
            Index           =   9
         End
         Begin VB.Menu mnu11 
            Caption         =   "¥´¥d²§±`³B²z§@·~"
            Index           =   10
         End
         Begin VB.Menu mnu11 
            Caption         =   "°·ÀË³ø§i¸ê®Æ"
            Index           =   11
         End
         Begin VB.Menu mnu11 
            Caption         =   "®È¹C¸É§Uª÷¸ê®Æ"
            Index           =   12
         End
         Begin VB.Menu mnu11 
            Caption         =   "¤u§@©Ò¦b¦a¸ê®Æ"
            Index           =   13
         End
         Begin VB.Menu mnu11 
            Caption         =   "Excel¾ã§å¶×¤J¨ê¥d°O¿ý"
            Index           =   14
         End
      End
      Begin VB.Menu mnu1 
         Caption         =   "¦~«×§@·~"
         Index           =   2
         Begin VB.Menu mnu12 
            Caption         =   "ºÝ¤È¡B¤¤¬î¼úª÷ºûÅ@"
            Index           =   1
         End
         Begin VB.Menu mnu12 
            Caption         =   "·s¦~«×¯S§O°²ºûÅ@"
            Index           =   2
         End
         Begin VB.Menu mnu12 
            Caption         =   "®È¹C¸É§Uª÷ºûÅ@"
            Index           =   3
         End
         Begin VB.Menu mnu12 
            Caption         =   "§À¤úºN±m¡B¦~¸ê¡B¥þ¶Ô¼úª÷ºûÅ@"
            Index           =   4
         End
      End
      Begin VB.Menu mnu1 
         Caption         =   "ÀÉ®×ºûÅ@"
         Index           =   3
         Begin VB.Menu mnu13 
            Caption         =   "Â¾ºÙ¥N¸¹¸ê®Æ"
            Index           =   1
         End
         Begin VB.Menu mnu13 
            Caption         =   "Â¾¦ì¥N¸¹¸ê®Æ"
            Index           =   2
         End
         Begin VB.Menu mnu13 
            Caption         =   "¾Ç¾ú¥N¸¹¸ê®Æ"
            Index           =   3
         End
         Begin VB.Menu mnu13 
            Caption         =   "°²§O¥N¸¹¸ê®Æ"
            Index           =   4
         End
         Begin VB.Menu mnu13 
            Caption         =   "²§°Ê­ì¦]¥N¸¹¸ê®Æ"
            Index           =   5
         End
         Begin VB.Menu mnu13 
            Caption         =   "¥X¥Í¦a¥N¸¹¸ê®Æ"
            Index           =   6
         End
         Begin VB.Menu mnu13 
            Caption         =   "¼úÃg¥N¸¹¸ê®Æ"
            Index           =   7
         End
         Begin VB.Menu mnu13 
            Caption         =   "¤H¨ÆÂ¾¥N¤Î¼f®Ö¥DºÞ³]©w"
            Index           =   8
         End
         Begin VB.Menu mnu13 
            Caption         =   "Ã±®Ö¥DºÞ¯S®í¹ï¶HªºÃ±®ÖÂ¾¥N"
            Index           =   9
         End
         Begin VB.Menu mnu13 
            Caption         =   "·s«Ø«ü¯¾¾ã§å¶×¤J"
            Index           =   10
         End
         Begin VB.Menu mnu13 
            Caption         =   "­û¤u«ü¯¾¥d¤ù¸ê®Æ"
            Index           =   11
         End
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "¬d¸ß"
      Index           =   2
      Begin VB.Menu mnu2 
         Caption         =   "­û¤u©m¦W¬d¸ß­û¤u¸ê®Æ"
         Index           =   1
      End
      Begin VB.Menu mnu2 
         Caption         =   "µ{¦¡¤½§i¬d¸ß"
         Index           =   2
      End
      Begin VB.Menu mnu2 
         Caption         =   "»OÆW¦a§}¶l»¼°Ï¸¹¬d¸ß"
         Index           =   3
      End
      Begin VB.Menu mnu2 
         Caption         =   "¥[¯Z³æ²§±`¬d¸ß"
         Index           =   4
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "¦C¦L"
      Index           =   3
      Begin VB.Menu mnu3 
         Caption         =   "¤£©w´Á"
         Index           =   1
         Begin VB.Menu mnu31 
            Caption         =   "­û¤u¦W¥U"
            Index           =   1
         End
         Begin VB.Menu mnu31 
            Caption         =   "­Ó¤H¤H¨Æ¸ê®Æ©ú²Ó"
            Index           =   2
         End
         Begin VB.Menu mnu31 
            Caption         =   "¤÷¡B¥À¿Ë¸`¦W±ø"
            Index           =   3
         End
         Begin VB.Menu mnu31 
            Caption         =   "³Ò¡B°·¡B¹Î«O¦W³æ"
            Index           =   4
         End
         Begin VB.Menu mnu31 
            Caption         =   "¥X¯Ê¶Ô¬ö¿ý"
            Index           =   5
         End
         Begin VB.Menu mnu31 
            Caption         =   "­Ó¤H¥X®t¬ö¿ý"
            Index           =   6
         End
         Begin VB.Menu mnu31 
            Caption         =   "­Ó¤H¥[¯Z¬ö¿ý"
            Index           =   7
         End
         Begin VB.Menu mnu31 
            Caption         =   "­Ó¤H½Ð°²¬ö¿ý"
            Index           =   8
         End
         Begin VB.Menu mnu31 
            Caption         =   "­Ó¤H¥X¯Ê¶Ô©ú²Óªí"
            Index           =   9
         End
         Begin VB.Menu mnu31 
            Caption         =   "­Ó¤H¼úÃg¸ê®Æ©ú²Óªí"
            Index           =   10
         End
         Begin VB.Menu mnu31 
            Caption         =   "®Ê¤É¡B¯u°£"
            Index           =   11
         End
         Begin VB.Menu mnu31 
            Caption         =   "ºÝ¤È¡B¤¤¬î¼úª÷¦W³æ"
            Index           =   12
         End
         Begin VB.Menu mnu31 
            Caption         =   "¦UÃþ¥N¸¹¸ê®Æ"
            Index           =   14
         End
         Begin VB.Menu mnu31 
            Caption         =   "Â¾°È¥N²z¤H¸ê®Æªí"
            Index           =   15
         End
         Begin VB.Menu mnu31 
            Caption         =   "¨C¤é°²³æÃ±¦¬©ú²Óªí"
            Index           =   16
         End
         Begin VB.Menu mnu31 
            Caption         =   "¦U¦¡¤H¨Æ¸ê®Æªí"
            Index           =   17
         End
      End
      Begin VB.Menu mnu3 
         Caption         =   "¤ë³ø"
         Index           =   2
         Begin VB.Menu mnu32 
            Caption         =   "¥X¯Ê¶Ô²Î­p"
            Index           =   1
         End
         Begin VB.Menu mnu32 
            Caption         =   "¥X¯Ê¶Ô¥[¯Z¤ë²Î­p"
            Index           =   2
         End
         Begin VB.Menu mnu32 
            Caption         =   "³¡ªù¥[¯Z®É¼Æ²Î­p"
            Index           =   3
         End
         Begin VB.Menu mnu32 
            Caption         =   "¥´¥d©ú²Ó¸ê®Æ"
            Index           =   4
         End
      End
      Begin VB.Menu mnu3 
         Caption         =   "¦~³ø"
         Index           =   3
         Begin VB.Menu mnu33 
            Caption         =   "¥X¯Ê¶Ô¦~²Î­p¤Î¥þ¶Ô¦W³æ"
            Index           =   1
         End
         Begin VB.Menu mnu33 
            Caption         =   "¦~²×¦Ò¶Ô"
            Index           =   2
         End
         Begin VB.Menu mnu33 
            Caption         =   "¦~²×¦ÒÁZ¦Ò®Ö"
            Index           =   3
         End
         Begin VB.Menu mnu33 
            Caption         =   "¾ú¦~¦ÒÁZ"
            Index           =   4
         End
         Begin VB.Menu mnu33 
            Caption         =   "¦U¦¡¦W³æ"
            Index           =   5
         End
         Begin VB.Menu mnu33 
            Caption         =   "¯S§O°²¦W³æ"
            Index           =   6
         End
         Begin VB.Menu mnu33 
            Caption         =   "§Ñ°O¥´¥d¦¸¼Æ"
            Index           =   7
         End
         Begin VB.Menu mnu33 
            Caption         =   "À³Ãº°·ÀË³ø§i²M³æ"
            Index           =   8
         End
         Begin VB.Menu mnu33 
            Caption         =   "§À¤ú©â¼ú¤¤¼ú¦W³æ"
            Index           =   9
         End
         Begin VB.Menu mnu33 
            Caption         =   "§À¤ú©â¼ú¦W±ø"
            Index           =   10
         End
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "¤@¯ë§@·~"
      Index           =   4
      Begin VB.Menu mnu4 
         Caption         =   "¹w¬ù§@·~"
         Index           =   1
      End
      Begin VB.Menu mnu4 
         Caption         =   "¥X¯Ê¶Ô§@·~"
         Index           =   2
         Begin VB.Menu mnu42 
            Caption         =   "ªí³æ"
            Index           =   1
            Begin VB.Menu mnu421 
               Caption         =   "¥Ø«eªí³æ"
               Index           =   1
            End
            Begin VB.Menu mnu421 
               Caption         =   "Â¾¥N/Ã±®Ö¥DºÞ¥N¶ñªí³æ"
               Index           =   2
            End
            Begin VB.Menu mnu421 
               Caption         =   "¥´¥d²§±`­Ó¤H³B²z"
               Index           =   3
            End
            Begin VB.Menu mnu421 
               Caption         =   "¤U¯Z¹O30¤ÀÄÁ­ì¦]½T»{"
               Index           =   4
            End
         End
         Begin VB.Menu mnu42 
            Caption         =   "Ã±®Ö"
            Index           =   2
            Begin VB.Menu mnu422 
               Caption         =   "Ã±®Ö§@·~"
               Index           =   1
            End
            Begin VB.Menu mnu422 
               Caption         =   "¨C¤ë¥X¯Ê¶Ô²Î­p½T»{"
               Index           =   2
            End
            Begin VB.Menu mnu422 
               Caption         =   "­û¤u­Ó¤H¸ê®Æ©ú²Ó½T»{"
               Index           =   3
            End
            Begin VB.Menu mnu422 
               Caption         =   "Ã±®Ö¤H­û²§°Ê§@·~"
               Index           =   4
            End
            Begin VB.Menu mnu422 
               Caption         =   "¥´¥d²§±`¥DºÞ³B²z"
               Index           =   5
            End
         End
         Begin VB.Menu mnu42 
            Caption         =   "¬d¸ß"
            Index           =   3
            Begin VB.Menu mnu423 
               Caption         =   "ªñ¤é½Ð°²¤½§GÄæ/¤u§@©Ò¦b¦a"
               Index           =   1
            End
            Begin VB.Menu mnu423 
               Caption         =   "¥X¯Ê¶Ô¬d¸ß"
               Index           =   2
            End
            Begin VB.Menu mnu423 
               Caption         =   "Â¾¥N/Ã±®Ö¥DºÞÃöÁp¬d¸ß"
               Index           =   3
            End
            Begin VB.Menu mnu423 
               Caption         =   "¥´¥d¸ê®Æ¬d¸ß"
               Index           =   4
            End
         End
      End
      Begin VB.Menu mnu4 
         Caption         =   "Á~¸ê¬d¸ß¨t²Î"
         Index           =   3
         Begin VB.Menu mnu43 
            Caption         =   "­û¤uÁ~¸ê©ú²Ó"
            Index           =   1
         End
         Begin VB.Menu mnu43 
            Caption         =   "³Ò«O/°·«O/³Ò°hª÷©ú²Ó"
            Index           =   2
         End
         Begin VB.Menu mnu43 
            Caption         =   "¦~«×¦U¶µ©Ò±o©ú²Ó"
            Index           =   3
         End
         Begin VB.Menu mnu43 
            Caption         =   "¦~²×¼úª÷©ú²Ó"
            Index           =   4
         End
         Begin VB.Menu mnu43 
            Caption         =   "Á~¸ê¬d¸ß±K½X­×§ï"
            Index           =   5
         End
      End
      Begin VB.Menu mnu4 
         Caption         =   "¹Ï®Ñ­É¾\¸ê®Æ¬d¸ß"
         Index           =   4
      End
      Begin VB.Menu mnu4 
         Caption         =   "±Ð¨|°V½mµn¿ý§@·~"
         Index           =   5
      End
      Begin VB.Menu mnu4 
         Caption         =   "­û¤u¤u§@µû»ù¸ê®Æ"
         Index           =   6
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "»¡©ú"
      Index           =   15
      Begin VB.Menu mnu15 
         Caption         =   "»¡©ú¥DÃD"
         Index           =   1
      End
      Begin VB.Menu mnu15 
         Caption         =   "¯Á¤Þ"
         Index           =   2
      End
      Begin VB.Menu mnu15 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnu15 
         Caption         =   "Ãö©ó"
         Index           =   4
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "³]©w"
      Index           =   16
      Begin VB.Menu mnu16 
         Caption         =   "³øªí¯È±i®æ¦¡³]©w"
         Index           =   0
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "µøµ¡"
      Index           =   99
      Begin VB.Menu mnu99 
         Caption         =   "³Ìªñ¶}±Òµe­±"
         Index           =   0
      End
   End
   Begin VB.Menu mnuChUser 
      Caption         =   "§ó§ï¨Ï¥ÎªÌ"
   End
   Begin VB.Menu mnuDML 
      Caption         =   "¬dºûÅ@¬ö¿ý"
      Index           =   0
      Visible         =   0   'False
   End
   Begin VB.Menu mnuPop 
      Caption         =   "¤½¥Î¼u¸õ¿ï³æ"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuPopItem 
         Caption         =   "·s¼W"
         Index           =   0
      End
      Begin VB.Menu mnuPopItem 
         Caption         =   "­×§ï"
         Index           =   1
      End
      Begin VB.Menu mnuPopItem 
         Caption         =   "§R°£"
         Index           =   2
      End
      Begin VB.Menu mnuPopItem 
         Caption         =   "ÀËµø"
         Index           =   3
      End
   End
   Begin VB.Menu mnuPop2 
      Caption         =   "¤½¥Î¼u¸õ¿ï³æ2"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuPopItem2 
         Caption         =   "°Å¤U(&T)"
         Index           =   0
      End
      Begin VB.Menu mnuPopItem2 
         Caption         =   "½Æ»s(&C)"
         Index           =   1
      End
      Begin VB.Menu mnuPopItem2 
         Caption         =   "¶K¤W(&P)"
         Index           =   2
      End
      Begin VB.Menu mnuPopItem2 
         Caption         =   "§R°£(&D)"
         Index           =   3
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo By Sindy 2011/2/17 SQLDate¤wÀË¬d
'Memo By Sindy 2010/11/25 ­û¤u½s¸¹Äæ¤w­×§ï
'Memo By Sindy 2010/7/30 ¤é´ÁÄæ¤w­×§ï
Option Explicit

Dim WithEvents eventConn As ADODB.Connection
Attribute eventConn.VB_VarHelpID = -1
Public bolReOpen As Boolean
'intPCaseKind¤À®×¤§¨t²Î¤ÀÃþ¡AintPWhere 0°ê¤º  1°ê¥~CF  2°ê¥~FC
Public intPCaseKind As Integer, intPWhere As Integer
Public m_wasMaximized As Boolean 'Added by Morgan µe­±³Ì¤p¤Æ«á§PÂ_­ì¨Ó¬O§_¬°³Ì¤j¤Æ¥Î
Public m_ChkIsOpenFrm180203 As Boolean 'Add By Sindy 2013/7/8


'±±¨î³s½u¶¢¸m¶W¹L30¤ÀÄÁ¦Û°ÊÃö³¬µ{¦¡
Private Sub eventConn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
   tmrConnect.Tag = 0
End Sub

Private Sub SwitchMenu(Optional bolEnable As Boolean = True)
   Dim mnuTmp As Menu
   For Each mnuTmp In mnuTitle
      If mnuTmp.Index <> 0 Then mnuTmp.Enabled = bolEnable
   Next
   If bolEnable = False Then Toolbar1.Visible = False
End Sub

Private Sub CloseAllChild()
   Dim frmTemp As Form
   For Each frmTemp In Forms
      If frmTemp.Name <> "mdiMain" Then Unload frmTemp
   Next
End Sub

Private Sub ReConnect()

      Timer1.Enabled = True
      Timer1.Interval = 100
      tmrConnect.Tag = 0
   
End Sub

Private Sub MDIForm_Activate()
   'Modify By Sindy 2025/11/3 §ï¬°¦@¥Î¨ç¼Æ
   Call MDIFormStarProc
   
   'Add By Sindy 2021/2/18 ÀË¬d¯S§O°²¬O§_¤w§ó·s
   strSql = "select count(*) from YEARVACATION where yv01=" & Left(strSrvDate(1), 4) & " and yv11=0"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If RsTemp.Fields(0) > 0 Then
         MsgBox "¯S§O°²©|¦³ " & RsTemp.Fields(0) & " µ§¥¼§ó·s¡A¬O§_·s¦~«×¥¼¾ã§å§ó·s¡I" & vbCrLf & _
                "¡]½Ð¬d©ú­ì¦]¡^", vbExclamation
      End If
   End If
   '2021/2/18 END
End Sub

'Add By Sindy 2025/11/3
Public Sub SetTmpForm()
   Set Tmpfrm180201 = frm180201
   Set Tmpfrm180101 = frm180101
   Set Tmpfrm180203_1 = frm180203_1
   Set Tmpfrm160102 = frm160102
   Set Tmpfrm160018 = frm160018
   Set Tmpfrm010035_2 = frm010035_2
End Sub
'Add By Sindy 2011/10/7
Public Sub SysStartCallForm()
   '¦¹¨ç¼Æ¦b¦U¨t²Î¤@±Ò°Ê®É,¦]¥X¯Ê¶Ô«Ý¿ì´£¥Ü¯Ç¤J¤§¬G,¦@¥Î·|¨Ï¥Î¨ì,©Ò¥H¤£¥i§R°£
End Sub

Private Sub MDIForm_Resize()
   'Added by Morgan 2011/12/14 ¬ö¿ý¬O§_¬°³Ì¤j¤Æª¬ºA
   If Me.WindowState = 2 Then
      m_wasMaximized = True
   ElseIf Me.WindowState = 0 Then
      m_wasMaximized = False
   End If
End Sub

Private Sub mnu11_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 1   '­û¤u°ò¥»¸ê®Æ
        If CheckUse("frm160001", strExec) = False Then
            Exit Sub
        End If
        frm160001.Show
    Case 2   '¥X¯Ê¶Ô¸ê®Æ
        If CheckUse("frm160002", strExec) = False Then
            Exit Sub
        End If
        frm160002.Show
    Case 3   '¥[¯Z¸ê®Æ
        If CheckUse("frm160003", strExec) = False Then
            Exit Sub
        End If
        frm160003.Show
    Case 4   '¥X®t¸ê®Æ
        If CheckUse("frm160004", strExec) = False Then
            Exit Sub
        End If
        frm160004.Show
    Case 5   '½Ð°²¸ê®Æ
        If CheckUse("frm160005", strExec) = False Then
            Exit Sub
        End If
        frm160005.Show
    'Add By Sindy 2024/10/15
    Case 6   '¸É¥ð°²¸ê®Æ
        If CheckUse("frm160017", strExec) = False Then
            Exit Sub
        End If
        frm160017.Show
    Case 7   '±B³à³ß¼y¸ê®Æ
        If CheckUse("frm160006", strExec) = False Then
            Exit Sub
        End If
        frm160006.Show
    Case 8   '¤H¨Æ²§°Ê¸ê®Æ
        If CheckUse("frm160007", strExec) = False Then
            Exit Sub
        End If
        frm160007.Show
    Case 9   '¼úÃg¸ê®Æ
        If CheckUse("frm160008", strExec) = False Then
            Exit Sub
        End If
        frm160008.Show
    'Add By Sindy 2013/6/25
    Case 10   '¥´¥d²§±`³B²z§@·~
        If CheckUse("frm160012", strExec) = False Then
            Exit Sub
        End If
'        'Add By Sindy 2019/2/21
'        If PUB_GetLock(UCase("frm160012"), "", "¥´¥d²§±`³B²z§@·~") = False Then
'            Exit Sub
'        End If
'        '2019/2/21 END
        frm160012.Show
    'Add By Sindy 2015/8/11
    Case 11  '°·ÀË³ø§i¸ê®Æ
        If CheckUse("frm160019", strExec) = False Then
            Exit Sub
        End If
        frm160019.Show
    'Add By Sindy 2019/7/24
    Case 12  '­û¤u®È¹C¸É§Uª÷¸ê®Æ
        If CheckUse("frm160020", strExec) = False Then
            Exit Sub
        End If
        frm160020.Show
    'Add By Sindy 2020/4/11
    Case 13  '¤u§@©Ò¦b¦a¸ê®Æ
        If CheckUse("frm160022", strExec) = False Then
            Exit Sub
        End If
        frm160022.Show
    'Add By Sindy 2021/6/4
    Case 14  'Excel¾ã§å¶×¤J¨ê¥d°O¿ý
        If CheckUse("frm160023", strExec) = False Then
            Exit Sub
        End If
        frm160023.Show
    End Select
End Sub

Private Sub mnu12_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 1   'ºÝ¤È¡B¤¤¬î¼úª÷ºûÅ@
        If CheckUse("frm160009", strExec) = False Then
            Exit Sub
        End If
        frm160009.Show
    Case 2   '·s¦~«×¯S§O°²ºûÅ@
        If CheckUse("frm160010", strExec) = False Then
            Exit Sub
        End If
        frm160010.Show
    'Add By Sindy 2019/7/25
    Case 3   '®È¹C¸É§Uª÷ºûÅ@
        If CheckUse("frm160021", strExec) = False Then
            Exit Sub
        End If
        frm160021.Show
    Case 4   '§À¤úºN±m¡B¦~¸ê¡B¥þ¶Ô¼úª÷ºûÅ@
        If CheckUse("frm170032", strExec) = False Then
            Exit Sub
        End If
        frm170032.Show
    Case Else
    End Select
End Sub

Private Sub mnu13_Click(Index As Integer)
ProSysState = ""
    ToolHide
    Select Case Index
    Case 1   'Â¾ºÙ¥N¸¹¸ê®Æ
        If CheckUse("frm160011A", strExec) = False Then
            Exit Sub
        End If
        ProSysState = "A"
        frm160011.Show
    Case 2   'Â¾¦ì¥N¸¹¸ê®Æ
        If CheckUse("frm160011B", strExec) = False Then
            Exit Sub
        End If
        ProSysState = "B"
        frm160011.Show
    Case 3   '¾Ç“ð¥N¸¹¸ê®Æ
        If CheckUse("frm160011C", strExec) = False Then
            Exit Sub
        End If
        ProSysState = "C"
        frm160011.Show
    Case 4   '°²§O¥N¸¹¸ê®Æ
        If CheckUse("frm160011D", strExec) = False Then
            Exit Sub
        End If
        ProSysState = "D"
        frm160011.Show
    Case 5    '²§°Ê­ì¦]¥N¸¹¸ê®Æ
        If CheckUse("frm160011E", strExec) = False Then
            Exit Sub
        End If
        ProSysState = "E"
        frm160011.Show
    Case 6    '¥X¥Í¦a¥N¸¹¸ê®Æ
        If CheckUse("frm160011F", strExec) = False Then
            Exit Sub
        End If
        ProSysState = "F"
        frm160011.Show
    Case 7    '¼úÃg¥N¸¹¸ê®Æ
        If CheckUse("frm160011H", strExec) = False Then
            Exit Sub
        End If
        ProSysState = "H"
        frm160011.Show
    Case 8    '¤H¨ÆÂ¾¥N¤Î¼f®Ö¥DºÞ³]©w
        If CheckUse("frm180401", strExec) = False Then
            Exit Sub
        End If
        frm180401.Show
    Case 9   '¯S®í¨­¥÷Â¾°È¥N²z¤H
        If CheckUse("frm180402", strExec) = False Then
            Exit Sub
        End If
        frm180402.Show
    'Added by Morgan 2013/7/15
    Case 10 '·s«Ø«ü¯¾¾ã§å¶×¤J
        If CheckUse("frm160013", strExec) = False Then
            Exit Sub
        End If
        frm160013.Show
    'Added by Morgan 2013/7/17
    Case 11 '­û¤u«ü¯¾¥d¤ù¸ê®Æ
        If CheckUse("frm160014", strExec) = False Then
            Exit Sub
        End If
        frm160014.Show
    Case Else
    End Select
End Sub

Private Sub mnu16_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 0   '³øªí¯È±i®æ¦¡³]©w
         frm880013.Show vbModal
   End Select
End Sub

Private Sub mnu2_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1   '¥H­û¤u©m¦W¬d¸ß­û¤u¸ê®Æ    '¼È¥Î
         'Modify by Amy 2014/04/30 Mark CheckUse
         'If CheckUse("frm100121_1", strExec) = False Then
         '   Exit Sub
         'End If
         frm100121_1.Show
      Case 2 'Add By Amy 2013/05/08 µ{¦¡¤½§i¬d¸ß
         frm100131.Show
      Case 3 'Add By Sindy 2015/3/20 »OÆW¦a§}¶l»¼°Ï¸¹¬d¸ß
         frm100134.Show
      'Add By Sindy 2013/8/7
      Case 4 '¥[¯Z³æ²§±`¬d¸ß
         If CheckUse("frm160501", strExec) = False Then
            Exit Sub
         End If
         frm160501.Show
   End Select
End Sub


Private Sub mnu31_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 1   '­û¤u¦W¥U¦C¦L
        If CheckUse("frm160101", strExec) = False Then
            Exit Sub
        End If
        frm160101.Show
    Case 2   '­Ó¤H¤H¨Æ¸ê®Æ©ú²Ó¦C¦L
        If CheckUse("frm160102", strExec) = False Then
            Exit Sub
        End If
        frm160102.Show
    Case 3   '¤÷¡B¥À¿Ë¸`¦W±ø¦C¦L
        If CheckUse("frm160103", strExec) = False Then
            Exit Sub
        End If
        frm160103.Show
    Case 4   '³Ò¡B°·¡B¹Î«O¶O¦W³æ¦C¦L
        If CheckUse("frm160104", strExec) = False Then
            Exit Sub
        End If
        frm160104.Show
    Case 5   '¥X¯Ê¶Ô¬ö¿ý¦C¦L
        If CheckUse("frm160105", strExec) = False Then
            Exit Sub
        End If
        frm160105.Show
    Case 6   '­Ó¤H¥X®t¬ö¿ý¦C¦L
        If CheckUse("frm160106", strExec) = False Then
            Exit Sub
        End If
        frm160106.Show
    Case 7   '­Ó¤H¥[¯Z¬ö¿ý¦C¦L
        If CheckUse("frm160107", strExec) = False Then
            Exit Sub
        End If
        frm160107.Show
    Case 8   '­Ó¤H½Ð°²¬ö¿ý¦C¦L
        If CheckUse("frm160108", strExec) = False Then
            Exit Sub
        End If
        frm160108.Show
    Case 9   '­Ó¤H¥X¯Ê¶Ô©ú²Óªí
        If CheckUse("frm160113", strExec) = False Then
            Exit Sub
        End If
        frm160113.Show
    '--2014/9/19 Add By Lydia ­Ó¤H¼úÃg¸ê®Æ©ú²Óªí
    Case 10   '­Ó¤H¼úÃg¸ê®Æ©ú²Óªí
        If CheckUse("frm160115", strExec) = False Then
            Exit Sub
        End If
        frm160115.Show
    Case 11   '®Ê¤É¡B¯u°£¦C¦L
        If CheckUse("frm160109", strExec) = False Then
            Exit Sub
        End If
        frm160109.Show
    Case 12   'ºÝ¤È¡B¤¤¬î¼úª÷¦W³æ
        If CheckUse("frm160110", strExec) = False Then
            Exit Sub
        End If
        frm160110.Show
'    Case 13   '­û¤u¤ë¥d¦W±ø
'        If CheckUse("frm160111", strExec) = False Then
'            Exit Sub
'        End If
'        frm160111.Show
    Case 14   '¦UÃþ¥N¸¹¸ê®Æ
        If CheckUse("frm160112", strExec) = False Then
            Exit Sub
        End If
        frm160112.Show
    Case 15   'Â¾°È¥N²z¤H¸ê®Æªí
        If CheckUse("frm180501", strExec) = False Then
            Exit Sub
        End If
        frm180501.Show
    Case 16   '¨C¤é°²³æÃ±¦¬©ú²Óªí
        If CheckUse("frm180502", strExec) = False Then
            Exit Sub
        End If
        frm180502.Show
    'Add By Sindy 2012/2/10
    Case 17   '¦U¦¡¤H¨Æ¸ê®Æªí
        If CheckUse("frm160114", strExec) = False Then
            Exit Sub
        End If
        frm160114.Show
    Case Else
    End Select
End Sub

Private Sub mnu32_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 1   '¥X¯Ê¶Ô²Î­p
        If CheckUse("frm160201", strExec) = False Then
            Exit Sub
        End If
        frm160201.Show
    Case 2   '¥X¯Ê¶Ô¥[¯Z¤ë²Î­p
        If CheckUse("frm160202", strExec) = False Then
            Exit Sub
        End If
        frm160202.Show
    Case 3   '³¡ªù¥[¯Z®É¼Æ²Î­p
        If CheckUse("frm160203", strExec) = False Then
            Exit Sub
        End If
        frm160203.Show
    'Add by Sindy 2013/8/5
    Case 4   '¥´¥d©ú²Ó¸ê®Æ
        If CheckUse("frm160204", strExec) = False Then
            Exit Sub
        End If
        frm160204.Show
    Case Else
    End Select
End Sub

Private Sub mnu33_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 1   '¥X¯Ê¶Ô¦~²Î­p¤Î¥þ¶Ô¦W³æ
        If CheckUse("frm160301", strExec) = False Then
            Exit Sub
        End If
        frm160301.Show
    Case 2   '¦~²×¦Ò¶Ô
        If CheckUse("frm160302", strExec) = False Then
            Exit Sub
        End If
        frm160302.Show
    Case 3   '¦~²×¦ÒÁZ¦Ò®Ö
        If CheckUse("frm160303", strExec) = False Then
            Exit Sub
        End If
        frm160303.Show
    Case 4   '¾ú¦~¦ÒÁZ
        If CheckUse("frm160304", strExec) = False Then
            Exit Sub
        End If
        frm160304.Show
    Case 5   '¦U¦¡¦W³æ
        If CheckUse("frm160305", strExec) = False Then
            Exit Sub
        End If
        frm160305.Show
    Case 6   '¯S§O°²¦W³æ
        If CheckUse("frm160306", strExec) = False Then
            Exit Sub
        End If
        frm160306.Show
    'Add By Sindy 2010/1/15
    Case 7   '§Ñ°O¥´¥d¦¸¼Æ
        If CheckUse("frm160307", strExec) = False Then
            Exit Sub
        End If
        frm160307.Show
    'Add By Sindy 2015/8/12
    Case 8   'À³Ãº°·ÀË³ø§i²M³æ
        If CheckUse("frm160308", strExec) = False Then
            Exit Sub
        End If
        frm160308.Show
        
    'Added by Morgan 2023/11/20
    Case 9 '§À¤ú©â¼ú¤¤¼ú¦W³æ
        If CheckUse("frm160309", strExec) = False Then
            Exit Sub
        End If
        frm160309.Show
    'Added by Morgan 2023/11/22
    Case 10 '§À¤ú©â¼ú¦W±ø
        If CheckUse("frm160310", strExec) = False Then
            Exit Sub
        End If
        frm160310.Show
    Case Else
    End Select
End Sub

Private Sub mnu4_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1 '¹w¬ù§@·~
         frm140112.Show
      Case 4 '¹Ï®Ñ­É¾\¸ê®Æ¬d¸ß Add by Amy 2017/02/3
         frm010035.Show
        If GetLoanRecordApply = True Then
            frm010035.bolLoanRecordApply = True
            Call frm010035.cmdLoanRecord_Click
         End If
      Case 5 'Add by Amy 2020/11/02 ±Ð¨|°V½mµn¤J§@·~
         frm140113.Show
      'Add By Sindy 2023/10/16
      Case 6 '­û¤u¤u§@µû»ù¸ê®Æ
         Call PUB_OpenFrm160016(frm160016)
   End Select
End Sub

Private Sub mnu421_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1   '¥Ø«eªí³æ
         frm180101.Show
      Case 2   'Â¾¥N/Ã±®Ö¥DºÞ¥N¶ñªí³æ
         frm180103.Show
      'Add By Sindy 2013/6/25
      Case 3   '¥´¥d²§±`­Ó¤H³B²z
         frm180105.Show
      'Add By Sindy 2025/10/7
      Case 4   '¤U¯Z¹O30¤ÀÄÁ­ì¦]½T»{
         frm160018.Show
   End Select
End Sub

Private Sub mnu422_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1   'Ã±®Ö§@·~
         frm180201.Show
      Case 2   '¨C¤ë¥X¯Ê¶Ô²Î­p½T»{
'         frm160201.intChoose = 1
'         frm160201.Hide
'         Call frm160201.cmdOK_Click(0)
         frm180203_1.Show
      Case 3   '­û¤u­Ó¤H¸ê®Æ©ú²Ó½T»{
         frm160102.intChoose = 1
         frm160102.Hide
         Call frm160102.cmdok_Click(0)
      Case 4 'Ã±®Ö¤H­û²§°Ê§@·~
         frm180104.Show
      'Add By Sindy 2013/6/25
      Case 5 '¥´¥d²§±`¥DºÞ³B²z
         frm180204.Show
   End Select
End Sub

Private Sub mnu423_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1 'ªñ¤é½Ð°²¤½§GÄæ
         frm180302.Show
      Case 2 '¥X¯Ê¶Ô¬d¸ß
         frm180301.Show
      Case 3 'Â¾¥N/¼f®Ö¥DºÞÃöÁp¬d¸ß
         frm180403.Show
      'Add By Sindy 2013/6/25
      Case 4 '¥´¥d¸ê®Æ¬d¸ß
         frm180303.Show
   End Select
End Sub

Private Sub mnuChUser_Click()
   frmChgUser.Show
End Sub

Private Sub mnuDML_Click(Index As Integer)
    frmDML.Show   '°ò¥»¸ê®ÆºûÅ@¬ö¿ý
End Sub

'±±¨î¤£¥i«þ¨©µe­±
Private Sub Timer3_Timer()
   Static dtNow As Date 'Added by Morgan 2024/8/7
      
On Error Resume Next 'Added by Morgan 2017/8/29 ­Y¦³¨ä¥L³nÅé¤]¦b¨Ï¥Î°Å¶KÃ¯®É·|µo¥Í521(µLªk¶}±Ò°Å¶KÃ¯)ªº¿ù»~(Ex.Word¶}±Ò°Å¶KÃ¯¨ÃÂ^¨úµe­±)

   'Added by Morgan 2024/8/7 ©w®É°õ¦æ¤@¦¸»yªk¥H½T«O¸óºô¬q³s½u®Éºô¸ô¤£·|³Q¤ÁÂ_
   If tmrConnect.Interval = 0 Then
      If Now > dtNow Then
         dtNow = DateAdd("n", cntAutoQueryInterval, Now)
         ClsLawReadRstMsg 1, "select * from dual"
      End If
   End If
   'end 2024/8/7
   
   '¹q¸£¤¤¤ßªº¤£ºÞ
   If Pub_StrUserSt03 = "M51" Or Pub_Can_Copy_Pic = True Then Exit Sub
   '¹ÏÀÉ¤~²M
   If Clipboard.GetFormat(1) = False And Clipboard.GetFormat(2) = True And Clipboard.GetFormat(3) = False Then
       Clipboard.Clear
   End If
End Sub

'±±¨î³s½u¶¢¸m¶W¹L10¤ÀÄÁ¦Û°ÊÂ÷½u
Private Sub tmrConnect_Timer()
   tmrConnect.Tag = tmrConnect.Tag + 1
   If tmrConnect.Tag = 10 Then
      Timer1.Enabled = False
      bolReOpen = False
      frmReopen.Show vbModal, Me
      If bolReOpen = True Then
         Call ReConnect
      Else
         Call mnu00_Click(1)
      End If
   End If
End Sub

Private Sub MDIForm_Load()
'±±¨î³s½u¶¢¸m¶W¹L30¤ÀÄÁ¦Û°ÊÃö³¬µ{¦¡
If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then
   Set eventConn = cnnConnection
   tmrConnect.Interval = 60000
End If
       DisableControl Me
Dim strSysKind As String
Dim lngValue, lngBufferSize As Long, intCounter As Integer
Dim strUserId As String * 10, strLocalId As String

    '­Yµn¤J¦¨¥\
    If pub_str_LoginSucceeded = "1" Then
        Me.Timer1.Interval = 100
       strSysKind = GetSystemKindByNick
       '¥i¥H¬d¸ßºûÅ@¬ö¿ý
       If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or Pub_StrUserSt03 = "M51" Then
         mnuDML(0).Visible = True
         mnuChUser.Visible = True 'Added by Sindy 2013/7/23
       Else
         mnuDML(0).Visible = False
         mnuChUser.Visible = False 'Added by Sindy 2013/7/23
       End If

'       'Added by Morgan 2016/1/22 Á~¸ê¬d¸ß´ú¸Õ
'       If Pub_StrUserSt03 = "M51" Then
'         mnu4(3).Visible = True
'       Else
'         mnu4(3).Visible = False
'       End If
'       'end
'
        If Me.ImageList1.ListImages.Count <= 0 Then
           ImageList1.ListImages.add , "graphic1", LoadPicture(strPicPath & "misc41.ico")
           ImageList1.ListImages.add , "graphic2", LoadPicture(strPicPath & "note16.ico")
           ImageList1.ListImages.add , "graphic3", LoadPicture(strPicPath & "erase02.ico")
           ImageList1.ListImages.add , "graphic4", LoadPicture(strPicPath & "drive03.ico")
           ImageList1.ListImages.add , "graphic5", LoadPicture(strPicPath & "trash02.ico")
           ImageList1.ListImages.add , "graphic6", LoadPicture(strPicPath & "explorer.ico")
           ImageList1.ListImages.add , "graphic7", LoadPicture(strPicPath & "printfld.ico")
           ImageList1.ListImages.add , "graphic8", LoadPicture(strPicPath & "first.ico")
           ImageList1.ListImages.add , "graphic9", LoadPicture(strPicPath & "prior.ico")
           ImageList1.ListImages.add , "graphic10", LoadPicture(strPicPath & "next.ico")
           ImageList1.ListImages.add , "graphic11", LoadPicture(strPicPath & "last.ico")
           ImageList1.ListImages.add , "graphic12", LoadPicture(strPicPath & "net14.ico")
           ImageList1.ListImages.add , "graphic13", LoadPicture(strPicPath & "w95mbx01.ico")
           Toolbar1.ImageList = ImageList1
           Toolbar1.Buttons.add , "function1", , tbrDefault, "graphic1"
           Toolbar1.Buttons.add , "none1", , tbrSeparator
           Toolbar1.Buttons.add , "none2", , tbrSeparator
           Toolbar1.Buttons.add , "function2", , tbrDefault, "graphic2"
           Toolbar1.Buttons.add , "function3", , tbrDefault, "graphic3"
           Toolbar1.Buttons.add , "function4", , tbrDefault, "graphic4"
           Toolbar1.Buttons.add , "function12", , tbrDefault, "graphic13"
           Toolbar1.Buttons.add , "function5", , tbrDefault, "graphic5"
           Toolbar1.Buttons.add , "function6", , tbrDefault, "graphic6"
        '   Toolbar1.Buttons.Add , "function7", , tbrDefault, "graphic7"
           Toolbar1.Buttons.add , "none5", , tbrSeparator
           Toolbar1.Buttons.add , "none3", , tbrSeparator
           Toolbar1.Buttons.add , "none4", , tbrSeparator
           Toolbar1.Buttons.add , "function8", , tbrDefault, "graphic8"
           Toolbar1.Buttons.add , "function9", , tbrDefault, "graphic9"
           Toolbar1.Buttons.add , "function10", , tbrDefault, "graphic10"
           Toolbar1.Buttons.add , "function11", , tbrDefault, "graphic11"
           Toolbar1.Buttons.Item(1).ToolTipText = "Ãö³¬(Esc)"
           Toolbar1.Buttons.Item(4).ToolTipText = "·s¼W(F2)"
           Toolbar1.Buttons.Item(5).ToolTipText = "­×§ï(F3)"
           Toolbar1.Buttons.Item(6).ToolTipText = "¦sÀÉ(F9)"
           Toolbar1.Buttons.Item(7).ToolTipText = "©ñ±ó(F10)"
           Toolbar1.Buttons.Item(8).ToolTipText = "§R°£(F5)"
           Toolbar1.Buttons.Item(9).ToolTipText = "¬d¸ß(F4)"
        '   Toolbar1.Buttons.Item(10).ToolTipText = "¦C¦L(F7)"
           Toolbar1.Buttons.Item(13).ToolTipText = "²Ä¤@µ§(Home)"
           Toolbar1.Buttons.Item(14).ToolTipText = "¤W¤@µ§(PageUp)"
           Toolbar1.Buttons.Item(15).ToolTipText = "¤U¤@µ§(PageDown)"
           Toolbar1.Buttons.Item(16).ToolTipText = "³Ì«á(End)"
        End If
       tool4_enabled
       strFormName = MsgText(601)
       strExitControl = MsgText(602)
        If Me.StatusBar1.Panels.Count < 5 Then
            For intCounter = 1 To 4
               StatusBar1.Panels.add
            Next intCounter
        End If
       StatusBar1.Height = 300
       StatusBar1.Panels.Item(1).Width = 5500
       StatusBar1.Panels.Item(2).Width = 1000
       StatusBar1.Panels.Item(3).Text = CFDate(ACDate(ServerDate))
       StatusBar1.Panels.Item(4).Text = time
'       Me.Icon = LoadPicture(strIcoPath)
       ToolHide
       Systemkind_g = GetSystemKindByNick
       Systemkind_g_P = GetSystemKindByNickP
       Systemkind_g_T = GetSystemKindByNickT
       Systemkind_g_TnoS = GetSystemKindByNickTnoS
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim frm As Form
'Ãö³¬©|¥¼Ãö³¬ªº¤lµøµ¡
For Each frm In Forms
    If frm.Name <> mdiMain.Name Then
        Unload frm
    End If
Next
PUB_AddAuditLog AL_µn¥X 'Added by Morgan 2025/7/31

Set mdiMain = Nothing
End Sub
'¥[¤Á´«³s½u¿ï¾Ü
Private Sub mnu00_Click(Index As Integer)
   Select Case Index
      Case 0
         If PUB_Connect2DB(True) = False Then
            Unload Me
         End If
      Case 1
         Unload Me
   End Select
End Sub

Public Sub ToolShow()
   Toolbar1.Visible = True
   StatusBar1.Visible = True
End Sub

'*************************************************
'  ¤u¨ã¦C«ö¶s¥¢®Ä³]©w1
'
'*************************************************
Public Sub tool1_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = True
   Toolbar1.Buttons.Item(5).Enabled = True
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = True
   Toolbar1.Buttons.Item(9).Enabled = True
   Toolbar1.Buttons.Item(10).Enabled = True
   Toolbar1.Buttons.Item(13).Enabled = True
   Toolbar1.Buttons.Item(14).Enabled = True
   Toolbar1.Buttons.Item(15).Enabled = True
   Toolbar1.Buttons.Item(16).Enabled = True
End Sub


'*************************************************
'  ¤u¨ã¦C«ö¶s¥¢®Ä³]©w2
'
'*************************************************
Public Sub tool2_enabled()
   Toolbar1.Buttons.Item(1).Enabled = False
   Toolbar1.Buttons.Item(4).Enabled = False
   Toolbar1.Buttons.Item(5).Enabled = False
   Toolbar1.Buttons.Item(6).Enabled = True
   Toolbar1.Buttons.Item(7).Enabled = True
   Toolbar1.Buttons.Item(8).Enabled = False
   Toolbar1.Buttons.Item(9).Enabled = False
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = False
   Toolbar1.Buttons.Item(14).Enabled = False
   Toolbar1.Buttons.Item(15).Enabled = False
   Toolbar1.Buttons.Item(16).Enabled = False
End Sub


'*************************************************
'  ¤u¨ã¦C«ö¶s¥¢®Ä³]©w3
'
'*************************************************
Public Sub tool3_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = False
   Toolbar1.Buttons.Item(5).Enabled = False
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = False
   Toolbar1.Buttons.Item(9).Enabled = False
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = False
   Toolbar1.Buttons.Item(14).Enabled = False
   Toolbar1.Buttons.Item(15).Enabled = False
   Toolbar1.Buttons.Item(16).Enabled = False
End Sub


'*************************************************
'  ¤u¨ã¦C«ö¶s¥¢®Ä³]©w4
'
'*************************************************
Public Sub tool4_enabled()
   Toolbar1.Buttons.Item(1).Enabled = False
   Toolbar1.Buttons.Item(4).Enabled = False
   Toolbar1.Buttons.Item(5).Enabled = False
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = False
   Toolbar1.Buttons.Item(9).Enabled = False
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = False
   Toolbar1.Buttons.Item(14).Enabled = False
   Toolbar1.Buttons.Item(15).Enabled = False
   Toolbar1.Buttons.Item(16).Enabled = False
End Sub

'*************************************************
'  ¤u¨ã¦C«ö¶s¥¢®Ä³]©w5
'
'*************************************************
Public Sub tool5_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = False
   Toolbar1.Buttons.Item(5).Enabled = False
   Toolbar1.Buttons.Item(6).Enabled = True
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = True
   Toolbar1.Buttons.Item(9).Enabled = True
   Toolbar1.Buttons.Item(10).Enabled = True
   Toolbar1.Buttons.Item(13).Enabled = True
   Toolbar1.Buttons.Item(14).Enabled = True
   Toolbar1.Buttons.Item(15).Enabled = True
   Toolbar1.Buttons.Item(16).Enabled = True
End Sub

'*************************************************
'  ¤u¨ã¦C«ö¶s¥¢®Ä³]©w6
'
'*************************************************
Public Sub tool6_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = False
   Toolbar1.Buttons.Item(5).Enabled = False
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = False
   Toolbar1.Buttons.Item(9).Enabled = False
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = True
   Toolbar1.Buttons.Item(14).Enabled = True
   Toolbar1.Buttons.Item(15).Enabled = True
   Toolbar1.Buttons.Item(16).Enabled = True
End Sub

'*************************************************
'  ¤u¨ã¦C«ö¶s¥¢®Ä³]©w7
'
'*************************************************
Public Sub tool7_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = False
   Toolbar1.Buttons.Item(5).Enabled = True
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = False
   Toolbar1.Buttons.Item(9).Enabled = False
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = False
   Toolbar1.Buttons.Item(14).Enabled = False
   Toolbar1.Buttons.Item(15).Enabled = False
   Toolbar1.Buttons.Item(16).Enabled = False
End Sub

'*************************************************
'  ¤u¨ã¦C«ö¶s¥¢®Ä³]©w8
'
'*************************************************
Public Sub tool8_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = False
   Toolbar1.Buttons.Item(5).Enabled = True
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = False
   Toolbar1.Buttons.Item(9).Enabled = True
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = True
   Toolbar1.Buttons.Item(14).Enabled = True
   Toolbar1.Buttons.Item(15).Enabled = True
   Toolbar1.Buttons.Item(16).Enabled = True
End Sub

'*************************************************
'  ¤u¨ã¦C«ö¶s¥¢®Ä³]©w9
'
'*************************************************
Public Sub tool9_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = True
   Toolbar1.Buttons.Item(5).Enabled = False
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = False
   Toolbar1.Buttons.Item(9).Enabled = False
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = False
   Toolbar1.Buttons.Item(14).Enabled = False
   Toolbar1.Buttons.Item(15).Enabled = False
   Toolbar1.Buttons.Item(16).Enabled = False
End Sub

'*************************************************
'  ¤u¨ã¦C«ö¶s¥¢®Ä³]©w10
'
'*************************************************
Public Sub tool10_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = True
   Toolbar1.Buttons.Item(5).Enabled = True
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = True
   Toolbar1.Buttons.Item(9).Enabled = False
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = False
   Toolbar1.Buttons.Item(14).Enabled = False
   Toolbar1.Buttons.Item(15).Enabled = False
   Toolbar1.Buttons.Item(16).Enabled = False
End Sub

'*************************************************
'  ¤u¨ã¦C«ö¶s¥¢®Ä³]©w11
'
'*************************************************
Public Sub tool11_enabled()
   Toolbar1.Buttons.Item(1).Enabled = False
   Toolbar1.Buttons.Item(4).Enabled = False
   Toolbar1.Buttons.Item(5).Enabled = False
   Toolbar1.Buttons.Item(6).Enabled = True
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = False
   Toolbar1.Buttons.Item(9).Enabled = False
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = False
   Toolbar1.Buttons.Item(14).Enabled = False
   Toolbar1.Buttons.Item(15).Enabled = False
   Toolbar1.Buttons.Item(16).Enabled = False
End Sub

'*************************************************
'  ¤u¨ã¦C«ö¶s¥¢®Ä³]©w12
'
'*************************************************
Public Sub tool12_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = True
   Toolbar1.Buttons.Item(5).Enabled = True
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = True
   Toolbar1.Buttons.Item(9).Enabled = False
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = False
   Toolbar1.Buttons.Item(14).Enabled = False
   Toolbar1.Buttons.Item(15).Enabled = False
   Toolbar1.Buttons.Item(16).Enabled = False
End Sub

'*************************************************
'  ¤u¨ã¦C«ö¶s¥¢®Ä³]©w13
'
'*************************************************
Public Sub tool13_enabled()
   Toolbar1.Buttons.Item(1).Enabled = True
   Toolbar1.Buttons.Item(4).Enabled = False
   Toolbar1.Buttons.Item(5).Enabled = False
   Toolbar1.Buttons.Item(6).Enabled = False
   Toolbar1.Buttons.Item(7).Enabled = False
   Toolbar1.Buttons.Item(8).Enabled = False
   Toolbar1.Buttons.Item(9).Enabled = True
   Toolbar1.Buttons.Item(10).Enabled = False
   Toolbar1.Buttons.Item(13).Enabled = False
   Toolbar1.Buttons.Item(14).Enabled = False
   Toolbar1.Buttons.Item(15).Enabled = False
   Toolbar1.Buttons.Item(16).Enabled = False
End Sub

Private Sub mnu15_Click(Index As Integer)
ToolHide

End Sub

Public Sub ToolHide()
   Toolbar1.Visible = False
   StatusBar1.Visible = False
End Sub

Private Sub mnu99_Click(Index As Integer)
Dim frm As Form
    
    For Each frm In Forms
        If frm.Name <> "mdiMain" Then
            If frm.Name = mnu99(Index).Tag Then
                '±N¤lµøµ¡±Æ¦b³»¼h
                frm.ZOrder (0)
                Exit For
            End If
        End If
    Next
End Sub

Private Sub Timer1_Timer()
Dim frm As Form

   '±±¨î"µøµ¡"Menu
   MenuForFormControl
   StatusBar1.Panels.Item(4).Text = time
   
   mnuTitle(99).Visible = mnuTitle(99).Enabled
End Sub

Private Sub Timer2_Timer()
    '­Yµn¤J¥¢±Ñ
    If pub_str_LoginSucceeded <> "1" Then
        Me.Timer1.Interval = 0
    '­Yµn¤J¦¨¥\
    Else
        Me.Timer1.Interval = 100
        Me.Timer2.Interval = 0
        'MDIForm_Load
        Me.Show
    End If
End Sub

Private Sub MenuForFormControl()
Dim frm As Form
Dim ii As Integer
Dim objMnu99 As Menu
Dim intMaxIndex As Integer
Dim blnFormNameMatch As Boolean

On Error Resume Next
   '­YµL¥ô¦ó¤lµøµ¡
   If Forms.Count <= 1 Then
       If Me.mnuTitle(99).Enabled = True Then
           Me.mnuTitle(99).Enabled = False
           For Each objMnu99 In Me.mnu99
               If objMnu99.Index = 0 Then
                   Me.mnu99(objMnu99.Index).Tag = ""
               Else
                   Unload Me.mnu99(objMnu99.Index)
               End If
           Next
       End If
   '­Y¦³¤lµøµ¡
   Else
       If Me.mnuTitle(99).Enabled = False Then Me.mnuTitle(99).Enabled = True
       '­Y¤lµøµ¡¼Æ»Pµøµ¡menu¼Æ¤£¦P
       If Forms.Count - 1 <> Me.mnu99.Count Then
           For Each frm In Forms
               If frm.Name <> "mdiMain" And frm.Caption <> "" Then
                   '­Y¤lµøµ¡³QÁôÂÃ
                   If frm.Visible = False Then
                       For Each objMnu99 In Me.mnu99
                           If frm.Name = Me.mnu99(objMnu99.Index).Tag Then
                               If Me.mnu99(objMnu99.Index).Enabled = True Then Me.mnu99(objMnu99.Index).Enabled = False
                               Exit For
                           End If
                       Next
                   '­Y¤lµøµ¡¥¼³QÁôÂÃ
                   Else
                       blnFormNameMatch = False
                       For Each objMnu99 In Me.mnu99
                           If frm.Name = Me.mnu99(objMnu99.Index).Tag Then
                               If Me.mnu99(objMnu99.Index).Enabled = False Then Me.mnu99(objMnu99.Index).Enabled = True
                               blnFormNameMatch = True
                               Exit For
                           End If
                       Next
                       '­Y¤lµøµ¡¥¼¥X²{¦bµøµ¡Menu¤W
                       If blnFormNameMatch = False Then
                           For Each objMnu99 In mnu99
                               intMaxIndex = Me.mnu99(objMnu99.Index).Index
                           Next
                           Load Me.mnu99(intMaxIndex + 1)
                           Me.mnu99(intMaxIndex + 1).Caption = frm.Caption
                           Me.mnu99(intMaxIndex + 1).Tag = frm.Name
                           Me.mnu99(intMaxIndex + 1).Enabled = True
                           Exit For
                       End If
                   End If
               End If
           Next
           For Each objMnu99 In Me.mnu99
               blnFormNameMatch = False
               For Each frm In Forms
                   If frm.Name <> "mdiMain" And frm.Caption <> "" Then
                       blnFormNameMatch = False
                       If frm.Name = Me.mnu99(objMnu99.Index).Tag Then
                           blnFormNameMatch = True
                           Exit For
                       End If
                   End If
               Next
               If blnFormNameMatch = False Then
                   Exit For
               End If
           Next
           '­Yµøµ¡Menu¬Û¹ïÀ³ªº¤lµøµ¡¤£¦s¦b
           If blnFormNameMatch = False Then
               If objMnu99.Index = 0 Then
                   For Each objMnu99 In Me.mnu99
                       If objMnu99.Index <> 0 Then
                           Unload Me.mnu99(objMnu99.Index)
                       End If
                   Next
                   ii = 0
                   For Each frm In Forms
                       If frm.Name <> "mdiMain" And frm.Caption <> "" Then
                           If ii = 0 Then
                               Me.mnu99(ii).Caption = frm.Caption
                               Me.mnu99(ii).Tag = frm.Name
                               If Me.mnu99(ii).Enabled = False Then Me.mnu99(ii).Enabled = True
                           Else
                               Load Me.mnu99(ii)
                               Me.mnu99(ii).Caption = frm.Caption
                               Me.mnu99(ii).Tag = frm.Name
                               Me.mnu99(ii).Enabled = True
                           End If
                           ii = ii + 1
                       End If
                   Next
               Else
                   Unload Me.mnu99(objMnu99.Index)
               End If
           End If
       '­Y¤lµøµ¡¼Æ»Pµøµ¡menu¼Æ¬Ò¬°1
       ElseIf Forms.Count - 1 = 1 And Me.mnu99.Count = 1 Then
           For Each frm In Forms
               If frm.Name <> "mdiMain" And frm.Caption <> "" Then
                   If frm.Name <> Me.mnu99(0).Tag Then
                       For Each objMnu99 In mnu99
                           intMaxIndex = Me.mnu99(0).Index
                       Next
                       If Me.mnu99(0).Enabled = False Then Me.mnu99(0).Enabled = True
                       Me.mnu99(0).Caption = frm.Caption
                       Me.mnu99(0).Tag = frm.Name
                       Exit For
                   End If
               End If
           Next
       End If
   End If
   '­Y¦³¤lµøµ¡
   If Forms.Count - 1 >= 1 Then
      If Not mdiMain.ActiveForm Is Nothing Then 'Added by Morgan 2015/10/30
       For Each objMnu99 In Me.mnu99
           If mdiMain.ActiveForm.Name = Me.mnu99(objMnu99.Index).Tag Then
               If Me.mnu99(objMnu99.Index).Checked = False Then Me.mnu99(objMnu99.Index).Checked = True
           Else
               If Me.mnu99(objMnu99.Index).Checked = True Then Me.mnu99(objMnu99.Index).Checked = False
           End If
       Next
      End If 'Added by Morgan 2015/10/30
   End If
End Sub

Private Sub mnuPopItem_Click(Index As Integer)
   'Modified by Morgan 2021/4/26
   'frm140112.ShowNextForm Index
   frm140112.Timer2.Enabled = True
   frm140112.Timer2.Tag = Index
End Sub

'Added by Morgan 2016/1/22
Private Sub mnu43_Click(Index As Integer)
   Select Case Index
      Case 1 '­û¤uÁ~¸ê©ú²Ó
         If PUB_SalaryEnabled Then
            frm170236.Show: Exit Sub
         Else
            frm170107.setNextForm "frm170236"
         End If
      Case 2 '³Ò«O/°·«O/³Ò°hª÷©ú²Ó
         If PUB_SalaryEnabled Then
            frm170237.Show: Exit Sub
         Else
            frm170107.setNextForm "frm170237"
         End If
      Case 3 '¦~«×¦U¶µ©Ò±o©ú²Ó
         If PUB_SalaryEnabled Then
            frm170238.Show: Exit Sub
         Else
            frm170107.setNextForm "frm170238"
         End If
      Case 4 '¦~²×¼úª÷©ú²Ó
         If PUB_SalaryEnabled Then
            frm170239.Show: Exit Sub
         Else
            frm170107.setNextForm "frm170239"
         End If
      Case 5 'Á~¸ê¬d¸ß±K½X­×§ï
         frm170107.setNextForm "frm170108"
'      Case 6 'Á~¸ê¬d¸ßÂ÷½u®É¶¡³]©w­×§ï (¥ý¤£°µ)
   End Select
   
   frm170107.Show
End Sub

'Added by Morgan 2016/1/22
'Á~¸êµe­±­p®É¾¹:60¬í
Private Sub tmrSalary_Timer()
   tmrSalary.Tag = Val(tmrSalary.Tag) + 1
   If Val(tmrSalary.Tag) > 60 Then
      tmrSalary.Enabled = False
      Pub_CloseSalaryQueryForm
   End If
End Sub

'Add By Sindy 2020/5/29
'¥H¦WºÙ¨ú±oªí³æ--³q¥Î¤£¥i§R
Public Function GetForm(pFormName As String) As Form
   Select Case pFormName
   '·s¼W±M®×·|¥Î¨ìªºForm
   Case "frm180301"
         Set GetForm = frm180301
   End Select
End Function

'Added by Morgan 2021/4/22
'½Æ»s¶K¤W¼u¸õµøµ¡
Public Sub PopupMenu2(oTextBox As Control)
   If oTextBox.Enabled = True And oTextBox.Locked = False Then
      mnuPopItem2(0).Enabled = False
      mnuPopItem2(1).Enabled = False
      mnuPopItem2(2).Enabled = False
      mnuPopItem2(3).Enabled = False
      If oTextBox.SelLength > 0 Then
         mnuPopItem2(0).Enabled = True
         mnuPopItem2(1).Enabled = True
         mnuPopItem2(3).Enabled = True
      End If
      If Clipboard.GetText <> "" Then
         mnuPopItem2(2).Enabled = True
      End If
      PopupMenu mnuPop2
   End If
End Sub

'Added by Morgan 2021/4/22
'½Æ»s¶K¤W¿ï³æ
Private Sub mnuPopItem2_Click(Index As Integer)
   Select Case Index
   Case 0 '°Å¤U
      SendKeys "+{DELETE}"
   Case 1 '½Æ»s
      SendKeys "^C"
   Case 2 '¶K¤W
      SendKeys "^V"
   Case 3 '§R°£
      SendKeys "{DELETE}"
   End Select
End Sub
