VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100102_1 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "¥H¥Ó½Ð¤H¬d¸ß"
   ClientHeight    =   6080
   ClientLeft      =   3780
   ClientTop       =   3700
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6080
   ScaleWidth      =   8950
   Begin VB.CommandButton cmdMemo 
      BackColor       =   &H00C0FFC0&
      Caption         =   "¬d¸ß¸m´«¦r"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   1390
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   56
      Top             =   45
      Width           =   1050
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Åã¥Ü¦³µL®×¥ó"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   30
      TabIndex        =   55
      Top             =   270
      Width           =   1665
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "±Hµo«H¨ç-©¹¨Ó°O¿ý"
      Height          =   345
      Index           =   12
      Left            =   5790
      TabIndex        =   54
      Top             =   2190
      Width           =   1845
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "±H¥ó¬d¸ß"
      Height          =   400
      Index           =   11
      Left            =   3050
      TabIndex        =   19
      Top             =   45
      Width           =   885
   End
   Begin VB.CheckBox Check2 
      Caption         =   "§t¹ï³y"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7920
      TabIndex        =   53
      Top             =   1860
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "¦C¦L¹ï³y¸ê®Æ"
      Enabled         =   0   'False
      Height          =   300
      Index           =   10
      Left            =   7440
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   30
      Top             =   830
      Width           =   1515
   End
   Begin VB.CheckBox Check1 
      Caption         =   "§t§ë¸êªk°È¶}©Ý¸ê®Æ"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7170
      TabIndex        =   50
      Top             =   1508
      Width           =   1750
   End
   Begin VB.OptionButton Option2 
      Caption         =   "E-Mail¡G"
      Height          =   180
      Index           =   3
      Left            =   3820
      TabIndex        =   48
      Top             =   1545
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Height          =   300
      Left            =   4785
      TabIndex        =   8
      Top             =   1485
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "±µ¬¢¤H/Ápµ¸¤H"
      Height          =   345
      Index           =   9
      Left            =   3600
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   26
      Top             =   465
      Width           =   1530
   End
   Begin VB.TextBox Text11 
      Height          =   300
      Left            =   7056
      TabIndex        =   6
      Top             =   1170
      Width           =   1035
   End
   Begin VB.OptionButton Option2 
      Caption         =   "ID¡G"
      Height          =   180
      Index           =   4
      Left            =   6360
      TabIndex        =   46
      Top             =   1230
      Width           =   660
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "©¹¨Ó°O¿ý"
      Height          =   345
      Index           =   8
      Left            =   5130
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   27
      Top             =   465
      Width           =   1170
   End
   Begin VB.OptionButton Option2 
      Caption         =   "­t³d¤H¡G"
      Height          =   180
      Index           =   2
      Left            =   30
      TabIndex        =   34
      Top             =   1545
      Width           =   1100
   End
   Begin VB.CheckBox ChkPCT 
      Caption         =   "¬O§_Åã¥ÜPCT ®×"
      Height          =   225
      Left            =   3960
      TabIndex        =   13
      Top             =   2280
      Width           =   1635
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ªk°È¶i«×"
      Height          =   345
      Index           =   7
      Left            =   6300
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   28
      Top             =   465
      Width           =   1170
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "¬ÛÃö¦h¥Ó½Ð¤H"
      Height          =   345
      Index           =   6
      Left            =   7455
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   29
      Top             =   465
      Width           =   1515
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "±M§Q¬ÛÃö®×"
      Height          =   400
      Index           =   5
      Left            =   7300
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   24
      Top             =   45
      Width           =   1050
   End
   Begin VB.TextBox txtCountry 
      Height          =   300
      Index           =   1
      Left            =   2025
      MaxLength       =   4
      TabIndex        =   17
      Top             =   2820
      Width           =   852
   End
   Begin VB.TextBox txtCountry 
      Height          =   300
      Index           =   0
      Left            =   975
      MaxLength       =   4
      TabIndex        =   16
      Top             =   2820
      Width           =   852
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "°ê¤ºA4¦W±ø"
      Height          =   400
      Index           =   4
      Left            =   6180
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   23
      Top             =   45
      Width           =   1100
   End
   Begin VB.CheckBox chk 
      Caption         =   "©Ò¦³¨t²ÎÃþ§O"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   30
      TabIndex        =   42
      Top             =   30
      Width           =   1665
   End
   Begin VB.Frame Frame2 
      Height          =   350
      Left            =   5300
      TabIndex        =   39
      Top             =   750
      Width           =   2100
      Begin VB.OptionButton Option3 
         Caption         =   "¼Ò½k¤ñ¹ï"
         Height          =   180
         Index           =   1
         Left            =   1050
         TabIndex        =   41
         Top             =   144
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton Option3 
         Caption         =   "¦r­º¤ñ¹ï"
         Height          =   180
         Index           =   0
         Left            =   72
         TabIndex        =   40
         Top             =   144
         Width           =   1020
      End
   End
   Begin VB.TextBox Text8 
      Height          =   300
      Left            =   6540
      MaxLength       =   1
      TabIndex        =   10
      Top             =   1830
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   350
      Left            =   5310
      TabIndex        =   35
      Top             =   2220
      Visible         =   0   'False
      Width           =   2436
      Begin VB.OptionButton Option1 
         Caption         =   "¤é¤å"
         Height          =   180
         Index           =   2
         Left            =   1656
         TabIndex        =   38
         Top             =   135
         Width           =   732
      End
      Begin VB.OptionButton Option1 
         Caption         =   "­^¤å"
         Height          =   180
         Index           =   1
         Left            =   900
         TabIndex        =   37
         Top             =   135
         Width           =   732
      End
      Begin VB.OptionButton Option1 
         Caption         =   "¤¤¤å"
         Height          =   180
         Index           =   0
         Left            =   72
         TabIndex        =   36
         Top             =   135
         Value           =   -1  'True
         Width           =   732
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "¥Ó½Ð¤H/±µ¬¢¤H/Ápµ¸¤H¦WºÙ¡G"
      Height          =   180
      Index           =   1
      Left            =   30
      TabIndex        =   33
      Top             =   915
      Width           =   2560
   End
   Begin VB.OptionButton Option2 
      Caption         =   "¥Ó½Ð¤H½s¸¹¡G"
      Height          =   180
      Index           =   0
      Left            =   30
      TabIndex        =   32
      Top             =   585
      Value           =   -1  'True
      Width           =   1380
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   2880
      Left            =   30
      TabIndex        =   31
      Top             =   3180
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   5080
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   17
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "·s²Ó©úÅé-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   17
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      MaxLength       =   9
      TabIndex        =   4
      Top             =   510
      Width           =   1932
   End
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   975
      MaxLength       =   4
      TabIndex        =   14
      Top             =   2490
      Width           =   852
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   975
      MaxLength       =   7
      TabIndex        =   11
      Top             =   2145
      Width           =   852
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   975
      TabIndex        =   9
      Top             =   1830
      Width           =   2772
   End
   Begin VB.TextBox Text7 
      Height          =   300
      Left            =   2025
      MaxLength       =   4
      TabIndex        =   15
      Top             =   2490
      Width           =   852
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   2025
      MaxLength       =   7
      TabIndex        =   12
      Top             =   2145
      Width           =   852
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "´M§ä"
      Default         =   -1  'True
      Height          =   400
      Left            =   2450
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   18
      Top             =   45
      Width           =   600
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ãö«Y¥ø·~"
      Height          =   400
      Index           =   2
      Left            =   5280
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   22
      Top             =   45
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "®×¥ó"
      Height          =   400
      Index           =   1
      Left            =   4670
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   21
      Top             =   45
      Width           =   600
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "¥Ó½Ð¤H"
      Height          =   400
      Index           =   0
      Left            =   3940
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   20
      Top             =   45
      Width           =   720
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "µ²§ô"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   8350
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   25
      Top             =   45
      Width           =   600
   End
   Begin MSForms.TextBox Text9 
      Height          =   336
      Left            =   1152
      TabIndex        =   7
      Top             =   1476
      Width           =   1704
      VariousPropertyBits=   671105051
      BackColor       =   16777215
      Size            =   "2999;593"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   330
      Left            =   2610
      TabIndex        =   5
      Top             =   840
      Width           =   2600
      VariousPropertyBits=   671105051
      BackColor       =   16777215
      Size            =   "4586;582"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "¿é¤J¦WºÙ¤§¯S¨ú³¡¤À, ¤£­n¨ú°ê®a,¬Ù¥÷,«°¥«,¨Ò¡G¤£¥i¿é¬ü°Ó..,¼sªF..,¼s¦{.."
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   2
      Left            =   20
      TabIndex        =   52
      Top             =   1200
      Width           =   5808
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "µù¡G¬õ¦â¤£¥i©Ó±µ®×¥ó¡þ¶À©³¬°«Ý¬¡¤Æ«È¤á"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   2892
      TabIndex        =   51
      Top             =   2976
      Width           =   3420
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1980
      X2              =   1860
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "¼Ò½k¤ñ¹ï"
      Height          =   180
      Left            =   6350
      TabIndex        =   49
      Top             =   1545
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "¼Ò½k¤ñ¹ï"
      Height          =   180
      Left            =   8136
      TabIndex        =   47
      Top             =   1236
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "¼Ò½k¤ñ¹ï"
      Height          =   180
      Left            =   2976
      TabIndex        =   45
      Top             =   1548
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð°ê®a¡G"
      Height          =   180
      Left            =   60
      TabIndex        =   44
      Top             =   2850
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "¡¯¡GÂÂªº¦WºÙ¡@¢C¡G¦³§b±b¡@    ¡´¡G¯S®í«È¤á   ¡ò¡G¤£±o¥N²z¡@ ¡¿¡GµL®×¥ó"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   930
      Left            =   7770
      TabIndex        =   43
      Top             =   2140
      Width           =   1260
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1980
      X2              =   1860
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1860
      X2              =   1980
      Y1              =   2295
      Y2              =   2295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "¨t²ÎÃþ§O¡G                                                               (ALL¡G¥þ³¡)"
      Height          =   180
      Left            =   30
      TabIndex        =   3
      Top             =   1890
      Width           =   4725
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "¦¬¤å¤é´Á¡G"
      Height          =   180
      Left            =   60
      TabIndex        =   2
      Top             =   2175
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "®×¥ó©Ê½è¡G"
      Height          =   180
      Left            =   60
      TabIndex        =   1
      Top             =   2520
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "¬O§_§t¨Ó¨ç¸ê®Æ¡G           ¡]N¡G¤£§t¡^"
      Height          =   180
      Left            =   4980
      TabIndex        =   0
      Top             =   1890
      Width           =   2955
   End
End
Attribute VB_Name = "frm100102_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2024/03/13 ®³±¼A4¦W±ø¦Lªí¾÷Combo1ªºª«¥ó©Mµ{¦¡
'Memo by Lydia 2021/12/16 §ï¦¨Form2.0 ; GrdDataList§ï¦r«¬=·s²Ó©úÅé-ExtB¡BText2¡BText9
'Memo by Amy 2013/11/06 ¦X¨Ö·s«È¤á¬d¸ßfrm100132¥\¯à(¤w¦³¬d¹ï³y),®³±¼¬dµL¸ê®Æ¬d¹ï³y¥\¯à 11/7®³±¼¤¤¡B­^¡B¤é¬d¸ß¿ï¶µ
'Memo By Sindy 2012/12/3 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo By Sindy 2011/2/17 SQLDate¤wÀË¬d
'Memo By Sindy 2010/11/25 ­û¤u½s¸¹Äæ¤w­×§ï
'sonia 2010/8/26 ¤é´ÁÄæ¤w­×§ï
'2007/10/24 ®³±¼ 2006 ¦~«eªºµù¸Ñ  nickc
'Modify by Morgan 2008/8/11 ­ì±µ¬¢¤H¬d¸ßÄæ¦ì¨Ö¤J¥Ó½Ð¤H¦WºÙ¬d¸ß
Option Explicit

Dim i As Long, j As Long
Dim StrTag As String, StrToGrid As String
Dim strSql As String, lngCounter As Long, lngCounterI As Long
Public cmdState As Integer
Dim m_dbl_LeftMargin As Double
Dim m_dbl_TopMargin  As Double
Dim SeekPrintL As Integer
Dim SeekPrint As Integer
Dim m_bolPrintRight As Boolean
'Add by Amy 2013/11/06
Dim StrToPrint As String '°O¿ý½s¸¹ for ¹ï³y¦C¦L
Dim strTp(3) As String, ColName() As String
Dim PLeft() As Integer, intCounter As Integer, intRecord As Integer, intPage As Integer, kk As Integer
Dim bolPrint As Boolean '¬O§_¦³¹ï³y
'end 2013/11/06
Public IsSearchNew As Boolean 'Modify by Amy 2014/04/30 ¬d·s«È¤á
Dim m_blnColOrderAsc As Boolean 'Add by Amy 2020/06/16 Äæ¦ì¸ê®Æ¥Ñ¤p¨ì¤j±Æ§Ç
Dim strField() As String 'Add by Amy 2023/03/08
Dim strQueryChangTxt As String 'Add by Amy 2023/08/17 ¸m´«¦r
Dim m_pub_QL05 As String 'Add By Sindy 2025/8/13 ¥u°O¿ý©ó¦¹Form


'Modify by Amy 2023/08/24 +IsRelation
Private Sub SetDataListWidth(Optional ByVal IsRelation As Boolean = False)
   grdDataList.row = 0
   grdDataList.col = 0: grdDataList.Text = "V"
   grdDataList.ColWidth(0) = 200
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 1: grdDataList.Text = "½s¸¹"
   grdDataList.ColWidth(1) = 1200
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 2: grdDataList.Text = "¦WºÙ"
   grdDataList.ColWidth(2) = 4000
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 3: grdDataList.Text = "°êÄy"
   grdDataList.ColWidth(3) = 1200
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 4: grdDataList.Text = "´¼Åv¤H­û"
   grdDataList.ColWidth(4) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 5: grdDataList.Text = "ª¬ºA"
   grdDataList.ColWidth(5) = 1000
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 6: grdDataList.Text = "³Æµù"
   grdDataList.ColWidth(6) = 2000
   grdDataList.CellAlignment = flexAlignLeftCenter
   'Add by Amy 2013/11/06
   '¦]¬d¸ßªA°È¹ï³y¸ê®Æ»Ý¨Ìsp09§ì´¼Åv¤H­û¸ê®Æ,¬G¥[¥Ó½Ð°ê®a
   grdDataList.col = 7: grdDataList.Text = "¥Ó½Ð°ê®a"
   grdDataList.ColWidth(7) = 0
   '§ì¨ú¹ï³yÄæ¦ì for ¦C¦L
   grdDataList.col = 8: grdDataList.Text = "Á`¦¬¤å¸¹"
   grdDataList.ColWidth(8) = 0
   grdDataList.col = 9: grdDataList.Text = "®×¥ó©Ê½è"
   grdDataList.ColWidth(9) = 0
   grdDataList.col = 10: grdDataList.Text = "¦¬¤å¤é"
   grdDataList.ColWidth(10) = 0
   'end 2013/11/06

   'Added by Lydia 2017/02/14 ÃöÁp¥ø·~
   'Modify by Amy 2019/09/17 §ï¬°¤é´Á§PÂ_ ­ì:Äæ¦ì¼Æ
   If strSrvDate(1) < °ê¥~³¡ÃöÁp¥ø·~±Ò¥Î¤é Then 'Added by Lydia 2017/12/28
        grdDataList.col = 11: grdDataList.Text = "ÃöÁp½s¸¹"
        grdDataList.ColWidth(11) = 0
        grdDataList.col = 12: grdDataList.Text = "ÃöÁp¦WºÙ"
        grdDataList.ColWidth(12) = 0
        grdDataList.col = 13: grdDataList.Text = "ÃöÁpÃö«Y"
        grdDataList.ColWidth(13) = 0
        grdDataList.col = 14: grdDataList.Text = "ÃöÁp»¡©ú"
        grdDataList.ColWidth(14) = 0
        grdDataList.FixedCols = 0
   End If  'Added by Lydia 2017/12/28
   'end 2017/02/14
   'Modify by Amy 2022/08/19 +ORGN
   grdDataList.col = 15: grdDataList.Text = "ORGN"
   grdDataList.ColWidth(15) = 0
   grdDataList.col = 16: grdDataList.Text = "«Ý¬¡¤Æ«È¤á"
   grdDataList.ColWidth(16) = 0
   grdDataList.FixedCols = 0
   'end 2019/09/17
   
   'Modify by Amy 2023/08/24 Á×§K¨S§ï¨ì,±qstrMenu1·h¹L¨Ó
   'ÃöÁp¥ø·~
   If IsRelation = True Then
      'Added by Lydia 2017/12/05 §ï¥Ñ±Ò¥Î¤é±±¨î
      If strSrvDate(1) >= °ê¥~³¡ÃöÁp¥ø·~±Ò¥Î¤é Then
         'Added by Lydia 2017/02/14 Äæ¼e½Õ¾ã
         grdDataList.FixedCols = 3 '©T©w½s¸¹©M¦WºÙ
         Call PUB_SetMSFGridColor(Me.grdDataList, "15") '©³¦â³]©w¬°ªÅ¥Õ
         grdDataList.ColWidth(2) = 1200 '¦WºÙ
         grdDataList.ColWidth(3) = 800 '°êÄy
         grdDataList.ColWidth(6) = 1200 '³Æµù
         grdDataList.ColWidth(11) = 1000 'ÃöÁp½s¸¹
         grdDataList.ColWidth(12) = 1200 'ÃöÁp¦WºÙ
         grdDataList.ColWidth(13) = 1200 'ÃöÁpÃö«Y
         grdDataList.ColWidth(14) = 1200 'ÃöÁp»¡©ú
         'end 2017/02/14
      End If
   End If
     
End Sub

'Add by Amy 2023/03/08 ÅÜ°ÊªºÄæ¦ì
Private Sub GetField()
    ReDim strField(grdDataList.Cols - 1)
    For j = 0 To grdDataList.Cols - 1
        strField(j) = grdDataList.TextMatrix(0, j)
    Next j
End Sub

Private Function GetValue(pFieldN As String) As Integer
    Dim jj As Integer
 
    For jj = 1 To UBound(strField)
        If UCase(strField(jj)) = UCase(pFieldN) Then
            GetValue = jj
            Exit For
        End If
    Next jj
End Function
'end 2023/03/08

Private Sub chk_Click()
   If Me.chk.Value = vbChecked Then
       Me.Text3.Text = "ALL"
   Else
       Me.Text3.Text = Systemkind_g
   End If
End Sub

'Mark by Amy 2023/08/24 §ï§ì¦@¥Î
Public Sub PubShowNextData_Old()
'Dim blnPrintAdd As Boolean
'Dim ii As Integer
'Dim j As Integer
'Dim strTmp As String
'Dim strCaseNo As String 'Add by Amy 2014/04/07 ¥»©Ò®×¸¹(for ¹ï³y)
'Dim bA4Print As Variant  'Added by Lydia 2016/11/10 ¬O§_¦C¦LA4¦W±ø¿ï¶µ
'
'   'Modify by Amy 2023/03/08 Äæ¦ì§ï°ÊºA
'   Select Case cmdState
'      Case 0 '¥Ó½Ð¤H¸ê®Æ
'            Me.Enabled = False
'            For i = 1 To GrdDataList.Rows - 1
'            GrdDataList.col = 0
'            GrdDataList.row = i
'            If Trim(GrdDataList.Text) = "V" Then
'               GrdDataList.col = 0
'               GrdDataList.Text = ""
'               'Add By Sindy 2012/3/21
'               GrdDataList.col = 1
'               'Add by Amy 2019/09/17 ¬¡¤Æ«È¤áÅã¥Ü¾ã¦C¶À,­Y¦³§b±b½s¸¹©³¬°¬õ¨ä¥LÄæ¬°¶À
'               If GrdDataList.TextMatrix(GrdDataList.row, GetValue("«Ý¬¡¤Æ«È¤á")) = "0" And Right(GrdDataList.TextMatrix(GrdDataList.row, GetValue("½s¸¹")), 1) <> "¡¯" Then
'                    For j = 0 To GrdDataList.Cols - 1
'                        '§b±b
'                        If Right(GrdDataList.Text, 1) = "$" And j = 1 Then
'                            GrdDataList.CellBackColor = &HFF& '¬õ¦â
'                        '¬¡¤Æ«È¤á
'                        Else
'                            GrdDataList.col = j
'                            GrdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'               'Add by Amy 2019/08/28 +«È¤áª¬ºA¬° ¾E²¾¤£©ú/¸Ñ´²/¼o¤î/ºM¾P/°±·~/¦º¤` Åã¥Ü¶Â©³
'               'Modify by Amy 2022/05/23 ®³±¼ ¾E²¾¤£©ú ¤Î °±·~
'               ElseIf (Left(GrdDataList.Text, 1) = "Y" Or Left(GrdDataList.Text, 1) = "X" Or Left(GrdDataList.Text, 1) = "R") _
'                  And (GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¸Ñ´²" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¼o¤î" _
'                      Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "ºM¾P" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¦º¤`") Then
'                    For j = 0 To GrdDataList.Cols - 1
'                       GrdDataList.col = j
'                       GrdDataList.CellBackColor = &H0 '¶Â¦â
'                       GrdDataList.CellForeColor = &HFF00FF '¯»¬õ¦â
'                    Next j
'               'Modify by Amy 2013/12/10 +§PÂ_¹ï³y
'               ElseIf Right(GrdDataList.Text, 1) = "¡ò" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¹ï³y" Then
'                  For j = 0 To GrdDataList.Cols - 1
'                     GrdDataList.col = j
'                     GrdDataList.CellBackColor = &H8080FF
'                  Next j
'               Else
'               '2012/3/21 End
'                  For j = 0 To GrdDataList.Cols - 1
'                     If j <> 1 Then
'                         GrdDataList.col = j
'                         GrdDataList.CellBackColor = QBColor(15)
'                     End If
'                  Next j
'               End If
'               If fnSaveParentForm(Me) = False Then
'                   Me.Enabled = True
'                   Exit Sub
'               End If
'               GrdDataList.col = 1
'               Screen.MousePointer = vbHourglass
'               'Modify by Morgan 2007/12/13 ¥[§PÂ_²Ä¤@½X¤Á¤£¦Pµe­±
'               strTmp = Pub_RplStr(GrdDataList.Text)
'               Select Case Left(strTmp, 1)
'                  Case "X"
'                     If Mid(strTmp, 10, 1) = "-" Then
'                        strTmp = Left(strTmp, 9)
'                     End If
'                     frm100101_11.Show
'                     frm100101_11.Tag = strTmp
'                     frm100101_11.StrMenu
'                  Case "Y" '¥N²z¤H
'                     'Add by Sindy 98/03/05
'                     '+§PÂ_¦³Åv­­ªº¤~¯à¬d¥N²z¤Hªº®×¥ó¸ê®Æ
'                     If bolFNation = True Then
'                        If Mid(strTmp, 10, 1) = "-" Then
'                           strTmp = Left(strTmp, 9)
'                        End If
'                        frm100101_10.Show
'                        frm100101_10.Tag = strTmp
'                        frm100101_10.StrMenu
'                     '2011/5/6 add by sonia
'                     Else
'                        Me.Show
'                        MsgBox "±zµL¬d¸ß°ê¥~¥N²z¤H¸ê®ÆÅv­­¡I", vbInformation
'                     '2011/5/6 end
'                     End If
'                  Case "R"
'                     'Modify By Sindy 2009/06/24 §PÂ_¬O°ê¥~©Î¬O°ê¤º¼ç¦b«È¤á
'                     strExc(0) = "select * from potcustomer where pcu01(+)='" & Left(strTmp, 8) & "' and pcu02(+)='" & Mid(strTmp, 9, 1) & "' "
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                     strExc(2) = ""
'                     If intI = 1 Then
'                        strExc(2) = "" & RsTemp.Fields(0)
'                     End If
'                     If strExc(2) <> "" Then '°ê¥~
'                        frm100101_14.Show
'                        frm100101_14.Tag = strTmp
'                        frm100101_14.StrMenu
'                     Else '°ê¤º
'                        frm100101_21.Show
'                        frm100101_21.Tag = strTmp
'                        frm100101_21.StrMenu
'                     End If
'                  'Add by Amy 2015/03/27 +«È¤áºÝ¥­¥x±b¸¹
'                  Case "¥­"
'                     'Modify by Amy 2015/04/15 §ï¥H¥­¥x½s¸¹§ìÅv­­
'                     If PUB_ChkCustWebLimit(GrdDataList.TextMatrix(GrdDataList.row, GetValue("¦¬¤å¤é")), strUserNum) = True Then
'                        frm100101_27.Show
'                        frm100101_27.Tag = Trim(GrdDataList.TextMatrix(GrdDataList.row, GetValue("¦¬¤å¤é")))
'                        frm100101_27.StrMenu
'                     Else
'                        Me.Show
'                        MsgBox "±zµLÅv­­¬d¸ß¦¹«È¤áºÝ¥­¥x±b¸¹¡I", vbInformation
'                     End If
'                  'Add By Sindy 2009/07/22
'                  Case Else
'                     'Modify By Sindy 2012/3/21 +¤£±o¥N²z®×¥ó¤§«È¤á©Î¥N²z¤H
'                     If InStr(strTmp, "-") = 0 Then
'                        frm100101_25.Show
'                        frm100101_25.Tag = strTmp
'                        frm100101_25.StrMenu
'                     Else
'                     '2012/3/21 End
'                        frm100101_22.Show
'                        frm100101_22.Tag = strTmp
'                        frm100101_22.StrMenu
'                     End If
'                  '2009/07/22 End
'               End Select
'               'end 2007/12/13
'               Screen.MousePointer = vbDefault
'               GrdDataList.col = 0
'               GrdDataList.Text = ""
'               'Add By Sindy 2012/3/21
'               GrdDataList.col = 1
'               'Add by Amy 2019/09/17 ¬¡¤Æ«È¤áÅã¥Ü¾ã¦C¶À,­Y¦³§b±b½s¸¹©³¬°¬õ¨ä¥LÄæ¬°¶À
'               If GrdDataList.TextMatrix(GrdDataList.row, GetValue("«Ý¬¡¤Æ«È¤á")) = "0" And Right(GrdDataList.TextMatrix(GrdDataList.row, GetValue("½s¸¹")), 1) <> "¡¯" Then
'                    For j = 0 To GrdDataList.Cols - 1
'                        '§b±b
'                        If Right(GrdDataList.Text, 1) = "$" And j = 1 Then
'                            GrdDataList.CellBackColor = &HFF& '¬õ¦â
'                        '¬¡¤Æ«È¤á
'                        Else
'                            GrdDataList.col = j
'                            GrdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'               'Add by Amy 2019/08/28 +«È¤áª¬ºA¬° ¾E²¾¤£©ú/¸Ñ´²/¼o¤î/ºM¾P/°±·~/¦º¤` Åã¥Ü¶Â©³
'               'Modify by Amy 2022/05/23 ®³±¼ ¾E²¾¤£©ú ¤Î °±·~
'               ElseIf (Left(GrdDataList.Text, 1) = "Y" Or Left(GrdDataList.Text, 1) = "X" Or Left(GrdDataList.Text, 1) = "R") _
'                  And (GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¸Ñ´²" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¼o¤î" _
'                      Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "ºM¾P" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¦º¤`") Then
'                    For j = 0 To GrdDataList.Cols - 1
'                       GrdDataList.col = j
'                       GrdDataList.CellBackColor = &H0 '¶Â¦â
'                       GrdDataList.CellForeColor = &HFF00FF '¯»¬õ¦â
'                    Next j
'               'Modify by Amy 2013/12/10 +§PÂ_¹ï³y
'               ElseIf Right(GrdDataList.Text, 1) = "¡ò" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¹ï³y" Then
'                  For j = 0 To GrdDataList.Cols - 1
'                     GrdDataList.col = j
'                     GrdDataList.CellBackColor = &H8080FF
'                  Next j
'               Else
'               '2012/3/21 End
'                  For j = 0 To GrdDataList.Cols - 1
'                     If j <> 1 Then
'                         GrdDataList.col = j
'                         GrdDataList.CellBackColor = QBColor(15)
'                     End If
'                  Next j
'               End If
'               Me.Enabled = True
'               Exit Sub
'            End If
'            Next i
'            Me.Enabled = True
'      Case 1 '®×¥ó¸ê®Æ
'            Me.Enabled = False
'            For i = 1 To GrdDataList.Rows - 1
'            GrdDataList.col = 0
'            GrdDataList.row = i
'            If Trim(GrdDataList.Text) = "V" Then
'               GrdDataList.col = 0
'               GrdDataList.Text = ""
'               'Add By Sindy 2012/3/21
'               GrdDataList.col = 1
'               'Add by Amy 2019/09/17 ¬¡¤Æ«È¤áÅã¥Ü¾ã¦C¶À,­Y¦³§b±b½s¸¹©³¬°¬õ¨ä¥LÄæ¬°¶À
'                If GrdDataList.TextMatrix(GrdDataList.row, GetValue("«Ý¬¡¤Æ«È¤á")) = "0" And Right(GrdDataList.TextMatrix(GrdDataList.row, GetValue("½s¸¹")), 1) <> "¡¯" Then
'                    For j = 0 To GrdDataList.Cols - 1
'                        '§b±b
'                        If Right(GrdDataList.Text, 1) = "$" And j = 1 Then
'                            GrdDataList.CellBackColor = &HFF& '¬õ¦â
'                        '¬¡¤Æ«È¤á
'                        Else
'                            GrdDataList.col = j
'                            GrdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'               'Add by Amy 2019/08/28 +«È¤áª¬ºA¬° ¾E²¾¤£©ú/¸Ñ´²/¼o¤î/ºM¾P/°±·~/¦º¤` Åã¥Ü¶Â©³
'               'Modify by Amy 2022/05/23 ®³±¼ ¾E²¾¤£©ú ¤Î °±·~
'               ElseIf (Left(GrdDataList.Text, 1) = "Y" Or Left(GrdDataList.Text, 1) = "X" Or Left(GrdDataList.Text, 1) = "R") _
'                  And (GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¸Ñ´²" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¼o¤î" _
'                      Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "ºM¾P" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¦º¤`") Then
'                    For j = 0 To GrdDataList.Cols - 1
'                       GrdDataList.col = j
'                       GrdDataList.CellBackColor = &H0 '¶Â¦â
'                       GrdDataList.CellForeColor = &HFF00FF '¯»¬õ¦â
'                    Next j
'               'Modify by Amy 2013/12/10 +§PÂ_¹ï³y
'               ElseIf Right(GrdDataList.Text, 1) = "¡ò" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¹ï³y" Then
'                  For j = 0 To GrdDataList.Cols - 1
'                     GrdDataList.col = j
'                     GrdDataList.CellBackColor = &H8080FF
'                  Next j
'               Else
'               '2012/3/21 End
'                  For j = 0 To GrdDataList.Cols - 1
'                     If j <> 1 Then
'                         GrdDataList.col = j
'                         GrdDataList.CellBackColor = QBColor(15)
'                     End If
'                  Next j
'               End If
'               GrdDataList.col = 1
'               If Not IsNull(GrdDataList.Text) Then
'                  If fnSaveParentForm(Me) = False Then
'                      Me.Enabled = True
'                      Exit Sub
'                  End If
'
'                  'Modify by Amy 2014/04/07 +¥H¥»©Ò®×¸¹§ì®×¥ó¸ê®Æ
'                  If GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¹ï³y" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¨ä¥L¬ÛÃö¤H" Then
'                    strCaseNo = Pub_RplStr(GrdDataList.Text)
'                    strTmp = GetPrjPeopleNum1(strCaseNo)
'                  Else
'                    strTmp = Pub_RplStr(GrdDataList.Text)
'                  End If
'                  'end 2014/05/07
'
'                  Select Case Left(strTmp, 1)
'                  Case "X" '¥Ó½Ð¤H
'                     Screen.MousePointer = vbHourglass
'                     With frm100102_2
'                        .Show
'                        .Tag = strTmp
'                        'add b nickc 2007/12/21
'                        .ChkPCT = Me.ChkPCT
'                        'Modify by Amy 2014/05/07
'                        If strCaseNo <> "" Then
'                            .m_CaseNo = strCaseNo
'                            .StrMenu4 '¹ï³y¸ê®Æ¶i¤JªÌ
'                        Else
'                            'Modify by Morgan 2008/11/26
'                            '¬°¨Ï¬d¸ß®×¥óµe­±¦@¥Î±ø¥ó§ï°Ñ¼Æ¤è¦¡¶Ç»¼¥B¬d¸ßµ²ªG§ï»P¥N²z¤H¬d¸ß¬Û¦P
'                            .m_Sys = Text3
'                            .m_Type = "1"
'                            .m_Date1 = Text4
'                            .m_Date2 = Text5
'                            .m_Pty1 = Text6
'                            .m_Pty2 = Text7
'                            .m_CKind = Text8
'                            .m_Cty1 = txtCountry(0)
'                            .m_Cty2 = txtCountry(1)
'                            'end 2008/11/26
'                            .StrMenu
'                        End If
'                        'end 2014/05/07
'                    End With
'                    Screen.MousePointer = vbDefault
'
'                  Case "Y" '¥N²z¤H
'                     'Add by Morgan 2008/11/21
'                     '+§PÂ_¦³Åv­­ªº¤~¯à¬d¥N²z¤Hªº®×¥ó¸ê®Æ
'                     If bolFNation = True Then
'                        Screen.MousePointer = vbHourglass
'                        'Add by Morgan 2008/8/12
'                        If Mid(strTmp, 10, 1) = "-" Then
'                           strTmp = Left(strTmp, 9)
'                        End If
'
'                        With frm100114_2
'                        .Show
'                        .Tag = strTmp
'                        'add by nickc 2007/12/21
'                        .ChkPCT = Me.ChkPCT
'                        'Modify by Morgan 2008/11/21
'                        '¬°¨Ï¬d¸ß®×¥óµe­±¦@¥Î±ø¥ó§ï°Ñ¼Æ¤è¦¡¶Ç»¼¥B¬d¸ßµ²ªG§ï»P¥N²z¤H¬d¸ß¬Û¦P
'                        '.StrMenu2
'                        .m_Sys = Text3
'                        .m_Type = "1"
'                        .m_Date1 = Text4
'                        .m_Date2 = Text5
'                        .m_Pty1 = Text6
'                        .m_Pty2 = Text7
'                        .m_CKind = Text8
'                        .m_Cty1 = txtCountry(0)
'                        .m_Cty2 = txtCountry(1)
'                        .StrMenu
'                        'end 2008/11/21
'                        End With
'                        Screen.MousePointer = vbDefault
'                     '2011/5/6 add by sonia
'                     Else
'                        Me.Show
'                        MsgBox "±zµL¬d¸ß°ê¥~¥N²z¤H®×¥ó¸ê®ÆÅv­­¡I", vbInformation
'                     '2011/5/6 end
'                     End If
'                  Case "R" '¼ç¦b«È¤á
'                     Me.Show
'                     MsgBox "¸Ó½s¸¹¬°¼ç¦b«È¤á¤£·|¦³®×¥ó¸ê®Æ¡I", vbInformation
'                  Case Else
'                     Me.Show
'                  End Select
'                  Me.Enabled = True
'                  Exit Sub
'              End If
'            End If
'            Next i
'            Me.Enabled = True
'      Case 2 'Ãö«Y¥ø·~
'            Me.Enabled = False
'            strExc(9) = "" 'Added by Lydia 2017/08/18 ¤Ä¿ï²M³æ
'            'Added by Lydia 2017/12/05 §ï¥Ñ±Ò¥Î¤é±±¨î
'            If strSrvDate(1) < °ê¥~³¡ÃöÁp¥ø·~±Ò¥Î¤é Then
'               cnnConnection.Execute "DELETE FROM R100102 where id='" & strUserNum & "' "
'            End If
'            'end 2017/12/05
'            For i = 1 To GrdDataList.Rows - 1
'              GrdDataList.col = 0
'              GrdDataList.row = i
'              If Trim(GrdDataList.Text) = "V" Then
'                  GrdDataList.col = 0
'                  GrdDataList.Text = ""
'                  'Add By Sindy 2012/3/21
'                  GrdDataList.col = 1
'                  'Add by Amy 2019/09/17 ¬¡¤Æ«È¤áÅã¥Ü¾ã¦C¶À,­Y¦³§b±b½s¸¹©³¬°¬õ¨ä¥LÄæ¬°¶À
'                  If GrdDataList.TextMatrix(GrdDataList.row, GetValue("«Ý¬¡¤Æ«È¤á")) = "0" And Right(GrdDataList.TextMatrix(GrdDataList.row, GetValue("½s¸¹")), 1) <> "¡¯" Then
'                        For j = 0 To GrdDataList.Cols - 1
'                            '§b±b
'                            If Right(GrdDataList.Text, 1) = "$" And j = 1 Then
'                                GrdDataList.CellBackColor = &HFF& '¬õ¦â
'                            '¬¡¤Æ«È¤á
'                            Else
'                                GrdDataList.col = j
'                                GrdDataList.CellBackColor = vbYellow
'                            End If
'                        Next
'                  'Add by Amy 2019/08/28 +«È¤áª¬ºA¬° ¾E²¾¤£©ú/¸Ñ´²/¼o¤î/ºM¾P/°±·~/¦º¤` Åã¥Ü¶Â©³
'                  'Modify by Amy 2022/05/23 ®³±¼ ¾E²¾¤£©ú ¤Î °±·~
'                  ElseIf (Left(GrdDataList.Text, 1) = "Y" Or Left(GrdDataList.Text, 1) = "X" Or Left(GrdDataList.Text, 1) = "R") _
'                    And (GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¸Ñ´²" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¼o¤î" _
'                        Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "ºM¾P" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¦º¤`") Then
'                      For j = 0 To GrdDataList.Cols - 1
'                         GrdDataList.col = j
'                         GrdDataList.CellBackColor = &H0 '¶Â¦â
'                         GrdDataList.CellForeColor = &HFF00FF '¯»¬õ¦â
'                      Next j
'                  'Modify by Amy 2013/12/10 +§PÂ_¹ï³y
'                  ElseIf Right(GrdDataList.Text, 1) = "¡ò" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¹ï³y" Then
'                     For j = 0 To GrdDataList.Cols - 1
'                        GrdDataList.col = j
'                        GrdDataList.CellBackColor = &H8080FF
'                     Next j
'                  Else
'                  '2012/3/21 End
'                     For j = 0 To GrdDataList.Cols - 1
'                        If j <> 1 Then
'                            GrdDataList.col = j
'                            GrdDataList.CellBackColor = QBColor(15)
'                        End If
'                     Next j
'                  End If
'                  GrdDataList.col = 1
'                  'Add By Sindy 2011/01/03 ÀË¬d°ê¤º¥~Åv­­
'                  If CheckSR12(Pub_RplStr(GrdDataList.Text)) = True Then
'                  '2011/01/03 End
'                     Screen.MousePointer = vbHourglass
'                     'Modified by Lydia 2017/12/05 §ï¥Ñ±Ò¥Î¤é±±¨î
'                     If strSrvDate(1) < °ê¥~³¡ÃöÁp¥ø·~±Ò¥Î¤é Then
'                         Call StrMenu(Pub_RplStr(GrdDataList.Text))
'                     Else
'                         'Added by Lydia 2017/02/14 §ìÃöÁp¥ø·~§ï¦¨¼Ò²Õ,¼È¦sR100114_1
'                         'Modified by Lydia 2017/08/18 ¬O§_²M°£¥ý«e°O¿ý
'                         'j = PUB_GetR100114_1(Me.Name, Pub_RplStr(GrdDataList.Text))
'                         j = PUB_GetR100114_1(IIf(strExc(9) = "", True, False), Me.Name, Pub_RplStr(GrdDataList.Text))
'                         strExc(9) = strExc(9) & IIf(strExc(9) <> "", ",", "") & Pub_RplStr(GrdDataList.Text)
'                         'end 2017/08/18
'                     End If
'                     'end 2017/12/05
'
'                     cmdOK(2).Enabled = False
'                     Screen.MousePointer = vbDefault
'                  End If
'              End If
'            Next i
'            'Modified by Lydia 2017/12/05 §ï¥Ñ±Ò¥Î¤é±±¨î
'            If strSrvDate(1) < °ê¥~³¡ÃöÁp¥ø·~±Ò¥Î¤é Then
'                Call StrMenu1
'            Else
'                'Added Lydia 2017/02/14 §ìÃöÁp¥ø·~§ï¦¨¼Ò²Õ,¼È¦sR100114_1
'                If j > 1 Then Call StrMenu1
'            End If
'            'end 2017/12/05
'
'            Me.Enabled = True
'      Case 3 'µ²§ô
'         'Added by Lydia 2016/10/28 µ²§ô®É¶]¦C¦LA4¦W±ø²M³æ
'          If PUB_AddAddressA4List("", strExc(0)) Then
'          End If
'          If Val(strExc(0)) > 0 Then
'             'Midified by Lydia 2016/11/10 ¼W¥[©ñ±ó=§R°£°O¿ý
'             'If MsgBox("©|¦³" & strExc(0) & "±iA4¦W±ø¥¼¦C¦L¡A²{¦b¬O§_­n¦L¡H ", vbInformation + vbYesNo) = vbYes Then
'             'Modified by Lydia 2017/11/22 +°ê¤º
'             bA4Print = MsgBox("©|¦³" & strExc(0) & "±i°ê¤ºA4¦W±ø¥¼¦C¦L¡A²{¦b¬O§_­n¦L¡H (¬O:¦C¦L¡A§_:¤U¦¸¦C¦L¡A¨ú®ø:§R°£A4¦W±ø)", vbInformation + vbYesNoCancel)
'             If bA4Print = 6 Then  '¦C¦L
'                'Modified by Lydia 2017/11/03 §ï¦¨¾Þ§@¤¶­±
''                Load frm083014
''                frm083014.Hide
''                frm083014.Opt1(4).Value = True
''                frm083014.Text1(0).Text = strExc(0)
''                frm083014.Text1(3).Text = "1"
''                frm083014.Text1(4).Text = "1"
''                frm083014.SetPrinter Combo1
''                frm083014.cmdPrint_Click
''                Set Printer = Printers(SeekPrint)
''                Printer.Orientation = SeekPrintL
''                Unload frm083014
'                frm083014.iStiu = 1
'                frm083014.Show
'                Me.Hide
'                'end 2017/11/03
'             'Added by Lydia 2016/11/10
'             ElseIf bA4Print = 2 Then '¨ú®ø
'                cnnConnection.Execute "delete from AddressA4List where aal01='" & strUserNum & "' "
'             End If
'          End If
'          'end 2016/10/28
'
'          fnCloseAllFrm100
'
'      Case 4 '¦a§}±ø
'          Screen.MousePointer = vbHourglass
'          blnPrintAdd = False
'          'Modified by Morgan 2021/6/23
'          'Set Printer = Printers(Combo1.ListIndex)
'          PUB_RestorePrinter Combo1
'          'end 2021/6/23
'          For ii = 1 To Me.GrdDataList.Rows - 1
'              If Me.GrdDataList.TextMatrix(ii, GetValue("V")) = "V" Then
'                  strTmp = Pub_RplStr(Me.GrdDataList.TextMatrix(ii, GetValue("½s¸¹")))
'                  If Left(strTmp, 1) = "X" Then
'                     'Add By Sindy 2015/8/4
'                     strExc(3) = "select pcc01,pcc02 from PotCustCont where pcc01='" & Left(strTmp, 8) & "'"
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(3))
'                     If intI = 1 Then
'                        If RsTemp.RecordCount > 1 Then
'                           strExc(3) = "select pcc05 from customer,PotCustCont where cu01='" & Left(strTmp, 8) & "' and cu02='" & Mid(strTmp, 9, 1) & "' and cu01=pcc01(+) and cu127=pcc02(+)"
'                           intI = 1
'                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(3))
'                           If intI = 1 Then
'                              strExc(4) = "" & RsTemp.Fields(0)
'                           End If
''                           If MsgBox("¦¹«È¤á¦³¤@­Ó¥H¤W±µ¬¢¤H¡A¦¹¥\¯à¥u¦L¥X¹w³]±µ¬¢¤H" & strExc(4) & "¡A¬O§_½T©w¤´­n¦C¦L¡H" & vbCrLf & _
''                                     "(¨ä¥L±µ¬¢¤H½Ð¦Ü ®×¥ó¸ê®Æ¤Î¶i«×¬d¸ß ¦C¦L) ­Y­n¦C¦L¹w³]±µ¬¢¤H, ½Ð¿ï¾Ü¡u¬O¡v", vbYesNo) = vbNo Then
''                              Screen.MousePointer = vbDefault
''                              Exit Sub
''                           End If
'                           If MsgBox("¦¹«È¤á¦³¤@­Ó¥H¤W±µ¬¢¤H¡A¦¹¥\¯à¥u¦L¥X¹w³]±µ¬¢¤H" & strExc(4) & "¡A¬O§_½T©w¤´­n¦C¦L¡H" & vbCrLf & _
'                                     "­Y­n¦C¦L¡u¹w³]±µ¬¢¤H¡v, ½Ð¿ï¾Ü¡u¬O¡v", vbYesNo) = vbNo Then
'                              'Screen.MousePointer = vbDefault
'                              Call cmdOK_Click(9)
'                              Exit Sub
'                           End If
'                        End If
'                     End If
'                     '2015/8/4 END
'
'                     'Modified by Lydia 2016/10/28 §ï¦s¦bA4¦W±ø²M³æ,µ²§ô®É¶]¦C¦L
''                     blnPrintAdd = True
''                     Load frm083014
''                     frm083014.Hide
''                     frm083014.Opt1(0).Value = True
''                     'Add by Morgan 2008/8/26 +¥i¦L±µ¬¢¤H
''                     If Mid(strTmp, 10, 1) = "-" Then
''                        frm083014.m_ContactNo = Mid(strTmp, 11)
''                        strTmp = Left(strTmp, 9)
''                     End If
''                     'end 2008/8/26
''                     frm083014.Text1(0).Text = strTmp
''                     frm083014.Text1(4).Text = "1"
''                     frm083014.SetPrinter Printer.DeviceName
''                     frm083014.cmdPrint_Click
''                     Unload frm083014
'                     If PUB_AddAddressA4List(strTmp, strExc(0)) Then
'                        blnPrintAdd = True
'                     End If
'                     'Modified by Lydia 2017/11/22 +°ê¤º
'                     If Val(strExc(0)) > 0 Then cmdOK(4).Caption = "°ê¤ºA4¦W±ø (" & Val(strExc(0)) & ")"
'                     'end 2016/10/28
'
'                  End If
'              End If
'          Next ii
'          Screen.MousePointer = vbDefault
'          If blnPrintAdd = False Then
'              'Modified by Lydia 2016/11/04 ¦a§}±ø=>A4¦W±ø
'              MsgBox "½Ð¤Ä¿ï±ý¦C¦LA4¦W±øªº¸ê®Æ!!!", vbExclamation + vbOKOnly
'          Else
'              'ShowPrintOk 'Remove by Lydia 2016/10/28
'          End If
'          '¦L§¹¹w³]¦^¹w³]¦Lªí¾÷
'          'Move by
'          'Set Printer = Printers(SeekPrint)
'          'Printer.Orientation = SeekPrintL
'      Case 5
'           Me.Enabled = False
'           StrTag = ""
'           For i = 1 To GrdDataList.Rows - 1
'           GrdDataList.col = 0
'           GrdDataList.row = i
'           If Trim(GrdDataList.Text) = "V" Then
'               GrdDataList.col = 0
'               GrdDataList.Text = ""
'               'Add By Sindy 2012/3/21
'               GrdDataList.col = 1
'               'Add by Amy 2019/09/17 ¬¡¤Æ«È¤áÅã¥Ü¾ã¦C¶À,­Y¦³§b±b½s¸¹©³¬°¬õ¨ä¥LÄæ¬°¶À
'               If GrdDataList.TextMatrix(GrdDataList.row, GetValue("«Ý¬¡¤Æ«È¤á")) = "0" And Right(GrdDataList.TextMatrix(GrdDataList.row, GetValue("½s¸¹")), 1) <> "¡¯" Then
'                    For j = 0 To GrdDataList.Cols - 1
'                        '§b±b
'                        If Right(GrdDataList.Text, 1) = "$" And j = 1 Then
'                            GrdDataList.CellBackColor = &HFF& '¬õ¦â
'                        '¬¡¤Æ«È¤á
'                        Else
'                            GrdDataList.col = j
'                            GrdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'               'Add by Amy 2019/08/28 +«È¤áª¬ºA¬° ¾E²¾¤£©ú/¸Ñ´²/¼o¤î/ºM¾P/°±·~/¦º¤` Åã¥Ü¶Â©³
'               'Modify by Amy 2022/05/23 ®³±¼ ¾E²¾¤£©ú ¤Î °±·~
'               ElseIf (Left(GrdDataList.Text, 1) = "Y" Or Left(GrdDataList.Text, 1) = "X" Or Left(GrdDataList.Text, 1) = "R") _
'                  And (GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¸Ñ´²" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¼o¤î" _
'                      Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "ºM¾P" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¦º¤`") Then
'                    For j = 0 To GrdDataList.Cols - 1
'                       GrdDataList.col = j
'                       GrdDataList.CellBackColor = &H0 '¶Â¦â
'                       GrdDataList.CellForeColor = &HFF00FF '¯»¬õ¦â
'                    Next j
'               'Modify by Amy 2013/12/10 +§PÂ_¹ï³y
'               ElseIf Right(GrdDataList.Text, 1) = "¡ò" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¹ï³y" Then
'                  For j = 0 To GrdDataList.Cols - 1
'                     GrdDataList.col = j
'                     GrdDataList.CellBackColor = &H8080FF
'                  Next j
'               Else
'               '2012/3/21 End
'                  For j = 0 To GrdDataList.Cols - 1
'                     If j <> 1 Then
'                         GrdDataList.col = j
'                         GrdDataList.CellBackColor = QBColor(15)
'                     End If
'                  Next j
'               End If
'               GrdDataList.col = 1
'               If Not IsNull(GrdDataList.Text) Then
'                  If fnSaveParentForm(Me) = False Then
'                      Me.Enabled = True
'                      Exit Sub
'                  End If
'                  Screen.MousePointer = vbHourglass
'                  frm100101_h.Show
'                  frm100101_h.KeyString = Pub_RplStr(GrdDataList.Text)
'                  frm100101_h.SearchKind = "«È¤á½s¸¹"
'                  frm100101_h.StrMenu
'                  Screen.MousePointer = vbDefault
'                  Me.Enabled = True
'                  Exit Sub
'               End If
'           End If
'           Next i
'           Me.Enabled = True
'      Case 6
'           Me.Enabled = False
'           StrTag = ""
'           For i = 1 To GrdDataList.Rows - 1
'           GrdDataList.col = 0
'           GrdDataList.row = i
'           If Trim(GrdDataList.Text) = "V" Then
'               GrdDataList.col = 0
'               GrdDataList.Text = ""
'               'Add By Sindy 2012/3/21
'               GrdDataList.col = 1
'               'Add by Amy 2019/09/17 ¬¡¤Æ«È¤áÅã¥Ü¾ã¦C¶À,­Y¦³§b±b½s¸¹©³¬°¬õ¨ä¥LÄæ¬°¶À
'               If GrdDataList.TextMatrix(GrdDataList.row, GetValue("«Ý¬¡¤Æ«È¤á")) = "0" And Right(GrdDataList.TextMatrix(GrdDataList.row, GetValue("½s¸¹")), 1) <> "¡¯" Then
'                    For j = 0 To GrdDataList.Cols - 1
'                        '§b±b
'                        If Right(GrdDataList.Text, 1) = "$" And j = 1 Then
'                            GrdDataList.CellBackColor = &HFF& '¬õ¦â
'                        '¬¡¤Æ«È¤á
'                        Else
'                            GrdDataList.col = j
'                            GrdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'               'Add by Amy 2019/08/28 +«È¤áª¬ºA¬° ¾E²¾¤£©ú/¸Ñ´²/¼o¤î/ºM¾P/°±·~/¦º¤` Åã¥Ü¶Â©³
'               'Modify by Amy 2022/05/23 ®³±¼ ¾E²¾¤£©ú ¤Î °±·~
'               ElseIf (Left(GrdDataList.Text, 1) = "Y" Or Left(GrdDataList.Text, 1) = "X" Or Left(GrdDataList.Text, 1) = "R") _
'                  And (GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¸Ñ´²" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¼o¤î" _
'                      Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "ºM¾P" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¦º¤`") Then
'                    For j = 0 To GrdDataList.Cols - 1
'                       GrdDataList.col = j
'                       GrdDataList.CellBackColor = &H0 '¶Â¦â
'                       GrdDataList.CellForeColor = &HFF00FF '¯»¬õ¦â
'                    Next j
'               'Modify by Amy 2013/12/10 +§PÂ_¹ï³y
'               ElseIf Right(GrdDataList.Text, 1) = "¡ò" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¹ï³y" Then
'                  For j = 0 To GrdDataList.Cols - 1
'                     GrdDataList.col = j
'                     GrdDataList.CellBackColor = &H8080FF
'                  Next j
'               Else
'               '2012/3/21 End
'                  For j = 0 To GrdDataList.Cols - 1
'                     If j <> 1 Then
'                         GrdDataList.col = j
'                         GrdDataList.CellBackColor = QBColor(15)
'                     End If
'                  Next j
'               End If
'               GrdDataList.col = 1
'               If Not IsNull(GrdDataList.Text) Then
'                  If fnSaveParentForm(Me) = False Then
'                      Me.Enabled = True
'                      Exit Sub
'                  End If
'                  Screen.MousePointer = vbHourglass
'                  frm100102_4.Show
'                  frm100102_4.KeyString = Pub_RplStr(GrdDataList.Text)
'                  frm100102_4.StrMenu
'                  Screen.MousePointer = vbDefault
'                  Me.Enabled = True
'                  Exit Sub
'               End If
'           End If
'           Next i
'           Me.Enabled = True
'      Case 7 'ªk°È®×¥ó
'            Me.Enabled = False
'            For i = 1 To GrdDataList.Rows - 1
'            GrdDataList.col = 0
'            GrdDataList.row = i
'            If Trim(GrdDataList.Text) = "V" Then
'               GrdDataList.col = 0
'               GrdDataList.Text = ""
'               'Add By Sindy 2012/3/21
'               GrdDataList.col = 1
'               'Add by Amy 2019/09/17 ¬¡¤Æ«È¤áÅã¥Ü¾ã¦C¶À,­Y¦³§b±b½s¸¹©³¬°¬õ¨ä¥LÄæ¬°¶À
'               If GrdDataList.TextMatrix(GrdDataList.row, GetValue("«Ý¬¡¤Æ«È¤á")) = "0" And Right(GrdDataList.TextMatrix(GrdDataList.row, GetValue("½s¸¹")), 1) <> "¡¯" Then
'                    For j = 0 To GrdDataList.Cols - 1
'                        '§b±b
'                        If Right(GrdDataList.Text, 1) = "$" And j = 1 Then
'                            GrdDataList.CellBackColor = &HFF& '¬õ¦â
'                        '¬¡¤Æ«È¤á
'                        Else
'                            GrdDataList.col = j
'                            GrdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'               'Add by Amy 2019/08/28 +«È¤áª¬ºA¬° ¾E²¾¤£©ú/¸Ñ´²/¼o¤î/ºM¾P/°±·~/¦º¤` Åã¥Ü¶Â©³
'               'Modify by Amy 2022/05/23 ®³±¼ ¾E²¾¤£©ú ¤Î °±·~
'               ElseIf (Left(GrdDataList.Text, 1) = "Y" Or Left(GrdDataList.Text, 1) = "X" Or Left(GrdDataList.Text, 1) = "R") _
'                  And (GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¸Ñ´²" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¼o¤î" _
'                      Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "ºM¾P" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¦º¤`") Then
'                    For j = 0 To GrdDataList.Cols - 1
'                       GrdDataList.col = j
'                       GrdDataList.CellBackColor = &H0 '¶Â¦â
'                       GrdDataList.CellForeColor = &HFF00FF '¯»¬õ¦â
'                    Next j
'               'Modify by Amy 2013/12/10 +§PÂ_¹ï³y
'               ElseIf Right(GrdDataList.Text, 1) = "¡ò" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¹ï³y" Then
'                  For j = 0 To GrdDataList.Cols - 1
'                     GrdDataList.col = j
'                     GrdDataList.CellBackColor = &H8080FF
'                  Next j
'               Else
'               '2012/3/21 End
'                  For j = 0 To GrdDataList.Cols - 1
'                     If j <> 1 Then
'                         GrdDataList.col = j
'                         GrdDataList.CellBackColor = QBColor(15)
'                     End If
'                  Next j
'               End If
'               GrdDataList.col = 1
'               If Not IsNull(GrdDataList.Text) Then
'                  If fnSaveParentForm(Me) = False Then
'                      Me.Enabled = True
'                      Exit Sub
'                  End If
'                  '¥Ó½Ð¤H
'                  If UCase(Mid(GrdDataList.Text, 1, 1)) = "X" Then
'                     Screen.MousePointer = vbHourglass
'                     With frm100102_2
'                     .Show
'                     .Tag = Pub_RplStr(GrdDataList.Text)
'                     'add b nickc 2007/12/21
'                     .ChkPCT = Me.ChkPCT
'                     .bolIsL = True
'                     'Modify by Morgan 2008/11/26
'                     '¬°¨Ï¬d¸ß®×¥óµe­±¦@¥Î±ø¥ó§ï°Ñ¼Æ¤è¦¡¶Ç»¼¥B¬d¸ßµ²ªG§ï»P¥N²z¤H¬d¸ß¬Û¦P
'                     .bolIsL = True
'                     .m_Sys = Text3
'                     .m_Type = "1"
'                     .m_Date1 = Text4
'                     .m_Date2 = Text5
'                     .m_Pty1 = Text6
'                     .m_Pty2 = Text7
'                     .m_CKind = Text8
'                     .m_Cty1 = txtCountry(0)
'                     .m_Cty2 = txtCountry(1)
'                     'end 2008/11/26
'                     .StrMenu
'                     End With
'                     Screen.MousePointer = vbDefault
'                  '¥N²z¤H
'                  Else
'                     'Add by Morgan 2008/11/21
'                     '+§PÂ_¦³Åv­­ªº¤~¯à¬d¥N²z¤Hªº®×¥ó¸ê®Æ
'                     If bolFNation = True Then
'                        Screen.MousePointer = vbHourglass
'                        With frm100114_2
'                        .Show
'                        .Tag = Pub_RplStr(GrdDataList.Text)
'                        'add b nickc 2007/12/21
'                        .ChkPCT = Me.ChkPCT
'                        'Modify by Morgan 2008/11/21
'                        '¬°¨Ï¬d¸ß®×¥óµe­±¦@¥Î±ø¥ó§ï°Ñ¼Æ¤è¦¡¶Ç»¼¥B¬d¸ßµ²ªG§ï»P¥N²z¤H¬d¸ß¬Û¦P
'                        '.StrMenu2
'                        .bolIsL = True
'                        .m_Sys = Text3
'                        .m_Type = "1"
'                        .m_Date1 = Text4
'                        .m_Date2 = Text5
'                        .m_Pty1 = Text6
'                        .m_Pty2 = Text7
'                        .m_CKind = Text8
'                        .m_Cty1 = txtCountry(0)
'                        .m_Cty2 = txtCountry(1)
'                        .StrMenu
'                        'end 2008/11/21
'                        End With
'                        Screen.MousePointer = vbDefault
'                     '2011/5/6 add by sonia
'                     Else
'                        Me.Show
'                        MsgBox "±zµL¬d¸ß°ê¥~¥N²z¤H®×¥ó¸ê®ÆÅv­­¡I", vbInformation
'                     '2011/5/6 end
'                     End If
'                  End If
'                  Me.Enabled = True
'                  Exit Sub
'              End If
'            End If
'            Next i
'            Me.Enabled = True
'      'Add by Morgan 2007/12/14
'      Case 8 '©¹¨Ó°O¿ý
'            Me.Enabled = False
'            For i = 1 To GrdDataList.Rows - 1
'            GrdDataList.col = 0
'            GrdDataList.row = i
'            If Trim(GrdDataList.Text) = "V" Then
'               GrdDataList.col = 0
'               GrdDataList.Text = ""
'               'Add By Sindy 2012/3/21
'               GrdDataList.col = 1
'               'Add by Amy 2019/09/17 ¬¡¤Æ«È¤áÅã¥Ü¾ã¦C¶À,­Y¦³§b±b½s¸¹©³¬°¬õ¨ä¥LÄæ¬°¶À
'               If GrdDataList.TextMatrix(GrdDataList.row, GetValue("«Ý¬¡¤Æ«È¤á")) = "0" And Right(GrdDataList.TextMatrix(GrdDataList.row, GetValue("½s¸¹")), 1) <> "¡¯" Then
'                    For j = 0 To GrdDataList.Cols - 1
'                        '§b±b
'                        If Right(GrdDataList.Text, 1) = "$" And j = 1 Then
'                            GrdDataList.CellBackColor = &HFF& '¬õ¦â
'                        '¬¡¤Æ«È¤á
'                        Else
'                            GrdDataList.col = j
'                            GrdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'               'Add by Amy 2019/08/28 +«È¤áª¬ºA¬° ¾E²¾¤£©ú/¸Ñ´²/¼o¤î/ºM¾P/°±·~/¦º¤` Åã¥Ü¶Â©³
'               'Modify by Amy 2022/05/23 ®³±¼ ¾E²¾¤£©ú ¤Î °±·~
'               ElseIf (Left(GrdDataList.Text, 1) = "Y" Or Left(GrdDataList.Text, 1) = "X" Or Left(GrdDataList.Text, 1) = "R") _
'                  And (GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¸Ñ´²" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¼o¤î" _
'                      Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "ºM¾P" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¦º¤`") Then
'                    For j = 0 To GrdDataList.Cols - 1
'                       GrdDataList.col = j
'                       GrdDataList.CellBackColor = &H0 '¶Â¦â
'                       GrdDataList.CellForeColor = &HFF00FF '¯»¬õ¦â
'                    Next j
'               'Modify by Amy 2013/12/10 +§PÂ_¹ï³y
'               ElseIf Right(GrdDataList.Text, 1) = "¡ò" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¹ï³y" Then
'                  For j = 0 To GrdDataList.Cols - 1
'                     GrdDataList.col = j
'                     GrdDataList.CellBackColor = &H8080FF
'                  Next j
'               Else
'               '2012/3/21 End
'                  For j = 0 To GrdDataList.Cols - 1
'                     If j <> 1 Then
'                         GrdDataList.col = j
'                         GrdDataList.CellBackColor = QBColor(15)
'                     End If
'                  Next j
'               End If
'               If fnSaveParentForm(Me) = False Then
'                   Me.Enabled = True
'                   Exit Sub
'               End If
'               GrdDataList.col = 1
'               Screen.MousePointer = vbHourglass
'               strTmp = Pub_RplStr(GrdDataList.Text)
'
'               'Modify By Sindy 2010/02/23 §PÂ_¬O°ê¥~©Î¬O°ê¤º¼ç¦b«È¤á
'               '«È¤áÀÉ
'               strExc(3) = "select cu12,cu13 from customer where cu01(+)='" & Left(strTmp, 8) & "' and cu02(+)='" & Mid(strTmp, 9, 1) & "' "
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(3))
'               strExc(4) = ""
'               If intI = 1 Then
'                  strExc(4) = "" & RsTemp.Fields("cu12")
'               End If
'               '¼ç¦b«È¤áÀÉ
'               strExc(0) = "select * from potcustomer where pcu01(+)='" & Left(strTmp, 8) & "' and pcu02(+)='" & Mid(strTmp, 9, 1) & "' "
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               strExc(2) = ""
'               If intI = 1 Then
'                  strExc(2) = "" & RsTemp.Fields(0)
'               End If
''               If strExc(2) <> "" Or Left(Trim(strTmp), 1) = "Y" Or Left(Trim(strExc(4)), 1) = "F" Then '°ê¥~
'                  frm100101_15.Show
'                  frm100101_15.Tag = strTmp
'                  'Modify By Sindy 2020/5/18
'                  'Modify By Sindy 2020/5/25 + (Left(Trim(strTmp), 1) = "Y" And Left(Pub_StrUserSt03, 1) = "F")
'                  'If strExc(2) <> "" Or Left(Trim(strTmp), 1) = "Y" Or Left(Trim(strExc(4)), 1) = "F" Then '°ê¥~
'                  'Modify By Sindy 2021/3/25 + Or Left(Trim(strTmp), 1) = "¥­"
'                  If strExc(2) <> "" Or _
'                     (Left(Trim(strTmp), 1) = "Y" And Left(Pub_StrUserSt03, 1) = "F") Or _
'                     Left(Trim(strExc(4)), 1) = "F" Or Pub_StrUserSt03 = "M51" Or Left(Trim(strTmp), 1) = "¥­" Then '°ê¥~
'                     frm100101_15.m_quyDataKind = 0 '°ê¥~
'                     frm100101_15.StrMenu
'                  Else
'                     frm100101_15.m_quyDataKind = 1 '°ê¤º
'                     frm100101_15.StrMenu2
'                  End If
'                  '2020/5/18 END
''               Else
''                  frm100101_20.Show
''                  frm100101_20.Tag = strTmp
''                  frm100101_20.StrMenu
''               End If
'
'               Screen.MousePointer = vbDefault
'               GrdDataList.col = 0
'               GrdDataList.Text = ""
'               'Add By Sindy 2012/3/21
'               GrdDataList.col = 1
'               'Add by Amy 2019/09/17 ¬¡¤Æ«È¤áÅã¥Ü¾ã¦C¶À,­Y¦³§b±b½s¸¹©³¬°¬õ¨ä¥LÄæ¬°¶À
'               If GrdDataList.TextMatrix(GrdDataList.row, GetValue("«Ý¬¡¤Æ«È¤á")) = "0" And Right(GrdDataList.TextMatrix(GrdDataList.row, GetValue("½s¸¹")), 1) <> "¡¯" Then
'                    For j = 0 To GrdDataList.Cols - 1
'                        '§b±b
'                        If Right(GrdDataList.Text, 1) = "$" And j = 1 Then
'                            GrdDataList.CellBackColor = &HFF& '¬õ¦â
'                        '¬¡¤Æ«È¤á
'                        Else
'                            GrdDataList.col = j
'                            GrdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'               'Add by Amy 2019/08/28 +«È¤áª¬ºA¬° ¾E²¾¤£©ú/¸Ñ´²/¼o¤î/ºM¾P/°±·~/¦º¤` Åã¥Ü¶Â©³
'               'Modify by Amy 2022/05/23 ®³±¼ ¾E²¾¤£©ú ¤Î °±·~
'               ElseIf (Left(GrdDataList.Text, 1) = "Y" Or Left(GrdDataList.Text, 1) = "X" Or Left(GrdDataList.Text, 1) = "R") _
'                  And (GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¸Ñ´²" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¼o¤î" _
'                      Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "ºM¾P" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¦º¤`") Then
'                    For j = 0 To GrdDataList.Cols - 1
'                       GrdDataList.col = j
'                       GrdDataList.CellBackColor = &H0 '¶Â¦â
'                       GrdDataList.CellForeColor = &HFF00FF '¯»¬õ¦â
'                    Next j
'               'Modify by Amy 2013/12/10 +§PÂ_¹ï³y
'               ElseIf Right(GrdDataList.Text, 1) = "¡ò" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¹ï³y" Then
'                  For j = 0 To GrdDataList.Cols - 1
'                     GrdDataList.col = j
'                     GrdDataList.CellBackColor = &H8080FF
'                  Next j
'               Else
'               '2012/3/21 End
'                  For j = 0 To GrdDataList.Cols - 1
'                     If j <> 1 Then
'                         GrdDataList.col = j
'                         GrdDataList.CellBackColor = QBColor(15)
'                     End If
'                  Next j
'               End If
'               Me.Enabled = True
'               Exit Sub
'            End If
'            Next i
'            Me.Enabled = True
'      'Add by Morgan 2008/7/23
'      Case 9 'Ápµ¸¤H
'            Me.Enabled = False
'            For i = 1 To GrdDataList.Rows - 1
'            GrdDataList.col = 0
'            GrdDataList.row = i
'            If Trim(GrdDataList.Text) = "V" Then
'               GrdDataList.col = 0
'               GrdDataList.Text = ""
'               'Add By Sindy 2012/3/21
'               GrdDataList.col = 1
'               'Add by Amy 2019/09/17 ¬¡¤Æ«È¤áÅã¥Ü¾ã¦C¶À,­Y¦³§b±b½s¸¹©³¬°¬õ¨ä¥LÄæ¬°¶À
'               If GrdDataList.TextMatrix(GrdDataList.row, GetValue("«Ý¬¡¤Æ«È¤á")) = "0" And Right(GrdDataList.TextMatrix(GrdDataList.row, GetValue("½s¸¹")), 1) <> "¡¯" Then
'                    For j = 0 To GrdDataList.Cols - 1
'                        '§b±b
'                        If Right(GrdDataList.Text, 1) = "$" And j = 1 Then
'                            GrdDataList.CellBackColor = &HFF& '¬õ¦â
'                        '¬¡¤Æ«È¤á
'                        Else
'                            GrdDataList.col = j
'                            GrdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'               'Add by Amy 2019/08/28 +«È¤áª¬ºA¬° ¾E²¾¤£©ú/¸Ñ´²/¼o¤î/ºM¾P/°±·~/¦º¤` Åã¥Ü¶Â©³
'               'Modify by Amy 2022/05/23 ®³±¼ ¾E²¾¤£©ú ¤Î °±·~
'               ElseIf (Left(GrdDataList.Text, 1) = "Y" Or Left(GrdDataList.Text, 1) = "X" Or Left(GrdDataList.Text, 1) = "R") _
'                  And (GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¸Ñ´²" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¼o¤î" _
'                      Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "ºM¾P" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¦º¤`") Then
'                    For j = 0 To GrdDataList.Cols - 1
'                       GrdDataList.col = j
'                       GrdDataList.CellBackColor = &H0 '¶Â¦â
'                       GrdDataList.CellForeColor = &HFF00FF '¯»¬õ¦â
'                    Next j
'               'Modify by Amy 2013/12/10 +§PÂ_¹ï³y
'               ElseIf Right(GrdDataList.Text, 1) = "¡ò" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¹ï³y" Then
'                  For j = 0 To GrdDataList.Cols - 1
'                     GrdDataList.col = j
'                     GrdDataList.CellBackColor = &H8080FF
'                  Next j
'               Else
'               '2012/3/21 End
'                  For j = 0 To GrdDataList.Cols - 1
'                     If j <> 1 Then
'                         GrdDataList.col = j
'                         GrdDataList.CellBackColor = QBColor(15)
'                     End If
'                  Next j
'               End If
'               If fnSaveParentForm(Me) = False Then
'                   Me.Enabled = True
'                   Exit Sub
'               End If
'               GrdDataList.col = 1
'               Screen.MousePointer = vbHourglass
'               strTmp = Pub_RplStr(GrdDataList.Text)
'               'Modify by Morgan 2008/8/5 °ê¤º¥~«È¤á¶]¤£¦Pµe­±
'               Select Case Left(strTmp, 1)
'                  'Add by Morgan 2008/9/1 ¼ç¦b«È¤á¶]¥Ó½Ð¤H¸ê®Æµe­±
'                  Case "R"
'                     frm100101_14.Show
'                     frm100101_14.Tag = strTmp
'                     frm100101_14.StrMenu
'                  Case Else
'                     strExc(2) = "F"
'                     If Left(strTmp, 1) = "X" Then
'                        strExc(0) = "select st03 from customer,staff where cu01(+)='" & Left(strTmp, 8) & "' and cu02(+)='" & Mid(strTmp, 9, 1) & "' and st01(+)=cu13"
'                        intI = 1
'                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                        If intI = 1 Then
'                           strExc(2) = "" & RsTemp.Fields(0)
'                        End If
'                     End If
'                     If Left(strExc(2), 1) = "F" Then
'                        frm100101_17.Show
'                        frm100101_17.Tag = strTmp
'                        frm100101_17.StrMenu
'                     Else
'                        frm100101_18.Show
'                        'Added by Lydia 2016/10/28
'                        frm100101_18.SetParent Me
'                        frm100101_18.Label2(1).Visible = False
'                        frm100101_18.Combo1.Visible = False
'                        frm100101_18.CmdOk1(1).Visible = False
'                        frm100101_18.CmdOk1(2).Caption = Me.cmdOK(4).Caption
'                        'end 2016/10/28
'                        frm100101_18.Tag = strTmp
'                        frm100101_18.CmdOk1(2).Enabled = m_bolPrintRight 'Add by Morgan 2008/8/26
'                        frm100101_18.StrMenu
'                     End If
'               End Select
'               'end 2008/8/5
'               Screen.MousePointer = vbDefault
'               GrdDataList.col = 0
'               GrdDataList.Text = ""
'               'Add By Sindy 2012/3/21
'               GrdDataList.col = 1
'               'Add by Amy 2019/09/17 ¬¡¤Æ«È¤áÅã¥Ü¾ã¦C¶À,­Y¦³§b±b½s¸¹©³¬°¬õ¨ä¥LÄæ¬°¶À
'               If GrdDataList.TextMatrix(GrdDataList.row, GetValue("«Ý¬¡¤Æ«È¤á")) = "0" And Right(GrdDataList.TextMatrix(GrdDataList.row, GetValue("½s¸¹")), 1) <> "¡¯" Then
'                    For j = 0 To GrdDataList.Cols - 1
'                        '§b±b
'                        If Right(GrdDataList.Text, 1) = "$" And j = 1 Then
'                            GrdDataList.CellBackColor = &HFF& '¬õ¦â
'                        '¬¡¤Æ«È¤á
'                        Else
'                            GrdDataList.col = j
'                            GrdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'               'Add by Amy 2019/08/28 +«È¤áª¬ºA¬° ¾E²¾¤£©ú/¸Ñ´²/¼o¤î/ºM¾P/°±·~/¦º¤` Åã¥Ü¶Â©³
'               'Modify by Amy 2022/05/23 ®³±¼ ¾E²¾¤£©ú ¤Î °±·~
'               ElseIf (Left(GrdDataList.Text, 1) = "Y" Or Left(GrdDataList.Text, 1) = "X" Or Left(GrdDataList.Text, 1) = "R") _
'                  And (GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¸Ñ´²" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¼o¤î" _
'                      Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "ºM¾P" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¦º¤`") Then
'                    For j = 0 To GrdDataList.Cols - 1
'                       GrdDataList.col = j
'                       GrdDataList.CellBackColor = &H0 '¶Â¦â
'                       GrdDataList.CellForeColor = &HFF00FF '¯»¬õ¦â
'                    Next j
'               'Modify by Amy 2013/12/10 +§PÂ_¹ï³y
'               ElseIf Right(GrdDataList.Text, 1) = "¡ò" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¹ï³y" Then
'                  For j = 0 To GrdDataList.Cols - 1
'                     GrdDataList.col = j
'                     GrdDataList.CellBackColor = &H8080FF
'                  Next j
'               Else
'               '2012/3/21 End
'                  For j = 0 To GrdDataList.Cols - 1
'                     If j <> 1 Then
'                         GrdDataList.col = j
'                         GrdDataList.CellBackColor = QBColor(15)
'                     End If
'                  Next j
'               End If
'               Me.Enabled = True
'               Exit Sub
'            End If
'            Next i
'            Me.Enabled = True
'      'Add by Amy 2013/11/06
'      Case 10 '¦C¦L¹ï³y¸ê®Æ
'            'Modify by Amy 2014/02/21 §ï¦L¼È¦s¸ê®Æ
'            'PrintDataA4
'            PrintDataA4_Temp
'            'end 2014/02/21
'      'Add By Sindy 2014/5/12
'      Case 11 '¥H¥Ó½Ð¤H¬d³Ìªñ(¤@­Ó¤ë)¥H¤ºªº±H°e¸ê®Æ
'            Me.Enabled = False
'            For i = 1 To GrdDataList.Rows - 1
'               GrdDataList.col = 0
'               GrdDataList.row = i
'               If Trim(GrdDataList.Text) = "V" Then
'                  GrdDataList.col = 0
'                  GrdDataList.Text = ""
'                  GrdDataList.col = 1
'                  'Add by Amy 2019/09/17 ¬¡¤Æ«È¤áÅã¥Ü¾ã¦C¶À,­Y¦³§b±b½s¸¹©³¬°¬õ¨ä¥LÄæ¬°¶À
'                  If GrdDataList.TextMatrix(GrdDataList.row, GetValue("«Ý¬¡¤Æ«È¤á")) = "0" And Right(GrdDataList.TextMatrix(GrdDataList.row, GetValue("½s¸¹")), 1) <> "¡¯" Then
'                    For j = 0 To GrdDataList.Cols - 1
'                        '§b±b
'                        If Right(GrdDataList.Text, 1) = "$" And j = 1 Then
'                            GrdDataList.CellBackColor = &HFF& '¬õ¦â
'                        '¬¡¤Æ«È¤á
'                        Else
'                            GrdDataList.col = j
'                            GrdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'                  'Add by Amy 2019/08/28 +«È¤áª¬ºA¬° ¾E²¾¤£©ú/¸Ñ´²/¼o¤î/ºM¾P/°±·~/¦º¤` Åã¥Ü¶Â©³
'                  'Modify by Amy 2022/05/23 ®³±¼ ¾E²¾¤£©ú ¤Î °±·~
'                  ElseIf (Left(GrdDataList.Text, 1) = "Y" Or Left(GrdDataList.Text, 1) = "X" Or Left(GrdDataList.Text, 1) = "R") _
'                    And (GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¸Ñ´²" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¼o¤î" _
'                        Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "ºM¾P" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¦º¤`") Then
'                      For j = 0 To GrdDataList.Cols - 1
'                         GrdDataList.col = j
'                         GrdDataList.CellBackColor = &H0 '¶Â¦â
'                         GrdDataList.CellForeColor = &HFF00FF '¯»¬õ¦â
'                      Next j
'                  '§PÂ_¬O§_¬°¹ï³y,Åã¥Ü¤£¦PÃC¦â
'                  ElseIf Right(GrdDataList.Text, 1) = "¡ò" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¹ï³y" Then
'                     For j = 0 To GrdDataList.Cols - 1
'                        GrdDataList.col = j
'                        GrdDataList.CellBackColor = &H8080FF
'                     Next j
'                  Else
'                     For j = 0 To GrdDataList.Cols - 1
'                        If j <> 1 Then
'                            GrdDataList.col = j
'                            GrdDataList.CellBackColor = QBColor(15)
'                        End If
'                     Next j
'                  End If
'                  GrdDataList.col = 1
'                  strTmp = Pub_RplStr(GrdDataList.Text)
'                  If Left(Trim(strTmp), 1) = "X" Then
'                     Screen.MousePointer = vbHourglass
'                     If fnSaveParentForm(Me) = False Then
'                         Me.Enabled = True
'                         Exit Sub
'                     End If
'                     If Mid(strTmp, 10, 1) = "-" Then
'                        strTmp = Left(strTmp, 9)
'                     End If
'                     frm210145.intWorkItem = 0
'                     frm210145.Show
'                     frm210145.Tag = strTmp
'                     frm210145.lblAppl = GrdDataList.TextMatrix(i, GetValue("½s¸¹")) & GrdDataList.TextMatrix(i, GetValue("¦WºÙ"))
'                     Call frm210145.QueryData(False)
'                     Screen.MousePointer = vbDefault
'                  End If
'                  GrdDataList.col = 0
'                  GrdDataList.Text = ""
'                  GrdDataList.col = 1
'                  'Add by Amy 2019/09/17 ¬¡¤Æ«È¤áÅã¥Ü¾ã¦C¶À,­Y¦³§b±b½s¸¹©³¬°¬õ¨ä¥LÄæ¬°¶À
'                  If GrdDataList.TextMatrix(GrdDataList.row, GetValue("«Ý¬¡¤Æ«È¤á")) = "0" And Right(GrdDataList.TextMatrix(GrdDataList.row, GetValue("½s¸¹")), 1) <> "¡¯" Then
'                    For j = 0 To GrdDataList.Cols - 1
'                        '§b±b
'                        If Right(GrdDataList.Text, 1) = "$" And j = 1 Then
'                            GrdDataList.CellBackColor = &HFF& '¬õ¦â
'                        '¬¡¤Æ«È¤á
'                        Else
'                            GrdDataList.col = j
'                            GrdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'                  'Add by Amy 2019/08/28 +«È¤áª¬ºA¬° ¾E²¾¤£©ú/¸Ñ´²/¼o¤î/ºM¾P/°±·~/¦º¤` Åã¥Ü¶Â©³
'                  'Modify by Amy 2022/05/23 ®³±¼ ¾E²¾¤£©ú ¤Î °±·~
'                  ElseIf (Left(GrdDataList.Text, 1) = "Y" Or Left(GrdDataList.Text, 1) = "X" Or Left(GrdDataList.Text, 1) = "R") _
'                    And (GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¸Ñ´²" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¼o¤î" _
'                        Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "ºM¾P" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¦º¤`") Then
'                      For j = 0 To GrdDataList.Cols - 1
'                         GrdDataList.col = j
'                         GrdDataList.CellBackColor = &H0 '¶Â¦â
'                         GrdDataList.CellForeColor = &HFF00FF '¯»¬õ¦â
'                      Next j
'                  '§PÂ_¬O§_¬°¹ï³y,Åã¥Ü¤£¦PÃC¦â
'                  ElseIf Right(GrdDataList.Text, 1) = "¡ò" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("ª¬ºA")) = "¹ï³y" Then
'                     For j = 0 To GrdDataList.Cols - 1
'                        GrdDataList.col = j
'                        GrdDataList.CellBackColor = &H8080FF
'                     Next j
'                  Else
'                     For j = 0 To GrdDataList.Cols - 1
'                        If j <> 1 Then
'                            GrdDataList.col = j
'                            GrdDataList.CellBackColor = QBColor(15)
'                        End If
'                     Next j
'                  End If
'                  Me.Enabled = True
'                  Exit Sub
'               End If
'            Next i
'            Me.Enabled = True
'      'Add By Sindy 2019/10/8
'      Case 12 '±Hµo«H¨ç-©¹¨Ó°O¿ý
'         Me.Enabled = False
'         For i = 1 To GrdDataList.Rows - 1
'           GrdDataList.col = 0
'           GrdDataList.row = i
'           If Trim(GrdDataList.Text) = "V" Then
'               Screen.MousePointer = vbHourglass
'               GrdDataList.Text = ""
'               GrdDataList.col = 1
'               strTmp = Trim(GrdDataList.TextMatrix(i, GetValue("½s¸¹")))
'               If Len(strTmp) = 9 Or (Len(strTmp) = 12 And InStr(strTmp, "-") > 0) Then
'                  Me.Hide
'                  Set frm880022.m_PrevF = Me
'                  frm880022.m_strNo = Left(strTmp, 9)
'                  frm880022.m_PCC02 = IIf(InStr(strTmp, "-") > 0, Right(strTmp, 2), "")
'                  If frm880022.QueryData = True Then
'                     frm880022.Show 'vbModal
'                  End If
'                  Screen.MousePointer = vbDefault
'                  Me.Enabled = True
'                  Exit Sub
'               End If
'           End If
'         Next i
'         Screen.MousePointer = vbDefault
'         Me.Enabled = True
'      '2019/10/8 END
'      Case Else
'   End Select
'   'end 2023/03/08
End Sub

'Add by Amy 2023/08/24 ¾ã²z
Public Sub PubShowNextData()
Dim blnPrintAdd As Boolean, ii As Integer, strTmp As String, strRepCon As String

If cmdState = 10 Then
   strRepCon = Text2
   If Option3(0).Value = True Then
      strRepCon = strRepCon & " (¦r­º¤ñ¹ï)"
   ElseIf Option3(1).Value = True Then
      strRepCon = strRepCon & " (¼Ò½k¤ñ¹ï)"
   End If
   cmdOK(cmdState).Enabled = False
End If
If cmdState <> 4 Then
   Call PubShowNextForm(cmdState, Me, grdDataList, strField, _
      IIf(Check3.Value = vbChecked, True, False), IIf(ChkPCT.Value = vbChecked, True, False), _
     Text3, "1", Text4, Text5, Text6, Text7, txtCountry(0), txtCountry(1), Text8, m_bolPrintRight, , strRepCon)
   If cmdState = 10 Then cmdOK(cmdState).Enabled = True
   Exit Sub
End If

Me.Enabled = False
Screen.MousePointer = vbHourglass
   Select Case cmdState
      Case 4 '¦a§}±ø
         blnPrintAdd = False
         'PUB_RestorePrinter Combo1 'Mark by Lydia 2024/03/13
         For ii = 1 To Me.grdDataList.Rows - 1
            If Me.grdDataList.TextMatrix(ii, GetValue("V")) = "V" Then
               strTmp = Pub_RplStr(Me.grdDataList.TextMatrix(ii, GetValue("½s¸¹")))
               If Left(strTmp, 1) = "X" Then
                  strExc(3) = "select pcc01,pcc02 from PotCustCont where pcc01='" & Left(strTmp, 8) & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(3))
                  If intI = 1 Then
                     If RsTemp.RecordCount > 1 Then
                        strExc(3) = "select pcc05 from customer,PotCustCont where cu01='" & Left(strTmp, 8) & "' and cu02='" & Mid(strTmp, 9, 1) & "' and cu01=pcc01(+) and cu127=pcc02(+)"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(3))
                        If intI = 1 Then
                           strExc(4) = "" & RsTemp.Fields(0)
                        End If
                        If MsgBox("¦¹«È¤á¦³¤@­Ó¥H¤W±µ¬¢¤H¡A¦¹¥\¯à¥u¦L¥X¹w³]±µ¬¢¤H" & strExc(4) & "¡A¬O§_½T©w¤´­n¦C¦L¡H" & vbCrLf & _
                           "­Y­n¦C¦L¡u¹w³]±µ¬¢¤H¡v, ½Ð¿ï¾Ü¡u¬O¡v", vbYesNo) = vbNo Then
                           Call cmdok_Click(9)
                           Exit Sub
                        End If
                     End If 'RecordCount > 1
                  End If
                  If PUB_AddAddressA4List(strTmp, strExc(0)) Then
                     blnPrintAdd = True
                  End If
                  '°ê¤º
                  If Val(strExc(0)) > 0 Then cmdOK(4).Caption = "°ê¤ºA4¦W±ø (" & Val(strExc(0)) & ")"
               End If '= "X"
            End If '= "V"
         Next ii
         If blnPrintAdd = False Then
            '¦a§}±ø=>A4¦W±ø
            MsgBox "½Ð¤Ä¿ï±ý¦C¦LA4¦W±øªº¸ê®Æ!!!", vbExclamation + vbOKOnly
         End If
      Case Else
   End Select
   cmdOK(8).BackColor = &H8000000F
   Screen.MousePointer = vbDefault
   Me.Enabled = True
End Sub

Private Sub cmdMemo_Click()
   cmdState = 99
   If fnSaveParentForm(Me) = False Then
      Me.Enabled = True
      Exit Sub
   End If
   Me.Enabled = False
   Set frm100137.UpForm = Me
   frm100137.Show
   Me.Enabled = True
End Sub

Private Sub cmdok_Click(Index As Integer)
   'Memo by Amy 2023/08/24 index=4 [°ê¤ºA4¦W±ø] ¶s¦WºÙ¦³­×§ïPubShowNextForm¤]­n§ï
   'add by nickc 2007/01/12
   If Len(Trim(Me.Text3.Text)) = 0 Then
       Me.Text3.Text = "ALL"
   End If
   cmdState = Index
   PubShowNextData
End Sub

'Modify by Amy 2022/07/29 ¦WºÙ¬d¸ß»yªk§ï¦Ü¦@¥ÎFunction,¨Ã¾ã²zµ{¦¡
'Modify by Amy 2022/11/14 ­ì:Private
Public Sub cmdSearch_Click()
    Dim s As Integer
    Dim strCheckWay As String, strNo As String, Str01 As String, strFields As String
    Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
    Dim IsDevelop As Boolean, IsContrast As Boolean
    Dim strWhere_Case As String 'Add by Amy 2023/01/16 for  ¬dXY½s¸¹®×¥ó
    Dim strRtnVal As String 'Add by Amy 2023/08/17
On Error GoTo ErrHnd

    bolPrint = False '¥ý³]©wµL¹ï³y
    StrToPrint = ""
    lngCounterI = 0
    '¥Ó½Ð¤H½s¸¹
    If Option2(0).Value = True Then
        If Len(Trim(Text1)) = 0 Then
            s = MsgBox("±ø¥ó¤£¥iªÅ¥Õ", , "¿é¤J±ø¥ó¿ù»~")
            Text1.SetFocus
            Exit Sub
        End If
    End If
    '¦WºÙ
    If Option2(1).Value = True Then
        If Len(Trim(Text2)) = 0 Then
            s = MsgBox("±ø¥ó¤£¥iªÅ¥Õ", , "¿é¤J±ø¥ó¿ù»~")
            Text2.SetFocus
            Exit Sub
        End If
    End If
    '­t³d¤H
    If Option2(2).Value = True Then
        If Len(Trim(Text9)) = 0 Then
            s = MsgBox("±ø¥ó¤£¥iªÅ¥Õ", , "¿é¤J±ø¥ó¿ù»~")
            Text9.SetFocus
            Exit Sub
        End If
    End If
    'Email
    If Option2(3).Value = True Then
        If Len(Trim(Text10)) = 0 Then
            s = MsgBox("±ø¥ó¤£¥iªÅ¥Õ", , "¿é¤J±ø¥ó¿ù»~")
            Text10.SetFocus
            Exit Sub
        End If
    End If
    'ID
    If Option2(4).Value = True Then
        If Len(Trim(Text11)) = 0 Then
            s = MsgBox("±ø¥ó¤£¥iªÅ¥Õ", , "¿é¤J±ø¥ó¿ù»~")
            Text11.SetFocus
            Exit Sub
        End If
    End If
    
   'Add by Amy 2023/08/17 ÄÝ©ó¬d¸ß¸m´«¦r¼u°T®§
   If Option2(1).Value = True Then
      If ChkQuryChangetxt(Text2, strRtnVal) = True Then
         frm100137_1.Caption = "°T®§"
         frm100137_1.txtOrg = Text2
         frm100137_1.txtChg = strRtnVal
         frm100137_1.Show vbModal
      End If
   End If
   
   ClearQueryLog (Me.Name) '²M°£¬d¸ß¦Lªí°O¿ýÀÉÄæ¦ì
   Screen.MousePointer = vbHourglass
   grdDataList.Clear
   grdDataList.Rows = 2
   SetDataListWidth
  
   'Modify by Amy 2022/08/19 +OrgN
   strFields = ",'' AS ÃöÁp½s¸¹,'' AS ÃöÁp¦WºÙ,'' AS ÃöÁpÃö«Y,'' AS ÃöÁp»¡©ú,'' AS OrgN "
   
    If Option2(0).Value = True Then
'*** ¥Ó½Ð¤H½s¸¹ ***
        '¼ç¦b«È¤á
        If UCase(Left(Trim(Text1), 1)) = "R" Then
            strSql = "Select' ' as V ,pcu01||pcu02||Decode(pcu02,'0','','¡¯') as ½s¸¹,Nvl(pcu08,Decode(pcu03,null,pcu07,RTrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06))) as ¦WºÙ,NA03 as °êÄy,pcu38 as ´¼Åv¤H­û,pcu39 as ª¬ºA,pcu40 as ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From PotCustomer,Nation,Staff Where pcu09=na01(+) And pcu01='" & Left(GetNewFagent(Trim(Text1)), 8) & "' And substr(LTrim(pcu38),1,5)=st01(+)"
            strSql = strSql & " Union All " & _
                        "Select ' ' as V ,poc01||poc02||Decode(poc02,'0','','¡¯') as ½s¸¹,Nvl(poc03,Decode(poc23,null,poc27,RTrim(poc23||' '||poc24||' '||poc25||' '||poc26))) as ¦WºÙ,NA03 as °êÄy,poc13 as ´¼Åv¤H­û,poc14 as ª¬ºA,poc15 as ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From PotCustomer1,Nation,Staff Where poc04=na01(+) And poc01='" & Left(GetNewFagent(Trim(Text1)), 8) & "' And poc13=st01(+)"
        Else
            strSql = "Select ' ' as V ,cu01||cu02||Decode(cu02,'0','','¡¯')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','¡´','') as ½s¸¹,Nvl(cu04,Decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)) as ¦WºÙ,NA03 as °êÄy,ST02 as ´¼Åv¤H­û,Decode(cu142,null,cu80,GetDizhang(cu142,'Y')) as ª¬ºA,cu79 as ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From Customer,Nation,Staff Where CU10=na01(+) And cu01='" & Left(GetNewFagent(Trim(Text1)), 8) & "' And cu13=st01(+)"
            strSql = strSql & " Union All " & _
                        "Select ' ' as V,fa01||fa02||Decode(fa02,'0','','¡¯')||Decode(fa77,'Y','$','') as ½s¸¹,Nvl(fa04,Decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) as ¦WºÙ,NA03 as °êÄy,' ' as ´¼Åv¤H­û,Decode(fa103,null,fa69,GetDizhang(fa103,'Y')) as ª¬ºA, fa29 as ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From Fagent,Nation Where fa10=na01(+) And fa01='" & Left(GetNewFagent(Trim(Text1)), 8) & "' "
            strSql = strSql & " Union All " & _
                        "Select ' ' as V,nt01||Decode(nt21,null,'¡ò','') as ½s¸¹,Nvl(nt02,Decode(nt03,null,nt07,nt03||' '||nt04||' '||nt05||' '||nt06)) as ¦WºÙ,NA03 as °êÄy,ST02 as ´¼Åv¤H­û,Decode(nt21,null,'¤£±o¥N²z','') as ª¬ºA, Decode(nt21,null,'','ºM¾P¤é´Á¡G'||sqldatet(nt21)||'¡F')||nt20 as ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From NotAgent,Nation,Staff Where nt08=na01(+) And nt01='" & IIf(Len(Trim(Text1)) >= 3, Trim(Text1), Right("000" & Trim(Text1), 3)) & "' And nt18=st01(+)"
            'Add by Amy 2023/12/11 +­·ÀIÀË¬d¹ï¶H
            strSql = strSql & " Union All " & GetSearchRiskChkSql(1, Me.Name, Text1)
        End If
        pub_QL05 = pub_QL05 & ";" & Option2(0).Caption & Trim(Text1)
    ElseIf Option2(1).Value = True Then
'*** ¥Ó½Ð¤H¦WºÙ ***
        '¼Ò½k¤ñ¹ï
        If Option3(0).Value = False Then
            strCheckWay = ">0"
            pub_QL05 = pub_QL05 & ";" & Option3(0).Caption
        '¦r­º¤ñ¹ï
        Else
            strCheckWay = "=1"
            pub_QL05 = pub_QL05 & ";" & Option3(1).Caption
        End If
        '¹ï³y
        strSQL1 = " And cp01 In(" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
        strSQL2 = " And cp01 In(" & SQLGrpStr("", 1) & ") "
        StrSQL3 = " And cp01 In(" & SQLGrpStr("", 3) & ") "
        StrSQL4 = " And cp01 In(" & SQLGrpStr("", 4) & ") "
        strSQL5 = " And cp01 In(" & SQLGrpStr("", 5) & ") "
        '§t§ë¸êªk°È¶}©Ý
        If Check1.Value = 1 Then IsDevelop = True
        '§R°£¹ï³y¼È¦sÀÉ¸ê®Æ
        cnnConnection.Execute "Delete From R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "' "
        '§t¹ï³y
        If Check2.Value = 1 Then IsContrast = True
        
        strSql = GetSearchNameSql(Me.Name, Text2, strCheckWay, IsDevelop, IsContrast, strSQL1, strSQL2, StrSQL3, StrSQL4, strSQL5)
        pub_QL05 = pub_QL05 & ";" & Option2(1).Caption & Trim(Text2)
    ElseIf Option2(2).Value = True Then
'*** ­t³d¤H (­t³d¤H»P±µ¬¢¤H¤£¥Î§ì¥N²z¤HÀÉ¡A¦]¬°¨S¦³)***
        'Modify by Amy 2023/01/07 ¨ú¥N§ï¦@¥Î¨ç¼Æ
        'Modify by Amy 2023/06/26 §ï§ìReplaceSign DB¨ç¼Æ
'        strTp(0) = Pub_ReplaceSign(True, "cu07")
'        strTp(1) = Pub_ReplaceSign(False, Text9)
        strTp(0) = "ReplaceSign(TO_MULTI_BYTE(Upper(cu07)))"
        'Modify by Amy 2023/09/21 §ïGetSearchNameSql»P¦P¼gªk,§_«h·|§ìªº«ÜºC
'        strTp(1) = Pub_GetField("Dual", "1=1", "ReplaceSign(TO_MULTI_BYTE(Upper('" & ChgSQL(Text9) & "')))")
        'strSql = "Select ' ' as V,cu01||cu02||Decode(cu02,'0','','¡¯')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','¡´','') as ½s¸¹,Nvl(cu04,Decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)) as ¦WºÙ,NA03 as °êÄy,ST02 as ´¼Åv¤H­û,Decode(cu142,null,cu80,GetDizhang(CU142,'Y')) as ª¬ºA,cu79 as ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & _
                " From Customer,Nation,Staff,(Select Distinct cu01 as A1 From Customer Where InStr(" & strTp(0) & ",'" & strTp(1) & "')>=1 ) A Where cu10=na01(+) And cu01=A.A1 And cu13=st01(+)"
        strTp(1) = ",(Select ReplaceSign(TO_MULTI_BYTE(Upper('" & ChgSQL(Text9) & "'))) kw From Dual) x "
        strSql = "Select ' ' as V,cu01||cu02||Decode(cu02,'0','','¡¯')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','¡´','') as ½s¸¹,Nvl(cu04,Decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)) as ¦WºÙ,NA03 as °êÄy,ST02 as ´¼Åv¤H­û,Decode(cu142,null,cu80,GetDizhang(CU142,'Y')) as ª¬ºA,cu79 as ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & _
                " From Customer,Nation,Staff,(Select Distinct cu01 as A1 From Customer" & strTp(1) & " Where InStr(cu07(+),kw)>=1 And CU01 is not null  ) A Where cu10=na01(+) And cu01=A.A1 And cu13=st01(+)"
        'end 2023/09/21
        pub_QL05 = pub_QL05 & ";" & Option2(2).Caption & Trim(Text9)
    ElseIf Option2(3).Value = True Then
'*** E-Mail ***
        'Modified by Lydia 2024/09/18 +°]°È°Æ¥»«H½cCU200
        strSql = "Select ' ' as V,cu01||cu02||Decode(cu02,'0','','¡¯')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','¡´','') as ½s¸¹,Nvl(cu04,Decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)) as ¦WºÙ,na03 as °êÄy,ST02 as ´¼Åv¤H­û,Decode(cu142,null,CU80,GetDizhang(cu142,'Y')) as ª¬ºA,cu79 as ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & _
                    " From Customer,Nation,Staff Where (Instr(NLS_Upper(cu20),'" & UCase(ChgSQL(Trim(Text10))) & "')>0 Or Instr(NLS_Upper(cu115),'" & UCase(ChgSQL(Trim(Text10))) & "')>0 Or Instr(NLS_Upper(cu116),'" & UCase(ChgSQL(Trim(Text10))) & "')>0  Or Instr(NLS_Upper(cu117),'" & UCase(ChgSQL(Trim(Text10))) & "')>0 Or Instr(NLS_Upper(cu118),'" & UCase(ChgSQL(Trim(Text10))) & "') > 0 Or Instr(NLS_Upper(CU200),'" & UCase(ChgSQL(Trim(Text10))) & "')>0 )  And cu10=na01(+) And cu13=st01(+)"
        strSql = strSql & " Union All " & _
                    "Select ' ' as V,pcu01||pcu02||Decode(pcu02,'0','','¡¯') as ½s¸¹,Nvl(pcu08,Decode(pcu03,null,pcu07,RTrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06))) as ¦WºÙ,na03 as °êÄy,pcu38 as ´¼Åv¤H­û,PCU39 as ª¬ºA,PCU40 as ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & _
                    " From PotCustomer,Nation,Staff Where (Instr(NLS_Upper(pcu18),'" & UCase(ChgSQL(Trim(Text10))) & "') >0 ) And pcu09=na01(+) And SubStr(LTrim(pcu38),1,5)=st01(+)"
        strSql = strSql & " Union All " & _
                    "Select ' ' as V,poc01||poc02||Decode(poc02,'0','','¡¯') as ½s¸¹,Nvl(poc03,Decode(poc23,null,poc27,RTrim(poc23||' '||poc24||' '||poc25||' '||poc26))) as ¦WºÙ,na03 as °êÄy,poc13 as ´¼Åv¤H­û,poc14 as ª¬ºA,poc15 as ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & _
                    " From PotCustomer1,Nation,Staff Where (Instr(NLS_Upper(poc09),'" & UCase(ChgSQL(Trim(Text10))) & "') >0 ) And poc04=na01(+) And poc13=st01(+)"
        'Modified by Lydia 2024/09/18 +°]°È°Æ¥»«H½cFA134
        strSql = strSql & " Union All " & _
                    "Select ' ' as V,fa01||fa02||Decode(fa02,'0','','¡¯')||Decode(fa77,'Y','$','') as ½s¸¹,Nvl(fa04,Decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) as ¦WºÙ,na03 as °êÄy,' ' as ´¼Åv¤H­û,Decode(fa103,null,FA69,GetDizhang(fa103,'Y')) as ª¬ºA, fa29 as ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & _
                    " From fagent,nation Where (Instr(NLS_Upper(fa16),'" & UCase(ChgSQL(Trim(Text10))) & "')> 0 or Instr(NLS_Upper(fa79),'" & UCase(ChgSQL(Trim(Text10))) & "')> 0 or Instr(NLS_Upper(fa105),'" & UCase(ChgSQL(Trim(Text10))) & "')> 0 or Instr(NLS_Upper(fa80),'" & UCase(ChgSQL(Trim(Text10))) & "')> 0 or Instr(NLS_Upper(fa81),'" & UCase(ChgSQL(Trim(Text10))) & "') > 0 Or Instr(NLS_Upper(fa82),'" & UCase(ChgSQL(Trim(Text10))) & "') > 0 or Instr(NLS_Upper(fa134),'" & UCase(ChgSQL(Trim(Text10))) & "')> 0 ) And fa10=na01(+) "
        strSql = strSql & " Union All " & _
                    "Select ' ' as V,pcc01||'0-'||pcc02 as ½s¸¹,Nvl(pcc05,Nvl(pcc03,pcc04)) as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H­û,' ' as ª¬ºA,PCC13 as ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & _
                    " From PotCustCont Where (Instr(NLS_Upper(pcc08),'" & UCase(ChgSQL(Trim(Text10))) & "') > 0 )  "
        '§t§ë¸êªk°È¶}©Ý
        If Check1.Value = 1 Then
            strSql = strSql & " Union All " & _
                    "Select ' ' as V,ecd02||'-'||LPAD(ecd01,6,'0') as ½s¸¹,ecd03||' '||ecd04 as ¦WºÙ,NA03 as °êÄy,' ' as ´¼Åv¤H­û,'§ëªk¶}©Ý'||Decode(ecd15,null,null,'-'||ecd15) as ª¬ºA,ecd16 as ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & _
                    " From ExPandCusDetail,ExPandCusattr,Nation Where (instr(NLS_Upper(ecd13),'" & UCase(ChgSQL(Trim(Text10))) & "') > 0 ) And ecd10=na01(+) And ecd02=eca01(+) "
        End If
        'Add By Sindy 2023/8/21 + ¹q¤l³ø¯S®í¦W³æ
        strSql = strSql & " Union All " & _
                    "Select ' ' as V,'¹q¤l³ø¯S®í¦W³æ-'||TBNP09 as ½s¸¹,TBNP01 as ¦WºÙ,'' as °êÄy,'' as ´¼Åv¤H­û,TBNP10 as ª¬ºA,'' as ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & _
                    " From TMBulletinNp Where (instr(NLS_Upper(TBNP01),'" & UCase(ChgSQL(Trim(Text10))) & "') > 0 ) And TBNP08='M' "
        '2023/8/21 END
        pub_QL05 = pub_QL05 & ";" & Option2(3).Caption & Trim(Text10)
    ElseIf Option2(4).Value = True Then
'*** ID ***
        strSql = "Select ' ' as V,cu01||cu02||Decode(cu02,'0','','¡¯')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','¡´','') as ½s¸¹,Nvl(cu04,Decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)) as ¦WºÙ,NA03 as °êÄy,ST02 as ´¼Åv¤H­û,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) as ª¬ºA,CU79 as ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & _
                    " From Customer,Nation,Staff,(Select Distinct cu01 as A1 From Customer Where InStr(cu11,'" & ChgSQL(Trim(Text11)) & "')>=1 ) A Where cu10=na01(+) And cu01=A.A1 And cu13=st01(+)"
        'Add by Amy 2023/12/11 +­·ÀIÀË¬d¹ï¶H
         strSql = strSql & " Union All " & GetSearchRiskChkSql(1, Me.Name, Text1)
        pub_QL05 = pub_QL05 & ";" & Option2(4).Caption & Trim(Text11)
    End If
    
    '¦WºÙ
    If Option2(1).Value = True Then
        'Modify by Amy 2022/08/19 ¦]¦WºÙ«e¥[§ä¨ì¤§¤¤ or ­^ or ¤éÄæ¦ì,¾É­P¦P½s¸¹µLªk±Æ©ó¤@°_ ­ì:Order by Upper(¦WºÙ),½s¸¹
        'ex: ¬d SONN & PARTNER 2µ§(Y45656000/1)¤Î§ëªk981-000001,2µ§Y½s¸¹µLªk±Æ¤@°_
        strSql = "Select X.*,Decode(Ocu01,null, '',NVL(Ocu03,0)) as OCU03 From (" & strSql & ") X, OldCustomer Where substr(½s¸¹,1,8)= ocu01(+) Order by Upper(OrgN) "
    Else
        strSql = "Select X.*,Decode(Ocu01,null, '',NVL(Ocu03,0)) as OCU03 From (" & strSql & ") X, OldCustomer Where substr(½s¸¹,1,8)= ocu01(+) Order by ½s¸¹ "
    End If
    '§t§ë¸êªk°È¶}©Ý
    If Check1.Value = 1 Then
        pub_QL05 = pub_QL05 & ";" & Check1.Caption
    End If
    '§t¹ï³y
    If Check2.Value = 1 Then
        pub_QL05 = pub_QL05 & ";" & Check2.Caption
    End If
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/13 °O¿ý¦¹Formªº¬d¸ß±ø¥ó
    If adoRecordset.RecordCount <> 0 Then
        InsertQueryLog (adoRecordset.RecordCount)
        If Not cmdOK(0).Enabled Then cmdOK(0).Enabled = True
        If Not cmdOK(1).Enabled Then cmdOK(1).Enabled = True
        If Not cmdOK(2).Enabled Then cmdOK(2).Enabled = True
        If Not cmdOK(5).Enabled Then cmdOK(5).Enabled = True
        If Not cmdOK(6).Enabled Then cmdOK(6).Enabled = True
        If Not cmdOK(7).Enabled Then cmdOK(7).Enabled = True
        Set grdDataList.Recordset = adoRecordset
    Else
        InsertQueryLog (0)
        Pub_Can_Copy_Pic = True
        ShowNoData
        Pub_Can_Copy_Pic = False
        cmdOK(0).Enabled = False
        cmdOK(1).Enabled = False
        cmdOK(2).Enabled = False
        cmdOK(5).Enabled = False
        cmdOK(6).Enabled = False
        cmdOK(7).Enabled = False
        grdDataList.Clear
    End If
    Me.grdDataList.Visible = False
    SetDataListWidth
    CheckOC
    
    'Modify by Amy 2023/03/08 Äæ¦ì§ï°ÊºA
    With Me.grdDataList
        If .Rows > 0 Then
            For i = 1 To .Rows - 1
                .row = i
                .col = 1
                .CellForeColor = &H0 '¦r¶Â¦â
                'Modify by Amy 2023/08/24 ÅÜ¦â§ï¬°¦@¥Î¨ç¼Æ(ÅÜ¦â³]©w¥H¦@¥Î¨ç¼Æ¬°¥D-»P¨q¬Â½T»{¹L)
'                'Add by Amy 2023/01/16 +X ©Î Y ½s¸¹­YµL®×¥óÅã¥Ü¡¿
'                If Check3.Value = vbChecked And (Left(.Text, 1) = «È¤á½s¸¹ Or Left(.Text, 1) = ¥N²z¤H½s¸¹) Then
'                    If ChkXYCase(Left(.Text, 9)) = False Then
'                        .Text = .Text & "¡¿"
'                    End If
'                End If
'                'end 2023/01/16
'                '¬¡¤Æ«È¤áÅã¥Ü¾ã¦C¶À,­Y¦³§b±b½s¸¹©³¬°¬õ¨ä¥LÄæ¬°¶À
'                If .TextMatrix(i, GetValue("«Ý¬¡¤Æ«È¤á")) = "0" And Right(.Text, 1) <> "¡¯" Then
'                    For j = 0 To .Cols - 1
'                        If Right(.Text, 1) = "$" And j = 1 Then
'                        Else
'                            .col = j
'                            .CellBackColor = vbYellow
'                        End If
'                    Next
'                '§b±b
'                ElseIf Right(.Text, 1) = "$" Then
'                    .CellBackColor = &HFF& '¬õ¦â
'                '«È¤áª¬ºA¬° ¾E²¾¤£©ú/¼o¤î/ºM¾P/¦º¤` Åã¥Ü¶Â©³¯»¦r
'                ElseIf (Left(.Text, 1) = "Y" Or Left(.Text, 1) = "X" Or Left(.Text, 1) = "R") _
'                  And (.TextMatrix(i, GetValue("ª¬ºA")) = "¸Ñ´²" Or .TextMatrix(i, GetValue("ª¬ºA")) = "¼o¤î" Or .TextMatrix(i, GetValue("ª¬ºA")) = "ºM¾P" Or .TextMatrix(i, GetValue("ª¬ºA")) = "¦º¤`") Then
'                    For j = 0 To .Cols - 1
'                        .col = j
'                        .CellBackColor = &H0 '¶Â¦â
'                        .CellForeColor = &HFF00FF '¯»¬õ¦â
'                    Next j
'                ElseIf Right(.Text, 1) = "¡ò" Or .TextMatrix(i, GetValue("ª¬ºA")) = "¹ï³y" Or .TextMatrix(i, GetValue("ª¬ºA")) = "¨ä¥L¬ÛÃö¤H" Then
                    'Modify by Amy 2023/09/26 ¨Ìª¬ºA§ó·s´¼Åv¤H­û§ï¬°¦@¥Î¨ç¼Æ
'                    '¹ï³y­«§ì´¼Åv¤H¸ê®Æ
                    If Me.grdDataList.TextMatrix(i, GetValue("ª¬ºA")) = "¹ï³y" Or .TextMatrix(i, GetValue("ª¬ºA")) = "¨ä¥L¬ÛÃö¤H" Then
                        bolPrint = True '¦³¹ï³y¸ê®Æ
'                        strNo = Pub_RplStr(.TextMatrix(i, GetValue("½s¸¹")))
'                        StrToPrint = strNo & ","
'                        Str01 = SystemNumber(strNo, 1)
'                        Select Case Str01
'                            Case "FCP", "FG"
'                                .TextMatrix(i, GetValue("´¼Åv¤H­û")) = GetPrjSalesNM(PUB_GetFCPSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                            Case "FCL", "LIN"
'                                .TextMatrix(i, GetValue("´¼Åv¤H­û")) = GetPrjSalesNM(PUB_GetFCLSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                            Case "FCT"
'                                .TextMatrix(i, GetValue("´¼Åv¤H­û")) = GetPrjSalesNM(PUB_GetFCTSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                            Case "S"
'                                If .TextMatrix(i, GetValue("¥Ó½Ð°ê®a")) = "000" Then
'                                    .TextMatrix(i, GetValue("´¼Åv¤H­û")) = GetPrjSalesNM(PUB_GetFCTSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                                Else
'                                    .TextMatrix(i, GetValue("´¼Åv¤H­û")) = GetPrjSalesNM(PUB_GetAKindSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                                End If
'                            Case Else
'                                .TextMatrix(i, GetValue("´¼Åv¤H­û")) = GetPrjSalesNM(PUB_GetAKindSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                        End Select
'                        .TextMatrix(i, GetValue("®×¥ó©Ê½è")) = .TextMatrix(i, GetValue("®×¥ó©Ê½è")) & PUB_GetRelateCasePropertyName(.TextMatrix(i, GetValue("Á`¦¬¤å¸¹")), "1")
'                        '§ó·s´¼Åv¤H­û¦Ü¼È¦sÀÉ
'                        strExc(0) = "Update R100102_1 Set R021003='" & .TextMatrix(i, GetValue("´¼Åv¤H­û")) & "' Where R021014='" & Str01 & "' And R021015='" & SystemNumber(strNo, 2) & "' And R021016='" & SystemNumber(strNo, 3) & "' And R021017='" & SystemNumber(strNo, 4) & "' "
'                        cnnConnection.Execute strExc(0)
                    End If
'                    '¤£±o¥N²z/¹ï³y
'                    If Right(.Text, 1) = "¡ò" Or .TextMatrix(i, GetValue("ª¬ºA")) = "¹ï³y" Then
'                        For j = 0 To .Cols - 1
'                            .col = j
'                            .CellBackColor = &H8080FF
'                        Next j
'                    End If
'                '°w¹ïCW03=7.´C¤¶¥­¥x,Åã¥Ü¾ï¦â
'                ElseIf Left(.TextMatrix(i, GetValue("½s¸¹")), 1) = "¥­" And .TextMatrix(i, GetValue("®×¥ó©Ê½è")) = "7" Then
'                    .CellBackColor = &H80FF& '¾ï¦â
'                End If
'                '°ê¤º¥~¼ç¦b«È¤á ´¼Åv¤H­ûÄæ»Ý­«§ì¸ê®Æ(¥i¯à¦hµ§)
'                If Left(.Text, 1) = "R" Then
'                    '.TextMatrix(i, GetValue("´¼Åv¤H­û")) = GetDevelopP(.TextMatrix(i, GetValue("´¼Åv¤H­û")))
'                End If
                Call UpdQuerySales(Me.Name, grdDataList, strField)
                'end 2023/09/26
                Call SetMSGridColorQCus(0, Me.Name, grdDataList, strField, IIf(Check3.Value = vbChecked, True, False))
                'end 2023/08/24
            Next i
        End If
    End With
   
    '­Y¥u¦³¤@µ§¸ê®Æ, «hª½±µ³]©w¬°ÂI¿ï¦¹µ§¸ê®Æ
    'Modify by Amy 2023/08/24 ­ìµ{¦¡¼g¦Ü¦@¥Î
    cmdOK(8).BackColor = &H8000000F
    Call SetGridOneData
    'end 2023/08/24
   'end 2023/03/08
   Me.grdDataList.Visible = True
   If bolPrint Then
        cmdOK(10).Enabled = True
   Else
        cmdOK(10).Enabled = False
   End If
   Screen.MousePointer = vbDefault
   Exit Sub

ErrHnd:
    If Err.Number = -2147217900 Then
        MsgBox "¿é¤Jªº¤å¦rµLªk¬d¸ß,½Ð¹q¸£¤¤¤ß¨ó§U¡I"
    Else
        MsgBox Err.Description
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdSearchOLD_Click()
'Dim StrSQLa As String
'Dim strCheckWay As String
''Add by Amy 2013/11/06
'Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
'Dim strSwhSQL1 As String, strSwhSQL2 As String
'Dim strSubSQL1 As String, strSubSQL2 As String
'Dim strNo As String, Str01 As String
'Dim strFields As String 'Added by Lydia 2017/02/14 ³]©wÃöÁp¥N¸¹Äæ¦ì
'
''Add by Amy 2015/03/27 +ErrHnd ³y¦r¡u‹Ü¡v·|¿ù,¥Ø«e³y¦rµLªk¤ñ¹ï(¥Ñ©ó³y¦r«D³Ì«á¤@­Ó¦r¤]¬d¤£¥X,¬Gµ{¦¡¤£§ï)
'On Error GoTo ErrHnd
'
'bolPrint = False '¥ý³]©wµL¹ï³y
'StrToPrint = ""
''end 2013/11/06
'
'   lngCounterI = 0
'   Dim s As Integer
'
'   If Option2(0).Value = True Then
'       If Len(Trim(Text1)) = 0 Then
'           s = MsgBox("±ø¥ó¤£¥iªÅ¥Õ", , "¿é¤J±ø¥ó¿ù»~")
'           Text1.SetFocus
'           Exit Sub
'       End If
'   End If
'   If Option2(1).Value = True Then
'       If Len(Trim(Text2)) = 0 Then
'           s = MsgBox("±ø¥ó¤£¥iªÅ¥Õ", , "¿é¤J±ø¥ó¿ù»~")
'           Text2.SetFocus
'           Exit Sub
'       End If
'   End If
'   'add by nickc 2007/10/24
'   If Option2(2).Value = True Then
'       If Len(Trim(Text9)) = 0 Then
'           s = MsgBox("±ø¥ó¤£¥iªÅ¥Õ", , "¿é¤J±ø¥ó¿ù»~")
'           Text9.SetFocus
'           Exit Sub
'       End If
'   End If
'
'   'add by Toni 2008/12/03
'   If Option2(3).Value = True Then
'       If Len(Trim(Text10)) = 0 Then
'           s = MsgBox("±ø¥ó¤£¥iªÅ¥Õ", , "¿é¤J±ø¥ó¿ù»~")
'           Text10.SetFocus
'           Exit Sub
'       End If
'   End If
'
'   'add by nickc 2008/05/02
'   If Option2(4).Value = True Then
'       If Len(Trim(Text11)) = 0 Then
'           s = MsgBox("±ø¥ó¤£¥iªÅ¥Õ", , "¿é¤J±ø¥ó¿ù»~")
'           Text11.SetFocus
'           Exit Sub
'       End If
'   End If
'
'   ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/22 ²M°£¬d¸ß¦Lªí°O¿ýÀÉÄæ¦ì
'   Screen.MousePointer = vbHourglass
'   GrdDataList.Clear
'   GrdDataList.Rows = 2
'   SetDataListWidth
'   StrSQLa = ""
'   strFields = ",'' AS ÃöÁp½s¸¹,'' AS ÃöÁp¦WºÙ,'' AS ÃöÁpÃö«Y,'' AS ÃöÁp»¡©ú " 'Added by Lydia 2017/02/14
'   '­Y¬°°ê¤º´¼Åv¤H­û©Î°ê¤º¤uµ{®v, ¤£¥i¬d¥N²z¤H¸ê®Æ
'   'Modify By Sindy 2011/01/04 ¨ú®ø
'   'If bolFNation = False Then
'   '    StrSQLa = " And FA01<'Y' "
'   'End If
'
'   'Modify by Amy 2013/10/30 Åª¨úFagent¤ÎCustomerªºª¬ºAÄæ®É¡A¥ýÀË¬dFA103©ÎCU142¡A¦³­ÈÅã¥Ü ³B²z±¡§Îªº¤º®e¡AµL­È¤~§ì­ìª¬ºAÄæ¦ì
'   'Modify by Amy 2013/09/27 +trim±¼ªÅ¥Õ¥hÀË¬d:½s¸¹,¦WºÙ,ID,­t³d¤H,E-Mail
'   'Modify by Morgan 2007/12/14 µ{¦¡ÅÞ¿è¾ã²z
'   '¥Ó½Ð¤H½s¸¹
'   If Option2(0).Value = True Then
'      'Modify by Amy 2013/11/06 +¥Ó½Ð°ê®a/Á`¦¬¤å¸¹/®×¥ó©Ê½è/¦¬¤å¤é
'      'Modify by Morgan 2007/12/13 ¥[¥i¬d¼ç¦b«È¤á
'      If UCase(Left(Trim(Text1), 1)) = "R" Then
'         'Modify by Amy 2020/03/16 ´¼Åv¤H­û ­ì:st02 ,¦]¶}µo¤H­û¥i¯à¦h¤H
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = "SELECT ' ' AS V ,PCU01||PCU02||Decode(PCU02,'0','','¡¯') AS ½s¸¹,NVL(PCU08,Decode(PCU03,null,PCU07,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06))) AS ¦WºÙ,NA03 AS °êÄy,pcu38 AS ´¼Åv¤H­û,PCU39 AS ª¬ºA,PCU40 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " FROM POTCUSTOMER,NATION,staff WHERE PCU09=NA01(+) AND PCU01='" & Left(GetNewFagent(Trim(Text1)), 8) & "' and substr(LTrim(PCU38),1,5)=ST01(+)"
'         'Add By Sindy 2011/10/11
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT ' ' AS V ,PoC01||PoC02||Decode(PoC02,'0','','¡¯') AS ½s¸¹,NVL(PoC03,Decode(PoC23,null,PoC27,RTRIM(PoC23||' '||PoC24||' '||PoC25||' '||PoC26))) AS ¦WºÙ,NA03 AS °êÄy,poc13 AS ´¼Åv¤H­û,PoC14 AS ª¬ºA,PoC15 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " FROM POTCUSTOMER1,NATION,staff WHERE PoC04=NA01(+) AND PoC01='" & Left(GetNewFagent(Trim(Text1)), 8) & "' and poc13=ST01(+)"
'         'end 2020/03/16
'      Else
'         'edit by nickc 2008/01/03 ¥[¤J¯S®í«È¤á
'         'strSQL = "SELECT ' ' AS V ,CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','') AS ½s¸¹,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,CU80 AS ª¬ºA,CU79 AS ³Æµù FROM CUSTOMER,NATION,STAFF WHERE CU10=NA01(+) AND CU01='" & Left(GetNewFagent(Text1), 8) & "' AND CU13=ST01(+)"
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = "SELECT ' ' AS V ,CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','¡´','') AS ½s¸¹,NVL(CU04,Decode(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS ª¬ºA,CU79 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " FROM CUSTOMER,NATION,STAFF WHERE CU10=NA01(+) AND CU01='" & Left(GetNewFagent(Trim(Text1)), 8) & "' AND CU13=ST01(+)"
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','¡¯')||Decode(fa77,'Y','$','') AS ½s¸¹,NVL(fa04,Decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) as ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS ª¬ºA, FA29 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " FROM fagent,nation where fa10=na01(+) and fa01='" & Left(GetNewFagent(Trim(Text1)), 8) & "' " & StrSQLa
'         'Add By Sindy 2012/3/21
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'¡ò','') AS ½s¸¹,NVL(NT02,Decode(NT03,null,NT07,NT03||' '||NT04||' '||NT05||' '||NT06)) as ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(NT21,null,'¤£±o¥N²z','') AS ª¬ºA, Decode(NT21,null,'','ºM¾P¤é´Á¡G'||sqldatet(NT21)||'¡F')||NT20 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " FROM notagent,nation,STAFF where nt08=na01(+) and nt01='" & IIf(Len(Trim(Text1)) >= 3, Trim(Text1), Right("000" & Trim(Text1), 3)) & "' AND nt18=ST01(+)"
'      End If
'      pub_QL05 = pub_QL05 & ";" & Option2(0).Caption & Trim(Text1) 'Add By Sindy 2010/10/22
'
'   '¥Ó½Ð¤H¦WºÙ
'   ElseIf Option2(1).Value = True Then
'      '¥H½s¸¹©Î¦WºÙ
'        '¼Ò½k¤ñ¹ï
'        If Option3(0).Value = False Then
'           strCheckWay = ">0"
'           pub_QL05 = pub_QL05 & ";" & Option3(0).Caption 'Add By Sindy 2010/10/22
'        '¦r­º¤ñ¹ï
'        Else
'           strCheckWay = "=1"
'           pub_QL05 = pub_QL05 & ";" & Option3(1).Caption 'Add By Sindy 2010/10/22
'        End If
'        'Add by Amy 2013/11/06
'        strTp(3) = ChgSQL(UCase(Trim(Text2)))
'        '¹ï³y
'        strSQL1 = " AND CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
'        strSQL2 = " AND CP01 IN (" & SQLGrpStr("", 1) & ") "
'        StrSQL3 = " AND CP01 IN (" & SQLGrpStr("", 3) & ") "
'        StrSQL4 = " AND CP01 IN (" & SQLGrpStr("", 4) & ") "
'        strSQL5 = " AND CP01 IN (" & SQLGrpStr("", 5) & ") "
'        'end 2013/11/06
'
''Modify by Amy 2013/11/19 ®³±¼¤¤­^¤é
''        '¤¤¤å
''        If Option1(0).Value = True Then
''            pub_QL05 = pub_QL05 & ";" & Option1(0).Caption 'Add By Sindy 2010/10/22
''            'Modify by Amy 2013/11/06 +¥Ó½Ð°ê®a/Á`¦¬¤å¸¹/®×¥ó©Ê½è/¦¬¤å¤é
''            'edit by nickc 2008/01/03 ¥[¤J¯S®í«È¤á
''            'strSQL = "SELECT ' ' as V,CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','') AS ½s¸¹,CU04 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,CU80 AS ª¬ºA,CU79 AS ³Æµù FROM CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(CU04,'" & ChgSQL(Text2) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+)"
''            strSql = "SELECT ' ' as V,CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','¡´','') AS ½s¸¹,CU04 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(CU142,null,CU80,Decode(CU142,'A','¦P·N©è±b¤¤',Decode(CU142,'B','«Å§i¯}²£','±b´Ú³B²z¤¤'))) AS ª¬ºA,CU79 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é FROM CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(CU04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+)"
''            'Add by Morgan 2007/12/13 ¥[¥i¬d¼ç¦b«È¤á
''            strSql = strSql & " union all SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','¡¯') AS ½s¸¹,PCU08 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,PCU39 AS ª¬ºA,PCU40 AS ³Æµù,' ' AS ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From potcustomer,nation,staff, (Select Distinct pcu01 As A1 From potcustomer Where instr(pcu08,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A where pcu09=na01(+) and pcu01=A.A1 and substr(LTrim(PCU38),1,5)=ST01(+)"
''            'end 2007/12/13
''            'Add By Sindy 98/03/19
''            strSql = strSql & " union all SELECT ' ' AS V,POC01||POC02||Decode(POC02,'0','','¡¯') AS ½s¸¹,POC03 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,POC14 AS ª¬ºA,POC15 AS ³Æµù,' ' AS ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From potcustomer1,nation,staff, (Select Distinct poc01 As A1 From potcustomer1 Where instr(poc03,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A where poc04=na01(+) and poc01=A.A1 and poc13=ST01(+)"
''            '98/03/19 End
''            strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','¡¯')||Decode(fa77,'Y','$','') AS ½s¸¹,fa04 as ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,Decode(FA103,null,FA69,Decode(FA103,'A','¦P·N©è±b¤¤',Decode(FA103,'B','«Å§i¯}²£','±b´Ú³B²z¤¤'))) AS ª¬ºA, FA29 AS ³Æµù,' ' AS ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From fagent,nation, (Select Distinct FA01 As A1 From Fagent Where instr(fa04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A where fa10=na01(+) AND FA01=A.A1 " & StrSQLa
''            'Add by Morgan 2007/12/19 ¥[¥i¬dÁpµ¸¤H
''            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC05 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(CU142,null,CU80,Decode(CU142,'A','¦P·N©è±b¤¤',Decode(CU142,'B','«Å§i¯}²£','±b´Ú³B²z¤¤'))) AS ª¬ºA,PCC13 AS ³Æµù,' ' AS ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,CUSTOMER,NATION,STAFF WHERE CU10=NA01(+) AND CU13=ST01(+) AND CU01(+)=PCC01 AND CU02='0' "
''            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC05 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,PCU39 AS ª¬ºA,PCC13 AS ³Æµù,' ' AS ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,potcustomer,nation,staff where pcu09=na01(+) AND PCU01(+)=PCC01 AND PCU02='0' and substr(LTrim(PCU38),1,5)=ST01(+) "
''            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC05 AS ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,Decode(FA103,null,FA69,Decode(FA103,'A','¦P·N©è±b¤¤',Decode(FA103,'B','«Å§i¯}²£','±b´Ú³B²z¤¤'))) AS ª¬ºA, PCC13 AS ³Æµù,' ' AS ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' " & StrSQLa
''            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC05 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,POC14 AS ª¬ºA,PCC13 AS ³Æµù,' ' AS ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,potcustomer1,nation,staff where poc04=na01(+) AND POC01(+)=PCC01 AND POC02='0' and poc13=ST01(+) "
''            'end 2007/12/19
''            'Add By Sindy 2012/3/21
''            strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'¡ò','') AS ½s¸¹,NT02 as ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(NT21,null,'¤£±o¥N²z','') AS ª¬ºA, Decode(NT21,null,'','ºM¾P¤é´Á¡G'||sqldatet(NT21)||'¡F')||NT20 AS ³Æµù,' ' AS ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From notagent,nation,STAFF, (Select Distinct nt01 As A1 From notagent Where instr(nt02,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A where nt08=na01(+) AND nt01=A.A1 AND nt18=ST01(+)"
''
''            'Add by Amy 2013/11/06 +¹ï³y
''            strSubSQL1 = " And InStr(CP40,'" & ChgSQL(Trim(Text2)) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(CP50,'" & ChgSQL(Trim(Text2)) & "') " & strCheckWay
''            strSwhSQL1 = " CP40>' ' "
''            strSwhSQL2 = " CP50>' ' "
''            '°Ó¼Ð
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,Decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(TM29,'Y','¡¯','')||Decode(length(nvl(tm57,'')),null,'','¡´') as ½s¸¹, CP40 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL1 & _
''                        " Union  Select ' ' as V,Decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(TM29,'Y','¡¯','')||Decode(length(nvl(tm57,'')),null,'','¡´') as ½s¸¹, CP50 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL2
''            '±M§Q
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,Decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,'Y','¡¯','')||Decode(length(nvl(pa108,'')),null,'','¡´') as ½s¸¹,CP40 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL1 & _
''                        " Union  Select ' ' as V,Decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,'Y','¡¯','')||Decode(length(nvl(pa108,'')),null,'','¡´') as ½s¸¹,CP50 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL2
''            'ªk°È
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(LC08,'Y','¡¯','')||Decode(length(nvl(LC34,'')),null,'','¡´') AS ½s¸¹,CP40 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,' ' AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) " & StrSQL3 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(LC08,'Y','¡¯','')||Decode(length(nvl(LC34,'')),null,'','¡´') AS ½s¸¹,CP50 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a ,CP09 as Á`¦¬¤å¸¹,' ' AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) " & StrSQL3 & strSubSQL2
''            'ÅU°Ý
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(HC09,'Y','¡¯','')||Decode(length(nvl(HC19,'')),null,'','¡´') as ½s¸¹,CP40 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,' ' AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) " & StrSQL4 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(HC09,'Y','¡¯','')||Decode(length(nvl(HC19,'')),null,'','¡´') as ½s¸¹,CP50 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,' ' AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) " & StrSQL4 & strSubSQL2
''            'ªA°È
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(SP15,'Y','¡¯','')||Decode(length(nvl(SP61,'')),null,'','¡´') as ½s¸¹,CP40 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,SP09 as¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,' ' AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & strSQL5 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(SP15,'Y','¡¯','')||Decode(length(nvl(SP61,'')),null,'','¡´') as ½s¸¹,CP50 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,SP09 as¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,' ' AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & strSQL5 & strSubSQL2
''            'end 2013/11/06
''
''        '­^¤å
''        ElseIf Option1(1).Value = True Then
''            pub_QL05 = pub_QL05 & ";" & Option1(1).Caption 'Add By Sindy 2010/10/22
''            'Modify by Amy 2013/11/06 +¥Ó½Ð°ê®a/Á`¦¬¤å¸¹/®×¥ó©Ê½è/¦¬¤å¤é
''            'edit by nickc 2008/01/03 ¥[¤J¯S®í«È¤á
''            'strSQL = "SELECT ' ' AS V,CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','') AS ½s¸¹,cu05||' '||cu88||' '||cu89||' '||cu90 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,CU80 AS ª¬ºA,CU79 AS ³Æµù FROM CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(upper(cu05||' '||cu88||' '||cu89||' '||cu90),'" & UCase(ChgSQL(Text2)) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+)"
''            strSql = "SELECT ' ' AS V,CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','¡´','') AS ½s¸¹,cu05||' '||cu88||' '||cu89||' '||cu90 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(CU142,null,CU80,Decode(CU142,'A','¦P·N©è±b¤¤',Decode(CU142,'B','«Å§i¯}²£','±b´Ú³B²z¤¤'))) AS ª¬ºA,CU79 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(upper(cu05||' '||cu88||' '||cu89||' '||cu90),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+)"
''            'Add by Morgan 2007/12/13 ¥[¥i¬d¼ç¦b«È¤á
''            strSql = strSql & " union all SELECT ' ' AS V ,PCU01||PCU02||Decode(PCU02,'0','','¡¯') AS ½s¸¹,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06) AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,PCU39 AS ª¬ºA,PCU40 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From potcustomer,nation,staff, (Select Distinct pcu01 As A1 From potcustomer Where instr(upper(pcu03||' '||pcu04||' '||pcu05||' '||pcu06),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & " ) A where pcu09=na01(+) and pcu01=A.A1 and substr(LTrim(PCU38),1,5)=ST01(+)"
''            'end 2007/12/13
''            'Add By Sindy 2010/02/12
''            strSql = strSql & " union all SELECT ' ' AS V ,POC01||POC02||Decode(POC02,'0','','¡¯') AS ½s¸¹,RTRIM(POC23||' '||POC24||' '||POC25||' '||POC26) AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,POC14 AS ª¬ºA,POC15 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From potcustomer1,nation,staff, (Select Distinct poc01 As A1 From potcustomer1 Where instr(upper(poc23||' '||poc24||' '||poc25||' '||poc26),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & " ) A where poc04=na01(+) and poc01=A.A1 and poc13=ST01(+)"
''            '2010/02/12 End
''            strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','¡¯')||Decode(fa77,'Y','$','') AS ½s¸¹,fa05||' '||fa63||' '||fa64||' '||fa65 as ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,Decode(FA103,null,FA69,Decode(FA103,'A','¦P·N©è±b¤¤',Decode(FA103,'B','«Å§i¯}²£','±b´Ú³B²z¤¤'))) AS ª¬ºA, FA29 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From fagent,nation, (Select Distinct FA01 As A1 From Fagent Where instr(upper(fa05||' '||fa63||' '||fa64||' '||fa65),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & " ) A where fa10=na01(+) AND FA01=A.A1 " & StrSQLa
''            'Add by Morgan 2007/12/19 ¥[¥i¬dÁpµ¸¤H
''            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC03 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(CU142,null,CU80,Decode(CU142,'A','¦P·N©è±b¤¤',Decode(CU142,'B','«Å§i¯}²£','±b´Ú³B²z¤¤'))) AS ª¬ºA,PCC13 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & ") A,CUSTOMER,NATION,STAFF WHERE CU10=NA01(+) AND CU13=ST01(+) AND CU01(+)=PCC01 AND CU02='0' "
''            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC03 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,PCU39 AS ª¬ºA,PCC13 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & ") A,potcustomer,nation,staff where pcu09=na01(+) AND PCU01(+)=PCC01 AND PCU02='0' and substr(LTrim(PCU38),1,5)=ST01(+) "
''            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC03 AS ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,Decode(FA103,null,FA69,Decode(FA103,'A','¦P·N©è±b¤¤',Decode(FA103,'B','«Å§i¯}²£','±b´Ú³B²z¤¤'))) AS ª¬ºA, PCC13 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' " & StrSQLa
''            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC03 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,POC14 AS ª¬ºA,PCC13 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & ") A,potcustomer1,nation,staff where poc04=na01(+) AND POC01(+)=PCC01 AND POC02='0' and poc13=ST01(+) "
''            'end 2007/12/19
''            'Add By Sindy 2012/3/21
''            strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'¡ò','') AS ½s¸¹,NT03||' '||NT04||' '||NT05||' '||NT06 as ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(NT21,null,'¤£±o¥N²z','') AS ª¬ºA, Decode(NT21,null,'','ºM¾P¤é´Á¡G'||sqldatet(NT21)||'¡F')||NT20 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From notagent,nation,STAFF, (Select Distinct nt01 As A1 From notagent Where instr(upper(nt03||' '||nt04||' '||nt05||' '||nt06),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & " ) A where nt08=na01(+) AND nt01=A.A1 AND nt18=ST01(+)"
''
''            'Add by Amy 2013/11/06 +¹ï³y
''            strSubSQL1 = " And InStr(Upper(CP41),'" & UCase(ChgSQL(Trim(Text2))) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP51),'" & UCase(ChgSQL(Trim(Text2))) & "') " & strCheckWay
''            strSwhSQL1 = " CP41>' ' "
''            strSwhSQL2 = " CP51>' ' "
''            '°Ó¼Ð
''            strSql = strSql & " Union " & _
''                         "Select ' ' as V,Decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(TM29,'Y','¡¯','')||Decode(length(nvl(tm57,'')),null,'','¡´') as ½s¸¹, CP41 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL1 & _
''                         " Union Select ' ' as V,Decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(TM29,'Y','¡¯','')||Decode(length(nvl(tm57,'')),null,'','¡´') as ½s¸¹, CP51 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL2
''            '±M§Q
''            strSql = strSql & " Union " & _
''                         "Select ' ' as V,Decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,'Y','¡¯','')||Decode(length(nvl(pa108,'')),null,'','¡´') as ½s¸¹,CP41 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL1 & _
''                         " Union Select ' ' as V,Decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,'Y','¡¯','')||Decode(length(nvl(pa108,'')),null,'','¡´') as ½s¸¹,CP51 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL2
''            'ªk°È
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(LC08,'Y','¡¯','')||Decode(length(nvl(LC34,'')),null,'','¡´') AS ½s¸¹,CP41 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) " & StrSQL3 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(LC08,'Y','¡¯','')||Decode(length(nvl(LC34,'')),null,'','¡´') AS ½s¸¹,CP51 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) " & StrSQL3 & strSubSQL2
''            'ÅU°Ý
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(HC09,'Y','¡¯','')||Decode(length(nvl(HC19,'')),null,'','¡´') as ½s¸¹,CP41 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) " & StrSQL4 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(HC09,'Y','¡¯','')||Decode(length(nvl(HC19,'')),null,'','¡´') as ½s¸¹,CP51 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) " & StrSQL4 & strSubSQL2
''            'ªA°È
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(SP15,'Y','¡¯','')||Decode(length(nvl(SP61,'')),null,'','¡´') as ½s¸¹,CP41 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,SP09 as¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & strSQL5 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(SP15,'Y','¡¯','')||Decode(length(nvl(SP61,'')),null,'','¡´') as ½s¸¹,CP51 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,SP09 as¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & strSQL5 & strSubSQL2
''            'end 2013/11/06
''
''        '¤é¤å
''        ElseIf Option1(2).Value = True Then
''            pub_QL05 = pub_QL05 & ";" & Option1(2).Caption 'Add By Sindy 2010/10/22
''            'Modify by Amy 2013/11/06 +¥Ó½Ð°ê®a/Á`¦¬¤å¸¹/®×¥ó©Ê½è/¦¬¤å¤é
''            'edit by nickc 2008/01/03 ¥[¤J¯S®í«È¤á
''            'strSQL = "SELECT ' ' AS V,CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','') AS ½s¸¹,CU06 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,CU80 AS ª¬ºA,CU79 AS ³Æµù FROM CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(CU06,'" & ChgSQL(Text2) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+)"
''            strSql = "SELECT ' ' AS V,CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','¡´','') AS ½s¸¹,CU06 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(CU142,null,CU80,Decode(CU142,'A','¦P·N©è±b¤¤',Decode(CU142,'B','«Å§i¯}²£','±b´Ú³B²z¤¤'))) AS ª¬ºA,CU79 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(CU06,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+)"
''            'Add by Morgan 2007/12/13 ¥[¥i¬d¼ç¦b«È¤á
''            strSql = strSql & " union all SELECT ' ' AS V ,PCU01||PCU02||Decode(PCU02,'0','','¡¯') AS ½s¸¹,PCU07 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,PCU39 AS ª¬ºA,PCU40 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From potcustomer,nation,staff, (Select Distinct pcu01 As A1 From potcustomer Where instr(pCU07,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A where pcu09=na01(+) and pcu01=A.A1 and substr(LTrim(PCU38),1,5)=ST01(+)"
''            'end 2007/12/13
''            'Add By Sindy 2010/02/12
''            strSql = strSql & " union all SELECT ' ' AS V ,POC01||POC02||Decode(POC02,'0','','¡¯') AS ½s¸¹,POC27 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,POC14 AS ª¬ºA,POC15 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From potcustomer1,nation,staff, (Select Distinct poc01 As A1 From potcustomer1 Where instr(POC27,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A where poc04=na01(+) and poc01=A.A1 and poc13=ST01(+)"
''            '2010/02/12 End
''            strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','¡¯')||Decode(fa77,'Y','$','') AS ½s¸¹,fa06 as ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,Decode(FA103,null,FA69,Decode(FA103,'A','¦P·N©è±b¤¤',Decode(FA103,'B','«Å§i¯}²£','±b´Ú³B²z¤¤'))) AS ª¬ºA, FA29 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From fagent,nation, (Select Distinct FA01 As A1 From Fagent Where instr(fa06,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A where fa10=na01(+) AND FA01=A.A1 " & StrSQLa
''            'Add by Morgan 2007/12/19 ¥[¥i¬dÁpµ¸¤H
''            strSql = strSql & " union all select ' ' as V,CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','') AS ½s¸¹,CU06 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(CU142,null,CU80,Decode(CU142,'A','¦P·N©è±b¤¤',Decode(CU142,'B','«Å§i¯}²£','±b´Ú³B²z¤¤'))) AS ª¬ºA,CU79 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From CUSTOMER,NATION,STAFF, (Select Distinct pcc01 As A1 From potcustcont Where instr(pcc04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+)"
''            strSql = strSql & " union all SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','¡¯') AS ½s¸¹,PCU07 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,PCU39 AS ª¬ºA,PCU40 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From potcustomer,nation,staff, (Select Distinct pcc01 As A1 From potcustcont Where instr(pcc04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A where pcu09=na01(+) and pcu01=A.A1 and substr(LTrim(PCU38),1,5)=ST01(+)"
''            strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','¡¯')||Decode(fa77,'Y','$','') AS ½s¸¹,fa06 as ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,Decode(FA103,null,FA69,Decode(FA103,'A','¦P·N©è±b¤¤',Decode(FA103,'B','«Å§i¯}²£','±b´Ú³B²z¤¤'))) AS ª¬ºA, FA29 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From fagent,nation, (Select Distinct pcc01 As A1 From potcustcont Where instr(pcc04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A where fa10=na01(+) AND FA01=A.A1 " & StrSQLa
''            strSql = strSql & " union all SELECT ' ' AS V,POC01||POC02||Decode(POC02,'0','','¡¯') AS ½s¸¹,POC27 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,POC14 AS ª¬ºA,POC15 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From potcustomer1,nation,staff, (Select Distinct pcc01 As A1 From potcustcont Where instr(pcc04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A where poc04=na01(+) and poc01=A.A1 and poc13=ST01(+)"
''            'end 2007/12/19
''            'Add by Morgan 2007/12/19 ¥[¥i¬dÁpµ¸¤H
''            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC04 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(CU142,null,CU80,Decode(CU142,'A','¦P·N©è±b¤¤',Decode(CU142,'B','«Å§i¯}²£','±b´Ú³B²z¤¤'))) AS ª¬ºA,PCC13 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,CUSTOMER,NATION,STAFF WHERE CU10=NA01(+) AND CU13=ST01(+) AND CU01(+)=PCC01 AND CU02='0' "
''            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC04 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,PCU39 AS ª¬ºA,PCC13 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,potcustomer,nation,staff where pcu09=na01(+) AND PCU01(+)=PCC01 AND PCU02='0' and substr(LTrim(PCU38),1,5)=ST01(+) "
''            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC04 AS ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,Decode(FA103,null,FA69,Decode(FA103,'A','¦P·N©è±b¤¤',Decode(FA103,'B','«Å§i¯}²£','±b´Ú³B²z¤¤'))) AS ª¬ºA, PCC13 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' " & StrSQLa
''            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC04 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,POC14 AS ª¬ºA,PCC13 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,potcustomer1,nation,staff where poc04=na01(+) AND POC01(+)=PCC01 AND POC02='0' and poc13=ST01(+) "
''            'end 2007/12/19
''            'Add By Sindy 2012/3/21
''            strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'¡ò','') AS ½s¸¹,NT07 as ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(NT21,null,'¤£±o¥N²z','') AS ª¬ºA, Decode(NT21,null,'','ºM¾P¤é´Á¡G'||sqldatet(NT21)||'¡F')||NT20 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é From notagent,nation,STAFF, (Select Distinct nt01 As A1 From notagent Where instr(nt07,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A where nt08=na01(+) AND nt01=A.A1 AND nt18=ST01(+)"
''
''            'Add by Amy 2013/11/06 +¹ï³y
''            strSubSQL1 = " And InStr(CP42,'" & ChgSQL(Trim(Text2)) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(CP52,'" & ChgSQL(Trim(Text2)) & "') " & strCheckWay
''            strSwhSQL1 = " CP42>' ' "
''            strSwhSQL2 = " CP52>' ' "
''            '°Ó¼Ð
''            strSql = strSql & " Union " & _
''                         "Select ' ' as V,Decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(TM29,'Y','¡¯','')||Decode(length(nvl(tm57,'')),null,'','¡´') as ½s¸¹, CP42 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL1 & _
''                         " Union Select ' ' as V,Decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(TM29,'Y','¡¯','')||Decode(length(nvl(tm57,'')),null,'','¡´') as ½s¸¹, CP52 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL2
''            '±M§Q
''            strSql = strSql & " Union " & _
''                         "Select ' ' as V,Decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,'Y','¡¯','')||Decode(length(nvl(pa108,'')),null,'','¡´') as ½s¸¹,CP42 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL1 & _
''                         " Union Select ' ' as V,Decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,'Y','¡¯','')||Decode(length(nvl(pa108,'')),null,'','¡´') as ½s¸¹,CP52 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL2
''            'ªk°È
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(LC08,'Y','¡¯','')||Decode(length(nvl(LC34,'')),null,'','¡´') AS ½s¸¹,CP42 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) " & StrSQL3 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(LC08,'Y','¡¯','')||Decode(length(nvl(LC34,'')),null,'','¡´') AS ½s¸¹,CP52 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) " & StrSQL3 & strSubSQL2
''            'ÅU°Ý
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(HC09,'Y','¡¯','')||Decode(length(nvl(HC19,'')),null,'','¡´') as ½s¸¹,CP42 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) " & StrSQL4 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(HC09,'Y','¡¯','')||Decode(length(nvl(HC19,'')),null,'','¡´') as ½s¸¹,CP52 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) " & StrSQL4 & strSubSQL2
''            'ªA°È
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(SP15,'Y','¡¯','')||Decode(length(nvl(SP61,'')),null,'','¡´') as ½s¸¹,CP42 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,SP09 as¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & strSQL5 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(SP15,'Y','¡¯','')||Decode(length(nvl(SP61,'')),null,'','¡´') as ½s¸¹,CP52 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,SP09 as¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & strSQL5 & strSubSQL2
''
''            'end 2013/11/06
''        End If
'
'    'Modify by Amy 2015/03/27 ®³±¼¹ï³y®×¥ó½s¸¹²Å¸¹,+«È¤áºÝ¥­¥x±b¸¹¸ê®Æ
'    'Modified by Lydia 2019/12/26
'    'cnnConnection.Execute "Delete From R100102_1 Where ID='" & strUserNum & "' "
'    cnnConnection.Execute "Delete From R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "' "
'
'    If Check2.Value = 1 Then '§t¹ï³y
'           'Modify by Amy 2014/02/21 ¹ï³y¥Ñ¤U·h¤W¨Ó§ï»yªk¦s¦Ü¼È¦sÀÉ
''Modified by Lydia 2019/12/26 §ï¦¨¤½¥Î¼Ò²ÕPub_ProcR100102_1
''            '¹ï³y(¤¤)
''            strSubSQL1 = " And InStr(Upper(CP40),'" & strTp(3) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP50),'" & strTp(3) & "') " & strCheckWay
''            strSwhSQL1 = " CP40>' ' "
''            strSwhSQL2 = " CP50>' ' "
''            '°Ó¼Ð
''            '§ï¦¨¼Ò²Õ
''            strSql = "Insert Into R100102_1 (r021001,r021002,r021003,r021004,r021005,r021006,r021007,r021008,r021009,r021010,r021011,r021012,r021013,r021014,r021015,r021016,r021017,r021018,ID) " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹, CP40 as ¦WºÙ,' ' as ´¼Åv¤H,'1' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
''                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
''                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+) " & strSQL1 & strSubSQL1
''
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹, CP50 as ¦WºÙ,' ' as ´¼Åv¤H,'2' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
''                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
''                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+) " & strSQL1 & strSubSQL2
''            '±M§Q
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP40 as ¦WºÙ,' ' as ´¼Åv¤H,'1' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é, " & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
''                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
''                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+) " & strSQL2 & strSubSQL1
''
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP50 as ¦WºÙ,' ' as ´¼Åv¤H,'2' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é, " & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
''                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
''                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+) " & strSQL2 & strSubSQL2
''            'ªk°È
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP40 as ¦WºÙ,' ' as ´¼Åv¤H,'1' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é, " & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
''                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
''                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+)" & StrSQL3 & strSubSQL1
''
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP50 as ¦WºÙ,' ' as ´¼Åv¤H,'2' as ª¬ºA,' ' as ¥Ó½Ð°ê®a ,CP09 as Á`¦¬¤å¸¹,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é, " & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
''                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
''                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+)" & StrSQL3 & strSubSQL2
''            'ÅU°Ý
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP40 as ¦WºÙ,' ' as ´¼Åv¤H,'1' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) AS ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
''                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
''                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+)" & StrSQL4 & strSubSQL1
''
''            strSql = strSql & " Union " & _
''                        " Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP50 as ¦WºÙ,' ' as ´¼Åv¤H,'2' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) AS ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
''                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
''                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+)" & StrSQL4 & strSubSQL2
''            'ªA°È
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP40 as ¦WºÙ,' ' as ´¼Åv¤H,'1' as ª¬ºA,SP09 as¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
''                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
''                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+)" & strSQL5 & strSubSQL1
''
''           strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP50 as ¦WºÙ,' ' as ´¼Åv¤H,'2' as ª¬ºA,SP09 as¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
''                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
''                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+)" & strSQL5 & strSubSQL2
''
''            '¹ï³y(­^)
''            strSubSQL1 = " And InStr(Upper(CP41),'" & strTp(3) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP51),'" & strTp(3) & "') " & strCheckWay
''            strSwhSQL1 = " CP41>' ' "
''            strSwhSQL2 = " CP51>' ' "
''            '°Ó¼Ð
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹, CP41 as ¦WºÙ,' ' as ´¼Åv¤H,'1' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                           "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                           "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
''                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
''                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+)" & strSQL1 & strSubSQL1
''
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹, CP51 as ¦WºÙ,' ' as ´¼Åv¤H,'2' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                           "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                           "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
''                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
''                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+)" & strSQL1 & strSubSQL2
''            '±M§Q
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP41 as ¦WºÙ,' ' as ´¼Åv¤H,'1' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                           "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                           "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
''                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
''                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+)" & strSQL2 & strSubSQL1
''
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP51 as ¦WºÙ,' ' as ´¼Åv¤H,'2' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
''                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
''                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+)" & strSQL2 & strSubSQL2
''            'ªk°È
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP41 as ¦WºÙ,' ' as ´¼Åv¤H,'1' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) as ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
''                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
''                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+)" & StrSQL3 & strSubSQL1
''
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP51 as ¦WºÙ,' ' as ´¼Åv¤H,'2' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) as ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
''                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
''                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+)" & StrSQL3 & strSubSQL2
''            'ÅU°Ý
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP41 as ¦WºÙ,' ' as ´¼Åv¤H,'1' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) as ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
''                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
''                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+)" & StrSQL4 & strSubSQL1
''
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP51 as ¦WºÙ,' ' as ´¼Åv¤H,'2' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) as ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
''                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
''                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+)" & StrSQL4 & strSubSQL2
''            'ªA°È
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP41 as ¦WºÙ,' ' as ´¼Åv¤H,'1' as ª¬ºA,SP09 as¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) as ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
''                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
''                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+)" & strSQL5 & strSubSQL1
''
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP51 as ¦WºÙ,' ' as ´¼Åv¤H,'2' as ª¬ºA,SP09 as¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) as ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS ¥Ó½Ð¤H1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS ¥Ó½Ð¤H2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS ¥Ó½Ð¤H4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
''                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
''                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+)" & strSQL5 & strSubSQL2
''
''            '¹ï³y(¤é)
''            strSubSQL1 = " And InStr(Upper(CP42),'" & strTp(3) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP52),'" & strTp(3) & "') " & strCheckWay
''            strSwhSQL1 = " CP42>' ' "
''            strSwhSQL2 = " CP52>' ' "
''            '°Ó¼Ð
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹, CP42 as ¦WºÙ,' ' as ´¼Åv¤H,'1' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS ¥Ó½Ð¤H1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS ¥Ó½Ð¤H2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS ¥Ó½Ð¤H4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
''                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
''                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+)" & strSQL1 & strSubSQL1
''
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹, CP52 as ¦WºÙ,' ' as ´¼Åv¤H,'2' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS ¥Ó½Ð¤H1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS ¥Ó½Ð¤H2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS ¥Ó½Ð¤H4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
''                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
''                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+)" & strSQL1 & strSubSQL2
''            '±M§Q
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP42 as ¦WºÙ,' ' as ´¼Åv¤H,'1' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS ¥Ó½Ð¤H1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS ¥Ó½Ð¤H2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS ¥Ó½Ð¤H4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
''                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
''                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+)" & strSQL2 & strSubSQL1
''
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP52 as ¦WºÙ,' ' as ´¼Åv¤H,'2' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS ¥Ó½Ð¤H1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS ¥Ó½Ð¤H2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS ¥Ó½Ð¤H4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
''                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
''                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+)" & strSQL2 & strSubSQL2
''            'ªk°È
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP42 as ¦WºÙ,' ' as ´¼Åv¤H,'1' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) as ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS ¥Ó½Ð¤H1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS ¥Ó½Ð¤H2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS ¥Ó½Ð¤H4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
''                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
''                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+)" & StrSQL3 & strSubSQL1
''
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP52 as ¦WºÙ,' ' as ´¼Åv¤H,'2' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) as ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS ¥Ó½Ð¤H1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS ¥Ó½Ð¤H2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS ¥Ó½Ð¤H4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
''                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
''                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+)" & StrSQL3 & strSubSQL2
''            'ÅU°Ý
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP42 as ¦WºÙ,' ' as ´¼Åv¤H,'1' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) as ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS ¥Ó½Ð¤H1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS ¥Ó½Ð¤H2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS ¥Ó½Ð¤H4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
''                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
''                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+)" & StrSQL4 & strSubSQL1
''
''            strSql = strSql & " Union " & _
''                        " Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP52 as ¦WºÙ,' ' as ´¼Åv¤H,'2' as ª¬ºA,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) as ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS ¥Ó½Ð¤H1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS ¥Ó½Ð¤H2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS ¥Ó½Ð¤H4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
''                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
''                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+)" & StrSQL4 & strSubSQL2
''            'ªA°È
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP42 as ¦WºÙ,' ' as ´¼Åv¤H,'1' as ª¬ºA,SP09 as¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) as ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS ¥Ó½Ð¤H1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS ¥Ó½Ð¤H2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS ¥Ó½Ð¤H4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
''                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
''                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+)" & strSQL5 & strSubSQL1
''
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as ½s¸¹,CP52 as ¦WºÙ,' ' as ´¼Åv¤H,'2' as ª¬ºA,SP09 as¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) as ®×¥ó©Ê½è,''||cp05 as ¦¬¤å¤é," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS ¥Ó½Ð¤H1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS ¥Ó½Ð¤H2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS ¥Ó½Ð¤H3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS ¥Ó½Ð¤H4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS ¥Ó½Ð¤H5,CP01,CP02,CP03,CP04,CP10 AS ®×¥ó©Ê½è½s¸¹,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
''                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
''                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+)" & strSQL5 & strSubSQL2
''
''           cnnConnection.Execute strSql
''
''           '§R°£¹ï³y»P¥Ó½Ð¤H¬Û¦P¸ê®Æ
''           strSql = "Delete From R100102_1 Where ID='" & strUserNum & "' And (ltrim(rtrim(R021002))=ltrim(rtrim(R021008)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021009)) " & _
''                       "Or ltrim(rtrim(R021002))=ltrim(rtrim(R021010)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021011)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021012))) "
''           cnnConnection.Execute strSql
''           'end 2014/02/21
''
''           'Add by Amy 2014/03/17 ±N©Ò¦³°Ó¼Ð®×InStr(R021014,'T')¥B®×¥ó©Ê½è¬°1202(®Ö»é«e¥ý¦æ³qª¾)ªÌª¬ºA§ï¬° ¨ä¥L¬ÛÃö¤H
''           'Modify by Amy 2015/12/03 ¼W¥[°Ó¼Ð®×(CFC/S) ®×¥ó©Ê½è202(¥Ó½Ð·N¨£®Ñ)¤Î303(©µ´Á)ªÌ ª¬ºA§ï¬° ¨ä¥L¬ÛÃö¤H
''           strSql = "Update R100102_1 Set R021004='2' Where (InStr(R021014,'T')>0 or R021014='CFC' or R021014='S') And (R021018='1202' or R021018='202' or R021018='303')"
''           cnnConnection.Execute strSql
''           'end 2014/03/17
''           'Add by Amy 2015/12/03 ©Ò¦³±M§Q®×¥ó©Ê½è404(©µ´Á) ªÌª¬ºA§ï¬° ¨ä¥L¬ÛÃö¤H
''           strSql = "Update R100102_1 Set R021004='2' Where (InStr(R021014,'P')>0 or R021014='FG') And R021018='404' "
''           cnnConnection.Execute strSql
''           'end 2015/12/03
'           Call Pub_ProcR100102_1(strUserNum & "@" & Me.Name, strSQL1, strSQL2, StrSQL3, StrSQL4, strSQL5, strTp(3), strCheckWay)
''end 2019/12/26
'      End If
'            '¬dcustomer «È¤á ÀÉ
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = "SELECT ' ' as V,CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','¡´','') AS ½s¸¹,CU04 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS ª¬ºA,CU79 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " FROM CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(CU04,'" & strTp(3) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+)"
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all SELECT ' ' AS V,CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','¡´','') AS ½s¸¹,cu05||' '||cu88||' '||cu89||' '||cu90 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS ª¬ºA,CU79 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(upper(cu05||' '||cu88||' '||cu89||' '||cu90),'" & strTp(3) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+)"
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all SELECT ' ' AS V,CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','¡´','') AS ½s¸¹,CU06 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS ª¬ºA,CU79 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(CU06,'" & strTp(3) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+)"
'
'            '¬dFagent ¥N²z¤H ÀÉ
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','¡¯')||Decode(fa77,'Y','$','') AS ½s¸¹,fa04 as ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS ª¬ºA, FA29 AS ³Æµù,' ' AS ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From fagent,nation, (Select Distinct FA01 As A1 From Fagent Where instr(fa04,'" & strTp(3) & "')" & strCheckWay & " ) A where fa10=na01(+) AND FA01=A.A1 " & StrSQLa
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','¡¯')||Decode(fa77,'Y','$','') AS ½s¸¹,fa05||' '||fa63||' '||fa64||' '||fa65 as ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS ª¬ºA, FA29 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From fagent,nation, (Select Distinct FA01 As A1 From Fagent Where instr(upper(fa05||' '||fa63||' '||fa64||' '||fa65),'" & strTp(3) & "')" & strCheckWay & " ) A where fa10=na01(+) AND FA01=A.A1 " & StrSQLa
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','¡¯')||Decode(fa77,'Y','$','') AS ½s¸¹,fa06 as ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS ª¬ºA, FA29 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From fagent,nation, (Select Distinct FA01 As A1 From Fagent Where instr(fa06,'" & strTp(3) & "')" & strCheckWay & " ) A where fa10=na01(+) AND FA01=A.A1 " & StrSQLa
'
'            'Modify by Amy 2015/04/15 «È¤áºÝ¥­¥x±b¸¹¸ê®Æ
'            'Modified by Lydia 2017/02/14 + strfields
'            'Modify By Sindy 2021/3/25 '' as ®×¥ó©Ê½è, => CW03 as ®×¥ó©Ê½è,
'            strSql = strSql & " union all Select ' ' as V,'¥­¥x'||CW01 AS ½s¸¹, CW12 AS ¦WºÙ,'¥­¥x' AS °êÄy,' ' AS ´¼Åv¤H­û,Nvl(CW19,'') AS ª¬ºA,'' AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,CW03 as ®×¥ó©Ê½è,CW01 as ¦¬¤å¤é" & strFields & " From CustWeb Where InStr(Upper(CW12),'" & strTp(3) & "') " & strCheckWay
'
'            '¬dpotcustomer °ê¥~¼ç¦b«È¤á ÀÉ
'            'Modify by Amy 2020/03/16 ´¼Åv¤H­û ­ì:st02 ,¦]¶}µo¤H­û¥i¯à¦h¤H
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','¡¯') AS ½s¸¹,PCU08 AS ¦WºÙ,NA03 AS °êÄy,pcu38 AS ´¼Åv¤H­û,PCU39 AS ª¬ºA,PCU40 AS ³Æµù,' ' AS ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From potcustomer,nation,staff, (Select Distinct pcu01 As A1 From potcustomer Where instr(pcu08,'" & strTp(3) & "')" & strCheckWay & ") A where pcu09=na01(+) and pcu01=A.A1 and substr(LTrim(PCU38),1,5)=ST01(+)"
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','¡¯') AS ½s¸¹,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06) AS ¦WºÙ,NA03 AS °êÄy,pcu38 AS ´¼Åv¤H­û,PCU39 AS ª¬ºA,PCU40 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From potcustomer,nation,staff, (Select Distinct pcu01 As A1 From potcustomer Where instr(upper(pcu03||' '||pcu04||' '||pcu05||' '||pcu06),'" & strTp(3) & "')" & strCheckWay & " ) A where pcu09=na01(+) and pcu01=A.A1 and substr(LTrim(PCU38),1,5)=ST01(+)"
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','¡¯') AS ½s¸¹,PCU07 AS ¦WºÙ,NA03 AS °êÄy,pcu38 AS ´¼Åv¤H­û,PCU39 AS ª¬ºA,PCU40 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From potcustomer,nation,staff, (Select Distinct pcu01 As A1 From potcustomer Where instr(pCU07,'" & strTp(3) & "')" & strCheckWay & " ) A where pcu09=na01(+) and pcu01=A.A1 and substr(LTrim(PCU38),1,5)=ST01(+)"
'
'            '¬dpotcustomer1 °ê¤º¼ç¦b«È¤á ÀÉ
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all SELECT ' ' AS V,POC01||POC02||Decode(POC02,'0','','¡¯') AS ½s¸¹,POC03 AS ¦WºÙ,NA03 AS °êÄy,poc13 AS ´¼Åv¤H­û,POC14 AS ª¬ºA,POC15 AS ³Æµù,' ' AS ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From potcustomer1,nation,staff, (Select Distinct poc01 As A1 From potcustomer1 Where instr(poc03,'" & strTp(3) & "')" & strCheckWay & ") A where poc04=na01(+) and poc01=A.A1 and poc13=ST01(+)"
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all SELECT ' ' AS V,POC01||POC02||Decode(POC02,'0','','¡¯') AS ½s¸¹,RTRIM(POC23||' '||POC24||' '||POC25||' '||POC26) AS ¦WºÙ,NA03 AS °êÄy,poc13 AS ´¼Åv¤H­û,POC14 AS ª¬ºA,POC15 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From potcustomer1,nation,staff, (Select Distinct poc01 As A1 From potcustomer1 Where instr(upper(poc23||' '||poc24||' '||poc25||' '||poc26),'" & strTp(3) & "')" & strCheckWay & " ) A where poc04=na01(+) and poc01=A.A1 and poc13=ST01(+)"
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all SELECT ' ' AS V,POC01||POC02||Decode(POC02,'0','','¡¯') AS ½s¸¹,POC27 AS ¦WºÙ,NA03 AS °êÄy,poc13 AS ´¼Åv¤H­û,POC14 AS ª¬ºA,POC15 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From potcustomer1,nation,staff, (Select Distinct poc01 As A1 From potcustomer1 Where instr(POC27,'" & strTp(3) & "')" & strCheckWay & " ) A where poc04=na01(+) and poc01=A.A1 and poc13=ST01(+)"
'            'end 2020/03/16
'
'            '¬dNotAgent ¤£±o¥N²z®×¥ó¤§«È¤á©Î¥N²z¤H ÀÉ
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'¡ò','') AS ½s¸¹,NT02 as ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(NT21,null,'¤£±o¥N²z','') AS ª¬ºA, Decode(NT21,null,'','ºM¾P¤é´Á¡G'||sqldatet(NT21)||'¡F')||NT20 AS ³Æµù,' ' AS ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From notagent,nation,STAFF, (Select Distinct nt01 As A1 From notagent Where instr(nt02,'" & strTp(3) & "')" & strCheckWay & ") A where nt08=na01(+) AND nt01=A.A1 AND nt18=ST01(+)"
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'¡ò','') AS ½s¸¹,NT03||' '||NT04||' '||NT05||' '||NT06 as ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(NT21,null,'¤£±o¥N²z','') AS ª¬ºA, Decode(NT21,null,'','ºM¾P¤é´Á¡G'||sqldatet(NT21)||'¡F')||NT20 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From notagent,nation,STAFF, (Select Distinct nt01 As A1 From notagent Where instr(upper(nt03||' '||nt04||' '||nt05||' '||nt06),'" & strTp(3) & "')" & strCheckWay & " ) A where nt08=na01(+) AND nt01=A.A1 AND nt18=ST01(+)"
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'¡ò','') AS ½s¸¹,NT07 as ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(NT21,null,'¤£±o¥N²z','') AS ª¬ºA, Decode(NT21,null,'','ºM¾P¤é´Á¡G'||sqldatet(NT21)||'¡F')||NT20 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From notagent,nation,STAFF, (Select Distinct nt01 As A1 From notagent Where instr(nt07,'" & strTp(3) & "')" & strCheckWay & " ) A where nt08=na01(+) AND nt01=A.A1 AND nt18=ST01(+)"
'
'            '¬dÁpµ¸¤H(¤¤¤å)
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC05 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS ª¬ºA,PCC13 AS ³Æµù,' ' AS ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From (Select * From potcustcont Where instr(pcc05,'" & strTp(3) & "')" & strCheckWay & ") A,CUSTOMER,NATION,STAFF WHERE CU10=NA01(+) AND CU13=ST01(+) AND CU01(+)=PCC01 AND CU02='0' "
'            'Modify by Amy 2020/03/16 ´¼Åv¤H­û ­ì:st02 ,¦]¶}µo¤H­û¥i¯à¦h¤H
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC05 AS ¦WºÙ,NA03 AS °êÄy,pcu38 AS ´¼Åv¤H­û,PCU39 AS ª¬ºA,PCC13 AS ³Æµù,' ' AS ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From (Select * From potcustcont Where instr(pcc05,'" & strTp(3) & "')" & strCheckWay & ") A,potcustomer,nation,staff where pcu09=na01(+) AND PCU01(+)=PCC01 AND PCU02='0' and substr(LTrim(PCU38),1,5)=ST01(+) "
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC05 AS ¦WºÙ,NA03 AS °êÄy,poc13 AS ´¼Åv¤H­û,POC14 AS ª¬ºA,PCC13 AS ³Æµù,' ' AS ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From (Select * From potcustcont Where instr(pcc05,'" & strTp(3) & "')" & strCheckWay & ") A,potcustomer1,nation,staff where poc04=na01(+) AND POC01(+)=PCC01 AND POC02='0' and poc13=ST01(+) "
'            'end 2020/03/16
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC05 AS ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS ª¬ºA, PCC13 AS ³Æµù,' ' AS ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From (Select * From potcustcont Where instr(pcc05,'" & strTp(3) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' " & StrSQLa
'
'            '¬dÁpµ¸¤H(­^¤å)
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC03 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS ª¬ºA,PCC13 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From (Select * From potcustcont Where instr(upper(pcc03),'" & strTp(3) & "')" & strCheckWay & ") A,CUSTOMER,NATION,STAFF WHERE CU10=NA01(+) AND CU13=ST01(+) AND CU01(+)=PCC01 AND CU02='0' "
'            'Modify by Amy 2020/03/16 ´¼Åv¤H­û ­ì:st02 ,¦]¶}µo¤H­û¥i¯à¦h¤H
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC03 AS ¦WºÙ,NA03 AS °êÄy,pcu38 AS ´¼Åv¤H­û,PCU39 AS ª¬ºA,PCC13 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From (Select * From potcustcont Where instr(upper(pcc03),'" & strTp(3) & "')" & strCheckWay & ") A,potcustomer,nation,staff where pcu09=na01(+) AND PCU01(+)=PCC01 AND PCU02='0' and substr(LTrim(PCU38),1,5)=ST01(+) "
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC03 AS ¦WºÙ,NA03 AS °êÄy,poc13 AS ´¼Åv¤H­û,POC14 AS ª¬ºA,PCC13 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From (Select * From potcustcont Where instr(upper(pcc03),'" & strTp(3) & "')" & strCheckWay & ") A,potcustomer1,nation,staff where poc04=na01(+) AND POC01(+)=PCC01 AND POC02='0' and poc13=ST01(+) "
'            'end 2020/03/16
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC03 AS ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS ª¬ºA, PCC13 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From (Select * From potcustcont Where instr(upper(pcc03),'" & strTp(3) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' " & StrSQLa
'
'            '¬dÁpµ¸¤H(¤é¤å)
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC04 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS ª¬ºA,PCC13 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From (Select * From potcustcont Where instr(PCC04,'" & strTp(3) & "')" & strCheckWay & ") A,CUSTOMER,NATION,STAFF WHERE CU10=NA01(+) AND CU13=ST01(+) AND CU01(+)=PCC01 AND CU02='0' "
'            'Modify by Amy 2020/03/16 ´¼Åv¤H­û ­ì:st02 ,¦]¶}µo¤H­û¥i¯à¦h¤H
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC04 AS ¦WºÙ,NA03 AS °êÄy,pcu38 AS ´¼Åv¤H­û,PCU39 AS ª¬ºA,PCC13 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From (Select * From potcustcont Where instr(PCC04,'" & strTp(3) & "')" & strCheckWay & ") A,potcustomer,nation,staff where pcu09=na01(+) AND PCU01(+)=PCC01 AND PCU02='0' and substr(LTrim(PCU38),1,5)=ST01(+) "
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC04 AS ¦WºÙ,NA03 AS °êÄy,poc13 AS ´¼Åv¤H­û,POC14 AS ª¬ºA,PCC13 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From (Select * From potcustcont Where instr(PCC04,'" & strTp(3) & "')" & strCheckWay & ") A,potcustomer1,nation,staff where poc04=na01(+) AND POC01(+)=PCC01 AND POC02='0' and poc13=ST01(+) "
'            'end 2020/03/16
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS ½s¸¹,PCC04 AS ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS ª¬ºA, PCC13 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From (Select * From potcustcont Where instr(PCC04,'" & strTp(3) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' " & StrSQLa
'
'        'Modify by Amy 2014/04/30
'        If Check2.Value = 1 Then
'            '§ì¼È¦sÀÉ¹ï³y
'            'Modified by Lydia 2017/02/14 + strfields
'            'Modified by Lydia 2019/12/26 +@+Me.name
'            'Modify by Amy 2020/09/04 +all ¦]¬d ª÷§ù À³¥X²{2µ§,¤¤/¤é¤å³£¦³¿é
'            strSql = strSql & " union all Select ' ' as V,R021001 AS ½s¸¹,R021002 AS ¦WºÙ,'' AS °êÄy,'' AS ´¼Åv¤H­û,Decode(R021004,'1','¹ï³y','¨ä¥L¬ÛÃö¤H') AS ª¬ºA,'' AS ³Æµù,'' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "' And R021004<3 "
'        End If
'        'end 2014/04/30
'        'end 2015/03/27
'
'             'Mark 2014/02/21 ©¹¤W·h
''            '¹ï³y(¤¤)
''            strSubSQL1 = " And InStr(Upper(CP40),'" & strTp(3) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP50),'" & strTp(3) & "') " & strCheckWay
''            strSwhSQL1 = " CP40>' ' "
''            strSwhSQL2 = " CP50>' ' "
''            '°Ó¼Ð
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,Decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(TM29,'Y','¡¯','')||Decode(length(nvl(tm57,'')),null,'','¡´') as ½s¸¹, CP40 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL1 & _
''                        " Union  Select ' ' as V,Decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(TM29,'Y','¡¯','')||Decode(length(nvl(tm57,'')),null,'','¡´') as ½s¸¹, CP50 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL2
''            '±M§Q
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,Decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,'Y','¡¯','')||Decode(length(nvl(pa108,'')),null,'','¡´') as ½s¸¹,CP40 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL1 & _
''                        " Union  Select ' ' as V,Decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,'Y','¡¯','')||Decode(length(nvl(pa108,'')),null,'','¡´') as ½s¸¹,CP50 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL2
''            'ªk°È
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(LC08,'Y','¡¯','')||Decode(length(nvl(LC34,'')),null,'','¡´') AS ½s¸¹,CP40 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(LC08,'Y','¡¯','')||Decode(length(nvl(LC34,'')),null,'','¡´') AS ½s¸¹,CP50 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a ,CP09 as Á`¦¬¤å¸¹,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL2
''            'ÅU°Ý
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(HC09,'Y','¡¯','')||Decode(length(nvl(HC19,'')),null,'','¡´') as ½s¸¹,CP40 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(HC09,'Y','¡¯','')||Decode(length(nvl(HC19,'')),null,'','¡´') as ½s¸¹,CP50 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL2
''            'ªA°È
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(SP15,'Y','¡¯','')||Decode(length(nvl(SP61,'')),null,'','¡´') as ½s¸¹,CP40 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,SP09 as¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(SP15,'Y','¡¯','')||Decode(length(nvl(SP61,'')),null,'','¡´') as ½s¸¹,CP50 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,SP09 as¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL2
''
''            '¹ï³y(­^)
''            strSubSQL1 = " And InStr(Upper(CP41),'" & strTp(3) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP51),'" & strTp(3) & "') " & strCheckWay
''            strSwhSQL1 = " CP41>' ' "
''            strSwhSQL2 = " CP51>' ' "
''            '°Ó¼Ð
''            strSql = strSql & " Union " & _
''                         "Select ' ' as V,Decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(TM29,'Y','¡¯','')||Decode(length(nvl(tm57,'')),null,'','¡´') as ½s¸¹, CP41 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL1 & _
''                         " Union Select ' ' as V,Decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(TM29,'Y','¡¯','')||Decode(length(nvl(tm57,'')),null,'','¡´') as ½s¸¹, CP51 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL2
''            '±M§Q
''            strSql = strSql & " Union " & _
''                         "Select ' ' as V,Decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,'Y','¡¯','')||Decode(length(nvl(pa108,'')),null,'','¡´') as ½s¸¹,CP41 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL1 & _
''                         " Union Select ' ' as V,Decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,'Y','¡¯','')||Decode(length(nvl(pa108,'')),null,'','¡´') as ½s¸¹,CP51 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL2
''            'ªk°È
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(LC08,'Y','¡¯','')||Decode(length(nvl(LC34,'')),null,'','¡´') AS ½s¸¹,CP41 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(LC08,'Y','¡¯','')||Decode(length(nvl(LC34,'')),null,'','¡´') AS ½s¸¹,CP51 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL2
''            'ÅU°Ý
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(HC09,'Y','¡¯','')||Decode(length(nvl(HC19,'')),null,'','¡´') as ½s¸¹,CP41 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(HC09,'Y','¡¯','')||Decode(length(nvl(HC19,'')),null,'','¡´') as ½s¸¹,CP51 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL2
''            'ªA°È
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(SP15,'Y','¡¯','')||Decode(length(nvl(SP61,'')),null,'','¡´') as ½s¸¹,CP41 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,SP09 as¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(SP15,'Y','¡¯','')||Decode(length(nvl(SP61,'')),null,'','¡´') as ½s¸¹,CP51 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,SP09 as¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL2
''
''            '¹ï³y(¤é)
''            strSubSQL1 = " And InStr(Upper(CP42),'" & strTp(3) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP52),'" & strTp(3) & "') " & strCheckWay
''            strSwhSQL1 = " CP42>' ' "
''            strSwhSQL2 = " CP52>' ' "
''            '°Ó¼Ð
''            strSql = strSql & " Union " & _
''                         "Select ' ' as V,Decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(TM29,'Y','¡¯','')||Decode(length(nvl(tm57,'')),null,'','¡´') as ½s¸¹, CP42 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL1 & _
''                         " Union Select ' ' as V,Decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(TM29,'Y','¡¯','')||Decode(length(nvl(tm57,'')),null,'','¡´') as ½s¸¹, CP52 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL2
''            '±M§Q
''            strSql = strSql & " Union " & _
''                         "Select ' ' as V,Decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,'Y','¡¯','')||Decode(length(nvl(pa108,'')),null,'','¡´') as ½s¸¹,CP42 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL1 & _
''                         " Union Select ' ' as V,Decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,'Y','¡¯','')||Decode(length(nvl(pa108,'')),null,'','¡´') as ½s¸¹,CP52 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL2
''            'ªk°È
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(LC08,'Y','¡¯','')||Decode(length(nvl(LC34,'')),null,'','¡´') AS ½s¸¹,CP42 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(LC08,'Y','¡¯','')||Decode(length(nvl(LC34,'')),null,'','¡´') AS ½s¸¹,CP52 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL2
''            'ÅU°Ý
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(HC09,'Y','¡¯','')||Decode(length(nvl(HC19,'')),null,'','¡´') as ½s¸¹,CP42 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(HC09,'Y','¡¯','')||Decode(length(nvl(HC19,'')),null,'','¡´') as ½s¸¹,CP52 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,' ' as ¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL2
''            'ªA°È
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(SP15,'Y','¡¯','')||Decode(length(nvl(SP61,'')),null,'','¡´') as ½s¸¹,CP42 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¹ï³y' as ª¬ºA,' ' as ³Æµù,SP09 as¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(SP15,'Y','¡¯','')||Decode(length(nvl(SP61,'')),null,'','¡´') as ½s¸¹,CP52 as ¦WºÙ,' ' as °êÄy,' ' as ´¼Åv¤H,'¨ä¥L¬ÛÃö¤H' as ª¬ºA,' ' as ³Æµù,SP09 as¥Ó½Ð°ê®a,CP09 as Á`¦¬¤å¸¹,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) as ®×¥ó©Ê½è,Nvl(To_Char(cp05-19110000),'') as ¦¬¤å¤é " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL2
'             'end Mark 2014/02/21
'
''end 2013/11/19
'        pub_QL05 = pub_QL05 & ";" & Option2(1).Caption & Trim(Text2) 'Add By Sindy 2010/10/22
'
'        ' Add By Sindy 98/02/13 ¶}©Ý«È¤á
'        If Check1.Value = 1 Then
'            'Modify by Amy 2013/11/06 +¥Ó½Ð°ê®a/Á`¦¬¤å¸¹/®×¥ó©Ê½è/¦¬¤å¤é
'            'Modify by Amy 2013/09/27 ­ì¥uÀË¬decd11,ecd12«oÅã¥Üecd03,ecd04
'            'strSql = strSql & " union all SELECT ' ' AS V,ecd02||'-'||LPAD(ecd01,6,'0') AS ½s¸¹,NVL(ecd03,'')||NVL(ecd04,'') AS ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,'§ëªk¶}©Ý'||Decode(ecd15,null,null,'-'||ecd15) AS ª¬ºA,ecd16 AS ³Æµù From expandcusdetail, expandcusattr, nation,(Select Distinct nvl(ecd01,'')||nvl(ecd02,'') as A1 From expandcusdetail Where instr(ecd11,'" & ChgSQL(Trim(Text2)) & "') " & strCheckWay & " or instr(ecd12,'" & ChgSQL(Trim(Text2)) & "') " & strCheckWay & ") A Where ecd10=na01(+) and ecd02=eca01(+) and nvl(ecd01,'')||nvl(ecd02,'')=A.A1 "
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all SELECT ' ' AS V,ecd02||'-'||LPAD(ecd01,6,'0') AS ½s¸¹,NVL(ecd03,'')||NVL(ecd04,'') AS ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,'§ëªk¶}©Ý'||Decode(ecd15,null,null,'-'||ecd15) AS ª¬ºA,ecd16 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From expandcusdetail, expandcusattr, nation,(Select Distinct nvl(ecd01,'')||nvl(ecd02,'') as A1 From expandcusdetail Where instr(Upper(ecd03),'" & ChgSQL(UCase(Trim(Text2))) & "') " & strCheckWay & " or instr(Upper(ecd04),'" & ChgSQL(UCase(Trim(Text2))) & "') " & strCheckWay & ") A Where ecd10=na01(+) and ecd02=eca01(+) and nvl(ecd01,'')||nvl(ecd02,'')=A.A1 "
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all SELECT ' ' AS V,ecd02||'-'||LPAD(ecd01,6,'0') AS ½s¸¹,NVL(ecd11,'')||NVL(ecd12,'') AS ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,'§ëªk¶}©Ý'||Decode(ecd15,null,null,'-'||ecd15) AS ª¬ºA,ecd16 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " From expandcusdetail, expandcusattr, nation,(Select Distinct nvl(ecd01,'')||nvl(ecd02,'') as A1 From expandcusdetail Where instr(Upper(ecd11),'" & ChgSQL(UCase(Trim(Text2))) & "') " & strCheckWay & " or instr(Upper(ecd12),'" & ChgSQL(UCase(Trim(Text2))) & "') " & strCheckWay & ") A Where ecd10=na01(+) and ecd02=eca01(+) and nvl(ecd01,'')||nvl(ecd02,'')=A.A1 "
'        End If
'        ' 98/02/13 End
'
'   'add by nickc 2007/10/24  ­t³d¤H»P±µ¬¢¤H¤£¥Î§ì¥N²z¤HÀÉ¡A¦]¬°¨S¦³
'   ElseIf Option2(2).Value = True Then
'       'Modify by Amy 2013/11/06 +¥Ó½Ð°ê®a/Á`¦¬¤å¸¹/®×¥ó©Ê½è/¦¬¤å¤é
'       'edit by nickc 2008/01/03 ¥[¤J¯S®í«È¤á
'       'strSQL = "SELECT ' ' as V,CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','') AS ½s¸¹,CU04 AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,CU80 AS ª¬ºA,CU79 AS ³Æµù FROM CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(CU07,'" & ChgSQL(Text9) & "')>=1 ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+)"
'       'Modified by Lydia 2017/02/14 + strfields
'       strSql = "SELECT ' ' as V,CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','¡´','') AS ½s¸¹,NVL(CU04,Decode(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS ª¬ºA,CU79 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " FROM CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(CU07,'" & ChgSQL(Trim(Text9)) & "')>=1 ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+)"
'       pub_QL05 = pub_QL05 & ";" & Option2(2).Caption & Trim(Text9) 'Add By Sindy 2010/10/22
'
'   'E-Mail  add by Toni 2008/12/03
'   ElseIf Option2(3).Value = True Then
'        'Modify by Amy 2013/11/06 +¥Ó½Ð°ê®a/Á`¦¬¤å¸¹/®×¥ó©Ê½è/¦¬¤å¤é
'        'Modified by Lydia 2017/02/14 + strfields
'        strSql = "SELECT ' ' as V,CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','¡´','') AS ½s¸¹,NVL(CU04,Decode(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS ª¬ºA,CU79 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " FROM CUSTOMER,NATION,staff  Where (instr(NLS_Upper(CU20),'" & UCase(ChgSQL(Trim(Text10))) & "')>0 Or instr(NLS_Upper(CU115),'" & UCase(ChgSQL(Trim(Text10))) & "')>0 or instr(NLS_Upper(CU116),'" & UCase(ChgSQL(Trim(Text10))) & "')>0  or instr(NLS_Upper(CU117),'" & UCase(ChgSQL(Trim(Text10))) & "')>0 or instr(NLS_Upper(CU118),'" & UCase(ChgSQL(Trim(Text10))) & "') > 0 )  and CU10=NA01(+)  AND CU13=ST01(+)"
'        'Modify by Amy 2020/03/16 ´¼Åv¤H­û ­ì:st02 ,¦]¶}µo¤H­û¥i¯à¦h¤H
'        'Modified by Lydia 2017/02/14 + strfields
'        strSql = strSql & " union all SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','¡¯') AS ½s¸¹,NVL(PCU08,Decode(PCU03,null,PCU07,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06))) AS ¦WºÙ,NA03 AS °êÄy,pcu38 AS ´¼Åv¤H­û,PCU39 AS ª¬ºA,PCU40 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " FROM potcustomer,nation,staff Where (instr(NLS_Upper(pcu18),'" & UCase(ChgSQL(Trim(Text10))) & "') >0 ) and pcu09=na01(+) and substr(LTrim(PCU38),1,5)=ST01(+)"
'        'Add By Sindy 98/03/19
'        'Modified by Lydia 2017/02/14 + strfields
'        strSql = strSql & " union all SELECT ' ' AS V,POC01||POC02||Decode(POC02,'0','','¡¯') AS ½s¸¹,NVL(PoC03,Decode(PoC23,null,PoC27,RTRIM(PoC23||' '||PoC24||' '||PoC25||' '||PoC26))) AS ¦WºÙ,NA03 AS °êÄy,poc13 AS ´¼Åv¤H­û,POC14 AS ª¬ºA,POC15 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " FROM potcustomer1,nation,staff Where (instr(NLS_Upper(poc09),'" & UCase(ChgSQL(Trim(Text10))) & "') >0 ) and poc04=na01(+) and poc13=ST01(+)"
'        '98/03/19 End
'        'end 2020/03/16
'
'        'Modified by Lydia 2017/02/14 + strfields
'        'Modified by Lydia 2018/07/20 +FA105 °]°È«H½c(CF)
'        'strSql = strSql & " union all SELECT ' ' AS V,FA01||FA02||Decode(FA02,'0','','¡¯')||Decode(fa77,'Y','$','') AS ½s¸¹,NVL(fa04,Decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) as ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS ª¬ºA, FA29 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " FROM fagent,nation Where (instr(NLS_Upper(fa16),'" & UCase(ChgSQL(Trim(Text10))) & "')> 0 or instr(NLS_Upper(fa79),'" & UCase(ChgSQL(Trim(Text10))) & "')> 0 or instr(NLS_Upper(fa80),'" & UCase(ChgSQL(Trim(Text10))) & "')> 0 or instr(NLS_Upper(fa81),'" & UCase(ChgSQL(Trim(Text10))) & "') > 0 Or InStr(NLS_Upper(fa82),'" & UCase(ChgSQL(Trim(Text10))) & "') > 0 )  and fa10=na01(+) " & StrSQLa
'        strSql = strSql & " union all SELECT ' ' AS V,FA01||FA02||Decode(FA02,'0','','¡¯')||Decode(fa77,'Y','$','') AS ½s¸¹,NVL(fa04,Decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) as ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS ª¬ºA, FA29 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & _
'                    " FROM fagent,nation Where (instr(NLS_Upper(fa16),'" & UCase(ChgSQL(Trim(Text10))) & "')> 0 or instr(NLS_Upper(fa79),'" & UCase(ChgSQL(Trim(Text10))) & "')> 0 or instr(NLS_Upper(fa105),'" & UCase(ChgSQL(Trim(Text10))) & "')> 0 or instr(NLS_Upper(fa80),'" & UCase(ChgSQL(Trim(Text10))) & "')> 0 or instr(NLS_Upper(fa81),'" & UCase(ChgSQL(Trim(Text10))) & "') > 0 Or InStr(NLS_Upper(fa82),'" & UCase(ChgSQL(Trim(Text10))) & "') > 0 )  and fa10=na01(+) " & StrSQLa
'        'Modified by Lydia 2017/02/14 + strfields
'        strSql = strSql & " union all SELECT ' ' AS V,PCC01||'0-'||PCC02 AS ½s¸¹,NVL(PCC05,NVL(PCC03,PCC04)) AS ¦WºÙ,' ' AS °êÄy,' ' AS ´¼Åv¤H­û,' ' AS ª¬ºA,PCC13 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " FROM PotCustCont Where (instr(NLS_Upper(PCC08),'" & UCase(ChgSQL(Trim(Text10))) & "') > 0 )  "
'
'        pub_QL05 = pub_QL05 & ";" & Option2(3).Caption & Trim(Text10) 'Add By Sindy 2010/10/22
'
'        ' Add By Sindy 98/02/13 ¶}©Ý«È¤á
'        If Check1.Value = 1 Then
'            'Modify by Amy 2013/11/06 +¥Ó½Ð°ê®a/Á`¦¬¤å¸¹/®×¥ó©Ê½è/¦¬¤å¤é
'            'Modify by Amy 2013/09/27 ­ì:ecd15 AS ª¬ºA
'            'Modified by Lydia 2017/02/14 + strfields
'            strSql = strSql & " union all SELECT ' ' AS V,ecd02||'-'||LPAD(ecd01,6,'0') AS ½s¸¹,ecd03||' '||ecd04 AS ¦WºÙ,NA03 AS °êÄy,' ' AS ´¼Åv¤H­û,'§ëªk¶}©Ý'||Decode(ecd15,null,null,'-'||ecd15) AS ª¬ºA,ecd16 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " FROM expandcusdetail, expandcusattr, nation Where (instr(NLS_Upper(ecd13),'" & UCase(ChgSQL(Trim(Text10))) & "') > 0 ) and ecd10=na01(+) and ecd02=eca01(+) "
'        End If
'        ' 98/02/13 End
'
'   'add by nickc 2008/05/02
'   ElseIf Option2(4).Value = True Then
'       'Modify by Amy 2013/11/06 +¥Ó½Ð°ê®a/Á`¦¬¤å¸¹/®×¥ó©Ê½è/¦¬¤å¤é
'       'Modified by Lydia 2017/02/14 + strfields
'       strSql = "SELECT ' ' as V,CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','¡´','') AS ½s¸¹,NVL(CU04,Decode(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS ¦WºÙ,NA03 AS °êÄy,ST02 AS ´¼Åv¤H­û,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS ª¬ºA,CU79 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é" & strFields & " FROM CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(CU11,'" & ChgSQL(Trim(Text11)) & "')>=1 ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+)"
'       pub_QL05 = pub_QL05 & ";" & Option2(4).Caption & Trim(Text11) 'Add By Sindy 2010/10/22
'   End If
'
'   '2008/12/3 add by sonia
'   'Modify by Amy 2019/09/17 ¥[«Ý¬¡¤Æ«È¤á
'   If Option2(1).Value = True Then
'      'Modify by Amy 2014/01/15 +½s¸¹±Æ
'      strSql = "select X.*,Decode(Ocu01,null, '',NVL(Ocu03,0)) as OCU03 from (" & strSql & ") X, OldCustomer Where substr(½s¸¹,1,8)= ocu01(+) order by upper(¦WºÙ),½s¸¹ "
'   Else
'      strSql = "select X.*,Decode(Ocu01,null, '',NVL(Ocu03,0)) as OCU03 from (" & strSql & ") X, OldCustomer Where substr(½s¸¹,1,8)= ocu01(+) order by ½s¸¹ "
'   End If
'   'end 2019/09/17
'   '2008/12/3 end
'
'   If Check1.Value = 1 Then
'      pub_QL05 = pub_QL05 & ";" & Check1.Caption 'Add By Sindy 2010/10/22
'   End If
'
'   CheckOC
'   adoRecordset.CursorLocation = adUseClient
'   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If adoRecordset.RecordCount <> 0 Then
'       InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/10/22
'       If Not cmdOK(0).Enabled Then cmdOK(0).Enabled = True
'       If Not cmdOK(1).Enabled Then cmdOK(1).Enabled = True
'       If Not cmdOK(2).Enabled Then cmdOK(2).Enabled = True
'       If Not cmdOK(5).Enabled Then cmdOK(5).Enabled = True
'       If Not cmdOK(6).Enabled Then cmdOK(6).Enabled = True
'       If Not cmdOK(7).Enabled Then cmdOK(7).Enabled = True
'       Set GrdDataList.Recordset = adoRecordset
'   Else
'       InsertQueryLog (0) 'Add By Sindy 2010/10/22
'       'Modify by Amy 2013/11/06 Mark If Option2(1).Value = True And Trim(Text2) <> "" Then ¤£»Ý¦A§ä¹ï³y
''       'Add By Sindy 2010/02/05
''       If Option2(1).Value = True And Trim(Text2) <> "" Then
''          Pub_Can_Copy_Pic = True 'Added by Morgan 2011/12/26
''          MsgBox "«D¥»©Ò«È¤á©Î¥N²z¤H¡A¨t²Î·|¦A·j´M®×¥ó¹ï³y¸ê®Æ¡A½Ðª`·N¬O§_¦³Âù¤è¥N²z±¡§Î¡I", vbInformation, "¨S¦³¸ê®Æ " & Now
''          Pub_Can_Copy_Pic = False 'Added by Morgan 2011/12/26
''          Me.Enabled = False
''          frm100110_3.Show 'Added by Morgan 2012/8/8 ­n¥ý©I¥s¤~¤£·|Ä²µo¨ä¥Lµøµ¡ªº Form_Activate ¨Æ¥ó
''          If fnSaveParentForm(Me) = False Then
''             Me.Enabled = True
''             Exit Sub
''          End If
''          Screen.MousePointer = vbHourglass
''          'Me.Hide 'Removed by Morgan 2012/8/8 ¤£»Ý­n
'''          frm100110_1.Option1(1).Value = True
'''          frm100110_1.txt1(1) = Trim(Text2)
'''          frm100110_1.Hide
''          'frm100110_3.Show 'Removed by Morgan 2012/8/8 ²¾¨ì¤W­±
''          Call frm100110_3.StrMenu_2(Trim(Text2))
'''          Unload frm100110_1
''          Screen.MousePointer = vbDefault
''
''   '       Do
''   '       DoEvents
''   '       If bolToEndByNick = True Then Unload Me: Exit Sub
''   '       Loop Until Not frm100110_3.Visible
''   '       Unload frm100110_3
''
''          Me.Enabled = True
''   '       If frm100110_3.Visible = False Then
''   '         Me.Show
''   '       End If
''          Exit Sub
''       '2010/02/05 End
''       Else
'          'Add by Amy 2013/11/06 +µe­±°T®§¶}©ñ¥i¦C¦L
'          Pub_Can_Copy_Pic = True
'          ShowNoData
'          Pub_Can_Copy_Pic = False
'          'end 2013/11/06
'          cmdOK(0).Enabled = False
'          cmdOK(1).Enabled = False
'          cmdOK(2).Enabled = False
'          cmdOK(5).Enabled = False
'          cmdOK(6).Enabled = False
'          cmdOK(7).Enabled = False
'          GrdDataList.Clear
''       End If
'   End If
'
'   Me.GrdDataList.Visible = False 'Add by Amy 2013/11/06
'   SetDataListWidth
'   CheckOC
'
'   With Me.GrdDataList
'        If .Rows > 0 Then 'Add by Amy 2013/11/19
'            For i = 1 To .Rows - 1
'                .row = i
'                .col = 1
'                .CellForeColor = &H0   '¦r¶Â¦â 'Modfiy by Amy 2019/08/29 ­ì:ForeColor ¬d»ö¤j·|¾ã­ÓÅÜ¶Â
'                'Add by Amy 2019/09/17 ¬¡¤Æ«È¤áÅã¥Ü¾ã¦C¶À,­Y¦³§b±b½s¸¹©³¬°¬õ¨ä¥LÄæ¬°¶À
'                If .TextMatrix(i, 15) = "0" And Right(.Text, 1) <> "¡¯" Then
'                    For j = 0 To .Cols - 1
'                        If Right(.Text, 1) = "$" And j = 1 Then
'                        Else
'                            .col = j
'                            .CellBackColor = vbYellow
'                        End If
'                    Next
'                ElseIf Right(.Text, 1) = "$" Then '§b±b
'                    .CellBackColor = &HFF& '¬õ¦â
'                    'Add By Sindy 2012/3/21
'                'Add by Amy 2019/08/28 +«È¤áª¬ºA¬° ¾E²¾¤£©ú/¸Ñ´²/¼o¤î/ºM¾P/°±·~/¦º¤` Åã¥Ü¶Â©³¯»¦r
'                'Modify by Amy 2022/05/23 ®³±¼ ¾E²¾¤£©ú ¤Î °±·~
'                ElseIf (Left(.Text, 1) = "Y" Or Left(.Text, 1) = "X" Or Left(.Text, 1) = "R") _
'                  And (.TextMatrix(i, 5) = "¸Ñ´²" Or .TextMatrix(i, 5) = "¼o¤î" Or .TextMatrix(i, 5) = "ºM¾P" Or .TextMatrix(i, 5) = "¦º¤`") Then
'                        For j = 0 To .Cols - 1
'                            .col = j
'                            .CellBackColor = &H0 '¶Â¦â
'                            .CellForeColor = &HFF00FF '¯»¬õ¦â  'Modfiy by Amy 2019/08/29 ­ì:ForeColor
'                        Next j
'                ElseIf Right(.Text, 1) = "¡ò" Or .TextMatrix(i, 5) = "¹ï³y" Or .TextMatrix(i, 5) = "¨ä¥L¬ÛÃö¤H" Then
'                    'Modify by Amy 2013/11/06 ¹ï³y­«§ì´¼Åv¤H¸ê®Æ
'                    If Me.GrdDataList.TextMatrix(i, 5) = "¹ï³y" Or .TextMatrix(i, 5) = "¨ä¥L¬ÛÃö¤H" Then
'                        bolPrint = True '¦³¹ï³y¸ê®Æ
'                        strNo = Pub_RplStr(.TextMatrix(i, 1))
'                        StrToPrint = strNo & ","
'                        Str01 = SystemNumber(strNo, 1)
'                        Select Case Str01
'                            Case "FCP", "FG"
'                                .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetFCPSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                            Case "FCL", "LIN"
'                                .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetFCLSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                            Case "FCT"
'                                .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetFCTSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                            Case "S"
'                                If .TextMatrix(i, 7) = "000" Then
'                                    .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetFCTSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                                Else
'                                    .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetAKindSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                                End If
'                            Case Else
'                                .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetAKindSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                        End Select
'                        .TextMatrix(i, 9) = .TextMatrix(i, 9) & PUB_GetRelateCasePropertyName(.TextMatrix(i, 8), "1")
'                        'Add by Amy 2014/02/21 §ó·s´¼Åv¤H­û¦Ü¼È¦sÀÉ
'                        strExc(0) = "Update R100102_1 Set R021003='" & .TextMatrix(i, 4) & "' Where R021014='" & Str01 & "' And R021015='" & SystemNumber(strNo, 2) & "' And R021016='" & SystemNumber(strNo, 3) & "' And R021017='" & SystemNumber(strNo, 4) & "' "
'                        cnnConnection.Execute strExc(0)
'                        'end 2014/02/21
'                    End If
'                    'end 2013/11/06
'                    If Right(.Text, 1) = "¡ò" Or .TextMatrix(i, 5) = "¹ï³y" Then
'                        For j = 0 To .Cols - 1
'                            .col = j
'                            .CellBackColor = &H8080FF
'                        Next j
'                    End If
'                    '2012/3/21 End
'
'                'Add By Sindy 2021/3/25 °w¹ïCW03=7.´C¤¶¥­¥x¡A¦b¬d¸ß¨t²ÎÅã¥Üµ²ªG¬°¾ï¦â
'                ElseIf Left(.TextMatrix(i, 1), 1) = "¥­" And .TextMatrix(i, 9) = "7" Then
'                    .CellBackColor = &H80FF& '¾ï¦â
'                '2021/3/25 END
'                End If
'
'                'Add by Amy 2020/03/16 °ê¤º¥~¼ç¦b«È¤á ´¼Åv¤H­ûÄæ»Ý­«§ì¸ê®Æ(¥i¯à¦hµ§)
'                If Left(.Text, 1) = "R" Then
'                    .TextMatrix(i, 4) = GetDevelopP(.TextMatrix(i, 4))
'                End If
'            Next i
'      End If 'end 2013/11/19
'   End With
'
'   '­Y¥u¦³¤@µ§¸ê®Æ, «hª½±µ³]©w¬°ÂI¿ï¦¹µ§¸ê®Æ
'   With Me.GrdDataList
'      If .Rows = 2 Then
'         .row = 1
'         .col = 1
'         If .Text <> "" Then
'           .Visible = False
'           .row = 1
'           .col = 0
'           .Text = "V"
'           For i = 0 To .Cols - 1
'               'Modify By Sindy 2012/3/21 old:If i <> 1 Then
'               If i <> 1 And (i = 2 And Right(.TextMatrix(1, 1), 1) = "¡ò") = False Then
'                   .col = i
'                   .CellBackColor = &HFFC0C0
'               End If
'           Next i
'           'Add by Amy 2020/10/15 ¤Ä¿ï®É§PÂ_¦³©¹¨Ó°O¿ý,©¹¨Ó°O¿ý¶sÅÜ¦â
'           Call ChkContactRecordBT(.TextMatrix(1, 0), Left(.TextMatrix(1, 1), 8))
'           .Visible = True
'         End If
'      End If
'   End With
'   'Add by Amy 2013/11/06
'   Me.GrdDataList.Visible = True
'   If bolPrint Then
'        cmdOK(10).Enabled = True
'   Else
'        cmdOK(10).Enabled = False
'   End If
'   'end 2013/11/06
'   Screen.MousePointer = vbDefault
'   Exit Sub
'
'ErrHnd:
'    If Err.Number = -2147217900 Then
'        MsgBox "¿é¤Jªº¤å¦rµLªk¬d¸ß,½Ð¹q¸£¤¤¤ß¨ó§U¡I"
'    Else
'        MsgBox Err.Description
'    End If
'    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   pub_QL05 = m_pub_QL05 'Add By Sindy 2025/8/13 ÁÙ­ì¦¹Formªº¬d¸ß±ø¥ó°O¿ý
End Sub

Private Sub Form_Load()
   'Memo by Amy 2023/08/24 index=4 [°ê¤ºA4¦W±ø] ¶s¦WºÙ¦³­×§ïPubShowNextForm¤]­n§ï
   
   bolToEndByNick = False
   MoveFormToCenter Me
   'Frame2.Left = 1470 'Modify 2013/12/04 Add by Amy 2013/11/19 ÁôÂÃ¤¤­^¤é,§ïframe2¦ì¸m
   SetDataListWidth
   GetField 'Add by Amy 2023/03/08
   cmdOK(0).Enabled = False
   cmdOK(1).Enabled = False
   cmdOK(2).Enabled = False
   cmdOK(5).Enabled = False
   cmdOK(4).Enabled = False
   cmdOK(6).Enabled = False
   cmdOK(7).Enabled = False
   Option2(0).Value = True
   Option1(0).Enabled = False
   Option1(1).Enabled = False
   Option1(2).Enabled = False
   Option3(0).Enabled = False
   Option3(1).Enabled = False
   'Modify by Amy 2014/04/30 ¥Ñ¬d¥»©Ò«È¤á¿ï¶µ¶i¤J ¹w³] ¬d¦r­º ¤£¬d¹ï³y
   If IsSearchNew = False Then
        Option3(0).Value = True
        Check2.Value = 0
   Else
        Option3(0).Value = False
        Check2.Value = 1
   End If
   'end 2014/04/30
   
   '2011/12/6 modify by sonia
   'Text3 = Systemkind_g
   Me.chk.Value = vbChecked
   Text3 = "ALL"
   '2011/12/6 end
   bolToEndByNick = False
   m_bolPrintRight = IsUserHasRightOfFunction("frm100102_1", strPrint, False)
   Me.cmdOK(4).Enabled = m_bolPrintRight
   cmdState = -1
   Label2(0).Caption = Label2(0).Caption & "¡þµµ©³¬°­·ÀIÄµ¥Ü" 'Modify by Amy 2024/01/31 +­·ÀIÀË¬d¹ï¶H,®³±¼­·ÀIÄµ¥Ü±Ò¥Î¤é
   ' Add By Sindy 98/02/16
   'MODIFY BY SONIA 2015/5/20 ¦]P31¤ÎF31¤H­û¨Ö¤JL02,¦ý¤º¥~ªk¤£¶}©ñÅv­­,¬G§ï¥Î­û¤uµ¥¯Å±±¨î
   'If Pub_StrUserSt03 = "F31" Or Pub_StrUserSt03 = "F41" Then
   If Pub_strUserST05 >= "51" And Pub_strUserST05 <= "55" Then
      Check1.Value = 1
   Else
      Check1.Value = 0
   End If
   ' 98/02/16 End
   
   'Added by Lydia 2016/11/04 Åã¥Ü¥¼¦C¦LªºA4¦W±ø¼Æ¶q
    If PUB_AddAddressA4List("", strExc(0)) Then
    End If
    'Modified by Lydia 2017/11/22 +°ê¤º
    If Val(strExc(0)) > 0 Then cmdOK(4).Caption = "°ê¤ºA4¦W±ø (" & Val(strExc(0)) & ")"
   'end 2016/11/04
   
   'Added by Lydia 2017/12/05 §ï¥Ñ±Ò¥Î¤é±±¨î
   If strSrvDate(1) >= °ê¥~³¡ÃöÁp¥ø·~±Ò¥Î¤é Then cmdOK(2).Caption = "ÃöÁp¥ø·~"
   'Add by Amy 2023/08/17 ¬d¸ß¸m´«¦r ¶s¥u¦³¹q¸£¤¤¤ß¤~¥X²{
   cmdMemo.Visible = False
   If Pub_StrUserSt03 = "M51" Then cmdMemo.Visible = True
   'end 2023/08/17
   Check2.Visible = False 'Add by Amy 2023/09/14 µ{¦¡¥Î,¬GÁôÂÃ
   m_blnColOrderAsc = True 'Add by Amy 2020/06/16
   SeekPrintL = Printer.Orientation
   'Mark by Lydia 2024/03/13
   'PUB_SetPrinter Me.Name, Me.Combo1, , , SeekPrint, , , True  'Modified by Moragn 2021/6/23 +¥uÅã¥Ü¦³®Äªº¦Lªí¾÷°Ñ¼Æ
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '­Y¦Lªí¾÷©Î°¾²¾­È¦³ÅÜ°Ê, «h§ó·s¦C¦L³]©w
   'Mark by Lydia 2024/03/13
   'If Me.Combo1.Text <> Me.Combo1.Tag Then
   '    PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, 0, 0, Me.Combo1.Text
   'End If
   'end 2024/03/13
   'Modified by Morgan 2021/6/23
   'Set Printer = Printers(SeekPrint)
   'Mark by Lydia 2024/03/13
   'PUB_RestorePrinter Combo1.List(SeekPrint)
   ''end 2021/6/23
   'If SeekPrintL <> 0 Then
   '    Printer.Orientation = SeekPrintL
   'End If
   'end 2024/03/13
   'Set frm100102_1 = Nothing 'Remove by Lydia 2021/12/16 Form2.0·|¦³°ÝÃD¡A§ï¦b©I¥s®É²M°£°O¾ÐÅéÅÜ¼Æ
End Sub

'ÃöÁp¥ø·~(°ê¥~³¡ÃöÁp¥ø·~±Ò¥Î¤é«eªº§ìªk)
Sub StrMenu(StrToGrid)
   '¤w¥Ó½Ð¤H¬d¸ß¤§¸ê®Æ®w
   '¥H½s¸¹ LIKE
   'edit by nickc 2008/01/03 ¥[¤J¯S®í«È¤á
   'strSQL = "SELECT CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$',''),NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),NA03,CU80,CU79 FROM CUSTOMER,NATION WHERE CU10=NA01(+) AND CU01>='" & Left(StrToGrid, 6) & "00' AND CU01<='" & Left(StrToGrid, 6) & "zz' "
   strSql = "SELECT CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','¡´',''),NVL(CU04,Decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,CU80,CU79 FROM CUSTOMER,NATION WHERE CU10=NA01(+) AND CU01>='" & Left(StrToGrid, 6) & "00' AND CU01<='" & Left(StrToGrid, 6) & "zz' "
   strSql = strSql & " union SELECT FA01||FA02||Decode(FA02,'0','','¡¯')||Decode(fa77,'Y','$',''),Decode(FA10,'013',NVL(FA04,Decode(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),'020',NVL(FA04,Decode(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),Decode(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)),NA03,' ',FA29 FROM FAGENT,NATION WHERE FA01>='" & Left(StrToGrid, 6) & "00' AND FA01<='" & Left(StrToGrid, 6) & "zz' AND fa10=NA01(+) "
   'Add By Sindy 98/03/19
   strSql = strSql & " union  SELECT PCU01||PCU02||Decode(PCU02,'0','','¡¯'),NVL(PCU08,Decode(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)),NA03,PCU39,PCU40 FROM PotCustomer,Nation WHERE PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz'   AND NA01(+)=PCU09"
   strSql = strSql & " union  SELECT POC01||POC02||Decode(POC02,'0','','¡¯'),POC03,NA03,POC14,POC15 FROM PotCustomer1,Nation WHERE POC01>='" & Left(StrToGrid, 6) & "00' AND POC01<='" & Left(StrToGrid, 6) & "zz'   AND NA01(+)=POC04"
   '¶Ç¤JR1®É§ä¥X¬ÛÃöªºX
   strSql = strSql & " union  SELECT CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','¡´',''),NVL(CU04,Decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,CU80,CU79 " & _
                                                    "From CUSTOMER, PotCustomer1, Nation " & _
                                               "WHERE CU10=NA01(+) " & _
                                                    "AND POC01>='" & Left(StrToGrid, 6) & "00' AND POC01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND CU01>=(substr(POC16,1,6)||'00') AND CU01<=(substr(POC16,1,6)||'zz') " & _
                                                    "AND POC16 is not null "
   '§ä¥XR1ªºÃö«Y¥ø·~
   strSql = strSql & " union  SELECT POC01||POC02||Decode(POC02,'0','','¡¯'),POC03,NA03,POC14,POC15 " & _
                                                    "From PotCustomer1, Nation " & _
                                                "WHERE NA01(+)=POC04 " & _
                                                     "AND POC16>='" & Left(StrToGrid, 6) & "00' AND POC16<='" & Left(StrToGrid, 6) & "zz' " & _
                                                     "AND POC16 is not null "
   '¶Ç¤JR1®É§ä¥X¬ÛÃöªºR
   strSql = strSql & " union  SELECT PCU01||PCU02||Decode(PCU02,'0','','¡¯'),NVL(PCU08,Decode(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)),NA03,PCU39,PCU40 " & _
                                                    "From PotCustomer, Nation, PotCustomer1 " & _
                                               "WHERE NA01(+)=PCU09 " & _
                                                    "AND POC01>='" & Left(StrToGrid, 6) & "00' AND POC01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND PCU47>=(substr(POC16,1,6)||'00') AND PCU47<=(substr(POC16,1,6)||'zz') " & _
                                                    "AND POC16 is not null AND PCU47 is not null "
   '98/03/19 End
   'Add By Sindy 2009/06/24
   '¶Ç¤JR®É§ä¥X¬ÛÃöªºX
   strSql = strSql & " union  SELECT CU01||CU02||Decode(CU02,'0','','¡¯')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','¡´',''),NVL(CU04,Decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,CU80,CU79 " & _
                                                    "From CUSTOMER, PotCustomer, Nation " & _
                                               "WHERE CU10=NA01(+) " & _
                                                    "AND PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND CU01>=(substr(PCU47,1,6)||'00') AND CU01<=(substr(PCU47,1,6)||'zz') " & _
                                                    "AND PCU47 is not null "
   '¶Ç¤JR®É§ä¥X¬ÛÃöªºY
   strSql = strSql & " union  SELECT FA01||FA02||Decode(FA02,'0','','¡¯'),NVL(FA04,Decode(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),NA03,FA69,FA29 " & _
                                                    "From Fagent, PotCustomer, Nation " & _
                                                "WHERE NA01(+)=FA10 " & _
                                                     "AND PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                     "AND FA01>=(substr(PCU47,1,6)||'00') AND FA01<=(substr(PCU47,1,6)||'zz') " & _
                                                     "AND PCU47 is not null "
   '§ä¥XRªºÃö«Y¥ø·~
   strSql = strSql & " union  SELECT PCU01||PCU02||Decode(PCU02,'0','','¡¯'),NVL(PCU08,Decode(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)),NA03,PCU39,PCU40 " & _
                                                    "From PotCustomer, Nation " & _
                                               "WHERE NA01(+)=PCU09 " & _
                                                    "AND PCU47>='" & Left(StrToGrid, 6) & "00' AND PCU47<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND PCU47 is not null "
   '¶Ç¤JR®É§ä¥X¬ÛÃöªºR1
   strSql = strSql & " union  SELECT POC01||POC02||Decode(POC02,'0','','¡¯'),POC03,NA03,POC14,POC15 " & _
                                                    "From PotCustomer1, Nation, PotCustomer " & _
                                               "WHERE NA01(+)=POC04 " & _
                                                    "AND PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND POC16>=(substr(PCU47,1,6)||'00') AND POC16<=(substr(PCU47,1,6)||'zz') " & _
                                                    "AND PCU47 is not null AND POC16 is not null "
   '2009/06/24 End
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 Then
       adoRecordset.MoveFirst
       Do While adoRecordset.EOF = False
       strSql = "INSERT INTO R100102 values ('"
       If Not IsNull(adoRecordset.Fields(0)) Then
           strSql = strSql + ChgSQL(CheckStr(adoRecordset.Fields(0))) + "','"
       Else
           strSql = strSql + "','"
       End If
       If Not IsNull(adoRecordset.Fields(1)) Then
           strSql = strSql + ChgSQL(CheckStr(adoRecordset.Fields(1))) + "','"
       Else
           strSql = strSql + "','"
       End If
       If Not IsNull(adoRecordset.Fields(2)) Then
           strSql = strSql + ChgSQL(CheckStr(adoRecordset.Fields(2))) + "','" & strUserNum & "')"
       Else
           strSql = strSql + "','" & strUserNum & "')"
       End If
       cnnConnection.Execute strSql
       adoRecordset.MoveNext
       Loop
   Else
       ShowNoData
       Screen.MousePointer = vbDefault
       Exit Sub
   End If
   CheckOC
End Sub

'ÃöÁp¥ø·~
Sub StrMenu1()
    'Dim k As Integer  'Add by Amy 2019/10/05
    
   ''Add by Amy 2013/12/10 +¥Ó½Ð°ê®a/Á`¦¬¤å¸¹,/®×¥ó©Ê½è/¦¬¤å¤éÄæ¦ì
    'Modified by Lydia 2017/12/05 §ï¥Ñ±Ò¥Î¤é±±¨î
    If strSrvDate(1) < °ê¥~³¡ÃöÁp¥ø·~±Ò¥Î¤é Then
        'Modify by Amy 2019/10/05  +4­Ó''->ÃöÁp½s¸¹/¦WºÙ/Ãö«Y/»¡©ú Á×§K¥[Äæ¦ì§xÃø
        'Modified by Lydia 2020/05/07 +'00' as R11401
        strSql = "SELECT '' AS V,R06001 AS ½s¸¹,R06002 AS ¦WºÙ,R06003 AS °êÄy,ST02 AS ´¼Åv¤H­û,CU80 AS ª¬ºA,CU79 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é,'' as ÃöÁp½s¸¹,'' as ÃöÁp¦WºÙ,'' as ÃöÁpÃö«Y,'' as ÃöÁp»¡©ú,'00' as R11401 FROM R100102,CUSTOMER,STAFF where id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='X' AND SUBSTR(R06001,1,8)=CU01(+) AND SUBSTR(R06001,9,1)=CU02(+) AND CU13=ST01(+) "
        'Add By Sindy 98/03/19
        'Modify by Amy 2019/10/05 ­ì:Union All §âAll  ®³±¼ ex:X29973 ¦³¨âµ§(¤@µ§¬°§ó¦W)->¨âµ§¤Ä¿ï->«ö¡uÃö«Y¥ø·~¡v->¤£À³¥X²{¥|µ§
        'Modify by Amy 2020/03/16 ­ì:st02 ,¦]¶}µo¤H­û¥i¯à¦h¤H
        strSql = strSql & "UNION SELECT '' AS V,R06001 AS ½s¸¹,R06002 AS ¦WºÙ,R06003 AS °êÄy,pcu38 AS ´¼Åv¤H­û,PCU39 AS ª¬ºA,PCU40 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é,'' as ÃöÁp½s¸¹,'' as ÃöÁp¦WºÙ,'' as ÃöÁpÃö«Y,'' as ÃöÁp»¡©ú,'00' as R11401 FROM R100102,POTCUSTOMER,staff where id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='R' AND SUBSTR(R06001,1,8)=PCU01 AND SUBSTR(R06001,9,1)=PCU02 and substr(LTrim(PCU38),1,5)=ST01(+) "
        strSql = strSql & "UNION SELECT '' AS V,R06001 AS ½s¸¹,R06002 AS ¦WºÙ,R06003 AS °êÄy,poc13 AS ´¼Åv¤H­û,POC14 AS ª¬ºA,POC15 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é,'' as ÃöÁp½s¸¹,'' as ÃöÁp¦WºÙ,'' as ÃöÁpÃö«Y,'' as ÃöÁp»¡©ú,'00' as R11401 FROM R100102,POTCUSTOMER1,STAFF where id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='R' AND SUBSTR(R06001,1,8)=POC01 AND SUBSTR(R06001,9,1)=POC02 and POC13=ST01(+) "
        'end 2020/03/16
        '98/03/19 End
        strSql = strSql & "UNION SELECT '' AS V,R06001 AS ½s¸¹,R06002 AS ¦WºÙ,R06003 AS °êÄy,' ' AS ´¼Åv¤H­û,FA69 AS ª¬ºA,FA29 AS ³Æµù,' ' as ¥Ó½Ð°ê®a,'' as Á`¦¬¤å¸¹,'' as ®×¥ó©Ê½è,'' as ¦¬¤å¤é,'' as ÃöÁp½s¸¹,'' as ÃöÁp¦WºÙ,'' as ÃöÁpÃö«Y,'' as ÃöÁp»¡©ú,'00' as R11401 FROM R100102,FAGENT where id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='Y' AND SUBSTR(R06001,1,8)=FA01(+) AND SUBSTR(R06001,9,1)=FA02(+) "
        'strSql = strSql & "ORDER BY ½s¸¹" 'Remove by Amy 2019/10/05 +¬¡¤Æ«È¤á
   Else
        'Added by Lydia 2017/02/14 §ìÃöÁp¥ø·~§ï¦¨¼Ò²Õ,¼È¦sR100114_1
        'Modified by Lydia 2020/05/07 +R11401
        strSql = "SELECT '' AS V,R11402 AS ½s¸¹,R11403 AS ¦WºÙ,NVL(NA03,R11405) AS °êÄy ,ST02 AS ´¼Åv¤H­û,R11407 AS ª¬ºA,R11408 AS ³Æµù,' ' AS ¥Ó½Ð°ê®a,'' AS Á`¦¬¤å¸¹,'' AS ®×¥ó©Ê½è,'' AS ¦¬¤å¤é," & _
                 "R11409 AS ÃöÁp½s¸¹,DECODE(SUBSTR(R11409,1,1)," & _
                 "'X',DECODE(SIGN(INSTR('000,001,002,003,004,005,006,007,008,009,013,020',C1.CU10)),0,DECODE(C1.CU05,NULL,NVL(C1.CU04,C1.CU06),C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90),NVL(C1.CU04,DECODE(C1.CU05,NULL,C1.CU06,C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)))," & _
                 "'Y',DECODE(SIGN(INSTR('000,001,002,003,004,005,006,007,008,009,013,020',F1.FA10)),0,DECODE(F1.FA05,NULL,NVL(F1.FA04,F1.FA06),F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65),NVL(F1.FA04,DECODE(F1.FA05,NULL,F1.FA06,F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65))) " & _
                 ",R11409) AS ÃöÁp¦WºÙ," & _
                 "R11410 AS ÃöÁpÃö«Y, R11411 AS ÃöÁp»¡©ú,R11401 FROM R100114_1,STAFF,NATION,CUSTOMER C1,FAGENT F1 " & _
                 "WHERE ID='" & strUserNum & "' AND FORMID='" & UCase(Me.Name) & "' AND R11406=ST01(+) AND R11405=NA01(+) " & _
                 "AND SUBSTR(R11409,1,8)=C1.CU01(+) AND '0'=C1.CU02(+) AND SUBSTR(R11409,1,8)=F1.FA01(+) AND '0'=F1.FA02(+) "
        'strSql = strSql & "ORDER BY R11401,R11402,R11409 " 'Remove by Amy 2019/10/05 +¬¡¤Æ«È¤á
        'end 2017/02/14
   End If
   'end 2020/03/16
   'end 2017/12/05
   
   'Added by Amy 2019/10/05 +¬¡¤Æ«È¤á
   'Modified by Lydia 2020/05/07 ­«·s¾ã²zSQL
   'strSql = "Select X.*,Decode(Ocu01,null, '',NVL(Ocu03,0)) as OCU03 from (" & strSql & ") X, OldCustomer Where substr(½s¸¹,1,8)= ocu01(+) "
   'Modified by Lydia 2023/08/23 §ó¦WOCU03=>«Ý¬¡¤Æ«È¤á; ¼W¥[ORGNÄæ¦ì
   strSql = "Select X.V, X.½s¸¹, X.¦WºÙ, X.°êÄy, X.´¼Åv¤H­û, X.ª¬ºA, X.³Æµù, X.¥Ó½Ð°ê®a, X.Á`¦¬¤å¸¹, X.®×¥ó©Ê½è, X.¦¬¤å¤é, X.ÃöÁp½s¸¹, X.ÃöÁp¦WºÙ, X.ÃöÁpÃö«Y, X.ÃöÁp»¡©ú, " & _
               "'' as ORGN, Decode(Ocu01,null, '',NVL(Ocu03,0)) as «Ý¬¡¤Æ«È¤á from (" & strSql & ") X, OldCustomer Where substr(½s¸¹,1,8)= ocu01(+) "
   If strSrvDate(1) < °ê¥~³¡ÃöÁp¥ø·~±Ò¥Î¤é Then
        strSql = strSql & " ORDER BY ½s¸¹"
   Else
        'Modified by Lydia 2020/05/07 ­«·s¾ã²zSQL
        'strSql = strSql & " ORDER BY R11401,R11402,R11409 "
        strSql = strSql & " ORDER BY R11401, ½s¸¹, ÃöÁp½s¸¹"
   End If
   'end 2019/10/05
   
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 Then
       Set grdDataList.Recordset = adoRecordset
       'Modify by Amy 2023/08/24 ­ìµ{¦¡·h¦ÜSetDataListWidth
        SetDataListWidth (True)
   End If
   CheckOC
  
   'Add by Amy 2019/10/05 +©Ò¦³ÃC¦âÅã¥Ü
   grdDataList.Visible = False
   'Modify by Amy 2023/03/08 Äæ¦ì§ï°ÊºA
   If grdDataList.Rows > 0 Then
        For i = 1 To grdDataList.Rows - 1
            grdDataList.row = i
            grdDataList.col = 1
            grdDataList.CellForeColor = &H0   '¦r¶Â¦â ex:¬d»ö¤j·|¾ã­ÓÅÜ¶Â
            'Modify by Amy 2023/08/24 ÅÜ¦â§ï¦@¥Î¨ç¼Æ
            'Modify by Amy 2023/09/26 ¨Ìª¬ºA§ó·s´¼Åv¤H­û§ï¬°¦@¥Î¨ç¼Æ
            Call UpdQuerySales(Me.Name, grdDataList, strField)
            'end 2023/09/26
            Call SetMSGridColorQCus(0, Me.Name, grdDataList, strField, IIf(Check3.Value = vbChecked, True, False))
            'end 2023/08/24
        Next i
   End If
   
   '­Y¥u¦³¤@µ§¸ê®Æ , «hª½±µ³]©w¬°ÂI¿ï¦¹µ§¸ê®Æ
   'Modify by Amy 2023/08/24 ­ìµ{¦¡§ï¦¨¦@¥ÎSetGridOneData,Á×§K¦³¨S§ï¨ìªº
   cmdOK(8).BackColor = &H8000000F
   Call SetGridOneData
   'end 2023/08/24
   grdDataList.Visible = True
   'end 2019/10/05
End Sub

Private Sub GrdDataList_Click()
   Dim strCopyTxt As String ' Add by Amy 2014/04/25 ½Æ»s½s¸¹¤å¦r
   
   grdDataList.row = grdDataList.MouseRow
   
   'Modify by Amy 2014/04/25 +¿ï¨ì½s¸¹Äæ=½Æ»s
   'Modify by Amy 2023/03/08 Äæ¦ì§ïÅÜ°Ê
   grdDataList.col = grdDataList.MouseCol
   If grdDataList.col = 1 Then
        grdDataList.CellForeColor = &H0 '¶Â¦â
        'Modify by Amy 2020/09/04 ¤£¤p¤ß«ö¨ìÄæ¦ì¦WºÙ¤]·|copy
        If grdDataList.row > 0 Then
            strCopyTxt = grdDataList.TextMatrix(grdDataList.row, grdDataList.col)
        End If
        If strCopyTxt <> "" Then
            '½Æ»s½s¸¹¦Ü°Å¶KÃ¯
            Clipboard.Clear 'Added by Lydia 2021/12/20 ¹w³]²M°£°Å¶KÃ¯; µo²{Clipboard.SetText«e¥¼²M°£°Å¶KÃ¯¡ACtrl+V¶K¨ìForm2.0¤¸¥ó·|±a¤J½Æ»s¤§«eªº¤W¤@µ§ªº½Æ»s¤º®e
            Clipboard.SetText strCopyTxt
            grdDataList.CellBackColor = QBColor(7)
            MsgBox "½s¸¹¤w½Æ»s", , MsgText(21)
        
            '³]¦^­ì¥»ÃC¦â
            'Modify by Amy 2023/08/24 §ï¼g¦Ü¦@¥Î¨ç¼Æ
'            'Add by Amy 2019/09/17 ¬¡¤Æ«È¤áÅã¥Ü¾ã¦C¶À,­Y¦³§b±b½s¸¹©³¬°¬õ¨ä¥LÄæ¬°¶À
'            If grdDataList.TextMatrix(grdDataList.row, GetValue("«Ý¬¡¤Æ«È¤á")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("½s¸¹")), 1) <> "¡¯" Then
'                '§b±b
'                If Right(grdDataList.TextMatrix(grdDataList.row, GetValue("½s¸¹")), 1) = "$" Then
'                    grdDataList.CellBackColor = &HFF& '¬õ¦â
'                '¬¡¤Æ«È¤á
'                Else
'                    grdDataList.CellBackColor = vbYellow
'                End If
'            'Modify by Amy 2019/08/28 +«È¤áª¬ºA¬° ¾E²¾¤£©ú/¸Ñ´²/¼o¤î/ºM¾P/°±·~/¦º¤` Åã¥Ü¶Â©³
'            'Modify by Amy 2022/05/23 ®³±¼ ¾E²¾¤£©ú ¤Î °±·~
'            ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'              And (grdDataList.TextMatrix(grdDataList.row, GetValue("ª¬ºA")) = "¸Ñ´²" Or grdDataList.TextMatrix(grdDataList.row, GetValue("ª¬ºA")) = "¼o¤î" _
'                  Or grdDataList.TextMatrix(grdDataList.row, GetValue("ª¬ºA")) = "ºM¾P" Or grdDataList.TextMatrix(grdDataList.row, GetValue("ª¬ºA")) = "¦º¤`") Then
'                grdDataList.CellBackColor = &H0 '¶Â¦â
'                grdDataList.CellForeColor = &HFF00FF '¯»¬õ¦â
'            ElseIf Right(grdDataList.Text, 1) = "¡ò" Or grdDataList.TextMatrix(grdDataList.row, GetValue("ª¬ºA")) = "¹ï³y" Then
'                grdDataList.CellBackColor = &H8080FF
'            Else
'                grdDataList.CellBackColor = QBColor(15)
'            End If
            Call SetMSGridColorQCus(2, Me.Name, grdDataList, strField, IIf(Check3.Value = vbChecked, True, False))
        End If
        Exit Sub
   End If
   'end 2014/04/25
   
   grdDataList.Visible = False
   grdDataList.col = 0
   If grdDataList.row <> 0 Then
       If grdDataList.Text = "V" Then
            grdDataList.Text = ""
            'Modify by Amy 2023/08/24 §ï¼g¦Ü¦@¥Î¨ç¼Æ
'            'Add By Sindy 2012/3/21
'            grdDataList.col = 1
'            'Add by Amy 2019/09/17 ¬¡¤Æ«È¤áÅã¥Ü¾ã¦C¶À,­Y¦³§b±b½s¸¹©³¬°¬õ¨ä¥LÄæ¬°¶À
'            If grdDataList.TextMatrix(grdDataList.row, GetValue("«Ý¬¡¤Æ«È¤á")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("½s¸¹")), 1) <> "¡¯" Then
'                 For i = 0 To grdDataList.Cols - 1
'                    '§b±b
'                    If Right(grdDataList.Text, 1) = "$" And i = 1 Then
'                        grdDataList.CellBackColor = &HFF& '¬õ¦â
'                    '¬¡¤Æ«È¤á
'                    Else
'                        grdDataList.col = i
'                        grdDataList.CellBackColor = vbYellow
'                    End If
'                Next
'            'Modify by Amy 2019/08/28 +«È¤áª¬ºA¬° ¾E²¾¤£©ú/¸Ñ´²/¼o¤î/ºM¾P/°±·~/¦º¤` Åã¥Ü¶Â©³
'            'Modify by Amy 2022/05/23 ®³±¼ ¾E²¾¤£©ú ¤Î °±·~
'            ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'              And (grdDataList.TextMatrix(grdDataList.row, GetValue("ª¬ºA")) = "¸Ñ´²" Or grdDataList.TextMatrix(grdDataList.row, GetValue("ª¬ºA")) = "¼o¤î" _
'                  Or grdDataList.TextMatrix(grdDataList.row, GetValue("ª¬ºA")) = "ºM¾P" Or grdDataList.TextMatrix(grdDataList.row, GetValue("ª¬ºA")) = "¦º¤`") Then
'                For i = 0 To grdDataList.Cols - 1
'                    grdDataList.col = i
'                    grdDataList.CellBackColor = &H0 '¶Â¦â
'                    grdDataList.CellForeColor = &HFF00FF '¯»¬õ¦â
'                Next i
'            ElseIf Right(grdDataList.Text, 1) = "¡ò" Or grdDataList.TextMatrix(grdDataList.row, GetValue("ª¬ºA")) = "¹ï³y" Then
'               For i = 0 To grdDataList.Cols - 1
'                  grdDataList.col = i
'                  grdDataList.CellBackColor = &H8080FF
'               Next i
'            Else
'            '2012/3/21 End
'               For i = 0 To grdDataList.Cols - 1
'                  If i <> 1 Then
'                     grdDataList.col = i
'                     grdDataList.CellBackColor = QBColor(15)
'                  End If
'               Next i
'            End If
            Call SetMSGridColorQCus(0, Me.Name, grdDataList, strField, IIf(Check3.Value = vbChecked, True, False))
       '¤Ä¿ï
       Else
            grdDataList.Text = "V"
            'Modify by Amy 2023/08/24 §ï¼g¦Ü¦@¥Î¨ç¼Æ
'            For i = 0 To grdDataList.Cols - 1
'               'Modify By Sindy 2012/3/21 old:If i <> 1 Then
'               'Mofify by Amy 2013/12/10 +§PÂ_¹ï³y
'               If i <> 1 And (i = 2 And Right(grdDataList.TextMatrix(grdDataList.MouseRow, GetValue("½s¸¹")), 1) = "¡ò") = False Then
'                   grdDataList.col = i
'                   grdDataList.CellBackColor = &HFFC0C0
'               End If
'            Next i
            Call SetMSGridColorQCus(1, Me.Name, grdDataList, strField, IIf(Check3.Value = vbChecked, True, False))
       End If
       'Add by Amy 2020/10/15 ¤Ä¿ï®É§PÂ_¦³©¹¨Ó°O¿ý,©¹¨Ó°O¿ý¶sÅÜ¦â
       'Modify by Amy 2023/08/24 bug-Ápµ¸¤H¤]·|¦³©¹¨Ó°O¿ý,¬G®³±¼½s¸¹¥u¨ú8½X
       strExc(10) = grdDataList.TextMatrix(grdDataList.row, GetValue("½s¸¹"))
       If Left(strExc(10), 1) = "X" Or Left(strExc(10), 1) = "Y" Or Left(strExc(10), 1) = "R" Or Left(strExc(10), 2) = "¥­¥x" Then
         Call ChkContactRecordBT(grdDataList.TextMatrix(grdDataList.row, GetValue("V")), strExc(10))
       End If
   End If
   grdDataList.Visible = True
End Sub

'Add by Amy 2020/06/16 +±Æ§Ç
Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If grdDataList.MouseCol < 0 Or grdDataList.MouseRow < 0 Then Exit Sub
    
    grdDataList.col = grdDataList.MouseCol
    grdDataList.row = grdDataList.MouseRow
    If grdDataList.col = 2 Then grdDataList.col = 15 'Modify by Amy 2022/08/19 ¦WºÙ¥HOrgN±Æ
    If Me.grdDataList.row < 1 And Me.grdDataList.Text <> "V" Then
        If m_blnColOrderAsc = True Then
            Me.grdDataList.Sort = 5 '¦r¦êª@¾­
            m_blnColOrderAsc = False
        Else
            Me.grdDataList.Sort = 6 '¦r¦ê­°¾­
            m_blnColOrderAsc = True
        End If
    End If
End Sub

'add by nickc 2007/06/13
Private Sub Option1_Click(Index As Integer)
   'If Index = 1 Then
   '    CloseIme
   '    Text2.SetFocus
   'Else
   '    OpenIme
   '    Text2.SetFocus
   'End If
   'Modify By Sindy 2010/02/25
   Call Text2_GotFocus
End Sub

Private Sub Option2_Click(Index As Integer)
   Select Case Index
      Case 0
           If Option2(0).Value = True Then
              Option2(1).Value = False
              'add by nickc 2007/10/24
              Option2(2).Value = False
              'add by nickc 2008/05/02
              Option2(4).Value = False
              
              Option1(0).Enabled = False
              Option1(1).Enabled = False
              Option1(2).Enabled = False
              Option3(0).Enabled = False
              Option3(1).Enabled = False
           End If
      Case 1
           If Option2(1).Value = True Then
              Option1(0).Enabled = True
              'add by nickc 2007/10/24
              Option2(2).Value = False
              'add by nickc 2008/05/02
              Option2(4).Value = False
              
              Option1(0).Value = True
              Option1(1).Enabled = True
              Option1(2).Enabled = True
              Option2(0).Value = False
              Option3(0).Enabled = True
              Option3(1).Enabled = True
              'Modify by Amy 2014/04/30 ¥Ñ¬d¥»©Ò«È¤á¿ï¶µ¶i¤J ¹w³] ¬d¦r­º ¤£¬d¹ï³y
              If IsSearchNew = False Then
                    Option3(0).Value = True
                    Check2.Value = 0
              Else
                    Option3(0).Value = False
                    Check2.Value = 1
                End If
                'Option3(1).Value = True    '2012/3/28 ADD BY SONIA
                'end 2014/04/30
           End If
      'add by nickc 2007/10/24
      Case 2
           If Option2(2).Value = True Then
              Option2(0).Value = False
              Option2(1).Value = False
              'add by nickc 2008/05/02
              Option2(4).Value = False
              
              Option1(0).Enabled = False
              Option1(1).Enabled = False
              Option1(2).Enabled = False
              Option3(0).Enabled = False
              Option3(1).Enabled = False
           End If
           
      'add by Toni 2008/12/03
      Case 3
         If Option2(3).Value = True Then
              Option2(0).Value = False
              Option2(1).Value = False
              Option2(2).Value = False
              Option2(4).Value = False
              
              Option1(0).Enabled = False
              Option1(1).Enabled = False
              Option1(2).Enabled = False
              Option3(0).Enabled = False
              Option3(1).Enabled = False
         End If
      
      'add by nickc 2008/05/02
      Case 4
           If Option2(4).Value = True Then
              Option2(0).Value = False
              Option2(1).Value = False
              Option2(2).Value = False
              Option1(0).Enabled = False
              Option1(1).Enabled = False
              Option1(2).Enabled = False
              Option3(0).Enabled = False
              Option3(1).Enabled = False
              Text11_GotFocus
           End If
      Case Else
   End Select
End Sub

Private Sub Text1_GotFocus()
   'Me.Option2(0).Value = True
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1)
   'edit by nickc 2007/06/06
'   'Add by Morgan 2006/4/11 §PÂ_§@·~¨t²Î95,98¤~¤Á
'   If pub_OS = 1 Then
'      Text2.IMEMode = 2
'      Debug.Print Me.Text2.IMEMode & ":c1-->" & Now
'   End If
   CloseIme
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option2(0).Value = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'add by Toni 2008/12/03
Private Sub Text10_GotFocus()
   Me.Option2(3).Value = True
   Text10.SelStart = 0
   Text10.SelLength = Len(Text10)
   CloseIme
End Sub

Private Sub Text10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option2(3).Value = True
End Sub

'add by nickc 2008/05/02
Private Sub Text11_GotFocus()
   Me.Option2(4).Value = True
   Text11.SelStart = 0
   Text11.SelLength = Len(Text11)
   CloseIme
End Sub

'Add by Amy 2013/09/27
Private Sub Text11_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

'add by nickc 2008/05/02
Private Sub Text11_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option2(4).Value = True
End Sub

Private Sub Text2_GotFocus()
   Me.Option2(1).Value = True
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2)
   'Add by Amy 2013/12/10
   If Left(Pub_StrUserSt03, 1) = "F" Then
        CloseIme
   Else
        OpenIme
   End If
   'end 2013/12/10
'   If pub_OS = 1 Then
      'Modify by Amy 2013/12/04 Mark±¼
'      '­^¤å
'      If Option1(1).Value = True Then
'         'edit by nickc 2007/06/06
'         'Me.Text2.IMEMode = 2
'         CloseIme
'      Else
'         'edit by nickc 2007/06/06
'         'Me.Text2.IMEMode = 1
'         OpenIme
'      End If
'   End If

End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option2(1).Value = True
End Sub

'Add by Morgan 2006/6/12
Private Sub Text2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   'If pub_OS = 1 Then
      'Modify by Amy 2013/12/04 Mark±¼
'      '­^¤å
'      If Option1(1).Value = True Then
'         'edit by nickc 2007/06/06 ¤Á´«¿é¤Jªk§ï¥ÎAPI
'         'Me.Text2.IMEMode = 2
'         CloseIme
'      Else
'         'edit by nickc 2007/06/06 ¤Á´«¿é¤Jªk§ï¥ÎAPI
'         'Me.Text2.IMEMode = 1
'         OpenIme
'      End If
   'End If
End Sub

Private Sub Text3_GotFocus()
   Text3.SelStart = 0
   Text3.SelLength = Len(Text3)
   CloseIme
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   Text4.SelStart = 0
   Text4.SelLength = Len(Text4)
   CloseIme
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_LostFocus()
   If PUB_CheckKeyInDate(Me.Text4) = -1 Then
      Me.Text4.SetFocus
      Text4_GotFocus
      Exit Sub
   End If
End Sub

Private Sub Text5_GotFocus()
   Text5.SelStart = 0
   Text5.SelLength = Len(Text5)
   CloseIme
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_LostFocus()
   If PUB_CheckKeyInDate(Me.Text5) = -1 Then
      Me.Text5.SetFocus
      Text5_GotFocus
      Exit Sub
   End If
   If Not nickChgRan(Text4, Text5, "¦¬¤å¤é´Á") Then
      Text4.SetFocus
      Text4_GotFocus
      Exit Sub
   End If
End Sub

Private Sub Text6_GotFocus()
   Text6.SelStart = 0
   Text6.SelLength = Len(Text6)
   CloseIme
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_GotFocus()
   Text7.SelStart = 0
   Text7.SelLength = Len(Text7)
   CloseIme
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_LostFocus()
   If Not nickChgRan(Text6, Text7, "®×¥ó©Ê½è") Then
      Text6.SetFocus
      Text6_GotFocus
   End If
End Sub

Private Sub Text8_GotFocus()
      Text8.SelStart = 0
      Text8.SelLength = Len(Text8)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text8_LostFocus()
Dim s
   If InStr(1, "nN ", Text8) = 0 Then
       s = MsgBox("¶È­­¿é¤J N ©ÎªÅ¥Õ", , "USER ¿é¤J¿ù»~")
       Text8.SetFocus
       Text8.SelStart = 0
       Text8.SelLength = Len(Text8)
   End If
End Sub

'add by nickc 2007/10/24
Private Sub Text9_GotFocus()
   Me.Option2(2).Value = True
   Text9.SelStart = 0
   Text9.SelLength = Len(Text9)
   OpenIme
End Sub

'add by nickc 2007/10/24
Private Sub Text9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option2(2).Value = True
End Sub

Private Sub txtCountry_GotFocus(Index As Integer)
   TextInverse txtCountry(Index)
   CloseIme
End Sub

Private Sub txtCountry_LostFocus(Index As Integer)
   If Index = 1 Then
      If Not nickChgRan(txtCountry(0), txtCountry(1), "¥Ó½Ð°ê®a") Then
         txtCountry(0).SetFocus
         txtCountry_GotFocus 0
      End If
   End If
End Sub

'Mark by Amy 2023/09/20 §ï¦¨¦@¥Î¨ç¼Æ
'Add by Amy 2014/02/21 (PrintDataA4 §R°£¤£¥Î'Add by Amy 2013/11/06)
Private Sub PrintDataA4_Temp()
'    Dim rsPrint As New ADODB.Recordset
'    Dim strPrint As String
'    Dim ii As Integer, jj As Integer
'On Error GoTo Checking
'    intCounter = 1: intRecord = 1: intPage = 1
'
'    Screen.MousePointer = vbHourglass
'    Printer.PaperSize = PUB_GetPaperSize(9) '³]©w¯È±i A4
'    Printer.Orientation = 1 'ª½¦L
'    PrintHeadA4
'
'    Printer.FontBold = False
'    'Modify by Amy 2020/09/08 ID+ªí³æ
'    strPrint = "Select R021001,R021002,R021003,Decode(R021004,'1','¹ï³y','¨ä¥L¬ÛÃö¤H'),R021006,R021007,Nvl(To_Char(R021008-19110000),'') " & _
'                 "From R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "' Order by R021002,R021001"
'    intI = 1
'    Set rsPrint = ClsLawReadRstMsg(intI, strPrint)
'    If intI = 1 Then
'        rsPrint.MoveFirst
'        For ii = 0 To rsPrint.RecordCount - 1
'            If intRecord > 45 Then
'                intPage = intPage + 1
'                intRecord = 1
'                Printer.NewPage
'                intCounter = 1
'                PrintHeadA4
'                Printer.FontBold = False
'            End If
'            For jj = 0 To rsPrint.Fields.Count - 1
'                If jj = rsPrint.Fields.Count - 1 Then
'                    Printer.CurrentX = PLeft(jj + 1) - 300 - Printer.TextWidth(rsPrint.Fields(jj).Value) '³Ì¥kÃä
'                Else
'                    Printer.CurrentX = PLeft(jj)
'                End If
'                Printer.CurrentY = 300 + intCounter * 300
'
'                Select Case jj
'                    Case 0 '¥»©Ò®×¸¹
'                        Printer.Print Pub_RplStr(rsPrint.Fields(jj).Value)
'                    Case 1 '¦WºÙ
'                        Printer.Print StrToStr(rsPrint.Fields(jj).Value, 10)
'                    Case 2, 3, 4 '´¼Åv¤H­û,ª¬ºA,Á`¦¬¤å¸¹
'                        Printer.Print rsPrint.Fields(jj).Value
'                    Case 5 '®×¥ó©Ê½è
'                        Printer.Print StrToStr(rsPrint.Fields(jj).Value, 6)
'                    Case 6  '¦¬¤å¤é
'                        Printer.Print ChangeTStringToTDateString(rsPrint.Fields(jj).Value)
'                    Case Else
'                End Select
'            Next jj
'            intCounter = intCounter + 1
'            intRecord = intRecord + 1
'            rsPrint.MoveNext
'        Next ii
'    End If
'    Printer.EndDoc
'    Screen.MousePointer = vbDefault
'
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'   Screen.MousePointer = vbDefault
End Sub
'end 2014/02/21

Private Sub PrintHeadA4()
'   If intPage = 1 Then
'        GetPleft
'        strTp(0) = "¥H¥Ó½Ð¤H¬d¸ß"
'        strTp(1) = ""
'
'        If Option3(0).Value = True Then
'            strTp(1) = strTp(1) & "(¦r­º¤ñ¹ï)"
'        ElseIf Option3(1).Value = True Then
'            strTp(1) = strTp(1) & "(¼Ò½k¤ñ¹ï)"
'        End If
'   End If
'   strTp(2) = "¦WºÙ¡G" & strTp(3) & Space(6) & strTp(1)
'
'   Printer.FontSize = 17
'   Printer.FontBold = True
'   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strTp(0)) / 2)
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print strTp(0)
'
'   Printer.FontSize = 12
'   intCounter = intCounter + 2
'   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strTp(2)) / 2)
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print strTp(2)
'   'Printer.Line (Printer.ScaleWidth / 2 - ((Printer.TextWidth(strTp(2)) - Printer.TextWidth("¦WºÙ¡G")) / 2) + 300, Printer.CurrentY + 30)-(Printer.ScaleWidth / 2 + Printer.TextWidth(strTp(2)) / 2, Printer.CurrentY + 30)
'
'   intCounter = intCounter + 1
'   Printer.CurrentX = 0
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print "¾Þ§@¤H­û¡G" & StaffQuery(strUserNum)
'   Printer.CurrentX = 8800
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print "¬d¸ß¤é´Á¡G" & CFDate(ACDate(ServerDate))
''   intCounter = intCounter + 1
''   Printer.CurrentX = 12000
''   Printer.CurrentY = 300 + intCounter * 300
''   Printer.Print "­¶¦¸: " & intPage
'    intCounter = intCounter + 1
'    For kk = 1 To UBound(PLeft)
'        Printer.CurrentX = PLeft(kk - 1) + (PLeft(kk) - PLeft(kk - 1) - Printer.TextWidth(ColName(kk)) - 100) / 2
'        Printer.CurrentY = 300 + intCounter * 300
'        Printer.Print ColName(kk)
'        Printer.Line (PLeft(kk - 1), Printer.CurrentY)-(PLeft(kk) - 100, Printer.CurrentY)
'    Next kk
'    intCounter = intCounter + 1
End Sub

Private Sub GetPleft()
'   ReDim PLeft(0 To 7)
'   ReDim ColName(1 To 7)
'   PLeft(0) = 100
'   PLeft(1) = PLeft(0) + 2000: ColName(1) = "¥»©Ò®×¸¹"
'   PLeft(2) = PLeft(1) + 2700: ColName(2) = "    ¦W       ºÙ    "
'   PLeft(3) = PLeft(2) + 1200: ColName(3) = "´¼Åv¤H­û"
'   PLeft(4) = PLeft(3) + 1500: ColName(4) = " ª¬  ºA "
'   PLeft(5) = PLeft(4) + 1300: ColName(5) = "Á`¦¬¤å¸¹"
'   PLeft(6) = PLeft(5) + 1800: ColName(6) = "®×¥ó©Ê½è"
'   PLeft(7) = PLeft(6) + 1200: ColName(7) = "¦¬¤å¤é"
End Sub
'end 2013/11/06
'end 2023/09/20 ¤£¨Ï¥Î

'Add by Amy 2020/10/15 ¤Ä¿ï®É§PÂ_¦³©¹¨Ó°O¿ý,©¹¨Ó°O¿ý¶sÅÜ¦â
Private Sub ChkContactRecordBT(ByVal stChk As String, ByVal stKey As String)
    'Memo by Amy 2023/09/27  ­ì2023/08/24 ±N«ö¶sÂê¦í,¦³¸ê®Æ¤~¥i«ö,User «ö¦¹¶s·s¼W,¬G¤£Âê
    cmdOK(8).BackColor = &H8000000F
    If stChk = "V" And PUB_ChkContactRecord(stKey) = True Then
        cmdOK(8).BackColor = vbYellow
    End If
End Sub

'Add by Amy 2023/08/24 ¬d¸ß¥u¦³¤@µ§¸ê®ÆGridÃC¦â³]©w
Private Sub SetGridOneData()
     grdDataList.Visible = False
     With Me.grdDataList
        If .Rows = 2 Then
            .row = 1
            .col = 1
            If .Text <> "" Then
                .row = 1
                .col = 0
                .Text = "V"
                'ÅÜ¹L¦â¤´»Ý­n¦A¶],¦]¬°¥u¦³¤@µ§¿ï¨ú®É­nÅÜÂÅ¦â
                Call SetMSGridColorQCus(1, Me.Name, grdDataList, strField, IIf(Check3.Value = vbChecked, True, False))
                '¤Ä¿ï®É§PÂ_¦³©¹¨Ó°O¿ý,©¹¨Ó°O¿ý¶sÅÜ¦â
                strExc(10) = grdDataList.TextMatrix(grdDataList.row, GetValue("½s¸¹"))
                If Left(strExc(10), 1) = "X" Or Left(strExc(10), 1) = "Y" Or Left(strExc(10), 1) = "R" Or Left(strExc(10), 2) = "¥­¥x" Then
                  Call ChkContactRecordBT(grdDataList.TextMatrix(grdDataList.row, GetValue("V")), strExc(10))
                End If
            End If
        End If
   End With
   grdDataList.Visible = True
End Sub
              


