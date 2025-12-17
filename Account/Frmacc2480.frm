VERSION 5.00
Begin VB.Form Frmacc2480 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "FC½Ð´Ú³æ"
   ClientHeight    =   4320
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5640
   Begin VB.CommandButton CmdHelp 
      Caption         =   "PDF reDirect »¡©ú"
      Height          =   495
      Left            =   4170
      TabIndex        =   20
      Top             =   2970
      Width           =   1185
   End
   Begin VB.CheckBox Check2 
      Caption         =   "¥u²£¥ÍLEDES¹q¤l±b³æ"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3150
      TabIndex        =   19
      Top             =   1620
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   315
      Left            =   5130
      ScaleHeight     =   264
      ScaleWidth      =   288
      TabIndex        =   15
      Top             =   90
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txtOutMode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1845
      MaxLength       =   1
      TabIndex        =   13
      Text            =   "1"
      Top             =   2970
      Width           =   345
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3150
      TabIndex        =   1
      Top             =   390
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1230
      TabIndex        =   0
      Top             =   390
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "¦C¦L(&P)"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   135
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   2
      Top             =   780
      Width           =   5370
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "¼Ð·¢Åé"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1230
      TabIndex        =   3
      Top             =   1200
      Width           =   4300
   End
   Begin VB.TextBox txtCopy 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1224
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "3"
      Top             =   1620
      Width           =   705
   End
   Begin VB.TextBox txtAdd 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1380
      MaxLength       =   1
      TabIndex        =   5
      Top             =   2070
      Width           =   705
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "¼Ð·¢Åé"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1860
      TabIndex        =   6
      Top             =   2520
      Width           =   3630
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "ª`·N : ¦C¦L¤¤½Ð¤Å¾Þ§@¨t²Î©ÎWord¡I"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   810
      TabIndex        =   18
      Top             =   60
      Width           =   4125
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '³z©ú
      Caption         =   "½Ð´Ú¹ï¶H: X5208400, X4831001, X4831000, Y4830906, Y4830907, X5349500, X3224200, Y4725001, Y2245700  ¥Ó½Ð¤H: X5509400, X7286900"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   450
      TabIndex        =   17
      Top             =   3600
      Width           =   5055
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "¤U¦C¹ï¶H¦C¦L®æ¦¡¯S§O¡G"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   16
      Top             =   3360
      Width           =   2310
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "½Ð´Ú³æ¿é¥X¤è¦¡¡G       (1:¦Lªí¾÷  2:¹q¤lÀÉ)"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   14
      Top             =   3000
      Width           =   3960
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '³z©ú
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2910
      TabIndex        =   12
      Top             =   390
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '³z©ú
      Caption         =   "D/N No.¡G"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   11
      Top             =   390
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   45
      Top             =   1230
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '³z©ú
      Caption         =   "¦Lªí¾÷¡G"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '³z©ú
      Caption         =   "¦C¦L¥÷¼Æ¡G            (¥÷)"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   9
      Top             =   1650
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '³z©ú
      Caption         =   "¦C¦L¦a§}±ø¡G            (Y : ¬O)"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   8
      Top             =   2100
      Width           =   2745
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '³z©ú
      Caption         =   "¦a§}±ø¦Lªí¾÷¡G"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   7
      Top             =   2535
      Width           =   1725
   End
End
Attribute VB_Name = "Frmacc2480"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/8/17 ¤é¤å¤w§ï§ìTable
'Memo by Morgan 2022/8/4 µ{¦¡½X¤Ó¦h¡A¥B¥Ø«e¤w¬Ò¥iÂà¦sPDF®æ¦¡¡A¨ú®ø¦C¦L¦¨¹Ï¤ùªºµ{¦¡½X¡A¤]¥i¬Ù¥h¤£¥²­nªººûÅ@¤ÎÀË¬d¡C
'Memo By Sonia 2012/12/4 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo By Sindy 2011/2/17 SQLDate¤wÀË¬d
'2010/12/1 memo by sonia ­û¤u½s¸¹Äæ¤w­×§ï
'Memo by Morgan 2010/7/27 ¤é´ÁÄæ¤w­×§ï
'Memo by Morgan 2015/2/25
'***** ª`·N¡G­Y¸ê®Æ§t¦³¸õ¦æ²Å¸¹¥i¯à·|¾É­P Word ´å¼Ð¦ì¸m²¾°Ê¤£¥¿½T¦Óµo¥Í¿ù»~ *****

Option Explicit

Public adoacc1k0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset

Public m_bBeCalled As Boolean 'Add by Morgan 2008/4/3 ¬O§_³Q©I¥s
Public m_CallPrevForm As String  'Added by Lydia 2020/01/06 ©I¥s½Ð´Ú³æªºµ{¦¡¦WºÙ
Public m_bEMail As Boolean 'Add by Morgan 2008/4/3 ¬O§_¥HEMail±H°e
Public m_bPaper As Boolean 'Add by Morgan 2009/10/19 ¬O§_EMail¦P®É±H¯È¥»
Public m_SavePath As String 'Add by Morgan 2009/4/13 ¹q¤lÀÉ¦s©ñ¸ô®|
Public m_sEBillingMsg As String 'Add by Morgan 2010/11/9
Public m_bAddDate As Boolean 'Add by Morgan 2011/6/23 ¹q¤lÀÉ¦W¬O§_¥[¤é´Á
Public m_iPageCount As Integer 'Add by Morgan 2011/6/24 ½Ð´Ú³æ­¶¼Æ
Public m_iCopies As Integer 'Add by Morgan 2011/7/8 «ü©w¦C¦L¥÷¼Æ
Public m_bolOneAddr As Boolean 'Added by Morgan 2012/10/3 ¾ã§å½Ð´Ú¥u­n¦C¦L1±i(¸Ë¤@­Ó«H«Ê)--³¯ª÷½¬
Public m_bLedesOnly As Boolean 'Added by Morgan 2018/6/27 ¨ú±oLEDES¤º®e(¤ë±b³æ®æ¦¡¥Î)
Public m_bEditDoc As Boolean '¯S®í½Ð´Ú 'Added by Lydia 2015/04/15 +¾ã§å½Ð´Ú³æcall
Public m_bPAID As Boolean 'Added by Lydia 2020/06/23 ¤w¥I´Ú=>¥[¡uPAID³¹¡v
Public m_strOutErr As String 'Added by Lydia 2020/09/10 §PÂ_PDFÀÉ®×¬O§_¦s¦b
Public m_bolNoPic As Boolean 'Added by Morgan 2023/12/20 ­¶­º§À¤£©ñ¹Ï(©ú²Ó­n¦X¨Ö¨ì¤ë±b³æ­º­¶®É¨Ï¥Î)
Public m_strInvoiceNo As String 'Added by Morgan 2024/1/24 «ü©w½Ð´Ú³æ¸¹

Dim strSql As String
Dim strNo As String
Dim douAmount As Double
Dim strAmount As String
Dim intLength As Integer
Dim intCounter As Integer
Dim strLanguage As String
Private Const intDefault As Integer = 500
Private Const intTop As Integer = 600
Dim strNewPage As String
Dim prnPrint As Printer
Dim strPrinter As String
Dim strCurr As String
Dim strRemark As String
Dim intAddSpaceRow As Integer
'Add By Cheng 2003/02/07
'¥»©Ò®×¸¹
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
'Add By Cheng 2003/02/24
Dim m_DetailTopStart As Double '¦C¦L©ú²Óªº°_©lÂI
'Add By Cheng 2003/03/19
Dim m_strOriExcRate As Double '­ì©l¶×²v
Dim intRow As Integer
'Add By Cheng 2004/04/27
Dim m_strA1K01 As String '½Ð´Ú³æ¸¹
'End
Dim m_TotOffFees As Long 'Add by Morgan 2004/10/4 Total Official Fees
Dim m_A1k03 As String 'Add by Morgan 2006/10/14
Dim m_iPages As Integer
Dim m_strDN As String '¥¿¦b¦C¦Lªº½Ð´Ú³æ¸¹
Dim bolPTUSDCase As Boolean '¬O§_¬° P,T ¦L¬üª÷

'Add by Morgan 2008/4/3
Dim m_b2Printer As Boolean '¬O§_¦L¯È¥»
Dim m_b2Picture As Boolean '¬O§_¦s¹q¤lÀÉ
Dim m_b2Word As Boolean '¬O§_¥HWord¿é¥X(ÂÂªº½Ð´Ú³æ¤´¥HPrinter¤è¦¡¿é¥X) 'Add by Morgan 2011/1/5
Dim m_bMsg As Boolean '¬O§_´£¿ô¦³¦s¹q¤lÀÉ
Dim m_tmp As String '¼È¦s¦C¦L¤º®e
Dim m_strCaseNo As String '¹q¤lÀÉ¥»©Ò®×¸¹
Dim douExtRate As Double '¦r«¬¦ì¸mÁY©ñ¤ñ
Dim m_EFilePath As String '¹q¤lÀÉ¸ô®|
Dim m_A1k27 As String 'Add by Morgan 2012/4/17
Dim m_A1k28 As String 'Add by Morgan 2008/5/29
Dim m_DNRate As Double  '2009/5/18 add by sonia
Dim m_DNCurr As String  '2009/5/18 add by sonia
Dim m_bEBilling As Boolean 'Add by Morgan 2010/11/4
Dim strLedes() As String 'Add by Morgan 2010/11/5
Dim iUpper As Integer 'Add by Morgan 2010/11/5
Dim m_ItemDesc As String '¶µ¥Ø»¡©ú Add by Morgan 2010/11/9
Dim m_ItemHDesc As String '¶µ¥Ø»¡©ú«e­±³¡¥÷ Add by Morgan 2013/10/24
Dim m_ItemTDesc As String '¶µ¥Ø»¡©ú«á­±³¡¥÷ Add by Morgan 2013/10/24
Dim m_ItemXDesc As String '¶µ¥Ø»¡©ú¯S§O Add by Morgan 2013/10/29
Dim m_strFA126 As String 'Add By Sindy 2021/3/3

'Add by Morgan 2010/11/24
'¼È¦s­n¦C¦Lªº¸ê®Æ¥H«K²£¥ÍWord®æ¦¡
Dim m_Head() As String 'ªíÀY:¦C¦L¹ï¶H..., 4Äæ x N¦C
Dim m_Subject() As String '¼ÐÃD:®×¥ó¦WºÙ..., 3Äæ x N¦C
Dim m_iSubject As Integer
Private Type INVITEM
   IDesc As String
   ICode As String
   iCur As String
   iAmt As String
   IDate As String
   IDescHead As String 'Added by Morgan 2013/10/24 ¶µ¥Ø»¡©ú«e­±³¡¥÷(­ì¨Ó»¡©ú)
   IDescTail As String 'Added by Morgan 2013/10/24 ¶µ¥Ø»¡©ú«á­±³¡¥÷(§é¦©...)
   iNo As String 'Added by Morgan 2013/10/29 ½Ð´Ú¶µ¥Ø¥N½X
   IDescX As String 'Added by Morgan 2013/10/29 ¶µ¥Ø»¡©ú¯S§O(¦³³W¶O¤º®e)
   IXAmt As String 'Added by Morgan 2014/8/28 ¯S®í½Ð´Ú³æ½Ð´Úª÷ÃB
'Added by Lydia 2015/04/09 ¤¤¤åª©-¾ã§å½Ð´Ú³æ
   IChiCno As String '®×¸¹
   IChiCna As String '¦WºÙ
   IChiCls As String 'Ãþ§O
   IChiApp As String 'µù¥U¸¹/¥Ó½Ð®×¸¹
   IChiA1k01 As String '½Ð´Ú³æ¸¹
   IChiAmt As String '½Ð´Úª÷ÃB(­ì¹ô§O)
   IChiUAmt As String '½Ð´Úª÷ÃB
   INtAmt As String 'Added by Morgan 2016/8/5 ¥x¹ô½Ð´Úª÷ÃB
   iAmtNoDisc As String 'Added by Morgan 2020/8/6 §é¦©¿úª÷ÃB
End Type

Dim m_Item() As INVITEM '½Ð´Ú¶µ¥Ø
Dim m_iItem As Integer
Dim m_Sum() As String '5Äæ x N¦C(¦X­p,¹ô§O,ª÷ÃB)
Dim m_Footer() As String '2Äæ x N¦C(¦X­p,¹ô§O,ª÷ÃB)

'Added by Morgan 2014/8/21
Dim m_PlusFormNo As String '¯S®í¦C¦L¹ï¶H¥N½X
Dim m_PlusHead() As String 'ªíÀY:¦C¦L¹ï¶H..., 4Äæ x N¦C
Dim m_PlusFooter() As String '2Äæ x N¦C(¦X­p,¹ô§O,ª÷ÃB)
Dim m_PdfDone As Boolean '¬O§_¤wÂàpdf
'end 2014/8/21

Dim m_Title As String
Dim strCust1 As String  '2010/11/26 add by sonia
Dim m_bSpecial1 As Boolean 'Add by Morgan 2010/12/3 ¯S®í½Ð´Ú³æ
Dim m_bDowX As Boolean 'Added by Morgan 2017/8/17 ¬O§_¬° Dow ªº¯S®í®æ¦¡
Dim m_bDowN As Boolean 'Added by Morgan 2020/8/6 Dow §é¦©¤£¥t¦C
Dim m_bSpecial2 As Boolean 'Add by Morgan 2012/6/12 ¯S®í½Ð´Ú³æ2
Dim m_bSpecial3 As Boolean 'Add by Morgan 2012/10/26 ¯S®í½Ð´Ú³æ3
Dim m_bSpecial4 As Boolean 'Add by Morgan 2014/2/24 ¯S®í½Ð´Ú³æ5
Dim m_bSpecial5 As Boolean 'Added by Lydia 2016/03/03 ¯S®í½Ð´Ú³æ(X72869)
Dim m_bSpecialNew1 As Boolean 'Add by Morgan 2014/2/18 ·s®æ¦¡¯S®í½Ð´Ú³æ1
Dim m_bSpecialNew2 As Boolean 'Add by Morgan 2016/8/5 ·s®æ¦¡¯S®í½Ð´Ú³æ2
Dim m_bSpecialNew3 As Boolean 'Add by Morgan 2018/3/22 ·s®æ¦¡¯S®í½Ð´Ú³æ3
Dim m_bSpecialNew4 As Boolean 'Add by Morgan 2018/4/11 ¥N²z¤H Y27696000 Bobst Mex SA±b³æ®æ¦¡
'Added by MOrgan 2024/3/25 (FMP)Y20049000+X47325000/X47325000, (FCP)Y20049000+X47325C10/X47325C12 ¯S®í½Ð´Ú³æ
Dim m_bSpecialNew5 As Boolean
Dim m_EngCP10List As String
'end 2024/3/25
Dim m_dblDiscTot As Double 'Add by Morgan 2010/12/3 §é¦©Á`ÃB
Dim m_dblNoDiscAmtTot As Double 'Add by Morgan 2011/5/24 ¥¼§é¦©½Ð´ÚÁ`ÃB
Dim m_ClearItemDesc As String '¶µ¥Ø»¡©ú(Â²³æª©) Add by Morgan 2010/12/7
Const m_LineH As Integer = 280 'Add by Morgan 2010/12/24 ¦C°ª
'Add by Morgan 2011/1/5
Dim m_bSaveWord As Boolean
Dim m_bPrintWord As Boolean
Dim m_iSpCopies As Integer
Dim m_strA1K02 As String '½Ð´Ú¤é´Á
Dim m_iLedesVer As Integer 'Add by Morgan 2011/2/25 LEDESª©¥» 1=1998B,2=1998BI
Dim m_iCols As Integer 'Add by Moragn 2011/5/13 LEDES Äæ¦ì¼Æ
Dim m_A1k08 As String 'Add by Morgan 2011/4/1
Dim bolUserClick As Boolean 'Add by Morgan 2011/4/26
Dim adoLEDES As ADODB.Recordset 'Added by Morgan 2012/4/24
Dim m_bolAddAddrOK As Boolean 'Added by Morgan 2012/10/3 ¾ã§å½Ð´Ú¥u­n¦C¦L1±i(¸Ë¤@­Ó«H«Ê)--³¯ª÷½¬
'Added by Morgan 2012/10/31
Dim m_b2PDF As Boolean '¬O§_¦C¦LPDF
Dim m_bPrint2Pdf As Boolean '¬O§_¦C¦L¸Ë¸m¬° PDF ¦Lªí¾÷
Dim m_bWord2Pdf As Boolean
'Added by Morgan 2012/12/7
Dim m_iPrintCurrType As Integer '¦C¦L¹ô§O®æ¦¡:1.¯Â¥x¹ô 2.¥x¹ô+¥~¹ô¦X­p 3.¯Â¥~¹ô 4.¥~¹ô+¬üª÷¦X­p
Dim m_bolNewBill As Boolean '¬O§_¥Î·sµ{¦¡
Dim m_DUsdRate As Double 'Add By Sindy 2013/1/10 ½Ð´Ú¹ô§O¹ï¬üª÷¶×²v
'Add By Sindy 2013/1/28
Dim bolIsFMP As Boolean
Dim strFMPFee99RMB As String
Dim strA1L16 As String
Dim strCompDate As Double
'2013/1/28 End

Dim bolFMPCase As Boolean 'Added by Morgan 2013/5/13
Dim bolNewForm As Boolean 'Added by Morgan 2013/10/25 ½Ð´Ú³æ·s®æ¦¡
Dim m_bShowWord As Boolean 'Added by Morgan 2014/1/23 Word¬O§_¥iÁôÂÃ°õ¦æ
Dim m_WordLeft As Long, m_WordTop As Long 'Added by Morgan 2014/6/24

'Added by Lydia 2015/04/09 ¤¤¤åª©-¾ã§å½Ð´Ú³æ
Public m_bolChiDB As Boolean '³]©w¬°¾ã§å½Ð´Ú³æ¦C¦L
Public m_ChiSys As String '¨t²Î§O
Public m_ChiApply As String '¥Ó½Ð¤H (½Ð´Ú¹ï¶H)
Public m_ChiCust As String '«È¤á
Public m_Chi2Word As Boolean '¿é¥XWordÀÉ(¥[«HÀY«H§À)
Dim bol_ChiDB As Boolean '³vµ§§PÂ_¬O§_¥i¾ã§å½Ð´Ú³æ
Dim midStr(0 To 3) As String '¼È¦s¸ê®Æ
Public m_ChiArrNO As String 'Added by Lydia 2015/08/10 ¶Ç-¾ã§å5.¤¤¤å½Ð´Ú³æªº³æ¸¹
Dim m_TM23 As String 'Add by Amy 2017/01/11
Dim m_Activity As String 'Added by Morgan 2019/2/27
Dim m_CaseName As String 'Added by Morgan 2019/5/29
Dim m_OsPrinter  As String 'Added by Lydia 2019/12/27 §@·~¨t²Î¹w³]¦Lªí¾÷; ­ì¥»¦@¥ÎÅÜ¼Æpub_OsPrinter¦]¬°±qFCP¦~¶Oµo¤å->½Ð´Ú¨ç¦C¦L->½Ð´Ú³æ¦C¦L¡A«h¤¤¶¡¹Lµ{OS¹w³]¦Lªí¾÷¦³ÅÜ¤Æ¡A³y¦¨µLªk¦^¨ìµo¤å«eªºOS¹w³]¦Lªí¾÷¡C
Dim m_Att1 As String, m_Att2 As String 'Added by Morgan 2020/3/17 Ápµ¸¤H1(­^),Ápµ¸¤H2(­^)
Dim iPicNo As Integer, iPicNo2 As Integer 'Added by Morgan 2020/3/31
Dim strUniText As String 'Added by Morgan 2022/7/27
Dim strAppNo As String 'Modified by Morgan 2023/7/10 §ï¬°¥þ°ìÅÜ¼Æ
Dim m_CaseNo As String 'Added by Morgan 2023/10/13 «È¤á®×¥ó®×¸¹
Dim m_CustNoList As String 'Added by Morgan 2024/3/25
Dim m_LD16 As String 'Added by Morgan 2025/3/14

Private Sub Check2_Click()
   If Check2.Value = vbChecked Then
      txtOutMode = "1"
   End If
End Sub

'Added by Lydia 2021/02/23
Private Sub cmdHelp_Click()
Dim strMsg As String
   '¹ï©ó Printer ª«¥ó¦Ó¨¥¡A¦h¥÷¦C¦L¥i¥H¦³¨âºØ¼Ò¦¡¡A±N¾ã­Ó¤å¥ó³v¥÷½Æ»s©Î±N¤å¥ó¤¤ªº¨C¤@­¶¶i¦æ½Æ»s¡A¯à§_³v­¶½Æ»s¨ú¨M©ó¦Lªí¾÷ÅX°Êµ{¦¡¡C¹ï©ó¤£¤ä´©³v­¶½Æ»sªº¦Lªí¾÷¡A¥i¥H³]©w Copies = 1¡AµM«á¦bµ{¦¡¤¤¨Ï¥Î°j°é¡A±N¾ã­Ó¤å¥ó¦C¦L¦h¥÷¡C
   '¥Ø«e¦Ò¼{¶]°j°éªº³t«×¡A¤£±Ä¥Î¦h¦¸¦C¦L
   strMsg = "¦C¦L¤º®e¥u¦³¤@­¶¡A¦C¦L¥÷¼ÆCopies¥u¤ä´©¤@¥÷¡F" & vbCrLf & _
                 "¦C¦L¤º®e¦³¦h­¶¡A¦C¦L¥÷¼ÆCopies¤ä´©¤@¥÷¥H¤W¡C"
   MsgBox strMsg, vbInformation + vbOKOnly, "PDF reDirect »¡©ú"

End Sub

'Modify By Cheng 2003/01/08
'Private Sub Command2_Click()
Public Sub Command2_Click()
   Dim bEditDoc As Boolean
   
   If Me.Visible Then bolUserClick = IIf(Me.ActiveControl.Name = "Command2", True, False)  'Add by Morgan 2011/4/26
   
   If FormCheck = False Then
      'MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   PUB_RestorePrinter Combo1
   Forms(0).Enabled = False 'Added by Morgan 2014/7/17
   
   'Added by Morgan 2017/8/7
   bEditDoc = m_bEditDoc
   If txtOutMode = "3" Then
      txtOutMode = "1"
      m_bEditDoc = True
   End If
   'end 2017/8/7
   
   m_strOutErr = "" 'Added by Lydia 2020/09/10
   
   PrintData

   m_bEditDoc = bEditDoc 'Added by Morgan 2017/8/7
   
   Forms(0).Enabled = True 'Added by Morgan 2014/7/17
   If strCon10 <> MsgText(602) Then
      FormClear
   End If
   PUB_RestorePrinter strPrinter
   Screen.MousePointer = vbDefault
   StatusView MsgText(100)
   
   'Added by Lydia 2020/09/10 «D¥~³¡©I¥s¡Aµo¥Í¿ù»~µoemail³qª¾
   If (m_bBeCalled = False And m_CallPrevForm = "") And m_strOutErr <> "" Then
        PUB_SendMail strUserNum, strUserNum, "", "½Ð´Ú³æ¹q¤lÀÉ²£¥Í¥¢±Ñ", "½Ð´Ú³æ¹q¤lÀÉ²£¥Í¥¢±Ñ¡G" & vbCrLf & Replace(m_strOutErr, "¡®", vbCrLf) & vbCrLf
        MsgBox "½Ð´Ú³æ¹q¤lÀÉ²£¥Í¥¢±Ñ¡G" & vbCrLf & Replace(m_strOutErr, "¡®", vbCrLf) & vbCrLf & "½Ð°Ñ¦Ò¡I", vbInformation, Me.Caption & "-¹q¤lÀÉ²£¥Í¥¢±Ñ"
   End If
   'end 2020/09/10
   
End Sub

Private Sub Form_Activate()
'edit by nickc 2007/02/08
'   '93.3.16 ADD BY SONIA
'   If IsObject(mdiMain) Then
'      mdiMain.toolshow
'   End If
'   '93.3.16 END
   Dim formCnt As Integer
   For formCnt = 0 To Forms.Count - 1
       If UCase(Forms(formCnt).Name) = "MDIMAIN" Then
             Forms(formCnt).ToolShow
             Exit For
       End If
   Next
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      StatusView MsgText(100)
   End If
End Sub

Private Sub Form_Load()
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim ii As Integer
   
   PUB_InitForm Me, Me.Width, Me.Height, strBackPicPath4
   
   'Modified by Morgan 2021/5/18 +¥uÅã¥Ü¦³®Äªº¦Lªí¾÷
   PUB_SetPrinter Me.Name, Combo1, strPrinter, , , , , True
   PUB_SetPrinter Me.Name, Combo2, , , , , , True
   
   'Removed by Morgan 2015/6/22
   'Check1.Value = Val(PUB_GetLastDate(Me.Name, pub_HostName & Me.Check1.Name))
   
   StatusView MsgText(100)

'Modified by Morgan 2014/3/7 ¬Y¨Çª¬ªp¦C¦L·|¶]±¼,«ì´_¤@«ß³£Åã¥Ü
'   'Added by Morgan 2014/1/23
'   strExc(0) = Pub_GetSpecMan("WordÅã¥Ü°õ¦æ¹q¸£²M³æ")
'   If InStr(strExc(0), pub_HostName) > 0 Then
'      m_bShowWord = True
'   End If
'   'end 2014/1/23
   m_bShowWord = True
'end 2014/3/7

   'Added by Morgan 2014/7/18
   'Modified by Morgan 2015/1/29 ¶}©ñ³£¯à¥Î
   'If Pub_StrUserSt03 = "M51" Then
      Check2.Visible = True
   'End If
   'end 2014/7/18
   
   'Removed by Morgan 2021/12/15 ¨ú®ø,Word³£¤w§ï·sª©¥i¤ä´©Âàpdf¥\¯à
   'If PUB_PdfCreatorNameInWord = "" Then PUB_PdfCreatorNameInWord = PUB_GetCreatorNameInWord 'Added by Morgan 2017/9/27
   'end 2021/12/15
   
   'Added by Lydia 2021/02/23
   If Pub_StrUserSt03 <> "M51" Then
       CmdHelp.Visible = False
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Modify By Cheng 2003/03/24
    '­Y¬O§_¦C¦L¦a§}±ø¤W"Y"
    If Me.txtAdd.Text = "Y" Then
        'Add By Cheng 2003/01/29
        '¦C¦L¦a§}±ø
        PUB_PrintAddressList strUserNum, Me.Combo2.Text
        '§R°£¦a§}±ø¦Cªí¸ê®Æ
        PUB_DeleteAddressList strUserNum
        'ªì©l¤Æ§Ç¸¹
        pub_AddressListSN = 0
        '¦Lªí¾÷³]¦^¹w³]¦Lªí¾÷
        For Each prnPrint In Printers
           If prnPrint.DeviceName = strPrinter Then
              Set Printer = prnPrint
           End If
        Next
    End If
    'Add By Cheng 2003/03/17
    '­Y¦Lªí¾÷ÅÜ°Ê, «h§ó·s¦C¦L³]©w
    If Me.Combo1.Text <> Me.Combo1.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
    End If
    '­Y¦Lªí¾÷ÅÜ°Ê, «h§ó·s¦C¦L³]©w
    If Me.Combo2.Text <> Me.Combo2.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
    End If
    
    'Removed by Morgan 2015/6/22
    'PUB_SaveLastDate Me.Name, pub_HostName & Me.Check1.Name, Check1.Value 'Added by Morgan 2014/6/30
    
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set adoLEDES = Nothing 'Added by Morgan 2012/4/24
   Set Frmacc2480 = Nothing
End Sub

Private Sub Text1_GotFocus()
   CloseIme 'Add by Morgan 2008/7/2
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
   CloseIme 'Add by Morgan 2008/7/2
   'Add by Morgan 2008/5/30 °_¨´¹w³]¤@¼Ë-³¢
   If Text2 = "" And Text1 <> "" Then
      Text2 = Text1
   End If
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
' ²M°£µe­±
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
   Text2 = ""
   If Me.Visible Then Text1.SetFocus
   'Added by Lydia 2015/04/09 ¤¤¤åª©-¾ã§å½Ð´Ú³æ
   m_bolChiDB = False: m_Chi2Word = False
   m_ChiSys = ""
   m_ChiApply = ""
   m_ChiCust = ""
End Sub

'*************************************************
' ¦C¦L©ú²Ó¸ê®Æ
'
'*************************************************
Private Sub PrintData()
   Dim strDescription As String
   Dim strAnnuity As String
   Dim i As Integer
   Dim s As Integer
   Dim strTemp1 As Variant
   Dim strTemp2 As Variant
   Dim strSendDate As String
   Dim intYear As Integer
   Dim StrSQLa As String
   Dim ii As Integer
   Dim strA1K01 As String '½Ð´Ú³æ¸¹
   Dim intTMKindCnt As Integer '°Ó«~Ãþ§O¼Æ
   Dim dblGrpAmt As Double '¬Û¦P½Ð´Ú¶µ¥Øª÷ÃB¦X­p¼Æ
   'Add by Morgan 2010/11/9
   Dim dblGrpAmtNoDisc As Double '¬Û¦P½Ð´Ú¶µ¥Øª÷ÃB¦X­p¼Æ(¥¼§é¦©)
   Dim strAmountNoDisc As String '½Ð´Úª÷ÃB(¥¼§é¦©)
   'End 2010/11/9
   Dim strA1J02 As String '½Ð´Ú¶µ¥Ø¥N¸¹
   Dim strA1J06 As String 'Add by Morgan 2005/3/31
   Dim iItemNo As Integer
   Dim douNTDollar  As Double '¥x¹ô¥[Á`
   
   'add by nickc 2007/02/08
   Dim douUSDollar As Double
   'Add by Morgan 2008/3/13
   Dim iCopys As Integer
   Dim strDNMemoAlertList As String 'D/N³Æµù¤w´£¿ô²M³æ
   Dim str1stNo As String 'Added by Morgan 2018/1/15 ¬d¦W²Ä1Ãþ¶µ¦¸
   Dim strItemNoFrom As String, strItemNoTo As String 'Add by Morgan 2010/10/12
   Dim douUSAmount As Double, douUSAmountNoDisc As Double 'Add by Morgan 2010/12/29
   Dim arrName
   Dim strSpecialList As String 'Add by Morgan 2011/10/11 ¯S®í½Ð´Ú³æ²M³æ
   Dim strSpecialL2 As String 'Added by Lydia 2015/04/15 +¾ã§å½Ð´Ú³æ²M³æ
   Static bCanPDF As Boolean  'Added by Morgan 2012/11/6 ¬O§_¦³¦w¸Ë PDFCreator
   
   Dim dblGrpAmtFAmt As Double, dblGrpAmtNoDiscFAmt As Double 'Add By Sindy 2013/3/29
   Dim dbl_A1L0507FAmt As Double, dbl_A1L05FAmt As Double 'Add By Sindy 2013/3/29
   Dim mChiCNo As String, mChiOAmt As String, mChiUAmt As String 'Added by Lydia 2015/04/09 °O¿ý¦X¨Ö®×¸¹+¶µ¥Ø,½Ð´Ú¹ô§O
   Dim stIDate As String 'Added by Morgan 2023/8/29
   Dim douDiscount As Double 'Added by Morgan 2023/10/18
   
On Error GoTo Checking
   
   'ªì©l¤Æ¥»©Ò®×¸¹ÅÜ¼Æ
   m_CP01 = "": m_CP02 = "": m_CP03 = "": m_CP04 = ""
   strA1K01 = ""
   intLength = 0: douAmount = 0
   strNewPage = ""
   intYear = 0
   'Add by Morgan 2010/11/9
   Erase strLedes
   iUpper = 0
   m_sEBillingMsg = ""
   m_ItemDesc = ""
   m_ItemHDesc = "" 'Added by Morgan 2013/10/24
   m_ItemTDesc = "" 'Added by Morgan 2013/10/24
   m_ItemXDesc = "" 'Added by Morgan 2013/10/24
   'end 2010/11/9
   m_ClearItemDesc = "" 'Add by Morgan 2010/12/7
   mChiCNo = "": mChiOAmt = "": mChiUAmt = ""  'Added by Lydia 2015/04/09 ¤¤¤åª©-¾ã§å½Ð´Ú³æ
   
   'Added by Lydia 2020/09/08 ³B²z¦C¦L°O¿ý©M¬d¸ß±ø¥ó¤Î­pºâ; µ{§Ç¤Ó¤j©â¥X¬°¤lµ{§Ç
   'Modify by Sindy 2021/10/26 ¤S¥X²{µ{§Ç¤Ó¤j,¬d¸ß«áªº­pºâ¤]²¾¤J¦¹¨ç¼Æ
   'Modified by Morgan 2024/11/1
   'Call PrintData_1(strSpecialList, strSpecialL2, strA1k01)
   If PrintData_1(strSpecialList, strSpecialL2, strA1K01) = False Then Exit Sub
   'end 2024/11/1
   
   'Added by Morgan 2023/11/17
   '¥~±M¤é¤å½Ð´Ú³æ¥u½Ð³W¶O®É½Ð´Ú¶µ¥Ø±a¥D¶µªº±Ô­z--±Ó²ú
   StrSQLa = "update ACCRPT428 a set R42809=(select a1j16 from acc1j0 where a1j01=R42822 and a1j02||'99'=R42840)" & _
      " Where R42801='" & strUserNum & "' and substr(R42840,-2)='99' and exists(select * from caseprogress where cp60=R42831 and cp12 like 'F2%')" & _
      " and not exists(select * from ACCRPT428 b where R42801=a.R42801 and R42831=a.R42831 and R42840||'99'=a.R42840)" & _
      " and exists(select a1j16 from acc1j0 where a1j01=R42822 and a1j02||'99'=R42840 and a1j16 is not null)"
   adoTaie.Execute StrSQLa, intI
   'end 2023/11/17
   
   'Modify by Morgan 2011/2/23 +IDate
   'Modify By Sindy 2011/3/7 +fa108
   'Modified by Morgan 2012/12/6 +a1k33
   'Modified by Sindy 2013/3/29 +,R42856 A1L0507FAmt,R42857 A1L05FAmt
   'Modified by Morgan 2013/10/8 ¨ú®ø fa103(R42851),fa104(R42852),fa105(R42853)
   'Modified by Morgan 2014/2/18 +a2616(R42851)
   StrSQLa = "Select R42802 As a1k27, R42803 As a1l07, R42804 As a1l05, R42805 As a1j04, R42806 As a1l06, R42807 As a1j05, R42808 As a1j06, R42809 As a1j16, R42810 As a1j10, " & _
                   "R42811 As fa05, R42812 As fa63, R42813 As fa64, R42814 As fa65, R42815 As fa32, R42816 As fa18, R42817 As a1k02, R42818 As fa33, R42819 As fa19, R42820 As fa34, " & _
                   "R42821 As fa20, R42822 As a1k13, R42823 As a1k14, R42824 As a1k15, R42825 As a1k16, R42826 As fa21, R42827 As fa22, R42828 As fa35, R42829 As a1k03, R42830 As fa36, " & _
                   "R42831 As a1k01, R42832 As fa06, R42833 As fa23, R42834 As a1k04, R42835 As a1k10, R42836 As a1l02, R42837 As fa43, R42838 As Curr, R42839 As cu102, R42840 As a1j02" & _
                   ",R42841 As a1k05, R42842 As FA04, R42843 As FA17, R42844 a1j03, R42845 a1k08, R42846 a1k11, R42847 a1k28" & _
                   ",R42848 a1j18,R42849 a1j19,R42850 a1j20,R42851 a2616,R42854 fa108" & _
                   ",R42855 a2604,A1K33,R42856 A1L0507FAmt,R42857 A1L05FAmt,null as IDate" & _
                   " From ACCRPT428,ACC1K0 Where R42801='" & strUserNum & "' AND A1K01(+)=R42831 Order By R42831 Asc, R42829 Asc, R42836 Asc"
   adoacc1k0.CursorLocation = adUseClient
   adoacc1k0.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc1k0.RecordCount = 0 Then adoacc1k0.Close: GoTo OnlySpecial 'Add by Morgan 2011/10/11
   
   'Add by Morgan 2010/4/13 ÂàÂ÷½u¸ê®Æ¶°¥H«K¥i­×§ï¤º®e
   'Modify by Amy 2014/06/24 +FormName §ï¼È¦sTB
   Set RsTemp = PUB_CreateRecordset(adoacc1k0, , , , Me.Name)
   Set adoacc1k0 = RsTemp.Clone
   'end 2010/4/13
   
   txtCopy.Text = IIf(Val("0" & Me.txtCopy.Text) = 0, 3, Val("0" & Me.txtCopy.Text))

   'Add by Morgan 2008/4/22
   '«D¾ã§å¦C¦L¥B¥u¦L³æµ§®É­Y¦³¦s¹q¤lÀÉ«h´£¿ô
   'Modified by Morgan 2018/6/27 +m_bLedesOnly ¤]¤£¼u°T®§
   If m_bLedesOnly = False And m_bBeCalled = False And Text1 = Text2 Then
      m_bMsg = True
   End If
   
'Added by Morgan 2012/11/6
   If Not bCanPDF Then
      'ÀË¬d¬O§_¦³¦w¸ËPDFCreator
      If PUB_PrinterIndex("PDFCreator") >= 0 Then
         bCanPDF = True
      End If
   End If
   m_b2PDF = False
   m_bPrint2Pdf = False
   
   'Added by Morgan 2025/8/19
   '¤º°Óµ{§Ç¾Þ§@®É¤@«ß¥u¦sPDFÀÉ¦Ü¤½¥Î¹q¸£
   If Not m_bEditDoc And Left(Pub_StrUserSt03, 2) = "P2" Then
      m_b2Printer = False
      m_b2Picture = False
      m_bPrint2Pdf = False
      m_bPrintWord = False
      m_b2PDF = True
   End If
   'end 2025/8/19
   
PrintPdfStart:

   adoacc1k0.MoveFirst
'end 2012/11/6
   
   '¥u²£¥Í¹q¤lÀÉ
   If txtOutMode = "2" Then
      m_b2Printer = False
      m_b2Picture = True
      'Added by Morgan 2014/3/3 ±±¨î¥u­n¶]1¦¸
      If bCanPDF Then
         m_b2Picture = False
         m_b2PDF = True
         m_PdfDone = False 'Added by Morgan 2014/8/26
         m_bPrint2Pdf = True
      End If
      'end 2014/3/3
   End If

   'Modified by Morgan 2012/11/6
   'If m_b2Printer Then
   If m_b2Printer Or m_bPrint2Pdf Then
      Printer.Orientation = 1
      '¹w³]¦C¦L¦r«¬
      Printer.FontSize = 12
      Printer.Font.Name = "Times New Roman"
      '³]©w¦C¦L¥÷¼Æ
      If m_bPrint2Pdf Then
         Printer.Copies = 1
      Else
         Printer.Copies = Val(txtCopy.Text)
      End If
   End If
   
   'ÅÜ¼Æªì©l¤Æ
   strNo = ""
   m_iPages = 1
   douAmount = 0
   douUSDollar = 0
   douNTDollar = 0
   strNewPage = ""
   m_strA1K01 = ""
   m_dblDiscTot = 0 'Add by Morgan 2018/9/18
   m_dblNoDiscAmtTot = 0 'Add by Morgan 2018/9/18
   
   Do While adoacc1k0.EOF = False
      'Modified by Lydia 2015/04/08 ¾ã§å½Ð´Ú³æ¤£´«­¶
'      If strNo <> adoacc1k0.Fields("a1k01").Value & adoacc1k0.Fields("a1k27").Value Then
'         If douAmount <> 0 Then
      If strNo <> adoacc1k0.Fields("a1k01").Value & adoacc1k0.Fields("a1k27").Value Then
         If (m_bolChiDB = False And douAmount <> 0) Then
            If m_b2Printer Then
               Printer.Line (0 + intDefault, m_DetailTopStart + intCounter * m_LineH - 200 + intTop)-(10000 + intDefault, m_DetailTopStart + intCounter * m_LineH - 200 + intTop)
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.Line (douExtRate * (0 + intDefault), douExtRate * (m_DetailTopStart + intCounter * m_LineH - 200 + intTop))-(douExtRate * (10000 + intDefault), douExtRate * (m_DetailTopStart + intCounter * m_LineH - 200 + intTop))
            'End If
            'end 2022/8/4
            
            'Added by Morgan 2023/10/18
            '§é¦©Á`©M¥t¦C³Ì«áªÌ
            If m_bEBilling And douDiscount <> 0 Then
               iUpper = iUpper + 1
               ReDim Preserve strLedes(m_iCols, iUpper)
               For intI = 1 To 8
                  strLedes(intI, iUpper) = strLedes(intI, iUpper - 1)
               Next
               strLedes(9, iUpper) = iUpper
               strLedes(10, iUpper) = "IF"
               strLedes(11, iUpper) = ""
               strLedes(12, iUpper) = Round(douDiscount, 4)
               strLedes(13, iUpper) = strLedes(12, iUpper)
               strLedes(14, iUpper) = strLedes(14, iUpper - 1)
               strLedes(15, iUpper) = ""
               strLedes(16, iUpper) = ""
               strLedes(17, iUpper) = ""
               strLedes(18, iUpper) = strLedes(18, iUpper - 1)
               strLedes(19, iUpper) = "Discount"
               strLedes(20, iUpper) = strLedes(20, iUpper - 1)
               strLedes(21, iUpper) = ""
               '22~
               For intI = 22 To m_iCols
                  strLedes(intI, iUpper) = strLedes(intI, iUpper - 1)
               Next
            End If
            douDiscount = 0
            'end 2023/10/18
   
            PrintSum
            
            If m_b2Printer And Not m_b2PDF And Not m_b2Picture Then
               '·s¼W¦a§}±ø¦Cªí¸ê®Æ
               If Me.txtAdd.Text = "Y" Then
                  'Modified by Morgan 2012/10/3
                  '¾ã§å½Ð´Ú¥u­n¦C¦L1±i
                  If Not (m_bolOneAddr And m_bolAddAddrOK) Then
                     pub_AddressListSN = pub_AddressListSN + 1
                     PUB_AddNewAddressList strUserNum, m_CP01, m_CP02, m_CP03, m_CP04, "" & pub_AddressListSN, "0", GetCP10(strA1K01)
                     m_bolAddAddrOK = True
                  End If
               End If
            End If
            
            If m_b2Printer Then
               'Modified by Morgan 2012/11/6
               'Printer.NewPage
               If m_bPrint2Pdf Then
                  Printer.EndDoc
                  frmPDF.EndtProcess
                  'Added by Lydia 2020/02/15 §PÂ_ÀÉ®×¬O§_¦s¦b, ¶W¹L®É¶¡´NÄ~Äò
                  If Me.Tag <> "" And InStr(Me.Tag, "\") > 0 Then
                    'Modified by Lydia 2020/09/10 ¶W¹L®É¶¡¡Aª½±µ°O¿ý¥¢±Ñ²M³æ
                    'If PUB_ChkFileStatus(Me.Tag) = False Then
                    If PUB_ChkFileStatus(Me.Tag, False, m_strOutErr) = False Then
                    End If
                  End If
                  'end 2020/02/15
                  Unload frmPDF
               Else
                  Printer.NewPage
               End If
            End If
            
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   If m_iPages > 1 Then
            '      SetPic m_iPages, True
            '   Else
            '      SetPic 0, True
            '   End If
            'End If
            'end 2022/8/4
            
            m_iPages = 1
            douAmount = 0
            douUSDollar = 0
            douNTDollar = 0
            strNewPage = ""
            m_strA1K01 = ""
            str1stNo = "" 'Added by Morgan 2018/10/15 ¤é¤å¬d¦W²Ä1Ãþ­n³æ¿W¦C
         End If 'If douAmount <> 0 Then
         
         strNo = adoacc1k0.Fields("a1k01").Value & adoacc1k0.Fields("a1k27").Value
         strA1K01 = "" & adoacc1k0.Fields("a1k01").Value
         m_strA1K01 = m_strA1K01 & "'" & adoacc1k0.Fields("a1k01").Value & "',"
         strRemark = "" & adoacc1k0.Fields("a1k05").Value
         
         '°O¿ý¥»©Ò®×¸¹
         m_CP01 = "" & adoacc1k0.Fields("a1k13").Value
         m_CP02 = "" & adoacc1k0.Fields("a1k14").Value
         m_CP03 = "" & adoacc1k0.Fields("a1k15").Value
         m_CP04 = "" & adoacc1k0.Fields("a1k16").Value
         m_strCaseNo = m_CP01 & m_CP02 & IIf(m_CP03 & m_CP04 <> "000", m_CP03 & m_CP04, "")  'Add by Morgan 2008/4/8
         
         m_A1k03 = "" & adoacc1k0.Fields("a1k03").Value
         m_A1k27 = "" & adoacc1k0.Fields("a1k27").Value 'Add by Morgan 2012/4/17
         m_A1k28 = "" & adoacc1k0.Fields("a1k28").Value 'Add by Morgan 2008/5/29
         m_A1k08 = "" & adoacc1k0.Fields("a1k08").Value 'Add by Morgan 2011/4/1
         '°O¿ý¶×²v
         m_strOriExcRate = "" & adoacc1k0.Fields("a1k10").Value
         m_DNCurr = "" & adoacc1k0.Fields("Curr").Value
         
         'Modied by Morgan 2012/12/7
         '·sÂÂ½Ð´Ú³æ¥H¬O§_¦³³]©w¦C¦L¹ô§O®æ¦¡(a1k33)§PÂ_
         If Not IsNull(adoacc1k0("a1k33")) Then
            m_bolNewBill = True
            m_iPrintCurrType = Val(adoacc1k0("a1k33"))
            
            'Add By Sindy 2021/3/3 ¨ú±o°Ó¼Ð½Ð´Ú³æ¤§¥»©Ò±b¤á
            Call PUB_GetDefaultCurrPrintType(m_CP01, m_A1k28, m_DNCurr, , m_CP02, m_CP03, m_CP04, , m_strFA126)
            '2021/3/3 END
         Else
            m_bolNewBill = False
            'Modify By Sindy 2016/11/29 + , , m_CP02, m_CP03, m_CP04
            'm_iPrintCurrType = PUB_GetDefaultCurrPrintType(m_CP01, m_A1k27, m_DNCurr)
            'Modified by Morgan 2018/4/27
            'm_iPrintCurrType = PUB_GetDefaultCurrPrintType(m_CP01, m_A1k27, m_DNCurr, , m_CP02, m_CP03, m_CP04)
            'Add By Sindy 2021/3/3 ¨ú±o°Ó¼Ð½Ð´Ú³æ¤§¥»©Ò±b¤á + ,m_strFA126
            m_iPrintCurrType = PUB_GetDefaultCurrPrintType(m_CP01, m_A1k28, m_DNCurr, , m_CP02, m_CP03, m_CP04, m_A1k27, m_strFA126)
            '2016/11/29 END
         End If
         'end 2012/12/7
            
         'Add by Morgan 2011/1/5
         m_strA1K02 = "" & adoacc1k0.Fields("a1k02").Value
         Call PrintData_2(strA1K01, strDNMemoAlertList)  'Added by Lydia 2020/09/08 §PÂ_¯S®í«È¤á¡þ®æ¦¡; µ{§Ç¤Ó¤j©â¥X¬°¤lµ{§Ç

         intCounter = 0
         iItemNo = 0
         m_strDN = strA1K01
         
         '¥u²£¥Í¹q¤lÀÉ
         If txtOutMode = "2" Then
            m_b2Printer = False
            m_b2Picture = True
            m_bEBilling = False 'Added by Morgan 2012/10/31
         
         Else
            'Added by Morgan 2025/8/20 ¤º°Óµ{§Ç¾Þ§@®É¤@«ß¥u¦sPDFÀÉ¦Ü¤½¥Î¹q¸£
            If Left(Pub_StrUserSt03, 2) = "P2" Then
               m_b2Printer = False
            Else
            'end 2025/8/20
               m_b2Printer = True
            End If
            m_bEBilling = SetEBilling(m_A1k28) 'Add by Morgan 2010/11/5 ¬O§_¹q¤l½Ð´Ú
         End If
         
         'Add by Morgan 2011/3/14 1998BI
         If m_bEBilling Then
            'Modified by Morgan 2012/4/24 LEDED ³]©w§ï§ì LEDED ¸ê®Æªí
            strExc(0) = "select * from LEDES where LD01='" & m_A1k28 & "'"
            intI = 1
            Set adoLEDES = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               m_iLedesVer = Val("" & adoLEDES.Fields("ld15"))
               If m_iLedesVer = 0 Then m_iLedesVer = 1
               If m_iLedesVer = 2 Then
                  m_iCols = 52
               Else
                  m_iCols = 24
               End If
               m_LD16 = "" & adoLEDES.Fields("ld16") 'Added by Morgan 2025/3/14
            Else
               m_bEBilling = False
               m_sEBillingMsg = m_sEBillingMsg & vbCrLf & adoacc1k0("a1k01") & " LEDES ³]©wÅª¨ú¥¢±Ñ¡AµLªk²£¥Í¹q¤l±b³æ¡I"
            End If
            'end 2012/4/24
            
            'end 2012/4/24
            Erase strLedes
            iUpper = 0
         End If
         'end 2010/11/5
         
         'Add by Morgan 2010/11/25
         m_b2Word = False
         m_bSaveWord = False
         m_bPrintWord = False
         
         'Added by Morgan 2016/7/27
         '¥u¦L LEDES ¹q¤l±b³æ
         If Check2.Value = vbChecked Then
            m_b2Picture = False
            m_b2Printer = False
            m_bPrint2Pdf = False
         'end 2016/7/27
         
         'Add by Morgan 2011/2/16
         'EATON ½Ð´Ú³æ¯S®í(­n¸É¥N½X®æ¦¡»Pledes¤£¦P)¥B­n¦spdf
         ElseIf m_A1k28 = "Y20438000" Then
            m_b2Word = True
            m_bSaveWord = True
            m_b2Picture = False
            m_b2Printer = False
            
         '¯S®í½Ð´Ú
         ElseIf m_bEditDoc Then
            m_b2Word = True
            m_b2Picture = False
            m_b2Printer = False
            
         Else
            'Modified by Morgan 2013/10/31
            'If m_bSpecial1 Or m_bSpecial2 Or m_bSpecial3 Then
            'Modified by Morgan 2014/2/24 + m_bSpecial4
            'Modified by Lydia 2016/03/03 + m_bSpecial5
            If m_bSpecial1 Or m_bSpecial2 Or m_bSpecial3 Or m_bSpecial4 Or m_bSpecial5 Or bolNewForm Then
               m_b2Word = True
            End If

            'Add by Morgan 2008/4/7
            '­Y¥Ñ¨ä¥Lµ{¦¡©I¥s®É¹q¤lÀÉ³]©w·Óm_bEMail
            If m_bBeCalled Then
               m_b2Picture = m_bEMail
               iCopys = 1
            '­Y¤£¬O¥Ñ¨ä¥Lµe­±©I¥s¥B¿é¥X¤è¦¡¬°¦Lªí¾÷®É®É»Ý§PÂ_¬O§_­n²£¥Í¹q¤lÀÉ
            ElseIf m_b2Printer Then
               'Modified by Morgan 2014/5/30 ¥[§PÂ_¬O§_¦³³] D/N e¤Æ
               m_b2Picture = PUB_GetEMailFlag(m_CP01 & m_CP02 & m_CP03 & m_CP04, , strA1K01, m_bPaper, , True)
            End If
            
         End If
         
         'Added by Morgan 2012/11/6
         m_bWord2Pdf = False
         '¦C¦LPDF
         If m_bPrint2Pdf Then
            '¹q¤lÀÉ
            If m_b2Picture Then
               m_b2Picture = False
               If m_b2Word Then
                  m_b2Printer = False
                  m_bPrintWord = False
                  If Not m_b2PDF Then 'Added by Morgan 2017/9/26
                     m_bWord2Pdf = True
                  End If
                  m_iSpCopies = 1
               Else
                  m_b2Printer = True
                  MyNewPage True
               End If
            '«D¹q¤lÀÉ¥»¦¸¤£°Ê§@
            Else
               m_b2Printer = False
               m_b2Word = False
            End If
            
         '¤@¯ë¦C¦L
         Else
         'end 2012/11/6
         
            'Add by Morgan 2009/9/29
            If m_b2Printer Then
               
               '­Y¦³³]©w­n²£¥Í¹q¤lÀÉ®É¤´»Ý¦C¦L¤@¥÷
               'Modified by Morgan 2011/5/13 +m_bEBilling
               If (m_b2Picture And Not m_bPaper) Or m_bEBilling Then
                  'Modified by Morgan 2022/8/18 ­Ó®×³]©wÀu¥ý Ex:FCT-031774 (E¤ÆºA¼Ë¤Ó¦h,»Ýµø¹ê»Úª¬ªp­×§ï³W«h)
                  'iCopys = 1
                  iCopys = GetSpecificCopy(Val(txtCopy.Text), m_A1k28, m_CP01, m_CP02, m_CP03, m_CP04, True)
                  
               'Add by Morgan 2011/4/26
               '­Y¬Oª½±µ°õ¦æ¥»µ{¦¡(¨Ï¥ÎªÌ«ö¦C¦L«ö¶s)®É§ìµe­±¥÷¼Æ
               ElseIf bolUserClick Then
                  'Modified by Morgan 2014/7/14 ­Y¦³³]©w®É¤´¥H³]©w¬°¥D--¨q¬Â
                  'iCopys = Val(txtCopy.Text)
                  iCopys = GetSpecificCopy(Val(txtCopy.Text), m_A1k28, m_CP01, m_CP02, m_CP03, m_CP04)
                  'end 2014/7/14
               Else
                  'Modify by Morgan 2011/7/8 ¥i«ü©w¥÷¼Æ
                  If m_iCopies > 0 Then
                     iCopys = m_iCopies
                  Else
                     iCopys = GetSpecificCopy(Val(txtCopy.Text), m_A1k28, m_CP01, m_CP02, m_CP03, m_CP04)
                  End If
               End If
               'Added by Lydia 2020/07/16 ¦~ÃÒ¶O½Ð´Ú¨ç¡Ge¤Æ¤£¥X¯È¥»¡A«DE¤Æ¥X¯È¥»¤G¥÷; ­Y¦³«ü©w¥÷¼Æ>1«h´î¤@¥÷,«ü©w¥÷¼Æ=1«h¤£¥X¯È¥»
               'Modified by Lydia 2021/01/21 ¼W¥[¹ê¼f½Ð´Ú¨ç
               'If m_CallPrevForm = "frm060307" Then
               If m_CallPrevForm = "frm060307" Or m_CallPrevForm = "frm060306_7" Then
                   'Added by Lydia 2020/12/23 ­Y¬OE+±H¦L¤@¥÷,©T©w³£­n¦L; ex.Y53715000(E+±H)¥u¦L¤@¥÷
                   If m_bEMail = True And m_bPaper = True Then
                        If iCopys > 1 Then iCopys = iCopys - 1
                   Else
                   'end 2020/12/23
                        If iCopys > 0 Then iCopys = iCopys - 1
                   End If 'Added by Lydia 2020/12/23
               End If
               'end 2020/07/16
               
               'Add by Morgan 2011/1/6
               If m_b2Word Then
                  m_b2Printer = False
                  m_bPrintWord = True
                  m_iSpCopies = iCopys
               'Added by Morgan 2022/10/26
               ElseIf iCopys = 0 Then
                  m_b2Printer = False
               'end 2022/10/26
               Else
               'end 2011/1/6
                  '¥÷¼Æ¦³ÅÜ®É­n­«³]
                  If iCopys <> Printer.Copies Then
                     Printer.EndDoc
                     Printer.Orientation = 1
                     Printer.Copies = iCopys
                  End If
                  Printer.FontSize = 12
                  Printer.Font.Name = "Times New Roman"
               End If
            End If
            'end 2009/9/29
            
            If m_b2Picture Then
               'Add by Morgan 2011/1/6
               '­Y­n²£¥Í¹q¤lÀÉ¥B¥iª½±µ¿é¥X¬°PDF®æ¦¡®É,¤£¥²²£¥Í¹ÏÀÉ
               If m_b2PDF Or bCanPDF Then
                  m_b2Picture = False
                  m_b2PDF = True
                  m_PdfDone = False 'Added by Morgan 2014/8/26
               ElseIf m_b2Word Then
                  m_b2Picture = False
                  m_bSaveWord = True
               Else
               'end 2011/1/6
               
                  Picture1.AutoRedraw = True
                  'Picture1.Height = 16836
                  'Picture1.Width = 11904
                  NewPic
               End If
            End If
            
         End If 'Added by Morgan 2012/11/6 'If m_bPrint2Pdf Then

        
         'Added by Lydia 2015/04/08 §PÂ_¾ã§å½Ð´Ú³æ¥u¦b²Ä¤@¦¸°µªì©l¤Æ
         'If m_b2Word Then
         If (m_bolChiDB = False And m_b2Word) Or (m_bolChiDB And adoacc1k0.AbsolutePosition = 1) Then
            InitVar 'Added by Morgan 2020/8/6 µ{§Ç¤Ó¤j©â¥X¬°¤lµ{§Ç
         End If
         
         'Added by Lydia 2015/04/09 + ¬O§_¬°¤¤¤å¾ã§å½Ð´Ú³æ§PÂ_(bol_ChiDB)
         bol_ChiDB = False
         PrintHead
         If m_bolChiDB = True And bol_ChiDB = False Then GoTo Chi_JumpPDM
         'end 2015/04/09
         
         intCounter = intCounter + intAddSpaceRow
         
         'Add by Morgan 2011/1/6 ³]©wWord¦sÀÉ¸ô®|
         If m_bSaveWord Then
            m_EFilePath = GetPath
            If Dir(m_EFilePath, vbDirectory) = "" Then
               MkDir m_EFilePath
            End If
            m_EFilePath = m_EFilePath & "\" & m_strCaseNo & "_DN" & m_strDN & ".doc"
         End If
         
         bolFMPCase = False 'Added by Morgan 2013/5/13
         'Add By Sindy 2013/1/28
         bolIsFMP = False
         strExc(0) = "select cp01,cp12,a1k22,a1k19,cp27 from caseprogress,acc1k0 where cp60='" & strA1K01 & "' and cp60=a1k01(+) and cp01='P' and substr(cp12,1,1)='F'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            bolFMPCase = True 'Added by Morgan 2013/5/13
            If Not IsNull(RsTemp.Fields("a1k19")) And RsTemp.Fields("a1k19") > 0 Then
               strCompDate = DBDATE(RsTemp.Fields("a1k19"))
            ElseIf Not IsNull(RsTemp.Fields("a1k22")) And RsTemp.Fields("a1k22") > 0 Then
               strCompDate = DBDATE(RsTemp.Fields("a1k22"))
            End If
            If strCompDate >= AccFMPImputCurrStarDate Then
               bolIsFMP = True
            End If
         End If
         '2013/1/28 End
      End If
      
      'Add by Morgan 2011/2/23
      'Modified by Morgan 2014/2/24 +m_bSpecial4
      'Modified by Lydia 2016/03/03 +m_bSpecial5
      'Modified by Morgan 2018/3/22
      If m_bSpecial1 Or m_bSpecial4 Or m_bSpecial5 Or m_bSpecialNew3 Then
         'Modified by Morgan 2017/3/22 ³W¶O¤]­n§ì¸Ó®×¥ó©Ê½èªºµo¤å¤é
         'strExc(0) = "select cp27 from caseprogress where cp60='" & strA1k01 & "' and cp10='" & adoacc1k0.Fields("a1j02") & "' and cp27>0 order by cp27 asc"
         strExc(1) = adoacc1k0.Fields("a1j02")
         'Added by Morgan 2023/8/29
         If m_CP01 & adoacc1k0.Fields("a1j02") = "FCTA0199" Then
            strExc(1) = "10199"
         End If
         'end 2023/8/29
         If Right(strExc(1), 2) = "99" Then strExc(1) = Left(strExc(1), Len(strExc(1)) - 2)
         strExc(0) = "select cp27 from caseprogress where cp60='" & strA1K01 & "' and cp10='" & strExc(1) & "' and cp27>0 order by cp27 asc"
         'end 2017/3/22
         
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            adoacc1k0.Fields("IDate") = RsTemp.Fields("cp27")
         'Added by Morgan 2018/3/22
         ElseIf m_bSpecialNew3 Then
            adoacc1k0.Fields("IDate") = "--"
         'end 2018/3/22
         ElseIf Not m_bSpecial4 Then
            adoacc1k0.Fields("IDate") = DBDATE(m_strA1K02)
         End If
      End If
      'end 2011/2/23
      
      'Add By Sindy 2013/1/28
      strFMPFee99RMB = ""
      strA1L16 = ""
      If bolIsFMP And Right(adoacc1k0.Fields("a1J02").Value, 2) = "99" Then
         strExc(0) = "select * from acc1L0 where a1L01='" & strA1K01 & "' and a1L04='" & adoacc1k0.Fields("a1J02").Value & "' and (a1L16='RMB' or (a1L18 is not null and a1L18>0))"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.RecordCount > 0 Then
               strA1L16 = "" & RsTemp.Fields("a1L16")
               If Val("" & RsTemp.Fields("a1L18")) > 0 Then
                  strFMPFee99RMB = " ( RMB " & PUB_ChgFormat(RsTemp.Fields("a1L18"), True) & " ) "
               Else
                  strFMPFee99RMB = " ( RMB " & PUB_ChgFormat(RsTemp.Fields("a1L17"), True) & " ) "
               End If
            End If
         End If
      End If
      '2013/1/28 End
      
      m_ItemDesc = "" 'Add by Morgan 2010/11/9
      m_ItemHDesc = "" 'Add by Morgan 2013/10/24
      m_ItemTDesc = "" 'Add by Morgan 2013/10/24
      m_ClearItemDesc = "" 'Add by Morgan 2010/12/7
      
      'Added by Morgan 2020/5/11
      'Y54047ªi­µ²Ä¤G¦¸(§t)¥H«áªº¼f¬d·N¨£³qª¾(1202)¡Ï²Ä¤G¦¸(§t)¥H«áªº¥Ó´_(205)--Tim
      'CodeÅÜ§ó¬°PA699,±Ô­z+subsequent
      If m_A1k28 = "Y54047000" And m_CP01 = "FCP" Then
         If adoacc1k0.Fields("a1j02") = "1202" Or adoacc1k0.Fields("a1j02") = "205" Then
            strExc(0) = "select cp09 from caseprogress a where cp60='" & adoacc1k0.Fields("a1k01") & "' and cp10='" & adoacc1k0.Fields("a1j02") & "'" & _
               " and exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02" & _
               " and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp10(+)=a.cp10 and b.cp05<a.cp05 and b.cp57 is null)"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               adoacc1k0.Fields("a1j18") = "PA699"
               If adoacc1k0.Fields("a1j02") = "1202" Then
                  adoacc1k0.Fields("a1j04") = "Receiving subsequent Official Action / Official Report (Letter), reporting the same to you"
               ElseIf adoacc1k0.Fields("a1j02") = "205" Then
                  adoacc1k0.Fields("a1j04") = "Receiving your instructions, preparing and filing a subsequent response"
               End If
            End If
         End If
      End If
      'end 2020/5/11
            
      'Add by Morgan 2010/4/13 FCPªº940­n§ï¦L·s¥Ó½Ð®×µ{§Ç
      If Left(m_CP01 & adoacc1k0.Fields("a1j02"), 6) = "FCP940" Then
         strExc(1) = Mid(adoacc1k0.Fields("a1j02"), 4)
         'Modify by Morgan 2010/11/9 +a1j18,a1j19,a1j20
         strExc(0) = "select a1j02,a1j03,a1j04,a1j05,a1j06,a1j16,a1j10,a1j18,a1j19,a1j20 from acc1j0 where a1j01='" & m_CP01 & "' and a1j02 in (select cp10||'" & strExc(1) & "' from caseprogress where cp01=a1j01 AND cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "' and cp57 is null and cp10 in ('101','102','103','105'))"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            adoacc1k0.Fields("a1j02") = RsTemp.Fields("a1j02")
            adoacc1k0.Fields("a1j03") = RsTemp.Fields("a1j03")
            adoacc1k0.Fields("a1j04") = RsTemp.Fields("a1j04")
            adoacc1k0.Fields("a1j05") = RsTemp.Fields("a1j05")
            adoacc1k0.Fields("a1j06") = RsTemp.Fields("a1j06")
            adoacc1k0.Fields("a1j16") = RsTemp.Fields("a1j16")
            adoacc1k0.Fields("a1j10") = RsTemp.Fields("a1j10")
            'Add by Morgan 2010/11/9
            adoacc1k0.Fields("a1j18") = RsTemp.Fields("a1j18")
            adoacc1k0.Fields("a1j19") = RsTemp.Fields("a1j19")
            adoacc1k0.Fields("a1j20") = RsTemp.Fields("a1j20")
         End If
      'Add by Morgan 2011/3/4
      ElseIf m_CP01 & adoacc1k0.Fields("a1j02") = "FCTA0199" Then
         adoacc1k0.Fields("a1j02") = "10199"
      End If
      'end 2010/4/13
      
      'Add by Morgan 2011/1/6 ·s¦¡¼Ëªº PA520 ­n´«¦¨ PA530
      If m_CP01 = "FCP" And adoacc1k0.Fields("a1j18") = "PA520" Then
         If ChkIs103(m_CP01, m_CP02, m_CP03, m_CP04) Then
            adoacc1k0.Fields("a1j18") = "PA530"
         End If
      End If
      'Added by Lydia 2016/03/03 §ìOrder number
      If m_bSpecial5 Then
         If Right(adoacc1k0.Fields("a1j02"), 2) = "99" Then
            adoacc1k0.Fields("a1j18") = "6442000865" '³W¶O
         Else
            adoacc1k0.Fields("a1j18") = "6442000864"
         End If
      End If
      'end 2016/03/03
      'Add by Morgan 2010/12/23
      If Not IsNull(adoacc1k0.Fields("a1j03")) Then
         'Modified by Morgan 2025/8/21
         'adoacc1k0.Fields("a1j03") = ParseDesc(adoacc1k0.Fields("a1l05"), adoacc1k0.Fields("a1j03"))
         adoacc1k0.Fields("a1j03") = PUB_ParseItemDesc(adoacc1k0.Fields("a1l05"), adoacc1k0.Fields("a1j03"))
      End If
      If Not IsNull(adoacc1k0.Fields("a1j04")) Then
         'Modified by Morgan 2025/8/21
         'adoacc1k0.Fields("a1j04") = ParseDesc(adoacc1k0.Fields("a1l05"), adoacc1k0.Fields("a1j04"), adoacc1k0.Fields("a1j04") & adoacc1k0.Fields("a1j05") & adoacc1k0.Fields("a1j06"))
         adoacc1k0.Fields("a1j04") = PUB_ParseItemDesc(adoacc1k0.Fields("a1l05"), adoacc1k0.Fields("a1j04"))
      End If
      If Not IsNull(adoacc1k0.Fields("a1j05")) Then
         'Modified by Morgan 2025/8/21
         'adoacc1k0.Fields("a1j05") = ParseDesc(adoacc1k0.Fields("a1l05"), adoacc1k0.Fields("a1j05"), adoacc1k0.Fields("a1j04") & adoacc1k0.Fields("a1j05") & adoacc1k0.Fields("a1j06"))
         adoacc1k0.Fields("a1j05") = PUB_ParseItemDesc(adoacc1k0.Fields("a1l05"), adoacc1k0.Fields("a1j05"))
      End If
      If Not IsNull(adoacc1k0.Fields("a1j06")) Then
         'Modified by Morgan 2025/8/21
         'adoacc1k0.Fields("a1j06") = ParseDesc(adoacc1k0.Fields("a1l05"), adoacc1k0.Fields("a1j06"), adoacc1k0.Fields("a1j04") & adoacc1k0.Fields("a1j05") & adoacc1k0.Fields("a1j06"))
         adoacc1k0.Fields("a1j06") = PUB_ParseItemDesc(adoacc1k0.Fields("a1l05"), adoacc1k0.Fields("a1j06"))
      End If
      If Not IsNull(adoacc1k0.Fields("a1j16")) Then
         'Modified by Morgan 2025/8/21
         'adoacc1k0.Fields("a1j16") = ParseDesc(adoacc1k0.Fields("a1l05"), adoacc1k0.Fields("a1j16"))
         adoacc1k0.Fields("a1j16") = PUB_ParseItemDesc(adoacc1k0.Fields("a1l05"), adoacc1k0.Fields("a1j16"))
      End If
      'end 2010/12/23
         
      If intCounter > 24 Then
         strNewPage = MsgText(602)
         intCounter = 0
         If m_b2Printer Then
            'Modified by Morgan 2012/10/31
            'Printer.NewPage
            MyNewPage
         End If
         'Removed by Morgan 2022/8/4
         'If m_b2Picture Then
         '   SetPic m_iPages
         'End If
         'end 2022/8/4
         m_iPages = m_iPages + 1
      End If
      dblGrpAmt = 0
      dblGrpAmtNoDisc = 0 'Add by Morgan 2010/11/9
      'Add By Sindy 2013/3/29
      dblGrpAmtFAmt = 0
      dblGrpAmtNoDiscFAmt = 0
      '2013/3/29 End

      '2009/5/18 add by sonia ½Ð´Ú¹ô§O«D¬üª÷®É§ì½Ð´Ú¹ô§O¹ï¥x¹ô¶×²v
      m_DNCurr = "" & adoacc1k0.Fields("Curr").Value
      dbl_A1L0507FAmt = "" & adoacc1k0.Fields("A1L0507FAmt").Value 'Add by Sindy 2013/3/29
      dbl_A1L05FAmt = "" & adoacc1k0.Fields("A1L05FAmt").Value 'Add by Sindy 2013/3/29
      'Modify By Sindy 2013/1/28
'      If bolIsFMP = True Then
''         If Right(Trim(adoacc1k0.Fields("A1J02").Value), 2) = "99" Then
'            '¥H½Ð´Ú¹ô§O¹w¦ôµ²¶×¶×²v´«ºâ½Ð´Úª÷ÃB
'            m_DNRate = PUB_GetAcc210(2, m_DNCurr, m_strA1K02)
''         Else
''            '¥H½Ð´Ú¹ô§Oªº½Ð´Ú¶×²v´«ºâ½Ð´Úª÷ÃB
''            m_DNRate = m_strOriExcRate
''         End If
'      Else
         m_DNRate = m_strOriExcRate '´N¬OA1k10Äæ¦ì­È
'      End If
      '2013/1/28 End
      If m_DNCurr <> "USD" Then
         '§ì½Ð´Ú¹ô§O¹ï¥x¹ô¶×²v
         'm_DNRate = PUB_GetUSXRate_1("" & adoacc1k0.Fields("a1k02").Value, m_DNCurr)
         '§ì½Ð´Ú¹ô§O¹ï¬üª÷¶×²v
         'Modify By Sindy 2013/1/10
         'm_strOriExcRate = PUB_GetDNRate("" & adoacc1k0.Fields("a1k02").Value, m_DNCurr)
         m_DUsdRate = PUB_GetDNRate("" & adoacc1k0.Fields("a1k02").Value, m_DNCurr)
         '2013/1/10 End
         'Add by Morgan 2010/11/9
         If m_bEBilling Then
            'Modified by Morgan 2021/3/24 §ï¯Â¥~¹ô¤]¥i Ex:X11003742(EUR) --Anny
            'Modified by Morgan 2021/5/20 ¥~¹ô+¬üª÷¤]¥i
            'Modified by Morgan 2021/5/25 ¯Â¥x¹ô¤]¥i
            If m_iPrintCurrType <> 1 And m_iPrintCurrType <> 3 And m_iPrintCurrType <> 4 Then
               m_sEBillingMsg = m_sEBillingMsg & vbCrLf & adoacc1k0("a1k01") & "µLªk²£¥Í LEDES ¹q¤l±b³æ¡I" & vbCrLf & vbCrLf & "**¬üª÷½Ð´Ú©Î¦C¦L¤è¦¡¬°¯Â¥~¹ô¤~¥i²£¥Í¹q¤l±b³æ**"
               m_bEBilling = False
            End If
         End If
         'end 2010/11/9
      End If
      '2009/5/18 end
      
      '¬Û¦P½Ð´Ú¶µ¥Ø¦X¨Ö
      '­Y¨t²ÎÃþ§O¬°FCT, ½Ð´Ú¶µ¥Ø¬°"101", "10199"
      'Modify by Morgan 2004/10/4 ¥[ FCT 715,71599,716,71699,717,71799, S 001
      'Modify by Morgan 2006/5/2 ¥[ FCT 303,30399
      'Modify by Morgan 2006/11/2 ¥[ FCT 60199, 60399, 60599
      'If "" & adoacc1k0.Fields("a1k13").Value = "FCT" And ("" & adoacc1k0.Fields("a1j02").Value = "101" Or "" & adoacc1k0.Fields("a1j02").Value = "10199") Then
      'Modify by Morgan 2007/12/17 ¥[FCT 001,02,S 02
      'Modify by Morgan 2008/5/16 +FCT 308,30899
      'Modify by Morgan 2011/3/10 +FCT 2026,108
      'Modify by Morgan 2011/3/18 +FCT 1012
      'Modified by Morgan 2011/12/28 +FCT 2011
      'Modified by Morgan 2014/8/13 +FCT 102,10299
      'Modified by Morgan 2015/1/29 §ï¨t²Î§O¬°FCT©ÎS®É³£¶µ¥Ø³£¦X¨Ö --³¯ª÷½¬
      'If ("" & adoacc1k0.Fields("a1k13").Value = "FCT" And (adoacc1k0("a1j02") = "102" Or adoacc1k0("a1j02") = "10299" _
         Or adoacc1k0("a1j02") = "2011" Or adoacc1k0("a1j02") = "1012" Or adoacc1k0("a1j02") = "108" Or adoacc1k0("a1j02") = "2026" _
         Or "" & adoacc1k0("a1j02") = "308" Or adoacc1k0("a1j02") = "30899" Or adoacc1k0("a1j02") = "101" Or adoacc1k0("a1j02") = "10199" _
         Or adoacc1k0("a1j02") = "715" Or adoacc1k0("a1j02") = "71599" Or adoacc1k0("a1j02") = "716" Or adoacc1k0("a1j02") = "71699" _
         Or adoacc1k0("a1j02") = "717" Or adoacc1k0("a1j02") = "71799" Or adoacc1k0("a1j02") = "303" _
         Or adoacc1k0("a1j02") = "30399" Or adoacc1k0("a1j02") = "60199" Or adoacc1k0("a1j02") = "60399" _
         Or adoacc1k0("a1j02") = "60599" Or adoacc1k0("a1j02") = "001" Or adoacc1k0("a1j02") = "02")) _
         Or (adoacc1k0("a1k13") = "S" And (adoacc1k0("a1j02") = "001" Or adoacc1k0("a1j02") = "02")) Then
      If "" & adoacc1k0.Fields("a1k13").Value = "FCT" Or "" & adoacc1k0.Fields("a1k13").Value = "S" Then
      'end 2015/1/29
         'Modify By Sindy 2011/3/7
'         If IsNull(adoacc1k0.Fields("fa43").Value) = False Then
'            strCurr = adoacc1k0.Fields("fa43").Value
         If IsNull(adoacc1k0.Fields("fa108").Value) = False Then
            strCurr = adoacc1k0.Fields("fa108").Value
         Else
            strCurr = ""
         End If
         stIDate = "" & adoacc1k0.Fields("IDate").Value 'Added by Morgan 2023/8/29
         strA1J02 = "" & adoacc1k0.Fields("a1j02").Value
         strItemNoFrom = "" & adoacc1k0.Fields("a1l02").Value 'Add by Morgan 2010/10/13
         'Modify by Morgan 2008/7/2 ÁÙ­n§PÂ_½Ð´Ú³æ¸¹¤@¼Ë§_«h¥u¦³¤@­Ó¶µ¥Ø¥B¬Û¦P©Ê½è®É·|¥u¦L¤@±i½Ð´Ú³æ
         'Do While "" & adoacc1k0("a1j02").Value = strA1J02
         Do While "" & adoacc1k0("a1j02").Value = strA1J02 And "" & adoacc1k0("a1k01").Value = strA1K01
            'Added by Morgan 2013/7/30
            '¬Û¦P½Ð´Ú¶µ¥Ø¦ýª÷ÃB¤£¤@©w¬Û¦P,À³¸Ó­n³v¤@Åª¨ú
            dbl_A1L0507FAmt = "" & adoacc1k0.Fields("A1L0507FAmt").Value
            dbl_A1L05FAmt = "" & adoacc1k0.Fields("A1L05FAmt").Value
            'end 2013/7/30
            
            strItemNoTo = "" & adoacc1k0.Fields("a1l02").Value 'Add by Morgan 2010/10/13
            'Modified by Morgan 2012/12/7
            'Select Case strCurr
            '   Case "U"
            '      dblGrpAmt = dblGrpAmt + (Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / Val(adoacc1k0.Fields("a1k10").Value)
            '      dblGrpAmtNoDisc = dblGrpAmtNoDisc + (Val(adoacc1k0.Fields("a1l05").Value)) / Val(adoacc1k0.Fields("a1k10").Value) 'Add by Morgan 2010/11/9
            '   '2009/5/18 modify by sonia
            '   'Case Else
            '   '   dblGrpAmt = dblGrpAmt + (Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value))
            '   Case "N"
            '      dblGrpAmt = dblGrpAmt + (Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value))
            '      dblGrpAmtNoDisc = dblGrpAmtNoDisc + Val(adoacc1k0.Fields("a1l05").Value)  'Add by Morgan 2010/11/9
            '   Case Else
            '      If m_DNCurr = "USD" Then
            '         dblGrpAmt = dblGrpAmt + (Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value))
            '         dblGrpAmtNoDisc = dblGrpAmtNoDisc + Val(adoacc1k0.Fields("a1l05").Value)  'Add by Morgan 2010/11/9
            '      Else
            '         dblGrpAmt = dblGrpAmt + (Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / m_DNRate
            '         dblGrpAmtNoDisc = dblGrpAmtNoDisc + Val(adoacc1k0.Fields("a1l05").Value) / m_DNRate 'Add by Morgan 2010/11/9
            '      End If
            '   '2009/5/18 end
            'End Select
            Select Case m_iPrintCurrType
            Case 3, 4 '3.¯Â¥~¹ô, 4.¥~¹ô+¬üª÷¦X­p
               '·s½Ð´Ú³æ¥~¹ô¤p¼Æ³£±Ë¥h
               If m_bolNewBill Then
                  'Modify By Sindy 2013/1/28
'                  If m_DNCurr = "USD" Then
'                     dblGrpAmt = dblGrpAmt + Fix(Format((Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / Val(adoacc1k0.Fields("a1k10").Value)))
'                     dblGrpAmtNoDisc = dblGrpAmtNoDisc + Fix(Format((Val(adoacc1k0.Fields("a1l05").Value)) / Val(adoacc1k0.Fields("a1k10").Value)))
'                  Else
                     
                     'Modify By Sindy 2013/3/29
                     'dblGrpAmt = dblGrpAmt + Fix(Format((Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / m_DNRate))
                     'dblGrpAmtNoDisc = dblGrpAmtNoDisc + Fix(Format(Val(adoacc1k0.Fields("a1l05").Value) / m_DNRate))
                     dblGrpAmt = dblGrpAmt + Trunc(dbl_A1L0507FAmt)
                     dblGrpAmtNoDisc = dblGrpAmtNoDisc + Trunc(dbl_A1L05FAmt)
                     '2013/3/29 End
                     
'                  End If
                  '2013/1/28 End
               Else
                  'Modify By Sindy 2013/1/28
'                  If m_DNCurr = "USD" Then
'                     dblGrpAmt = dblGrpAmt + (Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / Val(adoacc1k0.Fields("a1k10").Value)
'                     dblGrpAmtNoDisc = dblGrpAmtNoDisc + (Val(adoacc1k0.Fields("a1l05").Value)) / Val(adoacc1k0.Fields("a1k10").Value)
'                  Else
                     
                     'Modify By Sindy 2013/3/29
                     'dblGrpAmt = dblGrpAmt + (Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / m_DNRate
                     'dblGrpAmtNoDisc = dblGrpAmtNoDisc + Val(adoacc1k0.Fields("a1l05").Value) / m_DNRate
                     dblGrpAmt = dblGrpAmt + dbl_A1L0507FAmt
                     dblGrpAmtNoDisc = dblGrpAmtNoDisc + dbl_A1L05FAmt
                     '2013/3/29 End
                     
'                  End If
                  '2013/1/28 End
               End If
            Case Else '1.¯Â¥x¹ô 2.¥x¹ô+¥~¹ô¦X­p
               dblGrpAmt = dblGrpAmt + (Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value))
               dblGrpAmtNoDisc = dblGrpAmtNoDisc + Val(adoacc1k0.Fields("a1l05").Value)
               'Add By Sindy 2013/3/29
               dblGrpAmtFAmt = dblGrpAmtFAmt + dbl_A1L0507FAmt
               dblGrpAmtNoDiscFAmt = dblGrpAmtNoDiscFAmt + dbl_A1L05FAmt
               '2013/3/29 End
            End Select
            'end 2012/12/7
            
            'Added by Morgan 2018/1/15 ¤é¤å¬d¦W²Ä1Ãþ­n³æ¿W¦C
            If strLanguage = "3" And (m_CP01 = "S" Or m_CP01 = "FCT") And strA1J02 = "001" Then
               If str1stNo = "" Then
                  str1stNo = strItemNoFrom
                  adoacc1k0.MoveNext
                  Exit Do
               End If
            End If
            'end 2018/1/15
            adoacc1k0.Fields("IDate") = stIDate 'Added by Morgan 2023/8/29
            adoacc1k0.MoveNext
            If adoacc1k0.EOF Then Exit Do
         Loop
         adoacc1k0.MovePrevious
         
         'Modify by Morgan 2011/3/4
         'strDescription = GetNewDescription("" & adoacc1k0("a1k01").Value, strA1J02, strItemNoFrom, strItemNoTo)
         'FCT °Ó¼Ð¥Ó½Ð³W¶O»¡©ú¯S®í
         If m_CP01 = "FCT" And strA1J02 = "10199" Then
            strDescription = GetNewDescFCT10199("" & adoacc1k0("a1k01").Value, strA1J02, strItemNoFrom, strItemNoTo)
         'Added by Morgan 2018/11/20
         'Modified by Morgan 2020/1/9 +FCT-1013
         ElseIf (m_CP01 = "S" And strA1J02 = "0011") Or (m_CP01 = "FCT" And strA1J02 = "1013") Then
            strExc(0) = "select a1l14 from acc1l0 where a1l01='" & adoacc1k0("a1k01").Value & "' and a1l02='" & adoacc1k0("a1l02").Value & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If (m_CP01 = "S" And strA1J02 = "0011") Then
                  strDescription = " (NTD500 x " & RsTemp("a1l14") & "­Ó¶W¹L)"
               ElseIf (m_CP01 = "FCT" And strA1J02 = "1013") Then
                  strDescription = " (NTD80 x " & RsTemp("a1l14") & ")"
               End If
            End If
         'end 2018/11/20
         'Added by Morgan 2024/9/16 Y25061¤Î¨äÃöÁp½s¸¹¡BY54470¡BY54469¡BY55363 ¥N²z¶OÄæ¦ì¤£Åã¥Ü¥»©Ò±µ¬¢³æ¶ñ¼gªº¥x¹ô­pºâ¦¡--¨H¨Î¿o
         ElseIf InStr("Y54470000,Y54469000,Y55363000", m_A1k28) > 0 Or Left(m_A1k28, 6) = "Y25061" And Right(strA1J02, 2) <> "99" Then
            strDescription = ""
         'end 2024/9/16
         Else
            strDescription = GetNewDescription("" & adoacc1k0("a1k01").Value, strA1J02, strItemNoFrom, strItemNoTo)
         End If
      '¨ä¥L·ÓÂÂ
      Else
         'Add By Sindy 2013/5/21 Ex.X10206995
         If m_iPrintCurrType = 1 Or m_iPrintCurrType = 2 Then
            dblGrpAmtFAmt = dbl_A1L0507FAmt
            dblGrpAmtNoDiscFAmt = dbl_A1L05FAmt
         End If
         '2013/5/21 End
         If adoacc1k0.Fields("a1l07").Value <> 0 Then
            If (100 - Val((Format(adoacc1k0.Fields("a1l07").Value / adoacc1k0.Fields("a1l05").Value * 100, DAmount)))) = 100 Then
               strDescription = ""
            Else
               'Modify by Sindy 2013/4/9 ·í3.¯Â¥~¹ô®ÉÅã¥Ü½Ð´Ú¹ôª÷ÃB
               'Modified by Morgan 2016/8/5 Y45204000 ©ú²Ó­n¦L¥x¹ô
               'If m_iPrintCurrType = "3" Then
               If m_iPrintCurrType = "3" And m_bSpecialNew2 = False Then
               'end 2016/8/5
                  strDescription = " ( " & m_DNCurr & " " & PUB_ChgFormat(adoacc1k0.Fields("A1L05FAmt").Value, True) & " x " & (100 - Val((Format(adoacc1k0.Fields("a1l07").Value / adoacc1k0.Fields("a1l05").Value * 100, DAmount)))) & "% )"
               Else
                  strDescription = " ( NTD " & PUB_ChgFormat(adoacc1k0.Fields("a1l05").Value, True) & " x " & (100 - Val((Format(adoacc1k0.Fields("a1l07").Value / adoacc1k0.Fields("a1l05").Value * 100, DAmount)))) & "% )"
               End If
               '2013/4/9 End
            End If
         Else
            strDescription = ""
         End If
      End If
      
      m_ItemTDesc = strDescription & strFMPFee99RMB 'Add by Morgan 2013/10/24
      
      'Modify By Sindy 2011/3/7
      If CheckSys(m_CP01) = "2" Or CheckSys(m_CP01) = "6" Then
         If IsNull(adoacc1k0.Fields("fa108").Value) = False Then
             strCurr = adoacc1k0.Fields("fa108").Value
         Else
            strCurr = ""
         End If
      '2011/3/7 End
      Else
         If IsNull(adoacc1k0.Fields("fa43").Value) = False Then
             strCurr = adoacc1k0.Fields("fa43").Value
         Else
            strCurr = ""
         End If
      End If
      
      'Memo by Morgan 2012/12/14 ¤U­±µ{§Ç»Ý«O¯d¦]¬°¦C¦LÂÂ½Ð´Ú³æ­n¥Î
      'Add by Morgan 2006/9/29 P,Tªº ¥xÆW®× ­Y¨S¦³«ü©w®É¥Î ¬üª÷ ½Ð´Ú
      bolPTUSDCase = False
      If (Left(m_CP01, 1) = "P" Or Left(m_CP01, 1) = "T") Then
         If GetPrjNation1(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04) = "000" Then
            'Remove by Morgan 2008/3/21 --¤£¹w³]¤F
            'If strCurr = "" Then
            '   strCurr = "U"
            'End If
            'end 2008/3/21
            'Modified by Morgan 2012/12/7
            'If strCurr <> "N" Then
            If m_iPrintCurrType <> 1 Then
            'end 2012/12/7
               bolPTUSDCase = True
            End If
         End If
      End If
      
      If m_b2Printer Then
         Printer.CurrentX = 0 + intDefault
         Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
      End If
      'Removed by Morgan 2022/8/4
      'If m_b2Picture Then
      '   Picture1.CurrentX = (0 + intDefault) * douExtRate
      '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
      'End If
      'end 2022/8/4
      Select Case strLanguage
         Case "1" '¤¤¤å
            iItemNo = iItemNo + 1
            If Not IsNull(adoacc1k0.Fields("a1j03").Value) Then
               'Modify by Morgan 2006/11/17
               'If m_CP01 = "T" Then
               If Left(m_CP01, 1) = "T" Then
                  strA1J06 = "¡@¡¸¡@" & adoacc1k0.Fields("a1j03").Value
               Else
                  'Modified by Morgan 2015/9/2 ¦P®×¥ó©Ê½è³W¶OªA°È¶O¦L¦P¤@¦æ«á§Ç¦¸·|¤£¥¿½T¬G¨ú®ø
                  'strA1J06 = iItemNo & ". " & adoacc1k0.Fields("a1j03").Value
                  If m_CP01 = "FCP" Then
                     strA1J06 = "¡· " & adoacc1k0.Fields("a1j03").Value
                  Else
                     strA1J06 = "" & adoacc1k0.Fields("a1j03").Value
                  End If
                  'end 2015/9/2
               End If
               'Modified by Lydia 2015/04/09
               If m_bolChiDB And bol_ChiDB Then
                  strA1J06 = "" & adoacc1k0.Fields("a1j03").Value
               End If
               
               m_ClearItemDesc = strA1J06 'Add by Morgan 2010/12/7
               m_ItemDesc = strA1J06 & strDescription 'Add by Morgan 2010/11/9
               m_ItemHDesc = strA1J06 'Add by Morgan 2013/10/24
               
               If CountLength(strA1J06 & strDescription) <= 70 Then
                  PutData strA1J06 & strDescription, intCounter, 0, m_DetailTopStart
                  'Added by Lydia 2015/04/09
                  If m_bolChiDB And bol_ChiDB Then
                     PutData strA1J06 & strDescription, intCounter, 4, m_DetailTopStart
                     PutData adoacc1k0.Fields("a1k01").Value, intCounter, 5, m_DetailTopStart
                  End If
                    
               Else
                  PrintDropLine strA1J06 & strDescription, "", intRow + intCounter, 70, -400
               End If
            End If
            
         Case "2" '­^¤å
            If Not IsNull(adoacc1k0.Fields("a1j04").Value) Then
               If adoacc1k0.Fields("a1j04").Value = "Official fees" Then
                  If "" & adoacc1k0.Fields("a1l06").Value <> "" Then
                     'Modify by Sindy 2013/7/4
                     'm_ClearItemDesc = adoacc1k0.Fields("a1j04").Value & PUB_GetA1L06Text("" & adoacc1k0.Fields("a1l06").Value) 'Add by Morgn 2010/12/7
                     m_ClearItemDesc = adoacc1k0.Fields("a1j04").Value & PUB_GetA1L06Text("" & adoacc1k0.Fields("a1l06").Value) & IIf(strFMPFee99RMB <> "", strFMPFee99RMB, "") 'Add by Morgn 2010/12/7
                     'm_ItemDesc = adoacc1k0.Fields("a1j04").Value & PUB_GetA1L06Text("" & adoacc1k0.Fields("a1l06").Value) & strDescription 'Add by Morgan 2010/11/9
                     m_ItemDesc = adoacc1k0.Fields("a1j04").Value & PUB_GetA1L06Text("" & adoacc1k0.Fields("a1l06").Value) & strDescription & IIf(strFMPFee99RMB <> "", strFMPFee99RMB, "") 'Add by Morgan 2010/11/9
                     m_ItemHDesc = adoacc1k0.Fields("a1j04").Value & PUB_GetA1L06Text("" & adoacc1k0.Fields("a1l06").Value) 'Add by Morgan 2013/10/24
                     m_ItemXDesc = m_ItemHDesc 'Added by Morgan 2013/10/29
                     
                     '2013/7/4 END
                     'Modify By Sindy 2013/1/28
                     'PutData adoacc1k0.Fields("a1j04").Value & PUB_GetA1L06Text("" & adoacc1k0.Fields("a1l06").Value) & strDescription, intCounter, 0, m_DetailTopStart
                     PutData adoacc1k0.Fields("a1j04").Value & PUB_GetA1L06Text("" & adoacc1k0.Fields("a1l06").Value) & strDescription & IIf(strFMPFee99RMB <> "", strFMPFee99RMB, ""), intCounter, 0, m_DetailTopStart
                     '2013/1/28 End
                  Else
                     'Modify by Sindy 2013/7/4
                     'm_ClearItemDesc = adoacc1k0.Fields("a1j04").Value 'Add by Morgn 2010/12/7
                     m_ClearItemDesc = adoacc1k0.Fields("a1j04").Value & IIf(strFMPFee99RMB <> "", strFMPFee99RMB, "") 'Add by Morgn 2010/12/7
                     'm_ItemDesc = adoacc1k0.Fields("a1j04").Value & strDescription 'Add by Morgan 2010/11/9
                     m_ItemDesc = adoacc1k0.Fields("a1j04").Value & strDescription & IIf(strFMPFee99RMB <> "", strFMPFee99RMB, "") 'Add by Morgan 2010/11/9
                     m_ItemHDesc = adoacc1k0.Fields("a1j04").Value 'Add by Morgan 2013/10/24
                     
                     '2013/7/4 END
                     '2010/11/1 modify by sonia ³W¶O­Y¦³¿é§ì§é¦©¤]­n¦L
                     'PutData adoacc1k0.Fields("a1j04").Value, intCounter, 0, m_DetailTopStart
                     'Modify By Sindy 2013/1/28
                     'PutData adoacc1k0.Fields("a1j04").Value & strDescription, intCounter, 0, m_DetailTopStart
                     PutData adoacc1k0.Fields("a1j04").Value & strDescription & IIf(strFMPFee99RMB <> "", strFMPFee99RMB, ""), intCounter, 0, m_DetailTopStart
                     '2013/1/28 End
                  End If
               Else
                  If IsNull(adoacc1k0.Fields("a1j05").Value) Then
                     'Modify by Sindy 2013/7/4
                     'm_ClearItemDesc = adoacc1k0.Fields("a1j04").Value 'Add by Morgan 2010/12/7
                     m_ClearItemDesc = adoacc1k0.Fields("a1j04").Value & IIf(strFMPFee99RMB <> "", strFMPFee99RMB, "") 'Add by Morgan 2010/12/7
                     'm_ItemDesc = adoacc1k0.Fields("a1j04").Value & strDescription 'Add by Morgan 2010/11/9
                     m_ItemDesc = adoacc1k0.Fields("a1j04").Value & strDescription & IIf(strFMPFee99RMB <> "", strFMPFee99RMB, "") 'Add by Morgan 2010/11/9
                     m_ItemHDesc = adoacc1k0.Fields("a1j04").Value 'Add by Morgan 2013/10/24
                     
                     '2013/7/4 END
                     If CountLength(adoacc1k0.Fields("a1j04").Value & strDescription) <= 70 Then
                        'Modify By Sindy 2013/4/23
                        'PutData adoacc1k0.Fields("a1j04").Value & strDescription, intCounter, 0, m_DetailTopStart
                        PutData adoacc1k0.Fields("a1j04").Value & strDescription & IIf(strFMPFee99RMB <> "", strFMPFee99RMB, ""), intCounter, 0, m_DetailTopStart
                        '2013/4/23 End
                     Else
                        'Modify By Sindy 2013/4/23
                        'PutData adoacc1k0.Fields("a1j04").Value, intCounter, 0, m_DetailTopStart
                        PutData adoacc1k0.Fields("a1j04").Value & IIf(strFMPFee99RMB <> "", strFMPFee99RMB, ""), intCounter, 0, m_DetailTopStart
                        '2013/4/23 End
                        If Trim(strDescription) <> "" Then
                           intCounter = intCounter + 1
                           PrintDropLine Trim(strDescription), "", intRow + intCounter, 70, -400
                        End If
                     End If
                  Else
                     Select Case adoacc1k0.Fields("a1j02").Value
                        Case "601", "605" '»âÃÒ¤Î¦~¶O, ¦~¶O
                           'Modified by Morgan 2013/5/3 +P
                           'Modified by Morgan 2015/12/22 +CFP
                           If adoacc1k0.Fields("a1k13").Value = "FCP" Or adoacc1k0.Fields("a1k13").Value = "P" Or adoacc1k0.Fields("a1k13").Value = "CFP" Then
                           'end 2013/5/3
                              
                              'Modified by Morgan 2024/11/4 §ï¨ç¼Æ¦@¥Î
                              'If adoquery.State = adStateOpen Then
                              '   adoquery.Close
                              'End If
                              ''ªì©l¤ÆÅÜ¼Æ
                              'intYear = 0: strAnnuity = "": strSendDate = ""
                              'adoquery.CursorLocation = adUseClient
                              'adoquery.Open "select cp27, pa72, pa73 from caseprogress, patent where cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and cp60 = '" & adoacc1k0.Fields("a1k01").Value & "' And CP10='" & adoacc1k0.Fields("a1j02").Value & "' ", adoTaie, adOpenStatic, adLockReadOnly
                              'If adoquery.RecordCount <> 0 Then
                              '   If IsNull(adoquery.Fields("pa73").Value) = False Then
                              '      strTemp1 = Split(UCase(adoquery.Fields("pa72").Value), ",")
                              '      strTemp2 = Split(UCase(adoquery.Fields("pa73").Value), ",")
                              '      For i = 0 To UBound(strTemp2)
                              '         If Val(strTemp2(i)) = adoquery.Fields("cp27").Value Then
                              '            If Val(strSendDate) <> adoquery.Fields("cp27").Value Then
                              '               Select Case Val(strTemp1(i))
                              '                  Case 1
                              '                     strAnnuity = strTemp1(i) & "st to "
                              '                  Case 2
                              '                     strAnnuity = strTemp1(i) & "nd to "
                              '                  Case 3
                              '                     strAnnuity = strTemp1(i) & "rd to "
                              '                  Case Else
                              '                     strAnnuity = strTemp1(i) & "th to "
                              '               End Select
                              '               strSendDate = strTemp2(i)
                              '            End If
                              '            s = i
                              '            intYear = intYear + 1
                              '         End If
                              '      Next i
                              '   End If
                              'End If
                              'adoquery.Close
                              'If intYear > 1 Then
                              '   Select Case Val(strTemp1(s))
                              '      Case 1
                              '         strAnnuity = strAnnuity & strTemp1(s) & "st"
                              '      Case 2
                              '         strAnnuity = strAnnuity & strTemp1(s) & "nd"
                              '      Case 3
                              '         strAnnuity = strAnnuity & strTemp1(s) & "rd"
                              '      Case Else
                              '         strAnnuity = strAnnuity & strTemp1(s) & "th"
                              '   End Select
                              'ElseIf strAnnuity <> "" Then
                              '   strAnnuity = Mid(strAnnuity, 1, Len(strAnnuity) - 4)
                              'End If
                              
                              'strA1J06 = "" & adoacc1k0.Fields("a1j04").Value
                              'If adoacc1k0.Fields("a1j02").Value = "601" Then
                              '   strA1J06 = Replace(strA1J06, "1st", strAnnuity)
                              'Else
                              '   strA1J06 = strA1J06 & strAnnuity
                              'End If
                              
                              strA1J06 = "" & adoacc1k0.Fields("a1j04").Value
                              strA1J06 = PUB_GetAnnuityDesc(adoacc1k0.Fields("a1k01").Value, adoacc1k0.Fields("a1j02").Value, strA1J06)
                              'end 2024/11/1
                              m_ClearItemDesc = strA1J06 'Add by Morgan 2010/12/7
                              m_ItemDesc = strA1J06 'Add by Morgan 2010/11/9
                              m_ItemHDesc = strA1J06 'Add by Morgan 2013/10/24
                              
                              PutData strA1J06, intCounter, 0, m_DetailTopStart
                              'end 2007/12/5
                           Else
                              m_ClearItemDesc = adoacc1k0.Fields("a1j04").Value 'Add by Morgan 2010/12/7
                              m_ItemDesc = adoacc1k0.Fields("a1j04").Value 'Add by Morgan 2010/11/9
                              m_ItemHDesc = adoacc1k0.Fields("a1j04").Value 'Add by Morgan 2013/10/24
                              PutData adoacc1k0.Fields("a1j04").Value, intCounter, 0, m_DetailTopStart
                           End If
                        Case Else
                           m_ClearItemDesc = adoacc1k0.Fields("a1j04").Value 'Add by Morgan 2010/12/7
                           m_ItemDesc = adoacc1k0.Fields("a1j04").Value 'Add by Morgan 2010/11/9
                           m_ItemHDesc = adoacc1k0.Fields("a1j04").Value 'Add by Morgan 2013/10/24
                           
                           PutData adoacc1k0.Fields("a1j04").Value, intCounter, 0, m_DetailTopStart
                           
                     End Select
                  End If
               End If
               If IsNull(adoacc1k0.Fields("a1j05").Value) = False Then
                  intCounter = intCounter + 1
                  If m_b2Printer Then
                     Printer.CurrentX = 0 + intDefault
                     Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
                  End If
                  'Removed by Morgan 2022/8/4
                  'If m_b2Picture Then
                  '   Picture1.CurrentX = (0 + intDefault) * douExtRate
                  '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
                  'End If
                  'end 2022/8/4
                  If IsNull(adoacc1k0.Fields("a1j06").Value) Then
                     m_ClearItemDesc = m_ClearItemDesc & " " & adoacc1k0.Fields("a1j05").Value 'Add by Morgan 2010/12/7
                     m_ItemDesc = m_ItemDesc & " " & adoacc1k0.Fields("a1j05").Value & strDescription 'Add by Morgan 2010/11/9
                     m_ItemHDesc = m_ItemHDesc & " " & adoacc1k0.Fields("a1j05").Value  'Add by Morgan 2013/10/24
                     
                     
                     If CountLength(adoacc1k0.Fields("a1j05").Value & strDescription) <= 70 Then
                        PutData adoacc1k0.Fields("a1j05").Value & strDescription, intCounter, 0, m_DetailTopStart
                     Else
                        PutData adoacc1k0.Fields("a1j05").Value, intCounter, 0, m_DetailTopStart
                        If Trim(strDescription) <> "" Then
                           intCounter = intCounter + 1
                           PrintDropLine Trim(strDescription), "", intRow + intCounter, 70, -400
                        End If
                     End If
                  Else
                     m_ClearItemDesc = m_ClearItemDesc & " " & adoacc1k0.Fields("a1j05").Value 'Add by Morgan 2010/12/7
                     m_ItemDesc = m_ItemDesc & " " & adoacc1k0.Fields("a1j05").Value  'Add by Morgan 2010/11/9
                     m_ItemHDesc = m_ItemHDesc & " " & adoacc1k0.Fields("a1j05").Value  'Add by Morgan 2013/10/24
                     
                     PutData adoacc1k0.Fields("a1j05").Value, intCounter, 0, m_DetailTopStart
                  End If
               End If
               If IsNull(adoacc1k0.Fields("a1j06").Value) = False Then
                  intCounter = intCounter + 1
                  If m_b2Printer Then
                     Printer.CurrentX = 0 + intDefault
                     Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
                  End If
                  'Removed by Morgan 2022/8/4
                  'If m_b2Picture Then
                  '   Picture1.CurrentX = (0 + intDefault) * douExtRate
                  '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
                  'End If
                  'end 2022/8/4
                  m_ClearItemDesc = m_ClearItemDesc & " " & adoacc1k0.Fields("a1j06").Value 'Add by Morgan 2010/12/7
                  m_ItemDesc = m_ItemDesc & " " & adoacc1k0.Fields("a1j06").Value & strDescription  'Add by Morgan 2010/11/9
                  m_ItemHDesc = m_ItemHDesc & " " & adoacc1k0.Fields("a1j06").Value  'Add by Morgan 2013/10/24
                  
                  If CountLength(adoacc1k0.Fields("a1j06").Value & strDescription) <= 70 Then
                     PutData adoacc1k0.Fields("a1j06").Value & strDescription, intCounter, 0, m_DetailTopStart
                  Else
                     PutData adoacc1k0.Fields("a1j06").Value, intCounter, 0, m_DetailTopStart
                     
                     If Trim(strDescription) <> "" Then
                        intCounter = intCounter + 1
                        PrintDropLine Trim(strDescription), "", intRow + intCounter, 70, -400
                     End If
                  End If
               End If
                  
               Set201ItemDesc 'Added by Morgan 2024/5/9 µ{§Ç¤Ó¤j¡Aµ{¦¡½X²¾¨ì¨ç¼Æ
               
               'Added by Morgan 2019/10/2
               'Y54225B10 Syngenta ªº½Ð´Ú³æ¡]LEDES¡^½Ð´Ú¶µ¥Ø106¥D±iÀu¥ýÅvªº­^¤å±Ô­z«á¤è¥[¤W©T©wª÷ÃB USD35x[Àu¥ýÅv¼Æ¶q]
               If adoacc1k0.Fields("a1k13").Value = "FCP" And m_A1k28 = "Y54225B10" And adoacc1k0.Fields("a1j02") = "106" Then
                  strExc(0) = "select count(*) from pridate where pd01='" & adoacc1k0.Fields("a1k13") & "' and pd02='" & adoacc1k0.Fields("a1k14") & "' and pd03='" & adoacc1k0.Fields("a1k15") & "' and pd04='" & adoacc1k0.Fields("a1k16") & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If RsTemp(0) > 0 Then
                        m_ItemDesc = m_ItemDesc & "(USD35x" & RsTemp(0) & ")"
                     End If
                  End If
               End If
               'end 2019/10/2
               
               'Added by Morgan 2019/8/30
               'Y52206 (Teradyne (Asia) Pte Ltd) ½Ð´Ú±Ô­z¯S®í»Ý¨D--Tim
               strExc(1) = ""
               If adoacc1k0.Fields("a1k03").Value = "Y52206000" And (adoacc1k0.Fields("a1k13").Value = "FCP" Or adoacc1k0.Fields("a1k13").Value = "P") Then
                  If adoacc1k0.Fields("a1j02") = "203" Then
                     If Is203Amendent(adoacc1k0.Fields("a1k01")) = True Then
                        strExc(1) = "[Amendment]"
                     Else
                        strExc(1) = "[Filing billing]"
                     End If
                  Else
                     Select Case "" & adoacc1k0.Fields("a1j02")
                     Case "03", "101", "102", "103", "106", "201", "202", "209", "307"
                        strExc(1) = "[Filing billing]"
                     'Modified by Morgan 2024/11/13 +447¦A¼f¬d¥[³t¼f¬d
                     Case "416", "435", "422", "431", "447"
                        strExc(1) = "[Examination/Search]"
                     Case "1002", "1202", "107", "205", "404", "407", "408"
                        strExc(1) = "[OA billing]"
                     Case "927"
                        strExc(1) = "[OA translation]"
                     Case "204"
                        strExc(1) = "[Amendment]"
                     Case "601", "926", "402"
                        strExc(1) = "[Grant billing]"
                     Case "02", "401", "908", "929"
                        strExc(1) = "[Transfer/Docket/Correspondence/Disbursements]"
                     End Select
                  End If
               End If
               m_ItemDesc = strExc(1) & m_ItemDesc
               m_ItemHDesc = strExc(1) & m_ItemHDesc
               'end 2019/8/30
               
               'Added by Morgan 2024/4/17
               'Y48110B20 (BOZICEVIC, FIELD & FRANCIS LLP) ½Ð´Ú±Ô­z«e­±¥[µo¤å¤é--Kahn
               If adoacc1k0.Fields("a1k03").Value = "Y48110B20" And (adoacc1k0.Fields("a1k13").Value = "FCP" Or adoacc1k0.Fields("a1k13").Value = "P") Then
                  '*¤¤»¡½Ð´Ú (µo©ú¥Ó½Ð¡BÂ½Ä¶¡B¥´¦r¡B¸É¤å¥ó¡B¥D±iÀu¥ýÅvµ¥½Ð´Ú)¡Aµo¤å¤é½Ð§ìµo©ú¥Ó½Ð¤§µo¤å¤é
                  strExc(2) = adoacc1k0.Fields("a1j02")
                  If Right(strExc(2), 2) = "99" Then
                     strExc(2) = Left(strExc(2), Len(strExc(2)) - 1)
                  End If
                  If strExc(2) = "106" Or strExc(2) = "201" Or strExc(2) = "209" Or strExc(2) = "210" Or strExc(2) = "202" Then
                     strExc(3) = " and cp10 in (" & NewCasePtyList & ")"
                  Else
                     strExc(3) = " and cp10='" & strExc(2) & "'"
                  End If
                  strExc(0) = "select sqldatew(cp27) DDate, 1 srt from caseprogress where cp60='" & adoacc1k0.Fields("a1k01").Value & "'" & strExc(3) & _
                     " union select sqldatew(min(cp27)) DDate, 2 srt from caseprogress where cp60='" & adoacc1k0.Fields("a1k01").Value & "' and cp09<'C' order by 2,1"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     m_ItemDesc = Format(RsTemp("DDate"), "mm/dd/yyyy") & ": " & m_ItemDesc
                     m_ItemHDesc = Format(RsTemp("DDate"), "mm/dd/yyyy") & ": " & m_ItemHDesc
                  End If
               End If
               'end 2024/4/17
            End If
            
         Case "3" '¤é¤å
            If Not IsNull(adoacc1k0.Fields("a1j16").Value) Then
               'Add by Morgan 2005/4/1
               'Added by Morgan 2018/1/15 ¤é¤å¬d¦W
               If strLanguage = "3" And (m_CP01 = "S" Or m_CP01 = "FCT") And strA1J02 = "001" Then
                  'Modified by Morgan 2018/3/12 ­×§ï¦³§é¦©®æ¦¡--³¯ª÷½¬
                  'Modified by Morgan 2022/7/27
                  'strA1J06 = "¤â“g®Æ ("
                  strA1J06 = PUB_GetUniText(Me.Name, "¶O¥Î") & " ("
                  'end 2022/7/27
                  If str1stNo = strItemNoFrom Then
                     'Modified by Morgan 2022/7/27
                     'strA1J06 = strA1J06 & "1üP¤À¥Ø"
                     strA1J06 = strA1J06 & "1" & PUB_GetUniText(Me.Name, "Ãþ§O")
                     'end 2022/7/27
                  Else
                     'Modified by Morgan 2022/7/27
                     'strA1J06 = strA1J06 & "2üP¤À¥Ø¥H­°"
                     strA1J06 = strA1J06 & "2" & PUB_GetUniText(Me.Name, "Ãþ§O") & "¥H­°"
                     'end 2022/7/27
                  End If
                  strA1J06 = strA1J06 & "¡þNT$" & Format(adoacc1k0.Fields("a1l05").Value, DDollar)
                  
                  If adoacc1k0.Fields("a1l07") > 0 Then
                     strA1J06 = strA1J06 & " x " & Format((1 - (adoacc1k0.Fields("a1l07") / adoacc1k0.Fields("a1l05").Value)) * 100, DAmount) & "%"
                  End If
                  If str1stNo <> strItemNoFrom Then
                     'Modified by Morgan 2022/7/27
                     'strA1J06 = strA1J06 & " x " & (Val(strItemNoTo) - Val(strItemNoFrom) + 1) & "üP¤À"
                     strA1J06 = strA1J06 & " x " & (Val(strItemNoTo) - Val(strItemNoFrom) + 1) & PUB_GetUniText(Me.Name, "°Ï") & "¤À"
                     'end 2022/7/27
                  End If
                  strA1J06 = strA1J06 & ")"
                  strDescription = ""
                  'end 2018/3/12
               Else
               'end 2018/1/15
                  strA1J06 = adoacc1k0.Fields("a1j16").Value
               End If
               Select Case adoacc1k0.Fields("a1j02").Value
                  Case "601", "605" '»âÃÒ¤Î¦~¶O, ¦~¶O
                     'Modified by Morgan 2014/10/27 +P
                     If adoacc1k0.Fields("a1k13").Value = "FCP" Or adoacc1k0.Fields("a1k13").Value = "P" Then
                        
                        'Modified by Morgan 2024/11/4 §ï¨ç¼Æ¦@¥Î
                        'If adoquery.State = adStateOpen Then
                        '   adoquery.Close
                        'End If
                        'Dim iStart As Integer, iEnd As Integer
                        ''ªì©l¤ÆÅÜ¼Æ
                        'intYear = 0: strAnnuity = "": strSendDate = "": iStart = -1: iEnd = -1
                        'adoquery.CursorLocation = adUseClient
                        ''­×§ï½Ð´Ú³æ¸¹ªºÅÜ¼Æ
                        'adoquery.Open "select cp27, pa72, pa73 from caseprogress, patent where cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and cp60 = '" & adoacc1k0.Fields("a1k01").Value & "' And CP10='" & adoacc1k0.Fields("a1j02").Value & "' ", adoTaie, adOpenStatic, adLockReadOnly
                        'If adoquery.RecordCount <> 0 Then
                        '   If "" & adoquery.Fields("pa73") <> "" Then
                        '      strTemp1 = Split("" & adoquery.Fields("pa72"), ",")
                        '      strTemp2 = Split("" & adoquery.Fields("pa73"), ",")
                        '      If UBound(strTemp1) = UBound(strTemp2) Then
                        '         For i = 0 To UBound(strTemp2)
                        '            If Val(strTemp2(i)) = adoquery.Fields("cp27").Value Then
                        '               strAnnuity = strTemp1(i)
                        '               If iStart = -1 Then iStart = i
                        '               intYear = intYear + 1
                        '            ElseIf Val(strTemp2(i)) > adoquery.Fields("cp27").Value Then
                        '               Exit For
                        '            End If
                        '         Next i
                        '         iEnd = i - 1
                        '      End If
                        '   End If
                        'End If
                        'adoquery.Close
                        'If iStart <> -1 Then
                        '   If iEnd > iStart Then
                        '      strAnnuity = strTemp1(iStart) & " ~ " & strTemp1(iEnd)
                        '   Else
                        '      strAnnuity = strTemp1(iStart)
                        '   End If
                        '   'Modify by Morgan 2005/7/19 ¥[¦Ò¼{601
                        '   'strA1J06 = Replace(strA1J06, "²Ä  ¦~", "²Ä " & strAnnuity & " ¦~")
                        '   If adoacc1k0.Fields("a1j02").Value = "601" Then
                        '      strA1J06 = Replace(strA1J06, "²Ä 1 ¦~", "²Ä " & strAnnuity & " ¦~")
                        '   Else
                        '      strA1J06 = Replace(strA1J06, "²Ä  ¦~", "²Ä " & strAnnuity & " ¦~")
                        '   End If
                        'End If
                        strA1J06 = PUB_GetAnnuityDesc(adoacc1k0.Fields("a1k01").Value, adoacc1k0.Fields("a1j02").Value, strA1J06, 3)
                        'end 2024/11/4
                     End If
               End Select
               
               m_ClearItemDesc = strA1J06 'Add by Morgan 2010/12/7
               m_ItemDesc = strA1J06 & strDescription  'Add by Morgan 2010/11/9
               m_ItemHDesc = strA1J06   'Add by Morgan 2013/10/24
               
               'Add By Sindy 2013/1/28
               If Trim(strA1J06) = "¬F©²®Æª÷" And Trim(strFMPFee99RMB) <> "" Then
                  strDescription = strDescription & strFMPFee99RMB
                  'Add by Sindy 2013/7/4
                  m_ClearItemDesc = m_ClearItemDesc & strFMPFee99RMB
                  m_ItemDesc = m_ItemDesc & strFMPFee99RMB
                  '2013/7/4 END
               End If
               '2013/1/28 End
               If CountLength(strA1J06 & strDescription) <= 70 Then
                  If m_b2Printer Then
                     'Modified by Morgan 2022/8/4
                     'Printer.Print strA1J06 & strDescription
                     PUB_PrintUnicodeText strA1J06 & strDescription, Printer.CurrentX, Printer.CurrentY, 0
                     'end 2022/8/4
                  End If
                  'Removed by Morgan 2022/8/4
                  'If m_b2Picture Then
                  '   Picture1.Print strA1J06 & strDescription
                  'End If
                  'end 2022/8/4
               Else
                  'Modify by Morgan 2007/10/3 ­Y½Ð´Ú¶µ¥Ø¨S¶W¹L®É³æ¿W¦L
                  If CountLength(strA1J06) <= 70 Then
                     If m_b2Printer Then
                        'Modified by Morgan 2022/8/4
                        'Printer.Print strA1J06
                        PUB_PrintUnicodeText strA1J06, Printer.CurrentX, Printer.CurrentY, 0
                        'end 2022/8/4
                     End If
                     'Removed by Morgan 2022/8/4
                     'If m_b2Picture Then
                     '   Picture1.Print strA1J06
                     'End If
                     'end 2022/8/4
                     intCounter = intCounter + 1
                     PrintDropLine strDescription, "", intRow + intCounter, 70, -400
                  Else
                     PrintDropLine strA1J06 & strDescription, "", intRow + intCounter, 70, -400
                  End If
                  'end 2007/10/3
               End If
            End If
      End Select
      
      'Added by Morgan 2025/10/3
      '½Ð´Ú¹ï¶HY56199 Coupang Corpªº¤¤¶¡µ{§Ç­n±ahourly rate 125©M®É¼Æ
      If m_A1k28 = "Y56199000" And (adoacc1k0.Fields("a1k13").Value = "FCP" Or adoacc1k0.Fields("a1k13").Value = "P") And InStr("'1202','1002','1201','205','204','107','422','431','407','408'", "'" & adoacc1k0("a1j02") & "'") > 0 Then
         strExc(0) = "select cp113 from caseprogress where cp60='" & adoacc1k0("a1k01") & "' and cp10='" & adoacc1k0("a1j02") & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            m_ItemDesc = m_ItemDesc & " (Hourly Rate : USD 125 ; Amount of Hour : " & RsTemp(0) & ")"
         End If
      End If
      'end 2025/10/3
               
      'Added by Morgan 2024/2/17 ¤é±M¦¬¤å®×¥ó·s®×Â½Ä¶201¦³¬Û¦ü®×®É«á­±¦h±a¤@¬q»¡©ú--³¯·¶ªÚ
      'Modified by Morgan 2024/3/25 X3029900 ¿n¤ô¤Æ¾Ç P,FCPªº201Â½Ä¶¶O­n±a­ì¤å¦r¼Æ --Mable
      'Modified by Morgan 2024/6/24 +X4612500 ¤jª÷ P,FCPªº201Â½Ä¶¶O­n±a­ì¤å¦r¼Æ --Mable
      If (adoacc1k0.Fields("a1k13").Value = "FCP" Or adoacc1k0.Fields("a1k13").Value = "P") And adoacc1k0.Fields("a1j02") = "201" Then
         strExc(0) = "select tf23,tf19,tf20,pa77 from caseprogress,transfee,staff,patent where cp60='" & adoacc1k0.Fields("a1k01") & "' and cp10='" & adoacc1k0.Fields("a1j02") & "' and tf01(+)=cp09 and st01(+)=cp13 and st93='J21' and pa01(+)=substr(tf20,1,length(tf20)-9) and pa02(+)=substr(tf20,-9,6) and pa03(+)=substr(tf20,-3,1) and pa04(+)=substr(tf20,-2)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Select Case strLanguage
               Case "2" '­^¤å
                  'Added by Morgan 2024/3/25
                  If InStr(m_CustNoList, "X3029900") > 0 Or InStr(m_CustNoList, "X4612500") > 0 Then
                     intI = InStr(m_ItemDesc, ")")
                     If intI > 0 Then
                        m_ItemDesc = Left(m_ItemDesc, intI - 1) & "; Total " & Format(RsTemp(0), "#,###") & " words" & Mid(m_ItemDesc, intI)
                     End If
                  End If
                  'end 2024/3/25
                  If Not IsNull(RsTemp("tf20")) Then
                     m_ItemDesc = m_ItemDesc & vbCrLf & vbCrLf & "* The specification for this case has overlapping content with the previous case (Your Ref. " & RsTemp("pa77") & "), so an appropriate discount has already been applied to the translation fee." & vbCrLf
                  
                  'Added by Morgan 2024/6/24 µL¬Û¦ü®×»¡©ú(±a¤é¤å)
                  ElseIf InStr(m_CustNoList, "X4612500") > 0 Then
                     m_ItemDesc = m_ItemDesc & vbCrLf & vbCrLf & PUB_GetUniText(Me.Name, "·s®×Â½Ä¶µL¬Û¦ü®×»¡©ú")
                  'end 2024/6/24
                  End If
               Case "3" '¤é¤å
                  'Added by Morgan 2024/3/25
                  If InStr(m_CustNoList, "X3029900") > 0 Or InStr(m_CustNoList, "X4612500") > 0 Then
                     intI = InStr(m_ItemDesc, ")")
                     If intI > 0 Then
                        m_ItemDesc = Left(m_ItemDesc, intI - 1) & "; ­p" & Format(RsTemp(0), "#,###") & "¦r" & Mid(m_ItemDesc, intI)
                     End If
                  End If
                  'end 2024/3/25
                  If Not IsNull(RsTemp("tf20")) Then
                     m_ItemDesc = m_ItemDesc & vbCrLf & vbCrLf & Replace(PUB_GetUniText(Me.Name, "·s®×Â½Ä¶¦³¬Û¦ü®×»¡©ú"), "XXXXX", RsTemp("pa77")) & vbCrLf
                  
                  'Added by Morgan 2024/6/24 µL¬Û¦ü®×»¡©ú
                  ElseIf InStr(m_CustNoList, "X4612500") > 0 Then
                     m_ItemDesc = m_ItemDesc & vbCrLf & vbCrLf & PUB_GetUniText(Me.Name, "·s®×Â½Ä¶µL¬Û¦ü®×»¡©ú")
                  'end 2024/6/24
                  End If
            End Select
         End If
      End If
      'end 2024/2/17
               
      m_iItem = m_iItem + 1
     'Modified by Lydia 2016/03/03 +m_bSpecial5
      If m_bSpecial1 Or m_bSpecial5 Then
         'Added by Morgan 2020/8/6
         If m_bDowN = True Then
            SetItemWordArray m_Item, m_iItem, 1, m_ItemDesc
         Else
         'end 2020/8/6
            SetItemWordArray m_Item, m_iItem, 1, m_ClearItemDesc
         End If 'Added by Morgan 2020/8/6
         SetItemWordArray m_Item, m_iItem, 5, Format(adoacc1k0.Fields("IDate"), "####/##/##")
      'Added by Morgan 2014/2/25
      'Modified by Morgan 2018/3/22 +m_bSpecialNew3
      ElseIf m_bSpecial4 Or m_bSpecialNew3 Then
         SetItemWordArray m_Item, m_iItem, 1, m_ItemDesc
         SetItemWordArray m_Item, m_iItem, 5, Format(adoacc1k0.Fields("IDate"), "####/##/##")
      'end 2014/2/25
      Else
         SetItemWordArray m_Item, m_iItem, 1, m_ItemDesc
      End If
      
      'Added by Morgan 2013/10/29
      SetItemWordArray m_Item, m_iItem, 6, m_ItemHDesc
      SetItemWordArray m_Item, m_iItem, 7, m_ItemTDesc
      SetItemWordArray m_Item, m_iItem, 8, "" & adoacc1k0.Fields("a1j02").Value
      SetItemWordArray m_Item, m_iItem, 9, m_ItemXDesc
      'end 2013/10/29
      
      'Add by Morgan 2004/10/4 Y27766000, Offical fees ¥[¦L¬üª÷
      If "" & adoacc1k0.Fields("a1k03").Value = "Y27766000" And Len("" & adoacc1k0.Fields("a1j02").Value) = 5 And Right("" & adoacc1k0.Fields("a1j02").Value, 2) = "99" Then
         'Modify By Sindy 2013/1/28
         'strAmount = Int((Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / Val(adoacc1k0.Fields("a1k10").Value))
         
         'Modify By Sindy 2013/3/29
         'strAmount = Int((Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / m_DNRate)
         strAmount = Trunc(dbl_A1L0507FAmt)
         '2013/3/29 End
         
         '2013/1/28 End
         m_TotOffFees = m_TotOffFees + Val(strAmount)
         'Modified by Morgan 2013/1/3 ª÷ÃB³£¥[¦L.00
         'strAmount = "( USD  " & Format(strAmount, DDollar) & " )"
         strAmount = "( " & m_DNCurr & "  " & Format(strAmount, FDollar) & " )"
         If m_b2Printer Then
            Printer.CurrentX = 7300 + intDefault - Printer.TextWidth(strAmount)
            Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
            'Modified by Morgan 2022/8/4
            'Printer.Print strAmount
            PUB_PrintUnicodeText strAmount, Printer.CurrentX, Printer.CurrentY, 0
            'end 2022/8/4
         End If
         'Removed by Morgan 2022/8/4
         'If m_b2Picture Then
         '   Picture1.CurrentX = (7300 + intDefault) * douExtRate - Picture1.TextWidth(strAmount)
         '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
         '   Picture1.Print strAmount
         'End If
         'end 2022/8/4
         SetItemWordArray m_Item, m_iItem, 2, strAmount  'Add by Morgan 2010/11/29
      
      'Add by Morgan 2010/12/3
     'Modified by Lydia 2016/03/03 +m_bSpecial5
     ' ElseIf m_bSpecial1 Then
      ElseIf m_bSpecial1 Or m_bSpecial5 Then
         If m_bSpecial5 Then
            m_tmp = "" & adoacc1k0.Fields("a1j18")
         Else
            m_tmp = "" & adoacc1k0.Fields("a1j18") & adoacc1k0.Fields("a1j19")
         End If
         SetItemWordArray m_Item, m_iItem, 2, m_tmp
         
      'Modified by Morgan 2012/11/2
      'ElseIf m_bSpecial2 Then
         'm_tmp = "" & adoacc1k0.Fields("a1j18") & adoacc1k0.Fields("a1j19")
      ElseIf m_bSpecial2 Or m_bSpecial3 Then
         'Modified by Morgan 2014/2/18 §ï§ìa2616
         'If IsNull(adoacc1k0.Fields("a2604")) Then
         If IsNull(adoacc1k0.Fields("a2616")) Then
            m_tmp = " "
            'm_sEBillingMsg = m_sEBillingMsg & vbCrLf & adoacc1k0("a1k13") & "¤§½Ð´Ú¶µ¥Ø" & adoacc1k0("a1j02") & "µL¹ïÀ³¤§¥N½X¡I"
         Else
            'Modified by Morgan 2014/2/18 §ï§ìa2616
            'm_tmp = adoacc1k0.Fields("a2604")
            m_tmp = adoacc1k0.Fields("a2616")
         End If
      'end 2012/11/2
         SetItemWordArray m_Item, m_iItem, 2, m_tmp
      'end 2010/12/3
      
      'Added by Morgan 2014/2/18
      ElseIf m_bSpecialNew1 Then
         m_tmp = "" & adoacc1k0.Fields("a2616")
         
         'Added by Morgan 2019/2/27
         '¨S³]©wªº§ì¥D¶µ(­ì«h¤W¤@±i½Ð´Ú³æ¥u·|¦³¤@­Ó¥D¶µ)
         If m_A1k28 = "Y48904000" Then
            If m_Activity = "" Then m_Activity = m_tmp
            If m_tmp = "" Then m_tmp = m_Activity
         End If
         'end 2019/2/27
         
         If m_tmp = "" Then m_tmp = " "
         SetItemWordArray m_Item, m_iItem, 2, m_tmp
      'end 2014/2/18
      End If
      
      '¦C¦Lª÷ÃB
      'Modified by Morgan 2012/12/7
      'Select Case strCurr
      '   Case "U"
      '      m_tmp = "USD"
      '   '2009/5/18 modify by sonia
      '   'Case Else
      '   '   m_tmp = "NTD"
      '   Case "N"
      '      m_tmp = "NTD"
      '   Case Else
      '      If m_DNCurr = "USD" Then
      '         m_tmp = "NTD"
      '      Else
      '         m_tmp = m_DNCurr
      '      End If
      '   '2009/5/18 end
      'End Select
      
      Select Case m_iPrintCurrType
      Case 3, 4 '¯Â¥~¹ô,¥~¹ô+¬üª÷¦X­p
         m_tmp = m_DNCurr
      Case Else
         m_tmp = "NTD"
      End Select
      'end 2012/12/7
      
      If m_b2Printer Then
         Printer.CurrentX = 8000 + intDefault
         Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
         'Modified by Morgan 2022/8/4
         'Printer.Print m_tmp
         PUB_PrintUnicodeText m_tmp, Printer.CurrentX, Printer.CurrentY, 0
         'end 2022/8/4
      End If
      'Removed by Morgan 2022/8/4
      'If m_b2Picture Then
      '   Picture1.CurrentX = (8000 + intDefault) * douExtRate
      '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
      '   Picture1.Print m_tmp
      'End If
      'end 2022/8/4
      SetItemWordArray m_Item, m_iItem, 3, m_tmp 'Add by Morgan 2010/11/29
      
      If IsNull(adoacc1k0.Fields("a1l05").Value) = False Or dblGrpAmt <> 0 Then
         If dblGrpAmt <> 0 Then
            strAmount = PUB_ChgFormat("" & dblGrpAmt, True)
            strAmountNoDisc = PUB_ChgFormat("" & dblGrpAmtNoDisc, False) 'Add by Morgan 2010/11/9
            douAmount = douAmount + dblGrpAmt
         Else
            'Modified by Morgan 2012/12/7
'            Select Case strCurr
'               Case "U"
'                  strAmount = PUB_ChgFormat((Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / Val(adoacc1k0.Fields("a1k10").Value), True)
'                  strAmountNoDisc = PUB_ChgFormat(Val(adoacc1k0.Fields("a1l05").Value) / Val(adoacc1k0.Fields("a1k10").Value), False)  'Add by Morgan 2010/11/9
'                  'Modify by Morgan 2006/9/29
'                  'douAmount = douAmount + (Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / Val(adoacc1k0.Fields("a1k10").Value)
'                  'P,T¥xÆW®×®É¥u¦L¾ã¼Æ
'                  If bolPTUSDCase Then
'                     douNTDollar = douNTDollar + Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)
'                     If douNTDollar >= Val("" & adoacc1k0.Fields("a1k11")) Then
'                        strAmount = Val("" & adoacc1k0.Fields("a1k08")) - douAmount
'                        douAmount = Val("" & adoacc1k0.Fields("a1k08"))
'                     Else
'                        strAmount = Int(strAmount)
'                        douAmount = douAmount + Val(strAmount)
'                     End If
'                  Else
'                     douAmount = douAmount + (Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / Val(adoacc1k0.Fields("a1k10").Value)
'                  End If
'                  'end 2006/9/29
'               '2009/5/18 modify by sonia
'               'Case Else
'               '   strAmount = PUB_ChgFormat(Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value), True)
'               '   douAmount = douAmount + Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)
'               Case "N"
'                  strAmount = PUB_ChgFormat(Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value), True)
'                  strAmountNoDisc = PUB_ChgFormat(Val(adoacc1k0.Fields("a1l05").Value), False)  'Add by Morgan 2010/11/9
'                  douAmount = douAmount + Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)
'               Case Else
'                  If m_DNCurr = "USD" Then
'                     strAmount = PUB_ChgFormat(Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value), True)
'                     strAmountNoDisc = PUB_ChgFormat(Val(adoacc1k0.Fields("a1l05").Value), False)  'Add by Morgan 2010/11/9
'                     douAmount = douAmount + Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)
'                  Else
'                     strAmount = PUB_ChgFormat((Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / m_DNRate, True)
'                     strAmountNoDisc = PUB_ChgFormat(Val(adoacc1k0.Fields("a1l05").Value) / m_DNRate, False) 'Add by Morgan 2010/11/9
'                     douAmount = douAmount + (Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / m_DNRate
'                  End If
'               '2009/5/18 end
'            End Select
            '
            Select Case m_iPrintCurrType
            Case 3, 4 '3.¯Â¥~¹ô, 4.¥~¹ô+¬üª÷¦X­p
               '·s½Ð´Ú³æ¥~¹ô¤p¼Æ³£±Ë¥h
               If m_bolNewBill Then
                  'Modify By Sindy 2013/1/28
'                  If m_DNCurr = "USD" Then
'                     strAmount = PUB_ChgFormat(Fix(Format((Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / Val(adoacc1k0.Fields("a1k10").Value))), True)
'                     strAmountNoDisc = PUB_ChgFormat(Fix(Format(Val(adoacc1k0.Fields("a1l05").Value) / Val(adoacc1k0.Fields("a1k10").Value))), False)
'                     douAmount = douAmount + Val(Format(strAmount))
'                  Else
                     
                     'Modify By Sindy 2013/3/29
                     'strAmount = PUB_ChgFormat(Fix(Format((Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / m_DNRate)), True)
                     'strAmountNoDisc = PUB_ChgFormat(Fix(Format(Val(adoacc1k0.Fields("a1l05").Value) / m_DNRate)), False)
                     
                     'Added by Morgan 2019/8/30 BASF Â½Ä¶¶O ¬üª÷­nÅã¥Ü¦Ü¤p¼Æ²Ä2¦ì
                     'Modified by Morgan 2022/2/18 +927¨ä¥LÂ½Ä¶ Ex:X11102382-- Ryan
                     'Modified by Morgan 2022/9/2 +209ÀËµø¤¤»¡-- Tim
                     'Modified by Morgan 2025/10/31 +FG Ex:X11404796
                     If (adoacc1k0.Fields("a1k28") = "Y45814010" Or adoacc1k0.Fields("a1k28") = "Y33268010") And (adoacc1k0.Fields("a1k13") = "FCP" Or adoacc1k0.Fields("a1k13") = "FG" Or adoacc1k0.Fields("a1k13") = "P" Or adoacc1k0.Fields("a1k13") = "CFP") And (adoacc1k0.Fields("a1j02") = "201" Or adoacc1k0.Fields("a1j02") = "927" Or adoacc1k0.Fields("a1j02") = "209") Then
                        strAmount = PUB_ChgFormat("" & dbl_A1L0507FAmt, True)
                        strAmountNoDisc = PUB_ChgFormat("" & dbl_A1L05FAmt, False)
                        
                     Else
                     'end 2019/8/30
                     
                        strAmount = PUB_ChgFormat(Trunc(dbl_A1L0507FAmt), True)
                        strAmountNoDisc = PUB_ChgFormat(Trunc(dbl_A1L05FAmt), False)
                        
                     End If 'Added by Morgan 2019/8/30
                     
                     
                     '2013/3/29 End
                     douAmount = douAmount + Val(Format(strAmount))
'                  End If
                  '2013/1/28 End
               Else
                  If m_DNCurr = "USD" Then
                     'Modify By Sindy 2013/1/28
'                     strAmount = PUB_ChgFormat((Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / Val(adoacc1k0.Fields("a1k10").Value), True)
'                     strAmountNoDisc = PUB_ChgFormat(Val(adoacc1k0.Fields("a1l05").Value) / Val(adoacc1k0.Fields("a1k10").Value), False)
                     
                     'Modify By Sindy 2013/3/29
                     'strAmount = PUB_ChgFormat((Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / m_DNRate, True)
                     'strAmountNoDisc = PUB_ChgFormat(Val(adoacc1k0.Fields("a1l05").Value) / m_DNRate, False)
                     strAmount = PUB_ChgFormat(CStr(dbl_A1L0507FAmt), True)
                     strAmountNoDisc = PUB_ChgFormat(CStr(dbl_A1L05FAmt), False)
                     '2013/3/29 End
                     
                     '2013/1/28 End
                     'P,T¥xÆW®×®É¥u¦L¾ã¼Æ
                     If bolPTUSDCase Then
                        douNTDollar = douNTDollar + Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)
                        If douNTDollar >= Val("" & adoacc1k0.Fields("a1k11")) Then
                           strAmount = Val("" & adoacc1k0.Fields("a1k08")) - douAmount
                           douAmount = Val("" & adoacc1k0.Fields("a1k08"))
                        Else
                           strAmount = Int(strAmount)
                           douAmount = douAmount + Val(strAmount)
                        End If
                     Else
                        douAmount = douAmount + Val(Format(strAmount))
                     End If
                  Else
                     'Modify By Sindy 2013/3/29
                     'Modify By Sindy 2013/1/28
                     'strAmount = PUB_ChgFormat((Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / m_DNRate, True)
                     'strAmountNoDisc = PUB_ChgFormat(Val(adoacc1k0.Fields("a1l05").Value) / m_DNRate, False)
                     'douAmount = douAmount + (Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / m_DNRate
                     strAmount = PUB_ChgFormat(CStr(dbl_A1L0507FAmt), True)
                     strAmountNoDisc = PUB_ChgFormat(CStr(dbl_A1L05FAmt), False)
                     douAmount = douAmount + dbl_A1L0507FAmt
                     '2013/3/29 End
                     
'                     strAmount = PUB_ChgFormat((Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / Val(adoacc1k0.Fields("a1k10").Value), True)
'                     strAmountNoDisc = PUB_ChgFormat(Val(adoacc1k0.Fields("a1l05").Value) / Val(adoacc1k0.Fields("a1k10").Value), False)
'                     douAmount = douAmount + (Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)) / Val(adoacc1k0.Fields("a1k10").Value)
                     '2013/1/28 End
                  End If
               End If
               
            Case Else '1.¯Â¥x¹ô 2.¥x¹ô+¥~¹ô¦X­p
               strAmount = PUB_ChgFormat(Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value), True)
               strAmountNoDisc = PUB_ChgFormat(Val(adoacc1k0.Fields("a1l05").Value), False)
               douAmount = douAmount + Val(adoacc1k0.Fields("a1l05").Value) - Val(adoacc1k0.Fields("a1l07").Value)
            End Select
            'end 2012/12/7
            
         End If
         
         'Modified by Morgan 2013/1/3 ª÷ÃB¤@²v³£­n¦L.00
         
'         'Added by Morgan 2012/8/3 ¬üª÷¾ã¼Æ¤]­n¦L .00
'         'If m_tmp = "USD" Then
'            strAmount = Format(strAmount, FDollar)
'         End If
'         'end 2012/8/3

         strAmount = Format(strAmount, FDollar)
         'end 2013/1/3
         
         'Added by Morgan 2014/2/24
         If m_bSpecial4 Then
            m_tmp = " "
            If Val("" & adoacc1k0.Fields("IDate")) > 0 Then
               'µL±ø¥ó¶i¦ì¨ì¤p¼Æ¨â¦ì
               strExc(1) = -1 * Int(-100 * Val(strAmountNoDisc) / 6000) / 100
               If Val(strExc(1)) > 0 Then
                  m_tmp = strExc(1)
               End If
            End If
            
            SetItemWordArray m_Item, m_iItem, 2, m_tmp
         End If
         'end 2014/2/24
      
         If m_b2Printer Then
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 10000 + intDefault - intLength
            Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
            'Modified by Morgan 2022/8/4
            'Printer.Print strAmount
            PUB_PrintUnicodeText strAmount, Printer.CurrentX, Printer.CurrentY, 0
            'end 2022/8/4
         End If
         'Removed by Morgan 2022/8/4
         'If m_b2Picture Then
         '   intLength = Picture1.TextWidth(strAmount)
         '   Picture1.CurrentX = (10000 + intDefault) * douExtRate - intLength
         '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
         '   Picture1.Print strAmount
         'End If
         'end 2022/8/4
         
         '¦C¦L©ú²Óª÷ÃB
         
         'Modified by Lydia 2016/03/03 + m_bSpecial5
         'Modified by Morgan 2020/8/6 Y22457000,Y48048000,Y52322000 §é¦©¤£¥t¦C--Tim
         If m_bSpecial1 Or m_bSpecial5 Then
            'Added by Morgan 2020/8/6
            If m_bDowN = True Then
               SetItemWordArray m_Item, m_iItem, 4, strAmount
               SetItemWordArray m_Item, m_iItem, 19, Format(strAmountNoDisc, FDollar)
            Else
            'end 2020/8/6
               SetItemWordArray m_Item, m_iItem, 4, Format(strAmountNoDisc, FDollar) 'Add by Morgan 2010/12/3
            End If 'Added by Morgan 2020/8/6
            m_dblNoDiscAmtTot = m_dblNoDiscAmtTot + Val(Format(strAmountNoDisc))
            'm_dblDiscTot = m_dblDiscTot + Val(strAmountNoDisc) - Val(Format(strAmount))
         Else
            SetItemWordArray m_Item, m_iItem, 4, strAmount 'Add by Morgan 2010/11/29
            SetItemWordArray m_Item, m_iItem, 18, Format(Val("" & adoacc1k0.Fields("a1l05")) - Val("" & adoacc1k0.Fields("a1l07")), FDollar) 'Add by Morgan 2016/8/5
            'Added by Morgan 2014/8/28
            If m_PlusFormNo <> "" Then
               'Modified by Morgan 2015/6/11 ­×¥¿¥H§é¦©¼Æ±Àºâ¥¼§é¦©ª÷ÃB·|¦³¤p¼Æ°ÝÃD Ex.X10408788
               If Right(m_Item(m_iItem).iNo, 2) <> "99" And Val(strAmountNoDisc) > 0 Then
                  'Modified by Morgan 2015/6/18 ¶Ç¯u¶O¡BÂø¶O°£¥~--Kimi
                  If m_Item(m_iItem).iNo = "01" Or m_Item(m_iItem).iNo = "02" Then
                     strExc(0) = strAmount
                  'Modified by Morgan 2020/6/23 §ï¦³¥´§é®É¥Î­ì»ù¡A¨S¥´§é®É´£°ª2¦¨--Kimi
                  'ElseIf 100 * Round(Val(Format(strAmount)) / Val(strAmountNoDisc), 2) = 72 Then
                  '   strExc(0) = Format(Round(Val(strAmountNoDisc) * 0.9), FDollar)
                  'Else
                  '   strExc(0) = Format(Val(strAmountNoDisc) * (Round(Val(Format(strAmount)) / Val(strAmountNoDisc), 2) + 0.2), FDollar)
                  ElseIf Val(Format(strAmount)) = Val(strAmountNoDisc) Then
                     strExc(0) = Format(Trunc(Val(strAmountNoDisc) * 1.2), FDollar)
                  Else
                     strExc(0) = Format(strAmountNoDisc, FDollar)
                  'end 2020/6/20
                  End If
                  'end 2015/6/18
                  SetItemWordArray m_Item, m_iItem, 10, strExc(0)
               End If
               'end 2015/6/11
            End If
            'end 2014/8/28
            
         End If
         
         'Added by Lydia 2015/04/09 +¾ã§å½Ð´Ú³æ-©ú²Ó¸ê®Æ
         If m_bolChiDB And bol_ChiDB Then
             SetItemWordArray m_Item, m_iItem, 11, midStr(0) '®×¸¹
             SetItemWordArray m_Item, m_iItem, 12, midStr(1) '¦WºÙ
             SetItemWordArray m_Item, m_iItem, 13, midStr(2) 'Ãþ§O
             SetItemWordArray m_Item, m_iItem, 14, midStr(3) 'µù¥U¸¹/¥Ó½Ð¸¹
             strExc(8) = Format(strAmount, "###0.00")
             If m_iPrintCurrType = 2 Then
                strExc(7) = Format(Trunc(Val(strExc(8)) / Val(m_DNRate)), "###0")
             Else
                strExc(7) = Format(Trunc(Val(strExc(8)) * Val(m_DUsdRate)), "###0")
             End If
             If mChiCNo = adoacc1k0.Fields("a1k01") & adoacc1k0.Fields("a1j02") Then
                mChiOAmt = Format(Val(mChiOAmt) + Val(strExc(8)), "###0")
                mChiUAmt = Format(Val(mChiUAmt) + Val(strExc(7)), "###0")
             Else
                'Modified by Lydia 2020/09/08 «D³W¶O¤~+99
                'mChiCNo = adoacc1k0.Fields("a1k01") & adoacc1k0.Fields("a1j02") & "99"
                mChiCNo = adoacc1k0.Fields("a1k01") & adoacc1k0.Fields("a1j02") & IIf(Right("" & adoacc1k0.Fields("a1j02"), 2) = "99", "", "99")
                mChiOAmt = strExc(8)
                mChiUAmt = strExc(7)
             End If
             SetItemWordArray m_Item, m_iItem, 15, adoacc1k0.Fields("a1k01")
             SetItemWordArray m_Item, m_iItem, 16, mChiOAmt '(³W¶O+ªA°È¶O)=½Ð´Úª÷ÃB
             SetItemWordArray m_Item, m_iItem, 17, mChiUAmt
         End If
         
         'Add by Morgan 2010/11/5
         '¬üª÷½Ð´Ú
         'Modified by Morgan 2012/12/7
         'If strCurr = "U" Then
         If m_iPrintCurrType = 3 And m_DNCurr = "USD" Then
         'end 2012/12/7
            douUSAmount = Format(strAmount)
            douUSAmountNoDisc = Format(strAmountNoDisc)
         'Modify by Sindy 2013/1/28
'         ElseIf adoacc1k0("a1k10") > 0 Then
'            'Modified by Morgan 2012/8/31 ©ú²Ó¨ú¤p¼Æ¨â¦ì«á±Ë¥h(
'            'douUSAmount = Val(Int(Format(strAmount) / adoacc1k0("a1k10")))
'            'douUSAmountNoDisc = Val(Int(Format(strAmountNoDisc) / adoacc1k0("a1k10")))
'
'            'Modified by Morgan 2012/12/7
'            '·s½Ð´Ú³æ¥~¹ô¤p¼Æ³£±Ë¥h
'            If m_bolNewBill  Then
'               douUSAmount = Fix(Format(Format(strAmount) / adoacc1k0("a1k10")))
'               douUSAmountNoDisc = Fix(Format(Format(strAmountNoDisc) / adoacc1k0("a1k10")))
'            Else
'            'end 2012/12/7
'               douUSAmount = Fix((100 * Format(strAmount)) / adoacc1k0("a1k10")) / 100
'               douUSAmountNoDisc = Fix((100 * Format(strAmountNoDisc)) / adoacc1k0("a1k10")) / 100
'            End If 'Added by Morgan 2012/12/7
         ElseIf m_DNRate > 0 Then
            'Modified by Morgan 2012/8/31 ©ú²Ó¨ú¤p¼Æ¨â¦ì«á±Ë¥h(
            'douUSAmount = Val(Int(Format(strAmount) / adoacc1k0("a1k10")))
            'douUSAmountNoDisc = Val(Int(Format(strAmountNoDisc) / adoacc1k0("a1k10")))
            
            'Modified by Morgan 2012/12/7
            '·s½Ð´Ú³æ¥~¹ô¤p¼Æ³£±Ë¥h
            If m_bolNewBill Then
               'Add By Sindy 2013/3/29
               'douUSAmount = Trunc(Format(strAmount) / m_DNRate)
               'douUSAmountNoDisc = Trunc(Format(strAmountNoDisc) / m_DNRate)
               'Modifid by Morgan 2017/10/18 HP§é¦©ª÷ÃB­n­pºâ¨ì¤p¼Æ1¦ì
               'douUSAmountNoDisc = Trunc(dblGrpAmtNoDiscFAmt)
               If m_A1k28 = "Y48292000" Then
                  douUSAmount = Trunc(dblGrpAmtFAmt, 1)
                  douUSAmountNoDisc = Trunc(dblGrpAmtNoDiscFAmt, 1)
               Else
                  douUSAmount = Trunc(dblGrpAmtFAmt)
                  douUSAmountNoDisc = Trunc(dblGrpAmtNoDiscFAmt)
               End If
               'end 2017/10/18
               '2013/3/29 End
            Else
            'end 2012/12/7
               'Add By Sindy 2013/3/29
               'douUSAmount = Trunc(Format(strAmount) / m_DNRate, 2)
               'douUSAmountNoDisc = Trunc(Format(strAmountNoDisc) / m_DNRate, 2)
               douUSAmount = Trunc(dblGrpAmtFAmt, 2)
               douUSAmountNoDisc = Trunc(dblGrpAmtNoDiscFAmt, 2)
               '2013/3/29 End
            End If 'Added by Morgan 2012/12/7
         '2013/1/28 End
         ElseIf m_bEBilling Then
            m_bEBilling = False
            m_sEBillingMsg = m_sEBillingMsg & vbCrLf & adoacc1k0("a1k01") & "¶×²v¿ù»~¡AµLªk²£¥Í¹q¤l±b³æ¡I"
         End If
            
         'Modify by Morgan 2010/12/29 ­×§ï LINE_ITEM_UNIT_COST,LINE_ITEM_UNIT_COST
         If m_bEBilling Then
            iUpper = iUpper + 1
            ReDim Preserve strLedes(m_iCols, iUpper)
         
            With adoacc1k0
            '¥u­n³]©w²Ä¤@µ§,¨ä¥L¬Û¦P
            If iUpper = 1 Then
               '1 INVOICE_DATE
               strLedes(1, iUpper) = DBDATE(.Fields("a1k02"))
               '2 'INVOICE_NUMBER
               strLedes(2, iUpper) = .Fields("a1k01")
            
               '3 CLIENT_ID(­Y¨S¦³³]©w®É¥Î¥N²z¤H½s¸¹)
               'Modified by Morgan 2012/4/24 LEDED ³]©w§ï§ì LEDED ¸ê®Æªí
               'strLedes(3, iUpper) = "" & .Fields("fa103")
               strLedes(3, iUpper) = "" & adoLEDES.Fields("ld02")
               If strLedes(3, iUpper) = "" Then
                  strLedes(3, iUpper) = m_A1k28
               End If
               
               '4 LAW_FIRM_MATTER_ID(¥»©Ò®×¸¹)
               strLedes(4, iUpper) = .Fields("a1k13") & "-" & .Fields("a1k14") & IIf(.Fields("a1k15") & .Fields("a1k16") = "000", "", "-" & .Fields("a1k15") & "-" & .Fields("a1k16"))
               'Added by Morgan 2017/6/20 Y34126 LAW_FIRM_MATTER_ID ³Ì«á­±­n©ñTW(the 2 character WIPO Country Cod)
               If m_A1k28 = "Y34126000" Then
                  strLedes(4, iUpper) = strLedes(4, iUpper) & "TW"
               End If
               'end 2017/6/20
               
               '5 INVOICE_TOTAL
               'Added by Morgan 2021/3/15 X82720000 General Mills, Inc. ¥Î¥x¹ô½Ð´Ú--Ali
               If m_A1k28 = "X82720000" Then
                  strLedes(5, iUpper) = .Fields("a1k11")
               Else
               'end 2021/3/15
                  strLedes(5, iUpper) = .Fields("a1k08")
               End If
               
               '6 BILLING_START_DATE
               'Modify by Morgan 2011/9/26 °t¦XHP²Î¤@§ï§ì³Ì¦­¦¬¤å¤é
               'strLedes(6, iUpper) = strLedes(1, iUpper)
               'Added by Morgan 2016/8/5
               'Y54237 BRISTOL-MYERS SQUIBB CO. Patent Department ±N±b³æ¤é´Áªº°_¨´¤é/²×¤î¤é§ï¬°±b³æ¤é´Á-- ³¯©É»T
               'Modified by Morgan 2017/10/23 +Y54869000 Albemarle Corp --¼ï¤l·L
               'Modified by Morgan 2019/6/14 +Y55134000 Albemarle Germany GmbH--Ryan
               If m_A1k28 = "Y54237000" Or m_A1k28 = "Y54869000" Or m_A1k28 = "Y55134000" Then
                  strLedes(6, iUpper) = strLedes(1, iUpper)
               'end 2016/8/5
               'Added by Morgan 2018/6/25 +Dow¯S®í³]©w
               '±b³æ°_¤é§ì³Ì¦­µo¤å¤é
               ElseIf m_bSpecial1 Then
                  strLedes(6, iUpper) = GetCp27(.Fields("a1k01"))
               'end 2018/6/25
               'Added by Morgan 2023/7/10
               '½Ð´Ú¤é·í¤ë1¸¹
               ElseIf m_A1k28 = "Y48279000" Then
                  strLedes(6, iUpper) = Left(strLedes(1, iUpper), 6) & "01"
               'Added by Morgan 2025/6/19 Y2232700 MKS Inc.
               '³Ì±ßµo¤å¤é·í¤ë1¸¹¨ì·í¤ë©³
               'Removed by Morgan 2025/9/8 §ï¬°±b³æ¤é·í¤ë²Ä¤@¤Ñ¨ì·í¤ë³Ì«á¤@¤Ñ(¤U­±)
               'ElseIf m_A1k28 = "Y22327000" Then
               '   strLedes(6, iUpper) = Left(GetCp27(m_strDN, , True), 6) & "01"
               '   strLedes(7, iUpper) = CompDate(2, -1, CompDate(1, 1, strLedes(6, iUpper)))
               End If
            
               If strLedes(6, iUpper) = "" Then
                  strLedes(6, iUpper) = GetMinCp05(.Fields("a1k01"))
               End If
            
               '7 BILLING_END_DATE
               If strLedes(7, iUpper) = "" Then strLedes(7, iUpper) = strLedes(1, iUpper)
            
               'Added by Morgan 2016/8/23
               'Y52234010 Abercrombie & Fitch Europe Sagl, BILLING_START_DATE¤ÎBILLING_END_DATE¶·Åã¥Ü¬°·í¤ë¥÷¤§²Ä¤@¤Ñ¤Î³Ì«á¤@¤Ñ--¤Bà±»Z
               'Modified by Morgan 2019/1/28 +X69455 Columbia Sportswear North America, Inc. --Ã¹ÝÂ¾ç
               'Modified by Morgan 2020/10/28 +Y54869000 --³¯«W¥c
               'Modified by Morgan 2023/8/29 +Y54096000 --Tim
               'Modified by Morgan 2025/4/14 +X56347000 --Tim
               'Modified by Morgan 2025/4/21 +Y56065000 --Tim
               'Modified by Morgan 2025/8/5 +X28716030 --Tim
               'Modified by Morgan 2025/9/8 +Y22327000 --Lisa
               If m_A1k28 = "Y52234010" Or m_A1k28 = "X69455000" Or m_A1k28 = "Y54869000" Or m_A1k28 = "Y54096000" Or m_A1k28 = "X56347000" Or m_A1k28 = "Y56065000" Or m_A1k28 = "X28716030" Or m_A1k28 = "Y22327000" Then
                  strLedes(6, iUpper) = Left(strLedes(1, iUpper), 6) & "01"
                  strLedes(7, iUpper) = CompDate(2, -1, CompDate(1, 1, strLedes(6, iUpper)))
               End If
               
               'Added by Morgan 2024/6/24 X76126000 --Kahn
               If m_A1k28 = "X76126000" Then
                  'Billing start date:AÃþ¦¬¤å³Ì¤jµo¤å¤é·í¤ë1¸¹
                  strLedes(6, iUpper) = GetCp27(.Fields("a1k01"), , True, True)
                  strLedes(6, iUpper) = Left(strLedes(6, iUpper), 6) & "01"
                  'Billing end date:AÃþ¦¬¤å³Ì¤jµo¤å¤é·í¤ë³Ì«á¤@¤Ñ
                  strLedes(7, iUpper) = CompDate(2, -1, CompDate(1, 1, strLedes(6, iUpper)))
               End If
               'end 2024/6/24

               '8 INVOICE_DESCRIPTION
               'Modified by Morgan 2013/5/9 Teva ªº INVOICE_DESCRIPTION «e­±­n¥[ "PO number -",¥Ø«e©T©w¥Î "1206041-"
               If m_A1k28 = "Y53713000" Then
                  'Added by Morgan 2014/7/25
                  If strLedes(4, iUpper) = "FCP-050364" Then
                     strLedes(8, iUpper) = "PO-129164-" & .Fields("a1j04")
                  Else
                  'end 2014/7/25
                     strLedes(8, iUpper) = "1206041-" & .Fields("a1j04")
                  End If
               'Added by  Morgan 2019/5/29 Nordson ªº DESCRIPTION ­n©ñ®×¥ó¦WºÙ
               ElseIf m_A1k28 = "X70124000" Then
                  strLedes(8, iUpper) = m_CaseName
               'end 2019/5/29
               'Added by Morgan 2020/8/6-- ³¯«W¥c
               'X82995¡]Onto Innovation, Inc.¡^©ñ«È¤á®×¥ó®×¸¹
               'Removed by Morgan 2020/10/23 ¨ú®ø-- ³¯«W¥c
               'ElseIf m_A1k28 = "X82995000" Then
               '   strLedes(8, iUpper) = GetCustCaseNo(.Fields("a1k13"), .Fields("a1k14"), .Fields("a1k15"), .Fields("a1k16"))
               'end 2020/10/23
               'end 2020/8/6
               'Added by Morgan 2021/4/9
               ElseIf m_A1k28 = "Y52216000" Then
                  strLedes(8, iUpper) = GetYourRefNo1(.Fields("a1k13"), .Fields("a1k14"), .Fields("a1k15"), .Fields("a1k16"))
               'end 2021/4/9
               
               'Added by Morgan 2023/10/13 X7692800 (BeiGene, LTD)--Tim
               'INVOICE_DESCRIPTION Åã¥Ü­Ó®×ªº¥Ó½Ð®×¸¹¡B±M§Q¦WºÙtitle¡B«È¤á®×¥ó®×¸¹
               ElseIf m_A1k28 = "X76928000" Then
                  strLedes(8, iUpper) = "[Appl. No.]" & strAppNo & "[Title]" & m_CaseName & "[BeiGene Ref.]" & m_CaseNo
               'end 2023/10/13
               
               Else
                  'Added by Morgan 2023/1/3
                  If m_ItemDesc <> "" Then
                     strLedes(8, iUpper) = m_ItemDesc
                  Else
                  'end 2023/1/3
                     strLedes(8, iUpper) = "" & .Fields("a1j04")
                  End If
                  'Added by Morgan 2015/10/30
                  'Mondelez Global LLC­n¨D±NAnaqua number(«È¤á®×¥ó®×¸¹)©ñ¦bwork description¤¤
                  If m_A1k28 = "Y54037000" Then
                     strLedes(8, iUpper) = GetCustCaseNo(.Fields("a1k13"), .Fields("a1k14"), .Fields("a1k15"), .Fields("a1k16")) & " " & strLedes(8, iUpper)
                  End If
                  'end 2015/10/30
               End If
               
               'Added by Morgan 2017/3/31
               'Y54225B10 Syngenta ªº INVOICE_DESCRIPTION «e­±­n¥[[PO number],¥Ø«ePO number=CLIENT_ID"
               'Modified by Morgan 2017/5/2 +Y48309070--§d°ê¦w
               'Modified by Morgan 2019/2/1 PO_NUBMER§ï¿W¥ßÄæ¦ì(»PCLIENT_ID¤£¦P)
               If m_A1k28 = "Y54225B10" Or m_A1k28 = "Y48309070" Then
                  'strLedes(8, iUpper) = "[" & adoLEDES.Fields("ld02") & "]" & strLedes(8, iUpper)
                  strLedes(8, iUpper) = "[" & adoLEDES.Fields("ld17") & "]" & strLedes(8, iUpper)
               End If
               'end 2017/3/31
               
               'Added by Morgan 2025/4/14 --Tim
               If m_A1k28 = "X56347000" Then
                  strLedes(8, iUpper) = "(Matter Name: " & GetYourRefNo1(.Fields("a1k13"), .Fields("a1k14"), .Fields("a1k15"), .Fields("a1k16")) & ")" & strLedes(8, iUpper)
               End If
               'end 2025/4/14
               
               'Added by Morgan 2022/8/8--Kahn
               If m_A1k28 = "X82720000" Then
                  strLedes(8, iUpper) = "[All services covered by this invoice were provided outside the United States]" & strLedes(8, iUpper)
               End If
               'end 2022/8/8
               
               'Added by Morgan 2025/1/13 X82916010 Zuffa, LLC¯S®í³]©w --Kahn
               'Modified by Morgan 2025/1/16 ¦A¥[¥Ó½Ð¸¹ --Kahn
               If m_A1k28 = "X82916010" Then
                  strLedes(8, iUpper) = "(Application No. " & strAppNo & ")(" & strLedes(4, iUpper) & ")" & strLedes(8, iUpper)
                  strLedes(4, iUpper) = m_A1k28
               End If
               'end 2025/1/13
               
               'Added by Morgan 2023/7/12
               If m_A1k28 = "Y48279000" Then
                  strLedes(8, iUpper) = strLedes(8, iUpper) & " for Application No. " & strAppNo
               End If
               'end 2023/7/12
               
               '20 LAW_FIRM_ID(­Y¨S¦³³]©w®É¥Î±M§Qªk«ß¨Æ°È©Ò²Î½s)
               'Modified by Morgan 2012/4/24 LEDED ³]©w§ï§ì LEDED ¸ê®Æªí
               'strLedes(20, iUpper) = "" & .Fields("fa105")
               strLedes(20, iUpper) = "" & adoLEDES.Fields("ld14")
               If strLedes(20, iUpper) = "" Then
                  strLedes(20, iUpper) = "04146457"
               End If
               
               '24 CLIENT_MATTER_ID(­Y¨S¦³³]©w®É¥Î©¼©Ò®×¸¹)
               strLedes(24, iUpper) = GetClientMatterID(.Fields("a1k13"), .Fields("a1k14"), .Fields("a1k15"), .Fields("a1k16"), .Fields("a1k01"), m_A1k28) 'CLIENT_MATTER_ID
            Else
               strLedes(1, iUpper) = strLedes(1, 1)
               strLedes(2, iUpper) = strLedes(2, 1)
               strLedes(3, iUpper) = strLedes(3, 1)
               strLedes(4, iUpper) = strLedes(4, 1)
               strLedes(5, iUpper) = strLedes(5, 1)
               strLedes(6, iUpper) = strLedes(6, 1)
               strLedes(7, iUpper) = strLedes(7, 1)
               strLedes(8, iUpper) = strLedes(8, 1)
               strLedes(20, iUpper) = strLedes(20, 1)
               strLedes(24, iUpper) = strLedes(24, 1)
            End If
            
            '9 LINE_ITEM_NUMBER(¶µ¦¸)
            strLedes(9, iUpper) = iUpper
            
            '10 EXP/FEE/INV_ADJ_TYPE(¶µ¥ØÃþ§O)
            If .Fields("a1j19") <> "" Then
               strLedes(10, iUpper) = "E"
            Else
               strLedes(10, iUpper) = "F"
            End If
                     
            'Modify by Morgan 2011/5/12 ¥~°Ó§é¦©¥t¦C,³æ¦ì¼Æ©T©w¥Î1
            'Modified by Morgan 2014/7/17 §é¦©§ï¤£¥t¦C
            'Modified by Morgan 2021/5/20 +T
            'Modified by Morgan 2022/10/14 +TS
            If .Fields("a1k13") = "FCT" Or .Fields("a1k13") = "S" Or .Fields("a1k13") = "T" Or .Fields("a1k13") = "TS" Then
               '11 LINE_ITEM_NUMBER_OF_UNITS(³æ¦ì¼Æ/°Ó«~Ãþ§O¼Æ)
               strLedes(11, iUpper) = "1"
               '12 LINE_ITEM_ADJUSTMENT_AMOUNT(§é¦©)
               strLedes(12, iUpper) = "0" 'Modified by Morgan 2013/5/27 ¨S¦³§é¦©§ï©ñ0(­ì©ñNULL,¦]BASF¦³­n¨D°®¯Ü³£§ï¤Ï¥¿­ì³W©w¬O³£¥i¥H)
               
               'Added by Morgan 2021/3/24 +¯Â¥~¹ô
               'Modified by Morgan 2021/5/21 +¥~¹ô+¬üª÷
               If m_iPrintCurrType = 3 Or m_iPrintCurrType = 4 Then
                  '13 LINE_ITEM_TOTAL(§é¦©«á)
                  strLedes(13, iUpper) = Format(strAmount)
                  '12 LINE_ITEM_ADJUSTMENT_AMOUNT
                  strLedes(12, iUpper) = Format(strAmount) - Format(strAmountNoDisc)
              Else
              'end 2021/3/24
                  'Added by Morgan 2014/7/17 §é¦©§ï¤£¥t¦C
                  If douUSAmountNoDisc <> douUSAmount Then
                     strLedes(12, iUpper) = Round(douUSAmount - douUSAmountNoDisc, 4)
                  End If
                  'end 2014/7/17
                  '13 LINE_ITEM_TOTAL(½Ð´Úª÷ÃB)
                  'Modified by Morgan 2014/7/17 §é¦©§ï¤£¥t¦C
                  'strLedes(13, iUpper) = Round(douUSAmountNoDisc, 4)
                  strLedes(13, iUpper) = Round(douUSAmount, 4)
                  'end 2014/7/17
                  
               End If
               
               '21 LINE_ITEM_UNIT_COST(³æ»ù)
               'Modified by Morgan 2014/7/17 §é¦©§ï¤£¥t¦C
               'strLedes(21, iUpper) = strLedes(13, iUpper)
               strLedes(21, iUpper) = Val(strLedes(13, iUpper)) - Val(strLedes(12, iUpper))
               'end 2014/7/17
               
               'Added by Morgan 2019/5/14 Nordson ¤£¥i¶W¹L³]©wªº Rate (¥Ø«e¬°400),¶W¹L®É§ï¥Î 100 ¦A­«·s­pºâ®É¼Æ Ex:X10801665,X10802990
               If m_A1k28 = "X70124000" And Val("" & adoLEDES.Fields("ld16")) > 0 Then
                  If Val(strLedes(21, iUpper)) > Val(adoLEDES.Fields("ld16")) Then
                     strLedes(21, iUpper) = "100"
                     strLedes(11, iUpper) = Round((Val(strLedes(13, iUpper)) - Val(strLedes(12, iUpper))) / Val(strLedes(21, iUpper)), 4)
                  End If
               End If
               'end 2019/5/14
               
            Else
               'Modified by Morgan 2012/10/30 Rockwell ¹q¤l±b³æ¥u¤¹³\¤u®É¨ì¤p¼Æ¤@¦ì¬G¨C¶µ¦¸ªºªº½Ð´Úª÷ÃB¥²¶·¬°¾ã¼Æ
               If m_A1k28 <> "Y48292000" Then 'Added by Morgan 2017/10/18 HP§é¦©ª÷ÃB­n­pºâ¨ì¤p¼Æ1¦ì
                  douUSAmount = Trunc(douUSAmount)
                  douUSAmountNoDisc = Trunc(douUSAmountNoDisc)
               End If
               
               'Added by Morgan 2018/9/18
               'Dow§é¦©¥t¦C
               'Modified by Morgan 2020/2/17 Y22457000,Y48048000,Y52322000 ¨ú®ø--Tim
               If m_bSpecial1 And m_bDowN = False Then
                  '13 LINE_ITEM_TOTAL(§é¦©«á)
                  strLedes(13, iUpper) = Round(douUSAmountNoDisc, 4)
                  '12 LINE_ITEM_ADJUSTMENT_AMOUNT
                  strLedes(12, iUpper) = "0"
                  
               'Added by Morgan 2023/10/13--Tim
               'BeiGene §é¦©Á`©M¥t¦C³Ì«á
               ElseIf m_A1k28 = "X76928000" Then
                  douDiscount = douDiscount + Round(douUSAmount - douUSAmountNoDisc, 4)
                  strLedes(13, iUpper) = Round(douUSAmountNoDisc, 4)
                  strLedes(12, iUpper) = "0"
               'end 2023/10/13
               
               'Added by Morgan 2021/3/15 X82720000 General Mills, Inc. ¥Î¥x¹ô½Ð´Ú--Ali
               ElseIf m_A1k28 = "X82720000" Then
                  '13 LINE_ITEM_TOTAL(§é¦©«á)
                  strLedes(13, iUpper) = Format(strAmount)
                  '12 LINE_ITEM_ADJUSTMENT_AMOUNT
                  If strAmountNoDisc <> Format(strAmount) Then
                     strLedes(12, iUpper) = Format(strAmount) - Val(strAmountNoDisc)
                  Else
                     strLedes(12, iUpper) = "0"
                  End If
               
               'Added by Morgan 2021/3/24 +¯Â¥~¹ô
               'Modified by Morgan 2021/5/21 +¥~¹ô+¬üª÷
               'Modified by Morgan 2021/5/25 +¯Â¥x¹ô
               ElseIf m_iPrintCurrType = 1 Or m_iPrintCurrType = 3 Or m_iPrintCurrType = 4 Then
                  '13 LINE_ITEM_TOTAL(§é¦©«á)
                  strLedes(13, iUpper) = Format(strAmount)
                  '12 LINE_ITEM_ADJUSTMENT_AMOUNT
                  strLedes(12, iUpper) = Format(strAmount) - Format(strAmountNoDisc)
               'end 2021/3/24
               Else
               
                  '13 LINE_ITEM_TOTAL(§é¦©«á)
                  strLedes(13, iUpper) = Round(douUSAmount, 4)
                  '12 LINE_ITEM_ADJUSTMENT_AMOUNT
                  If douUSAmountNoDisc <> douUSAmount Then
                     strLedes(12, iUpper) = Round(douUSAmount - douUSAmountNoDisc, 4)
                  Else
                     strLedes(12, iUpper) = "0" 'Modified by Morgan 2013/5/27 ¨S¦³§é¦©§ï©ñ0(­ì©ñNULL,¦]BASF¦³­n¨D°®¯Ü³£§ï¤Ï¥¿­ì³W©w¬O³£¥i¥H)
                  End If
                  
               End If
               
               '21 LINE_ITEM_UNIT_COST(³æ»ù)
               '11 LINE_ITEM_NUMBER_OF_UNITS(³æ¦ì¼Æ/¤u®É)
               If strLedes(10, iUpper) = "F" Then
                  strLedes(11, iUpper) = ""
                  'Added by Morgan 2013/6/27
                  'BASF ªº TimeKeeper Rate ­n¥Î US$455
                  'Modified by Morgan 2014/1/3 §ï§ì³]©w­È
                  'If m_A1k28 = "Y45814010" Then
                  '   strLedes(21, iUpper) = 455
                  'Added by Morgan 2018/6/25 +Dow¯S®í³]©w(«D¤¤¶¡µ{§Ç§ì¹w³]³]©w TIMEKEEPER_ID=DY,TIMEKEEPER_NAME=David Yen, LINE_ITEM_UNIT_COST=10)
                  'Modified by Morgan 2019/1/4 +Dow¤ñ·Ó¯È¥»¦A¥[§PÂ_½Ð´Ú¶µ¥Ø Ex:X10719747
                  'If m_bSpecial1 And Not m_bDowX Then
                  If m_bSpecial1 And Not (m_bDowX And ChkDowXFormat(.Fields("a1j02"), True) = True) Then
                     'Added by Morgan 2019/5/27 Y55199 (Dow AgroSciences LLC) ©T©w³ø»ù
                     'Modified by Morgan 2020/3/24 +Y48048 Dow Silicones Corporation¡BY22457 THE DOW CHEMICAL COMPANY¡BY55240 DuPont --Lisa
                     'Modified by Morgan 2020/4/21 +Y52322 (Dow Toray Co., Ltd.)--Tim
                     'Modified by Morgan 2022/3/4 +Y55423 --Kimi
                     'Modified by Morgan 2025/6/30 +Y52322B10 --Tim
                     If m_A1k28 = "Y55199000" Or m_A1k28 = "Y48048000" Or m_A1k28 = "Y22457000" Or m_A1k28 = "Y55240000" Or m_A1k28 = "Y52322000" Or m_A1k28 = "Y55423000" Or m_A1k28 = "Y52322B10" Then
                        strLedes(21, iUpper) = Val(strLedes(13, iUpper)) - Val(strLedes(12, iUpper))
                        strLedes(11, iUpper) = "1"
                        
                        'Added by Morgan 2020/4/9 Lisa
                        '½Ð´Úª÷ÃB¶W¹L100®É§ï¥Î75
                        If Val(strLedes(21, iUpper)) > 100 Then
                           strLedes(21, iUpper) = 75
                           strLedes(11, iUpper) = ""
                        End If
                        'end 2020/4/9
                     Else
                     'end 2019/5/27
                        strLedes(21, iUpper) = 10
                     End If
                  'end 2018/6/25
                  
                  'Added by Morgan 2019/11/21 ¦] syngenta ®É¼Æ¥u¯à¤p¼Æ1¦ì,§ï®É¼Æ©T©w¥Î1¶O²v³]½Ð´Úª÷ÃB --²øÞ±¤Z
                  'Modified by Morgan 2023/8/29 +Y54096000 ©T©w³ø»ù,RateµL¤W­­--Tim
                  'Modified by Morgan 2024/7/9 +Y53893000 --Kahn
                  'Modified by Morgan 2024/9/2 -Y53893000,§ï¤U­±³æ¿W³]©w --Kahn
                  'Modified by Morgan 2025/4/21 +Y56065000 ©T©w³ø»ù--Tim
                  'Modified by Morgan 2025/8/5 +X28716030 --Tim
                  ElseIf m_A1k28 = "Y54225B10" Or m_A1k28 = "Y54096000" Or m_A1k28 = "Y56065000" Or m_A1k28 = "X28716030" Then
                     strLedes(21, iUpper) = Val(strLedes(13, iUpper)) - Val(strLedes(12, iUpper))
                     strLedes(11, iUpper) = "1"
                  'end 2019/11/13
                  'Added by Morgan 2021/5/25
                  'Modified by Morgan 2025/3/14 +Y56142000--Franny
                  ElseIf m_A1k28 = "Y19893030" Or m_A1k28 = "Y56142000" Then
                     If InStr("'1202','205','1002','107','203','431','422','204','903'", "'" & .Fields("a1j02") & "'") > 0 Then
                        strLedes(21, iUpper) = adoLEDES.Fields("ld16")
                     Else
                        strLedes(21, iUpper) = Val(strLedes(13, iUpper)) - Val(strLedes(12, iUpper))
                        strLedes(11, iUpper) = "1"
                     End If
                  'end 2021/5/25
                  ElseIf adoLEDES.Fields("ld16") > 0 Then
                     strLedes(21, iUpper) = adoLEDES.Fields("ld16")
                  'end 2014/1/3
                  Else
                  'end 2013/6/27
                     'Hourly rate
                     strLedes(21, iUpper) = 10
                  End If
                  
                  
                  'Added by Morgan 2024/9/2 Y53893000 «D¤¤¶¡µ{§Ç
                  'Modified by Morgan 2024/10/7 +Y55014 Nalco, an Ecolab Company --Kahn
                  'Modified by Morgan 2025/5/19 +X56347000 (Entegris, Inc.)--Tim
                  '1.¦U¦¬¤å©Ê½è (°£¤F¤U¦C¤¤¶¡µ{§Ç©Ê½è)ªºUnits (LINE_ITEM_NUMBER_OF_UNITS):³]©w©T©w¬°1¡AUNIT COST:½Ð³]©T©wUSD280 (USD350*80%) (­ì¥ý³]©w¬O¨Ì¥»©Ò¯u¹ê½Ð´Úª÷ÃB)¡A¨Ã¼W¥[½Õ¾ãÄæ¦ìADJUSTMENT (LINE_ITEM_ADJUSTMENT_AM OUNT):½Õ¾ã¦¨¥»©Ò½Ð´Úª÷ÃB
                  '2.¤¤¶¡µ{§Ç (CÃþ³ø§i/1202; ®Ö»é/1002; ¥Ó´_/205; ¦A¼f/107 ; ¥D°Ê­×¥¿/203; ³qª¾­×¥¿/1201 ; ­×¥¿/204 ; ½Ð¨D­±¸ß/407 ; ­±¸ß/408 ; ¥[³t¼f¬d/422 ; °ª³t¼f¬d/431; ³qª¾¾Ü¤@¥Ó´_/1232 ; ¾Ü¤@¥Ó´_/239; ¨ÌÂ¾Åv¹q¸Ü³qª¾­×¥¿/1225;§ó¥¿/402; ³Ì«á³qª¾)¡A½Ð¨ó§U±NUnits Cost³]©wUSD280 (USD350*80%)¡AUnits«h¹ïÀ³¯B°Ê (¶È¥i¨ì¤p¼ÆÂI«á²Ä¤@¦ì)
                  If m_A1k28 = "Y53893000" Or m_A1k28 = "Y55014000" Or m_A1k28 = "X56347000" Then
                     'Modified by Morgan 2024/11/13 +447¦A¼f¬d¥[³t¼f¬d
                     If InStr("'1202','1002','205','107','203','1201','204','407','408','422','431','447','1232','239','1225','402','1227'", "'" & .Fields("a1j02") & "'") = 0 Then
                        strLedes(11, iUpper) = "1"
                        strLedes(12, iUpper) = Val(strLedes(13, iUpper)) - Val(strLedes(21, iUpper))
                        
                        'Added by Morgan 2025/5/19
                        If m_A1k28 = "X56347000" Then
                           strLedes(19, iUpper) = "(Flate Fee)"
                        End If
                        'end 2025/5/19
                     End If
                  End If
                  
                  If strLedes(11, iUpper) = "" Then
                     'Modified by Morgan 2014/8/5
                     'strLedes(11, iUpper) = Round((Val(strLedes(13, iUpper)) - Val(strLedes(12, iUpper))) / Val(strLedes(21, iUpper)), 4)
                     'Rockwell ªºRate¦³½Õ¾ã,­Y½Ð´Úª÷ÃB­Ó¦ì¼Æ¬°©_¼Æ®É·|¦³¤p¼Æ¨â¦ì·|¦³¿ù»~,¬G§ï¬°¥|±Ë¤­¤J¨ì¤p¼Æ1¦ì¦A¥H¦¹­«ºâ½Ð´Úª÷ÃB
                     If m_A1k28 = "Y46295000" Then
                        strLedes(11, iUpper) = Round((Val(strLedes(13, iUpper)) - Val(strLedes(12, iUpper))) / Val(strLedes(21, iUpper)), 1)
                        strLedes(13, iUpper) = Val(strLedes(11, iUpper)) * Val(strLedes(21, iUpper)) + Val(strLedes(12, iUpper))
                     Else
                        strLedes(11, iUpper) = Round((Val(strLedes(13, iUpper)) - Val(strLedes(12, iUpper))) / Val(strLedes(21, iUpper)), 4)
                     End If
                     'end 2014/8/5
                  End If
               Else
                  'Modified by Morgan 2012/5/31 À³©ñ§é¦©«eª÷ÃB
                  'strLedes(21, iUpper) = strLedes(13, iUpper)
                  strLedes(21, iUpper) = Val(strLedes(13, iUpper)) - Val(strLedes(12, iUpper))
                  strLedes(11, iUpper) = "1"
               End If
            End If
            
            '14 LINE_ITEM_DATE
            'Added by Morgan 2018/6/25 +Dow¯S®í³]©w(§ìµo¤å¤é)
            If m_bSpecial1 Then
               strLedes(14, iUpper) = GetCp27(.Fields("a1k01"), .Fields("a1j02"))
            End If
            'end 2018/6/25
            If strLedes(14, iUpper) = "" Then
               strLedes(14, iUpper) = strLedes(1, iUpper)
            End If
            
            '15 LINE_ITEM_TASK_CODE
            'Modified by Morgan 2020/8/3 X70124 (Nordson Corporation)°Ó¼Ð®×¥ó¹w³]TR630--§õÞ³®¦
            If m_A1k28 = "X70124000" And .Fields("a1k13") = "FCT" And Not IsNull(.Fields("a1j18")) And IsNull(.Fields("a2604")) Then
               strLedes(15, iUpper) = "TR630"
            Else
               strLedes(15, iUpper) = "" & .Fields("a1j18")
            End If
            
            'Added by Morgan 2025/1/16
            'ÀË¬d¬O§_¬°¤À³Î®×¥B¥t¦³³]©wTASK_CODE
            If strLedes(15, iUpper) <> "" Then
               SetDivCaseCode .Fields("a1k13"), .Fields("a1k14"), .Fields("a1k15"), .Fields("a1k16"), .Fields("a1k27"), .Fields("a1j02"), strLedes(15, iUpper)
            End If
            'end 2025/1/16
            
            '16 LINE_ITEM_EXPENSE_CODE
            strLedes(16, iUpper) = "" & .Fields("a1j19")
            'Added by Morgan 2017/10/18 HPªº¶W­¶¶W¶µ¶O­Y»P¹ê¼f¤@°_½Ð´Ú®É¥Î¹ê¼f³W¶OªºCode
            If m_A1k28 = "Y48292000" And (.Fields("a1j02") = "93899" Or .Fields("a1j02") = "93999") Then
               Set AdoRecordSet3 = adoacc1k0.Clone
               AdoRecordSet3.Find "a1k01='" & .Fields("a1k01") & "'"
               If Not AdoRecordSet3.EOF Then
                  AdoRecordSet3.Find "a1j02='41699'"
                  If Not AdoRecordSet3.EOF Then
                     strLedes(16, iUpper) = "" & AdoRecordSet3.Fields("a1j19")
                  End If
               End If
            End If
            'end 2017/10/18
            
            '17 LINE_ITEM_ACTIVITY_CODE
            strLedes(17, iUpper) = "" & .Fields("a1j20")
            
            '18 TIMEKEEPER_ID
            '22 TIMEKEEPER_NAME
            '23 TIMEKEEPER_CLASSIFICATION
            If strLedes(10, iUpper) = "F" Then
               'Modified by Morgan 2012/4/24 LEDED ³]©w§ï§ì LEDED ¸ê®Æªí
               'strLedes(18, iUpper) = "" & .Fields("fa104")
               strLedes(18, iUpper) = "" & adoLEDES.Fields("ld11")
               'Added by Morgan 2018/6/25 +Dow¯S®í³]©w(«D¤¤¶¡µ{§Ç§ì¹w³]³]©w TIMEKEEPER_ID=DY,TIMEKEEPER_NAME=David Yen, LINE_ITEM_UNIT_COST=10)
               'Modified by Morgan 2019/1/4 +Dow¤ñ·Ó¯È¥»¦A¥[§PÂ_½Ð´Ú¶µ¥Ø Ex:X10719747
               'If m_bSpecial1 And Not m_bDowX Then
               If m_bSpecial1 And Not (m_bDowX And ChkDowXFormat(.Fields("a1j02"), True) = True) Then
                  'Added by Morgan 2019/5/27 Y55199 (Dow AgroSciences LLC)¯S®í³]©w
                  If m_A1k28 = "Y55199000" Then
                     strLedes(18, iUpper) = "NONE"
                     strLedes(22, iUpper) = "NONE"
                  Else
                  'end 2019/5/27
                     strLedes(18, iUpper) = "DY"
                     strLedes(22, iUpper) = "David Yen"
                  End If
               'end 2018/6/25
               ElseIf strLedes(18, iUpper) <> "" Then
                  'Modified by Morgan 2012/4/24 LEDED ³]©w§ï§ì LEDED ¸ê®Æªí
                  'If strLedes(18, iUpper) = "DY" Then
                  '   strLedes(22, iUpper) = "David Yen"
                  'Else
                  strLedes(22, iUpper) = "" & adoLEDES.Fields("ld12")
                  If strLedes(22, iUpper) = "" Then
                  'end 2012/4/24
                     strLedes(22, iUpper) = strLedes(18, iUpper)
                  End If
                  
               Else
                  'Modify by Morgan 2011/5/12 ¥~°Ó¹w³] Fred,¥~±M¹w³] David
                  'Modified by Morgan 2022/10/14 +T,TS
                  If .Fields("a1k13") = "FCT" Or .Fields("a1k13") = "S" Or .Fields("a1k13") = "FCL" Or .Fields("a1k13") = "TS" Or .Fields("a1k13") = "T" Then
                     strLedes(18, iUpper) = "Fred Yen"
                     strLedes(22, iUpper) = "Fred Yen"
                  Else
                     strLedes(18, iUpper) = "DY"
                     strLedes(22, iUpper) = "David Yen"
                  End If
               End If
                  
               'Modified by Morgan 2012/4/24 LEDED ³]©w§ï§ì LEDED ¸ê®Æªí
               ''Add by Morgan 2011/8/11 --³¯©É»T
               'If m_A1k28 = "Y49165000" Then
               '   strLedes(23, iUpper) = "LA"
               'Else
               ''end 2011/8/11
               '   strLedes(23, iUpper) = "PT" '©T©w,­Y»Ý«ü©w¦A¥[Äæ¦ì
               'End If
               strLedes(23, iUpper) = "" & adoLEDES.Fields("ld13")
               If strLedes(23, iUpper) = "" Then
                  strLedes(23, iUpper) = "PT"
               End If
               'end 2012/4/24
            End If
               
            '19 LINE_ITEM_DESCRIPTION
            'Modify by Morgan 2011/5/19 ¥~°Ó§é¦©¥t¦C¶µ¥Ø»¡©ú±aÂ²³æª©
            'Modified by Morgan 2022/10/14 +T,TS
            If .Fields("a1k13") = "FCT" Or .Fields("a1k13") = "S" Or .Fields("a1k13") = "T" Or .Fields("a1k13") = "TS" Then
               If m_ClearItemDesc <> "" Then
                  strLedes(19, iUpper) = m_ClearItemDesc
               Else
                  strLedes(19, iUpper) = "" & .Fields("a1j04")
               End If
            Else
               'Added by Morgan 2018/9/18
               'Dow§é¦©¥t¦C¶µ¥Ø»¡©ú±aÂ²³æª©
               'Modified by Morgan 2020/2/17 Y22457000,Y48048000,Y52322000 ¨ú®ø--Tim
               If m_bSpecial1 And m_bDowN = False Then
                  strLedes(19, iUpper) = m_ClearItemDesc
               'Added by Morgan 2023/10/13
               'BeiGene §é¦©Á`©M¥t¦C³Ì«á,¶µ¥Ø»¡©ú±aÂ²³æª©
               ElseIf m_A1k28 = "X76928000" Then
                  strLedes(19, iUpper) = m_ClearItemDesc
               'end 2023/10/13
               Else
                  
                  'Modified by Morgan 2025/5/19 «e­±¥i¯à¦³³]©wªþ¥[¤å¦r
                  'If m_ItemDesc <> "" Then
                  '   strLedes(19, iUpper) = m_ItemDesc
                  'Else
                  '   strLedes(19, iUpper) = .Fields("a1j04")
                  'End If
                  If m_ItemDesc <> "" Then
                     strLedes(19, iUpper) = strLedes(19, iUpper) & m_ItemDesc
                  Else
                     strLedes(19, iUpper) = strLedes(19, iUpper) & .Fields("a1j04")
                  End If
                  'end 2025/5/19
               End If
               
               'Added by Morgan 2015/10/30
               'Mondelez Global LLC­n¨D±NAnaqua number(«È¤á®×¥ó®×¸¹)©ñ¦bwork description¤¤
               If m_A1k28 = "Y54037000" Then
                  strLedes(19, iUpper) = GetCustCaseNo(.Fields("a1k13"), .Fields("a1k14"), .Fields("a1k15"), .Fields("a1k16")) & " " & strLedes(19, iUpper)
               'Added by Morgan 2025/7/11 LINE_ITEM_DESCRIPTION «e­±¤]­n¥[¥Ó½Ð¸¹¤Î¥»©Ò¸¹(¦PINVOICE_DESCRIPTION)--Franny
               ElseIf m_A1k28 = "X82916010" Then
                  strLedes(19, iUpper) = "(Application No. " & strAppNo & ")(" & .Fields("a1k13") & "-" & .Fields("a1k14") & IIf(.Fields("a1k15") & .Fields("a1k16") = "000", "", "-" & .Fields("a1k15") & "-" & .Fields("a1k16")) & ")" & strLedes(19, iUpper)
               End If
               'end 2015/10/30
            End If
            'Added by Morgan 2018/6/27 Syngenta ªº LINE_ITEM_DESCRIPTION «e­±­n¥[[©¼©Ò®×¸¹]
            'Modified by Morgan 2021/4/9 +Y52216000
            If m_A1k28 = "Y54225B10" Or m_A1k28 = "Y48309070" Or m_A1k28 = "Y52216000" Then
               strExc(1) = GetYourRefNo1(.Fields("a1k13"), .Fields("a1k14"), .Fields("a1k15"), .Fields("a1k16"))
               
               'Added by Morgan 2020/3/17 FG/FCL/LIN®×¶·¥t±a¥X1st+2ndÁpµ¸¤H¦WºÙ--Franny
               If m_A1k28 = "Y54225B10" Then
                  'Added by Morgan 2022/2/10 ­Y¦³«È¤á®×¥ó®×¸¹®ÉÀu¥ý --Ryan
                  strExc(2) = GetCustCaseNo(.Fields("a1k13"), .Fields("a1k14"), .Fields("a1k15"), .Fields("a1k16"))
                  If strExc(2) <> "" Then strExc(1) = strExc(2)
                  'end 2022/2/10
                  
                  'Modified by Morgan 2021/2/4 ¤w§ï¥Î98BI¥B¨ú®ø©¼¸¹«e«áªº[]
                  If (.Fields("a1k13") = "FG" Or .Fields("a1k13") = "FCL" Or .Fields("a1k13") = "LIN") Then
                     strLedes(19, iUpper) = strExc(1) & " [" & m_Att1 & IIf(m_Att2 <> "", " and " & m_Att2, "") & "]" & strLedes(19, iUpper)
                  Else
                     strLedes(19, iUpper) = strExc(1) & " " & strLedes(19, iUpper)
                  End If
                  'end 2021/2/4
               Else
               'end 2020/3/17
                  strLedes(19, iUpper) = "[" & strExc(1) & "]" & strLedes(19, iUpper)
               End If
               
            'Added by Morgan 2018/10/8 Serengeti Y52878B10 ªº LINE_ITEM_DESCRIPTION ¥[ Your Ref: ©¼©Ò®×¸¹ ¨ì«á­±
            ElseIf m_A1k28 = "Y52878B10" Then
               strLedes(19, iUpper) = strLedes(19, iUpper) & " Your Ref: " & GetYourRefNo1(.Fields("a1k13"), .Fields("a1k14"), .Fields("a1k15"), .Fields("a1k16"))
            'Added by Morgan 2023/7/10
            ElseIf m_A1k28 = "Y48279000" Then
               strLedes(19, iUpper) = strLedes(19, iUpper) & " for Application No. " & strAppNo
            End If
            'end 2018/6/27
         
            'Add by Morgan 2011/3/14
            If m_iLedesVer = 2 Then
               
               '25 PO_NUMBER
               'Modified by Morgan 2021/1/5 Y54225B10 §ïBI®æ¦¡
               'strLedes(25, iUpper) = ""
               strLedes(25, iUpper) = "" & adoLEDES.Fields("ld17")
               
               'Added by Morgan 2013/5/27 BASF¦PCLIENT_MATTER_ID
               'Modified by Morgan 2014/1/20 BASF¥Î¹q¤l±b³æªº¥N½X¤w§ï
               'If m_A1k28 = "Y45814010" Then
               If m_A1k28 = "Y33268010" Then
                  strLedes(25, iUpper) = strLedes(24, iUpper)
               End If
               'end 2013/5/27
               '27 MATTER_NAME
               'Modify by Morgan 2011/5/12 ¥~°Ó±a®×¥ó¦WºÙ
               If .Fields("a1k13") = "FCT" Or .Fields("a1k13") = "S" Then
                  strLedes(27, iUpper) = GetCaseName(m_CP01 & m_CP02 & m_CP03 & m_CP04)
               Else
                  strLedes(27, iUpper) = "Taiwan patent/trademark"
               End If
               '28 INVOICE_TAX_TOTAL
               strLedes(28, iUpper) = ""
               '29 INVOICE_NET_TOTAL(½Ð´ÚÁ`ÃB)
               strLedes(29, iUpper) = strLedes(5, iUpper)
               '30 INVOICE_CURRENCY
               strLedes(30, iUpper) = "USD"
               If strLedes(22, iUpper) <> "" Then
                  'Added by Morgan 2013/5/28
                  If InStr(strLedes(22, iUpper), ",") > 0 Then
                     arrName = Split(strLedes(22, iUpper), ",")
                     '31 TIMEKEEPER_LAST_NAME
                     strLedes(31, iUpper) = Trim(arrName(LBound(arrName)))
                     '32 TIMEKEEPER_FIRST_NAME
                     strLedes(32, iUpper) = Trim(arrName(UBound(arrName)))
                  Else
                     arrName = Split(strLedes(22, iUpper), " ")
                     '31 TIMEKEEPER_LAST_NAME
                     strLedes(31, iUpper) = arrName(UBound(arrName))
                     '32 TIMEKEEPER_FIRST_NAME
                     strLedes(32, iUpper) = arrName(LBound(arrName))
                  End If
               End If
               '33 ACCOUNT_TYPE
               strLedes(33, iUpper) = "O"
               '34 LAW_FIRM_NAME
               strLedes(34, iUpper) = "Tai E International Patent & Law Office"
               '35 LAW_FIRM_ADDRESS_1
               strLedes(35, iUpper) = "9Fl., No. 112,"
               '36 LAW_FIRM_ADDRESS_2
               strLedes(36, iUpper) = "Sec. 2, Chang-An E. Rd"
               '37 LAW_FIRM_CITY
               strLedes(37, iUpper) = "Taipei"
               '38 LAW_FIRM_STATEorREGION
               strLedes(38, iUpper) = ""
               '39 LAW_FIRM_POSTCODE
               strLedes(39, iUpper) = "10491"
               '40 LAW_FIRM_COUNTRY
               strLedes(40, iUpper) = "TWN"
                  
'Modified by Morgan 2012/4/24 LEDED ³]©w§ï§ì LEDED ¸ê®Æªí

               '26 CLIENT_TAX_ID
               strLedes(26, iUpper) = "" & adoLEDES.Fields("ld03")
               '41 CLIENT_NAME
               strLedes(41, iUpper) = "" & adoLEDES.Fields("ld04")
               '42 CLIENT_ADDRESS_1
               strLedes(42, iUpper) = "" & adoLEDES.Fields("ld05")
               '43 CLIENT_ADDRESS_2
               strLedes(43, iUpper) = "" & adoLEDES.Fields("ld06")
               '44 CLIENT_CITY
               strLedes(44, iUpper) = "" & adoLEDES.Fields("ld07")
               '45 CLIENT_STATEorREGION
               strLedes(45, iUpper) = "" & adoLEDES.Fields("ld08")
               '46 CLIENT_POSTCODE
               strLedes(46, iUpper) = "" & adoLEDES.Fields("ld09")
               '47 CLIENT_COUNTRY
               strLedes(47, iUpper) = "" & adoLEDES.Fields("ld10")
'end 2012/4/24
                  
               '48 LINE_ITEM_TAX_RATE
               strLedes(48, iUpper) = ""
               '49 LINE_ITEM_TAX_TOTAL
               '= (LINE_ITEM_UNIT_COST * LINE_ITEM_NUMBER_OF_UNITS + LINE_ITEM_ADJUSTMENT_AMOUNT) * LINE_ITEM_TAX_RATE
               'Modify by Morgan 2011/5/12 ¥~°Ó¦³¥N²z¤H­n¨D¤£­n¶ñ,§ï±±¨î°Ó¼Ð³£ªÅµÛ
               'Modified by Morgan 2017/9/8 X18064010 ¤£¯àªÅ --¸U§Ó¼w
               'Modified by Morgan 2017/12/25 Y54688000 ¤£¯àªÅ --§d°ê¦w
               If m_A1k28 = "X18064010" Or m_A1k28 = "Y54688000" Then
                  strLedes(49, iUpper) = "0"
               ElseIf (.Fields("a1k13") = "FCT" Or .Fields("a1k13") = "S") Then
                  strLedes(49, iUpper) = ""
               Else
                  strLedes(49, iUpper) = "0"
               End If
               '50 LINE_ITEM_TAX_TYPE
               strLedes(50, iUpper) = ""
               '51 INVOICE_REPORTED_TAX_TOTAL
               strLedes(51, iUpper) = ""
               '52 INVOICE_TAX_CURRENCY
               strLedes(52, iUpper) = ""
                              
               'Added by Morgan 2018/8/30 BASF¯S®í»Ý¨D --Lisa
               If (m_A1k28 = "Y33268010" And .Fields("a1k13") = "FCP") Then
                  UpdateLEDES
               End If
               'end 2018/8/30
               
               'Added by Morgan 2021/2/4
               If m_A1k28 = "Y54225B10" And (.Fields("a1k13") = "FCP" Or .Fields("a1k13") = "FG") Then
                  UpdateLEDES2
               End If
               'end 2021/2/4
               
            'Added by Morgan 2023/6/19 +Y48279000,X48279000 --Kahn
            Else
               If (m_A1k28 = "Y48279000" Or m_A1k28 = "X48279000") And (.Fields("a1k13") = "FCP" Or .Fields("a1k13") = "P") Then
                  UpdateLEDES
               End If
            'end 2023/6/19
            End If
            
            UpdateLEDES3 .Fields("a1k13"), .Fields("a1k14"), .Fields("a1k15"), .Fields("a1k16"), .Fields("a1j02") 'Added by Morgan 2020/5/21
            
            'Add by Morgan 2011/5/12 §é¦©¥t¦C
            'Modified by Morgan 2018/9/18 Dow§é¦©³Ì«á¦C
            'If douUSAmountNoDisc <> douUSAmount Then
            'Modified by Morgan 2023/10/18 ±Æ°£§é¦©Á`­p©ó³Ì«á¦CªÌ( And douDiscount = 0)
            If douUSAmountNoDisc <> douUSAmount And m_bSpecial1 = False And douDiscount = 0 Then
            'end 2018/9/18
               If Val(strLedes(12, iUpper)) = 0 Then
                  iUpper = iUpper + 1
                  ReDim Preserve strLedes(m_iCols, iUpper)
                  '1~8
                  For intI = 1 To 8
                     strLedes(intI, iUpper) = strLedes(intI, iUpper - 1)
                  Next
                  strLedes(9, iUpper) = iUpper
                  strLedes(10, iUpper) = "IF"
                  strLedes(11, iUpper) = ""
                  strLedes(12, iUpper) = Round(douUSAmount - douUSAmountNoDisc, 4)
                  strLedes(13, iUpper) = strLedes(12, iUpper)
                  strLedes(14, iUpper) = strLedes(14, iUpper - 1)
                  strLedes(15, iUpper) = "IF999"
                  strLedes(16, iUpper) = ""
                  strLedes(17, iUpper) = ""
                  strLedes(18, iUpper) = strLedes(18, iUpper - 1)
                  strLedes(19, iUpper) = "Discount"
                  strLedes(20, iUpper) = strLedes(20, iUpper - 1)
                  strLedes(21, iUpper) = ""
                  '22~
                  For intI = 22 To m_iCols
                     strLedes(intI, iUpper) = strLedes(intI, iUpper - 1)
                  Next
                  
               End If
            End If
            End With
         End If
      End If
      '¶µ¥Ø¶¡­n¦hªÅ¤@¦æ
      intCounter = intCounter + 2
'Added by Lydia 2015/04/09
Chi_JumpPDM:
      adoacc1k0.MoveNext
   Loop
   
   'Added by Morgan 2023/10/18
   '§é¦©Á`©M¥t¦C³Ì«áªÌ
   If m_bEBilling And douDiscount <> 0 Then
      iUpper = iUpper + 1
      ReDim Preserve strLedes(m_iCols, iUpper)
      For intI = 1 To 8
         strLedes(intI, iUpper) = strLedes(intI, iUpper - 1)
      Next
      strLedes(9, iUpper) = iUpper
      strLedes(10, iUpper) = "IF"
      strLedes(11, iUpper) = ""
      strLedes(12, iUpper) = Round(douDiscount, 4)
      strLedes(13, iUpper) = strLedes(12, iUpper)
      strLedes(14, iUpper) = strLedes(14, iUpper - 1)
      strLedes(15, iUpper) = ""
      strLedes(16, iUpper) = ""
      strLedes(17, iUpper) = ""
      strLedes(18, iUpper) = strLedes(18, iUpper - 1)
      strLedes(19, iUpper) = "Discount"
      strLedes(20, iUpper) = strLedes(20, iUpper - 1)
      strLedes(21, iUpper) = ""
      '22~
      For intI = 22 To m_iCols
         strLedes(intI, iUpper) = strLedes(intI, iUpper - 1)
      Next
   End If
   douDiscount = 0
   'end 2023/10/18
   
   If douAmount <> 0 Then
   
      If m_b2Printer Then
         Printer.Line (0 + intDefault, m_DetailTopStart + intCounter * m_LineH - 200 + intTop)-(10000 + intDefault, m_DetailTopStart + intCounter * m_LineH - 200 + intTop)
      End If
      'Removed by Morgan 2022/8/4
      'If m_b2Picture Then
      '   Picture1.Line (douExtRate * (0 + intDefault), douExtRate * (m_DetailTopStart + intCounter * m_LineH - 200 + intTop))-(douExtRate * (10000 + intDefault), douExtRate * (m_DetailTopStart + intCounter * m_LineH - 200 + intTop))
      'End If
      'end 2022/8/4
      PrintSum
      
      If Check2.Value = vbChecked Then GoTo CloseFlag 'Added by Morgan 2014/8/5
      
      If m_b2Printer And Not m_b2PDF And Not m_b2Picture Then
         '·s¼W¦a§}±ø¦Cªí¸ê®Æ
         If Me.txtAdd.Text = "Y" Then
            'Modified by Morgan 2012/10/3
            '¾ã§å½Ð´Ú¥u­n¦C¦L1±i
            If Not (m_bolOneAddr And m_bolAddAddrOK) Then
               pub_AddressListSN = pub_AddressListSN + 1
               PUB_AddNewAddressList strUserNum, m_CP01, m_CP02, m_CP03, m_CP04, "" & pub_AddressListSN, "0", GetCP10(strA1K01)
               m_bolAddAddrOK = True
            End If
         End If
      End If
   End If
      
   If m_b2Printer Then
      Printer.EndDoc
      'Added by Morgan 2012/10/31
      If m_bPrint2Pdf Then
         frmPDF.EndtProcess
         'Added by Lydia 2020/02/15 §PÂ_ÀÉ®×¬O§_¦s¦b, ¶W¹L®É¶¡´NÄ~Äò
         If Me.Tag <> "" And InStr(Me.Tag, "\") > 0 Then
            'Modified by Lydia 2020/09/10 ¶W¹L®É¶¡¡Aª½±µ°O¿ý¥¢±Ñ²M³æ
            'If PUB_ChkFileStatus(Me.Tag) = False Then
            If PUB_ChkFileStatus(Me.Tag, False, m_strOutErr) = False Then
            End If
         End If
         'end 2020/02/15
         Unload frmPDF
      End If
   End If

   'Removed by Morgan 2022/8/4
   'If m_b2Picture Then
   '   If m_iPages > 1 Then
   '      SetPic m_iPages, True
   '   Else
   '      SetPic 0, True
   '   End If
   'End If
   'end 2022/8/4
   
   'Added by Morgan 2012/11/6
   '­Y­nÂàPDF
   'Modified by Morgan 2014/8/26
   'If m_b2PDF Then
   'Added by Lydia 2015/04/10 §PÂ_¬O§_¥u¿é¥Xword,«K¤£ÂàPDF
  ' If m_Chi2Word = False Then
        If m_b2PDF And Not m_PdfDone Then
        'end 2014/8/26
           '¥¼Âà
           If Not m_bPrint2Pdf Then
              m_bPrint2Pdf = True
              GoTo PrintPdfStart
           End If
        End If
        'end 2012/11/6
 '  End If
   'end 2015/04/10
CloseFlag:

   adoacc1k0.Close
   
   'Add by Morgan 2010/11/19
   If m_bMsg Then
      'Modified by Morgan 2012/10/31
      'If m_bSaveWord Or m_b2Picture Then
      If m_bSaveWord Or m_b2Picture Or m_b2PDF Then
         MsgBox "¹q¤lÀÉ¤w¦s©ó [ " & m_EFilePath & " ]¡I"
      End If
      
      If m_sEBillingMsg <> "" Then
         If MsgBox(m_sEBillingMsg & vbCrLf & vbCrLf & "¬O§_­n¦C¦L¥»°T®§??", vbYesNo) = vbYes Then
            'Modified by Morgan 2022/8/4
            'Printer.Print m_sEBillingMsg
            PUB_PrintUnicodeText m_sEBillingMsg, Printer.CurrentX, Printer.CurrentY, 0
            'end 2022/8/4
            Printer.EndDoc
         End If
      End If
   End If
   
OnlySpecial:
   'Add by Morgan 2011/10/11
'Added by Lydia 2015/04/15
   'If strSpecialList <> "" Then MsgBox "¤U¦C³æ¸¹¬°¯S®í½Ð´Ú³æ¤£¥i¦C¦L¡I" & vbCrLf & vbCrLf & strSpecialList
   strExc(7) = ""
   If strSpecialList <> "" Then strExc(7) = "¤U¦C³æ¸¹¬°¯S®í½Ð´Ú³æ¤£¥i¦C¦L¡I" & vbCrLf & vbCrLf & strSpecialList
   If strSpecialL2 <> "" Then
     If Len(strExc(7)) > 0 Then strExc(7) = strExc(7) & vbCrLf & vbCrLf
     strExc(7) = strExc(7) & "¤U¦C³æ¸¹¬°¾ã§å½Ð´Ú³æ¤£¥i¦C¦L¡I" & vbCrLf & vbCrLf & strSpecialL2
   End If
   If strExc(7) <> "" Then MsgBox strExc(7)
   
   Exit Sub 'Added by Morgan 2020/8/19 ¦]¥i¯à·|´Ý¯d¥i©¿²¤ªº¿ù»~¨Æ¥ó
   
Checking:
   If Err.Number <> 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If
End Sub

'Add by Morgan 2009/11/23
'PrintHead µ{§Ç¤Ó¤j©â¥X
Private Sub PrintCustData(strCustName() As String, strLanguage)
   Dim ii As Integer
   ii = 0
   While Not adoquery.EOF
      Select Case strLanguage
         Case "1"
            If IsNull(adoquery.Fields("cu04").Value) = False Then
               If "" & adoacc1k0.Fields("a1k04").Value = "Y" Then
                  If ii > 0 Then intRow = intRow + 1 'Add by Morgan 2010/11/29
                  PutData adoquery.Fields("cu04").Value, intRow
                  SetWordArray m_Head, intRow, 2, adoquery.Fields("cu04").Value 'Add by Morgan 2010/12/1
               End If
               strCustName(ii) = strCustName(ii) & "" & adoquery.Fields("cu04").Value
            End If
            
         Case "2"
            If IsNull(adoquery.Fields("cu05").Value) = False Then
               If "" & adoacc1k0.Fields("a1k04").Value = "Y" Then
                  If ii > 0 Then intRow = intRow + 1 'Add by Morgan 2010/11/29
                  PutData adoquery.Fields("cu05").Value, intRow
                  SetWordArray m_Head, intRow, 2, adoquery.Fields("cu05").Value 'Add by Morgan 2010/11/24
               End If
               strCustName(ii) = strCustName(ii) & "" & adoquery.Fields("cu05").Value
               If IsNull(adoquery.Fields("cu88").Value) = False Then
                  If "" & adoacc1k0.Fields("a1k04").Value = "Y" Then
                     intRow = intRow + 1
                     PutData adoquery.Fields("cu88").Value, intRow
                     SetWordArray m_Head, intRow, 2, adoquery.Fields("cu88").Value 'Add by Morgan 2010/11/24
                  End If
                  strCustName(ii) = strCustName(ii) & " " & adoquery.Fields("cu88").Value
               End If
               If IsNull(adoquery.Fields("cu89").Value) = False Then
                  If "" & adoacc1k0.Fields("a1k04").Value = "Y" Then
                     intRow = intRow + 1
                     PutData adoquery.Fields("cu89").Value, intRow
                     SetWordArray m_Head, intRow, 2, adoquery.Fields("cu89").Value 'Add by Morgan 2010/11/24
                  End If
                  strCustName(ii) = strCustName(ii) & " " & adoquery.Fields("cu89").Value
               End If
               If IsNull(adoquery.Fields("cu90").Value) = False Then
                  If "" & adoacc1k0.Fields("a1k04").Value = "Y" Then
                     intRow = intRow + 1
                     PutData adoquery.Fields("cu90").Value, intRow
                     SetWordArray m_Head, intRow, 2, adoquery.Fields("cu90").Value 'Add by Morgan 2010/11/24
                  End If
                  strCustName(ii) = strCustName(ii) & " " & adoquery.Fields("cu90").Value
               End If
               
            ElseIf IsNull(adoquery.Fields("cu06").Value) = False Then
               '­Y¬°FCP®×
               If "" & adoacc1k0.Fields("A1K13").Value = "FCP" Then
                  If "" & adoacc1k0.Fields("a1k04").Value = "Y" Then
                     If ii > 0 Then intRow = intRow + 1 'Add by Morgan 2010/11/29
                     PutData adoquery.Fields("cu06").Value, intRow
                     SetWordArray m_Head, intRow, 2, adoquery.Fields("cu06").Value 'Add by Morgan 2010/11/24
                  End If
                  strCustName(ii) = strCustName(ii) & "" & adoquery.Fields("cu06").Value
               End If
            End If
            
         Case "3"
            If IsNull(adoquery.Fields("cu06").Value) = False Then
               If "" & adoacc1k0.Fields("a1k04").Value = "Y" Then
                  '2010/11/26 modify BY SONIA Y34271000+X18031000ªº¥ª¤W¨¤Payer¥u¦L¥Ó½Ð¤H1
                  'PutData adoquery.Fields("cu06").Value, intRow
                  'Modified by Morgan 2015/7/24 §ï»Pa1k04³]©w¤@­P§PÂ_a1k28(­ìa1k03)
                  If ii > 0 And "" & adoacc1k0("a1k13").Value = "FCP" And ("" & adoacc1k0("a1k28").Value = "Y34271000" And strCust1 = "X18031000") Then
                  Else
                     If ii > 0 Then intRow = intRow + 1 'Add by Morgan 2010/11/29
                     PutData adoquery.Fields("cu06").Value, intRow
                     'Modified by Morgan 2014/3/28
                     'SetWordArray m_Head, intRow, 2, adoquery.Fields("cu06").Value 'Add by Morgan 2010/12/1
                     SetWordArray m_Head, intRow - 7, 2, adoquery.Fields("cu06").Value
                     'end 2014/3/28
                  End If
                  '2010/11/26 END
               End If
               strCustName(ii) = strCustName(ii) & "" & adoquery.Fields("cu06").Value
               
            'Add by Morgan 2009/11/23
            ElseIf IsNull(adoquery.Fields("cu05").Value) = False Then
               If "" & adoacc1k0.Fields("a1k04").Value = "Y" Then
                  '2010/11/26 modify BY SONIA Y34271000+X18031000ªº¥ª¤W¨¤Payer¥u¦L¥Ó½Ð¤H1
                  'PutData adoquery.Fields("cu05").Value, intRow
                  'Modified by Morgan 2015/7/24 §ï»Pa1k04³]©w¤@­P§PÂ_a1k28(­ìa1k03)
                  If ii > 0 And "" & adoacc1k0("a1k13").Value = "FCP" And ("" & adoacc1k0("a1k28").Value = "Y34271000" And strCust1 = "X18031000") Then
                  Else
                     If ii > 0 Then intRow = intRow + 1 'Add by Morgan 2010/11/29
                     PutData adoquery.Fields("cu05").Value, intRow
                     'Modified by Morgan 2014/3/28
                     'SetWordArray m_Head, intRow, 2, adoquery.Fields("cu05").Value 'Add by Morgan 2010/12/1
                     SetWordArray m_Head, intRow - 7, 2, adoquery.Fields("cu05").Value
                     'end 2014/3/28
                  End If
                  '2010/11/26 END
               End If
               strCustName(ii) = strCustName(ii) & "" & adoquery.Fields("cu05").Value
            
               If IsNull(adoquery.Fields("cu88").Value) = False Then
                  If "" & adoacc1k0.Fields("a1k04").Value = "Y" Then
                     '2010/11/26 modify BY SONIA Y34271000+X18031000ªº¥ª¤W¨¤Payer¥u¦L¥Ó½Ð¤H1
                     'intRow = intRow + 1
                     'PutData adoquery.Fields("cu88").Value, intRow
                     'Modified by Morgan 2015/7/24 §ï»Pa1k04³]©w¤@­P§PÂ_a1k28(­ìa1k03)
                     If ii > 0 And "" & adoacc1k0("a1k13").Value = "FCP" And ("" & adoacc1k0("a1k28").Value = "Y34271000" And strCust1 = "X18031000") Then
                     Else
                        intRow = intRow + 1
                        PutData adoquery.Fields("cu88").Value, intRow
                        'Modified by Morgan 2014/3/28
                        'SetWordArray m_Head, intRow, 2, adoquery.Fields("cu88").Value 'Add by Morgan 2010/12/1
                        SetWordArray m_Head, intRow - 7, 2, adoquery.Fields("cu88").Value
                        'end 2014/3/28
                     End If
                     '2010/11/26 END
                  End If
                  strCustName(ii) = strCustName(ii) & " " & adoquery.Fields("cu88").Value
               End If
               If IsNull(adoquery.Fields("cu89").Value) = False Then
                  If "" & adoacc1k0.Fields("a1k04").Value = "Y" Then
                     '2010/11/26 modify BY SONIA Y34271000+X18031000ªº¥ª¤W¨¤Payer¥u¦L¥Ó½Ð¤H1
                     'intRow = intRow + 1
                     'PutData adoquery.Fields("cu89").Value, intRow
                     'Modified by Morgan 2015/7/24 §ï»Pa1k04³]©w¤@­P§PÂ_a1k28(­ìa1k03)
                     If ii > 0 And "" & adoacc1k0("a1k13").Value = "FCP" And ("" & adoacc1k0("a1k28").Value = "Y34271000" And strCust1 = "X18031000") Then
                     Else
                        intRow = intRow + 1
                        PutData adoquery.Fields("cu89").Value, intRow
                        'Modified by Morgan 2014/3/28
                        'SetWordArray m_Head, intRow, 2, adoquery.Fields("cu89").Value 'Add by Morgan 2010/12/1
                        SetWordArray m_Head, intRow - 7, 2, adoquery.Fields("cu89").Value
                        'end 2014/3/28
                     End If
                     '2010/11/26 END
                  End If
                  strCustName(ii) = strCustName(ii) & " " & adoquery.Fields("cu89").Value
               End If
               If IsNull(adoquery.Fields("cu90").Value) = False Then
                  If "" & adoacc1k0.Fields("a1k04").Value = "Y" Then
                     '2010/11/26 modify BY SONIA Y34271000+X18031000ªº¥ª¤W¨¤Payer¥u¦L¥Ó½Ð¤H1
                     'intRow = intRow + 1
                     'PutData adoquery.Fields("cu90").Value, intRow
                     'Modified by Morgan 2015/7/24 §ï»Pa1k04³]©w¤@­P§PÂ_a1k28(­ìa1k03)
                     If ii > 0 And "" & adoacc1k0("a1k13").Value = "FCP" And ("" & adoacc1k0("a1k28").Value = "Y34271000" And strCust1 = "X18031000") Then
                     Else
                        intRow = intRow + 1
                        PutData adoquery.Fields("cu90").Value, intRow
                        'Modified by Morgan 2014/3/28
                        'SetWordArray m_Head, intRow, 2, adoquery.Fields("cu90").Value 'Add by Morgan 2010/12/1
                        SetWordArray m_Head, intRow - 7, 2, adoquery.Fields("cu90").Value
                        'end 2014/3/28
                     End If
                     '2010/11/26 END
                  End If
                  strCustName(ii) = strCustName(ii) & " " & adoquery.Fields("cu90").Value
               End If
            End If
      End Select
      ii = ii + 1
      adoquery.MoveNext
   Wend
End Sub

'Added by Morgan 2015/11/19 PrintHead¤Ó¤jµLªk°õ¦æ©î³¡¤Àµ{¦¡¥X¨Ó
Private Sub PrintHeadPart0()
      strLanguage = GetLanguage("" & adoacc1k0.Fields("a1k13").Value, "" & adoacc1k0.Fields("a1k14").Value, "" & adoacc1k0.Fields("a1k15").Value, "" & adoacc1k0.Fields("a1k16").Value, "" & adoacc1k0.Fields("a1k01").Value)
      Select Case strLanguage
         Case "1" '¤¤
            If m_b2Printer Then
               Printer.Font.Name = "²Ó©úÅé"
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.Font.Name = "²Ó©úÅé"
            'End If
            'end 2022/8/4
         Case "2" '­^
            If m_b2Printer Then
               Printer.Font.Name = "Times New Roman"
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.Font.Name = "Times New Roman"
            'End If
            'end 2022/8/4
         Case "3" '¤é
            intRow = intRow + 2
            If m_b2Printer Then
               Printer.Font.Name = "²Ó©úÅé"
               Printer.Font.Size = 18
               Printer.CurrentX = (10000 - Printer.TextWidth("½Ð¡@¨D¡@®Ñ")) / 2 + intDefault
               Printer.CurrentY = 1500 + intRow * m_LineH + intTop
               'Modified by Morgan 2022/8/4
               'Printer.Print "½Ð¡@¨D¡@®Ñ"
               PUB_PrintUnicodeText "½Ð¡@¨D¡@®Ñ", Printer.CurrentX, Printer.CurrentY, 0
               'end 2022/8/4
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.Font.Name = "²Ó©úÅé"
            '   Picture1.Font.Size = 18 * douExtRate
            '   Picture1.CurrentX = ((10000 - Picture1.TextWidth("½Ð¡@¨D¡@®Ñ")) / 2 + intDefault) * douExtRate
            '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
            '   Picture1.Print "½Ð¡@¨D¡@®Ñ"
            'End If
            'end 2022/8/4
            m_Title = "½Ð¡@¨D¡@®Ñ" 'Add by Morgan 2011/8/26
            intRow = intRow + 1
            If m_b2Printer Then
               Printer.Font.Size = 12
               Printer.CurrentX = 7000 + intDefault
               Printer.CurrentY = 1500 + intRow * m_LineH + intTop
               'Modified by Morgan 2022/8/4
               'Printer.Print "No." & adoacc1k0.Fields("a1k01").Value
               PUB_PrintUnicodeText "No." & adoacc1k0.Fields("a1k01").Value, Printer.CurrentX, Printer.CurrentY, 0
               'end 2022/8/4
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.Font.Size = 12 * douExtRate
            '   Picture1.CurrentX = (7000 + intDefault) * douExtRate
            '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
            '   Picture1.Print "No." & adoacc1k0.Fields("a1k01").Value
            'End If
            'end 2022/8/4
            intRow = intRow + 4
      End Select
End Sub

'Added by Morgan 2015/8/3 PrintHead¤Ó¤jµLªk°õ¦æ©î³¡¤Àµ{¦¡¥X¨Ó
Private Sub PrintHeadPart1()

   If "" & adoacc1k0.Fields("a1k04").Value = "Y" Then
      'Modified by Morgan 2015/7/24 §ï»Pa1k04³]©w¤@­P§PÂ_a1k28¥B¤£­­©w¥Ó½Ð¤H
      'If "" & adoacc1k0("a1k13").Value = "FCP" And ("" & adoacc1k0("a1k03").Value = "Y34271000" And strCustNo = "X18031000") Then
      If "" & adoacc1k0("a1k13").Value = "FCP" And ("" & adoacc1k0("a1k28").Value = "Y34271000") Then
      'end 2015/7/24
         intRow = intRow + 1
         
         'Modified by Morgan 2014/3/28
         'SetWordArray m_Head, intRow, 1, "Advance Payer : " 'Add by Morgan 2010/11/24
         If strLanguage = "3" Then
            SetWordArray m_Head, intRow - 7, 1, "Advance Payer : "
         Else
            SetWordArray m_Head, intRow, 1, "Advance Payer : "
         End If
         'end 2014/3/28
         
         PutData "Advance Payer : ", intRow, 0
      'Add by Morgan 2004/12/16 FCT ¤é¤å½Ð´Ú³æ
      ElseIf "" & adoacc1k0("a1k13").Value = "FCT" And strLanguage = "3" Then
         intRow = intRow + 1
         'Modified by Morgan 2022/7/27
         'strExc(1) = "¡]’Iþê¡G" & PUB_GetFAgentName("" & adoacc1k0("a1k03").Value, "3") & "¡^"
         strExc(1) = "¡]" & PUB_GetUniText(Me.Name, "Âà¥æ") & "¡G" & PUB_GetFAgentName("" & adoacc1k0("a1k03").Value, "3") & "¡^"
         'end 2022/7/27
         SetWordArray m_Head, intRow - 7, 1, strExc(1) 'Added by Morgan 2014/3/28
         PutData strExc(1), intRow, , 1600
         intRow = intRow - 1
      Else
         intRow = intRow + 1
         
         'Modified by Morgan 2014/3/28
         'SetWordArray m_Head, intRow, 1, "C/O" 'Add by Morgan 2010/11/24
         If strLanguage = "3" Then
            SetWordArray m_Head, intRow - 7, 1, "C/O"
         Else
         
            'Added by Morgan 2023/1/3 MCT®×ªíÀY¤W²¾¨âÄæÁ×§K¸õ­¶--®Û­^
            If strLanguage = "1" And Left("" & adoacc1k0.Fields("a1k13").Value, 1) = "T" Then
               SetWordArray m_Head, intRow - 2, 1, "C/O"
            Else
            'end 2023/1/3
               SetWordArray m_Head, intRow, 1, "C/O"
               
            End If
         End If
         'end 2014/3/28
         
         PutData "C/O", intRow, 0
         intRow = intRow - 1
      End If
   End If
End Sub
'*************************************************
'  ©ïÀY¦C¦L
'
'*************************************************
Private Sub PrintHead()
Dim strCustName(4) As String '¥Ó½Ð¤H¦WºÙ
Dim strSystemName As String
Dim strCaseName As String
Dim strProperty As String
Dim strPatentNo As String
Dim strConNo As String
Dim strTradeMarkYes As String
Dim ii As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strTM28 As String '¨÷©v©Ê½è
Dim strTM09 As String '°Ó«~Ãþ§O
Dim strTM32 As String '°Ó«~²Õ¸s
Dim strPA08, strPA09 As String 'Added by Morgan 2013/5/14
Dim strPrintText As String
'Add by Morgan 2010/11/25
Dim strCache As String, strCache1 As String, strCache2 As String
Dim strCustNo As String
Dim iHeadRow As Integer 'Add by Morgan 2010/12/1
Dim intRRow As Integer 'Added by Morgan 2013/9/27
Dim varTmp As Variant 'Add By Sindy 2015/9/25
Dim strTitle As String 'Added by Morgan 2016/3/15
Dim strClientMatterID As String, strAccNo As String 'Added by Morgan 2018/3/22

On Error GoTo ErrHnd

   'Added by Morgan 2016/3/15 +Y54391000 ¨Ã±N³æ¾Ú¦WºÙ¥ÎÅÜ¼Æ±±¨î
   'Modified by Morgan 2016/3/17 +Y27766000,Y52963000"
   'Modified by Morgan 2017/12/19 +Y20701000 --³¯¨Ø­s
   'Modified by Morgan 2022/7/8 +Y45666000--Franny
   'Modified by Morgan 2022/9/27 +Y54093 --³¯ª÷½¬
   'Modified by Morgan 2025/7/11 +Y55033 --Anny
   If m_A1k28 = "Y52960000" Or m_A1k28 = "Y54391000" Or m_A1k28 = "Y27766000" Or m_A1k28 = "Y52963000" Or m_A1k28 = "Y20701000" Or m_A1k28 = "Y45666000" Or m_A1k28 = "Y54093000" Or m_A1k28 = "Y55033000" Then
      strTitle = "  INVOICE "
   Else
      strTitle = "DEBIT NOTE"
   End If
   'end 2016/3/15

   'Added by Morgan 2013/5/13 ±M§Q³Bµ{§Ç¦C¦L½Ð´Ú³æ®É,­n±a¥X«HÀY
   'Modify By Sindy 2015/7/13 ¶®®S¸ò¨q¬Â»¡­n¥Î±M§Qªk«ß«HÀY
   'If Pub_StrUserSt03 = "P12" And m_b2Printer And Not bolNewForm Then PrintPicture 7
   If (m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") And m_b2Printer And Not bolNewForm Then
      If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é Then
         PUB_GetLetterPicID "2", m_CP01, iPicNo, iPicNo2, 2, True, Pub_StrUserSt03
         PrintPicture iPicNo, iPicNo2
      Else
         PrintPicture 5, 9
      End If
   End If
   'Modify By Sindy 2015/11/4
   'Modified by Morgan 2015/11/6 ÁÙ­ì§_«h¦³¸ê®Æ·|¨S¦L¥X¨Ó
   intRow = 2
   'intRow = 0
   'end 2015/11/6
   '2015/11/4 END
   strCustNo = "": strCust1 = ""
   'Modified by Lydia 2015/04/09
    midStr(0) = "": midStr(1) = "": midStr(2) = "": midStr(3) = ""
    m_TM23 = "" 'Add by Amy2017/01/11
   
   '±M§Q¦³5­Ó¥Ó½Ð¤H
   '2010/11/26 modify by sonia °t¦XY34271000+X18031000ªºFCP®×¤§¥ª¤W¨¤¥u¦L¥Ó½Ð¤H1¥[§ì¥Ó½Ð¤H1¨Ñ§PÂ_
   'Modify By Sindy 2011/2/21 ¼W¥[LC43,LC44,LC45,LC46,HC24,HC25,HC26,HC27
   'Modify By Sindy 2016/6/15 1 srt ==> decode(cu01,null,'',1) srt ¦]¬°X10509287µL¥Ó½Ð¤H¸ê®Æ
   '                            cu01¤Îcu02§¡¥[(+)
        strSql = "select cu05, cu88, cu89, cu90, pa85 as Lang, PA26 As CustNo, CU06, cu04,decode(cu01,null,'',1) srt, PA26 AS CUST1 from patent, customer where substr(pa26, 1, 8) = cu01(+) and substr(pa26, 9, 1) = cu02(+) and pa01 = '" & adoacc1k0.Fields("a1k13").Value & "' and pa02 = '" & adoacc1k0.Fields("a1k14").Value & "' and pa03 = '" & adoacc1k0.Fields("a1k15").Value & "' and pa04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, pa85 as Lang, PA27 As CustNo, CU06, cu04,decode(cu01,null,'',2) srt, PA26 AS CUST1 from patent, customer where substr(pa27, 1, 8) = cu01(+) and substr(pa27, 9, 1) = cu02(+) and pa01 = '" & adoacc1k0.Fields("a1k13").Value & "' and pa02 = '" & adoacc1k0.Fields("a1k14").Value & "' and pa03 = '" & adoacc1k0.Fields("a1k15").Value & "' and pa04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, pa85 as Lang, PA28 As CustNo, CU06, cu04,decode(cu01,null,'',3) srt, PA26 AS CUST1 from patent, customer where substr(pa28, 1, 8) = cu01(+) and substr(pa28, 9, 1) = cu02(+) and pa01 = '" & adoacc1k0.Fields("a1k13").Value & "' and pa02 = '" & adoacc1k0.Fields("a1k14").Value & "' and pa03 = '" & adoacc1k0.Fields("a1k15").Value & "' and pa04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, pa85 as Lang, PA29 As CustNo, CU06, cu04,decode(cu01,null,'',4) srt, PA26 AS CUST1 from patent, customer where substr(pa29, 1, 8) = cu01(+) and substr(pa29, 9, 1) = cu02(+) and pa01 = '" & adoacc1k0.Fields("a1k13").Value & "' and pa02 = '" & adoacc1k0.Fields("a1k14").Value & "' and pa03 = '" & adoacc1k0.Fields("a1k15").Value & "' and pa04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, pa85 as Lang, PA30 As CustNo, CU06, cu04,decode(cu01,null,'',5) srt, PA26 AS CUST1 from patent, customer where substr(pa30, 1, 8) = cu01(+) and substr(pa30, 9, 1) = cu02(+) and pa01 = '" & adoacc1k0.Fields("a1k13").Value & "' and pa02 = '" & adoacc1k0.Fields("a1k14").Value & "' and pa03 = '" & adoacc1k0.Fields("a1k15").Value & "' and pa04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, tm53 as Lang, TM23 As CustNo, CU06, cu04,decode(cu01,null,'',1) srt, TM23 AS CUST1 from trademark, customer where substr(tm23, 1, 8) = cu01(+) and substr(tm23, 9, 1) = cu02(+) and tm01 = '" & adoacc1k0.Fields("a1k13").Value & "' and tm02 = '" & adoacc1k0.Fields("a1k14").Value & "' and tm03 = '" & adoacc1k0.Fields("a1k15").Value & "' and tm04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, tm53 as Lang, TM78 As CustNo, CU06, cu04,decode(cu01,null,'',2) srt, TM23 AS CUST1 from trademark, customer where substr(TM78, 1, 8) = cu01(+) and substr(TM78, 9, 1) = cu02(+) and tm01 = '" & adoacc1k0.Fields("a1k13").Value & "' and tm02 = '" & adoacc1k0.Fields("a1k14").Value & "' and tm03 = '" & adoacc1k0.Fields("a1k15").Value & "' and tm04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, tm53 as Lang, TM79 As CustNo, CU06, cu04,decode(cu01,null,'',3) srt, TM23 AS CUST1 from trademark, customer where substr(TM79, 1, 8) = cu01(+) and substr(TM79, 9, 1) = cu02(+) and tm01 = '" & adoacc1k0.Fields("a1k13").Value & "' and tm02 = '" & adoacc1k0.Fields("a1k14").Value & "' and tm03 = '" & adoacc1k0.Fields("a1k15").Value & "' and tm04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, tm53 as Lang, TM80 As CustNo, CU06, cu04,decode(cu01,null,'',4) srt, TM23 AS CUST1 from trademark, customer where substr(TM80, 1, 8) = cu01(+) and substr(TM80, 9, 1) = cu02(+) and tm01 = '" & adoacc1k0.Fields("a1k13").Value & "' and tm02 = '" & adoacc1k0.Fields("a1k14").Value & "' and tm03 = '" & adoacc1k0.Fields("a1k15").Value & "' and tm04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, tm53 as Lang, TM81 As CustNo, CU06, cu04,decode(cu01,null,'',5) srt, TM23 AS CUST1 from trademark, customer where substr(TM81, 1, 8) = cu01(+) and substr(TM81, 9, 1) = cu02(+) and tm01 = '" & adoacc1k0.Fields("a1k13").Value & "' and tm02 = '" & adoacc1k0.Fields("a1k14").Value & "' and tm03 = '" & adoacc1k0.Fields("a1k15").Value & "' and tm04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, '' as Lang, LC11 As CustNo, CU06, cu04,decode(cu01,null,'',1) srt, LC11 AS CUST1 from lawcase, customer where substr(lc11, 1, 8) = cu01(+) and substr(lc11, 9, 1) = cu02(+) and lc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and lc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and lc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and lc04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, '' as Lang, LC43 As CustNo, CU06, cu04,decode(cu01,null,'',2) srt, LC11 AS CUST1 from lawcase, customer where substr(lc43, 1, 8) = cu01(+) and substr(lc43, 9, 1) = cu02(+) and lc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and lc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and lc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and lc04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, '' as Lang, LC44 As CustNo, CU06, cu04,decode(cu01,null,'',3) srt, LC11 AS CUST1 from lawcase, customer where substr(lc44, 1, 8) = cu01(+) and substr(lc44, 9, 1) = cu02(+) and lc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and lc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and lc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and lc04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, '' as Lang, LC45 As CustNo, CU06, cu04,decode(cu01,null,'',4) srt, LC11 AS CUST1 from lawcase, customer where substr(lc45, 1, 8) = cu01(+) and substr(lc45, 9, 1) = cu02(+) and lc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and lc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and lc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and lc04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, '' as Lang, LC46 As CustNo, CU06, cu04,decode(cu01,null,'',5) srt, LC11 AS CUST1 from lawcase, customer where substr(lc46, 1, 8) = cu01(+) and substr(lc46, 9, 1) = cu02(+) and lc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and lc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and lc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and lc04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, '' as Lang, HC05 As CustNo, CU06, cu04,decode(cu01,null,'',1) srt, HC05 AS CUST1 from hirecase, customer where substr(hc05, 1, 8) = cu01(+) and substr(hc05, 9, 1) = cu02(+) and hc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and hc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and hc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and hc04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, '' as Lang, HC24 As CustNo, CU06, cu04,decode(cu01,null,'',2) srt, HC05 AS CUST1 from hirecase, customer where substr(hc24, 1, 8) = cu01(+) and substr(hc24, 9, 1) = cu02(+) and hc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and hc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and hc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and hc04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, '' as Lang, HC25 As CustNo, CU06, cu04,decode(cu01,null,'',3) srt, HC05 AS CUST1 from hirecase, customer where substr(hc25, 1, 8) = cu01(+) and substr(hc25, 9, 1) = cu02(+) and hc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and hc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and hc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and hc04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, '' as Lang, HC26 As CustNo, CU06, cu04,decode(cu01,null,'',4) srt, HC05 AS CUST1 from hirecase, customer where substr(hc26, 1, 8) = cu01(+) and substr(hc26, 9, 1) = cu02(+) and hc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and hc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and hc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and hc04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, '' as Lang, HC27 As CustNo, CU06, cu04,decode(cu01,null,'',5) srt, HC05 AS CUST1 from hirecase, customer where substr(hc27, 1, 8) = cu01(+) and substr(hc27, 9, 1) = cu02(+) and hc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and hc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and hc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and hc04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, sp34 as Lang, SP08 As CustNo, CU06, cu04,decode(cu01,null,'',1) srt, SP08 AS CUST1 from servicepractice, customer where substr(sp08, 1, 8) = cu01(+) and substr(sp08, 9, 1) = cu02(+) and sp01 = '" & adoacc1k0.Fields("a1k13").Value & "' and sp02 = '" & adoacc1k0.Fields("a1k14").Value & "' and sp03 = '" & adoacc1k0.Fields("a1k15").Value & "' and sp04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, sp34 as Lang, SP58 As CustNo, CU06, cu04,decode(cu01,null,'',2) srt, SP08 AS CUST1 from servicepractice, customer where substr(SP58, 1, 8) = cu01(+) and substr(SP58, 9, 1) = cu02(+) and sp01 = '" & adoacc1k0.Fields("a1k13").Value & "' and sp02 = '" & adoacc1k0.Fields("a1k14").Value & "' and sp03 = '" & adoacc1k0.Fields("a1k15").Value & "' and sp04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, sp34 as Lang, SP59 As CustNo, CU06, cu04,decode(cu01,null,'',3) srt, SP08 AS CUST1 from servicepractice, customer where substr(SP59, 1, 8) = cu01(+) and substr(SP59, 9, 1) = cu02(+) and sp01 = '" & adoacc1k0.Fields("a1k13").Value & "' and sp02 = '" & adoacc1k0.Fields("a1k14").Value & "' and sp03 = '" & adoacc1k0.Fields("a1k15").Value & "' and sp04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, sp34 as Lang, SP65 As CustNo, CU06, cu04,decode(cu01,null,'',4) srt, SP08 AS CUST1 from servicepractice, customer where substr(SP65, 1, 8) = cu01(+) and substr(SP65, 9, 1) = cu02(+) and sp01 = '" & adoacc1k0.Fields("a1k13").Value & "' and sp02 = '" & adoacc1k0.Fields("a1k14").Value & "' and sp03 = '" & adoacc1k0.Fields("a1k15").Value & "' and sp04 = '" & adoacc1k0.Fields("a1k16").Value & "' " & _
           "union select cu05, cu88, cu89, cu90, sp34 as Lang, SP66 As CustNo, CU06, cu04,decode(cu01,null,'',5) srt, SP08 AS CUST1 from servicepractice, customer where substr(SP66, 1, 8) = cu01(+) and substr(SP66, 9, 1) = cu02(+) and sp01 = '" & adoacc1k0.Fields("a1k13").Value & "' and sp02 = '" & adoacc1k0.Fields("a1k14").Value & "' and sp03 = '" & adoacc1k0.Fields("a1k15").Value & "' and sp04 = '" & adoacc1k0.Fields("a1k16").Value & "' order by srt"
   adoquery.CursorLocation = adUseClient
   adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   '¦³°ò¥»ÀÉ
   If adoquery.RecordCount <> 0 Then
      strCust1 = "" & adoquery("CUST1").Value    '2010/11/26 ADD BY SONIA
      strCustNo = "" & adoquery("CustNo").Value
      'Added by Morgan 2024/3/25
      m_CustNoList = strCustNo
      If adoquery.RecordCount > 1 Then
         adoquery.MoveNext
         Do While Not adoquery.EOF
            If "" & adoquery("CustNo") <> "" Then
               m_CustNoList = m_CustNoList & "," & adoquery("CustNo")
            End If
            adoquery.MoveNext
         Loop
         adoquery.MoveFirst
      End If
      'end 2024/3/25
      
      'Add by Amy 2017/01/11 +§ì¥Ó½Ð¤H (for MCTF§PÂ_)
      If Left("" & adoacc1k0.Fields("a1k13").Value, 1) = "T" Then
          m_TM23 = "" & adoquery.Fields("CUST1")
      End If
      'end 2017/01/11
      '­Y°ò¥»ÀÉ¦³©w½Z»y¤å
      If IsNull(adoquery.Fields("Lang").Value) = False Then
         'Modified by Morgan 2017/9/20 FCP57495½Ð­Ó®×³]©w±b³æ»y¨¥:­^¤å --¦ó²QµØ
         'strLanguage = adoquery.Fields("Lang").Value
         'Modified by Morgan 2018/1/5
         'If adoacc1k0.Fields("a1k13").Value = "FCP" And adoacc1k0.Fields("a1k14").Value = "057495" And adoacc1k0.Fields("a1k15").Value = "0" And adoacc1k0.Fields("a1k16").Value = "00" Then
         '   strLanguage = "2"
         'Else
         If Not GetBillLanguage(strLanguage) Then
         'end 2018/1/5
            strLanguage = adoquery.Fields("Lang").Value
         End If
         'end 2017/9/20
      '­Y°ò¥»ÀÉµL©w½Z»y¤å
      Else
          strLanguage = GetLanguage("" & adoacc1k0.Fields("a1k13").Value, "" & adoacc1k0.Fields("a1k14").Value, "" & adoacc1k0.Fields("a1k15").Value, "" & adoacc1k0.Fields("a1k16").Value, "" & adoacc1k0.Fields("a1k01").Value)
      End If
      'Added by Lydia 2015/04/09 ¤¤¤åª©-¾ã§å½Ð´Ú³æ
      If m_bolChiDB = True Then
         '¤£²Å±ø¥ó,¤£¦C¦L
         If strLanguage <> "1" Or (m_ChiCust <> "" And m_ChiCust <> "" & adoquery("CustNo").Value) Then
            GoTo Chi_JumpPHM
         Else
            bol_ChiDB = True
         End If
      End If
      Select Case strLanguage
         Case "1" '¤¤
            If m_b2Printer Then
               Printer.Font.Name = "²Ó©úÅé"
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.Font.Name = "²Ó©úÅé"
            'End If
            'end 2022/8/4
            
         Case "2" '­^
            If m_b2Printer Then
               Printer.Font.Name = "Times New Roman"
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.Font.Name = "Times New Roman"
            'End If
            'end 2022/8/4
         Case "3" '¤é
            intRow = intRow + 2
            m_Title = "½Ð¡@¨D¡@®Ñ" 'Add by Morgan 2010/12/1
            If m_b2Printer Then
               Printer.Font.Name = "²Ó©úÅé"
               Printer.Font.Size = 18
               Printer.CurrentX = (10000 - Printer.TextWidth(m_Title)) / 2 + intDefault
               Printer.CurrentY = 1500 + intRow * m_LineH + intTop
               'Modified by Morgan 2022/8/4
               'Printer.Print m_Title
               PUB_PrintUnicodeText m_Title, Printer.CurrentX, Printer.CurrentY, 0
               'end 2022/8/4
            End If
            'Add by Morgan 2008/4/7 ²£¥Í¹q¤lÀÉ¦P®É¦L¤@¥÷
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.Font.Name = "²Ó©úÅé"
            '   Picture1.Font.Size = 18 * douExtRate
            '   Picture1.CurrentX = ((10000 - Picture1.TextWidth(m_Title)) / 2 + intDefault) * douExtRate
            '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
            '   Picture1.Print m_Title
            'End If
            'end 2022/8/4
            intRow = intRow + 1
            
            'Modified by Morgan 2022/8/4
            'If m_b2Printer Then
            '   Printer.Font.Size = 12
            '   Printer.CurrentX = 7000 + intDefault
            '   Printer.CurrentY = 1500 + intRow * m_LineH + intTop
            '   Printer.Print "No." & adoacc1k0.Fields("a1k01").Value
            'End If
            'If m_b2Picture Then
            '   Picture1.Font.Size = 12 * douExtRate
            '   Picture1.CurrentX = (7000 + intDefault) * douExtRate
            '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
            '   Picture1.Print "No." & adoacc1k0.Fields("a1k01").Value
            'End If
            PutData "No." & adoacc1k0.Fields("a1k01").Value, intRow, 7000, , 12
            'end 2022/8/4
            
            intRow = intRow + 4
      End Select
      
      'Added by Morgan 2020/1/14 ¥N²z¤H Y55364 KARL MAYER R&D GmbH ¥ª¤W¦L¯S©w EMail--¼B¿³ªN
      If m_A1k03 = "Y55364000" Then
         'Modified by Morgan 2022/11/16
         'm_tmp = "winfried.herr@karlmayer.com"
         m_tmp = "ying.tang@karlmayer.com"
         PutData m_tmp, 1, 0
         SetWordArray m_Head, 1, 1, m_tmp
         
      'Added by Morgan 2020/3/18 --Ryan
      ElseIf m_A1k03 = "Y53861010" And strCust1 = "X45899000" Then
         m_tmp = "Attention to Catherine Chenal N'Kaoua"
         PutData m_tmp, 1, 0
         SetWordArray m_Head, 1, 1, m_tmp
      'end 2020/3/18
      End If
      'end 2020/1/14
      
      If "" & adoacc1k0.Fields("a1k04").Value = "Y" Then
         'Modified by Morgan 2015/7/24 §ï»Pa1k04³]©w¤@­P§PÂ_a1k28¥B¤£­­©w¥Ó½Ð¤H
         'If "" & adoacc1k0("a1k13").Value = "FCP" And ("" & adoacc1k0("a1k03").Value = "Y34271000" And strCustNo = "X18031000") Then
         If "" & adoacc1k0("a1k13").Value = "FCP" And ("" & adoacc1k0("a1k28").Value = "Y34271000") Then
         'end 2015/7/24
            'Modified by Morgan 2014/3/28
            'SetWordArray m_Head, intRow, 1, "Payer : " 'Add by Morgan 2010/11/24
            If strLanguage = "3" Then
               SetWordArray m_Head, intRow - 7, 1, "Payer : "
            Else
               SetWordArray m_Head, intRow, 1, "Payer : "
            End If
            'end 2014/3/28
            
            PutData "Payer : ", intRow, 0
            intRow = intRow + 1
         End If
      End If
      strCustName(0) = "": strCustName(1) = "": strCustName(2) = "": strCustName(3) = "": strCustName(4) = ""
      PrintCustData strCustName, strLanguage
   '¨S°ò¥»ÀÉ
   Else
      PrintHeadPart0
      strCustName(0) = "": strCustName(1) = "": strCustName(2) = "": strCustName(3) = "": strCustName(4) = ""
   End If
   adoquery.Close
   
   PrintHeadPart1 'Added by Morgan 2015/8/3
   Select Case strLanguage
      Case "1" '¤¤¤å
         
         If Left("" & adoacc1k0.Fields("a1k13").Value, 1) <> "T" Then 'Added by Morgan 2022/12/30 MCT®×ªíÀY¤W²¾¨âÄæÁ×§K¸õ­¶--®Û­^
            intRow = intRow + 2
         End If
         '¥_¨Ê¤¤¥_(Y31671)¦L­^¤å¦WºÙ
         '2012/5/17 MODIFY BY SONIA ¥[Å±¦¨«ß®v¨Æ°È©Ò(Y52618)
         '2012/12/24 MODIFY BY SONIA µL­^¤å§ï§ì¤¤¤å
         'Modified by Morgan 2014/4/14 +½Ð´Ú¹ï¶H¬° X71814
         'Modified by Morgan 2014/4/15 +½Ð´Ú¹ï¶H¬° X71837
         'Added by Lydia 2015/04/08  §PÂ_¾ã§å½Ð´Ú³æ¥u°O¿ý²Ä¤@µ§
         'Modified by Morgan 2017/8/14 +Y54592 --®Û­^
         If m_bolChiDB = False Or (m_bolChiDB And adoacc1k0.AbsolutePosition = 1) Then
            If (Left(m_A1k03, 6) = "Y31671" Or Left(m_A1k03, 6) = "Y52618" Or Left(m_A1k28, 6) = "X71814" Or Left(m_A1k28, 6) = "X71837" Or Left(m_A1k28, 6) = "Y54592") And Not IsNull("" & adoacc1k0.Fields("fa05").Value) Then
               '¥N²z¤H­^¤å¦WºÙ
               PutData "" & adoacc1k0.Fields("fa05").Value, intRow
               PutData "" & adoacc1k0.Fields("fa63").Value, intRow + 1
               PutData "" & adoacc1k0.Fields("fa64").Value, intRow + 2
               PutData "" & adoacc1k0.Fields("fa65").Value, intRow + 3
               'Add by Morgan 2010/12/1
               SetWordArray m_Head, intRow, 2, "" & adoacc1k0.Fields("fa05").Value
               SetWordArray m_Head, intRow + 1, 2, "" & adoacc1k0.Fields("fa63").Value
               SetWordArray m_Head, intRow + 2, 2, "" & adoacc1k0.Fields("fa64").Value
               SetWordArray m_Head, intRow + 3, 2, "" & adoacc1k0.Fields("fa65").Value
            Else
               ''¥N²z¤H¤¤¤å¦WºÙ
               If IsNull(adoacc1k0.Fields("fa04").Value) = False Then
                  PutData "" & adoacc1k0.Fields("fa04").Value, intRow
                  SetWordArray m_Head, intRow, 2, "" & adoacc1k0.Fields("fa04").Value 'Add by Morgan 2010/12/1
                  'Added by Morgan 2015/11/19
                  '¥_¨Ê»ÈÀs¥[¦L¤¤¤å¦a§}
                  'modify by sonia 2019/6/13 +Y5245900
                  'modify by sonia 2019/7/10 §ï¥þ³¡³£¦L
                  'If m_A1k28 = "Y51333010" Or m_A1k28 = "Y52459000" Then
                  If bolNewForm Then 'Added by Morgan 2022/11/24 ÂÂ®æ¦¡±b³æ¤£¦L¡A§_«h·|»\¨ì Ex:X09906508
                     If Not IsNull(adoacc1k0.Fields("fa17").Value) Then
                        PutData "" & adoacc1k0.Fields("fa17").Value, intRow
                        SetWordArray m_Head, intRow + 1, 2, adoacc1k0.Fields("fa17")
                     End If
                  End If
                  'End If
                  'End 2015/11/19
               'Add By Sindy 2016/8/24 ­Y¨S¦³¤¤¤å§ï§ì­^¤å,µL­^¤å¤~¦A§ì¤é¤å ex:Y54478000
               ElseIf IsNull(adoacc1k0.Fields("fa05").Value) = False Then
                  '¥N²z¤H­^¤å¦WºÙ
                  PutData "" & adoacc1k0.Fields("fa05").Value, intRow
                  PutData "" & adoacc1k0.Fields("fa63").Value, intRow + 1
                  PutData "" & adoacc1k0.Fields("fa64").Value, intRow + 2
                  PutData "" & adoacc1k0.Fields("fa65").Value, intRow + 3
                  'Add by Morgan 2010/12/1
                  SetWordArray m_Head, intRow, 2, "" & adoacc1k0.Fields("fa05").Value
                  
                  'Modified by Morgan 2024/2/20 µ{¦¡¤Ó¤j§ï¨ç¼Æ¨Ã¼W¥[¯S©w½Ð´Ú¹ï¶H­n¦L­^¤å¦a§}
                  'SetWordArray m_Head, intRow + 1, 2, "" & adoacc1k0.Fields("fa63").Value
                  'SetWordArray m_Head, intRow + 2, 2, "" & adoacc1k0.Fields("fa64").Value
                  'SetWordArray m_Head, intRow + 3, 2, "" & adoacc1k0.Fields("fa65").Value
                  SetEngName intRow
                  'end 2024/2/20
               '¥N²z¤H¤é¤å¦WºÙ
               ElseIf IsNull(adoacc1k0.Fields("fa06").Value) = False Then
                  'Add by Morgan 2005/12/6
                  ''Modified by Morgan 2015/7/24 §ï»Pa1k04³]©w¤@­P§PÂ_a1k28¥B¤£­­©w¥Ó½Ð¤H
                  If "" & adoacc1k0("a1k13").Value = "FCP" And ("" & adoacc1k0("a1k28").Value = "Y34271000") Then
                     intRow = intRow + 1
                  End If
                  m_tmp = adoacc1k0.Fields("fa06").Value & "¡@" & "±s¤¤"
                  PutData m_tmp, intRow
                  SetWordArray m_Head, intRow, 2, m_tmp
               '2016/8/24 END
               End If
            End If
            
            If "" & adoacc1k0.Fields("a1k02").Value <> "" Then
               m_tmp = Format(DBDATE(adoacc1k0.Fields("a1k02").Value), "#### ¦~ ## ¤ë ## ¤é")
               PutData m_tmp, intRow, 5500
               SetWordArray m_Head, intRow, 3, m_tmp 'Add by Morgan 2010/12/1
            End If
         End If
         
         '©¼©Ò®×¸¹
         intRow = intRow + 1
         adoquery.CursorLocation = adUseClient
         'Modify by Morgan 2006/12/15 TS¬d¦W®×Ãþ§O§ì²Õ¸s
         '2007/7/24 MODIFY BY SONIA TS¬d¦W®×µL²Õ¸s®É¦LÃþ§O
         '2009/9/21 MODIFY BY SONIA ¥[¤À©Ò®×¸¹Dno(¥¨¨Ê®×¸¹)
         'Modify By Sindy 2015/7/8 tm05 ==> nvl(tm131,tm05)
         'modify by sonia 2016/12/23 +tm08
         adoquery.Open "select pa77 as Yno, pa48 as Cno, decode(pa09,'020',ptm04,ptm03) as MName, pa05 as Cname, pa11 as Ano, pa26 as Custno, pa22, null as tm15, null as Yes, '' As TM12, '' As TM16, '' as TM09, '' AS TM32, pa47 as Dno,pa08 from patent, systemkind, patenttrademarkmap where pa01 = sk01 and pa08 = ptm02 (+) and sk02 = ptm01 and pa01 = '" & adoacc1k0.Fields("a1k13").Value & "' and pa02 = '" & adoacc1k0.Fields("a1k14").Value & "' and pa03 = '" & adoacc1k0.Fields("a1k15").Value & "' and pa04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                       "select tm45 as Yno, tm35 as Cno, decode(tm10,'020',ptm04,ptm03) as MName, Rtrim(Ltrim(nvl(tm131,tm05)||' '||tm06)) as Cname, tm12 as Ano, tm23 as Custno, null as pa22, tm15, '1' as Yes, TM12, TM16, TM09, TM32, tm34 as Dno,tm08 as pa08 from trademark, systemkind, patenttrademarkmap where tm01 = sk01 and tm08 = ptm02 (+) and sk02 = ptm01 and tm01 = '" & adoacc1k0.Fields("a1k13").Value & "' and tm02 = '" & adoacc1k0.Fields("a1k14").Value & "' and tm03 = '" & adoacc1k0.Fields("a1k15").Value & "' and tm04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                       "select lc23 as Yno, lc17 as Cno, '' as MName, nvl(lc05,lc06) as Cname, '' as Ano, lc11 as Custno, null as pa22, null as tm15, null as Yes, '' As TM12, '' As TM16, '' As TM09, '' AS TM32, lc16 as Dno,'' as pa08 from lawcase where lc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and lc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and lc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and lc04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                       "select sp27 as Yno, sp29 as Cno, '' as MName, nvl(sp05,sp06) as Cname, sp11 as Ano, sp08 as Custno, null as pa22, null as tm15, null as Yes, '' As TM12, '' As TM16, sp73 As TM09, SP74 AS TM32, sp28 as Dno,'' as pa08 from servicepractice where sp01 = '" & adoacc1k0.Fields("a1k13").Value & "' and sp02 = '" & adoacc1k0.Fields("a1k14").Value & "' and sp03 = '" & adoacc1k0.Fields("a1k15").Value & "' and sp04 = '" & adoacc1k0.Fields("a1k16").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            m_CaseNo = "" & adoquery.Fields("Cno").Value  'Added by Morgan 2023/10/13
            strPA08 = "" & adoquery.Fields("pa08").Value  'add by sonia 2016/12/23
            m_tmp = "¶Q¤è¨÷¸¹: "
            m_tmp = m_tmp & GetYourRef(adoacc1k0.Fields("A1K13").Value, adoacc1k0.Fields("A1K01").Value, "" & adoquery.Fields("Yno").Value)
            'end 2016/1/12
            'Modified by Lydia 2015/04/09 §PÂ_«D¾ã§å½Ð´Ú³æ,¦C¦L©ïÀY¸ê®Æ
            If m_bolChiDB = False Then
                PutData m_tmp, intRow, 5500
                SetWordArray m_Head, intRow, 3, m_tmp 'Add by Morgan 2010/12/1
            End If
            
            intRow = intRow + 1
            m_tmp = "¥»©Ò®×¸¹: " & adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value
            If "" & adoacc1k0.Fields("a1k15") & adoacc1k0.Fields("a1k16").Value <> "000" Then
               m_tmp = m_tmp & "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value
            End If
            'Modified by Lydia 2015/04/09 +§PÂ_«D¾ã§å½Ð´Ú³æ,¦C¦L©ïÀY¸ê®Æ
            If m_bolChiDB = True Then
               '¼È¦s©ú²Ó
                midStr(0) = adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & IIf(adoacc1k0.Fields("a1k15") & adoacc1k0.Fields("a1k16").Value <> "000", "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value, "")
            Else
                PutData m_tmp, intRow, 5500
                SetWordArray m_Head, intRow, 3, m_tmp 'Add by Morgan 2010/12/1
            End If
            
            'Modified by Lydia 2015/04/09 §PÂ_«D¾ã§å½Ð´Ú³æ,¦C¦L©ïÀY¸ê®Æ
            If m_bolChiDB = False Then
                If IsNull(adoquery.Fields("Cno").Value) = False Then
                   intRow = intRow + 1
                   m_tmp = "«È¤á®×¸¹: " & adoquery.Fields("Cno").Value
                   PutData m_tmp, intRow, 5500
                   SetWordArray m_Head, intRow, 3, m_tmp 'Add by Morgan 2010/12/1
                End If
                
                '2009/9/21 ADD BY SONIA ¥[¤À©Ò®×¸¹Dno(¥¨¨Ê®×¸¹)
                If IsNull(adoquery.Fields("Dno").Value) = False Then
                   intRow = intRow + 1
                   m_tmp = "¥¨¨Ê®×¸¹: " & adoquery.Fields("Dno").Value
                   PutData m_tmp, intRow, 5500
                   SetWordArray m_Head, intRow, 3, m_tmp 'Add by Morgan 2010/12/1
                End If
                '2009/9/21 END
            End If
            '­Y¬°®Ö­ã¥B¦³¼f©w¸¹®É, ¦L¼f©w¸¹, §_«h¦L¥Ó½Ð®×¸¹
            If "" & adoquery.Fields("TM16").Value = "1" And "" & adoquery.Fields("TM15").Value <> "" Then
               strConNo = "" & adoquery.Fields("TM15").Value
            Else
               strConNo = ""
            End If
            strPatentNo = "" & adoquery.Fields("pa22").Value
            strTradeMarkYes = "" & adoquery.Fields("Yes").Value
            strSystemName = "" & adoquery.Fields("MName").Value
            strCaseName = "" & adoquery.Fields("Cname").Value
            m_CaseName = strCaseName 'Added by Morgan 2019/5/29
            strAppNo = "" & adoquery.Fields("Ano").Value
            strCustNo = "" & adoquery.Fields("Custno").Value
            strTM09 = "" & adoquery.Fields("TM09").Value
            strTM32 = "" & adoquery.Fields("TM32").Value
            
            'Modified by Lydia 2015/04/09 ¼È¦s©ú²Ó
            midStr(1) = strCaseName
            midStr(2) = strTM09
            midStr(3) = IIf(strConNo <> "", strConNo, strAppNo)
         End If
         adoquery.Close
'end Added by Lydia 2015/04/09 «D-¤¤¤åª©¾ã§å½Ð´Ú³æ,¦C¦L©ïÀY¸ê®Æ
''''''''''
         '¥N²z¤H/¥Ó½Ð¤H°]°È½s¸¹
         'Modify by Morgan 2010/8/20 ­n§ì½Ð´Ú¹ï¶Hªº--David
         'Modified by Morgan 2012/4/17 §ï§ì¦C¦L¹ï¶H Ex.X10102561 --David,³¯ª÷½¬
         'strExc(1) = m_A1k28
         strExc(1) = m_A1k27
         'end 2012/4/17
         If Left(strExc(1), 1) = "Y" Then
            'Add By Sindy 2011/3/7 +FA106
            StrSQLa = "select FA28,FA106 from fagent where fa01='" & Left(strExc(1), 8) & "' and fa02='" & Mid(strExc(1), 9) & "'"
         Else
            'Add By Sindy 2011/3/7 +CU146
            StrSQLa = "select CU33,CU146 from customer where cu01='" & Left(strExc(1), 8) & "' and cu02='" & Mid(strExc(1), 9) & "'"
         End If
         'end 2010/8/20
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            'Modify by Morgan 2010/8/20 ­n§ì½Ð´Ú¹ï¶Hªº--David
            'If "" & rsA("FA28").Value <> "" Then
            '   m_tmp = "" & rsA("FA28").Value
            'ElseIf "" & rsA("CU33").Value <> "" Then
            '   m_tmp = "" & rsA("CU33").Value
            'Else
            '   m_tmp = ""
            'End If
            'Modify By Sindy 2011/3/7
            If CheckSys("" & adoacc1k0.Fields("A1K13").Value) = "2" Or _
               CheckSys("" & adoacc1k0.Fields("A1K13").Value) = "6" Then
               m_tmp = "" & rsA(1)
            '2011/3/7 End
            Else
               m_tmp = "" & rsA(0)
            End If
            'end 2010/8/20
            
            intRow = intRow + 1 'Added by Morgan 2022/5/5
            PutData m_tmp, intRow, 5500
            SetWordArray m_Head, intRow, 3, m_tmp 'Add by Morgan 2010/12/1
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         
         Call PrintHead2 'Add By Sindy 2009/07/20
            
         intRow = intRow + 1
         If m_b2Printer Then
            Printer.Line (0 + intDefault, 1500 + intRow * m_LineH + 350 + intTop)-(10000 + intDefault, 1500 + intRow * m_LineH + 350 + intTop)
         End If
         'Removed by Morgan 2022/8/4
         'If m_b2Picture Then
         '   Picture1.Line (douExtRate * (0 + intDefault), douExtRate * (1500 + intRow * m_LineH + 350 + intTop))-(douExtRate * (10000 + intDefault), douExtRate * (1500 + intRow * m_LineH + 350 + intTop))
         'End If
         'end 2022/8/4
         intRow = intRow + 1
   
      Case "2" '­^¤å
         '¥N²z¤H­^¤å¦WºÙ
         If IsNull(adoacc1k0.Fields("fa05").Value) = False Then
            intRow = intRow + 1
            m_tmp = adoacc1k0.Fields("fa05").Value
            'Added by Morgan 2020/10/7 ¤é¥N <Y52259> NAGAI & ASSOCIATES­n¨D¥»©Ò´£¨Ñ¤§±b³æ°£¨ä¨Æ°È©Ò¦W¥~,¶·¦A¥[µù¨ä©Òªø©m¦W¡uFuyuki NAGAI¡v--³¢©É¼ü
            If m_A1k27 = "Y52259000" Then m_tmp = RTrim(m_tmp) & " Fuyuki NAGAI"
            'end 2020/10/7
            'Added by Morgan 2023/2/20 --²øÞ±¤Z
            If m_A1k27 = "Y55261000" And adoacc1k0.Fields("a1k13") = "FCP" Then
               m_tmp = "AVV. DAVIDE BRESNER" & vbCrLf & m_tmp
            End If
            'end 2023/2/20
            PutData m_tmp, intRow
            SetWordArray m_Head, intRow, 2, m_tmp 'Add by Morgan 2010/11/24
         End If
         If IsNull(adoacc1k0.Fields("fa63").Value) = False Then
            intRow = intRow + 1
            m_tmp = adoacc1k0.Fields("fa63").Value
            PutData m_tmp, intRow
            SetWordArray m_Head, intRow, 2, m_tmp 'Add by Morgan 2010/11/24
         End If
         If IsNull(adoacc1k0.Fields("fa64").Value) = False Then
            intRow = intRow + 1
            m_tmp = adoacc1k0.Fields("fa64").Value
            PutData m_tmp, intRow
            SetWordArray m_Head, intRow, 2, m_tmp 'Add by Morgan 2010/11/24
         End If
         If IsNull(adoacc1k0.Fields("fa65").Value) = False Then
            intRow = intRow + 1
            m_tmp = adoacc1k0.Fields("fa65").Value
            PutData m_tmp, intRow
            SetWordArray m_Head, intRow, 2, m_tmp 'Add by Morgan 2010/11/24
         End If
         
         intRow = intRow + 1
         '­^¤å¦a§}
         If IsNull(adoacc1k0.Fields("fa32").Value) = False Then
            m_tmp = adoacc1k0.Fields("fa32").Value
            PutData m_tmp, intRow
            SetWordArray m_Head, intRow, 2, m_tmp 'Add by Morgan 2010/11/24
            
         ElseIf IsNull(adoacc1k0.Fields("fa18").Value) = False Then
            m_tmp = adoacc1k0.Fields("fa18").Value
            PutData m_tmp, intRow
            SetWordArray m_Head, intRow, 2, m_tmp 'Add by Morgan 2010/11/24
         End If
         
         intRRow = 2
         'Added by Morgan 2022/8/5
         'Y55751 Birkenstock IP GmbH--Franny,Peggy
         If m_A1k28 = "Y55751000" Then
            m_tmp = "Invoice Date:"
         Else
         'end 2022/8/5
            m_tmp = "Date:"
         End If
         
         PutData m_tmp, intRRow, 5500
         SetWordArray m_Head, intRRow, 3, m_tmp   'Add by Morgan 2010/11/24
         If "" & adoacc1k0.Fields("a1k02").Value <> "" Then
            'Modified by Morgan 2018/4/11
            If m_bSpecialNew4 Then
               m_tmp = GetEngDate(adoacc1k0.Fields("a1k02").Value, 1)
            Else
               m_tmp = GetEngDate(adoacc1k0.Fields("a1k02").Value) 'Modified by Morgan 2017/7/7 §ï¥Î¨ç¼Æ
            End If
            'end 2018/4/11
            PutData m_tmp, intRRow, 6500
            SetWordArray m_Head, intRRow, 4, m_tmp 'Add by Morgan 2010/11/24
         End If
         
         'Added by Morgan 2022/8/5
         'Y55751 Birkenstock IP GmbH--Franny
         If m_A1k28 = "Y55751000" And (m_CP01 = "FCP" Or m_CP01 = "P") Then
            intRRow = intRRow + 1
            m_tmp = "Service Date:"
            SetWordArray m_Head, intRRow, 3, m_tmp
            
            m_tmp = GetEngDate(GetCp27(adoacc1k0.Fields("a1k01"), , True))
            SetWordArray m_Head, intRRow, 4, m_tmp
         End If
         'end 2022/8/5
         
         If IsNull(adoacc1k0.Fields("fa32").Value) Then
            intRow = intRow + 1
            If IsNull(adoacc1k0.Fields("fa19").Value) = False Then
               m_tmp = adoacc1k0.Fields("fa19")
               PutData m_tmp, intRow
               SetWordArray m_Head, intRow, 2, m_tmp 'Add by Morgan 2010/11/24
            End If
         Else
            If IsNull(adoacc1k0.Fields("fa33").Value) = False Then
               intRow = intRow + 1
               m_tmp = adoacc1k0.Fields("fa33")
               PutData m_tmp, intRow
               SetWordArray m_Head, intRow, 2, m_tmp 'Add by Morgan 2010/11/24
            End If
         End If
         
         intRRow = intRRow + 1
         adoquery.CursorLocation = adUseClient
         '2007/7/24 MODIFY BY SONIA TS¬d¦W®×µL²Õ¸s®É¦LÃþ§O
         'Modified by Morgan 2012/11/30 +Ápµ¸¤H1(­^) TM39 (Syngenta ªº Reuester)
         'Modify By Sindy 2015/7/8 tm05 ==> nvl(tm131,tm05)
         'modify by sonia 2016/12/23 +tm08
         'Modified by Morgan 2017/7/7 +pa159
         adoquery.Open "select pa77 as Yno, pa48 as Cno, ptm05 as MName, pa06 as Cname, pa11 as Ano, pa26 as Custno, pa22, null as tm15, null as Yes, '' As TM12, '' As TM16, '' As TM09, '' AS TM32, PA52 As TM39,pa08,pa159,pa55 from patent, systemkind, patenttrademarkmap where pa01 = sk01 and pa08 = ptm02 (+) and sk02 = ptm01 and pa01 = '" & adoacc1k0.Fields("a1k13").Value & "' and pa02 = '" & adoacc1k0.Fields("a1k14").Value & "' and pa03 = '" & adoacc1k0.Fields("a1k15").Value & "' and pa04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                       "select tm45 as Yno, tm35 as Cno, ptm05 as MName, Rtrim(Ltrim(nvl(tm131,tm05)||' '||tm06)) as Cname, tm12 as Ano, tm23 as Custno, null as pa22, tm15, '1' as Yes, TM12, TM16, TM09, TM32, TM39,tm08 as pa08,tm127 as pa159,tm42 as pa55 from trademark, systemkind, patenttrademarkmap where tm01 = sk01 and tm08 = ptm02 (+) and sk02 = ptm01 and tm01 = '" & adoacc1k0.Fields("a1k13").Value & "' and tm02 = '" & adoacc1k0.Fields("a1k14").Value & "' and tm03 = '" & adoacc1k0.Fields("a1k15").Value & "' and tm04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                       "select lc23 as Yno, lc17 as Cno, '' as MName, nvl(lc05,lc06) as Cname, '' as Ano, lc11 as Custno, null as pa22, null as tm15, null as Yes, '' As TM12, '' As TM16, '' As TM09, '' AS TM32, LC19 As TM39,'' as pa08,'' as pa159,'' as pa55 from lawcase where lc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and lc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and lc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and lc04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                       "select sp27 as Yno, sp29 as Cno, '' as MName, nvl(sp05,sp06) as Cname, sp11 as Ano, sp08 as Custno, null as pa22, null as tm15, null as Yes, '' As TM12, '' As TM16, sp73 As TM09, SP74 AS TM32, SP30 As TM39,'' as pa08,sp84 as pa159,sp75 as pa55 from servicepractice where sp01 = '" & adoacc1k0.Fields("a1k13").Value & "' and sp02 = '" & adoacc1k0.Fields("a1k14").Value & "' and sp03 = '" & adoacc1k0.Fields("a1k15").Value & "' and sp04 = '" & adoacc1k0.Fields("a1k16").Value & "'", adoTaie, adOpenStatic, adLockReadOnly

         If adoquery.RecordCount <> 0 Then
            m_CaseNo = "" & adoquery.Fields("Cno").Value  'Added by Morgan 2023/10/13
            m_Att1 = "" & adoquery.Fields("TM39").Value 'Added by Morgan 2020/3/17
            m_Att2 = "" & adoquery.Fields("pa55").Value 'Added by Morgan 2020/3/17
            
            strClientMatterID = "" & adoquery.Fields("pa159").Value  'Added by Morgan 2018/3/22
            strPA08 = "" & adoquery.Fields("pa08").Value  'add by sonia 2016/12/23
            
            'Added by Morgan 2025/10/3
            'Y56199 Coupang Corp¤£¦L©¼¸¹§ï¦L«È¤á®×¥ó®×¸¹(IDF No.)¤ÎClient Matter ID(Coupang Reference No.)
            If m_A1k28 = "Y56199000" Then
               m_tmp = "IDF No.: "
               SetWordArray m_Head, intRRow, 3, m_tmp
               If m_CaseNo <> "" Then
                  SetWordArray m_Head, intRRow, 4, m_CaseNo
               End If
               intRRow = intRRow + 1
               m_tmp = "Coupang Reference No.: " & strClientMatterID
               SetWordArray m_Head, intRRow, 3, m_tmp
            Else
            'end 2025/10/3
            
               '¼ÐÃD Your Ref »P¸ê®Æ¤£­n¤À¶}¦L
                '­Y¬°FCP®×
                m_tmp = "Your Ref: "
                SetWordArray m_Head, intRRow, 3, m_tmp 'Add by Morgan 2010/11/24
                
                'Added by Morgan 2025/6/13
                '½Ð´Ú¹ï¶HY55666000 NOVOCURE GMBH ©¼¸¹Äæ¦ì (Your ref:) Àu¥ý§ì«È¤á®×¥ó®×¸¹--Franny
                'Modified by Morgan 2025/6/25 §ó¥N«á«È¤á®×¸¹·|§ï©ñ¨ì©¼¸¹ Ex:X11408163--Franny
                If m_A1k28 = "Y55666000" Then
                    If m_CaseNo <> "" Then
                      strExc(1) = m_CaseNo
                   Else
                      strExc(1) = "" & adoquery.Fields("Yno").Value
                   End If
                Else
                'end 2025/6/12
                   strExc(1) = GetYourRef(adoacc1k0.Fields("A1K13").Value, adoacc1k0.Fields("A1K01").Value, "" & adoquery.Fields("Yno").Value)
                End If
                
                If strExc(1) <> "" Then
                   SetWordArray m_Head, intRRow, 4, strExc(1)
                   m_tmp = m_tmp & strExc(1)
                End If
                'end 2016/1/12
                
                PutData m_tmp, intRRow, 5500
            End If
            
            If IsNull(adoacc1k0.Fields("fa32").Value) Then
               intRow = intRow + 1
               If IsNull(adoacc1k0.Fields("fa20").Value) = False Then
                  m_tmp = adoacc1k0.Fields("fa20").Value
                  PutData m_tmp, intRow
                  SetWordArray m_Head, intRow, 2, m_tmp 'Add by Morgan 2010/11/24
               End If
               
            ElseIf IsNull(adoacc1k0.Fields("fa34").Value) = False Then
               intRow = intRow + 1
               m_tmp = adoacc1k0.Fields("fa34").Value
               PutData m_tmp, intRow
               SetWordArray m_Head, intRow, 2, m_tmp 'Add by Morgan 2010/11/24
            End If
                        
            intRRow = intRRow + 1
            m_tmp = "Our Ref:"
            PutData m_tmp, intRRow, 5500
            SetWordArray m_Head, intRRow, 3, m_tmp 'Add by Morgan 2010/11/24
            '­Y¥»©Ò®×¸¹«á¤T½X¬°000«h¤£¦L¦¹¤T½X
            m_tmp = adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value
            If "" & adoacc1k0.Fields("a1k15") & adoacc1k0.Fields("a1k16").Value <> "000" Then
                m_tmp = m_tmp & "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value
            End If
            PutData m_tmp, intRRow, 6500
            SetWordArray m_Head, intRRow, 4, m_tmp 'Add by Morgan 2010/11/24
            
            If IsNull(adoacc1k0.Fields("fa32").Value) Then
               If IsNull(adoacc1k0.Fields("fa21").Value) = False Then
                  intRow = intRow + 1
                  m_tmp = adoacc1k0.Fields("fa21").Value
                  PutData m_tmp, intRow
                  SetWordArray m_Head, intRow, 2, m_tmp 'Add by Morgan 2010/11/24
               End If
               
            ElseIf IsNull(adoacc1k0.Fields("fa35").Value) = False Then
               intRow = intRow + 1
               m_tmp = adoacc1k0.Fields("fa35").Value
               PutData m_tmp, intRow
               SetWordArray m_Head, intRow, 2, m_tmp 'Add by Morgan 2010/11/24
            End If
            
            'Added by Morgan 2020/1/14 ¥N²z¤HY54600 Giorgio Armani SpA ¯S®í½Ð´Ú³æ--³¯ª÷½¬
            If m_A1k03 = "Y54600000" Then
               intRow = intRow + 1
               m_tmp = "VAT number : IT10985020964"
               PutData m_tmp, intRow
               SetWordArray m_Head, intRow, 2, m_tmp
               intRow = intRow + 1
               m_tmp = "Fiscal code: 02342990153"
               PutData m_tmp, intRow
               SetWordArray m_Head, intRow, 2, m_tmp
            End If
            'end 2020/1/14
            
            'Removed by Morgan 2022/8/5
            'intRRow = 4 'Added by Morgan 2013/9/27
            'end 2022/8/5
            
            'Added by Morgan 2017/3/22
            If m_bSpecial1 Then
               'Modified by Morgan 2017/8/17 +m_bDowX±±¨î
               If m_bDowX Then
                  m_tmp = "Timekeeper name: Jacky Wang"
               'Added by Morgan 2019/5/27
               ElseIf m_A1k28 = "Y55199000" Then
                  m_tmp = "Timekeeper name: NONE"
               'end 2019/5/27
               Else
                  m_tmp = "Timekeeper name: David Yen"
               End If
               
               intRRow = intRRow + 1
               PutData m_tmp, intRRow, 5500
               SetWordArray m_Head, intRRow, 3, m_tmp
               
               If m_bDowX Then
                  m_tmp = "Timekeeper identification number: JW"
               'Added by Morgan 2019/5/27
               ElseIf m_A1k28 = "Y55199000" Then
                  m_tmp = "Timekeeper identification number: NONE"
               'end 2019/5/27
               Else
                  m_tmp = "Timekeeper identification number: DY"
               End If
               intRRow = intRRow + 1
               PutData m_tmp, intRRow, 5500
               SetWordArray m_Head, intRRow, 3, m_tmp
               
               If m_bDowX Then
                  m_tmp = "Rate: USD120"
               Else
                  m_tmp = "Rate: N.A."
               End If
               intRRow = intRRow + 1
               PutData m_tmp, intRRow, 5500
               SetWordArray m_Head, intRRow, 3, m_tmp
               'end 2017/8/17
            End If
            'end 2017/3/22
            
            'Added by Morgan 2018/1/24
            '¦C¦L¹ï¶H Y54983000 Hilti Corporation ¥[¦L BU--³¯¨Ø­s
            'Removed by Morgan 2019/9/3 ¨ú®ø,§ï³]°]°È½s¸¹"BU: BU Anchors"--²øÞ±¤Z
            'If m_A1k27 = "Y54983000" Then
            '   intRRow = intRRow + 1
            '   m_tmp = "BU:"
            '   PutData m_tmp, intRRow, 5500
            '   SetWordArray m_Head, intRRow, 3, m_tmp
            '   m_tmp = "BU Direct Fastening"
            '   PutData m_tmp, intRRow, 6500
            '   SetWordArray m_Head, intRRow, 4, m_tmp
            'End If
            'end 2019/9/3
            'end 2018/1/24
            
            'Added by Morgan 2017/2/16 ¥N²z¤HY52679 Mattel, In.¤§½Ð´Ú³æ¼W¥[µo¤å¤é
            If m_A1k03 = "Y52679000" Then
               m_tmp = "Date when the work was performed:"
               intRRow = intRRow + 1
               PutData m_tmp, intRRow, 5500
               SetWordArray m_Head, intRRow, 3, m_tmp
               
               StrSQLa = "select cp27 from caseprogress where cp60='" & adoacc1k0.Fields("A1K01").Value & "' and cp27>19221111 order by cp27 desc"
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  m_tmp = ChgEngDate(rsA(0))
                  intRRow = intRRow + 1
                  PutData m_tmp, intRRow, 5500
                  SetWordArray m_Head, intRRow, 3, m_tmp
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               'Added by Morgan 2018/1/25 --³¯ª÷½¬
               If (adoacc1k0.Fields("a1k13") = "FCT" Or adoacc1k0.Fields("a1k13") = "S") Then
                  'Modified by Morgan 2019/11/12 Ms. Ling Peng¤wÂ÷Â¾--³¯ª÷½¬
                  'm_tmp = "Handling attorney's name: Ms. Ling Peng"
                  'Modified by Morgan 2023/9/28 --®}´ð?
                  'm_tmp = "Handling attorney's name: Ms. Rie Miyake"
                  m_tmp = "Handling attorney's name: Mr.Michael Moore"
                  'end 2023/9/28
                  intRRow = intRRow + 1
                  PutData m_tmp, intRRow, 5500
                  SetWordArray m_Head, intRRow, 3, m_tmp
               End If
               'end 2018/1/25
            End If
            'end 2017/2/16
            
            'Add By Sindy 2015/9/25 +Instruction No
            StrSQLa = "select cp09,cp64 from caseprogress where cp60='" & adoacc1k0.Fields("A1K01").Value & "' and instr(upper(cp64),upper('Instruction No'))>0"
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               strPrintText = Mid(rsA.Fields("cp64"), InStr(rsA.Fields("cp64"), "Instruction No"))
               varTmp = Split(strPrintText, ";")
               If UBound(varTmp) > 0 Then
                  intRRow = intRRow + 1
                  PutData CStr(varTmp(0)), intRRow, 5500
                  SetWordArray m_Head, intRRow, 3, CStr(varTmp(0))
               End If
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '2015/9/25 END
            
            'Modified by Morgan 2013/1/17
            If m_bSpecial3 Then
               m_tmp = ""
               If adoacc1k0.Fields("A1K13").Value = "FCT" Then
                  'Added by Morgan 2017/5/2
                  'Modified by Morgan 2018/1/9 --¶À«w¹F
                  If adoacc1k0.Fields("A1K02").Value > 1070000 Then
                     m_tmp = "Purchase Order: 8700820697"
                  ElseIf adoacc1k0.Fields("A1K02").Value > 1060000 Then
                     m_tmp = "Purchase Order: 8700711072"
                  'Modified by Morgan 2015/1/30
                  'm_tmp = "Purchase Order: 8700233813"
                  ElseIf adoacc1k0.Fields("A1K02").Value > 1040000 Then
                     m_tmp = "Purchase Order: 8700472463"
                  ElseIf adoacc1k0.Fields("A1K02").Value > 1030000 Then
                     m_tmp = "Purchase Order: 8700347328"
                  Else
                     m_tmp = "Purchase Order: 8700233813"
                  End If
                  'end 2015/1/30
               End If
               
               If m_tmp <> "" Then
                  intRRow = intRRow + 1
                  PutData m_tmp, intRRow, 5500
                  SetWordArray m_Head, intRRow, 3, m_tmp
               End If
               
               'Modified by Morgan 2019/2/14 µ{§Ç¤Ó¤j§ï¥Î¨ç¼Æ
               'm_tmp = "Requester:"
               'intRRow = intRRow + 1
               'PutData m_tmp, intRRow, 5500
               'SetWordArray m_Head, intRRow, 3, m_tmp
               'm_tmp = "" & adoquery.Fields("TM39").Value
               'PutData m_tmp, intRRow, 6500
               'SetWordArray m_Head, intRRow, 4, m_tmp
               PrintRightHeadCol "Requester:", intRRow, "" & adoquery.Fields("TM39").Value
               'end 2019/2/14
            
'Modified by Morgan 2019/9/9 ­ìµ{¦¡©â¥X¼g¦¨¨ç¼Æ PrintCaseNo
            'Modified Morgan 2025/10/3 ½Ð´Ú¹ï¶HY56199 Coupang CorpÄæ¦ì¦WºÙ¤£¦P¥B©T©w­n¦L§ï¦b«e­±³B²z
            'Else
            ElseIf m_A1k28 <> "Y56199000" Then
            'end 2025/10/3
               PrintCaseNo intRRow
'end 2019/9/9
            End If
            
            'Added by Morgan 2016/10/26 --Elisa
            'Modified by Morgan 2019/5/29 +Y47084
            'Modified by Morgan 2020/4/8 +Y45541--Franny
            'Modified by Morgan 2021/3/18 +Y20372--Franny
            'Modified by Morgan 2024/1/24
            'If InStr("X45349010,Y53983000,Y53983B10,Y47084000,Y45541000,Y20372000", adoacc1k0.Fields("a1k28").Value) > 0 Then
            If adoacc1k0.Fields("a1k28").Value = "Y45541000" Then
               m_tmp = "Thomas Kohl (S193979)"
               intRRow = intRRow + 1
               PutData m_tmp, intRRow, 5500
               SetWordArray m_Head, intRRow, 3, m_tmp
            ElseIf InStr("X45349010,Y53983000,Y53983B10,Y47084000,Y45541000,Y20372000", adoacc1k0.Fields("a1k28").Value) > 0 Then
            'end 2024/1/24
               m_tmp = "Contact Person: " & adoquery.Fields("TM39").Value
               intRRow = intRRow + 1
               PutData m_tmp, intRRow, 5500
               SetWordArray m_Head, intRRow, 3, m_tmp
               'Added by Morgan 2020/4/8
               'Modified by Morgan 2021/3/18 +Y20372--Franny
               'Modified by Morgan 2024/1/24
               'If InStr("Y45541000,Y20372000", adoacc1k0.Fields("a1k28").Value) > 0 Then
               If adoacc1k0.Fields("a1k28").Value = "Y20372000" Then
               'end 2024/1/24
                  If Not IsNull(adoquery.Fields("pa55").Value) Then
                     m_tmp = "¡@¡@¡@¡@¡@¡@ " & adoquery.Fields("pa55").Value
                     intRRow = intRRow + 1
                     PutData m_tmp, intRRow, 5500
                     SetWordArray m_Head, intRRow, 3, m_tmp
                  End If
               End If
               'end 2020/4/8
               
            'Added by Morgan 2022/4/22 ¥N²z¤H Y55719 ASSA ABLOY AB¤§©Ò¦³FCP®×±b³æ»ÝÅã¥Ü1st Ápµ¸¤H¡C--Franny
            ElseIf adoacc1k0.Fields("a1k03").Value = "Y55719000" And (adoacc1k0.Fields("a1k13") = "FCP" Or adoacc1k0.Fields("a1k13") = "FG" Or adoacc1k0.Fields("a1k13") = "P" Or adoacc1k0.Fields("a1k13") = "PS") Then
               intRRow = intRRow + 1
               m_tmp = "Attention:"
               PutData m_tmp, intRRow, 5500
               SetWordArray m_Head, intRRow, 3, m_tmp
               m_tmp = "" & adoquery.Fields("TM39").Value
               PutData m_tmp, intRRow, 6500
               SetWordArray m_Head, intRRow, 4, m_tmp
            'end 2022/4/22
            
            'Added by Morgan 2021/3/9 ¬ü°ê¥N²z¤HFOLEY & LARDNER¡A³]©w±b³æ¤@¨Ö±a¥X­Ó®×²Ä¤@Ápµ¸¤H--¼ï¤l·L
            ElseIf Left(adoacc1k0.Fields("a1k03").Value, 6) = "Y33940" And (adoacc1k0.Fields("a1k13") = "FCP" Or adoacc1k0.Fields("a1k13") = "FG" Or adoacc1k0.Fields("a1k13") = "P" Or adoacc1k0.Fields("a1k13") = "PS") Then
               m_tmp = "Requester: " & adoquery.Fields("TM39").Value
               intRRow = intRRow + 1
               PutData m_tmp, intRRow, 5500
               SetWordArray m_Head, intRRow, 3, m_tmp
            'end 2021/3/9
            
            'Added by Morgan 2017/6/21 --³¯¨Ø­s
            '¥N²z¤HY52622000 Keltie LLP ±b³æ¼W¥[Ápµ¸¤HÄæ¦ì¤Î®×¥ó¦WºÙ
            'Modified by Morgan 2017/10/17 +¥~°Ó¤£­n¦L--³¯ª÷½¬
            'Modified by Morgan 2018/3/20 +Y47124000 Dentons US LLP(Washington Office)¤ÎY47124010 Dentons US LLP(San Diego Office) --Tim
            'Modified by Morgan 2018/3/22 +Y20600000 BREVALEX --Lina
            'Modified by Morgan 2018/3/26 +Y53992000 Eureka IP Consulting --Lina
            'Modified by Morgan 2018/4/20 +Y51971000 Outokumpu Oyj IP Managemet --Lina
            'Modified by Morgan 2018/12/10 +Y48631000 WALLINGER RICKER SCHLOTTER TOSTMANN Patent- und Rechtsanwaelte Partnerschaft mbB --Anny
            'Modified by Lydia 2018/12/26 +Y46672000 LONZA AG PATENTABTEILING¥H¤Î Y46672010 LONZA LTD
            'Modified by Morgan 2019/2/15 +Y51409B10,Y51409B20,Y5140900 --Joseph
            'Modified by Morgan 2020/1/16 +Y33412010--Franny
            'Modified by Morgan 2020/3/6 +Y53598000--Ryan
            'Modified by Morgan 2024/6/14 +Y4794600,Y47946B1,Y47946B2 --Tim
            'Modified by Morgan 2025/8/4 +Y47946B3 --Tim
            'Modified by Morgan 2025/11/18 +Y2165207 --Teddy
            ElseIf InStr("Y52622000,Y47124000,Y47124010,Y20600000,Y53992000,Y51971000,Y48631000,Y46672000,Y46672010Y51409B10,Y51409B20,Y51409000,Y34412010,Y53598000,Y47946000,Y47946B10,Y47946B20,Y47946B30,Y21652070", adoacc1k0.Fields("a1k28").Value) > 0 And (adoacc1k0.Fields("a1k13") = "FCP" Or adoacc1k0.Fields("a1k13") = "FG" Or adoacc1k0.Fields("a1k13") = "P" Or adoacc1k0.Fields("a1k13") = "PS") Then
               intRRow = intRRow + 1
               'Modified by Morgan 2018/12/10
               'm_tmp = "Attn: "
               If adoacc1k0.Fields("a1k28") = "Y48631000" Then
                  m_tmp = "Orderer:"
               'end 2020/4/8
               Else
                  m_tmp = "Attn:"
               End If
               'end 2018/12/10
               PutData m_tmp, intRRow, 5500
               SetWordArray m_Head, intRRow, 3, m_tmp
               m_tmp = "" & adoquery.Fields("TM39").Value
               PutData m_tmp, intRRow, 6500
               SetWordArray m_Head, intRRow, 4, m_tmp
               
               'Added by Morgan 2018/3/26
               If Not IsNull(adoquery.Fields("pa55").Value) Then
                  intRRow = intRRow + 1
                  m_tmp = ""
                  PutData m_tmp, intRRow, 5500
                  SetWordArray m_Head, intRRow, 3, m_tmp
                  m_tmp = "" & adoquery.Fields("pa55").Value
                  PutData m_tmp, intRRow, 6500
                  SetWordArray m_Head, intRRow, 4, m_tmp
               End If
               'end 2018/3/26
            'end  2017/6/21
            End If
            'end 2016/10/26
            
            
            'Added by Morgan 2017/4/7
            If adoacc1k0.Fields("a1k28").Value = "Y33611B50" Then
               m_tmp = "Matter Code: " & GetClientMatterID(adoacc1k0.Fields("a1k13"), adoacc1k0.Fields("a1k14"), adoacc1k0.Fields("a1k15"), adoacc1k0.Fields("a1k16"), adoacc1k0.Fields("a1k01"), adoacc1k0.Fields("a1k28"))
               intRRow = intRRow + 1
               PutData m_tmp, intRRow, 5500
               SetWordArray m_Head, intRRow, 3, m_tmp
            End If
            'end 2017/4/7
            
            '­Y¬°®Ö­ã¥B¦³¼f©w¸¹®É, ¦L¼f©w¸¹, §_«h¦L¥Ó½Ð®×¸¹
            If "" & adoquery.Fields("TM16").Value = "1" And "" & adoquery.Fields("TM15").Value <> "" Then
               strConNo = "" & adoquery.Fields("TM15").Value
            Else
               strConNo = ""
            End If
            strPatentNo = "" & adoquery.Fields("pa22").Value
            strTradeMarkYes = "" & adoquery.Fields("Yes").Value
            strSystemName = "" & adoquery.Fields("MName").Value
            strCaseName = "" & adoquery.Fields("Cname").Value
            m_CaseName = strCaseName 'Added by Morgan 2019/5/29
            strAppNo = "" & adoquery.Fields("Ano").Value
            strCustNo = "" & adoquery.Fields("Custno").Value
            strTM09 = "" & adoquery.Fields("TM09").Value
            strTM32 = "" & adoquery.Fields("TM32").Value
         End If
         adoquery.Close
         If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If IsNull(adoacc1k0.Fields("fa22").Value) = False Then
               intRow = intRow + 1
               m_tmp = adoacc1k0.Fields("fa22").Value
               PutData m_tmp, intRow
               SetWordArray m_Head, intRow, 2, m_tmp 'Add by Morgan 2010/11/24
            End If
         Else
            If IsNull(adoacc1k0.Fields("fa36").Value) = False Then
               intRow = intRow + 1
               m_tmp = adoacc1k0.Fields("fa36").Value
               PutData m_tmp, intRow
               SetWordArray m_Head, intRow, 2, m_tmp 'Add by Morgan 2010/11/24
            End If
         End If
                  
         '¥N²z¤H/¥Ó½Ð¤H°]°È½s¸¹
         strAccNo = ""
         'Modify by Morgan 2010/8/20 ­n§ì½Ð´Ú¹ï¶Hªº--David
         'Modified by Morgan 2012/4/17 §ï§ì¦C¦L¹ï¶H Ex.X10102561 --David,³¯ª÷½¬
         'strExc(1) = m_A1k28
         strExc(1) = m_A1k27
         'end 2012/4/17
         
         'Added by Morgan 2020/3/5
         'Y33611030¤§­Ó®×¥»¨­¦³³]©wClient_Matter_ID¡A½ÐÀu¥ýÅã¥Ü­Ó®×Client_Matter_ID¡A­YµL³]©w­Ó®×Client_Matter_ID¤~Åã¥Ü¦C¦L¹ï¶H¤§°]°È½s¸¹
         If m_A1k28 = "Y33611030" Then strAccNo = strClientMatterID
         If strAccNo = "" Then
         'end 2020/3/5
         
            If Left(strExc(1), 1) = "Y" Then
               'Add By Sindy 2011/3/7 +FA106
               StrSQLa = "select FA28,FA106 from fagent where fa01='" & Left(strExc(1), 8) & "' and fa02='" & Mid(strExc(1), 9) & "'"
            Else
               'Add By Sindy 2011/3/7 +CU146
               StrSQLa = "select CU33,CU146 from customer where cu01='" & Left(strExc(1), 8) & "' and cu02='" & Mid(strExc(1), 9) & "'"
            End If
            'end 2010/8/20
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
            
               'Modify by Morgan 2010/8/20 ­n§ì½Ð´Ú¹ï¶Hªº--David
               'If "" & rsA("FA28").Value <> "" Then
               '   strAccNo = "" & rsA("FA28").Value
               'ElseIf "" & rsA("CU33").Value <> "" Then
               '   strAccNo = "" & rsA("CU33").Value
               'Else
               '   strAccNo = ""
               'End If
               'Modify By Sindy 2011/3/7
               If CheckSys("" & adoacc1k0.Fields("A1K13").Value) = "2" Or _
                  CheckSys("" & adoacc1k0.Fields("A1K13").Value) = "6" Then
                  strAccNo = "" & rsA(1)
               '2011/3/7 End
               Else
                  strAccNo = "" & rsA(0)
               End If
               'end 2010/8/20
            End If
            
         End If 'Added by Morgan 2020/3/5
         
         If strAccNo <> "" Then
            'Modfiied by Morgan 2021/7/1 +Y3326801 BASF Corporation --Anny
            If m_A1k28 <> "Y45814010" And m_A1k28 <> "Y33268010" Then 'Added by Morgan 2017/7/7 BASF ¦b¤U­±¦L
               'Added by Morgan 2015/9/30
               intRRow = intRRow + 1
               If intRRow < 6 Then intRRow = 6
               'end 2015/9/30
               
               'Added by Morgan 2024/8/2 °]°È½s¸¹¦L¦a§}¤U¤è,¥kÃä§ï¦L¥x¤@²Î½s --Franny
               If m_A1k27 = "Y55033000" And adoacc1k0.Fields("a1k13") = "FCP" Then
                  If m_Head(2, UBound(m_Head, 2)) <> "" Then intRow = intRow + 1
                  intRow = intRow + 1
                  SetWordArray m_Head, intRow, 2, strAccNo
                  strAccNo = "Tai E TAX ID. No.: 04146457"
               End If
               'end 2024/8/2
   
               'Modified by Morgan 2013/7/16 ¦³«È¤á®×¥ó®×¸¹®É·|­«Å|
               'PutData strAccNo, 6, 5500
               'Modified by Morgan 2013/9/27 ·|À£¨ì½Ð´Ú³æ¸¹§ï½Õ¾ã¤W­±Äæ¦ì¦ì¸m­pºâ¤è¦¡--Ex. X10212984
               'PutData strAccNo, 0, 5500, 3300
               PutData strAccNo, intRRow, 5500
               'end 2013/9/27
               'end 2013/7/16
               SetWordArray m_Head, intRRow, 3, strAccNo 'Add by Morgan 2010/11/24
            End If 'Added by Morgan 2017/7/7
         End If
         
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         
         PrintSpecialHead intRRow, strClientMatterID, strAccNo 'Added by Morgan 2017/12/11 µ{§Ç¤Ó¤j§ï¼g¬°¨ç¼Æ
         
         If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If IsNull(adoacc1k0.Fields("cu102").Value) = False Then
               intRow = intRow + 1
               m_tmp = adoacc1k0.Fields("cu102")
               PutData m_tmp, intRow
               SetWordArray m_Head, intRow, 2, m_tmp 'Add by Morgan 2010/11/24
            End If
         End If
         If intRow > 9 Then
            intAddSpaceRow = 1
         Else
            intAddSpaceRow = 0
         End If
         If intRRow > 6 Then intRow = intRRow 'Added by Morgan 2015/9/30
         'Add by Morgan 2011/4/14
         '¦Ü¤Ö­n¦³ 6 ¦æ,§_«h¦³¥i¯à¦C¦L·|­«Å|(¦]¬°«HÀY¥kÃä¤å¦r¬O©T©w¦ì¸m¤è¦¡¦C¦L)
         If intRow < 6 Then
            intRow = 6
         End If
         intRow = intRow + 1
         
         'Modified by Morgan 2022/8/4
         'If m_b2Printer Then
         '   Printer.CurrentX = 4000 + intDefault
         '   Printer.CurrentY = 1500 + intRow * m_LineH + intTop
         '   Printer.Print strTitle
         '   Printer.CurrentX = 7000 + intDefault
         '   Printer.CurrentY = 1500 + intRow * m_LineH + intTop
         '   Printer.Print "No."
         '   Printer.CurrentX = 7500 + intDefault
         '   Printer.CurrentY = 1500 + intRow * m_LineH + intTop
         '   Printer.Print adoacc1k0.Fields("a1k01").Value
         '   Printer.Line (0 + intDefault, 1500 + intRow * m_LineH + 350 + intTop)-(10000 + intDefault, 1500 + intRow * m_LineH + 350 + intTop)
         'End If
         'If m_b2Picture Then
         '   Picture1.CurrentX = (4000 + intDefault) * douExtRate
         '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
         '   Picture1.Print strTitle
         '   Picture1.CurrentX = (7000 + intDefault) * douExtRate
         '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
         '   Picture1.Print "No."
         '   Picture1.CurrentX = (7500 + intDefault) * douExtRate
         '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
         '   Picture1.Print adoacc1k0.Fields("a1k01").Value
         '   Picture1.Line (douExtRate * (0 + intDefault), douExtRate * (1500 + intRow * m_LineH + 350 + intTop))-(douExtRate * (10000 + intDefault), douExtRate * (1500 + intRow * m_LineH + 350 + intTop))
         'End If
         PutData strTitle, intRow, 4000
         PutData "No. " & adoacc1k0.Fields("a1k01").Value, intRow, 7000
         If m_b2Printer Then
            Printer.Line (0 + intDefault, 1500 + intRow * m_LineH + 350 + intTop)-(10000 + intDefault, 1500 + intRow * m_LineH + 350 + intTop)
         End If
         'end 2022/8/4
                  
         intRow = intRow + 1
         'Modified by Morgan 2015/10/30
         'Modified by Morgan 2016/3/15
         'If m_A1k28 = "Y52960000" Then
         '   m_Title = "INVOICE"
         'Else
         '   m_Title = "DEBIT NOTE" 'Add by Morgan 2010/12/1
         'End If
         m_Title = strTitle
         'end 2016/3/15
         'end 2015/10/30
         
      Case "3" '¤é¤å
        '¥N²z¤H¦WºÙ
        If IsNull(adoacc1k0.Fields("fa06").Value) = False Then
            'Add by Morgan 2005/12/6
            ''Modified by Morgan 2015/7/24 §ï»Pa1k04³]©w¤@­P§PÂ_a1k28¥B¤£­­©w¥Ó½Ð¤H
            If "" & adoacc1k0("a1k13").Value = "FCP" And ("" & adoacc1k0("a1k28").Value = "Y34271000") Then
               intRow = intRow + 1
            End If
            m_tmp = adoacc1k0.Fields("fa06").Value & "¡@" & "±s¤¤"
            PutData m_tmp, intRow
            SetWordArray m_Head, intRow - 7, 2, m_tmp 'Add by Morgan 2010/12/1
            'Added by Morgan 2015/11/19
            '¥_¨Ê»ÈÀs¥[¦L¤¤¤å¦a§}
            'modify by sonia 2019/6/13 +Y5245900
            'modify by sonia 2019/7/10 §ï¥þ³¡³£¦L
            'Modified by Morgan 2019/7/11 ÁÙ­ì,¤é¤åÁÙ¬O­n±±¨î
            If m_A1k28 = "Y51333010" Or m_A1k28 = "Y52459000" Then
               If Not IsNull(adoacc1k0.Fields("fa17").Value) Then
                  PutData "" & adoacc1k0.Fields("fa17").Value, intRow
                  SetWordArray m_Head, intRow - 6, 2, adoacc1k0.Fields("fa17")
               End If
            End If
            'End 2015/11/19
        'Add by Morgan 2009/11/23
        '¥N²z¤H­^¤å¦WºÙ
        ElseIf IsNull(adoacc1k0.Fields("fa05").Value) = False Then
            intRow = intRow + 1
            PutData adoacc1k0.Fields("fa05").Value, intRow
            SetWordArray m_Head, intRow - 7, 2, adoacc1k0.Fields("fa05").Value 'Add by Morgan 2010/12/1
            If IsNull(adoacc1k0.Fields("fa63").Value) = False Then
               intRow = intRow + 1
               PutData adoacc1k0.Fields("fa63").Value, intRow
               SetWordArray m_Head, intRow - 7, 2, adoacc1k0.Fields("fa63").Value 'Add by Morgan 2010/12/1
            End If
            If IsNull(adoacc1k0.Fields("fa64").Value) = False Then
               intRow = intRow + 1
               PutData adoacc1k0.Fields("fa64").Value, intRow
               SetWordArray m_Head, intRow - 7, 2, adoacc1k0.Fields("fa64").Value 'Add by Morgan 2010/12/1
            End If
            If IsNull(adoacc1k0.Fields("fa65").Value) = False Then
               intRow = intRow + 1
               PutData adoacc1k0.Fields("fa65").Value, intRow
               SetWordArray m_Head, intRow - 7, 2, adoacc1k0.Fields("fa65").Value 'Add by Morgan 2010/12/1
            End If
        End If
         intRow = intRow + 1
         '½Ð´Ú¤é´Á
         m_tmp = Format(ChangeTStringToWString("" & adoacc1k0.Fields("A1K02").Value), "####¦~##¤ë##¤é")
         PutData m_tmp, 0, 7000, 4200
         'Modified by Morgan 2013/12/16 ¤é¤å©ïÀY¥i¯à¸ûªø¤S¤£©y¸õ¦æ,¤é´Á­n©ñ¤U¤@¦C(¦Pª½±µ¦C¦L)
         iHeadRow = 3
         'end 2013/12/16
         SetWordArray m_Head, iHeadRow, 3, m_tmp  'Add by Morgan 2010/12/1
         
        intRow = intRow + 1
        adoquery.CursorLocation = adUseClient
         '2007/7/24 MODIFY BY SONIA TS¬d¦W®×µL²Õ¸s®É¦LÃþ§O
         'Modified by Morgan 2013/5/14 +pa08,pa09
         'Modify By Sindy 2015/7/8 tm05 ==> nvl(tm131,tm05)
         'Modified by Morgan 2020/9/30 pa06 as Cname-->nvl(pa07,pa06) as Cname
        adoquery.Open "select pa77 as Yno, pa48 as Cno, ptm06 as MName, nvl(pa07,pa06) as Cname, pa11 as Ano, pa26 as Custno, pa22, null as tm15, null as Yes, '' As TM12, '' As TM16, '' As TM09, '' AS TM32,pa08,pa09 from patent, systemkind, patenttrademarkmap where pa01 = sk01 and pa08 = ptm02 (+) and sk02 = ptm01 and pa01 = '" & adoacc1k0.Fields("a1k13").Value & "' and pa02 = '" & adoacc1k0.Fields("a1k14").Value & "' and pa03 = '" & adoacc1k0.Fields("a1k15").Value & "' and pa04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                        "select tm45 as Yno, tm35 as Cno, ptm06 as MName, Rtrim(Ltrim(nvl(tm131,tm05)||' '||tm06)) as Cname, tm12 as Ano, tm23 as Custno, null as pa22, tm15, '1' as Yes, TM12, TM16, TM09, TM32,tm08 as pa08,tm10 as pa09 from trademark, systemkind, patenttrademarkmap where tm01 = sk01 and tm08 = ptm02 (+) and sk02 = ptm01 and tm01 = '" & adoacc1k0.Fields("a1k13").Value & "' and tm02 = '" & adoacc1k0.Fields("a1k14").Value & "' and tm03 = '" & adoacc1k0.Fields("a1k15").Value & "' and tm04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                        "select lc23 as Yno, lc17 as Cno, '' as MName, nvl(lc05,lc06) as Cname, '' as Ano, lc11 as Custno, null as pa22, null as tm15, null as Yes, '' As TM12, '' As TM16, '' As TM09, '' AS TM32,'' as pa08,'000' as pa09 from lawcase where lc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and lc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and lc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and lc04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                        "select sp27 as Yno, sp29 as Cno, '' as MName, nvl(sp05,sp06) as Cname, sp11 as Ano, sp08 as Custno, null as pa22, null as tm15, null as Yes, '' As TM12, '' As TM16, sp73 As TM09, SP74 AS TM3,'' as pa08,sp09 as pa09 from servicepractice where sp01 = '" & adoacc1k0.Fields("a1k13").Value & "' and sp02 = '" & adoacc1k0.Fields("a1k14").Value & "' and sp03 = '" & adoacc1k0.Fields("a1k15").Value & "' and sp04 = '" & adoacc1k0.Fields("a1k16").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
        If adoquery.RecordCount <> 0 Then
            m_CaseNo = "" & adoquery.Fields("Cno").Value  'Added by Morgan 2023/10/13
            'Modified by Morgan 2022/7/27
            'm_tmp = "¶Q¤è¾ã²zµf†A¡G"
            m_tmp = "¶Q¤è¾ã²zµf" & PUB_GetUniText(Me.Name, "¸¹") & "¡G"
            'end 2022/7/27
            '­Y¬°FCP®×
            'Modified by Morgan 2015/8/3 +°Ó¼Ð©µ®i
            If InStr(",P,CFP,FCP,T,FCT,CFT,TF,", "," & adoacc1k0.Fields("A1K13").Value & ",") > 0 Then
               If InStr(",P,CFP,FCP,", "," & adoacc1k0.Fields("A1K13").Value & ",") > 0 Then
                  StrSQLa = "Select PA106 From Patent, CaseProgress Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP60='" & adoacc1k0.Fields("A1K01").Value & "' And CP10='605'  and pa76 is not null"
               Else
                  StrSQLa = "Select TM65 From TRADEMARK, CaseProgress Where TM01=CP01 And TM02=CP02 And TM03=TM03 And TM04=CP04 And CP60='" & adoacc1k0.Fields("A1K01").Value & "' And CP10='102' and tm33  is not null"
               End If
            'end 2015/8/3
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  m_tmp = m_tmp & rsA.Fields(0).Value
               Else
                  If Not IsNull(adoquery.Fields("Yno").Value) Then
                     m_tmp = m_tmp & adoquery.Fields("Yno").Value
                  End If
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            '­Y«DFCP®×
            Else
               If Not IsNull(adoquery.Fields("Yno").Value) Then
                  m_tmp = m_tmp & adoquery.Fields("Yno").Value
               End If
            End If
            'End
            
            PutData m_tmp, 0, 5500, 4500
            iHeadRow = iHeadRow + 1 'Add by Morgan 2010/12/1
            SetWordArray m_Head, iHeadRow, 3, m_tmp 'Add by Morgan 2010/12/1
            
            intRow = intRow + 1
            'Modified by Morgan 2022/7/27
            'm_tmp = "¹ú©Ò¾ã²zµf†A¡G" & Replace(adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value, "-0-00", "")
            m_tmp = "¹ú©Ò¾ã²zµf" & PUB_GetUniText(Me.Name, "¸¹") & "¡G" & Replace(adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value, "-0-00", "")
            'end 2022/7/27
            PutData m_tmp, 0, 5500, 4800
            iHeadRow = iHeadRow + 1 'Add by Morgan 2010/12/1
            SetWordArray m_Head, iHeadRow, 3, m_tmp 'Add by Morgan 2010/12/1
            
            intRow = intRow + 1
            m_tmp = ""
            If IsNull(adoquery.Fields("Cno").Value) = False Then
               m_tmp = "Case No.:" & adoquery.Fields("Cno").Value
            ElseIf adoacc1k0.Fields("a1k03").Value = "Y20438020" Then
               m_tmp = "Vendor Code # 88125-0"
            End If
            If m_tmp <> "" Then
               PutData m_tmp, 0, 5500, 5100
               iHeadRow = iHeadRow + 1 'Add by Morgan 2010/12/1
               SetWordArray m_Head, iHeadRow, 3, m_tmp 'Add by Morgan 2010/12/1
            End If
            '­Y¬°®Ö­ã¥B¦³¼f©w¸¹®É, ¦L¼f©w¸¹, §_«h¦L¥Ó½Ð®×¸¹
            If "" & adoquery.Fields("TM16").Value = "1" And "" & adoquery.Fields("TM15").Value <> "" Then
               strConNo = "" & adoquery.Fields("TM15").Value
            Else
               strConNo = ""
            End If
            strPatentNo = "" & adoquery.Fields("pa22").Value
            strTradeMarkYes = "" & adoquery.Fields("Yes").Value
            strSystemName = "" & adoquery.Fields("MName").Value
            strCaseName = "" & adoquery.Fields("Cname").Value
            m_CaseName = strCaseName 'Added by Morgan 2019/5/29
            strAppNo = "" & adoquery.Fields("Ano").Value
            strCustNo = "" & adoquery.Fields("Custno").Value
            strTM09 = "" & adoquery.Fields("TM09").Value
            strTM32 = "" & adoquery.Fields("TM32").Value
            strPA08 = "" & adoquery.Fields("pa08").Value 'Added by Morgan 2013/5/14
            strPA09 = "" & adoquery.Fields("pa09").Value 'Added by Morgan 2013/5/14
         End If
         adoquery.Close
         intRow = intRow + 1
        '¥N²z¤H/¥Ó½Ð¤H°]°È½s¸¹
        'Modify by Morgan 2010/8/20 ­n§ì½Ð´Ú¹ï¶Hªº--David
         'Modified by Morgan 2012/4/17 §ï§ì¦C¦L¹ï¶H Ex.X10102561 --David,³¯ª÷½¬
         'strExc(1) = m_A1k28
         strExc(1) = m_A1k27
         'end 2012/4/17
        If Left(strExc(1), 1) = "Y" Then
            'Add By Sindy 2011/3/7 +FA106
            StrSQLa = "select FA28,FA106 from fagent where fa01='" & Left(strExc(1), 8) & "' and fa02='" & Mid(strExc(1), 9) & "'"
         Else
            'Add By Sindy 2011/3/7 +CU146
            StrSQLa = "select CU33,CU146 from customer where cu01='" & Left(strExc(1), 8) & "' and cu02='" & Mid(strExc(1), 9) & "'"
         End If
         'end 2010/8/20
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            'Modify by Morgan 2010/8/20 ­n§ì½Ð´Ú¹ï¶Hªº--David
            'If "" & rsA("FA28").Value <> "" Then
            '    m_tmp = "" & rsA("FA28").Value
            'ElseIf "" & rsA("CU33").Value <> "" Then
            '    m_tmp = "" & rsA("CU33").Value
            'Else
            '   m_tmp = ""
            'End If
            'Modify By Sindy 2011/3/7
            If CheckSys("" & adoacc1k0.Fields("A1K13").Value) = "2" Or _
               CheckSys("" & adoacc1k0.Fields("A1K13").Value) = "6" Then
               m_tmp = "" & rsA(1)
            '2011/3/7 End
            Else
               m_tmp = "" & rsA(0)
            End If
            'end 2010/8/20
            If m_tmp <> "" Then
               PutData m_tmp, 0, 5500, 5400
               iHeadRow = iHeadRow + 1 'Add by Morgan 2010/12/1
               SetWordArray m_Head, iHeadRow, 3, m_tmp 'Add by Morgan 2010/12/1
            End If
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        
         'Added by Morgan 2024/3/25
         If m_bSpecialNew5 Then
            m_tmp = "½Ð¨DªÌ¡G¯S³\¤é¥»³¡³¡ªø"
            iHeadRow = iHeadRow + 1
            SetWordArray m_Head, iHeadRow, 3, m_tmp
            
            m_tmp = String(4, "¡@") & GetStaffName(GetDeptMan("J00", 2))
            iHeadRow = iHeadRow + 1
            SetWordArray m_Head, iHeadRow, 3, m_tmp
            
            m_tmp = PUB_GetUniText(Me.Name, "­p®É¦¬¶O") & "¡G" & IIf(m_CP01 = "P", "USD200", "NTD6,000")
            iHeadRow = iHeadRow + 1
            SetWordArray m_Head, iHeadRow, 3, m_tmp
            
            intRow = intRow + 2
         End If
         'end 2024/3/25
   
        m_tmp = ReportSum(130)
        PutData m_tmp, intRow
        SetWordArray m_Head, intRow - 7, 2, m_tmp 'Add by Morgan 2010/12/1
        
         intRow = intRow + 1
         If m_b2Printer Then
            Printer.Line (0 + intDefault, 1500 + intRow * m_LineH + 350 + intTop)-(10000 + intDefault, 1500 + intRow * m_LineH + 350 + intTop)
         End If
         'Removed by Morgan 2022/8/4
         'If m_b2Picture Then
         '   Picture1.Line (douExtRate * (0 + intDefault), douExtRate * (1500 + intRow * m_LineH + 350 + intTop))-(douExtRate * (10000 + intDefault), douExtRate * (1500 + intRow * m_LineH + 350 + intTop))
         'End If
         'end 2022/8/4
         intRow = intRow + 1
   End Select

'******************************************************************************
   If adoquery.State = adStateOpen Then
       adoquery.Close
   End If
   strTM28 = ""
   '­Y¨t²ÎÃþ§O¬°FCT
   If ("" & adoacc1k0.Fields("a1k13").Value = "FCT" Or "" & adoacc1k0.Fields("a1k13").Value = "T") Then
       StrSQLa = "Select * From Trademark Where " & ChgTradeMark("" & adoacc1k0.Fields("a1k13").Value & "" & adoacc1k0.Fields("a1k14").Value & "" & adoacc1k0.Fields("a1k15").Value & "" & adoacc1k0.Fields("a1k16").Value) & " And TM28 Is Not Null And TM28<>'1' "
       rsA.CursorLocation = adUseClient
       rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
       '­Y¦³®×¥ó©Ê½è¬°²§Ä³, µû©w, ¼o¤î
       If rsA.RecordCount > 0 Then
           strTM28 = "" & rsA("TM28").Value
           strConNo = "" & rsA("TM15").Value
       End If
       If rsA.State <> adStateClosed Then rsA.Close
       Set rsA = Nothing
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select cp10, cp01 from caseprogress where cp60 = '" & adoacc1k0.Fields("a1k01").Value & "' and cp10 >= '101' and cp10 <= '105'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      Select Case adoquery.Fields("cp01").Value
         Case "CFT", "FCT", "T"
            If adoquery.Fields("cp10").Value = "101" Then
               strProperty = "New "
            Else
               strProperty = ""
            End If
         Case "FCP", "FG"
            strProperty = ""
         Case Else
            strProperty = "New "
      End Select
   Else
      strProperty = ""
   End If
   adoquery.Close
   intRow = intRow + 1
   If m_b2Printer Then
      Printer.CurrentX = 0 + intDefault
      Printer.CurrentY = 1500 + intRow * m_LineH + intTop
   End If
   'Removed by Morgan 2022/8/4
   'If m_b2Picture Then
   '   Picture1.CurrentX = (0 + intDefault) * douExtRate
   '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
   'End If
   'end 2022/8/4
   strCache1 = "": strCache2 = "" 'Add by Morgan 2010/12/1
   strPrintText = ""
   '­Y¬°°Ó¼Ð®×
   If strTradeMarkYes = "1" Then
      'Modify by Morgan 2006/11/22 ¥[¤¤¤å
      Select Case strTM28
         Case "2" '²§Ä³
            If strLanguage = "1" Then
               strCache1 = "¥DÃD¡G"
               'Modify by Morgan 2007/1/15 ¥[¦L®×¥ó©Ê½è
               strCache2 = GetNationName(strLanguage) & "°Ó¼Ð²§Ä³®×" & GetCasePropertyName("" & adoacc1k0.Fields("A1K01").Value)
               'modify by sonia 2017/3/7 T-208134 ²§Ä³®×¤§³¯­z·N¨£®Ñ¶¥¬qµLµù¥U¸¹
               'strPrintText = "µù¥U¸¹¡G" & strConNo
               If strConNo <> "" Then
                  strPrintText = "µù¥U¸¹¡G" & strConNo
               Else
                  strPrintText = "¥Ó½Ð®×¸¹¡G" & strAppNo
               End If
               'end 2017/3/7
            Else
               strCache1 = "Re: "
               strCache2 = "Opposition Action against " & GetNationName(strLanguage) & "Mark"
               'Modify by Morgan 2007//2/2 --³¯ª÷½¬
               'strPrintText = "Approved No. " & strConNo
               strPrintText = "Registration No. " & strConNo
               'end 2007/2/2
            End If
            
         Case "3" 'µû©w
            If strLanguage = "1" Then
               strCache1 = "¥DÃD¡G"
               'Modify by Morgan 2007/1/15 ¥[¦L®×¥ó©Ê½è
               strCache2 = GetNationName(strLanguage) & "°Ó¼Ðµû©w®×" & GetCasePropertyName("" & adoacc1k0.Fields("A1K01").Value)
               strPrintText = "µù¥U¸¹¡G" & strConNo
            Else
               strCache1 = "Re: "
               strCache2 = "Invalidation Action against " & GetNationName(strLanguage) & "Mark"
               strPrintText = "Registration No. " & strConNo
            End If
         
         Case "4" 'Á|µo(¼o¤î)
            If strLanguage = "1" Then
               strCache1 = "¥DÃD¡G"
               'Modify by Morgan 2007/1/15 ¥[¦L®×¥ó©Ê½è
               strCache2 = GetNationName(strLanguage) & "°Ó¼Ð¼o¤î®×" & GetCasePropertyName("" & adoacc1k0.Fields("A1K01").Value)
               strPrintText = "µù¥U¸¹¡G" & strConNo
            Else
               strCache1 = "Re: "
               strCache2 = "Revocation Action against " & GetNationName(strLanguage) & "Mark"
               strPrintText = "Registration No. " & strConNo
            End If
            
         Case Else '¨ä¥L
            Select Case strLanguage
               Case "1"
                  strCache1 = "¥DÃD¡G"
                  'Modified by Lydia 2015/04/09 ¾ã§å½Ð´Ú³æªº®×¥ó©Ê½è¦b©ú²Ó
                  'strCache2 = GetNationName(strLanguage) & "°Ó¼Ð" & GetCasePropertyName("" & adoacc1k0.Fields("A1K01").Value)
                  strCache2 = GetNationName(strLanguage) & "°Ó¼Ð" & IIf(m_bolChiDB = True And bol_ChiDB = True, "", GetCasePropertyName("" & adoacc1k0.Fields("A1K01").Value))
               Case "2"
                  strCache1 = "Re: "
                  If strConNo <> "" Then
                     'Modified by Morgan 2012/12/19 ¥h±¼¥kÃäªºªÅ¥Õ
                     strCache2 = RTrim(GetNationName(strLanguage) & strProperty & strSystemName) & " Registration No. " & strConNo
                  Else
                     'Modified by Morgan 2012/12/19 ¥h±¼¥kÃäªºªÅ¥Õ
                     strCache2 = RTrim(GetNationName(strLanguage) & strProperty & strSystemName) & " Application No. " & strAppNo
                  End If
               Case "3"
                  strCache1 = "¥ó¦W¡G"
                  If CheckStr(adoacc1k0.Fields("A1K13").Value) = "FCT" Then
                     If CheckStr(adoacc1k0.Fields("A1j02").Value) = "101" Then
                        'Modified by Morgan 2022/8/3
                        'strCache2 = "¥x“ø°Ó¼Ð" & GetCasePropertyName("" & adoacc1k0.Fields("A1K01").Value) & "µn“÷¥XÄ@No." & strAppNo
                        strCache2 = "¥x" & PUB_GetUniText(Me.Name, "ÆW") & "°Ó¼Ð" & GetCasePropertyName("" & adoacc1k0.Fields("A1K01").Value) & "µn" & PUB_GetUniText(Me.Name, "¿ý") & "¥XÄ@No." & strAppNo
                        'end 2022/8/3
                     Else
                        'Modify by Morgan 2011/5/4 FCT ®×­Y¦³¼f©w¸¹«h¤£¥²¦L¥Ó½Ð¸¹
                        'strCache2 = "¥x“ø°Ó¼Ð" & GetCasePropertyName("" & adoacc1k0.Fields("A1K01").Value) & "µn“÷¥XÄ@No." & strAppNo & IIf(strConNo <> "", "¡]µn“÷²Ä" & strConNo & "†A¡^", "")
                        'Modified by Morgan 2022/8/3
                        'strCache2 = "¥x“ø°Ó¼Ð" & GetCasePropertyName("" & adoacc1k0.Fields("A1K01").Value) & IIf(strConNo <> "", "¡]µn“÷²Ä" & strConNo & "†A¡^", "µn“÷¥XÄ@No." & strAppNo)
                        strCache2 = "¥x" & PUB_GetUniText(Me.Name, "ÆW") & "°Ó¼Ð" & GetCasePropertyName("" & adoacc1k0.Fields("A1K01").Value) & IIf(strConNo <> "", "¡]µn" & PUB_GetUniText(Me.Name, "¿ý") & "²Ä" & strConNo & PUB_GetUniText(Me.Name, "¸¹") & "¡^", "µn" & PUB_GetUniText(Me.Name, "¿ý") & "¥XÄ@No." & strAppNo)
                        'end 2022/8/3
                     End If
                  Else
                     '2011/6/21 MODIFY BY SONIA §ì¥Ó½Ð°ê®a¦WºÙ¥B­Y¦³¼f©w¸¹«h¤£¥²¦L¥Ó½Ð¸¹
                     'strCache2 = "¥x“ø°Ó¼Ð" & GetCasePropertyName("" & adoacc1k0.Fields("A1K01").Value) & "µn“÷¥XÄ@No." & strAppNo & IIf(strConNo <> "", "¡]µn“÷²Ä" & strConNo & "†A¡^", "")
                     'Modified by Morgan 2022/8/3
                     'strCache2 = GetNationName(strLanguage) & "°Ó¼Ð" & GetCasePropertyName("" & adoacc1k0.Fields("A1K01").Value) & IIf(strConNo <> "", "µn“÷²Ä" & strConNo & "†A", "µn“÷¥XÄ@No." & strAppNo)
                     strCache2 = GetNationName(strLanguage) & "°Ó¼Ð" & GetCasePropertyName("" & adoacc1k0.Fields("A1K01").Value) & IIf(strConNo <> "", "µn" & PUB_GetUniText(Me.Name, "¿ý") & "²Ä" & strConNo & PUB_GetUniText(Me.Name, "¸¹"), "µn" & PUB_GetUniText(Me.Name, "¿ý") & "¥XÄ@No." & strAppNo)
                     'end 2022/8/3
                  End If
            End Select
      End Select
      
   'Add by Morgan 2006/11/17
   ElseIf m_CP01 = "TS" Then
      Select Case strLanguage
         Case "1"
            strCache1 = "¥DÃD¡G"
            strCache2 = GetNationName(strLanguage) & "°Ó¼Ð" & GetCasePropertyName("" & adoacc1k0.Fields("A1K01").Value)
         '2012/4/11 ADD BY SONIA
         Case "2"
            strCache1 = "Re: "
            strCache2 = "Trademark Search in " & GetNationName(strLanguage)
         '2012/4/11 END
      End Select
   
   'Add by Morgan 2007/4/19
   ElseIf m_CP01 = "TM" Then
      Select Case strLanguage
         Case "1"
            strCache1 = "¥DÃD¡G"
            strCache2 = GetNationName(strLanguage) & GetCasePropertyName("" & adoacc1k0.Fields("A1K01").Value)
            
         Case Else
            strCache1 = "Re: "
            strCache2 = "Monitoring system (C.C.C. Code) in " & GetNationName(strLanguage)
      End Select
      
   '­Y«D°Ó¼Ð®×
   Else
      Select Case m_CP01
         '­Y¨t²ÎÃþ§O¬°"S"
         Case "S"
            'Modify By Sindy 2017/5/5
            If strLanguage = "3" And _
               CheckStr(adoacc1k0.Fields("A1j02").Value) = "001" Then 'S¤é¤å001¬d¦W
               strCache1 = "¥ó¦W¡G" & vbCrLf & vbCrLf
               'Modified by Morgan 2022/8/3
               'strCache2 = strCache2 & "¥x“ø°Ó¼Ð½Õ¬d" & vbCrLf
               strCache2 = strCache2 & "¥x" & PUB_GetUniText(Me.Name, "ÆW") & "°Ó¼Ð½Õ¬d" & vbCrLf
               'end 2022/8/3
               'Modify By Sindy 2018/6/27 µL¦WºÙ®É,³s¼ÐÃD³£¤£Åã¥Ü
               If strCustName(0) <> "" Or strCustName(1) <> "" Or _
                  strCustName(2) <> "" Or strCustName(3) <> "" Or strCustName(4) <> "" Then
               '2018/6/27 END
                  'Modified by Morgan 2022/8/4
                  'strCache2 = strCache2 & "þç¨Ì“ß¤H¡G"
                  strCache2 = strCache2 & PUB_GetUniText(Me.Name, "«È¤á¡G")
                  'end 2022/8/4
                  If strCustName(0) <> "" Then
                     strCache1 = strCache1 & vbCrLf
                     strCache2 = strCache2 & strCustName(0) & vbCrLf
                  End If
                  If strCustName(1) <> "" Then
                     strCache1 = strCache1 & vbCrLf
                     strCache2 = strCache2 & "¡@¡@¡@¡@¡@" & strCustName(1) & vbCrLf
                  End If
                  If strCustName(2) <> "" Then
                     strCache1 = strCache1 & vbCrLf
                     strCache2 = strCache2 & "¡@¡@¡@¡@¡@" & strCustName(2) & vbCrLf
                  End If
                  If strCustName(3) <> "" Then
                     strCache1 = strCache1 & vbCrLf
                     strCache2 = strCache2 & "¡@¡@¡@¡@¡@" & strCustName(3) & vbCrLf
                  End If
                  If strCustName(4) <> "" Then
                     strCache1 = strCache1 & vbCrLf
                     strCache2 = strCache2 & "¡@¡@¡@¡@¡@" & strCustName(4) & vbCrLf
                  End If
               End If
               'Modified by Morgan 2022/8/4
               'strCache2 = strCache2 & "°Ó¡@¼Ð¡G" & GetCaseName(m_CP01 & m_CP02 & m_CP03 & m_CP04) & vbCrLf & "üP¡@¤À¡G²Ä" & strTM09 & "Ãþ"
               strCache2 = strCache2 & "°Ó¡@¼Ð¡G" & GetCaseName(m_CP01 & m_CP02 & m_CP03 & m_CP04) & vbCrLf & PUB_GetUniText(Me.Name, "°Ï") & "¡@¤À¡G²Ä" & strTM09 & "Ãþ"
               'end 2022/8/4
            Else
            '2017/5/5 END
               strCache1 = "Re: "
               'Modified by Morgan 2012/12/19 ¥h±¼¥kÃäªºªÅ¥Õ
               'Modified by Morgan 2022/12/19 ¥u¦³½Ð¡u¬d¦W¡v®É¤~±a¡A¨ä¾lªÅ¥Õ--³¯ª÷½¬
               strCache2 = ""
               If ChkItemExist(adoacc1k0.Fields("A1K01").Value, "001") Then
                  strCache2 = RTrim(GetNationName(strLanguage) & strProperty & strSystemName) & " Trademark Search for " & GetCaseName(m_CP01 & m_CP02 & m_CP03 & m_CP04)
               End If
            End If
            
         'Add by Morgan 2006/6/27
         Case "TD"
            strCache1 = "Re: "
            'Modified by Morgan 2012/12/19 ¥h±¼¥kÃäªºªÅ¥Õ
            strCache2 = RTrim(GetNationName(strLanguage) & strProperty & strSystemName) & " domain name: " & strCaseName
            
         'Modified by Morgan 2015/3/16 +PS Ex. X10403137
         Case "FG", "PS"
            strCache1 = "Re: "
            strCache2 = GetCaseName(m_CP01 & m_CP02 & m_CP03 & m_CP04, strLanguage)
            
         Case Else '¨ä¥L¨t²ÎÃþ§O
            Select Case strLanguage
               Case "1"
                  strCache1 = "¥DÃD¡G"
                  '2009/9/2 ADD BY SONIA
                  If m_CP01 = "TT" Then
                     strCache2 = GetCaseName(m_CP01 & m_CP02 & m_CP03 & m_CP04, strLanguage)
                  Else
                  '2009/9/2 END
                     'Modified by Morgan 2024/7/8
                     'strCache2 = GetNationName(strLanguage) & strSystemName & "¥Ó½Ð"
                     If m_CP01 = "P" Or m_CP01 = "FCP" Or m_CP01 = "CFP" Then
                        strCache1 = "®×¥ó¡G"
                        strCache2 = GetNationName(strLanguage) & strSystemName & "±M§Q¥Ó½Ð"
                     Else
                        strCache2 = GetNationName(strLanguage) & strSystemName & "¥Ó½Ð"
                     End If
                     'end 2024/7/8
                     'Modify by Morgan 2008/5/30 ¦³¥Ó½Ð¸¹¤~¦L
                     'strCache2 = strCache2 & "²Ä " & IIf(strAppNo = "", "        ", strAppNo) & " ¸¹"
                     If strAppNo <> "" Then
                        strCache2 = strCache2 & "²Ä " & strAppNo & " ¸¹"
                     End If
                  End If
                  
               Case "2"
                  strCache1 = "Re: "
                  If strPatentNo <> "" Then
                     'Modified by Morgan 2012/12/19 ¥h±¼¥kÃäªºªÅ¥Õ
                     strCache2 = RTrim(GetNationName(strLanguage) & strProperty & strSystemName) & " Application No. " & strAppNo & " (Patent No. " & strPatentNo & ")"
                  Else
                     'Modified by Morgan 2012/12/19 ¥h±¼¥kÃäªºªÅ¥Õ
                     strCache2 = RTrim(GetNationName(strLanguage) & strProperty & strSystemName) & " Application No. " & strAppNo
                  End If
                  
               Case "3"
                  strCache1 = "¥ó¦W¡G"
                  'Modified by Morgan 2013/5/13
                  'strCache2 = "¥x“ø" & strSystemName & "¥XÄ@²Ä" & IIf(strAppNo = "", "        ", strAppNo) & "†A"
                  If m_CP01 = "P" And strPA09 = "013" And strPA08 = "1" Then
                     'Modified by Morgan 2022/8/4
                     'strCache2 = "­»´ä¼Ð·Ç¯S³\µn“÷²Ä" & IIf(strAppNo = "", "        ", strAppNo) & "†A"
                     strCache2 = "­»´ä¼Ð·Ç¯S³\µn" & PUB_GetUniText(Me.Name, "¿ý") & "²Ä" & IIf(strAppNo = "", "        ", strAppNo) & PUB_GetUniText(Me.Name, "¸¹")
                     'end 2022/8/4
                  Else
                     'Modified by Morgan 2022/8/4
                     'strCache2 = GetNationName(strLanguage) & strSystemName & "¥XÄ@²Ä" & IIf(strAppNo = "", "        ", strAppNo) & "†A"
                     strCache2 = GetNationName(strLanguage) & strSystemName & "¥XÄ@²Ä" & IIf(strAppNo = "", "        ", strAppNo) & PUB_GetUniText(Me.Name, "¸¹")
                     'end 2022/8/4
                  End If
            End Select
      End Select
   End If
   
  'Added by Lydia 2015/04/08  §PÂ_¾ã§å½Ð´Ú³æªí­º¥u°O¿ý²Ä¤@µ§
   If m_bolChiDB = False Or (m_bolChiDB And adoacc1k0.AbsolutePosition = 1) Then
        If strCache1 & strCache2 <> "" Then
          If m_b2Printer Then
             'Modified by Morgan 2022/8/4
             'Printer.Print strCache1 & strCache2
             PUB_PrintUnicodeText strCache1 & strCache2, Printer.CurrentX, Printer.CurrentY, 0
             'end 2022/8/4
          End If
          'Removed by Morgan 2022/8/4
          'If m_b2Picture Then
          '   Picture1.Print strCache1 & strCache2
          'End If
          'end 2022/8/4
          'Add by Morgan 2010/11/25
          m_iSubject = 1
          SetWordArray m_Subject, 1, 1, Trim(strCache1)
          SetWordArray m_Subject, 1, 2, Trim(strCache2)
        End If
        
        If strCustNo = "X22232010" Then
          intRow = intRow + 1
          PutData strCaseName, intRow, 400
          m_iSubject = m_iSubject + 1 'Add by Morgan 2010/11/25
          SetWordArray m_Subject, m_iSubject, 2, strCaseName 'Add by Morgan 2010/11/25
        End If
   End If
   '­Y¬°°Ó¼Ð®×¥ó
   If strTradeMarkYes = "1" Then
      Select Case strLanguage
         Case "1"
           'Added by Lydia 2015/04/08  §PÂ_¾ã§å½Ð´Ú³æ¥u°O¿ý²Ä¤@µ§
           If m_bolChiDB = False Or (m_bolChiDB And adoacc1k0.AbsolutePosition = 1) Then
                intRow = intRow + 1
                If m_b2Printer Then
                   Printer.CurrentX = Printer.TextWidth("¥DÃD¡G") + intDefault
                   Printer.CurrentY = 1500 + intRow * m_LineH + intTop
                End If
                'Removed by Morgan 2022/8/4
                'If m_b2Picture Then
                '   Picture1.CurrentX = Picture1.TextWidth("¥DÃD¡G") + (intDefault) * douExtRate
                '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
                'End If
                'end 2022/8/4
                If CountLength(strCustName(0)) <= 80 Then
                   If m_b2Printer Then
                      'Modified by Morgan 2022/8/4
                      'Printer.Print "¥Ó½Ð¤H¡G" & strCustName(0)
                      PUB_PrintUnicodeText "¥Ó½Ð¤H¡G" & strCustName(0), Printer.CurrentX, Printer.CurrentY, 0
                      'end 2022/8/4
                   End If
                   'Removed by Morgan 2022/8/4
                   'If m_b2Picture Then
                   '   Picture1.Print "¥Ó½Ð¤H¡G" & strCustName(0)
                   'End If
                   'end 2022/8/4
                Else
                    'Modified by Morgan 2012/3/12
                    'PrintDropLine "" & strCustName(0), "¥Ó½Ð¤H¡G", intRow, 80
                    PrintDropLine "" & strCustName(0), "¥Ó½Ð¤H¡G", intRow, 80, Printer.TextWidth("¥DÃD¡G") - 400
                End If
                'Add by Morgan 2010/12/1
                m_iSubject = m_iSubject + 1
                SetWordArray m_Subject, m_iSubject, 2, "¥Ó½Ð¤H¡G" & strCustName(0)
            
                '¨ä¥Lªº¥Ó½Ð¤H
                For ii = 1 To 4
                   If strCustName(ii) <> "" Then
                      intRow = intRow + 1
                      If m_b2Printer Then
                         Printer.CurrentX = Printer.TextWidth("¥DÃD¡G") + intDefault + Printer.TextWidth("¥Ó½Ð¤H¡G")
                         Printer.CurrentY = 1500 + intRow * m_LineH + intTop
                         'Modified by Morgan 2022/8/4
                         'Printer.Print strCustName(ii)
                         PUB_PrintUnicodeText strCustName(ii), Printer.CurrentX, Printer.CurrentY, 0
                         'end 2022/8/4
                      End If
                      'Removed by Morgan 2022/8/4
                      'If m_b2Picture Then
                      '   Picture1.CurrentX = (Picture1.TextWidth("¥DÃD¡G") + Picture1.TextWidth("¥Ó½Ð¤H¡G")) + intDefault * douExtRate
                      '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
                      '   Picture1.Print strCustName(ii)
                      'End If
                      'end 2022/8/4
                      'Add by Morgan 2010/12/1
                      m_iSubject = m_iSubject + 1
                      SetWordArray m_Subject, m_iSubject, 2, "¡@¡@¡@¡@" & strCustName(ii)
                   End If
                Next ii
           End If
                intRow = intRow + 1
            'Modified by Lydia 2015/04/09 +§PÂ_«D¾ã§å½Ð´Ú³æ,¦C¦L©ïÀY¸ê®Æ
            If m_bolChiDB = False Then
                If m_b2Printer Then
                   Printer.CurrentX = Printer.TextWidth("¥DÃD¡G") + intDefault
                   Printer.CurrentY = 1500 + intRow * m_LineH + intTop
                   'Modified by Morgan 2022/8/4
                   'Printer.Print "°Ó¼Ð¡G" & strCaseName
                   PUB_PrintUnicodeText "°Ó¼Ð¡G" & strCaseName, Printer.CurrentX, Printer.CurrentY, 0
                   'end 2022/8/4
                End If
                'Removed by Morgan 2022/8/4
                'If m_b2Picture Then
                '   Picture1.CurrentX = Picture1.TextWidth("¥DÃD¡G") + (intDefault) * douExtRate
                '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
                '   Picture1.Print "°Ó¼Ð¡G" & strCaseName
                'End If
                'end 2022/8/4
                'Add by Morgan 2010/12/1
                m_iSubject = m_iSubject + 1
                SetWordArray m_Subject, m_iSubject, 2, "°Ó¼Ð¡G" & strCaseName
                      
                intRow = intRow + 1
                If m_b2Printer Then
                   Printer.CurrentX = Printer.TextWidth("¥DÃD¡G") + intDefault
                   Printer.CurrentY = 1500 + intRow * m_LineH + intTop
                   'Modified by Morgan 2022/8/4
                   'Printer.Print "Ãþ§O¡G²Ä " & strTM09 & " Ãþ"
                   PUB_PrintUnicodeText "Ãþ§O¡G²Ä " & strTM09 & " Ãþ", Printer.CurrentX, Printer.CurrentY, 0
                   'end 2022/8/4
                End If
                'Removed by Morgan 2022/8/4
                'If m_b2Picture Then
                '   Picture1.CurrentX = Picture1.TextWidth("¥DÃD¡G") + (intDefault) * douExtRate
                '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
                '   Picture1.Print "Ãþ§O¡G²Ä " & strTM09 & " Ãþ"
                'End If
                'end 2022/8/4
                'Add by Morgan 2010/12/1
                m_iSubject = m_iSubject + 1
                SetWordArray m_Subject, m_iSubject, 2, "Ãþ§O¡G²Ä " & strTM09 & " Ãþ"
                
                intRow = intRow + 1
                If m_b2Printer Then
                   Printer.CurrentX = Printer.TextWidth("¥DÃD¡G") + intDefault
                   Printer.CurrentY = 1500 + intRow * m_LineH + intTop
                End If
                'Removed by Morgan 2022/8/4
                'If m_b2Picture Then
                '   Picture1.CurrentX = Picture1.TextWidth("¥DÃD¡G") + (intDefault) * douExtRate
                '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
                'End If
                'end 2022/8/4
                If strPrintText <> "" Then
                   If m_b2Printer Then
                      'Modified by Morgan 2022/8/4
                      'Printer.Print strPrintText
                      PUB_PrintUnicodeText strPrintText, Printer.CurrentX, Printer.CurrentY, 0
                      'end 2022/8/4
                   End If
                   'Removed by Morgan 2022/8/4
                   'If m_b2Picture Then
                   '   Picture1.Print strPrintText
                   'End If
                   'end 2022/8/4
                   'Add by Morgan 2010/12/1
                   m_iSubject = m_iSubject + 1
                   SetWordArray m_Subject, m_iSubject, 2, strPrintText
                Else
                   If strConNo <> "" Then
                      m_tmp = "µù¥U¸¹¡G" & strConNo
                   Else
                      m_tmp = "¥Ó½Ð®×¸¹¡G" & strAppNo
                   End If
                   If m_b2Printer Then
                      'Modified by Morgan 2022/8/4
                      'Printer.Print m_tmp
                      PUB_PrintUnicodeText m_tmp, Printer.CurrentX, Printer.CurrentY, 0
                      'end 2022/8/4
                   End If
                   'Removed by Morgan 2022/8/4
                   'If m_b2Picture Then
                   '   Picture1.Print m_tmp
                   'End If
                   'end 2022/8/4
                   'Add by Morgan 2010/12/1
                   m_iSubject = m_iSubject + 1
                   SetWordArray m_Subject, m_iSubject, 2, m_tmp
                End If
            End If
            'end Modified by Lydia 2015/04/09
         Case "2"
            intRow = intRow + 1
            If strTM28 = "2" Or strTM28 = "3" Or strTM28 = "4" Then
               m_tmp = strPrintText & " " & """" & strCaseName & """"
               PutData m_tmp, intRow, 400
            Else
               m_tmp = "Mark: " & strCaseName
               'Modified by Morgan 2012/8/28 °Ó¼Ð¦WºÙ·|¶W¹L¤@¦æ,Ex.X10109181
               'PutData m_tmp, intRow, 400
               If CountLength(strCaseName) <= 90 Then
                  PutData m_tmp, intRow, 400
               Else
                  PrintDropLine strCaseName, "Mark: ", intRow, 90
               End If
               'end 2012/8/28
            End If
            'Add by Morgan 2010/11/25
            m_iSubject = m_iSubject + 1
            SetWordArray m_Subject, m_iSubject, 2, m_tmp
            
           If strPA08 <> "7" Then     'add by sonia 2016/12/23 «DÃÒ©ú¼Ð³¹¤~­n¦L°Ó«~Ãþ§O
               'Add By Sindy 2016/9/5
               intRow = intRow + 1
               'Modified by Morgan 2017/8/25 Class:->Class(es):
               m_tmp = "Class(es): " & strTM09
               If CountLength(strTM09) <= 90 Then
                  PutData m_tmp, intRow, 400
               Else
                  PrintDropLine strTM09, "Class(es): ", intRow, 90
               End If
               m_iSubject = m_iSubject + 1
               SetWordArray m_Subject, m_iSubject, 2, m_tmp
               '2016/9/5 END
            End If          'add by sonia 2016/12/23
            
            intRow = intRow + 1
            If m_b2Printer Then
               Printer.CurrentX = 400 + intDefault
               Printer.CurrentY = 1500 + intRow * m_LineH + intTop
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.CurrentX = (400 + intDefault) * douExtRate
            '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
            'End If
            'end 2022/8/4
            If CountLength(strCustName(0)) <= 80 Then
               If m_b2Printer Then
                  'Modified by Morgan 2022/8/4
                  'Printer.Print IIf(strTM28 = "", "In the name of ", "Applicant: ") & strCustName(0)
                  PUB_PrintUnicodeText IIf(strTM28 = "", "In the name of ", "Applicant: ") & strCustName(0), Printer.CurrentX, Printer.CurrentY, 0
                  'end 2022/8/4
               End If
               'Removed by Morgan 2022/8/4
               'If m_b2Picture Then
               '   Picture1.Print IIf(strTM28 = "", "In the name of ", "Applicant: ") & strCustName(0)
               'End If
               'end 2022/8/4
            Else
                PrintDropLine "" & strCustName(0), IIf(strTM28 = "", "In the name of ", "Applicant: "), intRow, 80
            End If
            
            'Add by Morgan 2010/11/25
            m_iSubject = m_iSubject + 1
            SetWordArray m_Subject, m_iSubject, 2, IIf(strTM28 = "", "In the name of ", "Applicant: ") & strCustName(0)
            '¨ä¥Lªº¥Ó½Ð¤H
            For ii = 1 To 4
               If strCustName(ii) <> "" Then
                  intRow = intRow + 1
                  If m_b2Printer Then
                     Printer.CurrentX = 400 + intDefault + Printer.TextWidth(IIf(strTM28 = "", "In the name of ", "Applicant: "))
                     Printer.CurrentY = 1500 + intRow * m_LineH + intTop
                     'Modified by Morgan 2022/8/4
                     'Printer.Print strCustName(ii)
                     PUB_PrintUnicodeText strCustName(ii), Printer.CurrentX, Printer.CurrentY, 0
                     'end 2022/8/4
                  End If
                  'Removed by Morgan 2022/8/4
                  'If m_b2Picture Then
                  '   Picture1.CurrentX = (400 + intDefault) * douExtRate + Picture1.TextWidth(IIf(strTM28 = "", "In the name of ", "Applicant: "))
                  '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
                  '   Picture1.Print strCustName(ii)
                  'End If
                  'end 2022/8/4
                  'Add by Morgan 2010/11/25
                  m_iSubject = m_iSubject + 1
                  SetWordArray m_Subject, m_iSubject, 3, strCustName(ii)
               End If
            Next ii
            
         Case "3"
            intRow = intRow + 1
            If m_b2Printer Then
               Printer.CurrentX = Printer.TextWidth("¥ó¦W¡G") + intDefault
               Printer.CurrentY = 1500 + intRow * m_LineH + intTop
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.CurrentX = Picture1.TextWidth("¥ó¦W¡G") + (intDefault) * douExtRate
            '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
            'End If
            'end 2022/8/4
            If CountLength(strCustName(0)) <= 80 Then
               If m_b2Printer Then
                  'Modified by Morgan 2022/8/4
                  'Printer.Print "¥XÄ@¤H¡G" & strCustName(0)
                  PUB_PrintUnicodeText "¥XÄ@¤H¡G" & strCustName(0), Printer.CurrentX, Printer.CurrentY, 0
                  'end 2022/8/4
               End If
               'Removed by Morgan 2022/8/4
               'If m_b2Picture Then
               '   Picture1.Print "¥XÄ@¤H¡G" & strCustName(0)
               'End If
               'end 2022/8/4
            Else
                'Modified by Morgan 2012/3/12
                'PrintDropLine "" & strCustName(0), "¥XÄ@¤H¡G", intRow, 80
                PrintDropLine "" & strCustName(0), "¥XÄ@¤H¡G", intRow, 80, Printer.TextWidth("¥ó¦W¡G") - 400
            End If
            
            'Add by Morgan 2010/12/1
            m_iSubject = m_iSubject + 1
            SetWordArray m_Subject, m_iSubject, 2, "¥XÄ@¤H¡G" & strCustName(0)
            
            '¨ä¥Lªº¥Ó½Ð¤H
            For ii = 1 To 4
               If strCustName(ii) <> "" Then
                  intRow = intRow + 1
                  If m_b2Printer Then
                     Printer.CurrentX = Printer.TextWidth("¥ó¦W¡G") + intDefault + Printer.TextWidth("¥XÄ@¤H¡G")
                     Printer.CurrentY = 1500 + intRow * m_LineH + intTop
                     'Modified by Morgan 2022/8/4
                     'Printer.Print strCustName(ii)
                     PUB_PrintUnicodeText strCustName(ii), Printer.CurrentX, Printer.CurrentY, 0
                     'end 2022/8/4
                  End If
                  'Removed by Morgan 2022/8/4
                  'If m_b2Picture Then
                  '   Picture1.CurrentX = (Picture1.TextWidth("¥ó¦W¡G") + Picture1.TextWidth("¥XÄ@¤H¡G")) + intDefault * douExtRate
                  '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
                  '   Picture1.Print strCustName(ii)
                  'End If
                  'end 2022/8/4
                  'Add by Morgan 2010/12/1
                  m_iSubject = m_iSubject + 1
                  SetWordArray m_Subject, m_iSubject, 2, "¡@¡@¡@¡@" & strCustName(ii)
               End If
            Next ii
        
            intRow = intRow + 1
            'Modified by Morgan 2022/8/4
            'strExc(1) = "°Ó¼Ð¦W†ï¡G" & strCaseName
            strExc(1) = "°Ó¼Ð¦W" & PUB_GetUniText(Me.Name, "ºÙ") & "¡G" & strCaseName
            'end 2022/8/4
            If m_b2Printer Then
               Printer.CurrentX = Printer.TextWidth("¥ó¦W¡G") + intDefault
               Printer.CurrentY = 1500 + intRow * m_LineH + intTop
               'Modified by Morgan 2022/8/4
               'Printer.Print strExc(1)
               PUB_PrintUnicodeText strExc(1), Printer.CurrentX, Printer.CurrentY, 0
               'end 2022/8/4
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.CurrentX = Picture1.TextWidth("¥ó¦W¡G") + (intDefault) * douExtRate
            '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
            '   Picture1.Print strExc(1)
            'End If
            'end 2022/8/4
            'Add by Morgan 2010/12/1
            m_iSubject = m_iSubject + 1
            SetWordArray m_Subject, m_iSubject, 2, strExc(1)
                  
            intRow = intRow + 1
            'Modify by Morgan 2009/6/17
            strExc(1) = Left(strTM09, 68)
            'Modified by Morgan 2022/8/4
            'strExc(1) = "°Ó«~üP¤À¡G" & strExc(1)
            strExc(1) = "°Ó«~" & PUB_GetUniText(Me.Name, "°Ï") & "¤À¡G" & strExc(1)
            'end 2022/8/4
            strTM09 = Mid(strTM09, 69)
            If m_b2Printer Then
               Printer.CurrentX = Printer.TextWidth("¥ó¦W¡G") + intDefault
               Printer.CurrentY = 1500 + intRow * m_LineH + intTop
               'Modified by Morgan 2022/8/4
               'Printer.Print strExc(1)
               PUB_PrintUnicodeText strExc(1), Printer.CurrentX, Printer.CurrentY, 0
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.CurrentX = Picture1.TextWidth("¥ó¦W¡G") + (intDefault) * douExtRate
            '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
            '   Picture1.Print strExc(1)
            'End If
            'end 2022/8/4
            'Add by Morgan 2010/12/1
            m_iSubject = m_iSubject + 1
            SetWordArray m_Subject, m_iSubject, 2, strExc(1)
            
            Do While strTM09 <> ""
               strExc(1) = Left(strTM09, 68)
               strTM09 = Mid(strTM09, 69)
               intRow = intRow + 1
               If m_b2Printer Then
                  Printer.CurrentX = Printer.TextWidth(String(8, "¡@")) + intDefault
                  Printer.CurrentY = 1500 + intRow * m_LineH + intTop
                  'Modified by Morgan 2022/8/4
                  'Printer.Print strExc(1)
                  PUB_PrintUnicodeText strExc(1), Printer.CurrentX, Printer.CurrentY, 0
                  'end 2022/8/4
               End If
               'Removed by Morgan 2022/8/4
               'If m_b2Picture Then
               '   Picture1.CurrentX = Picture1.TextWidth(String(8, "¡@")) + (intDefault) * douExtRate
               '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
               '   Picture1.Print strExc(1)
               'End If
               'end 2022/8/4
               'Add by Morgan 2010/12/1
               m_iSubject = m_iSubject + 1
               'Modified by Morgan 2019/4/12
               'SetWordArray m_Subject, m_iSubject, 2, "¡@¡@¡@¡@¡@¡@¡@¡@" & strExc(1)
               SetWordArray m_Subject, m_iSubject, 2, "¡@¡@¡@¡@¡@" & strExc(1)
            Loop
      End Select
      
      '¦hªÅ¤@¦æ
      intRow = intRow + 1
      
   'Add by Morgan 2006/11/17
   ElseIf m_CP01 = "TS" Then
      Select Case strLanguage
         Case "1"
            'Add by Morgan 2007/2/2 ¥[°Ó¼Ð¦WºÙ
            'Added by Lydia 2015/07/03  §PÂ_¾ã§å½Ð´Ú³æ¥u°O¿ý²Ä¤@µ§
            If m_bolChiDB = False Or (m_bolChiDB And adoacc1k0.AbsolutePosition = 1) Then
                intRow = intRow + 1
                m_tmp = "°Ó¼Ð¡G" & GetCaseName(m_CP01 & m_CP02 & m_CP03 & m_CP04, "1")
                If m_b2Printer Then
                   Printer.CurrentX = Printer.TextWidth("¥DÃD¡G") + intDefault
                   Printer.CurrentY = 1500 + intRow * m_LineH + intTop
                   'Modified by Morgan 2022/8/4
                   'Printer.Print m_tmp
                   PUB_PrintUnicodeText m_tmp, Printer.CurrentX, Printer.CurrentY, 0
                   'end 2022/8/4
                End If
                'Removed by Morgan 2022/8/4
                'If m_b2Picture Then
                '   Picture1.CurrentX = Picture1.TextWidth("¥DÃD¡G") + (intDefault) * douExtRate
                '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
                '   Picture1.Print m_tmp
                'End If
                'end 2022/8/4
                'Add by Morgan 2010/12/1
                m_iSubject = m_iSubject + 1
                SetWordArray m_Subject, m_iSubject, 2, m_tmp
            End If
            'end 2015/07/03
            
            'end 2007/2/2
            intRow = intRow + 1
            'Modified by Lydia 2015/07/03 +§PÂ_«D¾ã§å½Ð´Ú³æ,¦C¦L©ïÀY¸ê®Æ
            If m_bolChiDB = False Then
               '2007/7/24 modify by sonia µL²Õ¸s®É¦LÃþ§O
               'Printer.Print "Ãþ§O¡G²Ä " & strTM09 & " ²Õ¸s"
               If strTM32 <> "" Then
                  m_tmp = "Ãþ§O¡G²Ä " & strTM32 & " ²Õ¸s"
               Else
                  m_tmp = "Ãþ§O¡G²Ä " & strTM09 & " Ãþ"
               End If
               '2007/7/24 end
               If m_b2Printer Then
                  Printer.CurrentX = Printer.TextWidth("¥DÃD¡G") + intDefault
                  Printer.CurrentY = 1500 + intRow * m_LineH + intTop
                  'Modified by Morgan 2022/8/4
                  'Printer.Print m_tmp
                  PUB_PrintUnicodeText m_tmp, Printer.CurrentX, Printer.CurrentY, 0
                  'end 2022/8/4
               End If
               'Removed by Morgan 2022/8/4
               'If m_b2Picture Then
               '   Picture1.CurrentX = Picture1.TextWidth("¥DÃD¡G") + (intDefault) * douExtRate
               '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
               '   Picture1.Print m_tmp
               'End If
               'end 2022/8/4
               'Add by Morgan 2010/12/1
               m_iSubject = m_iSubject + 1
               SetWordArray m_Subject, m_iSubject, 2, m_tmp
               
               intRow = intRow + 1
            '2012/4/11 add by sonia
            End If
            'end 2015/07/03
         Case "2"
            intRow = intRow + 1
            m_tmp = " for " & GetCaseName(m_CP01 & m_CP02 & m_CP03 & m_CP04) & " in class " & strTM09
            If m_b2Printer Then
               Printer.CurrentX = Printer.TextWidth("Re¡G") + intDefault
               Printer.CurrentY = 1500 + intRow * m_LineH + intTop
               'Modified by Morgan 2022/8/4
               'Printer.Print m_tmp
               PUB_PrintUnicodeText m_tmp, Printer.CurrentX, Printer.CurrentY, 0
               'end 2022/8/4
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.CurrentX = Picture1.TextWidth("Re¡G") + (intDefault) * douExtRate
            '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
            '   Picture1.Print m_tmp
            'End If
            'end 2022/8/4
            'Add by Morgan 2010/12/1
            m_iSubject = m_iSubject + 1
            SetWordArray m_Subject, m_iSubject, 2, m_tmp
            intRow = intRow + 1
         '2012/4/11 end
      End Select
      
   '­Y«D°Ó¼Ð®×¥ó
   Else
      '­Y®×¥ó¨t²ÎÃþ§O«D"S"
      If m_CP01 <> "S" Then
         'Modify by Morgan 2006/9/21 FG ·|¨S¦³¥Ó½Ð¤H
         Select Case strLanguage
            Case "1"
            'Added by Lydia 2015/07/03  §PÂ_¾ã§å½Ð´Ú³æ¥u°O¿ý²Ä¤@µ§
            If m_bolChiDB = False Or (m_bolChiDB And adoacc1k0.AbsolutePosition = 1) Then
               intRow = intRow + 1
               m_tmp = "¥Ó½Ð¤H¡G" & strCustName(0)
               If strCustName(0) <> "" Then
                  If m_b2Printer Then
                     Printer.CurrentX = Printer.TextWidth("¥DÃD¡G") + intDefault
                     Printer.CurrentY = 1500 + intRow * m_LineH + intTop
                     'Modified by Morgan 2022/8/4
                     'Printer.Print m_tmp
                     PUB_PrintUnicodeText m_tmp, Printer.CurrentX, Printer.CurrentY, 0
                     'end 2022/8/4
                  End If
                  'Removed by Morgan 2022/8/4
                  'If m_b2Picture Then
                  '   Picture1.CurrentX = (Picture1.TextWidth("¥DÃD¡G") + intDefault) * douExtRate
                  '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
                  '   Picture1.Print m_tmp
                  'End If
                  'end 2022/8/4
               End If
               'Add by Morgan 2010/12/1
               m_iSubject = m_iSubject + 1
               SetWordArray m_Subject, m_iSubject, 2, m_tmp
            
               '¨ä¥Lªº¥Ó½Ð¤H
               For ii = 1 To 4
                  If strCustName(ii) <> "" Then
                     intRow = intRow + 1
                     If m_b2Printer Then
                        Printer.CurrentX = Printer.TextWidth("¥DÃD¡G") + intDefault + Printer.TextWidth("¥Ó½Ð¤H¡G")
                        Printer.CurrentY = 1500 + intRow * m_LineH + intTop
                        'Modified by Morgan 2022/8/4
                        'Printer.Print strCustName(ii)
                        PUB_PrintUnicodeText strCustName(ii), Printer.CurrentX, Printer.CurrentY, 0
                        'end 2022/8/4
                     End If
                     'Removed by Morgan 2022/8/4
                     'If m_b2Picture Then
                     '   Picture1.CurrentX = (Picture1.TextWidth("¥DÃD¡G") + Picture1.TextWidth("¥Ó½Ð¤H¡G")) + intDefault * douExtRate
                     '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
                     '   Picture1.Print strCustName(ii)
                     'End If
                     'end 2022/8/4
                     'Add by Morgan 2010/12/1
                     m_iSubject = m_iSubject + 1
                     SetWordArray m_Subject, m_iSubject, 2, "¡@¡@¡@¡@" & strCustName(ii)
               
                  End If
               Next ii
            End If
            'end 2015/07/03
            
               intRow = intRow + 1
            'Modified by Lydia 2015/07/03 +§PÂ_«D¾ã§å½Ð´Ú³æ,¦C¦L©ïÀY¸ê®Æ
            If m_bolChiDB = False Then
               If m_b2Printer Then
                  Printer.CurrentX = Printer.TextWidth("¥DÃD¡G") + intDefault
                  Printer.CurrentY = 1500 + intRow * m_LineH + intTop
                  'Modified by Morgan 2022/8/4
                  'Printer.Print "®×¥ó¦WºÙ¡G" & strCaseName
                  PUB_PrintUnicodeText "®×¥ó¦WºÙ¡G" & strCaseName, Printer.CurrentX, Printer.CurrentY, 0
                  'end 2022/8/4
               End If
               'Removed by Morgan 2022/8/4
               'If m_b2Picture Then
               '   Picture1.CurrentX = Picture1.TextWidth("¥DÃD¡G") + (intDefault) * douExtRate
               '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
               '   Picture1.Print "®×¥ó¦WºÙ¡G" & strCaseName
               'End If
               'end 2022/8/4
               'Add by Morgan 2010/12/1
               m_iSubject = m_iSubject + 1
               SetWordArray m_Subject, m_iSubject, 2, "®×¥ó¦WºÙ¡G" & strCaseName
            End If
            'end 2015/07/03
            
            Case "2"
               'Added by Morgan 2020/10/12 ½Ð´Ú¹ï¶HY5548300 Sagacious Advanced Research Center Inc. ­n±a®×¥ó¦WºÙ --·¨¬M·O
               If m_A1k28 = "Y55483000" And (m_CP01 = "P" Or m_CP01 = "FCP" Or m_CP01 = "CFP") Then
                  intRow = intRow + 1
                  m_iSubject = m_iSubject + 1
                  SetWordArray m_Subject, m_iSubject, 2, "Title: " & strCaseName
               End If
               'end 2020/10/12
               
               intRow = intRow + 1
               If strCustName(0) <> "" Then
                  'Modify by Morgan 2010/12/27 ¥[§é¦æ±±¨î
                  'PutData "Applicant: " & strCustName(0), intRow, 400
                  PrintDropLine strCustName(0), "Applicant: ", intRow, 80
                  
                  'Add by Morgan 2010/11/25
                  m_iSubject = m_iSubject + 1
                  SetWordArray m_Subject, m_iSubject, 2, "Applicant: " & strCustName(0)
                  
               End If
               '¨ä¥Lªº¥Ó½Ð¤H
               For ii = 1 To 4
                  If strCustName(ii) <> "" Then
                     intRow = intRow + 1
                     If m_b2Printer Then
                        Printer.CurrentX = 400 + intDefault + Printer.TextWidth("Applicant: ")
                        Printer.CurrentY = 1500 + intRow * m_LineH + intTop
                        'Modified by Morgan 2022/8/4
                        'Printer.Print strCustName(ii)
                        PUB_PrintUnicodeText strCustName(ii), Printer.CurrentX, Printer.CurrentY, 0
                        'end 2022/8/4
                     End If
                     'Removed by Morgan 2022/8/4
                     'If m_b2Picture Then
                     '   Picture1.CurrentX = (400 + intDefault) * douExtRate + Picture1.TextWidth("Applicant: ")
                     '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
                     '   Picture1.Print strCustName(ii)
                     'End If
                     'end 2022/8/4
                     'Add by Morgan 2010/11/25
                     m_iSubject = m_iSubject + 1
                     SetWordArray m_Subject, m_iSubject, 3, strCustName(ii)
                     
                  End If
               Next ii
               'Add by Morgan 2010/9/29
               'Modified by Morgan 2017/6/21 +Y52622000  --³¯¨Ø­s
               'Modified by Morgan 2017/7/7 +¦C¦L¹ï¶H=Y54443 -- ³¯¼W¼s
               'Modified by Morgan 2018/8/22 +¦C¦L¹ï¶H=Y54875 -- ¬x°ö³ó
               If m_A1k28 = "Y52146000" Or m_A1k28 = "Y52622000" Or m_A1k27 = "Y54443000" Or m_A1k27 = "Y54875000" Then
               'end 2017/6/21
                  intRow = intRow + 1
                  PrintDropLine strCaseName, "Title: ", intRow, 70
                  
                  'Add by Morgan 2010/11/25
                  m_iSubject = m_iSubject + 1
                  'Modified by Morgan 2017/6/21
                  SetWordArray m_Subject, m_iSubject, 2, "Title:"
                  SetWordArray m_Subject, m_iSubject, 3, strCaseName
                  'SetWordArray m_Subject, m_iSubject, 2, "Title:" & strCaseName
                  'end 2017/6/21
               End If
               
            Case "3"
               'Added by Morgan 2020/9/30 ¥N²z¤HY49456 Ogoshi ®×¥ó­n±a®×¥ó¦WºÙ --³¢©É¼ü
               If m_A1k03 = "Y49456000" And (m_CP01 = "P" Or m_CP01 = "FCP" Or m_CP01 = "CFP") Then
                  intRow = intRow + 1
                  m_iSubject = m_iSubject + 1
                  SetWordArray m_Subject, m_iSubject, 2, "¡u" & strCaseName & "¡v"
               End If
               'end 2020/9/30
               
               intRow = intRow + 1
               If strCustName(0) <> "" Then
                  If m_b2Printer Then
                     Printer.CurrentX = Printer.TextWidth("¥ó¦W¡G") + intDefault
                     Printer.CurrentY = 1500 + intRow * m_LineH + intTop
                     'Modified by Morgan 2022/8/4
                     'Printer.Print "¥XÄ@¤H¡G" & strCustName(0)
                     PUB_PrintUnicodeText "¥XÄ@¤H¡G" & strCustName(0), Printer.CurrentX, Printer.CurrentY, 0
                     'end 2022/8/4
                  End If
                  'Removed by Morgan 2022/8/4
                  'If m_b2Picture Then
                  '   Picture1.CurrentX = Picture1.TextWidth("¥ó¦W¡G") + (intDefault) * douExtRate
                  '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
                  '   Picture1.Print "¥XÄ@¤H¡G" & strCustName(0)
                  'End If
                  'end 2022/8/4
                  'Add by Morgan 2010/12/1
                  m_iSubject = m_iSubject + 1
                  SetWordArray m_Subject, m_iSubject, 2, "¥XÄ@¤H¡G" & strCustName(0)
               End If
               '¨ä¥Lªº¥Ó½Ð¤H
               For ii = 1 To 4
                  If strCustName(ii) <> "" Then
                     intRow = intRow + 1
                     If m_b2Printer Then
                        Printer.CurrentX = Printer.TextWidth("¥ó¦W¡G") + intDefault + Printer.TextWidth("¥XÄ@¤H¡G")
                        Printer.CurrentY = 1500 + intRow * m_LineH + intTop
                        'Modified by Morgan 2022/8/4
                        'Printer.Print strCustName(ii)
                        PUB_PrintUnicodeText strCustName(ii), Printer.CurrentX, Printer.CurrentY, 0
                        'end 2022/8/4
                     End If
                     'Removed by Morgan 2022/8/4
                     'If m_b2Picture Then
                     '   Picture1.CurrentX = (Picture1.TextWidth("¥ó¦W¡G") + Picture1.TextWidth("¥XÄ@¤H¡G")) + intDefault * douExtRate
                     '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
                     '   Picture1.Print strCustName(ii)
                     'End If
                     'end 2022/8/4
                     'Add by Morgan 2010/12/1
                     m_iSubject = m_iSubject + 1
                     SetWordArray m_Subject, m_iSubject, 2, "¡@¡@¡@¡@" & strCustName(ii)
                  
                  End If
               Next ii
         End Select
      End If
      '¦hªÅ¤@¦æ
      intRow = intRow + 1
   End If
   
   intRow = intRow + 1
   If bolNewForm = False Then 'Added by Morgan 2013/10/25
      Select Case strLanguage
         Case "2"
            PutData ReportSum(83), intRow, 0
            
            'Add by Morgan 2010/11/25
            m_iSubject = m_iSubject + 1
            SetWordArray m_Subject, m_iSubject, 1, ReportSum(83)
                     
            '¦hªÅ¤@¦æ
            intRow = intRow + 1
       End Select
    End If 'Added by Morgan 2013/10/25
    If intAddSpaceRow = 0 Then intRow = intRow + 1
    m_DetailTopStart = 1500 + intRow * m_LineH
   Exit Sub
   
ErrHnd:
   MsgBox Err.Description, vbCritical
'Added by Lydia 2015/04/09
Chi_JumpPHM:
   
   If adoquery.State <> adStateClosed Then adoquery.Close
End Sub

Private Sub PrintHead2()
   'Modify by Morgan 2006/11/17
   'If ("" & adoacc1k0.Fields("A1K13").Value = "T") Then
   If Left("" & adoacc1k0.Fields("A1K13").Value, 1) = "T" Then
      'Add By Sindy 2009/07/20
      If Left(m_A1k03, 6) = "Y52588" Then
         m_tmp = "Invoice ±b³æ"
      '2009/07/20 End
      'Added by Morgan 2018/7/12--®Û­^
      ElseIf Left(m_A1k03, 6) = "Y53839" Then
         m_tmp = "INVOICE"
      'end 2018/7/12
      Else
         m_tmp = "¦¬¡@¶O¡@³q¡@ª¾¡@³æ"
      End If
      
   'Added by Morgan 2016/8/5 --Kimi
   ElseIf m_A1k28 = "Y54391000" Then
      m_tmp = "¡@µo¡@¡@¡@¡@¡@²¼¡@"
   'end 2016/8/5
   Else
      m_tmp = "½Ð¡@¡@¡@´Ú¡@¡@¡@³æ"
   End If
   m_Title = m_tmp 'Add by Morgan 2010/12/1
   If m_b2Printer Then
      Printer.Font.Size = 18
      Printer.Font.Bold = True
      strExc(1) = Printer.TextHeight(m_tmp)
      Printer.CurrentX = (10000 - Printer.TextWidth(m_tmp)) / 2 + intDefault
      Printer.CurrentY = 1500 + intRow * m_LineH + intTop + 600 - strExc(1)
      'Modified by Morgan 2022/8/4
      'Printer.Print m_tmp
      PUB_PrintUnicodeText m_tmp, Printer.CurrentX, Printer.CurrentY, 0
      'end 2022/8/4
   End If
   'Removed by Morgan 2022/8/4
   'If m_b2Picture Then
   '   Picture1.Font.Size = 18 * douExtRate
   '   Picture1.Font.Bold = True
   '   strExc(1) = Picture1.TextHeight(m_tmp)
   '   Picture1.CurrentX = ((10000 - Picture1.TextWidth(m_tmp)) / 2 + intDefault) * douExtRate
   '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop + 600 - strExc(1)) * douExtRate
   '   Picture1.Print m_tmp
   'End If
   'end 2022/8/4
   m_tmp = "½s¸¹: " & adoacc1k0.Fields("a1k01").Value
   If m_b2Printer Then
      Printer.Font.Size = 12
      Printer.Font.Bold = True
      strExc(1) = Printer.TextHeight(m_tmp)
      Printer.CurrentX = 8000 + intDefault
      Printer.CurrentY = 1500 + intRow * m_LineH + intTop + 600 - strExc(1)
      'Modified by Morgan 2022/8/4
      'Printer.Print m_tmp
      PUB_PrintUnicodeText m_tmp, Printer.CurrentX, Printer.CurrentY, 0
      'end 2022/8/4
      Printer.Font.Bold = False
   End If
   'Removed by Morgan 2022/8/4
   'If m_b2Picture Then
   '   Picture1.Font.Size = 12 * douExtRate
   '   Picture1.Font.Bold = True
   '   strExc(1) = Picture1.TextHeight(m_tmp)
   '   Picture1.CurrentX = (8000 + intDefault) * douExtRate
   '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop + 600 - strExc(1)) * douExtRate
   '   Picture1.Print m_tmp
   '   Picture1.Font.Bold = False
   'End If
   'end 2022/8/4
End Sub

'*************************************************
' ¦X­p¦ì¸m
'
'*************************************************
Private Sub PrintSum()
   Dim StrSQLa As String
'   Dim rsA As New ADODB.Recordset
   Dim lngBoxX As Long, lngBoxY As Long 'Add by Morgan 2006/7/5
   Dim strCache As String 'Add by Morgan 2010/11/30
   Dim dblRMBRate As Double 'Add by Sindy 2013/1/28
   Dim arrTxt() As String 'Added by Morgan 2013/11/14
   Dim ii As Integer 'Added by Morgan 2013/11/14
   Dim a1k28Na01 As String 'Added by Lydia 2015/05/06
   
   Select Case strLanguage
      Case "1": m_tmp = "Á`¡@­p"
      Case "2": m_tmp = "TOTAL"
      Case "3": m_tmp = "¦X¡@­p"
   End Select
   
   'ºû«ù¥Ø«e¦æ¼Æ
   intCounter = intCounter
   
   If (intCounter + 3) > 24 Then
      If strNewPage <> MsgText(602) Then
         intCounter = 0
         If m_b2Printer Then
            'Modified by Morgan 2012/10/31
            'Printer.NewPage
            MyNewPage
         End If
         'Removed by Morgan 2022/8/4
         'If m_b2Picture Then
         '   SetPic m_iPages
         'End If
         'end 2022/8/4
         m_iPages = m_iPages + 1
      End If
   End If
   
   If m_b2Printer Then
      Printer.CurrentX = 5000 + intDefault
      Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
      'Modified by Morgan 2022/8/4
      'Printer.Print m_tmp
      PUB_PrintUnicodeText m_tmp, Printer.CurrentX, Printer.CurrentY, 0
      'end 2022/8/4
   End If
   'Removed by Morgan 2022/8/4
   'If m_b2Picture Then
   '   Picture1.CurrentX = (5000 + intDefault) * douExtRate
   '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
   '   Picture1.Print m_tmp
   'End If
   'end 2022/8/4
   SetWordArray m_Sum, 1, 1, m_tmp 'Add by Morgan 2010/11/30
   
   strAmount = PUB_ChgFormat("" & douAmount, True)
   
   'Modified by Morgan 2012/12/7
'   Select Case strCurr
'      Case "U"
'         m_tmp = "USD"
'         'Add by Morgan 2011/4/1 ¬üª÷¦X­p§ï§ì a1k08 (»P¶Ê´Ú­n¤@­P)
'         strAmount = PUB_ChgFormat(m_A1k08, True)
'
'      '2009/5/18 modify by sonia
'      'Case Else: m_tmp = "NTD"
'      Case "N"
'         m_tmp = "NTD"
'      Case Else
'         If m_DNCurr = "USD" Then
'            m_tmp = "NTD"
'         Else
'            m_tmp = m_DNCurr
'         End If
'      '2009/5/18 END
'   End Select
   Select Case m_iPrintCurrType
   Case 3, 4 '¯Â¥~¹ô,¥~¹ô+¬üª÷¦X­p
      m_tmp = m_DNCurr
      '***** Sindy 2013/4/11 ¦¹¬qµ{¦¡ºû«ù¤£°Ê,¦]¹ïAccFMPImputCurrStarDate¤W½u«eªºÂÂ¸ê®Æ¦³¨ä·N¸q *****
      '¬üª÷¦X­p§ï§ì a1k08(»P¶Ê´Ú­n¤@­P)--·sªº±b³æÀ³¸Ó¤£·|¦A¦³(¬üª÷¦X­p=©ú²Ó¬üª÷¥[Á`)
      If m_tmp = "USD" Then
         strAmount = PUB_ChgFormat(m_A1k08, True)
      End If
      '***** 2013/4/11 End *****
   Case Else
      m_tmp = "NTD"
   End Select
   'end 2012/12/7
   
   'Modified by Morgan 2013/1/3 ª÷ÃB³£¥[¦L.00
'   'Added by Morgan 2012/8/3 ¬üª÷¾ã¼Æ¤]­n¦L .00
'   If m_tmp = "USD" Then
'      strAmount = Format(strAmount, FDollar)
'   End If
'   'end 2012/8/3
   strAmount = Format(strAmount, FDollar)
   'end 2013/1/3
         
   If m_b2Printer Then
      Printer.CurrentX = 8000 + intDefault
      Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
      'Modified by Morgan 2022/8/4
      'Printer.Print m_tmp
      PUB_PrintUnicodeText m_tmp, Printer.CurrentX, Printer.CurrentY, 0
      'end 2022/8/4
      intLength = Printer.TextWidth(strAmount)
      Printer.CurrentX = 10000 + intDefault - intLength
      Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
      'Modified by Morgan 2022/8/4
      'Printer.Print strAmount
      PUB_PrintUnicodeText strAmount, Printer.CurrentX, Printer.CurrentY, 0
      'end 2022/8/4
   End If
   'Removed by Morgan 2022/8/4
   'If m_b2Picture Then
   '   Picture1.CurrentX = (8000 + intDefault) * douExtRate
   '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
   '   Picture1.Print m_tmp
   '   intLength = Picture1.TextWidth(strAmount)
   '   Picture1.CurrentX = (10000 + intDefault) * douExtRate - intLength
   '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
   '   Picture1.Print strAmount
   'End If
   'end 2022/8/4
   
   'Add by Morgan 2010/12/3
   'Modified by Lydia 2016/03/03 + m_bSpecial5
   'Modified by Morgan 2020/2/17 Y22457000,Y48048000,Y52322000 ¨ú®ø--Tim
   If (m_bSpecial1 And m_bDowN = False) Or m_bSpecial5 Then
   
      strAmount = Format(strAmount)
      'Modify by Morgan 2011/5/24 §é¦©§ï¥Î¥¼§é¦©½Ð´Úª÷ÃB´î½Ð´Ú©ú²Ó¥[Á`­pºâ¤~¤£·|¦³»~®t
      If m_dblNoDiscAmtTot <> strAmount Then
         
         m_dblDiscTot = m_dblNoDiscAmtTot - Val(strAmount)
         
         m_iItem = m_iItem + 1
         SetItemWordArray m_Item, m_iItem, 1, "Negotiated Discount"
         SetItemWordArray m_Item, m_iItem, 2, "A111"
         SetItemWordArray m_Item, m_iItem, 3, m_tmp
         'Modified by Morgan 2013/1/3 ª÷ÃB³£¥[¦L.00
         'SetItemWordArray m_Item, m_iItem, 4, PUB_ChgFormat(-1 * m_dblDiscTot, True)
         SetItemWordArray m_Item, m_iItem, 4, Format(-1 * m_dblDiscTot, FDollar)
         'end 2013/1/3
         SetItemWordArray m_Item, m_iItem, 5, Format(DBDATE(m_strA1K02), "####/##/##")
         
         'Added by Morgan 2018/9/18
         If m_bEBilling Then
               iUpper = iUpper + 1
               ReDim Preserve strLedes(m_iCols, iUpper)
               '1~8
               For intI = 1 To 8
                  strLedes(intI, iUpper) = strLedes(intI, 1)
               Next
               strLedes(9, iUpper) = iUpper
               strLedes(10, iUpper) = "IF"
               strLedes(11, iUpper) = "1"
               strLedes(12, iUpper) = Format(-1 * m_dblDiscTot)
               strLedes(13, iUpper) = strLedes(12, iUpper)
               strLedes(14, iUpper) = strLedes(1, iUpper)
               strLedes(15, iUpper) = "L720"
               strLedes(16, iUpper) = ""
               strLedes(17, iUpper) = "A111"
               strLedes(18, iUpper) = ""
               strLedes(19, iUpper) = "Negotiated Discount of " & GetDisc(strLedes(2, iUpper)) & "."
               strLedes(20, iUpper) = strLedes(20, iUpper - 1)
               strLedes(21, iUpper) = ""
               '22~
               For intI = 22 To m_iCols
                  strLedes(intI, iUpper) = strLedes(intI, 1)
               Next
         End If
         'end 2018/9/18
      End If
      strAmount = Format(strAmount, "#,###.00") 'Dow ¤p¼Æ«á¨â¦ì00¤]­n±a
      SetWordArray m_Sum, 1, 2, m_tmp
      SetWordArray m_Sum, 1, 3, strAmount
   Else
      SetWordArray m_Sum, 1, 2, m_tmp 'Add by Morgan 2010/11/30
      SetWordArray m_Sum, 1, 3, strAmount  'Add by Morgan 2010/11/30
   End If
   
   intCounter = intCounter + 1
   
   'Modified by Morgan 2012/12/7
   'Select Case strCurr
   '   Case "U"
   '   Case "N"
   '   Case Else
   Select Case m_iPrintCurrType
      'Modify by Sindy 2013/1/10
      Case 2 '¥x¹ô+¥~¹ô¦X­p
   'end 2012/12/7
         If Val(m_TotOffFees) > 0 Then
            strAmount = "( Total Official fees " & m_DNCurr & "  " & Format(m_TotOffFees, FDollar) & " )"
            If m_b2Printer Then
               Printer.CurrentX = 7300 + intDefault - Printer.TextWidth(strAmount)
               Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
               'Modified by Morgan 2022/8/4
               'Printer.Print strAmount
               PUB_PrintUnicodeText strAmount, Printer.CurrentX, Printer.CurrentY, 0
               'end 2022/8/4
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.CurrentX = (7300 + intDefault) * douExtRate - Picture1.TextWidth(strAmount)
            '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
            '   Picture1.Print strAmount
            'End If
            'end 2022/8/4
            SetWordArray m_Sum, 2, 1, strAmount
            m_TotOffFees = 0
         End If
      
         If m_b2Printer Then
            Printer.CurrentX = 8000 + intDefault
            Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
            'Modified by Morgan 2022/8/4
            'Printer.Print m_DNCurr
            PUB_PrintUnicodeText m_DNCurr, Printer.CurrentX, Printer.CurrentY, 0
            'end 2022/8/4
         End If
         'Removed by Morgan 2022/8/4
         'If m_b2Picture Then
         '   Picture1.CurrentX = (8000 + intDefault) * douExtRate
         '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
         '   Picture1.Print m_DNCurr
         'End If
         'end 2022/8/4
         SetWordArray m_Sum, 2, 2, m_DNCurr
         
         strAmount = m_A1k08
         If bolPTUSDCase Then
            strAmount = Int(Val(strAmount))
         Else
            strAmount = Val(strAmount)
         End If
         
         'Modified by Morgan 2013/1/3 ª÷ÃB³£­n¦L.00
         'strAmount = PUB_ChgFormat(strAmount, True)
         strAmount = Format(strAmount, FDollar)
         'end 2013/1/3

         If m_b2Printer Then
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 10000 + intDefault - intLength
            Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
            'Modified by Morgan 2022/8/4
            'Printer.Print strAmount
            PUB_PrintUnicodeText strAmount, Printer.CurrentX, Printer.CurrentY, 0
            'end 2022/8/4
         End If
         'Removed by Morgan 2022/8/4
         'If m_b2Picture Then
         '   intLength = Picture1.TextWidth(strAmount)
         '   Picture1.CurrentX = (10000 + intDefault) * douExtRate - intLength
         '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
         '   Picture1.Print strAmount
         'End If
         'end 2022/8/4
         SetWordArray m_Sum, 2, 3, strAmount
         
         intCounter = intCounter + 1
      Case 4 '¥~¹ô+¬üª÷¦X­p
   'end 2012/12/7

         If m_b2Printer Then
            Printer.CurrentX = 8000 + intDefault
            Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
            Printer.Print "USD"
         End If
         'Removed by Morgan 2022/8/4
         'If m_b2Picture Then
         '   Picture1.CurrentX = (8000 + intDefault) * douExtRate
         '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
         '   Picture1.Print "USD"
         'End If
         'end 2022/8/4
         SetWordArray m_Sum, 2, 2, "USD" 'Add by Morgan 2010/11/30

         'Modify By Sindy 2012/12/27 Mark ¦]§¡³æ¤@½Ð´Ú½s¸¹§@¦C¦L,ª½±µ¤Þ¥Î«e­±¤wÅª¨ú m_A1k08 ÅÜ¼Æ§Y¥i
'         StrSQLa = "Select Sum(A1K08) From ACC1K0 Where A1K01 In (" & Left(m_strA1K01, Len(m_strA1K01) - 1) & ") "
'         rsA.CursorLocation = adUseClient
'         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'         If rsA.RecordCount > 0 Then
'            'Modify by Morgan 2008/3/21 P,T ¥xÆW®×¨ú¾ã¼Æ
'            'strAmount = Val("" & rsA.Fields(0).Value)
'            If bolPTUSDCase Then
'               strAmount = Int(Val("" & rsA.Fields(0).Value))
'            Else
'               strAmount = Val("" & rsA.Fields(0).Value)
'            End If
'            'end 2008/3/21
'            strAmount = PUB_ChgFormat(strAmount, True)
'         Else
'             strAmount = "0"
'         End If
'         If rsA.State <> adStateClosed Then rsA.Close
'         Set rsA = Nothing
'         'End
         '2012/12/27 End
         
         'Modify by Morgan 2008/3/21 P,T ¥xÆW®×¨ú¾ã¼Æ
         'strAmount = Val("" & rsA.Fields(0).Value)
         'Add By Sindy 2012/12/27
         If m_DNCurr = "USD" Then
            strAmount = m_A1k08
         Else
            '¨Ì½Ð´Úª÷ÃB¤Î½Ð´Ú¹ô§O¹ï¬üª÷¶×²v¨Ó´«ºâ¬üª÷¨ú¾ã¼Æ
            'Add By Sindy 2013/4/10
            'strAmount = trunc(m_A1k08 * m_DUsdRate)
            strAmount = Trunc(strAmount * m_DUsdRate)
         End If
         '2012/12/27 End
         If bolPTUSDCase Then
            strAmount = Int(Val(strAmount))
         Else
            strAmount = Val(strAmount)
         End If
         'end 2008/3/21
         'Modified by Morgan 2013/1/3 ª÷ÃB³£­n¦L.00
         'strAmount = PUB_ChgFormat(strAmount, True)
         strAmount = Format(strAmount, FDollar)
         'end 2013/1/3
         
         If m_b2Printer Then
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 10000 + intDefault - intLength
            Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
            'Modified by Morgan 2022/8/4
            'Printer.Print strAmount
            PUB_PrintUnicodeText strAmount, Printer.CurrentX, Printer.CurrentY, 0
            'end 2022/8/4
         End If
         'Removed by Morgan 2022/8/4
         'If m_b2Picture Then
         '   intLength = Picture1.TextWidth(strAmount)
         '   Picture1.CurrentX = (10000 + intDefault) * douExtRate - intLength
         '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
         '   Picture1.Print strAmount
         'End If
         'end 2022/8/4
         SetWordArray m_Sum, 2, 3, strAmount  'Add by Morgan 2010/11/30
                  
         If (m_b2Printer Or m_b2Picture) Then UpdateA1K38 m_strDN, Format(strAmount) 'Added by Morgan 2021/1/14 §ó·s½Ð´Ú³æ¬üª÷Á`ÃB
         
         intCounter = intCounter + 1
   End Select
   
   If m_b2Printer Then
      Printer.CurrentX = 8000 + intDefault
      Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
      Printer.Print String(IIf(strLanguage = "2", 18, 17), "v")
   End If
   'Removed by Morgan 2022/8/4
   'If m_b2Picture Then
   '   Picture1.CurrentX = (8000 + intDefault) * douExtRate
   '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
   '   Picture1.Print String(IIf(strLanguage = "2", 18, 17), "v")
   'End If
   'end 2022/8/4
   intCounter = intCounter + 1
   'Modify by Morgan 2010/12/24
   'If (intCounter + 6) > 24 Then
   If (intCounter + 6) > 27 Then
      If strNewPage <> MsgText(602) Then
         intCounter = 0
         If m_b2Printer Then
            'Modified by Morgan 2012/10/31
            'Printer.NewPage
            MyNewPage
         End If
         'Removed by Morgan 2022/8/4
         'If m_b2Picture Then
         '   SetPic m_iPages
         'End If
         'end 2022/8/4
         m_iPages = m_iPages + 1
      End If
   End If
   
   'Add By Sindy 2013/1/28
   strFMPFee99RMB = ""
   strA1L16 = ""
   If bolIsFMP Then
      strExc(0) = "select a1L16,sum(a1L17) as a1L17,sum(a1L18) as a1L18 from acc1L0 where a1L01=" & Left(m_strA1K01, Len(m_strA1K01) - 1) & " and substr(a1L04,-2)='99' and (a1L16='RMB' or (a1L18 is not null and a1L18>0)) group by a1L01,a1L16 order by a1L01,a1L16 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.RecordCount > 0 Then
            strA1L16 = "" & RsTemp.Fields("a1L16")
            If Val("" & RsTemp.Fields("a1L18")) > 0 Then
               strFMPFee99RMB = " ( RMB " & PUB_ChgFormat(RsTemp.Fields("a1L18"), True) & " ) "
            Else
               strFMPFee99RMB = " ( RMB " & PUB_ChgFormat(RsTemp.Fields("a1L17"), True) & " ) "
            End If
         End If
      End If
   End If
   '2013/1/28 End
   
    '§PÂ_©w½Z»y¤å
   Select Case strLanguage
      'Add by Morgan 2006/9/27
      Case "1" '¤¤
         If intCounter >= 21 Then
            strNewPage = MsgText(602)
            intCounter = 0
            If m_b2Printer Then
               'Modified by Morgan 2012/10/31
               'Printer.NewPage
               MyNewPage
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   SetPic m_iPages
            'End If
            'end 2022/8/4
            m_iPages = m_iPages + 1
            m_DetailTopStart = 1500 + 4 * m_LineH
         End If
         
         'Modify by Morgan 2006/11/17
         'If m_CP01 = "T" Then
         'Modify by Morgan 2008/3/19¡@¥x»È±b¸¹²Î¤@
         'If Left(m_CP01, 1) = "T" Then
            '»È¦æ¦WºÙ(­^)
            intCounter = intCounter + 3
            
            'Added by Morgan 2024/7/8
            If m_A1k03 = "Y56042000" And Left(strCust1, 8) = "X5655900" Then
               m_Title = "½Ð¡@´Ú¡@³æ"
            'end 2024/7/8
            
            'Modify by Morgan 2011/3/25 P§ï»PT¦P
            'If Left(m_CP01, 1) <> "T" Then
            ElseIf Left(m_CP01, 1) <> "T" And Left(m_CP01, 1) <> "P" Then
               m_tmp = "¥I´Ú®É½Ðµù©ú½Ð´Ú³æ½s¸¹"
               PutData m_tmp, intCounter, 0, m_DetailTopStart
               strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/12/1
               
               intCounter = intCounter + 1
               m_tmp = "¥»©Ò¤§»È¦æ±b¸¹¬°¡G"
               PutData m_tmp, intCounter, 0, m_DetailTopStart
               strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/12/1
               
               intCounter = intCounter + 1
            End If
            
            'Add by Morgan 2008/5/29 «È¤á¬°"·s¤é¿³(Ä¬¦{)"®É±b¤á¥Î"¤Ñ¬z¤T¤¸"
            '2011/12/9 modify by sonia §ï¤W®ü°ß·½
            If m_A1k28 = "X14843050" Then
               intCounter = intCounter + 1
               m_tmp = "»È¦æ:«Ø³]»È¦æ±×¤g¸ô¤ä¦æ"
               PutData m_tmp, intCounter, 0, m_DetailTopStart
               strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/12/1
               
               intCounter = intCounter + 1
               m_tmp = "±b¤á¦WºÙ:¤W®ü°ß·½±M§Q¥N²z¦³­­¤½¥q"
               PutData m_tmp, intCounter, 0, m_DetailTopStart
               strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/12/1
               
               intCounter = intCounter + 1
               m_tmp = "±b¸¹¡G31001554600050002379"
               PutData m_tmp, intCounter, 0, m_DetailTopStart
               strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/12/1
            
            'Added by Morgan 2024/7/9
            ElseIf m_A1k03 = "Y56042000" And Left(strCust1, 8) = "X5655900" Then
               strCache = vbCrLf
               strCache = strCache & "¡i±b¤á¸ê°T¡j"
               strCache = strCache & vbCrLf & "±b¤á¦WºÙ¡G¥x¤@°ê»Ú´¼¼z°]²£¨Æ°È©Ò"
               strCache = strCache & vbCrLf & "±b¸¹¡G003001305688"
               strCache = strCache & vbCrLf & "»È¦æ¦WºÙ¡G»OÆW»È¦æ"
               strCache = strCache & vbCrLf & "»È¦æ¥N½X¡G004"
               strCache = strCache & vbCrLf & "¤À¦æ¦WºÙ¤Î¤À¦æ¥N½X¡G¥x»ÈÀç·~³¡"
               strCache = strCache & vbCrLf
               strCache = strCache & vbCrLf
               strCache = strCache & "¡i½Ð´Ú¤H¸ê°T¡j"
               strCache = strCache & vbCrLf & "¥x¤@°ê»Ú´¼¼z°]²£¨Æ°È©Ò"
               strCache = strCache & vbCrLf & "¥x¥_¥«ªø¦wªF¸ô2¬q112¸¹9¼Ó"
               strCache = strCache & vbCrLf & "¹q¸Ü¡G(02)25061023¤À¾÷428"
               strCache = strCache & vbCrLf & "Email: ipdept@taie.com.tw"
            
            Else
               
               m_tmp = ReportSum(71001)
               PutData m_tmp, intCounter, 0, m_DetailTopStart
               strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/12/1
               
               '»È¦æ¦a§}(­^)
               intCounter = intCounter + 1
               m_tmp = ReportSum(72)
               'Added by Morgan 2023/8/14 Y53195¤§©Ò¦³®×¥ó½Ð´Ú³æªº¦a§}¡A¨ú®øR. O. C. --§d°û²ñ,³¢¶®®S,¾G¤Ñ¶³
               If Left(m_A1k03, 8) = "Y5319500" Then
                  m_tmp = Replace(m_tmp, ", R.O.C.", "")
               End If
               'end 2023/8/14
               PutData m_tmp, intCounter, 0, m_DetailTopStart
               strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/12/1
               
               'S.W.I.F.T. Address
               intCounter = intCounter + 1
               m_tmp = ReportSum(73001)
               PutData m_tmp, intCounter, 0, m_DetailTopStart
               strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/12/1
               
               '±b¤á¦WºÙ(­^)
               intCounter = intCounter + 1
               m_tmp = ReportSum(85)
               PutData m_tmp, intCounter, 0, m_DetailTopStart
               strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/12/1
               
               '±b¸¹
               intCounter = intCounter + 1
               m_tmp = ReportSum(74)
               PutData m_tmp, intCounter, 0, m_DetailTopStart
               strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/12/1

               'Add by Morgan 2008/3/21--³¢¶®®S
               '¬üª÷¶×²v
               'Modify by Morgan 2008/4/8 °Ó¼Ð¤£¥Î¦L--¸­¤jX09703287
'Remove by Morgan 2011/8/3 ¤£¥²¦A¦L--³¢¶®®S
'               If Left(m_CP01, 1) = "P" Then
'                  intCounter = intCounter + 1
'                  '2009/6/10 MODIFY BY SONIA
'                  'm_tmp = ReportSum(75) & m_strOriExcRate
'                  If m_DNCurr <> "USD" Then
'                     m_tmp = "Currency Rate: " & m_DNCurr & "1.00=USD" & Format(Val("0" & m_strOriExcRate))
'                  Else
'                     m_tmp = ReportSum(75) & Format(Val("0" & m_strOriExcRate), FDollar)
'                  End If
'                  '2009/6/10 END
'                  PutData m_tmp, intCounter, 0, m_DetailTopStart
'                  strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/12/1
'               End If
'end 2011/8/3
               'end 2008/3/21

            End If
         'End If
         
         'Added by Lydia 2015/07/07 RMB+¬üª÷­nÅã¥Ü¶×²v
         If m_iPrintCurrType = 4 And m_DNCurr = "RMB" Then
            intCounter = intCounter + 1
            m_tmp = "Currency Rate: " & m_DNCurr & "1.00=USD" & Format(Val("0" & m_DUsdRate))
            PutData m_tmp, intCounter, 0, m_DetailTopStart
            strCache = strCache & vbCrLf & m_tmp
         End If
         
         'Modify by Morgan 2011/3/25 P®×¤]­n¦L¤H¥Á¹ô±b¤á
         'If Left(m_CP01, 1) = "T" Then
         'Modified by Lydia 2015/05/06 §ï¦¨½Ð´Ú¹ï¶H°êÄy¬°¤j³°ªÌ
         'If Left(m_CP01, 1) = "T" Or Left(m_CP01, 1) = "P" Then
         If Left(m_A1k28, 1) = "Y" Then
            a1k28Na01 = GetPrjNationNumber(m_A1k28) '¥N²z¤H°êÄy
         ElseIf Left(m_A1k28, 1) = "X" Then
            a1k28Na01 = GetPrjNationNumber1(m_A1k28) '¥Ó½Ð¤H°êÄy
         Else
            a1k28Na01 = ""
         End If
         
         'Modified by Lydia 2015/10/05 X74358000­»´ä«È¤á­n¦L¤U¦C¸ê®Æ
         'If a1k28Na01 = "020" Then
         'Modified by Morgan 2017/8/8 +X71814000 --®Û­^
         'Modified by Morgan 2018/6/25 +Y54478000 --®Û­^
         If a1k28Na01 = "020" Or m_A1k28 = "X74358000" Or m_A1k28 = "X71814000" Or m_A1k28 = "Y54478000" Then
            'Removed by Morgan 2018/7/12 ¨ú®ø--°û²ñ:¤jÃB¤ä¥I¦æ¸¹¬O¥»©Ò¶×¬üª÷¨ì¤j³°©Ò»Ý­nªº¸ê®Æ
            ''¤jÃB¦æ¸¹   2014/4/7 ADD BY SONIA
            'intCounter = intCounter + 1
            ''¤jÃB¦æ¸¹:104290000362
            ''Modified by Lydia 2015/06/08 ¤jÃB¦æ¸¹:104290000362 ¦A§ï¬° ¤jÃB¤ä¥I¦æ¸¹ 104290000379
            ''m_tmp = "¤jÃB¦æ¸¹:104290000362"
            'm_tmp = "¤jÃB¤ä¥I¦æ¸¹:104290000379"
            'PutData m_tmp, intCounter, 0, m_DetailTopStart
            'strCache = strCache & vbCrLf & m_tmp
            ''2014/4/7 END
            'end 2018/7/12
            
            'Added by Morgan 2016/8/11
            'Y54391°¢¦è¦³¦â¤Ñ§»·ç¬ìÖº§÷®Æ¦³­­³d¥ô¤½¥q ±b³æ»Ý¥X²{¥x¹ô¬üª÷¶×²v--ªô¤l·ì
            If m_A1k28 = "Y54391000" Then
               intCounter = intCounter + 1
               m_tmp = "¥I´Ú¶×²v: USD1.00=NTD" & PUB_GetUSXRate_1(m_strA1K02, "USD")
               PutData m_tmp, intCounter, 0, m_DetailTopStart
               strCache = strCache & vbCrLf & m_tmp
               
            'Added by Morgan 2022/6/30 ¥¼½T»{--Ãý¥à
'            ElseIf m_A1k28 = "Y52754000" Then
'               intCounter = intCounter + 1
'               m_tmp = "Currency Rate: " & m_DNCurr & "1.00=NTD" & Format(Val("0" & m_DNRate))
'               PutData m_tmp, intCounter, 0, m_DetailTopStart
'               strCache = strCache & vbCrLf & m_tmp
            End If
            'end 2016/8/11
            
            'Modify by Morgan 2011/4/14 ¦ô­pÀ³¸ÓÁÙ¥i¥[2¦æ(´î¤Ö¸õ­¶¾÷²v)
            'If intCounter >= 21 Then
            If intCounter >= 23 Then
               strNewPage = MsgText(602)
               intCounter = 0
               If m_b2Printer Then
                  'Modified by Morgan 2012/10/31
                  'Printer.NewPage
                  MyNewPage
               End If
               'Removed by Morgan 2022/8/4
               'If m_b2Picture Then
               '   SetPic m_iPages
               'End If
               'end 2022/8/4
               m_iPages = m_iPages + 1
               
               m_DetailTopStart = 1500 + 4 * m_LineH
            End If
            'Modify by Morgan 2006/10/14
            '¼sªF¬Ù°Ó¼Ð¨Æ°È©Ò(Y46505)
            
'2009/8/4 CANCEL BY SONIA
'            If Left(m_A1k03, 6) = "Y46505" Then
'               '³æ¦ì¦WºÙ
'               intCounter = intCounter + 2
'               If m_b2Printer Then
'                  Printer.CurrentX = 0 + intDefault
'                  Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
'                  Printer.Print "³æ¦ì¦WºÙ¡G¥_¨Ê¥¨¨Êª¾ÃÑ²£Åv¥N²z¦³­­¤½¥q"
'               End If
'               If m_b2Picture Then
'                  Picture1.CurrentX = (0 + intDefault) * douExtRate
'                  Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
'                  Picture1.Print "³æ¦ì¦WºÙ¡G¥_¨Ê¥¨¨Êª¾ÃÑ²£Åv¥N²z¦³­­¤½¥q"
'               End If
'               '¶}¤á»È¦æ
'               intCounter = intCounter + 1
'               If m_b2Printer Then
'                  Printer.CurrentX = 0 + intDefault
'                  Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
'                  Printer.Print "¶}¤á»È¦æ¡G¤¤«H¹ê·~»È¦æÁ`¦æÀç·~³¡"
'               End If
'               If m_b2Picture Then
'                  Picture1.CurrentX = (0 + intDefault) * douExtRate
'                  Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
'                  Picture1.Print "¶}¤á»È¦æ¡G¤¤«H¹ê·~»È¦æÁ`¦æÀç·~³¡"
'               End If
'               '¦æ¸¹
'               intCounter = intCounter + 1
'               If m_b2Printer Then
'                  Printer.CurrentX = 0 + intDefault
'                  Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
'                  Printer.Print "¦æ¡@¡@¸¹¡G470"
'               End If
'               If m_b2Picture Then
'                  Picture1.CurrentX = (0 + intDefault) * douExtRate
'                  Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
'                  Picture1.Print "¦æ¡@¡@¸¹¡G470"
'               End If
'               '±b¸¹
'               intCounter = intCounter + 1
'               If m_b2Printer Then
'                  Printer.CurrentX = 0 + intDefault
'                  Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
'                  Printer.Print "±b¡@¡@¸¹¡G7111010182400048333"
'               End If
'               If m_b2Picture Then
'                  Picture1.CurrentX = (0 + intDefault) * douExtRate
'                  Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
'                  Picture1.Print "±b¡@¡@¸¹¡G7111010182400048333"
'               End If
'
'               intCounter = intCounter + 2
'               'ª`·N¨Æ¶µ
'               m_tmp = "¡°¡@¶Q¤½¥q¥i±N´Ú¶µ¶×¦Ü¥_¨Ê©Î¥xÆW¤§»È¦æ½ã¤á¡A"
'               If m_b2Printer Then
'                  Printer.FontBold = True
'                  Printer.CurrentX = 0 + intDefault
'                  Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
'                  Printer.Print m_tmp & "±©¤_¶×´Ú¦Z½Ð°È¥²±N¶×´Ú¾ÌÃÒ¶Ç¯u¦Ü¥x¥_"
'               End If
'               If m_b2Picture Then
'                  Picture1.FontBold = True
'                  Picture1.CurrentX = (0 + intDefault) * douExtRate
'                  Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
'                  Picture1.Print m_tmp & "±©¤_¶×´Ú¦Z½Ð°È¥²±N¶×´Ú¾ÌÃÒ¶Ç¯u¦Ü¥x¥_"
'               End If
'            Else
'2009/8/4 END

            If m_CP01 <> "FCP" Then 'Added by Morgan 2015/9/2 FCP¤£¦L--Kimi
            
               'Added by Morgan 2020/5/8 ³¢¶®®S«È¤áªºP®×¤£­n¦L¤j³°±b¤á--¬Â¬Â
               strExc(1) = ""
               'If m_CP01 = "P" Then   'cancel by sonia 2021/1/11 X11000333¬°³¢¶®®S¦¬¤å¤§CFP®×CFP-032167
                  strExc(0) = GetCuSales(strCust1, strExc(1))
               'End If                 'cancel by sonia 2021/1/11 X11000333¬°³¢¶®®S¦¬¤å¤§CFP®×CFP-032167
               'Modify By Sindy 2021/3/3 °Ó¼Ð½Ð´Ú³æ¤§¥»©Ò±b¤á:1.¥u¦L¥x»È±b¤á
               If strExc(1) = "79075" Or m_strFA126 = "1" Then
                  strCache = strCache & vbCrLf
                  SetWordArray m_Footer, 2, 1, "¶Q¤½¥q¤_¶×´Ú¦Z½Ð°È¥²±N¶×´Ú¾ÌÃÒ¶Ç¯u¦Ü¥x¥_©Ò¡A§_«h¥»©ÒµLªkª¾±x¡@¶Q¤½¥q¤w¶×´Ú¡C(¶Ç¯u¸¹½X¡G886 2 25011666)"
                  
                  intCounter = intCounter + 2
                  m_tmp = "¡°¡@¶Q¤½¥q¤_¶×´Ú¦Z½Ð°È¥²±N¶×´Ú¾ÌÃÒ¶Ç¯u¦Ü¥x¥_©Ò¡A§_«h¥»©ÒµLªkª¾±x¡@¶Q¤½¥q¤w¶×´Ú¡C"
                  PutData m_tmp, intCounter, 0, m_DetailTopStart
                  intCounter = intCounter + 1
                  m_tmp = "¡@¡@(¶Ç¯u¸¹½X¡G886 2 25011666)"
                  PutData m_tmp, intCounter, 0, m_DetailTopStart
               Else
               'end 2020/5/8
         
               
                  '2009/8/4§ï»È¦æ±b¤á,­ì¬°¤¤°ê¤u°Ó»È¦æ¤W®üªF¦w¸ô¤ä¦æ,ªL®Ê³¹,1001239101213902786*
                  '»È¦æ
                  intCounter = intCounter + 2
                  m_tmp = "»È¦æ¡G©Û°Ó»È¦æ¥_¨Ê¤À¦æª÷¿Äµó¤ä¦æ"
   
                  PutData m_tmp, intCounter, 0, m_DetailTopStart
                  strCache = strCache & vbCrLf & vbCrLf & m_tmp 'Add by Morgan 2010/12/1
                  
                  '±b¤á
                  intCounter = intCounter + 1
                  'modify by sonia 2021/1/8 ªL®Ê³¹§ïªL´º­§  2021/2/3¦A§ï¦^ªL®Ê³¹
                  'Modify By Sindy 2021/3/3 °Ó¼Ð½Ð´Ú³æ¤§¥»©Ò±b¤á:2.¥x»È±b¤á+¸³¨Æªø 3.¥x»È±b¤á+Á`¸g²z
                  If m_strFA126 = "3" Then
                     m_tmp = "½ã¤á¦WºÙ¡GªL´º­§¡]¤H¥Á¹ô­Ó¤H½ã¤á¡^"
                  Else
                  '2021/3/3 END
                     m_tmp = "½ã¤á¦WºÙ¡GªL®Ê³¹¡]¤H¥Á¹ô­Ó¤H½ã¤á¡^"
                  End If
                  PutData m_tmp, intCounter, 0, m_DetailTopStart
                  strCache = strCache & vbCrLf & m_tmp  'Add by Morgan 2010/12/1
                  
                  '±b¸¹
                  intCounter = intCounter + 1
                  'modify by sonia 2018/7/17¬¡¦s800100603817111§ïª÷¥d6226 0901 0488 1723
                  'm_tmp = "½ã¸¹¡G800100603817111"
                  'modify by sonia 2021/1/8 ªL®Ê³¹6226 0901 0488 1723§ïªL´º­§6214 8601 0005 4796    2021/2/3¦A§ï¦^ªL®Ê³¹
                  'Modify By Sindy 2021/3/3 °Ó¼Ð½Ð´Ú³æ¤§¥»©Ò±b¤á:2.¥x»È±b¤á+¸³¨Æªø 3.¥x»È±b¤á+Á`¸g²z
                  If m_strFA126 = "3" Then
                     m_tmp = "½ã¸¹¡G6214 8601 0005 4796"
                  Else
                  '2021/3/3 END
                     m_tmp = "½ã¸¹¡G6226 0901 0488 1723"
                  End If
                  PutData m_tmp, intCounter, 0, m_DetailTopStart
                  strCache = strCache & vbCrLf & m_tmp & vbCrLf  'Add by Morgan 2010/12/1
    
                  'ª`·N¨Æ¶µ
                  intCounter = intCounter + 2
                  
                  'Add by Morgan 2010/12/1
                  '³Æµù¤º®e­Y¦³ÅÜ°Ê»Ý¦P¨B­×§ï runWordChinese
                  SetWordArray m_Footer, 2, 1, "¶Q¤½¥q¥i±N´Ú¶µ¶×¦Ü¥»©Ò¥_¨Ê©Î¥xÆW¤§»È¦æ½ã¤á¡A±©¤_¶×´Ú¦Z½Ð°È¥²±N¶×´Ú¾ÌÃÒ¶Ç¯u¦Ü¥x¥_©Ò¡A§_«h¥»©ÒµLªkª¾±x¡@¶Q¤½¥q¤w¶×´Ú¡C(¶Ç¯u¸¹½X¡G886 2 25011666)"
                  'end 2010/12/1

                  m_tmp = "¡°¡@¶Q¤½¥q¥i±N´Ú¶µ¶×¦Ü¥»©Ò¥_¨Ê©Î¥xÆW¤§»È¦æ½ã¤á¡A"
                  strExc(1) = "±©¤_¶×´Ú¦Z½Ð°È¥²±N¶×´Ú¾ÌÃÒ¶Ç¯u¦Ü¥x¥_"
                  If m_b2Printer Then
                     Printer.FontBold = True
                     Printer.CurrentX = 0 + intDefault
                     Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
                     'Modified by Morgan 2022/8/4
                     'Printer.Print m_tmp & strExc(1)
                     PUB_PrintUnicodeText m_tmp & strExc(1), Printer.CurrentX, Printer.CurrentY, 0
                     'end 2022/8/4
                  End If
                  'Removed by Morgan 2022/8/4
                  'If m_b2Picture Then
                  '   Picture1.FontBold = True
                  '   Picture1.CurrentX = (0 + intDefault) * douExtRate
                  '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
                  '   Picture1.Print m_tmp & strExc(1)
                  'End If
                  'end 2022/8/4
   '2009/8/4CANCEL End If
                  If m_b2Printer Then
                     Printer.Line (0 + intDefault + Printer.TextWidth(m_tmp), m_DetailTopStart + intCounter * m_LineH + intTop + Printer.TextHeight("¡@") + 10)-(0 + intDefault + Printer.TextWidth(m_tmp) + Printer.TextWidth(strExc(1)), m_DetailTopStart + intCounter * m_LineH + intTop + Printer.TextHeight("¡@") + 10)
                     Printer.Line (0 + intDefault + Printer.TextWidth(m_tmp), m_DetailTopStart + intCounter * m_LineH + intTop + Printer.TextHeight("¡@") + 40)-(0 + intDefault + Printer.TextWidth(m_tmp) + Printer.TextWidth(strExc(1)), m_DetailTopStart + intCounter * m_LineH + intTop + Printer.TextHeight("¡@") + 40)
                  End If
                  'Removed by Morgan 2022/8/4
                  'If m_b2Picture Then
                  '   Picture1.Line (douExtRate * (0 + intDefault) + Picture1.TextWidth(m_tmp), douExtRate * (m_DetailTopStart + intCounter * m_LineH + intTop) + Picture1.TextHeight("¡@") + 10)-(douExtRate * (0 + intDefault) + Picture1.TextWidth(m_tmp) + Picture1.TextWidth(strExc(1)), douExtRate * (m_DetailTopStart + intCounter * m_LineH + intTop) + Picture1.TextHeight("¡@") + 10)
                  '   Picture1.Line (douExtRate * (0 + intDefault) + Picture1.TextWidth(m_tmp), douExtRate * (m_DetailTopStart + intCounter * m_LineH + intTop) + Picture1.TextHeight("¡@") + 40)-(douExtRate * (0 + intDefault) + Picture1.TextWidth(m_tmp) + Picture1.TextWidth(strExc(1)), douExtRate * (m_DetailTopStart + intCounter * m_LineH + intTop) + Picture1.TextHeight("¡@") + 40)
                  'End If
                  'end 2022/8/4
                  intCounter = intCounter + 1

                  strExc(1) = "©Ò¡A§_«h¥»©ÒµLªkª¾±x¡@¶Q¤½¥q¤w¶×´Ú"
                  If m_b2Printer Then
                     Printer.CurrentX = 0 + intDefault + Printer.TextWidth("¡@¡@")
                     Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
                     'Modified by Morgan 2022/8/4
                     'Printer.Print strExc(1) & "¡C(¶Ç¯u¸¹½X¡G886 2 25011666)"
                     PUB_PrintUnicodeText strExc(1) & "¡C(¶Ç¯u¸¹½X¡G886 2 25011666)", Printer.CurrentX, Printer.CurrentY, 0
                     'end 2022/8/4
                     Printer.Line (0 + intDefault + Printer.TextWidth("¡@¡@"), m_DetailTopStart + intCounter * m_LineH + intTop + Printer.TextHeight("¡@") + 10)-(0 + intDefault + Printer.TextWidth("¡@¡@") + Printer.TextWidth(strExc(1)), m_DetailTopStart + intCounter * m_LineH + intTop + Printer.TextHeight("¡@") + 10)
                     Printer.Line (0 + intDefault + Printer.TextWidth("¡@¡@"), m_DetailTopStart + intCounter * m_LineH + intTop + Printer.TextHeight("¡@") + 40)-(0 + intDefault + Printer.TextWidth("¡@¡@") + Printer.TextWidth(strExc(1)), m_DetailTopStart + intCounter * m_LineH + intTop + Printer.TextHeight("¡@") + 40)
                     Printer.FontBold = False
                  End If
                  'Removed by Morgan 2022/8/4
                  'If m_b2Picture Then
                  '   Picture1.CurrentX = (0 + intDefault) * douExtRate + Picture1.TextWidth("¡@¡@")
                  '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
                  '   Picture1.Print strExc(1) & "¡C(¶Ç¯u¸¹½X¡G886 2 25011666)"
                  '   Picture1.Line (douExtRate * (0 + intDefault) + Picture1.TextWidth("¡@¡@"), douExtRate * (m_DetailTopStart + intCounter * m_LineH + intTop) + Picture1.TextHeight("¡@") + 10)-(douExtRate * (0 + intDefault) + Picture1.TextWidth("¡@¡@") + Picture1.TextWidth(strExc(1)), douExtRate * (m_DetailTopStart + intCounter * m_LineH + intTop) + Picture1.TextHeight("¡@") + 10)
                  '   Picture1.Line (douExtRate * (0 + intDefault) + Picture1.TextWidth("¡@¡@"), douExtRate * (m_DetailTopStart + intCounter * m_LineH + intTop) + Picture1.TextHeight("¡@") + 40)-(douExtRate * (0 + intDefault) + Picture1.TextWidth("¡@¡@") + Picture1.TextWidth(strExc(1)), douExtRate * (m_DetailTopStart + intCounter * m_LineH + intTop) + Picture1.TextHeight("¡@") + 40)
                  '   Picture1.FontBold = False
                  'End If
                  'end 2022/8/4
                  
               End If 'Added by Morgan 2020/5/8
            
            End If 'Added by Morgan 2015/9/2
            
         End If
         SetWordArray m_Footer, 1, 1, strCache 'Add by Morgan 2010/12/1
         
      Case "2" '­^
         'Added by Morgan 2020/8/21 --Tim
         If m_A1k28 = "Y55435000" Then
            strCache = vbCrLf
            strCache = strCache & "¥I´Ú¤è¦¡¡G"
            strCache = strCache & vbCrLf & "½Ð±N¤W­z´Ú¶µ¶×¤J¥»¤½¥q¥H¤U±b¤á"
            strCache = strCache & vbCrLf & "»È¦æ¦WºÙ¡G·ç¿³»È¦æ ªø¦w¤À¦æ"
            strCache = strCache & vbCrLf & "¤á¦W¡G¥x¤@°ê»Ú´¼¼z°]²£¨Æ°È©Ò¡@±b¸¹¡G0075 21 1756680"
            strCache = strCache & vbCrLf & "¥I´Ú®É½Ð±N¬ÛÃö½Ð´Ú³æ¸¹¥HE-mail³qª¾¡A¥H§Q¨R±b¡AÁÂÁÂ±zªº¤ä«ù»P¦X§@¡C"
            strCache = strCache & vbCrLf
            SetWordArray m_Footer, 1, 1, strCache
            SetWordArray m_Footer, 1, 2, "Wire" & vbCrLf & "Transfer" & vbCrLf & "Preferred"
         
         
         'Added by Morgan 2025/9/1 --Joanne
         ElseIf m_A1k28 = "Y54444000" Then
            'Removed by Morgan 2025/9/1 Äæ¦ì­n²ÊÅé¡A§ï¦bWord¤º±±¨î
            'strCache = vbCrLf
            'strCache = strCache & "Account Holder's Name: Tai E International Patent and Law Office"
            'strCache = strCache & vbCrLf & "Beneficiary Address: "
            'strCache = strCache & vbCrLf & "COUNTRY: Taiwan, R.O.C."
            'strCache = strCache & vbCrLf & "COUNTRY SUBDIVISION: Taipei"
            'strCache = strCache & vbCrLf & "TOWN NAME: N/A"
            'strCache = strCache & vbCrLf & "STREET NAME: 9Fl.,No.112, Sec.2, Chang-An E. Rd"
            'strCache = strCache & vbCrLf & "POST CODE: 10491"
            
            'strCache = strCache & vbCrLf & "Beneficiary Bank Name: " & ReportSum(71001)
            'strCache = strCache & vbCrLf & "Address of the Beneficiary Bank: "
            'strCache = strCache & vbCrLf & "COUNTRY: Taiwan, R.O.C."
            'strCache = strCache & vbCrLf & "COUNTRY SUBDIVISION: Taipei"
            'strCache = strCache & vbCrLf & "TOWN NAME: N/A"
            'strCache = strCache & vbCrLf & "STREET NAME: 120 , Sec. 1. Chongqing S. Rd."
            'strCache = strCache & vbCrLf & "POST CODE:100005"
            'strCache = strCache & vbCrLf & ReportSum(74) 'Account No.((Multi-Currency Account)
            'strCache = strCache & vbCrLf & ReportSum(121) 'Account No.(for Taiwan currency)
            'strCache = strCache & vbCrLf & ReportSum(73001) 'SWIFT Code
            'If m_DNCurr <> "NTD" And m_iPrintCurrType <> 3 Then
            '   strCache = strCache & vbCrLf & "Currency Rate: " & m_DNCurr & "1.00=NTD" & Format(Val("0" & m_DNRate))
            'End If
            'strCache = strCache & vbCrLf
            'SetWordArray m_Footer, 1, 1, strCache
            'end 2025/9/1
            
            SetWordArray m_Footer, 1, 2, "Wire" & vbCrLf & "Transfer" & vbCrLf & "Preferred"
         'end 2025/9/1
         Else
         
            If m_b2Printer Then
               Printer.CurrentX = 0 + intDefault
               Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
               'Modified by Morgan 2022/8/4
               'Printer.Print ReportSum(71001)
               PUB_PrintUnicodeText ReportSum(71001), Printer.CurrentX, Printer.CurrentY, 0
               'end 2022/8/4
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.CurrentX = (0 + intDefault) * douExtRate
            '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
            '   Picture1.Print ReportSum(71001)
            'End If
            'end 2022/8/4
            strCache = ReportSum(71001) 'Add by Morgan 2010/11/30
            
            intCounter = intCounter + 1
            
            m_tmp = ReportSum(72)
            'Added by Morgan 2023/8/14 Y53195¤§©Ò¦³®×¥ó½Ð´Ú³æªº¦a§}¡A¨ú®øR. O. C. --§d°û²ñ,³¢¶®®S,¾G¤Ñ¶³
            If Left(m_A1k03, 8) = "Y5319500" Then
               m_tmp = Replace(m_tmp, ", R.O.C.", "")
            End If
               
            '93.6.9 MODIFY BY SONIA David­n¨D¥[¦L¦a§}
            If m_b2Printer Then
               Printer.CurrentX = 0 + intDefault
               Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
               'Modified by Morgan 2022/8/4
               'Printer.Print m_tmp
               PUB_PrintUnicodeText m_tmp, Printer.CurrentX, Printer.CurrentY, 0
               'end 2022/8/4
            End If
            'Add by Morgan 2008/4/7 ²£¥Í¹q¤lÀÉ¦P®É¦L¤@¥÷
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.CurrentX = (0 + intDefault) * douExtRate
            '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
            '   Picture1.Print m_tmp
            'End If
            'end 2022/8/4
            strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/11/30
            
            intCounter = intCounter + 1
            '93.6.9 END
            If m_b2Printer Then
               Printer.CurrentX = 0 + intDefault
               Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
               'Modified by Morgan 2022/8/4
               'Printer.Print ReportSum(73001)
               PUB_PrintUnicodeText ReportSum(73001), Printer.CurrentX, Printer.CurrentY, 0
               'end 2022/8/4
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.CurrentX = (0 + intDefault) * douExtRate
            '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
            '   Picture1.Print ReportSum(73001)
            'End If
            'end 2022/8/4
            strCache = strCache & vbCrLf & ReportSum(73001) 'Add by Morgan 2010/11/30
            
            'Add by Morgan 2006/7/5 ¥[¦L "Wire Transfer Preferred" -- Ä¬°ÆÁ`
            lngBoxX = intDefault + 7500
            lngBoxY = m_DetailTopStart + intCounter * m_LineH + intTop
            If m_b2Printer Then
               Printer.DrawWidth = 5
               Printer.Line (lngBoxX, lngBoxY - 100)-(lngBoxX + 1200, lngBoxY + 900), , B
               Printer.DrawWidth = 1
               Printer.CurrentX = lngBoxX + 150
               Printer.CurrentY = lngBoxY
               Printer.Print "Wire"
               Printer.CurrentX = lngBoxX + 150
               Printer.CurrentY = lngBoxY + 300
               Printer.Print "Transfer"
               Printer.CurrentX = lngBoxX + 150
               Printer.CurrentY = lngBoxY + 600
               Printer.Print "Preferred"
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.DrawWidth = 5
            '   Picture1.Line (douExtRate * (lngBoxX), douExtRate * (lngBoxY - 100))-(douExtRate * (lngBoxX + 1200), douExtRate * (lngBoxY + 900)), , B
            '   Picture1.DrawWidth = 1
            '   Picture1.CurrentX = (lngBoxX + 150) * douExtRate
            '   Picture1.CurrentY = (lngBoxY) * douExtRate
            '   Picture1.Print "Wire"
            '   Picture1.CurrentX = (lngBoxX + 150) * douExtRate
            '   Picture1.CurrentY = (lngBoxY + 300) * douExtRate
            '   Picture1.Print "Transfer"
            '   Picture1.CurrentX = (lngBoxX + 150) * douExtRate
            '   Picture1.CurrentY = (lngBoxY + 600) * douExtRate
            '   Picture1.Print "Preferred"
            'End If
            'end 2022/8/4
            'end 2006/7/5
            
           SetWordArray m_Footer, 1, 2, "Wire" & vbCrLf & "Transfer" & vbCrLf & "Preferred" 'Add by Morgan 2010/11/30
           
            intCounter = intCounter + 1
            If m_b2Printer Then
               Printer.CurrentX = 0 + intDefault
               Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
               'Modified by Morgan 2022/8/4
               'Printer.Print ReportSum(85)
               PUB_PrintUnicodeText ReportSum(85), Printer.CurrentX, Printer.CurrentY, 0
               'end 2022/8/4
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.CurrentX = (0 + intDefault) * douExtRate
            '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
            '   Picture1.Print ReportSum(85)
            'End If
            'end 2022/8/4
            strCache = strCache & vbCrLf & ReportSum(85) 'Add by Morgan 2010/11/30
            
            intCounter = intCounter + 1
            If m_b2Printer Then
               Printer.CurrentX = 0 + intDefault
               Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
               'Modified by Morgan 2022/8/4
               'Printer.Print ReportSum(74)
               PUB_PrintUnicodeText ReportSum(74), Printer.CurrentX, Printer.CurrentY, 0
               'end 2022/8/4
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.CurrentX = (0 + intDefault) * douExtRate
            '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
            '   Picture1.Print ReportSum(74)
            'End If
            'end 2022/8/4
            strCache = strCache & vbCrLf & ReportSum(74) 'Add by Morgan 2010/11/30
            
   'Removed by Morgan 2013/12/17 ¨ú®ø--ÔÑÞ±
   '         intCounter = intCounter + 1
   '         '¥[¼Ú¤¸±b¸¹
   '         If m_b2Printer Then
   '            Printer.CurrentX = 0 + intDefault
   '            Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
   '            Printer.Print ReportSum(129)
   '         End If
   '         If m_b2Picture Then
   '            Picture1.CurrentX = (0 + intDefault) * douExtRate
   '            Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
   '            Picture1.Print ReportSum(129)
   '         End If
   '         strCache = strCache & vbCrLf & ReportSum(129) 'Add by Morgan 2010/11/30
   'end 2013/12/17
            
            intCounter = intCounter + 1
            If m_b2Printer Then
               Printer.CurrentX = 0 + intDefault
               Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
               'Modified by Morgan 2022/8/4
               'Printer.Print ReportSum(121)
               PUB_PrintUnicodeText ReportSum(121), Printer.CurrentX, Printer.CurrentY, 0
               'end 2022/8/4
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   Picture1.CurrentX = (0 + intDefault) * douExtRate
            '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
            '   Picture1.Print ReportSum(121)
            'End If
            'end 2022/8/4
            strCache = strCache & vbCrLf & ReportSum(121) 'Add by Morgan 2010/11/30
            
            'Modified by Morgan 2016/9/6 m_bSpecialNew2 ®æ¦¡¤]­n¦L¶×²v
            If m_DNCurr <> "NTD" And (m_iPrintCurrType <> 3 Or m_bSpecialNew2) Then 'Add By Sindy 2013/1/17 +if ¯Â¥~¹ô¤£¶·Åã¥Ü½Ð´Ú¶×²v
               intCounter = intCounter + 1
               '2009/6/10 MODIFY BY SONIA
               'm_tmp = ReportSum(75)
               If m_DNCurr <> "USD" Then
                  'Modify By Sindy 2013/1/10
                  'm_tmp = "Currency Rate: " & m_DNCurr & "1.00=USD"
                  m_tmp = "Currency Rate: " & m_DNCurr & "1.00=NTD"
                  '2013/1/10 End
               Else
                  m_tmp = ReportSum(75)
               End If
               '2009/6/10 END
               If m_b2Printer Then
                  Printer.CurrentX = 0 + intDefault
                  Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
                  'Modified by Morgan 2022/8/4
                  'Printer.Print m_tmp
                  PUB_PrintUnicodeText m_tmp, Printer.CurrentX, Printer.CurrentY, 0
                  'end 2022/8/4
                  Printer.CurrentX = 3150 + intDefault
                  Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
                  '2009/6/10 MODIFY BY SONIA
                  'Printer.Print Format(Val("0" & m_strOriExcRate), FDollar)
                  'Printer.Print Format(Val("0" & m_strOriExcRate))
                  Printer.Print Format(Val("0" & m_DNRate)) 'Modify by Sindy 2013/1/28
                  '2009/6/10 END
               End If
               'Removed by Morgan 2022/8/4
               'If m_b2Picture Then
               '   Picture1.CurrentX = (0 + intDefault) * douExtRate
               '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
               '   Picture1.Print m_tmp
               '   Picture1.CurrentX = (3150 + intDefault) * douExtRate
               '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
               '   '2009/6/10 MODIFY BY SONIA
               '   'Picture1.Print Format(Val("0" & m_strOriExcRate), FDollar)
               '   'Picture1.Print Format(Val("0" & m_strOriExcRate))
               '   Picture1.Print Format(Val("0" & m_DNRate)) 'Modify by Sindy 2013/1/28
               '   '2009/6/10 END
               'End If
               'end 2022/8/4
               strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/11/30
               'strCache = strCache & Format(Val("0" & m_strOriExcRate)) 'Add by Morgan 2010/11/30
               strCache = strCache & Format(Val("0" & m_DNRate)) 'Modify by Sindy 2013/1/28
            End If
            SetWordArray m_Footer, 1, 1, strCache 'Add by Morgan 2010/11/30
            'Add By Sindy 2013/1/28
            If m_DNCurr <> "NTD" And strFMPFee99RMB <> "" And strA1L16 = "RMB" Then
               'a¡×·í®É¤§½Ð´Ú¹ô§O¹w¦ôµ²¶×¶×²v¡þ·í®É¤§RMB³ø»ù¶×²v  (¶×²vÄæ¥u§ì¤p¼Æ¤T¦ì)
               dblRMBRate = Trunc(Trunc(PUB_GetAcc210(2, m_DNCurr, m_strA1K02), 3) / Trunc(PUB_GetAcc210(1, "RMB", m_strA1K02), 3), 3)
               intCounter = intCounter + 1
               m_tmp = "Currency Rate: " & m_DNCurr & "1.00=RMB"
               If m_b2Printer Then
                  Printer.CurrentX = 0 + intDefault
                  Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
                  'Modified by Morgan 2022/8/4
                  'Printer.Print m_tmp
                  PUB_PrintUnicodeText m_tmp, Printer.CurrentX, Printer.CurrentY, 0
                  'end 2022/8/4
                  Printer.CurrentX = 3150 + intDefault
                  Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
                  Printer.Print Format(Val(dblRMBRate))
               End If
               'Removed by Morgan 2022/8/4
               'If m_b2Picture Then
               '   Picture1.CurrentX = (0 + intDefault) * douExtRate
               '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
               '   Picture1.Print m_tmp
               '   Picture1.CurrentX = (3150 + intDefault) * douExtRate
               '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
               '   Picture1.Print Format(Val(dblRMBRate))
               'End If
               'end 2022/8/4
               strCache = strCache & vbCrLf & m_tmp
               strCache = strCache & Format(Val(dblRMBRate))
               SetWordArray m_Footer, 1, 1, strCache
            End If
            '2013/1/28 End
            
         End If
            
         intCounter = intCounter + 1
         If m_b2Printer Then
            Printer.FontBold = True
            Printer.CurrentX = 0 + intDefault
            Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
            'Modified by Morgan 2022/8/4
            'Printer.Print ReportSum(86001)
            PUB_PrintUnicodeText ReportSum(86001), Printer.CurrentX, Printer.CurrentY, 0
            'end 2022/8/4
            Printer.FontBold = False
         End If
         'Removed by Morgan 2022/8/4
         'If m_b2Picture Then
         '   Picture1.FontBold = True
         '   Picture1.CurrentX = (0 + intDefault) * douExtRate
         '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
         '   Picture1.Print ReportSum(86001)
         '   Picture1.FontBold = False
         'End If
         'end 2022/8/4
         strCache = "Please return a copy of the invoice(s) or indicate the invoice number(s) paid with remittance" 'Add by Morgan 2010/11/30
         'Added by Morgan 2012/11/30
         If m_bSpecial3 Then
            strCache = strCache & vbCrLf & vbCrLf & "Our invoice has been issued electronically."
         End If
         'end 2012/11/30
      
         SetWordArray m_Footer, 2, 1, strCache 'Add by Morgan 2010/11/30
         
         
         '­Y¦³³Æµù¸ê®Æ
         'edit by nick 2004/07/05 FCT ®É¡Aa1k05 ¤£¦L¡A¤£ºÞ¥ô¦ó»y¨¥
         '2013/10/24 MODIFY BY SONIA ¥[¤JA1K34¬GFCT¥[¦L¦¹Äæ
         If strRemark <> "" Then
         'If strRemark <> "" And m_CP01 <> "FCT" Then
            'Modified by Morgan 2013/11/14 ­×¥¿²ÊÅé,¸õ¦æ°ÝÃD
            arrTxt = Split(strRemark, vbCrLf)
            For ii = LBound(arrTxt) To UBound(arrTxt)
               intCounter = intCounter + 1
               If m_b2Printer Then
                  Printer.FontBold = True
                  Printer.CurrentX = 0 + intDefault + Printer.TextWidth("PS: ")
                  Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
                  'Modified by Morgan 2022/8/4
                  'Printer.Print arrTxt(ii)
                  PUB_PrintUnicodeText arrTxt(ii), Printer.CurrentX, Printer.CurrentY, 0
                  'end 2022/8/4
                  Printer.FontBold = False
               End If
               'Removed by Morgan 2022/8/4
               'If m_b2Picture Then
               '   Picture1.FontBold = True
               '   Picture1.CurrentX = (0 + intDefault) * douExtRate + Picture1.TextWidth("PS: ")
               '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
               '   Picture1.Print arrTxt(ii)
               '   Picture1.FontBold = False
               'End If
               'end 2022/8/4
            Next
            'end 2013/11/14
            SetWordArray m_Footer, 3, 1, strRemark  'Add by Morgan 2010/11/30
         End If
         
        
    Case "3" '¤é
        'Add by Morgan 2005/4/13 ±±¨î¸õ­¶
        'Modify by Morgan 2006/3/24
        'intCounter = intCounter + 3
        'If intCounter >= 20 Then
        If intCounter >= 15 Then
            strNewPage = MsgText(602)
            intCounter = 0
            If m_b2Printer Then
               'Modified by Morgan 2012/10/31
               'Printer.NewPage
               MyNewPage
            End If
            'Removed by Morgan 2022/8/4
            'If m_b2Picture Then
            '   SetPic m_iPages
            'End If
            'end 2022/8/4
            m_iPages = m_iPages + 1
            m_DetailTopStart = 1500 + 4 * m_LineH
        End If
         
         
         If bolFMPCase = False Then 'Added by Morgan 2013/5/13
            'Modify By Sindy 2013/1/28
            'm_tmp = ReportSum(75001) & Format(Val("0" & m_strOriExcRate), FDollar)
            If m_DNCurr <> "NTD" And strFMPFee99RMB <> "" And strA1L16 = "RMB" Then
               'a¡×·í®É¤§½Ð´Ú¹ô§O¹w¦ôµ²¶×¶×²v¡þ·í®É¤§RMB³ø»ù¶×²v  (¶×²vÄæ¥u§ì¤p¼Æ¤T¦ì)
               dblRMBRate = Trunc(Trunc(PUB_GetAcc210(2, m_DNCurr, m_strA1K02), 3) / Trunc(PUB_GetAcc210(1, "RMB", m_strA1K02), 3), 3)
               'Modified by Morgan 2016/2/19
               'm_tmp = "²{¦bÇU¬°´ÀÇè¡ÐÇÄ¡GUSD1.00=NTD " & Format(Val("0" & m_DNRate), FDollar) & "¡F" & m_DNCurr & "1.00=RMB " & Format(Val(dblRMBRate))
               'Modified by Morgan 2022/8/17
               'm_tmp = "²{¦bÇU¬°´ÀÇè¡ÐÇÄ¡G" & m_DNCurr & "1.00=NTD " & Format(Val("0" & m_DNRate), FDollar) & "¡F" & m_DNCurr & "1.00=RMB " & Format(Val(dblRMBRate))
               m_tmp = PUB_GetUniText(Me.Name, "»È1") & m_DNCurr & "1.00=NTD " & Format(Val("0" & m_DNRate), FDollar) & "¡F" & m_DNCurr & "1.00=RMB " & Format(Val(dblRMBRate))
               'end 2022/8/17
               PutData m_tmp, intCounter, 0, m_DetailTopStart
               strCache = strCache & vbCrLf & m_tmp
            Else
            '2013/1/28 End
               'Modified by Morgan 2016/2/19
               'm_tmp = ReportSum(75001) & Format(Val("0" & m_DNRate), FDollar)
               'Modified by Morgan 2022/8/17
               'm_tmp = "²{¦bÇU¬°´ÀÇè¡ÐÇÄ¡G" & m_DNCurr & "1.00=NTD " & Format(m_DNRate)
               m_tmp = PUB_GetUniText(Me.Name, "»È1") & m_DNCurr & "1.00=NTD " & Format(m_DNRate)
               'end 2022/8/17
               'end 2016/2/19
               PutData m_tmp, intCounter, 0, m_DetailTopStart
               strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/12/1
            End If
            intCounter = intCounter + 1
         End If 'Added by Morgan 2013/5/13
         
         m_tmp = ReportSum(71002)
         PutData m_tmp, intCounter, 0, m_DetailTopStart
         strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/12/1
                  
         intCounter = intCounter + 1
         'Add by Morgan 2005/4/4 »È¦æ¦a§}
         m_tmp = ReportSum(72001)
         PutData m_tmp, intCounter, 0, m_DetailTopStart
         strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/12/1
         
         intCounter = intCounter + 1
         m_tmp = ReportSum(73001)
         PutData m_tmp, intCounter, 0, m_DetailTopStart
         strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/12/1
         
         intCounter = intCounter + 1
        'Add by Morgan 2005/4/4 ¤½¥q¦W
         m_tmp = ReportSum(11601)
         PutData m_tmp, intCounter, 0, m_DetailTopStart
         strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/12/1
                  
         intCounter = intCounter + 1
         m_tmp = ReportSum(74001)
         PutData m_tmp, intCounter, 0, m_DetailTopStart
         strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/12/1
         
         intCounter = intCounter + 1
         m_tmp = ReportSum(12101)
         PutData m_tmp, intCounter, 0, m_DetailTopStart
         strCache = strCache & vbCrLf & m_tmp 'Add by Morgan 2010/12/1
         
         SetWordArray m_Footer, 1, 1, strCache 'Add by Morgan 2010/12/1
         
         'Added by Lydia 2020/08/07 ¼W¥[¦C¦L³Æµù(a1k05)  ; ¦]¬°X55778000±ä¹Fªº¥N²z¤HY55339000ªº©w½Z­^¤å±q­^¤å§ï¬°¤é¤å¡A©Ò¥H¤é¤å½Ð´Ú³æ§À³¡¼W¥[¦C¦L³Æµù(¤ñ·Ó­^¤åª©)
         If strRemark <> "" Then
            arrTxt = Split(strRemark, vbCrLf)
            For ii = LBound(arrTxt) To UBound(arrTxt)
               intCounter = intCounter + 1
               If m_b2Printer Then
                  Printer.FontBold = True
                  Printer.CurrentX = 0 + intDefault + Printer.TextWidth("PS: ")
                  Printer.CurrentY = m_DetailTopStart + intCounter * m_LineH + intTop
                  'Modified by Morgan 2022/8/4
                  'Printer.Print arrTxt(ii)
                  PUB_PrintUnicodeText arrTxt(ii), Printer.CurrentX, Printer.CurrentY, 0
                  'end 2022/8/4
                  Printer.FontBold = False
               End If
               'Removed by Morgan 2022/8/4
               'If m_b2Picture Then
               '   Picture1.FontBold = True
               '   Picture1.CurrentX = (0 + intDefault) * douExtRate + Printer.TextWidth("PS: ")
               '   Picture1.CurrentY = (m_DetailTopStart + intCounter * m_LineH + intTop) * douExtRate
               '   Picture1.Print arrTxt(ii)
               '   Picture1.FontBold = False
               'End If
               'end 2022/8/4
            Next
            SetWordArray m_Footer, 2, 1, strRemark '¦C¦L¦Û°Ê·|¥[¤WPS:
         End If
         'end 2020/08/07
         intCounter = intCounter + 1
    End Select
   
   'Add by Morgan 2011/6/24
   If m_b2Printer Then
      m_iPageCount = Printer.Page
   End If
   
   'Add by Morgan 2010/11/5
   If m_bEBilling And (m_b2Printer Or m_b2Word Or Check2.Value = vbChecked) Then
      If WriteLEDES(strLedes, m_iLedesVer) Then
         If m_bMsg Then
            MsgBox "¹q¤l±b³æÀÉ¤w¦s©ó [ " & m_EFilePath & " ]¡I"
            'Added by Morgan 2014/7/18
            If Check2.Value = vbChecked Then
               Exit Sub
            End If
            'end 2014/7/18
         End If
      End If
   End If
   
   'Add by Morgan 2010/12/1
   If m_b2Word Then
      'Modified by Lydia 2019/12/27 ­ì¥»¦@¥ÎÅÜ¼Æpub_OsPrinter¦]¬°±qFCP¦~¶Oµo¤å->½Ð´Ú¨ç¦C¦L->½Ð´Ú³æ¦C¦L¡A«h¤¤¶¡¹Lµ{OS¹w³]¦Lªí¾÷¦³ÅÜ¤Æ¡A³y¦¨µLªk¦^¨ìµo¤å«eªºOS¹w³]¦Lªí¾÷
      'pub_OsPrinter = PUB_GetOsDefaultPrinter
      m_OsPrinter = PUB_GetOsDefaultPrinter
      If m_bWord2Pdf Then
         MyNewPage True
         PUB_SetOsDefaultPrinter Printer.DeviceName
      Else
         PUB_SetOsDefaultPrinter Combo1
      End If
      PUB_SetWordActivePrinter
      
      Select Case strLanguage
         Case "1" '¤¤
            runWordChinese
            
         Case "2" '­^
            runWordEnglish
            
         Case "3" '¤é
            runWordJapanese
      End Select
      
        'Added by Morgan 2012/11/1
        If m_bWord2Pdf Then
           frmPDF.EndtProcess
           'Added by Lydia 2020/02/15 §PÂ_ÀÉ®×¬O§_¦s¦b, ¶W¹L®É¶¡´NÄ~Äò
           If Me.Tag <> "" And InStr(Me.Tag, "\") > 0 Then
             'Modified by Lydia 2020/09/10 ¶W¹L®É¶¡¡Aª½±µ°O¿ý¥¢±Ñ²M³æ
             'If PUB_ChkFileStatus(Me.Tag) = False Then
             If PUB_ChkFileStatus(Me.Tag, False, m_strOutErr) = False Then
             End If
           End If
           'end 2020/02/15
           Unload frmPDF
        'Added by Morgan 2014/8/28
        '¯S®í½Ð´Ú³æ
        ElseIf m_PlusFormNo <> "" Then
           PUB_SetOsDefaultPrinter Combo1
           PUB_SetWordActivePrinter
           runWordJapaneseNew True
        'end 2014/8/28
        
        End If

      'Modified by Lydia 2019/12/27 ­ì¥»¦@¥ÎÅÜ¼Æpub_OsPrinter¦]¬°±qFCP¦~¶Oµo¤å->½Ð´Ú¨ç¦C¦L->½Ð´Ú³æ¦C¦L¡A«h¤¤¶¡¹Lµ{OS¹w³]¦Lªí¾÷¦³ÅÜ¤Æ¡A³y¦¨µLªk¦^¨ìµo¤å«eªºOS¹w³]¦Lªí¾÷
      'PUB_SetOsDefaultPrinter pub_OsPrinter
      PUB_SetOsDefaultPrinter m_OsPrinter
      m_b2Word = False
   End If

End Sub

'*************************************************
'  µe­±¿é¤JÀË¬d
'
'*************************************************
Public Function FormCheck() As Boolean
   'Added by Morgan 2017/2/8
   If Check2.Value = vbChecked And txtOutMode <> "1" Then
      MsgBox "¤Ä¿ï¡i¥u²£¥ÍLEDES¹q¤l±b³æ¡j®É¡i½Ð´Ú³æ¿é¥X¤è¦¡¡j¥²¶·¬°¡i¦Lªí¾÷¡j¡I", vbExclamation
      Exit Function
   End If
   'end 2017/2/8
   
   'Modified by Morgan 2023/7/12
   'If Text1 <> MsgText(601) Then
   '   FormCheck = True
   '   Exit Function
   'End If
   'If Text2 <> MsgText(601) Then
   '   FormCheck = True
   '   Exit Function
   'End If
   If Text1 <> MsgText(601) And Text2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   'end 2023/7/12
   FormCheck = False
   
   MsgBox "D/N No.¤£¥i¬°ªÅ¥Õ¡I", vbExclamation 'Added by Morgan 2017/2/8
End Function

'Add By Cheng 2003/02/24
'¨ú±o®×¥ó¦WºÙ
'Modify by Morgan 2006/9/21 strLang:0=¤¤+­^+¤é 1=¤¤ 2=­^ 3=¤é
Private Function GetCaseName(strCP0104 As String, Optional strLang As String = "0") As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select PA05,PA06,PA07 From Patent Where " & ChgPatent(strCP0104)
'Modify By Sindy 2015/7/8 tm05 ==> nvl(tm131,tm05)
StrSQLa = StrSQLa & " union Select nvl(tm131,tm05) TM05,TM06,TM07 From Trademark Where " & ChgTradeMark(strCP0104)
StrSQLa = StrSQLa & " union Select LC05,LC06,LC07 From Lawcase Where " & ChgLawcase(strCP0104)
StrSQLa = StrSQLa & " union Select SP05,SP06,SP07 From ServicePractice Where " & ChgService(strCP0104)
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    'Modify By Cheng 2003/05/16
'    GetCaseName = "" & rsA.Fields(1).Value
   Select Case strLang
      Case "1"
         GetCaseName = "" & rsA.Fields(0).Value
      Case "2"
         GetCaseName = "" & rsA.Fields(1).Value
      Case "3"
         GetCaseName = "" & rsA.Fields(2).Value
      Case Else
         GetCaseName = Trim("" & rsA.Fields(0).Value & " " & rsA.Fields(1).Value & " " & rsA.Fields(2).Value)
   End Select
   
Else
    GetCaseName = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Add By Cheng 2003/03/13
'­pºâ¦r¼Æ
Private Function CountLength(strWord As String) As Double
Dim ii As Integer
Dim strChr As String
    
CountLength = 0
If strWord <> "" Then
    For ii = 1 To Len(strWord)
        strChr = Mid(strWord, ii, 1)
        If Asc(strChr) >= 65 And Asc(strChr) <= 90 Then
            CountLength = CountLength + 1.5
        ElseIf Asc(strChr) >= 128 Then
            CountLength = CountLength + 2
        'Add by Morgan 2004/12/17
        ElseIf Asc(strChr) < 0 Then
            CountLength = CountLength + 2
        Else
            CountLength = CountLength + 1
        End If
    Next ii
End If
End Function
'¤å¦r§é¦æ
Private Sub PrintDropLine(strWord As String, strLineTitle As String, intRow As Integer, intLineChrs As Integer, Optional intBeginDiff As Integer = 0)
   Dim ii As Integer
   Dim jj As Integer
   Dim kk As Integer
   Dim dblChrCnt As Double
   Dim strArr
   Dim strWordPrint As String
   Dim intIntiRow As Integer

   kk = 0
   dblChrCnt = 0
   strWordPrint = ""
   intIntiRow = intRow
   If Trim(strWord) <> "" Then
      strArr = Split(RTrim(strWord), " ")
      For ii = LBound(strArr) To UBound(strArr)
      '­Y°}¦C¬°ªÅ¦r¦êªº³]©w¬°ªÅ¥Õ
         If strArr(ii) = "" Then strArr(ii) = " "
            For jj = 1 To Len(strArr(ii))
               If Asc(Mid(strArr(ii), jj, 1)) >= 65 And Asc(Mid(strArr(ii), jj, 1)) <= 90 Then
                  dblChrCnt = dblChrCnt + 1.5
               'Modify by Morgan 2005/3/31 ³y¦r¤p©ó0
               'ElseIf (Asc(Mid(strArr(ii), jj, 1)) >= 128
               ElseIf (Asc(Mid(strArr(ii), jj, 1)) >= 128 Or Asc(Mid(strArr(ii), jj, 1)) < 0) Then
                  dblChrCnt = dblChrCnt + 2
               Else
                  dblChrCnt = dblChrCnt + 1
               End If
            Next jj
            If dblChrCnt + 1 > intLineChrs Then
               If kk = 0 Then
                  If m_b2Printer Then
                     Printer.CurrentX = 400 + intDefault + intBeginDiff
                     Printer.CurrentY = 1500 + intRow * m_LineH + intTop
                     'Modified by Morgan 2022/8/4
                     'Printer.Print strLineTitle & strWordPrint
                     PUB_PrintUnicodeText strLineTitle & strWordPrint, Printer.CurrentX, Printer.CurrentY, 0
                     'end 2022/8/4
                  End If
                  'Removed by Morgan 2022/8/4
                  'If m_b2Picture Then
                  '   Picture1.CurrentX = (400 + intDefault + intBeginDiff) * douExtRate
                  '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
                  '   Picture1.Print strLineTitle & strWordPrint
                  'End If
                  'end 2022/8/4
               Else
                  If m_b2Printer Then
                     Printer.CurrentX = 400 + intDefault + Printer.TextWidth(strLineTitle) + intBeginDiff
                     Printer.CurrentY = 1500 + intRow * m_LineH + intTop
                     'Modified by Morgan 2022/8/4
                     'Printer.Print strWordPrint
                     PUB_PrintUnicodeText strWordPrint, Printer.CurrentX, Printer.CurrentY, 0
                     'end 2022/8/4
                  End If
                  'Removed by Morgan 2022/8/4
                  'If m_b2Picture Then
                  '   Picture1.CurrentX = (400 + intDefault + intBeginDiff) * douExtRate + Picture1.TextWidth(strLineTitle)
                  '   Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
                  '   Picture1.Print strWordPrint
                  'End If
                  'end 2022/8/4
               End If
               kk = kk + 1
               intRow = intRow + 1
               dblChrCnt = 0
               For jj = 1 To Len(strArr(ii))
                  If Asc(Mid(strArr(ii), jj, 1)) >= 65 And Asc(Mid(strArr(ii), jj, 1)) <= 90 Then
                     dblChrCnt = dblChrCnt + 1.5
                  'Modify by Morgan 2005/3/31 ³y¦r¤p©ó0
                  'ElseIf Asc(Mid(strArr(ii), jj, 1)) >= 128 Then
                  ElseIf (Asc(Mid(strArr(ii), jj, 1)) >= 128 Or Asc(Mid(strArr(ii), jj, 1)) < 0) Then
                     dblChrCnt = dblChrCnt + 2
                  Else
                     dblChrCnt = dblChrCnt + 1
                  End If
               Next jj
               strWordPrint = ""
               strWordPrint = strWordPrint & strArr(ii) & IIf(strArr(ii) = " ", "", " ")
               dblChrCnt = dblChrCnt + 1
            Else
               strWordPrint = strWordPrint & strArr(ii) & IIf(strArr(ii) = " ", "", " ")
               dblChrCnt = dblChrCnt + 1
            End If
      Next ii
      If strWordPrint <> "" Then
         If m_b2Printer Then
            If kk = 0 Then
               Printer.CurrentX = 400 + intDefault + intBeginDiff
               Printer.CurrentY = 1500 + intRow * m_LineH + intTop
               'Modified by Morgan 2022/8/4
               'Printer.Print strLineTitle & strWordPrint
               PUB_PrintUnicodeText strLineTitle & strWordPrint, Printer.CurrentX, Printer.CurrentY, 0
               'end 2022/8/4
            Else
               Printer.CurrentX = 400 + intDefault + Printer.TextWidth(strLineTitle) + intBeginDiff
               Printer.CurrentY = 1500 + intRow * m_LineH + intTop
               'Modified by Morgan 2022/8/4
               'Printer.Print strWordPrint
               PUB_PrintUnicodeText strWordPrint, Printer.CurrentX, Printer.CurrentY, 0
               'end 2022/8/4
            End If
         End If
         'Removed by Morgan 2022/8/4
         'If m_b2Picture Then
         '   If kk = 0 Then
         '      Picture1.CurrentX = (400 + intDefault + intBeginDiff) * douExtRate
         '      Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
         '      Picture1.Print strLineTitle & strWordPrint
         '   Else
         '      Picture1.CurrentX = (400 + intDefault + intBeginDiff) * douExtRate + Picture1.TextWidth(strLineTitle)
         '      Picture1.CurrentY = (1500 + intRow * m_LineH + intTop) * douExtRate
         '      Picture1.Print strWordPrint
         '   End If
         'End If
         'end 2022/8/4
      End If
      intCounter = intCounter + (intRow - intIntiRow)
   End If
End Sub

Private Sub txtOutMode_Change()
   If txtOutMode = "2" Then
      Check2.Value = vbUnchecked
   End If
End Sub

Private Sub txtOutMode_GotFocus()
   TextInverse txtOutMode
   CloseIme
End Sub

Private Sub txtOutMode_KeyPress(KeyAscii As Integer)
   If KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtAdd_GotFocus()
    'Add By Cheng 2003/03/19
    TextInverse Me.txtAdd
End Sub

Private Sub txtAdd_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2003/03/19
    KeyAscii = UpperCase(KeyAscii)
    Select Case KeyAscii
    Case 8, 89
        'µL°Ê§@
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txtCopy_GotFocus()
    'Add By Cheng 2003/03/18
    TextInverse Me.txtCopy
End Sub

Private Sub txtCopy_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2003/03/18
    KeyAscii = UpperCase(KeyAscii)
    Select Case KeyAscii
    'Modify by Morgan 2005/6/23 ¥[52(4)
    'Case 8, 49, 50, 51
    Case 8, 49, 50, 51, 52
        'µL°Ê§@
    Case Else
        KeyAscii = 51
    End Select
End Sub

'Add By Cheng 2003/04/03
Private Function GetNationName(strKind As String) As String
'strKind : 2 ­^¤å
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   GetNationName = ""
   'Modified by Morgan 2013/5/13 +na01
   StrSQLa = "Select NA03, NA04,na01 From Nation, Patent Where NA01=PA09 And " & ChgPatent(m_CP01 & m_CP02 & m_CP03 & m_CP04)
   StrSQLa = StrSQLa & " Union Select NA03, NA04,na01 From Nation, Trademark Where NA01=TM10 And " & ChgTradeMark(m_CP01 & m_CP02 & m_CP03 & m_CP04)
   StrSQLa = StrSQLa & " Union Select NA03, NA04,na01 From Nation, Lawcase Where NA01=LC15 And " & ChgLawcase(m_CP01 & m_CP02 & m_CP03 & m_CP04)
   StrSQLa = StrSQLa & " Union Select NA03, NA04,na01 From Nation, Hirecase Where '000'=NA01 And " & ChgHirecase(m_CP01 & m_CP02 & m_CP03 & m_CP04)
   StrSQLa = StrSQLa & " Union Select NA03, NA04,na01 From Nation, Servicepractice Where NA01=SP09 And " & ChgService(m_CP01 & m_CP02 & m_CP03 & m_CP04)
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       GetNationName = IIf(strKind = "2", "" & rsA.Fields(1).Value, "" & rsA.Fields(0).Value)
       '2011/6/21 ADD BY SONIA
       If strKind = "3" Then
         If GetNationName = "¥xÆW" Then
            'Modified by Morgan 2022/8/3
            'GetNationName = "¥x“ø"
            GetNationName = PUB_GetUniText(Me.Name, "¥xÆW")
            'end 2022/8/3
         'Added by Morgan 2013/5/13
         ElseIf rsA.Fields("na01") = "020" Then
            'Modified by Morgan 2022/8/3
            'GetNationName = "¤¤üÂ"
            GetNationName = PUB_GetUniText(Me.Name, "¤¤°ê")
            'end 2022/8/3
         End If
       End If
       '2011/6/21 END
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   'Add By Cheng 2003/04/14
   '2009/7/3 MODIFY BY SONIA ­^¤å¤~¥[ªÅ®æ
   'If GetNationName <> "" Then GetNationName = GetNationName & " "
   If GetNationName <> "" And strKind = "2" Then GetNationName = GetNationName & " "

End Function
'¨ú±o©w½Z»y¤å
'Modify by Morgan 2006/5/29 ¥[½Ð´Ú³æ¸¹ p_A1K01 ¥H§PÂ_¬O§_¬°¦~¶O½Ð´Ú¨Ã§ïCall¦@¥Î¨ç¼ÆPUB_GetLanguage
Public Function GetLanguage(strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, ByVal p_A1K01 As String) As String
   Dim strLanguage As String
   
   'Added by Morgan 2017/9/20 FCP57495½Ð­Ó®×³]©w±b³æ»y¨¥:­^¤å --¦ó²QµØ
   'Modified by Morgan 2018/1/5
   'If strCP01 = "FCP" And strCP02 = "057495" And strCP03 = "0" And strCP04 = "00" Then
   '   GetLanguage = "2"
   If GetBillLanguage(strLanguage) Then
      GetLanguage = strLanguage
      Exit Function
   End If
   'end 2018/1/5
   'end 2017/9/20

   'Add by Morgan 2006/5/29
   Dim strCP10 As String, strSysKind As String
   If strCP01 = "FCP" And p_A1K01 <> "" Then
      strSysKind = "1"
      strSql = "select * from ACC1L0 where A1L01='" & p_A1K01 & "' AND A1L04='605'"
      CheckOC
      With adoRecordset
         .CursorLocation = adUseClient
         .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
         If .RecordCount > 0 Then
            strCP10 = "605"
         End If
      End With
   End If
   
   GetLanguage = PUB_GetLanguage(strCP01, strCP02, strCP03, strCP04, strCP10, strSysKind)
   
End Function

'Add By Cheng 2003/10/16
'¨ú±o®×¥ó©Ê½è
Private Function GetCP10(strA1K01 As String) As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   
   GetCP10 = ""
   'Modify by Morgan 2007/4/2 ¥[§PÂ_¬ÛÃöÁ`¦¬¤å¸¹ªº®×¥ó©Ê½è
   StrSQLa = "Select C1.CP01 ,C1.CP10,C2.CP10 CP10x From Caseprogress C1, Caseprogress C2 Where C1.CP60='" & strA1K01 & "' AND C2.CP09(+)=C1.CP43"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      Do While Not rsA.EOF
         GetCP10 = "" & rsA("CP10").Value
         Select Case "" & rsA("CP01").Value
            Case "FCP", "CFP", "P"
               If "" & rsA("CP10").Value = "605" Then
                  Exit Do
               ElseIf "" & rsA("CP10x").Value = "605" Then
                  GetCP10 = "" & rsA("CP10x").Value
                  Exit Do
               End If
            Case "FCT", "CFT", "T", "TF"
               If "" & rsA("CP10").Value = "102" Then
                  Exit Do
               ElseIf "" & rsA("CP10x").Value = "102" Then
                  GetCP10 = "" & rsA("CP10x").Value
                  Exit Do
               End If
         End Select
         rsA.MoveNext
      Loop
   End If
   'end 2007/4/2
   
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing

End Function

'Add By Cheng 2003/12/11
'¨ú±o°Ó«~Ãþ§O¼Æ
Private Function GetTMKindCnt(strTM01 As String, strTM02 As String, strTM03 As String, strTM04 As String) As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetTMKindCnt = 0
StrSQLa = "Select TM09 From Trademark Where " & ChgTradeMark(strTM01 & strTM02 & strTM03 & strTM04)
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    If "" & rsA.Fields(0).Value = "" Then
        GetTMKindCnt = 1
    Else
        GetTMKindCnt = UBound(Split("" & rsA.Fields(0).Value, ",")) + 1
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'Add by Morgan 2011/3/4
Private Function GetNewDescFCT10199(stra1l01 As String, stra1l04 As String, strFrom As String, strTo As String) As String
   Dim strCon As String, strSql As String
   Dim iR As Integer, adoRst As ADODB.Recordset
   Dim strDesc As String, strLastItemDesc As String, strThisItemDesc As String
   Dim dblOverFee As Double, iSameItemCnt As Integer
   Dim strAdd As String
   Dim dblFee As Double
   Dim iNameDiscCount As Integer
   Dim boleFiling As Boolean
   Dim stA1l02 As String 'Added by Morgan 2014/3/5
   
   strCon = " and a1l02>='" & strFrom & "' and a1l02<='" & strTo & "'"
   
   '¬O§_¹q¤l°e¥ó
   'Modify By Sindy 2020/1/10 and cp118='Y' => and cp118 is not null
   'strExc(0) = "select a1l02 from caseprogress,acc1l0 where cp60='" & stra1l01 & "' and cp01 in ('T','FCT') and cp10='101' and cp118='Y' and a1l01(+)=cp60" & strCon & " order by a1l02"
   strExc(0) = "select a1l02 from caseprogress,acc1l0 where cp60='" & stra1l01 & "' and cp01 in ('T','FCT') and cp10='101' and cp118 is not null and a1l01(+)=cp60" & strCon & " order by a1l02"
   '2020/1/10 END
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      boleFiling = True
      stA1l02 = RsTemp("a1l02") 'Added by Morgan 2014/3/5
   Else
      boleFiling = False
   End If
   iNameDiscCount = 0
   
   'Modify by Morgan 2011/7/28 °Ó«~¼Æ¦³¶W¹L­n¥ý¦C¦L--ªü½¬ Ex.X10007911
   'strSql = "select * from acc1l0 where a1l01='" & stra1l01 & "'" & strCon & " and a1l05>0 order by a1l02"
   'Modified by Morgan 2014/3/5 ·Ó°Ó«~¼Æ±Æ§Ç
   'strSql = "select * from acc1l0 where a1l01='" & stra1l01 & "'" & strCon & " and a1l05>0 order by decode(nvl(a1l14,0),0,1,0),a1l02"
   strSql = "select * from acc1l0 where a1l01='" & stra1l01 & "'" & strCon & " and a1l05>0 order by nvl(a1l14,0) desc,a1l02"
   
   iR = 1
   Set adoRst = ClsLawReadRstMsg(iR, strSql)
   If iR = 1 Then
      iSameItemCnt = 1
      With adoRst
      dblFee = .Fields("a1l05")
      '¹q¤l°e¥ó¥i´î 300
      If boleFiling Then
         If .Fields("a1l02") = stA1l02 Then 'Added by Morgan 2014/3/5
            dblFee = dblFee + 300
         End If
      End If
      
      '°Ó¼Ð¦WºÙ²Å¦X³W©w¥i´î§K 300
      If .Fields("a1l15") = "Y" Then
         dblFee = dblFee + 300
         'Modified by Morgan 2013/2/4
         'iNameDiscCount = iNameDiscCount + 300
         iNameDiscCount = iNameDiscCount + 1
         'End 2013/2/4
      End If
      
      '°Ó«~¼Æ¦³¶W¹L
      If .Fields("a1l14") > 0 Then
         If .Fields("a1l04") = "A0199" Then
            dblOverFee = 500 * .Fields("a1l14")
            'Added by Morgan 2015/3/3
            'If m_iPrintCurrType = 3 Or m_iPrintCurrType = 4 Then
            '   strLastItemDesc = "(" & m_DNCurr & " " & PUB_ChgFormat((dblFee - dblOverFee) / m_DNRate, True) & " + " & m_DNCurr & " " & PUB_ChgFormat(500 / m_DNRate, True) & " x " & .Fields("a1l14") & ")"
            'Else
            'end 2015/3/3
               strLastItemDesc = "(NTD " & Format(dblFee - dblOverFee, DDollar) & " + NTD 500 x " & .Fields("a1l14") & ")"
            'End If 'Added by Morgan 2015/3/3
         Else
            dblOverFee = 200 * .Fields("a1l14")
            'Added by Morgan 2015/3/3
            'If m_iPrintCurrType = 3 Or m_iPrintCurrType = 4 Then
            '   strLastItemDesc = "(" & m_DNCurr & " " & PUB_ChgFormat((dblFee - dblOverFee) / m_DNRate, True) & " + " & m_DNCurr & " " & PUB_ChgFormat(200 / m_DNRate, True) & " x " & .Fields("a1l14") & ")"
            'Else
            'end 2015/3/3
               strLastItemDesc = "(NTD " & Format(dblFee - dblOverFee, DDollar) & " + NTD 200 x " & .Fields("a1l14") & ")"
            'End If 'Added by Morgan 2015/3/3
         End If
      ElseIf .RecordCount > 1 Then
         'Added by Morgan 2015/3/3
         'If m_iPrintCurrType = 3 Or m_iPrintCurrType = 4 Then
         '   strLastItemDesc = m_DNCurr & " " & PUB_ChgFormat(dblFee / m_DNRate, True)
         'Else
         'end 2015/3/3
            strLastItemDesc = "NTD " & Format(dblFee, DDollar)
         'End If 'Added by Morgan 2015/3/3
      Else
         'Added by Morgan 2013/2/20
         If boleFiling Then
            'Added by Morgan 2015/3/3
            'If m_iPrintCurrType = 3 Or m_iPrintCurrType = 4 Then
            '   strDesc = m_DNCurr & " " & PUB_ChgFormat(dblFee / m_DNRate, True)
            'Else
            'end 2015/3/3
               strDesc = "NTD " & Format(dblFee, DDollar)
            'End If 'Added by Morgan 2015/3/3
         End If
         'end 2013/2/20
         GoTo NoNeedDesc
      End If
      .MoveNext
      Do While Not .EOF
         dblFee = .Fields("a1l05")
         
         'Added by Morgan 2014/3/5
         '¹q¤l°e¥ó¥i´î 300
         If boleFiling Then
            If .Fields("a1l02") = stA1l02 Then
               dblFee = dblFee + 300
            End If
         End If
         'end 2014/3/5
      
         '°Ó¼Ð¦WºÙ²Å¦X³W©w¥i´î§K 300
         If .Fields("a1l15") = "Y" Then
            dblFee = dblFee + 300
            'Modified by Morgan 2013/2/4
            'iNameDiscCount = iNameDiscCount + 300
            iNameDiscCount = iNameDiscCount + 1
            'End 2013/2/4
         End If
      
         '°Ó«~¼Æ¦³¶W¹L
         If .Fields("a1l14") > 0 Then
            If .Fields("a1l04") = "A0199" Then
               dblOverFee = 500 * .Fields("a1l14")
               'Added by Morgan 2015/3/3
               'If m_iPrintCurrType = 3 Or m_iPrintCurrType = 4 Then
               '   strThisItemDesc = "(" & m_DNCurr & " " & PUB_ChgFormat((dblFee - dblOverFee) / m_DNRate, True) & " + " & m_DNCurr & " " & PUB_ChgFormat(500 / m_DNRate, True) & " x " & .Fields("a1l14") & ")"
               'Else
               'end 2015/3/3
                  strThisItemDesc = "(NTD " & Format(dblFee - dblOverFee, DDollar) & " + NTD 500 x " & .Fields("a1l14") & ")"
               'End If 'Added by Morgan 2015/3/3
            Else
               dblOverFee = 200 * .Fields("a1l14")
               'Added by Morgan 2015/3/3
               'If m_iPrintCurrType = 3 Or m_iPrintCurrType = 4 Then
               '   strThisItemDesc = "(" & m_DNCurr & " " & PUB_ChgFormat((dblFee - dblOverFee) / m_DNRate, True) & " + " & m_DNCurr & " " & PUB_ChgFormat(200 / m_DNRate, True) & " x " & .Fields("a1l14") & ")"
               'Else
               'end 2015/3/3
                  strThisItemDesc = "(NTD " & Format(dblFee - dblOverFee, DDollar) & " + NTD 200 x " & .Fields("a1l14") & ")"
               'End If 'Added by Morgan 2015/3/3
            End If
         Else
            'Added by Morgan 2015/3/3
            'If m_iPrintCurrType = 3 Or m_iPrintCurrType = 4 Then
            '   strThisItemDesc = m_DNCurr & " " & PUB_ChgFormat(dblFee / m_DNRate, True)
            'Else
            'end 2015/3/3
               strThisItemDesc = "NTD " & Format(dblFee, DDollar)
            'End If 'Added by Morgan 2015/3/3
         End If
         If strThisItemDesc = strLastItemDesc Then
            iSameItemCnt = iSameItemCnt + 1
         Else
            If strDesc <> "" Then
               strAdd = " + "
            End If
            '¤º®e¦³­«½Æ
            If iSameItemCnt > 1 Then
               strDesc = strDesc & strAdd & strLastItemDesc & " x " & iSameItemCnt
            Else
               strDesc = strDesc & strAdd & strLastItemDesc
            End If
            strLastItemDesc = strThisItemDesc
            iSameItemCnt = 1
         End If
         .MoveNext
      Loop
      If strDesc <> "" Then
         strAdd = " + "
      End If
      If iSameItemCnt > 1 Then
         strDesc = strDesc & strAdd & strLastItemDesc & " x " & iSameItemCnt
      Else
         strDesc = strDesc & strAdd & strLastItemDesc
      End If
      
NoNeedDesc:

      'Modified by Morgan 2013/2/20
      If boleFiling Then
         If iNameDiscCount > 0 Then
            'Added by Morgan 2015/3/3
            'If m_iPrintCurrType = 3 Or m_iPrintCurrType = 4 Then
            '   strDesc = strDesc & " - (" & m_DNCurr & " " & PUB_ChgFormat(300 / m_DNRate, True)
            'Else
            'end 2015/3/3
               strDesc = strDesc & " - (NTD 300"
            'End If 'Added by Morgan 2015/3/3
            
            'Modified by Morgan 2013/2/4
            'strDesc = strDesc & " + NTD 300 x " & iNameDiscCount
            If iNameDiscCount > 1 Then
               'Added by Morgan 2015/3/3
               'If m_iPrintCurrType = 3 Or m_iPrintCurrType = 4 Then
               '   strDesc = strDesc & " + " & m_DNCurr & " " & PUB_ChgFormat(300 / m_DNRate, True) & " x " & iNameDiscCount
               'Else
               'end 2015/3/3
                  strDesc = strDesc & " + NTD 300 x " & iNameDiscCount
               'End If 'Added by Morgan 2015/3/3
            Else
               'Added by Morgan 2015/3/3
               'If m_iPrintCurrType = 3 Or m_iPrintCurrType = 4 Then
               '   strDesc = strDesc & " + " & m_DNCurr & " " & PUB_ChgFormat(300 / m_DNRate, True) & " "
               'Else
               'end 2015/3/3
                  strDesc = strDesc & " + NTD 300 "
               'End If 'Added by Morgan 2015/3/3
            End If
            'end 2013/2/4
            strDesc = strDesc & ")"
         Else
            'Added by Morgan 2015/3/3
            'If m_iPrintCurrType = 3 Or m_iPrintCurrType = 4 Then
            '   strDesc = strDesc & " - " & m_DNCurr & " " & PUB_ChgFormat(300 / m_DNRate, True)
            'Else
            'end 2015/3/3
               strDesc = strDesc & " - NTD 300"
            'End If 'Added by Morgan 2015/3/3
         End If
      End If
      
      If InStr(strDesc, " x ") > 0 Or InStr(strDesc, " + ") Or InStr(strDesc, " - ") > 0 Then
         If InStr(strDesc, "(") > 0 Then
            strDesc = "[" & strDesc & "]"
         Else
            strDesc = "(" & strDesc & ")"
         End If
      End If
      'end 2013/2/20
      End With
   End If
   GetNewDescFCT10199 = strDesc

   Set adoRst = Nothing
End Function

'Add By Cheng 2003/12/11
'Modify by Morgan 2010/10/13 +strFrom,strTo ½Ð´Ú¶µ¦¸°_¨´
Private Function GetNewDescription(stra1l01 As String, stra1l04 As String, Optional strFrom As String, Optional strTo As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strDisc As String '§é¦©ª÷ÃB
Dim arrDesc() As String
Dim ii As Integer
Dim jj As Integer
Dim strAmt As String
'Add by Morgan 2010/10/13
Dim strCon As String

If strFrom <> "" Then
   strCon = strCon & " and a1l02>='" & strFrom & "'"
End If
If strTo <> "" Then
   strCon = strCon & " and a1l02<='" & strTo & "'"
End If
'end 2010/10/13
GetNewDescription = ""
'Add By Cheng 2004/05/12
StrSQLa = "Select Count(*) From acc1l0 Where a1l01='" & stra1l01 & "' And a1l04='" & stra1l04 & "' " & strCon
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'­Y¥u¦³¤@µ§
If Val("" & rsA.Fields(0).Value) <= 1 Then
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    '§PÂ_¬O§_¦³§é¦©
    StrSQLa = "Select Nvl(a1l05, 0), Nvl(a1l07, 0), Count(*) From acc1l0 Where a1l01='" & stra1l01 & "' And a1l04='" & stra1l04 & "'" & strCon & " Group By Nvl(a1l05, 0), Nvl(a1l07, 0) Order By 2 Desc, 1 Desc "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        If Val("" & rsA.Fields(1).Value) <= 0 Then
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            Exit Function
        End If
    Else
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        Exit Function
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
'End
'Modify by Morgan 2008/1/3
'StrSQLa = "Select Nvl(a1l05, 0), Nvl(a1l07, 0), Count(*) From acc1l0 Where a1l01='" & stra1l01 & "' And a1l04='" & stra1l04 & "' Group By Nvl(a1l05, 0), Nvl(a1l07, 0) Order By 2 Desc, 1 Desc "
'Modify by Morgan 2008/3/28 »P¥~°Ó¦P¤¯½T»{«á§ï¬°¨Ì¶µ¦¸¶¶§Ç¦C¦L
'StrSQLa = "Select Nvl(a1l05, 0), Nvl(a1l07, 0), Count(*) From acc1l0 Where a1l01='" & stra1l01 & "' And a1l04='" & stra1l04 & "' Group By Nvl(a1l05, 0), Nvl(a1l07, 0) Order By 3 asc,2 desc, 1 desc"
StrSQLa = "Select Nvl(a1l05, 0), Nvl(a1l07, 0), Count(*),min(a1l02) From acc1l0 Where a1l01='" & stra1l01 & "' And a1l04='" & stra1l04 & "'" & strCon & " Group By Nvl(a1l05, 0), Nvl(a1l07, 0) Order By 4 asc"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    ii = 0
    ReDim arrDesc(ii)
    strAmt = "" & rsA.Fields(0).Value
    strDisc = "" & rsA.Fields(1).Value
    'Add By Sindy 2013/7/25 ÀË¬d¬O§_¦³98¥N²z¤HªA°È¶O
    StrSQLa = "Select Nvl(a1l05, 0), Nvl(a1l07, 0), Count(*),min(a1l02) From acc1l0 Where a1l01='" & stra1l01 & "' And a1l04='" & Trim(stra1l04) & "98'" & strCon & " Group By Nvl(a1l05, 0), Nvl(a1l07, 0) Order By 4 asc"
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, StrSQLa)
    If intI = 1 Then
      If Not IsNull(RsTemp(0)) Then
         strAmt = strAmt + Val("" & RsTemp(0))
         strDisc = strDisc + Val("" & RsTemp(1))
      End If
    End If
    '2013/7/25 END
    
    While Not rsA.EOF
        '­Y§é¦©ª÷ÃB¤£¦P
        If strDisc <> "" & rsA.Fields(1).Value Then
            arrDesc(ii) = Left(arrDesc(ii), Len(arrDesc(ii)) - 3)
            If strDisc <> "0" Then
                'Modified by Morgan 2013/10/31 °t¦X·s®æ¦¡´î¤ÖªÅ¥Õ¥HÁ×§K®e©ö¸õ¦æ
                'arrDesc(ii) = " ( " & arrDesc(ii) & " x " & Format((1 - (Val(strDisc) / Val(strAmt))) * 100, DAmount) & "% ) "
                arrDesc(ii) = "(" & arrDesc(ii) & " x " & Format((1 - (Val(strDisc) / Val(strAmt))) * 100, DAmount) & "%)"
            End If
            
            strAmt = "" & rsA.Fields(0).Value 'Add by Morgan 2005/2/15 ª÷ÃB¤]­n­«§ì¡A§_«h§é¦©·|ºâ¿ù
            strDisc = "" & rsA.Fields(1).Value
            ii = ii + 1
            ReDim Preserve arrDesc(ii)
            'Modify By Cheng 2004/04/26
'            arrDesc(iI) = arrDesc(iI) & " NTD " & Format(rsA.Fields(0).Value, FDollar) & IIf(rsA.Fields(2).Value = 1, "", " x " & rsA.Fields(2).Value) & " + "
            'Modify by Morgan 2004/7/28 ¥h±¼NTD«e­±ªÅ¥Õ
            'Added by Morgan 2015/3/3
            'If m_iPrintCurrType = 3 Or m_iPrintCurrType = 4 Then
            '   arrDesc(ii) = arrDesc(ii) & m_DNCurr & " " & PUB_ChgFormat(rsA.Fields(0).Value / m_DNRate, True) & IIf(rsA.Fields(2).Value = 1, "", " x " & rsA.Fields(2).Value) & " + "
            'Else
            'end 2015/3/3
               arrDesc(ii) = arrDesc(ii) & "NTD " & PUB_ChgFormat(rsA.Fields(0).Value, True) & IIf(rsA.Fields(2).Value = 1, "", " x " & rsA.Fields(2).Value) & " + "
            'End If 'Added by Morgan 2015/3/3
            'End
        '­Y§é¦©ª÷ÃB¬Û¦P
        Else
            'Modify By Cheng 2004/04/26
'            arrDesc(iI) = arrDesc(iI) & " NTD " & Format(rsA.Fields(0).Value, FDollar) & IIf(rsA.Fields(2).Value = 1, "", " x " & rsA.Fields(2).Value) & " + "
            'Modify by Morgan 2004/7/28 ¥h±¼NTD«e­±ªÅ¥Õ
            'Modify By Sindy 2013/7/25 strAmt:¦³¥i¯à§t98¥N²z¤HªA°È¶O
            'Modified by Morgan 2013/7/30 §ï§PÂ_²Ä¤@µ§,¦]¬°¨S¦³§é¦©©Î¬Û¦P§é¦©®É ii ¤£ÅÜ·|¾É­P«á­±ªº¶µ¥Ø¤]³£¥Î²Ä¤@µ§ªºª÷ÃB
            'If ii = 0 Then
            If rsA.AbsolutePosition = 1 Then
            'end 2013/7/30
               'Added by Morgan 2015/3/3
               'If m_iPrintCurrType = 3 Or m_iPrintCurrType = 4 Then
               '   arrDesc(ii) = arrDesc(ii) & m_DNCurr & " " & PUB_ChgFormat(strAmt / m_DNRate, True) & IIf(rsA.Fields(2).Value = 1, "", " x " & rsA.Fields(2).Value) & " + "
               'Else
               'end 2015/3/3
                  arrDesc(ii) = arrDesc(ii) & "NTD " & PUB_ChgFormat(strAmt, True) & IIf(rsA.Fields(2).Value = 1, "", " x " & rsA.Fields(2).Value) & " + "
               'End If 'Added by Morgan 2015/3/3
            Else
            '2013/7/25 END
               'Added by Morgan 2015/3/3
               'If m_iPrintCurrType = 3 Or m_iPrintCurrType = 4 Then
               '   arrDesc(ii) = arrDesc(ii) & m_DNCurr & " " & PUB_ChgFormat(rsA.Fields(0).Value / m_DNRate, True) & IIf(rsA.Fields(2).Value = 1, "", " x " & rsA.Fields(2).Value) & " + "
               'Else
               'end 2015/3/3
                  arrDesc(ii) = arrDesc(ii) & "NTD " & PUB_ChgFormat(rsA.Fields(0).Value, True) & IIf(rsA.Fields(2).Value = 1, "", " x " & rsA.Fields(2).Value) & " + "
               'End If 'Added by Morgan 2015/3/3
            End If
            'End
        End If
        rsA.MoveNext
    Wend
    arrDesc(ii) = Left(arrDesc(ii), Len(arrDesc(ii)) - 3)
    If strDisc <> "0" Then
        'Modified by Morgan 2013/10/31 °t¦X·s®æ¦¡´î¤ÖªÅ¥Õ¥HÁ×§K®e©ö¸õ¦æ
        'arrDesc(ii) = " ( " & arrDesc(ii) & " x " & Format((1 - (Val(strDisc) / Val(strAmt))) * 100, DAmount) & "% ) "
        arrDesc(ii) = "(" & arrDesc(ii) & " x " & Format((1 - (Val(strDisc) / Val(strAmt))) * 100, DAmount) & "%)"
    End If
    For jj = 0 To ii
        GetNewDescription = GetNewDescription & arrDesc(jj) & " + "
    Next jj
    '93.6.6 modify by sonia
    'If GetNewDescription <> "" Then
    If GetNewDescription <> "" And ii = 0 And strDisc <> "0" Then
        GetNewDescription = Left(GetNewDescription, Len(GetNewDescription) - 3)
    Else
    '93.6.6 end
      'Modified by Morgan 2013/10/31 °t¦X·s®æ¦¡´î¤ÖªÅ¥Õ¥HÁ×§K®e©ö¸õ¦æ,¥~¼h§ï¥Î¤¤¬A©·
      'GetNewDescription = " ( " & Left(GetNewDescription, Len(GetNewDescription) - 3) & " ) "
      If InStr(GetNewDescription, "(") > 0 Then
         GetNewDescription = "[" & Left(GetNewDescription, Len(GetNewDescription) - 3) & "]"
      Else
         GetNewDescription = "(" & Left(GetNewDescription, Len(GetNewDescription) - 3) & ")"
      End If
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Add By Cheng 2004/05/13
'¨ú±o®×¥ó¦WºÙ
Private Function GetCasePropertyName(strCP60 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'Add by Amy 2017/01/11
Dim strMCTF As String 'MCTF¯S®í¤H­ûlist
Dim strCU13(0) As String '«È¤áÀÉ´¼Åv¤H­û

GetCasePropertyName = ""
'Modified by Morgan 2021/7/27 +¦¬¤å¸¹±Æ§Ç Ex:X11009884(T-234948)·|¨ì¶W¶µ¶O--®Û­^
StrSQLa = "Select * From CaseProgress, CasePropertyMap Where CPM01=CP01 And CPM02=CP10 And CP60='" & strCP60 & "' order by cp09 asc"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   Select Case strLanguage
     Case "1"
         GetCasePropertyName = "" & rsA("CPM03").Value
     Case "2"
         GetCasePropertyName = "" & rsA("CPM10").Value
     Case "3"
         GetCasePropertyName = "" & rsA("CPM13").Value
   End Select
   
   'Add by Morgan 2006/11/22 ³¡¤À®×¥ó©Ê½è»Ý§ì«D'C'Ãþ¬ÛÃöÁ`¦¬¤å¸¹ªº®×¥ó©Ê½è
   If InStr("303,306,307,612,613", rsA.Fields("CP10")) > 0 And Not IsNull(rsA.Fields("CP43")) Then
      Do While Not IsNull(rsA.Fields("CP43"))
         StrSQLa = "Select * From CaseProgress, CasePropertyMap Where CPM01=CP01 And CPM02=CP10 And CP09='" & rsA.Fields("CP43") & "'"
         rsA.Close
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            If rsA.Fields("CP09") < "C" Then
               Select Case strLanguage
                 Case "1"
                     GetCasePropertyName = "" & rsA("CPM03").Value & GetCasePropertyName
                 Case "2"
                     GetCasePropertyName = "" & rsA("CPM10").Value & GetCasePropertyName
                 Case "3"
                     GetCasePropertyName = "" & rsA("CPM13").Value & GetCasePropertyName
               End Select
               Exit Do
            End If
         Else
            Exit Do
         End If
      Loop
   End If
   'end 2006/11/22
   'Add By Sindy 2012/11/27 ªL®Û­^:¥Ó½Ð¤£¿ìµ²®×½Ð´Ú¼ÐÃD­n©T©w
   'modify by sonia 2016/12/7 ¥¨¨Ê¦¬¤å®×¥ó¤]­n
   'Modify by Amy 2017/01/11 +´¼Åv¤H­ûMCFT¶}ÀY±±¨î,¨Ã§ï¨t²Î§O¬°T¦rÀY
   strMCTF = Pub_GetSpecMan("MCTF", True)
   strExc(0) = GetCusORFagentData(m_TM23, "CU13", strCU13())
   If Left(rsA("CP01").Value, 1) = "T" And Not IsNull(rsA("CP57").Value) And _
      rsA("CP10").Value = "101" And (rsA("CP13").Value = "67002" Or rsA("CP13").Value = "96029" Or rsA("CP13").Value = "96030" Or Left(strCU13(0), 4) = "MCTF") Then
      GetCasePropertyName = "µù¥U¥Ó½Ð-ªñ¦üµ²®×¤£¿ì"
   End If
   '2012/11/27 End
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

Private Sub LoadHeadPic()
   Dim strPicFileName As String
   Dim iNo As Integer
   
   strPicFileName = App.path & "\$TmpHead.jpg"
   
   'Added by Morgan 2020/3/31
   If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é Then
      PUB_GetLetterPicID "2", , iNo, , , , , True
   Else
   'end 2020/3/31
   
      'Modify by Morgan 2011/8/11 §ï¥Î§ó²M´·ªº¹ÏÀÉ(¸û¤j)
      'iNo = 6
      iNo = 16
      'Modify by Morgan 2008/5/26 +¤j³°«HÀY
      If m_A1k28 <> "" Then
         If Left(m_A1k28, 1) = "X" Then
            strExc(0) = "SELECT CU10 FROM CUSTOMER WHERE CU01='" & Left(m_A1k28, 8) & "' and CU02='" & Mid(m_A1k28, 9) & "'"
         Else
            strExc(0) = "SELECT FA10 FROM FAGENT WHERE FA01='" & Left(m_A1k28, 8) & "' and FA02='" & Mid(m_A1k28, 9) & "'"
         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp(0) = "020" Then
               iNo = 8
            End If
         End If
      End If
      'End If
   End If
   
   If PUB_ReadDB2File(strPicFileName, iNo) = False Then
      Exit Sub
   End If
   
   Set Picture1.Picture = LoadPicture(strPicFileName)
   Picture1.AutoSize = True
End Sub

Private Sub NewPic()
   Dim strPicFileName As String
   strPicFileName = App.path & "\$tmp_*.jpg"
   If Dir(strPicFileName) <> "" Then
      Kill strPicFileName
   End If
   LoadHeadPic
   Picture1.AutoRedraw = True
   'Picture1.Height = 16836
   'Picture1.Width = 11904
   '¹w³]¦C¦L¦r«¬
   douExtRate = Picture1.Height / 16836
   Picture1.FontSize = 12 * douExtRate
   Picture1.Font.Name = "Times New Roman"
End Sub

Private Function GetPath()
   Dim strSubDir As String
   'Add by Morgan 2009/4/13
   '«ü©w¦sÀÉ¸ô®|
   If m_SavePath <> "" Then
      GetPath = m_SavePath
   Else
      'Modify by Morgan 2008/5/23 ¿é¥X¤è¦¡¿ï¹q¤lÀÉ®É¦s®à­±
      If txtOutMode = "2" Then
         GetPath = PUB_Getdesktop
      Else
         'Modify by Morgan 2011/1/6 ¥[¤W¼h¥Ø¿ý=¨t²Î§O+¥»©Ò¸¹«e2½X
         'GetPath = PUB_GetEFilePath(m_CP01) & "\" & m_strCaseNo
         'Modified by Morgan 2012/4/2 §ï¬° ¤W¼h¥Ø¿ý=¨t²Î§O\¥»©Ò¸¹«e3½X\
         'strSubDir = PUB_GetEFilePath(m_CP01) & "\" & m_CP01 & Left(m_CP02, 2)
         strSubDir = PUB_GetEFilePath(m_CP01) & "\" & m_CP01
         If Dir(strSubDir, vbDirectory) = "" Then
            MkDir strSubDir
         End If
         'Added by Morgan 2012/4/2
         strSubDir = strSubDir & "\" & Left(m_CP02, 3)
         If Dir(strSubDir, vbDirectory) = "" Then
            MkDir strSubDir
         End If
         'end 2012/4/2
         GetPath = strSubDir & "\" & m_strCaseNo
      End If
   End If
   
End Function
Private Sub SetPic(idx As Integer, Optional pIsFinal As Boolean)
   Dim strPicFileName As String
   Dim strPicFileName1 As String
   Dim objImg As StdPicture
   Dim m_Image As New cImage
   Dim m_Jpeg  As New cJpeg
   Dim strFolder As String
   Dim strFileName As String
   Dim bol2Pdf As Boolean
   
   m_EFilePath = GetPath
   
   strFolder = m_EFilePath
   
   If Dir(strFolder, vbDirectory) = "" Then
      MkDir strFolder
   End If
   
   strFileName = m_strCaseNo & IIf(m_bAddDate, "_" & strSrvDate(1), "") & "_DN" & m_strDN
   
   'Modify by Morgan 2011/8/9
   ' strPicFileName1 = strFolder & "\" & strFileName
   If pub_PdfEnable Then
      strPicFileName1 = ".\" & strFileName
   Else
      strPicFileName1 = strFolder & "\" & strFileName
   End If
   'end 2011/8/9
   
   If idx > 0 Then
      strPicFileName1 = strPicFileName1 & "_P" & idx
   End If
   strPicFileName1 = strPicFileName1 & ".jpg"
   
   RidFile strPicFileName1
   PUB_SavePic Picture1, strPicFileName1
   
   'Add by Morgan 2011/8/9 ÂàPDF
   If pIsFinal And pub_PdfEnable Then
      If idx > 0 Then
         PUB_Trans2Pdf ".\" & strFileName & "*.jpg", strFileName & ".pdf", True, strFolder
      Else
         PUB_Trans2Pdf ".\" & strFileName & ".jpg", strFileName & ".pdf", False, strFolder
      End If
   End If
   'end 2011/8/9
   
   '­n¥ÎÂÐ»\ªº§_«h·|¿ù»~--VB Bug
   'Picture1.Cls
   'Picture1.Line (0, 0)-(Picture1.Width, Picture1.Height), QBColor(15), BF
   LoadHeadPic
   
End Sub

'Add by Morgan 2008/3/13
'ÀË¬d¦C¦L¥÷¼Æªº¦C¥~ª¬ªp
'Modify by Morgan 2009/9/25 +°ò¥»ÀÉ¤Î«È¤áÀÉ¤]¥i³]¥÷¼Æ
Private Function GetSpecificCopy(ByVal p_iCopy As Integer, ByVal A1k03 As String, ByVal A1k13 As String, ByVal A1k14 As String, ByVal A1k15 As String, ByVal A1k16 As String, Optional ByVal bEMail As Boolean = False) As Integer
   Dim iCopy As Integer, stSQL As String, iR As Integer, iKind As Integer
   Dim stCustNo As String 'Added by Morgan 2014/6/27
   
   'iCopy = p_iCopy 'Removed by Morgan 2014/6/27
   
   'Modify by Morgan 2010/9/6 +S
   'Modified by Morgan 2013/6/5 §ï¤£­­¨î¨t²Î§O(Ex:FMP)
   'If A1k13 = "FCP" Or A1k13 = "FG" Or A1k13 = "FCT" Or A1k13 = "S" Then
      
      If ClsPDGetSystemKind(A1k13, iKind) Then
         'Add by Morgan 2009/9/25
         Select Case iKind
            Case 1
               stSQL = "select pa154,pa155,pa26 from patent where pa01='" & A1k13 & "' and pa02='" & A1k14 & "' and pa03='" & A1k15 & "' and pa04='" & A1k16 & "'"
            Case 2
               stSQL = "select tm125,tm126,tm23 from trademark where tm01='" & A1k13 & "' and tm02='" & A1k14 & "' and tm03='" & A1k15 & "' and tm04='" & A1k16 & "'"
            Case Else
               stSQL = "select sp82,sp83,sp08 from servicepractice where sp01='" & A1k13 & "' and sp02='" & A1k14 & "' and sp03='" & A1k15 & "' and sp04='" & A1k16 & "'"
         End Select
         iR = 1
         Set RsTemp = ClsLawReadRstMsg(iR, stSQL)
         If iR = 1 Then
            stCustNo = "" & RsTemp.Fields(2) 'Added by Morgan 2014/6/27
            If RsTemp.Fields(0) > 0 Then
               iCopy = Val("" & RsTemp.Fields(0))
               
            'Added by Morgan 2022/8/18
            ElseIf bEMail Then
               iCopy = 1
            'end 2022/8/18
            
            Else
         'end 2009/9/25
               If Left(A1k03, 1) = "Y" Then
                  'Add by Morgan 2009/10/14
                  '¥N²z¤H¬°Y52013(Tsujimaru)¥B¥Ó½Ð¤H¬°X47660(¤éªF¹q¤u)®É,¨ä¹ï¥~©w½Z«H¨ç¤Î½Ð´Ú³æ§¡¦C¦L4¥÷
                  If A1k13 = "FCP" And A1k03 = "Y52013000" And RsTemp.Fields(2) = "X47660000" Then
                     iCopy = 4
                  Else
                  'end 2009/10/14
                     'Modified by Morgan 2018/3/16 FCT®×¤é¥»¥N²z¤H½Ð´Ú³æ¹w³]¦L2¥÷ +fa10 --³¯ª÷½¬
                     stSQL = "select decode('" & A1k13 & "','FCT',fa90,'S',fa90,fa89) cpy,fa10 from fagent" & _
                        " where fa01='" & Left(A1k03, 8) & "' and fa02='" & Mid(A1k03, 9) & "'"
                     iR = 1
                     Set RsTemp = ClsLawReadRstMsg(iR, stSQL)
                     If iR = 1 Then
                        If Val("" & RsTemp.Fields(0)) > 0 Then
                           iCopy = Val("" & RsTemp.Fields(0))
                        'Added by Morgan 2018/3/16
                        'Modified by Morgan 2018/5/14 +S®×¹w³]1¥÷ --³¯ª÷½¬
                        ElseIf (A1k13 = "FCT" Or A1k13 = "S") And Left(RsTemp("fa10"), 3) = "011" Then
                           '§ï¹w³]¥÷¼Æ,¯S®í³]©w¤´Àu¥ý
                           If p_iCopy = 3 Then
                              If A1k13 = "S" Then
                                 p_iCopy = 1
                              Else
                                 p_iCopy = 2
                              End If
                           End If
                        'end 2018/3/16
                        End If
                     End If
                  End If
               'Add by Morgan 2009/9/25
               ElseIf Left(A1k03, 1) = "X" Then
                  stSQL = "select decode('" & A1k13 & "','FCT',cu136,'S',cu136,cu135) cpy from customer" & _
                     " where cu01='" & Left(A1k03, 8) & "' and cu02='" & Mid(A1k03, 9) & "'"
                  iR = 1
                  Set RsTemp = ClsLawReadRstMsg(iR, stSQL)
                  If iR = 1 Then
                     If Val("" & RsTemp.Fields(0)) > 0 Then
                        iCopy = Val("" & RsTemp.Fields(0))
                     End If
                  End If
               End If
               
               'Added by Morgan 2014/6/27
               'Modified by Morgan 2018/3/16
               'If stCustNo <> "" Then
               If iCopy = 0 And stCustNo <> "" Then
               'end 2018/3/16
                  stSQL = "select decode('" & A1k13 & "','FCT',cu136,'S',cu136,cu135) cpy from customer" & _
                     " where cu01='" & Left(stCustNo, 8) & "' and cu02='" & Mid(stCustNo, 9) & "'"
                  iR = 1
                  Set RsTemp = ClsLawReadRstMsg(iR, stSQL)
                  If iR = 1 Then
                     If Val("" & RsTemp.Fields(0)) > 0 Then
                        iCopy = Val("" & RsTemp.Fields(0))
                     End If
                  End If
               End If
               'end 2014/6/27
               
            End If
         End If
      End If
   'End If
   If iCopy = 0 Then iCopy = p_iCopy 'Added by Morgan 2014/6/27
   GetSpecificCopy = iCopy
End Function

'Modified by Morgan 2022/8/4 +p_FontSize
Private Sub PutData(p_sData As String, p_iRow As Integer, Optional p_lX As Double = 500, Optional p_lY As Double = 1500, Optional p_FontSize As Integer)
   If p_sData <> "" Then
      If m_b2Printer Then
         If p_FontSize > 0 Then Printer.Font.Size = p_FontSize 'Added by Morgan 2022/8/4
         Printer.CurrentX = p_lX + intDefault
         Printer.CurrentY = p_lY + p_iRow * m_LineH + intTop
         'Modified by Morgan 2022/8/4
         'Printer.Print p_sData
         PUB_PrintUnicodeText p_sData, Printer.CurrentX, Printer.CurrentY, 0
         'end 2022/8/4
      End If
      'Removed by Morgan 2022/8/4
      'If m_b2Picture Then
      '   Picture1.CurrentX = (p_lX + intDefault) * douExtRate
      '   Picture1.CurrentY = (p_lY + p_iRow * m_LineH + intTop) * douExtRate
      '   Picture1.Print p_sData
      'End If
      'end 2022/8/4
   End If
End Sub
'Add by Morgan 2010/11/5
'¬O§_¥Î¹q¤l±b³æ
Private Function SetEBilling(p_A1k28 As String) As Boolean
   Dim stSQL As String
   Dim iR As Integer
   Dim RsTemp  As ADODB.Recordset
   
   If Left(p_A1k28, 1) = "Y" Then
      stSQL = "select fa102,fa01||fa02 from fagent where fa01='" & Left(p_A1k28, 8) & "' and fa02='" & Mid(p_A1k28, 9) & "'"
   Else
      stSQL = "select cu141,cu01||cu02 from customer where cu01='" & Left(p_A1k28, 8) & "' and cu02='" & Mid(p_A1k28, 9) & "'"
   End If
   iR = 1
   Set RsTemp = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      If RsTemp.Fields(0) = "Y" Then
         SetEBilling = True
      End If
   End If
   Set RsTemp = Nothing
End Function

'Added by Morgan 2015/10/3
Private Function GetCustCaseNo(p_a1k13 As String, p_a1k14 As String, p_a1k15 As String, p_a1k16 As String) As String
   Dim stSQL As String
   Dim iR As Integer
   Dim rsQuery  As ADODB.Recordset
   
   stSQL = "select pa48 from patent where pa01='" & p_a1k13 & "' and pa02='" & p_a1k14 & "'" & _
         " and pa03='" & p_a1k15 & "' and pa04='" & p_a1k16 & "'"
   iR = 1
   Set rsQuery = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      GetCustCaseNo = "" & rsQuery.Fields(0)
   End If
   Set rsQuery = Nothing
End Function

'Add by Morgan 2010/11/5
Private Function GetClientMatterID(p_a1k13 As String, p_a1k14 As String, p_a1k15 As String, p_a1k16 As String, p_A1K01 As String, Optional FAgentNo As String) As String
   Dim stSQL As String
   Dim iR As Integer
   Dim rsQuery  As ADODB.Recordset
   Dim strRefNo As String
   
   'Added by Morgan 2013/6/27 HP§ì«È¤á®×¥ó®×¸¹
   'Modified by Morgan 2015/10/14 +Y
   'Modified by Morgan 2016/3/17 +Lawcase(¥Ø«e¥u¦³¤@®×½Ð´ÚÄæ¦ìÁÙ¨S·s¼W,¥ý¤â°Ê»s§@)
   'Modified by Morgan 2020/4/20 Lawcase +LC51
   'Modified by Morgan 2023/7/4 +Y54570000
   'Modified by Morgan 2024/8/1 +X21660000 --Lisa
   If FAgentNo = "Y48292000" Or FAgentNo = "Y54332000" Or FAgentNo = "Y54570000" Or FAgentNo = "X21660000" Then
      stSQL = "select nvl(pa159,pa48) from patent where pa01='" & p_a1k13 & "' and pa02='" & p_a1k14 & "'" & _
         " and pa03='" & p_a1k15 & "' and pa04='" & p_a1k16 & "'" & _
         " union select nvl(tm127,tm35) from trademark where tm01='" & p_a1k13 & "' and tm02='" & p_a1k14 & "'" & _
         " and tm03='" & p_a1k15 & "' and tm04='" & p_a1k16 & "'" & _
         " union select nvl(sp84,sp29) from servicepractice where sp01='" & p_a1k13 & "' and sp02='" & p_a1k14 & "'" & _
         " and sp03='" & p_a1k15 & "' and sp04='" & p_a1k16 & "'" & _
         " union select nvl(lc51,lc17) from lawcase where lc01='" & p_a1k13 & "' and lc02='" & p_a1k14 & "'" & _
         " and lc03='" & p_a1k15 & "' and lc04='" & p_a1k16 & "'"
         
      
      
   'Added by Morgan 2017/3/30
   'Y54225B10ªº±M§Q®×¹w³] Matter Number ¬°SYN2017005730
   'Modified by Morgan 2017/5/2 Y48309070ªº°Ó¼Ð®×¤]¹w³] Matter Number ¬°SYN2017005730
   'ElseIf FAgentNo = "Y54225B10" And (p_a1k13 = "FCP" Or p_a1k13 = "FG" Or p_a1k13 = "P" Or p_a1k13 = "PS" Or p_a1k13 = "CFP" Or p_a1k13 = "CPS") Then
   '   stSQL = "select nvl(pa159,'SYN2017005730') from patent where pa01='" & p_a1k13 & "' and pa02='" & p_a1k14 & "'" & _
         " and pa03='" & p_a1k15 & "' and pa04='" & p_a1k16 & "'" & _
         " union select nvl(sp84,'SYN2017005730') from servicepractice where sp01='" & p_a1k13 & "' and sp02='" & p_a1k14 & "'" & _
         " and sp03='" & p_a1k15 & "' and sp04='" & p_a1k16 & "'"
   'Modified by Morgan 2019/1/10 Y48309070 108/1/1°_§ï¬° SYN2018006894(­ì¬° SYN2017005730 »P Y54225B10 ¦P)
   'ElseIf FAgentNo = "Y54225B10" Or FAgentNo = "Y48309070" Then
   ElseIf FAgentNo = "Y48309070" Then
      stSQL = "select nvl(pa159,'SYN2018006894') from patent where pa01='" & p_a1k13 & "' and pa02='" & p_a1k14 & "'" & _
         " and pa03='" & p_a1k15 & "' and pa04='" & p_a1k16 & "'" & _
         " union select nvl(tm127,'SYN2018006894') from trademark where tm01='" & p_a1k13 & "' and tm02='" & p_a1k14 & "'" & _
         " and tm03='" & p_a1k15 & "' and tm04='" & p_a1k16 & "'" & _
         " union select nvl(sp84,'SYN2018006894') from servicepractice where sp01='" & p_a1k13 & "' and sp02='" & p_a1k14 & "'" & _
         " and sp03='" & p_a1k15 & "' and sp04='" & p_a1k16 & "'" & _
         " union select 'SYN2018006894' from lawcase where lc01='" & p_a1k13 & "' and lc02='" & p_a1k14 & "'" & _
         " and lc03='" & p_a1k15 & "' and lc04='" & p_a1k16 & "'"
   ElseIf FAgentNo = "Y54225B10" Then
   'end 2019/1/10
      stSQL = "select nvl(pa159,'SYN2017005730') from patent where pa01='" & p_a1k13 & "' and pa02='" & p_a1k14 & "'" & _
         " and pa03='" & p_a1k15 & "' and pa04='" & p_a1k16 & "'" & _
         " union select nvl(tm127,'SYN2017005730') from trademark where tm01='" & p_a1k13 & "' and tm02='" & p_a1k14 & "'" & _
         " and tm03='" & p_a1k15 & "' and tm04='" & p_a1k16 & "'" & _
         " union select nvl(sp84,'SYN2017005730') from servicepractice where sp01='" & p_a1k13 & "' and sp02='" & p_a1k14 & "'" & _
         " and sp03='" & p_a1k15 & "' and sp04='" & p_a1k16 & "'" & _
         " union select 'SYN2017005730' from lawcase where lc01='" & p_a1k13 & "' and lc02='" & p_a1k14 & "'" & _
         " and lc03='" & p_a1k15 & "' and lc04='" & p_a1k16 & "'"
   'end 2017/3/30
   
   'Added by Morgan 2017/4/6
   ElseIf FAgentNo = "Y33611B50" And (p_a1k13 = "FCP" Or p_a1k13 = "FG" Or p_a1k13 = "P" Or p_a1k13 = "PS" Or p_a1k13 = "CFP" Or p_a1k13 = "CPS") Then
      stSQL = "select pa159 from patent where pa01='" & p_a1k13 & "' and pa02='" & p_a1k14 & "'" & _
         " and pa03='" & p_a1k15 & "' and pa04='" & p_a1k16 & "'" & _
         " union select sp84 from servicepractice where sp01='" & p_a1k13 & "' and sp02='" & p_a1k14 & "'" & _
         " and sp03='" & p_a1k15 & "' and sp04='" & p_a1k16 & "'"
   'end 2017/4/6
   
   Else
   'end 2013/6/27
            
      stSQL = "select pa159,pa77 from patent where pa01='" & p_a1k13 & "' and pa02='" & p_a1k14 & "'" & _
         " and pa03='" & p_a1k15 & "' and pa04='" & p_a1k16 & "'" & _
         " union select tm127,tm45 from trademark where tm01='" & p_a1k13 & "' and tm02='" & p_a1k14 & "'" & _
         " and tm03='" & p_a1k15 & "' and tm04='" & p_a1k16 & "'" & _
         " union select sp84,sp27 from servicepractice where sp01='" & p_a1k13 & "' and sp02='" & p_a1k14 & "'" & _
         " and sp03='" & p_a1k15 & "' and sp04='" & p_a1k16 & "'" & _
         " union select lc51,lc23 from lawcase where lc01='" & p_a1k13 & "' and lc02='" & p_a1k14 & "'" & _
         " and lc03='" & p_a1k15 & "' and lc04='" & p_a1k16 & "'"
   End If
   iR = 1
   Set rsQuery = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      strRefNo = "" & rsQuery.Fields(0)
      If strRefNo = "" And rsQuery.Fields.Count > 1 Then
         'Added by Morgan 2017/4/7 §ì©¼¸¹ªº¤~­n¦Ò¼{§ó¥Nªºª¬ªp
         'strRefNo = "" & rsQuery.Fields(1)
         strRefNo = GetYourRef(p_a1k13, p_A1K01, "" & rsQuery.Fields(1))
      End If
   End If
   GetClientMatterID = strRefNo
   
   'Removed by Morgan 2017/4/7 §ì©¼¸¹ªº¤~­n¦Ò¼{§ó¥Nªºª¬ªp
   ''Modified by Morgan 2016/1/12 ¦³¥i¯à§ó¥N¹L
   'GetClientMatterID = GetYourRef(p_a1k13, p_A1K01, strRefNo)
   
   Set rsQuery = Nothing
End Function


'Add by Morgan 2010/11/5
'Modify by Morgan 2011/2/25 +1998BI®æ¦¡
'iFormat:®æ¦¡ 1=1998B,2=1998BI
Private Function WriteLEDES(arrData() As String, Optional iFormat As Integer = 1) As Boolean
   Dim F1 As Integer, ii As Integer, jj As Integer, kk As Integer, stOut As String
   Dim strColHead As String, arrCol
   Dim strTot As String, strAmt As String
   Dim strPath As String
   Dim strColErr As String, strValErr As String
   Dim strMsg As String
   Dim stTmp As String, rr As Integer
   Dim dblEFeeAmt As Double 'LEDESªA°È¶OÁ`ÃB Added by Morgan 2018/5/16
   Dim strFTot As String, strETot As String, strFTotUS As String, strETotUS As String, idxF As Integer, idxE As Integer 'Added by Morgan 2021/5/21
   
   If iFormat = 2 Then
      strColHead = "INVOICE_DATE|INVOICE_NUMBER|CLIENT_ID|LAW_FIRM_MATTER_ID|INVOICE_TOTAL|BILLING_START_DATE|BILLING_END_DATE|INVOICE_DESCRIPTION|LINE_ITEM_NUMBER|EXP/FEE/INV_ADJ_TYPE|LINE_ITEM_NUMBER_OF_UNITS|LINE_ITEM_ADJUSTMENT_AMOUNT|LINE_ITEM_TOTAL|LINE_ITEM_DATE|LINE_ITEM_TASK_CODE|LINE_ITEM_EXPENSE_CODE|LINE_ITEM_ACTIVITY_CODE|TIMEKEEPER_ID|LINE_ITEM_DESCRIPTION|LAW_FIRM_ID|LINE_ITEM_UNIT_COST|TIMEKEEPER_NAME|TIMEKEEPER_CLASSIFICATION|CLIENT_MATTER_ID|PO_NUMBER|CLIENT_TAX_ID|MATTER_NAME|INVOICE_TAX_TOTAL|INVOICE_NET_TOTAL|INVOICE_CURRENCY|TIMEKEEPER_LAST_NAME|TIMEKEEPER_FIRST_NAME|ACCOUNT_TYPE|LAW_FIRM_NAME|LAW_FIRM_ADDRESS_1|LAW_FIRM_ADDRESS_2|LAW_FIRM_CITY|LAW_FIRM_STATEorREGION|LAW_FIRM_POSTCODE|LAW_FIRM_COUNTRY|CLIENT_NAME|CLIENT_ADDRESS_1|CLIENT_ADDRESS_2|CLIENT_CITY|CLIENT_STATEorREGION|CLIENT_POSTCODE|CLIENT_COUNTRY|LINE_ITEM_TAX_RATE|LINE_ITEM_TAX_TOTAL|LINE_ITEM_TAX_TYPE|INVOICE_REPORTED_TAX_TOTAL|INVOICE_TAX_CURRENCY"
   Else
      strColHead = "INVOICE_DATE|INVOICE_NUMBER|CLIENT_ID|LAW_FIRM_MATTER_ID|INVOICE_TOTAL|BILLING_START_DATE|BILLING_END_DATE|INVOICE_DESCRIPTION|LINE_ITEM_NUMBER|EXP/FEE/INV_ADJ_TYPE|LINE_ITEM_NUMBER_OF_UNITS|LINE_ITEM_ADJUSTMENT_AMOUNT|LINE_ITEM_TOTAL|LINE_ITEM_DATE|LINE_ITEM_TASK_CODE|LINE_ITEM_EXPENSE_CODE|LINE_ITEM_ACTIVITY_CODE|TIMEKEEPER_ID|LINE_ITEM_DESCRIPTION|LAW_FIRM_ID|LINE_ITEM_UNIT_COST|TIMEKEEPER_NAME|TIMEKEEPER_CLASSIFICATION|CLIENT_MATTER_ID"
   End If
   
   arrCol = Split(strColHead, "|")
   kk = UBound(arrData, 2)
   
   '¸ê®ÆÀË¬d
   strColErr = "": strValErr = ""
   For ii = 1 To kk
      If arrData(10, ii) = "F" Then
         If arrData(18, ii) = "" Then
            strColErr = strColErr & ",TIMEKEPPER_ID"
         End If
         If arrData(15, ii) = "" Then
            strColErr = strColErr & ",TASK_CODE"
         End If
         'Added by Morgan 2024/3/19
'         If arrData(17, ii) = "" Then
'            strColErr = strColErr & ",ACTIVITY_CODE"
'         End If
         'end 2024/3/19
      End If
      If strColErr <> "" Then Exit For
   Next
   For ii = 1 To kk
      For jj = 1 To UBound(arrData, 1)
         If InStr(arrData(jj, ii), "|") + InStr(arrData(jj, ii), "[]") > 0 Then
            strValErr = strValErr & "," & arrCol(jj - 1)
         End If
         'Added by Morgan 2012/10/12
         '°Ó¼Ð·|¦³¤¤¤å­n­ç°£(¤£¤ä´©)
         If GetTextLength(arrData(jj, ii)) <> Len(arrData(jj, ii)) Then
            stTmp = ""
            For rr = 1 To Len(arrData(jj, ii))
               If GetTextLength(Mid(arrData(jj, ii), rr, 1)) = 1 Then
                  stTmp = stTmp & Mid(arrData(jj, ii), rr, 1)
               End If
            Next
            arrData(jj, ii) = stTmp
         End If
         'end 2012/10/12
      Next
      If strValErr <> "" Then Exit For
   Next
   
   If arrData(3, 1) = "" Then
      strColErr = strColErr & ",CLIENT_ID"
   End If
   
   'Modified by Morgan 2017/8/23 X18064010 ¤£»Ý­n Client Matter ID -- ¸U§Ó¼w
   If arrData(24, 1) = "" And m_A1k28 <> "X18064010" Then
      strColErr = strColErr & ",CLIENT_MATTER_ID"
   End If
   
   If strColErr & strValErr <> "" Then
      strMsg = "½Ð´Ú³æ " & arrData(2, 1) & " ¦]¤U¦C°ÝÃDµLªk²£¥Í¹q¤l±b³æ!!" & vbCrLf
      If strColErr <> "" Then
         strMsg = strMsg & vbCrLf & vbTab & Mid(strColErr, 2) & " Äæ¦ì­È¬°¤£¥iªÅ¥Õ!!"
      End If
      If strValErr <> "" Then
         strMsg = strMsg & vbCrLf & vbTab & Mid(strValErr, 2) & " Äæ¦ì­È¤£¥i§t«O¯d²Å¸¹ | ©Î [] !!"
      End If
      m_sEBillingMsg = m_sEBillingMsg & vbCrLf & strMsg
      Exit Function
   End If
   
   strAmt = ""
   '¬üª÷½Ð´Ú®É¦X­p¬°¦U©ú²Ó¥[Á`
   'Modified by Morgan 2012/12/7
   'If strCurr = "U" Then
   If m_iPrintCurrType = 3 Then
   'end 2012/12/7
      For ii = 1 To kk
         strAmt = Val(strAmt) + Val(arrData(13, ii))
      Next
      For ii = 1 To kk
         arrData(5, ii) = Val(strAmt)
         'Add by Morgan 2011/3/14
         If iFormat = 2 Then
            arrData(29, ii) = arrData(5, ii)
         End If
      Next
   
   'Added by Morgan 2021/5/20
   '¥~¹ô+¬üª÷®É
   ElseIf m_iPrintCurrType = 4 Then
      strExc(0) = "select a1k38 from acc1k0 where a1k01='" & m_strDN & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp(0) > 0 Then
            idxF = 0: idxE = 0
            For ii = 1 To kk
               arrData(5, ii) = RsTemp(0)
               'ªA°È¶O
               If arrData(10, ii) = "F" Then
                  strFTot = Val(strFTot) + Val(arrData(13, ii))
               '³W¶O
               Else
                  strETot = Val(strETot) + Val(arrData(13, ii))
               End If
               
               arrData(13, ii) = Trunc(Val(arrData(13, ii)) * m_DUsdRate) 'LINE_ITEM_TOTAL
               arrData(12, ii) = Trunc(Val(arrData(12, ii)) * m_DUsdRate) 'LINE_ITEM_ADJUSTMENT_AMOUNT
               arrData(21, ii) = (Val(arrData(13, ii)) - Val(arrData(12, ii))) / Val(arrData(11, ii))  'LINE_ITEM_UNIT_COST
               
               'ªA°È¶O(¬üª÷)
               If arrData(10, ii) = "F" Then
                  strFTotUS = Val(strFTotUS) + Val(arrData(13, ii))
                  idxF = ii
               '³W¶O(¬üª÷)
               Else
                  strETotUS = Val(strETotUS) + Val(arrData(13, ii))
                  idxE = ii
               End If
            Next
            
            strFTot = Trunc(Val(strFTot) * m_DUsdRate)
            If idxF > 0 And Val(strFTotUS) <> Val(strFTot) Then
               arrData(13, idxF) = Val(strFTot) - (Val(strFTotUS) - Val(arrData(13, idxF)))  'LINE_ITEM_TOTAL
               arrData(21, idxF) = (Val(arrData(13, idxF)) - Val(arrData(12, idxF))) / Val(arrData(11, idxF))  'LINE_ITEM_UNIT_COST
            End If
            
            strETot = Trunc(Val(strETot) * m_DUsdRate)
            If idxE > 0 And Val(strETotUS) <> Val(strETot) Then
               arrData(13, idxE) = Val(strETot) - (Val(strETotUS) - Val(arrData(13, idxE))) 'LINE_ITEM_TOTAL
               arrData(21, idxE) = (Val(arrData(13, idxE)) - Val(arrData(12, idxE))) / Val(arrData(11, idxE))  'LINE_ITEM_UNIT_COST
            End If
         End If
      End If
   'end 2021/5/20
      
   '¥x¹ô½Ð´Ú®É¦X­p»P©ú²Ó­Y¦³®tÃB®É½Õ¾ã¨ì³Ì«á¶µ¦¸
   Else
      strTot = arrData(5, 1)
      For ii = 1 To kk - 1
         strAmt = Val(strAmt) + Val(arrData(13, ii))
      Next
      arrData(13, kk) = Val(strTot) - Val(strAmt)
      If arrData(10, kk) = "F" Then
         'Modified by Morgan 2014/8/5
         'arrData(11, kk) = Round((Val(arrData(13, kk)) - Val(arrData(12, kk))) / Val(arrData(21, kk)), 4)
         'Rockwell ³æ»ù§ï20,¤u®É·|­pºâ¥X¤p¼Æ¨â¦ì¦Ó³y¦¨¤W¶Ç¿ù»~,¬G§ï¬°¥|±Ë¤­¤J¨ì¤p¼Æ1¦ì¦A¥H¦¹­«ºâ½Ð´Úª÷ÃB --Elvan
         If m_A1k28 = "Y46295000" Then
            arrData(11, kk) = Round((Val(arrData(13, kk)) - Val(arrData(12, kk))) / Val(arrData(21, kk)), 1)
            arrData(13, kk) = Val(arrData(11, kk)) * Val(arrData(21, kk)) + Val(arrData(12, kk))
            strAmt = "" 'Added by Morgan 2017/10/18
            For ii = 1 To kk
               strAmt = Val(strAmt) + Val(arrData(13, ii))
            Next
            For ii = 1 To kk
               arrData(5, ii) = Val(strAmt)
               If iFormat = 2 Then
                  arrData(29, ii) = arrData(5, ii)
               End If
            Next
         Else
            'Added by Morgan 2022/1/13 ¥~°Ó§ï¶O²v(³æ¦ìºû«ù¥Î1)
            If InStr(arrData(4, kk), "FCT-") = 1 Or InStr(arrData(4, kk), "S-") = 1 Or InStr(arrData(4, kk), "T-") = 1 Then
                arrData(21, kk) = Val(arrData(13, kk)) - Val(arrData(12, kk))
            Else
            'end 2022/1/13
               arrData(11, kk) = Round((Val(arrData(13, kk)) - Val(arrData(12, kk))) / Val(arrData(21, kk)), 4)
            End If
         End If
         'end 2014/8/5
      Else
         arrData(21, kk) = Val(arrData(13, kk)) - Val(arrData(12, kk))
      End If
   End If
   
   'Added by Morgan 2018/5/16 Y53942 Xperi Corporation ¦³ "ªA°È¶OÁ`ÃB x 0.5%" ªº¥­¥x§é
   If m_A1k28 = "Y53942000" Then
      For ii = 1 To kk
         If arrData(10, ii) = "F" Then
            dblEFeeAmt = dblEFeeAmt + Val(arrData(13, ii))
         End If
      Next
      If dblEFeeAmt > 0 Then
         strAmt = Round(-1 * dblEFeeAmt * 0.005, 2)
         strLedes(5, 1) = Round(Val(strLedes(5, 1)) + Val(strAmt), 2)
         For ii = 2 To kk
            strLedes(5, ii) = strLedes(5, 1)
         Next
         
         iUpper = iUpper + 1
         ReDim Preserve strLedes(m_iCols, iUpper)
         '1~8
         For intI = 1 To 8
            strLedes(intI, iUpper) = strLedes(intI, iUpper - 1)
         Next
         strLedes(9, iUpper) = iUpper
         strLedes(10, iUpper) = "IF"
         strLedes(12, iUpper) = strAmt
         strLedes(13, iUpper) = strLedes(12, iUpper)
         strLedes(14, iUpper) = strLedes(14, iUpper - 1)
         strLedes(19, iUpper) = "E-billing Adjustment Discount"
         strLedes(20, iUpper) = strLedes(20, iUpper - 1)
         strLedes(24, iUpper) = strLedes(24, iUpper - 1)
         kk = iUpper
      End If
   End If
   'end 2018/5/16
   
   'Added by Morgan 2018/6/27
   If m_bLedesOnly Then
      strLedes(1, 0) = iFormat 'LEDESª©¥»
      WriteLEDES = True
      Exit Function
   End If
   'end2018/6/27
   
   m_EFilePath = GetPath
   If Dir(m_EFilePath, vbDirectory) = "" Then
      MkDir m_EFilePath
   End If
   strPath = m_EFilePath & "\" & m_strCaseNo & "_DN" & m_strDN & ".txt"
   
   F1 = FreeFile
   Open strPath For Output As F1
   
   If iFormat = 2 Then
      Print #F1, "LEDES98BI V2[]"
   Else
      Print #F1, "LEDES1998B[]"
   End If
   
   Print #F1, strColHead & "[]"
   For ii = 1 To kk
      stOut = arrData(1, ii)
      For jj = 2 To UBound(arrData, 1)
         stOut = stOut & "|" & arrData(jj, ii)
      Next
      stOut = stOut & "[]"
      Print #F1, stOut
   Next
   Close #F1
   WriteLEDES = True
End Function

'Modify by Morgan 2011/1/5 ¥[¥i¦C¦L¤Î¦sÀÉ
Private Sub runWordChinese()

   'Added by Morgan 2013/10/31
   If bolNewForm Then
      runWordChineseNew
      Exit Sub
   End If
   'end 2013/10/31
   
   Dim stTmp As String
   Dim iRow As Integer, iCol As Integer, lngCellWidth As Long
   Dim strLastCountry As String, iFCol As Integer, iCols As Integer, bolPrintCountry As Boolean
   Dim strFontSize As String
   Dim iResumeCnt As Integer
   Dim iPos1 As Integer, iPos2 As Integer
   Dim bVisible As Boolean, stFileName As String, iHeadCount As Integer
   Dim oShape
   Dim ii As Integer 'Added by Morgan 2014/3/4
   Dim dblOffFeeSub As Double
   Dim dblAttFeeSub As Double
   
   strFontSize = 12
   
   'Modified by Lydia 2019/04/09 §ï¦¨¦@¥Î¼Ò²Õ
   'If NewDoc(bVisible, m_WordLeft, m_WordTop) = False Then Exit Sub 'Added by Morgan 2014/6/26
   If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = False Then Exit Sub
   
On Error GoTo ErrHnd

'Removed by Morgan 2014/6/26
'   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
'
'   g_WordAp.Documents.Add
'
'   bVisible = g_WordAp.Visible
'
'   '¤£Åã¥Ü¥i¯à·|¦³°ÝÃD
'   'If pub_OS = 1 Or m_bEditDoc Then
'      g_WordAp.Visible = True
'      g_WordAp.WindowState = wdWindowStateMaximize
'      g_WordAp.ActiveWindow.ActivePane.View.Zoom.Percentage = 100 'Added by Morgan 2012/11/30
'   'Else
'   '   g_WordAp.Visible = False
'   'End If

   With g_WordAp.Application
      'ª©­±³]©w
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.5)
      'Modified by Morgan 2014/1/23 ­¶­º¥[°ª(²Ä2­¶¦L¯È¥»·|À£¨ì«HÀY)
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(3.53)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.3)
      'Modify by Morgan 2011/5/13 ·s«H¯È¤W¤U³£¦³¹Ï
      '.Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(3)
      .Selection.PageSetup.FooterDistance = .CentimetersToPoints(3)
      
      .Selection.PageSetup.CharsLine = 40
      .Selection.PageSetup.LinesPage = 38
      
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = strFontSize
      '.Selection.ParagraphFormat.DisableLineHeightGrid = True
      
      '«O¯d«HÀYªÅ¶¡(2¦æ)
      '.Selection.TypeParagraph 'Removed by Morgan 2014/1/23 ­¶­º¤w¥[°ª
      .Selection.TypeParagraph
      
      '·s¼Wªí®æ(1*4)
      .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=4
      
      With .Selection.Tables(1)
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
      End With
    
      '³]©wªí®æ°ª«×Äæ¼e
      .Selection.SelectRow
      'Modified by Morgan 2014/2/20
      '.Selection.Cells.VerticalAlignment = wdAlignVerticalBottom
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.9), RulerStyle:=wdAdjustProportional
      .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(8.5), RulerStyle:=wdAdjustProportional
      .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
      .Selection.Cells(3).VerticalAlignment = wdCellAlignVerticalTop 'Added by Morgan 2012/3/20
      '.Selection.Cells(4).Width = .CentimetersToPoints(6)
      
      .Selection.InsertRows UBound(m_Head, 2)
      .Selection.Collapse Direction:=wdCollapseStart
      'ªíÀY¦C1
      For iRow = 1 To UBound(m_Head, 2)
         .Selection.Cells(1).SetHeight RowHeight:=12, HeightRule:=wdRowHeightAtLeast
         If m_Head(1, iRow) <> "" Then
            .Selection.TypeText Text:=m_Head(1, iRow)
         End If
         If m_Head(2, iRow) <> "" Then
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:=m_Head(2, iRow)
         Else
            .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
            .Selection.Cells.Merge
         End If
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         If m_Head(3, iRow) <> "" Then
            .Selection.TypeText Text:=m_Head(3, iRow)
         End If
         If m_Head(4, iRow) <> "" Then
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:=m_Head(4, iRow)
         Else
            .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
            .Selection.Cells.Merge
         End If
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
      Next

      .Selection.SelectRow
      .Selection.Cells.Merge
      .Selection.InsertRows 3
      .Selection.Collapse Direction:=wdCollapseStart
      With .Selection.Cells.Borders(wdBorderBottom)
         .LineStyle = wdLineStyleSingle
         .LineWidth = wdLineWidth050pt
         .ColorIndex = wdAuto
      End With
      
      .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
      .Selection.TypeText Text:="                           "
      .Selection.Font.Size = 14
      .Selection.Font.Bold = True
      .Selection.TypeText Text:=m_Title
      .Selection.Font.Size = strFontSize
      .Selection.TypeText Text:="      ½s¸¹: " & m_strDN
      .Selection.Font.Bold = False
      .Selection.MoveRight Unit:=wdCharacter, Count:=2
      '¼ÐÃD
      For iRow = 1 To UBound(m_Subject, 2)
         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
         If m_Subject(1, iRow) <> "" Then
            .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
            If m_Subject(2, iRow) = "" Then
               .Selection.TypeText Text:=m_Subject(1, iRow)
            Else
               .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.TypeText Text:=m_Subject(1, iRow)
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.TypeText Text:=m_Subject(2, iRow)
            End If
            
         Else
            .Selection.Cells(1).SetHeight RowHeight:=16, HeightRule:=wdRowHeightAtLeast
            If m_Subject(2, iRow) <> "" Then
               .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.TypeText Text:=m_Subject(2, iRow)
               
            ElseIf m_Subject(3, iRow) <> "" Then
               .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.9), RulerStyle:=wdAdjustProportional
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.MoveRight Unit:=wdCell, Count:=2
               .Selection.TypeText Text:=m_Subject(3, iRow)
            End If
         End If
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
         .Selection.InsertRows 1
         .Selection.Collapse Direction:=wdCollapseStart
      Next
      '©ú²Ó
      For iRow = 1 To UBound(m_Item)
         'Modify By Sindy 2021/10/29
         If Right(m_Item(iRow).iNo, 2) = "99" Then
            dblOffFeeSub = dblOffFeeSub + Val(Format(m_Item(iRow).iAmt))
         Else
            dblAttFeeSub = dblAttFeeSub + Val(Format(m_Item(iRow).iAmt))
         End If
         '2021/10/29 END
         
         .Selection.ParagraphFormat.SpaceBefore = 6
         .Selection.ParagraphFormat.SpaceAfter = 6
         .Selection.Cells.VerticalAlignment = wdAlignVerticalBottom
         If m_Item(iRow).ICode <> "" Then
            .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=False
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(10.5), RulerStyle:=wdAdjustProportional
            'Modified by Lydia 2016/03/03 + m_bSpecial5
            If m_bSpecial1 Or m_bSpecial5 Then
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(12.1), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.4), RulerStyle:=wdAdjustProportional
            Else
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(10.5), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(4), RulerStyle:=wdAdjustProportional
            End If
            .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(1.1), RulerStyle:=wdAdjustProportional
         Else
            .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=False
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.5), RulerStyle:=wdAdjustProportional
            .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.1), RulerStyle:=wdAdjustProportional
         End If
         
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.ParagraphFormat.RightIndent = .CentimetersToPoints(0.6)
         .Selection.TypeText Text:=m_Item(iRow).IDesc
         
         If m_Item(iRow).ICode <> "" Then
            .Selection.MoveRight Unit:=wdCell, Count:=1
            '.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
            '.Selection.ParagraphFormat.RightIndent = .CentimetersToPoints(0.6)
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Selection.Font.Bold = True
            .Selection.TypeText Text:=m_Item(iRow).ICode
            .Selection.Font.Bold = False
         End If
         
         .Selection.MoveRight Unit:=wdCell, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.TypeText Text:=m_Item(iRow).iCur
         
         .Selection.MoveRight Unit:=wdCell, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.TypeText Text:=m_Item(iRow).iAmt
         
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
         .Selection.InsertRows 1
         .Selection.Cells.Merge
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.ParagraphFormat.RightIndent = .CentimetersToPoints(0)
      Next
      UpdateA1K3940 dblAttFeeSub, dblOffFeeSub 'Add by Sindy 2021/10/29 §ó·s¥~¹ôªA°È¶O³W¶Oª÷ÃB
      
      '¦X­p
      .Selection.Cells(1).SetHeight RowHeight:=18, HeightRule:=wdRowHeightAtLeast
      .Selection.SelectRow
      With .Selection.Cells.Borders(wdBorderTop)
         .LineStyle = wdLineStyleSingle
         .LineWidth = wdLineWidth050pt
         .ColorIndex = wdAuto
      End With
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.5), RulerStyle:=wdAdjustProportional
      .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.1), RulerStyle:=wdAdjustProportional
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.Font.Bold = True
      .Selection.Font.Size = 13
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.TypeText Text:="                                " & m_Sum(1, 1)
      .Selection.Font.Size = strFontSize
      .Selection.Font.Bold = False
      
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.TypeText Text:=m_Sum(2, 1)
      
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.TypeText Text:=m_Sum(3, 1)
      .Selection.MoveRight Unit:=wdCharacter, Count:=2
      If UBound(m_Sum, 2) > 1 Then
         .Selection.Cells(1).SetHeight RowHeight:=18, HeightRule:=wdRowHeightAtLeast
         .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.5), RulerStyle:=wdAdjustProportional
         .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.1), RulerStyle:=wdAdjustProportional
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.ParagraphFormat.RightIndent = .CentimetersToPoints(0.6)
         .Selection.TypeText Text:=m_Sum(1, 2)
         
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.TypeText Text:=m_Sum(2, 2)
         
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.TypeText Text:=m_Sum(3, 2)
         
         UpdateA1K38 m_strDN, Format(m_Sum(3, 2)) 'Added by Morgan 2021/1/14 §ó·s½Ð´Ú³æ¬üª÷Á`ÃB
         
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.InsertRows 1
         .Selection.ParagraphFormat.RightIndent = .CentimetersToPoints(0)
         .Selection.Collapse Direction:=wdCollapseStart
      End If
      
      .Selection.SelectRow
      .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
      .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.5), RulerStyle:=wdAdjustProportional
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
      .Selection.Paragraphs.Alignment = wdAlignParagraphDistribute
      .Selection.TypeText Text:="vvvvvvvvvvvvv"
      
      'ªí§À1
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.InsertRows 1
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.Collapse Direction:=wdCollapseStart
      With .ActiveDocument.Bookmarks
         .add Range:=g_WordAp.Application.Selection.Range, Name:="BreakPos"
         .DefaultSorting = wdSortByLocation
         .ShowHidden = False
      End With
      .Selection.TypeText Text:=m_Footer(1, 1)
      
      If m_Footer(2, 1) <> "" Then
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.Cells.Split NumRows:=3, NumColumns:=2, MergeBeforeSplit:=False
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.Cells(1).SetHeight RowHeight:=36, HeightRule:=wdRowHeightAtLeast
         .Selection.MoveDown Unit:=wdLine, Count:=1
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
         With .Selection.Cells
             With .Borders(wdBorderLeft)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             With .Borders(wdBorderRight)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             With .Borders(wdBorderTop)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             With .Borders(wdBorderBottom)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
             .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
             .Borders.Shadow = False
         End With
         .Selection.ParagraphFormat.LeftIndent = .CentimetersToPoints(0.2)
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
         .Selection.Cells(1).SetHeight RowHeight:=52, HeightRule:=wdRowHeightAtLeast
         .Selection.TypeText Text:=m_Footer(2, 1)
         .Selection.HomeKey
         .Selection.MoveDown Unit:=wdLine, Count:=1
         .Selection.Cells(1).SetHeight RowHeight:=0, HeightRule:=wdRowHeightAtLeast
         .Selection.MoveDown Unit:=wdLine, Count:=1
      Else
         .Selection.MoveRight Unit:=wdCharacter, Count:=3
      End If
      
      '³Æµù
      If UBound(m_Footer, 2) > 1 Then
         .Selection.SelectRow
         .Selection.Font.Bold = True
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
         .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.TypeText Text:="¡°"
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         iPos1 = InStr(m_Footer(1, 2), "±©¤_¶×´Ú¦Z")
         If iPos1 > 0 Then
            iPos2 = InStr(iPos1, m_Footer(1, 2), "¡C")
         End If
         If iPos1 > 0 And iPos2 > 0 Then
            .Selection.TypeText Text:=Left(m_Footer(1, 2), iPos1 - 1)
            .Selection.Font.Underline = wdUnderlineDouble
            .Selection.TypeText Text:=Mid(m_Footer(1, 2), iPos1, iPos2 - iPos1)
            .Selection.Font.Underline = wdUnderlineNone
            .Selection.TypeText Text:=Mid(m_Footer(1, 2), iPos2)
         Else
            .Selection.TypeText Text:=m_Footer(1, 2)
         End If
         
         For iRow = 3 To UBound(m_Footer, 2)
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.InsertRows 1
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:=m_Footer(1, iRow)
         Next
         .Selection.Font.Bold = False
      End If
      
      .Selection.WholeStory
      .Selection.Font.Name = "Times New Roman"
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .ActiveDocument.Repaginate
      '¶W¹L1­¶®É´¡¤J­¶½X
      If .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) > 1 Then
         .Selection.GoTo what:=wdGoToBookmark, Name:="BreakPos"
         .Selection.InsertBreak Type:=wdPageBreak
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .ActiveDocument.Repaginate
         If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
            .ActiveWindow.ActivePane.View.Type = wdPageView
         Else
            .ActiveWindow.View.Type = wdPageView
         End If
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
         .Selection.TypeParagraph
         .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldPage
         .Selection.TypeText Text:="/"
         .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldEmpty, Text:="NUMPAGES ", PreserveFormatting:=True
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
         
      End If
      .ActiveDocument.Bookmarks("BreakPos").Delete
      '.ActiveWindow.View.TableGridlines = False
      .Selection.HomeKey Unit:=wdStory
      
      If m_bPrintWord And m_iSpCopies > 0 Then
         .ActiveDocument.PrintOut Background:=False, Copies:=m_iSpCopies, Collate:=True
         m_iPageCount = .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) 'Added by Morgan 2015/2/26
      End If
      'Modify by Morgan 2011/7/15 ¤º±M§ï­n±a«HÀY--³¢
      'If m_bSaveWord Then
      'Modified by Morgan 2014/2/26 +¥~°Ó³£±a«HÀY
      'Modified by Morgan 2014/8/25 §ï¤Á´«¦Lªí¾÷¤è¦¡¦C¦LPDF,¤£¥Î­«¶]Word
      'If m_bWord2Pdf Or m_bSaveWord Or (m_bEditDoc And (Pub_StrUserSt03 = "P12" Or Left(Pub_StrUserSt03, 2) = "F1")) Then
      'Modify By Sindy 2015/7/13 ¶®®S¸ò¨q¬Â»¡­n¥Î±M§Qªk«ß«HÀY
      'If m_b2PDF Or m_bWord2Pdf Or m_bSaveWord Or (m_bEditDoc And (Pub_StrUserSt03 = "P12" Or Left(Pub_StrUserSt03, 2) = "F1")) Then
      If m_b2PDF Or m_bWord2Pdf Or m_bSaveWord Or (m_bEditDoc And ((m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") Or Left(Pub_StrUserSt03, 2) = "F1")) Then
         
         'Added by Morgan 2020/3/31
         If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é Then
            PUB_GetLetterPicID "2", m_CP01, iPicNo, iPicNo2, 1, True, Pub_StrUserSt03
            If PUB_ReadDB2File(stFileName, iPicNo) Then
               For ii = 1 To .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) '¨C­¶³£­n¦³«HÀY§À
                  Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                  oShape.ZOrder 4
                  oShape.LockAnchor = True
                  oShape.LockAspectRatio = -1
                  oShape.Width = .CentimetersToPoints(21)
                  oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                  oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                  oShape.Left = .CentimetersToPoints(0)
                  oShape.Top = .CentimetersToPoints(0)
                  oShape.WrapFormat.Type = wdWrapNone
                  .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
               Next ii
               .Selection.HomeKey Unit:=wdStory
               If iPicNo2 > 0 Then
                  If PUB_ReadDB2File(stFileName, iPicNo2) Then
                     For ii = 1 To .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)
                        Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                        oShape.ZOrder 4
                        oShape.LockAnchor = True
                        oShape.LockAspectRatio = -1
                        oShape.Width = .CentimetersToPoints(21)
                        oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                        oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                        oShape.Left = .CentimetersToPoints(0)
                        oShape.Top = .CentimetersToPoints(27.6)
                        oShape.WrapFormat.Type = wdWrapNone
                        .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
                     Next ii
                  End If
                  .Selection.HomeKey Unit:=wdStory
               End If
            End If
         Else
         'end 2020/3/31
         
            'Modify By Sindy 2015/7/17
            'If Pub_StrUserSt03 = "P12" Then
            If (m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") Then
            '2015/7/17 END
               '´¡¤J¹Ï¤ùÀÉ®×(¥Î¸õ­¶²Å¸¹§PÂ_­¶¼Æ)
               If PUB_ReadDB2File(stFileName, 19) Then
                  .Selection.HomeKey Unit:=wdStory, Extend:=wdMove
                  Do
                     Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                     oShape.ZOrder 5
                     oShape.LockAnchor = True
                     oShape.LockAspectRatio = -1
                     oShape.Width = 546.5
                     oShape.WrapFormat.Type = wdWrapNone
                     oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                     oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                     oShape.Left = .CentimetersToPoints(1)
                     oShape.Top = .CentimetersToPoints(1)
                     iHeadCount = iHeadCount + 1
                     .ActiveDocument.Repaginate
                     '¨S¦³¸õ­¶²Å¸¹¦ý­¶¼Æ¤j©ó¤w¦L«HÀY¼Æ
                     If iHeadCount < .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) Then
                        .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
                     Else
                        Exit Do
                     End If
                  Loop
               End If
            Else
            
               '´¡¤J¹Ï¤ùÀÉ®×(¥Î¸õ­¶²Å¸¹§PÂ_­¶¼Æ)
               If PUB_ReadDB2File(stFileName, 7) Then
                  .Selection.HomeKey Unit:=wdStory, Extend:=wdMove
                  Do
                     Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                     oShape.ZOrder 5
                     oShape.LockAnchor = True
                     oShape.LockAspectRatio = -1
                     oShape.Width = 546.5
                     oShape.WrapFormat.Type = wdWrapNone
                     oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                     oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                     oShape.Left = .CentimetersToPoints(1)
                     oShape.Top = .CentimetersToPoints(1)
                     iHeadCount = iHeadCount + 1
                     .ActiveDocument.Repaginate
                     '¨S¦³¸õ­¶²Å¸¹¦ý­¶¼Æ¤j©ó¤w¦L«HÀY¼Æ
                     If iHeadCount < .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) Then
                        .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
                     Else
                        Exit Do
                     End If
                  Loop
               End If
            End If
            
         End If 'Added by Morgan 2020/3/31
         
         If m_bWord2Pdf Then
            .ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
            
         'Added by Morgan 2014/8/26
         '§ï¤Á´«¦Lªí¾÷¤è¦¡¦C¦LPDF,¤£¥Î­«¶]Word
         ElseIf m_b2PDF Then
            PrintWord2PDF
         'end 2014/8/26
         
         End If
         
         If m_bSaveWord Then
            RidFile m_EFilePath
            .ActiveDocument.SaveAs m_EFilePath
         End If
      End If
   End With
   'Modified by Lydia 2019/04/09 §ï¦¨¦@¥Î¼Ò²Õ
   'RePosWord bVisible, m_WordLeft, m_WordTop 'Added by Morgan 2014/6/26
   Pub_RePosWord g_WordAp, bVisible, m_WordLeft, m_WordTop
   
   If m_bEditDoc Then
      g_WordAp.Visible = True
      g_WordAp.WindowState = wdWindowStateMaximize
      g_WordAp.Activate
   Else
      g_WordAp.ActiveDocument.Close wdDoNotSaveChanges
      If g_WordAp.Visible = False Then
         g_WordAp.Quit wdDoNotSaveChanges
         Set g_WordAp = Nothing 'Added by Lydia 2017/12/12 Á×§K§Ö³t¶}±ÒWord,µ{¦¡¥X¿ù
      Else
         g_WordAp.Visible = True
      End If
   End If
   
   Exit Sub
   
ErrHnd:
   If Err.Number <> 0 Then
'Modified by Morgan 2014/6/26
'      If iResumeCnt > 3 Then
'         MsgBox "¿ù»~ : " & Err.Description, vbCritical
'      Else
'         iResumeCnt = iResumeCnt + 1
'         Select Case Err.NUMBER
'            Case 91:
'               g_WordAp.Documents.Add
'               Resume Next
'            Case 462:
'               Set g_WordAp = New Word.Application
'               Resume
'            Case Else:
'               MsgBox "¿ù»~ : " & Err.Description, vbCritical
'         End Select
'      End If
      MsgBox "¿ù»~ : " & Err.Description, vbCritical
'end 2014/6/26
Resume
   End If
End Sub

'Modify by Morgan 2011/1/5 ¥[¥i¦C¦L¤Î¦sÀÉ
Private Sub runWordEnglish()
   
   'Added by Morgan 2013/10/25
   If bolNewForm Then
      runWordEnglishNew
      Exit Sub
   End If
   'end 2013/10/25
   
   
   Dim iRow As Integer, iCol As Integer, lngCellWidth As Long
   Dim strLastCountry As String, iFCol As Integer, iCols As Integer, bolPrintCountry As Boolean
   Dim strFontSize As String
   Dim iResumeCnt As Integer
   Dim bVisible As Boolean, stFileName As String, iHeadCount As Integer
   Dim strText As String
   Dim oShape
   Dim strLstStr As String
   Dim ii As Integer 'Added by Morgan 2014/3/4
   Dim pCount As Integer 'Added by Lydia 2016/03/03
   Dim dblOffFeeSub As Double
   Dim dblAttFeeSub As Double
   
   strFontSize = 12
   
   'Modified by Lydia 2019/04/09 §ï¦¨¦@¥Î¼Ò²Õ
   'If NewDoc(bVisible, m_WordLeft, m_WordTop) = False Then Exit Sub 'Added by Morgan 2014/6/26
   If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = False Then Exit Sub
   
On Error GoTo ErrHnd

'Removed by Morgan 2014/6/26
'   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
'
'   g_WordAp.Documents.Add
'
'   bVisible = g_WordAp.Visible
'
'   '¤£Åã¥Ü¥i¯à·|¦³°ÝÃD
'   'If pub_OS = 1 Or m_bEditDoc Then
'      g_WordAp.Visible = True
'      g_WordAp.WindowState = wdWindowStateMaximize
'      g_WordAp.ActiveWindow.ActivePane.View.Zoom.Percentage = 100 'Added by Morgan 2012/11/30
'   'Else
'   '   g_WordAp.Visible = False
'   'End If
   
   With g_WordAp.Application
      'ª©­±³]©w
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.5)
      'Modified by Morgan 2014/1/23 ­¶­º¥[°ª(²Ä2­¶¦L¯È¥»·|À£¨ì«HÀY)
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(3.53)
      'Modify By Sindy 2015/11/9 ¦]¥~±M­n¥Î¶}µ¡«H«Ê
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(4.3)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4)
      '2015/11/9 END
      'Modify by Morgan 2011/5/13 ·s«H¯È¤W¤U³£¦³¹Ï
      '.Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(3)
      .Selection.PageSetup.FooterDistance = .CentimetersToPoints(3)
      
      .Selection.PageSetup.CharsLine = 40
      .Selection.PageSetup.LinesPage = 38
      
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = strFontSize
      '.Selection.ParagraphFormat.DisableLineHeightGrid = True
      
      '«O¯d«HÀYªÅ¶¡(2¦æ)
      '.Selection.TypeParagraph 'Removed by Morgan 2014/1/23 ­¶­º¤w¥[°ª
      .Selection.TypeParagraph '¤£¥iMark,¦]Word2007¦b´¡«HÀY®É¦ì¸m·|¿ù,·|´¡¤J¨ìªí®æ¸Ì
       
      'Added by Morgan 2012/1/4 97 »P 2007 ¹w³]¤£¦P»Ý«ü©w¦æ¶Z
      With .Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
        .DisableLineHeightGrid = True
      End With
      
      '·s¼Wªí®æ(1*4)
      .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=4
      
      With .Selection.Tables(1)
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
      End With
      
      '³]©wªí®æ°ª«×Äæ¼e
      .Selection.SelectRow
      'Modified by Morgan 2014/2/20
      '.Selection.Cells.VerticalAlignment = wdAlignVerticalBottom
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.9), RulerStyle:=wdAdjustProportional
      .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(8.5), RulerStyle:=wdAdjustProportional
      .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
      .Selection.Cells(3).VerticalAlignment = wdCellAlignVerticalTop 'Added by Morgan 2012/3/20
      '.Selection.Cells(4).Width = .CentimetersToPoints(6)
      
      .Selection.InsertRows UBound(m_Head, 2)
      .Selection.Collapse Direction:=wdCollapseStart
      intI = 0 'Added by Morgan 2014/2/20
      'ªíÀY¦C1
      For iRow = 1 To UBound(m_Head, 2)
         .Selection.Cells(1).SetHeight RowHeight:=12, HeightRule:=wdRowHeightAtLeast
         If m_Head(1, iRow) <> "" Then
            .Selection.TypeText Text:=m_Head(1, iRow)
         End If
         If m_Head(2, iRow) <> "" Then
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:=m_Head(2, iRow)
            
            'Added by Morgan 2014/2/20 Äæ¦ì¦X¨Ö§_«h­Y©¼©Ò®×¸¹©Î«È¤á®×¥ó®×¸¹¤Óªø¦³§é¦æ®É¦a§}·|¹jªÅ¥Õ¦æ Ex.X10302374
            If iRow > 1 Then
               If m_Head(2, iRow - 1) <> "" Then
                  .Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
                  .Selection.Cells.Merge
                  .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
                  intI = intI + 1
               End If
            End If
            'end 2014/2/20
            
         Else
            intI = 0 'Added by Morgan 2016/2/17
            .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
            .Selection.Cells.Merge
         End If
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         
         'Added by Morgan 2014/2/20
         If intI > 0 Then
            .Selection.MoveDown Unit:=wdLine, Count:=intI
         End If
         'end 2014/2/20
         
         If m_Head(3, iRow) <> "" Then
            .Selection.TypeText Text:=m_Head(3, iRow)
         End If
         If m_Head(4, iRow) <> "" Then
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:=m_Head(4, iRow)
         Else
            .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
            .Selection.Cells.Merge
         End If
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
      Next
      
      .Selection.SelectRow
      .Selection.Cells.Merge
      .Selection.InsertRows 3
      .Selection.Collapse Direction:=wdCollapseStart
      
      'Modified by Morgan 2017/3/22 Dow ®æ¦¡ªíÀY¸ê®Æ¸û¦h·|¤ÓÀ½§ï±N½Ð´Ú³æ¸¹³o¦C¥[°ª¨Ã««ª½¾a¤¤¹ï»ô
      '.Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
      .Selection.Cells(1).SetHeight RowHeight:=30, HeightRule:=wdRowHeightAtLeast
      .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
      'end 2017/3/22

      .Selection.TypeText Text:="                           "
      .Selection.Font.Size = 14
      .Selection.Font.Bold = True
      .Selection.TypeText Text:=m_Title
      .Selection.Font.Size = strFontSize
      .Selection.TypeText Text:="      NO. " & m_strDN
      .Selection.Font.Bold = False
      
      'Added by Morgan 2012/6/11
'Removed by Morgan 2012/7/24 ¥ý¨ú®ø--¶À¬ü¬Ã
'      If m_bSpecial1 Then
'         .Selection.MoveRight Unit:=wdCharacter, Count:=1
'         .Selection.InsertRows 1
'         .Selection.Font.Bold = True
'         .Selection.Cells.Split NumRows:=1, NumColumns:=5, MergeBeforeSplit:=False
'         .Selection.Cells(1).SetHeight RowHeight:=20, HeightRule:=wdRowHeightAtLeast
'         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.9), RulerStyle:=wdAdjustProportional
'         .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(4.1), RulerStyle:=wdAdjustProportional
'         .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.6), RulerStyle:=wdAdjustProportional
'         .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(3.5), RulerStyle:=wdAdjustProportional
'         .Selection.Collapse Direction:=wdCollapseStart
'         .Selection.MoveRight Unit:=wdCell, Count:=1
'         .Selection.TypeText Text:="TimeKeeper"
'         .Selection.MoveRight Unit:=wdCell, Count:=1
'         .Selection.TypeText Text:="ID"
'         .Selection.MoveRight Unit:=wdCell, Count:=1
'         .Selection.TypeText Text:="Hours"
'         .Selection.MoveRight Unit:=wdCell, Count:=1
'         .Selection.TypeText Text:="Rate/Hour"
'         .Selection.MoveRight Unit:=wdCharacter, Count:=1
'         .Selection.InsertRows 1
'         .Selection.Font.Bold = False
'         .Selection.Collapse Direction:=wdCollapseStart
'         .Selection.MoveRight Unit:=wdCell, Count:=1
'         .Selection.TypeText Text:="Jacky Wang"
'         .Selection.MoveRight Unit:=wdCell, Count:=1
'         .Selection.TypeText Text:="JW"
'         .Selection.MoveRight Unit:=wdCell, Count:=1
'         .Selection.TypeText Text:="2.5"
'         .Selection.MoveRight Unit:=wdCell, Count:=1
'         .Selection.TypeText Text:="US$200"
'         .Selection.MoveRight Unit:=wdCharacter, Count:=2
'         .Selection.InsertRows 1
'         .Selection.Collapse Direction:=wdCollapseStart
'      End If
'end 2012/7/24
      'end 2012/6/11
      
      With .Selection.Cells.Borders(wdBorderBottom)
         .LineStyle = wdLineStyleSingle
         .LineWidth = wdLineWidth050pt
         .ColorIndex = wdAuto
      End With
      .Selection.MoveRight Unit:=wdCharacter, Count:=2
      
      '¼ÐÃD
      For iRow = 1 To UBound(m_Subject, 2)
         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
         If m_Subject(1, iRow) <> "" Then
            .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
            If m_Subject(2, iRow) = "" Then
               .Selection.TypeText Text:=m_Subject(1, iRow)
            Else
               .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.TypeText Text:=m_Subject(1, iRow)
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.TypeText Text:=m_Subject(2, iRow)
            End If
            
         Else
            .Selection.Cells(1).SetHeight RowHeight:=16, HeightRule:=wdRowHeightAtLeast
            If m_Subject(2, iRow) <> "" Then
               .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.TypeText Text:=m_Subject(2, iRow)
               strLstStr = m_Subject(2, iRow)
            ElseIf m_Subject(3, iRow) <> "" Then
               .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
               'Modified by Morgan 2012/6/29 ­n¦Ò¼{«e¸m¤å¦r¦³¨â²Õ("In the name of ","Applicant: ")
               If InStr(strLstStr, "In the name of") = 1 Then
                  .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.8), RulerStyle:=wdAdjustProportional
               Else
                  .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
               End If
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.MoveRight Unit:=wdCell, Count:=2
               .Selection.TypeText Text:=m_Subject(3, iRow)
            End If
         End If

         .Selection.MoveRight Unit:=wdCharacter, Count:=2
         .Selection.InsertRows 1
         .Selection.Collapse Direction:=wdCollapseStart
      Next
      '©ú²Ó
      For iRow = 1 To UBound(m_Item)
         'Modify By Sindy 2021/10/29
         If Right(m_Item(iRow).iNo, 2) = "99" Then
            dblOffFeeSub = dblOffFeeSub + Val(Format(m_Item(iRow).iAmt))
         Else
            dblAttFeeSub = dblAttFeeSub + Val(Format(m_Item(iRow).iAmt))
         End If
         '2021/10/29 END
         
         .Selection.ParagraphFormat.SpaceBefore = 6
         .Selection.ParagraphFormat.SpaceAfter = 6
         .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
         If m_Item(iRow).ICode <> "" Then
            'Modified by Lydia 2016/03/03 + m_bSpecial5
            'Modified by Morgan 2017/8/17 Dow ¥[Äæ¦ì Billable Time
            'If m_bSpecial1 Or m_bSpecial5 Then
            If m_bSpecial1 Then
               .Selection.Cells.Split NumRows:=1, NumColumns:=5, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(6.3), RulerStyle:=wdAdjustProportional
               .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.7), RulerStyle:=wdAdjustProportional
               .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(3), RulerStyle:=wdAdjustProportional
            ElseIf m_bSpecial5 Then
            'end 2017/8/17
               .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(9), RulerStyle:=wdAdjustProportional
               .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(3), RulerStyle:=wdAdjustProportional
               
            ElseIf m_bSpecial2 Or m_bSpecial3 Then
               .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(12.5), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
               .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(1.1), RulerStyle:=wdAdjustProportional
               
            'Added by Morgan 2014/2/24
            ElseIf m_bSpecial4 Then
               .Selection.Cells.Split NumRows:=1, NumColumns:=5, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(8), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(3.7), RulerStyle:=wdAdjustProportional
               .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
               .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(1.1), RulerStyle:=wdAdjustProportional
               
            'end 2014/2/24
            Else
               .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(10.5), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(4), RulerStyle:=wdAdjustProportional
               .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(1.1), RulerStyle:=wdAdjustProportional
            End If
         Else
            .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=False
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.5), RulerStyle:=wdAdjustProportional
            .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.1), RulerStyle:=wdAdjustProportional
         End If
         
         .Selection.Collapse Direction:=wdCollapseStart
         
         .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         'Modified by Lydia 2016/03/03 + m_bSpecial5
         If m_bSpecial1 Or m_bSpecial5 Then
            'Äæ¦ì¦WºÙ
            If iRow = 1 Then
               .Selection.Font.Bold = True
               .Selection.TypeText Text:="DATE"
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
               .Selection.Font.Bold = True
               .Selection.TypeText Text:="DESCRIPTION"
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
               .Selection.Font.Bold = True
               'Added by Lydia 2016/03/03 +m_bSpecial5
               If m_bSpecial5 Then
                  .Selection.TypeText Text:="Order number"
               Else
                  'Added by Morgan 2017/8/17 Dow ¥[Äæ¦ì Billable Time
                  .Selection.TypeText Text:="Billable Time"
                  .Selection.MoveRight Unit:=wdCell, Count:=1
                  .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
                  .Selection.Font.Bold = True
                  'end 2017/8/17
                  .Selection.TypeText Text:="UTBMS Code"
               End If
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
               .Selection.Font.Bold = True
               .Selection.TypeText Text:="Amount " & m_Item(iRow).iCur
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
               .Selection.InsertRows 1
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            End If
            .Selection.Font.Bold = False
            .Selection.TypeText Text:=m_Item(iRow).IDate
            .Selection.MoveRight Unit:=wdCell, Count:=1
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         
         'Added by Morgan 2014/2/24
         ElseIf m_bSpecial4 Then
            'Äæ¦ì¦WºÙ
            If iRow = 1 Then
               .Selection.Font.Bold = True
               .Selection.TypeText Text:="Description of the work"
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
               .Selection.Font.Bold = True
               .Selection.TypeText Text:="Date of the service"
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
               .Selection.Font.Bold = True
               .Selection.TypeText Text:="Hours spent"
               .Selection.MoveRight Unit:=wdCell, Count:=2
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
               .Selection.InsertRows 1
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            End If
         End If
         
         .Selection.ParagraphFormat.RightIndent = .CentimetersToPoints(0.6)
         .Selection.Font.Bold = False
         .Selection.TypeText Text:=m_Item(iRow).IDesc
         
         'Added by Morgan 2014/2/24
         If m_bSpecial4 Then
            .Selection.MoveRight Unit:=wdCell, Count:=1
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Selection.Font.Bold = False
            .Selection.Cells(1).VerticalAlignment = wdAlignVerticalBottom
            .Selection.TypeText Text:=m_Item(iRow).IDate
         End If
         'end 2014/2/24
         
         'Added by Morgan 2017/8/17 Dow ¥[Äæ¦ì Billable Time
         If m_bSpecial1 Then
            .Selection.MoveRight Unit:=wdCell, Count:=1
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Selection.Font.Bold = False
            .Selection.Cells(1).VerticalAlignment = wdAlignVerticalTop
            If m_bDowX Then
               If ChkDowXFormat(m_Item(iRow).iNo, True) Then
                  'Modified by Morgan 2020/8/6
                  If m_bDowN = True Then 'Added by Morgan 2020/10/12
                     '.Selection.TypeText Text:=Round(m_Item(iRow).iAmt / 120, 1)
                     .Selection.TypeText Text:=Round(m_Item(iRow).iAmtNoDisc / 120, 1)
                  
                  'Added by Morgan 2020/10/12
                  Else
                     .Selection.TypeText Text:=Round(m_Item(iRow).iAmt / 120, 1)
                  End If
                  'end 2020/10/12
                  'end 2020/8/6
               End If
            End If
         End If
         'end 2017/8/17
         
         If m_Item(iRow).ICode <> "" Then
            .Selection.MoveRight Unit:=wdCell, Count:=1
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Selection.Font.Bold = False
            'Added by Morgan 2014/2/24
            If m_bSpecial4 Then
               .Selection.Cells(1).VerticalAlignment = wdAlignVerticalBottom
            End If
            'end 2014/2/24
            .Selection.TypeText Text:=m_Item(iRow).ICode
         End If
         'Modified by Lydia 2016/03/03 + m_bSpecial5
         'If Not m_bSpecial1 Then
         If Not (m_bSpecial1 Or m_bSpecial5) Then
            .Selection.MoveRight Unit:=wdCell, Count:=1
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Selection.Font.Bold = False
            .Selection.Cells(1).VerticalAlignment = wdAlignVerticalBottom
            .Selection.TypeText Text:=m_Item(iRow).iCur
         End If
         
         .Selection.MoveRight Unit:=wdCell, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.Font.Bold = False
         'Modified by Lydia 2016/03/03 + m_bSpecial5
         If m_bSpecial1 Or m_bSpecial5 Then
            strText = Format(m_Item(iRow).iAmt, FDollar)
            
            If m_bSpecial5 Then 'Added by Morgan 2017/3/22 Dow ®æ¦¡(m_bSpecial1)¨ú®ø$²Å¸¹--ªô¤l·ì
               If iRow = 1 Then
                  'Modified by Morgan 2016/9/8 «D¬üª÷¤£­n¦L$
                  If m_Item(iRow).iCur = "USD" Then
                     strText = "$" & Right(String(12, " ") & strText, 9)
                  Else
                     strText = " " & Right(String(12, " ") & strText, 9)
                  End If
               End If
            End If
            
            If Val(m_Item(UBound(m_Item)).iAmt) < 0 And iRow = UBound(m_Item) - 1 Then
               strText = Right(String(10, " ") & strText, 10)
               .Selection.Font.Underline = wdUnderlineSingle
               .Selection.TypeText Text:=strText
               .Selection.Font.Underline = wdUnderlineNone
            Else
               .Selection.TypeText Text:=strText
            End If
         Else
            .Selection.Cells(1).VerticalAlignment = wdAlignVerticalBottom
            .Selection.TypeText Text:=m_Item(iRow).iAmt
         End If
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
         .Selection.InsertRows 1
         .Selection.Cells.Merge
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.ParagraphFormat.RightIndent = .CentimetersToPoints(0)
      Next
      UpdateA1K3940 dblAttFeeSub, dblOffFeeSub 'Add by Sindy 2021/10/29 §ó·s¥~¹ôªA°È¶O³W¶Oª÷ÃB
            
      '¦X­p
      .Selection.Cells(1).SetHeight RowHeight:=18, HeightRule:=wdRowHeightAtLeast
      .Selection.SelectRow
      'Modified by Lydia 2016/03/03 + m_bSpecial5
      If m_bSpecial1 Or m_bSpecial5 Then
         
         .Selection.Cells.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
         .Selection.Cells.Borders(wdBorderBottom).LineWidth = wdLineWidth050pt
         .Selection.Cells.Borders(wdBorderBottom).ColorIndex = wdAuto
         
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(11.5), RulerStyle:=wdAdjustProportional
         .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(3), RulerStyle:=wdAdjustProportional
         .Selection.ParagraphFormat.SpaceBefore = 6
         .Selection.ParagraphFormat.SpaceAfter = 6
         .Selection.Cells.VerticalAlignment = wdAlignVerticalBottom
         .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
         
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.Font.Bold = True
         .Selection.TypeText Text:="Total " & m_Item(1).iCur & ":"
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.Font.Bold = False
         .Selection.Font.Underline = wdUnderlineSingle
         
         'Added by Morgan 2017/3/22 Dow ®æ¦¡(m_bSpecial1)¨ú®ø$²Å¸¹--ªô¤l·ì
         If m_bSpecial1 Then
            .Selection.TypeText Text:=m_Sum(3, 1)
         Else
         'end 2017/3/22
         
            'Modified by Morgan 2016/9/8 «D¬üª÷¤£­n¦L$
            If m_Item(1).iCur = "USD" Then
               .Selection.TypeText Text:="$" & Right(String(10, " ") & m_Sum(3, 1), 9)
            Else
               .Selection.TypeText Text:=" " & Right(String(10, " ") & m_Sum(3, 1), 9)
            End If
            'end 2016/9/8
            
         End If 'Added by Morgan 2017/3/22
         .Selection.Font.Underline = wdUnderlineNone
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
      Else
         With .Selection.Cells.Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .ColorIndex = wdAuto
         End With
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
         If m_bSpecial4 Then
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.2), RulerStyle:=wdAdjustProportional
         Else
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.5), RulerStyle:=wdAdjustProportional
         End If
         .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.1), RulerStyle:=wdAdjustProportional
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.Font.Bold = True
         .Selection.Font.Size = 13
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.TypeText Text:="                                " & m_Sum(1, 1)
         .Selection.Font.Size = strFontSize
         .Selection.Font.Bold = False
         
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.TypeText Text:=m_Sum(2, 1)
      
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.TypeText Text:=m_Sum(3, 1)
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
      End If
      
      If UBound(m_Sum, 2) > 1 Then
         .Selection.Cells(1).SetHeight RowHeight:=18, HeightRule:=wdRowHeightAtLeast
         .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
         If m_bSpecial4 Then
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.2), RulerStyle:=wdAdjustProportional
         Else
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.5), RulerStyle:=wdAdjustProportional
         End If
         .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.1), RulerStyle:=wdAdjustProportional
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.ParagraphFormat.RightIndent = .CentimetersToPoints(0.6)
         .Selection.TypeText Text:=m_Sum(1, 2)
         
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.TypeText Text:=m_Sum(2, 2)
         
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.TypeText Text:=m_Sum(3, 2)
         
         UpdateA1K38 m_strDN, Format(m_Sum(3, 2)) 'Added by Morgan 2021/1/14 §ó·s½Ð´Ú³æ¬üª÷Á`ÃB
         
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.InsertRows 1
         .Selection.ParagraphFormat.RightIndent = .CentimetersToPoints(0)
         .Selection.Collapse Direction:=wdCollapseStart
      End If
      
      .Selection.SelectRow
      .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
      .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
      If m_bSpecial4 Then
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.2), RulerStyle:=wdAdjustProportional
      Else
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.5), RulerStyle:=wdAdjustProportional
      End If
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      'Modified by Lydia 2016/03/03 + m_bSpecial5
      'If Not m_bSpecial1 Then
      If Not (m_bSpecial1 Or m_bSpecial5) Then
         .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
         .Selection.Paragraphs.Alignment = wdAlignParagraphDistribute
         .Selection.TypeText Text:="vvvvvvvvvvvvv"
         .Selection.Cells(1).SetHeight RowHeight:=12, HeightRule:=wdRowHeightAtLeast 'Added by Morgan 2014/2/25
      End If
      
      'ªí§À1
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      'Modified by Morgan 2014/2/24
      '.Selection.InsertRows 1
      If m_bSpecial4 Then
         .Selection.InsertRows 2
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.TypeText Text:="Timekeeper of the service: Isao Lee"
         .Selection.TypeParagraph
         .Selection.TypeText Text:="Rate of professional' fees: 6,000 NTD"
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.MoveDown Unit:=wdLine, Count:=1
         .Selection.SelectRow
      Else
         .Selection.InsertRows 1
      End If
      'end 2014/2/24
     
      .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleNone
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.Collapse Direction:=wdCollapseStart
      With .ActiveDocument.Bookmarks
         .add Range:=g_WordAp.Application.Selection.Range, Name:="BreakPos"
         .DefaultSorting = wdSortByLocation
         .ShowHidden = False
      End With
      .Selection.TypeText Text:=m_Footer(1, 1)
      
      If m_Footer(2, 1) <> "" Then
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.Cells.Split NumRows:=3, NumColumns:=1, MergeBeforeSplit:=False
         .Selection.Cells(1).SetHeight RowHeight:=36, HeightRule:=wdRowHeightAtLeast
         
         'Modified by Morgan 2013/5/28 ­×¥¿¿Ã¹õ¥iÅã¥Ü¨â­¶®É·|¦³°ÝÃD
         '.Selection.MoveDown Unit:=wdLine, Count:=1
         '.Selection.MoveLeft Unit:=wdCharacter, Count:=2
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.MoveDown Unit:=wdLine, Count:=2
         'end 2013/5/28
         
         
         .Selection.Cells(1).Borders(wdBorderTop).LineStyle = wdLineStyleNone
         .Selection.MoveUp Unit:=wdLine, Count:=1
         .Selection.Cells(1).Borders(wdBorderTop).LineStyle = wdLineStyleNone
         .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
         With .Selection.Cells(1)
             With .Borders(wdBorderLeft)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             With .Borders(wdBorderRight)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             With .Borders(wdBorderTop)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             With .Borders(wdBorderBottom)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             '.Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
             '.Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
             '.Borders.Shadow = False
         End With
         .Selection.ParagraphFormat.LeftIndent = .CentimetersToPoints(0.2)
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
         .Selection.Cells(1).SetHeight RowHeight:=52, HeightRule:=wdRowHeightAtLeast
         .Selection.TypeText Text:=m_Footer(2, 1)
         .Selection.HomeKey
         .Selection.MoveDown Unit:=wdLine, Count:=1
         .Selection.Cells(1).SetHeight RowHeight:=0, HeightRule:=wdRowHeightAtLeast
         '.Selection.MoveDown Unit:=wdLine, Count:=1
         .Selection.EndKey Unit:=wdStory
         'Modified by Morgan 2014/8/22 ­Yµe­±Åã¥Ü¥ª¥k¨â­¶¥B´å¼Ð¦b²Ä¤G­¶ªº²Ä¤@¦C®É¤£·|²¾¨ì²Ä¤@­¶ªº³Ì«á¤@¦C
         '.Selection.MoveUp Unit:=wdLine, Count:=1
         .Selection.MoveLeft Unit:=wdCharacter, Count:=1
         'end 2014/8/22
      Else
         .Selection.MoveRight Unit:=wdCharacter, Count:=3
      End If
      
      '³Æµù
      If UBound(m_Footer, 2) > 1 Then
         .Selection.SelectRow
         .Selection.Font.Bold = True
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
         .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.TypeText Text:="PS:"
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.TypeText Text:=m_Footer(1, 2)
         For iRow = 3 To UBound(m_Footer, 2)
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.InsertRows 1
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:=m_Footer(1, iRow)
         Next
         .Selection.Font.Bold = False
      End If
      
      .Selection.WholeStory
      .Selection.Font.Name = "Times New Roman"
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .ActiveDocument.Repaginate
      '¶W¹L1­¶®É´¡¤J­¶½X
      If .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) > 1 Then
         .Selection.GoTo what:=wdGoToBookmark, Name:="BreakPos"
         pCount = .Selection.Information(wdActiveEndPageNumber)
         If .Selection.Information(wdActiveEndPageNumber) = 1 Then
            .Selection.InsertBreak Type:=wdPageBreak
            .Selection.TypeParagraph
            .Selection.TypeParagraph
         End If
         .ActiveDocument.Repaginate
         If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
            .ActiveWindow.ActivePane.View.Type = wdPageView
         Else
            .ActiveWindow.View.Type = wdPageView
         End If
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
         .Selection.TypeParagraph
         .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldPage
         .Selection.TypeText Text:="/"
         .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldEmpty, Text:="NUMPAGES ", PreserveFormatting:=True
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
      End If
      .ActiveDocument.Bookmarks("BreakPos").Delete
      '.ActiveWindow.View.TableGridlines = False
      .Selection.HomeKey Unit:=wdStory
      
      'Removed by Morgan 2020/4/1 +¥~°Ó¤]­n±a«HÀY--³¯ª÷½¬
      'If Not (m_CP01 = "FCP" Or m_CP01 = "FG" Or m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") Then 'Add By Sindy 2015/7/9 +if ¥~±M¦b¤U­±¥[«HÀY«á¤~¦C¦L
      'Modified by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
      'Removed by Morgan 2025/8/19 ¤º°Ó§ï³æµ§©Î¾ã§å³£­n«HÀY --®Û­^
      'If Left(Pub_StrUserSt03, 2) = "P2" Then
      'end 2025/8/19
      
         If m_bPrintWord And m_iSpCopies > 0 Then
            .ActiveDocument.PrintOut Background:=False, Copies:=m_iSpCopies, Collate:=True
            m_iPageCount = .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) 'Added by Morgan 2015/2/26
         End If
      
      'End If
      'end 2020/5/7
      'end 2020/4/1
      
      'Modified by Morgan 2014/2/26 +¥~°Ó³£±a«HÀY
      'Modified by Morgan 2014/8/25 §ï¤Á´«¦Lªí¾÷¤è¦¡¦C¦LPDF,¤£¥Î­«¶]Word
      'If m_bWord2Pdf Or m_bSaveWord Or (m_bEditDoc And Left(Pub_StrUserSt03, 2) = "F1") Then
      'Modify By Sindy 2015/7/8 +¥~±M³£±a«HÀY( Or Left(Pub_StrUserSt03, 2) = "F2")
      'Removed by Morgan 2020/4/1 +¥~°Ó¤]­n±a«HÀY--³¯ª÷½¬
      'If m_b2PDF Or m_bWord2Pdf Or m_bSaveWord Or (m_bEditDoc And Left(Pub_StrUserSt03, 2) = "F1") Or (m_CP01 = "FCP" Or m_CP01 = "FG" Or m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") Then
      'Added by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
      'Removed by Morgan 2025/8/19 ¤º°Ó§ï³æµ§©Î¾ã§å³£­n«HÀY--®Û­^
      'If m_b2PDF Or m_bWord2Pdf Or m_bSaveWord Or Left(Pub_StrUserSt03, 2) <> "P2" Then
      'end 2025/8/19
      'end 2020/5/7
      'end 2020/4/1
         'Modified by Morgan 2020/3/31
         'If PUB_ReadDB2File(stFileName, 5) Then
         If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é Then
            PUB_GetLetterPicID "2", m_CP01, iPicNo, iPicNo2, 2, True, Pub_StrUserSt03
         Else
            iPicNo = 5
            iPicNo2 = 9
         End If
         If PUB_ReadDB2File(stFileName, iPicNo) Then
         'end 2020/3/31
         
            'Modified by Morgan 2014/3/4 ¨C­¶³£­n¦³«HÀY§À
            For ii = 1 To .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)
               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.ZOrder 4
               oShape.LockAnchor = True
               oShape.LockAspectRatio = -1
               oShape.Width = .CentimetersToPoints(21)
               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
               oShape.Left = .CentimetersToPoints(0)
               'Modify By Sindy 2015/11/9 ¦]¥~±M­n¥Î¶}µ¡«H«Ê
               'oShape.Top = .CentimetersToPoints(0.5)
               oShape.Top = .CentimetersToPoints(0)
               '2015/11/9 END
               oShape.WrapFormat.Type = wdWrapNone
               .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
               'Added by Lydia 2016/03/04 ¤Á³Î¶W¹L1­¶ªº©ú²Óªí®æ
               If ii < pCount Then
                  .Selection.Collapse Direction:=wdCollapseStart
                  .Selection.InsertBreak Type:=wdSectionBreakContinuous
                  '¦^¨ì¤W¤@­¶ªº¥½ºÝ,¼W¥[2¦æªÅ¥Õ
                  .Selection.GoTo what:=wdGoToLine, which:=wdGoToPrevious, Count:=1
                  .Selection.TypeParagraph
                  .Selection.TypeParagraph
               End If
            Next ii
            .Selection.HomeKey Unit:=wdStory
            
            'Modified by Morgan 2020/3/31
            'If PUB_ReadDB2File(stFileName, 9) Then
            If iPicNo2 > 0 Then
               If PUB_ReadDB2File(stFileName, iPicNo2) Then
            'end 2020/3/31
                  For ii = 1 To .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)
                     Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                     oShape.ZOrder 4
                     oShape.LockAnchor = True
                     oShape.LockAspectRatio = -1
                     oShape.Width = .CentimetersToPoints(21)
                     oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                     oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                     oShape.Left = .CentimetersToPoints(0)
                     'Added by Morgan 2020/3/31
                     If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é Then
                        oShape.Top = .CentimetersToPoints(27.6)
                     Else
                     'end 2020/3/31
                        oShape.Top = .CentimetersToPoints(27)
                     End If 'Added by Morgan 2020/3/31
                     oShape.WrapFormat.Type = wdWrapNone
                     .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
                  Next ii
                  .Selection.HomeKey Unit:=wdStory
               End If
               'end 2014/3/4
            End If 'Added by Morgan 2020/3/31
            
         'End If 'Removed by Morgan 2025/8/19
         
         'Add By Sindy 2015/7/9 ¥~±M¦b¦¹³B¤~¦C¦L
         'If (m_CP01 = "FCP" Or m_CP01 = "FG" Or m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") Then 'Removed by Morgan 2020/4/1 +¥~°Ó¤]­n±a«HÀY--³¯ª÷½¬
         'Added by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
         'Removed by Morgan 2025/8/19 ¤º°Ó§ï³æµ§©Î¾ã§å³£­n«HÀY--®Û­^
         'If Left(Pub_StrUserSt03, 2) <> "P2" Then
         'end 2020/5/7
            If m_bPrintWord And m_iSpCopies > 0 Then
               .ActiveDocument.PrintOut Background:=False, Copies:=m_iSpCopies, Collate:=True
               m_iPageCount = .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) 'Added by Morgan 2015/2/26
            End If
         'End If 'Added by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
         'End If 'Removed by Morgan 2020/4/1 +¥~°Ó¤]­n±a«HÀY--³¯ª÷½¬
         '2015/7/9 END
         
         If m_bWord2Pdf Then
            .ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
            
         'Added by Morgan 2014/8/26
         '§ï¤Á´«¦Lªí¾÷¤è¦¡¦C¦LPDF,¤£¥Î­«¶]Word
         ElseIf m_b2PDF Then
            PrintWord2PDF
         'end 2014/8/26
            
         End If
         
         If m_bSaveWord Then
            RidFile m_EFilePath
            .ActiveDocument.SaveAs m_EFilePath
         End If
         
      End If 'Added by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
      'End If 'Removed by Morgan 2020/4/1 +¥~°Ó¤]­n±a«HÀY--³¯ª÷½¬
   End With
   'Modified by Lydia 2019/04/09 §ï¦¨¦@¥Î¼Ò²Õ
   'RePosWord bVisible, m_WordLeft, m_WordTop 'Added by Morgan 2014/6/26
   Pub_RePosWord g_WordAp, bVisible, m_WordLeft, m_WordTop
   
   If m_bEditDoc Then
      g_WordAp.Visible = True
      g_WordAp.WindowState = wdWindowStateMaximize
      g_WordAp.Activate
   Else
      g_WordAp.ActiveDocument.Close wdDoNotSaveChanges
      If bVisible = False Then
         g_WordAp.Quit wdDoNotSaveChanges
         Set g_WordAp = Nothing 'Added by Lydia 2017/12/12 Á×§K§Ö³t¶}±ÒWord,µ{¦¡¥X¿ù
      Else
         g_WordAp.Visible = True
      End If
   End If
   
   Exit Sub
   
ErrHnd:
   If Err.Number <> 0 Then
'Modified by Morgan 2014/6/26
'      If iResumeCnt > 3 Then
'         MsgBox "¿ù»~ : " & Err.Description, vbCritical
'      Else
'         iResumeCnt = iResumeCnt + 1
'         Select Case Err.NUMBER
'            Case 91:
'               g_WordAp.Documents.Add
'               Resume Next
'            Case 462:
'               Set g_WordAp = New Word.Application
'               Resume
'            Case Else:
'               MsgBox "¿ù»~ : " & Err.Description, vbCritical
'         End Select
'      End If
      MsgBox "¿ù»~ : " & Err.Description, vbCritical
'end 2014/6/26
   End If
End Sub

'Modify by Morgan 2011/1/5 ¥[¥i¦C¦L¤Î¦sÀÉ
Private Sub runWordJapanese()

   'Added by Morgan 2013/10/25 ´ú¸Õ¥Î
   If bolNewForm Then
      runWordJapaneseNew
      Exit Sub
   End If
   'end 2013/10/25
   
   Dim stTmp As String
   Dim iRow As Integer, iCol As Integer, lngCellWidth As Long
   Dim strLastCountry As String, iFCol As Integer, iCols As Integer, bolPrintCountry As Boolean
   Dim strFontSize As String
   Dim iResumeCnt As Integer
   Dim bVisible As Boolean, stFileName As String, iHeadCount As Integer
   Dim oShape
   Dim ii As Integer 'Added by Morgan 2014/3/4
   Dim dblOffFeeSub As Double
   Dim dblAttFeeSub As Double
   
   strFontSize = 12
   
   'Modified by Lydia 2019/04/09 §ï¦¨¦@¥Î¼Ò²Õ
   'If NewDoc(bVisible, m_WordLeft, m_WordTop) = False Then Exit Sub 'Added by Morgan 2014/6/26
   If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = False Then Exit Sub
   
On Error GoTo ErrHnd

'Removed by Morgan 2014/6/26
'   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
'
'   g_WordAp.Documents.Add
'
'   bVisible = g_WordAp.Visible
'
'   '¤£Åã¥Ü¥i¯à·|¦³°ÝÃD
'   'If pub_OS = 1 Or m_bEditDoc Then
'      g_WordAp.Visible = True
'      g_WordAp.WindowState = wdWindowStateMaximize
'      g_WordAp.ActiveWindow.ActivePane.View.Zoom.Percentage = 100 'Added by Morgan 2012/11/30
'   'Else
'   '   g_WordAp.Visible = False
'   'End If
   
   With g_WordAp.Application
      'ª©­±³]©w
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.5)
      'Modified by Morgan 2014/1/23 ­¶­º¥[°ª(²Ä2­¶¦L¯È¥»·|À£¨ì«HÀY)
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(3.53)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.3)
      'Modify by Morgan 2011/5/13 ·s«H¯È¤W¤U³£¦³¹Ï
      '.Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(3)
      .Selection.PageSetup.FooterDistance = .CentimetersToPoints(3)
      
      .Selection.PageSetup.CharsLine = 40
      .Selection.PageSetup.LinesPage = 38
      
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = strFontSize
      '.Selection.ParagraphFormat.DisableLineHeightGrid = True
      
      '«O¯d«HÀYªÅ¶¡(2¦æ)
      '.Selection.TypeParagraph 'Removed by Morgan 2014/1/23 ­¶­º¤w¥[°ª
      .Selection.TypeParagraph
      
      '·s¼Wªí®æ(1*4)
      .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=4
      
      With .Selection.Tables(1)
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
      End With
    
      '³]©wªí®æ°ª«×Äæ¼e
      .Selection.SelectRow
      'Modified by Morgan 2014/2/20
      '.Selection.Cells.VerticalAlignment = wdAlignVerticalBottom
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.9), RulerStyle:=wdAdjustProportional
      .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(8.5), RulerStyle:=wdAdjustProportional
      .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
      .Selection.Cells(3).VerticalAlignment = wdCellAlignVerticalTop 'Added by Morgan 2012/3/20
      '.Selection.Cells(4).Width = .CentimetersToPoints(6)
      
      .Selection.InsertRows UBound(m_Head, 2) + 1
      .Selection.Collapse Direction:=wdCollapseStart
            
      .Selection.SelectRow
      .Selection.Cells.Merge
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.Cells.VerticalAlignment = wdAlignVerticalBottom
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.ParagraphFormat.SpaceBefore = .CentimetersToPoints(0.6)
      .Selection.ParagraphFormat.SpaceAfter = .CentimetersToPoints(0.6)
      .Selection.TypeText Text:="                           "
      .Selection.Font.Size = 18
      .Selection.Font.Bold = True
      .Selection.TypeText Text:=m_Title
      .Selection.Font.Size = strFontSize
      .Selection.TypeText Text:="              NO. " & m_strDN
      .Selection.Font.Bold = False
      .Selection.MoveRight Unit:=wdCharacter, Count:=2
      
      'ªíÀY¦C1
      For iRow = 1 To UBound(m_Head, 2)
         .Selection.Cells(1).SetHeight RowHeight:=12, HeightRule:=wdRowHeightAtLeast
         If m_Head(1, iRow) <> "" Then
            .Selection.TypeText Text:=m_Head(1, iRow)
         End If
         If m_Head(2, iRow) <> "" Then
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:=m_Head(2, iRow)
         Else
            .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
            .Selection.Cells.Merge
         End If
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         If m_Head(3, iRow) <> "" Then
            .Selection.TypeText Text:=m_Head(3, iRow)
         End If
         If m_Head(4, iRow) <> "" Then
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:=m_Head(4, iRow)
         Else
            .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
            .Selection.Cells.Merge
         End If
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
      Next

      .Selection.SelectRow
      .Selection.Cells.Merge
      .Selection.InsertRows 3
      .Selection.Collapse Direction:=wdCollapseStart
      With .Selection.Cells.Borders(wdBorderBottom)
         .LineStyle = wdLineStyleSingle
         .LineWidth = wdLineWidth050pt
         .ColorIndex = wdAuto
      End With
      .Selection.MoveRight Unit:=wdCharacter, Count:=2
      '¼ÐÃD
      For iRow = 1 To UBound(m_Subject, 2)
         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
         If m_Subject(1, iRow) <> "" Then
            .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
            If m_Subject(2, iRow) = "" Then
               .Selection.TypeText Text:=m_Subject(1, iRow)
            Else
               .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.TypeText Text:=m_Subject(1, iRow)
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.TypeText Text:=m_Subject(2, iRow)
            End If
            
         Else
            .Selection.Cells(1).SetHeight RowHeight:=16, HeightRule:=wdRowHeightAtLeast
            If m_Subject(2, iRow) <> "" Then
               .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.TypeText Text:=m_Subject(2, iRow)
               
            ElseIf m_Subject(3, iRow) <> "" Then
               .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.9), RulerStyle:=wdAdjustProportional
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.MoveRight Unit:=wdCell, Count:=2
               .Selection.TypeText Text:=m_Subject(3, iRow)
            End If
         End If
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
         .Selection.InsertRows 1
         .Selection.Collapse Direction:=wdCollapseStart
      Next
      '©ú²Ó
      For iRow = 1 To UBound(m_Item)
         'Modify By Sindy 2021/10/29
         If Right(m_Item(iRow).iNo, 2) = "99" Then
            dblOffFeeSub = dblOffFeeSub + Val(Format(m_Item(iRow).iAmt))
         Else
            dblAttFeeSub = dblAttFeeSub + Val(Format(m_Item(iRow).iAmt))
         End If
         '2021/10/29 END
         
         .Selection.ParagraphFormat.SpaceBefore = 6
         .Selection.ParagraphFormat.SpaceAfter = 6
         .Selection.Cells.VerticalAlignment = wdAlignVerticalBottom
         If m_Item(iRow).ICode <> "" Then
            .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=False
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(10.5), RulerStyle:=wdAdjustProportional
            'Modified by Lydia 2016/03/03 +m_bSpecial5
            If m_bSpecial1 Or m_bSpecial5 Then
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(12.1), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.4), RulerStyle:=wdAdjustProportional
            Else
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(10.5), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(4), RulerStyle:=wdAdjustProportional
            End If
            .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(1.1), RulerStyle:=wdAdjustProportional
         Else
            .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=False
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.5), RulerStyle:=wdAdjustProportional
            .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.1), RulerStyle:=wdAdjustProportional
         End If
         
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.ParagraphFormat.RightIndent = .CentimetersToPoints(0.6)
         .Selection.TypeText Text:=m_Item(iRow).IDesc
         
         If m_Item(iRow).ICode <> "" Then
            .Selection.MoveRight Unit:=wdCell, Count:=1
            '.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
            '.Selection.ParagraphFormat.RightIndent = .CentimetersToPoints(0.6)
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Selection.Font.Bold = True
            .Selection.TypeText Text:=m_Item(iRow).ICode
            .Selection.Font.Bold = False
         End If
         
         .Selection.MoveRight Unit:=wdCell, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.TypeText Text:=m_Item(iRow).iCur
         
         .Selection.MoveRight Unit:=wdCell, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.TypeText Text:=m_Item(iRow).iAmt
         
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
         .Selection.InsertRows 1
         .Selection.Cells.Merge
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.ParagraphFormat.RightIndent = .CentimetersToPoints(0)
      Next
      UpdateA1K3940 dblAttFeeSub, dblOffFeeSub 'Add by Sindy 2021/10/29 §ó·s¥~¹ôªA°È¶O³W¶Oª÷ÃB
      
      '¦X­p
      .Selection.Cells(1).SetHeight RowHeight:=18, HeightRule:=wdRowHeightAtLeast
      .Selection.SelectRow
      With .Selection.Cells.Borders(wdBorderTop)
         .LineStyle = wdLineStyleSingle
         .LineWidth = wdLineWidth050pt
         .ColorIndex = wdAuto
      End With
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.5), RulerStyle:=wdAdjustProportional
      .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.1), RulerStyle:=wdAdjustProportional
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.Font.Bold = True
      .Selection.Font.Size = 13
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.TypeText Text:="                                " & m_Sum(1, 1)
      .Selection.Font.Size = strFontSize
      .Selection.Font.Bold = False
      
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.TypeText Text:=m_Sum(2, 1)
      
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.TypeText Text:=m_Sum(3, 1)
      .Selection.MoveRight Unit:=wdCharacter, Count:=2
      If UBound(m_Sum, 2) > 1 Then
         .Selection.Cells(1).SetHeight RowHeight:=18, HeightRule:=wdRowHeightAtLeast
         .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.5), RulerStyle:=wdAdjustProportional
         .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.1), RulerStyle:=wdAdjustProportional
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.ParagraphFormat.RightIndent = .CentimetersToPoints(0.6)
         .Selection.TypeText Text:=m_Sum(1, 2)
         
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.TypeText Text:=m_Sum(2, 2)
         
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         .Selection.TypeText Text:=m_Sum(3, 2)
         
         UpdateA1K38 m_strDN, Format(m_Sum(3, 2)) 'Added by Morgan 2021/1/14 §ó·s½Ð´Ú³æ¬üª÷Á`ÃB
         
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.InsertRows 1
         .Selection.ParagraphFormat.RightIndent = .CentimetersToPoints(0)
         .Selection.Collapse Direction:=wdCollapseStart
      End If
      
      .Selection.SelectRow
      .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
      .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.5), RulerStyle:=wdAdjustProportional
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
      .Selection.Paragraphs.Alignment = wdAlignParagraphDistribute
      .Selection.TypeText Text:="vvvvvvvvvvvvv"
      
      'ªí§À1
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.InsertRows 1
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.Collapse Direction:=wdCollapseStart
      With .ActiveDocument.Bookmarks
         .add Range:=g_WordAp.Application.Selection.Range, Name:="BreakPos"
         .DefaultSorting = wdSortByLocation
         .ShowHidden = False
      End With
      .Selection.TypeText Text:=m_Footer(1, 1)
      
      If m_Footer(2, 1) <> "" Then
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.Cells.Split NumRows:=3, NumColumns:=2, MergeBeforeSplit:=False
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.Cells(1).SetHeight RowHeight:=36, HeightRule:=wdRowHeightAtLeast
         .Selection.MoveDown Unit:=wdLine, Count:=1
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
         With .Selection.Cells
             With .Borders(wdBorderLeft)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             With .Borders(wdBorderRight)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             With .Borders(wdBorderTop)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             With .Borders(wdBorderBottom)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
             .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
             .Borders.Shadow = False
         End With
         .Selection.ParagraphFormat.LeftIndent = .CentimetersToPoints(0.2)
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
         .Selection.Cells(1).SetHeight RowHeight:=52, HeightRule:=wdRowHeightAtLeast
         .Selection.TypeText Text:=m_Footer(2, 1)
         .Selection.HomeKey
         .Selection.MoveDown Unit:=wdLine, Count:=1
         .Selection.Cells(1).SetHeight RowHeight:=0, HeightRule:=wdRowHeightAtLeast
         .Selection.MoveDown Unit:=wdLine, Count:=1
      Else
         .Selection.MoveRight Unit:=wdCharacter, Count:=3
      End If
      
      '³Æµù
      If UBound(m_Footer, 2) > 1 Then
         .Selection.SelectRow
         .Selection.Font.Bold = True
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
         .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.TypeText Text:="PS:"
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.TypeText Text:=m_Footer(1, 2)
         For iRow = 3 To UBound(m_Footer, 2)
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.InsertRows 1
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:=m_Footer(1, iRow)
         Next
         .Selection.Font.Bold = False
      End If
      
      .Selection.WholeStory
      .Selection.Font.Name = "²Ó©úÅé"
      .Selection.WholeStory
      .Selection.Font.Name = "Times New Roman"
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .ActiveDocument.Repaginate
      '¶W¹L1­¶®É´¡¤J­¶½X
      If .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) > 1 Then
         .Selection.GoTo what:=wdGoToBookmark, Name:="BreakPos"
         .Selection.InsertBreak Type:=wdPageBreak
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .ActiveDocument.Repaginate
         If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
            .ActiveWindow.ActivePane.View.Type = wdPageView
         Else
            .ActiveWindow.View.Type = wdPageView
         End If
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
         .Selection.TypeParagraph
         .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldPage
         .Selection.TypeText Text:="/"
         .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldEmpty, Text:="NUMPAGES ", PreserveFormatting:=True
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
      End If
      .ActiveDocument.Bookmarks("BreakPos").Delete
      '.ActiveWindow.View.TableGridlines = False
      .Selection.HomeKey Unit:=wdStory
      
      'Removed by Morgan 2020/4/1 +¥~°Ó¤]­n±a«HÀY--³¯ª÷½¬
      'If Not (m_CP01 = "FCP" Or m_CP01 = "FG" Or m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") Then 'Add By Sindy 2015/7/9 +if ¥~±M¦b¤U­±¥[«HÀY«á¤~¦C¦L
      'Modified by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
      'Removed by Morgan 2025/8/19 ¤º°Ó§ï³æµ§©Î¾ã§å³£­n«HÀY--®Û­^
      'If Left(Pub_StrUserSt03, 2) = "P2" Then
      'end 2025/8/19
      
         If m_bPrintWord And m_iSpCopies > 0 Then
            .ActiveDocument.PrintOut Background:=False, Copies:=m_iSpCopies, Collate:=True
            m_iPageCount = .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) 'Added by Morgan 2015/2/26
         End If
         
      'End If
      'end 2020/5/7
      'end 2020/4/1
      
      'Modified by Morgan 2014/2/26 +¥~°Ó³£±a«HÀY
      'Modified by Morgan 2014/8/25 §ï¤Á´«¦Lªí¾÷¤è¦¡¦C¦LPDF,¤£¥Î­«¶]Word
      'If m_bWord2Pdf Or m_bSaveWord Or (m_bEditDoc And Left(Pub_StrUserSt03, 2) = "F1") Then
      'Modify By Sindy 2015/7/8 +¥~±M³£±a«HÀY( Or Left(Pub_StrUserSt03, 2) = "F2")
      'Removed by Morgan 2020/4/1 +¥~°Ó¤]­n±a«HÀY--³¯ª÷½¬
      'If m_b2PDF Or m_bWord2Pdf Or m_bSaveWord Or (m_bEditDoc And Left(Pub_StrUserSt03, 2) = "F1") Or (m_CP01 = "FCP" Or m_CP01 = "FG" Or m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS" Or m_CP01 = "FCT") Then
      'Added by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
      'Removed by Morgan 2025/8/19 ¤º°Ó§ï³æµ§©Î¾ã§å³£­n«HÀY--®Û­^
      'If m_b2PDF Or m_bWord2Pdf Or m_bSaveWord Or Left(Pub_StrUserSt03, 2) <> "P2" Then
      'end 2025/8/19
      'end 2020/5/7
      'end 2020/4/1
         'Modified by Morgan 2020/3/31
         'If PUB_ReadDB2File(stFileName, 5) Then
         If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é Then
            PUB_GetLetterPicID "2", m_CP01, iPicNo, iPicNo2, 3, True, Pub_StrUserSt03
         Else
            iPicNo = 5
            iPicNo2 = 9
         End If
         If PUB_ReadDB2File(stFileName, iPicNo) Then
         'end 2020/3/31
            'Modified by Morgan 2014/3/4 ¨C­¶³£­n¦³«HÀY§À
            For ii = 1 To .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)
               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.ZOrder 4
               oShape.LockAnchor = True
               oShape.LockAspectRatio = -1
               oShape.Width = .CentimetersToPoints(21)
               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
               oShape.Left = .CentimetersToPoints(0)
               oShape.Top = .CentimetersToPoints(0.5)
               .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
            Next ii
            .Selection.HomeKey Unit:=wdStory
               
            'If PUB_ReadDB2File(stFileName, 9) Then
            If iPicNo2 > 0 Then
               If PUB_ReadDB2File(stFileName, iPicNo2) Then
            'end 2020/3/31
                  For ii = 1 To .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)
                     Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                     oShape.ZOrder 4
                     oShape.LockAnchor = True
                     oShape.LockAspectRatio = -1
                     oShape.Width = .CentimetersToPoints(21)
                     oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                     oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                     oShape.Left = .CentimetersToPoints(0)
                     'Added by Morgan 2020/3/31
                     If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é Then
                        oShape.Top = .CentimetersToPoints(27.6)
                     Else
                     'end 2020/3/31
                        oShape.Top = .CentimetersToPoints(27)
                     End If 'Added by Morgan 2020/3/31
                     .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
                  Next ii
                  .Selection.HomeKey Unit:=wdStory
               End If
               'end 2014/3/4
            End If 'Added by Morgan 2020/3/31
            
         End If
         
         'Add By Sindy 2015/7/9 ¥~±M¦b¦¹³B¤~¦C¦L
         'If (m_CP01 = "FCP" Or m_CP01 = "FG" Or m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") Then 'Removed by Morgan 2020/4/1 +¥~°Ó¤]­n±a«HÀY--³¯ª÷½¬
         'Added by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
         'Removed by Morgan 2025/8/19 ¤º°Ó§ï³æµ§©Î¾ã§å³£­n«HÀY--®Û­^
         'If Left(Pub_StrUserSt03, 2) <> "P2" Then
         'end 2025/8/19
         'end 2020/5/7
         
            If m_bPrintWord And m_iSpCopies > 0 Then
               .ActiveDocument.PrintOut Background:=False, Copies:=m_iSpCopies, Collate:=True
               m_iPageCount = .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) 'Added by Morgan 2015/2/26
            End If
            
         'End If 'Added by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
         'End If 'Removed by Morgan 2020/4/1 +¥~°Ó¤]­n±a«HÀY--³¯ª÷½¬
         '2015/7/9 END
         
         If m_bWord2Pdf Then
            .ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
            
         'Added by Morgan 2014/8/26
         '§ï¤Á´«¦Lªí¾÷¤è¦¡¦C¦LPDF,¤£¥Î­«¶]Word
         ElseIf m_b2PDF Then
            PrintWord2PDF
         'end 2014/8/26
         
         End If
         
         If m_bSaveWord Then
            RidFile m_EFilePath
            .ActiveDocument.SaveAs m_EFilePath
         End If
         
      'End If 'Added by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
      'End If 'Removed by Morgan 2020/4/1 +¥~°Ó¤]­n±a«HÀY--³¯ª÷½¬
      
   End With
   'Modified by Lydia 2019/04/09 §ï¦¨¦@¥Î¼Ò²Õ
   'RePosWord bVisible, m_WordLeft, m_WordTop 'Added by Morgan 2014/6/26
   Pub_RePosWord g_WordAp, bVisible, m_WordLeft, m_WordTop
   
   If m_bEditDoc Then
      g_WordAp.Visible = True
      g_WordAp.WindowState = wdWindowStateMaximize
      g_WordAp.Activate
   Else
      g_WordAp.ActiveDocument.Close wdDoNotSaveChanges
      If bVisible = False Then
         g_WordAp.Quit wdDoNotSaveChanges
         Set g_WordAp = Nothing 'Added by Lydia 2017/12/12 Á×§K§Ö³t¶}±ÒWord,µ{¦¡¥X¿ù
      Else
         g_WordAp.Visible = True
      End If
   End If

   
   Exit Sub
   
ErrHnd:
   If Err.Number <> 0 Then
'Modified by Morgan 2014/6/26
'      If iResumeCnt > 3 Then
'         MsgBox "¿ù»~ : " & Err.Description, vbCritical
'      Else
'         iResumeCnt = iResumeCnt + 1
'         Select Case Err.NUMBER
'            Case 91:
'               g_WordAp.Documents.Add
'               Resume Next
'            Case 462:
'               Set g_WordAp = New Word.Application
'               Resume
'            Case Else:
'               MsgBox "¿ù»~ : " & Err.Description, vbCritical
'         End Select
'      End If
      MsgBox "¿ù»~ : " & Err.Description, vbCritical
'end 2014/6/26
   End If
End Sub
'Add by Morgan 2010/11/24
'Àx¦sWord®æ¦¡¸ê®Æ
Private Sub SetWordArray(pArray() As String, pRow As Integer, pIndex As Integer, pData As String)
   Dim iSizeD1 As Integer
   'Add by Morgan 2010/11/24
   If m_b2Word Then
      If UBound(pArray, 2) < pRow Then
         iSizeD1 = UBound(pArray, 1)
         ReDim Preserve pArray(iSizeD1, pRow)
      End If
      pArray(pIndex, pRow) = pArray(pIndex, pRow) & pData
   End If
End Sub

'Add by Morgan 2011/2/23
'Àx¦sWord®æ¦¡¸ê®Æ(½Ð´Ú¶µ¥Ø)
'pIndex:1=IDesc,2=ICode,3=ICur,4=IAmt,5=IDate
Private Sub SetItemWordArray(pArray() As INVITEM, pRow As Integer, pIndex As Integer, pData As String)
   Dim iSizeD1 As Integer
   'Add by Morgan 2010/11/24
   If m_b2Word Then
      If UBound(pArray) < pRow Then
         ReDim Preserve pArray(pRow)
      End If
      Select Case pIndex
         Case 1
            pArray(pRow).IDesc = pArray(pRow).IDesc & pData
            
         Case 2
            pArray(pRow).ICode = pArray(pRow).ICode & pData
         
         Case 3
            pArray(pRow).iCur = pArray(pRow).iCur & pData
            
         Case 4
            pArray(pRow).iAmt = pArray(pRow).iAmt & pData
         
         Case 5
            pArray(pRow).IDate = pArray(pRow).IDate & pData
            
         'Added by Morgan 2013/10/29
         Case 6
            pArray(pRow).IDescHead = pArray(pRow).IDescHead & pData
         Case 7
            pArray(pRow).IDescTail = pArray(pRow).IDescTail & pData
         Case 8
            pArray(pRow).iNo = pArray(pRow).iNo & pData
         Case 9
            pArray(pRow).IDescX = pArray(pRow).IDescX & pData
         'end 2013/10/29
         'Added by Morgan 2014/8/28
         Case 10
            pArray(pRow).IXAmt = pArray(pRow).IXAmt & pData
'Added by Lydia 2015/04/09 ¤¤¤åª©-¾ã§å½Ð´Ú³æ
         Case 11
            pArray(pRow).IChiCno = pArray(pRow).IChiCno & pData  '®×¸¹
         Case 12
            pArray(pRow).IChiCna = pArray(pRow).IChiCna & pData '¦WºÙ
         Case 13
            pArray(pRow).IChiCls = pArray(pRow).IChiCls & pData 'Ãþ§O
         Case 14
            pArray(pRow).IChiApp = pArray(pRow).IChiApp & pData 'µù¥U¸¹/¥Ó½Ð®×¸¹
         Case 15
            pArray(pRow).IChiA1k01 = pArray(pRow).IChiA1k01 & pData '½Ð´Ú³æ¸¹
         Case 16
            pArray(pRow).IChiAmt = pArray(pRow).IChiAmt & pData '½Ð´Úª÷ÃB(­ì¹ô§O)
         Case 17
            pArray(pRow).IChiUAmt = pArray(pRow).IChiUAmt & pData '½Ð´Úª÷ÃB
         'Added by Morgan 2016/8/5
         Case 18
            pArray(pRow).INtAmt = pArray(pRow).INtAmt & pData '¥x¹ô½Ð´Úª÷ÃB
         'Added by Morgan 2020/8/6
         Case 19
            pArray(pRow).iAmtNoDisc = pArray(pRow).iAmtNoDisc & pData '§é¦©«e½Ð´Úª÷ÃB
      End Select
   End If
   
End Sub

'Add by Morgan 2010/12/23
'¶µ¥Ø»¡©ú¸ÑªR
'Removed by Morgan 2025/8/21 §ï¤½¥Î¨ç¼Æ PUB_ParseItemDesc
'Private Function ParseDesc(ByVal pAmount As Double, ByVal pDesc As String, Optional pWholeDesc As String) As String
'   Dim iPos1 As Integer, iPos2 As Integer
'   Dim strUnit As String, strUnitPrice As String, strQty As String
'   Dim dblUnit As Double, dblUnitPrice As Double
'   If InStr(pDesc, "[QTY]") > 0 Or InStr(pDesc, "[$") > 0 Or InStr(pDesc, "[#") > 0 Then
'      '³æ»ù
'      iPos1 = InStr(pDesc, "[$")
'      If iPos1 > 0 Then
'         iPos2 = InStr(pDesc, "$]")
'         If iPos2 > iPos1 Then
'            strUnitPrice = Mid(pDesc, iPos1 + 2, iPos2 - iPos1 - 2)
'            If IsNumeric(strUnitPrice) Then
'               dblUnitPrice = Format(strUnitPrice)
'            End If
'            pDesc = Left(pDesc, iPos1 - 1) & strUnitPrice & Mid(pDesc, iPos2 + 2)
'         End If
'      End If
'      '³æ¦ì
'      dblUnit = 1
'      iPos1 = InStr(pDesc, "[#")
'      If iPos1 > 0 Then
'         iPos2 = InStr(pDesc, "#]")
'         If iPos2 > iPos1 Then
'            strUnit = Mid(pDesc, iPos1 + 2, iPos2 - iPos1 - 2)
'            If IsNumeric(strUnit) Then
'               dblUnit = Format(strUnit)
'            End If
'            pDesc = Left(pDesc, iPos1 - 1) & dblUnit & Mid(pDesc, iPos2 + 2)
'         End If
'      End If
'      '¼Æ¶q
'      If pAmount > 0 And dblUnitPrice > 0 And dblUnit > 0 Then
'         strQty = Round(pAmount / dblUnitPrice * dblUnit)
'      End If
'      pDesc = Replace(pDesc, "[QTY]", strQty)
'   End If
'   ParseDesc = pDesc
'End Function

'Add by Morgan 2011/1/6
Private Function ChkIs103(pPA01 As String, pPA02 As String, pPA03 As String, pPA04 As String) As Boolean
   Dim stSQL As String, iR As Integer, rstX As ADODB.Recordset
   
   stSQL = "select pa08 from patent where pa01='" & pPA01 & "'" & _
      " and pa02='" & pPA02 & "' and pa03='" & pPA03 & "' and pa04='" & pPA04 & "'"
   iR = 1
   Set rstX = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      If rstX.Fields(0) = "3" Then
         ChkIs103 = True
      End If
   End If
   Set rstX = Nothing
End Function
'Add by Morgan 2011/9/26
'½Ð´Ú³æ³Ì¦­¦¬¤å¤é
Private Function GetMinCp05(pBillNo As String) As String
   Dim stSQL As String, intR As Integer
   Dim rsTmp As ADODB.Recordset
   stSQL = "select min(CP05) from caseprogress where cp60='" & pBillNo & "'"
   intR = 1
   Set rsTmp = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      GetMinCp05 = rsTmp.Fields(0)
   End If
End Function
'Added by Morgan 2018/6/25
'Modified by Morgan 2022/8/5 +pMax:§ì³Ì¤jµo¤å¤é
'Modified by Morgan 2024/6/24 +pAKind:¥u§ìAÃþ
'½Ð´Ú³æ³Ì¦­µo¤å¤é
Private Function GetCp27(pBillNo As String, Optional pCP10 As String, Optional pMax As Boolean = False, Optional pAKind As Boolean = False) As String
   Dim stSQL As String, intR As Integer, stCon As String
   Dim rsTmp As ADODB.Recordset
   
   If pCP10 <> "" Then
      If Right(pCP10, 2) = "99" Then
         stCon = " and cp10='" & Left(pCP10, Len(pCP10) - 2) & "'"
      Else
         stCon = " and cp10='" & pCP10 & "'"
      End If
   End If
   
   If pAKind Then stCon = stCon & " and cp09<'B'" 'Added by Morgan 2024/6/24
   
   stSQL = "select cp27 from caseprogress where cp60='" & pBillNo & "' and cp27>0" & stCon
   If pMax Then
      stSQL = stSQL & " order by cp27 desc"
   Else
      stSQL = stSQL & " order by cp27 asc"
   End If
   intR = 1
   Set rsTmp = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      GetCp27 = rsTmp.Fields(0)
   End If
End Function

'Added by Morgan 2012/10/31
Private Sub MyNewPage(Optional bolIs1st As Boolean)
   Dim iPicNo As Integer, iPicNo2 As Integer
   Dim strFolder As String
   Dim strFileName As String
   
   'Added by Morgan 2020/3/31
   If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é Then
      PUB_GetLetterPicID "2", m_CP01, iPicNo, iPicNo2, 2, True, Pub_StrUserSt03
   Else
   'end 2020/3/31
      iPicNo = 5
      iPicNo2 = 9
   End If 'Added by Morgan 2020/3/31
   
   If bolIs1st = False Then
      Printer.NewPage
      
      'Added by Morgan 2013/5/13 ±M§Q³Bµ{§Ç¦C¦L½Ð´Ú³æ®É,­n±a¥X«HÀY
      'Modify By Sindy 2015/7/13 ¶®®S¸ò¨q¬Â»¡­n¥Î±M§Qªk«ß«HÀY
      'If Pub_StrUserSt03 = "P12" And m_b2Printer Then PrintPicture 7  'Added by Morgan 2013/5/13 ±M§Q³Bµ{§Ç¦C¦L½Ð´Ú³æ®É,­n±a¥X«HÀY
      If (m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") And m_b2Printer Then
         PrintPicture iPicNo, iPicNo2
      End If
   
   End If
   
   If m_bPrint2Pdf Then
      If bolIs1st Then
         m_EFilePath = GetPath
         strFolder = m_EFilePath
         If Dir(strFolder, vbDirectory) = "" Then
            MkDir strFolder
         End If
         strFileName = m_strCaseNo & IIf(m_bAddDate, "_" & strSrvDate(1), "") & "_DN" & m_strDN
         frmPDF.Show
         frmPDF.StartProcess strFolder, strFileName
         Me.Tag = strFolder & "\" & strFileName & ".pdf" 'Added by Lydia 2017/02/18 °O¿ý-½Ð´Ú³æpdfÀÉ¸ô®|
         EfileNameFCP_08 = EfileNameFCP_08 & ";" & m_EFilePath & "\" & strFileName & ".pdf" 'Add By Sindy 2015/7/9
      End If
      
      If Not m_bWord2Pdf Then
         'Added by Morgan 2020/3/31
         If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é Then
         '²Î¤@¥Îipdept
         Else
         'end 2020/3/31
            If Left(m_A1k28, 1) = "X" Then
               strExc(0) = "SELECT CU10 FROM CUSTOMER WHERE CU01='" & Left(m_A1k28, 8) & "' and CU02='" & Mid(m_A1k28, 9) & "'"
            Else
               strExc(0) = "SELECT FA10 FROM FAGENT WHERE FA01='" & Left(m_A1k28, 8) & "' and FA02='" & Mid(m_A1k28, 9) & "'"
            End If
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp(0) = "020" Then
                  iPicNo = 7
                  iPicNo2 = 0
               End If
            End If
         End If
         PrintPicture iPicNo, iPicNo2
      End If
   End If
End Sub
'Added by Morgan 2012/10/31
Private Sub PrintPicture(iPicNo As Integer, Optional iPicNo2 As Integer)
   Dim tObj As New StdPicture, stFileName As String, pWidth As Long, pHeight As Long
   pWidth = 21 * 567
   If PUB_ReadDB2File(stFileName, iPicNo) Then
      Set tObj = pvGetStdPicture(stFileName)
      pHeight = tObj.Height * (pWidth / tObj.Width)
      Printer.PaintPicture tObj, 0, Int(0.5 * 567), pWidth, pHeight
   End If
   If iPicNo2 > 0 Then
      If PUB_ReadDB2File(stFileName, iPicNo2) Then
         Set tObj = pvGetStdPicture(stFileName)
         pHeight = tObj.Height * (pWidth / tObj.Width)
         Printer.PaintPicture tObj, 0, Int(27 * 567), pWidth, pHeight
      End If
   End If
   Set tObj = Nothing
End Sub

'Added by Morgan 2014/6/24
'Mark by Lydia 2019/04/09 §ï¦¨¦@¥Î¼Ò²Õ
'Private Function NewDoc(Optional ByRef pVisible As Boolean, Optional ByRef PLeft As Long, Optional ByRef pTop As Long) As Boolean
'
'   Dim iResumeCnt As Integer
'
'On Error GoTo ErrHnd
'
'   If TypeName(g_WordAp) <> "Application" Then
'      Set g_WordAp = New Word.Application
'   End If
'   PLeft = g_WordAp.Left
'   pTop = g_WordAp.Top
'
'   g_WordAp.Documents.add
'
'   pVisible = g_WordAp.Visible
'
'   '¤£Åã¥Ü¥i¯à·|¦³°ÝÃD
'   'Modified by Morgan 2015/6/22
'   'g_WordAp.Visible = False
'   'If m_bShowWord Then
'      g_WordAp.Visible = True
'      'If Me.Check1.Value = vbChecked Then
'      If Not g_LetterDebug Then
'         g_WordAp.WindowState = wdWindowStateNormal
'         If g_WordAp.WindowState = wdWindowStateNormal Then
'            g_WordAp.Move Screen.Width / 20, Screen.Height / 20
'         End If
'      End If
'      'Else
'      '   g_WordAp.WindowState = wdWindowStateMaximize
'      'End If
'
'      g_WordAp.ActiveWindow.ActivePane.View.Zoom.Percentage = 100
'   'End If
'   'end 2015/6/22
'   NewDoc = True
'   Exit Function
'
'ErrHnd:
'   'Resume
'   If Err.Number <> 0 Then
'      If iResumeCnt > 3 Then
'         MsgBox "¿ù»~ : " & Err.Description, vbCritical
'      Else
'         iResumeCnt = iResumeCnt + 1
'         Select Case Err.Number
'            Case 91:
'               g_WordAp.Documents.add
'               Resume Next
'            Case 462:
'               Set g_WordAp = New Word.Application
'               Resume
'            Case Else:
'               MsgBox "¿ù»~ : " & Err.Description, vbCritical
'         End Select
'      End If
'   End If
'End Function
''Added by Morgan 2014/6/24
'Private Sub RePosWord(ByRef pVisible As Boolean, ByRef PLeft As Long, ByRef pTop As Long)
'   g_WordAp.Visible = pVisible
'   'Modified by Morgan 2015/6/22
'   'If Me.Check1 = vbChecked Then
'      g_WordAp.Move PLeft, pTop
'   'End If
'   'end 2015/6/22
'End Sub
'end 2019/04/09

'Added by Morgan 2013/10/25 ·s®æ¦¡
Private Sub runWordEnglishNew()
   Dim iRow As Integer, iCol As Integer, lngCellWidth As Long
   Dim strLastCountry As String, iFCol As Integer, iCols As Integer, bolPrintCountry As Boolean
   Dim strFontSize As String
   Dim iResumeCnt As Integer
   Dim bVisible As Boolean, stFileName As String, iHeadCount As Integer
   Dim strText As String, strText1 As String
   Dim oShape
   Dim strLstStr As String
   Dim strLstNo As String
   Dim dblOffFeeSub As Double
   Dim dblAttFeeSub As Double
   Dim dblNtOffFeeSub As Double
   Dim dblNtAttFeeSub As Double
   Dim bolNewRow As Boolean
   Dim ii As Integer
   Dim iRowCount As Integer
   'Added by Morgan 2019/3/28
   Dim ItemCache As INVITEM
   Dim bolPrintActSub As Boolean
   Dim iMergeRows As Integer
   Dim bolMergeRow As Boolean
   Dim dblOffFeeActSub As Double
   Dim dblAttFeeActSub As Double
   'end 2019/3/28
   Dim dblA1K39 As Double, dblA1K40 As Double 'Add By Sindy 2021/10/26
   
   'Added by Morgan 2022/5/18
   If (m_A1k28 = "Y55666000" And m_CallPrevForm = "Frmacc24l0") Then
      ii = 1
      For iRow = 1 To UBound(m_Head, 2)
         m_Head(1, iRow) = ""
         'Modified by Morgan 2022/11/24
         'm_Head(2, iRow) = ""
         If m_Head(2, iRow) <> "" Then
            If ii < iRow Then
               m_Head(2, ii) = m_Head(2, iRow)
               m_Head(2, iRow) = ""
               ii = ii + 1
            End If
         End If
         
         If InStr("Date:,Your Ref:,Our Ref:", Trim(m_Head(3, iRow))) = 0 Then
            m_Head(3, iRow) = ""
            m_Head(4, iRow) = ""
         End If
      Next
      
      'Added by Morgan 2022/11/24
      m_Head(3, 1) = "Service period: " & Format(ChangeTStringToWDateString(m_strA1K02), "mmmm, yyyy")
      m_Head(2, ii + 1) = PUB_GetACCNO(m_A1k28)
      'end 2022/11/24
      m_Title = "          "
   End If
   'end 2022/5/18
   
   strFontSize = 12
      
   'Modified by Lydia 2019/04/09 §ï¦¨¦@¥Î¼Ò²Õ
   'If NewDoc(bVisible, m_WordLeft, m_WordTop) = False Then Exit Sub 'Added by Morgan 2014/6/26
   If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = False Then Exit Sub
   
On Error GoTo ErrHnd
   
   With g_WordAp.Application
      'ª©­±³]©w
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.5)
      'Modified by Morgan 2014/1/23 ­¶­º¥[°ª(²Ä2­¶¦L¯È¥»·|À£¨ì«HÀY)
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(3.53)
      'Modify By Sindy 2015/11/9 ¦]¥~±M­n¥Î¶}µ¡«H«Ê
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(4.3)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4)
      '2015/11/9 END
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(3)
      .Selection.PageSetup.FooterDistance = .CentimetersToPoints(3)
      
      .Selection.PageSetup.CharsLine = 40
      .Selection.PageSetup.LinesPage = 38
      
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = strFontSize
      
      'Added by Lydia 2020/06/23 «HÀY¤U¤è¥[¦L¡uPAID³¹¡v
      If m_bPAID = True Then
          If PUB_ReadDB2File(stFileName, "57") = True Then
               'ªÅ4¦æ
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               .Selection.TypeParagraph

               .Selection.MoveUp Unit:=wdLine, Count:=3 '²¾¨ì²Ä¤@¦æ
               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.Left = .CentimetersToPoints(13.5)
               .Selection.MoveDown Unit:=wdLine, Count:=3 '²¾¨ì³Ì«á¤@¦æ
          Else
               .Selection.TypeParagraph
          End If
      Else
      'end 2020/06/23
         '«O¯d«HÀYªÅ¶¡(2¦æ)
         '.Selection.TypeParagraph 'Removed by Morgan 2014/1/23 ­¶­º¤w¥[°ª
         .Selection.TypeParagraph '¤£¥iMark,¦]Word2007¦b´¡«HÀY®É¦ì¸m·|¿ù,·|´¡¤J¨ìªí®æ¸Ì
      End If 'Added by Lydia 2020/06/23
      
      '¦æ¶Z
      With .Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
        .DisableLineHeightGrid = True
      End With
      
      '·s¼Wªí®æ(1*4)
      .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=4
      
      With .Selection.Tables(1)
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
      End With
    
      '³]©wªí®æ°ª«×Äæ¼e
      .Selection.SelectRow
      'Modified by Morgan 2014/2/20
      '.Selection.Cells.VerticalAlignment = wdAlignVerticalBottom
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.9), RulerStyle:=wdAdjustProportional
      'Modified by Morgan 2021/7/1 BASF ¥kÃä¥[¼e Ex:X11006502 --Anny
      If m_A1k28 = "Y45814010" Or m_A1k28 = "Y33268010" Then
         .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(8.2), RulerStyle:=wdAdjustProportional
      Else
         .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(8.7), RulerStyle:=wdAdjustProportional
      End If
      'end 2021/7/2
      
      'Added by Morgan 2022/8/5
      'Y55751 Birkenstock IP GmbH--Franny,Peggy
      If m_A1k28 = "Y55751000" Then
         .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.6), RulerStyle:=wdAdjustProportional
      Else
      'end 2022/8/5
      
         .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
      End If 'Added by Morgan 2022/8/5
      .Selection.Cells(3).VerticalAlignment = wdCellAlignVerticalTop
      
      .Selection.InsertRows UBound(m_Head, 2)
      .Selection.Collapse Direction:=wdCollapseStart
      iRowCount = 0 'Added by Morgan 2014/2/20
      'ªíÀY¦C1
      For iRow = 1 To UBound(m_Head, 2)
         .Selection.Cells(1).SetHeight RowHeight:=12, HeightRule:=wdRowHeightAtLeast
         If m_Head(1, iRow) <> "" Then
            .Selection.TypeText Text:=m_Head(1, iRow)
         End If
         If m_Head(2, iRow) <> "" Then
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:=m_Head(2, iRow)
            
            'Added by Morgan 2014/2/20 Äæ¦ì¦X¨Ö§_«h­Y©¼©Ò®×¸¹©Î«È¤á®×¥ó®×¸¹¤Óªø¦³§é¦æ®É¦a§}·|¹jªÅ¥Õ¦æ Ex: X10302374
            If iRow > 1 Then
               If m_Head(2, iRow - 1) <> "" Then
                  .Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
                  .Selection.Cells.Merge
                  iRowCount = iRowCount + 1
               End If
            End If
            'end 2014/2/20
            
         Else
            iRowCount = 0 'Added by Morgan 2016/2/17
            .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
            .Selection.Cells.Merge
         End If
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         
         'Added by Morgan 2014/2/20
         If iRowCount > 0 Then
            'Modified by Morgan 2019/12/5 «È¤á®×¥ó®×¸¹­Y¦X¨ÖÄæ¦ì¥B¦³§é¦æ®É¦C¼Æ·|¼W¥[ Ex: X10818755
            '.Selection.MoveDown unit:=wdLine, Count:=intI
            For ii = 1 To iRowCount
               .Selection.MoveEnd Unit:=wdCell
               .Selection.MoveDown Unit:=wdLine, Count:=1
            Next
            'end 2019/12/5
         End If
         'end 2014/2/20
         
         If m_Head(3, iRow) <> "" Then
            'Added by Morgan 2017/7/7
            'Modfiied by Morgan 2021/7/1 +Y3326801 BASF Corporation --Anny
            'Modified by Morgan 2021/12/7 +Y51467020 Saurer Spinning Solutions GmbH & Co. KG --Franny
            If m_A1k28 = "Y45814010" Or m_A1k28 = "Y33268010" Or m_A1k28 = "Y51467020" Then
               If InStr(m_Head(3, iRow), "Service Period:") = 1 Then
                  .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
                  'Modified by Morgan 2021/12/8 ­×¥¿§Ç¼Æµü¿ù»~°ÝÃD(¥Ø«e¥u·|¦³st¤Îth¨âºØ,¨ä¥Lµ¥¦³»Ý­n¦A¥[)
                  intI = InStr(m_Head(3, iRow), " to")
                  strExc(1) = Left(m_Head(3, iRow), intI - 1)
                  strExc(2) = Mid(m_Head(3, iRow), intI)
                  .Selection.TypeText Text:=strExc(1)
                  If Right(strExc(1), 2) = " 1" Then
                     .Selection.Font.Superscript = True
                     .Selection.TypeText Text:="st"
                     .Selection.Font.Superscript = False
                  End If
                  intI = InStrRev(strExc(2), ",")
                  strExc(1) = Left(strExc(2), intI - 1)
                  strExc(2) = Mid(strExc(2), intI)
                  .Selection.TypeText Text:=strExc(1)
                  If Right(strExc(1), 2) = "31" Then
                     .Selection.Font.Superscript = True
                     .Selection.TypeText Text:="st"
                     .Selection.Font.Superscript = False
                  ElseIf InStr(" 1, 2, 3,21,22,23,31", Right(strExc(1), 2)) = 0 Then
                     .Selection.Font.Superscript = True
                     .Selection.TypeText Text:="th"
                     .Selection.Font.Superscript = False
                  End If
                  .Selection.TypeText Text:=strExc(2)
                  'end 2021/12/8
               Else
                  .Selection.TypeText Text:=m_Head(3, iRow)
               End If
            Else
               .Selection.TypeText Text:=m_Head(3, iRow)
            End If
         End If
         If m_Head(4, iRow) <> "" Then
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:=m_Head(4, iRow)
         Else
            .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
            .Selection.Cells.Merge
         End If
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
      Next

      .Selection.SelectRow
      .Selection.Cells.Merge
      .Selection.InsertRows 3
      .Selection.Collapse Direction:=wdCollapseStart
      
      .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
      .Selection.TypeText Text:="                           "
      .Selection.Font.Size = 14
      .Selection.Font.Bold = True
      .Selection.TypeText Text:=m_Title
      .Selection.Font.Size = strFontSize
      'Modified by Morgan 2024/1/24
      '.Selection.TypeText Text:="      NO. " & m_strDN
      If m_strInvoiceNo <> "" Then
         .Selection.TypeText Text:="      " & m_strInvoiceNo
      Else
         .Selection.TypeText Text:="      NO. " & m_strDN
      End If
      'end 2024/1/24
      .Selection.Font.Bold = False
      
      With .Selection.Cells.Borders(wdBorderBottom)
         .LineStyle = wdLineStyleSingle
         .LineWidth = wdLineWidth050pt
         .ColorIndex = wdAuto
      End With
      
      .Selection.MoveRight Unit:=wdCharacter, Count:=2
      
      '¼ÐÃD
      For iRow = 1 To UBound(m_Subject, 2)
         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
         If m_Subject(1, iRow) <> "" Then
            .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
            If m_Subject(2, iRow) = "" Then
               .Selection.TypeText Text:=m_Subject(1, iRow)
            Else
               .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.TypeText Text:=m_Subject(1, iRow)
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.TypeText Text:=m_Subject(2, iRow)
            End If
            
         Else
            .Selection.Cells(1).SetHeight RowHeight:=16, HeightRule:=wdRowHeightAtLeast
            If m_Subject(3, iRow) <> "" Then
               .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
               'Addded by Morgan 2017/6/21
               If m_Subject(2, iRow) = "Title:" Then
                  .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.2), RulerStyle:=wdAdjustProportional
                  .Selection.Collapse Direction:=wdCollapseStart
                  .Selection.MoveRight Unit:=wdCell, Count:=1
                  .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
                  .Selection.TypeText Text:=m_Subject(2, iRow)
                  .Selection.MoveRight Unit:=wdCell, Count:=1
               Else
               'end 2017/6/21
                  '«e¸m¤å¦r¥i¯à¦³¨â²Õ("In the name of ","Applicant: ")
                  If InStr(strLstStr, "In the name of") = 1 Then
                     .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.8), RulerStyle:=wdAdjustProportional
                  Else
                     .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
                  End If
                  .Selection.Collapse Direction:=wdCollapseStart
                  .Selection.MoveRight Unit:=wdCell, Count:=2
               End If 'Addded by Morgan 2017/6/21
               .Selection.TypeText Text:=m_Subject(3, iRow)
            ElseIf m_Subject(2, iRow) <> "" Then
               strLstStr = m_Subject(2, iRow)
               
               'Added by Morgan 2023/7/26 ¥Ó½Ð¤H¥i¯à·|§é¦æ ex:X11207075
               If Left(strLstStr, 10) = "Applicant:" Or Left(strLstStr, 14) = "In the name of" Then
                  .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=False
                  .Selection.Cells.VerticalAlignment = wdCellAlignVerticalTop
                  .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
                  If InStr(strLstStr, "In the name of") = 1 Then
                     .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.8), RulerStyle:=wdAdjustProportional
                  Else
                     .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
                  End If
                  .Selection.Collapse Direction:=wdCollapseStart
                  .Selection.MoveRight Unit:=wdCell, Count:=1
                  If InStr(strLstStr, "In the name of") = 1 Then
                     .Selection.TypeText Text:=Left(strLstStr, 14)
                     .Selection.MoveRight Unit:=wdCell, Count:=1
                     .Selection.TypeText Text:=Trim(Mid(strLstStr, 15))
                  Else
                     .Selection.TypeText Text:=Left(strLstStr, 10)
                     .Selection.MoveRight Unit:=wdCell, Count:=1
                     .Selection.TypeText Text:=Trim(Mid(strLstStr, 11))
                  End If
                  
               Else
               'end 2023/7/26
               
                  .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
                  .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
                  .Selection.Collapse Direction:=wdCollapseStart
                  .Selection.MoveRight Unit:=wdCell, Count:=1
                  .Selection.TypeText Text:=m_Subject(2, iRow)
                  
               End If 'Added by Morgan 2023/7/26
               
               
            End If
         End If
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
         .Selection.InsertRows 1
      Next
      
      'Added by Morgan 2021/8/23
      'Modified by Morgan 2022/9/16 ¥~°Ó¨ú®ø--³¯ª÷½¬ ,¥~±M¥ý«O¯dµ¥¥N²z¤H«ü¥Ü--Anny
      'If m_A1k27 = "X64826000" Then
      If m_A1k27 = "X64826000" And m_CP01 = "FCP" Then
      'end 2022/9/16
         intI = UBound(m_Subject, 2)
         .Selection.MoveUp Unit:=wdLine, Count:=intI
         .Selection.SelectRow
         iCol = .Selection.Cells.Count
         strExc(1) = .Selection.Cells(iCol).Width '­ì¨ÓÄæ¼e
         strExc(2) = .CentimetersToPoints(5.9)
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.MoveRight Unit:=wdCell, Count:=iCol - 1
         .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
         .Selection.Cells(1).SetWidth ColumnWidth:=Val(strExc(1)) - Val(strExc(2)), RulerStyle:=wdAdjustProportional
         .Selection.Cells(2).SetWidth ColumnWidth:=Val(strExc(2)), RulerStyle:=wdAdjustProportional

         .Selection.MoveDown Unit:=wdLine, Count:=1
         .Selection.SelectRow
         iCol = .Selection.Cells.Count
         strExc(1) = .Selection.Cells(iCol).Width '­ì¨ÓÄæ¼e
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.MoveRight Unit:=wdCell, Count:=iCol - 1
         .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
         .Selection.Cells(1).SetWidth ColumnWidth:=Val(strExc(1)) - Val(strExc(2)), RulerStyle:=wdAdjustProportional
         .Selection.Cells(2).SetWidth ColumnWidth:=Val(strExc(2)), RulerStyle:=wdAdjustProportional

         'Added by Morgan 2022/4/18 FCP¤]­n¼ÐÃD¦C¼Æ¤£¦P
         If m_CP01 = "FCP" Then
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.MoveRight Unit:=wdCell, Count:=1
            .Selection.MoveUp Unit:=wdLine, Count:=intI - 1
            .Selection.MoveDown Unit:=wdLine, Count:=intI - 1, Extend:=wdExtend
            .Selection.Cells.Merge
            
         Else
         'end 2022/4/18
         
            .Selection.MoveDown Unit:=wdLine, Count:=1
            .Selection.SelectRow
            iCol = .Selection.Cells.Count
            strExc(1) = .Selection.Cells(iCol).Width '­ì¨ÓÄæ¼e
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.MoveRight Unit:=wdCell, Count:=iCol - 1
            .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
            .Selection.Cells(1).SetWidth ColumnWidth:=Val(strExc(1)) - Val(strExc(2)), RulerStyle:=wdAdjustProportional
            .Selection.Cells(2).SetWidth ColumnWidth:=Val(strExc(2)), RulerStyle:=wdAdjustProportional
   
            .Selection.MoveDown Unit:=wdLine, Count:=1
            .Selection.SelectRow
            iCol = .Selection.Cells.Count
            strExc(1) = .Selection.Cells(iCol).Width '­ì¨ÓÄæ¼e
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.MoveRight Unit:=wdCell, Count:=iCol - 1
            .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
            .Selection.Cells(1).SetWidth ColumnWidth:=Val(strExc(1)) - Val(strExc(2)), RulerStyle:=wdAdjustProportional
            .Selection.Cells(2).SetWidth ColumnWidth:=Val(strExc(2)), RulerStyle:=wdAdjustProportional
            
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.MoveRight Unit:=wdCell, Count:=1
            .Selection.MoveUp Unit:=wdLine, Count:=intI - 1
            .Selection.MoveDown Unit:=wdLine, Count:=intI - 1, Extend:=wdExtend
            .Selection.Cells.Merge
         End If

         strExc(3) = "CAPEX" & vbCrLf & "Cost Centre: C00038" & vbCrLf & "Capex Code: PCIPTMNEW" & vbCrLf & "Nominal: 102900 - IP - Prof Fees"
         .Selection.Font.Size = 10
         .Selection.Font.Bold = True
         .Selection.Font.ColorIndex = wdRed
         .Selection.TypeText Text:=strExc(3)
         .Selection.MoveDown Unit:=wdLine, Count:=1
         
      End If
      'end 2021/8/23
      
      If iRow = 1 Then 'Added by Morgan 2019/12/5 ¤¤¶¡¤£¥[ªÅ¥Õ¦C¥H­°§C¸õ­¶¾÷²v
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.InsertRows 1
      End If
      
      '©ú²ÓªíÀY
      With .Selection.Cells(1)
         With .Borders(wdBorderLeft)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderRight)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderTop)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderBottom)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
      End With
      
      If m_iPrintCurrType = 1 Or m_iPrintCurrType = 2 Then
         strText = "NTD"
      Else
         strText = m_DNCurr
      End If
      
      .Selection.SelectRow
      .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
      .Selection.Cells.Shading.Texture = wdTexture5Percent
      
      'Added by Morgan 2014/2/18 +m_bSpecialNew1 ®æ¦¡
      If m_bSpecialNew1 Then
         .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=False
         'BASF
         'Modfiied by Morgan 2021/7/1 +Y3326801 BASF Corporation --Anny
         If m_A1k28 = "Y45814010" Or m_A1k28 = "Y33268010" Then
            'Modified by Morgan 2018/3/8 ½Õ¾ã¼e«×
            'Modified by Lydia 2021/09/15 ½Õ¾ã¼e«× (1)10.3=>10 , (3)2.5=>2.8 ; ex.X11013257
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(10), RulerStyle:=wdAdjustProportional
            .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.9), RulerStyle:=wdAdjustProportional
            .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.8), RulerStyle:=wdAdjustProportional
         'Longitude
         ElseIf m_A1k28 = "Y54179000" Then
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(9.5), RulerStyle:=wdAdjustProportional
            .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
            .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustProportional
         
         'Added by Morgan 2019/2/27
         'Advanced Energy Industries, Inc.
         ElseIf m_A1k28 = "Y48904000" Then
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(8.5), RulerStyle:=wdAdjustProportional
            .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(3.2), RulerStyle:=wdAdjustProportional
            .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustProportional
            
         End If
      'end 2014/2/18
      'Added by Morgan 2016/8/5
      ElseIf m_bSpecialNew2 Then
         .Selection.Cells.Split NumRows:=1, NumColumns:=5, MergeBeforeSplit:=False
         'Modified by Morgan 2025/1/6 ªA°È¶O¥~¹ô>=1000.00®É·|§é¦æ,½Õ¾ã¼e«×(»¡©ú-0.2cm) Ex:X11319131
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(10), RulerStyle:=wdAdjustProportional
         .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
         .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(1.6), RulerStyle:=wdAdjustProportional
         .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
         .Selection.InsertRows 1
      'end 2016/8/5
      
      'Added by Morgan 2018/3/22
      ElseIf m_bSpecialNew3 Then
         .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=False
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(4.7), RulerStyle:=wdAdjustProportional
         .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(7), RulerStyle:=wdAdjustProportional
         .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustProportional
      'end 2018/3/22
      
      'Added by Morgan 2021/5/25
      'Modified by Morgan 2025/3/14 +Y56142000--Franny
      ElseIf m_A1k28 = "Y19893030" Or m_A1k28 = "Y56142000" Then
         .Selection.Cells.Split NumRows:=1, NumColumns:=5, MergeBeforeSplit:=False
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(8.8), RulerStyle:=wdAdjustProportional
         .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
         .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
         .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
      'end 2021/5/25
      Else
         .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=False
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(11.7), RulerStyle:=wdAdjustProportional
         .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustProportional
      End If
      
      .Selection.Cells(1).SetHeight RowHeight:=12, HeightRule:=wdRowHeightAtLeast
      .Selection.Collapse Direction:=wdCollapseStart
      If m_bSpecialNew2 Then
         .Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
         .Selection.Cells.Merge
      End If
      
      'Added by Morgan 2018/3/22
      If m_bSpecialNew3 Then
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.TypeText Text:="Date the work was performed"
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
      End If
      
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText Text:="Item"
      
      'Added by Morgan 2014/2/18 +m_bSpecialNew1 ®æ¦¡
      If m_bSpecialNew1 Then
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         'BASF
         'Modfiied by Morgan 2021/7/1 +Y3326801 BASF Corporation --Anny
         If m_A1k28 = "Y45814010" Or m_A1k28 = "Y33268010" Then
            .Selection.TypeText Text:="Position"
         'Longitude
         ElseIf m_A1k28 = "Y54179000" Then
            .Selection.TypeText Text:="Billing Code"
         'Advanced Energy Industries, Inc.
         ElseIf m_A1k28 = "Y48904000" Then
            .Selection.TypeText Text:="Activity"
         End If
      End If
      'end 2014/2/18
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      
      'Added by Morgan 2016/8/5
      If m_bSpecialNew2 Then
         .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
         .Selection.Cells.Merge
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.TypeText Text:="Official fee"
         .Selection.MoveRight Unit:=wdCell, Count:=1
         .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
         .Selection.Cells.Merge
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.TypeText Text:="Attorney fee"
         .Selection.MoveDown Unit:=wdLine, Count:=1
         .Selection.MoveLeft Unit:=wdCell, Count:=3
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.TypeText Text:="NTD"
         .Selection.MoveRight Unit:=wdCell, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.TypeText Text:=strText
         .Selection.MoveRight Unit:=wdCell, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.TypeText Text:="NTD"
         .Selection.MoveRight Unit:=wdCell, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.TypeText Text:=strText
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.InsertRows 1
         .Selection.Cells.Shading.Texture = wdTextureNone
         .Selection.Collapse Direction:=wdCollapseStart
      'Added by Morgan 2021/5/25
      'Modified by Morgan 2025/3/14 +Y56142000--Franny
      ElseIf m_A1k28 = "Y19893030" Or m_A1k28 = "Y56142000" Then
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.TypeText Text:="Hourly Rate/" & vbCrLf & "Flat Fee"
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.TypeText Text:="Time" & vbCrLf & "(Hour)"
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.TypeText Text:="Official fee" & vbCrLf & strText
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.TypeText Text:="Attorney fee" & vbCrLf & strText
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.InsertRows 1
         .Selection.Cells.Shading.Texture = wdTextureNone
         .Selection.Collapse Direction:=wdCollapseStart
      'end 2021/5/25
      
      Else
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.TypeText Text:="Official fee"
         .Selection.TypeParagraph
         .Selection.TypeText Text:=strText
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.TypeText Text:="Attorney fee"
         .Selection.TypeParagraph
         .Selection.TypeText Text:=strText
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.InsertRows 1
         .Selection.Cells.Shading.Texture = wdTextureNone
         .Selection.Collapse Direction:=wdCollapseStart
         
      End If
      
      strLstNo = ""
      
      'Added by Morgan 2019/3/28
      'Y48904000 Advanced Energy Industries, Inc.
      '­«·s¾ã²z:¬Û¦PªºÃþ§O(Activity)­n¬Û¾F¨Ã¤p­p
      If m_A1k28 = "Y48904000" Then
         For iRow = 1 To UBound(m_Item) - 1
            For iCol = iRow + 1 To UBound(m_Item)
               If m_Item(iCol).ICode = m_Item(iRow).ICode Then
                  If iCol > iRow + 1 Then
                     ItemCache = m_Item(iRow + 1)
                     m_Item(iRow + 1) = m_Item(iCol)
                     m_Item(iCol) = ItemCache
                  End If
                  Exit For
               End If
            Next iCol
         Next iRow
      End If
      'end 2019/3/28
      
      '©ú²Ó
      For iRow = 1 To UBound(m_Item)
         strText = ""
         
         If Right(m_Item(iRow).iNo, 2) = "99" Then
            dblOffFeeSub = dblOffFeeSub + Val(Format(m_Item(iRow).iAmt))
            dblNtOffFeeSub = dblNtOffFeeSub + Val(Format(m_Item(iRow).INtAmt))
         Else
            dblAttFeeSub = dblAttFeeSub + Val(Format(m_Item(iRow).iAmt))
            dblNtAttFeeSub = dblNtAttFeeSub + Val(Format(m_Item(iRow).INtAmt))
         End If
         '­Y¬°«e¶µªº³W¶O«h¦X¨Ö
         If m_Item(iRow).iNo = strLstNo & "99" Then
            'Added by Morgan 2017/7/7
            'BASF
            'Modfiied by Morgan 2021/7/1 +Y3326801 BASF Corporation --Anny
            If m_A1k28 = "Y45814010" Or m_A1k28 = "Y33268010" Then
                'Modified by Morgan 2018/3/21 +­×¥¿¦³§é¦©»¡©ú·|·í°ÝÃD Ex:X10704103
                If m_Item(iRow - 1).IDescTail = "" Then
                    .Selection.InsertRows 1
                    .Selection.Collapse Direction:=wdCollapseStart
                    .Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
                    .Selection.Cells.Merge
                    .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
                    .Selection.MoveRight Unit:=wdCell, Count:=1
                    .Selection.MoveDown Unit:=wdLine, Count:=1
                Else
                    .Selection.MoveLeft Unit:=wdCharacter, Count:=1
                    .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
                    .Selection.MoveLeft Unit:=wdCell, Count:=1
                    .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
                    .Selection.MoveLeft Unit:=wdCell, Count:=1
                    .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
                End If
               
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
               .Selection.Font.Bold = False
               .Selection.TypeText Text:=m_Item(iRow).ICode
               
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
               .Selection.Font.Bold = False
               .Selection.TypeText Text:=m_Item(iRow).iAmt & "(1PU)"
               
               .Selection.MoveRight Unit:=wdCharacter, Count:=2
               
            ElseIf m_Item(iRow).IDescTail <> "" Or m_Item(iRow).IDescX <> "" Then
                              
               If m_Item(iRow).IDescX <> "" Then
                  strText = m_Item(iRow).IDescX
               Else
                  strText = "Official fee"
               End If
               
               If m_Item(iRow).IDescTail <> "" Then
                  strText = strText & ":" & m_Item(iRow).IDescTail
               End If
               
               .Selection.InsertRows 1
               .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleNone
               'Added by Morgan 2017/12/14
               If m_bSpecialNew2 Then
                  .Selection.Cells.Split NumRows:=1, NumColumns:=6, MergeBeforeSplit:=True
                  .Selection.Cells(1).Borders(wdBorderRight).LineStyle = wdLineStyleNone
                  .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.6), RulerStyle:=wdAdjustProportional
                  'Modified by Morgan 2025/1/6 ªA°È¶O¥~¹ô>=1000.00®É·|§é¦æ,½Õ¾ã¼e«×(»¡©ú-0.2cm) Ex:X11319131
                  .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(9.4), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(1.6), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(5).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
               'Added by Morgan 2018/12/3
               ElseIf m_bSpecialNew3 Then
                  .Selection.Cells.Split NumRows:=1, NumColumns:=5, MergeBeforeSplit:=True
                  .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(4.7), RulerStyle:=wdAdjustFirstColumn
                  .Selection.Cells(2).Borders(wdBorderRight).LineStyle = wdLineStyleNone
                  .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(0.6), RulerStyle:=wdAdjustFirstColumn
                  .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(6.4), RulerStyle:=wdAdjustFirstColumn
                  .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustFirstColumn
               'end 2018/12/3
               Else
               'end 2017/12/14
                  .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=True
                  .Selection.Cells(1).Borders(wdBorderRight).LineStyle = wdLineStyleNone
                  .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.6), RulerStyle:=wdAdjustFirstColumn
                  .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(11.1), RulerStyle:=wdAdjustFirstColumn
                  .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustFirstColumn
               End If
               .Selection.Collapse Direction:=wdCollapseStart
               If m_bSpecialNew3 Then .Selection.MoveRight Unit:=wdCharacter, Count:=1 'Added by Morgan 2018/12/3
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
               .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
               .Selection.TypeText Text:=strText
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
               'Added by Morgan 2017/12/14
               If m_bSpecialNew2 Then
                  .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                  .Selection.TypeText Text:=m_Item(iRow).INtAmt
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
               End If
               'end 2017/12/14
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
               .Selection.TypeText Text:=m_Item(iRow).iAmt
               'Added by Morgan 2017/12/14
               If m_bSpecialNew2 Then
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
               End If
               'end 2017/12/14
               .Selection.MoveRight Unit:=wdCharacter, Count:=2
               
            Else
               'Added by Morgan 2016/8/5
               If m_bSpecialNew2 Then
                  .Selection.MoveLeft Unit:=wdCell, Count:=4
               'end 2016/8/5
               Else
                  .Selection.MoveLeft Unit:=wdCell, Count:=2
               End If
               
               If m_Item(iRow - 1).IDescTail <> "" Then
                  .Selection.MoveUp Unit:=wdLine, Count:=1
               End If
               
               'Added by Morgan 2016/8/5
               If m_bSpecialNew2 Then
                  .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                  .Selection.TypeText Text:=m_Item(iRow).INtAmt
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
               End If
               'end 2016/8/5
               
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
               'Modified by Morgan 2018/4/11
               If m_bSpecialNew4 Then
                  .Selection.TypeText Text:=m_Item(iRow).iCur & " " & m_Item(iRow).iAmt
               Else
                  .Selection.TypeText Text:=m_Item(iRow).iAmt
                  'Added by Morgan 2025/10/23 Y52019000¯Â¥~¹ô­n¥[¦L¥x¹ôª÷ÃB -- Kahn
                  If m_A1k28 = "Y52019000" And m_iPrintCurrType = "3" And m_CP01 = "FCP" Then
                     .Selection.TypeParagraph
                     .Selection.TypeText Text:="(NTD" & Format(m_Item(iRow).INtAmt, "#,##0") & ")"
                  End If
                  'end 2025/10/23
               End If
               'end 2018/4/11
               
               'Added by Morgan 2016/8/5
               If m_bSpecialNew2 Then
                  .Selection.MoveRight Unit:=wdCell, Count:=2
               'end 2016/8/5
               Else
                  .Selection.MoveRight Unit:=wdCell, Count:=1
               End If
               
               If m_Item(iRow - 1).IDescTail <> "" Then
                  .Selection.MoveDown Unit:=wdLine, Count:=1
               ElseIf .Selection.Text <> "" Then
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
               End If
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
            End If
         
         Else
            .Selection.InsertRows 1
            '¶µ¥Ø¥Îµê½u¤À¹j
            If iRow > 1 Then
               .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleDashLargeGap
               .Selection.Cells.Borders(wdBorderTop).LineWidth = wdLineWidth050pt
               'Added by Morgan 2019/3/29
               If m_A1k28 = "Y48904000" Then
                  If m_Item(iRow).ICode <> m_Item(iRow - 1).ICode Then
                     .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
                  End If
               End If
               'end 2019/3/29
            End If
            
            'Added by Morgan 2014/2/18
            If m_bSpecialNew1 Then
               .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=True
               'BASF
               'Modfiied by Morgan 2021/7/1 +Y3326801 BASF Corporation --Anny
               If m_A1k28 = "Y45814010" Or m_A1k28 = "Y33268010" Then
                  'Modified by Morgan 2018/3/8 ½Õ¾ã¼e«×
                  'Modified by Lydia 2021/09/15 ½Õ¾ã¼e«× (1)10.3=>10 , (3)2.5=>2.8 ; ex.X11013257
                  .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(10), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.9), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.8), RulerStyle:=wdAdjustProportional
               'Longitude
               ElseIf m_A1k28 = "Y54179000" Then
                  .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(9.5), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustProportional
                  
               'Added by Morgan 2019/2/27
               'Advanced Energy Industries, Inc.
               ElseIf m_A1k28 = "Y48904000" Then
                  .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(8.5), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(3.2), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustProportional
               
               End If
               
            'end 2014/2/18
            'Added by Morgan 2016/8/5
            ElseIf m_bSpecialNew2 Then
               .Selection.Cells.Split NumRows:=1, NumColumns:=5, MergeBeforeSplit:=True
               'Modified by Morgan 2025/1/6 ªA°È¶O¥~¹ô>=1000.00®É·|§é¦æ,½Õ¾ã¼e«×(»¡©ú-0.2cm) Ex:X11319131
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(10), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
               .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(1.6), RulerStyle:=wdAdjustProportional
               .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
            'end 2016/8/5
            
            'Added by Morgan 2018/3/22
            ElseIf m_bSpecialNew3 Then
               .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=True
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(4.7), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(7), RulerStyle:=wdAdjustProportional
               .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustProportional
            'end 2018/3/22
            
            'Added by Morgan 2021/5/25
            'Modified by Morgan 2025/3/14 +Y56142000--Franny
            ElseIf m_A1k28 = "Y19893030" Or m_A1k28 = "Y56142000" Then
               .Selection.Cells.Split NumRows:=1, NumColumns:=5, MergeBeforeSplit:=True
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(8.8), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
               .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
               .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
            'end 2021/5/25
      
            Else
               .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(11.7), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustProportional
            End If
            
            .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
            .Selection.Collapse Direction:=wdCollapseStart
            
            'Added by Morgan 2018/3/22
            If m_bSpecialNew3 Then
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
               .Selection.TypeText Text:=m_Item(iRow).IDate
               .Selection.MoveRight Unit:=wdCell, Count:=1
            End If
            'end 2018/3/22
            
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            bolNewRow = False
            
            '³Ì«á¤@¶µ
            If iRow = UBound(m_Item) Then
               .Selection.TypeText Text:=m_Item(iRow).IDesc
            '«á­±±µ¬Û¦P½Ð´Ú¶µ¥Øªº³W¶O
            ElseIf m_Item(iRow + 1).iNo = m_Item(iRow).iNo & "99" Then
               .Selection.TypeText Text:=m_Item(iRow).IDescHead
               bolNewRow = True
            '¨ä¥L
            Else
               .Selection.TypeText Text:=m_Item(iRow).IDesc
            End If
            
            'Added by Morgan 2014/2/18
            If m_bSpecialNew1 Then
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
               .Selection.Font.Bold = False
               'Added by Morgan 2019/3/29 ¬Û¦P¤ÀÃþ¤£¦L(­n¦X¨ÖÀx¦s®æ)
               If m_A1k28 = "Y48904000" And iRow > 1 Then
                  If m_Item(iRow).ICode <> m_Item(iRow - 1).ICode Then
                     .Selection.TypeText Text:=m_Item(iRow).ICode
                  End If
               Else
               'end 2019/3/29
               
                  .Selection.TypeText Text:=m_Item(iRow).ICode
                  
               End If 'Added by Morgan 2019/3/29
            End If
            'end 2014/2/18
            
            If Right(m_Item(iRow).iNo, 2) = "99" Then
               'Modified by Morgan 2016/8/5
               '.Selection.MoveRight Unit:=wdCell, Count:=1
               If m_bSpecialNew2 Then
                  .Selection.MoveRight Unit:=wdCell, Count:=1
                  .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                  .Selection.Font.Bold = False
                  .Selection.TypeText Text:=m_Item(iRow).INtAmt
                  .Selection.MoveRight Unit:=wdCell, Count:=1
               'Added by Morgan 2021/5/25
               'Modified by Morgan 2025/3/14 +Y56142000--Franny
               ElseIf m_A1k28 = "Y19893030" Or m_A1k28 = "Y56142000" Then
                  .Selection.MoveRight Unit:=wdCell, Count:=1
                  .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                  .Selection.Font.Bold = False
                  .Selection.TypeText "Flat Fee"
                  .Selection.MoveRight Unit:=wdCell, Count:=2
               'end 2021/5/25
               Else
                  .Selection.MoveRight Unit:=wdCell, Count:=1
               End If
               'end 2016/8/5
            Else
               'Modified by Morgan 2016/8/5
               '.Selection.MoveRight Unit:=wdCell, Count:=2
               If m_bSpecialNew2 Then
                  .Selection.MoveRight Unit:=wdCell, Count:=3
                  .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                  .Selection.Font.Bold = False
                  .Selection.TypeText Text:=m_Item(iRow).INtAmt
                  .Selection.MoveRight Unit:=wdCell, Count:=1
               'Added by Morgan 2021/5/25
               'Modified by Morgan 2025/3/14 +Y56142000--Franny
               ElseIf m_A1k28 = "Y19893030" Or m_A1k28 = "Y56142000" Then
                  .Selection.MoveRight Unit:=wdCell, Count:=1
                  .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                  .Selection.Font.Bold = False
                  If InStr("'1202','205','1002','107','203','431','422','204','903'", "'" & m_Item(iRow).iNo & "'") > 0 Then
                     .Selection.TypeText "Hourly Rate"
                     .Selection.MoveRight Unit:=wdCell, Count:=1
                     .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                     'Modified by Morgan 2025/3/14
                     '.Selection.TypeText Round(Format(m_Item(iRow).iAmt) / 4500, 2)
                     If Val(m_LD16) > 0 Then
                        .Selection.TypeText Round(Format(m_Item(iRow).iAmt) / Val(m_LD16), 2)
                     ElseIf m_A1k28 = "Y19893030" Then
                        .Selection.TypeText Round(Format(m_Item(iRow).iAmt) / 4500, 2)
                     ElseIf m_A1k28 = "Y56142000" Then
                        .Selection.TypeText Round(Format(m_Item(iRow).iAmt) / 149, 2)
                     End If
                     'end 2025/3/14
                     .Selection.MoveRight Unit:=wdCell, Count:=2
                  Else
                     .Selection.TypeText "Flat Fee"
                     .Selection.MoveRight Unit:=wdCell, Count:=3
                  End If
               Else
                  .Selection.MoveRight Unit:=wdCell, Count:=2
               End If
               'end 2016/8/5
            End If
            
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
            .Selection.Font.Bold = False
            
            'Added by Morgan 2017/7/7 BASF
            'Modfiied by Morgan 2021/7/1 +Y3326801 BASF Corporation --Anny
            If m_A1k28 = "Y45814010" Or m_A1k28 = "Y33268010" Then
               .Selection.TypeText Text:=m_Item(iRow).iAmt & "(1PU)"
            'Added by Morgan 2018/4/11
            ElseIf m_bSpecialNew4 Then
                  .Selection.TypeText Text:=m_Item(iRow).iCur & " " & m_Item(iRow).iAmt
            'end 2018/4/11
            Else
            'end 2017/7/7
               .Selection.TypeText Text:=m_Item(iRow).iAmt
               'Added by Morgan 2025/10/23 Y52019000¯Â¥~¹ô­n¥[¦L¥x¹ôª÷ÃB -- Kahn
               If m_A1k28 = "Y52019000" And m_iPrintCurrType = "3" And m_CP01 = "FCP" Then
                  .Selection.TypeParagraph
                  .Selection.TypeText Text:="(NTD" & Format(m_Item(iRow).INtAmt, "#,##0") & ")"
               End If
               'end 2025/10/23
            End If 'Added by Morgan 2017/7/7
            
            If Right(m_Item(iRow).iNo, 2) = "99" Then
               'Modified by Morgan 2016/8/5
               '.Selection.MoveRight Unit:=wdCharacter, Count:=2
               If m_bSpecialNew2 Then
                  .Selection.MoveRight Unit:=wdCharacter, Count:=3
               Else
                  .Selection.MoveRight Unit:=wdCharacter, Count:=2
               End If
               'end 2016/8/5
            Else
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
            End If
            
             If bolNewRow And m_Item(iRow).IDescTail <> "" Then
             
               .Selection.InsertRows 1
               .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleNone
               
               'Added by Morgan 2014/3/10
               If m_bSpecialNew1 Then
                  .Selection.Cells(1).Split NumRows:=1, NumColumns:=2
                  .Selection.Cells(1).Borders(wdBorderRight).LineStyle = wdLineStyleNone
                  .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.6), RulerStyle:=wdAdjustFirstColumn
               'end 2014/3/10
               
               'Added by Morgan 2016/8/5
               ElseIf m_bSpecialNew2 Then
                  .Selection.Cells.Split NumRows:=1, NumColumns:=6, MergeBeforeSplit:=True
                  .Selection.Cells(1).Borders(wdBorderRight).LineStyle = wdLineStyleNone
                  .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.6), RulerStyle:=wdAdjustFirstColumn
                  'Modified by Morgan 2025/1/6 ªA°È¶O¥~¹ô>=1000.00®É·|§é¦æ,½Õ¾ã¼e«×(»¡©ú-0.2cm) Ex:X11319131
                  .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(9.4), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(1.6), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(5).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
               'end 2016/8/5
               
               'Added by Morgan 2023/8/29
               ElseIf m_bSpecialNew3 Then
                  .Selection.Cells.Split NumRows:=1, NumColumns:=5, MergeBeforeSplit:=True
                  .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(4.7), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(0.6), RulerStyle:=wdAdjustFirstColumn
                  .Selection.Cells(2).Borders(wdBorderRight).LineStyle = wdLineStyleNone
                  .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(6.4), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustProportional
               
               'Added by Morgan 2021/6/4
               'Modified by Morgan 2025/3/14 +Y56142000--Franny
               ElseIf m_A1k28 = "Y19893030" Or m_A1k28 = "Y56142000" Then
                  .Selection.Cells.Split NumRows:=1, NumColumns:=6, MergeBeforeSplit:=True
                  .Selection.Cells(1).Borders(wdBorderRight).LineStyle = wdLineStyleNone
                  .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.6), RulerStyle:=wdAdjustFirstColumn
                  .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(8.2), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(5).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
               'end 2021/5/25
            
               Else
                  .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=True
                  .Selection.Cells(1).Borders(wdBorderRight).LineStyle = wdLineStyleNone
                  .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.6), RulerStyle:=wdAdjustFirstColumn
                  .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(11.1), RulerStyle:=wdAdjustFirstColumn
                  .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustFirstColumn
               End If
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
               'Added by Morgan 2023/8/29
               If m_bSpecialNew3 Then
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
               End If
               'end 2023/8/29
               strText = "Attorney fee:" & m_Item(iRow).IDescTail
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
               .Selection.TypeText Text:=strText
               
               'Added by Morgan 2014/3/10
               If m_bSpecialNew1 Then
                  .Selection.MoveRight Unit:=wdCell, Count:=3
               'end 2014/3/10
               'Added by Morgan 2016/8/5
               ElseIf m_bSpecialNew2 Then
                  .Selection.MoveRight Unit:=wdCell, Count:=4
               'end 2016/8/5
               'Added by Morgan 2023/8/29
               ElseIf m_bSpecialNew3 Then
                  .Selection.MoveRight Unit:=wdCell, Count:=3
               'Added by Morgan 2021/6/4
               'Modified by Morgan 2025/3/14 +Y56142000--Franny
               ElseIf m_A1k28 = "Y19893030" Or m_A1k28 = "Y56142000" Then
                  .Selection.MoveRight Unit:=wdCell, Count:=4
               Else
                  .Selection.MoveRight Unit:=wdCell, Count:=2
               End If
               
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
            End If
               
         End If
         strLstNo = m_Item(iRow).iNo
         
         'Added by Morgan 2019/3/28
         If m_A1k28 = "Y48904000" Then
            If Right(m_Item(iRow).iNo, 2) = "99" Then
               dblOffFeeActSub = dblOffFeeActSub + Val(Format(m_Item(iRow).iAmt))
            Else
               dblAttFeeActSub = dblAttFeeActSub + Val(Format(m_Item(iRow).iAmt))
            End If
            
            '¦X¨Ö¤ÀÃþÀx¦s®æ
            bolMergeRow = False
            If iRow = UBound(m_Item) Then
               If iRow > 1 And iMergeRows > 0 Then
                  bolMergeRow = True
               End If
            ElseIf m_Item(iRow + 1).ICode <> m_Item(iRow).ICode Then
               bolMergeRow = True
               
            ElseIf m_Item(iRow + 1).iNo <> m_Item(iRow).iNo & "99" Then
               iMergeRows = iMergeRows + 1
            End If
            
            If bolMergeRow And iMergeRows > 0 Then
               .Selection.MoveLeft Unit:=wdCell, Count:=3
               .Selection.MoveUp Unit:=wdLine, Count:=iMergeRows
               .Selection.MoveDown Unit:=wdLine, Count:=iMergeRows, Extend:=wdExtend
               .Selection.Cells.Merge
               .Selection.MoveDown Unit:=wdLine, Count:=1
               .Selection.SelectRow
               bolPrintActSub = True
'
'               If bolPrintActSub Then
'                  '¦C¦L¤ÀÃþ¤p­p(·|¼vÅT½Ð´Úª÷ÃB»Ý¦A½T»{)
'               End If

               iMergeRows = 0
               dblOffFeeActSub = 0
               dblAttFeeActSub = 0
            End If
         End If
         'end 2019/3/28
      Next
      
      '¤p­p
      .Selection.InsertRows 1
      .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleDouble
      .Selection.Cells.Borders(wdBorderTop).LineWidth = wdLineWidth050pt
      .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
      .Selection.Cells(1).SetHeight RowHeight:=30, HeightRule:=wdRowHeightAtLeast
      If m_bSpecialNew2 Then  'Added by Morgan 2016/8/5
         If .Selection.Cells.Count = 6 Then
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
            .Selection.Cells.Merge
         End If
         
      Else
         .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
         'Modfiied by Morgan 2021/7/1 +Y3326801 BASF Corporation --Anny
         If m_A1k28 = "Y45814010" Or m_A1k28 = "Y33268010" Then
             'Modified by Lydia 2021/09/15 ½Õ¾ã¼e«× (1)12.2=>11.9 , (2)2.5=>2.8 ; ex.X11013257
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(11.9), RulerStyle:=wdAdjustProportional
            .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.8), RulerStyle:=wdAdjustProportional
         'Added by Morgan 2021/5/25
         'Modified by Morgan 2025/3/14 +Y56142000--Franny
         ElseIf m_A1k28 = "Y19893030" Or m_A1k28 = "Y56142000" Then
            .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=True
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(8.8), RulerStyle:=wdAdjustProportional
            .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(4), RulerStyle:=wdAdjustProportional
            .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
         'end 2021/5/25
         Else
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(11.7), RulerStyle:=wdAdjustProportional
            .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustProportional
         End If
         
      End If 'Added by Morgan 2016/8/5
      
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText Text:="Subtotal"
      'Added by Morgan 2016/8/5
      If m_bSpecialNew2 Then
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         If dblNtOffFeeSub > 0 Then
            strText = Format(dblNtOffFeeSub, "#,##0.00")
            .Selection.TypeText Text:=strText
         End If
      End If
      'end 2016/8/5
      
      'Added by Morgan 2021/5/25
      'Modified by Morgan 2025/3/14 +Y56142000--Franny
      If m_A1k28 = "Y19893030" Or m_A1k28 = "Y56142000" Then
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
      End If
      'end 2021/5/25
      
      '³W¶O
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      If dblOffFeeSub > 0 Then
         strText = Format(dblOffFeeSub, "#,##0.00")
         dblA1K40 = dblOffFeeSub
         'Modified by Morgan 2018/4/11
         If m_bSpecialNew4 Then
            .Selection.TypeText Text:=m_Sum(2, 1) & " " & strText
         Else
            .Selection.TypeText Text:=strText
            'Added by Morgan 2025/10/23 Y52019000¯Â¥~¹ô­n¥[¦L¥x¹ôª÷ÃB -- Kahn
            If m_A1k28 = "Y52019000" And m_iPrintCurrType = "3" And m_CP01 = "FCP" Then
               .Selection.TypeParagraph
               .Selection.TypeText Text:="(NTD" & Format(dblNtOffFeeSub, "#,##0") & ")"
            End If
            'end 2025/10/23
         End If
         'end 2018/4/11
         If m_iPrintCurrType = 2 Or m_iPrintCurrType = 4 Then
            .Selection.TypeParagraph
            If m_iPrintCurrType = 2 Then
               dblA1K40 = Trunc(dblOffFeeSub / Val(m_DNRate))
               strText = "(" & m_DNCurr & Format(dblA1K40, "#,##0.00") & ")"
            Else
'               dblA1K40 = Trunc(dblOffFeeSub * Val(m_DUsdRate))
               strText = "(USD" & Format(Trunc(dblOffFeeSub * Val(m_DUsdRate)), "#,##0.00") & ")"
            End If
            .Selection.TypeText Text:=strText
         End If
         
         'Add By Sindy 2021/10/26 °O¿ý ½Ð´Ú³W¶O¥~¹ôª÷ÃB
         strSql = "Update ACC1K0 Set A1K40=" & dblA1K40 & " Where A1K01='" & m_strDN & "'"
         adoTaie.Execute strSql, intI
         '2021/10/26 END
      Else
         'Add By Sindy 2021/12/6 °O¿ý ½Ð´Ú³W¶O¥~¹ôª÷ÃB
         strSql = "Update ACC1K0 Set A1K40=" & dblOffFeeSub & " Where A1K01='" & m_strDN & "'"
         adoTaie.Execute strSql, intI
         '2021/12/6 END
         
         If m_iPrintCurrType = 2 Or m_iPrintCurrType = 4 Then
            .Selection.TypeParagraph
         End If
      End If
      
      'Added by Morgan 2025/4/10 --Franny
      If m_A1k28 = "Y55295000" Then
         .Selection.TypeParagraph
         .Selection.TypeText Text:="610320"
      End If
      'end 2025/4/10
      
      'Added by Morgan 2016/8/5
      If m_bSpecialNew2 Then
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         If dblNtAttFeeSub > 0 Then
            strText = Format(dblNtAttFeeSub, "#,##0.00")
            .Selection.TypeText Text:=strText
         End If
      End If
      'end 2016/8/5
      
      'ªA°È¶O
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      If dblAttFeeSub > 0 Then
         strText = Format(dblAttFeeSub, "#,##0.00")
         dblA1K39 = dblAttFeeSub
         'Modified by Morgan 2018/4/11
         If m_bSpecialNew4 Then
            .Selection.TypeText Text:=m_Sum(2, 1) & " " & strText
         Else
            .Selection.TypeText Text:=strText
            'Added by Morgan 2025/10/23 Y52019000¯Â¥~¹ô­n¥[¦L¥x¹ôª÷ÃB -- Kahn
            If m_A1k28 = "Y52019000" And m_iPrintCurrType = "3" And m_CP01 = "FCP" Then
               .Selection.TypeParagraph
               .Selection.TypeText Text:="(NTD" & Format(dblNtAttFeeSub, "#,##0") & ")"
            End If
            'end 2025/10/23
         End If
         'end 2018/4/11
         If m_iPrintCurrType = 2 Or m_iPrintCurrType = 4 Then
            .Selection.TypeParagraph
            If m_iPrintCurrType = 2 Then
               dblA1K39 = Trunc(dblAttFeeSub / Val(m_DNRate))
               strText = "(" & m_DNCurr & Format(dblA1K39, "#,##0.00") & ")"
            Else
'               dblA1K39 = Trunc(dblAttFeeSub * Val(m_DUsdRate))
               strText = "(USD" & Format(Trunc(dblAttFeeSub * Val(m_DUsdRate)), "#,##0.00") & ")"
            End If
            .Selection.TypeText Text:=strText
         End If
         
         'Add By Sindy 2021/10/26 °O¿ý ½Ð´ÚªA°È¶O¥~¹ôª÷ÃB
         strSql = "Update ACC1K0 Set A1K39=" & dblA1K39 & " Where A1K01='" & m_strDN & "'"
         adoTaie.Execute strSql, intI
         '2021/10/26 END
      Else
         'Add By Sindy 2021/12/6 °O¿ý ½Ð´ÚªA°È¶O¥~¹ôª÷ÃB
         strSql = "Update ACC1K0 Set A1K39=" & dblAttFeeSub & " Where A1K01='" & m_strDN & "'"
         adoTaie.Execute strSql, intI
         '2021/12/6 END
         
         If m_iPrintCurrType = 2 Or m_iPrintCurrType = 4 Then
            .Selection.TypeParagraph
         End If
      End If
      
      'Added by Morgan 2025/4/10 --Franny
      If m_A1k28 = "Y55295000" Then
         .Selection.TypeParagraph
         .Selection.TypeText Text:="636420"
      End If
      'end 2025/4/10
      
      .Selection.MoveRight Unit:=wdCharacter, Count:=2
      
      '¦X­p
      .Selection.SelectRow
      .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
      
      'Added by Morgan 2016/8/5
      If m_bSpecialNew2 Then
         'Modified by Morgan 2025/1/6 ªA°È¶O¥~¹ô>=1000.00®É·|§é¦æ,½Õ¾ã¼e«×(»¡©ú-0.2cm) Ex:X11319131
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(10), RulerStyle:=wdAdjustProportional
      'end 2016/8/5
      Else
         'Modfiied by Morgan 2021/7/1 +Y3326801 BASF Corporation --Anny
         If m_A1k28 = "Y45814010" Or m_A1k28 = "Y33268010" Then
            'Modified by Lydia 2021/09/15 ½Õ¾ã¼e«× (1)12.2=>11.9 ; ex.X11013257
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(11.9), RulerStyle:=wdAdjustProportional
         'Added by Morgan 2021/5/25
         'Modified by Morgan 2025/3/14 +Y56142000--Franny
         ElseIf m_A1k28 = "Y19893030" Or m_A1k28 = "Y56142000" Then
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(8.8), RulerStyle:=wdAdjustProportional
         'end 2021/5/25
         Else
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(11.7), RulerStyle:=wdAdjustProportional
         End If
      End If
      
      .Selection.Cells(1).SetHeight RowHeight:=30, HeightRule:=wdRowHeightAtLeast
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.Font.Bold = True
      .Selection.Font.Size = 12
      .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText Text:="Total"
      .Selection.Font.Size = strFontSize
      
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Font.Bold = True
      'Added by Morgan 2016/8/5
      If m_bSpecialNew2 Then
         strText = "NTD" & Format(dblNtOffFeeSub + dblNtAttFeeSub, FDollar)
         .Selection.TypeText Text:=strText
         .Selection.TypeParagraph
      End If
      'end 2016/8/5
      
      'Modified by Morgan 2018/4/11
      If m_bSpecialNew4 Then
         strText = m_Sum(2, 1) & " " & m_Sum(3, 1)
      Else
         strText = m_Sum(2, 1) & m_Sum(3, 1)
      End If
      'end 2018/4/11
      .Selection.TypeText Text:=strText
      
      If m_iPrintCurrType = 2 Or m_iPrintCurrType = 4 Then
         .Selection.TypeParagraph
         If m_iPrintCurrType = 2 Then
            strText = m_DNCurr & Format(Trunc(dblOffFeeSub / Val(m_DNRate)) + Trunc(dblAttFeeSub / Val(m_DNRate)), "#,##0.00")
         Else
            strText = "USD" & Format(Trunc(dblOffFeeSub * Val(m_DUsdRate)) + Trunc(dblAttFeeSub * Val(m_DUsdRate)), "#,##0.00")
            
            UpdateA1K38 m_strDN, Trunc(dblOffFeeSub * Val(m_DUsdRate)) + Trunc(dblAttFeeSub * Val(m_DUsdRate)) 'Added by Morgan 2021/1/14 §ó·s½Ð´Ú³æ¬üª÷Á`ÃB
         End If
         .Selection.TypeText Text:=strText
      
      'Added by Morgan 2025/10/23 Y52019000¯Â¥~¹ô­n¥[¦L¥x¹ôª÷ÃB -- Kahn
      ElseIf m_A1k28 = "Y52019000" And m_iPrintCurrType = "3" And m_CP01 = "FCP" Then
         .Selection.TypeParagraph
         .Selection.TypeText Text:="(NTD" & Format(dblNtOffFeeSub + dblNtAttFeeSub, "#,##0") & ")"
      'end 2025/10/23
      End If
      
      .Selection.Font.Bold = False
      .Selection.MoveRight Unit:=wdCharacter, Count:=3
      .Selection.InsertRows 1
      .Selection.Cells.Merge
      .Selection.Cells(1).Borders(wdBorderTop).LineStyle = wdLineStyleNone 'Added by Morgan 2014/8/15
      
      '¥u¦³®æ¦¡2(¥x¹ô+¥~¹ô¦X­p)¤~±a
      If m_iPrintCurrType = 2 Then
         .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
         .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.6), RulerStyle:=wdAdjustFirstColumn
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         strText = "Remarks: "
         strText1 = "The total sum in " & m_DNCurr & " is calculated based on the subtotal official fees in " & m_DNCurr
         strText1 = strText1 & " plus the subtotal attorney fees in " & m_DNCurr & " rather than directly divided from NTD"
         strText1 = strText1 & " based on the current rate. Minor differences in rounding up or rounding down to the nearest complete"
         strText1 = strText1 & " dollar has caused the one " & m_DNCurr & " variation."
         'Added by Morgan 2021/7/27
         If InStr("Y5222000,Y52220B1,Y52220B2", Left(m_A1k28, 8)) > 0 Then
            .Selection.TypeText Text:=strText
            .Selection.TypeParagraph
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.InsertRows 1
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustFirstColumn
            .Selection.Collapse Direction:=wdCollapseStart
            
            .Selection.TypeText Text:="*"
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:=strText1
            .Selection.TypeParagraph
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.InsertRows 1
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:="*"
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:="None of the services for which they are being paid were performed on U.S. soil and are, therefore, not U.S. sourced."
            .Selection.TypeParagraph
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
         Else
         'end 2021/7/27
         
            .Selection.TypeText Text:=strText & strText1
            .Selection.TypeParagraph
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            
         End If 'Added by Morgan 2021/7/27
         
         .Selection.InsertRows 1
         .Selection.Cells.Merge
      End If
 
   If Not (m_A1k28 = "Y55666000" And m_CallPrevForm = "Frmacc24l0") Then 'Added by Morgan 2022/5/18
 
      'ªí§À1
      .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(14.5), RulerStyle:=wdAdjustProportional
      
      .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleNone
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.Collapse Direction:=wdCollapseStart
      With .ActiveDocument.Bookmarks
         .add Range:=g_WordAp.Application.Selection.Range, Name:="BreakPos"
         .DefaultSorting = wdSortByLocation
         .ShowHidden = False
      End With
      
      .Selection.TypeText Text:=m_Footer(1, 1)
      
      'Added by Morgan 2025/9/1 Y54444000­¶§À¯S®í»Ý¨D--Joanne
      If m_A1k28 = "Y54444000" Then
         .Selection.Font.Bold = True
         .Selection.TypeText Text:="Account Holder's Name: "
         .Selection.Font.Bold = False
         .Selection.TypeText Text:="Tai E International Patent and Law Office" & vbCrLf
         .Selection.Font.Bold = True
         .Selection.TypeText Text:="Beneficiary Address: " & vbCrLf
         .Selection.Font.Bold = False
         .Selection.TypeText Text:="COUNTRY: Taiwan, R.O.C." & vbCrLf
         .Selection.TypeText Text:="COUNTRY SUBDIVISION: Taipei" & vbCrLf
         .Selection.TypeText Text:="TOWN NAME: N/A" & vbCrLf
         .Selection.TypeText Text:="STREET NAME: 9Fl.,No.112, Sec.2, Chang-An E. Rd" & vbCrLf
         .Selection.TypeText Text:="POST CODE: 10491" & vbCrLf
         .Selection.Font.Bold = True
         .Selection.TypeText Text:="Beneficiary Bank Name: "
         .Selection.Font.Bold = False
         .Selection.TypeText Text:="Bank of Taiwan DEPT. OF BUSINESS" & vbCrLf
         .Selection.Font.Bold = True
         .Selection.TypeText Text:="Address of the Beneficiary Bank: " & vbCrLf
         .Selection.Font.Bold = False
         .Selection.TypeText Text:="COUNTRY: Taiwan, R.O.C." & vbCrLf
         .Selection.TypeText Text:="COUNTRY SUBDIVISION: Taipei" & vbCrLf
         .Selection.TypeText Text:="TOWN NAME: N/A" & vbCrLf
         .Selection.TypeText Text:="STREET NAME: 120 , Sec. 1. Chongqing S. Rd." & vbCrLf
         .Selection.TypeText Text:="POST CODE:100005" & vbCrLf
         .Selection.Font.Bold = True
         .Selection.TypeText Text:="Account No.: "
         .Selection.Font.Bold = False
         .Selection.TypeText Text:="003007052646 (Multi-Currency Account)" & vbCrLf
         .Selection.Font.Bold = True
         .Selection.TypeText Text:="Account No.: "
         .Selection.Font.Bold = False
         .Selection.TypeText Text:="003001305688 (for Taiwan currency)" & vbCrLf
         .Selection.Font.Bold = True
         .Selection.TypeText Text:="SWIFT Code: "
         .Selection.Font.Bold = False
         .Selection.TypeText Text:="BKTWTWTP"
         If m_DNCurr <> "NTD" And m_iPrintCurrType <> 3 Then
            .Selection.TypeText Text:=vbCrLf
            .Selection.Font.Bold = True
            .Selection.TypeText Text:="Currency Rate: "
            .Selection.Font.Bold = False
            .Selection.TypeText Text:=m_DNCurr & "1.00=NTD" & Format(Val("0" & m_DNRate))
         End If
      End If
      'end 2025/9/1
      
      If m_Footer(2, 1) <> "" Then
         .Selection.MoveRight Unit:=wdCell, Count:=1, Extend:=wdMove
         .Selection.Cells.Split NumRows:=3, NumColumns:=1, MergeBeforeSplit:=False
         .Selection.Cells(1).SetHeight RowHeight:=22, HeightRule:=wdRowHeightAtLeast
         
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdMove
         .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
         With .Selection.Cells(1)
             With .Borders(wdBorderLeft)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             With .Borders(wdBorderRight)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             With .Borders(wdBorderTop)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             With .Borders(wdBorderBottom)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
         End With
         .Selection.ParagraphFormat.LeftIndent = .CentimetersToPoints(0.2)
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
         .Selection.Cells(1).SetHeight RowHeight:=52, HeightRule:=wdRowHeightAtLeast
         .Selection.TypeText Text:=m_Footer(2, 1)
         .Selection.HomeKey
         .Selection.MoveDown Unit:=wdLine, Count:=1
         .Selection.Cells(1).SetHeight RowHeight:=0, HeightRule:=wdRowHeightAtLeast
         .Selection.EndKey Unit:=wdStory
         'Modified by Morgan 2014/8/22 ­Yµe­±Åã¥Ü¥ª¥k¨â­¶¥B´å¼Ð¦b²Ä¤G­¶ªº²Ä¤@¦C®É¤£·|²¾¨ì²Ä¤@­¶ªº³Ì«á¤@¦C
         '.Selection.MoveUp Unit:=wdLine, Count:=1
         .Selection.MoveLeft Unit:=wdCharacter, Count:=1
         'end 2014/8/22
         
      Else
         .Selection.MoveRight Unit:=wdCharacter, Count:=3
      End If
      
      '³Æµù
      If UBound(m_Footer, 2) > 1 Then
         .Selection.SelectRow
         .Selection.Font.Bold = True
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
         .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.TypeText Text:="PS:"
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.TypeText Text:=m_Footer(1, 2)
         For iRow = 3 To UBound(m_Footer, 2)
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.InsertRows 1
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:=m_Footer(1, iRow)
         Next
         .Selection.Font.Bold = False
      End If
   
      .Selection.WholeStory
      .Selection.Font.Name = "Times New Roman"
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .ActiveDocument.Repaginate
      
   If m_A1k28 <> "Y54444000" Then 'Added by Morgan 2025/9/1 Y54444000 ­¶§À­n¦L¸û¦h¸ê°T¡A¥i¤£¥²±±¨î¸õ­¶--Joanne
      
      '¶W¹L1­¶®É´¡¤J­¶½X
      If .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) > 1 Then
      
'Removed by Morgan 2017/6/21 ²¾¨ì¤U­±(²Ä1­¶¤º®e¦h®É·|²Ä2­¶·|ÅÜªÅ¥Õ­¶)
'         .Selection.GoTo what:=wdGoToBookmark, Name:="BreakPos"
'         If .Selection.Information(wdActiveEndPageNumber) = 1 Then
'            .Selection.InsertBreak Type:=wdPageBreak
'            .Selection.TypeParagraph
'            .Selection.TypeParagraph
'         End If
'         .ActiveDocument.Repaginate
'end 2017/6/21
         
         If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
            .ActiveWindow.ActivePane.View.Type = wdPageView
         Else
            .ActiveWindow.View.Type = wdPageView
         End If
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
         .Selection.TypeParagraph
         .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldPage
         .Selection.TypeText Text:="/"
         .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldEmpty, Text:="NUMPAGES ", PreserveFormatting:=True
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
         
'Added by Morgan 2017/6/21
         .Selection.GoTo what:=wdGoToBookmark, Name:="BreakPos"
         If .Selection.Information(wdActiveEndPageNumber) = 1 Then
            .Selection.InsertBreak Type:=wdPageBreak
            .Selection.TypeParagraph
            .ActiveDocument.Repaginate
         Else
            .Selection.SplitTable '­n¤À³Îªí®æ§_«h«HÀY·|©ñ¦bªí®æ¤º
         End If
         .Selection.TypeParagraph
'end 2017/6/21
         
      End If
      
   End If 'Added by Morgan 2025/9/1
   
      .ActiveDocument.Bookmarks("BreakPos").Delete
      .Selection.HomeKey Unit:=wdStory
      
      'Removed by Morgan 2020/4/1 +¥~°Ó¤]­n±a«HÀY--³¯ª÷½¬
      'If Not (m_CP01 = "FCP" Or m_CP01 = "FG" Or m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") Then 'Add By Sindy 2015/7/9 +if ¥~±M¦b¤U­±¥[«HÀY«á¤~¦C¦L
      'Modified by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
      'Removed by Morgan 2025/8/19 ¤º°Ó§ï³æµ§©Î¾ã§å³£­n«HÀY--®Û­^
      'If Left(Pub_StrUserSt03, 2) = "P2" Then
      '   If m_bPrintWord And m_iSpCopies > 0 Then
      '      .ActiveDocument.PrintOut Background:=False, Copies:=m_iSpCopies, Collate:=True
      '      m_iPageCount = .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) 'Added by Morgan 2015/2/26
      '   End If
      'End If
      'end 2025/8/19
      'end 2020/5/7
      'end 2020/4/1
      
      'Modified by Morgan 2014/2/26 +¥~°Ó³£±a«HÀY
      'Modified by Morgan 2014/8/25 §ï¤Á´«¦Lªí¾÷¤è¦¡¦C¦LPDF,¤£¥Î­«¶]Word
      'If m_bWord2Pdf Or m_bSaveWord Or (m_bEditDoc And Left(Pub_StrUserSt03, 2) = "F1") Then
      'Modify By Sindy 2015/7/8 +¥~±M³£±a«HÀY( Or Left(Pub_StrUserSt03, 2) = "F2")
      'Removed by Morgan 2020/4/1 +¥~°Ó¤]­n±a«HÀY--³¯ª÷½¬
      'If m_b2PDF Or m_bWord2Pdf Or m_bSaveWord Or (m_bEditDoc And Left(Pub_StrUserSt03, 2) = "F1") Or (m_CP01 = "FCP" Or m_CP01 = "FG" Or m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") Then
      'Added by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
      'Modified by Morgan 2023/12/20 +m_bolNoPic
      'Modified by Morgan 2025/8/19 ¤º°Ó§ï³æµ§©Î¾ã§å³£­n«HÀY--®Û­^
      'If m_bolNoPic = False And (m_b2PDF Or m_bWord2Pdf Or m_bSaveWord Or Left(Pub_StrUserSt03, 2) <> "P2") Then
      If m_bolNoPic = False Then
      'end 2025/8/19
      'end 2020/5/7
      'end 2020/4/1
        
         'Modified by Morgan 2020/3/31
         'If PUB_ReadDB2File(stFileName, 5) Then
         If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é Then
            PUB_GetLetterPicID "2", m_CP01, iPicNo, iPicNo2, 2, True, Pub_StrUserSt03
         Else
            iPicNo = 5
            iPicNo2 = 9
         End If
         
         
      'Added by Morgan 2022/5/30 ¦C¦L¹ï¶H Y55761010 Bosch Corporation ­¶­º/§À¯Â¤å¦r®æ¦¡ -Franny
      'Modified by Morgan 2022/6/22 +Y5576100 -Franny
      If Left(m_A1k27, 8) = "Y5576101" Or Left(m_A1k27, 8) = "Y5576100" Then
         '­¶­º/
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.Font.Name = "Calibri"
         .Selection.Font.Size = 14
         .Selection.TypeText Text:="   Tai E International Patent & Law Office"
         
         '­¶§À
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
         .Selection.PageSetup.FooterDistance = .CentimetersToPoints(2)
         
         .Selection.Font.Name = "Calibri"
         .Selection.Font.Size = 9
         .Selection.TypeText Text:="9Fl., No. 112, Sec.2, Chang-An E. Rd., Taipei 10491, Taiwan. R.O.C."
         .Selection.TypeParagraph
         .Selection.TypeText Text:="TEL: 886-2-25061023. FAX: 886-2-25068147 (General). 25064319 (Patent). 25090804 (Trademark)"
         
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
         .Selection.EndKey Unit:=wdStory
      Else
      'end 2022/5/30
         
         If PUB_ReadDB2File(stFileName, iPicNo) Then
         'end 2020/3/31
            For ii = 1 To .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) 'Added by Morgan 2014/3/4 ¨C­¶³£­n¦³«HÀY§À
               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.ZOrder 4
               oShape.LockAnchor = True
               oShape.LockAspectRatio = -1
               oShape.Width = .CentimetersToPoints(21)
               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
               oShape.Left = .CentimetersToPoints(0)
               'Modify By Sindy 2015/11/9 ¦]¥~±M­n¥Î¶}µ¡«H«Ê
               'oShape.Top = .CentimetersToPoints(0.5)
               oShape.Top = .CentimetersToPoints(0)
               '2015/11/9 END
               oShape.WrapFormat.Type = wdWrapNone
               .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
               .Selection.EndKey Unit:=wdStory 'Added by Morgan 2018/7/12
            Next ii
            .Selection.HomeKey Unit:=wdStory
            
            'Modified by Morgan 2020/3/31
            'If PUB_ReadDB2File(stFileName, 9) Then
            If iPicNo2 > 0 Then
               If PUB_ReadDB2File(stFileName, iPicNo2) Then
            'end 2020/3/31
                  For ii = 1 To .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)
                     Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                     oShape.ZOrder 4
                     oShape.LockAnchor = True
                     oShape.LockAspectRatio = -1
                     oShape.Width = .CentimetersToPoints(21)
                     oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                     oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                     oShape.Left = .CentimetersToPoints(0)
                     'Added by Morgan 2020/3/31
                     If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é Then
                        oShape.Top = .CentimetersToPoints(27.6)
                     Else
                     'end 2020/3/31
                        oShape.Top = .CentimetersToPoints(27)
                     End If 'Added by Morgan 2020/3/31
                     oShape.WrapFormat.Type = wdWrapNone
                     .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
                     .Selection.EndKey Unit:=wdStory 'Added by Morgan 2018/7/12
                  Next ii
                  .Selection.HomeKey Unit:=wdStory
               End If
               'end 2014/3/4
            End If 'Added by Morgan 2020/3/31
            
         End If
         
      End If 'Added by Morgan 2022/5/30 ¦C¦L¹ï¶H Y55761010 Bosch Corporation ­¶­º/§À¯Â¤å¦r®æ¦¡
         
         If m_bWord2Pdf Then
            .ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
         'Added by Morgan 2014/8/26
         '§ï¤Á´«¦Lªí¾÷¤è¦¡¦C¦LPDF,¤£¥Î­«¶]Word
         ElseIf m_b2PDF Then
            PrintWord2PDF
         'end 2014/8/26
         End If
         
         'Move by Lydia 2020/12/25 ±qIf m_bWord2Pdf Then ¤W­±²¾¤U¨Ó; ¦]¬°12/23ªºFCP¦~¶O¾ã§åµo¤å¨ÌÂÂ¦³¯È¥»PDF¥¼¦s¤JTyping2
         'Add By Sindy 2015/7/9 ¥~±M¦b¦¹³B¤~¦C¦L
         'If (m_CP01 = "FCP" Or m_CP01 = "FG" Or m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") Then 'Removed by Morgan 2020/4/1 +°Ó¼Ð¤]­n±a«HÀY--®Û­^
         'Added by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
         'Removed by Morgan 2025/8/19 ¤º°Ó§ï³æµ§©Î¾ã§å³£­n«HÀY--®Û­^
         'If Left(Pub_StrUserSt03, 2) <> "P2" Then
         'end 2025/8/19
         'end 2020/5/7
         
            If m_bPrintWord And m_iSpCopies > 0 Then
               .ActiveDocument.PrintOut Background:=False, Copies:=m_iSpCopies, Collate:=True
               m_iPageCount = .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) 'Added by Morgan 2015/2/26
            End If
            
         'End If 'Added by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
         'End If 'Removed by Morgan 2020/4/1 +°Ó¼Ð¤]­n±a«HÀY--®Û­^
         '2015/7/9 END
         'end ----Move by Lydia 2020/12/25
         
         If m_bSaveWord Then
            RidFile m_EFilePath
            .ActiveDocument.SaveAs m_EFilePath
         End If
      End If 'Added by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
      'End If 'Removed by Morgan 2020/4/1 +°Ó¼Ð¤]­n±a«HÀY--®Û­^
      
   End If 'Added by Morgan 2022/5/18
   
   End With
   'Modified by Lydia 2019/04/09 §ï¦¨¦@¥Î¼Ò²Õ
   'RePosWord bVisible, m_WordLeft, m_WordTop 'Added by Morgan 2014/6/26
   Pub_RePosWord g_WordAp, bVisible, m_WordLeft, m_WordTop
   
   If m_bEditDoc Then
      g_WordAp.Visible = True
      g_WordAp.Activate
   Else
      g_WordAp.ActiveDocument.Close wdDoNotSaveChanges
      If bVisible = False Then
         g_WordAp.Quit wdDoNotSaveChanges
         Set g_WordAp = Nothing 'Added by Lydia 2017/12/12 Á×§K§Ö³t¶}±ÒWord,µ{¦¡¥X¿ù
      Else
         g_WordAp.Visible = True
      End If
   End If
   
   Exit Sub
   
ErrHnd:
   'Resume
   If Err.Number <> 0 Then
'Modified by Morgan 2014/6/24
'      If iResumeCnt > 3 Then
'         MsgBox "¿ù»~ : " & Err.Description, vbCritical
'      Else
'         iResumeCnt = iResumeCnt + 1
'         Select Case Err.Number
'            Case 91:
'               g_WordAp.Documents.Add
'               Resume Next
'            Case 462:
'               Set g_WordAp = New Word.Application
'               Resume
'            Case Else:
'               MsgBox "¿ù»~ : " & Err.Description, vbCritical
'         End Select
'      End If
      MsgBox "¿ù»~ : " & Err.Description, vbCritical
'end 2014/6/24
   End If
End Sub

'Added by Morgan 2013/10/30 ·s®æ¦¡
'Modified by Morgan 2014/8/21 +pBolSpecial ¯S®í¦C¦L¹ï¶H½Ð´Ú³æ
'Modified by Morgan 2015/11/20 ªíÀY§ï5Äæ,²Ä2Äæ«á¥[¤@ªÅ¥ÕÄæ¥H«K°Ï¹j¦a§}»P«á­±ªº¸ê®Æ
Private Sub runWordJapaneseNew(Optional pBolPlusForm As Boolean = False)

   Dim stTmp As String
   Dim iRow As Integer, iCol As Integer, lngCellWidth As Long
   Dim strLastCountry As String, iFCol As Integer, iCols As Integer, bolPrintCountry As Boolean
   Dim strFontSize As String
   Dim iResumeCnt As Integer
   Dim iPos1 As Integer, iPos2 As Integer, iPos3 As Integer
   Dim bVisible As Boolean, stFileName As String, iHeadCount As Integer
   Dim oShape
   Dim strText As String
   Dim strLstNo As String
   Dim dblOffFeeSub As Double
   Dim dblAttFeeSub As Double
   'Added by Morgan 2024/3/26 ¥N¦¬¥N¥I
   Dim dblCFAttFee As Double
   Dim dblCFAttFeeSub As Double
   'end 2024/3/26
   Dim bolNewRow As Boolean
   Dim bolSkip As Boolean 'Added by Morgan 2013/12/16
   Dim ii As Integer 'Added by Morgan 2014/3/4
   'Added by Morgan 2014/8/21
   Dim arrHead() As String
   Dim arrSubject() As String
   Dim arrItem() As INVITEM
   Dim arrFooter() As String
   Dim bolAddRemark As Boolean 'Added by Morgan 2015/4/16
   Dim dblA1K39 As Double, dblA1K40 As Double 'Add By Sindy 2021/10/26
   
   If m_A1k28 = "Y52075000" Or m_A1k28 = "Y52075010" Then bolAddRemark = True 'Added by Morgan 2015/4/17
   
   If pBolPlusForm Then
      SetPlusFormData
      arrHead = m_PlusHead
      arrFooter = m_PlusFooter
   Else
      arrHead = m_Head
      arrFooter = m_Footer
   End If
   arrSubject = m_Subject
   arrItem = m_Item
   'end 2014/8/21
   
   strFontSize = 12
   
   'Modified by Lydia 2019/04/09 §ï¦¨¦@¥Î¼Ò²Õ
   'If NewDoc(bVisible, m_WordLeft, m_WordTop) = False Then Exit Sub 'Added by Morgan 2014/6/26
   If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = False Then Exit Sub
   
On Error GoTo ErrHnd

'Removed by Morgan 2014/6/26
'   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
'
'   g_WordAp.Documents.Add
'
'   bVisible = g_WordAp.Visible
'
'   '¤£Åã¥Ü¥i¯à·|¦³°ÝÃD
'   If m_bShowWord Then
'      g_WordAp.Visible = True
'      g_WordAp.WindowState = wdWindowStateMaximize
'      g_WordAp.ActiveWindow.ActivePane.View.Zoom.Percentage = 100
'   Else
'      g_WordAp.Visible = False
'   End If
   
   With g_WordAp.Application
      'ª©­±³]©w
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.5)
      'Modified by Morgan 2014/1/23 ­¶­º¥[°ª(²Ä2­¶¦L¯È¥»·|À£¨ì«HÀY)
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(3.53)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.3)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(3)
      .Selection.PageSetup.FooterDistance = .CentimetersToPoints(3)
      
      .Selection.PageSetup.CharsLine = 40
      .Selection.PageSetup.LinesPage = 38
      
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = strFontSize
      
      'Added by Morgan 2015/4/16
      If bolAddRemark Then
         .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
         .Selection.ParagraphFormat.LineSpacing = 9
      End If
      'end 2015/4/16
      
      'Added by Lydia 2020/06/23 «HÀY¤U¤è¥[¦L¡uPAID³¹¡v
      If m_bPAID = True Then
          If PUB_ReadDB2File(stFileName, "57") = True Then
               'ªÅ5¦æ (¦]¬°­¶­º½d³ò,¤ñ­^¤å¦h¤@¦æ)
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               If bolAddRemark Then '¦æ¶Z¸û¤p,¦h2¦æ
                    .Selection.TypeParagraph
                    .Selection.TypeParagraph
               End If
               .Selection.MoveUp Unit:=wdLine, Count:=4 '²¾¨ì²Ä¤@¦æ
               If bolAddRemark Then .Selection.MoveUp Unit:=wdLine, Count:=2
               
               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.Left = .CentimetersToPoints(13.5)
               .Selection.MoveDown Unit:=wdLine, Count:=4 '²¾¨ì³Ì«á¤@¦æ
               If bolAddRemark Then .Selection.MoveDown Unit:=wdLine, Count:=2
          Else
               .Selection.TypeParagraph
          End If
      Else
      'end 2020/06/23
      '«O¯d«HÀYªÅ¶¡(2¦æ)
         '.Selection.TypeParagraph 'Removed by Morgan 2014/1/23 ­¶­º¤w¥[°ª
         .Selection.TypeParagraph
      End If 'Added by Lydia 2020/06/23
      
      '¦æ¶Z
      With .Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
        .DisableLineHeightGrid = True
      End With
      
      
      '·s¼Wªí®æ(1*5)
      'Modified by Morgan 2015/4/22
      .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=5
      With .Selection.Tables(1)
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
      End With
      
      'Added by Morgan 2015/4/16
      If bolAddRemark Then
         .Selection.SelectRow
         .Selection.InsertRows 1
         .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(10.6), RulerStyle:=wdAdjustProportional
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.MoveRight Unit:=wdCell, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.Font.Size = 10
         .Selection.Font.Bold = False
         .Selection.TypeText Text:="This is an original official electronic invoice;"
         .Selection.TypeParagraph
         .Selection.TypeText Text:="(No signature is required.)"
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
      End If
      'end 2015/4/16
      'end 2015/4/22
      
      '³]©wªí®æ°ª«×Äæ¼e
      .Selection.SelectRow
      'Modified by Morgan 2014/2/20
      '.Selection.Cells.VerticalAlignment = wdAlignVerticalBottom
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.9), RulerStyle:=wdAdjustProportional
      .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(8), RulerStyle:=wdAdjustProportional
      .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(0.7), RulerStyle:=wdAdjustProportional
      .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
      .Selection.Cells(4).VerticalAlignment = wdCellAlignVerticalTop
      
      .Selection.InsertRows UBound(arrHead, 2) + 1
      .Selection.Collapse Direction:=wdCollapseStart
            
      .Selection.SelectRow
      .Selection.Cells.Merge
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.Cells.VerticalAlignment = wdAlignVerticalBottom
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      'Modified by Morgan 2015/4/16
      '.Selection.ParagraphFormat.SpaceBefore = .CentimetersToPoints(0.6)
      'Modified by Morgan 2018/8/30
      'Á×§K®ö¶O¯È±i¤Î¸`¬Ù§@·~®É¶¡,¤W¶¡¶Z©T©w³]¬°0.2cm¥H´î¤Ö¸õ­¶¾÷²v Ex:X10712759 --ªü½¬
      'If bolAddRemark Then
      '   .Selection.ParagraphFormat.SpaceBefore = .CentimetersToPoints(0.2)
      'Else
      '   .Selection.ParagraphFormat.SpaceBefore = .CentimetersToPoints(0.6)
      'End If
      .Selection.ParagraphFormat.SpaceBefore = .CentimetersToPoints(0.2)
      'end 2018/8/30
      'end 2015/4/16
      .Selection.ParagraphFormat.SpaceAfter = .CentimetersToPoints(0.6)
      .Selection.TypeText Text:="                           "
      .Selection.Font.Size = 18
      .Selection.Font.Bold = True
      .Selection.TypeText Text:=m_Title
      .Selection.Font.Size = strFontSize
      .Selection.TypeText Text:="              NO. " & m_strDN
      .Selection.Font.Bold = False
      .Selection.MoveRight Unit:=wdCharacter, Count:=2
      
      iPos1 = 0
      iPos2 = 0
      
      'ªíÀY¦C1
      For iRow = 1 To UBound(arrHead, 2)
         bolSkip = False
         .Selection.Cells(1).SetHeight RowHeight:=12, HeightRule:=wdRowHeightAtLeast
         If arrHead(1, iRow) <> "" Then
            If iPos1 = 0 Then iPos1 = iRow
            .Selection.TypeText Text:=arrHead(1, iRow)
         End If
         If arrHead(2, iRow) <> "" Then
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            If arrHead(3, iRow) & arrHead(4, iRow) = "" Then
               .Selection.MoveRight Unit:=wdCharacter, Count:=4, Extend:=wdExtend
               .Selection.Cells.Merge
               bolSkip = True
               If iPos3 < iPos2 Then iPos3 = iRow
            Else
               If iPos2 = 0 Then iPos2 = iRow
            End If
            .Selection.TypeText Text:=arrHead(2, iRow)
            
         ElseIf arrHead(1, iRow) <> "" Then
            .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
            .Selection.Cells.Merge
         Else
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
         End If
         
         If bolSkip = False Then
            .Selection.MoveRight Unit:=wdCharacter, Count:=2
            If arrHead(3, iRow) <> "" Then
               .Selection.TypeText Text:=arrHead(3, iRow)
            End If
            
            If arrHead(4, iRow) <> "" Then
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
               .Selection.TypeText Text:=arrHead(4, iRow)
            Else
               .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
               .Selection.Cells.Merge
            End If
         End If
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
      Next
      
      'Added by Morgan 2015/11/20
      If m_A1k28 = "Y51333010" Then
         '±N²Ä2Äæ¦X¨Ö
         If iPos2 > iPos1 And (iPos3 = 0 Or iPos3 > iPos2) Then
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            intI = UBound(arrHead, 2) - iPos3 + 2
            .Selection.MoveUp Unit:=wdLine, Count:=intI
            If iPos3 > iPos2 Then
               intI = iPos3 - iPos2 - 1
               .Selection.MoveUp Unit:=wdLine, Count:=intI, Extend:=wdExtend
            End If
            .Selection.Cells.Merge
            
            If iPos3 = 0 Then
               intI = 1
            Else
               intI = UBound(arrHead, 2) - iPos3 + 2
            End If
            .Selection.MoveDown Unit:=wdLine, Count:=intI
         End If
      End If
      
      .Selection.SelectRow
      .Selection.Cells.Merge
      
      'Added by Morgan 2019/3/27 ½Õ¾ãª©­±¯à¦L¦b¤@­¶--³¯ª÷½¬
      '¨ú®ø[¥ó¦W]¤Î[¶µ¥Ø]¤W¤èªºªÅ¥Õ¦æ
      If m_CP01 = "FCT" Then
         .Selection.InsertRows 2
         .Selection.Collapse Direction:=wdCollapseStart
         With .Selection.Cells.Borders(wdBorderTop)
         .LineStyle = wdLineStyleSingle
         .LineWidth = wdLineWidth050pt
         .ColorIndex = wdAuto
         End With
      Else
      'end 2019/3/27
      
         .Selection.InsertRows 3
         .Selection.Collapse Direction:=wdCollapseStart
         With .Selection.Cells.Borders(wdBorderBottom)
         .LineStyle = wdLineStyleSingle
         .LineWidth = wdLineWidth050pt
         .ColorIndex = wdAuto
         End With
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
         
      End If 'Added by Morgan 2019/3/27
      
      '¼ÐÃD
      For iRow = 1 To UBound(arrSubject, 2)
         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
         If arrSubject(1, iRow) <> "" Then
            .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
            If arrSubject(2, iRow) = "" Then
               .Selection.TypeText Text:=arrSubject(1, iRow)
            Else
               .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.TypeText Text:=arrSubject(1, iRow)
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.TypeText Text:=arrSubject(2, iRow)
            End If
            
         Else
            .Selection.Cells(1).SetHeight RowHeight:=16, HeightRule:=wdRowHeightAtLeast
            If arrSubject(2, iRow) <> "" Then
               'Added by Morgan 2023/7/7 ¥Ó½Ð¤H¥i¯à·|§é¦æ ex:X11208950
               If Left(arrSubject(2, iRow), 4) = "¥XÄ@¤H¡G" Or Left(arrSubject(2, iRow), 4) = "¡@¡@¡@¡@" Then
                  .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=False
                  .Selection.Cells.VerticalAlignment = wdCellAlignVerticalTop
                  .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.86), RulerStyle:=wdAdjustProportional
                  .Selection.Collapse Direction:=wdCollapseStart
                  .Selection.MoveRight Unit:=wdCell, Count:=1
                  .Selection.TypeText Text:=Left(arrSubject(2, iRow), 4)
                  .Selection.MoveRight Unit:=wdCell, Count:=1
                  .Selection.TypeText Text:=Mid(arrSubject(2, iRow), 5)
               Else
               'end 2023/7/7
                  .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
                  .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
                  .Selection.Collapse Direction:=wdCollapseStart
                  .Selection.MoveRight Unit:=wdCell, Count:=1
                  .Selection.TypeText Text:=arrSubject(2, iRow)
               End If 'Added by Morgan 2023/7/7
               
            ElseIf arrSubject(3, iRow) <> "" Then
               .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.9), RulerStyle:=wdAdjustProportional
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.MoveRight Unit:=wdCell, Count:=2
               .Selection.TypeText Text:=arrSubject(3, iRow)
            End If
         End If
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
         .Selection.InsertRows 1
      Next
      
      'Added by Morgan 2019/3/27 ½Õ¾ãª©­±¯à¦L¦b¤@­¶--³¯ª÷½¬
      '¨ú®ø[¥ó¦W]¤Î[¶µ¥Ø]¤W¤èªºªÅ¥Õ¦æ
      If m_CP01 = "FCT" Then
         .Selection.Collapse Direction:=wdCollapseStart
      Else
      'end 2019/3/27
      
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.InsertRows 1
         
      End If 'Added by Morgan 2019/3/27
      
      '©ú²ÓªíÀY
      With .Selection.Cells(1)
         With .Borders(wdBorderLeft)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderRight)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderTop)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderBottom)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
      End With
      
      If m_iPrintCurrType = 1 Or m_iPrintCurrType = 2 Then
         strText = "NTD"
      Else
         strText = m_DNCurr
      End If
      
      .Selection.SelectRow
      .Selection.Cells.Shading.Texture = wdTexture5Percent 'Added by Morgan 2014/2/18
      .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
      If m_bSpecialNew5 Then
         If m_CP01 = "P" Then
            .Selection.Cells.Split NumRows:=1, NumColumns:=5, MergeBeforeSplit:=False
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(8.5), RulerStyle:=wdAdjustProportional
            .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
            .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
            .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
         Else
            .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=False
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(9.75), RulerStyle:=wdAdjustProportional
            .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.25), RulerStyle:=wdAdjustProportional
            .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.6), RulerStyle:=wdAdjustProportional
         End If
      Else
         .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=False
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(11.7), RulerStyle:=wdAdjustProportional
         .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustProportional
      End If
      .Selection.Cells(1).SetHeight RowHeight:=12, HeightRule:=wdRowHeightAtLeast
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText Text:="¶µ¡@¥Ø"
      'Added by Morgan 2024/3/26
      If m_bSpecialNew5 Then
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.TypeText Text:="§@·~®É¶¡"
         .Selection.TypeParagraph
         .Selection.TypeText Text:="HR"
         If m_CP01 = "P" Then
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Selection.TypeText Text:=PUB_GetUniText(Me.Name, "¥N¦¬¥N¥I")
            .Selection.TypeParagraph
            .Selection.TypeText Text:=strText
         End If
      End If
      'end 2024/3/26
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText Text:="¬F©²®Æª÷"
      .Selection.TypeParagraph
      .Selection.TypeText Text:=strText
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      'Modified by Morgan 2022/8/17
      '.Selection.TypeText Text:="¹ú©Ò¤â“g®Æ"
      .Selection.TypeText Text:=PUB_GetUniText(Me.Name, "¥»©Ò¶O¥Î")
      'end 2022/8/17
      .Selection.TypeParagraph
      .Selection.TypeText Text:=strText
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.InsertRows 1
      .Selection.Cells.Shading.Texture = wdTextureNone 'Added by Morgan 2014/2/18
      .Selection.Collapse Direction:=wdCollapseStart
      
      strLstNo = ""
      '©ú²Ó
      For iRow = 1 To UBound(arrItem)
         strText = ""
         dblCFAttFee = 0 'Added by Morgan 2024/3/26
         
         If Right(arrItem(iRow).iNo, 2) = "99" Then
            dblOffFeeSub = dblOffFeeSub + Val(Format(arrItem(iRow).iAmt))
         Else
            'Modified by Morgan 2014/8/28 +¯S®í½Ð´Ú³æ
            If pBolPlusForm Then
               dblAttFeeSub = dblAttFeeSub + Val(Format(arrItem(iRow).IXAmt))
            Else
               dblAttFeeSub = dblAttFeeSub + Val(Format(arrItem(iRow).iAmt))
            End If
            'end 2014/8/28
         End If
         
         '­Y¬°«e¶µªº³W¶O«h¦X¨Ö
         If arrItem(iRow).iNo = strLstNo & "99" Then
            'Modified by Morgan 2014/10/13
            'If arrItem(iRow).IDescTail <> "" Or arrItem(iRow).IDescX <> "" Then
            'Added by Morgan 2014/12/30
            If pBolPlusForm Then
               .Selection.MoveLeft Unit:=wdCell, Count:=2
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
               .Selection.TypeText Text:=arrItem(iRow).iAmt
               .Selection.MoveRight Unit:=wdCell, Count:=1
               If .Selection.Text <> "" Then
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
               End If
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
               
            'end 2014/12/30
            ElseIf (arrItem(iRow).IDescTail <> "" Or arrItem(iRow).IDescX <> "") Then
               If arrItem(iRow).IDescX <> "" Then
                  strText = arrItem(iRow).IDescX
               Else
                  strText = "¬F©²®Æª÷"
               End If
               If arrItem(iRow).IDescTail <> "" Then
                  strText = strText & "¡G" & arrItem(iRow).IDescTail
               End If
               
               .Selection.InsertRows 1
               .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleNone
               'Adde by Morgan 2024/3/26
               If m_bSpecialNew5 Then
                  If m_CP01 = "P" Then
                     .Selection.Cells.Split NumRows:=1, NumColumns:=6, MergeBeforeSplit:=True
                     .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.6), RulerStyle:=wdAdjustProportional
                     .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(7.9), RulerStyle:=wdAdjustProportional
                     .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
                     .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
                     .Selection.Cells(5).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
                  Else
                     .Selection.Cells.Split NumRows:=1, NumColumns:=5, MergeBeforeSplit:=True
                     .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.6), RulerStyle:=wdAdjustProportional
                     .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(9.15), RulerStyle:=wdAdjustProportional
                     .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.25), RulerStyle:=wdAdjustProportional
                     .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.6), RulerStyle:=wdAdjustProportional
                  End If
               Else
               'end 2023/3/26
                  .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=True
                  .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.6), RulerStyle:=wdAdjustFirstColumn
                  .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(11.1), RulerStyle:=wdAdjustFirstColumn
                  .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustFirstColumn
               End If
               .Selection.Cells(1).Borders(wdBorderRight).LineStyle = wdLineStyleNone
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
               .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
               .Selection.TypeText Text:=strText
               'Adde by Morgan 2024/3/26
               If m_bSpecialNew5 Then
                  If m_CP01 = "P" Then
                     .Selection.MoveRight Unit:=wdCell, Count:=2
                  Else
                     .Selection.MoveRight Unit:=wdCell, Count:=1
                  End If
               End If
               'end 2024/3/26
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
               .Selection.TypeText Text:=arrItem(iRow).iAmt
               .Selection.MoveRight Unit:=wdCharacter, Count:=2
               
            Else
               .Selection.MoveLeft Unit:=wdCell, Count:=2
               
               If arrItem(iRow - 1).IDescTail <> "" Then
                  .Selection.MoveUp Unit:=wdLine, Count:=1
               End If
               
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
               .Selection.TypeText Text:=arrItem(iRow).iAmt
               .Selection.MoveRight Unit:=wdCell, Count:=1
               If arrItem(iRow - 1).IDescTail <> "" Then
                  .Selection.MoveDown Unit:=wdLine, Count:=1
               ElseIf .Selection.Text <> "" Then
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
               End If
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
            End If
         
         Else
            .Selection.InsertRows 1
            '¶µ¥Ø¥Îµê½u¤À¹j
            If iRow > 1 Then
               .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleDashLargeGap
               .Selection.Cells.Borders(wdBorderTop).LineWidth = wdLineWidth050pt
            End If
            'Adde by Morgan 2024/3/26
            If m_bSpecialNew5 Then
               If m_CP01 = "P" Then
                  .Selection.Cells.Split NumRows:=1, NumColumns:=5, MergeBeforeSplit:=True
                  .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(8.5), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
               Else
                  .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=True
                  .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(9.75), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.25), RulerStyle:=wdAdjustProportional
                  .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.6), RulerStyle:=wdAdjustProportional
               End If
            Else
            'end 2024/3/26
               .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(11.7), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustProportional
            End If
            .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            
            bolNewRow = False
            
            'Added by Morgan 2014/10/13
            If pBolPlusForm Then
               .Selection.TypeText Text:=arrItem(iRow).IDescHead
               If Right(arrItem(iRow).iNo, 2) = "99" Then
                  .Selection.MoveRight Unit:=wdCell, Count:=1
               Else
                  .Selection.MoveRight Unit:=wdCell, Count:=2
               End If
               
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
               .Selection.Font.Bold = False
               
               If Right(arrItem(iRow).iNo, 2) = "99" Then
                  .Selection.TypeText Text:=arrItem(iRow).iAmt
                  .Selection.MoveRight Unit:=wdCharacter, Count:=2
               Else
                  .Selection.TypeText Text:=arrItem(iRow).IXAmt
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
               End If
               
            Else
            'end 2014/10/13
            
               '³Ì«á¤@¶µ
               If iRow = UBound(arrItem) Then
                  .Selection.TypeText Text:=arrItem(iRow).IDesc
               '«á­±±µ¬Û¦P½Ð´Ú¶µ¥Øªº³W¶O
               ElseIf arrItem(iRow + 1).iNo = arrItem(iRow).iNo & "99" Then
                  .Selection.TypeText Text:=arrItem(iRow).IDescHead
                  bolNewRow = True
               '¨ä¥L
               Else
                  .Selection.TypeText Text:=arrItem(iRow).IDesc
               End If
               
               'Added by Morgan 2024/3/26
               If m_bSpecialNew5 Then
                  'Added by Morgan 2024/12/19 ¥N¦¬¥N¥I¤£¥²­­©w¤uµ{®v©Ó¿ìªº¶µ¥Ø--Kimi
                  dblCFAttFee = 0
                  If m_CP01 = "P" Then
                     If Len(arrItem(iRow).iNo) = 3 Then
                        dblCFAttFee = GetA1L17(m_strDN, arrItem(iRow).iNo & "98")
                     End If
                  End If
                  'end 2024/12/19
                  
                  .Selection.MoveRight Unit:=wdCell, Count:=1
                  If InStr(m_EngCP10List, "," & arrItem(iRow).iNo & ",") > 0 Then
                     .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                     .Selection.Font.Bold = False
                     If m_CP01 = "P" Then
                        'Removed by Morgan 2024/12/19 §ï¨ì¥~¼h,¥N¦¬¥N¥I¤£¥²­­©w¤uµ{®v©Ó¿ìªº¶µ¥Ø--Kimi
                        'dblCFAttFee = 0
                        'If Len(arrItem(iRow).iNo) = 3 Then
                        '   dblCFAttFee = GetA1L17(m_strDN, arrItem(iRow).iNo & "98")
                        'End If
                        
                        'Modified by Morgan 2024/4/26
                        '.Selection.TypeText Text:=Round(arrItem(iRow).iAmt / 200, 2) '§@·~®É¶¡
                        .Selection.TypeText Text:=Round((arrItem(iRow).iAmt - dblCFAttFee) / 200, 2) '§@·~®É¶¡
                        'end 2024/4/26
                        .Selection.MoveRight Unit:=wdCell, Count:=1
                        
                        'Removed by Morgan 2024/12/19 §ï¨ì¥~¼h,¥N¦¬¥N¥I¤£¥²­­©w¤uµ{®v©Ó¿ìªº¶µ¥Ø--Kimi
                        '.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                        '.Selection.Font.Bold = False
                        'If dblCFAttFee > 0 Then
                        '   dblCFAttFeeSub = dblCFAttFeeSub + dblCFAttFee
                        '   .Selection.TypeText Text:=Format(dblCFAttFee, "#,##0.00") '¥N¦¬¥N¥I
                        'End If
                        'end 2024/12/19
                        
                     Else
                        .Selection.TypeText Text:=Round(arrItem(iRow).iAmt / 6000, 2) '§@·~®É¶¡
                     End If
                     
                  ElseIf m_CP01 = "P" Then
                     .Selection.MoveRight Unit:=wdCell, Count:=1
                  End If
                  
                  'Added by Morgan 2024/12/19 ¥N¦¬¥N¥I¤£¥²­­©w¤uµ{®v©Ó¿ìªº¶µ¥Ø--Kimi
                  If m_CP01 = "P" Then
                     If dblCFAttFee > 0 Then
                        dblCFAttFeeSub = dblCFAttFeeSub + dblCFAttFee
                        .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                        .Selection.Font.Bold = False
                        .Selection.TypeText Text:=Format(dblCFAttFee, "#,##0.00") '¥N¦¬¥N¥I
                     End If
                  End If
                  'end 2024/12/19
               End If
               'end 2024/3/26
               If Right(arrItem(iRow).iNo, 2) = "99" Then
                  .Selection.MoveRight Unit:=wdCell, Count:=1
               Else
                  .Selection.MoveRight Unit:=wdCell, Count:=2
               End If
               
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
               .Selection.Font.Bold = False
               'Added by Morgan 2024/3/26
               If dblCFAttFee > 0 Then
                  .Selection.TypeText Text:=Format(Val(Format(arrItem(iRow).iAmt)) - dblCFAttFee, "#,##0.00")
               Else
               'end 2024/3/26
                  .Selection.TypeText Text:=arrItem(iRow).iAmt
               End If
               
               If Right(arrItem(iRow).iNo, 2) = "99" Then
                  .Selection.MoveRight Unit:=wdCharacter, Count:=2
               Else
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
               End If
               
               If bolNewRow And arrItem(iRow).IDescTail <> "" Then
                  .Selection.InsertRows 1
                  .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleNone
                  'Adde by Morgan 2024/3/26
                  If m_bSpecialNew5 Then
                     If m_CP01 = "P" Then
                        .Selection.Cells.Split NumRows:=1, NumColumns:=6, MergeBeforeSplit:=True
                        .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.6), RulerStyle:=wdAdjustProportional
                        .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(7.9), RulerStyle:=wdAdjustProportional
                        .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
                        .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
                        .Selection.Cells(5).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
                     Else
                        .Selection.Cells.Split NumRows:=1, NumColumns:=5, MergeBeforeSplit:=True
                        .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.6), RulerStyle:=wdAdjustProportional
                        .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(9.15), RulerStyle:=wdAdjustProportional
                        .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.25), RulerStyle:=wdAdjustProportional
                        .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.6), RulerStyle:=wdAdjustProportional
                     End If
                  Else
                  'end 2023/3/26
                     .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=True
                     .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.6), RulerStyle:=wdAdjustFirstColumn
                     .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(11.1), RulerStyle:=wdAdjustFirstColumn
                     .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustFirstColumn
                  End If
                  .Selection.Cells(1).Borders(wdBorderRight).LineStyle = wdLineStyleNone
                  .Selection.Collapse Direction:=wdCollapseStart
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
                  'Modified by Morgan 2022/7/27
                  'strText = "¹ú©Ò¤â“g®Æ¡G" & arrItem(iRow).IDescTail
                  strText = PUB_GetUniText(Me.Name, "¥»©Ò¶O¥Î") & "¡G" & arrItem(iRow).IDescTail
                  'end 2022/7/27
                  .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
                  .Selection.TypeText Text:=strText
                  'Adde by Morgan 2024/3/26
                  If m_bSpecialNew5 Then
                     If m_CP01 = "P" Then
                        .Selection.MoveRight Unit:=wdCell, Count:=2
                     Else
                        .Selection.MoveRight Unit:=wdCell, Count:=1
                     End If
                  End If
                  'end 2024/3/26
                  .Selection.MoveRight Unit:=wdCell, Count:=2
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
               End If
               
            End If 'Added by Morgan 2014/10/13
               
         End If
         strLstNo = arrItem(iRow).iNo
      Next
      
      '¤p­p
      .Selection.InsertRows 1
      'Added by Morgan 2024/3/26
      If m_bSpecialNew5 Then
         If m_CP01 = "P" Then
            .Selection.Cells.Split NumRows:=1, NumColumns:=5, MergeBeforeSplit:=True
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(8.5), RulerStyle:=wdAdjustProportional
            .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
            .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
            .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
         Else
            .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=True
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(9.75), RulerStyle:=wdAdjustProportional
            .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.25), RulerStyle:=wdAdjustProportional
            .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.6), RulerStyle:=wdAdjustProportional
         End If
      Else
      'end 2024/3/26
         .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(11.7), RulerStyle:=wdAdjustProportional
         .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustProportional
      End If
      
      .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleDouble
      .Selection.Cells.Borders(wdBorderTop).LineWidth = wdLineWidth050pt
      .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
      .Selection.Cells(1).SetHeight RowHeight:=30, HeightRule:=wdRowHeightAtLeast
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText Text:="¤p¡@­p"
      
      
      'Added by Morgan 2024/3/26
      If m_bSpecialNew5 Then
         If m_CP01 = "P" Then
            .Selection.MoveRight Unit:=wdCharacter, Count:=2
            '¥N¦¬¥N¥I
            If dblCFAttFeeSub > 0 Then
               strText = Format(dblCFAttFeeSub, "#,##0.00")
               .Selection.TypeText Text:=strText
            End If
         Else
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
         End If
      End If
      'end 2024/3/26
      
      '³W¶O
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      If dblOffFeeSub > 0 Then
         strText = Format(dblOffFeeSub, "#,##0.00")
         dblA1K40 = dblOffFeeSub
         .Selection.TypeText Text:=strText
         
         If m_iPrintCurrType = 2 Or m_iPrintCurrType = 4 Then
            .Selection.TypeParagraph
            If m_iPrintCurrType = 2 Then
               dblA1K40 = Trunc(dblOffFeeSub / Val(m_DNRate))
               strText = "(" & m_DNCurr & Format(dblA1K40, "#,##0.00") & ")"
            Else
'               dblA1K40 = Trunc(dblOffFeeSub * Val(m_DUsdRate))
               strText = "(USD" & Format(Trunc(dblOffFeeSub * Val(m_DUsdRate)), "#,##0.00") & ")"
            End If
            .Selection.TypeText Text:=strText
         End If
         
         'Add By Sindy 2021/10/26 °O¿ý ½Ð´Ú³W¶O¥~¹ôª÷ÃB
         If pBolPlusForm = False Then
            strSql = "Update ACC1K0 Set A1K40=" & dblA1K40 & " Where A1K01='" & m_strDN & "'"
            adoTaie.Execute strSql, intI
         End If
         '2021/10/26 END
      Else
         'Add By Sindy 2021/12/6 °O¿ý ½Ð´Ú³W¶O¥~¹ôª÷ÃB
         If pBolPlusForm = False Then
            strSql = "Update ACC1K0 Set A1K40=" & dblOffFeeSub & " Where A1K01='" & m_strDN & "'"
            adoTaie.Execute strSql, intI
         End If
         '2021/12/6 END
      End If
      
      'ªA°È¶O
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      If dblAttFeeSub > 0 Then
         'Added by Morgan 2024/3/26
         If dblCFAttFeeSub > 0 Then
            strText = Format(dblAttFeeSub - dblCFAttFeeSub, "#,##0.00")
         Else
         'end 2024/3/26
            strText = Format(dblAttFeeSub, "#,##0.00")
         End If
         dblA1K39 = dblAttFeeSub
         .Selection.TypeText Text:=strText
         
         If m_iPrintCurrType = 2 Or m_iPrintCurrType = 4 Then
            .Selection.TypeParagraph
            If m_iPrintCurrType = 2 Then
               dblA1K39 = Trunc(dblAttFeeSub / Val(m_DNRate))
               strText = "(" & m_DNCurr & Format(dblA1K39, "#,##0.00") & ")"
            Else
'               dblA1K39 = Trunc(dblAttFeeSub * Val(m_DUsdRate))
               strText = "(USD" & Format(Trunc(dblAttFeeSub * Val(m_DUsdRate)), "#,##0.00") & ")"
            End If
            .Selection.TypeText Text:=strText
         End If
         
         'Add By Sindy 2021/10/26 °O¿ý ½Ð´ÚªA°È¶O¥~¹ôª÷ÃB
         If pBolPlusForm = False Then
            strSql = "Update ACC1K0 Set A1K39=" & dblA1K39 & " Where A1K01='" & m_strDN & "'"
            adoTaie.Execute strSql, intI
         End If
         '2021/10/26 END
      Else
         'Add By Sindy 2021/12/6 °O¿ý ½Ð´ÚªA°È¶O¥~¹ôª÷ÃB
         If pBolPlusForm = False Then
            strSql = "Update ACC1K0 Set A1K39=" & dblAttFeeSub & " Where A1K01='" & m_strDN & "'"
            adoTaie.Execute strSql, intI
         End If
         '2021/12/6 END
      End If
      
      .Selection.MoveRight Unit:=wdCharacter, Count:=2

      '¦X­p
      .Selection.SelectRow
      .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
      'Added by Morgan 2024/3/26
      If m_bSpecialNew5 Then
         If m_CP01 = "P" Then
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(8.5), RulerStyle:=wdAdjustProportional
         Else
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(9.75), RulerStyle:=wdAdjustProportional
         End If
      Else
      'end 2024/3/26
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(11.7), RulerStyle:=wdAdjustProportional
      End If
      .Selection.Cells(1).SetHeight RowHeight:=30, HeightRule:=wdRowHeightAtLeast
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.Font.Bold = True
      .Selection.Font.Size = 12
      .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText Text:="¦X¡@­p"
      .Selection.Font.Size = strFontSize
      
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Font.Bold = True
      'Modified by Morgan 2014/8/28 +¯S®í½Ð´Ú³æ
      If pBolPlusForm Then
         strText = m_Sum(2, 1) & Format(dblOffFeeSub + dblAttFeeSub, "#,##0.00")
      Else
         strText = m_Sum(2, 1) & m_Sum(3, 1)
      End If
      .Selection.TypeText Text:=strText
      
      If m_iPrintCurrType = 2 Or m_iPrintCurrType = 4 Then
         .Selection.TypeParagraph
         If m_iPrintCurrType = 2 Then
            strText = m_DNCurr & Format(Trunc(dblOffFeeSub / Val(m_DNRate)) + Trunc(dblAttFeeSub / Val(m_DNRate)), "#,##0.00")
         Else
            strText = "USD" & Format(Trunc(dblOffFeeSub * Val(m_DUsdRate)) + Trunc(dblAttFeeSub * Val(m_DUsdRate)), "#,##0.00")
            
            UpdateA1K38 m_strDN, Trunc(dblOffFeeSub * Val(m_DUsdRate)) + Trunc(dblAttFeeSub * Val(m_DUsdRate)) 'Added by Morgan 2021/1/14 §ó·s½Ð´Ú³æ¬üª÷Á`ÃB
         End If
         .Selection.TypeText Text:=strText
      End If
      
      .Selection.Font.Bold = False
      .Selection.MoveRight Unit:=wdCharacter, Count:=3
      .Selection.InsertRows 1
      .Selection.Cells.Merge
      .Selection.Cells(1).Borders(wdBorderTop).LineStyle = wdLineStyleNone 'Added by Morgan 2014/8/15
      
      'Added by Morgan 2016/8/8
      With .ActiveDocument.Bookmarks
         .add Range:=g_WordAp.Application.Selection.Range, Name:="BreakPos"
         .DefaultSorting = wdSortByLocation
         .ShowHidden = False
      End With
      'end 2016/8/8
      
      '¥u¦³®æ¦¡2(¥x¹ô+¥~¹ô¦X­p)¤~±a
      If m_iPrintCurrType = 2 And m_DNCurr = "USD" Then
         .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
         .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.9), RulerStyle:=wdAdjustFirstColumn
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.Font.Size = 10
         .Selection.TypeText Text:="µù¡G"
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         
         .Selection.Font.Size = 10
         'Modified by Morgan 2022/7/27
         'strText = "¦X­pÇUÄæÇR¥Üþêþò" & m_DNCurr & "ÇV¡B¦X­pÇUÄæÇRþÝÆêþùÇUNTDþÞÇp´«ºâþêþòÇiÇUþúÇVÇQþâ¡B¤p­pÇUÄæÇRþÝþäÇr" & m_DNCurr & "ÇU ñÃBþúÆèÇqÇeþì¡C "
         strText = PUB_GetUniText(Me.Name, "µù¤@") & m_DNCurr & PUB_GetUniText(Me.Name, "µù¤G") & m_DNCurr & PUB_GetUniText(Me.Name, "µù¤T")
         'end 2022/7/27
         .Selection.TypeText Text:=strText
         .Selection.TypeParagraph
         .Selection.Font.Size = 10
         'Modified by Morgan 2022/7/27
         'strText = "Çeþò¡B¤p­pÇUÄæÇRþÝþäÇr" & m_DNCurr & "ÇV¡B¤p“g“|¥H¤UÇy¤ÁÇq¬BþùþùÆêÇeþìÇUþú¡BüÁ»ÚÇU½Ð¨DþìÇ`þàª÷ÃBÇO1" & m_DNCurr & "µ{«×ÇU’½®tþßÆèÇqÇeþì¡Cþç¤F©Ó¤UþèÆê¡C "
         strText = PUB_GetUniText(Me.Name, "µù¥|") & m_DNCurr & PUB_GetUniText(Me.Name, "µù¤­") & m_DNCurr & PUB_GetUniText(Me.Name, "µù¤»")
         'end 2022/7/27
         .Selection.TypeText Text:=strText
         '.Selection.TypeParagraph
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.InsertRows 1
         .Selection.Font.Size = strFontSize
         .Selection.Cells.Merge
      End If
      
      'ªí§À1
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.Collapse Direction:=wdCollapseStart
'Removed by Morgan 2016/8/8 ¶µ¥Ø¦h®É·|¦b¸õ­¶²Å¸¹«e´N¶W¹L¤@­¶¡A§ï²¾¨ì¦X­p«á­±
'      With .ActiveDocument.Bookmarks
'         .add Range:=g_WordAp.Application.Selection.Range, Name:="BreakPos"
'         .DefaultSorting = wdSortByLocation
'         .ShowHidden = False
'      End With
'end 2016/8/8
      .Selection.TypeText Text:=arrFooter(1, 1)
      
      If arrFooter(2, 1) <> "" Then
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.Cells.Split NumRows:=3, NumColumns:=2, MergeBeforeSplit:=False
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.Cells(1).SetHeight RowHeight:=36, HeightRule:=wdRowHeightAtLeast
         .Selection.MoveDown Unit:=wdLine, Count:=1
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
         With .Selection.Cells
             With .Borders(wdBorderLeft)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             With .Borders(wdBorderRight)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             With .Borders(wdBorderTop)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             With .Borders(wdBorderBottom)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
             .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
             .Borders.Shadow = False
         End With
         .Selection.ParagraphFormat.LeftIndent = .CentimetersToPoints(0.2)
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
         .Selection.Cells(1).SetHeight RowHeight:=52, HeightRule:=wdRowHeightAtLeast
         .Selection.TypeText Text:=arrFooter(2, 1)
         .Selection.HomeKey
         .Selection.MoveDown Unit:=wdLine, Count:=1
         .Selection.Cells(1).SetHeight RowHeight:=0, HeightRule:=wdRowHeightAtLeast
         .Selection.MoveDown Unit:=wdLine, Count:=1
      Else
         .Selection.MoveRight Unit:=wdCharacter, Count:=3
      End If
      
      '³Æµù
      If UBound(arrFooter, 2) > 1 Then
         .Selection.SelectRow
         .Selection.Font.Bold = True
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
         .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.TypeText Text:="PS:"
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.TypeText Text:=arrFooter(1, 2)
         For iRow = 3 To UBound(arrFooter, 2)
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.InsertRows 1
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:=arrFooter(1, iRow)
         Next
         .Selection.Font.Bold = False
      End If
      
      .Selection.WholeStory
      .Selection.Font.Name = "²Ó©úÅé"
      .Selection.WholeStory
      .Selection.Font.Name = "Times New Roman"
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .ActiveDocument.Repaginate
      '¶W¹L1­¶®É´¡¤J­¶½X
      If .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) > 1 Then
         .Selection.GoTo what:=wdGoToBookmark, Name:="BreakPos"
         .Selection.InsertBreak Type:=wdPageBreak
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .ActiveDocument.Repaginate
         If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
            .ActiveWindow.ActivePane.View.Type = wdPageView
         Else
            .ActiveWindow.View.Type = wdPageView
         End If
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
         .Selection.TypeParagraph
         .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldPage
         .Selection.TypeText Text:="/"
         .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldEmpty, Text:="NUMPAGES ", PreserveFormatting:=True
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
      End If
      .ActiveDocument.Bookmarks("BreakPos").Delete
      .Selection.HomeKey Unit:=wdStory
      
      'Removed by Morgan 2020/4/1 +¥~°Ó¤]­n±a«HÀY--³¯ª÷½¬
      'If Not (m_CP01 = "FCP" Or m_CP01 = "FG" Or m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") Then 'Add By Sindy 2015/7/9 +if ¥~±M¦b¤U­±¥[«HÀY«á¤~¦C¦L
      'Modified by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
      'Removed by Morgan 2025/8/19 ¤º°Ó§ï³æµ§©Î¾ã§å³£­n«HÀY--®Û­^
      'If Left(Pub_StrUserSt03, 2) = "P2" Then
      '   If m_bPrintWord And m_iSpCopies > 0 Then
      '      .ActiveDocument.PrintOut Background:=False, Copies:=m_iSpCopies, Collate:=True
      '      m_iPageCount = .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) 'Added by Morgan 2015/2/26
      '   End If
      'End If
      'end 2025/8/19
      'end 2020/5/7
      'end 2020/4/1
      
      'Modified by Morgan 2014/2/26 +¥~°Ó³£±a«HÀY
      'Modified by Morgan 2014/8/25 §ï¤Á´«¦Lªí¾÷¤è¦¡¦C¦LPDF,¤£¥Î­«¶]Word
      'If m_bWord2Pdf Or m_bSaveWord Or (m_bEditDoc And Left(Pub_StrUserSt03, 2) = "F1") Then
      'Modify By Sindy 2015/7/8 +¥~±M³£±a«HÀY( Or Left(Pub_StrUserSt03, 2) = "F2")
      'Removed by Morgan 2020/4/1 +¥~°Ó¤]­n±a«HÀY--³¯ª÷½¬
      'If m_b2PDF Or m_bWord2Pdf Or m_bSaveWord Or (m_bEditDoc And Left(Pub_StrUserSt03, 2) = "F1") Or (m_CP01 = "FCP" Or m_CP01 = "FG" Or m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") Then
      'Added by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
      'Removed by Morgan 2025/8/19 ¤º°Ó§ï³æµ§©Î¾ã§å³£­n«HÀY--®Û­^
      'If m_b2PDF Or m_bWord2Pdf Or m_bSaveWord Or Left(Pub_StrUserSt03, 2) <> "P2" Then
      'end 2025/8/19
      'end 2020/5/7
      'end 2020/4/1
      
         'Modified by Morgan 2020/3/31
         'If PUB_ReadDB2File(stFileName, 5) Then
         If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é Then
            PUB_GetLetterPicID "2", m_CP01, iPicNo, iPicNo2, 3, True, Pub_StrUserSt03
         Else
            iPicNo = 5
            iPicNo2 = 9
         End If
         If PUB_ReadDB2File(stFileName, iPicNo) Then
         'end 2020/3/31
            'Modified by Morgan 2014/3/4 ¨C­¶³£­n¦³«HÀY§À
            For ii = 1 To .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)
               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.ZOrder 4
               oShape.LockAnchor = True
               oShape.LockAspectRatio = -1
               oShape.Width = .CentimetersToPoints(21)
               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
               oShape.Left = .CentimetersToPoints(0)
               oShape.Top = .CentimetersToPoints(0.5)
               .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
            Next ii
            .Selection.HomeKey Unit:=wdStory
            
            'Modified by Morgan 2020/3/31
            'If PUB_ReadDB2File(stFileName, 9) Then
            If iPicNo2 > 0 Then
               If PUB_ReadDB2File(stFileName, iPicNo2) Then
            'end 2020/3/31
                  For ii = 1 To .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)
                     Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                     oShape.ZOrder 4
                     oShape.LockAnchor = True
                     oShape.LockAspectRatio = -1
                     oShape.Width = .CentimetersToPoints(21)
                     oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                     oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                     oShape.Left = .CentimetersToPoints(0)
                     'Added by Morgan 2020/3/31
                     If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é Then
                        oShape.Top = .CentimetersToPoints(27.6)
                     Else
                     'end 2020/3/31
                        oShape.Top = .CentimetersToPoints(27)
                     End If 'Added by Morgan 2020/3/31
                     .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
                  Next ii
                  .Selection.HomeKey Unit:=wdStory
               End If
               'end 2014/3/4
            End If 'Added by Morgan 2020/3/31
            
         End If
         
         If m_bWord2Pdf Then
            .ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
            
         'Added by Morgan 2014/8/26
         '§ï¤Á´«¦Lªí¾÷¤è¦¡¦C¦LPDF,¤£¥Î­«¶]Word
         ElseIf m_b2PDF Then
            PrintWord2PDF pBolPlusForm
         'end 2014/8/26
         
         End If
         
         'Move by Lydia 2020/12/25 ±qIf m_bWord2Pdf Then ¤W­±²¾¤U¨Ó; ¦]¬°12/23ªºFCP¦~¶O¾ã§åµo¤å¨ÌÂÂ¦³¯È¥»PDF¥¼¦s¤JTyping2
         'Add By Sindy 2015/7/9 ¥~±M¦b¦¹³B¤~¦C¦L
         'If (m_CP01 = "FCP" Or m_CP01 = "FG" Or m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") Then 'Removed by Morgan 2020/4/1 +¥~°Ó¤]­n±a«HÀY--³¯ª÷½¬
         'Added by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
         'Removed by Morgan 2025/8/19 ¤º°Ó§ï³æµ§©Î¾ã§å³£­n«HÀY--®Û­^
         'If Left(Pub_StrUserSt03, 2) <> "P2" Then
         'end 2025/8/19
         'end 2020/5/7
         
            If m_bPrintWord And m_iSpCopies > 0 Then
               .ActiveDocument.PrintOut Background:=False, Copies:=m_iSpCopies, Collate:=True
               m_iPageCount = .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) 'Added by Morgan 2015/2/26
            End If
            
         'End If 'Added by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
         'End If 'Removed by Morgan 2020/4/1 +¥~°Ó¤]­n±a«HÀY--³¯ª÷½¬
         '2015/7/9 END
         'end --- Move by Lydia 2020/12/25
         
         If m_bSaveWord Then
            RidFile m_EFilePath
            .ActiveDocument.SaveAs m_EFilePath
         End If
         
      'End If 'Added by Morgan 2020/5/7 ¤º°Ó§ï¦^¤£­n«HÀY--®Û­^
      'End If 'Removed by Morgan 2020/4/1 +¥~°Ó¤]­n±a«HÀY--³¯ª÷½¬
      
   End With
   'Modified by Lydia 2019/04/09 §ï¦¨¦@¥Î¼Ò²Õ
   'RePosWord bVisible, m_WordLeft, m_WordTop 'Added by Morgan 2014/6/26
   Pub_RePosWord g_WordAp, bVisible, m_WordLeft, m_WordTop
   
   If m_bEditDoc Then
      g_WordAp.Visible = True
      g_WordAp.Activate
   Else
      g_WordAp.ActiveDocument.Close wdDoNotSaveChanges
      If bVisible = False Then
         g_WordAp.Quit wdDoNotSaveChanges
         Set g_WordAp = Nothing 'Added by Lydia 2017/12/12 Á×§K§Ö³t¶}±ÒWord,µ{¦¡¥X¿ù
      Else
         g_WordAp.Visible = True
      End If
   End If
   
   Exit Sub
   
ErrHnd:
   If Err.Number <> 0 Then
'Modified by Morgan 2014/6/26
'      If iResumeCnt > 3 Then
'         MsgBox "¿ù»~ : " & Err.Description, vbCritical
'      Else
'         iResumeCnt = iResumeCnt + 1
'         Select Case Err.NUMBER
'            Case 91:
'               g_WordAp.Documents.Add
'               Resume Next
'            Case 462:
'               Set g_WordAp = New Word.Application
'               Resume
'            Case Else:
'               MsgBox "¿ù»~ : " & Err.Description, vbCritical
'         End Select
'      End If
      MsgBox "¿ù»~ : " & Err.Description, vbCritical
'end 2014/6/26
   End If
End Sub

'Added by Morgan 2013/10/31 ·s®æ¦¡
'Modified by Morgan 2015/11/20 ªíÀY§ï5Äæ,²Ä2Äæ«á¥[¤@ªÅ¥ÕÄæ¥H«K°Ï¹j¦WºÙ¦a§}»P«á­±ªº¸ê®Æ
Private Sub runWordChineseNew()

   Dim stTmp As String
   Dim iRow As Integer, iCol As Integer, lngCellWidth As Long
   Dim strLastCountry As String, iFCol As Integer, iCols As Integer, bolPrintCountry As Boolean
   Dim strFontSize As String
   Dim iResumeCnt As Integer
   Dim iPos1 As Integer, iPos2 As Integer
   Dim bVisible As Boolean, stFileName As String, iHeadCount As Integer
   Dim strText As String
   Dim oShape
   Dim strLstNo As String
   Dim dblOffFeeSub As Double
   Dim dblAttFeeSub As Double
   Dim bolNewRow As Boolean
   Dim ii As Integer 'Added by Morgan 2014/3/4
   Dim mColNum As Integer 'Added by Lydia 2015/04/09 °O¿ý©ú²ÓÄæ¼Æ
   Dim mColW(1 To 10) As Single   'Added by Lydia 2015/04/09 °O¿ý©ú²ÓÄæ¼e
   Dim dbChiAmt As Double, dbChiUAmt As Double 'Added by Lydia 2015/04/09 Á`­p->½Ð´Úª÷ÃB,¥~¹ô
   Dim rsA As New ADODB.Recordset
   Dim dblA1K39 As Double, dblA1K40 As Double 'Add By Sindy 2021/10/26
   
   strFontSize = 12
   'Modified by Lydia 2019/04/09 §ï¦¨¦@¥Î¼Ò²Õ
   'If NewDoc(bVisible, m_WordLeft, m_WordTop) = False Then Exit Sub 'Added by Morgan 2014/6/26
   If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = False Then Exit Sub
   
On Error GoTo ErrHnd

'Removed by Morgan 2014/6/26
'   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
'
'   g_WordAp.Documents.add
'
'   bVisible = g_WordAp.Visible
'
'   '¤£Åã¥Ü¥i¯à·|¦³°ÝÃD
'   If m_bShowWord Then
'      g_WordAp.Visible = True
'      g_WordAp.WindowState = wdWindowStateMaximize
'      g_WordAp.ActiveWindow.ActivePane.View.Zoom.Percentage = 100
'   Else
'      g_WordAp.Visible = False
'   End If
   
   With g_WordAp.Application
      'ª©­±³]©w
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.5)
      'Modified by Morgan 2014/1/23 ­¶­º¥[°ª(²Ä2­¶¦L¯È¥»·|À£¨ì«HÀY)
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(3.53)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.3)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(3)
      .Selection.PageSetup.FooterDistance = .CentimetersToPoints(3)
      
      .Selection.PageSetup.CharsLine = 40
      .Selection.PageSetup.LinesPage = 38
      
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = strFontSize
      
      'Added by Lydia 2020/06/23 «HÀY¤U¤è¥[¦L¡uPAID³¹¡v
      If m_bPAID = True Then
          If PUB_ReadDB2File(stFileName, "57") = True Then
               'ªÅ5¦æ (¦]¬°­¶­º½d³ò,¤ñ­^¤å¦h¤@¦æ)
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               .Selection.MoveUp Unit:=wdLine, Count:=4 '²¾¨ì²Ä¤@¦æ
               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.Left = .CentimetersToPoints(13.5)
               .Selection.MoveDown Unit:=wdLine, Count:=4 '²¾¨ì³Ì«á¤@¦æ
          Else
               .Selection.TypeParagraph
          End If
      Else
      'end 2020/06/23
      '«O¯d«HÀYªÅ¶¡(2¦æ)
         '.Selection.TypeParagraph 'Removed by Morgan 2014/1/23 ­¶­º¤w¥[°ª
         .Selection.TypeParagraph
      End If 'Added by Lydia 2020/06/23
      
      '¦æ¶Z
      With .Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
        .DisableLineHeightGrid = True
      End With
      
      '·s¼Wªí®æ(1*5)
      .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=5
      
      With .Selection.Tables(1)
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
      End With
    
      '³]©wªí®æ°ª«×Äæ¼e
      .Selection.SelectRow
      'Modified by Morgan 2014/2/20
      '.Selection.Cells.VerticalAlignment = wdAlignVerticalBottom
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.9), RulerStyle:=wdAdjustProportional
      .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(8), RulerStyle:=wdAdjustProportional
      .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(0.7), RulerStyle:=wdAdjustProportional
      .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
      .Selection.Cells(4).VerticalAlignment = wdCellAlignVerticalTop
      
      .Selection.InsertRows UBound(m_Head, 2)
      .Selection.Collapse Direction:=wdCollapseStart
      
      'Added by Morgan 2024/7/9
      If m_A1k03 = "Y56042000" And Left(strCust1, 8) = "X5655900" Then
         .Selection.HomeKey Unit:=wdStory
         '.Selection.ParagraphFormat.SpaceAfter = .CentimetersToPoints(0.6)
         .Selection.TypeText Text:="                           "
         .Selection.Font.Size = 18
         .Selection.Font.Bold = True
         .Selection.TypeText Text:=m_Title
         .Selection.Font.Size = strFontSize
         .Selection.TypeText Text:="              NO. " & m_strDN
         .Selection.Font.Bold = False
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
      End If
      'end 2024/7/9
      
      iPos1 = 0
      iPos2 = 0
      
      'ªíÀY¦C1
      For iRow = 1 To UBound(m_Head, 2)
         .Selection.Cells(1).SetHeight RowHeight:=12, HeightRule:=wdRowHeightAtLeast
         If m_Head(1, iRow) <> "" Then
            If iPos1 = 0 Then iPos1 = iRow
            .Selection.TypeText Text:=m_Head(1, iRow)
         End If
         If m_Head(2, iRow) <> "" Then
            If iPos2 = 0 Then iPos2 = iRow
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:=m_Head(2, iRow) '½Ð´Ú¹ï¶H
         ElseIf m_Head(1, iRow) <> "" Then
            .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
            .Selection.Cells.Merge
         Else
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
         End If
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
         
         If m_Head(3, iRow) <> "" Then
            .Selection.TypeText Text:=m_Head(3, iRow) '½Ð´Ú¤é´Á;¶Q¤è¨÷¸¹ ;¥»©Ò®×¸¹
         End If
         If m_Head(4, iRow) <> "" Then
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:=m_Head(4, iRow)
         Else
            .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
            .Selection.Cells.Merge
         End If
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
      Next
      
      'Added by Morgan 2015/11/20
      'Modified by Morgan 2024/2/20 +Y51817010
      'Modified by Morgan 2024/9/6 +Y52459030 ¨Ã§ï§PÂ_¦C¦L¹ï¶H--®Û­^
      If m_A1k27 = "Y51333010" Or m_A1k27 = "Y51817010" Or m_A1k27 = "Y52459030" Then
         '±N²Ä2Äæ¦X¨Ö
         If iPos2 > iPos1 Then
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.MoveUp Unit:=wdLine, Count:=1
            intI = UBound(m_Head, 2) - iPos2
            .Selection.MoveUp Unit:=wdLine, Count:=intI, Extend:=wdExtend
            .Selection.Cells.Merge
            .Selection.MoveDown Unit:=wdLine, Count:=1
         End If
      End If
      'end 2015/11/20
            
      .Selection.SelectRow
      .Selection.Cells.Merge
      .Selection.InsertRows 4
      .Selection.Collapse Direction:=wdCollapseStart
      .Selection.MoveDown Unit:=wdLine, Count:=1
      
      If m_A1k03 = "Y56042000" And Left(strCust1, 8) = "X5655900" Then
         .Selection.TypeText Text:="    ½Ð´Ú¤º®e¦p¤U¡G"
      Else
      
         .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
         .Selection.TypeText Text:="                           "
         .Selection.Font.Size = 14
         .Selection.Font.Bold = True
         .Selection.TypeText Text:=m_Title '½Ð´Ú³æÃþ§O: ex.¦¬¶O³qª¾³æ
         'Added by Lydia 2015/04/08 ¾ã§å½Ð´Ú³æªíÀY¤£¦L½s¸¹(¦b©ú²Ó)
         If m_bolChiDB = False Then
           .Selection.Font.Size = strFontSize
           .Selection.TypeText Text:="      ½s¸¹: " & m_strDN '½Ð´Ú³æ¸¹
           .Selection.Font.Bold = False
         End If
         
      End If
      
      With .Selection.Cells.Borders(wdBorderBottom)
         .LineStyle = wdLineStyleSingle
         .LineWidth = wdLineWidth050pt
         .ColorIndex = wdAuto
      End With
      
      .Selection.MoveRight Unit:=wdCharacter, Count:=2
      '¼ÐÃD
      For iRow = 1 To UBound(m_Subject, 2)
         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
         If m_Subject(1, iRow) <> "" Then
            .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
            If m_Subject(2, iRow) = "" Then
               .Selection.TypeText Text:=m_Subject(1, iRow)
            Else
               .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.TypeText Text:=m_Subject(1, iRow) '¥DÃD:
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.TypeText Text:=m_Subject(2, iRow)
            End If
            
         Else
            .Selection.Cells(1).SetHeight RowHeight:=16, HeightRule:=wdRowHeightAtLeast
            If m_Subject(2, iRow) <> "" Then
               .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.TypeText Text:=m_Subject(2, iRow) 'iROW=2 ¥Ó½Ð¤H,iROW=3 ¦WºÙ °Ó¼Ð¡GCUP,4 Ãþ§O, 5 µù¥U¸¹
               
            ElseIf m_Subject(3, iRow) <> "" Then
               .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=False
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.9), RulerStyle:=wdAdjustProportional
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.MoveRight Unit:=wdCell, Count:=2
               .Selection.TypeText Text:=m_Subject(3, iRow)
            End If
         End If
         .Selection.MoveRight Unit:=wdCharacter, Count:=2
         .Selection.InsertRows 1
      Next
      
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Selection.InsertRows 1
      
      '©ú²ÓªíÀY
      With .Selection.Cells(1)
         With .Borders(wdBorderLeft)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderRight)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderTop)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
          With .Borders(wdBorderBottom)
              .LineStyle = wdLineStyleSingle
              .LineWidth = wdLineWidth100pt
              .ColorIndex = wdAuto
          End With
      End With
      
      If m_iPrintCurrType = 1 Or m_iPrintCurrType = 2 Then
         strText = "NTD"
      Else
         strText = m_DNCurr
      End If
      
      .Selection.SelectRow
      .Selection.Cells.Shading.Texture = wdTexture5Percent 'Added by Morgan 2014/2/18
      .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
      'Added by Lydia 2015/04/09 +¤¤¤åª©¾ã§å½Ð´Ú³æ(©ú²Ó)
      If m_bolChiDB Then
        '¯È¥»SampleÄæ¼e,PDF©ïÀY·|§é¦æ
'         mColW(1) = 1.11: mColW(2) = 2: mColW(3) = 2.19: mColW(4) = 2.58: mColW(5) = 1.9
'         mColW(6) = 2.12: mColW(7) = 1.06: mColW(8) = 1.48: mColW(9) = 1.9: mColW(10) = 1.06
         mColW(1) = 1.11: mColW(2) = 2: mColW(3) = 2.19: mColW(4) = 2.28: mColW(5) = 1.9
         mColW(6) = 2.32: mColW(7) = 1.26: mColW(8) = 1.48: mColW(9) = 1.9: mColW(10) = 1.06
         If m_iPrintCurrType = 2 Or m_iPrintCurrType = 4 Then
            mColNum = 10
         Else
            mColNum = 9
         End If
        .Selection.Cells.Split NumRows:=1, NumColumns:=mColNum, MergeBeforeSplit:=False
        For ii = 1 To mColNum - 1
           If m_iPrintCurrType = 2 Or m_iPrintCurrType = 4 Then
              .Selection.Cells(ii).SetWidth ColumnWidth:=.CentimetersToPoints(mColW(ii)), RulerStyle:=wdAdjustProportional
           Else
               If ii < mColNum - 1 Then
                  .Selection.Cells(ii).SetWidth ColumnWidth:=.CentimetersToPoints(mColW(ii)), RulerStyle:=wdAdjustProportional
               End If
           End If
        Next ii
        .Selection.Cells(1).SetHeight RowHeight:=12, HeightRule:=wdRowHeightAtLeast
        .Selection.Collapse Direction:=wdCollapseStart
        'column(1)
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.TypeText Text:="Ãþ§O"
        'column(2)
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Selection.TypeText Text:="¥»©Ò®×¸¹"
        .Selection.TypeParagraph
        .Selection.TypeText Text:="©¼©Ò®×¸¹"
        'column(3)
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Selection.TypeText Text:="µù¥U¸¹¼Æ/"
        .Selection.TypeParagraph
        .Selection.TypeText Text:="¥Ó½Ð®×¸¹"
        'column(4)
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Selection.TypeText Text:="°Ó¼Ð¦WºÙ"
        'column(5)
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Selection.TypeText Text:="®×¥ó©Ê½è"
        'column(6)
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Selection.TypeText Text:="½Ð´Ú³æ¸¹"
        'column(7)
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.TypeText Text:="³W¶O"
        .Selection.TypeParagraph
        .Selection.TypeText Text:=strText
        'column(8)
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.TypeText Text:="ªA°È¶O"
        .Selection.TypeParagraph
        .Selection.TypeText Text:=strText
        'column(9)
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.TypeText Text:="½Ð´Úª÷ÃB"
        .Selection.TypeParagraph
        .Selection.TypeText Text:=strText
        'column(10)
         If m_iPrintCurrType = 2 Or m_iPrintCurrType = 4 Then
            If m_iPrintCurrType = 4 Then
               strExc(5) = "¬üª÷": strExc(6) = "USD"
            Else
                '§ì¥~¹ô
                strExc(4) = "Select A1Y02 From ACC1Y0 Where A1Y01='" & m_DNCurr & "' "
                rsA.CursorLocation = adUseClient
                rsA.Open strExc(4), cnnConnection, adOpenStatic, adLockReadOnly
                strExc(5) = "": strExc(6) = m_DNCurr
                If rsA.RecordCount > 0 Then
                   strExc(5) = rsA.Fields(0)
                End If
            End If
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Selection.TypeText Text:=strExc(5)
            .Selection.TypeParagraph
            .Selection.TypeText Text:=strExc(6)
         End If
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
      Else
        .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=False
        .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(11.7), RulerStyle:=wdAdjustProportional
        .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustProportional
        .Selection.Cells(1).SetHeight RowHeight:=12, HeightRule:=wdRowHeightAtLeast
        .Selection.Collapse Direction:=wdCollapseStart
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.TypeText Text:="¶µ¡@¥Ø"
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.TypeText Text:="³W¡@¶O"
        .Selection.TypeParagraph
        .Selection.TypeText Text:=strText '³W¶O¹ô§O
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.TypeText Text:="ªA°È¶O"
        .Selection.TypeParagraph
        .Selection.TypeText Text:=strText  'ªA°È¶O¹ô§O
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
      End If
      .Selection.InsertRows 1
      .Selection.Cells.Shading.Texture = wdTextureNone 'Added by Morgan 2014/2/18
      .Selection.Collapse Direction:=wdCollapseStart
      
      strLstNo = ""
      '©ú²Ó
      For iRow = 1 To UBound(m_Item) '©ú²Ó¶µ¥Ø¼Æ
         strText = ""
        If Right(m_Item(iRow).iNo, 2) = "99" Then '¤p­p³W¶O©MªA°È¶O
           dblOffFeeSub = dblOffFeeSub + Val(Format(m_Item(iRow).iAmt))
        Else
           dblAttFeeSub = dblAttFeeSub + Val(Format(m_Item(iRow).iAmt))
        End If

         '­Y¬°«e¶µªº³W¶O«h¦X¨Ö
         If m_Item(iRow).IChiCno & m_Item(iRow).iNo = strLstNo & "99" Then
            If m_Item(iRow).IDescTail <> "" Or m_Item(iRow).IDescX <> "" Then '¶µ¥Ø»¡©ú¯S§O(¦³³W¶O¤º®e)
               If m_Item(iRow).IDescX <> "" Then
                  strText = m_Item(iRow).IDescX
               Else
                  strText = "Official fee"
               End If
               
               If m_Item(iRow).IDescTail <> "" Then
                  strText = strText & ":" & m_Item(iRow).IDescTail
               End If
               
               .Selection.InsertRows 1
               .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleNone
                'Added by Lydia 2015/04/09 +¤¤¤åª©¾ã§å½Ð´Ú³æ(©ú²Ó),¤£¦L§é¦©
                If m_bolChiDB Then
                    .Selection.Collapse Direction:=wdCollapseStart
                    .Selection.MoveRight Unit:=wdCharacter, Count:=1
                    .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    .Selection.TypeText Text:=m_Item(iRow).IChiCls
                    .Selection.MoveRight Unit:=wdCell, Count:=1
                    .Selection.TypeText Text:=m_Item(iRow).IChiCno
                    .Selection.MoveRight Unit:=wdCell, Count:=1
                    .Selection.TypeText Text:=m_Item(iRow).IChiApp
                    .Selection.MoveRight Unit:=wdCell, Count:=1
                    .Selection.TypeText Text:=m_Item(iRow).IChiCna
                    .Selection.MoveRight Unit:=wdCell, Count:=2
                    .Selection.TypeText Text:=m_Item(iRow).IChiA1k01
                    .Selection.MoveRight Unit:=wdCell, Count:=1
                    .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                    .Selection.TypeText Text:=Format(m_Item(iRow).iAmt)
                    .Selection.MoveRight Unit:=wdCharacter, Count:=mColNum - 6 '¤U¤@¦æ
                Else
                '----------------------------------
                    .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=True
                    .Selection.Cells(1).Borders(wdBorderRight).LineStyle = wdLineStyleNone
                    .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.6), RulerStyle:=wdAdjustFirstColumn
                    .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(11.1), RulerStyle:=wdAdjustFirstColumn
                    .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustFirstColumn
                    .Selection.Collapse Direction:=wdCollapseStart
                    .Selection.MoveRight Unit:=wdCharacter, Count:=1
                    .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
                    .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
                    .Selection.TypeText Text:=strText
                    .Selection.MoveRight Unit:=wdCharacter, Count:=1
                    .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                    .Selection.TypeText Text:=m_Item(iRow).iAmt
                    .Selection.MoveRight Unit:=wdCharacter, Count:=2
                End If
            Else  '¨S¦³¶µ¥Ø»¡©ú¯S§O(¦³³W¶O¤º®e)
               'Added by Lydia 2015/04/09 +¤¤¤åª©¾ã§å½Ð´Ú³æ(©ú²Ó)
                If m_bolChiDB Then
                    .Selection.MoveLeft Unit:=wdCell, Count:=2  '¦^¨ìªA°È¶O
                    .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                    .Selection.TypeText Text:=Format(m_Item(iRow).iAmt)
                    
                    .Selection.MoveRight Unit:=wdCell, Count:=2 '¦^¨ì½Ð´Úª÷ÃB
                    .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                    .Selection.TypeText Text:=Format(m_Item(iRow).IChiAmt)
                    dbChiAmt = dbChiAmt + Val(m_Item(iRow).IChiAmt)
                    If m_iPrintCurrType = 2 Or m_iPrintCurrType = 4 Then
                        .Selection.MoveRight Unit:=wdCharacter, Count:=1
                        .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                        .Selection.TypeText Text:=Format(m_Item(iRow).IChiUAmt)
                        dbChiUAmt = dbChiUAmt + Val(m_Item(iRow).IChiUAmt)
                        
                        UpdateA1K38 m_Item(iRow).IChiA1k01, Format(m_Item(iRow).IChiUAmt) 'Added by Morgan 2021/1/14 §ó·s½Ð´Ú³æ¬üª÷Á`ÃB
                    End If
                    If .Selection.Text <> "" Then
                       .Selection.MoveRight Unit:=wdCharacter, Count:=1
                    End If
                    .Selection.MoveRight Unit:=wdCharacter, Count:=1
                
                Else
                '--------------------
                    .Selection.MoveLeft Unit:=wdCell, Count:=2
                    If m_Item(iRow - 1).IDescTail <> "" Then
                       .Selection.MoveUp Unit:=wdLine, Count:=1
                    End If
                    .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                    .Selection.TypeText Text:=m_Item(iRow).iAmt

                    .Selection.MoveRight Unit:=wdCell, Count:=1
                    If m_Item(iRow - 1).IDescTail <> "" Then
                       .Selection.MoveDown Unit:=wdLine, Count:=1
                    ElseIf .Selection.Text <> "" Then
                       .Selection.MoveRight Unit:=wdCharacter, Count:=1
                    End If
                    .Selection.MoveRight Unit:=wdCharacter, Count:=1
                End If
            End If
         
         Else    '«D¦X¨Ö
            .Selection.InsertRows 1
            'Added by Lydia 2015/04/09 +¤¤¤åª©¾ã§å½Ð´Ú³æ(©ú²Ó),¤£¦L§é¦©
            If m_bolChiDB Then
               .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
               
               bolNewRow = False
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
               .Selection.TypeText Text:=m_Item(iRow).IChiCls
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.TypeText Text:=m_Item(iRow).IChiCno
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.TypeText Text:=m_Item(iRow).IChiApp
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.TypeText Text:=m_Item(iRow).IChiCna
               .Selection.MoveRight Unit:=wdCell, Count:=1
               '³Ì«á¤@¶µ
               If iRow = UBound(m_Item) Then
                  .Selection.TypeText Text:=m_Item(iRow).IDesc
               '«á­±±µ¬Û¦P½Ð´Ú¶µ¥Øªº³W¶O
               ElseIf m_Item(iRow + 1).iNo = m_Item(iRow).iNo & "99" Then
                  .Selection.TypeText Text:=m_Item(iRow).IDescHead
                  bolNewRow = True
               '¨ä¥L
               Else
                  .Selection.TypeText Text:=m_Item(iRow).IDesc
               End If
               .Selection.MoveRight Unit:=wdCell, Count:=1
               .Selection.TypeText Text:=m_Item(iRow).IChiA1k01
               If Right(m_Item(iRow).iNo, 2) = "99" Then
                  .Selection.MoveRight Unit:=wdCell, Count:=1 '¦^¨ì³W¶O
               Else
                  .Selection.MoveRight Unit:=wdCell, Count:=2 '¦^¨ìªA°È¶O
               End If
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
               .Selection.Font.Bold = False
               .Selection.TypeText Text:=Format(m_Item(iRow).iAmt)
               '³æµ§-½Ð´Úª÷ÃB
               If iRow = UBound(m_Item) Then
                  strExc(5) = ""
               Else
                  strExc(5) = m_Item(iRow + 1).IChiCno
               End If
               If iRow = UBound(m_Item) Or strExc(5) <> m_Item(iRow).IChiCno Then  '³Ì«á¤@µ§©M¤£¦P½Ð´Ú³æ
                  If m_Item(iRow).IChiAmt <> "" Then
                     'Added by Lydia 2020/09/08 ¥u¦³³W¶O¨S¦³ªA°È¶O,ªA°È¶O=ªÅ¥Õ
                     If Right(m_Item(iRow).iNo, 2) = "99" Then '
                        .Selection.MoveRight Unit:=wdCell, Count:=1
                        .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                        .Selection.Font.Bold = False
                     End If
                     'end 2020/09/08
                     .Selection.MoveRight Unit:=wdCell, Count:=1
                     .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                     .Selection.Font.Bold = False
                     .Selection.TypeText Text:=Format(m_Item(iRow).IChiAmt)
                     dbChiAmt = dbChiAmt + Val(m_Item(iRow).IChiAmt)
                  End If
                  If mColNum > 9 Then  '½Ð´Úª÷ÃB(USD)
                     .Selection.MoveRight Unit:=wdCell, Count:=1
                     .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                     .Selection.Font.Bold = False
                     .Selection.TypeText Text:=Format(m_Item(iRow).IChiUAmt)
                     dbChiUAmt = dbChiUAmt + Val(m_Item(iRow).IChiUAmt)
                     UpdateA1K38 m_Item(iRow).IChiA1k01, Format(m_Item(iRow).IChiUAmt) 'Added by Morgan 2021/1/14 §ó·s½Ð´Ú³æ¬üª÷Á`ÃB
                  End If
               '¦P½Ð´Ú³æ,¦X¨Ö¶µ¥Ø©MµL¦X¨Ö¶µ¥Ø
               Else
                 If (Right(m_Item(iRow).iNo, 2) = "99" And m_Item(iRow - 1).iNo & "99" = m_Item(iRow).iNo) Or _
                     (Right(m_Item(iRow).iNo, 2) <> "99" And m_Item(iRow + 1).iNo <> m_Item(iRow).iNo & "99") Then
                     If m_Item(iRow).IChiAmt <> "" Then
                        .Selection.MoveRight Unit:=wdCell, Count:=1
                        .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                        .Selection.Font.Bold = False
                        .Selection.TypeText Text:=Format(m_Item(iRow).IChiAmt)
                        dbChiAmt = dbChiAmt + Val(m_Item(iRow).IChiAmt)
                     End If
                     If mColNum > 9 Then
                        .Selection.MoveRight Unit:=wdCell, Count:=1
                        .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                        .Selection.Font.Bold = False
                        .Selection.TypeText Text:=Format(m_Item(iRow).IChiUAmt)
                        dbChiUAmt = dbChiUAmt + Val(m_Item(iRow).IChiUAmt)
                        UpdateA1K38 m_Item(iRow).IChiA1k01, Format(m_Item(iRow).IChiUAmt) 'Added by Morgan 2021/1/14 §ó·s½Ð´Ú³æ¬üª÷Á`ÃB
                     End If
                  End If
               End If
               .Selection.MoveRight Unit:=wdCell, Count:=1
            Else
            '---------------------------------------
                '¶µ¥Ø¥Îµê½u¤À¹j
                If iRow > 1 Then
                   .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleDashLargeGap
                   .Selection.Cells.Borders(wdBorderTop).LineWidth = wdLineWidth050pt
                End If
               .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(11.7), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustProportional
               .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
               .Selection.Collapse Direction:=wdCollapseStart
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
               
               bolNewRow = False
               '³Ì«á¤@¶µ
               If iRow = UBound(m_Item) Then
                  .Selection.TypeText Text:=m_Item(iRow).IDesc
               '«á­±±µ¬Û¦P½Ð´Ú¶µ¥Øªº³W¶O
               ElseIf m_Item(iRow + 1).iNo = m_Item(iRow).iNo & "99" Then
                  .Selection.TypeText Text:=m_Item(iRow).IDescHead
                  bolNewRow = True
               '¨ä¥L
               Else
                  .Selection.TypeText Text:=m_Item(iRow).IDesc
               End If
               
               If Right(m_Item(iRow).iNo, 2) = "99" Then
                  .Selection.MoveRight Unit:=wdCell, Count:=1
               Else
                  .Selection.MoveRight Unit:=wdCell, Count:=2
               End If
               
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
               .Selection.Font.Bold = False
               .Selection.TypeText Text:=m_Item(iRow).iAmt
               
                If Right(m_Item(iRow).iNo, 2) = "99" Then
                   .Selection.MoveRight Unit:=wdCharacter, Count:=2
                Else
                   .Selection.MoveRight Unit:=wdCharacter, Count:=1
                End If
                
                 If bolNewRow And m_Item(iRow).IDescTail <> "" Then
                   .Selection.InsertRows 1
                   .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleNone
                   .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=True
                   .Selection.Cells(1).Borders(wdBorderRight).LineStyle = wdLineStyleNone
                   .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.6), RulerStyle:=wdAdjustFirstColumn
                   .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(11.1), RulerStyle:=wdAdjustFirstColumn
                   .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustFirstColumn
                   .Selection.Collapse Direction:=wdCollapseStart
                   .Selection.MoveRight Unit:=wdCharacter, Count:=1
                   strText = "Attorney fee:" & m_Item(iRow).IDescTail
                   .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
                   .Selection.TypeText Text:=strText
                   .Selection.MoveRight Unit:=wdCell, Count:=2
                   .Selection.MoveRight Unit:=wdCharacter, Count:=1
                End If
            End If
         End If
         'Added by Lydia 2015/04/09
         If m_bolChiDB Then
            'Modified by Lydia 2020/09/08 «D³W¶O¤~+99
            'strLstNo = m_Item(iRow).IChiCno & m_Item(iRow).iNo
            strLstNo = m_Item(iRow).IChiCno & IIf(Right(m_Item(iRow).iNo, 2) = "99", Mid(m_Item(iRow).iNo, 1, Len(m_Item(iRow).iNo) - 2), m_Item(iRow).iNo)
         Else
            strLstNo = m_Item(iRow).iNo
         End If
      Next iRow
      
      'Added by Lydia 2015/04/08 +§PÂ_¬O§_¬°¾ã§å½Ð´Ú³æ(¤£¦L¤p­p)
      If m_bolChiDB Then
        .Selection.SelectRow
        .Selection.Collapse Direction:=wdCollapseStart
        .Selection.MoveRight Unit:=wdCharacter, Count:=7
        .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        .Selection.TypeText Text:="Á`­p"
        .Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
         strText = Format(m_Sum(3, 1)) 'Á`­p(°}¦CÀx¦s,¹ô§O+ª÷ÃB)
        .Selection.TypeText Text:=strText
         If m_iPrintCurrType = 2 Or m_iPrintCurrType = 4 Then
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
            '§ï¥Î¼È¦sÅÜ¼Æ
            .Selection.TypeText Text:=Format(dbChiUAmt, "###0")
         End If
          '´¡¤J¤U¤@¦C(1®æ),¥u¦³¤U¤è®æ½u,Ãþ¦üªí®æ¦b³Ì¥½ºÝµe½u
          .Selection.Font.Bold = False
          .Selection.MoveRight Unit:=wdCharacter, Count:=3
          .Selection.InsertRows 1
          .Selection.Cells.Merge
          
      Else
          '¤p­p
          .Selection.InsertRows 1
          .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
          .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleDouble
          .Selection.Cells.Borders(wdBorderTop).LineWidth = wdLineWidth050pt
          .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
          .Selection.Cells(1).SetHeight RowHeight:=30, HeightRule:=wdRowHeightAtLeast
          .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(11.7), RulerStyle:=wdAdjustProportional
          .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.9), RulerStyle:=wdAdjustProportional
          .Selection.Collapse Direction:=wdCollapseStart
          
          .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
          .Selection.TypeText Text:="¤p¡@­p"
          
          '³W¶O
          .Selection.MoveRight Unit:=wdCharacter, Count:=1
          .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
          If dblOffFeeSub > 0 Then
             strText = Format(dblOffFeeSub, "#,##0.00")
             dblA1K40 = dblOffFeeSub
             .Selection.TypeText Text:=strText
             If m_iPrintCurrType = 2 Or m_iPrintCurrType = 4 Then
                .Selection.TypeParagraph
                If m_iPrintCurrType = 2 Then
                   dblA1K40 = Trunc(dblOffFeeSub / Val(m_DNRate))
                   strText = "(" & m_DNCurr & Format(dblA1K40, "#,##0.00") & ")"
                Else
'                   dblA1K40 = Trunc(dblOffFeeSub * Val(m_DUsdRate))
                   strText = "(USD" & Format(Trunc(dblOffFeeSub * Val(m_DUsdRate)), "#,##0.00") & ")"
                End If
                .Selection.TypeText Text:=strText
             End If
             
             'Add By Sindy 2021/10/26 °O¿ý ½Ð´Ú³W¶O¥~¹ôª÷ÃB
             strSql = "Update ACC1K0 Set A1K40=" & dblA1K40 & " Where A1K01='" & m_strDN & "'"
             adoTaie.Execute strSql, intI
             '2021/10/26 END
          Else
             'Add By Sindy 2021/12/6 °O¿ý ½Ð´Ú³W¶O¥~¹ôª÷ÃB
             strSql = "Update ACC1K0 Set A1K40=" & dblOffFeeSub & " Where A1K01='" & m_strDN & "'"
             adoTaie.Execute strSql, intI
             '2021/12/6 END
          End If
          
          'ªA°È¶O
          .Selection.MoveRight Unit:=wdCharacter, Count:=1
          .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
          If dblAttFeeSub > 0 Then
             strText = Format(dblAttFeeSub, "#,##0.00")
             dblA1K39 = dblAttFeeSub
             .Selection.TypeText Text:=strText '¤p­pªA°È¶O
             
             If m_iPrintCurrType = 2 Or m_iPrintCurrType = 4 Then
                .Selection.TypeParagraph
                If m_iPrintCurrType = 2 Then
                   dblA1K39 = Trunc(dblAttFeeSub / Val(m_DNRate))
                   strText = "(" & m_DNCurr & Format(dblA1K39, "#,##0.00") & ")"
                Else
'                   dblA1K39 = Trunc(dblAttFeeSub * Val(m_DUsdRate))
                   strText = "(USD" & Format(Trunc(dblAttFeeSub * Val(m_DUsdRate)), "#,##0.00") & ")"
                End If
                .Selection.TypeText Text:=strText
             End If
             
             'Add By Sindy 2021/10/26 °O¿ý ½Ð´ÚªA°È¶O¥~¹ôª÷ÃB
             strSql = "Update ACC1K0 Set A1K39=" & dblA1K39 & " Where A1K01='" & m_strDN & "'"
             adoTaie.Execute strSql, intI
             '2021/10/26 END
          Else
             'Add By Sindy 2021/12/6 °O¿ý ½Ð´ÚªA°È¶O¥~¹ôª÷ÃB
             strSql = "Update ACC1K0 Set A1K39=" & dblAttFeeSub & " Where A1K01='" & m_strDN & "'"
             adoTaie.Execute strSql, intI
             '2021/12/6 END
          End If
          
          .Selection.MoveRight Unit:=wdCharacter, Count:=2
          
         'Added by Morgan 2021/5/4 Y54810·¸®õ(«n¨Ê)¥[¦Lµ|ª÷¤Î§tµ|Á`­p
         'Modified by Morgan 2021/5/6 +X83843000 Ex:CFP-032375--³¢
         If m_A1k28 = "Y54810000" Or m_A1k28 = "X83843000" Then
            .Selection.SelectRow
            .Selection.InsertRows 1
            .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
            .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleDouble
            .Selection.Cells.Borders(wdBorderTop).LineWidth = wdLineWidth050pt
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(11.7), RulerStyle:=wdAdjustProportional
            .Selection.Cells(1).SetHeight RowHeight:=30, HeightRule:=wdRowHeightAtLeast
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.Font.Size = 12
            .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Selection.TypeText Text:="¤p¡@­p¡@¥[¡@Á`"
            .Selection.Font.Size = strFontSize
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            strText = m_Sum(3, 1)
            .Selection.TypeText Text:=strText
            .Selection.MoveRight Unit:=wdCharacter, Count:=2
            
            .Selection.SelectRow
            .Selection.InsertRows 1
            .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
            .Selection.Cells.Borders(wdBorderTop).LineStyle = wdLineStyleDouble
            .Selection.Cells.Borders(wdBorderTop).LineWidth = wdLineWidth050pt
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(11.7), RulerStyle:=wdAdjustProportional
            .Selection.Cells(1).SetHeight RowHeight:=30, HeightRule:=wdRowHeightAtLeast
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.Font.Size = 12
            .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Selection.TypeText Text:="µ|¡@ª÷"
            .Selection.Font.Size = strFontSize
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            'Modified by Morgan 2021/9/23 «È¤á­n¨D¶O¥Î§ïª½±µ¡Ñ1.06´N¦n¡AµL¶·¡Ò0.992--¼ïÃý¥à
            'strText = Round(Val(Format(m_Sum(3, 1))) / 0.9928 * 1.06 - Val(Format(m_Sum(3, 1))), 2)
            strText = Round(Val(Format(m_Sum(3, 1))) * 1.06 - Val(Format(m_Sum(3, 1))), 2)
            'end 2021/9/23
            .Selection.TypeText Text:=strText
            .Selection.MoveRight Unit:=wdCharacter, Count:=2
         End If
         'end 2021/5/4
         
         '¦X­p
         .Selection.SelectRow
         .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(11.7), RulerStyle:=wdAdjustProportional
         .Selection.Cells(1).SetHeight RowHeight:=30, HeightRule:=wdRowHeightAtLeast
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.Font.Size = 12
         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.Font.Bold = True
         .Selection.TypeText Text:="Á`¡@­p"
         .Selection.Font.Size = strFontSize
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.Font.Bold = True
         'Added by Morgan 2021/5/4 Y54810·¸®õ(«n¨Ê)¥[¦Lµ|ª÷¤Î§tµ|Á`­p
         'Modified by Morgan 2021/5/6 +X83843000 Ex:CFP-032375--³¢
         If m_A1k28 = "Y54810000" Or m_A1k28 = "X83843000" Then
            'Modified by Morgan 2021/9/23 «È¤á­n¨D¶O¥Î§ïª½±µ¡Ñ1.06´N¦n¡AµL¶·¡Ò0.992--¼ïÃý¥à
            'strText = m_Sum(2, 1) & Round(Val(Format(m_Sum(3, 1))) / 0.9928 * 1.06, 2)
            strText = m_Sum(2, 1) & Round(Val(Format(m_Sum(3, 1))) * 1.06, 2)
            'end 2021/9/23
         Else
         'end 2021/5/4
            strText = m_Sum(2, 1) & m_Sum(3, 1)
         End If
         .Selection.TypeText Text:=strText
          
         If m_iPrintCurrType = 2 Or m_iPrintCurrType = 4 Then
             .Selection.TypeParagraph
             If m_iPrintCurrType = 2 Then
                strText = m_DNCurr & Format(Trunc(dblOffFeeSub / Val(m_DNRate)) + Trunc(dblAttFeeSub / Val(m_DNRate)), "#,##0.00")
             Else
               strText = "USD" & Format(Trunc(dblOffFeeSub * Val(m_DUsdRate)) + Trunc(dblAttFeeSub * Val(m_DUsdRate)), "#,##0.00")
                
               UpdateA1K38 m_strDN, Trunc(dblOffFeeSub * Val(m_DUsdRate)) + Trunc(dblAttFeeSub * Val(m_DUsdRate)) 'Added by Morgan 2021/1/14 §ó·s½Ð´Ú³æ¬üª÷Á`ÃB
             End If
             .Selection.TypeText Text:=strText
         End If

         .Selection.Font.Bold = False
         .Selection.MoveRight Unit:=wdCharacter, Count:=3
         .Selection.InsertRows 1
         .Selection.Cells.Merge
      End If
      
      '¥Ø«e¥u¦³®æ¦¡2(¥x¹ô+¥~¹ô¦X­p)
      If m_iPrintCurrType = 2 Then
      '¤¤¤å¨S¦³¦L¶×²v¤£¥²¦L»¡©ú
      End If
      
      'ªí§À1
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.Collapse Direction:=wdCollapseStart
      With .ActiveDocument.Bookmarks
         .add Range:=g_WordAp.Application.Selection.Range, Name:="BreakPos"
         .DefaultSorting = wdSortByLocation
         .ShowHidden = False
      End With
      .Selection.TypeText Text:=m_Footer(1, 1) 'Name of Bank: Bank of Taiwan, ~½ã¸¹¡G
      
      If m_Footer(2, 1) <> "" Then
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         .Selection.Cells.Split NumRows:=3, NumColumns:=2, MergeBeforeSplit:=False
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.Cells(1).SetHeight RowHeight:=36, HeightRule:=wdRowHeightAtLeast
         .Selection.MoveDown Unit:=wdLine, Count:=1
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
         With .Selection.Cells
             With .Borders(wdBorderLeft)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             With .Borders(wdBorderRight)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             With .Borders(wdBorderTop)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             With .Borders(wdBorderBottom)
                 .LineStyle = wdLineStyleSingle
                 .LineWidth = wdLineWidth100pt
                 .ColorIndex = wdAuto
             End With
             .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
             .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
             .Borders.Shadow = False
         End With
         .Selection.ParagraphFormat.LeftIndent = .CentimetersToPoints(0.2)
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
         .Selection.Cells(1).SetHeight RowHeight:=52, HeightRule:=wdRowHeightAtLeast
         .Selection.TypeText Text:=m_Footer(2, 1)
         .Selection.HomeKey
         .Selection.MoveDown Unit:=wdLine, Count:=1
         .Selection.Cells(1).SetHeight RowHeight:=0, HeightRule:=wdRowHeightAtLeast
         .Selection.MoveDown Unit:=wdLine, Count:=1
      Else
         .Selection.MoveRight Unit:=wdCharacter, Count:=3
      End If
      
      '³Æµù
      If UBound(m_Footer, 2) > 1 Then
         .Selection.SelectRow
         .Selection.Font.Bold = True
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
         .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=True
         .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
         .Selection.Collapse Direction:=wdCollapseStart
         .Selection.TypeText Text:="¡°"
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         iPos1 = InStr(m_Footer(1, 2), "±©¤_¶×´Ú¦Z")
         If iPos1 > 0 Then
            iPos2 = InStr(iPos1, m_Footer(1, 2), "¡C")
         End If
         If iPos1 > 0 And iPos2 > 0 Then
            .Selection.TypeText Text:=Left(m_Footer(1, 2), iPos1 - 1)
            .Selection.Font.Underline = wdUnderlineDouble
            .Selection.TypeText Text:=Mid(m_Footer(1, 2), iPos1, iPos2 - iPos1)
            .Selection.Font.Underline = wdUnderlineNone '±©¤_¶×´Ú«á->µe©³½u
            .Selection.TypeText Text:=Mid(m_Footer(1, 2), iPos2)
         Else
            .Selection.TypeText Text:=m_Footer(1, 2)
         End If
         
         For iRow = 3 To UBound(m_Footer, 2)
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.InsertRows 1
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:=m_Footer(1, iRow)
         Next
         .Selection.Font.Bold = False
      End If
      
      .Selection.WholeStory
      .Selection.Font.Name = "Times New Roman"
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .ActiveDocument.Repaginate
      '¶W¹L1­¶®É´¡¤J­¶½X
      If .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) > 1 Then
         .Selection.GoTo what:=wdGoToBookmark, Name:="BreakPos"
         .Selection.InsertBreak Type:=wdPageBreak
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .ActiveDocument.Repaginate
         If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
            .ActiveWindow.ActivePane.View.Type = wdPageView
         Else
            .ActiveWindow.View.Type = wdPageView
         End If
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
         .Selection.TypeParagraph
         .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldPage
         .Selection.TypeText Text:="/"
         .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldEmpty, Text:="NUMPAGES ", PreserveFormatting:=True
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
         
      End If
      .ActiveDocument.Bookmarks("BreakPos").Delete
      .Selection.HomeKey Unit:=wdStory
      
      'Added by Morgan 2013/12/12 ¤º±M³£¥Î¥Õ¯È¦L
      'Modify By Sindy 2015/7/13 ¶®®S¸ò¨q¬Â»¡­n¥Î±M§Qªk«ß«HÀY
      'If Pub_StrUserSt03 = "P12" Then
      'Remove by Lydia 2016/09/09 ´ú¸ÕX10508660,µo²{¦³­«ÂÐ«HÀY
'      If (m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") Then
'      '2015/7/13 END
'         '´¡¤J¹Ï¤ùÀÉ®×(¥Î¸õ­¶²Å¸¹§PÂ_­¶¼Æ)
'         If PUB_ReadDB2File(stFileName, 7) Then
'            .Selection.HomeKey Unit:=wdStory, Extend:=wdMove
'            Do
'               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
'               oShape.ZOrder 5
'               oShape.LockAnchor = True
'               oShape.LockAspectRatio = -1
'               oShape.Width = 546.5
'               oShape.WrapFormat.Type = wdWrapNone
'               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
'               oShape.Left = .CentimetersToPoints(1)
'               oShape.Top = .CentimetersToPoints(1)
'               iHeadCount = iHeadCount + 1
'               .ActiveDocument.Repaginate
'               '¨S¦³¸õ­¶²Å¸¹¦ý­¶¼Æ¤j©ó¤w¦L«HÀY¼Æ
'               If iHeadCount < .ActiveDocument.BuiltinDocumentProperties(wdPropertyPages) Then
'                  .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
'               Else
'                  Exit Do
'               End If
'            Loop
'         End If
'      End If
      'end 2013/12/12
      
      'Add By Sindy 2015/7/9 +if ±M§Q³B§ï¦b¤U­±¥[«HÀY«á¤~¦C¦L
      'Modified by Morgan 2016/8/5 +FCP,FG  --Kimi
      'Removed by Morgan 2020/4/8 °Ó¼Ð¤]­n±a«HÀY--®Û­^
      'Modified by Morgan 2020/4/23 ¤º°Ó§ï¦^¤£­n«HÀY--´ðªå
      'If Not (m_CP01 = "FCP" Or m_CP01 = "FG" Or m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") Then
      'Modified by Lydia 2020/09/08 ¾ã§å½Ð´Ú³æ¤£­n«HÀY
      'If Left(Pub_StrUserSt03, 2) = "P2" Then
      'Removed by Morgan 2025/8/19 ¤º°Ó§ï³æµ§©Î¾ã§å³£­n«HÀY--®Û­^
      'If Left(Pub_StrUserSt03, 2) = "P2" Or (m_bolChiDB = True And bol_ChiDB = True) Then
      '   'Added by Lydia 2015/04/10 §PÂ_¬O§_¥u¿é¥Xword (m_Chi2Word)
      '   'If m_bPrintWord And m_iSpCopies > 0 Then
      '   If m_bPrintWord And m_iSpCopies > 0 Then 'And m_Chi2Word = False Then
      '      .ActiveDocument.PrintOut Background:=False, Copies:=m_iSpCopies, Collate:=True
      '      m_iPageCount = .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) 'Added by Morgan 2015/2/26
      '   End If
      'End If
      'end 2025/8/19
      'end 2020/4/23
      'end 2020/4/8
      
      'Modify by Morgan 2011/7/15 ¤º±M§ï­n±a«HÀY--³¢
      'Modified by Morgan 2013/12/12 ¤º±M±±¨î²¾¨ì¤W­±
      'Modified by Morgan 2014/2/26 +¥~°Ó³£±a«HÀY
      'Modified by Morgan 2014/8/25 §ï¤Á´«¦Lªí¾÷¤è¦¡¦C¦LPDF,¤£¥Î­«¶]Word
      'If m_bWord2Pdf Or m_bSaveWord Or (m_bEditDoc And Left(Pub_StrUserSt03, 2) = "F1") Then
       'Added by Lydia 2015/04/10 §PÂ_¬O§_¥u¿é¥Xword (m_Chi2Word)
      'If m_b2PDF Or m_bWord2Pdf Or m_bSaveWord Or (m_bEditDoc And Left(Pub_StrUserSt03, 2) = "F1") Then
      'Modify By Sindy 2015/7/9 + Or (m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS")
      'Modified by Morgan 2016/8/5 +FCP,FG ¤]­n«HÀY  --Kimi
      'Removed by Morgan 2020/4/8 °Ó¼Ð¤]­n±a«HÀY--®Û­^
      'If m_Chi2Word Or m_b2PDF Or m_bWord2Pdf Or m_bSaveWord Or (m_bEditDoc And Left(Pub_StrUserSt03, 2) = "F1") Or (m_CP01 = "FCP" Or m_CP01 = "FG" Or m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") Then
      'Added by Morgan 2020/4/23 ¤º°Ó§ï¦^¤£­n«HÀY--´ðªå
      'Modified by Lydia 2020/09/08 ¾ã§å½Ð´Ú³æ¤£­n«HÀY
      'If m_b2PDF Or m_bWord2Pdf Or m_bSaveWord Or Left(Pub_StrUserSt03, 2) <> "P2" Then
      'end 2020/4/23
      'Removed by Morgan 2025/8/19 ¤º°Ó§ï³æµ§©Î¾ã§å³£­n«HÀY --®Û­^
      'If m_b2PDF Or m_bWord2Pdf Or m_bSaveWord Or Not (Left(Pub_StrUserSt03, 2) = "P2" Or (m_bolChiDB = True And bol_ChiDB = True)) Then
      'end 2025/8/19
         'Added by Morgan 2020/3/31
         If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é Then
            PUB_GetLetterPicID "2", m_CP01, iPicNo, iPicNo2, 1, True, Pub_StrUserSt03
            If PUB_ReadDB2File(stFileName, iPicNo) Then
               For ii = 1 To .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) '¨C­¶³£­n¦³«HÀY§À
                  Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                  oShape.ZOrder 4
                  oShape.LockAnchor = True
                  oShape.LockAspectRatio = -1
                  oShape.Width = .CentimetersToPoints(21)
                  oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                  oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                  oShape.Left = .CentimetersToPoints(0)
                  oShape.Top = .CentimetersToPoints(0)
                  oShape.WrapFormat.Type = wdWrapNone
                  .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
               Next ii
               .Selection.HomeKey Unit:=wdStory
               If iPicNo2 > 0 Then
                  If PUB_ReadDB2File(stFileName, iPicNo2) Then
                     For ii = 1 To .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)
                        Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                        oShape.ZOrder 4
                        oShape.LockAnchor = True
                        oShape.LockAspectRatio = -1
                        oShape.Width = .CentimetersToPoints(21)
                        oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                        oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                        oShape.Left = .CentimetersToPoints(0)
                        oShape.Top = .CentimetersToPoints(27.6)
                        oShape.WrapFormat.Type = wdWrapNone
                        .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
                     Next ii
                  End If
                  .Selection.HomeKey Unit:=wdStory
               End If
            End If
         Else
         'end 2020/3/31
      
            'Modify By Sindy 2015/7/17
            'If Pub_StrUserSt03 = "P12" Then
            If (m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") Then
            '2015/7/17 END
               '´¡¤J¹Ï¤ùÀÉ®×(¥Î¸õ­¶²Å¸¹§PÂ_­¶¼Æ)
               If PUB_ReadDB2File(stFileName, 19) Then
                  .Selection.HomeKey Unit:=wdStory, Extend:=wdMove
                  Do
                     Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                     oShape.ZOrder 5
                     oShape.LockAnchor = True
                     oShape.LockAspectRatio = -1
                     oShape.Width = 546.5
                     oShape.WrapFormat.Type = wdWrapNone
                     oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                     oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                     oShape.Left = .CentimetersToPoints(1)
                     oShape.Top = .CentimetersToPoints(1)
                     iHeadCount = iHeadCount + 1
                     .ActiveDocument.Repaginate
                     '¨S¦³¸õ­¶²Å¸¹¦ý­¶¼Æ¤j©ó¤w¦L«HÀY¼Æ
                     If iHeadCount < .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) Then
                        .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
                     Else
                        Exit Do
                     End If
                  Loop
               End If
            'Added by Morgan 2016/8/5  --Kimi
            'FCP¦L­^¤å«HÀY
            ElseIf m_CP01 = "FCP" Or m_CP01 = "FG" Then
               
               If PUB_ReadDB2File(stFileName, 5) Then
                  For ii = 1 To .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) '¨C­¶³£­n¦³«HÀY§À
                     Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                     oShape.ZOrder 4
                     oShape.LockAnchor = True
                     oShape.LockAspectRatio = -1
                     oShape.Width = .CentimetersToPoints(21)
                     oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                     oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                     oShape.Left = .CentimetersToPoints(0)
                     '¥~±M­n¥Î¶}µ¡«H«Ê
                     oShape.Top = .CentimetersToPoints(0)
                     oShape.WrapFormat.Type = wdWrapNone
                     .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
                  Next ii
                  .Selection.HomeKey Unit:=wdStory
                  
                  If PUB_ReadDB2File(stFileName, 9) Then
                     For ii = 1 To .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)
                        Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                        oShape.ZOrder 4
                        oShape.LockAnchor = True
                        oShape.LockAspectRatio = -1
                        oShape.Width = .CentimetersToPoints(21)
                        oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                        oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                        oShape.Left = .CentimetersToPoints(0)
                        oShape.Top = .CentimetersToPoints(27)
                        oShape.WrapFormat.Type = wdWrapNone
                        .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
                     Next ii
                  End If
                  .Selection.HomeKey Unit:=wdStory
               End If
            'end 2016/8/5
            Else
            
               '´¡¤J¹Ï¤ùÀÉ®×(¥Î¸õ­¶²Å¸¹§PÂ_­¶¼Æ)
               If PUB_ReadDB2File(stFileName, 7) Then
                  .Selection.HomeKey Unit:=wdStory, Extend:=wdMove
                  Do
                     Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                     oShape.ZOrder 5
                     oShape.LockAnchor = True
                     oShape.LockAspectRatio = -1
                     oShape.Width = 546.5
                     oShape.WrapFormat.Type = wdWrapNone
                     oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                     oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                     oShape.Left = .CentimetersToPoints(1)
                     oShape.Top = .CentimetersToPoints(1)
                     iHeadCount = iHeadCount + 1
                     .ActiveDocument.Repaginate
                     '¨S¦³¸õ­¶²Å¸¹¦ý­¶¼Æ¤j©ó¤w¦L«HÀY¼Æ
                     If iHeadCount < .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) Then
                        .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
                     Else
                        Exit Do
                     End If
                  Loop
               End If
            End If 'Added by Morgan 2020/3/31
            
         End If
         
         If m_bWord2Pdf Then
            .ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
            
         'Added by Morgan 2014/8/26
         '§ï¤Á´«¦Lªí¾÷¤è¦¡¦C¦LPDF,¤£¥Î­«¶]Word
         ElseIf m_b2PDF Then
            PrintWord2PDF
         'end 2014/8/26
         
         End If
         
         'Move by Lydia 2020/12/25 ±qIf m_bWord2Pdf Then ¤W­±²¾¤U¨Ó; ¦]¬°12/23ªºFCP¦~¶O¾ã§åµo¤å¨ÌÂÂ¦³¯È¥»PDF¥¼¦s¤JTyping2
         'Add By Sindy 2015/7/9 ±M§Q³B§ï¦b¦¹³B¤~¦C¦L
         'Modified by Morgan 2016/8/5 +FCP,FG  --Kimi
         'Modified by Morgan 2020/4/23 ¤º°Ó§ï¦^¤£­n«HÀY--´ðªå
         'If (m_CP01 = "FCP" Or m_CP01 = "FG" Or m_CP01 = "P" Or m_CP01 = "PS" Or m_CP01 = "CFP" Or m_CP01 = "CPS") Then 'Removed by Morgan 2020/4/8 °Ó¼Ð¤]­n±a«HÀY--®Û­^
         'Removed by Morgan 2025/8/19 ¤º°Ó§ï³æµ§©Î¾ã§å³£­n«HÀY --®Û­^
         'If Left(Pub_StrUserSt03, 2) <> "P2" Then
         'end 2025/8/19
         'end 2020/4/23
            
            If m_bPrintWord And m_iSpCopies > 0 Then 'And m_Chi2Word = False Then
               .ActiveDocument.PrintOut Background:=False, Copies:=m_iSpCopies, Collate:=True
               m_iPageCount = .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) 'Added by Morgan 2015/2/26
            End If
            
         'End If
         '2015/7/9 END
         'end -- 2020/12/25
         
         If m_bSaveWord Then
            RidFile m_EFilePath
            .ActiveDocument.SaveAs m_EFilePath
         End If

      'End If Removed by Morgan 2025/8/19
            
   End With
   'Modified by Lydia 2019/04/09 §ï¦¨¦@¥Î¼Ò²Õ
   'RePosWord bVisible, m_WordLeft, m_WordTop 'Added by Morgan 2014/6/26
   Pub_RePosWord g_WordAp, bVisible, m_WordLeft, m_WordTop
   
      If m_bEditDoc Then
         g_WordAp.Visible = True
         g_WordAp.Activate
      Else
         g_WordAp.ActiveDocument.Close wdDoNotSaveChanges
         If bVisible = False Then
            g_WordAp.Quit wdDoNotSaveChanges
            Set g_WordAp = Nothing 'Added by Lydia 2017/12/12 Á×§K§Ö³t¶}±ÒWord,µ{¦¡¥X¿ù
         Else
            g_WordAp.Visible = True
         End If
      End If

   
   Exit Sub
   
ErrHnd:
   If Err.Number <> 0 Then
'Modified by Morgan 2014/6/26
'      If iResumeCnt > 3 Then
'         MsgBox "¿ù»~ : " & Err.Description, vbCritical
'      Else
'         iResumeCnt = iResumeCnt + 1
'         Select Case Err.NUMBER
'            Case 91:
'               g_WordAp.Documents.Add
'               Resume Next
'            Case 462:
'               Set g_WordAp = New Word.Application
'               Resume
'            Case Else:
'               MsgBox "¿ù»~ : " & Err.Description, vbCritical
'         End Select
'      End If
      MsgBox "¿ù»~ : " & Err.Description, vbCritical
'end 2014/6/26
   End If
End Sub
'Added by Morgan 2014/8/21 ¯S®í¦C¦L¹ï¶H¥N½X
Private Function GetPlusFormNo(pA1K28 As String, pCNo1 As String, pCNo2 As String, pCNo3 As String, pCNo4 As String) As String
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   
   'Modified by Morgan 2018/7/27 ¶Ê´Ú¤£¥²¦L--°û²ñ,Kimi
   'If pA1K28 = "Y51333010" Then
   'Modified by Lydia 2020/01/06 §PÂ_©I¥s½Ð´Ú³æªºªí³æ¦WºÙ¡A¥u­­¨î¶Ê´Ú³æ¨Óªº¦C¦L¤£¥Î¥[¦L¯S®í½Ð´Ú³æ
   'If pA1K28 = "Y51333010" And Not m_bBeCalled Then
   'end 2018/7/27
   'Memo by Lydia 2020/01/06 Y51333010(»ÈÀs)
   If pA1K28 = "Y51333010" And m_CallPrevForm <> "Frmacc2470" Then
   
      stSQL = "select pa166 from patent where pa01='" & pCNo1 & "' and pa02='" & pCNo2 & "' and pa03='" & pCNo3 & "' and pa04='" & pCNo4 & "' and pa166 is not null"
      intR = 1
      Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         GetPlusFormNo = rsQuery(0)
      End If
      Set rsQuery = Nothing
   End If
End Function

'Added by Morgan 2014/8/26
'ÂàPDF
Private Sub PrintWord2PDF(Optional pBolPlusForm As Boolean)
   Dim strFolder As String
   Dim strFileName As String
   
   'Added by Morgan 2025/8/20
   '¤º°Óµ{§Ç¾Þ§@®É¤@«ß¥u¦sPDFÀÉ¦Ü¤½¥Î¹q¸£
   'ÀÉ¦W=½Ð´Ú³æ¸¹.invoice.pdf
   'Modified by Morgan 2025/8/27 m_SavePath¤w³]©w®É°£¥~(¶Ê´Ú³æ©I¥s)
   If Left(Pub_StrUserSt03, 2) = "P2" And m_SavePath = "" Then
      If (UCase(pub_DbTerminalName) <> UCase(¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ) Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0) Then
         m_EFilePath = PUB_Getdesktop
      Else
         m_EFilePath = Pub_GetSpecMan("MCTInvoicePath")
      End If
      strFolder = m_EFilePath
      strFileName = m_strDN & ".invoice"
   Else
   'end 2025/8/20
      
      m_EFilePath = GetPath
      strFolder = m_EFilePath
      If Dir(strFolder, vbDirectory) = "" Then
         MkDir strFolder
      End If
      strFileName = m_strCaseNo & IIf(m_bAddDate, "_" & strSrvDate(1), "") & "_DN" & m_strDN & IIf(pBolPlusForm, "_2", "")
      
   End If 'Added by Morgan 2025/8/19
   
   Me.Tag = strFolder & "\" & strFileName & ".pdf" 'Added by Lydia 2017/02/18 °O¿ý-½Ð´Ú³æpdfÀÉ¸ô®|
   EfileNameFCP_08 = EfileNameFCP_08 & ";" & m_EFilePath & "\" & strFileName & ".pdf" 'Add By Sindy 2018/1/24
   
   'Added by Morgan 2017/9/15
   If pub_Word2Pdf Then
      g_WordAp.ActiveDocument.ExportAsFixedFormat OutputFileName:=Me.Tag, ExportFormat:=17, OpenAfterExport:=False
      'Added by Lydia 2020/02/15 §PÂ_ÀÉ®×¬O§_¦s¦b, ¶W¹L®É¶¡´NÄ~Äò
      'Modified by Lydia 2020/09/10 ¶W¹L®É¶¡¡Aª½±µ°O¿ý¥¢±Ñ²M³æ
      'If PUB_ChkFileStatus(Me.Tag) = False Then
      If PUB_ChkFileStatus(Me.Tag, False, m_strOutErr) = False Then
      End If
      'end 2020/02/15
   Else
   'end 2017/9/15
      frmPDF.Show
      frmPDF.StartProcess strFolder, strFileName
      '¤Á´«¦Lªí¾÷
      If PUB_PdfCreatorNameInWord = "" Then PUB_PdfCreatorNameInWord = PUB_GetCreatorNameInWord 'Added by Morgan 2017/10/6
      g_WordAp.ActivePrinter = PUB_PdfCreatorNameInWord
      g_WordAp.ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
      frmPDF.EndtProcess
      'Added by Lydia 2020/02/15 §PÂ_ÀÉ®×¬O§_¦s¦b, ¶W¹L®É¶¡´NÄ~Äò
      'Modified by Lydia 2020/09/10 ¶W¹L®É¶¡¡Aª½±µ°O¿ý¥¢±Ñ²M³æ
      'If PUB_ChkFileStatus(Me.Tag) = False Then
      If PUB_ChkFileStatus(Me.Tag, False, m_strOutErr) = False Then
      End If
      'end 2020/02/15
      Unload frmPDF
   End If
   
   m_PdfDone = True
End Sub

'Added by Morgan 2014/8/26
'³]©w¯S®í¦C¦L¹ï¶H¸ê®Æ
Private Sub SetPlusFormData()
   Dim intHeadRow As Integer
   Dim arrNo() As String, ii As Integer
   
   If m_PlusFormNo = "" Then Exit Sub
   
   Erase m_PlusHead
   ReDim m_PlusHead(4, 1)
   Erase m_PlusFooter
   ReDim m_PlusFooter(2, 1)
            
   intHeadRow = 0
   '¯S®í¦C¦L¹ï¶H
   'Modified by Morgan 2015/6/10 +¦h­Ó¥Î³r¸¹¤À¹j
   strExc(0) = ""
   arrNo = Split(m_PlusFormNo, ",")
   For ii = LBound(arrNo) To UBound(arrNo)
      If Trim(arrNo(ii)) <> "" Then
         arrNo(ii) = RTrim(arrNo(ii))
         If Left(arrNo(ii), 1) = "Y" Then
            strExc(0) = strExc(0) & IIf(strExc(0) <> "", " union ", "") & "SELECT fa06, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20 FROM FAGENT WHERE FA01='" & Left(arrNo(ii), 8) & "' AND FA02='" & Mid(arrNo(ii), 9) & "'"
         Else
            strExc(0) = strExc(0) & IIf(strExc(0) <> "", " union ", "") & "SELECT cu06 as fa06, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20" & _
               " FROM customer WHERE CU01='" & Left(arrNo(ii), 8) & "' AND CU02='" & Mid(arrNo(ii), 9) & "'"
         End If
      End If
   Next
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         '¤é¤å¦WºÙ
         If IsNull(.Fields("fa06").Value) = False Then
             m_tmp = .Fields("fa06").Value & "¡@" & "±s¤¤"
             intHeadRow = intHeadRow + 1
             SetWordArray m_PlusHead, intHeadRow, 1, m_tmp
         '­^¤å¦WºÙ
         ElseIf IsNull(adoacc1k0.Fields("fa05").Value) = False Then
             intHeadRow = intHeadRow + 1
             SetWordArray m_PlusHead, intHeadRow, 2, .Fields("fa05").Value
             If IsNull(adoacc1k0.Fields("fa63").Value) = False Then
                intHeadRow = intHeadRow + 1
                SetWordArray m_PlusHead, intHeadRow, 1, .Fields("fa63").Value
             End If
             If IsNull(adoacc1k0.Fields("fa64").Value) = False Then
                intHeadRow = intHeadRow + 1
                SetWordArray m_PlusHead, intHeadRow, 1, .Fields("fa64").Value
             End If
             If IsNull(adoacc1k0.Fields("fa65").Value) = False Then
                intHeadRow = intHeadRow + 1
                SetWordArray m_PlusHead, intHeadRow, 1, .Fields("fa65").Value
             End If
         End If
         .MoveNext
      Loop
      End With
   End If
   'end 2015/6/10
   
   '½Ð´Ú¤é´Á
   m_tmp = Format(ChangeTStringToWString("" & m_strA1K02), "####¦~##¤ë##¤é")
   intHeadRow = intHeadRow + 1
   SetWordArray m_PlusHead, intHeadRow, 3, m_tmp
   'Modify By Sindy 2025/9/17 tm05 ==> nvl(tm131,tm05) ¦³©w½Z°Ó¼Ð¦WºÙ®É,«h§ì¦¹Äæ¦ì­ÈÅã¥Ü,µL®É,´N¥ÎTM05
   strExc(0) = "select pa77 as Yno, pa48 as Cno, ptm06 as MName, pa06 as Cname, pa11 as Ano, pa26 as Custno, pa22, null as tm15, null as Yes, '' As TM12, '' As TM16, '' As TM09, '' AS TM32,pa08,pa09 from patent, systemkind, patenttrademarkmap where pa01 = sk01 and pa08 = ptm02 (+) and sk02 = ptm01 and pa01 = '" & m_CP01 & "' and pa02 = '" & m_CP02 & "' and pa03 = '" & m_CP03 & "' and pa04 = '" & m_CP04 & "' union " & _
      "select tm45 as Yno, tm35 as Cno, ptm06 as MName, Rtrim(Ltrim(nvl(tm131,tm05)||' '||tm06)) as Cname, tm12 as Ano, tm23 as Custno, null as pa22, tm15, '1' as Yes, TM12, TM16, TM09, TM32,tm08 as pa08,tm10 as pa09 from trademark, systemkind, patenttrademarkmap where tm01 = sk01 and tm08 = ptm02 (+) and sk02 = ptm01 and tm01 = '" & m_CP01 & "' and tm02 = '" & m_CP02 & "' and tm03 = '" & m_CP03 & "' and tm04 = '" & m_CP04 & "' union " & _
      "select lc23 as Yno, lc17 as Cno, '' as MName, nvl(lc05,lc06) as Cname, '' as Ano, lc11 as Custno, null as pa22, null as tm15, null as Yes, '' As TM12, '' As TM16, '' As TM09, '' AS TM32,'' as pa08,'000' as pa09 from lawcase where lc01 = '" & m_CP01 & "' and lc02 = '" & m_CP02 & "' and lc03 = '" & m_CP03 & "' and lc04 = '" & m_CP04 & "' union " & _
      "select sp27 as Yno, sp29 as Cno, '' as MName, nvl(sp05,sp06) as Cname, sp11 as Ano, sp08 as Custno, null as pa22, null as tm15, null as Yes, '' As TM12, '' As TM16, sp73 As TM09, SP74 AS TM3,'' as pa08,sp09 as pa09 from servicepractice where sp01 = '" & m_CP01 & "' and sp02 = '" & m_CP02 & "' and sp03 = '" & m_CP03 & "' and sp04 = '" & m_CP04 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      If Not IsNull(.Fields("Cno").Value) Then
         'Modified by Morgan 2015/8/21 --Kimi
         'm_tmp = "¶Q©Ò¾ã²zµf†A¡G" & .Fields("Cno").Value
         m_tmp = "¶Q¤è¾ã²zµf" & PUB_GetUniText(Me.Name, "¸¹") & "¡G" & .Fields("Cno").Value
         intHeadRow = intHeadRow + 1
         SetWordArray m_PlusHead, intHeadRow, 3, m_tmp
      End If
      m_tmp = "»ÈÀs¾ã²zµf" & PUB_GetUniText(Me.Name, "¸¹") & "¡G" & .Fields("Yno").Value
      intHeadRow = intHeadRow + 1
      SetWordArray m_PlusHead, intHeadRow, 3, m_tmp
      End With
   Else
      'm_tmp = "¶Q©Ò¾ã²zµf†A¡G"
      'intHeadRow = intHeadRow + 1
      'SetWordArray m_PlusHead, intHeadRow, 3, m_tmp
      m_tmp = "»ÈÀs¾ã²zµf" & PUB_GetUniText(Me.Name, "¸¹") & "¡G"
      intHeadRow = intHeadRow + 1
      SetWordArray m_PlusHead, intHeadRow, 3, m_tmp
   End If
   
   m_tmp = "¹ú©Ò¾ã²zµf" & PUB_GetUniText(Me.Name, "¸¹") & "¡G" & Replace(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04, "-0-00", "")
   intHeadRow = intHeadRow + 1
   SetWordArray m_PlusHead, intHeadRow, 3, m_tmp
   
   m_tmp = ReportSum(130)
   intHeadRow = intHeadRow + 1
   SetWordArray m_PlusHead, intHeadRow, 1, m_tmp
   
   '¶×²v
   If m_DNCurr <> "NTD" And m_iPrintCurrType <> 3 Then
      If m_DNCurr <> "USD" Then
         m_tmp = "Currency Rate: " & m_DNCurr & "1.00=NTD"
      Else
         m_tmp = ReportSum(75)
      End If
      m_tmp = vbCrLf & m_tmp & Format(Val("0" & m_DNRate), "0.00")
      SetWordArray m_PlusFooter, 1, 1, m_tmp
   End If
End Sub
'Added by Morgan 2016/1/12
Private Function GetYourRef(pA1k13 As String, pA1k01 As String, pRefNo As String) As String
   Dim StrSQLa As String, intQ As Integer, strRefNo As String
   Dim rsQuery As ADODB.Recordset
   
   If InStr(",P,CFP,FCP,T,FCT,CFT,TF,", "," & pA1k13 & ",") > 0 Then
      '±M§Q
      If InStr(",P,CFP,FCP,", "," & pA1k13 & ",") > 0 Then
         StrSQLa = "Select PA106 From Patent, CaseProgress Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP60='" & pA1k01 & "' And CP10='605'  and pa76 is not null"
      '°Ó¼Ð
      Else
         StrSQLa = "Select TM65 From TRADEMARK, CaseProgress Where TM01=CP01 And TM02=CP02 And TM03=TM03 And TM04=CP04 And CP60='" & pA1k01 & "' And CP10='102' and tm33  is not null"
      End If
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, StrSQLa)
      '¦³¦~¶O¥N²z¤H«h§ì¦~¶O©¼©Ò®×¸¹
      If intQ = 1 Then
         If PUB_GetFCCaseNo(pA1k01, strRefNo, True) Then
            GetYourRef = strRefNo
         Else
            'Modify By Sindy 2016/1/29
            'GetYourRef = rsQuery(0).Value
            GetYourRef = "" & rsQuery(0).Value
            '2016/1/29 END
         End If
      '¤@¯ë
      '¥ý§ì²§°Ê
      ElseIf PUB_GetFCCaseNo(pA1k01, strRefNo) Then
         GetYourRef = strRefNo
      'end 2014/2/17
      Else
         GetYourRef = pRefNo
      End If
      
   '¨ä¥L
   ElseIf PUB_GetFCCaseNo(pA1k01, strRefNo) Then
      GetYourRef = strRefNo
   Else
      GetYourRef = pRefNo
   End If
   Set rsQuery = Nothing
End Function

'¨ú±o­^¤å¤é´Á®æ¦¡
'Modified by Moran 2018/4/11 +pFormat
Private Function GetEngDate(pDate As String, Optional pFormat As Integer = 0) As String
   Dim stDate As String
   stDate = AFDate(DBDATE(pDate))
   If pFormat = 1 Then
      'Modified by Morgan 2025/3/17
      'stDate = Format(stDate, "DD/MM/YYYY")
      stDate = Format(stDate, "DD") & "/" & Format(stDate, "MM") & "/" & Format(stDate, "YYYY")
      'end 2025/3/17
   Else
      If Month(stDate) = 5 Then
         stDate = Format(stDate, "mmm d, yyyy")
      Else
         stDate = Format(stDate, "mmm. d, yyyy")
      End If
   End If
   GetEngDate = stDate
End Function
'Added by Morgan 2017/8/17
'ÀË¬d Dow ±b³æªº¯S®í®æ¦¡
Private Function ChkDowXFormat(pKey As String, Optional pChkCode As Boolean = False) As Boolean
   'Modified by Morgan 2021/3/3 +903±M§Q½Õ¬d--§d±mµÙ
   'Modified by Morgan 2022/6/24 +433»~Ä¶­q¥¿ --§õ¹D©û
   'Modified by Morgan 2024/11/7 +206¸É¥R»¡©ú--Franny
   'Modified by Lydia 2025/06/09 +402§ó¥¿--Tim
   Const ItemNoList As String = "'1202','205','1002','107','203','206','431','422','204','903','433','402'"
   Dim stSQL As String, iQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   If pChkCode Then
      If InStr(ItemNoList, "'" & pKey & "'") > 0 Then
         ChkDowXFormat = True
      End If
   Else
      stSQL = "select * from acc1l0 where a1l01='" & pKey & "' and a1l04 in (" & ItemNoList & ") and rownum<2"
      iQ = 1
      Set rsQuery = ClsLawReadRstMsg(iQ, stSQL)
      If iQ = 1 Then
         ChkDowXFormat = True
      End If
   End If
   Set rsQuery = Nothing
End Function
'Added by Morgan 2022/12/19
'ÀË¬d¬O§_§t¯S©w½Ð´Ú¶µ¥Ø
Private Function ChkItemExist(pKey As String, pItemCode As String) As Boolean
   Dim stSQL As String, iQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select * from acc1l0 where a1l01='" & pKey & "' and a1l04='" & pItemCode & "' and rownum<2"
   iQ = 1
   Set rsQuery = ClsLawReadRstMsg(iQ, stSQL)
   If iQ = 1 Then
      ChkItemExist = True
   End If
   Set rsQuery = Nothing
End Function
Private Sub PrintSpecialHead(ByRef intRRow As Integer, ByVal pClientMatterID As String, ByVal pAccNo As String)
   'Added by Morgan 2023/3/6 Y55864000 HOIBERG AB c/o HOIBERG--Franny
   If m_A1k27 = "Y55864000" Then
      intRRow = intRRow + 1
      m_tmp = "Org. No. 559058-3158"
      PutData m_tmp, intRRow, 5500
      SetWordArray m_Head, intRRow, 3, m_tmp
   End If
   'end 2023/3/6
   
   'Added by Morgan 2017/7/7
   'BASF
   'Modfiied by Morgan 2021/7/1 +Y3326801 BASF Corporation --Anny
   If m_A1k28 = "Y45814010" Or m_A1k28 = "Y33268010" Then
      intRRow = intRRow + 1
      If m_A1k03 <> "Y52263010" Then 'Added by Morgan 2020/4/9 ¥N²z¤HY52263010ªº®×¥ó°£¥~--Franny
         m_tmp = "Record ID: " & pClientMatterID
         PutData m_tmp, intRRow, 5500
         SetWordArray m_Head, intRRow, 3, m_tmp
      End If
        
      m_tmp = "Purchase Order No.: " & pAccNo
      intRRow = intRRow + 1
      PutData m_tmp, intRRow, 5500
      SetWordArray m_Head, intRRow, 3, m_tmp
      
      'Added by Morgan 2021/7/1
      m_tmp = "Valid Period of the PO: 01.06.2021-31.05.2024"
      intRRow = intRRow + 1
      PutData m_tmp, intRRow, 5500
      SetWordArray m_Head, intRRow, 3, m_tmp
      'end 2021/7/1
      
      m_tmp = "VAT ID: DE149145247"
      intRRow = intRRow + 1
      PutData m_tmp, intRRow, 5500
      SetWordArray m_Head, intRRow, 3, m_tmp
   End If
   
   'Modified by Morgan 2021/12/7 +Y51467020 Saurer Spinning Solutions GmbH & Co. KG --Franny
   If m_A1k28 = "Y45814010" Or m_A1k28 = "Y33268010" Or m_A1k28 = "Y51467020" Then
      'Service Period=½Ð´Ú¤éªº¦¸¤ë1¸¹¨ì¤ë©³
      'modify by sonia 2017/7/13 §ï¬°½Ð´Ú¤éªº«e¤ë1¸¹¨ì¤ë©³
      'strExc(1) = CompDate(1, 1, Left(m_strA1K02, 5) & "01")
      strExc(1) = CompDate(1, -1, Left(m_strA1K02, 5) & "01")
      'end 2017/7/13
      strExc(2) = CompDate(1, 1, strExc(1))
      strExc(2) = CompDate(2, -1, strExc(2))
      strExc(3) = GetEngDate(strExc(1))
      strExc(4) = GetEngDate(strExc(2))
      'Modified by Morgan 2021/12/7 §ï¦C¦L®É¦A¥[§Ç¼Æµü(­ì¨Ó¦³¦³¿ù,°£1,2,3,21,22,23,31¥~¡A¨ä¥L³£¥[ th)
      'strExc(3) = Left(strExc(3), InStr(strExc(3), " ") - 1)
      'm_tmp = "Service Period: " & strExc(3) & " 1st" & " to " & strExc(3) & " " & Right(strExc(2), 2) & "st, " & Left(strExc(1), 4)
      m_tmp = "Service Period: " & Left(strExc(3), InStr(strExc(3), ",") - 1) & " to " & strExc(4)
      'end 2021/12/7
      intRRow = intRRow + 1
      PutData m_tmp, intRRow, 5500
      SetWordArray m_Head, intRRow, 3, m_tmp
   
   'Added by Morgan 2025/6/10 +Y55506000 --Tim
   'Service Period=³Ì¦­¦¬¤å¤é¨ì³Ì±ßµo¤å¤é
   ElseIf m_A1k03 = "Y55506000" Then
      strExc(1) = GetMinCp05(m_strDN)
      strExc(2) = GetCp27(m_strDN, , True)
      strExc(3) = GetEngDate(strExc(1))
      strExc(4) = GetEngDate(strExc(2))
      If Right(strExc(3), 4) = Right(strExc(4), 4) Then
         m_tmp = "Service Period: " & Left(strExc(3), InStr(strExc(3), ",") - 1) & " to " & strExc(4)
      Else
         m_tmp = "Service Period: " & vbCrLf & strExc(3) & " to " & strExc(4)
      End If
      intRRow = intRRow + 1
      PutData m_tmp, intRRow, 5500
      SetWordArray m_Head, intRRow, 3, m_tmp
   'end 2025/6/10
   End If
   
   'Lundbeck
   If m_A1k28 = "Y45493000" Then
      
      intRow = intRow + 1
      m_tmp = "Attention Patent Invoices"
      PutData m_tmp, intRow
      SetWordArray m_Head, intRow, 2, m_tmp
            
      m_tmp = "Lundbeck Costcenter: 10008133"
      intRRow = intRRow + 1
      PutData m_tmp, intRRow, 5500
      SetWordArray m_Head, intRRow, 3, m_tmp
      
      m_tmp = "Lundbeck Account: 51330200"
      intRRow = intRRow + 1
      PutData m_tmp, intRRow, 5500
      SetWordArray m_Head, intRRow, 3, m_tmp
      
      m_tmp = "Time Period Covered: "
      intRRow = intRRow + 1
      PutData m_tmp, intRRow, 5500
      SetWordArray m_Head, intRRow, 3, m_tmp
      
      strExc(1) = GetMinCp05(m_strDN)
      strExc(1) = GetEngDate(strExc(1))
      strExc(2) = GetEngDate(adoacc1k0("a1k02"))
      m_tmp = strExc(1) & " ~ " & strExc(2)
      intRRow = intRRow + 1
      PutData m_tmp, intRRow, 5500
      SetWordArray m_Head, intRRow, 3, m_tmp
   End If
   
   'Added by Morgan 2018/4/25 --§d±mµÙ
   If m_A1k28 = "Y54997000" Then
      m_tmp = "Your cost center: " & pClientMatterID
      intRRow = intRRow + 1
      PutData m_tmp, intRRow, 5500
      SetWordArray m_Head, intRRow, 3, m_tmp
   End If
   
   'Added by Morgan 2018/11/29 --Lina
   If m_A1k28 = "Y34412010" Then
      m_tmp = "Contract No.: CM0051A22420350579"
      If intRRow < 7 Then
         intRRow = 7
      Else
         intRRow = intRRow + 1
      End If
      PutData m_tmp, intRRow, 5500
      SetWordArray m_Head, intRRow, 3, m_tmp
   End If
   
   'Added by Morgan 2018/3/22 --Lina
   If m_A1k28 = "Y20600000" Then
      If (adoacc1k0.Fields("a1k13") = "FCP" Or adoacc1k0.Fields("a1k13") = "P") Then
         m_tmp = "Instruction Date: " & Format(GetMinCp05(m_strDN), "####/##/##")
         intRRow = intRRow + 1
         PutData m_tmp, intRRow, 5500
         SetWordArray m_Head, intRRow, 3, m_tmp
         
      'Added by Morgan 2018/5/25 --³¯ª÷½¬
      ElseIf (adoacc1k0.Fields("a1k13") = "FCT" Or adoacc1k0.Fields("a1k13") = "S") Then
         m_tmp = "Attn: Christophe Saliou"
         intRRow = intRRow + 1
         PutData m_tmp, intRRow, 5500
         SetWordArray m_Head, intRRow, 3, m_tmp
         
         m_tmp = "Instruction Date: "
         intRRow = intRRow + 1
         PutData m_tmp, intRRow, 5500
         SetWordArray m_Head, intRRow, 3, m_tmp
      'end 2018/5/25
      End If
   End If
   
   'Added by Morgan 2022/4/19 --³¯ª÷½¬
   'Y54038 ThyssenKrupp Intellectual Property GmbH
   If m_A1k28 = "Y54038000" Then
      If (adoacc1k0.Fields("a1k13") = "FCT" Or adoacc1k0.Fields("a1k13") = "S") Then
         m_tmp = "Performance Period: "
         intRRow = intRRow + 1
         PutData m_tmp, intRRow, 5500
         SetWordArray m_Head, intRRow, 3, m_tmp
      End If
   End If
   'end 2022/4/19
   
   'Added by Morgan 2018/4/11
   If m_bSpecialNew4 Then
      m_tmp = "VAT Number: 04146457"
      intRRow = intRRow + 1
      SetWordArray m_Head, intRRow, 3, m_tmp
      m_tmp = "VAT Rate: 0"
      intRRow = intRRow + 1
      SetWordArray m_Head, intRRow, 3, m_tmp
      m_tmp = "Amount VAT: 0"
      intRRow = intRRow + 1
      SetWordArray m_Head, intRRow, 3, m_tmp
      m_tmp = "Total Amount VAT included: 0"
      intRRow = intRRow + 1
      SetWordArray m_Head, intRRow, 3, m_tmp
      m_tmp = "Attn: Ms. Susanna Hostettler"
      intRRow = intRRow + 1
      SetWordArray m_Head, intRRow, 3, m_tmp
      'Modified by Morgan 2020/7/1 --²øÞ±¤Z
      'm_tmp = "(Susanna.Hostettler@bobst.com)"
      m_tmp = "Susanna.Hostettler@bobst.com"
      'end 2020/7/1
      intRRow = intRRow + 1
      SetWordArray m_Head, intRRow, 3, m_tmp
   End If
   
   'Added by Morgan 2019/6/3 --³¯ª÷½¬
   If m_A1k28 = "Y43169000" Then
      intRRow = intRRow + 1
      m_tmp = "Codes:"
      PutData m_tmp, intRRow, 5500
      SetWordArray m_Head, intRRow, 3, m_tmp
      m_tmp = "L824701"
      PutData m_tmp, intRRow, 6500
      SetWordArray m_Head, intRRow, 4, m_tmp
      
      intRRow = intRRow + 1
      m_tmp = "I8S60015"
      PutData m_tmp, intRRow, 6500
      SetWordArray m_Head, intRRow, 4, m_tmp
   End If
   'end 2019/6/3
   
   'Added by Morgan 2019/9/19 --³¯ª÷½¬
   If m_A1k03 = "Y52543000" Then
      'Modified by Morgan 2023/2/21
      'm_tmp = "Purchase Order Number: 4300004798"
      m_tmp = "Vendor Number: 20718822"
      'end 2023/2/21
      intRRow = intRRow + 1
      PutData m_tmp, intRRow, 5500
      SetWordArray m_Head, intRRow, 3, m_tmp
   End If
   'end 2019/9/19
   
   'Added by Morgan 2020/7/8 --²øÞ±¤Z
   If m_A1k28 = "Y52418000" Then
      m_tmp = "Tai E Business ID: 04146457"
      intRRow = intRRow + 1
      PutData m_tmp, intRRow, 5500
      SetWordArray m_Head, intRRow, 3, m_tmp
   End If
   'end 2020/7/8
   
   'Added by Morgan 2021/5/24 --Ryan
   If m_A1k28 = "Y19893030" Then
      m_tmp = "Hourly Rate: NTD" & IIf(m_LD16 <> "", m_LD16, "4500")
      intRRow = intRRow + 1
      PutData m_tmp, intRRow, 5500
      SetWordArray m_Head, intRRow, 3, m_tmp
   'Added by Morgan 2025/3/14
   ElseIf m_A1k28 = "Y56142000" Then
      m_tmp = "Hourly Rate: USD" & IIf(m_LD16 <> "", m_LD16, "149")
      intRRow = intRRow + 1
      PutData m_tmp, intRRow, 5500
      SetWordArray m_Head, intRRow, 3, m_tmp
   End If
   'end 2021/5/24
   
   'Added by Morgan 2020/11/12 --³¯ª÷½¬
   'Removed by Morgan 2021/4/27 ¨ú®ø¡A§ï©ñ«È¤á®×¥ó®×¸¹--³¯ª÷½¬
   'If strCust1 = "X19893020" And (adoacc1k0.Fields("A1K13").Value = "FCT" Or adoacc1k0.Fields("A1K13").Value = "S") Then
   '   m_tmp = "Contact: Brigitte Kuster"
   '   intRRow = intRRow + 1
   '   PutData m_tmp, intRRow, 5500
   '   SetWordArray m_Head, intRRow, 3, m_tmp
   'End If
   'end 2021/4/27
   'end 2020/11/12
   
   'Added by Morgan 2022/5/13 --³¯ª÷½¬
   If m_A1k28 = "X83548010" Then
      m_tmp = "Corporate registration no: 556734-2026"
      intRRow = intRRow + 1
      PutData m_tmp, intRRow, 5500
      SetWordArray m_Head, intRRow, 3, m_tmp
   End If
   'end 2022/5/13
   
   'Added by Morgan 2022/8/5
   'Y55751 Birkenstock IP GmbH--Franny
   If m_A1k28 = "Y55751000" Then
      intRow = intRow + 1
      
      If (adoacc1k0.Fields("a1k13") = "FCP" Or adoacc1k0.Fields("a1k13") = "P") Then
         m_tmp = "Attn: Mrs. Daniela Denner, Mr. Marvin Petzold"
         SetWordArray m_Head, intRow, 2, m_tmp
         intRow = intRow + 2
      End If
      
      m_tmp = "Cost Center: 35007100"
      SetWordArray m_Head, intRow, 2, m_tmp
      
      intRow = intRow + 1
      If (adoacc1k0.Fields("a1k13") = "FCP" Or adoacc1k0.Fields("a1k13") = "P") Then
         m_tmp = "Internal Order Number: 61672"
      Else
         m_tmp = "Internal Order Number: 61671"
      End If
      SetWordArray m_Head, intRow, 2, m_tmp
   End If
   'end 2022/8/5
   
   'Added by Morgan 2023/2/20 --²øÞ±¤Z
   If m_A1k27 = "Y55261000" And adoacc1k0.Fields("a1k13") = "FCP" Then
      If m_Head(2, UBound(m_Head, 2)) <> "" Then intRow = intRow + 1
      m_tmp = "FISCAL CODE BRSDVD80E24F205D"
      SetWordArray m_Head, intRow, 2, m_tmp
   End If
   'end 2023/2/20
   
   'Added by Morgan 2023/3/22 --Izumi
   If m_A1k27 = "Y52150B20" And adoacc1k0.Fields("a1k13") = "FCP" Then
      If m_Head(2, UBound(m_Head, 2)) <> "" Then intRow = intRow + 1
      intRow = intRow + 1
      m_tmp = "Attn: " & m_Att1
      SetWordArray m_Head, intRow, 2, m_tmp
   End If
   'end 2023/3/22
   
   'Added by Morgan 2024/12/13 ½Ð´Ú¹ï¶H¬°Y55822030®É¥[¦L®×¥óÁpµ¸¤H --®Û­^
   If m_A1k28 = "Y55822030" Then
      If m_Head(2, UBound(m_Head, 2)) <> "" Then intRow = intRow + 1
      intRow = intRow + 1
      SetWordArray m_Head, intRow, 2, m_Att1
   End If
   'end 2024/12/13
End Sub
'Added by Morgan 2018/1/5
'±b³æ»y¤å
Private Function GetBillLanguage(pLanguage As String) As Boolean
   'Y54770+X48886 ³]©w±b³æ»y¨¥:­^¤å --¦ó²QµØ
   'Modified by Morgan 2019/12/13 ¥Ó½Ð¤H§ï§ì8½X(X48886000¦³§ó¦W)--Kimi
   If m_A1k03 = "Y54770000" And Left(strCust1, 8) = "X4888600" Then
      pLanguage = "2"
      GetBillLanguage = True
   'Added by Morgan 2018/5/15 ½Ð´Ú¹ï¶H¬OY54975»õ®õºë±K¤u·~¡]«n¨Ê¡^¦³­­¤½¥q¡A¥Ó½Ð¤H«ü¥Ü¡A­n¥Î¤¤¤å±b³æ--Joseph Lo
   ElseIf m_A1k28 = "Y54975000" Then
      pLanguage = "1"
      GetBillLanguage = True
   'Added by Morgan 2020/10/19 Y51817040(King & Wood Mallesons)³]©w±b³æ»y¨¥:­^¤å --§d±mµÙ
   ElseIf m_A1k03 = "Y51817040" Then
      pLanguage = "2"
      GetBillLanguage = True
   'Added by Lydia 2024/06/28 ½Ð³]©wY56042000 + X56559000 ±b³æ»y¨¥¬°¤¤¤å --§d±mµÙ
   ElseIf m_A1k03 = "Y56042000" And Left(strCust1, 8) = "X5655900" Then
      pLanguage = "1"
      GetBillLanguage = True
      
   'Added by Morgan 2024/12/12 ½Ð´Ú¹ï¶H¬°Y55822030®É³]©w±b³æ»y¨¥©T©w¬°­^¤å --®Û­^
   ElseIf m_A1k28 = "Y55822030" Then
      pLanguage = "2"
      GetBillLanguage = True
      
   End If
End Function

'Added by Morgan 2018/6/27
Public Sub GetLedes(pArr() As String)
   pArr() = strLedes()
End Sub
'Added by Morgan 2018/8/30
Private Sub UpdateLEDES()
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   With adoacc1k0
   stSQL = "select * from acc260" & _
      " where a2601='" & Left(m_A1k28, 8) & "' and a2602='" & .Fields("a1k13") & "'" & _
      " and a2603='" & .Fields("a1j02") & "'"
   End With
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      '18 TIMEKEEPER_ID
      If "" & .Fields("a2617") <> "" Then
         strLedes(18, iUpper) = "" & .Fields("a2617")
      End If
      '22 TIMEKEEPER_NAME
      If "" & .Fields("a2620") <> "" Then
         strLedes(22, iUpper) = "" & .Fields("a2620")
      End If
      
      If Val("" & .Fields("a2621")) > 0 Then
         '21 LINE_ITEM_UNIT_COST(Rate)
         strLedes(21, iUpper) = "" & .Fields("a2621")
         '11 LINE_ITEM_NUMBER_OF_UNITS
         strLedes(11, iUpper) = Round((Val(strLedes(13, iUpper)) - Val(strLedes(12, iUpper))) / Val(strLedes(21, iUpper)), 4)
      End If
      
      If m_iLedesVer = 2 Then 'Added by Morgan 2023/6/19
         '32 TIMEKEEPER_FIRST_NAME
         If "" & .Fields("a2618") <> "" Then
            strLedes(32, iUpper) = "" & .Fields("a2618")
         End If
         '31 TIMEKEEPER_LAST_NAME
         If "" & .Fields("a2619") <> "" Then
            strLedes(31, iUpper) = "" & .Fields("a2619")
         End If
      End If
      End With
   End If
   Set rsQuery = Nothing
End Sub

'Added by Morgan 2021/2/4
'Y54225B10 Syngenta ¯S®í³]©w
Private Sub UpdateLEDES2()
   If strLedes(10, iUpper) = "F" Then
      '10 EXP/FEE/INV_ADJ_TYPE(¶µ¥ØÃþ§O)
      strLedes(10, iUpper) = "IF"
      '21 LINE_ITEM_UNIT_COST(Rate)
      strLedes(21, iUpper) = ""
      '11 LINE_ITEM_NUMBER_OF_UNITS
      strLedes(11, iUpper) = "1"
      '12 LINE_ITEM_ADJUSTMENT_AMOUNT(§é¦©)
      strLedes(12, iUpper) = strLedes(13, iUpper)
      '15 LINE_ITEM_TASK_CODE
      strLedes(15, iUpper) = ""
      '17 LINE_ITEM_ACTIVITY_CODE
      strLedes(17, iUpper) = ""
      '18 TIMEKEEPER_ID
      strLedes(18, iUpper) = ""
      '22 TIMEKEEPER_NAME
      strLedes(22, iUpper) = ""
      '23 TIMEKEEPER_CLASSIFICATION
      strLedes(23, iUpper) = ""
      '31 TIMEKEEPER_LAST_NAME
      strLedes(31, iUpper) = ""
      '32 TIMEKEEPER_FIRST_NAME
      strLedes(32, iUpper) = ""
   End If
   
   '24 CLIENT_MATTER_ID
   strLedes(24, iUpper) = "206113"
   '48 LINE_ITEM_TAX_RATE
   strLedes(48, iUpper) = "0"
End Sub
'Added by Morgan 2018/9/18
Private Function GetDisc(pDocNo As String) As String
   Dim stSQL As String
   Dim intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select * from acc1l0 where a1l01='" & pDocNo & "' and a1l07>0 and a1l19>0 order by a1l02"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      GetDisc = (rsQuery("a1l19") * 100) & "%"
   End If
   
   Set rsQuery = Nothing
End Function

'Added by Morgan 2019/2/14 ±q PrintHead ©â¥X
Private Sub PrintRightHeadCol(pVaule1 As String, ByRef intRRow As Integer, Optional pVaule2 As String = "")
   intRRow = intRRow + 1
   PutData pVaule1, intRRow, 5500
   SetWordArray m_Head, intRRow, 3, pVaule1
   If pVaule2 <> "" Then
      PutData pVaule2, intRRow, 6500
      SetWordArray m_Head, intRRow, 4, pVaule2
   End If
End Sub
'Added by Morgan 2019/8/30
'ÀË¹î¥D°Ê­×¥¿µo¤å¤é¬O§_±ß©ó¤¤»¡
Private Function Is203Amendent(pDnNo As String) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select * from caseprogress a where cp60='" & pDnNo & "' and cp10='203' and cp27>0" & _
      " and exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04" & _
      " and b.cp10 in ('201','209','210','235') and b.cp27<a.cp27)"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      Is203Amendent = True
   End If
   Set rsQuery = Nothing
End Function

'Added by Morgan 2019/9/9
'±q PrintHead ©â¥X
Private Sub PrintCaseNo(pRRow As Integer)
   Dim stTmp As String
   
   'Added by Morgan 2019/2/14--³¯ª÷½¬
   If strCust1 = "X70124000" And (adoacc1k0.Fields("A1K13").Value = "FCT" Or adoacc1k0.Fields("A1K13").Value = "S") Then
      PrintRightHeadCol "Nordson ref.: " & adoquery.Fields("Cno").Value, pRRow
   'Added by Morgan 2019/2/15--³¯ª÷½¬
   'Modified by Morgan 2019/4/11 +Y34440B6--³¯ª÷½¬
   'Modify By Sindy 2021/3/3 X74676010 => X74676000
   'Modified by Morgan 2022/8/16 ¨ú®ø X74676000--³¯ª÷½¬
   'ElseIf (strCust1 = "X74676000" Or adoacc1k0.Fields("a1k03").Value = "Y34440B60") And (adoacc1k0.Fields("A1K13").Value = "FCT" Or adoacc1k0.Fields("A1K13").Value = "S") Then
   ElseIf (adoacc1k0.Fields("a1k03").Value = "Y34440B60") And (adoacc1k0.Fields("A1K13").Value = "FCT" Or adoacc1k0.Fields("A1K13").Value = "S") Then
      PrintRightHeadCol "Instructor: " & adoquery.Fields("Cno").Value, pRRow
   'end 2019/2/14
   'Added by Morgan 2019/5/13--³¯ª÷½¬
   ElseIf (adoacc1k0.Fields("a1k03").Value = "Y51409000" Or adoacc1k0.Fields("a1k03").Value = "Y51409B10") And (adoacc1k0.Fields("A1K13").Value = "FCT" Or adoacc1k0.Fields("A1K13").Value = "S") Then
      PrintRightHeadCol "Attorney: " & adoquery.Fields("Cno").Value, pRRow
   'end 2019/5/13
   'Added by Morgan 2019/5/17--³¯ª÷½¬
   ElseIf adoacc1k0.Fields("a1k03").Value = "Y52622000" And (adoacc1k0.Fields("A1K13").Value = "FCT" Or adoacc1k0.Fields("A1K13").Value = "S") Then
      PrintRightHeadCol "Keltie fee earner: " & adoquery.Fields("Cno").Value, pRRow
   'end 2019/5/17
   
   'Added by Morgan 2019/9/9--³¯ª÷½¬
   'Modified by Morgan 2023/3/14 +Y53817
   ElseIf InStr("Y53598000,Y53817000", adoacc1k0.Fields("a1k03").Value) > 0 And (adoacc1k0.Fields("A1K13").Value = "FCT" Or adoacc1k0.Fields("A1K13").Value = "S") Then
      PrintRightHeadCol "Instructing attorney: " & adoquery.Fields("Cno").Value, pRRow
   'end 2019/5/17
   
   'Added by Morgan 2019/9/5 --Franny
   ElseIf (adoacc1k0.Fields("a1k03").Value = "Y52875000" Or adoacc1k0.Fields("a1k03").Value = "Y54983000") And (adoacc1k0.Fields("A1K13").Value = "FCP" Or adoacc1k0.Fields("A1K13").Value = "FG") Then
      PrintRightHeadCol "BU:", pRRow, "" & adoquery.Fields("Cno").Value
   'end 2019/9/5
   
   'Added by Morgan 2023/12/14 --Franny
   ElseIf adoacc1k0.Fields("a1k03").Value = "Y55948000" Then
      PrintRightHeadCol "PO Number: " & adoquery.Fields("Cno").Value, pRRow
      
   'Added by Morgan 2019/12/5--³¯ª÷½¬
   ElseIf (strCust1 = "X64775000" Or strCust1 = "X61139010") And (adoacc1k0.Fields("A1K13").Value = "FCT" Or adoacc1k0.Fields("A1K13").Value = "S") Then
      PrintRightHeadCol "Client Matter Name: " & adoquery.Fields("Cno").Value, pRRow
      stTmp = GetClientMatterID(adoacc1k0.Fields("a1k13"), adoacc1k0.Fields("a1k14"), adoacc1k0.Fields("a1k15"), adoacc1k0.Fields("a1k16"), adoacc1k0.Fields("a1k01"), adoacc1k0.Fields("a1k28"))
      PrintRightHeadCol "Client Matter ID: " & stTmp, pRRow
      
   'Added by Morgan 2020/1/17--³¯ª÷½¬
   ElseIf adoacc1k0.Fields("a1k03").Value = "Y55374000" And (adoacc1k0.Fields("A1K13").Value = "FCT" Or adoacc1k0.Fields("A1K13").Value = "S") Then
      PrintRightHeadCol "Attention: " & adoquery.Fields("Cno").Value, pRRow
      
   'Added by Morgan 2021/4/27 --³¯ª÷½¬
   'Modified by Morgan 2022/5/13 +½Ð´Ú¹ï¶H X8354801--³¯ª÷½¬
   ElseIf (strCust1 = "X19893020" Or adoacc1k0.Fields("a1k28").Value = "X83548010") And (adoacc1k0.Fields("A1K13").Value = "FCT" Or adoacc1k0.Fields("A1K13").Value = "S") Then
      PrintRightHeadCol "Contact: " & adoquery.Fields("Cno").Value, pRRow
   'end 2021/4/27
   
   'Added by Morgan 2022/9/16 --³¯ª÷½¬,Anny
   ElseIf m_A1k27 = "X64826000" Then
      PrintRightHeadCol "Purchase Order number: " & adoquery.Fields("Cno").Value, pRRow
   
   'Added by Morgan 2025/5/27
   'Potter Clarkson (Y20115)©Ò¦³°Ó¼Ð®×¥ó½Ð´Ú³æ¤§¡u«È¤á®×¥ó®×¸¹¡v(Case No.) Äæ¦ì¦WºÙ§ó·s¬°¡uPurchase Order Number¡v--¶À«w¹F/Ú{«º
   ElseIf adoacc1k0.Fields("a1k03").Value = "Y20115000" And (adoacc1k0.Fields("A1K13").Value = "FCT" Or adoacc1k0.Fields("A1K13").Value = "S") Then
      PrintRightHeadCol "Purchase Order number: " & adoquery.Fields("Cno").Value, pRRow
   
   'end 2022/9/16
   '¦³«È¤á®×¥ó®×¸¹
   ElseIf IsNull(adoquery.Fields("Cno").Value) = False Then
      'Added by Morgan 2021/4/20 --Ryan
      If adoacc1k0.Fields("a1k03").Value = "Y53603020" Then
         PrintRightHeadCol adoquery.Fields("Cno").Value, pRRow
      'end 2021/4/20
      Else
      'end 2020/1/17
         PrintRightHeadCol "Case No.:", pRRow, adoquery.Fields("Cno").Value
      End If
   
   'µL«È¤á®×¥ó®×¸¹
   Else
   
      If adoacc1k0.Fields("a1k03").Value = "Y20438020" Then
         PrintRightHeadCol "Vendor Code # 88125-0", pRRow
         
      'Added by Morgan 2017/5/31 --³¯¼W¼s
      ElseIf adoacc1k0.Fields("a1k28").Value = "Y54117000" Then
         PrintRightHeadCol "VAT : BE 0543573053", pRRow
      'end 2017/5/31
      
      '2010/10/29 add by sonia ¥N²z¤HY52341020¥BFCT¦LInternal Cost Center: DE94-1260
      '2012/6/1 MODIFY BY SONIA §ï¬°(¥N²z¤H¤Î¦C¦L¹ï¶H¬°Y5234102¥B½Ð´Ú¹ï¶H¬°Y52822),©Î(¥N²z¤H¤Î¦C¦L¹ï¶H¬°Y52341¥B½Ð´Ú¹ï¶H¬°Y5282201)³£­n¦L¦¹
      'ElseIf "" & adoacc1k0.Fields("A1K13").Value = "FCT" And adoacc1k0.Fields("a1k03").Value = "Y52341020" Then
      ElseIf "" & adoacc1k0.Fields("A1K13").Value = "FCT" And ((adoacc1k0.Fields("a1k03").Value = "Y52341020" And adoacc1k0.Fields("a1k27").Value = "Y52341020" And adoacc1k0.Fields("a1k28").Value = "Y52822000") Or (adoacc1k0.Fields("a1k03").Value = "Y52341000" And adoacc1k0.Fields("a1k27").Value = "Y52341000" And adoacc1k0.Fields("a1k28").Value = "Y52822010")) Then
         PrintRightHeadCol "Internal Cost Center: DE94-1260", pRRow
         PrintRightHeadCol "Contact Person: Lorenzo Fanti", pRRow
         PrintRightHeadCol "Cost Center: EX9131: Trademarks - Sandoz", pRRow
      '2010/10/29 end
      
      '2010/11/12 add by sonia ½Ð´Ú¹ï¶HY33989010©Ò¦³®×¥ó¦LYour VAT No.: 1188263
      '2011/12/9 modify by sonia §ï¬°118263
      ElseIf adoacc1k0.Fields("a1k28").Value = "Y33989010" Then
         PrintRightHeadCol "Your VAT No.: 118263", pRRow
         PrintRightHeadCol "Our VAT No.: 04146457", pRRow
      '2010/10/29 end
      End If
   End If
   
End Sub

'Added by Morgan 2020/5/21
Private Sub UpdateLEDES3(pA1k13 As String, pA1k14 As String, pA1k15 As String, pA1k16 As String, pA1j02 As String)
   Dim stRefNo As String
   
   'GE Y53971000 ¯S®í»Ý¨D
   'Added by Morgan 2019/12/30 --Ryan
   'Modified by Morgan 2021/10/19 +Y53971B1
   If (m_A1k28 = "Y53971000" Or m_A1k28 = "Y53971B10") And (pA1k13 = "FCP" Or pA1k13 = "P" Or pA1k13 = "FG" Or pA1k13 = "PS") Then
      If strLedes(10, iUpper) = "F" Then
         '21 LINE_ITEM_UNIT_COST(³æ»ù)
         '11 LINE_ITEM_NUMBER_OF_UNITS(³æ¦ì¼Æ/¤u®É)
         strLedes(21, iUpper) = Val(strLedes(13, iUpper)) - Val(strLedes(12, iUpper))
         strLedes(11, iUpper) = "1"
         'Added by Morgan 2021/6/22 ¶O²v¶W¹L¥Ø«e³]©w(USD350)®É¦Û°Ê½Õ¾ã¶O²v¤Î®É¼Æ
         If adoLEDES.Fields("ld16") > 0 And Val(strLedes(21, iUpper)) > adoLEDES.Fields("ld16") Then
            Do While (Val(strLedes(21, iUpper)) > adoLEDES.Fields("ld16"))
               strLedes(21, iUpper) = Val(strLedes(21, iUpper)) / 2
               strLedes(11, iUpper) = Val(strLedes(11, iUpper)) * 2
            Loop
         End If
         'end 2021/6/22
      End If
      
      'Added by Morgan 2021/6/22 --Anny
      stRefNo = GetYourRefNo1(pA1k13, pA1k14, pA1k15, pA1k16)
      'INVOICE_DESCRIPTION «e­±¥[©¼¸¹
      If iUpper = 1 Then
         strLedes(8, iUpper) = stRefNo & " " & strLedes(8, iUpper)
      End If
      'end 2021/6/23
      
      'LINE_ITEM_DESCRIPTION «e­±¥[©¼¸¹
      strLedes(19, iUpper) = stRefNo & " " & strLedes(19, iUpper)
      'end 2021/6/22
   End If
   'end 2019/12/30
   
   'Y55350000 ¯S®í»Ý¨D
   'Modified by Morgan 2020/7/23 +P,FG,PS
   If m_A1k28 = "Y55350000" And (pA1k13 = "FCP" Or pA1k13 = "P" Or pA1k13 = "FG" Or pA1k13 = "PS") Then
      'INVOICE_DESCRIPTION ­n©ñ©¼¸¹¤ÎÁpµ¸¤H1
      stRefNo = GetYourRefNo1(pA1k13, pA1k14, pA1k15, pA1k16)
      strLedes(8, iUpper) = "Merck Reference or Docket: " & stRefNo & "; Attorney: " & m_Att1
      
      ' ¥Ø«e¥u¦³ TASK_CODE:PA400, Activity_code:A111
      If strLedes(10, iUpper) = "F" Then
         strLedes(15, iUpper) = "PA400"
         strLedes(17, iUpper) = "A111"
      End If
      
      'Matter Number ¨Ì½Ð´Ú¶µ¥Ø¤£¦P
      If Left(pA1j02, 3) = "605" Then
         strLedes(24, iUpper) = "20069053"
         'EXPENSE_CODE
         If strLedes(10, iUpper) = "E" Then
            strLedes(16, iUpper) = "E130"
         End If
      ElseIf pA1j02 = "201" Or pA1j02 = "927" Then
         strLedes(24, iUpper) = "2011016"
         'EXPENSE_CODE
         If strLedes(10, iUpper) = "E" Then
            strLedes(16, iUpper) = "E125"
         End If
      Else
         strLedes(24, iUpper) = "20069051"
         'EXPENSE_CODE
         If strLedes(10, iUpper) = "E" Then
            If pA1j02 = "93999" Then
               strLedes(16, iUpper) = "E133"
            Else
               strLedes(16, iUpper) = "E129"
            End If
         End If
      End If
   End If
End Sub
'Added by Morgan 2020/8/6
'µ{§Ç¤Ó¤j©â¥X¬°¤lµ{§Ç
Private Sub InitVar()
   Erase m_Head
   ReDim m_Head(4, 1)
   Erase m_Subject
   ReDim m_Subject(3, 1)
   m_iSubject = 0
   Erase m_Item
   'Modified by Lydia 2015/04/08 ¾ã§å½Ð´Ú³æ°}¦Cºû¼Æ¤£¦P
   'ReDim m_Item(1)
   If m_bolChiDB Then
     ReDim m_Item(adoacc1k0.RecordCount)
   Else
     ReDim m_Item(1)
   End If
   m_iItem = 0
   Erase m_Sum
   'Modified by Morgan 2013/10/29 4.³W¶O,5.ªA°È¶O
   'ReDim m_Sum(3, 1)
   ReDim m_Sum(5, 1)
   Erase m_Footer
   ReDim m_Footer(2, 1)
   m_dblDiscTot = 0
   m_dblNoDiscAmtTot = 0
End Sub

'Added by Lydia 2020/09/08 ³B²z¦C¦L°O¿ý©M¬d¸ß±ø¥ó
'Modified by Morgan 2024/11/1 §ïsub¬°function¥H«K¥Dµ{¦¡§PÂ_¬O§_Ä~Äò
Private Function PrintData_1(ByRef strSpecialList As String, ByRef strSpecialL2 As String, ByRef strA1K01 As String) As Boolean
Dim StrSQLa As String
Dim stConStaff As String 'Added by Morgan 2018/3/21
Dim bolChkIsACCRPT428Exists As Boolean 'Add By Sindy 2013/5/7
Dim dblR42856 As Double, dblR42857 As Double
Dim ii As Integer
Dim stR42836 As String 'Added by Morgan 2024/7/9
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/22 ²M°£¬d¸ß¦Lªí°O¿ýÀÉÄæ¦ì
   If m_bolChiDB = True Then
      strSql = MsgText(601)
      If m_ChiSys <> "" Then
        strSql = strSql & " and a1k13='" & m_ChiSys & "'"
        pub_QL05 = pub_QL05 & ";¨t²ÎÃþ§O:" & m_ChiSys
      End If
      If m_ChiApply <> "" Then
        strSql = strSql & " and a1k28='" & m_ChiApply & "'"
        pub_QL05 = pub_QL05 & ";½Ð´Ú¹ï¶H:" & m_ChiApply
      End If
      If m_ChiCust <> "" Then
         pub_QL05 = pub_QL05 & ";«È¤á½s¸¹:" & m_ChiCust
      End If
      'Added by Lydia 2015/08/10
      If m_ChiArrNO <> "" Then
         strSql = strSql & " and a1k01 in (" & m_ChiArrNO & ") "
         strExc(1) = Replace(m_ChiArrNO, "'", "")
        
         pub_QL05 = pub_QL05 & ";½Ð´Ú½s¸¹:" & strExc(1)
      End If
   Else
   'end 2015/04/09
        strSql = MsgText(601)
        If Text1 <> MsgText(601) Then
            strSql = strSql & " and a1k01>='" & Text1 & "'"
        End If
        If Text2 <> MsgText(601) Then
            strSql = strSql & " and a1k01<='" & Text2 & "'"
        End If
        If Text1 <> MsgText(601) Or Text2 <> MsgText(601) Then
           pub_QL05 = pub_QL05 & ";" & Label1 & Text1 & "-" & Text2 'Add By Sindy 2010/12/22
        End If
        'Add by Morgan 2005/1/4 ÂÂ¸ê®Æ¤£¦L
        strSql = strSql & " AND a1k02>=920201"
        
        If Trim(txtCopy) <> "" Then
           pub_QL05 = pub_QL05 & ";" & Left(Label3, 5) & txtCopy 'Add By Sindy 2010/12/22
        End If
        If Trim(txtAdd) <> "" Then
           pub_QL05 = pub_QL05 & ";" & Left(Label4, 6) & txtAdd 'Add By Sindy 2010/12/22
        End If
        If Trim(txtOutMode) = "1" Then
           pub_QL05 = pub_QL05 & ";" & Left(Label5, 8) & "1:¦Lªí¾÷" 'Add By Sindy 2010/12/22
        Else
           pub_QL05 = pub_QL05 & ";" & Left(Label5, 8) & "2:¹q¤lÀÉ" 'Add By Sindy 2010/12/22
        End If
   End If  'Added by Lydia 2015/04/09 ¤¤¤åª©-¾ã§å½Ð´Ú³æ

   
   'Added by Morgan 2018/3/21
   '°£¹q¸£¤¤¤ß¤Î°]°È³B¥~¨ä¥L¤H­û¤£¥i¸ó¤j³¡ªù¦C¦L½Ð´Ú³æ(¦]µo¥Í°ê¥~³¡¦@¥Î¸ê®Æ§¨·|¦³«D¸Ó³¡ªùªº½Ð´Ú³æ¹q¤lÀÉ¸ê®Æ Ex.FCT_workflow©³¤U¦³FCP¸ê®Æ§¨)
   stConStaff = ""
   'Modified by Morgan 2024/11/1 ½Ð´Ú³æ«Ø¥ß¤H¥i¯à´«³¡ªù,¥ý²¤¹L¾ã§å¦C¦L And m_bBeCalled = False Ex:B3033
   If Pub_StrUserSt03 <> "M51" And Pub_StrUserSt03 <> "M31" And m_bBeCalled = False Then
      stConStaff = " and exists(select * from staff s1 where s1.st01 in (a1k21,a1k24) and substr(s1.ST03,1,2)='" & Left(Pub_StrUserSt03, 2) & "') "
   End If
   'end 2018/3/21
   
'Remove by Morgan 2011/10/11 ²¾¨ì¤U­±¦^°é¤º§PÂ_¤~¯à¼u°T®§
'   If Not m_bEditDoc Then 'Add by Morgan 2010/11/25
'      'Modify by Morgan 2010/5/13 ¯S®í½Ð´Ú³æ¤£¦L
'      strSql = strSql & " AND a1k32 is null"
'   End If

   '¤£­­¨î¤w¦¬´Ú¸ê®Æ
   'Modify by Morgan 2006/9/26 ¥[a1j03, fa10
   'Modify by Morgan 2006/4/10 ¥[fa70
   'Modify by Morgan 2010/11/4 +a1j18,a1j19,a1j20,fa103,fa104,fa105
   'Modify by Morgan 2010/11/22 +¥ý§ì«È»s¤Æ½Ð´Ú¶µ¥Ø¸ê®Æ
   'pSQL = "select a1k27, a1l07, a1l05, a1j04, a1l06, a1j05, a1j06, a1j16, a1j10, fa05, fa63, fa64, fa65, fa32, fa18, a1k02, fa33, fa19, fa34, fa20, a1k13, a1k14, a1k15, a1k16, fa21, fa22, fa35, a1k03, fa36, a1k01, fa06, fa23, a1k04, a1k10, a1l02, fa43, a1k18 as Curr, fa70 as cu102, a1j02, a1k05, FA04, FA17, a1j03, a1k08, a1k11,a1k28,a1j18,a1j19,a1j20,fa103,fa104,fa105 from acc1k0, acc1l0, fagent, acc1j0 where a1k01 = a1l01 (+) and substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and a1l03 = a1j01 (+) and a1l04 = a1j02 (+) and (a1k12 is null or a1k12 = 0) and a1k32 is null" & strSql & " union " & _
             "select a1k27, a1l07, a1l05, a1j04, a1l06, a1j05, a1j06, a1j16, a1j10, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, a1k02, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, a1k13, a1k14, a1k15, a1k16, cu27 as fa21, cu28 as fa22, cu68 as fa35, a1k03, cu69 as fa36, a1k01, cu06 as fa06, cu29 as fa23, a1k04, a1k10, a1l02, cu76 as fa43, a1k18 as Curr, cu102, a1j02, a1k05, CU04 As FA04, CU23 As FA17, a1j03, a1k08, a1k11,a1k28,a1j18,a1j19,a1j20,cu142 fa103,cu143 fa104,cu144 fa105 From acc1k0, acc1l0, customer, acc1j0 where a1k01 = a1l01 (+) and substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) = cu02 and a1l03 = a1j01 (+) and a1l04 = a1j02 (+) and (a1k12 is null or a1k12 = 0) and a1k32 is null" & strSql & " order by a1k01 asc, a1k03 asc, a1l02 asc"
   
   'Modify By Sindy 2011/3/7 +fa108
   'Modify By Sindy 2013/3/28 +a1k33,a1l16,a1l17,a1l04
   'Modify by Sindy 2013/5/8 ­×§ïwhere±ø¥ó­ì¬° and a1j02(+)=a1l04 ==> and a1j02(+)=decode(substr(a1l04,-2),98,substr(a1l04,1,length(a1l04)-2),a1l04)
   '                                           and a2603(+)=a1l04 ==> and a2603(+)=decode(substr(a1l04,-2),98,substr(a1l04,1,length(a1l04)-2),a1l04)
   'Modified by Morgan 2013/10/8 ¨ú®ø fa103,fa104,fa105
   'Modified by Morgan 2014/2/18 +a2616
   'Modified by Morgan 2014/2/24 acc260 §ï§ì¦C¦L¹ï¶H
   'Modified by Morgan 2018/3/21 +staff ±±¨î¸ó¤j³¡ªù¦C¦L°ÝÃD
   'Modified by Morgan 2018/10/18 ­×¥¿½Ð´Ú¶µ¥Ø¬°98°ÝÃD X10715343
   StrSQLa = "select a1k27, a1l07,a1l05,decode(a2607,null,a1j04,a2607) a1j04,a1l06,decode(a2607,null,a1j05,a2608) a1j05,decode(a2607,null,a1j06,a2609) a1j06, a1j16, a1j10, fa05, fa63, fa64, fa65, fa32, fa18, a1k02, fa33, fa19, fa34, fa20, a1k13, a1k14, a1k15, a1k16, fa21, fa22, fa35, a1k03, fa36, a1k01, fa06, fa23, a1k04, a1k10, a1l02, fa43, Curr, cu102, a1j02, a1k05, FA04, FA17, a1j03, a1k08, a1k11,a1k28,decode(a2601,null,a1j18,a2604) a1j18,decode(a2601,null,a1j19,a2605) a1j19,decode(a2601,null,a1j20,a2606) a1j20,a2616,'' as vAY,'' as vAZ,fa108,a2604,a1k32,a1k33,a1l16,a1l17,a1l04" & _
      " from (select a1k27, a1l07, a1l05, a1l06, fa05, fa63, fa64, fa65, fa32, fa18, a1k02, fa33, fa19, fa34, fa20, a1k13, a1k14, a1k15, a1k16, fa21, fa22, fa35, a1k03, fa36, a1k01, fa06, fa23, a1k04, a1k10, a1l02, fa43, a1k18 as Curr, fa70 as cu102, a1k05, FA04, FA17, a1k08, a1k11,a1k28,a1l03,a1l04,fa108,a1k32,a1k33,a1l16,a1l17,a1k21,a1k24 from acc1k0, acc1l0, fagent where a1k01=a1l01(+) and substr(a1k27, 1, 8)=fa01 and substr(a1k27, 9, 1)= fa02 and (a1k12 is null or a1k12=0)" & strSql & _
      " union select a1k27, a1l07, a1l05, a1l06, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, a1k02, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, a1k13, a1k14, a1k15, a1k16, cu27 as fa21, cu28 as fa22, cu68 as fa35, a1k03, cu69 as fa36, a1k01, cu06 as fa06, cu29 as fa23, a1k04, a1k10, a1l02, cu76 as fa43, a1k18 as Curr, cu102, a1k05, CU04 As FA04, CU23 As FA17, a1k08, a1k11,a1k28,a1l03,a1l04,cu148 as fa108,a1k32,a1k33,a1l16,a1l17,a1k21,a1k24 From acc1k0, acc1l0, customer where a1k01=a1l01(+) and substr(a1k27, 1, 8)=cu01 and substr(a1k27, 9, 1)=cu02 and (a1k12 is null or a1k12=0)" & strSql & _
      ") X,acc1j0,acc260" & _
      " where a1j01(+)=a1l03" & _
      " and a1j02(+)=decode(a1l04,'98',a1l04,decode(substr(a1l04,-2),'98',substr(a1l04,1,length(a1l04)-2),a1l04))" & _
      " and a2601(+)=substr(a1k27,1,8)" & _
      " and a2602(+)=a1l03" & _
      " and a2603(+)=decode(a1l04,'98',a1l04,decode(substr(a1l04,-2),'98',substr(a1l04,1,length(a1l04)-2),a1l04))" & stConStaff & _
      " order by a1k01 asc, a1k03 asc, a1l02 asc"
   
   strCon10 = ""
   adoacc1k0.CursorLocation = adUseClient
   adoacc1k0.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc1k0.RecordCount = 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/22
'Remove by Morgan 2011/10/11 ²¾¨ì¤U­±¦^°é¤º§PÂ_¤~¯à¼u°T®§
'      If Not m_bEditDoc Then 'Add by Morgan 2010/11/25
'         'Add by Morgan 2010/5/13
'         adoacc1k0.Close
'         StrSQLa = "select 1 from acc1k0 where (a1k12 is null or a1k12 = 0) and a1k32 is not null" & strSql
'         adoacc1k0.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'         If adoacc1k0.RecordCount <> 0 Then
'             strCon10 = MsgText(602)
'             MsgBox "¯S®í½Ð´Ú³æ¤£¥i¦C¦L¡I"
'             adoacc1k0.Close
'             Exit Sub
'         End If
'         'end 2010/5/13
'      End If
      
      '92.10.6 ADD BY SONIA
      adoacc1k0.Close
      'Modify By Sindy 2011/3/7 +fa108
      StrSQLa = "select a1k27, a1l07, a1l05, a1j04, a1l06, a1j05, a1j06, a1j16, a1j10, fa05, fa63, fa64, fa65, fa32, fa18, a1k02, fa33, fa19, fa34, fa20, a1k13, a1k14, a1k15, a1k16, fa21, fa22, fa35, a1k03, fa36, a1k01, fa06, fa23, a1k04, a1k10, a1l02, fa43, a1k18 as Curr, '' as cu102, a1j02, a1k05, FA04, FA17,fa108 from acc1k0, acc1l0, fagent, acc1j0 where a1k01 = a1l01 (+) and substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and a1l03 = a1j01 (+) and a1l04 = a1j02 (+) and a1k29 is NOT null and (a1k12 is null or a1k12 = 0)" & strSql & " union " & _
                  "select a1k27, a1l07, a1l05, a1j04, a1l06, a1j05, a1j06, a1j16, a1j10, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, a1k02, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, a1k13, a1k14, a1k15, a1k16, cu27 as fa21, cu28 as fa22, cu68 as fa35, a1k03, cu69 as fa36, a1k01, cu06 as fa06, cu29 as fa23, a1k04, a1k10, a1l02, cu76 as fa43, a1k18 as Curr, cu102, a1j02, a1k05, CU04 As FA04, CU23 As FA17,cu148 as fa108 From acc1k0, acc1l0, customer, acc1j0 where a1k01 = a1l01 (+) and substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) = cu02 and a1l03 = a1j01 (+) and a1l04 = a1j02 (+) and a1k29 is NOT null and (a1k12 is null or a1k12 = 0)" & strSql & " order by a1k01 asc, a1k03 asc, a1l02 asc"
      adoacc1k0.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc1k0.RecordCount <> 0 Then
          strCon10 = MsgText(602)
          MsgBox MsgText(147), , MsgText(5)
          adoacc1k0.Close
          Exit Function
      Else
      '92.10.6 END
          strCon10 = MsgText(602)
          MsgBox MsgText(28), , MsgText(5)
          adoacc1k0.Close
          Exit Function
      End If
   Else
      InsertQueryLog (adoacc1k0.RecordCount) 'Add By Sindy 2010/12/22
      adoTaie.Execute "Delete From ACCRPT428 Where R42801='" & strUserNum & "' "
      strSpecialList = "" 'Add by Morgan 2011/10/11
      strSpecialL2 = "" 'Added by Lydia 2015/04/15
      While Not adoacc1k0.EOF
          'Add by Morgan 2011/10/11
          'Modified by Morgan 2016/7/27 LEDES¹q¤l±b³æ­n¯à²£¥Í
          If Not m_bEditDoc And Check2.Value <> vbChecked Then
             'Added by Lydia 2015/04/15 + ¾ã§å½Ð´Ú³æ(a1k32=C)
             'If adoacc1k0.Fields("a1k32") = "Y" Then
             If Not IsNull(adoacc1k0.Fields("a1k32")) Then
                If strA1K01 <> adoacc1k0.Fields("a1k01") Then
                   If adoacc1k0.Fields("a1k32") = "Y" Then
                      strSpecialList = strSpecialList & adoacc1k0.Fields("a1k01") & vbCrLf
                   Else
                      strSpecialL2 = strSpecialL2 & adoacc1k0.Fields("a1k01") & vbCrLf
                   End If
                  'end 2015/04/15
                   strA1K01 = adoacc1k0.Fields("a1k01")
                End If
                GoTo NextRec
             End If
          End If
         'end 2011/10/11
         
         'Add By Sindy 2012/12/27 98µ²§Àªº¬°¥N¦¬¥N¥I,ª½±µ±Nª÷ÃB¥[¨ì¦P¶µ¥Ø
         'Modify By Sindy 2013/5/7
         bolChkIsACCRPT428Exists = False
         If Right(Trim("" & adoacc1k0.Fields("a1l04").Value), 2) = "98" Then 'Added by Morgan 2023/11/17 98µ²§Àªº¤~»Ý­nÀË¬d
            'Modified by Morgan 2024/7/9 + order by R42836 desc (½Ð´Ú¶µ¥Ø¥i¯à·|­«½Æ,§ì³Ì¤jªº§Ç¸¹§ó·s Ex:X11309679)
            strExc(0) = "select * from ACCRPT428 where R42801='" & strUserNum & _
                      "' and R42831='" & adoacc1k0.Fields("a1k01").Value & _
                      "' and R42840='" & adoacc1k0.Fields("a1j02").Value & _
                      "' order by R42836 desc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               bolChkIsACCRPT428Exists = True
               stR42836 = RsTemp("R42836") 'Added by Morgan 2024/7/9
            End If
         End If 'Added by Morgan 2023/11/17
         
         If Right(Trim("" & adoacc1k0.Fields("a1l04").Value), 2) = "98" And bolChkIsACCRPT428Exists = True Then
         '2013/5/7 End
            '­pºâ¥~¹ôª÷ÃB R42856,R42857
            dblR42856 = GetDebitNoteFAmt("" & adoacc1k0.Fields("a1k01"), _
                                         "" & adoacc1k0.Fields("Curr").Value, _
                                         "" & adoacc1k0.Fields("a1k02").Value, _
                                         "" & adoacc1k0.Fields("a1l04").Value, _
                                         "" & adoacc1k0.Fields("a1l05").Value, _
                                         "" & adoacc1k0.Fields("a1l07").Value, _
                                         "" & adoacc1k0.Fields("a1l16").Value, _
                                         "" & adoacc1k0.Fields("a1l17").Value, _
                                         "" & adoacc1k0.Fields("a1k33").Value, _
                                         "" & adoacc1k0.Fields("a1k10").Value, _
                                         dblR42857)
            '" And R42840='" & Left(Trim("" & adoacc1k0.Fields("a1l04").Value), Len(Trim("" & adoacc1k0.Fields("a1l04").Value)) - 2) & "' "
            'Modified by Morgan 2024/7/9 ½Ð´Ú¶µ¥Ø¥i¯à·|­«½Æ,§ì³Ì¤jªº§Ç¸¹§ó·s Ex:X11309679
            StrSQLa = "Update ACCRPT428 Set R42803=R42803+" & Val("" & adoacc1k0.Fields("a1l07").Value) & _
                      ", R42804=R42804+" & Val("" & adoacc1k0.Fields("a1l05").Value) & _
                      ", R42856=R42856+" & dblR42856 & _
                      ", R42857=R42857+" & dblR42857 & _
                      " Where R42801='" & strUserNum & "'" & _
                      " And R42831='" & adoacc1k0.Fields("a1k01").Value & "'" & _
                      " And R42840='" & adoacc1k0.Fields("a1j02").Value & "' and R42836='" & stR42836 & "'"
            adoTaie.Execute StrSQLa, intI
            
         Else
         '2012/12/27 End
            intI = 0
            '­Y¨t²ÎÃþ§O¬°T, ­Y¦³03³W¶O®É¤£¦L, ª½±µ±Nª÷ÃB¥[¨ì01¶µ¥Ø
            'Modify by Morgan 2007/4/19 T¶}ÀYªº³£­n
            'If "" & adoacc1k0.Fields("a1k13").Value = "T" And "" & adoacc1k0.Fields("a1j02").Value = "03" Then
            If Left("" & adoacc1k0.Fields("a1k13").Value, 1) = "T" And "" & adoacc1k0.Fields("a1l04").Value = "03" Then
               'Add By Sindy 2013/3/29
               '­pºâ¥~¹ôª÷ÃB R42856,R42857
               dblR42856 = GetDebitNoteFAmt("" & adoacc1k0.Fields("a1k01"), _
                                            "" & adoacc1k0.Fields("Curr").Value, _
                                            "" & adoacc1k0.Fields("a1k02").Value, _
                                            "" & adoacc1k0.Fields("a1l04").Value, _
                                            "" & adoacc1k0.Fields("a1l05").Value, _
                                            "" & adoacc1k0.Fields("a1l07").Value, _
                                            "" & adoacc1k0.Fields("a1l16").Value, _
                                            "" & adoacc1k0.Fields("a1l17").Value, _
                                            "" & adoacc1k0.Fields("a1k33").Value, _
                                            "" & adoacc1k0.Fields("a1k10").Value, _
                                            dblR42857)
               '2013/3/29 End
               StrSQLa = "Update ACCRPT428 Set R42803=R42803+" & Val("" & adoacc1k0.Fields("a1l07").Value) & _
                         ", R42804=R42804+" & Val("" & adoacc1k0.Fields("a1l05").Value) & _
                         ", R42856=R42856+" & dblR42856 & _
                         ", R42857=R42857+" & dblR42857 & _
                         " Where R42801='" & strUserNum & "'" & _
                         " And R42831='" & adoacc1k0.Fields("a1k01").Value & "'" & _
                         " And R42840='01'"
               adoTaie.Execute StrSQLa, intI
               
            'Modify by Morgan 2007/7/4 ¨S§ó·s¨ìªº§ï·s¼W
            'Else
            End If
            If intI = 0 Then
            'end 2007/7/4
               StrSQLa = "(" & CNULL(strUserNum) & ","
               m_tmp = "R42801"
               'Modify by Morgan 2011/10/11 a1k32 ¤£¥²¼g¼È¦s
               'For ii = 0 To adoacc1k0.Fields.Count - 1
               'Modify by Sindy 2013/3/28 a1k33 ¤£¥²¼g¼È¦s
               'For ii = 0 To adoacc1k0.Fields.Count - 2
               For ii = 0 To 53
                   StrSQLa = StrSQLa & CNULL(ChgSQL("" & adoacc1k0.Fields(ii).Value)) & ","
                   m_tmp = m_tmp & ",R428" & Format(ii + 2, "0#")
               Next ii
               'Add By Sindy 2013/3/29
               '­pºâ¥~¹ôª÷ÃB R42856,R42857
               'Added by Morgan 2019/8/30 BASF Â½Ä¶¶O ¬üª÷­nÅã¥Ü¦Ü¤p¼Æ²Ä2¦ì
               'Modified by Morgan 2022/2/18 +927¨ä¥LÂ½Ä¶ Ex:X11102382-- Ryan
               'Modified by Morgan 2022/9/2 +209ÀËµø¤¤»¡-- Tim
               'Modified by Morgan 2025/10/31 +FG Ex:X11404796
               If (adoacc1k0.Fields("a1k28") = "Y45814010" Or adoacc1k0.Fields("a1k28") = "Y33268010") And (adoacc1k0.Fields("a1k13") = "FCP" Or adoacc1k0.Fields("a1k13") = "FG" Or adoacc1k0.Fields("a1k13") = "P" Or adoacc1k0.Fields("a1k13") = "CFP") And (adoacc1k0.Fields("a1l04") = "201" Or adoacc1k0.Fields("a1l04") = "927" Or adoacc1k0.Fields("a1l04") = "209") Then
                  'Modify By Sindy 2025/4/2 ¶}©ñ¥i¥H¿é¤J¹ô§O,¹ô§O¬Û¦P¤S¨S¦³§é¦©°ÝÃD,¤£¥Î¦A´«ºâª½±µ¨Ï¥ÎA1L17
                  If "" & adoacc1k0.Fields("a1L16").Value = "" & adoacc1k0.Fields("Curr").Value And Val("" & adoacc1k0.Fields("a1L07")) = 0 Then
                     dblR42856 = Val("" & adoacc1k0.Fields("a1l17"))
                     dblR42857 = Val("" & adoacc1k0.Fields("a1l17"))
                  Else
                  '2025/4/2 END
                     dblR42856 = Format((Val("" & adoacc1k0.Fields("a1l05")) - Val("" & adoacc1k0.Fields("a1l07"))) / Val("" & adoacc1k0.Fields("a1k10")), "#.00")
                     dblR42857 = Format(Val("" & adoacc1k0.Fields("a1l05")) / Val("" & adoacc1k0.Fields("a1k10")), "#.00")
                  End If
               Else
               'end 2019/8/30
             
                  dblR42856 = GetDebitNoteFAmt("" & adoacc1k0.Fields("a1k01"), _
                                               "" & adoacc1k0.Fields("Curr").Value, _
                                               "" & adoacc1k0.Fields("a1k02").Value, _
                                               "" & adoacc1k0.Fields("a1l04").Value, _
                                               "" & adoacc1k0.Fields("a1l05").Value, _
                                               "" & adoacc1k0.Fields("a1l07").Value, _
                                               "" & adoacc1k0.Fields("a1l16").Value, _
                                               "" & adoacc1k0.Fields("a1l17").Value, _
                                               "" & adoacc1k0.Fields("a1k33").Value, _
                                               "" & adoacc1k0.Fields("a1k10").Value, _
                                               dblR42857)
                                               
               End If 'Added by Morgan 2019/8/30
               StrSQLa = StrSQLa & dblR42856 & "," & dblR42857
               m_tmp = m_tmp & ",R42856,R42857"
               '2013/3/29 End
               'StrSQLa = Left(StrSQLa, Len(StrSQLa) - 1)
               StrSQLa = StrSQLa & ")"
               'Modify by Morgan 2006/9/26 §ï«ü©wÄæ¦ì¦WºÙ,³o¼Ë·s¼WÄæ¦ì®É¤~¤£·|¦³¿ù
               'StrSQLa = "Insert Into ACCRPT428 Values " & StrSQLa
               StrSQLa = "Insert Into ACCRPT428 (" & m_tmp & ") Values " & StrSQLa
               adoTaie.Execute StrSQLa, intI
               
            End If
         End If
NextRec:
         adoacc1k0.MoveNext
      Wend
      
      If adoacc1k0.State <> adStateClosed Then adoacc1k0.Close
      Set adoacc1k0 = Nothing
   End If
   
   PrintData_1 = True 'Added by Morgan 2024/11/1
   
End Function

'Added by Lydia 2020/09/08 §PÂ_¯S®í«È¤á¡þ®æ¦¡
Private Sub PrintData_2(ByVal pA1k01 As String, ByRef pstrDNMemoAlertList As String)
         m_bSpecial1 = False
         m_bSpecial2 = False
         m_bSpecial3 = False
         m_bSpecial4 = False 'Added by Morgan 2014/2/24
         m_bSpecial5 = False 'Added by Lydia 2016/03/03
         bolNewForm = False 'Added by Morgan 2013/10/31
         m_bSpecialNew1 = False 'Added by Morgan 2014/2/18
         m_bSpecialNew2 = False 'Added by Morgan 2016/9/9
         m_bSpecialNew3 = False 'Added by Morgan 2018/3/22
         m_bSpecialNew4 = False 'Added by Morgan 2018/4/11
         m_bSpecialNew5 = False 'Added by Morgan 2024/3/25
         m_Activity = "" 'Added by Morgan 2019/2/27
         m_bDowN = False 'Added by Morgan 2020/8/6
         
         'Modified by Morgan 2013/10/31 §ï¦@¥Î,½Ð´Ú§@·~¤]­n§PÂ_
         'If Left(m_A1k28, 8) = "Y2245700" Then
         '   m_bSpecial1 = True
         ''Added by Morgan 2012/6/13 +¦L«È¤á½Ð´Ú¶µ¥Ø¥N½X
         'ElseIf InStr("Y5349500,X3224200,Y4725001", Left(m_A1k28, 8)) > 0 Then
         '   m_bSpecial2 = True
         ''Modified by Morgan 2012/11/12
         ''Modified by Morgan 2013/1/22 +Y4830907 °Ó¼Ð½Ð´Ú
         'ElseIf InStr("X5208400,X4831001,X4831000,Y4830906,Y4830907", Left(m_A1k28, 8)) > 0 Then
         '   m_bSpecial3 = True
         'End If
         intI = PUB_GetBillFormat(m_A1k28, m_CP01, m_CP02, m_CP03, m_CP04)
         If intI = 1 Then
            m_bSpecial1 = True
            m_bDowX = ChkDowXFormat(pA1k01) 'Added by Morgan 2017/8/17
            'Added by Morgan 2020/8/6
            'Dow §é¦©¤£¥t¦C
            'Modified by Morgan 2022/3/4 +Y55423 --Kimi
            'Modified by Morgan 2022/3/11 -Y55423 --Kimi
            'Modified by Morgan 2025/6/30 +Y52322B10 --Tim
            If InStr("Y22457000,Y48048000,Y52322000,Y52322B10", m_A1k28) > 0 Then
               m_bDowN = True
            End If
            'end 2020/8/6
         ElseIf intI = 2 Then
            m_bSpecial2 = True
         ElseIf intI = 3 Then
            m_bSpecial3 = True
         'Added by Morgan 2014/2/24
         ElseIf intI = 4 Then
            m_bSpecial4 = True
         'Added by Lydia 2016/03/03
         ElseIf intI = 5 Then
            m_bSpecial5 = True
         'Added by Morgan 2013/12/4 ½Ð´Ú¤é´Á>=1021205§ï¥Î·s®æ¦¡
         ElseIf Val(m_strA1K02) >= 1021205 Then
            bolNewForm = True
         End If
         'end 2013/12/4
         
         'Added by Morgan 2014/2/18 BASF SE ®æ¦¡¯S§O
         'Modified by Morgan 2017/6/22 +Y54179000 Longitude Licensing Ltd -- ³¯¨Ø­s
         'Modified by Morgan 2019/2/27 +Y48904000 Advanced Energy Industries, Inc.
         'Modfiied by Morgan 2021/7/1 +Y3326801 BASF Corporation --Anny
         If InStr("Y45814010,Y54179000,Y48904000,Y33268010", m_A1k28) > 0 Then
            m_bSpecialNew1 = True
         'Added by Morgan 2016/8/5 SOEI PATENT & LAW FIRM ¯S®í®æ¦¡,©ú²Ó¦P®É­n¦C¥x¹ô¤Î¬üª÷ --³¢©É¼ü
         ElseIf m_A1k28 = "Y45204000" Then
            m_bSpecialNew2 = True
         'Added by Morgan 2018/3/22
         'Modified by Morgan 2025/7/10 +Y4794600,Y47946B1,Y47946B2 -- Tim
         'Modified by Morgan 2025/8/4 +Y47946B3 --Tim
         ElseIf InStr("Y27840000,Y27840B10,Y4794600,Y47946B1,Y47946B2,Y47946B3", Left(m_A1k28, 8)) > 0 Then
            m_bSpecialNew3 = True
         'Added by Morgan 2018/4/11
         ElseIf m_A1k03 = "Y27696000" And m_iPrintCurrType = 3 Then
            m_bSpecialNew4 = True
         'Added by Morgan 2024/3/25
         ElseIf m_A1k03 = "Y20049000" Then
            
            strCust1 = GetPrjPeopleNum1(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04)
            If (m_CP01 = "P" And (strCust1 = "X47325C10" Or strCust1 = "X47325C12")) Or (m_CP01 = "FCP" And (strCust1 = "X47325000" Or strCust1 = "X47325004")) Then
               'Added by Morgan 2024/12/24
               'Removed by Morgan 2024/12/25 §ï¦^,ºû«ù¾ã±i³£¨S¦³¤uµ{®v©Ó¿ìªº½Ð´Ú¶µ¥Ø®É¶]¤@¯ë®æ¦¡
               'm_EngCP10List = ""
               'm_bSpecialNew5 = True
               'end 2024/12/25
               'end 2024/12/24
               strExc(0) = "select cp10 from acc1l0,caseprogress,staff where a1l01='" & adoacc1k0("a1k01") & "' and cp60(+)=a1l01 and cp10(+)=a1l04 and st01(+)=cp14 and st03='F21'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  m_EngCP10List = "," & RsTemp.GetString(, , , ",")
                  m_bSpecialNew5 = True 'Removed by Morgan 2024/12/24 'Added by Morgan 2024/12/25 §ï¦^,ºû«ù¾ã±i³£¨S¦³¤uµ{®v©Ó¿ìªº½Ð´Ú¶µ¥Ø®É¶]¤@¯ë®æ¦¡
               End If
            End If
         End If
         'end 2014/2/18
         
         m_PlusFormNo = GetPlusFormNo(m_A1k28, m_CP01, m_CP02, m_CP03, m_CP04) 'Added by Morgan 2014/8/26
         
         'Add by Morgan 2008/6/11 D/N³Æµù´£¿ô
         If Not m_bBeCalled And InStr(pstrDNMemoAlertList, m_A1k28) = 0 Then
            'Modify by Morgan 2011/2/10
            'PUB_CheckDNMemo m_A1k28
            If Pub_StrUserSt03 = "F12" Then
               'Modify By Sindy 2011/3/3
               'If PUB_CheckDNMemo(m_A1k28, True) = False Then
               If PUB_CheckDNMemo(m_A1k28, True, m_CP01) = False Then
               '2011/3/3 End
                  Printer.KillDoc
                  If adoacc1k0.State <> adStateClosed Then adoacc1k0.Close
                  Exit Sub
               End If
            Else
               'Modify By Sindy 2011/3/3
               'PUB_CheckDNMemo m_A1k28
               PUB_CheckDNMemo m_A1k28, , m_CP01
            End If
            'end 2011/2/10
            pstrDNMemoAlertList = pstrDNMemoAlertList & m_A1k28 & ";"
         End If
         
End Sub
'Added by Morgan 2021/1/14 §ó·s½Ð´Ú³æ¬üª÷Á`ÃB
Private Sub UpdateA1K38(pA1k01 As String, pA1K38 As String)
   Dim intR As Integer
   cnnConnection.Execute "update acc1k0 set a1k38=" & pA1K38 & " where a1k01='" & pA1k01 & "'", intR
End Sub

'Add By Sindy 2021/10/29
Private Sub UpdateA1K3940(dblAttFeeSub As Double, dblOffFeeSub As Double)
   'ÀË¬d©ú²Ó¥~¹ôª÷ÃB¬O²Å¦X¥~¹ôÁ`ª÷ÃB¶Ü?¤£²Å¶·­n­«·s­pºâ
   If m_A1k08 <> (dblOffFeeSub + dblAttFeeSub) Then
      strExc(0) = "select sum(R42856) from ACCRPT428 where R42801='" & strUserNum & "'" & _
                  " and R42831='" & m_strDN & "' and substr(R42840,-2)=99"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.Fields(0) > 0 Then
            dblOffFeeSub = Trunc(RsTemp.Fields(0))
         End If
      End If
      dblAttFeeSub = m_A1k08 - dblOffFeeSub
   End If
   '°O¿ý ½Ð´ÚªA°È¶O,³W¶O¥~¹ôª÷ÃB
   strSql = "Update ACC1K0 Set A1K39=" & dblAttFeeSub & ",A1K40=" & dblOffFeeSub & " Where A1K01='" & m_strDN & "'"
   adoTaie.Execute strSql, intI
End Sub

'Added by Morgan 2022/8/4 ¦]µ{§Ç¤Ó¤j¡A§ï¼g¦¨¨ç¼Æ¦C¦L´î¤Öµ{¦¡½X
Private Sub PictureText(pText As String, Px As Long, Py As Long)
   Picture1.CurrentX = Px * douExtRate
   Picture1.CurrentY = Py * douExtRate
   Picture1.Print pText
End Sub

'Added by Morgan 2024/2/20
'¤¤¤å½Ð´Ú³æµL¤¤¤å¦W®É§ï§ì­^¤å
Private Sub SetEngName(pRow As Integer)
   Dim intLRow As Integer
   
   intLRow = pRow
   If Not IsNull(adoacc1k0.Fields("fa63")) Then
      intLRow = intLRow + 1
      SetWordArray m_Head, intLRow, 2, "" & adoacc1k0.Fields("fa63").Value
   End If
   If Not IsNull(adoacc1k0.Fields("fa64")) Then
      intLRow = intLRow + 1
      SetWordArray m_Head, intLRow, 2, "" & adoacc1k0.Fields("fa64").Value
   End If
   If Not IsNull(adoacc1k0.Fields("fa65")) Then
      intLRow = intLRow + 1
      SetWordArray m_Head, intLRow, 2, "" & adoacc1k0.Fields("fa65").Value
   End If
   '½Ð´Ú¹ï¶HY51817010¤¤¤å®æ¦¡½Ð´Ú³æ­n¥[¦L­^¤å¦a§}--®Û­^
   'Modified by Morgan 2024/2/20 +Y51817010 ¨Ã§ï§PÂ_¦C¦L¹ï¶H--®Û­^
   'Modified by Morgan 2025/5/28 +Y52459040 --®Û­^
   If m_A1k27 = "Y51817010" Or m_A1k27 = "Y52459030" Or m_A1k27 = "Y52459040" Then
      If IsNull(adoacc1k0.Fields("fa32").Value) = False Then
         intLRow = intLRow + 1
         m_tmp = adoacc1k0.Fields("fa32").Value
         SetWordArray m_Head, intLRow, 2, m_tmp
         If IsNull(adoacc1k0.Fields("fa33").Value) = False Then
            intLRow = intLRow + 1
            m_tmp = adoacc1k0.Fields("fa33").Value
            SetWordArray m_Head, intLRow, 2, m_tmp
         End If
         If IsNull(adoacc1k0.Fields("fa34").Value) = False Then
            intLRow = intLRow + 1
            m_tmp = adoacc1k0.Fields("fa34").Value
            SetWordArray m_Head, intLRow, 2, m_tmp
         End If
         If IsNull(adoacc1k0.Fields("fa35").Value) = False Then
            intLRow = intLRow + 1
            m_tmp = adoacc1k0.Fields("fa35").Value
            SetWordArray m_Head, intLRow, 2, m_tmp
         End If
         If IsNull(adoacc1k0.Fields("fa36").Value) = False Then
            intLRow = intLRow + 1
            m_tmp = adoacc1k0.Fields("fa36").Value
            SetWordArray m_Head, intLRow, 2, m_tmp
         End If
         
      ElseIf IsNull(adoacc1k0.Fields("fa18").Value) = False Then
         intLRow = intLRow + 1
         m_tmp = adoacc1k0.Fields("fa18").Value
         SetWordArray m_Head, intLRow, 2, m_tmp
         If IsNull(adoacc1k0.Fields("fa19").Value) = False Then
            intLRow = intLRow + 1
            m_tmp = adoacc1k0.Fields("fa19").Value
            SetWordArray m_Head, intLRow, 2, m_tmp
         End If
         If IsNull(adoacc1k0.Fields("fa20").Value) = False Then
            intLRow = intLRow + 1
            m_tmp = adoacc1k0.Fields("fa20").Value
            SetWordArray m_Head, intLRow, 2, m_tmp
         End If
         If IsNull(adoacc1k0.Fields("fa21").Value) = False Then
            intLRow = intLRow + 1
            m_tmp = adoacc1k0.Fields("fa21").Value
            SetWordArray m_Head, intLRow, 2, m_tmp
         End If
         If IsNull(adoacc1k0.Fields("fa22").Value) = False Then
            intLRow = intLRow + 1
            m_tmp = adoacc1k0.Fields("fa22").Value
            SetWordArray m_Head, intLRow, 2, m_tmp
         End If
      End If
   End If
End Sub
'Added by Morgan 2024/3/26
'¥~¹ô½Ð´Úª÷ÃB
Private Function GetA1L17(pA1L01 As String, pA1L04 As String) As Double
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   'Modified by Morgan 2024/5/28 ³°¥N¶O¥Î¤£¥´§é--Kimi
   'stSQL = "select round(a1l17*a1l19,0) from acc1l0 where a1l01='" & pA1L01 & "' and a1l04='" & pA1L04 & "' and a1l17>0"
   stSQL = "select a1l17 from acc1l0 where a1l01='" & pA1L01 & "' and a1l04='" & pA1L04 & "' and a1l17>0"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      GetA1L17 = rsQuery(0)
   End If
   Set rsQuery = Nothing
End Function

'Added by Morgan 2024/5/9
Private Sub Set201ItemDesc()
   'Added by Morgan 2018/3/8
   'Y45814010 BASF Â½Ä¶¶O­n¦L­ì¤å¦r¼Æ
   'Modified by Morgan 2019/6/28 +Y54225B10 Syngenta ªº½Ð´Ú³æ¡]LEDES¡^½Ð´Ú¶µ¥Ø209ÀËµø¤¤»¡ªº­^¤å±Ô­z«á¤è¥[¤W+­^¤å¦r¼Æ
   'Modified by Morgan 2021/7/1 +Y3326801 BASF Corporation --Anny
   'Modified by Morgan 2021/9/23 +Y45898000 CABINET HECKE (SAS) ªº 201 & 209 --Franny
   'Modified by Morgan 2022/1/6 +Y55483 Vivek Singh ªº 201 & 209 --Kahn
   'Modified by Morgan 2024/5/9 +Y19893030 Fresenius Medical Care AG ªº 201 --Izumi
   'Modified by Morgan 2025/3/14 +Y56142000--Franny
   If (adoacc1k0.Fields("a1k13").Value = "FCP" Or adoacc1k0.Fields("a1k13").Value = "P") And ( _
         ((m_A1k28 = "Y45814010" Or m_A1k28 = "Y33268010") And adoacc1k0.Fields("a1j02") = "201") _
      Or (m_A1k28 = "Y54225B10" And adoacc1k0.Fields("a1j02") = "209") _
      Or ((m_A1k28 = "Y45898000" Or m_A1k28 = "Y55483000") And (adoacc1k0.Fields("a1j02") = "201" Or adoacc1k0.Fields("a1j02") = "209")) _
      Or ((m_A1k28 = "Y19893030" Or m_A1k28 = "Y56142000") And adoacc1k0.Fields("a1j02") = "201")) Then
      'Modified by Morgan 2023/6/20 +tf19
      strExc(0) = "select tf23,tf19 from caseprogress,transfee where cp60='" & adoacc1k0.Fields("a1k01") & "' and cp10='" & adoacc1k0.Fields("a1j02") & "' and tf01(+)=cp09"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Modified by Morgan 2019/6/28 +Y54225B10¤§±b³æ¡]LEDES±b³æ¡^¨ä½Ð´Ú¶µ¥Ø209ÀËµø¤¤»¡­^¤å±Ô­z«á¤è¥[¤W+­^¤å¦r¼Æ
         If m_A1k28 = "Y54225B10" And adoacc1k0.Fields("a1j02") = "209" Then
            'Modified by Morgan 2019/8/22 + Unit Cost: 0.07 --Franny
            m_ItemDesc = m_ItemDesc & ": Unit Cost: 0.07, English Words: " & Format(RsTemp(0), "#,###")
         'Added by Morgan 2022/1/6
         ElseIf m_A1k28 = "Y55483000" Then
            If adoacc1k0.Fields("a1j02") = "201" Then
               m_ItemDesc = m_ItemDesc & "(English Words: " & Format(RsTemp(0), "#,###") & " / USD 15 per 100 English words)"
            Else
               m_ItemDesc = m_ItemDesc & "(English Words: " & Format(RsTemp(0), "#,###") & ")"
            End If
         'end 2022/1/6
         'Added by Morgan 2023/6/20  BASF¦³¬Û¦ü«×­n±a¹ê»Ú­pºâª÷ÃBªº½Ð´Ú­^¤å¦r¼Æ--Franny
         ElseIf RsTemp("tf19") > 0 And (m_A1k28 = "Y45814010" Or m_A1k28 = "Y3326801") Then
            m_ItemDesc = m_ItemDesc & "(Adjusted: " & Trunc(RsTemp("tf23") * (100 - RsTemp("tf19")) / 100) & " words)"
         'end 2023/6/20
         'Added by Morgan 2024/5/9
         'Modified by Morgan 2025/3/14 +Y56142000--Franny
         ElseIf m_A1k28 = "Y19893030" Or m_A1k28 = "Y56142000" Then
            If RsTemp("tf19") > 0 Then
               m_ItemDesc = m_ItemDesc & "(" & Trunc(RsTemp("tf23") * (100 - RsTemp("tf19")) / 100) & " words)"
            Else
               m_ItemDesc = m_ItemDesc & "(" & RsTemp(0) & " words)"
            End If
         'end 2024/5/9
         Else
            m_ItemDesc = m_ItemDesc & "(Total: " & RsTemp(0) & " words)"
         End If
      End If
   End If
   'end 2018/3/8
   
   'Added by Morgan 2024/5/9
   'Modified by Morgan 2025/3/14 +Y56142000--Franny
   If (adoacc1k0.Fields("a1k13").Value = "FCP" Or adoacc1k0.Fields("a1k13").Value = "P") And (m_A1k28 = "Y19893030" Or m_A1k28 = "Y56142000") And adoacc1k0.Fields("a1j02") = "03" Then
      strExc(0) = "select nvl(pa64,0)+nvl(pa65,0)+nvl(pa67,0) pages,tf19 from caseprogress,transfee,patent" & _
         " where cp60='" & adoacc1k0.Fields("a1k01") & "' and cp10='201' and tf01(+)=cp09" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp(0) > 0 Then
            If RsTemp("tf19") > 0 Then
               m_ItemDesc = m_ItemDesc & "(" & Trunc(RsTemp(0) * (100 - RsTemp("tf19")) / 100) & " pages)"
            Else
               m_ItemDesc = m_ItemDesc & "(" & RsTemp(0) & " pages)"
            End If
         End If
      End If
   End If
   'end 2024/5/9
End Sub

'Added by Morgan 2025/1/16
'ÀË¬d¬O§_¬°¤À³Î®×¥B¥t¦³³]©wTASK_CODE(¥Ø«e¥uÀË¬dP,FCP)
Private Sub SetDivCaseCode(cp01 As String, cp02 As String, cp03 As String, cp04 As String, A1K27 As String, A1L04 As String, ByRef pTASK_CODE As String)
   Dim cp(4) As String
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   If cp01 = "P" Or cp01 = "FCP" Then
      cp(1) = cp01
      cp(2) = cp02
      cp(3) = cp03
      cp(4) = cp04
      If PUB_ChkCPExist(cp(), "307") Then
         stSQL = "select a2623 from acc260 where a2601='" & Left(A1K27, 8) & "' and a2602='" & cp01 & "' and a2603='" & A1L04 & "' and a2623 is not null"
         intQ = 1
         Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
         If intQ = 1 Then
            pTASK_CODE = rsQuery(0)
         End If
      End If
   End If
   Set rsQuery = Nothing
End Sub
