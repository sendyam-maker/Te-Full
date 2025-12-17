VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc14o0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "°ê¤º¦¬¾Ú²£¥ÍINVOICE"
   ClientHeight    =   5400
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8988
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8988
   Begin VB.Frame Frame1 
      Caption         =   "­«¶]INVOICE"
      Height          =   588
      Left            =   5520
      TabIndex        =   35
      Top             =   4632
      Width           =   3252
      Begin VB.CommandButton cmdRun 
         Caption         =   "Word"
         Height          =   300
         Left            =   2184
         TabIndex        =   38
         Top             =   216
         Width           =   876
      End
      Begin VB.TextBox Text2 
         Height          =   324
         Left            =   1272
         MaxLength       =   6
         TabIndex        =   37
         Top             =   192
         Width           =   732
      End
      Begin VB.Label Label2 
         Caption         =   "INVOICE½s¸¹:"
         Height          =   252
         Left            =   48
         TabIndex        =   36
         Top             =   240
         Width           =   1236
      End
   End
   Begin VB.ComboBox Combo4 
      Height          =   276
      ItemData        =   "Frmacc14o0.frx":0000
      Left            =   7032
      List            =   "Frmacc14o0.frx":000D
      Style           =   2  '³æ¯Â¤U©Ô¦¡
      TabIndex        =   34
      Top             =   829
      Width           =   1872
   End
   Begin VB.TextBox txtCU196 
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      MaxLength       =   1
      TabIndex        =   33
      Top             =   810
      Width           =   408
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '¥­­±
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '¨S¦³®Ø½u
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   405
      Left            =   4950
      MultiLine       =   -1  'True
      TabIndex        =   31
      Text            =   "Frmacc14o0.frx":002D
      Top             =   90
      Width           =   3915
   End
   Begin VB.ComboBox Combo3 
      Height          =   276
      ItemData        =   "Frmacc14o0.frx":0078
      Left            =   4710
      List            =   "Frmacc14o0.frx":0085
      Style           =   2  '³æ¯Â¤U©Ô¦¡
      TabIndex        =   3
      Top             =   1500
      Width           =   1770
   End
   Begin VB.TextBox txtFee 
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4710
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1860
      Width           =   1050
   End
   Begin VB.TextBox txtRate2 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7635
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1860
      Width           =   1050
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "Frmacc14o0.frx":00A5
      Left            =   1140
      List            =   "Frmacc14o0.frx":00A7
      Style           =   2  '³æ¯Â¤U©Ô¦¡
      TabIndex        =   11
      Top             =   4650
      Width           =   1770
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "²M°£"
      Height          =   285
      Left            =   1860
      TabIndex        =   9
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdDelRow 
      Caption         =   "§R°£"
      Height          =   285
      Left            =   1035
      TabIndex        =   8
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdAddRow 
      Caption         =   "¥[¤J"
      Height          =   285
      Left            =   210
      TabIndex        =   7
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   2460
      Picture         =   "Frmacc14o0.frx":00A9
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   1
      Top             =   120
      Width           =   350
   End
   Begin VB.CommandButton cmdWord 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Word(&W)"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3240
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   12
      Top             =   4590
      Width           =   2196
   End
   Begin VB.TextBox txtNo 
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1260
      MaxLength       =   9
      TabIndex        =   0
      Top             =   90
      Width           =   1170
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      ItemData        =   "Frmacc14o0.frx":01AB
      Left            =   1710
      List            =   "Frmacc14o0.frx":01AD
      Style           =   2  '³æ¯Â¤U©Ô¦¡
      TabIndex        =   2
      Top             =   1500
      Width           =   1770
   End
   Begin VB.TextBox txtRate 
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7635
      TabIndex        =   4
      Top             =   1500
      Width           =   1050
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   2265
      Left            =   90
      TabIndex        =   10
      Top             =   2250
      Width           =   8775
      _ExtentX        =   15473
      _ExtentY        =   4001
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      FormatString    =   "V|¦¬¾Ú½s¸¹|¦¬¾Ú¤é´Á|¦¬¾Úª÷ÃB|«È¤á½s¸¹|´¼Åv¤H­û|¦¬¾Ú©ïÀY"
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
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "½Ð´Ú¶×¤J»È¦æ¸ê®Æ¡G"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   9
      Left            =   4824
      TabIndex        =   32
      Top             =   870
      Width           =   1836
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "ª÷ÃBÅã¥Ü¡G"
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
      Index           =   8
      Left            =   3660
      TabIndex        =   30
      Top             =   1560
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "¤âÄò¶O¡G"
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
      Index           =   7
      Left            =   3870
      TabIndex        =   29
      Top             =   1920
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "´«ºâ«á¶×²v¡G"
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
      Index           =   6
      Left            =   6300
      TabIndex        =   28
      Top             =   1920
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "®M¥Î«HÀY¡G"
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
      Index           =   5
      Left            =   120
      TabIndex        =   27
      Top             =   4710
      Width           =   1050
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '³z©ú
      Caption         =   "ª`·N¡G¦b²£¥ÍINVOICE®É¡A¤£­n¨Ï¥ÎWord¡I"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   210
      TabIndex        =   26
      Top             =   5070
      Width           =   4935
   End
   Begin MSForms.Label lblA0K04 
      Height          =   285
      Left            =   1260
      TabIndex        =   25
      Top             =   1230
      Width           =   7530
      VariousPropertyBits=   19
      Caption         =   "LblFM2"
      Size            =   "13282;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "¦¬¾Ú©ïÀY¡G"
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
      Index           =   4
      Left            =   240
      TabIndex        =   24
      Top             =   1230
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "´¼Åv¤H­û¡G"
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
      Index           =   3
      Left            =   2760
      TabIndex        =   23
      Top             =   870
      Width           =   1050
   End
   Begin MSForms.Label lblA0K20 
      Height          =   240
      Left            =   3840
      TabIndex        =   22
      Top             =   870
      Width           =   1680
      VariousPropertyBits=   19
      Caption         =   "LblFM2"
      Size            =   "2963;423"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "¦¬¾Ú½s¸¹¡G"
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
      Index           =   1
      Left            =   240
      TabIndex        =   21
      Top             =   150
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblAmt 
      BackStyle       =   0  '³z©ú
      Caption         =   "lblAmt"
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
      Left            =   3840
      TabIndex        =   20
      Top             =   510
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "INVOICE¹ô§O¡G"
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
      Index           =   53
      Left            =   240
      TabIndex        =   19
      Top             =   1560
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "¦¬¾Úª÷ÃB¡G"
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
      Index           =   13
      Left            =   2775
      TabIndex        =   18
      Top             =   510
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "«È¤á½s¸¹¡G"
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
      Index           =   0
      Left            =   240
      TabIndex        =   17
      Top             =   870
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "¶×¡@²v¡G"
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
      Index           =   2
      Left            =   6720
      TabIndex        =   16
      Top             =   1560
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "¦¬¾Ú¤é´Á¡G"
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
      Index           =   17
      Left            =   240
      TabIndex        =   15
      Top             =   510
      Width           =   1050
   End
   Begin VB.Label lblA0K03 
      BackStyle       =   0  '³z©ú
      Caption         =   "lblA0K03"
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
      Left            =   1260
      TabIndex        =   14
      Top             =   870
      Width           =   1680
   End
   Begin VB.Label lblA0K02 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "lblA0K02"
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
      Left            =   1260
      TabIndex        =   13
      Top             =   510
      Width           =   840
   End
End
Attribute VB_Name = "Frmacc14o0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/27 Form2.0¤w­×§ï
'Create By Sindy 2014/5/7
Option Explicit

Dim m_FileName As String
Dim m_FileName2 As String 'Add By Sindy 2018/11/16
Dim m_FileName3 As String 'Add By Sindy 2019/9/24
Dim m_FileName4 As String 'Add By Sindy 2021/5/10
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String, m_CP10 As String
Dim m_A1J04 As String, m_A1J05 As String, m_A1J06 As String
Dim pPrevRow As Integer 'Add By Sindy 2014/12/15
Dim m_Amount As String, m_Tot_Amount As String '¥~¹ôª÷ÃB
Dim m_strA0k03 As String   'Added by Lydia 2023/11/13 «È¤á½s¸¹(²Ä¤@±i²¼¾Ú)
Dim m_strA0k40 As String   'Added by Lydia 2023/11/13 INVOICE¬y¤ô¸¹(²Ä¤@±i²¼¾Ú)
Dim m_strA0k11 As String  'Added by Lydia 2025/07/30 ¦¬¾Ú¤½¥q§O(²Ä¤@±i²¼¾Ú)
Dim m_A0K11 As String   'Added by Lydia 2025/07/30 ¦¬¾Ú¤½¥q§O

Private Function TxtValidate() As Boolean
   TxtValidate = False
   
   'Add By Sindy 2014/12/15
   If GRD1.Rows < 2 Or (GRD1.Rows = 2 And GRD1.TextMatrix(1, 1) = "") Then
      MsgBox "½Ð¿é¤J¦¬¾Ú½s¸¹¡I"
      txtNo.SetFocus
      Exit Function
   End If
   '2014/12/15 END
   
   'Added by Lydia 2024/11/21 ­«¶]INVOICE
   If Frame1.Visible = True And Trim(Text2) <> "" Then
      m_strA0k40 = Trim(Text2)
   Else
   'end 2024/11/21
      'Added by Lydia 2023/11/13
      m_strA0k40 = AccAutoNo("Z", 3, Left(strSrvDate(2), 3), "0", True, True)
      If Right(m_strA0k40, 3) = "999" Then
         MsgBox "¬y¤ô¸¹¶W¹L³Ì¤jªø«×¡A½Ð³qª¾¹q¸£¤¤¤ß³B²z¡I", vbCritical, MsgText(5)
         Exit Function
      Else
         AccSaveAutoNo "Z", Right(m_strA0k40, 3), Left(strSrvDate(2), 3), "0"
      End If
      'end 2023/11/13
   End If
   TxtValidate = True
End Function

Private Sub JCallWordPrint()
Dim strFileName As String
Dim strName As String
Dim strText As String, strTemp As String
Dim i As Integer, ii As Integer, j As Integer
Dim rsA As ADODB.Recordset
Dim strCustAmtNo As String '«È¤á°]°È½s¸¹
Dim strPatentNo As String, strConNo As String
Dim strTradeMarkYes As String, strSystemName As String
Dim strCaseName As String, strAppNo As String
Dim strTM09 As String, strTM32 As String
Dim strTM28 As String, strProperty As String, strCache2 As String, strPrintText As String
Dim strLanguage As String
Dim strCustName(4) As String '¥Ó½Ð¤H¦WºÙ
Dim dblTotAmt As Double 'Add By Sindy 2014/12/16
Dim strNo As String

On Error GoTo ErrHand
   
   '§PÂ_word¬O§_¤w¶}±Ò
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If
  
   'ÀÉ¦W:«È¤á½s¸¹-¦¬¾Ú½s¸¹-¦¬¾Ú©ïÀY.doc
   strFileName = GRD1.TextMatrix(1, 4) & "-" & GRD1.TextMatrix(1, 1) & "-" & GRD1.TextMatrix(1, 6) & ".doc"
   If Dir(PUB_Getdesktop & "\" & strFileName) <> "" Then
      Kill PUB_Getdesktop & "\" & strFileName
   End If
   'Add By Sindy 2018/11/16
   'Modified by Lydia 2025/07/30 1=>2
   If Left(Combo2.Text, 1) = "2" Then '´¼¼z©Ò
      g_WordAp.Documents.Open App.path & "\" & strUserNum & "\" & m_FileName2
   '2018/11/16 END
   'Add By Sindy 2019/9/24
   ElseIf Left(Combo2.Text, 1) = "L" Then 'ªk«ß©Ò
      g_WordAp.Documents.Open App.path & "\" & strUserNum & "\" & m_FileName3
   '2019/9/24 END
   'Add By Sindy 2021/5/10
   ElseIf Left(Combo2.Text, 1) = "J" Then '´¼Åv¤½¥q
      g_WordAp.Documents.Open App.path & "\" & strUserNum & "\" & m_FileName4
   '2021/5/10 END
   Else
      g_WordAp.Documents.Open App.path & "\" & strUserNum & "\" & m_FileName
   End If
   g_WordAp.ActiveDocument.SaveAs PUB_Getdesktop & "\" & strFileName
   g_WordAp.ActiveDocument.Close
   g_WordAp.Documents.Open PUB_Getdesktop & "\" & strFileName
   strCustAmtNo = ""
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      'Add By Sindy 2014/12/15
      '¥ýÀË¬d¬O§_¬°¦hµ§ªº¦¬¾Ú½s¸¹,­Y¬O,¥ý½Æ»s/¶K¤WÀx¦s®æ
      If GRD1.Rows > 2 Then
         .Selection.Find.ClearFormatting
         .Selection.Find.Text = "Re:"
         .Selection.Find.Execute
         .Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
         .Selection.SelectRow
         .Selection.Copy
         For ii = 1 To GRD1.Rows - 2
            .Selection.Paste
         Next ii
      End If
      '2014/12/15 END
      
      'Add By Sindy 2019/9/24
      strNo = ""
      For j = 1 To GRD1.Rows - 1 '¥i¦hµ§¦¬¾Ú½s¸¹
         If GRD1.TextMatrix(j, 1) <> "" Then
            If strNo = "" Then
               strNo = GRD1.TextMatrix(j, 1)
            Else
               strNo = strNo & "¡B" & GRD1.TextMatrix(j, 1)
            End If
         End If
      Next j
      '2019/9/24 END
      
      Call doCaseQuery(GRD1.TextMatrix(1, 1)) 'Add By Sindy 2020/7/27
      '¼ÐÃD
      'Modified by Lydia 2023/11/13 4=>5 +INVOICE°O¿ý¬y¤ô½s¸¹
      For i = 1 To 5
         strName = ""
         strText = ""
         Select Case i
            Case 1
               strName = "¨t²Î¤é´Á"
               '­Y¬°¤­¤ë¥÷
               If Month(AFDate(strSrvDate(1))) = 5 Then
                  strText = Format(AFDate(strSrvDate(1)), "mmmm d, yyyy")
               '­Y«D¬°¤­¤ë¥÷
               Else
                  strText = Format(AFDate(strSrvDate(1)), "mmmm d, yyyy")
               End If
            Case 2
               strName = "«È¤á¸ê®Æ"
               '±M§Q¦³5­Ó¥Ó½Ð¤H
               strSql = "select cu05, cu88, cu89, cu90, pa85 as Lang, PA26 As CustNo, CU06, cu04,1 srt, PA26 AS CUST1 from patent, customer where substr(pa26, 1, 8) = cu01 and substr(pa26, 9, 1) = cu02 and pa01 = '" & m_CP01 & "' and pa02 = '" & m_CP02 & "' and pa03 = '" & m_CP03 & "' and pa04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, pa85 as Lang, PA27 As CustNo, CU06, cu04,2 srt, PA26 AS CUST1 from patent, customer where substr(pa27, 1, 8) = cu01 and substr(pa27, 9, 1) = cu02 and pa01 = '" & m_CP01 & "' and pa02 = '" & m_CP02 & "' and pa03 = '" & m_CP03 & "' and pa04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, pa85 as Lang, PA28 As CustNo, CU06, cu04,3 srt, PA26 AS CUST1 from patent, customer where substr(pa28, 1, 8) = cu01 and substr(pa28, 9, 1) = cu02 and pa01 = '" & m_CP01 & "' and pa02 = '" & m_CP02 & "' and pa03 = '" & m_CP03 & "' and pa04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, pa85 as Lang, PA29 As CustNo, CU06, cu04,4 srt, PA26 AS CUST1 from patent, customer where substr(pa29, 1, 8) = cu01 and substr(pa29, 9, 1) = cu02 and pa01 = '" & m_CP01 & "' and pa02 = '" & m_CP02 & "' and pa03 = '" & m_CP03 & "' and pa04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, pa85 as Lang, PA30 As CustNo, CU06, cu04,5 srt, PA26 AS CUST1 from patent, customer where substr(pa30, 1, 8) = cu01 and substr(pa30, 9, 1) = cu02 and pa01 = '" & m_CP01 & "' and pa02 = '" & m_CP02 & "' and pa03 = '" & m_CP03 & "' and pa04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, tm53 as Lang, TM23 As CustNo, CU06, cu04,1 srt, TM23 AS CUST1 from trademark, customer where substr(tm23, 1, 8) = cu01 and substr(tm23, 9, 1) = cu02 and tm01 = '" & m_CP01 & "' and tm02 = '" & m_CP02 & "' and tm03 = '" & m_CP03 & "' and tm04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, tm53 as Lang, TM78 As CustNo, CU06, cu04,2 srt, TM23 AS CUST1 from trademark, customer where substr(TM78, 1, 8) = cu01 and substr(TM78, 9, 1) = cu02 and tm01 = '" & m_CP01 & "' and tm02 = '" & m_CP02 & "' and tm03 = '" & m_CP03 & "' and tm04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, tm53 as Lang, TM79 As CustNo, CU06, cu04,3 srt, TM23 AS CUST1 from trademark, customer where substr(TM79, 1, 8) = cu01 and substr(TM79, 9, 1) = cu02 and tm01 = '" & m_CP01 & "' and tm02 = '" & m_CP02 & "' and tm03 = '" & m_CP03 & "' and tm04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, tm53 as Lang, TM80 As CustNo, CU06, cu04,4 srt, TM23 AS CUST1 from trademark, customer where substr(TM80, 1, 8) = cu01 and substr(TM80, 9, 1) = cu02 and tm01 = '" & m_CP01 & "' and tm02 = '" & m_CP02 & "' and tm03 = '" & m_CP03 & "' and tm04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, tm53 as Lang, TM81 As CustNo, CU06, cu04,5 srt, TM23 AS CUST1 from trademark, customer where substr(TM81, 1, 8) = cu01 and substr(TM81, 9, 1) = cu02 and tm01 = '" & m_CP01 & "' and tm02 = '" & m_CP02 & "' and tm03 = '" & m_CP03 & "' and tm04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, '' as Lang, LC11 As CustNo, CU06, cu04,1 srt, LC11 AS CUST1 from lawcase, customer where substr(lc11, 1, 8) = cu01 and substr(lc11, 9, 1) = cu02 and lc01 = '" & m_CP01 & "' and lc02 = '" & m_CP02 & "' and lc03 = '" & m_CP03 & "' and lc04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, '' as Lang, LC43 As CustNo, CU06, cu04,1 srt, LC11 AS CUST1 from lawcase, customer where substr(lc43, 1, 8) = cu01 and substr(lc43, 9, 1) = cu02 and lc01 = '" & m_CP01 & "' and lc02 = '" & m_CP02 & "' and lc03 = '" & m_CP03 & "' and lc04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, '' as Lang, LC44 As CustNo, CU06, cu04,1 srt, LC11 AS CUST1 from lawcase, customer where substr(lc44, 1, 8) = cu01 and substr(lc44, 9, 1) = cu02 and lc01 = '" & m_CP01 & "' and lc02 = '" & m_CP02 & "' and lc03 = '" & m_CP03 & "' and lc04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, '' as Lang, LC45 As CustNo, CU06, cu04,1 srt, LC11 AS CUST1 from lawcase, customer where substr(lc45, 1, 8) = cu01 and substr(lc45, 9, 1) = cu02 and lc01 = '" & m_CP01 & "' and lc02 = '" & m_CP02 & "' and lc03 = '" & m_CP03 & "' and lc04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, '' as Lang, LC46 As CustNo, CU06, cu04,1 srt, LC11 AS CUST1 from lawcase, customer where substr(lc46, 1, 8) = cu01 and substr(lc46, 9, 1) = cu02 and lc01 = '" & m_CP01 & "' and lc02 = '" & m_CP02 & "' and lc03 = '" & m_CP03 & "' and lc04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, '' as Lang, HC05 As CustNo, CU06, cu04,1 srt, HC05 AS CUST1 from hirecase, customer where substr(hc05, 1, 8) = cu01 and substr(hc05, 9, 1) = cu02 and hc01 = '" & m_CP01 & "' and hc02 = '" & m_CP02 & "' and hc03 = '" & m_CP03 & "' and hc04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, '' as Lang, HC24 As CustNo, CU06, cu04,1 srt, HC05 AS CUST1 from hirecase, customer where substr(hc24, 1, 8) = cu01 and substr(hc24, 9, 1) = cu02 and hc01 = '" & m_CP01 & "' and hc02 = '" & m_CP02 & "' and hc03 = '" & m_CP03 & "' and hc04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, '' as Lang, HC25 As CustNo, CU06, cu04,1 srt, HC05 AS CUST1 from hirecase, customer where substr(hc25, 1, 8) = cu01 and substr(hc25, 9, 1) = cu02 and hc01 = '" & m_CP01 & "' and hc02 = '" & m_CP02 & "' and hc03 = '" & m_CP03 & "' and hc04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, '' as Lang, HC26 As CustNo, CU06, cu04,1 srt, HC05 AS CUST1 from hirecase, customer where substr(hc26, 1, 8) = cu01 and substr(hc26, 9, 1) = cu02 and hc01 = '" & m_CP01 & "' and hc02 = '" & m_CP02 & "' and hc03 = '" & m_CP03 & "' and hc04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, '' as Lang, HC27 As CustNo, CU06, cu04,1 srt, HC05 AS CUST1 from hirecase, customer where substr(hc27, 1, 8) = cu01 and substr(hc27, 9, 1) = cu02 and hc01 = '" & m_CP01 & "' and hc02 = '" & m_CP02 & "' and hc03 = '" & m_CP03 & "' and hc04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, sp34 as Lang, SP08 As CustNo, CU06, cu04,1 srt, SP08 AS CUST1 from servicepractice, customer where substr(sp08, 1, 8) = cu01 and substr(sp08, 9, 1) = cu02 and sp01 = '" & m_CP01 & "' and sp02 = '" & m_CP02 & "' and sp03 = '" & m_CP03 & "' and sp04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, sp34 as Lang, SP58 As CustNo, CU06, cu04,2 srt, SP08 AS CUST1 from servicepractice, customer where substr(SP58, 1, 8) = cu01 and substr(SP58, 9, 1) = cu02 and sp01 = '" & m_CP01 & "' and sp02 = '" & m_CP02 & "' and sp03 = '" & m_CP03 & "' and sp04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, sp34 as Lang, SP59 As CustNo, CU06, cu04,3 srt, SP08 AS CUST1 from servicepractice, customer where substr(SP59, 1, 8) = cu01 and substr(SP59, 9, 1) = cu02 and sp01 = '" & m_CP01 & "' and sp02 = '" & m_CP02 & "' and sp03 = '" & m_CP03 & "' and sp04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, sp34 as Lang, SP65 As CustNo, CU06, cu04,4 srt, SP08 AS CUST1 from servicepractice, customer where substr(SP65, 1, 8) = cu01 and substr(SP65, 9, 1) = cu02 and sp01 = '" & m_CP01 & "' and sp02 = '" & m_CP02 & "' and sp03 = '" & m_CP03 & "' and sp04 = '" & m_CP04 & "' " & _
                  "union select cu05, cu88, cu89, cu90, sp34 as Lang, SP66 As CustNo, CU06, cu04,5 srt, SP08 AS CUST1 from servicepractice, customer where substr(SP66, 1, 8) = cu01 and substr(SP66, 9, 1) = cu02 and sp01 = '" & m_CP01 & "' and sp02 = '" & m_CP02 & "' and sp03 = '" & m_CP03 & "' and sp04 = '" & m_CP04 & "' order by srt"
               intI = 1
               Set rsA = ClsLawReadRstMsg(intI, strSql)
               '¦³°ò¥»ÀÉ
               If intI = 1 Then
                  '­Y°ò¥»ÀÉ¦³©w½Z»y¤å
                  If IsNull(rsA.Fields("Lang").Value) = False Then
                      strLanguage = rsA.Fields("Lang").Value
                  '­Y°ò¥»ÀÉµL©w½Z»y¤å
                  Else
                      strLanguage = GetLanguage(m_CP01, m_CP02, m_CP03, m_CP04, GRD1.TextMatrix(1, 1))
                  End If
                  strCustName(0) = "": strCustName(1) = "": strCustName(2) = "": strCustName(3) = "": strCustName(4) = ""
                  PrintCustData strCustName, strLanguage, rsA
               '¨S°ò¥»ÀÉ
               Else
                  strLanguage = GetLanguage(m_CP01, m_CP02, m_CP03, m_CP04)
                  strCustName(0) = "": strCustName(1) = "": strCustName(2) = "": strCustName(3) = "": strCustName(4) = ""
               End If
               
               '¥N²z¤H/¥Ó½Ð¤H¦WºÙ¡B¦a§}¤Î°]°È½s¸¹
               'Modify By Sindy 2014/11/27 §ï¥Î¦¬¾Ú©ïÀY§ì¸ê®Æ
   '            strSql = "select FA05,FA63,FA64,FA65,FA18,FA19,FA20,FA21,FA22,FA32,FA33,FA34,FA35,FA36,FA28,FA106,'' CU102 from fagent where FA01='" & Left(lblA0K03, 8) & "' and FA02='" & Right(lblA0K03, 1) & "'" & _
   '                     " Union" & _
   '                     " select CU05,CU88,CU89,CU90,CU24,CU25,CU26,CU27,CU28,CU65,CU66,CU67,CU68,CU69,CU33,CU146,CU102 from customer where CU01='" & Left(lblA0K03, 8) & "' and CU02='" & Right(lblA0K03, 1) & "'"
               'Modified by Lydia 2024/01/18 +ChgSQL
               strSql = "select FA05,FA63,FA64,FA65,FA18,FA19,FA20,FA21,FA22,FA32,FA33,FA34,FA35,FA36,FA28,FA106,'' CU102,1 sort from fagent" & _
                        " where FA05||' '||FA63||' '||FA64||' '||FA65='" & ChgSQL(Trim(GRD1.TextMatrix(1, 6))) & "' OR FA06='" & ChgSQL(Trim(GRD1.TextMatrix(1, 6))) & "'" & _
                        " Union" & _
                        " select CU05,CU88,CU89,CU90,CU24,CU25,CU26,CU27,CU28,CU65,CU66,CU67,CU68,CU69,CU33,CU146,CU102,2 sort from customer" & _
                        " where CU05||' '||CU88||' '||CU89||' '||CU90='" & ChgSQL(Trim(GRD1.TextMatrix(1, 6))) & "' OR CU06='" & ChgSQL(Trim(GRD1.TextMatrix(1, 6))) & "'" & _
                        " Union" & _
                        " select A4201 FA05,'' FA63,'' FA64,'' FA65,A4203 FA18,'' FA19,'' FA20,'' FA21,'' FA22,'' FA32,'' FA33,'' FA34,'' FA35,'' FA36,'' FA28,'' FA106,'' CU102,3 sort from acc420" & _
                        " where A4201='" & ChgSQL(Trim(GRD1.TextMatrix(1, 6))) & "'" & _
                        " order by sort asc"
               intI = 1
               Set rsA = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  If IsNull(rsA.Fields("fa05").Value) = False Then
                     strText = strText & rsA.Fields("fa05").Value & vbCrLf
                  End If
                  If IsNull(rsA.Fields("fa63").Value) = False Then
                     strText = strText & rsA.Fields("fa63").Value & vbCrLf
                  End If
                  If IsNull(rsA.Fields("fa64").Value) = False Then
                     strText = strText & rsA.Fields("fa64").Value & vbCrLf
                  End If
                  If IsNull(rsA.Fields("fa65").Value) = False Then
                     strText = strText & rsA.Fields("fa65").Value & vbCrLf
                  End If
                  '­^¤å¦a§}
                  If IsNull(rsA.Fields("fa32").Value) = False Then
                     strText = strText & rsA.Fields("fa32").Value & vbCrLf
                  ElseIf IsNull(rsA.Fields("fa18").Value) = False Then
                     strText = strText & rsA.Fields("fa18").Value & vbCrLf
                  End If
                  If IsNull(rsA.Fields("fa32").Value) Then
                     If IsNull(rsA.Fields("fa19").Value) = False Then
                        strText = strText & rsA.Fields("fa19") & vbCrLf
                     End If
                  Else
                     If IsNull(rsA.Fields("fa33").Value) = False Then
                        strText = strText & rsA.Fields("fa33") & vbCrLf
                     End If
                  End If
                  If IsNull(rsA.Fields("fa32").Value) Then
                     If IsNull(rsA.Fields("fa20").Value) = False Then
                        strText = strText & rsA.Fields("fa20").Value & vbCrLf
                     End If
                  ElseIf IsNull(rsA.Fields("fa34").Value) = False Then
                     strText = strText & rsA.Fields("fa34").Value & vbCrLf
                  End If
                  If IsNull(rsA.Fields("fa32").Value) Then
                     If IsNull(rsA.Fields("fa21").Value) = False Then
                        strText = strText & rsA.Fields("fa21").Value & vbCrLf
                     End If
                  ElseIf IsNull(rsA.Fields("fa35").Value) = False Then
                     strText = strText & rsA.Fields("fa35").Value & vbCrLf
                  End If
                  If IsNull(rsA.Fields("fa32").Value) Then
                     If IsNull(rsA.Fields("fa22").Value) = False Then
                        strText = strText & rsA.Fields("fa22").Value & vbCrLf
                     End If
                  Else
                     If IsNull(rsA.Fields("fa36").Value) = False Then
                        strText = strText & rsA.Fields("fa36").Value & vbCrLf
                     End If
                  End If
                  If IsNull(rsA.Fields("fa32").Value) Then
                     If IsNull(rsA.Fields("cu102").Value) = False Then
                        strText = strText & rsA.Fields("cu102") & vbCrLf
                     End If
                  End If
                  '°]°È½s¸¹
                  If CheckSys(m_CP01) = "2" Or CheckSys(m_CP01) = "6" Then '°Ó¼Ð
                     strCustAmtNo = "" & rsA(1)
                  Else
                     strCustAmtNo = "" & rsA(0)
                  End If
               Else
                  strText = Trim(GRD1.TextMatrix(1, 6))
               End If
            Case 3
               strName = "¥Ó½Ð¤H"
               strText = "Applicant: " & strCustName(0)
               '¨ä¥Lªº¥Ó½Ð¤H
               For ii = 1 To 4
                  If strCustName(ii) <> "" Then
                     strText = strText & vbCrLf & "         " & strCustName(ii)
                  End If
               Next ii
            'Add By Sindy 2019/9/24
            Case 4 '¦¬¾Ú½s¸¹
               strName = "¦¬¾Ú½s¸¹"
               'Modified by Lydia 2023/11/13
               'strText = strNo
               strText = "( " & strNo & " )"
            '2019/9/24 END
            'Added by Lydia 2023/11/13
            Case 5 'INVOICE¬y¤ô½s¸¹
               strName = "¬y¤ô½s¸¹"
               strText = "No." & m_strA0k40
            'end 2023/11/13
         End Select
         If Trim(strName) <> "" Then
            .Selection.Find.ClearFormatting
            .Selection.Find.Text = "|#" & strName & "#|"
            .Selection.Find.Replacement.Text = ""
            .Selection.Find.Forward = True
            .Selection.Find.Wrap = wdFindContinue
            .Selection.Find.Format = False
            .Selection.Find.MatchCase = False
            .Selection.Find.MatchWholeWord = False
            .Selection.Find.MatchWildcards = False
            .Selection.Find.MatchSoundsLike = False
            .Selection.Find.MatchAllWordForms = False
            .Selection.Find.MatchByte = True
            .Selection.Find.Execute
            .Selection.Delete
            .Selection.Font.ColorIndex = wdBlack
            .Selection.TypeText strText
         End If
      Next i
      
      For j = 1 To GRD1.Rows - 1 '¥i¦hµ§¦¬¾Ú½s¸¹
         '¤º®e
         For i = 1 To 6
            strName = ""
            strText = ""
            If i = 1 Then
               strName = "®×¸¹"
               If GRD1.Rows - 1 > 1 Then '¤@µ§¥H¤W¦¬¾Ú½s¸¹,¥»©Ò®×¸¹­n­«§ì
                  Call doCaseQuery(GRD1.TextMatrix(j, 1))
                  
                  '±M§Q¦³5­Ó¥Ó½Ð¤H
                  strSql = "select cu05, cu88, cu89, cu90, pa85 as Lang, PA26 As CustNo, CU06, cu04,1 srt, PA26 AS CUST1 from patent, customer where substr(pa26, 1, 8) = cu01 and substr(pa26, 9, 1) = cu02 and pa01 = '" & m_CP01 & "' and pa02 = '" & m_CP02 & "' and pa03 = '" & m_CP03 & "' and pa04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, pa85 as Lang, PA27 As CustNo, CU06, cu04,2 srt, PA26 AS CUST1 from patent, customer where substr(pa27, 1, 8) = cu01 and substr(pa27, 9, 1) = cu02 and pa01 = '" & m_CP01 & "' and pa02 = '" & m_CP02 & "' and pa03 = '" & m_CP03 & "' and pa04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, pa85 as Lang, PA28 As CustNo, CU06, cu04,3 srt, PA26 AS CUST1 from patent, customer where substr(pa28, 1, 8) = cu01 and substr(pa28, 9, 1) = cu02 and pa01 = '" & m_CP01 & "' and pa02 = '" & m_CP02 & "' and pa03 = '" & m_CP03 & "' and pa04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, pa85 as Lang, PA29 As CustNo, CU06, cu04,4 srt, PA26 AS CUST1 from patent, customer where substr(pa29, 1, 8) = cu01 and substr(pa29, 9, 1) = cu02 and pa01 = '" & m_CP01 & "' and pa02 = '" & m_CP02 & "' and pa03 = '" & m_CP03 & "' and pa04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, pa85 as Lang, PA30 As CustNo, CU06, cu04,5 srt, PA26 AS CUST1 from patent, customer where substr(pa30, 1, 8) = cu01 and substr(pa30, 9, 1) = cu02 and pa01 = '" & m_CP01 & "' and pa02 = '" & m_CP02 & "' and pa03 = '" & m_CP03 & "' and pa04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, tm53 as Lang, TM23 As CustNo, CU06, cu04,1 srt, TM23 AS CUST1 from trademark, customer where substr(tm23, 1, 8) = cu01 and substr(tm23, 9, 1) = cu02 and tm01 = '" & m_CP01 & "' and tm02 = '" & m_CP02 & "' and tm03 = '" & m_CP03 & "' and tm04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, tm53 as Lang, TM78 As CustNo, CU06, cu04,2 srt, TM23 AS CUST1 from trademark, customer where substr(TM78, 1, 8) = cu01 and substr(TM78, 9, 1) = cu02 and tm01 = '" & m_CP01 & "' and tm02 = '" & m_CP02 & "' and tm03 = '" & m_CP03 & "' and tm04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, tm53 as Lang, TM79 As CustNo, CU06, cu04,3 srt, TM23 AS CUST1 from trademark, customer where substr(TM79, 1, 8) = cu01 and substr(TM79, 9, 1) = cu02 and tm01 = '" & m_CP01 & "' and tm02 = '" & m_CP02 & "' and tm03 = '" & m_CP03 & "' and tm04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, tm53 as Lang, TM80 As CustNo, CU06, cu04,4 srt, TM23 AS CUST1 from trademark, customer where substr(TM80, 1, 8) = cu01 and substr(TM80, 9, 1) = cu02 and tm01 = '" & m_CP01 & "' and tm02 = '" & m_CP02 & "' and tm03 = '" & m_CP03 & "' and tm04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, tm53 as Lang, TM81 As CustNo, CU06, cu04,5 srt, TM23 AS CUST1 from trademark, customer where substr(TM81, 1, 8) = cu01 and substr(TM81, 9, 1) = cu02 and tm01 = '" & m_CP01 & "' and tm02 = '" & m_CP02 & "' and tm03 = '" & m_CP03 & "' and tm04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, '' as Lang, LC11 As CustNo, CU06, cu04,1 srt, LC11 AS CUST1 from lawcase, customer where substr(lc11, 1, 8) = cu01 and substr(lc11, 9, 1) = cu02 and lc01 = '" & m_CP01 & "' and lc02 = '" & m_CP02 & "' and lc03 = '" & m_CP03 & "' and lc04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, '' as Lang, LC43 As CustNo, CU06, cu04,1 srt, LC11 AS CUST1 from lawcase, customer where substr(lc43, 1, 8) = cu01 and substr(lc43, 9, 1) = cu02 and lc01 = '" & m_CP01 & "' and lc02 = '" & m_CP02 & "' and lc03 = '" & m_CP03 & "' and lc04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, '' as Lang, LC44 As CustNo, CU06, cu04,1 srt, LC11 AS CUST1 from lawcase, customer where substr(lc44, 1, 8) = cu01 and substr(lc44, 9, 1) = cu02 and lc01 = '" & m_CP01 & "' and lc02 = '" & m_CP02 & "' and lc03 = '" & m_CP03 & "' and lc04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, '' as Lang, LC45 As CustNo, CU06, cu04,1 srt, LC11 AS CUST1 from lawcase, customer where substr(lc45, 1, 8) = cu01 and substr(lc45, 9, 1) = cu02 and lc01 = '" & m_CP01 & "' and lc02 = '" & m_CP02 & "' and lc03 = '" & m_CP03 & "' and lc04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, '' as Lang, LC46 As CustNo, CU06, cu04,1 srt, LC11 AS CUST1 from lawcase, customer where substr(lc46, 1, 8) = cu01 and substr(lc46, 9, 1) = cu02 and lc01 = '" & m_CP01 & "' and lc02 = '" & m_CP02 & "' and lc03 = '" & m_CP03 & "' and lc04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, '' as Lang, HC05 As CustNo, CU06, cu04,1 srt, HC05 AS CUST1 from hirecase, customer where substr(hc05, 1, 8) = cu01 and substr(hc05, 9, 1) = cu02 and hc01 = '" & m_CP01 & "' and hc02 = '" & m_CP02 & "' and hc03 = '" & m_CP03 & "' and hc04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, '' as Lang, HC24 As CustNo, CU06, cu04,1 srt, HC05 AS CUST1 from hirecase, customer where substr(hc24, 1, 8) = cu01 and substr(hc24, 9, 1) = cu02 and hc01 = '" & m_CP01 & "' and hc02 = '" & m_CP02 & "' and hc03 = '" & m_CP03 & "' and hc04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, '' as Lang, HC25 As CustNo, CU06, cu04,1 srt, HC05 AS CUST1 from hirecase, customer where substr(hc25, 1, 8) = cu01 and substr(hc25, 9, 1) = cu02 and hc01 = '" & m_CP01 & "' and hc02 = '" & m_CP02 & "' and hc03 = '" & m_CP03 & "' and hc04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, '' as Lang, HC26 As CustNo, CU06, cu04,1 srt, HC05 AS CUST1 from hirecase, customer where substr(hc26, 1, 8) = cu01 and substr(hc26, 9, 1) = cu02 and hc01 = '" & m_CP01 & "' and hc02 = '" & m_CP02 & "' and hc03 = '" & m_CP03 & "' and hc04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, '' as Lang, HC27 As CustNo, CU06, cu04,1 srt, HC05 AS CUST1 from hirecase, customer where substr(hc27, 1, 8) = cu01 and substr(hc27, 9, 1) = cu02 and hc01 = '" & m_CP01 & "' and hc02 = '" & m_CP02 & "' and hc03 = '" & m_CP03 & "' and hc04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, sp34 as Lang, SP08 As CustNo, CU06, cu04,1 srt, SP08 AS CUST1 from servicepractice, customer where substr(sp08, 1, 8) = cu01 and substr(sp08, 9, 1) = cu02 and sp01 = '" & m_CP01 & "' and sp02 = '" & m_CP02 & "' and sp03 = '" & m_CP03 & "' and sp04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, sp34 as Lang, SP58 As CustNo, CU06, cu04,2 srt, SP08 AS CUST1 from servicepractice, customer where substr(SP58, 1, 8) = cu01 and substr(SP58, 9, 1) = cu02 and sp01 = '" & m_CP01 & "' and sp02 = '" & m_CP02 & "' and sp03 = '" & m_CP03 & "' and sp04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, sp34 as Lang, SP59 As CustNo, CU06, cu04,3 srt, SP08 AS CUST1 from servicepractice, customer where substr(SP59, 1, 8) = cu01 and substr(SP59, 9, 1) = cu02 and sp01 = '" & m_CP01 & "' and sp02 = '" & m_CP02 & "' and sp03 = '" & m_CP03 & "' and sp04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, sp34 as Lang, SP65 As CustNo, CU06, cu04,4 srt, SP08 AS CUST1 from servicepractice, customer where substr(SP65, 1, 8) = cu01 and substr(SP65, 9, 1) = cu02 and sp01 = '" & m_CP01 & "' and sp02 = '" & m_CP02 & "' and sp03 = '" & m_CP03 & "' and sp04 = '" & m_CP04 & "' " & _
                     "union select cu05, cu88, cu89, cu90, sp34 as Lang, SP66 As CustNo, CU06, cu04,5 srt, SP08 AS CUST1 from servicepractice, customer where substr(SP66, 1, 8) = cu01 and substr(SP66, 9, 1) = cu02 and sp01 = '" & m_CP01 & "' and sp02 = '" & m_CP02 & "' and sp03 = '" & m_CP03 & "' and sp04 = '" & m_CP04 & "' order by srt"
                  intI = 1
                  Set rsA = ClsLawReadRstMsg(intI, strSql)
                  '¦³°ò¥»ÀÉ
                  If intI = 1 Then
                     '­Y°ò¥»ÀÉ¦³©w½Z»y¤å
                     If IsNull(rsA.Fields("Lang").Value) = False Then
                        strLanguage = rsA.Fields("Lang").Value
                     '­Y°ò¥»ÀÉµL©w½Z»y¤å
                     Else
                        strLanguage = GetLanguage(m_CP01, m_CP02, m_CP03, m_CP04, GRD1.TextMatrix(j, 1))
                     End If
                  '¨S°ò¥»ÀÉ
                  Else
                     strLanguage = GetLanguage(m_CP01, m_CP02, m_CP03, m_CP04)
                  End If
               End If
               
               'TS¬d¦W®×µL²Õ¸s®É¦LÃþ§O
               '+Ápµ¸¤H1(­^) TM39 (Syngenta ªº Reuester)
               'ptm05 as MName
               strSql = "select pa77 as Yno, pa48 as Cno, 'Patent' as MName, nvl(nvl(pa06,pa05),pa07) as Cname, pa11 as Ano, pa26 as Custno, pa22, null as tm15, null as Yes, '' As TM12, '' As TM16, '' As TM09, '' AS TM32, PA52 As TM39 from patent, systemkind, patenttrademarkmap where pa01 = sk01 and pa08 = ptm02 (+) and sk02 = ptm01 and pa01 = '" & m_CP01 & "' and pa02 = '" & m_CP02 & "' and pa03 = '" & m_CP03 & "' and pa04 = '" & m_CP04 & "' union " & _
                        "select tm45 as Yno, tm35 as Cno, 'Trademark' as MName, nvl(nvl(tm06,tm05),tm07) as Cname, tm12 as Ano, tm23 as Custno, null as pa22, tm15, '1' as Yes, TM12, TM16, TM09, TM32, TM39 from trademark, systemkind, patenttrademarkmap where tm01 = sk01 and tm08 = ptm02 (+) and sk02 = ptm01 and tm01 = '" & m_CP01 & "' and tm02 = '" & m_CP02 & "' and tm03 = '" & m_CP03 & "' and tm04 = '" & m_CP04 & "' union " & _
                        "select lc23 as Yno, lc17 as Cno, '' as MName, nvl(nvl(lc06,lc05),lc07) as Cname, '' as Ano, lc11 as Custno, null as pa22, null as tm15, null as Yes, '' As TM12, '' As TM16, '' As TM09, '' AS TM32, LC19 As TM39 from lawcase where lc01 = '" & m_CP01 & "' and lc02 = '" & m_CP02 & "' and lc03 = '" & m_CP03 & "' and lc04 = '" & m_CP04 & "' union " & _
                        "select sp27 as Yno, sp29 as Cno, '' as MName, nvl(nvl(sp06,sp05),sp07) as Cname, sp11 as Ano, sp08 as Custno, null as pa22, null as tm15, null as Yes, '' As TM12, '' As TM16, sp73 As TM09, SP74 AS TM32, SP30 As TM39 from servicepractice where sp01 = '" & m_CP01 & "' and sp02 = '" & m_CP02 & "' and sp03 = '" & m_CP03 & "' and sp04 = '" & m_CP04 & "'"
               intI = 1
               Set rsA = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  '¼ÐÃD Your Ref »P¸ê®Æ¤£­n¤À¶}¦L
                  '­Y¬°FCP®×
                  strTemp = ""
                  'Add By Sindy 2016/7/18 +ÀË¬d°Ó¼Ð©µ®i
                  'If m_CP01 = "FCP" Then
                  If m_CP01 = "FCP" Or InStr("T,FCT,CFT,TF", m_CP01) > 0 Then
                     If m_CP01 = "FCP" Then
                        strSql = "Select PA106 From Patent, CaseProgress Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP60='" & GRD1.TextMatrix(j, 1) & "' And CP10='605' And CP01='FCP' And PA76 Is Not Null"
                     Else
                        strSql = "Select TM65 From Trademark, CaseProgress Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP60='" & GRD1.TextMatrix(j, 1) & "' And CP10='102' And CP01 in('T','FCT','CFT','TF') And TM33 Is Not Null"
                     End If
                  '2016/7/18 END
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        If GetFCCaseNo(GRD1.TextMatrix(j, 1), strExc(1), True) = True Then
                           strTemp = strExc(1)
                        Else
                           strTemp = RsTemp.Fields(0).Value
                        End If
                     '¨ä¥L
                     ElseIf GetFCCaseNo(GRD1.TextMatrix(j, 1), strExc(1)) = True Then
                        strTemp = strExc(1)
                     'Modified by Lydia 2025/07/04 Yno §ï§ì Cno
                     ElseIf Not IsNull(rsA.Fields("Cno").Value) Then
                        strTemp = rsA.Fields("Cno").Value
                     End If
                  '¨ä¥L
                  ElseIf GetFCCaseNo(GRD1.TextMatrix(j, 1), strExc(1)) = True Then
                     strTemp = strExc(1)
                  'Modified by Lydia 2025/07/04 Yno §ï§ì Cno
                  ElseIf Not IsNull(rsA.Fields("Cno").Value) Then
                     strTemp = rsA.Fields("Cno").Value
                  End If
                  If strTemp <> "" Then  'Memo by Lydia 2025/07/04 °ê¥~®×¤~¥Î©¼©Ò®×¸¹¡A°ê¤º®×¥Î«È¤á®×¥ó®×¸¹
                     strText = strText & "Your Ref: " & strTemp & vbCrLf
                  End If
                  'Our Ref
                  strText = strText & "Our Ref : " & m_CP01 & "-" & m_CP02
                  '­Y¥»©Ò®×¸¹«á¤T½X¬°000«h¤£¦L¦¹¤T½X
                  If m_CP03 & m_CP04 <> "000" Then
                      strText = strText & "-" & m_CP03 & "-" & m_CP04
                  End If
                  If IsNull(rsA.Fields("pa22").Value) = False Then
                     strPatentNo = rsA.Fields("pa22").Value
                  Else
                     strPatentNo = ""
                  End If
                  '­Y¬°®Ö­ã¥B¦³¼f©w¸¹®É, ¦L¼f©w¸¹, §_«h¦L¥Ó½Ð®×¸¹
                  If "" & rsA.Fields("TM16").Value = "1" And "" & rsA.Fields("TM15").Value <> "" Then
                     strConNo = "" & rsA.Fields("TM15").Value
                  Else
                     strConNo = ""
                  End If
                  If IsNull(rsA.Fields("Yes").Value) = False Then
                     strTradeMarkYes = rsA.Fields("Yes").Value
                  Else
                     strTradeMarkYes = ""
                  End If
                  If IsNull(rsA.Fields("MName").Value) Then
                     strSystemName = ""
                  Else
                     strSystemName = rsA.Fields("MName").Value
                  End If
                  If IsNull(rsA.Fields("Cname").Value) Then
                     strCaseName = ""
                  Else
                     strCaseName = rsA.Fields("Cname").Value
                  End If
                  If IsNull(rsA.Fields("Ano").Value) Then
                     strAppNo = ""
                  Else
                     strAppNo = rsA.Fields("Ano").Value
                  End If
                  If IsNull(rsA.Fields("Custno").Value) Then
                     strCustNo = ""
                  Else
                     strCustNo = rsA.Fields("Custno").Value
                  End If
                  strTM09 = "" & rsA.Fields("TM09").Value
                  If IsNull(rsA.Fields("TM32").Value) Then
                     strTM32 = ""
                  Else
                     strTM32 = rsA.Fields("TM32").Value
                  End If
               End If
               rsA.Close
            ElseIf i = 2 Then
               strName = "RE"
               
               strTM28 = ""
               '­Y¨t²ÎÃþ§O¬°FCT
               If m_CP01 = "FCT" Or m_CP01 = "T" Then
                  strSql = "Select tm15,tm28 From Trademark Where " & ChgTradeMark(m_CP01 & m_CP02 & m_CP03 & m_CP04) & " And TM28 Is Not Null And TM28<>'1'"
                  intI = 1
                  Set rsA = ClsLawReadRstMsg(intI, strSql)
                  '­Y¦³®×¥ó©Ê½è¬°²§Ä³, µû©w, ¼o¤î
                  If intI = 1 Then
                     strTM28 = "" & rsA("TM28").Value
                     strConNo = "" & rsA("TM15").Value
                  End If
               End If
               
               strSql = "select cp10, cp01 from caseprogress where cp60 = '" & GRD1.TextMatrix(j, 1) & "' and cp10>='101' and cp10<='105'"
               intI = 1
               Set rsA = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  Select Case rsA.Fields("cp01").Value
                     Case "CFT", "FCT", "T"
                        If rsA.Fields("cp10").Value = "101" Then
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
               
               strCache2 = ""
               strPrintText = ""
               '­Y¬°°Ó¼Ð®×
               If strTradeMarkYes = "1" Then
                  Select Case strTM28
                     Case "2" '²§Ä³
                           strCache2 = "Opposition Action against " & GetNationName(2) & "Mark"
                           strPrintText = "Registration No. " & strConNo
                     Case "3" 'µû©w
                           strCache2 = "Invalidation Action against " & GetNationName(2) & "Mark"
                           strPrintText = "Registration No. " & strConNo
                     Case "4" 'Á|µo(¼o¤î)
                           strCache2 = "Revocation Action against " & GetNationName(2) & "Mark"
                           strPrintText = "Registration No. " & strConNo
                     Case Else '¨ä¥L
                           If strConNo <> "" Then
                              '¥h±¼¥kÃäªºªÅ¥Õ
                              strCache2 = RTrim(GetNationName(2) & strProperty & strSystemName) & " Registration " & IIf(strConNo <> "", "No. " & strConNo, "")
                           Else
                              '¥h±¼¥kÃäªºªÅ¥Õ
                              strCache2 = RTrim(GetNationName(2) & strProperty & strSystemName) & " Application " & IIf(strAppNo <> "", "No. " & strAppNo, "")
                           End If
                  End Select
               ElseIf m_CP01 = "TS" Then
                  strCache2 = "Trademark Search in " & GetNationName(2)
               ElseIf m_CP01 = "TM" Then
                  strCache2 = "Monitoring system (C.C.C. Code) in " & GetNationName(2)
               '­Y«D°Ó¼Ð®×
               Else
                  Select Case m_CP01
                     '­Y¨t²ÎÃþ§O¬°"S"
                     Case "S"
                        '¥h±¼¥kÃäªºªÅ¥Õ
                        strCache2 = RTrim(GetNationName(2) & strProperty & strSystemName) & " Trademark Search for " & GetCaseName(m_CP01 & m_CP02 & m_CP03 & m_CP04)
                     Case "TD"
                        '¥h±¼¥kÃäªºªÅ¥Õ
                        strCache2 = RTrim(GetNationName(2) & strProperty & strSystemName) & " domain name: " & strCaseName
                     Case "FG"
                        strCache2 = GetCaseName(m_CP01 & m_CP02 & m_CP03 & m_CP04, strLanguage)
                     Case Else '¨ä¥L¨t²ÎÃþ§O
                        If strPatentNo <> "" Then
                           '¥h±¼¥kÃäªºªÅ¥Õ
                           strCache2 = RTrim(GetNationName(2) & strProperty & strSystemName) & " Application " & IIf(strAppNo <> "", "No. " & strAppNo, "") '& " (Patent No. " & strPatentNo & ")"
                        Else
                           '¥h±¼¥kÃäªºªÅ¥Õ
                           strCache2 = RTrim(GetNationName(2) & strProperty & strSystemName) & " Application " & IIf(strAppNo <> "", "No. " & strAppNo, "")
                        End If
                  End Select
               End If
               If strCache2 <> "" Then
                  strText = strText & strCache2 '& vbCrLf
               End If
   
'               '­Y¬°°Ó¼Ð®×¥ó
               If strTradeMarkYes = "1" Then
                  If strTM28 = "2" Or strTM28 = "3" Or strTM28 = "4" Then
                     strText = strText & strPrintText
                  End If
               ElseIf m_CP01 = "TS" Then
                  strText = strText & " for " & GetCaseName(m_CP01 & m_CP02 & m_CP03 & m_CP04) & " in class " & strTM09 & vbCrLf
               End If
            ElseIf i = 3 Then
               strName = "Title"
               If m_CP01 <> "TS" And m_CP01 <> "S" Then
                  strText = ChgSQL(strCaseName) & vbCrLf
               End If
            ElseIf i = 4 Then
               strName = "´Ú¶µ"
               strText = Trim(m_A1J04) & " " & Trim(m_A1J05) & " " & m_A1J06
            'Add By Sindy 2021/5/18
            ElseIf i = 5 Then
               strName = "¹ô§OD"
               '1.¥þ³¡ 2.¶È¥x¹ô 3.¶È¥~¹ô
               If Left(Combo3.Text, 1) = "1" Then
                  strText = "NTD" & vbCrLf & Left(Combo1.Text, 3)
               ElseIf Left(Combo3.Text, 1) = "2" Then
                  strText = "NTD"
               Else
                  strText = Left(Combo1.Text, 3)
               End If
            '2021/5/18 END
            ElseIf i = 6 Then
               strName = "ª÷ÃB"
   '            'Modify By Sindy 2014/10/9
   '            strName = "¹ô§O"
   '            strText = Left(Combo1, 3)
   '         ElseIf i = 9 Then
   ''            strName = "Á`­p"
   ''            strText = Format(lblAmt, FDollar)
   '            'Modify By Sindy 2014/10/9
   '            strName = "¥~¹ôª÷ÃB"
   '            m_Tot_Amount = 0
   '            If Val(lblAmt) > 0 Then
   '               m_Tot_Amount = Format(Round(Val(lblAmt) / Val(txtRate), 2), FDollar)
   '            End If
   '            strText = m_Tot_Amount
   '         ElseIf i = 10 Then
   '            strName = "¹ô§O"
   '            strText = Left(Combo1, 3)
   '         ElseIf i = 11 Then
   '            strName = "¥~¹ôª÷ÃB"
   '            m_Tot_Amount = 0
   '            If Val(lblAmt) > 0 Then
   '               m_Tot_Amount = Format(Round(Val(lblAmt) / Val(txtRate), 2), FDollar)
   '            End If
   '            strText = m_Tot_Amount
   '         ElseIf i = 12 Then
   '            'Modify By Sindy 2014/10/9 ¨ú®ø
   '            strName = "±b¸¹¸ê®Æ"
   ''            strText = strText & ReportSum(71001) & vbCrLf
   ''            strText = strText & ReportSum(72) & vbCrLf
   ''            strText = strText & ReportSum(73001) & vbCrLf
   ''            strText = strText & ReportSum(85) & vbCrLf
   ''            strText = strText & ReportSum(74) & vbCrLf
   ''            strText = strText & ReportSum(121) & vbCrLf
   '            strText = strText & "Currency Rate: " & Left(Combo1, 3) & "1.00=NTD" & txtRate
   '         ElseIf i = 13 Then
   '            'Modify By Sindy 2014/10/9 ¨ú®ø
   '            strName = "PS"
   '            strText = Trim(Replace(ReportSum(86001), "PS:", ""))
               'strText = Format(GRD1.TextMatrix(j, 3), FDollar)
               'Modify By Sindy 2021/5/18
               '1.¥þ³¡ 2.¶È¥x¹ô 3.¶È¥~¹ô
               If Left(Combo3.Text, 1) = "1" Then
                  strText = Format(GRD1.TextMatrix(j, 3), FDollar) & vbCrLf & Format(GRD1.TextMatrix(j, 7), FDollar)
               ElseIf Left(Combo3.Text, 1) = "2" Then
                  strText = Format(GRD1.TextMatrix(j, 3), FDollar)
               Else
                  strText = Format(GRD1.TextMatrix(j, 7), FDollar)
               End If
               '2021/5/18 END
               dblTotAmt = dblTotAmt + Val(GRD1.TextMatrix(j, 3))
            End If
            .Selection.WholeStory
            .Selection.Copy
            If Trim(strName) <> "" Then
               .Selection.Find.ClearFormatting
               .Selection.Find.Text = "|#" & strName & "#|"
               .Selection.Find.Replacement.Text = ""
               .Selection.Find.Forward = True
               .Selection.Find.Wrap = wdFindContinue
               .Selection.Find.Format = False
               .Selection.Find.MatchCase = False
               .Selection.Find.MatchWholeWord = False
               .Selection.Find.MatchWildcards = False
               .Selection.Find.MatchSoundsLike = False
               .Selection.Find.MatchAllWordForms = False
               .Selection.Find.MatchByte = True
               .Selection.Find.Execute
               .Selection.Delete
               .Selection.Font.ColorIndex = wdBlack
               .Selection.TypeText strText
            End If
         Next i
      Next j
      '¦X­p
      'Modified by Lydia 2023/11/13 2=>3 +¶×¤J»È¦æ¸ê®Æ
      For i = 1 To 3
         strName = ""
         strText = ""
         Select Case i
            Case 1
               strName = "¹ô§OT"
               'strText = Left(Combo1, 3)
               'Modify By Sindy 2021/5/18
               '1.¥þ³¡ 2.¶È¥x¹ô 3.¶È¥~¹ô
               If Left(Combo3.Text, 1) = "1" Then
                  strText = "NTD" & vbCrLf & Left(Combo1.Text, 3)
               ElseIf Left(Combo3.Text, 1) = "2" Then
                  strText = "NTD"
               Else
                  strText = Left(Combo1.Text, 3)
               End If
               '2021/5/18 END
            Case 2
               strName = "Á`­p"
'               m_Tot_Amount = 0
'               If dblTotAmt > 0 Then
'                  m_Tot_Amount = Format(Round(dblTotAmt / Val(txtRate), 2), FDollar) '¤p¼Æ¦ì2,¥|±Ë¤­¤J
'               End If
'               strText = m_Tot_Amount
               '1.¥þ³¡ 2.¶È¥x¹ô 3.¶È¥~¹ô
               If Left(Combo3.Text, 1) = "1" Then
                  strText = Format(dblTotAmt, FDollar) & vbCrLf & Format(m_Tot_Amount, FDollar)
               ElseIf Left(Combo3.Text, 1) = "2" Then
                  strText = Format(dblTotAmt, FDollar)
               Else
                  strText = Format(m_Tot_Amount, FDollar)
               End If
               '2021/5/18 END
            'Added by Lydia 2023/11/13
            Case 3
               strName = "¶×¤J»È¦æ¸ê®Æ"
               strText = ""
               If txtCU196 <> "" Then
                  strTemp = txtCU196
               Else
                  strTemp = "3" '¹w³]3.¥x¤@´¼¼z (µØ»È¥~¹ô)
               End If
               strSql = "select * from rptaccount where ra01='CU196' and ra02='" & Format(Val(strTemp), "00") & "' "
               Set rsA = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  strText = strText & "Name of Bank¡G" & rsA.Fields("RA04") & vbCrLf
                  If "" & rsA.Fields("RA05") <> "" Then
                     strText = strText & "Bank Address¡G" & rsA.Fields("RA05") & vbCrLf
                  End If
                  If "" & rsA.Fields("RA08") <> "" Then
                     strText = strText & "SWIFT code¡G" & rsA.Fields("RA08") & vbCrLf
                  End If
                  If "" & rsA.Fields("RA06") <> "" Then
                     strText = strText & "Account Name¡G" & rsA.Fields("RA06") & vbCrLf
                  End If
                  If "" & rsA.Fields("RA07") <> "" Then
                     strText = strText & "Account No¡G" & rsA.Fields("RA07") & vbCrLf
                  End If
               End If
            'end 2023/11/13
         End Select
         If Trim(strName) <> "" Then
            .Selection.Find.ClearFormatting
            .Selection.Find.Text = "|#" & strName & "#|"
            .Selection.Find.Replacement.Text = ""
            .Selection.Find.Forward = True
            .Selection.Find.Wrap = wdFindContinue
            .Selection.Find.Format = False
            .Selection.Find.MatchCase = False
            .Selection.Find.MatchWholeWord = False
            .Selection.Find.MatchWildcards = False
            .Selection.Find.MatchSoundsLike = False
            .Selection.Find.MatchAllWordForms = False
            .Selection.Find.MatchByte = True
            .Selection.Find.Execute
            .Selection.Delete
            .Selection.Font.ColorIndex = wdBlack
            .Selection.TypeText strText
         End If
      Next i
      'Added by Lydia 2023/11/13
      .Selection.WholeStory
      .Selection.Font.Name = "Times New Romans"
      'end 2023/11/13
      .ActiveWindow.ActivePane.View.Zoom.Percentage = 100 'Add By Sindy 2021/5/26 (·PÄ±¨S¦³¤°»ò®ÄªG)
      .ActiveDocument.Save
   End With
   g_WordAp.ActiveDocument.Close
   g_WordAp.Quit
   Set g_WordAp = Nothing
   'Added by Lydia 2023/11/13 §ó·sINVOICE¬y¤ô¸¹
   If m_strA0k40 <> "" And strNo <> "" Then
      strExc(1) = "UPDATE ACC0K0 SET A0K40=" & CNULL(m_strA0k40) & " WHERE A0K01 IN (" & GetAddStr(Replace(strNo, "¡B", ",")) & ") AND A0K40 IS NULL "
      cnnConnection.Execute strExc(1)
      strExc(1) = "UPDATE CUSTOMER SET CU196=" & CNULL(txtCU196) & " WHERE CU01='" & Mid(m_strA0k03, 1, 8) & "' AND CU02='" & Mid(m_strA0k03, 9, 1) & "' AND CU196 IS NULL "
      Pub_SeekTbLog strExc(1)
      cnnConnection.Execute strExc(1)
   End If
   Call cmdClear_Click
   'end 2023/11/13
   
   MsgBox "ÀÉ®×¤w²£¥Í¦Ü®à­±¡I"
   
   Exit Sub
   
ErrHand:
   If Err.Number = 462 Then '»·ºÝ¦øªA¾¹¤£¦s¦b©ÎµLªk¨Ï¥Î
      GoTo RestarWord
   ElseIf Err.Number <> 0 Then
      MsgBox (Err.Description)
      If Not g_WordAp Is Nothing Then
         g_WordAp.Quit
         Set g_WordAp = Nothing
      End If
   End If
End Sub

'Add By Sindy 2014/12/15
Private Sub cmdAddRow_Click()
Dim bolChk As Boolean
Dim ii As Integer
   
   If Len(txtNo) = 0 Then
      MsgBox "½Ð¿é¤J¦¬¾Ú½s¸¹¡I"
      txtNo.SetFocus
      Exit Sub
   End If
   
   If Trim(Combo1.Text) = "" Then
      MsgBox "½ÐÂI¿ïINVOICE¹ô§O¡I"
      Combo1.SetFocus
      Exit Sub
   End If
   
   If Val(txtRate) = 0 Then
      MsgBox "½Ð¿é¤J¶×²v¡I"
      txtRate.SetFocus
      Exit Sub
   End If
   'Added by Lydia 2023/11/13
   Call SetCombo4(IIf(txtCU196.Text = "", "0", txtCU196.Text))
   If Trim(txtCU196) = "" Or Trim(Combo4.Text) = "" Then
      MsgBox "½Ð¿é¤J½Ð´Ú¶×¤J»È¦æ¸ê®Æ¡I"
      txtCU196.SetFocus
      txtCU196_GotFocus
      Exit Sub
   End If
   If m_strA0k03 <> "" And m_strA0k03 <> lblA0K03 Then
      If MsgBox("²Ä¤@µ§²¼¾Úªº«È¤á½s¸¹¡G" & m_strA0k03 & vbCrLf & "¤£¦P«È¤á½s¸¹¡A¬O§_Ä~Äò¥[¤J¡H", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
         txtNo.SetFocus
         txtNo_GotFocus
         Exit Sub
      End If
   End If
   'end 2023/11/13
   
   'ÀË¬d¦¬¾Ú½s¸¹
   bolChk = True
   If txtNo = "" Or Trim(lblA0K02) = "" Then Exit Sub
   For ii = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(ii, 1) = txtNo Then
         bolChk = False
         Exit For
      End If
   Next ii
   If Not bolChk Then
      MsgBox "¦¬¾Ú½s¸¹¤£¥i­«ÂÐ !", vbCritical
      txtNo.SetFocus
      Exit Sub
   End If
   
   If Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 1) <> "" Then
      GRD1.AddItem ""
   End If
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 1) = txtNo
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 2) = lblA0K02
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 3) = lblAmt
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 4) = lblA0K03
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 5) = lblA0K20
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 6) = lblA0K04
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 8) = m_A0K11 'Added by Lydia 2025/07/30
   ClearAll '²MªÅÄæ¦ì
   
   Call AmtCount '­pºâª÷ÃB/¶×²v
   
   cmdWord.Enabled = True
End Sub

Private Sub AmtCount()
Dim j As Integer
Dim dblAmt_Tot As Double, dblAmt As Double
   
   '¦hµ§¦¬¾Ú
   m_Amount = 0
   For j = 1 To GRD1.Rows - 1
      m_Amount = m_Amount + Val(GRD1.TextMatrix(j, 3)) '¥x¹ô¦X­pª÷ÃB
   Next j
   
   '***************************************
   '¤p¼Æ¦ì2,¥|±Ë¤­¤J
   '***************************************
   '¥~¹ô¦X­pª÷ÃB
   m_Tot_Amount = Format(Round((m_Amount + Val(txtFee)) / Val(txtRate), 2), FDollar)
   '´«ºâ«á¶×²v
   If Val(txtFee) > 0 Then
      'Modify By Sindy 2021/6/7 ¦]¬°¶×²v´«ºâ¥u¨ì¤p¼ÆÂI«á2¦ì, ¥H­PÀ³¸Óª÷ÃB¤@­Pªº¨C¤@µ§¦³»~®t,½Õ¾ã¬°¤p¼ÆÂI«á6¦ì
      'txtRate2 = Format(Round(m_Amount / m_Tot_Amount, 2), FDollar)
      txtRate2 = Format(Round(m_Amount / m_Tot_Amount, 6), "###,###,###,###.000000")
   Else
      txtRate2 = ""
   End If
   
   '©ú²Ó
   dblAmt_Tot = 0
   For j = 1 To GRD1.Rows - 1
      '¥~¹ôª÷ÃB
      dblAmt = Format(Round(Val(GRD1.TextMatrix(j, 3)) / IIf(Val(txtRate2) > 0, txtRate2, txtRate), 2), FDollar)
      Me.GRD1.TextMatrix(j, 7) = dblAmt
      'Modify By Sindy 2021/6/7
      dblAmt_Tot = dblAmt_Tot + dblAmt
'      '³Ì«á¤@µ§¬O¥Î¾lÃB
'      If GRD1.Rows - 1 > 1 Then
'         If j = GRD1.Rows - 1 Then
'            Me.GRD1.TextMatrix(j, 7) = m_Tot_Amount - dblAmt_Tot
'         Else
'            dblAmt_Tot = dblAmt_Tot + dblAmt
'         End If
'      Else
'         Me.GRD1.TextMatrix(j, 7) = m_Tot_Amount '¥u¦³¤@µ§
'      End If
   Next j
   m_Tot_Amount = dblAmt_Tot
   '2021/6/7 END
End Sub

'Add By Sindy 2014/12/15
Private Sub cmdDelRow_Click()
Dim ii As Integer
Dim bolFind As Boolean, bolClear As Boolean 'Added by Lydia 2023/11/13

   'Modified by Lydia 2023/11/13 §ï§PÂ_V
   'If pPrevRow = 1 And GRD1.Rows = 2 Then
   '   For ii = 0 To 6
   '      GRD1.TextMatrix(pPrevRow, ii) = ""
   '   Next ii
   'Else
   '   If pPrevRow > 0 Then
   '      Call GRD1.RemoveItem(pPrevRow)
   '   Else
   '      Exit Sub
   '   End If
   'End If
   'pPrevRow = pPrevRow - 1
   'ClearAll '²MªÅÄæ¦ì
   'If pPrevRow > 0 Then cmdWord.Enabled = True
   For ii = 1 To GRD1.Rows - 1
      If "" & GRD1.TextMatrix(ii, 0) = "V" And "" & GRD1.TextMatrix(ii, 1) <> "" Then
         If GRD1.Rows = 2 And ii = 1 Then  'µLªk²¾°£³Ì«á¤@µ§
            For intI = 0 To GRD1.Cols - 1
               GRD1.TextMatrix(ii, intI) = ""
            Next intI
            bolClear = True
         Else
            Call GRD1.RemoveItem(ii)
         End If
         bolFind = True
         Exit For
      End If
   Next ii
   If bolFind = False Then
      MsgBox "½Ð¤Ä¿ï¸ê®Æ¡I", vbCritical
   Else
      pPrevRow = 0
      ClearAll
      If GRD1.Rows > 1 And bolClear = False Then
         cmdWord.Enabled = True
         m_strA0k03 = "" & GRD1.TextMatrix(1, 4) '²Ä¤@µ§
         m_strA0k11 = "" & GRD1.TextMatrix(1, 8) 'Added by Lydia 2025/07/30
      Else
         cmdWord.Enabled = False
         m_strA0k03 = ""
         m_strA0k11 = "" 'Added by Lydia 2025/07/30
      End If
   End If
   'end 2023/11/13
End Sub

'Add By Sindy 2014/12/15
Private Sub cmdClear_Click()
   'Modified by Lydia 2023/11/13
   'ClearAll
   ClearAll True
   GRD1.Clear
   SetGrd
   txtNo = ""
   txtNo.SetFocus
End Sub

'Add By Sindy 2014/12/15
Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   '                        0    1           2           3           4           5           6          7           8
   'Modified by Lydia 2025/07/30
   'arrGridHeadText = Array("V", "¦¬¾Ú½s¸¹", "¦¬¾Ú¤é´Á", "¦¬¾Úª÷ÃB", "«È¤á½s¸¹", "´¼Åv¤H­û", "¦¬¾Ú©ïÀY", "¥~¹ôª÷ÃB")
   'arrGridHeadWidth = Array(200, 1000, 800, 1000, 1000, 1000, 3000, 0)
   arrGridHeadText = Array("V", "¦¬¾Ú½s¸¹", "¦¬¾Ú¤é´Á", "¦¬¾Úª÷ÃB", "«È¤á½s¸¹", "´¼Åv¤H­û", "¦¬¾Ú©ïÀY", "¥~¹ôª÷ÃB", "A0K11")
   arrGridHeadWidth = Array(200, 1000, 800, 1000, 1000, 1000, 3000, 0, 0)
   'end 2025/07/30
   GRD1.Visible = False
   GRD1.Clear 'Added by Lydia 2023/11/13
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

'Add By Sindy 2014/12/15
Private Sub Command3_Click()
   Call doQuery
End Sub

Private Sub doQuery()
Dim Rs As ADODB.Recordset
   
   Call ClearAll
   
   Screen.MousePointer = vbHourglass
   
   'Modified by Lydia 2023/11/13
   'strExc(0) = "select a0k01,sqldatet(a0k02) a0k02,st02,sum(a0j09-nvl(a1u07,0)+a0j10-nvl(a1u09,0)) amt,a0k03,a0k04" & _
               " from acc1u0,(" & _
               "select a0k01,a0k02,st02,a0j01,a0j07,a0j25,a0k10,a0j09,a0j10,a0k03,a0k04" & _
               " From acc0j0,acc0k0,staff" & _
               " Where (a0k09 Is Null Or a0k09 = 0)" & _
               " and (a0k37<>'N' or a0k37 is null)" & _
               " and a0k20=st01(+)" & _
               " and a0k01='" & txtNo & "' and a0k01=a0j13) d" & _
               " where d.a0k01=a1u02(+)" & _
               " and d.a0j01=a1u03(+)" & _
               " and d.a0k10=a1u01(+)" & _
               " group by a0k01,a0k02,st02,a0k03,a0k04"
   'Memo by Lydia 2024/11/21 ­Y¦³²§°Ê,½Ð¤@¨Ö½Õ¾ãcmdRun
   'Modified by Lydia 2025/07/30 +A0K11
   strExc(0) = "select a0k01,substr(sqldatet(a0k02),1,9) a0k02,st02,sum(a0j09-nvl(a1u07,0)+a0j10-nvl(a1u09,0)) amt,a0k03,a0k04,cu196,a0k40,A0K11" & _
               " from acc1u0,(" & _
               "select a0k01,a0k02,st02,a0j01,a0j07,a0j25,a0k10,a0j09,a0j10,a0k03,a0k04,cu196,a0k40,A0K11" & _
               " From acc0j0,acc0k0,staff,customer" & _
               " Where (a0k09 Is Null Or a0k09 = 0)" & _
               " and (a0k37<>'N' or a0k37 is null)" & _
               " and a0k20=st01(+)" & _
               " and a0k01='" & txtNo & "' and a0k01=a0j13 and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+)) d" & _
               " where d.a0k01=a1u02(+)" & _
               " and d.a0j01=a1u03(+)" & _
               " and d.a0k10=a1u01(+)" & _
               " group by a0k01,a0k02,st02,a0k03,a0k04,cu196,a0k40,A0K11"
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      MsgBox "µL¥ô¦ó½Ð´Ú³æ¸ê®Æ¡I", , MsgText(5)
   Else
      'Added by Lydia 2023/11/13 §PÂ_¬O§_¤w¶}INVOICE
      If "" & Rs.Fields("a0k40") <> "" Then
         MsgBox "¤w¶}¥ßINVOICE½s¸¹¡G" & Rs.Fields("a0k40"), vbInformation, "¤w¶}¥ßINVOICE"
         GoTo JumpToSkip
      End If
      'end 2023/11/13
      cmdWord.Enabled = True
      lblA0K02.Caption = Rs.Fields("a0k02")
      lblAmt.Caption = Rs.Fields("amt")
      lblA0K03.Caption = Rs.Fields("a0k03")
      lblA0K20.Caption = Rs.Fields("st02")
      lblA0K04.Caption = Rs.Fields("a0k04")
      'Added by Lydia 2023/11/13 ½Ð´Ú¶×¤J»È¦æ¸ê®Æ
      txtCU196.Locked = True
      Combo4.Locked = True
      If txtCU196 = "" Then
         txtCU196 = "" & Rs.Fields("CU196")
         If txtCU196 = "" Then
            txtCU196.Locked = False
            Combo4.Locked = False
         End If
         Call SetCombo4(IIf(txtCU196.Text = "", "0", txtCU196.Text))
         txtCU196.Tag = txtCU196.Text
      End If
      If m_strA0k03 = "" Then
         m_strA0k03 = lblA0K03.Caption
         'Added by Lydia 2025/07/30
         m_strA0k11 = "" & Rs.Fields("a0k11")
         Combo2.Text = m_strA0k11 & " " & CompNameQuery(m_strA0k11, "4")
         '¦¬¾Ú¤½¥q§O­Y¬°J¤½¥q¡A«h½Ð´Ú¶×¤J»È¦æ¸ê®Æ¹w³]¬°"5.¥x¤@´¼Åv(µØ»È¥x¹ô)"
         If m_strA0k11 = "J" And txtCU196 <> "5" Then
            txtCU196 = "5"
            Call SetCombo4(IIf(txtCU196.Text = "", "0", txtCU196.Text))
            txtCU196.Tag = txtCU196.Text
         End If
         'end 2025/07/30
      End If
      'end 2023/11/13
      Call doCaseQuery(txtNo)
      
      txtRate.SetFocus
   End If
   
JumpToSkip: 'Added by Lydia 2023/11/13

   Set Rs = Nothing
   Screen.MousePointer = vbDefault
End Sub

Private Sub doCaseQuery(strKey As String)
Dim rs1 As ADODB.Recordset
   
   m_CP01 = ""
   m_CP02 = ""
   m_CP03 = ""
   m_CP04 = ""
   m_CP10 = ""
   m_A1J04 = ""
   m_A1J05 = ""
   m_A1J06 = ""
   m_A0K11 = "" 'Added by Lydia 2025/07/30
   
   'modify by sonia 2018/1/17 ±Æ§Ç±ø¥ó+a0j01 asc
   'Modified by Lydia 2025/07/30 +A0K11
   strExc(0) = "select cp01,cp02,cp03,cp04,cp10,a1j04,a1j05,a1j06,A0K11 from acc0k0,acc0j0,caseprogress,acc1j0" & _
               " where a0k01='" & strKey & "' and a0k01=a0j13" & _
               " and a0j01=cp09(+)" & _
               " and cp01=a1j01(+) and cp10=a1j02(+)" & _
               " order by a0j25 asc,a0j01 asc"
   intI = 1
   Set rs1 = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rs1.MoveFirst
      m_CP01 = rs1.Fields("cp01")
      m_CP02 = rs1.Fields("cp02")
      m_CP03 = rs1.Fields("cp03")
      m_CP04 = rs1.Fields("cp04")
      m_CP10 = rs1.Fields("cp10")
      m_A1J04 = "" & rs1.Fields("a1j04")
      m_A1J05 = "" & rs1.Fields("a1j05")
      m_A1J06 = "" & rs1.Fields("a1j06")
      m_A0K11 = "" & rs1.Fields("A0K11") 'Added by Lydia 2025/07/30
   End If
   
   'Add By Sindy 2015/11/5 §ìACC1J0¤§A1J04,A1J05,A1J06®É,­Y§ì¤£¨ì©ÎªÌA1J04¬°ªÅªÌ,
   '                       ¦A§ï§ì®×¥ó©Ê½èÀÉCASEPROPERTYMAP¤§­^¤å®×¥ó©Ê½è¦WºÙCPM10
   If m_A1J04 = "" And m_CP01 <> "" And m_CP10 <> "" Then
      strExc(0) = "select cpm10 from CASEPROPERTYMAP" & _
                  " where cpm01='" & m_CP01 & "' and cpm02='" & m_CP10 & "'"
      intI = 1
      Set rs1 = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_A1J04 = "" & rs1.Fields("cpm10")
      End If
   End If
   '2015/11/5 END
   
   Set rs1 = Nothing
End Sub

'Add By Sindy 2014/12/15
Private Sub Grd1_Click()
Dim nCol As Integer, nRow As Integer
Dim iCol As Integer
   
   With GRD1
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow > 0 And .TextMatrix(nRow, 1) <> "" Then
      nCol = .col
      If pPrevRow > 0 Then
         If pPrevRow <> nRow Then
            .row = pPrevRow
            .TextMatrix(pPrevRow, 0) = ""
            If .FixedCols > 0 Then
               .col = .FixedCols - 1
               .CellBackColor = .BackColorFixed
               .CellForeColor = .ForeColor
            End If
            For iCol = .FixedCols To .Cols - 1
               .col = iCol
               .CellBackColor = .BackColor
            Next
         End If
      End If
   
      If nRow > 0 Then
         .row = nRow
         .TextMatrix(nRow, 0) = "V"
         If .FixedCols > 0 Then
            .col = .FixedCols - 1
            .CellBackColor = .BackColorSel
            .CellForeColor = .ForeColorSel
         End If
         For iCol = .FixedCols To .Cols - 1
           .col = iCol
           .CellBackColor = &HFFC0C0
         Next
      End If
      .col = nCol
      pPrevRow = nRow
   End If
   .Visible = True
   End With
End Sub

Private Sub cmdWord_Click()
   If TxtValidate Then
      Screen.MousePointer = vbHourglass
      
      Call AmtCount '­pºâª÷ÃB/¶×²v
       
      Call JCallWordPrint
      
      Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  ¥\¯àÁä©w¸q
'
'*************************************************
Public Sub KeyDefine(KeyCode As Integer)
On Error GoTo Checking
   
   Select Case KeyCode
      Case vbKeyF12
         If FormCheck Then
            Screen.MousePointer = vbHourglass
            doQuery
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         
      Case Else
         KeyEnter KeyCode
   End Select
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
   Exit Sub
   
Checking:
   Screen.MousePointer = vbDefault
   MsgBox Err.Description, , MsgBox(5)
End Sub

'*************************************************
'  µe­±¿é¤JÀË¬d
'
'*************************************************
Public Function FormCheck() As Boolean
   FormCheck = False
   
   If Trim(txtNo) = "" Then
      MsgBox "¦¬¾Ú½s¸¹¤£¥iªÅ¥Õ¡I", , MsgText(5)
      txtNo.SetFocus
      Exit Function
   End If
   
   FormCheck = True
End Function

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
Dim intRow As Integer, intUSD As Integer
   
   'Modified by Lydia 2023/11/13 ªí³æªì©l¤Æ
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 9045 '6405
'   Me.Height = 5700 '3465
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath4)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   PUB_InitForm Me, 9060, 5820, strBackPicPath4, lngWidth, lngHeight
   'end 2023/11/13
      
   cmdWord.Enabled = False
   lblA0K02.Caption = ""
   lblAmt.Caption = ""
   lblA0K03.Caption = ""
   lblA0K20.Caption = ""
   lblA0K04.Caption = ""
   
   intRow = 0
   strExc(0) = "SELECT A1Y01||'-'||A1Y02 FROM ACC1Y0"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         Combo1.AddItem RsTemp.Fields(0)
         If Left(RsTemp.Fields(0), 3) = "USD" Then
            intUSD = intRow
         End If
         intRow = intRow + 1
         RsTemp.MoveNext
      Loop
   End If
   Combo1.ListIndex = intUSD
   
   'Add By Sindy 2019/9/24 ®M¥Î«HÀY
   Combo2.Clear
   'Combo2.AddItem "" 'Modify By Sindy °û²ñ:®M¥Î«HÀY"ªÅ¥Õ"¨ú®ø
   'Modified by Lydia 2025/07/30
   'Combo2.AddItem "1 ´¼¼z©Ò"
   'Combo2.AddItem "L ªk«ß©Ò"
   'Combo2.AddItem "J ´¼Åv¤½¥q"
   Combo2.AddItem "2 " & CompNameQuery("2", "4")
   Combo2.AddItem "L " & CompNameQuery("L", "4")
   Combo2.AddItem "J " & CompNameQuery("J", "4")
   'end 2025/07/30
   Combo2.ListIndex = 0
   '2019/9/24 END
      
   Combo3.ListIndex = 0 'Add By Sindy 2021/5/18
   
   m_FileName = "$$°ê¤º¦¬¾ÚINVOICE.doc"
   If Dir(App.path & "\" & strUserNum & "\" & m_FileName) <> "" Then
      Kill App.path & "\" & strUserNum & "\" & m_FileName
   End If
   Call PUB_GetSampleFile(m_FileName, "M31-000004-0-00", , App.path & "\" & strUserNum & "\")
   'Add By Sindy 2018/11/16
   m_FileName2 = "$$°ê¤º¦¬¾ÚINVOICE_2.doc"
   If Dir(App.path & "\" & strUserNum & "\" & m_FileName2) <> "" Then
      Kill App.path & "\" & strUserNum & "\" & m_FileName2
   End If
   Call PUB_GetSampleFile(m_FileName2, "M31-000004-1-00", , App.path & "\" & strUserNum & "\")
   '2018/11/16 END
   'Add By Sindy 2019/9/24
   m_FileName3 = "$$°ê¤º¦¬¾ÚINVOICE_3.doc"
   If Dir(App.path & "\" & strUserNum & "\" & m_FileName3) <> "" Then
      Kill App.path & "\" & strUserNum & "\" & m_FileName3
   End If
   Call PUB_GetSampleFile(m_FileName3, "M31-000004-2-00", , App.path & "\" & strUserNum & "\")
   '2019/9/24 END
   'Add By Sindy 2021/5/10
   m_FileName4 = "$$°ê¤º¦¬¾ÚINVOICE_4.doc"
   If Dir(App.path & "\" & strUserNum & "\" & m_FileName4) <> "" Then
      Kill App.path & "\" & strUserNum & "\" & m_FileName4
   End If
   Call PUB_GetSampleFile(m_FileName4, "M31-000004-3-00", , App.path & "\" & strUserNum & "\")
   '2021/5/10 END
   
   'Modify By Sindy 2014/12/15
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
   GRD1.Clear
   SetGrd
   '2014/12/15 END
   
   'Added by Lydia 2023/11/13 ½Ð´Ú¶×¤J»È¦æ¸ê®Æ¹w³]²M³æ
   Call SetCombo4
   txtCU196.Text = ""
   'end 2023/11/13
   
   'Added by Lydia 2024/11/21
   If Pub_StrUserSt03 <> "M51" Then
      Frame1.Visible = False
   End If
   'end 2024/11/21
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHand

   If Not g_WordAp Is Nothing Then
      g_WordAp.Visible = True
      g_WordAp.Quit
CloseWord:
      Set g_WordAp = Nothing
   End If
   
   StatusClear
   strFormName = MsgText(601)
   MenuEnabled
   Set Frmacc14o0 = Nothing
   
   Exit Sub
   
ErrHand:
   If Err.Number = 462 Then '»·ºÝ¦øªA¾¹¤£¦s¦b©ÎµLªk¨Ï¥Î
      GoTo CloseWord
   ElseIf Err.Number <> 0 Then
      MsgBox (Err.Description)
   End If
End Sub

'Modified by Lydia 2023/11/13 + bolReset
Private Sub ClearAll(Optional ByVal bolReset As Boolean = False)
   lblA0K02.Caption = ""
   lblAmt.Caption = ""
   lblA0K03.Caption = ""
   lblA0K20.Caption = ""
   lblA0K04.Caption = ""
   cmdWord.Enabled = False
   'Added by Lydia 2023/11/13
   If bolReset = True Then
      txtCU196.Text = "": txtCU196.Tag = ""
      m_strA0k03 = ""
      Combo4.ListIndex = 0
      Debug.Print "11"
      txtCU196.Locked = False
      Combo4.Locked = False
      SetGrd
      m_strA0k11 = "" 'Added by Lydia 2025/07/30
   End If
End Sub

'Add By Sindy 2021/5/13
Private Sub txtRate2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   txtRate2.ToolTipText = "(¦¬¾Úª÷ÃB+¤âÄò¶O)/ ¶×²v =´«ºâ«á¥~¹ô" & vbCrLf & _
                       " ¦¬¾Úª÷ÃB/ ´«ºâ«á¥~¹ô =´«ºâ«á¶×²v"
End Sub

Private Sub txtFee_GotFocus()
   InverseTextBox txtFee
End Sub

Private Sub txtFee_KeyPress(KeyAscii As Integer)
   If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9"))) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
      Exit Sub
   End If
End Sub

Private Sub txtNo_GotFocus()
   InverseTextBox txtNo
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Call ClearAll
End Sub

Private Sub txtRate_GotFocus()
   InverseTextBox txtRate
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
   If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9"))) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
      Exit Sub
   End If
End Sub

'********** »P Frmacc2480 ¬Û¦P **********
'¨ú±o©w½Z»y¤å
'¥[½Ð´Ú³æ¸¹ p_A1K01 ¥H§PÂ_¬O§_¬°¦~¶O½Ð´Ú¨Ã§ïCall¦@¥Î¨ç¼ÆPUB_GetLanguage
Public Function GetLanguage(strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, Optional ByVal p_A1K01 As String) As String
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
'¨ú±o®×¥ó¦WºÙ
'strLang:0=¤¤+­^+¤é 1=¤¤ 2=­^ 3=¤é
Private Function GetCaseName(strCP0104 As String, Optional strLang As String = "0") As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select PA05,PA06,PA07 From Patent Where " & ChgPatent(strCP0104)
StrSQLa = StrSQLa & " union Select TM05,TM06,TM07 From Trademark Where " & ChgTradeMark(strCP0104)
StrSQLa = StrSQLa & " union Select LC05,LC06,LC07 From Lawcase Where " & ChgLawcase(strCP0104)
StrSQLa = StrSQLa & " union Select SP05,SP06,SP07 From ServicePractice Where " & ChgService(strCP0104)
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
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
'PrintHead µ{§Ç¤Ó¤j©â¥X
Private Sub PrintCustData(strCustName() As String, strLanguage, tmpRsA As ADODB.Recordset)
Dim ii As Integer
   ii = 0
   tmpRsA.MoveFirst
   While Not tmpRsA.EOF
      Select Case strLanguage
         Case "1"
            If IsNull(tmpRsA.Fields("cu04").Value) = False Then
               strCustName(ii) = strCustName(ii) & "" & tmpRsA.Fields("cu04").Value
            End If
         Case "2"
            If IsNull(tmpRsA.Fields("cu05").Value) = False Then
               strCustName(ii) = strCustName(ii) & "" & tmpRsA.Fields("cu05").Value
               If IsNull(tmpRsA.Fields("cu88").Value) = False Then
                  strCustName(ii) = strCustName(ii) & " " & tmpRsA.Fields("cu88").Value
               End If
               If IsNull(tmpRsA.Fields("cu89").Value) = False Then
                  strCustName(ii) = strCustName(ii) & " " & tmpRsA.Fields("cu89").Value
               End If
               If IsNull(tmpRsA.Fields("cu90").Value) = False Then
                  strCustName(ii) = strCustName(ii) & " " & tmpRsA.Fields("cu90").Value
               End If
            ElseIf IsNull(tmpRsA.Fields("cu06").Value) = False Then
               strCustName(ii) = strCustName(ii) & "" & tmpRsA.Fields("cu06").Value
            End If
         Case "3"
            If IsNull(tmpRsA.Fields("cu06").Value) = False Then
               strCustName(ii) = strCustName(ii) & "" & tmpRsA.Fields("cu06").Value
            ElseIf IsNull(tmpRsA.Fields("cu05").Value) = False Then
               strCustName(ii) = strCustName(ii) & "" & tmpRsA.Fields("cu05").Value
               If IsNull(tmpRsA.Fields("cu88").Value) = False Then
                  strCustName(ii) = strCustName(ii) & " " & tmpRsA.Fields("cu88").Value
               End If
               If IsNull(tmpRsA.Fields("cu89").Value) = False Then
                  strCustName(ii) = strCustName(ii) & " " & tmpRsA.Fields("cu89").Value
               End If
               If IsNull(tmpRsA.Fields("cu90").Value) = False Then
                  strCustName(ii) = strCustName(ii) & " " & tmpRsA.Fields("cu90").Value
               End If
            End If
      End Select
      ii = ii + 1
      tmpRsA.MoveNext
   Wend
End Sub
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
      ElseIf Asc(strChr) < 0 Then
         CountLength = CountLength + 2
      Else
         CountLength = CountLength + 1
      End If
   Next ii
End If
End Function
Private Function GetNationName(strKind As String) As String
'strKind : 2 ­^¤å
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   GetNationName = ""
   StrSQLa = "Select NA03, NA04,na01 From Nation, Patent Where NA01=PA09 And " & ChgPatent(m_CP01 & m_CP02 & m_CP03 & m_CP04)
   StrSQLa = StrSQLa & " Union Select NA03, NA04,na01 From Nation, Trademark Where NA01=TM10 And " & ChgTradeMark(m_CP01 & m_CP02 & m_CP03 & m_CP04)
   StrSQLa = StrSQLa & " Union Select NA03, NA04,na01 From Nation, Lawcase Where NA01=LC15 And " & ChgLawcase(m_CP01 & m_CP02 & m_CP03 & m_CP04)
   StrSQLa = StrSQLa & " Union Select NA03, NA04,na01 From Nation, Hirecase Where '000'=NA01 And " & ChgHirecase(m_CP01 & m_CP02 & m_CP03 & m_CP04)
   StrSQLa = StrSQLa & " Union Select NA03, NA04,na01 From Nation, Servicepractice Where NA01=SP09 And " & ChgService(m_CP01 & m_CP02 & m_CP03 & m_CP04)
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       GetNationName = IIf(strKind = "2", "" & rsA.Fields(1).Value, "" & rsA.Fields(0).Value)
       If strKind = "3" Then
         If GetNationName = "¥xÆW" Then
            'Modified by Lydia 2023/11/13 §ï¥Î¼Ò²Õ¨ú±o
            'GetNationName = "¥x“ø"
            GetNationName = "¥x" & PUB_GetUniText("¦@¥Î", "ÆW")
         ElseIf rsA.Fields("na01") = "020" Then
            'Modified by Lydia 2023/11/13 §ï¥Î¼Ò²Õ¨ú±o
            'GetNationName = "¤¤üÂ"
            GetNationName = "¤¤" & PUB_GetUniText("¦@¥Î", "°ê")
         End If
       End If
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   If GetNationName <> "" And strKind = "2" Then GetNationName = GetNationName & " "
End Function
'********** End **********

'Åª¨ú©¼©Ò®×¸¹²§°Ê¸ê®Æ - °ê¤º¦¬¾Ú
Private Function GetFCCaseNo(pCP60 As String, ByRef opRefNo As String, Optional pIs605or102 As Boolean = False) As Boolean
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   
   opRefNo = ""
   stSQL = "select FL06 from acc0k0,FCCaseNoLog where a0k01='" & pCP60 & "' and FL01='" & m_CP01 & "' and FL02='" & m_CP02 & "' and FL03='" & m_CP03 & "' and FL04='" & m_CP04 & "' and FL07='" & IIf(pIs605or102 = True, "1", "2") & "' and FL09*1000000+FL10>(a0k24+19110000)*1000000+a0k25 order by FL09 asc,FL10 asc"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      opRefNo = "" & rsQuery(0)
      GetFCCaseNo = True
   End If
   Set rsQuery = Nothing
End Function

'Added by Lydia 2023/11/13 ½Ð´Ú¶×¤J»È¦æ¸ê®Æ
Private Sub SetCombo4(Optional ByVal pTxt As String)
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   
   If pTxt = "" Then  '¹w³]²M³æ
      Combo4.Clear
      stSQL = "select * from rptaccount where ra01='CU196' order by ra02"
      intR = 1
      Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         rsQuery.MoveFirst
         Combo4.AddItem "   "
         Do While Not rsQuery.EOF
            Combo4.AddItem "" & rsQuery.Fields("ra03")
            rsQuery.MoveNext
         Loop
      End If
      Set rsQuery = Nothing
      Combo4.ListIndex = 0
   Else
      If Val(pTxt) > 0 And Val(pTxt) < Combo4.ListCount Then
         Combo4.ListIndex = Val(pTxt)
      Else
         Combo4.ListIndex = 0
      End If
   End If
End Sub
'Added by Lydia 2023/11/13
Private Sub txtCU196_GotFocus()
   TextInverse txtCU196
End Sub
'Added by Lydia 2023/11/13
Private Sub txtCU196_Validate(Cancel As Boolean)
   If txtCU196.Tag <> txtCU196.Text Then
      Call SetCombo4(IIf(txtCU196.Text = "", "0", txtCU196.Text))
   End If
End Sub
'Added by Lydia 2023/11/13
Private Sub Combo4_Click()
   If txtCU196.Locked = False And txtCU196.Enabled = True Then
      'Added by Lydia 2024/12/10
      If Combo4.ListIndex = 0 Then
         txtCU196.Text = ""
      Else
      'end 2024/12/10
         txtCU196.Text = Combo4.ListIndex
      End If
   End If
End Sub

'Added by Lydia 2024/11/21
Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

'Added by Lydia 2024/11/21 ­«¶]INVOICE
Private Sub cmdRun_Click()
   
   If Len(Trim(Text2)) <> 6 Then Exit Sub
   
   If Trim(Combo1.Text) = "" Then
      MsgBox "½ÐÂI¿ïINVOICE¹ô§O¡I"
      Combo1.SetFocus
      Exit Sub
   End If
   
   If Val(txtRate) = 0 Then
      MsgBox "½Ð¿é¤J¶×²v¡I"
      txtRate.SetFocus
      Exit Sub
   End If
   'Modifeid by Lydia 2025/07/30 +A0K11
   strExc(0) = "select a0k01,substr(sqldatet(a0k02),1,9) a0k02,st02,sum(a0j09-nvl(a1u07,0)+a0j10-nvl(a1u09,0)) amt,a0k03,a0k04,cu196,a0k40,A0K11" & _
               " from acc1u0,(" & _
               "select a0k01,a0k02,st02,a0j01,a0j07,a0j25,a0k10,a0j09,a0j10,a0k03,a0k04,cu196,a0k40,A0K11" & _
               " From acc0j0,acc0k0,staff,customer" & _
               " Where (a0k09 Is Null Or a0k09 = 0)" & _
               " and (a0k37<>'N' or a0k37 is null)" & _
               " and a0k20=st01(+)" & _
               " and a0k40='" & Trim(Text2) & "' and a0k01=a0j13 and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+)) d" & _
               " where d.a0k01=a1u02(+)" & _
               " and d.a0j01=a1u03(+)" & _
               " and d.a0k10=a1u01(+)" & _
               " group by a0k01,a0k02,st02,a0k03,a0k04,cu196,a0k40,A0K11"
   strExc(0) = strExc(0) & " order by 1,2"
   intI = 1
   SetGrd
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      
      RsTemp.MoveFirst
      
      txtCU196 = "" & RsTemp.Fields("CU196")
      'Added by Lydia 2025/07/30
      m_strA0k11 = "" & RsTemp.Fields("a0k11")
      Combo2.Text = m_strA0k11 & " " & CompNameQuery(m_strA0k11, "4")
      'end 2025/07/30
      Call SetCombo4(IIf(txtCU196.Text = "", "0", txtCU196.Text))
      If Trim(txtCU196) = "" Or Trim(Combo4.Text) = "" Then
         MsgBox "½Ð¿é¤J½Ð´Ú¶×¤J»È¦æ¸ê®Æ¡I"
         txtCU196.SetFocus
         txtCU196_GotFocus
         Exit Sub
      End If
      Do While Not RsTemp.EOF
         If Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 1) <> "" Then
            GRD1.AddItem ""
         End If
         Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 1) = "" & RsTemp.Fields("a0k01")
         Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 2) = "" & RsTemp.Fields("a0k02")
         Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 3) = "" & RsTemp.Fields("amt")
         Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 4) = "" & RsTemp.Fields("a0k03")
         Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 5) = "" & RsTemp.Fields("st02")
         Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 6) = "" & RsTemp.Fields("a0k04")
         Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 8) = "" & RsTemp.Fields("a0k11") 'Added by Lydia 2025/07/30
         RsTemp.MoveNext
      Loop
      
      If TxtValidate Then
         Screen.MousePointer = vbHourglass
         
         Call AmtCount '­pºâª÷ÃB/¶×²v
          
         Call JCallWordPrint
         
         Screen.MousePointer = vbDefault
      End If
   End If
   
End Sub
