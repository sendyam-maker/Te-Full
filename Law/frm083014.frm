VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm083014 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "¦a§}±ø¦C¦L"
   ClientHeight    =   5232
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   12564
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5232
   ScaleWidth      =   12564
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Caption         =   "Frame2"
      Height          =   5295
      Left            =   6240
      TabIndex        =   37
      Top             =   0
      Width           =   6255
      Begin VB.TextBox txtStartPage 
         Alignment       =   2  '¸m¤¤¹ï»ô
         Height          =   264
         Left            =   1536
         MaxLength       =   2
         TabIndex        =   51
         Text            =   "1"
         Top             =   600
         Width           =   300
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   1080
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   43
         Top             =   4800
         Width           =   3840
      End
      Begin VB.CommandButton Command3 
         Caption         =   "µ²§ô"
         Height          =   400
         Left            =   4800
         TabIndex        =   41
         Top             =   120
         Width           =   800
      End
      Begin VB.CommandButton Command2 
         Caption         =   "¦C¦L"
         Height          =   400
         Left            =   3720
         TabIndex        =   40
         Top             =   120
         Width           =   800
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "§R°£"
         Height          =   400
         Left            =   2640
         Style           =   1  '¹Ï¤ù¥~Æ[
         TabIndex        =   39
         Top             =   120
         Width           =   800
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid1 
         Height          =   3732
         Left            =   120
         TabIndex        =   38
         Top             =   912
         Width           =   6012
         _ExtentX        =   10605
         _ExtentY        =   6583
         _Version        =   393216
         Cols            =   6
         FormatString    =   "V|¤é¡@´Á|¶¶§Ç|¦¬ ¥ó ¤H|¡@¡@¦¬¡@¥ó¡@¦a¡@§}¡@¡@|X,Y½s¸¹/¥»©Ò®×¸¹"
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
         _Band(0).Cols   =   6
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "±qA4¦W±øªº²Ä        ±i¶K¯È¶}©l¦C¦L"
         Height          =   180
         Index           =   3
         Left            =   384
         TabIndex        =   50
         Top             =   648
         Width           =   2748
      End
      Begin VB.Label lblCnt 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "·s²Ó©úÅé"
            Size            =   11.4
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   720
         TabIndex        =   45
         Top             =   195
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "¦@¡@¡@ ±i¦a§}±ø"
         BeginProperty Font 
            Name            =   "·s²Ó©úÅé"
            Size            =   11.4
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   193
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "¦Lªí¾÷"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   42
         Top             =   4800
         Width           =   765
      End
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "A4§å¦¸¦C¦L"
      Height          =   255
      Index           =   4
      Left            =   2640
      TabIndex        =   36
      Top             =   2130
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   16
      Left            =   1470
      MaxLength       =   9
      TabIndex        =   35
      Text            =   "R"
      Top             =   1462
      Width           =   1095
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "¼ç¦b«È¤á¡G"
      Height          =   255
      Index           =   3
      Left            =   150
      TabIndex        =   34
      Top             =   1470
      Width           =   1215
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "¥Ó  ½Ð  ¤H¡G"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   5
      Top             =   810
      Width           =   1215
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "¥N  ²z  ¤H¡G"
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   6
      Top             =   1140
      Width           =   1215
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "¾÷Ãö¥N¸¹¡G"
      Height          =   255
      Index           =   2
      Left            =   150
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "¦C¦L(&P)"
      Height          =   400
      Left            =   3750
      TabIndex        =   20
      Top             =   180
      Width           =   800
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "µ²§ô(&X)"
      Height          =   400
      Left            =   4575
      TabIndex        =   21
      Top             =   180
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1470
      MaxLength       =   9
      TabIndex        =   8
      Text            =   "X"
      Top             =   802
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1470
      MaxLength       =   9
      TabIndex        =   9
      Text            =   "Y"
      Top             =   1132
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   1470
      MaxLength       =   5
      TabIndex        =   10
      Top             =   1792
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   1470
      TabIndex        =   11
      Text            =   "1"
      Top             =   2100
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   4
      Left            =   1470
      MaxLength       =   1
      TabIndex        =   12
      Top             =   2400
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "³]©w"
      Height          =   660
      Left            =   270
      TabIndex        =   22
      Top             =   3930
      Width           =   4728
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   17
         Top             =   240
         Width           =   3840
      End
      Begin VB.Label Label2 
         Caption         =   "¦Lªí¾÷"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   23
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   6
      Left            =   1050
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1710
      TabIndex        =   18
      Top             =   4650
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   1710
      TabIndex        =   19
      Top             =   4950
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   9
      Left            =   1050
      MaxLength       =   3
      TabIndex        =   1
      Top             =   390
      Width           =   525
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   10
      Left            =   1650
      MaxLength       =   6
      TabIndex        =   2
      Top             =   390
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   11
      Left            =   2730
      MaxLength       =   1
      TabIndex        =   3
      Top             =   390
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   12
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   4
      Top             =   390
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   5
      Left            =   2445
      MaxLength       =   1
      TabIndex        =   16
      Text            =   "Y"
      Top             =   3690
      Width           =   375
   End
   Begin MSForms.TextBox textFM2 
      Height          =   300
      Index           =   2
      Left            =   1470
      TabIndex        =   15
      Top             =   3360
      Width           =   2310
      VariousPropertyBits=   671107099
      Size            =   "4075;529"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textFM2 
      Height          =   300
      Index           =   1
      Left            =   1470
      TabIndex        =   14
      Top             =   3030
      Width           =   2310
      VariousPropertyBits=   671107099
      Size            =   "4075;529"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textFM2 
      Height          =   300
      Index           =   0
      Left            =   1470
      TabIndex        =   13
      Top             =   2700
      Width           =   2310
      VariousPropertyBits=   671107099
      Size            =   "4075;529"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   49
      Top             =   1800
      Width           =   3000
      Caption         =   "lblFM2"
      Size            =   "5292;450"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   16
      Left            =   2640
      TabIndex        =   48
      Top             =   1470
      Width           =   3000
      Caption         =   "lblFM2"
      Size            =   "5292;450"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   47
      Top             =   1140
      Width           =   3000
      Caption         =   "lblFM2"
      Size            =   "5292;450"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   46
      Top             =   810
      Width           =   3000
      Caption         =   "lblFM2"
      Size            =   "5292;450"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ápµ¸¤H³¡ªù¡G                                                     (¤é¤å¥Î)"
      Height          =   180
      Index           =   9
      Left            =   390
      TabIndex        =   33
      Top             =   3390
      Width           =   4125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¦C¦L¥÷¼Æ¡G"
      Height          =   180
      Index           =   0
      Left            =   390
      TabIndex        =   32
      Top             =   2175
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "»y¤å¡G                            (1.¤¤¤å 2.­^¤å )"
      Height          =   180
      Index           =   0
      Left            =   390
      TabIndex        =   31
      Top             =   2490
      Width           =   3000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¥»©Ò®×¸¹¡G"
      Height          =   180
      Index           =   5
      Left            =   165
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "¾î¶b°¾²¾­È(X)¡G¡@¡@¡@¡@¡@¡@(³æ¦ì¤½¤À)"
      Height          =   180
      Index           =   0
      Left            =   300
      TabIndex        =   29
      Top             =   4710
      Width           =   3240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Áa¶b°¾²¾­È(Y)¡G¡@¡@¡@¡@¡@¡@(³æ¦ì¤½¤À)"
      Height          =   180
      Index           =   1
      Left            =   300
      TabIndex        =   28
      Top             =   5010
      Width           =   3240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¥»©Ò®×¸¹¡G"
      Height          =   180
      Index           =   6
      Left            =   150
      TabIndex        =   27
      Top             =   450
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   1260
      X2              =   3510
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ápµ¸¤H1¡G"
      Height          =   180
      Index           =   7
      Left            =   390
      TabIndex        =   26
      Top             =   2790
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ápµ¸¤H2¡G"
      Height          =   180
      Index           =   8
      Left            =   390
      TabIndex        =   25
      Top             =   3090
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¬O§_§t¤£±HÂø»xªº¹ï¶H¡G¡@¡@¡@(Y¡G§t)"
      Height          =   180
      Index           =   4
      Left            =   390
      TabIndex        =   24
      Top             =   3720
      Width           =   3840
   End
End
Attribute VB_Name = "frm083014"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/05/02 §ï¦¨Form 2.0 ¦C¦L¡G³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
'Memo by Lydia 2022/05/02 §ï¦¨Form2.0 ; MGrid1§ï¦r«¬=·s²Ó©úÅé-ExtB ; Label1(1)=>lblFM2(0), Label1(2)=>lblFM2(1), Label1(17)=>lblFM2(16), Label1(3)=>lblFM2(3), Text1(13)=>textFM2(0), Text1(14)=>textFM2(1), Text1(15)=>textFM2(2)
'Memo By Sindy 2012/12/3 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo By Sindy 2011/2/17 SQLDate¤wÀË¬d
'Memo By Sindy 2010/11/25 ­û¤u½s¸¹Äæ¤w­×§ï
'Memo by Morgan 2010/7/27 ¤é´ÁÄæ¤w­×§ï
Option Explicit

Public m_ContactNo As String '±µ¬¢¤H½s¸¹

' ¬y¤ô¸¹
Dim m_PageNo As Integer
'******** 90.11.14   nick
Dim m_PrinterName As String
Dim Prn As Printer
'******** 91.08.06   nick
Dim Rs083 As New ADODB.Recordset
Dim StrSQL083 As String
Dim tmp083 As String
'Add By Cheng 2003/12/22
Dim tmp083_1 As String
'End
'**************************
Dim strSql As String, SeekPrint As Integer, SeekPrintL As Integer, j As Integer, i As Integer
'Add By Cheng 2002/12/20
Dim m_CaseNo As String
'Add By Cheng 2002/02/30
Dim m_dbl_LeftMargin  As Double '¾î¶b°¾²¾­È
Dim m_dbl_TopMargin  As Double 'Áa¶b°¾²¾­È
'Add by Morgan 2004/11/4
Dim m_Contact1(1 To 3) As String 'Ápµ¸¤H1
Dim m_Contact2(1 To 3) As String 'Ápµ¸¤H2
Dim m_ContactDep As String 'Ápµ¸¤H³¡ªù
Dim m_ContactDep2 As String 'Ápµ¸¤H³¡ªù(¦C¦L¥Î)
'Added by Lydia 2017/11/03
Public iStiu As Integer ' 1=A4¦a§}±ø¤¶­±
Dim mESeqNo As String '¼È¦sTB½s¸¹
Dim strPrinter As String '­ì¥»¹w³]¦Lªí¾÷
Dim bolSetDone As Boolean 'Added by Morgan 2017/11/15
Public m_InputNo As String '©I¥s¶Ç¤Jªº½s¸¹ Added by Morgan 2017/11/14
Dim Xo As Long, Yo As Long 'Added by Lydia 2022/05/02
Dim oControl As Object 'Added by Lydia 2022/05/02
Dim strTmp As String 'Added by Lydia 2023/11/08
'Added by Lydia 2025/10/08 ¦Ë¾ä¦a§}±ø«ü©wÁpµ¸¤H¸ê®Æ
Public p_FScon1 As String 'Ápµ¸¤H1
Public p_FScon2 As String 'Ápµ¸¤H2
Public p_FSconDept As String  'Ápµ¸¤H³¡ªù
Public p_SpecLan As String '«ü©w«È¤á©w½Z¬°¤é¤å
'end 2025/10/08

' ³]©w¦C¦L¦a§}±øªº¬y¤ô¸¹ (¤j©óµ¥©ó1)
Public Sub SetPageNo(ByVal nPageNo As Integer)
   m_PageNo = nPageNo
End Sub

' ³]©w¦C¦Lªº¦Lªí¾÷
Public Sub SetPrinter(ByVal strPrinterName As String)
   'Dim m_PrinterName As String
   m_PrinterName = strPrinterName
End Sub

'Add By Cheng 2002/12/20
'³]©w¥»©Ò®×¸¹
Public Sub SetCaseNo(ByVal strCaseNo As String)
    m_CaseNo = strCaseNo
    'Add By Cheng 2003/01/14
    '­Y°l¥[¸¹¤Î¦h°ê¦hÃþ½X¬Ò¬°0«h¥u¦L
    m_CaseNo = IIf(Right(m_CaseNo, 5) = "-0-00", Left(m_CaseNo, Len(m_CaseNo) - 5), m_CaseNo)
End Sub

Public Sub cmdBack_Click()
    bolToEndByNick = True
    Unload Me
End Sub

Public Sub cmdPrint_Click()
'Add By Cheng 2002/09/10
Dim strTempName As String
   
    'Modify By Cheng 2003/03/31
    '­Y«D¾ã§å¦C¦L¦a§}±ø
    If pub_blnBatchPrintAddress = False Then
      '2008/11/07 add by Toni
      'Modified by Lydia 2016/10/28 +Opt1(4)
      If opt1(0).Value = False And opt1(1).Value = False And opt1(2).Value = False And opt1(3).Value = False And opt1(4).Value = False Then
            MsgBox "½Ð¿é¤J¦C¦L¿ï¶µ"
            Exit Sub
      End If
      'end 2008/11/07
      
        '¿ï¾Ü¥Ó½Ð¤H
        If opt1(0).Value = True Then
           If Text1(0) = "" Or Text1(0) = "X" Then
              Text1(0).SetFocus
              MsgBox "¥Ó½Ð¤H¤£±o¬°ªÅ­È !", vbCritical
              Exit Sub
           End If
           'Add By Cheng 2002/09/10
           'edit by nickc 2007/02/02 ¤£¥Î dll ¤F
           'If objPublicData.GetCustomer(Text1(0), strTempName) = False Then
           If ClsPDGetCustomer(Text1(0), strTempName) = False Then
              Me.lblFM2(0).Caption = ""
              Me.Text1(0).SetFocus
              Text1_GotFocus 0
              Exit Sub
           Else
              If strTempName <> "" Then
                 lblFM2(0).Caption = strTempName
              Else
                 lblFM2(0).Caption = ""
              End If
           End If
        '¿ï¾Ü¥N²z¤H
        ElseIf opt1(1).Value = True Then
           If Text1(1) = "" Or Text1(1) = "Y" Then
              Text1(1).SetFocus
              MsgBox "¥N²z¤H¤£±o¬°ªÅ­È !", vbCritical
              Exit Sub
           End If
           'Add By Cheng 2002/09/10
           'edit by nickc 2007/02/02 ¤£¥Î dll ¤F
           'If objPublicData.GetAgent(Text1(1), strTempName) = False Then
           If ClsPDGetAgent(Text1(1), strTempName) = False Then
              Me.lblFM2(1).Caption = ""
              Me.Text1(1).SetFocus
              Text1_GotFocus 1
              Exit Sub
           Else
              If strTempName <> "" Then
                 lblFM2(1).Caption = strTempName
              Else
                 lblFM2(1).Caption = ""
              End If
           End If
        '¿ï¾Ü¾÷Ãö¥N¸¹
        ElseIf opt1(2).Value = True Then
           If Text1(2) = "" Then
              Text1(2).SetFocus
              MsgBox "¾÷Ãö¥N¸¹¤£±o¬°ªÅ­È !", vbCritical
              Exit Sub
           End If
           'Add By Cheng 2002/09/10
           'edit by nickc 2007/02/05 ¤£¥Î dll ¤F
           'If objLawDll.GetGovName(Text1(2), strTempName) = False Then
           If ClsPDGetGovName(Text1(2), strTempName) = False Then
              Me.lblFM2(3).Caption = ""
              Me.Text1(2).SetFocus
              Text1_GotFocus 2
              Exit Sub
           Else
              If strTempName <> "" Then
                 lblFM2(3).Caption = strTempName
              Else
                 lblFM2(3).Caption = ""
              End If
           End If
           
         '¼ç¦b«È¤á
         'add by Toni 2008/10/27
         ElseIf opt1(3).Value = True Then
            If Text1(16) = "" Or Text1(16) = "R" Then
               Text1(16).SetFocus
               MsgBox "¼ç¦b«È¤á¤£±o¬°ªÅ­È !", vbCritical
               Exit Sub
            End If
            
             If ClsPCUGetContact(Text1(16), strTempName) = False Then
                Me.lblFM2(16).Caption = ""
                Me.Text1(16).SetFocus
                Text1_GotFocus 16
                Exit Sub
            Else
               If strTempName <> "" Then
                  lblFM2(16).Caption = strTempName
               Else
                  lblFM2(16).Caption = ""
               End If
           End If
        End If
        If Text1(3) = "" Then
           Text1(3).SetFocus
           MsgBox "¦C¦L¥÷¼Æ¤£±o¬°ªÅ­È !", vbCritical
           Exit Sub
        End If
        If Text1(4) = "" Then
           Text1(4).SetFocus
           MsgBox "¦C¦L»y¤å¤£±o¬°ªÅ­È !", vbCritical
           Exit Sub
        End If
    End If
    
'Modified by Morgan 2017/11/15 ¥~³¡©I¥s®É¥u­n³]©w¤@¦¸¤~¯à±±¨î¦L¦b¤@¥÷¤å¥ó¤º
If m_InputNo = "" Or Not bolSetDone Then
   '***************  90.11.14  NICKC
   ' ³]©w¦Lªí¾÷
    If IsEmptyText(m_PrinterName) Then
        'Modify by Morgan 2008/5/21 ¹w³]¦Lªí¾÷¤]¥i¿ï
        'If Combo1.ListIndex >= SeekPrint Then
        '    j = Combo1.ListIndex + 1
        'Else
        '    j = Combo1.ListIndex
        'End If
        j = Combo1.ListIndex
        'end 2008/5/21
        Set Printer = Printers(j)
    Else
        For Each Prn In Printers
            If Prn.DeviceName = m_PrinterName Then
                Set Printer = Prn
                Exit For
            End If
        Next
    End If
    Printer.Orientation = 1
   '****************************************
End If

   DoEvents
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/4 ²M°£¬d¸ß¦Lªí°O¿ýÀÉÄæ¦ì
   Screen.MousePointer = vbHourglass
   
   '°O¿ý¿é¤J±ø¥ó
   If Len(Text1(9)) > 0 And Len(Text1(10)) > 0 Then
      If Len(Text1(11)) = 0 Then Text1(11) = "0"
      If Len(Text1(12)) = 0 Then Text1(12) = "00"
      pub_QL05 = pub_QL05 & ";" & Label1(6) & Text1(9) & "-" & Text1(10) & "-" & Text1(11) & "-" & Text1(12) 'Add By Sindy 2010/10/4
   End If
   If opt1(0).Value = True Then
      pub_QL05 = pub_QL05 & ";" & opt1(0).Caption & Text1(0) & lblFM2(0) 'Add By Sindy 2010/10/4
   End If
   If opt1(1).Value = True Then
      pub_QL05 = pub_QL05 & ";" & opt1(1).Caption & Text1(1) & lblFM2(1) 'Add By Sindy 2010/10/4
   End If
   If opt1(2).Value = True Then
      pub_QL05 = pub_QL05 & ";" & opt1(2).Caption & Text1(2) & lblFM2(3) 'Add By Sindy 2010/10/4
   End If
   If opt1(3).Value = True Then
      pub_QL05 = pub_QL05 & ";" & opt1(3).Caption & Text1(16) & lblFM2(16) 'Add By Sindy 2010/10/4
   End If
   If Len(Text1(3)) > 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(0) & Text1(3) 'Add By Sindy 2010/10/4
   End If
   If Len(Text1(4)) > 0 Then
      If Text1(4) = "1" Then
         pub_QL05 = pub_QL05 & ";" & Left(Label2(0), 3) & "¤¤¤å" 'Add By Sindy 2010/10/4
      ElseIf Text1(4) = "2" Then
         pub_QL05 = pub_QL05 & ";" & Left(Label2(0), 3) & "­^¤å" 'Add By Sindy 2010/10/4
      End If
   End If
   If Len(textFM2(0)) > 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & textFM2(0) 'Add By Sindy 2010/10/4
   End If
   If Len(textFM2(1)) > 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(8) & textFM2(1) 'Add By Sindy 2010/10/4
   End If
   If Len(textFM2(2)) > 0 Then
      pub_QL05 = pub_QL05 & ";" & Left(Label1(9), 6) & textFM2(2) 'Add By Sindy 2010/10/4
   End If
   If Len(Text1(5)) > 0 Then
      pub_QL05 = pub_QL05 & ";" & Left(Label1(4), 11) & "§t" 'Add By Sindy 2010/10/4
   End If
   
    'Modify By Cheng 2003/03/31
    '­Y¬°¾ã§å¦C¦L¦a§}±ø
    If pub_blnBatchPrintAddress = True Then
        'Modified by Lydia 2019/04/10 PK: ¨Ï¥ÎªÌ±b¸¹@¹q¸£¦WºÙ(pub_HostName)
        'PrintCaseBatch strUserNum
        PrintCaseBatch strUserNum & "@" & pub_HostName
    '­Y«D¾ã§å¦C¦L¦a§}±ø
    Else
        '­Y¬°¤º±M¤H­û   '2007/9/12 ¥[¤J¤º°Ó¤H­û(¦] frm020311)
        If GetStaffDepartment(strUserNum) = "P12" Or GetStaffDepartment(strUserNum) = "P22" Then
            '­Y¥¼¥Ñ¥~³¡µ{¦¡¶Ç¤J¥»©Ò®×¸¹, ¥B¦³¿é¤J¥»©Ò®×¸¹
            '2008/11/10 Modify Toni
            If m_CaseNo = "" And ((Me.Text1(9).Text <> "" And IsNull(Text1(9)) = True) And Me.Text1(10).Text <> "") Then
            'If m_CaseNo = "" And (Me.Text1(9).Text <> "" And Me.Text1(10).Text <> "") Then
                Me.SetCaseNo Me.Text1(9).Text & "-" & Me.Text1(10).Text & "-" & Left(Me.Text1(11).Text & "0", 1) & "-" & Left(Me.Text1(12).Text & "00", 2)
            End If
        End If
        'Added by Lydia 2016/10/28 +A4¦C¦L
        If opt1(4).Value = True Then
            If PrintCaseBatchA4 = True Then
               FormClear
            End If
        Else
        'end 2016/10/28
            If PrintCase = True Then
               'Add By Cheng 2003/05/20
               '²M°£µe­±Äæ¦ì
               FormClear
            End If
        End If
    End If
    
   If m_InputNo = "" Then 'Added by Morgan 2017/11/15
      '¦L§¹«á¹w³]¦^¹w³]¦Lªí¾÷
      Set Printer = Printers(SeekPrint)
      Printer.Orientation = SeekPrintL
   End If
   
   bolToEndByNick = True
   Screen.MousePointer = vbDefault
   'Add By Cheng 2002/12/20
   '²M°£¶Ç¤Jªº¥»©Ò®×¸¹­È
   m_CaseNo = ""
End Sub

Private Function PrintCase() As Boolean

   Dim i As Integer
   Dim St As String
   Dim Page As Integer
   Dim iPrint As Integer
   Dim IntF As Integer
   Dim PriType As Integer
   Dim j As Integer
   Dim Prn As Printer
   Dim nRow As Integer
   Dim intBgnRow As Integer '°_©l¦C¼Æ
   Dim iCurrentX As Integer, iHeight As Integer
   Dim stAddr() As String '¦a§}±ø¬ÛÃö¸ê®Æ
   
On Error GoTo ErrHand

    '³]©w°¾²¾­È
    m_dbl_LeftMargin = CDbl(Me.Text1(7).Text) * 567
    m_dbl_TopMargin = CDbl(Me.Text1(8).Text) * 567

    PriType = Val(Text1(4).Text)
    'Added by Lydia 2025/10/08 ¦Ë¾ä¦a§}±ø¡G«ü©w¼ç¦b«È¤á¥i¥H§ì¤é¤å
    If PriType = "3" And p_SpecLan = "3" Then
       strExc(0) = "select nvl(pcu26,'N') as jaddr FROM PotCustomer WHERE pcu01='" & Mid(ChangeCustomerL(m_InputNo), 1, 8) & "' and pcu02='0' "
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       If intI = 1 Then
          If "" & RsTemp.Fields("jaddr") = "N" Then
             PriType = "2"
          End If
       End If
    End If
    'end 2025/10/08
    'Added by Lydia 2023/11/27 ±q¦a§}¨Ó¹w³]©w½Z»y¤å¡A¨S¦³¤¤¤å¦a§}§ï­^¤å¦a§}¡Fex.Y55554000
    strExc(0) = ""
    If PriType = 1 And (opt1(0).Value = True Or opt1(1).Value = True Or opt1(3).Value = True) Then
      If opt1(0).Value = True Then
         strExc(0) = "select nvl(cu23,'2') pLang from customer where " & ChgCustomer(Text1(0).Text)
      ElseIf opt1(1).Value = True Then
         strExc(0) = "select nvl(fa17,'2') pLang from fagent where " & ChgFagent(Text1(1).Text)
      ElseIf opt1(3).Value = True Then
         strExc(0) = "select nvl(pcu27,'2') pLang from potcustomer where " & ChgPotCustomer(Text1(16).Text)
      End If
      If strExc(0) <> "" Then
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If "" & RsTemp.Fields("pLang") = "2" Then
               Text1(4).Text = "2"
               PriType = 2
            End If
         End If
      End If
    End If

   'Add by Morgan 2004/11/29
   tmp083 = textFM2(0).Text
   tmp083_1 = textFM2(1).Text
   '2004/11/29
   m_ContactDep2 = textFM2(2).Text 'Add by Morgan 2006/10/24
   
   '¥Ó½Ð¤H
   If opt1(0).Value = True Then
   
      St = ChgCustomer(Text1(0).Text)
      If Text1(5) <> "Y" Then
          St = St & " AND (CU32<>'N' OR CU32 IS NULL)"
      End If
      
      '»y¤å
      Select Case PriType
         Case 1  '7
               strExc(0) = "SELECT Nvl(CU80,CU30), Decode(CU80, Null, DECODE(CU31,NULL,SUBSTR(CU23,1,20),SUBSTR(CU31,1,20)),''), Decode(CU80, Null, DECODE(CU31,NULL,SUBSTR(CU23,21,15),SUBSTR(CU31,21,15)),'')," & _
               "nvl(SUBSTR(nvl(cu104,CU04),1,20),CU05),decode(nvl(cu104,CU04),null,CU06,SUBSTR(nvl(cu104,CU04),21,20)),CU08,CU01||CU02 FROM CUSTOMER " & _
               "WHERE " & St
            Page = 1
         Case 2  '11
            '***************************** start
            ' ªô¤p©j»¡Ápµ¸¤H¥u§ì1¤£§ìÁpµ¸¤H2
            ' 910806   nickc   ­Y¦³¥»©Ò®×¸¹«h¥ý§ì°ò¥»ÀÉÁpµ¸¤H¡A­Y¨S¦³«h¡A¥N²z¤H¦³­È´N§ì¥N²z¤HªºÁpµ¸¤H¡A­Y¨S¦³«h§ì¥Ó½Ð¤Hªº¥N²z¤H¡A­Y³£¨S¦³´NªÅ¥Õ¡A¦ý¤£¯à¸õ¦æ¡A­Y¨S¦³¥»©Ò®×¸¹´N§ì¥Ó½Ð¤H©Î¥N²z¤HªºÁpµ¸¤H
            'strExc(0) = "SELECT CU05,CU88,CU89,CU90,DECODE(CU65,'',CU24,CU65)," & _
               "DECODE(CU65,'',CU25,CU66),DECODE(CU65,'',CU26,CU67)," & _
               "DECODE(CU65,'',CU27,CU68),DECODE(CU65,'',CU28,CU69)," & _
               "DECODE(CU65,'',CU102,''),CU01||CU02 FROM CUSTOMER WHERE " & St
            '­Y¦³¥»©Ò®×¸¹
            If Trim(Text1(6).Text) <> "" Then
            
               'Modify by Morgan 2007/5/4 Ápµ¸¤H§ìµe­±ªº¤£¥²¦A­«§ì¡A§_«h­Y¦³¸ê®Æ®ÉµLªk²M°£¤£¦L
               'Modify By Sindy 2016/1/27 ¬ü°ê¥u§ìÁpµ¸¤H1
                strExc(0) = "SELECT '" & ChgSQL(tmp083) & "','" & IIf(Left(ChangeCustomerL(GetPrjNationNumber1(Text1(1))), 3) = "101", "", ChgSQL(tmp083_1)) & "', CU05,CU88,CU89,CU90,Decode(CU80, Null, DECODE(CU65,'',CU24,CU65), CU80)," & _
                   "Decode(CU80, Null, DECODE(CU65,'',CU25,CU66), ''), Decode(CU80, Null, DECODE(CU65,'',CU26,CU67), '')," & _
                   "Decode(CU80, Null, DECODE(CU65,'',CU27,CU68), ''), Decode(CU80, Null, DECODE(CU65,'',CU28,CU69), '')," & _
                   "Decode(CU80, Null, DECODE(CU65,'',CU102,''), '') FROM CUSTOMER WHERE " & St
                
            '­Y¥¼¿é¤J¥»©Ò®×¸¹
            Else
               'Modify by Morgan 2007/5/4 Ápµ¸¤H§ìµe­±ªº¤£¥²¦A­«§ì¡A§_«h­Y¦³¸ê®Æ®ÉµLªk²M°£¤£¦L
               'strExc(0) = "SELECT " & IIf(tmp083 = "", "cu59", "'" & ChgSQL(tmp083) & "'") & "," & IIf(tmp083_1 = "", "CU62", "'" & ChgSQL(tmp083_1) & "'") & ",CU05,CU88,CU89,CU90, Decode(CU80, Null, DECODE(CU65,'',CU24,CU65), CU80),"
               'Modify By Sindy 2016/1/27 ¬ü°ê¥u§ìÁpµ¸¤H1
               'Modified by Morgan 2017/11/14 +¥[Ápµ¸¤H³¡ªù
               strExc(0) = "SELECT '" & ChgSQL(tmp083) & "','" & ChgSQL(m_ContactDep2) & "','" & IIf(Left(ChangeCustomerL(GetPrjNationNumber1(Text1(1))), 3) = "101", "", ChgSQL(tmp083_1)) & "',CU05,CU88,CU89,CU90, Decode(CU80, Null, DECODE(CU65,'',CU24,CU65), CU80),"
               'End 2007/5/4
               strExc(0) = strExc(0) & "Decode(CU80, Null, DECODE(CU65,'',CU25,CU66), ''), Decode(CU80, Null, DECODE(CU65,'',CU26,CU67), '')," & _
                     "Decode(CU80, Null, DECODE(CU65,'',CU27,CU68), ''), Decode(CU80, Null, DECODE(CU65,'',CU28,CU69), '')," & _
                     "Decode(CU80, Null, DECODE(CU65,'',CU102,''), '') FROM CUSTOMER WHERE " & St
               
            End If
            '***************************************************   end
            Page = 2
         Case 3  '6
            'Add By Cheng 2004/04/01
            '­Y¦³¿é¤J¥»©Ò®×¸¹
            If Me.Text1(6).Text <> "" Then
               'Modify by Morgan 2007/5/4 Ápµ¸¤H§ìµe­±ªº¤£¥²¦A­«§ì¡A§_«h­Y¦³¸ê®Æ®ÉµLªk²M°£¤£¦L
               'Modify by Morgan 2006/10/25 ¥[Ápµ¸¤H³¡ªù
                strExc(0) = "SELECT Decode(CU80, Null, SUBSTR(CU29,1,20), CU80), Decode(CU80, Null, SUBSTR(CU29,21,15), '')," & _
                                "SUBSTR(CU06,1,20),SUBSTR(CU06,21,20)" & IIf(m_ContactDep2 <> "", ",'" & ChgSQL(m_ContactDep2) & "'", "") & ",'" & ChgSQL(tmp083) & "','" & ChgSQL(tmp083_1) & "', CU01||CU02 FROM Customer " & _
                                "WHERE " & St
            '­Y¥¼¿é¤J¥»©Ò®×¸¹
            Else
               'Modify by Morgan 2004/11/30 Ápµ¸¤H¹w³]§ìµe­±¿é¤J
               'Modify by Morgan 2006/10/25 ¥[Ápµ¸¤H³¡ªù
               strExc(0) = "SELECT Decode(CU80, Null, SUBSTR(CU29,1,20), CU80), Decode(CU80, Null, SUBSTR(CU29,21,15), ''),"
               'Modify by Morgan 2007/5/4 Ápµ¸¤H§ìµe­±ªº¤£¥²¦A­«§ì¡A§_«h­Y¦³¸ê®Æ®ÉµLªk²M°£¤£¦L
               'strExc(0) = strExc(0) & "SUBSTR(CU06,1,20),SUBSTR(CU06,21,20), " & IIf(m_ContactDep2 = "", "CU114", "'" & ChgSQL(m_ContactDep2) & "'") & "," & IIf(tmp083 = "", "cu60", "'" & ChgSQL(tmp083) & "'") & ", " & IIf(tmp083_1 = "", "cu63", "'" & ChgSQL(tmp083_1) & "'") & ",CU01||CU02 FROM CUSTOMER WHERE " & St
               strExc(0) = strExc(0) & "SUBSTR(CU06,1,20),SUBSTR(CU06,21,20), '" & ChgSQL(m_ContactDep2) & "','" & ChgSQL(tmp083) & "','" & ChgSQL(tmp083_1) & "',CU01||CU02 FROM CUSTOMER WHERE " & St
            End If
            Page = 1
      End Select
   '¥N²z¤H
   ElseIf opt1(1).Value = True Then
      St = ChgFagent(Text1(1).Text)
      If Text1(5) <> "Y" Then
         St = St & " AND (FA24<>'N' OR FA24 IS NULL)"
      End If
      Select Case PriType
         Case 1  '5
            'Modified by Morgan 2019/2/26
            '¤¤¤å¦a§}¥i¯à¶W¹L35¦r(¤¤¼Æ¦r²V¦X)¡A¤£¥iºIÂ_¡CEx:FCP-59172(Y55054)
            'Ápµ¸¤H1À³¸Ó¤]­n¦L
            'strExc(0) = "SELECT SUBSTR(FA17,1,20),SUBSTR(FA17,21,15)," & _
               "SUBSTR(FA04,1,20),SUBSTR(FA04,21,20),FA01||FA02 FROM FAGENT " & _
               "WHERE " & St
            'Page = 4
            'Modified by Lydia 2025/10/09 ©w½Z§O¬°¤¤¤å¡A¦ý¬O¥u¦³­^¤å¤½¥q¦WºÙ¡FY55713
            'strExc(0) = "SELECT SUBSTR(FA17,1,20),SUBSTR(FA17,21)," & _
               "SUBSTR(FA04,1,20),SUBSTR(FA04,21),'" & ChgSQL(tmp083) & "','',FA01||FA02 FROM FAGENT " & _
               "WHERE " & St
            strExc(0) = "SELECT SUBSTR(FA17,1,20),SUBSTR(FA17,21)," & _
               "DECODE(FA04,NULL,FA05,SUBSTR(FA04,1,20)) NAME1,DECODE(FA04,NULL,FA63,SUBSTR(FA04,21)) AS NAME2,'" & ChgSQL(tmp083) & "','',FA01||FA02 FROM FAGENT " & _
               "WHERE " & St
               
            Page = 1
            'end 2019/2/26
            
         Case 2  '11
            '**************************** start
            ' ªô¤p©j»¡Ápµ¸¤H¥u§ì1¤£§ìÁpµ¸¤H2
            ' 910806   nickc   ­Y¦³¥»©Ò®×¸¹«h¥ý§ì°ò¥»ÀÉ¡A­Y¥N²z¤H¦³­È´N§ì¥N²z¤HªºÁpµ¸¤H¡A­Y¨S¦³«h§ì¥Ó½Ð¤Hªº¥N²z¤H¡A­Y³£¨S¦³´NªÅ¥Õ¡A¦ý¤£¯à¸õ¦æ¡A­Y¨S¦³¥»©Ò®×¸¹´N§ì¥Ó½Ð¤H©Î¥N²z¤HªºÁpµ¸¤H
            'strExc(0) = "SELECT FA05,FA63,FA64,FA65," & _
               "FA18,FA19,FA20,FA21,FA22,FA70,FA01||FA02 FROM FAGENT,NATION WHERE " & St & " AND FA55=NA01(+)"
            If Trim(Text1(6).Text) <> "" Then
               'Modify by Morgan 2007/5/4 Ápµ¸¤H§ìµe­±ªº¤£¥²¦A­«§ì¡A§_«h­Y¦³¸ê®Æ®ÉµLªk²M°£¤£¦L
               'Modify By Sindy 2016/1/27 ¬ü°ê¥u§ìÁpµ¸¤H1
                strExc(0) = "SELECT '" & ChgSQL(tmp083) & "', '" & IIf(Left(ChangeCustomerL(GetPrjNationNumber(Text1(1))), 3) = "101", "", ChgSQL(tmp083_1)) & "', FA05,FA63,FA64,FA65," & _
                   " Decode(FA32,Null,FA18,FA32), Decode(FA32,Null,FA19,FA33), Decode(FA32,Null,FA20,FA34), Decode(FA32,Null,FA21,FA35), Decode(FA32,Null,FA22,FA36), Decode(FA32,Null,FA70) FROM FAGENT,NATION WHERE " & St & " AND FA55=NA01(+)"
            Else
                  'Modify by Morgan 2007/5/4 Ápµ¸¤H§ìµe­±ªº¤£¥²¦A­«§ì¡A§_«h­Y¦³¸ê®Æ®ÉµLªk²M°£¤£¦L
                  'strExc(0) = "SELECT " & IIf(tmp083 = "", "FA08", "'" & ChgSQL(tmp083) & "'") & ", " & IIf(tmp083_1 = "", "FA53", "'" & ChgSQL(tmp083_1) & "'") & ",FA05,FA63,FA64,FA65,"
                  'Modify By Sindy 2016/1/27 ¬ü°ê¥u§ìÁpµ¸¤H1
                  'Modified by Morgan 2017/11/14 +¥[Ápµ¸¤H³¡ªù
                  strExc(0) = "SELECT '" & ChgSQL(tmp083) & "','" & ChgSQL(m_ContactDep2) & "', '" & IIf(Left(ChangeCustomerL(GetPrjNationNumber(Text1(1))), 3) = "101", "", ChgSQL(tmp083_1)) & "',FA05,FA63,FA64,FA65,"
                  'end 2007/5/4
                  strExc(0) = strExc(0) & " Decode(FA32,Null,FA18,FA32), Decode(FA32,Null,FA19,FA33), Decode(FA32,Null,FA20,FA34), Decode(FA32,Null,FA21,FA35), Decode(FA32,Null,FA22,FA36), Decode(FA32,Null,FA70) FROM FAGENT,NATION WHERE " & St & " AND FA55=NA01(+)"
            End If
            '****************************** end
            Page = 2
         Case 3  '5
            'Add By Cheng 2004/04/01
            '­Y¦³¿é¤J¥»©Ò®×¸¹
            If Me.Text1(6).Text <> "" Then
               'Modify by Morgan 2007/5/4 Ápµ¸¤H§ìµe­±ªº¤£¥²¦A­«§ì¡A§_«h­Y¦³¸ê®Æ®ÉµLªk²M°£¤£¦L
               'Modify by Morgan 2006/10/25 ¥[Ápµ¸¤H³¡ªù
               strExc(0) = "SELECT SUBSTR(FA23,1,20),SUBSTR(FA23,21,15)," & _
                                "SUBSTR(FA06,1,20),SUBSTR(FA06,21,20)" & IIf(m_ContactDep2 <> "", ",'" & ChgSQL(m_ContactDep2) & "'", "") & ",'" & ChgSQL(tmp083) & "','" & ChgSQL(tmp083_1) & "', FA01||FA02 FROM FAGENT " & _
                                "WHERE " & St
            '­Y¥¼¿é¤J¥»©Ò®×¸¹
            Else
               'Modify by Morgan 2004/11/30 Ápµ¸¤H¹w³]§ìµe­±¿é¤J
               'Modify by Morgan 2006/10/25 ¥[Ápµ¸¤H³¡ªù
                'Modify by Morgan 2007/5/4 Ápµ¸¤H§ìµe­±ªº¤£¥²¦A­«§ì¡A§_«h­Y¦³¸ê®Æ®ÉµLªk²M°£¤£¦L
                'strExc(0) = "SELECT SUBSTR(FA23,1,20),SUBSTR(FA23,21,15),SUBSTR(FA06,1,20),SUBSTR(FA06,21,20), " & IIf(m_ContactDep2 = "", "FA78", "'" & ChgSQL(m_ContactDep2) & "'") & ", " & IIf(tmp083 = "", "FA09", "'" & ChgSQL(tmp083) & "'") & ", " & IIf(tmp083_1 = "", "FA54", "'" & ChgSQL(tmp083_1) & "'") & ", FA01||FA02 FROM FAGENT WHERE " & St
                strExc(0) = "SELECT SUBSTR(FA23,1,20),SUBSTR(FA23,21,15),SUBSTR(FA06,1,20),SUBSTR(FA06,21,20), '" & ChgSQL(m_ContactDep2) & "', '" & ChgSQL(tmp083) & "','" & ChgSQL(tmp083_1) & "', FA01||FA02 FROM FAGENT WHERE " & St
                'end 2007/5/4
            End If
            Page = 1
      End Select
   '¾÷Ãö¤å¸¹
   ElseIf opt1(2).Value = True Then '5
      strExc(0) = "SELECT OR05,SUBSTR(OR06,1,15),SUBSTR(OR06,16,15),OR02,OR01 " & _
         "FROM ORGANIZATION WHERE OR01='" & Text1(2).Text & "'"
      Page = 4
   '¼ç¦b«È¤á
   '2008/10/28 add by toni
   ElseIf opt1(3).Value = True Then
      St = ChgPotCustomer(Text1(16).Text)
      If Text1(5) <> "Y" Then
         St = St & " AND (PCU34<>'N' OR PCU34 IS NULL)"
      End If
      'Mark Lydia 2020/11/17 ¼ç¦b«È¤á(¤¤)¨Ì¦a§}§PÂ_©w½Z»y¤å;¦]¬°¤¤¤åª©­±¬O¤@¦æ¦C¦L (P.S³Ì¦n¬O°w¹ï¯S©w«È¤á)
'      If Opt1(3).Value = True And PriType = "1" Then
'         strExc(1) = "select Pcu01,Pcu02,Pcu36, Decode(Pcu27,Null,Decode(Pcu20,Null,'3','2'),'1') Ntype from potcustomer " & _
'                          "WHERE " & St
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
'         If intI = 1 Then
'              If "" & RsTemp.Fields("ntype") = "2" Then '¤À¦¨­^¤å©M¤¤¤åª©­±
'                  PriType = 2
'              Else
'                  PriType = 1
'              End If
'         End If
'      End If
      'end 2020/11/17
         
      '»y¤å
      Select Case PriType
         Case 1  '7
'         '¤¤¤å¦a§},¤¤¤å¦WºÙ
'         strExc(0) = "SELECT decode(pcu39,NULL, SUBSTR(PCU27,1,20)),decode(pcu39,NULL,SUBSTR(PCU27,21,15))," & _
'         "SUBSTR(PCU08,1,20),SUBSTR(PCU08,21,20), PCU01||PCU02 FROM PotCustomer " & _
'         "WHERE " & St
         '2008/12/04 ¼W¥[SUBSTR(PCU08,41,20)¸Ñ¨M¤¤¤å¦a§}¹Lªø°ÝÃD add by Toni
         'Modify by Morgan 2007/5/4 +Ápµ¸¤H§ìµe­±
         'strExc(0) = "SELECT decode(pcu39,NULL, SUBSTR(PCU27,1,20)),decode(pcu39,NULL,SUBSTR(PCU27,21,15))," & _
         "SUBSTR(PCU08,1,20),SUBSTR(PCU08,21,20),SUBSTR(PCU08,41,20),  PCU01||PCU02 FROM PotCustomer " & _
         "WHERE " & St
         'Modified by Lydia 2020/11/17 2021¦Ë¾äªº¼ç¦b«È¤áªº©w½Z»y¤å:1-¤¤¤å,¦ý¬O¨S¦³¤¤¤å¸ê®Æ; »y¤å¶¶§Ç:¤¤¤å>­^¤å>¤é¤å
         'strExc(0) = "SELECT '" & ChgSQL(tmp083) & "' as a1,'" & ChgSQL(tmp083_1) & "' as a2, decode(pcu39,NULL, PCU27,null) faddr," & _
                          "PCU08 as fname FROM PotCustomer WHERE " & St
         strExc(0) = "SELECT " & IIf(tmp083 <> "", "'" & ChgSQL(tmp083) & "'", "nvl(pcu08 ,decode(pcu03,null,pcu07,pcu03||' '||pcu04||' '||pcu05||' '||pcu06))") & " as a1,'" & _
                           ChgSQL(tmp083_1) & " ' as a2, decode(pcu39,NULL, nvl(PCU27,decode(PCU20,NULL,pcu26 ,pcu20||' '||pcu21||' '||pcu22||' '||pcu23||' '||pcu24||' '||pcu25)),null) faddr," & _
                          "PCU08 as fname FROM PotCustomer WHERE " & St
                          
'          strExc(0) = "SELECT SUBSTR(PCU27,1,20),SUBSTR(PCU27,21,15)," & _
'               "SUBSTR(PCU08,1,20),SUBSTR(PCU08,21,20),PCU01||PCU02 FROM PotCustomer " & _
'               "WHERE " & St
'
           'Page = 4
           '2008/12/04 ¼W¥[SUBSTR(PCU08,41,20)¸Ñ¨M¤¤¤å¦a§}¹Lªø°ÝÃD add by Toni
           'Modified by Lydia 2018/11/01 §ïª©­±
           'Page = 3
           Page = 1
         Case 2  '11
               '­^¤å
               'Modify by Morgan 2007/5/4 Ápµ¸¤H§ìµe­±ªº¤£¥²¦A­«§ì¡A§_«h­Y¦³¸ê®Æ®ÉµLªk²M°£¤£¦L
               'strExc(0) = "SELECT " & IIf(tmp083 = "", "cu59", "'" & ChgSQL(tmp083) & "'") & "," & IIf(tmp083_1 = "", "CU62", "'" & ChgSQL(tmp083_1) & "'") & ",CU05,CU88,CU89,CU90, Decode(CU80, Null, DECODE(CU65,'',CU24,CU65), CU80),"
               strExc(0) = "SELECT '" & ChgSQL(tmp083) & "','" & ChgSQL(tmp083_1) & "',PCU03,Decode(PCU39,Null, DECODE(PCU29,'',PCU20,Pcu29),PCU39),"
               strExc(0) = strExc(0) & " DECODE(PCU29,'',PCU21,PCU30), DECODE(PCU29,'',PCU22,PCU31)," & _
                     "Decode(PCU39,Null, DECODE(PCU29,'',PCU23,PCU32),PCU39),Decode(PCU39,Null,DECODE(PCU29,'',PCU24,PCU33),PCU39)" & _
                     ",Decode(PCU39,Null,DECODE(PCU29,'',PCU25),PCU39)  FROM PotCustomer WHERE " & St
               
            
            '***************************************************   end
            Page = 2
         'Added by Lydia 2025/10/08
         Case 3  '¤é¤å
            strExc(0) = "SELECT " & IIf(tmp083 <> "", "'" & ChgSQL(tmp083) & "'", "nvl(pcu07,nvl(pcu08 ,decode(pcu03,null,pcu07,pcu03||' '||pcu04||' '||pcu05||' '||pcu06)))") & " as a1,'" & _
                           ChgSQL(tmp083_1) & " ' as a2, decode(pcu39,NULL, nvl(pcu26,nvl(PCU27,decode(PCU20,NULL,pcu26 ,pcu20||' '||pcu21||' '||pcu22||' '||pcu23||' '||pcu24||' '||pcu25))),NULL) faddr," & _
                          "nvl(pcu07,nvl(pcu08 ,decode(pcu03,null,pcu07,pcu03||' '||pcu04||' '||pcu05||' '||pcu06))) as fname FROM PotCustomer WHERE " & St
            Page = 1
         'edn 2025/10/08
      End Select
   End If
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      InsertQueryLog (0) 'Add By Sindy 2010/10/4
      Exit Function
   End If
   
   PrintCase = True '2008/10/31 ADD BY SONIA

   Select Case Page
      Case 1
         IntF = 7
      Case 2
         IntF = 11
      Case 3
         IntF = 6
      Case 4
         IntF = 5
   End Select
   
   'Modified by Morgan 2017/11/15 ¥~³¡©I¥s®É¥u­n³]©w¤@¦¸¤~¯à±±¨î¦L¦b¤@¥÷¤å¥ó¤º
   If m_InputNo = "" Or Not bolSetDone Then
      'Modify by Morgan 2006/4/27 ¥Ø«eXP¦Û©w¯È±i»Ý¤â°Ê³]©w¨Ã±N¦Lªí¾÷¹w³]¬°¸Ó¯È±i
      '95
      If pub_OS = "1" Then
         Printer.Height = 2880
         Printer.Width = 10000
      'NT ¶·¥ýµ²§ô¤å¥ó,§_«h¯È±i¤£·|¥Î³ß¦n³]©w
      Else
         'Modify by Morgan 2008/3/27 §ï§ì³]©w¦nªº­È
         'Printer.Orientation = 1
         'Printer.EndDoc
         Printer.PaperSize = PUB_GetPaperSize(2)
         'end 2008/3/27
      End If
         
      Printer.KillDoc
      
      bolSetDone = True
   End If
   'end 2017/11/15
   
   '¥ªÃä¬É
   iCurrentX = 0 + m_dbl_LeftMargin
    'Modify By Cheng 2003/01/15
    '¦C¦L®æ¦¡³]©w¬°­^¤å¦a§}±ø®æ¦¡, ¦C°ª¤£¦P
   If Page <> 2 Then
      iHeight = 280
   Else
      'Modify by Morgan 2009/4/17 ¦C°ª¤£°÷³¡¤À¦r¥Àªº¤U¥b¬q·|³QºI(Ex:g,y)
      'iHeight = 230
      iHeight = 270
   End If
   Printer.Font.Size = 12
   
    'Add By Cheng 2003/07/03
    '³]©w¦C¦L¦r«¬
    '­Y¬°¤¤¤é¤å
    If Me.Text1(4).Text = "1" Or Me.Text1(4).Text = 3 Then
        Printer.Font.Name = "²Ó©úÅé"
    '­Y¬°­^¤å
    Else
        Printer.Font.Name = "Times New Roman"
    End If
    InsertQueryLog (RsTemp.RecordCount) 'Add By Sindy 2010/10/4
   '¥÷¼Æ
   For j = 1 To Val(Text1(3).Text)
      RsTemp.MoveFirst
      iPrint = 1
      With RsTemp
         Select Case Page
             Case 2
               IntF = .Fields.Count 'Add by Morgan 2007/8/29
               Do While Not .EOF
                  nRow = 0
                  For i = 0 To IntF - 1
                     Printer.CurrentX = iCurrentX
                     If IsNull(.Fields(i)) = False Then
                        If IsEmptyText(.Fields(i)) = False Then
                           ' »y¤å¬°­^¤å®É¤£ªÅ¦æ
                           If Text1(4) = "2" Then
                              Printer.CurrentY = nRow * iHeight + m_dbl_TopMargin
                           Else
                              Printer.CurrentY = i * iHeight + m_dbl_TopMargin
                           End If
                           nRow = nRow + 1
                        End If
                     End If
                     
                     'Added by Lydia 2022/05/02
                     Xo = Printer.CurrentX
                     Yo = Printer.CurrentY
                     'end 2022/05/02
                                          
                     If IsNull(.Fields(i)) = False Then
                        If IsEmptyText(.Fields(i)) = False Then
                           'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                           'Printer.Print .Fields(i)
                           PUB_PrintUnicodeText .Fields(i), Xo, Yo, 0
                        End If
                     End If
                     
                     If i = 9 Then
                        Printer.CurrentX = 3200 + iCurrentX
                        If Text1(4) = "2" Then
                           Printer.CurrentY = (nRow - 1) * iHeight + m_dbl_TopMargin
                        Else
                           Printer.CurrentY = (i - 1) * iHeight + m_dbl_TopMargin
                        End If
                     End If
                  Next
                  
                  'Added by Morgan 2017/11/14
                  If m_InputNo <> "" And PriType = 2 Then
                     Printer.CurrentX = 3200 + iCurrentX
                     Printer.CurrentY = Printer.CurrentY + iHeight
                     If Printer.CurrentY + Printer.TextHeight("Y") < Printer.Height Then
                        Printer.Print m_InputNo
                     ElseIf InStr(UCase(Printer.DeviceName), "PDF") = 1 Then
                        Printer.Print m_InputNo
                     Else
                        Debug.Print m_InputNo
                     End If
                  End If
                  'end 2017/11/14
                  
                  iPrint = iPrint + 1
                  Printer.NewPage
                  .MoveNext
               Loop
            Case Else
               Do While Not .EOF
                  '³]©w°_©l¦C¼Æ
                  Select Case IntF
                     Case 7
                         intBgnRow = 1
                     Case 6
                         intBgnRow = 1.5
                     Case 5
                         intBgnRow = 2
                  End Select
                  nRow = 0
               
                  'Modify by Morgan 2008/8/8 ¥Ó½Ð¤H¤¤¤å¦a§}§ï©I¥s¦@¥Î¨ç¼Æ
                  Erase stAddr
                  If opt1(0).Value = True And Text1(4) = "1" Then
                     ReDim stAddr(6)
                     '«È¤á½s¸¹
                     stAddr(6) = ChangeCustomerL(Text1(0).Text)

                        If PUB_GetAddrRef(stAddr(6), Text1(9).Text, Text1(10).Text, Text1(11).Text, Text1(12).Text, stAddr(3), stAddr(5), stAddr(0), stAddr(1), , m_ContactNo) = True Then
                           If Len(stAddr(1)) > 20 Then
                              stAddr(2) = Mid(stAddr(1), 21)
                              stAddr(1) = Left(stAddr(1), 20)
                           End If
                           If Len(stAddr(3)) > 20 Then
                              stAddr(4) = Mid(stAddr(3), 21)
                              stAddr(3) = Left(stAddr(3), 20)
                           End If
                        End If
                  'Added by Lydia 2018/11/01 ¼ç¦b«È¤á¤¤¤å¦a§}
                  'Modified by Lydia 2025/10/08 ¦Ë¾ä¦a§}±ø¡G«ü©w¼ç¦b«È¤á¥i¥H§ì¤é¤å
                  'ElseIf opt1(3).Value = True And Text1(4) = "1" Then
                  ElseIf opt1(3).Value = True And (Text1(4) = "1" Or p_SpecLan = "3") Then
                       ReDim stAddr(IntF)
                       i = 0
                       For intI = 0 To 3
                           If intI < 2 Then  'Ápµ¸¤H
                                If "" & .Fields(intI) <> "" Then
                                    stAddr(i) = "" & .Fields(intI)
                                    i = i + 1
                                End If
                           Else '¦a§}©M¦WºÙ
                                strExc(1) = "" & .Fields(intI)
                                If GetTextLength(strExc(1)) > 40 Then '¶W¹L20¦r­n´«¦æ,³Ì¦h2¦æ
                                    stAddr(i) = Trim(convForm("" & .Fields(intI), 40))
                                    strExc(1) = Replace(strExc(1), stAddr(i), "")
                                    i = i + 1
                                    stAddr(i) = strExc(1)
                                    i = i + 1
                                Else
                                    stAddr(i) = "" & .Fields(intI)
                                    i = i + 1
                                End If
                           End If
                       Next intI
                  'end 2018/11/01
                  Else
                     ReDim stAddr(IntF - 1)
                     For i = 0 To IntF - 1
                        stAddr(i) = "" & .Fields(i)
                     Next
                  End If
              
               
                  For i = 0 To IntF - 1
                     Printer.CurrentX = iCurrentX
                     If IsEmpty(stAddr(i)) = False Then
                        ' »y¤å¬°­^¤å®É¤£ªÅ¦æ
                        If Text1(4) = "2" Then
                           Printer.CurrentY = (nRow + intBgnRow) * iHeight + m_dbl_TopMargin
                        Else
                           Printer.CurrentY = (i + intBgnRow) * iHeight + m_dbl_TopMargin
                        End If
                        nRow = nRow + 1
                     End If
                     
                     'Added by Lydia 2022/05/02
                     Xo = Printer.CurrentX
                     Yo = Printer.CurrentY
                     'end 2022/05/02
                     
                     If IsEmptyText(stAddr(i)) = False Then
                        If i = 6 Then
                           '«È¤á¥N¸¹«á¥[¥»©Ò®×¸¹
                           Printer.Print stAddr(i) & IIf(m_CaseNo <> "", "¡@( " & m_CaseNo & " )", "")
                        Else
                           '­Y¬°¤¤¤å¦a§}
                           'Modify by Morgan 2009/5/7
                           'If Me.Text1(4).Text = "1" Then
                           If Me.Text1(4).Text = "1" And opt1(0).Value = True Then
                              '¦C¦LÄæ¦ì¬°¤½¥q¦WºÙ©Î­Ó¤H¦WºÙ
                              '93.5.23 ADD BY SONIA
                              strExc(0) = stAddr(i)
                              If strExc(0) = "¼B¬ì¨}¡D·Å©y¬Â" Then
                                 strExc(0) = "·Å©y¬Â"
                              End If
                              '93.5.23 END
                              If i = 3 Then
                                 '­Y¬°­Ó¤H¤á®É
                                 If stAddr(3) <> "" And stAddr(5) = "" Then
                                     'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                                     'Printer.Print strExc(0) & "¡@¡@¡@¡@¡@§g¡@¶v±Ò"
                                     PUB_PrintUnicodeText strExc(0) & "¡@¡@¡@¡@¡@§g¡@¶v±Ò", Xo, Yo, 0
                                 Else
                                     'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                                     'Printer.Print strExc(0)
                                     PUB_PrintUnicodeText strExc(0), Xo, Yo, 0
                                 End If
                              ElseIf i = 5 Then
                                 '­Y¬°¤½¥q¤á®É
                                 If stAddr(3) <> "" And stAddr(5) <> "" Then
                                     'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                                     'Printer.Print stAddr(i) & "¡@¡@¡@¡@¡@§g¡@¶v±Ò"
                                     PUB_PrintUnicodeText stAddr(i) & "¡@¡@¡@¡@¡@§g¡@¶v±Ò", Xo, Yo, 0
                                 Else
                                     'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                                     'Printer.Print stAddr(i)
                                     PUB_PrintUnicodeText stAddr(i), Xo, Yo, 0
                                 End If
                              Else
                                 'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                                 'Printer.Print stAddr(i)
                                 PUB_PrintUnicodeText stAddr(i), Xo, Yo, 0
                              End If
                           '­Y«D¤¤¤å¦a§}
                           Else
                              'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                              'Printer.Print stAddr(i)
                              PUB_PrintUnicodeText stAddr(i), Xo, Yo, 0
                           End If
                        End If
                     End If
                     
                  Next
                  
                  Printer.CurrentX = 4000 + iCurrentX
                  'Printer.CurrentY = (i - 1) * iHeight
                  If Text1(4) = "2" Then
                    'Modify By Cheng 2003/01/15
                    '¥[°_©l¦C¼Æ
'                     Printer.CurrentY = (nRow - 1) * iHeight
                     Printer.CurrentY = ((nRow - 1) + intBgnRow) * iHeight + m_dbl_TopMargin
                  Else
                    'Modify By Cheng 2003/01/15
                    '¥[°_©l¦C¼Æ
'                     Printer.CurrentY = (i - 1) * iHeight
                     Printer.CurrentY = ((i - 1) + intBgnRow) * iHeight + m_dbl_TopMargin
                  End If
                     
                  ' 90.07.12 modify by louis
                  'Printer.Print Format(iPrint, "000000")
                  'Added by Morgan 2017/11/14
                  If m_InputNo <> "" Then
                     If PriType = 2 Then
                        If Printer.CurrentY + Printer.TextHeight("Y") < Printer.Height Then
                           Printer.Print m_InputNo
                        ElseIf InStr(UCase(Printer.DeviceName), "PDF") = 1 Then
                           Printer.Print m_InputNo
                        Else
                           Debug.Print m_InputNo
                        End If
                     End If
                  Else
                  'end 2017/11/14
                     If m_PageNo > 0 Then
                        Printer.Print Format(m_PageNo, "000000")
                     Else
                        Printer.Print Format(iPrint, "000000")
                     End If
                  End If 'Added by Morgan 2017/11/14
                  iPrint = iPrint + 1
                  Printer.NewPage
                  .MoveNext
               Loop
         End Select
      End With
   Next
   
   If m_InputNo = "" Then Printer.EndDoc
   
    Exit Function
ErrHand:
    MsgBox Err.Description
    Resume Next
End Function

Private Sub Form_Load()
'Add By Cheng 2003/01/30
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim ii As Integer
   
   'Added by Lydia 2017/11/03
   If iStiu = 1 Then
       Me.Width = 6300
       Frame2.Left = 0
   Else
       Me.Width = 5800
   End If
   'end 2017/11/03
      
   MoveFormToCenter Me
   Opt1_Click 0
   
   'Modify by Morgan 2008/5/20 §ï¤£±Æ°£¹w³]¦Lªí¾÷
   PUB_SetPrinter Me.Name, Me.Combo1, , , SeekPrint, Me.Text1(7), Me.Text1(8)
   SeekPrintL = Printer.Orientation
   'end 2008/5/20
   
   'Added by Lydia 2017/11/03 A4¦a§}±ø
   If iStiu = 1 Then
       Frame2.Visible = True
       'Modified by Lydia 2017/11/22 +°ê¤º
       Me.Caption = "°ê¤ºA4¦a§}±ø¦C¦L"
       PUB_SetPrinter Me.Name, Me.Combo2, strPrinter, , SeekPrint
       SeekPrintL = Printer.Orientation
       lblCnt.Caption = ""
       If ReadA4List(True) = False Then
       End If
   Else
       Frame2.Visible = False
   'end 2017/11/03
        '¹w³]¥»©Ò®×¸¹ªº¨t²ÎÃþ§O
        'Modify by Morgan 2005/12/14 ¥[±±¨î°]°È¤]¥i¥Î
        If Forms(0).Name = "mdimain" Then
           If Forms(0).intPCaseKind = ±M§Q Then
               Select Case Forms(0).intPWhere
               Case 0 '°ê¤º
                   Me.Text1(9).Text = "P"
               Case 1 '°ê¥~_CF
                   Me.Text1(9).Text = "CFP"
               Case 2 '°ê¥~_FC
                   Me.Text1(9).Text = "FCP"
               End Select
           ElseIf Forms(0).intPCaseKind = °Ó¼Ð Then
               Select Case Forms(0).intPWhere
               Case 0 '°ê¤º
                   Me.Text1(9).Text = "T"
               Case 1 '°ê¥~_CF
                   Me.Text1(9).Text = "CFT"
               Case 2 '°ê¥~_FC
                   Me.Text1(9).Text = "FCT"
               End Select
           End If
        End If
        
        'add by sonia 2016/6/14 ¦]¤¤­^¤å¦a§}±øÁa¶b°¾²¾­È(Y)¦ì¸m¤£¦P, ¬G°]°È³B¤H­û¾Þ§@®É¦Û°Ê¨Ì»y¤å§ïÁa¶b°¾²¾­È(Y)ªº­È
        If Pub_StrUserSt03 = "M31" Then
           MsgBox "°]°È³B¤H­û¾Þ§@®É¡A¨t²Î·|¦Û°Ê¨Ì»y¤å§ïÅÜÁa¶b°¾²¾­È(Y)¡I", vbExclamation + vbOKOnly
        End If
        'end 2016/6/14
   End If
   
   'Added by Lydia 2022/05/02
   For Each oControl In lblFM2
       oControl.Caption = ""
   Next
   For Each oControl In textFM2
       oControl.Text = ""
   Next
   'end 2022/05/02
End Sub

Private Sub Opt1_Click(Index As Integer)
 Dim i As Integer
On Error Resume Next
   'Modify by toni 2008/10/28
   'For i = 0 To 2
   For i = 0 To 3
      If i <> 3 Then
         Text1(i).Enabled = False
      Else
         Text1(16).Enabled = False
      End If
   Next
   'end 2008/10/28

   If Index = 3 Then
      Text1(16).Enabled = True
      Text1(16).SetFocus
   Else
      Text1(Index).Enabled = True
      Text1(Index).SetFocus
   End If
   'end 2008/10/28
   
    Select Case Index
    Case 0
        If Me.Text1(0).Text = "" Then Me.Text1(0).Text = "X"
    Case 1
        If Me.Text1(1).Text = "" Then Me.Text1(1).Text = "Y"
    'add by Toni 2008/10/28
    Case 3
         If Me.Text1(16).Text = "" Then Me.Text1(16).Text = "R"
    'end 2008/10/28
    End Select
End Sub
'Add by Morgan 2004/11/4
Private Sub Text1_Change(Index As Integer)
   '©w½Z»y¤åÅÜ§ó®ÉÁpµ¸¤H¤]­nÅÜ
      
   If Index = 4 Then
      setContact
      'add by sonia 2016/6/14 ¦]¤¤­^¤å¦a§}±øÁa¶b°¾²¾­È(Y)¦ì¸m¤£¦P, ¬G°]°È³B¤H­û¾Þ§@®É¦Û°Ê¨Ì»y¤å§ïÁa¶b°¾²¾­È(Y)ªº­È
      If Pub_StrUserSt03 = "M31" Then
         If Text1(4) = "1" Then
            Text1(8) = "-0.5"
         Else
            Text1(8) = 0
         End If
      End If
      'end 2016/6/14
   End If
   
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Select Case Index
    Case 0 '¥Ó½Ð¤H¥N¸¹
        If Me.Text1(0).Text <> "" Then
            Me.Text1(0).SelStart = 1
        End If
    Case 1 '¥N²z¤H¥N¸¹
        If Me.Text1(1).Text <> "" Then
            Me.Text1(1).SelStart = 1
        End If
   'add by toni 2008/10/27
   '¼ç¦b«È¤á
   Case 16
         If Me.Text1(16).Text <> "" Then
            Me.Text1(16).SelStart = 1
        End If
   'end 2008/10/27
   Case Else
        TextInverse Text1(Index)
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   'Modify by Morgan 2007/5/4 ¥[§PÂ_Ápµ¸¤Hµ¥Äæ¦ì¤£¥Î
   If Index <> 5 And Index <> 13 And Index <> 14 Then
      KeyAscii = UpperCase(KeyAscii)
   End If
   'end 2007/5/4
   Select Case Index
      Case 4
         '2005/7/4 MODIFY BY SONIA ¨ú®ø¤é¤å¦a§}±ø§ï¦L­^¤å¦a§}±ø
         'If (KeyAscii > 51 Or KeyAscii < 49) And KeyAscii <> 8 Then
         If (KeyAscii > 50 Or KeyAscii < 49) And KeyAscii <> 8 Then
         '2005/7/4 END
            KeyAscii = 0
            Beep
         'Add by Morgan 2004/11/30
         Else
            If KeyAscii = 49 Then
               textFM2(0).Text = ""
               textFM2(1).Text = ""
               textFM2(2).Text = ""
               textFM2(0).Enabled = False
               textFM2(1).Enabled = False
               textFM2(2).Enabled = False
               textFM2(0).BackColor = QBColor(7)
               textFM2(1).BackColor = QBColor(7)
               textFM2(2).BackColor = QBColor(7)
            Else
               textFM2(0).Enabled = True
               textFM2(1).Enabled = True
               textFM2(2).Enabled = True
               textFM2(0).BackColor = QBColor(15)
               textFM2(1).BackColor = QBColor(15)
               textFM2(2).BackColor = QBColor(15)
            End If
         End If
      Case 5
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      
   End Select
End Sub

Private Sub getContact(ByVal p_stSQL As String)
   Erase m_Contact1
   Erase m_Contact2
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open p_stSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         'Added by Morgan 2021/4/8 Y52884B10 (SOMAR CORPORATION ª¾ªº°]²£³¡)¦a§}±ø¦C¦L¦¬¥ó¤H©T©w±a¥X¡Gþçû{’U¡@ªø§ø¡@¬L«Û¡@¼Ë --§d±mµÙ
         If opt1(1).Value = True And Text1(1) = "Y52884B10" Then
            'Modified by Lydia 2023/11/08 §ï¥Î¼Ò²Õ¨ú±o
            'm_Contact1(1) = "þçû{’U¡@ªø§ø¡@¬L«Û¡@¼Ë"
            'm_Contact1(2) = "þçû{’U¡@ªø§ø¡@¬L«Û¡@¼Ë"
            'm_Contact1(3) = "þçû{’U¡@ªø§ø¡@¬L«Û¡@¼Ë"
            strTmp = PUB_GetUniText(Me.Name, "«ü©wÁpµ¸¤H1")
            m_Contact1(1) = strTmp
            m_Contact1(2) = strTmp
            m_Contact1(3) = strTmp
            'end 2023/11/08
         Else
         'end 2021/4/8
         
            m_Contact1(1) = "" & .Fields(0)
            m_Contact1(2) = "" & .Fields(1)
            m_Contact1(3) = "" & .Fields(2)
            m_Contact2(1) = "" & .Fields(3)
            m_Contact2(2) = "" & .Fields(4)
            m_Contact2(3) = "" & .Fields(5)
            'Add by Morgan 2006/10/24 Ápµ¸¤H³¡ªù
            m_ContactDep = "" & .Fields(6)
         End If
         
         '2008/10/31 ADD BY SONIA ©w½Z»y¤å
         Me.Text1(4).Text = "" & .Fields(7)
         'Modified by Lydia 2025/10/08 ¦Ë¾ä¦a§}±ø¡G«ü©w¼ç¦b«È¤á¥i¥H§ì¤é¤å
         'If Me.Text1(4).Text = "3" Then
         If Me.Text1(4).Text = "3" And p_SpecLan <> "3" Then
            Me.Text1(4).Text = "2"
         End If
         '2008/10/31 END
      End If
   End With
   CheckOC3
   
ErrHnd:
      If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub setContact()
   'Added by Lydia 2025/10/08 ¦Ë¾ä¦a§}±ø«ü©wÁpµ¸¤H¸ê®Æ¡AÁ×§KForm­«§ì
   If p_FScon1 <> "" Or p_FScon2 <> "" Or p_FSconDept <> "" Then
      textFM2(0).Text = p_FScon1
      textFM2(1).Text = p_FScon2
      textFM2(2).Text = p_FSconDept
   Else
   'end 2025/10/08
      textFM2(0).Text = "": textFM2(1).Text = "": textFM2(2).Text = ""
      Select Case Text1(4).Text
         Case "1"
            textFM2(0).Text = m_Contact1(1)
            textFM2(1).Text = m_Contact2(1)
         Case "2"
            textFM2(0).Text = m_Contact1(2)
            textFM2(1).Text = m_Contact2(2)
         Case "3"
            textFM2(0).Text = m_Contact1(3)
            textFM2(1).Text = m_Contact2(3)
            'Add by Morgan 2006/10/24 ¥[Ápµ¸¤H³¡ªù
            If m_Contact1(3) <> "" Or m_Contact2(3) <> "" Then
               textFM2(2).Text = m_ContactDep
            End If
      End Select
   End If 'Added by Lydia 2025/10/08
End Sub

'Modify By Sindy 2015/8/4
'Private Sub Text1_LostFocus(Index As Integer)
Public Sub Text1_LostFocus(Index As Integer)
'2015/8/4 END
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
    Select Case Index
       'Add by Morgan 2004/11/4
       Case 0, 1
         Screen.MousePointer = vbHourglass
         If Text1(6).Text = "" And Me.opt1(Index).Value = True Then
            If Index = 0 Then
               'Modify by Morgan 2006/10/24 ¥[Ápµ¸¤H³¡ªù
               'strSQL = "SELECT cu58,cu59,cu60,cu61,CU62,cu63 FROM CUSTOMER WHERE " & ChgCustomer(Text1(0).Text)
               '2008/10/31 MODIFY BY SONIA ¥[©w½Z»y¤å
               'strSQL = "SELECT cu58,cu59,cu60,cu61,CU62,cu63,cu114 FROM CUSTOMER WHERE " & ChgCustomer(Text1(0).Text)
               strSql = "SELECT cu58,cu59,cu60,cu61,CU62,cu63,cu114,CU64 FROM CUSTOMER WHERE " & ChgCustomer(Text1(0).Text)
               If Text1(5) <> "Y" Then
                  strSql = strSql & " AND (CU32<>'N' OR CU32 IS NULL)"
               End If
            ElseIf Index = 1 Then
               'Modify by Morgan 2006/10/24 ¥[Ápµ¸¤H³¡ªù
               'strSQL = "SELECT FA07,fa08,FA09,FA52,FA53,FA54 FROM FAGENT WHERE " & ChgFagent(Text1(1).Text)
               '2008/10/31 MODIFY BY SONIA ¥[©w½Z»y¤å
               'strSQL = "SELECT FA07,fa08,FA09,FA52,FA53,FA54,FA78 FROM FAGENT WHERE " & ChgFagent(Text1(1).Text)
               strSql = "SELECT FA07,fa08,FA09,FA52,FA53,FA54,FA78,FA31 FROM FAGENT WHERE " & ChgFagent(Text1(1).Text)
               If Text1(5) <> "Y" Then
                  strSql = strSql & " AND (FA24<>'N' OR FA24 IS NULL)"
               End If
            End If
            getContact strSql
            setContact
         End If
         Screen.MousePointer = vbDefault
      '2004/11/4 end
   
    'Modify By Cheng 2003/03/04
'    Case 6 '¥»©Ò®×¸¹
    Case 12 '¥»©Ò®×¸¹
        '­Y¦³¿é¤J¥»©Ò®×¸¹
        If Me.Text1(9).Text <> "" And Me.Text1(10).Text <> "" Then
            Screen.MousePointer = vbHourglass
            'Modify By Cheng 2004/03/31
            '»y¤å§ìªk : °ò¥»ÀÉ©w½Z»y¤å-->FC¥N²z¤H©w½Z»y¤å-->¥Ó½Ð¤H©w½Z»y¤å
'            'Modify By Cheng 2003/05/20
            '±M§Q
            'Modify by Morgan 2004/11/4
            'strSQLA = "Select PA75, PA26, Nvl(PA85, Nvl(FA31, CU64)) From Patent,Fagent,Customer Where " & ChgPatent(Me.Text1(9).Text & Me.Text1(10).Text & Me.Text1(11).Text & Me.Text1(12).Text) & " And substr(PA75,1,8)=FA01(+) And substr(PA75,9,1)=FA02(+) And substr(PA26,1,8)=CU01(+) And substr(PA26,9,1)=CU02(+) "
            StrSQLa = "Select PA75, PA26, Nvl(PA85, Nvl(FA31, CU64)),PA51,PA52,PA53,PA54,PA55,PA56,PA139 From Patent,Fagent,Customer Where " & ChgPatent(Me.Text1(9).Text & Me.Text1(10).Text & Me.Text1(11).Text & Me.Text1(12).Text) & " And substr(PA75,1,8)=FA01(+) And substr(PA75,9,1)=FA02(+) And substr(PA26,1,8)=CU01(+) And substr(PA26,9,1)=CU02(+) "
            
            '°Ó¼Ð
            'Modify by Morgan 2004/11/4
            'strSQLA = strSQLA & " union Select TM44, TM23, Nvl(TM53, Nvl(FA31, CU64)) From Trademark,Fagent,Customer Where " & ChgTradeMark(Me.Text1(9).Text & Me.Text1(10).Text & Me.Text1(11).Text & Me.Text1(12).Text) & " And substr(TM44,1,8)=FA01(+) And substr(TM44,9,1)=FA02(+) And substr(TM23,1,8)=CU01(+) And substr(TM23,9,1)=CU02(+) "
            StrSQLa = StrSQLa & " union Select TM44, TM23, Nvl(TM53, Nvl(FA31, CU64)),TM38,TM39,TM40,TM41,TM42,TM43,TM76 From Trademark,Fagent,Customer Where " & ChgTradeMark(Me.Text1(9).Text & Me.Text1(10).Text & Me.Text1(11).Text & Me.Text1(12).Text) & " And substr(TM44,1,8)=FA01(+) And substr(TM44,9,1)=FA02(+) And substr(TM23,1,8)=CU01(+) And substr(TM23,9,1)=CU02(+) "
            
            'ªA°È
            'Modify by Morgan 2004/11/4
            'strSQLA = strSQLA & " union Select SP26, SP08, Nvl(SP34, Nvl(FA31, CU64)) From ServicePractice,Fagent,Customer Where " & ChgService(Me.Text1(9).Text & Me.Text1(10).Text & Me.Text1(11).Text & Me.Text1(12).Text) & " And substr(SP26,1,8)=FA01(+) And substr(SP26,9,1)=FA02(+) And substr(SP08,1,8)=CU01(+) And substr(SP08,9,1)=CU02(+) "
            'Modify by Morgan 2005/10/11 ¥[sp30
            'StrSQLa = StrSQLa & " union Select SP26, SP08, Nvl(SP34, Nvl(FA31, CU64)),'','','','','','' From ServicePractice,Fagent,Customer Where " & ChgService(Me.Text1(9).Text & Me.Text1(10).Text & Me.Text1(11).Text & Me.Text1(12).Text) & " And substr(SP26,1,8)=FA01(+) And substr(SP26,9,1)=FA02(+) And substr(SP08,1,8)=CU01(+) And substr(SP08,9,1)=CU02(+) "
            'Modify by Morgan 2010/12/14 ¥[sp75
            'StrSQLa = StrSQLa & " union Select SP26, SP08, Nvl(SP34, Nvl(FA31, CU64)),'',sp30,'','','','',SP71 From ServicePractice,Fagent,Customer Where " & ChgService(Me.Text1(9).Text & Me.Text1(10).Text & Me.Text1(11).Text & Me.Text1(12).Text) & " And substr(SP26,1,8)=FA01(+) And substr(SP26,9,1)=FA02(+) And substr(SP08,1,8)=CU01(+) And substr(SP08,9,1)=CU02(+) "
            StrSQLa = StrSQLa & " union Select SP26, SP08, Nvl(SP34, Nvl(FA31, CU64)),sp30,sp30,sp30,sp75,sp75,sp75,SP71 From ServicePractice,Fagent,Customer Where " & ChgService(Me.Text1(9).Text & Me.Text1(10).Text & Me.Text1(11).Text & Me.Text1(12).Text) & " And substr(SP26,1,8)=FA01(+) And substr(SP26,9,1)=FA02(+) And substr(SP08,1,8)=CU01(+) And substr(SP08,9,1)=CU02(+) "
            'End
            
            '2009/12/10 add by sonia¥[ªk°È¤ÎÅU°Ý
            StrSQLa = StrSQLa & " union Select LC22, LC11, Nvl(FA31, CU64),LC18,LC19,LC20,'','','',LC39 From LAWCASE,Fagent,Customer Where " & ChgLawcase(Me.Text1(9).Text & Me.Text1(10).Text & Me.Text1(11).Text & Me.Text1(12).Text) & " And substr(LC22,1,8)=FA01(+) And substr(LC22,9,1)=FA02(+) And substr(LC11,1,8)=CU01(+) And substr(LC11,9,1)=CU02(+) "
            StrSQLa = StrSQLa & " union Select '', HC05, CU64,'','','','','','','' From HIRECASE,Customer Where " & ChgHirecase(Me.Text1(9).Text & Me.Text1(10).Text & Me.Text1(11).Text & Me.Text1(12).Text) & " and substr(HC05,1,8)=CU01(+) And substr(HC05,9,1)=CU02(+) "
            '2009/12/10 end
            
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                '­Y¦³°ê¥~¥N²z¤H¸ê®Æ
                If "" & rsA.Fields(0).Value <> "" Then
                    Me.opt1(1).Value = True
                    Me.Text1(1).Text = "" & rsA.Fields(0).Value
                    Text1_Validate 1, False
                    'Add By Cheng 2003/05/20
                    '©w½Z»y¤å
                    Me.Text1(4).Text = "" & rsA.Fields(2).Value
                    
                    'Modify by Morgan 2005/7/22 ¤é¤å§ï­^¤å
                    'Modified by Lydia 2025/10/08 ¦Ë¾ä¦a§}±ø¡G«ü©w¼ç¦b«È¤á¥i¥H§ì¤é¤å
                    'If Me.Text1(4).Text = "3" Then
                    If Me.Text1(4).Text And p_SpecLan <> "3" Then
                       Me.Text1(4).Text = "2"
                    End If

                    Me.Text1(0).Text = "": Me.lblFM2(0).Caption = ""
                    Me.Text1(2).Text = "": Me.lblFM2(3).Caption = ""
                    Me.Text1(16).Text = "": Me.lblFM2(16).Caption = ""   '2013/3/5 add by sonia
                '­Y¦³¥Ó½Ð¤H¸ê®Æ
                ElseIf "" & rsA.Fields(1).Value <> "" Then
                    Me.opt1(0).Value = True
                    Me.Text1(0).Text = "" & rsA.Fields(1).Value
                    Text1_Validate 0, False
                    'Add By Cheng 2003/05/20
                    '©w½Z»y¤å
                    Me.Text1(4).Text = "" & rsA.Fields(2).Value
                    
                    'Modify by Morgan 2005/7/22 ¤é¤å§ï­^¤å
                    'Modified by Lydia 2025/10/08 ¦Ë¾ä¦a§}±ø¡G«ü©w¼ç¦b«È¤á¥i¥H§ì¤é¤å
                    'If Me.Text1(4).Text = "3" Then
                    If Me.Text1(4).Text And p_SpecLan <> "3" Then
                       Me.Text1(4).Text = "2"
                    End If
                    
                    Me.Text1(1).Text = "": Me.lblFM2(1).Caption = ""
                    Me.Text1(2).Text = "": Me.lblFM2(3).Caption = ""
                    Me.Text1(16).Text = "": Me.lblFM2(2).Caption = ""   '2013/3/5 add by sonia
               End If
                'Add By Cheng 2003/03/07
                Me.Text1(6).Text = Me.Text1(9).Text & Me.Text1(10).Text & Me.Text1(11).Text & Me.Text1(12).Text
                
               'Added by Morgan 2021/4/8 Y52884B10 (SOMAR CORPORATION ª¾ªº°]²£³¡)¦a§}±ø¦C¦L¦¬¥ó¤H©T©w±a¥X¡Gþçû{’U¡@ªø§ø¡@¬L«Û¡@¼Ë --§d±mµÙ
               If opt1(1).Value = True And Text1(1) = "Y52884B10" Then
                  'Modified by Lydia 2023/11/08 §ï¥Î¼Ò²Õ¨ú±o
                  'm_Contact1(1) = "þçû{’U¡@ªø§ø¡@¬L«Û¡@¼Ë"
                  'm_Contact1(2) = "þçû{’U¡@ªø§ø¡@¬L«Û¡@¼Ë"
                  'm_Contact1(3) = "þçû{’U¡@ªø§ø¡@¬L«Û¡@¼Ë"
                  strTmp = PUB_GetUniText(Me.Name, "«ü©wÁpµ¸¤H1")
                  m_Contact1(1) = strTmp
                  m_Contact1(2) = strTmp
                  m_Contact1(3) = strTmp
                  'end 2023/11/08
               Else
               'end 2021/4/8
               
                  'Add by Morgan 2004/11/4
                  m_Contact1(1) = "" & rsA.Fields("PA51")
                  m_Contact1(2) = "" & rsA.Fields("PA52")
                  m_Contact1(3) = "" & rsA.Fields("PA53")
                  m_Contact2(1) = "" & rsA.Fields("PA54")
                  m_Contact2(2) = "" & rsA.Fields("PA55")
                  m_Contact2(3) = "" & rsA.Fields("PA56")
                  m_ContactDep = "" & rsA.Fields("PA139")
                  
               End If
               
               setContact
               '2004/11/4 end
            Else
                MsgBox "µL¦¹®×¸¹¸ê®Æ!!!", vbExclamation + vbOKOnly
                'Add By Cheng 2003/05/20
                '¿ï¶µ
                Me.opt1(0).Value = True
                Me.Text1(0).Text = ""
                Me.Text1(1).Text = ""
                Me.Text1(2).Text = ""
                Me.lblFM2(0).Caption = ""
                Me.lblFM2(1).Caption = ""
                Me.lblFM2(3).Caption = ""
                '©w½Z»y¤å
                Me.Text1(4).Text = ""
                'Add By Cheng 2003/03/07
                Me.Text1(6).Text = ""
                Me.Text1(9).SetFocus
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            Screen.MousePointer = vbDefault
        Else
            'Add By Cheng 2004/04/01
            Me.Text1(6).Text = ""
            'End
        End If
    'add by Toni 2008/10/28
    Case 16
        If Me.Text1(16).Text <> "" And opt1(3).Value = True Then
            Screen.MousePointer = vbHourglass
            
            StrSQLa = "Select PCC01,PCC02,PCC05,PCC03,PCC04,PCC06,PCU34,PCU36 From PotCustCont,PotCustomer Where " & ChgPotCustomer(Me.Text1(16)) & " AND PCC01=PCU01(+)"
               If Text1(5) <> "Y" Then
                  StrSQLa = StrSQLa & " AND (PCU34<>'N' OR PCU34 IS NULL)"
               End If
            
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
           
                '­Y¦³Ápµ¸¤H¸ê®Æ
                If "" & rsA.Fields(0).Value <> "" Then
                    Me.opt1(3).Value = True
                    Me.Text1(16).Text = "" & rsA.Fields(0).Value
                    'Modified by Lydia 2022/10/11 debug
                    'Text1_Validate 1, False
                    Text1_Validate 16, False
                    'Add By Cheng 2003/05/20
                    '©w½Z»y¤å
                    Me.Text1(4).Text = "" & rsA.Fields(7).Value
                    
                    'Modify by Morgan 2005/7/22 ¤é¤å§ï­^¤å
                    'Modified by Lydia 2025/10/08 ¦Ë¾ä¦a§}±ø¡G«ü©w¼ç¦b«È¤á¥i¥H§ì¤é¤å
                    'If Me.Text1(4).Text = "3" Then
                    If Me.Text1(4).Text And p_SpecLan <> "3" Then
                       Me.Text1(4).Text = "2"
                    End If
                    
                    Me.Text1(0).Text = "": Me.lblFM2(0).Caption = ""
                    Me.Text1(1).Text = "": Me.lblFM2(1).Caption = ""
                    Me.Text1(2).Text = "": Me.lblFM2(3).Caption = ""
                     
                End If
               
                '2008/10/30 add by Toni
               If rsA.RecordCount > 0 Then
               Erase m_Contact1
               Erase m_Contact2
                  'Added by Lydia 2025/10/08 ¦Ë¾ä¦a§}±ø«ü©wÁpµ¸¤H¸ê®Æ¡AÁ×§KForm­«§ì
                  If p_FScon1 <> "" Or p_FScon2 <> "" Or p_FSconDept <> "" Then
                     m_Contact1(1) = p_FScon1
                     m_Contact1(2) = p_FScon2
                     m_Contact1(3) = p_FSconDept
                  Else
                  'end 2025/10/08
                     With rsA
                        For i = 0 To .RecordCount - 1
                        If .RecordCount >= 2 Then
                           .MoveFirst
                           m_Contact1(1) = "" & rsA.Fields("PCC05")
                           m_Contact1(2) = "" & rsA.Fields("PCC03")
                           m_Contact1(3) = "" & rsA.Fields("PCC04")
                           
                           .MoveNext
                           m_Contact2(1) = "" & rsA.Fields("PCC05")
                           m_Contact2(2) = "" & rsA.Fields("PCC03")
                           m_Contact2(3) = "" & rsA.Fields("PCC04")
                        Else
                           m_Contact1(1) = "" & rsA.Fields("PCC05")
                           m_Contact1(2) = "" & rsA.Fields("PCC03")
                           m_Contact1(3) = "" & rsA.Fields("PCC04")
                           Exit For
                        End If
                           
                        Next i
                     End With
                  End If 'Added by Lydia 2025/10/08
                  m_ContactDep = "" & rsA.Fields("PCC06")
              End If
              '2008/10/30
              
              setContact
               '2004/11/4 end
'2013/3/5 cancel by sonia
'            Else
'                MsgBox "µL¦¹Ápµ¸¤H¸ê®Æ!!!", vbExclamation + vbOKOnly
'                'Add By Cheng 2003/05/20
'                '¿ï¶µ
'                Me.opt1(0).Value = True
'                Me.Text1(0).Text = ""
'                Me.Text1(1).Text = ""
'                Me.Text1(2).Text = ""
'                'add by Toni 2008/10/28
'                Me.Text1(16).Text = ""
'                Me.lblfm2(2).Caption = ""
'                'end 2008/10/28
'                Me.lblfm2(0).Caption = ""
'                Me.lblfm2(1).Caption = ""
'                Me.lblfm2(3).Caption = ""
'                '©w½Z»y¤å
'                Me.Text1(4).Text = ""
'                'Add By Cheng 2003/03/07
'                Me.Text1(6).Text = ""
'                Me.Text1(9).SetFocus
'2013/3/5 end
            'Added by Lydia 2022/10/11
            Else  'µLÁpµ¸¤H=>±a¤J¼ç¦b«È¤á¸ê®Æ
               StrSQLa = "Select PCU01,PCU02,PCU34,PCU36 From PotCustCont,PotCustomer Where " & ChgPotCustomer(Me.Text1(16)) & " AND PCU01=PCC01(+)"
               If Text1(5) <> "Y" Then
                  StrSQLa = StrSQLa & " AND (PCU34<>'N' OR PCU34 IS NULL)"
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                    Me.opt1(3).Value = True
                    Me.Text1(16).Text = "" & rsA.Fields("PCU01")
                    Text1_Validate 16, False
                    '©w½Z»y¤å
                    Me.Text1(4).Text = "" & rsA.Fields("PCU36")
                    
                    'Modify by Morgan 2005/7/22 ¤é¤å§ï­^¤å
                    'Modified by Lydia 2025/10/08 ¦Ë¾ä¦a§}±ø¡G«ü©w¼ç¦b«È¤á¥i¥H§ì¤é¤å
                    'If Me.Text1(4).Text = "3" Then
                    If Me.Text1(4).Text And p_SpecLan <> "3" Then
                       Me.Text1(4).Text = "2"
                    End If
                    
                    Me.Text1(0).Text = "": Me.lblFM2(0).Caption = ""
                    Me.Text1(1).Text = "": Me.lblFM2(1).Caption = ""
                    Me.Text1(2).Text = "": Me.lblFM2(3).Caption = ""
               End If
            'end 2022/10/11
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            Screen.MousePointer = vbDefault
        End If
    End Select
End Sub

'2013/3/4 add by sonia
Private Sub Text1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Index = 0 Then
      opt1(1).Value = False: opt1(2).Value = False: opt1(3).Value = False
      opt1(0).Value = True
   End If
End Sub
'2013/3/4 end

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
 Dim strTempName As String, St As String
 
   If Text1(Index) = "" Then
        'Modify By Cheng 2003/02/26
        '­Y¥¼¿é¤J¥N¸¹«h²MªÅ¦WºÙ
'      If Index <> 3 And Index <> 4 And Index <> 5 Then Label1(Index + 1).Caption = ""
' Modify by Toni ¼ç¦b«È¤átext1(16)
        Select Case Index
        'Case 0, 1, 2
        Case 0, 1, 2, 16
            'Modified by Lydia 2022/05/02
            'Label1(Index + 1).Caption = ""
            lblFM2(Index).Caption = ""
        End Select
      Exit Sub
   End If
    'Add By Cheng 2003/12/25
    Select Case Index
    Case 0
        If Me.Text1(0).Text = "X" Then Exit Sub
    Case 1
        If Me.Text1(1).Text = "Y" Then Exit Sub
    
    'add by toni 2008/10/28
    Case 16
       If Me.Text1(16).Text = "R" Then Exit Sub
    'end 2008/10/28
    End Select
    'End
   Select Case Index
      Case 0 '¥Ó½Ð¤H
         'edit by nickc 2007/02/02 ¤£¥Î dll ¤F
         'If objPublicData.GetCustomer(Text1(Index), strTempName) = False Then
         If ClsPDGetCustomer(Text1(Index), strTempName) = False Then
            Cancel = True
            'Modified by Lydia 2022/05/02
            'Label1(Index + 1).Caption = ""
            lblFM2(Index).Caption = ""
         Else
            If strTempName <> "" Then
               'Modifed by Lydia 2022/05/02
               'Label1(Index + 1).Caption = strTempName
               lblFM2(Index).Caption = strTempName
            Else
               'Modified by Lydia 2022/05/02
               'Label1(Index + 1).Caption = ""
               lblFM2(Index).Caption = ""
            End If
         End If
      Case 1 '¥N²z¤H
         'edit by nickc 2007/02/02 ¤£¥Î dll ¤F
         'If objPublicData.GetAgent(Text1(Index), strTempName) = False Then
         If ClsPDGetAgent(Text1(Index), strTempName) = False Then
            Cancel = True
            'Modified by Lydia 2022/05/02
            'Label1(Index + 1).Caption = ""
            lblFM2(Index).Caption = ""
         Else
            If strTempName <> "" Then
               'Modifed by Lydia 2022/05/02
               'Label1(Index + 1).Caption = strTempName
               lblFM2(Index).Caption = strTempName
            Else
               'Modified by Lydia 2022/05/02
               'Label1(Index + 1).Caption = ""
               lblFM2(Index).Caption = ""
            End If
         End If
      Case 2 '¾÷Ãö¥N¸¹
         'edit by nickc 2007/02/05 ¤£¥Î dll ¤F
         'If objLawDll.GetGovName(Text1(Index), strTempName) = False Then
         If ClsPDGetGovName(Text1(Index), strTempName) = False Then
            Cancel = True
            Label1(Index + 1).Caption = ""
         Else
            If strTempName <> "" Then
               Label1(Index + 1).Caption = strTempName
            Else
               Label1(Index + 1).Caption = ""
            End If
         End If
       '2008/11/10 add by toni
    Case 9
      If IsCorrectSysKind(Text1(9).Text) = False Then
         Cancel = True
         MsgBox "¦¹¨t²ÎÃþ§O¤£¦s¦b,½Ð¬d®Ö"
         TextInverse Text1(9)
         Exit Sub
      End If
    'end 2008/11/10
      '¼ç¦b«È¤á
      'add by Toni 2008/10/28
      Case 16
         If ClsPCUGetContact(Text1(16), strTempName) = False Then
            Cancel = True
            'Modified by Lydia 2022/05/02
            'Label1(Index + 1).Caption = ""
            lblFM2(Index).Caption = ""
         Else
            If strTempName <> "" Then
               'Modifed by Lydia 2022/05/02
               'Label1(Index + 1).Caption = strTempName
               lblFM2(Index).Caption = strTempName
            Else
               'Modified by Lydia 2022/05/02
               'Label1(Index + 1).Caption = ""
               lblFM2(Index).Caption = ""
            End If
         End If
         'end 2008/10/28
   End Select
   If Cancel Then TextInverse Text1(Index)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2003/01/30
    '­Y¦Lªí¾÷©Î°¾²¾­È¦³ÅÜ°Ê, «h§ó·s¦C¦L³]©w
    If Me.Combo1.Text <> Me.Combo1.Tag Or Me.Text1(7).Text <> Me.Text1(7).Tag Or Me.Text1(8).Text <> Me.Text1(8).Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, Me.Text1(7).Text, Me.Text1(8).Text, Me.Combo1.Text
    End If
    'Added by Lydia 2017/11/03 ­Y¦Lªí¾÷ÅÜ°Ê, «h§ó·s¦C¦L³]©w
    If Me.Combo2.Text <> Me.Combo2.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
    End If
   'end 2017/11/03
   Set Printer = Printers(SeekPrint)
   Printer.Orientation = SeekPrintL
  
   iStiu = 0  'Added by Lydia 2017/11/03

Set frm083014 = Nothing
End Sub

'Add By Cheng 2003/01/28
'§å¦¸¦C¦L¦a§}±ø
Private Sub PrintCaseBatch(strAL01 As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim i As Integer
Dim St As String
Dim Page As Integer
Dim iPrint As Integer
Dim IntF As Integer
Dim PriType As Integer
Dim j As Integer
Dim Prn As Printer
Dim nRow As Integer
Dim intBgnRow As Integer '°_©l¦C¼Æ
Dim strA1K28 As String
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim blnPrintCC As Boolean '¬O§_¦C¦L°Æ¥»¦¬¨ü¤H¦a§}±ø
Dim stAddr() As String '¦a§}±ø¬ÛÃö¸ê®Æ
Dim dblCnt As Double 'Add By Sindy 2010/10/4
Dim bolBeAsked As Boolean 'Added by Morgan 2012/10/3
   
On Error Resume Next
    
    'Modify By Sindy 2021/3/30 + , AL09, AL10
    StrSQLa = "Select PA75, PA26, PA01, PA02, PA03, PA04, AL06, AL08, Nvl(PA76, Nvl(CU96, Nvl(PA75, PA26))), PA86, PA87, AL09, AL10 From AddressList, Patent, Fagent, Customer Where AL02=PA01 And AL03=PA02 And AL04=PA03 And AL05=PA04 And AL01='" & strAL01 & "' And substr(PA75,1,8)=FA01(+) And substr(PA75,9,1)=FA02(+) And substr(PA26,1,8)=CU01(+) And substr(PA26,9,1)=CU02(+) "
    StrSQLa = StrSQLa & " UNION Select TM44, TM23, TM01, TM02, TM03, TM04, AL06, AL08, Nvl(TM33, Nvl(FA66, Nvl(CU98, Nvl(TM44, TM23)))), TM54, TM55, AL09, AL10 From AddressList, Trademark, Fagent, Customer Where AL02=TM01 And AL03=TM02 And AL04=TM03 And AL05=TM04 And AL01='" & strAL01 & "' And substr(TM44,1,8)=FA01(+) And substr(TM44,9,1)=FA02(+) And substr(TM23,1,8)=CU01(+) And substr(TM23,9,1)=CU02(+) "
    StrSQLa = StrSQLa & " UNION Select LC22, LC11, LC01, LC02, LC03, LC04, AL06, AL08, '', '', '', AL09, AL10 From AddressList, Lawcase Where AL02=LC01 And AL03=LC02 And AL04=LC03 And AL05=LC04 And AL01='" & strAL01 & "' "
    StrSQLa = StrSQLa & " UNION Select '', HC05, HC01, HC02, HC03, HC04, AL06, AL08, '', '', '', AL09, AL10 From AddressList, Hirecase Where AL02=HC01 And AL03=HC02 And AL04=HC03 And AL05=HC04 And AL01='" & strAL01 & "' "
    StrSQLa = StrSQLa & " UNION Select SP26 , SP08 , SP01 , SP02 , SP03 , SP04 , AL06, AL08, '', '', '', AL09, AL10 From AddressList, ServicePractice Where AL02=SP01 And AL03=SP02 And AL04=SP03 And AL05=SP04 And AL01='" & strAL01 & "' "
    StrSQLa = StrSQLa & " Order By AL06 "
    
    If rsA.State <> adStateClosed Then rsA.Close
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    '­Y¦³¦a§}±ø¸ê®Æ
    If rsA.RecordCount > 0 Then
      If MsgBox("·Ç³Æ¦C¦L¦a§}±ø¡A½Ð§ó´«¯È±i!!!", vbExclamation + vbOKCancel) = vbOK Then

RePrint:
      'Added by Morgan 2012/10/3 ­Y¹O®É­Ô­«·sµn¤J»Ý­«§ì¸ê®Æ
      If rsA.State = adStateClosed Then
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      End If
      'end 2012/10/3
      
      dblCnt = 0
      'Modify by Morgan 2006/4/27 ¥Ø«eXP¦Û©w¯È±i»Ý¤â°Ê³]©w¨Ã±N¦Lªí¾÷¹w³]¬°¸Ó¯È±i
      '95
      If pub_OS = "1" Then
         Printer.Height = 2880
         Printer.Width = 10000
      'NT ¶·¥ýµ²§ô¤å¥ó,§_«h¯È±i¤£·|¥Î³ß¦n³]©w
      Else
         'Modify by Morgan 2008/6/30
         'Printer.Orientation = 1
         'Printer.EndDoc
         Printer.PaperSize = PUB_GetPaperSize(2)
         'end 2008/6/30
      End If

            'Add by Morgan 2005/3/2 »Ý­«·s³]©w§_«h­«¦L®É½s¸¹·|Ä~Äò¥[
            iPrint = 1
            '²¾¦Ü²Ä¤@µ§¸ê®Æ
            rsA.MoveFirst
            While Not rsA.EOF
               
                Select Case "" & rsA.Fields(2).Value & rsA.Fields(7).Value
                Case "FCP605", "CFP605", "P605", "FCT102", "CFT102", "T102", "TF102"
                    '­Y¦³¥N²z¤H
                    If Left("" & rsA.Fields(8).Value, 1) = "Y" Then
                        frm083014.opt1(1).Value = True
                        frm083014.Text1(1).Text = "" & rsA.Fields(8).Value
                    '­YµL¥N²z¤H«h¦L¥Ó½Ð¤H
                    ElseIf Left("" & rsA.Fields(8).Value, 1) = "X" Then
                        frm083014.opt1(0).Value = True
                        frm083014.Text1(0).Text = "" & rsA.Fields(8).Value
                    '­YµL¥Ó½Ð¤H
                    Else
                        MsgBox "¥»©Ò®×¸¹¡G" & rsA.Fields(2).Value & "-" & rsA.Fields(3).Value & "-" & rsA.Fields(4).Value & "-" & rsA.Fields(5).Value & " µL¥Ó½Ð¤H¸ê®Æ!!!", vbExclamation + vbOKOnly
                        GoTo NextRecord
                    End If
                Case Else
                    '­Y¦³¥N²z¤H
                    If "" & rsA.Fields(0).Value <> "" Then
                        frm083014.opt1(1).Value = True
                        frm083014.Text1(1).Text = "" & rsA.Fields(0).Value
                    '­YµL¥N²z¤H«h¦L¥Ó½Ð¤H
                    ElseIf "" & rsA.Fields(1).Value <> "" Then
                        frm083014.opt1(0).Value = True
                        frm083014.Text1(0).Text = "" & rsA.Fields(1).Value
                    '­YµL¥Ó½Ð¤H
                    Else
                        MsgBox "¥»©Ò®×¸¹¡G" & rsA.Fields(2).Value & "-" & rsA.Fields(3).Value & "-" & rsA.Fields(4).Value & "-" & rsA.Fields(5).Value & " µL¥Ó½Ð¤H¸ê®Æ!!!", vbExclamation + vbOKOnly
                        GoTo NextRecord
                    End If
                End Select
                '»y¤å
                'Modify by Morgan 2006/6/2
                'frm083014.Text1(4).Text = GetLetterLanguage(rsA.Fields(2).Value, rsA.Fields(3).Value, rsA.Fields(4).Value, rsA.Fields(5).Value)
                frm083014.Text1(4).Text = PUB_GetLanguage(rsA.Fields(2).Value, rsA.Fields(3).Value, rsA.Fields(4).Value, rsA.Fields(5).Value, "" & rsA.Fields(7).Value)
                
                '2005/7/4 ADD BY SONIA ¨ú®ø¤é¤å¦a§}±ø§ï¦L­^¤å¦a§}±ø
                'Modified by Lydia 2025/10/08 ¦Ë¾ä¦a§}±ø¡G«ü©w¼ç¦b«È¤á¥i¥H§ì¤é¤å
                'If frm083014.Text1(4).Text = "3" Then
                If frm083014.Text1(4).Text And p_SpecLan <> "3" Then
                   frm083014.Text1(4).Text = "2"
                End If
                '2005/7/4 END
                '¥»©Ò®×¸¹
                Me.Text1(6).Text = "" & rsA.Fields(2).Value & rsA.Fields(3).Value & rsA.Fields(4).Value & rsA.Fields(5).Value
                m_CaseNo = ""
                '­Y¬°¤º±M¤H­û   '2007/9/12 ¥[¤J¤º°Ó¤H­û(¦] frm020311)
                If GetStaffDepartment(strUserNum) = "P12" Or GetStaffDepartment(strUserNum) = "P22" Then
                    '­Y¥¼¥Ñ¥~³¡µ{¦¡¶Ç¤J¥»©Ò®×¸¹, ¥B¦³¿é¤J¥»©Ò®×¸¹
                    'Modify By Cheng 2003/04/30
'                    If m_CaseNo = "" And (Me.Text1(9).Text <> "" And Me.Text1(10).Text <> "") Then
                    If m_CaseNo = "" Then
                        'Modify By Cheng 2003/04/30
'                        Me.SetCaseNo Me.Text1(9).Text & "-" & Me.Text1(10).Text & "-" & Left(Me.Text1(11).Text & "0", 1) & "-" & Left(Me.Text1(12).Text & "00", 2)
                        Me.SetCaseNo rsA.Fields(2).Value & "-" & rsA.Fields(3).Value & "-" & Left(rsA.Fields(4).Value & "0", 1) & "-" & Left(rsA.Fields(5).Value & "00", 2)
                    End If
                End If
'*************************************************
                '¹w³]¤£¦C¦L°Æ¥»¦¬¨ü¤H
                blnPrintCC = False
                '¹w³]µLÁpµ¸¤H
                tmp083 = ""
                tmp083_1 = ""
PrintCC:
                '³]©w°¾²¾­È
                m_dbl_LeftMargin = CDbl(Me.Text1(7).Text) * 567
                m_dbl_TopMargin = CDbl(Me.Text1(8).Text) * 567
               
                'Add By Sindy 2021/3/30
                '«ü©w¦¬¥ó¤H
                If "" & rsA.Fields("AL09") <> "" Then
                   If Left(rsA.Fields("AL09"), 1) = "X" Then
                      frm083014.opt1(0).Value = True
                      frm083014.Text1(0).Text = ChangeCustomerL(rsA.Fields("AL09"))
                   ElseIf Left(rsA.Fields("AL09"), 1) = "Y" Then
                      frm083014.opt1(1).Value = True
                      frm083014.Text1(1).Text = ChangeCustomerL(rsA.Fields("AL09"))
                   End If
                End If
                '2021/3/30 END
               
                PriType = Val(Text1(4).Text)
                '¥Ó½Ð¤H
                If opt1(0).Value = True Then
                    St = ChgCustomer(Text1(0).Text)
                    If Text1(5) <> "Y" Then
                        St = St & " AND (CU32<>'N' OR CU32 IS NULL)"
                    End If
                    '»y¤å
                    Select Case PriType
                        Case 1  '7
                            strExc(0) = "SELECT Nvl(CU80,CU30), Decode(CU80, Null, DECODE(CU31,NULL,SUBSTR(CU23,1,20),SUBSTR(CU31,1,20)), ''), Decode(CU80, Null, DECODE(CU31,NULL,SUBSTR(CU23,21,15),SUBSTR(CU31,21,15)), '')," & _
                                "SUBSTR(decode(cu104,null,CU04,cu104),1,20),SUBSTR(decode(cu104,null,CU04,cu104),21,20),CU08,CU01||CU02 FROM CUSTOMER " & _
                                "WHERE " & St
                            Page = 1
                        Case 2  '11
                            '­Y¦³¥»©Ò®×¸¹
                            If Trim(Text1(6).Text) <> "" Then
                                '­Y«D¦C¦L°Æ¥»¦¬¨ü¤H
                                If blnPrintCC = False Then
                                    'Modify by Morgan 2010/12/14 §ï©I¥s¦@¥Î¨ç¼Æ³]©w
                                    SetCaseContact Text1(6).Text, "" & rsA.Fields(2), "" & rsA.Fields(7), 0, 2, tmp083, tmp083_1
                                    'Add By Sindy 2021/3/30
                                    'Ápµ¸¤H¦WºÙ1
                                    If "" & rsA.Fields("AL10") <> "" Then
                                       tmp083 = rsA.Fields("AL10")
                                       tmp083_1 = ""
                                    End If
                                    '2021/3/30 END
'                                    '­Y¬°±M§Q¦~¶O
'                                    If ("" & rsA.Fields(2).Value = "FCP" Or "" & rsA.Fields(2).Value = "CFP" Or "" & rsA.Fields(2).Value = "P") And ("" & rsA.Fields(7).Value = "605") Then
'                                        '­Y¦³¦~¶O¥N²z¤H, Ápµ¸¤Hª½±µ§ì¬Û¦PÀÉ®×ªº¦~¶OÁpµ¸¤H(¤£½×¬O§_¦³­È)
'                                        StrSQL083 = "select Decode(PA76, Null, Nvl(PA52, Nvl(FA08, CU59)), PA135) from patent,customer,fagent " & _
'                                                    " where " & ChgPatent(Replace(Me.Text1(6).Text, "-", "")) & " And substr(pa26,1,8) = cu01(+) " & _
'                                                    " and nvl(substr(pa26,9,1),'0')=cu02(+) and substr(pa75,1,8)=fa01(+) and nvl(substr(pa75,9,1),'0')=fa02(+) "
'                                    Else
'                                        StrSQL083 = "select nvl(pa52, Nvl(FA08, CU59)) from patent,customer,fagent " & _
'                                                    " where " & ChgPatent(Replace(Me.Text1(6).Text, "-", "")) & " And substr(pa26,1,8) = cu01(+) " & _
'                                                    " and nvl(substr(pa26,9,1),'0')=cu02(+) and substr(pa75,1,8)=fa01(+) and nvl(substr(pa75,9,1),'0')=fa02(+) "
'                                    End If
'                                    '­Y¬°°Ó¼Ð©µ®i
'                                    If ("" & rsA.Fields(2).Value = "FCT" Or "" & rsA.Fields(2).Value = "CFT" Or "" & rsA.Fields(2).Value = "T" Or "" & rsA.Fields(2).Value = "TF") And ("" & rsA.Fields(7).Value = "102") Then
'                                        '­Y¦³©µ®i¥N²z¤H, Ápµ¸¤Hª½±µ§ì¬Û¦PÀÉ®×ªº¦~¶OÁpµ¸¤H(¤£½×¬O§_¦³­È)
'                                        StrSQL083 = StrSQL083 & " union select Decode(TM33, Null, Nvl(TM39, Nvl(FA08, CU59)), TM71) from Trademark, customer, fagent " & _
'                                                    " where " & ChgTradeMark(Replace(Me.Text1(6).Text, "-", "")) & " And substr(tm23,1,8) = cu01(+) " & _
'                                                    " and nvl(substr(tm23,9,1),'0')=cu02(+) and substr(tm44,1,8)=fa01(+) and nvl(substr(tm44,9,1),'0')=fa02(+) "
'                                    Else
'                                        StrSQL083 = StrSQL083 & " union select nvl(tm39, Nvl(FA08, CU59)) from Trademark, customer, fagent " & _
'                                                    " where " & ChgTradeMark(Replace(Me.Text1(6).Text, "-", "")) & " And substr(tm23,1,8) = cu01(+) " & _
'                                                    " and nvl(substr(tm23,9,1),'0')=cu02(+) and substr(tm44,1,8)=fa01(+) and nvl(substr(tm44,9,1),'0')=fa02(+) "
'                                    End If
'                                    Set Rs083 = New ADODB.Recordset
'                                    Rs083.CursorLocation = adUseClient
'                                    Rs083.Open StrSQL083, cnnConnection, adOpenStatic, adLockReadOnly
'                                    If Rs083.RecordCount > 0 Then
'                                        tmp083 = CheckStr(Rs083.Fields(0).Value)
'                                    Else
'                                        tmp083 = ""
'                                    End If
'                                    'Add By Cheng 2003/12/23
'                                    'Ápµ¸¤H2­^
'                                    '­Y¬°±M§Q¦~¶O
'                                    If ("" & rsA.Fields(2).Value = "FCP" Or "" & rsA.Fields(2).Value = "CFP" Or "" & rsA.Fields(2).Value = "P") And ("" & rsA.Fields(7).Value = "605") Then
'                                        '­Y¦³¦~¶O¥N²z¤H, Ápµ¸¤Hª½±µ§ì¬Û¦PÀÉ®×ªº¦~¶OÁpµ¸¤H(¤£½×¬O§_¦³­È)
'                                        StrSQL083 = "select Decode(PA76, Null, Decode(pa52, Null, Decode(FA08, Null, Decode(CU59, Null, '', CU62), FA53), PA55), '') from patent,customer,fagent " & _
'                                                    " where " & ChgPatent(Replace(Me.Text1(6).Text, "-", "")) & " And substr(pa26,1,8) = cu01(+) " & _
'                                                    " and nvl(substr(pa26,9,1),'0')=cu02(+) and substr(pa75,1,8)=fa01(+) and nvl(substr(pa75,9,1),'0')=fa02(+) "
'                                    Else
'                                        StrSQL083 = "select Decode(pa52, Null, Decode(FA08, Null, Decode(CU59, Null, '', CU62), FA53), PA55) from patent,customer,fagent " & _
'                                                    " where " & ChgPatent(Replace(Me.Text1(6).Text, "-", "")) & " And substr(pa26,1,8) = cu01(+) " & _
'                                                    " and nvl(substr(pa26,9,1),'0')=cu02(+) and substr(pa75,1,8)=fa01(+) and nvl(substr(pa75,9,1),'0')=fa02(+) "
'                                    End If
'                                    '­Y¬°°Ó¼Ð©µ®i
'                                    If ("" & rsA.Fields(2).Value = "FCT" Or "" & rsA.Fields(2).Value = "CFT" Or "" & rsA.Fields(2).Value = "T" Or "" & rsA.Fields(2).Value = "TF") And ("" & rsA.Fields(7).Value = "102") Then
'                                        '­Y¦³©µ®i¥N²z¤H, Ápµ¸¤Hª½±µ§ì¬Û¦PÀÉ®×ªº¦~¶OÁpµ¸¤H(¤£½×¬O§_¦³­È)
'                                        StrSQL083 = StrSQL083 & " union select Decode(TM33, Null, Decode(tm39, Null, Decode(FA08, Null, Decode(CU59, Null, '', CU62), FA53), TM42), '') from Trademark, customer, fagent " & _
'                                                    " where " & ChgTradeMark(Replace(Me.Text1(6).Text, "-", "")) & " And substr(tm23,1,8) = cu01(+) " & _
'                                                    " and nvl(substr(tm23,9,1),'0')=cu02(+) and substr(tm44,1,8)=fa01(+) and nvl(substr(tm44,9,1),'0')=fa02(+) "
'                                    Else
'                                        StrSQL083 = StrSQL083 & " union select Decode(tm39, Null, Decode(FA08, Null, Decode(CU59, Null, '', CU62), FA53), TM42) from Trademark, customer, fagent " & _
'                                                    " where " & ChgTradeMark(Replace(Me.Text1(6).Text, "-", "")) & " And substr(tm23,1,8) = cu01(+) " & _
'                                                    " and nvl(substr(tm23,9,1),'0')=cu02(+) and substr(tm44,1,8)=fa01(+) and nvl(substr(tm44,9,1),'0')=fa02(+) "
'                                    End If
'                                    Set Rs083 = New ADODB.Recordset
'                                    Rs083.CursorLocation = adUseClient
'                                    Rs083.Open StrSQL083, cnnConnection, adOpenStatic, adLockReadOnly
'                                    If Rs083.RecordCount > 0 Then
'                                        tmp083_1 = CheckStr(Rs083.Fields(0).Value)
'                                    Else
'                                        tmp083_1 = ""
'                                    End If
'                                    'End
                                    'end 2010/12/14
                                    
                                End If
                                'Modify By Cheng 2003/12/23
                                '¥[Ápµ¸¤H2­^
'                                strExc(0) = "SELECT '" & ChgSQL(tmp083) & "',CU05,CU88,CU89,CU90, Decode(CU80, Null, DECODE(CU65,'',CU24,CU65), CU80)," & _
'                                   "Decode(CU80, Null, DECODE(CU65,'',CU25,CU66), ''), Decode(CU80, Null, DECODE(CU65,'',CU26,CU67), '')," & _
'                                   "Decode(CU80, Null, DECODE(CU65,'',CU27,CU68), ''), Decode(CU80, Null, DECODE(CU65,'',CU28,CU69), '')," & _
'                                   "Decode(CU80, Null, DECODE(CU65,'',CU102,''), '') FROM CUSTOMER WHERE " & St
                                'Modify By Sindy 2016/1/27 ¬ü°ê¥u§ìÁpµ¸¤H1
                                strExc(0) = "SELECT '" & ChgSQL(tmp083) & "', '" & IIf(Left(ChangeCustomerL(GetPrjNationNumber1(Text1(1))), 3) = "101", "", ChgSQL(tmp083_1)) & "',CU05,CU88,CU89,CU90, Decode(CU80, Null, DECODE(CU65,'',CU24,CU65), CU80)," & _
                                   "Decode(CU80, Null, DECODE(CU65,'',CU25,CU66), ''), Decode(CU80, Null, DECODE(CU65,'',CU26,CU67), '')," & _
                                   "Decode(CU80, Null, DECODE(CU65,'',CU27,CU68), ''), Decode(CU80, Null, DECODE(CU65,'',CU28,CU69), '')," & _
                                   "Decode(CU80, Null, DECODE(CU65,'',CU102,''), '') FROM CUSTOMER WHERE " & St
                                'End
                            '­Y¥¼¿é¤J¥»©Ò®×¸¹
                            Else
                                'Modify By Cheng 2003/12/23
                                '¥[Ápµ¸¤H2­^
'                                strExc(0) = "SELECT cu59,CU05,CU88,CU89,CU90, Decode(CU80, Null, DECODE(CU65,'',CU24,CU65), CU80)," & _
'                                   "Decode(CU80, Null, DECODE(CU65,'',CU25,CU66), ''), Decode(CU80, Null, DECODE(CU65,'',CU26,CU67), '')," & _
'                                   "Decode(CU80, Null, DECODE(CU65,'',CU27,CU68), ''), Decode(CU80, Null, DECODE(CU65,'',CU28,CU69), '')," & _
'                                   "Decode(CU80, Null, DECODE(CU65,'',CU102,''), '') FROM CUSTOMER WHERE " & St
                                strExc(0) = "SELECT cu59, CU62, CU05,CU88,CU89,CU90, Decode(CU80, Null, DECODE(CU65,'',CU24,CU65), CU80)," & _
                                   "Decode(CU80, Null, DECODE(CU65,'',CU25,CU66), ''), Decode(CU80, Null, DECODE(CU65,'',CU26,CU67), '')," & _
                                   "Decode(CU80, Null, DECODE(CU65,'',CU27,CU68), ''), Decode(CU80, Null, DECODE(CU65,'',CU28,CU69), '')," & _
                                   "Decode(CU80, Null, DECODE(CU65,'',CU102,''), '') FROM CUSTOMER WHERE " & St
                                'End
                            End If
                            '***************************************************   end
                            Page = 2
                     Case 3  '6
                            '­Y¦³¥»©Ò®×¸¹
                            If Trim(Text1(6).Text) <> "" Then
                                '­Y«D¦C¦L°Æ¥»¦¬¨ü¤H
                                If blnPrintCC = False Then
                                    'Modify by Morgan 2010/12/14 §ï©I¥s¦@¥Î¨ç¼Æ³]©w
                                    SetCaseContact Text1(6).Text, "" & rsA.Fields(2), "" & rsA.Fields(7), 0, 3, tmp083, tmp083_1
'                                    '­Y¬°±M§Q¦~¶O
'                                    If ("" & rsA.Fields(2).Value = "FCP" Or "" & rsA.Fields(2).Value = "CFP" Or "" & rsA.Fields(2).Value = "P") And ("" & rsA.Fields(7).Value = "605") Then
'                                        '­Y¦³¦~¶O¥N²z¤H, Ápµ¸¤Hª½±µ§ì¬Û¦PÀÉ®×ªº¦~¶OÁpµ¸¤H(¤£½×¬O§_¦³­È)
'                                        StrSQL083 = "select Decode(PA76, Null, Nvl(PA53, Nvl(FA09, CU60)), PA135) from patent,customer,fagent " & _
'                                                    " where " & ChgPatent(Replace(Me.Text1(6).Text, "-", "")) & " And substr(pa26,1,8) = cu01(+) " & _
'                                                    " and nvl(substr(pa26,9,1),'0')=cu02(+) and substr(pa75,1,8)=fa01(+) and nvl(substr(pa75,9,1),'0')=fa02(+) "
'                                    Else
'                                        StrSQL083 = "select nvl(pa53, Nvl(FA09, CU60)) from patent,customer,fagent " & _
'                                                    " where " & ChgPatent(Replace(Me.Text1(6).Text, "-", "")) & " And substr(pa26,1,8) = cu01(+) " & _
'                                                    " and nvl(substr(pa26,9,1),'0')=cu02(+) and substr(pa75,1,8)=fa01(+) and nvl(substr(pa75,9,1),'0')=fa02(+) "
'                                    End If
'                                    '­Y¬°°Ó¼Ð©µ®i
'                                    If ("" & rsA.Fields(2).Value = "FCT" Or "" & rsA.Fields(2).Value = "CFT" Or "" & rsA.Fields(2).Value = "T" Or "" & rsA.Fields(2).Value = "TF") And ("" & rsA.Fields(7).Value = "102") Then
'                                        '­Y¦³©µ®i¥N²z¤H, Ápµ¸¤Hª½±µ§ì¬Û¦PÀÉ®×ªº¦~¶OÁpµ¸¤H(¤£½×¬O§_¦³­È)
'                                        StrSQL083 = StrSQL083 & " union select Decode(TM33, Null, Nvl(TM40, Nvl(FA09, CU60)), TM71) from Trademark, customer, fagent " & _
'                                                    " where " & ChgTradeMark(Replace(Me.Text1(6).Text, "-", "")) & " And substr(tm23,1,8) = cu01(+) " & _
'                                                    " and nvl(substr(tm23,9,1),'0')=cu02(+) and substr(tm44,1,8)=fa01(+) and nvl(substr(tm44,9,1),'0')=fa02(+) "
'                                    Else
'                                        StrSQL083 = StrSQL083 & " union select nvl(tm40, Nvl(FA09, CU60)) from Trademark, customer, fagent " & _
'                                                    " where " & ChgTradeMark(Replace(Me.Text1(6).Text, "-", "")) & " And substr(tm23,1,8) = cu01(+) " & _
'                                                    " and nvl(substr(tm23,9,1),'0')=cu02(+) and substr(tm44,1,8)=fa01(+) and nvl(substr(tm44,9,1),'0')=fa02(+) "
'                                    End If
'                                    Set Rs083 = New ADODB.Recordset
'                                    Rs083.CursorLocation = adUseClient
'                                    Rs083.Open StrSQL083, cnnConnection, adOpenStatic, adLockReadOnly
'                                    If Rs083.RecordCount > 0 Then
'                                        tmp083 = CheckStr(Rs083.Fields(0).Value)
'                                    Else
'                                        tmp083 = ""
'                                    End If
'                                    'Add By Cheng 2003/12/23
'                                    'Ápµ¸¤H2­^
'                                    '­Y¬°±M§Q¦~¶O
'                                    If ("" & rsA.Fields(2).Value = "FCP" Or "" & rsA.Fields(2).Value = "CFP" Or "" & rsA.Fields(2).Value = "P") And ("" & rsA.Fields(7).Value = "605") Then
'                                        '­Y¦³¦~¶O¥N²z¤H, Ápµ¸¤Hª½±µ§ì¬Û¦PÀÉ®×ªº¦~¶OÁpµ¸¤H(¤£½×¬O§_¦³­È)
'                                        StrSQL083 = "select Decode(PA76, Null, Decode(pa53, Null, Decode(FA09, Null, Decode(CU60, Null, '', CU63), FA54), PA56), '') from patent,customer,fagent " & _
'                                                    " where " & ChgPatent(Replace(Me.Text1(6).Text, "-", "")) & " And substr(pa26,1,8) = cu01(+) " & _
'                                                    " and nvl(substr(pa26,9,1),'0')=cu02(+) and substr(pa75,1,8)=fa01(+) and nvl(substr(pa75,9,1),'0')=fa02(+) "
'                                    Else
'                                        StrSQL083 = "select Decode(pa53, Null, Decode(FA09, Null, Decode(CU60, Null, '', CU63), FA54), PA56) from patent,customer,fagent " & _
'                                                    " where " & ChgPatent(Replace(Me.Text1(6).Text, "-", "")) & " And substr(pa26,1,8) = cu01(+) " & _
'                                                    " and nvl(substr(pa26,9,1),'0')=cu02(+) and substr(pa75,1,8)=fa01(+) and nvl(substr(pa75,9,1),'0')=fa02(+) "
'                                    End If
'                                    '­Y¬°°Ó¼Ð©µ®i
'                                    If ("" & rsA.Fields(2).Value = "FCT" Or "" & rsA.Fields(2).Value = "CFT" Or "" & rsA.Fields(2).Value = "T" Or "" & rsA.Fields(2).Value = "TF") And ("" & rsA.Fields(7).Value = "102") Then
'                                        '­Y¦³©µ®i¥N²z¤H, Ápµ¸¤Hª½±µ§ì¬Û¦PÀÉ®×ªº¦~¶OÁpµ¸¤H(¤£½×¬O§_¦³­È)
'                                        StrSQL083 = StrSQL083 & " union select Decode(TM33, Null, Decode(tm40, Null, Decode(FA09, Null, Decode(CU60, Null, '', CU63), FA54), TM43), '') from Trademark, customer, fagent " & _
'                                                    " where " & ChgTradeMark(Replace(Me.Text1(6).Text, "-", "")) & " And substr(tm23,1,8) = cu01(+) " & _
'                                                    " and nvl(substr(tm23,9,1),'0')=cu02(+) and substr(tm44,1,8)=fa01(+) and nvl(substr(tm44,9,1),'0')=fa02(+) "
'                                    Else
'                                        StrSQL083 = StrSQL083 & " union select Decode(tm40, Null, Decode(FA09, Null, Decode(CU60, Null, '', CU63), FA54), TM43) from Trademark, customer, fagent " & _
'                                                    " where " & ChgTradeMark(Replace(Me.Text1(6).Text, "-", "")) & " And substr(tm23,1,8) = cu01(+) " & _
'                                                    " and nvl(substr(tm23,9,1),'0')=cu02(+) and substr(tm44,1,8)=fa01(+) and nvl(substr(tm44,9,1),'0')=fa02(+) "
'                                    End If
'                                    Set Rs083 = New ADODB.Recordset
'                                    Rs083.CursorLocation = adUseClient
'                                    Rs083.Open StrSQL083, cnnConnection, adOpenStatic, adLockReadOnly
'                                    If Rs083.RecordCount > 0 Then
'                                        tmp083_1 = CheckStr(Rs083.Fields(0).Value)
'                                    Else
'                                        tmp083_1 = ""
'                                    End If
'                                    'End
                                    'end 2010/12/14
                                End If
'                                strExc(0) = "SELECT '" & ChgSQL(tmp083) & "', '" & ChgSQL(tmp083_1) & "',CU05,CU88,CU89,CU90, Decode(CU80, Null, DECODE(CU65,'',CU24,CU65), CU80)," & _
'                                   "Decode(CU80, Null, DECODE(CU65,'',CU25,CU66), ''), Decode(CU80, Null, DECODE(CU65,'',CU26,CU67), '')," & _
'                                   "Decode(CU80, Null, DECODE(CU65,'',CU27,CU68), ''), Decode(CU80, Null, DECODE(CU65,'',CU28,CU69), '')," & _
'                                   "Decode(CU80, Null, DECODE(CU65,'',CU102,''), '') FROM CUSTOMER WHERE " & St
                                strExc(0) = "SELECT Decode(CU80, Null, SUBSTR(CU29,1,20), CU80), Decode(CU80, Null, SUBSTR(CU29,21,15), '')," & _
                                   "SUBSTR(CU06,1,20),SUBSTR(CU06,21,20),'" & ChgSQL(tmp083) & "', '" & ChgSQL(tmp083_1) & "',CU01||CU02 FROM CUSTOMER " & _
                                   "WHERE " & St
                                'End
                            '­Y¥¼¿é¤J¥»©Ò®×¸¹
                            Else
                                strExc(0) = "SELECT Decode(CU80, Null, SUBSTR(CU29,1,20), CU80), Decode(CU80, Null, SUBSTR(CU29,21,15), '')," & _
                                   "SUBSTR(CU06,1,20),SUBSTR(CU06,21,20),CU60,CU63,CU01||CU02 FROM CUSTOMER " & _
                                   "WHERE " & St
                            End If
'                            Page = 3
                            Page = 1
                  End Select
               '¥N²z¤H
               ElseIf opt1(1).Value = True Then
                  St = ChgFagent(Text1(1).Text)
                  If Text1(5) <> "Y" Then
                     St = St & " AND (FA24<>'N' OR FA24 IS NULL)"
                  End If
                  Select Case PriType
                     Case 1  '5
                        strExc(0) = "SELECT SUBSTR(FA17,1,20),SUBSTR(FA17,21,15)," & _
                           "SUBSTR(FA04,1,20),SUBSTR(FA04,21,20),FA01||FA02 FROM FAGENT " & _
                           "WHERE " & St
                        Page = 4
                     Case 2  '11
                        '­Y¦³¿é¥»©Ò®×¸¹
                        If Trim(Text1(6).Text) <> "" Then
                            '­Y«D¦C¦L°Æ¥»¦¬¨ü¤H
                            If blnPrintCC = False Then
                                'Modify by Morgan 2010/12/14 §ï©I¥s¦@¥Î¨ç¼Æ³]©w
                                 SetCaseContact Text1(6).Text, "" & rsA.Fields(2), "" & rsA.Fields(7), 1, 2, tmp083, tmp083_1
                                 'Add By Sindy 2021/3/30
                                 'Ápµ¸¤H¦WºÙ1
                                 If "" & rsA.Fields("AL10") <> "" Then
                                    tmp083 = rsA.Fields("AL10")
                                    tmp083_1 = ""
                                 End If
                                 '2021/3/30 END
'                                '­Y¬°±M§Q¦~¶O
'                                If ("" & rsA.Fields(2).Value = "FCP" Or "" & rsA.Fields(2).Value = "CFP" Or "" & rsA.Fields(2).Value = "P") And ("" & rsA.Fields(7).Value = "605") Then
'                                    StrSQL083 = "select Decode(PA76, Null, Nvl(PA52, Nvl(FA08, CU59)), PA135) from patent,customer,fagent " & _
'                                                " where " & ChgPatent(Replace(Text1(6).Text, "-", "")) & " and substr(pa26,1,8) = cu01(+) " & _
'                                                " and nvl(substr(pa26,9,1),'0')=cu02(+) and substr(pa75,1,8)=fa01(+) and nvl(substr(pa75,9,1),'0')=fa02(+) "
'                                Else
'                                    StrSQL083 = "select nvl(pa52, Nvl(FA08, CU59)) from patent,customer,fagent " & _
'                                                " where " & ChgPatent(Replace(Text1(6).Text, "-", "")) & " and substr(pa26,1,8) = cu01(+) " & _
'                                                " and nvl(substr(pa26,9,1),'0')=cu02(+) and substr(pa75,1,8)=fa01(+) and nvl(substr(pa75,9,1),'0')=fa02(+) "
'                                End If
'                                '­Y¬°°Ó¼Ð©µ®i
'                                If ("" & rsA.Fields(2).Value = "FCT" Or "" & rsA.Fields(2).Value = "CFT" Or "" & rsA.Fields(2).Value = "T" Or "" & rsA.Fields(2).Value = "TF") And ("" & rsA.Fields(7).Value = "102") Then
'                                    StrSQL083 = StrSQL083 & " union select Decode(TM33, Null, Nvl(TM39, Nvl(FA08, CU59)), TM71) from Trademark, customer, fagent " & _
'                                                " where " & ChgTradeMark(Replace(Text1(6).Text, "-", "")) & " and substr(tm23,1,8) = cu01(+) " & _
'                                                " and nvl(substr(tm23,9,1),'0')=cu02(+) and substr(tm44,1,8)=fa01(+) and nvl(substr(tm44,9,1),'0')=fa02(+) "
'                                Else
'                                    StrSQL083 = StrSQL083 & " union select nvl(tm39, Nvl(FA08, CU59)) from Trademark, customer, fagent " & _
'                                                " where " & ChgTradeMark(Replace(Text1(6).Text, "-", "")) & " and substr(tm23,1,8) = cu01(+) " & _
'                                                " and nvl(substr(tm23,9,1),'0')=cu02(+) and substr(tm44,1,8)=fa01(+) and nvl(substr(tm44,9,1),'0')=fa02(+) "
'                                End If
'                                Set Rs083 = New ADODB.Recordset
'                                Rs083.CursorLocation = adUseClient
'                                Rs083.Open StrSQL083, cnnConnection, adOpenStatic, adLockReadOnly
'                                If Rs083.RecordCount > 0 Then
'                                    tmp083 = CheckStr(Rs083.Fields(0).Value)
'                                Else
'                                    tmp083 = ""
'                                End If
'                                'Add By Cheng 2003/12/23
'                                'Ápµ¸¤H2­^
'                                '­Y¬°±M§Q¦~¶O
'                                If ("" & rsA.Fields(2).Value = "FCP" Or "" & rsA.Fields(2).Value = "CFP" Or "" & rsA.Fields(2).Value = "P") And ("" & rsA.Fields(7).Value = "605") Then
'                                    StrSQL083 = "select Decode(PA76, Null, Decode(pa52, Null, Decode(FA08, Null, Decode(CU59, Null, '', CU62), FA53), PA55), '') from patent,customer,fagent " & _
'                                                " where " & ChgPatent(Replace(Text1(6).Text, "-", "")) & " and substr(pa26,1,8) = cu01(+) " & _
'                                                " and nvl(substr(pa26,9,1),'0')=cu02(+) and substr(pa75,1,8)=fa01(+) and nvl(substr(pa75,9,1),'0')=fa02(+) "
'                                Else
'                                    StrSQL083 = "select Decode(pa52, Null, Decode(FA08, Null, Decode(CU59, Null, '', CU62), FA53), PA55) from patent,customer,fagent " & _
'                                                " where " & ChgPatent(Replace(Text1(6).Text, "-", "")) & " and substr(pa26,1,8) = cu01(+) " & _
'                                                " and nvl(substr(pa26,9,1),'0')=cu02(+) and substr(pa75,1,8)=fa01(+) and nvl(substr(pa75,9,1),'0')=fa02(+) "
'                                End If
'                                '­Y¬°°Ó¼Ð©µ®i
'                                If ("" & rsA.Fields(2).Value = "FCT" Or "" & rsA.Fields(2).Value = "CFT" Or "" & rsA.Fields(2).Value = "T" Or "" & rsA.Fields(2).Value = "TF") And ("" & rsA.Fields(7).Value = "102") Then
'                                    StrSQL083 = StrSQL083 & " union select Decode(TM33, Null, Decode(tm39, Null, Decode(FA08, Null, Decode(CU59, Null, '', CU62), FA53), TM42), '') from Trademark, customer, fagent " & _
'                                                " where " & ChgTradeMark(Replace(Text1(6).Text, "-", "")) & " and substr(tm23,1,8) = cu01(+) " & _
'                                                " and nvl(substr(tm23,9,1),'0')=cu02(+) and substr(tm44,1,8)=fa01(+) and nvl(substr(tm44,9,1),'0')=fa02(+) "
'                                Else
'                                    StrSQL083 = StrSQL083 & " union select Decode(tm39, Null, Decode(FA08, Null, Decode(CU59, Null, '', CU62), FA53), TM42) from Trademark, customer, fagent " & _
'                                                " where " & ChgTradeMark(Replace(Text1(6).Text, "-", "")) & " and substr(tm23,1,8) = cu01(+) " & _
'                                                " and nvl(substr(tm23,9,1),'0')=cu02(+) and substr(tm44,1,8)=fa01(+) and nvl(substr(tm44,9,1),'0')=fa02(+) "
'                                End If
'                                Set Rs083 = New ADODB.Recordset
'                                Rs083.CursorLocation = adUseClient
'                                Rs083.Open StrSQL083, cnnConnection, adOpenStatic, adLockReadOnly
'                                If Rs083.RecordCount > 0 Then
'                                    tmp083_1 = CheckStr(Rs083.Fields(0).Value)
'                                Else
'                                    tmp083_1 = ""
'                                End If
'                                'End
                                 'end 2010/12/14
                            End If
                            'Modify By Cheng 2003/12/23
                            '¥[Ápµ¸¤H2­^
'                            strExc(0) = "SELECT '" & ChgSQL(tmp083) & "',FA05,FA63,FA64,FA65," & _
'                               " Decode(FA32,Null,FA18,FA32), Decode(FA32,Null,FA19,FA33), Decode(FA32,Null,FA20,FA34), Decode(FA32,Null,FA21,FA35), Decode(FA32,Null,FA22,FA36), FA70 FROM FAGENT,NATION WHERE " & St & " AND FA55=NA01(+)"
                            'Modify By Sindy 2016/1/27 ¬ü°ê¥u§ìÁpµ¸¤H1
                            strExc(0) = "SELECT '" & ChgSQL(tmp083) & "','" & IIf(Left(ChangeCustomerL(GetPrjNationNumber(Text1(1))), 3) = "101", "", ChgSQL(tmp083_1)) & "', FA05,FA63,FA64,FA65," & _
                               " Decode(FA32,Null,FA18,FA32), Decode(FA32,Null,FA19,FA33), Decode(FA32,Null,FA20,FA34), Decode(FA32,Null,FA21,FA35), Decode(FA32,Null,FA22,FA36), Decode(FA32,Null,FA70) FROM FAGENT,NATION WHERE " & St & " AND FA55=NA01(+)"
                            'End
                        '­Y¥¼¿é¥»©Ò®×¸¹
                        Else
                            'Modify By Cheng 2003/12/23
                            '¥[Ápµ¸¤H2­^
'                            strExc(0) = "SELECT fa08,FA05,FA63,FA64,FA65," & _
'                               " Decode(FA32,Null,FA18,FA32), Decode(FA32,Null,FA19,FA33), Decode(FA32,Null,FA20,FA34), Decode(FA32,Null,FA21,FA35), Decode(FA32,Null,FA22,FA36), FA70 FROM FAGENT,NATION WHERE " & St & " AND FA55=NA01(+)"
                            strExc(0) = "SELECT fa08, FA53, FA05,FA63,FA64,FA65," & _
                               " Decode(FA32,Null,FA18,FA32), Decode(FA32,Null,FA19,FA33), Decode(FA32,Null,FA20,FA34), Decode(FA32,Null,FA21,FA35), Decode(FA32,Null,FA22,FA36), Decode(FA32,Null,FA70) FROM FAGENT,NATION WHERE " & St & " AND FA55=NA01(+)"
                            'End
                        End If
                        '****************************** end
                        Page = 2
                     Case 3  '5
                        '­Y¦³¿é¥»©Ò®×¸¹
                        If Trim(Text1(6).Text) <> "" Then
                            '­Y«D¦C¦L°Æ¥»¦¬¨ü¤H
                            If blnPrintCC = False Then
                                'Modify by Morgan 2010/12/14 §ï©I¥s¦@¥Î¨ç¼Æ³]©w
                                 SetCaseContact Text1(6).Text, "" & rsA.Fields(2), "" & rsA.Fields(7), 1, 3, tmp083, tmp083_1
'                                '­Y¬°±M§Q¦~¶O
'                                If ("" & rsA.Fields(2).Value = "FCP" Or "" & rsA.Fields(2).Value = "CFP" Or "" & rsA.Fields(2).Value = "P") And ("" & rsA.Fields(7).Value = "605") Then
'                                    StrSQL083 = "select Decode(PA76, Null, Nvl(PA53, Nvl(FA09, CU60)), PA135) from patent,customer,fagent " & _
'                                                " where " & ChgPatent(Replace(Text1(6).Text, "-", "")) & " and substr(pa26,1,8) = cu01(+) " & _
'                                                " and nvl(substr(pa26,9,1),'0')=cu02(+) and substr(pa75,1,8)=fa01(+) and nvl(substr(pa75,9,1),'0')=fa02(+) "
'                                Else
'                                    StrSQL083 = "select nvl(pa53, Nvl(FA09, CU60)) from patent,customer,fagent " & _
'                                                " where " & ChgPatent(Replace(Text1(6).Text, "-", "")) & " and substr(pa26,1,8) = cu01(+) " & _
'                                                " and nvl(substr(pa26,9,1),'0')=cu02(+) and substr(pa75,1,8)=fa01(+) and nvl(substr(pa75,9,1),'0')=fa02(+) "
'                                End If
'                                '­Y¬°°Ó¼Ð©µ®i
'                                If ("" & rsA.Fields(2).Value = "FCT" Or "" & rsA.Fields(2).Value = "CFT" Or "" & rsA.Fields(2).Value = "T" Or "" & rsA.Fields(2).Value = "TF") And ("" & rsA.Fields(7).Value = "102") Then
'                                    StrSQL083 = StrSQL083 & " union select Decode(TM33, Null, Nvl(TM40, Nvl(FA09, CU60)), TM71) from Trademark, customer, fagent " & _
'                                                " where " & ChgTradeMark(Replace(Text1(6).Text, "-", "")) & " and substr(tm23,1,8) = cu01(+) " & _
'                                                " and nvl(substr(tm23,9,1),'0')=cu02(+) and substr(tm44,1,8)=fa01(+) and nvl(substr(tm44,9,1),'0')=fa02(+) "
'                                Else
'                                    StrSQL083 = StrSQL083 & " union select nvl(tm40, Nvl(FA09, CU60)) from Trademark, customer, fagent " & _
'                                                " where " & ChgTradeMark(Replace(Text1(6).Text, "-", "")) & " and substr(tm23,1,8) = cu01(+) " & _
'                                                " and nvl(substr(tm23,9,1),'0')=cu02(+) and substr(tm44,1,8)=fa01(+) and nvl(substr(tm44,9,1),'0')=fa02(+) "
'                                End If
'                                Set Rs083 = New ADODB.Recordset
'                                Rs083.CursorLocation = adUseClient
'                                Rs083.Open StrSQL083, cnnConnection, adOpenStatic, adLockReadOnly
'                                If Rs083.RecordCount > 0 Then
'                                    tmp083 = CheckStr(Rs083.Fields(0).Value)
'                                Else
'                                    tmp083 = ""
'                                End If
'                                'Add By Cheng 2003/12/23
'                                'Ápµ¸¤H2­^
'                                '­Y¬°±M§Q¦~¶O
'                                If ("" & rsA.Fields(2).Value = "FCP" Or "" & rsA.Fields(2).Value = "CFP" Or "" & rsA.Fields(2).Value = "P") And ("" & rsA.Fields(7).Value = "605") Then
'                                    StrSQL083 = "select Decode(PA76, Null, Decode(pa53, Null, Decode(FA09, Null, Decode(CU60, Null, '', CU63), FA54), PA56), '') from patent,customer,fagent " & _
'                                                " where " & ChgPatent(Replace(Text1(6).Text, "-", "")) & " and substr(pa26,1,8) = cu01(+) " & _
'                                                " and nvl(substr(pa26,9,1),'0')=cu02(+) and substr(pa75,1,8)=fa01(+) and nvl(substr(pa75,9,1),'0')=fa02(+) "
'                                Else
'                                    StrSQL083 = "select Decode(pa53, Null, Decode(FA09, Null, Decode(CU60, Null, '', CU63), FA54), PA56) from patent,customer,fagent " & _
'                                                " where " & ChgPatent(Replace(Text1(6).Text, "-", "")) & " and substr(pa26,1,8) = cu01(+) " & _
'                                                " and nvl(substr(pa26,9,1),'0')=cu02(+) and substr(pa75,1,8)=fa01(+) and nvl(substr(pa75,9,1),'0')=fa02(+) "
'                                End If
'                                '­Y¬°°Ó¼Ð©µ®i
'                                If ("" & rsA.Fields(2).Value = "FCT" Or "" & rsA.Fields(2).Value = "CFT" Or "" & rsA.Fields(2).Value = "T" Or "" & rsA.Fields(2).Value = "TF") And ("" & rsA.Fields(7).Value = "102") Then
'                                    StrSQL083 = StrSQL083 & " union select Decode(TM33, Null, Decode(tm40, Null, Decode(FA09, Null, Decode(CU60, Null, '', CU62), FA54), TM43), '') from Trademark, customer, fagent " & _
'                                                " where " & ChgTradeMark(Replace(Text1(6).Text, "-", "")) & " and substr(tm23,1,8) = cu01(+) " & _
'                                                " and nvl(substr(tm23,9,1),'0')=cu02(+) and substr(tm44,1,8)=fa01(+) and nvl(substr(tm44,9,1),'0')=fa02(+) "
'                                Else
'                                    StrSQL083 = StrSQL083 & " union select Decode(tm40, Null, Decode(FA09, Null, Decode(CU60, Null, '', CU63), FA54), TM43) from Trademark, customer, fagent " & _
'                                                " where " & ChgTradeMark(Replace(Text1(6).Text, "-", "")) & " and substr(tm23,1,8) = cu01(+) " & _
'                                                " and nvl(substr(tm23,9,1),'0')=cu02(+) and substr(tm44,1,8)=fa01(+) and nvl(substr(tm44,9,1),'0')=fa02(+) "
'                                End If
'                                Set Rs083 = New ADODB.Recordset
'                                Rs083.CursorLocation = adUseClient
'                                Rs083.Open StrSQL083, cnnConnection, adOpenStatic, adLockReadOnly
'                                If Rs083.RecordCount > 0 Then
'                                    tmp083_1 = CheckStr(Rs083.Fields(0).Value)
'                                Else
'                                    tmp083_1 = ""
'                                End If
'                                'End
                                 'end 2010/12/14
                            End If
                            'Modify By Cheng 2003/12/23
                            '¥[Ápµ¸¤H2­^
'                            strExc(0) = "SELECT '" & ChgSQL(tmp083) & "','" & ChgSQL(tmp083_1) & "', FA05,FA63,FA64,FA65," & _
'                               " Decode(FA32,Null,FA18,FA32), Decode(FA32,Null,FA19,FA33), Decode(FA32,Null,FA20,FA34), Decode(FA32,Null,FA21,FA35), Decode(FA32,Null,FA22,FA36), FA70 FROM FAGENT,NATION WHERE " & St & " AND FA55=NA01(+)"
                            strExc(0) = "SELECT SUBSTR(FA23,1,20),SUBSTR(FA23,21,15)," & _
                               "SUBSTR(FA06,1,20),SUBSTR(FA06,21,20),'" & ChgSQL(tmp083) & "','" & ChgSQL(tmp083_1) & "',FA01||FA02 FROM FAGENT " & _
                               "WHERE " & St
                            'End
                        '­Y¥¼¿é¥»©Ò®×¸¹
                        Else
'                            strExc(0) = "SELECT SUBSTR(FA23,1,20),SUBSTR(FA23,21,15)," & _
'                               "SUBSTR(FA06,1,20),SUBSTR(FA06,21,20),FA01||FA02 FROM FAGENT " & _
'                               "WHERE " & St
                            strExc(0) = "SELECT SUBSTR(FA23,1,20),SUBSTR(FA23,21,15)," & _
                               "SUBSTR(FA06,1,20),SUBSTR(FA06,21,20),FA09,FA54,FA01||FA02 FROM FAGENT " & _
                               "WHERE " & St
                        End If
'                        Page = 4
                        Page = 1
                  End Select
               '¾÷Ãö¤å¸¹
               ElseIf opt1(2).Value = True Then '5
                  strExc(0) = "SELECT OR05,SUBSTR(OR06,1,15),SUBSTR(OR06,16,15),OR02,OR01 " & _
                     "FROM ORGANIZATION WHERE OR01='" & Text1(2).Text & "'"
                  Page = 4
               '¼ç¦b«È¤á¾ã§å¦C¦L add by toni 2008/11/05
               ElseIf opt1(3).Value = True Then
                     St = ChgPotCustomer(Text1(16).Text)
                     If Text1(5) <> "Y" Then
                        St = St & " AND (PCU34<>'N' OR PCU34 IS NULL)"
                     End If
                     '»y¤å
                     Select Case PriType
                        Case 1  '7
               '         '¤¤¤å¦a§},¤¤¤å¦WºÙ
                        strExc(0) = "SELECT decode(pcu39,NULL, SUBSTR(PCU27,1,20)),decode(pcu39,NULL,SUBSTR(PCU27,21,15))," & _
                        "SUBSTR(PCU08,1,20),SUBSTR(PCU08,21,20),  PCU01||PCU02 FROM PotCustomer " & _
                        "WHERE " & St
               
                          Page = 4
                        Case 2  '11
                              '­^¤å
                              strExc(0) = "SELECT '" & ChgSQL(tmp083) & "','" & ChgSQL(tmp083_1) & "',PCU03,Decode(PCU39,Null, DECODE(PCU29,'',PCU20,Pcu29),PCU39),"
                              strExc(0) = strExc(0) & " DECODE(PCU29,'',PCU21,PCU30), DECODE(PCU29,'',PCU22,PCU31)," & _
                                    "Decode(PCU39,Null, DECODE(PCU29,'',PCU23,PCU32),PCU39),Decode(PCU39,Null,DECODE(PCU29,'',PCU24,PCU33),PCU39)" & _
                                    ",Decode(PCU39,Null,DECODE(PCU29,'',PCU25),PCU39)  FROM PotCustomer WHERE " & St
                              
                           
                           '***************************************************   end
                           Page = 2
                     End Select
               'end  2008/11/05
               End If
               
               intI = 0
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
                'Modify By Cheng 2003/04/03
'               If intI <> 1 Then Exit Sub
               If intI <> 1 Then GoTo NextRecord
               Select Case Page
                  Case 1
                     IntF = 7
                  Case 2
                     IntF = 11
                  Case 3
                     IntF = 6
                  Case 4
                     IntF = 5
               End Select
               '¥ªÃä¬É
               Dim iCurrentX As Integer, iHeight As Integer
               iCurrentX = 0 + m_dbl_LeftMargin
                'Modify By Cheng 2003/01/15
                '¦C¦L®æ¦¡³]©w¬°­^¤å¦a§}±ø®æ¦¡, ¦C°ª¤£¦P
               If Page <> 2 Then
                  iHeight = 280
               Else
                  'Modify by Morgan 2009/4/17 ¦C°ª¤£°÷³¡¤À¦r¥Àªº¤U¥b¬q·|³QºI(Ex:g,y)
                  'iHeight = 230
                  iHeight = 270
               End If
               Printer.Font.Size = 12
               
                'Add By Cheng 2003/07/03
                '³]©w¦C¦L¦r«¬
                '­Y¬°¤¤¤é¤å
                If Me.Text1(4).Text = "1" Or Me.Text1(4).Text = 3 Then
                    Printer.Font.Name = "²Ó©úÅé"
                '­Y¬°­^¤å
                Else
                    Printer.Font.Name = "Times New Roman"
                End If
               For j = 1 To Val(Text1(3).Text)
                  RsTemp.MoveFirst
'                  iPrint = 1
                  With RsTemp
                     Select Case Page
                        Case 2
                           Do While Not .EOF
                              nRow = 0
                              IntF = .Fields.Count 'Add by Morgan 2007/8/29
                              For i = 0 To IntF - 1
                                 Printer.CurrentX = iCurrentX
                                 If IsNull(.Fields(i)) = False Then
                                    If IsEmptyText(.Fields(i)) = False Then
                                       ' »y¤å¬°­^¤å®É¤£ªÅ¦æ
                                       If Text1(4) = "2" Then
                                          Printer.CurrentY = nRow * iHeight + m_dbl_TopMargin
                                       Else
                                          Printer.CurrentY = i * iHeight + m_dbl_TopMargin
                                       End If
                                       nRow = nRow + 1
                                    End If
                                 End If
                                 
                                 'Added by Lydia 2022/05/02
                                 Xo = Printer.CurrentX
                                 Yo = Printer.CurrentY
                                 'end 2022/05/02
                                 
                                 If IsNull(.Fields(i)) = False Then
                                    If IsEmptyText(.Fields(i)) = False Then
                                       'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                                       'Printer.Print .Fields(i)
                                       PUB_PrintUnicodeText .Fields(i), Xo, Yo, 0
                                    End If
                                 End If
                                'Modify By Cheng 2002/12/20
                                 If i = 9 Then
                                    Printer.CurrentX = 3200 + iCurrentX
                                    If Text1(4) = "2" Then
                                       Printer.CurrentY = (nRow - 1) * iHeight + m_dbl_TopMargin
                                    Else
                                       Printer.CurrentY = (i - 1) * iHeight + m_dbl_TopMargin
                                    End If
                                 End If
                              Next
                              iPrint = iPrint + 1
'                              Printer.NewPage
                              .MoveNext
                           Loop
                           
                        Case Else
                        
                           Do While Not .EOF
                              '³]©w°_©l¦C¼Æ
                              Select Case IntF
                                 Case 7
                                     intBgnRow = 1
                                 Case 6
                                     intBgnRow = 1.5
                                 Case 5
                                     intBgnRow = 2
                              End Select
                              nRow = 0
                              
                              'Modify by Morgan 2008/8/8 ¥Ó½Ð¤H¤¤¤å¦a§}§ï©I¥s¦@¥Î¨ç¼Æ
                              If opt1(0).Value = True And Text1(4) = "1" Then
                                 ReDim stAddr(6)
                                 '«È¤á½s¸¹
                                 stAddr(6) = ChangeCustomerL(Text1(0).Text)
                                 If PUB_GetAddrRef(stAddr(6), "" & rsA.Fields(2), "" & rsA.Fields(3), "" & rsA.Fields(4), "" & rsA.Fields(5), stAddr(3), stAddr(5), stAddr(0), stAddr(1)) = True Then
                                    If Len(stAddr(1)) > 20 Then
                                       stAddr(2) = Mid(stAddr(1), 21)
                                       stAddr(1) = Left(stAddr(1), 20)
                                    End If
                                    If Len(stAddr(3)) > 20 Then
                                       stAddr(4) = Mid(stAddr(1), 21)
                                       stAddr(3) = Left(stAddr(1), 20)
                                    End If
                                 End If
                              Else
                                 ReDim stAddr(IntF - 1)
                                 For i = 0 To IntF - 1
                                    stAddr(i) = "" & .Fields(i)
                                 Next
                              End If
                              
                              For i = 0 To IntF - 1
                                 Printer.CurrentX = iCurrentX
                                 If IsEmptyText(stAddr(i)) = False Then
                                    ' »y¤å¬°­^¤å®É¤£ªÅ¦æ
                                    If Text1(4) = "2" Then
                                       '¥[°_©l¦C¼Æ
                                       Printer.CurrentY = (nRow + intBgnRow) * iHeight + m_dbl_TopMargin
                                    Else
                                       '¥[°_©l¦C¼Æ
                                       Printer.CurrentY = (i + intBgnRow) * iHeight + m_dbl_TopMargin
                                    End If
                                    nRow = nRow + 1
                                 End If
                                 
                                 'Added by Lydia 2022/05/02
                                 Xo = Printer.CurrentX
                                 Yo = Printer.CurrentY
                                 'end 2022/05/02
                     
                                 If IsEmptyText(stAddr(i)) = False Then
                                     If i = 6 Then
                                         '«È¤á¥N¸¹«á¥[¥»©Ò®×¸¹
                                         Printer.Print stAddr(i) & IIf(m_CaseNo <> "", "¡@( " & m_CaseNo & " )", "")
                                     Else
                                         '­Y¬°¤¤¤å¦a§}
                                         'Modify by Morgan 2009/5/7
                                         'If Me.Text1(4).Text = "1" Then
                                         If Me.Text1(4).Text = "1" And opt1(0).Value = True Then
                                             '¦C¦LÄæ¦ì¬°¤½¥q¦WºÙ©Î­Ó¤H¦WºÙ
                                             '93.5.23 ADD BY SONIA
                                             strExc(0) = stAddr(i)
                                             If strExc(0) = "¼B¬ì¨}¡D·Å©y¬Â" Then
                                                strExc(0) = "·Å©y¬Â"
                                             End If
                                             '93.5.23 END
                                             If i = 3 Then
                                                 '­Y¬°­Ó¤H¤á®É
                                                 If stAddr(3) <> "" And stAddr(5) = "" Then
                                                     'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                                                     'Printer.Print strExc(0) & "¡@¡@¡@¡@¡@§g¡@¶v±Ò"
                                                     PUB_PrintUnicodeText strExc(0) & "¡@¡@¡@¡@¡@§g¡@¶v±Ò", Xo, Yo, 0
                                                 Else
                                                     'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                                                     'Printer.Print strExc(0)
                                                     PUB_PrintUnicodeText strExc(0), Xo, Yo, 0
                                                 End If
                                             ElseIf i = 5 Then
                                                 '­Y¬°¤½¥q¤á®É
                                                 If stAddr(3) <> "" And stAddr(5) <> "" Then
                                                     'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                                                     'Printer.Print stAddr(i) & "¡@¡@¡@¡@¡@§g¡@¶v±Ò"
                                                     PUB_PrintUnicodeText stAddr(i) & "¡@¡@¡@¡@¡@§g¡@¶v±Ò", Xo, Yo, 0
                                                 Else
                                                     'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                                                     'Printer.Print stAddr(i)
                                                     PUB_PrintUnicodeText stAddr(i), Xo, Yo, 0
                                                 End If
                                             Else
                                                 'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                                                 'Printer.Print stAddr(i)
                                                 PUB_PrintUnicodeText stAddr(i), Xo, Yo, 0
                                             End If
                                         '­Y«D¤¤¤å¦a§}
                                         Else
                                             'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                                             'Printer.Print stAddr(i)
                                             PUB_PrintUnicodeText stAddr(i), Xo, Yo, 0
                                         End If
                                     End If
                                 End If
                              Next
                              
                              Printer.CurrentX = 4000 + iCurrentX
                              If Text1(4) = "2" Then
                                 Printer.CurrentY = ((nRow - 1) + intBgnRow) * iHeight + m_dbl_TopMargin
                              Else
                                 Printer.CurrentY = ((i - 1) + intBgnRow) * iHeight + m_dbl_TopMargin
                              End If
'                              If m_PageNo > 0 Then
                              If Val("0" & m_PageNo) > 0 Then
                                 Printer.Print Format(m_PageNo, "000000")
                              Else
                                 Printer.Print Format(iPrint, "000000")
                              End If
                              iPrint = iPrint + 1
'                              Printer.NewPage
                              .MoveNext
                           Loop
                     End Select
                  End With
               Next
'*************************************************
                '­Y¦³°Æ¥»¦¬¨ü¤H, ¥B¥¼¦C¦L°Æ¥»¦¬¨ü¤H
                If "" & rsA.Fields(9).Value <> "" And blnPrintCC = False Then
                    Printer.NewPage
                    If Left("" & rsA.Fields(9).Value, 1) = "Y" Then
                        Me.opt1(1).Value = True
                        Me.Text1(1).Text = "" & rsA.Fields(9).Value
                    ElseIf Left("" & rsA.Fields(9).Value, 1) = "X" Then
                        Me.opt1(0).Value = True
                        Me.Text1(0).Text = "" & rsA.Fields(9).Value
                    End If
                    tmp083 = "" & rsA.Fields(10).Value
                    tmp083_1 = ""
                    blnPrintCC = True
                    GoTo PrintCC
                End If
                dblCnt = dblCnt + 1 'Add By Sindy 2010/10/4
NextRecord:
                rsA.MoveNext
                If rsA.EOF = False Then
                    Printer.NewPage
                Else
                    Printer.EndDoc
                End If
            Wend
            InsertQueryLog (dblCnt) 'Add By Sindy 2010/10/4
            'Add By Cheng 2003/02/12
            '¥i­«ÂÐ¦C¦L¦a§}±ø
            If MsgBox("¦a§}±ø¤w¦C¦L§¹²¦¡A±z¬O§_­n­«·s¦C¦L???", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
                GoTo RePrint
            End If
        End If
    End If
    pub_blnBatchPrintAddress = False
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End Sub

'Add By Cheng 2003/05/20
Private Sub FormClear()
    On Error Resume Next
    '2008/11/12 move by sonia ²¾¨ì¤U­±¦A²M
    ''¥»©Ò®×¸¹
    'Me.Text1(9).Text = ""
    'Me.Text1(10).Text = ""
    'Me.Text1(11).Text = ""
    'Me.Text1(12).Text = ""
    
    '¿ï¶µ
    Me.Text1(0).Text = "X"
    Me.Text1(1).Text = "Y"
    Me.Text1(2).Text = ""
    Me.lblFM2(0).Caption = ""
    Me.lblFM2(1).Caption = ""
    Me.lblFM2(3).Caption = ""
    Me.Text1(3).Text = "1"
    '2008/10/30 add by Toni
    Me.Text1(16).Text = "R"
    Me.lblFM2(2) = ""
    'end 2008/10/30
    'Modify By Cheng 2004/03/18
    '­Y­ì¦³¿é®×¸¹«h²MªÅ»y¤åÄæ¦ì
'    Me.Text1(4).Text = ""
    If Me.Text1(9).Text <> "" And Me.Text1(10).Text <> "" Then
        Me.Text1(4).Text = ""
    End If
    'End
    Me.Text1(5).Text = "Y"
    'Modify By Cheng 2003/12/17
    '³]©w¿é¤JµJÂI
'    Me.Opt1(0).Value = True
'    Me.Text1(9).SetFocus
    If Me.Text1(9).Text <> "" And Me.Text1(10).Text <> "" Then
        Me.Text1(9).SetFocus
    ElseIf Me.opt1(0).Value = True Then
        Me.Text1(0).SetFocus
    ElseIf Me.opt1(1).Value = True Then
        Me.Text1(1).SetFocus
    Else
        Me.Text1(2).SetFocus
    End If
    Text1(6).Text = "" 'Add by Morgan 2004/11/4
    '2008/11/12 move by sonia ±q¤W­±²¾¹L¨Ó,§_«h¤W­±§PÂ_·|µL®Ä
    '¥»©Ò®×¸¹
    Me.Text1(9).Text = ""
    Me.Text1(10).Text = ""
    Me.Text1(11).Text = ""
    Me.Text1(12).Text = ""
    
End Sub
'Add by Morgan 2008/7/29
'¨ú±o­Ó®×Ápµ¸¤H½s¸¹
Private Function GetCaseContactNo(p_CaseNo1 As String, p_CaseNo2 As String, p_CaseNo3 As String, p_CaseNo4 As String, p_CustNo As String) As String
   Dim stSysNo As String, stSQL As String, intR As Integer, stCaseNo As String
   stCaseNo = Trim(p_CaseNo1 & p_CaseNo2 & p_CaseNo3 & p_CaseNo4)
   '­Y¦³¥»©Ò®×¸¹
   If stCaseNo <> "" And p_CustNo <> "" Then
     stSysNo = CheckSys(p_CaseNo1)
     Select Case stSysNo
        Case "1"
            stSQL = "select pa149 from patent where " & ChgPatent(stCaseNo) & " and pa26='" & ChangeCustomerL(p_CustNo) & "'"
        Case "2"
            stSQL = "select tm123 from trademark where " & ChgTradeMark(stCaseNo) & " and tm23='" & ChangeCustomerL(p_CustNo) & "'"
        Case "3"
            stSQL = "select lc42 from Lawcase where " & ChgLawcase(stCaseNo) & " and lc11='" & ChangeCustomerL(p_CustNo) & "'"
        Case "4"
            stSQL = "select hc23 from Hirecase where " & ChgHirecase(stCaseNo) & " and hc05='" & ChangeCustomerL(p_CustNo) & "'"
        Case Else
            stSQL = "select sp78 from Servicepractice where " & ChgService(stCaseNo) & " and sp08='" & ChangeCustomerL(p_CustNo) & "'"
     End Select
     intR = 1
     Set RsTemp = ClsLawReadRstMsg(intR, stSQL)
     If intR = 1 Then
        GetCaseContactNo = "" & RsTemp(0)
     End If
   End If
End Function
'Add by Morgan 2008/8/8
'¥Ó½Ð¤H¤¤¤å¦a§}
Sub PrintAppChinese(ByRef nRow As Integer, ByRef i As Integer, ByRef intBgnRow As Integer, ByRef iHeight As Integer, iCurrentX As Integer)
   Dim stAddr(6) As String '¦a§}±ø¬ÛÃö¸ê®Æ
   '«È¤á½s¸¹
   stAddr(6) = Text1(0).Text
   If PUB_GetAddrRef(stAddr(6), Text1(9).Text, Text1(10).Text, Text1(11).Text, Text1(12).Text, stAddr(3), stAddr(5), stAddr(0), stAddr(1)) = True Then
      If Len(stAddr(1)) > 20 Then
         stAddr(2) = Mid(stAddr(1), 21)
         stAddr(1) = Left(stAddr(1), 20)
      End If
      If Len(stAddr(3)) > 20 Then
         stAddr(4) = Mid(stAddr(1), 21)
         stAddr(3) = Left(stAddr(1), 20)
      End If
   End If
   
   nRow = 0
   For i = 0 To 6
      Printer.CurrentX = iCurrentX
      Printer.CurrentY = (i + intBgnRow) * iHeight + m_dbl_TopMargin
      nRow = nRow + 1
      If i = 6 Then
          '«È¤á¥N¸¹«á¥[¥»©Ò®×¸¹
          Printer.Print stAddr(i) & IIf(m_CaseNo <> "", "¡@( " & m_CaseNo & " )", "")
      ElseIf i = 3 Then
         '­Y¬°­Ó¤H¤á®É
         If stAddr(3) <> "" And stAddr(5) = "" Then
             Printer.Print stAddr(3) & "¡@¡@¡@¡@¡@§g¡@¶v±Ò"
         Else
             Printer.Print stAddr(3)
         End If
      ElseIf i = 5 Then
         If stAddr(5) <> "" Then
            Printer.Print stAddr(5) & "¡@¡@¡@¡@¡@§g¡@¶v±Ò"
          End If
      Else
          Printer.Print stAddr(i)
      End If
   Next
   
End Sub
'add by Toni 2008/10/28
Private Function ChgPotCustomer(ByVal strTemp As String) As String
 On Error GoTo ErrHand
   If strTemp = "" Then GoTo ErrHand
   
   If Len(strTemp) = 9 Then
      ChgPotCustomer = "PCU01='" & Left(strTemp, 8) & "' AND PCU02='" & Right(strTemp, 1) & "'"
   Else
      ChgPotCustomer = "PCU01='" & strTemp & String(8 - Len(strTemp), "0") & "' AND PCU02='0'"
   End If
   Exit Function
ErrHand:
   ChgPotCustomer = "PCU01 IS NULL AND PCU02 IS NULL"
End Function
'end 2008/10/28

'add by Toni 2008/10/28
Private Function ClsPCUGetContact(ByRef strAgent As String, ByRef strAgentName As String) As Boolean
   Dim PCU01 As String, PCU02 As String, iPos As Integer
   iPos = InStr(strAgent, "-")
   PCU01 = Left(strAgent & "000", 8)
   PCU02 = Right(strAgent & "0", 1)
   
   strAgent = PCU01 & "-" & PCU02
   strExc(0) = "select nvl(pcu08,nvl(pcu03,pcu07)) from PotCustomer where pcu01='" & Left(PCU01, 8) & "'  and pcu02='" & Right(PCU02, 1) & "'"

   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strAgentName = "" & RsTemp.Fields(0)
      ClsPCUGetContact = True
   '2008/10/31 ADD BY SONIA
   Else
      MsgBox "¼ç¦b«È¤á¥N½X¿ù»~!", vbCritical + vbOKOnly, MsgText(9001)
   '2008/10/31 END
   End If
End Function
'end 2008/10/28

'Add by Morgan 2010/12/14
'³]©wÁpµ¸¤H
'In:pCaseNo=¥»©Ò®×¸¹,pSys=¨t²Î§O,pProperty=®×¥ó©Ê½è,pWho=¦C¦L¹ï¶H(0:¥Ó½Ð¤H,1:¥N²z¤H),pLanguage=©w½Z»y¤å(1:¤¤,2:­^,3:¤é)
'Out:pContact1=Ápµ¸¤H1,pContact2=Ápµ¸¤H2
Private Sub SetCaseContact(ByVal pCaseNo As String, ByVal pSys As String, ByVal pProperty As String, _
   ByVal pWho As String, ByVal pLanguage As String, ByRef pContact1 As String, ByRef pContact2 As String)
   Dim stSQL As String, rsTmp As ADODB.Recordset, iR As Integer
   Dim stCt1 As String, stCt2 As String
   
   pCaseNo = Replace(pCaseNo, "-", "")
   '±M§Q
   If (pSys = "FCP" Or pSys = "CFP" Or pSys = "P") Then
      Select Case pLanguage
         Case 1 '¤¤
            '¥Ó½Ð¤H
            If pWho = 0 Then
               stCt1 = "nvl(pa51,CU58)"
               stCt2 = "Decode(pa51, Null, CU61, PA54)"
               
            '¥N²z¤H
            ElseIf pWho = 1 Then
               stCt1 = "nvl(pa51, FA07)"
               stCt2 = "Decode(pa51, Null, FA52, PA54)"
                           
            Else
               stCt1 = "nvl(pa51, Nvl(FA07, CU58))"
               stCt2 = "Decode(pa51, Null, Decode(FA07, Null, Decode(CU58, Null, '', CU61), FA52), PA54)"
            End If
            
         Case 2 '­^
            '¥Ó½Ð¤H
            If pWho = 0 Then
               stCt1 = "nvl(pa52, CU59)"
               stCt2 = "Decode(pa52, Null, CU62, PA55)"
               
            '¥N²z¤H
            ElseIf pWho = 1 Then
               stCt1 = "nvl(pa52, FA08)"
               stCt2 = "Decode(pa52, Null, FA53, PA55)"
               
            Else
               stCt1 = "nvl(pa52, Nvl(FA08, CU59))"
               stCt2 = "Decode(pa52, Null, Decode(FA08, Null, Decode(CU59, Null, '', CU62), FA53), PA55)"
            End If
            
         Case Else '¤é
            '¥Ó½Ð¤H
            If pWho = 0 Then
               stCt1 = "nvl(pa53, CU60)"
               stCt2 = "Decode(pa53, Null, CU63, PA56)"
               
            '¥N²z¤H
            ElseIf pWho = 1 Then
               stCt1 = "nvl(pa53, FA09)"
               stCt2 = "Decode(pa53, Null, FA54, PA56)"
            
            Else
               stCt1 = "Nvl(PA53, Nvl(FA09, CU60))"
               stCt2 = "Decode(pa53, Null, Decode(FA09, Null, Decode(CU60, Null, '', CU63), FA54), PA56)"
            End If
      End Select
      
      '¦~¶O
      If pProperty = "605" Then
         stCt1 = "Decode(PA76, Null, " & stCt1 & ", PA135)"
         stCt2 = "Decode(PA76, Null, " & stCt2 & ", '')"
      End If
      
      stSQL = "select " & stCt1 & " as c1," & stCt2 & " as c2" & _
         " from patent,customer,fagent " & _
         " where " & ChgPatent(pCaseNo) & " and cu01(+)=substr(pa26,1,8) " & _
         " and cu02(+)=substr(pa26,9) and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9)"
   
   '°Ó¼Ð
   ElseIf (pSys = "FCT" Or pSys = "CFT" Or pSys = "T" Or pSys = "TF") Then
      Select Case pLanguage
         Case 1 '¤¤
            '¥Ó½Ð¤H
            If pWho = 0 Then
               stCt1 = "nvl(TM38,CU58)"
               stCt2 = "Decode(TM38, Null, CU61, TM41)"
               
            '¥N²z¤H
            ElseIf pWho = 1 Then
               stCt1 = "nvl(TM38, FA07)"
               stCt2 = "Decode(TM38, Null, FA52, TM41)"
            
            Else
               stCt1 = "nvl(TM38, Nvl(FA07, CU58))"
               stCt2 = "Decode(TM38, Null, Decode(FA07, Null, Decode(CU58, Null, '', CU61), FA52), TM41)"
            End If
            
         Case 2 '­^
            '¥Ó½Ð¤H
            If pWho = 0 Then
               stCt1 = "nvl(TM39, CU59)"
               stCt2 = "Decode(TM39, Null, CU62, TM42)"
            
            '¥N²z¤H
            ElseIf pWho = 1 Then
               stCt1 = "nvl(TM39, FA08)"
               stCt2 = "Decode(TM39, Null, FA53, TM42)"
            
            Else
               stCt1 = "nvl(TM39, Nvl(FA08, CU59))"
               stCt2 = "Decode(TM39, Null, Decode(FA08, Null, Decode(CU59, Null, '', CU62), FA53), TM42)"
            End If
            
         Case Else '¤é
            '¥Ó½Ð¤H
            If pWho = 0 Then
               stCt1 = "nvl(TM40, CU60)"
               stCt2 = "Decode(TM40, Null, CU63, TM43)"
            
            '¥N²z¤H
            ElseIf pWho = 1 Then
               stCt1 = "nvl(TM40, FA09)"
               stCt2 = "Decode(TM40, Null, FA54, TM43)"
            
            Else
               stCt1 = "Nvl(TM40, Nvl(FA09, CU60))"
               stCt2 = "Decode(TM40, Null, Decode(FA09, Null, Decode(CU60, Null, '', CU63), FA54), TM43)"
            End If
      End Select
   
      '©µ®i
      If pProperty = "102" Then
         stCt1 = "Decode(TM33, Null, " & stCt1 & ", TM71)"
         stCt2 = "Decode(TM33, Null, " & stCt2 & ", '')"
      End If
      
      stSQL = "select " & stCt1 & " as c1," & stCt2 & " as c2" & _
         " from Trademark, customer, fagent " & _
         " where " & ChgTradeMark(pCaseNo) & " and cu01(+)=substr(tm23,1,8) " & _
         " and cu02(+)=substr(tm23,9) and fa01(+)=substr(tm44,1,8) and fa02(+)=substr(tm44,9)"
   
   'ªk°È
   ElseIf InStr(pSys, "L") > 0 Then
      Select Case pLanguage
         Case 1 '¤¤
            '¥Ó½Ð¤H
            If pWho = 0 Then
               stCt1 = "nvl(LC18, CU58)"
               stCt2 = "Decode(LC18, Null, CU61, '')"
            
            '¥N²z¤H
            ElseIf pWho = 1 Then
               stCt1 = "nvl(LC18, FA07)"
               stCt2 = "Decode(LC18, Null, FA52, '')"
            
            Else
               stCt1 = "nvl(LC18, Nvl(FA07, CU58))"
               stCt2 = "Decode(LC18, Null, Decode(FA07, Null, Decode(CU58, Null, '', CU61), FA52), '')"
            End If
            
         Case 2 '­^
            '¥Ó½Ð¤H
            If pWho = 0 Then
               stCt1 = "nvl(LC19, CU59)"
               stCt2 = "Decode(LC19, Null, CU62, '')"
            
            '¥N²z¤H
            ElseIf pWho = 1 Then
               stCt1 = "nvl(LC19, FA08)"
               stCt2 = "Decode(LC19, Null, FA53, '')"
            
            Else
               stCt1 = "nvl(LC19, Nvl(FA08, CU59))"
               stCt2 = "Decode(LC19, Null, Decode(FA08, Null, Decode(CU59, Null, '', CU62), FA53), '')"
            End If
            
         Case Else '¤é
            '¥Ó½Ð¤H
            If pWho = 0 Then
               stCt1 = "nvl(LC20, CU60)"
               stCt2 = "Decode(LC20, Null, CU63, '')"
            
            '¥N²z¤H
            ElseIf pWho = 1 Then
               stCt1 = "nvl(LC20, FA09)"
               stCt2 = "Decode(LC20, Null, FA54, '')"
            
            Else
               stCt1 = "Nvl(LC20, Nvl(FA09, CU60))"
               stCt2 = "Decode(LC20, Null, Decode(FA09, Null, Decode(CU60, Null, '', CU63), FA54), '')"
            End If
      End Select
   
      stSQL = "select " & stCt1 & " as c1," & stCt2 & " as c2" & _
         " from LAWCASE, customer, fagent " & _
         " where " & ChgLawcase(pCaseNo) & " and cu01(+)=substr(LC11,1,8) " & _
         " and cu02(+)=substr(LC11,9) and fa01(+)=substr(LC22,1,8) and fa02(+)=substr(LC22,9)"
         
   'ÅU°Ý
   ElseIf pSys = "LA" Then
      Select Case pLanguage
         Case 1 '¤¤
               stCt1 = "CU58"
               stCt2 = "CU61"
            
         Case 2 '­^
               stCt1 = "CU59"
               stCt2 = "CU62"
            
         Case Else '¤é
               stCt1 = "CU60"
               stCt2 = "CU63"
               
      End Select
   
      stSQL = "select " & stCt1 & " as c1," & stCt2 & " as c2" & _
         " from hirecase, customer" & _
         " where " & ChgHirecase(pCaseNo) & " and cu01(+)=substr(hc05,1,8) " & _
         " and cu02(+)=substr(hc05,9)"
         
   'ªA°È
   Else
      Select Case pLanguage
         Case 1 '¤¤
            '¥Ó½Ð¤H
            If pWho = 0 Then
               stCt1 = "nvl(SP30, CU58)"
               stCt2 = "Decode(SP30, Null, CU61, SP75)"
            
            '¥N²z¤H
            ElseIf pWho = 1 Then
               stCt1 = "nvl(SP30, FA07)"
               stCt2 = "Decode(SP30, Null, FA52, SP75)"
            
            Else
               stCt1 = "nvl(SP30, Nvl(FA07, CU58))"
               stCt2 = "Decode(SP30, Null, Decode(FA07, Null, Decode(CU58, Null, '', CU61), FA52), SP75)"
            End If
            
         Case 2 '­^
            '¥Ó½Ð¤H
            If pWho = 0 Then
               stCt1 = "nvl(SP30, CU58)"
               stCt2 = "Decode(SP30, Null, CU61, SP75)"
            
            '¥N²z¤H
            ElseIf pWho = 1 Then
               stCt1 = "nvl(SP30, FA07)"
               stCt2 = "Decode(SP30, Null, FA52, SP75)"
            
            Else
               stCt1 = "nvl(SP30, Nvl(FA07, CU58))"
               stCt2 = "Decode(SP30, Null, Decode(FA07, Null, Decode(CU58, Null, '', CU61), FA52), SP75)"
            End If
            
         Case Else '¤é
            '¥Ó½Ð¤H
            If pWho = 0 Then
               stCt1 = "nvl(SP30, CU60)"
               stCt2 = "Decode(SP30, Null, CU63, SP75)"
            
            '¥N²z¤H
            ElseIf pWho = 1 Then
               stCt1 = "nvl(SP30, FA09)"
               stCt2 = "Decode(SP30, Null, FA54, SP75)"
            
            Else
               stCt1 = "Nvl(SP30, Nvl(FA09, CU60))"
               stCt2 = "Decode(SP30, Null, Decode(FA09, Null, Decode(CU60, Null, '', CU63), FA54), SP75)"
            End If
            
      End Select
   
      stSQL = "select " & stCt1 & " as c1," & stCt2 & " as c2" & _
         " from ServicePractice, customer, fagent " & _
         " where " & ChgService(pCaseNo) & " and cu01(+)=substr(SP08,1,8) " & _
         " and cu02(+)=substr(SP08,9) and fa01(+)=substr(SP26,1,8) and fa02(+)=substr(SP26,9)"
   
   End If
   
   iR = 1
   Set rsTmp = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      pContact1 = "" & rsTmp.Fields(0)
      pContact2 = "" & rsTmp.Fields(1)
   End If
   Set rsTmp = Nothing
End Sub

'Added By Lydia 2016/10/28 A4§å¦¸¦C¦L¦a§}±ø
'Modified by Lydia 2017/11/03 +«ü©w¦C¦L pNoList
Private Function PrintCaseBatchA4(Optional ByVal pNoList As String) As Boolean
Dim rsA As New ADODB.Recordset
Dim i As Integer
Dim iPrint As Integer
Dim IntF As Integer
Dim j As Integer
Dim intBgnRow As Integer '°_©l¦C¼Æ
Dim stAddr() As String '¦a§}±ø¬ÛÃö¸ê®Æ
Dim iHeight As Integer '¦C°ª
Dim m_HT As Single, m_WT  As Single '³æ±i°ª«×/¼e«×
Dim intTop As Double, intLeft As Double '¦C¦L¸ê®Æªº¤W¡B¥ªÃä¬É
Dim m_PrtOrientation As Integer '¦C¦L¤è¦V
Dim m_PrtScaleMode As Integer '¦C¦L®y¼Ð³æ¦ì
Dim d_Top As Double, d_Left As Double '¦Lªí¾÷ªº³Ì¤p¿é¥XÃä¬É
Dim lngTop As Double '¹w³]¦C¦Lªº¤WÃä¬É
Dim lngLeft As Double '¹w³]¦C¦Lªº¥ªÃä¬É
Dim byTwips As Integer '¨C¤½¤Àªº³æ¦ì 'twips ,¨C¤½¤À=567
Dim iStkrNo As Integer 'Added by Morgan 2024/3/14

On Error Resume Next
      
'³æ¦ì: Twips
byTwips = 567

'³æ¦ì: ¤½¤À
lngLeft = 1.3: lngTop = 0.3 '¹w³]¦C¦LªºÃä¬É
m_HT = 3.7: m_WT = 10.5 '³æ±i°ª«×/¼e«×
   
   '¦C¦Lªº¸ê®Æ«á­±¤~§ì
   'Added by Lydia 2017/11/03 «ü©w¦C¦L
   If pNoList <> "" Then
      strExc(0) = "SELECT * FROM ADDRESSA4LIST Where AAL01||AAL02||AAL03 IN (" & GetAddStr(pNoList) & ") "
   Else
   'end 2017/11/03
      strExc(0) = "SELECT * FROM ADDRESSA4LIST Where AAL01='" & strUserNum & "' "
   End If 'end 2017/11/03
   
   strExc(0) = strExc(0) & " order by 1,2,3" 'Added by Morgan 2024/3/14 ­n¨Ì·Ó·s¼Wªº¶¶§Ç¦C¦L--ÀR¿Z
   
   intI = 0
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      InsertQueryLog (0)
      Exit Function
   End If
If MsgBox("·Ç³Æ¦C¦LA4¦W±ø¡A½Ð§ó´«¯È±i!!!", vbExclamation + vbOKCancel) = vbOK Then

RePrint:
   PrintCaseBatchA4 = True
   
   '³]©w¯È±i©M¤è¦V
   Printer.PaperSize = 9 'A4
   Printer.Orientation = 1 'ª½¦L
   '¨ú±o¹w³]¦Lªí¾÷³]©w­È
   m_PrtOrientation = Printer.Orientation
   m_PrtScaleMode = Printer.ScaleMode
   '¦Lªí¾÷ªº¿é¥XÃä¬É by ¤½¤À
   d_Top = Format((Printer.Height - Printer.ScaleHeight) / byTwips / 2, "0.000")
   d_Left = Format((Printer.Width - Printer.ScaleWidth) / byTwips / 2, "0.000")
   
   Printer.KillDoc
    
   '¦C¦L®æ¦¡³]©w¬°­^¤å¦a§}±ø®æ¦¡, ¦C°ª¤£¦P
   iHeight = 255

   IntF = 7
   
   Printer.Font.Size = 12
   Printer.Font.Name = "Times New Roman"

    InsertQueryLog (rsA.RecordCount)
    
    '¨C­¶¹w³]¦C¦L°_ÂI
    intTop = 0 '²Ä1~2±i¤£¯d¶¡¹j '(lngTop - d_Top) * byTwips
    intLeft = (lngLeft - d_Left) * byTwips
    
    '¥÷¼Æ
   For j = 1 To Val(Text1(3).Text)
      rsA.MoveFirst
      'Modified by Morgan 2024/3/14 ¼W¥[¥i«ü©w±q²Ä´X±i¶K¯È¶}©l¦L(¤W¦¸¨S¥Î§¹ªº¶K¯È)
      'iPrint = 1
      If j = 1 Then
         iPrint = Val(txtStartPage)
         If iPrint > 1 Then
            If iPrint = 2 Then
               m_dbl_TopMargin = intTop: m_dbl_LeftMargin = intLeft
            ElseIf iPrint Mod 2 = 0 Then
               '©_¼Æ±i¦C¦Lªº¤WÃä¬É:¦b¯È¤Wªºµ´¹ï¦ì¸m(¦©°£¦Lªí¾÷ªº³Ì¤p¿é¥X¤§¤WÃä¬É) + ³æ±i¹w³]¤WÃä¬É
               m_dbl_TopMargin = Format((((iPrint - 1) \ 2) * m_HT) + lngTop - d_Top, "0.000") * byTwips
            End If
         End If
      Else
         iPrint = 1
      End If
      iStkrNo = iPrint - 1
      'end 2024/3/14
      With rsA
         Do While Not .EOF
            'Modified by Morgan 2024/3/14
            'If rsA.AbsolutePosition Mod 16 = 1 Then
            '   If rsA.AbsolutePosition > 16 Then Printer.NewPage
            iStkrNo = iStkrNo + 1
            If iStkrNo Mod 16 = 1 Then
               If iStkrNo > 16 Then Printer.NewPage
               iPrint = 1
            'end 2024/3/14
               m_dbl_TopMargin = intTop: m_dbl_LeftMargin = intLeft
            Else
               '½Õ¾ã¦ì¸m
                If iPrint Mod 2 = 1 Then
                   If iPrint Mod 15 = 0 Then '²Ä15~16±i°£¥~,¹ê´ú¯È±i26
                      m_dbl_TopMargin = (26 - d_Top) * byTwips
                   Else
                      '©_¼Æ±i¦C¦Lªº¤WÃä¬É:¦b¯È¤Wªºµ´¹ï¦ì¸m(¦©°£¦Lªí¾÷ªº³Ì¤p¿é¥X¤§¤WÃä¬É) + ³æ±i¹w³]¤WÃä¬É
                      m_dbl_TopMargin = Format(((iPrint \ 2) * m_HT) + lngTop - d_Top, "0.000") * byTwips
                   End If
                   m_dbl_LeftMargin = intLeft
                Else
                   '°¸¼Æ±i¦C¦Lªº¥ªÃä¬É:¦b¯È¤Wªºµ´¹ï¦ì¸m(¦©°£¦Lªí¾÷ªº³Ì¤p¿é¥X¤§¥ªÃä¬É)
                   m_dbl_LeftMargin = (m_WT + 1 - d_Left) * byTwips
                End If
            End If
            '³]©w°_©l¦C¼Æ
            intBgnRow = 0
         
            '¥Ó½Ð¤H¤¤¤å¦a§}§ï©I¥s¦@¥Î¨ç¼Æ
            Erase stAddr
            ReDim stAddr(6)
            
            'Modified by Lydia 2016/11/04 ¦³¨âºØ¨Ó·½
            If Mid(.Fields("AAL04"), 1, 1) = "X" Then
                '«È¤á½s¸¹(±q¥Ó½Ð¤H¬d¸ß¨ú±o)
                'stAddr(6) = Mid(.Fields("A07"), 1, 9)
                stAddr(6) = Mid(.Fields("AAL04"), 1, 9)
                strExc(1) = IIf(Mid(.Fields("AAL04"), 10, 1) = "-", Mid(.Fields("AAL04"), 11), "") '¹w³]±µ¬¢¤H
                'Modified by Lydia 2017/11/03 ±j¨îÅª±µ¬¢¤H¦WºÙ
                'If PUB_GetAddrRef(stAddr(6), "", "", "", "", stAddr(3), stAddr(5), stAddr(0), stAddr(1), , strExc(1)) = True Then
                'Modified by Morgan 2025/1/6 ´¼Åv¦LA4¦a§}±ø®É³£­n±a¥X¦a§} Ex:X54243090
                'If PUB_GetAddrRef(stAddr(6), "", "", "", "", stAddr(3), stAddr(5), stAddr(0), stAddr(1), , strExc(1), IIf(strExc(1) <> "", True, False)) = True Then
                If PUB_GetAddrRef(stAddr(6), "", "", "", "", stAddr(3), stAddr(5), stAddr(0), stAddr(1), , strExc(1), IIf(strExc(1) <> "" Or iStiu = 1, True, False)) = True Then
                'end 2025/1/6
                   If Len(stAddr(1)) > 20 Then
                      stAddr(2) = Mid(stAddr(1), 21)
                      stAddr(1) = Left(stAddr(1), 20)
                   End If
                   If Len(stAddr(3)) > 20 Then
                      stAddr(4) = Mid(stAddr(3), 21)
                      stAddr(3) = Left(stAddr(3), 20)
                   End If
                End If
            Else
                '¥»©Ò®×¸¹(±q®×¥ó¸ê®Æ¤Î¶i«×¬d¸ß¨ú±o)
                stAddr(6) = .Fields("AAL04")
                ChgCaseNo stAddr(6), strExc
                If PUB_GetAddrRef("", strExc(1), strExc(2), strExc(3), strExc(4), stAddr(3), stAddr(5), stAddr(0), stAddr(1)) = True Then
                   If Len(stAddr(1)) > 20 Then
                      stAddr(2) = Mid(stAddr(1), 21)
                      stAddr(1) = Left(stAddr(1), 20)
                   End If
                   If Len(stAddr(3)) > 20 Then
                      stAddr(4) = Mid(stAddr(3), 21)
                      stAddr(3) = Left(stAddr(3), 20)
                   End If
                End If
                stAddr(6) = strExc(1) & "-" & strExc(2) & IIf(strExc(3) & strExc(4) = "000", "", "-" & strExc(3) & "-" & strExc(4))
            End If
            'end 2016/11/04
            
            For i = 0 To IntF - 1
               Printer.CurrentX = m_dbl_LeftMargin
               If IsEmpty(stAddr(i)) = False Then
                  Printer.CurrentY = i * iHeight + m_dbl_TopMargin
               End If
            
               'Added by Lydia 2022/05/02
               Xo = Printer.CurrentX
               Yo = Printer.CurrentY
               'end 2022/05/02
                     
               If IsEmptyText(stAddr(i)) = False Then
                    '­Y¬°¤¤¤å¦a§},¦C¦LÄæ¦ì¬°¤½¥q¦WºÙ©Î­Ó¤H¦WºÙ
                    strExc(0) = stAddr(i)
                    If strExc(0) = "¼B¬ì¨}¡D·Å©y¬Â" Then
                       strExc(0) = "·Å©y¬Â"
                    End If
                    If i = 3 Then
                       '­Y¬°­Ó¤H¤á®É
                       If stAddr(3) <> "" And stAddr(5) = "" Then
                           'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                           'Printer.Print strExc(0) & "¡@¡@¡@¡@¡@§g¡@¶v±Ò"
                           PUB_PrintUnicodeText strExc(0) & "¡@¡@¡@¡@¡@§g¡@¶v±Ò", Xo, Yo, 0
                       Else
                           'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                           'Printer.Print strExc(0)
                           PUB_PrintUnicodeText strExc(0), Xo, Yo, 0
                       End If
                    ElseIf i = 5 Then
                       '­Y¬°¤½¥q¤á®É
                       If stAddr(3) <> "" And stAddr(5) <> "" Then
                           'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                           'Printer.Print stAddr(i) & "¡@¡@¡@¡@¡@§g¡@¶v±Ò"
                           PUB_PrintUnicodeText stAddr(i) & "¡@¡@¡@¡@¡@§g¡@¶v±Ò", Xo, Yo, 0
                       Else
                           'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                           'Printer.Print stAddr(i)
                           PUB_PrintUnicodeText stAddr(i), Xo, Yo, 0
                       End If
                    Else
                       'Modified by Lydia 2022/05/02 ³v¦rÀË¬dUnicode¤å¦r§ï¥H¹Ï¤ù¤è¦¡¦C¦L
                       'Printer.Print stAddr(i)
                       PUB_PrintUnicodeText stAddr(i), Xo, Yo, 0
                    End If
               End If
               intBgnRow = intBgnRow + 1
            Next
            
            'Mark by Lydia 2016/11/04 ¦W±ø¤£¦L¬y¤ô¸¹
            'Printer.CurrentX = 4000 + m_dbl_LeftMargin
            'Printer.CurrentY = (intBgnRow - 1) * iHeight + m_dbl_TopMargin
            'Printer.Print Format(rsA.Fields("AAL03"), "000000")

            iPrint = iPrint + 1
            .MoveNext
         Loop

      End With
   Next
    Printer.EndDoc

    '¥i­«ÂÐ¦C¦L¦a§}±ø
    If MsgBox("A4¦W±ø¤w¦C¦L§¹²¦¡A±z¬O§_­n­«·s¦C¦L???", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
        GoTo RePrint
    Else
       'Added by Lydia 2017/11/03 «ü©w¦C¦L
        If pNoList <> "" Then
           cnnConnection.Execute " DELETE FROM ADDRESSA4LIST Where AAL01||AAL02||AAL03 IN (" & GetAddStr(pNoList) & ") "
        Else
        'end 2017/11/03
           cnnConnection.Execute "delete from AddressA4List where aal01='" & strUserNum & "' "
        End If 'end 2017/11/03
    End If

End If

    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End Function

'Added by Lydia 2017/11/03 Åª¨úA4¦a§}±øªº³sµ¸¦a§}
Private Function ReadA4List(Optional ByVal bolAll As Boolean = False) As Boolean
Dim rsRd As New ADODB.Recordset
Dim rsAD As New ADODB.Recordset
Dim intR As Integer
Dim stSQL As String
Dim SNo As String, sCustName As String, sContact As String, sZip As String, sAddr As String

    stSQL = "select aal01,aal02,aal03,aal04 from addressa4list where aal01='" & strUserNum & "' order by aal02,aal03 "
    intR = 1
    Set rsRd = ClsLawReadRstMsg(intR, stSQL)
    If intR = 1 Then
       Set rsAD = PUB_CreateRecordset(rsRd, , , , Me.Name, mESeqNo)
       cnnConnection.BeginTrans
          With rsAD
              .MoveFirst
              Do While Not .EOF
                 sCustName = ""
                 sContact = ""
                 sZip = ""
                 sAddr = ""
                 If Mid("" & .Fields(3), 1, 1) = "X" Then
                    SNo = Mid("" & .Fields(3), 1, 9)
                    strExc(1) = IIf(Mid("" & .Fields(3), 10, 1) = "-", Mid("" & .Fields(3), 11), "") '¹w³]±µ¬¢¤H
                    'Modified by Morgan 2025/1/6 ´¼Åv¦LA4¦a§}±ø®É³£­n±a¥X¦a§} Ex:X54243090
                    'If PUB_GetAddrRef(SNo, "", "", "", "", sCustName, sContact, sZip, sAddr, , strExc(1), IIf(strExc(1) <> "", True, False)) = True Then
                    If PUB_GetAddrRef(SNo, "", "", "", "", sCustName, sContact, sZip, sAddr, , strExc(1), IIf(strExc(1) <> "" Or iStiu = 1, True, False)) = True Then
                    'end 2025/1/6
                    End If
                 Else
                    SNo = "" & .Fields(3)
                    ChgCaseNo SNo, strExc
                    If PUB_GetAddrRef("", strExc(1), strExc(2), strExc(3), strExc(4), sCustName, sContact, sZip, sAddr) = True Then
                    End If
                 End If
                   'Modified by Lydia 2019/12/06 +ChgSql
                   stSQL = "update rdatafactory set r005=" & CNULL(ChgSQL(sCustName)) & ", r006=" & CNULL(ChgSQL(sAddr)) & ", r007=" & CNULL(ChgSQL(sContact)) & ", r008=" & CNULL(ChgSQL(sZip)) & _
                            " where FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno='" & mESeqNo & "' and rowseq=" & CNULL(.AbsolutePosition)
                   cnnConnection.Execute stSQL, intR
                  .MoveNext
              Loop
          End With
       cnnConnection.CommitTrans
       
       intR = 1
       '«D±µ¬¢¤H
       stSQL = "select ' ' v, substr(sqldatet(aal02),1,10) aal02, aal03,nvl(r005,r007) name1,nvl(r006,'') addr1,aal04 from addressa4list,rdatafactory " & _
               "where aal01='" & strUserNum & "'  and instr(aal04,'-') = 0 and aal01=R001(+) and to_char(aal02)=R002(+) and to_char(aal03)=R003(+) and aal04=R004(+) " & _
               "and FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno='" & mESeqNo & "' "
       '±µ¬¢¤H
       stSQL = stSQL & " union all select ' ' v, substr(sqldatet(aal02),1,10) aal02, aal03,nvl(r007,r005) name1,nvl(r006,'') addr1,aal04 from addressa4list,rdatafactory " & _
               "where aal01='" & strUserNum & "'  and instr(aal04,'-') > 0 and aal01=R001(+) and to_char(aal02)=R002(+) and to_char(aal03)=R003(+) and aal04=R004(+) " & _
               "and FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno='" & mESeqNo & "' "
       stSQL = stSQL & " order by aal02,aal03 "
       Set rsRd = ClsLawReadRstMsg(intR, stSQL)
       If intR = 1 Then
          MGrid1.FixedCols = 0
          Set MGrid1.Recordset = rsRd
          Call SetGrd(rsRd.RecordCount + 1, bolAll)
          MGrid1.FixedCols = 1
          lblCnt.Caption = rsRd.RecordCount
       Else
          MGrid1.Clear
          Call SetGrd
          lblCnt.Caption = "0"
       End If
       
       ReadA4List = True
    End If
    
    Set rsRd = Nothing
    Set rsAD = Nothing
End Function

Private Sub SetGrd(Optional ByRef iR As Integer = 2, Optional ByVal bolAll As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   Dim iCol As Integer
   arrGridHeadText = Array("v", "¤é¡@´Á", "¶¶§Ç", "¦¬ ¥ó ¤H", "¡@¦¬¡@¥ó¡@¦a¡@§}¡@", "X,Y½s¸¹/¥»©Ò®×¸¹")
   arrGridHeadWidth = Array(270, 840, 0, 1200, 2600, 1500)

   With MGrid1

       .Visible = False
       .Cols = UBound(arrGridHeadText) + 1
       .Rows = iR
       For iRow = 0 To .Cols - 1
          .row = 0
          .col = iRow
          .Text = arrGridHeadText(iRow)
          .ColWidth(iRow) = arrGridHeadWidth(iRow)
          .CellAlignment = flexAlignCenterCenter
       Next

       If bolAll = True Then
          For iRow = 1 To iR - 1
             .row = iRow
             .col = 0
             .Text = "v"
             For iCol = 0 To .Cols - 1
                .col = iCol
                .CellBackColor = &HFFC0C0
             Next iCol
          Next iRow
       End If
       .Visible = True
   End With
End Sub

Private Sub MGrid1_Click()
   Dim intRow As Integer
   If MGrid1.row > 0 Then
      intRow = MGrid1.row
      GridClick MGrid1, intRow, 0
   End If
End Sub

'A4¦a§}±ø-§R°£
Private Sub Command1_Click()
Dim ii As Integer
Dim stSQL As String

   With MGrid1
      For ii = 1 To .Rows - 1
         .row = ii
         If Trim("" & .TextMatrix(ii, 0)) <> "" Then
            stSQL = "delete from addressa4list where aal01='" & strUserNum & "' and aal02='" & TransDate(Trim(Replace("" & .TextMatrix(ii, 1), "/", "")), 2) & "' and aal03='" & Trim("" & .TextMatrix(ii, 2)) & "' "
            cnnConnection.Execute stSQL
         End If
      Next ii
   End With
   If stSQL <> "" Then
      If ReadA4List = False Then
         MsgBox "µL¸ê®Æ¥i¨Ñ¦C¦L !"
         Call Command3_Click
      End If
   End If
End Sub

'A4¦a§}±ø-¦C¦L
Private Sub Command2_Click()
Dim ii As Integer
Dim stSQL As String

   With MGrid1
      For ii = 1 To .Rows - 1
         If Trim("" & .TextMatrix(ii, 0)) <> "" Then
            stSQL = stSQL & strUserNum & TransDate(Trim(Replace("" & .TextMatrix(ii, 1), "/", "")), 2) & Trim("" & .TextMatrix(ii, 2)) & ","
         End If
      Next ii
   End With
   If stSQL <> "" Then
      PUB_RestorePrinter Me.Combo2.Text
      If PrintCaseBatchA4(stSQL) = True Then
        If ReadA4List = False Then
           Call Command3_Click
        End If
      End If
      PUB_RestorePrinter strPrinter
   End If
End Sub

'A4¦a§}±ø-µ²§ô
Private Sub Command3_Click()
    Unload Me
End Sub

'Added by Lydia 2022/05/02
Private Sub textFM2_GotFocus(Index As Integer)
    TextInverse textFM2(Index)
End Sub

Private Sub txtStartPage_GotFocus()
   TextInverse txtStartPage
End Sub

'Added by Morgan 2024/3/14
Private Sub txtStartPage_Validate(Cancel As Boolean)
   If Val(txtStartPage) < 1 Or Val(txtStartPage) > 16 Then
      MsgBox "¥u¯à¿é¤J1~16ªº¼Æ¦r¡I", vbExclamation
      Cancel = True
   End If
End Sub
