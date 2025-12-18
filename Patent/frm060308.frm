VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060308 
   BorderStyle     =   1  '≥ÊΩu©T©w
   Caption         =   "¥¡≠≠∫ﬁ®Ó™Ì"
   ClientHeight    =   5832
   ClientLeft      =   1908
   ClientTop       =   1980
   ClientWidth     =   4272
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5832
   ScaleWidth      =   4272
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   15
      Left            =   1230
      MaxLength       =   1
      TabIndex        =   16
      Text            =   "1"
      Top             =   4380
      Width           =   315
   End
   Begin VB.Frame Frame1 
      Caption         =   "≥]©w"
      Height          =   744
      Left            =   72
      TabIndex        =   34
      Top             =   4980
      Width           =   3825
      Begin VB.OptionButton Option2 
         Caption         =   "æÓ¶L"
         Height          =   180
         Index           =   1
         Left            =   2175
         TabIndex        =   20
         Top             =   492
         Width           =   1275
      End
      Begin VB.OptionButton Option2 
         Caption         =   "™Ω¶L"
         Height          =   180
         Index           =   0
         Left            =   645
         TabIndex        =   19
         Top             =   516
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.ComboBox Combo1 
         Height          =   260
         Left            =   765
         Style           =   2  '≥ÊØ¬§U©‘¶°
         TabIndex        =   18
         Top             =   180
         Width           =   2880
      End
      Begin VB.Label Label2 
         Caption         =   "¶L™Ìæ˜"
         Height          =   180
         Left            =   105
         TabIndex        =   35
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "µ≤ßÙ(&X)"
      Height          =   400
      Index           =   1
      Left            =   3315
      TabIndex        =   22
      Top             =   36
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "¶C¶L(&P)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2520
      TabIndex        =   21
      Top             =   36
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   14
      Left            =   1230
      MaxLength       =   1
      TabIndex        =   17
      Top             =   4680
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   13
      Left            =   2445
      MaxLength       =   9
      TabIndex        =   15
      Top             =   4080
      Width           =   1080
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   12
      Left            =   1230
      MaxLength       =   9
      TabIndex        =   14
      Top             =   4080
      Width           =   1080
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   2445
      MaxLength       =   9
      TabIndex        =   13
      Top             =   3780
      Width           =   1080
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   1230
      MaxLength       =   9
      TabIndex        =   12
      Top             =   3780
      Width           =   1080
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1215
      MaxLength       =   6
      TabIndex        =   11
      Top             =   3470
      Width           =   1080
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1230
      MaxLength       =   6
      TabIndex        =   10
      Top             =   3180
      Width           =   1080
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1230
      MaxLength       =   6
      TabIndex        =   9
      Top             =   2880
      Width           =   1080
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1230
      TabIndex        =   8
      Top             =   2580
      Width           =   2940
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1230
      MaxLength       =   1
      TabIndex        =   7
      Top             =   2270
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2445
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1980
      Width           =   1080
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1230
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1980
      Width           =   1080
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2445
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1680
      Width           =   1080
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1230
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1680
      Width           =   1080
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1230
      TabIndex        =   0
      Top             =   1380
      Width           =   2304
   End
   Begin VB.OptionButton Option1 
      Caption         =   "™k©w¥¡≠≠°G"
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   4
      Top             =   2010
      Width           =   1200
   End
   Begin VB.OptionButton Option1 
      Caption         =   "•ª©“¥¡≠≠°G"
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   1740
      Value           =   -1  'True
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "•N≤z§H∞Íƒy°G"
      Height          =   180
      Index           =   12
      Left            =   90
      TabIndex        =   42
      Top             =   4430
      Width           =   1130
   End
   Begin VB.Label Label1 
      Caption         =   "(1.§È•ª 2.´D§È•ª)"
      Height          =   180
      Index           =   11
      Left            =   1620
      TabIndex        =   41
      Top             =   4430
      Width           =   1550
   End
   Begin MSForms.Label lbl1 
      Height          =   200
      Index           =   2
      Left            =   2340
      TabIndex        =   40
      Top             =   3510
      Width           =   1820
      VariousPropertyBits=   27
      Size            =   "3201;344"
      FontName        =   "∑s≤”©˙≈È-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   200
      Index           =   0
      Left            =   2340
      TabIndex        =   39
      Top             =   2910
      Width           =   1820
      VariousPropertyBits=   27
      Size            =   "3201;344"
      FontName        =   "∑s≤”©˙≈È-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   200
      Index           =   1
      Left            =   2340
      TabIndex        =   38
      Top             =   3210
      Width           =   1820
      VariousPropertyBits=   27
      Size            =   "3201;344"
      FontName        =   "∑s≤”©˙≈È-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "¢±≥¯™Ì±N¶C•XFMP§uµ{Æv(•ºßπΩZ©Œ¬Ωƒ∂•ºÆ÷°@ΩZßπ¶®)§Œ©”øÏ≤’¥¡≠≠°C"
      BeginProperty Font 
         Name            =   "∑s≤”©˙≈È"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   135
      TabIndex        =   37
      Top             =   930
      Width           =   3930
   End
   Begin VB.Label Label3 
      Caption         =   "¢∞¶C¶L•ª©“¥¡≠≠±¯•Û™∫•º¶¨§Â≥¯™Ì±N•]ßt¨˘©w°@¥¡≠≠∏®©Û∏”∞œ∂°™∫ FMP Æ◊•Û°C"
      BeginProperty Font 
         Name            =   "∑s≤”©˙≈È"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   135
      TabIndex        =   36
      Top             =   480
      Width           =   3930
   End
   Begin VB.Line Line4 
      X1              =   2100
      X2              =   3072
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line3 
      X1              =   2070
      X2              =   2790
      Y1              =   3920
      Y2              =   3920
   End
   Begin VB.Line Line2 
      X1              =   1920
      X2              =   2964
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Line Line1 
      X1              =   1970
      X2              =   2906
      Y1              =   1770
      Y2              =   1770
   End
   Begin VB.Label Label1 
      Caption         =   "(1.§§§Â 2.≠^§Â 3.§È§Â)"
      Height          =   180
      Index           =   10
      Left            =   1620
      TabIndex        =   33
      Top             =   4740
      Width           =   1920
   End
   Begin VB.Label Label1 
      Caption         =   "(1.•º¶¨§Â 2.§w¶¨§Â)"
      Height          =   180
      Index           =   9
      Left            =   1620
      TabIndex        =   32
      Top             =   2310
      Width           =   1550
   End
   Begin VB.Label Label1 
      Caption         =   "≥¯™ÌÆÊ¶°°G"
      Height          =   180
      Index           =   8
      Left            =   90
      TabIndex        =   31
      Top             =   4710
      Width           =   920
   End
   Begin VB.Label Label1 
      Caption         =   "•N≤z§H°G"
      Height          =   180
      Index           =   7
      Left            =   90
      TabIndex        =   30
      Top             =   4100
      Width           =   740
   End
   Begin VB.Label Label1 
      Caption         =   "•”Ω–§H°G"
      Height          =   180
      Index           =   6
      Left            =   90
      TabIndex        =   29
      Top             =   3810
      Width           =   740
   End
   Begin VB.Label Label1 
      Caption         =   "∫ﬁ®Ó§H°G"
      Height          =   180
      Index           =   5
      Left            =   90
      TabIndex        =   28
      Top             =   3510
      Width           =   740
   End
   Begin VB.Label Label1 
      Caption         =   "©”øÏ§H°G"
      Height          =   180
      Index           =   4
      Left            =   90
      TabIndex        =   27
      Top             =   3210
      Width           =   740
   End
   Begin VB.Label Label1 
      Caption         =   "¥º≈v§H≠˚°G"
      Height          =   180
      Index           =   3
      Left            =   90
      TabIndex        =   26
      Top             =   2910
      Width           =   1010
   End
   Begin VB.Label Label1 
      Caption         =   "Æ◊•Û© ΩË°G"
      Height          =   180
      Index           =   2
      Left            =   90
      TabIndex        =   25
      Top             =   2580
      Width           =   920
   End
   Begin VB.Label Label1 
      Caption         =   "¶C¶LßO°G"
      Height          =   180
      Index           =   1
      Left            =   90
      TabIndex        =   24
      Top             =   2310
      Width           =   740
   End
   Begin VB.Label Label1 
      Caption         =   "®t≤Œ√˛ßO°G"
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   23
      Top             =   1440
      Width           =   915
   End
End
Attribute VB_Name = "frm060308"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/5 Form2.0§w≠◊ßÔ
'Memo By Morgan 2012/12/10 ¥º≈v§H≠˚ƒÊ§w≠◊ßÔ
'Modify by Morgan 2011/5/26 +≠^§Â¶aß}6 CU28->CU28||rtrim(' '||cu102),FA22->FA22||rtrim(' '||FA70)(ªP≠^§Â¶aß}5¶X®÷)
'Memo by Morgan2010/12/27 •”Ω–Æ◊∏πƒÊ§w≠◊ßÔ
'2010/12/6 memo by sonia ≠˚§uΩs∏πƒÊ§w≠◊ßÔ
'Memo by Morgan2010/8/13 §È¥¡ƒÊ§w≠◊ßÔ
Option Explicit

'≠´æ„ by Morgan 2006/2/7
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer
Dim strTemp3(0 To 9) As String, StrSQL6 As String
Dim strSQL8 As String 'Add By Sindy 2021/4/15
Dim strSQL2 As String, iPrint As Integer, Page As Integer
Dim strTemp(0 To 25) As String, PrintPage As Boolean
Dim PLeft(0 To 25) As Integer, strTemp1 As Variant, strTemp2 As Variant, Bol1 As Boolean
Dim STRSTRING As String, SeekPrint As Integer, SeekPrintL As Integer
Dim ChkPro(1 To 12) As Boolean, ChkPrintPro(1 To 12) As Boolean
Dim m_ColCustName As String '•”Ω–§H¶W∫ŸƒÊ¶Ï
Dim m_ColCustAdd As String '•”Ω–§H¶aß}ƒÊ¶Ï
Dim m_ColAgName As String '•N≤z§H¶W∫ŸƒÊ¶Ï
Dim m_ColAgAdd As String '•N≤z§H¶aß}ƒÊ¶Ï
Dim blnClkSure As Boolean 'ßP¬_¨Oß_´ˆ§UΩT©w´ˆ∂s
Dim stPA(1 To 4) As String, iPos As Integer, stCaseNo As String
Dim StrSQL3 As String, StrSQL4 As String, strSQL5 As String, strPS As String 'Add by Morgan 2009/12/8
'Add by Morgan 2011/3/15
Dim strPrinter As String
Dim m_intDefaultOri As Integer 'πw≥]¶L™Ìæ˜¶C¶L§Ë¶V
'Add by Amy 2020/03/31
Dim strCmp_C As String '§Ω•q¶W∫Ÿ-§§§Â
Dim strCmp_J As String '§Ω•q¶W∫Ÿ-§È§Â


Private Sub DoPrint()
   StrSQL3 = "": StrSQL4 = "": strSQL5 = "": strPS = "" 'Add by Morgan 2009/12/8
   blnClkSure = False
   
   'Modify by Morgan 2011/3/15 ≤æ®Ï
'         '≠Y¶≥øÈ§J•N≤z§H©Œ•”Ω–§H±¯•ÛÆ…, ¶C¶L§j±i≥¯™Ì
'         If Me.txt1(10).Text <> "" Or Me.txt1(11).Text <> "" Or Me.txt1(12).Text <> "" Or Me.txt1(13).Text <> "" Then
'           Set Printer = Printers(Combo1.ListIndex)
'           If Option2(0).Value = True Then
'              Printer.Orientation = 1
'           Else
'              Printer.Orientation = 2
'           End If
'         '≠Y•ºøÈ§J•N≤z§H§Œ•”Ω–§H±¯•ÛÆ…, ¶C¶LæÓ¶VA4≥¯™Ì
'         Else
'           Set Printer = Printers(SeekPrint)
'           Printer.Orientation = 2
'         End If
      
      Printer.EndDoc
      If Option2(0).Value = True Then
         Printer.Orientation = 1
      Else
         Printer.Orientation = 2
      End If
   'end 2011/3/15
     DoEvents
     
      'Add By Sindy 2023/5/22
      If txt1(5) = "1" Then '1.•º¶¨§Â
         If Len(txt1(15)) = 0 Then
            s = MsgBox("•N≤z§H∞Íƒy§£•i™≈•’!!", , "USER øÈ§Jø˘ª~")
            txt1(15).SetFocus
            Exit Sub
         End If
      End If
      '2023/5/22 END
     
     If Len(txt1(0)) = 0 Then
         s = MsgBox("®t≤ŒßO§£•i™≈•’!!", , "USER øÈ§Jø˘ª~")
         txt1(0).SetFocus
         Exit Sub
     Else
         If Option1(0).Value = True Then
            If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
               Me.txt1(1).SetFocus
               txt1_GotFocus 1
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
               Me.txt1(2).SetFocus
               txt1_GotFocus 2
               Exit Sub
            End If
            If Me.txt1(1).Text <> "" And Me.txt1(2).Text <> "" Then
               If Val(Me.txt1(1).Text) > Val(Me.txt1(2).Text) Then
                  MsgBox "•ª©“¥¡≠≠Ωd≥ÚøÈ§Jø˘ª~!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(1).SetFocus
                  txt1_GotFocus 1
                  Exit Sub
               End If
            End If
            
            If Len(txt1(2)) = 0 Then
                s = MsgBox("•ª©“¥¡≠≠∞œ∂°§£•i™≈•’!!", , "USER øÈ§Jø˘ª~")
                If Len(txt1(1)) = 0 Then txt1(1).SetFocus
                Exit Sub
            Else
                If Len(txt1(5)) = 0 Then
                    s = MsgBox("¶C¶LßO§£•i™≈•’!!", , "USER øÈ§Jø˘ª~")
                    txt1(5).SetFocus
                    Exit Sub
                Else
                    If Len(txt1(14)) = 0 Then
                        s = MsgBox("≥¯™ÌÆÊ¶°§£•i™≈•’!!", , "USER øÈ§Jø˘ª~")
                        txt1(14).SetFocus
                        Exit Sub
                    Else
                        lbl1(0).Caption = GetPrjSales(txt1(7), "¥º≈v§H≠˚")
                        If Me.txt1(7).Text <> "" Then
                           If Me.txt1(7).Text = Me.lbl1(0).Caption Then
                              Me.lbl1(0).Caption = ""
                              Me.txt1(7).SetFocus
                              txt1_GotFocus 7
                              Exit Sub
                           End If
                        End If
                        lbl1(1).Caption = GetPrjSales(txt1(8))
                        If Me.txt1(8).Text <> "" Then
                           If Me.txt1(8).Text = Me.lbl1(1).Caption Then
                              Me.lbl1(1).Caption = ""
                              Me.txt1(8).SetFocus
                              txt1_GotFocus 8
                              Exit Sub
                           End If
                        End If
                        lbl1(2).Caption = GetPrjSales(txt1(9), "∫ﬁ®Ó§H")
                        If Me.txt1(9).Text <> "" Then
                           If Me.txt1(9).Text = Me.lbl1(2).Caption Then
                              Me.lbl1(2).Caption = ""
                              Me.txt1(9).SetFocus
                              txt1_GotFocus 9
                              Exit Sub
                           End If
                        End If
                        
                        If Len(Trim(txt1(10))) <> 0 Or Len(Trim(txt1(11))) <> 0 Then
                            If Left(txt1(10), 6) <> Left(txt1(11), 6) Then
                                s = MsgBox("•”Ω–§H´e 6 ΩX•≤∂∑¨€¶P", , "USER øÈ§Jø˘ª~")
                                blnClkSure = True
                                If Len(Trim(txt1(10))) = 0 Then txt1(10).SetFocus
                                Me.txt1(10).SetFocus
                                Exit Sub
                            End If
                        End If
                        If Me.txt1(10).Text <> "" And Me.txt1(11).Text <> "" Then
                           If Me.txt1(10).Text > Me.txt1(11).Text Then
                              MsgBox "•”Ω–§HΩd≥ÚøÈ§Jø˘ª~!!!", vbExclamation + vbOKOnly
                              blnClkSure = True
                              Me.txt1(10).SetFocus
                              txt1_GotFocus 10
                              Exit Sub
                           End If
                        End If
                        If Len(Trim(txt1(12))) <> 0 Or Len(Trim(txt1(13))) <> 0 Then
                            If Left(txt1(12), 6) <> Left(txt1(13), 6) Then
                                s = MsgBox("•N≤z§H´e 6 ΩX•≤∂∑¨€¶P", , "USER øÈ§Jø˘ª~")
                                blnClkSure = True
                                If Len(Trim(txt1(12))) = 0 Then txt1(12).SetFocus
                                Exit Sub
                            End If
                        End If
                        If Me.txt1(12).Text <> "" And Me.txt1(13).Text <> "" Then
                           If Me.txt1(12).Text > Me.txt1(13).Text Then
                              MsgBox "•N≤z§HΩd≥ÚøÈ§Jø˘ª~!!!", vbExclamation + vbOKOnly
                              blnClkSure = True
                              Me.txt1(12).SetFocus
                              txt1_GotFocus 12
                              Exit Sub
                           End If
                        End If
                        
                        Screen.MousePointer = vbHourglass
                        Me.Enabled = False
                        StrMenu
                        Me.Enabled = True
                        Screen.MousePointer = vbDefault
                    End If
                End If
            End If
         Else
            If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
               Me.txt1(3).SetFocus
               txt1_GotFocus 3
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(4)) = -1 Then
               Me.txt1(4).SetFocus
               txt1_GotFocus 4
               Exit Sub
            End If
            If Me.txt1(3).Text <> "" And Me.txt1(4).Text <> "" Then
               If Val(Me.txt1(3).Text) > Val(Me.txt1(4).Text) Then
                  MsgBox "™k©w¥¡≠≠Ωd≥ÚøÈ§Jø˘ª~!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(3).SetFocus
                  txt1_GotFocus 3
                  Exit Sub
               End If
            End If
            
            If Len(txt1(4)) = 0 Then
                s = MsgBox("™k©w¥¡≠≠∞œ∂°§£•i™≈•’!!", , "USER øÈ§Jø˘ª~")
                If Len(txt1(3)) = 0 Then txt1(3).SetFocus
                Exit Sub
            Else
                If Len(txt1(5)) = 0 Then
                    s = MsgBox("¶C¶LßO§£•i™≈•’!!", , "USER øÈ§Jø˘ª~")
                    txt1(5).SetFocus
                    Exit Sub
                Else
                    If Len(txt1(14)) = 0 Then
                        s = MsgBox("≥¯™ÌÆÊ¶°§£•i™≈•’!!", , "USER øÈ§Jø˘ª~")
                        txt1(14).SetFocus
                        Exit Sub
                    Else
                        lbl1(0).Caption = GetPrjSales(txt1(7))
                        If Me.txt1(7).Text <> "" Then
                           If Me.txt1(7).Text = Me.lbl1(0).Caption Then
                              Me.txt1(7).SetFocus
                              txt1_GotFocus 7
                              Exit Sub
                           End If
                        End If
                        lbl1(1).Caption = GetPrjSales(txt1(8))
                        If Me.txt1(8).Text <> "" Then
                           If Me.txt1(8).Text = Me.lbl1(1).Caption Then
                              Me.txt1(8).SetFocus
                              txt1_GotFocus 8
                              Exit Sub
                           End If
                        End If
                        lbl1(2).Caption = GetPrjSales(txt1(9))
                        If Me.txt1(9).Text <> "" Then
                           If Me.txt1(9).Text = Me.lbl1(2).Caption Then
                              Me.txt1(9).SetFocus
                              txt1_GotFocus 9
                              Exit Sub
                           End If
                        End If
                        If Len(Trim(txt1(10))) <> 0 Or Len(Trim(txt1(11))) <> 0 Then
                            If Left(txt1(10), 6) <> Left(txt1(11), 6) Then
                                s = MsgBox("•”Ω–§H´e 6 ΩX•≤∂∑¨€¶P", , "USER øÈ§Jø˘ª~")
                                blnClkSure = True
                                If Len(Trim(txt1(10))) = 0 Then txt1(10).SetFocus
                                Exit Sub
                            End If
                        End If
                        If Me.txt1(10).Text <> "" And Me.txt1(11).Text <> "" Then
                           If Me.txt1(10).Text > Me.txt1(11).Text Then
                              MsgBox "•”Ω–§HΩd≥ÚøÈ§Jø˘ª~!!!", vbExclamation + vbOKOnly
                              blnClkSure = True
                              Me.txt1(10).SetFocus
                              txt1_GotFocus 10
                              Exit Sub
                           End If
                        End If
                        If Len(Trim(txt1(12))) <> 0 Or Len(Trim(txt1(13))) <> 0 Then
                            If Left(txt1(12), 6) <> Left(txt1(13), 6) Then
                                s = MsgBox("•N≤z§H´e 6 ΩX•≤∂∑¨€¶P", , "USER øÈ§Jø˘ª~")
                                blnClkSure = True
                                If Len(Trim(txt1(12))) = 0 Then txt1(12).SetFocus
                                Exit Sub
                            End If
                        End If
                        If Me.txt1(12).Text <> "" And Me.txt1(13).Text <> "" Then
                           If Me.txt1(12).Text > Me.txt1(13).Text Then
                              MsgBox "•N≤z§HΩd≥ÚøÈ§Jø˘ª~!!!", vbExclamation + vbOKOnly
                              blnClkSure = True
                              Me.txt1(12).SetFocus
                              txt1_GotFocus 12
                              Exit Sub
                           End If
                        End If
                        
                        Screen.MousePointer = vbHourglass
                        Me.Enabled = False
                        StrMenu
                        Me.Enabled = True
                        Screen.MousePointer = vbDefault
                    End If
                End If
            End If
         End If
     End If
End Sub

Private Sub cmdok_Click(Index As Integer)
   Me.Tag = "" 'πw≥]≠» Add By Sindy 2023/5/23
   Select Case Index
      Case 0 '¶C¶L
         PUB_RestorePrinter Combo1
         'Modify By Sindy 2023/10/27 EP33≠n¶^¬k•Œ¶b≠^§ÂÆ÷ßπ§È,ßÔßÏEP39.Æ÷ΩZßπ¶®§È
         DoPrint
         PUB_RestorePrinter strPrinter, m_intDefaultOri
      Case 1 'µ≤ßÙ
           Unload Me
      Case Else
   End Select
End Sub

Sub StrMenu()
   cnnConnection.Execute "DELETE FROM R060308_1 WHERE ID='" & strUserNum & "' "
   cnnConnection.Execute "DELETE FROM R060308_2 WHERE ID='" & strUserNum & "' "
   cnnConnection.Execute "DELETE FROM R060308_3 WHERE ID='" & strUserNum & "' "
   cnnConnection.Execute "DELETE FROM R060308_4 WHERE ID='" & strUserNum & "' "
   cnnConnection.Execute "DELETE FROM R060308_5 WHERE ID='" & strUserNum & "' "
   cnnConnection.Execute "DELETE FROM R060308_6 WHERE ID='" & strUserNum & "' "
   cnnConnection.Execute "DELETE FROM R060308_7 WHERE ID='" & strUserNum & "' "
   Erase strTemp 'Added by Lydia 2019/06/24
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/8 ≤M∞£¨d∏ﬂ¶L™Ì∞Oø˝¿…ƒÊ¶Ï
   Select Case Val(txt1(5))
   Case 1
      pub_QL05 = pub_QL05 & ";" & Label1(1) & "1.•º¶¨§Â" 'Add By Sindy 2010/12/8
      Process1    '•º¶¨§Â
   Case 2
      pub_QL05 = pub_QL05 & ";" & Label1(1) & "2.§w¶¨§Â" 'Add By Sindy 2010/12/8
      Process2    '§w¶¨§Â
   Case Else
   End Select
End Sub

'Modify by Morgan 2010/1/11 +¨˘©w¥¡≠≠
Sub GetPleft1_A4()
Dim intP As Integer 'Added by Lydia 2019/06/24

   Erase PLeft
   intP = 0
   PLeft(0) = 500 '•ª©“¥¡≠≠ 1000
   'Modified by Lydia 2019/06/24
'   PLeft(1) = PLeft(0) + 1000 '™k©w¥¡≠≠ 1000
'   PLeft(12) = PLeft(1) + 1000 '¨˘©w¥¡≠≠ 1000
'   PLeft(2) = PLeft(12) + 1000 '•ª©“Æ◊∏π 1700
'   PLeft(3) = PLeft(2) + 1700 'Æ◊•Û¶W∫Ÿ 1000
'   PLeft(4) = PLeft(3) + 1000 '§U§@µ{ß« 1000
'   PLeft(5) = PLeft(4) + 2000  '¥º≈v§H≠˚ 1000
'   PLeft(6) = PLeft(5) + 900 '•N≤z§H 1100
'   PLeft(7) = PLeft(6) + 1200 '•N≤z§H∞Íƒy 1600
'   PLeft(8) = PLeft(7) + 1200 'ª‚√“¶€∞ •N√∫ 800
'   PLeft(9) = PLeft(8) + 800 '¶~∂O¶€∞ •N√∫ 800
'   PLeft(10) = PLeft(9) + 800 'Æ÷ΩZ§H 1000
'   PLeft(11) = PLeft(10) + 1000 '≥∆µ˘
   If txt1(6) <> "202" Then 'Added by Lydia 2020/05/18
      intP = intP + 1
      PLeft(intP) = PLeft(intP - 1) + 1000 '™k©w¥¡≠≠
      intP = intP + 1
      PLeft(intP) = PLeft(intP - 1) + 1000 '¨˘©w¥¡≠≠
      'Add By Sindy 2023/5/29
      If txt1(5) = "1" Then '•º¶¨§Â
         intP = intP + 1
         PLeft(intP) = PLeft(intP - 1) + 1000 'µo§Â§È
      End If
      '2023/5/29 END
      intP = intP + 1
      PLeft(intP) = PLeft(intP - 1) + 1000 '•ª©“Æ◊∏π
      'Modify By Sindy 2023/5/29
      If txt1(5) = "1" Then '•º¶¨§Â
         intP = intP + 1
         PLeft(intP) = PLeft(intP - 1) + 1700 '§U§@µ{ß«
      Else
         intP = intP + 1
         PLeft(intP) = PLeft(intP - 1) + 1700 'Æ◊•Û¶W∫Ÿ
         intP = intP + 1
         PLeft(intP) = PLeft(intP - 1) + 1000 '§U§@µ{ß«
      End If
      '2023/5/29 END
      intP = intP + 1
      PLeft(intP) = PLeft(intP - 1) + 2000 '¥º≈v§H≠˚
      intP = intP + 1
      PLeft(intP) = PLeft(intP - 1) + 900 '•N≤z§H
      intP = intP + 1
      PLeft(intP) = PLeft(intP - 1) + 1200 '•N≤z§H∞Íƒy
      intP = intP + 1
      PLeft(intP) = PLeft(intP - 1) + 1200 'ª‚√“¶€∞ •N√∫
      intP = intP + 1
      PLeft(intP) = PLeft(intP - 1) + 800 '¶~∂O¶€∞ •N√∫
      intP = intP + 1
      PLeft(intP) = PLeft(intP - 1) + 800 'πÍºf¶€∞ •N√∫
      intP = intP + 1
      'Modify By Sindy 2023/5/29
      If txt1(5) = "1" Then '•º¶¨§Â
         PLeft(intP) = PLeft(intP - 1) + 800 '≥∆µ˘
      Else
      '2023/5/29 END
         PLeft(intP) = PLeft(intP - 1) + 800 'Æ÷ΩZ§H
      End If
      'Modify By Sindy 2023/5/29
      If txt1(5) = "2" Then '§w¶¨§Â
         intP = intP + 1
         PLeft(intP) = PLeft(intP - 1) + 800 '≥∆µ˘
      End If
   'Added by Lydia 2020/05/18 ∏…§Â•Û¥¡≠≠∫ﬁ®Ó™Ì
   Else
      intP = intP + 1
      PLeft(intP) = PLeft(intP - 1) + 1100 '™k©w¥¡≠≠
      intP = intP + 1
      PLeft(intP) = PLeft(intP - 1) + 1100 '¨˘©w¥¡≠≠
      'Add By Sindy 2023/5/29
      If txt1(5) = "1" Then '•º¶¨§Â
         intP = intP + 1
         PLeft(intP) = PLeft(intP - 1) + 1100 'µo§Â§È
      End If
      '2023/5/29 END
      intP = intP + 1
      PLeft(intP) = PLeft(intP - 1) + 1100 '•ª©“Æ◊∏π
      'Modify By Sindy 2023/5/29
      If txt1(5) = "1" Then '•º¶¨§Â
         intP = intP + 1
         PLeft(intP) = PLeft(intP - 1) + 1700 '§U§@µ{ß«
      Else
         intP = intP + 1
         PLeft(intP) = PLeft(intP - 1) + 1700 'Æ◊•Û¶W∫Ÿ
         intP = intP + 1
         PLeft(intP) = PLeft(intP - 1) + 2100 '§U§@µ{ß«
      End If
      '2023/5/29 END
      intP = intP + 1
      PLeft(intP) = PLeft(intP - 1) + 2000 '¥º≈v§H≠˚
      intP = intP + 1
      PLeft(intP) = PLeft(intP - 1) + 900 '•N≤z§H
      intP = intP + 1
      PLeft(intP) = PLeft(intP - 1) + 1200 '•N≤z§H∞Íƒy
      intP = intP + 1
      PLeft(intP) = PLeft(intP - 1) + 1200 'ª‚√“¶€∞ •N√∫
      intP = intP + 1
      PLeft(intP) = PLeft(intP - 1) + 800 '¶~∂O¶€∞ •N√∫
      intP = intP + 1
      PLeft(intP) = PLeft(intP - 1) + 800 'πÍºf¶€∞ •N√∫
      'Add By Sindy 2023/5/29
      If txt1(5) = "1" Then '•º¶¨§Â
         intP = intP + 1
         PLeft(intP) = 500 '≥∆µ˘
      Else
      '2023/5/29 END
         intP = intP + 1
         PLeft(intP) = PLeft(intP - 1) + 800 'Æ÷ΩZ§H
         intP = intP + 1
         PLeft(intP) = 500 '≥∆µ˘
      End If
   End If
   'end 2020/05/18
End Sub

Sub GetPleft3_A4()
   Erase PLeft
   PLeft(0) = 500 '•ª©“Æ◊∏π
   PLeft(1) = 1500 '™k©w¥¡≠≠
   PLeft(2) = 2500 '•ª©“Æ◊∏π
   PLeft(3) = 4200 'Æ◊•Û¶W∫Ÿ
   PLeft(4) = 7500 - 800 'Æ◊•Û© ΩË
   PLeft(5) = 8500 - 800 '•N≤z§H
   PLeft(6) = 9500 '•N≤z§H∞Íƒy
   PLeft(7) = 11000 'Æ÷ΩZ§H
   PLeft(8) = 13000 '∂i´◊≥∆µ˘
   PLeft(9) = 11850 'ßπΩZ§È
End Sub

Sub GetPleft4()
   Erase PLeft
   PLeft(0) = 500      '¥¡≠≠
   PLeft(1) = 1500 + 200     '§U§@µ{ß«
   PLeft(2) = 1500 + 4500 + 2300 '∂Q©“Æ◊∏π
   PLeft(3) = 1500 + 9000 - 1000 + 2000 'case no
   PLeft(4) = 17000 - 2000   '•ª©“Æ◊∏π
   PLeft(5) = 8500 - 300 + 400      'Æ◊•Û¶W∫Ÿ
   PLeft(6) = 11000 + 950 + 400     '•”Ω–§H
   PLeft(7) = 15000 + 400 + 2000    '•”Ω–Æ◊∏π
   PLeft(8) = 17000 + 400        '±MßQ∏π
End Sub

Sub GetPleft5()
   Erase PLeft
   PLeft(0) = 500    '¥¡≠≠
   PLeft(1) = 1500 + 200    '§U§@µ{ß«
   PLeft(2) = 1500 + 4500 + 2300 '∂Q©“Æ◊∏π
   PLeft(3) = 1500 + 9000 + 1000 'case no
   PLeft(4) = 17000 - 2000 '•ª©“Æ◊∏π
   PLeft(5) = 8500 + 400    'Æ◊•Û¶W∫Ÿ
   PLeft(6) = 15000 + 400 + 2000 '•”Ω–Æ◊∏π
   PLeft(7) = 17000 + 400    '±MßQ∏π
   PLeft(8) = 11000 + 950 + 400 '•”Ω–§H
End Sub

Sub GetPleft6()
   Erase PLeft
   PLeft(0) = 500       '¥¡≠≠
   PLeft(1) = 1700     '∂Q©“Æ◊∏π
   PLeft(2) = 1700 + 4500   'case no
   PLeft(3) = 10500    '•ª©“Æ◊∏π
   PLeft(4) = 6000 - 150 + 400     'Æ◊•Û¶W∫Ÿ
   PLeft(5) = 10500 + 1000 + 200 + 400 - 3000   '•”Ω–§H
   PLeft(6) = 15000 + 400 + 2000     '•”Ω–Æ◊∏π
   PLeft(7) = 17000 + 400     '±MßQ∏π
End Sub

Sub GetPleft7()
   Erase PLeft
   PLeft(0) = 500    '¥¡≠≠
   PLeft(1) = 1500 + 200     '∂Q©“Æ◊∏π
   PLeft(2) = 1700 + 4500   'case no
   PLeft(3) = 10500    '•ª©“Æ◊∏π
   PLeft(4) = 6000 + 400 - 150   'Æ◊•Û¶W∫Ÿ
   PLeft(5) = 15000 + 400 + 2000 '•”Ω–Æ◊∏π
   PLeft(6) = 17000 + 400    '±MßQ∏π
   PLeft(7) = 10500 + 1000 + 600 - 3000 '•”Ω–§H
End Sub
'Modify by Morgan 2010/1/11 +¨˘©w¥¡≠≠
Sub GetPleft1_416() 'πÍ≈Èºf¨d∫ﬁ®Ó™Ì
   Erase PLeft
   PLeft(0) = 500 '•ª©“¥¡≠≠ 1000
   PLeft(1) = PLeft(0) + 1000 '™k©w¥¡≠≠ 1000
   PLeft(2) = PLeft(1) + 1000 '¨˘©w¥¡≠≠ 1000
   'Modify By Sindy 2023/5/29
   If Me.txt1(5) = "1" Then '•º¶¨§Â
      PLeft(4) = PLeft(2) + 1000 '•ª©“Æ◊∏π 1700
      PLeft(8) = PLeft(4) + 1700  '•N≤z§H∞Íƒy 1800
      PLeft(12) = PLeft(8) + 1800 '≥∆µ˘
   Else
      PLeft(3) = PLeft(2) + 1000 '•ª©“Æ◊∏π 1700
      PLeft(8) = PLeft(3) + 1700  '•N≤z§H∞Íƒy 1800
      PLeft(13) = PLeft(8) + 1800 '≥∆µ˘
   End If
   '2023/5/29 END
End Sub

Sub PrintDatil1_A4()
Dim intTotCol As Integer
Dim intNoteCol As Integer
   'Modified by Lydia 2019/06/24 ¶bPrintPro1_A4§w≥B≤z
'   strTemp(3) = StrToStr(strTemp(3), 4)
'   strTemp(4) = strTemp(4) 'StrToStr(strTemp(4), 4)
'   strTemp(5) = StrToStr(strTemp(5), 4)
'   strTemp(6) = StrToStr(strTemp(6), 5)
'   strTemp(7) = StrToStr(strTemp(7), 7)
'   strTemp(10) = StrToStr(strTemp(10), 4)
'   strTemp(11) = StrToStr(strTemp(11), 5)
   
   'Add By Sindy 2023/5/23
   If txt1(5) = "1" Then '•º¶¨§Â
      intTotCol = 12
      intNoteCol = 12
   Else
      intTotCol = 13
      intNoteCol = 13
   End If
   For i = 0 To intTotCol
   '2023/5/23 END
       'Added by Lydia 2020/05/18 +ßP¬_"∏…§Â•Û¥¡≠≠∫ﬁ®Ó™Ì"=>¥´¶Ê
       If txt1(6) = "202" And i = intNoteCol Then
           If strTemp(i) <> "" Then
               iPrint = iPrint + 300
               Printer.CurrentX = PLeft(i)
               Printer.CurrentY = iPrint
               Printer.Print "≥∆µ˘°G" & strTemp(i)
           End If
       Else
       'end 2020/05/18
           Printer.CurrentX = PLeft(i)
           Printer.CurrentY = iPrint
           Printer.Print strTemp(i)
       End If 'Added by Lydia 2020/05/18
   Next i
   
   iPrint = iPrint + 300
End Sub

'cancel by sonia 2014/11/19 §£•Œ§F
Sub PrintDatil3()
   strTemp(3) = StrToStr(strTemp(3), 8)
   strTemp(4) = StrToStr(strTemp(4), 6)
   strTemp(5) = StrToStr(strTemp(5), 4)
   strTemp(6) = StrToStr(strTemp(6), 4)
   strTemp(7) = StrToStr(strTemp(7), 6)
   strTemp(8) = StrToStr(strTemp(8), 12)
   strTemp(9) = StrToStr(strTemp(9), 8)
   strTemp(10) = StrToStr(strTemp(10), 8)
   strTemp(11) = StrToStr(strTemp(11), 9)
   For i = 0 To 11
       Printer.CurrentX = PLeft(i)
       Printer.CurrentY = iPrint
       Printer.Print strTemp(i)
   Next i
   iPrint = iPrint + 300
End Sub

Sub PrintDatil3_A4()
   strTemp(3) = StrToStr(strTemp(3), 8)
   strTemp(4) = StrToStr(strTemp(4), 4)
   strTemp(5) = StrToStr(strTemp(5), 6)
   strTemp(6) = StrToStr(strTemp(6), 8)
   strTemp(7) = StrToStr(strTemp(7), 8)
   strTemp(8) = StrToStr(strTemp(8), 5)
   'Modify by Morgan 2007/3/27 •[ßπΩZ§È
   'For i = 0 To 8
   For i = 0 To 9
       Printer.CurrentX = PLeft(i)
       Printer.CurrentY = iPrint
       Printer.Print strTemp(i)
   Next i
   iPrint = iPrint + 300
End Sub

Sub PrintDatil4()
   strTemp(0) = StrToStr(strTemp(0), 5.5)
   strTemp(1) = StrToStr(strTemp(1), 32)
   strTemp(2) = StrToStr(strTemp(2), 15)
   strTemp(3) = StrToStr(strTemp(3), 15)
   strTemp(5) = StrToStr(strTemp(5), 30)
   strTemp(6) = StrToStr(strTemp(6), 41)
   strTemp(7) = StrToStr(strTemp(7), 5)
   strTemp(8) = StrToStr(strTemp(8), 4)
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(0)
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(1)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(2)
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(3)
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(4)
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(7)
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(1) + 1000
   Printer.CurrentY = iPrint
   Printer.Print strTemp(5)
   Printer.CurrentX = 9000
   Printer.CurrentY = iPrint
   Printer.Print strTemp(6)
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(8)
   iPrint = iPrint + 300
End Sub

Sub PrintDatil5()
   strTemp(0) = StrToStr(strTemp(0), 5.5)
   strTemp(1) = StrToStr(strTemp(1), 32)
   strTemp(2) = StrToStr(strTemp(2), 15)
   strTemp(3) = StrToStr(strTemp(3), 15)
   strTemp(5) = StrToStr(strTemp(5), 30)
   strTemp(6) = StrToStr(strTemp(6), 5)
   strTemp(7) = StrToStr(strTemp(7), 4)
   strTemp(8) = StrToStr(strTemp(8), 41)
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(0)
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(1)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(2)
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(3)
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(4)
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(6)
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(1) + 1000
   Printer.CurrentY = iPrint
   Printer.Print strTemp(5)
   Printer.CurrentX = 9000
   Printer.CurrentY = iPrint
   Printer.Print strTemp(8)
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(7)
   iPrint = iPrint + 300
End Sub

Sub PrintDatil6()
   strTemp(0) = StrToStr(strTemp(0), 5.5)
   strTemp(1) = StrToStr(strTemp(1), 21)
   strTemp(2) = StrToStr(strTemp(2), 21)
   strTemp(4) = StrToStr(strTemp(4), 30)
   strTemp(5) = StrToStr(strTemp(5), 41)
   strTemp(6) = StrToStr(strTemp(6), 5)
   strTemp(7) = StrToStr(strTemp(7), 4)
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(0)
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(1)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(2)
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(3)
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(6)
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(1) + 1000
   Printer.CurrentY = iPrint
   Printer.Print strTemp(4)
   Printer.CurrentX = 9000
   Printer.CurrentY = iPrint
   Printer.Print strTemp(5)
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(7)
   iPrint = iPrint + 300
End Sub

Sub PrintDatil7()
   strTemp(0) = StrToStr(strTemp(0), 5.5)
   strTemp(1) = StrToStr(strTemp(1), 21)
   strTemp(2) = StrToStr(strTemp(2), 21)
   strTemp(4) = StrToStr(strTemp(4), 30)
   strTemp(5) = StrToStr(strTemp(5), 5)
   strTemp(6) = StrToStr(strTemp(6), 4)
   strTemp(7) = StrToStr(strTemp(7), 41)
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(0)
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(1)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(2)
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(3)
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(5)
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(1) + 1000
   Printer.CurrentY = iPrint
   Printer.Print strTemp(4)
   Printer.CurrentX = 9000
   Printer.CurrentY = iPrint
   Printer.Print strTemp(7)
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(6)
   iPrint = iPrint + 300
End Sub
'Add By Morgan 2004/12/28 πÍ≈Èºf¨d¥¡≠≠∫ﬁ®Ó™Ì
Sub PrintDatil1_416()
   'Modify By Sindy 2023/5/29
   If Me.txt1(5) = "1" Then '•º¶¨§Â
      For i = 0 To 12
         Select Case i
            Case 0, 1, 2, 4, 8, 12
               If i = 8 Then strTemp(i) = StrToStr(strTemp(i), 7)
               If i = 12 Then strTemp(i) = StrToStr(strTemp(i), 5)
               Printer.CurrentX = PLeft(i)
               Printer.CurrentY = iPrint
               Printer.Print strTemp(i)
         End Select
      Next i
   Else
      For i = 0 To 13
         Select Case i
            Case 0, 1, 2, 3, 8, 13
               If i = 8 Then strTemp(i) = StrToStr(strTemp(i), 7)
               If i = 13 Then strTemp(i) = StrToStr(strTemp(i), 5)
   '2023/5/29 END
               Printer.CurrentX = PLeft(i)
               Printer.CurrentY = iPrint
               Printer.Print strTemp(i)
         End Select
      Next i
   End If
   '2023/5/29 END
   iPrint = iPrint + 300
End Sub

'Modify By Sindy 2023/5/23 Optional ByVal strRptType As String = "1" : ≥¯™Ì§¡≠∂•H 1.∫ﬁ®Ó§H 2.¥º≈v§H≠˚
Sub PrintTitle1_A4(Optional ByVal strRptType As String = "1")
Dim intCol As Integer
   
   GetPleft1_A4
   iPrint = 500
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.Font.Name = "≤”©˙≈È"
   If txt1(6) <> "202" Then 'Added by Lydia 2020/05/18
        If Option1(0).Value = True Then
           Printer.CurrentX = 8000 - (Printer.TextWidth(GetTitleNick & "•ª©“¥¡≠≠∫ﬁ®Ó™Ì") / 2)
           Printer.CurrentY = iPrint
           Printer.Print GetTitleNick & "•ª©“¥¡≠≠∫ﬁ®Ó™Ì"
        Else
           Printer.CurrentX = 8000 - (Printer.TextWidth(GetTitleNick & "™k©w¥¡≠≠∫ﬁ®Ó™Ì") / 2)
           Printer.CurrentY = iPrint
           Printer.Print GetTitleNick & "™k©w¥¡≠≠∫ﬁ®Ó™Ì"
        End If
   'Added by Lydia 2020/05/18 ∏…§Â•Û¥¡≠≠∫ﬁ®Ó™Ì
   Else
        If Option1(0).Value = True Then
           Printer.CurrentX = 8000 - (Printer.TextWidth(GetTitleNick & "∏…§Â•Û•ª©“¥¡≠≠∫ﬁ®Ó™Ì") / 2)
           Printer.CurrentY = iPrint
           Printer.Print GetTitleNick & "∏…§Â•Û•ª©“¥¡≠≠∫ﬁ®Ó™Ì"
        Else
           Printer.CurrentX = 8000 - (Printer.TextWidth(GetTitleNick & "∏…§Â•Û™k©w¥¡≠≠∫ﬁ®Ó™Ì") / 2)
           Printer.CurrentY = iPrint
           Printer.Print GetTitleNick & "∏…§Â•Û™k©w¥¡≠≠∫ﬁ®Ó™Ì"
        End If
   End If
   'end 2020/05/18
   
   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   If Option1(0).Value = True Then
      Printer.CurrentX = 8000 - (Printer.TextWidth("•ª©“¥¡≠≠°G" & ChangeTStringToTDateString(txt1(1)) & "°–" & ChangeTStringToTDateString(txt1(2))) / 2)
      Printer.CurrentY = iPrint
      Printer.Print "•ª©“¥¡≠≠°G" & ChangeTStringToTDateString(txt1(1)) & "°–" & ChangeTStringToTDateString(txt1(2))
   Else
      Printer.CurrentX = 8000 - (Printer.TextWidth("™k©w¥¡≠≠°G" & ChangeTStringToTDateString(txt1(3)) & "°–" & ChangeTStringToTDateString(txt1(4))) / 2)
      Printer.CurrentY = iPrint
      Printer.Print "™k©w¥¡≠≠°G" & ChangeTStringToTDateString(txt1(3)) & "°–" & ChangeTStringToTDateString(txt1(4))
   End If
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "¶C¶L§H°G" & strUserName
   Printer.CurrentX = 13600
   Printer.CurrentY = iPrint
   Printer.Print "¶C¶L§È¥¡°G" & Format(GetTaiwanTodayDate, "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   'Modify By Sindy 2023/5/23
   If strRptType = "2" Then
      Printer.Print "¥º≈v§H≠˚°G" & GetPrjSalesNM(strTemp3(0))
   Else
   '2023/5/23 END
      Printer.Print "∫ﬁ®Ó§H°G" & GetPrjSalesNM(strTemp3(0))
   End If
   
   'Add by Morgan 2009/12/8
   If strPS <> "" Then
      Printer.CurrentX = 8000 - Printer.TextWidth(strPS) / 2
      Printer.CurrentY = iPrint
      Printer.Print strPS
   End If
   
   Printer.CurrentX = 13600
   Printer.CurrentY = iPrint
   Printer.Print "≠∂°@°@¶∏°G" & str(Page)
   iPrint = iPrint + 300
   Printer.Font.Size = 10
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(170, "-")
   
   'Add by Morgan 2010/1/11 +¨˘©w¥¡≠≠
   iPrint = iPrint + 300
   'Modified by Lydia 2019/06/24 PLeft(8)=>PLeft(9)
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = iPrint
   Printer.Print "ª‚√“¶€"
   'Modified by Lydia 2019/06/24 PLeft(9)=>PLeft(10)
   Printer.CurrentX = PLeft(10)
   Printer.CurrentY = iPrint
   Printer.Print "¶~∂O¶€"
   'end 2010/1/11
   'Added by Lydia 2019/06/24 πÍºf¶€∞ •N√∫
   Printer.CurrentX = PLeft(11)
   Printer.CurrentY = iPrint
   Printer.Print "πÍºf¶€"
   
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "•ª©“¥¡≠≠"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "™k©w¥¡≠≠"
   
   'Add by Morgan 2010/1/11
   'Modified by Lydia 2019/06/24 PLeft(12)=>PLeft(2)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "¨˘©w¥¡≠≠"
   'end 2010/1/11
   intCol = 2
   
   'Add By Sindy 2023/5/23 ≤Qµÿª°ºW•[≈„•‹µo§Â§È
   If txt1(5) = "1" Then '•º¶¨§Â
      intCol = intCol + 1
      Printer.CurrentX = PLeft(intCol)
      Printer.CurrentY = iPrint
      Printer.Print "µo§Â§È"
   End If
   '2023/5/23 END
   
   intCol = intCol + 1
   'Modified by Lydia 2019/06/24 PLeft(2)=>PLeft(3)
   Printer.CurrentX = PLeft(intCol)
   Printer.CurrentY = iPrint
   Printer.Print "•ª©“Æ◊∏π"
   
   'Modify By Sindy 2023/5/23 ≤Qµÿª°•º¶¨§Â§£≈„•‹
   If txt1(5) = "2" Then '§w¶¨§Â
   '2023/5/23 END
      intCol = intCol + 1
      'Modified by Lydia 2019/06/24 PLeft(3)=>PLeft(4)
      Printer.CurrentX = PLeft(intCol)
      Printer.CurrentY = iPrint
      Printer.Print "Æ◊•Û¶W∫Ÿ"
   End If
   
   intCol = intCol + 1
   'Modified by Lydia 2019/06/24 PLeft(4)=>PLeft(5)
   Printer.CurrentX = PLeft(intCol)
   Printer.CurrentY = iPrint
   If txt1(5) = "1" Then
      Printer.Print "§U§@µ{ß«"
   Else
      Printer.Print "Æ◊•Û© ΩË"
   End If
   
   intCol = intCol + 1
   'Modified by Lydia 2019/06/24 PLeft(5)=>PLeft(6)
   Printer.CurrentX = PLeft(intCol)
   Printer.CurrentY = iPrint
   If txt1(5) = "1" Then '•º¶¨§Â
      'Modify By Sindy 2023/5/23
      If strRptType = "2" Then
         Printer.Print "∫ﬁ®Ó§H"
      Else
      '2023/5/23 END
         Printer.Print "¥º≈v§H≠˚"
      End If
   Else
      Printer.Print "©”øÏ§H"
   End If
   
   intCol = intCol + 1
   'Modified by Lydia 2019/06/24 PLeft(6)=>PLeft(7)
   Printer.CurrentX = PLeft(intCol)
   Printer.CurrentY = iPrint
   Printer.Print "•N≤z§H"
   
   intCol = intCol + 1
   'Modified by Lydia 2019/06/24 PLeft(7)=>PLeft(8)
   Printer.CurrentX = PLeft(intCol)
   Printer.CurrentY = iPrint
   Printer.Print "•N≤z§H∞Íƒy"
   
   intCol = intCol + 1
   'Modified by Lydia 2019/06/24 PLeft(8)=>PLeft(9)
   Printer.CurrentX = PLeft(intCol)
   Printer.CurrentY = iPrint
   Printer.Print "∞ •N√∫"
   
   intCol = intCol + 1
   'Modified by Lydia 2019/06/24 PLeft(9)=>PLeft(10)
   Printer.CurrentX = PLeft(intCol)
   Printer.CurrentY = iPrint
   Printer.Print "∞ •N√∫"
   
   intCol = intCol + 1
   'Added by Lydia 2019/06/24 πÍºf¶€∞ •N√∫
   Printer.CurrentX = PLeft(intCol)
   Printer.CurrentY = iPrint
   Printer.Print "∞ •N√∫"
   'end 2019/06/24
   
   'Modify By Sindy 2023/5/23 ≤Qµÿª°•º¶¨§Â§£≈„•‹
   If txt1(5) = "2" Then '§w¶¨§Â
   '2023/5/23 END
      intCol = intCol + 1
      'Modified by Lydia 2019/06/24 PLeft(10)=>PLeft(12)
      Printer.CurrentX = PLeft(intCol)
      Printer.CurrentY = iPrint
      Printer.Print "Æ÷ΩZ§H"
   End If
   
   If txt1(6) <> "202" Then 'Added by Lydia 2020/05/18 "∏…§Â•Û¥¡≠≠∫ﬁ®Ó™Ì"©Ô¿Y§£¶L≥∆µ˘
       intCol = intCol + 1
       'Modified by Lydia 2019/06/24 PLeft(11)=>PLeft(13)
       Printer.CurrentX = PLeft(intCol)
       Printer.CurrentY = iPrint
       Printer.Print "≥∆µ˘"
   End If 'Added by Lydia 2020/05/18
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(170, "-")
   iPrint = iPrint + 300
End Sub

Sub PrintTitle3_A4()
   Printer.EndDoc 'Add By Sindy 2019/4/30
   Printer.Orientation = 2 'Add By Sindy 2019/4/30
   GetPleft3_A4
   iPrint = 500
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.Font.Name = "≤”©˙≈È"
   If Option1(0).Value = True Then
      Printer.CurrentX = 8000 - (Printer.TextWidth(GetTitleNick & "•ª©“¥¡≠≠∫ﬁ®Ó™Ì") / 2)
      Printer.CurrentY = iPrint
      Printer.Print GetTitleNick & "•ª©“¥¡≠≠∫ﬁ®Ó™Ì"
   Else
      Printer.CurrentX = 8000 - (Printer.TextWidth(GetTitleNick & "™k©w¥¡≠≠∫ﬁ®Ó™Ì") / 2)
      Printer.CurrentY = iPrint
      Printer.Print GetTitleNick & "™k©w¥¡≠≠∫ﬁ®Ó™Ì"
   End If
   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   If Option1(0).Value = True Then
      Printer.CurrentX = 8000 - (Printer.TextWidth("•ª©“¥¡≠≠°G" & ChangeTStringToTDateString(txt1(1)) & "°–" & ChangeTStringToTDateString(txt1(2))) / 2)
      Printer.CurrentY = iPrint
      Printer.Print "•ª©“¥¡≠≠°G" & ChangeTStringToTDateString(txt1(1)) & "°–" & ChangeTStringToTDateString(txt1(2))
   Else
      Printer.CurrentX = 8000 - (Printer.TextWidth("™k©w¥¡≠≠°G" & ChangeTStringToTDateString(txt1(3)) & "°–" & ChangeTStringToTDateString(txt1(4))) / 2)
      Printer.CurrentY = iPrint
      Printer.Print "™k©w¥¡≠≠°G" & ChangeTStringToTDateString(txt1(3)) & "°–" & ChangeTStringToTDateString(txt1(4))
   End If
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "¶C¶L§H°G" & strUserName
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "¶C¶L§È¥¡°G" & Format(GetTaiwanTodayDate, "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "©”øÏ§H°G" & GetPrjSalesNM(strTemp3(0)) & IIf(Left(strTemp3(0), 1) = "F", " " & strTemp3(0), "")
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "≠∂°@°@¶∏°G" & str(Page)
   iPrint = iPrint + 300
   Printer.Font.Size = 10
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(170, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "•ª©“¥¡≠≠"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "™k©w¥¡≠≠"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "•ª©“Æ◊∏π"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "Æ◊•Û¶W∫Ÿ"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   If txt1(5) = "1" Then
      Printer.Print "§U§@µ{ß«"
   Else
      Printer.Print "Æ◊•Û© ΩË"
   End If
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "•N≤z§H"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "•N≤z§H∞Íƒy"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "Æ÷ΩZ§H"
   'Add by Morgan 2007/3/26
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = iPrint
   Printer.Print "ßπΩZ§È"
   'end 2007/3/26
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = iPrint
   Printer.Print "∂i´◊≥∆µ˘"
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(170, "-")
   iPrint = iPrint + 300
End Sub

Sub PrintTitle4()
   GetPleft4
   Select Case Val(txt1(14))
      Case 1
          iPrint = 500
          If PrintPage = True Then
              PrintPage = False
              Printer.Font.Size = 12
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "To °G"
              For i = 0 To 3
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("To °G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "¶aß}°G"
              For i = 4 To 8
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("¶aß}°G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              iPrint = iPrint + 300
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "From " & strCmp_C 'Modify by Amy 2020/03/31 ≠Ï:"From •x§@∞Íª⁄±MßQ™k´ﬂ®∆∞»©“"
              iPrint = iPrint + 300
          End If
          Printer.Font.Size = 22
          Printer.Font.Bold = True
          Printer.Font.Underline = True
          Printer.Font.Name = "≤”©˙≈È"
          Printer.CurrentX = 9500 - (Printer.TextWidth("´»§·Æ◊•Û¥¡≠≠∫ﬁ®Ó™Ì") / 2)
          Printer.CurrentY = iPrint
          Printer.Print "´»§·Æ◊•Û¥¡≠≠∫ﬁ®Ó™Ì"
          iPrint = iPrint + 500
          Printer.Font.Size = 12
          Printer.Font.Bold = False
          Printer.Font.Underline = False
          If Option1(0).Value = True Then
               Printer.CurrentX = 9500 - (Printer.TextWidth("¥¡≠≠°G" & Format(ChangeTStringToTDateString(txt1(1)), "@@@@@@@@@@") & "°–" & ChangeTStringToTDateString(txt1(2))) / 2)
               Printer.CurrentY = iPrint
               Printer.Print "¥¡≠≠°G" & Format(ChangeTStringToTDateString(txt1(1)), "@@@@@@@@@@") & "°–" & ChangeTStringToTDateString(txt1(2))
          Else
               Printer.CurrentX = 9500 - (Printer.TextWidth("¥¡≠≠°G" & Format(ChangeTStringToTDateString(txt1(3)), "@@@@@@@@@@") & "°–" & ChangeTStringToTDateString(txt1(4))) / 2)
               Printer.CurrentY = iPrint
               Printer.Print "¥¡≠≠°G" & Format(ChangeTStringToTDateString(txt1(3)), "@@@@@@@@@@") & "°–" & ChangeTStringToTDateString(txt1(4))
          End If
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "¶C¶L§È¥¡°G" & Format(GetTaiwanTodayDate, "##/##/##")
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "≠∂    ¶∏°G" & str(Page)
          iPrint = iPrint + 300
          Printer.Font.Size = 10
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(0)
          Printer.CurrentY = iPrint
          Printer.Print "¥¡  ≠≠"
          Printer.CurrentX = PLeft(1)
          Printer.CurrentY = iPrint
          If txt1(5) = "1" Then
            Printer.Print "§U§@µ{ß«"
          Else
            Printer.Print "Æ◊•Û© ΩË"
          End If
          Printer.CurrentX = PLeft(2)
          Printer.CurrentY = iPrint
          Printer.Print "∂Q©“Æ◊∏π"
          Printer.CurrentX = PLeft(3)
          Printer.CurrentY = iPrint
          Printer.Print "Case No"
          Printer.CurrentX = PLeft(4)
          Printer.CurrentY = iPrint
          Printer.Print "•ª©“Æ◊∏π"
          Printer.CurrentX = PLeft(7)
          Printer.CurrentY = iPrint
          Printer.Print "•”Ω–Æ◊∏π"
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(1) + 1000
          Printer.CurrentY = iPrint
          Printer.Print "Æ◊•Û¶W∫Ÿ"
          Printer.CurrentX = 9000
          Printer.CurrentY = iPrint
          Printer.Print "•”Ω–§H"
          Printer.CurrentX = PLeft(7)
          Printer.CurrentY = iPrint
          Printer.Print "±M ßQ ∏π"
          iPrint = iPrint + 300
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
      Case 2
          iPrint = 500
          If PrintPage = True Then
              PrintPage = False
              Printer.Font.Size = 12
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "To °G"
              For i = 0 To 3
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("To °G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "Address°G"
              For i = 4 To 8
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("Address°G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              iPrint = iPrint + 300
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "From Tai E International Patent & Law Office"
              iPrint = iPrint + 300
          End If
          Printer.Font.Size = 22
          Printer.Font.Bold = True
          Printer.Font.Underline = True
          Printer.Font.Name = "≤”©˙≈È"
          Printer.CurrentX = 9500 - (Printer.TextWidth("Schedule Of Deadline") / 2)
          Printer.CurrentY = iPrint
          Printer.Print "Schedule Of Deadline"
          iPrint = iPrint + 500
          Printer.Font.Size = 12
          Printer.Font.Bold = False
          Printer.Font.Underline = False
          If Option1(0).Value = True Then
               Printer.CurrentX = 9500 - (Printer.TextWidth("From " & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmm DD,YYYY")) / 2)
               Printer.CurrentY = iPrint
               Printer.Print "From " & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmm DD,YYYY")
          Else
               Printer.CurrentX = 9500 - (Printer.TextWidth("From " & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(3))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(4))), "mmm DD,YYYY")) / 2)
               Printer.CurrentY = iPrint
               Printer.Print "From " & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(3))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(4))), "mmm DD,YYYY")
          End If
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "Date°G" & Format(Format(GetTodayDate, "####/##/##"), "mmm DD,YYYY")
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "Page°G" & str(Page)
          iPrint = iPrint + 300
          Printer.Font.Size = 10
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(0)
          Printer.CurrentY = iPrint
          Printer.Print "Deadline"
          Printer.CurrentX = PLeft(1)
          Printer.CurrentY = iPrint
          Printer.Print "Outstanding"
          Printer.CurrentX = PLeft(2)
          Printer.CurrentY = iPrint
          Printer.Print "Your Ref"
          Printer.CurrentX = PLeft(3)
          Printer.CurrentY = iPrint
          Printer.Print "Case No"
          Printer.CurrentX = PLeft(4)
          Printer.CurrentY = iPrint
          Printer.Print "Our Ref"
          Printer.CurrentX = PLeft(7)
          Printer.CurrentY = iPrint
          Printer.Print "Appln. No."
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(1) + 1000
          Printer.CurrentY = iPrint
          Printer.Print "Title"
          Printer.CurrentX = 9000
          Printer.CurrentY = iPrint
          Printer.Print "Applicant"
          Printer.CurrentX = PLeft(7)
          Printer.CurrentY = iPrint
          Printer.Print "Pat. No."
          iPrint = iPrint + 300
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
      Case 3
          iPrint = 500
          If PrintPage = True Then
              PrintPage = False
              Printer.Font.Size = 12
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "TO °G"
              For i = 0 To 3
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("TO °G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "¶Ì©“°G"
              For i = 4 To 8
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("¶Ì©“°G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              iPrint = iPrint + 300
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              'Modified by Lydia 2018/03/16
              'Printer.Print "From •x§@êÔª⁄˚IßQ™k´ﬂ®∆∞»©“"
              Printer.Print "From " & strCmp_J 'Modfiy by Amy 2020/03/31 ≠Ï:Taie_Jpn_Title
              iPrint = iPrint + 300
          End If
          Printer.Font.Size = 22
          Printer.Font.Bold = True
          Printer.Font.Underline = True
          Printer.Font.Name = "≤”©˙≈È"
          Printer.CurrentX = 9500 - (Printer.TextWidth("¥¡≠≠∫ﬁ≤z™Ì") / 2)
          Printer.CurrentY = iPrint
          Printer.Print "¥¡≠≠∫ﬁ≤z™Ì"
          iPrint = iPrint + 500
          Printer.Font.Size = 12
          Printer.Font.Bold = False
          Printer.Font.Underline = False
          If Option1(0).Value = True Then
               Printer.CurrentX = 9500 - (Printer.TextWidth("¥¡≠≠°G" & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmm DD,YYYY")) / 2)
               Printer.CurrentY = iPrint
               Printer.Print "¥¡≠≠°G" & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmm DD,YYYY")
          Else
               Printer.CurrentX = 9500 - (Printer.TextWidth("¥¡≠≠°G" & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(3))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(4))), "mmm DD,YYYY")) / 2)
               Printer.CurrentY = iPrint
               Printer.Print "¥¡≠≠°G" & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(3))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(4))), "mmm DD,YYYY")
          End If
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "ß@¶®§È°G" & Format(Format(GetTodayDate, "####/##/##"), "mmm DD,YYYY")
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "«a∆„«¥°G" & str(Page)
          iPrint = iPrint + 300
          Printer.Font.Size = 10
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(0)
          Printer.CurrentY = iPrint
          Printer.Print "¥¡  ≠≠"
          Printer.CurrentX = PLeft(1)
          Printer.CurrentY = iPrint
          Printer.Print "¶∏«U§‚ë“˛‡"
          Printer.CurrentX = PLeft(2)
          Printer.CurrentY = iPrint
          Printer.Print "∂Qæ„≤zµfêÓ"
          Printer.CurrentX = PLeft(3)
          Printer.CurrentY = iPrint
          Printer.Print "Case No."
          Printer.CurrentX = PLeft(4)
          Printer.CurrentY = iPrint
          Printer.Print "íJæ„≤zµfêÓ"
          Printer.CurrentX = PLeft(7)
          Printer.CurrentY = iPrint
          Printer.Print "•Xƒ@µfêÓ"
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(1) + 1000
          Printer.CurrentY = iPrint
          Printer.Print "Æ◊•Û«U¶Wëb"
          Printer.CurrentX = 9000
          Printer.CurrentY = iPrint
          Printer.Print "•Xƒ@§H"
          Printer.CurrentX = PLeft(7)
          Printer.CurrentY = iPrint
          Printer.Print "ØS≥\ê€µfêÓ"
          iPrint = iPrint + 300
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
      Case Else
   End Select
End Sub

Sub PrintTitle5()
   GetPleft5
   Select Case Val(txt1(14))
      Case 1
          iPrint = 500
          If PrintPage = True Then
              PrintPage = False
              Printer.Font.Size = 12
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "To °G"
              For i = 0 To 3
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("To °G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "¶aß}°G"
              For i = 4 To 8
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("¶aß}°G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              iPrint = iPrint + 300
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "From " & strCmp_C 'Modify by Amy 2020/03/31 ≠Ï:"From •x§@∞Íª⁄±MßQ™k´ﬂ®∆∞»©“"
              iPrint = iPrint + 300
          End If
          Printer.Font.Size = 22
          Printer.Font.Bold = True
          Printer.Font.Underline = True
          Printer.Font.Name = "≤”©˙≈È"
          Printer.CurrentX = 9500 - (Printer.TextWidth("´»§·Æ◊•Û¥¡≠≠∫ﬁ®Ó™Ì") / 2)
          Printer.CurrentY = iPrint
          Printer.Print "´»§·Æ◊•Û¥¡≠≠∫ﬁ®Ó™Ì"
          iPrint = iPrint + 500
          Printer.Font.Size = 12
          Printer.Font.Bold = False
          Printer.Font.Underline = False
          If Option1(0).Value = True Then
               Printer.CurrentX = 9500 - (Printer.TextWidth("¥¡≠≠°G" & Format(ChangeTStringToTDateString(txt1(1)), "@@@@@@@@@@") & "°–" & ChangeTStringToTDateString(txt1(2))) / 2)
               Printer.CurrentY = iPrint
               Printer.Print "¥¡≠≠°G" & Format(ChangeTStringToTDateString(txt1(1)), "@@@@@@@@@@") & "°–" & ChangeTStringToTDateString(txt1(2))
          Else
               Printer.CurrentX = 9500 - (Printer.TextWidth("¥¡≠≠°G" & Format(ChangeTStringToTDateString(txt1(3)), "@@@@@@@@@@") & "°–" & ChangeTStringToTDateString(txt1(4))) / 2)
               Printer.CurrentY = iPrint
               Printer.Print "¥¡≠≠°G" & Format(ChangeTStringToTDateString(txt1(3)), "@@@@@@@@@@") & "°–" & ChangeTStringToTDateString(txt1(4))
          End If
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "¶C¶L§È¥¡°G" & Format(GetTaiwanTodayDate, "##/##/##")
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "≠∂    ¶∏°G" & str(Page)
          iPrint = iPrint + 300
          Printer.Font.Size = 10
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(0)
          Printer.CurrentY = iPrint
          Printer.Print "¥¡  ≠≠"
          Printer.CurrentX = PLeft(1)
          Printer.CurrentY = iPrint
          If txt1(5) = "1" Then
            Printer.Print "§U§@µ{ß«"
          Else
            Printer.Print "Æ◊•Û© ΩË"
          End If
          Printer.CurrentX = PLeft(2)
          Printer.CurrentY = iPrint
          Printer.Print "∂Q©“Æ◊∏π"
          Printer.CurrentX = PLeft(3)
          Printer.CurrentY = iPrint
          Printer.Print "Case No"
          Printer.CurrentX = PLeft(4)
          Printer.CurrentY = iPrint
          Printer.Print "•ª©“Æ◊∏π"
          Printer.CurrentX = PLeft(6)
          Printer.CurrentY = iPrint
          Printer.Print "•”Ω–Æ◊∏π"
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(1) + 1000
          Printer.CurrentY = iPrint
          Printer.Print "Æ◊•Û¶W∫Ÿ"
          Printer.CurrentX = 9000
          Printer.CurrentY = iPrint
          Printer.Print "•” Ω– §H"
          Printer.CurrentX = PLeft(6)
          Printer.CurrentY = iPrint
          Printer.Print "±M ßQ ∏π"
          iPrint = iPrint + 300
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
      Case 2
          iPrint = 500
          If PrintPage = True Then
              PrintPage = False
              Printer.Font.Size = 12
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "To °G"
              For i = 0 To 3
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("To °G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "Address°G"
              For i = 4 To 8
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("Address°G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              iPrint = iPrint + 300
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "From Tai E International Patent & Law Office"
              iPrint = iPrint + 300
          End If
          Printer.Font.Size = 22
          Printer.Font.Bold = True
          Printer.Font.Underline = True
          Printer.Font.Name = "≤”©˙≈È"
          Printer.CurrentX = 9500 - (Printer.TextWidth("Schedule Of Deadline") / 2)
          Printer.CurrentY = iPrint
          Printer.Print "Schedule Of Deadline"
          iPrint = iPrint + 500
          Printer.Font.Size = 12
          Printer.Font.Bold = False
          Printer.Font.Underline = False
          If Option1(0).Value = True Then
               Printer.CurrentX = 9500 - (Printer.TextWidth("From " & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmm DD,YYYY")) / 2)
               Printer.CurrentY = iPrint
              Printer.Print "From " & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmm DD,YYYY")
          Else
               Printer.CurrentX = 9500 - (Printer.TextWidth("From " & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(3))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(4))), "mmm DD,YYYY")) / 2)
               Printer.CurrentY = iPrint
                Printer.Print "From " & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(3))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(4))), "mmm DD,YYYY")
          End If
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "Date°G" & Format(Format(GetTodayDate, "####/##/##"), "mmm DD,YYYY")
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "Page°G" & str(Page)
          iPrint = iPrint + 300
          Printer.Font.Size = 10
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(0)
          Printer.CurrentY = iPrint
          Printer.Print "Deadline"
          Printer.CurrentX = PLeft(1)
          Printer.CurrentY = iPrint
          Printer.Print "Outstanding"
          Printer.CurrentX = PLeft(2)
          Printer.CurrentY = iPrint
          Printer.Print "Your Ref"
          Printer.CurrentX = PLeft(3)
          Printer.CurrentY = iPrint
          Printer.Print "Case No"
          Printer.CurrentX = PLeft(4)
          Printer.CurrentY = iPrint
          Printer.Print "Our Ref"
          Printer.CurrentX = PLeft(6)
          Printer.CurrentY = iPrint
          Printer.Print "Appln. No."
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(1) + 1000
          Printer.CurrentY = iPrint
          Printer.Print "Title"
          Printer.CurrentX = 9000
          Printer.CurrentY = iPrint
          Printer.Print "Applicant"
          Printer.CurrentX = PLeft(6)
          Printer.CurrentY = iPrint
          Printer.Print "Pat. No."
          iPrint = iPrint + 300
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
      Case 3
          iPrint = 500
          If PrintPage = True Then
              PrintPage = False
              Printer.Font.Size = 12
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "TO °G"
              For i = 0 To 3
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("TO °G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "¶Ì©“°G"
              For i = 4 To 8
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("¶Ì©“°G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              iPrint = iPrint + 300
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              'Modified by Lydia 2018/03/16
              'Printer.Print "From •x§@êÔª⁄˚IßQ™k´ﬂ®∆∞»©“"
              Printer.Print "From " & strCmp_J 'Modify by Amy 2020/03/31 ≠Ï:Taie_Jpn_Title
              iPrint = iPrint + 300
          End If
          Printer.Font.Size = 22
          Printer.Font.Bold = True
          Printer.Font.Underline = True
          Printer.Font.Name = "≤”©˙≈È"
          Printer.CurrentX = 9500 - (Printer.TextWidth("¥¡≠≠∫ﬁ≤z™Ì") / 2)
          Printer.CurrentY = iPrint
          Printer.Print "¥¡≠≠∫ﬁ≤z™Ì"
          iPrint = iPrint + 500
          Printer.Font.Size = 12
          Printer.Font.Bold = False
          Printer.Font.Underline = False
          If Option1(0).Value = True Then
               Printer.CurrentX = 9500 - (Printer.TextWidth("¥¡≠≠°G" & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmm DD,YYYY")) / 2)
               Printer.CurrentY = iPrint
              Printer.Print "¥¡≠≠°G" & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmm DD,YYYY")
          Else
               Printer.CurrentX = 9500 - (Printer.TextWidth("¥¡≠≠°G" & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(3))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(4))), "mmm DD,YYYY")) / 2)
               Printer.CurrentY = iPrint
              Printer.Print "¥¡≠≠°G" & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(3))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(4))), "mmm DD,YYYY")
          End If
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "ß@¶®§È°G" & Format(Format(GetTodayDate, "####/##/##"), "mmm DD,YYYY")
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "«a∆„«¥°G" & str(Page)
          iPrint = iPrint + 300
          Printer.Font.Size = 10
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(0)
          Printer.CurrentY = iPrint
          Printer.Print "¥¡  ≠≠"
          Printer.CurrentX = PLeft(1)
          Printer.CurrentY = iPrint
          Printer.Print "¶∏«U§‚ë“˛‡"
          Printer.CurrentX = PLeft(2)
          Printer.CurrentY = iPrint
          Printer.Print "∂Qæ„≤zµfêÓ"
          Printer.CurrentX = PLeft(3)
          Printer.CurrentY = iPrint
          Printer.Print "Case No."
          Printer.CurrentX = PLeft(4)
          Printer.CurrentY = iPrint
          Printer.Print "íJæ„≤zµfêÓ"
          Printer.CurrentX = PLeft(6)
          Printer.CurrentY = iPrint
          Printer.Print "•Xƒ@µfêÓ"
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(1) + 1000
          Printer.CurrentY = iPrint
          Printer.Print "Æ◊•Û«U¶Wëb"
          Printer.CurrentX = 9000
          Printer.CurrentY = iPrint
          Printer.Print "•Xƒ@§H"
          Printer.CurrentX = PLeft(6)
          Printer.CurrentY = iPrint
          Printer.Print "ØS≥\ê€µfêÓ"
          iPrint = iPrint + 300
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
      Case Else
   End Select
End Sub

Sub PrintTitle6()
   GetPleft6
   Select Case Val(txt1(14))
      Case 1
          iPrint = 500
          If PrintPage = True Then
              PrintPage = False
              Printer.Font.Size = 12
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "To °G"
              For i = 0 To 3
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("To °G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "¶aß}°G"
              For i = 4 To 8
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("¶aß}°G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              iPrint = iPrint + 300
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "From " & strCmp_C 'Modify by Amy 2020/03/31 ≠Ï:"From •x§@∞Íª⁄±MßQ™k´ﬂ®∆∞»©“"
              iPrint = iPrint + 300
          End If
          Printer.Font.Size = 22
          Printer.Font.Bold = True
          Printer.Font.Underline = True
          Printer.Font.Name = "≤”©˙≈È"
          Printer.CurrentX = 9500 - (Printer.TextWidth("´»§·¶~∂O¥¡≠≠∫ﬁ®Ó™Ì") / 2)
          Printer.CurrentY = iPrint
          Printer.Print "´»§·¶~∂O¥¡≠≠∫ﬁ®Ó™Ì"
          iPrint = iPrint + 500
          Printer.Font.Size = 12
          Printer.Font.Bold = False
          Printer.Font.Underline = False
          If Option1(0).Value = True Then
               Printer.CurrentX = 9500 - (Printer.TextWidth("¥¡≠≠°G" & Format(ChangeTStringToTDateString(txt1(1)), "@@@@@@@@@@") & "°–" & ChangeTStringToTDateString(txt1(2))) / 2)
               Printer.CurrentY = iPrint
              Printer.Print "¥¡≠≠°G" & Format(ChangeTStringToTDateString(txt1(1)), "@@@@@@@@@@") & "°–" & ChangeTStringToTDateString(txt1(2))
          Else
               Printer.CurrentX = 9500 - (Printer.TextWidth("¥¡≠≠°G" & Format(ChangeTStringToTDateString(txt1(3)), "@@@@@@@@@@") & "°–" & ChangeTStringToTDateString(txt1(4))) / 2)
               Printer.CurrentY = iPrint
              Printer.Print "¥¡≠≠°G" & Format(ChangeTStringToTDateString(txt1(3)), "@@@@@@@@@@") & "°–" & ChangeTStringToTDateString(txt1(4))
          End If
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "¶C¶L§È¥¡°G" & Format(GetTaiwanTodayDate, "##/##/##")
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "≠∂    ¶∏°G" & str(Page)
          iPrint = iPrint + 300
          Printer.Font.Size = 10
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(0)
          Printer.CurrentY = iPrint
          Printer.Print "¥¡  ≠≠"
          Printer.CurrentX = PLeft(1)
          Printer.CurrentY = iPrint
          Printer.Print "∂Q©“Æ◊∏π"
          Printer.CurrentX = PLeft(2)
          Printer.CurrentY = iPrint
          Printer.Print "Case No"
          Printer.CurrentX = PLeft(3)
          Printer.CurrentY = iPrint
          Printer.Print "•ª©“Æ◊∏π"
          Printer.CurrentX = PLeft(6)
          Printer.CurrentY = iPrint
          Printer.Print "•”Ω–Æ◊∏π"
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(1) + 1000
          Printer.CurrentY = iPrint
          Printer.Print "Æ◊•Û¶W∫Ÿ"
          Printer.CurrentX = 9000
          Printer.CurrentY = iPrint
          Printer.Print "•”Ω–§H"
          Printer.CurrentX = PLeft(6)
          Printer.CurrentY = iPrint
          Printer.Print "±M ßQ ∏π"
          iPrint = iPrint + 300
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
      Case 2
          iPrint = 500
          If PrintPage = True Then
              PrintPage = False
              Printer.Font.Size = 12
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "To °G"
              For i = 0 To 3
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("To °G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "Address°G"
              For i = 4 To 8
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("Address°G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              iPrint = iPrint + 300
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "From Tai E International Patent & Law Office"
              iPrint = iPrint + 300
          End If
          Printer.Font.Size = 22
          Printer.Font.Bold = True
          Printer.Font.Underline = True
          Printer.Font.Name = "≤”©˙≈È"
          Printer.CurrentX = 9500 - (Printer.TextWidth("Schedule Of Annuity Deadline") / 2)
          Printer.CurrentY = iPrint
          Printer.Print "Schedule Of Annuity Deadline"
          iPrint = iPrint + 500
          Printer.Font.Size = 12
          Printer.Font.Bold = False
          Printer.Font.Underline = False
          If Option1(0).Value = True Then
               Printer.CurrentX = 9500 - (Printer.TextWidth("From " & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmm DD,YYYY")) / 2)
               Printer.CurrentY = iPrint
              Printer.Print "From " & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmm DD,YYYY")
          Else
               Printer.CurrentX = 9500 - (Printer.TextWidth("From " & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(3))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(4))), "mmm DD,YYYY")) / 2)
               Printer.CurrentY = iPrint
              Printer.Print "From " & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(3))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(4))), "mmm DD,YYYY")
          End If
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "Date°G" & Format(Format(GetTodayDate, "####/##/##"), "mmm DD,YYYY")
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "Page°G" & str(Page)
          iPrint = iPrint + 300
          Printer.Font.Size = 10
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(0)
          Printer.CurrentY = iPrint
          Printer.Print "Deadline"
          Printer.CurrentX = PLeft(1)
          Printer.CurrentY = iPrint
          Printer.Print "Your Ref"
          Printer.CurrentX = PLeft(2)
          Printer.CurrentY = iPrint
          Printer.Print "Case No"
          Printer.CurrentX = PLeft(3)
          Printer.CurrentY = iPrint
          Printer.Print "Our Ref"
          Printer.CurrentX = PLeft(6)
          Printer.CurrentY = iPrint
          Printer.Print "Appln. No."
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(1) + 1000
          Printer.CurrentY = iPrint
          Printer.Print "Title"
          Printer.CurrentX = 9000
          Printer.CurrentY = iPrint
          Printer.Print "Applicant"
          Printer.CurrentX = PLeft(6)
          Printer.CurrentY = iPrint
          Printer.Print "Pat. No."
          iPrint = iPrint + 300
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
      Case 3
          iPrint = 500
          If PrintPage = True Then
              PrintPage = False
              Printer.Font.Size = 12
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "TO °G"
              For i = 0 To 3
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("TO °G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "¶Ì©“°G"
              For i = 4 To 8
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("¶Ì©“°G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              iPrint = iPrint + 300
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              'Modified by Lydia 2018/03/16
              'Printer.Print "•x§@êÔª⁄˚IßQ™k´ﬂ®∆∞»©“"
              Printer.Print strCmp_J 'Modify by Amy 2020/03/31 ≠Ï:Taie_Jpn_Title
              iPrint = iPrint + 300
          End If
          Printer.Font.Size = 22
          Printer.Font.Bold = True
          Printer.Font.Underline = True
          Printer.Font.Name = "≤”©˙≈È"
          Printer.CurrentX = 9500 - (Printer.TextWidth("¶~™˜∫ﬁ≤z™Ì") / 2)
          Printer.CurrentY = iPrint
          Printer.Print "¶~™˜∫ﬁ≤z™Ì"
          iPrint = iPrint + 500
          Printer.Font.Size = 12
          Printer.Font.Bold = False
          Printer.Font.Underline = False
          If Option1(0).Value = True Then
               Printer.CurrentX = 9500 - (Printer.TextWidth("¥¡≠≠°G" & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmm DD,YYYY")) / 2)
               Printer.CurrentY = iPrint
              Printer.Print "¥¡≠≠°G" & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmm DD,YYYY")
          Else
               Printer.CurrentX = 9500 - (Printer.TextWidth("¥¡≠≠°G" & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(3))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(4))), "mmm DD,YYYY")) / 2)
               Printer.CurrentY = iPrint
              Printer.Print "¥¡≠≠°G" & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(3))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(4))), "mmm DD,YYYY")
          End If
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "ß@¶®§È°G" & Format(Format(GetTodayDate, "####/##/##"), "mmm DD,YYYY")
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "«a∆„«¥°G" & str(Page)
          iPrint = iPrint + 300
          Printer.Font.Size = 10
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(0)
          Printer.CurrentY = iPrint
          Printer.Print "¥¡°@≠≠"
          Printer.CurrentX = PLeft(1)
          Printer.CurrentY = iPrint
          Printer.Print "∂Qæ„≤zµfêÓ"
          Printer.CurrentX = PLeft(2)
          Printer.CurrentY = iPrint
          Printer.Print "Case No."
          Printer.CurrentX = PLeft(3)
          Printer.CurrentY = iPrint
          Printer.Print "íJæ„≤zµfêÓ"
          Printer.CurrentX = PLeft(6)
          Printer.CurrentY = iPrint
          Printer.Print "•Xƒ@µfêÓ"
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(1) + 1000
          Printer.CurrentY = iPrint
          Printer.Print "Æ◊•Û«U¶Wëb"
          Printer.CurrentX = 9000
          Printer.CurrentY = iPrint
          Printer.Print "•Xƒ@§H"
          Printer.CurrentX = PLeft(6)
          Printer.CurrentY = iPrint
          Printer.Print "ØS≥\ê€µfêÓ"
          iPrint = iPrint + 300
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
      Case Else
   End Select
End Sub

Sub PrintTitle7()
   GetPleft7
   Select Case Val(txt1(14))
      Case 1
          iPrint = 500
          If PrintPage = True Then
              PrintPage = False
              Printer.Font.Size = 12
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "To °G"
              For i = 0 To 3
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("To °G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "¶aß}°G"
              For i = 4 To 8
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("¶aß}°G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              iPrint = iPrint + 300
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "From " & strCmp_C 'Modify by Amy 2020/03/31 ≠Ï:"From •x§@∞Íª⁄±MßQ™k´ﬂ®∆∞»©“"
              iPrint = iPrint + 300
          End If
          Printer.Font.Size = 22
          Printer.Font.Bold = True
          Printer.Font.Underline = True
          Printer.Font.Name = "≤”©˙≈È"
          Printer.CurrentX = 9500 - (Printer.TextWidth("´»§·¶~∂O¥¡≠≠∫ﬁ®Ó™Ì") / 2)
          Printer.CurrentY = iPrint
          Printer.Print "´»§·¶~∂O¥¡≠≠∫ﬁ®Ó™Ì"
          iPrint = iPrint + 500
          Printer.Font.Size = 12
          Printer.Font.Bold = False
          Printer.Font.Underline = False
          If Option1(0).Value = True Then
               Printer.CurrentX = 9500 - (Printer.TextWidth("¥¡≠≠°G" & Format(ChangeTStringToTDateString(txt1(1)), "@@@@@@@@@@") & "°–" & ChangeTStringToTDateString(txt1(2))) / 2)
               Printer.CurrentY = iPrint
              Printer.Print "¥¡≠≠°G" & Format(ChangeTStringToTDateString(txt1(1)), "@@@@@@@@@@") & "°–" & ChangeTStringToTDateString(txt1(2))
          Else
               Printer.CurrentX = 9500 - (Printer.TextWidth("¥¡≠≠°G" & Format(ChangeTStringToTDateString(txt1(3)), "@@@@@@@@@@") & "°–" & ChangeTStringToTDateString(txt1(4))) / 2)
               Printer.CurrentY = iPrint
              Printer.Print "¥¡≠≠°G" & Format(ChangeTStringToTDateString(txt1(3)), "@@@@@@@@@@") & "°–" & ChangeTStringToTDateString(txt1(4))
          End If
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "¶C¶L§È¥¡°G" & Format(GetTaiwanTodayDate, "##/##/##")
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "≠∂    ¶∏°G" & str(Page)
          iPrint = iPrint + 300
          Printer.Font.Size = 10
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(0)
          Printer.CurrentY = iPrint
          Printer.Print "¥¡  ≠≠"
          Printer.CurrentX = PLeft(1)
          Printer.CurrentY = iPrint
          Printer.Print "∂Q©“Æ◊∏π"
          Printer.CurrentX = PLeft(2)
          Printer.CurrentY = iPrint
          Printer.Print "Case No"
          Printer.CurrentX = PLeft(3)
          Printer.CurrentY = iPrint
          Printer.Print "•ª©“Æ◊∏π"
          Printer.CurrentX = PLeft(5)
          Printer.CurrentY = iPrint
          Printer.Print "•”Ω–Æ◊∏π"
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(1) + 1000
          Printer.CurrentY = iPrint
          Printer.Print "Æ◊•Û¶W∫Ÿ"
          Printer.CurrentX = 9000
          Printer.CurrentY = iPrint
          Printer.Print "•” Ω– §H"
          Printer.CurrentX = PLeft(5)
          Printer.CurrentY = iPrint
          Printer.Print "±M ßQ ∏π"
          iPrint = iPrint + 300
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
      Case 2
          iPrint = 500
          If PrintPage = True Then
              PrintPage = False
              Printer.Font.Size = 12
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "To °G"
              For i = 0 To 3
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("To °G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "Address°G"
              For i = 4 To 8
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("Address°G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              iPrint = iPrint + 300
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "From Tai E International Patent & Law Office"
              iPrint = iPrint + 300
          End If
          Printer.Font.Size = 22
          Printer.Font.Bold = True
          Printer.Font.Underline = True
          Printer.Font.Name = "≤”©˙≈È"
          Printer.CurrentX = 9500 - (Printer.TextWidth("Schedule Of Annuity Deadline") / 2)
          Printer.CurrentY = iPrint
          Printer.Print "Schedule Of Annuity Deadline"
          iPrint = iPrint + 500
          Printer.Font.Size = 12
          Printer.Font.Bold = False
          Printer.Font.Underline = False
          If Option1(0).Value = True Then
               Printer.CurrentX = 9500 - (Printer.TextWidth("From " & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmm DD,YYYY")) / 2)
               Printer.CurrentY = iPrint
              Printer.Print "From " & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmm DD,YYYY")
          Else
               Printer.CurrentX = 9500 - (Printer.TextWidth("From " & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(3))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(4))), "mmm DD,YYYY")) / 2)
               Printer.CurrentY = iPrint
              Printer.Print "From " & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(3))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(4))), "mmm DD,YYYY")
          End If
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "Date°G" & Format(Format(GetTodayDate, "####/##/##"), "mmm DD,YYYY")
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "Page°G" & str(Page)
          iPrint = iPrint + 300
          Printer.Font.Size = 10
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(0)
          Printer.CurrentY = iPrint
          Printer.Print "Deadline"
          Printer.CurrentX = PLeft(1)
          Printer.CurrentY = iPrint
          Printer.Print "Your Ref"
          Printer.CurrentX = PLeft(2)
          Printer.CurrentY = iPrint
          Printer.Print "Case No"
          Printer.CurrentX = PLeft(3)
          Printer.CurrentY = iPrint
          Printer.Print "Our Ref"
          Printer.CurrentX = PLeft(5)
          Printer.CurrentY = iPrint
          Printer.Print "Appln. No."
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(1) + 1000
          Printer.CurrentY = iPrint
          Printer.Print "Title"
          Printer.CurrentX = 9000
          Printer.CurrentY = iPrint
          Printer.Print "Applicant"
          Printer.CurrentX = PLeft(5)
          Printer.CurrentY = iPrint
          Printer.Print "Pat. No."
          iPrint = iPrint + 300
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
      Case 3
          iPrint = 500
          If PrintPage = True Then
              PrintPage = False
              Printer.Font.Size = 12
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "TO °G"
              For i = 0 To 3
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("TO °G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              Printer.Print "¶Ì©“°G"
              For i = 4 To 8
                  If Len(strTemp3(i)) <> 0 Then
                      Printer.CurrentX = 600 + Printer.TextWidth("¶Ì©“°G")
                      Printer.CurrentY = iPrint
                      Printer.Print strTemp3(i)
                      iPrint = iPrint + 300
                  End If
              Next i
              iPrint = iPrint + 300
              Printer.CurrentX = 500
              Printer.CurrentY = iPrint
              'Modified by Lydia 2018/03/16
              'Printer.Print "•x§@êÔª⁄˚IßQ™k´ﬂ®∆∞»©“"
              Printer.Print strCmp_J 'Modify by Amy 2020/03/31 ≠Ï:Taie_Jpn_Title
              iPrint = iPrint + 300
          End If
          Printer.Font.Size = 22
          Printer.Font.Bold = True
          Printer.Font.Underline = True
          Printer.Font.Name = "≤”©˙≈È"
          Printer.CurrentX = 9500 - (Printer.TextWidth("¶~™˜∫ﬁ≤z™Ì") / 2)
          Printer.CurrentY = iPrint
          Printer.Print "¶~™˜∫ﬁ≤z™Ì"
          iPrint = iPrint + 500
          Printer.Font.Size = 12
          Printer.Font.Bold = False
          Printer.Font.Underline = False
          If Option1(0).Value = True Then
               Printer.CurrentX = 9500 - (Printer.TextWidth("¥¡≠≠°G" & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmm DD,YYYY")) / 2)
               Printer.CurrentY = iPrint
              Printer.Print "¥¡≠≠°G" & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(1))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmm DD,YYYY")
          Else
               Printer.CurrentX = 9500 - (Printer.TextWidth("¥¡≠≠°G" & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(3))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(4))), "mmm DD,YYYY")) / 2)
               Printer.CurrentY = iPrint
              Printer.Print "¥¡≠≠°G" & Format(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(3))), "mmm DD,YYYY"), "@@@@@@@@@@") & " To " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(4))), "mmm DD,YYYY")
          End If
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "ß@¶®§È°G" & Format(Format(GetTodayDate, "####/##/##"), "mmm DD,YYYY")
          iPrint = iPrint + 300
          Printer.CurrentX = 16000
          Printer.CurrentY = iPrint
          Printer.Print "«a∆„«¥°G" & str(Page)
          iPrint = iPrint + 300
          Printer.Font.Size = 10
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(0)
          Printer.CurrentY = iPrint
          Printer.Print "¥¡°@≠≠"
          Printer.CurrentX = PLeft(1)
          Printer.CurrentY = iPrint
          Printer.Print "∂Qæ„≤zµfêÓ"
          Printer.CurrentX = PLeft(2)
          Printer.CurrentY = iPrint
          Printer.Print "Case No."
          Printer.CurrentX = PLeft(3)
          Printer.CurrentY = iPrint
          Printer.Print "íJæ„≤zµfêÓ"
          Printer.CurrentX = PLeft(5)
          Printer.CurrentY = iPrint
          Printer.Print "•Xƒ@µfêÓ"
          iPrint = iPrint + 300
          Printer.CurrentX = PLeft(1) + 1000
          Printer.CurrentY = iPrint
          Printer.Print "Æ◊•Û«U¶Wëb"
          Printer.CurrentX = 9000
          Printer.CurrentY = iPrint
          Printer.Print "•Xƒ@§H"
          Printer.CurrentX = PLeft(5)
          Printer.CurrentY = iPrint
          Printer.Print "ØS≥\ê€µfêÓ"
          iPrint = iPrint + 300
          Printer.CurrentX = 500
          Printer.CurrentY = iPrint
          Printer.Print String(280, "-")
          iPrint = iPrint + 300
      Case Else
   End Select
End Sub

'Add By Morgan 2004/12/28 πÍ≈Èºf¨d¥¡≠≠∫ﬁ®Ó™Ì
Sub PrintTitle1_416()
   GetPleft1_416
   iPrint = 500
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.Font.Name = "≤”©˙≈È"
   If Option1(0).Value = True Then
      Printer.CurrentX = 8000 - (Printer.TextWidth("πÍ≈Èºf¨d•ª©“¥¡≠≠∫ﬁ®Ó™Ì") / 2)
      Printer.CurrentY = iPrint
      Printer.Print "πÍ≈Èºf¨d•ª©“¥¡≠≠∫ﬁ®Ó™Ì"
   Else
      Printer.CurrentX = 8000 - (Printer.TextWidth("πÍ≈Èºf¨d™k©w¥¡≠≠∫ﬁ®Ó™Ì") / 2)
      Printer.CurrentY = iPrint
      Printer.Print "πÍ≈Èºf¨d™k©w¥¡≠≠∫ﬁ®Ó™Ì"
   End If
   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   If Option1(0).Value = True Then
      Printer.CurrentX = 8000 - (Printer.TextWidth("•ª©“¥¡≠≠°G" & ChangeTStringToTDateString(txt1(1)) & "°–" & ChangeTStringToTDateString(txt1(2))) / 2)
      Printer.CurrentY = iPrint
      Printer.Print "•ª©“¥¡≠≠°G" & ChangeTStringToTDateString(txt1(1)) & "°–" & ChangeTStringToTDateString(txt1(2))
   Else
      Printer.CurrentX = 8000 - (Printer.TextWidth("™k©w¥¡≠≠°G" & ChangeTStringToTDateString(txt1(3)) & "°–" & ChangeTStringToTDateString(txt1(4))) / 2)
      Printer.CurrentY = iPrint
      Printer.Print "™k©w¥¡≠≠°G" & ChangeTStringToTDateString(txt1(3)) & "°–" & ChangeTStringToTDateString(txt1(4))
   End If
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "¶C¶L§H°G" & strUserName
   Printer.CurrentX = 13600
   Printer.CurrentY = iPrint
   Printer.Print "¶C¶L§È¥¡°G" & Format(GetTaiwanTodayDate, "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "∫ﬁ®Ó§H°G" & GetPrjSalesNM(strTemp3(0))
   
   'Add by Morgan 2009/12/8
   If strPS <> "" Then
      Printer.CurrentX = 8000 - Printer.TextWidth(strPS) / 2
      Printer.CurrentY = iPrint
      Printer.Print strPS
   End If
   
   Printer.CurrentX = 13600
   Printer.CurrentY = iPrint
   Printer.Print "≠∂°@°@¶∏°G" & str(Page)
   iPrint = iPrint + 300
   Printer.Font.Size = 10
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(170, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "•ª©“¥¡≠≠"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "™k©w¥¡≠≠"
   
   'Add by Morgan 2010/1/11
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "¨˘©w¥¡≠≠"
   'end 2010/1/11
   
   'Modify By Sindy 2023/5/29
   If Me.txt1(5) = "1" Then '•º¶¨§Â
      Printer.CurrentX = PLeft(4)
      Printer.CurrentY = iPrint
      Printer.Print "•ª©“Æ◊∏π"
      Printer.CurrentX = PLeft(8)
      Printer.CurrentY = iPrint
      Printer.Print "•N≤z§H∞Íƒy"
      Printer.CurrentX = PLeft(12)
      Printer.CurrentY = iPrint
      Printer.Print "≥∆µ˘"
   '2023/5/29 END
   Else
      Printer.CurrentX = PLeft(3)
      Printer.CurrentY = iPrint
      Printer.Print "•ª©“Æ◊∏π"
      Printer.CurrentX = PLeft(8)
      Printer.CurrentY = iPrint
      Printer.Print "•N≤z§H∞Íƒy"
      Printer.CurrentX = PLeft(13)
      Printer.CurrentY = iPrint
      Printer.Print "≥∆µ˘"
   End If
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(170, "-")
   iPrint = iPrint + 300
End Sub

'∫ﬁ®Ó§H
Function PrintPro1_A4() As Boolean
Dim intP As Integer 'Added by Lydia 2019/06/24

   If Option1(0).Value = True Then
       'Modified by Lydia 2019/06/24 +R036018
       strSql = "SELECT R036001,R036002,R036003,R036004,R036005,R036007,R036009,R036010,R036012,R036013,R036014,R036015,R036016,R036017,R036018 " & _
               " FROM R060308_1 WHERE ID='" & strUserNum & "' ORDER BY decode(R036001,'','0',r036001),decode(SUBSTR(r036002, decode(length(r036002), 8, 1, 9, 2), 8), '', '0', SUBSTR(r036002, decode(length(r036002), 8, 1, 9, 2), 8)),r036004 "
   Else
       'Modified by Lydia 2019/06/24 +R036018
       strSql = "SELECT R036001,R036002,R036003,R036004,R036005,R036007,R036009,R036010,R036012,R036013,R036014,R036015,R036016,R036017,R036018 " & _
               " FROM R060308_1 WHERE ID='" & strUserNum & "' ORDER BY decode(R036001,'','0',r036001),decode(SUBSTR(r036003, decode(length(r036003), 8, 1, 9, 2), 8), '', '0', SUBSTR(r036003, decode(length(r036003), 8, 1, 9, 2), 8)),r036004 "
   End If
   CheckOC
   Page = 1
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           strTemp3(0) = CheckStr(.Fields(0))
            'Modify by Morgan 2004/12/28 πÍ≈Èºf¨d416Æ…ßÔ•X∑s™Ì
            'PrintTitle1_A4
            If txt1(6) = "416" Then
               PrintTitle1_416
            Else 'Memo by Lydia 2020/05/18 +∏…§Â•Û¥¡≠≠∫ﬁ®Ó™Ì
               PrintTitle1_A4
            End If
            '2004/12/28 end
           Do While .EOF = False
               'Modified by Lydia 2019/06/24 ßÔ∂∂ß«
               'For i = 0 To 12
               '    strTemp(i) = CheckStr(.Fields(i + 1))
               'Next i
               intP = 0
               strTemp(intP) = "" & .Fields("R036002") '•ª©“¥¡≠≠
               intP = intP + 1
               strTemp(intP) = "" & .Fields("R036003") '™k©w¥¡≠≠
               intP = intP + 1
               strTemp(intP) = "" & .Fields("R036017") '¨˘©w¥¡≠≠
               'Add By Sindy 2023/5/29
               If txt1(5) = "1" Then '•º¶¨§Â
                  intP = intP + 1
                  strTemp(intP) = "" & .Fields("R036015") 'µo§Â§È
               End If
               '2023/5/29 END
               intP = intP + 1
               strTemp(intP) = "" & .Fields("R036004") '•ª©“Æ◊∏π
               'Modify By Sindy 2023/5/29 •º¶¨§Â§£•X≤{
               If txt1(5) = "2" Then '§w¶¨§Â
               '2023/5/29 END
                  intP = intP + 1
                  'Added by Lydia 2020/05/18
                  If txt1(6) = "202" Then
                      strTemp(intP) = convForm("" & .Fields("R036005"), 20) 'Æ◊•Û¶W∫Ÿ
                  Else
                  'end 2020/05/18
                      strTemp(intP) = convForm("" & .Fields("R036005"), 8) 'Æ◊•Û¶W∫Ÿ
                  End If 'Added by Lydia 2020/05/18
               End If
               intP = intP + 1
               strTemp(intP) = convForm("" & .Fields("R036007"), 20) '§U§@µ{ß«°˛Æ◊•Û© ΩË
               intP = intP + 1
               If GetPrjSalesNM("" & .Fields("R036009")) = "" Then '©”øÏ§H
                  strTemp(intP) = convForm("" & .Fields("R036009"), 8)
               Else
                  strTemp(intP) = convForm(GetPrjSalesNM("" & .Fields("R036009")), 8)
               End If
               intP = intP + 1
               strTemp(intP) = convForm("" & .Fields("R036010"), 10) '•N≤z§H
               intP = intP + 1
               strTemp(intP) = convForm("" & .Fields("R036012"), 10) '•N≤z§H∞Íƒy
               intP = intP + 1
               strTemp(intP) = "" & .Fields("R036013") 'ª‚√“¶€∞ •N√∫
               intP = intP + 1
               strTemp(intP) = "" & .Fields("R036014") '¶~∂O¶€∞ •N√∫
               intP = intP + 1
               strTemp(intP) = "" & .Fields("R036018") 'πÍºf¶€∞ •N√∫
               'Modify By Sindy 2023/5/29 •º¶¨§Â§£•X≤{
               If txt1(5) = "2" Then '§w¶¨§Â
               '2023/5/29 END
                  intP = intP + 1
                  strTemp(intP) = "" & .Fields("R036015") 'Æ÷ΩZ§H
               End If
               intP = intP + 1
               'Added by Lydia 2020/05/18
               If txt1(6) = "202" Then
                  strTemp(intP) = "" & .Fields("R036016")
               Else
               'end 2020/05/18
                  strTemp(intP) = convForm("" & .Fields("R036016"), 20)  '≥∆µ˘ , 16
               End If 'Added by Lydia 2020/05/18
               'end 2019/06/24
               
               If strTemp3(0) <> CheckStr(.Fields(0)) Then
                   strTemp3(0) = CheckStr(.Fields(0))
                   Printer.CurrentX = 500
                   Printer.CurrentY = iPrint
                   Printer.Print String(170, "-")
                   iPrint = iPrint + 300
                   Printer.CurrentX = 500
                   Printer.CurrentY = iPrint
                   'edit by nick 2004/07/13
                   'Printer.Print "* ™Ì•‹πO•ª©“¥¡≠≠, V ™Ì•‹∑Ì§È•ª©“¥¡≠≠, # ™Ì•‹©”øÏ§H©|•º≥q™æ•D∫ﬁæ˜√ˆ®”®Á "
                   'Modify by Morgan 2004/8/16
                   'Printer.Print "* ™Ì•‹πO•ª©“¥¡≠≠, V ™Ì•‹∑Ì§È•ª©“¥¡≠≠, # ™Ì•‹©”øÏ§H©|•º≥q™æ•D∫ﬁæ˜√ˆ®”®Á, ™Ì•‹©|•º§Ωßi°A¶~∂O¨∞πw¶Ù≠» "
                   Printer.Print "* ™Ì•‹πO•ª©“¥¡≠≠, V ™Ì•‹∑Ì§È•ª©“¥¡≠≠, # ™Ì•‹©”øÏ§H©|•º≥q™æ•D∫ﬁæ˜√ˆ®”®Á, ! ™Ì•‹©|•º§Ωßi°A¶~∂O¨∞πw¶Ù≠», & ™Ì•‹6≠”§ÎπO√∫¥¡≠≠"
                   Page = Page + 1
                   Printer.NewPage
                   
                  'Modify by Morgan 2004/12/28 πÍ≈Èºf¨d416Æ…ßÔ•X∑s™Ì
                  'PrintTitle1_A4
                  If txt1(6) = "416" Then
                     PrintTitle1_416
                  Else    'Memo by Lydia 2020/05/18 +∏…§Â•Û¥¡≠≠∫ﬁ®Ó™Ì
                     PrintTitle1_A4
                  End If
               End If
               'Modify by Morgan 2004/12/28 πÍ≈Èºf¨d416Æ…ßÔ•X∑s™Ì
               'PrintDatil1_A4
               If txt1(6) = "416" Then
                  PrintDatil1_416
               Else
                  PrintDatil1_A4
               End If
               'Modified by Lydia 2019/06/24
               'If iPrint >= 10000 Then
               If iPrint >= 11000 Then
                   Page = Page + 1
                   Printer.NewPage
                  'Modify by Morgan 2004/12/28 πÍ≈Èºf¨d416Æ…ßÔ•X∑s™Ì
                  'PrintTitle1_A4
                  If txt1(6) = "416" Then
                     PrintTitle1_416
                  Else 'Memo by Lydia 2020/05/18 +∏…§Â•Û¥¡≠≠∫ﬁ®Ó™Ì
                     PrintTitle1_A4
                  End If
               End If
               .MoveNext
           Loop
       End With
   Else
       PrintPro1_A4 = False
       ChkPrintPro(1) = False
       Exit Function
   End If
   CheckOC
   DoEvents
   
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(170, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   'Modify by Morgan 2004/8/16
   'Printer.Print "* ™Ì•‹πO•ª©“¥¡≠≠, V ™Ì•‹∑Ì§È•ª©“¥¡≠≠, # ™Ì•‹©”øÏ§H©|•º≥q™æ•D∫ﬁæ˜√ˆ®”®Á "
   Printer.Print "* ™Ì•‹πO•ª©“¥¡≠≠, V ™Ì•‹∑Ì§È•ª©“¥¡≠≠, # ™Ì•‹©”øÏ§H©|•º≥q™æ•D∫ﬁæ˜√ˆ®”®Á, ! ™Ì•‹©|•º§Ωßi°A¶~∂O¨∞πw¶Ù≠», & ™Ì•‹6≠”§ÎπO√∫¥¡≠≠"

   Printer.EndDoc: DoEvents
   PrintPro1_A4 = True
   ChkPrintPro(1) = True
End Function

'Modify By Sindy 2023/5/23 ∞—¶“PrintPro1_A4,≤£•X¥º≈v§H≠˚∫ﬁ®Ó™Ì
'¥º≈v§H≠˚
Function PrintPro2_A4() As Boolean
Dim intP As Integer 'Added by Lydia 2019/06/24

   If Option1(0).Value = True Then
       'Modified by Lydia 2019/06/24 +R036018
       strSql = "SELECT R036001,R036002,R036003,R036004,R036005,R036007,R036009,R036010,R036012,R036013,R036014,R036015,R036016,R036017,R036018 " & _
               " FROM R060308_1 WHERE ID='" & strUserNum & "' ORDER BY decode(R036009,'','0',R036009),decode(SUBSTR(r036002, decode(length(r036002), 8, 1, 9, 2), 8), '', '0', SUBSTR(r036002, decode(length(r036002), 8, 1, 9, 2), 8)),r036004 "
   Else
       'Modified by Lydia 2019/06/24 +R036018
       strSql = "SELECT R036001,R036002,R036003,R036004,R036005,R036007,R036009,R036010,R036012,R036013,R036014,R036015,R036016,R036017,R036018 " & _
               " FROM R060308_1 WHERE ID='" & strUserNum & "' ORDER BY decode(R036009,'','0',R036009),decode(SUBSTR(r036003, decode(length(r036003), 8, 1, 9, 2), 8), '', '0', SUBSTR(r036003, decode(length(r036003), 8, 1, 9, 2), 8)),r036004 "
   End If
   CheckOC
   Page = 1
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           strTemp3(0) = CheckStr(.Fields("R036009")) '¥º≈v§H≠˚
            'Modify by Morgan 2004/12/28 πÍ≈Èºf¨d416Æ…ßÔ•X∑s™Ì
            'PrintTitle1_A4
'            If txt1(6) = "416" Then
'               PrintTitle1_416
'            Else 'Memo by Lydia 2020/05/18 +∏…§Â•Û¥¡≠≠∫ﬁ®Ó™Ì
               Call PrintTitle1_A4(2)
'            End If
            '2004/12/28 end
           Do While .EOF = False
               'Modified by Lydia 2019/06/24 ßÔ∂∂ß«
               'For i = 0 To 12
               '    strTemp(i) = CheckStr(.Fields(i + 1))
               'Next i
               intP = 0
               strTemp(intP) = "" & .Fields("R036002") '•ª©“¥¡≠≠
               intP = intP + 1
               strTemp(intP) = "" & .Fields("R036003") '™k©w¥¡≠≠
               intP = intP + 1
               strTemp(intP) = "" & .Fields("R036017") '¨˘©w¥¡≠≠
               intP = intP + 1
               strTemp(intP) = "" & .Fields("R036015") 'µo§Â§È
               intP = intP + 1
               strTemp(intP) = "" & .Fields("R036004") '•ª©“Æ◊∏π
'               intP = intP + 1
'               'Added by Lydia 2020/05/18
'               If txt1(6) = "202" Then
'                   strTemp(intP) = convForm("" & .Fields("R036005"), 20) 'Æ◊•Û¶W∫Ÿ
'               Else
'               'end 2020/05/18
'                   strTemp(intP) = convForm("" & .Fields("R036005"), 8) 'Æ◊•Û¶W∫Ÿ
'               End If 'Added by Lydia 2020/05/18
               intP = intP + 1
               strTemp(intP) = convForm("" & .Fields("R036007"), 20) '§U§@µ{ß«°˛Æ◊•Û© ΩË
               intP = intP + 1
               strTemp(intP) = convForm("" & GetPrjSalesNM(.Fields("R036001")), 8) '∫ﬁ®Ó§H
               intP = intP + 1
               strTemp(intP) = convForm("" & .Fields("R036010"), 10) '•N≤z§H
               intP = intP + 1
               strTemp(intP) = convForm("" & .Fields("R036012"), 10) '•N≤z§H∞Íƒy
               intP = intP + 1
               strTemp(intP) = "" & .Fields("R036013") 'ª‚√“¶€∞ •N√∫
               intP = intP + 1
               strTemp(intP) = "" & .Fields("R036014") '¶~∂O¶€∞ •N√∫
               intP = intP + 1
               strTemp(intP) = "" & .Fields("R036018") 'πÍºf¶€∞ •N√∫
'               intP = intP + 1
'               strTemp(intP) = "" & .Fields("R036015") 'Æ÷ΩZ§H
               intP = intP + 1
               'Added by Lydia 2020/05/18
               If txt1(6) = "202" Then
                  strTemp(intP) = "" & .Fields("R036016")
               Else
               'end 2020/05/18
                  strTemp(intP) = convForm("" & .Fields("R036016"), 20)  '≥∆µ˘ , 16
               End If 'Added by Lydia 2020/05/18
               'end 2019/06/24
               
               If strTemp3(0) <> CheckStr(.Fields("R036009")) Then '¥º≈v§H≠˚
                   strTemp3(0) = CheckStr(.Fields("R036009"))
                   Printer.CurrentX = 500
                   Printer.CurrentY = iPrint
                   Printer.Print String(170, "-")
                   iPrint = iPrint + 300
                   Printer.CurrentX = 500
                   Printer.CurrentY = iPrint
                   'edit by nick 2004/07/13
                   'Printer.Print "* ™Ì•‹πO•ª©“¥¡≠≠, V ™Ì•‹∑Ì§È•ª©“¥¡≠≠, # ™Ì•‹©”øÏ§H©|•º≥q™æ•D∫ﬁæ˜√ˆ®”®Á "
                   'Modify by Morgan 2004/8/16
                   'Printer.Print "* ™Ì•‹πO•ª©“¥¡≠≠, V ™Ì•‹∑Ì§È•ª©“¥¡≠≠, # ™Ì•‹©”øÏ§H©|•º≥q™æ•D∫ﬁæ˜√ˆ®”®Á, ™Ì•‹©|•º§Ωßi°A¶~∂O¨∞πw¶Ù≠» "
                   Printer.Print "* ™Ì•‹πO•ª©“¥¡≠≠, V ™Ì•‹∑Ì§È•ª©“¥¡≠≠, # ™Ì•‹©”øÏ§H©|•º≥q™æ•D∫ﬁæ˜√ˆ®”®Á, ! ™Ì•‹©|•º§Ωßi°A¶~∂O¨∞πw¶Ù≠», & ™Ì•‹6≠”§ÎπO√∫¥¡≠≠"
                   Printer.EndDoc: DoEvents
                   
                   Printer.Orientation = 2
                   'Page = Page + 1
                   Page = 1
                   'Printer.NewPage
                   
'                  'Modify by Morgan 2004/12/28 πÍ≈Èºf¨d416Æ…ßÔ•X∑s™Ì
'                  'PrintTitle1_A4
'                  If txt1(6) = "416" Then
'                     PrintTitle1_416
'                  Else    'Memo by Lydia 2020/05/18 +∏…§Â•Û¥¡≠≠∫ﬁ®Ó™Ì
                     Call PrintTitle1_A4(2)
'                  End If
               End If
               'Modify by Morgan 2004/12/28 πÍ≈Èºf¨d416Æ…ßÔ•X∑s™Ì
               'PrintDatil1_A4
               If txt1(6) = "416" Then
                  PrintDatil1_416
               Else
                  PrintDatil1_A4
               End If
               'Modified by Lydia 2019/06/24
               'If iPrint >= 10000 Then
               If iPrint >= 11000 Then
                   Page = Page + 1
                   Printer.NewPage
'                  'Modify by Morgan 2004/12/28 πÍ≈Èºf¨d416Æ…ßÔ•X∑s™Ì
'                  'PrintTitle1_A4
'                  If txt1(6) = "416" Then
'                     PrintTitle1_416
'                  Else     'Memo by Lydia 2020/05/18 +∏…§Â•Û¥¡≠≠∫ﬁ®Ó™Ì
                     Call PrintTitle1_A4(2)
'                  End If
               End If
               .MoveNext
           Loop
       End With
   Else
       PrintPro2_A4 = False
       ChkPrintPro(2) = False
       Exit Function
   End If
   CheckOC
   DoEvents
   
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(170, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   'Modify by Morgan 2004/8/16
   'Printer.Print "* ™Ì•‹πO•ª©“¥¡≠≠, V ™Ì•‹∑Ì§È•ª©“¥¡≠≠, # ™Ì•‹©”øÏ§H©|•º≥q™æ•D∫ﬁæ˜√ˆ®”®Á "
   Printer.Print "* ™Ì•‹πO•ª©“¥¡≠≠, V ™Ì•‹∑Ì§È•ª©“¥¡≠≠, # ™Ì•‹©”øÏ§H©|•º≥q™æ•D∫ﬁæ˜√ˆ®”®Á, ! ™Ì•‹©|•º§Ωßi°A¶~∂O¨∞πw¶Ù≠», & ™Ì•‹6≠”§ÎπO√∫¥¡≠≠"

   Printer.EndDoc: DoEvents
   PrintPro2_A4 = True
   ChkPrintPro(2) = True
End Function

'Add By Sindy 2015/12/31
Private Sub RunIsMail(strRecvId As String)
Dim strFilePathName As String, strSubject As String, strContent As String
Dim strTempFolder As String
   
   strTempFolder = App.path & "\" & "$$TempFolder"
   
   'Add By Sindy 2023/5/23
   If UCase(pub_DbTerminalName) <> UCase(•ø¶°∏ÍÆ∆Æwπq∏£¶W∫Ÿ) Or _
      InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or _
      Pub_StrUserSt03 = "M51" Then
      If Me.Tag = "§£±“∞ ±H´H" Then
         Printer.EndDoc
         Exit Sub
      Else
         If MsgBox("±µµ€∑|±“∞ ±H´H•\Ø‡°]±H§w¶¨§Â¥¡≠≠Æ◊•Û≤M≥Êµπ§uµ{Æv°^" & vbCrLf & "°A¶≥≠n±Hµo´H•Û∂‹°H" & vbCrLf & strTempFolder, vbYesNo + vbDefaultButton2) = vbNo Then
            Me.Tag = "§£±“∞ ±H´H"
            Printer.EndDoc
            Exit Sub
         End If
      End If
   End If
   '2023/5/23 END
   
   If Dir(strTempFolder, vbDirectory) = "" Then
      MkDir strTempFolder
   End If
   
   If Option1(0).Value = True Then
      strFilePathName = txt1(1) & "~" & txt1(2) & "§w¶¨§Â¥¡≠≠Æ◊•Û_" & strRecvId
   Else
      strFilePathName = txt1(3) & "~" & txt1(4) & "§w¶¨§Â¥¡≠≠Æ◊•Û_" & strRecvId
   End If
   frmPDF.Show
   frmPDF.StartProcess strTempFolder, strFilePathName
   Printer.EndDoc
   frmPDF.EndtProcess
   Unload frmPDF
   strSubject = strFilePathName
   'Modify By Sindy 2019/4/23 + µ˘°G≥¯™Ì¨∞æÓ¶°ÆÊ¶°°A¶]¶π¶C¶LÆ…∞O±o≠n™Ω¶°ßÔæÓ¶°¶C¶L°C
   strContent = "Ω–∞—¶“™˛•Û°I" & vbCrLf & vbCrLf & _
                "®√Ω–øÌ¶u©”øÏ¥¡≠≠•B©Û•ª©“¥¡≠≠´e•Ê•Iµ{ß«≤’¶P§Ø∞e•Û°C" & vbCrLf & vbCrLf & _
                "µ˘°G≥¯™Ì¨∞æÓ¶°ÆÊ¶°°A¶]¶π¶C¶LÆ…∞O±o≠n™Ω¶°ßÔæÓ¶°¶C¶L°C"
   PUB_SendMail strUserNum, strRecvId, "", strSubject, strContent, , strTempFolder & "\" & strFilePathName & ".pdf", , , , , , , , True
End Sub

Function PrintPro3_A4() As Boolean
   'Modify By Sindy 2015/12/31 decode(R038001,'','0',r038001)==> IIf(txt1(8) = "", "R038015", "decode(R038001,'','0',r038001)")
   '                           R038001 ==> IIf(txt1(8) = "", "R038015", "R038001")
   If Option1(0).Value = True Then
       strSql = "SELECT " & IIf(txt1(8) = "", "R038015", "R038001") & ",R038002,R038003,R038004,R038005,R038007,R038009,R038011,R038012,R038013,R038014 " & _
               " FROM R060308_3 WHERE ID='" & strUserNum & "'" & _
               " ORDER BY " & IIf(txt1(8) = "", "R038015", "decode(R038001,'','0',r038001)") & ",decode(SUBSTR(r038002, decode(length(r038002), 8, 1, 9, 2), 8), '', '0', SUBSTR(r038002, decode(length(r038002), 8, 1, 9, 2), 8)),r038004 "
   Else
       strSql = "SELECT " & IIf(txt1(8) = "", "R038015", "R038001") & ",R038002,R038003,R038004,R038005,R038007,R038009,R038011,R038012,R038013,R038014 " & _
               " FROM R060308_3 WHERE ID='" & strUserNum & "'" & _
               " ORDER BY " & IIf(txt1(8) = "", "R038015", "decode(R038001,'','0',r038001)") & ",decode(SUBSTR(r038002, decode(length(r038003), 8, 1, 9, 2), 8), '', '0', SUBSTR(r038003, decode(length(r038003), 8, 1, 9, 2), 8)),r038004 "
   End If
   CheckOC
   Page = 1
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           strTemp3(0) = CheckStr(.Fields(0))
           PrintTitle3_A4
           Do While .EOF = False
               For i = 0 To 8
                   strTemp(i) = CheckStr(.Fields(i + 1))
               Next i
               strTemp(9) = "" & .Fields("R038014") 'Add by Morgan 2007/3/27
               If strTemp3(0) <> CheckStr(.Fields(0)) Then
                   Printer.CurrentX = 500
                   Printer.CurrentY = iPrint
                   Printer.Print String(170, "-")
                   iPrint = iPrint + 300
                   Printer.CurrentX = 500
                   Printer.CurrentY = iPrint
                   'Modify by Morgan 2004/8/16
                   'Printer.Print "* ™Ì•‹πO•ª©“¥¡≠≠, V ™Ì•‹∑Ì§È•ª©“¥¡≠≠, # ™Ì•‹©”øÏ§H©|•º≥q™æ•D∫ﬁæ˜√ˆ®”®Á "
                   Printer.Print "* ™Ì•‹πO•ª©“¥¡≠≠, V ™Ì•‹∑Ì§È•ª©“¥¡≠≠, # ™Ì•‹©”øÏ§H©|•º≥q™æ•D∫ﬁæ˜√ˆ®”®Á, ! ™Ì•‹©|•º§Ωßi°A¶~∂O¨∞πw¶Ù≠», & ™Ì•‹6≠”§ÎπO√∫¥¡≠≠"
                   
                   'Add By Sindy 2015/12/31
                   If txt1(8) = "" Then
                     Call RunIsMail(strTemp3(0))
                     Page = 1
                   Else
                   '2015/12/31 END
                     Page = Page + 1
                     Printer.NewPage
                   End If
                   
                   strTemp3(0) = CheckStr(.Fields(0))
                   PrintTitle3_A4
               End If
               PrintDatil3_A4
               'Modified by Lydia 2019/06/24
               'If iPrint >= 10000 Then
               If iPrint >= 11000 Then
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle3_A4
               End If
               .MoveNext
           Loop
       End With
   Else
       PrintPro3_A4 = False
       ChkPrintPro(3) = False
       Exit Function
   End If
   CheckOC
   DoEvents
                   Printer.CurrentX = 500
                   Printer.CurrentY = iPrint
                   Printer.Print String(170, "-")
                   iPrint = iPrint + 300
                   Printer.CurrentX = 500
                   Printer.CurrentY = iPrint
                   'Modify by Morgan 2004/8/16
                   'Printer.Print "* ™Ì•‹πO•ª©“¥¡≠≠, V ™Ì•‹∑Ì§È•ª©“¥¡≠≠, # ™Ì•‹©”øÏ§H©|•º≥q™æ•D∫ﬁæ˜√ˆ®”®Á "
                   Printer.Print "* ™Ì•‹πO•ª©“¥¡≠≠, V ™Ì•‹∑Ì§È•ª©“¥¡≠≠, # ™Ì•‹©”øÏ§H©|•º≥q™æ•D∫ﬁæ˜√ˆ®”®Á, ! ™Ì•‹©|•º§Ωßi°A¶~∂O¨∞πw¶Ù≠», & ™Ì•‹6≠”§ÎπO√∫¥¡≠≠"
   
   'Add By Sindy 2015/12/31
   If txt1(8) = "" Then
     Call RunIsMail(strTemp3(0))
   Else
   '2015/12/31 END
      Printer.EndDoc
   End If
   DoEvents
   PrintPro3_A4 = True
   ChkPrintPro(3) = True
End Function

Function PrintPro4() As Boolean
   For i = 0 To 8
      strTemp3(i) = ""
   Next i
   strSql = "SELECT * FROM R060308_4 WHERE ID='" & strUserNum & "' ORDER BY R039001,R039002,R039003,R039004,R039005,R039006,R039007,R039008,R039009,R039010,R039013 "
   CheckOC
   Page = 1
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           For i = 0 To 8
               strTemp3(i) = CheckStr(.Fields(i))
           Next i
           PrintPage = True
           PrintTitle4
           Do While .EOF = False
               strTemp(0) = CheckStr(.Fields(9))
               strTemp(1) = CheckStr(.Fields(10))
               strTemp(2) = CheckStr(.Fields(11))
               strTemp(3) = CheckStr(.Fields(16))
               strTemp(4) = CheckStr(.Fields(12))
               strTemp(5) = CheckStr(.Fields(13))
               strTemp(6) = CheckStr(.Fields(0))
               strTemp(7) = CheckStr(.Fields(14))
               strTemp(8) = CheckStr(.Fields(15))
               If strTemp3(0) <> CheckStr(.Fields(0)) Then
                   strTemp3(0) = CheckStr(.Fields(0))
                   PrintPage = True
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle4
               End If
               PrintDatil4
               If iPrint >= 14000 Then
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle4
               End If
               .MoveNext
           Loop
       End With
   Else
       PrintPro4 = False
       ChkPrintPro(4) = False
       Exit Function
   End If
   CheckOC
   DoEvents
   Printer.EndDoc
   PrintPro4 = True
   ChkPrintPro(4) = True
End Function

Function PrintPro5() As Boolean
   For i = 0 To 8
      strTemp3(i) = ""
   Next i
   strSql = "SELECT * FROM R060308_5 WHERE ID='" & strUserNum & "' ORDER BY R040001,R040002,R040003,R040004,R040005,R040006,R040007,R040008,R040009,R040010,R040013 "
   CheckOC
   Page = 1
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           For i = 0 To 8
               strTemp3(i) = CheckStr(.Fields(i))
           Next i
           PrintPage = True
           PrintTitle5
           Do While .EOF = False
               'For I = 0 To 8
               '    StrTemp(I) = CheckStr(.Fields(I + 9))
               'Next I
               strTemp(0) = CheckStr(.Fields(9))
               strTemp(1) = CheckStr(.Fields(10))
               strTemp(2) = CheckStr(.Fields(11))
               strTemp(3) = CheckStr(.Fields(17))
               strTemp(4) = CheckStr(.Fields(12))
               strTemp(5) = CheckStr(.Fields(13))
               strTemp(6) = CheckStr(.Fields(14))
               strTemp(7) = CheckStr(.Fields(15))
               'Modify By Cheng 2002/11/15
   '            strTemp(8) = CheckStr(.Fields(16))
               strTemp(8) = "" & .Fields(0).Value & " " & .Fields(1).Value & " " & .Fields(2).Value & " " & .Fields(3).Value
               If strTemp3(0) <> CheckStr(.Fields(0)) Then
                   strTemp3(0) = CheckStr(.Fields(0))
                   PrintPage = True
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle5
               End If
               PrintDatil5
               If iPrint >= 14000 Then
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle5
               End If
               .MoveNext
           Loop
       End With
   Else
       PrintPro5 = False
       ChkPrintPro(5) = False
       Exit Function
   End If
   CheckOC
   DoEvents
   Printer.EndDoc
   PrintPro5 = True
   ChkPrintPro(5) = True
End Function

Function PrintPro6() As Boolean
   For i = 0 To 8
      strTemp3(i) = ""
   Next i
   strSql = "SELECT * FROM R060308_6 WHERE ID='" & strUserNum & "' ORDER BY R041001,R041002,R041003,R041004,R041005,R041006,R041007,R041008,R041009,R041010,R041013 "
   CheckOC
   Page = 1
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           strTemp3(0) = CheckStr(.Fields(0))
           strTemp3(1) = CheckStr(.Fields(1))
           strTemp3(2) = CheckStr(.Fields(2))
           strTemp3(3) = CheckStr(.Fields(3))
           strTemp3(4) = CheckStr(.Fields(4))
           strTemp3(5) = CheckStr(.Fields(5))
           strTemp3(6) = CheckStr(.Fields(6))
           strTemp3(7) = CheckStr(.Fields(7))
           strTemp3(8) = CheckStr(.Fields(8))
           PrintPage = True
           PrintTitle6
           Do While .EOF = False
               strTemp(0) = CheckStr(.Fields(9))
               strTemp(1) = CheckStr(.Fields(10))
               strTemp(2) = CheckStr(.Fields(11))
               strTemp(3) = CheckStr(.Fields(12))
               strTemp(4) = CheckStr(.Fields(13))
               strTemp(5) = CheckStr(.Fields(0))
               strTemp(6) = CheckStr(.Fields(14))
               strTemp(7) = CheckStr(.Fields(15))
               If strTemp3(0) <> CheckStr(.Fields(0)) Then
                   strTemp3(0) = CheckStr(.Fields(0))
                   PrintPage = True
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle6
               End If
               PrintDatil6
               If iPrint >= 14000 Then
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle6
               End If
               .MoveNext
           Loop
       End With
   Else
       PrintPro6 = False
       ChkPrintPro(6) = False
       Exit Function
   End If
   CheckOC
   DoEvents
   Printer.EndDoc
   PrintPro6 = True
   ChkPrintPro(6) = True
End Function

Function PrintPro7() As Boolean
   For i = 0 To 8
      strTemp3(i) = ""
   Next i
   strSql = "SELECT * FROM R060308_7 WHERE ID='" & strUserNum & "' ORDER BY R042001,R042002,R042003,R042004,R042005,R042006,R042007,R042008,R042009,R042010,R042013 "
   CheckOC
   Page = 1
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           strTemp3(0) = CheckStr(.Fields(0))
           strTemp3(1) = CheckStr(.Fields(1))
           strTemp3(2) = CheckStr(.Fields(2))
           strTemp3(3) = CheckStr(.Fields(3))
           strTemp3(4) = CheckStr(.Fields(4))
           strTemp3(5) = CheckStr(.Fields(5))
           strTemp3(6) = CheckStr(.Fields(6))
           strTemp3(7) = CheckStr(.Fields(7))
           strTemp3(8) = CheckStr(.Fields(8))
           PrintPage = True
           PrintTitle7
           Do While .EOF = False
               For i = 0 To 7
                   strTemp(i) = CheckStr(.Fields(i + 9))
               Next i
               If strTemp3(0) <> CheckStr(.Fields(0)) Then
                   strTemp3(0) = CheckStr(.Fields(0))
                   PrintPage = True
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle7
               End If
               PrintDatil7
               If iPrint >= 14000 Then
                   Page = Page + 1
                   Printer.NewPage
                   PrintTitle7
               End If
               .MoveNext
           Loop
       End With
   Else
       PrintPro7 = False
       ChkPrintPro(7) = False
       Exit Function
   End If
   CheckOC
   DoEvents
   Printer.EndDoc
   PrintPro7 = True
   ChkPrintPro(7) = True
End Function

Function PrintPro8_A4() As Boolean
   PrintPro8_A4 = PrintPro1_A4 '•º¶¨§Â§Œ§w¶¨§Â≥¯™Ì¶@•Œ
   ChkPrintPro(8) = ChkPrintPro(1)
End Function

Function PrintPro9() As Boolean
   PrintPro9 = PrintPro4
   ChkPrintPro(9) = ChkPrintPro(4)
End Function

Function PrintPro10() As Boolean
   PrintPro10 = PrintPro5
   ChkPrintPro(10) = ChkPrintPro(5)
End Function

Function PrintPro11() As Boolean
   PrintPro11 = PrintPro6
   ChkPrintPro(11) = ChkPrintPro(6)
End Function

Function PrintPro12() As Boolean
   PrintPro12 = PrintPro7
   ChkPrintPro(12) = ChkPrintPro(7)
End Function

'•º¶¨§Â
Sub Process1()
   strSQL1 = ""
   strSQL2 = ""
   StrSQL6 = ""
   StrSQL3 = ""
   StrSQL4 = ""
   strSQL5 = ""
   strSQL1 = strSQL1 & " AND (PA57<>'Y' OR PA57 IS NULL) "
   strSQL2 = strSQL2 & " AND (SP15<>'Y' OR SP15 IS NULL) "
   If Len(txt1(6)) <> 0 Then
      StrSQL6 = StrSQL6 & " AND NP07 IN (" & GetAddStr(txt1(6)) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(6) 'Add By Sindy 2010/12/8
      'Add By Sindy 2021/12/8
      '•º¶¨§Â:≥ÊøW§UÆ◊•Û© ΩË¨∞202∏…§Â•Û'±∆∞£§U§@µ{ß«§ß°i´»§·¥£®—§§ª°°j°i≠^ª°∞—¶“•ª°j°∞¶]•X202•º¶¨§Â≥¯™ÌÆ…∂°¬I§”¶≠°AµL™k∂ °A¨Gª›±∆∞£≥o2∂µ
      '¶˝≠Y•º≥ÊøW§UÆ◊•Û© ΩËÆ…, •ø±`∂ ≥¯™ÌÆ…§¥ª›¶C•X(•ÿ´e∑|¶C•XßY∫˚´˘, §£ª›±∆∞£)
      If txt1(6) = "202" Then
         StrSQL6 = StrSQL6 & " AND instr(np15,'´»§·¥£®—§§ª°')=0 AND instr(np15,'≠^§Â∞—¶“•ª')=0"
      End If
      '2021/12/8 END
   End If
   strSQL1 = strSQL1 + StrSQL6
   strSQL2 = strSQL2 + StrSQL6
   StrSQL6 = " AND NP06 IS NULL "
   If Len(txt1(7)) <> 0 Then
       'Removed by Morgan 2020/4/13 NP10•iØ‡¨O¬˜¬æ§H≠˚,ßÔ∑sºWº»¶sÆ…πL¬o
       'strSQL1 = strSQL1 + " AND NP10='" & txt1(7) & "' "
       'strSQL2 = strSQL2 + " AND NP10='" & txt1(7) & "' "
       'end 2020/4/13
       pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(7) & lbl1(0) 'Add By Sindy 2010/12/8
   End If
   '©”øÏ§H
   If Len(txt1(8)) <> 0 Then
      'Modify By Sindy 2015/12/28 øÈ§J™L´H©˜Æ…,•Á§]≠nøÈ•X•t2≠”Ωs∏π∏ÍÆ∆
      If txt1(8) = "68007" Then
         strSQL1 = strSQL1 + " AND CP14 in('68007','68091','68092') "
         strSQL2 = strSQL2 + " AND CP14 in('68007','68091','68092') "
      Else
      '2015/12/28 END
         strSQL1 = strSQL1 + " AND CP14='" & txt1(8) & "' "
         strSQL2 = strSQL2 + " AND CP14='" & txt1(8) & "' "
      End If
      pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(8) & lbl1(1) 'Add By Sindy 2010/12/8
   End If
   '•”Ω–§H
   If Len(Trim(txt1(10))) <> 0 And Len(Trim(txt1(11))) <> 0 Then
       strSQL1 = strSQL1 + " AND ((PA26>='" & GetNewFagent(txt1(10)) & "' AND PA26<='" & GetNewFagent(txt1(11)) & "') OR (PA27>='" & GetNewFagent(txt1(10)) & "' AND PA27<='" & GetNewFagent(txt1(11)) & "') OR (PA28>='" & GetNewFagent(txt1(10)) & "' AND PA28<='" & GetNewFagent(txt1(11)) & "') OR (PA29>='" & GetNewFagent(txt1(10)) & "' AND PA29<='" & GetNewFagent(txt1(11)) & "') OR (PA30>='" & GetNewFagent(txt1(10)) & "' AND PA30<='" & GetNewFagent(txt1(11)) & "')) "
       strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(txt1(10)) & "' AND SP08<='" & GetNewFagent(txt1(11)) & "') OR (SP58<='" & GetNewFagent(txt1(10)) & "' AND SP58<='" & GetNewFagent(txt1(11)) & "') OR (SP59>='" & GetNewFagent(txt1(10)) & "' AND SP59<='" & GetNewFagent(txt1(11)) & "') OR (SP65>='" & GetNewFagent(txt1(10)) & "' AND SP65<='" & GetNewFagent(txt1(11)) & "') OR (SP66>='" & GetNewFagent(txt1(10)) & "' AND SP66<='" & GetNewFagent(txt1(11)) & "')) "
   Else
       If Len(Trim(txt1(10))) <> 0 And Len(Trim(txt1(11))) = 0 Then
           strSQL1 = strSQL1 + " AND (PA26>='" & GetNewFagent(txt1(10)) & "' OR PA27>='" & GetNewFagent(txt1(10)) & "' OR PA28>='" & GetNewFagent(txt1(10)) & "' OR PA29>='" & GetNewFagent(txt1(10)) & "' OR PA30>='" & GetNewFagent(txt1(10)) & "') "
           strSQL2 = strSQL2 + " AND (SP08>='" & GetNewFagent(txt1(10)) & "' OR SP58>='" & GetNewFagent(txt1(10)) & "' OR SP59>='" & GetNewFagent(txt1(10)) & "' OR SP65>='" & GetNewFagent(txt1(10)) & "' OR SP66>='" & GetNewFagent(txt1(10)) & "') "
       Else
           If Len(Trim(txt1(10))) = 0 And Len(Trim(txt1(11))) <> 0 Then
               strSQL1 = strSQL1 + " AND (PA26<='" & GetNewFagent(txt1(11)) & "' OR PA27<='" & GetNewFagent(txt1(11)) & "' OR PA28<='" & GetNewFagent(txt1(11)) & "' OR PA29<='" & GetNewFagent(txt1(11)) & "' OR PA30<='" & GetNewFagent(txt1(11)) & "') "
               strSQL2 = strSQL2 + " AND (SP08<='" & GetNewFagent(txt1(11)) & "' OR SP58<='" & GetNewFagent(txt1(11)) & "' OR SP59<='" & GetNewFagent(txt1(11)) & "' OR SP65<='" & GetNewFagent(txt1(11)) & "' OR SP66<='" & GetNewFagent(txt1(11)) & "') "
           End If
       End If
   End If
   If Len(Trim(txt1(10))) <> 0 Or Len(Trim(txt1(11))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(6) & txt1(10) & "-" & txt1(11) 'Add By Sindy 2010/12/8
   End If
   '•N≤z§H
   If Len(Trim(txt1(12))) <> 0 And Len(Trim(txt1(13))) <> 0 Then
       strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(12)) & "' AND PA75<='" & GetNewFagent(txt1(13)) & "' "
       strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(12)) & "' AND SP26<='" & GetNewFagent(txt1(13)) & "' "
   Else
       If Len(Trim(txt1(12))) <> 0 And Len(Trim(txt1(13))) = 0 Then
           strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(12)) & "' "
           strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(12)) & "' "
       Else
           If Len(Trim(txt1(12))) = 0 And Len(Trim(txt1(13))) <> 0 Then
               strSQL1 = strSQL1 + " AND PA75<='" & GetNewFagent(txt1(13)) & "' "
               strSQL2 = strSQL2 + " AND SP26<='" & GetNewFagent(txt1(13)) & "' "
           End If
       End If
   End If
   If Len(Trim(txt1(12))) <> 0 Or Len(Trim(txt1(13))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & txt1(12) & "-" & txt1(13) 'Add By Sindy 2010/12/8
   End If
   
   'Add By Sindy 2023/5/22
   '•N≤z§H∞Íƒy:(1.§È•ª 2.´D§È•ª)
   If Len(Trim(txt1(15))) <> 0 Then
      If Trim(txt1(15)) = "1" Then
         strSQL1 = strSQL1 + " AND substr(FA10,1,3)='011'"
         strSQL2 = strSQL2 + " AND substr(FA10,1,3)='011'"
      Else
         strSQL1 = strSQL1 + " AND substr(FA10,1,3)<>'011'"
         strSQL2 = strSQL2 + " AND substr(FA10,1,3)<>'011'"
      End If
      pub_QL05 = pub_QL05 & ";" & Label1(12) & txt1(15) & " " & Label1(11)
   End If
   '2023/5/22 END
   
   '•ª©“¥¡≠≠
   If Option1(0).Value = True Then
      'Add by Morgan 2009/12/8 +FMPÆ◊±¯•Û
      StrSQL3 = strSQL1 & " AND NP02='P' AND EXISTS(SELECT * FROM STAFF X WHERE X.ST01=NP10 AND SUBSTR(X.ST15,1,1)='F')"
      StrSQL4 = strSQL2 & " AND NP02='PS' AND EXISTS(SELECT * FROM STAFF X WHERE X.ST01=NP10 AND SUBSTR(X.ST15,1,1)='F')"
      strSQL5 = StrSQL6
      If Len(txt1(1)) <> 0 Then
         'Modify By Sindy 2021/5/5 •[ßP¬_¨˘©w¥¡≠≠
         'StrSQL6 = StrSQL6 & " AND NP08>=" & Val(ChangeTStringToWString(txt1(1))) & " AND NP08<=" & Val(ChangeTStringToWString(txt1(2)))
         'Modify By Sindy 2021/11/5 •º¶¨§Â°G≠Ï•H•ª©“¥¡≠≠§Œ¨˘©w¥¡≠≠ (ßR∞£•ª©“¥¡≠≠±¯•Û)
         'StrSQL6 = StrSQL6 & " AND ((NP08>=" & Val(ChangeTStringToWString(txt1(1))) & " AND NP08<=" & Val(ChangeTStringToWString(txt1(2))) & ") or (NP23>=" & Val(ChangeTStringToWString(txt1(1))) & " AND NP23<=" & Val(ChangeTStringToWString(txt1(2))) & "))"
         'Modify By Sindy 2021/11/12 ≤Qµÿ:PÆ◊≠n∫˚´˘•Œ•ª©“¥¡≠≠∂ ¥¡≠≠
         'StrSQL6 = StrSQL6 & " AND (NP23>=" & Val(ChangeTStringToWString(txt1(1))) & " AND NP23<=" & Val(ChangeTStringToWString(txt1(2))) & ")"
         StrSQL6 = StrSQL6 & " AND (((NP08>=" & Val(ChangeTStringToWString(txt1(1))) & " AND NP08<=" & Val(ChangeTStringToWString(txt1(2))) & ") and NP02 in('P','CFP','PS','CPS')) or (NP23>=" & Val(ChangeTStringToWString(txt1(1))) & " AND NP23<=" & Val(ChangeTStringToWString(txt1(2))) & "))"
         
         '2021/5/5 END
         'Modified by Morgan 2012/5/25 ±∆∞£©“≠≠®÷®Ï§U≠±
         'strSQL5 = strSQL5 & " AND NP23>=" & Val(ChangeTStringToWString(txt1(1))) & " AND NP23<=" & Val(ChangeTStringToWString(txt1(2)))
         strSQL5 = strSQL5 & " AND ((NP23>=" & Val(DBDATE(txt1(1))) & " AND NP23<=" & Val(DBDATE(txt1(2))) & ") and not (NP08>=" & Val(DBDATE(txt1(1))) & " AND NP08<=" & Val(DBDATE(txt1(2))) & "))"
      Else
         'Modify By Sindy 2021/5/5 •[ßP¬_¨˘©w¥¡≠≠
         'StrSQL6 = StrSQL6 & " AND NP08<=" & Val(ChangeTStringToWString(txt1(2)))
         'Modify By Sindy 2021/11/5 •º¶¨§Â°G≠Ï•H•ª©“¥¡≠≠§Œ¨˘©w¥¡≠≠ (ßR∞£•ª©“¥¡≠≠±¯•Û)
         'StrSQL6 = StrSQL6 & " AND (NP08<=" & Val(ChangeTStringToWString(txt1(2))) & " or NP23<=" & Val(ChangeTStringToWString(txt1(2))) & ")"
         'Modify By Sindy 2021/11/12 ≤Qµÿ:PÆ◊≠n∫˚´˘•Œ•ª©“¥¡≠≠∂ ¥¡≠≠
         'StrSQL6 = StrSQL6 & " AND NP23<=" & Val(ChangeTStringToWString(txt1(2)))
         StrSQL6 = StrSQL6 & " AND ((NP08<=" & Val(ChangeTStringToWString(txt1(2))) & " and NP02 in('P','CFP','PS','CPS')) or NP23<=" & Val(ChangeTStringToWString(txt1(2))) & ")"
         
         '2021/5/5 END
         'Modified by Morgan 2012/5/25 ±∆∞£©“≠≠®÷®Ï§U≠±
         'strSQL5 = strSQL5 & " AND NP23<=" & Val(ChangeTStringToWString(txt1(2)))
         strSQL5 = strSQL5 & " AND (NP23<=" & Val(DBDATE(txt1(2))) & " and not NP08<=" & Val(DBDATE(txt1(2))) & ") "
      End If
      pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/8
   '™k©w¥¡≠≠
   Else
       If Option1(1).Value = True Then
           If Len(txt1(3)) <> 0 Then
           StrSQL6 = StrSQL6 + " AND NP09>=" & Val(ChangeTStringToWString(txt1(3))) & " "
           End If
           StrSQL6 = StrSQL6 + " AND NP09<=" & Val(ChangeTStringToWString(txt1(4))) '& " and np08>=" & Val(GetTodayDate)
           pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/12/8
       End If
   End If
   
   If Len(txt1(0)) <> 0 Then
      'Modified by Morgan 2012/5/25 +∞Í•~≥°∫ﬁ®Ó™∫FMPÆ◊¥¡≠≠
      'strSQL1 = strSQL1 + " AND NP02 IN (" & SQLGrpStr(txt1(0), 1) & ") "
      'strSQL2 = strSQL2 + " AND NP02 IN (" & SQLGrpStr(txt1(0), 5) & ") "
      strSQL1 = strSQL1 + " AND (NP02 IN (" & SQLGrpStr(txt1(0), 1) & ") or (np02 in ('P','CFP') and exists(select * from staff x where x.st01=np10 and SUBSTR(st15,1,1)='F'))) "
      strSQL2 = strSQL2 + " AND (NP02 IN (" & SQLGrpStr(txt1(0), 5) & ") or (np02 in ('PS','CPS') and exists(select * from staff x where x.st01=np10 and SUBSTR(st15,1,1)='F'))) "
      'end 2012/5/25
      pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/8
   End If
   
   'Pro1   ∫ﬁ®Ó§H(•º¶¨§Â)
   'Pro2   ¥º≈v§H≠˚(•º¶¨§Â) --2007/3/26 ®˙Æ¯¥º≈v§H≠˚≥¯™Ì
   'Pro3   ©”øÏ§H(§w¶¨§Â)
   'Pro4   ´»§·Æ◊•Û(•”Ω–§H)(•º¶¨§Â)
   'Pro5   ´»§·Æ◊•Û(•N≤z§H)(•º¶¨§Â)
   'Pro6   ¶~∂O(•”Ω–§H)(•º¶¨§Â)
   'Pro7   ¶~∂O(•N≤z§H)(•º¶¨§Â)
   
   '≠Y•ºøÈ§J•”Ω–§H§Œ•N≤z§H
   If Len(txt1(10)) = 0 And Len(txt1(11)) = 0 And Len(txt1(12)) = 0 And Len(txt1(13)) = 0 Then
       Pro1
       strSql = "SELECT DISTINCT R036001 FROM R060308_1 WHERE ID='" & strUserNum & "' "
       CheckOC
       adoRecordset.CursorLocation = adUseClient
       adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
           With adoRecordset
               InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/8
               .MoveFirst
               Do While .EOF = False
                   cnnConnection.Execute "DELETE FROM R060308_2 WHERE ID='" & strUserNum & "' AND R037001='" & ChgSQL(CheckStr(.Fields(0))) & "' "
                   .MoveNext
               Loop
           End With
       Else
         InsertQueryLog (0) 'Add By Sindy 2010/12/8
       End If
       CheckOC
       If Not ChkPro(1) Then
         ShowNoData
         Exit Sub
       End If
       Printer.Orientation = 2
       DoEvents
       'If Printer.Orientation <> 2 Then Printer.Orientation = 2
       If ChkPro(1) Then
         PrintPro1_A4 '∫ﬁ®Ó§H
         'Add By Sindy 2023/5/23
         If txt1(6) <> "416" Then
            Printer.Orientation = 2
            DoEvents
            PrintPro2_A4 '¥º≈v§H≠˚
         End If
         '2023/5/23 END
       End If
       
       If ChkPrintPro(1) Or ChkPrintPro(2) Then
          ShowPrintOk
       End If
   Else
      '≠Y¶≥øÈ§J•”Ω–§H
       If Len(txt1(10)) <> 0 Then
           'Pro4
           'Pro6
           If Not Pro4 And Not Pro6 Then
               InsertQueryLog (0) 'Add By Sindy 2010/12/8
               ShowNoData
               Exit Sub
           End If
           InsertQueryLog ("") 'Add By Sindy 2010/12/8
           If ChkPro(4) Then PrintPro4
           If ChkPro(6) Then PrintPro6
           If ChkPrintPro(4) Or ChkPrintPro(6) Then
               ShowPrintOk
            End If
      '≠Y¶≥øÈ§J•N≤z§H
       Else
           'Pro5
           'Pro7
           If Not Pro5 And Not Pro7 Then
               InsertQueryLog (0) 'Add By Sindy 2010/12/8
               ShowNoData
               Exit Sub
           End If
           InsertQueryLog ("") 'Add By Sindy 2010/12/8
           If ChkPro(5) Then PrintPro5
           If ChkPro(7) Then PrintPro7
           If ChkPrintPro(5) Or ChkPrintPro(7) Then
               ShowPrintOk
            End If
       End If
   End If
End Sub

Function Pro1() As Boolean
'Pro1  ∫ﬁ®Ó§H(•º¶¨§Â)
   Dim strTmp As String, strTmp1 As String
   Dim strTmp2 As String '≠˚§u•N∏π
   Dim strNote As String
   
    'Pro1   ∫ﬁ®Ó§H(•º¶¨§Â)
   'πw≥]µL∏ÍÆ∆
   ChkPro(1) = False
   'Modify by Morgan 2010/1/11 +NP23
   '•N≤z§H≠^-->§§-->§È
   'Modify by Morgan 2006/2/7 •[∫ﬁ®Ó§H±±®Ó
'   strSQL = "SELECT '',NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(PA05,NVL(PA06,PA07)),PA11,DECODE(PA09,'000',CPM03,CPM04),'¥º≈v§H≠˚',ST02,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)),PA77,NVL(NA03,NA04),PA71,PA70,' ',NP15,CP09,CP27,np22,np07,pa14,pa09 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,STAFF,CASEPROPERTYMAP,FAGENT,NATION WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND NP10=ST01(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND FA10=NA01(+) AND NP01=CP09(+) " & strSQL1 & StrSQL6
'   strSQL = strSQL + " union all select '',NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),SP11,DECODE(SP09,'000',CPM03,CPM04),'¥º≈v§H≠˚',ST02,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)),SP27,NVL(NA03,NA04),' ',' ',' ',NP15,CP09,CP27,np22,np07,0,'' FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF,CASEPROPERTYMAP,FAGENT,NATION WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND NP10=ST01(+)  AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND FA10=NA01(+) AND NP01=CP09(+) " & strSQL2 & StrSQL6
   'modify by sonia 2014/11/14 ≠Y≠”Æ◊™∫FCP¶~∂O¶€∞ •N√∫PA70§ŒFCPª‚√“¶€∞ •N√∫PA71•º≥]©w,•N≤z§H©Œ´»§·¶≥≥]©w§]≠n±a
   'strSql = "SELECT '' C00,NP08 C01,NP09 C02,NP02||'-'||NP03||'-'||NP04||'-'||NP05 C03,NVL(PA05,NVL(PA06,PA07)) C04,PA11 C05,DECODE(PA09,'000',CPM03,CPM04) C06,'¥º≈v§H≠˚' C07,ST02 C08,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)) C09,PA77 C10,NVL(NA03,NA04) C11,PA71 C12,PA70 C13,' ' C14,NP15 C15,CP09 C16,CP27 C17,np22 C18,np07,pa14,pa09,PA76,pa75,pa26, np23 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,STAFF,CASEPROPERTYMAP,FAGENT,NATION WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND NP10=ST01(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND FA10=NA01(+) AND NP01=CP09(+) " & strSQL1 & StrSQL6
   '2014/12/1 modify by sonia ∂»ª‚√“601§Œ¶~∂O605§~≠n±a•X¶€∞ •N√∫ƒÊ
   'Modify By Sindy 2016/3/1 +,np01
   'Modified by Lydia 2019/05/31 ∞Íƒy≠≠®Ó5≠”¶r
   'Modified by Lydia 2019/06/24 +FCPπÍºf¶€∞ •N√∫C19
   'strSql = "SELECT '' C00,NP08 C01,NP09 C02,NP02||'-'||NP03||'-'||NP04||'-'||NP05 C03,NVL(PA05,NVL(PA06,PA07)) C04,PA11 C05,DECODE(PA09,'000',CPM03,CPM04) C06,'¥º≈v§H≠˚' C07,ST02 C08,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)) C09,PA77 C10,SUBSTR(NVL(NA03,NA04),1,5) C11,decode(np07,'601',nvl(PA71,nvl(fa42,cu75)),null) C12,decode(np07,'605',nvl(PA70,nvl(fa41,cu74)),null) C13,' ' C14,NP15 C15,CP09 C16,CP27 C17,np22 C18,np07,pa14,pa09,PA76,pa75,pa26,np23,np01 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,STAFF,CASEPROPERTYMAP,FAGENT,NATION,CUSTOMER WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND NP10=ST01(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND FA10=NA01(+) AND NP01=CP09(+) " & strSQL1 & StrSQL6
   'Modified by Morgan 2024/1/13 ¶~∂O¶€∞ •N√∫Y§~≈„•‹(N™Ìª‚√“´·§£ƒÚøÏ,≠Yª›≠n¿≥•t¶CƒÊ¶Ï)--Bobbie nvl(fa41,cu74)->decode(nvl(fa41,cu74),'Y','Y')
   strSql = "SELECT '' C00,NP08 C01,NP09 C02,NP02||'-'||NP03||'-'||NP04||'-'||NP05 C03,NVL(PA05,NVL(PA06,PA07)) C04,PA11 C05,DECODE(PA09,'000',CPM03,CPM04) C06,'¥º≈v§H≠˚' C07,ST02 C08,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)) C09,PA77 C10,SUBSTR(NVL(NA03,NA04),1,5) C11,decode(np07,'601',nvl(PA71,nvl(fa42,cu75)),null) C12,decode(np07,'605',nvl(PA70,decode(nvl(fa41,cu74),'Y','Y')),null) C13,' ' C14,NP15 C15,CP09 C16,CP27 C17,np22 C18,np07,pa14,pa09,PA76,pa75,pa26,np23,np01,DECODE(NP07,'416',NVL(PA69,NVL(FA124,CU177)),NULL) AS C19 " & _
               "FROM NEXTPROGRESS,CASEPROGRESS,PATENT,STAFF,CASEPROPERTYMAP,FAGENT,NATION,CUSTOMER WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND NP10=ST01(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND FA10=NA01(+) AND NP01=CP09(+) " & strSQL1 & StrSQL6
   'end 2014/11/14
   'Modified by Lydia 2019/06/24 +FCPπÍºf¶€∞ •N√∫C19
   strSql = strSql + " union all select '',NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),SP11,DECODE(SP09,'000',CPM03,CPM04),'¥º≈v§H≠˚',ST02,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)),SP27,SUBSTR(NVL(NA03,NA04),1,5),' ',' ',' ',NP15,CP09,CP27,np22,np07,0,'',NULL,sp26,sp08,np23,np01,'' as C19 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF,CASEPROPERTYMAP,FAGENT,NATION WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND NP10=ST01(+)  AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND FA10=NA01(+) AND NP01=CP09(+) " & strSQL2 & StrSQL6
   'Add by Morgan 2009/12/8 +FMPÆ◊
   If StrSQL3 <> "" Then
      'modify by sonia 2014/11/14 ≠Y≠”Æ◊™∫FCP¶~∂O¶€∞ •N√∫PA70§ŒFCPª‚√“¶€∞ •N√∫PA71•º≥]©w,•N≤z§H©Œ´»§·¶≥≥]©w§]≠n±a
      'strSql = strSql + " union all SELECT '' C00,NP08 C01,NP09 C02,NP02||'-'||NP03||'-'||NP04||'-'||NP05 C03,NVL(PA05,NVL(PA06,PA07)) C04,PA11 C05,DECODE(PA09,'000',CPM03,CPM04) C06,'¥º≈v§H≠˚' C07,ST02 C08,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)) C09,PA77 C10,NVL(NA03,NA04) C11,PA71 C12,PA70 C13,' ' C14,NP15 C15,CP09 C16,CP27 C17,np22 C18,np07,pa14,pa09,PA76,pa75,pa26, np23 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,STAFF,CASEPROPERTYMAP,FAGENT,NATION WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND NP10=ST01(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND FA10=NA01(+) AND NP01=CP09(+) " & StrSQL3 & strSQL5
      '2014/12/1 modify by sonia ∂»ª‚√“601§Œ¶~∂O605§~≠n±a•X¶€∞ •N√∫ƒÊ
      'Modified by Lydia 2019/06/24 +FCPπÍºf¶€∞ •N√∫C19
      'strSql = strSql + " union all SELECT '' C00,NP08 C01,NP09 C02,NP02||'-'||NP03||'-'||NP04||'-'||NP05 C03,NVL(PA05,NVL(PA06,PA07)) C04,PA11 C05,DECODE(PA09,'000',CPM03,CPM04) C06,'¥º≈v§H≠˚' C07,ST02 C08,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)) C09,PA77 C10,SUBSTR(NVL(NA03,NA04),1,5) C11,decode(np07,'601',nvl(PA71,nvl(fa42,cu75)),null) C12,decode(np07,'605',nvl(PA70,nvl(fa41,cu74)),null) C13,' ' C14,NP15 C15,CP09 C16,CP27 C17,np22 C18,np07,pa14,pa09,PA76,pa75,pa26,np23,np01 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,STAFF,CASEPROPERTYMAP,FAGENT,NATION,CUSTOMER WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND NP10=ST01(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND FA10=NA01(+) AND NP01=CP09(+) " & StrSQL3 & strSQL5
      'Modified by Morgan 2024/1/13 ¶~∂O¶€∞ •N√∫Y§~≈„•‹(N™Ìª‚√“´·§£ƒÚøÏ,≠Yª›≠n¿≥•t¶CƒÊ¶Ï)--Bobbie nvl(fa41,cu74)->decode(nvl(fa41,cu74),'Y','Y')
      strSql = strSql + " union all SELECT '' C00,NP08 C01,NP09 C02,NP02||'-'||NP03||'-'||NP04||'-'||NP05 C03,NVL(PA05,NVL(PA06,PA07)) C04,PA11 C05,DECODE(PA09,'000',CPM03,CPM04) C06,'¥º≈v§H≠˚' C07,ST02 C08,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)) C09,PA77 C10,SUBSTR(NVL(NA03,NA04),1,5) C11,decode(np07,'601',nvl(PA71,nvl(fa42,cu75)),null) C12,decode(np07,'605',nvl(PA70,decode(nvl(fa41,cu74),'Y','Y')),null) C13,' ' C14,NP15 C15,CP09 C16,CP27 C17,np22 C18,np07,pa14,pa09,PA76,pa75,pa26,np23,np01,DECODE(NP07,'416',NVL(PA69,NVL(FA124,CU177)),NULL) AS C19 " & _
                              "FROM NEXTPROGRESS,CASEPROGRESS,PATENT,STAFF,CASEPROPERTYMAP,FAGENT,NATION,CUSTOMER WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND NP10=ST01(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND FA10=NA01(+) AND NP01=CP09(+) " & StrSQL3 & strSQL5
      'end 2014/11/14
      strPS = "•ª™Ì•]ßt¨˘©w¥¡≠≠≤≈¶X•ª©“¥¡≠≠±¯•Û§ßFMPÆ◊•Û"
   End If
   If StrSQL4 <> "" Then
      'Modified by Lydia 2019/06/24 +FCPπÍºf¶€∞ •N√∫C19
      strSql = strSql + " union all select '',NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),SP11,DECODE(SP09,'000',CPM03,CPM04),'¥º≈v§H≠˚',ST02,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)),SP27,SUBSTR(NVL(NA03,NA04),1,5),' ',' ',' ',NP15,CP09,CP27,np22,np07,0,'',NULL,sp26,sp08,np23,np01,'' AS C19 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF,CASEPROPERTYMAP,FAGENT,NATION WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND NP10=ST01(+)  AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND FA10=NA01(+) AND NP01=CP09(+) " & StrSQL4 & strSQL5
   End If
   
   'Modify by Morgan 2007/4/25
   '¶~∂O∫ﬁ®Ó§H∂∂ß«:pa76->pa75->pa26.cu96->pa26(PA76•iØ‡¨O´»§·Ωs∏π)
   'Modified by Lydia 2017/02/13 +FMP∫ﬁ®Ó§H
   'Modified by Morgan 2020/5/12 ¶~∂O∫ﬁ®Ó§HßÔßÏÆ◊•Û•N≤z§H™∫∫ﬁ®Ó§H(§£•≤¶“º{¨Oß_¶≥¶~∂O•N≤z§H)
   'If strSrvDate(1) < FMP∫ﬁ®Ó§H±“•Œ§È Then
   '     strSql = "SELECT Y.*,NA16 FROM (" & _
           " select X.*,NVL(NVL(F1.FA10,C1.CU10),NVL(F2.FA10,NVL(F4.FA10,C2.CU10))) FA10" & _
           " from (" & strSql & ") X, FAGENT F1, FAGENT F2,FAGENT F4,CUSTOMER C1, CUSTOMER C2" & _
           " WHERE NP07='605' AND F1.FA01(+)=SUBSTR(PA76,1,8) AND F1.FA02(+)=SUBSTR(PA76||'0',9,1)" & _
           " AND C1.CU01(+)=SUBSTR(PA76,1,8) AND C1.CU02(+)=SUBSTR(PA76||'0',9,1)" & _
           " AND F2.FA01(+)=SUBSTR(PA75,1,8) AND F2.FA02(+)=SUBSTR(PA75||'0',9,1)" & _
           " AND C2.CU01(+)=SUBSTR(PA26,1,8) AND C2.CU02(+)=SUBSTR(PA26||'0',9,1)" & _
           " AND F4.FA01(+)=SUBSTR(C2.CU96,1,8) AND F4.FA02(+)=SUBSTR(C2.CU96||'0',9,1)" & _
           " Union All" & _
           " select X.*,NVL(FA10,CU10) FA10" & _
           " from (" & strSql & ") X, FAGENT,CUSTOMER" & _
           " WHERE NP07<>'605' AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75||'0',9,1)" & _
           " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26||'0',9,1)" & _
           ") Y, NATION WHERE NA01(+)=FA10"
   'Else
   '     strSql = "SELECT Y.*,DECODE(SUBSTR(C03,1,2),'P-',NVL(NA79,NA16),'PS',NVL(NA79,NA16),NA16) NA16 FROM (" & _
           " select X.*,NVL(NVL(F1.FA10,C1.CU10),NVL(F2.FA10,NVL(F4.FA10,C2.CU10))) FA10" & _
           " from (" & strSql & ") X, FAGENT F1, FAGENT F2,FAGENT F4,CUSTOMER C1, CUSTOMER C2" & _
           " WHERE NP07='605' AND F1.FA01(+)=SUBSTR(PA76,1,8) AND F1.FA02(+)=SUBSTR(PA76||'0',9,1)" & _
           " AND C1.CU01(+)=SUBSTR(PA76,1,8) AND C1.CU02(+)=SUBSTR(PA76||'0',9,1)" & _
           " AND F2.FA01(+)=SUBSTR(PA75,1,8) AND F2.FA02(+)=SUBSTR(PA75||'0',9,1)" & _
           " AND C2.CU01(+)=SUBSTR(PA26,1,8) AND C2.CU02(+)=SUBSTR(PA26||'0',9,1)" & _
           " AND F4.FA01(+)=SUBSTR(C2.CU96,1,8) AND F4.FA02(+)=SUBSTR(C2.CU96||'0',9,1)" & _
           " Union All" & _
           " select X.*,NVL(FA10,CU10) FA10" & _
           " from (" & strSql & ") X, FAGENT,CUSTOMER" & _
           " WHERE NP07<>'605' AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75||'0',9,1)" & _
           " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26||'0',9,1)" & _
           ") Y, NATION WHERE NA01(+)=FA10"
   'End If
   'Modify By Sindy 2021/5/10 X.*, => distinct X.*,
   strSql = "SELECT Y.*,DECODE(SUBSTR(C03,1,2),'P-',NVL(NA79,NA16),'PS',NVL(NA79,NA16),NA16) NA16 FROM (" & _
           " select distinct X.*,NVL(FA10,CU10) FA10" & _
           " from (" & strSql & ") X, FAGENT,CUSTOMER" & _
           " WHERE FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75||'0',9,1)" & _
           " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26||'0',9,1)" & _
           ") Y, NATION WHERE NA01(+)=FA10"
   'end 2020/5/12
   'end 2017/02/13
   
   If txt1(9).Text <> "" Then
      'Modified by Lydia 2017/02/13 +FMP∫ﬁ®Ó§H
      If strSrvDate(1) < FMP∫ﬁ®Ó§H±“•Œ§È Then
        strSql = strSql & " AND NA16 ='" & txt1(9).Text & "'"
      Else
        strSql = strSql & " AND DECODE(SUBSTR(C03,1,2),'P-',NVL(NA79,NA16),'PS',NVL(NA79,NA16),NA16) ='" & txt1(9).Text & "'"
      End If
      'end 2017/02/13
      
      pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(9) & lbl1(2) 'Add By Sindy 2010/12/8
   End If
   'end 2007/4/25
   
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           DoEvents
           Do While .EOF = False
               For i = 0 To 15
                   strTemp(i) = ChgSQL(CheckStr(.Fields(i)))
               Next i
               If strTemp(1) < Format(Now, "YYYYMMDD") And Len(strTemp(1)) = 8 Then
                   strTemp(1) = "*" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
               Else
                   If strTemp(1) = Format(Now, "YYYYMMDD") And Len(strTemp(1)) = 8 Then
                       strTemp(1) = "V" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                   Else
                       If Mid(CheckStr(.Fields(16)), 1, 1) = "C" And Len(CheckStr(.Fields(17))) = 0 Then
                           strTemp(1) = "#" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                       Else
                           'add by nick 2004/07/13 •[§J¶~∂Oæn∞OßP¬_
                           If CheckStr(adoRecordset.Fields("np07").Value) = "605" And (SystemNumber(strTemp(3), 1) = "FCP" Or SystemNumber(strTemp(3), 1) = "P") And CheckStr(adoRecordset.Fields("pa09").Value) = "000" And CheckStr(adoRecordset.Fields("pa14").Value) = "" Then
                               strTemp(1) = "!" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                           'Add by Morgan 2004/8/16
                           ElseIf CheckStr(adoRecordset.Fields("np07").Value) = "605" And (SystemNumber(strTemp(3), 1) = "FCP" Or SystemNumber(strTemp(3), 1) = "P") Then
                              
                              stCaseNo = strTemp(3)
                              For j = 1 To 4
                                 iPos = InStr(stCaseNo, "-")
                                 If iPos > 0 Then
                                    stPA(j) = Left(stCaseNo, iPos - 1)
                                    stCaseNo = Mid(stCaseNo, iPos + 1)
                                 Else
                                    stPA(j) = stCaseNo
                                 End If
                              Next j
                              If PUB_IfCtrlDateExtended(stPA, strTemp(2)) = True Then
                                 strTemp(1) = "&" & ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                              Else
                                 strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                              End If
                           'Add end
                           Else
                               strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                           End If
                       End If
                   End If
               End If
               
               strTemp(2) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(2)))
               
'Modify by Morgan 2006/2/9 ßÔ•—§W≠±ªy™k±±®Ó
'               '®˙±oFCP∫ﬁ®Ó§H(NA16)
               strTemp(0) = "" & .Fields("NA16")
'2006/2/9 end
               
               '®˙±oFCP©”øÏ¥º≈v§H≠˚(NA51)
               strTmp = Replace(strTemp(3), "-", "")
               'Modified by Morgan 2012/4/26 +P,CFP
               'If Left(strTmp, 3) = "FCP" Then
               strExc(1) = Left(strTemp(3), InStr(strTemp(3), "-") - 1)
               If strExc(1) = "FCP" Or strExc(1) = "P" Or strExc(1) = "CFP" Then
               'end 2012/4/26
                  If GetFCPSales(strTmp, strTmp1, strTmp2) Then
                     'Removed by Morgan 2020/4/13 ≤Œ§@§U≠±¿À¨d
                     'If Me.txt1(7).Text <> "" Then
                     '   If Me.txt1(7).Text <> strTmp2 Then GoTo NextRecord
                     'End If
                     'end 2020/4/13
                     strTemp(8) = strTmp2 'strTmp1
                  End If
               'Modified by Morgan 2012/4/26 +PS,CPS
               'ElseIf Left(strTmp, 2) = "FG" Then
               ElseIf strExc(1) = "FG" Or strExc(1) = "PS" Or strExc(1) = "CPS" Then
                  If GetFGSales(strTmp, strTmp1, strTmp2) Then
                     'Removed by Morgan 2020/4/13 ≤Œ§@§U≠±¿À¨d
                     'If Me.txt1(7).Text <> "" Then
                     '   If Me.txt1(7).Text <> strTmp2 Then GoTo NextRecord
                     'End If
                     'end 2020/4/13
                     strTemp(8) = strTmp2 'strTmp1
                  End If
               End If
               
               'Added by Morgan 2020/4/13
               If Me.txt1(7).Text <> "" Then
                  If Me.txt1(7).Text <> strTmp2 Then GoTo NextRecord
               End If
               'end 2020/4/13
                     
               'Add by Morgan 2010/1/12
               '¨˘©w¥¡≠≠
               If Not IsNull(.Fields("NP23")) Then
                  strTemp(16) = ChangeTStringToTDateString(TransDate(.Fields("NP23"), 1))
               Else
                  strTemp(16) = ""
               End If
               
               'Add By Sindy 2016/3/1 ¿À¨d¨Oß_§w¶¨©µ¥¡,≠Y¨O,§U§@µ{ß«+(§w¶¨©µ¥¡)
               strExc(0) = "select cp09 from caseprogress" & _
                           " where cp43='" & adoRecordset.Fields("np01").Value & "'" & _
                           " and cp10=404" & _
                           " and cp27||cp57 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strTemp(6) = StrToStr(strTemp(6), 4) & "(§w¶¨©µ¥¡)"
               End If
               '2016/3/1 END

               strTemp(17) = "" & .Fields("C19") 'Added by Lydia 2019/06/24 FCPπÍºf¶€∞ •N√∫
               
               'Modify By Sindy 2023/5/22 ≠ÏÆ÷ΩZ§HƒÊ¶ÏßÔ©Òµo§Â§È(R036015)
               If Not IsNull(.Fields("C17")) Then
                  strTemp(14) = ChangeTStringToTDateString(TransDate(.Fields("C17"), 1))
               Else
                  strTemp(14) = ""
               End If
               '2023/5/22 END
               
               'Modified by Lydia 2019/06/24
               'strSql = "insert into R060308_1(R036001,R036002,R036003,R036004,R036005,R036006,R036007,R036008,R036009,R036010,R036011,R036012,R036013,R036014,R036015,R036016,ID,R036017) VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & strUserNum & "','" & strTemp(16) & "')"
               strSql = "insert into R060308_1(R036001,R036002,R036003,R036004,R036005,R036006,R036007,R036008,R036009,R036010,R036011,R036012,R036013,R036014,R036015,R036016,ID,R036017,R036018) " & _
                           "VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & strUserNum & "','" & strTemp(16) & "','" & strTemp(17) & "')"
                           
               cnnConnection.Execute strSql
               '¶≥∏ÍÆ∆
               ChkPro(1) = True
NextRecord:
               .MoveNext
               DoEvents
           Loop
       End With
   Else
      Pro1 = False
      ChkPro(1) = False
      Exit Function
   End If
   CheckOC
   If ChkPro(1) = False Then
      Pro1 = False
      ChkPro(1) = False
      Exit Function
   End If
   
   Pro1 = True
   ChkPro(1) = True
End Function

Function Pro3() As Boolean
'Pro3   ©”øÏ§H(§w¶¨§Â)
Dim strNote As String 'Add By Sindy 2015/12/28
Dim strR038015 As String 'Add By Sindy 2015/12/31

'πw≥]µL∏ÍÆ∆
ChkPro(3) = False

'Modify by Morgan 2007/4/3 •[±±®Ó§£¶L©”øÏ§H¨∞µ{ß«™∫
'Modify By Sindy 2015/12/28 + ,CP01,CP43
'Modified by Lydia 2019/05/31 ∞Íƒy≠≠®Ó5≠”¶r
strSql = "SELECT ALL S1.ST01,CP06,CP07,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X03,NVL(PA05,NVL(PA06,PA07)) X04,PA11,DECODE(PA09,'000',CPM03,CPM04) X06,S2.ST02,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)) X07,PA77,SUBSTR(NVL(NA03,NA04),1,5) X08,' ' X09,CP64,CP09,CP10,PA76,PA75,PA26,CP01,CP43 " & _
         "FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2,CASEPROPERTYMAP,FAGENT,NATION " & _
         "WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND FA10=NA01(+) AND S1.ST03<>'F22' " & _
         strSQL1 & StrSQL6
strSql = strSql + " union select S1.ST01,CP06,CP07,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),SP11,DECODE(SP09,'000',CPM03,CPM04),S2.ST02,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)),SP27,SUBSTR(NVL(NA03,NA04),1,5),' ',CP64,CP09,CP10,NULL,SP26,SP08,CP01,CP43 " & _
         "FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,FAGENT,NATION " & _
         "WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND FA10=NA01(+) AND S1.ST03<>'F22' " & _
         strSQL2 & StrSQL6
'Add By Sindy 2021/4/15 ºW•[∑sÆ◊¬Ωƒ∂©“§∫§uµ{Æv©”øÏ¥¡≠≠
strSql = strSql + " union SELECT ALL decode(S3.ST01,NULL,S1.ST01,S3.ST01),CP06,CP07,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X03,NVL(PA05,NVL(PA06,PA07)) X04,PA11,DECODE(PA09,'000',CPM03,CPM04) X06,S2.ST02,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)) X07,PA77,SUBSTR(NVL(NA03,NA04),1,5) X08,' ' X09,CP64,CP09,CP10,PA76,PA75,PA26,CP01,CP43 " & _
         "FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2,CASEPROPERTYMAP,FAGENT,NATION,STAFF_IDMAP,STAFF S3 " & _
         "WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND FA10=NA01(+) AND S1.ST03<>'F22' " & _
         "AND SIM02(+)=CP14 AND S3.ST01(+)=SIM01 " & _
         "AND SUBSTRB(S1.ST15,1,1)='F' AND (S3.ST15 IS NULL OR S1.ST15='F52') " & _
         "AND EXISTS(select ts2.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1') " & _
         strSQL1 & strSQL8 & " AND CP10='201' AND exists(select * from engineerprogress x where x.ep02=cp09 and ep09||" & IIf(strSrvDate(1) >= FCPÆ÷ßπ§ÈßÔ•ŒEP39, "ep39", "ep33") & " is null)"
strSql = strSql + " union select decode(S3.ST01,NULL,S1.ST01,S3.ST01),CP06,CP07,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),SP11,DECODE(SP09,'000',CPM03,CPM04),S2.ST02,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)),SP27,SUBSTR(NVL(NA03,NA04),1,5),' ',CP64,CP09,CP10,NULL,SP26,SP08,CP01,CP43 " & _
         "FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,FAGENT,NATION,STAFF_IDMAP,STAFF S3 " & _
         "WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND FA10=NA01(+) AND S1.ST03<>'F22' " & _
         "AND SIM02(+)=CP14 AND S3.ST01(+)=SIM01 " & _
         "AND SUBSTRB(S1.ST15,1,1)='F' AND (S3.ST15 IS NULL OR S1.ST15='F52') " & _
         "AND EXISTS(select ts2.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1') " & _
         strSQL2 & strSQL8 & " AND CP10='201' AND exists(select * from engineerprogress x where x.ep02=cp09 and ep09||" & IIf(strSrvDate(1) >= FCPÆ÷ßπ§ÈßÔ•ŒEP39, "ep39", "ep33") & " is null)"
'2021/4/15 END

'Modify by Morgan 2006/2/14 •[∫ﬁ®Ó§H±¯•Ûªy™k
'strSQL = "select * from (" & strSQL & ") X where " & " not exists(select * from caseprogress a where a.cp43= X.cp09 and a.cp10='907')"
If Me.txt1(9).Text = "" Then
   'Modify by Morgan 2007/3/27 •[ßπΩZ§È
   'strSQL = "select * from (" & strSQL & ") X where " & " not exists(select * from caseprogress a where a.cp43= X.cp09 and a.cp10='907')"
   'Modify by Sindy 2021/5/6 +,decode(ep09||ep33,null,'',decode(EP04,null,'',EP04)) EP04
   strSql = "select X.*,' ' FA10,' ' NA16,DECODE(CP10,'201',EP09) C16,decode(ep09||" & IIf(strSrvDate(1) >= FCPÆ÷ßπ§ÈßÔ•ŒEP39, "ep39", "ep33") & ",null,'',decode(EP04,null,'',EP04)) EP04 from (" & strSql & ") X,ENGINEERPROGRESS where EP02(+)=CP09 AND " & " not exists(select * from caseprogress a where a.cp43= X.cp09 and a.cp10='907')"
Else
   pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(9) & lbl1(2) 'Add By Sindy 2010/12/8
   '¶~∂O∫ﬁ®Ó§H∂∂ß«:pa76->pa75->pa26.cu96->pa26(PA76•iØ‡¨O´»§·Ωs∏π)
   'Modify by Morgan 2007/3/27 •[ßπΩZ§È
   'Modified by Lydia 2017/02/13 +FMP∫ﬁ®Ó§H
   'Modified by Morgan 2020/5/12 ¶~∂O∫ﬁ®Ó§HßÔßÏÆ◊•Û•N≤z§H™∫∫ﬁ®Ó§H(§£•≤¶“º{¨Oß_¶≥¶~∂O•N≤z§H)
   'If strSrvDate(1) < FMP∫ﬁ®Ó§H±“•Œ§È Then
   '     strSql = "SELECT Y.*,NA16,DECODE(CP10,'201',EP09) C16 FROM (" & _
           " select X.*,NVL(NVL(F1.FA10,C1.CU10),NVL(F2.FA10,NVL(F4.FA10,C2.CU10))) FA10" & _
           " from (" & strSql & ") X, FAGENT F1, FAGENT F2,FAGENT F4,CUSTOMER C1, CUSTOMER C2" & _
           " WHERE CP10='605' AND F1.FA01(+)=SUBSTR(PA76,1,8) AND F1.FA02(+)=SUBSTR(PA76||'0',9,1)" & _
           " AND C1.CU01(+)=SUBSTR(PA76,1,8) AND C1.CU02(+)=SUBSTR(PA76||'0',9,1)" & _
           " AND F2.FA01(+)=SUBSTR(PA75,1,8) AND F2.FA02(+)=SUBSTR(PA75||'0',9,1)" & _
           " AND C2.CU01(+)=SUBSTR(PA26,1,8) AND C2.CU02(+)=SUBSTR(PA26||'0',9,1)" & _
           " AND F4.FA01(+)=SUBSTR(C2.CU96,1,8) AND F4.FA02(+)=SUBSTR(C2.CU96||'0',9,1)" & _
           " Union All" & _
           " select X.*,NVL(FA10,CU10) FA10" & _
           " from (" & strSql & ") X, FAGENT,CUSTOMER" & _
           " WHERE CP10<>'605' AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75||'0',9,1)" & _
           " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26||'0',9,1)" & _
           ") Y, NATION,ENGINEERPROGRESS WHERE NA01(+)=FA10 AND EP02(+)=CP09 AND NA16 ='" & txt1(9).Text & "' AND  not exists(select * from caseprogress a where a.cp43= Y.cp09 and a.cp10='907')"
   'Else
   '     strSql = "SELECT Y.*,DECODE(CP01,'P',NVL(NA79,NA16),'PS',NVL(NA79,NA16),NA16) NA16,DECODE(CP10,'201',EP09) C16 FROM (" & _
           " select X.*,NVL(NVL(F1.FA10,C1.CU10),NVL(F2.FA10,NVL(F4.FA10,C2.CU10))) FA10" & _
           " from (" & strSql & ") X, FAGENT F1, FAGENT F2,FAGENT F4,CUSTOMER C1, CUSTOMER C2" & _
           " WHERE CP10='605' AND F1.FA01(+)=SUBSTR(PA76,1,8) AND F1.FA02(+)=SUBSTR(PA76||'0',9,1)" & _
           " AND C1.CU01(+)=SUBSTR(PA76,1,8) AND C1.CU02(+)=SUBSTR(PA76||'0',9,1)" & _
           " AND F2.FA01(+)=SUBSTR(PA75,1,8) AND F2.FA02(+)=SUBSTR(PA75||'0',9,1)" & _
           " AND C2.CU01(+)=SUBSTR(PA26,1,8) AND C2.CU02(+)=SUBSTR(PA26||'0',9,1)" & _
           " AND F4.FA01(+)=SUBSTR(C2.CU96,1,8) AND F4.FA02(+)=SUBSTR(C2.CU96||'0',9,1)" & _
           " Union All" & _
           " select X.*,NVL(FA10,CU10) FA10" & _
           " from (" & strSql & ") X, FAGENT,CUSTOMER" & _
           " WHERE CP10<>'605' AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75||'0',9,1)" & _
           " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26||'0',9,1)" & _
           ") Y, NATION,ENGINEERPROGRESS WHERE NA01(+)=FA10 AND EP02(+)=CP09 AND DECODE(CP01,'P',NVL(NA79,NA16),'PS',NVL(NA79,NA16),NA16) ='" & txt1(9).Text & "' AND  not exists(select * from caseprogress a where a.cp43= Y.cp09 and a.cp10='907')"
   'End If
   'Modify by Sindy 2021/5/6 +,decode(ep09||ep33,null,'',decode(EP04,null,'',EP04)) EP04
   strSql = "SELECT Y.*,DECODE(CP01,'P',NVL(NA79,NA16),'PS',NVL(NA79,NA16),NA16) NA16,DECODE(CP10,'201',EP09) C16,decode(ep09||" & IIf(strSrvDate(1) >= FCPÆ÷ßπ§ÈßÔ•ŒEP39, "ep39", "ep33") & ",null,'',decode(EP04,null,'',EP04)) EP04 FROM (" & _
           " select X.*,NVL(FA10,CU10) FA10" & _
           " from (" & strSql & ") X, FAGENT,CUSTOMER" & _
           " WHERE FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75||'0',9,1)" & _
           " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26||'0',9,1)" & _
           ") Y, NATION,ENGINEERPROGRESS WHERE NA01(+)=FA10 AND EP02(+)=CP09 AND DECODE(CP01,'P',NVL(NA79,NA16),'PS',NVL(NA79,NA16),NA16) ='" & txt1(9).Text & "' AND  not exists(select * from caseprogress a where a.cp43= Y.cp09 and a.cp10='907')"
   'end 2020/5/12
   'end 2017/02/13
End If
'2006/2/14 END

CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'edit by nickc 2007/02/08
'k = 0

If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        Do While .EOF = False
            'Modified by Lydia 2021/06/28 ≤{¶b¶≥24ƒÊ
            'For i = 0 To 22 '12
            For i = 0 To 23
                strTemp(i) = ChgSQL(CheckStr(.Fields(i)))
            Next i
            If strTemp(1) < strSrvDate(1) And Len(strTemp(1)) = 8 Then
                strTemp(1) = "*" & ChangeWStringToTDateString(strTemp(1))
            Else
                If strTemp(1) = strSrvDate(1) And Len(strTemp(1)) = 8 Then
                    strTemp(1) = "V" & ChangeWStringToTDateString(strTemp(1))
                Else
                    strTemp(1) = ChangeWStringToTDateString(strTemp(1))
                End If
            End If
            strTemp(2) = ChangeWStringToTDateString(strTemp(2))
            
'Modify by Morgan 2006/2/9 ßÔ•—ªy™k±±®Ó
'            '®˙±oFCP∫ﬁ®Ó§H
'2006/2/9 END
            
            'Modify By Sindy 2015/12/28 ≥∆µ˘
            strNote = PUB_GetFCPAddQuyNotes(.Fields("cp01"), .Fields("CP09"), .Fields("CP10"), "" & .Fields("cp43"))
            If strNote <> "" Then
               strTemp(12) = strNote & strTemp(12)
            End If
            '2015/12/28 END
            
            '®˙±oEP04.Æ÷ΩZ§H
            strTemp(11) = GetPrjSalesNM(Get060308_1(CheckStr(.Fields(13)))) '∂«§J¡`¶¨§Â∏π
            'Modify by Morgan 2007/3/27 •[ßπΩZ§È
            'Modify by Sindy 2015/12/29 + R038015:¶¨´H§H
            strR038015 = ""
            If txt1(8) = "" Then 'µLøÈ§J©”øÏ§H±¯•Û, ≠n±H´H
               'Add By Sindy 2021/5/6
               'Modified by Lydia 2021/06/28 index¶≥ª~
               'If strTemp(21) <> "" Then
               '   If strTemp(22) = "" Then
               If strTemp(22) <> "" Then 'ßπΩZ§È : DECODE(CP10,'201',EP09) C16
                  If strTemp(23) = "" Then 'Æ÷ΩZ§H:¶≥ßπΩZ§ÈÆ…,¶πƒÊ¶Ï§~∑|©Ò§JÆ÷ΩZ§H decode(ep09||ep33,null,'',decode(EP04,null,'',EP04)) EP04
               'end 2021/06/28
                     '≥q™æ•D∫ﬁ
                     strExc(0) = "SELECT pa150,TCT10 FROM patent,(SELECT DISTINCT cp01 TCTcp01,cp02 TCTcp02,cp03 TCTcp03,cp04 TCTcp04,TCT10 FROM TRANSCASETITLE,caseprogress WHERE tct01=cp09(+) and cp09 is not null) TCT" & _
                                 " WHERE pa01='" & SystemNumber(strTemp(3), 1) & "'" & _
                                 " and pa02='" & SystemNumber(strTemp(3), 2) & "'" & _
                                 " and pa03='" & SystemNumber(strTemp(3), 3) & "'" & _
                                 " and pa04='" & SystemNumber(strTemp(3), 4) & "'" & _
                                 " and TCTcp01(+)=pa01 AND TCTcp02(+)=pa02 AND TCTcp03(+)=pa03 AND TCTcp04(+)=pa04"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        If "" & RsTemp.Fields("TCT10") <> "" Then
                           strR038015 = PUB_GetFCPEngSup(RsTemp.Fields("TCT10"), True) '§uµ{Æv•D∫ﬁ(∞∆≤z)
                        ElseIf "" & RsTemp.Fields("pa150") <> "" Then
                           strR038015 = Pub_GetFCPGrpMan(RsTemp.Fields("pa150"))
                        End If
                     End If
                  Else
                     'Modified by Lydia 2021/06/28 index¶≥ª~
                     'strR038015 = strTemp(22) '±HÆ÷ΩZ§H
                     strR038015 = strTemp(23)
                  End If
               End If
               If strR038015 = "" Then
               '2021/5/6 END
                  If Left(strTemp(0), 1) <> "F" Then
                     strR038015 = strTemp(0) '±H§uµ{Æv
                  Else
'                     '¬Ωƒ∂§H≠˚:¶≥Æ÷ΩZ§H±HÆ÷ΩZ§H,µL´h±H§uµ{Æv•D∫ﬁ
'                     If strTemp(11) <> "" Then
'                        strR038015 = Get060308_1(CheckStr(.Fields(13)))
'                        If Left(strR038015, 1) = "F" Then
'                           strExc(0) = "select st01,st02 from staff where st26=" & _
'                                       "(select st26 from staff where st01='" & strR038015 & "' and st26 is not null)" & _
'                                       " and st04='1' and SUBSTR(st01,1,1)<>'F'"
'                           intI = 1
'                           strR038015 = ""
'                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                           If intI = 1 Then
'                              strR038015 = RsTemp.Fields(0)
'                           End If
'                        End If
'                     End If
'                     If strR038015 = "" Then
'                        strExc(0) = "SELECT pa150 FROM patent" & _
'                                    " WHERE pa01='" & SystemNumber(strTemp(3), 1) & "'" & _
'                                    " and pa02='" & SystemNumber(strTemp(3), 2) & "'" & _
'                                    " and pa03='" & SystemNumber(strTemp(3), 3) & "'" & _
'                                    " and pa04='" & SystemNumber(strTemp(3), 4) & "'"
'                        intI = 1
'                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                        If intI = 1 Then
'                           If "" & RsTemp.Fields("pa150") <> "" Then
'                              'Modified by Lydia 2019/01/09
'                              'strR038015 = IIf(RsTemp.Fields("pa150") = "1", Pub_GetSpecMan("T"), IIf(RsTemp.Fields("pa150") = "2", Pub_GetSpecMan("R"), IIf(RsTemp.Fields("pa150") = "3", Pub_GetSpecMan("S"), Pub_GetSpecMan("T1"))))
'                              strR038015 = Pub_GetFCPGrpMan("" & RsTemp.Fields("pa150"))
'                           End If
'                        End If
'                     End If
                     'Modify By Sindy 2021/5/6
                     strExc(0) = "SELECT st01,st02,st15,SIM01,SIM02" & _
                                 " FROM STAFF_IDMAP,staff" & _
                                 " WHERE ST01=SIM02(+) AND st01='" & strTemp(0) & "'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        If RsTemp.Fields("st15") = "F51" Then '∞Í•~≥°•~¬Ω
                           strR038015 = Pub_GetSpecMan("M") '∞Í•~≥°±MßQ≥B¬Ωƒ∂•º•ÊΩZ∫ﬁ®Ó§H
                        Else
                           If "" & RsTemp.Fields("SIM01") <> "" Then
                              strR038015 = RsTemp.Fields("SIM01") '©“§∫¬Ωƒ∂
'                           Else
'                              strR038015 = RsTemp.Fields("st01")
                           End If
                        End If
                     End If
                     If strR038015 = "" Then strR038015 = strTemp(0) '¶π¨q¿≥∏”Run§£®Ï
                     '2021/5/6 END
                  End If
               End If
            End If
            strSql = "insert into R060308_3(R038001,R038002,R038003,R038004,R038005,R038006,R038007,R038008,R038009,R038010,R038011,R038012,R038013,ID,R038014,R038015) VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & strUserNum & "','" & ChangeWStringToTDateString("" & .Fields("C16")) & "','" & strR038015 & "') "
            'end 2007/3/27
            cnnConnection.Execute strSql
            'πw≥]µL∏ÍÆ∆
            ChkPro(3) = True
NextRecord:
            .MoveNext
            DoEvents
        Loop
    End With
Else
   Pro3 = False
   ChkPro(3) = False
   Exit Function
End If
CheckOC
If ChkPro(3) = False Then
   Pro3 = False
   ChkPro(3) = False
   Exit Function
End If

Pro3 = True
ChkPro(3) = True
End Function

Function Pro4() As Boolean
   'Pro4   ´»§·Æ◊•Û(•”Ω–§H)(•º¶¨§Â)
   '****************************************************
   '•N≤z§H¶W∫Ÿ, •”Ω–§H¶W∫Ÿ, ¶aß}, Æ◊•Û© ΩË¶W∫Ÿ, Æ◊•Û¶W∫Ÿ
   '≠Y§§§Â≥¯™Ì  §§-->≠^-->§È
   '≠Y≠^§Â≥¯™Ì  ≠^-->§È-->NULL
   '≠Y§È§Â≥¯™Ì  §È-->≠^-->NULL
   '****************************************************
   Select Case Val(txt1(14))
      Case 1      '§§§Â
           pub_QL05 = pub_QL05 & ";" & Label1(8) & "1.§§§Â" 'Add By Sindy 2010/12/8
           RefreshColData 1, "•”Ω–§H"
           If Option1(0).Value = True Then
              'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
              strSql = " SELECT ALL " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',NP08,Nvl(DECODE(PA09,'000',CPM03,CPM04),Nvl(CPM10,CPM13)),PA77,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(PA05,Nvl(PA06,PA07)),PA11,PA22,PA48,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=PA01 AND NP03=PA02 AND NP04=PA03 AND NP05=PA04 AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND NP01=CP09 AND NP07<>605 " & strSQL1 & StrSQL6
              strSql = strSql + " union all select " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',NP08,Nvl(DECODE(SP09,'000',CPM03,CPM04),Nvl(CPM10,CPM13)),SP27,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(SP05,Nvl(SP06,SP07)),SP11,'',SP29,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=SP01 AND NP03=SP02 AND NP04=SP03 AND NP05=SP04 AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND NP01=CP09  AND NP07<>605 " & strSQL2 & StrSQL6
           Else
              'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
              strSql = " SELECT ALL " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',NP09,Nvl(DECODE(PA09,'000',CPM03,CPM04),Nvl(CPM10,CPM13)),PA77,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(PA05,Nvl(PA06,PA07)),PA11,PA22,PA48,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=PA01 AND NP03=PA02 AND NP04=PA03 AND NP05=PA04 AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND NP01=CP09  AND NP07<>605 " & strSQL1 & StrSQL6
              strSql = strSql + " union all select " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',NP09,Nvl(DECODE(SP09,'000',CPM03,CPM04),Nvl(CPM10,CPM13)),SP27,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(SP05,Nvl(SP06,SP07)),SP11,'',SP29,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=SP01 AND NP03=SP02 AND NP04=SP03 AND NP05=SP04 AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND NP01=CP09  AND NP07<>605 " & strSQL2 & StrSQL6
           End If
      Case 2      '≠^§Â
           pub_QL05 = pub_QL05 & ";" & Label1(8) & "2.≠^§Â" 'Add By Sindy 2010/12/8
           RefreshColData 2, "•”Ω–§H"
           If Option1(0).Value = True Then
              'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
              strSql = " SELECT ALL " & m_ColCustName & "," & m_ColCustAdd & ",CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,NP08,Nvl(cpm10,CPM13),PA77,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(PA06,PA07),PA11,PA22,PA48,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=PA01 AND NP03=PA02 AND NP04=PA03 AND NP05=PA04 AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND NP01=CP09 AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND NP07<>605 " & strSQL1 & StrSQL6
              strSql = strSql + " union all select " & m_ColCustName & "," & m_ColCustAdd & ",CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,NP08,Nvl(CPM10,CPM13),SP27,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(SP06,SP07),SP11,'',SP29,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=SP01 AND NP03=SP02 AND NP04=SP03 AND NP05=SP04 AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND NP01=CP09  AND NP07<>605 " & strSQL2 & StrSQL6
           Else
              'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
              strSql = " SELECT ALL " & m_ColCustName & "," & m_ColCustAdd & ",CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,NP09,Nvl(CPM10,CPM13),PA77,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(PA06,PA07),PA11,PA22,PA48,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=PA01 AND NP03=PA02 AND NP04=PA03 AND NP05=PA04 AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+))  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND NP01=CP09  AND NP07<>605 " & strSQL1 & StrSQL6
              strSql = strSql + " union all select " & m_ColCustName & "," & m_ColCustAdd & ",CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,NP09,Nvl(CPM10,CPM13),SP27,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(SP06,SP07),SP11,'',SP29,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=SP01 AND NP03=SP02 AND NP04=SP03 AND NP05=SP04 AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND NP01=CP09  AND NP07<>605 " & strSQL2 & StrSQL6
           End If
      Case 3      '§È§Â
           pub_QL05 = pub_QL05 & ";" & Label1(8) & "3.§È§Â" 'Add By Sindy 2010/12/8
           RefreshColData 3, "•”Ω–§H"
           If Option1(0).Value = True Then
              'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
              strSql = " SELECT ALL " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',NP08,Nvl(CPm13,CPM10),PA77,NP02||'-'||NP03||'-'||NP04||'-'||NP05,nvl(PA07,pa06),PA11,PA22,PA48,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=PA01 AND NP03=PA02 AND NP04=PA03 AND NP05=PA04 AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+))  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND NP01=CP09 AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND NP07<>605 " & strSQL1 & StrSQL6
              strSql = strSql + " union all select " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',NP08,Nvl(CPM13,CPM10),SP27,NP02||'-'||NP03||'-'||NP04||'-'||NP05,nvl(SP07,sp06),SP11,'',SP29,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=SP01 AND NP03=SP02 AND NP04=SP03 AND NP05=SP04 AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND NP01=CP09  AND NP07<>605 " & strSQL2 & StrSQL6
           Else
              'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
              strSql = " SELECT ALL " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',NP09,Nvl(CPM13,CPM10),PA77,NP02||'-'||NP03||'-'||NP04||'-'||NP05,nvl(PA07,pa06),PA11,PA22,PA48,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=PA01 AND NP03=PA02 AND NP04=PA03 AND NP05=PA04 AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND NP01=CP09  AND NP07<>605 " & strSQL1 & StrSQL6
              strSql = strSql + " union all select " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',NP09,Nvl(CPM13,CPM10),SP27,NP02||'-'||NP03||'-'||NP04||'-'||NP05,nvl(SP07,sp06),SP11,'',SP29,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=SP01 AND NP03=SP02 AND NP04=SP03 AND NP05=SP04 AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND NP01=CP09  AND NP07<>605 " & strSQL2 & StrSQL6
           End If
      Case Else
   End Select
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'edit by nickc 2007/02/08
   'k = 0
   
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           Do While .EOF = False
               For i = 0 To 23
                   strTemp(i) = ChgSQL(CheckStr(.Fields(i)))
               Next i
               If Len(strTemp(4)) = 0 And Len(strTemp(5)) = 0 And Len(strTemp(6)) = 0 And Len(strTemp(7)) = 0 And Len(strTemp(8)) = 0 Then
                   For i = 4 To 16
                       strTemp(i) = strTemp(i + 5)
                   Next i
               Else
                   For i = 9 To 16
                       strTemp(i) = strTemp(i + 5)
                   Next i
               End If
               Select Case Val(txt1(14))
               Case 1      '§§§Â
                     strTemp(9) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(9)))
               Case 2, 3
                     strTemp(9) = Format(ChangeWStringToWDateString(strTemp(9)), "mmm DD,YYYY")
               Case Else
               End Select
   
               strSql = "insert into R060308_4 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & ChgSQL(strTemp(16)) & "','" & strUserNum & "') "
               cnnConnection.Execute strSql
               .MoveNext
               DoEvents
           Loop
       End With
   Else
      Pro4 = False
      ChkPro(4) = False
      Exit Function
   End If
   CheckOC
   Pro4 = True
   ChkPro(4) = True
End Function

Function Pro5() As Boolean
   'Pro5   ´»§·Æ◊•Û(•N≤z§H)(•º¶¨§Â)
   '****************************************************
   '•N≤z§H¶W∫Ÿ, •”Ω–§H¶W∫Ÿ, ¶aß}, Æ◊•Û© ΩË¶W∫Ÿ, Æ◊•Û¶W∫Ÿ
   '≠Y§§§Â≥¯™Ì  §§-->≠^-->§È
   '≠Y≠^§Â≥¯™Ì  ≠^-->§È-->NULL
   '≠Y§È§Â≥¯™Ì  §È-->≠^-->NULL
   '****************************************************
   Select Case Val(txt1(14))
   Case 1      '§§§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "1.§§§Â" 'Add By Sindy 2010/12/8
        RefreshColData 1, "•N≤z§H"
        If Option1(0).Value = True Then
           strSql = " SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',NP08,Nvl(DECODE(PA09,'000',CPM03,CPM04),Nvl(CPM10,CPM13)),PA77,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(PA05,Nvl(PA06,PA07)),PA11,PA22,PA48,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CASEPROPERTYMAP, Fagent WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA26,9,1))=FA02(+) AND NP01=CP09(+) AND NP07<>605 " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',NP08,Nvl(DECODE(SP09,'000',CPM03,CPM04),Nvl(CPM10,CPM13)),SP27,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(SP05,Nvl(SP06,SP07)),SP11,'',SP29,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP, Fagent WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP08,9,1))=FA02(+) AND NP01=CP09(+)  AND NP07<>605 " & strSQL2 & StrSQL6
        Else
           strSql = "SELECT " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',NP09,Nvl(DECODE(PA09,'000',CPM03,CPM04),Nvl(CPM10,CPM13)),PA77,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(PA05,Nvl(PA06,PA07)),PA11,PA22,CU04,PA48,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CASEPROPERTYMAP,FAGENT,CUSTOMER WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND NP01=CP09(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+)  AND NP07<>605 " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',NP09,Nvl(DECODE(SP09,'000',CPM03,CPM04),Nvl(CPM10,CPM13)),SP27,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(SP05,Nvl(SP06,SP07)),SP11,'',CU04,SP29,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,FAGENT,CUSTOMER WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND NP01=CP09(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+)  AND NP07<>605 " & strSQL2 & StrSQL6
        End If
   Case 2      '≠^§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "2.≠^§Â" 'Add By Sindy 2010/12/8
        RefreshColData 2, "•N≤z§H"
        If Option1(0).Value = True Then
           strSql = "SELECT " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,fa22||rtrim(' '||fa70) FA22,NP08,Nvl(CPM10,CPM13),PA77,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(PA06,PA07),PA11,PA22,CU05||CU88||CU89||CU90,PA48,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND NP01=CP09(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+)  AND NP07<>605 " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,fa22||rtrim(' '||fa70) FA22,NP08,Nvl(CPM10,CPM13),SP27,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(SP06,SP07),SP11,'',CU05||CU88||CU89||CU90,SP29,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND NP01=CP09(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+)  AND NP07<>605 " & strSQL2 & StrSQL6
        Else
           strSql = "SELECT " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,fa22||rtrim(' '||fa70) FA22,NP09,Nvl(CPM10,CPM13),PA77,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(PA06,PA07),PA11,PA22,CU05||CU88||CU89||CU90,PA48,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND NP01=CP09(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+)  AND NP07<>605  " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,fa22||rtrim(' '||fa70) FA22,NP09,Nvl(CPM10,CPM13),SP27,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(SP06,SP07),SP11,'',CU05||CU88||CU89||CU90,SP29,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND NP01=CP09(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+)  AND NP07<>605 " & strSQL2 & StrSQL6
        End If
   Case 3      '§È§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "3.§È§Â" 'Add By Sindy 2010/12/8
        RefreshColData 3, "•N≤z§H"
        If Option1(0).Value = True Then
           strSql = "SELECT " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',NP08,Nvl(CPM13,CPM10),PA77,NP02||'-'||NP03||'-'||NP04||'-'||NP05,nvl(PA07,pa06),PA11,PA22,CU06,PA48,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND NP01=CP09(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+)  AND NP07<>605 " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',NP08,Nvl(CPm13,CPM10),SP27,NP02||'-'||NP03||'-'||NP04||'-'||NP05,nvl(SP07,sp06),SP11,'',CU06,SP29,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND NP01=CP09(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+)  AND NP07<>605 " & strSQL2 & StrSQL6
        Else
           strSql = "SELECT " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',NP09,Nvl(CPM13,CPM10),PA77,NP02||'-'||NP03||'-'||NP04||'-'||NP05,nvl(PA07,pa06),PA11,PA22,CU06,PA48,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND NP01=CP09(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+)  AND NP07<>605 " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',NP09,Nvl(CPm13,CPM10),SP27,NP02||'-'||NP03||'-'||NP04||'-'||NP05,nvl(SP07,sp06),SP11,'',CU06,SP29,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND NP01=CP09(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+)  AND NP07<>605 " & strSQL2 & StrSQL6
        End If
   Case Else
   End Select
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'edit by nickc 2007/02/08
   'k = 0
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           Do While .EOF = False
               For i = 0 To 22
                   strTemp(i) = ChgSQL(CheckStr(.Fields(i)))
               Next i
               If Len(strTemp(4)) = 0 And Len(strTemp(5)) = 0 And Len(strTemp(6)) = 0 And Len(strTemp(7)) = 0 And Len(strTemp(8)) = 0 Then
                   For i = 4 To 17
                       strTemp(i) = strTemp(i + 5)
                   Next i
               Else
                   For i = 9 To 17
                       strTemp(i) = strTemp(i + 5)
                   Next i
               End If
               Select Case Val(txt1(14))
               Case 1      '§§§Â
                               strTemp(9) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(9)))
               Case 2, 3
                               strTemp(9) = Format(ChangeWStringToWDateString(strTemp(9)), "mmm DD,YYYY")
               Case Else
               End Select
               strSql = "insert into R060308_5 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "','" & strUserNum & "') "
               cnnConnection.Execute strSql
               .MoveNext
               DoEvents
           Loop
       End With
   Else
      Pro5 = False
      ChkPro(5) = False
      Exit Function
   End If
   CheckOC
   Pro5 = True
   ChkPro(5) = True
End Function

Function Pro6() As Boolean
   '****************************************************
   '•N≤z§H¶W∫Ÿ, •”Ω–§H¶W∫Ÿ, ¶aß}, Æ◊•Û© ΩË¶W∫Ÿ, Æ◊•Û¶W∫Ÿ
   '≠Y§§§Â≥¯™Ì  §§-->≠^-->§È
   '≠Y≠^§Â≥¯™Ì  ≠^-->§È-->NULL
   '≠Y§È§Â≥¯™Ì  §È-->≠^-->NULL
   '****************************************************
   
   'Pro6   ¶~∂O(•”Ω–§H)(•º¶¨§Â)
   Select Case Val(txt1(14))
   Case 1      '§§§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "1.§§§Â" 'Add By Sindy 2010/12/8
        RefreshColData 1, "•”Ω–§H"
        If Option1(0).Value = True Then
           'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
           strSql = "SELECT " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',NP08,PA77,PA48,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(PA05,Nvl(PA06,PA07)),PA11,PA22,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND NP01=CP09(+) AND NP07=605 " & strSQL1 & StrSQL6
           strSql = strSql + " union  SELECT " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',NP08,SP27,SP29,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(SP05,Nvl(SP06,SP07)),SP11,'',CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND NP01=CP09(+) AND NP07=605 " & strSQL2 & StrSQL6
        Else
           'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
           strSql = "SELECT " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',NP09,PA77,PA48,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(PA05,Nvl(PA06,PA07)),PA11,PA22,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND NP01=CP09(+) AND NP07=605 " & strSQL1 & StrSQL6
           strSql = strSql + " union  SELECT " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',NP09,SP27,SP29,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(SP05,Nvl(SP06,SP07)),SP11,'',CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND NP01=CP09(+) AND NP07=605 " & strSQL2 & StrSQL6
        End If
   Case 2      '≠^§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "2.≠^§Â" 'Add By Sindy 2010/12/8
        RefreshColData 2, "•”Ω–§H"
        If Option1(0).Value = True Then
           'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
           strSql = "SELECT " & m_ColCustName & "," & m_ColCustAdd & ",CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,NP08,PA77,PA48,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(PA06,PA07),PA11,PA22,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND NP01=CP09(+) AND NP07=605 " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColCustName & "," & m_ColCustAdd & ",CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,NP08,SP27,SP29,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(SP06,SP07),SP11,'',CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND NP01=CP09(+) AND NP07=605 " & strSQL2 & StrSQL6
        Else
           'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
           strSql = "SELECT " & m_ColCustName & "," & m_ColCustAdd & ",CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,NP09,PA77,PA48,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(PA06,PA07),PA11,PA22,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND NP01=CP09(+) AND NP07=605 " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColCustName & "," & m_ColCustAdd & ",CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,NP09,SP27,SP29,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(SP06,SP07),SP11,'',CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND NP01=CP09(+) AND NP07=605 " & strSQL2 & StrSQL6
        End If
   Case 3      '§È§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "3.§È§Â" 'Add By Sindy 2010/12/8
        RefreshColData 3, "•”Ω–§H"
        If Option1(0).Value = True Then
           'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
           strSql = "SELECT " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',NP08,PA77,PA48,NP02||'-'||NP03||'-'||NP04||'-'||NP05,nvl(PA07,pa06),PA11,PA22,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND NP01=CP09(+) AND NP07=605 " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',NP08,SP27,SP29,NP02||'-'||NP03||'-'||NP04||'-'||NP05,nvl(SP07,sp06),SP11,'',CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND NP01=CP09(+) AND NP07=605 " & strSQL2 & StrSQL6
        Else
           'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
           strSql = "SELECT " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',NP09,PA77,PA48,NP02||'-'||NP03||'-'||NP04||'-'||NP05,nvl(PA07,pa06),PA11,PA22,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND NP01=CP09(+) AND NP07=605 " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',NP09,SP27,SP29,NP02||'-'||NP03||'-'||NP04||'-'||NP05,nvl(SP07,sp06),SP11,'',CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND NP01=CP09(+) AND NP07=605 " & strSQL2 & StrSQL6
        End If
   Case Else
   End Select
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'edit by nickc 2007/02/08
   'k = 0
   
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           Do While .EOF = False
               For i = 0 To 20
                   strTemp(i) = ChgSQL(CheckStr(.Fields(i)))
               Next i
               If Len(strTemp(4)) = 0 And Len(strTemp(5)) = 0 And Len(strTemp(6)) = 0 And Len(strTemp(7)) = 0 And Len(strTemp(8)) = 0 Then
                   For i = 4 To 15
                       strTemp(i) = strTemp(i + 5)
                   Next i
               Else
                   For i = 9 To 15
                       strTemp(i) = strTemp(i + 5)
                   Next i
               End If
               Select Case Val(txt1(14))
               Case 1      '§§§Â
                               strTemp(9) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(9)))
               Case 2, 3
                             strTemp(9) = Format(ChangeWStringToWDateString(strTemp(9)), "mmm DD,YYYY")
               Case Else
               End Select
               strSql = "insert into R060308_6 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & (strTemp(4)) & "','" & (strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & (strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & strUserNum & "') "
               cnnConnection.Execute strSql
               .MoveNext
               DoEvents
           Loop
       End With
   Else
      Pro6 = False
      ChkPro(6) = False
      Exit Function
   End If
   CheckOC
   Pro6 = True
   ChkPro(6) = True
End Function

Function Pro7() As Boolean
   'Pro7   ¶~∂O(•N≤z§H)(•º¶¨§Â)
   '****************************************************
   '•N≤z§H¶W∫Ÿ, •”Ω–§H¶W∫Ÿ, ¶aß}, Æ◊•Û© ΩË¶W∫Ÿ, Æ◊•Û¶W∫Ÿ
   '≠Y§§§Â≥¯™Ì  §§-->≠^-->§È
   '≠Y≠^§Â≥¯™Ì  ≠^-->§È-->NULL
   '≠Y§È§Â≥¯™Ì  §È-->≠^-->NULL
   '****************************************************
   
   Select Case Val(txt1(14))
   Case 1      '§§§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "1.§§§Â" 'Add By Sindy 2010/12/8
        RefreshColData 1, "•N≤z§H"
        If Option1(0).Value = True Then
           strSql = "SELECT " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',NP08,PA77,PA48,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(PA05,Nvl(PA06,PA07)),PA11,PA22,CU04,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,FAGENT,CUSTOMER WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND NP01=CP09(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND NP07=605 " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',NP08,SP27,SP29,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(SP05,Nvl(SP06,SP07)),SP11,'',CU04,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,FAGENT,CUSTOMER WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,1,8),'','0',SUBSTR(SP26,1,8))=FA02(+) AND NP01=CP09(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND NP07=605" & strSQL2 & StrSQL6
        Else
           strSql = "SELECT " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',NP09,PA77,PA48,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(PA05,Nvl(PA06,PA07)),PA11,PA22,CU04,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,FAGENT,CUSTOMER WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND NP01=CP09(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND NP07=605 " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',NP09,SP27,SP29,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(SP05,Nvl(SP06,SP07)),SP11,'',CU04,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,FAGENT,CUSTOMER WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,1,8),'','0',SUBSTR(SP26,1,8))=FA02(+) AND NP01=CP09(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND NP07=605" & strSQL2 & StrSQL6
        End If
   Case 2      '≠^§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "2.≠^§Â" 'Add By Sindy 2010/12/8
        RefreshColData 2, "•N≤z§H"
        If Option1(0).Value = True Then
           strSql = "SELECT " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,fa22||rtrim(' '||fa70) FA22,NP08,PA77,PA48,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(PA06,PA07),PA11,PA22,CU05||CU88||CU89||CU90,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND NP01=CP09(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND NP07=605 " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,fa22||rtrim(' '||fa70) FA22,NP08,SP27,SP29,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(SP06,SP07),SP11,'',CU05||CU88||CU89||CU90,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,1,8),'','0',SUBSTR(SP26,1,8))=FA02(+) AND NP01=CP09(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND NP07=605 " & strSQL2 & StrSQL6
        Else
           strSql = "SELECT " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,fa22||rtrim(' '||fa70) FA22,NP09,PA77,PA48,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(PA06,PA07),PA11,PA22,CU05||CU88||CU89||CU90,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND NP01=CP09(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND NP07=605 " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,fa22||rtrim(' '||fa70) FA22,NP09,SP27,SP29,NP02||'-'||NP03||'-'||NP04||'-'||NP05,Nvl(SP06,SP07),SP11,'',CU05||CU88||CU89||CU90,CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,1,8),'','0',SUBSTR(SP26,1,8))=FA02(+) AND NP01=CP09(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND NP07=605 " & strSQL2 & StrSQL6
        End If
   Case 3      '§È§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "3.§È§Â" 'Add By Sindy 2010/12/8
        RefreshColData 3, "•N≤z§H"
        If Option1(0).Value = True Then
           '´»§·¶W∫Ÿ(§È=>≠^=>"")
           strSql = "SELECT " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',NP08,PA77,PA48,NP02||'-'||NP03||'-'||NP04||'-'||NP05,nvl(PA07,pa06),PA11,PA22,NVL(CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND NP01=CP09(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND NP07=605 " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',NP08,SP27,SP29,NP02||'-'||NP03||'-'||NP04||'-'||NP05,nvl(SP07,sp06),SP11,'',NVL(CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND DECODE(SUBSTR(SP26,1,8),'','0',SUBSTR(SP26,1,8))=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND NP01=CP09(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND NP07=605 " & strSQL2 & StrSQL6
        Else
           '´»§·¶W∫Ÿ(§È=>≠^=>"")
           strSql = "SELECT " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',NP09,PA77,PA48,NP02||'-'||NP03||'-'||NP04||'-'||NP05,nvl(PA07,pa06),PA11,PA22,NVL(CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND NP01=CP09(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND NP07=605 " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',NP09,SP27,SP29,NP02||'-'||NP03||'-'||NP04||'-'||NP05,nvl(SP07,sp06),SP11,'',NVL(CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CP09,CP27 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND DECODE(SUBSTR(SP26,1,8),'','0',SUBSTR(SP26,1,8))=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND NP01=CP09(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND NP07=605 " & strSQL2 & StrSQL6
        End If
   Case Else
   End Select
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'edit by nickc 2007/02/08
   'k = 0
   
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           Do While .EOF = False
               For i = 0 To 21
                   strTemp(i) = ChgSQL(CheckStr(.Fields(i)))
               Next i
               If Len(strTemp(4)) = 0 And Len(strTemp(5)) = 0 And Len(strTemp(6)) = 0 And Len(strTemp(7)) = 0 And Len(strTemp(8)) = 0 Then
                   For i = 4 To 16
                       strTemp(i) = strTemp(i + 5)
                   Next i
               Else
                   For i = 9 To 16
                       strTemp(i) = strTemp(i + 5)
                   Next i
               End If
               Select Case Val(txt1(14))
               Case 1      '§§§Â
                               strTemp(9) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(9)))
               Case 2, 3
                               strTemp(9) = Format(ChangeWStringToWDateString(strTemp(9)), "mmm DD,YYYY")
               Case Else
               End Select
               strSql = "insert into R060308_7 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & ChgSQL(strTemp(16)) & "','" & strUserNum & "') "
               cnnConnection.Execute strSql
               .MoveNext
               DoEvents
           Loop
       End With
   Else
      Pro7 = False
      ChkPro(7) = False
      Exit Function
   End If
   CheckOC
   Pro7 = True
   ChkPro(7) = True
End Function

Function Pro8() As Boolean
'Pro8  ∫ﬁ®Ó§H(§w¶¨§Â)
Dim strNote As String 'Add By Sindy 2015/12/28
   
   'πw≥]µL∏ÍÆ∆
   ChkPro(8) = False
   
   '•N≤z§H≠^-->§§-->§È
   'modify by sonia 2014/11/14 ≠Y≠”Æ◊™∫FCP¶~∂O¶€∞ •N√∫PA70§ŒFCPª‚√“¶€∞ •N√∫PA71•º≥]©w,•N≤z§H©Œ´»§·¶≥≥]©w§]≠n±a
   'strSql = "SELECT '' X00,CP06,CP07,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X03,NVL(PA05,NVL(PA06,PA07)) X04,PA11,DECODE(PA09,'000',CPM03,CPM04) X06,'©”øÏ§H' X07,ST02,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)) X09,PA77,NVL(NA03,NA04) X11,PA71,PA70,' ' X14,CP64,CP09,CP27,CP10,PA76,PA75,PA26 FROM CASEPROGRESS,PATENT,STAFF,CASEPROPERTYMAP,FAGENT,NATION WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP14=ST01(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND FA10=NA01(+) " & strSQL1 & StrSQL6
   '2014/12/1 modify by sonia ∂»ª‚√“601§Œ¶~∂O605§~≠n±a•X¶€∞ •N√∫ƒÊ
   'Modify By Sindy 2015/12/28 + ,CP01,CP43
   'Modified by Lydia 2019/05/31 ∞Íƒy≠≠®Ó5≠”¶r
   'Modified by Lydia 2019/06/24 +FCPπÍºf¶€∞ •N√∫C19
   'strSql = "SELECT '' X00,CP06,CP07,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X03,NVL(PA05,NVL(PA06,PA07)) X04,PA11,DECODE(PA09,'000',CPM03,CPM04) X06,'©”øÏ§H' X07,ST02,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)) X09,PA77,SUBSTR(NVL(NA03,NA04),1,5) X11,decode(cp10,'601',nvl(PA71,nvl(fa42,cu75)),null) PA71,decode(cp10,'605',nvl(PA70,nvl(fa41,cu74)),null) PA70,' ' X14,CP64,CP09,CP27,CP10,PA76,PA75,PA26,CP01,CP43 FROM CASEPROGRESS,PATENT,STAFF,CASEPROPERTYMAP,FAGENT,NATION,CUSTOMER WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP14=ST01(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND FA10=NA01(+) " & strSQL1 & StrSQL6
   'Modified by Morgan 2024/1/13 ¶~∂O¶€∞ •N√∫Y§~≈„•‹(N™Ìª‚√“´·§£ƒÚøÏ,≠Yª›≠n¿≥•t¶CƒÊ¶Ï)--Bobbie nvl(fa41,cu74)->decode(nvl(fa41,cu74),'Y','Y')
   strSql = "SELECT '' X00,CP06,CP07,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X03,NVL(PA05,NVL(PA06,PA07)) X04,PA11,DECODE(PA09,'000',CPM03,CPM04) X06,'©”øÏ§H' X07,S1.ST02,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)) X09,PA77,SUBSTR(NVL(NA03,NA04),1,5) X11,decode(cp10,'601',nvl(PA71,nvl(fa42,cu75)),null) PA71,decode(cp10,'605',nvl(PA70,decode(nvl(fa41,cu74),'Y','Y')),null) PA70,' ' X14,CP64,CP09,CP27,CP10,PA76,PA75,PA26,CP01,CP43,DECODE(CP10,'416',NVL(PA69,NVL(FA124,CU177)),NULL) AS C19 " & _
               "FROM CASEPROGRESS,PATENT,STAFF S1,CASEPROPERTYMAP,FAGENT,NATION,CUSTOMER WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND S1.ST01(+)=CP14 AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND FA10=NA01(+) " & _
               strSQL1 & StrSQL6
   'end 2014/11/14
   'Modify By Sindy 2015/12/28 + ,CP01,CP43
   'Modified by Lydia 2019/06/24 +FCPπÍºf¶€∞ •N√∫C19
   strSql = strSql + " union select '',CP06,CP07,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),SP11,DECODE(SP09,'000',CPM03,CPM04),'©”øÏ§H',S1.ST02,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)),SP27,SUBSTR(NVL(NA03,NA04),1,5),' ',' ',' ',CP64,CP09,CP27,CP10,NULL,SP26,SP08,CP01,CP43,'' as C19 " & _
               "FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,CASEPROPERTYMAP,FAGENT,NATION WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND S1.ST01(+)=CP14 AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND FA10=NA01(+) " & _
               strSQL2 & StrSQL6
   'Add By Sindy 2021/4/15 ºW•[∑sÆ◊¬Ωƒ∂©“§∫§uµ{Æv©”øÏ¥¡≠≠
   'Modified by Morgan 2024/1/13 ¶~∂O¶€∞ •N√∫Y§~≈„•‹(N™Ìª‚√“´·§£ƒÚøÏ,≠Yª›≠n¿≥•t¶CƒÊ¶Ï)--Bobbie nvl(fa41,cu74)->decode(nvl(fa41,cu74),'Y','Y')
   strSql = strSql + " union SELECT '' X00,CP06,CP07,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X03,NVL(PA05,NVL(PA06,PA07)) X04,PA11,DECODE(PA09,'000',CPM03,CPM04) X06,'©”øÏ§H' X07,decode(S2.ST02,NULL,S1.ST02,S2.ST02),NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)) X09,PA77,SUBSTR(NVL(NA03,NA04),1,5) X11,decode(cp10,'601',nvl(PA71,nvl(fa42,cu75)),null) PA71,decode(cp10,'605',nvl(PA70,decode(nvl(fa41,cu74),'Y','Y')),null) PA70,' ' X14,CP64,CP09,CP27,CP10,PA76,PA75,PA26,CP01,CP43,DECODE(CP10,'416',NVL(PA69,NVL(FA124,CU177)),NULL) AS C19 " & _
               "FROM CASEPROGRESS,PATENT,STAFF S1,CASEPROPERTYMAP,FAGENT,NATION,CUSTOMER,STAFF_IDMAP,STAFF S2 WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND S1.ST01(+)=CP14 AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND FA10=NA01(+) " & _
               "AND SIM02(+)=CP14 AND S2.ST01(+)=SIM01 " & _
               "AND SUBSTRB(S1.ST15,1,1)='F' AND (S2.ST15 IS NULL OR S1.ST15='F52') " & _
               "AND EXISTS(select ts2.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1') " & _
               strSQL1 & strSQL8 & " AND CP10='201' AND exists(select * from engineerprogress x where x.ep02=cp09 and ep09||" & IIf(strSrvDate(1) >= FCPÆ÷ßπ§ÈßÔ•ŒEP39, "ep39", "ep33") & " is null)"
   strSql = strSql + " union select '',CP06,CP07,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),SP11,DECODE(SP09,'000',CPM03,CPM04),'©”øÏ§H',decode(S2.ST02,NULL,S1.ST02,S2.ST02),NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)),SP27,SUBSTR(NVL(NA03,NA04),1,5),' ',' ',' ',CP64,CP09,CP27,CP10,NULL,SP26,SP08,CP01,CP43,'' as C19 " & _
               "FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,CASEPROPERTYMAP,FAGENT,NATION,STAFF_IDMAP,STAFF S2 WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND S1.ST01(+)=CP14 AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND FA10=NA01(+) " & _
               "AND SIM02(+)=CP14 AND S2.ST01(+)=SIM01 " & _
               "AND SUBSTRB(S1.ST15,1,1)='F' AND (S2.ST15 IS NULL OR S1.ST15='F52') " & _
               "AND EXISTS(select ts2.st01 from staff ts1,staff ts2 where ts1.st01=cp14 and ts1.st26 is not null and ts2.st26=ts1.st26 and ts2.st04='1') " & _
               strSQL2 & strSQL8 & " AND CP10='201' AND exists(select * from engineerprogress x where x.ep02=cp09 and ep09||" & IIf(strSrvDate(1) >= FCPÆ÷ßπ§ÈßÔ•ŒEP39, "ep39", "ep33") & " is null)"
   '2021/4/15 END
   
   'Modify by Morgan 2006/2/14 •[∫ﬁ®Ó§H±¯•Ûªy™k
   'strSQL = "select * from (" & strSQL & ") X where " & " not exists(select * from caseprogress a where a.cp43= X.cp09 and a.cp10='907')"
   '¶~∂O∫ﬁ®Ó§H∂∂ß«:pa76->pa75->pa26.cu96->pa26(PA76•iØ‡¨O´»§·Ωs∏π)
   'Modified by Lydia 2017/02/13 +FMP∫ﬁ®Ó§H
   'Modified by Morgan 2020/5/12 ¶~∂O∫ﬁ®Ó§HßÔßÏÆ◊•Û•N≤z§H™∫∫ﬁ®Ó§H(§£•≤¶“º{¨Oß_¶≥¶~∂O•N≤z§H)
   'If strSrvDate(1) < FMP∫ﬁ®Ó§H±“•Œ§È Then
   '     strSql = "SELECT Y.*,NA16 FROM (" & _
           " select X.*,NVL(NVL(F1.FA10,C1.CU10),NVL(F2.FA10,NVL(F4.FA10,C2.CU10))) FA10" & _
           " from (" & strSql & ") X, FAGENT F1, FAGENT F2,FAGENT F4,CUSTOMER C1, CUSTOMER C2" & _
           " WHERE CP10='605' AND F1.FA01(+)=SUBSTR(PA76,1,8) AND F1.FA02(+)=SUBSTR(PA76||'0',9,1)" & _
           " AND C1.CU01(+)=SUBSTR(PA76,1,8) AND C1.CU02(+)=SUBSTR(PA76||'0',9,1)" & _
           " AND F2.FA01(+)=SUBSTR(PA75,1,8) AND F2.FA02(+)=SUBSTR(PA75||'0',9,1)" & _
           " AND C2.CU01(+)=SUBSTR(PA26,1,8) AND C2.CU02(+)=SUBSTR(PA26||'0',9,1)" & _
           " AND F4.FA01(+)=SUBSTR(C2.CU96,1,8) AND F4.FA02(+)=SUBSTR(C2.CU96||'0',9,1)" & _
           " Union All" & _
           " select X.*,NVL(FA10,CU10) FA10" & _
           " from (" & strSql & ") X, FAGENT,CUSTOMER" & _
           " WHERE CP10<>'605' AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75||'0',9,1)" & _
           " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26||'0',9,1)" & _
           ") Y, NATION WHERE NA01(+)=FA10"
   'Else
   '     strSql = "SELECT Y.*,DECODE(CP01,'P',NVL(NA79,NA16),'PS',NVL(NA79,NA16),NA16) NA16 FROM (" & _
           " select X.*,NVL(NVL(F1.FA10,C1.CU10),NVL(F2.FA10,NVL(F4.FA10,C2.CU10))) FA10" & _
           " from (" & strSql & ") X, FAGENT F1, FAGENT F2,FAGENT F4,CUSTOMER C1, CUSTOMER C2" & _
           " WHERE CP10='605' AND F1.FA01(+)=SUBSTR(PA76,1,8) AND F1.FA02(+)=SUBSTR(PA76||'0',9,1)" & _
           " AND C1.CU01(+)=SUBSTR(PA76,1,8) AND C1.CU02(+)=SUBSTR(PA76||'0',9,1)" & _
           " AND F2.FA01(+)=SUBSTR(PA75,1,8) AND F2.FA02(+)=SUBSTR(PA75||'0',9,1)" & _
           " AND C2.CU01(+)=SUBSTR(PA26,1,8) AND C2.CU02(+)=SUBSTR(PA26||'0',9,1)" & _
           " AND F4.FA01(+)=SUBSTR(C2.CU96,1,8) AND F4.FA02(+)=SUBSTR(C2.CU96||'0',9,1)" & _
           " Union All" & _
           " select X.*,NVL(FA10,CU10) FA10" & _
           " from (" & strSql & ") X, FAGENT,CUSTOMER" & _
           " WHERE CP10<>'605' AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75||'0',9,1)" & _
           " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26||'0',9,1)" & _
           ") Y, NATION WHERE NA01(+)=FA10"
   'End If
   strSql = "SELECT Y.*,DECODE(CP01,'P',NVL(NA79,NA16),'PS',NVL(NA79,NA16),NA16) NA16 FROM (" & _
           " select X.*,NVL(FA10,CU10) FA10" & _
           " from (" & strSql & ") X, FAGENT,CUSTOMER" & _
           " WHERE FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75||'0',9,1)" & _
           " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26||'0',9,1)" & _
           ") Y, NATION WHERE NA01(+)=FA10"
   'end 2020/5/12
   'end 2017/02/13
   
   If txt1(9).Text <> "" Then
      'Modified by Lydia 2017/02/13 +FMP∫ﬁ®Ó§H
      If strSrvDate(1) < FMP∫ﬁ®Ó§H±“•Œ§È Then
         strSql = strSql & " AND NA16 ='" & txt1(9).Text & "'"
      Else
         strSql = strSql & " AND DECODE(CP01,'P',NVL(NA79,NA16),'PS',NVL(NA79,NA16),NA16) ='" & txt1(9).Text & "'"
      End If
      'end 2017/02/13
      
      pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(9) & lbl1(2) 'Add By Sindy 2010/12/8
   End If
   '2006/2/14 END

   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'edit by nickc 2007/02/08
   'k = 0
   
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           Do While .EOF = False
               For i = 0 To 15
                   strTemp(i) = ChgSQL(CheckStr(.Fields(i)))
               Next i
               
               If strTemp(1) < strSrvDate(1) And Len(strTemp(1)) = 8 Then
                   strTemp(1) = "*" & ChangeWStringToTDateString(strTemp(1))
               Else
                   If strTemp(1) = strSrvDate(1) And Len(strTemp(1)) = 8 Then
                       strTemp(1) = "V" & ChangeWStringToTDateString(strTemp(1))
                   Else
                       strTemp(1) = ChangeWStringToTDateString(strTemp(1))
                   End If
               End If
               
               strTemp(2) = ChangeWStringToTDateString(strTemp(2))
               
'Modify by Morgan 2006/2/9 ßÔ•—§W≠±ªy™k±±®Ó
'               '®˙±oFCP∫ﬁ®Ó§H
               strTemp(0) = "" & .Fields("NA16")
'2006/2/9 END
               
               'Modify By Sindy 2015/12/28 ≥∆µ˘
               strNote = PUB_GetFCPAddQuyNotes(.Fields("cp01"), .Fields("CP09"), .Fields("CP10"), "" & .Fields("cp43"))
               If strNote <> "" Then
                  strTemp(15) = strNote & strTemp(15)
               End If
               '2015/12/28 END
               
               '®˙±oÆ÷ΩZ§H
               strTemp(14) = GetPrjSalesNM(Get060308_1(CheckStr(.Fields(16))))

               strTemp(16) = "" & .Fields("C19") 'Added by Lydia 2019/06/24 FCPπÍºf¶€∞ •N√∫
               
               'Modified by Lydia 2019/06/24
               'strSql = "insert into R060308_1(R036001,R036002,R036003,R036004,R036005,R036006,R036007,R036008,R036009,R036010,R036011,R036012,R036013,R036014,R036015,R036016,ID) VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & strUserNum & "') "
               strSql = "insert into R060308_1(R036001,R036002,R036003,R036004,R036005,R036006,R036007,R036008,R036009,R036010,R036011,R036012,R036013,R036014,R036015,R036016,ID,R036018) " & _
                           "VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & strUserNum & "','" & ChgSQL(strTemp(16)) & "') "
               cnnConnection.Execute strSql
               '¶≥∏ÍÆ∆
               ChkPro(8) = True
NextRecord:
               .MoveNext
               DoEvents
           Loop
       End With
   Else
      Pro8 = False
      ChkPro(8) = False
      Exit Function
   End If
   CheckOC
   If ChkPro(8) = False Then
      Pro8 = False
      ChkPro(8) = False
      Exit Function
   End If
   
   Pro8 = True
   ChkPro(8) = True
End Function

Function Pro9() As Boolean
   'Pro9   ´»§·Æ◊•Û(•”Ω–§H)(§w¶¨§Â)
   '****************************************************
   '•N≤z§H¶W∫Ÿ, •”Ω–§H¶W∫Ÿ, ¶aß}, Æ◊•Û© ΩË¶W∫Ÿ, Æ◊•Û¶W∫Ÿ
   '≠Y§§§Â≥¯™Ì  §§-->≠^-->§È
   '≠Y≠^§Â≥¯™Ì  ≠^-->§È-->NULL
   '≠Y§È§Â≥¯™Ì  §È-->≠^-->NULL
   '****************************************************
   
   Select Case Val(txt1(14))
   Case 1      '§§§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "1.§§§Â" 'Add By Sindy 2010/12/8
        RefreshColData 1, "•”Ω–§H"
        If Option1(0).Value = True Then
           'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
           strSql = " SELECT " & m_ColCustName & "," & m_ColCustAdd & ",'' X01,'' X02,'' X03,'' X04,'' X05,CP06,Nvl(DECODE(PA09,'000',CPM03,CPM04),Nvl(CPM10,CPm13)) X06,PA77,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X07,Nvl(PA05,Nvl(PA06,PA07)) X08,PA11,PA22,PA48,CP09,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',CP06,Nvl(DECODE(SP09,'000',CPM03,CPM04),Nvl(CPM10,CPM13)),SP27,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(SP05,Nvl(SP06,SP07)),SP11,'',SP29,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) " & strSQL2 & StrSQL6
        Else
           'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
           strSql = " SELECT " & m_ColCustName & "," & m_ColCustAdd & ",'' X01,'' X02,'' X03,'' X04,'' X05,CP07,Nvl(DECODE(PA09,'000',CPM03,CPM04),Nvl(CPM10,CPM13)) X06,PA77,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X07,Nvl(PA05,Nvl(PA06,PA07)) X08,PA11,PA22,PA48,CP09,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',CP07,Nvl(DECODE(SP09,'000',CPM03,CPM04),Nvl(CPM10,CPM13)),SP27,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(SP05,Nvl(SP06,SP07)),SP11,'',SP29,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) " & strSQL2 & StrSQL6
        End If
   Case 2      '≠^§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "2.≠^§Â" 'Add By Sindy 2010/12/8
        RefreshColData 2, "•”Ω–§H"
        If Option1(0).Value = True Then
           'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
           strSql = " SELECT " & m_ColCustName & "," & m_ColCustAdd & ",CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,CP06,Nvl(cpm10,CPM13) X01,PA77,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X02,Nvl(PA06,PA07) X03,PA11,PA22,PA48,CP09,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColCustName & "," & m_ColCustAdd & ",CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,CP06,Nvl(CPM10,CPM13),SP27,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(SP06,SP07),SP11,'',SP29,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) " & strSQL2 & StrSQL6
        Else
           'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
           strSql = " SELECT " & m_ColCustName & "," & m_ColCustAdd & ",CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,CP07,Nvl(CPM10,CPM13) X01,PA77,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X02,Nvl(PA06,PA07) X03,PA11,PA22,PA48,CP09,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColCustName & "," & m_ColCustAdd & ",CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,CP07,Nvl(CPM10,CPM13),SP27,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(SP06,SP07),SP11,'',SP29,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) " & strSQL2 & StrSQL6
        End If
   Case 3      '§È§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "3.§È§Â" 'Add By Sindy 2010/12/8
        RefreshColData 3, "•”Ω–§H"
        If Option1(0).Value = True Then
           'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
           strSql = " SELECT " & m_ColCustName & "," & m_ColCustAdd & ",'' X01,'' X02,'' X03,'' X04,'' X05,CP06,Nvl(CPM13,CPM10) X06,PA77,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X07,nvl(PA07,pa06) X08,PA11,PA22,PA48,CP09,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) " & strSQL1 & StrSQL6
           strSql = strSql + " union all select CU06,'','','',CU29,'','','','','','','','','',CP06,Nvl(CPM13,CPM10),SP27,CP01||'-'||CP02||'-'||CP03||'-'||CP04,nvl(SP07,sp06),SP11,'',SP29,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) " & strSQL2 & StrSQL6
        Else
           'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT:+CP09,CP27
           strSql = " SELECT " & m_ColCustName & "," & m_ColCustAdd & ",'' X01,'' X02,'' X03,'' X04,'' X05,CP07,Nvl(CPM13,CPM10) X06,PA77,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X07,nvl(PA07,pa06) X08,PA11,PA22,PA48,CP09,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) " & strSQL1 & StrSQL6
           strSql = strSql + " union all select " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',CP07,Nvl(CPM13,CPM10),SP27,CP01||'-'||CP02||'-'||CP03||'-'||CP04,nvl(SP07,sp06),SP11,'',SP29,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) " & strSQL2 & StrSQL6
        End If
   Case Else
   End Select
   
   'Modify by Morgan 2005/1/28 •[πL¬o§w¶¨§£ƒÚøÏ™∫®”®Á
   strSql = "select * from (" & strSql & ") X where " & " not exists(select * from caseprogress a where a.cp43= X.cp09 and a.cp10='907')"
   '2005/1/28 end
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'edit by nickc 2007/02/08
   'k = 0
   
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           Do While .EOF = False
               For i = 0 To 21
                   strTemp(i) = ChgSQL(CheckStr(.Fields(i)))
               Next i
               If Len(strTemp(4)) = 0 And Len(strTemp(5)) = 0 And Len(strTemp(6)) = 0 And Len(strTemp(7)) = 0 And Len(strTemp(8)) = 0 Then
                   For i = 4 To 16
                       strTemp(i) = strTemp(i + 5)
                   Next i
               Else
                   For i = 9 To 16
                       strTemp(i) = strTemp(i + 5)
                   Next i
               End If
               Select Case Val(txt1(14))
               Case 1      '§§§Â
                               strTemp(9) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(9)))
               Case 2, 3
                              strTemp(9) = Format(ChangeWStringToWDateString(strTemp(9)), "mmm DD,YYYY")
               Case Else
               End Select
               strSql = "insert into R060308_4 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & ChgSQL(strTemp(16)) & "','" & strUserNum & "') "
               cnnConnection.Execute strSql
               .MoveNext
               DoEvents
           Loop
       End With
   Else
      Pro9 = False
      ChkPro(9) = False
      Exit Function
   End If
   CheckOC
   Pro9 = True
   ChkPro(9) = True
End Function

Function Pro10() As Boolean
   'Pro10  ´»§·Æ◊•Û(•N≤z§H)(§w¶¨§Â)
   '****************************************************
   '•N≤z§H¶W∫Ÿ, •”Ω–§H¶W∫Ÿ, ¶aß}, Æ◊•Û© ΩË¶W∫Ÿ, Æ◊•Û¶W∫Ÿ
   '≠Y§§§Â≥¯™Ì  §§-->≠^-->§È
   '≠Y≠^§Â≥¯™Ì  ≠^-->§È-->NULL
   '≠Y§È§Â≥¯™Ì  §È-->≠^-->NULL
   '****************************************************
      
   Select Case Val(txt1(14))
   Case 1      '§§§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "1.§§§Â" 'Add By Sindy 2010/12/8
        RefreshColData 1, "•N≤z§H"
        If Option1(0).Value = True Then
   'Modify by Morgan 2005/1/31 ®˙ßO¶W•H´K®œ•ŒµÍ¿¿∏ÍÆ∆™Ì
   '        StrSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',CP06,Nvl(DECODE(PA09,'000',CPM03,CPM04),Nvl(CPM10,CPM13)),PA77,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(PA05,Nvl(PA06,PA07)),PA11,PA22,CU04,PA48,CP09,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,FAGENT,CUSTOMER WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) " & strSQL1 & StrSQL6
           strSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",'' X01,'' X02,'' X03,'' X04,'' X05,CP06,Nvl(DECODE(PA09,'000',CPM03,CPM04),Nvl(CPM10,CPM13)) X07,PA77,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X08,Nvl(PA05,Nvl(PA06,PA07)) X09,PA11,PA22,CU04,PA48,CP09,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,FAGENT,CUSTOMER WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) " & strSQL1 & StrSQL6
   '2005/1/31 end
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',CP06,Nvl(DECODE(SP09,'000',CPM03,CPM04),Nvl(CPM10,CPM13)),SP27,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(SP05,Nvl(SP06,SP07)),SP11,'',CU04,SP29,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,FAGENT,CUSTOMER WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) " & strSQL2 & StrSQL6
        Else
   'Modify by Morgan 2005/1/31 ®˙ßO¶W•H´K®œ•ŒµÍ¿¿∏ÍÆ∆™Ì
   '        StrSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',CP07,Nvl(DECODE(PA09,'000',CPM03,CPM04),Nvl(CPM10,CPM13)),PA77,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(PA05,Nvl(PA06,PA07)),PA11,PA22,CU04,PA48,CP09,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,FAGENT,CUSTOMER WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+)  AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) " & strSQL1 & StrSQL6
           strSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",'' X01,'' X02,'' X03,'' X04,'' X05,CP07,Nvl(DECODE(PA09,'000',CPM03,CPM04),Nvl(CPM10,CPM13)) X07,PA77,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X08,Nvl(PA05,Nvl(PA06,PA07)) X09,PA11,PA22,CU04,PA48,CP09,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,FAGENT,CUSTOMER WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+)  AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) " & strSQL1 & StrSQL6
   '2005/1/31 END
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',CP07,Nvl(DECODE(SP09,'000',CPM03,CPM04),Nvl(CPM10,CPM13)),SP27,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(SP05,Nvl(SP06,SP07)),SP11,'',CU04,SP29,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,FAGENT,CUSTOMER WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) " & strSQL2 & StrSQL6
        End If
   Case 2      '≠^§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "2.≠^§Â" 'Add By Sindy 2010/12/8
        RefreshColData 2, "•N≤z§H"
        If Option1(0).Value = True Then
   'Modify by Morgan 2005/1/31 ®˙ßO¶W•H´K®œ•ŒµÍ¿¿∏ÍÆ∆™Ì
   '        StrSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,FA22,CP06,Nvl(CPM10,CPM13),PA77,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(PA06,PA07),PA11,PA22,CU05||CU88||CU89||CU90,PA48,CP09,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) " & strSQL1 & StrSQL6
           strSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,fa22||rtrim(' '||fa70) FA22,CP06,Nvl(CPM10,CPM13) X01,PA77,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X02,Nvl(PA06,PA07) X03,PA11,PA22,CU05||CU88||CU89||CU90 X04,PA48,CP09,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) " & strSQL1 & StrSQL6
   '2005/1/31 END
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,fa22||rtrim(' '||fa70) FA22,CP06,Nvl(CPM10,CPM13),SP27,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(SP06,SP07),SP11,'',CU05||CU88||CU89||CU90,SP29,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) " & strSQL2 & StrSQL6
        Else
   'Modify by Morgan 2005/1/31 ®˙ßO¶W•H´K®œ•ŒµÍ¿¿∏ÍÆ∆™Ì
   '        StrSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,FA22,CP07,Nvl(CPM10,CPM13),PA77,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(PA06,PA07),PA11,PA22,CU05||CU88||CU89||CU90,PA48,CP09,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+)  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) " & strSQL1 & StrSQL6
           strSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,fa22||rtrim(' '||fa70) FA22,CP07,Nvl(CPM10,CPM13) X01,PA77,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X02,Nvl(PA06,PA07) X03,PA11,PA22,CU05||CU88||CU89||CU90 X04,PA48,CP09,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+)  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) " & strSQL1 & StrSQL6
   '2005/1/31 END
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,fa22||rtrim(' '||fa70) A22,CP07,Nvl(CPM10,CPM13),SP27,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(SP06,SP07),SP11,'',CU05||CU88||CU89||CU90,SP29,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) " & strSQL2 & StrSQL6
        End If
   Case 3      '§È§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "3.§È§Â" 'Add By Sindy 2010/12/8
        RefreshColData 3, "•N≤z§H"
        If Option1(0).Value = True Then
   'Modify by Morgan 2005/1/31 ®˙ßO¶W•H´K®œ•ŒµÍ¿¿∏ÍÆ∆™Ì
   '        StrSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',CP06,Nvl(CPM13,CPM10),PA77,CP01||'-'||CP02||'-'||CP03||'-'||CP04,nvl(PA07,pa06),PA11,PA22,CU06,PA48,CP09,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+)  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) " & strSQL1 & StrSQL6
           strSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",'' X01,'' X02,'' X03,'' X04,'' X05,CP06,Nvl(CPM13,CPM10) X06,PA77,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X07,nvl(PA07,pa06) X08,PA11,PA22,CU06,PA48,CP09,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+)  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) " & strSQL1 & StrSQL6
   '2005/1/31 END
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',CP06,Nvl(CPm13,CPM10),SP27,CP01||'-'||CP02||'-'||CP03||'-'||CP04,nvl(SP07,sp06),SP11,'',CU06,SP29,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) " & strSQL2 & StrSQL6
        Else
   'Modify by Morgan 2005/1/31 ®˙ßO¶W•H´K®œ•ŒµÍ¿¿∏ÍÆ∆™Ì
   '        StrSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',CP07,Nvl(CPm13,CPM10),PA77,CP01||'-'||CP02||'-'||CP03||'-'||CP04,nvl(PA07,pa06),PA11,PA22,CU06,PA48,CP09,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) " & strSQL1 & StrSQL6
           strSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",'' X01,'' X02,'' X03,'' X04,'' X05,CP07,Nvl(CPm13,CPM10) X06,PA77,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X07,nvl(PA07,pa06) X08,PA11,PA22,CU06,PA48,CP09,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) " & strSQL1 & StrSQL6
   '2005/1/31 END
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',CP07,Nvl(CPm13,CPM10),SP27,CP01||'-'||CP02||'-'||CP03||'-'||CP04,nvl(SP07,sp06),SP11,'',CU06,SP29,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) " & strSQL2 & StrSQL6
        End If
   Case Else
   End Select
   
   'Modify by Morgan 2005/1/28 •[πL¬o§w¶¨§£ƒÚøÏ™∫®”®Á
   strSql = "select * from (" & strSql & ") X where " & " not exists(select * from caseprogress a where a.cp43= X.cp09 and a.cp10='907')"
   '2005/1/28 end
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           DoEvents
           Do While .EOF = False
               For i = 0 To 22
                   strTemp(i) = ChgSQL(CheckStr(.Fields(i)))
               Next i
               If Len(strTemp(4)) = 0 And Len(strTemp(5)) = 0 And Len(strTemp(6)) = 0 And Len(strTemp(7)) = 0 And Len(strTemp(8)) = 0 Then
                   For i = 4 To 17
                       strTemp(i) = strTemp(i + 5)
                   Next i
               Else
                   For i = 9 To 17
                       strTemp(i) = strTemp(i + 5)
                   Next i
               End If
               Select Case Val(txt1(14))
               Case 1      '§§§Â
                               strTemp(9) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(9)))
               Case 2, 3
                               strTemp(9) = Format(ChangeWStringToWDateString(strTemp(9)), "mmm DD,YYYY")
               Case Else
               End Select
               strSql = "insert into R060308_5 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "','" & strUserNum & "') "
               cnnConnection.Execute strSql
               .MoveNext
               DoEvents
           Loop
       End With
   Else
      Pro10 = False
      ChkPro(10) = False
      Exit Function
   End If
   CheckOC
   Pro10 = True
   ChkPro(10) = True
End Function

Function Pro11() As Boolean
   'Pro11  ¶~∂O(•”Ω–§H)(§w¶¨§Â)
   '****************************************************
   '•N≤z§H¶W∫Ÿ, •”Ω–§H¶W∫Ÿ, ¶aß}, Æ◊•Û© ΩË¶W∫Ÿ, Æ◊•Û¶W∫Ÿ
   '≠Y§§§Â≥¯™Ì  §§-->≠^-->§È
   '≠Y≠^§Â≥¯™Ì  ≠^-->§È-->NULL
   '≠Y§È§Â≥¯™Ì  §È-->≠^-->NULL
   '****************************************************
      
   Select Case Val(txt1(14))
   Case 1      '§§§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "1.§§§Â" 'Add By Sindy 2010/12/8
        RefreshColData 1, "•”Ω–§H"
        If Option1(0).Value = True Then
           'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
           strSql = "SELECT ALL " & m_ColCustName & "," & m_ColCustAdd & ",'' X01,'' X02,'' X03,'' X04,'' X05,CP06,PA77,PA48,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X06,Nvl(PA05,Nvl(PA06,PA07)) X07,PA11,PA22,CP09,CP27 FROM CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND CP10='605' " & strSQL1 & StrSQL6
           strSql = strSql + " union ALL SELECT " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',CP06,SP27,SP29,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(SP05,Nvl(SP06,SP07)),SP11,'',CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND CP10='605' " & strSQL2 & StrSQL6
        Else
           'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
           strSql = "SELECT ALL " & m_ColCustName & "," & m_ColCustAdd & ",'' X01,'' X02,'' X03,'' X04,'' X05,CP07,PA77,PA48,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X06,Nvl(PA05,Nvl(PA06,PA07)) X07,PA11,PA22,CP09,CP27 FROM CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND CP10='605' " & strSQL1 & StrSQL6
           strSql = strSql + " union ALL SELECT " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',CP07,SP27,SP29,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(SP05,Nvl(SP06,SP07)),SP11,'',CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND CP10='605' " & strSQL2 & StrSQL6
        End If
   Case 2      '≠^§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "2.≠^§Â" 'Add By Sindy 2010/12/8
        RefreshColData 2, "•”Ω–§H"
        If Option1(0).Value = True Then
           'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
           strSql = "SELECT ALL " & m_ColCustName & "," & m_ColCustAdd & ",CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,CP06,PA77,PA48,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X01,Nvl(PA06,PA07) X02,PA11,PA22,CP09,CP27 FROM CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND CP10='605' " & strSQL1 & StrSQL6
           strSql = strSql + " union ALL SELECT " & m_ColCustName & "," & m_ColCustAdd & ",CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,CP06,SP27,SP29,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(SP06,SP07),SP11,'',CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND CP10='605' " & strSQL2 & StrSQL6
        Else
           'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
           strSql = "SELECT ALL " & m_ColCustName & "," & m_ColCustAdd & ",CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,CP07,PA77,PA48,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X01,Nvl(PA06,PA07) X02,PA11,PA22,CP09,CP27 FROM CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND CP10='605' " & strSQL1 & StrSQL6
           strSql = strSql + " union ALL SELECT " & m_ColCustName & "," & m_ColCustAdd & ",CU24,CU25,CU26,CU27,CU28||rtrim(' '||cu102) CU28,CP07,SP27,SP29,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(SP06,SP07),SP11,'',CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND CP10='605' " & strSQL2 & StrSQL6
        End If
   Case 3      '§È§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "3.§È§Â" 'Add By Sindy 2010/12/8
        RefreshColData 3, "•”Ω–§H"
        If Option1(0).Value = True Then
           'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
           strSql = "SELECT ALL " & m_ColCustName & "," & m_ColCustAdd & ",'' X01,'' X02,'' X03,'' X04,'' X05,CP06,PA77,PA48,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X06,nvl(PA07,pa06) X07,PA11,PA22,CP09,CP27 FROM CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND CP10='605' " & strSQL1 & StrSQL6
           strSql = strSql + " union ALL SELECT " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',CP06,SP27,SP29,CP01||'-'||CP02||'-'||CP03||'-'||CP04,nvl(SP07,sp06),SP11,'',CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND CP10='605' " & strSQL2 & StrSQL6
        Else
           'Modified by Lydia 2025/01/06 +•N≤z§HFAGENT
           strSql = "SELECT ALL " & m_ColCustName & "," & m_ColCustAdd & ",'' X01,'' X02,'' X03,'' X04,'' X05,CP07,PA77,PA48,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X06,nvl(PA07,pa06) X07,PA11,PA22,CP09,CP27 FROM CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND CP10='605' " & strSQL1 & StrSQL6
           strSql = strSql + " union ALL SELECT " & m_ColCustName & "," & m_ColCustAdd & ",'','','','','',CP07,SP27,SP29,CP01||'-'||CP02||'-'||CP03||'-'||CP04,nvl(SP07,sp06),SP11,'',CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND CP10='605' " & strSQL2 & StrSQL6
        End If
   Case Else
   End Select
   
   'Modify by Morgan 2005/1/28 •[πL¬o§w¶¨§£ƒÚøÏ™∫®”®Á
   strSql = "select * from (" & strSql & ") X where " & " not exists(select * from caseprogress a where a.cp43= X.cp09 and a.cp10='907')"
   '2005/1/28 end
   
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'edit by nickc 2007/02/08
   'k = 0
   
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           Do While .EOF = False
               For i = 0 To 20
                   strTemp(i) = ChgSQL(CheckStr(.Fields(i)))
               Next i
               If Len(strTemp(4)) = 0 And Len(strTemp(5)) = 0 And Len(strTemp(6)) = 0 And Len(strTemp(7)) = 0 And Len(strTemp(8)) = 0 Then
                   For i = 4 To 15
                       strTemp(i) = strTemp(i + 5)
                   Next i
               Else
                   For i = 9 To 15
                       strTemp(i) = strTemp(i + 5)
                   Next i
               End If
               Select Case Val(txt1(14))
               Case 1      '§§§Â
                               strTemp(9) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(9)))
               Case 2, 3
                               strTemp(9) = Format(ChangeWStringToWDateString(strTemp(9)), "mmm DD,YYYY")
               Case Else
               End Select
               strSql = "insert into R060308_6 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & strUserNum & "') "
               cnnConnection.Execute strSql
               .MoveNext
               DoEvents
           Loop
       End With
   Else
      Pro11 = False
      ChkPro(11) = False
      Exit Function
   End If
   CheckOC
   Pro11 = True
   ChkPro(11) = True
End Function

Function Pro12() As Boolean
   'Pro12  ¶~∂O(•N≤z§H)(§w¶¨§Â)
   '****************************************************
   '•N≤z§H¶W∫Ÿ, •”Ω–§H¶W∫Ÿ, ¶aß}, Æ◊•Û© ΩË¶W∫Ÿ, Æ◊•Û¶W∫Ÿ
   '≠Y§§§Â≥¯™Ì  §§-->≠^-->§È
   '≠Y≠^§Â≥¯™Ì  ≠^-->§È-->NULL
   '≠Y§È§Â≥¯™Ì  §È-->≠^-->NULL
   '****************************************************
   
   Select Case Val(txt1(14))
   Case 1      '§§§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "1.§§§Â" 'Add By Sindy 2010/12/8
        RefreshColData 1, "•N≤z§H"
        If Option1(0).Value = True Then
   'Modify by Morgan 2005/1/28 ®˙ßO¶W•H´K®œ•ŒµÍ¿¿∏ÍÆ∆™Ì
   '        StrSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',CP06,PA77,PA48,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(PA05,nvl(PA06,PA07)),PA11,PA22,CU04,CP09,CP27 FROM CASEPROGRESS,PATENT,FAGENT,CUSTOMER WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND CP10='605' " & strSQL1 & StrSQL6
           strSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",'' X01,'' X02,'' X03,'' X04,'' X05,CP06,PA77,PA48,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X06,Nvl(PA05,nvl(PA06,PA07)) X07,PA11,PA22,CU04,CP09,CP27 FROM CASEPROGRESS,PATENT,FAGENT,CUSTOMER WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND CP10='605' " & strSQL1 & StrSQL6
   '2005/1/31 END
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',CP06,SP27,SP29,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(SP05,Nvl(SP06,SP07)),SP11,'',CU04,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,FAGENT,CUSTOMER WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND CP10='605' " & strSQL2 & StrSQL6
        Else
   'Modify by Morgan 2005/1/28 ®˙ßO¶W•H´K®œ•ŒµÍ¿¿∏ÍÆ∆™Ì
   '        StrSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',CP07,PA77,PA48,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(PA05,Nvl(PA06,PA07)),PA11,PA22,CU04,CP09,CP27 FROM CASEPROGRESS,PATENT,FAGENT,CUSTOMER WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND CP10='605' " & strSQL1 & StrSQL6
           strSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",'' X01,'' X02,'' X03,'' X04,'' X05,CP07,PA77,PA48,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X06,Nvl(PA05,Nvl(PA06,PA07)) X07,PA11,PA22,CU04,CP09,CP27 FROM CASEPROGRESS,PATENT,FAGENT,CUSTOMER WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND CP10='605' " & strSQL1 & StrSQL6
   '2005/1/31 END
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',CP07,SP27,SP29,CP01||'-'||CP02||'-'||CP03||'-'||CP04,nvl(SP05,Nvl(SP06,SP07)),SP11,'',CU04,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,FAGENT,CUSTOMER WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND CP10='605' " & strSQL2 & StrSQL6
        End If
   Case 2      '≠^§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "2.≠^§Â" 'Add By Sindy 2010/12/8
        RefreshColData 2, "•N≤z§H"
        If Option1(0).Value = True Then
   'Modify by Morgan 2005/1/28 ®˙ßO¶W•H´K®œ•ŒµÍ¿¿∏ÍÆ∆™Ì
   '        StrSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,FA22,CP06,PA77,PA48,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(PA06,PA07),PA11,PA22,CU05||CU88||CU89||CU90,CP09,CP27 FROM CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND CP10='605' " & strSQL1 & StrSQL6
           strSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,fa22||rtrim(' '||fa70) FA22,CP06,PA77,PA48,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X01,Nvl(PA06,PA07) X02,PA11,PA22,CU05||CU88||CU89||CU90 X03,CP09,CP27 FROM CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND CP10='605' " & strSQL1 & StrSQL6
   '2005/1/31 END
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,fa22||rtrim(' '||fa70) FA22,CP06,SP27,SP29,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(SP06,SP07),SP11,'',CU05||CU88||CU89||CU90,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+)  AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND CP10='605' " & strSQL2 & StrSQL6
        Else
   'Modify by Morgan 2005/1/28 ®˙ßO¶W•H´K®œ•ŒµÍ¿¿∏ÍÆ∆™Ì
   '        StrSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,FA22,CP07,PA77,PA48,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(PA06,PA07),PA11,PA22,CU05||CU88||CU89||CU90,CP09,CP27 FROM CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND CP10='605' " & strSQL1 & StrSQL6
           strSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,fa22||rtrim(' '||fa70) FA22,CP07,PA77,PA48,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X01,Nvl(PA06,PA07) X02,PA11,PA22,CU05||CU88||CU89||CU90 X03,CP09,CP27 FROM CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND CP10='605' " & strSQL1 & StrSQL6
   '2005/1/31 END
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",FA18,FA19,FA20,FA21,fa22||rtrim(' '||fa70) FA22,CP07,SP27,SP29,CP01||'-'||CP02||'-'||CP03||'-'||CP04,Nvl(SP06,SP07),SP11,'',CU05||CU88||CU89||CU90,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+)  AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND CP10='605' " & strSQL2 & StrSQL6
        End If
   Case 3      '§È§Â
        pub_QL05 = pub_QL05 & ";" & Label1(8) & "3.§È§Â" 'Add By Sindy 2010/12/8
        RefreshColData 3, "•N≤z§H"
        If Option1(0).Value = True Then
   'Modify by Morgan 2005/1/28 ®˙ßO¶W•H´K®œ•ŒµÍ¿¿∏ÍÆ∆™Ì
   '        StrSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',CP06,PA77,PA48,CP01||'-'||CP02||'-'||CP03||'-'||CP04,nvl(PA07,pa06),PA11,PA22,CU06,CP09,CP27 FROM CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND CP10='605' " & strSQL1 & StrSQL6
           strSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",'' X01,'' X02,'' X03,'' X04,'' X05,CP06,PA77,PA48,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X06,nvl(PA07,pa06) X07,PA11,PA22,CU06,CP09,CP27 FROM CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND CP10='605' " & strSQL1 & StrSQL6
   '2005/1/31 END
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',CP06,SP27,SP29,CP01||'-'||CP02||'-'||CP03||'-'||CP04,nvl(SP07,sp06),SP11,'',CU06,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND CP10='605' " & strSQL2 & StrSQL6
        Else
   'Modify by Morgan 2005/1/28 ®˙ßO¶W•H´K®œ•ŒµÍ¿¿∏ÍÆ∆™Ì
   '        StrSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',CP06,PA77,PA48,CP01||'-'||CP02||'-'||CP03||'-'||CP04,nvl(PA07,pa06),PA11,PA22,CU06,CP09,CP27 FROM CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND CP10='605' " & strSQL1 & StrSQL6
           strSql = "SELECT ALL " & m_ColAgName & "," & m_ColAgAdd & ",'' X01,'' X02,'' X03,'' X04,'' X05,CP06,PA77,PA48,CP01||'-'||CP02||'-'||CP03||'-'||CP04 X06,nvl(PA07,pa06) X07,PA11,PA22,CU06,CP09,CP27 FROM CASEPROGRESS,PATENT,CUSTOMER,FAGENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),'','0',SUBSTR(PA75,9,1))=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND CP10='605' " & strSQL1 & StrSQL6
   '2005/1/31 END
           strSql = strSql + " union all select " & m_ColAgName & "," & m_ColAgAdd & ",'','','','','',CP06,SP27,SP29,CP01||'-'||CP02||'-'||CP03||'-'||CP04,nvl(SP07,sp06),SP11,'',CU06,CP09,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,FAGENT WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),'','0',SUBSTR(SP26,9,1))=FA02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND CP10='605' " & strSQL2 & StrSQL6
        End If
   Case Else
   End Select
   
   'Modify by Morgan 2005/1/28 •[πL¬o§w¶¨§£ƒÚøÏ™∫®”®Á
   strSql = "select * from (" & strSql & ") X where " & " not exists(select * from caseprogress a where a.cp43= X.cp09 and a.cp10='907')"
   '2005/1/28 end
   
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'edit by nickc 2007/02/08
   'k = 0
   
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           DoEvents
           Do While .EOF = False
               For i = 0 To 21
                   strTemp(i) = ChgSQL(CheckStr(.Fields(i)))
               Next i
               If Len(strTemp(4)) = 0 And Len(strTemp(5)) = 0 And Len(strTemp(6)) = 0 And Len(strTemp(7)) = 0 And Len(strTemp(8)) = 0 Then
                   For i = 4 To 16
                       strTemp(i) = strTemp(i + 5)
                   Next i
               Else
                   For i = 9 To 16
                       strTemp(i) = strTemp(i + 5)
                   Next i
               End If
               Select Case Val(txt1(14))
               Case 1      '§§§Â
                               strTemp(9) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(9)))
               Case 2, 3
                               strTemp(9) = Format(ChangeWStringToWDateString(strTemp(9)), "mmm DD,YYYY")
               Case Else
               End Select
               strSql = "insert into R060308_7 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & ChgSQL(strTemp(16)) & "','" & strUserNum & "') "
               cnnConnection.Execute strSql
               .MoveNext
               DoEvents
           Loop
       End With
   Else
      Pro12 = False
      ChkPro(12) = False
      Exit Function
   End If
   CheckOC
   Pro12 = True
   ChkPro(12) = True
End Function

'§w¶¨§Â
Sub Process2()
Dim PrinterIndex As Integer 'Add By Sindy 2015/12/31
   
   strSQL1 = ""
   strSQL2 = ""
   StrSQL6 = ""
   
   If Len(txt1(0)) <> 0 Then
      'Modified by Morgan 2012/5/25 +∞Í•~≥°©”øÏ™∫FMPÆ◊•ºßπ∑d©Œ¬Ωƒ∂Æ÷ΩZßπ¶®µ{ß«
      'strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
      'strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
      strSQL1 = strSQL1 + " AND (CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") or (cp01 in ('P','CFP') and SUBSTR(cp12,1,1)='F' and exists(select * from staff x where x.st01=CP14 and st03 like 'F%')" & _
         " and exists(select * from engineerprogress x where x.ep02=cp09 and (ep09||" & IIf(strSrvDate(1) >= FCPÆ÷ßπ§ÈßÔ•ŒEP39, "ep39", "ep33") & " is null or cp10||" & IIf(strSrvDate(1) >= FCPÆ÷ßπ§ÈßÔ•ŒEP39, "ep39", "ep33") & "='201')))) "
      strSQL2 = strSQL2 + " AND (CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") or (cp01 in ('PS','CPS') and SUBSTR(cp12,1,1)='F' and exists(select * from staff x where x.st01=CP14 and st03 like 'F%')" & _
         " and exists(select * from engineerprogress x where x.ep02=cp09 and (ep09||" & IIf(strSrvDate(1) >= FCPÆ÷ßπ§ÈßÔ•ŒEP39, "ep39", "ep33") & " is null or cp10||" & IIf(strSrvDate(1) >= FCPÆ÷ßπ§ÈßÔ•ŒEP39, "ep39", "ep33") & "='201')))) "
      'end 2012/5/25
      pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/8
   End If
   
   strSQL1 = StrSQL6 + strSQL1 & " AND (PA57<>'Y' OR PA57 IS NULL) "
   strSQL2 = StrSQL6 + strSQL2 & " AND (SP15<>'Y' OR SP15 IS NULL) "
   StrSQL6 = " AND CP27 IS NULL AND CP57 IS NULL "
   If Len(txt1(6)) <> 0 Then
      StrSQL6 = StrSQL6 & " AND CP10 IN (" & GetAddStr(txt1(6)) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(6) 'Add By Sindy 2010/12/8
   End If
   strSQL1 = strSQL1 + StrSQL6
   strSQL2 = strSQL2 + StrSQL6
   StrSQL6 = ""
   strSQL8 = "" 'Add By Sindy 2021/4/15
   
   '•ª©“¥¡≠≠
   If Option1(0).Value = True Then
      If Len(txt1(1)) <> 0 Then
         'Modify By Sindy 2021/5/5 •[ßP¬_´¸©w∞e•Û§È
         'StrSQL6 = StrSQL6 + " AND CP06>=" & Val(ChangeTStringToWString(txt1(1))) & " "
         StrSQL6 = StrSQL6 + " AND ((CP06>=" & Val(ChangeTStringToWString(txt1(1))) & " and CP06<=" & Val(ChangeTStringToWString(txt1(2))) & ") or (CP142>=" & Val(ChangeTStringToWString(txt1(1))) & " and CP142<=" & Val(ChangeTStringToWString(txt1(2))) & "))"
         '2021/5/5 END
         strSQL8 = strSQL8 + " AND CP48>=" & Val(ChangeTStringToWString(txt1(1))) & " AND CP48<=" & Val(ChangeTStringToWString(txt1(2))) 'Add By Sindy 2021/4/15
      Else
         'Modify By Sindy 2021/5/5 •[ßP¬_´¸©w∞e•Û§È
         'StrSQL6 = StrSQL6 + " AND CP06<=" & Val(ChangeTStringToWString(txt1(2))) '& " and cp06>=" & Val(GetTodayDate)
         StrSQL6 = StrSQL6 + " AND (CP06<=" & Val(ChangeTStringToWString(txt1(2))) & " or CP142<=" & Val(ChangeTStringToWString(txt1(2))) & ")" '& " and cp06>=" & Val(GetTodayDate)
         '2021/5/5 END
         strSQL8 = strSQL8 + " AND CP48<=" & Val(ChangeTStringToWString(txt1(2))) 'Add By Sindy 2021/4/15
      End If
      pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/8
   '™k©w¥¡≠≠
   Else
       If Option1(1).Value = True Then
           If Len(txt1(3)) <> 0 Then
           StrSQL6 = StrSQL6 + " AND CP07>=" & Val(ChangeTStringToWString(txt1(3))) & " "
           strSQL8 = strSQL8 + " AND CP48>=" & Val(ChangeTStringToWString(txt1(3))) & " " 'Add By Sindy 2021/4/15
           End If
           StrSQL6 = StrSQL6 + " AND CP07<=" & Val(ChangeTStringToWString(txt1(4))) '& " and cp06>=" & Val(GetTodayDate)
           strSQL8 = strSQL8 + " AND CP48<=" & Val(ChangeTStringToWString(txt1(4))) 'Add By Sindy 2021/4/15
           pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/12/8
       End If
   End If
   
   If Len(txt1(7)) <> 0 Then
       strSQL1 = strSQL1 + " AND CP13='" & txt1(7) & "' "
       strSQL2 = strSQL2 + " AND CP13='" & txt1(7) & "' "
       pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(7) & lbl1(0) 'Add By Sindy 2010/12/8
   End If
   '©”øÏ§H
   If Len(txt1(8)) <> 0 Then
      'Modify By Sindy 2015/12/28 øÈ§J™L´H©˜Æ…,•Á§]≠nøÈ•X•t2≠”Ωs∏π∏ÍÆ∆
      If txt1(8) = "68007" Then
         strSQL1 = strSQL1 + " AND CP14 in('68007','68091','68092') "
         strSQL2 = strSQL2 + " AND CP14 in('68007','68091','68092') "
      Else
      '2015/12/28 END
         strSQL1 = strSQL1 + " AND CP14='" & txt1(8) & "' "
         strSQL2 = strSQL2 + " AND CP14='" & txt1(8) & "' "
      End If
      pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(8) & lbl1(1) 'Add By Sindy 2010/12/8
   End If
   If Len(Trim(txt1(10))) <> 0 And Len(Trim(txt1(11))) <> 0 Then
       strSQL1 = strSQL1 + " AND ((PA26>='" & GetNewFagent(txt1(10)) & "' AND PA26<='" & GetNewFagent(txt1(11)) & "') OR (PA27>='" & GetNewFagent(txt1(10)) & "' AND PA27<='" & GetNewFagent(txt1(11)) & "') OR (PA28>='" & GetNewFagent(txt1(10)) & "' AND PA28<='" & GetNewFagent(txt1(11)) & "') OR (PA29>='" & GetNewFagent(txt1(10)) & "' AND PA29<='" & GetNewFagent(txt1(11)) & "') OR (PA30>='" & GetNewFagent(txt1(10)) & "' AND PA30<='" & GetNewFagent(txt1(11)) & "')) "
       strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(txt1(10)) & "' AND SP08<='" & GetNewFagent(txt1(11)) & "') OR (SP58<='" & GetNewFagent(txt1(10)) & "' AND SP58<='" & GetNewFagent(txt1(11)) & "') OR (SP59>='" & GetNewFagent(txt1(10)) & "' AND SP59<='" & GetNewFagent(txt1(11)) & "') OR (SP65>='" & GetNewFagent(txt1(10)) & "' AND SP65<='" & GetNewFagent(txt1(11)) & "') OR (SP66>='" & GetNewFagent(txt1(10)) & "' AND SP66<='" & GetNewFagent(txt1(11)) & "')) "
   Else
       If Len(Trim(txt1(10))) <> 0 And Len(Trim(txt1(11))) = 0 Then
           strSQL1 = strSQL1 + " AND (PA26>='" & GetNewFagent(txt1(10)) & "' OR PA27>='" & GetNewFagent(txt1(10)) & "' OR PA28>='" & GetNewFagent(txt1(10)) & "' OR PA29>='" & GetNewFagent(txt1(10)) & "' OR PA30>='" & GetNewFagent(txt1(10)) & "') "
           strSQL2 = strSQL2 + " AND (SP08>='" & GetNewFagent(txt1(10)) & "' OR SP58>='" & GetNewFagent(txt1(10)) & "' OR SP59>='" & GetNewFagent(txt1(10)) & "' OR SP65>='" & GetNewFagent(txt1(10)) & "' OR SP66>='" & GetNewFagent(txt1(10)) & "') "
       Else
           If Len(Trim(txt1(10))) = 0 And Len(Trim(txt1(11))) <> 0 Then
               strSQL1 = strSQL1 + " AND (PA26<='" & GetNewFagent(txt1(11)) & "' OR PA27<='" & GetNewFagent(txt1(11)) & "' OR PA28<='" & GetNewFagent(txt1(11)) & "' OR PA29<='" & GetNewFagent(txt1(11)) & "' OR PA30<='" & GetNewFagent(txt1(11)) & "') "
               strSQL2 = strSQL2 + " AND (SP08<='" & GetNewFagent(txt1(11)) & "' OR SP58<='" & GetNewFagent(txt1(11)) & "' OR SP59<='" & GetNewFagent(txt1(11)) & "' OR SP65<='" & GetNewFagent(txt1(11)) & "' OR SP66<='" & GetNewFagent(txt1(11)) & "') "
           End If
       End If
   End If
   If Len(Trim(txt1(10))) <> 0 Or Len(Trim(txt1(11))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(6) & txt1(10) & "-" & txt1(11) 'Add By Sindy 2010/12/8
   End If
   '•N≤z§H
   If Len(Trim(txt1(12))) <> 0 And Len(Trim(txt1(13))) <> 0 Then
       strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(12)) & "' AND PA75<='" & GetNewFagent(txt1(13)) & "' "
       strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(12)) & "' AND SP26<='" & GetNewFagent(txt1(13)) & "' "
   Else
       If Len(Trim(txt1(12))) <> 0 And Len(Trim(txt1(13))) = 0 Then
           strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(12)) & "' "
           strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(12)) & "' "
       Else
           If Len(Trim(txt1(12))) = 0 And Len(Trim(txt1(13))) <> 0 Then
               strSQL1 = strSQL1 + " AND PA75<='" & GetNewFagent(txt1(13)) & "' "
               strSQL2 = strSQL2 + " AND SP26<='" & GetNewFagent(txt1(13)) & "' "
           End If
       End If
   End If
   If Len(Trim(txt1(12))) <> 0 Or Len(Trim(txt1(13))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & txt1(12) & "-" & txt1(13) 'Add By Sindy 2010/12/8
   End If
   'Pro8  ∫ﬁ®Ó§H(§w¶¨§Â)
   'Pro2   ¥º≈v§H≠˚(•º¶¨§Â)
   'Pro3   ©”øÏ§H(§w¶¨§Â)
   'Pro9   ´»§·Æ◊•Û(•”Ω–§H)(§w¶¨§Â)
   'Pro10  ´»§·Æ◊•Û(•N≤z§H)(§w¶¨§Â)
   'Pro11  ¶~∂O(•”Ω–§H)(§w¶¨§Â)
   'Pro12  ¶~∂O(•N≤z§H)(§w¶¨§Â)
   
   '≠Y•ºøÈ§J•”Ω–§H§Œ•N≤z§H
   If Len(txt1(10)) = 0 And Len(txt1(11)) = 0 And Len(txt1(12)) = 0 And Len(txt1(13)) = 0 Then
       Pro8
       Pro3
       strSql = "SELECT DISTINCT R036001 FROM R060308_1 WHERE ID='" & strUserNum & "' "
       CheckOC
       adoRecordset.CursorLocation = adUseClient
       adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
           With adoRecordset
               InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/8
               .MoveFirst
               Do While .EOF = False
                   cnnConnection.Execute "DELETE FROM R060308_3 WHERE ID='" & strUserNum & "' AND R038001='" & ChgSQL(CheckStr(.Fields(0))) & "' "
                   .MoveNext
               Loop
           End With
       Else
         InsertQueryLog (0) 'Add By Sindy 2010/12/8
       End If
       CheckOC
       If Not ChkPro(8) And Not ChkPro(3) Then
         ShowNoData
         Exit Sub
       End If
       
       Printer.Orientation = 2
       DoEvents
       'If Printer.Orientation <> 2 Then Printer.Orientation = 2
       If ChkPro(8) Then PrintPro8_A4
       
      'Add By Sindy 2015/12/31
      If txt1(8) = "" Then
         '¿À¨d¨Oß_¶≥¶w∏ÀPDFCreator
         PrinterIndex = -1
         For i = 0 To Printers.Count - 1
            If UCase(Printers(i).DeviceName) = UCase$("PDFCreator") Then
               PrinterIndex = i
               Exit For
            End If
         Next i
         If PrinterIndex < 0 Then
            MsgBox "Ω–≥q™æπq∏£§§§ﬂ¶w∏ÀPDFCreator !!!"
            Exit Sub
         End If
         PUB_RestorePrinter Printers(PrinterIndex).DeviceName
      End If
      '2015/12/31 END
      
      Printer.Orientation = 2
      DoEvents
      'If Printer.Orientation <> 2 Then Printer.Orientation = 2
      If ChkPro(3) Then PrintPro3_A4
      
      If ChkPrintPro(8) Or ChkPrintPro(3) Then
         ShowPrintOk
      End If
      
   '≠Y¶≥øÈ§J•”Ω–§H©Œ•N≤z§H
   Else
       '≠Y¶≥øÈ§J•”Ω–§H
       If Len(txt1(10)) <> 0 Then
           'Pro9
           'Pro11
            If Not Pro9 And Not Pro11 Then
               InsertQueryLog (0) 'Add By Sindy 2010/12/8
               ShowNoData
               Exit Sub
            End If
            InsertQueryLog ("") 'Add By Sindy 2010/12/8
            If ChkPro(9) Then PrintPro9
            If ChkPro(11) Then PrintPro11
            If ChkPrintPro(9) Or ChkPrintPro(11) Then
               ShowPrintOk
            End If
       '≠Y¶≥øÈ§J•N≤z§H
       Else
           'Pro10
           'Pro12
            If Not Pro10 And Not Pro12 Then
               InsertQueryLog (0) 'Add By Sindy 2010/12/8
               ShowNoData
               Exit Sub
            End If
            InsertQueryLog ("") 'Add By Sindy 2010/12/8
            If ChkPro(10) Then PrintPro10
            If ChkPro(12) Then PrintPro12
            If ChkPrintPro(10) Or ChkPrintPro(12) Then
               ShowPrintOk
            End If
       End If
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txt1(0) = GetSystemKindByNick
   j = 0
   
'Modify by Morgan 2011/3/15 ßÔ¶@•Œ•B§£≠n±∆∞£πw≥]¶L™Ìæ˜
'   strSql = Printer.DeviceName
'   SeekPrintL = Printer.Orientation
'   For i = 0 To Printers.Count - 1
'       Set Printer = Printers(i)
'       Combo1.AddItem Printer.DeviceName, i
'       If Printer.DeviceName = strSql Then
'           SeekPrint = i
'       End If
'   Next i
'   Combo1.Text = Combo1.List(SeekPrint)
'
'
'   If SeekPrintL = 1 Then
'       Option2(0).Value = True
'   Else
'       Option2(1).Value = True
'   End If
   m_intDefaultOri = Printer.Orientation
   PUB_SetPrinter Me.Name, Combo1, strPrinter
   Option2(1).Value = True
'end 2011/3/15
   'Add by Amy  2020/03/31 §Ω•q¶W∫Ÿ
   strCmp_C = CompNameQuery(2, 1)
   strCmp_J = CompNameQuery(2, 3)
   'end 2020/03/31
   Option1_Click 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '≠Y¶L™Ìæ˜≈‹∞ , ´hßÛ∑s¶C¶L≥]©w
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   Set frm060308 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   If Option1(0).Value = True Then
      txt1(1).Enabled = True
      txt1(2).Enabled = True
      txt1(3).Enabled = False
      txt1(4).Enabled = False
   Else
      txt1(1).Enabled = False
      txt1(2).Enabled = False
      txt1(3).Enabled = True
      txt1(4).Enabled = True
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
       cmdOK(0).SetFocus
   End If
End Sub
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
   Case 5
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   Case 14
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
'add by nickc 2007/02/08
Dim StrTempP, StrTempP2

   Select Case Index
      Case 0
           If Len(txt1(0)) <> 0 Then
              STRSTRING = ""
              StrTempP = Split(Replace(txt1(0), ",,", ""), ",")
              StrTempP2 = Split(Replace(GetSystemKindByNick, ",,", ""), ",")
              For i = 0 To UBound(StrTempP)
                  s = 0
                  For j = 0 To UBound(StrTempP2)
                      If StrTempP(i) = StrTempP2(j) Then
                          s = 1
                      End If
                  Next j
                  If s = 0 Then
                      STRSTRING = STRSTRING + StrTempP(i) + " "
                  End If
              Next i
              If Len(STRSTRING) <> 0 Then
                  s = MsgBox(STRSTRING + " §£¨O " + strUserNum + " ™∫≈v≠≠!!", , "ƒµßi!!!")
                  txt1(0).SetFocus
                  Exit Sub
              End If
              If Len(txt1(7)) <> 0 Then
                  lbl1(1).Caption = GetPrjState4(StrTempP(0) + "---", txt1(7))
              End If
            End If
      Case 2
         If blnClkSure = False Then
           If RunNick(txt1(1), txt1(2)) Then
               txt1(1).SetFocus
               txt1_GotFocus (1)
               Exit Sub
           End If
         Else
            blnClkSure = False
         End If
      Case 4
         If blnClkSure = False Then
           If RunNick(txt1(3), txt1(4)) Then
             txt1(3).SetFocus
             txt1_GotFocus (3)
             Exit Sub
            End If
         Else
            blnClkSure = False
         End If
      Case 5
         If Me.txt1(5).Text <> "" Then
           Select Case Val(txt1(5))
           Case 1, 2
           Case Else
                s = MsgBox("¶C¶LßO•uØ‡ 1 ©Œ 2 !!", , "USER øÈ§Jø˘ª~")
                txt1(5).SetFocus
                txt1(5).SelStart = 0
                txt1(5).SelLength = Len(txt1(5))
                Exit Sub
           End Select
         End If
      Case 7
           lbl1(0).Caption = GetPrjSales(txt1(7), "¥º≈v§H≠˚")
            If Me.txt1(7).Text <> "" Then
               If Me.txt1(7).Text = Me.lbl1(0).Caption Then
                  Me.lbl1(0).Caption = ""
                  Me.txt1(7).SetFocus
                  txt1_GotFocus 7
                  Exit Sub
               End If
            End If
      Case 8
           lbl1(1).Caption = GetPrjSales(txt1(8))
            If Me.txt1(8).Text <> "" Then
               If Me.txt1(8).Text = Me.lbl1(1).Caption Then
                  Me.lbl1(1).Caption = ""
                  Me.txt1(8).SetFocus
                  txt1_GotFocus 8
                  Exit Sub
               End If
            End If
      Case 9
           lbl1(2).Caption = GetPrjSales(txt1(9), "∫ﬁ®Ó§H")
            If Me.txt1(9).Text <> "" Then
               If Me.txt1(9).Text = Me.lbl1(2).Caption Then
                  Me.lbl1(2).Caption = ""
                  Me.txt1(9).SetFocus
                  txt1_GotFocus 9
                  Exit Sub
               End If
            End If
      Case 11
         If blnClkSure = False Then
            If Left(Me.txt1(10).Text, 6) <> Left(Me.txt1(11).Text, 6) Then
               MsgBox "•”Ω–§H•N∏π´e§ªΩX•≤∂∑¨€¶P!!!", vbExclamation + vbOKOnly
               Me.txt1(10).SetFocus
               txt1_GotFocus 10
               Exit Sub
            End If
            If RunNick(txt1(10), txt1(11)) Then
                txt1(10).SetFocus
                txt1_GotFocus (10)
                Exit Sub
             End If
         Else
            blnClkSure = False
         End If
      Case 13
         If blnClkSure = False Then
            If Left(Me.txt1(12).Text, 6) <> Left(Me.txt1(13).Text, 6) Then
               MsgBox "•N≤z§H•N∏π´e§ªΩX•≤∂∑¨€¶P!!!", vbExclamation + vbOKOnly
               Me.txt1(12).SetFocus
               txt1_GotFocus 12
               Exit Sub
            End If
           If RunNick(txt1(12), txt1(13)) Then
               txt1(12).SetFocus
               txt1_GotFocus (12)
               Exit Sub
           End If
         Else
            blnClkSure = False
         End If
      Case 14
         If Me.txt1(14).Text <> "" Then
           Select Case Val(txt1(14))
           Case 1
           Case 2, 3
                If Len(txt1(10)) = 0 And Len(txt1(11)) = 0 And Len(txt1(12)) = 0 And Len(txt1(13)) = 0 Then
                   s = MsgBox("¶C¶LÆÊ¶°¨∞ 2 ©Œ 3 Æ…, •”Ω–§H•N∏π∞œ∂°©Œ•N≤z§H•N∏π∞œ∂°§£•i™≈•’!!", , "USER øÈ§Jø˘ª~")
                   txt1(14).SetFocus
                   txt1(14).SelStart = 0
                   txt1(14).SelLength = Len(txt1(14))
                   Exit Sub
                End If
           Case Else
                s = MsgBox("≥¯™ÌÆÊ¶°•uØ‡ 1 ©Œ 2 ©Œ 3 !!", , "USER øÈ§Jø˘ª~")
                txt1(14).SetFocus
                txt1(14).SelStart = 0
                txt1(14).SelLength = Len(txt1(14))
                Exit Sub
           End Select
         End If
   End Select
End Sub

Public Function Get060308_1(ByRef Strindex As String) As String           'Æ÷ΩZ§H        NICK
   '∂«§J¡`¶¨§Â∏π
   Dim SQLTEMP As String
   strSql = "SELECT EP04 FROM ENGINEERPROGRESS WHERE EP02='" & Strindex & "' "
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
       If CheckStr(AdoRecordSet3.Fields(0)) <> "" Then
           Get060308_1 = CheckStr(AdoRecordSet3.Fields(0))
       Else
           Get060308_1 = ""
       End If
   Else
       Get060308_1 = ""
   End If
   CheckOC3
End Function

Private Sub RefreshColData(Index As Integer, strKind As String)
   Select Case Index
   Case 1 '§§§Â≥¯™Ì
      If strKind = "•”Ω–§H" Then
   'Modify by Morgan 2005/1/28 ®˙ßO¶W•H´K®œ•ŒµÍ¿¿∏ÍÆ∆™Ì
   '      m_ColCustName = "Nvl(CU04,Nvl(CU05,CU06)),Decode(CU04,Null,Decode(CU05,Null,CU88,''),''),Decode(CU04,Null,Decode(CU05,Null,CU89,''),''),Decode(CU04,Null,Decode(CU05,Null,CU90,''),'')"
   '      m_ColCustAdd = "Nvl(CU23,Nvl(CU65,Nvl(CU24,CU29))),Decode(CU23,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU25),CU66),''),Decode(CU23,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU26),CU67),''),Decode(CU23,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU27),CU68),''),Decode(CU23,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU28),CU69),'')"
         m_ColCustName = "Nvl(CU04,Nvl(CU05,CU06)) CN01,Decode(CU04,Null,Decode(CU05,Null,CU88,''),'') CN02,Decode(CU04,Null,Decode(CU05,Null,CU89,''),'') CN03,Decode(CU04,Null,Decode(CU05,Null,CU90,''),'') CN04"
         m_ColCustAdd = "Nvl(CU23,Nvl(CU65,Nvl(CU24,CU29))) CA01,Decode(CU23,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU25),CU66),'') CA02,Decode(CU23,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU26),CU67),'') CA03,Decode(CU23,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU27),CU68),'') CA04,Decode(CU23,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU28||rtrim(' '||cu102)),CU69),'') CA05"
   '2005/1/28 END
      Else '•N≤z§H
   'Modify by Morgan 2005/1/28 ®˙ßO¶W•H´K®œ•ŒµÍ¿¿∏ÍÆ∆™Ì
   '      m_ColAgName = "Nvl(FA04,Nvl(FA05,FA06)),Decode(FA04,Null,Decode(FA05,Null,'',FA63),''),Decode(FA04,Null,Decode(FA05,Null,'',FA64),''),Decode(FA04,Null,Decode(FA05,Null,'',FA65),'')"
   '      m_ColAgAdd = "Nvl(FA17,Nvl(FA32,Nvl(FA18,FA23))),Decode(FA17,Null,Decode(FA32,Null,Decode(FA18,Null,'',FA19),FA33),''),Decode(FA17,Null,Decode(FA32,Null,Decode(FA18,Null,'',FA20),FA34),''),Decode(FA17,Null,Decode(FA32,Null,Decode(FA18,Null,'',FA21),FA35),''),Decode(FA17,Null,Decode(FA32,Null,Decode(FA18,Null,'',FA22),FA36),'')"
         m_ColAgName = "Nvl(FA04,Nvl(FA05,FA06)) AN01,Decode(FA04,Null,Decode(FA05,Null,'',FA63),'') AN02,Decode(FA04,Null,Decode(FA05,Null,'',FA64),'') AN03,Decode(FA04,Null,Decode(FA05,Null,'',FA65),'') AN04"
         m_ColAgAdd = "Nvl(FA17,Nvl(FA32,Nvl(FA18,FA23))) AA01,Decode(FA17,Null,Decode(FA32,Null,Decode(FA18,Null,'',FA19),FA33),'') AA02,Decode(FA17,Null,Decode(FA32,Null,Decode(FA18,Null,'',FA20),FA34),'') AA03,Decode(FA17,Null,Decode(FA32,Null,Decode(FA18,Null,'',FA21),FA35),'') AA04,Decode(FA17,Null,Decode(FA32,Null,Decode(FA18,Null,'',fa22||rtrim(' '||fa70)),FA36),'') AA05"
   '2005/1/28 END
      End If
   Case 2 '≠^§Â≥¯™Ì
      If strKind = "•”Ω–§H" Then
   'Modify by Morgan 2005/1/28 ®˙ßO¶W•H´K®œ•ŒµÍ¿¿∏ÍÆ∆™Ì
   '      m_ColCustName = "Nvl(CU05,CU06),Decode(CU05,Null,'',CU88),Decode(CU05,Null,'',CU89),Decode(CU05,Null,'',CU90)"
   '      m_ColCustAdd = "Nvl(CU65,Nvl(CU24,CU29)),Decode(CU65,Null,Decode(CU24,Null,'',CU25),CU66),Decode(CU65,Null,Decode(CU24,Null,'',CU26),CU67),Decode(CU65,Null,Decode(CU24,Null,'',CU27),CU68),Decode(CU65,Null,Decode(CU24,Null,'',CU28),CU69)"
         m_ColCustName = "Nvl(CU05,CU06) CN01,Decode(CU05,Null,'',CU88),Decode(CU05,Null,'',CU89) CN02,Decode(CU05,Null,'',CU90) CN03"
         m_ColCustAdd = "Nvl(CU65,Nvl(CU24,CU29)) CA01,Decode(CU65,Null,Decode(CU24,Null,'',CU25),CU66) CA02,Decode(CU65,Null,Decode(CU24,Null,'',CU26),CU67) CA03,Decode(CU65,Null,Decode(CU24,Null,'',CU27),CU68) CA04,Decode(CU65,Null,Decode(CU24,Null,'',CU28||rtrim(' '||cu102)),CU69) CA05"
   '2005/1/28 END
      Else '•N≤z§H
   'Modify by Morgan 2005/1/28 ®˙ßO¶W•H´K®œ•ŒµÍ¿¿∏ÍÆ∆™Ì
   '      m_ColAgName = "Nvl(FA05,FA06),Decode(FA05,Null,'',FA63),Decode(FA05,Null,'',FA64),Decode(FA05,Null,'',FA65)"
   '      m_ColAgAdd = "Nvl(FA32,Nvl(FA18,FA23)),Decode(FA32,Null,Decode(FA18,Null,'',FA19),FA33),Decode(FA32,Null,Decode(FA18,Null,'',FA20),FA34),Decode(FA32,Null,Decode(FA18,Null,'',FA21),FA35),Decode(FA32,Null,Decode(FA18,Null,'',FA22),FA36)"
         m_ColAgName = "Nvl(FA05,FA06) AN01,Decode(FA05,Null,'',FA63) AN02,Decode(FA05,Null,'',FA64) AN03,Decode(FA05,Null,'',FA65) AN04"
         m_ColAgAdd = "Nvl(FA32,Nvl(FA18,FA23)) AA01,Decode(FA32,Null,Decode(FA18,Null,'',FA19),FA33) AA02,Decode(FA32,Null,Decode(FA18,Null,'',FA20),FA34) AA03,Decode(FA32,Null,Decode(FA18,Null,'',FA21),FA35) AA04,Decode(FA32,Null,Decode(FA18,Null,'',fa22||rtrim(' '||fa70)),FA36) AA05"
   '2005/1/28 END
      End If
   Case 3 '§È§Â≥¯™Ì
      If strKind = "•”Ω–§H" Then
   'Modify by Morgan 2005/1/28 ®˙ßO¶W•H´K®œ•ŒµÍ¿¿∏ÍÆ∆™Ì
   '      m_ColCustName = "Nvl(CU06,CU05),Decode(CU06,Null,CU88,''),Decode(CU06,Null,CU89,''),Decode(CU06,Null,CU90,'')"
   '      m_ColCustAdd = "Nvl(CU29,Nvl(CU65,CU24)),Decode(CU29,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU25),CU66),''),Decode(CU29,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU26),CU67),''),Decode(CU29,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU27),CU68),''),Decode(CU29,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU28),CU69),'')"
         m_ColCustName = "Nvl(CU06,CU05) CN01,Decode(CU06,Null,CU88,'') CN02,Decode(CU06,Null,CU89,'') CN03,Decode(CU06,Null,CU90,'') CN04"
         m_ColCustAdd = "Nvl(CU29,Nvl(CU65,CU24)) CA01,Decode(CU29,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU25),CU66),'') CA02,Decode(CU29,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU26),CU67),'') CA03,Decode(CU29,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU27),CU68),'') CA04,Decode(CU29,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU28||rtrim(' '||cu102)),CU69),'') CA05"
   '2005/1/28 END
      Else '•N≤z§H
   'Modify by Morgan 2005/1/28 ®˙ßO¶W•H´K®œ•ŒµÍ¿¿∏ÍÆ∆™Ì
   '      m_ColAgName = "Nvl(FA06,FA05),Decode(FA06,Null,FA63,''),Decode(FA06,Null,FA64,''),Decode(FA06,Null,FA65,'')"
   '      m_ColAgAdd = "Nvl(FA23,Nvl(FA32,FA18)),Decode(FA23,Null,Decode(FA32,Null,FA19,FA33),''),Decode(FA23,Null,Decode(FA32,Null,FA20,FA34),''),Decode(FA23,Null,Decode(FA32,Null,FA21,FA35),''),Decode(FA23,Null,Decode(FA32,Null,FA22,FA36),'')"
         m_ColAgName = "Nvl(FA06,FA05) AN01,Decode(FA06,Null,FA63,'') AN02,Decode(FA06,Null,FA64,'') AN03,Decode(FA06,Null,FA65,'') AN04"
         m_ColAgAdd = "Nvl(FA23,Nvl(FA32,FA18)) AA01,Decode(FA23,Null,Decode(FA32,Null,FA19,FA33),'') AA02,Decode(FA23,Null,Decode(FA32,Null,FA20,FA34),'') AA03,Decode(FA23,Null,Decode(FA32,Null,FA21,FA35),'') AA04,Decode(FA23,Null,Decode(FA32,Null,fa22||rtrim(' '||fa70),FA36),'') AA05"
   '2005/1/28 END
      End If
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) = "" Then Exit Sub
   Select Case Index
      Case 1, 2, 3, 4 '•ª©“¥¡≠≠∞_, ®¥, ™k©w¥¡≠≠∞_, ®¥
         Cancel = Not ChkDate(txt1(Index))
         If Cancel Then TextInverse txt1(Index)
      'Add By Sindy 2023/5/23
      Case 5
         txt1(15).Enabled = True
         If txt1(15) = "" Then txt1(15) = "1"
         If txt1(Index) = "2" Then
            txt1(15) = ""
            txt1(15).Enabled = False
         End If
         '2023/5/23 END
   End Select
End Sub
