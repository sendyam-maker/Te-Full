VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm170238 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "¦~«×¦U¶µ©Ò±o©ú²Ó"
   ClientHeight    =   5770
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5770
   ScaleWidth      =   8930
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   8600
      _ExtentX        =   15152
      _ExtentY        =   7426
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "©Ò±o¸ê®Æ"
      TabPicture(0)   =   "frm170238.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grdDataList1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "¦Û¥I«O¶O"
      TabPicture(1)   =   "frm170238.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grdDataList2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList1 
         Height          =   3700
         Left            =   -74900
         TabIndex        =   11
         Top             =   360
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   6544
         _Version        =   393216
         BackColor       =   -2147483628
         Cols            =   13
         FixedCols       =   3
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "·s²Ó©úÅé"
            Size            =   9.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   13
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList2 
         Height          =   3700
         Left            =   100
         TabIndex        =   12
         Top             =   360
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   6544
         _Version        =   393216
         BackColor       =   -2147483628
         Cols            =   13
         FixedCols       =   3
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "·s²Ó©úÅé"
            Size            =   9.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   13
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4635
      Top             =   0
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   1845
      Max             =   200
      Min             =   150
      TabIndex        =   5
      Top             =   5490
      Value           =   200
      Width           =   4785
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   765
      MaxLength       =   3
      TabIndex        =   0
      Top             =   510
      Width           =   600
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7920
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   2
      Top             =   420
      Width           =   800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "¬d¸ß(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7020
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   1
      Top             =   420
      Width           =   800
   End
   Begin MSForms.ComboBox cboUser 
      Height          =   300
      Left            =   765
      TabIndex        =   13
      Top             =   120
      Width           =   2400
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4233;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblStaffNo 
      AutoSize        =   -1  'True
      Caption         =   "­û¤u¡G"
      Height          =   180
      Left            =   225
      TabIndex        =   9
      Top             =   180
      Width           =   540
   End
   Begin VB.Label lblTimeOut 
      Appearance      =   0  '¥­­±
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "­Y±z¥¼Ä~Äò²¾°Ê·Æ¹«,±N·|©ó 59 ¬í«áµn¥X"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5220
      TabIndex        =   8
      Top             =   90
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Á~¸ê¸ê®Æ¿@²H³]©w"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   5520
      Width           =   1560
   End
   Begin VB.Label lblTest 
      Alignment       =   2  '¸m¤¤¹ï»ô
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000005&
      BorderStyle     =   1  '³æ½u©T©w
      Caption         =   "³o¬O¿@²H³]©w¹wÄý"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6795
      TabIndex        =   6
      Top             =   5520
      Width           =   1905
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "PS: ÂI¿ïªí®æ¤º¼Æ¦rÄæ¦ì¥i¬d¸ß©ú²Ó"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   5250
      Width           =   3030
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "¦~«×¡G"
      Height          =   180
      Index           =   0
      Left            =   225
      TabIndex        =   3
      Top             =   555
      Width           =   540
   End
End
Attribute VB_Name = "frm170238"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/10 Form2.0¤w­×§ï(cboUser)
'created by sonia 2016/2/25
Option Explicit
Dim m_iCol As Integer, m_iRow As Integer
Dim m_StaffNoCon As String

Private Sub cboUser_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboUser_Validate(Cancel As Boolean)
   Dim ii As Integer
   For ii = 0 To cboUser.ListCount - 1
      If InStr(cboUser.List(ii), cboUser) > 0 Then
         cboUser.ListIndex = ii
      End If
   Next
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdSearch_Click()
      
   If txtYear = "" Then
      MsgBox "½Ð¿é¤J¦~«×¡I", vbExclamation
      txtYear.SetFocus
      Exit Sub
   End If
   
   '©|¥¼±Ò¥Î
   If Val(Pub_MaxYBYear) = 0 Then
      MsgBox "©|µL¥ô¦ó¦©Ãº¸ê®Æ¥i¬d¸ß¡I", vbExclamation
      txtYear.SetFocus
      Exit Sub
   End If
   
   '¤£¥i¦­©ó±Ò¥Î¦~¤ë
   If Val(txtYear) < Val(Left(Pub_StartYM, 4) - 1912) Then
      MsgBox "¤£¥i¦­©ó±Ò¥Î¦~«× " & Val(Left(Pub_StartYM, 4) - 1912) & " ¦~¡I", vbExclamation
      txtYear.SetFocus
      Exit Sub
   End If
   
   '¤£¥i¤j©ó¦~²×¤J±b³Ì¤j¦~«×,­Y¬°·í¦~«×®É¨t²Î¤é¥²¶·>3/1
   If Val(txtYear) = Val(Pub_MaxYBYear) - 1911 And Val(Right(strSrvDate(2), 4)) < 301 Then
      MsgBox Val(txtYear) & "¦~«×¸ê®Æ©ó3/1°_¶}©ñ¬d¸ß¡I", vbExclamation
      txtYear.SetFocus
      Exit Sub
   ElseIf Val(txtYear) > Val(Pub_MaxYBYear) - 1912 Then
      MsgBox "¤£¥i¤j©ó¤w¶}©ñ¦~«× " & Val((Pub_MaxYBYear) - 1912) & " ¦~¡I", vbExclamation
      txtYear.SetFocus
      Exit Sub
   End If
   
   cboUser_Validate False
   If cboUser.ListIndex < 0 Then
      MsgBox "½Ð¿ï¾Ü­û¤u¡I", vbExclamation
      cboUser.SetFocus
      Exit Sub
   End If
   
   QueryData

End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   PUB_AddSalaryUser cboUser, False
   
   If Val(Pub_MaxYBYear) <> 0 Then
      txtYear = Val(Pub_MaxYBYear) - 1912
   Else
      txtYear = Val(Left(Pub_StartYM, 4) - 1912)
   End If
   
   PUB_SetForeColorScroll HScroll1
   SetGridColor
   PUB_EnableSalaryTimer
   SSTab1.Tab = 0
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170238 = Nothing
End Sub

Private Sub SetGridColor()
   lblTest.ForeColor = PUB_GetColor(HScroll1.Value)
   grdDataList1.ForeColor = lblTest.ForeColor
   grdDataList2.ForeColor = lblTest.ForeColor
End Sub

Private Sub HScroll1_Change()
   HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
   SetGridColor
   PUB_SaveForeColor HScroll1
End Sub

Private Sub Timer1_Timer()
   PUB_ShowSalaryCountDown lblTimeOut
End Sub

Private Sub txtYear_GotFocus()
   TextInverse txtYear
   CloseIme
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

'©Ò±o¸ê®Æ
Private Sub SetGrid1()
Dim iCol As Integer, iRow As Integer
Dim varGridWidth() As Variant
   
   varGridWidth = Array(1100, 2250, 1100, 1100, 1100, 1100, 700, 700, 1100, 1600)
   SetGridDataListWidth grdDataList1, varGridWidth()
   With grdDataList1
      .ColWordWrapOption(0) = True '¦Û°Ê§é¦æ
     .Visible = False
      
      .ColAlignmentFixed(0) = flexAlignCenterCenter
      .ColAlignment(0) = flexAlignCenterCenter
      .RowHeight(0) = 1.7 * .RowHeight(1)
      For iCol = 0 To .Cols - 1
         Select Case iCol
            Case 0, 1, 2, 8
               .ColAlignmentFixed(iCol) = flexAlignCenterCenter
               .ColAlignment(iCol) = flexAlignLeftCenter
            Case 6, 7
               .ColAlignmentFixed(iCol) = flexAlignCenterCenter
               .ColAlignment(iCol) = flexAlignCenterCenter
            Case Else
               .ColAlignmentFixed(iCol) = flexAlignCenterCenter
               .ColAlignment(iCol) = flexAlignRightCenter
         End Select
      Next
      For iRow = 1 To .Rows - 1
         .RowHeight(iRow) = 1.7 * .RowHeight(iRow)
      Next
     
     .Visible = True
   End With

End Sub

'¦Û¥I«O¶O
Private Sub SetGrid2()
Dim iCol As Integer, iRow As Integer
Dim varGridWidth() As Variant
   
   '“³ÂÃ­û¤u½s¸¹
   varGridWidth = Array(0, 1100, 2250, 1100, 1100, 1100)
   SetGridDataListWidth grdDataList2, varGridWidth()
   With grdDataList2
      .ColWordWrapOption(0) = True '¦Û°Ê§é¦æ
     .Visible = False
      
      .ColAlignmentFixed(0) = flexAlignCenterCenter
      .ColAlignment(0) = flexAlignCenterCenter
      .RowHeight(0) = 1.7 * .RowHeight(1)
      For iCol = 0 To .Cols - 1
         Select Case iCol
            Case 0, 1, 2
               .ColAlignmentFixed(iCol) = flexAlignCenterCenter
               .ColAlignment(iCol) = flexAlignLeftCenter
            Case Else
               .ColAlignmentFixed(iCol) = flexAlignCenterCenter
               .ColAlignment(iCol) = flexAlignRightCenter
         End Select
      Next
     
      For iRow = 1 To .Rows - 1
         .RowHeight(iRow) = 1.7 * .RowHeight(iRow)
      Next
     
     .Visible = True
   End With
End Sub

Private Sub QueryData()
Dim arrTmp() As String
   
   arrTmp = Split(cboUser, " ")
   cboUser.Tag = arrTmp(1)
   
   '­nÅª¥X©Ò¦³¨­¥÷ÃÒ¦rÃÒ»Pµe­±©Ò¿ï¤H­û¬Û¦Pªº©Ò¦³©Ò±o¸ê®Æ
   '54ªÑ§QµL³Ò°hª÷¦Û´£(¸ÓÄæ¬°ªÑ§Q²bÃB
   strExc(0) = "SELECT ID25 ©Ò±o¤H¥N¸¹,A0802 ¤½¥q¦WºÙ,DECODE(ID05,'50','50Á~¸ê©Ò±o','51','51¯²ª÷','54','54ªÑ§Q','92','92ºÖ§Qª÷¡þ¦þª÷','93','93°hÂ¾©Ò±o','9A','9A°õ¦æ·~°È©Ò±o','9B','9BÁ¿ºt¶O',ID05) ©Ò±o®æ¦¡," & _
               " TO_CHAR(ID08,'99,999,999') µ¹¥IÁ`ÃB©ÎªÑ§QÁ`ÃB,TO_CHAR(ID09,'99,999,999') ¦©Ãºµ|ÃB©Î¥i¦©©èµ|ÃB,TO_CHAR(ID10,'99,999,999') µ¹¥I²bÃB©ÎªÑ§Q²bÃB,TO_CHAR(ID22) °_©l¤ë,TO_CHAR(ID23) ºI¤î¤ë," & _
               " DECODE(ID27,'1','ªÑ§Q','2','°õ¦æ·~°È©Ò±o¤ÎÁ¿ºt¶O','3','Á~¸ê©Ò±o','5','¯²ª÷','9','°hÂ¾©Ò±o','A','¨ä¥L',ID27) ©Ò±oÃþ§O,TO_CHAR(TO_NUMBER(DECODE(ID05,'54',0,DECODE(SUBSTR(ID17,1,1),'0',SUBSTR(ID17,1,10),0))),'99,999,999') ³Ò°hª÷¦Û´£,A0801 FROM INCOMEDATA,ACC080 " & _
               " WHERE ID14=" & Val(txtYear) + 1911 & " AND ID03=A0807(+) AND ID06 IN (SELECT ST26 FROM STAFF WHERE ST01='" & cboUser.Tag & "') " & _
               " UNION ALL SELECT 'Á`­p',NULL,NULL,TO_CHAR(SUM(NVL(ID08,0)),'99,999,999'),TO_CHAR(SUM(NVL(ID09,0)),'99,999,999'),TO_CHAR(SUM(NVL(ID10,0)),'99,999,999'),NULL,NULL,NULL,TO_CHAR(SUM(TO_NUMBER(DECODE(ID05,'54',0,DECODE(SUBSTR(ID17,1,1),'0',SUBSTR(ID17,1,10),0)))),'99,999,999'),NULL FROM INCOMEDATA " & _
               " WHERE ID14=" & Val(txtYear) + 1911 & " AND ID06 IN (SELECT ST26 FROM STAFF WHERE ST01='" & cboUser.Tag & "') ORDER BY 1,A0801,3"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With grdDataList1
      .FixedCols = 0
      Set .Recordset = RsTemp.Clone
      .FixedCols = 3
      End With
      SetGrid1
   End If
   
   'modify by sonia 2023/12/7 ­nÅª¥X©Ò¦³¨­¥÷ÃÒ¦rÃÒ»Pµe­±©Ò¿ï¤H­û¬Û¦Pªº©Ò¦³©Ò±o¸ê®Æ
   '¦U¶µ«O¶O ¸É¥R«O¶O§ï¨Ì©Ò±o¤H¥N¸¹§ì¸ê®Æ 201601¥H«á¤w¸g¨S¦³²Ä¤G®a¤½¥qªº¸ê®Æ,¬G¨ú®øsubstr(sm01,1,1)||replace(substr(sm01,2),'A','0')ªº¼gªk
   'strExc(0) = "select substr(sm01,1,1)||replace(substr(sm01,2),'A','0') ­û¤u½s¸¹,sm01 ©Ò±o¤H¥N¸¹,a0802 ¤½¥q¦WºÙ,TO_CHAR(x.SL,'99,999,999') ³Ò«O¶O,TO_CHAR(x.SH,'99,999,999') °·«O¶O,TO_CHAR(y.SH2,'99,999,999') ¸É¥R«O¶O " & _
               " from (SELECT sm01,substr(sm01,1,1)||replace(substr(sm01,2),'A','0') sid,sm37,SUM(SM14) sL,SUM(SM15) sH from (select sm01,sm02,sm14,sm15,sm37 FROM SALARYMONTH " & _
               " WHERE NVL(SM14,0)+NVL(SM15,0)>0 and substr(sm02,1,4)=" & Val(txtYear) + 1911 & " and substr(sm01,1,1)||replace(substr(sm01,2),'A','0')>='" & cboUser.Tag & "' and substr(sm01,1,1)||replace(substr(sm01,2),'A','0')<='" & cboUser.Tag & "'" & _
               " union all select od03,od14,decode(od04,'31',od05,'32',-1*od05) sm14,decode(od04,'35',od05,'36',-1*od05) sm15,sm37 from salarymonth,othersalarydata" & _
               " where od03(+)=sm01 and od14(+)=sm02 and od04 in ('31','32','35','36') and substr(sm02,1,4)=" & Val(txtYear) + 1911 & " and substr(sm01,1,1)||replace(substr(sm01,2),'A','0')>='" & cboUser.Tag & "' and substr(sm01,1,1)||replace(substr(sm01,2),'A','0')<='" & cboUser.Tag & "'" & _
               " ) GROUP BY sm01,sm37) x,Staff s1, " & _
               " (select min(st01) y0,nvl(st26,st01) y1,nhi11 y2,sum(nhi06) sH2 from nhi2nd,staff where st01(+)=nhi01 and st01>'6' and substr(nhi02,1,4)=" & Val(txtYear) + 1911 & " group by nvl(st26,st01),nhi11" & _
               " ) y,acc080 where y0(+)=sid and y2(+)=sm37 and s1.st01(+)=sid and sm37=a0801(+) " & _
               "union all select null,'Á`­p',null,TO_CHAR(sum(x.SL),'99,999,999'),TO_CHAR(sum(x.SH),'99,999,999'),TO_CHAR(sum(y.SH2),'99,999,999') " & _
               " from (SELECT sm01,substr(sm01,1,1)||replace(substr(sm01,2),'A','0') sid,sm37,SUM(SM14) sL,SUM(SM15) sH from (select sm01,sm02,sm14,sm15,sm37 FROM SALARYMONTH " & _
               " WHERE NVL(SM14,0)+NVL(SM15,0)>0 and substr(sm02,1,4)=" & Val(txtYear) + 1911 & " and substr(sm01,1,1)||replace(substr(sm01,2),'A','0')>='" & cboUser.Tag & "' and substr(sm01,1,1)||replace(substr(sm01,2),'A','0')<='" & cboUser.Tag & "'" & _
               " union all select od03,od14,decode(od04,'31',od05,'32',-1*od05) sm14,decode(od04,'35',od05,'36',-1*od05) sm15,sm37 from salarymonth,othersalarydata" & _
               " where od03(+)=sm01 and od14(+)=sm02 and od04 in ('31','32','35','36') and substr(sm02,1,4)=" & Val(txtYear) + 1911 & " and substr(sm01,1,1)||replace(substr(sm01,2),'A','0')>='" & cboUser.Tag & "' and substr(sm01,1,1)||replace(substr(sm01,2),'A','0')<='" & cboUser.Tag & "'" & _
               " ) GROUP BY sm01,sm37) x,Staff s1, " & _
               " (select min(st01) y0,nvl(st26,st01) y1,nhi11 y2,sum(nhi06) sH2 from nhi2nd,staff where st01(+)=nhi01 and st01>'6' and substr(nhi02,1,4)=" & Val(txtYear) + 1911 & " group by nvl(st26,st01),nhi11" & _
               " ) y,acc080 where y0(+)=sid and y2(+)=sm37 and s1.st01(+)=sid and sm37=a0801(+) " & _
               " ORDER BY 1,2,3"
   '¦U¶µ«O¶O
   strExc(0) = "select sm01 ­û¤u½s¸¹,sm01 ©Ò±o¤H¥N¸¹,a0802 ¤½¥q¦WºÙ,TO_CHAR(sum(sL),'99,999,999') ³Ò«O¶O,TO_CHAR(sum(sH),'99,999,999') °·«O¶O,TO_CHAR(sum(sH2),'99,999,999') ¸É¥R«O¶O  from acc080," & _
               " (SELECT sm01,sm01 sid,sm37,SUM(SM14) sL,SUM(SM15) sH,0 sH2 from " & _
               " (select sm01,sm02,sm14,sm15,sm37 FROM SALARYMONTH,staff WHERE NVL(SM14,0)+NVL(SM15,0)>0 and substr(sm02,1,4)=" & Val(txtYear) + 1911 & " and sm01=st01(+) and st26 in (select st26 from staff where st01='" & cboUser.Tag & "')" & _
               " union all select od03,od14,decode(od04,'31',od05,'32',-1*od05) sm14,decode(od04,'35',od05,'36',-1*od05) sm15,sm37 from salarymonth,othersalarydata,staff where od03(+)=sm01 and od14(+)=sm02 and od04 in ('31','32','35','36') and substr(sm02,1,4)=" & Val(txtYear) + 1911 & " and sm01=st01(+) and st26 in (select st26 from staff where st01='" & cboUser.Tag & "')) GROUP BY sm01,sm37" & _
               " union all select min(st01) st01,nhi01 sid,nhi11 sm37,0 sL,0 sH,sum(nhi06) sH2 from nhi2nd,staff where substr(nhi02,1,4)=" & Val(txtYear) + 1911 & " and nhi01=st01(+) and st26 in (select st26 from staff where st01='" & cboUser.Tag & "') group by nhi01,nhi11 " & _
               " union all select 'Á`­p',null,null,sum(SL),sum(SH),sum(SH2) " & _
               " from (SELECT sm01,sm01 sid,sm37,SUM(SM14) sL,SUM(SM15) sH,0 sH2 from " & _
               " (select sm01,sm02,sm14,sm15,sm37 FROM SALARYMONTH,staff WHERE NVL(SM14,0)+NVL(SM15,0)>0 and substr(sm02,1,4)=" & Val(txtYear) + 1911 & " and sm01=st01(+) and st26 in (select st26 from staff where st01='" & cboUser.Tag & "')" & _
               "  union all select od03,od14,decode(od04,'31',od05,'32',-1*od05) sm14,decode(od04,'35',od05,'36',-1*od05) sm15,sm37 from salarymonth,othersalarydata,staff where od03(+)=sm01 and od14(+)=sm02 and od04 in ('31','32','35','36') and substr(sm02,1,4)=" & Val(txtYear) + 1911 & " and sm01=st01(+) and st26 in (select st26 from staff where st01='" & cboUser.Tag & "')) GROUP BY sm01,sm37" & _
               " union all select min(st01) st01,nhi01 sid,nhi11 sm37,0 sL,0 sH,sum(nhi06) sH2 from nhi2nd,staff where substr(nhi02,1,4)=" & Val(txtYear) + 1911 & " and nhi01=st01(+) and st26 in (select st26 from staff where st01='" & cboUser.Tag & "') group by nhi01,nhi11 ) " & _
               " ) where sm37=a0801(+) group by sm01,a0802 having sum(sL)+sum(sH)+sum(sH2)>0 ORDER BY 1,2"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With grdDataList2
      .FixedCols = 0
      Set .Recordset = RsTemp.Clone
      .FixedCols = 3
      End With
      SetGrid2
   End If
   
   SSTab1.Tab = 0
   
End Sub

Private Sub grdDataList1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_ResetSalaryTimer Me
End Sub

Private Sub grdDataList2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_ResetSalaryTimer Me
End Sub

