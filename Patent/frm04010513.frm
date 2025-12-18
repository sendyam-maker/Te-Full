VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010513 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "¥N²z¤H¤w¶¶½Z"
   ClientHeight    =   5748
   ClientLeft      =   -600
   ClientTop       =   2988
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9336
   Begin VB.TextBox txtCode 
      Height          =   288
      Index           =   4
      Left            =   3450
      MaxLength       =   2
      TabIndex        =   3
      Top             =   660
      Width           =   492
   End
   Begin VB.TextBox txtCode 
      Height          =   288
      Index           =   3
      Left            =   3060
      MaxLength       =   1
      TabIndex        =   2
      Top             =   660
      Width           =   372
   End
   Begin VB.TextBox txtCode 
      Height          =   288
      Index           =   2
      Left            =   1830
      MaxLength       =   6
      TabIndex        =   1
      Top             =   660
      Width           =   1212
   End
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   288
      Index           =   1
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   0
      Top             =   660
      Width           =   732
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4212
      Left            =   120
      TabIndex        =   5
      Top             =   1380
      Width           =   9072
      _ExtentX        =   16002
      _ExtentY        =   7430
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "¦¬¤å¸¹|¦¬¤å¤é|µo¤å¤é|®×¥ó©Ê½è|¥N²z¤H|©¼©Ò®×¸¹|¦¬¹F¤é"
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
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "¶¶½Z¸ê®Æ(&F)"
      Default         =   -1  'True
      Height          =   400
      Index           =   2
      Left            =   6255
      TabIndex        =   6
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      Height          =   400
      Index           =   0
      Left            =   7530
      TabIndex        =   7
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8388
      TabIndex        =   8
      Top             =   60
      Width           =   800
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1080
      TabIndex        =   4
      Top             =   1020
      Width           =   8115
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "14314;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblNation 
      Height          =   252
      Left            =   5940
      TabIndex        =   12
      Top             =   660
      Width           =   372
   End
   Begin VB.Label Label11 
      Caption         =   "¥Ó½Ð°ê®a¡G"
      Height          =   252
      Left            =   4980
      TabIndex        =   13
      Top             =   660
      Width           =   972
   End
   Begin MSForms.Label lblCountryName 
      Height          =   270
      Left            =   6300
      TabIndex        =   11
      Top             =   660
      Width           =   2535
      VariousPropertyBits=   27
      Size            =   "4471;476"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "¥»©Ò®×¸¹¡G"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   660
      Width           =   972
   End
   Begin VB.Label Label24 
      Caption         =   "®×¥ó¦WºÙ¡G"
      Height          =   252
      Left            =   120
      TabIndex        =   9
      Top             =   1020
      Width           =   972
   End
End
Attribute VB_Name = "frm04010513"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 §ï¦¨Form2.0 (cboCaseName,lblCountryName)
'Memo By Morgan 2012/12/11 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo by Morgan2010/8/11 ¤é´ÁÄæ¤w­×§ï
'Create by Morgan 2009/11/17
Option Explicit

'intLastRow¤W¤@¦¸¤Ï¥ÕªºRow
'blnOKtoShow¨M©w¬O§_­n¤Ï¥Õ
Dim intLastRow As Integer, blnOKtoShow As Boolean
Dim m_936CP09 As String, m_bolFMP As Boolean
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String, m_AppNo As String
Dim m_Done As Boolean
'2016/10/7 END
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/10/31 ¬O§_¬°¾ÈµØ®×

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If intLastRow > 0 Then
            If MsgBox("¬O§_½T©w­n¸Ñ°£¶¶½Z´Á­­¡I(­Y¸Ó¦¬¤å¸¹¥¼¦¬¹F¤]±N¦Û°Ê§ó·s¬°¨t²Î¤é)", vbYesNo + vbDefaultButton2) = vbYes Then
               'Add By Sindy 2020/7/20
               If m_strIR01 <> "" Then
                  '¤U¸ü«H¥óÀÉ,ÀË¬d«H¥ó¬O§_¶}±Ò¤¤,¥H§K«á­±¤W¶Ç¨÷©v°Ï·|µLªkÀx¦s
                  'Modify By Sindy 2022/11/10 + IIf(lblNation <> ¥xÆW°ê®a¥N¸¹, "PAT", "RX")
                  If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, "", IIf(lblNation <> ¥xÆW°ê®a¥N¸¹, "PAT", "RX"), , True) = False Then
                     Screen.MousePointer = vbDefault
                     Exit Sub
                  End If
               End If
               '2020/7/20 END
               
               If FormSave = True Then
                  'Added by Lydia 2023/10/04 FMP®×«Ý«È¤á³Ì²×«ü¥Ü¬ÛÃö±±ºÞ
                  'If m_bolFMP2 = True Then 'Added by Lydia 2023/10/31 §PÂ_¾ÈµØ®× 'Mark by Lydia 2024/01/31 FMP®×¬O¦b¼Ò²Õ¤º§PÂ_«D¾ÈµØ®×ªºEmail¤£¦P
                     If PUB_ChkFMP970mail("1", txtCode(1), txtCode(2), txtCode(3), txtCode(4)) = True Then
                     End If
                  'End If
                  'end 2023/10/04
                  
                  'Modified by Morgan 2016/9/9 ¤j³°®×¤w¹q¤l¤Æ,§ïFMP®×¤~­n¦L--«~Á¨
                  'If m_936CP09 <> "" Then
                  If m_936CP09 <> "" And m_bolFMP Then
                     'Modified by Morgan 2023/5/24 §ï¦s¨÷©v°Ï¤£¦L¯È¥»--«~Á¨
                     g_PrtForm001.PrintCForm m_936CP09, , , True
                  End If
                  'Add By Sindy 2016/10/7
                  If Me.m_strIR01 <> "" Then
                     Unload Me
                     'Modify By Sindy 2022/5/20
                     'frm04010519.GoNext
                     Forms(0).Tmpfrm04010519.GoNext
                     Set Forms(0).Tmpfrm04010519 = Nothing
                     '2022/5/20 END
                  Else
                  '2016/10/7 END
                     QueryData
                  End If
               End If
            End If
         End If
      Case 1
         Unload Me
      Case 2
         'Add By Sindy 2017/12/27
         If m_strIR01 <> "" Then
            If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> txtCode(1) & txtCode(2) & txtCode(3) & txtCode(4) Then
               MsgBox "«H¥ó¿é¤J¥²¶·»P«H¥ó¥»©Ò®×¸¹(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")¤@­P¡I"
               Exit Sub
            End If
         End If
         '2017/12/27 END
         If TxtValidate = True Then
            QueryData
         End If
   End Select
End Sub

Private Function FormSave() As Boolean
   Dim strNP01 As String, strTmp As String
   Dim cp(1 To 14) As String
   Dim strCP48 As String, strCP14 As String, strFA10 As String, strPA09 As String, strCP10 As String
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd
   
   strNP01 = grdDataList.TextMatrix(intLastRow, 0)
   strSql = "Update Nextprogress set np06='Y' where np01='" & strNP01 & "' AND NP06 IS NULL AND NP07='994'"
   cnnConnection.Execute strSql, intI
   strSql = "update caseprogress set cp46=" & strSrvDate(1) & " where cp09='" & strNP01 & "' and cp46 is null"
   cnnConnection.Execute strSql, intI
   
   '·s¼W¦^ÂÐ©e¥ô¥N²z¤H
   'Modified by Morgan 2012/9/6 +«ü©w´£¥Óªº´Á­­
   strExc(0) = "select cp01,cp02,cp03,cp04,cp05,cp07,cp10,cp12,cp14,pa09,fa10,np09" & _
      " from caseprogress,patent,fagent,nextprogress where cp09='" & strNP01 & "'" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9)" & _
      " and np01(+)=cp09 and np07(+)='995' and np06(+) is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strCP10 = "936"
      With RsTemp
      strFA10 = "" & .Fields("fa10")
      strPA09 = "" & .Fields("pa09")
      cp(1) = .Fields("cp01")
      cp(2) = .Fields("cp02")
      cp(3) = .Fields("cp03")
      cp(4) = .Fields("cp04")
      cp(5) = .Fields("cp05")
      cp(7) = "" & .Fields("cp07")
      cp(9) = strNP01 'Added by Lydia 2017/11/01 ¤º±M-¥N²z¤H¤w¶¶½Zªº©Ó¿ì¤H¨S¦³¹w³](P-118690, BA6039998)
      'Added by Morgan 2012/9/6 ­Y¦³«ü©w´£¥Ó¤é«h¥H¸Ó¤é¬°ªk­­
      If .Fields("NP09") > 0 Then
         cp(7) = "" & .Fields("np09")
      End If
      cp(10) = .Fields("cp10")
      '2009/12/30 MODIFY BY SONIA
      'cp(12) = .Fields("cp12")
      cp(13) = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
      cp(12) = GetSalesArea(cp(13))
      '2009/12/30 END
      cp(14) = "" & .Fields("cp14")
      
      'FMP
      'Modified by Lydia 2023/10/31 §PÂ_«D¥xÆW®× strPA09 <> "000"
      'Modified by Lydia 2025/09/30 §ï¥Î¼Ò²Õ
      'If Left(cp(12), 1) = "F" And strPA09 <> "000" Then
      '   m_bolFMP = True 'Added by Morgan 2016/9/9
      m_bolFMP = PUB_ChkIsFMP(cp(1), cp(2), cp(3), cp(4), strPA09)
      If m_bolFMP = True Then
      'end 2025/09/30
         'Modified by Morgan 2017/10/11 FMP¹w³]©Ó¿ì¤H¤ñ·ÓFCP
         'strCP14 = PUB_GetFmpCP14(cp, strFA10)
         strCP14 = PUB_GetFCPPromoterNo(cp(9), "936", "" & .Fields("cp14"))
         'end 2017/10/11
         '·s®×§ì³]©w
         If InStr("101,102,103", cp(10)) > 0 Then
            strCP48 = PUB_GetDeadLine(cp(5), cp(7), 6)
         Else
            '¥Ø«e³]7¤Ñ
            strCP48 = Pub_GetHandleDay("FCP", strPA09, strCP10)
         End If
         'Add by Morgan 2011/7/15 FMP¥»©Ò´Á­­³]¬°¨t²Î¤é+7¤é“ï¤Ñ
         cp(6) = CompDate(2, 7, strSrvDate(1))
         'Added by Morgan 2022/4/1 ¥»©Ò´Á­­­n¦A§ì¤u§@¤é--Sharon
         cp(6) = PUB_GetWorkDay1(cp(6), True)
         If Val(cp(6)) > Val(cp(7)) Then cp(6) = cp(7) '¤£¥i±ß©óªk­­
      Else
         'm_bolFMP = False 'Added by Morgan 2016/9/9 'Mark by Lydia 2025/09/30
         'Modified by Morgan 2018/11/29
         '­Y©Ó¿ì¤H¬°µ{§Ç®É§ï§ì°ê¤º®×¤uµ{®v
         'strCP14 = cp(14)
         strCP14 = ""
         If GetStaffDepartment(cp(14)) = "P12" Then
            strCP14 = PUB_GetInCaseCP14(cp(1), cp(2), cp(3), cp(4))
         End If
         If strCP14 = "" Then strCP14 = cp(14)
         'end 2018/11/29
         
         '¥Ø«e³]3¤Ñ
         strCP48 = Pub_GetHandleDay(cp(1), strPA09, strCP10)
         'Added by Morgan 2012/9/6
         '©Ò­­=ªk­­-2¤u§@¤Ñ
         'MODIFY BY SONIA 2014/4/3 «DFMP®×¨ú®øªk­­,¥»©Ò´Á­­¦P©Ó¿ì´Á­­³]¬°3¤u§@¤Ñ
         'If cp(7) <> "" Then
         '   cp(6) = CompDate(2, -1, cp(7))
         '   cp(6) = CompWorkDay(2, cp(6), 1)
         'End If
         cp(7) = ""
         cp(6) = strCP48
         '2014/4/3 END
      End If
      'Added by Lydia 2023/10/31 §PÂ_¾ÈµØ®×
      m_bolFMP2 = False
      If m_bolFMP = True Then
         m_bolFMP2 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, cp(1), cp(2), cp(3), cp(4))
      End If
      'end 2023/10/31
   
      If Val(cp(6)) > 0 And Val(strCP48) > Val(cp(6)) Then strCP48 = cp(6)
      
      'Add by Morgan 2010/10/6 §ó·s¬ÛÃö¦¬¤å¸¹¬Û¦P¥B¬°·|½Z(924)ªº»ô³Æ¤é
      strSql = "Update engineerprogress set ep06=" & strSrvDate(1) & _
         " where ep02 in (select cp09 from caseprogress where cp43='" & strNP01 & "'" & _
         " and cp10='924' and cp27 is null) and ep06 is null"
      cnnConnection.Execute strSql, intI
      'end 2010/10/6
      
      If Val(strCP48) > 0 Then
         If Left(cp(12), 1) = "F" Or PUB_IfSetCP48() Then 'Add by Morgan 2010/10/6
            strSql = "Update caseprogress set cp48=" & strCP48 & " where cp43='" & strNP01 & "'" & _
               " and cp10='924' and cp27 is null and cp48 is null"
            cnnConnection.Execute strSql, intI
         'Add by Morgan 2010/10/6
         Else
            strCP48 = ""
         End If
         'end 2010/10/6
      End If
      
      m_936CP09 = AutoNo("B", 6)
      '2009/12/30 MODIFY BY SONIA
      'strSQL = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05" & _
         ",CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48) " & _
         "select cp01,cp02,cp03,cp04," & strSrvDate(1) & _
         ",'" & m_936CP09 & "','" & strCP10 & "','90',cp12,cp13,'" & strCP14 & "','N','N','N',cp09," & CNULL(strCP48, True) & _
         " from caseprogress where cp09='" & strNP01 & "'"
      'Modify by Morgan 2011/7/15 +CP06,CP07
      strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,cp06,cp07" & _
         ",CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48) " & _
         "select cp01,cp02,cp03,cp04," & strSrvDate(1) & "," & CNULL(cp(6), True) & "," & CNULL(cp(7), True) & _
         ",'" & m_936CP09 & "','" & strCP10 & "','90','" & cp(12) & "','" & cp(13) & "','" & strCP14 & "','N','N','N',cp09," & CNULL(strCP48, True) & _
         " from caseprogress where cp09='" & strNP01 & "'"
      '2009/12/30 END
      cnnConnection.Execute strSql, intI
      
      'Add by Morgan 2010/10/6 §ó·s¦^¥N»ô³Æ¤é
      strSql = "Update engineerprogress set ep06=" & strSrvDate(1) & _
         " where ep02='" & m_936CP09 & "'"
      cnnConnection.Execute strSql, intI
      'end 2010/10/6
      End With
   End If
   
   'Add by Sindy 2016/10/7
   If m_strIR01 <> "" Then
      'Added by Morgan 2018/11/29 ¶¶½Z«H¥ó­nÂk¨÷
      'Modified by Morgan 2018/12/5 «H¥ó­nÂk¨ì¦^ÂÐ©e¥ô¥N²z¤Hµ{§Ç--«~Á¨
      'Modify By Sindy 2022/11/10 + IIf(lblNation <> ¥xÆW°ê®a¥N¸¹, "PAT", "RX")
      If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, m_936CP09, IIf(Pub_StrUserSt03 = "F22", "ALTR", IIf(lblNation <> ¥xÆW°ê®a¥N¸¹, "PAT", "RX"))) = False Then
         GoTo ErrHnd
      End If
      'end 2018/11/29
      
      'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", m_936CP09, "")
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm04010513", IIf(Pub_StrUserSt03 = "F22", m_936CP09, "")
   End If
   '2016/10/7 END
   
  
   cnnConnection.CommitTrans
   FormSave = True
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If
   
End Function

Private Sub QueryData()
   Dim stCode(1 To 4) As String
   stCode(1) = txtCode(1)
   stCode(2) = txtCode(2)
   stCode(3) = Right("0" & txtCode(3), 1)
   stCode(4) = Right("00" & txtCode(4), 2)
   
    'Add by Lydia 2014/10/31 ¶}©ñ¥~±Mµ{§Ç¤H­û(31,33,34)¥i¶i¤J±M§Q³B¨t²Î¾Þ§@FMP¾ÈµØ®×¥ó¡A¦ý«D¦¹Ãþ®×¥ó®É¥~±Mµ{§Ç¤H­û¤£¥i¾Þ§@¡C
     If FMP2open = True Then
        If PUB_FMPtoCheck(0, 1, Pub_strUserST05, stCode(1), stCode(2), stCode(3), stCode(4)) = False Then
          txtCode(2).SetFocus
          Exit Sub
        End If
     End If
   
   'Modified by Morgan 2016/9/9 +CP12
   strExc(0) = "select cp09 ¦¬¤å¸¹,sqldatet(cp05) ¦¬¤å¤é,sqldatet(cp27) µo¤å¤é" & _
      ",cpm04 ®×¥ó©Ê½è,NVL(fa04,NVL(fa05,fa06)) ¥N²z¤H,cp45 ©¼©Ò®×¸¹,sqldatet(cp46) ¦¬¹F¤é,sqldatet(np08) ¶¶½Z´Á­­,CP12" & _
      " from nextprogress,caseprogress,fagent,casepropertymap where np02='" & stCode(1) & "' and np03='" & stCode(2) & "'" & _
      " and np04='" & stCode(3) & "' and np05='" & stCode(4) & "' and np06 is null and np07='994' and cp09(+)=np01 and fa01(+)=substr(cp44,1,8) and fa02(+)=substr(cp44,9) and cpm01(+)=cp01 and cpm02(+)=cp10"
   intI = 1
   Set grdDataList.Recordset = ClsLawReadRstMsg(intI, strExc(0))
   grdDataList.FormatString = "¦¬¤å¸¹|¦¬¤å¤é|µo¤å¤é|®×¥ó©Ê½è|¥N²z¤H|©¼©Ò®×¸¹|¦¬¹F¤é|¶¶½Z´Á­­"
   SetDataListWidth
   intLastRow = 0
   If grdDataList.Rows > 1 Then
      ShowBar grdDataList, intLastRow, 7
      cmdOK(0).Enabled = True
      cmdOK(0).Default = True
   Else
      MsgBox "¬dµL¸ê®Æ¡I"
      cmdOK(0).Enabled = False
      cmdOK(2).Default = True
   End If
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean, oText As TextBox
   For Each oText In txtCode
      txtCode_Validate oText.Index, Cancel
      If Cancel = True Then Exit Function
   Next
   TxtValidate = True
End Function

Private Sub Form_Activate()
   If txtCode(1) = "" Then
      txtCode(1) = "P"
      txtCode(2).SetFocus
   End If
   'Added by Sindy 2016/10/7
   If m_strIR01 <> "" And m_Done = False Then
      txtCode(1).Text = m_strCP01
      txtCode(2).Text = m_strCP02
      txtCode(3).Text = m_strCP03
      txtCode(4).Text = m_strCP04
      cmdOK(2).Value = True
      m_Done = True
      Me.Caption = Me.Caption & "¡]«H¥ó½s¸¹:" & m_strIR01 & "-" & m_strIR03 & "¡^"
   End If
   '2016/10/7 END
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
    'Add by Lydia 2014/10/31 ¶}©ñ¥~±Mµ{§Ç¤H­û(31,33,34)¥i¶i¤J±M§Q³B¨t²Î¾Þ§@FMP¾ÈµØ®×¥ó¡A¦ý«D¦¹Ãþ®×¥ó®É¥~±Mµ{§Ç¤H­û¤£¥i¾Þ§@¡C
    FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010513 = Nothing
End Sub

Private Sub grdDataList_DblClick()
   cmdok_Click 0
End Sub

Private Sub grdDataList_GotFocus()
   GridGotFocus grdDataList
End Sub

Private Sub grdDataList_LostFocus()
   GridLostFocus grdDataList
End Sub

Private Sub grdDataList_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then grdDataList_DblClick
End Sub

Private Sub grdDataList_RowColChange()
   If intLastRow <> grdDataList.row Then
      If blnOKtoShow Then
         blnOKtoShow = False
         ShowBar grdDataList, intLastRow, 7
         blnOKtoShow = True
      End If
   End If
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   TextInverse txtCode(Index)
   CloseIme
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
   
   If Index = 2 Or Index = 4 Then
      Cancel = Not GetPatent
   End If
End Sub

Private Sub SetDataListWidth()
Dim varGridWidth() As Variant

varGridWidth = Array(990, 765, 765, 1425, 2225, 1020, 765, 805)
SetGridDataListWidth grdDataList, varGridWidth()
blnOKtoShow = True
End Sub

Private Sub lblNation_Change()
   Dim strTemp As String
   If lblNation = "" Then Exit Sub
   If ClsPDGetNation(lblNation, strTemp) Then
      lblCountryName.Caption = strTemp
   End If
End Sub

Private Function GetPatent() As Boolean
   Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String, strNation As String
   If Len(txtCode(2)) <> txtCode(2).MaxLength Then
      MsgBox "¥»©Ò®×¸¹¿é¤J¿ù»~¡I"
   Else
      txtCode(3) = Right("0" & txtCode(3), 1)
      txtCode(4) = Right("00" & txtCode(4), 2)
      If ClsPDCheckCaseCodeIsExist(txtCode(1), txtCode(2), txtCode(3), txtCode(4), strCaseName1, strCaseName2, strCaseName3, , strNation, , , False) Then
         lblNation = strNation
         SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
         GetPatent = True
      Else
         MsgBox "¥»©Ò®×¸¹¿é¤J¿ù»~¡I"
      End If
   End If
End Function

