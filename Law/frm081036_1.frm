VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm081036_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "TIPS案請款階段分配比例-年度結算作業"
   ClientHeight    =   4920
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8424
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   8424
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "智財協作"
      Height          =   660
      Left            =   72
      TabIndex        =   25
      Top             =   2112
      Width           =   8316
      Begin VB.Label lblData 
         Caption         =   "lblData"
         Height          =   228
         Index           =   11
         Left            =   7560
         TabIndex        =   35
         Top             =   312
         Width           =   684
      End
      Begin VB.Label lblData 
         Caption         =   "lblData"
         Height          =   228
         Index           =   10
         Left            =   5784
         TabIndex        =   34
         Top             =   312
         Width           =   684
      End
      Begin VB.Label lblData 
         Caption         =   "lblData"
         Height          =   228
         Index           =   9
         Left            =   3216
         TabIndex        =   33
         Top             =   312
         Width           =   684
      End
      Begin VB.Label lblData 
         Caption         =   "lblData"
         Height          =   228
         Index           =   8
         Left            =   1536
         TabIndex        =   32
         Top             =   312
         Width           =   684
      End
      Begin VB.Label Label1 
         Caption         =   "專業點數%加總："
         Height          =   228
         Index           =   7
         Left            =   96
         TabIndex        =   29
         Top             =   312
         Width           =   1428
      End
      Begin VB.Label Label1 
         Caption         =   "點數加總："
         Height          =   228
         Index           =   10
         Left            =   2304
         TabIndex        =   28
         Top             =   312
         Width           =   948
      End
      Begin VB.Label Label1 
         Caption         =   "顧服獎金%加總："
         Height          =   228
         Index           =   12
         Left            =   4392
         TabIndex        =   27
         Top             =   312
         Width           =   1428
      End
      Begin VB.Label Label1 
         Caption         =   "點數加總："
         Height          =   228
         Index           =   14
         Left            =   6576
         TabIndex        =   26
         Top             =   312
         Width           =   996
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid1 
      Height          =   1908
      Left            =   48
      TabIndex        =   18
      Top             =   2904
      Width           =   8268
      _ExtentX        =   14584
      _ExtentY        =   3366
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CheckBox Check1 
      Caption         =   "包含主管已確認的年度"
      ForeColor       =   &H000000FF&
      Height          =   276
      Left            =   4512
      TabIndex        =   17
      Top             =   168
      Width           =   2148
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFC0&
      Caption         =   "主管確認"
      Height          =   324
      Left            =   6528
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   1248
      Width           =   1140
   End
   Begin VB.TextBox txtCase 
      Height          =   285
      Index           =   0
      Left            =   1116
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "ACS"
      Top             =   168
      Width           =   495
   End
   Begin VB.TextBox txtCase 
      Height          =   285
      Index           =   1
      Left            =   1668
      MaxLength       =   6
      TabIndex        =   1
      Top             =   168
      Width           =   855
   End
   Begin VB.TextBox txtCase 
      Height          =   285
      Index           =   2
      Left            =   2568
      MaxLength       =   1
      TabIndex        =   2
      Top             =   168
      Width           =   345
   End
   Begin VB.TextBox txtCase 
      Height          =   285
      Index           =   3
      Left            =   2976
      MaxLength       =   2
      TabIndex        =   3
      Top             =   168
      Width           =   495
   End
   Begin VB.CommandButton CmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   360
      Left            =   3528
      TabIndex        =   4
      Top             =   144
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&E)"
      Height          =   360
      Left            =   7488
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblData 
      Caption         =   "lblData"
      Height          =   228
      Index           =   14
      Left            =   5592
      TabIndex        =   37
      Top             =   1296
      Width           =   804
   End
   Begin VB.Label lblData 
      Caption         =   "lblData"
      Height          =   228
      Index           =   12
      Left            =   3216
      TabIndex        =   36
      Top             =   1296
      Width           =   804
   End
   Begin VB.Label lblData 
      Caption         =   "lblData"
      Height          =   228
      Index           =   7
      Left            =   1584
      TabIndex        =   31
      Top             =   1680
      Width           =   684
   End
   Begin VB.Label lblData 
      Caption         =   "lblData"
      Height          =   228
      Index           =   5
      Left            =   1104
      TabIndex        =   30
      Top             =   1296
      Width           =   588
   End
   Begin VB.Label lblSaleTot 
      Height          =   228
      Left            =   7632
      TabIndex        =   24
      Top             =   1680
      Width           =   684
   End
   Begin VB.Label lblRecTot 
      Height          =   228
      Left            =   5904
      TabIndex        =   23
      Top             =   1680
      Width           =   684
   End
   Begin VB.Label Label1 
      Caption         =   "顧服獎金："
      Height          =   228
      Index           =   9
      Left            =   6672
      TabIndex        =   22
      Top             =   1680
      Width           =   924
   End
   Begin VB.Label Label1 
      Caption         =   "專業點數："
      Height          =   228
      Index           =   6
      Left            =   4944
      TabIndex        =   21
      Top             =   1680
      Width           =   924
   End
   Begin VB.Label lblSalePoint 
      Caption         =   "lblSalePoint"
      ForeColor       =   &H00FF0000&
      Height          =   228
      Left            =   4032
      TabIndex        =   20
      Top             =   1680
      Width           =   684
   End
   Begin VB.Label Label1 
      Caption         =   "智權業績獎金比例："
      ForeColor       =   &H00FF0000&
      Height          =   228
      Index           =   4
      Left            =   2352
      TabIndex        =   19
      Top             =   1680
      Width           =   1668
   End
   Begin VB.Label Label1 
      Caption         =   "主管確認日期："
      Height          =   228
      Index           =   16
      Left            =   4272
      TabIndex        =   16
      Top             =   1296
      Width           =   1308
   End
   Begin VB.Label Label1 
      Caption         =   "年度入帳總點數："
      Height          =   228
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   1476
   End
   Begin VB.Label Label1 
      Caption         =   "結算通知日期："
      Height          =   228
      Index           =   1
      Left            =   1896
      TabIndex        =   14
      Top             =   1296
      Width           =   1308
   End
   Begin VB.Label Label1 
      Caption         =   "請款年度："
      Height          =   228
      Index           =   8
      Left            =   120
      TabIndex        =   13
      Top             =   1296
      Width           =   948
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   228
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   216
      Width           =   948
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   228
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   888
      Width           =   948
   End
   Begin VB.Label Label1 
      Caption         =   "當事人1："
      Height          =   228
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   568
      Width           =   888
   End
   Begin MSForms.Label lblFM2 
      Height          =   260
      Index           =   0
      Left            =   1140
      TabIndex        =   9
      Top             =   552
      Width           =   888
      Size            =   "1566;459"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   264
      Index           =   1
      Left            =   2040
      TabIndex        =   8
      Top             =   552
      Width           =   6252
      Size            =   "11028;466"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1140
      TabIndex        =   6
      Top             =   864
      Width           =   7164
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "12636;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1476
      X2              =   3066
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "frm081036_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2025/04/18
Option Explicit
Dim intLastRow As Integer '記錄MGrid1勾選最後一筆
Dim m_AT1(1 To 14) As String 'Index:6~11記錄全部系統別的總計
Dim strQuery As String, intQ As Integer
Dim rsQuery As New ADODB.Recordset
Dim oObj As Object
Dim colAT105 As Integer, colAT106 As Integer

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdok_Click()

   If Trim(txtCase(0).Text) = "" Or Len(Trim(txtCase(1).Text)) < 6 Then
      MsgBox "請輸入本所案號！", vbExclamation, "檢核資料"
      Exit Sub
   End If
   If m_AT1(1) & m_AT1(2) & m_AT1(3) & m_AT1(4) <> txtCase(0) & txtCase(1) & txtCase(2) & txtCase(3) Then
      MsgBox "輸入本所案號後，請執行查詢功能！", vbExclamation, "檢核資料"
      Exit Sub
   End If
   If "" & m_AT1(14) <> "" Then
      MsgBox "主管已確認，不可變更！", vbExclamation, "檢核資料"
      Exit Sub
   End If
   If Trim(lblData(5)) = "" Or Trim(lblData(12)) = "" Then
      MsgBox "請點選請款年度！", vbExclamation, "檢核資料"
      Exit Sub
   End If
   If ShowAT1detail(lblData(5), m_AT1(6)) = False Then
      MsgBox "目前無" & lblData(5) & "年度結算記錄，請重新查詢！", vbExclamation, "檢核資料"
      Exit Sub
   Else
      For Each oObj In lblData
         If m_AT1(oObj.Index) <> oObj.Caption Then
            If m_AT1(6) = "NONE" And InStr("08,09,10,11", Format(oObj.Index, "00")) > 0 Then
            Else
               MsgBox "年度結算有異動，請重新查詢！", vbExclamation, "檢核資料"
               Exit Sub
            End If
         End If
      Next
   End If
   strExc(0) = "select cp01||'-'||cp02||decode(cp03||cp04,'000',null,'-'||cp03||'-'||cp04) as caseno from acs_tips_rate,caseprogress " & _
               "where atr06=cp09(+) and nvl(atr09,0)+nvl(atr10,0)=0 and atr01='" & m_AT1(1) & "' and atr02='" & m_AT1(2) & "' and atr03='" & m_AT1(3) & "' and atr04='" & m_AT1(4) & "' and atr05='" & m_AT1(5) & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strExc(1) = RsTemp.GetString(adClipString, , , ",")
      If MsgBox("尚有下列智財協作案件未輸入分配比例，是否繼續主管確認？" & vbCrLf & Replace(strExc(1), ",", vbCrLf), vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   
   cmdOK.Enabled = False
   Screen.MousePointer = vbHourglass
   If SaveData = False Then
   Else
      QueryData2
   End If
   Screen.MousePointer = vbDefault
   PUB_SendMailCache
   
   Exit Sub
   
End Sub

Private Sub cmdQuery_Click()
   
   If Len(txtCase(1)) <> 6 Then
      MsgBox "請輸入本所案號！！", vbExclamation
      txtCase(1).SetFocus
      txtCase_GotFocus 1
      Exit Sub
   End If
   If Trim(txtCase(2)) = "" Then txtCase(2) = "0"
   If Trim(txtCase(3)) = "" Then txtCase(3) = "00"
   ClearForm False

   strQuery = "select lc01,lc02,lc03,lc04,lc05,lc06,lc07,lc11 as custno,nvl(cu04,nvl(cu05,cu06)) custname,cp10 " & _
             "from caseprogress,lawcase,customer " & _
             "where cp01='" & txtCase(0) & "' and cp02='" & txtCase(1) & "' and cp03='" & txtCase(2) & "' and cp04='" & txtCase(3) & "' " & _
             "and cp31='Y' and cp159=0 and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) " & _
             "and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) "
   intQ = 0
   Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
   If intQ = 0 Then
       Exit Sub
   Else
       If InStr(ACSforTIPSstep, "'" & rsQuery.Fields("cp10") & "'") = 0 Then
           MsgBox "查無TIPS進度！", vbInformation
           txtCase(1).SetFocus
           txtCase_GotFocus 1
           Exit Sub
       End If
       m_AT1(1) = "" & rsQuery.Fields("lc01")
       m_AT1(2) = "" & rsQuery.Fields("lc02")
       m_AT1(3) = "" & rsQuery.Fields("lc03")
       m_AT1(4) = "" & rsQuery.Fields("lc04")
       intQ = 0
       Combo1.AddItem "中：" & rsQuery.Fields("lc05"), 0
       If "" & rsQuery.Fields("lc05") <> "" And intQ = 0 Then intQ = 1
       Combo1.AddItem "英：" & rsQuery.Fields("lc06"), 1
       If "" & rsQuery.Fields("lc06") <> "" And intQ = 0 Then intQ = 2
       Combo1.AddItem "日：" & rsQuery.Fields("lc07"), 2
       If "" & rsQuery.Fields("lc07") <> "" And intQ = 0 Then intQ = 3
       Combo1.ListIndex = intQ - 1
       lblFM2(0).Caption = "" & rsQuery.Fields("custno")
       lblFM2(1).Caption = "" & rsQuery.Fields("custname")
       
        '(智權人員)智權業務獎金比例
       strExc(0) = "select nvl(atr08,0) * 0.01 as atr08 from acs_tips_rate where atr05='1' and atr01='" & m_AT1(1) & "' and atr02='" & m_AT1(2) & "' and atr03='" & m_AT1(3) & "' and atr04='" & m_AT1(4) & "' "
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       If intI = 1 Then
          lblSalePoint = Val("" & RsTemp.Fields("atr08")) * 100
       End If
       
       QueryData2
   End If
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   
   Frame1.BackColor = &H8000000F
   SetGrd1 True
   ClearForm True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsQuery = Nothing
   Set frm081036_1 = Nothing
End Sub

Private Sub ClearForm(ByVal bolAll As Boolean)

   If bolAll = True Then
      For Each oObj In txtCase
         If oObj.Index = 0 Then
            oObj.Text = "ACS"
         Else
            oObj.Text = ""
         End If
         oObj.Tag = oObj.Text
      Next
      For intI = 1 To 4
          m_AT1(intI) = ""
      Next intI
      Combo1.Clear
      For Each oObj In lblFM2
         oObj.Caption = ""
         oObj.Tag = ""
      Next
      intLastRow = 0
      lblSalePoint = ""  '(智權人員)智權業績獎金比例
   End If
   
   cmdOK.Visible = False
   For Each oObj In lblData
      oObj.Caption = ""
      oObj.Tag = ""
   Next
   lblRecTot = ""
   lblSaleTot = ""
   For intI = 5 To UBound(m_AT1)
      m_AT1(intI) = ""
   Next
   
End Sub

Private Sub QueryData2()

    ClearForm False
    SetGrd1 True
    
    If m_AT1(1) <> "" And m_AT1(2) <> "" Then
       strQuery = "SELECT '' V, AT105,AT106,NVL(AT107,0) AT107,NVL(AT108,0) AT108,NVL(AT109,0) AT109,NVL(AT110,0) AT110,NVL(AT111,0) AT111,SUBSTR(SQLDATET(AT112),1,10) AT112,SUBSTR(SQLDATET(AT114),1,10) AT114 " & _
                 "From ACS_TIPS_RATE1 WHERE AT101='" & m_AT1(1) & "' AND AT102='" & m_AT1(2) & "' AND AT103='" & m_AT1(3) & "' AND AT104='" & m_AT1(4) & "' "
       If Check1.Value = False Then
          strQuery = strQuery & "AND NVL(AT114,0)=0 "
       End If
       strQuery = strQuery & "ORDER BY AT105,AT106 "

       Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
       If intQ = 1 Then
          MGrid1.FixedCols = 0
          Set MGrid1.Recordset = rsQuery
          Call SetGrd1
          MGrid1.FixedCols = 4
       End If
    End If
End Sub

Private Sub txtCase_GotFocus(Index As Integer)
   TextInverse txtCase(Index)
End Sub

Private Sub txtCase_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCase_LostFocus(Index As Integer)
   If Index > 1 And Trim(txtCase(Index)) = "" Then
      If Index = 2 Then
           txtCase(2) = "0"
      ElseIf Index = 3 Then
           txtCase(3) = "00"
      End If
   End If
End Sub

Private Sub lblData_GotFocus(Index As Integer)
   TextInverse lblData(Index)
End Sub

Private Sub SetGrd1(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
 
   arrGridHeadText = Array("V", "請款年度", "協作系統別", "入帳總點數", "專業%加總", _
             "專業點數", "顧服%加總", "顧服獎金", "結算通知日期", "主管確認日期")
   arrGridHeadWidth = Array(260, 900, 1000, 1000, 1000, _
                   1000, 1000, 1000, 1200, 1200)
   
   MGrid1.Visible = False
   MGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
       MGrid1.Clear
       MGrid1.Rows = 2
   End If
       
   For iRow = 0 To MGrid1.Cols - 1
      MGrid1.row = 0
      MGrid1.col = iRow
      MGrid1.Text = arrGridHeadText(iRow)
      MGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      MGrid1.CellAlignment = flexAlignCenterCenter
   Next
   If colAT105 = 0 Then
      colAT105 = PUB_MGridGetId("請款年度", MGrid1)
      colAT106 = PUB_MGridGetId("協作系統別", MGrid1)
   End If
   
   For intI = 1 To MGrid1.Rows - 1
      MGrid1.row = intI
      For iRow = 1 To MGrid1.Cols - 1
         MGrid1.col = iRow
         '置中
         'If InStr("03,04,05,06,07", Format(iRow, "00")) > 0 Then
         If iRow >= 1 Or iRow <= 7 Then
            MGrid1.CellAlignment = flexAlignCenterCenter
         End If
      Next iRow
   Next intI
   
   MGrid1.Visible = True
End Sub

Private Sub MGrid1_Click()
Dim intRow As Integer, intCol As Integer
   
   With MGrid1
      If .MouseRow > 0 Then
         intRow = .MouseRow
         intCol = .MouseCol
         .row = intRow
         '----單選
         GridClick MGrid1, intLastRow, 0, 0, , "V"
         intLastRow = intRow
         .col = intCol
         
         ClearForm False
         
         If "" & .TextMatrix(intRow, 0) = "V" And "" & .TextMatrix(intRow, colAT105) <> "" And colAT105 > 0 Then
             If ShowAT1detail("" & .TextMatrix(intRow, colAT105), "" & .TextMatrix(intRow, colAT106)) = False Then
                 MsgBox "目前無" & .TextMatrix(intRow, colAT105) & "年度結算記錄，請重新查詢！"
             Else
                 For Each oObj In lblData
                    If m_AT1(6) = "NONE" And InStr("08,09,10,11", Format(oObj.Index, "00")) > 0 Then
                       oObj.Caption = ""
                    Else
                       oObj.Caption = m_AT1(oObj.Index)
                    End If
                    oObj.Tag = oObj.Caption
                 Next
                 If m_AT1(14) <> "" Then
                    cmdOK.Visible = False
                 Else
                    cmdOK.Visible = True
                 End If
                 lblRecTot = lblData(7)
                 lblSaleTot = Round(Val(lblData(7)) * (1 - (Val(lblSalePoint) / 100)), 2)
             End If
         End If
       End If
   End With
End Sub

Private Function ShowAT1detail(ByVal pAT105 As String, ByVal pAT106 As String) As Boolean
   
   ShowAT1detail = False
   
   For intI = 5 To 11
      m_AT1(intI) = ""
   Next
   If Val(pAT105) > 0 Then
      strExc(0) = "select * from ACS_TIPS_RATE1 WHERE AT101='" & m_AT1(1) & "' and AT102='" & m_AT1(2) & "' and AT103='" & m_AT1(3) & "' and AT104='" & m_AT1(4) & "' and AT105='" & pAT105 & "' " & _
                  "order by AT105, AT106 "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         RsTemp.MoveFirst
         m_AT1(5) = pAT105
         If pAT106 = "NONE" Then
             m_AT1(6) = "NONE"
         Else
             m_AT1(6) = "ALL"
         End If
         m_AT1(7) = Val("" & RsTemp.Fields("at107"))
         Do While Not RsTemp.EOF
            m_AT1(8) = Val(m_AT1(8)) + Val("" & RsTemp.Fields("at108"))
            m_AT1(9) = Val(m_AT1(9)) + Val("" & RsTemp.Fields("at109"))
            m_AT1(10) = Val(m_AT1(10)) + Val("" & RsTemp.Fields("at110"))
            m_AT1(11) = Val(m_AT1(11)) + Val("" & RsTemp.Fields("at111"))
            If Val(m_AT1(12)) < Val("" & RsTemp.Fields("at112")) Then
               m_AT1(12) = ChangeWStringToTDateString("" & RsTemp.Fields("at112"))
            End If
            If m_AT1(13) = "" And "" & RsTemp.Fields("at113") <> "" Then
               m_AT1(13) = "" & RsTemp.Fields("at113")
               m_AT1(14) = ChangeWStringToTDateString("" & RsTemp.Fields("at114"))
            End If
            RsTemp.MoveNext
         Loop
         ShowAT1detail = True
      End If
   End If
End Function

Private Function SaveData() As Boolean
Dim strCon As String
   
   SaveData = False
   strExc(0) = "select * from ACS_TIPS_RATE1 WHERE AT101='" & m_AT1(1) & "' AND AT102='" & m_AT1(2) & "' AND AT103='" & m_AT1(3) & "' AND AT104='" & m_AT1(4) & "' AND AT105='" & m_AT1(5) & "' " & _
               "AND AT106<>'NONE' order by AT105,AT106 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         strCon = strCon & vbCrLf & RsTemp.Fields("AT106") & "：專業點數" & Val("" & RsTemp.Fields("at109")) & "，專業業務獎金(顧服獎金)" & Val("" & RsTemp.Fields("at111"))
         RsTemp.MoveNext
      Loop
   Else
      strCon = strCon & vbCrLf & "NONE"
   End If
   '總專業點數=繳款總點數、總專業業務獎金(顧服獎金)=繳款總點數*(100-智權獎金%)
   strCon = m_AT1(5) & "年繳款總點數" & Val(lblData(7)) & "：專業點數" & Val(lblData(7)) & "，專業業務獎金(顧服獎金)" & Round(Val(lblData(7)) * (1 - (Val(lblSalePoint) / 100)), 2) & vbCrLf & _
            "智財協作點數請依以下進行調整：" & strCon
   
   If strCon <> "" Then
      cnnConnection.BeginTrans
         strSql = "Update ACS_TIPS_RATE1 SET AT113='" & strUserNum & "', AT114=TO_CHAR(SYSDATE,'YYYYMMDD') WHERE AT101='" & m_AT1(1) & "' AND AT102='" & m_AT1(2) & "' AND AT103='" & m_AT1(3) & "' AND AT104='" & m_AT1(4) & "' AND AT105='" & m_AT1(5) & "' "
         cnnConnection.Execute strSql
         strSql = "Update ACS_TIPS_RATE SET ATR12=TO_CHAR(SYSDATE,'YYYYMMDD') WHERE atr01='" & m_AT1(1) & "' AND atr02='" & m_AT1(2) & "' AND atr03='" & m_AT1(3) & "' AND atr04='" & m_AT1(4) & "' AND atr05='" & m_AT1(5) & "' "
         cnnConnection.Execute strSql
         strExc(1) = Pub_GetSpecMan("財務處總帳人員")
         strExc(2) = Pub_GetSpecMan("ACS郵件通知主管")
         If strExc(1) <> "" Then
             strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09) " & _
                     "values('" & strUserNum & "','" & strExc(1) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss') " & _
                     ",'" & m_AT1(1) & "-" & m_AT1(2) & IIf(m_AT1(3) & m_AT1(4) = "000", "", "-" & m_AT1(3) & "-" & m_AT1(4)) & "(" & m_AT1(5) & "年)最後一階段款項已入帳，請財務處進行點數分配', " & _
                     "'" & ChgSQL(strCon) & "','" & strExc(2) & "') "
             cnnConnection.Execute strSql
         End If
      cnnConnection.CommitTrans
   End If
   SaveData = True
   Exit Function
   
ErrHandle:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox "存檔失敗：" & vbCrLf & Err.Description
   End If
End Function
