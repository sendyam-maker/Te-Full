VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060406 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專程序發文逾期統計表"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5070
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   4
      Left            =   1440
      TabIndex        =   5
      Top             =   2340
      Width           =   555
      VariousPropertyBits=   679495707
      MaxLength       =   1
      Size            =   "979;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   6
      Left            =   210
      TabIndex        =   15
      Top             =   2400
      Width           =   4515
      VariousPropertyBits=   8388627
      Caption         =   "報表種類：　　　(1-統計, 2-明細)"
      Size            =   "7964;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   5
      Left            =   210
      TabIndex        =   14
      Top             =   3540
      Width           =   615
      VariousPropertyBits=   8388627
      Caption         =   "備註："
      Size            =   "1085;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "1. 本統計表/明細表僅統計發文日超過承辦期限的案件。"
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   13
      Top             =   3810
      Width           =   4365
   End
   Begin VB.Label Label1 
      Caption         =   "2. 統計方式為承辦期限減去發文日所得的工作天數作為數值，負值即代表遲延天數。另後附括號以表示遲延的總件數。"
      Height          =   435
      Index           =   1
      Left            =   210
      TabIndex        =   12
      Top             =   4020
      Width           =   4800
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2580
      X2              =   2790
      Y1              =   480
      Y2              =   480
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   4
      Left            =   1470
      TabIndex        =   11
      Top             =   1560
      Width           =   3165
      ForeColor       =   16711680
      VariousPropertyBits=   8388627
      Caption         =   "(用,區隔性質，空白：表示全部)"
      Size            =   "5583;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.CommandButton CmdFM2 
      Height          =   450
      Left            =   930
      TabIndex        =   6
      Top             =   2910
      Width           =   2985
      BackColor       =   12648384
      Caption         =   "列印(&P)"
      Size            =   "5265;794"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CheckBox ChkFM2_1 
      Height          =   345
      Left            =   1470
      TabIndex        =   4
      Top             =   1860
      Width           =   2445
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "4313;609"
      Value           =   "0"
      Caption         =   "限程序大項工作的性質"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   3
      Left            =   1440
      TabIndex        =   3
      Top             =   1140
      Width           =   2985
      VariousPropertyBits=   679495707
      Size            =   "5265;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   3
      Left            =   210
      TabIndex        =   10
      Top             =   1230
      Width           =   1245
      VariousPropertyBits=   8388627
      Caption         =   "案件性質："
      Size            =   "2196;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   2
      Left            =   2640
      TabIndex        =   9
      Top             =   765
      Width           =   1065
      VariousPropertyBits=   8388627
      Caption         =   "Name"
      Size            =   "1879;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   690
      Width           =   1065
      VariousPropertyBits=   679495707
      MaxLength       =   6
      Size            =   "1879;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   1
      Left            =   210
      TabIndex        =   8
      Top             =   765
      Width           =   1065
      VariousPropertyBits=   8388627
      Caption         =   "承辦人："
      Size            =   "1879;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   1
      Left            =   2880
      TabIndex        =   1
      Top             =   270
      Width           =   1065
      VariousPropertyBits=   679495707
      MaxLength       =   7
      Size            =   "1879;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   270
      Width           =   1065
      VariousPropertyBits=   679495707
      MaxLength       =   7
      Size            =   "1879;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   7
      Top             =   345
      Width           =   1275
      VariousPropertyBits=   8388627
      Caption         =   "發文日期："
      Size            =   "2249;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
End
Attribute VB_Name = "frm060406"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2019/08/19 外專程序發文逾期統計表
'Memo by Lydia 2019/08/19 使用Form 2.0 (Label,TextBox,CommandBox,CheckBox)
Option Explicit

Dim oText As Control
Dim PLeft(0 To 20) As Integer, iPrint As Long, m_iMaxNum As Integer, m_iColMax As Integer
Dim strPrinter As String
Dim m_PrtOrientation As Integer
Dim m_PaperSize As Integer
Dim rsAD As New ADODB.Recordset
Dim m_stGrp As String, m_stTotal As String, m_stCount As String
Dim arrNameList() As String '個人資料
Dim m_iRound As Integer '橫軸頁數
Dim m_iPages As Integer '列印頁數

Private Sub GetPleft1()
   Dim ii As Integer
      
   Erase PLeft
   
   m_iMaxNum = 10 '每頁最大人數
   m_iColMax = 3 '名稱每列字數
   
   PLeft(0) = 500
   PLeft(1) = 500
   PLeft(2) = PLeft(1) + 3120
   For ii = 0 To m_iMaxNum
      PLeft(3 + ii) = PLeft(2 + ii) + 1200
   Next
End Sub

Private Sub GetPleft3()
   
   Erase PLeft
      
   PLeft(0) = 500
   PLeft(1) = 500 '承辦人
   PLeft(2) = PLeft(1) + 1680 '本所案號
   PLeft(3) = PLeft(2) + 2160 '發文日
   PLeft(4) = PLeft(3) + 1200 '承辦期限
   PLeft(5) = PLeft(4) + 1200 '天數差
   PLeft(6) = PLeft(5) + 960 '案件性質
   PLeft(7) = PLeft(6) + 3120
End Sub

Private Sub ChkFM2_1_Click()
    If ChkFM2_1.Value = True Then
        txtFM2(3).Enabled = False
        txtFM2(3).Text = ""
    Else
        txtFM2(3).Enabled = True
    End If
End Sub

Private Sub CmdFM2_Click()
Dim bolTmp As Boolean
    
    For Each oText In txtFM2
         If (oText.Index = 0 Or oText.Index = 1) And Trim(oText.Text) = "" Then
              MsgBox "發文日期區間不可空白！", vbCritical, "檢核條件"
              txtFM2_GotFocus oText.Index
              txtFM2(oText.Index).SetFocus
              Exit Sub
         End If
         
         Call txtFM2_Validate(oText.Index, bolTmp)
         If bolTmp = True Then
              txtFM2_GotFocus oText.Index
              txtFM2(oText.Index).SetFocus
              Exit Sub
         End If
    Next
    
    If txtFM2(0) > txtFM2(1) Then
        MsgBox "發文日期起值不可大於迄值！", vbCritical, "檢核條件"
        txtFM2_GotFocus 0
        txtFM2(oText.Index0).SetFocus
        Exit Sub
    End If
   
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call doQuery
    Me.Enabled = True
    Screen.MousePointer = vbDefault
      
End Sub

Private Sub Form_Load()

    MoveFormToCenter Me
    
    ClearForm
    
    ChkFM2_1.Value = True '預設為程序大項工作
    
    strPrinter = Printer.DeviceName
    m_PrtOrientation = Printer.Orientation
    m_PaperSize = Printer.PaperSize
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060406 = Nothing
End Sub

Private Sub ClearForm()

    For Each oText In txtFM2
         oText.Text = ""
    Next
    LblFM2(2).Caption = ""
End Sub

Private Sub doQuery()
Dim stPty As String
Dim stCon As String
Dim ii As Integer, iFrom As Integer, iTo As Integer
   
   '發文日期
   If txtFM2(0).Text <> "" Then
       stCon = stCon & " AND NVL(CP27,0)>=" & DBDATE(txtFM2(0))
   End If
   If txtFM2(1).Text <> "" Then
       stCon = stCon & " AND NVL(CP27,0)<=" & DBDATE(txtFM2(1))
   End If
   '承辦人
   If txtFM2(2).Text <> "" Then
       stCon = stCon & " AND CP14=" & CNULL(txtFM2(2))
   End If
   '案件性質
   If ChkFM2_1.Value = True Then
       '程序大項工作：1.告准函1917、2.專利證書1603、3.公開公報1229、4.專利權消滅1604、5.通知年費逾期1605、6.期限通知-年費1913-605
       stCon = stCon & " AND CP10 IN ('1917','1603','1229','1604','1605','1913') "
   ElseIf txtFM2(3).Text <> "" Then
       stCon = stCon & " AND CP10 IN (" & GetAddStr(Trim(txtFM2(3))) & ") "
   End If
   strExc(0) = "SELECT CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) CASENO,CP09," & _
                "CP14,ST02,CP10,DECODE(NVL(PA09,'000'),'000',CPM03,CPM04) C1,WORKDAYDIFF(CP27,CP48) C2,substr(sqldatet(CP27),1,9) as CP27,substr(sqldatet(CP48),1,9) as CP48 " & _
               "FROM CASEPROGRESS A, STAFF,PATENT,CASEPROPERTYMAP " & _
               "WHERE CP48+0>0  AND CP01='FCP' AND SUBSTR(CP09,1,1)<>'D' " & stCon & _
               "AND CP14 IS NOT NULL AND ST01(+)=CP14 AND ST03='F22' " & _
               "AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA01 IS NOT NULL " & _
               "AND CP01=CPM01(+) AND CP10=CPM02(+) "
   If txtFM2(4) = "2" Then
       '明細
       strSql = "SELECT CP14 AS R01,ST02 AS R02,CASENO AS R03,CP27 AS R04,CP48 AS R05, C2 AS R06,C1 AS R07,CP10 FROM ( " & _
                   strExc(0) & ")  WHERE C2 < 0 ORDER BY CP14,CP10,CP27, CASENO "
   Else
       '統計
       strSql = "SELECT CP14 AS R01,ST02 AS R02,CP10 AS R03,SUBSTR(C1,1,20) R04,SUM(C2) R05,SUM(1) R06 FROM ( " & _
                    strExc(0) & ") WHERE C2<0 GROUP BY CP14,ST02,CP10,SUBSTR(C1,1,20) ORDER BY CP10,CP14"
   End If

   intI = 1
   Set rsAD = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
        '印表機設定A4
        Printer.PaperSize = vbPRPSA4
        If txtFM2(4) = "2" Then
           '明細-直印
           Printer.Orientation = 1
           GetPleft3
        Else
           '統計-橫印
           Printer.Orientation = 2
           GetPleft1
        End If

        m_iPages = 0
        m_stTotal = ""
        m_stCount = ""
        stPty = ""
        With rsAD
            SetNameList
            .MoveFirst
            m_iRound = 1
            PrintTitle
            
            If txtFM2(4) = "2" Then '明細
               Do While Not .EOF
                  NewLine , True
                  Printer.CurrentX = PLeft(1)
                  Printer.CurrentY = iPrint
                  Printer.Print "" & .Fields("R02") '承辦人
                  Printer.CurrentX = PLeft(2)
                  Printer.CurrentY = iPrint
                  Printer.Print "" & .Fields("R03") '本所案號
                  Printer.CurrentX = PLeft(3)
                  Printer.CurrentY = iPrint
                  Printer.Print "" & .Fields("R04") '發文日
                  Printer.CurrentX = PLeft(4)
                  Printer.CurrentY = iPrint
                  Printer.Print "" & .Fields("R05") '承辦期限
                  strExc(1) = "" & .Fields("R06") '天數差
                  Printer.CurrentX = PLeft(6) - 240 - Printer.TextWidth(strExc(1))
                  Printer.CurrentY = iPrint
                  Printer.Print strExc(1)
                  Printer.CurrentX = PLeft(6)
                  Printer.CurrentY = iPrint
                  Printer.Print "" & .Fields("R07") '案件性質
                  m_stCount = Val(m_stCount) + 1
                  .MoveNext
               Loop
               PrintCount
               
            Else '統計
'-----------------------------------
                Do While Not .EOF
                   '列印-案件性質
                   If stPty <> "" & .Fields("R04") Then
                      NewLine , True
                      stPty = "" & .Fields("R04")
                      Printer.CurrentX = PLeft(1)
                      Printer.CurrentY = iPrint
                      Printer.Print StrToStr(stPty, 12)
                   End If
                   '比對-員工編號符合,才列出
                   iFrom = (m_iRound - 1) * m_iMaxNum + 1
                   If UBound(arrNameList, 2) > iFrom + m_iMaxNum - 1 Then
                      iTo = iFrom + m_iMaxNum - 1
                   Else
                      iTo = UBound(arrNameList, 2)
                   End If
                   For ii = iFrom To iTo
                      If arrNameList(1, ii) = .Fields("R01") Then
                         If Not IsNull(.Fields("R05")) Then
                            strExc(1) = .Fields("R05") & Format("(" & .Fields("R06") & ")", "@@@@")
                            Printer.CurrentX = PLeft(IIf(ii Mod m_iMaxNum = 0, m_iMaxNum, ii Mod m_iMaxNum) + 2) - 240 - Printer.TextWidth(strExc(1))
                            Printer.CurrentY = iPrint
                            Printer.Print strExc(1)
                            arrNameList(3, ii) = Val(arrNameList(3, ii)) + Val("" & .Fields("R05"))
                            arrNameList(4, ii) = Val(arrNameList(4, ii)) + Val("" & .Fields("R06"))
                         End If
                         Exit For
                      End If
                   Next
                   .MoveNext
                   If .EOF Then '尚未列印所有資料,因為欄位數超過1頁
                      PrintSubTot
                      '欄位數超過
                      If UBound(arrNameList, 2) > m_iRound * m_iMaxNum Then
                         m_iRound = m_iRound + 1
                         PrintTitle
                         .MoveFirst
                      End If
                   End If
                Loop
                PrintTotal
'-----------------------------------
            End If
        End With
        Printer.EndDoc
        ShowPrintOk
        '還原設定
        Printer.PaperSize = m_PaperSize
        Printer.Orientation = m_PrtOrientation

   Else
      MsgBox "無符合資料!!"
   End If
   Set rsAD = Nothing
End Sub

Private Sub SetNameList()
   Dim ii As Integer, stLstCP14 As String
   Set RsTemp = rsAD.Clone
   RsTemp.Sort = "R01,R03"
   With RsTemp
        .MoveFirst
        Erase arrNameList
        ii = 0
        stLstCP14 = ""
        Do While Not .EOF
           If stLstCP14 <> "" & .Fields("R01") Then
              ii = ii + 1
              ReDim Preserve arrNameList(4, ii) As String
              arrNameList(1, ii) = "" & .Fields("R01")
              arrNameList(2, ii) = "" & .Fields("R02")
              arrNameList(3, ii) = 0     '小計(天數)
              arrNameList(4, ii) = 0     '小計(天數)
              stLstCP14 = "" & .Fields("R01")
           End If
           .MoveNext
        Loop
   End With
End Sub

Private Sub txtFM2_GotFocus(Index As Integer)
    TextInverse txtFM2(Index)
End Sub

Private Sub txtFM2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
     If Index = 2 Or Index = 3 Then
         KeyAscii = UpperCase(KeyAscii)
     Else
         KeyAscii = Pub_NumAscii(KeyAscii)
     End If
End Sub

Private Sub txtFM2_Validate(Index As Integer, Cancel As Boolean)
Dim strText As String
    
    Select Case Index
         Case 0, 1 '發文日期
            If PUB_CheckKeyInDate(txtFM2(Index)) = -1 Then
               GoTo EXITSUB
            End If
         Case 2 '承辦人
            LblFM2(2).Caption = ""
            If txtFM2(Index).Text <> "" Then
                If ClsPDGetStaff(txtFM2(Index), strText) = True Then
                    LblFM2(2).Caption = strText
                Else
                    GoTo EXITSUB
                End If
            End If
         Case 4 '報表種類
            If txtFM2(Index) <> "1" And txtFM2(Index) <> "2" Then
                 MsgBox "請輸入報表種類1或2 ！", vbCritical, "檢核條件"
                 GoTo EXITSUB
            End If
    End Select
    
    Exit Sub
    
EXITSUB:
    txtFM2(Index).SetFocus
    txtFM2_GotFocus Index
    Cancel = True
    
End Sub

Private Sub NewLine(Optional iHeight As Integer = 400, Optional bDrawLine As Boolean)
   iPrint = iPrint + iHeight
   If iPrint > Printer.ScaleHeight - 800 Then
      If bDrawLine Then
         iPrint = iPrint - iHeight
         PrintLine
      End If
      PrintTitle
      iPrint = iPrint + 300
   End If
End Sub

Private Sub PrintLine(Optional iType As Integer = 0)
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   If iType = 1 Then
      Printer.Line (PLeft(1), iPrint + 150)-(Printer.ScaleWidth - 200, iPrint + 150)
   Else
      Printer.Print String(Round((Printer.ScaleWidth - PLeft(1) - 200) / Printer.TextWidth("-")), "-")
   End If
End Sub

Private Sub PrintTitle()
   Dim stCon As String
   Dim stTmp As String
   
   m_iPages = m_iPages + 1
      
   If m_iPages > 1 Then Printer.NewPage
   
   stTmp = Me.Caption
   If txtFM2(4) = "2" Then stTmp = Replace(stTmp, "統計", "明細")
   
   iPrint = 500
   Printer.FontName = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(stTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print stTmp
      
   
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.Font.Size = 12
      
   iPrint = iPrint + 500
   stTmp = "發文日期：" & Format(ChangeTStringToTDateString(txtFM2(0)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txtFM2(1))
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(stTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print stTmp

   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   
   Printer.CurrentX = Printer.Width - 3400
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   iPrint = iPrint + 300
   
   If ChkFM2_1.Value = True Then
       stTmp = "限程序大項工作"
   Else
       stTmp = IIf(Trim(txtFM2(3)) = "", "ALL", Trim(txtFM2(3)))
   End If
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "案件性質：" & stTmp
   
   If Trim(txtFM2(2)) <> "" Then
        stTmp = "承辦人：" & LblFM2(2).Caption
        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(stTmp)) / 2
        Printer.CurrentY = iPrint
        Printer.Print stTmp
   End If
   
   Printer.CurrentX = Printer.Width - 3400
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(m_iPages)
   PrintLine 1

   If txtFM2(4) = "2" Then
      PrintTitle2 '明細
   Else
      PrintTitle1 '統計
   End If
End Sub

'列印欄位名稱(統計)
Private Sub PrintTitle1()
   Dim ii As Integer, jj As Integer
   Dim iY1 As Long, iY2 As Long
   
   iPrint = iPrint + 300
   iY1 = iPrint
   iY2 = iPrint
      
   For ii = 1 To m_iMaxNum
      jj = m_iMaxNum * (m_iRound - 1) + ii
      If UBound(arrNameList, 2) >= jj Then
         strExc(0) = StrToStr(arrNameList(2, jj), Val(m_iColMax))
         If strExc(0) <> arrNameList(2, jj) Then
            iY2 = iY1 + 300
            Exit For
         End If
      End If
   Next
   
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iY2
   Printer.Print "案件性質"
   
   For ii = 1 To m_iMaxNum
      jj = m_iMaxNum * (m_iRound - 1) + ii
      If UBound(arrNameList, 2) >= jj Then
         Printer.CurrentX = PLeft(ii + 1)
         
         strExc(0) = StrToStr(arrNameList(2, jj), Val(m_iColMax))
         If strExc(0) = arrNameList(2, jj) Then
            Printer.CurrentY = iY2
            Printer.Print arrNameList(2, jj)
         Else
            Printer.CurrentY = iY1
            Printer.Print strExc(0)
            Printer.CurrentX = PLeft(ii + 1)
            Printer.CurrentY = iY2
            Printer.Print StrToStr(Mid(arrNameList(2, jj), Len(strExc(0)) + 1), Val(m_iColMax))
         End If
      Else
         Exit For
      End If
   Next
   iPrint = iY2 - 100
   PrintLine 1
End Sub

'列印欄位名稱(明細)
Sub PrintTitle2()
  
   iPrint = iPrint + 300
   
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "承辦人"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "發文日"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "承辦期限"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "天數差"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "案件性質"
   
   iPrint = iPrint - 100
   PrintLine 1
End Sub

Private Sub PrintSubTot()
   Dim ii As Integer, iFrom As Integer, iTo As Integer
   
   PrintLine 1
   NewLine
   
   Printer.CurrentX = PLeft(2) - 480 - Printer.TextWidth("合計")
   Printer.CurrentY = iPrint
   Printer.Print "合計"
   
   iFrom = (m_iRound - 1) * m_iMaxNum + 1
   If UBound(arrNameList, 2) > iFrom + m_iMaxNum - 1 Then
      iTo = iFrom + m_iMaxNum - 1
   Else
      iTo = UBound(arrNameList, 2)
   End If
   For ii = iFrom To iTo
      strExc(1) = arrNameList(3, ii) & Format("(" & arrNameList(4, ii) & ")", "@@@@")
      Printer.CurrentX = PLeft(IIf(ii Mod m_iMaxNum = 0, m_iMaxNum, ii Mod m_iMaxNum) + 2) - 240 - Printer.TextWidth(strExc(1))
      Printer.CurrentY = iPrint
      Printer.Print strExc(1)
      m_stTotal = Val(m_stTotal) + Val(arrNameList(3, ii))
      m_stCount = Val(m_stCount) + Val(arrNameList(4, ii))
   Next
End Sub

Private Sub PrintTotal()
   NewLine
   Printer.CurrentX = PLeft(2) - 480 - Printer.TextWidth("總計")
   Printer.CurrentY = iPrint
   Printer.Print "總計"
   
   strExc(1) = m_stTotal & Format("(" & m_stCount & ")", "@@@@")
   Printer.CurrentX = PLeft(3) - 240 - Printer.TextWidth(strExc(1))
   Printer.CurrentY = iPrint
   Printer.Print strExc(1)
End Sub

Private Sub PrintCount()
   PrintLine 1
   NewLine
   
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "共 " & m_stCount & " 筆資料"
End Sub

