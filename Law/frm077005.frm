VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm077005 
   BorderStyle     =   1  '單線固定
   Caption         =   "智財訴訟案需專業部配合通知補收文作業"
   ClientHeight    =   4452
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7908
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4452
   ScaleWidth      =   7908
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   6900
      TabIndex        =   5
      Top             =   195
      Width           =   855
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "提供書狀意見"
      Height          =   315
      Index           =   1
      Left            =   2430
      TabIndex        =   7
      Top             =   1740
      Width           =   1515
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "配合開庭"
      Height          =   315
      Index           =   0
      Left            =   1110
      TabIndex        =   6
      Top             =   1740
      Width           =   1125
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "通知(&E)"
      Height          =   375
      Left            =   4350
      TabIndex        =   8
      Top             =   1710
      Width           =   855
   End
   Begin VB.CommandButton CmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   375
      Left            =   5970
      TabIndex        =   4
      Top             =   195
      Width           =   855
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   3
      Left            =   2970
      MaxLength       =   2
      TabIndex        =   3
      Top             =   315
      Width           =   495
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   2
      Left            =   2570
      MaxLength       =   1
      TabIndex        =   2
      Top             =   315
      Width           =   345
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   1
      Left            =   1660
      MaxLength       =   6
      TabIndex        =   1
      Top             =   315
      Width           =   855
   End
   Begin VB.TextBox txtCase 
      Height          =   300
      Index           =   0
      Left            =   1110
      MaxLength       =   3
      TabIndex        =   0
      Top             =   315
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2085
      Left            =   120
      TabIndex        =   11
      Top             =   2190
      Width           =   7665
      _ExtentX        =   13526
      _ExtentY        =   3683
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      AllowUserResizing=   3
      FormatString    =   "V|收文日|總收文號|備註主題(案件性質)|承辦人|智權人員"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
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
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1470
      X2              =   3060
      Y1              =   480
      Y2              =   480
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1110
      TabIndex        =   18
      Top             =   1035
      Width           =   6615
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "11668;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   300
      Index           =   2
      Left            =   1110
      TabIndex        =   17
      Top             =   1380
      Width           =   6435
      Size            =   "11351;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   300
      Index           =   1
      Left            =   2010
      TabIndex        =   16
      Top             =   675
      Width           =   5535
      Size            =   "9763;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   300
      Index           =   0
      Left            =   1110
      TabIndex        =   15
      Top             =   675
      Width           =   885
      Size            =   "1561;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "案件屬性："
      Height          =   225
      Index           =   4
      Left            =   90
      TabIndex        =   14
      Top             =   1380
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "當事人1："
      Height          =   225
      Index           =   3
      Left            =   90
      TabIndex        =   13
      Top             =   735
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   225
      Index           =   2
      Left            =   90
      TabIndex        =   12
      Top             =   1050
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "通知類型："
      Height          =   225
      Index           =   1
      Left            =   90
      TabIndex        =   10
      Top             =   1785
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   360
      Width           =   945
   End
End
Attribute VB_Name = "frm077005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/04/26 Form2.0已修改 lblFM2(index)、Combo1 ;  MSHFlexGrid1改字型=新細明體-ExtB
'Create by Lydia 2020/06/16 智財訴訟案需專業部配合通知補收文作業
Option Explicit
Dim intLastRow As Integer '記錄勾選最後一筆
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim strTmpQ As String
Dim intQ As Integer
Dim rsQuery As New ADODB.Recordset

Private Sub Chk1_Click(Index As Integer)
    If Index = 0 And Chk1(Index).Value = 1 Then
        Chk1(1).Value = 1 '配合開庭->自動+提供書狀意見
    End If
End Sub

Private Sub Cmd1_Click()
Dim intX As Integer
Dim strCP09 As String, strLOS15 As String

   If lblFM2(0).Caption = "" Then
       MsgBox "請輸入本所案號，並且查詢資料！", vbInformation
       Exit Sub
   End If
   If Chk1(0).Value = 0 And Chk1(1).Value = 0 Then
       MsgBox "請選擇通知類型！", vbInformation
       Exit Sub
   End If
   If InStr(lblFM2(2).Caption, "專利") = 0 And InStr(lblFM2(2).Caption, "商標") = 0 And InStr(lblFM2(2).Caption, "著作權") = 0 Then
       MsgBox "案件屬性應有專利、商標或著作權！", vbExclamation
       Exit Sub
   End If
   
   With MSHFlexGrid1
       For intX = 1 To .Rows - 1
          If .TextMatrix(intX, 0) = "v" And "" & .TextMatrix(intX, 2) <> "" Then
               strCP09 = "" & .TextMatrix(intX, 2)
               Exit For
          End If
       Next
   End With
   
   If strCP09 = "" Then
       MsgBox "請先選取一道收文！", vbInformation
   Else
       Cmd1.Enabled = False
       If FormSave(strCP09) = True Then
           PUB_SendMailCache
           Call doQuery(False)
       Else
           MsgBox "存檔失敗！", vbCritical
       End If
       Cmd1.Enabled = True
   End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdQuery_Click()
    Call doQuery(True)
End Sub

Private Sub doQuery(ByVal bolMsg As Boolean)
Dim m_Nation As String
    If Trim(txtCase(0).Text) = "" Or Len(Trim(txtCase(1).Text)) < 6 Then
        MsgBox "請輸入本所案號！", vbExclamation, "檢核資料"
        Exit Sub
    End If
    
    Call ClearForm(False)
    If txtCase(2) = "" Then txtCase(2) = "0"
    If txtCase(3) = "" Then txtCase(3) = "00"
     
    strTmpQ = "select lc01,lc02,lc03,lc04,lc05,lc06,lc07,lc11 as custno,nvl(cu04,nvl(cu05,cu06)) custname,lc47,lc15 " & _
                    "from lawcase,customer where lc01='" & txtCase(0) & "' and lc02='" & txtCase(1) & "' and lc03='" & txtCase(2) & "' and lc04='" & txtCase(3) & "' and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) "
    intQ = 0
    Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
    If intQ = 0 Then
        Exit Sub
    End If
    
    intQ = 0
    Combo1.AddItem "中：" & rsQuery.Fields("lc05"), 0
    If rsQuery.Fields("lc05") <> "" Then intQ = 1
    Combo1.AddItem "英：" & rsQuery.Fields("lc06"), 1
    If rsQuery.Fields("lc06") <> "" Then intQ = 2
    Combo1.AddItem "日：" & rsQuery.Fields("lc07"), 2
    If rsQuery.Fields("lc07") <> "" Then intQ = 3
    Combo1.ListIndex = intQ - 1
    
    lblFM2(0).Caption = "" & rsQuery.Fields("custno")
    lblFM2(1).Caption = "" & rsQuery.Fields("custname")
    lblFM2(2).Caption = "" & rsQuery.Fields("lc47")
    m_Nation = "" & rsQuery.Fields("lc15")
    
    Call SetGrd(True) '清空
    '抓未發文之A類案源收文1101~1104(民事委任律師)
    'Modified by Lydia 2020/06/19 改成以會計科目判斷
    'strTmpQ = "select ' ' V,sqldatet(cp05) cp05t,cp09, cp64||'('||" & IIf(m_Nation = "000", "cpm03", "cpm04") & "||')' as subject,st02 as cp14n,los04 as los04_1,los04 as los04_2,los15 " & _
                     "From caseprogress, casepropertymap, staff, LawOfficeSource " & _
                     "where cp01='" & txtCase(0) & "' and cp02='" & txtCase(1) & "' and cp03='" & txtCase(2) & "' and cp04='" & txtCase(3) & "' and cp10>='1101' and cp10<='1104' " & _
                     "and cp158=0 and cp159=0 and cp01=cpm01(+) and cp10=cpm02(+) and cp14=st01(+) and cp09=los06(+) and substr(los02,1,1) = 'A' "
    'Modified by Lydia 2025/07/04 只要控制未取消收文，拿掉and cp158=0
    strTmpQ = "select ' ' V,sqldatet(cp05) cp05t,cp09, cp64||'('||" & IIf(m_Nation = "000", "cpm03", "cpm04") & "||')' as subject,st02 as cp14n,los04 as los04_1,los04 as los04_2,los15 " & _
                     "From caseprogress, casepropertymap, staff, LawOfficeSource " & _
                     "where cp01='" & txtCase(0) & "' and cp02='" & txtCase(1) & "' and cp03='" & txtCase(2) & "' and cp04='" & txtCase(3) & "' and cpm11 in ('414101','416111','416121') " & _
                     "and cp159=0 and cp01=cpm01(+) and cp10=cpm02(+) and cp14=st01(+) and cp09=los06(+) and substr(los02,1,1) = 'A' "
    strTmpQ = strTmpQ & "order by cp05,cp04 "
    intQ = 1
    Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
    If intQ = 1 Then
         MSHFlexGrid1.FixedCols = 0
         Set MSHFlexGrid1.Recordset = rsQuery
         Call SetGrd
    Else
         If bolMsg = True Then MsgBox "查無A類案源資料！", vbInformation
    End If
End Sub

Private Sub Form_Load()

    MoveFormToCenter Me
    Call ClearForm(True)
    Call SetGrd(True)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm077005 = Nothing
End Sub

Private Sub ClearForm(ByVal bolResetCase As Boolean)
Dim oObj
    
    If bolResetCase = True Then
        For Each oObj In txtCase
            oObj.Text = ""
        Next
    End If
    
    For Each oObj In lblFM2
       oObj.Caption = ""
    Next

    Chk1(0).Value = 0
    Chk1(1).Value = 0
    Combo1.Clear
    
End Sub

Private Sub MSHFlexGrid1_Click()

   If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 1) <> "" Then
       GridClick MSHFlexGrid1, intLastRow, 0, 0
   End If
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   getGrdColRow MSHFlexGrid1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   MSHFlexGrid1.col = nCol
   MSHFlexGrid1.row = nRow
   If Me.MSHFlexGrid1.row < 1 And Me.MSHFlexGrid1.Text <> "V" Then
        If m_blnColOrderAsc = True Then
           Me.MSHFlexGrid1.Sort = 5 '字串昇冪
           m_blnColOrderAsc = False
        Else
           Me.MSHFlexGrid1.Sort = 6 '字串降冪
           m_blnColOrderAsc = True
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

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer, iR As Integer
Dim strTmp As String
 
   arrGridHeadText = Array("V", "收文日", "總收文號", "備註主題(案件性質)", "承辦人", "介紹人", "LOS04_2")
   arrGridHeadWidth = Array(260, 860, 1000, 3000, 900, 900, 0)
        
   MSHFlexGrid1.Visible = False
   MSHFlexGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
         MSHFlexGrid1.Clear
         MSHFlexGrid1.Rows = 2
   End If
       
   For iRow = 0 To MSHFlexGrid1.Cols - 1
      MSHFlexGrid1.row = 0
      MSHFlexGrid1.col = iRow
      MSHFlexGrid1.Text = arrGridHeadText(iRow)
      MSHFlexGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
   Next
   
   For iR = 1 To MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.row = iR
        For iRow = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.col = iRow
           '只顯示介紹人員第一人
           If iRow = 5 Then
              strTmp = "" & MSHFlexGrid1.TextMatrix(iR, iRow)
              If strTmp <> "" Then
                  If InStr(strTmp, ",") > 0 Then strTmp = Mid(strTmp, 1, InStr(strTmp, ",") - 1)
                  MSHFlexGrid1.TextMatrix(iR, iRow) = GetStaffName(strTmp, True)
                  MSHFlexGrid1.TextMatrix(iR, iRow + 1) = strTmp
              End If
           End If
           MSHFlexGrid1.CellAlignment = flexAlignLeftCenter ' 文字=>靠左  flexAlignCenterCenter
        Next iRow
   Next iR

   MSHFlexGrid1.Visible = True
End Sub

Private Function FormSave(ByVal pKeyNo As String) As Boolean
Dim strGrp As String
Dim strTitleNo As String, strTitleName As String
Dim tmpArr As Variant
Dim intP As Integer

   Chk1(0).Tag = "": Chk1(1).Tag = ""
   tmpArr = Split(lblFM2(2).Caption, ",")
   For intP = 0 To UBound(tmpArr)
      If Trim(tmpArr(intP)) <> "" Then
         Select Case tmpArr(intP)
              Case "專利"
                   strGrp = "B1P"
                   Chk1(0).Tag = "226"  '配合開庭
                   Chk1(1).Tag = "225" '提供書狀意見
              Case "商標", "著作權"
                   strGrp = "B1T"
                   Chk1(0).Tag = "213" '配合開庭
                   Chk1(1).Tag = "212" '提供書狀意見
         End Select
         If strGrp <> "" Then Exit For
      End If
   Next intP
   
   If strGrp <> "" Then
        For intP = 0 To 1
           If Chk1(intP).Value = 1 And Chk1(intP).Tag <> "" Then
               strTitleNo = strTitleNo & "," & Chk1(intP).Tag
               strTitleName = strTitleName & "、" & Chk1(intP).Caption
           End If
        Next intP
        strTitleNo = Mid(strTitleNo, 2)
        strTitleName = Mid(strTitleName, 2)
        
        strExc(0) = "select * from LawOfficeSource where LOS06=" & CNULL(pKeyNo)
        intP = 1
        Set RsTemp = ClsLawReadRstMsg(intP, strExc(0))
        If intP = 1 Then
            cnnConnection.BeginTrans
                '先EMAIL通知智權人員補收配合開庭(因為LOS01會被清空)
                Call PUB_AddMailCache_LOS("3", "" & RsTemp.Fields("LOS15"), strTitleName)
                '案件屬性有勾專利或商標或著作權時，若案件性質為訴訟案件性質[cpm11 in ('414101','416111','416121')]
                '抓未發文程序，列出案件性質(B1P/B1T-配合開庭)供勾選並回存LOS19以便預設在接洽單輸入畫面，勾選配合開庭時也自動勾選。
                strSql = "Update LawOfficeSource set LOS01=null, LOS02='B1',LOS19='" & strTitleNo & "' where LOS15='" & RsTemp.Fields("LOS15") & "' "
                cnnConnection.Execute strSql, intP
                If "" & RsTemp.Fields("LOS10") <> "" Then
                    '並重新計算TT-999999費用點數更新回去並改案件性質為736(服務費)；再EMAIL通知智權人員補收配合開庭226(B1)
                    strSql = "Update CaseProgress set CP10='736' where CP09='" & RsTemp.Fields("LOS10") & "' "
                    cnnConnection.Execute strSql, intP
                    
                    PUB_UpdateTTFee RsTemp.Fields("LOS15") 'Added by Morgan 2021/1/13 費用點數更新
                End If
            cnnConnection.CommitTrans
        End If
   End If
   
   FormSave = True

ErrHandle:
   If Err.Number <> 0 Then
        cnnConnection.RollbackTrans
   End If
   
End Function

