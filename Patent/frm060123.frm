VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060123 
   BorderStyle     =   1  '單線固定
   Caption         =   "電子送件稽核"
   ClientHeight    =   3996
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8904
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3996
   ScaleWidth      =   8904
   Begin VB.CommandButton cmdPath 
      Height          =   315
      Left            =   7170
      Picture         =   "frm060123.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   3375
      Width           =   330
   End
   Begin VB.CommandButton cmdBatch 
      Caption         =   "電子稽核(&A)"
      Height          =   315
      Left            =   7500
      TabIndex        =   11
      Top             =   3375
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2415
      Left            =   120
      TabIndex        =   10
      Top             =   870
      Width           =   8655
      _ExtentX        =   15261
      _ExtentY        =   4255
      _Version        =   393216
      Cols            =   12
      AllowUserResizing=   3
      FormatString    =   "V|發文日|本所案號|案件名稱|案件性質|申請人|承辦人|智權人員|下一程序|本所期限|法定期限|申請案號"
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
      _Band(0).Cols   =   12
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "稽核完成(&S)"
      Height          =   375
      Index           =   3
      Left            =   6540
      TabIndex        =   6
      Top             =   120
      Width           =   1160
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "下一程序(&N)"
      Height          =   375
      Index           =   2
      Left            =   5220
      TabIndex        =   5
      Top             =   450
      Width           =   1160
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "案件進度(&D)"
      Height          =   375
      Index           =   1
      Left            =   4020
      TabIndex        =   4
      Top             =   450
      Width           =   1160
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   7740
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "查詢"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3180
      TabIndex        =   3
      Top             =   450
      Width           =   615
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   1
      Left            =   2220
      MaxLength       =   7
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Index           =   0
      Left            =   1140
      MaxLength       =   7
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1140
      TabIndex        =   0
      Top             =   120
      Width           =   1965
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3466;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCount 
      Alignment       =   1  '靠右對齊
      Height          =   255
      Left            =   7380
      TabIndex        =   16
      Top             =   660
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "PS: 電子稽核之發文日抓管制日期迄日"
      Height          =   195
      Index           =   3
      Left            =   5550
      TabIndex        =   15
      Top             =   3780
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "送件/繳費清單CSV檔資料夾："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   14
      Top             =   3435
      Width           =   2355
   End
   Begin MSForms.TextBox txt2Path 
      Height          =   285
      Left            =   2550
      TabIndex        =   13
      Top             =   3360
      Width           =   4605
      VariousPropertyBits=   679495711
      Size            =   "8123;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   2140
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Caption         =   "管制日期："
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   9
      Top             =   510
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "發文人員："
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "frm060123"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/19 Form2.0已修改
'Create by Lydia 2019/03/20 電子送件稽核(外專新版)
Option Explicit

Public cmdState As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim colCase As Integer '本所案號-欄位
Dim colCp09 As Integer '收文號-欄位


Private Sub Cmd1_Click(Index As Integer)
   If Index = 0 Then '查詢
      If txtField(0) <> "" And txtField(1) <> "" And txtField(0) > txtField(1) Then
           MsgBox "管制日期起值不可大於迄值 !"
           txtField(0).SetFocus
           Call txtField_GotFocus(0)
           Exit Sub
      End If
      If doQuery() = False Then
      End If
   Else
      cmdState = Index
      PubShowNextData
   End If
End Sub

Private Sub cmdBatch_Click()
   Dim bolReQuery As Boolean
   
   'Added by Lydia 2020/03/09 先檢查表單
   If PUB_CheckFormExist("frm100106_1") Then
       MsgBox "請先關閉共同查詢〔以期限管制日查詢〕畫面！"
       Exit Sub
   End If
   If PUB_CheckFormExist("frm100106_2") Then
       MsgBox "請先關閉共同查詢〔以期限管制日查詢〕畫面！"
       Exit Sub
   End If
   'end 2020/03/09
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   If txt2Path = "" Then
      MsgBox "請先設定CSV資料夾!", vbExclamation
   ElseIf Dir(txt2Path, vbDirectory) = "" Then
      MsgBox "CSV資料夾讀取失敗!", vbExclamation
   Else
      doBatch bolReQuery
   End If
   Me.Enabled = True
   Screen.MousePointer = vbDefault

   'Modified by Lydia 2020/03/09 外專為避免漏執行查詢法定期限未發文案件，將每日電子送件稽核與查詢法定期限結合一起。
                                                   '人工稽核(稽核完成Cmd(3)) 則不用執行共同查詢
   'If bolReQuery Then Cmd1(0).Value = True '有稽核完成時自動重新查詢
   If bolReQuery Then
        If CheckUse("frm100106_1", strExec) Then
            frm100106_1.Show
            frm100106_1.cmdState = 0
            frm100106_1.opt1(1).Value = True
            frm100106_1.txt2(0) = Me.txtField(0)
            'Modify By Sindy 2023/4/13 淑華說要多加2個工作天(不含當天)
            'frm100106_1.txt2(1) = Me.txtField(1)
            frm100106_1.txt2(1) = TransDate(CompWorkDay(3, strSrvDate(1), 0), 1)
            '2023/4/13 END
            frm100106_1.Check1.Value = 1
            frm100106_1.PubShowNextData
        End If
        Cmd1(0).Value = True
   End If
   'end 2020/03/09
   
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPath_Click()
   Dim fName As String, strStartFolder As String
   
   If Dir(txt2Path & "\", vbDirectory) <> "" Then strStartFolder = txt2Path
   
   fName = PUB_GetFolder(Me.hWnd, strStartFolder, "請選取資料夾:")
   If fName <> "" Then 'they did not hit cancel
      txt2Path = fName
      SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", txt2Path
   End If
End Sub

Private Sub Form_Load()

    MoveFormToCenter Me
    
    txtField(0) = TransDate(CompWorkDay(5, strSrvDate(1), 1), 1)
    txtField(1) = strSrvDate(2)
    
    'Added by Morgan 2019/7/19
    txt2Path.Text = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
    'Added by Lydia 2024/07/22
    If InStr(UCase(txt2Path), "\\TYPING2\") > 0 Then
      txt2Path = Replace(txt2Path, "\\Typing2\", "\\" & strTyping2Path & "\")
      txt2Path = Replace(txt2Path, "\\TYPING2\", "\\" & strTyping2Path & "\")
      txt2Path = Replace(txt2Path, "\\typing2\", "\\" & strTyping2Path & "\")
    End If
    'end 2024/07/22
    
    If txt2Path = "" Then
      'Modified by Lydia 2024/07/22 改用變數
      'txt2Path = "\\Typing2\fcp程序\電子送件稽核"
      txt2Path = "\\" & strTyping2Path & "\fcp程序\電子送件稽核"
    End If
    'end 2019/7/19
    
    SetGrd True
    Call SetCombo1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060123 = Nothing
End Sub

Public Function doQuery() As Boolean
Dim strCon As String
Dim intQ As Integer
Dim rsQuery As New ADODB.Recordset
      
   '管制日期
   If txtField(0) <> "" Then
       strCon = strCon & " and cp27>=" & TransDate(txtField(0), 2)
   Else
       strCon = strCon & " and cp27>=" & CompWorkDay(5, strSrvDate(1), 1)
   End If
   If txtField(1) <> "" Then
       strCon = strCon & " and cp27<=" & TransDate(txtField(1), 2)
   Else
       strCon = strCon & " and cp27<=" & strSrvDate(1)
   End If
   '發文人員
   If Trim(Left(Combo1.Text, 6)) <> "" Then
       strCon = strCon & " and cp83=" & CNULL(Trim(Left(Combo1.Text, 6)))
   End If
   
    SetGrd True
    'Modified by Morgan 2019/9/17 有發文規費但無"已繳費"也要列出--敏莉
    strSql = "select '' as chk1,sqldatet(cp27) cp27t,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as caseno," & _
                 "nvl(pa05,nvl(pa06,pa07)) as casename,c1.cpm03,nvl(cu04,nvl(cu05,cu06)) as custname,s1.st02 as cp14n,s2.st02 as cp13n," & _
                 "pa11,s3.st02 as cp83n,cp83,CP09 " & _
                 "from caseprogress,patent,customer,staff s1,staff s2,staff s3 ,CasePropertyMap c1 " & _
                 "where cp01='FCP' And Instr(Cp64,'智慧局收文文號')>0 And (instr(cp64,'電子送件已稽核')=0 or (cp84>0 and instr(cp64,'已繳費')=0)) " & strCon & _
                 " And cp118 is not null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) " & _
                 "and cp14=s1.st01(+) and cp13=s2.st01(+) And CP01=c1.cpm01(+) And CP10=c1.cpm02(+) " & _
                 "and cp83=s3.st01(+) and substr(s3.st03,1,1)='F' "
    strSql = strSql & " order by 2, 3, cp09 "
    
    intQ = 1
    Set rsQuery = ClsLawReadRstMsg(intQ, strSql)
    MSHFlexGrid1.FixedCols = 0
    lblCount.Caption = "共 " & rsQuery.RecordCount & " 筆" 'Added by Morgan 2019/7/24"
    If intQ = 1 Then
         doQuery = True
         Set MSHFlexGrid1.Recordset = rsQuery
         SetGrd False
         MSHFlexGrid1.FixedCols = 5
    Else
         doQuery = False
    End If

   Set rsQuery = Nothing
End Function

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("V", "發文日", "本所案號", "案件名稱", "案件性質", "申請人", "承辦人", "智權人員", "申請案號", "發文人員", "CP83", "CP09")
   arrGridHeadWidth = Array(260, 840, 1000, 1000, 1000, 1000, 800, 920, 1000, 920, 0, 0)
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
   If colCase = 0 Then
      colCase = PUB_MGridGetId("本所案號", MSHFlexGrid1)
      colCp09 = PUB_MGridGetId("CP09", MSHFlexGrid1)
   End If
   
   For intI = 1 To MSHFlexGrid1.Rows - 1
      MSHFlexGrid1.row = intI
      For iRow = 0 To MSHFlexGrid1.Cols - 1
         MSHFlexGrid1.col = iRow
         MSHFlexGrid1.CellBackColor = &H80000005
         If iRow = 0 Then
            MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
         Else
            MSHFlexGrid1.CellAlignment = flexAlignLeftCenter
         End If
      Next iRow
   Next intI
        
   MSHFlexGrid1.Visible = True
End Sub

Private Sub MSHFlexGrid1_Click()
Dim intRow As Integer
Dim lngColor As Long
   With MSHFlexGrid1
       If .MouseRow > 0 Then
          intRow = .MouseRow
          .row = intRow
          .col = 4
          lngColor = .CellBackColor
          GridClick MSHFlexGrid1, intRow, 0, 0, 4, "V", lngColor
       End If
   End With
End Sub

Public Sub PubShowNextData()
Dim inX As Integer, inY As Integer
Dim Str01 As String
Dim lngColor As Long
Dim TempList As String
    
    If cmdState = 3 Then
        Cmd1(cmdState).Enabled = False '避免重複按到
    End If
    
    For inX = 1 To MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.row = inX
       MSHFlexGrid1.col = 0
       If Trim(MSHFlexGrid1.Text) = "V" Then
           MSHFlexGrid1.Text = ""
           MSHFlexGrid1.col = 0
           MSHFlexGrid1.CellBackColor = MSHFlexGrid1.BackColor
           MSHFlexGrid1.col = 4
           lngColor = MSHFlexGrid1.CellBackColor
           For inY = 1 To 3
               MSHFlexGrid1.col = inY
               MSHFlexGrid1.CellBackColor = lngColor
           Next inY
           '本所案號
           Str01 = Trim(MSHFlexGrid1.TextMatrix(inX, colCase))
           Str01 = Pub_RplStr(Str01) '清掉符號
           If InStr(Str01, "-") > 0 Then
               If InStrRev(Str01, "-") < 6 Then Str01 = Str01 & "-0-00"
           End If

           If Replace(Str01, "-", "") <> "" Then
                Select Case cmdState
                    Case 1 '進度檔
                        If fnSaveParentForm(Me) = False Then
                            Me.Enabled = True
                            Exit Sub
                        End If
                         frm100101_2.Show
                         frm100101_2.Tag = Pub_RplStr(Str01)
                         frm100101_2.StrMenu
                    Case 2 '下一程序檔
                         Call ChgCaseNo(Replace(Str01, "-", ""), strExc)
                         Call frm075007_1.SetParent(Me)
                         With frm075007_1
                              .Show
                              .textNP02 = strExc(1)
                              .textNP03 = strExc(2)
                              .textNP04 = strExc(3)
                              .textNP05 = strExc(4)
                              .cmdQuery_Click
                         End With
                    Case 3 '稽核完成
                         strExc(1) = Trim("" & MSHFlexGrid1.TextMatrix(inX, colCp09))
                         If strExc(1) <> "" And (TempList = "" Or (TempList <> "" And InStr(TempList, strExc(1)) = 0)) Then
                            TempList = TempList & strExc(1) & ","
                         End If
                End Select
           End If
           If cmdState <> 3 Then Exit For
       End If
    Next inX
    Me.Enabled = True
    
    If cmdState = 3 And TempList <> "" Then
        If SaveData(TempList) = True Then
            MsgBox ("電子送件稽核完成!")
        End If
    End If
    If cmdState = 3 Then
        Cmd1(cmdState).Enabled = True '避免重複按到
        Call Cmd1_Click(0)
    End If
    
    Exit Sub
    
ErrHand01:
    If Err.Number <> 0 Then
         MsgBox Err.Description
    End If

End Sub

Private Sub Combo1_Click()
      If Combo1.Tag <> "" And Combo1.Tag <> Combo1.Text Then
          If doQuery() = False Then
          End If
      End If
      Combo1.Tag = Combo1.Text
End Sub

Private Sub SetCombo1()

   Combo1.Clear
   strExc(0) = "select st01,st02 from staff a where st03='F22' and st04='1' order by 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not RsTemp.EOF
         If .Fields("st01") = strUserNum Then
            Combo1.AddItem .Fields("st01") & " " & .Fields("st02"), 0
            Combo1.Tag = .Fields("st01") & " " & .Fields("st02")
         Else
            Combo1.AddItem .Fields("st01") & " " & .Fields("st02")
         End If
      .MoveNext
      Loop
      End With
   End If
   Combo1.AddItem "      全部"
   If Combo1.Tag <> "" Then
      Combo1.ListIndex = 0
   Else
      Combo1.ListIndex = Combo1.ListCount - 1
   End If
   
   Call Cmd1_Click(0)
End Sub

Private Sub txtField_GotFocus(Index As Integer)
    TextInverse txtField(Index)
End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)

    If txtField(Index) = "" Then
        MsgBox "管制日期不可空白!", vbCritical
        GoTo JumpExit
    Else
        If CheckIsTaiwanDate(txtField(Index)) = False Then
            GoTo JumpExit
        End If
    End If
    If Index = 0 And txtField(0) > txtField(1) Then
         MsgBox "管制日期起值不可大於迄值 !"
         GoTo JumpExit
    End If
    
    Exit Sub
    
JumpExit:
    Cancel = True
    txtField(Index).SetFocus
    Call txtField_GotFocus(Index)
End Sub

Private Function SaveData(ByVal mList As String) As Boolean
Dim intP As Integer
Dim tmpArr As Variant

    tmpArr = Split(mList, ",")
    
    cnnConnection.BeginTrans
    For intP = 0 To UBound(tmpArr)
        If Trim(tmpArr(intP)) <> "" Then
            
            strSql = "Update CaseProgress set CP64='電子送件已稽核;'||CP64 Where CP09='" & tmpArr(intP) & "' and instr(CP64,'電子送件已稽核') = 0 "
            cnnConnection.Execute strSql, intI
            'Added by Morgan 2019/8/16 有發文規費的+已繳費--敏莉(與電子稽核一致)
            strSql = "Update CaseProgress set CP64='已繳費;'||CP64 Where CP09='" & tmpArr(intP) & "' and cp84>0 and instr(CP64,'已繳費') = 0"
            cnnConnection.Execute strSql, intI
            'end 2019/8/16
        End If
    Next intP
    cnnConnection.CommitTrans
    
    SaveData = True
    Exit Function
    
ErrHandle:
    If Err.Number <> "" Then
         MsgBox Err.Description
         cnnConnection.RollbackTrans
    End If
End Function

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

'Added by Morgan 2019/7/19
'電子稽核
Private Sub doBatch(Optional pUpdated As Boolean)
   Dim strFile As String
   Dim strText As String
   Dim arrRow() As String
   Dim arrCell() As String
   Dim ii As Integer, jj As Integer
   Dim bol1stRow As Boolean, iType As Integer, strMsg As String, strList As String
   Dim iCol1 As Integer, iCol2 As Integer, iCol3 As Integer, iCol4 As Integer
   Dim stVal1 As String, stVal2 As String, stVal3 As String, stVal4 As String
   Dim iErrCnt As Integer, iGoodCnt As Integer
   Dim bolNoChk3 As Boolean
   Dim arrErrMail() As String
   
On Error GoTo ErrHnd

   strFile = Dir(txt2Path & "\*.CSV")
   If strFile = "" Then
      If MsgBox("資料夾內無CSV檔！" & vbCrLf & vbCrLf & "是否要繼續檢查 " & txtField(1).Text & " 未繳費案件？" & vbCrLf & vbCrLf & "請確認無CSV檔可下載才可按 ""是"" 繼續！否則請先下載CSV檔後再進行稽核作業。", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
         '檢查有發文規費未繳費案件
         Call ChkData3(txtField(1).Text)
      End If
   Else

      Do
         iType = 0: bol1stRow = True: strMsg = "": strList = "": iErrCnt = 0: iGoodCnt = 0
         
         'Modified by Morgan 2019/7/24
         '收據Big5,送件UTF-8
         If InStr(UCase(strFile), "RECEIPT") > 0 Then
            strText = PUB_ReadTextFile(txt2Path & "\" & strFile)
         Else
            strText = PUB_ReadTextFile(txt2Path & "\" & strFile, "UTF-8")
         End If
         arrRow = Split(strText, vbCrLf)
         For ii = LBound(arrRow) To UBound(arrRow)
            If arrRow(ii) <> "" Then
               '欄位名稱
               If bol1stRow = True Then
                  bol1stRow = False
                  'Modified by Morgan 2019/7/25 +申請案號
                  If InStr(arrRow(ii), "收文文號") > 0 And InStr(arrRow(ii), "自訂案件編號") > 0 And InStr(arrRow(ii), "送件時間") > 0 And InStr(arrRow(ii), "申請案號") > 0 Then
                     iType = 1 '送件
                     
                  'Modified by Morgan 2019/7/24 未扣款前CSV檔沒有扣款日及收據號
                  'ElseIf InStr(arrRow(ii), "自訂案件編號") > 0 And InStr(arrRow(ii), "規費金額") > 0 And InStr(arrRow(ii), "收據日期") > 0 And InStr(arrRow(ii), "收據號碼") > 0 Then
                  ElseIf InStr(arrRow(ii), "自訂案件編號") > 0 And InStr(arrRow(ii), "規費金額") > 0 Then
                  
                     iType = 2 '收據
                  Else
                     MsgBox "【" & strFile & "】檔案內容有誤，請確認！", vbCritical
                     Exit Sub
                  End If
                  
                  arrCell = Split(arrRow(ii), ",")
                  iCol1 = -1: iCol2 = -1: iCol3 = -1
                  For jj = LBound(arrCell) To UBound(arrCell)
                     arrCell(jj) = Replace(arrCell(jj), """", "")
                     
                     '通用
                     If arrCell(jj) = "自訂案件編號" Then
                        iCol1 = jj
                     '送件
                     ElseIf iType = 1 Then
                        If arrCell(jj) = "收文文號" Then
                           iCol2 = jj
                        ElseIf arrCell(jj) = "送件時間" Then
                           iCol3 = jj
                        ElseIf arrCell(jj) = "申請案號" Then
                           iCol4 = jj
                        End If
                     '繳費
                     ElseIf iType = 2 Then
                        If arrCell(jj) = "規費金額" Then
                           iCol2 = jj
                           
                        'Removed by Morgan 2019/7/24 未扣款前CSV檔沒有扣款日及收據號
                        'ElseIf arrCell(jj) = "收據日期" Then
                        '   iCol3 = jj
                        'ElseIf arrCell(jj) = "收據號碼" Then
                        '   iCol4 = jj
                        'end 2019/7/24
                        
                        End If
                     End If
                  Next
                  
                  If iCol1 = -1 Then strMsg = strMsg & "自訂案件編號"
                  If iType = 1 Then
                     If iCol2 = -1 Then strMsg = strMsg & IIf(strMsg = "", "", ",") & "收文文號"
                     If iCol3 = -1 Then strMsg = strMsg & IIf(strMsg = "", "", ",") & "送件時間"
                     If iCol4 = -1 Then strMsg = strMsg & IIf(strMsg = "", "", ",") & "申請案號"
                  Else
                     If iCol2 = -1 Then strMsg = strMsg & IIf(strMsg = "", "", ",") & "規費金額"
                     
                     'Removed by Morgan 2019/7/24 未扣款前CSV檔沒有扣款日及收據號
                     'If iCol3 = -1 Then strMsg = strMsg & IIf(strMsg = "", "", ",") & "收據日期"
                     'If iCol4 = -1 Then strMsg = strMsg & IIf(strMsg = "", "", ",") & "收據號碼"
                     'end 2019/7/24
                     
                  End If
                  
                  If strMsg <> "" Then
                     MsgBox "【" & strFile & "】檔案下列欄位讀取失敗，請確認！" & vbCrLf & vbCrLf & strMsg, vbCritical
                     Exit Sub
                  End If
                  
               '資料
               Else
                  'Modified by Morgan 2019/7/26 案件名稱會有,號
                  'arrCell = Split(arrRow(ii), ",")
                  arrCell = Split(arrRow(ii), """,""")
                  stVal1 = Replace(arrCell(iCol1), """", "")
                  stVal2 = Replace(arrCell(iCol2), """", "")
                  If Left(stVal1, 3) = "FCP" Then '只做 FCP,忽略 P --敏莉
                     '送件清單稽核
                     If iType = 1 Then
                        stVal3 = Replace(arrCell(iCol3), """", "")
                        stVal4 = Replace(arrCell(iCol4), """", "")
                        If ChkData1(stVal1, stVal2, stVal3, stVal4, strMsg) = True Then
                           iGoodCnt = iGoodCnt + 1
                           pUpdated = True
                        Else
                           iErrCnt = iErrCnt + 1
                           If iErrCnt < 6 Then
                              strList = strList & strMsg & vbCrLf & vbCrLf
                           End If
                        End If
                        
                     '繳費清單稽核
                     Else
                     
                        'Modified by Morgan 2019/7/24 未扣款前CSV檔沒有扣款日及收據號
                        If ChkData2(stVal1, stVal2, txtField(1).Text, strMsg) = True Then
                           iGoodCnt = iGoodCnt + 1
                           pUpdated = True 'Added by Morgan 2021/5/28
                        Else
                           iErrCnt = iErrCnt + 1
                           If iErrCnt < 6 Then
                              strList = strList & strMsg & vbCrLf & vbCrLf
                           End If
                        'end 2019/7/24
                        
                        End If
                     End If
                  End If
               End If
            End If
         Next
         
         If strList <> "" Then
            If iErrCnt >= 6 Then strList = strList & "...等 "
            strList = "【" & strFile & "】稽核有異常！" & iGoodCnt & "筆成功，" & iErrCnt & " 筆失敗如下:" & vbCrLf & vbCrLf & strList
            MsgBox strList, vbExclamation
            'Kill txt2Path & "\" & strFile 'Removed by Morgan 2023/4/25 (敏莉又取消)若稽核異常則由主管人工稽核故也刪除--敏莉
         Else
            MsgBox "【" & strFile & "】稽核完成共" & iGoodCnt & "筆！", vbInformation
            Kill txt2Path & "\" & strFile
         End If
         
         'Modified by Morgan 2019/12/31 不論前面是否有異常都要繼續稽核
         If iType = 2 Then
            '檢查未繳費案件
            Call ChkData3(txtField(1).Text)
         End If
         
         If iType = 2 Then bolNoChk3 = True
         
         strFile = Dir()
      Loop While (strFile <> "")
      
      If bolNoChk3 = False Then
         If MsgBox("本次無繳費清單CSV要稽核，是否要檢查 " & txtField(1).Text & " 未繳費案件？", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
            Call ChkData3(txtField(1).Text)
         End If
      End If
      
   End If

   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub
'Added by Morgan 2019/7/19
'送件清單稽核
'pCaseNo:自訂案件編號(本所案號FCP-XXXXXX), pRecNo:收文文號, pAppDate:送件時間, pAppNo:申請案號, pErrMsg:錯誤訊息
Private Function ChkData1(pCaseNo As String, pRecNo As String, pAppDate As String, pAppNo As String, Optional pErrMsg As String) As Boolean
   Dim CNo() As String
   Dim iCount As Integer
   Dim strAppDate As String
   Dim strMsg As String
   Dim stHandler As String, stBoss As String
   Dim stSubj As String
   
   'Modified by Morgan 2019/10/2 一申請書多程序時後面會加.[發文數]
   'CNo() = Split(pCaseNo, "-")
   If InStr(pCaseNo, ".") > 0 Then
      CNo() = Split(pCaseNo, ".")
      iCount = Val(CNo(1))
      CNo() = Split(CNo(0), "-")
   Else
      iCount = 1 'Added by Morgan 2020/10/13 申請書漏書發文數也要稽核, Ex:FCP-35530
      CNo() = Split(pCaseNo, "-")
   End If
   'end 2019/10/2
   
   ReDim Preserve CNo(3) As String
   
   If CNo(2) = "" Then CNo(2) = "0"
   If CNo(3) = "" Then CNo(3) = "00"
   
   '智慧局清單不同地方下載日期格式不同(有民國也有西元)
   strAppDate = Left(DBDATE(pAppDate), 8)
   
   pErrMsg = ""
   strExc(0) = "select cp09,cp27,cpm03,cp84,cp83,st52,sqldatet(cp27) adate,pa11" & _
      " from caseprogress,casepropertymap,staff,patent" & _
      " where cp01='" & CNo(0) & "' and cp02='" & CNo(1) & "' and cp03='" & CNo(2) & "' and cp04='" & CNo(3) & "'" & _
      " and instr(cp64,'智慧局收文文號:" & pRecNo & ";')>0 and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp83" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      If iCount > 0 And iCount <> .RecordCount Then
         strMsg = "錯誤：" & vbTab & vbTab & " 此案件應有" & iCount & "道發文(系統發文數為 " & .RecordCount & ")，有漏/誤發文情況，請確認後由主管人工稽核。"
      Else
         ChkData1 = True
         strMsg = ""
         Do While Not .EOF
            
            If "" & .Fields("cp27") = strAppDate And "" & .Fields("pa11") = pAppNo Then
               strSql = "update caseprogress set cp64='電子送件已稽核;'||cp64 where cp09='" & .Fields("cp09") & "' and instr(cp64,'電子送件已稽核;')=0"
               cnnConnection.Execute strSql, intI
            Else
               ChkData1 = False
               stHandler = "" & .Fields("cp83")
               stBoss = "" & .Fields("st52")
               
               If strMsg <> "" Then strMsg = strMsg & vbCrLf & vbCrLf
               
               strMsg = strMsg & "案件性質：" & vbTab & .Fields("cpm03") & vbCrLf
               strMsg = strMsg & "發文日：" & vbTab & .Fields("adate") & vbCrLf
               If "" & .Fields("cp27") <> strAppDate Then
                  strMsg = strMsg & "錯誤：" & vbTab & vbTab & "發文日不符"
               End If
               
               If "" & .Fields("pa11") <> pAppNo Then
                  strMsg = strMsg & "本所申請案號：" & vbTab & .Fields("pa11") & vbCrLf
                  strMsg = strMsg & "E-SET申請案號：" & vbTab & pAppNo & vbCrLf
                  strMsg = strMsg & "錯誤：" & vbTab & vbTab & "申請案號與E-SET不符"
               End If
            End If
            .MoveNext
         Loop
      End If
      End With
   Else
      strMsg = strMsg & "錯誤：" & vbTab & vbTab & "收文文號與系統不符或無發文"
   End If
   
   If ChkData1 = False Then
      If stHandler = "" Then
         stHandler = PUB_GetFCPHandler(CNo(0), CNo(1), CNo(2), CNo(3))
         stBoss = PUB_GetFCPProSup(stHandler)
      End If
            
      If stHandler = "" Then stHandler = strUserNum 'Added by Morgan 2021/12/2 沒有收件者時發給自己 Ex:本所案號輸錯時 --Sharon,Phoebe
      
      pErrMsg = "收文文號：" & vbTab & pRecNo & vbCrLf & _
               "本所案號：" & vbTab & pCaseNo & vbCrLf & _
               "送件時間：" & vbTab & pAppDate & "(E-SET)" & vbCrLf & _
               strMsg
      
      stSubj = pCaseNo & "電子送件稽核異常通知！(收文文號：" & pRecNo & ")"
      'Modified by Morgan 2020/6/12 稽核人員也要收到通知--淑華
      PUB_SendMail strUserNum, stHandler, "", stSubj, pErrMsg, , , , , , stBoss & ";" & strUserNum
   End If
   
End Function

'Added by Morgan 2019/7/19
'Modified by Morgan 2019/7/24 未扣款前CSV檔沒有扣款日及收據號,改寫
'繳費清單稽核
'pCaseNo:自訂案件編號(本所案號FCP-XXXXXX),pFee:規費金額,pAppDate:本所發文日(預設當日),pErrMsg:錯誤訊息
Private Function ChkData2(pCaseNo As String, pFee As String, Optional pAppDate As String, Optional pErrMsg As String) As Boolean
   Dim stCP27 As String
   Dim CNo() As String
   Dim stErrMsg As String
   Dim stHandler As String, stBoss As String
   Dim stSubj As String
   
   'Modified by Morgan 2019/10/2 一申請書多程序時後面
   'CNo() = Split(pCaseNo, "-")
   If InStr(pCaseNo, ".") > 0 Then
      CNo() = Split(pCaseNo, ".")
      CNo() = Split(CNo(0), "-")
   Else
      CNo() = Split(pCaseNo, "-")
   End If
   'end 2019/10/2
   
   ReDim Preserve CNo(3) As String
   
   If CNo(2) = "" Then CNo(2) = "0"
   If CNo(3) = "" Then CNo(3) = "00"
   
   If pAppDate = "" Then
      stCP27 = strSrvDate(1)
   Else
      stCP27 = DBDATE(pAppDate)
   End If
   
   pErrMsg = ""
   '抓當日發文的來比對()--敏莉
   strExc(0) = "select IPONO,sum(cp84) Fee,max(cp152) FDate,count(distinct cp152) DCount,max(cp83) cp83,max(st52) st52" & _
      " from (select GETTEXTVALUE(CP64,'智慧局收文文號:',';') IPONO,CP09,CP84,CP152,cp83,st52" & _
      " from caseprogress,casepropertymap,staff" & _
      " where cp01='" & CNo(0) & "' and cp02='" & CNo(1) & "' and cp03='" & CNo(2) & "' and cp04='" & CNo(3) & "'" & _
      " and cp27=" & stCP27 & " and cp84>0 and cp118='A'" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp83) X group by IPONO"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      stErrMsg = ""
      Do While Not .EOF
         stHandler = "" & .Fields("cp83")
         stBoss = "" & .Fields("st52")
         
         If stErrMsg <> "" Then stErrMsg = stErrMsg & vbCrLf & vbCrLf
         
         stErrMsg = stErrMsg & "收文文號:" & vbTab & .Fields("IPONO") & vbCrLf
         stErrMsg = stErrMsg & "發文規費:" & vbTab & .Fields("Fee") & vbCrLf
         '規費
         If .Fields("Fee") <> pFee Then
            stErrMsg = stErrMsg & "錯誤：" & vbTab & vbTab & "規費金額不符"
         '同收文文號有不同扣款日
         ElseIf .Fields("DCount") <> 1 Then
            stErrMsg = stErrMsg & "錯誤：" & vbTab & vbTab & "扣款日異常"
            
         '上稽核註記
         Else
            'Modified by Morgan 2019/10/2 判斷有發文規費且自動扣款的才更新
            strSql = "update caseprogress set cp64='已繳費;'||cp64" & _
               " where cp01='" & CNo(0) & "' AND CP02='" & CNo(1) & "' AND CP03='" & CNo(2) & "' AND CP04='" & CNo(3) & "'" & _
               " AND INSTR(CP64,'智慧局收文文號:" & .Fields("IPONO") & ";')>0 and instr(cp64,'已繳費;')=0" & _
               " and cp27=" & stCP27 & " and cp84>0 and cp118='A'"
            cnnConnection.Execute strSql, intI
            ChkData2 = True
            Exit Do
         End If
         .MoveNext
      Loop
      End With
   Else
      stErrMsg = "錯誤：" & vbTab & vbTab & "E-SET有繳費，但系統無符合電子送件且有規費之資料"
   End If
   
   If ChkData2 = False Then
      If stHandler = "" Then
         stHandler = PUB_GetFCPHandler(CNo(0), CNo(1), CNo(2), CNo(3))
         stBoss = PUB_GetFCPProSup(stHandler)
      End If
      
      If stHandler = "" Then stHandler = strUserNum 'Added by Morgan 2023/8/18 沒有收件者時發給自己 Ex:本所案號輸錯時 --Sharon,Phoebe
   
      pErrMsg = "本所案號:" & vbTab & pCaseNo & vbCrLf & _
               "規費金額：" & vbTab & pFee & vbCrLf & _
               stErrMsg
      
      stSubj = pCaseNo & "電子送件繳費稽核異常通知！"
      'Modified by Morgan 2020/6/12 稽核人員也要收到通知--淑華
      PUB_SendMail strUserNum, stHandler, "", stSubj, pErrMsg, , , , , , stBoss & ";" & strUserNum
   End If
   
End Function

'Added by Morgan 2019/7/23
'檢查有發文無未繳費案件
Private Sub ChkData3(Optional pAppDate As String)
   Dim stCP27 As String, stMsg As String, stMsgList As String
   Dim stHandler As String, stBoss As String, iMsgCount As Integer
   Dim stSubj As String, stMailContent As String
   
   If pAppDate = "" Then
      stCP27 = strSrvDate(1)
   Else
      stCP27 = DBDATE(pAppDate)
   End If
   
   'Modified by Moran 2019/7/24 異常通知發文人及其主管
   strExc(0) = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CNo,CP09" & _
      ",cpm03,CP84,GETTEXTVALUE(CP64,'智慧局收文文號:',';') IPONO,cp83,st52" & _
      " from caseprogress,casepropertymap,staff" & _
      " where cp27=" & stCP27 & " and cp84>0 and cp118='A' and cp01='FCP'" & _
      " and instr(CP64,'智慧局收文文號:')>0 and instr(CP64,'已繳費;')=0" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp83" & _
      " order by cp83,IPONO"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      MsgBox "無未繳費案件！", vbInformation
   Else
      With RsTemp
      stMsgList = "有發文規費但無E-SET繳費紀錄共 " & .RecordCount & " 筆如下：" & vbCrLf & vbCrLf
      stHandler = .Fields("cp83")
      stBoss = "" & .Fields("st52")
      Do While Not .EOF
         If stHandler <> .Fields("cp83") Then
            stSubj = "有發文規費但無E-SET繳費紀錄通知(共 " & iMsgCount & " 筆)"
            'Modified by Morgan 2020/6/12 稽核人員也要收到通知--淑華
            PUB_SendMail strUserNum, stHandler, "", stSubj, stMailContent, , , , , , stBoss & ";" & strUserNum
            
            stHandler = .Fields("cp83")
            stBoss = "" & .Fields("st52")
            
            stMailContent = ""
            iMsgCount = 0
         End If
         
         iMsgCount = iMsgCount + 1
         
         stMsg = "收文文號：" & vbTab & .Fields("IPONO") & vbCrLf
         stMsg = stMsg & "本所案號：" & vbTab & .Fields("CNo") & vbCrLf
         stMsg = stMsg & "案件性質：" & vbTab & .Fields("cpm03") & vbCrLf
         stMsg = stMsg & "發文規費：" & vbTab & .Fields("cp84") & vbCrLf & vbCrLf
         
         If .AbsolutePosition < 6 Then
            stMsgList = stMsgList & stMsg
         ElseIf .AbsolutePosition = 6 Then
            stMsgList = stMsgList & "...等 "
         End If
         
         stMailContent = stMailContent & stMsg
         
         .MoveNext
      Loop
      End With
      stSubj = "有發文規費但無E-SET繳費紀錄通知(共 " & iMsgCount & " 筆)"
      'Modified by Morgan 2020/6/12 稽核人員也要收到通知--淑華
      PUB_SendMail strUserNum, stHandler, "", stSubj, stMailContent, , , , , , stBoss & ";" & strUserNum
      
      MsgBox stMsgList, vbExclamation
   End If
End Sub


