VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm077002 
   BorderStyle     =   1  '單線固定
   Caption         =   "顧問明細及統計"
   ClientHeight    =   5472
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8208
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   8300
   Begin VB.CommandButton cmdRead 
      Caption         =   "顧問記錄(&R)"
      Height          =   405
      Left            =   4530
      TabIndex        =   18
      Top             =   90
      Width           =   1245
   End
   Begin VB.CheckBox Check1 
      Caption         =   "含取消收文資料"
      Height          =   255
      Left            =   3930
      TabIndex        =   4
      Top             =   1003
      Width           =   2655
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3195
      Left            =   90
      TabIndex        =   17
      Top             =   2160
      Width           =   8025
      _ExtentX        =   14161
      _ExtentY        =   5630
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      AllowUserResizing=   3
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   405
      Left            =   6930
      TabIndex        =   16
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   405
      Left            =   5880
      TabIndex        =   15
      Top             =   90
      Width           =   915
   End
   Begin VB.TextBox txtQ 
      Height          =   280
      Index           =   7
      Left            =   4860
      MaxLength       =   6
      TabIndex        =   8
      Top             =   1710
      Width           =   800
   End
   Begin VB.TextBox txtQ 
      Height          =   280
      Index           =   6
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1710
      Width           =   600
   End
   Begin VB.TextBox txtQ 
      Height          =   280
      Index           =   5
      Left            =   930
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1710
      Width           =   600
   End
   Begin VB.TextBox txtQ 
      Height          =   280
      Index           =   4
      Left            =   930
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1350
      Width           =   345
   End
   Begin VB.TextBox txtQ 
      Height          =   280
      Index           =   3
      Left            =   2070
      MaxLength       =   7
      TabIndex        =   3
      Top             =   990
      Width           =   880
   End
   Begin VB.TextBox txtQ 
      Height          =   280
      Index           =   2
      Left            =   930
      MaxLength       =   7
      TabIndex        =   2
      Top             =   990
      Width           =   880
   End
   Begin VB.TextBox txtQ 
      Height          =   280
      Index           =   1
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "1"
      Top             =   630
      Width           =   345
   End
   Begin VB.TextBox txtQ 
      Height          =   280
      Index           =   0
      Left            =   930
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "1"
      Top             =   630
      Width           =   345
   End
   Begin MSForms.Label lblName 
      Height          =   300
      Left            =   5700
      TabIndex        =   19
      Top             =   1710
      Width           =   1470
      VariousPropertyBits=   27
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "諮詢人員："
      Height          =   225
      Index           =   5
      Left            =   3930
      TabIndex        =   14
      Top             =   1740
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "業務區：       　　－       　"
      Height          =   225
      Index           =   4
      Left            =   180
      TabIndex        =   13
      Top             =   1740
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "所　別：　　 （1北所 2中所 3南所 4高所）"
      Height          =   225
      Index           =   3
      Left            =   150
      TabIndex        =   12
      Top             =   1380
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "日　期：　　　　 　－"
      Height          =   225
      Index           =   2
      Left            =   180
      TabIndex        =   11
      Top             =   1020
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "查詢別：　     (1.明細 2.統計)"
      Height          =   225
      Index           =   1
      Left            =   3930
      TabIndex        =   10
      Top             =   658
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "收發文：　　   (1.收文 2.發文)"
      Height          =   225
      Index           =   0
      Left            =   180
      TabIndex        =   9
      Top             =   660
      Width           =   2415
   End
End
Attribute VB_Name = "frm077002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/11 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB、lblName
'Create by Lydia 2020/04/20 顧問明細及統計
Option Explicit
Dim intLastRow As Integer '記錄勾選最後一筆
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    cmdRead.Visible = False 'Added by Lydia 2020/05/27
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm077002 = Nothing
End Sub

Private Sub MSHFlexGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   MSHFlexGrid1.ToolTipText = ""
   'Modified by Lydia 2020/05/27
   'If MSHFlexGrid1.MouseRow > 0 And MSHFlexGrid1.MouseCol > 0 And txtQ(1) = "1" And MSHFlexGrid1.Cols >= 9 Then
   If MSHFlexGrid1.MouseRow > 0 And MSHFlexGrid1.MouseCol > 0 And txtQ(1) = "1" And MSHFlexGrid1.Cols >= 10 Then
      If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.MouseRow, MSHFlexGrid1.MouseCol) <> "" Then
         MSHFlexGrid1.ToolTipText = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.MouseRow, MSHFlexGrid1.MouseCol)
      End If
   Else
      MSHFlexGrid1.ToolTipText = ""
   End If
   
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow MSHFlexGrid1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   MSHFlexGrid1.col = nCol
   MSHFlexGrid1.row = nRow
   If Me.MSHFlexGrid1.row < 1 And Me.MSHFlexGrid1.Text <> "V" Then
      If InStr("時數,點數", Me.MSHFlexGrid1.Text) > 0 Then
         If m_blnColOrderAsc = True Then
            Me.MSHFlexGrid1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.MSHFlexGrid1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.MSHFlexGrid1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.MSHFlexGrid1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

'Added by Lydia 2020/05/27
Private Sub MSHFlexGrid1_Click()
   If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 1) <> "" Then
       GridClick MSHFlexGrid1, intLastRow, 0, 1
   End If
End Sub

Private Sub txtQ_GotFocus(Index As Integer)
    TextInverse txtQ(Index)
End Sub

Private Sub txtQ_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
         Case 5, 6, 7  '業務區、諮詢人員
              KeyAscii = UpperCase(KeyAscii)
    End Select
End Sub

Private Sub txtQ_Validate(Index As Integer, Cancel As Boolean)
Dim strMsg As String

    Select Case Index
         Case 0      '收／發文
              If Trim(txtQ(Index)) = "" Then
                  strMsg = "收發文不可空白！"
              ElseIf txtQ(Index) <> "1" And txtQ(Index) <> "2" Then
                  strMsg = "收發文請輸入 1 或 2！"
              End If
         Case 1      '明細／統計
              If Trim(txtQ(Index)) = "" Then
                  strMsg = "查詢別不可空白！"
              ElseIf txtQ(Index) <> "1" And txtQ(Index) <> "2" Then
                  strMsg = "查詢別請輸入 1 或 2！"
              End If
         Case 2, 3  '日期區間
              If IsEmptyText(txtQ(Index)) = False Then
                  If CheckIsTaiwanDate(txtQ(Index), False) = False Then
                     strMsg = "日期格式不正確！"
                  End If
              End If
         Case 4      '所別
              If IsEmptyText(txtQ(Index)) = False Then
                  If txtQ(Index) <> "1" And txtQ(Index) <> "2" And txtQ(Index) <> "3" And txtQ(Index) <> "4" Then
                     strMsg = "所別請輸入 1 ~ 4！"
                  End If
              End If
         Case 5, 6  '業務區
              '...
         Case 7      '諮詢人員
              If IsEmptyText(txtQ(Index)) = False Then
                  lblName.Caption = GetStaffName(txtQ(Index), True)
                  If lblName.Caption = "" Then
                     strMsg = "諮詢人員代號不存在！"
                  End If
              Else
                  lblName.Caption = ""
              End If
    End Select
    
    If strMsg <> "" Then GoTo ExceptExit
    
    Cancel = False
    Exit Sub
    
ExceptExit:
    Cancel = True
    MsgBox strMsg, vbExclamation, "檢核資料"
    
    txtQ(Index).SetFocus
    txtQ_GotFocus (Index)
End Sub

'查詢
Private Sub cmdQuery_Click()
Dim strTmpQ As String, strCon As String
Dim intQ As Integer
Dim rsQuery As New ADODB.Recordset
Dim tBox As TextBox
Dim tmpBol As Boolean

   For Each tBox In txtQ
       Call txtQ_Validate(tBox.Index, tmpBol)
       If tmpBol = True Then
           Exit Sub
       End If
   Next

   '日期區間
   If Trim(txtQ(2)) = "" And Trim(txtQ(3)) = "" Then
        MsgBox "請輸入日期區間！", vbExclamation, "檢核資料"
        txtQ(2).SetFocus
        Call txtQ_GotFocus(2)
        Exit Sub
   ElseIf Trim(txtQ(2)) <> "" And Trim(txtQ(3)) <> "" And Val(txtQ(2)) > Val(txtQ(3)) Then
        MsgBox "日期起值不可大於迄值！", vbExclamation, "檢核資料"
        txtQ(2).SetFocus
        Call txtQ_GotFocus(2)
        Exit Sub
   End If

   'Added By Lydia 2023/04/20
   ClearQueryLog (Me.Name) '清除查詢印表記錄檔欄位
   pub_QL05 = pub_QL05 & ";" & Me.Caption & ";收發文：" & txtQ(0) & IIf(txtQ(0) = "1", "收文", "發文")
   pub_QL05 = pub_QL05 & ";查詢別：" & txtQ(1) & IIf(txtQ(0) = "1", "明細", "統計")
   If Trim(txtQ(2) & txtQ(3)) <> "" Then pub_QL05 = pub_QL05 & ";日期：" & txtQ(2) & "-" & txtQ(3)
   If Trim(txtQ(4)) <> "" Then pub_QL05 = pub_QL05 & ";所別：" & txtQ(4)
   If Trim(txtQ(5) & txtQ(6)) <> "" Then pub_QL05 = pub_QL05 & ";業務區：" & txtQ(5) & "-" & txtQ(6)
   If Trim(txtQ(7)) <> "" Then pub_QL05 = pub_QL05 & ";諮詢人員：" & txtQ(7)
   If Check1.Value = 1 Then pub_QL05 = pub_QL05 & ";含取消收文資料"
   'end 2023/04/20
   
   '收／發文日期區間
   If Trim(txtQ(2)) <> "" Then
       strCon = strCon & " and " & IIf(txtQ(0) = "1", "cp05", "cp27") & ">=" & DBDATE(txtQ(2))
   End If
   If Trim(txtQ(3)) <> "" Then
       strCon = strCon & " and " & IIf(txtQ(0) = "1", "cp05", "cp27") & "<=" & DBDATE(txtQ(3))
   End If
   '所別
   If Trim(txtQ(4)) <> "" Then
       strCon = strCon & " and st06='" & txtQ(4) & "'"
   End If
   '業務區
   If Trim(txtQ(5)) <> "" Then
       strCon = strCon & " and st03>='" & txtQ(5) & "'"
   End If
   If Trim(txtQ(6)) <> "" Then
       strCon = strCon & " and st03<='" & txtQ(6) & "'"
   End If
   '諮詢人員
   If Trim(txtQ(7)) <> "" Then
       strCon = strCon & " and st01='" & txtQ(7) & "'"
   End If
   '含取消收文資料
   If Check1.Value = 0 Then
       strCon = strCon & " and nvl(cp57,0)=0 "
   End If

   '非A類: 工作時數=0不顯示0
   'Modified by Lydia 2020/05/27 明細查詢+V
   'Added by Lydia 2023/12/26
   If DBDATE(txtQ(3)) >= 新部門啟用日 Then
      strTmpQ = "select " & IIf(txtQ(1) = "1", " ' ' as V, ", "") & " sqldatet(cp05) cp05,cp09,cpm03,nvl(a0921,a0901) as a0901,nvl(a0922,a0902) as a0902,st01,st02,decode(cp113,0,null,cp113) as cp113,sqldatet(cp27) cp27,cp64,nvl(cp18,0) dot,sqldatet(cp57) cp57 " & _
                       "From caseprogress, staff, acc090, ACC090NEW, casepropertymap where cp01='LA' and cp02='999999' and cp03='0' and cp04='00' and cp13=st01(+) and st03=a0901(+) AND ST93=A0921(+) and cp09 not like 'A%' and cp01=cpm01(+) and cp10=cpm02(+) " & strCon
   Else
   'end 2023/12/26
      strTmpQ = "select " & IIf(txtQ(1) = "1", " ' ' as V, ", "") & " sqldatet(cp05) cp05,cp09,cpm03,a0901,a0902,st01,st02,decode(cp113,0,null,cp113) as cp113,sqldatet(cp27) cp27,cp64,nvl(cp18,0) dot,sqldatet(cp57) cp57 " & _
                       "From caseprogress, staff, acc090,casepropertymap where cp01='LA' and cp02='999999' and cp03='0' and cp04='00' and cp13=st01(+) and st03=a0901(+) and cp09 not like 'A%' and cp01=cpm01(+) and cp10=cpm02(+) " & strCon
   End If
   'A類：A類收文資料則業務區、諮詢人、時數放空白
   'Modified by Lydia 2020/05/27 明細查詢+V
   strTmpQ = strTmpQ & " union all select  " & IIf(txtQ(1) = "1", " ' ' as V, ", "") & " sqldatet(cp05) cp05,cp09,cpm03,' ' as a0901,' ' as a0902,' ' as st01, ' ' as st02,null as cp113,sqldatet(cp27) cp27,cp64,nvl(cp18,0) dot,sqldatet(cp57) cp57 " & _
                    "From caseprogress, staff, acc090,casepropertymap where cp01='LA' and cp02='999999' and cp03='0' and cp04='00' and cp13=st01(+) and st03=a0901(+) and cp09 like 'A%' and cp01=cpm01(+) and cp10=cpm02(+) " & strCon

   If txtQ(1) = "2" Then '統計
      'Modified by Lydia 2020/05/27 統一查詢+V
      strTmpQ = "select ' ' as V, a0901,a0902,st01,st02,count(cp09) a1,sum(cp113) as a2,sum(dot) as a3 from (" & strTmpQ & ") group by a0901,a0902,st01,st02"
      strTmpQ = strTmpQ & " order by a0901,st01 "
   ElseIf txtQ(1) = "1" Then '明細
      strTmpQ = strTmpQ & " order by cp05,a0901,cp09 "
   End If

   cmdRead.Visible = False 'Added by Lydia 2020/05/27
   Call SetGrd(True)
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
   If intQ = 1 Then
      InsertQueryLog (rsQuery.RecordCount) 'Added by Lydia 2023/04/20
      MSHFlexGrid1.FixedCols = 0
      Set MSHFlexGrid1.Recordset = rsQuery
      Call SetGrd
      If txtQ(1) = "1" Then cmdRead.Visible = True 'Added by Lydia 2020/05/27 明細才顯示
   Else
      InsertQueryLog (0) 'Added by Lydia 2023/04/20
      MsgBox "查無資料！", vbInformation
   End If
   
   Set rsQuery = Nothing
   
End Sub

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer, iR As Integer
Dim strTmp As String

   If txtQ(1) = "1" Then
        'Modified by Lydia 2020/05/27 統一查詢+V (260)
        arrGridHeadText = Array("V", "收文日", "總收文號", "案件性質", "A0901", "業務區", "ST01", "諮詢人", "時數", "發文日", "進度備註", "點數", "取消收文日")
        If Check1.Value = 0 Then
            arrGridHeadWidth = Array(260, 860, 1000, 1000, 0, 900, 0, 900, 650, 860, 2000, 900, 0)
        Else
            arrGridHeadWidth = Array(260, 860, 1000, 1000, 0, 900, 0, 900, 650, 860, 2000, 900, 1000)
        End If
   Else
        'Modified by Lydia 2020/05/27 統一查詢+V (260)
        arrGridHeadText = Array("V", "A0901", "業務區", "ST01", "諮詢人", "次數", "時數", "點數")
        arrGridHeadWidth = Array(260, 0, 900, 0, 900, 900, 900, 900)
   End If
   
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
           '數值=>靠右
           'Modified by Lydia 2020/05/27
           'If (txtQ(1) = "1" And InStr("07,10", Format(iRow, "00")) > 0) Or (txtQ(1) = "2" And InStr("04,05,06", Format(iRow, "00")) > 0) Then
           '    If (txtQ(1) = "1" And iRow = 10) Or (txtQ(1) = "2" And iRow = 6) Then
           If (txtQ(1) = "1" And InStr("08,11", Format(iRow, "00")) > 0) Or (txtQ(1) = "2" And InStr("05,06,07", Format(iRow, "00")) > 0) Then
               If (txtQ(1) = "1" And iRow = 11) Or (txtQ(1) = "2" And iRow = 7) Then
                    strTmp = Format(Val(MSHFlexGrid1.TextMatrix(iR, iRow)), "#,##0.000")  '點數顯示小數點3位
                    MSHFlexGrid1.TextMatrix(iR, iRow) = strTmp
               End If
               MSHFlexGrid1.CellAlignment = flexAlignRightCenter
           Else
               MSHFlexGrid1.CellAlignment = flexAlignLeftCenter ' 文字=>靠左  flexAlignCenterCenter
           End If
        Next iRow
   Next iR
   
   MSHFlexGrid1.Visible = True
End Sub

'Added by Lydia 2020/05/27
Private Sub cmdRead_Click()
Dim intX As Integer
Dim strCP09 As String

   With MSHFlexGrid1
       For intX = 1 To .Rows - 1
          If .TextMatrix(intX, 0) = "v" And "" & .TextMatrix(intX, 2) <> "" Then
             If Left("" & .TextMatrix(intX, 2), 1) <> "B" Then
                 MsgBox "非顧問記錄收文！", vbInformation
                 Exit Sub
             Else
                 strCP09 = strCP09 & "," & .TextMatrix(intX, 2)
             End If
          End If
       Next
   End With
   
   If strCP09 <> "" Then '進入顧問記錄畫面
        If PUB_CheckFormExist("frm077001") Then
            MsgBox "請先關閉〔顧問記錄〕畫面！"
            Exit Sub
        End If
        frm077001.SetData "LA", 0, True
        frm077001.SetData "999999", 1, False
        frm077001.SetData "0", 2, False
        frm077001.SetData "00", 3, False
        frm077001.SetData Mid(strCP09, 2), 7, False
        Set frm077001.m_PrevForm = Me
        frm077001.Show
        frm077001.QueryData
        Me.Hide
   Else
        MsgBox "請勾選收文！", vbInformation
   End If
End Sub
