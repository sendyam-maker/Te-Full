VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm060104_i 
   BorderStyle     =   1  '單線固定
   Caption         =   "繳年費CSV檔產生"
   ClientHeight    =   5916
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8508
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5916
   ScaleWidth      =   8508
   Begin VB.CommandButton Command1 
      Caption         =   "刪除"
      Height          =   252
      Index           =   1
      Left            =   3960
      TabIndex        =   16
      Top             =   336
      Width           =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新增"
      Height          =   252
      Index           =   0
      Left            =   3336
      TabIndex        =   15
      Top             =   336
      Width           =   600
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   1
      Left            =   1368
      MaxLength       =   3
      TabIndex        =   14
      Text            =   "FCP"
      Top             =   324
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   2
      Left            =   1848
      MaxLength       =   6
      TabIndex        =   13
      Top             =   324
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   3
      Left            =   2688
      MaxLength       =   1
      TabIndex        =   12
      Top             =   324
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   4
      Left            =   2928
      MaxLength       =   2
      TabIndex        =   11
      Top             =   324
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   1
      Left            =   168
      TabIndex        =   10
      Top             =   360
      Width           =   1236
   End
   Begin VB.OptionButton Option1 
      Caption         =   "大批"
      Height          =   180
      Index           =   0
      Left            =   168
      TabIndex        =   9
      Top             =   96
      Width           =   948
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "重新整理(&R)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   4632
      TabIndex        =   6
      Top             =   120
      Width           =   1155
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "選擇..."
      Height          =   315
      Left            =   7245
      TabIndex        =   3
      Top             =   684
      Width           =   825
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   1428
      TabIndex        =   2
      Top             =   660
      Width           =   5796
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "產生CSV檔(&C)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5928
      TabIndex        =   1
      Top             =   120
      Width           =   1308
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7260
      TabIndex        =   0
      Top             =   120
      Width           =   800
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3336
      Top             =   5376
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   4425
      Left            =   150
      TabIndex        =   5
      Top             =   1080
      Width           =   8145
      _ExtentX        =   14372
      _ExtentY        =   7811
      _Version        =   393216
      Cols            =   13
      FixedCols       =   2
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
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
      _Band(0).Cols   =   13
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "已勾選案件數："
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   5610
      Width           =   1260
   End
   Begin VB.Label lblCount 
      Height          =   180
      Left            =   1665
      TabIndex        =   7
      Top             =   5610
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CSV檔案路徑："
      Height          =   180
      Index           =   0
      Left            =   168
      TabIndex        =   4
      Top             =   720
      Width           =   1236
   End
End
Attribute VB_Name = "frm060104_i"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/18 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Created by Morgan 2012/8/28
Option Explicit
Dim m_strCaseList As String 'Added by Morgan 2025/10/17

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
   Case 0
      ReadData
   Case 1
      If CheckData Then
         BuileFile
      End If
   Case 2
      Unload Me
   End Select
End Sub

Private Sub cmdOpen_Click()
   Dim strPath As String
   strPath = GetSaveName(txtPath)
   If strPath <> "" Then
      txtPath = strPath
   End If
End Sub

Private Function GetSaveName(ByVal pFileName As String) As String
   Dim strPath As String, strFileName As String
   
   If InStrRev(pFileName, "\") > 0 Then
      strPath = Left(pFileName, InStrRev(pFileName, "\") - 1)
      strFileName = Mid(pFileName, InStrRev(pFileName, "\") + 1)
   Else
      'strPath = PUB_Getdesktop
      'Modified by Morgan 2015/8/13 CSV檔只留最後一次的
      'strPath = EFilePath & "\CSV\" & Format(Now, "YYYYMMDD")
      strPath = EFilePath & "\CSV"
      strFileName = pFileName
   End If
   
On Error GoTo ErrHnd

   If Dir(strPath, vbDirectory) = "" Then
      MkDir strPath
   End If

   With CommonDialog1
      .CancelError = True
      .FileName = strFileName
      .Filter = "CSV 檔 (*.CSV)|*.CSV"
      .InitDir = strPath
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowSave
      If .FileName <> "" Then
         GetSaveName = .FileName
      End If
   End With
   
   Exit Function
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Function

Private Sub Command1_Click(Index As Integer)
   Dim strTmp As String, ii As Integer
   If Index = 0 Then
      If Text1(2).Text <> "" Then
         strTmp = Text1(1) & Text1(2)
         If Text1(3).Text = "" Then
            strTmp = strTmp & "0"
         Else
            strTmp = strTmp & Text1(3).Text
         End If
         If Text1(4).Text = "" Then
            strTmp = strTmp & "00"
         Else
            strTmp = strTmp & Text1(4).Text
         End If
         
         m_strCaseList = m_strCaseList & "," & strTmp
         Text1(2).Text = ""
         Text1(2).SetFocus
      End If
   Else
      With GrdDataList
      For ii = 1 To .Rows - 1
         If .TextMatrix(ii, 0) = "V" Then
            strTmp = "," & Left(Replace(.TextMatrix(ii, 1), "-", "") & "000", 12)
            m_strCaseList = Replace(m_strCaseList, strTmp, "")
         End If
      Next ii
      End With
   End If
   cmdOK(0).Value = True
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   'txtPath = PUB_Getdesktop & "\" & Format(Now, "YYYYMMDD") & "_" & Format(Now, "HHMMSS") & ".CSV"
   'Modified by Morgan 2015/8/13 CSV檔只留最後一次的
   'txtPath = EFilePath & "\CSV\" & Format(Now, "YYYYMMDD") & "\" & Format(Now, "YYYYMMDD") & "_" & Format(Now, "HHMMSS") & ".CSV"
   'Modified by Morgan 2023/6/28 新版網頁只接受小寫的副檔名
   'txtPath = EFilePath & "\CSV\" & Format(Now, "YYYYMMDD") & "_" & Format(Now, "HHMMSS") & ".CSV"
   'Modified by Morgan 2025/10/17 檔名增加員工號
   txtPath = EFilePath & "\CSV\" & Format(Now, "YYYYMMDD") & "_" & Format(Now, "HHMMSS") & "_" & strUserNum & ".csv"
   'end 2023/6/28
   
   'Modified by Morgan 2025/10/17
   'SetDataListWidth
   'ReadData
   Option1(0).Value = True
   'end 2025/10/17
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060104_i = Nothing
End Sub
Private Sub ReadData()
   Dim stCon As String
   
   'Added by Morgan 2025/10/17
   If m_strCaseList <> "" Then
      If m_strCaseList = "X" Then
         stCon = " and 1=0"
      Else
         stCon = " and instr('" & m_strCaseList & "',cp01||cp02||cp03||cp04)>0"
      End If
   End If
   'end 2025/10/17
   
   lblCount = ""
   '排除補繳案件(仍從單筆發文作業)
   'Modified by Morgan 2022/3/3 +CP164,條件+指定日期方式(CP164)=之前(2)也要列出
   'Modified by Morgan 2022/11/21 指定當日繳納一律不顯示(都單筆繳)--Sharon
   strExc(0) = "select 'V',cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)" & _
      ",st02,sqldatet(cp07),sqldatet(cp142)||decode(cp142,null,null,decode(cp164,'1','(當天)','2','(之前)','3','(之後)')),pa11,pa22,cp53,cp54,pa01||pa02||pa03||pa04 CaseNo,pa08,0 Fee,'' Red,'' RedType,PA16" & _
      ",pa14" & _
      " From caseprogress, patent,staff" & _
      " where cp05>" & (strSrvDate(1) - 10000) & " and cp10='605'" & stCon & _
      " and cp01||cp27||cp57='FCP'" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and nvl(cp141,'1')<>'4' and nvl(cp141||cp164,'1')<>'31' and (cp142 is null or cp142<=to_char(sysdate,'yyyymmdd') or cp164='2')" & _
      " and st01(+)=cp14" & _
      " order by cp14,cp01,cp02"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If RsTemp.RecordCount = 0 Then
      SetDataListWidth True
   Else
      GrdDataList.Visible = False
      GrdDataList.FixedCols = 0
      Set GrdDataList.Recordset = RsTemp
      SetDataListWidth
      GrdDataList.FixedCols = 2
      If RsTemp.RecordCount > 0 Then
         lblCount = RsTemp.RecordCount
         CheckData True
      End If
      GrdDataList.Visible = True
   End If
End Sub

Private Sub SetDataListWidth(Optional pReset As Boolean)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iCol As Integer
   arrGridHeadText = Array("V", "本所案號", "承辦人", "法定期限", "指定日期" _
                     , "申請案號", "證書號", "繳費起年", "繳費迄年", "PA0104", "專利種類", "規費", "是否減免", "減免類型", "准駁", "公告日")
   arrGridHeadWidth = Array(250, 1100, 700, 850, 1300 _
                     , 1100, 850, 820, 820, 0, 0, 0, 0, 0, 0, 0)
                        
   GrdDataList.Cols = UBound(arrGridHeadText) + 1
   
   'Added by Morgan 2025/10/17
   If pReset = True Then
      GrdDataList.Clear
      GrdDataList.Rows = 2
   End If
   'end 2025/10/17
   
   For iCol = 0 To GrdDataList.Cols - 1
      GrdDataList.row = 0
      GrdDataList.col = iCol
      GrdDataList.Text = arrGridHeadText(iCol)
      GrdDataList.ColWidth(iCol) = arrGridHeadWidth(iCol)
      GrdDataList.CellAlignment = flexAlignCenterCenter
   Next iCol
   GrdDataList.ColAlignment(7) = flexAlignRightCenter
   GrdDataList.ColAlignment(8) = flexAlignRightCenter
   GrdDataList.ColAlignment(0) = flexAlignCenterCenter
   GrdDataList.ColAlignmentFixed(0) = flexAlignCenterCenter
   GrdDataList.BackColor = &HFFC0C0
End Sub

Private Sub grdDataList_Click()
'   intI = grdDataList.MouseRow
'   grdDataList.row = intI
'   If grdDataList.row <> 0 Then
'      ClickGrid
'   End If
End Sub

Private Sub grdDataList_SelChange()
   GrdDataList.row = GrdDataList.MouseRow
   If GrdDataList.row <> 0 Then
      ClickGrid
   End If
End Sub

Private Sub ClickGrid()
Dim i As Integer
GrdDataList.Visible = False
GrdDataList.col = 0
If GrdDataList.Text = "V" Then
   lblCount = Val(lblCount) - 1
     GrdDataList.Text = ""
     For i = 2 To GrdDataList.Cols - 1
          GrdDataList.col = i
          GrdDataList.CellBackColor = QBColor(15)
    Next i
Else
   lblCount = Val(lblCount) + 1
     GrdDataList.Text = "V"
     For i = 2 To GrdDataList.Cols - 1
         GrdDataList.col = i
         GrdDataList.CellBackColor = &HFFC0C0
     Next i
End If
GrdDataList.Visible = True
End Sub

Private Function CheckData(Optional pAutoCheck As Boolean) As Boolean
   Dim ii As Integer
   Dim strCP81 As String, strDiscType As String, lngDisc As Long, strFee As String, bolIsDelay As Boolean, strDueDate As String
   Dim bolChecked As Boolean
   
   If txtPath = "" Then MsgBox "請輸入CSV檔案路徑！"
   
   With GrdDataList
   For ii = 1 To .Rows - 1
      If .TextMatrix(ii, 0) = "V" Then
         bolChecked = True
         If .TextMatrix(ii, 5) = "" Then
            If pAutoCheck Then
               '.TextMatrix(ii, 0) = ""
               .row = ii
               ClickGrid
            Else
               MsgBox .TextMatrix(ii, 1) & " 申請案號空白，請修正！"
               Exit For
            End If
         ElseIf .TextMatrix(ii, 6) = "" Then
            If pAutoCheck Then
               '.TextMatrix(ii, 0) = ""
               .row = ii
               ClickGrid
            Else
               MsgBox .TextMatrix(ii, 1) & " 申請案號空白，請修正！"
               Exit For
            End If
         ElseIf .TextMatrix(ii, 7) = "" Then
            If pAutoCheck Then
               '.TextMatrix(ii, 0) = ""
               .row = ii
               ClickGrid
            Else
               MsgBox .TextMatrix(ii, 1) & " 繳費起年空白，請修正！"
               Exit For
            End If
         ElseIf .TextMatrix(ii, 8) = "" Then
            If pAutoCheck Then
               '.TextMatrix(ii, 0) = ""
               .row = ii
               grdDataList_SelChange
            Else
               MsgBox .TextMatrix(ii, 1) & " 繳費迄年空白，請修正！"
               Exit For
            End If
            
         ElseIf Not pAutoCheck Then
            '減免
            strCP81 = PUB_GetCaseDiscStat(.TextMatrix(ii, 9), strDiscType)
            '逾繳
            '原法限(遇假日順延)
            strDueDate = CompDate(0, Val(.TextMatrix(ii, 7)) - 1, .TextMatrix(ii, 15))
            strDueDate = CompDate(2, -1, strDueDate)            '
            strDueDate = PUB_GetWorkDay1(strDueDate, False)
            If Val(strSrvDate(1)) > Val(strDueDate) Then
               bolIsDelay = True
            Else
               bolIsDelay = False
            End If
            PUB_GetPatentYearFee "000", .TextMatrix(ii, 10), "Y00000000", 年費, .TextMatrix(ii, 7), .TextMatrix(ii, 8), bolIsDelay, strCP81, .TextMatrix(ii, 15), strSrvDate(2), strFee, , lngDisc
            .TextMatrix(ii, 11) = strFee
            If lngDisc > 0 Then
               .TextMatrix(ii, 12) = "Y"
               '0:自然人,1:中小企業,2:學校
               '自然人(本所代碼:1)
               If Left(strDiscType, 1) = "1" Then
                  .TextMatrix(ii, 13) = "0"
               '學校(本所代碼:2)
               ElseIf Left(strDiscType, 1) = "2" Then
                  .TextMatrix(ii, 13) = "2"
               '中小企業(本所代碼:3)
               Else
                  .TextMatrix(ii, 13) = "1"
               End If
            Else
               .TextMatrix(ii, 12) = ""
               .TextMatrix(ii, 13) = ""
            End If
         End If
      End If
   Next
   If bolChecked = False Then MsgBox "請至少點選一筆資料！": Exit Function
   If ii = .Rows Then CheckData = True
   End With
End Function

'Modified by Morgan 2023/6/20 e網通新網頁改用UTF-8編碼的CSV(以LF符號斷行)
Private Sub BuileFile()
   Dim ii As Integer
   Dim ff As Integer, strData As String
   Dim strPath As String, strFileName As String
   
On Error GoTo ErrHnd

   If InStrRev(txtPath, "\") > 0 Then
      strPath = Left(txtPath, InStrRev(txtPath, "\") - 1)
      strFileName = Mid(txtPath, InStrRev(txtPath, "\") + 1)
   Else
      'Modified by Morgan 2015/8/13 CSV檔只留最後一次的
      'strPath = EFilePath & "\CSV\" & Format(Now, "YYYYMMDD")
      strPath = EFilePath & "\CSV"
      strFileName = txtPath
   End If
   
   If Dir(strPath, vbDirectory) = "" Then
      MkDir strPath
   'Modified by Morgan 2025/10/17 因可能多人操作，改只刪除自己的檔案
   'ElseIf Dir(strPath & "\*.CSV") <> "" Then
   '   Kill strPath & "\*.CSV"
   Else
      strExc(1) = strPath & "\*_" & strUserNum & ".CSV"
      If Dir(strExc(1)) <> "" Then
         Kill strExc(1)
      End If
   'end 2025/10/17
   End If
   
   'Removed by Morgan 2023/6/20
   'If ff > 0 Then Close #ff
   'ff = FreeFile
   'Open txtPath For Output As ff
   'end 2023/6/20
   
   With GrdDataList
   strData = ""
   For ii = 1 To .Rows - 1
      If .TextMatrix(ii, 0) = "V" Then
         'Mofified by Morgan 2023/6/20
         'strData = ""
         If strData <> "" Then strData = strData & vbLf 'e網通只接受以LF符號斷行的CSV檔(用CRLF會格式錯誤)
         'end 2023/6/20
         strData = strData & .TextMatrix(ii, 5) & "," & .TextMatrix(ii, 6) & "," & .TextMatrix(ii, 7) & "," & .TextMatrix(ii, 8) & "," & .TextMatrix(ii, 11)
         strData = strData & ",A" 'Added by Morgan 2012/12/11 +收據抬頭(固定用A 專利權人)--靜芳
         If .TextMatrix(ii, 12) = "Y" Then
            strData = strData & ",Y," & .TextMatrix(ii, 13)
         Else
            strData = strData & ",,"
         End If
         strData = strData & ",1" 'Added by Morgan 2023/6/20 電子收據
         
         'Print #ff, strData 'Removed by Morgan 2023/6/20
      End If
   Next
   End With
   
   'Modified by Morgan 2023/6/20 轉存成UTF-8
   'Close ff
   SaveUTF8NoBOM txtPath, strData
   'end 2023/6/19
   
   MsgBox "CSV 檔已產生！"
   Exit Sub
   
ErrHnd:
   MsgBox Err.Description
   
End Sub

'Added by Morgan 2023/6/19
Private Sub SaveUTF8NoBOM(filePath, Text)
   Const adSaveCreateNotExist = 1
   Const adSaveCreateOverWrite = 2
   Const adModeReadWrite = 3
   Const adTypeBinary = 1
   Const adTypeText = 2
   
   Dim StreamUTF8 As New ADODB.Stream
   Dim StreamUTF8NoBOM  As New ADODB.Stream
   
   With StreamUTF8
     .Charset = "UTF-8"
     .Type = adTypeText
     .Mode = adModeReadWrite
     .Open
     .WriteText Text
     .Position = 3
   End With
   
   With StreamUTF8NoBOM
     .Type = adTypeBinary
     .Mode = adModeReadWrite
     .Open
     StreamUTF8.CopyTo StreamUTF8NoBOM
     .SaveToFile filePath, adSaveCreateOverWrite
   End With
   
   StreamUTF8.Close
   StreamUTF8NoBOM.Close
   
   Set StreamUTF8 = Nothing
   Set StreamUTF8NoBOM = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   If Index = 0 Then
      m_strCaseList = ""
      Command1(0).Enabled = False
      Command1(1).Enabled = False
   Else
      Command1(0).Enabled = True
      Command1(1).Enabled = True
      If m_strCaseList = "" Then
         m_strCaseList = "X"
      End If
   End If
   cmdOK(0).Value = True
End Sub
