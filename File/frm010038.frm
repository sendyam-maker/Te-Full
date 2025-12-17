VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm010038 
   BorderStyle     =   1  '單線固定
   Caption         =   "公告本整批匯入"
   ClientHeight    =   5820
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7992
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   7992
   Begin VB.CommandButton cmdEmail 
      Caption         =   "Email"
      Height          =   285
      Left            =   4350
      TabIndex        =   17
      Top             =   5520
      Width           =   1125
   End
   Begin VB.TextBox txtDate 
      Height          =   315
      Left            =   3210
      TabIndex        =   16
      Top             =   5490
      Width           =   1095
   End
   Begin VB.CommandButton CmdTxt 
      Caption         =   "全選"
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton CmdTxt 
      Caption         =   "複製申請號"
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdTxt 
      Caption         =   "複製公告號"
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSGrd1 
      Height          =   3855
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   6795
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      FormatString    =   "V|公 告 號|公 告 日|申 請 案 號|本 所 案 號|案　　件　　名　　稱"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "結束(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6720
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "匯入(&E)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtPath1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Text            =   "C:\temp\公告本"
      Top             =   840
      Width           =   5745
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<="
      Height          =   315
      Left            =   7440
      TabIndex        =   0
      Top             =   840
      Width           =   345
   End
   Begin VB.Label lblDate 
      Caption         =   "公告日："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2370
      TabIndex        =   15
      Top             =   5520
      Width           =   825
   End
   Begin VB.Label lblCnt2 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   5520
      Width           =   315
   End
   Begin VB.Label Label4 
      Caption         =   "勾選，共 　 筆"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   150
      TabIndex        =   13
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "P.S. 智慧局檢索系統建議使用IE瀏覽器。"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   540
      Width           =   3285
   End
   Begin VB.Label Label2 
      Caption         =   "備註：可勾選多筆記錄後，按下複製按鈕即可複製多筆公告號或申請案號。"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1300
      Width           =   7215
   End
   Begin VB.Label LblCnt 
      Caption         =   "查詢，共 Ｘ 筆"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "電子檔存放路徑："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frm010038"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/11/30 Form2.0已修改 MSGrd1
'Create by Lydia 2019/02/01 公告本整批匯入
Option Explicit
Private Const cStartDate As String = "20190401" '開始下載公報期別
Dim rsAD As New ADODB.Recordset
Public cmdState As Integer '紀錄作用按鍵
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序

Dim intJ As Integer
Dim tmpArr As Variant

Dim colPA15 As Integer, colPA11 As Integer, colPA01 As Integer '記錄Grid的欄位
Dim strImportList As String '記錄匯入卷宗區的收文號
Dim strImportDList As String '記錄匯入卷宗區的期別
Dim strFileList As String '附件-調件清單
Private Const iLimit As Integer = 46    '單頁最大列數
Dim iPage As Integer, intCounter As Integer
Dim m_WordLeft As Long, m_WordTop As Long 'Added by Lydia 2019/04/09 Word開啟位置
Dim bVisible As Boolean 'Added by Lydia 2019/04/09

Private Sub Cmd1_Click(Index As Integer)

    Select Case Index
        Case 0 '查詢
             Screen.MousePointer = vbHourglass
             If QueryData(True) = False Then
             End If
             Screen.MousePointer = vbDefault
        Case 1 '匯入
             Screen.MousePointer = vbHourglass
             'Added by Lydia 2019/04/09
             Cmd1(0).Enabled = False
             Cmd1(1).Enabled = False
             CmdTxt(0).Enabled = False
             CmdTxt(1).Enabled = False
             CmdTxt(2).Enabled = False
             'end 2019/04/09
             strImportList = ""
             strFileList = ""
             strImportDList = ""
             If AutoUpdCpp() = True Then
             End If
             If strImportDList <> "" Then '發email
                If QueryData(False) = False Then
                End If
                '重整Grid
                If QueryData(True) = False Then
                End If
             End If
             'Added by Lydia 2019/04/09
             Cmd1(0).Enabled = True
             Cmd1(1).Enabled = True
             CmdTxt(0).Enabled = True
             CmdTxt(1).Enabled = True
             CmdTxt(2).Enabled = True
             'end 2019/04/09
             Screen.MousePointer = vbDefault
        Case 2 '結束
              Unload Me
    End Select
End Sub

Private Sub CmdTxt_Click(Index As Integer)
Dim intP As Integer
Dim iRow As Integer
Dim strCopyTxt As String ' 複製編號文字
 
    If MSGrd1.Rows < 2 Then Exit Sub

    For iRow = 1 To MSGrd1.Rows - 1
         If "" & MSGrd1.TextMatrix(iRow, 1) <> "" Then
            Select Case Index
                  Case 0 '複製公告號
                       If "" & MSGrd1.TextMatrix(iRow, 0) <> "" And "" & MSGrd1.TextMatrix(iRow, colPA15) <> "" Then
                           'Modified by Lydia 2021/09/15 原本智慧局網站搜尋可以用;區隔多筆進行查詢，現在只能 or 做區隔；經過「號碼檢索」和「簡易檢索」皆可使用。
                           'strCopyTxt = strCopyTxt & MSGrd1.TextMatrix(iRow, colPA15) & ";"
                           strCopyTxt = IIf(strCopyTxt <> "", strCopyTxt & " OR ", "") & MSGrd1.TextMatrix(iRow, colPA15)
                           intP = intP + 1
                       End If
                  Case 1 '複製申請號
                       If "" & MSGrd1.TextMatrix(iRow, 0) <> "" And "" & MSGrd1.TextMatrix(iRow, colPA11) <> "" Then
                           'Modified by Lydia 2021/09/15 原本智慧局網站搜尋可以用;區隔多筆進行查詢，現在只能 or 做區隔
                           'strCopyTxt = strCopyTxt & MSGrd1.TextMatrix(iRow, colPA11) & ";"
                           strCopyTxt = IIf(strCopyTxt <> "", strCopyTxt & " OR ", "") & MSGrd1.TextMatrix(iRow, colPA11)
                           intP = intP + 1
                       End If
                  Case 2 '全選/取消
                     MSGrd1.col = 0
                     MSGrd1.row = iRow
                     intP = intP + 1
                    If CmdTxt(Index).Caption = "全選" Then
                         MSGrd1.Text = "V"
                    Else
                         MSGrd1.Text = ""
                    End If
                    'Added by Lydia 2019/06/12 底色統一為白色
                    For intI = 0 To MSGrd1.Cols - 1
                          MSGrd1.col = intI
                          MSGrd1.CellBackColor = QBColor(15)
                    Next intI
            End Select
         End If
    Next iRow
    
    If strCopyTxt <> "" Then
        '複製編號至剪貼簿
        Clipboard.Clear
        Clipboard.SetText strCopyTxt
        If Index = 0 Then
            MsgBox "公告號已複製(" & intP & ") ", , MsgText(21)
        Else
            MsgBox "申請案號已複製(" & intP & ") ", , MsgText(21)
        End If
    ElseIf Index = 2 And intP > 0 Then
        If CmdTxt(Index).Caption = "全選" Then
            CmdTxt(Index).Caption = "取消"
            lblCnt2.Caption = intP  'Added by Lydia 2019/06/12 勾選筆數
        Else
            CmdTxt(Index).Caption = "全選"
            lblCnt2.Caption = "0"  'Added by Lydia 2019/06/12 勾選筆數
        End If
    End If
    
End Sub

Private Sub Form_Load()

   '刪除舊檔
   If Dir(App.path & "\*調卷清單.pdf") <> "" Then
         Kill App.path & "\*調卷清單.pdf"
   End If
   
   MoveFormToCenter Me
  
   strExc(1) = GetSetting("TAIE", "FileGAZ", UCase(Me.Name) & "Dir", "")
   If strExc(1) <> "" Then
      txtPath1.Text = strExc(1)
   '預設個人桌面
   Else
      txtPath1.Text = PUB_Getdesktop
   End If
   
   Call Cmd1_Click(0)
   
   'Added by Lydia 2023/04/12
   If Pub_StrUserSt03 <> "M51" Then
      lblDate.Visible = False
      txtDate.Visible = False
      cmdEmail.Visible = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm010038 = Nothing
End Sub

Private Sub Command2_Click()
Dim sFile
   
On Error GoTo ErrHnd
   
   With CommonDialog1
      .CancelError = True
      .FileName = "*.pdf"
      .Filter = "PDF檔案 (*.pdf)|*.pdf"
      If GetSetting("TAIE", "FileGAZ", UCase(Me.Name) & "Dir", "") <> "" Then
         .InitDir = GetSetting("TAIE", "FileGAZ", UCase(Me.Name) & "Dir", "")
      Else
         .InitDir = PUB_Getdesktop
      End If
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
            txtPath1.Text = sFile(0)
         Else
            '記錄路徑
            If InStr(.FileName, "\") > 0 Then
                SaveSetting "TAIE", "FileGAZ", UCase(Me.Name) & "Dir", Left(.FileName, InStrRev(.FileName, "\") - 1)
            End If
            txtPath1.Text = Left(.FileName, InStrRev(.FileName, "\") - 1)
         End If
      End If
   End With
   
   Exit Sub
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Function QueryData(ByVal bolRefresh As Boolean) As Boolean
Dim strGrp As String, tmpCont As String
Dim stCn1 As String, stCn2 As String
Dim strDTlist As String
Dim tmpArr As Variant

   If bolRefresh = True Then
      '清空及預設欄位值
      Call SetGrd(True)
      lblCnt2.Caption = "0"  'Added by Lydia 2019/06/12 勾選筆數
   End If
   
   stCn1 = " AND C1.CP27>=" & cStartDate
   '指定匯入的期別(同一期別匯完後才通知->判斷最後一筆)
   If bolRefresh = False And strImportDList <> "" Then
       stCn1 = ""
       tmpArr = Split(strImportDList, ",")
       For intI = 0 To UBound(tmpArr)
           If Trim(tmpArr(intI)) <> "" Then
               stCn1 = stCn1 & " OR C1.CP27=" & Trim(tmpArr(intI))
           End If
       Next intI
       If stCn1 <> "" Then
           stCn1 = "AND (" & Mid(stCn1, 4) & " ) "
       Else
            MsgBox "匯入的期別" & strImportDList & vbCrLf & ", 資料有問題請洽電腦中心！", vbCritical
            Exit Function
       End If
   End If
   
  '查詢->重整Grid
  If bolRefresh = True Then
      stCn2 = "SELECT '' AS CHK1,PA15,SQLDATET(PA14) AS PA14,PA11," & _
                    " DECODE(PA03||PA04,'000',PA01||'-'||PA02,PA01||'-'||PA02||'-'||PA03||'-'||PA04) AS CASENO,NVL(PA05,NVL(PA06,PA07)) PNAME," & _
                    " PA01 , PA02, PA03, PA04,PA75, PA26,PA08 " & _
                    " FROM CASEPROGRESS C1, PATENT " & _
                    " WHERE C1.CP01='FCP' AND C1.CP10='1228' " & stCn1 & _
                    "  AND C1.CP57 IS NULL AND C1.CP121 IS NULL AND C1.CP01=PA01(+) AND C1.CP02=PA02(+) AND C1.CP03=PA03(+) AND C1.CP04=PA04(+) " & _
                    " AND (PA01,PA02,PA03,PA04) IN (SELECT B.CP01,B.CP02,B.CP03,B.CP04 FROM CASEPROGRESS B WHERE B.CP01=PA01 AND B.CP02=PA02 AND B.CP03=PA03 AND B.CP04=PA04 AND CP10='926' AND CP159=0) "
      strSql = stCn2 & " ORDER BY PA14,PA08,PA01,PA02"
  '發email
  ElseIf strImportDList <> "" Then
        stCn2 = "SELECT C1.CP09,PA15,SQLDATET(PA14) AS PA14,PA11," & _
                    " DECODE(PA03||PA04,'000',PA01||'-'||PA02,PA01||'-'||PA02||'-'||PA03||'-'||PA04) AS CASENO,NVL(PA05,NVL(PA06,PA07)) PNAME," & _
                    " PA01 , PA02, PA03, PA04,PA75, PA26,DECODE(C1.CP121,'Y',1,0) CP121Y,DECODE(C1.CP121,'Y',0,1) CP121N,PA08,C1.CP158 " & _
                    " FROM CASEPROGRESS C1, PATENT " & _
                    " WHERE C1.CP01='FCP' AND C1.CP10='1228' " & stCn1 & _
                    "  AND C1.CP57 IS NULL AND C1.CP01=PA01(+) AND C1.CP02=PA02(+) AND C1.CP03=PA03(+) AND C1.CP04=PA04(+) " & _
                    " AND (PA01,PA02,PA03,PA04) IN (SELECT B.CP01,B.CP02,B.CP03,B.CP04 FROM CASEPROGRESS B WHERE B.CP01=PA01 AND B.CP02=PA02 AND B.CP03=PA03 AND B.CP04=PA04 AND CP10='926' AND CP159=0) "
        strSql = "SELECT '2' ord1,PA14,CP158,COUNT(CASENO) CNT,SUM(CP121Y) CP121Y,SUM(CP121N) CP121N FROM (" & stCn2 & ") GROUP BY PA14,CP158 "
        strSql = strSql & " ORDER BY 2,1"
  End If
  
  intJ = 1
  Set rsAD = ClsLawReadRstMsg(intJ, strSql)
  If intJ = 1 Then
        If bolRefresh = True Then
            Set MSGrd1.Recordset = rsAD
            LblCnt.Caption = "查詢，共 " & rsAD.RecordCount & " 筆"
            Call SetGrd(False)
            '記錄Grid的欄位
            If colPA11 = 0 Then
                colPA15 = PUB_MGridGetId("公 告 號", MSGrd1) '公告號
                colPA11 = PUB_MGridGetId("申請案號", MSGrd1) '申請案號
                colPA01 = PUB_MGridGetId("PA01", MSGrd1)   '本所案號(系統別)
            End If
        '發email
        ElseIf strImportDList <> "" Then
            If rsAD.RecordCount > 0 Then
                strExc(2) = "" 'Added by Lydia 2019/04/09
                With rsAD
                    .MoveFirst
                    Do While Not .EOF
                        '考慮不同期別的最後一筆,要逐期別判斷
                        '2019/03/12 若有先前公告本要重下的情況，請人員在email註明公告號，總務人員匯入後回覆email
                        If "" & .Fields("CP121Y") = "" & .Fields("CNT") Then
                            strDTlist = strDTlist & " OR C1.CP27=" & .Fields("CP158") '發文日和公告日可能不一致
                            strExc(2) = strExc(2) & "," & .Fields("PA14") 'Added by Lydia 2019/04/09
                        End If
                        .MoveNext
                    Loop
                End With
                If strDTlist <> "" Then '產生清單
                    strDTlist = "AND (" & Mid(strDTlist, 4) & " ) "
                    strSql = Replace(stCn2, stCn1, strDTlist)
                    strSql = strSql & " order by PA14, PA08,PA11 "
                    intJ = 1
                    Set rsAD = ClsLawReadRstMsg(intJ, strSql)
                    If intJ = 1 Then
                        '用Word編緝，最後轉存PDF檔，避免PDF Creater
                       Call WordList(rsAD)
                       
                       If strFileList <> "" Then
                            'Modified by Lydia 2019/04/09 主旨註明為匯入期別
                            'Modified by Lydia 2022/12/02 發文對象請改為 江如玉, Cc對象: 陳亭妙 , 李姿瑄
                            'Modified by Lydia 2023/03/07 收件者,CC改成系統特殊設定
                            PUB_SendMail strUserNum, Pub_GetSpecMan("外專程序-匯入公告本收件者"), "", "匯入" & Mid(strExc(2), 2) & "公告本", "請參考附件", , strFileList, , , , Pub_GetSpecMan("外專程序-匯入公告本副本")
                       End If
                    End If
                End If
            End If

        End If
  ElseIf bolRefresh = True Then
        ShowNoData
        LblCnt.Caption = "查詢，共  0  筆"
  End If
   
End Function

Private Function AutoUpdCpp() As Boolean
Dim intA As Integer
Dim fs, f
Dim strKey As String '公告號
Dim strKeyCP09 As String '公告公報收文號
Dim strFileName As String '檔案名稱
Dim strErrMsg As String '錯誤訊息
Dim strExSql As String '更新語法
Dim stReName As String

    strFileName = Dir(txtPath1.Text & "\*.pdf")
    If strFileName = "" Then Exit Function
    
    Do While strFileName <> ""
        '檢查檔案是否正在使用中
        If PUB_ChkFileOpening(txtPath1.Text & "\" & strFileName) = True Then
            MsgBox strFileName & vbCrLf & "檔案正在使用中，請關閉才可執行匯入！", vbExclamation
            strErrMsg = strErrMsg & strFileName & "：檔案正在使用中，請關閉才可執行匯入！" & vbCrLf
            GoTo JumpToNext
        End If
        '檢查檔案大小為 0 KB 有誤
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set f = fs.GetFile(txtPath1.Text & "\" & strFileName)
        If f.Size = 0 Then
           strErrMsg = strErrMsg & strFileName & "：檔案插入有誤，因檔案大小為 0 KB！" & vbCrLf
           GoTo JumpToNext
        End If
        '檔案名稱->抓公告號
        strKey = ""
        'Modified by Lydia2022/01/12 改成從檔名後端判斷；FCP-62288為改請衍生，所以下載檔名為TB001628620_GN__1_108307218D01_D216483，比一般申請多了D01
'        If InStr(UCase(strFileName), "I") > 0 Then '發明
'            intA = InStr(UCase(strFileName), "I")
'        ElseIf InStr(UCase(strFileName), "M") > 0 Then '新型
'            intA = InStr(UCase(strFileName), "M")
'        ElseIf InStr(UCase(strFileName), "D") > 0 Then '設明
'            intA = InStr(UCase(strFileName), "D")
        strExc(1) = Mid(UCase(strFileName), 1, Len(UCase(strFileName)) - 4)
        If InStr(strExc(1), "I") > 0 Then '發明
            intA = InStrRev(strExc(1), "I")
        ElseIf InStr(strExc(1), "M") > 0 Then '新型
            intA = InStrRev(strExc(1), "M")
        ElseIf InStr(strExc(1), "D") > 0 Then '設明
            intA = InStrRev(strExc(1), "D")
        'end 2022/01/12
        Else
        End If
        If intA > 0 Then
            strKey = Mid(strFileName, 1, Len(strFileName) - 4)
            strKey = Mid(strKey, intA)
        Else
            'strErrMsg = strErrMsg & strFileName & "：檔案名稱無公告號！" & vbCrLf 'Remove by Lydia 2019/05/17
            Kill txtPath1.Text & "\" & strFileName 'Added by Lydia 2019/05/17 因為總務人員無法理解錯誤訊息,所以改成直接刪除檔案;已與秀玲確認
            GoTo JumpToNext
        End If
        If strKey <> "" Then
            strExc(0) = " SELECT PA01,PA02,PA03,PA04,PA14,PA15,PA11,CP09,CP10,CP121,CP158,CPP02" & _
                             " FROM PATENT," & _
                             " (SELECT CP01,CP02,CP03,CP04,CP09,CP10,CP121,CP158 FROM CASEPROGRESS WHERE CP158>0 AND CP159=0 AND CP10='1228') V1," & _
                             " (SELECT CPP01,CPP02 FROM CASEPAPERPDF WHERE NVL(CPP10,'N') <> 'D' AND UPPER(CPP02) LIKE '%.GAZ.PDF') V2 " & _
                             " WHERE (PA15='" & strKey & "' OR PA13='" & strKey & "') " & _
                             " AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) AND CP09=CPP01(+) "
            intJ = 1
            Set RsTemp = ClsLawReadRstMsg(intJ, strExc(0))
            If intJ = 1 Then
                If "" & RsTemp.Fields("CP09") <> "" Then
                   If "" & RsTemp.Fields("PA01") = "FCP" Then
                        If "" & RsTemp.Fields("CP121") = "Y" Or "" & RsTemp.Fields("CPP02") <> "" Then
                            'strErrMsg = strErrMsg & strFileName & "：本所案號" & RsTemp.Fields("PA01") & "-" & RsTemp.Fields("PA02") & IIf(RsTemp.Fields("PA03") & RsTemp.Fields("PA04") <> "000", "-" & RsTemp.Fields("PA03") & "-" & RsTemp.Fields("PA04"), "") & _
                                      "，公告公報(" & RsTemp.Fields("CP09") & ")卷宗區已有公告本！" & vbCrLf 'Remove by Lydia 2019/05/17
                            Kill txtPath1.Text & "\" & strFileName 'Added by Lydia 2019/05/17
                            GoTo JumpToNext
                        Else
                            strKeyCP09 = "" & RsTemp.Fields("CP09")
                            '統一更名
                            If PUB_GetEmpFlowReNameFile(RsTemp.Fields("PA01"), RsTemp.Fields("PA02"), RsTemp.Fields("PA03"), RsTemp.Fields("PA04"), RsTemp.Fields("CP10"), RsTemp.Fields("PA01") & RsTemp.Fields("PA02") & "." & RsTemp.Fields("CP10") & ".pdf", stReName, True, 1, False, , , "GAZ") = False Then
                            End If
                            
                            If SaveAttFile_PDF(strKeyCP09, txtPath1.Text & "\" & strFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), False) = False Then
                                'strErrMsg = strErrMsg & strFileName & "：本所案號" & RsTemp.Fields("PA01") & "-" & RsTemp.Fields("PA02") & IIf(RsTemp.Fields("PA03") & RsTemp.Fields("PA04") <> "000", "-" & RsTemp.Fields("PA03") & "-" & RsTemp.Fields("PA04"), "") & _
                                          "，公告公報(" & RsTemp.Fields("CP09") & ")卷宗區上傳失敗！" & vbCrLf 'Remove by Lydia 2019/05/17
                                Kill txtPath1.Text & "\" & strFileName 'Added by Lydia 2019/05/17
                                GoTo JumpToNext
                            Else
                                strExSql = strExSql & "update caseprogress set cp121='Y' where cp09='" & strKeyCP09 & "' and cp121 is null ;"
                                If strImportList = "" Or (strImportList <> "" And InStr(strImportList, strKeyCP09) = 0) Then
                                    strImportList = strImportList & strKeyCP09 & ","
                                End If
                                If strImportDList = "" Or (strImportDList <> "" And InStr(strImportDList, "" & RsTemp.Fields("CP158")) = 0) Then
                                    strImportDList = strImportDList & "" & RsTemp.Fields("CP158") & ","
                                End If
                                Kill txtPath1.Text & "\" & strFileName
                            End If
                        End If
                   Else
                        'strErrMsg = strErrMsg & strFileName & "：本所案號" & RsTemp.Fields("PA01") & "-" & RsTemp.Fields("PA02") & IIf(RsTemp.Fields("PA03") & RsTemp.Fields("PA04") <> "000", "-" & RsTemp.Fields("PA03") & "-" & RsTemp.Fields("PA04"), "") & _
                                  "，非FCP案！" & vbCrLf 'Remove by Lydia 2019/05/17
                        Kill txtPath1.Text & "\" & strFileName 'Added by Lydia 2019/05/17
                        GoTo JumpToNext
                   End If
                Else
                     'strErrMsg = strErrMsg & strFileName & "：本所案號" & RsTemp.Fields("PA01") & "-" & RsTemp.Fields("PA02") & IIf(RsTemp.Fields("PA03") & RsTemp.Fields("PA04") <> "000", "-" & RsTemp.Fields("PA03") & "-" & RsTemp.Fields("PA04"), "") & _
                               "，未收文公告公報！" & vbCrLf 'Remove by Lydia 2019/05/17
                     Kill txtPath1.Text & "\" & strFileName 'Added by Lydia 2019/05/17
                     GoTo JumpToNext
                End If
            Else
                'strErrMsg = strErrMsg & strFileName & "：公告號查無案件基本資料！" & vbCrLf 'Remove by Lydia 2019/05/17
                Kill txtPath1.Text & "\" & strFileName 'Added by Lydia 2019/05/17
                GoTo JumpToNext
            End If
        End If
        
JumpToNext:
        strFileName = Dir()
    Loop
    
    '統一更新CP121
    If strExSql <> "" Then
        tmpArr = Empty
        tmpArr = Split(strExSql, ";")
        cnnConnection.BeginTrans
        For intJ = 0 To UBound(tmpArr)
            If Trim(tmpArr(intJ)) <> "" Then
                cnnConnection.Execute Trim(tmpArr(intJ)), intI
            End If
        Next intJ
        cnnConnection.CommitTrans
    End If
    
    If strErrMsg <> "" Then '錯誤訊息統一發email通知操作者
        MsgBox "請檢查公告本資料夾的檔案，錯誤訊息如下列：" & vbCrLf & strErrMsg, vbInformation, "公告本匯入錯誤訊息"  'Added by Lydia 2019/04/09 總務人員沒有一直開啟Outlook,所以加彈訊息
        PUB_SendMail strUserNum, strUserNum, "", "◎公告本匯入錯誤訊息", vbCrLf & strErrMsg
    End If
    
    Exit Function
    
End Function

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrMSGrd1HeadText, arrMSGrd1HeadWidth
   Dim iRow As Integer
   
   arrMSGrd1HeadText = Array("V", "公 告 號", "公 告 日", "申請案號", "本所 案 號", "案　件　名　稱", "PA01", "PA02", "PA03", "PA04", "PA75", "PA26")
   arrMSGrd1HeadWidth = Array(260, 900, 900, 1000, 1100, 3100, 0, 0, 0, 0, 0, 0)
   MSGrd1.Visible = False
   MSGrd1.Cols = UBound(arrMSGrd1HeadText) + 1
   If pReset = True Then
        MSGrd1.Clear
        MSGrd1.Rows = 2
   End If
   For iRow = 0 To MSGrd1.Cols - 1
      MSGrd1.row = 0
      MSGrd1.col = iRow
      MSGrd1.Text = arrMSGrd1HeadText(iRow)
      MSGrd1.ColWidth(iRow) = arrMSGrd1HeadWidth(iRow)
      MSGrd1.CellAlignment = flexAlignCenterCenter
   Next

   MSGrd1.Visible = True
End Sub

Private Sub MSGrd1_Click()
Dim strCopyTxt As String ' 複製編號文字

   MSGrd1.row = MSGrd1.MouseRow
   
   '選到編號欄=複製
   MSGrd1.col = MSGrd1.MouseCol
   If InStr("公 告 號,申請案號", MSGrd1.Text) = 0 Then
    If MSGrd1.col = colPA15 Or MSGrd1.col = colPA11 Then
         strCopyTxt = MSGrd1.TextMatrix(MSGrd1.row, MSGrd1.col)
         If strCopyTxt <> "" Then
             '複製編號至剪貼簿
             Clipboard.Clear
             Clipboard.SetText strCopyTxt
             MSGrd1.CellBackColor = QBColor(7)
             If MSGrd1.col = colPA15 Then
                 MsgBox strCopyTxt & "，公告號已複製", , MsgText(21)
             Else
               MsgBox strCopyTxt & "，申請案號已複製", , MsgText(21)
             End If
             
             '設回原本顏色
             MSGrd1.CellBackColor = QBColor(15)
         End If
         Exit Sub
    End If
   End If
   MSGrd1.Visible = False
   MSGrd1.col = 0
   If MSGrd1.row <> 0 Then
       If MSGrd1.Text = "V" Then
            MSGrd1.Text = ""
            MSGrd1.col = 1
            For intJ = 0 To MSGrd1.Cols - 1
               If intJ <> colPA15 And intJ <> colPA11 Then
                  MSGrd1.col = intJ
                  MSGrd1.CellBackColor = QBColor(15)
               End If
            Next intJ
            lblCnt2.Caption = Val(lblCnt2.Caption) - 1  'Added by Lydia 2019/06/12 勾選筆數
       Else
            MSGrd1.Text = "V"
            For intJ = 0 To MSGrd1.Cols - 1
               If intJ <> colPA15 And intJ <> colPA11 Then
                  MSGrd1.col = intJ
                  MSGrd1.CellBackColor = &HFFC0C0
               End If
            Next intJ
            lblCnt2.Caption = Val(lblCnt2.Caption) + 1  'Added by Lydia 2019/06/12 勾選筆數
       End If
   End If
   MSGrd1.Visible = True
End Sub

Private Sub MSGrd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow MSGrd1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   MSGrd1.col = nCol
   MSGrd1.row = nRow
   If Me.MSGrd1.row < 1 And Me.MSGrd1.Text <> "V" Then
      '全部都是文字(保留數值排序)
      'If InStr("公 告 日,申請案號", Me.MSGrd1.Text) > 0 Then
      '   If m_blnColOrderAsc = True Then
      '      Me.MSGrd1.Sort = 3  '數值昇冪
      '      m_blnColOrderAsc = False
      '   Else
      '      Me.MSGrd1.Sort = 4 '數值降冪
      '      m_blnColOrderAsc = True
      '   End If
      'Else
         If m_blnColOrderAsc = True Then
            Me.MSGrd1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.MSGrd1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      'End If
   End If
End Sub

'調卷清單
Private Sub WordList(ByRef mRs As ADODB.Recordset)
Dim intP As Integer
Dim strFileName As String
Dim strGrp As String

    mRs.MoveFirst
    
    Do While Not mRs.EOF
        If strGrp <> "" & mRs.Fields("PA14") Then
            If strGrp <> "" Then
                Pub_RePosWord g_WordAp, bVisible, m_WordLeft, m_WordTop 'Added by Lydia 2019/04/09 還原Word位置
                '先存檔
                'Modified by Lydia 2023/04/27 改模組
                'If PUB_PrintWord2PDF(g_WordAp, App.path, strFileName, strExc(1)) = True Then
                If PUB_PrintWord2File(g_WordAp, App.path, strFileName, strExc(1)) = True Then
                     strFileList = strFileList & App.path & "\" & strExc(1) & "*"
                End If
            End If
            'Modified by Lydia 2019/04/09
            'Call NewListPage("0", "" & mRs.Fields("PA14"))
            If NewListPage("0", "" & mRs.Fields("PA14")) = False Then GoTo ErrHandle
            intP = 1
        End If
        
        With g_WordAp.Application
            '流水號.申請案號(公告號)
            strExc(1) = Right(String(3, " ") & intP, 3) & "." & mRs.Fields("PA11") & "(" & mRs.Fields("PA15") & ")"
            .Selection.TypeText Text:=strExc(1)
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            '本所案號
            .Selection.TypeText Text:="" & mRs.Fields("CASENO")
            .Selection.MoveRight Unit:=wdCharacter, Count:=3
        End With
        
        intCounter = intCounter + 1
        If intCounter > iLimit Then
            'Modified by Lydia 2019/04/09
            'Call NewListPage("1", "" & mRs.Fields("PA14"))
            If NewListPage("1", "" & mRs.Fields("PA14")) = False Then GoTo ErrHandle
        End If
        strGrp = "" & mRs.Fields("PA14")
        strFileName = Replace("" & mRs.Fields("PA14"), "/", "") & "調卷清單"
        intP = intP + 1
        mRs.MoveNext
    Loop
    
    '存檔
    Pub_RePosWord g_WordAp, bVisible, m_WordLeft, m_WordTop 'Added by Lydia 2019/04/09 還原Word位置
    'Modified by Lydia 2023/04/27 改模組
    'If PUB_PrintWord2PDF(g_WordAp, App.path, strFileName, strExc(1)) = True Then
    If PUB_PrintWord2File(g_WordAp, App.path, strFileName, strExc(1)) = True Then
         strFileList = strFileList & App.path & "\" & strExc(1) & "*"
    End If

'Added by Lydia 2019/04/09
    Exit Sub
    
ErrHandle:
    MsgBox "調卷清單Word檔產生失敗，請通知電腦中心！", vbCritical
End Sub

'Modified by Lydia 2019/04/09
'Private Sub NewListPage(ByVal iType As String, ByVal kDate As String)
Private Function NewListPage(ByVal iType As String, ByVal kDate As String) As Boolean
    
    NewListPage = False 'Added by Lydia 2019/04/09
    '開啟Word檔
    If iType = "0" Then
          'Modifiec by Lydia 2019/04/09 改成模組
          'If TypeName(g_WordAp) <> "Application" Then Set g_WordAp = New Word.Application
          'g_WordAp.Documents.add
          If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = False Then Exit Function
          
          With g_WordAp.Application
              'Mark by Lydia 2019/04/09 改成模組
              '.Visible = True
              '.WindowState = wdWindowStateMaximize
              '.WindowState = wdWindowStateNormal
              '.ActiveWindow.ActivePane.View.Zoom.Percentage = 100
              'end 2019/04/09
             '邊界
             .Selection.PageSetup.LeftMargin = .CentimetersToPoints(1)
             .Selection.PageSetup.RightMargin = .CentimetersToPoints(1)
             .Selection.PageSetup.TopMargin = .CentimetersToPoints(1)
             .Selection.PageSetup.BottomMargin = .CentimetersToPoints(0.8)
          End With
          iPage = 0
    '跳頁
    Else
        g_WordAp.Selection.InsertBreak Type:=wdPageBreak
        g_WordAp.Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
    End If
      
    iPage = iPage + 1
    intCounter = 1
    '列印表頭
    With g_WordAp.Application
        
        '新增表格(1*3)
        If iPage = 1 Then
            .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=3
            With .Selection.Tables(1)
              .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
              .Borders(wdBorderRight).LineStyle = wdLineStyleNone
              .Borders(wdBorderTop).LineStyle = wdLineStyleNone
              .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
              .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
              .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
              .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
              .Borders.Shadow = False
            End With
            .Selection.SelectRow
            .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
            .Selection.Cells.SetHeight RowHeight:=16, HeightRule:=wdRowHeightExactly '固定列高
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(6), RulerStyle:=wdAdjustProportional
            .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(3), RulerStyle:=wdAdjustProportional
        Else
            .Selection.SelectRow
        End If
        
        .Selection.InsertRows iLimit
        .Selection.Collapse Direction:=wdCollapseStart
        .Selection.SelectRow
        .Selection.Cells.Merge
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.Cells.SetHeight RowHeight:=28, HeightRule:=wdRowHeightExactly '固定列高
        .Selection.Font.Size = 16
        .Selection.Font.Bold = True
        .Selection.TypeText Text:="公告本整批匯入-調卷清單"
        
        .Selection.MoveRight Unit:=wdCharacter, Count:=2
        
        .Selection.SelectRow
        .Selection.Font.Size = 12
        .Selection.Font.Bold = False
        
        .Selection.SelectRow
        .Selection.Collapse Direction:=wdCollapseStart
        .Selection.TypeText Text:="列印人員：" & strUserName
        .Selection.MoveRight Unit:=wdCharacter, Count:=2
        .Selection.TypeText Text:=String(8, "　") & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
        .Selection.MoveRight Unit:=wdCharacter, Count:=2
        .Selection.TypeText Text:="公告期別：" & kDate
        .Selection.MoveRight Unit:=wdCharacter, Count:=2
        .Selection.TypeText Text:=String(8, "　") & "頁　　數：" & iPage
        .Selection.SelectRow
        .Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle '用底部框線當做分隔線
        .Selection.Collapse Direction:=wdCollapseEnd
        intCounter = 4 '目前位置
    End With
    NewListPage = True 'Added by Lydia 2019/04/09
End Function

'Added by Lydia 2023/04/12
Private Sub txtDate_GotFocus()
   TextInverse txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
   If txtDate <> "" Then
      If CheckIsTaiwanDate(txtDate.Text) = False Then
         txtDate.SetFocus
         txtDate_GotFocus
         Cancel = True
         Exit Sub
      Else
         If txtDate > strSrvDate(2) Then
            MsgBox "公告日不可超過系統日!", vbOKOnly + vbExclamation
            txtDate.SetFocus
            txtDate_GotFocus
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub cmdEmail_Click()
Dim tmpBol As Boolean

   If Trim(txtDate) = "" Then
      MsgBox "公告日不可空白!", vbOKOnly + vbExclamation
      txtDate.SetFocus
      txtDate_GotFocus
      Exit Sub
   Else
      Call txtDate_Validate(tmpBol)
      If tmpBol = True Then
          Exit Sub
      End If
   End If
   
   strImportDList = DBDATE(txtDate)
   If QueryData(False) = True Then
       MsgBox "Email寄送完成!"
   End If
End Sub
'end 2023/04/12
