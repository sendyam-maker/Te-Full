VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm060124 
   BorderStyle     =   1  '單線固定
   Caption         =   "勘誤完備輸入"
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
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1200
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   157
      Width           =   2775
   End
   Begin VB.CommandButton CmdTxt 
      Caption         =   "全選"
      Height          =   300
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   680
   End
   Begin VB.CommandButton CmdTxt 
      Caption         =   "複製申請案號"
      Height          =   300
      Index           =   1
      Left            =   930
      TabIndex        =   8
      Top             =   960
      Width           =   1515
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7350
      Top             =   1050
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
      AllowUserResizing=   3
      FormatString    =   "V|勘 誤 日|申 請 案 號|本 所 案 號|事　　　　由|案　件　名　稱"
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
      Left            =   5700
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
      Left            =   4680
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
      Text            =   "C:\temp"
      Top             =   570
      Width           =   5745
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<="
      Height          =   315
      Left            =   7440
      TabIndex        =   0
      Top             =   570
      Width           =   345
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
   Begin VB.Label Label3 
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
   Begin VB.Label Label1 
      Caption         =   "承辦人員："
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
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   180
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "備註：可勾選多筆記錄後，按下複製按鈕即可複製多筆申請案號。"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
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
      Top             =   570
      Width           =   1695
   End
End
Attribute VB_Name = "frm060124"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/11/10 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB
'Create by Lydia 2019/05/23 勘誤完備輸入
Option Explicit
Dim rsAD As New ADODB.Recordset
Public cmdState As Integer '紀錄作用按鍵
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序

Dim intJ As Integer
Dim tmpArr As Variant

Dim colPA11 As Integer, colPA01 As Integer '記錄Grid的欄位

Private Sub Cmd1_Click(Index As Integer)

    Select Case Index
        Case 0 '查詢
             Screen.MousePointer = vbHourglass
             If QueryData(True) = False Then
             End If
             Screen.MousePointer = vbDefault
        Case 1 '匯入
             Screen.MousePointer = vbHourglass
             
             Cmd1(0).Enabled = False
             Cmd1(1).Enabled = False
             CmdTxt(1).Enabled = False
             CmdTxt(2).Enabled = False

             If AutoUpdCpp() = True Then
             End If
            '重整Grid
            If QueryData(True) = False Then
            End If
             
            Cmd1(0).Enabled = True
            Cmd1(1).Enabled = True
            CmdTxt(1).Enabled = True
            CmdTxt(2).Enabled = True
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
         If "" & MSGrd1.TextMatrix(iRow, colPA11) <> "" Then
            Select Case Index
                  Case 1 '複製申請號
                       If "" & MSGrd1.TextMatrix(iRow, 0) <> "" And "" & MSGrd1.TextMatrix(iRow, colPA11) <> "" Then
                           strCopyTxt = strCopyTxt & MSGrd1.TextMatrix(iRow, colPA11) & ";"
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
                    '底色統一為白色
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
        MsgBox "申請案號已複製(" & intP & ") ", , MsgText(21)
    ElseIf Index = 2 And intP > 0 Then
        If CmdTxt(Index).Caption = "全選" Then
            CmdTxt(Index).Caption = "取消"
            lblCnt2.Caption = intP  '勾選筆數
        Else
            CmdTxt(Index).Caption = "全選"
            lblCnt2.Caption = "0"  '勾選筆數
        End If
    End If
End Sub

Private Sub Combo1_Click()
    If Combo1.Tag <> "" And Combo1.Tag <> Combo1.Text Then
        Call Cmd1_Click(0)
    End If
    Combo1.Tag = Combo1.Text
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
  
   strExc(1) = GetSetting("TAIE", "FileCRC", UCase(Me.Name) & "Dir", "")
   If strExc(1) <> "" Then
      txtPath1.Text = strExc(1)
   '預設個人桌面
   Else
      txtPath1.Text = PUB_Getdesktop
   End If
   
   Call SetCombo1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060124 = Nothing
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

   If Combo1.Tag <> "" Then
      Combo1.ListIndex = 0
   Else
      Combo1.ListIndex = Combo1.ListCount - 1
   End If
   Combo1.Tag = Combo1.Text
   Call Cmd1_Click(0)
End Sub

Private Sub Command2_Click()
Dim sFile
   
On Error GoTo ErrHnd
   
   With CommonDialog1
      .CancelError = True
      .FileName = "*.pdf"
      .Filter = "PDF檔案 (*.pdf)|*.pdf"
      If GetSetting("TAIE", "FileCRC", UCase(Me.Name) & "Dir", "") <> "" Then
         .InitDir = GetSetting("TAIE", "FileCRC", UCase(Me.Name) & "Dir", "")
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
                SaveSetting "TAIE", "FileCRC", UCase(Me.Name) & "Dir", Left(.FileName, InStrRev(.FileName, "\") - 1)
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
Dim strCon As String

   If bolRefresh = True Then
      '清空及預設欄位值
      Call SetGrd(True)
      lblCnt2.Caption = "0"  '勾選筆數
   End If
   
   strCon = " AND C1.CP14='" & Trim(Left(Combo1, 6)) & "' "
   
   'Modified by Lydia 2023/08/25 +專利權延長415
   strSql = "SELECT '' AS V, SQLDATET(C1.CP48) AS C1CP48,PA11,C1.CP64 AS C1CP64,DECODE(PA03||PA04,'000',PA01||'-'||PA02,PA01||'-'||PA02||'-'||PA03||'-'||PA04) AS CASENO " & _
                ",NVL(PA05,NVL(PA06,PA07)) AS CASENAME,PA01,PA02,PA03,PA04 " & _
                "FROM CASEPROGRESS C1, CASEPROGRESS C2,PATENT " & _
                "WHERE C1.CP05>=20190501 AND C1.CP158=0 AND C1.CP159=0 AND SUBSTR(C1.CP09,1,1)='C' AND C1.CP10='1001' AND NVL(C1.CP121,'N')='N' " & strCon & _
                "AND C1.CP43=C2.CP09(+) AND C2.CP10 IN ('402','403','415') AND C1.CP01=PA01(+) AND C1.CP02=PA02(+) AND C1.CP03=PA03(+) AND C1.CP04=PA04(+) "
   strSql = strSql & " ORDER BY C1CP48,CASENO "

  intJ = 1
  Set rsAD = ClsLawReadRstMsg(intJ, strSql)
  If intJ = 1 Then
        If bolRefresh = True Then
            Set MSGrd1.Recordset = rsAD
            LblCnt.Caption = "查詢，共 " & rsAD.RecordCount & " 筆"
            Call SetGrd(False)
            '記錄Grid的欄位
            If colPA11 = 0 Then
                colPA11 = PUB_MGridGetId("申請案號", MSGrd1) '申請案號
                colPA01 = PUB_MGridGetId("PA01", MSGrd1)   '本所案號(系統別)
            End If
        End If
  ElseIf bolRefresh = True Then
        ShowNoData
        LblCnt.Caption = "查詢，共  0  筆"
  End If
   
End Function

Private Function AutoUpdCpp() As Boolean
Dim intA As Integer, RsUpd As New ADODB.Recordset
Dim fs, f
Dim strKey As String '申請號
Dim strKeyCP09 As String '核准收文號
Dim strFileName As String '檔案名稱
Dim strErrMsg As String '錯誤訊息
Dim strExSql As String '更新語法
Dim strMailList As String '發email的記錄
Dim tmpMail As Variant
Dim stReName As String
Dim strMemo As String 'Added by Lydia 2023/08/25
Dim strTo As String, strCC As String, strSubject As String, strContent As String 'Added by Lydia 2024/05/30
Dim strGrp As String 'Added by Lydia 2024/06/21

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
        'Added by Lydia 2024/06/21
        If Right(UCase(strFileName), 8) <> ".CRC.PDF" And Right(UCase(strFileName), 8) <> ".GAZ.PDF" Then
           strErrMsg = strErrMsg & strFileName & "：副檔名請預設為.CRC.PDF 或 .GAZ.PDF" & vbCrLf
           GoTo JumpToNext
        End If
        'end 2024/06/21
        '檔案名稱->抓申請號
        strKey = Mid(strFileName, 1, InStr(strFileName, ".") - 1)

        If strKey <> "" Then
            'Modified by Lydia 2023/08/25 +相關總收文號的案件性質
            'strExc(0) = " SELECT PA01,PA02,PA03,PA04,PA11,V1.CP09,V1.CP10,V1.CP121,V1.CP158,CPP02,C2.CP14 AS BCP14, ST03 AS BST03, ST02 AS BST02" & _
                             " FROM PATENT," & _
                             " (SELECT CP01,CP02,CP03,CP04,CP09,CP10,CP14,CP43,CP121,CP158 FROM CASEPROGRESS WHERE CP158=0 AND CP159=0 AND CP10='1001' AND SUBSTR(CP09,1,1)='C' ) V1," & _
                             " (SELECT CPP01,CPP02 FROM CASEPAPERPDF WHERE NVL(CPP10,'N') <> 'D' AND UPPER(CPP02) LIKE '%.CRC.PDF') V2" & _
                             " ,CASEPROGRESS C2, STAFF WHERE PA11='" & strKey & "' " & _
                             " AND PA01=V1.CP01(+) AND PA02=V1.CP02(+) AND PA03=V1.CP03(+) AND PA04=V1.CP04(+) AND V1.CP09=CPP01(+)" & _
                             " AND V1.CP43=C2.CP09(+) AND C2.CP14=ST01(+)"
            'Modified by Lydia 2024/06/21 +更正402核准的公告本GAZ
            'strExc(0) = " SELECT PA01,PA02,PA03,PA04,PA11,V1.CP09,V1.CP10,V1.CP121,V1.CP158,CPP02,C2.CP14 AS BCP14, ST03 AS BST03, ST02 AS BST02" & _
                             "  ,c2.cp09 as scp09,c2.cp10 as scp10,decode(c2.cp01,'P',cpm04,cpm03) as scp10name FROM PATENT," & _
                             " (SELECT CP01,CP02,CP03,CP04,CP09,CP10,CP14,CP43,CP121,CP158 FROM CASEPROGRESS WHERE CP158=0 AND CP159=0 AND CP10='1001' AND SUBSTR(CP09,1,1)='C' ) V1," & _
                             " (SELECT CPP01,CPP02 FROM CASEPAPERPDF WHERE NVL(CPP10,'N') <> 'D' AND UPPER(CPP02) LIKE '%.CRC.PDF') V2" & _
                             " ,CASEPROGRESS C2, STAFF,casepropertymap WHERE PA11='" & strKey & "' " & _
                             " AND PA01=V1.CP01(+) AND PA02=V1.CP02(+) AND PA03=V1.CP03(+) AND PA04=V1.CP04(+) AND V1.CP09=CPP01(+)" & _
                             " AND V1.CP43=C2.CP09(+) AND C2.CP14=ST01(+) and c2.cp01=cpm01(+) and c2.cp10=cpm02(+)"
            strExc(0) = " SELECT PA01,PA02,PA03,PA04,PA11,V1.CP09,V1.CP10,V1.CP121,V1.CP158,CPP02,G02,C2.CP14 AS BCP14, ST03 AS BST03, ST02 AS BST02" & _
                             "  ,c2.cp09 as scp09,c2.cp10 as scp10,decode(c2.cp01,'P',cpm04,cpm03) as scp10name FROM PATENT," & _
                             " (SELECT CP01,CP02,CP03,CP04,CP09,CP10,CP14,CP43,CP121,CP158 FROM CASEPROGRESS WHERE CP158=0 AND CP159=0 AND CP10='1001' AND SUBSTR(CP09,1,1)='C' ) V1," & _
                             " (SELECT CPP01,CPP02 FROM CASEPAPERPDF WHERE NVL(CPP10,'N') <> 'D' AND UPPER(CPP02) LIKE '%.CRC.PDF') V2," & _
                             " (SELECT CPP01 as G01,CPP02 as G02 FROM CASEPAPERPDF WHERE NVL(CPP10,'N') <> 'D' AND UPPER(CPP02) LIKE '%.GAZ.PDF') V3" & _
                             " ,CASEPROGRESS C2, STAFF,casepropertymap WHERE PA11='" & strKey & "' " & _
                             " AND PA01=V1.CP01(+) AND PA02=V1.CP02(+) AND PA03=V1.CP03(+) AND PA04=V1.CP04(+) AND V1.CP09=CPP01(+) AND V1.CP09=G01(+)" & _
                             " AND V1.CP43=C2.CP09(+) AND C2.CP14=ST01(+) and c2.cp01=cpm01(+) and c2.cp10=cpm02(+)"
            intJ = 1
            Set RsUpd = ClsLawReadRstMsg(intJ, strExc(0))
            If intJ = 1 Then
                If "" & RsUpd.Fields("CP09") <> "" Then
                   If "" & RsUpd.Fields("PA01") = "FCP" Then
                        'Added by Lydia 2024/06/21 區分種類
                        If Right(UCase(strFileName), 8) = ".CRC.PDF" Then
                           strGrp = "CRC"
                        ElseIf Right(UCase(strFileName), 8) = ".GAZ.PDF" Then
                           strGrp = "GAZ"
                        End If
                        'Modified by Lydia 2024/06/21
                        'If "" & RsUpd.Fields("CP121") = "Y" Or "" & RsUpd.Fields("CPP02") <> "" Then
                        If "" & RsUpd.Fields("CPP02") <> "" And strGrp = "CRC" Then
                            strErrMsg = strErrMsg & strFileName & "：本所案號" & RsUpd.Fields("PA01") & "-" & RsUpd.Fields("PA02") & IIf(RsUpd.Fields("PA03") & RsUpd.Fields("PA04") <> "000", "-" & RsUpd.Fields("PA03") & "-" & RsUpd.Fields("PA04"), "") & _
                                      "，核准(" & RsUpd.Fields("CP09") & ")卷宗區已有勘誤表！" & vbCrLf
                            GoTo JumpToNext
                        'Added by Lydia 2024/06/21
                        ElseIf "" & RsUpd.Fields("G02") <> "" And strGrp = "GAZ" Then
                            strErrMsg = strErrMsg & strFileName & "：本所案號" & RsUpd.Fields("PA01") & "-" & RsUpd.Fields("PA02") & IIf(RsUpd.Fields("PA03") & RsUpd.Fields("PA04") <> "000", "-" & RsUpd.Fields("PA03") & "-" & RsUpd.Fields("PA04"), "") & _
                                      "，核准(" & RsUpd.Fields("CP09") & ")卷宗區已有公告本！" & vbCrLf
                            GoTo JumpToNext
                        'end 2024/06/21
                        Else
                            strKeyCP09 = "" & RsUpd.Fields("CP09")
                            'strMemo = "(" & rsupd.Fields("scp10name") & ")"  'Added by Lydia 2023/08/25 專利權延長415加註 'Mark by Lydia 2024/05/30
                            '統一更名
                            'Modified by Lydia 2024/06/21 "CRC"=>strGrp
                            If PUB_GetEmpFlowReNameFile(RsUpd.Fields("PA01"), RsUpd.Fields("PA02"), RsUpd.Fields("PA03"), RsUpd.Fields("PA04"), RsUpd.Fields("CP10"), RsUpd.Fields("PA01") & RsUpd.Fields("PA02") & "." & RsUpd.Fields("CP10") & ".pdf", stReName, True, 1, False, , , strGrp) = False Then
                            End If
                            
                            If SaveAttFile_PDF(strKeyCP09, txtPath1.Text & "\" & strFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), False) = False Then
                                strErrMsg = strErrMsg & strFileName & "：本所案號" & RsUpd.Fields("PA01") & "-" & RsUpd.Fields("PA02") & IIf(RsUpd.Fields("PA03") & RsUpd.Fields("PA04") <> "000", "-" & RsUpd.Fields("PA03") & "-" & RsUpd.Fields("PA04"), "") & _
                                          "，核准(" & RsUpd.Fields("CP09") & ")卷宗區上傳失敗！" & vbCrLf
                                GoTo JumpToNext
                            Else
                                strExSql = strExSql & "update caseprogress set cp121='Y' where cp09='" & strKeyCP09 & "' and cp121 is null;"
                                'Modified by Lydia 2024/05/30
'                                '上發文日
'                                'Modified by Lydia 2019/06/19 如果操作者和承辦人非同一人,則更新承辦人
'                                'strExSql = strExSql & "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & strKeyCP09 & "' and cp27 is null ;"
'                                strExSql = strExSql & "update caseprogress set cp27=" & strSrvDate(1) & IIf(Trim(Left(Combo1.Text, 6)) <> strUserNum, ", cp14='" & strUserNum & "' ", "") & " where cp09='" & strKeyCP09 & "' and cp27 is null ;"
'                                'Kill txtPath1.Text & "\" & strFileName 　'發完email再刪檔
'                                '收件者
'                                strMailList = strMailList & PUB_GetFCPHandler(rsupd.Fields("pa01"), rsupd.Fields("pa02"), rsupd.Fields("pa03"), rsupd.Fields("pa04")) & ";"
'                                '主旨
'                                'Modified by Lydia 2023/08/25 專利權延長415加註+strMemo
'                                strMailList = strMailList & rsupd.Fields("PA01") & "-" & rsupd.Fields("PA02") & IIf(rsupd.Fields("PA03") & rsupd.Fields("PA04") = "000", "", "-" & rsupd.Fields("PA03") & "-" & rsupd.Fields("PA04")) & "已下載勘誤表" & strMemo & ";"
'                                '附件名稱
'                                strMailList = strMailList & strFileName & ";"
'                                '內文: 請調卷後交 承辦業務姓名(第1項承辦人為程序) 或工程師姓名 "(第1項承辦人為工程師)通知客戶，勘誤表已匯入卷宗區。
'                                strExc(1) = ""
'                                If "" & rsupd.Fields("bst03") = "F21" Then
'                                    strExc(1) = "" & rsupd.Fields("bst02")
'                                    If InStr(strExc(1), "林信昌") > 0 Then '因為"林信昌"有不同組別
'                                        strExc(1) = "林信昌"
'                                    End If
'                                Else
'                                    strExc(1) = GetStaffName(PUB_GetFCPSalesNo(rsupd.Fields("pa01"), rsupd.Fields("pa02"), rsupd.Fields("pa03"), rsupd.Fields("pa04")))
'                                End If
'                                'Modified by Lydia 2023/08/25 專利權延長415加註+strMemo
'                                strMailList = strMailList & "請調卷後交 " & strExc(1) & " 通知客戶，勘誤表" & strMemo & "已匯入卷宗區。"
'                                strMailList = strMailList & "" & "||"
'------------------------------------------------------------
                                '預設CC= 該區程序, 操作本人, backup
                                strCC = PUB_GetFCPHandler(RsUpd.Fields("pa01"), RsUpd.Fields("pa02"), RsUpd.Fields("pa03"), RsUpd.Fields("pa04"))
                                If strCC <> strUserNum Then strCC = strCC & ";" & strUserNum
                                strCC = strCC & ";backup"
                                '依內部收文的承辦人區分
                                'Modified by Lydia 2024/06/21 公告本通知工程師
                                'If "" & RsUpd.Fields("bst03") = "F21" Then
                                If "" & RsUpd.Fields("bst03") = "F21" Or strGrp = "GAZ" Then
                                    '承辦人為工程師：不上發文日，承辦人更新為工程師，並自動掛本所期限為+1週，承辦期限往前2天，
                                    '通知E-MAIL: 工程師、C.C.：工程師主管, 該區程序, 操作本人, backup
                                    If "" & RsUpd.Fields("bst03") = "F21" Then
                                       strTo = "" & RsUpd.Fields("bcp14")
                                    Else
                                       strTo = PUB_GetFCPPromoterNo(strKeyCP09, "1001")
                                    End If
                                    strExc(2) = PUB_GetFCPEngSup(strTo)
                                     
                                    strExc(5) = PUB_GetWorkDay1(CompDate(2, 7, strSrvDate(1)), True)
                                    strExc(6) = CompWorkDay(-3, strExc(5))
                                     
                                    strExSql = strExSql & "update caseprogress set cp14='" & strTo & "', cp06=" & strExc(5) & ", cp48=" & strExc(6) & " where cp09='" & strKeyCP09 & "' and cp27 is null ;"
                                     
                                    If strExc(2) <> "" Then strCC = strExc(2) & ";" & strCC
                                    
                                    'Modified by Lydia 2024/06/25 加註[INCOM
                                    'strSubject = "Our Ref: " & RsUpd.Fields("PA01") & "-" & RsUpd.Fields("PA02") & IIf(RsUpd.Fields("PA03") & RsUpd.Fields("PA04") = "000", "", "-" & RsUpd.Fields("PA03") & "-" & RsUpd.Fields("PA04")) & RsUpd.Fields("scp10name") & "已核准"
                                    strSubject = RsUpd.Fields("PA01") & "-" & RsUpd.Fields("PA02") & IIf(RsUpd.Fields("PA03") & RsUpd.Fields("PA04") = "000", "", "-" & RsUpd.Fields("PA03") & "-" & RsUpd.Fields("PA04")) & RsUpd.Fields("scp10name") & "已核准 Our Ref: " & RsUpd.Fields("PA01") & "-" & RsUpd.Fields("PA02") & IIf(RsUpd.Fields("PA03") & RsUpd.Fields("PA04") = "000", "", "-" & RsUpd.Fields("PA03") & "-" & RsUpd.Fields("PA04")) & " [INCOM." & RsUpd.Fields("cp10") & "]"
                                    'Modified by Lydia 2024/06/21
                                    'strContent = RsUpd.Fields("scp10name") & "已核准，勘誤表已下載匯入卷宗區" & vbCrLf
                                    'strContent = strContent & "1. 主管請分案。" & vbCrLf
                                    'strContent = strContent & "2. 工程師請報告客戶及寄勘誤表。" & vbCrLf
                                    strContent = RsUpd.Fields("scp10name") & "已核准，" & IIf(strGrp = "CRC", "勘誤表", "公告本") & "已下載匯入卷宗區，請報告客戶及寄" & IIf(strGrp = "CRC", "勘誤表", "公告本") & "。" & vbCrLf
                                    'end 2024/06/21
                                Else
                                     '非工程師：同時上發文日。通知E-MAIL: 承辦、C.C.：承辦主管, 該區程序, 操作本人, backup
                                    strExSql = strExSql & "update caseprogress set cp27=" & strSrvDate(1) & IIf(Trim(Left(Combo1.Text, 6)) <> strUserNum, ", cp14='" & strUserNum & "' ", "") & " where cp09='" & strKeyCP09 & "' and cp27 is null ;"
                                     
                                    strTo = PUB_GetFCPSalesNo(RsUpd.Fields("pa01"), RsUpd.Fields("pa02"), RsUpd.Fields("pa03"), RsUpd.Fields("pa04"))
                                    strExc(2) = PUB_GetFCPProSup(strTo)
                                    If strExc(2) <> "" Then strCC = strExc(2) & ";" & strCC
                                    'Modified by Lydia 2024/06/25 加註[INCOM
                                    'strSubject = "Our Ref: " & RsUpd.Fields("PA01") & "-" & RsUpd.Fields("PA02") & IIf(RsUpd.Fields("PA03") & RsUpd.Fields("PA04") = "000", "", "-" & RsUpd.Fields("PA03") & "-" & RsUpd.Fields("PA04")) & RsUpd.Fields("scp10name") & "公報已核准"
                                    strSubject = RsUpd.Fields("PA01") & "-" & RsUpd.Fields("PA02") & IIf(RsUpd.Fields("PA03") & RsUpd.Fields("PA04") = "000", "", "-" & RsUpd.Fields("PA03") & "-" & RsUpd.Fields("PA04")) & RsUpd.Fields("scp10name") & "公報已核准 Our Ref: " & RsUpd.Fields("PA01") & "-" & RsUpd.Fields("PA02") & IIf(RsUpd.Fields("PA03") & RsUpd.Fields("PA04") = "000", "", "-" & RsUpd.Fields("PA03") & "-" & RsUpd.Fields("PA04")) & " [INCOM." & RsUpd.Fields("cp10") & "]"
                                    strContent = RsUpd.Fields("scp10name") & "公報已核准，勘誤表已下載匯入卷宗區，請報告客戶及寄勘誤表。" & vbCrLf
                                End If
                                '排列:收件者|主旨|附件名稱|內文|CC，用@區隔
                                'Modified by Lydia 2024/06/21 判斷收件者為工程師，不夾帶附件
                                'strMailList = strMailList & strTo & "|" & strSubject & "|" & strFileName & "|" & strContent & "|" & strCC & "@"
                                strMailList = strMailList & strTo & "|" & strSubject & "|" & IIf("" & RsUpd.Fields("bst03") = "F21" Or strGrp = "GAZ", "AAAA" & strFileName, strFileName) & "|" & strContent & "|" & strCC & "@"
                                'end 2024/05/30
                            End If  '----SaveAttFile_PDF
                        End If
                   Else
                        strErrMsg = strErrMsg & strFileName & "：本所案號" & RsUpd.Fields("PA01") & "-" & RsUpd.Fields("PA02") & IIf(RsUpd.Fields("PA03") & RsUpd.Fields("PA04") <> "000", "-" & RsUpd.Fields("PA03") & "-" & RsUpd.Fields("PA04"), "") & _
                                  "，非FCP案！" & vbCrLf
                        Kill txtPath1.Text & "\" & strFileName
                        GoTo JumpToNext
                   End If
                Else
                     strErrMsg = strErrMsg & strFileName & "：本所案號" & RsUpd.Fields("PA01") & "-" & RsUpd.Fields("PA02") & IIf(RsUpd.Fields("PA03") & RsUpd.Fields("PA04") <> "000", "-" & RsUpd.Fields("PA03") & "-" & RsUpd.Fields("PA04"), "") & _
                               "，無未發文的核准進度！" & vbCrLf
                     GoTo JumpToNext
                End If
            Else
                strErrMsg = strErrMsg & strFileName & "：申請號查無案件基本資料！" & vbCrLf
                GoTo JumpToNext
            End If
        End If
        
JumpToNext:
        strFileName = Dir()
    Loop
    
    '統一更新CP121,CP27
    If strExSql <> "" Then
        tmpArr = Empty
        tmpArr = Split(strExSql, ";")
        cnnConnection.BeginTrans
        For intJ = 0 To UBound(tmpArr)
            If Trim(tmpArr(intJ)) <> "" Then
                cnnConnection.Execute Trim(tmpArr(intJ)), intA
            End If
        Next intJ
        cnnConnection.CommitTrans
        '整批發email
        tmpArr = Empty
        'Modified by Lydia 2024/05/30 用||區隔 => 用@區隔
        tmpArr = Split(strMailList, "@")
        For intJ = 0 To UBound(tmpArr)
            If Trim(tmpArr(intJ)) <> "" Then
                tmpMail = Empty
                'Modified by Lydia 2024/05/30
                'tmpMail = Split(tmpArr(intJ), ";")
                'PUB_SendMail strUserNum, tmpMail(0), "", tmpMail(1), tmpMail(3), , txtPath1.Text & "\" & tmpMail(2)
                tmpMail = Split(tmpArr(intJ), "|")
                'Modified by Lydia 2024/06/21 判斷不夾帶附件
                'PUB_SendMail strUserNum, tmpMail(0), "", tmpMail(1), tmpMail(3), , txtPath1.Text & "\" & tmpMail(2), , , , tmpMail(4)
                PUB_SendMail strUserNum, tmpMail(0), "", tmpMail(1), tmpMail(3), , IIf(Left(tmpMail(2), 4) = "AAAA", "", txtPath1.Text & "\" & tmpMail(2)), , , , tmpMail(4)
                'end 2024/06/30
                '刪檔
                Sleep 1000
                Kill txtPath1.Text & "\" & IIf(Left(tmpMail(2), 4) = "AAAA", Mid(tmpMail(2), 5), tmpMail(2))
            End If
        Next intJ
    End If
    
    If strErrMsg <> "" Then '錯誤訊息統一發email通知操作者
        PUB_SendMail strUserNum, strUserNum, "", "◎勘誤表匯入錯誤訊息", vbCrLf & strErrMsg
    End If
    
    Set RsUpd = Nothing
    Exit Function
    
End Function

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrMSGrd1HeadText, arrMSGrd1HeadWidth
   Dim iRow As Integer
   
   arrMSGrd1HeadText = Array("V", "勘 誤 日", "申請案號", "事　　　　由", "本所 案 號", "案　件　名　稱", "PA01", "PA02", "PA03", "PA04")
   arrMSGrd1HeadWidth = Array(260, 860, 1000, 3200, 1100, 1500, 0, 0, 0, 0)
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
   If MSGrd1.Text <> "申請案號" Then
    If MSGrd1.col = colPA11 Then
         strCopyTxt = MSGrd1.TextMatrix(MSGrd1.row, MSGrd1.col)
         If strCopyTxt <> "" Then
             '複製編號至剪貼簿
             Clipboard.Clear
             Clipboard.SetText strCopyTxt
             MSGrd1.CellBackColor = QBColor(7)
             MsgBox strCopyTxt & "，申請案號已複製", , MsgText(21)
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
               If intJ <> colPA11 Then
                  MSGrd1.col = intJ
                  MSGrd1.CellBackColor = QBColor(15)
               End If
            Next intJ
            lblCnt2.Caption = Val(lblCnt2.Caption) - 1  '勾選筆數
       Else
            MSGrd1.Text = "V"
            For intJ = 0 To MSGrd1.Cols - 1
               If intJ <> colPA11 Then
                  MSGrd1.col = intJ
                  MSGrd1.CellBackColor = &HFFC0C0
               End If
            Next intJ
            lblCnt2.Caption = Val(lblCnt2.Caption) + 1 '勾選筆數
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
