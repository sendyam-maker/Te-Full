VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm040120 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人來函匯入"
   ClientHeight    =   5736
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   8952
   Begin VB.TextBox textUser 
      Height          =   285
      Left            =   1395
      MaxLength       =   6
      TabIndex        =   0
      Top             =   30
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "無缺檔"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   0
      Left            =   6120
      TabIndex        =   9
      Top             =   1200
      Width           =   870
   End
   Begin VB.TextBox textDate 
      Height          =   285
      Left            =   1395
      MaxLength       =   7
      TabIndex        =   1
      Top             =   450
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   7980
      TabIndex        =   7
      Top             =   30
      Width           =   885
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "匯入(&T)"
      Height          =   345
      Left            =   5175
      TabIndex        =   4
      Top             =   30
      Width           =   885
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   345
      Left            =   7065
      TabIndex        =   6
      Top             =   30
      Width           =   885
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "重整(&R)"
      Default         =   -1  'True
      Height          =   345
      Left            =   6120
      TabIndex        =   5
      Top             =   30
      Width           =   885
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   765
      TabIndex        =   21
      Top             =   5400
      Width           =   4995
   End
   Begin VB.CommandButton cmdPath 
      Height          =   330
      Left            =   8550
      Picture         =   "frm040120.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   840
      Width           =   350
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "開啟"
      Height          =   345
      Left            =   2160
      TabIndex        =   8
      Top             =   1200
      Width           =   705
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "卷宗區"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   13
      Left            =   7065
      TabIndex        =   10
      Top             =   1200
      Width           =   870
   End
   Begin VB.Frame Frame3 
      Caption         =   "待匯入檔案："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3705
      Left            =   30
      TabIndex        =   18
      Top             =   1290
      Width           =   3960
      Begin VB.FileListBox File1 
         Height          =   1692
         Left            =   1395
         TabIndex        =   19
         Top             =   1260
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.CheckBox Check2 
         Caption         =   "列印"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   11
         Top             =   0
         Width           =   705
      End
      Begin VB.ListBox lstImport 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3312
         ItemData        =   "frm040120.frx":0102
         Left            =   90
         List            =   "frm040120.frx":0109
         TabIndex        =   23
         Top             =   270
         Width           =   3795
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "待匯入案件："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3705
      Left            =   3990
      TabIndex        =   16
      Top             =   1290
      Width           =   4890
      Begin VB.CheckBox Check1 
         Caption         =   "列印"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   12
         Top             =   0
         Width           =   705
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3345
         Left            =   60
         TabIndex        =   13
         Top             =   270
         Width           =   4770
         _ExtentX        =   8424
         _ExtentY        =   5906
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   "V|本所案號|案件性質|收文日|缺檔|說明"
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   1410
      TabIndex        =   2
      Top             =   840
      Width           =   7110
   End
   Begin VB.Frame Frame1 
      Height          =   465
      Left            =   30
      TabIndex        =   15
      Top             =   4890
      Width           =   8895
      Begin VB.TextBox txtProgressBar 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00FF0000&
         Enabled         =   0   'False
         Height          =   300
         Left            =   45
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   120
         Width           =   8820
      End
   End
   Begin VB.Label Label4 
      Caption         =   "　　　　　本所案號.案件性質.副檔名.PDF（ex.P105116.1701.inv.PDF）"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   2730
      TabIndex        =   27
      Top             =   630
      Width           =   5805
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      Caption         =   "lblUserName"
      Height          =   180
      Left            =   2730
      TabIndex        =   26
      Top             =   75
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "來函輸入人員："
      Height          =   180
      Left            =   90
      TabIndex        =   25
      Top             =   75
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "來函輸入日期："
      Height          =   180
      Left            =   90
      TabIndex        =   24
      Top             =   495
      Width           =   1260
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      Height          =   180
      Left            =   45
      TabIndex        =   22
      Top             =   5460
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "檔名規則：本所案號.副檔名.PDF（ex.P105116.inv.PDF）"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   2730
      TabIndex        =   17
      Top             =   390
      Width           =   4455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "檔案存放路徑："
      Height          =   180
      Left            =   90
      TabIndex        =   14
      Top             =   900
      Width           =   1260
   End
End
Attribute VB_Name = "frm040120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/11/30 Form2.0已修改 (無需修改)
'Create By Morgan 2016/6/2
Option Explicit

Public cmdState As Integer '紀錄作用按鍵
Public m_ProState As String 'Add by Amy 2020/01/09
Dim m_AttachPath As String

'列印報表用---
Dim PLeft() As Integer
Dim strTemp() As String
Dim iNowLine As Integer
Dim iRowHeight As Integer

Dim strPrinter As String
Dim dblMaxWidth As Double
Dim oFileSys As New FileSystemObject
Dim oFile As File

Dim m_iCols As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim intPrevRow As Integer
Dim mstrSTARTFOLDER As String '起始資料夾

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdImPort_Click()
   If ImportFile Then
      QueryData
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
cmdState = Index '紀錄作用按鍵
PubShowNextData
End Sub

Private Function ImportFile() As Boolean
   Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, strCP09 As String, strCP10 As String, strErr As String
   Dim iCP01 As Integer, iCP02 As Integer, iCP03 As Integer, iCP04 As Integer, iCp09 As Integer, iCP10 As Integer
   Dim strCaseNo As String
   Dim iTotRows As Integer
   Dim ii As Integer, jj As Integer
   Dim dblFCnt As Double
   Dim stSaveName As String '存檔的檔名-本所案號.案件性質.INVOICE.Pdf
   Dim bolUploadDone As Boolean
   Dim bolChkOk As Boolean, iCaseQty As Integer 'Add by Amy 2020/02/26 可上傳/本所案號數量Col
   'Add by Amy 2020/02/27
   Dim iCP10N As Integer, iRCP10 As Integer  '案件性質名稱 Col/相關案件性質Col
   Dim bolUseRCP10 As Boolean, strRCP10 As String '使用相關案件性質/相關案件性質
   
On Error GoTo ErrHnd

   If IsEmptyText(txtPath) = True Then
      MsgBox "請選擇檔案存放路徑！", vbOKOnly, "檢核資料"
      If cmdPath.Enabled Then cmdPath.SetFocus
      Exit Function
   'Modified by Morgan 2017/1/12
   'ElseIf oFileSys.FolderExists(txtPath) = False Then
   ElseIf PUB_ChkDir(txtPath) = False Then
      MsgBox "檔案存放路徑不存在，請重新選擇！"
      If cmdPath.Enabled Then cmdPath.SetFocus
      Exit Function
   ElseIf Dir(txtPath & "\*.pdf") = "" Then
      MsgBox txtPath.Text & " 資料夾內沒有pdf檔！"
      If cmdPath.Enabled Then cmdPath.SetFocus
      Exit Function
   ElseIf textDate = "" Then
      MsgBox "來函輸入日期不可空白！", vbExclamation
      If textDate.Enabled Then textDate.SetFocus
      Exit Function
   End If
   
   RefreshList
   
   If MSHFlexGrid1.TextMatrix(1, 1) <> "" Then
      iCP01 = PUB_MGridGetId("cp01", MSHFlexGrid1)
      iCP02 = PUB_MGridGetId("cp02", MSHFlexGrid1)
      iCP03 = PUB_MGridGetId("cp03", MSHFlexGrid1)
      iCP04 = PUB_MGridGetId("cp04", MSHFlexGrid1)
      iCp09 = PUB_MGridGetId("cp09", MSHFlexGrid1)
      iCP10 = PUB_MGridGetId("cp10", MSHFlexGrid1)
      iCaseQty = PUB_MGridGetId("CaseQty", MSHFlexGrid1) 'Add by Amy 2020/02/26
      'Add by Amy 2020/02/27
      iCP10N = PUB_MGridGetId("案件性質", MSHFlexGrid1)
      iRCP10 = PUB_MGridGetId("RCP10", MSHFlexGrid1)
      'end 2020/02/27
   End If
   
   txtProgressBar.Width = 0
   dblFCnt = lstImport.ListCount
   For ii = 0 To lstImport.ListCount - 1
      'Modfiy by Amy 2020/01/09 改用InputCaseGetSys
'      If Left(lstImport.List(ii), 3) = "CFP" Then
'         strCP01 = "CFP"
'      ElseIf Left(lstImport.List(ii), 3) = "CPS" Then
'         strCP01 = "CPS"
'      ElseIf Left(lstImport.List(ii), 2) = "PS" Then
'         strCP01 = "PS"
'      Else
'         strCP01 = "P"
'      End If
      strCP01 = InputCaseGetSys(lstImport.List(ii))
      strErr = "": strCP02 = "": strCP03 = "": strCP04 = "": strCP09 = "": strCP10 = ""
      bolChkOk = False 'Add by Amy 2020/02/26
      bolUseRCP10 = False: strRCP10 = "" 'Add by Amy 2020/02/27
         
      'Added by Morgan 2019/1/21
      '檔名中不可有中文字,否則會無法合併 Ex:CFP-29926
      For jj = 1 To Len(lstImport.List(ii))
         If Asc(Mid(lstImport.List(ii), jj, 1)) <= 0 Then
            strErr = "檔名不可有中文字!!!"
            Exit For
         End If
      Next jj
      'Add By Sindy 2019/6/13
      If InStr(lstImport.List(ii), "#") > 0 Then
         strErr = IIf(strErr <> "", strErr & ";", "") & "【#】符號為系統保留字，請重新命名！"
      End If
      If InStr(lstImport.List(ii), ",") > 0 Then
         strErr = IIf(strErr <> "", strErr & "; ", "") & "逗號[,]為系統保留字，請重新命名！"
      End If
      '2019/6/13 END
      'Added by Lydia 2021/09/06 增加對檔案和檔名的控制
        '9/5每日批次"更正FTP檔案路徑"無法變更CA9056362.CFP028736.1909.altr.pdf；
        '1.在109/9/14已提申之卷宗區多出一筆.altr .pdf，造成無法更名；
        '2.因為無法確定形成原因，所以只針對檔名做取消空白及換行符號的動作。
      '檢查檔案是否正在使用中
      If PUB_ChkFileOpening(txtPath.Text & "\" & lstImport.List(ii)) = True Then
         strErr = IIf(strErr <> "", strErr & "; ", "") & Trim(lstImport.List(ii)) & vbCrLf & "檔案正在使用中，請關閉才可執行匯入！"
      End If
      '檢查檔名
      If Trim(lstImport.List(ii)) <> PUB_GetSimpleName(Trim(lstImport.List(ii))) Then
         strErr = IIf(strErr <> "", strErr & "; ", "") & Trim(lstImport.List(ii)) & "，不符檔案命名原則：含非英數字。"
      End If
      'end 2021/09/06
      
      If strErr = "" Then
      'end 2019/1/21
         If PUB_GetCaseNoFromFileName(lstImport.List(ii), strCP01, strCP02, strCP03, strCP04, strErr) = True Then
            '檔案要上傳的收文號順序
            '1.有缺檔的 2.後發文的AB類
            strCP09 = ""
            With MSHFlexGrid1
            For jj = 1 To .Rows - 1
               If .TextMatrix(jj, iCP01) = strCP01 And .TextMatrix(jj, iCP02) = strCP02 And .TextMatrix(jj, iCP03) = strCP03 And .TextMatrix(jj, iCP04) = strCP04 Then
                  strCP09 = .TextMatrix(jj, iCp09)
                  strCP10 = .TextMatrix(jj, iCP10)
                  strRCP10 = .TextMatrix(jj, iRCP10)
                  'Modify by Amy 2020/02/26
                  '右方List中案號只有一筆
                  If Val(.TextMatrix(jj, iCaseQty)) = 1 Then
                        '若檔名先都輸.1001.->改一筆為.102.匯入一筆->另一筆再改.301. 若先判斷只有一筆會將.301.檔名也串入
                        If InStr(.TextMatrix(jj, iCP10N), "-") > 0 And InStr(lstImport.List(ii), "." & strRCP10 & ".") > 0 Then
                            bolUseRCP10 = True
                        End If
                        bolChkOk = True
                        Exit For
                  '右方List中案號超過一筆
                  ElseIf Val(.TextMatrix(jj, iCaseQty)) > 1 Then
                    '案件名稱有-(表示有相關案,需以相關案的案件性質上傳 ex:同一天有 核准-變更/核准-續展)
                    If InStr(.TextMatrix(jj, iCP10N), "-") > 0 Then
                        If InStr(lstImport.List(ii), "." & strRCP10 & ".") > 0 Then
                            bolUseRCP10 = True
                            bolChkOk = True
                            Exit For
                        End If
                    '檔名有案件性質
                    ElseIf InStr(lstImport.List(ii), "." & strCP10 & ".") > 0 Then
                        bolChkOk = True
                        Exit For
                    End If
                  End If
               End If
               'Add by Amy 2020/03/03 最後一筆都沒比對到
                If bolChkOk = False And jj = MSHFlexGrid1.Rows - 1 Then
                    strErr = "'無匹配資料！"
                End If
            Next
            End With
            
            'Add by Amy 2020/03/03 原UploadFile判斷左方List中需歸之檔案拆出
            If strCP09 = MsgText(601) Then
                If ChkUpload(strCP01, strCP02, strCP03, strCP04, strCP09, strCP10, strErr) = True Then
                    bolChkOk = True
                End If
            End If
            
            'Modify by Amy 2020/03/03 +if及iif
            If bolChkOk = True Then
                '上傳
                UploadFile lstImport.List(ii), txtPath, strCP01, strCP02, strCP03, strCP04, strCP09, strCP10, strErr, IIf(bolUseRCP10 = True, strRCP10, "")
            End If
            'end 2020/03/03
         End If
      End If
    
      If strErr <> "" Then
         lstImport.List(ii) = lstImport.List(ii) & " ..." & strErr
      Else
         lstImport.List(ii) = lstImport.List(ii) & " ...成功"
         lstImport.ItemData(ii) = 1
      End If
      
      txtProgressBar.Width = ii * (dblMaxWidth / dblFCnt): DoEvents
   Next
   txtProgressBar.Width = ii * (dblMaxWidth / dblFCnt): DoEvents
   SetListScroll lstImport
   
   jj = 0
   For ii = 0 To lstImport.ListCount - 1
      If lstImport.ItemData(ii - jj) = 1 Then
         lstImport.RemoveItem ii - jj
         jj = jj + 1
      End If
   Next
   
   ImportFile = True
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function

'Modify by Amy 2020/02/27 +相關案之案件性質pRCP10
Private Function UploadFile(pFileName As String, pPath As String, pCP01 As String, pCP02 As String, pCP03 As String, pCP04 As String, ByRef pCP09 As String, ByRef pCP10 As String, ByRef pErr As String, Optional ByVal pRCP10 As String = "") As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stFullPath As String, strFileName As String
   Dim stCon As String
   
On Error GoTo ErrHand
   
    'Mark by Amy 2020/03/03 程式改至ChkUpload
'   '左方List中需歸之檔案(可能後補)
'   If pCP09 = "" Then
'      'Modify by Amy 2020/02/27 +商標
'      '專利
'      If m_ProState = MsgText(601) Then
'        If textUser <> "" Then
'           stCon = " and cp65='" & textUser & "'"
'        End If
'        stSQL = "select cp09,cp10,pa09 from caseprogress,patent where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' and cp66=" & DBDATE(textDate) & stCon & " and cp09>'C' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 order by cp27 desc,cp09 asc"
'      '商標
'      Else
'        stSQL = "select cp09,cp10,tm10 pa09 from caseprogress,TradeMark,ServicePractice " & _
'                    "where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' " & _
'                    "and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 " & _
'                    "and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 " & _
'                    "and cp66=" & DBDATE(textDate) & stCon & " and cp09>'C'  order by cp27 desc,cp09 asc"
'      End If
'
'      intQ = 1
'      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
'      If intQ = 1 Then
'         If rsQuery("pa09") = "000" Then
'            Err.Raise 999, , "不可為台灣案"
'         Else
'            pCP09 = rsQuery(0)
'            pCP10 = rsQuery(1)
'         End If
'      Else
'         Err.Raise 999, , "該日" & IIf(textUser <> "" And m_ProState = "", textUser, "") & "沒輸入來函"
'      End If
'      'end 2020/02/27
'   End If
   
   'Modify by Amy 2020/02/27 +if 檔名若為相關案案件性質,需過濾
   strFileName = PUB_CaseNo2FileName(pCP01, pCP02, pCP03, pCP04) & "." & pCP10 & Mid(pFileName, InStr(pFileName, "."))
   If pRCP10 <> MsgText(601) Then
        strFileName = Replace(strFileName, "." & pRCP10 & ".", ".")
   'Modify by Amy 2020/02/26 +if 避免匯入檔名已有案件性質,需過濾
   ElseIf InStr(pFileName, "." & pCP10 & ".") > 0 Then
        strFileName = Replace(strFileName, "." & pCP10 & ".", ".")
   End If
   
   stFullPath = pPath & "\" & pFileName
   Set oFile = oFileSys.GetFile(stFullPath)
   
   If SaveAttFile_PDF(pCP09, stFullPath, strFileName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), True, , , True) Then
      'Added by Morgan 2023/9/7
      'P的OA來函有自動內部收文則同時將.altr更名為.order上傳至該內部收文卷宗區
      If pCP01 = "P" And (pCP10 = "1201" Or pCP10 = "1202") And Right(UCase(strFileName), 9) = ".ALTR.PDF" Then
         strExc(0) = "select CP09,CP10 from caseprogress where cp43='" & pCP09 & "' and cp10 in ('204','205') and cp09>'B' and not exists(select * from casepaperpdf where cpp01=cp09 and substr(UPPER(cpp02),-10)='.ORDER.PDF')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(1) = Left(strFileName, InStr(strFileName, ".") - 1) & "." & RsTemp("CP10") & ".ORDER.PDF"
            Call SaveAttFile_PDF(RsTemp("CP09"), stFullPath, strExc(1), Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), True, , , True)
         End If
      End If
      'end 2023/9/7
      oFile.Delete
      pErr = "" 'Add by Amy 2020/03/06
      UploadFile = True
   End If
   
ErrHand:
   If Err.NUMBER <> 0 Then
      pErr = Err.Description
   End If
   Set rsQuery = Nothing
End Function

Public Sub PubShowNextData()
   Dim iRow As Integer
   
   Me.Enabled = False
   Screen.MousePointer = vbHourglass
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      If Trim(.TextMatrix(iRow, 0)) = "V" Then
         Select Case cmdState
            'Added by Morgan 2016/7/6
            Case 0 '無缺檔
               If UpdateData(PUB_MGridGetValue(iRow, "cp09", MSHFlexGrid1), iRow) Then
                  PUB_SendMailCache 'Added by Morgan 2018/10/15
                  Me.Enabled = True
                  QueryData
               End If
            'end 2016/7/6
            Case 13 '卷宗區
               frm100101_L.m_strKey = PUB_MGridGetValue(iRow, "cp09", MSHFlexGrid1)   '總收文號
               frm100101_L.Hide
               frm100101_L.SetParent Me
               If frm100101_L.QueryData = True Then
                  frm100101_L.Show
                  Me.Hide
               End If
         End Select
         
         Exit For
      End If
   Next
   End With
   Screen.MousePointer = vbDefault
   Me.Enabled = True
End Sub

Private Function UpdateData(pCP09 As String, pRowID As Integer) As Boolean

On Error GoTo ErrHnd

   cnnConnection.BeginTrans

   strSql = "update caseprogress set cp121='Y' where cp09='" & pCP09 & "'"
   cnnConnection.Execute strSql
   
   strSql = "update letterprogress set lp02=0 where lp01='" & pCP09 & "'"
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   PUB_UpdateLP03 pCP09
      
   MailEngChk pRowID 'Adde by Morgan 2018/10/15
   
   cnnConnection.CommitTrans
   UpdateData = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Sub cmdOpen_Click()
   Dim stFileName As String
   Dim hLocalFile As Long
   Dim arrList() As String
   
   If lstImport.ListCount > 0 And lstImport.ListIndex > -1 Then
      If lstImport.ItemData(lstImport.ListIndex) = 1 Then
         MsgBox "檔案已匯入！！", vbInformation
      Else
         arrList = Split(lstImport.List(lstImport.ListIndex), " ")
         ShellExecute hLocalFile, "open", txtPath & "\" & arrList(0), vbNullString, vbNullString, 1
      End If
   End If
End Sub

Private Sub cmdPath_Click()
   Dim fName As String, strStartFolder As String
   
   If Dir(txtPath & "\", vbDirectory) <> "" Then strStartFolder = txtPath
   
   fName = PUB_GetFolder(Me.hWnd, strStartFolder, "請選取資料夾:")
   If fName <> "" Then 'they did not hit cancel
      txtPath = fName
      'SaveSetting "TAIE", "P", UCase(Me.Name) & "Dir", txtPath 'Removed by Morgan 2020/4/8 移到 unload (Wind10 網路路徑可能會選不到需要人工輸入)
   End If
   RefreshList
   
End Sub

Private Sub cmdPrint_Click()
   PUB_RestorePrinter cmbPrinter
   DoPrint
   PUB_RestorePrinter strPrinter
End Sub

Private Sub cmdQuery_Click()
   RefreshList
   QueryData True
End Sub


Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cmbPrinter, strPrinter
   dblMaxWidth = txtProgressBar.Width
   '讀取前次設定路徑
   txtPath.Text = GetSetting("TAIE", "P", UCase(Me.Name) & "Dir", "")
   txtPath.Tag = txtPath.Text 'Added by Morgan 2020/4/8
   If txtPath <> "" Then
      'Modified by Morgan 2017/1/12
      'If oFileSys.FolderExists(txtPath) = False Then
      If PUB_ChkDir(txtPath) = False Then
         MsgBox "副本存放路徑 [ " & txtPath & " ] 不存在，請重新設定！", vbCritical
         txtPath = "C:\"
      End If
   Else
      txtPath = "C:\"
   End If
   
   textDate = strSrvDate(2)
   textUser = strUserNum
   
   RefreshList
   QueryData
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   KillTemp
End Sub

Private Sub KillTemp()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Added by Morgan 2020/4/8
   If txtPath.Tag <> txtPath.Text Then
      SaveSetting "TAIE", "P", UCase(Me.Name) & "Dir", txtPath
   End If
   'end 2020/4/8
   
   If Me.cmbPrinter.Text <> Me.cmbPrinter.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
   End If
   Set oFileSys = Nothing
   Set frm040120 = Nothing
End Sub

Private Sub RefreshList()
   Dim ii As Integer
    
   lstImport.Clear
   File1.path = txtPath.Text
   File1.Refresh
   If File1.ListCount > 0 Then
      For ii = 0 To File1.ListCount - 1
         If UCase(Right(Trim(File1.List(ii)), 4)) = ".PDF" Then
            lstImport.AddItem Trim(File1.List(ii))
         End If
      Next
      cmdOpen.Enabled = True
   Else
      cmdOpen.Enabled = False
   End If
End Sub

Private Sub QueryData(Optional pShowMsg As Boolean = False)
   Dim ii As Integer
   Dim stVTB As String
   Dim iCp09 As Integer, iCP10 As Integer, iCPM26 As Integer, iCP145 As Integer, iLP02 As Integer
   Dim idx1 As Integer, iAltr As Integer, iNQty As Integer
   Dim rsQeury As ADODB.Recordset
   Dim bCancel As Boolean
   Dim stCon As String
   Dim strCountCase As String 'Add byAmy 2020/02/26
   
   'Added by Morgan 2016/7/7
   If textUser = "" Then
      MsgBox "來函輸入人員不可空白!!", vbExclamation
      If textUser.Enabled Then textUser.SetFocus
      Exit Sub
   Else
      textUser_Validate bCancel
      If bCancel Then
         If textUser.Enabled Then textUser.SetFocus
         Exit Sub
      'Moidfy by Amy 2020/01/09 商標輪流操作,故不需判斷建立人員-桂英(專利非一人操作,只允許匯自己的案件)
      ElseIf m_ProState = MsgText(601) Then
         'Modified by Morgan 2018/7/17 +cp68 CFP電子化
         stCon = " and (c1.cp65='" & textUser & "' or c1.cp68='" & textUser & "')"
      End If
   End If
   'end 2016/7/7
   
   intPrevRow = 0
   SetGrid True
   cmdOK(13).Enabled = False
   
   'Modify by Amy 2020/01/09 +if 增加商標,原專利抓服務業務時,加入系統別條件
   '專利
   If m_ProState = MsgText(601) Then
    '相關收文號有發文代理人(不一定有相關收文號，無需特別控制。Ex:通知年費逾期)
    '代理人來函檔案不存在(cp121 is null)
    'P案領證,年費已提申沒有附件--玲玲
    'P案專利證書(1603),核發-申請專利權評價報告(1008)沒有altr--玲玲
    'Modified by Morgan 2016/7/7 +來函輸入人員條件
    'Modified by Morgan 2016/7/27 副檔.data 的改為.req--玲玲
    'Modified by Morgan 2017/10/18 +代理人來函(ALTR) 可放 MSG 格式
    
    '未齊備收文號之檔案數(Qty)及是否有Alter檔(Altr)
    'Modified by Morgan 2018/7/18 +剔除空白回覆單(.BLANK.PDF)
    'Modified by Morgan 2018/9/27 +.FIL.PDF
    'Modified by Morgan 2021/3/19 +排除 .info. (IDS報價會先上傳卷宗區)--玲玲
    stVTB = "select cpp01,count(*) Qty,max(decode(sign(instr(upper(cpp02),'.ALTR.')),1,'Y')) Altr" & _
       ",max(decode(sign(instr(upper(cpp02),'.RECEIPT.PDF')+instr(upper(cpp02),'.FIL.PDF')),1,'Y')) Rcp from letterprogress,casepaperpdf" & _
       " where lp03=0 and lp01>'C' and cpp01(+)=lp01 and instr(upper(cpp02),'.CUS.PDF')=0 and instr(upper(cpp02),'.BLANK.PDF')=0 and instr(upper(cpp02),'.INFO.')=0 and cpp10<>'D'" & _
       " AND (SUBSTR(UPPER(CPP02),-4)='.PDF' or (SUBSTR(UPPER(CPP02),-4)='.MSG' and instr(upper(cpp02),'.ALTR.')>0))" & _
       " group by cpp01"
       
    'Modify by Amy 2020/02/26 +案號筆數,同一天同一案號筆數(商標可能有 同一案號不同案件性質需匯,若有未輸案件性質可能會歸錯)
    strCountCase = "Select c1.CP01||c1.CP02||c1.CP03||c1.CP04 QCNo,Count(*) CaseQty From Letterprogress, Caseprogress c1, Caseprogress c2,Patent,Servicepractice " & _
                             " Where lp03=0 And lp01>'C' And c1.cp09(+)=lp01 And c2.cp09(+)=c1.cp43" & _
                             " And not (c1.cp01='P' And c1.cp10='1909' And (c2.cp10='601' or c2.cp10='605'))" & stCon & _
                            " And pa01(+)=c1.cp01 And pa02(+)=c1.cp02 And pa03(+)=c1.cp03 And pa04(+)=c1.cp04 And nvl(pa09,sp09)<>'000' " & _
                            " And sp01(+)=c1.cp01 And sp02(+)=c1.cp02 And sp03(+)=c1.cp03 And sp04(+)=c1.cp04 And c1.cp01 in('" & Replace(GetSystemKindByNick, ",", "','") & "')" & _
                            " Group by c1.CP01||c1.CP02||c1.CP03||c1.CP04 "

    'Modified by Morgan 2018/9/26 +lp04
    'Modified by Morgan 2020/1/8 +lp43
    'Modify by Amy 2020/02/27 +相關總收文號案件性質名稱及編號
    strExc(0) = "select '' V,c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04) as 本所案號" & _
       ",m1.cpm04||decode(m2.cpm04,'','','-'||m2.cpm04)||GetRelateCasePropertyName(c1.cp09, '1') as 案件性質,sqldatet(c1.cp05) as 收文日,lp02-nvl(Qty,0)||'/'||lp02 as 缺檔" & _
       ",decode(c1.cp01||c1.cp10,'P1603','','P1008','',decode(Altr,'','altr')) 說明,lp02-nvl(Qty,0) NQty" & _
       ",c1.cp01,c1.cp02,c1.cp03,c1.cp04,c1.cp05,c1.cp09,c1.cp10,c1.cp43,c1.cp145,lp04,lp10,lp43" & _
       ",decode(c2.cp10,'307',decode(pa08,'1','inv','2','utl'),decode(lower(m2.cpm26),'data','req',m2.cpm26)) cpm26,lp02,lp19,rcp,CaseQty,c2.cp10 as RCp10" & _
       " From letterprogress, caseprogress c1, caseprogress c2, casepropertymap m1, casepropertymap m2,patent,servicepractice" & _
       ",(" & stVTB & "),(" & strCountCase & ") where lp03=0 and lp01>'C' and c1.cp09(+)=lp01 and c2.cp09(+)=c1.cp43 And c1.CP01||c1.CP02||c1.CP03||c1.CP04=QCNo(+)" & _
       " and not (c1.cp01='P' and c1.cp10='1909' and (c2.cp10='601' or c2.cp10='605'))" & stCon & _
       " and m1.cpm01(+)=c1.cp01 and m1.cpm02(+)=c1.cp10 and m2.cpm01(+)=c2.cp01 and m2.cpm02(+)=c2.cp10 and cpp01(+)=lp01" & _
       " and pa01(+)=c1.cp01 and pa02(+)=c1.cp02 and pa03(+)=c1.cp03 and pa04(+)=c1.cp04 and nvl(pa09,sp09)<>'000' " & _
       " and sp01(+)=c1.cp01 and sp02(+)=c1.cp02 and sp03(+)=c1.cp03 and sp04(+)=c1.cp04 and c1.cp01 in('" & Replace(GetSystemKindByNick, ",", "','") & "')" & _
       " order by c1.cp05 desc,c1.cp01 asc,c1.cp02 asc,c1.cp03 asc,c1.cp04 asc,c1.cp09 asc"
    'end 2020/02/26
   '商標
   Else
    'Momo by Amy T電子化-目前檔案數抓法同專利
    stVTB = "select cpp01,count(*) Qty,max(decode(sign(instr(upper(cpp02),'.ALTR.')),1,'Y')) Altr" & _
       ",max(decode(sign(instr(upper(cpp02),'.RECEIPT.PDF')+instr(upper(cpp02),'.FIL.PDF')),1,'Y')) Rcp from letterprogress,casepaperpdf" & _
       " where lp03=0 and lp01>'C' and cpp01(+)=lp01 and instr(upper(cpp02),'.CUS.PDF')=0 and instr(upper(cpp02),'.BLANK.PDF')=0 and cpp10<>'D'" & _
       " AND (SUBSTR(UPPER(CPP02),-4)='.PDF' or (SUBSTR(UPPER(CPP02),-4)='.MSG' and instr(upper(cpp02),'.ALTR.')>0))" & _
       " group by cpp01"
       
    'Modify by Amy 2020/02/26 +案號筆數,同一天同一案號筆數(商標可能有 同一案號不同案件性質需匯,若有未輸案件性質可能會歸錯)
    strCountCase = "Select c1.CP01||c1.CP02||c1.CP03||c1.CP04 QCNo,Count(*) CaseQty From Letterprogress, Caseprogress c1, TradeMark,Servicepractice " & _
                             "Where lp03=0 And lp01>'C' And c1.cp09(+)=lp01 And nvl(tm10,sp09)<>'000'  " & stCon & _
                             " And tm01(+)=c1.cp01 And tm02(+)=c1.cp02 And tm03(+)=c1.cp03 And tm04(+)=c1.cp04 " & _
                             " And sp01(+)=c1.cp01 And sp02(+)=c1.cp02 And sp03(+)=c1.cp03 And sp04(+)=c1.cp04 " & _
                             " And c1.cp01 in('" & Replace(GetSystemKindByNick, ",", "','") & "')" & _
                             " Group by CP01||CP02||CP03||CP04 "

    'Modify by Amy 2020/02/27 +相關總收文號案件性質名稱及編號
    strExc(0) = "select '' V,c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04) as 本所案號" & _
       ",m1.cpm04||GetRelateCasePropertyName(c1.cp09, '1') as 案件性質,sqldatet(c1.cp05) as 收文日,lp02-nvl(Qty,0)||'/'||lp02 as 缺檔" & _
       ",'' 說明,lp02-nvl(Qty,0) NQty" & _
       ",c1.cp01,c1.cp02,c1.cp03,c1.cp04,c1.cp05,c1.cp09,c1.cp10,c1.cp43,c1.cp145,lp04,lp10,lp43" & _
       ",'' cpm26,lp02,lp19,rcp,CaseQty,c2.cp10 as RCp10" & _
       " From letterprogress, caseprogress c1, caseprogress c2,  casepropertymap m1, TradeMark,servicepractice" & _
       ",(" & stVTB & "),(" & strCountCase & ") where lp03=0 and lp01>'C' and c1.cp09(+)=lp01 And c2.cp09(+)=c1.cp43 And c1.CP01||c1.CP02||c1.CP03||c1.CP04=QCNo(+) " & stCon & _
       " and m1.cpm01(+)=c1.cp01 and m1.cpm02(+)=c1.cp10 and cpp01(+)=lp01" & _
       " and tm01(+)=c1.cp01 and tm02(+)=c1.cp02 and tm03(+)=c1.cp03 and tm04(+)=c1.cp04 and nvl(tm10,sp09)<>'000' " & _
       " and sp01(+)=c1.cp01 and sp02(+)=c1.cp02 and sp03(+)=c1.cp03 and sp04(+)=c1.cp04 and c1.cp01 in('" & Replace(GetSystemKindByNick, ",", "','") & "')" & _
       " order by c1.cp05 desc,c1.cp01 asc,c1.cp02 asc,c1.cp03 asc,c1.cp04 asc,c1.cp09 asc"
   'end 2020/02/26
   End If
   'end 2020/01/09
   intI = 1
   With MSHFlexGrid1
   .FixedCols = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '若沒有資料時不可直接設定給 Grid 否則 MouseRow 會跑掉
   If intI = 1 Then
      Set .Recordset = RsTemp
      iCp09 = PUB_MGridGetId("cp09", MSHFlexGrid1)
      iCP10 = PUB_MGridGetId("cp10", MSHFlexGrid1)
      iCP145 = PUB_MGridGetId("cp145", MSHFlexGrid1)
      iCPM26 = PUB_MGridGetId("cpm26", MSHFlexGrid1)
      idx1 = PUB_MGridGetId("說明", MSHFlexGrid1)
      iNQty = PUB_MGridGetId("NQty", MSHFlexGrid1)
      iLP02 = PUB_MGridGetId("lp02", MSHFlexGrid1)
      For ii = 1 To .Rows - 1
         'Modify by Amy 2020/01/09 +if 增加商標,原專利抓服務業務時,加入系統別條件
         '專利
         If m_ProState = MsgText(601) Then
            '提申有副本(cp145='Y')要檢查案件性質副檔(inv,utl...)
            'Modified by Morgan 2019/8/30 +通知申請日1101 Ex:CFP-031113
            If (.TextMatrix(ii, iCP10) = "1101" Or .TextMatrix(ii, iCP10) = "1102" Or .TextMatrix(ii, iCP10) = "1909") Then
               If .TextMatrix(ii, iCP145) = "Y" Then
                  If .TextMatrix(ii, iCPM26) <> "" Then
                     strExc(0) = "select cpp02 from casepaperpdf where cpp01='" & .TextMatrix(ii, iCp09) & "' and instr(upper(cpp02),'." & UCase(.TextMatrix(ii, iCPM26)) & ".PDF')>0 and cpp10<>'D'"
                     intI = 1
                     Set rsQeury = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 0 Then
                        '填入缺檔
                        If .TextMatrix(ii, idx1) = "" Then
                           .TextMatrix(ii, idx1) = .TextMatrix(ii, iCPM26)
                        Else
                           .TextMatrix(ii, idx1) = .TextMatrix(ii, idx1) & "," & .TextMatrix(ii, iCPM26)
                        End If
                     End If
                  End If
               End If
            'Added by Morgan 2016/12/1
            'Modified by Morgan 2017/1/13 專利證書(1603)和核發(1008)要檢查公文
            'ElseIf Val(.TextMatrix(ii, iLP02)) > 1 Then
            'Modified by Morgan 2018/7/19 代理人通知修正1224除外(沒公文但可能有附件)
            'Modified by Morgan 2019/3/8 依職權電話通知修正1225沒有公文
            ElseIf (Val(.TextMatrix(ii, iLP02)) > 1 Or .TextMatrix(ii, iCP10) = "1603" Or .TextMatrix(ii, iCP10) = "1008") And _
               .TextMatrix(ii, iCP10) <> "1224" And .TextMatrix(ii, iCP10) <> "1225" Then
               strExc(0) = "select cpp02 from casepaperpdf where cpp01='" & .TextMatrix(ii, iCp09) & "' and instr(upper(cpp02),'." & .TextMatrix(ii, iCP10) & ".PDF')>0 and cpp10<>'D'"
               intI = 1
               Set rsQeury = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
                  '填入缺檔
                  If .TextMatrix(ii, idx1) = "" Then
                     .TextMatrix(ii, idx1) = "公文"
                  Else
                     .TextMatrix(ii, idx1) = .TextMatrix(ii, idx1) & ",公文"
                  End If
               End If
            'end 2016/12/1
            End If
            'Added by Morgan 2018/6/29 CFP電子化,若有用收據整批匯入功能時要取消此處控制
            If PUB_MGridGetValue(ii, "cp01", MSHFlexGrid1) = "CFP" And PUB_MGridGetValue(ii, "lp19", MSHFlexGrid1) = "Y" And PUB_MGridGetValue(ii, "rcp", MSHFlexGrid1) <> "Y" Then
               '填入缺檔
               If .TextMatrix(ii, idx1) = "" Then
                  .TextMatrix(ii, idx1) = "收據"
               Else
                  .TextMatrix(ii, idx1) = .TextMatrix(ii, idx1) & ",收據"
               End If
            End If
            'end 2018/6/29
            
         '商標
         Else
            '1702 通知修正/1706 其他來函 不會有公文
            If .TextMatrix(ii, iCP10) <> "1702" And .TextMatrix(ii, iCP10) <> "1706" Then
                strExc(0) = "select cpp02 from casepaperpdf where cpp01='" & .TextMatrix(ii, iCp09) & "' and instr(upper(cpp02),'." & .TextMatrix(ii, iCP10) & ".PDF')>0 and cpp10<>'D'"
                intI = 1
                Set rsQeury = ClsLawReadRstMsg(intI, strExc(0))
                If intI = 0 Then
                     '填入缺檔
                     'Modify By Sindy 2020/3/5 1909.已提申缺的檔案不會是公文,顯示申請書
                     If .TextMatrix(ii, iCP10) = "1909" Then
                        If .TextMatrix(ii, idx1) = "" Then
                            .TextMatrix(ii, idx1) = "申請書"
                        Else
                            .TextMatrix(ii, idx1) = .TextMatrix(ii, idx1) & ",申請書"
                        End If
                     Else
                     '2020/3/5 END
                        If .TextMatrix(ii, idx1) = "" Then
                            .TextMatrix(ii, idx1) = "公文"
                        Else
                            .TextMatrix(ii, idx1) = .TextMatrix(ii, idx1) & ",公文"
                        End If
                     End If
                End If
            End If
         End If
         'end 2020/01/09
         
         '沒缺檔且要檢查的檔案已存在
         If Trim(.TextMatrix(ii, idx1)) = "" And Val(.TextMatrix(ii, iNQty)) <= 0 Then
            
            cnnConnection.BeginTrans
            
On Error GoTo ErrHnd:
            strSql = "update caseprogress set cp121='Y' where cp09='" & .TextMatrix(ii, iCp09) & "'"
            cnnConnection.Execute strSql
            
            MailEngChk ii 'Adde by Morgan 2018/9/26
            
            cnnConnection.CommitTrans
            
            .RowHeight(ii) = 0
         End If
      Next
      
      PUB_UpdateLP03
      
      PUB_SendMailCache 'Added by Morgan 2018/9/27
      
      SetGrid
      cmdOK(13).Enabled = True
      MSHFlexGrid1.row = 1
      ShowBar MSHFlexGrid1, intPrevRow, MSHFlexGrid1.Cols - 1
      MSHFlexGrid1.TextMatrix(.row, 0) = "V"
   Else
      If pShowMsg Then MsgBox "沒有待匯入案件！", vbInformation
      PUB_UpdateLP03
   End If
   m_iCols = .Cols
   End With
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   
   Set rsQeury = Nothing
End Sub

'Adde by Morgan 2018/9/26
'CFP已提申1909有通知信且非自判時EMail通知判發人
Private Sub MailEngChk(pRowID As Integer)
   'Modified by Morgan 2020/1/8 沒有客戶函也要判發，改判斷 LP43
   'If PUB_MGridGetValue(pRowID, "cp01", MSHFlexGrid1) = "CFP" And PUB_MGridGetValue(pRowID, "CP10", MSHFlexGrid1) = "1909" And PUB_MGridGetValue(pRowID, "lp10", MSHFlexGrid1) = "Y" Then
   If PUB_MGridGetValue(pRowID, "cp01", MSHFlexGrid1) = "CFP" And PUB_MGridGetValue(pRowID, "CP10", MSHFlexGrid1) = "1909" And PUB_MGridGetValue(pRowID, "lp43", MSHFlexGrid1) = "Y" Then
      strExc(1) = PUB_MGridGetValue(pRowID, "lp04", MSHFlexGrid1)
      If strExc(1) <> "" Then
          strExc(0) = PUB_MGridGetValue(pRowID, "本所案號", MSHFlexGrid1) & "(" & PUB_MGridGetValue(pRowID, "案件性質", MSHFlexGrid1) & ")公文來函待判發通知!!"
          'Modified by Morgan 2018/11/26 +考慮判發人不是用承辦人系統
          'strExc(2) = "請至案件管理系統的 專利處\公文來函判發作業 進行判發。"
          strExc(2) = "【承辦人系統】：請至【專利處\公文來函判發作業】進行判發。" & vbCrLf & _
            "【專利及承辦人系統】：請至【承辦人\公文來函判發作業】進行判發。"
          strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
            " values ('" & strUserNum & "','" & strExc(1) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & ChgSQL(strExc(0)) & "','" & ChgSQL(strExc(2)) & "')"
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2018/9/26
End Sub

Private Sub DoPrint()
   Dim ii As Integer, jj As Integer
   Dim strFontName As String
   
   If Check1.Value <> vbChecked And Check2.Value <> vbChecked Then
      MsgBox "請勾選要列印的內容！", vbInformation
      Exit Sub
   End If
   
   strFontName = Printer.FontName
   Printer.FontName = "細明體"
   
   '待匯入案件
   If Check1.Value = 1 Then
      GetPleft 1
      PrintTitle 1
      For jj = 1 To MSHFlexGrid1.Rows - 1
         For ii = 1 To 4
            strTemp(ii) = "" & MSHFlexGrid1.TextMatrix(jj, ii)
         Next ii
         If (iNowLine + 2) * iRowHeight > Printer.ScaleHeight Then
            Printer.NewPage
            PrintTitle 1  '列印表頭
         End If
         PrintDetail '列印明細
      Next jj
      Printer.EndDoc
   End If
   
   '匯入結果
   If Check2.Value = 1 Then
      GetPleft 2
      PrintTitle 2
      For jj = lstImport.ListCount - 1 To 0 Step -1
         strTemp(1) = lstImport.List(jj)
         If (iNowLine + 2) * iRowHeight > Printer.ScaleHeight Then
            Printer.NewPage
            PrintTitle 2 '列印表頭
         End If
         
         PrintDetail '列印明細
         
      Next jj
      Printer.EndDoc
   End If
   
   Printer.FontName = strFontName
End Sub

Sub GetPleft(Optional pIndex As Integer = 1)
   iRowHeight = 300
   If pIndex = 1 Then
      ReDim PLeft(1 To 5)
      ReDim strTemp(1 To 5)
      
      PLeft(1) = 500
      PLeft(2) = 2500
      PLeft(3) = 6000
      PLeft(4) = 7500
      PLeft(5) = 9000
   Else
      ReDim PLeft(1 To 1)
      ReDim strTemp(1 To 1)
      
      PLeft(1) = 500
   End If
End Sub

Sub PrintTitle(Optional pIndex As Integer = 1)
Dim strTitle As String

iNowLine = 0
If pIndex = 1 Then
   strTitle = "待匯入案件"
Else
   strTitle = "匯入結果"
End If

iNowLine = 1

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(Me.Caption) / 2)
Printer.CurrentY = iNowLine * iRowHeight
Printer.Print Me.Caption

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

iNowLine = iNowLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = 900
Printer.Print "列印人員：" & strUserName
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
iNowLine = iNowLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 1200
Printer.Print "頁　　次：" & Printer.Page

iNowLine = 5
If pIndex = 1 Then
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iNowLine * iRowHeight
   Printer.Print MSHFlexGrid1.TextMatrix(0, 1)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iNowLine * iRowHeight
   Printer.Print MSHFlexGrid1.TextMatrix(0, 2)
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iNowLine * iRowHeight
   Printer.Print MSHFlexGrid1.TextMatrix(0, 3)
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iNowLine * iRowHeight
   Printer.Print MSHFlexGrid1.TextMatrix(0, 4)
Else
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iNowLine * iRowHeight
   Printer.Print strTitle
End If
iNowLine = iNowLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iNowLine * iRowHeight
Printer.Print String(148, "-")
iNowLine = iNowLine + 1
End Sub

Sub PrintDetail()
   Dim ii As Integer
   
   For ii = 1 To UBound(PLeft)
      Printer.CurrentX = PLeft(ii)
      Printer.CurrentY = iNowLine * iRowHeight
      Printer.Print strTemp(ii)
   Next ii
   iNowLine = iNowLine + 1
End Sub

Private Sub SetListScroll(oList As ListBox)
   Dim ii As Integer
   Dim lWnow As Long, lWmax As Long
   
   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next
   'Add By Sindy 2020/3/5 換算值似乎有問題,給預設值5000,至少不要讓使用者看不到匯入狀況的內容
   If lWmax < 5000 Then
      lWmax = 5000
   End If
   '2020/3/5 END
   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

Private Sub SetGrid(Optional pReset As Boolean = False)
   Dim arrMSHFlexGrid1HeadText, arrMSHFlexGrid1HeadWidth
   Dim iCol As Integer
   Dim iUbound As Integer

   arrMSHFlexGrid1HeadText = Array("V", "本所案號", "案件性質", "收文日", "缺檔", "說明")
   arrMSHFlexGrid1HeadWidth = Array(200, 1035, 1605, 800, 400, 1600)
   iUbound = UBound(arrMSHFlexGrid1HeadWidth)
   
   With MSHFlexGrid1
   .Visible = False
   If pReset = True Then
      .Clear
      .Rows = 2
   End If
   '.FixedCols = 2
   For iCol = 0 To .Cols - 1
      .row = 0
      .col = iCol
      If iCol <= iUbound Then
         .Text = arrMSHFlexGrid1HeadText(iCol)
         .ColWidth(iCol) = arrMSHFlexGrid1HeadWidth(iCol)
         .ColAlignment(iCol) = flexAlignLeftCenter
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   .Visible = True
   End With
End Sub

Private Sub lstImport_DblClick()
   MsgBox lstImport.List(lstImport.ListIndex)
End Sub

Private Sub MSHFlexGrid1_Click()
   Dim nCol As Long, nRow As Long, iRow As Integer
   
   With MSHFlexGrid1
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow > 0 And .TextMatrix(nRow, 1) <> "" Then
      If intPrevRow <> nRow Then
         '清除上一筆資料列反白
         If intPrevRow > 0 Then
            .row = intPrevRow
            ShowBar MSHFlexGrid1, intPrevRow, MSHFlexGrid1.Cols - 1
            MSHFlexGrid1.TextMatrix(.row, 0) = ""
         End If
         .row = nRow
         ShowBar MSHFlexGrid1, intPrevRow, MSHFlexGrid1.Cols - 1
         MSHFlexGrid1.TextMatrix(.row, 0) = "V"
      End If
   End If
   End With
End Sub

Private Sub textDate_GotFocus()
   InverseTextBox textDate
End Sub

Private Sub textDate_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textDate) = False Then
      If CheckIsTaiwanDate(textDate, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的輸入日期"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDate_GotFocus
         GoTo EXITSUB
      End If
      
      '發文日不能大於系統日
      If DBDATE(textDate) > strSrvDate(1) Then
         Cancel = True
         strMsg = "輸入日期不能大於系統日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDate_GotFocus
      End If
   End If
EXITSUB:
End Sub

Private Sub textUser_GotFocus()
   TextInverse textUser
End Sub

Private Sub textUser_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textUser_Validate(Cancel As Boolean)
   lblUserName = ""
   If textUser <> "" Then
      lblUserName = GetStaffName(textUser, True)
      If lblUserName = "" Then
         Cancel = True
      End If
   End If
End Sub

'Add by Amy 2020/03/03 從ImportFile拆出來修改,判斷是否需匯入(可能後補的檔案)
Private Function ChkUpload(pCP01 As String, pCP02 As String, pCP03 As String, pCP04 As String, ByRef pCP09 As String, ByRef pCP10 As String, ByRef pErr As String) As Boolean
    Dim stSQL As String, intQ As Integer
    Dim rsQuery As ADODB.Recordset
    Dim stCon As String
   
On Error GoTo ErrHand

    ChkUpload = False
   
    '左方List中需歸之檔案(可能後補)
    '專利
    If m_ProState = MsgText(601) Then
        If textUser <> "" Then
            stCon = " and cp65='" & textUser & "'"
        End If
        stSQL = "select cp09,cp10,pa09 from caseprogress,patent " & _
                     "where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' " & _
                     "and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 " & _
                     "and cp66=" & DBDATE(textDate) & stCon & " and cp09>'C'  order by cp27 desc,cp09 asc"
      '商標
    Else
        stSQL = "select cp09,cp10,Decode(tm01,null,sp09,tm10) pa09 from caseprogress,TradeMark,ServicePractice " & _
                    "where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' " & _
                    "and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 " & _
                    "and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 " & _
                    "and cp66=" & DBDATE(textDate) & stCon & " and cp09>'C'  order by cp27 desc,cp09 asc"
    End If
    intQ = 1
    Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
    If intQ = 1 Then
        If rsQuery("pa09") = "000" Then
            Err.Raise 999, , "不可為台灣案"
        Else
            pCP09 = rsQuery(0)
            pCP10 = rsQuery(1)
            ChkUpload = True
        End If
    Else
        'Modify by Amy 2020/03/03 +m_ProState="",專利才需Show人
        Err.Raise 999, , "該日" & IIf(textUser <> "" And m_ProState = "", textUser, "") & "沒輸入來函"
    End If
   
ErrHand:
   If Err.NUMBER <> 0 Then
      pErr = Err.Description
   End If
   Set rsQuery = Nothing
End Function
