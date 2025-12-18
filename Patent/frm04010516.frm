VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010516 
   BorderStyle     =   1  '單線固定
   Caption         =   "電子公文來函"
   ClientHeight    =   4752
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9132
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4752
   ScaleWidth      =   9132
   Begin VB.ComboBox cmbPrinter2 
      Height          =   276
      Left            =   1296
      TabIndex        =   18
      Top             =   4416
      Width           =   4320
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "刪除"
      Height          =   400
      Left            =   6435
      TabIndex        =   15
      Top             =   90
      Width           =   780
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "重整(&Q)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   5625
      TabIndex        =   12
      Top             =   90
      Width           =   780
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "輸入"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   7245
      TabIndex        =   11
      Top             =   90
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印清單"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5940
      TabIndex        =   10
      Top             =   540
      Width           =   1005
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "列印附件"
      Height          =   400
      Index           =   1
      Left            =   8010
      TabIndex        =   7
      Top             =   525
      Width           =   960
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "開啟附件"
      Height          =   400
      Index           =   0
      Left            =   6975
      TabIndex        =   6
      Top             =   525
      Width           =   1005
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   8145
      TabIndex        =   5
      Top             =   90
      Width           =   825
   End
   Begin VB.TextBox txtPDFPath 
      Height          =   315
      Left            =   1635
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "C:\Program Files\Adobe\Reader 8.0\Reader\AcroRd32.exe"
      Top             =   570
      Width           =   4245
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   276
      Left            =   1284
      TabIndex        =   0
      Top             =   4080
      Width           =   4320
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3015
      Left            =   90
      TabIndex        =   4
      Top             =   960
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   5313
      _Version        =   393216
      Cols            =   14
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "V|本所案號|申請案號|案由|簽收日期|處理期限|發文日期|發文文號|案件種類|檔案|送達時間|受送達人|案號類別|案號"
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
      _Band(0).Cols   =   14
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "證書印表機："
      Height          =   180
      Left            =   192
      TabIndex        =   19
      Top             =   4476
      Width           =   1080
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      Caption         =   "程序人員："
      Height          =   240
      Left            =   216
      TabIndex        =   17
      Top             =   192
      Visible         =   0   'False
      Width           =   948
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1200
      TabIndex        =   16
      Top             =   168
      Visible         =   0   'False
      Width           =   2088
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3678;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblTotal 
      Height          =   180
      Left            =   6435
      TabIndex        =   14
      Top             =   4140
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "總筆數："
      Height          =   180
      Index           =   0
      Left            =   5670
      TabIndex        =   13
      Top             =   4140
      Width           =   720
   End
   Begin VB.Label lblCount 
      Height          =   180
      Left            =   8145
      TabIndex        =   9
      Top             =   4140
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "已勾選筆數："
      Height          =   180
      Index           =   1
      Left            =   7020
      TabIndex        =   8
      Top             =   4140
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PDF執行檔路徑："
      Height          =   180
      Left            =   225
      TabIndex        =   3
      Top             =   630
      Width           =   1380
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "公文印表機："
      Height          =   180
      Left            =   180
      TabIndex        =   2
      Top             =   4140
      Width           =   1080
   End
End
Attribute VB_Name = "frm04010516"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modified by Morgan 2025/1/15 增加程序人員選單並刪除不再使用的物件及部分舊程式碼
'Memo by Morgan 2021/12/22 改成Form2.0 (MSHFlexGrid1,Printer列印未改)
'Created by Morgan 2014/1/9
Option Explicit

'執行各項功能的權限
Dim m_bDelete As Boolean

Dim m_AttachPath As String
'列印用
Dim PLeft() As Integer, iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500, ciColGap = 250
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim strPrinter As String
Dim strPrinter2 As String 'Added by Morgan 2025/2/17
Dim m_Sys As String
Dim m_iCols As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim oFileSys As New FileSystemObject
Dim oFile As File

Private Sub cmdDelete_Click()
   Screen.MousePointer = vbHourglass
   If DeleteCheck = True Then
      FormDelete
   End If
   Screen.MousePointer = vbDefault
End Sub

'Modified by Morgan 2014/5/5 改只能刪除非P案且不必勾選
'Modified by Morgan 2019/7/5 改不限非P案
Private Function DeleteCheck() As Boolean
   Dim iRow As Integer, iCount As Integer, idx As Integer
   Dim bConfirm As Boolean, bolNoPCase As Boolean

   iCount = 0
   With MSHFlexGrid1
   idx = GetFieldId("pa01")
   For iRow = 1 To .Rows - 1
'      If .TextMatrix(iRow, idx) = "P" Then
'         If .TextMatrix(iRow, 0) = "V" Then
'            .row = iRow
'            ClickGrid MSHFlexGrid1
'         End If
'      Else
'         If .TextMatrix(iRow, 0) = "" Then
'            .row = iRow
'            ClickGrid MSHFlexGrid1
'         End If
'      End If
      If .TextMatrix(iRow, 0) = "V" Then iCount = iCount + 1
   Next
   End With
   
   If iCount = 0 Then
      'MsgBox "無非P案可刪除！", vbExclamation
      MsgBox "請先勾選！", vbExclamation
      Exit Function
   Else
      If MsgBox("共有 " & iCount & " 筆電子來函紀錄將刪除，是否確定要繼續？" & vbCrLf & vbCrLf & "注意：請先確認來函內容後才可刪除！", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
         Exit Function
      End If
   End If
   DeleteCheck = True
   
End Function

Private Function FormDelete() As Boolean
   Dim iRow As Integer, stNo As String
   Dim stSys As String
   
On Error GoTo ErrHnd
   
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      .row = iRow
      If .TextMatrix(.row, 0) = "V" Then
         cnnConnection.BeginTrans
         
On Error GoTo ErrHndT

         stNo = GetValue(.row, "ed01")
         'Modified by Morgan 2017/5/11非P案不再刪除以進行測試
         'Modified by Morgan 2017/11/22 +還原(可刪除非本所案件)
         PUB_DelFtpFile2 stNo 'Added by Morgan 2015/4/15 檔案改放 FTP,必須在DB資料刪除前執行
         strSql = "delete casepaperpdf where cpp01='" & stNo & "'"
         cnnConnection.Execute strSql, intI
         
         'Modified by Morgan 2020/3/25
         'strSql = "delete edocument where ed01='" & stNo & "'"
         strSql = "update edocument set ed11='不收文' where ed01='" & stNo & "'"
         
         cnnConnection.Execute strSql, intI
         'stSys = GetValue(.row, "pa01")
         'If stSys <> "FCP" Then stSys = "FCT"
         'strSql = "update edocument set ed30='" & stSys & "' where ed01='" & stNo & "'"
         'cnnConnection.Execute strSql, intI
         'end 2017/11/22
         'end 2017/5/11
         
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 0) = "X"
         MSHFlexGrid1.RowHeight(MSHFlexGrid1.row) = 0
         cnnConnection.CommitTrans
         
         lblCount = Val(lblCount) - 1
         lblTotal = lblTotal - 1
         DoEvents
      End If
   Next
   End With
   
   
   FormDelete = True
   Exit Function
   
ErrHndT:
   cnnConnection.RollbackTrans
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
   Case 0
      Unload Me
   Case 1
      PUB_RestorePrinter cmbPrinter
      DoPrint
      PUB_RestorePrinter strPrinter
   Case 2
      QueryData
   Case 3
      Process
   End Select
End Sub

Private Sub Process()
   Dim iRow As Integer, idx As Integer, iColId As Integer
   Dim stAppNo As String
   
   'Modified by Morgan 2015/5/20 改要先輸入申請案號
   stAppNo = UCase(InputBox("請輸入申請案號：", Me.Caption))
   If stAppNo = "" Then Exit Sub
   
   With MSHFlexGrid1
      iColId = GetFieldId("pa01")
      For iRow = 1 To .Rows - 1
         
         'If .TextMatrix(.row, 0) = "V" Then
         If stAppNo = .TextMatrix(iRow, 2) And .TextMatrix(iRow, 0) <> "X" Then
            .row = iRow
            
            If .TextMatrix(iRow, iColId) <> "P" Then
               MsgBox "非本所P案，無法作業！", vbExclamation
               Exit Sub
            End If
            
            strExc(1) = GetFormCode("P", GetValue(iRow, "案由"))
            If strExc(1) <> "" Then
               Select Case strExc(1)
               '實審通知日輸入
               Case 通知實審日, "1217"
                  idx = 1
               '核准函輸入
               Case 核准, "1906"
                  idx = 2
               '核駁函輸入
               Case 核駁
                  idx = 3
               '專利權消滅函輸入
               Case 專利權消滅
                  idx = 4
               'Added by Morgan 2023/1/12
               Case 專利證書
                  idx = 6
               'end 2023/1/12
               '異議/舉發受理函輸入
               Case "1803", "1804"
                  idx = 7
               Case Else
                  idx = 5
               End Select
               If strExc(1) = "9999" Then strExc(1) = "" '9999為不確定的案件性質,不帶入輸入畫面
               
               OpenForm idx, strExc(1)
            Else
               PopupMenu mdiMain.mnuPopEDoc
            End If
            Exit For
         End If
      Next
      If iRow = .Rows Then
         'MsgBox "請至少點選一筆資料！", vbExclamation
         MsgBox "申請案號輸入錯誤！", vbExclamation
      End If
      End With
End Sub

Private Sub cmdOpen_Click(Index As Integer)
   With MSHFlexGrid1
   For intI = 1 To .Rows - 1
      If .TextMatrix(intI, 0) = "V" Then
         Exit For
      End If
   Next
   
   If intI < .Rows Then
      RunPdf Index
   Else
      MsgBox "請至少點選一筆資料！", vbExclamation
   End If
   End With
End Sub

'Added by Morgan 2025/1/15
Private Sub Combo1_Click()
   If Combo1.Visible = False Then Exit Sub
   If Combo1.Tag <> Combo1 Then
      cmdOK(2).Value = True
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   
   PUB_SetPrinter Me.Name, cmbPrinter, strPrinter
   PUB_SetPrinter Me.Name, cmbPrinter2, strPrinter2
   
   'Modified by Morgan 2017/5/11
   'SetFileAssociation
   txtPDFPath = PUB_SetFileAssociation
   'end 2017/5/11
   
   m_AttachPath = App.path & "\" & Pub_GetSpecMan("EDocPath")
   
   'Added by Morgan 2025/1/15
   If strSrvDate(1) >= P業務區劃分啟用日 Then
      Combo1.Visible = True
      Label4.Visible = True
      Call SetPatentP12Combo(Combo1, "P", Label4)
   End If
   'end 2025/1/15
   
   QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.cmbPrinter.Text <> Me.cmbPrinter.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
   End If
   'Added by Morgan 2025/2/17
   If Me.cmbPrinter2.Text <> Me.cmbPrinter2.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter2.Name, "0", "0", Me.cmbPrinter2.Text
   End If
   'end 2025/2/17
   KillTemp
   Set oFileSys = Nothing
   Set oFile = Nothing
   Set frm04010516 = Nothing
End Sub

Private Sub KillTemp()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
End Sub

'Removed by Morgan 2019/7/9
'Private Function ReadTextFile(Optional pCharset As String = "big5") As String
'   Dim adoStream As ADODB.Stream
'   Dim var_String As Variant
'   Dim strFileName As String
'
'   strFileName = txtPath
'   Set adoStream = New ADODB.Stream
'   'adoStream.Charset = "UTF-8"
'   adoStream.Charset = pCharset
'   adoStream.Open
'   adoStream.LoadFromFile strFileName
'   ReadTextFile = adoStream.ReadText
'   adoStream.Close
'   Set adoStream = Nothing
'End Function

Private Function GetStr(ByVal pContent As String) As String
   'Modified by Morgan 2014/8/21
   'If Left(pContent, 1) = """" And Right(pContent, 1) = """" Then
   '   pContent = Trim(Mid(pContent, 2, Len(pContent) - 2))
   'End If
   ''目前有欄位前面多一個雙引號
   'If Left(pContent, 1) = """" Then
   '   pContent = Mid(pContent, 2)
   'End If
   '
   'If Right(pContent, 1) = """" Then
   '   pContent = Left(pContent, Len(pContent) - 1)
   'End If
   pContent = Trim(pContent)
   'Added by Morgan 2015/3/19
   '又發現最後多個逗號
   If Right(pContent, 1) = "," Then
      pContent = Left(pContent, Len(pContent) - 1)
   End If
   'end 2015/3/19
   Do While Left(pContent, 1) = """"
      pContent = Mid(pContent, 2)
   Loop
   Do While Right(pContent, 1) = """"
      pContent = Left(pContent, Len(pContent) - 1)
   Loop
   'end 2014/8/21
   'Modified by Morgan 2014/9/17 改讀檔後就去除
   'GetStr = Replace(pContent, vbTab, "")
   GetStr = pContent
   'end 2014/9/17
End Function

Private Sub QueryData()
   
   Dim stCon As String
   'Added by Morgan 2025/1/15
   Dim rsQuery As ADODB.Recordset
   Dim mSeqNo As String, stVTB0 As String
   'end 2025/1/15
   
   'Added by Morgan 2017/5/11
   '電子公文全面上線後只需看P案
   If strSrvDate(1) >= 電子公文啟用日 Then
      stCon = " and ed10='P' and (pa01='P' or pa01 is null)"
   End If
   'end 2017/5/11
   
   'Modified by Morgan 2017/1/19 舉發案要抓進度檔的對造 ex.102207894N02
   'Modified by Morgan 2017/2/9 舉發案進度檔抓不到再用前9碼抓基本檔(舉發發文只會有原申請號,受理後才有Nxx) ex.105302234N01
   'strExc(0) = "select ED20 V,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) 本所案號" & _
      ",ed02 申請案號,ed04 案由,sqldatet(to_char(ed05,'yyyymmdd')) 簽收日期,nvl(ed19,sqldatet(ed18)) 處理期限" & _
      ",sqldatet(ed08) 發文日期,ed17||'字第'||ed01||'號' 發文文號,ed10 案件種類,ed09 檔案" & _
      ",to_char(ed03,'yyyy/mm/dd hh24:mi:ss') 送達時間,ed07 受送達人,ed15 案號類別,ed16 案號" & _
      ",to_char(ed05,'yyyy/mm/dd hh24:mi:ss') 簽收時間,ed06 簽收人,ed01,ed17,ed18,ed19,ed20,pa01,pa02,pa03,pa04" & _
      " from EDocument,patent where ed11='C' and length(ed02)=9 and pa11(+)=ed02" & _
      " union select ED20 V,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) 本所案號" & _
      ",ed02 申請案號,ed04 案由,sqldatet(to_char(ed05,'yyyymmdd')) 簽收日期,nvl(ed19,sqldatet(ed18)) 處理期限" & _
      ",sqldatet(ed08) 發文日期,ed17||'字第'||ed01||'號' 發文文號,ed10 案件種類,ed09 檔案" & _
      ",to_char(ed03,'yyyy/mm/dd hh24:mi:ss') 送達時間,ed07 受送達人,ed15 案號類別,ed16 案號" & _
      ",to_char(ed05,'yyyy/mm/dd hh24:mi:ss') 簽收時間,ed06 簽收人,ed01,ed17,ed18,ed19,ed20,pa01,pa02,pa03,pa04" & _
      " from EDocument,caseprogress,patent where ed11='C' and length(ed02)>9 and cp36(+)=ed02 and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " order by 2,ed01"
   '衍生設計申請號超過9碼,舉發案改判斷來函申請號是否含N
   'Modified by Morgan 2018/1/16 +重新申請若未收新案號時舊申請號手動放CP30以便輸入來函 Ex.106144935(FCP-058047)
   'Modified by Morgan 2025/1/15 +ed05,PID
   strExc(0) = "select ED20 V,decode(pa01,'','非本所案件',pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04)) 本所案號" & _
      ",ed02 申請案號,ed04||decode(ed28,'副本','-'||ed28) 案由,sqldatet(to_char(ed05,'yyyymmdd')) 簽收日期,nvl(ed19,sqldatet(ed18)) 處理期限" & _
      ",sqldatet(ed08) 發文日期,ed17||'字第'||ed01||'號' 發文文號,ed10 案件種類,ed09 檔案" & _
      ",to_char(ed03,'yyyy/mm/dd hh24:mi:ss') 送達時間,ed07 受送達人,ed15 案號類別,ed16 案號" & _
      ",to_char(ed05,'yyyy/mm/dd hh24:mi:ss') 簽收時間,ed06 簽收人,ed01,ed17,ed18,ed19,ed20,pa01,pa02,pa03,pa04,trunc(ed05) ed05,'' PID" & _
      " from EDocument,(select ed02 C0,pa01 C1,pa02 C2,pa03 C3,pa04 C4 from EDocument,patent where ed11='C' and instr(ed02,'N')=0 and pa11(+)=ed02 and pa01 is not null" & _
      " union select ed02 C0,cp01 C1,cp02 C2,cp03 C3,cp04 C4 from EDocument,caseprogress where ed11='C' and instr(ed02,'N')=0 and cp30(+)=ed02 and cp09 is not null" & _
      " union select ed02 C0,cp01 C1,cp02 C2,cp03 C3,cp04 C4 from EDocument,caseprogress where ed11='C' and instr(ed02,'N')>0 and cp36(+)=ed02 and cp09 is not null" & _
      " union select ed02 C0,pa01 C1,pa02 C2,pa03 C3,pa04 C4 from EDocument,patent where ed11='C' and instr(ed02,'N')>0 and pa11(+)=substr(ed02,1,9) and pa01 is not null and not exists(select * from caseprogress where cp36=ed02)" & _
      ") TT,patent where ed11='C' and C0(+)=ed02 and pa01(+)=C1 and pa02(+)=C2 and pa03(+)=C3 and pa04(+)=C4 and ed30 is null" & stCon & _
      " order by trunc(ed05),2,ed01"
   'end 2017/2/9
   
   intI = 1
   lblCount = 0
   lblTotal = 0
   
   With MSHFlexGrid1
   .FixedCols = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   
   'Added by Morgan 2025/1/15
   If intI = 1 And strSrvDate(1) >= P業務區劃分啟用日 And Combo1 <> "" Then
      Combo1.Tag = ""
      Set rsQuery = PUB_CreateRecordset(RsTemp, , , 300, Me.Name, mSeqNo)
      With rsQuery
         .MoveFirst
         Do While Not .EOF
            .Fields("PID") = PUB_GetPHandler(.Fields("本所案號"))
            .MoveNext
         Loop
         .UpdateBatch
         
         stVTB0 = "select R001 as " & .Fields(0).Name
         For intI = 2 To .Fields.Count
            stVTB0 = stVTB0 & ", R" & Format(intI, "000") & " as " & .Fields(intI - 1).Name
         Next
         stVTB0 = stVTB0 & " from Rdatafactory Where Id='" & strUserNum & "' And Formname='" & Me.Name & "'  And Seqno='" & mSeqNo & "'"
      End With
      strSql = "Select X.* From (" & stVTB0 & ") X where PID='" & Left(Combo1, 5) & "' order by ed05,2,ed01"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      Combo1.Tag = Combo1
   End If
   'end 2025/1/15
   
   '若沒有資料時不可直接設定給 Grid 否則 MouseRow 會跑掉
   If intI = 1 Then
      Set .Recordset = RsTemp
      SetCmdEnabled True
      lblTotal = RsTemp.RecordCount
      SetGrid
   Else
      SetCmdEnabled False
      SetGrid True
   End If
   m_iCols = .Cols
   End With
   
   Set rsQuery = Nothing
End Sub

Private Sub SetGrid(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGrd1HeadWidth
   Dim iUbound As Integer
   Dim iRow As Integer

   arrGrd1HeadWidth = Array(250, 1140, 1200, 2100, 825, 825, 825, 3150, 825, 5200, 1600, 825, 825, 925)
   iUbound = UBound(arrGrd1HeadWidth)
   
   With MSHFlexGrid1
   .Visible = False
   If pReset = True Then
      .Clear
      .Rows = 2
      '.RowHeight(1) = 0
   Else
      For iRow = 1 To .Rows - 1
         If .TextMatrix(iRow, 0) = "N" Then
            .row = iRow
            For iCol = 0 To .Cols - 1
               .col = iCol
               .CellBackColor = &H80000018
            Next
         End If
      Next
   End If
   .FixedCols = 2
   .FormatString = "V|本所案號|申請案號|案由|簽收日期|處理期限|發文日期|發文文號|案件種類|檔案|送達時間|受送達人|案號類別|案號"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGrd1HeadWidth(iCol)
         .ColAlignment(iCol) = flexAlignLeftCenter
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   
   .Visible = True
   End With
End Sub

Private Sub RunPdf(iAct As Integer)
   Dim stFiles As String, stSavePath As String
   Dim stFileName As String
   Dim hLocalFile As Long
   Dim arrFileName() As String
   Dim idx As Integer, iRow As Integer
   
   Dim program_name As String
   Dim process_id As Long
   Dim process_handle As Long
   Dim strPrinterName As String
   Dim iCopys As Integer, ii As Integer
   Dim iColId As Integer
   
   '列印
   'If iAct = 1 Then 'Removed by Morgan 2025/3/4 不用限制，因Win11下開啟第2個檔案有時會沒顯示
      program_name = txtPDFPath
      strPrinterName = cmbPrinter
      '因為第 2 個以後開啟的 Reader 才會印完後自動關閉,所以固定先開一個空的程式,全部印完後再關閉
      process_id = Shell(program_name, vbHide)
      process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
   'End If
   
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         With MSHFlexGrid1
         
         'Added by Morgan 2014/6/4
         '一個發文號只要印一份
         iColId = GetFieldId("ed01")
         For ii = 1 To iRow - 1
            If .TextMatrix(ii, 0) = "V" Then
               If .TextMatrix(ii, iColId) = .TextMatrix(iRow, iColId) Then
                  Exit For
               End If
            End If
         Next
         If ii = iRow Then
         'end 2014/6/4
            stFiles = ""
            stSavePath = ""
            If GetAttachFile(GetValue(iRow, "ed01"), stFiles, stSavePath) = True Then
               arrFileName = Split(stFiles, ";")
               For idx = LBound(arrFileName) To UBound(arrFileName)
                  If arrFileName(idx) <> "" Then
                     stFileName = stSavePath & "\" & arrFileName(idx)
                     If iAct = 1 Then
                        'Added by Morgan 2025/2/17
                        PUB_WaitUntilNoJob cmbPrinter
                        PUB_WaitUntilNoJob cmbPrinter2
                        If InStr(UCase(arrFileName(idx)), ".CERT.") > 0 Then
                           PrintOnePdf program_name, " /n /t """ & stFileName & """ """ & cmbPrinter2 & """"
                        Else
                        'end 2025/2/17
                           PrintOnePdf program_name, " /n /t """ & stFileName & """ """ & strPrinterName & """"
                        End If
                        
                        
                        
                        'Modified by Morgan 2014/6/20
                        '103/7/1 起有存電子信函的只要印 1 份
                        If Val(strSrvDate(1)) < 20140701 Then
                           'Added by Morgan 2014/4/14 平行期間加印 北所 1 份,分所 2 份
                           If GetValue(iRow, "pa01") <> "" Then
                              iCopys = 1
                              'Modified by Morgan 2014/6/20 +特殊設定A7所有編號視為北所人員
                              'strExc(0) = "select st06 from caseprogress,staff where cp01='" & GetValue(iRow, "pa01") & "' and cp02='" & GetValue(iRow, "pa02") & "' and cp03='" & GetValue(iRow, "pa03") & "' and cp04='" & GetValue(iRow, "pa04") & "' and st01(+)=cp13 order by cp05 desc"
                              strExc(0) = "select DECODE(instr(';'||replace(oMan,',',';')||';',';'||ST01||';'),0,ST06,'1') from caseprogress,staff, SetSpecMan where cp01='" & GetValue(iRow, "pa01") & "' and cp02='" & GetValue(iRow, "pa02") & "' and cp03='" & GetValue(iRow, "pa03") & "' and cp04='" & GetValue(iRow, "pa04") & "' and st01(+)=cp13 and ocode(+)='A7' order by cp05 desc"
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                              If intI = 1 Then
                                 '分所
                                 If RsTemp.Fields(0) <> "1" Then
                                    iCopys = 2
                                 End If
                              End If
                              For ii = 1 To iCopys
                                 PrintOnePdf program_name, " /n /t """ & stFileName & """ """ & strPrinterName & """"
                              Next
                           End If
                           'end 2014/4/14
                        End If
                     Else
                        ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
                     End If
                  End If
               Next
            End If
         End If 'Added by Morgan 2014/6/4
         End With
      End If
   Next
   End With
   
   If iAct = 1 Then
      If process_handle <> 0 Then
         TerminateProcess process_handle, 0&
         CloseHandle process_handle
      End If
      MsgBox "列印完畢!"
   End If
End Sub

Private Sub PrintOnePdf(ByVal program_name As String, parameters As String)

Dim process_id As Long
Dim process_handle As Long
    ' Start the program.
    On Error GoTo ShellError

    process_id = Shell(program_name & parameters, vbHide)

    On Error GoTo 0

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If

    Exit Sub

ShellError:
    MsgBox " " & _
        program_name & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
End Sub


Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      
   Dim nCol As Long, nRow As Long, iRow As Integer
   Dim stValue As String
      
   If nCol < 0 Or nRow < 0 Then Exit Sub
   With MSHFlexGrid1
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow = 0 Then
      If nCol = 0 Then
         For iRow = 1 To .Rows - 1
            If .TextMatrix(iRow, 0) = "" Then
               stValue = "V"
               Exit For
            '已刪除資料標示為 X
            ElseIf .TextMatrix(iRow, 0) = "V" Then
               stValue = ""
               Exit For
            End If
         Next
         
         For iRow = 1 To .Rows - 1
            If .TextMatrix(iRow, 0) <> "X" Then
               If .TextMatrix(iRow, 0) <> stValue Then
                  .row = iRow
                  ClickGrid MSHFlexGrid1
               End If
            End If
         Next
      Else
         .col = nCol
         If m_blnColOrderAsc = False Then '字串降冪
            .Sort = 5 '字串昇冪
            m_blnColOrderAsc = True
         Else
            .Sort = 6 '字串降冪
            m_blnColOrderAsc = False
         End If
      End If
   Else
      .row = nRow
      ClickGrid MSHFlexGrid1
   End If
   .Visible = True
   End With
End Sub

Private Sub ClickGrid(grdDataList As MSHFlexGrid)
   Dim iCol As Integer

   With grdDataList
   If .TextMatrix(grdDataList.row, 1) <> "" Then
      If .TextMatrix(.row, 0) = "V" Then
         lblCount = Val(lblCount) - 1
         .TextMatrix(.row, 0) = ""
         For iCol = .FixedCols To .Cols - 1
            .col = iCol
            .CellBackColor = .BackColor
          Next
      '已刪除資料標示為 X
      ElseIf .TextMatrix(.row, 0) = "" Then
         lblCount = Val(lblCount) + 1
         .TextMatrix(.row, 0) = "V"
         For iCol = .FixedCols To .Cols - 1
            .col = iCol
            .CellBackColor = &HFFC0C0
         Next
      End If
   End If
   End With
End Sub

Private Function GetAttachFile(ByVal strCPP01 As String, Optional ByRef pFileName As String, Optional ByRef pSavePath As String) As Boolean
   Dim stAttPath As String
   Dim lngSize As Long
   Dim iFileNo As Integer
   Dim bytes() As Byte

On Error GoTo ErrHnd

   If pSavePath = "" Then
      '建立暫存資料夾
      If Dir(m_AttachPath, vbDirectory) = "" Then
         MkDir m_AttachPath
      End If
      pSavePath = m_AttachPath
   End If
   
'Modified by Morgan 2015/3/23 讀取檔案改呼叫共用函數(要改為FTP方式)
   strExc(0) = "select cpp01,cpp02 from casepaperpdf where cpp01='" & strCPP01 & "' and lower(substr(cpp02,-4))= '.pdf'" & IIf(pFileName <> "", " and cpp02='" & ChgSQL(pFileName) & "'", "") & " order by 2 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         stAttPath = pSavePath & "\" & .Fields("cpp02")
         If Dir(stAttPath) <> "" Then Kill stAttPath
         
'Modified by Morgan 2015/3/23 讀取檔案改呼叫共用函數(要改為FTP方式)
'         lngSize = Val(.Fields("cpp03").Value)
'         ReDim bytes(lngSize)
'         If lngSize > 0 Then
'            bytes() = .Fields("cpp04").GetChunk(lngSize)
'         End If
'
'         iFileNo = FreeFile
'         Open stAttPath For Binary Access Write As #iFileNo
'         If lngSize > 0 Then Put #iFileNo, , bytes()
'         Close #iFileNo
         If PUB_GetAttachFile_CPP(.Fields("cpp01"), .Fields("cpp02"), pSavePath) = False Then
            Exit Function
         End If
'end 2015/3/23
         pFileName = IIf(pFileName = "", "", pFileName & ";") & .Fields("cpp02")
         .MoveNext
      Loop
      End With
      GetAttachFile = True
   End If
   Exit Function

ErrHnd:

   MsgBox Err.Description, vbCritical
   If iFileNo > 0 Then Close #iFileNo
End Function

Private Sub DoPrint()

   Dim iOrientation As Integer, iRow As Integer, iCol As Integer
   Dim strTemp() As String
   Dim strSys As String, strCaseNo As String, iCaseNoId As Integer
   Dim iRecs As Integer, iCases As Integer
   Dim ii As Integer
   Dim iColId As Integer
   Dim bPaper As Boolean
   
   iOrientation = Printer.Orientation
   Printer.PaperSize = 9
   Printer.Orientation = 1
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   With MSHFlexGrid1
      '依本所案號排序
      iCol = .col
      iCaseNoId = GetFieldId("本所案號")
      iColId = GetFieldId("ed01")
      .col = iCaseNoId
      .Sort = 5
      .col = iCol
      
      GetPleft
      ReDim strTemp(m_iCols - 1)
      iPage = 1
      m_Sys = GetValue(1, "pa01")
      PrintPageHeader
      PrintPageHeader1
      iRecs = 0
      iCases = 0
      strCaseNo = ""
      For iRow = 1 To .Rows - 1
         .row = iRow
         strExc(1) = GetValue(iRow, "pa01")
         If m_Sys <> strExc(1) Then
            Call PrintReportFooter(iRecs, iCases)
            Printer.NewPage
            m_Sys = strExc(1)
            iRecs = 0
            iCases = 0
            strCaseNo = ""
            PrintPageHeader
            PrintPageHeader1
         End If
         
         'Added by Morgan 2014/6/4 總筆數改為統計發文號筆數
         'iRecs = iRecs + 1
         For ii = 1 To iRow - 1
            If .TextMatrix(ii, iColId) = .TextMatrix(iRow, iColId) Then
               Exit For
            End If
         Next
         If ii = iRow Then
            iRecs = iRecs + 1
         End If
         'end 2014/6/4
         
         For iCol = LBound(strTemp) To UBound(strTemp)
            strTemp(iCol) = .TextMatrix(iRow, iCol)
            '本所案號重複不印
            If iCol = iCaseNoId Then
               If .TextMatrix(iRow, iCaseNoId) <> "---" Then
                  If .TextMatrix(iRow, iCaseNoId) = strCaseNo Then
                     strTemp(iCol) = ""
                  Else
                     'Added by Morgan 2015/6/23
                     If m_Sys = "P" Then
                        If PUB_GetEMailFlag(Replace(strTemp(iCol), "-", ""), , , bPaper) = True And bPaper = False Then
                           strTemp(iCol) = strTemp(iCol) & "＊"
                        End If
                     End If
                     'end 2015/6/23
                     iCases = iCases + 1
                  End If
               'Added by Morgan 2014/8/5
               '非P案改印發文號
               Else
                  strTemp(iCol) = .TextMatrix(iRow, iColId)
               'end 2014/8/5
               End If
            End If
         Next
         PrintDetail strTemp
         strCaseNo = .TextMatrix(iRow, iCaseNoId)
      Next
      Call PrintReportFooter(iRecs, iCases)
      Printer.EndDoc
      MsgBox "列印完成！"
   End With
   Printer.Orientation = iOrientation
   
   
End Sub

Private Sub GetPleft()
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   intI = m_iCols + 1
   ReDim PLeft(1 To intI)
   PLeft(1) = ciStartX
   PLeft(2) = PLeft(1) + Printer.TextWidth(String(8, "　")) + ciColGap
   PLeft(3) = PLeft(2) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(4) = PLeft(3) + Printer.TextWidth(String(12, "　")) + ciColGap
   PLeft(5) = PLeft(4) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(6) = PLeft(5) + Printer.TextWidth(String(5, "　")) + ciColGap
End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)

   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      PrintLine
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      If bolSubtotal Then
         PrintPageHeader1
         iPrint = iPrint + lngLineHeight
      End If
   End If
    
End Sub

Private Sub PrintLine()
   Dim iNo As Integer
   iNo = (Printer.ScaleWidth - Printer.CurrentX - 500) \ Printer.TextWidth("-")
   Printer.Print String(iNo, "-")
End Sub

Sub PrintPageHeader()
   Dim strPTmp As String
   iPrint = ciStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strPTmp = "電子機關來函清單"
   If m_Sys = "P" Then
      strPTmp = strPTmp & "(P案)"
   Else
      strPTmp = strPTmp & "(非P案)"
   End If
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False

   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   PrintNewLine
   'Added by Morgan 2015/6/23
   If m_Sys = "P" Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      Printer.Print "＊E化案件"
   End If
   'end 2015/6/23
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
    
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   PrintLine
End Sub

Sub PrintPageHeader1()

    Call PrintNewLine(False, 1)
    For intI = 0 To m_iCols - 1
      Select Case intI
         'Modified by Morgan 2014/8/5
         'Case 1, 2, 3, 4, 5, 6
         Case 1
            Printer.CurrentX = PLeft(intI)
            Printer.CurrentY = iPrint
            If m_Sys = "P" Then
               Printer.Print MSHFlexGrid1.TextMatrix(0, intI)
            Else
               Printer.Print "發文號"
            End If
         Case 2, 3, 4, 5, 6
         'end 2014/8/5
            Printer.CurrentX = PLeft(intI)
            Printer.CurrentY = iPrint
            Printer.Print MSHFlexGrid1.TextMatrix(0, intI)
      End Select
    Next
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    PrintLine
    
End Sub

Sub PrintDetail(strData() As String)
    Dim iCol As Integer
    PrintNewLine
    For iCol = LBound(strData) To UBound(strData)
      Select Case iCol
         Case 1, 2, 3, 4, 5, 6
            Printer.CurrentX = PLeft(iCol)
            Printer.CurrentY = iPrint
            If iCol = 3 Then
               Printer.Print convForm(strData(iCol), 24)
            Else
               Printer.Print strData(iCol)
            End If
      End Select
    Next
End Sub

'列印表尾
Private Sub PrintReportFooter(ByVal iRecCount As Integer, Optional iCaseCount As Integer = 0)

    Call PrintNewLine(True, 1)
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    PrintLine
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print IIf(iCaseCount > 0, "本所案號：" & iCaseCount, "") & vbTab & vbTab & vbTab & IIf(iRecCount > 0, "公文：" & iRecCount, "")
    'Printer.EndDoc
End Sub

Private Sub SetCmdEnabled(pEnabled As Boolean)
   cmdOK(3).Enabled = pEnabled
   cmdOK(1).Enabled = pEnabled
   cmdOpen(0).Enabled = pEnabled
   cmdOpen(1).Enabled = pEnabled
   cmdDelete.Enabled = pEnabled And m_bDelete
End Sub

Private Function CheckStatus(pDocNo As String, Optional bolUnRec As Boolean = True) As Boolean
   strExc(0) = "select * from edocument where ed01='" & pDocNo & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp("ed11") > "C" Then
         If bolUnRec = False Then
            CheckStatus = True
         End If
      ElseIf bolUnRec = True Then
         CheckStatus = True
      End If
   End If
End Function

Public Sub OpenForm(Index As Integer, Optional pCP10 As String)
   Dim strDocNo As String, strAppNo As String, strRecDate As String, strDocWord As String, strDeadLine As String
   
   
   strDocNo = GetValue(MSHFlexGrid1.row, "ed01")
   If CheckStatus(strDocNo) = False Then
      MsgBox "發文文號 " & strDocNo & " 已輸入!!", vbExclamation
      Exit Sub
   End If
   
   strDocWord = GetValue(MSHFlexGrid1.row, "ed17")
   strAppNo = GetValue(MSHFlexGrid1.row, "申請案號")
   strRecDate = Replace(GetValue(MSHFlexGrid1.row, "簽收日期"), "/", "")
   If GetValue(MSHFlexGrid1.row, "ed19") <> "" Then
      '處理期間固定次日起且以月份計算
      'Modified by Morgan 2014/8/18 有30日的期限 Ex.P-71356 103/8/18
      'strDeadLine = Val(GetValue(MSHFlexGrid1.row, "ed19"))
      strDeadLine = GetValue(MSHFlexGrid1.row, "ed19")
   Else
      strDeadLine = GetValue(MSHFlexGrid1.row, "ed18")
   End If
   
   Select Case Index
      Case 1   '實審通知日輸入
         frm04010501.m_DocWord = strDocWord
         frm04010501.m_DocNo = strDocNo
         frm04010501.m_AppNo = strAppNo
         frm04010501.m_RDate = strRecDate
         frm04010501.Show
         
      Case 2   '核准函輸入
         frm04010502_1.m_DocWord = strDocWord
         frm04010502_1.m_DocNo = strDocNo
         frm04010502_1.m_AppNo = strAppNo
         frm04010502_1.m_RDate = strRecDate
         frm04010502_1.m_DeadLine = strDeadLine
         frm04010502_1.Show
         
      Case 3  '核駁函輸入
         frm04010503_1.m_DocWord = strDocWord
         frm04010503_1.m_DocNo = strDocNo
         frm04010503_1.m_AppNo = strAppNo
         frm04010503_1.m_RDate = strRecDate
         frm04010503_1.m_DeadLine = strDeadLine
         frm04010503_1.Show
         
      Case 4   '專利權消滅函輸入
         frm04010511_1.m_DocWord = strDocWord
         frm04010511_1.m_DocNo = strDocNo
         frm04010511_1.m_AppNo = strAppNo
         frm04010511_1.m_RDate = strRecDate
         frm04010511_1.Show
         
      Case 5   '一般來函輸入
         frm04010504_1.m_DocWord = strDocWord
         frm04010504_1.m_DocNo = strDocNo
         frm04010504_1.m_AppNo = strAppNo
         frm04010504_1.m_RDate = strRecDate
         frm04010504_1.m_DeadLine = strDeadLine
         frm04010504_1.m_NewCP10 = pCP10
         frm04010504_1.Show
         
      'Added by Morgan 2022/12/19
      Case 6   '證書號數輸入
         frm04010505_1.m_DocWord = strDocWord
         frm04010505_1.m_DocNo = strDocNo
         frm04010505_1.m_AppNo = strAppNo
         frm04010505_1.m_RDate = strRecDate
         frm04010505_1.Show
         
      Case 7   '異議/舉發受理函輸入
         frm04010506_1.m_DocWord = strDocWord
         frm04010506_1.m_DocNo = strDocNo
         frm04010506_1.m_AppNo = strAppNo
         frm04010506_1.m_RDate = strRecDate
         frm04010506_1.Show
   End Select
   
End Sub

Public Sub GoNext()
   Dim iRow As Integer, stED01 As String, iColId As Integer
   
   With MSHFlexGrid1
   If .Rows > 2 Then
      '上刪除標記,高度設零
      stED01 = GetValue(.row, "ed01")
      If CheckStatus(stED01, False) = True Then
         If .TextMatrix(.row, 0) = "V" Then
            lblCount = Val(lblCount) - 1
         End If
         
         .TextMatrix(.row, 0) = "X"
         .RowHeight(.row) = 0
         lblTotal = Val(lblTotal) - 1
         
         iColId = GetFieldId("ed01")
         For iRow = 1 To .Rows - 1
            '同發文號且不收文的一併上刪除標記
            'Modified by Morgan 2014/6/4 改發文號相同都上刪除標記(讓與案兩方都為本所客戶 or 兩個客戶舉發同一案 Ex. 申請號:095215810)
            'If GetValue(iRow, "ed01") = stED01 And .TextMatrix(iRow, 0) = "N" Then
            If .TextMatrix(iRow, iColId) = stED01 And .TextMatrix(iRow, 0) <> "X" Then
               .TextMatrix(iRow, 0) = "X"
               .RowHeight(iRow) = 0
               lblTotal = Val(lblTotal) - 1
            End If
         Next
      End If
      
      For iRow = 1 To .Rows - 1
         '改要先輸入申請號
         If .TextMatrix(iRow, 0) <> "X" Then
            cmdOK(3).Value = True
            Exit For
         End If
      Next
   Else
      SetGrid True
   End If
   End With
End Sub
'檢查案由與案件性質代碼
Private Function GetFormCode(pSys As String, pReason As String) As String
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   '抓確定案件性質碼者
   stSQL = "select em03 from EDocCodeMap where em01='" & pSys & "' and em02='" & pReason & "' and em03 is not null and em04 is null"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      GetFormCode = rsQuery(0)
   End If
   Set rsQuery = Nothing
End Function

'Added by Morgan 2014/9/3
Private Function CopyFile(pFromFolder As String, pToFolder As String, pFileList As String) As Boolean
   Dim arrFiles() As String
   Dim idx As Integer
   Dim stFromPath As String
   Dim stToPath As String
   
On Error GoTo ErrHnd

   arrFiles = Split(pFileList, ";")
   For idx = LBound(arrFiles) To UBound(arrFiles)
      If arrFiles(idx) <> "" Then
         stFromPath = pFromFolder & "\" & arrFiles(idx)
         stToPath = pToFolder & "\" & arrFiles(idx)
         If oFileSys.FileExists(stFromPath) = True Then
            'Modified by Morgan 2017/1/12
            'If oFileSys.FolderExists(pToFolder) = True Then
            If PUB_ChkDir(pToFolder) = True Then
               Set oFile = oFileSys.GetFile(stFromPath)
               oFile.Copy stToPath, True
            Else
               MsgBox "目的資料夾不存在！" & vbCrLf & vbCrLf & "[ " & pToFolder & " ]"
            End If
         Else
            MsgBox "來源檔案不存在！" & vbCrLf & vbCrLf & "[ " & stFromPath & " ]"
         End If
      End If
   Next
   CopyFile = True
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function

Private Function ChkPdfOK(pFileName As String) As Boolean
   Dim stTmpFile As String
   Dim iTimes As Integer
   Dim strCmd As String
   
On Error GoTo ErrHnd

   stTmpFile = ".\$$Check.pdf"
   
   '檢查合併檔是否正確
   If Dir(stTmpFile) <> "" Then Kill stTmpFile
   'Modified by Morgan 2014/5/9 合併程式改放執行檔路徑
   'strCmd = "pdftk.exe " & pFileName & " cat output " & stTmpFile
   strCmd = pub_PdftkEXE & " " & pFileName & " cat output " & stTmpFile
   Shell strCmd
   For iTimes = 1 To 10
      If PUB_CheckIsRunning(pub_PdftkName) = True Then
         Sleep 1000
      Else
         Exit For
      End If
   Next
   If iTimes > 10 Then Exit Function
   ChkPdfOK = True
   
ErrHnd:

End Function

Private Function GetValue(pRow As Integer, pFieldName As String) As String
   Dim ii As Integer
   With MSHFlexGrid1
   For ii = 0 To .Cols - 1
      If UCase(.TextMatrix(0, ii)) = UCase(pFieldName) Then
         GetValue = .TextMatrix(pRow, ii)
         Exit For
      End If
   Next
   End With
End Function

Private Function GetFieldId(pFieldName As String) As Integer
   Dim ii As Integer
   With MSHFlexGrid1
   For ii = 0 To .Cols - 1
      If UCase(.TextMatrix(0, ii)) = UCase(pFieldName) Then
         GetFieldId = ii
         Exit For
      End If
   Next
   End With
End Function
