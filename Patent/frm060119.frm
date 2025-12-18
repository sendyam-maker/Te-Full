VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060119 
   BorderStyle     =   1  '單線固定
   Caption         =   "電子公文來函"
   ClientHeight    =   4428
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9132
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4428
   ScaleWidth      =   9132
   Begin VB.CommandButton cmdOK 
      Caption         =   "來函清單"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   4
      Left            =   4200
      TabIndex        =   21
      Top             =   90
      Width           =   1005
   End
   Begin VB.TextBox txtEDocNo 
      Height          =   270
      Left            =   1440
      TabIndex        =   19
      Top             =   600
      Width           =   1725
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消歸卷"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3240
      TabIndex        =   18
      Top             =   540
      Width           =   915
   End
   Begin VB.CommandButton cmdAttach 
      Caption         =   "歸卷"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   5130
      TabIndex        =   17
      Top             =   540
      Width           =   780
   End
   Begin VB.CommandButton cmdChgReason 
      Caption         =   "更正案由"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6480
      TabIndex        =   16
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdDelete 
      Cancel          =   -1  'True
      Caption         =   "刪除(&D)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3240
      TabIndex        =   15
      Top             =   90
      Visible         =   0   'False
      Width           =   915
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "重整(&Q)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   2
      Left            =   5670
      TabIndex        =   10
      Top             =   90
      Width           =   780
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "輸入"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   7425
      TabIndex        =   9
      Top             =   90
      Width           =   690
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印清單"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5940
      TabIndex        =   8
      Top             =   540
      Width           =   1005
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "列印附件"
      Height          =   400
      Index           =   1
      Left            =   8010
      TabIndex        =   5
      Top             =   525
      Width           =   960
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "開啟附件"
      Height          =   400
      Index           =   0
      Left            =   6975
      TabIndex        =   4
      Top             =   525
      Width           =   1005
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   8145
      TabIndex        =   3
      Top             =   90
      Width           =   825
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   855
      Style           =   2  '單純下拉式
      TabIndex        =   0
      Top             =   4050
      Width           =   4710
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2955
      Left            =   60
      TabIndex        =   2
      Top             =   990
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   5207
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
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   900
      TabIndex        =   13
      Top             =   150
      Width           =   2310
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4075;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "智慧局發文號："
      Height          =   180
      Left            =   180
      TabIndex        =   20
      Top             =   630
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "管制人："
      Height          =   180
      Left            =   180
      TabIndex        =   14
      Top             =   210
      Width           =   720
   End
   Begin VB.Label lblTotal 
      Height          =   180
      Left            =   6435
      TabIndex        =   12
      Top             =   4110
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "總筆數："
      Height          =   180
      Index           =   0
      Left            =   5670
      TabIndex        =   11
      Top             =   4110
      Width           =   720
   End
   Begin VB.Label lblCount 
      Height          =   180
      Left            =   8145
      TabIndex        =   7
      Top             =   4110
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "已勾選筆數："
      Height          =   180
      Index           =   1
      Left            =   7020
      TabIndex        =   6
      Top             =   4110
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "印表機："
      Height          =   180
      Left            =   135
      TabIndex        =   1
      Top             =   4110
      Width           =   720
   End
End
Attribute VB_Name = "frm060119"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/29 Form2.0已修改
'Created by Morgan 2017/5/9
Option Explicit

Public m_bolUnloadPrint As Boolean '結束是否列印

Dim m_bDelete As Boolean
Dim m_AttachPath As String
'列印用
Dim PLeft() As Integer, iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500, ciColGap = 250
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim strPrinter As String
Dim m_iCols As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim oFileSys As New FileSystemObject
Dim oFile As File
Dim m_PdfReader As String


Private Sub cmdAttach_Click()
   With MSHFlexGrid1
   For intI = 1 To .Rows - 1
      If .TextMatrix(intI, 0) = "V" Then
         Exit For
      End If
   Next
   If intI < .Rows Then
      Set2RecNo
   Else
      MsgBox "請至少點選一筆資料！", vbExclamation
   End If
   End With
End Sub

'歸卷(不收文)
Private Sub Set2RecNo()
   Dim iRow As Integer, strMsg As String
   Dim stDocWord As String, stDocNo As String, stRecNo As String, stCP01 As String, stCP02 As String, stCP03 As String, stCP04 As String, stCP10 As String
   
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         strMsg = GetValue(iRow, "本所案號") & " 案的 """ & GetValue(iRow, "案由") & """ 來函" & vbCrLf & "是否確定要歸卷？"
         If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            stDocNo = GetValue(iRow, "ed01")
            stDocWord = GetValue(iRow, "ed17")
            stCP01 = GetValue(iRow, "pa01")
            stCP02 = GetValue(iRow, "pa02")
            stCP03 = GetValue(iRow, "pa03")
            stCP04 = GetValue(iRow, "pa04")
            stCP10 = ChangeTDateStringToTString(GetValue(iRow, "簽收日期")) & ".odoc" 'Modified by Morgan 2018/10/23 +副檔名
            Set frm060119_1.fmParent = Me
            frm060119_1.strPatent = stCP01 & stCP02 & stCP03 & stCP04
            frm060119_1.Caption = "請選擇要歸卷的收文號"
            frm060119_1.Show vbModal
            'Modified by Morgan 2018/10/23
            'stRecNo = Me.Tag
            stRecNo = Left(Me.Tag, 9)
            stCP10 = Mid(Me.Tag, 11) & "." & stCP10
            'end 2018/10/23
            If stRecNo <> "" Then
               cnnConnection.BeginTrans
On Error GoTo ErrHnd

               PUB_UpdateEdocRec stDocNo, stRecNo, stCP01, stCP02, stCP03, stCP04, stCP10
               'Added by Morgan 2017/6/29 機關文號也要寫入歸卷的收文號(放最新的)--葉敏莉
               strSql = "update caseprogress set cp08='" & stDocWord & "字第" & stDocNo & "號' where cp09='" & stRecNo & "'"
               Pub_SeekTbLog strSql 'Added by Morgan 2019/2/22
               cnnConnection.Execute strSql, intI
               
               .TextMatrix(iRow, 0) = "X"
               .RowHeight(iRow) = 0
               cnnConnection.CommitTrans
               
               lblCount = Val(lblCount) - 1
               lblTotal = lblTotal - 1
               DoEvents
               
On Error GoTo ErrHnd2
            Else
               Exit For
            End If
         Else
            Exit For
         End If
      End If
   Next
   End With
   Exit Sub
   
ErrHnd:
   cnnConnection.RollbackTrans

ErrHnd2:
   MsgBox Err.Description, vbCritical
   
End Sub
'Added by Morgan 2019/2/22
Private Sub cmdCancel_Click()
   Dim iErr As Integer
   Dim bolOK As Boolean
   
   If txtEDocNo = "" Then
      MsgBox "請輸入智慧局發文號！", vbExclamation
      txtEDocNo.SetFocus
      Exit Sub
   End If
   
   strExc(0) = "select ed01,sqldatet(ed13) ed13,NVL(ed32,ed04||decode(ed28,'副本','-'||ed28)) 案由" & _
      ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CaseNo" & _
      ",sqldatet(cp27) cp27,cpm03 from edocument,caseprogress,casepropertymap" & _
      " where ed01='" & txtEDocNo & "' and cp09<'C' and cp09(+)=ed11" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strExc(1) = "本所案號：　　" & RsTemp("CaseNo") & vbCrLf & _
                  "發文日：　　　" & RsTemp("cp27") & vbCrLf & _
                  "案件性質：　　" & RsTemp("cpm03") & vbCrLf & vbCrLf & _
                  "智慧局發文號：" & RsTemp("ed01") & vbCrLf & _
                  "簽收日期：　　" & RsTemp("ed13") & vbCrLf & _
                  "案由：　　　　" & RsTemp("案由") & vbCrLf & vbCrLf & _
                  "是否確定要取消歸卷？"
      If MsgBox(strExc(1), vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   Else
      MsgBox "歸卷資料讀取失敗！", vbCritical
      Exit Sub
   End If
   
On Error GoTo ErrHnd:
   cnnConnection.BeginTrans

   iErr = 1
   strSql = "update casepaperpdf set cpp01='" & txtEDocNo & "',cpp02='$" & txtEDocNo & ".pdf',cpp10='U'" & _
      " where cpp01=( select ed11 from edocument where ed01='" & txtEDocNo & "')" & _
      " and exists(select * from edocument where ed01='" & txtEDocNo & "' and cpp02 like '%.'||(ed13-19110000)||'.%')"
   cnnConnection.Execute strSql, intI
   If intI <> 1 Then GoTo ErrHnd
   
   iErr = 2
   '清除機關文號
   strSql = "update caseprogress set cp08='' where (cp09,cp08) in (select ed11,ed17||'字第'||ed01||'號' from edocument where ed01='" & txtEDocNo & "')"
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   
   strSql = "update edocument set ed11='C' where ed01='" & txtEDocNo & "'"
   cnnConnection.Execute strSql, intI
   If intI <> 1 Then GoTo ErrHnd
      
   cnnConnection.CommitTrans
   MsgBox "歸卷已取消！", vbInformation
   txtEDocNo = ""
   cmdOK(2).Value = True
   Exit Sub
   
ErrHnd:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   Else
      strExc(0) = "歸卷取消失敗！"
      If iErr = 1 Then
         If intI = 0 Then
            strExc(0) = strExc(0) & "(卷宗區找不到已歸卷的pdf檔)"
         Else
            strExc(0) = strExc(0) & "(卷宗區找到超過1個已歸卷的pdf檔)"
         End If
      Else
         strExc(0) = strExc(0) & "(找不到符合的電子公文紀錄！)"
      End If
      MsgBox strExc(0), vbCritical
   End If
End Sub

'Added by Morgan 2019/2/22
Private Sub cmdChgReason_Click()
   Dim iIdx As Integer, intR As Integer
   
   With MSHFlexGrid1
   For intI = 1 To .Rows - 1
      If .TextMatrix(intI, 0) = "V" Then
         Exit For
      End If
   Next
   
   If intI < .Rows Then
      iIdx = GetFieldId("案由")
      strExc(1) = .TextMatrix(intI, iIdx)
      strExc(0) = InputBox("請輸入正確的案由：", "更正案由", strExc(1))
      If strExc(0) <> "" And strExc(0) <> strExc(1) Then
         strExc(2) = GetValue(intI, "ed01")
         strSql = "update edocument set ed32='" & ChgSQL(strExc(0)) & "' where ed01='" & strExc(2) & "'"
         cnnConnection.Execute strSql, intR
         If intR > 0 Then
            .TextMatrix(intI, iIdx) = strExc(0)
         End If
      End If
   Else
      MsgBox "請至少點選一筆資料！", vbExclamation
   End If
   End With
End Sub

Private Sub cmdDelete_Click()
   Screen.MousePointer = vbHourglass
   If DeleteCheck = True Then
      FormDelete
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Function DeleteCheck() As Boolean
   Dim iRow As Integer, iCount As Integer, idx As Integer
   Dim bConfirm As Boolean, bolNoPCase As Boolean


   iCount = 0
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then iCount = iCount + 1
   Next
   End With
   
   If iCount = 0 Then
      MsgBox "請勾選要刪除的資料！", vbExclamation
      Exit Function
   Else
      If MsgBox("共有 " & iCount & " 筆電子來函紀錄將刪除，是否確定要繼續？" & vbCrLf & vbCrLf & "注意：請先列印出紙本後才可刪除！", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
         Exit Function
      End If
   End If
   DeleteCheck = True
   
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
   'Added by Morgan 2020/4/10
   Case 4 '來函清單
      PUB_RestorePrinter cmbPrinter
      DoPrint2
      PUB_RestorePrinter strPrinter
   End Select
End Sub

'Added by Morgan 2020/4/10
Private Sub DoPrint2()
   Dim strDate As String
   strDate = strSrvDate(1)
   
   strDate = InputBox("請輸入電子公文來函日期！", "電子公文來函清單列印", strDate)
   If strDate <> "" Then
      If ChkDate(strDate) = True Then
         frm010027.Hide
         frm010027.cmbPrinter = Me.cmbPrinter
         frm010027.m_bCalled = True
         frm010027.ReportFCP strDate, , Trim(Left(Combo1.Text, 5))
         Unload frm010027
      End If
   End If
End Sub

Private Sub Process()
   Dim iRow As Integer, idx As Integer, iColId As Integer
   Dim stAppNo As String, iSRow As Integer
   
   stAppNo = UCase(InputBox("請輸入申請案號：", Me.Caption))
   If stAppNo = "" Then Exit Sub
   
   iColId = GetFieldId("pa01")
   'Modified by Morgan 2025/3/19 改有勾選的優先(一案同日有多來函時可方便選擇輸入)
   iSRow = 0
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      If stAppNo = .TextMatrix(iRow, 2) Then
         If .TextMatrix(iRow, 0) = "V" Then
            iSRow = iRow
            Exit For
         ElseIf .TextMatrix(iRow, 0) <> "X" Then
            iSRow = iRow
         End If
      End If
   Next
   End With
   
   If iSRow > 0 Then
      MSHFlexGrid1.row = iSRow '不可省略，後續會用現行位置抓資料
      strExc(1) = GetFormCode("P", GetValue(iSRow, "案由"))
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
         'Added by Morgan 2023/1/16
         Case 專利證書
            idx = 6
         '異議/舉發受理函輸入
         Case "1803", "1804"
            idx = 8
         Case Else
            idx = 5
         End Select
         If strExc(1) = "9999" Then strExc(1) = "" '9999為不確定的案件性質,不帶入輸入畫面
         
         OpenForm idx, strExc(1)
      Else
         PopupMenu mdiMain.mnuPopEDoc
      End If
   Else
      MsgBox "申請案號輸入錯誤！", vbExclamation
   End If
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

Private Sub Combo1_Click()
   If Me.Visible Then
      If Trim(Left(Combo1.Text, 5)) <> Combo1.Tag Then QueryData
   End If
   
   'Added by Morgan 2020/4/10
   If Trim(Left(Combo1.Text, 5)) = "" Then
      cmdOK(4).Enabled = False
   Else
      cmdOK(4).Enabled = True
   End If
   'end 2020/4/10
   
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   If m_bDelete Then cmdDelete.Visible = True
   PUB_SetPrinter Me.Name, cmbPrinter, strPrinter
  
   m_PdfReader = PUB_SetFileAssociation
   m_AttachPath = App.path & "\" & Pub_GetSpecMan("EDocPath")
   
   AddList
   QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   '若印表機變動, 則更新列印設定
   If Me.cmbPrinter.Text <> Me.cmbPrinter.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
   End If
   
   KillTemp
   Set oFileSys = Nothing
   Set oFile = Nothing
   Set frm060119 = Nothing
End Sub

Private Sub AddList()
   Combo1.Clear
   'Modified by Morgan 2017/6/2 改抓國家檔FCP管制人(原抓在職)
   strExc(0) = "select st01,st02 from staff a where st03='F22' and exists(select * from nation where na16=st01) order by 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not RsTemp.EOF
         If .Fields("st01") = strUserNum Then
            Combo1.AddItem .Fields("st01") & " " & .Fields("st02"), 0
            Combo1.Tag = strUserNum
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
End Sub

Private Sub KillTemp()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
End Sub

Private Sub QueryData()
   Dim strCon As String
   
   strCon = ""
   Combo1.Tag = Trim(Left(Combo1.Text, 5))
   If Combo1.Tag <> "" Then
      '不調卷的公文統一由整理來函的人員輸入
      'Removed by Morgan 2020/2/10 改各區自行處理
      'If Combo1.Tag = Pub_GetSpecMan("FCP來函整理人員") Then
      '   strCon = " or em07='N'"
      'Else
      '   strCon = " and em07 is null"
      'End If
      'end 2020/2/10
      
      strCon = " and ( na16='" & Left(Combo1.Text, 5) & "'" & strCon & ")"
   End If
   
   '測試期限加 ED30 is not null 條件以確保內專已列印紙本
   If strSrvDate(1) < 電子公文啟用日 Then
      strCon = strCon & " and ed30 is not null"
   End If
   
   '衍生設計申請號超過9碼,舉發案改判斷來函申請號是否含N
   'Modified by Morgan 2018/1/16 +重新申請若未收新案號時舊申請號手動放CP30以便輸入來函 Ex.106144935(FCP-058047)
   'modify by sonia 2024/7/4 因為申請案號095117221之電子公文，將倒數第3行之pa01='FCP'改為(pa01='FCP' or ed30='FCP')
   strExc(0) = "select ED20 V,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) 本所案號" & _
      ",ed02 申請案號,NVL(ed32,ed04||decode(ed28,'副本','-'||ed28)) 案由,sqldatet(to_char(ed05,'yyyymmdd')) 簽收日期,nvl(ed19,sqldatet(ed18)) 處理期限" & _
      ",sqldatet(ed08) 發文日期,ed17||'字第'||ed01||'號' 發文文號,ed10 案件種類,ed09 檔案" & _
      ",to_char(ed03,'yyyy/mm/dd hh24:mi:ss') 送達時間,ed07 受送達人,ed15 案號類別,ed16 案號" & _
      ",to_char(ed05,'yyyy/mm/dd hh24:mi:ss') 簽收時間,ed06 簽收人,ed01,ed17,ed18,ed19,ed20,pa01,pa02,pa03,pa04" & _
      " from EDocument,(select ed02 C0,pa01 C1,pa02 C2,pa03 C3,pa04 C4 from EDocument,patent where ed11='C' and ed10='P' and instr(ed02,'N')=0 and pa11(+)=ed02 and pa01='FCP'" & _
      " union select ed02 C0,cp01 C1,cp02 C2,cp03 C3,cp04 C4 from EDocument,caseprogress where ed11='C' and ed10='P' and instr(ed02,'N')=0 and cp30(+)=ed02 and cp01='FCP'" & _
      " union select ed02 C0,cp01 C1,cp02 C2,cp03 C3,cp04 C4 from EDocument,caseprogress where ed11='C' and ed10='P' and instr(ed02,'N')>0 and cp36(+)=ed02 and cp01='FCP'" & _
      " union select ed02 C0,pa01 C1,pa02 C2,pa03 C3,pa04 C4 from EDocument,patent where ed11='C' and ed10='P' and instr(ed02,'N')>0 and pa11(+)=substr(ed02,1,9) and pa01='FCP' and not exists(select * from caseprogress where cp36=ed02 and cp01='FCP')" & _
      ") TT,patent,fagent,nation,edoccodemap where ed11='C' and ed10='P' and C0(+)=ed02 and pa01(+)=C1 and pa02(+)=C2 and pa03(+)=C3 and pa04(+)=C4 and (pa01='FCP' or ed30='FCP') and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9) and na01(+)=fa10" & _
      " and em01(+)=ed10 and em02(+)=ed04" & strCon & _
      " order by trunc(ed05),em07,2,ed01"
   'end 2017/2/9
   
   intI = 1
   lblCount = 0
   lblTotal = 0
   
   With MSHFlexGrid1
   .FixedCols = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
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
   Dim ii As Integer
   Dim iColId As Integer
   
   '列印
   If iAct = 1 Then
      program_name = m_PdfReader
      strPrinterName = cmbPrinter
      '因為第 2 個以後開啟的 Reader 才會印完後自動關閉,所以固定先開一個空的程式,全部印完後再關閉
      process_id = SHELL(program_name, vbHide)
      process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
   End If
   
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
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
            stFiles = ""
            stSavePath = ""
            If GetAttachFile(GetValue(iRow, "ed01"), stFiles, stSavePath) = True Then
               arrFileName = Split(stFiles, ";")
               For idx = LBound(arrFileName) To UBound(arrFileName)
                  If arrFileName(idx) <> "" Then
                     stFileName = stSavePath & "\" & arrFileName(idx)
                     If iAct = 1 Then
                        PrintOnePdf program_name, " /n /t """ & stFileName & """ """ & strPrinterName & """"
                     Else
                        ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
                     End If
                  End If
               Next
            End If
         End If
      End If
   Next
   End With
   
   If iAct = 1 Then
      TerminateProcess process_handle, 0&
      CloseHandle process_handle
      MsgBox "列印完畢!"
   End If
End Sub

Private Sub PrintOnePdf(ByVal program_name As String, parameters As String)

Dim process_id As Long
Dim process_handle As Long
    ' Start the program.
    On Error GoTo ShellError

    process_id = SHELL(program_name & parameters, vbHide)

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
   
On Error GoTo ErrHnd

   If pSavePath = "" Then
      '建立暫存資料夾
      If Dir(m_AttachPath, vbDirectory) = "" Then
         MkDir m_AttachPath
      End If
      pSavePath = m_AttachPath
   End If

   strExc(0) = "select cpp01,cpp02 from casepaperpdf where cpp01='" & strCPP01 & "' and lower(substr(cpp02,-4))= '.pdf'" & IIf(pFileName <> "", " and cpp02='" & ChgSQL(pFileName) & "'", "") & " order by 2"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         stAttPath = pSavePath & "\" & .Fields("cpp02")
         If Dir(stAttPath) <> "" Then Kill stAttPath

         If PUB_GetAttachFile_CPP(.Fields("cpp01"), .Fields("cpp02"), pSavePath) = False Then
            Exit Function
         End If
         pFileName = IIf(pFileName = "", "", pFileName & ";") & .Fields("cpp02")
         .MoveNext
      Loop
      End With
      GetAttachFile = True
   End If
   Exit Function

ErrHnd:

   MsgBox Err.Description, vbCritical
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
      PrintPageHeader
      PrintPageHeader1
      iRecs = 0
      iCases = 0
      strCaseNo = ""
      For iRow = 1 To .Rows - 1
         .row = iRow
         '發文號筆數
         For ii = 1 To iRow - 1
            If .TextMatrix(ii, iColId) = .TextMatrix(iRow, iColId) Then
               Exit For
            End If
         Next
         If ii = iRow Then
            iRecs = iRecs + 1
         End If

         For iCol = LBound(strTemp) To UBound(strTemp)
            strTemp(iCol) = .TextMatrix(iRow, iCol)
            '本所案號重複不印
            If iCol = iCaseNoId Then
               If .TextMatrix(iRow, iCaseNoId) <> "---" Then
                  If .TextMatrix(iRow, iCaseNoId) = strCaseNo Then
                     strTemp(iCol) = ""
                  Else
                     iCases = iCases + 1
                  End If
               '非本所案改印發文號
               Else
                  strTemp(iCol) = .TextMatrix(iRow, iColId)
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
   strPTmp = "電子機關來函清單(FCP)"
   
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
    For intI = 1 To 6
      Printer.CurrentX = PLeft(intI)
      Printer.CurrentY = iPrint
      Printer.Print MSHFlexGrid1.TextMatrix(0, intI)
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
   Dim strDocDate As String
   Dim oForm As Form
   
   strDocNo = GetValue(MSHFlexGrid1.row, "ed01")
   If CheckStatus(strDocNo) = False Then
      MsgBox "發文文號 " & strDocNo & " 已輸入!!", vbExclamation
      Exit Sub
   End If
   
   strDocWord = GetValue(MSHFlexGrid1.row, "ed17")
   strAppNo = GetValue(MSHFlexGrid1.row, "申請案號")
   strRecDate = Replace(GetValue(MSHFlexGrid1.row, "簽收日期"), "/", "")
   strDocDate = Replace(GetValue(MSHFlexGrid1.row, "發文日期"), "/", "")
   If GetValue(MSHFlexGrid1.row, "ed19") <> "" Then
      strDeadLine = GetValue(MSHFlexGrid1.row, "ed19")
   Else
      strDeadLine = GetValue(MSHFlexGrid1.row, "ed18")
   End If
   
   Select Case Index
      Case 1   '實審通知日輸入
         Set oForm = frm06010601

      Case 2   '核准函輸入
         Set oForm = frm06010602_1
         
      Case 3  '核駁函輸入
         Set oForm = frm06010603_1
         
      Case 4   '專利權消滅函輸入
         Set oForm = frm06010608_1

      Case 5   '一般來函輸入
         Set oForm = frm06010604_1
      'Added by Morgan 2023/1/16
      Case 6   '證書號數輸入
         Set oForm = frm06010605_1
      'end 2023/1/16
      Case 8   '異議/舉發受理函輸入
         Set oForm = frm06010606_1
         
   End Select
   
   If Not oForm Is Nothing Then
      With oForm
      .m_DocWord = strDocWord
      .m_DocNo = strDocNo
      .m_DocDate = strDocDate
      .m_AppNo = strAppNo
      .m_RDate = strRecDate
      If strDeadLine <> "" Then
         .m_DeadLine = strDeadLine
      End If
      If pCP10 <> "" Then
         .m_NewCP10 = pCP10
      End If
      .Show
      End With
      Set oForm = Nothing
   End If
End Sub

Public Sub GoNext()
   Dim iRow As Integer, stED01 As String, iColId As Integer
   
   PUB_SendMailCache 'Added by Morgan 2017/8/16
   
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
            'cmdOK(3).Value = True
            Exit For
         End If
      Next
   Else
      SetGrid True
   End If
   End With
End Sub
'檢查案由與案件性質代碼
Private Function GetFormCode(pSys As String, pReason As String, Optional pDispute As Boolean) As String
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   '抓確定案件性質碼者
   stSQL = "select em03,em07 from EDocCodeMap where em01='" & pSys & "' and em02='" & pReason & "' and em03 is not null and em04 is null"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      GetFormCode = rsQuery(0)
      If rsQuery(1) = "Y" Then pDispute = True
   End If
   Set rsQuery = Nothing
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
         
         PUB_DelFtpFile2 stNo '檔案放 FTP,必須在DB資料刪除前執行
         strSql = "delete casepaperpdf where cpp01='" & stNo & "'"
         cnnConnection.Execute strSql, intI
         strSql = "delete edocument where ed01='" & stNo & "'"
         cnnConnection.Execute strSql, intI
         
         
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
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function

Private Sub txtEDocNo_GotFocus()
   TextInverse txtEDocNo
End Sub
