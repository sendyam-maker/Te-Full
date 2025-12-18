VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frm02010412 
   BorderStyle     =   1  '單線固定
   Caption         =   "電子公文來函"
   ClientHeight    =   4416
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9132
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4416
   ScaleWidth      =   9132
   Begin VB.CommandButton cmdChgReason 
      Caption         =   "更正案由"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   4950
      TabIndex        =   19
      Top             =   540
      Width           =   915
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   585
      Left            =   90
      TabIndex        =   18
      Top             =   30
      Width           =   4065
      _ExtentX        =   7176
      _ExtentY        =   1037
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frm02010412.frx":0000
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "刪除電子公文"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   4230
      TabIndex        =   17
      Top             =   90
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定地址條"
      Height          =   585
      Left            =   90
      TabIndex        =   14
      Top             =   3840
      Width           =   5460
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   15
         Top             =   210
         Width           =   4635
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   16
         Top             =   225
         Width           =   765
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdChgDep 
      Caption         =   "轉部門"
      Height          =   400
      Left            =   5715
      TabIndex        =   13
      Top             =   90
      Width           =   870
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "重整(&Q)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6615
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
      Left            =   810
      TabIndex        =   0
      Top             =   660
      Width           =   4080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2775
      Left            =   90
      TabIndex        =   2
      Top             =   1020
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   4890
      _Version        =   393216
      Cols            =   15
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "V|本所案號|申請案號|註冊號|案由|簽收日期|處理期限|發文日期|發文文號|案件種類|檔案|送達時間|受送達人|案號類別|案號"
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
      _Band(0).Cols   =   15
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblTotal 
      Height          =   180
      Left            =   6435
      TabIndex        =   12
      Top             =   3900
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "總筆數："
      Height          =   180
      Index           =   0
      Left            =   5670
      TabIndex        =   11
      Top             =   3900
      Width           =   720
   End
   Begin VB.Label lblCount 
      Height          =   180
      Left            =   8145
      TabIndex        =   7
      Top             =   3900
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "已勾選筆數："
      Height          =   180
      Index           =   1
      Left            =   7020
      TabIndex        =   6
      Top             =   3900
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "印表機："
      Height          =   180
      Left            =   90
      TabIndex        =   1
      Top             =   720
      Width           =   720
   End
End
Attribute VB_Name = "frm02010412"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/13 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB
'Created by Morgan 2017/3/21
Option Explicit

Public m_bolUnloadPrint As Boolean '結束是否列印
Public m_strHiddenFormName As String
Public m_TM14 As String 'Added by Morgan 2023/6/14

Dim m_bDelete As Boolean
Dim m_Sys As String '系統別
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

Private Sub cmdChgDep_Click()
   Screen.MousePointer = vbHourglass
   If UpdateCheck = True Then
      FormUpdate
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Function UpdateCheck() As Boolean
   Dim iRow As Integer, iCount As Integer
   Dim strMsg As String
   
   iCount = 0
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         strExc(1) = GetValue(iRow, "TM01")
         'Modified by Morgan 2018/9/25 +T(T舊案註冊號可能與FCT案相同 Ex:FCT-004385,T-105672)
         If strExc(1) = "FCT" Or strExc(1) = "" Or strExc(1) = "T" Then
            iCount = iCount + 1
            strMsg = GetValue(iRow, "本所案號") & " (申請案號:" & GetValue(iRow, "申請案號") & ") "
         Else
            MsgBox "只有FCT案或非本所案才可轉部門！", vbCritical
            Exit Function
         End If
      End If
   Next
   End With
   
   If iCount = 0 Then
      MsgBox "請先勾選要轉部門的來函！", vbExclamation
      Exit Function
   ElseIf iCount = 1 Then
      If MsgBox(strMsg & "將轉到" & IIf(m_Sys = "FCT", "內商", "外商") & "，是否確定要繼續？" & vbCrLf & vbCrLf & "注意：請將卷交" & IIf(m_Sys = "FCT", "內商", "外商") & "！", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
         Exit Function
      End If
   Else
      If MsgBox("共有 " & iCount & " 筆電子來函紀錄將轉到" & IIf(m_Sys = "FCT", "內商", "外商") & "，是否確定要繼續？" & vbCrLf & vbCrLf & "注意：請將卷交" & IIf(m_Sys = "FCT", "內商", "外商") & "！", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
         Exit Function
      End If
   End If
   UpdateCheck = True
   
End Function

'轉部門
Private Function FormUpdate() As Boolean
   Dim iRow As Integer, stNo As String
   
On Error GoTo ErrHnd
   
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      .row = iRow
      If .TextMatrix(.row, 0) = "V" Then
         cnnConnection.BeginTrans
         
On Error GoTo ErrHndT
         
         stNo = GetValue(.row, "ed01")
         'Modified by Morgan 2017/10/13
         'strSql = "update edocument set ed30='T' where ed01='" & GetValue(.row, "ed01") & "'"
         strSql = "update edocument set ed30='" & IIf(m_Sys = "FCT", "T", "FCT") & "' where ed01='" & GetValue(.row, "ed01") & "'"
         'end 2017/10/13
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
   
   
   FormUpdate = True
   Exit Function
   
ErrHndT:
   cnnConnection.RollbackTrans
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   
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
         
         PUB_DelFtpFile2 stNo '必須在DB資料刪除前執行
         strSql = "delete casepaperpdf where cpp01='" & stNo & "'"
         cnnConnection.Execute strSql, intI
         
         'Modified by Morgan 2020/3/25
         'strSql = "delete edocument where ed01='" & stNo & "'"
         strSql = "update edocument set ed11='不收文' where ed01='" & stNo & "'"
         
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

Private Function DeleteCheck() As Boolean
   Dim iRow As Integer, iCount As Integer, idx As Integer
   Dim bConfirm As Boolean, bolNoPCase As Boolean


   iCount = 0
   With MSHFlexGrid1
   idx = GetFieldId("本所案號")
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, idx) = "非本所案件" Then
'         If .TextMatrix(iRow, 0) = "" Then
'            .row = iRow
'            ClickGrid MSHFlexGrid1
'         End If
      Else
'Removed by Morgan 2021/7/26 取消非本所案件限制，因外商會重複送件，但會有1次沒繳錢
'         If .TextMatrix(iRow, 0) = "V" Then
'            .row = iRow
'            ClickGrid MSHFlexGrid1
'         End If
      End If
      If .TextMatrix(iRow, 0) = "V" Then iCount = iCount + 1
   Next
   End With
   
   If iCount = 0 Then
      MsgBox "無案件可刪除！", vbExclamation
      Exit Function
   Else
      If MsgBox("共有 " & iCount & " 筆電子來函紀錄將刪除，是否確定要繼續？" & vbCrLf & vbCrLf & "注意：建議先列印出紙本後才可刪除！", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
         Exit Function
      End If
   End If
   DeleteCheck = True
   
End Function
'Added by Morgan 2018/10/4
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

Private Sub cmdOK_Click(Index As Integer)
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
   Dim stAppNo As String, bolDispute As Boolean
   Dim arrRow() As Integer, iMatchRows As Integer, stInputList As String
   
   stAppNo = UCase(InputBox("請輸入註冊號/申請案號：", Me.Caption))
   If stAppNo = "" Then Exit Sub
   
   With MSHFlexGrid1
      iColId = GetFieldId("tm01")
      iMatchRows = 0
      For iRow = 1 To .Rows - 1
         If (stAppNo = .TextMatrix(iRow, 2) Or stAppNo = .TextMatrix(iRow, 3)) And .TextMatrix(iRow, 0) <> "X" And .TextMatrix(iRow, 0) <> "N" Then
            .row = iRow
            If IsNull(.TextMatrix(iRow, iColId)) Then
               MsgBox "非本所商標案，無法作業！", vbExclamation
               Exit Sub
            End If
      'Added by Morgan 2017/11/29 若同一案件有多個來函時要可選擇
            iMatchRows = iMatchRows + 1
            ReDim Preserve arrRow(iMatchRows) As Integer
            arrRow(iMatchRows) = iRow
            stInputList = stInputList & iMatchRows & ": " & GetValue(iRow, "案由") & vbCrLf
         End If
      Next
      
      If iMatchRows > 1 Then
         Do
            strExc(0) = InputBox(stInputList & vbCrLf & "請輸入1~" & iMatchRows, "案由選擇", "1")
            If strExc(0) = "" Then
               Exit Sub
            Else
               intI = Val(strExc(0))
               If intI > 0 And intI <= iMatchRows Then
                  iRow = arrRow(intI)
                  Exit Do
               Else
                  
               End If
            End If
         Loop
      ElseIf iMatchRows = 1 Then
         iRow = arrRow(iMatchRows)
      End If
      
      If iRow > 0 And iRow < .Rows Then
         .row = iRow
      'end 2017/11/29
         
         'Added by Morgan 2021/6/16
         strExc(1) = GetValue(iRow, "註冊號")
         'T與FCT共同控管案件
         'Modified by Morgan 2022/1/14 增加案件,改抓系統特殊設定
         'If strExc(1) = "01922108" Or strExc(1) = "01922109" Then
         strExc(0) = Pub_GetSpecMan("T與FCT共同管控案件") 'Added by Morgan 2022/1/14
         If InStr(";" & strExc(0) & ";", ";" & strExc(1) & ";") > 0 Then
         'end 2022/1/14
            strExc(0) = "select tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04) CNo from trademark where tm15='" & strExc(1) & "' and tm10='000' and tm01 in ('T','FCT') and tm01<>'" & .TextMatrix(iRow, iColId) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "本案有另一案號 " & RsTemp(0) & " ，請特別注意！" & vbCrLf & vbCrLf & "※ 請於公文函上註記與 " & RsTemp(0) & " 共同管控案件", vbExclamation
            End If
         End If
         'end if
         
            strExc(1) = GetFormCode("T", GetValue(iRow, "案由"), bolDispute)
            'Added by Morgan 2017/6/27
            '爭議案非爭議案由都自行點選
            strExc(2) = GetValue(iRow, "tm28")
            If strExc(2) <> "1" And bolDispute = False Then
               strExc(1) = ""
            End If
            'end 2017/6/27
            If strExc(1) <> "" Then
               If m_Sys = "T" Then
                  If bolDispute = False Then
                     Select Case strExc(1)
                     '核准審定書
                     Case 1001
                        idx = 1
                     '核駁審定書
                     Case 1002
                        idx = 2
                     '審查報告
                     Case 1201, 1202, 1205
                        idx = 3
                     'Added by Morgan 2023/1/13
                     '註冊證
                     Case 1701
                        idx = 4
                     'end 2023/1/13
                     '非爭議案取消催審期限
                     Case 1705
                        idx = 5
                     '商標案被禁止處分
                     Case 1614, 1615
                        idx = 6
                     '延期受理
                     Case 1005
                        idx = 7
                     '其他來函
                     Case Else
                        idx = 8
                        If strExc(1) = "9999" Then strExc(1) = "" '9999為不確定的案件性質,不帶入輸入畫面
                     End Select
                     OpenForm 1, idx, strExc(1)
                  
                  Else
                     Select Case strExc(1)
                     '爭議案勝訴輸入
                     Case 1003
                        idx = 1
                     '爭議案敗訴輸入
                     Case 1004
                        idx = 2
                     '被異議/被評定/被撤銷/對方補充理由/對方延期/通知復審答辯
                     Case 1601, 1602, 1603, 1604, 1605, 1606, 1609, 1611, 1404
                        idx = 3
                     '發回補理由/發回補答辯
                     Case 1612, 1613
                        idx = 4
                     '撤銷原處分／和解輸入
                     Case 1402, 1407
                        idx = 5
                     '受理
                     Case 1607
                        idx = 6
                     '延長審查時間
                     Case 1401
                        idx = 7
                     '對方撤回
                     Case 1610
                        idx = 8
                     '延期受理
                     Case 1005
                        idx = 10
                     '部分勝部分敗
                     Case 1006
                        idx = 11
                     '其他來函
                     Case Else
                        idx = 9
                        If strExc(1) = "9999" Then strExc(1) = "" '9999為不確定的案件性質,不帶入輸入畫面
                     End Select
                     OpenForm 2, idx, strExc(1)
                     
                  End If
                  
               'FCT
               Else
                  Select Case strExc(1)
                  '核准審定書
                  Case 1001
                     idx = 1
                  '核駁審定書
                  Case 1002
                     idx = 2
                  '審查報告
                  Case 1201, 1202, 1205
                     idx = 3
                  'Added by Morgan 2023/1/13
                  '註冊證
                  Case 1701
                     idx = 4
                  'end 2023/1/13
                  '非爭議案取消催審期限
                  Case 1705
                     idx = 5
                  '商標案被禁止處分
                  Case 1614, 1615
                     idx = 6
                  '延期受理
                  Case 1005
                     idx = 7
                  '其他來函
                  Case Else
                     idx = 8
                     If strExc(1) = "9999" Then strExc(1) = "" '9999為不確定的案件性質,不帶入輸入畫面
                  End Select
                  OpenForm 3, idx, strExc(1)
                  
               End If
            Else
               If m_Sys = "T" Then
                  PopupMenu mdiMain.mnuPopEDoc
               Else
                  PopupMenu mdiMain.mnuPopEDoc03
               End If
            End If
            'Exit For 'Removed by Morgan 2017/11/29
         End If
      'Next 'Removed by Morgan 2017/11/29
      If iRow = .Rows Then
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

Private Sub Form_Load()
   MoveFormToCenter Me
   
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   cmdDelete.Visible = m_bDelete
   
   PUB_SetPrinter Me.Name, cmbPrinter, strPrinter
   PUB_SetPrinter Me.Name, Combo1
  
   m_PdfReader = PUB_SetFileAssociation
   m_AttachPath = App.path & "\" & Pub_GetSpecMan("EDocPath")
   
   If intPWhere = 國內 Then
      m_Sys = "T"
      'cmdChgDep.Visible = False '改T也可轉部門但限定非本所案號
   Else
      m_Sys = "FCT"
   End If
   QueryData
   setTextColor 'Added by Morgan 2018/9/18
End Sub

'Added by Morgan 2018/9/18
Private Sub setTextColor()
   Dim iStart As Integer, iEnd As Integer
   With RichTextBox1
      .SelStart = 0
      .SelLength = Len(.Text)
      .SelColor = vbBlue
      iStart = InStr(.Text, "非本所案件")
      .SelStart = iStart - 1
      .SelLength = 5
      .SelColor = vbRed
      iStart = InStr(.Text, "申請案號")
      .SelStart = iStart - 1
      .SelLength = 4
      .SelColor = vbRed
      iStart = InStr(.Text, "本所案號")
      .SelStart = iStart - 1
      .SelLength = 4
      .SelColor = vbRed
      .SelLength = 0
   End With
End Sub

'Added by Morgan 2017/5/3
Private Sub UnloadPrint(pPrinter As String)
   PUB_PrintCaseCloseSheet strUserNum
   '刪除暫存資料
   PUB_DeleteCaseCloseSheet strUserNum
   '列印地址條
   PUB_PrintAddressList strUserNum, pPrinter
   '刪除地址條列表資料
   PUB_DeleteAddressList strUserNum
   '初始化序號
   pub_AddressListSN = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_bolUnloadPrint Then
      UnloadPrint Me.Combo1.Text
   End If
   
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   If Me.cmbPrinter.Text <> Me.cmbPrinter.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
   End If
   
   KillTemp
   Set oFileSys = Nothing
   Set oFile = Nothing
   
   If m_strHiddenFormName <> "" Then
      unloadForm
   End If
   
   Set frm02010412 = Nothing
End Sub

Private Sub KillTemp()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
End Sub

Private Sub QueryData()
   Dim strCon As String, strVTB As String
   
   '內商: 1.轉部門到T 2.商爭案由 3.T案 4.無基本檔者
   If m_Sys = "T" Then
      strCon = " and nvl(ed30,'T')<>'FCT' and (ed30='T' or tm28<>'1' or nvl(tm01,'T')='T' or em07='Y')"
   
   '外商: FCT案且為非商爭案由且未轉部門
   Else
      strCon = " and (ed30='FCT' or (tm01='FCT' and tm28='1' and em07 is null and ed30 is null))"
   End If
   
   'Removed by Morgan 2017/7/28
   '測試期限加 ED30 is not null 條件以確保內專已列印紙本
   'If strSrvDate(1) < 電子公文啟用日 Then
   '   strCon = strCon & " and ed30 is not null"
   'End If
   
   'modify by sonia 2017/7/12 第三句抓cp30資料時,顯示tm12欄改為cp30,以便user看公文直接輸入,
   'strExc(0) = "select ED20 V,decode(TM01,null,'非本所案件',TM01||'-'||TM02||decode(TM03||TM04,'000','','-'||TM03||'-'||TM04)) 本所案號" & _
      ",nvl(tm12,ed02) 申請案號,nvl(tm15,decode(ed15,'註冊號',ed16)) 註冊號,ed04||decode(ed28,'副本','-'||ed28) 案由,sqldatet(to_char(ed05,'yyyymmdd')) 簽收日期,nvl(ed19,sqldatet(ed18)) 處理期限" & _
      ",sqldatet(ed08) 發文日期,ed17||'字第'||substr(ed01,1,11)||'號' 發文文號,ed10 案件種類,ed09 檔案" & _
      ",to_char(ed03,'yyyy/mm/dd hh24:mi:ss') 送達時間,ed07 受送達人,ed15 案號類別,ed16 案號" & _
      ",to_char(ed05,'yyyy/mm/dd hh24:mi:ss') 簽收時間,ed06 簽收人,ed01,ed17,ed18,ed19,ed20,TM01,TM02,TM03,TM04,TM28" & _
      " from (select ed01 No,tm01,tm02,tm03,tm04,tm12,tm15,tm28 from EDocument,trademark" & _
      " where ed11='C' and ed10='T' and tm12(+)=ed02" & _
      " and not exists(select * from caseprogress where cp30=ed02)" & _
      " and not exists(select * from trademark t where t.tm15=ed16)" & _
      " union select ed01 No,tm01,tm02,tm03,tm04,tm12,tm15,tm28 from EDocument,trademark" & _
      " where ed11='C' and ed10='T' and tm15(+)=ed16 and tm01 is not null" & _
      " union select ed01 No,tm01,tm02,tm03,tm04,tm12,tm15,tm28 from EDocument,caseprogress,trademark" & _
      " where ed11='C' and ed10='T' and cp30(+)=ed02 and cp30 is not null" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      "),edocument,edoccodemap" & _
      " where ed01(+)=No and em01(+)=ed10 and em02(+)=ed04" & _
      " and em01(+)=ed10 and em02(+)=ed04  " & strCon & _
      " order by trunc(ed05),2,ed01,ed20 desc"
   'Modified by Morgan 2018/8/22 註冊號加已核准條件
   'Modified by Morgan 2018/10/2 對造會放審定號 Ex:T-216798
   'Modified by Morgan 2018/10/9 本所號抓法改與電子公文維護相同
   'Modified by Morgan 2020/4/21 對造只抓AB類 Ex:T-226861(對造號01201023)
   'Modified by Morgan 2021/6/16 改用ED27串本所案號
   'strVTB = "select ed01 No,tm01,tm02,tm03,tm04,tm12,tm15,tm28 from EDocument,trademark" & _
      " where ed11='C' and ed10='T' and tm15(+)=ed16 and tm16='1' and tm57 is null" & _
      " union select ed01 No,tm01,tm02,tm03,tm04,tm12,tm15,tm28 from EDocument,trademark" & _
      " where ed11='C' and ed10='T' and tm12(+)=ed02 and tm01 is not null" & _
      " and not exists(select * from trademark t where tm15=ed16 and t.tm16='1' and t.tm57 is null)" & _
      " union select distinct ed01 No,tm01,tm02,tm03,tm04,cp30,tm15,tm28 from EDocument,caseprogress,trademark" & _
      " where ed11='C' and ed10='T' and cp30(+)=ed02 and cp30 is not null" & _
      " and not exists(select * from trademark t where tm15=ed16 and t.tm16='1' and t.tm57 is null)" & _
      " and not exists(select * from trademark t where tm12=ed02)" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      " union select distinct ed01 No,tm01,tm02,tm03,tm04,tm12,nvl(tm15,ed16) tm15,tm28 from EDocument,caseprogress,trademark" & _
      " where ed11='C' and ed10='T' and cp36(+)=ed16 and cp36 is not null and cp09<'C'" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      " and not exists(select * from trademark t where tm15=ed16 and t.tm16='1' and t.tm57 is null)" & _
      " and not exists(select * from trademark t where tm12=ed02)" & _
      " and not exists(select * from caseprogress t where t.cp30=ed02)"
   
   'strExc(0) = "select ED20 V,decode(TM01,null,'非本所案件',TM01||'-'||TM02||decode(TM03||TM04,'000','','-'||TM03||'-'||TM04)) 本所案號" & _
      ",nvl(tm12,ed02) 申請案號,tm15 註冊號,NVL(ed32,ed04||decode(ed28,'副本','-'||ed28)) 案由,sqldatet(to_char(ed05,'yyyymmdd')) 簽收日期,nvl(ed19,sqldatet(ed18)) 處理期限" & _
      ",sqldatet(ed08) 發文日期,ed17||'字第'||substr(ed01,1,11)||'號' 發文文號,ed10 案件種類,ed09 檔案" & _
      ",to_char(ed03,'yyyy/mm/dd hh24:mi:ss') 送達時間,ed07 受送達人,ed15 案號類別,ed16 案號" & _
      ",to_char(ed05,'yyyy/mm/dd hh24:mi:ss') 簽收時間,ed06 簽收人,ed01,ed17,ed18,ed19,ed20,TM01,TM02,TM03,TM04,TM28" & _
      " from edocument,(" & strVTB & ") X,edoccodemap" & _
      " where ed11='C' and ed10='T' and No(+)=ed01 and em01(+)=ed10 and em02(+)=ed04" & _
      " and em01(+)=ed10 and em02(+)=ed04  " & strCon & _
      " order by trunc(ed05),2,ed01,ed20 desc"
   
   strExc(0) = "select ED20 V,decode(TM01,null,'非本所案件',TM01||'-'||TM02||decode(TM03||TM04,'000','','-'||TM03||'-'||TM04)) 本所案號" & _
      ",nvl(tm12,ed02) 申請案號,tm15 註冊號,NVL(ed32,ed04||decode(ed28,'副本','-'||ed28)) 案由,sqldatet(to_char(ed05,'yyyymmdd')) 簽收日期,nvl(ed19,sqldatet(ed18)) 處理期限" & _
      ",sqldatet(ed08) 發文日期,ed17||'字第'||substr(ed01,1,11)||'號' 發文文號,ed10 案件種類,ed09 檔案" & _
      ",to_char(ed03,'yyyy/mm/dd hh24:mi:ss') 送達時間,ed07 受送達人,ed15 案號類別,ed16 案號" & _
      ",to_char(ed05,'yyyy/mm/dd hh24:mi:ss') 簽收時間,ed06 簽收人,ed01,ed17,ed18,ed19,ed20,TM01,TM02,TM03,TM04,TM28" & _
      " from edocument,trademark,edoccodemap" & _
      " where ed11='C' and ed10='T'" & _
      " and tm01(+)=substr(ed27,1,length(ed27)-9) and tm02(+)=substr(ed27,-9,6) and tm03(+)=substr(ed27,-3,1) and tm04(+)=substr(ed27,-2)" & _
      " and em01(+)=ed10 and em02(+)=ed04  " & strCon & _
      " order by trunc(ed05),2,ed01,ed20 desc"
   'end 2021/6/16
   
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

   arrGrd1HeadWidth = Array(250, 1140, 1200, 1200, 2100, 825, 825, 825, 3150, 825, 5200, 1600, 825, 825, 925)
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
   .FormatString = "V|本所案號|申請案號|註冊號|案由|簽收日期|處理期限|發文日期|發文文號|案件種類|檔案|送達時間|受送達人|案號類別|案號"
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
   If iAct = 1 Then
      program_name = m_PdfReader
      strPrinterName = cmbPrinter
      '因為第 2 個以後開啟的 Reader 才會印完後自動關閉,所以固定先開一個空的程式,全部印完後再關閉
      process_id = Shell(program_name, vbHide)
      process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
   End If
   
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
                        PrintOnePdf program_name, " /n /t """ & stFileName & """ """ & strPrinterName & """"
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
   strExc(0) = "select cpp01,cpp02 from casepaperpdf where cpp01='" & strCPP01 & "' and lower(substr(cpp02,-4))= '.pdf'" & IIf(pFileName <> "", " and cpp02='" & ChgSQL(pFileName) & "'", "")
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
   PLeft(4) = PLeft(3) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(5) = PLeft(4) + Printer.TextWidth(String(12, "　")) + ciColGap
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
   If m_Sys = "T" Then
      strPTmp = strPTmp & "(內商)"
   ElseIf m_Sys = "FCT" Then
      strPTmp = strPTmp & "(外商)"
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

Public Sub OpenForm(pType As Integer, Index As Integer, Optional pCP10 As String)

   Dim strDocNo As String, strAppNo As String, strRecDate As String, strDocWord As String, strDeadLine As String
   Dim strRegNo As String
   Dim oForm As Form
   Dim iStiu As Integer
   
   
   strDocNo = GetValue(MSHFlexGrid1.row, "ed01")
   strRegNo = GetValue(MSHFlexGrid1.row, "註冊號")
   If CheckStatus(strDocNo) = False Then
      MsgBox "發文文號 " & strDocNo & " 已輸入!!", vbExclamation
      Exit Sub
   End If

   strDocWord = GetValue(MSHFlexGrid1.row, "ed17")
   strAppNo = GetValue(MSHFlexGrid1.row, "申請案號")
   strRecDate = Replace(GetValue(MSHFlexGrid1.row, "簽收日期"), "/", "")
   If GetValue(MSHFlexGrid1.row, "ed19") <> "" Then
      '處理期間(前面是數字後面接單位(Ex. 3個月 30日)
      strDeadLine = GetValue(MSHFlexGrid1.row, "ed19")
   Else
      '處理期限(民國年月日)
      strDeadLine = GetValue(MSHFlexGrid1.row, "ed18")
   End If
   
   'Modify By Sindy 2018/11/14 商標系統拆內外商
   Call mdiMain.EDocSubOpenForm(pType, Index, oForm, iStiu)
'   'T商申
'   If pType = 1 Then
'
'      Select Case Index
'      Case 1   '非爭議案核准輸入
'         Set oForm = frm02010401_1
'
'      Case 2   '非爭議案核駁輸入
'         Set oForm = frm02010402_1
'
'      Case 3:   '審查報告輸入
'         Set oForm = frm02010403_1
'
'      'Removed by Morgan 2017/4/24 證書只會有紙本
'      'Case 4:   '註冊證輸入
'      '   Set oForm = frm02010404_1
'
'      Case 5:   '非爭議案取消催審期限
'         Set oForm = frm02010405_1
'
'      Case 6:   '商標案被禁止處分
'         Set oForm = frm02010406_1
'
'      Case 7:   '延期受理
'         Set oForm = frm02010407_1
'
'      Case 8:   '其他來函輸入
'         Set oForm = frm02010408_1
'
'      'Removed by Morgan 2017/4/24 主管機關不會是智慧局--秀玲
'      'Case 9:   '服務業務結果輸入
'      '   Set oForm = frm02010409_1
'      'Removed by Morgan 2017/4/24 大陸才有已經沒有用了--秀玲
'      'Case 10:   '廣告刊出來函輸入
'      '   Set oForm = frm02010410_1
'
'      Case 11:   '智慧局註冊費通知函輸入
'         Set oForm = frm02010411_1
'
'      End Select
'
'   'T商爭
'   ElseIf pType = 2 Then
'
'      Select Case Index
'      Case 1:   '爭議案勝訴輸入
'         Set oForm = frm02010501_1
'
'      Case 2:   '爭議案敗訴輸入
'         Set oForm = frm02010502_1
'
'      Case 3:   '被異議/被評定/被撤銷/對方補充理由/對方延期/通知復審答辯
'         Set oForm = frm02010503_1
'
'      Case 4:   '發回補理由/發回補答辯
'         Set oForm = frm02010504_1
'
'      Case 5:   '撤銷原處分／和解輸入
'         Set oForm = frm02010505_1
'
'      Case 6:   '受理
'         Set oForm = frm02010506_1
'
'      Case 7:   '延長審查時間
'         Set oForm = frm02010507_1
'
'      Case 8:   '對方撤回
'         Set oForm = frm02010508_1
'
'      Case 9:   '其他來函輸入
'         Set oForm = frm02010408_1
'
'      Case 10:   '延期受理
'         Set oForm = frm02010407_1
'
'      Case 11:   '部分勝部分敗
'         Set oForm = frm02010509_1
'
'      End Select
'
'   'FCT商申
'   Else
'
'      Select Case Index
'      Case 1   '非爭議案核准輸入
'         Set oForm = frm03020401_01
'
'      Case 2   '非爭議案核駁輸入
'         Set oForm = frm03020402_01
'
'      Case 3:   '審查報告輸入
'         Set oForm = frm03020403_01
'
'      Case 5:   '非爭議案取消催審期限
'         Set oForm = frm03020405_01
'         iStiu = 1
'
'      Case 6:   '商標案被禁止處分
'         Set oForm = frm03020406_01
'
'      Case 7:   '延期受理
'         Set oForm = frm03020407_01
'
'      Case 8:   '其他來函輸入
'         Set oForm = frm03020408_01
'         iStiu = 1
'
'      Case 12:   '通知已轉他所
'         Set oForm = frm03020405_01
'         iStiu = 2
'      End Select
'   End If
   
   If Not oForm Is Nothing Then
      
      With oForm
      .m_DocWord = strDocWord
      .m_DocNo = strDocNo
      .m_AppNo = strAppNo
      .m_RDate = strRecDate
      .m_DeadLine = strDeadLine
      .m_NewCP10 = pCP10
      .m_RegNo = strRegNo
      If iStiu > 0 Then
         .iStiu = iStiu
      End If
      'Added by Morgan 2023/6/15 註冊證要傳入前次的公告日
      'Modified by Morgan 2024/1/9 +pType = "1"
      'Modified by Morgan 2024/2/2
      'If pType = "1" And Index = 4 Then
      If (pType = "1" Or pType = "3") And Index = 4 Then
         .m_TM14 = m_TM14 '要在 from load 後設定否則會被清除
      End If
      'end 2023/6/15
      .Show
      m_strHiddenFormName = .Name
      End With
   End If
   
   Set oForm = Nothing
End Sub

Private Function setForm() As Boolean
   Dim oForm As Form
   
On Error Resume Next

   For Each oForm In Forms
      If oForm.Name = m_strHiddenFormName Then
         oForm.m_Done = False
         setForm = True
         Exit For
      End If
   Next
End Function

Private Sub unloadForm()
   Dim oForm As Form
   
   For Each oForm In Forms
      If oForm.Name = m_strHiddenFormName Then
         Unload oForm
         Set oForm = Nothing
         Exit For
      End If
   Next

End Sub

Public Sub GoNext()
   Dim iRow As Integer, stED01 As String, iColId As Integer
      
   PUB_SendMailCache 'Added by Morgan 2017/8/16
   
   If m_strHiddenFormName <> "" Then
      If setForm = False Then
         m_strHiddenFormName = ""
      End If
   End If
   
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
