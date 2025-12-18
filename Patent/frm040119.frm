VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040119 
   BorderStyle     =   1  '單線固定
   Caption         =   "指示信判發作業"
   ClientHeight    =   5820
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8928
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8928
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "點我展開"
      Height          =   345
      Left            =   4410
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   0
      Width           =   4515
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5415
      Left            =   4410
      TabIndex        =   6
      Top             =   330
      Width           =   4515
      ExtentX         =   7964
      ExtentY         =   9551
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   3825
      Top             =   450
      _ExtentX        =   339
      _ExtentY        =   339
      _Version        =   393216
      Picture         =   "frm040119.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "結束"
      Height          =   345
      Left            =   3600
      TabIndex        =   4
      Top             =   0
      Width           =   780
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2925
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   4365
      _ExtentX        =   7705
      _ExtentY        =   5165
      _Version        =   393216
      Cols            =   6
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "V|@|本所案號|案件性質|案件名稱|發文日"
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
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   0
      TabIndex        =   7
      Top             =   4860
      Width           =   4380
      Begin VB.CommandButton Command1 
         Caption         =   "卷宗區"
         Height          =   345
         Index           =   2
         Left            =   1350
         TabIndex        =   19
         Top             =   510
         Width           =   780
      End
      Begin VB.CommandButton Command1 
         Caption         =   "結案單"
         Height          =   345
         Index           =   1
         Left            =   2160
         TabIndex        =   18
         Top             =   135
         Width           =   780
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H0080FFFF&
         Caption         =   "取消"
         Height          =   315
         Index           =   2
         Left            =   675
         Style           =   1  '圖片外觀
         TabIndex        =   13
         Top             =   150
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton Command1 
         Caption         =   "進度(&C)"
         Height          =   345
         Index           =   0
         Left            =   1350
         TabIndex        =   12
         Top             =   135
         Width           =   780
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H0080FFFF&
         Caption         =   "退回"
         Height          =   315
         Index           =   1
         Left            =   45
         Style           =   1  '圖片外觀
         TabIndex        =   10
         Top             =   150
         Width           =   645
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H0000FF00&
         Caption         =   "判發"
         Height          =   315
         Index           =   0
         Left            =   3435
         Style           =   1  '圖片外觀
         TabIndex        =   8
         Top             =   150
         Width           =   870
      End
      Begin VB.Label lblCount 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         Caption         =   "0 / 0"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   3060
         TabIndex        =   9
         Top             =   210
         Width           =   315
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "重整"
      Height          =   345
      Left            =   2760
      TabIndex        =   11
      Top             =   0
      Width           =   780
   End
   Begin VB.Frame Frame2 
      Caption         =   "退回意見"
      Height          =   1395
      Left            =   0
      TabIndex        =   14
      Top             =   3540
      Width           =   4380
      Begin MSForms.TextBox txtAF10 
         Height          =   1155
         Left            =   90
         TabIndex        =   15
         Top             =   180
         Width           =   4200
         VariousPropertyBits=   -1467987941
         ScrollBars      =   2
         Size            =   "7408;2037"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   240
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   930
      TabIndex        =   2
      Top             =   30
      Width           =   1530
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "2699;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "退回重送"
      Height          =   180
      Left            =   1980
      TabIndex        =   17
      Top             =   390
      Width           =   720
   End
   Begin VB.Label Label3 
      Appearance      =   0  '平面
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1755
      TabIndex        =   16
      Top             =   390
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "判發人員："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   45
      TabIndex        =   3
      Top             =   90
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(註:雙擊預覽並選取)"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   390
      Width           =   1605
   End
End
Attribute VB_Name = "frm040119"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/3 改成Form2.0 (MSHFlexGrid1,txtAF10,Combo1
'Created by Morgan 2015/11/2
Option Explicit

Public m_ProState As String '系統別 Added by Morgan 2018/8/16


Dim iPrevRow As Integer '前次點選列
Dim lTotRows As Long, lSelRows As Long
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序

Dim m_AttachPath As String
Dim oFileSys As New FileSystemObject
Dim oFile As File
Dim iColCP09 As Integer, iColCP10 As Integer, iColAF10 As Integer, iColCPP02 As Integer, iColCP140 As Integer


Private Sub Command1_Click(Index As Integer)
   
   Select Case Index
   Case 0 '進度
      PubShowNextData iPrevRow, MSHFlexGrid1, Index
   
   Case 1 '結案單
      ShowCloseSheet
      
   Case 2 '卷宗區
      PubShowNextData iPrevRow, MSHFlexGrid1, Index
      
   End Select
End Sub

Private Sub ShowCloseSheet()
   Dim intFCState As String, strSysKind As String  'Add by Amy 2025/06/11
   Dim strCCM18 As String 'Add by Amy 2025/06/27
   
   If iPrevRow > 0 Then
      strExc(1) = MSHFlexGrid1.TextMatrix(iPrevRow, iColCP140)
      If strExc(1) = "" Then
         'Removed by Morgan 2021/7/26 取消提醒，FMP的結案已改為以電子方式(EMail)--郭
         'MsgBox "非電子結案！", vbInformation
         'end 2021/7/26
         Command1(1).Enabled = False
      Else
         'Add by Amy 2025/06/11 +FC結案單
         intFCState = 0 '非FC結案單
         If strSrvDate(1) >= FCP結案單電子化啟用日 Then
            strExc(9) = MSHFlexGrid1.TextMatrix(iPrevRow, PUB_MGridGetId("本所案號", MSHFlexGrid1))
            strSysKind = SystemNumber(strExc(9), 1)
            'Modify by Amy 2025/06/30 發現舊資料會頁籤判斷會有問題FCP-065275 由P案轉入,1080730閉卷-年費時,由林士堯結案
            '       ex:FCP-065275 由P案轉入,1080730閉卷-年費時,由林士堯結案 / 外商承辦使用國內結案單操作結案 ex:T-242111(結案單號11203939)
            strCCM18 = Pub_GetField("CloseCaseMain", "CCM01='" & strExc(1) & "'", "CCM18")
            If (strSysKind = "P" Or strSysKind = "CFP") And strCCM18 = "F" Then intFCState = 2
            'end 2025/06/30
         End If
         frm210147_1.intFCState = intFCState
         'end 2025/06/11
                     
         Call frm210147_1.SetParent(Me)
         frm210147_1.Hide
         frm210147_1.cmdModify.Visible = False
         frm210147_1.cmdDel.Visible = False
         frm210147_1.cmdFile.Visible = False '檢視回覆單按鈕隱藏
         frm210147_1.txtF0301 = strExc(1)
         frm210147_1.Show
         frm210147_1.QueryData
         Me.Hide
      End If
   Else
      MsgBox "請先選擇一筆資料！", vbInformation
   End If
End Sub

'結案單回本畫面會呼叫,不可刪除
Public Function QueryData() As Boolean
   
End Function

Private Sub Command4_Click()
   If Command4.Caption = "點我展開" Then
      RePosForm True
   Else
      RePosForm False
   End If
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   WebBrowser1.Navigate "about:blank"
   m_AttachPath = App.path & "\" & strUserNum
   KillTemp
   Me.WindowState = 2
   SetCombo1
End Sub

Private Sub Form_Activate()
   Static bDone As Boolean
   If Me.WindowState = 0 Then Me.WindowState = 2
   If bDone = False Then
      Combo1_Click
      bDone = True
   End If
End Sub

Private Sub Combo1_Click()
   If Combo1.Tag <> Combo1 Then
      Command5.Value = True
      Combo1.Tag = Combo1
   End If
End Sub

Private Sub Form_Resize()
   If Me.WindowState = 0 Or Me.WindowState = 2 Then
      If Command4.Caption = "點我展開" Then
         RePosForm False
      Else
         RePosForm True
      End If
   End If
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      
   Dim nCol As Integer, nRow As Integer, iRow As Integer
   Dim stValue As String
   Dim stCP09 As String
   
   With MSHFlexGrid1
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow = 0 Then
      '紀錄前次點選的收文號
      If iPrevRow > 0 Then
         stCP09 = GetValue(iPrevRow, "cp09")
      End If
      
      .col = nCol
      If m_blnColOrderAsc = False Then '字串降冪
         .Sort = 5 '字串昇冪
         m_blnColOrderAsc = True
      Else
         .Sort = 6 '字串降冪
         m_blnColOrderAsc = False
      End If
               
      '重設排序後前次點選的位置
      If iPrevRow > 0 Then
         For iRow = 1 To .Rows - 1
            If stCP09 = GetValue(iRow, "cp09") Then
               iPrevRow = iRow
               Exit For
            End If
         Next
      End If
      
   ElseIf nRow > 0 Then
      .row = nRow
      .col = nCol
      If nCol = 0 Then
         ClickGrid MSHFlexGrid1
      End If
      SelectRow nRow
   End If
   
   .Visible = True
   End With
End Sub

Private Sub MSHFlexGrid1_DblClick()
   If MSHFlexGrid1.MouseRow > 0 Then
      intI = GetFieldId("cpp02", MSHFlexGrid1)
      If MSHFlexGrid1.TextMatrix(iPrevRow, intI) = "" Then
         MsgBox "指示信尚未上傳!!", vbExclamation
      Else
         ReadPdf
         If Command1(1).Enabled Then
            Command1(1).Value = True
         End If
      End If
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim iRow As Integer, bContinue As Boolean
   Dim iCol As Integer
   Dim bolShowForm As Boolean
   Dim strCP09 As String

   SetMouseBusy
   bContinue = False
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         bContinue = True
         
         If Index = 1 Then
            If cmdOK(1).Caption = "確定" Then
               If txtAF10 = "" Or txtAF10 = "請輸入退回意見!!" Then
                  MsgBox "請輸入退回意見!!", vbExclamation
                  txtAF10.SetFocus
                  GoTo EXITSUB
               End If

               'Added by Morgan 2022/1/4 檢查畫面輸入欄位是否含有Unicode文字
               If PUB_ChkUniText(Me, , True, "TextBox") = False Then
                  txtAF10.SetFocus
                  GoTo EXITSUB
               End If
               'end 2022/1/4
   
               iCol = GetFieldId("AF10", Me.MSHFlexGrid1)
               .TextMatrix(iRow, iCol) = txtAF10
               SelectRow iRow
            Else
               SetAF10 iRow, True
               GoTo EXITSUB
            End If
         Else
      
            SelectRow iRow
            
            '取消
            If Index = 2 Then
               SetAF10 iRow
               GoTo EXITSUB
            Else
            
               iCol = GetFieldId("Read", Me.MSHFlexGrid1)
               If .TextMatrix(iRow, iCol) = "" Then
                  .Visible = True
                  MsgBox "請開啟指示信後再行判發或退回!!", vbExclamation
                  GoTo EXITSUB
               End If
               
            End If
         End If
         Exit For
      End If
   Next
   End With
   
   If bContinue = False Then
      MsgBox "請先勾選(V)資料列！", vbInformation
   Else
      FormSave Index, MSHFlexGrid1
   End If
   
EXITSUB:

   SetMouseReady
End Sub

Private Sub Command5_Click()
   WebBrowser1.Navigate "about:blank": DoEvents 'Added by Morgan 2019/1/19
   SetMouseBusy
   ReadGridData
   SetMouseReady
   KillTemp 'Added by Morgan 2019/1/19 要先刪除暫存否則退會重送的定稿可能會讀到舊的
End Sub

Private Sub SetGrid(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGridHeadWidth
   Dim iUbound As Integer
   'Modified by Morgan 2018/7/24 本所案號欄加寬(顯示EPC子案流水號)
   arrGridHeadWidth = Array(240, 240, 1400, 800, 800, 825)
   iUbound = UBound(arrGridHeadWidth)
   
   With MSHFlexGrid1
   If pReset = True Then
      .Clear
      .Rows = 2
      
      iPrevRow = 0
      lTotRows = 0
      lSelRows = 0
      lblCount = lSelRows & " / " & lTotRows
   End If
   .FixedCols = 3
   .FormatString = "V| |本所案號|案件性質|案件名稱|發文日"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         .ColAlignment(iCol) = flexAlignLeftCenter
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   .row = 0
   .col = 1
   Set .CellPicture = PictureClip1.Picture
   End With
End Sub

Private Sub ReadGridData()
   Dim iRow As Integer, iCol As Integer
   Dim stCon As String
   Dim idx As Integer
   
   If Trim(Left(Combo1.Text, 6)) <> "" Then
      stCon = " and AF06='" & Trim(Left("" & Combo1.Text, 6)) & "'"
   End If
   
   SetGrid True
   'Modified by Morgan 2018/8/27 +排除未上傳的自行判發案件
   'Modified by Morgan 2020/5/5 指示信檔名改另外抓(要排除已刪除的檔案)
   strExc(0) = "select '' V,'' PDF,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",cpm04 案件性質,pa05 案件名稱,sqldatet(cp27) 發文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 CaseNo" & _
      ",'' Read,AF01,AF09,AF10,CP09,CP140,'' cpp02,cp10" & _
      " From AppForm, caseprogress, patent, casepropertymap" & _
      " where AF07=0 " & stCon & " and cp09(+)=AF01 and (af06<>af15" & _
      " or exists(select cpp02 from casepaperpdf where cpp01=af01" & _
      " and substr(upper(cpp02),-9)='.DATA.PDF' and cpp10<>'D'))" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 order by AF09,CP27"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With MSHFlexGrid1
      .Visible = False
      .FixedCols = 0
      Set .Recordset = RsTemp.Clone
      SetGrid
      
      lTotRows = RsTemp.RecordCount
      lblCount = lSelRows & " / " & lTotRows
      idx = GetFieldId("案件性質", MSHFlexGrid1)
      iColCP09 = GetFieldId("CP09", MSHFlexGrid1)
      iColCP10 = GetFieldId("CP10", MSHFlexGrid1)
      iColCPP02 = GetFieldId("CPP02", MSHFlexGrid1)
      iColAF10 = GetFieldId("AF10", MSHFlexGrid1)
      iColCP140 = GetFieldId("CP140", MSHFlexGrid1)
      For iRow = 1 To .Rows - 1
         .TextMatrix(iRow, idx) = .TextMatrix(iRow, idx) & PUB_GetRelateCasePropertyName(.TextMatrix(iRow, iColCP09), "1")
         
         'Added by Morgan 2020/5/5
         '指示信檔名改另外抓(要排除已刪除的檔案)
         .TextMatrix(iRow, iColCPP02) = GetDataPdf(.TextMatrix(iRow, iColCP09))
         'end 2020/5/5
         
         '指示信
         If .TextMatrix(iRow, iColCPP02) <> "" Then
            .row = iRow
            .col = 1
            Set .CellPicture = PictureClip1.Picture
         End If
         '退回
         If .TextMatrix(iRow, iColAF10) <> "" Then
            For iCol = 0 To .Cols - 1
               If iCol <> .FixedCols - 1 Then
                  .row = iRow
                  .col = iCol
                  .CellBackColor = cmdOK(1).BackColor
               End If
            Next
         End If
      Next
      .col = 1: .row = 1
      SelectRow 1
      .Visible = True
      End With
   Else
      SetAF10
      MsgBox "無待判發資料！", vbExclamation
   End If
End Sub

Private Function GetFieldId(pFieldName As String, ByRef FlexGrid As MSHFlexGrid) As Integer
   Dim iRow As Integer
   With FlexGrid
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         GetFieldId = iRow
         Exit For
      End If
   Next
   End With
End Function

Private Function GetValue(pRow As Integer, pFieldName As String) As String
   Dim iRow As Integer
   With MSHFlexGrid1
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         GetValue = .TextMatrix(pRow, iRow)
         Exit For
      End If
   Next
   End With
End Function

Private Function SetValue(pRow As Integer, pFieldName As String, pValue As String) As Boolean
   Dim iRow As Integer
   With MSHFlexGrid1
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         .TextMatrix(pRow, iRow) = pValue
         SetValue = True
         Exit Function
      End If
   Next
   End With
End Function

Private Function FormSave(pIdx As Integer, ByRef UpdFlexGrid As MSHFlexGrid) As Boolean
   Dim iRow As Integer, idxCaseNo As Integer
   Dim strCP09 As String, strCPP02 As String
   Dim strSub As String, strContent As String
   Dim bDone As Boolean
   
On Error GoTo ErrHnd
   
   With UpdFlexGrid
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         strCP09 = .TextMatrix(iRow, iColCP09)
         strCPP02 = .TextMatrix(iRow, iColCPP02)
         'Removed by Morgan 2015/11/24 改由程序傳送EMail
         ''判發要彈EMail給代理人的視窗
         'If pIdx = 0 Then
         '   ShowMailForm strCP09, strCPP02, bDone
         '   If Not bDone Then
         '      Exit Function
         '   End If
         'End If
         'end 2015/11/24
      
         cnnConnection.BeginTrans
         
On Error GoTo ErrHndT
         
         '判發
         If pIdx = 0 Then
            strSql = "update appform set af06='" & strUserNum & "',af07=" & strSrvDate(1) & " where af01='" & strCP09 & "'"
            cnnConnection.Execute strSql, intI
            
            'Removed by Morgan 2015/11/24 改由程序傳送EMail
            ''電子表單
            'iColCP140 = GetFieldId("CP140", UpdFlexGrid)
            'If .TextMatrix(iRow, iColCP140) <> "" Then
            '   strSql = "update flow003 set f0309='03' where f0301='" & .TextMatrix(iRow, iColCP140) & "' and f0309='09'"
            '   cnnConnection.Execute strSql, intI
            'End If
            'end 2015/11/24
            
         '退回
         Else
            strSql = "update appform set af09=sysdate,af10='" & ChgSQL(.TextMatrix(iRow, iColAF10)) & "' where af01='" & strCP09 & "'"
            cnnConnection.Execute strSql, intI
            
'Removed by Morgan 2015/11/25
'            '發EMail給程序
'            strExc(0) = "select cp83 from appform,caseprogress where af01='" & .TextMatrix(iRow, iColCP09) & "' and cp09(+)=af01"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               idxCaseNo = GetFieldId("本所案號", MSHFlexGrid1)
'               strSub = "指示信判發退回:" & .TextMatrix(iRow, idxCaseNo) & "(" & strCP09 & ")"
'               strContent = "本所案號：" & .TextMatrix(iRow, idxCaseNo)
'               strContent = strContent & vbCrLf & "案件名稱：" & .TextMatrix(iRow, GetFieldId("案件名稱", MSHFlexGrid1))
'               strContent = strContent & vbCrLf & "案件性質：" & .TextMatrix(iRow, GetFieldId("案件性質", MSHFlexGrid1))
'               strContent = strContent & vbCrLf & "發文日：" & .TextMatrix(iRow, GetFieldId("發文日", MSHFlexGrid1))
'               strContent = strContent & vbCrLf & "退回意見：" & .TextMatrix(iRow, iColAF10)
'
'               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'                  " values( '" & strUserNum & "','" & RsTemp("cp83") & "',to_char(sysdate,'yyyymmdd')" & _
'                  ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strSub) & "','" & ChgSQL(strContent) & "')"
'               cnnConnection.Execute strSql, intI
'            End If
            
            '刪除指示信
            PUB_DelFtpFile2 strCP09, " and instr(upper(cpp02),'.DATA.PDF')>0"
            strSql = "delete casepaperpdf where cpp01='" & strCP09 & "' and instr(upper(cpp02),'.DATA.PDF')>0"
            cnnConnection.Execute strSql, intI
         End If
         
         cnnConnection.CommitTrans
         
On Error GoTo ErrHnd
         If iRow = iPrevRow Then SelectRow 0
         .TextMatrix(iRow, 0) = "X"
         .RowHeight(iRow) = 0
         lSelRows = lSelRows - 1
         lTotRows = lTotRows - 1
         lblCount = lSelRows & " / " & lTotRows
         DoEvents
      End If
   Next
   End With
   
   WebBrowser1.Navigate "about:blank"
   FormSave = True
   PUB_SendMailCache
   Exit Function
   
ErrHndT:
   cnnConnection.RollbackTrans
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Private Sub SelectRow(pRow As Integer)
   Dim nCol As Integer, iCol As Integer, lColor As Long
   With MSHFlexGrid1
   nCol = .col
   If iPrevRow > 0 Then
      If iPrevRow <> pRow Then
         .row = iPrevRow
         
         iCol = GetFieldId("AF09", MSHFlexGrid1)
         If .TextMatrix(.row, iCol) <> "" Then
            lColor = cmdOK(1).BackColor
         Else
            lColor = .BackColor
         End If
         
         For iCol = 0 To .Cols - 1
            If iCol >= .FixedCols Then
               .col = iCol
               .CellBackColor = lColor
               
            ElseIf iCol = .FixedCols - 1 Then
               .col = iCol
               .CellBackColor = .BackColorFixed
               .CellForeColor = .ForeColor
            End If
         Next
         
         
      End If
   End If
   If pRow > 0 Then
      .row = pRow
      If .FixedCols > 0 Then
         .col = .FixedCols - 1
         .CellBackColor = .BackColorSel
         .CellForeColor = .ForeColorSel
      End If
         
      For iCol = .FixedCols To .Cols - 1
        .col = iCol
        .CellBackColor = &HFFC0C0
      Next
   End If
   .col = nCol
   iPrevRow = pRow
   SetAF10 pRow
   
   'Added by Morgan 2016/5/24  結案單控制
   Command1(1).Enabled = False
   If iPrevRow > 0 Then
      strExc(1) = MSHFlexGrid1.TextMatrix(iPrevRow, iColCP10)
      If strExc(1) = "907" Or strExc(1) = "913" Or strExc(1) = "925" Then
         Command1(1).Enabled = True
      End If
   End If
   'end 2016/5/24
   
   End With
End Sub

'帶出退回意見
Private Sub SetAF10(Optional pRow As Integer, Optional bolReturn As Boolean = False)
   Dim iCol As Integer
   
   With MSHFlexGrid1
   If pRow = 0 Then
      txtAF10 = ""
   Else
      iCol = GetFieldId("AF10", MSHFlexGrid1)
      txtAF10 = .TextMatrix(pRow, iCol)
   End If
   If txtAF10 <> "" Or bolReturn Then
      .Height = Frame1.Top - .Top - Frame2.Height
   Else
      .Height = Frame1.Top - .Top + 50
   End If
   
   If bolReturn Then
      If txtAF10 = "" Then
         txtAF10 = "請輸入退回意見!!"
      End If
      txtAF10.SelStart = 0
      txtAF10.SelLength = Len(txtAF10)
      txtAF10.SetFocus
      txtAF10.Locked = False
      .Enabled = False
      cmdOK(1).Caption = "確定"
      cmdOK(2).Visible = True
      cmdOK(0).Visible = False
   Else
      .Enabled = True
      txtAF10.Locked = True
      cmdOK(1).Caption = "退回"
      cmdOK(2).Visible = False
      cmdOK(0).Visible = True
   End If
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oFileSys = Nothing
   Set oFile = Nothing
   Set frm040119 = Nothing
End Sub

Private Sub ClickGrid(FlexGrid As MSHFlexGrid, Optional pSelected As Boolean = False)
   Dim iCol As Integer, iRow As Integer
   
   With FlexGrid
   iCol = GetFieldId("V", MSHFlexGrid1)
   If pSelected Then
      If .TextMatrix(.row, iCol) = "" Then
         lSelRows = lSelRows + 1
         .TextMatrix(.row, iCol) = "V"
      End If
      For iRow = 1 To .Rows - 1
         If iRow <> .row Then
            .TextMatrix(iRow, iCol) = ""
         End If
      Next
   ElseIf .TextMatrix(.row, iCol) = "V" Then
      lSelRows = lSelRows - 1
      .TextMatrix(.row, iCol) = ""
      
'   '已刪除資料標示為 X
'   ElseIf .Text = "" Then
'      iCol = GetFieldId("cpp02", MSHFlexGrid1)
'      If MSHFlexGrid1.TextMatrix(iPrevRow, iCol) = "" Then
'         .Visible = True
'         MsgBox "指示信尚未上傳不可勾選!!", vbExclamation
'         Exit Sub
'      End If
'
'      iCol = GetFieldId("Read", FlexGrid)
'      If .TextMatrix(.row, iCol) = "" Then
'         .Visible = True
'         MsgBox "請開啟指示信後再行勾選!!", vbExclamation
'         Exit Sub
'      End If
'
'      lSelRows = lSelRows + 1
'      .Text = "V"
   End If
   lblCount = lSelRows & " / " & lTotRows
   End With
End Sub

Private Sub SetMouseBusy()
   Screen.MousePointer = vbHourglass
   MSHFlexGrid1.MousePointer = vbHourglass
End Sub

Private Sub SetMouseReady()
   Screen.MousePointer = vbDefault
   MSHFlexGrid1.MousePointer = vbDefault
End Sub

Private Sub SetCombo1()
Combo1.Clear
'Modified by Morgan 2018/8/16 +CFP
If m_ProState = "CFP" Then
   Me.Caption = "CFP案指示信判發作業" 'Added by Morgan 2024/3/21
   strExc(0) = "select st01||' '||st02 Rvr from AppForm,caseprogress,staff" & _
      " where AF07=0 and af06 is not null and cp09(+)=af01 and cp01='CFP' and st01(+)=af06" & _
      " union select decode(s1.st01,null,s2.st01||' '||s2.st02,s1.st01||' '||s1.st02) Rvr" & _
      " From LETTERREVIEWER,staff s1,SetSpecMan,staff s2" & _
      " where lr02='2' and lr03='CFP' and lr07 is not null" & _
      " and s1.st01(+)=lr07 and OCODE(+)=lr07 and s2.st01(+)=oman order by 1 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         '自己排第一個
         If InStr(RsTemp(0), strUserNum) = 1 Then
            Combo1.AddItem RsTemp(0), 0
         Else
            Combo1.AddItem RsTemp(0)
         End If
         .MoveNext
      Loop
      End With
      Combo1.ListIndex = 0
   End If
Else
   Me.Caption = "P案指示信判發作業" 'Added by Morgan 2024/3/21
   strExc(1) = Pub_GetSpecMan("PS4")
   If strExc(1) <> "" Then
      strExc(2) = GetPrjSalesNM(strExc(1))
      Combo1.AddItem strExc(1) & " " & strExc(2)
      Combo1.ListIndex = 0
   End If
   Combo1.Enabled = False
End If

End Sub

'Modified by Morgan 2019/1/19 加重試3次後彈訊息(檔案被鎖住時無法刪除)
Private Sub KillTemp()
   Dim iTimes As Integer
On Error GoTo ErrHnd
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
   Exit Sub
   
ErrHnd:
   If iTimes < 2 Then
      iTimes = iTimes + 1
      Sleep 1000
      Resume
   Else
      'MsgBox "暫存檔無法清除！" & vbCrLf & vbCrLf & "請重新執行本作業，否則有可能載入的不是最新的定稿！", vbExclamation
   End If
   Err.Clear
End Sub

Private Sub ReadPdf()
   Dim stFileName As String
   Dim idx As Integer

   If iPrevRow = 0 Then
      MsgBox "請先點選欲預覽的資料列！", vbInformation
   Else
      SetMouseBusy
      With MSHFlexGrid1
      WebBrowser1.Navigate "about:blank": DoEvents
      
      idx = GetFieldId("cpp02", MSHFlexGrid1)
      stFileName = .TextMatrix(iPrevRow, idx)
      
      idx = GetFieldId("cp09", MSHFlexGrid1)
      If PUB_GetAttachFile_CPP(.TextMatrix(iPrevRow, idx), stFileName, m_AttachPath) = True Then
         'Modified by Morgan 2017/6/1
         'WebBrowser1.Navigate m_AttachPath & "\" & stFileName
         WebBrowser1.Navigate stFileName
         'end 2017/6/1
         SetValue iPrevRow, "Read", "Y"
         ClickGrid MSHFlexGrid1, True
      End If
      End With
      SetMouseReady
   End If
End Sub

Private Sub RePosForm(pFull As Boolean)
   Static lngLeft As Long
   Dim a1 As Integer
   If Forms(0).WindowState <> 1 Then
      If lngLeft = 0 Then lngLeft = Command4.Left
      If pFull = True Then
         WebBrowser1.Left = 0
         WebBrowser1.Width = Me.Width - 90
         WebBrowser1.Height = Me.Height - Command4.Height - 390
         Command4.Caption = "點我還原"
      Else
         WebBrowser1.Left = lngLeft
         Command4.Caption = "點我展開"
      End If
      WebBrowser1.Width = Me.Width - 90 - WebBrowser1.Left
      WebBrowser1.Height = Me.Height - Command4.Height - 390
      Command4.Left = WebBrowser1.Left
      Command4.Width = WebBrowser1.Width
      
      If MSHFlexGrid1.Enabled = True And txtAF10 = "" Then
         MSHFlexGrid1.Height = Me.Height - MSHFlexGrid1.Top - Frame1.Height - 350
         Frame1.Top = MSHFlexGrid1.Top + MSHFlexGrid1.Height - 50
         Frame2.Top = Frame1.Top - Frame2.Height + 50
      Else
         a1 = Me.Height - MSHFlexGrid1.Top - Frame1.Height - Frame2.Height - 450
         If a1 > 0 Then
            MSHFlexGrid1.Height = a1
            Frame2.Top = MSHFlexGrid1.Top + MSHFlexGrid1.Height + 50
            Frame1.Top = Frame2.Top + Frame2.Height
         End If
      End If
   End If
End Sub

Public Sub PubShowNextData(ByRef iPt As Integer, ByRef Fgrid As MSHFlexGrid, ByRef Index As Integer)
   If iPt = 0 Then Exit Sub
   If fnSaveParentForm(Me) = False Then
      Me.Enabled = True
      Exit Sub
   End If
   
   Me.Enabled = False
   Screen.MousePointer = vbHourglass
   
   Select Case Index
   Case 0
      frm100101_2.Show
      intI = GetFieldId("CaseNo", Fgrid)
      frm100101_2.Tag = Pub_RplStr(Fgrid.TextMatrix(iPt, intI))
      frm100101_2.cmdOK(5).Visible = False '下一筆按鈕隱藏
      frm100101_2.StrMenu
      
   Case 2
      frm100101_L.m_strKey = Fgrid.TextMatrix(iPt, iColCP09)
      frm100101_L.SetParent Me
      If frm100101_L.QueryData = True Then
         frm100101_L.Show
         Me.Hide
      End If
    End Select
    
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    
End Sub

Private Sub txtAF10_GotFocus()
   OpenIme
End Sub

Private Sub txtAF10_LostFocus()
   CloseIme
End Sub

Private Function GetDataPdf(pRecNo As String) As String
   Dim stSQL As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
   
   stSQL = "select cpp02 from casepaperpdf where cpp01='" & pRecNo & "' and substr(upper(cpp02),-9)='.DATA.PDF' and cpp10<>'D'"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      GetDataPdf = RsQ.Fields("cpp02")
   End If
   Set RsQ = Nothing
End Function
