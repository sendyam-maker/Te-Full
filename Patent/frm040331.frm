VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm040331 
   BorderStyle     =   1  '單線固定
   Caption         =   "智慧局年費通知核對清單"
   ClientHeight    =   3732
   ClientLeft      =   2796
   ClientTop       =   3948
   ClientWidth     =   5436
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3732
   ScaleWidth      =   5436
   Begin VB.TextBox txtBaseDate 
      Height          =   270
      Left            =   1980
      MaxLength       =   7
      TabIndex        =   20
      Top             =   390
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   765
      Left            =   180
      TabIndex        =   18
      Top             =   3750
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   1355
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame2 
      Caption         =   "逾期"
      Height          =   1215
      Left            =   180
      TabIndex        =   14
      Top             =   2430
      Width           =   5100
      Begin VB.TextBox txtPath 
         Height          =   315
         Index           =   1
         Left            =   1035
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   270
         Width           =   3975
      End
      Begin VB.TextBox txtDateLimit 
         Height          =   345
         Index           =   2
         Left            =   1035
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   690
         Width           =   1095
      End
      Begin VB.TextBox txtDateLimit 
         Height          =   345
         Index           =   3
         Left            =   2475
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "檔案路徑："
         Height          =   180
         Left            =   225
         TabIndex        =   17
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "法定期限："
         Height          =   180
         Left            =   225
         TabIndex        =   16
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "－"
         Height          =   180
         Left            =   2250
         TabIndex        =   15
         Top             =   780
         Width           =   180
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "正常"
      Height          =   1215
      Left            =   180
      TabIndex        =   10
      Top             =   1050
      Width           =   5100
      Begin VB.TextBox txtDateLimit 
         Height          =   345
         Index           =   1
         Left            =   2475
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   690
         Width           =   1095
      End
      Begin VB.TextBox txtDateLimit 
         Height          =   345
         Index           =   0
         Left            =   1035
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   690
         Width           =   1095
      End
      Begin VB.TextBox txtPath 
         Height          =   315
         Index           =   0
         Left            =   1035
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   270
         Width           =   3975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "－"
         Height          =   180
         Left            =   2250
         TabIndex        =   13
         Top             =   780
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "法定期限："
         Height          =   180
         Left            =   225
         TabIndex        =   12
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "檔案路徑："
         Height          =   180
         Left            =   225
         TabIndex        =   11
         Top             =   330
         Width           =   900
      End
   End
   Begin VB.TextBox txtMailDate 
      Height          =   270
      Left            =   1980
      MaxLength       =   7
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4530
      TabIndex        =   8
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3735
      TabIndex        =   7
      Top             =   60
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   765
      Left            =   180
      TabIndex        =   19
      Top             =   4590
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   1355
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "基準日期請輸入 5 或 20 號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   360
      TabIndex        =   22
      Top             =   60
      Width           =   2580
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "期限基準日期："
      Height          =   180
      Left            =   360
      TabIndex        =   21
      Top             =   420
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "電子郵件檔案日期："
      Height          =   180
      Left            =   360
      TabIndex        =   9
      Top             =   750
      Width           =   1620
   End
End
Attribute VB_Name = "frm040331"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
'Add by Morgan 2005/6/28 列印用
Dim PLeft() As Integer, iPrint As Integer, iPage As Integer, strTemp() As String
Dim m_iTitleFontSize As Single, m_iFontSize As Single
Dim m_iStartX As Integer, m_iStartY As Integer
Dim m_iPageHeight As Integer, m_iLineHeight As Integer, m_stTmp As String
Dim m_iMargin As Integer, m_Title As String
Dim m_iType As Integer '1=正常 2=逾期
Dim m_AttachPath As String 'Added by Morgan 2015/7/13
Dim m_TitleS As String  'Added by Morgan 2025/1/20

Private Sub ResetGrid(ByRef p_Grid As MSHFlexGrid)
   With p_Grid
      .Clear
      .Rows = 2
      .FixedRows = 1
      .FixedCols = 0
      'Modified by Morgan 2025/1/16 +PID
      .FormatString = "本所案號|申請案號|證書號|專利權人|專利名稱|繳納年次|繳納期限|加倍補繳期限|刪除|排序|續辦|下次年度|案號2|案號3|案號4|PID"
   End With
End Sub

Private Function LoadXLS(ByRef p_Recs1 As Integer, p_Recs2 As Integer) As Boolean

   Dim xlsAnnuity As New Excel.Application
   Dim wkbAnnuity As New Excel.Workbook
   Dim wksAnnuity As New Excel.Worksheet
   Dim iRow As Integer
   
On Error GoTo ErrHnd
   p_Recs1 = 0: p_Recs2 = 0
   If txtPath(0) <> "" Then
      If Dir(txtPath(0)) = "" Then
         If MsgBox("正常年費檔案不存在，是否繼續！", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
            txtPath(0).SetFocus
            Exit Function
         End If
      Else
         '正常
         Set wkbAnnuity = xlsAnnuity.Workbooks.Open(txtPath(0), 0, True)
         Set wksAnnuity = wkbAnnuity.Worksheets(1)
         
         With wksAnnuity
            iRow = 3
            ResetGrid MSHFlexGrid1
            Do While Not .Range("A" & iRow).Value = ""
               p_Recs1 = p_Recs1 + 1
               MSHFlexGrid1.Rows = p_Recs1 + 1
               'Modify by Morgan 2010/12/28 申請案號改碼數(不必轉)
               'MSHFlexGrid1.TextMatrix(p_Recs1, 1) = Format(.Range("A" & iRow).Value, "#")
               MSHFlexGrid1.TextMatrix(p_Recs1, 1) = "" & .Range("A" & iRow).Value
               MSHFlexGrid1.TextMatrix(p_Recs1, 2) = .Range("B" & iRow).Value
               MSHFlexGrid1.TextMatrix(p_Recs1, 3) = .Range("C" & iRow).Value
               MSHFlexGrid1.TextMatrix(p_Recs1, 4) = .Range("D" & iRow).Value
               MSHFlexGrid1.TextMatrix(p_Recs1, 5) = Format(.Range("E" & iRow).Value, "#")
               MSHFlexGrid1.TextMatrix(p_Recs1, 6) = Format(.Range("F" & iRow).Value, "#")
               MSHFlexGrid1.TextMatrix(p_Recs1, 7) = Format(.Range("G" & iRow).Value, "#")
               iRow = iRow + 1
            Loop
         End With
         xlsAnnuity.Workbooks.Close
      End If
   End If
   
   If txtPath(1) <> "" Then
      If Dir(txtPath(1)) = "" Then
         If MsgBox("逾期年費檔案不存在，是否繼續！", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
            txtPath(1).SetFocus
            Exit Function
         End If
      Else
         '逾期
         Set wkbAnnuity = xlsAnnuity.Workbooks.Open(txtPath(1), 0, True)
         Set wksAnnuity = wkbAnnuity.Worksheets(1)
         
         With wksAnnuity
            iRow = 3
            ResetGrid MSHFlexGrid2
            Do While Not .Range("A" & iRow).Value = ""
               p_Recs2 = p_Recs2 + 1
               MSHFlexGrid2.Rows = p_Recs2 + 1
               'Modify by Morgan 2010/12/28 申請案號改碼數(不必轉)
               'MSHFlexGrid2.TextMatrix(p_Recs2, 1) = Format(.Range("A" & iRow).Value, "#")
               MSHFlexGrid2.TextMatrix(p_Recs2, 1) = "" & .Range("A" & iRow).Value
               MSHFlexGrid2.TextMatrix(p_Recs2, 2) = .Range("B" & iRow).Value
               MSHFlexGrid2.TextMatrix(p_Recs2, 3) = .Range("C" & iRow).Value
               MSHFlexGrid2.TextMatrix(p_Recs2, 4) = .Range("D" & iRow).Value
               MSHFlexGrid2.TextMatrix(p_Recs2, 5) = Format(.Range("E" & iRow).Value, "#")
               MSHFlexGrid2.TextMatrix(p_Recs2, 6) = Format(.Range("F" & iRow).Value, "#")
               MSHFlexGrid2.TextMatrix(p_Recs2, 7) = Format(.Range("G" & iRow).Value, "#")
               iRow = iRow + 1
            Loop
         End With
         xlsAnnuity.Workbooks.Close
      End If
   End If
   xlsAnnuity.Quit
   
   LoadXLS = True
ErrHnd:
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   
   Set xlsAnnuity = Nothing
   Set wkbAnnuity = Nothing
   Set wksAnnuity = Nothing

End Function

Private Function Process1() As Boolean

   Dim iRow As Integer, iRecs As Integer, iXRow As Integer
   
 On Error GoTo ErrHnd
 
   m_iType = 1
 
   For iRow = 1 To MSHFlexGrid1.Rows - 1
      CompareData MSHFlexGrid1, iRow
   Next
 
   '新增未通知案件
   '以本所號排序,遞減
   MSHFlexGrid1.col = 0
   MSHFlexGrid1.Sort = flexSortGenericDescending
   For iRow = 1 To MSHFlexGrid1.Rows - 1
      If MSHFlexGrid1.TextMatrix(iRow, 0) <> "" Then
         iRecs = iRow
      End If
   Next
   
   'Modified by Morgan 2025/1/20 +,NP03,NP04,NP05
   strSql = "SELECT NP02||'-'||NP03||'-'||NP04||'-'||NP05 CNo,NP02,NP03,NP04,NP05 FROM NEXTPROGRESS,PATENT" & _
      " WHERE NP02='P' and NP06 is null and NP07='605' and NP09>=" & TransDate(txtDateLimit(0), 2) & " AND NP09<=" & TransDate(txtDateLimit(1), 2) & _
      " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA09='000' AND SUBSTR(PA14,5)=TO_CHAR(TO_DATE(NP09,'YYYYMMDD')+1,'MMDD')"
   
   strSql = strSql & " UNION ALL" & _
      " SELECT NP02||'-'||NP03||'-'||NP04||'-'||NP05 CNo,NP02,NP03,NP04,NP05 FROM NEXTPROGRESS,PATENT" & _
      " WHERE NP02='FCP' and NP06 is null and NP07='605' and NP09>=" & TransDate(txtDateLimit(0), 2) & " AND NP09<=" & TransDate(txtDateLimit(1), 2) & _
      " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA09='000' AND SUBSTR(PA14,5)=TO_CHAR(TO_DATE(NP09,'YYYYMMDD')+1,'MMDD')"
      
   strSql = strSql & " order by 1 desc"
      
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         '資料尋找起始列
         iXRow = 1
         Do While Not .EOF
            If FindData("" & .Fields("CNo"), iRecs, iXRow) = False Then
               MSHFlexGrid1.AddItem "" & .Fields("CNo")
               MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 9) = "" & .Fields("NP02")
               'Added by Morgan 2025/1/20
               MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 12) = "" & .Fields("NP03")
               MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 13) = "" & .Fields("NP04")
               MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 14) = "" & .Fields("NP05")
               'end 2025/1/20
            End If
            .MoveNext
         Loop
      End If
   End With
   
   '刪除不要的資料
   '以刪除旗標排序,遞增
   MSHFlexGrid1.col = 8
   MSHFlexGrid1.Sort = flexSortGenericAscending
   iRecs = 0
   For iRow = 1 To MSHFlexGrid1.Rows - 1
   
      'Added by Morgan 2025/1/20
      With MSHFlexGrid1
      If .TextMatrix(iRow, 9) = "P" And strSrvDate(1) >= P業務區劃分啟用日 Then
         If .TextMatrix(iRow, 12) <> "" Then
            .TextMatrix(iRow, 15) = PUB_GetPHandler(.TextMatrix(iRow, 9) & "-" & .TextMatrix(iRow, 12) & "-" & .TextMatrix(iRow, 13) & "-" & .TextMatrix(iRow, 14))
         End If
      End If
      End With
      'end 2025/1/20
      
      If MSHFlexGrid1.TextMatrix(iRow, 8) = "1" Then
         Exit For
      Else
         iRecs = iRow
      End If
   Next
   MSHFlexGrid1.Rows = iRecs + 1 '刪除不要的資料
   
   '以本所號排序,遞增
   MSHFlexGrid1.col = 0
   MSHFlexGrid1.Sort = flexSortGenericAscending
   
   'Added by Morgan 2025/1/20 再以管制人排序
   MSHFlexGrid1.col = 15
   MSHFlexGrid1.Sort = flexSortGenericAscending
   'end 2025/1/20
   
   InsertQueryLog (MSHFlexGrid1.Rows - 1) 'Add By Sindy 2010/12/2
   
   m_Title = "內專" & Me.Caption
   m_TitleS = m_Title 'Added by Morgan 2025/1/20
   If Not DoPrint(MSHFlexGrid1, "P") Then Exit Function
   
   'Removed by Morgan 2015/7/13 外專不核對--秀玲
   'm_Title = "FCP" & Me.Caption
   'If Not DoPrint(MSHFlexGrid1, "FCP") Then Exit Function
   'end 2015/7/13
  
   Process1 = True
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   'Resume

End Function

Private Function Process2() As Boolean

   Dim iRow As Integer, iRecs As Integer, iPPos As Integer
   Dim stPdfName As String
   
 On Error GoTo ErrHnd
 
   m_iType = 2

   With MSHFlexGrid2
      For iRow = 1 To .Rows - 1
         CompareData MSHFlexGrid2, iRow, 2
   
         'Added by Morgan 2025/1/21
         If .TextMatrix(iRow, 9) = "P" And strSrvDate(1) >= P業務區劃分啟用日 Then
            If .TextMatrix(iRow, 12) <> "" Then
               .TextMatrix(iRow, 15) = PUB_GetPHandler(.TextMatrix(iRow, 9) & "-" & .TextMatrix(iRow, 12) & "-" & .TextMatrix(iRow, 13) & "-" & .TextMatrix(iRow, 14))
            End If
         End If
         'end 2025/1/21
      Next
      '以本所號排序,遞增
      .col = 0
      .Sort = flexSortGenericAscending
            
      '以客戶名排序,遞增
      .col = 3
      .Sort = flexSortGenericAscending
      
      'Added by Morgan 2025/1/20 再以管制人排序
      .col = 15
      .Sort = flexSortGenericAscending
      'end 2025/1/20
   
      If AddCP("P") = False Then Exit Function 'Add by Morgan 2010/3/8
      
      m_Title = "內專" & Me.Caption & "(逾期)"
      m_TitleS = m_Title 'Added by Morgan 2025/1/21
      If DoPrint(MSHFlexGrid2, "P") = False Then Exit Function
      
      '以本所號排序,遞增
      .col = 0
      .Sort = flexSortGenericAscending
      InsertQueryLog (MSHFlexGrid2.Rows - 1) 'Add By Sindy 2010/12/2
      
      If AddCP("FCP") = False Then Exit Function 'Add by Morgan 2010/3/8
      m_Title = "FCP" & Me.Caption & "(逾期)"
      
      'Added by Morgan 2015/7/13
      stPdfName = m_Title & "_" & strSrvDate(1) & ".pdf"
      If Dir(m_AttachPath & "\" & stPdfName) <> "" Then
         Kill m_AttachPath & "\" & stPdfName
      End If
      frmPDF.Show
      frmPDF.StartProcess m_AttachPath, stPdfName
      'end 2015/7/13
      
      If DoPrint(MSHFlexGrid2, "FCP") = False Then Exit Function
      
      'Added by Morgan 2015/7/13
      frmPDF.EndtProcess
      Unload frmPDF
      'Modified by Lydia 2020/09/10 改成系統特殊設定
      'PUB_SendMail strUserNum, "85033", "", m_Title, "電子郵件檔案日期：" & txtMailDate & vbCrLf & "法定期限：" & txtDateLimit(2) & " - " & txtDateLimit(3), "", m_AttachPath & "\" & stPdfName
      ''end 2015/7/13
      strExc(0) = Pub_GetSpecMan("外專程序-通知年費逾期")
      If strExc(0) = "" Then
          MsgBox "外專程序-通知年費逾期人員編號不存在，請通知電腦中心！", vbInformation, "外專程序-通知年費逾期"
      Else
          PUB_SendMail strUserNum, strExc(0), "", m_Title, "電子郵件檔案日期：" & txtMailDate & vbCrLf & "法定期限：" & txtDateLimit(2) & " - " & txtDateLimit(3), "", m_AttachPath & "\" & stPdfName
      End If
      'end 2020/09/10
      
   End With
   
   Process2 = True
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description, vbCritical
   End If

End Function

'Add by Morgan 2010/3/8
'新增逾期補繳通知來函
Private Function AddCP(p_Sys As String) As Boolean
   Dim stCP01 As String, stCP02 As String, stCP03 As String, stCP04 As String, stCP64 As String
   Dim stCP05 As String, stCP06 As String, stCP07 As String, stCP09 As String
   Dim stCP13 As String, stCP12 As String, stCP20 As String, stCP16 As String
   Dim iRow As Integer
   Dim strDate(3) As String
   Dim strCP14 As String, strCP48 As String  'Added by Lydia 2019/05/31 預設承辦人和承辦期限
   Dim m_PA177 As String 'Added by Lydia 2023/07/28
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd

   stCP05 = DBDATE(txtMailDate)
   
   With MSHFlexGrid2
   For iRow = 1 To .Rows - 1
      'Add by Morgan 2010/3/8 新增逾期補繳通知來函
      If .TextMatrix(iRow, 9) = p_Sys Then
         stCP01 = .TextMatrix(iRow, 9)
         If stCP01 = "FCP" Or Val(.TextMatrix(iRow, 5)) >= Val(.TextMatrix(iRow, 11)) Then
            stCP02 = .TextMatrix(iRow, 12)
            stCP03 = .TextMatrix(iRow, 13)
            stCP04 = .TextMatrix(iRow, 14)
            m_PA177 = "" 'Added by Lydia 2023/07/28
            
            'Added by Morgan 2024/8/19 重下面移上來,因若已有通知紀錄時不可再更新繳年費紀錄,否則會造成資料異常 Ex:113/8/7
            strExc(0) = "select * from caseprogress where cp01='" & stCP01 & "' and cp02='" & stCP02 & "' and cp03='" & stCP03 & "' and cp04='" & stCP04 & "' and cp10='1605' and cp05=" & stCP05
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
            'end 2024/8/19
            
               'Add by Lydia 2015/01/28 系統自動更新基本檔的繳費年度與智慧局一致
               'Modified by Morgan 2015/11/19 +FCP也要補繳費年度,下一程序不要更新或上N(有不續辦但要管制的情形且不一定是6個月)--江如玉
               'If stCP01 = "P" And Val(.TextMatrix(iRow, 11)) < Val(.TextMatrix(iRow, 5)) Then
               If Val(.TextMatrix(iRow, 11)) < Val(.TextMatrix(iRow, 5)) Then
               'end 2015/11/19
                  strExc(2) = "": strExc(3) = "": strExc(4) = ""
                  For intI = Val(.TextMatrix(iRow, 11)) To Val(.TextMatrix(iRow, 5)) - 1 '減催繳的年度
                     If intI = 1 Then
                        strExc(2) = "1": strExc(3) = "19221111"
                     Else
                        strExc(2) = strExc(2) & "," & intI
                        strExc(3) = strExc(3) & ",19221111"
                        strExc(4) = strExc(4) & ","
                     End If
                  Next intI
                  strExc(0) = "update patent set PA72=PA72||'" & strExc(2) & "',PA73=PA73||'" & strExc(3) & "',PA74=PA74||'" & strExc(4) & "' where pa01='" & stCP01 & "' and pa02='" & stCP02 & "' " & _
                              "and pa03='" & stCP03 & "' and pa04='" & stCP04 & "' "
                  cnnConnection.Execute strExc(0)
                  'Added by Lydia 2015/06/03 +下一程序的法限,所限及備註更新 ex:P-81782(1040306.XLS),原閉卷後又收文
                  strExc(0) = " select NP01,NP06,NP07,NP08,NP09,NP15,NP22 from nextprogress WHERE NP02='P' and (NP06 is null or NP06='N') and NP07 in ('605','601') and NP09<=" & Val(.TextMatrix(iRow, 6)) + 19110000 & _
                              " and np03='" & stCP02 & "' and np04='" & stCP03 & "' and np05='" & stCP04 & "' order by np09 desc "
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                        strExc(9) = DBDATE(Val(.TextMatrix(iRow, 6))) '繳納期限,注意與本所資料不符會在後面加*號
                        strExc(8) = PUB_GetOurDeadline(strExc(9))
                       '下一程序的期限(本所管制的繳費年度)是否續辦=N
                       'strExc(0) = "update nextprogress set np06='N' WHERE NP02='P' and NP06 is null and NP07 in ('605','601') and NP09<=" & Val(.TextMatrix(iRow, 6)) + 19110000 & _
                                   " and np03='" & stCP02 & "' and np04='" & stCP03 & "' and np05='" & stCP04 & "' "
                       strExc(10) = ChangeWStringToTDateString(strSrvDate(1)) & "年費逾期補繳通知更新期限,原法限:" & ChangeWStringToTDateString(RsTemp.Fields("np09")) & _
                                    IIf(Len(RsTemp.Fields("np15")) > 0, ";" & Trim(RsTemp.Fields("np15")), "")
                       strExc(0) = "update nextprogress set np06='N',np08=" & CNULL(strExc(8), True) & ",np09=" & CNULL(strExc(9), True) & ",np15='" & strExc(10) & "'" & _
                                   " where NP02='P' and np03='" & stCP02 & "' and np04='" & stCP03 & "' and np05='" & stCP04 & "' and np22=" & RsTemp.Fields("np22")
                       cnnConnection.Execute strExc(0)
                  End If
                  'end 2015/06/03
               End If
               'end Add by Lydia 2015/01/28
            
            'Removed by Morgan 2024/8/19 移到上面
            'strExc(0) = "select * from caseprogress where cp01='" & stCP01 & "' and cp02='" & stCP02 & "' and cp03='" & stCP03 & "' and cp04='" & stCP04 & "' and cp10='1605' and cp05=" & stCP05
            'intI = 1
            'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            'If intI = 0 Then
            'end 2024/8/19
            
               stCP07 = DBDATE(.TextMatrix(iRow, 7))
               strDate(1) = stCP01     '系統別
               strDate(2) = "000"     '申請國家
               strDate(3) = stCP07  '法定期限
               GetCtrlDT strDate()
               If stCP01 = "P" Then
                  stCP06 = PUB_GetWorkDay1(strDate(0), True)
               Else
                  stCP06 = strDate(0)
               End If
               
               '收文號
               stCP09 = AutoNo("C", 6)
               stCP13 = PUB_GetAKindSalesNo(stCP01, stCP02, stCP03, stCP04)
               stCP12 = GetSalesArea(stCP13)
               strExc(3) = "" 'Added by Lydia 2019/06/17
               If stCP01 = "P" Then
                  stCP20 = "N"
               Else
                  'Modified by Lydia 2019/06/17 +是否閉卷銷卷(closecase)+年費逾期補繳通知函是否寄發
                  'Modified by Lydia 2023/07/28 FCP專利連結通知PA177
                  strExc(0) = "select pa26||pa27||pa28||pa29||pa30,pa75,pa57||pa108||pa167 as closecase,PA177 from patent where pa01='" & stCP01 & "' and pa02='" & stCP02 & "' and pa03='" & stCP03 & "' and pa04='" & stCP04 & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     strExc(3) = "" & RsTemp("closecase") 'Added by Lydia 2019/06/17 'Move by Lydia 2019/07/02 從下GetCP20移上來
                     stCP20 = PUB_GetCP20(stCP01, "1605", stCP16, "" & RsTemp(0), "" & RsTemp(1), stCP01 & stCP02 & stCP03 & stCP04)
                     m_PA177 = "" & RsTemp.Fields("PA177") 'Added by Lydia 2023/07/28
                  End If
               End If
               stCP16 = Val("" & stCP16)
               stCP64 = "未繳年度:" & Val(.TextMatrix(iRow, 5)) & ",原繳費期限:" & Val(.TextMatrix(iRow, 6))
               'Modified by Morgan 2012/11/5 +CP119
               'Modified by Lydia 2019/05/31 外專程序工作大項先不上發文日(整批發文)
               'strSql = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10" & _
                  ",cp12,cp13,cp20,cp26,cp32,CP64,CP16,CP17,CP18,CP119 ) values ('" & stCP01 & "'" & _
                  ",'" & stCP02 & "','" & stCP03 & "','" & stCP04 & "'," & stCP05 & _
                  "," & stCP06 & "," & stCP07 & ",'" & stCP09 & "','1605','" & stCP12 & "'" & _
                  ",'" & stCP13 & "','" & stCP20 & "','N','N'" & _
                  ",'" & ChgSQL(stCP64) & "'," & stCP16 & ",0," & stCP16 / 1000 & "," & stCP05 & ")"
               If stCP01 = "FCP" Then
                   strCP14 = Pub_GetSpecMan("外專程序-通知年費逾期")
                   'Added by Lydia 2019/06/17 已上閉卷的案件，各項大批進度檔發文日請先上111111
                   If strExc(3) <> "" Then
                       strExc(3) = "19221111"
                   Else
                   'end 2019/06/17
                       strCP48 = CompDate(2, 14, strSrvDate(1))
                   End If 'end 2019/06/17
               'Added by Morgan 2025/3/10
               Else
                  strCP14 = .TextMatrix(iRow, 15)
               'end 2025/3/10
               End If
               'Modified by Lydia 2019/06/17 已上閉卷的案件，各項大批進度檔發文日請先上111111
               'strSql = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10" & _
                  ",cp12,cp13,CP14,cp20,cp26,cp32,CP64,CP16,CP17,CP18,CP119,CP48 ) values ('" & stCP01 & "'" & _
                  ",'" & stCP02 & "','" & stCP03 & "','" & stCP04 & "'," & stCP05 & _
                  "," & stCP06 & "," & stCP07 & ",'" & stCP09 & "','1605','" & stCP12 & "'" & _
                  ",'" & stCP13 & "'," & CNULL(strCP14) & ",'" & stCP20 & "','N','N'" & _
                  ",'" & ChgSQL(stCP64) & "'," & stCP16 & ",0," & stCP16 / 1000 & "," & stCP05 & ", " & CNULL(strCP48, True) & " )"
               strSql = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10" & _
                  ",cp12,cp13,CP14,cp20,cp26,cp32,CP64,CP16,CP17,CP18,CP119,CP48,CP27 ) values ('" & stCP01 & "'" & _
                  ",'" & stCP02 & "','" & stCP03 & "','" & stCP04 & "'," & stCP05 & _
                  "," & stCP06 & "," & stCP07 & ",'" & stCP09 & "','1605','" & stCP12 & "'" & _
                  ",'" & stCP13 & "'," & CNULL(strCP14) & ",'" & stCP20 & "','N','N'" & _
                  ",'" & ChgSQL(stCP64) & "'," & stCP16 & ",0," & stCP16 / 1000 & "," & stCP05 & ", " & CNULL(strCP48, True) & ", " & CNULL(strExc(3), True) & " )"
               'end 2019/05/31
               cnnConnection.Execute strSql, intI
               
               PUB_DualCaseInform stCP09 'Added by Morgan 2022/4/7
               
               'Added by Lydia 2023/07/28 外專-FCP專利連結案管制：輸入通知年費逾繳1605自動收文「通知資訊變更961」,發一封Email給承辦工程師
               If stCP01 = "FCP" And m_PA177 = "Y" Then
                  strExc(1) = stCP01: strExc(2) = stCP02: strExc(3) = stCP03: strExc(4) = stCP04
                  If PUB_GetFCPlinkMC("4", stCP05, strExc, stCP09, "605", "1605", stCP12, stCP13) = True Then
                  End If
               End If
               'end 2023/07/28
            End If
         End If
      End If
   Next
   End With
   
   cnnConnection.CommitTrans
   AddCP = True
   Exit Function

ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
   
End Function

Private Function DoPrint(ByRef p_Grid As MSHFlexGrid, Optional ByVal p_Sys As String = "P") As Boolean
   Dim iRow As Integer, iRecs As Integer
   ReDim strTemp(9)
   Dim stPID As String 'Added by Morgan 2025/1/20
   
On Error GoTo ErrHnd

   GetPleft
   iRecs = 0
   iPage = 1
   With p_Grid
      For iRow = 1 To .Rows - 1
         If .TextMatrix(iRow, 9) = "" Or .TextMatrix(iRow, 9) = p_Sys Then
            'Added by Morgan 2025/1/20
            If stPID <> .TextMatrix(iRow, 15) Then
               If iRecs > 0 Then
                  Call PrintReportFooter(iRecs)
                  GetPleft
                  iRecs = 0
                  iPage = 1
               End If
               
               stPID = .TextMatrix(iRow, 15)
               If stPID <> "" Then
                  m_Title = m_TitleS & "(" & GetStaffName(stPID, True) & ")"
               End If
               PrintPageHeader
               PrintPageHeader1
            End If
            'end 2025/1/20
            
            iRecs = iRecs + 1
            If iRecs = 1 Then
               PrintPageHeader
               PrintPageHeader1
            End If
            strTemp(1) = .TextMatrix(iRow, 0)
            strTemp(2) = .TextMatrix(iRow, 1)
            strTemp(3) = .TextMatrix(iRow, 2)
            strTemp(4) = Left(.TextMatrix(iRow, 3), 9)
            strTemp(5) = Left(.TextMatrix(iRow, 4), 13)
            strTemp(6) = .TextMatrix(iRow, 5)
            strTemp(7) = .TextMatrix(iRow, 6)
            strTemp(8) = .TextMatrix(iRow, 7)
            strTemp(9) = " " & .TextMatrix(iRow, 10)
            PrintDetail
         End If
      Next
      If iRecs > 0 Then
         Call PrintReportFooter(iRecs)
      End If
   End With
   DoPrint = True
   Exit Function
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function
'本所資料對應電子檔資料
'p_CNo:本所號,p_jRows:最後資料範圍,p_iRows:起始資料範圍
Private Function FindData(ByVal p_CNo As String, ByVal p_jRows As Integer, Optional ByRef p_iRows As Integer = 1) As Boolean
   Dim iRow As Integer
   For iRow = p_iRows To p_jRows
      If MSHFlexGrid1.TextMatrix(iRow, 0) = p_CNo Then
         p_iRows = iRow + 1
         FindData = True
         Exit For
      ElseIf MSHFlexGrid1.TextMatrix(iRow, 0) < p_CNo Then
         p_iRows = iRow
         Exit For
      Else
         p_iRows = iRow + 1
      End If
   Next
End Function

'電子檔資料對應本所資料
'p_Type:資料種類1=正常2逾期
Private Sub CompareData(ByRef p_Grid As MSHFlexGrid, ByVal p_iRow As Integer, Optional ByVal p_Type As Integer = 1)

   Dim bol_Del As Boolean

   'Modify by Morgan 2007/3/27 追加案不用特別排除，因為申請號不同，若有重複則需改基本資料。
   '考慮可能有延期或不續辦固定抓期限最大的來比
   strSql = "select pa01,pa02,pa03,pa04,pa22,pa72,np09-19110000 np09,ai01,np06,pa57,pa108" & _
      " from patent,nextprogress,annuityinform where pa11='" & p_Grid.TextMatrix(p_iRow, 1) & "'" & _
      " and np02(+)=pa01 and np03(+)=pa02 and np04(+)=pa03 and np05(+)=pa04 and np07(+)='605' and ai01(+)=np22 order by np09 desc"

   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         bol_Del = True
         '本所號
         p_Grid.TextMatrix(p_iRow, 0) = "" & .Fields("PA01") & "-" & .Fields("PA02") & "-" & .Fields("PA03") & "-" & .Fields("PA04")
         'Add by Morgan 2010/3/10
         If Not IsNull(.Fields("pa57")) Then
            p_Grid.TextMatrix(p_iRow, 0) = p_Grid.TextMatrix(p_iRow, 0) & "＊"
         End If
         If Not IsNull(.Fields("pa108")) Then
            p_Grid.TextMatrix(p_iRow, 0) = p_Grid.TextMatrix(p_iRow, 0) & "●"
         End If
         'end 2010/3/10
         '證書號
         'Modified by Morgan 2017/9/29 Excel的證書號不再帶第1碼的英文,增加若本所第1碼不是數字時只抓第2碼以後的來比對(舊案可能只有數字故原判斷也要保留)--陳玲玲
         'If p_Grid.TextMatrix(p_iRow, 2) <> "" & .Fields("PA22") Then
         If p_Grid.TextMatrix(p_iRow, 2) = "" & .Fields("PA22") Or (IsNumeric(Left("" & .Fields("PA22"), 1)) = False And Mid("" & .Fields("PA22"), 2) = p_Grid.TextMatrix(p_iRow, 2)) Then
         Else
         'end 2017/9/29
            bol_Del = False
            p_Grid.TextMatrix(p_iRow, 2) = p_Grid.TextMatrix(p_iRow, 2) & "*"
         End If
         '繳納年次
         p_Grid.TextMatrix(p_iRow, 11) = UBound(Split("" & .Fields("PA72"), ",")) + 2
         If Val(p_Grid.TextMatrix(p_iRow, 5)) <> Val(p_Grid.TextMatrix(p_iRow, 11)) Then
            bol_Del = False
            If Val(p_Grid.TextMatrix(p_iRow, 5)) > Val(p_Grid.TextMatrix(p_iRow, 11)) Then
               p_Grid.TextMatrix(p_iRow, 5) = p_Grid.TextMatrix(p_iRow, 5) & "*"
            Else
               p_Grid.TextMatrix(p_iRow, 5) = p_Grid.TextMatrix(p_iRow, 5) & "<"
            End If
         End If
         
         '繳納期限
         If p_Grid.TextMatrix(p_iRow, 6) <> "" & .Fields("NP09") Then
            bol_Del = False
            p_Grid.TextMatrix(p_iRow, 6) = p_Grid.TextMatrix(p_iRow, 6) & "*"
         End If
         '系統別
         p_Grid.TextMatrix(p_iRow, 9) = "" & .Fields("PA01")
         
         '內專2006/01/01以後沒催過年費的也要印
         If p_Grid.TextMatrix(p_iRow, 9) = "P" And Val(txtMailDate) >= 20060101 Then
            If "" & .Fields("ai01") = "" Then
               bol_Del = False
               p_Grid.TextMatrix(p_iRow, 0) = p_Grid.TextMatrix(p_iRow, 0) & "⊙"
            End If
         End If
         
         'Add by Morgan 2010/3/8 正常若已續辦的也要印
         If p_Type = 1 And .Fields("np06") = "Y" Then
            bol_Del = False
         End If
         
         p_Grid.TextMatrix(p_iRow, 10) = "" & .Fields("np06")
         
         'Add by Morgan 2010/3/8
         p_Grid.TextMatrix(p_iRow, 12) = "" & .Fields("PA02")
         p_Grid.TextMatrix(p_iRow, 13) = "" & .Fields("PA03")
         p_Grid.TextMatrix(p_iRow, 14) = "" & .Fields("PA04")
         'end 2010/3/8
      End If
      '逾期的不刪都要印
      If p_Type = 1 And bol_Del = True Then
         '刪除旗標
         p_Grid.TextMatrix(p_iRow, 8) = "1"
      End If
   End With
End Sub

Private Sub cmdok_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Dim iRecs1 As Integer, iRecs2 As Integer
   Select Case Index
      Case 0
         If LoadXLS(iRecs1, iRecs2) = True Then
            ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/2 清除查詢印表記錄檔欄位
            If Len(txtBaseDate) <> 0 Then
               pub_QL05 = pub_QL05 & ";" & Label8 & txtBaseDate 'Add By Sindy 2010/12/2
            End If
            If Len(txtMailDate) <> 0 Then
               pub_QL05 = pub_QL05 & ";" & Label1 & txtMailDate 'Add By Sindy 2010/12/2
            End If
            If iRecs1 = 0 Then
               MsgBox "正常年費檔案內無資料！", vbExclamation
            Else
               pub_QL05 = pub_QL05 & ";正常年費;" & Label2 & txtPath(0) 'Add By Sindy 2010/12/2
               pub_QL05 = pub_QL05 & ";" & Label3 & txtDateLimit(0) & "-" & txtDateLimit(1) 'Add By Sindy 2010/12/2
               Process1
            End If
            If iRecs2 = 0 Then
               MsgBox "逾期年費檔案內無資料！", vbExclamation
            Else
               pub_QL05 = pub_QL05 & ";逾期年費;" & Label7 & txtPath(1) 'Add By Sindy 2010/12/2
               pub_QL05 = pub_QL05 & ";" & Label6 & txtDateLimit(2) & "-" & txtDateLimit(3) 'Add By Sindy 2010/12/2
               Process2
            End If
            MsgBox "作業完成！"
         End If
      Case 1
         Unload Me
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  
   MoveFormToCenter Me
   SetData
   SetFile
   'MSHFlexGrid1.Visible = True
   'MSHFlexGrid2.Visible = True
   'Added by Morgan 2015/7/13
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   KillTemp
   'end 2015/7/13
End Sub

Private Sub KillTemp()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
End Sub

Private Sub SetData()
   If txtBaseDate = "" Then
      '每個月5,20兩天通知
      '20號
      If Val(Right(strSrvDate(2), 2)) >= 20 Then
         txtBaseDate = strSrvDate(2) \ 100 & "20"
      '5號
      ElseIf Val(Right(strSrvDate(2), 2)) >= 5 Then
         txtBaseDate = strSrvDate(2) \ 100 & "05"
      '上月20
      Else
         txtBaseDate = TransDate(CompDate("1", -1, strSrvDate(1) \ 100 & "20"), 1)
      End If
   End If
   
   '5
   If Right(txtBaseDate, 2) = "05" Then
      '正常=2個月後的20~月底
      txtDateLimit(0) = TransDate(CompDate("1", 2, txtBaseDate \ 100 & "20"), 1)
      txtDateLimit(1) = TransDate(CompDate("2", -1, CompDate("1", 3, txtBaseDate \ 100 & "01")), 1)
      '逾期=3個月後的20~月底
      txtDateLimit(2) = TransDate(CompDate("1", 3, txtBaseDate \ 100 & "20"), 1)
      txtDateLimit(3) = TransDate(CompDate("2", -1, CompDate("1", 4, txtBaseDate \ 100 & "01")), 1)
   '20
   Else
      '正常=3個月後的01~19
      txtDateLimit(0) = TransDate(CompDate("1", 3, txtBaseDate \ 100 & "01"), 1)
      txtDateLimit(1) = TransDate(CompDate("1", 3, txtBaseDate \ 100 & "19"), 1)
      '逾期=4個月後的01~19
      txtDateLimit(2) = TransDate(CompDate("1", 4, txtBaseDate \ 100 & "01"), 1)
      txtDateLimit(3) = TransDate(CompDate("1", 4, txtBaseDate \ 100 & "19"), 1)
   End If
   
   txtMailDate = txtBaseDate
   
End Sub

Private Sub SetFile()
   'Modified by Morgan 2017/3/31 改新版Excel格式 .xls->.xlsx
   'Modified by Lydia 2024/07/22 改成變數設定
   'txtPath(0) = "\\PAT1\C\年費通知電子檔\正常年費\" & txtMailDate & ".xlsx"
   'txtPath(1) = "\\PAT1\C\年費通知電子檔\逾期年費\" & txtMailDate & ".xlsx"
   txtPath(0) = "\\" & strPat1Path & "\C\年費通知電子檔\正常年費\" & txtMailDate & ".xlsx"
   txtPath(1) = "\\" & strPat1Path & "\C\年費通知電子檔\逾期年費\" & txtMailDate & ".xlsx"
   'end 2024/07/22
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Morgan 2022/4/7
   Set frm040331 = Nothing
End Sub

Private Sub txtBaseDate_Validate(Cancel As Boolean)
   SetData
End Sub

Private Sub txtDateLimit_GotFocus(Index As Integer)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtDateLimit(Index).IMEMode = 2
   CloseIme
   TextInverse txtDateLimit(Index)
End Sub

Private Sub txtMailDate_GotFocus()
   CloseIme
   TextInverse txtMailDate
End Sub

Private Sub txtBaseDate_GotFocus()
   CloseIme
   TextInverse txtBaseDate
End Sub

Private Sub txtMailDate_Validate(Cancel As Boolean)
   SetFile
End Sub

Private Sub txtPath_GotFocus(Index As Integer)
   TextInverse txtPath(Index)
End Sub

Sub GetPleft()
   Printer.Orientation = 2
   m_iTitleFontSize = 22
   m_iFontSize = 12
   m_iStartX = 500
   m_iStartY = 500
   m_iPageHeight = Printer.ScaleHeight
   m_iLineHeight = 300
   m_iMargin = 500
   
   ReDim PLeft(10)
   PLeft(0) = 500
   '本所案號(2500)
   PLeft(1) = 500
   '申請案號(1400)
   PLeft(2) = PLeft(1) + 2500
   '證書號(1200)
   PLeft(3) = PLeft(2) + 1400
   '專利權人(2600)
   PLeft(4) = PLeft(3) + 1200
   '專利名稱(3400)
   PLeft(5) = PLeft(4) + 2600
   '繳納年次(1200)
   PLeft(6) = PLeft(5) + 3400
   '繳納期限(1200)
   PLeft(7) = PLeft(6) + 1200
   '加倍補繳期限(1200)
   PLeft(8) = PLeft(7) + 1200
   PLeft(9) = PLeft(8) + 1500
   PLeft(10) = PLeft(9) + 600
End Sub

Sub PrintPageHeader()
   iPrint = m_iStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = m_iTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(m_Title)) / 2
   Printer.CurrentY = iPrint
   Printer.Print m_Title
   iPrint = iPrint + 500
   Printer.Font.Size = m_iFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   PrintNewLine
   
   Printer.CurrentX = 6500
   Printer.CurrentY = iPrint
   Printer.Print "電子郵件檔案日期：" & txtMailDate
   
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = Printer.ScaleWidth - m_iMargin - 2500
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   PrintNewLine
   
   Printer.CurrentX = 6500
   Printer.CurrentY = iPrint
   If m_iType = 1 Then
      Printer.Print "法定期限：" & txtDateLimit(0) & " - " & txtDateLimit(1)
   Else
      Printer.Print "法定期限：" & txtDateLimit(2) & " - " & txtDateLimit(3)
   End If
   
   Printer.CurrentX = Printer.ScaleWidth - m_iMargin - 2500
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
   PrintNewLine
End Sub

Sub PrintPageHeader1()

   Call PrintNewLine(False, 1)
   
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "申請案號"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "證書號"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "專利權人"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "專利名稱"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "繳納年次"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "繳納期限"
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = iPrint
   Printer.Print "加倍補繳期限"
   'Add by Morgan 2008/5/12
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = iPrint
   Printer.Print "續辦"
   'end 2008/5/12
   PrintNewLine
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print String(132, "-")
    
End Sub
'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)

   Call PrintNewLine(True, 3)
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print String(132, "-")
   PrintNewLine
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "合計： " & iRecCount & " 筆"
   PrintMemo
   Printer.EndDoc
   
End Sub

Sub PrintDetail()

   Dim iCol As Integer
   
   PrintNewLine
   For iCol = LBound(strTemp) To UBound(strTemp)
      Printer.CurrentX = PLeft(iCol)
      Printer.CurrentY = iPrint
      Printer.Print strTemp(iCol)
   Next
    
End Sub

Private Sub PrintMemo()
   Printer.Font.Size = 10
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = m_iPageHeight - m_iMargin - Printer.TextHeight("註")
   Printer.Print "註：1.本所案號空白表示無本所資料 2.只有本所案號其他空白表示未通知 3.非本所案號欄位有加*號者表示與本所資料不符"
   Printer.CurrentX = m_iStartX
   Printer.Print "　　4.本所案號加＊號表示已閉卷 5.本所案號加●號表示已銷卷 6.本所案號加⊙號表示尚未催繳 7.繳納年次加<表示小於本所紀錄"
   Printer.Font.Size = m_iFontSize
End Sub

Private Sub PrintNewLine(Optional ByVal p_bolHeader1 As Boolean = True, Optional ByVal p_iExtraLines As Integer = 3)

   iPrint = iPrint + m_iLineHeight
   If iPrint >= (m_iPageHeight - m_iMargin - p_iExtraLines * m_iLineHeight) Then
      Printer.CurrentX = m_iStartX
      Printer.CurrentY = iPrint
      Printer.Print String(132, "-")
      PrintMemo
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      If p_bolHeader1 Then
         PrintPageHeader1
      End If
      iPrint = iPrint + m_iLineHeight
   End If
   
End Sub
