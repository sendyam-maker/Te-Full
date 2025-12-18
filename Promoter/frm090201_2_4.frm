VERSION 5.00
Begin VB.Form frm090201_2_4 
   Caption         =   "開庭/面詢紀錄上傳"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1920
   ScaleWidth      =   8880
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox textCPA06 
      Height          =   270
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "搬檔"
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "存檔(&S)"
      Height          =   435
      Index           =   0
      Left            =   6300
      TabIndex        =   6
      Top             =   90
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   435
      Index           =   1
      Left            =   7530
      TabIndex        =   5
      Top             =   90
      Width           =   1200
   End
   Begin VB.ListBox lstAtt 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1128
      ItemData        =   "frm090201_2_4.frx":0000
      Left            =   720
      List            =   "frm090201_2_4.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   630
      Width           =   7305
   End
   Begin VB.CommandButton cmdRemAtt 
      Caption         =   "-> 移除"
      Height          =   315
      Left            =   8010
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdAddAtt 
      Caption         =   "<- 新增"
      Height          =   345
      Left            =   8010
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdOpenAtt 
      Caption         =   "開啟"
      Height          =   315
      Left            =   8010
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   630
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "附件："
      Height          =   180
      Index           =   7
      Left            =   135
      TabIndex        =   4
      Top             =   660
      Width           =   540
   End
End
Attribute VB_Name = "frm090201_2_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/4 改成Form2.0 (無)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Created by Morgan 2012/4/9
Option Explicit

Public m_Key As String '收文號(FTP資料夾)
Dim m_bDelete As Boolean
Dim m_Appendix As String

Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194

Private Const cTableName As String = "CASEPROGRESSAPPENDIX" 'Added by Lydia 2017/08/09 指定FTP資料夾名稱

Private Sub SetListScroll()
   Dim s As String
   Static x As Long
   
   s = lstAtt.List(lstAtt.ListIndex)
   If x < TextWidth(s & "  ") Then
      x = TextWidth(s & "  ")
     If ScaleMode = vbTwips Then _
         x = x / Screen.TwipsPerPixelX  ' if twips change to pixels
     SendMessageByNum lstAtt.hWnd, LB_SETHORIZONTALEXTENT, x, 0
   End If
End Sub

Private Sub cmdAddAtt_Click()
   PUB_OpenDialog4List lstAtt
End Sub

Private Sub cmdOK_Click(Index As Integer)
   If Index = 0 Then
      Screen.MousePointer = vbHourglass
      If FormSave = False Then
         Screen.MousePointer = vbDefault
         MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
         Exit Sub
      End If
      Screen.MousePointer = vbDefault
   End If
   Unload Me
End Sub

'開啟附件
Private Sub cmdOpenAtt_Click()
'Added by Lydia 2017/08/09
Dim tmpArr As Variant, ii As Integer
Dim stFileName As String
Dim hLocalFile As Long
'end 2017/08/09

   If lstAtt.Text = "" Then
      MsgBox "請選擇欲開啟的附件！"
   Else
      'Added by Lydia 2017/08/09 判斷移檔日期
      If strSrvDate(1) >= CR_NewDate And textCPA06 <> "" Then
         tmpArr = Empty
         tmpArr = Split(textCPA06.Text, ",")
         ii = lstAtt.ListIndex
         If ii > UBound(tmpArr) Then Exit Sub
         If Trim(tmpArr(ii)) <> "" Then
            strExc(1) = Trim(Mid(lstAtt.Text, 1, InStrRev(lstAtt.Text, " (") - 1))
            stFileName = App.path & "\$$" & strExc(1)
            If PUB_GetFtpFile(Trim(tmpArr(ii)), stFileName, cTableName) Then
                ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
            End If
         End If
         
      'Removed by Morgan 2024/8/2 不用的標記為註解，檢查程式碼才知時可略過
      'Else
      ''end 2017/08/09
      '   PUB_OpenFtpFile m_Key, lstAtt.Text, , 案件進度附件存放路徑
      'end 2024/8/2
      
      End If 'end 2017/08/09
   End If
End Sub

Private Sub cmdRemAtt_Click()
   If InStr(lstAtt, "\") = 0 And m_bDelete = False Then
         MsgBox "已上傳檔案不可移除！"
   ElseIf PUB_RemoveList(lstAtt) = True Then
      cmdAddAtt.SetFocus
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   LoadAppendix
   
   'Added by Lydia 2017/08/09
   If Pub_StrUserSt03 <> "M51" Then cmd1.Visible = False
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

   PUB_KillTempFile "$$*.*" 'Added by Lydia 2017/08/09 清除暫存檔
   
   Set frm090201_2_4 = Nothing
End Sub

Private Sub LoadAppendix()
   lstAtt.Clear
   textCPA06.Text = "" 'Added by Lydia 2017/08/09
   m_Appendix = ""
   strExc(0) = "select * from CaseProgressAppendix where cpa01='" & m_Key & "' order by cpa04,cpa05"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         lstAtt.AddItem .Fields("cpa02"), lstAtt.ListCount
         m_Appendix = m_Appendix & "," & .Fields("cpa02")
         textCPA06 = textCPA06 & IIf(textCPA06 <> "", ",", "") & .Fields("cpa06")   'Added by Lydia 2017/08/09 FTP路徑
         .MoveNext
      Loop
      m_Appendix = Mid(m_Appendix, 2)
      End With
   End If
End Sub

Private Function FormSave() As Boolean
   Dim stSQL As String, stFileName As String
   Dim strNewFiles As String
   Dim arrFile1
   Dim ii As Integer, bolRemove As Boolean
   Dim iErr As Integer, sErrMsg As String

On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   For ii = 0 To lstAtt.ListCount - 1
      If InStr(lstAtt.List(ii), "\") > 0 Then
         stFileName = Mid(lstAtt.List(ii), InStrRev(lstAtt.List(ii), "\") + 1)
         stSQL = "insert into CaseProgressAppendix(cpa01,cpa02,cpa03,cpa04,cpa05)" & _
            " values('" & m_Key & "','" & ChgSQL(stFileName) & "','" & strUserNum & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss'))"
         cnnConnection.Execute stSQL, intI
      End If
   Next
   
   '上傳附件檔
   If PUB_UploadAtt(案件進度附件存放路徑, m_Key, lstAtt, iErr, sErrMsg) = False Then
      GoTo ErrHand
   End If
   
   '檔案有異動時，移掉的要刪除
   strNewFiles = PUB_ComposeAttList(lstAtt)
   bolRemove = False
   If strNewFiles <> m_Appendix Then
      arrFile1 = Split(m_Appendix, ",")
      For ii = LBound(arrFile1) To UBound(arrFile1)
         If InStr(strNewFiles & ",", arrFile1(ii) & ",") > 0 Then
            arrFile1(ii) = ""
         Else
            stSQL = "delete CaseProgressAppendix where cpa01='" & m_Key & "' and cpa02='" & ChgSQL(arrFile1(ii)) & "'"
            cnnConnection.Execute stSQL, intI
            bolRemove = True
         End If
      Next
      
      If bolRemove = True Then
         If PUB_RemoveAtt(案件進度附件存放路徑, m_Key, Join(arrFile1, ","), iErr, sErrMsg) = False Then
            GoTo ErrHand
         End If
         
      End If
   End If
   
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   ElseIf iErr <> 0 Then
      MsgBox sErrMsg, vbCritical
   End If
End Function

Private Sub lstAtt_Click()
   SetListScroll
End Sub

'Added by Lydia 2017/08/09 搬檔
'Removed by Morgan 2024/8/2 不用的標記為註解，檢查程式碼才知時可略過
'Private Sub Cmd1_Click()
'Dim stSQL As String, intR As Integer
'Dim rsQuery As ADODB.Recordset
'Dim stOldDir As String, stNewDir As String, stNewPath As String
'Dim oFileName As String, mFileName As String
'Dim strGrp As String, strList As String, strNameList As String
'Dim tmpArr As Variant
'Dim strTmpExc As String
'Dim stDownFile As String
'Dim strLost As String, strLostId As String
'
'   stOldDir = 案件進度附件存放路徑
'   stNewDir = PUB_GetFtpTableDir(stNewDir) & cTableName
'   stSQL = "select CPA01,CPA02 from CASEPROGRESSAPPENDIX " & _
'           "where NVL(CPA02,'N') <> 'N' AND NVL(CPA06,'N')='N' ORDER BY CPA01,CPA04,CPA05,CPA02 "
'
'   intR = 0
'   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
'   If intR = 1 Then
'      With rsQuery
'         MsgBox "開始工作，共" & .RecordCount & "筆記錄!"
'         .MoveFirst
'         Do While Not .EOF
'            '清除暫存檔
'            PUB_KillTempFile "$$*.*"
'
'            If strGrp <> "" & .Fields("CPA01") & .Fields("CPA02") Then
'               If strGrp <> "" Then
'                  strTmpExc = strTmpExc & "UPDATE CASEPROGRESSAPPENDIX SET CPA06='" & strList & "' WHERE CPA01||CPA02='" & strGrp & "' ;"
'               End If
'
'               If Left(strGrp, 9) <> "" & .Fields("CPA01") Then
'                  strNameList = ""
'               End If
'               strList = ""
'               strGrp = "" & .Fields("CPA01") & .Fields("CPA02")
'               tmpArr = Empty
'               tmpArr = Split("" & .Fields("CPA02"), ",")
'            End If
'
'            For intR = 0 To UBound(tmpArr)
'               If Trim(tmpArr(intR)) <> "" Then
'                   '先下載檔案
'                   stDownFile = ""
'                   '因為有附件檔名有包含刮號,直接到模組處理舊檔名
'                   strExc(1) = PUB_StringFilter(Trim(tmpArr(intR)))
'                   If InStr(strExc(1), "(") > 0 And InStr(strExc(1), " (") = 0 Then
'                      strExc(1) = Mid(strExc(1), 1, InStrRev(strExc(1), "(") - 1) & " " & Mid(strExc(1), InStrRev(strExc(1), "("))
'                   End If
'                   PUB_OpenFtpFile .Fields("CPA01"), strExc(1), , 案件進度附件存放路徑, False, stDownFile
'
'                   If stDownFile = "" Then
'                       strLostId = strLostId & .Fields("CPA01") & "," & IIf(Len(strLostId) > 50, vbCrLf, "")
'                       strLost = strLost & .Fields("CPA01") & "_" & Trim(tmpArr(intR)) & vbCrLf
'                   Else
'                        oFileName = Trim(tmpArr(intR))
'                        oFileName = Trim(Mid(oFileName, 1, InStrRev(oFileName, "(") - 1))
'                        '新-FTP檔名(非中文)
'                        mFileName = PUB_GetNewFileNameSec(oFileName, "2", strNameList, "" & .Fields("CPA01"))
'
'                        If PUB_PutFtpFile(stDownFile, "" & .Fields("CPA01"), mFileName, stNewPath, cTableName) = True Then
'                           strList = strList & IIf(strList <> "", ",", "") & stNewPath
'                        Else
'                           MsgBox "Error !"
'                           Exit Sub
'                        End If
'                   End If
'               End If
'            Next intR
'            .MoveNext
'         Loop
'
'         '最後一筆
'         strTmpExc = strTmpExc & "UPDATE CASEPROGRESSAPPENDIX SET CPA06='" & strList & "' WHERE CPA01||CPA02='" & strGrp & "' ;"
'      End With
'
'      '清除暫存檔
'      PUB_KillTempFile "$$*.*"
'
'      If strTmpExc <> "" Then
'         tmpArr = Empty
'         tmpArr = Split(strTmpExc, ";")
'         cnnConnection.BeginTrans
'           For intR = 0 To UBound(tmpArr)
'              If Trim(tmpArr(intR)) <> "" Then
'                 cnnConnection.Execute Trim(tmpArr(intR)), intI
'              End If
'           Next intR
'         cnnConnection.CommitTrans
'         MsgBox "工作結束!"
'      End If
'   End If
'
'   If strLost <> "" Then
'      PUB_SendMail "QPGMR", "A3034", "", 案件進度附件存放路徑 & "在NT2缺少檔案", "資料夾:" & strLostId & vbCrLf & vbCrLf & "檔案名稱:" & strLost
'   End If
'
'   Set rsQuery = Nothing
'   Exit Sub
'
'ErrHandle:
'   cnnConnection.RollbackTrans
'
'OutPort:
'   Exit Sub
'
'End Sub


