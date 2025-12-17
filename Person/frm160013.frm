VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm160013 
   BorderStyle     =   1  '單線固定
   Caption         =   "新建指紋整批匯入"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7950
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束"
      Height          =   345
      Left            =   6750
      TabIndex        =   6
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "下載刷卡紀錄"
      Height          =   345
      Index           =   1
      Left            =   270
      TabIndex        =   5
      Top             =   510
      Width           =   1545
   End
   Begin VB.CommandButton Command1 
      Caption         =   "批次建檔"
      Height          =   345
      Index           =   4
      Left            =   6300
      TabIndex        =   4
      Top             =   510
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      Caption         =   "員工號自動匹配"
      Height          =   345
      Index           =   3
      Left            =   4545
      TabIndex        =   3
      Top             =   510
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查詢待建檔指紋刷卡紀錄"
      Height          =   345
      Index           =   2
      Left            =   1980
      TabIndex        =   2
      Top             =   510
      Width           =   2445
   End
   Begin VB.CommandButton Command1 
      Caption         =   "指紋自動匯入"
      Height          =   345
      Index           =   0
      Left            =   270
      TabIndex        =   1
      Top             =   90
      Width           =   1545
   End
   Begin MSFlexGridLib.MSFlexGrid grdList 
      Height          =   3705
      Left            =   180
      TabIndex        =   0
      Top             =   960
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6535
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      FillStyle       =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frm160013"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/6/17 Form2.0已修改
'Created by Morgan 2013/7/15
Option Explicit

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Command1_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Select Case Index
   Case 0
      InitialGridList
      If PollingData(True) = True Then
         If QueryData(True) = True Then
            If AutoMatch(True) = True Then
               AutoCreate
            End If
         End If
      End If
   Case 1 '下載刷卡紀錄
      PollingData
   Case 2 '查詢待建檔指紋刷卡紀錄
      QueryData
   Case 3 '員工號自動匹配
      AutoMatch
   Case 4 '批次建檔
      AutoCreate
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Function AutoMatch(Optional pbolAuto As Boolean) As Boolean
   Dim ii As Integer, strLstNo As String
   If grdList.Rows < 2 Then
      MsgBox "無待匹配資料！"
   Else
      grdList.TopRow = 1
      For ii = 1 To grdList.Rows - 1
         grdList.row = ii
         If ii Mod 15 = 0 Then grdList.TopRow = ii
         If grdList.TextMatrix(ii, 0) = "" Then
            If strLstNo = grdList.TextMatrix(ii, 4) Then
               grdList.TextMatrix(ii, 0) = grdList.TextMatrix(ii - 1, 0)
               grdList.TextMatrix(ii, 1) = grdList.TextMatrix(ii - 1, 1)
            Else
            
               strLstNo = grdList.TextMatrix(ii, 4)
               If Len(strLstNo) = 4 Then
                  intI = (Format(Now, "YYYY") - 1911 - 100) \ 10
                  strLstNo = Chr(Asc("A") + intI) & strLstNo
               ElseIf Left(strLstNo, 1) < "6" Then
                  strLstNo = Chr(Asc("A") + Val(Left(strLstNo, 1)) - 1) & Mid(strLstNo, 2)
               End If
               'Modified by Morgan 2019/4/26 +st60
               strExc(0) = "select st01,st02,st04,nvl(st60,st02) st60 from staff where st01='" & strLstNo & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If RsTemp("st04") = "1" Then
                     grdList.TextMatrix(ii, 0) = RsTemp("st01")
                     grdList.TextMatrix(ii, 1) = RsTemp("st02")
                     grdList.TextMatrix(ii, 7) = RsTemp("st60") 'Added by Morgan 2019/4/26
                     strLstNo = grdList.TextMatrix(ii, 4)
                  Else
                     grdList.TextMatrix(ii, 6) = "已離職"
                  End If
               Else
                  grdList.TextMatrix(ii, 6) = "匹配失敗"
               End If
            End If
         End If
      Next
      AutoMatch = True
   End If
End Function

Private Sub AutoCreate()
   Dim ii As Integer, strLstNo As String
   Dim stFinger1 As String, stFinger2 As String, iRlt As Integer
   Dim jj As Integer, stDomain As String
   Dim arrIpList
   Dim okList As String, errList1 As String, errList2 As String, errList3 As String, errList4 As String, errIpList As String
   Dim iOkCount As Integer, iErrCount As Integer, stMsg As String
   Dim bolWriteOk As Boolean
   
   HTAips = GetHtaIP()
   If HTAips = "" Then
      MsgBox "考勤機IP未設定！"
      Exit Sub
   End If
   
   arrIpList = Split(HTAips, ";")
   
   If grdList.Rows < 2 Then
      MsgBox "無待建檔資料！", vbInformation
   Else
      grdList.TopRow = 1
      For ii = 1 To grdList.Rows - 1
         grdList.row = ii
         If ii Mod 15 = 0 Then grdList.TopRow = ii '顯示目前執行資料
         If grdList.TextMatrix(ii, 0) <> "" And grdList.TextMatrix(ii, 6) = "" Then
            If strLstNo = grdList.TextMatrix(ii, 4) Then
               grdList.col = 0
               grdList.CellForeColor = vbBlack
               grdList.col = 1
               grdList.CellForeColor = vbBlack
            Else
               strLstNo = grdList.TextMatrix(ii, 4)
               '設定指紋來源IP
               If grdList.TextMatrix(ii, 5) = "" Then
                  iErrCount = iErrCount + 1
                  errList1 = errList1 & grdList.TextMatrix(ii, 0) & "(" & grdList.TextMatrix(ii, 4) & ")" & grdList.TextMatrix(ii, 1) & vbCrLf
                  grdList.TextMatrix(ii, 6) = "指紋來源考勤機IP不明"
                  'If MsgBox("無法讀取 " & strLstNo & "(" & grdList.TextMatrix(ii, 1) & ") 的指紋，來源考勤機的IP未知!!是否要繼續??", vbYesNo + vbDefaultButton2 + vbCritical) = vbNo Then
                  '   Exit Function
                  'End If
               Else
                  HTAip = grdList.TextMatrix(ii, 5)
                  'Added by Morgan 2013/8/1
                  '讀指紋指定在執行檔跑時若連線後馬上呼叫會錯,所有故意連線後斷線
                  If HTAconnect() = False Then
                     iErrCount = iErrCount + 1
                     errList2 = errList2 & grdList.TextMatrix(ii, 0) & "(" & grdList.TextMatrix(ii, 4) & ")" & grdList.TextMatrix(ii, 1) & vbCrLf
                     grdList.TextMatrix(ii, 6) = "指紋讀取失敗"
                  Else
                     HTAclose
                  'end 2013/8/1
                     '讀取指紋
                     If HTAqueryFingerPrinter(strLstNo, stFinger1, stFinger2) = True Then
                        '更新資料庫
                        If SaveCard(grdList.TextMatrix(ii, 0), grdList.TextMatrix(ii, 0), stFinger1, stFinger2, True, strLstNo) = True Then
                           grdList.col = 0
                           grdList.CellForeColor = vbBlack
                           grdList.col = 1
                           grdList.CellForeColor = vbBlack
                           
                           stDomain = GetDomain(HTAip)
                           errIpList = ""
                           For jj = LBound(arrIpList) To UBound(arrIpList)
                              If arrIpList(jj) <> "" Then
                                 bolWriteOk = True
                                 If stDomain = GetDomain(arrIpList(jj)) Then
                                    bolWriteOk = False
                                    HTAip = arrIpList(jj)
                                    iRlt = 0
                                    If HTAqueryCard(strLstNo, , True) = True Then
                                       If HTAdeleteCard(strLstNo) = False Then
                                          iRlt = 1
                                       End If
                                    End If
                                    If iRlt = 0 Then
                                       'Modified by Morgan 2019/4/26 改用顯示姓名
                                       If HTAaddFingerPrinter(grdList.TextMatrix(ii, 0), grdList.TextMatrix(ii, 7), stFinger1, stFinger2) = True Then
                                          bolWriteOk = True
                                       End If
                                    End If
                                    If bolWriteOk = False Then
                                       errIpList = errIpList & "," & HTAip
                                    End If
                                 End If
                              End If
                           Next
                           If bolWriteOk = True And errIpList = "" Then
                              iOkCount = iOkCount + 1
                              okList = okList & grdList.TextMatrix(ii, 0) & "(" & grdList.TextMatrix(ii, 4) & ")" & grdList.TextMatrix(ii, 1) & vbCrLf
                              grdList.TextMatrix(ii, 6) = "OK"
                           Else
                              iErrCount = iErrCount + 1
                              errList4 = errList4 & grdList.TextMatrix(ii, 0) & "(" & grdList.TextMatrix(ii, 4) & ")" & grdList.TextMatrix(ii, 1) & "(失敗IP:" & Mid(errIpList, 2) & ")" & vbCrLf
                              grdList.TextMatrix(ii, 6) = "回寫考勤機失敗"
                              'If MsgBox(strLstNo & "(" & grdList.TextMatrix(ii, 1) & ") 的指紋回寫考勤機失敗!!是否要繼續??", vbYesNo + vbDefaultButton2 + vbCritical) = vbNo Then
                              '   Exit Function
                              'End If
                           End If
                        Else
                           iErrCount = iErrCount + 1
                           errList3 = errList3 & grdList.TextMatrix(ii, 0) & "(" & grdList.TextMatrix(ii, 4) & ")" & grdList.TextMatrix(ii, 1) & vbCrLf
                           grdList.TextMatrix(ii, 6) = "寫入資料庫失敗"
                           'If MsgBox(strLstNo & "(" & grdList.TextMatrix(ii, 1) & ") 的指紋寫入資料庫失敗!!是否要繼續??", vbYesNo + vbDefaultButton2 + vbCritical) = vbNo Then
                           '   Exit Function
                           'End If
                        End If
                     Else
                        iErrCount = iErrCount + 1
                        errList2 = errList2 & grdList.TextMatrix(ii, 0) & "(" & grdList.TextMatrix(ii, 4) & ")" & grdList.TextMatrix(ii, 1) & vbCrLf
                        grdList.TextMatrix(ii, 6) = "指紋讀取失敗"
                        'If MsgBox("無法讀取 " & strLstNo & "(" & grdList.TextMatrix(ii, 1) & ") 的指紋!!是否要繼續??", vbYesNo + vbDefaultButton2 + vbCritical) = vbNo Then
                        '   Exit Function
                        'End If
                     End If
                  End If
               End If
            End If
         End If
      Next
      
      stMsg = "匯入完畢! 成功 " & iOkCount & " 筆" & IIf(iErrCount > 0, "，失敗 " & iErrCount & " 筆", "") & "。"
      
      If iOkCount > 0 Then
         stMsg = stMsg & vbCrLf & vbCrLf & "成功清單入下：" & vbCrLf & okList
      End If
      If iErrCount > 0 Then
         stMsg = stMsg & vbCrLf & vbCrLf & "失敗清單入下："
      End If
      If errList1 <> "" Then
         stMsg = stMsg & vbCrLf & "*指紋來源考勤機IP不明：" & vbCrLf & errList1
      End If
      If errList2 <> "" Then
         stMsg = stMsg & vbCrLf & "*指紋讀取失敗：" & vbCrLf & errList2
      End If
      If errList3 <> "" Then
         stMsg = stMsg & vbCrLf & "*寫入資料庫失敗：" & vbCrLf & errList3
      End If
      If errList4 <> "" Then
         stMsg = stMsg & vbCrLf & "*回寫考勤機失敗(資料庫已寫入)：" & vbCrLf & errList4
      End If
      
      
      WriteLog stMsg
      MsgBox stMsg, vbInformation
   End If
End Sub

'寫記錄
Private Function WriteLog(oStrLog As String)
   Dim ffa As Integer
   ffa = FreeFile
   Open App.path & "\" & App.EXEName & ".log" For Append As ffa
   Print #ffa, Trim(Now) & "  ==>  " & oStrLog
   Close ffa
End Function

Private Function GetDomain(pIP) As String
   Dim iPos As Integer
   iPos = InStrRev(pIP, ".")
   If iPos > 1 Then
      GetDomain = Left(pIP, iPos - 1)
   End If
End Function

Private Function PollingData(Optional pbolAuto As Boolean) As Boolean
   Dim iRecs As Integer, iRecTot As Integer
   Dim arrIpList
   Dim ii As Integer
   Dim bResult As Boolean
   Dim iRtn As Integer
   Dim iTimes As Integer
   
   HTAips = GetHtaIP()
   If HTAips = "" Then
      MsgBox "考勤機IP未設定！"
      Exit Function
   End If
   
   arrIpList = Split(HTAips, ";")
   For ii = LBound(arrIpList) To UBound(arrIpList)
      HTAip = arrIpList(ii)
      If HTAip <> "" Then
         
         Pub_WriteSysLog "開始下載...(" & HTAip & ")"
         
         'Modified by Morgan 2015/3/12 改自動重試 3 次並可手動再重試(人事電腦第1次連線若有刷卡記錄會無法下載)
         iTimes = 1
         bResult = False
         bResult = HTAPolling(iRecs, True)
         Do While (bResult = False And iTimes < 3)
            Sleep 3000
            iTimes = iTimes + 1
            bResult = HTAPolling(iRecs, True)
         Loop
         
         If bResult = True Then
            iRecTot = iRecTot + iRecs
         Else
            iRtn = MsgBox("考勤機(" & HTAip & ") 刷卡紀錄接收失敗！是否要重試??", vbYesNoCancel + vbDefaultButton3 + vbCritical)
            If iRtn = vbCancel Then
               Exit For
            ElseIf iRtn = vbYes Then
               ii = ii - 1
            Else
               bResult = True
            End If
         End If
         'end 2015/3/12
      End If
   Next
   PollingData = bResult
   
   If pbolAuto = False Then
      MsgBox "已接收 " & iRecTot & " 筆!", vbInformation
   End If
End Function

Private Function QueryData(Optional pbolAuto As Boolean) As Boolean
   InitialGridList
   '抓1週內的刷卡紀錄
   'Modified by Morgan 2019/4/26 +st60
   strExc(0) = "select '' st01,'' st02" & _
      ",sqldatet(substr(max(pr01||pr02),1,8)) dt,sqltime(substr(max(pr01||pr02),9)) tm,pr03,min(pr09) ip" & _
      ",'' Rlt,'' st60 from pollrecord,staffcarddata where pr01>to_char(sysdate-7,'yyyymmdd')  and pr08='9'" & _
      " and scd02(+)=pr03 And scd01 is null" & _
      " group by pr03"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      UpdateGridList RsTemp
      QueryData = True
   Else
      MsgBox "無待建檔指紋刷卡紀錄!", vbInformation
   End If
End Function

Private Sub UpdateGridList(ByRef rsTmp As ADODB.Recordset)
   Dim iRow As Integer, iCol As Integer
   With rsTmp
   .MoveFirst
   grdList.Visible = False
   Do While Not .EOF
      grdList.Rows = grdList.Rows + 1
      iRow = grdList.Rows - 1
      grdList.row = iRow
      For iCol = 0 To .Fields.Count - 1
         grdList.col = iCol
         If iCol < 2 Then
            grdList.CellForeColor = vbRed
         End If
         grdList.TextMatrix(iRow, iCol) = "" & .Fields(iCol)
      Next
      .MoveNext
   Loop
   grdList.Visible = True
   End With
End Sub


' 初始化列表
Private Sub InitialGridList()
   With grdList
   .Clear
   .Rows = 1
   .Cols = 8
    
   .row = 0
   .col = 0
   .Text = "員工號"
   .CellAlignment = flexAlignCenterCenter
   .ColWidth(.col) = 660
   .ColAlignment(.col) = flexAlignCenterCenter
   
   .col = 1
   .Text = "姓名"
   .CellAlignment = flexAlignCenterCenter
   .ColWidth(.col) = 1000
   .ColAlignment(.col) = flexAlignCenterCenter
    
   .col = 2
   .Text = "刷卡日期"
   .CellAlignment = flexAlignCenterCenter
   .ColWidth(.col) = 1000
   .ColAlignment(.col) = flexAlignCenterCenter
       
   .col = 3
   .Text = "刷卡時間"
   .CellAlignment = flexAlignCenterCenter
   .ColWidth(.col) = 900
   .ColAlignment(.col) = flexAlignCenterCenter
    
   .col = 4
   .Text = "卡號"
   .CellAlignment = flexAlignCenterCenter
   .ColWidth(.col) = 700
   .ColAlignment(.col) = flexAlignCenterCenter
    
   .col = 5
   .Text = "IP"
   .CellAlignment = flexAlignCenterCenter
   .ColWidth(.col) = 1200
   .ColAlignment(.col) = flexAlignCenterCenter
   
   .col = 6
   .Text = "結果"
   .CellAlignment = flexAlignCenterCenter
   .ColWidth(.col) = 1750
   .ColAlignment(.col) = flexAlignCenterCenter
   
   .col = 7
   .Text = "顯示姓名"
   .CellAlignment = flexAlignCenterCenter
   .ColWidth(.col) = 0
   .ColAlignment(.col) = flexAlignCenterCenter
   End With
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   InitialGridList
   HTAips = GetHtaIP()
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160013 = Nothing
End Sub

Private Function SaveCard(pSCD01 As String, pSCD02 As String, pSCD03 As String, pSCD04 As String, Optional pUpdateRecord As Boolean, Optional pPR03 As String) As Boolean
   Dim stSQL As String, iRtn As Integer
   
On Error GoTo ErrHnd

   cnnConnection.BeginTrans

   stSQL = "update staffcarddata set scd03='" & pSCD03 & "',scd04='" & pSCD04 & "' where scd02='" & pSCD02 & "'"
   cnnConnection.Execute stSQL, iRtn
   If iRtn = 0 Then
      stSQL = "insert into staffcarddata(scd01,scd02,scd03,scd04) values ('" & pSCD01 & "','" & pSCD02 & "','" & pSCD03 & "','" & pSCD04 & "')"
      cnnConnection.Execute stSQL, iRtn
   End If
   If iRtn = 1 Then
      If pUpdateRecord = True Then
         stSQL = "update pollrecord set pr03='" & pSCD01 & "' where pr03='" & pPR03 & "'"
         cnnConnection.Execute stSQL, iRtn
      End If
      SaveCard = True
   End If
   cnnConnection.CommitTrans
   
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function

Private Sub grdList_DblClick()
   Dim iRow As Integer, strTmp As String, strNew As String
   
   iRow = grdList.row
   If iRow > 0 And grdList.col = 0 Then
      If grdList.CellForeColor = vbRed Then
         strTmp = grdList.TextMatrix(iRow, 0)
         strNew = InputBox("請輸入員工編號!!", "手動匹配", strTmp)
         If strNew <> "" Then
            'Modified by Morgan 2019/4/26 +st60
            strExc(0) = "select st02,scd01,nvl(st60,st02) st60 from staff,staffcarddata where st01='" & strNew & "' and scd01(+)=st01 and scd02(+)=st01"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If Not IsNull(RsTemp("scd01")) Then
                  MsgBox strNew & "(" & RsTemp("st02") & ") 指紋檔已存在，不可再新增!!請到維護作業更新!!", vbCritical
               Else
                  grdList.TextMatrix(iRow, 0) = strNew
                  grdList.TextMatrix(iRow, 1) = RsTemp("st02")
                  grdList.TextMatrix(iRow, 6) = ""
                  grdList.TextMatrix(iRow, 7) = RsTemp("st60") 'Added by Morgan 2019/4/26
               End If
            Else
               MsgBox "員工編號輸入錯誤!!", vbCritical
            End If
         End If
      End If
   End If
End Sub

