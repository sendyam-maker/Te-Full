VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm010035_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "個人記錄"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   1935
   ClientWidth     =   7560
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7560
   Begin VB.CommandButton CmdOK 
      Caption         =   "延期"
      Default         =   -1  'True
      Height          =   400
      Index           =   3
      Left            =   2520
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   840
      Width           =   675
   End
   Begin VB.Frame Frame1 
      Height          =   710
      Left            =   120
      TabIndex        =   9
      Top             =   675
      Width           =   2250
      Begin VB.CommandButton CmdOK 
         Caption         =   "刪除"
         Height          =   400
         Index           =   2
         Left            =   1520
         Style           =   1  '圖片外觀
         TabIndex        =   12
         Top             =   180
         Width           =   675
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "歸還"
         Height          =   400
         Index           =   1
         Left            =   780
         Style           =   1  '圖片外觀
         TabIndex        =   11
         Top             =   180
         Width           =   675
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "同意"
         Height          =   400
         Index           =   0
         Left            =   60
         Style           =   1  '圖片外觀
         TabIndex        =   10
         Top             =   180
         Width           =   675
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "全部"
      Height          =   180
      Index           =   2
      Left            =   2520
      TabIndex        =   8
      Top             =   300
      Value           =   -1  'True
      Width           =   1000
   End
   Begin VB.OptionButton Option1 
      Caption         =   "個人保管"
      Height          =   180
      Index           =   1
      Left            =   1200
      TabIndex        =   7
      Top             =   300
      Width           =   1300
   End
   Begin VB.OptionButton Option1 
      Caption         =   "借閱中"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   300
      Width           =   1000
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找"
      Height          =   400
      Left            =   4815
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束"
      Height          =   400
      Left            =   6675
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "書籍資料"
      Height          =   400
      Index           =   4
      Left            =   5505
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   840
      Width           =   950
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "借閱記錄"
      Height          =   400
      Index           =   5
      Left            =   6465
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   840
      Width           =   950
   End
   Begin VB.CommandButton cmdPrePage 
      Caption         =   "回前畫面"
      Height          =   400
      Left            =   5670
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   120
      Width           =   950
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   2700
      Left            =   120
      TabIndex        =   5
      Top             =   1490
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   4763
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "V| 編號|書名|ISBN|作者|譯者|類別|保管人|狀態|借閱人|借閱/延期日|上架日|出刊日"
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
      _Band(0).Cols   =   13
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7320
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frm010035_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/07/27 Form2.0已修改 GrdDataList
'2016/10/03 Create by Amy
Option Explicit

Public cmdState As Integer
Public bolLoanRecordApply As Boolean '保管人是否需簽核
Dim i As Integer
Dim arrField, intWidth

Private Sub cmdExit_Click()
    Unload Me
    frm010035.Show
    frm010035.PubShowNextData
End Sub

Private Sub cmdok_Click(Index As Integer)
    cmdState = Index
    PubShowNextData
End Sub

Private Sub cmdPrePage_Click()
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
End Sub

Public Sub cmdSearch_Click()
    Dim RsQ As New ADODB.Recordset
    Dim strQ1 As String, strQ2 As String, strQ As String
    
    '個人借閱中(排除 遺失/銷毀)
    strQ1 = "Select * From LoanRecord a ,BooksData Where LR08='" & strUserNum & "' " & _
                "And LR01||LR02=(Select Max(LR01||LR02) as LR01 From LoanRecord Where a.LR03=LR03) " & _
                "And Not Exists(Select * From LoanRecord b Where a.LR03=b.LR03(+) " & _
                "And LR02='X' And LR01=(Select Max(LR01) as LR01 From LoanRecord Where b.LR03=LR03)) " & _
                "And Not Exists(Select * From LoanRecord c Where a.LR03=c.LR03(+) " & _
                "And (LR02='Y' Or LR02='Z')  And LR01=(Select Max(LR01) as LR01 From LoanRecord Where c.LR03=LR03)) " & _
                "And BK01=LR03(+)  "
    '個人保管(排除 遺失/銷毀)
    strQ2 = "Select * From LoanRecord a ,BooksData Where BK10='" & strUserNum & "' And BK01=LR03(+) " & _
                "And LR01||LR02=(Select Max(LR01||LR02) as LR01 From LoanRecord Where a.LR03=LR03) " & _
                "And LR02<>'Y' And LR02<>'Z' "
    
    If Option1(0).Value = True Then
        strQ = strQ1
    ElseIf Option1(1).Value = True Then
        strQ = strQ2
    Else
        strQ = strQ1 & vbCrLf & " Union " & strQ2
    End If
    strQ = "Select '' as V,BK01 as 編號,Decode(BK04,null,BK05,Decode(BK05,null,BK04,BK04||'('||BK05||')')) as 書名,Nvl(BK02,'') as ISBN," & _
                "Decode(BK06,null,BK07,Decode(BK07,null,BK06,BK06||'('||BK07||')')) as 作者,Nvl(BK08,'') as 譯者," & _
                "Decode(BK03,'1','專利','2','商標','3','法律','4','電腦','5','其他') as 類別,k.ST02 as 保管人," & _
                "Decode(LR06,null,Decode(LR02, '1','借閱申請中', 'X','可借閱', 'Y','遺失', 'Z','銷毀', '延期申請中'), Decode(LR02, 'X','可借閱', 'Y','遺失', 'Z','銷毀', '借閱中')) as  狀態," & _
                "Decode(LR02,'X','',b.ST02) as 借閱人,Decode(LR02,'X','',Decode(LR06,null,'',sqldatet(LR04))) as 借閱日," & _
                "sqldatet(BK09) as 上架日,Decode(BK13,null,'',sqldatet(BK13)) as 出刊日, " & _
                "Nvl(LR06,'') as LR06,BK10,O.LR01,LR02,LR04,P.LR07,LR08 " & _
                "From (" & strQ & ") O,Staff k,Staff b, " & _
                "(Select LR01,LR07 From LoanRecord Where LR02='1') P " & _
                "Where BK10=k.ST01(+) And LR08=b.ST01(+) And O.LR01=P.LR01(+) And LR02<>'Y' And LR02<>'Z' " & _
                "Order by Decode(InStr(bk12,'申請中'),0,2,1),BK01"
   
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    grdDataList.Clear
    grdDataList.FixedCols = 0
    grdDataList.Rows = 2
    If RsQ.RecordCount > 0 Then
        Set grdDataList.Recordset = RsQ
    End If
    SetGridColor
    SetGridWidth
    grdDataList.FixedCols = 3
    grdDataList.ColAlignmentFixed(GetValue("編號")) = flexAlignLeftCenter
    grdDataList.ColAlignmentFixed(GetValue("書名")) = flexAlignLeftCenter
    RsQ.Close
    Set RsQ = Nothing
End Sub

Private Sub Form_Load()
    ReDim arrField(18)
    ReDim intWidth(18)
    arrField = Array("V", "編號", "書名", "ISBN", "作者", "譯者", "類別", "保管人", "狀態", "借閱人", "借閱/延期日", "上架日", "出刊日", _
                                "LR06", "BK10", "LR01", "LR02", "LR04", "LR07", "LR08")
    intWidth = Array(200, 500, 1300, 1000, 1000, 1000, 500, 700, 1000, 700, 1000, 1000, 1000, _
                                0, 0, 0, 0, 0, 0, 0)
    
    MoveFormToCenter Me
    grdDataList.FixedCols = 3 '固定前3欄
End Sub

Private Sub SetGridWidth()
   
    '設欄寬
    With grdDataList
        .FormatString = .FormatString
        For i = LBound(intWidth) To UBound(intWidth)
            .ColWidth(i) = intWidth(i)
            If intWidth(i) <> 0 Then .ColAlignment(i) = flexAlignLeftCenter
        Next i
    End With
End Sub

Private Sub SetGridColor()
    Dim j As Integer
    
    '設狀態欄為綠色
    With grdDataList
        .Visible = False
        For i = 1 To .Rows - 1
            .row = i
            If InStr(.TextMatrix(i, GetValue("狀態")), "申請中") > 0 Then
                .col = GetValue("狀態")
                .CellBackColor = &HC000&
            End If
        Next i
        .Visible = True
    End With
End Sub

Private Function GetValue(strField As String) As Integer
    Dim jj As Integer
 
    For jj = 1 To UBound(arrField)
       If UCase(arrField(jj)) = UCase(strField) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

Private Function EditRecord(CmdIdx As Integer, stLR01 As String, stLR02 As String, stLR03 As String, ByVal bolSelf As Boolean, Optional stDate As String = "") As Boolean
    Dim strExe As String
    
On Error GoTo ErrHand
    
    EditRecord = False
    Select Case CmdIdx
        Case 0 '同意
            strExe = "Update LoanRecord set LR05=" & stDate & ",LR06=" & strSrvDate(1) & " Where LR01='" & stLR01 & "' And LR02='" & stLR02 & "' "
            cnnConnection.Execute strExe
        Case 1 '歸還
            strExe = ServerTime
            strExe = IIf(Len(strExe) = 6, Left(strExe, 4), Left(strExe, 3))
            strExe = "Insert Into LoanRecord (LR01,LR02,LR03,LR08,LR09,LR10) Values(" & _
                        "'" & stLR01 & "' ,'X','" & stLR03 & "','" & strUserNum & "'," & strSrvDate(1) & ",'" & strExe & "')"
            cnnConnection.Execute strExe
        Case 2 '刪除
            strExe = "Delete From LoanRecord Where LR01='" & stLR01 & "' And LR02='" & stLR02 & "' "
            cnnConnection.Execute strExe
        Case 3 '延期
            strExe = ServerTime
            strExe = IIf(Len(strExe) = 6, Left(strExe, 4), Left(strExe, 3))
            If bolSelf = True Then
                cnnConnection.BeginTrans
                '保管人延期自動上保管人確認日
                strExe = "Insert Into LoanRecord (LR01,LR02,LR03,LR04,LR06,LR08,LR09,LR10) Values(" & _
                              "'" & stLR01 & "' ,'" & stLR02 & "','" & stLR03 & "'," & strSrvDate(1) & "," & strSrvDate(1) & ",'" & strUserNum & "'," & strSrvDate(1) & "," & strExe & ")"
                cnnConnection.Execute strExe
                '更新應還日
                strExe = DBDATE(DateAdd("m", 1, Format(strSrvDate(1), "####/##/##")))
                strExe = "Update LoanRecord set LR05=" & strExe & " Where LR01='" & stLR01 & "' And LR02='1' "
                cnnConnection.Execute strExe
                cnnConnection.CommitTrans
            Else
                strExe = "Insert Into LoanRecord (LR01,LR02,LR03,LR04,LR08,LR09,LR10) Values(" & _
                              "'" & stLR01 & "' ,'" & stLR02 & "','" & stLR03 & "'," & strSrvDate(1) & ",'" & strUserNum & "'," & strSrvDate(1) & "," & strExe & ")"
            cnnConnection.Execute strExe
            End If
    End Select
    
    EditRecord = True
    Exit Function
    
ErrHand:
    If CmdIdx = 3 And bolSelf = True Then cnnConnection.RollbackTrans
    MsgBox "程式有誤請洽電腦中心！" & vbCrLf & Err.Description
End Function

Private Sub Form_Unload(Cancel As Integer)
    bolLoanRecordApply = False
    Set frm010035_2 = Nothing
End Sub

Private Sub GrdDataList_Click()

    With grdDataList
        .Visible = False
        .col = 0
        If .row <> 0 Then
            If .Text = "V" Then
                .Text = ""
                For i = GetValue("ISBN") To .Cols - 1
                    .col = i
                    If i = GetValue("狀態") Then
                        If InStr(.TextMatrix(.row, GetValue("狀態")), "申請中") = 0 Then .CellBackColor = QBColor(15)
                    Else
                        .CellBackColor = QBColor(15)
                    End If
                Next i
            Else
                .Text = "V"
                For i = GetValue("ISBN") To .Cols - 1
                    .col = i
                    If i = GetValue("狀態") Then
                        If InStr(.TextMatrix(.row, GetValue("狀態")), "申請中") = 0 Then .CellBackColor = &HFFC0C0
                    Else
                        .CellBackColor = &HFFC0C0
                    End If
                Next i
            End If
        End If
        .Visible = True
    End With
End Sub

Public Sub PubShowNextData()
    Dim j As Integer
    Dim bolAccept As Boolean
    Dim strMsg As String, strTemp(2) As String
    Dim strTo As String, strContent As String, strSubject As String
    Dim stLR02 As String, stLR05 As String
                                                            
    With grdDataList
        For i = 1 To .Rows - 1
            bolAccept = False: strMsg = "": strTemp(0) = "": strTemp(1) = "": strTemp(2) = ""
            strTo = "": strContent = "": strSubject = ""
            stLR02 = "": stLR05 = ""
            
            If .TextMatrix(i, 0) = "V" Then
                .TextMatrix(i, 0) = ""
                .row = i
                For j = GetValue("ISBN") To .Cols - 1
                    .col = j
                    .CellBackColor = QBColor(15)
                Next j
                If cmdState > 3 Then
                    If fnSaveParentForm(Me) = False Then
                       Me.Enabled = True
                       Exit Sub
                    End If
                '遺失/銷毀 不可按 同意/歸還/刪除/延期 鈕
                ElseIf .TextMatrix(i, GetValue("狀態")) = "遺失" Or .TextMatrix(i, GetValue("狀態")) = "銷毀" Then
                    Exit Sub
                'Add by Amy 2020/02/11 可借閱不需按 同意/歸還/刪除/延期 鈕 ex:狀態為可借閱, 再按歸還 Insert 語法會錯
                ElseIf cmdState <= 3 And .TextMatrix(i, GetValue("狀態")) = "可借閱" Then
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                Select Case cmdState
                    Case 0
                        strMsg = "同意"
                        'Add by Amy 2020/02/11 借閱中 不需按 同意 鈕
                        If .TextMatrix(i, GetValue("狀態")) = "借閱中" Then
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                        'end 2020/0211
                        '保管人可操作
                        If .TextMatrix(i, GetValue("BK10")) = strUserNum And .TextMatrix(i, GetValue("LR08")) <> strUserNum Then
                            bolAccept = True
                        End If
                    Case 1
                        strMsg = "歸還"
                        '保管人可操作
                        'Modify by Amy 2020/02/11 原:<>"借閱申請中"
                        If .TextMatrix(i, GetValue("BK10")) = strUserNum And .TextMatrix(i, GetValue("狀態")) = "借閱中" Then
                            bolAccept = True
                        End If
                    Case 2
                        strMsg = "刪除"
                        '保管人只能 刪除 借閱中/借閱申請中/延期申請中 資料
                        If .TextMatrix(i, GetValue("BK10")) = strUserNum And _
                          (.TextMatrix(i, GetValue("狀態")) = "借閱中" Or .TextMatrix(i, GetValue("狀態")) = "借閱申請中" Or .TextMatrix(i, GetValue("狀態")) = "延期申請中") Then
                            If .TextMatrix(i, GetValue("狀態")) = "借閱中" And .TextMatrix(i, GetValue("LR01")) = GetMinLR01(.TextMatrix(i, GetValue("編號"))) Then
                                strMsg = "系統產生之第一筆閱記錄不可刪除！"
                            Else
                                bolAccept = True
                            End If
                        End If
                    Case 3
                        strMsg = "延期"
                        '借閱人可操作
                        'Modify by Amy 2020/02/11 原:<>"借閱申請中"
                        If .TextMatrix(i, GetValue("LR08")) = strUserNum And .TextMatrix(i, GetValue("狀態")) = "借閱中" Then
                            bolAccept = True
                        End If
                    Case Else
                        bolAccept = True
                End Select
                
                If bolAccept = False Then
                    If Len(strMsg) = 2 Then
                        MsgBox "您無此資料的「" & strMsg & "」權限！", vbInformation
                    Else
                        MsgBox strMsg, vbInformation
                    End If
                Else
                    If cmdState <= 3 Then
                        If cmdState = 3 Then
                            stLR02 = GetLR02(.TextMatrix(i, GetValue("LR01")))
                            If stLR02 = "X" Then MsgBox "延期次數已超過上限！", vbInformation: Exit Sub
                        Else
                            stLR02 = .TextMatrix(i, GetValue("LR02"))
                        End If
                        If MsgBox("您確定「" & strMsg & "」編號 " & .TextMatrix(i, GetValue("編號")) & "  的借閱記錄嗎？", vbYesNo) = vbYes Then
                            If cmdState = 0 Then stLR05 = DBDATE(DateAdd("m", 1, Format(.TextMatrix(i, GetValue("LR04")), "####/##/##")))
                            If EditRecord(cmdState, .TextMatrix(i, GetValue("LR01")), stLR02, .TextMatrix(i, GetValue("編號")), _
                              IIf(cmdState = 3 And .TextMatrix(i, GetValue("BK10")) = strUserNum, True, False), IIf(cmdState = 0, stLR05, "")) = False Then
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            Else
                                strContent = "圖書編號：" & .TextMatrix(i, GetValue("編號")) & vbCrLf & _
                                                    "ＩＳＢＮ：" & .TextMatrix(i, GetValue("ISBN")) & vbCrLf & _
                                                    "書　　名：" & .TextMatrix(i, GetValue("書名")) & vbCrLf & _
                                                    "作　　者：" & .TextMatrix(i, GetValue("作者")) & vbCrLf
                                Select Case cmdState
                                    Case 0 '同意
                                        '借閱申請同意
                                        If .TextMatrix(i, GetValue("狀態")) = "借閱申請中" Then
                                            strTo = .TextMatrix(i, GetValue("LR08"))
                                            strSubject = "圖書借閱申請，保管人 " & .TextMatrix(i, GetValue("保管人")) & " 已同意，請向保管人取書！應還日期：" & ChangeWStringToTDateString(stLR05)
                                            strContent = strContent & vbCrLf & "應還日期：" & ChangeWStringToTDateString(stLR05)
                                        '延期申請同意,不是保管人才發mail通知借閱
                                        ElseIf .TextMatrix(i, GetValue("BK10")) = strUserNum And .TextMatrix(i, GetValue("LR08")) <> strUserNum Then
                                            strTemp(0) = GetLR05(0, .TextMatrix(i, GetValue("LR01")), strTemp(1), strTemp(2))
                                            strTo = .TextMatrix(i, GetValue("LR08"))
                                            strSubject = "圖書延期申請，保管人 " & .TextMatrix(i, GetValue("保管人")) & " 已同意延期！應還日期：" & ChangeWStringToTDateString(stLR05)
                                            strContent = strContent & "借閱日期：" & ChangeWStringToTDateString(strTemp(1)) & vbCrLf & _
                                                                "延期次數：" & strTemp(2) & " 次（含本次）" & vbCrLf & _
                                                                "延期後應還日期：" & ChangeWStringToTDateString(strTemp(0)) & vbCrLf
                                        End If
                                    Case 1 '歸還
                                        '不是保管人歸還才發mail通知借閱人
                                        If .TextMatrix(i, GetValue("BK10")) <> .TextMatrix(i, GetValue("LR08")) Then
                                            strTemp(0) = GetLR05(1, .TextMatrix(i, GetValue("LR01")), strTemp(1))
                                            strTo = .TextMatrix(i, GetValue("LR08"))
                                            strSubject = "借閱圖書已做歸還處理通知！"
                                            strContent = strContent & "應還日期：" & ChangeWStringToTDateString(strTemp(0)) & vbCrLf & _
                                                                "歸還日期：" & ChangeWStringToTDateString(strTemp(1)) & vbCrLf
                                        End If
                                    Case 3 '延期
                                        strTemp(0) = GetLR05(0, .TextMatrix(i, GetValue("LR01")), strTemp(1), strTemp(2))
                                        '保管人延期彈訊息
                                        If .TextMatrix(i, GetValue("BK10")) = strUserNum Then
                                            MsgBox "圖書編號 「" & .TextMatrix(i, GetValue("編號")) & "」已延期" & vbCrLf & _
                                                          "應還日期為 " & ChangeWStringToTDateString(strTemp(0)), vbInformation
                                        '不是保管人延期才發mail給保管人
                                        Else
                                            strTo = .TextMatrix(i, GetValue("BK10"))
                                            strSubject = "圖書借閱延期申請，請至一般作業－＞圖書借閱資料查詢　之個人記錄 確認！"
                                            strContent = strContent & "借閱日期：" & ChangeWStringToTDateString(strTemp(1)) & vbCrLf & _
                                                                "延期次數：" & strTemp(2) & " 次（含本次）" & vbCrLf & _
                                                                "延期後應還日期：" & ChangeWStringToTDateString(strTemp(0)) & vbCrLf
                                        End If
                                End Select
                                If strTo <> MsgText(601) Then
                                    PUB_SendMail strUserNum, strTo, "", strSubject, strContent
                                End If
                            End If
                        End If
                    '書籍資料
                    ElseIf cmdState = 4 Then
                        If frm010034.QueryRecord(.TextMatrix(i, GetValue("編號"))) = True Then
                            frm010034.SetParent Me
                            frm010034.ToolBarSet -1
                            frm010034.Show
                            Screen.MousePointer = vbDefault
                            Me.Enabled = True
                            Exit Sub
                        End If
                    '借閱記錄
                    Else
                        If frm010035_3.QueryRecord(.TextMatrix(i, GetValue("編號"))) = True Then
                            frm010035_3.Show
                            Screen.MousePointer = vbDefault
                            Me.Enabled = True
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Next i
        If cmdState <= 3 Then Call cmdSearch_Click: Screen.MousePointer = vbDefault
    End With
End Sub

'取得序號
Private Function GetLR02(ByVal stLR01 As String, Optional bolNowName As Boolean = False) As String
    Dim rsA As New ADODB.Recordset
    Dim stQ As String
   
    '抓LoanRecord(借閱記錄)流水號
    stQ = "Select NVL(MAX(LR02),0) as LR02 From LoanRecord Where LR01='" & stLR01 & "' "
    
    rsA.CursorLocation = adUseClient
    rsA.Open stQ, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        '取得新編號
        If bolNowName = False Then
            If IsNumeric(rsA("LR02")) And Val(rsA("LR02")) < 9 Then
                GetLR02 = Val(rsA("LR02")) + 1
            ElseIf rsA("LR02") = "9" Then
                GetLR02 = "A"
            Else
                GetLR02 = Chr(Asc(rsA("LR02")) + 1)
            End If
        '取得目前編號,回傳名稱
        Else
            Select Case "" & rsA("LR02")
                Case ""
                    GetLR02 = ""
                Case "1"
                    GetLR02 = "借閱"
                Case "X"
                    GetLR02 = "歸還"
                Case "Y"
                    GetLR02 = "遺失"
                Case "Z"
                    GetLR02 = "銷毀"
                Case Else
                    GetLR02 = "延期"
            End Select
        End If
    End If
    rsA.Close
    Set rsA = Nothing
End Function

'取得借閱資料
Private Function GetLR05(ByVal intChoose As Integer, ByVal stLR01 As String, Optional ByRef stLR1 As String = "", Optional ByRef stLR2 As String = "") As String
    Dim rsA As New ADODB.Recordset
    Dim stQ As String, stField As String
    
    GetLR05 = "": stLR1 = "": stLR2 = ""
    Select Case intChoose
        Case 0 '回傳 借閱日/延期次數
            stQ = "Select LR05,LR04,CFreq From " & _
                     "(Select LR01,LR04,LR05 From LoanRecord Where LR01='" & stLR01 & "' And LR02='1'), " & _
                     "(Select LR01 as CLR01,Count(*) as CFreq From LoanRecord Where LR01='" & stLR01 & "' And LR02>'1' And LR02<'X' Group by LR01) " & _
                     "Where LR01=CLR01(+)"
        Case 1 '回傳 應還日/歸還日
            stQ = "Select LR05,LR09 From " & _
                     "(Select LR01,LR05 From LoanRecord Where LR01='" & stLR01 & "' And LR02='1'), " & _
                     "(Select LR01 as CLR01,LR09 From LoanRecord Where LR01='" & stLR01 & "' And LR02='X') " & _
                     "Where LR01=CLR01(+)"
    End Select
    
    rsA.CursorLocation = adUseClient
    rsA.Open stQ, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        GetLR05 = "" & rsA.Fields(0)
        stLR1 = "" & rsA.Fields(1)
        If intChoose = 0 Then stLR2 = Val("" & rsA.Fields(2))
    End If
    rsA.Close
    Set rsA = Nothing
End Function

'傳入圖書編號取得第一筆借閱資記錄編號
Private Function GetMinLR01(ByVal stLR03 As String) As String
    Dim rsA As New ADODB.Recordset
    Dim stQ As String, stField As String
    
    GetMinLR01 = ""
    stQ = "Select Min(LR01) as LR01 From LoanRecord Where LR03='" & stLR03 & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open stQ, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        GetMinLR01 = "" & rsA.Fields("LR01")
    End If
    rsA.Close
    Set rsA = Nothing
End Function

Private Sub grdDataList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If grdDataList.TextMatrix(1, 1) = MsgText(601) Then Exit Sub
    
    grdDataList.ToolTipText = ""
    If grdDataList.MouseRow <> 0 And grdDataList.MouseCol > 0 Then
        If grdDataList.TextMatrix(grdDataList.MouseRow, GetValue("LR07")) <> "" Then
            grdDataList.ToolTipText = "借閱備註：" & grdDataList.TextMatrix(grdDataList.MouseRow, GetValue("LR07"))
        End If
    End If
End Sub

