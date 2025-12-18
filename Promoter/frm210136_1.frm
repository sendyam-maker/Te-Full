VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210136_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標著作案件齊備日或急件維護"
   ClientHeight    =   5750
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8950
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   345
      Left            =   420
      TabIndex        =   28
      Top             =   2550
      Width           =   7965
      Begin VB.TextBox txtEP06 
         Height          =   285
         Left            =   750
         MaxLength       =   7
         TabIndex        =   30
         Text            =   "txtEP06"
         Top             =   0
         Width           =   1275
      End
      Begin VB.TextBox txtCP122 
         Height          =   285
         Left            =   4950
         MaxLength       =   1
         TabIndex        =   29
         Top             =   0
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "齊備日："
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   33
         Top             =   30
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "是否急件：              （Y/N）"
         Height          =   255
         Index           =   7
         Left            =   4020
         TabIndex        =   32
         Top             =   30
         Width           =   2820
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "(此欄只可修改一次)"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   12
         Left            =   2070
         TabIndex        =   31
         Top             =   30
         Width           =   1560
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Height          =   1470
      Left            =   60
      TabIndex        =   34
      Top             =   2010
      Width           =   8865
      Begin VB.TextBox txtTCD05 
         Height          =   285
         Left            =   840
         TabIndex        =   39
         Top             =   1020
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtTCD04 
         Height          =   285
         Left            =   90
         TabIndex        =   38
         Top             =   1020
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   35
         Text            =   "frm210136_1.frx":0000
         Top             =   270
         Width           =   1785
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
         Height          =   1395
         Left            =   1800
         TabIndex        =   36
         Top             =   30
         Width           =   7035
         _ExtentX        =   12418
         _ExtentY        =   2469
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   10
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label1 
         Caption         =   "補充資料內容："
         Height          =   255
         Index           =   13
         Left            =   525
         TabIndex        =   37
         Top             =   30
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "電腦中心取消齊備(&C)"
      Height          =   375
      Index           =   2
      Left            =   6930
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   930
      Width           =   1980
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "存檔(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   405
      Index           =   1
      Left            =   6930
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   30
      Width           =   780
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   0
      Left            =   7770
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   30
      Width           =   780
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   1965
      Left            =   60
      TabIndex        =   3
      Top             =   3720
      Width           =   8835
      _ExtentX        =   15593
      _ExtentY        =   3457
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "備註：在資料列上快速點二下即可查看明細資料"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   14
      Left            =   4920
      TabIndex        =   40
      Top             =   3510
      Width           =   3780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "齊備日歷史記錄："
      Height          =   180
      Index           =   9
      Left            =   90
      TabIndex        =   27
      Top             =   3510
      Width           =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "(已修改過齊備日期一次，不可再修改)"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   8
      Left            =   1620
      TabIndex        =   26
      Top             =   3510
      Width           =   3000
   End
   Begin VB.Label LabCP18 
      Height          =   255
      Left            =   4680
      TabIndex        =   25
      Top             =   1680
      Width           =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "點數："
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   24
      Top             =   1680
      Width           =   930
   End
   Begin VB.Label LabCP16 
      Height          =   255
      Left            =   1350
      TabIndex        =   23
      Top             =   1680
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "費用："
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   22
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   21
      Top             =   180
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   11
      Left            =   420
      TabIndex        =   20
      Top             =   480
      Width           =   900
   End
   Begin MSForms.Label LabTM05 
      Height          =   255
      Left            =   1350
      TabIndex        =   19
      Top             =   480
      Width           =   7440
      VariousPropertyBits=   27
      Size            =   "13123;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件性質："
      Height          =   255
      Index           =   2
      Left            =   420
      TabIndex        =   18
      Top             =   780
      Width           =   900
   End
   Begin VB.Label LabCP10 
      Height          =   255
      Left            =   1350
      TabIndex        =   17
      Top             =   780
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "收文日："
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   16
      Top             =   780
      Width           =   930
   End
   Begin VB.Label LabCP05 
      Height          =   255
      Left            =   4680
      TabIndex        =   15
      Top             =   780
      Width           =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "總收文號："
      Height          =   255
      Index           =   6
      Left            =   3720
      TabIndex        =   14
      Top             =   180
      Width           =   930
   End
   Begin VB.Label LabCP09 
      Height          =   255
      Left            =   4680
      TabIndex        =   13
      Top             =   180
      Width           =   1740
   End
   Begin VB.Label LabID 
      Height          =   255
      Left            =   1350
      TabIndex        =   12
      Top             =   180
      Width           =   1740
   End
   Begin VB.Label LabCP48 
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   1080
      Width           =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "承辦期限："
      Height          =   255
      Index           =   10
      Left            =   3720
      TabIndex        =   10
      Top             =   1080
      Width           =   930
   End
   Begin MSForms.Label LabCP14 
      Height          =   255
      Left            =   1350
      TabIndex        =   9
      Top             =   1080
      Width           =   2220
      VariousPropertyBits=   27
      Size            =   "3916;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "承辦人："
      Height          =   255
      Index           =   17
      Left            =   420
      TabIndex        =   8
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label LabCP07 
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   1380
      Width           =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "法定期限："
      Height          =   255
      Index           =   19
      Left            =   3720
      TabIndex        =   6
      Top             =   1380
      Width           =   930
   End
   Begin VB.Label LabCP06 
      Height          =   255
      Left            =   1350
      TabIndex        =   5
      Top             =   1380
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "本所期限："
      Height          =   255
      Index           =   21
      Left            =   420
      TabIndex        =   4
      Top             =   1380
      Width           =   900
   End
End
Attribute VB_Name = "frm210136_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/18 改成Form2.0 (grd1,grd2,LabTM05,LabCP14)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Create By Sindy 2012/5/7
Option Explicit

'紀錄作用按鍵
Public cmdState As Integer
Dim m_EP06 As String
Dim m_CP122 As String
Dim m_CP13 As String
Dim m_CP14 As String
Dim m_CP10 As String 'Added by Lydia 2018/12/10
Public WorkType As String 'Add By Sindy 2012/10/24 1.齊備日或急件維護 2.回覆補充資料作業
Dim i As Integer, j As Integer
Public bolNotData As Boolean
'Added by Lydia 2019/05/02
Dim m_TM01 As String
Dim m_TM10 As String
Dim m_CP05 As String
Dim m_CP06 As String

Private Sub SetDataListWidth()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer

arrGridHeadText = Array("操作日期", "操作人員", "操作時間", "輸入日期", "操作動作", "通知補充資料")
arrGridHeadWidth = Array(850, 850, 850, 850, 1300, 4000)
grd1.MergeCells = flexMergeRestrictColumns
grd1.Cols = UBound(arrGridHeadText) + 1
For iRow = 0 To grd1.Cols - 1
   grd1.row = 0
   grd1.col = iRow
   grd1.Text = arrGridHeadText(iRow)
   grd1.ColWidth(iRow) = arrGridHeadWidth(iRow)
   grd1.CellAlignment = flexAlignLeftCenter
Next
End Sub

'Add By Sindy 2012/10/24
Private Sub SetDataListWidth2()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer

arrGridHeadText = Array("有資料", "無資料", "補充資料")
arrGridHeadWidth = Array(700, 700, 5000)
Grid2.MergeCells = flexMergeRestrictColumns
Grid2.Cols = UBound(arrGridHeadText) + 1
For iRow = 0 To Grid2.Cols - 1
   Grid2.row = 0
   Grid2.col = iRow
   Grid2.Text = arrGridHeadText(iRow)
   Grid2.ColWidth(iRow) = arrGridHeadWidth(iRow)
   Grid2.CellAlignment = flexAlignLeftCenter
Next
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim longSeqno As Long
Dim Cancel As Boolean
Dim strCaseNo As String
Dim bolChk As Boolean, strText As String, strTemp As Variant 'Add By Sindy 2012/10/24
'Added by Lydia 2019/12/13
Dim intJ As Integer
Dim rsAD As New ADODB.Recordset

On Error GoTo ErrHnd

cmdState = Index
Select Case cmdState
Case 1
   '齊備日或急件維護
   If WorkType = "1" Then
      'Modify By Sindy 2019/2/12 txtEP06 <> "" ==> Val(txtEP06) > 0
      If (Val(txtEP06) > 0 And m_EP06 <> txtEP06) Or m_CP122 <> txtCP122 Then
         '檢查齊備日不可＜系統日
         'If txtEP06.Enabled = True And Val(txtEP06) > 0 And m_EP06 <> txtEP06 Then
         If txtEP06.Enabled = True And (Val(txtEP06) > 0 Or Val(m_EP06) = 0) Then
      '2019/2/12 END
            If Val(DBDATE(txtEP06)) < strSrvDate(1) Then
               MsgBox "齊備日不可小於系統日！", vbExclamation
               txtEP06.SetFocus
               Exit Sub
            End If
            Cancel = False
            Call txtEP06_Validate(Cancel)
            If Cancel = True Then
               txtEP06.SetFocus
               Exit Sub
            End If
         End If
         
         cnnConnection.BeginTrans
         If txtEP06 <> "" Then
            If txtEP06.Enabled = True And m_EP06 <> txtEP06 Then
               strSql = "update engineerprogress set ep06=" & DBDATE(txtEP06) & ",ep36=" & DBDATE(txtEP06) & " where ep02='" & Trim(LabCP09.Caption) & "'"
               cnnConnection.Execute strSql
               frm210136.m_EP06 = ChangeWStringToTDateString(DBDATE(txtEP06))
               strSql = "insert into tmctldate(tcd01,tcd02,tcd03,tcd04,tcd05,tcd06,tcd07)" & _
                        " values('" & Trim(LabCP09.Caption) & "','1','" & strUserNum & "'," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & "," & DBDATE(txtEP06) & ",'智權人員維護')"
               cnnConnection.Execute strSql
               'Add By Sindy 2013/7/3 加入 齊備日異動時，紀錄
               Pub_SaveLog strUserNum, "齊備日異動：" & DBDATE(m_EP06) & "==>" & DBDATE(txtEP06) & " ", SystemNumber(LabID.Caption, 1), SystemNumber(LabID.Caption, 2), SystemNumber(LabID.Caption, 3), SystemNumber(LabID.Caption, 4), LabCP09.Caption
            End If
         End If
         If txtCP122.Enabled = True And m_CP122 <> txtCP122 Then
            strSql = "update caseprogress set cp122='" & txtCP122 & "' where cp09='" & Trim(LabCP09.Caption) & "'"
            cnnConnection.Execute strSql
         End If
         '計算承辦期限
         'Memo by Lydia 2018/12/10 商申案計算方式與商爭案相同
         'Modified by Lydia 2019/05/02　商申案計算方式用原本的方式
         'If Trim(txtEP06) = "" Then
         '   strDate = PUB_TMdebateCountCP48(LabCP06.Caption, txtCP122, m_EP06, LabCP09.Caption, m_CP13)
         'Else
         '   strDate = PUB_TMdebateCountCP48(LabCP06.Caption, txtCP122, txtEP06, LabCP09.Caption, m_CP13)
         'End If
         If LabCP48.Caption = "" Then 'Added by Lydia 2019/06/03 已有承辦期限則不變更原承辦期限
            strExc(1) = Trim(txtEP06)
            If strExc(1) = "" Then strExc(1) = m_EP06
            strExc(1) = DBDATE(strExc(1))
            'Modified by Lydia 2022/07/15 T大陸案之齊備日管控: 限制T,FCT案
            'If InStr(TMdebate, m_CP10) > 0 Then '商爭案
            'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
            If (m_TM01 = "T" Or m_TM01 = "FCT") And InStr(TMdebate, m_CP10) > 0 And Not (m_TM01 = "FCT" And InStr(FCT_NotTMdebate, m_CP10) > 0) Then
                 strDate = PUB_TMdebateCountCP48(LabCP06.Caption, txtCP122, strExc(1), LabCP09.Caption, m_CP13)
            Else     '商申案
                 strDate = Pub_GetHandleDay(m_TM01, m_TM10, m_CP10, strExc(1), DBDATE(LabCP06.Caption), LabCP09.Caption)
                 'Added by Lydia 2019/12/13 內商申請101，增加判斷查名是否已齊備;
                         '起因： T-225444原本分案時文件+查名已齊備=>已上承辦期限CP48, 後來承辦在13:28增加勾選未完成的查名單=> 查名未齊備+承辦期限CP48=null;
                         '後來智權人員在13:39誤輸入文件齊備,在沒有判斷查名是否已齊備,所以又上了承辦期限
                 'Modified by Lydia 2022/07/15 T大陸案之齊備日管控
                 'If m_TM01 = "T" And m_TM10 = "000" And m_CP10 = "101" Then
                 If m_TM01 = "T" And (m_TM10 = "000" Or m_TM10 = "020") And m_CP10 = "101" Then
                     strExc(0) = "select nvl(cp143,0) cp143 from caseprogress where cp09='" & LabCP09.Caption & "' "
                     intJ = 1
                     Set rsAD = ClsLawReadRstMsg(intJ, strExc(0))
                     If intJ = 1 Then
                         If Val("" & rsAD.Fields("cp143")) = 0 Then
                             strDate = "" '查名未齊備，不更新承辦期限
                         End If
                     End If
                     Set rsAD = Nothing
                 End If
            End If
            'end 2019/05/02
            If strDate <> "" Then 'Added by Lydia 2019/06/28 T-173975的出具同意書無法計算承辦期限
                LabCP48.Caption = ChangeWStringToTDateString(strDate)
                strSql = "UPDATE CaseProgress SET CP48 = " & strDate & " " & _
                         "WHERE CP09 = '" & Trim(LabCP09.Caption) & "' "
                cnnConnection.Execute strSql
                frm210136.m_CP48 = LabCP48.Caption
            End If
         End If 'end 2019/06/03
         cnnConnection.CommitTrans
         
         frm210136.m_CP14 = m_CP14
         If Right(LabID, 5) = "-0-00" Then
            strCaseNo = Left(LabID, Len(LabID) - 5)
         Else
            strCaseNo = LabID
         End If
         'Modified by Lydia 2018/12/10 台灣商標爭議案=>台灣商標案
         'Modified by Lydia 2022/07/15 台灣商標案=>商標著作權案件
         frm210136.strSubject = strCaseNo & " 商標著作權案件承辦期限通知"
         frm210136.strContent = "本所案號：" + LabID + vbCrLf + _
                                "案件名稱：" + LabTM05 + vbCrLf + _
                                "案件性質：" + LabCP10 + vbCrLf + _
                                "收文日　：" + LabCP05 + vbCrLf + _
                                "承辦期限：" + LabCP48 + vbCrLf + _
                                "本所期限：" + LabCP06 + vbCrLf + _
                                "法定期限：" + LabCP07 + vbCrLf + _
                                "齊備日　：" + IIf(Trim(txtEP06) = "", ChangeWStringToTDateString(DBDATE(m_EP06)), ChangeWStringToTDateString(DBDATE(txtEP06))) + vbCrLf + _
                                "是否急件：" + txtCP122
         
         MsgBox "存檔完成！", vbExclamation
         Me.Hide
         Exit Sub
      Else
         MsgBox "無異動資料！", vbExclamation
         Exit Sub
      End If
      
   'Add By Sindy 2012/10/24
   '回覆補充資料作業
   ElseIf WorkType = "2" Then
      bolChk = True: strText = ""
      '檢查是否有逐筆回覆補件資料
      For i = 1 To Grid2.Rows - 1
         If Grid2.TextMatrix(i, 0) = "" And Grid2.TextMatrix(i, 1) = "" Then
            bolChk = False
         Else
            If strText <> "" Then strText = strText & ","
            If Grid2.TextMatrix(i, 0) = "V" Then
               strText = strText & "(Y)" & Trim(Grid2.TextMatrix(i, 2))
            End If
            If Grid2.TextMatrix(i, 1) = "V" Then
               strText = strText & "(N)" & Trim(Grid2.TextMatrix(i, 2))
            End If
         End If
      Next i
      If bolChk = False Then
         MsgBox "請逐筆點選補充資料的回覆狀況 !!", vbExclamation
         Exit Sub
      End If
      
      cnnConnection.BeginTrans
      '齊備日=系統日
      'Modified by Lydia 2019/05/02 改成工作日
'      txtEP06 = strSrvDate(2)
'      strSql = "update engineerprogress set ep06=" & strSrvDate(1) & ",ep36=" & strSrvDate(1) & " where ep02='" & Trim(LabCP09.Caption) & "'"
'      cnnConnection.Execute strSql
'      frm210136.m_EP06 = ChangeWStringToTDateString(strSrvDate(1))
'      '更新補充資料內容
'      strSql = "update tmctldate set tcd06=" & strSrvDate(1) & ",tcd08='" & strText & "'" & _
'               "where tcd01='" & Trim(LabCP09.Caption) & "' and tcd04=" & txtTCD04 & " and tcd05=" & txtTCD05
'      cnnConnection.Execute strSql
'      '計算承辦期限
'      'Memo by Lydia 2018/12/10 商申案計算方式與商爭案相同
'      strDate = PUB_TMdebateCountCP48(LabCP06.Caption, txtCP122, strSrvDate(1), LabCP09.Caption, m_CP13)
'      LabCP48.Caption = ChangeWStringToTDateString(strDate)
'      strSql = "UPDATE CaseProgress SET CP48 = " & strDate & " " & _
'               "WHERE CP09 = '" & Trim(LabCP09.Caption) & "' "
      strExc(1) = CompWorkDay(1, strSrvDate(1))
      txtEP06 = TransDate(strExc(1), 1)
      strSql = "update engineerprogress set ep06=" & strExc(1) & ",ep36=" & strExc(1) & " where ep02='" & Trim(LabCP09.Caption) & "'"
      cnnConnection.Execute strSql
      frm210136.m_EP06 = ChangeWStringToTDateString(strExc(1))
      '更新補充資料內容
      strSql = "update tmctldate set tcd06=" & strExc(1) & ",tcd08='" & strText & "'" & _
               "where tcd01='" & Trim(LabCP09.Caption) & "' and tcd04=" & txtTCD04 & " and tcd05=" & txtTCD05
      cnnConnection.Execute strSql
      'Add By Sindy 2021/5/7 加入 齊備日異動時，紀錄
      Pub_SaveLog strUserNum, "齊備日異動ep06：" & strExc(1) & " ", SystemNumber(LabID.Caption, 1), SystemNumber(LabID.Caption, 2), SystemNumber(LabID.Caption, 3), SystemNumber(LabID.Caption, 4), LabCP09.Caption
      '計算承辦期限
      'Memo by Lydia 2018/12/10 商申案計算方式與商爭案相同
      'Modified by Lydia 2019/05/02　商申案計算方式用原本的方式
      'Modified by Lydia 2022/07/15 TC案之文件齊備日管控
      'If m_CP10 <> "102" Then 'Added by Lydia 2019/06/03 排除延展案
      '      If InStr(TMdebate, m_CP10) > 0 Then
      If ((m_TM01 = "T" Or m_TM01 = "FCT") And m_CP10 <> "102") Or m_TM01 = "TC" Then
            'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
            If (m_TM01 = "T" Or m_TM01 = "FCT") And InStr(TMdebate, m_CP10) > 0 And Not (m_TM01 = "FCT" And InStr(FCT_NotTMdebate, m_CP10) > 0) Then
      'end 2022/07/15
                 strDate = PUB_TMdebateCountCP48(LabCP06.Caption, txtCP122, strExc(1), LabCP09.Caption, m_CP13)
            Else
                 strDate = Pub_GetHandleDay(m_TM01, m_TM10, m_CP10, strExc(1), DBDATE(LabCP06.Caption), LabCP09.Caption)
            End If
            'end 2019/05/02
            If strDate <> "" Then 'Added by Lydia 2019/09/18 分割案無法計算承辦期限
                LabCP48.Caption = ChangeWStringToTDateString(strDate)
                strSql = "UPDATE CaseProgress SET CP48 = " & strDate & " " & _
                         "WHERE CP09 = '" & Trim(LabCP09.Caption) & "' "
                cnnConnection.Execute strSql
                frm210136.m_CP48 = LabCP48.Caption
            End If 'end 2019/09/18
      End If 'end 2019/05/06
      
      cnnConnection.CommitTrans
      
      strTemp = Split(strText, ",")
      strText = ""
      For i = 0 To UBound(strTemp)
         If i > 0 Then
            strText = strText + "　　　　　　　"
         End If
         strText = strText + strTemp(i) + vbCrLf
      Next i
      frm210136.m_CP14 = m_CP14
      If Right(LabID, 5) = "-0-00" Then
         strCaseNo = Left(LabID, Len(LabID) - 5)
      Else
         strCaseNo = LabID
      End If
      'Modified by Lydia 2018/12/10 台灣商標爭議案=>台灣商標案
      'Modified by Lydia 2022/07/15 台灣商標案=>商標著作權案件
      frm210136.strSubject = strCaseNo & " 商標著作權案件承辦期限通知 (回覆通知補充資料)"
      frm210136.strContent = "本所案號：" + LabID + vbCrLf + _
                             "案件名稱：" + LabTM05 + vbCrLf + _
                             "案件性質：" + LabCP10 + vbCrLf + _
                             "收文日　：" + LabCP05 + vbCrLf + _
                             "承辦期限：" + LabCP48 + vbCrLf + _
                             "本所期限：" + LabCP06 + vbCrLf + _
                             "法定期限：" + LabCP07 + vbCrLf + _
                             "齊備日　：" + ChangeWStringToTDateString(DBDATE(txtEP06)) + vbCrLf + _
                             "是否急件：" + txtCP122 + vbCrLf + vbCrLf + _
                             "回覆補充資料：" + strText + vbCrLf
      
      MsgBox "存檔完成！", vbExclamation
      Me.Hide
      Exit Sub
   End If
Case 0
   Me.Hide
   Exit Sub
'Add By Sindy 2012/7/16 智權人員輸入齊備日要取消
Case 2 '電腦中心取消齊備
   '檢查是否有齊備日
   If Val(m_EP06) = 0 Then
      MsgBox "無齊備日！", vbExclamation
      Exit Sub
   End If
   
   cnnConnection.BeginTrans
   strSql = "update engineerprogress set ep06=0,ep36=0 where ep02='" & LabCP09 & "'"
   cnnConnection.Execute strSql
   strSql = "insert into tmctldate(tcd01,tcd02,tcd03,tcd04,tcd05,tcd06,tcd07) values('" & LabCP09 & "','4','" & strUserNum & "'," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & ",null,'電腦中心取消齊備')"
   cnnConnection.Execute strSql
   'Modified by Lydia 2022/07/15 TC案之文件齊備日管控
   'If m_CP10 <> "102" Then 'Added by Lydia 2019/06/03 排除延展案
   If ((m_TM01 = "T" Or m_TM01 = "FCT") And m_CP10 <> "102") Or m_TM01 = "TC" Then
        strSql = "update CaseProgress SET CP48=null WHERE CP09='" & LabCP09 & "'"
        cnnConnection.Execute strSql
   End If 'end 2019/06/03
   cnnConnection.CommitTrans
   
   frm210136.m_EP06 = ""
   frm210136.m_CP48 = ""
   MsgBox "已取消！", vbExclamation
   Me.Hide
   Exit Sub
'2012/7/16 End
End Select

ErrHnd:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   cmdState = -1
   'Add By Sindy 2012/7/16
   If Pub_StrUserSt03 = "M51" Then
      cmdok(2).Visible = True
   Else
      cmdok(2).Visible = False
   End If
   '2012/7/16 End
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm210136_1 = Nothing
End Sub

Public Function Process(strText As String) As Boolean
Dim intCnt As Integer, i As Integer
Dim strEP34 As String
Dim strTemp As Variant, intRow As Integer

On Error GoTo ErrHnd
   
   bolNotData = True
   Process = False
   
   'Modified by Lydia 2018/12/10 開放T台灣案管控文件齊備
   'strSql = "select sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,cpm03 as 案件性質,tm05 as 案件名稱," & _
            "s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(ep06) as 齊備日," & _
            "sqldatet(cp48) as 承辦期限,sqldatet(ep28) as 指定會稿日,ep34 as 是否會稿,sqldatet(ep07) as 會稿日,sqldatet(cp27) as 發文日," & _
            "cp16 As 費用, cp18 As 點數, cp64 As 進度備註, cp09 As 總收文號,cp122,cp13,cp14" & _
            " from caseprogress,engineerprogress,trademark,casepropertymap,staff s1,staff s2" & _
            " where cp01 in('T','FCT') and cp10 in (" & TMdebate & ")" & _
            " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)" & _
            " and cp09=ep02(+)" & _
            " and tm10='000'" & _
            " and cp01=cpm01(+) and cp10=cpm02(+)" & _
            " and cp13=s1.st01(+)" & _
            " and cp14=s2.st01(+)" & _
            " and cp09='" & strText & "'"
   'Modified by Lydia 2019/05/02 +tm01,tm10
   'Modified by Lydia 2022/07/15 T大陸案之齊備日管控: tm10='000' => tm10 in ('000','020')、cpm03 => decode(tm10,'000',cpm03,cpm04)
   strSql = "select sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,decode(tm10,'000',cpm03,cpm04) as 案件性質,tm05 as 案件名稱," & _
            "s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(ep06) as 齊備日," & _
            "sqldatet(cp48) as 承辦期限,sqldatet(ep28) as 指定會稿日,ep34 as 是否會稿,sqldatet(ep07) as 會稿日,sqldatet(cp27) as 發文日," & _
            "cp16 As 費用, cp18 As 點數, cp64 As 進度備註, cp09 As 總收文號,cp122,cp13,cp14,cp10,tm01,tm10" & _
            " from caseprogress,engineerprogress,trademark,casepropertymap,staff s1,staff s2" & _
            " where cp01 in('T','FCT') " & _
            " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)" & _
            " and cp09=ep02(+)" & _
            " and tm10 in ('000','020')" & _
            " and cp01=cpm01(+) and cp10=cpm02(+)" & _
            " and cp13=s1.st01(+)" & _
            " and cp14=s2.st01(+)" & _
            " and cp09='" & strText & "'"
   'Added by Lydia 2022/07/15 TC案之文件齊備日管控: 臺灣、大陸
   strSql = strSql & " Union select sqldatet(cp05) as 收文日,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,decode(sp09,'000',cpm03,cpm04) as 案件性質,sp05 as 案件名稱," & _
            "s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(ep06) as 齊備日," & _
            "sqldatet(cp48) as 承辦期限,sqldatet(ep28) as 指定會稿日,ep34 as 是否會稿,sqldatet(ep07) as 會稿日,sqldatet(cp27) as 發文日," & _
            "cp16 As 費用, cp18 As 點數, cp64 As 進度備註, cp09 As 總收文號,cp122,cp13,cp14,cp10,sp01,sp09" & _
            " from caseprogress,engineerprogress,servicepractice,casepropertymap,staff s1,staff s2" & _
            " where cp01 in('TC')" & _
            " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+)" & _
            " and cp09=ep02(+)" & _
            " and sp09 in ('000','020')" & _
            " and cp01=cpm01(+) and cp10=cpm02(+)" & _
            " and cp13=s1.st01(+)" & _
            " and cp14=s2.st01(+)" & _
            " and cp09='" & strText & "'"
   'end 2022/07/15
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Process = True
         LabID.Caption = "" & .Fields("本所案號")
         LabCP09.Caption = "" & .Fields("總收文號")
         LabTM05.Caption = "" & .Fields("案件名稱")
         LabCP10.Caption = "" & .Fields("案件性質")
         m_CP10 = "" & .Fields("cp10") 'Added by Lydia 2018/12/10
         LabCP05.Caption = "" & .Fields("收文日")
         LabCP14.Caption = "" & .Fields("承辦人")
         LabCP48.Caption = "" & .Fields("承辦期限")
         LabCP06.Caption = "" & .Fields("本所期限")
         LabCP07.Caption = "" & .Fields("法定期限")
         txtEP06.Text = ""
         'Added by Lydia 2019/05/02
         m_TM01 = "" & .Fields("tm01")
         m_TM10 = "" & .Fields("tm10")
         If Trim(.Fields("齊備日")) <> "" Then
            'txtEP06.Text = ChangeTDateStringToTString("" & .Fields("齊備日"))
            m_EP06 = ChangeTDateStringToTString("" & .Fields("齊備日"))
         End If
         LabCP16.Caption = "" & .Fields("費用")
         LabCP18.Caption = "" & .Fields("點數")
         txtCP122.Text = "" & .Fields("cp122")
         m_CP122 = "" & .Fields("cp122")
         strEP34 = "" & .Fields("是否會稿")
         m_CP13 = "" & .Fields("cp13")
         m_CP14 = "" & .Fields("cp14")
      Else
         LabID.Caption = ""
         LabCP09.Caption = ""
         LabTM05.Caption = ""
         LabCP10.Caption = ""
         LabCP05.Caption = ""
         LabCP14.Caption = ""
         LabCP48.Caption = ""
         LabCP06.Caption = ""
         LabCP07.Caption = ""
         txtEP06.Text = "": m_EP06 = ""
         LabCP16.Caption = ""
         LabCP18.Caption = ""
         txtCP122.Text = "": m_CP122 = ""
         m_CP10 = "" 'Added by Lydia 2018/12/10
         strEP34 = ""
         m_CP13 = ""
         m_CP14 = ""
         MsgBox "查無資料！", vbExclamation
         Me.Hide
         Exit Function
      End If
      '要會稿且收文費用<8000元，鎖住是否急件欄不可輸入
      txtCP122.Enabled = True
      'Modified by Lydia 2022/07/15 限制T臺灣案
      If m_TM01 = "T" And m_TM10 = "000" And strEP34 = "Y" And Val(LabCP16) < 8000 Then
         txtCP122.Enabled = False
      End If
   End With
   
   txtEP06.Enabled = True
   Label1(12).Visible = True: Label1(8).Visible = False
   '齊備日歷史記錄:以總收文號抓的商標案管制日期異動記錄檔異動類別=1的資料，依異動日期＋異動時間排序顯示
   'Modify By Sindy 2012/10/24 增加TCD02=5通知補充資料
   strSql = "select sqldatet(tcd04) as 操作日期,st02 as 操作人員,sqltime(tcd05) as 操作時間,sqldatet(tcd06) as 輸入日期,tcd07 as 操作動作,tcd08 as 通知補充資料" & _
            " From tmctldate,staff" & _
            " where tcd01='" & strText & "'" & _
            " and tcd02 in('1','5')" & _
            " and tcd03=st01(+)" & _
            " order by tcd04,tcd05 asc"
   CheckOC3
   grd1.Rows = 2
   grd1.Clear
   SetDataListWidth
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Set grd1.Recordset = AdoRecordSet3.Clone
         SetDataListWidth
         intCnt = 0
         For i = 1 To grd1.Rows - 1
            If grd1.TextMatrix(i, 4) = "收文" Then
               intCnt = intCnt + 1
            ElseIf grd1.TextMatrix(i, 4) = "收文取消齊備" Then
               intCnt = intCnt - 1
            ElseIf grd1.TextMatrix(i, 4) = "智權人員維護" Then
               intCnt = intCnt + 1
            End If
         Next i
         If intCnt >= 2 Then
            txtEP06.Enabled = False
            Label1(12).Visible = False: Label1(8).Visible = True
         End If
      End If
   End With
   
   cmdok(1).Enabled = False
   If txtEP06.Enabled = True Or txtCP122.Enabled = True Then
      cmdok(1).Enabled = True
   End If
   
   'Add By Sindy 2012/10/24 回覆補充資料作業
   'If WorkType = "2" Then
      strSql = "select sqldatet(tcd04) as 操作日期,st02 as 操作人員,sqltime(tcd05) as 操作時間,sqldatet(tcd06) as 輸入日期,tcd07 as 操作動作,tcd08 as 通知補充資料,TCD04,TCD05" & _
               " From tmctldate,staff" & _
               " where tcd01='" & strText & "'" & _
               " and tcd02 in('5')" & _
               " and tcd03=st01(+)" & _
               " order by tcd04,tcd05 desc"
      CheckOC3
      Grid2.Rows = 2
      Grid2.Clear
      SetDataListWidth2
      With AdoRecordSet3
         .CursorLocation = adUseClient
         .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If .RecordCount > 0 Then
            .MoveFirst
            If InStr(.Fields(5), "(Y)") = 0 And InStr(.Fields(5), "(N)") = 0 Then '未回覆
               cmdok(1).Enabled = True
               txtTCD04 = .Fields("TCD04")
               txtTCD05 = .Fields("TCD05")
               strTemp = Split(.Fields(5), ",")
               intRow = 0
               For i = 0 To UBound(strTemp)
                  intRow = intRow + 1
                  If intRow > 1 Then
                     Grid2.AddItem ""
                  End If
                  Grid2.TextMatrix(intRow, 2) = Trim(strTemp(i))
               Next i
            Else
               bolNotData = False
               'CmdOK(1).Enabled = False
            End If
         Else
            bolNotData = False
            'CmdOK(1).Enabled = False
         End If
      End With
   'End If
   '2012/10/24 End
   
   Exit Function
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

'Add By Sindy 2012/12/18
Private Sub GRD1_DblClick()
Dim i As Integer
   
   grd1.col = 0
   grd1.row = grd1.MouseRow
   grd1.Visible = False
   If grd1.row <> 0 Then
      If DBDATE(grd1.TextMatrix(grd1.row, 0)) <> "" Then
         If frm090201_b_2.StrMenu(Trim(LabCP09.Caption), DBDATE(grd1.TextMatrix(grd1.row, 0)), Format(grd1.TextMatrix(grd1.row, 2), "HHMMSS")) = True Then
            frm090201_b_2.Show vbModal
         Else
            grd1.Visible = True
            ShowNoData
            Exit Sub
         End If
      End If
   End If
   grd1.Visible = True
End Sub

'Add By Sindy 2012/12/18
Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
Dim i As Integer
   
   getGrdColRow grd1, x, y, nCol, nRow
   grd1.col = nCol
   grd1.row = nRow
End Sub

Private Sub txtCP122_GotFocus()
   TextInverse txtCP122
End Sub

Private Sub txtCP122_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtEP06_GotFocus()
   InverseTextBox txtEP06
End Sub

'Added by Lydia 2019/05/02 修改後無法直接觸發Validate
Private Sub txtEP06_LostFocus()
Dim tmpBol As Boolean
  Call txtEP06_Validate(tmpBol)
End Sub

Private Sub txtEP06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(txtEP06) = False Then
      If CheckIsTaiwanDate(txtEP06, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "齊備日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtEP06_GotFocus
      End If
      'Modified by Lydia 2019/05/02 輸入非工作日，改成往後推工作日; ex.T-217900齊備日輸入5/1(勞動節放假)
      'If ChkWork(ChangeTStringToWString(txtEP06)) = False Then
      '   Cancel = True
      '   txtEP06_GotFocus
      'End If
      strExc(1) = CompWorkDay(1, DBDATE(txtEP06))
      If strExc(1) <> DBDATE(txtEP06) Then
          MsgBox "輸入之日期不是工作天!!" & vbCrLf & "請輸入" & TransDate(strExc(1), 1), , "日期錯誤"
          txtEP06.SetFocus
          txtEP06_GotFocus
          Cancel = True
      End If
      'end 2019/05/02
   End If
End Sub

Private Sub txtEP06_KeyPress(KeyAscii As Integer)
   'KeyAscii = Pub_NumAscii(KeyAscii, True)    'Mark by Lydia 2019/04/23 桂英反應齊備日無法用Ctrl+C/V
End Sub

'Add By Sindy 2012/10/24
Private Sub Grid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow Grid2, x, y, nCol, nRow
   Grid2.col = nCol
   Grid2.row = nRow
End Sub

'Add By Sindy 2012/10/24
Private Sub Grid2_click()
Dim tmpMouseRow, tmpMouseCol
   
   Grid2.Visible = False
   tmpMouseRow = Grid2.row
   tmpMouseCol = Grid2.col
   Grid2.Visible = True
   If tmpMouseRow <> 0 And (tmpMouseCol <= 1) Then
      If tmpMouseCol = 0 Then
         Grid2.TextMatrix(tmpMouseRow, 0) = "V"
         Grid2.row = tmpMouseRow
         Grid2.col = 0
         Grid2.CellBackColor = &HFFC0C0
         Grid2.TextMatrix(tmpMouseRow, 1) = ""
         Grid2.row = tmpMouseRow
         Grid2.col = 1
         Grid2.CellBackColor = QBColor(15)
      ElseIf tmpMouseCol = 1 Then
         Grid2.TextMatrix(tmpMouseRow, 0) = ""
         Grid2.row = tmpMouseRow
         Grid2.col = 0
         Grid2.CellBackColor = QBColor(15)
         Grid2.TextMatrix(tmpMouseRow, 1) = "V"
         Grid2.row = tmpMouseRow
         Grid2.col = 1
         Grid2.CellBackColor = &HFFC0C0
      End If
   End If
End Sub
