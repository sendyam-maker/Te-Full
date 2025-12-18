VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090903 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專新案未命名區-待命名              "
   ClientHeight    =   4416
   ClientLeft      =   420
   ClientTop       =   4416
   ClientWidth     =   8916
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4416
   ScaleWidth      =   8916
   Begin VB.CommandButton cmdOK 
      Caption         =   "外文本(&P)"
      Height          =   400
      Index           =   4
      Left            =   5400
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   120
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "明細(&E)"
      Height          =   400
      Index           =   3
      Left            =   3720
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   120
      Width           =   795
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3495
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   8655
      _ExtentX        =   15261
      _ExtentY        =   6160
      _Version        =   393216
      Cols            =   10
      AllowUserResizing=   3
      FormatString    =   "V|  收文號  | 收文日 |本所案號   |案件性質|  譯畢期限  |命名人員|翻譯人員|已回報|案件名稱"
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
      _Band(0).Cols   =   10
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   400
      Index           =   2
      Left            =   7110
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   120
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料(&B)"
      Height          =   400
      Index           =   1
      Left            =   6195
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   120
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   7935
      TabIndex        =   2
      Top             =   120
      Width           =   800
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1170
      TabIndex        =   0
      Top             =   180
      Width           =   1800
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3175;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   230
      Width           =   900
   End
End
Attribute VB_Name = "frm090903"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/27 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB、Combo1
'Created by Lydia 2017/11/14 外專新案未命名區-待分案/待確認
Option Explicit
Public cmdState As Integer

Dim colTCT02 As Integer '譯畢期限日期欄位
Dim colTCT01 As Integer '收文別欄位=PK
Dim colCP01 As Integer  '本所案號欄位
Dim colTCT11 As Integer '命名人員已回報
Dim m_GrpMan As String  '各組工程師主管

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdok_Click(Index As Integer)
   If Index = 0 Then '查詢
      If doQuery(True) = False Then
      End If
   Else
      cmdState = Index
      PubShowNextData
   End If
End Sub

Public Sub PubShowNextData()
Dim inX As Integer, inY As Integer
Dim rsRd As New ADODB.Recordset
Dim intR As Integer
Dim Str01 As String
Dim lngColor As Long
Dim stUser As String
Dim hLocalFile As Long 'Added by Lydia 2018/06/21

    stUser = Trim(Mid(Combo1.Text, 1, 6))
    For inX = 1 To MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.row = inX
       MSHFlexGrid1.col = 0
       If Trim(MSHFlexGrid1.Text) = "V" Then
           MSHFlexGrid1.Text = ""
           MSHFlexGrid1.col = 0
           MSHFlexGrid1.CellBackColor = MSHFlexGrid1.BackColor
           MSHFlexGrid1.col = 4
           lngColor = MSHFlexGrid1.CellBackColor
           For inY = 1 To 3
               MSHFlexGrid1.col = inY
               MSHFlexGrid1.CellBackColor = lngColor
           Next inY
           If cmdState < 3 Then
              If fnSaveParentForm(Me) = False Then
                  Me.Enabled = True
                  Exit Sub
              End If
           End If
           '本所案號
           Str01 = Trim(MSHFlexGrid1.TextMatrix(inX, colCP01)) & "-" & Trim(MSHFlexGrid1.TextMatrix(inX, colCP01 + 1)) & "-" & Trim(MSHFlexGrid1.TextMatrix(inX, colCP01 + 2)) & "-" & Trim(MSHFlexGrid1.TextMatrix(inX, colCP01 + 3))
           If Replace(Str01, "-", "") <> "" Then
                Select Case cmdState
                    Case 1 '基本檔
                         frm100101_3.Show
                         frm100101_3.Tag = Pub_RplStr(Str01)
                         frm100101_3.StrMenu
                    Case 2 '進度檔
                         frm100101_2.Show
                         frm100101_2.Tag = Pub_RplStr(Str01)
                         frm100101_2.StrMenu
                    Case 3 '明細
                         Me.Hide
                         Call frm090903_1.SetParent(Me, Str01, Trim(MSHFlexGrid1.TextMatrix(inX, colTCT01)), Trim(Mid(Combo1.Text, 1, 6)))
                         frm090903_1.Show
                    'Added by Lydia 2017/12/27
                    Case 4 '外文本
On Error GoTo ErrHand01 'Added by Lydia 2018/03/23 無權限的錯誤要改訊息
                        'Added by Lydia 2020/01/20 開啟[原始檔區]
                        If InStr(cmdOK(cmdState).Caption, "原始檔") > 0 Then
                            If PUB_CheckFormExist("frm100101_M") Then
                                MsgBox "請先關閉共同查詢〔原始檔區〕畫面！"
                            Else
                                Call ChgCaseNo(Replace(Str01, "-", ""), strExc)
                                If PUB_ChkCPExist(strExc, cntEnglish_Vers, , strExc(0), , "D") = True Then 'English_Vers992
                                    frm100101_M.m_strKey = strExc(0)
                                    frm100101_M.SetParent Me
                                    If frm100101_M.QueryData = True Then
                                       frm100101_M.Show
                                       Me.Hide
                                    End If
                                Else
                                   MsgBox strExc(1) & "-" & strExc(2) & "在〔原始檔區〕的English_Vers收文號不存在!", vbInformation
                                End If
                            End If
                        Else
                        'end 2020/01/20
                            'Modified by Lydia 2018/05/09 +系統別
                            'Modifiede by Lydia 2021/12/06 (109/4/6)已將\\Typing2的"English_Vers"和"專利案件"的案件資料夾，全部搬到原始檔區
                            'strExc(1) = Pub_GetFCPcaseFilePath(Trim(MSHFlexGrid1.TextMatrix(inX, colCP01 + 1)), , Trim(MSHFlexGrid1.TextMatrix(inX, colCP01)))
                            'If Dir(strExc(1) & "\*.*") <> "" Then
                            '     'Modified by Lydia 2018/06/21 用檔案總管開啟放置1~2分鐘後,檔案總管會出錯(ex. A2037, A4041)
                            '     'SHELL "Explorer.exe " & strExc(1), vbNormalFocus  '開啟案件資料夾
                            '     ShellExecute hLocalFile, "explore", strExc(1), vbNullString, vbNullString, 1
                            '     Exit Sub
                            'Else
                            '     MsgBox Str01 & "在" & strExc(1) & "的資料夾不存在或無檔案!", vbInformation
                            'End If
                            strExc(1) = ""
                            'end 2021/12/06
                        End If 'Added by Lydia 2020/01/20
                    'end 2017/12/27
                End Select
           End If
           Exit For
       End If
    Next inX
    Me.Enabled = True
    
'Added by Lydia 2018/03/23
    Exit Sub
    
ErrHand01:
    If Err.Number <> 0 Then
         '全部錯誤訊息統一
         MsgBox "無法讀取" & strExc(1) & "，請通知電腦中心！", vbCritical
         Resume Next
    End If
'end 2018/03/23
End Sub

Private Sub Combo1_Click()
      'Added by Lydia 2017/12/25 直接查詢
      If Combo1.Tag <> "" And Combo1.Tag <> Combo1.Text Then
          If doQuery(True) = False Then
          End If
      End If
      Combo1.Tag = Combo1.Text
      'end 2017/12/25
End Sub

Private Sub Form_Load()
Dim tmpArr As Variant

    MoveFormToCenter Me
   
    Combo1.Clear
    Combo1.AddItem strUserNum & " " & strUserName
    '檢查當時是否需要為他人職代
    Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False)
    Combo1.ListIndex = 0
    
    If doQuery(False) = False Then
    End If
    
    'Added by Lydia 2020/01/20 專利案件和English_Vers檔案：判斷檔案上傳目的地
    If strSrvDate(1) >= XY特殊權限啟用日by檔案 Then
        cmdOK(4).Caption = Replace(cmdOK(4).Caption, "外文本", "原始檔")
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090902 = Nothing
End Sub

Public Function doQuery(ByVal bolMsg As Boolean) As Boolean
Dim strQuery As String
Dim intQ As Integer
Dim rsQuery As New ADODB.Recordset

   strQuery = Trim(Mid(Combo1.Text, 1, 6))
   
   If strQuery <> "" Then
       SetGrd True 'Added by Lydia 2018/01/03
      'Added by Lydia 2025/10/14 調整翻譯人員顯示
      strExc(3) = "DECODE(TCT27,'1','舜禹','2','捷恩凱','3','迅達','4','百靈','5','湃傳思','Z','其他-'||TCT28,'A',S1.ST02||'-下班','B',S1.ST02||'-上班',TCT27)"
      
      '抓未回報
      'Modified by Lydia 2018/03/06 + LPAD
      'Modified by Lydia 2018/09/11 排除閉卷和銷卷
      'Modified by Lydia 2025/10/14 調整翻譯人員顯示
      'strSql = "SELECT '' v, TCT01 AS 收文號,SUBSTR(SQLDATET(CP05),1,9) AS 收文日," & _
               "CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) AS 本所案號,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質," & _
               "DECODE(TCT02,NULL,'',SUBSTR(SQLDATET(TCT02),1,9)||' '||SQLTIME6(TCT03||'00')) AS 譯畢期限,S1.ST02 AS 命名人員," & _
               "DECODE(TCT27,'1','舜禹','2','捷恩凱','3','迅達','4','其他-'||TCT28,'A',S1.ST02||'-下班','B',S1.ST02||'-上班',TCT27) 翻譯人員," & _
               "DECODE(TCT11,NULL,'','Y') AS 已回報,NVL(TCT16,TCT17) 案件名稱" & _
               ",DECODE(TCT02,NULL,'2','1') ORD1,DECODE(TCT10,NULL,'1','2') ORD2,TCT02,LPAD(TCT03,4,'0') TCT03,CP01,CP02,CP03,CP04,TCT04,TCT07,TCT10 " & _
               "FROM TRANSCASETITLE,CASEPROGRESS,CASEPROPERTYMAP,PATENT,STAFF S1 " & _
               "WHERE TCT01=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and pa57 is null and pa108 is null " & _
               "AND CP01=CPM01(+) AND CP10=CPM02(+) AND TCT10=S1.ST01(+) AND TCT10='" & strQuery & "' AND NVL(TCT11,0) = 0 "
      strSql = "SELECT '' v, TCT01 AS 收文號,SUBSTR(SQLDATET(CP05),1,9) AS 收文日," & _
               "CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) AS 本所案號,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質," & _
               "DECODE(TCT02,NULL,'',SUBSTR(SQLDATET(TCT02),1,9)||' '||SQLTIME6(TCT03||'00')) AS 譯畢期限,S1.ST02 AS 命名人員," & _
               strExc(3) & " 翻譯人員," & _
               "DECODE(TCT11,NULL,'','Y') AS 已回報,NVL(TCT16,TCT17) 案件名稱" & _
               ",DECODE(TCT02,NULL,'2','1') ORD1,DECODE(TCT10,NULL,'1','2') ORD2,TCT02,LPAD(TCT03,4,'0') TCT03,CP01,CP02,CP03,CP04,TCT04,TCT07,TCT10 " & _
               "FROM TRANSCASETITLE,CASEPROGRESS,CASEPROPERTYMAP,PATENT,STAFF S1 " & _
               "WHERE TCT01=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and pa57 is null and pa108 is null " & _
               "AND CP01=CPM01(+) AND CP10=CPM02(+) AND TCT10=S1.ST01(+) AND TCT10='" & strQuery & "' AND NVL(TCT11,0) = 0 "
      strSql = strSql & " ORDER BY ORD1,TCT02,TCT03,ORD2 "
      
      If bolMsg = True Then
         intQ = 0
      Else
         intQ = 1
      End If
      Set rsQuery = ClsLawReadRstMsg(intQ, strSql)
      MSHFlexGrid1.FixedCols = 0
      If intQ = 1 Then
         doQuery = True
         Set MSHFlexGrid1.Recordset = rsQuery
         'Modified by Lydia 2018/01/03 Grid點選失效的情況
         'SetGrd (rsQuery.RecordCount + 1)
         SetGrd False
         MSHFlexGrid1.FixedCols = 5
      Else
         doQuery = False
         'Remove by Lydia 2018/01/03 Grid點選失效的情況:曾經無資料列後，又重新載入資料列，所以只能有資料才可以Set 資料來源
         'Set MSHFlexGrid1.Recordset = rsQuery
         'SetGrd
         'end 2018/01/03
      End If
   End If
   
   Set rsQuery = Nothing
End Function

'Modified by Lydia 2018/01/03 改成預設清空
'Private Sub SetGrd(Optional ByVal iR As Integer = 2)
Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   Dim pTime As String
   Dim lngColor As Long
   
   pTime = Mid(Format(ServerTime, "000000"), 1, 4)
   arrGridHeadText = Array("v", "收文號", "收文日", "本所案號", "案件性質", "譯畢期限", "命名人員", "翻譯人員", "已回報", "案件名稱", _
                          "ORD1", "ORD2", "TCT02", "TCT03", "CP01", "CP02", "CP03", "CP04", "TCT04", "TCT07", "TCT10")
   arrGridHeadWidth = Array(200, 970, 840, 1100, 1200, 1260, 860, 860, 600, 1200, _
                           0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
   MSHFlexGrid1.Visible = False
   MSHFlexGrid1.Cols = UBound(arrGridHeadText) + 1
   'Modified by Lydia 2018/01/03
   'MSHFlexGrid1.Rows = iR
   If pReset = True Then
        MSHFlexGrid1.Clear
        MSHFlexGrid1.Rows = 2
   End If
   'end 2018/01/03
   
   For iRow = 0 To MSHFlexGrid1.Cols - 1
      MSHFlexGrid1.row = 0
      MSHFlexGrid1.col = iRow
      MSHFlexGrid1.Text = arrGridHeadText(iRow)
      MSHFlexGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
   Next
   If colCP01 = 0 Then
      colTCT01 = PUB_MGridGetId("收文號", MSHFlexGrid1)
      colCP01 = PUB_MGridGetId("CP01", MSHFlexGrid1)
      colTCT02 = PUB_MGridGetId("TCT02", MSHFlexGrid1)
      colTCT11 = PUB_MGridGetId("已回報", MSHFlexGrid1)
   End If
   
   'Modified by Lydia 2018/01/03
   'For intI = 1 To iR - 1
   For intI = 1 To MSHFlexGrid1.Rows - 1
      MSHFlexGrid1.row = intI
      '有譯畢期限並且系統時間距離期限小於2小時並且命名人員尚未確認回報，則那條記錄顯示為紅色。
      If Trim("" & MSHFlexGrid1.TextMatrix(intI, colTCT11)) = "" And Val("" & MSHFlexGrid1.TextMatrix(intI, colTCT02)) > 0 _
         And Val("" & MSHFlexGrid1.TextMatrix(intI, colTCT02) & MSHFlexGrid1.TextMatrix(intI, colTCT02 + 1)) - Val(strSrvDate(1) & pTime) < 200 Then
         lngColor = &HFF&
      Else
         lngColor = &H80000005
      End If
      For iRow = 0 To MSHFlexGrid1.Cols - 1
         MSHFlexGrid1.col = iRow
         MSHFlexGrid1.CellBackColor = lngColor
         If iRow = colTCT11 Then
            MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
         End If
      Next iRow
   Next intI
   
   MSHFlexGrid1.Visible = True
End Sub

Private Sub MSHFlexGrid1_Click()
Dim intRow As Integer
Dim lngColor As Long
   With MSHFlexGrid1
       If .MouseRow > 0 Then
          intRow = .MouseRow
          .row = intRow
          .col = 4
          lngColor = .CellBackColor
          GridClick MSHFlexGrid1, intRow, 0, 0, 4, "V", lngColor
       End If
   End With
End Sub

Private Sub MSHFlexGrid1_DblClick()
  Call cmdok_Click(3) '明細
End Sub

