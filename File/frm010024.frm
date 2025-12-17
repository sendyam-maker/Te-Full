VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm010024 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文－主管機關"
   ClientHeight    =   5745
   ClientLeft      =   3780
   ClientTop       =   3690
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   Begin VB.CommandButton cmdOK 
      Caption         =   "補輸發文字號(&I)"
      Height          =   345
      Index           =   6
      Left            =   7410
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   390
      Width           =   1470
   End
   Begin VB.TextBox txtCP84 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   264
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1290
      Width           =   525
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   345
      Index           =   5
      Left            =   7230
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   15
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "基本資料(&B)"
      Height          =   345
      Index           =   4
      Left            =   6150
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   15
      Width           =   1080
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消發文(&C)"
      Height          =   345
      Index           =   3
      Left            =   6135
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   390
      Width           =   1260
   End
   Begin VB.Timer Timer1 
      Left            =   8520
      Top             =   1170
   End
   Begin VB.TextBox txtType 
      Height          =   264
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   3
      Top             =   990
      Width           =   525
   End
   Begin VB.TextBox txtCP27 
      Height          =   264
      Left            =   1560
      MaxLength       =   7
      TabIndex        =   0
      Top             =   30
      Width           =   1185
   End
   Begin VB.ComboBox cboListTime 
      Height          =   300
      ItemData        =   "frm010024.frx":0000
      Left            =   1560
      List            =   "frm010024.frx":0002
      TabIndex        =   2
      Top             =   660
      Width           =   1230
   End
   Begin VB.ComboBox cboListType 
      Height          =   300
      ItemData        =   "frm010024.frx":0004
      Left            =   1560
      List            =   "frm010024.frx":000E
      TabIndex        =   1
      Top             =   330
      Width           =   2625
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定發文(&S)"
      Height          =   345
      Index           =   1
      Left            =   4860
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   390
      Width           =   1260
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   4140
      Left            =   30
      TabIndex        =   5
      Top             =   1590
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   7303
      _Version        =   393216
      Cols            =   16
      FixedCols       =   3
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
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
      _Band(0).Cols   =   16
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "重新查詢(&F)"
      Height          =   345
      Index           =   0
      Left            =   4890
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   15
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   2
      Left            =   7980
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   15
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "有無規費："
      Height          =   255
      Index           =   7
      Left            =   180
      TabIndex        =   20
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "(1.有 2.無  空白.全部)"
      Height          =   255
      Index           =   6
      Left            =   2130
      TabIndex        =   19
      Top             =   1320
      Width           =   1845
   End
   Begin VB.Label Label1 
      Caption         =   "共　0　件"
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   5
      Left            =   6870
      TabIndex        =   18
      Top             =   1350
      Width           =   1605
   End
   Begin VB.Label Label1 
      Caption         =   "(1.未發文 2.所有資料)"
      Height          =   255
      Index           =   4
      Left            =   2130
      TabIndex        =   17
      Top             =   1020
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "類　別："
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   16
      Top             =   1020
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "發文日期："
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   15
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "送件時段："
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   14
      Top             =   690
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "部門別："
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   13
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frm010024"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/14 Form2.0已修改 GrdDataList
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
Option Explicit

Dim i As Integer, j As Integer
Dim lngCounterI As Long
Dim m_bolPrintRight As Boolean
Dim m_DBTime As Long '系統時間
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_intRow As Integer, m_intCol As Integer
Dim bolV As Boolean
Public cmdState As Integer
'Added by Lydia 2016/06/03
Dim bolChkDate As Boolean '是否詢問過17:30以後電子發文日期是否為翌日
Dim bolDate2 As Boolean  '確定改為翌日

Private Sub Timer1_Timer()
   m_DBTime = Format(Now, "HHMMSS")
End Sub

Private Sub SetDataListWidth()
Dim ii As Integer
grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "V"
grdDataList.ColWidth(0) = 200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "本所案號"
grdDataList.ColWidth(1) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "案件名稱"
grdDataList.ColWidth(2) = 1200
grdDataList.CellAlignment = flexAlignLeftCenter
grdDataList.col = 3: grdDataList.Text = "案件性質"
grdDataList.ColWidth(3) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
'Add By Sindy 2009/05/07
grdDataList.col = 4: grdDataList.Text = "主管機關"
grdDataList.ColWidth(4) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
'2009/05/07 End
grdDataList.col = 5: grdDataList.Text = "申請案號/審定號/對造號"
grdDataList.ColWidth(5) = 1100
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 6: grdDataList.Text = "有無規費"
grdDataList.ColWidth(6) = 600
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 7: grdDataList.Text = "是否算發文件數"
grdDataList.ColWidth(7) = 600
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 8: grdDataList.Text = "申請人"
grdDataList.ColWidth(8) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 9: grdDataList.Text = "發文部門"
grdDataList.ColWidth(9) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 10: grdDataList.Text = "發文人員"
grdDataList.ColWidth(10) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 11: grdDataList.Text = "專業部發文時間"
grdDataList.ColWidth(11) = 1600
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 12: grdDataList.Text = "發文室發文時間"
grdDataList.ColWidth(12) = 1600
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 13: grdDataList.Text = "發文字號"
grdDataList.ColWidth(13) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 14: grdDataList.Text = "進度備註"
grdDataList.ColWidth(14) = 1200
grdDataList.CellAlignment = flexAlignLeftCenter
grdDataList.col = 15: grdDataList.Text = "總收文號"
grdDataList.ColWidth(15) = 0
grdDataList.CellAlignment = flexAlignLeftCenter
'Added by Morgan 2014/1/6
For ii = 16 To grdDataList.Cols - 1
   grdDataList.ColWidth(16) = 0
Next
End Sub

Private Sub cmdOK_Click(Index As Integer)
cmdState = Index
PubShowNextData
Exit Sub
End Sub

Public Sub PubShowNextData()
Dim BolOk As Boolean, bolSameData As Boolean, bolCancel As Boolean
Dim strData As String, strData2 As String
Dim strCP131 As String, strCP132 As String, strCP28 As String, strCP124 As String
Dim nResponse
Dim strDPNum As String, intCnt As Integer

Select Case cmdState
   Case 0 '重新查詢
      Call SearchData
      
   Case 1 '確定發文
      bolV = False
      BolOk = True
      bolSameData = True
      strData = ""
      For i = 1 To grdDataList.Rows - 1
         If grdDataList.TextMatrix(i, 0) = "V" Then
            bolV = True
            If Trim(grdDataList.TextMatrix(i, 12)) <> "" Then '發文室發文時間
               BolOk = False
            End If
            If strData <> "" And strData <> Trim(grdDataList.TextMatrix(i, 1)) Then
               bolSameData = False
            End If
            strData = Trim(grdDataList.TextMatrix(i, 1)) '本所案號
            If strData = "" Then Exit Sub
         End If
      Next i
      If bolV = False Then
         MsgBox "請勾選欲發文的資料！", vbExclamation + vbOKOnly
         Exit Sub
      End If
      If BolOk = False Then
         MsgBox "勾選的資料中，已有發文室發文時間，請重新確認！", vbExclamation + vbOKOnly
         Exit Sub
      End If
      If bolSameData = False Then
         nResponse = MsgBox("案號不同，確定要發文嗎?", vbYesNo + vbCritical + vbDefaultButton2, "詢問")
         If nResponse = vbNo Then
            Exit Sub
         End If
      End If
      Call GoSend
      
   Case 2 '結束
      'fnCloseAllFrm100
      Unload Me
      Set frm010024 = Nothing
      
   Case 3 '取消發文
      bolV = False
      BolOk = True
      bolSameData = True
      strData = ""
      For i = 1 To grdDataList.Rows - 1
         If grdDataList.TextMatrix(i, 0) = "V" Then
            bolV = True
            If Trim(grdDataList.TextMatrix(i, 12)) = "" Then '發文室發文時間
               BolOk = False
            End If
            If strData <> "" And strData <> Trim(grdDataList.TextMatrix(i, 13)) Then '發文字號
               bolSameData = False
            End If
            'Modified by Morgan 2015/1/6
            'strData = Trim(grdDataList.TextMatrix(i, 13)) '發文字號
            strDPNum = GetValue(i, "cp28")
            strData = Mid(strDPNum, 4)
            'end 2015/1/6
            strData2 = Trim(grdDataList.TextMatrix(i, 9)) '發文部門
            If Trim(grdDataList.TextMatrix(i, 15)) = "" Then Exit Sub '總收文號
         End If
      Next i
      If bolV = False Then
         MsgBox "請勾選欲取消發文的資料！", vbExclamation + vbOKOnly
         Exit Sub
      End If
      If BolOk = False Then
         MsgBox "勾選的資料中，未有發文室發文時間，請重新確認！", vbExclamation + vbOKOnly
         Exit Sub
      End If
      If bolSameData = False Then
         MsgBox "發文字號不同，不可同時取消發文，請重新確認！", vbExclamation + vbOKOnly
         Exit Sub
      End If
      '發文字號
      'Modify By Sindy 2010/8/13 修改百年問題
      'strDPNum = (Left(strSrvDate(1), 4) - 1911) & Format(strData, "000000")
      'Removed by Morgan 2015/1/6
      'strDPNum = CompAutoNumberYear(Left(strSrvDate(1), 4) - 1911) & Format(strData, "000000")
      'end 2015/1/6
      '開啟取消發文視窗
      frm010024_1.txt1(0) = txtCP27    '發文日期
      frm010024_1.txt1(1) = strData2    '發文部門
      frm010024_1.txt1(2) = strDPNum '發文字號
      If frm010024_1.CheckShowList Then
         frm010024_1.Show vbModal
      End If
      strCP131 = frm010024_1.strCP131
      strCP132 = frm010024_1.strCP132
      bolCancel = frm010024_1.BolOk
      Unload frm010024_1
      Set frm010024_1 = Nothing
      If bolCancel Then
         '更新取消發文資料
         Call CancelSend(strDPNum, strCP131, strCP132)
      End If
      
   Case 4 '案件基本資料
      Me.Enabled = False
      For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
        Dim Str01 As String
        grdDataList.col = 0
        grdDataList.Text = ""
        For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            grdDataList.CellBackColor = QBColor(15)
            If j <= 2 Then
               grdDataList.CellBackColor = &H8000000F
            End If
        Next j
        grdDataList.col = 1
        Str01 = SystemNumber(grdDataList, 1)
        If Mid(UCase(Str01), 1, 1) = "N" Then
            Str01 = Mid(Str01, 2, 3)
        End If
        If Not IsNull(grdDataList.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Select Case Pub_RplStr(Str01)
                Case "CFP", "FCP", "P"   '專利
                      Screen.MousePointer = vbHourglass
                      frm100101_3.Show
                      frm100101_3.Tag = Pub_RplStr(grdDataList.Text)
                      frm100101_3.StrMenu
                      Screen.MousePointer = vbDefault
                Case "CFT", "FCT", "T", "TF"   '商標
                      Screen.MousePointer = vbHourglass
                      frm100101_4.Show
                      frm100101_4.Tag = Pub_RplStr(grdDataList.Text)
                      frm100101_4.StrMenu
                      Screen.MousePointer = vbDefault
                Case "CFL", "FCL", "L"          '法務
                      Screen.MousePointer = vbHourglass
                      frm100101_5.Show
                      frm100101_5.Tag = Pub_RplStr(grdDataList.Text)
                      frm100101_5.StrMenu
                      Screen.MousePointer = vbDefault
                Case "LA"            '顧問
                      Screen.MousePointer = vbHourglass
                      frm100101_6.Show
                      frm100101_6.Tag = Pub_RplStr(grdDataList.Text)
                      frm100101_6.StrMenu
                      Screen.MousePointer = vbDefault
                Case Else                  '服務
                     Select Case Pub_RplStr(Str01)
                         Case "TB"    '條碼
                            Screen.MousePointer = vbHourglass
                            frm100101_7.Show
                            frm100101_7.Tag = Pub_RplStr(grdDataList.Text)
                            frm100101_7.StrMenu
                            Screen.MousePointer = vbDefault
                         Case "TM"
                            Screen.MousePointer = vbHourglass
                            frm100101_8.Show
                            frm100101_8.Tag = Pub_RplStr(grdDataList.Text)
                            frm100101_8.StrMenu
                            Screen.MousePointer = vbDefault
                         Case "TD"
                            Screen.MousePointer = vbHourglass
                            frm100101_9.Show
                            frm100101_9.Tag = Pub_RplStr(grdDataList.Text)
                            frm100101_9.StrMenu
                            Screen.MousePointer = vbDefault
                         Case "TC", "CFC"
                            Screen.MousePointer = vbHourglass
                            frm100101_A.Show
                            frm100101_A.Tag = Pub_RplStr(grdDataList.Text)
                            frm100101_A.StrMenu
                            Screen.MousePointer = vbDefault
                         Case Else
                            Screen.MousePointer = vbHourglass
                            frm100101_B.Show
                            frm100101_B.Tag = Pub_RplStr(grdDataList.Text)
                            frm100101_B.StrMenu
                            Screen.MousePointer = vbDefault
                      End Select
            End Select
        End If
        Me.Enabled = True
        Exit Sub
     End If
     Next i
     Me.Enabled = True
     
   Case 5 '案件進度
     Me.Enabled = False
     For i = 1 To grdDataList.Rows - 1
     grdDataList.col = 0
     grdDataList.row = i
     If Trim(grdDataList.Text) = "V" Then
        grdDataList.col = 0
        grdDataList.Text = ""
        For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            grdDataList.CellBackColor = QBColor(15)
            If j <= 2 Then
               grdDataList.CellBackColor = &H8000000F
            End If
        Next j
         grdDataList.col = 1
         If Not IsNull(grdDataList.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm100101_2.Show
            frm100101_2.Tag = Pub_RplStr(grdDataList.Text)
            frm100101_2.StrMenu
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
     End If
     Next i
     Me.Enabled = True
     
   Case 6 '補輸發文字號
      bolV = False
      BolOk = True
      intCnt = 0
      strData = ""
      For i = 1 To grdDataList.Rows - 1
         If grdDataList.TextMatrix(i, 0) = "V" Then
            bolV = True
            intCnt = intCnt + 1
            If Trim(grdDataList.TextMatrix(i, 13)) <> "" Then '發文字號
               BolOk = False
            End If
            strData = Trim(grdDataList.TextMatrix(i, 15)) '總收文號
            If strData = "" Then Exit Sub
         End If
      Next i
      If bolV = False Then
         MsgBox "請勾選欲補輸發文字號的資料！", vbExclamation + vbOKOnly
         Exit Sub
      End If
      If BolOk = False Then
         MsgBox "必須為未發文資料，請重新確認！", vbExclamation + vbOKOnly
         Exit Sub
      End If
      If intCnt > 1 Then
         MsgBox "只可點選一筆資料，請重新確認！", vbExclamation + vbOKOnly
         Exit Sub
      End If
      '開啟補輸發文字號視窗
      frm010024_2.strCP09 = strData '總收文號
      If frm010024_2.CheckShowList Then
         frm010024_2.Show vbModal
      End If
      strCP28 = frm010024_2.strCP28
      strCP124 = frm010024_2.strCP124
      bolCancel = frm010024_2.BolOk
      Unload frm010024_2
      Set frm010024_2 = Nothing
      If bolCancel Then
         '更新發文字號資料
         Call UpdateSendNo(strData, strCP28, strCP124)
      End If
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   '更新系統時間
   m_DBTime = ServerTime
   time = Format(m_DBTime, "##:##:##")
   Timer1.Interval = 1000
   
   SetDataListWidth
   
   '發文日期
   txtCP27.Text = strSrvDate(2)
   
   '清單種類
   cboListType.Clear
   cboListType.AddItem "全部"
   cboListType.AddItem "內專"
   cboListType.AddItem "內商"
   cboListType.AddItem "外專"
   cboListType.AddItem "外商"
   cboListType.ListIndex = 0
   
   '送件時段
   cboListTime.Clear
   cboListTime.AddItem "上午"
   cboListTime.AddItem "下午"
   '12點前預設上午
   If m_DBTime < 120000 Then
      cboListTime.ListIndex = 0
   Else
      cboListTime.ListIndex = 1
   End If
   
   '類別
   txtType.Text = "1"
   '有無規費
   txtCP84.Text = ""
   
   m_bolPrintRight = IsUserHasRightOfFunction("frm010024", strPrint, False)
   
   cmdOK(0).Default = True
   Call SearchData
   cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm010024 = Nothing
End Sub

Private Sub GrdDataList_Click()
Dim strItem As String
   
   grdDataList.Visible = False
   
   '依點選的欄位做排序
   If grdDataList.MouseRow = 0 Then
      If grdDataList.MouseCol <> 0 Then
         m_intRow = grdDataList.MouseRow
         m_intCol = grdDataList.MouseCol
         grdDataList.row = m_intRow
         grdDataList.col = m_intCol
         Select Case m_intCol
            Case 5, 13
               '數字
               If m_blnColOrderAsc = True Then
                   Me.grdDataList.Sort = 3 '昇冪
                   m_blnColOrderAsc = False
               Else
                   Me.grdDataList.Sort = 4 '降冪
                   m_blnColOrderAsc = True
               End If
           Case Else
               '字串
               If m_blnColOrderAsc = True Then
                   Me.grdDataList.Sort = 5 '昇冪
                   m_blnColOrderAsc = False
               Else
                   Me.grdDataList.Sort = 6 '降冪
                   m_blnColOrderAsc = True
               End If
         End Select
      End If
   End If
   
   '勾選
   grdDataList.row = grdDataList.MouseRow
   grdDataList.col = 0
   If grdDataList.row <> 0 Then
      'If grdDataList.TextMatrix(GrdDataList.MouseRow, 13) = "" Then
         If grdDataList.Text = "V" Then
              grdDataList.Text = ""
              For i = 0 To grdDataList.Cols - 1
                  grdDataList.col = i
                  grdDataList.CellBackColor = QBColor(15)
                  If i <= 2 Then
                     grdDataList.CellBackColor = &H8000000F
                  End If
             Next i
         Else
              grdDataList.Text = "V"
              For i = 0 To grdDataList.Cols - 1
                  grdDataList.col = i
                  grdDataList.CellBackColor = &HFFC0C0
                  If i <= 2 Then
                     grdDataList.CellBackColor = &H8000000F
                  End If
              Next i
         End If
         
         'Add by Morgan 2011/4/26
         If grdDataList.TextMatrix(grdDataList.row, 4) <> "經濟部智慧財產局" Then
            grdDataList.col = 4
            grdDataList.CellBackColor = RGB(&HFF, &HA5, &H0)
         End If

      'End If
   End If
   
   grdDataList.Visible = True
   
   '控制按鈕為預設值
   bolV = False
   strItem = ""
   For i = 1 To grdDataList.Rows - 1
      If grdDataList.TextMatrix(i, 0) = "V" Then
         bolV = True
         If Trim(grdDataList.TextMatrix(i, 13)) = "" Then
            strItem = "1" '發文
         Else
            strItem = "3" '取消發文
         End If
         Exit For
      End If
   Next i
   If bolV = False Then
      cmdOK(0).SetFocus
   Else
      If strItem = "1" Then
         cmdOK(1).SetFocus
      ElseIf strItem = "3" Then
         cmdOK(3).SetFocus
      End If
   End If
End Sub

Private Sub GrdDataList_Sort()
   grdDataList.Visible = False
   '依點選的欄位做排序
   If m_intRow = 0 Then
      If m_intCol <> 0 Then
         grdDataList.row = m_intRow
         grdDataList.col = m_intCol
         Select Case m_intCol
            Case 5, 13
               '數字
               If m_blnColOrderAsc = False Then
                   Me.grdDataList.Sort = 3 '昇冪
               Else
                   Me.grdDataList.Sort = 4 '降冪
               End If
           Case Else
               '字串
               If m_blnColOrderAsc = False Then
                   Me.grdDataList.Sort = 5 '昇冪
               Else
                   Me.grdDataList.Sort = 6 '降冪
               End If
         End Select
      End If
   End If
   grdDataList.Visible = True
End Sub

Private Sub SearchData()
Dim strComWhere As String, strPSel As String, strTSel As String, strSSel As String
Dim dblAL05 As Double
Dim strSql As String, strWhere As String

   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
   'Modified by Morgan 2015/1/6 發文號改放收文號(顯示抓後6碼 To_number(substr(CP28,3,6))=>to_number(substr(CP28,4)),+CP28
   'Modify by Morgan 2011/4/27 若有對造號時申請號帶該號數
   'Modified by Morgan 2020/2/14 號數改先抓申請號沒有再抓對造號
   strPSel = "SELECT ' ' AS V,CP01||'-'||CP02||'-'||CP03||'-'||CP04,PA05,CPM03,CP130,NVL(PA11,CP36) PA11,decode(CP84,0,'無',null,'無','有'), " & _
             "CP123,NVL(CU04,CU05),A0902,S1.ST02,SqlTime(CP82), " & _
             "SqlTime (CP125), To_number(substr(CP28,4)), CP64,CP09,CP28 " & _
             "FROM Patent,CaseProgress,CasePropertyMap,Customer,Staff S1,ACC090 " & _
             "WHERE CP01 = PA01 And cp02 = pa02 And cp03 = pa03 And cp04 = pa04 " & _
             "AND substr(PA26,1,8)=CU01(+) " & _
             "AND substr(PA26,9,1)=CU02(+) " & _
             "AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
             "AND CP83=S1.ST01(+) " & _
             "AND S1.ST03=A0901(+) "
                     
   'Modify by Morgan 2011/7/19 商標不必抓進度檔--宋若蘭 T170527
   'modify by sonia 2013/11/12 有審定號時後面同時顯示申請案號
   'modify by sonia 2016/7/13 CFT申請英文證明304之審定號改抓CP30
   'strTSel = "SELECT ' ' AS V,CP01||'-'||CP02||'-'||CP03||'-'||CP04,TM05,CPM03,CP130,NVL(TM15,TM12)||decode(tm15,null,null,decode(tm12,null,null,'/'||tm12)),decode(CP84,0,'無',null,'無','有'), " & _
             "CP123,NVL(CU04,CU05),A0902,S1.ST02,SqlTime(CP82), " & _
             "SqlTime (CP125), To_number(substr(CP28,4)), CP64,CP09,CP28 " & _
             "FROM TradeMark,CaseProgress,CasePropertyMap,Customer,Staff S1,ACC090 " & _
             "WHERE CP01 = tm01 And cp02 = tm02 And cp03 = tm03 And cp04 = tm04 " & _
             "AND substr(TM23,1,8)=CU01(+) " & _
             "AND substr(TM23,9,1)=CU02(+) " & _
             "AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
             "AND CP83=S1.ST01(+) " & _
             "AND S1.ST03=A0901(+) "
   'Modified by Morgan 2020/2/14 號數改先抓申請號沒有再抓對造號
   strTSel = "SELECT ' ' AS V,CP01||'-'||CP02||'-'||CP03||'-'||CP04,TM05,CPM03,CP130,DECODE(CP01||CP10,'CFT304',CP30,NVL(TM15,NVL(TM12,CP36))||decode(tm15,null,null,decode(tm12,null,null,'/'||tm12))),decode(CP84,0,'無',null,'無','有'), " & _
             "CP123,NVL(CU04,CU05),A0902,S1.ST02,SqlTime(CP82), " & _
             "SqlTime (CP125), To_number(substr(CP28,4)), CP64,CP09,CP28 " & _
             "FROM TradeMark,CaseProgress,CasePropertyMap,Customer,Staff S1,ACC090 " & _
             "WHERE CP01 = tm01 And cp02 = tm02 And cp03 = tm03 And cp04 = tm04 " & _
             "AND substr(TM23,1,8)=CU01(+) " & _
             "AND substr(TM23,9,1)=CU02(+) " & _
             "AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
             "AND CP83=S1.ST01(+) " & _
             "AND S1.ST03=A0901(+) "
   'END 2016/7/13
   'Modified by Morgan 2017/10/2 著作權登記會沒有中文名稱 Ex:TC-010888
   'Modified by Morgan 2020/2/14 號數改先抓申請號沒有再抓對造號
   strSSel = "SELECT ' ' AS V,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,SP06) SP05,CPM03,CP130,NVL(SP11,CP36),decode(CP84,0,'無',null,'無','有'), " & _
             "CP123,NVL(CU04,CU05),A0902,S1.ST02,SqlTime(CP82), " & _
             "SqlTime (CP125), To_number(substr(CP28,4)), CP64,CP09,CP28 " & _
             "FROM ServicePractice,CaseProgress,CasePropertyMap,Customer,Staff S1,ACC090 " & _
             "WHERE CP01 = SP01 And cp02 = SP02 And cp03 = SP03 And cp04 = SP04 " & _
             "AND substr(SP08,1,8)=CU01(+) " & _
             "AND substr(SP08,9,1)=CU02(+) " & _
             "AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
             "AND CP83=S1.ST01(+) " & _
             "AND S1.ST03=A0901(+) "
                        
   '組共用Where條件
   '是否經發文室-主管機關不可為NULL
   strComWhere = " AND CP123 is not null "
   '發文日期
   If Len(Trim(txtCP27)) <> 0 Then strComWhere = strComWhere & " AND CP27=" & ChangeTStringToWString(txtCP27) & " "
   '發文室發文日為NULL的, 視為未發文
   If Trim(txtType) = "1" Then strComWhere = strComWhere & " AND CP124 is null "
   '有無規費 1.有 2.無
   If Trim(txtCP84) <> "" Then
      If Trim(txtCP84) = "1" Then
         strComWhere = strComWhere & " AND CP84 > 0 "
      ElseIf Trim(txtCP84) = "2" Then
         strComWhere = strComWhere & " AND (CP84 = 0 OR CP84 is null) "
      End If
   End If
   
   strSql = ""
   If Len(Trim(cboListType)) <> 0 Then
      '0.全部 1.內專
      If cboListType.ListIndex = 0 Or cboListType.ListIndex = 1 Then
         strWhere = " AND S1.ST03 like 'P1%' "
         dblAL05 = GetAppListTime(ChangeTStringToWString(txtCP27), "P1")
         If dblAL05 <> 0 And Len(Trim(cboListTime)) <> 0 Then
            If cboListTime.ListIndex = 0 Then
               strWhere = strWhere & " AND CP82 <= " & dblAL05 & " "
            ElseIf cboListTime.ListIndex = 1 Then
               strWhere = strWhere & " AND CP82 > " & dblAL05 & " "
            End If
         End If
         
         If strSql <> "" Then strSql = strSql & " union all "
         strSql = strSql & strPSel & strComWhere & strWhere
         strSql = strSql & " union all " & strSSel & strComWhere & strWhere
      End If
      
      '0.全部 2.內商
      If cboListType.ListIndex = 0 Or cboListType.ListIndex = 2 Then
         strWhere = " AND S1.ST03 like 'P2%' "
         dblAL05 = GetAppListTime(ChangeTStringToWString(txtCP27), "P2")
         If dblAL05 <> 0 And Len(Trim(cboListTime)) <> 0 Then
            If cboListTime.ListIndex = 0 Then
               strWhere = strWhere & " AND CP82 <= " & dblAL05 & " "
            ElseIf cboListTime.ListIndex = 1 Then
               strWhere = strWhere & " AND CP82 > " & dblAL05 & " "
            End If
         End If
         
         If strSql <> "" Then strSql = strSql & " union all "
         strSql = strSql & strTSel & strComWhere & strWhere
         strSql = strSql & " union all " & strSSel & strComWhere & strWhere
      End If
      
      '0.全部 3.外專
      If cboListType.ListIndex = 0 Or cboListType.ListIndex = 3 Then
         strWhere = " AND S1.ST03 like 'F2%' "
         dblAL05 = GetAppListTime(ChangeTStringToWString(txtCP27), "F2")
         If dblAL05 <> 0 And Len(Trim(cboListTime)) <> 0 Then
            If cboListTime.ListIndex = 0 Then
               strWhere = strWhere & " AND CP82 <= " & dblAL05 & " "
            ElseIf cboListTime.ListIndex = 1 Then
               strWhere = strWhere & " AND CP82 > " & dblAL05 & " "
            End If
         End If
         
         If strSql <> "" Then strSql = strSql & " union all "
         strSql = strSql & strPSel & strComWhere & strWhere
         strSql = strSql & " union all " & strSSel & strComWhere & strWhere
      End If
      
      '0.全部 4.外商
      If cboListType.ListIndex = 0 Or cboListType.ListIndex = 4 Then
         strWhere = " AND S1.ST03 like 'F1%' "
         dblAL05 = GetAppListTime(ChangeTStringToWString(txtCP27), "F1")
         If dblAL05 <> 0 And Len(Trim(cboListTime)) <> 0 Then
            If cboListTime.ListIndex = 0 Then
               strWhere = strWhere & " AND CP82 <= " & dblAL05 & " "
            ElseIf cboListTime.ListIndex = 1 Then
               strWhere = strWhere & " AND CP82 > " & dblAL05 & " "
            End If
         End If
         
         If strSql <> "" Then strSql = strSql & " union all "
         strSql = strSql & strTSel & strComWhere & strWhere
         strSql = strSql & " union all " & strSSel & strComWhere & strWhere
      End If
   End If
   strSql = "SELECT * FROM (" & strSql & ") Order By 12,2 ASC "
   
   Screen.MousePointer = vbHourglass
   grdDataList.Clear
   grdDataList.Rows = 2
   SetDataListWidth
   grdDataList.FixedCols = 0
   
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 Then
       Label1(5).Caption = "共　" & adoRecordset.RecordCount & "　件"
       Set grdDataList.Recordset = adoRecordset
   Else
       Label1(5).Caption = "共　0　件"
       ShowNoData
       grdDataList.Clear
   End If
   SetDataListWidth
   grdDataList.FixedCols = 3
   CheckOC
   Call GrdDataList_Sort
   
   'With Me.GrdDataList
   '   For i = 1 To .Rows - 1
   '      .row = i
   '      .col = 9
   '      If .Text <> "" Then
   '         For j = 0 To .Cols - 1
   '            .row = i
   '            .col = j
   '            .CellBackColor = &HC0FFC0   '&HFF&
   '         Next j
   '      End If
   '   Next i
   'End With
   
   '若只有一筆資料, 則直接設定為點選此筆資料
   With Me.grdDataList
      
      If .Rows = 2 Then
         .row = 1
         .col = 1
         If .Text <> "" Then
           .Visible = False
           .row = 1
           .col = 0
           .Text = "V"
           For i = 0 To .Cols - 1
               .col = i
               .CellBackColor = &HFFC0C0
               If i <= 2 Then
                 grdDataList.CellBackColor = &H8000000F
               End If
           Next i
           .Visible = True
         End If
      End If
      
      'Add by Morgan 2011/4/26
      .Visible = False
      For intI = 1 To .Rows - 1
         .row = intI
         If .TextMatrix(.row, 4) <> "經濟部智慧財產局" Then
            .col = 4
            .CellBackColor = RGB(&HFF, &HA5, &H0)
         End If
      Next
      .Visible = True
   End With
   
   Screen.MousePointer = vbDefault
End Sub

'讀取AppList送件時段
Public Function GetAppListTime(Strindex As String, StrIndex2 As String) As Double
Dim strSql As String
CheckOC3
GetAppListTime = 0
strSql = "SELECT * FROM AppList " & _
                "WHERE AL01='" & Strindex & "' " & _
                     "AND AL02='" & StrIndex2 & "' " & _
                "Order By AL05 ASC "
AdoRecordSet3.CursorLocation = adUseClient
AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
   GetAppListTime = AdoRecordSet3.Fields("AL05")
End If
CheckOC3
End Function

'Added by Morgan 2015/1/6
Private Function GetValue(pRow As Integer, pFieldName As String) As String
   Dim iRow As Integer
   With grdDataList
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         GetValue = .TextMatrix(pRow, iRow)
         Exit For
      End If
   Next
   End With
End Function

Private Sub GoSend()
Dim strAutoNumber As String
Dim strDPNum As String, strTime As String
Dim jj As Integer 'Added by Morgan 2015/1/8
Dim strNewCP124 As String 'Added by Lydia 2016/06/03

   strAutoNumber = ""
   Me.Enabled = False
   For i = 1 To grdDataList.Rows - 1
     grdDataList.col = 0
     grdDataList.row = i
     If Trim(grdDataList.Text) = "V" Then
         grdDataList.col = 0
         grdDataList.Text = ""
         For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            grdDataList.CellBackColor = QBColor(15)
            If j <= 2 Then
                grdDataList.CellBackColor = &H8000000F
            End If
         Next j
         grdDataList.col = 1
         
         'Modified by Morgan 2015/1/6 發文號改放收文號
         ''自動給號
         'If strAutoNumber = "" Then
         '   If ClsPDGetAutoNumber("DP", strAutoNumber, True, True) = False Then
         '      Exit Sub
         '   End If
         '   cnnConnection.BeginTrans
         'End If
         ''Modify By Sindy 2010/8/13 修改百年問題
         ''strDPNum = (Left(strSrvDate(1), 4) - 1911) & strAutoNumber
         'strDPNum = CompAutoNumberYear(Left(strSrvDate(1), 4) - 1911) & strAutoNumber
         If strAutoNumber = "" Then
            strDPNum = GetValue(i, "總收文號")
            'Added by Morgan 2015/1/8 若有多筆時抓要計件的
            If GetValue(i, "是否算發文件數") <> "Y" Then
               For jj = i + 1 To grdDataList.Rows - 1
                  If Trim(grdDataList.TextMatrix(jj, 0)) = "V" Then
                     If GetValue(jj, "是否算發文件數") = "Y" Then
                        strDPNum = GetValue(jj, "總收文號")
                        Exit For
                     End If
                  End If
               Next
            End If
            'end 2015/1/8
            strAutoNumber = Mid(strDPNum, 4)
            cnnConnection.BeginTrans
         End If
         'end 2015/1/6
         
         'Memo by Lydia 2020/09/28 依專業部發文日和系統日的差距 , 發文室發文日改三種類型:
         '1.系統日＝專業發文日：17:30前用系統日期和系統時間，17:30後詢問確定加1工作天和固定時間000001;
         '2.系統日＞專業發文日：用系統日期和系統時間; ex.T-223490 109/9/24
         '3.系統日＜專業發文日：用專業發文日和固定時間000002; ex.FCT-028183 99/8/11
         'end 2020/09/28
         If Val(strSrvDate(1)) < Val(ChangeTStringToWString(txtCP27)) Then
            'Modified by Lydia 2020/09/28 系統日＜專業發文日：用專業發文日和固定時間000002 ; ex.FCT-028183 99/8/11
            'strTime = "000001"
            strTime = "000002"
         Else
            'Modified by Lydia 2016/06/04 改成ServerTime
            'strTime = Format(time, "hhmmss")
            strTime = ServerTime
         End If
         'Added by Lydia 2016/06/03 發文室17:30以後電子發文日期為翌日
         strNewCP124 = ChangeTStringToWString(txtCP27)
         'Modified by Lydia 2020/09/28 系統日＝專業發文日：17:30前用系統日期和系統時間，17:30後詢問確定加1工作天和固定時間000001;
         'If bolChkDate = False And Val(strTime) >= 173000 Then '只詢問一次
         If bolChkDate = False And Val(strSrvDate(1)) = Val(ChangeTStringToWString(txtCP27)) And Val(strTime) >= 173000 Then  '只詢問一次
            bolChkDate = True
            If MsgBox("下午5:30以後，發文日期是否為次日?", vbInformation + vbYesNo) = vbYes Then bolDate2 = True
         End If
         If bolDate2 Then
            strTime = "000001"
            strNewCP124 = CompWorkDay(2, strNewCP124)
         End If
         'end 2016/06/03
         'Added by Lydia 2020/09/28 系統日＞專業發文日：用系統日期和系統時間; ex.T-223490 109/9/24
         If Val(strSrvDate(1)) > Val(ChangeTStringToWString(txtCP27)) Then
               strNewCP124 = strSrvDate(1)
         End If
         'end 2020/09/28
         
         'Modify By Sindy 2010/12/8 CP124為畫面上的發文日
'         strSql = "UPDATE CaseProgress " & _
'                                " SET CP124=" & strSrvDate(1) & "," & _
'                                         "CP125=" & Format(time, "hhmmss") & "," & _
'                                         "CP28='" & strDPNum & "'" & _
'                         " WHERE CP09='" & GrdDataList.TextMatrix(i, 15) & "' "
         'Modified by Lydia 2016/06/03 依詢問的結果更新日期
         'strSql = "UPDATE CaseProgress " & _
                                " SET CP124=" & ChangeTStringToWString(txtCP27) & "," & _
                                         "CP125=" & strTime & "," & _
                                         "CP28='" & strDPNum & "'" & _
                         " WHERE CP09='" & grdDataList.TextMatrix(i, 15) & "' "
         strSql = "UPDATE CaseProgress " & _
                                " SET CP124=" & strNewCP124 & "," & _
                                         "CP125=" & strTime & "," & _
                                         "CP28='" & strDPNum & "'" & _
                         " WHERE CP09='" & grdDataList.TextMatrix(i, 15) & "' "
         cnnConnection.Execute strSql
     End If
   Next i
   If strAutoNumber <> "" Then
      cnnConnection.CommitTrans
      MsgBox "發文字號為 ( " & CDbl(strAutoNumber) & " )!!!", vbExclamation + vbOKOnly, Me.Caption
   End If
   Call SearchData
   Me.Enabled = True
End Sub

Private Sub CancelSend(strDPNum As String, strCP131 As String, strCP132 As String)
   Me.Enabled = False
'   bolV = False
   
   'Modify By Sindy 2009/05/04
   '"CP132=" & DBDATE(strCP132) & " "
   cnnConnection.BeginTrans
   strSql = "UPDATE CaseProgress " & _
                          " SET CP124=null," & _
                                   "CP125=null," & _
                                   "CP28=null," & _
                                   "CP131='" & strCP131 & "', " & _
                                   "CP132=" & strSrvDate(1) & " " & _
                   " WHERE CP28='" & strDPNum & "' "
   cnnConnection.Execute strSql
   
'   For i = 1 To GrdDataList.Rows - 1
'     GrdDataList.col = 0
'     GrdDataList.row = i
'     If Trim(GrdDataList.Text) = "V" Then
'         If bolV = False Then
'            bolV = True
'            cnnConnection.BeginTrans
'         End If
'
'         GrdDataList.col = 0
'         GrdDataList.Text = ""
'         For j = 0 To GrdDataList.Cols - 1
'            GrdDataList.col = j
'            GrdDataList.CellBackColor = QBColor(15)
'            If j <= 2 Then
'                GrdDataList.CellBackColor = &H8000000F
'            End If
'         Next j
'         GrdDataList.col = 1
'
'         strSQL = "UPDATE CaseProgress " & _
'                                " SET CP124=null," & _
'                                         "CP125=null," & _
'                                         "CP28=null " & _
'                         " WHERE CP09='" & GrdDataList.TextMatrix(i, 15) & "' "
'         cnnConnection.Execute strSQL
'     End If
'   Next i
'   If bolV = True Then
      cnnConnection.CommitTrans
'   End If
   Call SearchData
   Me.Enabled = True
End Sub

Private Sub UpdateSendNo(strCP09 As String, strCP28 As String, strCP124 As String)
Dim strTime As String
   
   Me.Enabled = False
   
   cnnConnection.BeginTrans
   'Modify By Sindy 2010/12/8
   If Val(strCP124) > Val(strSrvDate(1)) Then
      strTime = "000001"
   Else
      strTime = Format(time, "hhmmss")
   End If
   strSql = "UPDATE CaseProgress " & _
                          " SET CP28='" & strCP28 & "'," & _
                                   "CP124=" & Val(strCP124) & "," & _
                                   "CP125=" & strTime & _
                   " WHERE CP09='" & strCP09 & "' "
   cnnConnection.Execute strSql
   cnnConnection.CommitTrans
   
   Call SearchData
   Me.Enabled = True
End Sub

Private Sub txtCP27_GotFocus()
    InverseTextBox txtCP27
End Sub

Private Sub txtCP27_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtCP27_Validate(Cancel As Boolean)
   If CheckIsTaiwanDate(txtCP27, False) = False Then
        Cancel = True
        MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
        Call txtCP27_GotFocus
        Exit Sub
    End If
End Sub

Private Sub cboListType_GotFocus()
    InverseTextBox cboListType
End Sub

Private Sub cboListType_KeyPress(KeyAscii As Integer)
'KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboListType_Validate(Cancel As Boolean)
If cboListType.Text <> "" Then
    Dim MyArr As String
    Dim MyArr2 As String
    Dim Myi As Integer
    MyArr = cboListType.Text
    For Myi = 0 To cboListType.ListCount - 1
        MyArr2 = cboListType.List(Myi)
        If MyArr = MyArr2 Then
            cboListType.Text = cboListType.List(Myi)
            Exit Sub
        End If
    Next Myi
    MsgBox "部門別輸入錯誤!!!", vbExclamation + vbOKOnly, Me.Caption
    Call cboListType_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub

Private Sub cboListTime_GotFocus()
    InverseTextBox cboListTime
End Sub

Private Sub cboListTime_KeyPress(KeyAscii As Integer)
'KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboListTime_Validate(Cancel As Boolean)
If cboListTime.Text <> "" Then
    Dim MyArr As String
    Dim MyArr2 As String
    Dim Myi As Integer
    MyArr = cboListTime.Text
    For Myi = 0 To cboListTime.ListCount - 1
        MyArr2 = cboListTime.List(Myi)
        If MyArr = MyArr2 Then
            cboListTime.Text = cboListTime.List(Myi)
            Exit Sub
        End If
    Next Myi
    MsgBox "送件時段輸入錯誤!!!", vbExclamation + vbOKOnly, Me.Caption
    Call cboListTime_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub

Private Sub txtType_GotFocus()
    InverseTextBox txtType
End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtType_Validate(Cancel As Boolean)
If txtType.Text <> "" Then
    Select Case txtType.Text
    Case 1, 2
    Case Else
         MsgBox "類別只能輸入1或2!!!", vbExclamation + vbOKOnly, Me.Caption
         Call txtType_GotFocus
         Cancel = True
         Exit Sub
    End Select
End If
End Sub

Private Sub txtCP84_GotFocus()
    InverseTextBox txtCP84
End Sub

Private Sub txtCP84_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtCP84_Validate(Cancel As Boolean)
If txtCP84.Text <> "" Then
    Select Case txtCP84.Text
    Case 1, 2
    Case Else
         MsgBox "有無規費只能輸入1或2或空白!!!", vbExclamation + vbOKOnly, Me.Caption
         Call txtCP84_GotFocus
         Cancel = True
         Exit Sub
    End Select
End If
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
Dim s As Integer

TxtValidate = False

'發文日期
If Len(Trim(txtCP27.Text)) = 0 Then
   s = MsgBox("發文日期不可空白", , "輸入條件錯誤")
   txtCP27.SetFocus
   Exit Function
End If

'部門別
If Len(Trim(cboListType.Text)) = 0 Then
   s = MsgBox("部門別不可空白", , "輸入條件錯誤")
   cboListType.SetFocus
   Exit Function
End If

'送件時段
If Len(Trim(cboListTime.Text)) = 0 Then
   s = MsgBox("送件時段不可空白", , "輸入條件錯誤")
   cboListTime.SetFocus
   Exit Function
End If

'類別
If Len(Trim(txtType.Text)) = 0 Then
   s = MsgBox("類別不可空白", , "輸入條件錯誤")
   txtType.SetFocus
   Exit Function
End If

If Me.txtCP27.Enabled = True Then
   Cancel = False
   txtCP27_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.cboListType.Enabled = True Then
   Cancel = False
   cboListType_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.cboListTime.Enabled = True Then
   Cancel = False
   cboListTime_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.txtType.Enabled = True Then
   Cancel = False
   txtType_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.txtCP84.Enabled = True Then
   Cancel = False
   txtCP84_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function
