VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm050101_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "分案"
   ClientHeight    =   5748
   ClientLeft      =   -2136
   ClientTop       =   1068
   ClientWidth     =   9324
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9324
   Begin VB.CommandButton cmdOK 
      Caption         =   "未註記(&N)"
      Height          =   400
      Index           =   4
      Left            =   3150
      TabIndex        =   34
      Top             =   20
      Width           =   1200
   End
   Begin VB.CommandButton cmdReceive 
      Caption         =   "收到註記(&R)"
      Height          =   400
      Index           =   0
      Left            =   720
      TabIndex        =   33
      Top             =   20
      Width           =   1200
   End
   Begin VB.CommandButton cmdReceive 
      Caption         =   "取消註記(&U)"
      Height          =   400
      Index           =   1
      Left            =   1935
      TabIndex        =   32
      Top             =   20
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   6000
      MaxLength       =   1
      TabIndex        =   16
      Top             =   1320
      Width           =   300
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   6720
      MaxLength       =   1
      TabIndex        =   17
      Top             =   1320
      Width           =   300
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3852
      Left            =   168
      TabIndex        =   18
      Top             =   1848
      Width           =   9132
      _ExtentX        =   16108
      _ExtentY        =   6795
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      RowSizingMode   =   1
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
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "未分案(&U)"
      Default         =   -1  'True
      Height          =   400
      Index           =   2
      Left            =   5520
      TabIndex        =   20
      Top             =   20
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "所有資料(&L)"
      Height          =   400
      Index           =   3
      Left            =   6492
      TabIndex        =   21
      Top             =   20
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "全部選取(&A)"
      Height          =   400
      Left            =   4395
      TabIndex        =   19
      Top             =   20
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1152
      Left            =   120
      TabIndex        =   24
      Top             =   540
      Width           =   4815
      Begin VB.OptionButton optChoose1 
         Caption         =   "電子收文未分案"
         CausesValidation=   0   'False
         Height          =   300
         Index           =   3
         Left            =   1920
         TabIndex        =   9
         Top             =   780
         Width           =   1692
      End
      Begin VB.Frame fraChoose1 
         BorderStyle     =   0  '沒有框線
         Enabled         =   0   'False
         Height          =   372
         Index           =   1
         Left            =   1320
         TabIndex        =   27
         Top             =   480
         Width           =   3372
         Begin VB.TextBox txtSystem 
            Height          =   264
            Left            =   60
            MaxLength       =   3
            TabIndex        =   4
            Top             =   0
            Width           =   492
         End
         Begin VB.Frame fraElse 
            BorderStyle     =   0  '沒有框線
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   492
            Left            =   660
            TabIndex        =   29
            Top             =   0
            Width           =   2412
            Begin VB.TextBox txtCode 
               Height          =   264
               Index           =   0
               Left            =   0
               MaxLength       =   6
               TabIndex        =   5
               Top             =   0
               Width           =   852
            End
            Begin VB.TextBox txtCode 
               Height          =   264
               Index           =   1
               Left            =   960
               MaxLength       =   1
               TabIndex        =   6
               Top             =   0
               Width           =   372
            End
            Begin VB.TextBox txtCode 
               Height          =   264
               Index           =   2
               Left            =   1440
               MaxLength       =   2
               TabIndex        =   7
               Top             =   0
               Width           =   492
            End
         End
         Begin VB.Frame fraTF 
            BorderStyle     =   0  '沒有框線
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   492
            Left            =   600
            TabIndex        =   28
            Top             =   0
            Visible         =   0   'False
            Width           =   2412
            Begin VB.TextBox txtTFCode 
               Height          =   264
               Index           =   3
               Left            =   1920
               MaxLength       =   2
               TabIndex        =   13
               Top             =   0
               Width           =   492
            End
            Begin VB.TextBox txtTFCode 
               Height          =   264
               Index           =   2
               Left            =   1440
               MaxLength       =   1
               TabIndex        =   12
               Top             =   0
               Width           =   372
            End
            Begin VB.TextBox txtTFCode 
               Height          =   264
               Index           =   1
               Left            =   960
               MaxLength       =   1
               TabIndex        =   11
               Top             =   0
               Width           =   372
            End
            Begin VB.TextBox txtTFCode 
               Height          =   264
               Index           =   0
               Left            =   60
               MaxLength       =   5
               TabIndex        =   10
               Top             =   0
               Width           =   852
            End
         End
      End
      Begin VB.Frame fraChoose1 
         BorderStyle     =   0  '沒有框線
         Height          =   372
         Index           =   0
         Left            =   1320
         TabIndex        =   26
         Top             =   180
         Width           =   2652
         Begin VB.TextBox txtDate 
            Height          =   264
            Index           =   0
            Left            =   48
            TabIndex        =   1
            Top             =   0
            Width           =   972
         End
         Begin VB.TextBox txtDate 
            Height          =   264
            Index           =   1
            Left            =   1380
            TabIndex        =   2
            Top             =   0
            Width           =   972
         End
         Begin VB.Line Line1 
            X1              =   1140
            X2              =   1260
            Y1              =   120
            Y2              =   120
         End
      End
      Begin VB.OptionButton optChoose1 
         Caption         =   "以前未分案"
         CausesValidation=   0   'False
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   780
         Width           =   1212
      End
      Begin VB.OptionButton optChoose1 
         Caption         =   "本所案號："
         CausesValidation=   0   'False
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1212
      End
      Begin VB.OptionButton optChoose1 
         Caption         =   "收文日："
         CausesValidation=   0   'False
         Height          =   276
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   180
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8448
      TabIndex        =   23
      Top             =   20
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   7620
      TabIndex        =   22
      Top             =   20
      Width           =   800
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   5040
      TabIndex        =   25
      Top             =   552
      Width           =   4215
      Begin VB.OptionButton optChoose2 
         Caption         =   "主管機關來函"
         Height          =   252
         Index           =   1
         Left            =   2040
         TabIndex        =   15
         Top             =   180
         Width           =   1455
      End
      Begin VB.OptionButton optChoose2 
         Caption         =   "接洽及內部收文單"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   180
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      Caption         =   "收文所別："
      Height          =   240
      Left            =   5040
      TabIndex        =   31
      Top             =   1365
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   6360
      X2              =   6600
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label2 
      Caption         =   "(1.北  2.中  3.南  4.高)"
      Height          =   240
      Left            =   7200
      TabIndex        =   30
      Top             =   1365
      Width           =   1695
   End
End
Attribute VB_Name = "frm050101_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/4 改成Form2.0 (grdDataList)
'Modified by Morgan 2021/12/8 符號"ˇ"改為"V"(因Ext-B字會變小)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
'intLastRow上一次反白的Row
'blnOKtoShow決定是否要反白
Dim intLastRow As Integer, blnOKtoShow As Boolean
'intOpt1為optChoose1中為True之選擇,intOpt2為optChoose2中為True之選擇
Dim intOpt1 As Integer, intOpt2 As Integer
'intChoose被選擇之收文號總數
Dim intChoose As Integer, strReceiveCode() As String, intNowReceiveCode As Integer
Dim intNone As Integer
' 暫存前次搜尋的方式 1:所有資料 2:未分案 3:未註記 4.電子收文未分案
Dim m_QueryType As Integer

'Add By Cheng 2002/04/23
Dim m_bln_FromThisForm As Boolean
Dim lngX As Long, lngY As Long 'Add by Morgan 2018/9/11
Dim stDefArea1 As String, stDefArea2 As String 'Add by Amy 2022/11/16

Public Sub ReChoose(ByRef intNowReceive As Integer, ByRef strReceive() As String)
intNowReceiveCode = intNowReceive
strReceiveCode() = strReceive()
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Dim i As Integer
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nRow As Integer
   Dim bFind As Boolean

   Select Case Index
      Case 0:
         ' 90.07.05 modify by louis
         If grdDataList.Rows < 2 Then
            strTit = "檢核資料"
            strMsg = "請先選取資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
         bFind = False
         For nRow = 1 To grdDataList.Rows - 1
            If grdDataList.TextMatrix(nRow, 0) = "V" Then
               bFind = True
               Exit For
            End If
         Next nRow
         If bFind = False Then
            strTit = "檢核資料"
            strMsg = "請先選取資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
         
         'Added by Morgan 2021/12/9
         '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
         If PUB_CheckFormExist("frm050101_2") = False Then
            Set frm050101_2 = Nothing
         End If
         'end 2021/12/9
         
         frm050101_2.Show
         'Add by Morgan 2004/4/20
         '若為主管機關來函時，轉本所案號不可輸入
         If optChoose2(1).Value = True And IsObject(frm050101_2) Then
            frm050101_2.txtCode(0).Enabled = False
            frm050101_2.txtCode(1).Enabled = False
            frm050101_2.txtCode(2).Enabled = False
            frm050101_2.txtCode(3).Enabled = False
         End If
         Me.Hide
      Case 1:
         Unload Me
      'Modify by Morgan 2008/2/12 +4 未註記
      Case 2, 3, 4: '未分案, 所有資料
         intChoose = 0
         Select Case Index
            Case 3 '所有資料
               m_QueryType = 1
            Case 2 '未分案
               intNone = 1
               m_QueryType = 2
            'Add by Morgan 2008/2/12
            Case 4 '未註記
               intNone = 3
               m_QueryType = 3
         End Select
         'Add By Sindy 2023/6/8
         If optChoose1(3).Value Then
            intNone = 1 '未分案
            m_QueryType = 4 '電子收文未分案
         End If
         '2023/6/8 END
         
         '以Index-2區分是未分案或全部
         ShowList (Index - 2)
         'Add By Cheng 2002/04/23
         If Index = 2 Then
            m_bln_FromThisForm = True
         Else
            m_bln_FromThisForm = False
         End If
         'Modify by Morgan 2003/12/23
         'If m_bln_FromThisForm = True And Me.grdDataList.Rows = 2 Then
         If m_bln_FromThisForm = True And Me.grdDataList.Rows = 2 And Me.Visible = True Then
         'Modify end 2003/12/23
         
            cmdSelectAll_Click
            cmdOK_Click 0
         End If
         m_bln_FromThisForm = False
   End Select
EXITSUB:
End Sub

Private Sub ShowList(ByRef intIndex As Integer)
Select Case intOpt1
             Case 0 '收文日
                       GetSeparateData intOpt1, intOpt2, intIndex, TransDate(txtDate(0), 2), TransDate(txtDate(1), 2)
             Case 1 '本所案號
                       'TF為馬德里案，另外判斷
                       If txtSystem = 馬德里案 Then
                          GetSeparateData intOpt1, intOpt2, intIndex, txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
                             IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3))
                       Else
                           GetSeparateData intOpt1, intOpt2, intIndex, txtSystem, txtCode(0), _
                              IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))
                       End If
            Case 2, 3 '以前未分案 Add By Sindy 2023/6/8 + 3 :電子收文未分案
                      GetSeparateData intOpt1, intOpt2, intIndex
End Select
End Sub

Private Sub cmdReceive_Click(Index As Integer)

   Dim i As Integer
   
   With grdDataList
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 0) = "V" Then
            strExc(1) = ""
            'Modified by Morgan 2018/9/11 CFP電子化-接洽單案件性質數量
            intI = Val(GetValue(i, "數量"))
            If intI > 0 Then
               If Index = 0 Then
                  strExc(1) = ",CP156=" & intI
               Else
                  strExc(1) = ",CP156=Null"
               End If
            End If
            'end 2018/9/11
            'Modify by Morgan 2018/10/12  加數量欄後移一欄
            'Modify by Amy 2022/10/21 原:.TextMatrix(i, 3)
            strSql = "Update Caseprogress Set CP86=" & IIf(Index = 0, "'Y'", "NULL") & strExc(1) & " Where CP09='" & .TextMatrix(i, GetValue(0, "總收文號")) & "'"
            cnnConnection.Execute strSql
            'Modify by Amy 2022/10/21 原:.TextMatrix(i, 4)
            .TextMatrix(i, GetValue(0, "註記")) = IIf(Index = 0, "Y", "")
         End If
      Next
   End With
   
End Sub

Private Sub cmdSelectAll_Click()
Dim i As Integer

For i = 1 To grdDataList.Rows - 1
       grdDataList.TextMatrix(i, 0) = "V"
Next
intChoose = grdDataList.Rows - 1
cmdOK(0).Enabled = True
End Sub

Private Sub Form_Activate()
If txtDate(0).Enabled Then
   txtDate(0).SetFocus
   txtDate_GotFocus 0
End If
'ShowList 0
End Sub

Private Sub GetSeparateData(ByRef strKind1 As Integer, ByRef strKind2 As Integer, ByRef strKind3 As Integer, Optional strSeparate1 As String, Optional strSeparate2 As String, Optional strSeparate3 As String, Optional strSeparate4 As String)
Dim varSaveCursor, i As Integer, j As Integer, k As Integer, intCounter As Integer
Dim strTit As String
Dim strMsg As String
Dim nResponse

Screen.MousePointer = vbHourglass
'93.6.27 modify by sonia
'Set grdDataList.Recordset = objPublicData.ReadSeparateCaseRst(intPCaseKind, intPWhere, strGroup, strKind1, strKind2, strKind3, strSeparate1, strSeparate2, strSeparate3, strSeparate4, intNone)
Set grdDataList.Recordset = ReadSeparateCaseRst(intPCaseKind, intPWhere, strGroup, strKind1, strKind2, strKind3, strSeparate1, strSeparate2, strSeparate3, strSeparate4, intNone)
grdDataList.ColAlignment(3) = flexAlignCenterCenter  'Add by Morgan 2004/9/13
'93.6.27 end
intNone = 0
SetDataListVision grdDataList, True, True
grdDataList.Visible = False
For i = 1 To grdDataList.Rows - 1
   'Modify By Cheng 2002/07/16
'   If ChangeWDateStringToWString(grdDataList.TextMatrix(i, 12)) <> "" And ChangeWDateStringToWString(grdDataList.TextMatrix(i, 12)) <= GetTodayDate And grdDataList.TextMatrix(i, 15) = "" Then
'   If ChangeWDateStringToWString(grdDataList.TextMatrix(i, 12)) <> "" And ChangeWDateStringToWString(grdDataList.TextMatrix(i, 12)) <= GetTodayDate And grdDataList.TextMatrix(i, 16) = "" Then
    '若有本所期限
   'Modify by Morgan 2004/9/13  加註記欄後移一欄
'   If grdDataList.TextMatrix(i, 12) <> "" Then
'        '若本所期限小於等於系統日或本所期限為假日且無發文日
'        If (ChangeTStringToWString(ChangeTDateStringToTString(grdDataList.TextMatrix(i, 12))) <= strSrvDate(1) Or WorkDayCheck(ChangeTStringToWString(ChangeTDateStringToTString(grdDataList.TextMatrix(i, 12)))) = True) And grdDataList.TextMatrix(i, 16) = "" Then
   'Modify by Morgan 2018/10/12  加數量欄後移一欄
   'Modify by Amy 2022/10/21 原:.TextMatrix(i, 14)/.TextMatrix(i, 18)
   'Modif by Amy 2022/11/09 +CP122=Y 顯示紅色
   If grdDataList.TextMatrix(i, GetValue(0, "本所期限")) <> "" Or grdDataList.TextMatrix(i, GetValue(0, "CP122")) = "Y" Then
        '若本所期限小於等於系統日或本所期限為假日且無發文日
        If ((ChangeTStringToWString(ChangeTDateStringToTString(grdDataList.TextMatrix(i, GetValue(0, "本所期限")))) <= strSrvDate(1) Or WorkDayCheck(ChangeTStringToWString(ChangeTDateStringToTString(grdDataList.TextMatrix(i, GetValue(0, "本所期限"))))) = True) And grdDataList.TextMatrix(i, GetValue(0, "發文日")) = "") _
          Or grdDataList.TextMatrix(i, GetValue(0, "CP122")) = "Y" Then
              'Modify By Cheng 2002/07/16
        '      For intCounter = 0 To 14
        
              'Modify by Morgan 2004/9/13
              'For intCounter = 0 To 16
              For intCounter = 0 To grdDataList.Cols - 1
                 grdDataList.row = i
                 grdDataList.col = intCounter
                 grdDataList.CellBackColor = &H8080FF '紅色
              Next
        End If
   Else
      'Modify By Cheng 2002/07/16
'      If grdDataList.TextMatrix(i, 13) = MsgText(9030) Then

      'Modify by Morgan 2004/9/13  加註記欄後移一欄
      'If grdDataList.TextMatrix(i, 14) = MsgText(9030) Then
      'Modify by Morgan 2018/10/12  加數量欄後移一欄
      'Modify by Amy 2022/10/21 原:.TextMatrix(i, 16)
      If grdDataList.TextMatrix(i, GetValue(0, "是否閉卷")) = MsgText(9030) Then
         'Modify By Cheng 2002/07/16
'         For intCounter = 0 To 14

         'Modify by Morgan 2004/9/13
         'For intCounter = 0 To 16
         For intCounter = 0 To grdDataList.Cols - 1
            grdDataList.row = i
            grdDataList.col = intCounter
            grdDataList.CellBackColor = &HFFFF&
         Next
      Else
         'Modify By Cheng 2002/0/16
'         If ChangeWDateStringToWString(grdDataList.TextMatrix(i, 14)) <> "" And ChangeWDateStringToWString(grdDataList.TextMatrix(i, 14)) <> "//" Then
         
         'Modify by Morgan 2004/9/13  加註記欄後移一欄
         'If ChangeWDateStringToWString(grdDataList.TextMatrix(i, 15)) <> "" And ChangeWDateStringToWString(grdDataList.TextMatrix(i, 15)) <> "//" Then
         'Modify by Morgan 2018/10/12  加數量欄後移一欄
         'Modify by Amy 2022/10/21 原:.TextMatrix(i, 17)
         If ChangeWDateStringToWString(grdDataList.TextMatrix(i, GetValue(0, "取消收文日"))) <> "" And ChangeWDateStringToWString(grdDataList.TextMatrix(i, GetValue(0, "取消收文日"))) <> "//" Then
            'Modify By Cheng 2002/07/16
'            For intCounter = 0 To 14

            'Modify by Morgan 2004/9/13
            'For intCounter = 0 To 16
            For intCounter = 0 To grdDataList.Cols - 1
               grdDataList.row = i
               grdDataList.col = intCounter
               grdDataList.CellBackColor = &HE0E0E0
            Next
         End If
     End If
   End If
   'Add by Amy 2022/12/23 需補件者,案件性質顯示 粉紅色
   If grdDataList.TextMatrix(i, GetValue(0, "目前表單狀態")) = "程序補件" Then
        grdDataList.row = i
        grdDataList.col = GetValue(0, "案件性質")
        grdDataList.CellBackColor = &HFF80FF     '粉紅色
   End If
Next
grdDataList.Visible = True
intLastRow = 0
If grdDataList.Rows > 1 Then
'   ShowBar grdDataList, intLastRow, 14
   If intNowReceiveCode < intChoose Then
      j = intNowReceiveCode
      For i = 1 To grdDataList.Rows - 1
             'Modify by Morgan 2018/10/12  加數量欄後移一欄
             'Modify by Amy 2022/10/21 原:.TextMatrix(i, 3)
             If grdDataList.TextMatrix(i, GetValue(0, "總收文號")) = strReceiveCode(j) Then
                grdDataList.TextMatrix(i, 0) = "V"
                k = k + 1
                j = j + 1
                If j = intChoose Then Exit For
             End If
      Next
   End If
   intChoose = k
   cmdSelectAll.Enabled = True
Else
   intChoose = 0
   cmdSelectAll.Enabled = False
End If
If intChoose = 0 Then
   cmdOK(0).Enabled = False
Else
   cmdOK(0).Enabled = True
End If
Screen.MousePointer = vbDefault

' 90.07.02 modify by louis ' 沒有資料顯示訊息
If grdDataList.Rows < 2 Then
   strTit = "查詢資料"
   strMsg = "沒有符合條件的資料!"
   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
End If

End Sub
Private Sub Form_Load()
   MoveFormToCenter Me
   intOpt1 = 0
   intOpt2 = 0
   SetDataListWidth
   'If intPWhere <> 國外_CF Then
      txtDate(0) = GetTaiwanTodayDate
      txtDate(1) = GetTaiwanTodayDate
   'Else
   '   txtDate(0).MaxLength = 8
   '   txtDate(1).MaxLength = 8
   '   txtDate(0) = GetTodayDate
   '   txtDate(1) = GetTodayDate
   'End If
   intChoose = 0
   '93.6.27 add by sonia
   Text1 = PUB_GetST06(strUserNum)
   'Modify by Amy 2022/11/02 接洽單電子收文上線後無紙本,北所進入預設顯示全部
   If Text1 = "1" And strSrvDate(1) >= 接洽單電子收文啟用日 Then
        Text2 = "4"
   Else
        Text2 = PUB_GetST06(strUserNum)
   End If
   '93.6.27 END
   'add by sonia 2020/8/17 專利國內部人員進入時改預設值(中所管理部及外專不改)
   If Left(Pub_StrUserSt03, 2) = "P1" Then
      txtDate(0) = TransDate(CompWorkDay(-2, strSrvDate(1), 1), 1)
      Text1 = "1"
      Text2 = "4"
   End If
   'end 2020/8/17
   'Add by Amy 2022/11/15 記錄預設收文所別
   stDefArea1 = Text1
   stDefArea2 = Text2
End Sub
Private Sub SetDataListWidth()
Dim varGridWidth() As Variant
'Modify by Morgan 2004/9/13 加註記
'varGridWidth = Array(300, 1000, 1000, 3000, 1000, 400, 650, 0, 200, 250, 1000, 1000)
'Modified by Morgan 2018/10/12 加數量
'Modify by Amy 2022/11/09 +CP122 並整理
'varGridWidth = Array(300, 500, 900, 1000, 500, 1800, 1000, 400, 650, 0, 200, 250, 800, 800)

'                    V    數量 收文日期 總收文號 註記 案件名稱
'                    s04 案件性質 s05 本所案號 s06 本所案號 s07 本所案號 s08 本所案號 s09 本所案號
'                    智權人員 目前表單狀態 承辦人 本所期限 是否算案件數 是否閉卷 取消收文日 發文日 CP122
varGridWidth = Array(300, 450, 900, 1000, 400, 1800, 1000, 400, 650, 0, 200, 250, 800, 800, 800, 800, 1200, 800, 1000, 800, 0)
SetGridDataListWidth grdDataList, varGridWidth()
blnOKtoShow = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
   intChoose = 0
   Set frm050101_1 = Nothing
End Sub
Private Sub grdDataList_GotFocus()
GridGotFocus grdDataList
End Sub
Private Sub grdDataList_LostFocus()
GridLostFocus grdDataList
End Sub
Private Sub grdDataList_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then GrdDataList_Click
End Sub
Private Sub GrdDataList_Click()
Dim strHeight As String

If grdDataList.TextMatrix(grdDataList.row, 0) = "V" Then
   grdDataList.TextMatrix(grdDataList.row, 0) = ""
   intChoose = intChoose - 1
Else
   
   'Added by Morgan 2018/9/11 CFP電子化-接洽單案件性質數量
   If Pub_StrUserSt03 = "P12" And strSrvDate(1) >= CFP第一階段電子化啟用日 Then
      With grdDataList
      If GetValue(.row, "註記") = "" And GetValue(.row, "數量") <> "-" Then
         frm040101_3.Label3.Caption = GetValue(.row, "本所案號") & " (" & GetValue(.row, "總收文號") & ")"
         strHeight = mdiMain.Top + Me.Top + .Top + lngY + (mdiMain.Height - mdiMain.ScaleHeight) + (Me.Height - Me.ScaleHeight)
         If Val(strHeight) + frm040101_3.Height > Val(mdiMain.Top + mdiMain.Height) Then
             strHeight = Val(strHeight) - frm040101_3.Height - Val(.RowHeight(1))
         End If
         frm040101_3.Move mdiMain.Left + Me.Left + .Left + lngX, Val(strHeight)
         frm040101_3.Show vbModal
         If Val(strPublicTemp) > 0 Then
            .TextMatrix(.row, GetValue(0, "數量")) = Val(strPublicTemp)
            strPublicTemp = ""
         Else
            strPublicTemp = ""
            Exit Sub
         End If
      End If
      End With
   End If
   'end 2018/9/11
   
   grdDataList.TextMatrix(grdDataList.row, 0) = "V"
   intChoose = intChoose + 1
End If
If intChoose = 0 Then
   cmdOK(0).Enabled = False
Else
   cmdOK(0).Enabled = True
   cmdOK(0).SetFocus
End If
End Sub

Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lngX = x
    lngY = y
End Sub

Private Sub grdDataList_RowColChange()
If intLastRow <> grdDataList.row Then
   If blnOKtoShow Then
      blnOKtoShow = False
'      ShowBar grdDataList, intLastRow, 11
      blnOKtoShow = True
   End If
End If
End Sub
Private Sub optChoose1_Click(Index As Integer)
'Add by Amy 2022/11/16 不是選 以前未分案,改回預設所別
If Index <> 2 Then
    Text1 = stDefArea1
    Text2 = stDefArea2
End If
'end 2022/11/16
intOpt1 = Index
Select Case Index
    Case 0 '收文日
        fraChoose1(0).Enabled = True
        fraChoose1(1).Enabled = False
        cmdOK(3).Enabled = True
        txtDate(0).SetFocus
Case 1 '本所案號
        fraChoose1(0).Enabled = False
        fraChoose1(1).Enabled = True
        cmdOK(3).Enabled = True
        txtSystem.SetFocus
Case 2, 3 '以前未分案
        fraChoose1(0).Enabled = False
        fraChoose1(1).Enabled = False
        cmdOK(3).Enabled = False
        cmdOK(2).SetFocus
        'Add by Amy 2022/11/16 避免資料量太多,造成溢位,選 以前未分案 帶User所別
        Text1 = PUB_GetST06(strUserNum)
        Text2 = Text1
End Select
End Sub
Private Sub optChoose2_Click(Index As Integer)
intOpt2 = Index
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtDate_GotFocus(Index As Integer)
txtDate(Index).SelStart = 0
txtDate(Index).SelLength = Len(txtDate(Index))
CloseIme
End Sub
Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
If CheckKeyIn(Index) = False Then
   Cancel = True
   txtDate_GotFocus Index
End If
End Sub
'93.6.27 ADD BY SONIA
' 收文所別(起)
Private Sub Text1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(Text1) = True Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "收文所別(起)不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Else
      If Text1 < "1" Or Text1 > "4" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "收文所別(起)只可為 '1'~'4' "
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text1_GotFocus
      End If
   End If
End Sub
' 收文所別(止)
Private Sub Text2_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(Text2) = True Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "收文所別(止)不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Else
      If Text2 < "1" Or Text2 > "4" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "收文所別(止)只可為 '1'~'4' "
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text2_GotFocus
      Else
         If Text2 < Text1 Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "收文所別範圍錯誤 "
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Text1_GotFocus
         End If
      End If
   End If
End Sub
'93.6.27 END
Private Function CheckKeyIn(intIndex As Integer) As Boolean
Select Case intIndex
             Case 0, 1
                        If txtDate(intIndex) = "" Then
                           CheckKeyIn = True
                        Else
                           If CheckIsTaiwanDate(txtDate(intIndex)) Then
                              CheckKeyIn = True
                           End If
                        End If
                        If CheckKeyIn = False Then Exit Function
                        If txtDate(0) <> "" And txtDate(1) = "" And intIndex = 1 Then
                           ShowMsg MsgText(9169)
                           CheckKeyIn = False
                        ElseIf txtDate(1) <> "" And Val(txtDate(0)) > Val(txtDate(1)) And intIndex = 1 Then
                           ShowMsg MsgText(9170)
                           CheckKeyIn = False
                        End If
End Select
End Function
Private Sub txtSystem_Change()
If txtSystem.Text = 馬德里案 Then
   fraTF.Visible = True
   fraElse.Visible = False
Else
   fraTF.Visible = False
   fraElse.Visible = True
End If
End Sub
Private Sub txtSystem_GotFocus()
txtSystem.SelStart = 0
txtSystem.SelLength = Len(txtSystem.Text)
CloseIme
End Sub
Private Sub txtSystem_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txtSystem_Validate(Cancel As Boolean)
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.GetGroupCase(txtSystem, strGroup) = False Then
If ClsPDGetGroupCase(txtSystem, strGroup) = False Then
   ShowMsg MsgText(9171)
   Cancel = True
   txtSystem_GotFocus
End If
End Sub
Private Sub txtTFCode_GotFocus(Index As Integer)
txtTFCode(Index).SelStart = 0
txtTFCode(Index).SelLength = Len(txtTFCode(Index).Text)
End Sub
Private Sub txtTFCode_Validate(Index As Integer, Cancel As Boolean)
If Len(txtTFCode(Index)) > 0 And Len(txtTFCode(Index)) < txtTFCode(Index).MaxLength Then
   ShowMsg MsgText(9172)
   Cancel = True
   txtTFCode_GotFocus Index
ElseIf Index = 3 Then
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
       IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3))) = False Then
   If ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
       IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3))) = False Then
   End If
End If
End Sub
Private Sub txtCode_GotFocus(Index As Integer)
txtCode(Index).SelStart = 0
txtCode(Index).SelLength = Len(txtCode(Index).Text)
CloseIme
End Sub
Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
If Len(txtCode(Index)) > 0 And Len(txtCode(Index)) < txtCode(Index).MaxLength Then
   ShowMsg MsgText(9172)
   Cancel = True
   txtCode_GotFocus Index
ElseIf Index = 2 Then
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtCode(0), _
       IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) = False Then
   If ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
       IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) = False Then
   End If
End If
End Sub

' 90.07.06 modify by louis (重新以原條件搜尋)
Public Sub RefreshData()
   Select Case m_QueryType
      Case 1: '前次搜尋方式為"所有資料"
         cmdOK_Click 3
      Case 2: '前次搜尋方式為"未分案"
        'Modify By Cheng 2002/10/30
'         cmdOK_Click 3
         cmdOK_Click 2
      'Add by Morgan 2008/2/12
      Case 3: '前次搜尋方式為"未註記"
         cmdOK_Click 4
   End Select
End Sub

'Add By Cheng 2003/07/22
'檢查本所期限是否為假日期限
Private Function WorkDayCheck(strDate As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

WorkDayCheck = False
If strDate = "" Then Exit Function
StrSQLa = "Select * From Workday Where WD01>" & strSrvDate(1) & " Order By 1 "
rsA.CursorLocation = adUseClient
'Add by Morgan 2003/12/31
rsA.MaxRecords = 1

rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    If Val(strDate) >= Val(strSrvDate(1)) And Val(strDate) < Val("" & rsA.Fields(0).Value) Then
        WorkDayCheck = True
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function
'93.6.27 ADD BY SONIA 自prjTaieDll 複製過來
'讀取分案資料,intCaseKind系統分類
Private Function ReadSeparateCaseRst(ByRef intCaseKind As Integer, ByRef intWhere As Integer, ByRef strGroup As String, ByRef intKind1 As Integer, ByRef intKind2 As Integer, ByRef intKind3 As Integer, Optional strSeparate1 As String, Optional strSeparate2 As String, Optional strSeparate3 As String, Optional strSeparate4 As String, Optional intNone As Integer) As ADODB.Recordset
Dim strSql As String, strSQL1 As String, rsRecordset As New ADODB.Recordset
Dim strDateLine As String
Dim strField As String 'Add by Amy 2015/01/22
Dim strWhere  As String 'Add by Amy 2022/11/16

'Add by Amy 2015/01/22 操作人員為北所且北所分案日期有值才顯示辦人
'Modify by Amy 2022/11/03 接洽單電子收文上線後直接顯示(cp14=cra09)
 If strSrvDate(1) >= 接洽單電子收文啟用日 Then
    strField = " staff1.st02 as s11, "
 Else
    If pub_strUserOffice = "1" Then
       strField = " Decode(Nvl(cp157,0),0,'',staff1.st02) as s11, "
    Else
       strField = " staff1.st02 as s11, "
    End If
 End If

strSql = "select " & SQLDate("cp05") & " s01,cp09 s02,"
'Modify By Cheng 2002/07/16
'加是否算案件數欄
'strSQL1 = "select substr(cp05,1,4)" + strDateLine + "||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2) s01,cp09 s02,nvl(sp05,nvl(sp06,sp07)) s03,decode(sp09," + CNULL(大陸國家代號) + ",cpm04,cpm03) s04,sp01 s05,sp02 s06,sp02 s07,replace(sp03,'0') s08,replace(sp04,'00') s09,staff.st02 s10,staff1.st02 s11, substr(cp06,1,4)" + strDateLine + "||'/'||substr(cp06,5,2)||'/'||substr(cp06,7,2) s12, sp15 s13, substr(cp57,1,4)" + strDateLine + "||'/'||substr(cp57,5,2)||'/'||substr(cp57,7,2) s14, cp27 s15 from caseprogress,servicepractice,casepropertymap,staff,staff staff1 where CP01 in ('CPS', '') AND cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=staff.st01(+) and cp14=staff1.st01(+)"

'Add by Morgan 2004/9/13 加 CP86
'Modify by Amy 2015/01/22 操作人員為北所且北所分案日期有值才顯示辦人
'Modified by Morgan 2018/9/11 +S17
'Modified by Morgan 2021/2/23 改國外部案件也要輸入數量
'Modify by Amy 2022/10/21 +CP140/CP157
'Modify by Amy 2022/11/09 +CP122
strSQL1 = "select " & SQLDate("cp05") & " s01,cp09 s02,nvl(sp05,nvl(sp06,sp07)) s03," & _
   "decode(sp09," + CNULL(大陸國家代號) + ",cpm04,cpm03) s04,sp01 s05,sp02 s06,sp02 s07,replace(sp03,'0') s08," & _
   "replace(sp04,'00') s09,staff.st02 s10," & strField & SQLDate("cp06") & " s12, sp15 s13, " & SQLDate("cp57") & _
   " s14," & SQLDate("cp27") & " s15,CP26 AS S16, CP86,decode( CP140||substr(cp09,1,1),'A',''||CP156,'-') as S17,cp140,cp157,cp122" & _
   " from caseprogress,servicepractice,casepropertymap,staff,staff staff1" & _
   " where CP01 in ('CPS', '') AND cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp13=staff.st01(+) and cp14=staff1.st01(+)"
Select Case intCaseKind
   Case 專利
      'Modify By Cheng 2002/07/16
      '加是否算案件數欄
'      strSQL = strSQL + "nvl(pa05,nvl(pa06,pa07)) s03,decode(pa09," + CNULL(大陸國家代號) + ",cpm04,cpm03) s04,pa01 s05,pa02 s06,pa02 s07,replace(pa03,'0') s08,replace(pa04,'00') s09,staff.st02 s10,staff1.st02 s11, substr(cp06,1,4)" + strDateLine + "||'/'||substr(cp06,5,2)||'/'||substr(cp06,7,2) s12, pa57 s13, substr(cp57,1,4)" + strDateLine + "||'/'||substr(cp57,5,2)||'/'||substr(cp57,7,2) s14, cp27 s15 from caseprogress,patent,casepropertymap,staff,staff staff1 where cp01 in ('CFP', '') AND CP01=PA01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
        'Modify By Cheng 2003/10/03
        '若CP04<>"00", 收文號為"B"類且已有發文日的資料不顯示
        'Begin
'      strSQL = strSQL + "nvl(pa05,nvl(pa06,pa07)) s03,decode(pa09," + CNULL(大陸國家代號) + ",cpm04,cpm03) s04,pa01 s05," & _
'         "pa02 s06,pa02 s07,replace(pa03,'0') s08,replace(pa04,'00') s09,staff.st02 s10,staff1.st02 s11," & _
'         SQLDate("cp06") & " s12, pa57 s13," & SQLDate("cp57") & "s14," & SQLDate("cp27") & " s15,cp26 as S16 from " & _
'         "caseprogress,patent,casepropertymap,staff,staff staff1 where cp01 in ('CFP', '') AND CP01=PA01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
      'Modify by Amy 2015/01/22 操作人員為北所且北所分案日期有值才顯示辦人
      'Modified by Morgan 2018/9/11 +S17
      'Modified by Morgan 2021/2/23 改國外部案件也要輸入數量
      'Modify by Amy 2022/10/21 +CP140/CP157
      'Modify by Amy 2022/11/09 +CP122
      strSql = strSql + "nvl(pa05,nvl(pa06,pa07)) s03,decode(pa09," + CNULL(大陸國家代號) + ",cpm04,cpm03) s04,pa01 s05," & _
         "pa02 s06,pa02 s07,replace(pa03,'0') s08,replace(pa04,'00') s09,staff.st02 s10," & strField & _
         SQLDate("cp06") & " s12, pa57 s13," & SQLDate("cp57") & "s14," & SQLDate("cp27") & " s15,cp26 as S16, CP86" & _
         ",decode( CP140||substr(cp09,1,1),'A',''||CP156,'-') as S17,cp140,cp157,cp122" & _
         " from caseprogress,patent,casepropertymap,staff,staff staff1" & _
         " where cp01 in ('CFP', '') AND CP01=PA01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
         " And ( CP04='00' Or ( ( CP04<>'00' And ( CP09<'B' Or CP09>'C' ) ) Or ( CP04<>'00' And ( CP09>'B' And CP09<'C' ) And CP27 Is Null ) ) ) "
        'End
End Select
strSql = strSql + " and cp01=cpm01(+) and cp10=cpm02(+) and cp13=staff.st01(+) and cp14=staff1.st01(+)"

If strSeparate1 = "" Then strSeparate1 = "0"
If strSeparate2 = "" Then strSeparate2 = "0"
Select Case intKind1
   Case 0
      strSql = strSql + " and cp05 between " + strSeparate1 + " and " + strSeparate2
      strSQL1 = strSQL1 + " and cp05 between " + strSeparate1 + " and " + strSeparate2
   Case 1
      strSql = strSql + " and cp01=" + CNULL(strSeparate1) + " and cp02=" + CNULL(strSeparate2) + " and cp03=" + CNULL(strSeparate3) + " and cp04=" + CNULL(strSeparate4)
      strSQL1 = strSQL1 + " and cp01=" + CNULL(strSeparate1) + " and cp02=" + CNULL(strSeparate2) + " and cp03=" + CNULL(strSeparate3) + " and cp04=" + CNULL(strSeparate4)
End Select

strSql = strSql + " and substr(cp09,1,1)"
strSQL1 = strSQL1 + " and substr(cp09,1,1)"

If intKind2 = 0 Then
   strSql = strSql + " in(" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")"
   strSQL1 = strSQL1 + " in(" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")"
Else
   strSql = strSql + "=" + CNULL(主管機關來函)
   strSQL1 = strSQL1 + "=" + CNULL(主管機關來函)
End If

If intNone = 1 Then
   'Modify by Amy 2015/01/22 北所需顯示北所分案日沒值的資料
   If pub_strUserOffice = "1" Then
      'modify by sonia 2015/11/10 郭說選未分案時已取消收文不出現CFP-026927,故加nvl(cp57,0)=0條件
      strSql = strSql & " and (cp14 is null or cp157 is null) and nvl(cp57,0)=0 "
      strSQL1 = strSQL1 & " and (cp14 is null or cp157 is null) and nvl(cp57,0)=0 "
   Else
      'modify by sonia 2015/11/10 郭說選未分案時已取消收文不出現CFP-026927,故加nvl(cp57,0)=0條件
      strSql = strSql & " and cp14 is null and nvl(cp57,0)=0"
      strSQL1 = strSQL1 & " and cp14 is null and nvl(cp57,0)=0"
   End If
End If

'Add by Morgan 2008/2/12
If intNone = 3 Then
   strSql = strSql & " AND CP86 is null"
   strSQL1 = strSQL1 & " AND CP86 is null"
End If

'93.6.27 add by sonia 加收文所別
If intOpt1 = 0 Then
   If Text1 <> "1" And Text2 <> "1" Then
      strSql = strSql & " AND staff.ST06 >= '" & Text1 & "' AND staff.ST06 <= '" & Text2 & "' "
      strSQL1 = strSQL1 & " AND staff.ST06 >= '" & Text1 & "' AND staff.ST06 <= '" & Text2 & "' "
   Else
      strSql = strSql & " AND ((staff.ST06 >= '" & Text1 & "' AND staff.ST06 <= '" & Text2 & "') OR staff.ST06='5') "
      strSQL1 = strSQL1 & " AND ((staff.ST06 >= '" & Text1 & "' AND staff.ST06 <= '" & Text2 & "') OR staff.ST06='5') "
   End If
End If
'93.6.27 END

'Modify by Amy 2022/11/16 未分案才加判斷條件
If m_QueryType = 2 Then
   strWhere = "And (F0308='A7' Or F0309='" & Flow_已分案 & "' Or F0301 IS NULL) "
'Modify By Sindy 2023/6/8
ElseIf m_QueryType = 4 Then '電子收文未分案
   strWhere = "And F0308='A7' And F0309='" & Flow_處理中 & "' And F0301 IS NOT NULL "
   '2023/6/8 END
End If
'Modify By Cheng 2002/07/16
'strSQL = "select '' V,s01 收文日期,s02 總收文號,s03 案件名稱,s04 案件性質,s05 本所案號,s06 本所案號,s07 本所案號,s08 本所案號,s09 本所案號,s10 智權人員,s11 承辦人, s12 本所期限, s13 是否閉卷, s14 取消收文日, s15 發文日 from (" + strSQL + " union " + strSQL1 + ") order by s02"
'Modify by Morgan 2004/9/13 加註記欄 CP86
'Modified by Morgan 2018/9/11 +s17 as 數量(接洽單收文數)
'Modify by Amy 2022/10/28 +目前表單狀態
'Modify by Amy 2022/11/09 +CP122
strSql = "select '' V, s17 as 數量,s01 收文日期,s02 總收文號, cp86 註記,s03 案件名稱,s04 案件性質,s05 本所案號,s06 本所案號,s07 本所案號,s08 本所案號,s09 本所案號,s10 智權人員" & _
                ",Decode(F0309,'" & Flow_處理中 & "',Decode(Decode(Decode(cp157,null,0,1),1,'已分案',F0309),'已分案','已分案'," & ShowFlow表單狀態中文 & ")," & ShowFlow表單狀態中文 & ") as 目前表單狀態,s11 承辦人, s12 本所期限,s16 as 是否算案件數 , s13 是否閉卷, s14 取消收文日, s15 發文日,CP122 " & _
            "from Flow003,(" + strSql + " union " + strSQL1 + ") Where CP140=F0301(+) " & strWhere & _
            "order by s02"
'end 2022/11/16
Set ReadSeparateCaseRst = ClsPDReadRst(strSql)
End Function

'93.6.27 ADD BY SONIA
Private Sub Text1_GotFocus()
   InverseTextBox Text1
   CloseIme
End Sub

Private Sub Text2_GotFocus()
   InverseTextBox Text2
   CloseIme
End Sub
'93.6.27 END

'Added by Morgan 2018/9/11
Private Function GetValue(pRow As Integer, pFieldName As String) As String
   Dim ii As Integer
   With Me.grdDataList
   For ii = 0 To .Cols - 1
      If UCase(.TextMatrix(0, ii)) = UCase(pFieldName) Then
         If pRow = 0 Then
            '回傳第幾欄
            GetValue = ii
         Else
            '回傳欄位內容
            GetValue = .TextMatrix(pRow, ii)
         End If
         Exit For
      End If
   Next
   End With
End Function
