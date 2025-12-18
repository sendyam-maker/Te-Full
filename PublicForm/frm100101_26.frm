VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_26 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "客戶端平台帳號資料查詢"
   ClientHeight    =   6210
   ClientLeft      =   180
   ClientTop       =   990
   ClientWidth     =   8955
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSaveAtt 
      Caption         =   "下載"
      Height          =   255
      Index           =   0
      Left            =   8400
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5430
      Width           =   615
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢帳號"
      Height          =   465
      Index           =   0
      Left            =   8400
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Frame Frame1 
      Caption         =   "複製區"
      ForeColor       =   &H000000C0&
      Height          =   465
      Left            =   1200
      TabIndex        =   21
      Top             =   2970
      Width           =   7170
      Begin VB.TextBox textCD04 
         BackColor       =   &H8000000F&
         Height          =   264
         Left            =   4350
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   150
         Width           =   1815
      End
      Begin VB.TextBox textCD03 
         BackColor       =   &H8000000F&
         Height          =   264
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   150
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "密　碼："
         Height          =   180
         Index           =   4
         Left            =   3630
         TabIndex        =   23
         Top             =   150
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "帳　號："
         Height          =   180
         Index           =   5
         Left            =   750
         TabIndex        =   22
         Top             =   150
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdOpenIE 
      Caption         =   "進入網站"
      Default         =   -1  'True
      Height          =   405
      Left            =   8400
      TabIndex        =   2
      Top             =   990
      Width           =   525
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   6840
      TabIndex        =   11
      Top             =   30
      Width           =   1230
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   0
      Left            =   8100
      TabIndex        =   12
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdOpenAtt 
      Caption         =   "開啟"
      Height          =   255
      Index           =   0
      Left            =   8400
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5130
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   180
      Index           =   0
      Left            =   1830
      MaxLength       =   9
      TabIndex        =   13
      Text            =   "X1256400"
      Top             =   210
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   1485
      Left            =   1200
      TabIndex        =   3
      Top             =   1440
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   2619
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   5
      FixedCols       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "帳號              |密碼                     |身份別           |下次更新日期  |註解                   "
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
      _Band(0).Cols   =   5
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   420
      Top             =   2490
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.ListBox lstUsers 
      Height          =   780
      Index           =   1
      Left            =   1200
      TabIndex        =   6
      Top             =   3480
      Width           =   7170
      VariousPropertyBits=   746586139
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "12647;1376"
      MatchEntry      =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstCW 
      Height          =   960
      Left            =   1200
      TabIndex        =   0
      Top             =   450
      Width           =   7170
      VariousPropertyBits=   746586139
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "12647;1693"
      MatchEntry      =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstAtt 
      Height          =   780
      Left            =   1200
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5040
      Width           =   7170
      VariousPropertyBits=   746586139
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "12647;1376"
      MatchEntry      =   0
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCW05 
      Height          =   780
      Left            =   1200
      TabIndex        =   7
      Top             =   4260
      Width           =   7170
      VariousPropertyBits=   -1466941409
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "12647;1376"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客　　戶："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   26
      Top             =   3480
      Width           =   900
   End
   Begin VB.Label Label4 
      Caption         =   "註：紅色為期限將到或過期，請通知管理者更新"
      ForeColor       =   &H000000C0&
      Height          =   765
      Left            =   60
      TabIndex        =   25
      Top             =   1680
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "註：點選平台資料列即可查詢帳號資料"
      ForeColor       =   &H000000C0&
      Height          =   585
      Left            =   60
      TabIndex        =   24
      Top             =   720
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "帳號資料："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   20
      Top             =   1470
      Width           =   900
   End
   Begin VB.Label Label23 
      Caption         =   "CREATE : 　　　  101/09/03  13:54:00          UPDATE : 　　　  101/09/04  09:21:44"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   19
      Top             =   5880
      Width           =   8490
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "平　　台："
      Height          =   180
      Left            =   180
      TabIndex        =   18
      Top             =   510
      Width           =   900
   End
   Begin MSForms.Label LblIn01 
      Height          =   285
      Left            =   2970
      TabIndex        =   17
      Top             =   210
      Width           =   1800
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "附件或憑證："
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   16
      Top             =   5040
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "備　　註："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   15
      Top             =   4260
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人/代理人編號："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   210
      Width           =   1665
   End
End
Attribute VB_Name = "frm100101_26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/06 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Created by Sindy 2012/9/20
Option Explicit

' 變數宣告區
Dim MyArr As Variant
Dim m_AttachPath As String
Dim i As Integer, j As Integer

'附件
Dim m_FilesRemoved() As String
Dim ii As Integer, jj As Integer
Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194

'紀錄作用按鍵
Public cmdState As Integer
Dim m_CW01 As String


Private Function BrowseForFolder(Optional sCaption As String = "請選擇欲儲存的位置", Optional sDefault As String) As String
    Const BIF_RETURNONLYFSDIRS = 1
    Const MAX_PATH = 260
    Dim lPos As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, tBrowse As BrowseInfo

    With tBrowse
        'Set the owner window
        .hwndOwner = GetActiveWindow        'Me.hWnd in VB
        .lpszTitle = sCaption
        .ulFlags = BIF_RETURNONLYFSDIRS     'Return only if the user selected a directory
    End With

    'Show the dialog
    lpIDList = SHBrowseForFolder(tBrowse)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        lPos = InStr(sPath, vbNullChar)
        If lPos Then
            BrowseForFolder = Left$(sPath, lPos - 1)
            If Right$(BrowseForFolder, 1) <> "\" Then
                BrowseForFolder = BrowseForFolder & "\"
            End If
        End If
    Else
        'User cancelled, return default path
        BrowseForFolder = sDefault
    End If
End Function

Public Sub PubShowNextData()
Select Case cmdState
   Case 1
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Case 0
      fnCloseAllFrm100
   Case Else
End Select
End Sub

Private Sub cmdOK_Click(Index As Integer)
'紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   'Modify By Sindy 2021\5\19
   'm_AttachPath = App.path & "\SeminarAttach"
   m_AttachPath = App.path & "\SeminarAttach\" & strUserNum
   '2021\5\19 END
   
   cmdState = -1
   Pub_Can_Copy_Pic = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Pub_Can_Copy_Pic = False
   KillAttach
   Set frm100101_26 = Nothing
End Sub

Private Sub KillAttach()
On Error Resume Next
   'Modify By Sindy 2021/2/3 刪不掉,改用函數
'   If Dir(m_AttachPath & "\.") <> "" Then
'      Kill m_AttachPath & "\*.*"
'   End If
   'Modify By Sindy 2021\5\19
   PUB_KillTempFile "SeminarAttach\" & strUserNum & "\*.*"
   '2021\5\19 END
   '2021/2/3 END
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow grd1, x, y, nCol, nRow
   grd1.col = nCol
   grd1.row = nRow
End Sub

Private Sub grd1_SelChange()
Dim tmpMouseRow
Dim strText As String
   
   grd1.Visible = False
   tmpMouseRow = grd1.row
   grd1.Visible = True
   If tmpMouseRow <> 0 And grd1.TextMatrix(tmpMouseRow, 0) <> "" Then
      grd1.row = tmpMouseRow
      grd1.col = 0
'      If grd1.CellBackColor = QBColor(15) Then
         grd1.Visible = False
         For j = 1 To grd1.Rows - 1
            grd1.row = j
            'Modify By Sindy 2016/12/13 + And DBDATE(GRD1.TextMatrix(j, 4)) <> 19221111 11/11/11:代表停止使用
            If DBDATE(grd1.TextMatrix(j, 4)) <> "" And DBDATE(grd1.TextMatrix(j, 4)) <> 19221111 And DBDATE(grd1.TextMatrix(j, 4)) <= CompWorkDay(4, strSrvDate(1), 0) Then
               For i = 0 To grd1.Cols - 1
                  grd1.col = i
                  grd1.CellBackColor = &H8080FF '紅色
               Next i
            Else
               For i = 0 To grd1.Cols - 1
                  grd1.col = i
                  grd1.CellBackColor = QBColor(15)
               Next i
            End If
         Next j
         grd1.row = tmpMouseRow
         For i = 0 To grd1.Cols - 1
            grd1.col = i
            'Modify By Sindy 2016/12/13 + And DBDATE(GRD1.TextMatrix(tmpMouseRow, 4)) <> 19221111 11/11/11:代表停止使用
            If i = 4 And DBDATE(grd1.TextMatrix(tmpMouseRow, 4)) <> "" And DBDATE(grd1.TextMatrix(tmpMouseRow, 4)) <> 19221111 And DBDATE(grd1.TextMatrix(tmpMouseRow, 4)) <= CompWorkDay(4, strSrvDate(1), 0) Then
               grd1.CellBackColor = &H8080FF '紅色
            Else
               grd1.CellBackColor = &HFFC0C0
            End If
         Next i
         textCD03 = grd1.TextMatrix(tmpMouseRow, 0) '帳號
         textCD04 = grd1.TextMatrix(tmpMouseRow, 1) '密碼
         grd1.Visible = True
'      End If
   End If
End Sub

Private Sub ShowMsg(ByVal St As String)
   MsgBox St, vbInformation
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("cw06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("cw06")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("cw06"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("cw07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("cw07")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("cw07"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("cw08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("cw08")) = False Then
         strTemp = rsSrcTmp.Fields("cw08")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("cw09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("cw09")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("cw09"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("cw10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("cw10")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("cw10"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("cw11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("cw11")) = False Then
         strTemp = rsSrcTmp.Fields("cw11")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

' 將資料庫中的資料更新到所有欄位中
Public Function StrMenu() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strCustID As String, strCon As String
   
   strCustID = Me.Tag
   m_CW01 = ""
   lstCW.Clear
   ClearField
   
   Text1(0) = strCustID
   LblIn01 = ""
   If Left(strCustID, 1) = "X" Then
      LblIn01 = GetPrjPeople1(strCustID, "1")
   ElseIf Left(strCustID, 1) = "Y" Then
      LblIn01 = GetPrjName1(strCustID)
   End If
   
   StrMenu = False
   
   If Pub_StrUserSt03 <> "M51" Then
      strCon = "and (instr(cd06,'" & strUserNum & "')>0 or cd06 is null) "
   End If
   '客戶所使用的平台清單
'   strSql = "SELECT distinct cw01,cw02,cw03,decode(cw03,'1','IP管理','2','檔案存取','3','電子帳單','4','憑證'),cw12 FROM custweb,custwebid " & _
'            "WHERE cw01=cd01 " & _
'            "and instr(cw04,'" & strCustID & "')>0 " & strCon & _
'            "order by cw03 asc,cw02 asc"
   'Modified by Morgan 2017/10/24 decode(cw03,'1','IP管理','2','檔案存取','3','電子帳單','4','憑證') -> PUB_GetCW03SQL
   strSql = "SELECT cw01,cw02,cw03," & PUB_GetCW03SQL & ",cw12,'1' as sort " & _
            "FROM (select * from custweb where instr(cw04,'" & strCustID & "')>0) cw1,custwebid " & _
            "Where cw01 = cd01(+) And cw12 Is Not Null " & strCon & _
            "Union SELECT cw01,cw17,cw03," & PUB_GetCW03SQL & ",cw12,'2' as sort " & _
            "FROM (select * from custweb where instr(cw04,'" & strCustID & "')>0) cw1,custwebid " & _
            "Where cw01 = cd01(+) And cw17 Is Not Null " & strCon & _
            "Union SELECT cw01,cw18,cw03," & PUB_GetCW03SQL & ",cw12,'3' as sort " & _
            "FROM (select * from custweb where instr(cw04,'" & strCustID & "')>0) cw1,custwebid " & _
            "Where cw01 = cd01(+) And cw18 Is Not Null " & strCon & _
            "order by cw03 desc,sort desc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      StrMenu = True
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         If m_CW01 = "" Then
            m_CW01 = rsTmp.Fields(0)
         End If
         lstCW.AddItem "（" & rsTmp.Fields(3) & "）" & rsTmp.Fields(4) & IIf(rsTmp.Fields(5) = 2, " -網址2", IIf(rsTmp.Fields(5) = 3, " -網址3", "")) & "：" & rsTmp.Fields(1) & _
         "                                                                                                        @" & _
         rsTmp.Fields(0), 0
         rsTmp.MoveNext
      Loop
      lstCW.Selected(lstCW.ListCount - 1) = True
'      SetListScroll lstCW
      Call UpdateCtrlData2(m_CW01) '查詢平台明細資料
   End If
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Function

'查詢平台明細資料
Private Sub UpdateCtrlData2(strCW01 As String)
Dim rsTmp2 As New ADODB.Recordset
Dim strSql As String, strCon As String
   
   ClearField
   Call SetGrd
   
   '平台資料
   strSql = "SELECT * FROM custweb " & _
            "WHERE cw01='" & strCW01 & "' "
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp2.RecordCount > 0 Then
      textCW05 = "" & rsTmp2.Fields("cw05")
      SetlstUsers 1, rsTmp2.Fields("cw04")
      '更新CUID
      UpdateCUID rsTmp2
      
      'Added by Morgan 2015/3/13
      'Y48292030 電子帳單提醒
      If Me.Tag = "Y48292030" And rsTmp2.Fields("cw03") = "3" Then
         If Val(Right(strSrvDate(1), 2)) >= 16 Then
            MsgBox "帳單不得於16日-31日上傳", vbInformation, "Y48292030 電子帳單"
         End If
      End If
      'end 2015/3/13
   End If
   rsTmp2.Close
   
   If Pub_StrUserSt03 <> "M51" Then
      strCon = "and (instr(cd06,'" & strUserNum & "')>0 or cd06 is null) "
   End If
   '帳號資料
   strSql = "SELECT cd03 as 帳號,cd04 as 密碼,decode(cd05,'1','管理者','2','使用者') as 身份別,sqldatet(cd09) as 建置日期,sqldatet(cd07) as 下次更新日期,cd08 as 註解 FROM custweb,custwebid " & _
            "WHERE cw01=cd01(+) " & _
            "and cw01='" & strCW01 & "' " & strCon & _
            "order by cd05 asc,cd03 asc"
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp2.RecordCount > 0 Then
      Set grd1.Recordset = rsTmp2
      For i = 1 To grd1.Rows - 1
         'Modify By Sindy 2016/12/13 + And DBDATE(GRD1.TextMatrix(i, 4)) <> 19221111 11/11/11:代表停止使用
         'Modify By Sindy 2018/3/22
         If Trim(grd1.TextMatrix(i, 4)) <> "" Then
         '2018/3/22 END
            If DBDATE(grd1.TextMatrix(i, 4)) <> "" And DBDATE(grd1.TextMatrix(i, 4)) <> 19221111 And DBDATE(grd1.TextMatrix(i, 4)) <= CompWorkDay(4, strSrvDate(1), 0) Then
               grd1.Visible = False
               grd1.row = i
               For j = 0 To grd1.Cols - 1
                    grd1.col = j
                    grd1.CellBackColor = &H8080FF '紅色
               Next j
               grd1.Visible = True
            End If
         End If
      Next i
   End If
   rsTmp2.Close
   
   '附件檔
   strExc(0) = "select cf02,cf03 from custwebfile where cf01=" & strCW01 & " order by 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         lstAtt.AddItem .Fields("cf02") & " (" & Round(.Fields("cf03") / 1024, 2) & " KB)", 0
'         lstAtt.ItemData(0) = 1
         .MoveNext
      Loop
      End With
      cmdOpenAtt(0).Enabled = True
      'cmdSelect(0).Enabled = True
      cmdSaveAtt(0).Enabled = True
   End If
   'If lstAtt.ListCount > 0 Then SetListScroll lstAtt

EXITSUB:
   Set rsTmp2 = Nothing
End Sub

Private Sub SetlstUsers(p_idx As Integer, p_stNums As String)
   Dim arrID
   lstUsers(p_idx).Clear
   If p_stNums <> "" Then
      Select Case p_idx
         Case 0 '員工編號
            strExc(0) = "select st01,st02 from staff where instr('" & p_stNums & "',st01)>0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               arrID = Split(p_stNums, ",")
               With RsTemp
               '照原順序排
               For intI = UBound(arrID) To LBound(arrID) Step -1
                  .MoveFirst
                  Do While Not .EOF
                     If .Fields("st01") = arrID(intI) Then
                        'lstUsers(p_idx).AddItem "" & .Fields(1), 0
                        lstUsers(p_idx).AddItem "" & .Fields(1) & "                                                            @" & .Fields(0), 0
                        'lstUsers(p_idx).ITEMDATA(0) = PUB_Id2Num(.Fields(0), CStr(p_idx)) 'Removed by Morgan 2016/9/8 非維護不用
                        .MoveLast
                     End If
                     .MoveNext
                  Loop
               Next
               End With
            End If
         Case 1, 2 '客戶編號
            strExc(0) = "select cu01||cu02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) from customer where cu01>' ' and instr('" & p_stNums & "',cu01||cu02)>0" & _
                        " union" & _
                        " select fa01||fa02,NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)) from fagent where fa01>' ' and instr('" & p_stNums & "',fa01||fa02)>0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               arrID = Split(p_stNums, ",")
               With RsTemp
               '照原順序排
               For intI = UBound(arrID) To LBound(arrID) Step -1
                  .MoveFirst
                  Do While Not .EOF
                     If .Fields(0) = arrID(intI) Then
                        lstUsers(p_idx).AddItem "" & .Fields(0) & " " & .Fields(1), 0
                        'lstUsers(p_idx).ITEMDATA(0) = PUB_Id2Num(.Fields(0), CStr(p_idx)) 'Removed by Morgan 2016/9/8 非維護不用
                        .MoveLast
                     End If
                     .MoveNext
                  Loop
               Next
               End With
            End If
      End Select
   End If
End Sub

'Removed by Morgan 2016/9/8 客戶/代理人末三碼會有B,C,非數字會有錯,改呼叫公用函數
''員工或客戶編號轉數字
'Public Function PUB_Id2Num(pID As String, strType As String) As Long
'   Select Case strType
'      Case 0 '員工編號
'         PUB_Id2Num = "&H" & pID
'      Case 1 '客戶編號
'         If Left(Trim(pID), 1) = "X" Then
'            PUB_Id2Num = "1" & Mid(Trim(pID), 2, Len(Trim(pID)) - 1)
'         ElseIf Left(Trim(pID), 1) = "Y" Then
'            PUB_Id2Num = "2" & Mid(Trim(pID), 2, Len(Trim(pID)) - 1)
'         End If
'   End Select
'End Function
'end 2016/9/8

'轉換使用者姓名
Private Function ChangeCD06CN(strID As String, ByRef strText As String) As Boolean
Dim strTempName As String
   
   ChangeCD06CN = False
   If strID <> "" Then
      MyArr = Split(strID, ",")
      strText = ""
      For j = 0 To UBound(MyArr)
         If ClsPDGetStaff(MyArr(j), strTempName) = True Then
            strText = strText & "," & strTempName
         Else
            strText = strText & "," & MyArr(j)
         End If
      Next j
      strText = Mid(strText, 2, Len(strText))
      ChangeCD06CN = True
   End If
End Function

Private Sub ClearField()
   textCW05.Text = Empty
   Label23.Caption = Empty
   
   '帳號資料
   grd1.Clear
   grd1.Rows = 2
   Call SetGrd
   textCD03.Text = Empty
   textCD04.Text = Empty
   
   lstAtt.Clear '操作手冊
'   cmdOpenAtt(0).Enabled = False
'   cmdSelect(0).Enabled = False
'   cmdSaveAtt(0).Enabled = False
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   arrGridHeadText = Array("帳號", "密碼", "身份別", "建置日期", "下次更新日期", "註解")
   arrGridHeadWidth = Array(1200, 1200, 800, 800, 1200, 2000)
   grd1.Visible = False
   grd1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To grd1.Cols - 1
      grd1.row = 0
      grd1.col = iRow
      grd1.Text = arrGridHeadText(iRow)
      grd1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      'grd1.CellAlignment = flexAlignCenterCenter
   Next
   grd1.Visible = True
End Sub

'進入網站
Private Sub cmdOpenIE_Click()
Dim myweb As Object
Dim strCW02 As String
   
   Dim hLocalFile As Long
   
   'Modify By Sindy 2020/12/29
   MyArr = Split(lstCW.List(lstCW.ListIndex), "：")
   strCW02 = Trim(MyArr(1))
   MyArr = Split(strCW02, "@")
   strCW02 = Trim(MyArr(0))
   'Added by Morgan 2021/4/15 Tymetrix360 指定用Chrome開啟--經理
   'Modify By Sindy 2021/8/25 薛:平台資料0041,請修改為由chrome 開始。（網站已不支援ＩＥ）
   If m_CW01 = "0026" Or m_CW01 = "0041" Then
      PUB_OpenURL strCW02, 1
   Else
   'end 2021/4/15
      ShellExecute hLocalFile, "open", strCW02, vbNullString, vbNullString, 1
   End If
   Exit Sub
   '2020/12/29 END
   
   Set myweb = CreateObject("InternetExplorer.Application")
   Screen.MousePointer = vbHourglass
   With myweb
      .Toolbar = 0
      .Visible = True ' 顯示IE
      MyArr = Split(lstCW.List(lstCW.ListIndex), "：")
      strCW02 = Trim(MyArr(1))
      MyArr = Split(strCW02, "@")
      strCW02 = Trim(MyArr(0))
      .Navigate strCW02 ' 瀏覽網址 www.lativ.com.tw/Home/Login
'      ' 等待網頁載入完成
'      Do While .Busy
'         DoEvents
'      Loop
      '.Document.All("email").Value = "xxxx" '帳號 login
      '.Document.All("pw").Value = "xxxx"    '密碼 passwd
      '.Document.All("submit").Click         '登入 signIn
   End With
   Set myweb = Nothing ' 釋放IE 物件
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdQuery_Click(Index As Integer)
   If lstCW.ListIndex >= 0 Then
      MyArr = Split(lstCW.List(lstCW.ListIndex), "@")
      If m_CW01 <> Trim(MyArr(1)) Then
         m_CW01 = Trim(MyArr(1))
         Call UpdateCtrlData2(m_CW01) '查詢平台明細資料
      End If
   End If
End Sub

Private Function GetAttachFile(ByRef pFileName As String, Optional pSavePath As String) As Boolean
   Dim stAttPath As String
   Dim lngSize As Long
   Dim iFileNo As Integer
   Dim bytes() As Byte
   
On Error GoTo ErrHnd
   
   If pSavePath = "" Then
      If Dir(m_AttachPath, vbDirectory) = "" Then
         MkDir m_AttachPath
      End If
      stAttPath = m_AttachPath & "\" & pFileName
      '檔案已存在時不必重新下載
      If Dir(stAttPath) <> "" Then
         'Kill stAttPath
         pFileName = stAttPath
         GetAttachFile = True
         Exit Function
      End If
   Else
      stAttPath = pSavePath
   End If
      
   strExc(0) = "select * from custwebfile b where cf01=" & m_CW01 & " and cf02='" & ChgSQL(pFileName) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Dir(stAttPath) <> "" Then Kill stAttPath
      
      'Add By Sindy 2017/5/25
      If "" & RsTemp.Fields("cf08") <> "" Then
         GetAttachFile = PUB_GetFtpFile(RsTemp.Fields("cf08"), stAttPath, UCase("custwebfile"))
      Else
      '2017/5/25 END
         With RsTemp
         lngSize = Val(.Fields("cf03").Value)
         ReDim bytes(lngSize)
         If lngSize > 0 Then bytes() = .Fields("cf04").GetChunk(lngSize)
         End With
         iFileNo = FreeFile
         Open stAttPath For Binary Access Write As #iFileNo
         If lngSize > 0 Then Put #iFileNo, , bytes()
         Close #iFileNo
      End If
      
      pFileName = stAttPath
      GetAttachFile = True
   End If
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   If iFileNo > 0 Then Close #iFileNo
End Function

'開啟附件
Private Sub cmdOpenAtt_Click(Index As Integer)
   Dim hLocalFile As Long
   Dim stFileName As String
   Dim strAtt As String, strType As String
   
   Screen.MousePointer = vbHourglass
   
   If Index = 0 Then
      strAtt = lstAtt.List(lstAtt.ListIndex)
   End If
   
   If strAtt = "" Then
      MsgBox "請選擇欲開啟的附件！"
   Else
      For ii = 0 To lstAtt.ListCount - 1
         If lstAtt.Selected(ii) Then
            stFileName = lstAtt.List(ii)
            'stFileName = strAtt
            If InStrRev(stFileName, " (") > 0 Then
               stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
            End If
            
            If InStr(stFileName, "\") = 0 Then
               If GetAttachFile(stFileName) = False Then
                  Exit Sub
               End If
            End If
            
            ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
         End If
      Next ii
   End If
   
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSelect_Click(Index As Integer)
   Dim ii As Integer, oList As Object
   If Index = 0 Then
      Set oList = lstAtt
   End If
   
   For ii = 0 To oList.ListCount - 1
      lstAtt.Selected(ii) = True
   Next
End Sub

Private Sub cmdSaveAtt_Click(Index As Integer)
   
   Dim stFileName As String, stFolderPath As String, stFullName As String
   Dim bMultiFile As Boolean
   Dim ii As Integer, oList As Object
   
   Screen.MousePointer = vbHourglass
   
   If Index = 0 Then
      Set oList = lstAtt
   End If
   
   stFileName = ""
   bMultiFile = False
   For ii = 0 To oList.ListCount - 1
      If oList.Selected(ii) Then
         stFileName = oList.List(ii)
         If stFileName <> "" Then
            bMultiFile = True
            Exit For
         Else
            stFileName = oList.List(ii)
         End If
      End If
   Next
   
   If stFileName = "" Then
      MsgBox "請選擇欲存檔的附件！"
   Else
      '多選
      If bMultiFile Then
         stFolderPath = BrowseForFolder()
         If stFolderPath <> "" Then
            For ii = 0 To oList.ListCount - 1
               If oList.Selected(ii) Then
                  stFileName = oList.List(ii)
                  If InStrRev(stFileName, " (") > 0 Then
                     stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
                  End If
                  stFullName = stFolderPath & stFileName
                  If stFullName <> "" Then
                     If Dir(stFullName) <> "" Then
                        If MsgBox("檔案[ " & stFileName & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2) = vbNo Then
                           stFullName = ""
                        End If
                     End If
                     If stFullName <> "" Then
                        If GetAttachFile(stFileName, stFullName) = False Then
                           MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                        End If
                     End If
                  End If
               End If
            Next
         End If
      
      Else
         stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
         stFullName = GetSaveName(stFileName)
         If stFullName <> "" Then
            If Dir(stFullName) <> "" Then
               If MsgBox("檔案[ " & stFileName & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2) = vbNo Then
                  stFullName = ""
               End If
            End If
            If stFullName <> "" Then
               If GetAttachFile(stFileName, stFullName) = False Then
                  MsgBox "無法儲存檔案[ " & stFileName & " ]！"
               End If
            End If
         End If
      End If
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Function GetSaveName(ByVal pFileName As String) As String
   
On Error GoTo ErrHnd

   With CommonDialog1
      .CancelError = True
      .FileName = pFileName
      .Filter = "All Files (*.*)|*.*"
      .InitDir = PUB_Getdesktop
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowSave
      If .FileName <> "" Then
         GetSaveName = .FileName
      End If
   End With
   
   Exit Function
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Function

Private Function GetFileName(ByVal FullPath As String) As String
   Dim stItem As String, iPos As Integer
   
   stItem = FullPath
   iPos = InStr(stItem, "\")
   Do While iPos > 0
      stItem = Mid(stItem, iPos + 1)
      iPos = InStr(stItem, "\")
   Loop
   
   If InStrRev(stItem, " (") > 0 And Right(stItem, 1) = ")" Then
      stItem = Left(stItem, InStrRev(stItem, " (") - 1)
   End If
   
   GetFileName = stItem
End Function

Private Sub SetListScroll(oList As Object)
   Dim ii As Integer
   Dim lWnow As Long, lWmax As Long
   
   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next
  
   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

Private Sub lstCW_Click()
   If lstCW.ListIndex >= 0 Then
      Call cmdQuery_Click(0)
   End If
End Sub
