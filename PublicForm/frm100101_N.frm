VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_N 
   BorderStyle     =   1  '單線固定
   Caption         =   "合約資料查詢"
   ClientHeight    =   6590
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6590
   ScaleWidth      =   5130
   Tag             =   "加班資料"
   Begin VB.CommandButton CmdOK 
      Caption         =   "回前畫面(&U)"
      Height          =   315
      Index           =   1
      Left            =   630
      TabIndex        =   9
      Top             =   30
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   405
      Left            =   30
      TabIndex        =   6
      Top             =   5460
      Width           =   4485
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "開啟"
         Height          =   315
         Left            =   90
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   60
         Width           =   570
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdList 
      Height          =   4305
      Left            =   30
      TabIndex        =   2
      Top             =   1080
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   7602
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V  |  合約編號|合約名稱|檔案名稱|機密|備註|智權人員"
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
   Begin VB.CommandButton cmdSaveAtt 
      Caption         =   "下載"
      Height          =   315
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   525
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   315
      Index           =   0
      Left            =   1875
      TabIndex        =   1
      Top             =   30
      Width           =   765
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4170
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.Label lblCusName 
      Height          =   255
      Left            =   1140
      TabIndex        =   10
      Top             =   615
      Width           =   3000
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "5292;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "客戶名稱："
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   8
      Top             =   615
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(註:雙擊單檔預覽)"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   870
      Width           =   1425
   End
   Begin VB.Label Label1 
      Caption         =   "客戶編號："
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Top             =   390
      Width           =   960
   End
   Begin VB.Label lblCT01 
      Height          =   180
      Left            =   1140
      TabIndex        =   3
      Top             =   390
      Width           =   1830
   End
End
Attribute VB_Name = "frm100101_N"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/24 改成Form2.0 ; GrdList改字型=新細明體-ExtB、lblCusName
'Create by Amy 2017/12/20
Option Explicit

'附件宣告區
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
' 變數宣告區
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Const GrdMaxW = 9865
Private Const stFolderN As String = "CONTRACT" '指定FTP資料夾名稱
Public m_strKey As String '客戶編號
Public cmdState As Integer '紀錄作用按鍵
Dim arrGridHead
Dim i As Integer


Public Function QueryData(ByVal stNo As String) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    QueryData = False
    grdList.Clear
    pub_QL05 = ";編號：" & stNo & "(合約資料)" 'Add By Sindy 2025/8/27
    
    'Modify by Amy 2019/04/30 機密分國內/外
    strQ = "Select '' V,CT01 合約編號,CT03 合約名稱,CT08 檔案名稱,Decode(CT05,'C','國內',Decode(CT05,'F','國外','')) 機密,ST02 智權人員,CT07 備註," & _
               "CT01,CT02,CT06,CT08,CT09 From Contract,Staff " & _
               "Where CT06=ST01(+) And CT02='" & stNo & "' Order by CT01"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        If pub_QL04 <> "" Then InsertQueryLog (RsQ.RecordCount) 'Add By Sindy 2025/8/27
        Call AddGrdItem(RsQ)  '檔案多筆需增加列表示
        QueryData = True
    Else
        If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/27
        grdList.Rows = 2
        MsgBox "無合約資料", , "警告!!"
    End If
    Call SetGrd
    lblCusName = GetCusName 'Modify by Amy 2020/03/04
    
    Set RsQ = Nothing
    If intQ = 0 Then cmdok_Click (1)
End Function

Private Sub cmdok_Click(Index As Integer)
    '紀錄作用按鍵
    cmdState = Index
    PubShowNextData
End Sub

Private Function GetValue(pFieldN As String) As Integer
   Dim jj As Integer
 
    For jj = 1 To UBound(arrGridHead)
       If UCase(arrGridHead(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

Private Sub SetGrd()
    Dim arrGridHeadWidth
    Dim iCol As Integer

    arrGridHeadWidth = Array(200, 800, 2500, 2500, 600, 1000, 1800, 0, 0, 0, 0, 0)
    grdList.Visible = False
    grdList.Cols = UBound(arrGridHead) + 1
    For iCol = 0 To grdList.Cols - 1
        grdList.row = 0
        grdList.col = iCol
        grdList.ColWidth(iCol) = arrGridHeadWidth(iCol)
        grdList.TextMatrix(grdList.row, iCol) = arrGridHead(iCol)
    Next
    grdList.Visible = True
End Sub

Private Sub AddGrdItem(ByRef rsTmp As ADODB.Recordset)
    Dim IsFirst As Boolean, intRow As Integer
    
    IsFirst = True: intRow = 0
    rsTmp.MoveFirst
    Do While rsTmp.EOF = False
        intRow = intRow + 1
        If InStr("" & rsTmp.Fields("CT08"), ",") > 0 Then
            Call InsertRow(intRow, rsTmp)
        Else
            If IsFirst = False Then grdList.AddItem ""
            For i = GetValue("合約編號") To UBound(arrGridHead)
                grdList.TextMatrix(intRow, i) = "" & rsTmp.Fields(i)
            Next i
        End If
        If IsFirst = True Then IsFirst = False
        rsTmp.MoveNext
    Loop
End Sub

Private Sub InsertRow(ByRef intRow As Integer, ByRef rsTmp As ADODB.Recordset)
    Dim arrCT08, arrCT09
    Dim j As Integer, k As Integer
    Dim stTmp As String
    
    arrCT08 = Split(rsTmp.Fields(GetValue("CT08")), ",")
    arrCT09 = Split(rsTmp.Fields(GetValue("CT09")), ",")
    For j = LBound(arrCT08) To UBound(arrCT08)
        If intRow > 1 Then grdList.AddItem ""
        
        For k = 1 To UBound(arrGridHead)
            stTmp = ""
            If (j = LBound(arrCT08) And k <> GetValue("檔案名稱") And k <> GetValue("CT08") And k <> GetValue("CT09")) _
              Or (j <> LBound(arrCT08) And (k = GetValue("機密") Or k = GetValue("CT01") Or k = GetValue("CT02") Or k = GetValue("CT06"))) Then
                '第一筆才顯示合約編號
                stTmp = "" & rsTmp.Fields(k)
            End If
            If k = GetValue("檔案名稱") Or k = GetValue("CT08") Then stTmp = arrCT08(j)
            If k = GetValue("CT09") Then stTmp = arrCT09(j)
            
            grdList.TextMatrix(intRow, k) = stTmp
        Next k
        intRow = intRow + 1
    Next j
    intRow = intRow - 1
    
End Sub

Private Sub cmdOpenAtt_Click()
    Dim hLocalFile As Long
    Dim stPath As String, stFileN As String
    
    For i = 1 To grdList.Rows - 1
        grdList.row = i
        grdList.col = 0
        If grdList.TextMatrix(grdList.row, GetValue("V")) = "V" Then
            grdList.TextMatrix(grdList.row, GetValue("V")) = ""
            Call SetGrdColor(i, False)
            stFileN = grdList.TextMatrix(grdList.row, GetValue("CT08"))
            stPath = grdList.TextMatrix(grdList.row, GetValue("CT09"))
            If PUB_GetFtpFile(stPath, stFileN, stFolderN) Then
                ShellExecute hLocalFile, "open", stFileN, vbNullString, vbNullString, 1
            End If
        End If
    Next i
    
End Sub

Private Sub cmdSaveAtt_Click()
    Dim bolSelect As Boolean, bolCancel As Boolean
    Dim stFolderPath As String, stFileN As String, stPath As String
    
    stFolderPath = BrowseForFolder()
    If stFolderPath = MsgText(601) Then Exit Sub
    
    For i = 1 To grdList.Rows - 1
        If grdList.TextMatrix(i, GetValue("V")) = "V" Then
            bolCancel = False
            stFileN = grdList.TextMatrix(grdList.row, GetValue("CT08"))
            stPath = grdList.TextMatrix(grdList.row, GetValue("CT09"))
            If SaveNextAtt(stFolderPath, stFileN, stPath) = False Then
                bolCancel = True
                Exit For
            End If
            grdList.TextMatrix(i, GetValue("V")) = ""
            bolSelect = True
            Call SetGrdColor(i, False)
        End If
    Next i
    If bolCancel = True Then Exit Sub
    If bolSelect = False Then
        MsgBox "無檔案可存檔！"
    Else
        MsgBox "檔案已存於" & stFolderPath
    End If
    
End Sub

Private Function SaveNextAtt(ByVal stFolderPath As String, ByVal stCT08 As String, ByVal stCT09 As String) As Boolean
    Dim stLocalPath As String
    
    SaveNextAtt = False
    stLocalPath = stFolderPath & stCT08
    If stLocalPath = MsgText(601) Then Exit Function
    
    If Dir(stLocalPath) <> "" Then
        stLocalPath = ""
        '因檔案若已存在,且為開啟中仍可刪除(.txt),故檔案已存在即不可再下載
        MsgBox "檔案[ " & stCT08 & " ]已存在不可再下載!!", , MsgText(5)
        Exit Function
    End If
    
    If Dir(stFolderPath, vbDirectory) = "" Then
        MkDir stFolderPath
    End If
    If PUB_GetFtpFile(stCT09, stLocalPath, stFolderN) = False Then
        MsgBox "無法儲存檔案[ " & stCT08 & " ]！"
        GoTo RunExit
    End If
        
    SaveNextAtt = True
    Exit Function
    
RunExit:
    Screen.MousePointer = vbDefault
End Function

Private Sub Form_Activate()
    If Me.WindowState = 0 Then Me.WindowState = 2 '最大化
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    arrGridHead = Array("V", "合約編號", "合約名稱", "檔案名稱", "機密", "智權人員", "備註", _
                            "CT01", "CT02", "CT06", "CT08", "CT09")
        
    If Pub_StrUserSt03 = "M51" And InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") Then
       Me.Height = 6600
    Else
       Me.Height = 6120
    End If
    grdList.Width = GrdMaxW
    Me.WindowState = 2 '最大化
End Sub

Private Sub Form_Resize()
    If Me.Height > 6000 Then
        grdList.Height = Me.Height - grdList.Top - Frame1.Height - 380
        Frame1.Top = grdList.Top + grdList.Height - 20
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm100101_N = Nothing
End Sub

Public Sub PubShowNextData()
    Select Case cmdState
       Case 1
          tmpBol = fnCancelNowFormAndShowParentForm(Me)
       Case 0
          fnCloseAllFrm100
       Case Else
    End Select
End Sub

Private Sub grdList_Click()
    Dim nRow As Integer
    
    grdList.row = grdList.MouseRow
    grdList.col = 0
    nRow = grdList.row
    If nRow = 0 Then Exit Sub
    
    '機密判斷權限,有權限才可勾選
    'Modify by Amy 2019/04/30 機密分為國內機密(原:Y 改為C)/國外機密(F)
    If grdList.TextMatrix(nRow, GetValue("V")) = "" And grdList.TextMatrix(nRow, GetValue("機密")) <> MsgText(601) _
      And grdList.TextMatrix(grdList.row, GetValue("CT09")) <> MsgText(601) Then
        If ChkContractLimit(False, Left(GetNewFagent(lblCT01.Caption), 8), True, grdList.TextMatrix(nRow, GetValue("機密"))) = False Then
            Exit Sub
        End If
    End If
    'end 2019/04/30
    
    grdList.Visible = False
    If grdList.TextMatrix(grdList.row, GetValue("V")) = "" Then
        If Trim(grdList.TextMatrix(grdList.row, GetValue("CT09"))) = MsgText(601) Then
            grdList.Visible = True
            Exit Sub
        Else
            grdList.TextMatrix(grdList.row, GetValue("V")) = "V"
        End If
        Call SetGrdColor(nRow, True)
    Else
        grdList.TextMatrix(grdList.row, GetValue("V")) = ""
        For i = 0 To grdList.Cols - 1
             grdList.col = i
             grdList.CellBackColor = QBColor(15)
       Next i
    End If
    
    grdList.Visible = True
End Sub

Private Sub SetGrdColor(intRow As Integer, bolSetColor As Boolean)
    Dim j As Integer
   
    grdList.row = intRow
    For j = 1 To grdList.Cols - 1
        grdList.col = j
        If bolSetColor = True Then
            grdList.CellBackColor = &HFFC0C0 '選取
        Else
            grdList.CellBackColor = QBColor(15)
        End If
    Next j
   
End Sub

Private Function GetCusName() As String
    Dim rsTmp As New ADODB.Recordset
    Dim stQ As String, intQ As Integer
    
    GetCusName = ""
    intQ = 1
    stQ = "Select Nvl(CU04,Nvl(CU05,CU06)) From Customer Where CU01='" & lblCT01 & "' And CU02='0' "
    Set rsTmp = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        '中->英->日
        GetCusName = "" & rsTmp.Fields(0)
    End If
    
    Set rsTmp = Nothing
End Function

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

