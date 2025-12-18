VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040112 
   BorderStyle     =   1  '單線固定
   Caption         =   "公文來函文檔整批匯入"
   ClientHeight    =   5736
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   8952
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   8952
   Begin VB.CommandButton cmdOK 
      Caption         =   "無缺檔"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   0
      Left            =   6060
      TabIndex        =   24
      Top             =   1176
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.FileListBox File1 
      Height          =   432
      Left            =   1440
      TabIndex        =   17
      Top             =   2040
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   615
      Top             =   2070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   555
      Left            =   2565
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   974
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   "檔案名稱"
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   810
      TabIndex        =   18
      Top             =   5430
      Width           =   5835
   End
   Begin VB.CommandButton cmdPath 
      Height          =   330
      Left            =   8595
      Picture         =   "frm040112.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   852
      Width           =   350
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "卷宗區"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   13
      Left            =   7032
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   1176
      Width           =   870
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "重整(&Q)"
      Height          =   345
      Left            =   7956
      TabIndex        =   13
      Top             =   1176
      Width           =   870
   End
   Begin VB.Frame Frame3 
      Caption         =   "匯入錯誤訊息："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3732
      Left            =   45
      TabIndex        =   11
      Top             =   1320
      Width           =   3960
      Begin VB.CheckBox Check2 
         Caption         =   "列印"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1500
         TabIndex        =   2
         Top             =   0
         Width           =   705
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3468
         ItemData        =   "frm040112.frx":0102
         Left            =   60
         List            =   "frm040112.frx":0104
         TabIndex        =   12
         Top             =   240
         Width           =   3840
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "缺檔案件："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3732
      Left            =   4035
      TabIndex        =   10
      Top             =   1320
      Width           =   4890
      Begin VB.CheckBox Check1 
         Caption         =   "列印"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1140
         TabIndex        =   3
         Top             =   0
         Width           =   705
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3432
         Left            =   48
         TabIndex        =   6
         Top             =   240
         Width           =   4776
         _ExtentX        =   8424
         _ExtentY        =   6054
         _Version        =   393216
         Cols            =   6
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "V|本所案號|案件性質|收文日|本所期限|缺檔數"
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   516
      Left            =   6828
      TabIndex        =   4
      Top             =   216
      Width           =   885
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "匯入(&T)"
      Height          =   516
      Left            =   5808
      TabIndex        =   1
      Top             =   216
      Width           =   885
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   1500
      TabIndex        =   0
      Text            =   "\\Pat1\OA_SCAN"
      Top             =   852
      Width           =   7065
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   516
      Left            =   7848
      TabIndex        =   5
      Top             =   216
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Height          =   465
      Left            =   30
      TabIndex        =   8
      Top             =   4980
      Width           =   8895
      Begin VB.TextBox Text2 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00FF0000&
         Enabled         =   0   'False
         Height          =   300
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
         Width           =   8820
      End
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1512
      TabIndex        =   26
      Top             =   480
      Visible         =   0   'False
      Width           =   2088
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3678;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      Caption         =   "輸入人員："
      Height          =   240
      Left            =   456
      TabIndex        =   25
      Top             =   528
      Visible         =   0   'False
      Width           =   948
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "1. 來函檔名規則：本所案號.PDF（ex.P123456.PDF）"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   3
      Left            =   96
      TabIndex        =   23
      Top             =   48
      Width           =   4056
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "2. 引證前案檔名規則：本所案號.說明.PDF（ ex.P123456.abc.PDF）"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   2
      Left            =   96
      TabIndex        =   22
      Top             =   264
      Width           =   5148
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "印表機："
      Height          =   180
      Left            =   90
      TabIndex        =   21
      Top             =   5490
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "總筆數："
      Height          =   180
      Index           =   0
      Left            =   7110
      TabIndex        =   20
      Top             =   5490
      Width           =   720
   End
   Begin VB.Label lblTotal 
      Height          =   180
      Left            =   7875
      TabIndex        =   19
      Top             =   5490
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "文檔存放路徑："
      Height          =   180
      Left            =   132
      TabIndex        =   7
      Top             =   912
      Width           =   1260
   End
End
Attribute VB_Name = "frm040112"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/11/30 Form2.0已修改 (無需修改)
'Created by Morgan 2014/3/27
Option Explicit

'ListBox 加卷軸
Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194
'---------------

'選擇資料夾
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
(lpBI As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
(ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
(ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
   hwndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type
'-------------

Public cmdState As Integer '紀錄作用按鍵
Public m_ProState As String 'Add by Amy 2020/01/08
Dim lPrevRow As Long '前次點選列

'列印報表用---
Dim PLeft(1 To 7) As Integer
Dim strTemp(1 To 7) As String
Dim iLine1 As Integer
Dim strPrinter As String
'-------------
Dim bolActived As Boolean
Dim dblMaxWidth As Double
Dim oFileSys As New FileSystemObject
Dim oFile As File

Private Sub Check1_Click()
   If MSHFlexGrid1.Rows > 1 Then
      If MSHFlexGrid1.TextMatrix(1, 1) = "" Then
         Check1.Value = 0
      End If
   End If
End Sub

Private Sub Check2_Click()
   If List1.ListCount = 0 Then
      Check2.Value = 0
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdImPort_Click()
   ImportFile
End Sub

Private Function ImportFile() As Boolean
   Dim strCP02 As String, strCP03 As String, strCP04 As String, strCP09 As String, strErr As String
   Dim iTotRows As Integer
   Dim ii As Integer
   Dim dblFCnt As Double
   Dim stSaveName As String
   Dim strFileName As String
   Dim bolUploadDone As Boolean
   Dim strCP01 As String 'Add by Amy 2020/01/08
   Dim bolChkOk As Boolean 'Add by Amy 2020/02/20 可上傳
   
On Error GoTo ErrHnd

   If IsEmptyText(txtPath) = True Then
      MsgBox "請選擇文檔存放路徑！", vbOKOnly, "檢核資料"
      cmdPath.SetFocus
      Exit Function
   'Modified by Morgan 2014/4/23
   'Dir 若磁碟機不存在或網路路徑第一層會發生執行階段錯誤
   'ElseIf Dir(txtPath & "\", vbDirectory) = "" Then
   'Modified by Morgan 2017/1/12
   'ElseIf oFileSys.FolderExists(txtPath) = False Then
   ElseIf PUB_ChkDir(txtPath) = False Then
      MsgBox "文檔存放路徑不存在，請重新選擇！"
      cmdPath.SetFocus
      Exit Function
   ElseIf Dir(txtPath & "\*.pdf") = "" Then
      MsgBox "資料夾 " & txtPath.Text & " 中沒有pdf檔！"
      cmdPath.SetFocus
      Exit Function
   End If
   
   Text2.Width = 0
   List1.Clear
   
   '檔名寫入grid2並排序
   Grid2.Clear
   Grid2.Cols = 1
   Grid2.Rows = 1
   File1.path = txtPath.Text
   File1.Refresh
   For dblFCnt = 0 To File1.ListCount - 1
      '檔名後4碼為.PDF者才須匯入
      If UCase(Right(Trim(File1.List(dblFCnt)), 4)) = ".PDF" Then
         Grid2.AddItem Trim(File1.List(dblFCnt))
      End If
   Next dblFCnt
   
   Grid2.col = 0
   Grid2.row = 0
   Me.Grid2.Sort = 5 '字串昇冪
   
   iTotRows = Grid2.Rows - 1
   For dblFCnt = 1 To iTotRows
      Text2.Width = dblMaxWidth / iTotRows * dblFCnt: DoEvents
      strErr = "": strCP02 = "": strCP03 = "": strCP04 = ""
      strFileName = UCase(Grid2.TextMatrix(dblFCnt, 0))
      'Add by Amy 2020/01/08
      '專利固定抓 P
      If m_ProState = MsgText(601) Then
        strCP01 = "P"
      '商標
      Else
        strCP01 = InputCaseGetSys(strFileName)
      End If
      'end 2020/01/08
      
      bolUploadDone = False
      'Modfiy by Amy 2020/01/08 原只有P用系統別固定,開放商標用改成抓變數strCP01
      If PUB_GetCaseNoFromFileName(strFileName, strCP01, strCP02, strCP03, strCP04, strErr) = False Then
         strErr = convForm(CheckStr(strFileName), 25) & strErr
         List1.AddItem UCase(strErr), 0: SetListScroll List1
      Else
         With Me.MSHFlexGrid1
         For ii = 1 To .Rows - 1
            bolChkOk = False 'Add by Amy 2020/02/20
            If .TextMatrix(ii, 0) <> "X" Then
               If GetValue(ii, "CNo") = strCP01 & strCP02 & strCP03 & strCP04 Then
                  If Val(GetValue(ii, "Qty")) >= Val(GetValue(ii, "lp02")) Then
                     strErr = "檔案數已超過！"
                     Exit For
                  Else
                     'Modified by Morgan 2014/5/2 CP02零開頭的案號前面的零要去掉
                     'stSaveName = "P" & strCP02 & IIf(strCP03 & strCP04 = "000", "", "-" & strCP03 & "-" & strCP04) & "." & Val(GetValue(ii, "cp10"))
                     'Modified by Morgan 2015/1/26 本所號的追加聯合碼改各自判斷 Ex.P123456-1,P123456-0-01
                     'Modify by Amy 2020/02/20 配合卷宗區檔名格式統一,改抓函數
                     'stSaveName = strCP01 & Val(strCP02) & IIf(strCP04 <> "00", "-" & strCP03 & "-" & strCP04, IIf(strCP03 <> "0", "-" & strCP03, "")) & "." & Val(GetValue(ii, "cp10"))
                     stSaveName = PUB_CaseNo2FileName(strCP01, strCP02, strCP03, strCP04) & "." & Val(GetValue(ii, "cp10"))
      'end 2020/01/08
                     If Val(GetValue(ii, "Qty")) = 0 Then
                        '只有1個檔案(右方案號只有一筆)
                        'Modify by Amy 2020/02/20 原:lp01-語法並未抓此欄位-bug,並加Val(GetValue(ii, "CaseQty")) = 1
                        'Modified by Morgan 2021/3/23 + 判斷檔名內只有1個"."
                        If Val(GetValue(ii, "lp02")) = 1 And Val(GetValue(ii, "CaseQty")) = 1 And InStr(strFileName, ".") = InStrRev(strFileName, ".") Then
                           stSaveName = stSaveName & ".pdf"
                           bolChkOk = True
                        '來函(檔名內只有1個".")
                        'Modify by Amy 2020/02/20 加Val(GetValue(ii, "CaseQty")) = 1
                        ElseIf InStr(strFileName, ".") = InStrRev(strFileName, ".") And Val(GetValue(ii, "CaseQty")) = 1 Then
                           stSaveName = stSaveName & ".pdf"
                           bolChkOk = True
                        '前案(檔名內不只一個".")
                        Else
                           'Modify by Amy 2020/02/27 案件名稱有-需更改檔名(表示有相關案,需以相關案的案件性質上傳 ex:同一天有 核准-變更/核准-續展)
                           If Val(GetValue(ii, "CaseQty")) > 1 And InStr(GetValue(ii, "案件性質"), "-") > 0 Then
                                If InStr(strFileName, "." & Val(GetValue(ii, "RCp10")) & ".") > 0 Then
                                    stSaveName = Replace(stSaveName & Mid(strFileName, InStr(strFileName, ".")), "." & Val(GetValue(ii, "RCp10")) & ".", ".")
                                    bolChkOk = True
                                End If
                           '當天案號只有一筆
                           ElseIf Val(GetValue(ii, "CaseQty")) = 1 Then
                                'Modified by Morgan 2014/4/24 保留原來檔名
                                'stSaveName = stSaveName & ".001.pdf"
                                stSaveName = stSaveName & Mid(strFileName, InStr(strFileName, "."))
                                'Add by Amy 2020/03/04 檔名已有案件性質不需再加
                                If InStr(strFileName, "." & Val(GetValue(ii, "cp10")) & ".") > 0 Then
                                    stSaveName = Replace(stSaveName, "." & Val(GetValue(ii, "cp10")) & ".", ".")
                                End If
                                bolChkOk = True
                           End If
                        End If
                     Else
                        'Modify by Amy 2020/02/27 案件名稱有-需更改檔名(表示有相關案,需以相關案的案件性質上傳 ex:同一天有 核准-變更/核准-續展)
                        If Val(GetValue(ii, "CaseQty")) > 1 And InStr(GetValue(ii, "案件性質"), "-") > 0 Then
                            If InStr(strFileName, "." & Val(GetValue(ii, "RCp10")) & ".") > 0 Then
                                stSaveName = Replace(stSaveName & Mid(strFileName, InStr(strFileName, ".")), "." & Val(GetValue(ii, "RCp10")) & ".", ".")
                                bolChkOk = True
                            End If
                        '當天案號只有一筆
                        ElseIf Val(GetValue(ii, "CaseQty")) = 1 Then
                            'Modified by Morgan 2014/4/24 保留原來檔名
                            'stSaveName = GetFileName(GetValue(ii, "cp09"), stSaveName)
                            stSaveName = stSaveName & Mid(strFileName, InStr(strFileName, "."))
                            'Add by Amy 2020/03/04 檔名已有案件性質不需再加
                            If InStr(strFileName, "." & Val(GetValue(ii, "cp10")) & ".") > 0 Then
                                stSaveName = Replace(stSaveName, "." & Val(GetValue(ii, "cp10")) & ".", ".")
                            End If
                            bolChkOk = True
                        End If
                     End If
                     
                     'Modify by Amy 2020/02/20 +if bolChkOk及檔名已有案件性質不需再加
                     If bolChkOk = True Then
                        strCP09 = GetValue(ii, "cp09")
                        'Added by Morgan 2014/6/17
                        '檢查檔名是否重複
                        If FileExist(strCP09, stSaveName) = True Then
                           strErr = "檔名重複！"
                           Exit For
                        End If
                        'end 2014/6/17
                        
                        If UploadPDF(txtPath & "\" & strFileName, strCP09, stSaveName, ii, strErr) = True Then
                           bolUploadDone = True
                           Kill txtPath & "\" & strFileName
                           Exit For 'Added by Morgan 2014/8/15
                        End If
                     End If
                     'end 2020/02/20
                  End If
               End If
            End If
         Next
         End With
         
         If bolUploadDone = False Then
            If strErr = "" Then
               strErr = "無匹配缺檔案件！"
            End If
            strErr = convForm(CheckStr(strFileName), 25) & strErr
            List1.AddItem UCase(strErr), 0: SetListScroll List1
         End If
      End If
   Next
   MsgBox "匯入完畢！"
   Call cmdQuery_Click
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function

Private Function FileExist(pRecNo As String, pFileName As String) As Boolean
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select 1 from casepaperpdf where cpp01='" & pRecNo & "' and upper(cpp02)=upper('" & pFileName & "')"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      FileExist = True
   End If
   Set rsQuery = Nothing
End Function

Private Function UploadPDF(pFullPath As String, pCRecNo As String, pSaveName As String, pRowID As Integer, ByRef pErrMsg As String) As Boolean
   Dim stSQL As String
   
'Removed by Morgan 2015/3/23 上傳檔案改呼叫共用函數(要改為FTP方式)
'   Dim iFileNo As Integer
'   Dim bytes() As Byte
'   Dim lngSize As Long '檔案大小
'   Dim Numblocks As Integer
'   Dim LeftOver As Long
'   Dim adoRst As New ADODB.Recordset
'   Dim idx3 As Integer
'   Const BlockSize = 500000

On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   
On Error GoTo ErrHndT

   Set oFile = oFileSys.GetFile(pFullPath)
   
'Modified by Morgan 2015/3/23 上傳檔案改呼叫共用函數(要改為FTP方式)
'   iFileNo = FreeFile
'   Open pFullPath For Binary Access Read As #iFileNo
'
'   lngSize = LOF(iFileNo)
'   stSQL = "select * from CasePaperPDF where rownum<1"
'   If adoRst.State <> adStateClosed Then adoRst.Close
'   With adoRst
'   .CursorLocation = adUseClient
'   .Open stSQL, cnnConnection, adOpenStatic, adLockOptimistic
'   .AddNew
'   .Fields("cpp01").Value = pCRecNo
'   .Fields("cpp02").Value = pSaveName
'   .Fields("cpp03").Value = lngSize
'   Numblocks = lngSize / BlockSize
'   LeftOver = lngSize Mod BlockSize
'
'   ReDim bytes(LeftOver)
'   Get #iFileNo, , bytes()
'   .Fields("cpp04").AppendChunk bytes()
'
'   ReDim bytes(BlockSize)
'   For idx3 = 1 To Numblocks
'       Get #iFileNo, , bytes()
'       .Fields("cpp04").AppendChunk bytes()
'   Next
'   Close #iFileNo
'   iFileNo = 0
'   .Fields("cpp08") = Format(oFile.DateLastModified, "YYYYMMDD")
'   .Fields("cpp09") = Format(oFile.DateLastModified, "HHMMSS")
'   .Fields("cpp10") = "Y"
'   .UPDATE
'   End With
   
   'Modify By Sindy 2015/5/14
   'SaveAttFile_PDF pCRecNo, pFullPath, pSaveName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), True, "4", , , True
   SaveAttFile_PDF pCRecNo, pFullPath, pSaveName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), True, , , True
   '2015/5/14 END
'end 2015/3/24
   
   SetValue pRowID, "Qty", Val(GetValue(pRowID, "Qty")) + 1
   If Val(GetValue(pRowID, "Qty")) = Val(GetValue(pRowID, "lp02")) Then
      UpdateLP03 pCRecNo, pRowID
   End If
   
   cnnConnection.CommitTrans
   UploadPDF = True
   Exit Function

ErrHndT:
   cnnConnection.RollbackTrans

ErrHnd:
   pErrMsg = Err.Description
   
'Removed by Morgan 2015/3/23 上傳檔案改呼叫共用函數(要改為FTP方式)
'   Set adoRst = Nothing
'   If iFileNo > 0 Then Close #iFileNo

End Function
Private Sub cmdok_Click(Index As Integer)
cmdState = Index '紀錄作用按鍵
PubShowNextData
Exit Sub
End Sub

Private Sub cmdPrint_Click()
   PUB_RestorePrinter cmbPrinter
   DoPrint
   PUB_RestorePrinter strPrinter
End Sub

Private Sub DoPrint()
   Dim i As Integer, j As Integer
   Dim strFontName As String
   
   '缺檔案件
   If Check1.Value = 1 Then
      iLine1 = 0
      For j = 1 To MSHFlexGrid1.Rows - 1
         Erase strTemp
         For i = 1 To 5
            strTemp(i) = MSHFlexGrid1.TextMatrix(j, i)
         Next i
         If iLine1 > 52 Or iLine1 = 0 Then
            If iLine1 > 0 Then Printer.NewPage: iLine1 = 0
            PrintTitle '列印表頭
         End If
         PrintDetail '列印明細
      Next j
      '匯入錯誤訊息
      If Check2.Value = 1 Then
         iLine1 = iLine1 + 2
         Printer.Font.Size = 16
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iLine1 * 300
         Printer.Print "匯入錯誤訊息："
         iLine1 = iLine1 + 2
         Printer.Font.Size = 12
         For j = List1.ListCount - 1 To 0 Step -1
            Erase strTemp
            strTemp(1) = List1.List(j)
            If iLine1 > 52 Then
               If iLine1 > 0 Then Printer.NewPage: iLine1 = 2
            End If
            strFontName = Printer.FontName
            Printer.FontName = "細明體"
            PrintDetail2 '列印明細
            Printer.FontName = strFontName
         Next j
      End If
      Printer.EndDoc
      
   '匯入錯誤訊息
   ElseIf Check2.Value = 1 Then
      iLine1 = 0
      For j = List1.ListCount - 1 To 0 Step -1
         For i = 1 To 1
            strTemp(i) = ""
         Next i
         strTemp(1) = List1.List(j)
         If iLine1 > 52 Or iLine1 = 0 Then
            If iLine1 > 0 Then Printer.NewPage: iLine1 = 0
            PrintTitle2 '列印表頭
         End If
         strFontName = Printer.FontName
         Printer.FontName = "細明體"
         PrintDetail2 '列印明細
         Printer.FontName = strFontName
      Next j
      Printer.EndDoc
   Else
      MsgBox "請勾選要列印的內容！", vbInformation
   End If
End Sub

Public Sub cmdQuery_Click()
   Screen.MousePointer = vbHourglass
   QueryData
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPath_Click()
   Dim lpIDList As Long
   Dim sBuffer As String
   Dim szTitle As String
   Dim tBrowseInfo As BrowseInfo
   szTitle = "請選擇文檔存放路徑"
   
   With tBrowseInfo
       .hwndOwner = Me.hWnd
       .lpszTitle = lstrcat(szTitle, "")
       .ulFlags = BIF_RETURNONLYFSDIRS
   End With
   
   lpIDList = SHBrowseForFolder(tBrowseInfo)
   
   If (lpIDList) Then
       sBuffer = Space(MAX_PATH)
       SHGetPathFromIDList lpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       SaveSetting "TAIE", "P", UCase(Me.Name) & "Dir", sBuffer
       txtPath.Text = sBuffer
   End If
   
End Sub

'Added by Morgan 2025/1/15
Private Sub Combo1_Click()
   If Combo1.Visible = False Then Exit Sub
   If Combo1.Tag <> Combo1 Then
      cmdQuery.Value = True
   End If
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cmbPrinter, strPrinter
   '讀取前次設定路徑
   txtPath.Text = GetSetting("TAIE", "P", UCase(Me.Name) & "Dir", "")
   'Added by Lydia 2024/07/22
   If InStr(txtPath, "\\SALE1\") > 0 Then
      txtPath = Replace(txtPath, "\\SALE1\", "\\" & strSale1Path & "\")
   End If
   If InStr(txtPath, "\\PAT1\") > 0 Then
      txtPath = Replace(txtPath, "\\PAT1\", "\\" & strPat1Path & "\")
   End If
   'end 2024/07/22
   If txtPath = "" Then
      txtPath = strDocImportPath
      'Add by Amy 2020/01/08 +商標
      If m_ProState = "T" Then
        'Modified by Lydia 2024/07/22 改成變數
        'txtPath = "\\SALE1\TM_OA_SCAN"
        txtPath = "\\" & strSale1Path & "\TM_OA_SCAN"
      End If
      'Modified by Morgan 2014/9/3
      'If Dir(txtPath & "\", vbDirectory) = "" Then
      'Modified by Morgan 2017/1/12
      'If oFileSys.FolderExists(txtPath) = False Then
      If PUB_ChkDir(txtPath) = False Then
         'Modify by Amy 2020/01/08 原:strDocImportPath
         MsgBox "預設文檔存放路徑 [ " & txtPath & " ] 不存在，請確認！", vbCritical
         txtPath = "C:\"
      End If
   End If
   '紀錄進度棒寬度
   dblMaxWidth = Text2.Width
   
   '更新已齊備來函
   UpdateLP03
   
   'Added by Morgan 2025/1/15
   If m_ProState = "" And strSrvDate(1) >= P業務區劃分啟用日 Then
      Combo1.Visible = True
      Label4.Visible = True
      Call SetPatentP12Combo(Combo1, "P", Label4)
   End If
   'end 2025/1/15
   
   '查詢缺檔來函
   cmdQuery.Value = True
   
   'Add By Sindy 2021/1/4 無缺檔 T-229147
   If m_ProState = "T" Then
     cmdOK(0).Visible = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.cmbPrinter.Text <> Me.cmbPrinter.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
   End If
   Set oFileSys = Nothing
   Set oFile = Nothing
   Set frm040112 = Nothing
End Sub

Private Sub QueryData()
   Dim ii As Integer
   Dim strCountCase As String 'Add by Amy 2020/02/20
   Dim stCon As String 'Added by Morgan 2025/1/15
   
   Combo1.Tag = "" 'Added by Morgan 2025/1/15
   
   lblTotal = 0
   lPrevRow = 0
   'Add by Amy 2020/01/08 +if 加入商標
   '專利
   If m_ProState = "" Then
         'Added by Morgan 2025/1/15
         If Combo1 <> "" Then
            stCon = " and c1.cp65='" & Left(Combo1, 5) & "'"
         End If
         'end 2025/1/15
      
        'Modified by Morgan 2015/1/7 排除通知申請日案號
        'Modified by Morgan 2016/6/6 +剔除1102,+只抓臺灣案
        'Modify by Amy 2020/02/20 +案號筆數,同一天同一案號筆數(商標可能有 同一案號不同案件性質需匯,若有未輸案件性質可能會歸錯)
        strCountCase = "Select CP01||CP02||CP03||CP04 QCNo ,Count(*) CaseQty From LetterProgress,CaseProgress,Patent " & _
                            "Where LP03=0 And LP01>'C' And CP09(+)=LP01 And CP10<>'1101' And CP10<>'1102' And PA09='000' " & _
                            "And PA01(+)=CP01 And PA02(+)=CP02 And PA03(+)=CP03 And PA04(+)=CP04 " & _
                            "Group by CP01||CP02||CP03||CP04 "
        'Modify by Amy 2020/02/27 +相關案之案件性質名稱及編號
        'Modified by Morgan 2021/3/19 排除 .info. (IDS報價會先上傳卷宗區)--玲玲
        strExc(0) = "select '' as V,c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04) as 本所案號" & _
           ",Decode(PA09,'000',CPM03,CPM04)||GetRelateCasePropertyName(c1.cp09, '1') as 案件性質,sqldatet(c1.cp05) as 收文日" & _
           ",sqldatet(c1.cp06) as 本所期限, lp02-nvl(Qty,0)||'/'||lp02 as 缺檔" & _
           ",c1.CP01,c1.CP02,c1.CP03,c1.CP04,c1.CP09,c1.CP10,lp02,nvl(Qty,0) Qty,c1.CP01||c1.CP02||c1.CP03||c1.CP04 CNo,CaseQty,c2.cp10 as RCp10" & _
           " From letterprogress, caseprogress c1, caseprogress c2, patent, casepropertymap" & _
           ",(select cpp01,count(*) Qty from letterprogress,casepaperpdf where lp03=0 and lp01>'C' and cpp01(+)=lp01 and instr(upper(cpp02),'.CUS.PDF')=0 and instr(upper(cpp02),'.INFO.')=0 and cpp10<>'D' AND SUBSTR(UPPER(CPP02),-4)='.PDF' group by cpp01)" & _
           ",(" & strCountCase & ")" & _
           " where lp03=0 and lp01>'C' and c1.cp09(+)=lp01 And c2.cp09(+)=c1.cp43 and c1.cp10<>'1101' and c1.cp10<>'1102' And c1.CP01||c1.CP02||c1.CP03||c1.CP04=QCNo(+) " & _
           " and pa01(+)=c1.cp01 and pa02(+)=c1.cp02 and pa03(+)=c1.cp03 and pa04(+)=c1.cp04 and pa09='000'" & stCon & _
           " and cpm01(+)=c1.cp01 and cpm02(+)=c1.cp10 and cpp01(+)=lp01 order by c1.cp05 desc,c1.cp09 desc"
    '內商
    ElseIf m_ProState = "T" Then
        'Modify by Amy 2020/02/20 +案號筆數,同一天同一案號筆數(商標可能有 同一案號不同案件性質需匯,若有未輸案件性質可能會歸錯)
        strCountCase = "Select CP01||CP02||CP03||CP04 QCNo ,Count(*) CaseQty From LetterProgress,CaseProgress,TradeMark,ServicePractice " & _
                                "Where LP03=0 And LP01>'C' And CP09(+)=LP01 And CP10<>'1101' And NVL(TM10,SP09)='000' " & _
                                "And TM01(+)=CP01 And TM02(+)=CP02 And TM03(+)=CP03 And TM04(+)=CP04 " & _
                                "And SP01(+)=CP01 And SP02(+)=CP02 And SP03(+)=CP03 And SP04(+)=CP04 " & _
                                "Group by CP01||CP02||CP03||CP04 "
        'Modify by Amy 2020/02/27 +相關案之案件性質名稱及編號
        strExc(0) = "select '' as V,c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04) as 本所案號" & _
           ",Decode(Nvl(tm10,sp09),'000',CPM03,CPM04)||GetRelateCasePropertyName(c1.cp09, '1') as 案件性質,sqldatet(c1.cp05) as 收文日" & _
           ",sqldatet(c1.cp06) as 本所期限, lp02-nvl(Qty,0)||'/'||lp02 as 缺檔" & _
           ",c1.CP01,c1.CP02,c1.CP03,c1.CP04,c1.CP09,c1.CP10,lp02,nvl(Qty,0) Qty,c1.CP01||c1.CP02||c1.CP03||c1.CP04 CNo,CaseQty,c2.cp10 as RCp10" & _
           " From letterprogress, caseprogress c1, caseprogress c2,TradeMark,ServicePractice, casepropertymap" & _
           ",(select cpp01,count(*) Qty from letterprogress,casepaperpdf where lp03=0 and lp01>'C' and cpp01(+)=lp01 and instr(upper(cpp02),'.CUS.PDF')=0 and cpp10<>'D' AND SUBSTR(UPPER(CPP02),-4)='.PDF' group by cpp01)" & _
           ",(" & strCountCase & ")" & _
           " where lp03=0 and lp01>'C' and c1.cp09(+)=lp01 And c2.cp09(+)=c1.cp43 and c1.cp10<>'1101' And c1.CP01||c1.CP02||c1.CP03||c1.CP04=QCNo(+)" & _
           " and tm01(+)=c1.cp01 and tm02(+)=c1.cp02 and tm03(+)=c1.cp03 and tm04(+)=c1.cp04 and Nvl(tm10,sp09)='000' " & _
           " and sp01(+)=c1.cp01 and sp02(+)=c1.cp02 and sp03(+)=c1.cp03 and sp04(+)=c1.cp04 " & _
           " and cpm01(+)=c1.cp01 and cpm02(+)=c1.cp10 and cpp01(+)=lp01 order by c1.cp05 desc,c1.cp09 desc"
    End If
    'end 2020/01/08
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   With MSHFlexGrid1
   .Visible = False
   .FixedCols = 0
   '若沒有資料時不可直接設定給 Grid 否則 MouseRow 會跑掉
   If intI = 1 Then
      Set .Recordset = RsTemp
      lblTotal = RsTemp.RecordCount
      SetGrid
      
      '預設選取第一筆
      lPrevRow = 1
      MSHFlexGrid1.row = lPrevRow
      ClickGrid MSHFlexGrid1
   Else
      SetGrid True
   End If
   .Visible = True
   End With
   
   Combo1.Tag = Combo1 'Added by Morgan 2025/1/15
End Sub

Private Sub SetGrid(Optional pReset As Boolean = False)
   Dim arrGrd1HeadWidth
   Dim iCol As Integer
   Dim iUbound As Integer
   
   arrGrd1HeadWidth = Array(250, 1140, 800, 825, 825, 600)
   iUbound = UBound(arrGrd1HeadWidth)
   
   With MSHFlexGrid1
   .Visible = False
   If pReset = True Then
      .Clear
      .Rows = 2
      '.RowHeight(1) = 0
   End If
   .FormatString = "V|本所案號|案件性質|收文日|本所期限|缺檔數"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGrd1HeadWidth(iCol)
         If iCol = 5 Then
            .ColAlignment(iCol) = flexAlignCenterCenter
         Else
            .ColAlignment(iCol) = flexAlignLeftCenter
         End If
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   .Visible = True
   End With
End Sub

Public Sub PubShowNextData()
   Dim iRow As Integer
   
   Me.Enabled = False
   Screen.MousePointer = vbHourglass
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      If Trim(.TextMatrix(iRow, 0)) = "V" Then
         Select Case cmdState
            'Added by Sindy 2021/1/4
            Case 0 '無缺檔
               If UpdateData(PUB_MGridGetValue(iRow, "cp09", MSHFlexGrid1), iRow) Then
                  Me.Enabled = True
                  QueryData
               End If
            'end 2021/1/4
            Case 13 '卷宗區
               frm100101_L.m_strKey = PUB_MGridGetValue(iRow, "cp09", MSHFlexGrid1) '總收文號
               frm100101_L.Hide
               frm100101_L.SetParent Me
               If frm100101_L.QueryData = True Then
                  frm100101_L.Show
                  Me.Hide
               End If
         End Select
         
         Exit For
      End If
   Next
   End With
   Screen.MousePointer = vbDefault
   Me.Enabled = True
End Sub

Private Function UpdateData(pCP09 As String, pRowID As Integer) As Boolean

On Error GoTo ErrHnd

   cnnConnection.BeginTrans

   strSql = "update letterprogress set lp02=0 where lp01='" & pCP09 & "'"
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   PUB_UpdateLP03 pCP09
   
   cnnConnection.CommitTrans
   UpdateData = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim nCol As Long, nRow As Long, lRow As Long
   'getGrdColRow MSHFlexGrid1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   With MSHFlexGrid1
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow > 0 Then
      If lPrevRow > 0 Then
         If lPrevRow <> nRow Then
            .row = lPrevRow
            ClickGrid MSHFlexGrid1
            .row = nRow
            ClickGrid MSHFlexGrid1
         End If
      Else
         .row = nRow
         ClickGrid MSHFlexGrid1
      End If
      lPrevRow = .row
   End If
   .Visible = True
   End With
End Sub

Private Sub ClickGrid(grdDataList As MSHFlexGrid)
   Dim iCol As Integer

   With grdDataList
   If .TextMatrix(grdDataList.row, 1) <> "" Then
      If .TextMatrix(.row, 0) = "V" Then
         .TextMatrix(.row, 0) = ""
         For iCol = .FixedCols To .Cols - 1
            .col = iCol
            .CellBackColor = .BackColor
          Next
      '已刪除資料標示為 X
      ElseIf .TextMatrix(.row, 0) = "" Then
         .TextMatrix(.row, 0) = "V"
         For iCol = .FixedCols To .Cols - 1
            .col = iCol
            .CellBackColor = &HFFC0C0
         Next
      End If
   End If
   End With
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 2500
PLeft(3) = 5000
PLeft(4) = 6500
PLeft(5) = 8000
End Sub

Sub PrintTitle()
GetPleft
iLine1 = 1

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("機關來函缺文檔清單") / 2)
Printer.CurrentY = iLine1 * 300
Printer.Print "機關來函缺文檔清單"

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

iLine1 = iLine1 + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = 900
Printer.Print "列印人員：" & strUserName
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
iLine1 = iLine1 + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 1200
Printer.Print "頁　　次：" & Printer.Page

iLine1 = 5
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine1 * 300
Printer.Print "本所案號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine1 * 300
Printer.Print "案件性質"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine1 * 300
Printer.Print "收文日"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine1 * 300
Printer.Print "本所期限"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine1 * 300
Printer.Print "缺檔數"

iLine1 = iLine1 + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine1 * 300
Printer.Print String(148, "-")
iLine1 = iLine1 + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
   For m_j = 1 To 5
      Printer.CurrentX = PLeft(m_j)
      Printer.CurrentY = iLine1 * 300
      Printer.Print strTemp(m_j)
   Next m_j
   iLine1 = iLine1 + 1
End Sub

Sub PrintTitle2()
GetPleft
iLine1 = 1

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("匯入錯誤訊息") / 2)
Printer.CurrentY = iLine1 * 300
Printer.Print "匯入錯誤訊息"

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

iLine1 = iLine1 + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = 900
Printer.Print "列印人員：" & strUserName
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
iLine1 = iLine1 + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 1200
Printer.Print "頁　　次：" & Printer.Page

iLine1 = 5
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine1 * 300
Printer.Print "錯誤訊息"

iLine1 = iLine1 + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine1 * 300
Printer.Print String(148, "-")
iLine1 = iLine1 + 1
End Sub

Sub PrintDetail2()
Dim m_j As Integer
   For m_j = 1 To 1
      Printer.CurrentX = PLeft(m_j)
      Printer.CurrentY = iLine1 * 300
      Printer.Print strTemp(m_j)
   Next m_j
   iLine1 = iLine1 + 1
End Sub

Private Sub SetListScroll(oList As ListBox)
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

Private Sub UpdateLP03(Optional pCRecNo As String, Optional pRowID As Integer)
   Dim stSQL As String, intR As Integer
   
   'Modified by Morgan 2014/9/11
   'stSQL = "UPDATE letterprogress SET lp03=" & strSrvDate(1) & ",lp05=decode(lp04,null," & strSrvDate(1) & ",lp05) WHERE lp03=0 " & IIf(pCRecNo <> "", " and lp01='" & pCRecNo & "'", "") & _
   '   " AND EXISTS(SELECT 1 FROM casepaperpdf WHERE cpp01=lp01 and instr(upper(cpp02),'.CUS.PDF')=0 and cpp10<>'D' AND SUBSTR(UPPER(CPP02),-4)='.PDF' HAVING COUNT(*)=lp02)"
   'cnnConnection.Execute stSQL, intR
   'If intR = 1 Then
   If PUB_UpdateLP03(pCRecNo) = True Then
   'end 2014/9/11
      If pRowID > 0 Then
         MSHFlexGrid1.TextMatrix(pRowID, 0) = "X"
         MSHFlexGrid1.RowHeight(pRowID) = 0
         lblTotal = Val(lblTotal) - 1
      End If
   End If
   
End Sub

'取不重複的檔名
Private Function GetFileName(pCRecNo As String, pFileName As String) As String
   Dim strGoodName As String, ii As Integer
   strSql = "select upper(cpp02) cpp02 from casepaperpdf where cpp01='" & pCRecNo & "' order by 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
   With RsTemp
   Do
      ii = ii + 1
      strGoodName = pFileName & "." & Format(ii, "000") & ".pdf"
      .MoveFirst
      .Find "cpp02 ='" & UCase(strGoodName) & "'"
      If .EOF Then
         GetFileName = strGoodName
         Exit Do
      End If
   Loop
   End With
   End If
   
End Function

Private Function GetValue(pRow As Integer, pFieldName As String) As String
   Dim ii As Integer
   With Me.MSHFlexGrid1
   For ii = 0 To .Cols - 1
      If UCase(.TextMatrix(0, ii)) = UCase(pFieldName) Then
         GetValue = .TextMatrix(pRow, ii)
         Exit For
      End If
   Next
   End With
End Function

Private Function SetValue(pRow As Integer, pFieldName As String, pValue As String) As Boolean
   Dim ii As Integer
   With Me.MSHFlexGrid1
   For ii = 0 To .Cols - 1
      If UCase(.TextMatrix(0, ii)) = UCase(pFieldName) Then
         .TextMatrix(pRow, ii) = pValue
         SetValue = True
         Exit Function
      End If
   Next
   End With
End Function

