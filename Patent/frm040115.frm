VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040115 
   BorderStyle     =   1  '單線固定
   Caption         =   "收據/回執整批匯入"
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
   Begin VB.CommandButton cmdQuery 
      Caption         =   "重整(&Q)"
      Height          =   345
      Left            =   7980
      TabIndex        =   20
      Top             =   1200
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "卷宗區"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   13
      Left            =   7140
      Style           =   1  '圖片外觀
      TabIndex        =   19
      Top             =   1200
      Width           =   820
   End
   Begin VB.CommandButton cmdUpdLP03 
      Caption         =   "無收據/回執"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   0
      Left            =   5920
      Style           =   1  '圖片外觀
      TabIndex        =   23
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Caption         =   "缺收據案件："
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
      Height          =   3684
      Left            =   4005
      TabIndex        =   9
      Top             =   1332
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
         TabIndex        =   10
         Top             =   0
         Width           =   705
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3408
         Left            =   48
         TabIndex        =   11
         Top             =   240
         Width           =   4776
         _ExtentX        =   8424
         _ExtentY        =   6011
         _Version        =   393216
         Cols            =   6
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "V|本所案號|案件性質|發文日|本所期限"
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
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
      Height          =   3684
      Left            =   15
      TabIndex        =   6
      Top             =   1332
      Width           =   3960
      Begin MSComDlg.CommonDialog cd1 
         Left            =   480
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
         Height          =   3288
         ItemData        =   "frm040115.frx":0000
         Left            =   60
         List            =   "frm040115.frx":0002
         TabIndex        =   8
         Top             =   264
         Width           =   3840
      End
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
         TabIndex        =   7
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.FileListBox File1 
      Height          =   432
      Left            =   1200
      TabIndex        =   22
      Top             =   2280
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   444
      Left            =   0
      TabIndex        =   13
      Top             =   4968
      Width           =   8895
      Begin VB.TextBox Text2 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00FF0000&
         Enabled         =   0   'False
         Height          =   300
         Left            =   24
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   120
         Width           =   8820
      End
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   1470
      TabIndex        =   12
      Text            =   "\\Pat1\Reciept_SCAN"
      Top             =   864
      Width           =   7065
   End
   Begin VB.CommandButton cmdPath 
      Height          =   330
      Left            =   8565
      Picture         =   "frm040115.frx":0004
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   864
      Width           =   350
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   276
      Left            =   780
      TabIndex        =   4
      Top             =   5424
      Width           =   5835
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   564
      Left            =   7770
      TabIndex        =   2
      Top             =   192
      Width           =   885
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "匯入(&T)"
      Height          =   564
      Left            =   5760
      TabIndex        =   1
      Top             =   192
      Width           =   885
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   564
      Left            =   6750
      TabIndex        =   0
      Top             =   192
      Width           =   885
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   555
      Left            =   4440
      TabIndex        =   21
      Top             =   0
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
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      Caption         =   "發文人員："
      Height          =   240
      Left            =   408
      TabIndex        =   26
      Top             =   552
      Visible         =   0   'False
      Width           =   948
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1464
      TabIndex        =   25
      Top             =   504
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "　　　　　　　2.非臺灣案只需輸入逗點前面數字"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   24
      Top             =   276
      Width           =   3912
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "收據存放路徑："
      Height          =   180
      Left            =   96
      TabIndex        =   18
      Top             =   924
      Width           =   1260
   End
   Begin VB.Label lblTotal 
      Height          =   180
      Left            =   7848
      TabIndex        =   17
      Top             =   5484
      Width           =   996
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "總筆數："
      Height          =   180
      Index           =   0
      Left            =   7080
      TabIndex        =   16
      Top             =   5484
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "印表機："
      Height          =   180
      Left            =   60
      TabIndex        =   15
      Top             =   5508
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收據檔名規則：1.申請案號.PDF（ex.012345678.PDF）"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   3
      Left            =   60
      TabIndex        =   3
      Top             =   72
      Width           =   4188
   End
End
Attribute VB_Name = "frm040115"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/11/30 Form2.0已修改 (無需修改)
'Created by Amy 2014/08/01
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
Public m_ProState As String 'Add By Sindy 2019/12/30 系統作業


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
   PUB_UpdateLP03 'Added by Morgan 2016/3/24
   ImportFile
End Sub

Private Function ImportFile() As Boolean
   Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, strCP09 As String, strErr As String
   Dim iTotRows As Integer
   Dim ii As Integer
   Dim dblFCnt As Double
   Dim stSaveName As String '存檔的檔名-本所案號.案件性質.RECEIPT.Pdf
   Dim strFileName As String '匯入的檔名-申請案號
   Dim bolUploadDone As Boolean
   Dim strFileCP10 As String, strCP10 As String, strChkCaseNo As String 'Add By Sindy 2020/2/12
   Dim intRepCnt As Integer 'Add By Sindy 2021/3/24
   
On Error GoTo ErrHnd
   
   If IsEmptyText(txtPath) = True Then
      MsgBox "請選擇文檔存放路徑！", vbOKOnly, "檢核資料"
      cmdPath.SetFocus
      Exit Function
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
      strErr = "": strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = "": strCP09 = ""
      strFileName = UCase(Grid2.TextMatrix(dblFCnt, 0))
      bolUploadDone = False
      
      'Add By Sindy 2020/2/12 檢查是否有輸入案件性質
      strFileCP10 = Left(strFileName, Len(strFileName) - 4)
      If InStr(strFileCP10, ".") > 0 Then
         strFileCP10 = Mid(strFileCP10, InStr(strFileCP10, ".") + 1)
         strChkCaseNo = Left(Left(strFileName, Len(strFileName) - 4), InStr(Left(strFileName, Len(strFileName) - 4), ".") - 1)
      Else
         strFileCP10 = ""
         strChkCaseNo = Left(strFileName, Len(strFileName) - 4)
      End If
      '2020/2/12 END
      
      'Add By Sindy 2020/2/12
      'If ChkPA11No(Left(strFileName, Len(strFileName) - 4), strCP01, strCP02, strCP03, strCP04, strErr) = False Then
      If ChkPA11No(strChkCaseNo, strCP01, strCP02, strCP03, strCP04, strErr) = False Then
      '2020/2/12 END
         strErr = convForm(CheckStr(strFileName), 25) & strErr
         List1.AddItem UCase(strErr), 0: SetListScroll List1
      Else
         With Me.MSHFlexGrid1
         For ii = 1 To .Rows - 1
            If .TextMatrix(ii, 0) <> "X" Then
               If GetValue(ii, "CNo") = strCP01 & strCP02 & strCP03 & strCP04 Then
                  strCP09 = GetValue(ii, "cp09")
                  strCP10 = GetValue(ii, "cp10") 'Add By Sindy 2020/2/12
                  
                  'Add By Sindy 2020/2/12 檢查是否有輸入案件性質
                  If strFileCP10 = "" Or strFileCP10 = strCP10 Then
                  '2020/2/12 END
                  
                  'Modify by Amy 2014/10/03
                  'If Val(GetValue(ii, "Qty")) >= Val(GetValue(ii, "lp02")) Then
                  '   strErr = "檔案數已超過！"
                  '   Exit For
                  'Else
                     'CP02零開頭的案號前面的零要去掉
                     'Modified by Morgan 2015/1/26 本所號的追加聯合碼改各自判斷 Ex.P123456-1,P123456-0-01
                     'Modified by Morgan 2016/5/27
                     'stSaveName = "P" & Val(strCP02) & IIf(strCP04 <> "00", "-" & strCP03 & "-" & strCP04, IIf(strCP03 <> "0", "-" & strCP03, "")) & "." & Val(GetValue(ii, "cp10")) & ".RECEIPT.pdf"
                     'Modify By Sindy 2021/3/24
                     'stSaveName = PUB_CaseNo2FileName(strCP01, strCP02, strCP03, strCP04) & "." & Val(GetValue(ii, "cp10")) & ".RECEIPT.pdf"
                     strSql = "select count(*) from casepaperpdf where cpp01='" & strCP09 & "' and instr(upper(cpp02),upper('.RECEIPT.pdf'))>0"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     intRepCnt = 0
                     If intI = 1 Then
                        intRepCnt = RsTemp.Fields(0)
                     End If
                     stSaveName = PUB_CaseNo2FileName(strCP01, strCP02, strCP03, strCP04) & "." & Val(GetValue(ii, "cp10")) & IIf(intRepCnt > 0, "." & intRepCnt + 1, "") & ".RECEIPT.pdf"
                     '2021/3/24 END
                     
                     '檢查檔名是否重複
                     If FileExist(strCP09, stSaveName) = True Then
                        strErr = "檔名重複！"
                        Exit For
                     End If
                     
                     If UploadPDF(txtPath & "\" & strFileName, strCP09, stSaveName, ii, strErr) = True Then
                        bolUploadDone = True
                        Kill txtPath & "\" & strFileName
                        Exit For '解 缺收據案件list 若有2筆以上不可再存
                     End If
                     
                  End If 'Add By Sindy 2020/2/12
                  
                  'End If
                  'end 2014/10/03
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

'Removed by Morgan 2015/3/24
'   Dim stSQL As String
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
   'Modified by Morgan 2016/11/10 原則上不需加判斷,因為失敗會觸發錯誤
   'SaveAttFile_PDF pCRecNo, pFullPath, pSaveName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), True, , , True
   If SaveAttFile_PDF(pCRecNo, pFullPath, pSaveName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), True, , , True) = False Then
      Err.Raise 999, , " 上傳失敗!!"
   End If
   '2015/5/14 END
'end 2015/3/23
   
   'Modify by Amy 2014/10/22
   'SetValue pRowID, "Qty", Val(GetValue(pRowID, "Qty")) + 1
   'If Val(GetValue(pRowID, "Qty")) = Val(GetValue(pRowID, "lp02")) Then
      UpdateLP03 pCRecNo, pRowID
   'End If
   
   cnnConnection.CommitTrans
   UploadPDF = True
   Exit Function

ErrHndT:
   cnnConnection.RollbackTrans

ErrHnd:
   pErrMsg = Err.Description
   
'Removed by Morgan 2015/3/24
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

Private Sub cmdUpdLP03_Click(Index As Integer)
    Dim iRow As Integer, intR As Integer
    Dim strCP09 As String, stSQL As String
    
    Me.Enabled = False
    With MSHFlexGrid1
        For iRow = 1 To .Rows - 1
            If Trim(.TextMatrix(iRow, 0)) = "V" Then
                Screen.MousePointer = vbHourglass
                'Modified by Morgan 2015/1/7 +確認,因為可能會按錯
                If MsgBox(.TextMatrix(iRow, 1) & " (" & .TextMatrix(iRow, 2) & ") 是否確定無收據/回執", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                  strCP09 = GetValue(iRow, "cp09")
                  '上齊備日,若自行判發上判發日並設定LP02=0,lp19=null
                  'Modified by Morgan 2016/6/6 考慮大陸案會有其他附件(altr)
                  'stSQL = "Update LetterProgress Set lp02=0,lp03=" & strSrvDate(1) & ",lp05=Decode(lp04,null," & strSrvDate(1) & ",lp05),lp19=null " & _
                              "Where lp01='" & strCP09 & "' "
                  'cnnConnection.Execute stSQL, intR
                  cnnConnection.BeginTrans
                  stSQL = "Update LetterProgress Set lp02=lp02-1,lp19=null " & _
                              "Where lp01='" & strCP09 & "' and lp19='Y'"
                  cnnConnection.Execute stSQL, intR
                  PUB_UpdateLP03 strCP09
                  cnnConnection.CommitTrans
                  'end 2016/6/6
                End If
                Exit For
            End If
        Next
    End With
    cmdQuery_Click
    Screen.MousePointer = vbDefault
    Me.Enabled = True
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
      'Modify By Sindy 2019/12/30
      If m_ProState = "T" Then
         'Modified by Lydia 2024/07/22 改成變數
         'txtPath = "\\SALE1\TM_Receipt_SCAN"
         txtPath = "\\" & strSale1Path & "\TM_Receipt_SCAN"
      Else
      '2019/12/30 END
         'Modified by Lydia 2024/07/22 改成變數
         'txtPath = "\\Pat1\Reciept_SCAN"
         txtPath = "\\" & strPat1Path & "\Reciept_SCAN"
      End If
      'Modified by Morgan 2017/1/12
      'If oFileSys.FolderExists(txtPath) = False Then
      If PUB_ChkDir(txtPath) = False Then
         MsgBox "預設收據存放路徑 [ " & txtPath & " ] 不存在，請確認！", vbCritical
         txtPath = "C:\"
      End If
   End If
   '紀錄進度棒寬度
   dblMaxWidth = Text2.Width
   '更新已齊備日(手動上傳)
   PUB_UpdateLP03
   
   'Added by Morgan 2025/1/15
   If m_ProState = "" And strSrvDate(1) >= P業務區劃分啟用日 Then
      Combo1.Visible = True
      Label4.Visible = True
      Call SetPatentP12Combo(Combo1, "P", Label4)
   End If
   'end 2025/1/15
   
   '查詢缺檔收據
   cmdQuery_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.cmbPrinter.Text <> Me.cmbPrinter.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
   End If
   Set oFileSys = Nothing
   Set oFile = Nothing
   Set frm040115 = Nothing
End Sub

Private Sub QueryData()
   Dim iRow As Integer, idx As Integer, iColCP09 As Integer, iColCP10 As Integer
   Dim stCon As String 'Added by Morgan 2025/1/15
   
   Combo1.Tag = "" 'Added by Morgan 2025/1/15
   
   lblTotal = 0
   lPrevRow = 0
   'Modify by Amy 2014/10/03 拿掉缺檔
'   strExc(0) = "Select '' as V,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號" & _
'      ",Decode(PA09,'000',CPM03,CPM04) as 案件性質,sqldatet(cp27) as 發文日" & _
'      ",sqldatet(cp06) as 本所期限, lp02-nvl(Qty,0)||'/'||lp02 as 缺檔" & _
'      ",CP01,CP02,CP03,CP04,CP09,CP10,lp02,nvl(Qty,0) Qty,CP01||CP02||CP03||CP04 CNo,Nvl(lp04,'') as lp04" & _
'      " From LetterProgress a, CaseProgress, Patent, CasePropertyMap" & _
'      ",(Select cpp01,count(*) Qty From LetterProgress,CasePaperPdf Where lp03=0 And cpp01(+)=lp01 And InStr(upper(cpp02),'.CUS.PDF')=0 And InStr(upper(cpp02),'.DAT.PDF')=0 And InStr(upper(cpp02),'.ORDER.PDF')=0 AND Substr(Upper(cpp02),-4)='.PDF' And cpp10<>'D' And lp19='Y' Group By cpp01)" & _
'      " Where lp03=0 And lp19='Y' And Not Exists (Select * From CasePaperPdf Where InStr(Upper(cpp02),'.RECEIPT.PDF')>0 And a.lp01=cpp01(+) And cpp10<>'D') And cp09(+)=lp01" & _
'      " And pa01(+)=cp01 And pa02(+)=cp02 And pa03(+)=cp03 And pa04(+)=cp04 And cpm01(+)=cp01 And cpm02(+)=cp10 And cpp01(+)=lp01 Order By cp27 Desc,cp09 Desc"
   'Modify By Sindy 2019/12/30
   'Modify By Sindy 2021/3/2 sqldatet(cp06) as 本所期限 => sqldatet(cp152) as 自動扣款日
   If m_ProState = "T" Then
      'Modify By Sindy 2021/3/24 有2張收據 T-228505 612.補充理由(經濟部,經濟部智慧財產局)
      '" And Not Exists (Select * From CasePaperPdf Where InStr(Upper(cpp02),'.RECEIPT.PDF')>0 And a.lp01=cpp01(+) And cpp10<>'D')"
      '=>
      '" AND Decode(cp130,NULL,0,Counting(cp130))>(SELECT count(*) FROM CasePaperPdf WHERE InStr(Upper(cpp02),'.RECEIPT.PDF')+InStr(Upper(cpp02),'.FIL.PDF')>0 AND cpp01=cp09 AND cpp10<>'D')"
      '=>取消
      'Modify By Sindy 2021/4/6 + And cp118||cp84<>'Y0' : 排除台灣電子送件分割子案 T-233039,T-233040
      strExc(0) = "Select '' as V,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號" & _
         ",Decode(TM10,'000',CPM03,CPM04) as 案件性質,sqldatet(cp27) as 發文日" & _
         ",sqldatet(cp152) as 自動扣款日" & _
         ",CP01,CP02,CP03,CP04,CP09,CP10,lp02,CP01||CP02||CP03||CP04 CNo,Nvl(lp04,'') as lp04,cp27,nvl(cp152,19221111) sort1,cp82" & _
         " From LetterProgress a, CaseProgress, Trademark, CasePropertyMap" & _
         " Where lp03=0 And lp19='Y' And cp118||cp84<>'Y0'" & _
         " And cp09(+)=lp01 and cp27>0" & _
         " And tm01=cp01 And tm02=cp02 And tm03=cp03 And tm04=cp04 And cpm01(+)=cp01 And cpm02(+)=cp10"
      'Modify By Sindy 2021/1/18
      strExc(0) = strExc(0) & " union " & _
                  "Select '' as V,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號" & _
         ",Decode(sp09,'000',CPM03,CPM04) as 案件性質,sqldatet(cp27) as 發文日" & _
         ",sqldatet(cp152) as 自動扣款日" & _
         ",CP01,CP02,CP03,CP04,CP09,CP10,lp02,CP01||CP02||CP03||CP04 CNo,Nvl(lp04,'') as lp04,cp27,nvl(cp152,19221111) sort1,cp82" & _
         " From LetterProgress a, CaseProgress, ServicePractice, CasePropertyMap" & _
         " Where lp03=0 And lp19='Y' And cp118||cp84<>'Y0'" & _
         " And cp09(+)=lp01 and cp27>0" & _
         " And sp01=cp01 And sp02=cp02 And sp03=cp03 And sp04=cp04 And cpm01(+)=cp01 And cpm02(+)=cp10"
      strExc(0) = strExc(0) & _
         " Order By sort1 asc,cp27 asc,cp82 asc,cp09 Desc"
      '2021/1/18 END
   Else
   
      'Added by Morgan 2025/1/15
      If Combo1 <> "" Then
         stCon = " and c1.cp83='" & Left(Combo1, 5) & "'"
      End If
      'end 2025/1/15
         
   '2019/12/30 END
      'Modified by Morgan 2020/3/23 +Srt:同日發文多張收據時紙本發文優先
      'Modify By Sindy 2021/3/2 sqldatet(cp06) as 本所期限 => sqldatet(cp152) as 自動扣款日
      'Modified by Morgan 2021/9/24 C類的自動扣款日及電子送件改抓相關收文號
      'Modified by Morgan 2025/4/21 有收據但未齊備的還是要列(因可能會有一個以上的主管機關)
      strExc(0) = "Select '' as V,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) as 本所案號" & _
         ",Decode(PA09,'000',CPM03,CPM04) as 案件性質,sqldatet(c1.cp27) as 發文日" & _
         ",sqldatet(decode(substr(c1.cp09,1,1),'C',c2.cp152,c1.cp152)) as 自動扣款日,decode(decode(substr(c1.cp09,1,1),'C',c2.cp118,c1.cp118),'',1,2) Srt" & _
         ",c1.CP01,c1.CP02,c1.CP03,c1.CP04,c1.CP09,c1.CP10,lp02,Pa01||Pa02||Pa03||Pa04 CNo,Nvl(lp04,'') as lp04" & _
         " From LetterProgress a, CaseProgress c1, Patent, CasePropertyMap, CaseProgress c2" & _
         " Where lp03=0 And lp19='Y' And c1.cp09(+)=lp01 and c1.cp27>0" & _
         " And pa01=c1.cp01 And pa02=c1.cp02 And pa03=c1.cp03 And pa04=c1.cp04 And cpm01(+)=c1.cp01 And cpm02(+)=c1.cp10 and c2.cp09(+)=c1.cp43" & stCon & " Order By c1.cp27 Desc,Srt asc,pa01 asc,pa02 asc,pa03 asc,pa04 asc,c1.cp09 Desc"
   End If
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
      
      'Added by Morgan 2016/6/6
      idx = PUB_MGridGetId("案件性質", MSHFlexGrid1)
      iColCP09 = PUB_MGridGetId("cp09", MSHFlexGrid1)
      iColCP10 = PUB_MGridGetId("cp10", MSHFlexGrid1)
      For iRow = 1 To .Rows - 1
         If .TextMatrix(iRow, iColCP10) = "1909" Then
            .TextMatrix(iRow, idx) = .TextMatrix(iRow, idx) & PUB_GetRelateCasePropertyName(.TextMatrix(iRow, iColCP09), "1")
         End If
      Next
      'end 2016/6/6
      
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
   '缺檔案數先不顯示
   arrGrd1HeadWidth = Array(250, 1140, 1400, 825, 825)
   iUbound = UBound(arrGrd1HeadWidth)
   
   With MSHFlexGrid1
   .Visible = False
   If pReset = True Then
      .Clear
      .Rows = 2
   End If
   .FormatString = "V|本所案號|案件性質|發文日|自動扣款日"
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
   
   Select Case cmdState
      Case 13 '卷宗區
         Me.Enabled = False
         With MSHFlexGrid1
         For iRow = 1 To .Rows - 1
            If Trim(.TextMatrix(iRow, 0)) = "V" Then
               Screen.MousePointer = vbHourglass
               frm100101_L.m_strKey = GetValue(iRow, "cp09") '總收文號
               frm100101_L.Hide
               frm100101_L.SetParent Me
               If frm100101_L.QueryData = True Then
                  frm100101_L.Show
                  Me.Hide
               End If
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
            End If
         Next
         End With
         Me.Enabled = True
   End Select
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim nCol As Long, nRow As Long, lRow As Long
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

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("收據缺文檔清單") / 2)
Printer.CurrentY = iLine1 * 300
Printer.Print "收據缺文檔清單"

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
Printer.Print "發文日"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine1 * 300
Printer.Print "自動扣款日"
'Mark by Amy 2014/10/03
'Printer.CurrentX = PLeft(5)
'Printer.CurrentY = iLine1 * 300
'Printer.Print "缺檔數"

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
'   Dim stSQL As String, intR As Integer
   
'   '要剔除客戶函、申請書、接洽單
'   stSQL = "Update LetterProgress Set lp03=" & strSrvDate(1) & ",lp05=Decode(lp04,null," & strSrvDate(1) & ",lp05) Where lp03=0 " & IIf(pCRecNo <> "", " And  lp01='" & pCRecNo & "'", "") & _
'      " And  Exists(Select 1 From CasePaperPdf Where cpp01=lp01 And InStr(upper(cpp02),'.CUS.PDF')=0 And InStr(upper(cpp02),'.DAT.PDF')=0 And InStr(upper(cpp02),'.ORDER.PDF')=0 AND Substr(Upper(cpp02),-4)='.PDF' And cpp10<>'D' Having Count(*)=lp02)"
'   cnnConnection.Execute stSQL, intR
   
   If PUB_UpdateLP03(pCRecNo) = True Then
      If pRowID > 0 Then
         MSHFlexGrid1.TextMatrix(pRowID, 0) = "X"
         MSHFlexGrid1.RowHeight(pRowID) = 0
         lblTotal = Val(lblTotal) - 1
      End If
   End If
   
End Sub

Private Function GetValue(pRow As Integer, pCaseNo As String) As String
   Dim ii As Integer
   With Me.MSHFlexGrid1
   For ii = 0 To .Cols - 1
      If UCase(.TextMatrix(0, ii)) = UCase(pCaseNo) Then
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

'確認申請案號
Private Function ChkPA11No(pPA11 As String, pPA01 As String, pPA02 As String, pPA03 As String, pPA04 As String, strErr As String) As Boolean
   Dim stSQL As String, adoRst As ADODB.Recordset, intR As Integer
   Dim strChkOK As String
   
   'Add By Sindy 2021/1/18 檢查是否為商標系統別
   strChkOK = ""
   If m_ProState = "T" Then
      stSQL = "SELECT * FROM systemkind WHERE sk02 IN('2','6')" & _
              " and substr('" & pPA11 & "',1,length(sk01))=sk01"
      intR = 1
      Set adoRst = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         strChkOK = "案號"
      End If
   End If
   '2021/1/18 END
   'If Len(pPA11) = 9 Or Len(pPA11) = 12 Then
   'Modify by Amy 2025/07/11 +Len(pPA11) = 13,P-111831 申請案號[201510395519.9],但代理人來的檔案為[2015103955199] 造成無法匯入-玲玲
   If ((Len(pPA11) = 9 Or Len(pPA11) = 12 Or Len(pPA11) = 13) And m_ProState <> "T") Or _
      (strChkOK = "" And m_ProState = "T") Then
      'Modify By Sindy 2019/12/30
      If m_ProState = "T" Then
         'Add By Sindy 2021/1/18 檢查是否為申請案號或審定號數或對造號數
         '審定號數:
         stSQL = "Select tm01 pa01,tm02 pa02,tm03 pa03,tm04 pa04,tm10 pa09" & _
                 " From trademark Where tm15='" & pPA11 & "'" & _
                 " And tm29||tm57 is null" & _
                 " union Select sp01 pa01,sp02 pa02,sp03 pa03,sp04 pa04,sp09 pa09" & _
                 " From servicepractice Where sp13='" & pPA11 & "'" & _
                 " And sp15||sp61 is null"
         intR = 1
         Set adoRst = ClsLawReadRstMsg(intR, stSQL)
         If intR = 0 Then
            '申請案號:
            stSQL = "Select tm01 pa01,tm02 pa02,tm03 pa03,tm04 pa04,tm10 pa09" & _
                    " From trademark Where tm12='" & pPA11 & "'" & _
                    " And tm29||tm57 is null" & _
                    " union Select sp01 pa01,sp02 pa02,sp03 pa03,sp04 pa04,sp09 pa09" & _
                    " From servicepractice Where sp11='" & pPA11 & "'" & _
                    " And sp15||sp61 is null"
            intR = 1
            Set adoRst = ClsLawReadRstMsg(intR, stSQL)
            If intR = 0 Then
               '對造號數:
               stSQL = "Select cp01 pa01,cp02 pa02,cp03 pa03,cp04 pa04" & _
                       " From caseprogress,trademark Where cp36='" & pPA11 & "'" & _
                       " And cp01=tm01(+) And cp02=tm02(+) And cp03=tm03(+) And cp04=tm04(+)" & _
                       " And tm29||tm57 is null And tm28<>'1' group by cp01,cp02,cp03,cp04"
               strChkOK = "對造號數"
            Else
               strChkOK = "申請案號"
            End If
         Else
            strChkOK = "審定號數"
         End If
      Else
         strChkOK = "申請案號"
      '2019/12/30 END
         stSQL = "Select pa01,pa02,pa03,pa04,pa09 " & _
             "From Patent Where  pa11='" & pPA11 & "' and pa09='000'" & _
             "And pa57||pa108 is null "
         'Added by Morgan 2016/6/1 +大陸案(只輸前面12碼)
         'Modify by Amy 2025/07/11 將大陸案語法拆成另一句,避免用Replace 取代[.],造成速度慢-Morgan
         '  ex:P-111831 申請案號[201510395519.9],但代理人來的檔案為[2015103955199] 造成無法匯入-玲玲
'         stSQL = stSQL & " union all Select pa01,pa02,pa03,pa04,pa09 " & _
'             "From Patent Where  pa11>='" & pPA11 & "' and pa11<='" & pPA11 & "Z' and pa09<>'000'" & _
'             "And pa57||pa108 is null "
         stSQL = stSQL & " union all Select pa01,pa02,pa03,pa04,pa09 " & _
             "From Patent Where  pa11>='" & Left(pPA11, 12) & "' and pa11<='" & Left(pPA11, 12) & "Z' and pa09='020' " & _
             "And pa57||pa108 is null "
         stSQL = stSQL & " union all Select pa01,pa02,pa03,pa04,pa09 " & _
             "From Patent Where  pa11>='" & pPA11 & "' and pa11<='" & pPA11 & "Z' and pa09<>'000' and pa09<>'020' " & _
             "And pa57||pa108 is null "
         'end 2016/6/1
      End If
      intR = 1
      Set adoRst = ClsLawReadRstMsg(intR, stSQL)
      If intR = 0 Then
         'Modify By Sindy 2021/1/18
         'strErr = "無對應的申請案號"
         strErr = "無對應的" & strChkOK
         '2021/1/18 END
         
      ElseIf intR = 1 Then
          If adoRst.RecordCount = 1 Then
             'Modify By Sindy 2019/12/30
             If m_ProState = "T" Then
                If Left(adoRst.Fields("pa01"), 1) = "T" Or adoRst.Fields("pa01") = "FCT" Then
                   ChkPA11No = True
                   pPA01 = adoRst.Fields("pa01")
                   pPA02 = adoRst.Fields("pa02")
                   pPA03 = adoRst.Fields("pa03")
                   pPA04 = adoRst.Fields("pa04")
                Else
                   strErr = "系統別錯誤"
                End If
             Else
             '2019/12/30 END
                If adoRst.Fields("pa01") = "P" Then
                    'Modified by Morgan 2016/6/1 非臺灣案也可匯入收據
                    'If adoRst.Fields("pa09") = "000" Then
                        ChkPA11No = True
                        pPA01 = adoRst.Fields("pa01")
                        pPA02 = adoRst.Fields("pa02")
                        pPA03 = adoRst.Fields("pa03")
                        pPA04 = adoRst.Fields("pa04")
                        'pPA09 = adoRst.Fields("pa09") 'Removed by Morgan 2016/7/6 沒用
                        'Exit Function 'Removed by Morgan 2016/7/6
                    'Else
                    '    strErr = "申請國家非台灣"
                    '    Exit Function
                    'End If
                    'end 2016/6/1
                Else
                    strErr = "系統別錯誤"
                    'Exit Function 'Removed by Morgan 2016/7/6
                End If
             End If
          Else
             '多筆無法確定歸哪筆,讓user 手動歸
             'Modify By Sindy 2021/1/18
             'strErr = "有" & adoRst.RecordCount & " 個本所案號的申請案號相同"
             strErr = "有" & adoRst.RecordCount & " 個本所案號的" & strChkOK & "相同"
             '2021/1/18 END
             'Exit Function 'Removed by Morgan 2016/7/6
          End If
      End If
      
   Else
      'Modify By Sindy 2019/12/30
      If m_ProState <> "T" Then
      '2019/12/30 END
         'Modified by Morgan 2016/7/6 開放也可以輸本所案號
         pPA01 = "P"
      End If
      If PUB_GetCaseNoFromFileName(pPA11, pPA01, pPA02, pPA03, pPA04, strErr) = True Then
         ChkPA11No = True
      Else
        strErr = "檔名錯誤"
      End If
   End If
    
'Added by Morgan 2016/7/6
   Set adoRst = Nothing
End Function
