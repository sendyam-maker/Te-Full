VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc2154 
   AutoRedraw      =   -1  'True
   Caption         =   "帳單/抵帳單電子檔"
   ClientHeight    =   2844
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5148
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2844
   ScaleWidth      =   5148
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4095
      TabIndex        =   7
      Top             =   630
      Width           =   900
   End
   Begin VB.CommandButton cmdSaveAtt 
      Caption         =   "下載"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   810
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   630
      Width           =   615
   End
   Begin VB.CommandButton cmdRemAtt 
      Caption         =   "刪除"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2070
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   630
      Width           =   615
   End
   Begin VB.CommandButton cmdAddAtt 
      Caption         =   "新增"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   630
      Width           =   615
   End
   Begin VB.TextBox txtAYF01 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1950
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   1
      Top             =   165
      Width           =   1395
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   240
      Top             =   5310
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   572
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Frmacc2154.frx":0000
      Height          =   1140
      Left            =   240
      TabIndex        =   0
      Top             =   5460
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   2011
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   20
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "axf03"
         Caption         =   "本所案號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "axf02"
         Caption         =   "總收文號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "cpm03"
         Caption         =   "案件性質"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "axf04"
         Caption         =   "帳單金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "axf14"
         Caption         =   "盈虧"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "ProcessProfit"
         Caption         =   "收文號盈虧 "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "axf12"
         Caption         =   "案件名稱"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "axf13"
         Caption         =   "收據抬頭"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         Size            =   284
         BeginProperty Column00 
            ColumnWidth     =   1404.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1284.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1332.284
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1404.284
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   3636.284
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   4356.284
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1695
      Left            =   135
      TabIndex        =   3
      Top             =   990
      Width           =   4890
      _ExtentX        =   8615
      _ExtentY        =   2985
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V| 檔案名稱| 最後修改時間|上傳時間"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.CommandButton cmdOpenAtt 
      Caption         =   "開啟"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   630
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "帳單/抵帳單編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   225
      TabIndex        =   2
      Top             =   210
      Width           =   1650
   End
End
Attribute VB_Name = "Frmacc2154"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/07 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB
'Created by Morgan 2016/6/27
Option Explicit

Dim m_SaveFolder As String
Dim m_AttachPath As String

Private Sub cmdAddAtt_Click()
   Dim stFileName As String
   Dim sFile() As String
   Dim ii As Integer
   Dim fs, s
   Dim f As File
   Dim bolAdd As Boolean
   
On Error GoTo ErrHnd

   'Added by Morgan 2024/3/21
   'Y55766德國專利局帳單付款前不可有電子檔
   strExc(0) = "select * from acc150 where a1501='" & txtAYF01 & "' and a1503='Y55766000' and nvl(a1520,0)=0"
   intI = 1
   Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      MsgBox "德國專利局帳單結匯前不應有電子檔！", vbExclamation
      Exit Sub
   End If
   'end 2024/3/21
         

   stFileName = "*.*"
   With CommonDialog1
   .CancelError = True
   .FileName = stFileName
   .Filter = "All Files (*.*)|*.*"
   If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "") <> "" Then
      .InitDir = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
   Else
      .InitDir = PUB_Getdesktop
   End If
   .MaxFileSize = 3000
   .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
   .ShowOpen
   If .FileName <> "" Then
      If InStr(.FileName, ChrW$(0)) > 0 Then
         sFile = Split(.FileName, ChrW$(0))
      Else
         ReDim sFile(1) As String
         sFile(0) = Left(.FileName, InStrRev(.FileName, "\") - 1)
         sFile(1) = Mid(.FileName, InStrRev(.FileName, "\") + 1)
      End If
      
      '記錄路徑
      SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", sFile(0)
      For ii = 1 To UBound(sFile)
         If InStr(sFile(ii), "\") > 0 Then
            stFileName = sFile(ii)
         Else
            stFileName = sFile(0) & "\" & sFile(ii)
         End If
         
         If Right(Trim(UCase(stFileName)), 4) <> ".PDF" Then
            MsgBox "格式不符,只可存放.PDF檔!!"
            GoTo EXITSUB
         End If
         Set fs = CreateObject("Scripting.FileSystemObject")
         Set f = fs.GetFile(stFileName)
         If f.Size = 0 Then
            ShowMsg sFile(ii) & MsgText(9221)
         ElseIf f.Size > 5242880 Then
            If MsgBox("檔案過大（容量超過5MB），確認是否要傳送？", vbYesNo, "警告") = vbNo Then
               GoTo EXITSUB
            End If
         End If
         If IsRecordExist(sFile(ii)) = False Then
            '存檔並刪除原檔
            If AddRecord(sFile(ii), stFileName, f) = True Then
               bolAdd = True
            Else
               GoTo EXITSUB
            End If
         End If
      Next ii
   End If
   End With
   
   
EXITSUB:
   If bolAdd Then
      OpenTable
   End If
   Exit Sub
   
ErrHnd:
   If Err.NUMBER <> 32755 Then
      MsgBox Err.Description
   End If
               
End Sub

Private Function AddRecord(pFileName As String, pFullFromPath As String, pFile As File) As Boolean
   Dim stSQL As String, iRecords As Integer
   Dim bolInTrans As Boolean
   Dim stFtpPath As String

On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   stSQL = "update ACC152 set ayf01=ayf01 where ayf01='" & txtAYF01 & "' and upper(ayf02)='" & ChgSQL(UCase(pFileName)) & "'"
   cnnConnection.Execute stSQL, iRecords
   If iRecords > 0 Then
      Err.Raise 999, , "檔名 " & pFileName & " 重複!!"
   End If
   
   If PUB_PutFtpFile(pFullFromPath, strSrvDate(1), pFileName, stFtpPath, "ACC152") Then
      stSQL = "insert into ACC152(ayf01,ayf02,ayf03,ayf04,ayf05,ayf06,ayf07,ayf08,ayf09) values('" & txtAYF01 & "','" & ChgSQL(pFileName) & "'," & pFile.Size & "," & Format(pFile.DateLastModified, "YYYYMMDD") & "," & Format(pFile.DateLastModified, "HHMMSS") & ",'" & ChgSQL(stFtpPath) & "','" & strUserNum & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'))"
      cnnConnection.Execute stSQL, iRecords
   Else
      Err.Raise 999, , " 檔案 " & pFileName & " 上傳失敗!!"
   End If

   cnnConnection.CommitTrans
   pFile.Delete
   AddRecord = True
   
ErrHand:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Function

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal stFileName As String) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim adoRst As ADODB.Recordset
   
   IsRecordExist = False
   
   stSQL = "SELECT ayf01,ayf02 FROM acc152 WHERE ayf01='" & txtAYF01 & "' and upper(ayf02)=upper('" & ChgSQL(stFileName) & "')"
   intQ = 1
   Set adoRst = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      IsRecordExist = True
      MsgBox "檔案 " & stFileName & " 已存在！", vbCritical
   End If
   
   Set adoRst = Nothing
End Function

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdOpenAtt_Click()
   Dim ii As Integer
   Dim bolCheck As Boolean
   
   With MSHFlexGrid1
   For ii = 1 To .Rows - 1
      If UCase(UCase(.TextMatrix(ii, 0))) = "V" Then
         bolCheck = True
         OpenAtt .TextMatrix(ii, 3)
      End If
   Next
   End With
   If Not bolCheck Then MsgBox "請點選要開啟的檔案", vbInformation
End Sub

Private Sub OpenAtt(pFileName As String)
   Dim stSaveFileName As String
   Dim hLocalFile As Long
   
   If PUB_GetAttachFile_Invoice(txtAYF01, pFileName, m_AttachPath, stSaveFileName) = True Then
      ShellExecute hLocalFile, "open", m_AttachPath & "\" & stSaveFileName, vbNullString, vbNullString, 1
   End If
End Sub

Private Sub cmdRemAtt_Click()
   Dim iRecord As Integer
   Dim ii As Integer, bolCheck As Boolean
   
On Error GoTo ErrHnd

   With MSHFlexGrid1
   bolCheck = False
   For ii = 1 To .Rows - 1
      If UCase(.TextMatrix(ii, 0)) = "V" Then
         bolCheck = True
         Exit For
      End If
   Next
   If Not bolCheck Then
      MsgBox "請先勾選要刪除的記錄！", vbExclamation: Exit Sub
   Else
      If MsgBox("是否確定要刪除？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then Exit Sub
   End If
   
   For ii = 1 To .Rows - 1
      If UCase(.TextMatrix(ii, 0)) = "V" Then
          '此處不包transaction 若FTP刪除成功但DB刪除失敗才會有log
          strSql = "delete acc152 where ayf01='" & txtAYF01 & "' and ayf02='" & ChgSQL(.TextMatrix(ii, 3)) & "'"
          Pub_SeekTbLog strSql
         If PUB_DelFtpFile2(txtAYF01, " and ayf02='" & ChgSQL(.TextMatrix(ii, 3)) & "'", "ACC152") Then
            cnnConnection.Execute strSql, iRecord
         Else
            Err.Raise 999, , " 刪除Ftp檔案失敗!!"
         End If
         .TextMatrix(ii, 0) = "X"
         .RowHeight(ii) = 0
      End If
   Next
   End With
   OpenTable
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Sub

Private Sub cmdSaveAtt_Click()
   Dim stFileName As String, stFolderPath As String, stFullName As String
   Dim bMultiFile As Boolean
   Dim ii As Integer
   
   Screen.MousePointer = vbHourglass
   
   stFileName = ""
   bMultiFile = False
   With MSHFlexGrid1
   For ii = 1 To .Rows - 1
      If UCase(.TextMatrix(ii, 0)) = "V" And Trim(MSHFlexGrid1.TextMatrix(ii, 3)) <> "" Then
         If stFileName <> "" Then
            bMultiFile = True
            Exit For
         Else
            stFileName = Trim(.TextMatrix(ii, 3))
         End If
      End If
   Next ii
   End With
   
   If stFileName = "" Then
      MsgBox "請勾選要下載的檔案！"
   Else
      '多選
      If bMultiFile Then
         If m_SaveFolder = "" Then m_SaveFolder = PUB_Getdesktop
         stFolderPath = PUB_GetFolder(Me.hWnd, m_SaveFolder, "請選擇欲儲存的位置:")
         If stFolderPath <> "" Then
            With MSHFlexGrid1
            For ii = 1 To .Rows - 1
               If UCase(.TextMatrix(ii, 0)) = "V" And Trim(MSHFlexGrid1.TextMatrix(ii, 1)) <> "" Then
                  stFileName = Trim(.TextMatrix(ii, 3))
                  stFullName = stFolderPath & "\" & stFileName
                  If stFullName <> "" Then
                     If Dir(stFullName) <> "" Then
                        If MsgBox("檔案[ " & Trim(.TextMatrix(ii, 1)) & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                           stFullName = ""
                        End If
                     End If
                     If stFullName <> "" Then
                        If PUB_GetAttachFile_Invoice(txtAYF01, stFileName, stFullName, stFileName) = False Then
                           MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                           GoTo RunExit
                        End If
                     End If
                  End If
               End If
            Next ii
            End With
         End If
      Else
         stFullName = GetSaveName(stFileName)
         If stFullName <> "" Then
            If Dir(stFullName) <> "" Then
               If MsgBox("檔案[ " & stFullName & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                  stFullName = ""
               End If
            End If
            If stFullName <> "" Then
               If PUB_GetAttachFile_Invoice(txtAYF01, stFileName, stFullName, stFileName) = False Then
                  MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                  GoTo RunExit
               End If
            End If
         End If
      End If
      If stFullName <> "" Then
         MsgBox "下載完成！"
      End If
   End If
RunExit:
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
   If Err.NUMBER <> 32755 Then
      MsgBox Err.Description
   End If
End Function

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height
   txtAYF01 = strItemNo
   OpenTable
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   KillTemp
End Sub

Private Sub KillTemp()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
End Sub

Private Sub OpenTable()
   
   'Memoed by Morgan 2016/6/30 一張帳單只能有一個檔案--郭雅娟
   strExc(0) = "select '' as V, ayf02||' ('||Round(ayf03 / 1024, 2)||' KB)' as 檔案名稱 " & _
      ", sqldatet(ayf08)||' '||sqltime(ayf09)||'('||st02||')' as 上傳時間,ayf02" & _
      " from acc152,staff where ayf01='" & txtAYF01 & "' and st01(+)=ayf07"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Set MSHFlexGrid1.Recordset = RsTemp
   SetGrid
   If intI = 1 Then
      cmdAddAtt.Enabled = False
      cmdOpenAtt.Enabled = True
      cmdSaveAtt.Enabled = True
      If Not (strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4)) Then
         cmdRemAtt.Enabled = True
      Else
         cmdRemAtt.Enabled = False
      End If
   Else
      If Not (strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4)) Then
         cmdAddAtt.Enabled = True
      Else
         cmdAddAtt.Enabled = False
      End If
      cmdRemAtt.Enabled = False
      cmdOpenAtt.Enabled = False
      cmdSaveAtt.Enabled = False
   End If
End Sub

Private Sub SetGrid(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGridHeadWidth
   Dim iUbound As Integer

   'arrGridHeadWidth = Array(240, 2300, 1480)
   arrGridHeadWidth = Array(240, 2000, 2300)
   iUbound = UBound(arrGridHeadWidth)
   
   With MSHFlexGrid1
   If pReset = True Then
      .Clear
      .Rows = 2
      .FormatString = "V|檔案名稱|最後修改時間|上傳時間"
   End If
   '.FixedCols = 2
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         .ColAlignment(iCol) = flexAlignLeftCenter
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   End With
End Sub

Private Sub MSHFlexGrid1_Click()
   Dim intRow As Integer
   If MSHFlexGrid1.row > 0 Then
      intRow = MSHFlexGrid1.row
      GridClick MSHFlexGrid1, intRow, 0
   End If
End Sub

Private Sub MSHFlexGrid1_DblClick()
   If MSHFlexGrid1.row > 0 Then
      If cmdOpenAtt.Enabled Then OpenAtt MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 3)
   End If
End Sub
