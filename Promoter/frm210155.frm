VERSION 5.00
Begin VB.Form frm210155 
   BorderStyle     =   1  '單線固定
   Caption         =   "CFP領證預估報價"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   5040
   Begin VB.CommandButton cmdSave 
      Caption         =   "更換PDF檔案"
      Default         =   -1  'True
      Height          =   400
      Left            =   1995
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   120
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   4200
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "開啟"
      Height          =   400
      Left            =   3405
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   120
      Width           =   800
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "版本："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "僅供參考，非實際報價。"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "frm210155"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2022/05/11 CFP領證預估報價
Option Explicit

' 變數宣告區
Dim m_AttachPath As String
Private Declare Function SendMessageByNum Lib "user32" _
   Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
   wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194

Public m_strSaveFiles As String '新增附件

Private Const strKEY01 = "14"
Dim strKEY02 As String
    
Private Sub cmdOpen_Click()
Dim hLocalFile As Long
Dim stFileName As String

Dim strTemp As String
   
   Screen.MousePointer = vbHourglass

   strKEY02 = Mid(Label3.Caption, InStr(Label3.Caption, "(") + 1, 9)
   stFileName = "CFP領證預估報價"
    If InStrRev(strTemp, "-") > 0 Then
       stFileName = stFileName & Mid(strTemp, InStrRev(strTemp, "-"))
    End If
    stFileName = "$$" & stFileName & ServerTime & ".pdf"
    If GetAttachFile(stFileName, strKEY01, strKEY02) = False Then
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
   
   Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSave_Click()

   Call frm090801_8.SetParent(Me)
   frm090801_8.m_strSaveFiles = Me.m_strSaveFiles
   frm090801_8.Label1.Visible = False
   frm090801_8.lblCaseNo.Visible = False

   frm090801_8.Show vbModal
   
   If m_strSaveFiles <> "" Then
      If InStr(CStr(m_strSaveFiles), "&") > 0 Then
         MsgBox "附件只能有一個!", vbCritical
         Exit Sub
      End If
      Screen.MousePointer = vbHourglass
      strKEY02 = strSrvDate(1)
      If SaveAttFile(strKEY01, strKEY02, m_strSaveFiles) = False Then
          Screen.MousePointer = vbDefault
          Exit Sub
      Else
          MsgBox "上傳完成!", vbInformation
          cmdOpen.Enabled = True
          Label3.Caption = ChangeTStringToTDateString(strSrvDate(2))
      End If
      Screen.MousePointer = vbDefault
   End If
   
End Sub

Private Sub Form_Load()
Dim bolUpd As Boolean

   MoveFormToCenter Me
   m_AttachPath = App.path
   '更換PDF檔案的權限
   bolUpd = False
   strSql = "select distinct decode(st01,null,SR01,st01) from staff_right,staff" & _
            " where upper(sr02)='" & UCase(Me.Name) & "' and sr03='Y' and sr04='Y' and sr05='Y' and sr01=st05(+)" & _
            " and decode(st01,null,SR01,st01)='" & strUserNum & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   
   If intI = 1 Then bolUpd = True
   
   If bolUpd = True Or Pub_StrUserSt03 = "M51" Then
      cmdSave.Visible = True
   Else
      cmdSave.Visible = False
   End If
   
   QueryData
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210155 = Nothing
End Sub

' 查詢資料
Private Sub QueryData()
Dim rsTmp As New ADODB.Recordset
Dim rsQuery As ADODB.Recordset
Dim strSql As String
Dim strQueryLimit As String
Dim intR As Integer
   
Dim rs As New ADODB.Recordset
   
   
   strSql = "SELECT PLF01,max(PLF02) FROM pricelistfile" & _
            " WHERE PLF01 in('" & strKEY01 & "')" & _
            " group by PLF01"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         '版本(啟用日期=上傳日期)
         Label3.Caption = ChangeTStringToTDateString(TransDate(rsTmp.Fields(1), 1))
         rsTmp.MoveNext
      Loop
   Else
      Label3.Caption = ""
      cmdOpen.Enabled = False
   End If
   
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
   Set rsQuery = Nothing
End Sub

Private Function GetAttachFile(ByRef pFileName As String, ByVal strKEY01 As String, ByVal strKEY02 As String, _
                               Optional pSavePath As String, Optional pFileSize As Integer = 0) As Boolean
   Dim stAttPath As String
   Dim lngSize As Long
   Dim iFileNo As Integer
   Dim bytes() As Byte
   
On Error GoTo ErrHnd
   
   strExc(0) = "select * from pricelistfile where PLF01='" & strKEY01 & "' and PLF02=" & Val(DBDATE(strKEY02))
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If pSavePath = "" Then
         If Dir(m_AttachPath, vbDirectory) = "" Then
            MkDir m_AttachPath
         End If
         stAttPath = m_AttachPath & "\" & pFileName
         '檔案已存在時
         If Dir(stAttPath) <> "" Then
            '檢查檔案是否正在使用中
            If PUB_ChkFileOpening(stAttPath) = True Then
               MsgBox stAttPath & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
               Exit Function
            End If
            Kill stAttPath
         End If
      Else
         stAttPath = pSavePath
      End If
      
      If Dir(stAttPath) <> "" Then Kill stAttPath
      
      'Add By Sindy 2017/5/31
      If "" & RsTemp.Fields("plf11") <> "" Then
         GetAttachFile = PUB_GetFtpFile(RsTemp.Fields("plf11"), stAttPath, UCase("PRICELISTFILE"))
      Else
      '2017/5/31 END
         With RsTemp
            lngSize = Val(.Fields("PLF03").Value)
            ReDim bytes(lngSize)
            If lngSize > 0 Then bytes() = .Fields("PLF04").GetChunk(lngSize)
         End With
         iFileNo = FreeFile
         Open stAttPath For Binary Access Write As #iFileNo
         If lngSize > 0 Then Put #iFileNo, , bytes()
         Close #iFileNo
      End If
      pFileName = stAttPath
      If pFileSize = 1 Then
         pFileName = pFileName & " (" & Round(RsTemp.Fields("PLF03") / 1024, 2) & " KB)"
      End If
      GetAttachFile = True
   End If
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   If iFileNo > 0 Then Close #iFileNo
End Function

Private Function SaveAttFile(strKEY01 As String, strKEY02 As String, stFilePath As String) As Boolean
Dim ii As Integer, jj As Integer
Dim iFileNo As Integer
Dim bytes() As Byte
Dim lngSize As Long '檔案大小
Dim adoRst As New ADODB.Recordset
Const BlockSize = 500000
Dim Numblocks As Integer
Dim LeftOver As Long
Dim stReName As String, strFtpPath As String 'Add By Sindy 2017/5/31
   
   SaveAttFile = True
   
   stFilePath = Left(stFilePath, InStrRev(stFilePath, " (") - 1)
   If iFileNo > 0 Then Close #iFileNo
   iFileNo = FreeFile
   Open stFilePath For Binary Access Read As #iFileNo
   lngSize = LOF(iFileNo)
   Close #iFileNo
   If lngSize = 0 Then
      SaveAttFile = False
      ShowMsg stFilePath & MsgText(9221)
      Exit Function
   End If
   
   cnnConnection.BeginTrans
   PUB_DelFtpFile2 strKEY01, " and plf02=" & strKEY02, UCase("pricelistfile") 'Add By Sindy 2017/5/31 檔案改放 FTP,必須在DB資料刪除前執行
   strSql = "delete from pricelistfile where plf01='" & strKEY01 & "' and plf02=" & strKEY02
   cnnConnection.Execute strSql, intI
   
   'Add By Sindy 2017/5/31
   '改上傳FTP File Server
   stReName = strKEY02 & "." & lngSize & "." & GetFileName(stFilePath)
   PUB_PutFtpFile stFilePath, strKEY01, stReName, strFtpPath, UCase("pricelistfile")
   If strFtpPath <> "" Then
      strSql = "insert into pricelistfile(plf01,plf02,plf03,plf11) " & _
               "values(" & CNULL(strKEY01) & "," & strKEY02 & _
               "," & lngSize & "," & CNULL(strFtpPath) & ")"
      cnnConnection.Execute strSql
   Else
      Err.Raise 999, , " 檔案 " & stFilePath & " 上傳失敗!!"
   End If
   
   cnnConnection.CommitTrans
End Function
