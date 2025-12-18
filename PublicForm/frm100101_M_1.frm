VERSION 5.00
Begin VB.Form frm100101_M_1 
   AutoRedraw      =   -1  'True
   Caption         =   "原始檔Word維護-修改"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  '手動
   ScaleHeight     =   6795
   ScaleWidth      =   8955
   Begin VB.ComboBox cboPrinter 
      Height          =   300
      Left            =   5325
      Style           =   2  '單純下拉式
      TabIndex        =   2
      Top             =   60
      Width           =   3585
   End
   Begin VB.TextBox txtPDFPath 
      Height          =   315
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "C:\Program Files\Adobe\Reader 8.0\Reader\AcroRd32.exe"
      Top             =   7890
      Width           =   7485
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   45
      TabIndex        =   4
      Top             =   -90
      Width           =   4470
      Begin VB.CommandButton Command2 
         Caption         =   "上傳+EMail"
         Enabled         =   0   'False
         Height          =   345
         Index           =   2
         Left            =   2205
         TabIndex        =   8
         Top             =   150
         Width           =   1050
      End
      Begin VB.CommandButton Command2 
         Caption         =   "上傳+列印"
         Height          =   345
         Index           =   0
         Left            =   1215
         TabIndex        =   7
         Top             =   150
         Width           =   960
      End
      Begin VB.CommandButton Command3 
         Caption         =   "取消"
         Height          =   345
         Left            =   3600
         TabIndex        =   6
         Top             =   150
         Width           =   825
      End
      Begin VB.CommandButton Command2 
         Caption         =   "上傳"
         Height          =   345
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   150
         Width           =   825
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "印表機："
      Height          =   180
      Index           =   4
      Left            =   4590
      TabIndex        =   3
      Top             =   105
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PDF執行檔路徑："
      Height          =   180
      Left            =   45
      TabIndex        =   1
      Top             =   7950
      Width           =   1380
   End
End
Attribute VB_Name = "frm100101_M_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/24 Form2.不用改; Form2.0已檢查 (無需修改的物件)
'Created by Lydia 2020/02/06 原始檔Word維護-修改 (從frm1105_1複製)
Option Explicit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal _
     hWndNewParent As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
     (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Dim bolOpened As Boolean

Dim m_PrevForm As Form
Dim m_CaseNo As String '本所案號
Dim m_CP(1 To 4) As String
Dim m_CP10 As String
Dim m_RecNo As String  'CPF01
Dim m_RecCPF02 As String 'CPF02
Dim m_Status As String '上傳前先檢查有重複檔案 , D = 刪除舊檔, A = 自動 + 系統日期(時間), 空白 = 不上傳
Dim m_Type As String '另外處理：2=是否有執行上傳，回傳給前表單

Dim bolActived As Boolean
Dim m_AttachPath As String
Dim m_Os_Printer As String
Dim m_DocFullPath As String
Dim m_WordWidth As Long, m_WordHeight As Long, m_WordTop As Long
Dim m_WordAp As Word.Application
Dim intLastRow As Integer

Public Sub SetFormParent(ByRef pForm As Form, ByRef pCaseNo As String, ByRef pCPF01 As String, ByRef pCPF02 As String, ByRef pStatus As String, ByRef pCP10 As String)
     Set m_PrevForm = pForm
     m_CaseNo = pCaseNo
     Call ChgCaseNo(m_CaseNo, m_CP)
     pCP10 = m_CP10
     m_RecNo = pCPF01
     m_RecCPF02 = pCPF02
     If Len(pStatus) > 1 Then
         m_Status = Left(pStatus, 1)
         m_Type = Mid(pStatus, 2) '另外處理：2=是否有執行上傳，回傳給前表單
     Else
         m_Status = pStatus
     End If
End Sub

Private Sub OpenDoc2()
   Dim iTimes As Integer
   Dim stWinName As String
   Dim hWnd As Long
   
On Error GoTo Err_Handler
      
   'Modified by Lydia 2020/02/20 改放在使用者\原檔名
   'm_DocFullPath = App.path & "\$" & m_RecNo & ".doc"
   m_DocFullPath = m_AttachPath & "\" & m_RecCPF02
   'Modified by Lydia 2020/02/25 Word選項->儲存->設為*.doc時，若指定docx檔在存Word時會出錯, 所以要指定存檔類型
   'g_WordAp.ActiveDocument.SaveAs m_DocFullPath
   g_WordAp.ActiveDocument.SaveAs FileName:=m_DocFullPath, FileFormat:=12 '(12=wdFormatXMLDocument, XML 文件格式)
   
   g_WordAp.Quit wdDoNotSaveChanges
   Set g_WordAp = Nothing
   Set m_WordAp = New Word.Application
   m_WordAp.Documents.Open m_DocFullPath
   m_WordAp.Visible = True

   'Modified by Morgan 2017/4/17 Word97的視窗名稱不同
   'Modified by Morgan 2019/2/22 Word2013的視窗名稱不同
   If Val(m_WordAp.Version) >= 15 Then
      hWnd = FindWindow(vbNullString, m_WordAp.ActiveWindow.Caption & " - Word")
   ElseIf Val(m_WordAp.Version) > 8 Then
      hWnd = FindWindow(vbNullString, m_WordAp.ActiveWindow.Caption & " - Microsoft Word")
   Else
      hWnd = FindWindow(vbNullString, "Microsoft Word - " & m_WordAp.ActiveWindow.Caption)
   End If
   'end 2017/4/17
   
   If hWnd <> 0 Then
      SetParent hWnd, Me.hWnd
      SetWordPos
      bolOpened = True
   End If
   
Err_Handler:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Sub

'------檔案回寫到FTP
Private Function Conver2Pdf(pIdx As Integer) As Boolean
   Dim strFullFileName As String, strPdfName As String
   Dim oFileSys 'As New FileSystemObject
   Dim oFile 'As File
   Dim boInTrans As Boolean
   Dim strDocName As String, strDocName2 As String
   Dim tmpTime As String
   Dim stReName As String
   
On Error GoTo ErrHnd

   Me.Enabled = False
   
   'Memo by Lydia 2020/02/06 保留：存PDF
'   strPdfName = "$" & m_RecNo & ".PDF"
'   strFullFileName = m_AttachPath & "\" & strPdfName
'   If Dir(strFullFileName) <> "" Then
'      Kill strFullFileName
'   End If
   'end 2020/02/06 保留：存PDF
   
   m_WordAp.ActiveDocument.Save
      
   '用Word轉Pdf功能
   'Memo by Lydia 2020/02/06 保留：存PDF
'   If pub_Word2Pdf Then
'      m_WordAp.ActiveDocument.ExportAsFixedFormat OutputFileName:=strFullFileName, ExportFormat:=17, OpenAfterExport:=False
'   Else
'      frmPDF.Show
'      frmPDF.StartProcess m_AttachPath, strPdfName
'      '不用IE
'      m_WordAp.PrintOut Background:=False, Copies:=1, Collate:=True
'
'      frmPDF.EndtProcess
'      Unload frmPDF
'   End If
   'end 2020/02/06 保留：存PDF
   
   '寫回原始檔
   strFullFileName = m_DocFullPath
   If Dir(strFullFileName) <> "" Then
      cnnConnection.BeginTrans
      boInTrans = True
         strDocName = m_RecCPF02
         strSql = "update CasePaperFile set cpf02=cpf02 where cpf01='" & m_RecNo & "' and upper(cpf02)=upper('" & strDocName & "')"
         cnnConnection.Execute strSql, intI
         If intI > 0 Then
            Select Case m_Status
                Case "D", "A"
                    If m_Status = "D" Then   '刪除舊檔
                         '指定bolConn=True
                         If DelAttFile_File(m_CaseNo, m_RecNo, strDocName, , True) = False Then
                             GoTo ErrOut
                         End If
                    ElseIf m_Status = "A" Then '副檔名:自動+日期時間
JumpToNewTime:
                          If tmpTime = "" Then
                             tmpTime = Format(ServerTime, "000000")
                          Else
                             tmpTime = Format(Val(tmpTime) + 1, "000000")
                          End If
                          strDocName2 = Mid(strDocName, 1, InStrRev(strDocName, ".")) & strSrvDate(2) & tmpTime & Mid(strDocName, InStrRev(strDocName, "."))
                          strDocName2 = PUB_GetReNameMax(strDocName2, m_CP(1), m_CP(2), m_CP(3), m_CP(4), m_CP10, 75)
                          strSql = "update CasePaperFile set cpf02=cpf02 where cpf01='" & m_RecNo & "' and upper(cpf02)=upper('" & strDocName2 & "')"
                          cnnConnection.Execute strSql, intI
                          If intI > 0 Then
                              GoTo JumpToNewTime
                          End If
                    End If
                Case Else  '不上傳
                    cnnConnection.RollbackTrans
                    MsgBox "原始檔[" & m_RecCPF02 & "]已存在，請先更名或刪除後再上傳！", vbExclamation
                    GoTo ErrOut
            End Select
         End If
          
         If PUB_GetEmpFlowReNameFile(m_CP(1), m_CP(2), m_CP(3), m_CP(4), m_CP10, IIf(strDocName2 <> "", strDocName2, strDocName), stReName, True, 0) = False Then GoTo ErrOut
         
         m_WordAp.Quit wdDoNotSaveChanges
         Set m_WordAp = Nothing
         
         Set oFileSys = CreateObject("Scripting.FileSystemObject")
         Set oFile = oFileSys.GetFile(strFullFileName)
        '上傳到原始檔區: 指定檔名
         If SaveAttFile_Org(m_RecNo, strFullFileName, stReName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), "A", IIf(strDocName2 <> "", strDocName2, strDocName)) = False Then
            cnnConnection.RollbackTrans
            GoTo ErrOut
         End If
      
      cnnConnection.CommitTrans
      boInTrans = False
      Conver2Pdf = True
   End If
   
ErrHnd:
   If Err.Number > 0 Then
      If boInTrans Then cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   
ErrOut:
   Me.Enabled = True
   Set oFileSys = Nothing
   Set oFile = Nothing

End Function

Private Sub PrintWordDoc()
   Dim oWordApp As Word.Application
   Set oWordApp = New Word.Application
On Error GoTo ErrHnd
   
   'Modified by Lydia 2020/02/20 改放在使用者\原檔名
   'oWordApp.Documents.Open FileName:=App.path & "\$" & m_RecNo & ".doc", ReadOnly:=True
   oWordApp.Documents.Open FileName:=m_DocFullPath, ReadOnly:=True
   oWordApp.ActiveDocument.PrintOut
   oWordApp.Quit wdDoNotSaveChanges
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   Set oWordApp = Nothing
End Sub

Private Sub Command2_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   If Conver2Pdf(Index) = True Then
      If m_Type <> "" Then m_Type = m_Type & "OK"
      Unload Me
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command3_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Dim lngMaxHeight As Long
   
   If bolActived = False Then
      bolActived = True
      Me.WindowState = 2
      lngMaxHeight = Me.Height - 200
      Me.WindowState = 0
      Me.Height = lngMaxHeight
      Me.Width = (Me.Height - 400 - Command2(0).Height) * 19 / 21
      Me.Top = 0
      
      '若用Word轉Pdf功能則不必設定印表機
      If pub_Word2Pdf = False Then
         m_Os_Printer = PUB_GetOsDefaultPrinter
         PUB_SetOsDefaultPrinter "PDFCreator"
         '不用IE
         PUB_SetWordActivePrinter '切換Word印表機到PDFCreator
      End If
      OpenDoc2
   End If
   Me.ZOrder

End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cboPrinter
   txtPDFPath = PUB_SetFileAssociation
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   Else
      'Modified by Lydia 2020/02/20 改用模組
      'KillTemp
      PUB_KillTempFile strUserNum & "\*.*"
   End If
   
   '預設功能按鈕
   Command2(0).Visible = False
   Command2(1).Visible = True
   Command2(2).Visible = False

End Sub
'Mark by Lydia 2020/02/20 保留: 定稿維護需要Word轉PDF, 所以放在不同層
'Private Sub KillTemp()
'On Error Resume Next
'
   'Kill App.path & "\$*.doc"   '將Word放在App.Path
   'If Dir(m_AttachPath & "\.") <> "" Then
   '   Kill m_AttachPath & "\*.*"   '將PDF放在App.Path\使用者帳號
   'End If
   'Err.Clear
'End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   If Not m_WordAp Is Nothing Then
      m_WordAp.Quit wdDoNotSaveChanges
      Set m_WordAp = Nothing
   End If
   '回原始檔區
   If Not m_PrevForm Is Nothing Then
      If UCase(m_PrevForm.Name) = UCase("frm100101_M") Then
         m_PrevForm.Show
         Call m_PrevForm.ReadAttachFile(0) 'Modify By Sindy 2020/3/17
      ElseIf m_Type <> "" Then '另外處理：2=是否有執行上傳，回傳給前表單
             If m_Type = "2OK" Then
                 m_PrevForm.bolAskPA174 = True
             ElseIf m_Type = "2" Then
                 m_PrevForm.bolAskPA174 = False
             End If
             m_PrevForm.PubShowNextData
      End If
   End If

   Set frm100101_M_1 = Nothing
End Sub

Private Sub Form_Resize()
   If Me.WindowState = 0 Or Me.WindowState = 2 Then
      m_WordTop = Frame1.Top + Frame1.Height
      m_WordWidth = Me.Width - 200
      m_WordHeight = Me.Height - 400 - Frame1.Height - txtPDFPath.Height
      txtPDFPath.Top = m_WordTop + m_WordHeight
      If bolOpened = True Then
         SetWordPos
      End If
      Label3.Top = txtPDFPath.Top + 50
   End If
End Sub

Private Sub SetWordPos()
m_WordAp.Width = m_WordWidth / 20
m_WordAp.Height = m_WordHeight / 20
m_WordAp.Move 0, m_WordTop / 20
End Sub
