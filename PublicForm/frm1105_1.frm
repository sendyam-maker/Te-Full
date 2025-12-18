VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm1105_1 
   AutoRedraw      =   -1  'True
   Caption         =   "定稿維護-修改"
   ClientHeight    =   8244
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8952
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   OLEDropMode     =   1  '手動
   ScaleHeight     =   8244
   ScaleWidth      =   8952
   Begin VB.Frame Frame4 
      Caption         =   "請選擇通知函的收文號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   45
      TabIndex        =   20
      Top             =   5190
      Visible         =   0   'False
      Width           =   4470
      Begin VB.CommandButton cmdRecNo 
         Caption         =   "確定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   90
         TabIndex        =   21
         Top             =   1770
         Width           =   2130
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   1485
         Left            =   90
         TabIndex        =   22
         Top             =   210
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   2625
         _Version        =   393216
         Cols            =   4
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "V|收文日|總收文號|案件性質"
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
         _Band(0).Cols   =   4
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "請選擇要合併的收文號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   45
      TabIndex        =   15
      Top             =   2820
      Visible         =   0   'False
      Width           =   7260
      Begin VB.CommandButton cmdJoinAct 
         Caption         =   "取消"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   2250
         TabIndex        =   17
         Top             =   1770
         Width           =   2130
      End
      Begin VB.CommandButton cmdJoinAct 
         Caption         =   "確定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   90
         TabIndex        =   16
         Top             =   1770
         Width           =   2130
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1275
         Left            =   90
         TabIndex        =   18
         Top             =   210
         Width           =   7050
         _ExtentX        =   12425
         _ExtentY        =   2244
         _Version        =   393216
         Cols            =   5
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "V|收文日|本所案號|案件性質|總收文號"
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
      Begin VB.CheckBox Check1 
         Caption         =   "開啟要合併的定稿Word"
         Height          =   315
         Left            =   90
         TabIndex        =   19
         Top             =   1470
         Value           =   1  '核取
         Width           =   2355
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "判發退回意見"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   45
      TabIndex        =   8
      Top             =   450
      Visible         =   0   'False
      Width           =   4470
      Begin VB.CommandButton Command1 
         Caption         =   "確定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   90
         TabIndex        =   10
         Top             =   1740
         Width           =   2130
      End
      Begin VB.CommandButton Command4 
         Caption         =   "確定刪除"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2250
         TabIndex        =   13
         Top             =   1740
         Visible         =   0   'False
         Width           =   2130
      End
      Begin MSForms.TextBox Text1 
         Height          =   1425
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   4245
         VariousPropertyBits=   -1466939365
         ScrollBars      =   2
         Size            =   "7488;2514"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
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
      Begin VB.CommandButton cmdJoin 
         Caption         =   "合併"
         Enabled         =   0   'False
         Height          =   345
         Left            =   630
         TabIndex        =   14
         Top             =   150
         Width           =   555
      End
      Begin VB.CommandButton Command2 
         Caption         =   "上傳+EMail"
         Enabled         =   0   'False
         Height          =   345
         Index           =   2
         Left            =   2205
         TabIndex        =   11
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
         Left            =   3870
         TabIndex        =   6
         Top             =   150
         Width           =   555
      End
      Begin VB.CommandButton Command2 
         Caption         =   "上傳"
         Height          =   345
         Index           =   1
         Left            =   3285
         TabIndex        =   5
         Top             =   150
         Width           =   555
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "刪除"
         Enabled         =   0   'False
         Height          =   345
         Left            =   45
         TabIndex        =   12
         Top             =   150
         Width           =   555
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
Attribute VB_Name = "frm1105_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/11/29 改成Form2.0 ; Text1、MSHFlexGrid1改字型=新細明體-ExtB、MSHFlexGrid2改字型=新細明體-ExtB
'Created by Morgan 2014/3/27
Option Explicit

'Added by Morgan 2016/9/20
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal _
     hWndNewParent As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
     (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Dim bolOpened As Boolean
'Dim stOsVerNo As String
'end 2016/9/20

Public m_PrevForm As Form 'Added by Morgan 2015/11/17
Public m_PdfName As String
Public m_RecNo As String
Public m_Subject As String 'Added by Morgan 2016/3/30
Public m_eCustNo As String 'Added by Morgan 2018/11/1

Dim bolActived As Boolean
Dim m_AttachPath As String
Dim m_Os_Printer As String
Dim m_DocFullPath As String
Dim m_CP140 As String 'Added by Morgan 2018/8/30
Dim m_WordWidth As Long, m_WordHeight As Long, m_WordTop As Long 'Added by Morgan 2018/8/31
Dim m_WordAp As Word.Application 'Added by Morgan 2018/10/15
Dim m_PropertyDesc As String  'Added by Morgan 2018/10/16
Dim intLastRow As Integer 'Added by Morgan 2018/11/1
Dim m_WordDoc As Word.Document 'Added by Morgan 2021/12/7
Dim m_bResize As Boolean  'Added by Morgan 2022/11/10

'Removed by Morgan 2018/8/31
'Private Sub OpenDoc()
'   Dim stDocFullPath As String
'   Dim iTimes As Integer
'
'On Error GoTo Err_Handler
'
'   stDocFullPath = App.path & "\$" & m_RecNo & ".doc"
'   g_WordAp.ActiveDocument.SaveAs stDocFullPath
'   g_WordAp.Quit wdDoNotSaveChanges
'   Set g_WordAp = Nothing
'   WebBrowser1.Visible = True
'   WebBrowser1.Navigate stDocFullPath
'   DoEvents
'   Do While WebBrowser1.Busy = True
'      Sleep 500
'   Loop
'   WebBrowser1.ExecWB OLECMDID_HIDETOOLBARS, OLECMDEXECOPT_DONTPROMPTUSER
'   Exit Sub
'
'Err_Handler:
'   If Err.NUMBER = -2147221248 Then
'      If iTimes < 3 Then
'         iTimes = iTimes + 1
'         Sleep 1000
'         Resume
'      End If
'   End If
'   MsgBox Err.Description, vbCritical
'End Sub

Private Sub OpenDoc2()
   Dim iTimes As Integer
   Dim stWinName As String
   Dim hWnd As Long
   
On Error GoTo Err_Handler
      
   m_DocFullPath = App.path & "\$" & m_RecNo & ".doc"
   g_WordAp.ActiveDocument.SaveAs m_DocFullPath
   
   'Added by Morgan 2018/10/15
   g_WordAp.Quit wdDoNotSaveChanges
   Set g_WordAp = Nothing
   Set m_WordAp = New Word.Application
   'Modified by Morgan 2021/12/7
   'm_WordAp.Documents.Open m_DocFullPath
   Set m_WordDoc = m_WordAp.Documents.Open(m_DocFullPath)
   m_WordAp.Visible = True
   'end 2018/10/15
   
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

Private Function Conver2Pdf(pIdx As Integer) As Boolean
   Dim strFullFileName As String, strPdfName As String
   Dim oFileSys As New FileSystemObject
   Dim oFile As File
   Dim boInTrans As Boolean
   Dim oFile2 As File, strDocName As String, strDocName2 As String 'Added by Morgan 2018/9/13
   Dim bDelPdf As Boolean

On Error GoTo ErrHnd

   Me.Enabled = False
   
   'Added by Morgan 2025/10/14 若畫面已開啟又再次被呼叫時會發生上傳的內容與案件不符的狀況，Ex:P-094024不續辦未上傳指示信又繼續不續辦P-091791。
   If InStr(m_WordDoc.Name, m_RecNo) = 0 Then
      MsgBox "定稿內容與將上傳的案件不符，不可繼續！" & vbCrLf & vbCrLf & "將上傳的收文號: " & m_RecNo & vbCrLf & "將上傳的檔案名: " & m_PdfName, vbCritical
      GoTo ErrOut
   End If
   'end 2025/10/14
   
   'oDocument.Activate
   'oDocument.Application.PrintOut Background:=False, Copies:=1, Collate:=True
   'SkipMark 'Removed by Morgan 2014/5/7 給客戶的定稿和留所的相同
   
   strPdfName = "$" & m_RecNo & ".PDF"
   strFullFileName = m_AttachPath & "\" & strPdfName
   If Dir(strFullFileName) <> "" Then
      Kill strFullFileName
   End If
   
   'Set g_WordAp = oDocument.Application
   'Set g_WordAp = OLE1.object.Application
   '轉pdf
   'Load frmPDF
   
   'Modified by Morgan 2021/12/7 改寫法，因Uuser可能會開其他的文件導致存錯檔案
   'm_WordAp.ActiveDocument.Save
   ''Added by Morgan 2017/8/9 用Word轉Pdf功能
   'If pub_Word2Pdf Then
   '   m_WordAp.ActiveDocument.ExportAsFixedFormat OutputFileName:=strFullFileName, ExportFormat:=17, OpenAfterExport:=False
   m_WordDoc.Save
   If pub_Word2Pdf Then
      m_WordDoc.ExportAsFixedFormat OutputFileName:=strFullFileName, ExportFormat:=17, OpenAfterExport:=False
   'end 2021/12/7
   Else
   'end 2017/8/9

      frmPDF.Show
      frmPDF.StartProcess m_AttachPath, strPdfName
   'Added by Morgan 2016/9/20
   'Win7無法用IE開Word改將Word開在Form內
   'Modified by Morgan 2016/9/26 改都不用IE
   'If Val(stOsVerNo) > 6 Then
      m_WordAp.PrintOut Background:=False, Copies:=1, Collate:=True
   'Else
   ''end 2016/9/20
   '   m_Os_Printer = PUB_GetOsDefaultPrinter
   '   PUB_SetOsDefaultPrinter Printer.DeviceName
   '   WebBrowser1.ExecWB OLECMDID_SAVE, OLECMDEXECOPT_DONTPROMPTUSER
   '   WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
   '   PUB_SetOsDefaultPrinter m_Os_Printer
   'End If 'Added by Morgan 2016/9/20
   'end 2016/9/26

      frmPDF.EndtProcess
      Unload frmPDF
   End If
   
   '寫回卷宗區並更新信函進度檔
   If Dir(strFullFileName) <> "" Then
      
      strSql = "update CasePaperPDF set cpp01=cpp01 where cpp01='" & m_RecNo & "' and upper(cpp02)=upper('" & m_PdfName & "')"
      cnnConnection.Execute strSql, intI
      If intI > 0 Then
         'Added by Morgan 2018/8/7 +CFP指示信需檢查卷宗區若已存在需人工刪除或更名(.data副檔名未強制規範,可能有其他內容)
         If InStr(UCase(m_PdfName), ".DATA.PDF") > 0 And Left(UCase(m_PdfName), 3) = "CFP" Then
            MsgBox "卷宗區已存在" & m_PdfName & "，請先更名或刪除後再上傳指示信！", vbExclamation
            GoTo ErrOut
         Else
         'end 2018/8/7
            If MsgBox("卷宗區已存在" & m_PdfName & "，是否要覆蓋？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
               GoTo ErrOut
            End If
         End If
         bDelPdf = True
      End If
      
      
      Set oFile = oFileSys.GetFile(strFullFileName)
      cnnConnection.BeginTrans
      boInTrans = True
      
      If bDelPdf Then
         PUB_DelFtpFile2 m_RecNo, " and upper(cpp02)=upper('" & m_PdfName & "')" 'Added by Morgan 2015/4/15 檔案改放 FTP,必須在DB資料刪除前執行
         
         strSql = "delete from CasePaperPDF where cpp01='" & m_RecNo & "' and upper(cpp02)=upper('" & m_PdfName & "')"
         cnnConnection.Execute strSql, intI
      End If
      
      If m_eCustNo = "" Then 'Added by Morgan 2018/11/1
      
         'Added by Morgan 2015/11/6 先鎖定以免自動轉PDF程式也在處理該筆資料
         If InStr(UCase(m_PdfName), ".DATA.PDF") > 0 Then
            strSql = "update AppForm set af02='" & strUserNum & "' where af01='" & m_RecNo & "'"
            cnnConnection.Execute strSql, intI
         Else
            strSql = "update letterprogress set lp08='" & strUserNum & "' where lp01='" & m_RecNo & "'"
            cnnConnection.Execute strSql, intI
         End If
         'end 2015/11/6
         
      End If 'Added by Morgan 2018/11/1
   
      'Added by Morgan 2018/9/13
      'Modified by Morgan 2019/1/17 考慮客戶函判發可能退回,改修改上傳也要同時存原始檔
      'Modified by Morgan 2019/1/19 +改P指示信也要存原始檔--玲玲
      'If (Right(UCase(m_PdfName), 9) = ".DATA.PDF" And Left(m_PdfName, 3) = "CFP") Or (Right(UCase(m_PdfName), 8) = ".CUS.PDF" And (Left(m_PdfName, 3) = "CFP" Or Left(m_PdfName, 1) = "P")) Then
'Modify By Sindy 2019/11/29 取消系統別限制
'      If (Left(m_PdfName, 3) = "CFP" Or Left(m_PdfName, 1) = "P") Then
'2019/11/29 END
         strDocName = Replace(UCase(m_PdfName), ".PDF", ".DOC")
         strSql = "update CasePaperFile set cpf02=cpf02 where cpf01='" & m_RecNo & "' and upper(cpf02)=upper('" & strDocName & "')"
         cnnConnection.Execute strSql, intI
         If intI > 0 Then
            'Modified by Morgan 2019/1/18
            'cnnConnection.RollbackTrans
            'MsgBox "原始檔[" & strDocName & "]已存在，請先更名或刪除後再上傳！", vbExclamation
            'GoTo ErrOut
            If Right(UCase(m_PdfName), 9) = ".DATA.PDF" Then
               strDocName2 = Replace(strDocName, ".DATA.DOC", ".OLD.DATA.DOC")
            Else
               strDocName2 = Replace(strDocName, ".CUS.DOC", ".OLD.CUS.DOC")
            End If
            strSql = "update CasePaperFile set cpf02=cpf02 where cpf01='" & m_RecNo & "' and upper(cpf02)=upper('" & strDocName2 & "')"
            cnnConnection.Execute strSql, intI
            If intI = 0 Then
               strSql = "update CasePaperFile set cpf02='" & strDocName2 & "' where cpf01='" & m_RecNo & "' and upper(cpf02)=upper('" & strDocName & "')"
               cnnConnection.Execute strSql, intI
            Else
               cnnConnection.RollbackTrans
               MsgBox "原始檔備份[" & strDocName2 & "]已存在，請先更名或刪除後再上傳！", vbExclamation
               GoTo ErrOut
            End If
            'end 2019/1/18
         End If
         
         'Modified by Morgan 2018/10/15
         'g_WordAp.Quit wdDoNotSaveChanges
         'Set g_WordAp = Nothing
         'Modified by Morgan 2021/12/7
         'm_WordAp.Quit wdDoNotSaveChanges
         m_WordDoc.Close wdDoNotSaveChanges
         Set m_WordDoc = Nothing
         If m_WordAp.Documents.Count = 0 Then
            m_WordAp.Quit wdDoNotSaveChanges
         End If
         Set m_WordAp = Nothing
         bolOpened = False 'Added by Morgan 2022/6/28
         'end 2021/12/7
         'end 2018/10/15
         
         Set oFile2 = oFileSys.GetFile(m_DocFullPath)
         If SaveAttFile_Org(m_RecNo, m_DocFullPath, strDocName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS")) = False Then
            cnnConnection.RollbackTrans
            GoTo ErrOut
         End If
         
         'Added by Morgan 2019/1/18
         '刪除原始檔備份
         If strDocName2 <> "" Then
            PUB_DelFtpFile2 m_RecNo, " and upper(cpf02)=upper('" & strDocName2 & "')", "CASEPAPERFILE"
            
            strSql = "delete from CasePaperFile where cpf01='" & m_RecNo & "' and upper(cpf02)=upper('" & strDocName2 & "')"
            cnnConnection.Execute strSql, intI
         End If
         'end 2019/1/18
'      End If
      'end 2018/9/13
      
      
      'Modify By Sindy 2015/5/14
      'SaveAttFile_PDF m_RecNo, strFullFileName, m_PdfName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, "4", , , True
      SaveAttFile_PDF m_RecNo, strFullFileName, m_PdfName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , , True
      '2015/5/14 END
      
      If m_eCustNo = "" Then 'Added by Morgan 2018/11/1
      
         If InStr(UCase(m_PdfName), ".DATA.PDF") > 0 Then
            strSql = "update AppForm set af02='" & strUserNum & "',af03=to_char(sysdate,'yyyymmdd') where af01='" & m_RecNo & "'"
            cnnConnection.Execute strSql, intI
            'Memo by Morgan 2016/5/12 自行判發指示信新增至卷宗區時trigger會上判發日
            'Added by Morgan 2015/12/21
            '指示信上傳EMail通知判發人員(非自行判發)
            strSql = "update AppForm set af07=af07 where af01='" & m_RecNo & "' and af07=0 and af06<>'" & strUserNum & "'"
            cnnConnection.Execute strSql, intI
            If intI = 1 Then
                strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                  " select '" & strUserNum & "' mc01,af06 mc02,to_char(sysdate,'yyyymmdd') mc03,to_char(sysdate,'hh24miss') mc04" & _
                  ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||'指示信已上傳待判發!!' mc07" & _
                  ",'如旨' mc08 from appform,caseprogress" & _
                  " where af01='" & m_RecNo & "' and cp09(+)=af01"
               cnnConnection.Execute strSql, intI
            End If
            'end 2015/12/21
         Else
            strSql = "update letterprogress set lp08='" & strUserNum & "',lp09=to_char(sysdate,'yyyymmdd') where lp01='" & m_RecNo & "'"
            cnnConnection.Execute strSql, intI
            
            'Added by Morgan 2019/10/1
            'CFP已提申1909有通知信且非自判時EMail通知判發人
            If InStr(UCase(m_PdfName), ".1909.CUS.PDF") > 0 And Left(m_PdfName, 3) = "CFP" Then
               '有退回,非自判
               strSql = "update letterprogress set lp04=lp04 where lp01='" & m_RecNo & "' and lp04 is not null and LP37 is not null"
               cnnConnection.Execute strSql, intI
               If intI = 1 Then
                  strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                    " select '" & strUserNum & "' mc01,lp04 mc02,to_char(sysdate,'yyyymmdd') mc03,to_char(sysdate,'hh24miss') mc04" & _
                    ",c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04)||cpm04||'已提申來函已上傳請判發!!' mc07" & _
                    ",'如旨' mc08 from letterprogress,caseprogress c1,caseprogress c2,casepropertymap" & _
                    " where lp01='" & m_RecNo & "' and c1.cp09(+)=lp01 and c2.cp09(+)=c1.cp43 and cpm01(+)=c2.cp01 and cpm02(+)=c2.cp10"
                  cnnConnection.Execute strSql, intI
               End If
            End If
            'end 2019/10/1
         End If
         
      End If 'Added by Morgan 2018/11/1
      
      cnnConnection.CommitTrans
      boInTrans = False
      Conver2Pdf = True

      '上傳+列印
      If pIdx = 0 Then 'Added by Morgan 2015/6/26
         'Added by Morgan 2022/3/22
         '若定稿有公司章則要去掉後才列印
         If m_AutoStampNameInWord <> "" Then
            pub_OsPrinter = PUB_GetOsDefaultPrinter
            PUB_SetOsDefaultPrinter cboPrinter
            PUB_SetWordActivePrinter
            Set m_WordAp = New Word.Application
            Set m_WordDoc = m_WordAp.Documents.Open(m_DocFullPath)
            m_WordDoc.Shapes(m_AutoStampNameInWord).Delete
            m_WordDoc.PrintOut Background:=False, Copies:=1, Collate:=True
            PUB_SetOsDefaultPrinter pub_OsPrinter
            m_WordDoc.Close wdDoNotSaveChanges
            Set m_WordDoc = Nothing
            If m_WordAp.Documents.Count = 0 Then
               m_WordAp.Quit wdDoNotSaveChanges
            End If
            Set m_WordAp = Nothing
         Else
         'end 2022/3/22
            PUB_PrintPDF strFullFileName, Me.cboPrinter
         End If
      'Added by Morgan 2016/3/30
      '上傳+EMail,上傳+列印+EMail
      ElseIf pIdx = 2 Then
         If Command2(2).Caption = "上傳+列印+EMail" Then
            PUB_PrintPDF strFullFileName, Me.cboPrinter
         End If
         
         'Added by Morgan 2018/11/1
         If m_eCustNo <> "" Then
            'Modified by Morgan 2021/12/2 改先Mail智權會稿
            'If PUB_ChkEmailBackUp(m_RecNo) = True Then
             '  PUB_SendECustLetter m_RecNo, m_eCustNo
            'End If
            PUB_SendECustLetter m_RecNo, m_eCustNo, True
            'end 2021/12/2
         Else
         'end 2018/11/1
            
            PUB_SendOrderLetterP m_RecNo, m_Subject
            
         End If 'Added by Morgan 2018/11/1
      'end 2016/3/30
      End If
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
   Set oFile2 = Nothing 'Added by Morgan 2018/9/13
   'PUB_SetOsDefaultPrinter m_Os_Printer
End Function

Private Sub PrintWordDoc()
   Dim oWordApp As Word.Application
   Set oWordApp = New Word.Application
On Error GoTo ErrHnd
   
   oWordApp.Documents.Open FileName:=App.path & "\$" & m_RecNo & ".doc", ReadOnly:=True
   oWordApp.ActiveDocument.PrintOut
   oWordApp.Quit wdDoNotSaveChanges
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   Set oWordApp = Nothing
End Sub

'Added by Morgan 2018/8/30
Private Sub cmdDelete_Click()
   Frame2.Visible = True
   Frame1.Enabled = False
   Frame2.Caption = "指示信刪除原因"
   Text1.Text = "請輸入刪除原因!!"
   Text1.Locked = False 'Added by Morgan 2020/9/7
   Text1.Tag = Text1.Text
   Command1.Width = Text1.Width / 2
   Command1.Caption = "取消"
   Command4.Visible = True
   SetWordPos
   Text1.SetFocus
   'If MsgBox("是否確定要刪除？", vbYesNo + vbExclamation + vbDefaultButton2) = vbYes Then
   '   If DeleteAppForm() = True Then
   '      Unload Me
   '   End If
   'End If
End Sub

'Added by Morgan 2018/8/30
Private Function DeleteAppForm() As Boolean
   
   cnnConnection.BeginTrans
On Error GoTo ErrHnd
   
   strSql = "delete appform  where af01='" & m_RecNo & "'"
   Pub_SeekTbLog strSql 'Added by Morgan 2018/9/4
   cnnConnection.Execute strSql, intI
   
   strSql = "update caseprogress set cp64=to_char(sysdate,'yyyy.mm.dd')||' 指示信：" & Text1.Text & "; '||cp64 where cp09='" & m_RecNo & "'"
   cnnConnection.Execute strSql, intI
   
   If m_CP140 <> "" Then OrderLetterFlowStatusUpdate m_CP140
      
   cnnConnection.CommitTrans
   DeleteAppForm = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Sub cmdJoin_Click()
Dim bolRun As Boolean
   
   'Modify By Sindy 2020/10/30
   If Left(UCase(m_PdfName), 1) = "T" Then
      If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
         bolRun = GetJoinData_T
      End If
   Else
      bolRun = GetJoinData
   End If
   If bolRun = True Then
   '2020/10/30 END
      Frame3.Top = Frame2.Top
      Frame3.Visible = True
      Frame1.Enabled = False
      SetWordPos
   End If
End Sub

Private Function JoinUpdate() As Boolean
   Dim iRow As Integer, intR As Integer
   Dim strCP09 As String
   
   cnnConnection.BeginTrans
On Error GoTo ErrHnd
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" And .TextMatrix(iRow, 5) = "" Then
         strCP09 = .TextMatrix(iRow, 4)
         'Modify By Sindy 2020/10/30
         If Left(UCase(m_PdfName), 1) = "T" Then
            strExc(0) = "已併入" & m_PropertyDesc & "通知函(" & _
                        IIf(frm1105.Text1(2) & frm1105.Text1(3) = "000", _
                        frm1105.Text1(0) & "-" & frm1105.Text1(1), _
                        frm1105.Text1(0) & "-" & frm1105.Text1(1) & "-" & frm1105.Text1(2) & "-" & frm1105.Text1(3)) & _
                        ":" & m_RecNo & ")告知客戶;"
         Else
         '2020/10/30 END
            strExc(0) = "已併入" & m_PropertyDesc & "通知函(" & m_RecNo & ")告知客戶;"
         End If
         'Modified by Morgan 2019/9/9 +更新確認人員時間(與共同查詢一致)
         strSql = "update letterprogress set lp06='" & strUserNum & "',lp07=to_char(sysdate,'yyyymmdd'),lp10='N',lp12='" & strExc(0) & "'||lp12,lp42='" & m_RecNo & "'" & _
            " where lp01='" & strCP09 & "'"
         cnnConnection.Execute strSql, intR
         
         PUB_DelFtpFile2 strCP09, " and SUBSTR(UPPER(cpp02),-8)='.CUS.PDF'"
        
         strSql = "delete casepaperpdf where cpp01='" & strCP09 & "' and SUBSTR(UPPER(cpp02),-8)='.CUS.PDF'"
         cnnConnection.Execute strSql, intR
      End If
   Next
   End With
   cnnConnection.CommitTrans
   JoinUpdate = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
End Function

Private Function GetJoinData() As Boolean
   Dim iRow As Integer, stLP11 As String
   
   m_PropertyDesc = ""
   'Modified by Morgan 2020/1/14 +lp11
   strExc(0) = "select CPM04,lp11 from caseprogress,casepropertymap,letterprogress" & _
      " where cp09='" & m_RecNo & "' and cpm01(+)=cp01 and cpm02(+)=cp10 and lp01(+)=cp09"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      m_PropertyDesc = "" & RsTemp(0)
      m_PropertyDesc = m_PropertyDesc & PUB_GetRelateCasePropertyName(m_RecNo, "1")
      stLP11 = "" & RsTemp("lp11") 'Added by Morgan 2020/1/14
   End If
   
   SetGrid MSHFlexGrid1, True
   'Modified by Morgan 2018/10/29 原判斷同日發文，改判斷發文室未發文都可合併，因有可能發文通知未寄出就已提申 Ex:CFP030287
   'Modified by Morgan 2018/11/8 +未確認(LP07=0) or 先前合併的收文號(lp42=a.cp09)
   'Modified by Morgan 2019/1/4 +發文人員為QPGMR時也可合併 Ex:CFP-030582(E化案件)
   'Modified by Morgan 2020/1/14 +掛號直寄可合併掛號直寄(非掛號直寄則不可) Ex:CFP-030075 實審+年費的期限通知
   strExc(0) = "select decode(lp42,null,'','V') V,sqldatet(b.cp05) 收文日,b.cp01||'-'||b.cp02||decode(b.cp03||b.cp04,'000','','-'||b.cp03||'-'||b.cp04) 本所案號" & _
      ",cpm04 案件性質,b.cp09 總收文號,lp42,lp04,lp05" & _
      " from caseprogress a,caseprogress b,letterprogress,casepropertymap" & _
      " where a.cp09='" & m_RecNo & "'" & _
      " and b.cp01(+)=a.cp01 and b.cp02(+)=a.cp02 and b.cp03(+)=a.cp03 and b.cp04(+)=a.cp04" & _
      " and b.cp27>19221111 and b.cp09<>a.cp09 and lp01(+)=b.cp09 and (lp15='N' or b.cp154='QPGMR')" & IIf(stLP11 = "", " and nvl(lp11,'N')<>'Y'", "") & _
      " and (lp07=0 or lp42=a.cp09) and cpm01(+)=b.cp01 and cpm02(+)=b.cp10"
      
   'Added by Morgan 2019/1/15 +EU子案也可被合併 Ex:CFP-030817
   strExc(0) = strExc(0) & " union select decode(lp42,null,'','V') V,sqldatet(b.cp05) 收文日,b.cp01||'-'||b.cp02||decode(b.cp03||b.cp04,'000','','-'||b.cp03||'-'||b.cp04) 本所案號" & _
      ",cpm04 案件性質,b.cp09 總收文號,lp42,lp04,lp05" & _
      " from caseprogress a,patent,caseprogress b,letterprogress,casepropertymap" & _
      " where a.cp09='" & m_RecNo & "'" & _
      " and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04 and pa09='239'" & _
      " and b.cp01(+)=a.cp01 and b.cp02(+)=a.cp02 and b.cp03<>a.cp03 and b.cp04(+)=a.cp04" & _
      " and b.cp27>19221111 and lp01(+)=b.cp09 and (lp15='N' or b.cp154='QPGMR') and nvl(lp11,'N')<>'Y'" & _
      " and (lp07=0 or lp42=a.cp09) and cpm01(+)=b.cp01 and cpm02(+)=b.cp10"
   'end 2019/1/15
   
   'Added by Morgan 2019/1/18 +EPC可以母案併子案或子案併子案 Ex:CFP-027176
   strExc(0) = strExc(0) & " union select decode(lp42,null,'','V') V,sqldatet(b.cp05) 收文日,b.cp01||'-'||b.cp02||decode(b.cp03||b.cp04,'000','','-'||b.cp03||'-'||b.cp04) 本所案號" & _
      ",cpm04 案件性質,b.cp09 總收文號,lp42,lp04,lp05" & _
      " from caseprogress a,patent,caseprogress b,letterprogress,casepropertymap" & _
      " where a.cp09='" & m_RecNo & "'" & _
      " and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)='00' and pa09='221'" & _
      " and b.cp01(+)=a.cp01 and b.cp02(+)=a.cp02 and b.cp03(+)=a.cp03 and b.cp04<>a.cp04" & _
      " and b.cp27>19221111 and lp01(+)=b.cp09 and (lp15='N' or b.cp154='QPGMR') and nvl(lp11,'N')<>'Y'" & _
      " and (lp07=0 or lp42=a.cp09)and cpm01(+)=b.cp01 and cpm02(+)=b.cp10"
   'end 2019/1/18
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With MSHFlexGrid1
      .Visible = False
      Set .Recordset = RsTemp
      SetGrid MSHFlexGrid1
      
      For iRow = 1 To .Rows - 1
         .TextMatrix(iRow, 3) = .TextMatrix(iRow, 3) & PUB_GetRelateCasePropertyName(.TextMatrix(iRow, 4), "1")
      Next
      
      .Visible = True
      End With
      GetJoinData = True
   Else
      MsgBox "無其他程序的通知函可合併！", vbExclamation
   End If

End Function

Private Function GetJoinData_T() As Boolean
   Dim iRow As Integer, stLP11 As String
   Dim strCP27 As String, strTM23 As String, strTM44 As String
   
   m_PropertyDesc = ""
   strExc(0) = "select DECODE('" & frm1105.m_strNationNo & "','000',CPM03,CPM04),lp11 from caseprogress,casepropertymap,letterprogress" & _
      " where cp09='" & m_RecNo & "' and cpm01(+)=cp01 and cpm02(+)=cp10 and lp01(+)=cp09"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      m_PropertyDesc = "" & RsTemp(0)
      m_PropertyDesc = m_PropertyDesc & PUB_GetRelateCasePropertyName(m_RecNo, "1")
      stLP11 = "" & RsTemp("lp11")
   End If
   
   SetGrid MSHFlexGrid1, True
   '原判斷同日發文，改判斷發文室未發文都可合併
   '未確認(LP07=0) or 先前合併的收文號(lp42=a.cp09)
   '發文人員為QPGMR時也可合併
   '掛號直寄可合併掛號直寄(非掛號直寄則不可)
   strExc(0) = "select *" & _
               " from caseprogress,trademark" & _
               " where cp09='" & m_RecNo & "'" & _
               " and cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and cp27 is not null and cp57 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strCP27 = RsTemp.Fields("cp27")
      strTM23 = RsTemp.Fields("tm23") '申請人1
      strTM44 = "" & RsTemp.Fields("tm44") 'FC代理人
      
      strExc(0) = "select decode(lp42,null,'','V') V,sqldatet(cp05) 收文日,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
         ",DECODE(tm10,'000',CPM03,CPM04) 案件性質,cp09 總收文號,lp42,lp04,lp05" & _
         " from caseprogress,letterprogress,trademark,casepropertymap" & _
         " where cp27=" & strCP27 & " and cp57 is null and cp09<>'" & m_RecNo & "'" & _
         " and cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and tm23='" & strTM23 & "'" & IIf(strTM44 <> "", " and tm44='" & strTM44 & "'", " and tm44 is null") & _
         " and lp01(+)=cp09 and (lp15='N' or cp154='QPGMR')" & IIf(stLP11 = "", " and nvl(lp11,'N')<>'Y'", "") & _
         " and (lp07=0 or lp42='" & m_RecNo & "') and cpm01(+)=cp01 and cpm02(+)=cp10"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With MSHFlexGrid1
         .Visible = False
         Set .Recordset = RsTemp
         SetGrid MSHFlexGrid1
         
         For iRow = 1 To .Rows - 1
            .TextMatrix(iRow, 3) = .TextMatrix(iRow, 3) & PUB_GetRelateCasePropertyName(.TextMatrix(iRow, 4), "1")
         Next
         
         .Visible = True
         End With
         GetJoinData_T = True
      Else
         MsgBox "無其他程序的通知函可合併！", vbExclamation
      End If
   Else
      MsgBox "無其他程序的通知函可合併！", vbExclamation
   End If

End Function

Private Sub SetGrid(oGrid As MSHFlexGrid, Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGridHeadWidth
   Dim iUbound As Integer

   If oGrid.Name = "MSHFlexGrid1" Then
      arrGridHeadWidth = Array(240, 800, 1400, 1900, 1000)
   Else
      arrGridHeadWidth = Array(240, 800, 1000, 1900)
   End If
   iUbound = UBound(arrGridHeadWidth)
   
   With oGrid
   If pReset = True Then
      .Clear
      .Rows = 2
   End If
   .FixedCols = 0
   If oGrid.Name = "MSHFlexGrid1" Then
      .FormatString = "V|收文日|本所案號|案件性質|總收文號"
   Else
      .FormatString = "V|發文日|總收文號|案件性質"
   End If
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

Private Sub OpenWord()
   Dim iRow As Integer, intR As Integer
   Dim strCP09 As String
   
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         strCP09 = .TextMatrix(iRow, 4)
         strExc(0) = "select ld01,ld04,ld10,ld11 from letterdemand where ld18='" & strCP09 & "' order by ld02 desc,ld03 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strUserNum = RsTemp("ld01") 'Added by Morgan 2022/4/8 被合併的定稿可能是不同人的，要切換使用者，否則會抓不到例外欄位資料
            NowPrint RsTemp("ld04"), RsTemp("ld10"), RsTemp("ld11"), True, strUserNum, , , , , , , , , False, , , , strCP09
            strUserNum = strUser1Num 'Added by Morgan 2022/4/8
         End If
      End If
   Next
   End With
End Sub

Private Sub cmdJoinAct_Click(Index As Integer)
   If Index = 0 Then
      If JoinUpdate() = False Then Exit Sub
      If Check1.Value = vbChecked Then
         OpenWord
      End If
   End If
   Frame3.Visible = False
   Frame1.Enabled = True
   SetWordPos
End Sub

Private Sub cmdRecNo_Click()
   Dim iRow As Integer
   Dim bolOK As Boolean
   
   With MSHFlexGrid2
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) <> "" Then
         m_RecNo = .TextMatrix(iRow, 2)
         m_PdfName = PUB_CaseNo2FileName(.TextMatrix(iRow, 4), .TextMatrix(iRow, 5), .TextMatrix(iRow, 6), .TextMatrix(iRow, 7)) & "." & .TextMatrix(iRow, 8) & ".CUS.PDF"
         m_DocFullPath = App.path & "\$" & m_RecNo & ".doc" 'Added by Morgan 2025/10/21
         LoadCPF m_RecNo 'Added by Morgan 2021/12/9
         bolOK = True
         Exit For
      End If
   Next
   If bolOK Then
      Frame4.Visible = False
      Frame1.Enabled = True
      SetWordPos
   Else
      MsgBox "請點選一筆收文號！", vbExclamation
   End If
   End With
End Sub

'Added by Morgan 2021/12/9
'載入原始檔區客戶函
Private Sub LoadCPF(pRecNo As String)
   Dim stSQL As String, intQ As Integer, stDoc As String
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select cpf13,cp01,cp02,cp03,cp04,cp10,cpf02 from CasePaperFile,caseprogress" & _
            " where cpf01='" & pRecNo & "' and cp09(+)=cpf01 and substr(upper(cpf02),-8)='.CUS.DOC'"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      stDoc = App.path & "\$TEMP"
      'end 2025/10/20
      With rsQuery
      If PUB_GetFtpFile(.Fields("cpf13"), stDoc, "CASEPAPERFILE", True) Then
         m_WordDoc.Close wdDoNotSaveChanges
         Set m_WordDoc = m_WordAp.Documents.Open(stDoc)
      End If
      End With
   End If
   m_WordDoc.SaveAs m_DocFullPath
End Sub

Private Sub Command1_Click()
   'Added by Lydia 2021/11/29 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
      Exit Sub
   End If
   'end 2021/11/29
   Frame2.Visible = False
   Frame1.Enabled = True
   SetWordPos 'Added by Morgan 2016/10/12
End Sub

Private Sub Command2_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   If Conver2Pdf(Index) = True Then
      Unload Me
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command3_Click()
   Unload Me
End Sub

'Added by Morgan 2018/8/31
Private Sub Command4_Click()
   If Trim(Text1.Text) = "" Or Trim(Text1.Text) = Text1.Tag Then
      MsgBox Text1.Tag, vbExclamation
      Text1.SetFocus
   ElseIf DeleteAppForm() = True Then
      Unload Me
   End If
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
      'Me.Left = 0
      
      'Added by Morgan 2017/8/9 若用Word轉Pdf功能則不必設定印表機
      If pub_Word2Pdf = False Then
         m_Os_Printer = PUB_GetOsDefaultPrinter
         PUB_SetOsDefaultPrinter "PDFCreator"
      'Added by Morgan 2016/9/20
      'Win7無法用IE開Word改將Word開在Form內
      'If Val(stOsVerNo) > 6 Then 'Removed by Morgan 2016/9/26 改都不用IE
         PUB_SetWordActivePrinter '切換Word印表機到PDFCreator
      End If
      
         OpenDoc2
      'Else
      '   OpenDoc
      'End If 'Addedd by Morgan 2016/9/20
      
      'Added by Morgan 2017/8/9 若用Word轉Pdf功能則不必設定印表機
      If pub_Word2Pdf = False Then
         PUB_SetOsDefaultPrinter m_Os_Printer
      End If
      
      'Added by Morgan 2015/11/10
      '檢查是否有指示信判發信退回意見
      If InStr(UCase(m_PdfName), ".DATA.PDF") > 0 Then
         'Modified by Morgan 2016/3/30 +判斷申請國家非臺灣時可點"上傳+EMail"按鈕
         strExc(0) = "select af10,cp12,pa09,af07,af06,cp01,cp04,cp140 from appform,caseprogress,patent where af01='" & m_RecNo & "' and cp09(+)=af01 and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With RsTemp
            If Not IsNull(.Fields("af10")) Then
               Text1 = .Fields("af10")
               Text1.Locked = True 'Added by Morgan 2020/9/7
               Frame1.Enabled = False
               Frame2.Visible = True
               Command1.Width = Text1.Width 'Added by Morgan 2018/8/31
               SetWordPos 'Added by Morgan 2016/10/12
            End If
            If "" & .Fields("pa09") <> "000" Then
               'Modified by Morgan 2016/5/20 不印指示信,只印寄件備份就好--玲玲
               'If Left(RsTemp("cp12"), 1) = "F" Then
               '   Command2(2).Caption = "上傳+列印+EMail"
               'Else
                  Command2(2).Caption = "上傳+EMail"
               'End If
               'end 2016/5/20
               
               'Modified by Morgan 2016/5/13 自判或已判發才可EMail
               If .Fields("af07") > 0 Or .Fields("af06") = strUserNum Then
                  Command2(2).Enabled = True
               End If
            End If
            'Added by Morgan 2018/8/30
            '開放EPC子案指示信可刪除
            If .Fields("cp01") = "CFP" Then
               If .Fields("cp04") <> "00" Then
                  cmdDelete.Enabled = True
                  m_CP140 = "" & .Fields("cp140")
               End If
            End If
            'end 2018/8/30
            End With
          End If
      'Added by Morgan 2018/11/1
      ElseIf m_eCustNo <> "" Then
         Command2(2).Enabled = True
         If GetRecNoList() = True Then
            Frame4.Top = Frame2.Top
            Frame4.Visible = True
            Frame1.Enabled = False
            SetWordPos
         End If
      'end 2018/11/1
      'Added by Morgan 2016/12/13
      Else
         'Added by Morgan 2019/1/17
         '檢查是否有客戶涵判發信退回意見
         strExc(0) = "select LP37 from letterprogress where lp01='" & m_RecNo & "' and LP37 is not null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Text1 = RsTemp("LP37")
            Text1.Locked = True 'Added by Morgan 2020/9/7
            Frame1.Enabled = False
            Frame2.Visible = True
            Command1.Width = Text1.Width
            SetWordPos
         End If
         'end 2019/1/17
         
         'Modify By Sindy 2020/10/30 + Or Left(UCase(m_PdfName), 1) = "T"
         If Left(UCase(m_PdfName), 3) = "CFP" Or Left(UCase(m_PdfName), 1) = "T" Then
            cmdJoin.Enabled = True 'Added by Morgan 2018/10/15 考慮附件過濾問題,P案先不開放
         End If
         
         strExc(0) = "select cp09 from caseprogress where cp43='" & m_RecNo & "' and cp10='990'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox "本信函另有副本信函請記得要一併修改！", vbExclamation
         End If
      End If
      'end 2015/11/10
   End If
   Me.ZOrder
   
End Sub

Private Sub Form_Click()
If m_bResize Then SetWordPos
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cboPrinter
   'Modified by Morgan 2014/9/9
   'SetFileAssociation
   txtPDFPath = PUB_SetFileAssociation
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   Else
      KillTemp
   End If
   'stOsVerNo = PUB_GetVersionNo 'Added by Morgan 2016/9/21
End Sub

Private Sub KillTemp()
On Error Resume Next
   Kill App.path & "\$*.doc"
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
   Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '更新已齊備日(手動上傳)
   PUB_UpdateLP03 'Add By Sindy 2025/1/2
   
   Set frm1105_1 = Nothing
   
   'Modified by Morgan 2021/12/7
   'If Not m_WordAp Is Nothing Then
   '   m_WordAp.Quit wdDoNotSaveChanges
   '   Set m_WordAp = Nothing
   'End If
   If Not m_WordDoc Is Nothing Then
      m_WordDoc.Close wdDoNotSaveChanges
   End If
   If Not m_WordAp Is Nothing Then
      If m_WordAp.Documents.Count = 0 Then
         m_WordAp.Quit wdDoNotSaveChanges
      End If
      Set m_WordAp = Nothing
   End If
   'end 2021/12/7
   
   PUB_SendMailCache 'Added by Morgan 2015/12/21
   
   'Added by Morgan 2015/11/17
   '回待處理區
   If Not m_PrevForm Is Nothing Then
      If UCase(m_PrevForm.Name) = UCase("frm210149") Then
         m_PrevForm.Show
         m_PrevForm.PubShowNextData
      End If
      'Added by Lydia 2024/03/04 ACS-TIPS案請款作業通知Email
      'Mark by Lydia 2024/03/15 保留
'      If UCase(m_PrevForm.Name) = UCase("frm071006") Then
'         Call Pub_ShowMailACS(m_RecNo, "")
'         Unload m_PrevForm
'      End If
      'end 2024/03/04
   End If
   'end 2015/11/17
End Sub

Private Sub Form_Resize()
   If Me.WindowState = 0 Or Me.WindowState = 2 Then
      m_WordTop = Frame1.Top + Frame1.Height
      m_WordWidth = Me.Width - 200
      m_WordHeight = Me.Height - 400 - Frame1.Height - txtPDFPath.Height
      txtPDFPath.Top = m_WordTop + m_WordHeight
      'Added by Morgan 2016/9/20
      If bolOpened = True Then
         SetWordPos
      End If
      'end 2016/9/20
      Label3.Top = txtPDFPath.Top + 50
   End If
End Sub

Private Sub SetWordPos()
On Error GoTo ErrHnd
m_bResize = False

If m_WordAp Is Nothing Then Exit Sub 'Added by Morgan 2022/6/28
m_WordAp.Width = m_WordWidth / 20
'Added by Morgan 2016/10/12
If Frame2.Visible Then
   m_WordAp.Height = (m_WordHeight - Frame2.Height) / 20
   m_WordAp.Move 0, (m_WordTop + Frame2.Height) / 20
'Added by Morgan 2018/10/15
ElseIf Frame3.Visible Then
   m_WordAp.Height = (m_WordHeight - Frame3.Height) / 20
   m_WordAp.Move 0, (m_WordTop + Frame3.Height) / 20
'Added by Morgan 2018/11/1
ElseIf Frame4.Visible Then
   m_WordAp.Height = (m_WordHeight - Frame4.Height) / 20
   m_WordAp.Move 0, (m_WordTop + Frame4.Height) / 20
Else
'end 2016/10/12
   m_WordAp.Height = m_WordHeight / 20
   m_WordAp.Move 0, m_WordTop / 20
End If 'Added by Morgan 2016/10/12
Exit Sub
ErrHnd:
   m_bResize = True
End Sub

Private Sub MSHFlexGrid1_Click()
   intI = MSHFlexGrid1.MouseRow
   If intI > 0 Then
      If MSHFlexGrid1.TextMatrix(intI, 0) = "" Then
         If MSHFlexGrid1.TextMatrix(intI, 7) = 0 And MSHFlexGrid1.TextMatrix(intI, 6) <> "" And MSHFlexGrid1.TextMatrix(intI, 6) <> strUserNum Then
            MsgBox "該程序尚未判發不可合併！", vbCritical
         Else
            MSHFlexGrid1.TextMatrix(intI, 0) = "V"
         End If
         
      ElseIf MSHFlexGrid1.TextMatrix(intI, 5) <> "" Then
         MsgBox "之前已合併不可取消！", vbExclamation
      Else
         MSHFlexGrid1.TextMatrix(intI, 0) = ""
      End If
   End If
End Sub

Private Sub MSHFlexGrid2_Click()
   GridClick MSHFlexGrid2, intLastRow, 0
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   OpenIme
End Sub

Private Sub Text1_LostFocus()
   CloseIme
End Sub

'Added by Morgan 2018/11/1
'已發文程序清單
Private Function GetRecNoList() As Boolean
   Dim iRow As Integer
   
   SetGrid MSHFlexGrid2, True
   
   strExc(0) = "select '' V,sqldatet(b.cp27) 發文日,b.cp09 總收文號,decode(nvl(tm10,sp09),'000',cpm03,cpm04) 案件性質,b.cp01,b.cp02,b.cp03,b.cp04,b.cp10" & _
      " from caseprogress a,caseprogress b,trademark,servicepractice,casepropertymap" & _
      " where a.cp09='" & m_RecNo & "'" & _
      " and b.cp01(+)=a.cp01 and b.cp02(+)=a.cp02 and b.cp03(+)=a.cp03 and b.cp04(+)=a.cp04" & _
      " and b.cp27>19221111" & _
      " and tm01(+)=a.cp01 and tm02(+)=a.cp02 and tm03(+)=a.cp03 and tm04(+)=a.cp04" & _
      " and sp01(+)=a.cp01 and sp02(+)=a.cp02 and sp03(+)=a.cp03 and sp04(+)=a.cp04" & _
      " and cpm01(+)=b.cp01 and cpm02(+)=b.cp10 order by b.cp27 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With MSHFlexGrid2
      .Visible = False
      Set .Recordset = RsTemp
      SetGrid MSHFlexGrid2
      
      For iRow = 1 To .Rows - 1
         .TextMatrix(iRow, 3) = .TextMatrix(iRow, 3) & PUB_GetRelateCasePropertyName(.TextMatrix(iRow, 2), "1")
      Next
      
      .Visible = True
      End With
      GetRecNoList = True
   Else
      MsgBox "無已發文程序！", vbExclamation
   End If

End Function
