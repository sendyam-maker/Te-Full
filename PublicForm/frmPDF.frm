VERSION 5.00
Begin VB.Form frmPDF 
   BackColor       =   &H80000018&
   BorderStyle     =   0  '沒有框線
   Caption         =   "建立PDF檔"
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5930
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   5930
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtStatus 
      BackColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   0
      Top             =   1050
      Width           =   5640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "PDF建立中...請稍候..."
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   26.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   5670
   End
End
Attribute VB_Name = "frmPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/15 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Created by Morgan 2012/10/31
Option Explicit

Const recDepth = 3

Private Const EM_FMTLINES = &HC8

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
 (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private WithEvents PDFCreator1 As PDFCreator.clsPDFCreator
Attribute PDFCreator1.VB_VarHelpID = -1
Private pErr As clsPDFCreatorError, opt As clsPDFCreatorOptions
Private noStart As Boolean, fac As Double, StartTime As Date
Public m_iStatusCode As Integer '1:成功 2:失敗
Dim iDefault As Integer 'Add By Sindy 2013/9/4


Private Sub Form_Load()
   Dim i As Integer
   
   Me.Top = Forms(0).Top + Forms(0).Height / 2 - Me.Height / 2
   Me.Left = Forms(0).Left + Forms(0).Width / 2 - Me.Width / 2
   
   'Added by Lydia 2019/01/16 檢查程式是否在執行中 ( 婉莘在國外付款明細表發email時, 第1個PDFCreater尚未結束,接著開啟造成錯誤 )
   For i = 1 To 10
      If PUB_CheckIsRunning("PDFCreator.exe") = True Then
         Sleep 1000
      Else
         Exit For
      End If
  Next
  'end 2019/01/16
  
   Set PDFCreator1 = New clsPDFCreator
   Set pErr = New clsPDFCreatorError
   With PDFCreator1
    .cVisible = True
    If .cStart("/NoProcessingAtStartup") = False Then
      If .cStart("/NoProcessingAtStartup", True) = False Then
       Exit Sub
      End If
      .cVisible = True
    End If
    ' Get the options
    Set opt = .cOptions
    .cClearCache
    noStart = False
   End With
   
   'Add By Sindy 2013/9/4
   For i = 0 To Printers.Count - 1
      If Printers(i).DeviceName = Printer.DeviceName Then iDefault = i
   Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim intX As Integer 'Added by Lydia 2019/01/16

 PDFCreator1.cPrinterStop = False
 If noStart = False Then
  PDFCreator1.cClose
  While PDFCreator1.cProgramIsRunning
   DoEvents
   Sleep 100
  Wend
 End If
  
 Set PDFCreator1 = Nothing
 Set pErr = Nothing
 Set opt = Nothing
 
 Set frmPDF = Nothing
End Sub

Public Sub StartProcess(Optional pPath As String, Optional pFileName As String)
 Dim strPath As String, strFileName As String, iPos As Integer
 
 m_iStatusCode = 0
 
   If pPath <> "" Then
      strPath = pPath
      If Dir(strPath, vbDirectory) = "" Then
         MkDir strPath
      End If
   Else
      strPath = PUB_Getdesktop
   End If
   
   If pFileName <> "" Then
      strFileName = pFileName
   Else
      strFileName = "@@" & Format(Now, "YYYYMMDDHHmmss")
   End If
   
   If UCase(Right(strFileName, 4)) <> ".PDF" Then
      strFileName = strFileName & ".pdf"
   End If
   
 With opt
  .AutosaveDirectory = strPath
  .AutosaveFilename = strFileName
  .UseAutosave = 1
  .UseAutosaveDirectory = 1
  .AutosaveFormat = 0 ' PDF
 End With
 Set PDFCreator1.cOptions = opt
 
 Set Printer = Printers(PrinterIndex("PDFCreator"))
End Sub

Public Sub EndtProcess()
On Error Resume Next 'Add By Sindy 2024/11/1

   PDFCreator1.cPrinterStop = False
   StartTime = Now
   Screen.MousePointer = vbHourglass
   'Modified by Morgan 2016/12/22 有可能沒有觸發 eReady 事件,加判斷沒有錯誤且未轉檔完
   'Do While m_iStatusCode = 0
   'Modified by Morgan 2016/12/29
   'Do While (m_iStatusCode = 0 And PDFCreator1.cError.Number = 0 And Not PDFCreator1.cIsConverted)
   Do While m_iStatusCode = 0
      If PDFCreator1.cIsConverted Then
         If Dir(PDFCreator1.cOutputFilename) <> "" Then
            PDFCreator1.cPrinterStop = True
            m_iStatusCode = 1
         End If
      End If
   'end 2016/12/29
   'end 2016/12/22
     DoEvents
     Sleep 1000
   Loop
   
   Set Printer = Printers(iDefault) 'Add By Sindy 2013/9/4
End Sub

Private Function PrinterIndex(Printername As String) As Long
 Dim i As Long
 For i = 0 To Printers.Count - 1
  If UCase(Printers(i).DeviceName) = UCase$(Printername) Then
   PrinterIndex = i
   Exit For
  End If
 Next i
End Function

Private Sub PDFCreator1_eReady()
 AddStatus """" & PDFCreator1.cOutputFilename & """ was created! (" & _
  DateDiff("s", StartTime, Now) & " seconds)"
 PDFCreator1.cPrinterStop = True
 m_iStatusCode = 1
 Screen.MousePointer = vbNormal
End Sub

Private Sub PDFCreator1_eError()
 Set pErr = PDFCreator1.cError
 AddStatus "Error[" & pErr.Number & "]: " & pErr.Description
 m_iStatusCode = 2
 Screen.MousePointer = vbNormal
End Sub

Private Sub AddStatus(Str1 As String)
 With txtStatus
  If LenB(.Text) = 0 Then
    .Text = time & ": " & Str1
   Else
    .Text = .Text & vbCrLf & time & ": " & Str1
  End If
  .SelStart = Len(.Text)
 End With
End Sub


